import httpx
import respx

from m365_migrate.graph_client import GraphClient
from m365_migrate.models import DriveGrant, DriveItem, PlannedDriveItem
from m365_migrate.workloads import files as fw

BASE = "https://graph.microsoft.com/v1.0"
SRC = "/users/jane@contoso.onmicrosoft.com/drive"
TGT = "/users/jane@fabrikam.onmicrosoft.com/drive"


def _folder(fid, name, parent="/drive/root:"):
    return {"id": fid, "name": name, "folder": {}, "size": 0, "parentReference": {"path": parent}}


def _file(fid, name, size=10, parent="/drive/root:", grants=None):
    item = {"id": fid, "name": name, "size": size, "parentReference": {"path": parent}}
    if grants:
        item["permissions"] = grants
    return item


def test_from_graph_parses_path_and_grants():
    data = _file(
        "f1",
        "Report.txt",
        parent="/drive/root:/Reports/2025",
        grants=[
            {"roles": ["write"], "grantedToV2": {"user": {"userPrincipalName": "bob@contoso.onmicrosoft.com"}}}
        ],
    )
    item = DriveItem.from_graph(data)
    assert item.parent_path == "Reports/2025"
    assert item.relative_path == "Reports/2025/Report.txt"
    assert not item.is_folder
    assert item.grants[0].upn == "bob@contoso.onmicrosoft.com"
    assert item.grants[0].roles == ["write"]


def test_resolve_drive_roots(config):
    src, tgt = fw.resolve_drive_roots(config, user="jane@contoso.onmicrosoft.com")
    assert src == SRC
    assert tgt == TGT

    src, tgt = fw.resolve_drive_roots(config, site="s1", target_site="s2")
    assert src == "/sites/s1/drive"
    assert tgt == "/sites/s2/drive"


def test_resolve_drive_roots_errors(config):
    import pytest

    with pytest.raises(ValueError):
        fw.resolve_drive_roots(config)
    with pytest.raises(ValueError):
        fw.resolve_drive_roots(config, site="s1")  # missing target site


@respx.mock
def test_discover_walks_folders_breadth_first(static_token):
    # Root has a folder and a file; the folder has one nested file.
    respx.get(f"{BASE}{SRC}/root/children").mock(
        return_value=httpx.Response(
            200,
            json={"value": [_folder("d1", "Reports"), _file("f0", "top.txt")]},
        )
    )
    respx.get(f"{BASE}{SRC}/items/d1/children").mock(
        return_value=httpx.Response(
            200,
            json={"value": [_file("f1", "inner.txt", parent="/drive/root:/Reports")]},
        )
    )
    client = GraphClient(static_token)
    items = fw.discover_drive_items(client, SRC)
    # Folder is listed before the item it contains.
    names = [i.relative_path for i in items]
    assert names == ["Reports", "top.txt", "Reports/inner.txt"]


def test_plan_skips_oversized_files(config):
    items = [
        DriveItem(id="d1", name="Reports", is_folder=True),
        DriveItem(id="f1", name="ok.txt", size=100),
        DriveItem(id="f2", name="big.bin", size=fw.SIMPLE_UPLOAD_LIMIT + 1),
    ]
    planned = fw.plan_drive_items(items, config)
    by_id = {p.source_id: p for p in planned}
    assert by_id["d1"].action == "copy"
    assert by_id["f1"].action == "copy"
    assert by_id["f2"].action == "skip"


def test_plan_rewrites_grant_upns(config):
    items = [
        DriveItem(
            id="f1",
            name="ok.txt",
            size=10,
            grants=[DriveGrant(upn="bob@contoso.onmicrosoft.com", roles=["read"])],
        )
    ]
    planned = fw.plan_drive_items(items, config)
    assert planned[0].target_grants[0].upn == "bob@fabrikam.onmicrosoft.com"


@respx.mock
def test_migrate_dry_run_writes_nothing(config, static_token):
    respx.get(f"{BASE}/users").mock(return_value=httpx.Response(200, json={"value": []}))
    planned = [
        PlannedDriveItem(source_id="d1", name="Reports", relative_path="Reports", is_folder=True, action="copy"),
        PlannedDriveItem(source_id="f1", name="ok.txt", relative_path="Reports/ok.txt", is_folder=False, size=10, action="copy"),
    ]
    client = GraphClient(static_token)
    results = fw.migrate_drive_items(client, client, planned, SRC, TGT, config, dry_run=True)
    assert results[0]["status"] == "would-create-folder"
    assert results[1]["status"] == "would-upload"


@respx.mock
def test_migrate_execute_creates_folder_uploads_and_grants(config, static_token):
    # Target user directory (for grant resolution) includes bob.
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "tb", "userPrincipalName": "bob@fabrikam.onmicrosoft.com"}]}
        )
    )
    make_folder = respx.post(f"{BASE}{TGT}/root/children").mock(
        return_value=httpx.Response(201, json={"id": "td1"})
    )
    download = respx.get(f"{BASE}{SRC}/items/f1/content").mock(
        return_value=httpx.Response(200, content=b"hello bytes")
    )
    upload = respx.put(f"{BASE}{TGT}/root:/Reports/ok.txt:/content").mock(
        return_value=httpx.Response(201, json={"id": "tf1"})
    )
    invite = respx.post(f"{BASE}{TGT}/root:/Reports/ok.txt:/invite").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    planned = [
        PlannedDriveItem(source_id="d1", name="Reports", relative_path="Reports", is_folder=True, action="copy"),
        PlannedDriveItem(
            source_id="f1",
            name="ok.txt",
            relative_path="Reports/ok.txt",
            is_folder=False,
            size=11,
            action="copy",
            target_grants=[DriveGrant(upn="bob@fabrikam.onmicrosoft.com", roles=["write"])],
        ),
    ]
    client = GraphClient(static_token)
    results = fw.migrate_drive_items(client, client, planned, SRC, TGT, config, dry_run=False)

    assert make_folder.called and download.called and upload.called and invite.called
    assert results[0]["status"] == "folder-created"
    assert results[1]["status"] == "uploaded"
    assert results[1]["grants"] == "granted:1"
    # The uploaded bytes are the ones downloaded from the source.
    assert upload.calls.last.request.content == b"hello bytes"
    # Grant invite uses the resolved target grantee and a write role.
    invite_body = invite.calls.last.request.content.decode()
    assert "bob@fabrikam.onmicrosoft.com" in invite_body
    assert '"write"' in invite_body


@respx.mock
def test_migrate_skips_unresolved_grant(config, static_token):
    respx.get(f"{BASE}/users").mock(return_value=httpx.Response(200, json={"value": []}))
    respx.get(f"{BASE}{SRC}/items/f1/content").mock(
        return_value=httpx.Response(200, content=b"x")
    )
    respx.put(f"{BASE}{TGT}/root:/ok.txt:/content").mock(
        return_value=httpx.Response(201, json={"id": "tf1"})
    )
    planned = [
        PlannedDriveItem(
            source_id="f1",
            name="ok.txt",
            relative_path="ok.txt",
            is_folder=False,
            size=1,
            action="copy",
            target_grants=[DriveGrant(upn="ghost@fabrikam.onmicrosoft.com", roles=["read"])],
        )
    ]
    client = GraphClient(static_token)
    results = fw.migrate_drive_items(client, client, planned, SRC, TGT, config, dry_run=False)
    assert results[0]["status"] == "uploaded"
    assert "grants" not in results[0]  # grantee not in target -> nothing reapplied
