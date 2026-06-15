import httpx
import respx

from m365_migrate.graph_client import GraphClient
from m365_migrate.models import SourceUser
from m365_migrate.workloads import users as uw

BASE = "https://graph.microsoft.com/v1.0"


def _user(uid, upn, user_type="Member", display="X"):
    return {"id": uid, "userPrincipalName": upn, "userType": user_type, "displayName": display}


@respx.mock
def test_discover_users_parses_graph(static_token):
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200, json={"value": [_user("1", "jane@contoso.onmicrosoft.com")]}
        )
    )
    client = GraphClient(static_token)
    found = uw.discover_users(client)
    assert len(found) == 1
    assert found[0].user_principal_name == "jane@contoso.onmicrosoft.com"


def test_plan_rewrites_skips_and_flags_conflicts(config):
    users = [
        SourceUser(id="1", user_principal_name="jane@contoso.onmicrosoft.com", display_name="Jane"),
        SourceUser(id="2", user_principal_name="guest@contoso.onmicrosoft.com", user_type="Guest"),
        SourceUser(id="3", user_principal_name="bob@contoso.onmicrosoft.com"),
    ]
    existing = {"bob@fabrikam.onmicrosoft.com"}  # already exists in target

    planned = uw.plan_users(users, config, existing_target_upns=existing)

    by_src = {p.source_id: p for p in planned}
    assert by_src["1"].action == "create"
    assert by_src["1"].target_upn == "jane@fabrikam.onmicrosoft.com"
    assert by_src["2"].action == "skip"  # guest filtered
    assert by_src["3"].action == "conflict"  # exists in target


def test_migrate_dry_run_writes_nothing(config, static_token):
    planned = uw.plan_users(
        [SourceUser(id="1", user_principal_name="jane@contoso.onmicrosoft.com", display_name="Jane")],
        config,
    )
    # No respx mock registered -> any real HTTP call would fail, proving dry run is inert.
    client = GraphClient(static_token)
    results = uw.migrate_users(client, planned, config, dry_run=True)
    assert results == [{"target_upn": "jane@fabrikam.onmicrosoft.com", "status": "would-create"}]


def test_from_graph_parses_manager_and_licenses():
    data = {
        "id": "1",
        "userPrincipalName": "jane@contoso.onmicrosoft.com",
        "userType": "Member",
        "assignedLicenses": [{"skuId": "sku-A"}, {"skuId": "sku-B"}],
        "manager": {"id": "9", "userPrincipalName": "boss@contoso.onmicrosoft.com"},
    }
    u = SourceUser.from_graph(data)
    assert u.manager_upn == "boss@contoso.onmicrosoft.com"
    assert u.assigned_sku_ids == ["sku-A", "sku-B"]


@respx.mock
def test_enrich_dry_run_reports_without_writing(config, static_token):
    # Target has both jane and her manager; one license SKU is available there.
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"},
                    {"id": "tb", "userPrincipalName": "boss@fabrikam.onmicrosoft.com"},
                ]
            },
        )
    )
    respx.get(f"{BASE}/subscribedSkus").mock(
        return_value=httpx.Response(200, json={"value": [{"skuId": "sku-A"}]})
    )
    users = [
        SourceUser(
            id="1",
            user_principal_name="jane@contoso.onmicrosoft.com",
            manager_upn="boss@contoso.onmicrosoft.com",
            assigned_sku_ids=["sku-A", "sku-unavailable"],
        )
    ]
    client = GraphClient(static_token)
    results = uw.enrich_users(client, users, config, dry_run=True)

    assert results[0]["manager"] == "would-set"
    # Only the SKU available in the target counts.
    assert results[0]["licenses"] == "would-assign:1"


@respx.mock
def test_enrich_execute_sets_manager_and_assigns_license(config, static_token):
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"},
                    {"id": "tb", "userPrincipalName": "boss@fabrikam.onmicrosoft.com"},
                ]
            },
        )
    )
    respx.get(f"{BASE}/subscribedSkus").mock(
        return_value=httpx.Response(200, json={"value": [{"skuId": "sku-A"}]})
    )
    mgr_ref = respx.put(f"{BASE}/users/tj/manager/$ref").mock(
        return_value=httpx.Response(204)
    )
    assign = respx.post(f"{BASE}/users/tj/assignLicense").mock(
        return_value=httpx.Response(200, json={"id": "tj"})
    )
    users = [
        SourceUser(
            id="1",
            user_principal_name="jane@contoso.onmicrosoft.com",
            manager_upn="boss@contoso.onmicrosoft.com",
            assigned_sku_ids=["sku-A"],
        )
    ]
    client = GraphClient(static_token)
    results = uw.enrich_users(client, users, config, dry_run=False)

    assert mgr_ref.called
    assert assign.called
    assert results[0]["manager"] == "set"
    assert results[0]["licenses"] == "assigned:1"
    # The manager $ref points at the resolved target manager id.
    body = mgr_ref.calls.last.request.content.decode()
    assert "/users/tb" in body


@respx.mock
def test_enrich_skips_user_missing_in_target(config, static_token):
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    respx.get(f"{BASE}/subscribedSkus").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    users = [SourceUser(id="1", user_principal_name="ghost@contoso.onmicrosoft.com")]
    client = GraphClient(static_token)
    results = uw.enrich_users(client, users, config, dry_run=False)
    assert results[0]["status"] == "skipped:not-in-target"


@respx.mock
def test_migrate_execute_creates_user(config, static_token):
    create = respx.post(f"{BASE}/users").mock(
        return_value=httpx.Response(201, json={"id": "new-id", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"})
    )
    planned = uw.plan_users(
        [SourceUser(id="1", user_principal_name="jane@contoso.onmicrosoft.com", display_name="Jane")],
        config,
    )
    client = GraphClient(static_token)
    results = uw.migrate_users(client, planned, config, dry_run=False)

    assert create.called
    assert results[0]["status"] == "created"
    assert results[0]["target_id"] == "new-id"
    # A real, non-placeholder password must have been generated.
    sent_body = create.calls.last.request.content.decode()
    assert "REPLACE_ME" not in sent_body
