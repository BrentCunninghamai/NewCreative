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
