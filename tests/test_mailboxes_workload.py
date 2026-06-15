import httpx
import respx

from m365_migrate.graph_client import GraphClient
from m365_migrate.models import PlannedMailbox, SourceMailbox
from m365_migrate.workloads import mailboxes as mw

BASE = "https://graph.microsoft.com/v1.0"

SETTINGS = {
    "timeZone": "Pacific Standard Time",
    "language": {"locale": "en-US", "displayName": "English (United States)"},
    "automaticRepliesSetting": {"status": "disabled"},
    "userPurpose": "user",  # read-only, must be dropped
}


def test_from_graph_keeps_only_settable_fields():
    mb = SourceMailbox.from_graph("jane@contoso.onmicrosoft.com", SETTINGS)
    assert mb.settings["timeZone"] == "Pacific Standard Time"
    assert "language" in mb.settings
    assert "userPurpose" not in mb.settings  # read-only stripped


@respx.mock
def test_discover_skips_users_without_mailbox(static_token):
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {"id": "u1", "userPrincipalName": "jane@contoso.onmicrosoft.com", "mail": "jane@contoso.com"},
                    {"id": "u2", "userPrincipalName": "nomail@contoso.onmicrosoft.com", "mail": None},
                    {"id": "u3", "userPrincipalName": "shared@contoso.onmicrosoft.com", "mail": "shared@contoso.com"},
                ]
            },
        )
    )
    # u1 has a mailbox; u3 has a mail address but no REST mailbox (404).
    respx.get(f"{BASE}/users/u1/mailboxSettings").mock(
        return_value=httpx.Response(200, json=SETTINGS)
    )
    respx.get(f"{BASE}/users/u3/mailboxSettings").mock(
        return_value=httpx.Response(404, json={"error": {"code": "MailboxNotEnabledForRESTAPI"}})
    )
    client = GraphClient(static_token, max_retries=0)
    found = mw.discover_mailboxes(client)
    assert [m.user_principal_name for m in found] == ["jane@contoso.onmicrosoft.com"]


def test_plan_settings_and_skips(config):
    mailboxes = [
        SourceMailbox(user_principal_name="jane@contoso.onmicrosoft.com", settings={"timeZone": "UTC"}),
        SourceMailbox(user_principal_name="ghost@contoso.onmicrosoft.com", settings={"timeZone": "UTC"}),
        SourceMailbox(user_principal_name="bare@contoso.onmicrosoft.com", settings={}),
    ]
    target_upns = {"jane@fabrikam.onmicrosoft.com", "bare@fabrikam.onmicrosoft.com"}
    planned = mw.plan_mailboxes(mailboxes, config, target_upns=target_upns)
    by_src = {p.source_upn: p for p in planned}
    assert by_src["jane@contoso.onmicrosoft.com"].action == "settings"
    assert by_src["jane@contoso.onmicrosoft.com"].target_upn == "jane@fabrikam.onmicrosoft.com"
    assert by_src["ghost@contoso.onmicrosoft.com"].action == "skip"  # not in target
    assert by_src["bare@contoso.onmicrosoft.com"].action == "skip"  # no settings


@respx.mock
def test_migrate_dry_run_writes_nothing(config, static_token):
    # get_target_user_ids needs the target users list even on a dry run, but no
    # PATCH mock is registered -> a write would fail, proving the dry run is inert.
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"}]},
        )
    )
    planned = [
        PlannedMailbox(
            source_upn="jane@contoso.onmicrosoft.com",
            target_upn="jane@fabrikam.onmicrosoft.com",
            action="settings",
            settings={"timeZone": "UTC"},
        )
    ]
    client = GraphClient(static_token)
    results = mw.migrate_mailboxes(client, planned, config, dry_run=True)
    assert results[0]["status"] == "would-update"


@respx.mock
def test_migrate_execute_patches_settings(config, static_token):
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"}]},
        )
    )
    patch = respx.patch(f"{BASE}/users/tj/mailboxSettings").mock(
        return_value=httpx.Response(200, json={"timeZone": "UTC"})
    )
    planned = [
        PlannedMailbox(
            source_upn="jane@contoso.onmicrosoft.com",
            target_upn="jane@fabrikam.onmicrosoft.com",
            action="settings",
            settings={"timeZone": "UTC"},
        )
    ]
    client = GraphClient(static_token)
    results = mw.migrate_mailboxes(client, planned, config, dry_run=False)

    assert patch.called
    assert results[0]["status"] == "updated"
    assert '"timeZone"' in patch.calls.last.request.content.decode()


@respx.mock
def test_migrate_skips_target_missing(config, static_token):
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    planned = [
        PlannedMailbox(
            source_upn="jane@contoso.onmicrosoft.com",
            target_upn="jane@fabrikam.onmicrosoft.com",
            action="settings",
            settings={"timeZone": "UTC"},
        )
    ]
    client = GraphClient(static_token)
    results = mw.migrate_mailboxes(client, planned, config, dry_run=False)
    assert results[0]["status"] == "skipped:not-in-target"
