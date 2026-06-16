import httpx
import respx

from m365_migrate.graph_client import GraphClient
from m365_migrate.models import PlannedGroup, SourceGroup
from m365_migrate.workloads import groups as gw

BASE = "https://graph.microsoft.com/v1.0"


def test_from_graph_classifies_kinds_and_members():
    unified = SourceGroup.from_graph(
        {
            "id": "g1",
            "displayName": "Marketing",
            "mailNickname": "marketing",
            "groupTypes": ["Unified"],
            "mailEnabled": True,
            "securityEnabled": False,
            "members": [
                {"id": "u1", "userPrincipalName": "jane@contoso.onmicrosoft.com"},
                {"id": "d1"},  # non-user member (e.g. device) ignored
            ],
        }
    )
    assert unified.kind == "microsoft365"
    assert unified.member_upns == ["jane@contoso.onmicrosoft.com"]

    sec = SourceGroup.from_graph(
        {"id": "g2", "mailNickname": "sec", "securityEnabled": True, "mailEnabled": False}
    )
    assert sec.kind == "security"

    dist = SourceGroup.from_graph(
        {"id": "g3", "mailNickname": "dl", "securityEnabled": False, "mailEnabled": True}
    )
    assert dist.kind == "distribution"


def test_plan_marks_create_exists_and_skip(config):
    groups = [
        SourceGroup(id="g1", mail_nickname="marketing", group_types=["Unified"], mail_enabled=True),
        SourceGroup(id="g2", mail_nickname="sec", security_enabled=True),
        SourceGroup(id="g3", mail_nickname="dl", mail_enabled=True),  # distribution
    ]
    planned = gw.plan_groups(groups, config, existing_target_groups={"sec": "tsec"})
    by_id = {p.source_id: p for p in planned}
    assert by_id["g1"].action == "create"
    assert by_id["g2"].action == "exists"
    assert by_id["g3"].action == "skip"
    assert "not provisionable" in by_id["g3"].reason


def test_plan_rewrites_member_upns(config):
    groups = [
        SourceGroup(
            id="g1",
            mail_nickname="sec",
            security_enabled=True,
            member_upns=["jane@contoso.onmicrosoft.com"],
        )
    ]
    planned = gw.plan_groups(groups, config)
    assert planned[0].target_member_upns == ["jane@fabrikam.onmicrosoft.com"]


def test_from_graph_parses_owners():
    g = SourceGroup.from_graph(
        {
            "id": "g1",
            "mailNickname": "marketing",
            "groupTypes": ["Unified"],
            "mailEnabled": True,
            "owners": [
                {"id": "u1", "userPrincipalName": "jane@contoso.onmicrosoft.com"},
                {"id": "sp1"},  # non-user owner (e.g. service principal) ignored
            ],
        }
    )
    assert g.owner_upns == ["jane@contoso.onmicrosoft.com"]


def test_plan_rewrites_owner_upns(config):
    groups = [
        SourceGroup(
            id="g1",
            mail_nickname="sec",
            security_enabled=True,
            owner_upns=["jane@contoso.onmicrosoft.com"],
        )
    ]
    planned = gw.plan_groups(groups, config)
    assert planned[0].target_owner_upns == ["jane@fabrikam.onmicrosoft.com"]


def test_microsoft365_group_body():
    p = PlannedGroup(
        source_id="g1",
        mail_nickname="marketing",
        display_name="Marketing",
        kind="microsoft365",
        action="create",
        description="Marketing team",
    )
    body = p.to_graph_body()
    assert body["groupTypes"] == ["Unified"]
    assert body["mailEnabled"] is True
    assert body["securityEnabled"] is False
    assert body["description"] == "Marketing team"


def test_from_graph_parses_dynamic_membership():
    g = SourceGroup.from_graph(
        {
            "id": "g1",
            "mailNickname": "sales",
            "groupTypes": ["Unified", "DynamicMembership"],
            "mailEnabled": True,
            "membershipRule": 'user.department -eq "Sales"',
            "membershipRuleProcessingState": "On",
        }
    )
    assert g.is_dynamic
    assert g.membership_rule == 'user.department -eq "Sales"'


def test_dynamic_group_body_includes_rule():
    p = PlannedGroup(
        source_id="g1",
        mail_nickname="sales",
        display_name="Sales",
        kind="microsoft365",
        action="create",
        is_dynamic=True,
        membership_rule='user.department -eq "Sales"',
    )
    body = p.to_graph_body()
    assert body["groupTypes"] == ["Unified", "DynamicMembership"]
    assert body["membershipRule"] == 'user.department -eq "Sales"'
    assert body["membershipRuleProcessingState"] == "On"


def test_security_group_body_drops_null_description():
    p = PlannedGroup(
        source_id="g2",
        mail_nickname="sec",
        display_name="Sec",
        kind="security",
        action="create",
    )
    body = p.to_graph_body()
    assert body["groupTypes"] == []
    assert body["securityEnabled"] is True
    assert body["mailEnabled"] is False
    assert "description" not in body  # null is stripped


@respx.mock
def test_sync_dry_run_reports_without_writing(config, static_token):
    # Target has the security group and jane; bob is not in the target yet.
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "tsec", "mailNickname": "sec"}]}
        )
    )
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"}]},
        )
    )
    # Membership read for the existing group (jane not yet a member).
    respx.get(f"{BASE}/groups/tsec/members").mock(
        return_value=httpx.Response(200, json={"value": []})
    )
    planned = [
        PlannedGroup(
            source_id="g1",
            mail_nickname="sec",
            display_name="Sec",
            kind="security",
            action="exists",
            target_member_upns=[
                "jane@fabrikam.onmicrosoft.com",
                "bob@fabrikam.onmicrosoft.com",
            ],
        )
    ]
    client = GraphClient(static_token)
    results = gw.sync_groups(client, planned, config, dry_run=True)

    assert results[0]["group"] == "exists"
    assert results[0]["members"] == "would-add:1"  # only jane resolves
    assert results[0]["members_unresolved"] == 1  # bob missing in target


@respx.mock
def test_sync_execute_creates_group_and_adds_member(config, static_token):
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(200, json={"value": []})  # nothing in target yet
    )
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"}]},
        )
    )
    create = respx.post(f"{BASE}/groups").mock(
        return_value=httpx.Response(201, json={"id": "new-grp", "mailNickname": "marketing"})
    )
    add_member = respx.post(f"{BASE}/groups/new-grp/members/$ref").mock(
        return_value=httpx.Response(204)
    )
    planned = [
        PlannedGroup(
            source_id="g1",
            mail_nickname="marketing",
            display_name="Marketing",
            kind="microsoft365",
            action="create",
            target_member_upns=["jane@fabrikam.onmicrosoft.com"],
        )
    ]
    client = GraphClient(static_token)
    results = gw.sync_groups(client, planned, config, dry_run=False)

    assert create.called
    assert add_member.called
    assert results[0]["group"] == "created"
    assert results[0]["members"] == "added:1"
    # The member $ref points at the resolved target user's directory object.
    body = add_member.calls.last.request.content.decode()
    assert "/directoryObjects/tj" in body


@respx.mock
def test_sync_dry_run_reports_owners(config, static_token):
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "tsec", "mailNickname": "sec"}]})
    )
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"}]},
        )
    )
    respx.get(f"{BASE}/groups/tsec/members").mock(return_value=httpx.Response(200, json={"value": []}))
    respx.get(f"{BASE}/groups/tsec/owners").mock(return_value=httpx.Response(200, json={"value": []}))
    planned = [
        PlannedGroup(
            source_id="g1",
            mail_nickname="sec",
            display_name="Sec",
            kind="security",
            action="exists",
            target_owner_upns=[
                "jane@fabrikam.onmicrosoft.com",
                "ghost@fabrikam.onmicrosoft.com",
            ],
        )
    ]
    client = GraphClient(static_token)
    results = gw.sync_groups(client, planned, config, dry_run=True)
    assert results[0]["owners"] == "would-add:1"  # only jane resolves
    assert results[0]["owners_unresolved"] == 1


@respx.mock
def test_sync_execute_adds_owner(config, static_token):
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "tsec", "mailNickname": "sec"}]})
    )
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"}]},
        )
    )
    respx.get(f"{BASE}/groups/tsec/members").mock(return_value=httpx.Response(200, json={"value": []}))
    respx.get(f"{BASE}/groups/tsec/owners").mock(return_value=httpx.Response(200, json={"value": []}))
    add_owner = respx.post(f"{BASE}/groups/tsec/owners/$ref").mock(return_value=httpx.Response(204))
    planned = [
        PlannedGroup(
            source_id="g1",
            mail_nickname="sec",
            display_name="Sec",
            kind="security",
            action="exists",
            target_owner_upns=["jane@fabrikam.onmicrosoft.com"],
        )
    ]
    client = GraphClient(static_token)
    results = gw.sync_groups(client, planned, config, dry_run=False)
    assert add_owner.called
    assert results[0]["owners"] == "added:1"
    body = add_owner.calls.last.request.content.decode()
    assert "/directoryObjects/tj" in body


@respx.mock
def test_sync_dynamic_group_creates_with_rule_and_skips_members(config, static_token):
    respx.get(f"{BASE}/groups").mock(return_value=httpx.Response(200, json={"value": []}))
    respx.get(f"{BASE}/users").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "tj", "userPrincipalName": "jane@fabrikam.onmicrosoft.com"}]},
        )
    )
    create = respx.post(f"{BASE}/groups").mock(
        return_value=httpx.Response(201, json={"id": "dyn", "mailNickname": "sales"})
    )
    add_member = respx.post(f"{BASE}/groups/dyn/members/$ref").mock(return_value=httpx.Response(204))
    planned = [
        PlannedGroup(
            source_id="g1",
            mail_nickname="sales",
            display_name="Sales",
            kind="microsoft365",
            action="create",
            is_dynamic=True,
            membership_rule='user.department -eq "Sales"',
            # Even if assigned members are present, a dynamic group must ignore them.
            target_member_upns=["jane@fabrikam.onmicrosoft.com"],
        )
    ]
    client = GraphClient(static_token)
    results = gw.sync_groups(client, planned, config, dry_run=False)

    assert create.called
    assert not add_member.called  # membership is rule-driven, not assigned
    assert results[0]["members"] == "dynamic-rule"
    body = create.calls.last.request.content.decode()
    assert "DynamicMembership" in body and "membershipRule" in body


@respx.mock
def test_sync_skips_distribution_group(config, static_token):
    respx.get(f"{BASE}/groups").mock(return_value=httpx.Response(200, json={"value": []}))
    respx.get(f"{BASE}/users").mock(return_value=httpx.Response(200, json={"value": []}))
    planned = [
        PlannedGroup(
            source_id="g3",
            mail_nickname="dl",
            display_name="DL",
            kind="distribution",
            action="skip",
            reason="distribution group not provisionable via Graph",
        )
    ]
    client = GraphClient(static_token)
    results = gw.sync_groups(client, planned, config, dry_run=False)
    assert results[0]["status"].startswith("skipped:")
