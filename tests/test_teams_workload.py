import httpx
import respx

from m365_migrate.graph_client import GraphClient
from m365_migrate.models import Channel, PlannedChannel, PlannedTeam, SourceTeam, group_is_team
from m365_migrate.workloads import teams as tw

BASE = "https://graph.microsoft.com/v1.0"


def _group(gid, nickname, *, team=True):
    data = {"id": gid, "displayName": nickname.title(), "mailNickname": nickname}
    data["resourceProvisioningOptions"] = ["Team"] if team else []
    return data


def test_group_is_team():
    assert group_is_team({"resourceProvisioningOptions": ["Team"]})
    assert group_is_team({"resourceProvisioningOptions": ["team"]})  # case-insensitive
    assert not group_is_team({"resourceProvisioningOptions": []})
    assert not group_is_team({})


def test_channel_default_detection():
    assert Channel(display_name="General").is_default
    assert Channel(display_name=" general ").is_default
    assert not Channel(display_name="Engineering").is_default


@respx.mock
def test_discover_only_returns_teams_with_channels(static_token):
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(
            200,
            json={"value": [_group("g1", "sales"), _group("g2", "plainsec", team=False)]},
        )
    )
    respx.get(f"{BASE}/teams/g1/channels").mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {"id": "c0", "displayName": "General", "membershipType": "standard"},
                    {"id": "c1", "displayName": "Deals", "membershipType": "standard"},
                ]
            },
        )
    )
    client = GraphClient(static_token)
    teams = tw.discover_teams(client)
    assert len(teams) == 1  # the non-team group is excluded
    assert teams[0].mail_nickname == "sales"
    assert [c.display_name for c in teams[0].channels] == ["General", "Deals"]


def test_plan_skips_team_without_target_group():
    teams = [SourceTeam(id="g1", mail_nickname="sales", display_name="Sales")]
    planned = tw.plan_teams(teams, existing_target_groups={})
    assert planned[0].action == "skip"
    assert "groups sync" in planned[0].reason


def test_plan_classifies_channels():
    team = SourceTeam(
        id="g1",
        mail_nickname="sales",
        display_name="Sales",
        channels=[
            Channel(display_name="General", membership_type="standard"),
            Channel(display_name="Deals", membership_type="standard"),
            Channel(display_name="Execs", membership_type="private"),
        ],
    )
    planned = tw.plan_teams([team], existing_target_groups={"sales": "tg1"})
    assert planned[0].action == "provision"
    by_name = {c.display_name: c for c in planned[0].channels}
    assert by_name["General"].action == "skip"  # auto-created
    assert by_name["Deals"].action == "create"
    assert by_name["Execs"].action == "skip"  # private not yet supported


@respx.mock
def test_migrate_dry_run_writes_nothing(static_token):
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(
            200, json={"value": [_group("tg1", "sales", team=False)]}
        )
    )
    planned = [
        PlannedTeam(
            source_id="g1",
            mail_nickname="sales",
            display_name="Sales",
            action="provision",
            channels=[PlannedChannel(display_name="Deals", action="create")],
        )
    ]
    client = GraphClient(static_token)
    results = tw.migrate_teams(client, planned, dry_run=True)
    assert results[0]["team"] == "would-enable"
    assert results[0]["channels"] == "would-create:1"


@respx.mock
def test_migrate_execute_enables_team_and_creates_channels(static_token):
    # Target group "sales" exists but is not yet a team.
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(
            200, json={"value": [_group("tg1", "sales", team=False)]}
        )
    )
    enable = respx.put(f"{BASE}/groups/tg1/team").mock(return_value=httpx.Response(201, json={"id": "tg1"}))
    # After enabling, the team has only the default General channel.
    respx.get(f"{BASE}/teams/tg1/channels").mock(
        return_value=httpx.Response(200, json={"value": [{"id": "c0", "displayName": "General"}]})
    )
    make_channel = respx.post(f"{BASE}/teams/tg1/channels").mock(
        return_value=httpx.Response(201, json={"id": "c1"})
    )
    planned = [
        PlannedTeam(
            source_id="g1",
            mail_nickname="sales",
            display_name="Sales",
            action="provision",
            channels=[
                PlannedChannel(display_name="Deals", action="create"),
                PlannedChannel(display_name="General", action="skip", reason="default"),
            ],
        )
    ]
    client = GraphClient(static_token)
    results = tw.migrate_teams(client, planned, dry_run=False)

    assert enable.called and make_channel.called
    assert results[0]["team"] == "enabled"
    assert results[0]["channels"] == "created:1"
    body = make_channel.calls.last.request.content.decode()
    assert "Deals" in body and "standard" in body


@respx.mock
def test_migrate_skips_existing_team_and_channel(static_token):
    # Target group "sales" is already a team; "Deals" channel already exists.
    respx.get(f"{BASE}/groups").mock(
        return_value=httpx.Response(200, json={"value": [_group("tg1", "sales", team=True)]})
    )
    respx.get(f"{BASE}/teams/tg1/channels").mock(
        return_value=httpx.Response(
            200, json={"value": [{"id": "c0", "displayName": "General"}, {"id": "c1", "displayName": "Deals"}]}
        )
    )
    create_route = respx.post(f"{BASE}/teams/tg1/channels").mock(return_value=httpx.Response(201, json={}))
    planned = [
        PlannedTeam(
            source_id="g1",
            mail_nickname="sales",
            display_name="Sales",
            action="provision",
            channels=[PlannedChannel(display_name="Deals", action="create")],
        )
    ]
    client = GraphClient(static_token)
    results = tw.migrate_teams(client, planned, dry_run=False)

    assert results[0]["team"] == "exists"
    assert results[0]["channels"] == "created:0"  # Deals already present
    assert not create_route.called
