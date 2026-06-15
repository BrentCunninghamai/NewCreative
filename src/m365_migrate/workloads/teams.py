"""Teams workload: discover, plan, and provision Microsoft Teams + channels.

A *team* is a Microsoft 365 group that has been Teams-enabled, so this workload
composes with the others rather than duplicating them:

* Team **membership** is the backing M365 group's membership — migrated by the
  groups workload (``groups sync``).
* Channel **files** live in the team's SharePoint document library — migrated by
  the files workload (``files migrate --site ...``).

This workload owns the Teams-specific layer, staged like the rest:

    discover  -> read teams (Teams-enabled M365 groups) and their channels
    plan      -> match each team to its target M365 group by mailNickname;
                 classify the team and each channel (read-only)
    migrate   -> enable Teams on the matching target group and recreate its
                 standard channels

Because a team rides on its M365 group, run ``groups sync`` first so the target
group exists to be Teams-enabled. Only **standard** channels are recreated; the
default *General* channel is created automatically with the team, and private /
shared channels (which need channel-scoped membership) are surfaced as ``skip``
for later work.
"""

from __future__ import annotations

from m365_migrate.graph_client import GraphClient, GraphError
from m365_migrate.models import (
    TEAM_GROUP_SELECT_FIELDS,
    Channel,
    PlannedChannel,
    PlannedTeam,
    SourceTeam,
    group_is_team,
)
from m365_migrate.workloads.groups import get_target_group_ids


def discover_teams(client: GraphClient) -> list[SourceTeam]:
    """Read all teams from the source tenant, including their channels."""
    select = ",".join(TEAM_GROUP_SELECT_FIELDS)
    raw = client.get_all("/groups", params={"$select": select, "$top": 999})
    teams: list[SourceTeam] = []
    for group in raw:
        if not group_is_team(group):
            continue
        channels_raw = client.get_all(f"/teams/{group['id']}/channels", params={"$top": 999})
        channels = [Channel.from_graph(c) for c in channels_raw]
        teams.append(SourceTeam.from_graph(group, channels))
    return teams


def get_target_team_nicknames(client: GraphClient) -> set[str]:
    """Return lowercased mailNicknames of target groups that are already teams."""
    raw = client.get_all(
        "/groups",
        params={"$select": "mailNickname,resourceProvisioningOptions", "$top": 999},
    )
    return {
        g["mailNickname"].lower()
        for g in raw
        if g.get("mailNickname") and group_is_team(g)
    }


def _plan_channel(channel: Channel) -> PlannedChannel:
    """Classify a single source channel for recreation."""
    if channel.is_default:
        action, reason = "skip", "default channel created automatically with the team"
    elif channel.membership_type != "standard":
        action, reason = "skip", f"{channel.membership_type} channel not yet supported"
    else:
        action, reason = "create", None
    return PlannedChannel(
        display_name=channel.display_name,
        description=channel.description,
        membership_type=channel.membership_type,
        action=action,
        reason=reason,
    )


def plan_teams(
    teams: list[SourceTeam],
    existing_target_groups: dict[str, str] | None = None,
) -> list[PlannedTeam]:
    """Build a team provisioning plan without writing anything.

    ``existing_target_groups`` maps lowercased target mailNickname -> group id. A
    team is ``provision`` only when its backing M365 group already exists in the
    target (otherwise ``skip`` until ``groups sync`` has run).
    """
    existing = existing_target_groups or {}
    planned: list[PlannedTeam] = []
    for team in teams:
        nickname = team.mail_nickname
        if not nickname:
            action, reason = "skip", "team has no mailNickname"
        elif nickname.lower() not in existing:
            action, reason = "skip", "target M365 group missing — run groups sync first"
        else:
            action, reason = "provision", None

        planned.append(
            PlannedTeam(
                source_id=team.id,
                mail_nickname=nickname,
                display_name=team.display_name,
                action=action,
                reason=reason,
                channels=[_plan_channel(c) for c in team.channels],
            )
        )
    return planned


def migrate_teams(
    client: GraphClient,
    planned: list[PlannedTeam],
    *,
    dry_run: bool = True,
) -> list[dict]:
    """Enable Teams on matching target groups and recreate standard channels.

    For a team, a Teams-enabled group has the same id as its team, so the group id
    doubles as the team id once enabled. With ``dry_run=True`` (default) nothing is
    written; planned actions are reported with a ``would-`` prefix.
    """
    target_groups = get_target_group_ids(client)
    target_teams = get_target_team_nicknames(client)
    results: list[dict] = []

    for p in planned:
        record: dict = {"mail_nickname": p.mail_nickname}

        if p.action == "skip":
            record["status"] = f"skipped:{p.reason}"
            results.append(record)
            continue

        nickname = (p.mail_nickname or "").lower()
        team_id = target_groups.get(nickname)
        if not team_id:
            record["status"] = "skipped:target group missing"
            results.append(record)
            continue

        already_team = nickname in target_teams

        # --- enable Teams on the group (team id == group id) ---
        if already_team:
            record["team"] = "exists"
        elif dry_run:
            record["team"] = "would-enable"
        else:
            try:
                client.put(f"/groups/{team_id}/team", json={})
                record["team"] = "enabled"
                already_team = True
            except GraphError as exc:
                record["team"] = f"error:{exc.status_code}"
                record["status"] = "error"
                results.append(record)
                continue

        # --- recreate standard channels ---
        to_create = [c for c in p.channels if c.action == "create"]
        if not to_create:
            record["channels"] = "none"
            record["status"] = "ok"
            results.append(record)
            continue

        if dry_run:
            record["channels"] = f"would-create:{len(to_create)}"
            record["status"] = "ok"
            results.append(record)
            continue

        # Skip channels that already exist on the target team (idempotent re-runs).
        existing_names = {
            c["displayName"].lower()
            for c in client.get_all(f"/teams/{team_id}/channels", params={"$top": 999})
            if c.get("displayName")
        }
        created, errors = 0, 0
        for channel in to_create:
            if channel.display_name.lower() in existing_names:
                continue
            try:
                client.post(
                    f"/teams/{team_id}/channels",
                    json={
                        "displayName": channel.display_name,
                        "description": channel.description,
                        "membershipType": "standard",
                    },
                )
                created += 1
            except GraphError:
                errors += 1
        record["channels"] = f"created:{created}"
        if errors:
            record["channels_errors"] = errors
        record["status"] = "ok"
        results.append(record)

    return results
