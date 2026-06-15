"""Domain models shared across workloads."""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel

# The Graph user properties we read from the source tenant. Keeping this list
# explicit (rather than ``$select=*``) keeps payloads small and predictable.
USER_SELECT_FIELDS = [
    "id",
    "userPrincipalName",
    "displayName",
    "givenName",
    "surname",
    "mail",
    "jobTitle",
    "department",
    "officeLocation",
    "mobilePhone",
    "accountEnabled",
    "userType",
    "usageLocation",
    "assignedLicenses",
]

# Expanded alongside the user so we learn each user's manager in one request.
USER_EXPAND = "manager($select=id,userPrincipalName)"


class SourceUser(BaseModel):
    """A user as read from the source tenant."""

    id: str
    user_principal_name: str
    display_name: str | None = None
    given_name: str | None = None
    surname: str | None = None
    mail: str | None = None
    job_title: str | None = None
    department: str | None = None
    office_location: str | None = None
    mobile_phone: str | None = None
    account_enabled: bool = True
    user_type: str = "Member"
    usage_location: str | None = None
    # UPN of this user's manager in the source tenant (from $expand=manager).
    manager_upn: str | None = None
    # License SKU IDs assigned to the user in the source tenant.
    assigned_sku_ids: list[str] = []

    @classmethod
    def from_graph(cls, data: dict[str, Any]) -> "SourceUser":
        manager = data.get("manager") or {}
        licenses = data.get("assignedLicenses") or []
        return cls(
            id=data["id"],
            user_principal_name=data["userPrincipalName"],
            display_name=data.get("displayName"),
            given_name=data.get("givenName"),
            surname=data.get("surname"),
            mail=data.get("mail"),
            job_title=data.get("jobTitle"),
            department=data.get("department"),
            office_location=data.get("officeLocation"),
            mobile_phone=data.get("mobilePhone"),
            account_enabled=data.get("accountEnabled", True),
            user_type=data.get("userType") or "Member",
            usage_location=data.get("usageLocation"),
            manager_upn=manager.get("userPrincipalName"),
            assigned_sku_ids=[lic["skuId"] for lic in licenses if lic.get("skuId")],
        )


class PlannedUser(BaseModel):
    """A user creation planned for the target tenant.

    ``action`` is one of: ``create`` (new user), ``skip`` (filtered out), or
    ``conflict`` (a user with the target UPN already exists in the target).
    """

    source_id: str
    source_upn: str
    target_upn: str
    display_name: str | None
    action: str
    reason: str | None = None

    def to_graph_body(self, *, default_usage_location: str | None = None) -> dict[str, Any]:
        """Build the Graph ``POST /users`` request body for this planned user."""
        mail_nickname = self.target_upn.split("@", 1)[0]
        body: dict[str, Any] = {
            "accountEnabled": True,
            "displayName": self.display_name or mail_nickname,
            "mailNickname": mail_nickname,
            "userPrincipalName": self.target_upn,
            "passwordProfile": {
                "forceChangePasswordNextSignIn": True,
                # Caller is expected to inject a real generated password before
                # sending; this placeholder is never sent as-is by the workload.
                "password": "REPLACE_ME",
            },
        }
        if default_usage_location:
            body["usageLocation"] = default_usage_location
        return body


# The Graph group properties we read from the source tenant.
GROUP_SELECT_FIELDS = [
    "id",
    "displayName",
    "mailNickname",
    "description",
    "groupTypes",
    "securityEnabled",
    "mailEnabled",
    "visibility",
]

# Expanded alongside each group so we learn its user members in one request.
GROUP_EXPAND = "members($select=id,userPrincipalName)"


class SourceGroup(BaseModel):
    """A group as read from the source tenant, with its user members."""

    id: str
    display_name: str | None = None
    mail_nickname: str | None = None
    description: str | None = None
    group_types: list[str] = []
    security_enabled: bool = False
    mail_enabled: bool = False
    visibility: str | None = None
    # UPNs of user members (non-user directory objects are ignored).
    member_upns: list[str] = []

    @property
    def is_unified(self) -> bool:
        """True for Microsoft 365 ("Unified") groups."""
        return any(t.lower() == "unified" for t in self.group_types)

    @property
    def kind(self) -> str:
        """Classify the group for migration handling.

        Returns one of ``microsoft365``, ``security``, ``mail-enabled-security``,
        ``distribution``, or ``unknown``. Only ``microsoft365`` and ``security``
        groups can be provisioned through Graph; the mail-enabled kinds require
        Exchange Online and are handled by a later workload.
        """
        if self.is_unified:
            return "microsoft365"
        if self.security_enabled and self.mail_enabled:
            return "mail-enabled-security"
        if self.security_enabled:
            return "security"
        if self.mail_enabled:
            return "distribution"
        return "unknown"

    @classmethod
    def from_graph(cls, data: dict[str, Any]) -> "SourceGroup":
        members = data.get("members") or []
        return cls(
            id=data["id"],
            display_name=data.get("displayName"),
            mail_nickname=data.get("mailNickname"),
            description=data.get("description"),
            group_types=data.get("groupTypes") or [],
            security_enabled=data.get("securityEnabled", False),
            mail_enabled=data.get("mailEnabled", False),
            visibility=data.get("visibility"),
            member_upns=[
                m["userPrincipalName"]
                for m in members
                if m.get("userPrincipalName")
            ],
        )


class PlannedGroup(BaseModel):
    """A group reconciliation planned for the target tenant.

    ``action`` is one of: ``create`` (provision a new target group), ``exists``
    (a matching target group is already present, just sync membership), or
    ``skip`` (the group kind cannot be provisioned via Graph).
    """

    source_id: str
    mail_nickname: str | None
    display_name: str | None
    kind: str
    action: str
    reason: str | None = None
    description: str | None = None
    # Member UPNs already rewritten to the target domain.
    target_member_upns: list[str] = []

    def to_graph_body(self) -> dict[str, Any]:
        """Build the Graph ``POST /groups`` request body for this planned group."""
        body: dict[str, Any] = {
            "displayName": self.display_name or self.mail_nickname,
            "mailNickname": self.mail_nickname,
            "description": self.description,
        }
        if self.kind == "microsoft365":
            body.update(groupTypes=["Unified"], mailEnabled=True, securityEnabled=False)
        else:  # security
            body.update(groupTypes=[], mailEnabled=False, securityEnabled=True)
        # Drop keys Graph would reject as null.
        return {k: v for k, v in body.items() if v is not None}


# The writable subset of Graph ``mailboxSettings``. ``userPurpose`` and similar
# read-only fields are intentionally excluded so they are never PATCHed back.
MAILBOX_SETTABLE_FIELDS = [
    "automaticRepliesSetting",
    "timeZone",
    "language",
    "workingHours",
    "dateFormat",
    "timeFormat",
    "delegateMeetingMessageDeliveryOptions",
]


class SourceMailbox(BaseModel):
    """A user's mailbox configuration as read from the source tenant.

    Only the writable ``mailboxSettings`` subset is captured; mailbox *content*
    (mail, calendar, contacts) is out of scope for the Graph layer and requires a
    native cross-tenant mailbox move.
    """

    user_principal_name: str
    settings: dict[str, Any] = {}

    @classmethod
    def from_graph(cls, upn: str, data: dict[str, Any]) -> "SourceMailbox":
        settings = {
            field: data[field]
            for field in MAILBOX_SETTABLE_FIELDS
            if data.get(field) is not None
        }
        return cls(user_principal_name=upn, settings=settings)


class PlannedMailbox(BaseModel):
    """A mailbox-settings migration planned for the target tenant.

    ``action`` is one of: ``settings`` (apply settings to the target mailbox) or
    ``skip`` (no target mailbox, or nothing to migrate).
    """

    source_upn: str
    target_upn: str
    action: str
    reason: str | None = None
    settings: dict[str, Any] = {}
