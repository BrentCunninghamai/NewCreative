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
