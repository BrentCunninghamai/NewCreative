"""Authentication against Microsoft Graph.

Uses the OAuth2 client-credentials flow (app-only) via ``azure-identity``. Each
tenant authenticates with its own app registration. Tokens are cached and
refreshed by the underlying credential object.
"""

from __future__ import annotations

from typing import Callable

from azure.identity import ClientSecretCredential

from m365_migrate.config import TenantConfig

GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"

# A token provider is any zero-arg callable returning a bearer token string.
TokenProvider = Callable[[], str]


def build_token_provider(tenant: TenantConfig) -> TokenProvider:
    """Return a callable that yields a valid Graph access token for ``tenant``.

    The returned provider can be passed to :class:`~m365_migrate.graph_client.GraphClient`.
    Token acquisition/refresh and caching are handled by ``azure-identity``.
    """
    credential = ClientSecretCredential(
        tenant_id=tenant.tenant_id,
        client_id=tenant.client_id,
        client_secret=tenant.client_secret,
    )

    def provider() -> str:
        return credential.get_token(GRAPH_DEFAULT_SCOPE).token

    return provider
