"""Shared test fixtures."""

import pytest

from m365_migrate.config import Config, Options, TenantConfig


@pytest.fixture
def config() -> Config:
    return Config(
        source=TenantConfig(
            tenant_id="src-tenant",
            client_id="src-client",
            client_secret="src-secret",
            primary_domain="contoso.onmicrosoft.com",
        ),
        target=TenantConfig(
            tenant_id="tgt-tenant",
            client_id="tgt-client",
            client_secret="tgt-secret",
            primary_domain="fabrikam.onmicrosoft.com",
        ),
        options=Options(rewrite_upn_domain=True, skip_guests=True, output_dir="out"),
    )


@pytest.fixture
def static_token():
    """A token provider that returns a fixed fake token (no network)."""
    return lambda: "fake-token"
