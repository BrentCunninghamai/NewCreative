import pytest

from m365_migrate.config import ConfigError, load_config

EXAMPLE = """
source:
  tenant_id: src
  client_id: cid
  client_secret: ${ENV:TEST_SRC_SECRET}
  primary_domain: contoso.onmicrosoft.com
target:
  tenant_id: tgt
  client_id: cid2
  client_secret: plain-secret
  primary_domain: fabrikam.onmicrosoft.com
options:
  rewrite_upn_domain: false
"""


def test_load_config_resolves_env(tmp_path, monkeypatch):
    monkeypatch.setenv("TEST_SRC_SECRET", "from-env")
    cfg_file = tmp_path / "config.yaml"
    cfg_file.write_text(EXAMPLE)

    cfg = load_config(cfg_file)

    assert cfg.source.client_secret == "from-env"
    assert cfg.target.client_secret == "plain-secret"
    assert cfg.options.rewrite_upn_domain is False
    # Defaulted option still present.
    assert cfg.options.skip_guests is True


def test_missing_env_raises(tmp_path, monkeypatch):
    monkeypatch.delenv("TEST_SRC_SECRET", raising=False)
    cfg_file = tmp_path / "config.yaml"
    cfg_file.write_text(EXAMPLE)
    with pytest.raises(ConfigError):
        load_config(cfg_file)


def test_missing_file_raises(tmp_path):
    with pytest.raises(ConfigError):
        load_config(tmp_path / "nope.yaml")
