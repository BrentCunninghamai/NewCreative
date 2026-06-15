"""Configuration loading and validation.

Config lives in a YAML file (see ``config.example.yaml``). Secret fields may use
the ``${ENV:VAR_NAME}`` syntax to pull values from the environment at load time,
which keeps credentials out of the file on disk.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

import yaml
from pydantic import BaseModel, Field

_ENV_PATTERN = re.compile(r"^\$\{ENV:([A-Z0-9_]+)\}$")


class TenantConfig(BaseModel):
    """Credentials and identity for a single Microsoft 365 tenant."""

    tenant_id: str
    client_id: str
    client_secret: str
    primary_domain: str


class Options(BaseModel):
    """Behavioral options for migration runs."""

    rewrite_upn_domain: bool = True
    skip_guests: bool = True
    output_dir: str = "out"


class Config(BaseModel):
    """Top-level configuration: a source tenant, a target tenant, and options."""

    source: TenantConfig
    target: TenantConfig
    options: Options = Field(default_factory=Options)


def _resolve_env(value: Any) -> Any:
    """Replace ``${ENV:NAME}`` strings with the corresponding env var value."""
    if isinstance(value, str):
        match = _ENV_PATTERN.match(value.strip())
        if match:
            name = match.group(1)
            resolved = os.environ.get(name)
            if resolved is None:
                raise ConfigError(
                    f"Environment variable {name!r} referenced in config is not set."
                )
            return resolved
        return value
    if isinstance(value, dict):
        return {k: _resolve_env(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_resolve_env(v) for v in value]
    return value


class ConfigError(Exception):
    """Raised when configuration is missing or invalid."""


def load_config(path: str | Path) -> Config:
    """Load, env-resolve, and validate a config file into a :class:`Config`."""
    path = Path(path)
    if not path.exists():
        raise ConfigError(
            f"Config file not found: {path}. "
            "Copy config.example.yaml to config.yaml and fill it in."
        )
    raw = yaml.safe_load(path.read_text()) or {}
    resolved = _resolve_env(raw)
    return Config.model_validate(resolved)
