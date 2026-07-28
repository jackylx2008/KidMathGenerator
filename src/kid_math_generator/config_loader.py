"""项目配置加载与环境变量展开。"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

import yaml

from logging_config import get_cloudstation_root

_ENV_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)(?::-([^}]*))?\}")


def load_config(
    config_path: str | Path,
    *,
    env_path: str | Path | None = None,
) -> dict[str, Any]:
    """读取 YAML 配置，并递归展开其中的环境变量标记。"""
    path = Path(config_path).resolve()
    local_env_path = Path(env_path).resolve() if env_path else path.with_name("common.env")
    load_env_file(local_env_path)
    os.environ.setdefault("CLOUDSTATION_ROOT", get_cloudstation_root())

    with path.open("r", encoding="utf-8") as file:
        loaded = yaml.safe_load(file) or {}

    if not isinstance(loaded, dict):
        raise ValueError(f"配置文件顶层必须是映射: {path}")

    return _expand_value(loaded)


def load_env_file(path: str | Path) -> None:
    """把本地环境文件注入进程环境，已有环境变量优先。"""
    env_file = Path(path)
    if not env_file.is_file():
        return

    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        name = name.strip()
        value = value.strip().strip("\"'")
        if name:
            os.environ.setdefault(name, value)


def _expand_value(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: _expand_value(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_expand_value(item) for item in value]
    if isinstance(value, str):
        expanded = _ENV_PATTERN.sub(_replace_env_marker, value)
        return str(Path(expanded).expanduser()) if expanded.startswith("~") else expanded
    return value


def _replace_env_marker(match: re.Match[str]) -> str:
    name, default = match.groups()
    value = os.getenv(name)
    if value not in (None, ""):
        return value
    if default is not None:
        return default
    raise ValueError(f"缺少环境变量: {name}")
