"""项目配置加载与环境变量展开。"""

from __future__ import annotations

import os
import platform
import re
from pathlib import Path
from typing import Any

import yaml

_ENV_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)(?::-([^}]*))?\}")
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_CLOUDSTATION_ENV_BY_PLATFORM = {
    "Windows": "CLOUDSTATION_ROOT_WINDOWS",
    "Darwin": "CLOUDSTATION_ROOT_MACOS",
    "Linux": "CLOUDSTATION_ROOT_LINUX",
}
_CLOUDSTATION_DEFAULT_BY_PLATFORM = {
    "Windows": r"D:\CloudStaion",
    "Darwin": "~/SynologyDrive/",
    "Linux": "~/CloudStation",
}


def load_config(
    config_path: str | Path,
    *,
    env_path: str | Path | None = None,
) -> dict[str, Any]:
    """读取 YAML 配置，并递归展开其中的环境变量标记。"""
    path = Path(config_path).resolve()
    local_env_path = (
        Path(env_path).resolve()
        if env_path
        else _PROJECT_ROOT / "common.env"
    )
    load_env_file(local_env_path)
    if not os.getenv("CLOUDSTATION_ROOT"):
        os.environ["CLOUDSTATION_ROOT"] = get_cloudstation_root()

    with path.open("r", encoding="utf-8") as file:
        loaded = yaml.safe_load(file) or {}

    if not isinstance(loaded, dict):
        raise ValueError(f"配置文件顶层必须是映射: {path}")

    return _expand_value(loaded)


def get_cloudstation_root() -> str:
    """按显式配置、当前平台配置和平台默认值的顺序解析同步根目录。"""
    explicit_root = os.getenv("CLOUDSTATION_ROOT")
    if explicit_root:
        return str(Path(explicit_root).expanduser())

    system = platform.system()
    platform_env_name = _CLOUDSTATION_ENV_BY_PLATFORM.get(system)
    platform_root = os.getenv(platform_env_name) if platform_env_name else None
    if not platform_root:
        platform_root = _CLOUDSTATION_DEFAULT_BY_PLATFORM.get(system, "~/CloudStation")
    return str(Path(platform_root).expanduser())


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
