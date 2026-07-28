"""入口层、编排层与模块层共享的应用上下文。"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True, slots=True)
class AppContext:
    """一次工作流运行所需的公共配置与路径。"""

    project_root: Path
    config: dict[str, Any]
    flow_name: str
    logger: logging.Logger

    @property
    def app_config(self) -> dict[str, Any]:
        return dict(self.config.get("app", {}))

    @property
    def flow_config(self) -> dict[str, Any]:
        flows = self.config.get("flows", {})
        flow_config = flows.get(self.flow_name)
        if not isinstance(flow_config, dict):
            raise KeyError(f"未找到工作流配置: flows.{self.flow_name}")
        return dict(flow_config)

    @property
    def output_dir(self) -> Path:
        configured = Path(self.app_config.get("output_dir", "output"))
        if configured.is_absolute():
            return configured
        return self.project_root / configured
