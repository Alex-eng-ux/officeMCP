"""配置管理."""

import os
from pathlib import Path

from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """Office MCP Server 配置."""

    # 允许操作的目录列表，分号分隔；默认用户主目录
    allowed_dirs: str = os.environ.get("USERPROFILE", "")

    # 默认是否允许覆盖文件
    default_overwrite: bool = False

    # 编辑前是否自动备份
    backup_before_edit: bool = True

    # Office 应用是否可见（调试用）
    visible: bool = False

    class Config:
        env_prefix = "OFFICE_MCP_"
        case_sensitive = False


settings = Settings()


def get_allowed_dirs() -> list[Path]:
    """获取允许操作的目录列表."""
    dirs = []
    for d in settings.allowed_dirs.split(";"):
        d = d.strip()
        if d:
            # 展开环境变量
            d = os.path.expandvars(d)
            dirs.append(Path(d).resolve())
    return dirs
