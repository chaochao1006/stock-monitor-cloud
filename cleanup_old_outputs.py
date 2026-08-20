from __future__ import annotations

import os
from datetime import datetime, timedelta
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("REPORT_BASE_DIR", PROJECT_DIR / "data"))

WORD_RETENTION_DAYS = 30
LOG_RETENTION_DAYS = 3


def delete_older_than(pattern: str, cutoff: datetime) -> list[Path]:
    """删除 data 目录下早于 cutoff 的指定类型文件。"""
    deleted: list[Path] = []
    if not DATA_DIR.exists():
        return deleted

    for path in DATA_DIR.rglob(pattern):
        if not path.is_file():
            continue
        modified = datetime.fromtimestamp(path.stat().st_mtime)
        if modified < cutoff:
            path.unlink()
            deleted.append(path)
    return deleted


def main() -> None:
    """清理云端输出，避免 GitHub 仓库长期无限变大。"""
    now = datetime.now()
    word_cutoff = now - timedelta(days=WORD_RETENTION_DAYS)
    log_cutoff = now - timedelta(days=LOG_RETENTION_DAYS)

    deleted_word = delete_older_than("*.docx", word_cutoff)
    deleted_logs = delete_older_than("*.log", log_cutoff)

    print(f"Word 报告保留最近 {WORD_RETENTION_DAYS} 天，已删除 {len(deleted_word)} 个旧文件。")
    for path in deleted_word:
        print(f"  deleted word: {path.relative_to(DATA_DIR)}")

    print(f"日志保留最近 {LOG_RETENTION_DAYS} 天，已删除 {len(deleted_logs)} 个旧文件。")
    for path in deleted_logs:
        print(f"  deleted log: {path.relative_to(DATA_DIR)}")

    print("Excel 和触发记录 TXT 不参与清理。")


if __name__ == "__main__":
    main()
