"""projectx 发布 CLI 子命令。

调用入口:
- python -m fundadmin.cli projectx init-db
- python -m fundadmin.cli projectx sync [--dates 2026-08-18,2026-08-19]
- python -m fundadmin.cli projectx verify

输出:
- stdout 打印每张表的同步统计 / 对账问题清单。

失败行为:
- PROJECTX_DB_URL 缺失时抛 RuntimeError；对账不一致时 verify 返回退出码 1。
"""

from __future__ import annotations

import typer

projectx_app = typer.Typer(help="projectx MySQL 发布（基金持仓/产品净值）", no_args_is_help=True)


@projectx_app.command("init-db")
def cli_init_db() -> None:
    """创建 projectx 基金表（幂等）。"""
    from .schema import init_db

    init_db()
    typer.echo("projectx: schema initialised")


@projectx_app.command("sync")
def cli_sync(
    dates: str | None = typer.Option(
        None, "--dates", help="仅同步指定估值日，逗号分隔（缺省全量镜像）"
    ),
    chunk: int = typer.Option(500, "--chunk", min=1, help="单批插入行数"),
) -> None:
    """把本地持仓/净值镜像到 projectx（先建表，再按估值日整快照替换）。"""
    from fundadmin.clients.schema import get_engine as local_engine

    from .schema import get_engine, init_db
    from .sync import sync_all

    target = get_engine()
    init_db(target)
    date_list = (
        [d.strip() for d in dates.split(",") if d.strip()] if dates else None
    )
    for stat in sync_all(source=local_engine(), target=target, dates=date_list, chunk=chunk):
        typer.echo(f"{stat['table']}: {stat['rows']} rows / {stat['dates']} dates")


@projectx_app.command("verify")
def cli_verify() -> None:
    """对账本地与 projectx 的行数一致性；不一致时退出码为 1。"""
    from fundadmin.clients.schema import get_engine as local_engine

    from .schema import get_engine
    from .sync import verify

    issues = verify(source=local_engine(), target=get_engine())
    if not issues:
        typer.echo("projectx: verify OK（持仓/净值与本地一致）")
        return
    for line in issues:
        typer.echo(f"[MISMATCH] {line}")
    raise typer.Exit(code=1)


__all__ = ["projectx_app"]
