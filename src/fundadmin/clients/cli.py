"""客户/净值/持仓 CLI 子命令。

调用入口:
- python -m fundadmin.cli clients init-db
- python -m fundadmin.cli clients import-clients --csv PATH
- python -m fundadmin.cli clients import-nav --xlsx PATH [--as-of DATE]
- python -m fundadmin.cli clients import-portfolio --xlsx PATH [--as-of DATE]

输出:
- stdout 打印每个子任务的写入行数。

失败行为:
- 路径或必需列不存在时抛 typer.BadParameter / FileNotFoundError，返回非零退出码。
"""

from __future__ import annotations

from pathlib import Path

import typer

clients_app = typer.Typer(help="客户台账 / 产品净值 / 持仓导入", no_args_is_help=True)


@clients_app.command("init-db")
def cli_init_db() -> None:
    """创建/补齐全部 SQLite 表。"""
    from fundadmin.clients.schema import init_db

    init_db()
    typer.echo("fundadmin: schema initialised")


@clients_app.command("import-clients")
def cli_import_clients(
    csv: Path = typer.Option(
        ..., "--csv", help="客户清单 CSV 路径", exists=True, dir_okay=False
    ),
) -> None:
    """导入客户主表（custname/prodcode/holding_shares/email/mobile）。"""
    from fundadmin.clients.ingest.client_nav import import_clients_csv

    rows = import_clients_csv(csv)
    typer.echo(f"clients: upserted {rows} rows")


@clients_app.command("import-nav")
def cli_import_nav(
    xlsx: Path = typer.Option(
        ..., "--xlsx", help="fund_portfolio_summary_<date>.xlsx 路径", exists=True, dir_okay=False
    ),
    as_of: str | None = typer.Option(None, "--as-of", help="估值日；缺省按文件名推断或当天"),
) -> None:
    """导入产品净值。"""
    from fundadmin.clients.ingest.client_nav import import_nav_xlsx

    rows = import_nav_xlsx(xlsx, as_of_date=as_of)
    typer.echo(f"product_nav_history: upserted {rows} rows")


@clients_app.command("import-portfolio")
def cli_import_portfolio(
    xlsx: Path = typer.Option(
        ..., "--xlsx", help="持仓模板 Excel 路径", exists=True, dir_okay=False
    ),
    as_of: str | None = typer.Option(None, "--as-of", help="行内 as_of_date 缺失时的默认估值日"),
) -> None:
    """导入产品持仓快照。"""
    from fundadmin.clients.ingest.portfolio_excel import import_portfolio_xlsx

    rows = import_portfolio_xlsx(xlsx, as_of_override=as_of)
    typer.echo(f"fund_portfolio_holdings: upserted {rows} rows")
