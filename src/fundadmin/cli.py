"""FundAdministration 统一 CLI 入口。

用途:
- clients：客户/净值/持仓导入（typer 子命令）。
- portfolio：券商持仓抓取 / 汇总 / 客户净值通知（透传给 argparse 实现 operations.main）。

调用入口:
- python -m fundadmin.cli --help
- python -m fundadmin.cli clients init-db
- python -m fundadmin.cli portfolio email-sync --help
"""

from __future__ import annotations

import typer

from fundadmin import __version__
from fundadmin.clients.cli import clients_app

app = typer.Typer(help="FundAdministration 基金运营 CLI", no_args_is_help=True)
app.add_typer(clients_app, name="clients")


@app.command(
    "portfolio",
    context_settings={
        "allow_extra_args": True,
        "ignore_unknown_options": True,
        "help_option_names": [],
    },
)
def portfolio(ctx: typer.Context) -> None:
    """券商持仓抓取 / 汇总 / 客户净值通知（透传 operations.main）。

    子命令：email-sync / sync-latest / build-products / build-cross-broker /
    prune-inbox / notify-clients。用 `portfolio --help` 查看 argparse 帮助。
    """
    from fundadmin.portfolio.operations import main as ops_main

    raise typer.Exit(code=int(ops_main(list(ctx.args))))


@app.command("version")
def version() -> None:
    """打印版本号。"""
    typer.echo(f"fundadmin {__version__}")


def main() -> None:
    app()


if __name__ == "__main__":
    main()
