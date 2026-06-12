"""跨券商持仓合并报表核心逻辑。

流程：
1. 加载 NAV（估值表）。
2. 分别加载中金、中信 USD、中信 HKD 持仓。
3. 按统一 ticker 合并、去重、求和。
4. 计算权重 = 个股市值 / NAV。
5. 格式化输出 Excel。
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

from fundadmin.portfolio.parsers.cicc import (
    parse_cicc_holdings,
    parse_cicc_valuation_navs,
    parse_cicc_valuation_xlsx_holdings,
)
from fundadmin.portfolio.parsers.citic import (
    parse_citic_fx,
    parse_citic_underlying,
    parse_citics_derivative_holdings,
)
from fundadmin.portfolio.parsers.common import clean_ticker
from fundadmin.portfolio.parsers.swhysc import (
    parse_swhysc_holdings,
    parse_swhysc_navs,
)
from fundadmin.portfolio.parsers.trades import (
    compute_position_cost_from_trades,
    parse_cicc_trades,
)
from fundadmin.portfolio.parsers.valuation_nav import parse_nav_from_valuation

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class CrossBrokerInput:
    """构建报表所需的输入文件集合。"""

    valuation_path: Path | None = None
    cicc_paths: list[Path] | None = None
    citic_usd_underlying_paths: list[Path] | None = None
    citic_usd_balance_paths: list[Path] | None = None
    citic_hkd_underlying_paths: list[Path] | None = None
    citic_hkd_balance_paths: list[Path] | None = None
    valuation_paths: list[Path] | None = None


def _ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _load_cicc_frames(paths: list[Path] | None) -> list[pd.DataFrame]:
    if not paths:
        return []
    frames: list[pd.DataFrame] = []
    for p in paths:
        if not p.exists():
            continue
        suffix = p.suffix.lower()
        if suffix == ".xls":
            # CICC .xls 估值表（如 SCD704/SNJ280）中的 3199 科目是收益互换会计记录，
            # 属于衍生品合约层面，并非底层个股持仓，不纳入合并。
            continue
        elif suffix == ".xlsx":
            # 中信证券代发的 CICC 场外衍生品估值表有 3 个 sheet（含 "互换标的信息"），
            # 与标准估值表格式（1 个 sheet）区分。
            # CICC 估值报告附件（如 2026-04-16弘运盛泰铂金2号私募证券投资基金.xlsx）
            # 含 "持仓" sheet，需单独解析。
            try:
                xl = pd.ExcelFile(p)
                sheet_names = xl.sheet_names
                is_derivative = len(sheet_names) == 3
                has_holdings_sheet = "持仓" in sheet_names
            except Exception:
                is_derivative = False
                has_holdings_sheet = False

            if is_derivative:
                try:
                    frames.append(parse_citics_derivative_holdings(p))
                except Exception:
                    # 如果解析失败（如 sheet 为空），回退到标准估值表解析
                    frames.append(parse_swhysc_holdings(p))
            elif has_holdings_sheet:
                try:
                    frames.append(parse_cicc_valuation_xlsx_holdings(p))
                except Exception:
                    frames.append(parse_swhysc_holdings(p))
            else:
                # CICC 的 .xlsx 估值表（如 SLL384/SXQ602/SQJ420 等）使用与申万相同的格式
                frames.append(parse_swhysc_holdings(p))
        else:
            frames.append(parse_cicc_holdings(p))
    return frames


def _load_citic_frames(
    underlying_paths: list[Path] | None,
    balance_paths: list[Path] | None,
) -> list[pd.DataFrame]:
    if not underlying_paths:
        return []

    # 预解析所有 balance 汇率
    fx_map: dict[str, float] = {}
    if balance_paths:
        for bp in balance_paths:
            if not bp.exists():
                continue
            try:
                # 对于 CITIC Statement Excel，Balance 在 "Balance" sheet
                sheet = "Balance" if bp.suffix.lower() in {".xlsx", ".xlsm", ".xls"} else 0
                fx_map[bp.stem] = parse_citic_fx(bp, sheet=sheet)
            except Exception:
                # 单个 Balance 文件解析失败时跳过它；保留其它文件继续。
                logger.warning("parse_citic_fx failed for %s; skipping", bp, exc_info=True)

    # 若只有一个有效汇率，当作全局汇率
    global_fx: float | None = None
    if len(fx_map) == 1:
        global_fx = list(fx_map.values())[0]

    frames: list[pd.DataFrame] = []
    for up in underlying_paths:
        if not up.exists():
            continue

        fx: float | None = global_fx
        if fx is None:
            # 按文件名相似度或同目录匹配
            up_stem = up.stem.lower()
            up_dir = up.parent
            for bp_stem, rate in fx_map.items():
                bp = next((p for p in balance_paths if p.stem == bp_stem), None)
                if bp is None:
                    continue
                # 同目录优先
                if bp.parent == up_dir:
                    fx = rate
                    break
                # 或文件名包含关系
                if bp_stem.lower() in up_stem or up_stem in bp_stem.lower():
                    fx = rate
                    break

        if fx is None:
            raise ValueError(f"无法为 {up} 找到匹配的 Balance/汇率文件")

        # 对于 CITIC Statement Excel，Underlying 在 "Underlying" sheet
        sheet = "Underlying" if up.suffix.lower() in {".xlsx", ".xlsm", ".xls"} else 0
        frames.append(parse_citic_underlying(up, fx, sheet=sheet))
    return frames


def _load_valuation_frames(paths: list[Path] | None) -> list[pd.DataFrame]:
    if not paths:
        return []
    frames: list[pd.DataFrame] = []
    for p in paths:
        if not p.exists():
            continue
        frames.append(parse_swhysc_holdings(p))
    return frames


def _cicc_cost_map(cicc_paths: list[Path] | None) -> dict[str, dict[str, Any]]:
    """从 CICC 全历史成交流水按移动加权法推算每个 ticker 的本地币持仓成本。

    CICC "持仓" sheet 无原生每股成本，需由"当日交易"成交流水推算。
    返回 {ticker_norm: {cost_price_local, cost_value_local, cost_ccy}}。
    """
    if not cicc_paths:
        return {}
    trade_frames: list[pd.DataFrame] = []
    for p in cicc_paths:
        if not p.exists() or p.suffix.lower() != ".xlsx":
            continue
        try:
            tx = parse_cicc_trades(p)
        except Exception:
            logger.debug("parse_cicc_trades failed for %s", p, exc_info=True)
            continue
        if not tx.empty:
            trade_frames.append(tx)
    if not trade_frames:
        return {}
    all_tx = pd.concat(trade_frames, ignore_index=True)
    # 跨文件去重：全历史在每日报告中重复出现，按业务键去重保留一条。
    dedup_keys = [
        "trade_date", "ticker", "direction", "open_close",
        "price_local", "quantity", "contract_no",
    ]
    present = [c for c in dedup_keys if c in all_tx.columns]
    all_tx = all_tx.drop_duplicates(subset=present, keep="first")
    cost = compute_position_cost_from_trades(all_tx)
    out: dict[str, dict[str, Any]] = {}
    for _, r in cost.iterrows():
        norm = clean_ticker(r["ticker"])
        out[norm] = {
            "cost_price_local": r.get("cost_price_local"),
            "cost_value_local": r.get("cost_value_local"),
            "cost_ccy": r.get("cost_ccy"),
        }
    return out


def build_cross_broker_report(
    *,
    trade_date: date,
    inputs: CrossBrokerInput,
    out_xlsx: Path,
) -> dict[str, Any]:
    """生成跨券商合并持仓报表。

    参数:
        trade_date: 交易日。
        inputs: 各类输入文件路径。
        out_xlsx: 输出 Excel 路径。

    返回:
        包含汇总信息的字典。
    """
    # 1. NAV（优先使用申万宏源估值表，其次 CICC 估值表，最后通用估值表）
    # 边界：每个文件解析的部分结果（仅 unit_nav 或仅 asset_nav）都保留，
    # 避免后续文件解析失败时把已找到的值重置为 None。
    unit_nav: float | None = None
    asset_nav: float | None = None
    nav_parse_attempts: list[str] = []

    def _update_navs(u: float | None, a: float | None) -> None:
        nonlocal unit_nav, asset_nav
        if u is not None and unit_nav is None:
            unit_nav = u
        if a is not None and asset_nav is None:
            asset_nav = a

    if inputs.valuation_paths:
        for p in inputs.valuation_paths:
            if p.exists():
                try:
                    u, a = parse_swhysc_navs(p)
                    _update_navs(u, a)
                    nav_parse_attempts.append(f"swhysc({p.name}) -> unit={u}, asset={a}")
                    if asset_nav is not None:
                        break
                except Exception as exc:
                    nav_parse_attempts.append(f"swhysc({p.name}) -> EXC: {exc!r}")

    if asset_nav is None and inputs.cicc_paths:
        for p in inputs.cicc_paths:
            if not p.exists():
                nav_parse_attempts.append(f"missing({p.name})")
                continue
            suffix = p.suffix.lower()
            try:
                if suffix == ".xls":
                    u, a = parse_cicc_valuation_navs(p)
                    nav_parse_attempts.append(f"cicc_xls({p.name}) -> unit={u}, asset={a}")
                elif suffix == ".xlsx":
                    # CICC 的 .xlsx 估值表使用与申万相同的 NAV 提取逻辑
                    u, a = parse_swhysc_navs(p)
                    nav_parse_attempts.append(f"swhysc_xlsx({p.name}) -> unit={u}, asset={a}")
                else:
                    continue
                _update_navs(u, a)
                if asset_nav is not None:
                    break
            except Exception as exc:
                nav_parse_attempts.append(f"{suffix}({p.name}) -> EXC: {exc!r}")

    if asset_nav is None and inputs.valuation_path and inputs.valuation_path.exists():
        try:
            asset_nav = parse_nav_from_valuation(inputs.valuation_path)
            nav_parse_attempts.append(
                f"legacy_valuation({inputs.valuation_path.name}) -> asset={asset_nav}"
            )
        except Exception as exc:
            nav_parse_attempts.append(
                f"legacy_valuation({inputs.valuation_path.name}) -> EXC: {exc!r}"
            )

    if asset_nav is None and unit_nav is None and nav_parse_attempts:
        # 失败兜底：把每次解析尝试摘要打到 logger，便于排查为何 NAV 缺失。
        logger.warning("NAV parsing failed; tried %d candidate(s):", len(nav_parse_attempts))
        for attempt in nav_parse_attempts:
            logger.warning("  - %s", attempt)

    # 用于权重计算的 NAV（优先资产净值，兜底单位净值）
    nav = asset_nav if asset_nav is not None else unit_nav

    # 2. 加载各账户持仓
    # CITIC Statement 和 CICC 场外衍生品估值表中的 "互换标的信息" 是同一批底层持仓的
    # 不同表示（前者是汇总视图，后者是合约级视图），避免重复计算。
    citic_frames = _load_citic_frames(inputs.citic_usd_underlying_paths, inputs.citic_usd_balance_paths)
    citic_frames += _load_citic_frames(inputs.citic_hkd_underlying_paths, inputs.citic_hkd_balance_paths)
    has_citic_holdings = any(not f.empty for f in citic_frames)

    cicc_paths = inputs.cicc_paths or []
    if has_citic_holdings:
        # 如果 CITIC Statement 已有持仓，过滤掉 CICC 场外衍生品估值表（3 个 sheet，
        # 由中信证券代发，与 CITIC Statement 底层标的数据重复）。
        # 保留 CICC 估值报告附件（含 "持仓" sheet）和标准估值表格式（SLL384/SXQ602/SQJ420 等），
        # 这些文件中的持仓与 CITIC Statement 互补，需要合并计算。
        filtered: list[Path] = []
        for p in cicc_paths:
            try:
                xl = pd.ExcelFile(p)
                is_derivative = len(xl.sheet_names) == 3
            except Exception:
                is_derivative = False
            if not is_derivative:
                filtered.append(p)
        cicc_paths = filtered

    frames: list[pd.DataFrame] = []
    frames.extend(_load_cicc_frames(cicc_paths))
    frames.extend(citic_frames)
    frames.extend(_load_valuation_frames(inputs.valuation_paths))

    valid_frames = [f for f in frames if not f.empty]
    if not valid_frames:
        raise ValueError("未找到任何有效持仓数据")

    holdings = pd.concat(valid_frames, ignore_index=True)

    # 补齐缺失的成本列（CICC 持仓 sheet 无原生成本，后续由成交流水回填）。
    for col in ("cost_price_local", "cost_value_local", "cost_ccy"):
        if col not in holdings.columns:
            holdings[col] = None
    # 成本金额转数值，确保 groupby 求和正确（None/空值视为缺失）。
    holdings["cost_value_local"] = pd.to_numeric(holdings["cost_value_local"], errors="coerce")

    # 3. 按 ticker 合并
    # 先标准化 ticker（去掉交易所后缀，如 TSLA.US -> TSLA，0286.HK -> 0286），
    # 确保不同来源的同一标的能够正确合并。
    # ticker 为空时用 company 兜底。
    holdings["ticker_norm"] = holdings["ticker"].map(clean_ticker)
    holdings["group_key"] = holdings["ticker_norm"].where(holdings["ticker_norm"].str.strip() != "", holdings["company"])

    def _first_non_null(s: pd.Series) -> Any:
        for v in s:
            if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip() != "":
                return v
        return None

    aggregated = (
        holdings.groupby("group_key", as_index=False)
        .agg(
            ticker=("ticker", "first"),
            company=("company", "first"),
            shares=("shares", "sum"),
            market_value_cny=("market_value_cny", "sum"),
            cost_value_local=("cost_value_local", "sum"),
            cost_ccy=("cost_ccy", _first_non_null),
            source_files=("source_file", lambda s: ",".join(s.drop_duplicates().tolist())),
        )
        .sort_values("market_value_cny", ascending=False, na_position="last")
        .reset_index(drop=True)
    )

    # 本地币每股成本 = 持仓成本加总 / 股数（成本看板使用 local currency）。
    # CICC 持仓无原生成本：用全历史成交流水按移动加权法回填。
    cicc_cost = _cicc_cost_map(cicc_paths)
    if cicc_cost:
        for i in aggregated.index:
            cv = aggregated.at[i, "cost_value_local"]
            if cv is not None and not (isinstance(cv, float) and pd.isna(cv)) and cv != 0:
                continue
            norm = clean_ticker(aggregated.at[i, "ticker"])
            info = cicc_cost.get(norm)
            if info:
                aggregated.at[i, "cost_value_local"] = info["cost_value_local"]
                aggregated.at[i, "cost_ccy"] = info["cost_ccy"]

    def _per_share_cost(row: pd.Series) -> Any:
        cv = row["cost_value_local"]
        sh = row["shares"]
        if cv is None or (isinstance(cv, float) and pd.isna(cv)) or cv == 0:
            return None
        if sh is None or (isinstance(sh, float) and pd.isna(sh)) or sh == 0:
            return None
        return cv / sh

    aggregated["cost_price_local"] = aggregated.apply(_per_share_cost, axis=1)

    # 4. 计算权重
    if nav and nav > 0:
        aggregated["weight"] = aggregated["market_value_cny"] / nav
    else:
        aggregated["weight"] = None

    # 保留原始数值 DataFrame，供可视化和通知模块使用
    aggregated_raw = aggregated.copy()

    # 5. 格式化输出
    fmt = aggregated.copy()
    fmt["shares"] = pd.to_numeric(fmt["shares"], errors="coerce").fillna(0).astype(int)
    fmt["market_value_cny"] = fmt["market_value_cny"].apply(
        lambda x: f"{x:,.2f}" if pd.notna(x) else ""
    )
    fmt["weight"] = fmt["weight"].apply(
        lambda x: f"{x:.2%}" if pd.notna(x) else ""
    )
    fmt["cost_price_local"] = fmt["cost_price_local"].apply(
        lambda x: f"{x:,.4f}" if pd.notna(x) else ""
    )
    fmt["cost_value_local"] = fmt["cost_value_local"].apply(
        lambda x: f"{x:,.2f}" if pd.notna(x) else ""
    )
    fmt["cost_ccy"] = fmt["cost_ccy"].fillna("")

    summary = pd.DataFrame(
        [
            {
                "trade_date": trade_date.isoformat(),
                "unit_nav": f"{unit_nav:,.4f}" if unit_nav is not None else "",
                "asset_nav": f"{asset_nav:,.2f}" if asset_nav is not None else "",
                "nav_for_weight": f"{nav:,.2f}" if nav is not None else "",
                "total_holdings": int(aggregated["group_key"].nunique()),
                "total_market_value_cny": f"{aggregated['market_value_cny'].sum():,.2f}" if aggregated["market_value_cny"].notna().any() else "",
            }
        ]
    )

    _ensure_dir(out_xlsx.parent)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        fmt.to_excel(writer, sheet_name="holdings", index=False)
        aggregated.to_excel(writer, sheet_name="holdings_raw", index=False)

    return {
        "trade_date": trade_date.isoformat(),
        "out_xlsx": str(out_xlsx),
        "unit_nav": unit_nav,
        "asset_nav": asset_nav,
        "nav": nav,
        "total_holdings": int(aggregated["group_key"].nunique()),
        "total_market_value_cny": float(aggregated["market_value_cny"].sum()) if aggregated["market_value_cny"].notna().any() else None,
        "holdings_raw": aggregated_raw,
    }


# ---- 按产品维度生成报表 ----

PRODUCT_CONFIG: list[dict[str, Any]] = [
    {
        "name": "全球视野",
        "citic_codes": ["104902"],
        "cicc_codes": ["SCD704"],
    },
    {
        "name": "种子",
        "citic_codes": ["105979"],
        "valuation_keywords": ["资产估值表"],
    },
    {
        "name": "铂金1号",
        "citic_codes": ["107244"],
        "cicc_codes": ["SLL384"],
    },
    {
        "name": "铂金2号",
        "citic_codes": ["111255"],
        "cicc_codes": ["SNJ280"],
    },
    {
        "name": "铂金8号",
        "citic_codes": ["115820"],
        "cicc_codes": ["SXQ602"],
    },
    {
        "name": "沐泽1号",
        "cicc_codes": ["SQJ420"],
    },
]


def _is_target_date(name: str, trade_date: date) -> bool:
    """检查文件名是否包含目标日期。"""
    date_str1 = trade_date.isoformat()  # 2026-04-16
    date_str2 = trade_date.strftime("%Y%m%d")  # 20260416
    return date_str1 in name or date_str2 in name


def _match_files_for_product(
    files: list[Path],
    cfg: dict[str, Any],
    trade_date: date,
) -> CrossBrokerInput:
    """根据产品配置从文件列表中匹配对应的持仓文件。"""
    citic_usd_u: list[Path] = []
    citic_usd_b: list[Path] = []
    citic_hkd_u: list[Path] = []
    citic_hkd_b: list[Path] = []
    cicc_paths: list[Path] = []
    valuation_paths_matched: list[Path] = []

    for f in files:
        name = f.name
        if not _is_target_date(name, trade_date):
            continue

        # CITIC
        for code in cfg.get("citic_codes", []):
            if code in name and "Statement" in name:
                if "USD" in name:
                    citic_usd_u.append(f)
                    citic_usd_b.append(f)
                elif "HKD" in name:
                    citic_hkd_u.append(f)
                    citic_hkd_b.append(f)

        # CICC 估值表（.xls 收益互换科目 或 .xlsx 标准估值表格式）
        for code in cfg.get("cicc_codes", []):
            if code in name and f.suffix.lower() in {".xls", ".xlsx"} and "Statement" not in name:
                cicc_paths.append(f)

        # 中信证券代发的 CICC 场外衍生品估值表（文件名含 citic_codes 如 104902/111255 等）
        for code in cfg.get("citic_codes", []):
            if code in name and f.suffix.lower() == ".xlsx" and "Statement" not in name and f not in cicc_paths:
                cicc_paths.append(f)

        # CICC 估值报告附件（文件名格式：{date}弘运盛泰{product_name}私募证券投资基金.xlsx）
        if cfg["name"] in name and "弘运盛泰" in name and f.suffix.lower() == ".xlsx" and "Statement" not in name and f not in cicc_paths and f not in valuation_paths_matched:
            cicc_paths.append(f)

        # 产品估值表（文件名包含关键词且为 xlsx）
        for kw in cfg.get("valuation_keywords", []):
            if kw in name and f.suffix.lower() == ".xlsx" and "Statement" not in name:
                valuation_paths_matched.append(f)

    # 对产品估值表去重并尝试筛选能解析出 NAV 的文件
    valid_valuation: list[Path] = []
    for p in sorted(set(valuation_paths_matched)):
        try:
            parse_swhysc_navs(p)
            valid_valuation.append(p)
        except Exception:
            # 解析失败说明该估值文件不属于本批，跳过即可；保留日志便于排查误匹配。
            logger.debug("parse_swhysc_navs probe failed for %s; skipping", p, exc_info=True)

    return CrossBrokerInput(
        cicc_paths=cicc_paths or None,
        citic_usd_underlying_paths=citic_usd_u or None,
        citic_usd_balance_paths=citic_usd_b or None,
        citic_hkd_underlying_paths=citic_hkd_u or None,
        citic_hkd_balance_paths=citic_hkd_b or None,
        valuation_paths=valid_valuation or None,
    )


def score_product_inputs_for_date(files: list[Path], trade_date: date) -> tuple[int, int]:
    """评估某个交易日的产品附件完整度。

    返回 (覆盖产品数, 覆盖来源数)，用于在同一收件目录存在多日附件时选择更完整的批次。
    """
    product_count = 0
    source_count = 0
    for cfg in PRODUCT_CONFIG:
        inputs = _match_files_for_product(files, cfg, trade_date)
        source_flags = [
            bool(inputs.cicc_paths),
            bool(inputs.citic_usd_underlying_paths),
            bool(inputs.citic_hkd_underlying_paths),
            bool(inputs.valuation_paths),
        ]
        if any(source_flags):
            product_count += 1
            source_count += sum(1 for flag in source_flags if flag)
    return product_count, source_count


def build_product_reports(
    *,
    trade_date: date,
    inbox_dir: Path,
    out_dir: Path,
) -> list[dict[str, Any]]:
    """按产品维度分别生成跨券商持仓报表。

    参数:
        trade_date: 交易日。
        inbox_dir: 下载附件目录（如 outputs/excels/fund_inbox/2026-04-16）。
        out_dir: 输出目录。

    返回:
        每个产品的汇总信息列表。
    """
    if not inbox_dir.exists():
        raise ValueError(f"inbox_dir not found: {inbox_dir}")

    # 收集该交易日所有文件
    files = [f for f in inbox_dir.iterdir() if f.is_file()]

    results: list[dict[str, Any]] = []
    for cfg in PRODUCT_CONFIG:
        pname = cfg["name"]
        inputs = _match_files_for_product(files, cfg, trade_date)

        # 检查是否有任何有效输入
        has_input = any([
            inputs.cicc_paths,
            inputs.citic_usd_underlying_paths,
            inputs.citic_hkd_underlying_paths,
            inputs.valuation_paths,
        ])
        if not has_input:
            logger.warning("%s: 未找到任何输入文件，跳过", pname)
            continue

        out_xlsx = out_dir / f"product_{pname}_{trade_date.isoformat()}.xlsx"
        try:
            payload = build_cross_broker_report(
                trade_date=trade_date,
                inputs=inputs,
                out_xlsx=out_xlsx,
            )
            payload["product_name"] = pname
            results.append(payload)
            logger.info("%s: %d holdings, NAV=%s", pname, payload["total_holdings"], payload["nav"])
        except ValueError as exc:
            if "未找到任何有效持仓数据" in str(exc):
                logger.warning("%s: 未找到有效持仓数据，跳过", pname)
                continue
            raise

    return results


def build_summary_excel(
    results: list[dict[str, Any]],
    *,
    trade_date: date,
    out_path: Path,
) -> Path:
    """生成产品汇总表格 Excel。

    参数:
        results: build_product_reports() 的返回列表。
        trade_date: 交易日。
        out_path: 输出文件路径。

    返回:
        写入后的文件路径。
    """
    rows: list[dict[str, Any]] = []
    for r in results:
        rows.append(
            {
                "trade_date": trade_date.isoformat(),
                "product_name": r.get("product_name", ""),
                "unit_nav": r.get("unit_nav"),
                "asset_nav": r.get("asset_nav"),
                "nav": r.get("nav"),
                "total_holdings": r.get("total_holdings"),
                "total_market_value_cny": r.get("total_market_value_cny"),
            }
        )

    df = pd.DataFrame(rows)
    _ensure_dir(out_path.parent)
    df.to_excel(out_path, index=False, engine="openpyxl")
    return out_path
