from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any

import pandas as pd


@dataclass
class BillingResult:
    new_master: pd.DataFrame
    prev_master: pd.DataFrame
    new_master_out: pd.DataFrame
    prev_master_out: pd.DataFrame
    abstracted: pd.DataFrame
    terminated: pd.DataFrame
    abstractions_2col: pd.DataFrame
    terminations_2col: pd.DataFrame
    counts: dict[str, int]


def _norm_system_id(value: Any) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".", 1)[0]
    return s


def _pick_col(df: pd.DataFrame, idx: int) -> pd.Series:
    if idx >= df.shape[1]:
        return pd.Series([None] * len(df))
    return df.iloc[:, idx]


def _header_id_hint(header: str) -> float:
    h = header.strip().lower()
    score = 0.0
    if any(k in h for k in (" id", "id", "lease", "ref")):
        score += 1.0
    if "property" in h and "name" not in h:
        score += 0.5
    if any(k in h for k in ("name", "description", "address", "abstracted", "terminated")):
        score -= 1.0
    return score


def _id_likeness_score(series: pd.Series) -> float:
    s = series.map(_norm_system_id)
    s = s.loc[s.ne("")]
    if s.empty:
        return -1.0
    numeric_ratio = s.str.match(r"^\d+(\.0+)?$", na=False).mean()
    prefix_ratio = s.str.match(r"^(15|18|24)\d*$", na=False).mean()
    return float(numeric_ratio + (2.0 * prefix_ratio))


def _maybe_swap_system_and_property(
    property_name: pd.Series,
    system_id: pd.Series,
) -> tuple[pd.Series, pd.Series]:
    sys_score = _id_likeness_score(system_id)
    prop_score = _id_likeness_score(property_name)
    if prop_score > sys_score:
        return system_id.astype(str), property_name.map(_norm_system_id)
    return property_name.astype(str), system_id.map(_norm_system_id)


def build_new_master(pay_df: pd.DataFrame, rec_df: pd.DataFrame, free_df: pd.DataFrame) -> pd.DataFrame:
    # Explicit keys by source:
    # - Lease/Sublease (Payable, Receivable): Lease ID
    # - Freehold: Property Ref
    pay_system = _pick_col(pay_df, 0)
    pay_property = _pick_col(pay_df, 7)
    rec_system = _pick_col(rec_df, 0)
    rec_property = _pick_col(rec_df, 1)
    free_system = _pick_col(free_df, 0)
    free_property = _pick_col(free_df, 1)

    pay = pd.DataFrame(
        {
            "property_name": pay_property,
            "system_id_raw": pay_system,
            "source": "Payable",
        }
    )

    rec = pd.DataFrame(
        {
            "property_name": rec_property,
            "system_id_raw": rec_system,
            "source": "Receivable",
        }
    )

    free_type = _pick_col(free_df, 4)
    free_filter = free_type.fillna("").astype(str).eq("Freehold")
    free = pd.DataFrame(
        {
            "property_name": free_property,
            "system_id_raw": free_system,
            "source": "Freehold",
        }
    )
    free = free.loc[free_filter].copy()

    combined = pd.concat([pay, rec, free], ignore_index=True)
    combined["property_name"] = combined["property_name"].fillna("").astype(str)
    combined["system_id"] = combined["system_id_raw"].map(_norm_system_id)

    return combined[["property_name", "system_id", "system_id_raw", "source"]]


def build_prev_master(prev_df: pd.DataFrame) -> pd.DataFrame:
    n_cols = prev_df.shape[1]
    if n_cols == 0:
        return pd.DataFrame({"property_name": [], "system_id": []})

    headers = [str(c).strip().lower() for c in prev_df.columns]

    explicit_system_idx: int | None = None
    for i, h in enumerate(headers):
        if "system_id" in h or "system id" in h or "lease id" in h:
            explicit_system_idx = i
            break

    candidate_idxs = list(range(min(5, n_cols)))
    if explicit_system_idx is None:
        scored: list[tuple[float, int]] = []
        for i in candidate_idxs:
            col = _pick_col(prev_df, i)
            score = _id_likeness_score(col) + _header_id_hint(headers[i])
            scored.append((score, i))
        scored.sort(reverse=True)
        system_idx = scored[0][1]
    else:
        system_idx = explicit_system_idx

    explicit_prop_idx: int | None = None
    for i, h in enumerate(headers):
        if i == system_idx:
            continue
        if "property" in h or "description" in h:
            explicit_prop_idx = i
            break

    if explicit_prop_idx is not None:
        property_idx = explicit_prop_idx
    else:
        fallback = [i for i in candidate_idxs if i != system_idx]
        property_idx = fallback[0] if fallback else system_idx

    system_id = _pick_col(prev_df, system_idx).map(_norm_system_id)
    property_name = _pick_col(prev_df, property_idx).fillna("").astype(str)
    property_name = property_name.where(
        property_name.str.strip().ne(""),
        _pick_col(prev_df, system_idx).fillna("").astype(str),
    )
    property_name, system_id = _maybe_swap_system_and_property(property_name, system_id)

    return pd.DataFrame(
        {
            "property_name": property_name,
            "system_id": system_id,
        }
    )


def build_diffs(new_master: pd.DataFrame, prev_master: pd.DataFrame) -> BillingResult:
    new_ids = set(new_master["system_id"].loc[new_master["system_id"].ne("")])
    prev_ids = set(prev_master["system_id"].loc[prev_master["system_id"].ne("")])

    abstracted = new_master.loc[~new_master["system_id"].isin(prev_ids) & new_master["system_id"].ne("")].copy()
    terminated = prev_master.loc[~prev_master["system_id"].isin(new_ids) & prev_master["system_id"].ne("")].copy()

    new_master_out = new_master[["property_name", "system_id"]].copy()
    new_master_out.insert(
        0,
        "abstracted_flag",
        new_master_out["system_id"].isin(abstracted["system_id"]).map(lambda v: "Abstracted" if v else ""),
    )

    prev_master_out = prev_master[["property_name", "system_id"]].copy()
    prev_master_out.insert(
        0,
        "terminated_flag",
        prev_master_out["system_id"].isin(terminated["system_id"]).map(lambda v: "Terminated" if v else ""),
    )

    count_prefix_24 = int(new_master["system_id"].str.startswith("24", na=False).sum())
    count_prefix_18 = int(new_master["system_id"].str.startswith("18", na=False).sum())
    count_prefix_15 = int(new_master["system_id"].str.startswith("15", na=False).sum())
    total_contracts = int(new_master.loc[new_master["system_id"].ne(""), "system_id"].nunique())
    term_count = int((prev_master_out["terminated_flag"] == "Terminated").sum())
    abs_count = int((new_master_out["abstracted_flag"] == "Abstracted").sum())

    abstractions_2col = abstracted[["system_id", "property_name"]].reset_index(drop=True)
    terminations_2col = terminated[["system_id", "property_name"]].reset_index(drop=True)

    return BillingResult(
        new_master=new_master,
        prev_master=prev_master,
        new_master_out=new_master_out,
        prev_master_out=prev_master_out,
        abstracted=abstracted,
        terminated=terminated,
        abstractions_2col=abstractions_2col,
        terminations_2col=terminations_2col,
        counts={
            "count_prefix_24": count_prefix_24,
            "count_prefix_18": count_prefix_18,
            "count_prefix_15": count_prefix_15,
            "total_contracts": total_contracts,
            "term_count": term_count,
            "abs_count": abs_count,
        },
    )


def _check_columns(df: pd.DataFrame, required_idx: list[int], name: str) -> list[str]:
    problems: list[str] = []
    for idx in required_idx:
        if idx >= df.shape[1]:
            problems.append(f"{name}: missing required column at position {idx + 1}")
    return problems


def run_sanity_checks(pay_df: pd.DataFrame, rec_df: pd.DataFrame, free_df: pd.DataFrame, prev_df: pd.DataFrame) -> list[str]:
    issues: list[str] = []
    issues.extend(_check_columns(pay_df, [0, 7], "Payable"))
    issues.extend(_check_columns(rec_df, [0, 1], "Receivable"))
    issues.extend(_check_columns(free_df, [0, 1, 4], "Freehold"))
    issues.extend(_check_columns(prev_df, [1], "Prev Master"))
    return issues
