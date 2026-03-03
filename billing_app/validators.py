from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd


@dataclass
class ValidationExceptionTables:
    missing_system_id_new: pd.DataFrame
    missing_system_id_prev: pd.DataFrame
    duplicate_system_id_new: pd.DataFrame
    duplicate_system_id_prev: pd.DataFrame
    multi_source_duplicates: pd.DataFrame


def _norm(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def validate_required_positions(df: pd.DataFrame, required_positions_1_based: list[int], label: str) -> list[str]:
    issues: list[str] = []
    for pos in required_positions_1_based:
        idx = pos - 1
        if idx >= df.shape[1]:
            issues.append(f"{label}: missing required column position {pos}.")
    return issues


def build_exception_tables(new_master: pd.DataFrame, prev_master: pd.DataFrame) -> ValidationExceptionTables:
    new_missing = new_master.loc[new_master["system_id"].map(_norm).eq("")].copy()
    prev_missing = prev_master.loc[prev_master["system_id"].map(_norm).eq("")].copy()

    new_dup_mask = new_master["system_id"].map(_norm).ne("") & new_master["system_id"].map(_norm).duplicated(keep=False)
    prev_dup_mask = prev_master["system_id"].map(_norm).ne("") & prev_master["system_id"].map(_norm).duplicated(keep=False)

    new_dups = new_master.loc[new_dup_mask].sort_values("system_id")
    prev_dups = prev_master.loc[prev_dup_mask].sort_values("system_id")

    by_source = (
        new_master.loc[new_master["system_id"].map(_norm).ne(""), ["system_id", "source"]]
        .drop_duplicates()
        .groupby("system_id")["source"]
        .nunique()
    )
    multi_source_ids = by_source.loc[by_source > 1].index
    multi_source = new_master.loc[new_master["system_id"].isin(multi_source_ids), ["system_id", "property_name", "source"]]

    return ValidationExceptionTables(
        missing_system_id_new=new_missing,
        missing_system_id_prev=prev_missing,
        duplicate_system_id_new=new_dups,
        duplicate_system_id_prev=prev_dups,
        multi_source_duplicates=multi_source,
    )


def summarize_exceptions(tables: ValidationExceptionTables) -> list[str]:
    warnings: list[str] = []
    if not tables.missing_system_id_new.empty:
        warnings.append(f"New master rows missing system_id: {len(tables.missing_system_id_new)}")
    if not tables.missing_system_id_prev.empty:
        warnings.append(f"Prev master rows missing system_id: {len(tables.missing_system_id_prev)}")
    if not tables.duplicate_system_id_new.empty:
        warnings.append(f"Duplicate system_id in new master: {tables.duplicate_system_id_new['system_id'].nunique()}")
    if not tables.duplicate_system_id_prev.empty:
        warnings.append(f"Duplicate system_id in prev master: {tables.duplicate_system_id_prev['system_id'].nunique()}")
    if not tables.multi_source_duplicates.empty:
        warnings.append(
            f"system_id values found in multiple sources: {tables.multi_source_duplicates['system_id'].nunique()}"
        )
    return warnings
