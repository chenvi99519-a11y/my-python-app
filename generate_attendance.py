# -*- coding: utf-8 -*-
#!/usr/bin/env python3
"""从“汇总.xlsx”自动生成各车队电子考勤表（兼容报表型表头）。"""

import argparse
from pathlib import Path


def _lazy_import_pandas():
    try:
        import pandas as pd  # type: ignore
    except ModuleNotFoundError as e:
        raise SystemExit("缺少依赖: pandas。请先执行: pip install pandas openpyxl") from e
    return pd


def _norm_text(v):
    return str(v).strip().replace(" ", "").replace("\n", "")


def _dedup_columns(cols):
    seen = {}
    out = []
    for c in cols:
        c = _norm_text(c)
        if c == "" or c.lower().startswith("unnamed"):
            c = "空列"
        n = seen.get(c, 0)
        seen[c] = n + 1
        out.append(c if n == 0 else f"{c}_{n+1}")
    return out


def _find_col(df, keywords, required=False):
    cols = list(df.columns)
    norm_map = {_norm_text(c).lower(): c for c in cols}
    for key in keywords:
        key_n = _norm_text(key).lower()
        for c_n, original in norm_map.items():
            if key_n in c_n:
                return original
    if required:
        raise KeyError(f"未找到列，关键词={list(keywords)}，可用列={cols}")
    return None


def _ensure_str_id(series):
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def _to_num(pd, s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _best_header_row(raw_df, keywords, scan_rows=20):
    best_idx = None
    best_score = -1
    for i in range(min(scan_rows, len(raw_df))):
        row_vals = [_norm_text(v).lower() for v in raw_df.iloc[i].tolist()]
        score = 0
        for kw in keywords:
            kw_n = _norm_text(kw).lower()
            if any(kw_n in cell for cell in row_vals):
                score += 1
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx if best_score > 0 else None


def _smart_read_sheet(path, sheet_name, header_keywords):
    pd = _lazy_import_pandas()
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    hdr = _best_header_row(raw, header_keywords, scan_rows=30)

    if hdr is None:
        # 回退默认读取
        return pd.read_excel(path, sheet_name=sheet_name)

    header = _dedup_columns(raw.iloc[hdr].tolist())
    body = raw.iloc[hdr + 1 :].copy()
    body.columns = header
    body = body.dropna(axis=1, how="all")
    body = body.dropna(axis=0, how="all")
    return body


def load_source_tables(path):
    pd = _lazy_import_pandas()
    xls = pd.ExcelFile(path)
    required = ["人员信息", "缺勤统计表", "电子加班表", "其他人员中夜班加班"]
    missing = [s for s in required if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"输入文件缺少工作表: {missing}；实际有: {xls.sheet_names}")

    tables = {
        "人员信息": _smart_read_sheet(path, "人员信息", ["职号", "工号", "姓名", "工种代码", "部门", "车队"]),
        "缺勤统计表": _smart_read_sheet(path, "缺勤统计表", ["职号", "工号", "姓名", "病", "事", "待岗"]),
        "电子加班表": _smart_read_sheet(path, "电子加班表", ["职号", "工号", "加班", "国假", "节日", "天数"]),
        "其他人员中夜班加班": _smart_read_sheet(path, "其他人员中夜班加班", ["职号", "工号", "中班", "夜班", "加班"]),
    }
    return tables


def build_attendance_table(source_tables, team_name):
    pd = _lazy_import_pandas()
    df_people = source_tables["人员信息"].copy()
    df_abs = source_tables["缺勤统计表"].copy()
    df_drv = source_tables["电子加班表"].copy()
    df_oth = source_tables["其他人员中夜班加班"].copy()

    p_id = _find_col(df_people, ["职号", "工号"], required=True)
    p_name = _find_col(df_people, ["姓名"], required=True)
    p_team = _find_col(df_people, ["部门名称", "车队", "部门"], required=True)
    p_job = _find_col(df_people, ["工种代码", "工种"], required=True)

    a_id = _find_col(df_abs, ["职号", "工号"], required=True)
    a_team = _find_col(df_abs, ["车队", "部门名称", "部门"], required=False)
    a_sick = _find_col(df_abs, ["病"], required=False)
    a_person = _find_col(df_abs, ["事"], required=False)
    a_wait = _find_col(df_abs, ["待岗"], required=False)
    a_note = _find_col(df_abs, ["备注"], required=False)

    d_id = _find_col(df_drv, ["职号", "工号"], required=True)
    d_ot = _find_col(df_drv, ["延时", "加班天数", "加班"], required=False)
    d_holiday = _find_col(df_drv, ["节日", "国假"], required=False)

    o_id = _find_col(df_oth, ["职号", "工号"], required=True)
    o_mid = _find_col(df_oth, ["中班"], required=False)
    o_night = _find_col(df_oth, ["夜班"], required=False)
    o_ot = _find_col(df_oth, ["公休", "加班"], required=False)
    o_delay = _find_col(df_oth, ["延时"], required=False)
    o_holiday = _find_col(df_oth, ["节日"], required=False)

    for d, c in [(df_people, p_id), (df_abs, a_id), (df_drv, d_id), (df_oth, o_id)]:
        d[c] = _ensure_str_id(d[c])

    base = df_people[df_people[p_team].astype(str).str.strip() == str(team_name).strip()][[p_id, p_name, p_team, p_job]].copy()
    base.columns = ["职号", "姓名", "车队", "工种代码"]

    # 缺勤：若缺勤表无车队列，则直接按职号合并
    abs_df = df_abs.copy()
    if a_team:
        abs_df = abs_df[abs_df[a_team].astype(str).str.strip() == str(team_name).strip()]

    keep = {a_id: "职号"}
    if a_sick:
        keep[a_sick] = "病假"
    if a_person:
        keep[a_person] = "事假"
    if a_wait:
        keep[a_wait] = "待岗"
    if a_note:
        keep[a_note] = "备注"
    abs_t = abs_df[list(keep.keys())].rename(columns=keep).drop_duplicates(subset=["职号"], keep="last")

    drv_keep = {d_id: "职号"}
    if d_ot:
        drv_keep[d_ot] = "驾驶员加班"
    if d_holiday:
        drv_keep[d_holiday] = "驾驶员节日加班"
    drv_t = df_drv[list(drv_keep.keys())].rename(columns=drv_keep).drop_duplicates(subset=["职号"], keep="last")

    oth_keep = {o_id: "职号"}
    if o_mid:
        oth_keep[o_mid] = "中班"
    if o_night:
        oth_keep[o_night] = "夜班"
    if o_ot:
        oth_keep[o_ot] = "其他加班"
    if o_delay:
        oth_keep[o_delay] = "其他延时加班"
    if o_holiday:
        oth_keep[o_holiday] = "其他节日加班"
    oth_t = df_oth[list(oth_keep.keys())].rename(columns=oth_keep).drop_duplicates(subset=["职号"], keep="last")

    out = base.merge(abs_t, on="职号", how="left").merge(drv_t, on="职号", how="left").merge(oth_t, on="职号", how="left")

    out["是否驾驶员"] = out["工种代码"].astype(str).str.strip().eq("301")
    drv_cols = [c for c in ["驾驶员加班", "驾驶员节日加班"] if c in out.columns]
    oth_cols = [c for c in ["中班", "夜班", "其他加班", "其他延时加班", "其他节日加班"] if c in out.columns]
    if drv_cols:
        out.loc[~out["是否驾驶员"], drv_cols] = 0
    if oth_cols:
        out.loc[out["是否驾驶员"], oth_cols] = 0

    for c in ["病假", "事假", "待岗", *drv_cols, *oth_cols]:
        if c in out.columns:
            out[c] = _to_num(pd, out[c])
    if "备注" in out.columns:
        out["备注"] = out["备注"].fillna("").astype(str)

    out.insert(0, "序", range(1, len(out) + 1))
    out["延时加班"] = (out["驾驶员加班"] if "驾驶员加班" in out.columns else 0) + (out["其他延时加班"] if "其他延时加班" in out.columns else 0)
    out["节日加班"] = (out["驾驶员节日加班"] if "驾驶员节日加班" in out.columns else 0) + (out["其他节日加班"] if "其他节日加班" in out.columns else 0)
    out["公休加班"] = out["其他加班"] if "其他加班" in out.columns else 0
    out["中班"] = out["中班"] if "中班" in out.columns else 0
    out["夜班"] = out["夜班"] if "夜班" in out.columns else 0
    out["应出勤天数"] = 18
    out["加班合计"] = out["延时加班"] + out["节日加班"] + out["公休加班"]

    ordered = ["序", "职号", "姓名", "车队", "工种代码", "是否驾驶员", "应出勤天数", "病假", "事假", "待岗", "中班", "夜班", "公休加班", "延时加班", "节日加班", "加班合计", "备注"]
    ordered = [c for c in ordered if c in out.columns]
    return out[ordered]


def build_all_teams(source_tables):
    df_people = source_tables["人员信息"]
    p_team = _find_col(df_people, ["部门名称", "车队", "部门"], required=True)
    teams = df_people[p_team].dropna().astype(str).str.strip()
    teams = teams[teams != ""].drop_duplicates().tolist()

    result = {}
    for t in teams:
        d = build_attendance_table(source_tables, t)
        if len(d) > 0:
            result[str(t)] = d
    return result


def export_workbook(team_tables, out_path):
    pd = _lazy_import_pandas()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        idx = []
        for team, df in team_tables.items():
            name = str(team)[:31] if str(team).strip() else "未命名车队"
            df.to_excel(writer, sheet_name=name, index=False)
            idx.append({"车队": team, "人数": len(df)})
        pd.DataFrame(idx).to_excel(writer, sheet_name="汇总索引", index=False)


def main():
    parser = argparse.ArgumentParser(description="由汇总.xlsx生成各车队电子考勤表（兼容标题行）")
    parser.add_argument("input", type=Path, nargs="?", help="输入Excel路径")
    parser.add_argument("-o", "--output", type=Path, default=Path("电子考勤汇总_自动生成.xlsx"), help="输出Excel路径")
    parser.add_argument("--team", type=str, default=None, help="仅生成指定车队")
    args = parser.parse_args()

    if args.input is None:
        parser.print_help()
        return

    tables = load_source_tables(args.input)
    if args.team:
        team_tables = {args.team: build_attendance_table(tables, args.team)}
    else:
        team_tables = build_all_teams(tables)

    if not team_tables:
        raise SystemExit("没有可导出的车队数据，请检查输入数据与表头。")

    export_workbook(team_tables, args.output)
    print(f"已生成: {args.output}")
    print(f"车队数: {len(team_tables)}")


if __name__ == "__main__":
    main()
