#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""接单交付工具箱 · deliver.py

闲鱼/猪八戒 Excel 自动化接单专用：一条命令出结果，交付成本压到接近 0。
零第三方依赖（复用 excel_merge_kit / excel_pivot_kit / excel_clean_kit 自研引擎）。

命令：
  info   in                      数据体检（接单前估价用）
  merge  a b c... -o out.xlsx    多表合并（可去重、带来源溯源）
  clean  in -o out.xlsx          清洗（去空格/空行/去重/填充）
  pivot  in -o out.xlsx --by 区域 --agg 金额:sum:金额合计
  split  in --by 城市 --outdir d 按列拆分成多个文件
  recon  A B --key 订单号 --outdir d  双表核对（对账专用，输出差异）

用法示例见每个子命令的 --help。
"""
import argparse
import os
import re
import sys
from typing import Dict, List, Optional, Sequence

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
# 两种布局都要能跑：① 引擎与本脚本同级（独立仓库）② 引擎在本仓库根目录的 kit 里
for _p in (HERE, ROOT,
           os.path.join(ROOT, "excel_merge_kit"),
           os.path.join(ROOT, "excel_pivot_kit"),
           os.path.join(ROOT, "excel_clean_kit")):
    if os.path.isdir(_p) and _p not in sys.path:
        sys.path.insert(0, _p)

try:
    import merge_engine as ME
    import pivot_engine as PE
    import clean_engine as CE
except ImportError as exc:  # pragma: no cover
    sys.exit("引擎导入失败，请确认 excel_*_kit 目录在本仓库根目录：%s" % exc)

AGGS = ("sum", "count", "avg", "max", "min")


# ---------- 通用小工具 ----------
def _num(s):
    """把字符串转成数字；转不了返回 None。支持 ¥ , % 千分位。"""
    if s is None:
        return None
    t = str(s).strip().replace(",", "").replace("¥", "").replace("￥", "")
    t = t.replace("%", "")
    if t in ("", "-", "--", "N/A", "NA", "null", "None"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def _cols(rows: Sequence[Dict]) -> List[str]:
    """保序列出所有出现过的列名。"""
    out: List[str] = []
    for r in rows:
        for k in r.keys():
            if k not in out:
                out.append(k)
    return out


def _safe(name: str) -> str:
    """文件名安全化（Windows 非法字符 + 空值）。"""
    t = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(name).strip())
    t = t.strip().rstrip(".")
    return t or "未分类"


def _ext(path: str) -> str:
    return os.path.splitext(path)[1].lower()


def _read(path: str) -> List[Dict[str, str]]:
    if not os.path.exists(path):
        sys.exit("文件不存在：%s" % path)
    try:
        return ME.read_any(path)
    except Exception as exc:
        sys.exit("读取失败 %s：%s" % (path, exc))


def _write(rows: Sequence[Dict], path: str, columns: Optional[Sequence[str]] = None) -> int:
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    if _ext(path) == ".csv":
        return ME.write_csv(rows, path, columns=columns)
    return ME.write_xlsx(rows, path, columns=columns)


def _fmt_rows(rows: Sequence[Dict], limit: int = 8) -> str:
    if not rows:
        return "  （空）"
    cols = _cols(rows)[:8]
    lines = ["  " + " | ".join(cols)]
    lines.append("  " + "-" * min(72, 12 * len(cols)))
    for r in rows[:limit]:
        lines.append("  " + " | ".join(str(r.get(c, ""))[:18] for c in cols))
    if len(rows) > limit:
        lines.append("  ... 共 %d 行（仅显示前 %d 行）" % (len(rows), limit))
    return "\n".join(lines)


# ---------- 子命令 ----------
def cmd_info(args):
    rows = _read(args.input)
    cols = _cols(rows)
    print("文件：%s" % os.path.abspath(args.input))
    print("行数：%d    列数：%d" % (len(rows), len(cols)))
    print("列名：%s" % ", ".join(cols))
    print()
    print("列体检（空值 / 重复值 / 是否数值列）：")
    for c in cols:
        vals = [str(r.get(c, "")).strip() for r in rows]
        blank = sum(1 for v in vals if v == "")
        nonempty = [v for v in vals if v != ""]
        dup = len(nonempty) - len(set(nonempty))
        numeric = sum(1 for v in nonempty if _num(v) is not None)
        isnum = "数值" if nonempty and numeric == len(nonempty) else ("部分数值" if numeric else "文本")
        print("  %-16s 空 %4d  重复 %4d  %s" % (c[:16], blank, dup, isnum))
    print()
    print("前几行预览：")
    print(_fmt_rows(rows))
    print()
    print("估价参考：行数 %d → 建议报价见 QUOTE.md" % len(rows))
    return 0


def cmd_merge(args):
    for p in args.inputs:
        if not os.path.exists(p):
            sys.exit("文件不存在：%s" % p)
    res = ME.merge(args.inputs, key=args.key, dedup=args.dedup)
    rows = res["rows"]
    columns = res["columns"] + (["__source__"] if "__source__" in _cols(rows) else [])
    n = _write(rows, args.output, columns=columns)
    print("合并完成：%d 个文件 → %d 行" % (len(args.inputs), len(rows)))
    print("输出：%s" % os.path.abspath(args.output))
    if args.dedup:
        print("去重键：%s" % (args.key or "（未指定，按整行）"))
    return 0


def cmd_clean(args):
    rows = _read(args.input)
    keys = [k.strip() for k in args.key.split(",")] if args.key else None
    out = CE.clean(
        rows,
        trim=not args.no_trim,
        drop_blanks=not args.keep_blanks,
        dedup=args.dedup,
        dedup_keys=keys,
        fill=args.fill,
    )
    n = _write(out, args.output)
    print("清洗完成：%d 行 → %d 行（减少 %d 行）" % (len(rows), len(out), len(rows) - len(out)))
    print("输出：%s" % os.path.abspath(args.output))
    return 0


def cmd_pivot(args):
    rows = _read(args.input)
    group_by = [g.strip() for g in args.by.split(",")]
    aggs = []
    for spec in args.agg:
        parts = spec.split(":")
        if len(parts) < 2:
            sys.exit("--agg 格式应为 列名:func[:别名]，例如 金额:sum:金额合计；收到：%s" % spec)
        col, func = parts[0].strip(), parts[1].strip()
        if func not in AGGS:
            sys.exit("未知聚合 %s，可选：%s" % (func, "/".join(AGGS)))
        aggs.append({"col": col, "func": func, "as": parts[2].strip() if len(parts) > 2 else None})
    for c in group_by + [a["col"] for a in aggs]:
        if c not in _cols(rows):
            print("警告：列「%s」不在文件里（将按空值处理）。实际列：%s"
                  % (c, ", ".join(_cols(rows))))
    res = PE.pivot(rows, group_by=group_by, aggs=aggs)
    n = _write(res["summary"], args.output, columns=res["group_cols"] + res["value_cols"])
    print("透视完成：%d 行 → %d 组" % (res["stats"]["input_rows"], res["stats"]["groups"]))
    print("输出：%s" % os.path.abspath(args.output))
    print()
    print(_fmt_rows(res["summary"], limit=15))
    return 0


def cmd_split(args):
    rows = _read(args.input)
    if args.by not in _cols(rows):
        sys.exit("列「%s」不存在。实际列：%s" % (args.by, ", ".join(_cols(rows))))
    buckets: Dict[str, List[Dict]] = {}
    for r in rows:
        buckets.setdefault(_safe(r.get(args.by, "")), []).append(r)
    outdir = os.path.abspath(args.outdir)
    os.makedirs(outdir, exist_ok=True)
    total = 0
    for name, group in sorted(buckets.items()):
        path = os.path.join(outdir, "%s.%s" % (name, "xlsx" if args.xlsx else "csv"))
        _write(group, path)
        total += 1
        print("  %-24s %4d 行 -> %s" % (name[:24], len(group), os.path.basename(path)))
    print()
    print("拆分完成：%d 行 → %d 个文件，目录：%s" % (len(rows), total, outdir))
    return 0


def cmd_recon(args):
    a = _read(args.a)
    b = _read(args.b)
    for tag, rows in (("A", a), ("B", b)):
        if args.key not in _cols(rows):
            sys.exit("表%s 缺少主键列「%s」。实际列：%s"
                     % (tag, args.key, ", ".join(_cols(rows))))
    da: Dict[str, Dict] = {}
    for r in a:
        da[str(r.get(args.key, "")).strip()] = r
    db: Dict[str, Dict] = {}
    for r in b:
        db[str(r.get(args.key, "")).strip()] = r

    only_a = [da[k] for k in da if k and k not in db]
    only_b = [db[k] for k in db if k and k not in da]
    common = [k for k in da if k and k in db]

    compare_cols = [c.strip() for c in args.cols.split(",")] if args.cols else None
    if compare_cols is None:
        compare_cols = [c for c in _cols(a) if c in _cols(b) and c != args.key]

    diffs = []
    for k in common:
        ra, rb = da[k], db[k]
        changed = []
        for c in compare_cols:
            va, vb = str(ra.get(c, "")).strip(), str(rb.get(c, "")).strip()
            na, nb = _num(va), _num(vb)
            if na is not None and nb is not None:
                same = abs(na - nb) < 1e-9
            else:
                same = va == vb
            if not same:
                changed.append({"主键": k, "列": c, "表A": va, "表B": vb})
        if changed:
            diffs.extend(changed)

    outdir = os.path.abspath(args.outdir)
    os.makedirs(outdir, exist_ok=True)
    ext = "xlsx" if args.xlsx else "csv"
    p1 = os.path.join(outdir, "1_仅在表A.%s" % ext)
    p2 = os.path.join(outdir, "2_仅在表B.%s" % ext)
    p3 = os.path.join(outdir, "3_数值不一致.%s" % ext)
    _write(only_a, p1)
    _write(only_b, p2)
    _write(diffs, p3, columns=["主键", "列", "表A", "表B"])

    print("核对完成（主键：%s）" % args.key)
    print("  表A %d 行 / 表B %d 行" % (len(a), len(b)))
    print("  一致（主键匹配）   : %d" % len(common))
    print("  仅在表A            : %d  -> %s" % (len(only_a), os.path.basename(p1)))
    print("  仅在表B            : %d  -> %s" % (len(only_b), os.path.basename(p2)))
    print("  数值/内容不一致    : %d  -> %s" % (len(diffs), os.path.basename(p3)))
    print()
    print("输出目录：%s" % outdir)
    if diffs:
        print()
        print("不一致明细（前 10 条）：")
        print(_fmt_rows(diffs, limit=10))
    return 0


# ---------- CLI ----------
def build():
    ap = argparse.ArgumentParser(
        prog="deliver.py",
        description="接单交付工具箱：Excel 合并 / 清洗 / 透视 / 拆分 / 核对（零依赖）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = ap.add_subparsers(dest="cmd")

    p = sub.add_parser("info", help="数据体检（接单前估价用）")
    p.add_argument("input")
    p.set_defaults(func=cmd_info)

    p = sub.add_parser("merge", help="多表合并")
    p.add_argument("inputs", nargs="+")
    p.add_argument("-o", "--output", default="合并结果.xlsx")
    p.add_argument("--key", default=None, help="去重主键列名")
    p.add_argument("--dedup", action="store_true", help="按主键去重（保留最新）")
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser("clean", help="清洗")
    p.add_argument("input")
    p.add_argument("-o", "--output", default="清洗结果.xlsx")
    p.add_argument("--key", default=None, help="去重依据列，逗号分隔")
    p.add_argument("--dedup", action="store_true")
    p.add_argument("--fill", default=None, help="空值填充为指定文本")
    p.add_argument("--no-trim", action="store_true", help="不去首尾空格")
    p.add_argument("--keep-blanks", action="store_true", help="保留全空行")
    p.set_defaults(func=cmd_clean)

    p = sub.add_parser("pivot", help="透视汇总")
    p.add_argument("input")
    p.add_argument("-o", "--output", default="透视汇总.xlsx")
    p.add_argument("--by", required=True, help="分组列，逗号分隔，如 区域")
    p.add_argument("--agg", nargs="+", required=True,
                   help="聚合，格式 列名:func[:别名]，如 金额:sum:金额合计 金额:count:笔数")
    p.set_defaults(func=cmd_pivot)

    p = sub.add_parser("split", help="按列拆分成多个文件")
    p.add_argument("input")
    p.add_argument("--by", required=True, help="拆分依据列，如 城市")
    p.add_argument("--outdir", default="拆分结果")
    p.add_argument("--xlsx", action="store_true", help="输出 xlsx（默认 csv）")
    p.set_defaults(func=cmd_split)

    p = sub.add_parser("recon", help="双表核对（对账）")
    p.add_argument("a")
    p.add_argument("b")
    p.add_argument("--key", required=True, help="主键列，如 订单号")
    p.add_argument("--cols", default=None, help="要比对的列，逗号分隔；默认全部同名列")
    p.add_argument("--outdir", default="核对结果")
    p.add_argument("--xlsx", action="store_true", help="输出 xlsx（默认 csv）")
    p.set_defaults(func=cmd_recon)

    return ap


def main(argv=None):
    ap = build()
    args = ap.parse_args(argv)
    if not getattr(args, "cmd", None):
        ap.print_help()
        return 1
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
