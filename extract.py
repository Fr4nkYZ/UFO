#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从指定 request.log 提取 step = N 的所有条目
使用方式:
    python extract.py logs/c4_2/request.log 10
"""

import argparse
import json
from pathlib import Path
from datetime import datetime


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Extract step-N entries from a request.log file"
    )
    # 位置参数：log_file, step
    p.add_argument("log_file", help="日志文件路径，如 logs/c4_2/request.log")
    p.add_argument("step", type=int, help="要提取的 step 值")
    return p.parse_args()


def extract_step(log_path: Path, step_val: int):
    """提取指定 step，并去掉 image_list 字段"""
    if not log_path.exists():
        raise FileNotFoundError(f"文件不存在: {log_path}")

    matches = []
    with log_path.open(encoding="utf-8") as fp:
        for ln, line in enumerate(fp, 1):
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except json.JSONDecodeError:
                continue

            if obj.get("step") == step_val:
                cleaned = obj.copy()
                cleaned.pop("image_list", None)
                matches.append({"line_number": ln, "data": cleaned})
    return matches


def main():
    args = parse_args()
    log_path = Path(args.log_file)
    step_val = args.step

    print(f"📖 正在读取 {log_path} ，提取 step={step_val} ...")
    try:
        items = extract_step(log_path, step_val)
    except Exception as exc:
        print(f"❌ 发生错误: {exc}")
        return

    print(f"🎯 共找到 {len(items)} 个条目")
    if not items:
        return

    # 打印
    print("\n" + "=" * 48)
    for idx, entry in enumerate(items, 1):
        print(f"\n--- 条目 {idx} (行号 {entry['line_number']}) ---")
        print(json.dumps(entry["data"], ensure_ascii=False, indent=2))

    # 保存
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = f"step{step_val}_extracted_{ts}.json"
    with open(out_file, "w", encoding="utf-8") as fp:
        json.dump([e["data"] for e in items], fp, ensure_ascii=False, indent=2)

    print(f"\n💾 结果已保存到: {out_file}")


if __name__ == "__main__":
    main()
