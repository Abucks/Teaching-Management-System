# -*- coding: utf-8 -*-
"""校验 macOS CI 工作流 YAML 结构（Windows 上也能验证）"""
import sys
from pathlib import Path

import yaml

P = Path(__file__).resolve().parent / "build-macos.yml"
d = yaml.safe_load(P.read_text(encoding="utf-8"))
fail = []


def ok(cond, name, detail=""):
    print(("  [PASS] " if cond else "  [FAIL] ") + name + ("" if cond else f"  {detail}"))
    if not cond:
        fail.append(name)


print("YAML 解析成功")
ok(d.get("name") == "Build macOS DMG", "workflow 名称正确", str(d.get("name")))
triggers = d.get("on") or d.get(True) or {}
ok("workflow_dispatch" in triggers, "支持手动触发", str(triggers))
ok(any(k == "push" for k in triggers), "支持 tag 推送触发", str(list(triggers)))

jobs = d.get("jobs", {})
ok(len(jobs) == 1, "包含 1 个 job", str(list(jobs)))
job = list(jobs.values())[0]
matrix = job.get("strategy", {}).get("matrix", {}).get("include", [])
labels = [m.get("label") for m in matrix]
archs = [m.get("arch") for m in matrix]
ok(labels == ["arm64", "x86_64"], "矩阵含 arm64 + x86_64", str(labels))
ok(archs == ["arm64", "x86_64"], "架构参数正确", str(archs))
ok([m.get("runner") for m in matrix] == ["macos-14", "macos-13"], "runner 版本正确",
   str([m.get("runner") for m in matrix]))

steps = job.get("steps", [])
names = [s.get("name", "") for s in steps]
ok(len(steps) >= 5, "步骤数量充足", str(len(steps)))
ok(any("checkout" in str(s.get("uses", "")) for s in steps), "有 checkout 步骤")
ok(any("setup-python" in str(s.get("uses", "")) for s in steps), "有 setup-python 步骤")
ok(any("upload-artifact" in str(s.get("uses", "")) for s in steps), "有上传产物步骤")
build_step = next((s for s in steps if "构建" in s.get("name", "")), None)
ok(build_step is not None, "存在构建步骤")
ok(build_step and "build_macos.py" in build_step.get("run", ""), "调用 build_macos.py")
ok(build_step and "--skip-deps" in build_step.get("run", ""), "CI 中跳过依赖安装步骤")
ok(build_step and "matrix.arch" in build_step.get("run", ""), "使用矩阵架构参数")
up = next((s for s in steps if "upload-artifact" in str(s.get("uses", ""))), None)
ok(up and up.get("with", {}).get("path", "").endswith("*.dmg"), "上传路径为 *.dmg",
   str(up.get("with", {}).get("path") if up else None))
ok(up and up.get("with", {}).get("if-no-files-found") == "error",
   "无产物时报错（避免静默失败）")

print(f"\n汇总: FAIL={len(fail)}")
sys.exit(1 if fail else 0)
