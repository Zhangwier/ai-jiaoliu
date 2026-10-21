# -*- coding: utf-8 -*-
"""一步：构建 -> 转 PDF -> 渲染 PNG。用法: python rr.py [outdir] [pages]"""
import sys, os, subprocess, shutil, glob
HERE = os.path.dirname(os.path.abspath(__file__))
NODE = r"D:\Software\Node\node.exe"
CLI  = r"C:\Users\MG\.dsh\profiles\web\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js"
PPTX = os.path.join(HERE, "build.pptx")
PDF  = os.path.join(HERE, "build.pdf")
OUT  = sys.argv[1] if len(sys.argv) > 1 else "rr"
PAGES = sys.argv[2] if len(sys.argv) > 2 else None

if os.path.exists(PDF):
    os.remove(PDF)
r = subprocess.run([NODE, CLI, "convert", "--input", PPTX, "--output", PDF],
                   capture_output=True, text=True)
if not os.path.exists(PDF):
    print("CONVERT FAILED", r.stdout[-2000:], r.stderr[-2000:]); sys.exit(1)
od = os.path.join(HERE, OUT)
if os.path.isdir(od):
    shutil.rmtree(od)
cmd = [NODE, CLI, "render", "--input", PDF, "--output-dir", od, "--dpi", "110"]
if PAGES:
    cmd += ["--pages", PAGES]
r = subprocess.run(cmd, capture_output=True, text=True)
print(r.stdout[-800:])
if r.returncode != 0:
    print(r.stderr[-2000:])
