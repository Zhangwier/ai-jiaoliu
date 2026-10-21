# -*- coding: utf-8 -*-
"""渲染：python rr.py <outdir> [pages]"""
import sys, os, subprocess, json, shutil
ROOT = r"D:\MG\program\ai交流\_v36"
NODE = r"D:\Software\Node\node.exe"
CLI = r"C:\Users\MG\.dsh\profiles\web\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js"
pptx = os.path.join(ROOT, "build.pptx")
pdf = os.path.join(ROOT, "build.pdf")
out = os.path.join(ROOT, sys.argv[1] if len(sys.argv) > 1 else "r")
pages = sys.argv[2] if len(sys.argv) > 2 else None
for p in (pdf,):
    if os.path.exists(p):
        os.remove(p)
if os.path.isdir(out):
    shutil.rmtree(out)
r = subprocess.run([NODE, CLI, "convert", "--input", pptx, "--output", pdf],
                   capture_output=True, text=True, encoding="utf-8", errors="replace")
print(r.returncode, (r.stdout or "")[-400:], (r.stderr or "")[-400:])
cmd = [NODE, CLI, "render", "--input", pdf, "--output-dir", out, "--dpi", "110"]
if pages:
    cmd += ["--pages", pages]
r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
print(r.returncode, (r.stdout or "")[-600:], (r.stderr or "")[-300:])
