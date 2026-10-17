# -*- coding: utf-8 -*-
"""统一视口 1200x680 截图 10 份鹈鹕结果，供同尺度比较（写进 _检查补充_20261017\\鹈鹕统一视口截图）。"""
import os, glob, subprocess, pathlib

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
ROOT = r"D:\MG\program\ai交流"
DRAW = os.path.join(ROOT, "AI对比", "ep1 识图")
OUT = os.path.join(ROOT, "_检查补充_20261017", "鹈鹕统一视口截图")
os.makedirs(OUT, exist_ok=True)

MAP = {
    "A_glm": "glm 5.3 flash trae",
    "B_seed": "seed 2.1 turbo trae",
    "C_deepseek": "deepseek V4.1flash 1M workbuddy",
    "D_opus": "opus 5.5 hark",
    "E_hy4": "hy4 preview 1M high workbuddy",
    "F_step5": "Step 5 preview workbuddy",
    "G_gpt": "gpt 6.1 sol high codex",
    "H_auto": "auto fast workbuddy",
    "I_qwen": "qwen 3.8 flash qorde",
    "J_minimax": "MiniMax M3 workbuddy",
}
ok = 0
for code, d in MAP.items():
    hits = glob.glob(os.path.join(DRAW, d, "*.html"))
    if not hits:
        print(code, "MISS"); continue
    png = os.path.join(OUT, f"{code}.png")
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--window-size=1200,680", "--virtual-time-budget=2500",
                    f"--screenshot={png}", pathlib.Path(hits[0]).as_uri()],
                   capture_output=True)
    if os.path.exists(png):
        ok += 1
        print(f"{code:12s} {os.path.getsize(png)} bytes")
print("screenshots:", ok)
