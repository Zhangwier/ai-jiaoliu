# -*- coding: utf-8 -*-
"""重整 核对/鹈鹕配置映射.json：把失效的绝对路径改成仓库内相对路径。"""
import json
import io

P = r"D:\MG\program\ai交流\核对\鹈鹕配置映射.json"
data = json.load(io.open(P, encoding="utf-8"))

for item in data:
    src = item.get("source", "")
    # 取倒数两段：配置目录 / 文件名
    parts = src.replace("/", "\\").split("\\")
    cfg, fname = parts[-2], parts[-1]
    item["source"] = "AI对比/ep1 识图/%s/%s" % (cfg, fname)

io.open(P, "w", encoding="utf-8").write(
    json.dumps(data, ensure_ascii=False, indent=2) + "\n")

for item in data:
    print(item["code"], item["source"])
