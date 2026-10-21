# -*- coding: utf-8 -*-
"""从源稿抽取先导片与封面帧"""
import os, zipfile, shutil
ROOT = r"D:\MG\program\ai交流\_v36"
SRC = r"D:\MG\program\ai交流\成品\AI交流分享-成品34.pptx"
z = zipfile.ZipFile(SRC)
data = z.read("ppt/media/media14.mp4")
p = os.path.join(ROOT, "teaser.mp4")
open(p, "wb").write(data)
print("video", len(data) / 1048576, "MB")
