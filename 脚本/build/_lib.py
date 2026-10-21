# -*- coding: utf-8 -*-
"""转发到 脚本/_lib.py，让子目录里的脚本可以直接 `python batch/xxx.py` 运行。"""
import os, sys
_p = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "_lib.py")
if not os.path.exists(_p):
    raise SystemExit("找不到 _lib.py: %s" % _p)
exec(compile(open(_p, encoding="utf-8").read(), _p, "exec"))
