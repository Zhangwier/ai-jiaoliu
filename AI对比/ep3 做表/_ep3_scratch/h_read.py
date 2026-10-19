import json, os
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
log = json.load(open(OUT + r'\e2_log.json', encoding='utf-8'))
def walk(o, p=''):
    if isinstance(o, dict):
        for k, v in o.items():
            yield from walk(v, p + '/' + str(k))
    elif isinstance(o, list):
        yield (p + f'[len={len(o)}]', o[:2] if o else [])
    else:
        yield (p, o)
print('=== e2_log 顶层键 ===', list(log.keys()))
for k, v in log.items():
    if isinstance(v, (str, int, float, bool)) or v is None:
        print(' ', k, '=', v)
    elif isinstance(v, list):
        print(' ', k, f'[list len={len(v)}]', json.dumps(v[:2], ensure_ascii=False)[:300])
    elif isinstance(v, dict):
        print(' ', k, f'[dict keys={list(v)[:12]}]')
        for kk, vv in list(v.items())[:30]:
            print('     ', kk, json.dumps(vv, ensure_ascii=False)[:200])
print('\n=== g2_summary ===')
g = json.load(open(OUT + r'\g2_summary.json', encoding='utf-8'))
print(json.dumps(g, ensure_ascii=False, indent=1)[:2500])
