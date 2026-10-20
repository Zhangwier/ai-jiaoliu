import json
d=json.load(open(r"work\cells_raw.json",encoding="utf-8"))
rows={}
for c in d["cells"]:
    rows.setdefault((c["table"],c["row"]),{})[c["col"]]=(c["text"],c["score"])
for (t,r),cols in sorted(rows.items()):
    if t=="L" and (r>560 or r<76): continue
    items=" | ".join(f"{k}:{v[0]}" for k,v in sorted(cols.items()))
    print(f"{t} {r:6.1f}  {items}")
print("SUMS:", d["sum_rows"])
