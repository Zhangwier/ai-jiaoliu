import json
d=json.load(open(r"work\cells_raw.json",encoding="utf-8"))
print("left_rows",[round(x,1) for x in d["left_rows"]])
print("right_rows",[round(x,1) for x in d["right_rows"]])
rows={}
for c in d["cells"]:
    rows.setdefault((c["table"],c["row"]),{})[c["col"]]=(c["text"],c["score"])
for (t,r),cols in sorted(rows.items()):
    items=" | ".join(f"{k}:{v[0]}({v[1]:.2f})" for k,v in sorted(cols.items()))
    print(f"{t} {r:6.1f}  {items}")
