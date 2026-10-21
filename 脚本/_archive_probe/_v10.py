import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
p = d[23]
for b in p.get_text("dict")["blocks"]:
    for l in b.get("lines", []):
        txt = "".join(s["text"] for s in l["spans"])
        if any(k in txt for k in ("梯队","紧随","其后","Anthropic","Google","xAI","国产追赶")):
            x0,y0,x1,y1 = l["bbox"]
            print(f"y {y0/28.3465:5.2f}-{y1/28.3465:5.2f}cm  x {x0/28.3465:5.2f}-{x1/28.3465:5.2f}cm  {txt!r}")
