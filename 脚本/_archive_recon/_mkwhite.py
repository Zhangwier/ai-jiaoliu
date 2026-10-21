from PIL import Image
import os, glob
src = r"D:\MG\program\ai交流\成品\_work29\logos4"
out = r"D:\MG\program\ai交流\成品\_work29\white"
os.makedirs(out, exist_ok=True)
for n in ["openai","anthropic","claude","gemini","google","deepseek","doubao","qwen","zhipu","moonshot","trae","qoder","cline","codex","kiro","meta","grok","mistral","hy","wenxin","minimax"]:
    f = os.path.join(src, n+".png")
    if not os.path.exists(f): print("miss", n); continue
    im = Image.open(f).convert("RGBA")
    px = im.load()
    w,h = im.size
    for y in range(h):
        for x in range(w):
            r,g,b,a = px[x,y]
            px[x,y] = (255,255,255,a)
    im.save(os.path.join(out, n+".png"))
    print("white", n)
