from pathlib import Path
import json,shutil,html,hashlib,zipfile,xml.etree.ElementTree as E,re
B=Path(__file__).resolve().parent
ROOT=Path(r'D:\MG\program\ai交流')
OUT=ROOT/'成品01'
OUT.mkdir(exist_ok=True)
shutil.copy2(B/'delivery'/'AI交流分享-成品01.pptx',OUT/'AI交流分享-成品01.pptx')
support=OUT/'配套';support.mkdir(exist_ok=True)
raw=support/'鹈鹕原始结果';raw.mkdir(exist_ok=True)
for p in (ROOT/'完善版01配套'/'鹈鹕原始结果').glob('*.html'):shutil.copy2(p,raw/p.name)
case=support/'案例材料';case.mkdir(exist_ok=True)
shutil.copy2(ROOT/'AI对比'/'ep1 识图'/'桥墩挡块.pdf',case/'桥墩挡块.pdf')
for p in (ROOT/'AI对比'/'ep2 做表').iterdir():
    if p.suffix in ['.doc','.docx']:shutil.copy2(p,case/p.name)
for p in (ROOT/'AI对比'/'ep2 做表'/'源文件').iterdir():
    if p.is_file():shutil.copy2(p,case/p.name)
shutil.copy2(ROOT/'完善版01配套'/'模型对比口径v3.md',case/'图纸检查点口径-原记录.md')
# Correct the distribution's explanation to the user's requested presentation order.
(case/'本版排序与口径.txt').write_text('本版图纸比较按命中数降序，同分按耗时升序。HY4 本次环境未完成看图，单列在表尾。\n原记录的耗时排序描述是历史版本，不是本版顺序。3 个事后检查点不构成整体准确率。\n报告金额采用汇总表口径，不含暂列金额。99 项金额变化，另 2 项仅数量变化。\n',encoding='utf8')
for name in ['补充_修编变更_两级口径.csv','补充_修编汇总表_公式复算.csv']:
    p=ROOT/'_检查补充_20261017'/name
    if p.exists():shutil.copy2(p,case/name)

manifest=json.loads((B/'manifest.json').read_text(encoding='utf8'))
sources=manifest['sources']
sources['S39']={'name':'GLM-5.3 官方工具与推理规格','url':'https://docs.z.ai/guides/llm/glm-5.3'}
projects={
'文档读取':{'MinerU':'https://github.com/opendatalab/MinerU','PaddleOCR':'https://github.com/PaddlePaddle/PaddleOCR','MarkItDown':'https://github.com/microsoft/markitdown'},
'知识库与流程':{'Dify':'https://github.com/langgenius/dify','RAGFlow':'https://github.com/infiniflow/ragflow','Open WebUI':'https://github.com/open-webui/open-webui','n8n':'https://github.com/n8n-io/n8n'},
'开发与协作':{'obra/superpowers':'https://github.com/obra/superpowers','mattpocock/skills':'https://github.com/mattpocock/skills','Fission-AI/OpenSpec':'https://github.com/Fission-AI/OpenSpec','open-gsd/gsd-core':'https://github.com/open-gsd/gsd-core','BMAD-METHOD':'https://github.com/bmad-code-org/BMAD-METHOD','oh-my-claudecode':'https://github.com/yeachan-heo/oh-my-claudecode','mindfold-ai/Trellis':'https://github.com/mindfold-ai/Trellis'}}
limited={'S03':'官方检索摘录与模型卡，客户端上限另核','S05':'官方型号页检索摘录，接口模态限制另核','S06':'官方帮助中心检索摘录，全文未完整读取','S07':'官方权重模型卡，不等于云端接口的全部限制','S20':'官方帮助中心检索摘录','S32':'原始研究页面本轮未成功读取，未引用 CVSS 数字','S34':'历史原文书目信息与链接，未全文审读','S35':'历史原文书目信息与链接，未全文审读','S36':'裁决原文未成功读取，未列精确赔付金额','S37':'榜单方法页可读，具体 Agent 配置未完整归档'}
css='body{margin:0;background:#f6f7f9;color:#18283b;font:17px/1.7 "Microsoft YaHei",sans-serif}main{max-width:1100px;margin:40px auto;padding:0 24px}h1,h2{color:#0b1f3a}h1{font-size:34px}h2{margin-top:38px;border-bottom:2px solid #1fa7b8;padding-bottom:8px}a{color:#087b83}section{scroll-margin-top:24px}li{margin:8px 0}.muted{color:#5d6c7e}table{border-collapse:collapse;width:100%;font-size:15px}th{background:#0b1f3a;color:white;text-align:left}td,th{padding:10px;border-bottom:1px solid #dce3ea}tr:nth-child(even){background:#ecf0f4}'
parts=[f'<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>AI交流分享：来源与案例</title><style>{css}</style><main><h1>AI 交流分享：来源与案例</h1><p>配套于 52 页成品。公开网页统一查询于 2026-10-18，案例仍为原始单次运行记录，没有重新运行模型。</p><p><a href="鹈鹕盲测.html">打开十份动画盲测</a>　<a href="#case">本地案例材料</a>　<a href="#sources">公开资料</a>　<a href="#limits">核查范围</a></p>']
for pg in manifest['slides']:
    parts.append(f'<section id="{pg["id"]}"><h2>{pg["id"]}　{html.escape(pg["title"])}</h2><ul>')
    for key in pg['sources']:
        if key in sources:
            r=sources[key];note=limited.get(key,'本轮读取原页面或对应官方资料')
            parts.append(f'<li><a href="{html.escape(r["url"],quote=True)}">{key}　{html.escape(r["name"])}</a> <span class="muted">{note}</span></li>')
        else:parts.append(f'<li>{html.escape(key)}</li>')
    if pg['id'] in ['A10','A11']:
        for group,links in projects.items():
            parts.append('<li>'+group+'：'+'　'.join(f'<a href="{u}">{html.escape(n)}</a>' for n,u in links.items())+'</li>')
    parts.append('</ul></section>')
parts.append('<section id="case"><h2>本地案例材料</h2><ul>')
for p in sorted(case.iterdir()):parts.append(f'<li><a href="案例材料/{html.escape(p.name,quote=True)}">{html.escape(p.name)}</a></li>')
parts.append('</ul></section><section id="sources"><h2>全部公开来源</h2><ol>')
for key,r in sources.items():parts.append(f'<li><a href="{html.escape(r["url"],quote=True)}">{key}　{html.escape(r["name"])}</a></li>')
parts.append('</ol></section><section id="limits"><h2>核查范围与待补</h2><ul><li>图纸使用三个事后选定检查点，不代表完整真值集，未发布整体准确率。</li><li>四份报告的同项摘录已逐字对照。参考函能支持部分原因表述，未完成每个变更项的完整因果真值标注。</li><li>AA 数据是统一查询日的具体配置记录，不冒充 2026-10-17 历史快照。TB 的原始 Agent 配置未完整归档，相关裸分数未采用。</li><li>MiMo 云端最大输出、部分接口模态和 Kimi 当期月费未完整确认，未填估计数。</li><li>分享人、实际设备与客户端、订阅和支付演示尚待指定，相关页面保留检查步骤。</li><li>动画页面保留原始输出，静态截图不构成对全部动作的逐帧质量判定。</li></ul></section></main></html>')
(support/'来源索引.html').write_text('\n'.join(parts),encoding='utf8')

names=['GLM 5.3 Flash · Trae','Seed 2.1 Turbo · Trae','DeepSeek V4.1 Flash 1M · WorkBuddy','Opus 5.5 · Hark','HY4 Preview 1M high · WorkBuddy','Step 5 Preview · WorkBuddy','GPT 6.1 Sol high · Codex','Auto fast · WorkBuddy','Qwen 3.8 Flash · Qoder','MiniMax M3 · WorkBuddy']
cards=''.join(f'<article><div class="letter">{chr(65+i)}</div><div class="viewport"><iframe src="鹈鹕原始结果/{chr(65+i)}.html" title="结果 {chr(65+i)}" width="1200" height="900"></iframe></div><p class="config">{html.escape(name)}</p><a href="鹈鹕原始结果/{chr(65+i)}.html" target="_blank">单独播放 {chr(65+i)}</a></article>' for i,name in enumerate(names))
blind=f'''<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>鹈鹕盲测：十份原始动画</title><style>{css}main{{max-width:1600px}}.grid{{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:20px}}article{{position:relative}}.viewport{{position:relative;aspect-ratio:4/3;overflow:hidden;background:white}}iframe{{border:0;position:absolute;top:0;left:0;transform-origin:top left}}.letter{{font-size:25px;font-weight:bold;color:#f28c38}}.config{{visibility:hidden;min-height:55px;font-weight:bold;color:#087b83;font-size:15px;margin:8px 0}}.revealed .config{{visibility:visible}}button{{font:inherit;padding:10px 22px;margin:0 12px 20px 0;background:#0b1f3a;color:white;border:0;cursor:pointer}}@media(max-width:950px){{.grid{{grid-template-columns:repeat(2,minmax(0,1fr))}}}}</style><main><h1>鹈鹕盲测：同一提示词，十份结果</h1><p>一次生成，无历史记忆，不同模型与客户端配置。原始 HTML 保持不变。</p><button id="reveal">揭晓配置</button><a href="来源索引.html#P11">查看来源与条件</a><div class="grid">{cards}</div></main><script>document.querySelector('#reveal').onclick=()=>{{let r=document.body.classList.toggle('revealed');document.querySelector('#reveal').textContent=r?'隐藏配置':'揭晓配置'}};const fit=()=>document.querySelectorAll('.viewport').forEach(v=>v.querySelector('iframe').style.transform='scale('+v.clientWidth/1200+')');new ResizeObserver(fit).observe(document.querySelector('.grid'));fit();</script></html>'''
(support/'鹈鹕盲测.html').write_text(blind,encoding='utf8')
(support/'鹈鹕配置映射.json').write_text(json.dumps([{'id':chr(65+i),'configuration':n,'file':'鹈鹕原始结果/'+chr(65+i)+'.html'} for i,n in enumerate(names)],ensure_ascii=False,indent=2),encoding='utf8')

src=OUT/'制作源文件';src.mkdir(exist_ok=True)
for name in ['build-deck.mjs','visual-pages.mjs','deck-data.mjs','brief.json','enrich.py','finalize.mjs','package.py']:shutil.copy2(B/name,src/name)
assets=src/'assets';assets.mkdir(exist_ok=True)
used=['pelican-crop.png','drawing-crop.png','report-crop.png','reason-reference.png']+[p.name for p in (B/'assets').glob('*1200x900.png')]+[p.name for p in (B/'assets').glob('*-excerpt.png')]
for name in used:shutil.copy2(B/'assets'/name,assets/name)
(OUT/'阅读说明.txt').write_text('AI交流分享 成品01\n\n52页：正文38页，附录14页。详细解释、完整提示词和来源放在每页备注。\nPowerPoint中可使用“视图—备注”或展开页面下方备注区阅读。\n\n第19页配置表按检查点命中数降序，同分按耗时升序，HY4本次环境限制单列。\n第11页放映时先匿名展示，单击同时揭晓10个配置。编辑视图和PDF直接显示映射。\n点鹈鹕缩略图可打开原始HTML，或直接打开 配套/鹈鹕盲测.html 同时播放。\n移动文件时请保留PPT与“配套”目录的同级关系。\n\n分享人、设备/客户端和订阅支付演示仍待指定，相关页面保留步骤。\n参数和证据尚未确认的细项已明确标注，具体范围见配套来源索引。\n原稿与原始实验输出均未覆盖。\n',encoding='utf8')
print('packaged',OUT,flush=True)
print('raw HTML hashes match',all(hashlib.sha256(p.read_bytes()).digest()==hashlib.sha256((ROOT/'完善版01配套'/'鹈鹕原始结果'/p.name).read_bytes()).digest() for p in raw.glob('*.html')),flush=True)
