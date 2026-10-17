from pathlib import Path
import json,win32com.client,zipfile,xml.etree.ElementTree as E,posixpath
B=Path(__file__).resolve().parent
m=json.loads((B/'manifest.json').read_text(encoding='utf8'))
try:
    prior=win32com.client.GetActiveObject('PowerPoint.Application')
    for i in range(prior.Presentations.Count,0,-1):
        old=prior.Presentations(i)
        if str(old.FullName).lower()==str(B/'candidate.pptx').lower():
            old.Saved=True;old.Close()
except Exception:pass
app=win32com.client.DispatchEx('PowerPoint.Application')
deck=app.Presentations.Open(str(B/'candidate.pptx'),False,False,False)
for link in m['links']:
    s=deck.Slides(link['slide'])
    try:sh=s.Shapes(link['shape'])
    except Exception:
        sh=next((v for v in s.Shapes if v.AlternativeText==link['shape']),None)
        if sh is None and link['shape'].startswith('pelican-'):
            pictures=[v for v in s.Shapes if v.Type==13]
            if len(pictures)==10:sh=pictures[int(link['shape'].split('-')[-1])]
        if sh is None:raise
        sh.Name=link['shape']
    act=sh.ActionSettings(1)
    act.Action=7
    if 'target' in link:
        target=deck.Slides(link['target'])
        act.Hyperlink.Address=''
        act.Hyperlink.SubAddress=f'{target.SlideID},{link["target"]},{m["slides"][link["target"]-1]["title"]}'
    else:
        act.Hyperlink.Address=link['address'].replace('/','\\')
        act.Hyperlink.SubAddress=''
s=deck.Slides(11)
for i in range(10):
    effect=s.TimeLine.MainSequence.AddEffect(s.Shapes(f'reveal-{i}'),1,0,1 if i==0 else 2)
    effect.Timing.Duration=0.2
for i in range(1,53):
    sl=deck.Slides(i)
    sl.SlideShowTransition.AdvanceOnClick=True
    sl.SlideShowTransition.AdvanceOnTime=False
    for sh in sl.NotesPage.Shapes:
        if sh.HasTextFrame and sh.TextFrame.HasText:
            sh.TextFrame.TextRange.Font.Name='Microsoft YaHei'
            sh.TextFrame.TextRange.Font.Size=12
deck.SaveAs(str(B/'enriched-com.pptx'),24)
print('enriched',deck.Slides.Count,'slides','reveal effects',s.TimeLine.MainSequence.Count,flush=True)
deck.Close()
if app.Presentations.Count==0:app.Quit()

# Store portable relative external links, independently of PowerPoint SaveAs.
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
R='{http://schemas.openxmlformats.org/package/2006/relationships}'
E.register_namespace('p',NS['p']);E.register_namespace('a',NS['a']);E.register_namespace('r',NS['r'])
with zipfile.ZipFile(B/'enriched-com.pptx') as z:
    data={n:z.read(n) for n in z.namelist()}
pr=E.fromstring(data['ppt/presentation.xml'])
rels={r.attrib['Id']:r.attrib['Target'] for r in E.fromstring(data['ppt/_rels/presentation.xml.rels'])}
slideparts=[posixpath.normpath('ppt/'+rels[x.attrib['{'+NS['r']+'}id']]) for x in pr.find('p:sldIdLst',NS)]
patched=0
for i,part in enumerate(slideparts,1):
    root=E.fromstring(data[part]);rp=posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'
    if rp not in data:continue
    relroot=E.fromstring(data[rp]);byid={x.attrib['Id']:x for x in relroot}
    for link in [a for a in m['links'] if a['slide']==i and 'address' in a]:
        nv=next((x for x in root.findall('.//p:cNvPr',NS) if x.attrib.get('name')==link['shape']),None)
        if nv is None:raise RuntimeError(link)
        hc=nv.find('a:hlinkClick',NS)
        if hc is None:raise RuntimeError('Missing link '+str(link))
        rid=hc.attrib['{'+NS['r']+'}id']
        byid[rid].set('Target',link['address'])
        byid[rid].set('TargetMode','External');patched+=1
    data[rp]=E.tostring(relroot,encoding='utf-8',xml_declaration=True)
# PowerPoint serializes binary floats to 17-digit strings. Restore the source
# numeric precision before generating portable embedded chart workbooks.
from decimal import Decimal
cn={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
for name in data:
    if name.startswith('ppt/charts/chart') and name.endswith('.xml'):
        cr=E.fromstring(data[name])
        for v in cr.findall('.//c:numCache/c:pt/c:v',cn)+cr.findall('.//c:numLit/c:pt/c:v',cn):
            v.text=format(Decimal(v.text).quantize(Decimal('0.000001')).normalize(),'f')
        data[name]=E.tostring(cr,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(B/'enriched.pptx','w',zipfile.ZIP_DEFLATED) as z:
    for n,b in data.items():z.writestr(n,b)
print('portable links',patched,flush=True)
