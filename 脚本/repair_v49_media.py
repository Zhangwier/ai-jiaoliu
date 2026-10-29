# -*- coding: utf-8 -*-
"""修复PowerPoint COM另存时误清理的原稿媒体引用。
只恢复原稿中未重绘的第12、142、149、152页，其它重绘页完全不动。
所有媒体复制到新唯一名称以避免Office重排后路径名冲突。
"""
from pathlib import Path
import zipfile,re,hashlib,os,shutil
BASE=Path(__file__).resolve().parent.parent/'成品'
OLD=BASE/'AI交流分享-成品48.pptx'
NEW=BASE/'AI交流分享-成品49.pptx'
TMP=BASE/'AI交流分享-成品49-媒体校正临时文件.pptx'
PARTS=(12,142,149,152)
assert OLD.is_file() and NEW.is_file() and not TMP.exists()
zold=zipfile.ZipFile(OLD)
znew=zipfile.ZipFile(NEW)
updates={}
add={}
for n in PARTS:
 slide='ppt/slides/slide%d.xml'%n
 rel='ppt/slides/_rels/slide%d.xml.rels'%n
 oldrel=zold.read(rel)
 parsed=oldrel.decode('utf-8')
 # 用原稿的slideXML、relationship；恢复媒体时使用唯一命名
 updates[slide]=zold.read(slide)
 for target in sorted(set(re.findall(r'Target="(\.\./media/[^"]+)"',parsed))):
  basename=target.split('/')[-1]
  origpath='ppt/media/'+basename
  assert origpath in zold.namelist(),origpath
  dest='v48safe_s%d_%s'%(n,basename)
  destpath='ppt/media/'+dest
  assert destpath not in znew.namelist() and destpath not in add
  add[destpath]=zold.read(origpath)
  parsed=parsed.replace(target,'../media/'+dest)
 updates[rel]=parsed.encode('utf-8')
 # 确保旧文件中的其它外部组件，如布局、批注仍可被新包访问
 for target in re.findall(r'Target="([^"]+)"',parsed):
  if target.startswith('../media/') or target.startswith('http'):continue
  if target.startswith('../'):
   pt='ppt/'+target.replace('../','')
   if pt not in znew.namelist():print('WARN_UNRESOLVED',n,pt)
with zipfile.ZipFile(TMP,'w') as out:
 for inf in znew.infolist():
  if inf.filename in updates:out.writestr(inf,updates[inf.filename])
  else:out.writestr(inf,znew.read(inf.filename))
 for k,data in add.items():out.writestr(k,data,compress_type=zipfile.ZIP_DEFLATED)
zold.close();znew.close()
with zipfile.ZipFile(TMP) as z:
 assert z.testzip() is None
 for name,value in add.items():assert z.read(name)==value
 from pptx import Presentation
 assert len(Presentation(str(TMP)).slides)==154
 print('RESTORED_REFERENCES',list(PARTS),'restored_media_files',len(add),'temp_size',TMP.stat().st_size)
 # 保护用户原版48与这次重制的其它页面
 for i in (7,17,23,39,40,50,54,58):
  assert z.read('ppt/slides/slide%d.xml'%i) == zipfile.ZipFile(NEW).read('ppt/slides/slide%d.xml'%i)
# 验证通过后原子替换我们的49结果；不涉及用户原始48
os.replace(TMP,NEW)
print('FINAL',NEW,NEW.stat().st_size)
