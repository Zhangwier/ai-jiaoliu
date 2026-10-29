# -*- coding: utf-8 -*-
"""基于成品48制作新的演讲正文（Office原生可编辑/可点击）。
保留154页、现有媒体/证据页与附录，另存49，不覆盖48。
"""
from pathlib import Path
import os,shutil,traceback
import win32com.client as win
ROOT=Path(__file__).resolve().parent.parent
SOURCE=ROOT/"成品"/"AI交流分享-成品48.pptx"
OUTPUT=ROOT/"成品"/"AI交流分享-成品49.pptx"
NOTE=ROOT/"核对"/"成品49_制作说明.md"
C={"navy":"#11253E","navy2":"#1E3850","ink":"#1C354D","muted":"#60758A","bg":"#F6F9FC","white":"#FFFFFF","stroke":"#DAE4EB","teal":"#109A94","orange":"#E68B50","red":"#D75C55","purple":"#746AA9","green":"#2D967A","soft":"#E8F2F5"}
def rgb(t):
 t=t.lstrip("#");return int(t[:2],16)+(int(t[2:4],16)<<8)+(int(t[4:6],16)<<16)
def shape(sl,x,y,w,h,color,rounded=True,line=None):
 sh=sl.Shapes.AddShape(5 if rounded else 1,x,y,w,h)
 sh.Fill.Solid();sh.Fill.ForeColor.RGB=rgb(color)
 if line:
  sh.Line.Visible=-1;sh.Line.ForeColor.RGB=rgb(line);sh.Line.Weight=1.0
 else:sh.Line.Visible=0
 return sh
def label(sl,txt,x,y,w,h,size=20,color=None,bold=False,align=1):
 sh=sl.Shapes.AddTextbox(1,x,y,w,h)
 tr=sh.TextFrame.TextRange;tr.Text=str(txt)
 tr.Font.Name="Microsoft YaHei";tr.Font.Size=size
 tr.Font.Bold=-1 if bold else 0;tr.Font.Color.RGB=rgb(color or C["ink"])
 tr.ParagraphFormat.Alignment=align
 sh.TextFrame.MarginLeft=4;sh.TextFrame.MarginRight=4
 sh.TextFrame.MarginTop=1;sh.TextFrame.MarginBottom=1
 sh.TextFrame.AutoSize=0;sh.TextFrame.VerticalAnchor=3
 return sh
def animate(sl,sh,mode=1):
 try:
  e=sl.TimeLine.MainSequence.AddEffect(sh,10,0,mode)
  e.Timing.Duration=.42
  COUNTS[1]+=1
 except Exception as e:ERRORS.append("动画页%s: %s"%(sl.SlideIndex,str(e)[:120]))
 return sh
def card(sl,x,y,w,h,heading,body,col,idx=None,animated=True):
 sh=shape(sl,x,y,w,h,C["white"],True,C["stroke"])
 shape(sl,x+2,y+13,5,h-26,col,False)
 label(sl,("%s"%(idx or "")),x+16,y+12,43,44,27,col,True)
 sh2=label(sl,heading,x+(62 if idx else 17),y+13,w-(78 if idx else 34),41,20,C["ink"],True)
 sh3=label(sl,body,x+18,y+63,w-35,h-74,15,C["muted"])
 if animated:
  animate(sl,sh,1);animate(sl,sh2,2);animate(sl,sh3,2)
 return sh
def arrow(sl,x1,y1,x2,y2,color=None):
 sh=sl.Shapes.AddLine(x1,y1,x2,y2);sh.Line.ForeColor.RGB=rgb(color or C["teal"]);sh.Line.Weight=2.4
 try:sh.Line.EndArrowheadStyle=3
 except:pass
 return sh
def foot(sl,s,n,dark=False):
 label(sl,s,44,509,795,20,10,"#B9CBDA" if dark else C["muted"])
 label(sl,"V49 / %02d"%n,841,509,74,19,10,"#9BDDD7" if dark else C["teal"],True,3)
def base(sl,n,title,sub,sec,source,dark=False):
 for j in range(sl.Shapes.Count,0,-1):sl.Shapes(j).Delete()
 b=C["navy"] if dark else C["bg"]
 shape(sl,0,0,960,540,b,False);shape(sl,0,0,8,540,C["orange"] if dark else C["teal"],False)
 label(sl,sec,42,17,720,25,11,"#A6E1DA" if dark else C["teal"],True)
 label(sl,title,42,42,878,51,27,C["white"] if dark else C["ink"],True)
 label(sl,sub,42,95,873,25,13,"#BDD2DD" if dark else C["muted"])
 shape(sl,42,125,877,2,"#36536E" if dark else C["stroke"],False)
 foot(sl,source,n,dark)
 COUNTS[0]+=1
def takeaway(sl,msg,y=432,dark=False):
 shape(sl,43,y,876,51,C["navy2"] if dark else C["soft"],True)
 shape(sl,44,y+5,6,41,C["orange"] if dark else C["teal"],False)
 label(sl,msg,62,y+7,845,38,18,C["white"] if dark else C["ink"],True)
def build_cards(sl,items):
 n=len(items)
 if n==2:xs=[53,501];w=406;h=220;y=175
 elif n==3:xs=[47,349,650];w=264;h=222;y=170
 elif n==4:xs=[45,271,497,723];w=192;h=186;y=205
 elif n==5:xs=[40,222,404,586,768];w=164;h=155;y=222
 elif n==7:xs=[40+i*129 for i in range(7)];w=120;h=136;y=219
 else:raise ValueError(n)
 for i,it in enumerate(items):
  c=[C["teal"],C["orange"],C["purple"],C["green"],C["red"]][i%5]
  if len(it)>=3 and it[2] in C:c=C[it[2]]
  hd,body=it[:2]
  z=card(sl,xs[i],y,w,h,hd,body,c,"%02d"%(i+1) if n<6 else None)
  if n in (4,5,7) and i<n-1:arrow(sl,xs[i]+w+3,y+h/2,xs[i+1]-5,y+h/2)
def build_bands(sl,items):
 for i,it in enumerate(items):
  y=157+i*57
  color=[C["teal"],C["orange"],C["purple"],C["green"],C["red"]][i]
  sh=shape(sl,68,y,826,49,C["navy2"],True)
  shape(sl,74,y+8,6,33,color,False)
  tx=label(sl,it,100,y+4,765,40,20,C["white"],True)
  animate(sl,sh);animate(sl,tx,2)

# 固定字段：章节、结论型标题、简短讲述提示、可视化内容、回收一句话、原始口径
SCENES={
 7:("01 认识AI","先分清：训练阶段和使用阶段","两种机制，别混成一种原因",[("训练","参数习得能力和倾向","teal"),("使用","输入、推理、生成、外部工具共同作用","orange")],"一次失败须回到当次输入、实际工具与证据核查。"),
 9:("01 认识AI","聊天模型如何形成？一条简化路线","各产品未必遵循完全相同的阶段",[("预训练","从序列中学习规律","teal"),("示范","学习如何遵循指令","green"),("偏好反馈","改善回答倾向","orange"),("检验","任务与安全评估","purple")],"这是典型路径示意，并非所有模型唯一的固定三步。"),
 12:("01 认识AI","分数涨了，任务却不一定做对","代理评分指标可能偏离真正工作目标",[("真实目标","核准图纸、合同和清单","teal"),("代理分数","数量多、形式完整、用时短","orange"),("可能偏离","为提高分数多报未经证实的问题","red")],"看起来努力，不等于实质上核验正确。"),
 15:("01 认识AI","幻觉：没有证据，也能写得很确定","不要把流畅文字当成已检索过的事实",[("模型文字","“依据第X条，应调整工程量”","orange"),("核验动作","返回真实合同条款或图纸逐项定位","teal")],"原件中找不到对应证据，就不得进入正式交付。"),
 17:("01 认识AI","长上下文：规则还在，可能没有被可靠调用","位置、干扰和任务结构都会影响信息利用",[("规则","单位统一采用 mm","teal"),("材料 A","规格、构件、附件","orange"),("材料 B","多页表格与附录","purple"),("生成","错误地写成 cm","red")],"机制示意，并非某个模型实测注意力权重。"),
 19:("01 认识AI","五种失误，是五组排查线索","训练与使用机制可能相互影响",[("迎合","顺着用户预设","teal"),("奖励偏移","代理目标跑偏","orange"),("幻觉","无来源事实","red"),("波动","输出不稳定","purple"),("长上下文","关键证据未被可靠调用","green")],"不要由单个输出直接倒推唯一训练原因。"),
 23:("02 选工具","模型 ≠ 运行环境：能力与权限需要分开看","原生多模态属于模型；文件与工具调用属于系统配置",[("Model","理解、推理、生成；有些模型可原生看图","teal"),("Harness","连接文件、程序、工具、权限、状态与任务流程","orange")],"能看图的模型，不代表每个客户端都有同等文件权限。"),
 27:("02 选工具","思考档位不保证越高越好","质量、耗时、成本存在任务依赖",[("低档","快，但复杂任务可能不足","teal"),("中档","很多任务的实用权衡","green"),("高档","难题可能获益，也可能收益递减","orange")],"同任务、同检查标准，再比较不同档位。"),
 29:("02 选工具","同品牌，不等于同一套功能","先核实入口、模型ID和权限",[("网页 App","套餐与UI能力","teal"),("API","调用参数、版本、计费","orange"),("Agent工具","终端、文件与自动化","purple"),("企业入口","组织合规与审计","green")],"使用记录必须附具体入口、版本和日期。"),
 30:("02 选工具","别背十二个产品名：先分清三类工作","具体工具与权益到附录再查",[("信息与搜索","回答问题；重点看引文来源","teal"),("办公文件","Office与PDF；重点看实际读写能力","orange"),("代码自动化","终端和仓库；重点看权限和回滚","purple")],"先确认能接入，再确认能验收。"),
 32:("02 选工具","换外壳值得测：但别只看一个分数","特定榜单与任务测试不能外推所有工作",[("厂商测试","本稿引用的8.7分只属于特定基准","teal"),("第三方复测","小样本不一定能区分稳定能力差异","orange")],"以自己的代表任务、成本与错误验收作为最终决策。"),
 33:("02 选工具","选配置，先问三个实际问题","一张排名图回答不了全部",[("读到了吗","真能打开图纸、表格和原始文件吗","teal"),("能交付吗","能输出可编辑、可复核的成果吗","orange"),("值不值","权限、时间和成本能接受吗","purple")],"任何一项不合格，高跑分也帮不了你。"),
 34:("02 选工具","按场景选组合，而不是选“全能王者”","产品建议以具体版本和单位政策为边界",[("Office为主","可把WorkBuddy＋DeepSeek作为候选试用","teal"),("代码为主","考虑Codex等具备终端验证能力的环境","orange"),("涉密资料","优先组织政策，不能授权外发则不上传","red")],"这是当时条件下的试用方向，不是永久排名。"),
 39:("03 真实案例","同一个问题，两份交付：先猜哪份靠谱","耗时和条目数是记录，不是准确率",[("A","29分40秒，提出26条疑点","orange"),("B","3分03秒，提出11条疑点","teal")],"此刻只投票；下一页才揭晓已复核检查点。"),
 40:("03 真实案例","揭晓：条目多，不代表错误找得更准","只按已复核事项表述，不做全面准确率宣传",[("B的检查点","K1/K2两个指定检查点均命中","teal"),("A的误报","所列26条中有8条已核为假错误","red"),("待核","另有21条待核，不能擅算对错","orange")],"结论仅针对本次配置、图纸和已复核范围。"),
 41:("03 真实案例","“我看过图”是自述，不是完整证据","读图方式与判断正确性必须分开验证",[("模型自述","回答称看了原图","orange"),("工具轨迹","记录实际打开和引用了什么","purple"),("原件核验","坐标、尺寸、符号、公式逐项比对","teal")],"未说读图过程，不代表一定没读；说看了也不代表看准。"),
 49:("03 真实案例","案例三：没有新材料，只有一次强烈反对","讲述重点是判断依据有没有变化",[("第一轮","模型作出了有条件判断","teal"),("后续输入","没有新增图纸、坐标和计算","orange"),("问题","它会先核验证据，还是直接改口？","purple")],"原对话保留于本地案例素材；这里不重写真实人物原话。"),
 50:("03 真实案例","没有新证据，立场却发生变化","一次具体观察不能直接证明某种训练因果",[("R1初判","给出一个带前提的结论","teal"),("用户反驳","新增可核证材料：0","orange"),("R3改口","判断方向发生反转","red")],"更正结论应当基于证据和复算，而非态度压力。"),
 54:("04 更可靠地用AI","v0→v6：不是把风险消灭，而是增加检查关口","同一个审图任务分七版逐步完善",[("v0","只下口令","orange"),("v1","范围目标","teal"),("v2","分步产物","teal"),("v3","上下文","teal"),("v4","检索证据","teal"),("v5","流程验收","teal"),("v6","独立复核","teal")],"新增措施是设计建议，是否有效还需要实际复测。"),
 55:("04 更可靠地用AI","别只说“帮我审完”：写一份任务合同","四项内容，把无形的要求变成可核标准",[("范围","材料、版本、要审和不审","teal"),("来源","页码、条款、表格、原图","green"),("边界","没把握标待核，不造依据","orange"),("验收","单位、公式、证据、文件","purple")],"清晰任务不等于零错误，最终仍须对成果验收。"),
 56:("04 更可靠地用AI","把“大任务”拆成四个可交接的工位","每一步留下别人能检查的产物",[("材料","版本和来源清单","teal"),("定位","原图的证据位置","green"),("复算","单位和计算底稿","orange"),("交付","结构化审查工作簿","purple")],"没有验收物的分工，只是在转移不确定性。"),
 57:("04 更可靠地用AI","上下文管理：保留关键规则，而不是保留每一句话","旧长聊与压缩交接，能否让证据不丢？",[("旧会话","版本混杂，约束埋在讨论中","orange"),("交接单","目标、规则、来源、未决问题和下一步","teal")],"换会话前必须保留有用证据和状态。"),
 58:("04 更可靠地用AI","RAG：先检索可能有用的证据，再回答","检索、重排和核验的组合依实现而不同",[("提问","明确待核事实","teal"),("召回","找候选材料","purple"),("选证据","对口径和页码","orange"),("生成","引用支持的回答","green")],"找不到可靠来源时，应说明缺失或拒绝猜测。"),
 59:("04 更可靠地用AI","RAG和Agent读文件，可以合作","两个路径最终都要回到同一份原始证据",[("RAG路径","索引/召回/选段落：聚焦，但可能漏召回","teal"),("Agent路径","用工具搜索/打开/比对文件：灵活，也有成本","purple")],"不是所有Agent都只能按文件名搜，也不是所有RAG都能精准查全。"),
 60:("04 更可靠地用AI","一张图纸的复核工作流：五步都有产物","从“找到问题”到“能拿给别人复核”",[("输入","版本确认","teal"),("识别","几何定位","green"),("计算","列公式","orange"),("核对","回到原图","purple"),("提交","证据工作簿","teal")],"出了争议退回原图和底稿，而不是继续圆话。"),
 61:("04 更可靠地用AI","被质疑时，先处理证据，再决定改口","给模型明确的更正规则",[("标记异议","用户究竟不同意什么？","orange"),("核对材料","有没有新原图、计算或条文？","purple"),("判断更正","有证据则改；无证据说明坚持理由","teal")],"本规则是拟议措施，本次未通过控制实验验证效果。"),
 62:("04 更可靠地用AI","工作流颗粒度：固定验收，开放方法","不是越细越好，也不是越自由越好",[("太松","口径和交付不可控","orange"),("太死","路径无法适应新情况","red"),("恰当","关键关口固定，探索路径灵活","teal")],"留给模型判断空间，也要保留审计追踪。"),
 63:("04 更可靠地用AI","多Agent的前提是：分得开、验得了、有人兜底","并行会增加成本，也可能传播同一个错误",[("资料定位","找到原图和出处","teal"),("独立复核","重新计算并反驳","purple"),("合并裁决","对冲突由人签字确认","orange")],"多Agent不是总优于单Agent，先看独立检验有没有意义。"),
 64:("04 更可靠地用AI","数据分三级处理；高风险动作要人批准","模型会操作不等于操作有授权",[("公开资料","仍需版权和版本核对","teal"),("内部资料","服从企业外发与存储政策","orange"),("敏感资料","默认不外发；删除付款分享人工确认","red")],"权限控制应当落在系统与工作流里，而不只是提示词。"),
 65:("04 更可靠地用AI","五类风险有了对应检查；仍由人最终验收","“有检查”不是“错误已经消失”",[("迎合","先查新增证据","teal"),("奖励偏移","改用实际交付验收","orange"),("幻觉","定位原始引用","purple"),("波动","关键输出交叉检查","green"),("长上下文","提取和保留重要规则","red")],"原始图纸、真实金额和合同条文，最终都由人核准。"),
 66:("结语","记住三个动作：选、交、核","把这张图拍下来，就能作为工作时的检查卡",[("选配置","确认模型、工具、权限与成本","teal"),("交任务","写清范围、来源、限制、验收","orange"),("核交付","打开原件、复算结果、标待核","purple")],"能生成不等于能交付，能交付不等于可以直接信任。")
}
# 保留含原生视频、GIF或图片的页面；只对没有媒体关系的纯文本解释页做整页重绘。
# 根据原始PPT包自动判断，避免开发时误删除原本已经优秀的动态素材。
import zipfile
with zipfile.ZipFile(SOURCE) as _zip:
 _names=set(_zip.namelist())
 _protected=set()
 for _i in SCENES:
  _rel="ppt/slides/_rels/slide%d.xml.rels"%_i
  if _rel in _names and b"/media/" in _zip.read(_rel):
   _protected.add(_i)
SCENES={_i:_content for _i,_content in SCENES.items() if _i not in _protected}
print("PRESERVE_MEDIA_SLIDES",sorted(_protected),flush=True)
COUNTS=[0,0];ERRORS=[]
def rewrite_text(sl,old,new):
 for sh in sl.Shapes:
  try:
   t=sh.TextFrame.TextRange.Text
   if old in t and len(t)<220:
    sh.TextFrame.TextRange.Text=t.replace(old,new)
    return True
  except:pass
 return False
def main():
 if not SOURCE.exists():raise RuntimeError("原稿不存在 "+str(SOURCE))
 if OUTPUT.exists():raise RuntimeError("新版文件已存在，拒绝直接覆盖："+str(OUTPUT))
 shutil.copy2(SOURCE,OUTPUT)
 app=win.DispatchEx("PowerPoint.Application")
 app.DisplayAlerts=1
 ppt=None
 try:
  ppt=app.Presentations.Open(str(OUTPUT),False,False,False)
  if ppt.Slides.Count!=154:raise RuntimeError("页数不符")
  for n,(sec,title,sub,items,summary) in SCENES.items():
   try:
    sl=ppt.Slides(n)
    base(sl,n,title,sub,sec,"示意/讲述版；详见来源与原始成品48的相应页",n in (39,49,65))
    if n==65:
     build_bands(sl,[a+"  →  "+b for a,b,*_ in items])
     takeaway(sl,summary,462,True)
    else:
     build_cards(sl,items)
     takeaway(sl,summary,437,n in (39,49))
    print("REDESIGNED",n,"cards",len(items),flush=True)
   except Exception as ex:
    ERRORS.append("第%s页重绘错误：%s"%(n,str(ex)[:160]))
    print("FAIL",n,str(ex)[:150],flush=True)
  fixes={
   13:[("逐 token 生成：回答是一步一步猜出来的","逐 token 生成：回答是一步一步构建出来的")],
   16:[("注意力按相关性组合上下文信息","注意力基于输入关系组合上下文信息")],
   20:[("其实只分属五种翻车","可由五类风险线索理解")],
   24:[("分数和价格不成正比","分数和价格衡量不同问题")],
   25:[("推理数学看 Astra，综合与干活看 Claude","公开榜单：不同任务，排名口径不同")],
   26:[("美国两强领先，国产头部约低 10 分","榜单格局依测试任务与模型版本而变")],
   51:[("为什么会改口：被反驳的对话里，谄媚率翻倍","研究提示：被反驳情境下的迎合现象"),
       ("它不是被说服了，是被压住了","在没有新证据时立场发生变化，原因未被单独验证")],
   52:[("五个点全亮了：三个案例，踩遍了训练和生成两边","三个真实案例，展示五类不同风险线索")],
   53:[("每升级一次，风险地图上就熄灭一个点","每升级一次，增加一个检查关口")]
  }
  for n,v in fixes.items():
   for old,new in v:
    if rewrite_text(ppt.Slides(n),old,new):print("TEXT",n,old[:13],flush=True)
  ppt.Save()
  print("SAVED",OUTPUT,"bytes",OUTPUT.stat().st_size,"pages",ppt.Slides.Count,"redraws",COUNTS[0],"effects",COUNTS[1],flush=True)
 finally:
  if ppt:
   try:ppt.Close()
   except:pass
  try:app.Quit()
  except:pass
def report():
 lines=[
 "# 成品49｜制作与回归记录","",
 "源稿：成品48（未覆盖）。输出：成品49，154页顺序不变。",
 "采用PowerPoint原生形状、文本和淡入时序；不是HTML截图伪装PPT。",
 "重绘核心正文页 %s 页：%s。"%(COUNTS[0],", ".join(map(str,SCENES.keys()))),
 "新增原生动画效果 %s 个。原13页的逐token动画、原22页鹈鹕多图与第2页媒体均未重制。"%COUNTS[1],
 "重点科学性文字修订：模型/工具、多模态、榜单、聊天改口、RAG、风险并非归零。",
 "附录68-144及隐藏页145-154未做结构性变更；原稿48仍可查看全文对照。",
 "边界：本版解决重点页面文字密度与可视讲述，不表示所有排名、价格和原始报告已经重新核验。",
 "边界：静态渲染不能证明点击动画或视频在所有Office版本均正常；须进行放映态验证。",
 "第40页原始十组数据未迁成新版列表，原始详细表见原成品48第40页；新版只呈现已核验的K1/K2、误报及待核分组。",
 "",
 "## 构建错误"
 ]+["- "+x for x in ERRORS] if ERRORS else [
 "# 成品49｜制作与回归记录","",
 "源稿成品48未覆盖；成品49保留154页。",
 "重绘正文页 %s 页：%s。"%(COUNTS[0],", ".join(map(str,SCENES.keys()))),
 "新增原生动画 %s 个。"%COUNTS[1],
 "2/13/22/35等已有媒体/动画保留；附录68-144及隐藏页145-154未结构性删改。",
 "第40页详细原表请参照成品48第40页。",
 "动画与视频仍需PowerPoint放映态确认；科学事实与企业资料需按原件复核。",
 "构建错误：无。"
 ]
 NOTE.write_text("\n".join(lines)+"\n",encoding="utf-8")
 print("REPORT",NOTE,flush=True)
if __name__=="__main__":
 try:main()
 finally:report()
