import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Presentation,PresentationFile} from 'file:///C:/Users/MG/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
import {sources,aa,tablePages,pageSources} from './deck-data.mjs';
import {renderVisual} from './visual-pages.mjs';
const B=path.dirname(fileURLToPath(import.meta.url));
const A=path.join(B,'assets');
const brief=JSON.parse(await fs.readFile(path.join(B,'brief.json'),'utf8'));
const P=Presentation.create({slideSize:{width:1280,height:720}});
const C={navy:'#0B1F3A',cyan:'#22A6B5',orange:'#F28C38',ink:'#18283B',muted:'#5D6C7E',paper:'#F6F7F9',white:'#FFFFFF',line:'#DCE3EA',teal:'#087B83'};
const FONT='Microsoft YaHei';
process.env.RUNTIME_NODE_MODULES='C:/Users/MG/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const {applyPresentationChartFont}=await import('file:///C:/Users/MG/.codex/plugins/cache/openai-primary-runtime/presentations/26.1015.12040/skills/presentations/container_tools/artifact_tool_utils.mjs');
let serial=0;const links=[];const chartOwners=[];const tableOwners=[];const manifest=[];
function clean(s){return String(s??'').replace(/\*\*/g,'').replace(/`/g,'').replace(/\[([^\]]+)\]\([^)]+\)/g,'$1').replace(/；/g,'，').trim();}
function wrap(s,width,size){
 return clean(s).split('\n').map(line=>{let out='',curr='',n=0;for(const ch of line){let k=/[\x00-\x7F]/.test(ch)?(ch===' '?0.3:0.59):1.02;if(n+k>width/size&&curr){out+=curr+'\n';curr='';n=0;}curr+=ch;n+=k;}return out+curr;}).join('\n');
}
function txt(sl,name,value,x,y,w,h,size=23,color=C.ink,bold=false,align='left',doWrap=true){
 const sh=sl.shapes.add({geometry:'textbox',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 sh.text=doWrap?wrap(value,w,size):clean(value);
 sh.text.style={typeface:FONT,fontSize:size,color,bold,alignment:align,verticalAlignment:'middle',autoFit:'none',wrap:'none',insets:0};return sh;
}
function node(sl,name,value,x,y,w,h,fill=C.white,size=23,color=C.ink){
 const sh=sl.shapes.add({geometry:'rect',name,position:{left:x,top:y,width:w,height:h},fill,line:{fill:C.line,width:1}});
 sh.text=wrap(value,w-24,size);sh.text.style={typeface:FONT,fontSize:size,color,bold:true,alignment:'center',verticalAlignment:'middle',autoFit:'none',insets:10};return sh;
}
function line(sl,name,x,y,w,h,color=C.cyan,width=2){return sl.shapes.add({geometry:'line',name,position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:color,width,style:'solid'}});}
async function pic(sl,file,name,x,y,w,h,fit='cover'){
 const b=await fs.readFile(path.join(A,file));const sh=sl.images.add({name,blob:b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength),contentType:'image/png',alt:name,fit,position:{left:x,top:y,width:w,height:h}});return sh;
}
function section(id){if(id.startsWith('A'))return '附录';let n=+id.slice(1);return n<=10?'认识 AI':n<=18?'模型与工具':n<=30?'工作实践':'上手与边界';}
function header(page,intro=''){
 const sl=P.slides.add();serial++;sl.background.fill=C.paper;
 txt(sl,'section',section(page.id),64,30,500,28,18,C.muted);
 txt(sl,'page',page.id.startsWith('A')?page.id:String(serial).padStart(2,'0'),1120,30,96,28,18,C.muted,false,'right');
 const size=page.title.length>27?36:page.title.length>22?39:42;
 txt(sl,'title',page.title,64,77,1152,63,size,C.navy,true,'left',false);
 if(intro)txt(sl,'intro',intro,66,149,1150,44,22,C.muted);
 const refs=pageSources[page.id]??['本地分享材料'];
 const friendly={P04:'学习关系示意',P05:'非线性关系示意',P11:'10 份原始 HTML 与统一视口截图',P14:'Artificial Analysis 具体型号页',P19:'本次模型输出与人工核对记录',P20:'本次模型输出、原图与人工复核',P22:'四份修编报告、源清单与复算底稿',P23:'四份修编报告与设计后服务函',P37:'网信办备案公告、Anthropic 使用政策',A02:'Artificial Analysis 具体型号页',A03:'各厂商官方型号文档',A04:'各产品官方订阅与价格页'};
 let labels=friendly[page.id]??refs.map(k=>sources[k]?k:k).join('；');if(labels.length>62)labels=labels.slice(0,58)+'…';
 txt(sl,'source','来源：'+labels+'　｜　2026-10-18',66,684,1065,24,13,C.muted,false,'left',false);
 txt(sl,'source-link','来源索引',1124,681,92,28,15,C.teal,false,'right');
 links.push({slide:serial,shape:'source-link',address:'配套/来源索引.html#'+page.id});
 sl.speakerNotes.text=refs.map(k=>sources[k]?`${k} ${sources[k].name}\n${sources[k].url}`:k).join('\n\n')+'\n\n查询与整理日期：2026-10-18。';
 manifest.push({id:page.id,title:page.title,slide:serial,sources:refs});return sl;
}
function takeaway(sl,main,sub='',y=605){
 if(main)txt(sl,'takeaway',main,66,y,1150,sub?40:60,25,C.navy,true);
 if(sub)txt(sl,'takeaway-detail',sub,66,y+43,1150,34,18,C.muted);
}
function mdtable(md){const lines=md.split('\n').filter(x=>x.trim().startsWith('|'));return lines.filter(x=>!/^\|[\s:|-]+\|?$/.test(x.trim())).map(x=>x.trim().replace(/^\||\|$/g,'').split('|').map(clean));}
function table(sl,values,widths,opts={}){
 const font=opts.font??22,headfont=opts.headfont??20;const x=opts.x??64,y=opts.y??215;
 let prepared=values.map((row,r)=>row.map((v,c)=>wrap(v,widths[c]-28,r===0?headfont:font)));
 let heights=prepared.map((row,r)=>Math.max(opts.minRow??37,Math.max(...row.map(t=>t.split('\n').length))*(r===0?headfont:font)*1.28+12));
 if(opts.heights)heights=opts.heights;
 const h=heights.reduce((a,b)=>a+b,0);
 if(y+h>(opts.maxBottom??593)) console.log('TABLE_HEIGHT',serial,y+h);
 const t=sl.tables.add({rows:values.length,columns:widths.length,left:x,top:y,width:widths.reduce((a,b)=>a+b,0),height:h,columnWidths:widths,values:prepared});
 t.styleOptions={headerRow:false,bandedRows:false};t.borders.assign({fill:'none',width:0});
 t.cells.block({row:0,column:0,rowCount:values.length,columnCount:widths.length}).assign({margins:{left:14,right:10,top:0,bottom:0}});
 for(let r=0;r<values.length;r++){t.rows[r].height=heights[r];for(let c=0;c<widths.length;c++){
   const cell=t.getCell(r,c);cell.fill=r===0?C.navy:(opts.highlight?.includes(r)?'#E4F2F2':(r%2===0?'#ECF0F4':C.paper));
   cell.text.style={typeface:FONT,fontSize:r===0?headfont:font,color:r===0?C.white:C.ink,bold:r===0||c===0,alignment:opts.center?.includes(c)?'center':'left',verticalAlignment:'middle',autoFit:'none',wrap:'none',insets:0};
 }}
 if(!tableOwners.includes(serial))tableOwners.push(serial);return t;
}
function listRows(sl,rows,{x=64,y=218,w=1152,step=100,labelW=240,font=24}={}){
 for(let i=0;i<rows.length;i++){const yy=y+i*step;txt(sl,`row-num-${i}`,String(i+1).padStart(2,'0'),x,yy,52,40,26,C.orange,true);txt(sl,`row-label-${i}`,rows[i][0],x+72,yy,labelW-72,44,font,C.navy,true);txt(sl,`row-body-${i}`,rows[i][1],x+labelW+26,yy,w-labelW-26,step-14,font,C.ink);if(i<rows.length-1)line(sl,`row-rule-${i}`,x,yy+step-15,w,0,C.line,1);}
}
function flow(sl,steps,{x=64,y=255,w=1152,h=115,gap=42,size=25}={}){
 const ww=(w-gap*(steps.length-1))/steps.length;
 steps.forEach((s,i)=>{node(sl,'flow-'+i,s,x+i*(ww+gap),y,ww,h,i%2?'#E4F2F2':C.white,size);if(i<steps.length-1)txt(sl,'arrow-'+i,'→',x+(i+1)*ww+i*gap+5,y+h/2-22,gap-10,44,28,C.cyan,true,'center',false);});
}
function chart(sl,type,config){const ch=sl.charts.add(type,config);applyPresentationChartFont(ch,{fontFamily:FONT});for(let i=0;i<(config.series?.length??0);i++){const st=ch.series.getItemAt(i).dataLabels?.textStyle;if(st)st.typeface=FONT;}chartOwners.push(serial);return ch;}
const ax={textStyle:{fontSize:18,fill:C.muted},line:{fill:C.line,width:1},majorGridlines:{fill:C.line,width:1}};
const override={};

override.P16=[['付费对象','通常买到什么','付款前核对什么'],['聊天产品会员','指定模型、功能与用量','地区、额度、自动续费和取消方式'],['API','按实际调用量计费的开发接口','输入、输出、缓存、工具调用的分项价格'],['工具或编程产品','客户端功能及套餐内用量','模型路由、自带 API、超额和赠额规则']];
override.P17=[['类型','主要用途','示例','选择时看什么'],['聊天入口','问答、写作、资料分析','DeepSeek、千问、Kimi\nChatGPT、Claude、Gemini','文件能力与地区可用性'],['任务客户端','本地文件、仓库与调试','Codex、Claude Code、Cursor\nQoder、Trae、WorkBuddy\nCline、OpenCode','环境权限、模型路由与执行反馈'],['工作流','串联重复步骤','n8n、Dify','失败处理、日志与人工审批'],['文档处理','将资料转为可处理内容','PaddleOCR、MinerU\nMarkItDown','扫描件、表格结构和完整性']];
override.P20=[['本次观察','对应结果','能得出的结论'],['找到较多检查点','Opus 3/3\nGPT、Qwen 2/3','本次检查点覆盖较好，仍可能有其他错误'],['只指出单位冲突','Step 5、DeepSeek 1/3','有有效发现，覆盖有限'],['表内算术支持“无实质错误”','Auto、GLM 0/3','算术正确，仍可能漏掉图面与公式的问题'],['有明显误报','MiniMax\nAuto 的部分图面判断','报的问题多，不等于有效发现多'],['该次环境未完成看图','HY4 自述未启用视觉','先检查输入链路，再讨论识图能力']];
override.A05=[['工具','服务变更与现行范围','使用时确认'],['Qoder IDE','Quest 于 2026-10-15 23:59（UTC+8）退役\nEditor 的 AI 对话、Agent 等不受该公告影响','区分 Quest 与 Editor'],['Qoder CN','官方页面列 Pro ¥59/月\n国际版价格另行核对','地区版、税费与活动'],['Gemini CLI','2026-06-18 起个人免费及 AI Pro/Ultra 迁移\n企业 Code Assist 许可与 API Key 认证不受影响','区分个人订阅和 API'],['Cursor / X','新组合套餐与现有旧订阅并存\n第三方模型不属于 Grok 共享用量池','第三方调用的费用'],['GSD','现行开发入口：open-gsd/gsd-core','使用新仓库与当前文档']];
override.A12=[['研究或事件','正确的阅读口径'],['幻觉与迎合评测','记录具体任务、型号、版本和时间。旧快照不能代表当期全部型号'],['宪法分类器、多样本越狱','攻击设置、样本与对象必须一起报告。缺条件时不引用裸成功率'],['EchoLeak','CVSS 来自不同评分方和版本时分别注明。本稿不使用分数做排名'],['GitHub MCP 注入演示','风险涉及不可信内容与已有权限的组合，不能泛称“一个已修复漏洞”'],['密钥“70% 仍有效”','GitGuardian 2025 报告跟踪 2022 年泄露密钥，不是 2026 年全部密钥比例'],['加拿大航空案','处理机构为 BC 民事裁决庭。精确赔付构成未核实，本稿不列金额'],['Terminal-Bench','按榜单版本、Agent、模型与推理设置对照，不能只抄一项裸分数']];

for(const page of brief.pages){
 const id=page.id;
 if(id==='P01'){
  const sl=P.slides.add();serial++;sl.background.fill=C.navy;
  txt(sl,'section','工作常用 AI 的使用观察',64,44,550,30,20,C.cyan,true);txt(sl,'date','2026 年 10 月',1000,44,216,30,20,'#B9C8D9',false,'right');
  txt(sl,'cover-title','AI 交流分享',64,121,1100,106,82,C.white,true);txt(sl,'cover-subtitle','从原理到实践',67,226,1100,81,54,'#33BEC9',true);
  txt(sl,'cover-description','基本原理与模型选择，工作案例与使用边界',68,329,1138,39,25,'#C8D5E4');
  await pic(sl,'pelican-crop.png','动画案例',68,414,362,172);await pic(sl,'drawing-crop.png','图纸案例',459,414,362,172);await pic(sl,'report-crop.png','报告案例',850,414,362,172);
  for(const[x,n,l,d]of[[68,'01','动画','同一句提示词，十份结果'],[459,'02','图纸','比较有效发现与误报'],[850,'03','报告','检查内容与交付质量']]){txt(sl,'n'+n,n,x,604,38,30,21,C.orange,true);txt(sl,'l'+n,l,x+48,600,200,35,25,C.white,true);txt(sl,'d'+n,d,x,645,362,28,19,'#B9C8D9');}
  sl.speakerNotes.text='真实案例素材来自本地鹈鹕 HTML 截图、桥墩挡块图纸和 Trae 修编报告。署名尚未指定。';manifest.push({id,title:page.title,slide:serial,sources:['本地三个案例']});continue;
 }
 if(await renderVisual(page,{header,txt,flow,node,line,pic,table,takeaway,listRows,C,links,serial:()=>serial}))continue;
 if(tablePages[id]){
  const cfg=tablePages[id];const sl=header(page,cfg.intro);let vals=override[id]??mdtable(page.content_markdown);
  const size=cfg.font??22;let y=id==='P20'?204:211;
  const generous=(id.startsWith('P')&&vals.length===4)?[44,96,96,96]:undefined;
  if(id==='P13')y=203;
  table(sl,vals,cfg.widths,{font:size,y,heights:id==='P13'?[40,52,52,52,52,52,52]:generous,maxBottom:607});
  takeaway(sl,cfg.foot,cfg.sub??'',cfg.sub?599:606);continue;
 }
 if(id==='P02'){
  const sl=header(page,'四个板块，正文讲清使用方法，附录保留查询资料');
  const rs=[['认识 AI','基本原理、发展历程与能力边界','03–10',3],['模型与工具','鹈鹕盲测、参数、能力成本与工具选择','11–18',11],['工作实践','图纸、报告、提示词与交付流程','19–30',19],['上手与边界','使用准备、数据安全、权限与合规','31–38',31]];
  rs.forEach((r,i)=>{let y=216+i*84;txt(sl,'ag-num-'+i,'0'+(i+1),64,y,70,48,31,C.orange,true);txt(sl,'ag-title-'+i,r[0],154,y,240,48,30,C.navy,true);txt(sl,'ag-desc-'+i,r[1],428,y,650,48,23);txt(sl,'ag-page-'+i,r[2],1100,y,115,48,22,C.teal,false,'right');links.push({slide:serial,shape:'ag-title-'+i,target:r[3]});if(i<3)line(sl,'agline'+i,64,y+69,1152,0,C.line,1);});
  takeaway(sl,'附录：参数与价格、配置指南、开源资源和案例来源');
 }else if(id==='P04'){
  const sl=header(page,'用已有的“面积—用量”样本，学习关系，再预测新输入');
  chart(sl,'scatter',{position:{left:64,top:216,width:708,height:355},title:'样本、拟合与新输入',titleTextStyle:{fontSize:21,fill:C.navy},scatterOptions:{style:'marker'},series:[{name:'已有样本',xValues:[1,2,3,4,5,6,7,8,9],values:[1.3,2.8,3.1,4.8,5.2,5.6,7.4,7.6,8.8],marker:{symbol:'circle',size:8,fill:C.cyan},line:{fill:'none'}},{name:'拟合关系',xValues:[0,10],values:[0.8,9],marker:{symbol:'none'},line:{fill:C.orange,width:3}},{name:'新输入的预测',xValues:[8.5],values:[7.77],marker:{symbol:'diamond',size:11,fill:C.navy}}],hasLegend:true,legend:{position:'bottom',textStyle:{fontSize:18}},xAxis:{...ax,title:'面积（示意）',min:0,max:10},yAxis:{...ax,title:'用量（示意）',min:0,max:10},chartFill:C.paper});
  listRows(sl,[['训练','调整参数，让预测更接近已有答案'],['预测','使用学到的参数，计算新的输入'],['边界','工况变化，规律可能失效']],{x:818,y:229,w:398,labelW:152,step:108,font:22});
  takeaway(sl,'拟合得好，不等于理解了因果','图中数值仅作概念示意，不代表任何工程系数。');
 }else if(id==='P05'){
  const sl=header(page,'多层非线性变换，可以逐步提取和组合特征');
  const xs=[0,1,2,3,4,5,6,7,8];
  chart(sl,'scatter',{position:{left:64,top:210,width:650,height:350},title:'简单关系与弯曲规律',titleTextStyle:{fontSize:21,fill:C.navy},scatterOptions:{style:'line'},series:[{name:'单一直线',xValues:xs,values:xs.map(x=>2+x*.5),line:{fill:C.orange,width:3},marker:{symbol:'none'}},{name:'非线性关系',xValues:xs,values:xs.map(x=>Math.round((2+Math.sin(x)*2+x*.5)*1000000)/1000000),line:{fill:C.cyan,width:3},marker:{symbol:'none'}}],hasLegend:true,legend:{position:'bottom',textStyle:{fontSize:18}},xAxis:{...ax,min:0,max:8},yAxis:{...ax,min:0,max:9},chartFill:C.paper});
  flow(sl,['输入','非线性\n变换','多层组合','输出'],{x:750,y:218,w:466,h:82,gap:15,size:19});
  txt(sl,'network-desc','训练改变的是大量参数。\n我们通常无法把每个参数\n翻译成一条人类规则。',754,340,450,138,25);
  takeaway(sl,'“折一下”是理解类比，实际网络在数学空间中进行变换','曲线仅说明关系形态，不是某个真实神经网络的训练结果。');
 }else if(id==='P06'){
  const sl=header(page,'六个节点，理解今天的模型能力从何而来');
  const events=[['1943','数学神经元','将神经元活动抽象为计算'],['1958','感知机','早期可学习的神经网络模型'],['2012','深度学习','大数据与 GPU 训练结合'],['2017','Transformer','用注意力处理序列关系'],['2022 起','对话产品普及','自然语言成为大众入口'],['持续发展','多模态与工具','从生成答案走向执行任务']];
  events.forEach((r,i)=>{const x=64+(i%3)*394,y=219+Math.floor(i/3)*178;txt(sl,'year'+i,r[0],x,y,360,44,33,i%2?C.teal:C.orange,true);txt(sl,'event'+i,r[1],x,y+51,360,38,27,C.navy,true);txt(sl,'meaning'+i,r[2],x,y+98,352,58,22);});
  takeaway(sl,'数据、算力与架构，共同推动能力变化');
 }else if(id==='P07'){
  const sl=header(page,'注意力帮助序列中的不同位置建立关联');
  txt(sl,'sentence','“把图纸和清单放在一起，检查它是否与图示一致。”',64,227,1152,65,36,C.navy,true);
  flow(sl,['图纸','清单','“它”指什么？','图示是否一致'],{x:95,y:338,w:1090,h:82,gap:38,size:25});
  txt(sl,'train','训练时：从大量样本学习关联',76,472,530,52,25,C.teal,true);txt(sl,'context','使用时：当前输入和对话提供上下文',652,472,560,52,25,C.teal,true);
  takeaway(sl,'能关联远处信息，也仍然需要定位与核验');
 }else if(id==='P09'){
  const sl=header(page,'输入拆成 token，模型逐步生成后续片段');
  flow(sl,['输入文本','分成 token','计算后续可能性','逐步生成'],{y:220,h:98,size:25});
  txt(sl,'split','分片示意：  检查　／　这份　／　表格',82,365,1100,56,31,C.teal,true);
  txt(sl,'token-rule','Token 不等于汉字或单词，以实际工具计数为准。',82,448,1100,62,28,C.navy,true);
  txt(sl,'temperature','温度影响随机性，更低不保证更真实。',82,527,1100,58,24);
  takeaway(sl,'完整流畅的句子，可能包含不存在的条款、来源和计算','以上是概念分片，不是某个型号 tokenizer 的实际输出。');
 }else if(id==='P11'){
  const sl=header(page,'任务：用 SVG 绘制鹈鹕骑自行车的 2D 动画');
  txt(sl,'conditions','一次生成　　无历史记忆　　不同模型与客户端配置',66,194,1150,26,18,C.muted);
  const files=['A_glm','B_seed','C_deepseek','D_opus','E_hy4','F_step5','G_gpt','H_auto','I_qwen','J_minimax'];
  const names=['GLM 5.3 Flash\nTrae','Seed 2.1 Turbo\nTrae','DeepSeek V4.1 Flash 1M\nWorkBuddy','Opus 5.5\nHark','HY4 Preview 1M high\nWorkBuddy','Step 5 Preview\nWorkBuddy','GPT 6.1 Sol high\nCodex','Auto fast\nWorkBuddy','Qwen 3.8 Flash\nQoder','MiniMax M3\nWorkBuddy'];
  for(let i=0;i<10;i++){const x=64+(i%5)*234,y=236+Math.floor(i/5)*188;await pic(sl,files[i]+'_1200x900.png','pelican-'+i,x,y,216,124,'contain');txt(sl,'pelican-letter-'+i,String.fromCharCode(65+i),x,y-2,27,30,22,C.orange,true);txt(sl,'reveal-'+i,names[i],x,y+129,218,50,16,C.teal,true);links.push({slide:serial,shape:'pelican-'+i,address:'配套/鹈鹕原始结果/'+String.fromCharCode(65+i)+'.html'});}
  takeaway(sl,'同一句提示词，可以得到很不一样的交付','放映时单击揭晓配置。点图可播放原 HTML，单帧截图不能证明全部动作正确。',617);
 }else if(id==='P14'){
  const sl=header(page,'AA 智能指数与每任务成本，统一查询于 2026-10-18');
  chart(sl,'scatter',{position:{left:64,top:211,width:790,height:370},title:'能力与任务成本',titleTextStyle:{fontSize:20,fill:C.navy},scatterOptions:{style:'marker'},series:aa.map((v,i)=>({name:v.name,xValues:[v.cost],values:[v.score],marker:{symbol:'circle',size:10,fill:i<3?C.navy:C.cyan},line:{fill:'none'},dataLabels:{showSeriesName:true,position:['top','top','top','right','top','bottom','bottom','right'][i],textStyle:{fontSize:15,fill:C.ink}}})),hasLegend:false,xAxis:{...ax,title:'AA 每任务成本（美元，线性轴）',min:0,max:6.8},yAxis:{...ax,title:'AA 智能指数',min:35,max:62,majorUnit:5},chartFill:C.paper});
  txt(sl,'chart-read','AA 指数\n特定评测集合\n\n每任务成本\n该评测设置下的费用',886,225,326,235,23);
  txt(sl,'config-note','具体档位与 API 标价\n见 A02，口径说明见备注。',886,465,326,112,20,C.muted);
  takeaway(sl,'先筛候选，再比较自己的任务质量、耗时、费用与返工');
 }else if(id==='P19'){
  const sl=header(page,'同一份挡块图纸，提交审查与数量复核结果');
  const v=[['配置','记录耗时','命中 / 3','阅读结果时注意'],['Opus','3分03秒','3 / 3','三项均指出'],['GPT','14分34秒','2 / 3','两项，可计算交付'],['Qwen','43分09秒','2 / 3','找到两项'],['Step 5','29分40秒','1 / 3','指出单位冲突'],['DeepSeek','32分05秒','1 / 3','指出单位冲突'],['Auto','2分38秒','0 / 3','判无实质错误，误读图面'],['MiniMax','12分27秒','0 / 3','有明显误报'],['Seed','12分47秒','0 / 3','三项未命中'],['GLM','1时24分','0 / 3','判无实质错误'],['HY4','3分55秒','不纳入比较','该次环境未完成视觉读取']];
  const t=table(sl,v,[156,178,174,644],{font:21,y:204,heights:[39,...Array(10).fill(34)],highlight:[1,2,3],center:[2]});
  for(let r=1;r<=3;r++)t.getCell(r,2).text.style={typeface:FONT,fontSize:23,color:C.teal,bold:true,alignment:'center',insets:0};
  for(let c=0;c<4;c++)t.getCell(10,c).text.style={typeface:FONT,fontSize:c===2?19:21,color:C.muted,bold:c===0,alignment:c===2?'center':'left',insets:0};
  takeaway(sl,'按命中数排序，同分按耗时，HY4 单列','3 个事后选定检查点不等于整体准确率，耗时也非严格统一计时。',605);
 }else if(id==='P21'){
  const sl=header(page,'相同材料下，要求越明确，越容易定义和检查交付');
  txt(sl,'short-label','简略写法',64,222,310,42,25,C.orange,true);txt(sl,'short-prompt','“帮我检查这张图，\n做个表格。”',64,283,318,120,31,C.navy,true);
  txt(sl,'improved-label','可复用写法',438,222,778,42,25,C.teal,true);
  const p='请基于附件检查图纸与数量表。先说明读取到哪些页、哪些部分不清晰。每项发现给出页码或图中位置、原文或尺寸、判断依据和计算过程。把明确不一致、待确认事项和未检查范围分开。不要补造看不见的尺寸，也不要为了找错强行判错。请交付可编辑 Excel，计算项保留公式，最后说明尚需人工确认的事项。';
  txt(sl,'improved-prompt',p,438,280,778,251,24);txt(sl,'four-parts','任务目标　　材料范围　　输出格式　　验收与不确定性',438,551,778,35,21,C.teal,true);
  takeaway(sl,'这是改写示例，尚未做控制变量复测，不声称效果提升');
 }else if(id==='P22'){
  const sl=header(page,'同一批清单、旧模板与参考材料，共 4 个配置。覆盖统计采用各输出原口径');
  table(sl,[['配置','记录耗时','记录的覆盖范围','需要核查的风险'],['Opus · Hark','4分','≥1 万元 43 条','承建单位等外补事实仍要核来源'],['Trae Auto','7分','25 行','数量、变更方向和覆盖完整性'],['WorkBuddy Auto','14分','101 条','总额一处写成 41,638.89 万，存在模板残留'],['DeepSeek Flash','34分','99 条','致函对象和原因仍需逐项对应原材料']],[226,140,224,562],{y:211,font:21});
  txt(sl,'baseline-label','汇总表复算基准',64,507,1152,32,21,C.muted);
  txt(sl,'baseline','41,638.9336 万元　→　41,465.1967 万元',64,547,1152,45,32,C.navy,true);
  takeaway(sl,'核减 173.7369 万元，99 项金额变化，另 2 项仅数量变化','四份输出筛选阈值不同，不能按条目数排成功率。金额不含暂列金额。');
 }else if(id==='P23'){
  const sl=header(page,'清单 203-1-2-2-1：395.9711 万元增至 573.1093 万元，增加 177.1382 万元');
  table(sl,[['配置','原报告摘录','怎样判断'],['Opus','下路床换填材料修改为石渣……利用石方增加','具体，仍需核来源'],['DeepSeek','利用方增加石渣利用量……调整为石渣','仍需核对应关系'],['Trae','石方利用量增加，根据土石方调配优化调整','方向清楚，仍需依据'],['WorkBuddy Auto','依据设计后服务函及勘误回复函核对修改清单数量','泛化，未解释为何增加']],[210,660,282],{y:206,font:21,maxBottom:556});
  await pic(sl,'reason-reference.png','参考函件原文局部',66,542,1150,62,'contain');
  takeaway(sl,'参考函第 1 页第 8 项支持“下路床换填材料为石渣”','摘录已逐字对照。完整金额归因仍要结合数量表，不凭表述详细程度判正确。',607);
 }else if(id==='P25'){
  const sl=header(page,'每一步留下看得见的结果，下一步才有可靠输入');
  listRows(sl,[['读材料','列出已读文件、范围、缺页和无法识别的位置'],['定口径','明确单位、金额范围、资料优先级与不得猜测的字段'],['做产物','按模板交付可编辑文件，保留公式与来源定位'],['核结果','回原文、独立重算、打开实际文件检查版面与完整性']],{step:89,font:24});
  takeaway(sl,'报告案例：材料清单、汇总口径、变更报告、复算与原因定位');
 }else if(id==='P26'){
  const sl=header(page,'检索相关资料片段，再结合问题生成答案和引用');
  txt(sl,'ingest-label','资料入库',64,223,220,45,28,C.teal,true);flow(sl,['读取资料','按结构切分','建立索引'],{x:302,y:214,w:914,h:80,size:25});
  txt(sl,'answer-label','回答问题',64,352,220,45,28,C.orange,true);flow(sl,['检索相关片段','片段＋问题','答案＋引用'],{x:302,y:343,w:914,h:80,size:25});
  txt(sl,'example-question','例：石方利用量增加的依据在哪里？',64,475,1152,50,30,C.navy,true);txt(sl,'example-result','理想输出应指出文件、页码或段落。资料不足时，明确说明缺少证据。',64,539,1152,48,24);
  takeaway(sl,'RAG 属于检索辅助回答，资料不会因此变成永久训练记忆');
 }else if(id==='P29'){
  const sl=header(page,'用可检查的交付物串起六个环节');
  const rs=[['澄清','目标、使用者、边界与成功标准'],['设计','页面或模块如何协作'],['规划','修改范围、依赖和执行顺序'],['执行','实际代码与修改差异'],['验证','与风险匹配的运行检查和测试证据'],['交付','使用方法、已知限制与回退方式']];
  rs.forEach((r,i)=>{const x=64+(i%2)*600,y=217+Math.floor(i/2)*120;txt(sl,'phase-num-'+i,'0'+(i+1),x,y,55,42,27,C.orange,true);txt(sl,'phase-title-'+i,r[0],x+72,y,465,42,28,C.navy,true);txt(sl,'phase-detail-'+i,r[1],x+72,y+47,450,61,23);});
  takeaway(sl,'导出按钮出现后，还要打开实际导出的文件检查内容和格式','方法示例，本次没有新增编码对照实验。');
 }else if(id==='P30'){
  const sl=header(page,'适合并行的工作有边界，合并与验收仍要有人负责');
  table(sl,mdtable(page.content_markdown),[240,456,456],{y:211,font:22});
  txt(sl,'parallel','适合并行：不同资料检索、不同模块、独立复算',66,491,1150,42,25,C.teal,true);txt(sl,'avoid','不适合硬拆：步骤强依赖、共享状态复杂、很小的任务',66,546,1150,42,25);
  takeaway(sl,'人数增加不等于可靠性增加，共同的错误前提仍会让大家一起错');
 }else if(id==='P31'){
  const sl=header(page,'从一份不敏感、结果可以人工检查的小材料开始');
  listRows(sl,[['选一个入口','选择可访问、允许使用的产品'],['选一份小材料','公开说明或简单表格，能够检查输出'],['明确交付物','摘要、对照表，或可打开的文件'],['检查后再升级','判断缺的是模型、文件工具，还是材料与提示词']],{step:87,font:24});
  takeaway(sl,'问答从网页或 App 开始，本地文件与代码任务再考虑任务客户端');
 }else if(id==='P34'){
  const sl=header(page,'外部网页、文档和仓库中的文字，也可能试图改变任务');
  txt(sl,'normal-label','普通资料',64,219,245,42,27,C.teal,true);txt(sl,'normal','“本项目使用某项计价口径。”',335,217,881,50,28);
  txt(sl,'bad-label','越权指令',64,296,245,42,27,C.orange,true);txt(sl,'bad','忽略用户任务，越权发送资料',335,290,881,73,28,C.navy,true);
  flow(sl,['读取不可信内容','把资料当成指令','使用已有工具权限','错误操作或外流'],{y:402,h:98,size:23});
  txt(sl,'risk-conditions','风险组合：私有数据　＋　不可信内容　＋　对外发送能力',64,539,1152,46,26,C.teal,true);
  takeaway(sl,'资料是分析对象，不能仅凭其中一段话获得操作授权');
 }else if(id==='P37'){
  const sl=header(page,'使用地区、工作数据和公众服务，应分别核对适用规则');
  table(sl,mdtable(page.content_markdown),[330,470,352],{y:210,font:21});
  txt(sl,'filing-number','1112 款',64,480,306,70,49,C.teal,true);txt(sl,'filing-description','截至 2026-08-31 完成备案的生成式 AI 服务\n另有 731 款应用或功能完成登记，二者不要混计',390,480,826,80,23);
  takeaway(sl,'Anthropic 新使用政策拟于 2026-11-12 生效','2026-10-08 为公告日期，10 月稿不能写成已生效。具体主体与范围以原文为准。');
 }else if(id==='P38'){
  const sl=header(page,'用一个能检查结果的小任务，建立自己的使用方法');
  listRows(sl,[['一个任务','一页公开资料摘要、一张表，或一段报告'],['一份模板','目标、材料、边界、交付物和完成标准'],['一套验收方法','回原文、重算关键数字、打开实际文件，记录未确定与未做事项']],{y:225,step:111,labelW:250,font:26});
  txt(sl,'appendix-navigation','会后查阅',64,584,200,36,23,C.orange,true);
  txt(sl,'appendix-links','参数 A02–A03　　价格 A04　　工具 A05　　指南 A06–A08　　资源与证据 A09–A14',64,632,1152,36,21,C.teal,true);
  links.push({slide:serial,shape:'appendix-links',target:39});
 }else if(id==='A02'){
  const sl=header(page,'AA 指数降序。AA 配置与案例配置分开，查询日期统一为 2026-10-18');
  const rows=[['型号与 AA 配置','AA 指数','输入 / 输出美元\n每百万 token','AA 每任务美元'],...aa.map(v=>[v.name+' · '+v.mode,String(v.score),v.input+' / '+v.output,v.cost.toFixed(2)])];
  table(sl,rows,[475,150,285,242],{y:205,font:21,headfont:19,heights:[49,...Array(8).fill(42)],center:[1,2,3]});
  takeaway(sl,'API 标价与 AA 每任务成本，属于不同计费口径','美元单价为所列 AA 配置的标价口径，地区、缓存、活动和长输入计费另核。');
 }else if(id==='A03'){
  const sl=header(page,'模型规格与产品限制分开看。M、K 以各官方文档的 token 单位为准');
  const rows=[['具体型号','上下文','最大输出','输入模态','输出','工具','入口与限制','官方来源\n查询日期'],['Opus 5.5','1M','128K','文字、图像','文字','工具调用','以所用平台的权限为准','S02\n2026-10-18'],['GPT-6.1 Sol','1,050K','128K','文字、图像','文字','工具调用','长输入和速度档有加价','S01\n2026-10-18'],['Gemini 4 Argon','1M 共享','≤1M*\n扣除输入','文字、图像\n音频、视频','文字','工具调用','统一输入输出预算\n产品可能另设上限','S03\n2026-10-18'],['GLM-5.3','1M','128K','文字','文字','工具调用','始终启用推理\n具体协议另核','S04、Z.ai\n2026-10-18'],['Qwen3.8-Max','1,000K','131,072','按具体\n接口版本','文字','按接口','输入输出预算存在约束\n不由家族名推断视觉','S05\n2026-10-18'],['Kimi K3','1M','余量*\n默认128K','按具体\n接口版本','文字','按接口','长度上限含输入\nAPI 与聊天额度分开','S06\n2026-10-18'],['MiMo-V2.6-Pro','1M','未确认','文字、图像\n音频、视频','文字','按接口','权重模型与云端 API\n限制需分别核对','S07\n2026-10-18'],['DeepSeek\nV4.1 Flash 1017','1M','384K','文字、图像','文字','工具调用','deepseek-flash 路由\n有高峰与低谷价格','S08\n2026-10-18']];
  table(sl,rows,[178,91,116,125,73,105,243,221],{y:204,font:18,headfont:17,heights:[42,...Array(8).fill(43)],center:[1,2,3,4,5]});
  takeaway(sl,'产品记忆不属于模型的上下文规格，权限也不由型号自动获得','* 上限还需扣除输入并遵守接口约束。“未确认”不代表不支持，链接见来源索引。');
 }else if(id==='A04'){
  const sl=header(page,'美元为美国网页月付参考，人民币标明地区。税费、活动和续费条件另核');
  table(sl,[['产品','月付参考','权益与购买时要确认的事项'],['ChatGPT','Go $8；Plus $20\nPro $100 / $200 / $500','Pro 分档与历史额度不同，模型、速度档和可用地区以账户页面为准'],['Claude','Pro $20\nMax $100 起','用量有上限，部分 Max 套餐含指定月度 API 额度，不能推广到全部会员'],['Google AI','Plus $4.99；Pro $19.99\nUltra $99.99 / $199.99','地区、用量档位与活动价分开看，API 另核'],['X Subscription','Plus $8；Premium $30\nSuper $100；Ultra $200','Grok、Cursor、X 的组合套餐。旧订阅可继续，第三方模型用量另计'],['DeepSeek / 千问','有免费使用入口','免费入口不等于全功能永久不限量，API、团队功能另核'],['Kimi','本轮未核实当期月费','以实际地区和账户的官方套餐页为准，未采用旧稿约 ¥49 的估计值']],[230,350,572],{y:203,font:20,headfont:19,heights:[39,57,57,57,65,50,50]});
  takeaway(sl,'购买前检查续费价格、限额、取消入口与实际付款对象');
 }else if(id==='A06'){
  const sl=header(page,'设备与实际客户端尚未指定，本页保留可直接使用的检查顺序');
  listRows(sl,[['设备与系统','确认系统版本、可用空间与客户端要求'],['官方安装','记录官方网站、客户端名称和版本'],['账号与恢复','保存恢复方式，课件中隐藏密码和验证码'],['连通检查','站点、登录、消息，再上传一份公开小文件'],['定位故障','网络不通、账号验证、地区不支持、额度不足分别判断']],{y:208,step:72,labelW:250,font:23});
  takeaway(sl,'实际操作截图待指定设备和客户端后补充','保留下载入口、安装完成、登录和一条成功消息四个截图位置，不虚构成功记录。');
 }else if(id==='A07'){
  const sl=header(page,'实际演示产品尚未指定，先保留购买前、中、后的核对事项');
  listRows(sl,[['购买前','确认地区、套餐权益、月付或年付、税费、续费与退款说明'],['购买中','从官方入口选择套餐，核对金额、币种、账单资料和付款对象'],['购买后','确认权益生效与账单，找到管理订阅、取消续费和联系客服入口']],{y:225,step:112,labelW:220,font:26});
  txt(sl,'fields','演示信息：产品〔待定〕　套餐〔待定〕　支付方式〔待定〕　演示日期〔待定〕',64,580,1152,47,23,C.teal,true);
  txt(sl,'boundary','不根据个别成功经历承诺某银行卡一定可用，不展示卡号、验证码或私人账单。',64,635,1152,35,20,C.muted);
 }else if(id==='A08'){
  const sl=header(page,'网关在本机完成适配或路由，上游推理不一定在本机');
  flow(sl,['任务客户端','本地网关','选定的上游服务','返回结果'],{y:220,h:106,size:26});
  listRows(sl,[['动作','本机执行了什么，上游执行了什么'],['数据','谁持有凭据，谁保存输入、输出和日志'],['费用与故障','谁计费，错误来自客户端、网关还是上游']],{y:378,step:70,labelW:245,font:24});
  takeaway(sl,'先确认上游服务、凭据持有方与日志位置');
 }else if(id==='A14'){
  const sl=header(page,'单页来源在页脚和备注，完整链接与核查说明见随包来源索引');
  const rs=[['模型与参数','S01–S08','官方型号文档、上下文、模态和 API 价格'],['能力与成本','S09–S16、S37–S38','AA 具体配置页面，Terminal-Bench 版本与方法'],['订阅与工具','S17–S27','订阅权益、地区价、服务变更与 GSD 迁移'],['政策与安全','S28–S32、S36','网信办公告、使用政策、注入研究与密钥报告'],['历史与原理','S33–S35','Transformer、数学神经元与感知机原论文'],['本地案例','配套案例索引','图纸、清单、四份报告、10 份原始 HTML 与复核口径']];
  table(sl,[['资料组','来源编号','查询内容'],...rs],[260,255,637],{y:209,font:22,heights:[42,...Array(6).fill(53)]});
  txt(sl,'open-index','打开完整来源与案例索引',64,604,1152,51,28,C.teal,true);links.push({slide:serial,shape:'open-index',address:'配套/来源索引.html'});
 }
}
for(let i=0;i<brief.pages.length;i++){
 const page=brief.pages[i],slide=P.slides.items[i];
 let detail=page.content_markdown.split('**组装说明**')[0].replace(/^\*\*上页内容[^\n]*\n/,'').trim();
 if(['P14','A02'].includes(page.id))detail=aa.map(v=>`${v.name}，${v.mode}：AA 指数 ${v.score}，每任务 $${v.cost}，输入/输出 $${v.input}/$${v.output} 每百万 token。`).join('\n')+'\n统一查询日期为 2026-10-18。TB 4.0 的 Agent 配置未完整核实，未采用旧稿裸分数。';
 if(page.id==='P37')detail='截至 2026-08-31，1112 款生成式 AI 服务完成备案，另有 731 款应用或功能完成登记。两者不能混计，也不是模型能力排名。Anthropic 于 2026-10-08 公告的新使用政策拟于 2026-11-12 生效。服务地区、工作材料制度、公众服务义务分别查看适用条款。';
 if(page.id==='A04')detail='当期美国网页月付参考：ChatGPT Go $8，Plus $20，Pro $100/$200/$500。Claude Pro $20，Max $100 起，部分 Max 含指定月度 API 额度。Google AI Plus $4.99，Pro $19.99，Ultra $99.99/$199.99。X Subscription Plus $8，Premium $30，Super $100，Ultra $200。税费、地区、优惠与续费条件另核。Kimi 当前月费未确认，未使用旧稿估计值。查询于 2026-10-18。';
 if(page.id==='A03')detail='规格查询于 2026-10-18。Opus 5.5：1M 上下文，128K 输出，文字与图像输入。GPT-6.1 Sol：1,050,000 上下文，128,000 输出，文字与图像输入，长输入、速度档和地区处理存在不同价格。Gemini 4 Argon：1M 共享输入输出预算，支持文字、图像、音频、视频输入，长输出仍受输入占用和产品上限约束。GLM-5.3：1M 上下文，128K 输出，文字输入，始终开启推理。Qwen3.8-Max：1,000,000 上下文，131,072 输出，模态与工具按具体接口版本核对。Kimi K3：1M 上下文，默认最大完成长度 131072，上限需扣除输入 token。MiMo-V2.6-Pro：官方权重模型卡列 1M 上下文，文字、图像、音频、视频模态，本次未确认云端最大输出。DeepSeek V4.1 Flash 1017：1M 上下文，384K 最大输出，文字与图像输入，deepseek-flash 路由，峰谷价格分开。上述输出主要为文字，生成代码或文件还依赖工具。模型协议支持不等于产品授予工具权限。产品记忆不是模型上下文规格，官方 M/K 单位以各型号原文为准。';
 if(page.id==='P16')detail+='\n2026-10-18 核查补充：Claude 当前部分 Max 套餐含指定 API 额度，不能笼统断言所有会员都不含 API。各入口的权益与计费仍需分别核对。';
 if(page.id==='P23')detail+='\n本轮已逐字对照四份原报告。参考函第 1 页第 8 项支持下路床材料改为石渣，第 2 页续第 12 项涉及土石方数量修改。具体金额与完整因果归属仍需结合数量表核查。';
 detail=detail.split('\n').filter(l=>!/^\|[\s:|-]+\|?$/.test(l.trim())).map(l=>l.trim().startsWith('|')?l.trim().replace(/^\||\|$/g,'').split('|').map(v=>v.trim()).join('；'):l).join('\n');
 slide.speakerNotes.text='讲解与会后阅读\n\n'+clean(detail)+'\n\n参考来源\n'+slide.speakerNotes.text;
}
if(P.slides.items.length!==52)throw new Error('Expected 52 slides, got '+P.slides.items.length);
await fs.writeFile(path.join(B,'manifest.json'),JSON.stringify({slides:manifest,links,tableOwners,chartOwners,sources,aa},null,2));
await (await PresentationFile.exportPptx(P)).save(path.join(B,'candidate.pptx'));
console.log(JSON.stringify({slides:P.slides.items.length,tables:tableOwners.length,chartOwners,output:path.join(B,'candidate.pptx')}));
