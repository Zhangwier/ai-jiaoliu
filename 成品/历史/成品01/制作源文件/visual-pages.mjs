// Main slides use a visual summary. Full explanations live in speaker notes.
export async function renderVisual(page,h){
 const {header,txt,flow,node,line,pic,table,takeaway,listRows,C}=h;
 const id=page.id;
 const ids=['P03','P08','P10','P12','P13','P15','P16','P17','P18','P20','P21','P22','P24','P25','P27','P28','P30','P31','P32','P33','P35','P36','P37','P38','A03','A04','A06','A07','A11','A12','A13'];
 if(!ids.includes(id))return false;
 const sl=header(page,'');
 const t=(name,value,x,y,w,hh=60,size=28,color=C.ink,bold=false)=>txt(sl,name,value,x,y,w,hh,size,color,bold);
 function columns(items,{y=225,w=1152,x=64,size=42,subSize=24,gap=48}={}){
  const cw=(w-gap*(items.length-1))/items.length;
  items.forEach((r,i)=>{const xx=x+i*(cw+gap);t('colnum'+i,'0'+(i+1),xx,y,cw,45,24,C.orange,true);t('coltitle'+i,r[0],xx,y+65,cw,86,size,C.navy,true);t('coldetail'+i,r[1],xx,y+170,cw,124,subSize,C.muted);});
 }
 if(id==='P03'){
  columns([['选配置','模型、工具、权限'],['给材料','目标、源文件、交付物'],['验结果','回原文、重算、打开文件']],{y:206,size:44});
  t('case-line','动画能否完成？　　图纸能否发现问题？　　报告能否交付？',64,582,1152,67,27,C.teal,true);
 }else if(id==='P08'){
  flow(sl,['预训练','监督微调','偏好优化／\n强化学习等'],{y:239,h:104,size:31,gap:58});
  const items=[['学习通用模式','语言与知识'],['学习示范回答','指令与格式'],['依据反馈调整','有用性与推理表现']];
  items.forEach((r,i)=>{const x=66+i*405;t('train1'+i,r[0],x,385,340,60,28,C.navy,true);t('train2'+i,r[1],x,458,340,60,24,C.muted);});
  takeaway(sl,'专业判断仍然需要依据与核验','各厂商的训练组合不同，图中为简化说明。');
 }else if(id==='P10'){
  t('hallucination','幻觉',64,225,510,100,66,C.navy,true);t('sycophancy','迎合',690,225,510,100,66,C.teal,true);
  t('hallucination-example','编出不存在的条款、来源',64,352,510,65,28);t('sycophancy-example','顺着“这里肯定有错”找理由',690,352,510,65,28);
  t('hallucination-action','定位原文',64,469,510,68,36,C.orange,true);t('sycophancy-action','检验反例',690,469,510,68,36,C.orange,true);
  takeaway(sl,'允许“材料不足”，关键事实回到原始资料');
 }else if(id==='P12'){
  flow(sl,['厂商／家族','型号与档位','产品入口','部署方式'],{y:242,h:108,size:28});
  const a=['谁提供模型','实际运行什么','网页、App 或 API','云端或本地'];
  a.forEach((v,i)=>t('model-layer'+i,v,64+i*299,390,255,90,26,C.muted));
  takeaway(sl,'开放权重　　许可证　　本地算力','三个条件分别确认，不能由品牌名称直接推断。');
 }else if(id==='P13'){
  const terms=[['上下文','一次能容纳多少信息'],['最大输出','一次能生成多少'],['输入模态','文字、图像、音频、视频'],['输出模态','模型直接输出什么'],['记忆','产品跨会话保留什么'],['推理与工具','思考预算和操作权限']];
  terms.forEach((r,i)=>{const x=64+(i%2)*600,y=208+Math.floor(i/2)*123;t('param'+i,r[0],x,y,210,60,32,C.navy,true);t('paramdesc'+i,r[1],x+224,y+4,330,72,24,C.muted);});
  takeaway(sl,'上传成功，还要确认图形和尺寸线是否真正进入读取链');
 }else if(id==='P15'){
  listRows(sl,[['看图','图像细节　文字识别　空间关系'],['表格','读取完整性　公式执行　单位'],['报告','资料覆盖　条款保留　引用'],['代码','环境操作　需求遵循　实际运行'],['问答','检索片段　来源定位　承认缺口']],{y:205,step:74,labelW:215,font:27});
  takeaway(sl,'同一材料做小样，同时看质量、费用与返工');
 }else if(id==='P16'){
  columns([['聊天会员','功能与用量'],['API','按调用量计费'],['工具订阅','客户端＋指定额度']],{y:206,size:43,subSize:28});
  takeaway(sl,'不同入口的权益与额度，分别核对','部分套餐含指定 API 额度，范围和条件以当期条款为准。');
 }else if(id==='P17'){
  const a=[['聊天入口','DeepSeek　千问　Kimi\nChatGPT　Claude　Gemini'],['任务客户端','Codex　Claude Code　Cursor\nQoder　Trae　WorkBuddy'],['工作流','n8n　Dify'],['文档处理','PaddleOCR　MinerU　MarkItDown']];
  a.forEach((r,i)=>{let x=64+(i%2)*600,y=207+Math.floor(i/2)*182;t('tooltype'+i,r[0],x,y,550,60,34,C.navy,true);t('tools'+i,r[1],x,y+68,550,87,25,C.teal);});
  takeaway(sl,'先确定任务，再看文件能力、执行权限与数据流向');
 }else if(id==='P18'){
  flow(sl,['用户任务','模型判断','工具执行','结果反馈'],{y:237,h:112,size:29});
  line(sl,'environment-span',362,409,851,0,C.cyan,3);
  t('environment','运行环境：上下文、调度、权限与反馈',362,432,851,65,29,C.teal,true);
  t('skill-mcp','Skill 约定完成方式　　　　MCP 连接工具和数据',64,535,1152,53,26,C.muted);
  takeaway(sl,'实际表现来自模型、客户端与设置的共同作用');
 }else if(id==='P20'){
  columns([['3 / 3','Opus'],['2 / 3','GPT　Qwen'],['1 / 3','Step 5　DeepSeek']],{y:208,size:64,subSize:30});
  t('other-results','Auto、MiniMax、Seed、GLM：0 / 3　　HY4：本次环境未完成看图',64,554,1152,58,23,C.muted);
  takeaway(sl,'命中、漏报与误报分开看，表内算术正确仍可能看错图');
 }else if(id==='P21'){
  t('briefprompt','“帮我检查这张图，做个表格。”',64,212,1152,89,39,C.navy,true);
  const a=[['任务目标','检查图纸与数量表'],['材料范围','先说明读取范围与缺口'],['输出格式','可编辑 Excel，保留公式'],['验收要求','给位置、依据与不确定性']];
  a.forEach((r,i)=>{let x=64+(i%2)*600,y=352+Math.floor(i/2)*107;t('prompt-key'+i,r[0],x,y,195,56,28,C.teal,true);t('prompt-value'+i,r[1],x+210,y,340,74,25);});
  takeaway(sl,'改写示例，尚未做控制变量复测');
 }else if(id==='P22'){
  const items=[['Opus · 4分','hark-opus5.5-4m-excerpt.png','43条，≥1万元','外补事实需核来源'],['Trae · 7分','trae-auto-7m-excerpt.png','25行','数量、方向与覆盖'],['Auto · 14分','workbuddy-auto-flat-14m-excerpt.png','101条','总额一处抄错'],['DeepSeek · 34分','workbuddy-deepseek-v4.1flash-34m-excerpt.png','99条','原因需对应原材料']];
  for(let i=0;i<4;i++){let x=64+i*295;const r=items[i];t('report-name'+i,r[0],x,210,268,48,25,C.navy,true);await pic(sl,r[1],'report-excerpt-'+i,x,277,264,142,'contain');t('report-count'+i,r[2],x,441,268,48,24,C.teal,true);t('report-risk'+i,r[3],x,496,268,60,23,C.muted);}
  takeaway(sl,'汇总表（不含暂列金额）核减 173.7369 万元','覆盖阈值不同，不能按条目数排成功率。99 项金额变化，另 2 项仅数量变化。');
 }else if(id==='P24'){
  flow(sl,['原始 PDF / Excel','提取结构','执行计算','实际文件'],{y:234,h:111,size:28});
  const a=[['读取','漏页、漏表'],['结构','错列、合并错位'],['计算','类型、单位、公式']];
  a.forEach((r,i)=>{const x=64+i*407;t('failure'+i,r[0],x,405,330,65,38,C.navy,true);t('failure-detail'+i,r[1],x,487,330,65,25,C.orange,true);});
  takeaway(sl,'打开文件，对照原页，并独立复算关键数字');
 }else if(id==='P25'){
  flow(sl,['读材料','定口径','做产物','核结果'],{y:238,h:119,size:32});
  const a=['文件与缺口','单位与范围','可编辑文件','原文与复算'];
  a.forEach((v,i)=>t('deliverable'+i,v,64+i*299,401,254,88,27,C.teal,true));
  takeaway(sl,'每一步留下可检查的结果');
 }else if(id==='P27'){
  flow(sl,['读取与切分','检索片段','组装上下文','回答与引用'],{y:215,h:92,size:26});
  t('acceptance-heading','最小验收题集',64,370,1152,64,35,C.navy,true);
  ['有明确答案','跨段落','新旧版本冲突','资料中无答案'].forEach((v,i)=>t('acceptance'+i,v,64+i*298,470,263,85,28,i===3?C.orange:C.teal,true));
  takeaway(sl,'先查检索片段，再检查回答与引用');
 }else if(id==='P28'){
  t('workflow-label','固定工作流',64,211,240,62,36,C.navy,true);flow(sl,['输入','预定义条件','结果'],{x:352,y:205,w:864,h:87,size:27});
  t('agent-label','Agent',64,377,240,62,39,C.teal,true);flow(sl,['观察','判断','操作','反馈'],{x:352,y:371,w:864,h:87,size:27,gap:27});
  t('wf-benefit','稳定、可追踪',352,306,864,48,23,C.muted);t('agent-benefit','适合需要探索和反复调整的任务',352,475,864,48,23,C.muted);
  takeaway(sl,'输入检查、计算与审批可固定，判断环节交给模型');
 }else if(id==='P30'){
  node(sl,'lead','主负责人',451,206,378,86,'#E4F2F2',34);
  line(sl,'branch-a',324,334,317,0,C.cyan,2);line(sl,'branch-b',641,334,315,0,C.cyan,2);line(sl,'lead-line',641,292,0,42,C.cyan,2);line(sl,'down-a',324,334,0,27,C.cyan,2);line(sl,'down-b',956,334,0,27,C.cyan,2);
  node(sl,'executor1','有边界的执行任务',109,361,430,100,C.white,29);node(sl,'reviewer','独立复核',741,361,430,100,C.white,29);
  t('executor-caption','分开文件与工作范围',109,481,430,58,24,C.muted);t('reviewer-caption','回到源材料检查',741,481,430,58,24,C.muted);
  takeaway(sl,'负责人合并并验收，人数增加不自动提高可靠性');
 }else if(id==='P31'){
  flow(sl,['选入口','给小材料','明确交付','检查结果'],{y:239,h:118,size:32});
  t('first-task','一页公开说明　／　一张简单表格',64,418,1152,91,40,C.teal,true);
  takeaway(sl,'先完成一个能人工判断结果的小任务，再决定是否升级工具');
 }else if(id==='P32'){
  const a=[['地区','服务是否支持'],['账号','验证与恢复'],['网络','登录、消息、文件'],['支付','账单、地区、套餐']];
  a.forEach((r,i)=>{let x=64+(i%2)*600,y=205+Math.floor(i/2)*181;t('setup'+i,r[0],x,y,550,85,54,C.navy,true);t('setup-desc'+i,r[1],x,y+94,550,60,28,C.teal);});
  takeaway(sl,'分别确认四个环节，再开始第一个任务');
 }else if(id==='P33'){
  flow(sl,['客户端','本地网关','上游模型服务'],{y:244,h:126,size:34,gap:67});
  t('data-route','数据发给谁？',64,435,530,90,41,C.navy,true);t('log-route','日志留在哪里？',690,435,526,90,41,C.teal,true);
  takeaway(sl,'网关在本机，不代表上游推理也在本机');
 }else if(id==='P35'){
  listRows(sl,[['EchoLeak','外部内容进入企业 AI 的信任链'],['GitHub MCP 演示','不可信资料滥用已有权限'],['密钥泄露','删除文本不等于撤销凭据']],{y:218,step:117,labelW:405,font:29});
  takeaway(sl,'工具连接可用，仍需核对内容可信度与操作权限');
 }else if(id==='P36'){
  const rows=[['公开资料','核来源与准确性',64,236,452],['内部材料','脱敏，使用组织允许的入口',119,350,397],['密钥与高敏感资料','专门凭据机制与明确授权',174,464,342]];
  rows.forEach((r,i)=>{node(sl,'data-level'+i,r[0],r[2],r[3],r[4],88,i===2?'#FBEBDD':C.white,28);t('data-action'+i,r[1],620,r[3],596,88,29,i===2?C.orange:C.teal,true);});
  takeaway(sl,'对外发送、付款、删除等动作，最终执行前核对对象与后果');
 }else if(id==='P37'){
  t('policy-checks','服务地区　　　　工作材料　　　　公众服务义务',64,205,1152,68,29,C.muted);
  t('filing-big','1112 款',64,316,536,105,64,C.teal,true);t('filing-scope','截至 2026-08-31\n完成备案的生成式 AI 服务',64,437,536,102,27);
  t('policy-date','2026-11-12',690,316,526,105,49,C.navy,true);t('policy-scope','Anthropic 新使用政策\n拟生效日期',690,437,526,102,27);
  takeaway(sl,'服务备案、应用登记与模型能力，不属于同一指标','公告日期为 2026-10-08，新政策在 10 月稿中尚未生效。');
 }else if(id==='P38'){
  columns([['一个任务','从小材料开始'],['一份模板','目标、材料、边界、交付'],['一套验收','原文、关键数值、实际文件']],{y:197,size:42,subSize:25});
  t('appendix-navigation','会后查阅',64,584,250,42,25,C.orange,true);t('appendix-links','参数　价格　工具　配置指南　资源与证据',64,637,1152,36,25,C.teal,true);
  h.links.push({slide:h.serial(),shape:'appendix-links',target:39});
 }else if(id==='A03'){
  table(sl,[['型号','上下文 token','最大输出 token','输入模态'],['Opus 5.5','1M','128K','文字、图像'],['GPT-6.1 Sol','1,050K','128K','文字、图像'],['Gemini 4 Argon','1M 共享','最多 1M，扣除输入','文字、图像、音频、视频'],['GLM-5.3','1M','128K','文字'],['Qwen3.8-Max','1,000K','131,072','按具体接口版本'],['Kimi K3','1M','上下文余量，默认 128K','按具体接口版本'],['MiMo-V2.6-Pro','1M','未确认','文字、图像、音频、视频'],['DeepSeek V4.1 Flash','1M','384K','文字、图像']],[330,190,320,312],{y:201,font:21,heights:[44,...Array(8).fill(44)]});
  takeaway(sl,'入口限制、工具支持、官方链接与单位说明放在备注中','产品记忆不属于上下文规格。“未确认”不代表不支持。');
 }else if(id==='A04'){
  table(sl,[['产品','美元月付参考','核对重点'],['ChatGPT','Go 8　Plus 20　Pro 100 / 200 / 500','地区、分档、历史额度'],['Claude','Pro 20　Max 100 起','用量上限与指定 API 赠额'],['Google AI','Plus 4.99　Pro 19.99　Ultra 99.99 起','地区、档位与活动'],['X Subscription','Plus 8　Premium 30　Super 100　Ultra 200','旧订阅与第三方模型费用'],['DeepSeek / 千问','有免费使用入口','API 与团队功能另核'],['Kimi','当期月费未确认','以实际官方账户页面为准']],[240,575,337],{y:204,font:21,heights:[45,...Array(6).fill(55)]});
  takeaway(sl,'购买前看续费价、限额、取消入口与付款对象','查询于 2026-10-18。美元为美国网页月付参考，不含税费，具体地区另核。');
 }else if(id==='A06'){
  listRows(sl,[['设备','系统版本、空间与要求'],['安装','官方入口与版本'],['账号','登录、验证与恢复'],['连通','消息、公开文件上传'],['故障','网络、地区、验证、额度分开查']],{y:205,step:75,labelW:230,font:29});
  takeaway(sl,'演示设备与客户端：待定');
 }else if(id==='A07'){
  columns([['购买前','地区、权益、续费'],['购买中','金额、币种、付款对象'],['购买后','生效、账单、取消入口']],{y:208,size:42,subSize:26});
  takeaway(sl,'演示产品与套餐：待定');
 }else if(id==='A11'){
  table(sl,[['项目','先解决什么'],['obra/superpowers','开发步骤与验证行为'],['mattpocock/skills','按需组合少量任务方法'],['Fission-AI/OpenSpec','需求与实现脱节'],['open-gsd/gsd-core','长期任务与跨会话进度'],['BMAD-METHOD','较完整的角色协作流程'],['oh-my-claudecode','已明确的并行子任务'],['mindfold-ai/Trellis','项目规范与交接上下文']],[560,592],{y:204,font:23,heights:[44,...Array(7).fill(48)]});
  takeaway(sl,'先解决一个痛点，用最少流程验证适用性');
 }else if(id==='A12'){
  table(sl,[['内容','阅读时必须保留'],['幻觉、迎合与越狱数字','具体型号、设置、任务和时间'],['EchoLeak CVSS','评分方与版本分别注明'],['GitHub MCP 演示','不可信资料与已有权限的组合'],['密钥“70% 仍有效”','2025 报告，跟踪 2022 年泄露群体'],['加拿大航空案','BC 民事裁决庭，未核实精确赔付构成'],['Terminal-Bench','榜单版本、Agent 与推理设置']],[420,732],{y:211,font:23,heights:[45,...Array(6).fill(56)]});
  takeaway(sl,'缺少研究条件时，不引用裸百分比');
 }else if(id==='A13'){
  columns([['鹈鹕','同提示词\n一次生成、无历史记忆'],['图纸','同图纸、十个配置\n三个事后检查点'],['报告','四个配置\n旧模板、清单与参考函']],{y:210,size:46,subSize:26});
  takeaway(sl,'模型、客户端、设置、原始输出和评价规则都需要留档','图纸按命中数排序。报告采用汇总表口径，分项 112 元差异单独说明。');
 }
 return true;
}
