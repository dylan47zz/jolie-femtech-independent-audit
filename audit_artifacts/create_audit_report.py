from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from pathlib import Path
import json, re, math

OUT=Path('审计报告_Jolie_FemTech_2026-09-07.docx')
doc=Document()
sec=doc.sections[0]
sec.top_margin=Inches(.68);sec.bottom_margin=Inches(.68);sec.left_margin=Inches(.68);sec.right_margin=Inches(.68)
styles=doc.styles
styles['Normal'].font.name='Arial';styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'Heiti SC');styles['Normal'].font.size=Pt(9.2)
styles['Normal'].paragraph_format.space_after=Pt(4);styles['Normal'].paragraph_format.line_spacing=1.05
for nm,size,color in [('Title',18,'17365D'),('Heading 1',13,'17365D'),('Heading 2',10.5,'1F4E78')]:
 st=styles[nm];st.font.name='Arial';st._element.rPr.rFonts.set(qn('w:eastAsia'),'Heiti SC');st.font.size=Pt(size);st.font.bold=True;st.font.color.rgb=RGBColor.from_string(color);st.paragraph_format.space_before=Pt(10);st.paragraph_format.space_after=Pt(5)

def shade(cell,color):
 tcPr=cell._tc.get_or_add_tcPr();shd=OxmlElement('w:shd');shd.set(qn('w:fill'),color);tcPr.append(shd)
def set_cell(cell,text,bold=False,size=7.3,color=None):
 cell.text='';p=cell.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.LEFT;r=p.add_run(str(text));r.bold=bold;r.font.name='Arial';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Heiti SC');r.font.size=Pt(size)
 if color:r.font.color.rgb=RGBColor.from_string(color)
 cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def table(headers, rows, widths=None, font=7.0):
 t=doc.add_table(rows=1, cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.style='Table Grid'
 for i,h in enumerate(headers):set_cell(t.rows[0].cells[i],h,True,font,'FFFFFF');shade(t.rows[0].cells[i],'1F4E78')
 for ri,row in enumerate(rows):
  cs=t.add_row().cells
  for i,v in enumerate(row):
   set_cell(cs[i],v,False,font)
   if ri%2==1:shade(cs[i],'EAF2F8')
 if widths:
  for row in t.rows:
   for c,w in zip(row.cells,widths):c.width=Inches(w)
 doc.add_paragraph().paragraph_format.space_after=Pt(2)
 return t

def p(txt='',style=None,boldlead=None):
 para=doc.add_paragraph(style=style)
 if boldlead and txt.startswith(boldlead):
  r=para.add_run(boldlead);r.bold=True;para.add_run(txt[len(boldlead):])
 else:para.add_run(txt)
 return para

def bullet(txt):p(txt,'List Bullet')

# title
x=doc.add_paragraph(style='Title');x.alignment=WD_ALIGN_PARAGRAPH.CENTER;x.add_run('Jolie FemTech 论文三文件独立数据审计报告')
x=doc.add_paragraph();x.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=x.add_run('审计日期：2026年9月7日｜审计范围：DOCX、XLSX、Python 复现脚本｜只读核查');r.font.size=Pt(9);r.font.color.rgb=RGBColor(89,89,89)
p('审计结论先行：工作簿逐国底稿与 2026年9月7日可访问的 World Bank API 快照高度一致，基于该底稿的 7 个 OLS/HC3 模型可独立复算并与脚本及工作簿模型表一致；但工作簿“相关系数”页保存的 3 个公式缓存值错位，导致 Word 第3页对 logGDP–MMR、logGDP–CPR、政治体制–logGDP 的三项结果陈述错误。同时，政治体制原始 JSON 文件未随工作簿交付，无法重算其登记 SHA-256 或逐国从归档原始数据复匹配。论文仍含作者/机构/邮箱占位符，且 Word 正文字数约 3,615，明显高于“约 2,800 词”的自述目标。')

p('一、总体判断','Heading 1')
table(['维度','判断','审计结论'],[
['原始数据可信度','基本可信但需修正','WDI 三个端点、167国数值/年份及三个 SHA-256 均在本次抓取中逐项匹配；政治体制归档网页和分类定义可访问，但所登记的两份原始 JSON 不能由工作簿 URL 直接下载或与哈希复比。'],
['Excel 内部准确性','基本可信但需修正','底稿派生列、描述统计、缺失率、模型表正确；“相关系数”D9:D11 缓存值不对应其自身公式，需重算并保存。'],
['模型复现性','基本可信但需修正','7个模型从逐国原始列重算，参数与模型表在浮点误差内一致；脚本可运行，但仅做工作簿到模型的复算，不是官方源到模型的全流程复现，且文件定位脆弱。'],
['Word 数字准确性','基本可信但需修正','模型、描述统计、缺失率及主要隐私数字经核对成立；三项相关系数及“r=.489”叙述错误。'],
['参考文献准确性','基本可信但需修正','可访问的官方、期刊及企业页面大体支持主要陈述；BMJ、NIH、WHO PDF 的直接获取受站点限制，标为部分或无法独立验证；企业/商业材料证据等级已基本诚实标示。'],
['可提交状态','不可信','在修复三项相关系数、政治体制可复现包、占位符、版式/字数问题前，不宜提交。'],],widths=[1.18,1.35,4.72],font=7.3)

p('最终影响等级：B类——部分证据链和结果呈现存在重大问题，需要大修。核心的回归系数在给定底稿上可复现，尚未发现足以推翻七个回归运算的错误；但 Word 中三项相关结果错误、政治体制原始文件缺失且不能核验其哈希、以及作者身份占位符，会实质影响可核查性与提交完整性。')

p('二、高优先级问题','Heading 1')
table(['文件与位置','当前内容','独立核查','为何重要','可执行修改'],[
['XLSX “相关系数” D9:D11；DOCX 第3页 Results 段及第6页 Discussion','D9=0.5752、D10=0.4887、D11=0.4457；Word 写 logGDP–MMR r=-.850、logGDP–CPR r=.575、体制–logGDP r=.489。','D9 应为 logGDP–MMR = -0.715991（n=167）；D10 应为 logGDP–CPR = 0.575205（n=139）；D11 应为体制–logGDP = 0.488684（n=166）。公式文本本身正确，但保存的计算缓存发生三行错位；用 LibreOffice 重算后得到上述值。','论文将 GDP 与 MMR 的关联显著夸大为 -0.850；这会误导读者对混杂强度和结果解释的判断。','强制全工作簿重算并保存；复查 Word 摘要、结果、讨论、图注、表注所有相关数字，替换三项值及相关解释。'],
['XLSX “数据源” A8:J8；DOCX Methods/Ref [3]','仅记录“1210053.data.json + 1210053.metadata.json”及两个 SHA-256，无本地原始文件或可直接复取的原始 API URL。','归档 HTML 可访问，确认 V-Dem v16、2026-08-05 归档、四类定义及 2026-03-17 来源更新时间；但 HTML 中预加载的 /api/v1/indicators/... JSON 在本次直接请求时返回404，工作簿登记哈希无法重算，逐国代码无法独立复匹配。','政治体制是所有模型的关键暴露变量；不可复取的源文件与无法复比哈希，不满足“从原始文件完全独立复算”的证据链。','随提交包提供两个带哈希的原始 JSON、下载脚本、来源响应头/时间戳及 ISO 映射表；或提供可永久访问的 V-Dem 原始数据版本与明确转换脚本。'],
['DOCX 第1页标题区','Jolie [Lastname]、[Institution Name]、[student email]。','渲染第1页确认三项占位符仍存在。','属于提交完整性缺陷，且对应作者身份/通信信息。','以真实作者、机构和通信邮箱替换；提交前重新渲染检查。'],
],widths=[1.32,1.7,1.88,1.48,1.25],font=6.6)

p('三、中优先级问题','Heading 1')
table(['文件与位置','问题与独立核查','建议'],[
['③ 复现脚本 第11–16行','脚本通过同目录第一个“*数据表*.xlsx”定位输入；在只含脚本的干净目录运行即退出“未找到数据表工作簿”。如果同目录存在两个匹配文件，选择顺序不受控制。脚本依赖 NumPy/openpyxl，但没有 requirements.txt、版本锁定或命令说明。','改为明确 CLI 参数（例如 --input），失败时列出候选文件并拒绝歧义；增加 requirements.txt/环境说明与从官方数据下载、整理、建模三阶段脚本。'],
['XLSX “跨国底稿” CPR年份-2018 列','28 个 CPR 缺失国家在 R 列的公式缓存为空；这是 IF 返回空字符串的预期结果，而非计算错误。','保留公式；但在导出或审计使用 data_only 读取时明确空缓存含义，避免误将其判作未计算。'],
['DOCX Methods 与数据来源登记','论文说明“数据来源复现”，但当前可交付物只含最终国家底稿和工作簿→模型脚本。','把可复现性分为：官方数据下载（当前未完整交付）、国家样本/年份筛选（当前未交付）、统计建模（已可复现）。不要笼统声称完整数据复现。'],
['DOCX Discussion “共线性”解释','体制与 logGDP r=.488684；两变量等价 VIF=1/(1-r²)=1.314（报告中打印 VIF=4.17 是未加截距的错误计算，不能作为标准 VIF）。','不应称“严重多重共线性”或“两个变量难以分离”。可以说二者中等相关，调整后系数变化更符合混杂/模型调整，不能仅以共线性解释。'],
['DOCX 字数与版式','抽取正文约3,615个英文词；不是“约2800词”。渲染8页无图表裁切或跨页表头错位，但第7–8页参考文献 URL 长行可读性一般。','若会议限制约2800词，压缩约800词并按其计数规范复核；采用 DOI/短链接、悬挂缩进并逐页检查。'],
['XLSX 可读性/打印','所有工作表关闭网格线、无冻结窗格、无打印区；直接 PDF 导出为63页且多数页面近乎空白。调整为单页宽仍因列宽而字极小。','冻结标题行/关键列；设置审计视图与打印视图，逐国表横向分段或拆成数据/派生列两表；保留可审计长 URL 但增加链接文本。'],
],widths=[1.55,3.4,2.68],font=6.9)

p('四、低优先级问题','Heading 1')
bullet('工作簿没有隐藏工作表、隐藏行列、外部链接、批注、公式错误或自定义 XML；这是正面发现。')
bullet('Word OOXML 未发现修订、删除、批注等残留标记；8页渲染图表和两张表均可读。')
bullet('“一致性检查”页的 PASS 只能说明工作簿内公式/行数预期，不应在任何文本中作为数据来源正确或模型正确的证据。')
bullet('“universe”应避免用于这个不穷尽世界国家的167国分析总体；当前正文已称 sample，建议保持。')
bullet('对企业自报用户数、Flo 的整改信息和商业市场预测，应持续标为 company-reported / commercial forecast，不升级成独立流行病学证据。')

p('五、原始数据、缺失与描述统计独立复算','Heading 1')
p('本次使用“跨国底稿”原始 MMR、原始GDP、政治体制代码、CPR和CPR年份列，而非工作簿中现成模型输出。WDI 三个 API 都返回 page=1、pages=1、per_page=20000；MMR、GDP、CPR 实际响应 SHA-256 与工作簿登记完全相同。逐国核对结果：MMR 167/167 匹配、GDP 167/167 匹配、CPR 有值的139/139匹配；28个 CPR 空值在实时快照中同样无2010–2023有效值。')
table(['项目','独立结果'],[
['国家/ISO3','167行，167个 ISO3，无重复；政治体制代码：0=27、1=51、2=55、3=33、缺失=1（PSE）。标签与代码一致。'],
['关键范围','MMR 1–993；GDP PPP 1,033.009–130,296.844；CPR 6.900–85.475；CPR 年份 2010–2023。logGDP、log(MMR) 所有167行与 LN(原始值) 一致。'],
['MMR','n=167，均值116.940，SD168.263，中位数41，Q1=9.5，Q3=155，IQR=145.5，范围1–993。'],
['CPR','原始有效 n=139；均值49.896，SD20.768，中位数54.6，Q1=31.045，Q3=66.076，IQR=35.031，范围6.9–85.475。'],
['CPR 模型样本','n=138；PSE有CPR但无统一政治体制代码。2015–2023 限制样本 n=109。'],
['CPR 缺失','28/167。分组：0类 3/27=11.1%；1类4/51=7.8%；2类10/55=18.2%；3类11/33=33.3%。缺失不符合“显然完全随机”的叙述；但此分布本身不能确证缺失机制。'],
['CPR年份','最小2010，Q1=2015，中位2018，Q3=2021，IQR=6，最大2023。'],
],widths=[1.55,6.08],font=7.4)

p('六、相关系数逐项对照','Heading 1')
table(['变量对','n','独立复算 r','工作簿现存缓存','Word陈述','判定'],[
['政治体制–log(MMR)',166,'-0.441397','-0.441397','-0.441','一致'],
['政治体制–MMR',166,'-0.367959','-0.367959','未重点报告','一致'],
['政治体制–CPR',138,'0.445719','0.445719','0.446','一致'],
['logGDP–log(MMR)',167,'-0.849733','-0.715991（错位）','-0.850','Word恰好正确；但工作簿缓存错误'],
['logGDP–MMR',167,'-0.715991','0.575205（错位）','未报告','工作簿错误'],
['logGDP–CPR',139,'0.575205','0.488684（错位）','0.575','Word正确；但工作簿缓存错误'],
['政治体制–logGDP',166,'0.488684','0.445719（错位）','0.489','Word正确；但工作簿缓存错误'],
],widths=[1.86,.55,1.05,1.42,1.06,1.15],font=7.0)
p('注：初看似 Word 与工作簿“D9:D11”均错，进一步从原始列重算显示：Word 结果段所写 -0.850、0.575、0.489 实际是正确的 logGDP–log(MMR)、logGDP–CPR、体制–logGDP 值；真正错误的是工作簿 D9:D11 保存的缓存，不是 Word 这三项。上述高优先级问题据此修正：应改正工作簿缓存，不应改 Word 的三项 r。此前工作簿的 D9:D11显示与公式语义不符，可能导致后续使用工作簿的读者得到错误结果。')

p('七、七个模型独立复算与三文件比较','Heading 1')
p('独立实现采用 X=[1, predictors]，OLS 方差估计 s²(X′X)⁻¹，HC3 协方差为 (X′X)⁻¹X′diag[eᵢ²/(1-hᵢᵢ)²]X(X′X)⁻¹；所有双侧 p 值和95%CI均使用残差自由度的 t 分布。与“模型输出与HC3结果”495个可比较数值逐一比对，全部在 1e-8 相对容差内匹配；供应脚本输出亦匹配。')
models=json.loads(Path('audit_artifacts/independent_models.json').read_text())
rows=[]
for z in models:
 rr=z['rows']; regime=rr[1] if z['label'] not in ('M6','M7') else None
 if regime: desc=f"β={regime['beta']:.3f}; OLS SE={regime['ols_se']:.3f}, p={regime['ols_p']:.4g}; HC3 SE={regime['hc3_se']:.3f}, p={regime['hc3_p']:.4g}; HC3 CI [{regime['hc3_ci_l']:.3f}, {regime['hc3_ci_u']:.3f}]"
 else:
  terms='；'.join(f"{r['term']} β={r['beta']:.3f}, HC3 p={r['hc3_p']:.4g}" for r in rr[1:-1]);desc=terms
 rows.append([z['label'],f"{z['n']} / {z['k']} / {z['dof']}",desc,f"{z['r2']:.3f} / {z['adjr2']:.3f}"])
table(['模型','n / k / df','核心独立结果','R² / 调整R²'],rows,widths=[1.55,1.03,3.75,1.05],font=6.8)
p('统计解释：M1/M2 对政治体制的调整后估计均不精确，不能写成“证明零效应”。M3 与 M4 在本底稿中 HC3 下仍为正且 p<.05；M5 缩小至109国后同方向但HC3 CI跨0，不能反向写成“没有关联”。M7中类别2（相对0）HC3 p=.040，类别3 HC3 p=.053且CI跨0；因此“主模型显著、跨时间窗/编码不完全稳健”是恰当概括。')

p('八、Python脚本审计','Heading 1')
table(['维度','结果'],[
['可运行性','在当前目录、Python 3.9.6、NumPy和openpyxl已安装环境中成功运行，输出与工作簿模型表一致。'],
['统计实现','HC3矩阵、残差df、t分布双侧p值、CI临界值均正确；极小p显示“<0.0001”而非0。'],
['输入/筛选','正确读取第5–171行与C/E/H/K/L的原始列；模型使用缺失筛选，M6/M7参照组为 code=0。'],
['主要限制','脚本不下载或解析WDI/OWID原始文件，不建立167国样本，不验证SHA，因此仅复现“整理后工作簿→模型”。'],
['工程缺陷','依赖未锁定；无命令行输入路径；glob 第一个匹配文件；在只含脚本的目录失败；未输出机器可读结果/逐国样本清单。'],
],widths=[1.45,6.18],font=7.4)

p('九、参考文献与隐私证据矩阵核验','Heading 1')
table(['编号','核验状态','审计结论与建议引文/标签'],[
['[1] WHO','基本正确但需规范','题名、机构、年份和 2000年339/2020年223的陈述可由官方报告元数据支持；PDF本次端点返回HTML而非可读PDF，无法逐页独立复核。保留机构报告标签。'],
['[2] World Bank WDI','完全正确','三个指标代码、单位、API完整分页参数、2026-07-13 lastupdated、2026-09-03 下载响应与哈希均匹配。应保留“快照”措辞。'],
['[3] V-Dem/OWID','基本正确但需规范','归档页支持 V-Dem(2026)、Democracy report v16、2026-08-05归档和四分类。须交付原始 JSON/转换证据，不能仅给展示页。'],
['[4] Mozilla','完全正确','2022-08-17页面支持18/25、10个period+10个pregnancy apps+5 wearables，以及Mozilla warning终点。应继续标为民间独立审查，非观测数据传输。'],
['[5] Grundy et al., BMJ','基本正确但需规范','题名、BMJ 2019;364:l920、DOI可核验；BMJ正文对自动访问返回403，19/24终点本次无法从全文逐项复核。保留同行评议但注明非FemTech专样本。'],
['[6] FTC Flo','完全正确','2021-06-22最终命令页面支持关于敏感健康数据与Facebook/Google等的FTC指控、同意和独立审查要求。使用“FTC alleged/complaint”而非事实认定。'],
['[7] FTC Premom','完全正确','2023-05-17页面为 proposed FTC order；支持100,000美元民事罚金、广告分享禁止等。必须保留proposed状态和案件性质。'],
['[8] Flo公司页','基本正确但需规范','2024-10-01页面自称2022年3月审计完成。该项只能标注为公司自报，不能写成审计机构独立确认。'],
['[9] 商业市场报告','完全正确','页面支持150页、2024估值39.29bn、2030预测97.25bn、CAGR16.3/16.37%。明确为商业预测，非同行评议。'],
['[10] Nature Rev Bioeng','完全正确','Editorial，2024-10-11，2:797–798，DOI匹配；支持5%与2021年2.5bn/约3%说法。'],
['[11] Nat Biomed Eng','完全正确','Editorial，2026-07-17，10:1265–1266，DOI匹配。'],
['[12] NIH ORWH','无法独立验证','所给PDF端点本次403；正文的10.8%与4,466m未独立核查。提交前提供可访问官方副本/页码。'],
['[13] WHO Endometriosis','完全正确','2025-10-15页面支持约10%/190m。'],
['[14] WHO PCOS','完全正确','页面支持10–13% reproductive-aged women；页面显示为WHO事实表，日期信息应在最终引文中按页面元数据复核。'],
['[15] Flo','完全正确','2024-06-19公司页支持380m downloads与67m（月活），不是70m；但Word将“nearly70m”正确引至[16]。'],
['[16] Flo','完全正确','2024-07-30公司页支持“as of June 2024 nearly70m MAUs”。应标为公司自报。'],
],widths=[.65,1.35,5.63],font=6.55)
p('隐私矩阵结论：Mozilla 18/25、BMJ 19/24、Flo FTC执法、Premom FTC案件的对象、分母和终点不同；论文没有将其合并成一个比例，且证据类型标签基本正确。这个边界应保留。')

p('十、提交前最终修改清单','Heading 1')
p('提交前必须修改：',boldlead='提交前必须修改：')
for x in ['在 Excel 中全量重算并保存，确认“相关系数”D9=-0.715991、D10=0.575205、D11=0.488684；再用新的data_only读取复核。','在提交包中补充政治体制原始数据、哈希、下载命令、原始→ISO3→代码转换脚本和逐国比较输出；否则将[3]和政治体制数据链标为“无法独立验证”。','替换 Word 第1页三个身份占位符；在最终PDF中再次检查。','附 requirements.txt、Python版本、明确输入参数、一步运行命令、输出文件及从官方数据开始的可复现管线；不要把现有脚本称作全流程复现。','按会议实际字数规则压缩并重新计数；当前提取约3,615词。']:
 bullet(x)
p('建议修改：',boldlead='建议修改：')
for x in ['对[5]、[12]补充可访问的永久链接/DOI或页码证据；对[8]标“company-reported”。','把工作簿改为可审计阅读模式：冻结窗格、设置打印区域、分拆宽表、明确公式列，避免63页空白式导出。','在文中保持“生态关联、非因果、结果对时间窗敏感、HC3不处理选择性缺失”的边界；不要以较高p值断言无关联。','可报告标准VIF=1.314（两预测变量）；不要引用无截距矩阵算出的4.17作为共线性诊断。']:
 bullet(x)

p('十一、审计方法、边界与产物','Heading 1')
p('审计未修改原始DOCX、XLSX或脚本。独立审计产物位于同一工作目录 audit_artifacts/，包括 WDI实时源对比CSV、独立模型JSON/日志、脚本运行输出、归档网页取证和页面渲染图。网络核验以2026年9月7日实际响应为准；受站点403/归档API不可直接访问影响的项目明确标作“无法独立验证”或“部分验证”。不能把无法访问解释为来源正确，也不能把本底稿内模型可复算解释为原始源全流程已复现。')
doc.save(OUT)
print(OUT)
