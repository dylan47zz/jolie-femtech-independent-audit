# 三个文件独立终审核查报告

**项目**：Jolie FemTech 英文会议论文及其数据、复现脚本  
**审计日期**：2026年9月7日  
**审计对象**：

1. `① 论文终稿_Jolie FemTech.docx`
2. `② 附件二_数据表_Jolie FemTech.xlsx`
3. `③ 复现脚本_reproduce_models.py`

**审计方式**：只读独立核查。未改动三个原始文件；没有采信工作簿中的 PASS 标记、已录入模型结果或相互一致性说明作为正确性证据。

---

## 一、总体结论与放行判断

### 1.1 总体结论

**结论：暂不可交付。影响等级：B类。**

B类的含义是：核心数据/模型并非整体失效，但可复现证据链、工作簿可用性和提交完整性存在重大缺口，必须大修后再向学生交付为“可提交版本”。

| 审计维度 | 结论 | 独立核查依据 |
|---|---|---|
| WDI 原始数据可信度 | **通过** | 2026年9月7日实际重新访问三个 World Bank API；MMR、GDP、CPR 三份响应的 SHA-256 均与工作簿登记完全一致；逐国“最新有效值（2010–2023）”比对为 MMR 167/167、GDP 167/167、CPR 有值国 139/139 匹配。 |
| 政治体制原始数据可信度 | **未能完全核实** | OWID 归档网页、V-Dem v16 版本、四分类标签及来源元数据可访问；但工作簿没有交付登记的 `1210053.data.json` 与 `metadata.json` 原始文件，且从归档页面露出的运行时 JSON 地址直接请求时返回 404，故无法重算两份 JSON 的 SHA-256 或逐国回配。 |
| Excel 内部计算 | **修正少量问题后可用** | 逐国底稿派生变量、描述统计、缺失统计和所有模型表经独立复算正确；但“相关系数”表 D9:D11 的**保存缓存值**与其公式语义不符。 |
| Python 模型复现 | **修正少量问题后可用** | 脚本在当前环境成功运行；其已打印的系数、SE、p、n、df、R²/调整R²与独立复算一致。但它不是官方源到模型的全流程复现，且未输出 t 值、置信区间和逐国样本清单。 |
| Word 数字准确性 | **修正少量问题后可用** | Word 的核心样本量、描述统计、模型点估计、SE、p、HC3 CI、缺失率及五项正文相关系数均与独立复算一致；但其依赖的 Excel 相关系数缓存错误，且仍存在身份占位符。 |
| 参考文献与隐私证据 | **基本可信但需补证** | 多数官方/期刊/公司页面和数字已实际访问核对；BMJ 正文、NIH PDF、WHO PDF 的直接正文访问受 403/非PDF响应限制，按“未能完全核实”处理，未推定正确。 |
| 版面/文件完整性 | **不具备最终提交状态** | Word 8页渲染无图表裁切、未见修订痕迹；但第1页保留作者、机构、邮箱占位符。Excel 无隐藏表/行/列、外部链接或批注，但默认导出为 63 页且无冻结窗格、打印区域，审计阅读体验不合格。 |

### 1.2 不应作出的统计解释

本审计**不**把以下表述视为同义：

- `p ≥ .05` ≠ “没有关联”或“证明零效应”；
- 跨规格不稳健 ≠ 结论必然相反；
- 估计不精确 ≠ 效果不存在；
- HC3 处理异方差 ≠ HC3 修复选择性缺失。

就当前底稿而言：M1/M2 的体制项均不精确；M3/M4 的体制项在 HC3 下仍为正且 `p<.05`；M5 在限制到 2015–2023 后仍为正但 HC3 CI 跨 0。最准确的结论是：**CPR 的主模型正向关联在当前底稿中存在，但对时间覆盖和体制编码敏感；MMR 的调整后关联小且不精确，不能被表述为“已证明不存在关联”。**

---

## 二、高、中、低严重程度问题清单

### 2.1 高严重程度问题

| 编号 | 文件与定位 | 原内容 | 独立核查结果与证据来源 | 问题原因 | 可直接采用的修改文本/操作 |
|---|---|---|---|---|---|
| H-01 | `② 附件二_数据表_Jolie FemTech.xlsx`，`相关系数!D9:D11` | D9 公式为 `CORREL(logGDP,MMR)`，保存值却是 `0.5752048369`；D10 保存值为 `0.4886844960`；D11 保存值为 `0.4457185935`。 | 从“跨国底稿”原始列独立复算：D9 应为 **-0.7159911511**（n=167）；D10 应为 **0.5752048370**（n=139）；D11 应为 **0.4886844960**（n=166）。对原工作簿用 LibreOffice 强制重算后，结果恢复为该正确顺序。公式文本本身正确，错误在保存缓存。 | 工作簿未在最后一次公式重算后正确保存，导致依赖缓存值的程序或审稿人读取到错误关联。 | **操作**：打开工作簿，执行“全工作簿重新计算”，保存并关闭；再以公式值模式复核 D9:D11。**应显示**：D9 `-0.7159911511`；D10 `0.5752048370`；D11 `0.4886844960`。 |
| H-02 | `② 附件二_数据表_Jolie FemTech.xlsx`，`数据源!A8:J8`；`①` 第2页 Methods（政治体制来源段）与第7页 Ref[3] | 只登记 `1210053.data.json + 1210053.metadata.json` 的文件名、哈希及归档展示页URL。 | OWID 归档页实际可访问，支持 V-Dem v16、归档日 2026-08-05、四分类和来源更新时间 2026-03-17；但未随附件提供原始 JSON。页面内运行时 JSON 路径直接请求返回 404，两个登记 SHA-256 无法重算，167国体制值无法由原始文件逐行复匹配。 | 关键暴露变量没有随交付包保存原始快照和可执行转换过程；展示页不是充分的数据归档。 | **在数据包新增**：`raw/owid/1210053.data.json`、`raw/owid/1210053.metadata.json`、`scripts/build_regime_2023.py`、`raw/owid/SHA256SUMS.txt`。**Methods 可替换为**：`The archived OWID source files, their SHA-256 hashes, and the country-code transformation script are supplied with the replication package. The script selects the 2023 observation, maps source entities to ISO3, and documents the separate-treatment rule for Palestine.` |
| H-03 | `① 论文终稿_Jolie FemTech.docx`，第1页标题区 | `Jolie [Lastname]`；`[Institution Name]`；`[student email]`。 | 渲染第1页及 DOCX 文本提取均确认三项占位符未替换。 | 最终稿未完成作者元数据。 | **替换模板**：`Jolie <Last name>`；`<Institution name>`；`Corresponding author: <student email>`。替换后重新导出 PDF，核查首版页面。 |
| H-04 | `③ 复现脚本_reproduce_models.py`，第11–16行；`①` 第7页 Ref[2] 的“cannot be reproduced from live API”表述 | 脚本自动选择同目录首个 `*数据表*.xlsx`；Ref[2]称该 WDI 快照不能从 live API 重现。 | 脚本在空目录失败；多个匹配文件时输入选择不确定。另一方面，本次实际重新请求 WDI 三个 URL，响应 SHA-256 与工作簿哈希完全一致，说明“不能从 live API 重现”在审计日并不成立。 | 输入文件定位不确定；文字把“未来可能变化”误写成“当前不可复现”。 | **脚本操作**：改为 `python reproduce_models.py --input "② 附件二_数据表_Jolie FemTech.xlsx"`。**Ref[2可替换句**：`The API response retrieved on 3 September 2026 is identified by filename and SHA-256. At the audit date, the live endpoint responses still matched those hashes; future source updates may change the response, so the raw snapshots should be retained in the replication package.` |

### 2.2 中严重程度问题

| 编号 | 文件与定位 | 原内容 | 独立核查结果与证据来源 | 问题原因 | 可直接采用的修改文本/操作 |
|---|---|---|---|---|---|
| M-01 | `③ 复现脚本_reproduce_models.py`，`ols()`函数及文件头声明 | 注释称可“完整复现三表全部数值”；脚本输出仅 β、两类SE、两类p、n/df/R²/调整R²。 | 脚本实际运行成功，已输出字段与独立复算一致；但不打印经典/HC3 t 值、经典/HC3 CI，也不写机器可读结果。因此不能单凭其输出验证“全部数值”。 | 文档承诺超过脚本可观察输出；缺少审计产物。 | **替换声明**：`This script reproduces the model estimates from the workbook-level country dataset. It writes coefficients, standard errors, t statistics, p values, confidence intervals, sample identifiers, and fit statistics to a machine-readable output file.` |
| M-02 | `③` 第1–3行、整个工作流；`② 分析复现说明!A2:A8` | 将脚本描述为“完整复现”。 | 脚本不下载 WDI/OWID、不验证哈希、不生成167国样本、不记录样本合并规则；只能做“整理后Excel→模型”。 | 未区分来源复现、数据整理复现、模型复现。 | **新增三步说明**：`Stage 1 downloads and hashes source files; Stage 2 constructs the 167-country analytic file and records exclusions; Stage 3 estimates M1–M7 from that analytic file.` |
| M-03 | `①` 第3页 Methods 缺失说明；`② 缺失分析!A1:F8` | Word正确说 HC3 不处理选择，但未给出缺失机制边界。 | CPR 缺失为 28/167；按体制 0–3 分别 11.1%、7.8%、18.2%、33.3%。缺失集中足以反驳“默认完全随机”的表述，但不能凭此证明特定 MNAR 机制。 | 缺少“不能从分组差异推出缺失机制”的方法边界。 | **Methods 追加**：`The observed variation in missingness by regime category is evidence against assuming MCAR by default, but it does not by itself identify a missingness mechanism. HC3 changes heteroskedasticity-robust uncertainty only; it does not correct selection into complete cases.` |
| M-04 | `①` 第6页 Discussion（`r=.489`句）；`③` 的潜在VIF解释 | 讨论称体制和收入“moderately related”，未直接称严重共线性。 | 独立 Pearson `r=.488684`；标准两预测变量 VIF = `1/(1-r²)=1.314`。不支持严重多重共线性。 | 如果后续以脚本中无截距矩阵 VIF=4.17 解释调整变化，会误用诊断。 | **可采用文本**：`Regime category and log GDP were moderately correlated (r=.489; two-predictor VIF=1.31), which does not indicate severe multicollinearity. The attenuation after adjustment should not be attributed to collinearity alone.` |
| M-05 | `①` 第1–8页全文 | 论文自述约2800词的背景下，正文实际约3615词。 | 从DOCX文字提取的英文 token 计数约 3,615；会议官方规则未随附件提供，不能断言其最终计数口径。 | 无投稿字数核查记录。 | **操作**：按会议系统规则重新计数；如限 2,800 词，删除约800词，并在提交前保存计数截图/导出记录。 |
| M-06 | `②` 全工作簿打印与审计视图 | 无冻结窗格、无打印区；默认PDF导出63页，许多页几乎为空。 | 实际渲染确认；工作簿无隐藏行列、外链、批注和错误值，但审计表阅读不可用。 | 宽表没有为打印/阅读设置布局。 | **操作**：冻结`跨国底稿`第4行和A:B列；将逐国表拆为“原始列”和“派生列”两张审计视图；设横向打印、重复标题行、合理打印区域。 |

### 2.3 低严重程度问题

| 编号 | 文件与定位 | 原内容 | 独立核查结果/原因 | 可直接采用的修改文本或操作 |
|---|---|---|---|---|
| L-01 | `② 跨国底稿!R5:R171` | `=IF($Lr="","",$Lr-2018)` 的28个缺失CPR年份行在缓存中为空。 | 这是IF返回空字符串的预期结果，不是公式错误；但 data-only 读者可能误判为未计算。 | 列头改为：`CPR year minus 2018 (blank when CPR year missing)`。 |
| L-02 | `①` 第7–8页参考文献 | 长URL占行、阅读负担大。 | 渲染无截断，但可读性一般。 | 使用悬挂缩进；优先 DOI、稳定官方链接或存档链接。 |
| L-03 | `② 一致性检查` | `PASS` 标记可能被误读为外部核验。 | 该页仅核对预设行数、范围和少量公式错误，不能验证来源真实性。 | A2改为：`PASS indicates only the listed internal workbook checks. It is not evidence of source authenticity or independent statistical validation.` |

---

## 三、数据来源真实性核查表

**核验方法**：没有采信 Excel 的“数据源”说明作为结论；对 WDI 三个完整API URL进行了实际访问、分页检查、响应哈希重算和逐国匹配；对 OWID 归档页实际访问并核对元数据。无法获得原始文件的项目标记为“未能完全核实”。

| 数据集 | 发布机构 | 工作簿版本/下载日 | 完整URL | 年份规则/单位 | 独立核查结果 | SHA-256 结果 | 结论 |
|---|---|---|---|---|---|---|---|
| Maternal mortality ratio, `SH.STA.MMRT` | World Bank, World Development Indicators | API last updated 2026-07-13；下载 2026-09-03 | `https://api.worldbank.org/v2/country/all/indicator/SH.STA.MMRT?format=json&per_page=20000` | 2010–2023最新非缺失；每10万活产的孕产妇死亡数 | 实际响应 `page=1, pages=1, per_page=20000`；167/167国家值和年份匹配 | 重算为 `d19e18021c680f5f910b13c9d9f93d4219824cc5b0f44350e1b69e8840b07c73`，与登记一致 | **通过** |
| GDP per capita, PPP, constant 2021 international $, `NY.GDP.PCAP.PP.KD` | World Bank, WDI | 同上 | `https://api.worldbank.org/v2/country/all/indicator/NY.GDP.PCAP.PP.KD?format=json&per_page=20000` | 2010–2023最新非缺失；2021不变价国际元/人 | 167/167国家值和年份匹配；`logGDP=LN(GDP)` 167/167正确 | 重算为 `c610e48829239b4b7fbdae26883093fa70f0c2e4cff6845d314beec722c7f6a5`，一致 | **通过** |
| Contraceptive prevalence, any method, `SP.DYN.CONU.ZS` | World Bank, WDI | 同上 | `https://api.worldbank.org/v2/country/all/indicator/SP.DYN.CONU.ZS?format=json&per_page=20000` | 2010–2023最新非缺失；15–49岁已婚女性任一避孕方法使用比例 | 139个有值国家值/年份匹配；28个空值国家在实时快照中同样无2010–2023有效值 | 重算为 `260237f909d0f09b84eae91b3746403ffaec82199b23c02a0669e9ff28cc76e5`，一致 | **通过** |
| Political regime / Regimes of the World | V-Dem v16，经 Our World in Data 处理 | OWID archive 2026-08-05 14:35 UTC；下载 2026-09-03 | `https://archive.ourworldindata.org/20260805-143540/grapher/political-regime.html` | 固定2023；0 closed autocracy、1 electoral autocracy、2 electoral democracy、3 liberal democracy | 归档页、V-Dem v16、四分类、来源更新时间2026-03-17可验证；PSE在底稿中保留但无统一体制值，逻辑自洽 | 登记的 data/metadata JSON 文件未交付；从归档页推导的运行时JSON URL本次返回404，无法复算 `03b999...`/`5b86dd...` | **部分验证，原始文件未能核实** |

### 3.1 167国逐国底稿复核摘要

| 检查项 | 独立结果 | 结论 |
|---|---|---|
| ISO3/国家 | 167行、167个唯一ISO3；未发现重复ISO3 | 通过 |
| MMR/GDP/CPR 来源匹配 | WDI 如上逐国比对全部匹配（CPR有值139国） | 通过 |
| MMR/GDP/CPR 年份 | 所有采用值均为2010–2023窗口内的最新可用观察 | 通过 |
| logGDP、log(MMR) | 与 LN(原始GDP)、LN(MMR) 分别 167/167一致 | 通过 |
| 体制代码及文字类别 | 0=27、1=51、2=55、3=33、PSE空值=1；文字类别一致 | 通过，但原始OWID文件未能重取 |
| CPR 原始有效样本 | 139；完整案例并入体制后 138；2015–2023限制样本109 | 通过 |
| CPR 缺失标记 | 28个 `Missing` 与原始CPR空值一一对应 | 通过 |
| 合理范围 | MMR 1–993；GDP 1033.009–130296.844；CPR 6.900–85.475；CPR年份2010–2023 | 通过 |

---

## 四、独立数据、缺失与相关系数复核表

### 4.1 描述统计

| 变量 | n | 均值 | 样本SD | 中位数 | Q1 | Q3 | IQR | 最小值 | 最大值 | Word/Excel 比较 |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|
| MMR | 167 | 116.940 | 168.263 | 41.0 | 9.5 | 155.0 | 145.5 | 1 | 993 | Word第3页所报均值116.9、中位数41、范围1–993一致；Excel描述统计一致 |
| log(MMR) | 167 | 3.637 | 1.676 | 3.714 | 2.250 | 5.043 | 2.794 | 0 | 6.901 | Excel一致；Word未完整报告 |
| GDP PPP | 167 | 25423.047 | 26036.227 | 15916.958 | 5627.125 | 39691.698 | 34064.574 | 1033.009 | 130296.844 | Excel一致；Word未完整报告 |
| logGDP | 167 | 9.571 | 1.170 | 9.675 | 8.635 | 10.589 | 1.954 | 6.940 | 11.778 | Excel一致 |
| 政治体制代码 | 166 | 1.566 | 0.987 | 2 | 1 | 2 | 1 | 0 | 3 | Excel一致；Word样本数一致 |
| CPR | 139 | 49.896 | 20.768 | 54.6 | 31.045 | 66.076 | 35.031 | 6.9 | 85.475 | Word第3页所报中位数54.6、范围6.9–85.5一致；Excel一致 |
| CPR年份 | 139 | 2017.741 | 3.564 | 2018 | 2015 | 2021 | 6 | 2010 | 2023 | Word第3页所报中位数2018、IQR=6一致；Excel一致 |

### 4.2 CPR缺失复核

| 政治体制类别 | 总数 | CPR观测 | CPR缺失 | 缺失率 | Word/Excel比较 |
|---|---:|---:|---:|---:|---|
| 0 Closed autocracy | 27 | 24 | 3 | 11.1% | 一致 |
| 1 Electoral autocracy | 51 | 47 | 4 | 7.8% | 一致 |
| 2 Electoral democracy | 55 | 45 | 10 | 18.2% | 一致 |
| 3 Liberal democracy | 33 | 22 | 11 | 33.3% | 一致 |
| 合计 | 166（有体制代码） | 138 | 28 | — | 另有 PSE：CPR有值、体制缺失；Word和Excel逻辑一致 |

28个 CPR 缺失国家为：BHR、BGR、CAN、HRV、CYP、CZE、DNK、GRC、HUN、ISL、ISR、KWT、LVA、LBN、LTU、LUX、MLT、NZL、NOR、ROU、SYC、SGP、SVK、SVN、SWE、SYR、ARE、URY。该分布足以使“默认MCAR”站不住脚，但不能单凭分组频率确定MNAR机制。

### 4.3 Pearson 相关系数：Word—Excel—独立重算

| 变量X | 变量Y | 配对n | Word报告 | Excel当前保存值 | 独立重算 | 判定 |
|---|---|---:|---:|---:|---:|---|
| 政治体制代码 | log(MMR) | 166 | -0.441 | -0.441397 | -0.441397 | 三者一致 |
| 政治体制代码 | MMR | 166 | 未报告 | -0.367959 | -0.367959 | Excel正确 |
| 政治体制代码 | CPR | 138 | 0.446 | 0.445719 | 0.445719 | 三者一致 |
| logGDP | log(MMR) | 167 | -0.850 | -0.849733 | -0.849733 | 三者一致 |
| logGDP | MMR | 167 | 未报告 | **0.575205（错误缓存）** | **-0.715991** | Excel不一致 |
| logGDP | CPR | 139 | 0.575 | **0.488684（错误缓存）** | 0.575205 | Word与独立重算一致；Excel不一致 |
| 政治体制代码 | logGDP | 166 | 0.489 | **0.445719（错误缓存）** | 0.488684 | Word与独立重算一致；Excel不一致 |

---

## 五、数据与模型复核表

### 5.1 独立复算方法

- 不使用 Excel 中现成模型结果；直接读取“跨国底稿”的 MMR、GDP、政治体制、CPR、CPR年份列。
- OLS：`β=(X'X)^−1X'y`；经典协方差：`s²(X'X)^−1`，`s²=e'e/(n−k)`。
- HC3：`(X'X)^−1 X' diag[eᵢ²/(1−hᵢᵢ)²] X (X'X)^−1`。
- 双侧 p 值与95%CI均使用残差自由度的 t 分布，未用正态近似；极小 p 以 `<0.0001` 表示，未写作 p=0。
- 对 Excel `模型输出与HC3结果` 的 26个系数行、每行19个可比较字段，共 **495个数值** 作比较；全部在浮点容差内一致。

### 5.2 全部系数的独立模型复算结果

| 模型 | 项 | n | k | df | β | 经典SE | 经典p | 经典95%CI | HC3 SE | HC3 p | HC3 95%CI | R² | 调整R² |
|---|---|---:|---:|---:|---:|---:|---:|---|---:|---:|---|---:|---:|
| M1 log(MMR)~有序体制+logGDP | 截距 | 166 | 3 | 163 | 15.250 | 0.592 | <0.0001 | [14.080,16.420] | 0.529 | <0.0001 | [14.206,16.294] | 0.731 | 0.728 |
| M1 log(MMR)~有序体制+logGDP | 政治体制代码（每升1类） | 166 | 3 | 163 | -0.053 | 0.079 | 0.5038 | [-0.209,0.103] | 0.075 | 0.4812 | [-0.202,0.095] | 0.731 | 0.728 |
| M1 log(MMR)~有序体制+logGDP | logGDP | 166 | 3 | 163 | -1.203 | 0.067 | <0.0001 | [-1.335,-1.072] | 0.061 | <0.0001 | [-1.324,-1.083] | 0.731 | 0.728 |
| M2 MMR原始尺度~有序体制+logGDP | 截距 | 166 | 3 | 163 | 1103.088 | 79.426 | <0.0001 | [946.252,1259.924] | 98.594 | <0.0001 | [908.401,1297.775] | 0.520 | 0.514 |
| M2 MMR原始尺度~有序体制+logGDP | 政治体制代码（每升1类） | 166 | 3 | 163 | -3.487 | 10.619 | 0.7430 | [-24.455,17.481] | 9.245 | 0.7065 | [-21.744,14.769] | 0.520 | 0.514 |
| M2 MMR原始尺度~有序体制+logGDP | logGDP | 166 | 3 | 163 | -102.336 | 8.947 | <0.0001 | [-120.003,-84.670] | 10.158 | <0.0001 | [-122.395,-82.278] | 0.520 | 0.514 |
| M3 CPR~有序体制+logGDP | 截距 | 138 | 3 | 135 | -39.685 | 12.580 | 0.0020 | [-64.564,-14.806] | 12.834 | 0.0024 | [-65.066,-14.304] | 0.373 | 0.363 |
| M3 CPR~有序体制+logGDP | 政治体制代码（每升1类） | 138 | 3 | 135 | 4.827 | 1.677 | 0.0047 | [1.510,8.144] | 1.982 | 0.0162 | [0.908,8.747] | 0.373 | 0.363 |
| M3 CPR~有序体制+logGDP | logGDP | 138 | 3 | 135 | 8.820 | 1.442 | <0.0001 | [5.969,11.672] | 1.584 | <0.0001 | [5.687,11.954] | 0.373 | 0.363 |
| M4 CPR~有序体制+logGDP+年份中心化 | 截距 | 138 | 4 | 134 | -38.760 | 13.202 | 0.0039 | [-64.871,-12.650] | 13.515 | 0.0048 | [-65.491,-12.029] | 0.373 | 0.359 |
| M4 CPR~有序体制+logGDP+年份中心化 | 政治体制代码（每升1类） | 138 | 4 | 134 | 4.800 | 1.687 | 0.0051 | [1.464,8.137] | 2.030 | 0.0195 | [0.786,8.815] | 0.373 | 0.359 |
| M4 CPR~有序体制+logGDP+年份中心化 | logGDP | 138 | 4 | 134 | 8.723 | 1.503 | <0.0001 | [5.749,11.696] | 1.631 | <0.0001 | [5.496,11.949] | 0.373 | 0.359 |
| M4 CPR~有序体制+logGDP+年份中心化 | CPR年份−2018 | 138 | 4 | 134 | -0.102 | 0.424 | 0.8112 | [-0.941,0.738] | 0.469 | 0.8287 | [-1.028,0.825] | 0.373 | 0.359 |
| M5 CPR（2015–2023）~有序体制+logGDP | 截距 | 109 | 3 | 106 | -44.284 | 14.497 | 0.0029 | [-73.027,-15.542] | 15.186 | 0.0043 | [-74.392,-14.176] | 0.342 | 0.330 |
| M5 CPR（2015–2023）~有序体制+logGDP | 政治体制代码（每升1类） | 109 | 3 | 106 | 3.265 | 2.015 | 0.1081 | [-0.730,7.261] | 2.362 | 0.1697 | [-1.417,7.947] | 0.342 | 0.330 |
| M5 CPR（2015–2023）~有序体制+logGDP | logGDP | 109 | 3 | 106 | 9.549 | 1.698 | <0.0001 | [6.183,12.915] | 1.892 | <0.0001 | [5.797,13.301] | 0.342 | 0.330 |
| M6 log(MMR)~体制虚拟变量+logGDP | 截距 | 166 | 5 | 161 | 15.056 | 0.666 | <0.0001 | [13.741,16.372] | 0.622 | <0.0001 | [13.828,16.285] | 0.732 | 0.726 |
| M6 log(MMR)~体制虚拟变量+logGDP | 体制1：选举型威权 vs 0类 | 166 | 5 | 161 | 0.103 | 0.210 | 0.6252 | [-0.311,0.516] | 0.214 | 0.6317 | [-0.319,0.524] | 0.732 | 0.726 |
| M6 log(MMR)~体制虚拟变量+logGDP | 体制2：选举民主 vs 0类 | 166 | 5 | 161 | -0.046 | 0.209 | 0.8279 | [-0.459,0.368] | 0.204 | 0.8232 | [-0.447,0.356] | 0.732 | 0.726 |
| M6 log(MMR)~体制虚拟变量+logGDP | 体制3：自由民主 vs 0类 | 166 | 5 | 161 | -0.119 | 0.257 | 0.6442 | [-0.627,0.389] | 0.241 | 0.6219 | [-0.595,0.357] | 0.732 | 0.726 |
| M6 log(MMR)~体制虚拟变量+logGDP | logGDP | 166 | 5 | 161 | -1.191 | 0.071 | <0.0001 | [-1.330,-1.052] | 0.067 | <0.0001 | [-1.323,-1.059] | 0.732 | 0.726 |
| M7 CPR~体制虚拟变量+logGDP | 截距 | 138 | 5 | 133 | -40.912 | 14.436 | 0.0053 | [-69.466,-12.359] | 14.729 | 0.0063 | [-70.046,-11.779] | 0.376 | 0.357 |
| M7 CPR~体制虚拟变量+logGDP | 体制1：选举型威权 vs 0类 | 138 | 5 | 133 | 3.390 | 4.195 | 0.4204 | [-4.907,11.688] | 4.943 | 0.4939 | [-6.386,13.166] | 0.376 | 0.357 |
| M7 CPR~体制虚拟变量+logGDP | 体制2：选举民主 vs 0类 | 138 | 5 | 133 | 10.613 | 4.255 | 0.0138 | [2.198,19.029] | 5.125 | 0.0403 | [0.477,20.749] | 0.376 | 0.357 |
| M7 CPR~体制虚拟变量+logGDP | 体制3：自由民主 vs 0类 | 138 | 5 | 133 | 12.477 | 5.639 | 0.0286 | [1.323,23.632] | 6.383 | 0.0527 | [-0.147,25.102] | 0.376 | 0.357 |
| M7 CPR~体制虚拟变量+logGDP | logGDP | 138 | 5 | 133 | 9.005 | 1.560 | <0.0001 | [5.920,12.090] | 1.697 | <0.0001 | [5.649,12.361] | 0.376 | 0.357 |

### 5.3 Word—Excel—Python模型一致性检查

| 项目 | Word位置/内容 | Excel位置 | Python脚本实际输出 | 独立复算 | 结论 |
|---|---|---|---|---|---|
| M1 主模型 | Word第4页：β=-.053、经典SE=.079、经典p=.504、HC3SE=.075、HC3p=.481、HC3CI[-.202,.095]、R²=.731/调整=.728 | `模型结果!A5:M15`；完整表 `模型输出与HC3结果` | 输出匹配上述 β、SE、p、n=166、df=163、R²/调整R² | 完全匹配 | 一致 |
| M2 原始MMR | Word第4页：β=-3.49、HC3p=.707 | `稳健性分析` 与完整模型表 | 输出β=-3.4874、HC3p=.7065 | 完全匹配 | 一致 |
| M3 CPR主模型 | Word第4页：β=4.83、经典SE=1.68、p=.0047、HC3SE=1.98、p=.016、CI[.91,8.75] | `模型结果!A6:M15` | 输出匹配 | 完全匹配 | 一致 |
| M4 CPR+年份 | Word第4页：β=4.80、HC3CI[.79,8.82]、p=.019 | `稳健性分析` 与完整模型表 | 输出匹配 | 完全匹配 | 一致 |
| M5 CPR限制年份 | Word第5页：β=3.27、HC3CI[-1.42,7.95]、p=.170 | `稳健性分析` 与完整模型表 | 输出匹配 | 完全匹配 | 一致 |
| M6 MMR类别模型 | Word第4页称三个对比均不显著，但未列数值 | `模型输出与HC3结果!A21:T25` | 输出5项系数 | 独立复算所有系数一致 | 一致；Word为概括性表述 |
| M7 CPR类别模型 | Word第5页：体制2 β=10.61、HC3p=.040；体制3 β=12.48、HC3p=.053 | `模型输出与HC3结果!A26:T30` | 输出匹配 | 完全匹配 | 一致 |
| 经典/HC3 t与CI | Word没有逐项完整列出 | Excel完整表列J:R存在 | 脚本**不打印**t和CI | 独立复算与Excel一致 | Excel正确；脚本交付不完整 |

---

## 六、Word—Excel—脚本一致性检查结论

### 6.1 已核对且一致的项目

1. 167国总体、MMR样本167、体制样本166、CPR原始有效139、CPR完整案例138、2015–2023 CPR样本109。
2. MMR、GDP、CPR的2010–2023“最新有效值”规则及WDI数据值/年份。
3. `logGDP`、`log(MMR)`、三类体制虚拟变量、CPR年份中心化变量的逻辑和数值。
4. M1–M7 所有回归系数、经典SE、经典p、HC3 SE、HC3 p、两类CI、R²、调整R²（Excel模型表与独立复算一致）。
5. Word中M1–M5的主报告模型结果、M7的两个类别对比、CPR缺失率和主要描述统计。
6. 隐私矩阵的18/25、19/24、Flo/Premom案件类型和“不可合并为单一比例”的边界。

### 6.2 已发现的不一致

1. Excel `相关系数!D9:D11` 保存值错误，详见 H-01；**Word三项核心相关系数本身正确**。
2. Python脚本的注释称“完整复现三表全部数值”，但实际输出不含 t 和置信区间，且不从官方源开始重建数据。
3. Excel数据源说明称快照不能从实时API复现；本次实时请求与SHA完全匹配，该句在审计日不成立。
4. Excel虽有内部 PASS 检查，但未捕捉到相关系数缓存错误；不能作为独立验证结论。

---

## 七、参考文献逐条核对表

| # | 作者/机构及准确题名 | 日期/期刊或机构/标识符 | 正文支持关系 | 独立核查结论 |
|---:|---|---|---|---|
| 1 | World Health Organization. *Trends in maternal mortality 2000 to 2020: estimates by WHO, UNICEF, UNFPA, World Bank Group and UNDESA/Population Division.* | WHO, Geneva, 2023；所列URL：`https://iris.who.int/bitstream/handle/10665/366225/9789240068759-eng.pdf` | 支持2000年339、2020年223的全球MMR背景数字。 | **基本正确但未能全文核实**：URL本次返回HTML而非可读PDF；题名/机构/年份可核。 |
| 2 | World Bank. *World Development Indicators.* 指标 `SH.STA.MMRT`、`SP.DYN.CONU.ZS`、`NY.GDP.PCAP.PP.KD`。 | WDI；API last updated 2026-07-13；三个完整URL见第三节。 | 支持国家级MMR、CPR、GDP及单位/年份。 | **完全正确**：真实端点、分页、值、年份与哈希均独立验证。 |
| 3 | V-Dem (2026) — processed by Our World in Data. *Democracy* [dataset]. V-Dem, *Democracy report v16* [original data]. | OWID archived 2026-08-05：`https://archive.ourworldindata.org/20260805-143540/grapher/political-regime.html` | 支持四类Regimes of the World及版本。 | **基本正确但需补原始文件**：网页元数据正确；原始JSON和哈希未能重算。 |
| 4 | Mozilla Foundation. *In post-Roe v. Wade era, Mozilla labels 18 of 25 popular period and pregnancy tracking tech with Privacy Not Included warning.* | 2022-08-17；`https://www.mozillafoundation.org/en/blog/in-post-roe-v-wade-era-mozilla-labels-18-of-25-popular-period-and-pregnancy-tracking-tech-with-privacy-not-included-warning/` | 支持18/25、10个经期app+10个孕期app+5个设备、Mozilla warning终点。 | **完全正确**。证据类型应为独立民间评审，不是观测到的数据传输比例。 |
| 5 | Grundy Q, Chiu K, Held F, Continella A, Bero L, Holz R. *Data sharing practices of medicines related apps and the mobile ecosystem: traffic, content, and network analysis.* | *BMJ*. 2019;364:l920. DOI `10.1136/bmj.l920`；`https://www.bmj.com/content/364/bmj.l920` | 支持24个medicine-related Android apps及19/24的网络传输结果。 | **基本正确但未能全文核实**：书目信息和DOI可验证；BMJ自动访问返回403，19/24需提交可访问PDF/全文页码。不能称其为FemTech专样本。 |
| 6 | US Federal Trade Commission. *FTC Finalizes Order with Flo Health, a Fertility-Tracking App that Shared Sensitive Health Data with Facebook, Google, and Others.* | 2021-06-22；`https://www.ftc.gov/news-events/news/press-releases/2021/06/ftc-finalizes-order-flo-health-fertility-tracking-app-shared-sensitive-health-data-facebook-google` | 支持FTC关于Flo分享敏感数据的指控、同意要求及独立审查要求。 | **完全正确**。正文应继续使用 alleged/complaint/FTC action，不应写成无条件事实认定。 |
| 7 | US Federal Trade Commission. *Ovulation Tracking App Premom Will be Barred from Sharing Health Data for Advertising Under Proposed FTC Order.* | 2023-05-17；`https://www.ftc.gov/news-events/news/press-releases/2023/05/ovulation-tracking-app-premom-will-be-barred-sharing-health-data-advertising-under-proposed-ftc` | 支持 proposed order、100,000美元民事罚金及广告分享限制。 | **完全正确**。必须保留 `proposed` 状态，不能改写为最终执行事实。 |
| 8 | Flo Health. *Flo's response to the FTC settlement.* | 更新 2024-10-01；`https://flo.health/newsroom/flo-response-ftc-settlement-update` | 支持公司称2022年3月已完成审计。 | **基本正确但需规范**：只属于公司自报，不能作为独立审计机构的验证。 |
| 9 | Grand View Research. *FemTech Market Size, Share, & Trends Analysis Report ... 2025–2030.* | Research and Markets页面；150 pages，November 2024；`https://www.researchandmarkets.com/reports/5595800/femtech-market-size-share-and-trends-analysis` | 支持2024年39.29bn、2030年97.25bn、CAGR约16.3%。 | **完全正确**，但仅为商业市场预测，不是同行评议公共卫生证据。 |
| 10 | *Funding research on women's health* [editorial]. | *Nature Reviews Bioengineering*. 2024;2:797–798. 2024-10-11. DOI `10.1038/s44222-024-00253-7`。 | 支持约5% R&D、2021年FemTech约2.5bn/约3%。 | **完全正确**；文章类型为editorial。 |
| 11 | *Design with women in mind* [editorial]. | *Nature Biomedical Engineering*. 2026;10(7):1265–1266. 2026-07-17. DOI `10.1038/s41551-026-01758-9`。 | 支持女性健康研究/工程的背景讨论。 | **完全正确**；文章类型为editorial。 |
| 12 | NIH Office of Research on Women's Health. *Perspectives on Advancing NIH Research to Inform and Improve the Health of Women: Executive Summary.* | 报告列为 updated 2022-03-01；`https://orwh.od.nih.gov/sites/orwh/files/docs/ORWH_WHC_ExecutiveSummary508.pdf` | 支持10.8%及$4,466m。 | **未能核实**：本次PDF端点返回403。提交前须提供可访问官方副本、存档链接或具体页码。 |
| 13 | World Health Organization. *Endometriosis.* | Fact sheet，2025-10-15；`https://www.who.int/news-room/fact-sheets/detail/endometriosis` | 支持约10% / 190 million。 | **完全正确**。 |
| 14 | World Health Organization. *Polycystic ovary syndrome.* | Fact sheet；`https://www.who.int/news-room/fact-sheets/detail/polycystic-ovary-syndrome` | 支持10–13% reproductive-aged women。 | **完全正确**：数值被页面支持；最终引用日期应按页面元数据再核。 |
| 15 | Flo Health. *Flo Health Research Reveals Menopause Makes Menstrual Cycles Longer and Less Predictable in Largest Study of its Kind.* | 2024-06-19；`https://flo.health/newsroom/flo-health-menstrual-cycle-patterns-study` | 支持380m downloads和67m MAUs。 | **完全正确**，但为公司自报。 |
| 16 | Flo Health. *Flo Health Secures More than $200M Investment from General Atlantic to Revolutionize Women's Health; First Purely Digital Consumer Women's Health App to Achieve Unicorn Status.* | 2024-07-30；`https://flo.health/newsroom/flo-health-raises-over-200m` | 支持“as of June 2024, nearly70m MAUs”。 | **完全正确**，但为公司自报。 |

---

## 八、最终修改清单

### 8.1 提交前必须修改

1. **重算并保存** Excel `相关系数!D9:D11`，确保正确缓存值为 `-0.7159911511`、`0.5752048370`、`0.4886844960`。
2. **补齐政治体制原始数据包**：两份原始JSON、SHA-256、下载时间/响应头、筛选/映射/编码脚本、PSE处理说明及逐国比对日志。
3. **修复脚本可复现性**：显式输入文件参数、依赖文件、Python版本、运行命令、输出CI/t/样本清单、机器可读结果；在论文和工作簿中把“完整复现”限定为真实能力。
4. **替换 Word 身份占位符**，并重新渲染全部页面。
5. **核对投稿字数**。若会议规则约2800词，按官方规则压缩当前约3615词正文。
6. **对Ref[12]提供可访问官方证据**；对Ref[5]提供全文/页码或可访问存档，以便核实19/24。

### 8.2 建议修改

1. 在 Methods 中加入缺失机制边界和“HC3不处理选择性缺失”的精确表述。
2. 在 Discussion 中如需讨论共线性，使用 `r=.489; VIF=1.31`，不要使用无截距计算出的4.17。
3. 将Excel设置为审计可读版本：冻结窗格、打印区、横向分页、公式列/手工列区分、长URL显示策略。
4. 为论文参考文献使用悬挂缩进和稳定URL/DOI，提升最终PDF可读性。
5. 将“PASS”页的标题和说明改为“内部工作簿检查”，避免被误认为外部数据验证。

### 8.3 已经核验通过，无须修改

1. WDI 三个数据端点的代码、单位、分页范围、2026-07-13更新时间、2026-09-03响应哈希及逐国数值/年份。
2. 167国样本、PSE体制缺失处理的**当前工作簿逻辑**、MMR/GDP/CPR完整案例计数。
3. 对数变换、三类虚拟变量、CPR年份中心化公式。
4. M1–M7 的 Excel 现存模型数字和 Python 脚本已打印结果。
5. Word中的主要回归数字、描述统计、CPR缺失分组结果、五项正文相关系数。
6. 隐私矩阵不将18/25、19/24和FTC个案混为一个比例，且证据类别边界基本合格。
7. Word图表、表格和页码渲染：未发现裁切、异常跨页、批注或修订残留。

---

## 九、最终放行意见

### 当前是否可以发给学生？

**不能作为“可提交的最终文件包”发给学生。**

如果目的是将本终审核查报告发给学生据此修改，可以发；但三个原始文件本身不应被标记为最终可交版本。

### 阻止交付的最少必要修改

满足以下五项后，可改判为“修正少量问题后可交付”：

1. 修复并保存 Excel `相关系数!D9:D11` 缓存；
2. 交付可重取/可校验的政治体制原始文件和转换脚本；
3. 修复 Python 输入定位、环境说明和输出完整性；
4. 替换 Word 首页身份占位符；
5. 核对并满足会议字数要求，同时补齐 NIH/BMJ 的可访问核验证据。

在上述修改完成前，不应把文件中的内部 PASS、模型表与脚本能够运行，描述为“数据来源、数据整理和模型均已完全独立复现”。

---

## 十、审计证据与可复查产物

本报告的分析产物保存在与原文件同一目录的 `audit_artifacts/` 中：

- `live_source_verification.txt`：三份WDI实时响应、SHA-256及逐国匹配结果；
- `independent_audit.txt` 和 `independent_models.json`：独立描述统计、相关、缺失和M1–M7模型复算；
- `模型完整独立复算表.csv`：全部26个系数项的可审计模型表；
- `supplied_script_output.txt`：供应Python脚本实际运行输出；
- `privacy_source_verify.txt`、`reference_source_evidence.txt`：隐私来源与文献网页的访问取证摘要；
- `docx_render_20260907/`：原Word八页渲染检查；
- `xlsx_pdf_20260907/`：原Excel默认PDF导出检查。

**证据边界声明**：本审计证明“当前已交付的逐国底稿 → M1–M7 模型”的数值可独立复算；在政治体制原始JSON未交付且无法从归档运行时端点重新取得的条件下，不能证明“所有官方原始文件 → 国家底稿 → 模型”的全流程已完全复现。
