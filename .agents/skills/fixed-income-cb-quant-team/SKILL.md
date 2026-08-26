---
name: fixed-income-cb-quant-team
description: Orchestrate a research-only quantitative team for fixed income and convertible bonds, with role separation across literature and domain research, point-in-time data governance, strategy design, statistical review, research engineering, independent replication, and report production. Use this skill whenever the user asks to build, invoke, or coordinate a quant research team; develop or audit a fixed-income or convertible-bond factor, signal, backtest, portfolio, dataset, methodology, replication, or academic report; or diagnose data leakage, overfitting, survivorship bias, multiple testing, or irreproducible results in those domains. Default to Chinese bond and exchange-listed convertible-bond markets unless another market is stated. Do not use for live trading, order routing, brokerage or account operations, intraday monitoring, investment-advice personalization, or simple copyediting unrelated to quantitative research.
---

# 固收与可转债量化研究团队

把主 Agent 作为唯一编排者，按任务需要调用项目中的专业 Agent。目标是形成可审计、可复现、能被反证的学术研究成果，而不是产生可直接下单的交易指令。

## 开始前读取

每次启动完整研究或审查流程时：

1. 读取 `references/team-charter.md`，确定角色边界和成果所有权。
2. 读取 `references/workflow-gates.md`，选择阶段、交接方式和验收门槛。
3. 涉及数据、因子、回测或市场制度时，读取 `references/domain-checklists.md`。
4. 需要写研究任务书、交接单、复现报告或正式报告时，从 `assets/` 复制相应模板并按任务裁剪。

简单概念问答不必加载所有参考文件，也不必调用子 Agent。

## 确立研究任务书

先从用户输入和项目文件中提取以下信息，只对会实质改变研究设计的缺口提问：

- 研究问题与拟检验机制；
- 资产范围：利率债、信用债、可转债或其子样本；
- 市场、样本期、频率、调仓周期和预测期限；
- 可用数据、数据来源及每个字段的实际可得时间；
- 基准、约束、成本和评价指标；
- 预期交付物：方案、代码、数据审计、回测、复现或报告；
- 已知限制、截止时间和不得修改的路径。

用户没有指定时，采用保守默认值并在任务书中显式列出；不要悄悄补齐关键假设。

## 选择最小角色集合

项目已注册以下自定义 Agent：

- `fiq_research_lead`：研究任务书、阶段决策和最终综合判断。
- `fiq_domain_researcher`：文献、制度和经济机制证据。
- `fiq_data_steward`：时点数据、清洗、口径和数据冻结。
- `fiq_strategy_researcher`：假设、因子、信号、组合与实验设计。
- `fiq_methodology_reviewer`：事前统计方案和事后方法审查。
- `fiq_research_engineer`：数据管道、回测实现、测试与可复现工程。
- `fiq_replication_reviewer`：冻结后的独立复现和反证。
- `fiq_research_editor`：通过验证后的报告与知识归档。

按任务选择：

| 任务 | 必选角色 | 视情况增加 |
|---|---|---|
| 新策略完整研究 | lead、domain、data、strategy、methodology、engineer、replication、editor | 无 |
| 数据质量或未来函数审查 | data、methodology | domain、replication |
| 既有策略复现 | replication、methodology | data、engineer |
| 因子构思与研究设计 | domain、strategy、methodology | data |
| 回测框架或研究代码建设 | engineer、data | methodology、replication |
| 已验证成果写报告 | editor | lead、domain、methodology |
| 快速概念说明 | 主 Agent 直接完成 | 通常不委派 |

不要为了展示团队规模而启动无关角色。

## 编排原则

1. 主 Agent 保留任务分解、角色选择、冲突裁决和最终答复权。专业 Agent 不再创建子 Agent。
2. 能独立开展的文献、数据可行性和事前方法设计可以并行；存在依赖的步骤必须按门槛顺序进行。
3. 每项成果设唯一负责人和明确路径。多个 Agent 不得同时修改同一文件。
4. 策略研究员提出定义，研究工程师实现；任何口径变化都要回交策略研究员确认。
5. 方法论审查员在开发前后各介入一次。策略开发者不得批准自己的统计结论。
6. 数据、代码、配置和依赖环境冻结后再启动独立复现。复现者只写新的复现输出，不修改冻结来源。
7. 编辑只整理已经通过门槛的证据；未经验证的发现必须标为探索性结果。

## 标准阶段与门槛

按 `references/workflow-gates.md` 执行以下阶段：

1. **G0 立项**：研究问题可证伪，范围和输出明确。
2. **G1 证据与数据可行性**：制度机制有证据，关键字段能按时点获得。
3. **G2 预分析批准**：主要假设、指标、样本划分和稳健性检验在看结果前确定。
4. **G3 实现一致性**：代码与策略定义一致，测试覆盖关键边界。
5. **G4 统计与经济审查**：控制过拟合、多重检验和现实约束，区分统计显著与经济意义。
6. **G5 冻结与复现**：冻结材料完整，独立运行得到可解释的一致结果。
7. **G6 报告发布**：结论逐项链接证据、版本、限制和负面结果。

任何门槛未通过时，输出“未通过原因—所需修正—责任角色—重新验收条件”，不要越级包装成最终结论。

## 固收与可转债研究底线

- 所有特征按真实可得时间构造；披露日不等于数据库入库日时采用更保守时间。
- 证券池必须包含历史到期、退市、赎回、转股或失去可交易资格的标的，避免幸存者偏差。
- 正确处理交易日历、停牌、涨跌停、上市初期、摘牌、付息、除权、转股价调整、强赎、回售和下修等事件。
- 同时报告毛收益和扣除保守成本后的研究收益；成本是假设，不是实盘承诺。
- 明确价格口径、收益口径、久期/凸性/信用暴露以及可转债股债性暴露。
- 参数搜索、因子筛选和样本切分必须留下完整实验账本；探索性发现不得冒充预先设定的检验。
- 没有数据或无法运行时，不虚构数字、图表、显著性或引用。

## 专业 Agent 的统一交接格式

要求每个专业 Agent 返回：

1. `任务与边界`
2. `使用的输入和版本`
3. `方法与关键假设`
4. `发现或完成项`
5. `证据与产物路径`
6. `风险、限制和未决问题`
7. `门槛建议：通过 / 有条件通过 / 不通过`
8. `下一接收角色`

使用 `assets/handoff-template.md` 保存正式交接记录。

## 最终答复要求

主 Agent 应当：

- 先说明研究问题是否得到支持，而不是先堆叠过程描述；
- 区分确认事实、实验结果、合理推断和未验证假设；
- 给出实际运行的样本、数据版本、代码/配置版本和关键口径；
- 同时报告有效结果、失败结果、敏感性和适用边界；
- 不把历史回测写成未来收益保证，不输出个性化买卖指令；
- 链接项目中的真实交付文件，并列出仍需用户决定的事项。

## 用户调用示例

- “使用 `$fixed-income-cb-quant-team`，研究可转债双低因子的样本外稳定性，先做预分析，不要直接写策略代码。”
- “调用固收量化团队审查这个回测是否有未来函数和幸存者偏差，只输出审查报告。”
- “让独立复现角色在冻结版本上重跑该策略，不允许修改原始代码。”
- “用完整团队完成从文献、数据、回测、复现到研究报告的流程。”

