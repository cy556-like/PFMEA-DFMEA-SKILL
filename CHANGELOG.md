# 变更记录

## [2.0.0] - 2026-06-27

### 新增

- 首次发布 PFMEA/DFMEA Skill v2.0.0
- 基于 AIAG & VDA FMEA 手册（2019 版）七步法
- 支持 DFMEA（设计 FMEA）与 PFMEA（过程 FMEA）两类分析
- 内置 S/O/D 评分表（10 级）+ AP 行动优先级矩阵（H/M/L）
- 5 套行业模板：电子电气、机械装配、表面处理、涂装、通用兜底
- 3 个 references 文件：七步法指南、SOD 评分表、失效链案例库
- FMEA 报告生成器（generate_fmea.py）：输出 xlsx（7 Sheet）+ docx
- 自动填充模式（auto_fill=True）：基于失效描述自动判定 S/O/D 评分
- 风险矩阵热力图：S×O 与 S×D 矩阵嵌入 xlsx
- 特殊特性识别：S=9-10 自动标识为 CC，S=8 且 AP=H/M 自动标识为 SC

### 关键决策

- 采用 2019 版 AP（Action Priority）替代旧版 RPN（Risk Priority Number）
- DFMEA 与 PFMEA 共用 AP 矩阵（手册明确两者一致）
- 暂不覆盖 FMEA-MSR（监视及系统响应 FMEA），后续版本规划
- 模板设计参考 8d-skill 的模板架构，保持一致的工程范式

### 已知限制

- 不支持 FMEA-MSR（监视及系统响应 FMEA）
- 不支持多语言（仅中文输出）
- 风险矩阵热力图依赖 openpyxl，未安装时跳过
- 失效链案例库目前覆盖 5 个行业，后续将扩展

---

## 规划中

### [2.1.0] - 计划

- 增加 FMEA-MSR 支持
- 增加控制计划（CP）自动生成
- 增加 DRBFM（设计评审基于失效模式）模式
- 增加多语言支持（英文/日文）

### [2.2.0] - 计划

- 增加 FMEA 与 APQP/PPAP 文档的自动关联
- 增加基于历史 FMEA 数据的智能推荐
- 增加客户特定要求（CSR）库
