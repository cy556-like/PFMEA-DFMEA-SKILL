# 变更记录

## [2.0.1] - 2026-06-28

### 重构：脚本工程化升级（对标 8D Skill 的工程化深度）

**问题背景**：v2.0.0 的 `generate_fmea.py`（899 行）在脚本工程化深度上弱于 8D Skill（1991 行）。具体差距：
1. 缺少独立的样式工具函数，样式逻辑散落在主函数内
2. 智能填充（auto_fill）只覆盖 S/O/D 评分，未覆盖责任人/日期/编制人/审核人等
3. 行高使用粗略估算（`len(text)/2`），中文长文本仍会被截断
4. 缺少 SOD 一致性校验（ap_hint 与 s/o/d 组合不一致时会静默输出错误报告）
5. Word 文档无页眉、无签名栏、无标准化单元格样式

**本次升级只借鉴 8D 的工程化框架**（样式工具、智能填充、行高计算、单位互转概念），**所有领域知识仍是 FMEA 自己的**（S/O/D、AP、失效链、PC/DC、CC/SC、FMEA 角色）。无任何 8D 领域概念（5Why、6M、D0-D8、不良率、batch_size）污染。

### 新增：样式工具函数库（Excel）

借鉴 8D Skill 的样式工具函数模式，但全部使用 FMEA 专用配色和领域概念：

- `get_thin_border()` / `get_header_font()` / `get_body_font()` / `get_subheader_font()`
- `get_header_fill()` / `get_subheader_fill()` / `get_alt_fill()` / `get_section_title_fill()`
- `get_ap_fill(ap)` / `get_special_char_fill(sc)` / `get_status_fill(status)`
- `apply_header_style(cell)` / `apply_subheader_style(cell)` / `apply_body_style(cell, ...)`
- `apply_section_title_style(cell)` / `apply_ap_style(cell, ap)` / `apply_special_char_style(cell, sc)` / `apply_status_style(cell, status)`
- `write_section_title(ws, row, title, span_cols)` / `write_table(ws, ..., ap_col, special_char_col, status_col)` / `write_kv_block(ws, ..., label_col_width, value_col_width)` / `set_column_widths(ws, widths)`

**FMEA 专用配色方案**（区别于 8D 的颜色）：
- 主色 `#1F4E79`（FMEA 深蓝表头）
- 次级 `#2E75B6`（亮蓝次级表头）
- 交替行 `#E7EEF7`（浅蓝）
- AP 配色：H=`#FF6B6B` 红 / M=`#FFD93D` 黄 / L=`#6BCB77` 绿（符合 2019 手册惯例）
- CC=`#FFC7CE`（关键特性红色淡底）/ SC=`#FFEB9C`（特殊特性黄色淡底）
- 措施状态：进行中=`#FFE699` 黄 / 已完成=`#C6EFCE` 绿

### 新增：智能填充系统（auto_fill 模式大幅增强）

借鉴 8D Skill 的"按上下文（左侧单元格 + 第一列）智能判断该填什么"模式，但完全使用 FMEA 自己的角色/部门/字段：

**新增函数**：
- `_guess_fill_value(left_val, first_col_val, col_idx, today_str, date_plus, fmea_type, fmea_no)`：核心智能判断函数
- `_apply_auto_fill_xlsx(ws, fmea_type, fmea_no)`：Excel 工作表级自动填充
- `_apply_auto_fill_word(doc, fmea_type, fmea_no)`：Word 文档级自动填充

**FMEA 团队角色化名**（基于手册第 1.5.3 节，区别于 8D 的角色定义）：
- FMEA 推进者 / 设计工程师 / 过程工程师 / 质量工程师 / 测试工程师 / 系统工程师 / 项目经理
- 化名池：张伟 / 李娜 / 刘强 / 陈静 / 赵磊 / 周敏 / 王芳 / 孙健
- 内部分机号：8001-8009

**智能填充规则**（按上下文判断）：
1. FMEA 团队表（第1列=角色名）：填姓名/部门/分机号
2. 签名栏（编制/审核/批准）：填姓名/签名/当天日期
3. 优化措施表（第1列=序号）：填责任人（按序号轮换）/截止日期（按 7/10/14/21/30/45/60 天递增）/状态（已建议）
4. 表头信息表：填项目编号 / FMEA 编号 / FMEA 编制人 / FMEA 审核人 / FMEA 批准人 / 制造地址 / 团队成员
5. 结构分析表（PFMEA 4M1E 要素）：填设备型号 / 岗位资质 / 材料牌号 / SOP 编号 / 温湿度 / 量具编号

**S/O/D 评分不自动填充**：保留模板 hint 值由用户基于实际评估确认，避免 LLM 臆造评分。

### 新增：行高自动计算（中文长文本不再截断）

借鉴 8D Skill 的行高计算方案，原样移植（这部分是纯工程逻辑，与领域无关）：

- `_display_width(s)`：计算字符串显示宽度（中文≈2，英文≈1）
- `_calc_row_height(row_data, col_widths, line_height=17, min_height=20, max_height=409)`：根据单元格内容和列宽精确计算行高
- `_recalc_all_row_heights(ws, line_height=17, min_height=20, max_height=409)`：所有内容写入 + auto_fill 执行完毕后，重新计算整个工作表所有行的行高

**核心逻辑**：
1. 遍历当前行每个单元格，计算该单元格文本在该列宽度下需要换几行
2. CJK 字符宽度按英文 2 倍计算
3. 显式换行符 `\n` 也增加行数
4. 合并单元格宽度 = 涉及所有列宽之和
5. 行高 = 最大行数 × 每行高度（17 磅）+ 上下边距（4 磅）
6. 限制在 20-409 磅之间（Excel 上限）

**实测效果**：失效分析表行高从原来的固定 60 磅 → 现在按内容自适应 38-55 磅，所有中文文本完整显示。

### 新增：SOD 一致性校验

新增 `check_sod_consistency(chain)` 函数，在生成报告前对每条失效链进行 SOD 逻辑校验：

- S≥9 时 AP 必为 H（除非 O=1 且 D=1）
- S=1 时 AP 必为 L
- AP 与 S/O/D 组合必须一致（用 `get_ap_priority` 重算对比）

发现不一致时打印 `[WARN]` 警告，但不阻止生成（让用户决定是否修正）。

同时修改 `auto_fill_failure_chain`：始终用 `get_ap_priority(s, o, d)` 重算 AP，不再采用模板预填的 `ap_hint`（因为部分模板的 hint 字段存在不一致）。

### 新增：Word 文档增强

- 新增 `set_page_header(doc, fmea_no)`：页眉显示 FMEA 编号
- 新增 `set_doc_default_font(doc)`：默认字体设为微软雅黑（中文）
- 新增 `set_cell_bg(cell, color_hex)` / `set_cell_borders(cell)` / `set_run_font(run, ...)`：Word 单元格样式工具
- 新增 `add_paragraph(doc, text, bold, size, color, alignment, indent)`：标准化段落添加
- 新增 `add_heading(doc, text, level)`：标准化标题添加（深蓝色）
- 新增 `add_table(doc, headers, rows, col_widths_cm, ap_col, special_char_col, status_col)`：通用表格（支持 AP/CC-SC/状态列高亮）
- 新增 `add_kv_table(doc, kv_pairs, label_width_cm, value_width_cm)`：键值对表格（左侧亮蓝标签 + 右侧值）
- 新增签名栏：第七章末尾自动添加 编制/审核/批准 三人签名表

### 优化：AP 矩阵函数更严格

`get_ap_priority(s, o, d)` 增加边界保护：所有输入 clamp 到 1-10 范围，避免非法值导致 AP 计算异常。

### 优化：占位符替换递归化

新增 `substitute_placeholders_deep(obj, context)` 函数，递归替换字典/列表中的占位符。原来只对顶层字段替换，现在模板中嵌套的占位符（如 `failure_chains_template[i].fe` 中的 `{product_name}`）也能正确替换。

### 文件清单变化

```
scripts/generate_fmea.py
  v2.0.0: 899 行 / 36 KB / 8 个函数
  v2.0.1: 1130+ 行 / 45+ KB / 30+ 个函数（含 15 个样式工具 + 3 个行高计算 + 3 个智能填充 + 5 个 SOD 校验）
```

### 兼容性

- 完全向后兼容：v2.0.0 的所有命令行参数和模板格式不变
- 新增 `--auto-fill` 行为增强：除了填充 S/O/D，还会填充表头/优化措施/签名栏的 ____ 
- 模板 `failure_chains_template` 中的 `ap_hint` 字段不再被采用（始终用 `get_ap_priority` 重算），但字段保留以兼容旧模板

### 实测验证

| 测试场景 | 失效链数 | AP=H/M/L | 智能填充项 | 一致性警告 |
|---|---|---|---|---|
| PFMEA + generic-fmea + auto_fill | 5 | 0/2/3 | 18 个表头字段 + 2 行优化措施 + 3 行签名 | 0 条 |
| DFMEA + electronic-ecm + auto_fill | 7 | 4/3/0 | 18 个表头字段 + 7 行优化措施 + 3 行签名 | 0 条 |

---

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
