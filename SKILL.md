# PFMEA/DFMEA Skill — 汽车行业 FMEA 分析报告生成器

> 基于 **AIAG & VDA FMEA 手册（2019 版）** 七步法（Seven-Step Approach），覆盖 DFMEA（设计 FMEA）与 PFMEA（过程 FMEA）两类分析。

---

## 一、概述

本 Skill 用于在用户描述产品/过程信息后，自动生成符合 2019 版 AIAG & VDA FMEA 手册要求的 FMEA 分析报告（.xlsx + .docx）。报告包含完整的七步法分析过程：从规划准备、结构分析、功能分析、失效分析、风险分析到优化措施和结果文件化，并基于 S/O/D 评分与 AP（Action Priority）矩阵输出风险优先级。

### 1.1 与 8D Skill 的区别

| 维度 | 8D Skill | PFMEA/DFMEA Skill |
|---|---|---|
| 触发场景 | 客户投诉 / SCAR / 8D 报告请求 | 设计/过程潜在失效预防分析 |
| 分析时点 | 事后问题解决 | 事前风险预防（APQP 阶段） |
| 输出物 | 8D 报告（D0-D8 八步） | FMEA 表格（七步法 + SOD + AP） |
| 核心方法 | 5Why + 6M 根因分析 | 失效链（FE→FM→FC）+ SOD 评分 |
| 评分体系 | 无 | S（严重度）+ O（频度）+ D（探测度）+ AP |
| 适用阶段 | 量产后问题处理 | 设计开发 + 过程开发阶段 |

### 1.2 关键约束

1. **必须使用 2019 版七步法**，不是旧版的"5T+6步"或更早的"RPN 评分"
2. **必须使用 AP（Action Priority）替代 RPN**，2019 版明确不再推荐 RPN 阈值
3. **DFMEA 与 PFMEA 的 AP 表相同**，但 FMEA-MSR 的 AP 表不同（本 Skill 暂不覆盖 MSR）
4. **评分必须基于手册的标准评分表**（详见 references/sod_scoring_tables.md），不可主观臆断

---

## 二、何时触发

### 2.1 触发关键词

- **明确 FMEA 字样**：DFMEA / PFMEA / FMEA-MSR / FMEA 评审 / FMEA 修订
- **设计/过程分析**：潜在失效模式 / 失效起因 / 失效影响 / 风险分析
- **评分相关**：严重度 S / 频度 O / 探测度 D / AP 行动优先级 / 措施优先级
- **方法论关键词**：七步法 / 参数图 / P-图 / 4M1E / 5M 分析 / 鱼骨图 / 结构树 / 功能树 / 失效链
- **行业场景**：APQP 阶段 / PPAP 提交 / IATF 16949 审核 / 控制计划 / 特殊特性 CC/SC

### 2.2 不应触发的场景

- ❌ 用户只要 8D 报告 → 使用 8d-skill
- ❌ 用户只要 SPC 控制图 → 使用 spc-chart-skill（如有）
- ❌ 用户只要过程能力指数 Cpk/Ppk → 使用 cpk-ppk-skill（如有）
- ❌ 用户问 FMEA 是什么（纯概念解释）→ 直接答概念即可，不需要生成报告

### 2.3 DFMEA vs PFMEA 判定

| 用户输入特征 | 判定为 |
|---|---|
| 提到"产品设计"、"零部件设计"、"DFMEA" | DFMEA |
| 提到"生产工艺"、"制造过程"、"PFMEA"、"工序" | PFMEA |
| 同时提到设计 + 过程 / 用户未明示 | **追问用户**："您是要做 DFMEA（设计 FMEA）还是 PFMEA（过程 FMEA）？" |

---

## 三、工作流

### Step 1：从用户输入提取关键信息

向用户提取以下字段（缺失的用追问，参考第 9.2 节追问模板）：

**通用字段（DFMEA/PFMEA 都需要）**

| 字段 | 必填 | 说明 | 示例 |
|---|---|---|---|
| `fmea_type` | ✅ | DFMEA 或 PFMEA | "DFMEA" |
| `product_name` | ✅ | 关注项目名称 | "前照灯 LED 模组" |
| `customer` | ✅ | 客户名 | "XX 汽车有限公司" |
| `project_no` | ⭕ | 项目编号 | "P2026-0123" |
| `team_members` | ⭕ | 团队成员 | "张三（设计）/李四（质量）/王五（工艺）" |
| `template` | ⭕ | 指定模板，否则自动匹配 | "electronic-ecm" |

**DFMEA 专属字段**

| 字段 | 必填 | 说明 |
|---|---|---|
| `system_level` | ✅ | 系统层级：整车 / 系统 / 子系统 / 组件 / 零件 |
| `design_responsibility` | ✅ | 设计责任方（自有设计 / 供应商设计 / 联合设计） |
| `interface_info` | ⭕ | 接口信息（机械/电气/软件接口） |

**PFMEA 专属字段**

| 字段 | 必填 | 说明 |
|---|---|---|
| `process_name` | ✅ | 工艺名称（如"注塑"、"焊接"、"装配"） |
| `process_steps` | ✅ | 工序清单（至少 3-5 个工序） |
| `manufacturing_site` | ⭕ | 制造地址 |

### Step 2：匹配最合适的模板

模板匹配优先级（参考 templates/INDEX.md）：

1. **第一优先级：用户明确指定 `template` 参数** → 直接使用
2. **第二优先级：产品类别关键字**
   - 含 "ECU/控制器/传感器/线束/PCB/电路" → `electronic-ecm`
   - 含 "齿轮/轴承/紧固件/轴/壳体/装配" → `mechanical-assembly`
   - 含 "电镀/热处理/氧化/表面处理/淬火" → `surface-treatment`
   - 含 "喷涂/电泳/漆面/涂装/喷漆" → `painting-coating`
3. **兜底：generic-fmea** → 无法明确匹配时使用

### Step 3：模板融合逻辑说明（generate_fmea_report_tool 自动完成，Agent 无需手动操作）

`generate_fmea_report_tool` 工具会：

1. 读取 `templates/<slug>/template.json`，获取该行业预填的失效链结构（FE/FM/FC）
2. 读取 `templates/<slug>/intro.md`，获取该行业的失效模式常识
3. 读取 `references/fmea_seven_step_guide.md`，注入七步法上下文
4. 读取 `references/sod_scoring_tables.md`，注入 S/O/D 评分表 + AP 矩阵
5. 读取 `references/failure_chain_examples.md`，注入失效链案例库
6. 将 `{product_name}` `{customer}` `{process_name}` 等占位符替换为用户输入
7. 调用 `scripts/generate_fmea.py` 生成 xlsx + docx 报告

### Step 3.5：动态失效链覆盖（可选，提升报告专业度）

如果用户在对话中已经讨论过具体的失效模式、起因、影响，可以将这些动态信息传入 `failure_chains` 参数，覆盖模板预填值：

```python
generate_fmea_report_tool(
    fmea_type="DFMEA",
    product_name="前照灯 LED 模组",
    customer="XX 汽车",
    template="electronic-ecm",
    failure_chains=[
        {
            "fe": "夜间行驶时前照灯突然熄灭，可能导致交通事故",
            "fm": "LED 模组开路",
            "fc": "LED 焊点因热应力开裂",
            "s": 10,  # 严重度：影响行车安全
            "o": 4,   # 频度：基于过往保修数据
            "d": 6,   # 探测度：老化测试可发现
            "ap": "H",  # 自动计算或显式传入
            "pc": "热仿真分析 + 焊点可靠性设计规范",
            "dc": "高温老化测试 1000h + 红外热成像"
        },
        # ... 更多失效链
    ]
)
```

### Step 4：调用 generate_fmea_report_tool 生成报告（唯一入口）

```python
# 基础调用（用模板预填失效链，空白处留 ____）
generate_fmea_report_tool(
    fmea_type="PFMEA",
    product_name="前保险杠总成",
    customer="XX 汽车",
    process_name="注塑成型",
    process_steps=["原料干燥", "注塑成型", "去毛刺", "外观检验", "包装"],
    template="generic-fmea",
    output_dir="~/Desktop"
)
```

```python
# 自动填充模式（用户说"你帮我填"/"给我示例"时启用）
generate_fmea_report_tool(
    fmea_type="DFMEA",
    product_name="电控单元 ECU",
    customer="XX 汽车",
    template="electronic-ecm",
    auto_fill=True,  # 自动填充 S/O/D 评分与措施
    output_dir="~/Desktop"
)
```

```python
# 进阶调用（传入动态失效链，覆盖模板预填）
generate_fmea_report_tool(
    fmea_type="PFMEA",
    product_name="传动轴总成",
    customer="XX 汽车",
    process_name="焊接",
    process_steps=["零件清洗", "自动焊接", "焊后检验", "热处理"],
    template="mechanical-assembly",
    failure_chains=[...],
    output_dir="~/Desktop"
)
```

### Step 5：向用户展示结果

报告生成后，向用户输出：

1. **七步法分析摘要**（先文字展示，让用户快速审阅）
   - 步骤一：项目信息、范围、团队
   - 步骤二：结构分析（结构树/过程流程图）
   - 步骤三：功能分析（功能树/参数图）
   - 步骤四：失效分析（失效链 FE→FM→FC）
   - 步骤五：风险分析（S/O/D 评分 + AP 优先级）
   - 步骤六：优化措施（PC/DC 改进 + 责任人 + 截止日期）
   - 步骤七：结果文件化（FMEA 表格 + 结论）
2. **风险统计**：高优先级（H）__项、中优先级（M）__项、低优先级（L）__项
3. **特殊特性清单**：识别出的 CC/SC 项
4. **下载链接**：FMEA.xlsx 与 FMEA.docx 文件路径

### 🔧 自动填充模式（auto_fill 参数）

当用户说"你帮我填"/"给我个示例"/"先填模板我看看"时，启用 `auto_fill=True`：

- **S 评分**：基于失效影响描述自动从 references/sod_scoring_tables.md 的 S 表中匹配
- **O 评分**：基于预防控制描述自动判定（无控制=9-10，标准控制=5-7，验证过的控制=2-4）
- **D 评分**：基于探测控制描述自动判定（无探测=10，通过/不通过=7-8，验证过的探测=2-4）
- **AP**：根据 S/O/D 组合查询 AP 矩阵自动判定（H/M/L）
- **措施**：对 AP=H 的失效链自动建议"加强预防控制"或"提升探测能力"

**禁止自动填充的内容**：
- ❌ 不可编造"过往保修数据"（除非用户已提供）
- ❌ 不可编造"客户投诉次数"（除非用户已提供）
- ❌ 不可编造"过程能力指数 Cpk"（除非用户已提供）

---

## 四、七步法详细说明

### 4.1 步骤一：规划与准备

**DFMEA**

- 5T 明确：Intent（目的）、Timing（时间）、Team（团队）、Task（任务）、Tools（工具）
- 项目边界确定：分析哪个系统层级（整车/系统/子系统/组件/零件）
- 基准 FMEA 引用：是否参考基础 FMEA / 家族 FMEA
- 工程团队协作接口：系统团队、安全团队、组件团队的接口

**PFMEA**

- 工艺边界确定：分析哪条生产线 / 哪个工序段
- 过程流程图（PFD）确认：列出所有工序步骤
- 基础 PFMEA 引用：是否参考类似产品 PFMEA
- 表头填写：公司名、制造地址、顾客名、年型、项目名、跨职能团队

### 4.2 步骤二：结构分析

**DFMEA**：使用结构树（Structure Tree）或方块图（Boundary Diagram）

```
整车
 └ 系统：照明系统
    └ 子系统：前照灯
       └ 组件：LED 模组
          └ 零件：LED 芯片 / 透镜 / 散热基板 / 驱动 IC
```

**PFMEA**：使用过程流程图 + 结构树 + 4M 工作要素

```
过程：注塑成型
 ├ 工序 1：原料干燥
 │   ├ 设备（M）：除湿干燥机
 │   ├ 人员（M）：操作工
 │   ├ 材料（M）：PC 原料
 │   └ 环境（M）：车间温湿度
 ├ 工序 2：注塑成型
 │   ├ 设备：注塑机
 │   ├ 人员：注塑工
 │   ...
```

### 4.3 步骤三：功能分析

**DFMEA**：使用功能树（Function Tree）或参数图（P-Diagram）

- 功能定义格式："主动动词 + 可测量名词"
- 例："将电能转换为光能"、"提供 800 lm 光通量"

参数图（P-图）要素：
- 输入（信号因素）
- 输出（预期输出 + 非预期输出/转向输出）
- 控制因素（可调整的设计参数）
- 噪声因素（5 类：组件间变化、随时间变化、顾客使用、外部环境、系统交互）

**PFMEA**：使用过程功能分析

- 工序功能 = "工序名 + 作用对象 + 期望结果"
- 例："注塑成型工序：将 PC 原料转化为符合图纸的保险杠壳体"
- 工作要素功能 = 4M 各自的功能

### 4.4 步骤四：失效分析

**核心概念：失效链（Failure Chain）**

```
失效影响（FE）── 失效模式（FM）── 失效起因（FC）
   ↑                ↑                ↑
为什么？         发生了什么？      为什么？
（后果）         （现象）         （根因）
```

**失效模式的 7 种类型**（手册原文）：

1. 功能丧失（无法操作、突然失效）
2. 功能退化（性能随时间损失）
3. 功能间歇（随机开始/停止）
4. 部分功能丧失（性能损失）
5. 非预期功能（错误时间/方向操作）
6. 功能超范围（超出可接受极限）
7. 功能延迟（非预期时间间隔后操作）

**DFMEA 失效起因来源**：
- 经典 P-图噪声因素（5 类）
- 接口失效（机械/电气/软件）
- 设计缺陷（材料选型/几何/容差）

**PFMEA 失效起因来源（4M/5M/6M 类型）**：
- 人（Man）：操作失误、培训不足
- 机（Machine）：设备故障、磨损、校准失效
- 料（Material）：来料异常、批次差异
- 法（Method）：SOP 缺失、参数错误
- 环（Environment）：温湿度、洁净度
- 测（Measurement）：测量系统误差

### 4.5 步骤五：风险分析

#### 4.5.1 S 严重度（Severity）评级

**表 D1（DFMEA）/ 表 P1（PFMEA）通用 10 级**

| S | 等级 | 标准 |
|---|---|---|
| 10 | 非常高 | 影响车辆操作安全，危及驾驶员/乘客/行人健康 |
| 9 | 非常高 | 不符合法规 |
| 8 | 高 | 预期寿命内失去车辆主要功能 |
| 7 | 高 | 预期寿命内降低车辆主要功能 |
| 6 | 中 | 失去车辆次要功能 |
| 5 | 中 | 降低车辆次要功能 |
| 4 | 低 | 外观/声音/振动/触感非常不舒服 |
| 3 | 低 | 外观/声音/振动/触感中度不舒服 |
| 2 | 低 | 外观/声音/振动/触感略微不舒服 |
| 1 | 非常低 | 没有可觉察的影响 |

⚠️ PFMEA 的 S 评分应与 DFMEA 一致（手册 1.4 节明确要求）

#### 4.5.2 O 频度（Occurrence）评级

**表 D2/P2 通用 10 级**

| O | 等级 | 标准（基于预防控制有效性） |
|---|---|---|
| 10 | 极高 | 新技术首次应用，无经验，预防控制不能预测现场绩效 |
| 9 | 非常高 | 技术创新设计首次应用，预防控制非针对特定性能 |
| 8 | 高 | 新应用内首次使用创新设计，预防控制不能可靠反映现场绩效 |
| 7 | 高 | 相似技术新型设计，预防控制提供有限性能指标 |
| 6 | 中 | 应用现有技术，类似应用，预防控制提供部分能力 |
| 5 | 中 | 成熟技术细节变化，预防控制能发现部分缺陷 |
| 4 | 高 | 与短期现场暴露几乎相同，预防控制能反映设计符合性 |
| 3 | 低 | 已知设计细微变化，预防控制能预测设计一致性 |
| 2 | 非常低 | 长期现场暴露几乎相同，预防控制显示设计符合性信心 |
| 1 | 非常低 | 通过预防控制完全消除失效起因 |

#### 4.5.3 D 探测度（Detection）评级

**表 D3/P3 通用 10 级**

| D | 等级 | 标准（探测方法成熟度 + 探测机会） |
|---|---|---|
| 10 | 非常低 | 尚未制定测试过程 |
| 9 | 非常低 | 未为探测失效模式/起因设计测试方法 |
| 8 | 低 | 新测试方法尚未验证 |
| 7 | 低 | 已验证测试方法，但测试时间较迟，失败将导致生产延迟 |
| 6 | 中 | 已验证测试方法，通过/不通过测试 |
| 5 | 中 | 失效测试 |
| 4 | 高 | 已验证测试方法，计划时间充分，可在生产前修改工装 |
| 3 | 高 | 失效测试，提前发现 |
| 2 | 非常高 | 老化测试，验证充分 |
| 1 | 非常高 | 测试证明不会出现失效，或探测方法总是能探测到 |

#### 4.5.4 AP 行动优先级（Action Priority）矩阵

**2019 版核心变化**：不再使用 RPN（=S×O×D），改用 AP 三档优先级

| S | O | D | AP |
|---|---|---|---|
| 9-10 | 任何 | 任何 | **H**（除非 O=1 且 D=1） |
| 9-10 | 2-4 | 1 | H |
| 9-10 | 1 | 1 | M |
| 6-8 | 6-10 | 任何 | H |
| 6-8 | 4-5 | 7-10 | H |
| 6-8 | 1-3 | 1-3 | L |
| 4-5 | 6-10 | 7-10 | H |
| 4-5 | 4-5 | 7-10 | M |
| 2-3 | 任何 | 任何 | L（除非 S=2-3 且 O=4-10 且 D=5-10 则为 M） |
| 1 | 任何 | 任何 | L |

**完整 1000 种组合详见**：references/sod_scoring_tables.md

**AP 处置原则**：
- **H（高优先级）**：必须评审并采取措施，或证明并记录当前控制足够有效
- **M（中优先级）**：应评审并采取措施，或公司自行决定证明并记录当前控制足够有效
- **L（低优先级）**：可以确定措施来改进预防或探测控制

⚠️ 对于 S=9-10 且 AP=H/M 的失效影响，建议至少由管理层评审

### 4.6 步骤六：优化

**优化措施分类**：
1. **预防控制（PC）改进**：降低 O 评分
   - 例：增加设计规范、增加 DOE、增加仿真分析
2. **探测控制（DC）改进**：降低 D 评分
   - 例：增加测试项目、提前测试时间、采用更敏感的测试方法
3. **严重度（S）改进**：通常难以降低，需通过设计变更
   - 例：增加冗余设计、增加故障安全机制

**措施状态**：
- 已建议（Suggested）
- 已决策（Decided）—— 含责任人、截止日期
- 已实施（Implemented）—— 含实施日期、验证结果
- 已关闭（Closed）—— 含关闭日期、关闭人

**措施有效性评估**：重新评估 S/O/D，计算 AP 是否降低

### 4.7 步骤七：结果文件化

**输出文件**：
1. **FMEA 表格（xlsx）**：包含 7 个 Sheet
   - Sheet 1：表头（项目信息）
   - Sheet 2：结构分析（结构树/过程流程图）
   - Sheet 3：功能分析（功能树/参数图）
   - Sheet 4：失效分析（失效链 FE→FM→FC）
   - Sheet 5：风险分析（S/O/D 评分 + AP + PC + DC）
   - Sheet 6：优化措施（措施清单 + 责任人 + 截止日期 + 状态）
   - Sheet 7：风险矩阵（S×O、S×D 热力图）
2. **FMEA 报告（docx）**：包含 7 个章节，文字描述 + 表格 + 图表

**特殊特性识别**：
- S=9-10 的失效影响 → CC（关键特性）
- S=8 且 AP=H/M 的失效影响 → SC（特殊特性）
- 在 PFMEA 中标识后，应同步到控制计划（CP）

---

## 五、模板选择指南

详见 `templates/INDEX.md`。

### 5.1 template.json 数据结构（v2.0+）

```json
{
  "slug": "electronic-ecm",
  "name": "电子电气 DFMEA 模板",
  "applicable_types": ["DFMEA"],  // DFMEA / PFMEA / Both
  "defect_types": ["电气失效", "短路", "开路", "EMC 超标", ...],
  "product_categories": ["ECU", "传感器", "线束", "PCBA", ...],
  
  "structure_tree_template": { ... },      // 步骤二：结构树预填
  "function_tree_template": { ... },       // 步骤三：功能树预填
  "failure_chains_template": [             // 步骤四：失效链预填
    {
      "fe": "____（失效影响）",
      "fm": "____（失效模式）",
      "fc": "____（失效起因）",
      "s_hint": 8,  // 建议评分
      "o_hint": 5,
      "d_hint": 6,
      "pc": "____（当前预防控制）",
      "dc": "____（当前探测控制）"
    }
  ],
  "optimization_measures": [ ... ],        // 步骤六：优化措施建议
  "special_characteristics": [ ... ]       // 特殊特性标识
}
```

### 5.2 行业标准参考

- **DFMEA**：参考 IATF 16949 第 8.3 条款（产品和服务的设计和开发）
- **PFMEA**：参考 IATF 16949 第 8.5 条款（生产和服务提供）
- **FMEA-MSR**：参考 ISO 26262 功能安全标准（本 Skill 暂不覆盖）

---

## 六、references/ 目录说明

| 文件 | 作用 |
|---|---|
| `fmea_seven_step_guide.md` | 七步法每一步的目的、关键活动、输出物、常见错误、IATF 关联 |
| `sod_scoring_tables.md` | S/O/D 三张评分表（10 级）+ 完整 1000 种 AP 组合矩阵 |
| `failure_chain_examples.md` | 各行业的失效链（FE→FM→FC）真实案例库 |

---

## 七、完整调用示例

### 7.1 简单 PFMEA（用模板预填，auto_fill 模式）

**用户**："帮我做一份前保险杠注塑成型的 PFMEA，客户是 XX 汽车，给我个示例。"

**Agent 调用**：

```python
generate_fmea_report_tool(
    fmea_type="PFMEA",
    product_name="前保险杠总成（注塑件）",
    customer="XX 汽车有限公司",
    process_name="注塑成型",
    process_steps=["原料干燥", "注塑成型", "去毛刺", "外观检验", "包装"],
    template="generic-fmea",
    auto_fill=True,
    output_dir="~/Desktop"
)
```

### 7.2 进阶 DFMEA（用户已讨论过失效链，覆盖模板）

**用户**："我们 ECU 在客户现场出现过 LED 驱动 IC 过热失效，S 评分应该是 10 分（影响行车安全），频度大概 4 分，探测度 6 分。做一份 DFMEA 把这个失效链写进去。"

**Agent 调用**：

```python
generate_fmea_report_tool(
    fmea_type="DFMEA",
    product_name="ECU 控制单元",
    customer="XX 汽车",
    system_level="组件",
    design_responsibility="自有设计",
    template="electronic-ecm",
    failure_chains=[
        {
            "fe": "ECU 控制单元在客户使用中过热失效，可能导致发动机控制异常，影响行车安全",
            "fm": "LED 驱动 IC 热失控",
            "fc": "IC 散热设计不足，热阻超标导致结温超过最高允许温度",
            "s": 10,
            "o": 4,
            "d": 6,
            "ap": "H",
            "pc": "热仿真分析 + IC 选型规范 + 散热基板设计",
            "dc": "高温老化测试 1000h + 红外热成像 + 客户现场温度监测"
        }
    ],
    output_dir="~/Desktop"
)
```

### 7.3 完整 PFMEA（多工序 + 多失效链）

```python
generate_fmea_report_tool(
    fmea_type="PFMEA",
    product_name="传动轴总成",
    customer="XX 汽车",
    process_name="焊接与机加",
    process_steps=["零件清洗", "自动焊接", "焊后检验", "热处理", "机加精加工"],
    template="mechanical-assembly",
    failure_chains=[
        # 工序 1：零件清洗
        {
            "process_step": "零件清洗",
            "fe": "焊接质量不合格，导致传动轴强度不足",
            "fm": "零件表面油污未清洗干净",
            "fc": "清洗剂浓度不足（人/法）",
            "s": 8, "o": 5, "d": 6, "ap": "H",
            "pc": "清洗 SOP 规定浓度范围 + 每班次检测浓度",
            "dc": "清洗后清洁度检验（白布擦拭法）"
        },
        # 工序 2：自动焊接
        {
            "process_step": "自动焊接",
            "fe": "传动轴在使用中断裂，可能导致车辆失控",
            "fm": "焊缝虚焊/未熔合",
            "fc": "焊接电流偏移（机）",
            "s": 10, "o": 4, "d": 5, "ap": "H",
            "pc": "焊接参数 SPC 监控 + 焊机每日点检",
            "dc": "超声波探伤 100% + 拉伸强度抽检"
        },
        # ... 更多失效链
    ],
    output_dir="~/Desktop"
)
```

---

## 八、关键约束

### 8.1 必须遵守

1. ✅ 必须使用 2019 版七步法，不能简化为旧版"5T+6步"
2. ✅ 必须使用 AP 替代 RPN，不能输出 RPN 评分
3. ✅ S/O/D 评分必须基于 references/sod_scoring_tables.md 的标准表
4. ✅ DFMEA 与 PFMEA 的 S 评分必须一致（手册 1.4 节要求）
5. ✅ 失效链必须按 FE→FM→FC 三级结构，不能跳级
6. ✅ PFMEA 必须按 4M1E（或 5M/6M）分类失效起因
7. ✅ 必须区分预防控制（PC）与探测控制（DC）
8. ✅ S=9-10 且 AP=H/M 的失效必须管理层评审

### 8.2 严禁行为

1. ❌ 严禁编造"过往保修数据"或"客户投诉次数"
2. ❌ 严禁编造"过程能力指数 Cpk"或"PPM 数据"
3. ❌ 严禁将 S/O/D 评分设为 0（最低为 1）
4. ❌ 严禁用 RPN 阈值（如 RPN>100 必须改进）判定措施
5. ❌ 严禁将 DFMEA 失效模式直接复制到 PFMEA（两者分析对象不同）
6. ❌ 严禁省略步骤二/三（结构/功能分析），直接做失效分析

---

## 九、文件清单

```
pfmea-dfmea-skill/
├── README.md                          项目说明（4.9 KB）
├── SKILL.md                           ★核心工作流定义（本文件）
├── VERSION                            当前版本号
├── CHANGELOG.md                       版本变更记录
├── .gitignore
├── scripts/
│   └── generate_fmea.py               ★FMEA 报告生成器
├── references/
│   ├── fmea_seven_step_guide.md       七步法详细指南
│   ├── sod_scoring_tables.md          S/O/D 评分表 + AP 矩阵
│   └── failure_chain_examples.md      失效链案例库
└── templates/
    ├── INDEX.md                       模板索引
    ├── electronic-ecm/                电子电气模板
    ├── mechanical-assembly/           机械装配模板
    ├── surface-treatment/             表面处理模板
    ├── painting-coating/              涂装模板
    └── generic-fmea/                  通用兜底模板
```

---

## 十、行业常识基准（🔴 严禁违反）

### 10.1 S 评分基准

- **S=10**：仅用于"影响行车安全"或"危及人身健康"的场景
  - 例：制动失灵、转向失控、安全气囊误爆、电池热失控
- **S=9**：仅用于"不符合法规"
  - 例：排放超标、噪音超国标、灯具不符合 ECE 法规
- **S=8**：用于"预期寿命内失去主要功能"
  - 例：发动机无法启动、空调完全不制冷、车窗无法升降
- **S=7**：用于"预期寿命内降低主要功能"
  - 例：发动机功率下降 10%、空调制冷不足
- **S=1-3**：外观类问题，不可与 S=9-10 混淆

### 10.2 O 评分基准

- **O=10**：仅用于"新技术首次应用、无任何经验"
- **O=9**：仅用于"技术创新设计首次应用"
- **O=5-6**：成熟技术 + 类似应用（最常见的量产产品评分区间）
- **O=1-2**：必须基于充分的现场数据，不可轻易给出
- ⚠️ 严禁 O=10 配合"已使用 10 年的成熟产品"——逻辑矛盾

### 10.3 D 评分基准

- **D=10**：仅用于"尚未制定测试过程"
- **D=9**：仅用于"未为探测失效模式/起因设计测试"
- **D=1-2**：必须基于"测试方法经验证总是能探测到"
- ⚠️ 严禁 D=10 配合"100% 在线自动检测"——逻辑矛盾

### 10.4 AP 判定合规性

- **S=9-10 时**：除非 O=1 且 D=1，否则 AP 必为 H
- **S=1 时**：AP 必为 L（不论 O/D 如何）
- 严禁出现 "S=10, O=10, D=10, AP=L" 这种荒谬组合

### 10.5 AskUserQuestion 追问时的合规示例

✅ **合规追问**：
- "您是要做 DFMEA 还是 PFMEA？"
- "分析的系统层级是整车、系统、子系统还是组件？"
- "客户在现场是否出现过该失效？是否涉及法规或安全？"
- "当前是否有预防控制措施？"
- "当前是否有探测控制措施？测试方法是否经过验证？"

❌ **不合规追问**：
- "您希望 S 评分是多少？"（应由 Agent 基于失效影响判定）
- "您希望 AP 是高、中还是低？"（应由 S/O/D 组合自动计算）
- "您要 RPN 大于多少？"（已废弃 RPN）

---

## 十一、版本与更新

- 当前版本：2.0.0（详见 VERSION）
- 发布日期：2026-06-27
- 变更记录：详见 CHANGELOG.md
- 参考标准：AIAG & VDA FMEA Handbook 2019

---

## 十二、与 JLAGENT 集成

本 Skill 设计为可独立运行，也可作为 JLAGENT 项目的 submodule。集成方式：

```bash
cd /path/to/JLAGENT
git submodule add https://github.com/cy556-like/PFMEA-DFMEA-SKILL.git skills/pfmea-dfmea-skill
git commit -m "feat(skills): 集成 PFMEA/DFMEA 分析 Skill"
```

后端注入逻辑参考 8d-skill 的 `_load_8d_skill_context()` 函数，新增 `_load_fmea_skill_context()` 即可。
