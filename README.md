# PFMEA/DFMEA Skill — 汽车行业 FMEA 分析报告生成器

> 基于 **AIAG & VDA FMEA 手册（2019 版）** 七步法（Seven-Step Approach）开发的智能 FMEA 分析 Skill，支持 **DFMEA（设计 FMEA）** 与 **PFMEA（过程 FMEA）** 两类分析，自动输出符合 IATF 16949 体系要求的 .xlsx 与 .docx 报告。

---

## 一、核心能力

| 能力 | 说明 |
|---|---|
| **七步法覆盖** | 步骤一规划准备 → 步骤二结构分析 → 步骤三功能分析 → 步骤四失效分析 → 步骤五风险分析 → 步骤六优化 → 步骤七结果文件化 |
| **DFMEA + PFMEA 双引擎** | 设计 FMEA 使用参数图（P-图）+ 结构树；过程 FMEA 使用过程流程图 + 4M1E 工作要素 |
| **SOD 评分自动判定** | 内置 2019 版 S/O/D 三张 10 级评分表 + 1000 种组合的 AP 行动优先级矩阵（H/M/L） |
| **5 套行业模板** | 电子电气（ECU/传感器）、机械装配（齿轮/轴承）、表面处理（电镀/热处理）、涂装（漆面/喷涂）、通用兜底 |
| **失效链自动构建** | FE（失效影响）→ FM（失效模式）→ FC（失效起因）三级链路，符合 VDA 标准格式 |
| **6M 因子分析** | 人/机/料/法/环/测 全维度根因排查，与鱼骨图（Ishikawa）兼容 |
| **预防控制 + 探测控制双栏** | PC（Prevention Control）与 DC（Detection Control）分离，符合新版手册要求 |
| **风险矩阵可视化** | 自动生成 S×O、S×D 风险矩阵热力图（嵌入 xlsx） |
| **特殊特性标识** | 自动标识 CC/SC（关键特性/特殊特性），关联 PFMEA 与控制计划 |

---

## 二、何时触发

当用户出现以下任意场景时，应触发本 Skill：

1. **明确提到 FMEA**："做一份 DFMEA"、"分析 PFMEA"、"FMEA 评审"、"FMEA-MSR"
2. **涉及设计/过程失效分析**："潜在失效模式"、"失效起因"、"失效影响"、"风险分析"
3. **涉及 SOD 评分**："严重度评分"、"频度评级"、"探测度评级"、"AP 高中低"
4. **涉及 2019 新版变化**："措施优先级 AP"、"七步法"、"FMEA-MSR"、"参数图 P-图"
5. **IATF 16949 体系审核场景**："APQP 阶段 FMEA"、"PPAP 提交 FMEA"、"控制计划关联"
6. **行业术语**："5M 分析"、"4M1E"、"鱼骨图"、"特殊特性"、"CC/SC"

**禁止触发**：用户只想要 8D 报告（请用 8d-skill）、只想要 SPC 控制图（请用 spc-chart-skill）。

---

## 三、目录结构

```
pfmea-dfmea-skill/
├── README.md                          本文件
├── SKILL.md                           ★核心工作流定义（Agent 主提示词）
├── VERSION                            当前版本号
├── CHANGELOG.md                       版本变更记录
├── .gitignore
├── scripts/
│   └── generate_fmea.py               ★FMEA 报告生成器（xlsx + docx）
├── references/
│   ├── fmea_seven_step_guide.md       七步法详细指南
│   ├── sod_scoring_tables.md          S/O/D 评分表 + AP 矩阵
│   └── failure_chain_examples.md      失效链（FE→FM→FC）案例库
└── templates/
    ├── INDEX.md                       模板索引
    ├── electronic-ecm/                电子电气（ECU/传感器/线束）
    │   ├── intro.md
    │   └── template.json
    ├── mechanical-assembly/           机械装配（齿轮/轴承/紧固件）
    │   ├── intro.md
    │   └── template.json
    ├── surface-treatment/             表面处理（电镀/热处理/氧化）
    │   ├── intro.md
    │   └── template.json
    ├── painting-coating/              涂装（漆面/喷涂/电泳）
    │   ├── intro.md
    │   └── template.json
    └── generic-fmea/                  通用兜底模板
        ├── intro.md
        └── template.json
```

---

## 四、与 JLAGENT 集成方式

本 Skill 设计为可独立运行，也可作为 JLAGENT 项目的 submodule。集成方式：

### 4.1 作为 submodule 添加

```bash
cd /path/to/JLAGENT
git submodule add https://github.com/cy556-like/PFMEA-DFMEA-SKILL.git skills/pfmea-dfmea-skill
git commit -m "feat(skills): 集成 PFMEA/DFMEA 分析 Skill"
git push
```

### 4.2 后端注入逻辑（参考 8d-skill 的 `_load_8d_skill_context`）

在 `app/agent/core.py` 中新增：

```python
def _load_fmea_skill_context(skill: str, user_input: str) -> str:
    """加载 PFMEA/DFMEA skill 的完整工作流上下文。"""
    if not skill or skill != "pfmea-dfmea-skill":
        return ""
    # 读取 SKILL.md
    # 匹配模板（electronic-ecm / mechanical-assembly / surface-treatment / painting-coating / generic-fmea）
    # 加载 template.json + references/
    # 拼接到 system prompt 末尾
    ...
```

### 4.3 前端按钮（参考 8d-skill）

在 `index.html` 的 skills 下拉菜单中：

```html
<button onclick="selectSkill('pfmea-dfmea-skill')" role="menuitem">
    📐 PFMEA/DFMEA分析
</button>
```

---

## 五、参考标准

- **AIAG & VDA FMEA Handbook**（2019 新版，本 Skill 的主要依据）
- **IATF 16949:2016** 汽车行业质量管理体系
- **VDA Volume 4 Part 2&3**（德国汽车工业联合会，已被本手册部分替代）
- **ISO 9001:2015** 第 6.1 条款（风险与机遇）
- **ISO 26262** 功能安全标准（与 FMEA-MSR 关联）

---

## 六、版本

- 当前版本：`2.0.2`（详见 VERSION）
- 变更记录：见 CHANGELOG.md
- 首次发布：2026-06-27
- 作者：cy556-like（参考 AIAG & VDA 2019 手册整理）

---

## 七、License

仅供 cy556-like 内部使用，基于公开标准（AIAG & VDA FMEA 手册 2019）的方法论整理。
