# FMEA 模板索引

> 5 套行业模板，覆盖汽车零部件常见 FMEA 分析场景。根据用户输入的产品类型智能匹配模板。

## 完整模板清单

| slug | 模板名称 | 适用 FMEA 类型 | 适用场景 | 一句话特点 |
|---|---|---|---|---|
| `electronic-ecm` | 电子电气 DFMEA 模板 | DFMEA / PFMEA | ECU、传感器、线束、PCBA | 电气失效、短路、开路、EMC 超标 |
| `mechanical-assembly` | 机械装配 DFMEA/PFMEA 模板 | DFMEA / PFMEA | 齿轮、轴承、紧固件、轴 | 断裂、磨损、松动、装配错装 |
| `surface-treatment` | 表面处理 PFMEA 模板 | PFMEA | 电镀、热处理、氧化、淬火 | 镀层附着力、变形超差、硬度不足 |
| `painting-coating` | 涂装 PFMEA 模板 | PFMEA | 喷涂、电泳、漆面 | 漆面颗粒、流挂、色差 |
| `generic-fmea` | 通用 FMEA 模板（兜底） | DFMEA / PFMEA | 其他/未分类 | 兜底模板，根据描述自适应 |

## 按产品类别快速筛

| 产品类别 | 首选模板 |
|---|---|
| ECU、传感器、线束、PCBA、连接器、继电器 | electronic-ecm |
| 齿轮、轴承、紧固件、轴、壳体、支架 | mechanical-assembly |
| 电镀件、热处理件、氧化件、淬火件 | surface-treatment |
| 喷涂件、电泳件、漆面件、涂装件 | painting-coating |
| 紧固件、橡胶件、塑料件、电子件（其他类） | generic-fmea |

## 按 FMEA 类型快速筛

| FMEA 类型 | 适用模板 |
|---|---|
| DFMEA（设计 FMEA） | electronic-ecm / mechanical-assembly / generic-fmea |
| PFMEA（过程 FMEA） | electronic-ecm / mechanical-assembly / surface-treatment / painting-coating / generic-fmea |

## 按失效模式快速筛

| 失效模式关键字 | 候选模板 |
|---|---|
| 短路/开路/EMC/电气失效 | electronic-ecm |
| 断裂/磨损/松动/装配错装 | mechanical-assembly |
| 附着力/镀层/变形/硬度 | surface-treatment |
| 漆面颗粒/流挂/色差 | painting-coating |
| 其他 | generic-fmea |

## 模板匹配优先级

1. **第一优先级：用户明确指定 `template` 参数** → 直接使用
2. **第二优先级：产品类别关键字**
   - 含 "ECU/控制器/传感器/线束/PCB/电路/继电器" → `electronic-ecm`
   - 含 "齿轮/轴承/紧固件/轴/壳体/支架/装配" → `mechanical-assembly`
   - 含 "电镀/热处理/氧化/表面处理/淬火/渗碳" → `surface-treatment`
   - 含 "喷涂/电泳/漆面/涂装/喷漆" → `painting-coating`
3. **兜底：generic-fmea** → 以上两者均无法明确匹配时使用

## 模板数据结构

每个 `template.json` 包含以下字段：

- `slug`：模板唯一标识
- `name`：模板中文名
- `applicable_types`：适用的 FMEA 类型（DFMEA / PFMEA / Both）
- `defect_types`：覆盖的缺陷类型清单
- `product_categories`：适用的产品类别
- `structure_tree_template`：步骤二结构分析预填
- `function_tree_template`：步骤三功能分析预填
- `failure_chains_template`：步骤四失效链预填（含 FE/FM/FC）
- `optimization_measures`：步骤六优化措施建议
- `special_characteristics`：特殊特性标识

占位符约定：

- `{product_name}`：替换为用户输入的产品名
- `{customer}`：替换为用户输入的客户名
- `{process_name}`：替换为用户输入的工艺名（PFMEA）

## 模板使用说明

### DFMEA 使用流程

1. 根据 `structure_tree_template` 构建结构树
2. 根据 `function_tree_template` 构建功能树
3. 根据 `failure_chains_template` 列出失效链
4. 评估每条失效链的 S/O/D，计算 AP
5. 根据 `optimization_measures` 制定措施
6. 识别 `special_characteristics` 中的 CC/SC

### PFMEA 使用流程

1. 根据 `structure_tree_template` 构建过程流程图
2. 根据 `function_tree_template` 列出工序功能
3. 根据 `failure_chains_template` 列出失效链（按 4M1E 分类）
4. 评估每条失效链的 S/O/D，计算 AP
5. 根据 `optimization_measures` 制定措施
6. 识别 `special_characteristics` 中的 CC/SC
