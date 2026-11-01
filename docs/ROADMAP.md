# Solidworks_AssetMate_MechDog 项目规划

> 一个会出图、也能证明这张图不能交的 SolidWorks MCP 服务。
> 仓库名与 MCP 注册名固定为 `Solidworks_AssetMate_MechDog`。

## 1. 目标与非目标

**目标**

1. 让智能体能安全地读、改、导出 SolidWorks 装配体、零件与工程图，不靠人肉点界面。
2. **对工程图给出可执行的交付判定**：不是"生成成功"，而是"缺哪 7 个尺寸、哪 3 项不满足，能不能交"。
3. 判定标准必须有实测依据，不靠拍脑袋——见第 2 节的真实图纸基线。
4. 每个工具都带「实测过的坑」，降低后来者的试错成本。这是与普通 API 封装的区别。
5. 所有写操作可预演、可回退，COM 失败不产生假结果。

**非目标**

- 不做建模/特征创建（与 solidpilot 等建模类工具重叠，且风险高）
- 不做云端服务、不做跨平台（SolidWorks COM 只在 Windows）
- 不做图形界面；交互由 AI 对话与返回的 JSON 承担
- **不做 DXF/DWG 逆向建模**（SolidPilot 已覆盖，我们只读 SolidWorks 原生图纸）
- 不追求「支持所有 SolidWorks 版本」，先保证 2020–2024 主路径可用

## 2. 现状（v0.8.2，2026-10-29）

> 本节是写入时的快照。最新进度以第 4 节路线图为准。

**48 个工具，三个能力域。**

| 域 | 工具数 | 代表能力 | 实测状态 |
|---|---|---|---|
| 装配与零件 | 13 | 上色 / BOM / 质量属性 / 材质 / 重命名 / **特征台账** | 46 组件 / 19 种零件全程跑通 |
| 工程图 | 15 | 建图 / 剖视 / 局部 / 展开 / 标注 / 公差 / **粗糙度** / 标题栏 / 导出 | 配电箱两张 A3 零警告 |
| 自检与判定 | 6 | 审计 / 注释清单 / 关联校验 / **台账对账** / **交付判定** / **基线抽样** | 100 张真实图 100% 可读 |

**真实图纸基线**（6346 张 SLDDRW 抽样 100 张，是全部阈值的依据）：

| 指标 | 平均 | 中位 | 范围 |
|---|---|---|---|
| 视图 | 3.1 | 3 | 1–9 |
| 尺寸 | 14.2 | 11 | 0–58 |
| 注释 | 46.5 | 36 | 0–160 |
| 中心标记 | 4.9 | 3 | 0–59 |

标注出现率：表面粗糙度 60%、形位公差 7%、零件序号 3%。

**能力边界**（详见 README）：形位公差写不进去（API 层缺失）、孔标注不可用（扫边性能）、
A3 最多 3 视图、批量读图必须逐张关闭、`material_library` 解析失效。

## 3. 用户与场景

| 用户 | 场景 | 关键诉求 |
|---|---|---|
| 机械工程师 | 出成套工程图、判断能否交付 | 一次出对、缺什么能说清 |
| 工艺/检验 | 审核外来图纸 | 能快速指出缺失与不可追溯的标注 |
| AI 使用者 | 用自然语言操作 CAD | 返回结构化、错误可读 |
| 二次开发者 | 把 CAD 能力接入自己的 Agent | 安装简单、接口稳定 |
| 逆向工程流程（自有） | 读取组件/配合生成 IR | 拓扑与属性完整、可交叉验证 |

## 4. 路线图

| 版本 | 主题 | 状态 |
|---|---|---|
| v0.1 | 上色 / 组件树 / 质量属性 / 预览 | ✅ 已发布 |
| v0.2 | BOM / 批量 STEP / 重命名 | ✅ 已发布 |
| v0.3 | 工程图链路：建图 / 视图 / 标注 / 导出 | ✅ 已发布 |
| v0.4 | 台账 / 尺寸公差 / 注释清单 / 关联校验 / 交付判定 / 对账 / 材料 / 粗糙度 / 基线校准 | ✅ |
| v0.4.6–0.4.9 | 基线抽样工具、静态检查（挡 NameError）、材料库 UTF-16、rules 契约、BOM 双算 | ✅ |
| v0.4.10–0.4.13 | **auto_fit 修复链**：实测安全区算比例 → 建完测量降档 → 重排前强制重建；三尺度验收 | ✅ |
| **v0.5** | 孔标注可用化 / 能力声明 / 出图前规划 | ✅ 已发布 |
| **v0.6** | 形位公差可创建 / 一站式出图流水线 `produce_drawing` | ✅ 已发布（0.6.1） |
| **v0.7** | 排版能力（视图摆位/尺寸删减/注记/件号/图纸直线）+ 钣金 QC + 展开依据声明 | ✅ 已发布（0.7.0） |
| **v0.8** | **BOM 明细栏可用**（10 行 4 列含数量列）+ 第四道静态防线 | ✅ 已发布（0.8.0–0.8.32） |
| v0.9 | 几何测距（最小折弯边/孔到折弯线）、尺寸精准筛选、图面利用率、形位公差符号 | 🔜 下一个 |
| v1.0 | 接口冻结、多配置与显示状态、升级说明 | 远期 |

### v0.5 / v0.6 已完成

| 条目 | 结果 |
|---|---|
| ~~auto_fit 对装配可靠~~ | ✅ 0.4.13：1:1 / 1:5 / 1:20 三种尺度全过 |
| ~~孔标注可用化~~ | ✅ 0.5.0：可见边 + `IsHoleCallout`，单视图 12 s；背面孔如实记 failed |
| ~~能力声明~~ | ✅ 0.5.0：`check_capabilities` 三类 + 证据 + 替代 |
| ~~出图前规划~~ | ✅ 0.5.0：`plan_drawing`，41 → 8 个尺寸 |
| ~~形位公差~~ | ✅ 0.6.0：根因是 pywin32 属性派发，改 `DISPATCH_METHOD` 即可；符号仍空 |
| ~~一站式流水线~~ | ✅ 0.6.1：`produce_drawing` 一次调用跑完 9 步 + 类别默认值 |

### v0.7 / v0.8 已完成

| 条目 | 结果 |
|---|---|
| ~~视图摆位~~ | ✅ 0.7.0：`view_positions` / `set_view` / `place_views`；`auto_fit` 不再覆盖显式 scale |
| ~~尺寸删减~~ | ✅ 0.7.0：`delete_dimensions` 按值保留，17 条 → 2 条 |
| ~~注记/件号/图纸直线~~ | ✅ 0.7.0：`add_note` / `add_balloon` / `add_sheet_line`（图纸草图单位是**米**） |
| ~~钣金 QC~~ | ✅ 0.7.0：`sheet_metal_qc` 板厚/弯折数/止裂槽；几何测距项标 `not_checked` |
| ~~展开依据声明~~ | ✅ 0.7.0：`flat_pattern_report` 声明展开由 SW 算、工具不读 K 因子 |
| ~~BOM 明细栏~~ | ✅ 0.8.0：Indented → TableType=2 → 隐藏重复列，10 行 4 列含数量列 |
| ~~同名函数防线~~ | ✅ 0.8.0：`static_check` 检测同作用域重复 `def` |
| ~~回归与冒烟体系~~ | ✅ 0.8.17：`regressions.py` 17 条静态断言 + `smoke.py` 29 项；每个 smoke 调用绑定到真实签名，server 包装层有静态契约检查 |
| ~~v0.8.6 反馈收尾~~ | ✅ 0.8.18：P2-12 `set_view` 对齐回执、P1-8 最近边失败诊断、P2-13 `precise_outline` 去掉 11.9 mm 虚胖、P3-15 两处过期文档；回归断言 17 → 21 |
| ~~v0.8.6 反馈收尾（续）~~ | ✅ 0.8.19：P1-6 `auto_dimension` 按模型 full name 自动删除重复副本（仅本次插入）、P0-2 新增 `set_dimension_position`（类型守卫 + 数值回读）、P1-7 `add_hole_callouts` 支持 `leader_x_m`/`leader_y_m` 引线落点；回归断言 21 → 24、smoke 29 → 30 |
| ~~v0.8.6 反馈收尾（补）~~ | ✅ 0.8.20：P2-9 `add_detail_view` 请求比例不再被父视图覆盖 + 空视图/比例被锁结构化报错、P2-10 形位公差符号格改中文文字 fallback（`GetFrameSymbols2` 回读校验，写不进就删框）、P2-11 表面粗糙度修正位置谎报（`SetPosition` + 回读）+ PDF 像素可见性自检；回归断言 24 → 27、smoke 30 → 31 |
| ~~P1-4 剖切算法~~ | ✅ 0.8.21：`add_section_view` 支持模型坐标剖切面（`cut_plane_normal`/`cut_plane_origin`，按视图 outline 与零件包围盒尺寸匹配定轴，斜切/平行图面/歧义全部结构化拒绝）+ 默认 `verify_cut` 用「剖切不改变未切方向尺寸」做 sanity（切错返回 `section_suspicious`，空剖返回 `section_empty`）；回归断言 27 → 30 |
| ~~P0-1 / P2-13 全面收尾~~ | ✅ 0.8.22：**P0-1** `get_sw` 用 18 ms 探针每 2 s 校验缓存、失败即丢；`call_method` 区分忙/断线（忙退避重试 2 次、断线丢缓存抛 `BridgeLost`）；新增 `ping` 工具做真实往返并报 `recovered`；未处理异常 flush stderr + 写日志；修掉 `_LAST_CALL` 无初值导致**首次失败**日志写不出。**P2-13** outline 单一真源：删死代码 `ink_outline`（三条取点路实测失败），投影轴推断与点命中改用 `precise_outline`，detail 视图返回 `precise_outline_m`；回归断言 30 → 34、smoke 31 → 32、工具 54 → 55 |
| ~~切割线初始位置~~ | ✅ 0.8.32：两个 bug 叠在一个表达式里。① `half = max(width,height)*0.62` **多除了一个 `scale`**——`width` 取自 `GetOutline` 本就是图纸米，实测 `0.1669×0.62/0.1 = 1.0347 m`，而 A3 只有 420 mm（这个 1.0347 正是 `GetLineInfo` 报的 ± 值）；② `CreateLine` 用绝对坐标**没加视图位置**，线画在了 x=0（图纸左边缘）。修法：`half` 改用 `precise_outline` 的宽高（不除 scale，也不受 11.9 mm 虚胖影响），`CreateLine` 加 `pos`。实测线长 0.1035 m、x = 视图中心 0.3133、y 覆盖视图上下各约 80 mm，`sanity ratio 1.0`；回归断言 64 → 65 |
| ~~气缸盖 P0-3（切割线跟随）~~ | ✅ 0.8.31：**四组接口实测后判定 API 无解，改为声明 + 流程**。`IDrSection.GetLineInfo` 是**模型坐标**（±1.03 m，A3 只有 420 mm），移动视图后数值不变；`SetLineInfo(VARIANT)` 确实能写（返回 True），但改的是**模型空间剖切平面**，等于切到别处；`IView.GetSectionLineInfo2` 只读且同样是模型坐标；`IView.GetSketchCount`/`GetSketches` **晚绑定下不存在**，触达不了图纸上的线。渲染核对还发现更基本的问题：`add_section_view` 建出的切割线**原本就不在父视图上**。故新增 `section_line_follows_view`（`unsupported`），workaround = 先摆视图再建剖面、已移过就重建；`add_section_view` docstring 同步警示。能力表 22 → 23 条、回归断言 63 → 64 |
| ~~气缸盖 P0-5~~ | ✅ 0.8.30：**证伪 + 给替代**。`InsertModelAnnotations3` 是文档级幂等——同一驱动尺寸插过一次，再次调用**一律返回 None**：`all_views=True/False`、`duplicateDims=True/False`、激活目标视图，三组组合实测第二视图尺寸数恒为 0。所以不再假装能做：`check_capabilities` 新增 `duplicate_dimension_to_view`（`unsupported`），workaround 指向 `list_drawing_edges` + `add_drawing_dimension(view_name=...)` 在该视图**重新标注**（实测视图2 尺寸 0 → 1，`value_m=0.11`）；`auto_dimension` docstring 同步警示。回归断言 62 → 63、能力表 21 → 22 条 |
| ~~气缸盖 P1-7~~ | ✅ 0.8.29：`set_view(break_alignment=True)` 先调 `IView.RemoveAlignment`（UI 的「解除对齐」）再写 `Position`，剖视不再被锁在父视图旁边。实测：不解除时请求 `y=parent+0.05` 被锁回 `0.2793`（`aligned_to_parent: true`），解除后落到 `0.3293`；回执带 `alignment_broken`。回归断言 61 → 62 |
| ~~气缸盖 P1-4 / P1-9~~ | ✅ 0.8.28：**P1-4** `set_title_block` 每字段只选**一个**单元格（GB 图框在框外还有一条 180° 倒置的「图样代号」，原实现会同时写两处，实测 code 出现在 `[0.053,0.285]` 与 `[0.391,0.022]`）——按「正的优先、再取离槽位最近」排序，跳过的写进 `skipped_extra_matches`；空单元格改为**新建 note**（`InsertNote` + `call_method(SetPosition2)`），解决 gb_a3.drwdot 返回 `planned: []`。**P1-9** `set_view` 改比例后调用 `_dodge_view_label`：把该视图所有 note 当标签障碍，超界的尺寸上移避让；顺带确认 `<VLLABEL>/<VLSCALEV>` **确实存在**（推翻 0.8.24 的「未复现」结论），并按可见宽度估算避免 40 mm 虚胖禁区。实测：停在标签 `[0.33,0.6371]` 的尺寸被移到 `[0.33,0.6497]`；回归断言 59 → 61 |
| ~~P0-1 尺寸摆位~~ | ✅ 0.8.27：`auto_dimension(place_outside=True)` 把本次插入的尺寸推到视图外**按侧分层**（`outside_gap_m` + n×`outside_step_m`），四侧就近分配、用 `precise_outline` 定位、只动本次插入、自动落盘。气缸盖实测：落在视图框内的尺寸 **21/25 → 0/25**，文本重叠 **17 对 → 6 对**（上 8 / 下 6 / 左 7 / 右 4）；回归断言 58 → 59 |
| ~~气缸盖反馈（尺寸布局链路）~~ | ✅ 0.8.26：**P0-4** `set_dimension_position` 增 `position_m`（一图 19 个 `D1` 可精确选，回执带 `disambiguated_by_position`）；**P0-2** `layout_dimensions` 把**其他**视图 outline 当固定障碍（同视图不挡，否则尺寸全被推出），并返回 `unresolved` / `unresolved_count` / `view_obstacles`；**P0-3** `set_view` 移动父视图时用 `GetBaseView` 检测派生视图并告警 `cut_lines_detached_warning`（切割线不会跟随）；**P1-1** `add_hole_callouts` 校验印刷直径与实测边径（`Ø99` 这类必拒）、被拒标注立即撤销、`verdict` 与失败数挂钩；**P1-2** `prune_orphan_dimensions` 增 `confirm` 门 + `caveat`（`GetAttachedEntities2` 返回空 ≠ 真孤立，实测清单里含 50 mm 总长）；**P1-3** `delivery_check` 带 `ledger_missing` 清单；**P1-5** `add_note` 返回 `bbox_final` / `shifted_m`；回归断言 51 → 58 |
| ~~第三方 v0.8.24 修复包~~ | ✅ 0.8.25：合入并逐项验证。**P0** `drawing.py` 五处 `_call_method(...)` 从未定义（只有别名 `_call_method_sf`）→ 每次删除抛 NameError 被 `except` 吞掉，`duplicates_removed` 从 0.8.19 起恒为 0；补模块级别名 + `dedupe_failed_count`。**N1** `delivery_check` 的 else 分支 `%d` 未格式化；**N2** `create_drawing` wrapper 补 `view_positions`；**N3** `auto_dimension` 新增 `dimension_filter` + `produce_drawing(prune_to_plan)`，让 plan 的 keep 8 真正落地（全不匹配则一条不删）；**N5** `place_views` 兼容 `view_name/x_m/y_m`；**P3-1** 拆 `missing_point` / `point_outside_views`；**P1-2** detail 视图创建前搜位、放不下返回 `no_space` 且不改图纸；新增通用防线 `assert_no_undefined_bare_calls`；回归断言 46 → 51 |
| ~~v0.8.14 反馈遗留五条~~ | ✅ 0.8.24：**P1-2** `list_circular_edges` 改用 `ICurve.CircleParams`（顶点路在轮廓边上恒为空，实测 0 → 23 条），`add_hole_callouts` 增 `verdict`（`no_native_callout` + 说明：边与相邻面都能选中仍返回 null，属几何限制）；**P2-4** 新增 `prune_orphan_dimensions`（按 name+position 删，实测 4 → 1，删后重跑校验）；**P3-1** `layout_dimensions` 把视图标签 note 当固定障碍（顺带修 `int(GetType())` 必抛导致类型过滤从未生效）；**P3-2** `add_note` 增 `avoid_collisions`；**P3-4** `set_title_block` 回执带单元格名与坐标；回归断言 39 → 45、工具 55 → 56 |
| ~~v0.8.14 反馈（标注删除 / 自检可信度）~~ | ✅ 0.8.23：**P0** `delete_annotation` 成功判定改为跨**全部视图**计数（删尺寸不改变图纸视图计数，旧的只看 `GetFirstView` 会把成功报成 `delete_failed`），四类标注实测可删、`position_m` 可区分 19 个同名 `D1`；**P1-1** 直径判定改用 `dimension_type == 5`（`GetText()` 无槽号对尺寸返回空，Ø 计数曾恒为 0），`list_drawing_annotations` 补 `dimension_type` / `raw_text` / 解析后 `text`；**P1-3** `pypdfium2` / `Pillow` 升为硬依赖且提示点名解释器；**P2-5** `model_clean` 由 blocker 降为 pending；**P3-5** reconcile 剔除 0.0 草图约束值；回归断言 34 → 39 |

### v0.9 待做

| 条目 | 为什么 | 难度 |
|---|---|---|
| **几何测距** | 最小折弯边（4t+R）、孔到折弯线（2t+R）需实体几何计算；当前 `not_checked` | 中高 |
| **尺寸精准筛选** | 规划说 8 个，实际导入 24 个；无按值筛接口，只能导后删 | 中 |
| **图面利用率** | 布局只保证不越界，视图偏小、留白多 | 中 |
| **形位公差符号** | 框格有、数值对，符号格空白 | 中（需找对写入方式） |
| **交付清单文件** | 自检结果还没落成可归档的一页纸 | 低 |
| 尺寸关联修复 | 8/29 读不到关联，需定性 | 中 |

排序原则：**先解决影响"能不能交付"的（几何测距、尺寸筛选、图面），再解决观感（符号）**。

## 5. 必须先补的工程能力

1. **批量处理框架**：读图必须逐张 `CloseDoc`（40+ 文档后全失败）；需要统一的批量入口，含进度与失败重试。
2. **超时与长任务**：单工具默认超时 300 s，批量走后台任务 + 状态轮询。
3. **COM 会话健壮性**：SW 崩溃/关闭后自动重连；文档堆积时主动清理。
4. **可观测性**：每个工具返回 `elapsed_ms`、`doc`、`warnings`；错误统一 `{ok:false, error, message, hint}`。
5. **无 COM 单测**：调色板校验、零件名归一化、JSON 结构、错误分支、统计分位数；真机回归单独跑。
6. **真机回归夹具**：监视器固定支架（46 组件）+ 配电箱（钣金）+ 100 张真实图纸抽样。
7. **基线可复现**：`survey_drawings` 固定 seed，改阈值前后必须能对比。

## 6. 质量门（每个工具合并前）

- [ ] 真机跑通：dry-run → apply → 复核 → 落盘验证
- [ ] 失败路径：文件不存在、只读、同名文档冲突、SW 未启动，均返回结构化错误
- [ ] 结果回读：不只信工具返回值（孔标注读孔径、尺寸读公差类型）
- [ ] 性能记录：写入 `docs/com-quirks.md` 与 README 实测表
- [ ] 中文路径/中文零件名可用
- [ ] 不修改原始测试文件：一律在副本上跑
- [ ] **判定类改动必须有真实数据支撑**（100 张基线或更严谨的样本）

## 7. 文档与传播

| 渠道 | 内容 | 频率 |
|---|---|---|
| `docs/com-quirks.md` | 每遇新坑即补，标注版本与机位 | 持续 |
| Release notes | 版本能力 + 实测数据 + 已知问题 | 每版 |
| README 首屏 | 一句话价值 + 交付判定示例 + 基线数据 | 每版更新 |
| MCP 目录/社区列表 | 提交服务器条目与简短示例 | v0.5 后一次 |
| 公众号「机械狗社区」 | 「AI 出的图能不能交」实操文，导流仓库 | 每大版本一篇 |

文章选题库：四个 COM 坑、100 张真实图纸告诉我们什么、为什么形位公差写不进去、
dry-run 的价值、SolidWorks 里的 AI 边界。

## 8. 风险与对策

| 风险 | 影响 | 对策 |
|---|---|---|
| **形位公差 API 缺失** | 永远无法宣称"完全合规" | 如实披露；真实图纸仅 7% 需要；调研 C# interop 兜底 |
| **孔标注性能** | 该环节自动化不可用 | 保留单点工具 + 界面手工；或用 Ø 尺寸替代 |
| COM 弹窗阻塞（保存/覆盖提示） | 进程挂死 | 写前静默保存；禁用交互式 API |
| 大装配/大批量性能 | 工具超时 | 缓存 + 后台任务 + 明确耗时预期 |
| 破坏性写操作 | 用户数据损坏 | 默认 dry-run、快照原值、只在副本测试 |
| 模型隐私 | 用户不愿上传 | 全本地运行，不联网、不上传 |
| 单人维护 | 迭代中断 | 每版保持小；工具模板化；坑清单即知识备份 |
| **判定标准失准** | 工具给出错误结论，比没有更危险 | 阈值必须来自实测基线；不确定的检测降级为 warning 而非 blocker |

## 9. 节奏

| 里程碑 | 内容 | 状态 |
|---|---|---|
| M1 | v0.1 上色闭环 | ✅ |
| M2 | v0.2 BOM / 导出 / 改名 | ✅ |
| M3 | v0.3 工程图链路 | ✅ |
| M4 | v0.4 台账 / 自检 / 判定 / 基线校准 | ✅ v0.4.5 |
| M5 | v0.5 孔标注 + 规划 + 检查单 | 待定 |
| M6 | v1.0 接口冻结 | 视 v0.5 反馈 |

## 10. 待决策

1. **形位公差要不要走 C# 插件**？官方两条 API 路径在 SW2024 都不存在，
   但 `IDrawingDoc` 的 QueryInterface 能拿到接口——需要在 C# 里实测 `InsertGtol` 是否真不可用。
2. **孔标注要不要放弃自动化**？方法已查清但扫边性能不达标；替代方案是用 Ø 尺寸标注。
3. **判定标准要不要按零件类型分档**？结构件（真实 7 张里 0 张有形位公差）与配合件要求完全不同。
4. 批量导出是否需要独立后台进程（避免阻塞 SW 主线程）？
5. 中文文档为主、英文 README 为辅，还是双语文档并列？
