"""Declare capability limits BEFORE the work starts.

The failure this prevents: `create_drawing` runs for 400 seconds on a hole-heavy sheet,
then `delivery_check` reports "0 hole callouts / 80 cylindrical faces". The caller spent the
time and got nothing. A preflight answer costs nothing.

Every entry states what is unsupported, why (with evidence), and what to do instead. This
is deliberately a hard-coded table rather than a probe: probing SolidWorks for a capability
it silently ignores is exactly how the 400-second wait happened.
"""

# status: "supported" | "partial" | "unsupported"
CAPABILITIES = [
    {
        "id": "view_layout",
        "status": "partial",
        "summary": "视图摆位与比例",
        "why": ("`create_drawing(view_positions=...)` 可按图纸米坐标显式摆位，`set_view` / `place_views` "
                "可事后移动缩放；自动布局在 4 视图以上会退回单列"),
        "workaround": "多视图请显式给 view_positions；auto_fit 不再覆盖调用方指定的 scale",
        "evidence": "2026-10-28 实测 4 视图 2x2：layout {fits: True, columns: 4, explicit: True}",
    },
    {
        "id": "outline_precision",
        "status": "verified",
        "summary": "视图 outline 恒定虚胖 11.9 mm，已用常数内缩统一修正",
        "why": ("`IView.GetOutline` 对每个视图的宽高各多报 11.9 mm（1:1 与 2:1 实测相同，"
                "是固定偏移而非投影误差）。所有判断——打包、重叠、出界、投影轴推断、"
                "点命中——统一走 `precise_outline()` 常数内缩，实测与真实投影吻合："
                "三视图 59.0×28.0 / 59.0×80.0 / 80.0×28.0（误差 <0.1 mm），"
                "detail 视图比虚胖值窄 11.92 mm。**没有逐顶点版本**："
                "`GetVisibleEntities2` 返回的是轮廓边（`GetStartVertex` 为 None）、"
                "`IBody2.GetVertices` 对全圆边实体返回 None、`ICurve` 在该版本没有 "
                "GetEndPoint/GetTessPoints，三条取点路均已实测排除。"),
        "workaround": ("需要原始值时仍读 `outline`，工具同时返回 `precise_outline_m`；"
                       "图纸尺寸（`ISheet`/`GetFirstView.GetOutline`）是精确的，"
                       "实测 420×297 mm，不受该偏移影响"),
        "evidence": ("2026-10-31 实测 灌装阀阀体 59×28×80：raw 70.88×39.88 → precise "
                     "58.98×27.98；drawing_audit 每行带 precise_width_m/precise_height_m；"
                     "detail 11.92 mm 差值"),
    },
    {
        "id": "bridge_resilience",
        "status": "verified",
        "summary": "COM 桥接自愈：存活探针 + 忙时有限重试 + 断线重连 + ping",
        "why": ("曾出现「连续约 50 分钟后所有调用报 Not connected，而 SolidWorks 本身正常」："
                "缓存的 dispatcher 变成死指针却永不刷新。现在 `get_sw()` 用 `RevisionNumber`"
                "（实测 18 ms）做存活探针，最多每 2 s 一次，失败立即丢缓存重新 GetActiveObject；"
                "`call_method()` 区分两类失败——忙（RPC_E_CALL_REJECTED）退避重试最多 2 次，"
                "断线（CO_E_OBJNOTCONNECTED 等）**不重试**而是丢缓存并抛 `BridgeLost` "
                "提示重试该工具；`ping()` 做真实往返，失败即重连重试一次并报 `recovered`。"
                "未处理的异常会 flush 到 stderr 并写入 logs/mcp-YYYYMMDD.log。"),
        "workaround": ("没有线程级硬超时：pywin32 在拥有对象的 STA 线程上派发，把调用丢到工作线程"
                       "再超时放弃会破坏 COM 公寓而不是取消调用。单次真正挂死仍会阻塞，"
                       "属 SolidWorks 侧状态，客户端无法安全中断"),
        "evidence": ("2026-10-31 实测：ping ok / 32.4.0 / 239.5 ms；污染缓存后 get_sw 自动丢弃"
                     "（stale_drops 0→1）并返回活连接；ping 报 recovered=true（reconnects 1→2）；"
                     "错误分类三例（busy / not connected / member not found）判定正确"),
    },
    {
        "id": "dimension_edit",
        "status": "partial",
        "summary": "尺寸删除、去重与单个移动",
        "why": ("`delete_dimensions` 按值保留/删除；`auto_dimension(eliminate_duplicates=True)` "
                "（0.8.19）按模型 full name 自动删除多视图重复副本，只动本次插入的尺寸、"
                "不碰用户已有尺寸；`set_dimension_position` 可把单个尺寸移到图纸坐标并回读数值。"
                "SolidWorks 无批量删除接口，实测 14 条一次调用 < 5 s（早期文档高估约 100 倍）；"
                "`prune_orphan_dimensions`（0.8.24）按 name + position 清掉挂不到实体的尺寸，"
                "`delete_dimensions` 按值删无法区分同值不同位置（一张阀体图有三个 23 mm）"),
        "workaround": "先用 plan_drawing 拿短名单再删；单条微调尺寸位置用 set_dimension_position",
        "evidence": ("2026-10-29 实测 14 条一次调用 < 5 s；2026-10-28 实测 17 条 → 2 条，"
                     "removed 15 / failed 0；0.8.19 full-name dedupe + 单尺寸移动 smoke 覆盖"),
    },
    {
        "id": "balloon",
        "status": "unsupported",
        "summary": "件号圆圈（BOM balloon）",
        "why": ("`add_balloon` 的两条路径都不成立。在 SW2024 SP4 实测："
                "`InsertBOMBalloon2` 全部参数都是必需的（DISP_E_PARAMNOTOPTIONAL，"
                "传 3/4/5 参都不够，完整签名需要 9 参含 Style/Size/上标/下标/引线样式）；"
                "`InsertBOMBalloon` 参数数目无效；`InsertBalloon` **名字解析不到**"
                "（DISP_E_UNKNOWNNAME，GetIDsOfNames 直接 False）。"
                "用户反馈在 SW2026 上 `InsertBalloon` 同样是 UNKNOWNNAME，与本地一致。"),
        "workaround": ("件号暂时用 add_note 写文字 + add_sheet_line 画引线；"
                       "或界面手工插（明细栏行号本身已经正确）"),
        "evidence": ("2026-10-29 实测 SW2024 SP4：InsertBOMBalloon2 → 非选择性的参数；"
                     "InsertBOMBalloon → 无效的参数数目；InsertBalloon → 未知名称。"
                     "用户反馈 SW2026 同款失败。"),
    },
    {
        "id": "delete_annotation",
        "status": "verified",
        "summary": "删除/移动标注（按 name；可用 position_m 区分重名）",
        "why": ("四类标注均已实测可删：note / 形位公差 / 表面粗糙度 / 尺寸，note 也已实测可移。"
                "关键是用 `IAnnotation.Select(False)` 选中——`Select3` 返回 False、"
                "`SelectByID2` 对「细节项目34」这类名字根本解析不到。"
                "**成功判定必须跨全部视图计数**：删除尺寸不改变图纸视图自身的注解数"
                "（实测图纸恒 83、工程图视图1 从 8→7、总数 122→121），只看 `GetFirstView` "
                "会把成功的删除误报成 `delete_failed`。"
                "SolidWorks 按草图复用 D1/D2/D3，一张图可能有 19 个同名 `D1`，"
                "这时传 `position_m` 指定具体那一个。"),
        "workaround": ("按 name 操作前先用 `list_drawing_annotations` 取名与 `position_m`——"
                       "名字是 `细节项目N` 形式，每次建图编号不同，不能硬编码；"
                       "重名尺寸务必带 `position_m`"),
        "evidence": ("2026-10-31 实测：note 移到 [0.045,0.15] 生效；GTOL 删除 125→124、"
                     "SFSYMBOL 删除、DIMENSION 删除 123→122 全部 removed=True；"
                     "重名 D1 用两组 position 分别定位到不同对象。"),
    },
    {
        "id": "dimension_display_text",
        "status": "verified",
        "summary": "尺寸的显示文本与类型（值槽是 GetText(1)，不是 GetText()）",
        "why": ("`IDisplayDimension.GetText()` **不带槽号返回空串**——尺寸的值与格式码在槽 1："
                "实测 Ø20 尺寸 `GetText(0)=''`、`GetText(1)='<MOD-DIAM>'`、"
                "`GetType=5`（直径；线性尺寸为 2）。0.8.23 起 `list_drawing_annotations` "
                "对尺寸返回 `dimension_type`、`raw_text`（原始槽 1）与解析后的 `text`"
                "（`<MOD-DIAM>`→`Ø`、`<DIM>`→数值）；`delivery_check` 的 Ø 判定改用 "
                "`dimension_type`，不再依赖显示文本。此前正因只读 `GetText()`，"
                "一张印着 Ø28/Ø23/Ø12/Ø18/Ø13 的图被判「0 处 Ø 尺寸」并降级 verdict。"),
        "workaround": ("判断是否直径看 `dimension_type == 5`；要比对文字用 `text`；"
                       "要核对原始编码用 `raw_text`"),
        "evidence": ("2026-10-31 实测：38 个尺寸里 18 个 `dimension_type==5`，解析后 text 为 Ø；"
                     "`delivery_check` 的 hole_callouts 由「0 处 Ø 尺寸」变为「17 处 Ø 尺寸」"),
    },
    {
        "id": "annotation_write",
        "status": "partial",
        "summary": "注记 / 图纸直线（件号见 balloon 条目）",
        "why": ("`add_note` / `add_sheet_line` 可用；图纸草图单位是**米**不是毫米。"
                "件号请单独看 `balloon` 条目——它是 unsupported，不要在这里当成可用"),
        "workaround": "传毫米会得到满屏放射斜线（坐标被当模型空间并乘视图比例）",
        "evidence": "2026-10-28 实测 add_note / add_sheet_line 均 ok；add_balloon 见 balloon 条目",
    },
    {
        "id": "bom_table",
        "status": "verified",
        "summary": "装配明细栏（原生 BOM）· 仅 SW2024 SP4 验证",
        "why": ("**仅 SW2024 SP4 验证**（2026-10-29）。SW2026 报告 `InsertBomTable4(Indented)` "
                "返回 None，两步法第一步就拿不到表——与本项「verified」的口径不符，"
                "故 status 保持 verified 但限定版本。"
                "**已解决**（2026-10-29 实测）。直接 InsertBomTable4(BomType=TopLevelOnly=2) 在本环境"
                "只出表头且无数量列，SetConfigurations 各种形式都不能修复。可行路径是两步："
                "先用 Indented(3) + 有效配置建表（能出 120 行 4 列，证明装配/模板/引用都正常），"
                "再把 BomFeature.TableType 切到 2——得到 9 行模块 + 一个重复数量列（5 列），"
                "隐藏最后一列即得 10 行 4 列目标形态。"),
        "workaround": ("用 add_bom_table；配置名传真实值（空则取视图的 ReferencedConfiguration）。"
                       "不要在该流程里调 SetConfigurations（会把表清回表头），"
                       "也不要用 SetColumnType 补 Quantity 列（SolidWorks 禁止普通列改数量列，返回 False）。"),
        "evidence": ("SW2024 SP4 实测 SAB1_0000_000.SLDASM：rows=10 cols=4 data_rows=9，"
                     "headers=['项目号','零件号','说明','数量']，保存重开保持。"
                     "SW2026 实测失败：InsertBomTable4(Indented) 返回 None。"),
    },
    {
        "id": "sheet_sketch_space",
        "status": "partial",
        "summary": "图纸草图坐标系",
        "why": "`add_sheet_line` 按米定位可用；细表格布线仍需多次调用",
        "workaround": "先算好每段线的米坐标，批量调用",
        "evidence": "2026-10-28 实测 ok（未做整表验证）",
    },
    {
        "id": "sheet_metal_qc",
        "status": "partial",
        "summary": "钣金工艺检查",
        "why": ("板厚 / 弯折数 / 止裂槽有无可判；`include_cylinders=True` 能给出圆柱面清单"
                "（按半径分组，孔与折弯圆角一并列出）。最小折弯边与孔到折弯线仍需判定"
                "「哪个圆柱面是孔、哪个是折弯圆角」——当前不可靠：半径差=板厚会把 Ø12/Ø14 "
                "同心孔误判为折弯（80 个圆柱面被配成 36 个折弯），面相邻性判据也无效"
                "（已知折弯的内外圆角同样不共享边），故这两项保持 not_checked"),
        "workaround": "not_checked 的项按 not_checked 处理，不要当作通过",
        "evidence": "2026-10-28 实测箱体 3 道弯折无止裂槽 → blocker",
    },
    {
        "id": "flat_pattern_calc",
        "status": "unsupported",
        "summary": "展开计算与校验",
        "why": ("展开由 SolidWorks 钣金引擎计算，工具只投影展平状态；K 因子与折弯表读不到"
                "（GetCustomBendAllowance 返回空）"),
        "workaround": "用 flat_pattern_report 拿 expected_note 的措辞，工艺参数由厂内折弯表定",
        "evidence": "2026-10-28 实测 computed_by_tool=false",
    },
    {
        "id": "hole_callout",
        "status": "pilot",
        "summary": "原生孔标注",
        "why": ("走「视图可见边 + Select2 + AddHoleCallout2」。SW2026 上**不弹模态对话框**"
                "（官方 Remarks 过时），单视图 13 条圆边 12 秒完成；"
                "可用 leader_x_m/leader_y_m（0.8.19）把引线文字钉在图纸米坐标内，"
                "返回值 echo 请求落点并逐条回读真实 position_m；"
                "但只有朝向观察者的孔边可标——同一孔的背面边返回 null"),
        "workaround": ("背面/被遮挡的孔如实记 failed 并由 Ø 尺寸补；"
                       "通风百叶等非孔圆不产生标注"),
        "evidence": ("2026-10-28 实测 door.SLDPRT：5/13 成功（8 X 18 ×2、Ø15 贯穿 ×3，文字完整、"
                     "引线自动引出），失败 8 个全为 Ø6 背面边（z=6.0）"),
    },
    {
        "id": "section_line_follows_view",
        "status": "unsupported",
        "summary": "移动父视图时让剖面切割线跟着走（0.8.32 已修正初始位置，跟随仍不支持）",
        "why": ("**API 层面无解**（2026-11-01 实测四组接口）。剖切线只有三种触达方式，都不对："
                "① `IDrSection.GetLineInfo` 返回**模型坐标**（±1.03 m，远超 A3 的 420 mm，"
                "不是图纸坐标），移动视图后数值不变；"
                "② `IDrSection.SetLineInfo(VARIANT 数组)` 确实可写（返回 True），"
                "但改的是**模型空间的剖切平面**——等于把剖面切到别的位置，不是让线跟随视图；"
                "③ `IView.GetSectionLineInfo2` 只读且同样是模型坐标；"
                "④ `IView.GetSketchCount` / `GetSketches` 在晚绑定下**不存在**，"
                "无法触达图纸上的那条线。"
                "**0.8.32 已修正初始位置**：原来 `half` 多除了一次 `scale`"
                "（`width` 取自 `GetOutline` 本就是图纸米，0.1669×0.62/0.1 = 1.0347 m，"
                "A3 只有 420 mm），且 `CreateLine` 用绝对坐标没加视图位置，线画在了 x=0。"
                "修好后线长 0.1035 m、x = 视图中心，实测 `GetLineInfo` 的 x 与视图一致、"
                "y 覆盖视图上下各约 80 mm。**但「跟随」仍做不到**：`SetLineInfo` 会把"
                "多段线重排成折返（实测把两段独立短线连成一条往返线）。"),
        "workaround": ("**先把视图摆到位，再建剖面**（`add_section_view` 是最后一步）。"
                       "已经移过就 `delete_view` 删掉剖面再重建——代价是该视图上的尺寸会丢，"
                       "所以务必先定死布局。每次改动后用 `export_drawing` + `render_drawing` "
                       "目视核对切割线是否压在父视图上"),
        "evidence": ("2026-11-01 履带轮 P05：移动父视图 +30 mm 前后各渲染一次，"
                     "两张图上切割线坐标一致（未跟随）；`GetLineInfo` 12 个值前后完全相同"),
    },
    {
        "id": "hole_callout_text_slots",
        "status": "partial",
        "summary": "孔标注文字：带深度的孔会连成 Ø1515",
        "why": ("`AddHoleCallout2` 产出的文字是 `<MOD-DIAM>` 格式码，必须用 `SetText` 改写，"
                "但**直径和深度在不同文字槽里**：只写直径时深度留在另一槽，Ø15 深15 渲染成 "
                "`Ø1515`（反馈里的 `Ø9 → Θ99` 就是同一件事）。三条路实测都不通——"
                "① 只写直径 → 深度粘连；② 写直径后清空其余槽 → **标注整条消失**；"
                "③ 完全不写 → 印出字面量 `<MOD-DIAM>`。"
                "通孔（无深度，如门板 Ø8）正常，只有带深度的孔受影响。"),
        "workaround": ("对带深度的孔改用 `add_drawing_dimension`（普通尺寸，文字可控），"
                       "或建完标注后人工改一次文字；`add_hole_callouts` 的回执里 "
                       "`text_source=SetText_depth_may_trail` 就是这条限制的标记"),
        "evidence": ("2026-11-01 门板 A3：Ø8×2 渲染正常（`Ø8`），Ø15×3 渲染成 `Ø1515`；"
                     "PDF 文本提取 + 渲染图双向确认"),
    },
    {
        "id": "duplicate_dimension_to_view",
        "status": "unsupported",
        "summary": "把同一条驱动尺寸补插到第二个视图（局部放大）",
        "why": ("**API 层面做不到**（2026-11-01 实测）。`InsertModelAnnotations3` 是文档级幂等："
                "同一模型尺寸一旦插入过，再次调用一律返回 None——`all_views=True/False`、"
                "`duplicateDims=True/False`、激活目标视图，全部组合都试过，第二个视图始终 0 条。"
                "具体实测：pass1 `all_views=True` 插 12 条（全在视图1）；"
                "pass2 激活视图2 + `all_views=False` + `duplicateDims=True` → 返回 None；"
                "pass3 `all_views=True` + `duplicateDims=True` → 也返回 None。"),
        "workaround": ("在目标视图**重新标注**，不要想复制：先 `list_drawing_edges(view_name=...)` "
                       "拿到该视图的边，再用 `add_drawing_dimension(view_name=..., edge_index=...)` "
                       "建一条新尺寸。它引用的是同一个模型几何，数值必然一致，只是多一条标注"),
        "evidence": ("2026-11-01 履带轮 P05：两视图实验，三组参数组合，第二视图尺寸数恒为 0；"
                     "`InsertModelAnnotations3` 返回值恒为 None"),
    },
    {
        "id": "hole_callout_batch",
        "status": "pilot",
        "summary": "批量孔标注（带 verdict：能否标注由几何决定）",
        "why": ("按孔去重（同圆心同径只标一次），单视图 25 秒预算，超时返回 budget_exceeded。"
                "0.8.24 起返回 `verdict`：`ok` / `no_native_callout` / `nothing_to_annotate`。"
                "`no_native_callout` 表示找到了孔但 `AddHoleCallout2` 全部返回 null——"
                "实测（阀体）边与相邻面都能选中、锚点就在圆上仍为 null，而同一调用在"
                "钣金拉伸孔上可用，属几何/构建限制，换 offset 或换视图不会改变，"
                "`native_callout_note` 会把这一点写清楚。"),
        "workaround": ("看到 `no_native_callout` 就不要再调参数重试：改用 Ø 尺寸或手动补标；"
                       "`list_circular_edges`（0.8.24 修好，改用 CircleParams 而非顶点）"
                       "可先列出圆边判断值不值得试"),
        "evidence": ("2026-10-28 实测钣金门板 Ø8/Ø15 成功；2026-10-31 实测阀体 8/8 全 null"
                     "（边/面均选中），`list_circular_edges` 由 0 条修正为 23 条"),
    },
    {
        "id": "geometric_tolerance",
        "status": "partial",
        "summary": "形位公差框格（符号为文字）",
        "why": ("**可以做**（2026-10-28 突破）。此前判定「不可创建」是调用方式错误："
                "pywin32 把 InsertGtol 当作属性派发（PROPERTYGET），必然报 "
                "DISP_E_MEMBERNOTFOUND，改用 DISPATCH_METHOD 即可拿到 Gtol 对象，"
                "SetFrameValues2 写入公差与基准并回读确认。"
                "0.8.20：符号格改用**可读文字**（位置度/平面度…）写入第一格，"
                "再用 GetFrameSymbols2 回读校验；`<MOD-*-PAR>` 格式码实测只会被存成"
                "字面文本、不解析成几何符号，故已弃用。写入不生效会删除框格并返回 "
                "`gtol_symbol_not_applied`，不再谎报符号已标注。"),
        "workaround": ("符号是文字不是真正的 GD&T 几何符号，出正式图纸请在界面替换；"
                       "或用 symbol='none' 只写公差值与基准"),
        "evidence": ("2026-10-28 实测 door.SLDPRT：2 个框格 frame1 回读正确；"
                     "2026-10-31 实测 GetFrameSymbols2 回读 ('位置度','','','','','')，"
                     "SetFrameSymbols2 返回 None 必须以回读判定"),
    },
    {
        "id": "surface_finish",
        "status": "partial",
        "summary": "表面粗糙度符号（位置修正 + 像素可见性自检）",
        "why": ("`InsertSurfaceFinishSymbol3` 14 参可创建符号，但**忽略传入的 x/y**——"
                "符号实际落在图纸原点 (0,0)，早期只回显请求值 position_m 造成谎报。"
                "0.8.20 起创建后用 `IAnnotation.SetPosition` 移到请求点并回读 "
                "`position_actual_m`；`verify_render=True` 前后各导出 PDF 渲染 PNG，"
                "在 (x,y) 窗口做像素差分证明符号确实可见（`visibility.visible`）。"
                "`verify_dwg=True` 只检查 DWG 文件非空——矢量格式无像素通道。"),
        "workaround": ("要看符号是否真在图面上，传 verify_render=True；"
                       "像素检查依赖 pypdfium2 + Pillow，缺失时 visibility.checked=false 诚实降级"),
        "evidence": ("2026-10-31 实测 door.SLDPRT：未修位置前 PDF diff 只在原点 "
                     "(0,1626)-(34,1684)；SetPosition 后 diff bbox=(1684,493)-(1736,551) "
                     "落在目标像素 (1701,549)，窗口 changed=265"),
    },
    {
        "id": "detail_view_scale",
        "status": "partial",
        "summary": "局部放大视图：空视图与比例锁定诊断",
        "why": ("局部视图默认 UseParentScale，创建时设置 ScaleDecimal 会被吞掉。"
                "0.8.20 起 `scale` 不再被父视图比例覆盖，创建后重试设置并回读；"
                "仍被锁则返回 `detail_scale_locked` error 而非静默；`edge_count==0` "
                "返回 `detail_empty`（边界圈落在源视图几何之外，视图是空的）。"),
        "workaround": ("`center` 必须落在源视图 outline 内的真实几何上；返回 detail_scale_locked "
                       "时保存重开后在 UI 设比例，或让 scale 等于父视图比例"),
        "evidence": ("2026-10-31 实测 door.SLDPRT：边界在内 edges=1、在外 edges=0；"
                     "`_apply_detail_scale(2.0)` 生效，视图名变为 2 : 1"),
    },
    {
        "id": "section_view",
        "status": "partial",
        "summary": "剖视图（模型坐标剖切面 + 剖切 sanity 校验）",
        "why": ("0.8.21 起 `cut_plane_normal` / `cut_plane_origin` 可用**模型坐标**指定剖切面；"
                "内部按「视图 deflated outline 与零件包围盒的尺寸匹配」确定图纸轴 ↔ 模型轴"
                "（`GetViewPlaneNormal` 不存在、`ModelToViewTransform` 的 16 个数不是行主序 "
                "4x4、`IMathPoint` 没有 MultiplyTransform，三者均已实测排除）。"
                "法向必须轴对齐（主分量 ≥0.9），斜切 / 平行图面 / 映射歧义分别返回 "
                "`cut_plane_oblique` / `cut_plane_not_in_view_plane` / `cut_plane_unmapped`。"
                "`verify_cut=True` 默认开启：垂直于图纸水平轴的剖切不改变宽度，故剖视宽必须"
                "等于源视图宽（竖切比高），容差 10%；超出返回 `section_suspicious`，"
                "`edge_count==0` 返回 `section_empty`，两者都保留视图供检查。"),
        "workaround": ("映射靠尺寸匹配，正立方体这类各向等长的零件会 `axes_ambiguous`，"
                       "此时改传 orientation + offset_m；部分剖（depth>0）尺寸本就缩小，"
                       "sanity 记 checked=false 不判定"),
        "evidence": ("2026-10-31 实测 灌装阀阀体 59×28×80：法向 +Y 过中心 → 轴映射 [0,1]、"
                     "horizontal、剖视 58.98×79.98、ratio 1.000；已知错位 offset +0.012 → "
                     "截面 14.4×49.98、ratio 0.244，被 `section_suspicious` 拦下；"
                     "法向 +X → vertical、ratio 1.04；法向平行视线 / 斜切均结构化拒绝"),
    },
    {
        "id": "assembly_bounding_box",
        "status": "partial",
        "summary": "装配体包围盒",
        "why": "IAssemblyDoc 没有 GetPartBox；组件 GetBox 在轻化状态下返回 None",
        "workaround": "比例由实测安全区推算，建图后测量越界再降档",
        "evidence": "2026-10-28 实测，1.6m 条码设备靠该方法得到 1:20",
    },
    {
        "id": "dxf_import",
        "status": "unsupported",
        "summary": "DXF/DWG 逆向建模",
        "why": "未实现，且与建模类工具重叠",
        "workaround": "用 SolidPilot 的 analyze_drawing",
        "evidence": "设计决策",
    },
    {
        "id": "geometry_creation",
        "status": "unsupported",
        "summary": "建模/特征创建",
        "why": "本工具定位是出图与自检，不做几何创建",
        "workaround": "用建模类工具（如 solidpilot）",
        "evidence": "设计决策",
    },
]

STATUS_LABEL = {"supported": "可做", "partial": "有限支持", "unsupported": "做不了"}


def check_capabilities(operation=None):
    """Return what this toolkit can and cannot do, optionally filtered to one operation.

    Call this before a long pipeline. It answers in milliseconds what would otherwise cost
    a full SolidWorks run to discover.
    """
    from ..sw_session import error, ok, runtime_info, sw_version

    entries = CAPABILITIES
    if operation:
        key = str(operation).strip().lower()
        entries = [c for c in CAPABILITIES
                   if key in c["id"].lower() or key in c["summary"]]
        if not entries:
            return error("unknown_operation",
                         "no capability entry matches %r" % operation,
                         "Known ids: %s" % ", ".join(c["id"] for c in CAPABILITIES))

    blocked = [c for c in entries if c["status"] == "unsupported"]
    partial = [c for c in entries if c["status"] == "partial"]
    payload = []
    for item in entries:
        row = dict(item)
        row["status_label"] = STATUS_LABEL.get(item["status"], item["status"])
        payload.append(row)

    # Version context belongs in the answer: "verified" means verified on one SolidWorks
    # build, and bom_table is SW2024 SP4 only. The PATH matters just as much — several copies
    # of this package can coexist on one machine, and an editable install pinned to an old
    # checkout is exactly what once sent a report after fixes that were already in the repo.
    runtime = runtime_info()
    return ok(
        server_name="Solidworks_AssetMate_MechDog",
        server_version=runtime.get("server_version"),
        server_path=runtime.get("server_path"),
        python_executable=runtime.get("python_executable"),
        python_version=runtime.get("python_version"),
        pid=runtime.get("pid"),
        sw_version=sw_version(),
        platform=__import__("platform").platform(),
        operation=operation,
        checked=len(entries),
        blocked=[c["id"] for c in blocked],
        partial=[c["id"] for c in partial],
        capabilities=payload,
        advice=("有 %d 项做不了。调用前请先看 workaround，不要跑完流程才发现。"
                % len(blocked)) if blocked else "全部可做。",
    )
