# SolidWorks COM 踩坑清单（pywin32 实测，SW 2024 / 32.4.0）

> 来源：2026-10-26 给风力仿生兽装配体按件配色时的真实调试过程。
> 标 ✅ 的是验证过能用的写法，❌ 的是"不报错但没效果 / 报类型不匹配"的写法。

## 一、外观（颜色）

### 外观有四层，别改错层

| 层 | 设置方式 | 影响范围 |
|---|---|---|
| 材质外观 | `set_part_material` / 材质库 | 零件 + 密度（会改质量属性） |
| 文档默认外观 | `doc.MaterialPropertyValues` | 该文档全部实体 |
| **组件外观** | **`comp.MaterialPropertyValues`** | **该实例（装配里按件上色就用这层）** |
| 面外观 | `Extension.SetMaterialPropertyValues`（选中面后） | 选中项 |

```python
# ✅ 逐组件
comp.MaterialPropertyValues = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [r, g, b, 1.0, 1.0, 0.88, 0.21, 0.0, 0.21])

# ❌ 文档级：会把装配里所有组件刷成同一色，且不报错
doc.Extension.SetMaterialPropertyValues(color_variant, 1)
```

### 9 个分量

```
[R, G, B, ambient, diffuse, specular, shininess, transparency, emission]
```
后 6 个**沿用默认** `[1, 1, 0.88, 0.21, 0, 0.21]`。把 ambient 降到 0.5 会让正红渲染成 `(72, 0, 0)` 的暗红——看起来像"没生效"。

### 数组必须封成 VARIANT

```python
comp.MaterialPropertyValues = [1, 0, 0, 1, 1, .88, .21, 0, 0]      # ❌ 静默失败
comp.MaterialPropertyValues = VARIANT(VT_ARRAY | VT_R8, [...])      # ✅
```

### 读回值 ≠ 文件里的值

装配引用的零件文档常驻内存，`CloseDoc` 后 `OpenDoc6` 会复用内存副本。判断"存没存进去"必须先清干净：

```python
sw.CloseAllDocuments(True)   # 再 OpenDoc6 才是从磁盘读
```

这一条害我误判了 20 分钟：文件明明是蓝的，读出来一直是白的。

## 二、打开文档

```python
# ❌ 参数不够 → "非选择性的参数"；Errors/Warnings 是 [out]
sw.OpenDoc6(path, 1, 1, "")

# ✅ byref VARIANT 占位
e = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
w = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
doc = sw.OpenDoc6(path, 1, 1, "", e, w)      # 1=part, 2=assembly；Options 1 = silent
```

注意：**SolidWorks 不允许同时打开两个同名文档**。若 `OutTXT\C Linkage.SLDPRT` 已开，再开 `D:\src\C Linkage.SLDPRT` 会返回 `None`——先把副本改名再开。

## 三、选择集

```python
comp.Select2(False, 0)          # ✅ 选组件
comp.Select4(False, None, False)  # ❌ 类型不匹配（Data 是 byref 指针）
ext.SelectByID2(name, "COMPONENT", 0,0,0, False, 0, None, 0)  # ❌ 末位参数 byref 报错
```

要改组件属性时，**不要走选择集**，直接拿 `IComponent2` 对象设属性最稳。

## 四、性能

- `asm.GetComponents(False)` 在 80 组件装配上要**数秒**；2026-10-26 在 46 组件的监视器支架上实测单次 **约 130 秒**（含首次打开零件的加载）。放进循环（每组调一次）会把会话拖到工具超时被杀。**只调一次，在内存里分组**。
- 本项目的做法：按文档路径缓存组件列表，`tree → dry-run → apply → tree` 只付一次 `GetComponents` 成本；用户外部增删组件后传 `refresh=true`。
- `SaveBMP(path, w, h)` 在 1400×1000 下约 5–10 s。
- 重开 80 组件装配 1–3 分钟，验证要挑时机。
- **关闭文档前必须先保存**。对未保存文档调 `CloseDoc` 会弹模态保存提示，自动化进程会被永久阻塞（无 UI 可点）。判断用 `GetSaveFlag`，保存用裸访问 `doc.SaveSilent`。
- **缓存不能只靠"读属性会抛异常"来判失效**。文档被外部关闭重开后，陈旧的 `IComponent2` 指针**不抛异常**，而是静默返回空串（`Name2=""`、`GetPathName=""`、`GetModelDoc2=None`）。此时任何依赖组件列表的读操作都会返回"成功但全空"的数据。校验必须同时满足：
  1. 缓存的 COM 文档身份与当前一致（`cached_ole == doc._oleobj_`，实测同对象返回 `True`、重开后为 `False`）；
  2. 首个组件的 `Name2` 非空。
  只做其中一条都会漏。

## 五、几个不可用的成员（本机实测）

这些在 pywin32 动态绑定下报"找不到成员"，不要浪费时间：

| 成员 | 现象 | 替代 |
|---|---|---|
| `SldWorks.GetMathUtility()` | 找不到成员 | 用组件现有 `Transform2` 的 `ArrayData` |
| `doc.Extension.CreateMassProperty` | 找不到成员 | 逐件用 MCP `analyze_model` 读质量属性，或手算合并 |
| `doc.SaveAs4(...)` | 非选择性的参数 | 用 MCP 的 save_document，或 `doc.SaveSilent` |
| `doc.Extension.GetMassProperties2(1, False)` | 类型不匹配 | 同上 |
| `SwApp.GetDocuments()` | 有时返回 `None` | 逐次 `OpenDoc6`，或先 `CloseAllDocuments` |

## 六、图像导出

```python
doc.SaveBMP(path, w, h)   # ✅ 但无论扩展名写 .png 还是 .bmp，实际都写 BMP（文件头 "BM"）
```

要 PNG：改名成 .bmp 再用图像库转，或用纯 Python 读 BMP + `zlib` 手写 PNG（无 Pillow 环境下的可行路径，约 30 行）。

## 七、保存文档（2026-10-26 实测新增）

`SaveSilent` 与 `GetSaveFlag` 在 pywin32 动态绑定下都是**属性**，不是可调用方法：

```python
doc.SaveSilent          # ✅ 访问即执行保存，返回 0/1（返回值不代表成功）
doc.SaveSilent()        # ❌ TypeError: 'int' object is not callable
doc.GetSaveFlag         # ✅ bool：True = 有未保存修改
doc.GetSaveFlag()       # ❌ TypeError: 'bool' object is not callable
```

返回值不可靠：实测保存成功后 `SaveSilent` 仍返回 `0`，应以 `GetSaveFlag` 在保存后变为 `False` 作为判据。

**重开后组件顺序会变。** 验证"颜色是否落盘"不能比对 `GetComponents()[0]`，必须按组件名（如 `P19-3`）查找同一实例。只关闭当前文档用 `sw.CloseDoc(title)`，不要为了验证而 `CloseAllDocuments`——会关掉用户正在编辑的其他文档。

## 八、实体级外观（2026-10-26 实测新增）

多实体零件按实体分色，官方文档推荐的两种写法在本机都不可靠：

| 写法 | 结果 |
|---|---|
| `body.MaterialPropertyValues = VARIANT(...)` | ❌ 不报错，回读永远是 `None`，颜色不变 |
| 选中实体 + `Extension.SetMaterialPropertyValues(color, 1)` | ⚠️ 小实体生效，758 面的主壳无效 |
| **逐面 `face.MaterialPropertyValues = VARIANT(...)`** | ✅ 全部生效，可回读，可靠 |

选择实体本身也有坑：`IBody2.Select2` 在晚绑定下报 `Python instance can not be converted to a COM object`，`SelectByID2` 的 Callout 参数必须传 `VARIANT(VT_BYREF | VT_DISPATCH, None)`：

```python
callout = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_DISPATCH, None)
doc.Extension.SelectByID2(name, "SOLIDBODY", 0, 0, 0, False, 0, callout, 0)
```

**结论：实体分色走面级。** 代价与注意点：

- 758 面约 12 秒；899 面的零件整体 apply 约 76 秒（含重建与保存）。
- 面级外观优先级最高，会覆盖原有的面外观/贴图，操作前务必用副本。
- `GetBodies2` 返回的实体**顺序不稳定**，跨调用必须按实体名匹配，不能用索引。

## 九、BOM 与重命名（2026-10-26 实测新增）

### 轻化组件读不到质量

装配默认轻化加载时，`comp.GetModelDoc2` 返回 `None`，`CreateMassProperty` 无从调用，19 个零件只能读到 1 个的质量。先一次性解析：

```python
doc.ResolveAllLightweightComponents(True)   # 返回值可能是 0，但文档确实加载了
```

之后 46/46 组件全部有 `GetModelDoc2`，总质量 0.126567 kg，与整机 `GetMassProperties` 一致。`CreateMassProperty` 只对**已加载**的文档有效，对未加载文档不会自动加载。

材料用 `GetMaterialPropertyName2("", VARIANT(VT_BYREF | VT_BSTR, ""))` 读取；未指定材料时返回**空字符串**，不是 `None`——判断时要 `strip()`。

### 重命名：实例名改不了，文件名可以

| 目标 | 结果 |
|---|---|
| `comp.Name = "x"` | ❌ `Property 'Name' can not be set` |
| `comp.Name2 = "x"` | ❌ 不报错，读回仍是旧名，`GetSaveFlag` 不变 |
| `doc.Extension.RenameDocument("x.SLDPRT")` | ❌ 返回 `2`，文件未改名（非 PDM 工作区） |
| **文件系统 `os.rename`** | ✅ 可用，但打开中的文件会失败 |

`IAssemblyDoc` 没有 `RenameComponent` / `RenameComponent2` 成员（用 `_oleobj_.GetIDsOfNames` 可直接探测成员是否存在，比 `dir()` 可靠，`dir()` 在 SldWorks 对象上会抛"找不到元素"）。

**结论：实例名只能 UI 改；工具用「另存 + 重定向」实现文件级重命名，并明确返回 `instance_rename_unsupported`，不造假。** `mode="file"` 会断装配引用，`mode="auto"` 会同步重定向引用。

### 装配打开时改不了零件文件名

装配会锁住组件文件，`os.rename` 直接报 `WinError 32 另一个程序正在使用此文件`。而"先关装配再改名"走不通——改名后装配找不到旧文件，直接断链。

可行路径是**另存 + 重定向**：

```python
part = open_document(old_path, SW_DOC_PART)
part.SaveAsSilent(new_path, False)      # ✅ 静默生成新文件（返回 0，文件已存在）
comp.ReplaceReference(new_path)         # ✅ 返回 1，单个组件立即改指向
doc.SaveSilent                          # 保存装配
```

实测：

- `SaveAsSilent` 会顺带更新已打开装配里指向该文档的引用；`ReplaceReference` 用于显式补刀，多实例要逐个调用。
- 保存并重开后路径与实例名都已更新，实例名按新文件名重排（`P01-1` → `RN_P01-1`）。
- 原文件保留，回退只需把引用改回去并删掉新文件。
- 关闭被装配引用的零件文档再改名无效（文件锁不释放），不要在这条路上浪费时间。

## 十、早绑定 gen_py 缓存让 byref VARIANT 全面崩溃（2026-10-27 用户实测）

现象：同一份代码，某些机器上 7 个工具全部失败：

```
TypeError: int() argument must be a string, a bytes-like object or a real number, not 'VARIANT'
```

根因：`%TEMP%\gen_py` 里存在 SldWorks 类型库的 makepy 缓存时，`GetActiveObject` 返回**早绑定** wrapper。它把 `OpenDoc6` 的 `Errors`/`Warnings` 声明为 byref int，在 `InvokeTypes` 里对传入的 `VARIANT` 调 `int()` → 直接崩。所有工具都要经过 `open_document`，所以全灭；只有纯文件系统的路径（`rename_components` 的 file 模式）幸存。

修复——拿到对象后强制晚绑定：

```python
import win32com.client.dynamic as dyn
_SW = dyn.Dispatch(win.GetActiveObject("SldWorks.Application"))
```

晚绑定路径下 `VARIANT(VT_BYREF | VT_I4)` 正常工作，且不依赖 gen_py 是否被别的程序生成过。

要点：

- **不要靠删 `gen_py` 缓解**——其他程序（含 makepy、其他 CAD 脚本）会再生成，问题必然复发。
- `EnsureDispatch` 对 SldWorks 会报 "This COM object can not automate the makepy process"，所以也没法用早绑定统一解决。
- 同类风险适用于所有自定义 `byref VARIANT` 调用（`SelectByID2` 的 Callout、`GetMaterialPropertyName2` 的 BSTR）。统一走晚绑定最省事。

## 十一、组件颜色读回失真（2026-10-27 用户实测）

在某些机器上，`recolor_assembly` 写入蓝/黄后：

- 渲染正确（预览图取色 `(24,64,195)` / `(251,181,5)`）
- 保存正确（关闭重开、`ResolveAllLightweightComponents` 后渲染依旧）
- 但 `comp.MaterialPropertyValues` 读回 `[1.0, 1.0, 1.0]`

同一台机器上面级 `IFace2.MaterialPropertyValues` 读回正常。出现条件与存在 gen_py 缓存、晚绑定包装早绑定对象一致，判断为**读路径的 VARIANT 解组问题，不是写入问题**。

对策：

- `recolor_assembly` 会**按 palette 每个零件键各抽 1 个实例**（最多 8 个）做读回校验，返回 `readback_sample` 与 `readback_verified`；为 `false` 时附 warning。
- 失真有两种形态：读回恒白 `[1, 1, 1]`，以及读回 `None`。**`None` 不能当作"该组件没有颜色"**，也不代表未设置外观。
- 失真按组件/会话随机分布，不是全文档属性，所以抽样必须覆盖每个零件类型，不能只取前 3 个组件。
- 此时**不要用 `previous_colors` 回退**——拿到可能是白色或 `None`。验收请用 `export_preview` 看图。
- `get_assembly_tree` 的 `rgb` 在这种环境下不可信，只看结构、路径、实例数。

另：`recolor_bodies` 的面级颜色在装配中会被**组件级颜色覆盖**（视口优先级：组件级 > 面级）。零件文件内颜色正确，但装配里看到的是组件色；工具已在 `note` 中提示。

## 十二、导出 STEP（2026-10-27 实测）

`doc.SaveAsSilent("x.STEP", False)` 可以直接写出有效的 STEP，而且**不会把文档路径改成 .STEP**——它按导出处理，不是真正的另存为：

- 46 组件装配 → 1.1 MB，约 2 秒
- `doc.GetPathName` 导出前后不变，可以连续导出、无需重开原文件
- 目标文件已存在时工具直接报冲突，不覆盖

PDF 必须来自 `.SLDDRW` 工程图，DXF 一般来自钣金平板型式或工程图视图；都不能从任意零件导出。`batch_export` 只接受 `step`，其余格式明确拒绝并给提示，不猜。

## 十三、工程图（2026-10-27 实测）

这一节是 4 个工程图工具的来源，每条都实机踩过。

### 模板路径必须是反斜杠

```python
sw.NewDocument(r"C:/.../gb_a3.drwdot", 0, 0, 0)   # ❌ 静默返回 None
sw.NewDocument(r"C://...//gb_a3.drwdot", 0, 0, 0)  # ✅
```

正斜杠不报错、不抛异常，只返回 None——是最容易误判成"模板坏了"的坑。模板本身可从
`C://ProgramData//SOLIDWORKS//<版本>//templates//*.drwdot` 找（GB 模板通常叫 `gb_a3.drwdot`）。

### 命名视图是本地化的

中文版 SolidWorks 的命名视图是 `*前视`、`*上视`、`*等轴测`，传 `*Front` 同样静默失败：

```python
names = list(model.GetModelViewNames)   # ('*正视于','*前视',...,'*等轴测',...)
```

正确做法是从模型读出真实名称再匹配，不要硬编码英文常量。

### 一批方法在晚绑定下是属性

| 写成方法 | 实际 |
|---|---|
| `doc.GetFirstView()` | `doc.GetFirstView`（属性，返回 sheet 的 IView） |
| `doc.GetViews()` | `doc.GetViews`（属性，返回**嵌套 tuple**） |
| `view.GetOutline()` | `view.GetOutline`（属性，返回 4 元素数组） |
| `view.GetDimensionCount()` | `view.GetDimensionCount`（属性） |

`GetViews` 的结构是 `((PyIDispatch, PyIDispatch, ...),)`，元素是裸指针，必须
`win32com.client.Dispatch(v)` 包装后才能读 `Name` / `Type` / `ScaleDecimal`。

### 视图插入前必须先开模型

`CreateDrawViewFromModelView3` 需要模型已加载；并且模型**必须先存盘**。工程图文档本身
可以在未保存状态下加视图，但模型不行。

### 新建视图读不到 outline

刚建好的视图 `GetOutline` 可能为空，设置 `ScaleDecimal` 之后才有值。因此比例不能靠"先建后量"，
要用 `model.GetPartBox(True)` 先估包围盒、选标准比例（1:1、1:2、1:5…），再一次性建视图并排版。

### GetOutline 恒定虚胖 11.9 mm（2026-10-29 实测，P2-13）

`IView.GetOutline` 返回的包围盒**宽和高各比真实投影大 11.9 mm**，与视图、比例无关：

| 视图 | 真实投影 | GetOutline | 差 |
|---|---|---|---|
| front (X×Y) | 59 × 28 | 70.9 × 39.9 | +11.9 / +11.9 |
| top (X×Z) | 59 × 80 | 70.9 × 91.9 | +11.9 / +11.9 |
| right (Z×Y) | 80 × 28 | 91.9 × 39.9 | +11.9 / +11.9 |
| A-A 剖视 (Y×Z) | 28 × 80 | 39.9 × 91.9 | +11.9 / +11.9 |

2:1 后仍是同一常数，说明是**固定偏移而非投影误差**。后果：`drawing_audit` 会把只在这层
虚胖边距里"相交"的两个视图误报 `views_overlap`，`_place_views` 摆位偏保守，auto_fit 会把
本可放下的视图缩小。

修法（`drawing.precise_outline`）：宽高各减 `OUTLINE_MARGIN_M = 0.0119`、四边各内缩 5.95 mm，
O(1)、零额外 COM 调用，与真实 ink bbox 吻合。**不要**在 audit 里逐视图遍历可见边顶点重算：
`ink_outline()` 那样做在 25 特征零件上实测约 200 s（见「四、性能」），只适合一次性交叉校验。

另外 `ink_outline` 早期版本返回的是**模型坐标**的 x/y，对 top/right 视图用错了投影轴；现在
统一用 `paper = view.Position + (model - centre) * scale` 投影到图纸米，与 `GetOutline`
同一坐标系后才能直接比较。

### 尺寸标注的注意事项

- 图纸坐标单位是米，A3 = 0.420 × 0.297。
- 标注必须点在投影边（`EDGE`）上，选不中时依次退化到 `SILHOUETTE`、`SKETCHSEGMENT`。
- 调用 `AddDimension2` 前要临时关闭 `swInputDimValOnCreate`，否则弹出数值输入框会卡死自动化。
- 剖视图必须先 `ActivateView` 目标视图，再选/画剖切线；否则 `CreateSectionViewAt5` 返回 null，
  At2 兜底甚至会让 SolidWorks 崩溃（RPC_E_SERVERFAULT）。

### 自动标注的三个 API（2026-10-27 实测）

```python
doc.InsertModelAnnotations3(0, 32768, True, True, False, True)   # 插入模型尺寸
view.AutoInsertCenterMarks2(1|4, 2, True, True, True, 0.0025, 0.0, True, True, 0.0)
doc.AddHoleCallout2(x, y, 0)                                     # 孔标注
```

- `InsertModelAnnotations3`：Option=0（FromEntireModel），Types 位掩码 32768（标记为工程图的尺寸）
  | 524288（全部驱动尺寸）；返回数组，长度即插入数量。P01 支架实测插入 6 个。
- `AutoInsertCenterMarks2` 是 **10 个参数**：InsertType(1=孔|4=槽)、Style(2=Single)、
  LinearSlotCenter、ArcSlotCenter、UseDocumentDefaults(True)、Size、Gap、ExtendedLines、
  CenterLineFont、Angle。**必须先 `ActivateView`**，否则不生效。返回 True 的视图数才是可靠信号，
  `GetCenterMarkCount` 在这个绑定下会低估（常读 0，但图上确实有标记）。
- `AddHoleCallout2` 必须点在孔的投影圆上，否则返回 None——工具据此返回 `callout_failed`，
  不伪造标注。
- `IModelDoc2.EditRebuild3` 是**属性**（bool），写 `EditRebuild3()` 会报 `'bool' object is not callable`。

### GB 模板的标题栏是写死的占位文字

新建的 GB 图纸标题栏显示 `"图样名称"`、`"图样代号"`（连引号一起），材料是 `<未指定>`。
这些**不是属性链接**，就是普通的 Note 文本，所以每张图都会原样显示占位词。

替换路径：

```python
sheet = doc.GetFirstView           # 属性
for ann in sheet.GetAnnotations:   # 属性，嵌套 tuple，需 Dispatch 包装
    if ann.GetType == 6:           # swAnnotationType_Note
        note = ann.GetSpecificAnnotation
        text = note.GetText        # 属性
        note.SetText("履带轮")      # 方法，返回 True
```

有的模板连占位词都没有——`gb_a3.drwdot` 的图样名称/代号格是**空文本**，文字匹配必然落空。
`set_title_block` 因此带位置兜底：文字匹配失败时按 GB A3 的格子坐标（`TITLE_SLOTS`）找最近的
注释再写入，返回里用 `via: "position"` 标出来。实测配电箱箱体：`name`/`code` 两格均按位置写入成功。

实测：一张 A3 图里 `图样代号` 出现 **2 次**：

| 位置 (m) | 角度 | 含义 |
|---|---|---|
| (0.391, 0.022) | 0° | 标题栏内的图样代号 |
| (0.053, 0.285) | **180°** | 图框左上角的倒置图号 |

左上角那处是 GB 折叠惯例：图纸对折后从背面能看到图号，所以模板里就是倒排的。
替换文字不会改变角度，因此它显示为倒着的编号——**这是标准，不是写错**。
`set_title_block(include_reversed=False)` 可以只改标题栏、保留倒置标记原样。

### 补标：按边索引选，不要指望坐标

`InsertModelAnnotations3` 只能搬模型已有的驱动尺寸。IR 重建的零件往往尺寸不全
（实测 P05：模型 38 个驱动尺寸，工程图只能拿到 27 个，缺的 7 个是基准面偏移、
7 个是草图内部尺寸，这两类不随模型项目导入），必须补标。

在工程图里补一个尺寸的可行顺序：

```python
# 1) 列边，拿索引
edges = view.GetVisibleEntities2(view, 1)        # 147 条边
edge  = win32com.client.Dispatch(edges[n])
# 2) 选边（Select2 可用，Select4 的 callout 参数会被拒）
edge.Select2(False, 0)
# 3) 加尺寸，坐标是图纸米
doc.AddDimension2(x, y - 0.015, 0).GetDimension2(0).SystemValue = 0.11
```

要点：

- **`edge.GetStartVertex.GetPoint` 返回的是模型坐标**，不是图纸坐标（P05 实测 x 为 -0.86），
  所以「按图纸坐标就近选边」是行不通的；要么按索引，要么让用户点在真实线上。
- `SelectByID2("", "EDGE", x, y, ...)` 在图纸中心这类空白处一定失败；点在真实投影线上才有机会命中。
  实测 `(0.1, 0.1)` 能命中，视图中心不能。所以工具先试坐标、失败再用 `edge_index`。
- `IEdge.Select4` 传 callout 会报"无效的参数数目"，用 `Select2(False, 0)`。
- `AddDimension2` 的 x/y 只是**尺寸文字位置**，与选中哪条边无关。

### 按图纸坐标补标（可用，2026-10-27 实测）

坐标选边本身行不通（`SelectByID2` 只认 VIEW），但可以**反过来投影**：
把每条可见边的中点从模型坐标换算到图纸坐标，再取最近的边选中。

```
paper = view.Position + (model - model_centre) * view.ScaleDecimal
```

三个前提：

1. `model_centre` 用该视图全部可见边端点的包围盒中心（不是零件 bbox）。
2. 要知道投影面：用视图 outline 的长宽比去匹配模型三组尺寸的长宽比
   （xy = 前视、xz = 俯视、zy = 侧视），最接近的一组即为投影面。
3. 换算只覆盖投影平面内的两轴，第三轴（深度）不参与比较。

实测：侧视图里取一条 275 mm 的边，按公式算出图纸点后补标成功（50 → 51 个尺寸）。

局限：投影面靠长宽比推断，**长宽比接近的视图可能判错**。优先用 `edge_index`
（`list_drawing_edges` 给出索引，稳定可靠），坐标路径作为便利入口。

### 尺寸去重怎么读、怎么比（2026-10-27 修正）

早前结论"读不回图纸尺寸"是错的，错在调用方式：

- `IView.GetDisplayDimensions` 确实只给 13 个 —— **不要用它**。
- 正确路径是 `IView.GetAnnotations` → `GetType() == 4` → `GetSpecificAnnotation.GetDimension2(0).SystemValue`，
  实测一张 50 尺寸的图能读到 **50 个**值。之前只看到 15 个，是因为值大量重复（5 mm 出现 5 次），
  不同值只有 11 种。
- 尺寸值是**投影长度**，不是三维边长：1253 mm 的边在侧视图里标注出来是 37.6 mm。
- 比较要用**整毫米**（`round(v*1000)`）：同一个尺寸在两个视图里最后几位浮点不同。
- 新增尺寸后 `doc.EditRebuild3` 才刷新注释列表，否则刚加的尺寸读不到。
- **`add_drawing_dimension` 必须保存**：只加尺寸不 `SaveSilent`，重开图纸就没了（这解释了
  为什么之前"补了尺寸却读不回来"）。

**去重为什么只能靠调用方传值**：判断"这个尺寸是否已有"必须知道它在新视图里会被标成多少，
而尺寸值是**投影长度**——同一条 1253 mm 的三维边，在主视图标 1253、在侧视图标 37.6。
读图纸现有值只能确认"哪些数值已存在"，无法预测新边在新视图的投影值。所以
`skip_duplicates` 默认关，只认调用方按目标视图投影长度传入的 `known_values`。

### 模板安全区要从模板自己量

把边距写死成 A3 数值（左 34mm / 下 60mm）在 A2 上会直接失败，不同模板的标题栏、
明细栏尺寸也不一样。正确做法是在**加任何模型视图之前**，从图纸的 `GetLines3` 量出来：

```python
sheet = doc.GetFirstView
segments = _template_lines(sheet, w, h)   # 必须在加视图前，否则会混入视图线
```

两个必须知道的细节：

1. **`GetLines3` 每条线是 12 个数**：前 6 个是元数据，后 6 个才是 `x1,y1,z1,x2,y2,z2`。
   按 stride 6 读，每隔一条就会读到 `(0, 2, 1, 1)` 这种垃圾，坐标范围会跑到几十万。
   用「坐标必须落在图纸范围内」过滤比猜偏移量可靠。
2. **明细栏右边界会把图框的行列编号带进来**（A3 量出 0.12 m），按它留边会吃掉 30% 宽度。
   实测 A3 合理值是左 0.0588 / 下 0.065 / 右 0.004 / 上 0.004，所以量到的带宽要按比例截断
   （左 ≤ 14% 纸宽、下 ≤ 24% 纸高）。

量出的安全区会随结果返回（`safe_area.source` 为 `template` 表示实测，`ratio` 表示回退）。

### 局部放大视图：12 个参数，最后一个不能是 0（2026-10-27 实测）

```python
doc.CreateDetailViewAt4(X, Y, Z, Style, Scale1, Scale2, Label, ShowLabel, Flip,
                        NoOutline, Shape, NumCircleSegments)
```

- **必须传满 12 个参数**，第 12 个是边界圆的分段数。传 0 会报 `非选择性的参数`
  （DISP_E_PARAMNOTOPTIONAL），不创建任何东西——这就是它一开始看起来"不支持"的原因。传 32 正常。
- 和剖视一样，要先在**源视图草图**里画边界圆并选中：
  `sk → CreateCircleByRadius(sx, sy, 0, r)`，坐标用 `(paper - view.Position) / scale`。
- 圆心必须落在**有几何的地方**（角部孔、折弯、圆角）；落在空白处会得到一个空的局部视图。
- **比例锁在父视图**：新建的局部视图 `UseParentScale = True`，此时写 `ScaleDecimal` 无效
  （读回仍是父比例），`Scale1/Scale2` 参数也不生效。顺序必须是
  `UseParentScale = False` → `ScaleDecimal = 2.0`，而且**新建视图要保存一次之后才吃得进去**。
  实测：创建后立即设置无效；`SaveSilent` 之后再设，`ScaleRatio` 从 `(1,2)` 变成 `(2,1)`。

### 剖视图：剖切线必须画在视图草图里（2026-10-27 实测）

`CreateSectionViewAt5(X, Y, Z, Label, Options, ExcludedComponents, Depth)` —— **7 个参数**，
但光调用没用，必须先有一条**画在该视图草图里的剖切线**并处于选中状态，否则返回 null。

可行流程（实测通过，得到"剖面视图 A-A"）：

```python
doc.ActivateView(view.Name)
sm = doc.SketchManager
sm.InsertSketch(True)
# 草图坐标 = (图纸点 - view.Position) / view.ScaleDecimal  ← 用转换实体标定过
sm.CreateLine(-half, 0, 0, half, 0, 0)      # 水平剖切线
sm.ActiveSketch.GetSketchSegments → segment.Select2(False, 0)
sm.InsertSketch(True)                        # 退出，选择会保留
section = doc.CreateSectionViewAt5(x, y, 0, "A", 4, None, 0)
```

要点：

- `Options = 4` 是 `swCreateSectionView_ChangeDirection`（翻转观察方向），0 为默认。
- **剖视图与父视图 x 对齐锁定**：`view.Position = (x, y)` 只有 y 生效（写 x 不报错但读回不变），
  所以剖视只能放在源视图的正上/正下方。布局时把它当"父视图的附属视图"，不要当独立视图排。
- 剖切线要**穿过几何**；画在视图外或空白处会返回 null（而且失败几次后视图状态会变差，建议重开图）。
- `Position` 决定局部视图放哪，与剖切线无关。
- 实测配电箱箱体：3 视图 → 剖视 A-A，尺寸从 50 涨到 68 个。
  但**4 个视图（主+俯+侧+剖视）在 A3 / 1:2 下空间不够**，会互相挤压或压到明细栏；
  改成「主视 + 侧视 + 剖视」3 个视图后 `drawing_audit` 零警告。视图多时优先减视图或换 A2。

### 钣金展开视图（2026-10-27 实测）

```python
view = doc.CreateFlatPatternViewFromModelView3(
    model_path, config_name, x, y, 0, hide_bend_lines, flip)
```

- **7 个参数**，没有"自动配置"：不传 config 就读模型的 `ConfigurationManager.ActiveConfiguration.Name`。
- 零件必须有 `FlatPattern` 特征（钣金件）；装配体不能展开，传装配会返回 null。
- 返回的 IView 与普通视图一样可以 `ScaleDecimal`、可以重排。
- 实测：配电箱门板（163×215×10 mm）展开后 97×127 mm @1:2，折弯线标注"上 90° R2.54"正确显示。

钣金件要标在**展开图**上（下料尺寸、孔位、折弯方向），不是折弯后的形状视图。

### 注释类操作必须自己保存

`InsertModelAnnotations3`、`AddDimension2`、`InsertCenterMark`、剖视/局部/展开视图创建之后，
文档是 dirty 的但**没有落盘**。之前只有 `add_drawing_dimension` 做了保存，结果是：

- 内存里 44 个尺寸，导出的 PDF 正常
- 磁盘上的 .SLDDRW 打开后**一个尺寸都没有**

给调用方的文件是空的。现在 `auto_dimension` / `add_center_marks` / `add_section_view` /
`add_detail_view` / `add_flat_pattern_view` / `add_dimension_tolerance` 全部 `SaveSilent`。

判断方法：任何"创建了对象"的操作，之后用 `drawing_dimension_values` 读一遍；读不到就是没保存。

### 圆柱面 ≠ 孔：IBody2.GetFaces 是方法

钣金件上每个 R2.54 折弯都贡献一个 Ø5.1 的圆柱面，所以"圆柱面清单"远多于"孔清单"。

- `IBody2.GetFaces` 是**方法**（必须 `GetFaces()`）；`IFeature.GetFaces` / `IComponent2.GetFaces`
  是属性。按属性调 body 的会静默拿到 bound method，扫描结果为空——这个坑让配电箱的 80 个圆柱面
  一开始一个都没扫到。
- 试过用「面法向 vs 径向」判孔/凸台：80 个面全部被判成同一类，与已知的 Ø6 安装孔矛盾，
  判据在本环境不可靠，已放弃。台账改为输出**全部圆柱面 + 直径/轴线/中心**，由调用方判断。

### 形位公差 / 粗糙度：能读不能写（2026-10-27 实测）

**能读**：形位公差的注释类型是 **5**（`swAnnotationType_Gtol`），读内容是

```python
gtol = annotation.GetSpecificAnnotation      # IGtol
gtol.GetFrameCount                            # 框格数
gtol.GetFrameValues(1)                        # ('0.01','','','','') 公差值 + 三个基准 + 符号
```

**不能写**，三条路都堵：

| 尝试 | 结果 |
|---|---|
| `doc.InsertGtol(...)` 0–4 个参数 | 全部 `找不到成员`（DISP_E_MEMBERNOTFOUND） |
| `doc.InsertSurfaceFinishSymbol2(...)` 4–9 个参数 | 全部 `参数数目` 错误 |
| 宏录制插入动作 | 只录到 `SelectByID2("细节项目275@图纸1", "GTOL", ...)`，插入本身不被捕获 |
| 复制粘贴：`doc.EditCopy()` 成功，`doc.EditPaste` / `EditPaste2` | **两个方法都不存在** |
| `gencache.EnsureDispatch` 早绑定 | `This COM object can not automate the makepy process`（SolidWorks 无类型库） |

**QueryInterface 试探（2026-10-27，最终结论）**

已知 `IDrawingDoc = {83A33D33-27C5-11CE-BFD4-00400513BB57}`、
`ISwDrawingDoc = {67930141-0BAC-4518-A40E-30AB15D6E9B8}`（注册表 HKCR\Interface 可查）。
用 `doc._oleobj_.QueryInterface(pythoncom.IID_IDispatch, IID)` 能成功拿到接口，绕开默认接口限制：

| 接口 | InsertGtol | InsertSurfaceFinishSymbol2 | InsertDatumTag |
|---|---|---|---|
| IDrawingDoc | 名字存在，**Invoke 报「找不到成员」** | 名字存在，报「非选择性的参数」 | 调用返回 None |
| ISwDrawingDoc | 同上 | 同上 | 同上 |

关键点：`GetIDsOfNames` 能找到方法名（类型库里有），但 `Invoke` 报 `DISP_E_MEMBERNOTFOUND`
—— 说明该版本（SolidWorks 2024）**运行时没有实现这个方法**，不是绑定方式的问题。
写 C# 插件也绕不过去。

**表面粗糙度有解（2026-10-27 查官方文档后实测通过）**

正确入口是 `IModelDocExtension.InsertSurfaceFinishSymbol3`，而且**必须给满 14 个参数**
（之前试 4–9 个，所以一直报「非选择性的参数」）：

```python
doc.Extension.InsertSurfaceFinishSymbol3(
    SymType, LeaderType, X, Y, Z, LaySymbol, ArrowType,
    MachAllowance, OtherVals, ProdMethod, SampleLen,
    MaxRoughness, MinRoughness, RoughnessSpacing)
```

实测 `(1, 1, 0.30, 0.10, 0.0, 0, 0, "", "", "", "", "3.2", "", "")` 成功返回 SFSymbol，
落盘后能在 `list_drawing_annotations` 里读回（`surface_finish` 类型）。
坐标是**图纸坐标**，给错位置会插在空白处，所以要么手动指定，要么先算好表面位置。

**形位公差确认无解（两条官方路径都不可用）**

官方文档给了两条路：

1. 旧：`ModelDoc2.InsertGtol()`（2016/2017 示例）→ 本机报 `DISP_E_MEMBERNOTFOUND`
2. 新：`IModelDocExtension.CreateGtolData()` + `ext.InsertGtol(data, x, y, z, ...)`（2015+ 推荐）
   → **本机 `CreateGtolData` / `InsertGtol` 两个成员都不存在**

所以不是参数问题、不是绑定问题，是该版本没有实现。C# 插件也绕不过去。

**形位公差四条路全堵（2026-10-28 在 SW2026 上逐一实测）**

| # | 路径 | 结果 |
|---|---|---|
| 1 | 图纸 `doc.InsertGtol()` | `DISP_E_MEMBERNOTFOUND` |
| 2 | 零件 `part.InsertGtol()` | 同样 `MEMBERNOTFOUND` |
| 3 | `IModelDocExtension.CreateGtolData` / `InsertGtol` | 成员不存在 |
| 4 | `InsertModelAnnotations3(0, swInsertGTols=0x20)` | 只能**搬运**模型已有 Gtol；模型侧创建不出来 → 无源可搬 |

补充排除的假设：不是"要先选中面"（选中状态对 MEMBERNOTFOUND 无影响）、
不是大小写（`InsertGTol` 同样失败）、不是对象类型（零件和图纸都试过）。

**注意 `MEMBERNOTFOUND` 的含义**：`GetIDsOfNames` 能找到方法名（类型库里有），
`Invoke` 却失败——说明该版本运行时**没有实现**，不是参数或调用顺序问题。
早绑定（C#）在同样运行时也会失败，除非 SolidWorks 提供别的入口。

**形位公差是能创建的——卡在 pywin32 的调用方式上（2026-10-28 突破）**

前面四条路都失败，原因不是 API 缺失，而是 **`doc.InsertGtol()` 被 win32com 当成属性读取
（DISPATCH_PROPERTYGET）**，于是必然返回 `DISP_E_MEMBERNOTFOUND`。

正确写法——用同一个 DISPID，但显式声明是方法：

```python
import pythoncom, win32com.client as win
ole = doc._oleobj_
dispid = ole.GetIDsOfNames("InsertGtol")
raw = ole.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True)   # ← 关键
gtol = win.Dispatch(raw)
gtol.SetFrameValues2(1, "0.05", "", "A", "", "")
anno = gtol.GetAnnotation
anno.SetPosition(0.30, 0.15, 0.0)
```

实测：插入 2 个框格，回读 `frame1=['0.05','','A','','']`、`['0.1','','B','C','']`，
`list_drawing_annotations` 计到 `geometric_tolerance: 2`，图上框格位置正确。

**留下的问题**：框格内的公差**符号**（平面度 ∇、位置度 ⌖）没写进去。
`SetFrameSymbols2` 的第一个字符串参数传 `<MOD-FLAT-PAR>` 这类格式码无效，传空得到空符号格。
需要找正确的符号写入方式（可能是 `swGdtSym_e` 枚举值，或特定的格式码语法）。

**方法论教训**：`DISP_E_MEMBERNOTFOUND` 有三种完全不同的来源——属性/方法派发方式错、
接口指针不对、真的没实现。只测一种调用方式就断言"API 不存在"，会把自己堵死两天。

**`IInsertGtol` / `INewGtol` 报的是另外的错误**（61836/61837 "无法读/写只读属性"），
说明这两个带 `I` 前缀的 dispatch 变体参数封送规则不同，暂未使用。

**结论**：表面粗糙度**可写**（`add_surface_finish`）；形位公差**可创建**（
`add_geometric_tolerance`），符号待补。

`list_drawing_annotations` 因此定位为**自检工具**：证明图上有没有形位公差、内容对不对。

### 读尺寸挂在什么上（§6.1 关联校验）

```python
ann.GetAttachedEntities2      # (entity, entity) —— 通常两条边
ann.GetAttachedEntityTypes    # (11, 11)  11 = EDGE
```

- 类型码：10 vertex / 11 edge / 12 face / 13 sketch_segment / 14 sketch_point。
- **没有关联实体的尺寸无法追溯**，按流程 §6.1 属于未完成；只挂在 sketch_segment 上的也要复核。
- 实测配电箱图 44 个尺寸：24 处挂边、7 处挂顶点、8 个 `type_1`、1 个 `type_0`、
  **21 个无关联**。孤儿比例这么高说明 `InsertModelAnnotations3` 导入的尺寸不一定绑到几何上，
  这正是「导入成功 ≠ 图纸正确」的实证。
- `delivery_check` 把「无关联尺寸」计入检查项，**默认按 warning**（`pending`），
  `strict_attachments=True` 才升级为 blocker（流程 §6.1 的字面读法）。

### 无关联尺寸是不是导入参数造成的？不是（2026-10-27 实测）

`InsertModelAnnotations3` 必须传 **6 个参数**（传 5 个报「非选择性的参数」）。逐一改参数试关联率：

| 调用 | 尺寸数 | 无关联 |
|---|---|---|
| `(0, T, True, False, False, True)` 现用 | 44 | 21（47.7%） |
| `(0, T, True, True, False, True)` 第4参数改 True | 36 | 17（47.2%） |
| `(0, T, True, True, True, True)` | 46 | 24（52%） |

**孤儿比例不随参数变化**，稳定在 47–52%。要么这些尺寸真的没绑几何，要么
`GetAttachedEntities2` 对它们就是不返回 —— 目前无法区分。

所以不能拿它一票否决：默认只报 warning，需要严格判定时传 `strict_attachments=True`。

### 装配体不能脱离引用复制（2026-10-28 教训）

把 `.SLDASM` 单独复制到临时目录去建图，组件引用全部断开，`CreateDrawViewFromModelView3`
会返回**没有几何的视图**：三个视图 outline 全是 0.0357 × 0.0252 m，`GetVisibleEntities2` 返回 0 条边，
看起来像"装配建图功能坏了"。

实际上原路径建图一切正常：`box = [0.8, 1.098, 0.8852]`、比例 1:20、视图有内容。

**结论**：测试装配必须保持引用完整——要么在原目录建图、要么整目录复制（含所有零件）。
只复制 `.SLDASM` 得到的"空白视图"不是产品缺陷。

### 材料：SetMaterialPropertyName2 与本地名称（2026-10-27 实测）

```python
part.SetMaterialPropertyName2("", "SolidWorks Materials.sldmat", "普通碳钢")
part.MaterialIdName     # -> 'solidworks materials|普通碳钢|9'
```

- 数据库名固定 `SolidWorks Materials.sldmat`；**材料名必须与本地材质库完全一致**。
  实测「普通碳钢」有效，「铝合金 6061」「不锈钢 304」写入后 `MaterialIdName` 为空 —— 名称不对。
- 写入后要 `EditRebuild3` 才会更新质量；没有材料时 SolidWorks 按 1000 kg/m³ 算，BOM 重量只是体积代理。
- 材质库 XML 在 `C:/ProgramData/SOLIDWORKS/SOLIDWORKS 2024\` 下能找到 `*.sldmat`，但
  `<material name="...">` 正则解析不出条目（格式/编码不符），自动列名称暂不可用。

### 真实图纸基线（100 张实测，2026-10-28）

用 `D:/142套solidworks工程图图纸模板` 里 1–3 号（6346 张 SLDDRW）随机抽 100 张跑
`drawing_audit` + `list_drawing_annotations`，**100/100 全部可读，单张约 7 秒**。结论：

| 指标 | 平均 | 中位 | 范围 |
|---|---|---|---|
| 视图 | 3.1 | 3 | 1–9 |
| **尺寸** | **14.2** | **11** | 0–58 |
| 注释 | 46.5 | 36 | 0–160 |
| 中心标记 | 4.9 | 3 | 0–59 |

制造类标注的出现率：

| 标注 | 有它的图纸 |
|---|---|
| **表面粗糙度** | **60 / 100** |
| **形位公差** | **7 / 100** |
| 零件序号 | 3 / 100 |

**据此改了三处判定**（原来都是我拍脑袋定的）：

1. `auto_dimension`：全量导入产生 44 个尺寸，是真实图纸（14 个）的 3 倍。现返回 `advisory` 提示，
   建议关闭 `include_unmarked` 或事后筛选。
2. `delivery_check`：**粗糙度升为常规检查项**（60% 的图有）、**形位公差降为可选**（仅 7%）。
   我们恰好是"粗糙度能写、形位公差写不了"，方向是对的。
3. `drawing_audit`：`no_dimensions` 原本按视图报警，**79/100 的真实图会误报**（轴测图、示意视图本来
   不带尺寸）。改为只在**整张图无任何尺寸**时报 `no_dimensions_anywhere`。复测 5 张真实图，误报归零。

**批量读图的硬性前提**：SolidWorks 同时开 40+ 文档后任何新文件都打不开
（"could not open the document"）。必须**每张读完就 `CloseDoc`**，否则第 2 张起全部失败。

### 台账对账：按值匹配要先过滤两类尺寸（2026-10-27 实测）

`reconcile_ledger` 把台账的驱动尺寸逐个到图纸上找对应值。直接比会有大量假缺失，两类要先剔掉：

| 类型 | 判断 | 例（配电箱） |
|---|---|---|
| 非长度尺寸 | `dim.GetType` 不是 0/2/3（线性/直径/半径） | 角度、坐标尺寸 8 个 |
| 超出包围盒 | 值 > 包围盒对角线 × 2 | 阵列间距 6000/8000，草图 2000/4000，共 4 个 |

过滤后：台账 58 个有效长度尺寸，图纸 22 个值，匹配 51，**覆盖率 87.9%，缺 7 个**。
缺的具体是 Ø30、Ø1、Ø4（沉孔）—— 这才是可执行的「补这些就齐了」清单。

注意匹配是按值且容差 0.5 mm，而图纸尺寸是投影长度，所以结果只能当待确认清单，不能当判决。

### AddHoleCallout2 必须先 SelectByID2 选边（2026-10-27 对照 SolidPilot 源码）

SolidPilot 的实现（`SolidWorksService.Drawing.cs`，AddHoleCallout）比我多做一步：

```csharp
modelDoc.ClearSelection2(true);
bool selected = modelDoc.Extension.SelectByID2("", "EDGE", px, py, 0, false, 0, null, 0);
if (!selected)
    selected = modelDoc.Extension.SelectByID2("", "SILHOUETTE", px, py, 0, false, 0, null, 0);
if (!selected) return "EDGE_NOT_FOUND";
object callout = drawingDoc.AddHoleCallout2(px, py, 0.0);
```

**选择集本身就是上下文**：先选中投影边，AddHoleCallout2 才知道标哪个孔、属于哪个视图。
只给坐标不选边，即使点位正确也返回 null —— 这就是我们之前全部失败的原因。

他们自己在注释里也承认这个方法的局限：
`coordinate selection is fragile for crowded geometry (KNOWN-LIMITATIONS #6)`。

**实测（配电箱结构图）**：按网格点试选 `EDGE`，15 个点里 12 个能选中 —— SelectByID2 本身有效；
但 `AddHoleCallout2` 在这些点上仍返回 null，因为网格落点命中的是**直线边**，不是孔的投影圆。
结论：方法正确，但必须先拿到**圆边上的精确点位**，而这需要扫边。

**扫边是硬瓶颈**：`GetVisibleEntities2(view, 1)` 取到 260 条边后，即使只用端点算弦长预筛
（不调 GetCurve），一个视图仍跑不完 400 秒。慢在实体枚举与逐边顶点读取，不在曲线判断。
要提速得换 API（如按视图分片、或用 `IView.GetPolylines7`），没在本次解决。

**当前结论**：孔标注在纯 Python 下可用性不足，建议界面手工或用 `add_drawing_dimension` 标 Ø 尺寸。

### 圆边识别：GetCurve 必须 Dispatch

```python
curve = win32com.client.Dispatch(edge.GetCurve)   # 少了这一步 IsCircle 恒失败
if curve.IsCircle: ...
```

配电箱侧视图 260 条可见边里有 **82 条是圆**，之前"找不到圆"纯粹是没包 Dispatch。
但逐边取曲线在大视图上很慢（260 边 > 300 s），批量孔标注要先按直径/位置过滤再取。

### PDF 只能来自工程图

`export_drawing` 的 PDF/DWG/DXF 都以 `.SLDDRW` 为输入；`batch_export` 的 STEP 以模型为输入。
两者不要混用——PDF 从零件导出会得到空壳或直接失败。

## 十四、一条通用原则

> **先读现值，再改。**
> 凡是"多层属性"（外观、显示状态、配置、材质），第一步读一次目标对象的现值，它会直接告诉你该改哪一层。跳过这一步，靠猜 API 语义，返工成本是它的 10 倍。

## 十五、标注身份、类型与位置（0.8.19，P1-6 / P0-2 / P1-7）

**尺寸身份是 `IDimension.FullName`，不是 `IAnnotation.GetName`。** `GetName` 只是图纸内
唯一的 `细节项目N` 标签；同一个驱动尺寸 `D1@草图1@零件.Part` 在 front / top / right
三个视图各插一次，会得到三个互不相同的 `细节项目N`。`auto_dimension` 靠
`InsertModelAnnotations3` 自带的 `eliminate_duplicates` 标志去重**没有用**（实测 56 个
尺寸散在 4 个视图）。可行做法：插入前快照所有 `GetName`，插入后 `EditRebuild3`，
再按 `IDimension.FullName`（回退 `Name`）分组，保留首次出现，其余用
`annotation.Select(False)` + `EditDelete` 删除。**只处理本次新增的标注**——
`before_names` 快照把用户先前放的尺寸全部排除，否则一次自动标注会删掉别人的成果。

**`IDisplayDimension.GetAnnotation` 必须按方法调用。** 晚绑定下写 `dim.GetAnnotation`
会被当 PROPERTYGET 派发，拿不到 `IAnnotation`；用 `call_method(dim, "GetAnnotation")`
再 `Dispatch` 才能读 `GetPosition`。这是 `add_hole_callouts` 回读引线真实落点
（`position_m`）的依据——请求落点 `anchor_m` 与真实落点可能不一致，必须都返回。

**`IAnnotation.GetType == 4` 才是尺寸。** note / GTOL / 粗糙度是别的类型。
`set_dimension_position` 先查类型，非尺寸直接回 `not_a_dimension`，避免名字打错时
静默移动错对象；`SetPosition(x, y, 0)` 之后必须 `EditRebuild3` 再回读
`GetPosition`（section / detail 视图会保持投影对齐而拒绝部分移动，所以要返回
`requested` 与 `position_after_m` 供调用方判断）。

**`EditRebuild3` 是方法不是属性。** `auto_dimension` 里曾写成裸 `doc.EditRebuild3`
属性访问，插入后 `view.GetAnnotations` 仍是旧集合，去重遍历扑空；必须
`call_method(doc, "EditRebuild3")` 真正调用才会刷新。

## 十六、局部视图、形位公差符号与表面粗糙度（0.8.20，P2-9 / P2-10 / P2-11）

### 局部视图：`scale` 参数曾被父视图比例覆盖

`add_detail_view` 里写过 `scale = float(target.ScaleDecimal)`，把用户请求的
2.0（2:1）直接换成父视图的 0.5，于是 `requested_scale` 也报 0.5、
`scale_locked` 恒为 False——一个永远抓不到的"锁"。修法：请求比例存
`requested_scale`，父视图比例单独存 `parent_scale`，**只**用来换算边界圈草图坐标
（`(paper - view.Position) / parent_scale`）；创建后 `_apply_detail_scale(doc,
detail, requested_scale)` 设 `UseParentScale=False` + `ScaleDecimal` 并回读。
实测 `_apply_detail_scale(2.0)` 生效，视图名从 `局部视图 A (1 : 2)` 变成
`局部视图 A (2 : 1)`。

**判空信号**：`IView.GetVisibleComponents` 在边界圈落在源视图几何之内时返回
1 个组件、`GetVisibleEntities2(component, 1)` 得到 1 条边；落在几何之外返回
`None`、0 条边。`GetVisibleEntities2(None, 1)` 会报"类型不匹配"，
`GetPolylines7` / `GetEdgeCount` / `IDisplayData.GetEdgeCount` 在这个晚绑定环境
不可用，别用它们判空。所以创建后数一次可见边，0 条就返回 `detail_empty`
（视图是空的，别报 ok:true）；比例被锁则返回 `detail_scale_locked` error。

### 形位公差：`<MOD-*-PAR>` 不是符号，是字面文本

`IGtol.SetFrameSymbols2(1, code, ...)` **返回 None**，成功失败都靠回读判定：
`GetFrameSymbols2(1)` 返回 6 元组，第一格就是刚写入的内容。实测
`SetFrameSymbols2(1, "<MOD-POS-PAR>", ...)` 回读 `('<MOD-POS-PAR>', ...)`——
格式码不会被解析成位置度符号，只会原样显示在框格里；而
`SetFrameSymbols2(1, "位置度", ...)` 回读 `('位置度', ...)`。
因此 0.8.20 起 `GTOL_SYMBOLS` 全部改成可读中文文字，写入后用
`GetFrameSymbols2` 校验，不一致就 `IAnnotation.Select(False)` + `EditDelete`
删掉刚建的框格并返回 `gtol_symbol_not_applied`。**这是文字标签，不是真正的
GD&T 几何符号**，出正式图纸仍需在界面替换。（另注：`GetFrameValues2` 在新建的
空 Gtol 上报成员找不到，回退 `GetFrameValues` 可用。）

### 表面粗糙度：`InsertSurfaceFinishSymbol3` 忽略 x/y

14 参调用能创建符号、`list_drawing_annotations` 也能列出 `surface_finish: 1`，
但 annotation 的 `GetPosition` 实测是 `[0.0, 0.0, 0.0]`（type 7）——
传入的 x/y 被忽略，符号落在图纸原点。早期只回显请求值
`position_m=[x,y]` 属于谎报。修法：创建后 `IAnnotation.SetPosition(x, y, 0)`
（实测返回 True、回读 `[0.3, 0.2, 0]`），再 `EditRebuild3`。

**像素可见性自检的坐标映射**：SolidWorks 图纸原点在**左下**，渲染出的 PNG
原点在**左上**，所以 y 要翻转：
`px = x / sheet_w * image.width`，`py = (1 - y / sheet_h) * image.height`。
在插入符号前后各 `SaveAsSilent` 一份 PDF、`render_drawing` 成 PNG，
用 `ImageChops.difference` 在 (px,py) 的 ±50 px 窗口统计变化像素。
实测修好位置后 diff bbox `(1684,493)-(1736,551)` 正好落在目标像素
`(1701,549)`，窗口 changed=265；未修位置前 diff 只在左下角原点。
DWG 是矢量、没有像素通道，只能检查文件非空，可见性以 PDF 为准。

## 十七、模型坐标 → 视图坐标：三条 API 路都堵死，最后靠尺寸匹配（0.8.21，P1-4）

要支持「用模型坐标指定剖切面」，必须先知道**图纸的两个轴分别对应模型的哪个轴**。
2026-10-31 在 SW2026 + 灌装阀阀体（59×28×80）上依次排除了三条路：

| 试的路 | 结果 |
|---|---|
| `IView.GetViewPlaneNormal` | 属性访问 `AttributeError`，按方法调用 `DISP_E_UNKNOWNNAME`（未知名称）——该版本不存在此 API |
| `IView.ModelToViewTransform` | **存在**，返回 16 个 double，但**不是行主序 4×4**：平移量与一个孤立的 `-1` 对不上，按任何常见分组都凑不出正交矩阵 |
| `IMathPoint.MultiplyTransform` | `DISP_E_MEMBERNOTFOUND`（找不到成员），无法让 SolidWorks 自己算变换 |

**还有一条更隐蔽的坑**：`IView.GetVisibleEntities2(component, 1)` 返回的是
**轮廓边（silhouette edge）**，不是 `IEdge`——`GetStartVertex` 返回 `None`，
只有 `GetCurve` 有值。实测 3 个视图各 10/14/11 个实体，取顶点全军覆没。
这直接意味着 `ink_outline()` 里的 `edge.GetStartVertex.GetPoint` **永远拿不到顶点**，
它只能返回 `None`；该函数保留为"理论上的精确交叉校验"，**不要依赖它出数**。
要数顶点只能走 `IBody2.GetEdges`（从零件本体拿），那是另一条更慢的路。

**最终方案：尺寸匹配（唯一可行且可解释的路）**

1. `IView.GetReferencedModelName` 拿模型路径（属性、方法都可用）；
2. `open_document` 打开零件，`GetPartBox(True)` 拿包围盒 → extent + centre；
3. 视图的 **deflated outline**（`GetOutline` 减 11.9 mm，见第十六节 P2-13）宽高，
   去和包围盒三轴 extent 的**两两组合**做匹配，误差最小且唯一者胜；
4. 误差超 `max(w,h)×5%` → `axes_mismatch`；两组误差相同（正立方体）→ `axes_ambiguous`，
   **拒绝而不是猜**——猜错就是切错平面，比报错严重得多。

**读 part box 会切换 ActiveDoc**：`open_document(model_path)` 之后必须重新
`_open_drawing(drawing_path)` + `Doc.ActivateView(name)` 才碰视图草图，
否则 `doc.SketchManager` 画到零件上去。

**offset 语义（实测锁定）**：`offset_m = (origin[dominant] - centre[dominant]) × scale`，
图纸米；`orientation != "vertical"` 时内部仍取反（0.8.10 的 top 视图草图 y 指向 -Z 那条）。
所以调用方始终按「模型坐标 − 零件中心」思考，不需要关心符号。

**剖切 sanity 的几何依据（不是启发式）**：剖切平面垂直于图纸**水平轴**时，
它**不可能**改变图纸水平方向的尺寸——剖视图的宽必须等于源视图的宽（竖切则比高）。
容差 10%，实测：

| 场景 | 源视图 deflated 宽 | 剖视 deflated 宽 | ratio | 判定 |
|---|---|---|---|---|
| 法向 +Y 过中心（正确） | 58.98 mm | 58.98 mm | 1.000 | ok |
| offset +0.012（切进 Ø20 孔，错） | 58.98 mm | 14.40 mm | 0.244 | `section_suspicious` |

`depth > 0` 的部分剖尺寸**本来就该缩小**，所以 sanity 记 `checked=false` 并写明原因，
不做判定——这与「宁可 not_checked 不要错数」一致。

**六场景真机结果（2026-10-31）**：法向 +Y → `horizontal` / 轴映射 `[0,1]` / ratio 1.000；
法向 +X → `vertical` / ratio 1.04；法向平行视线 → `cut_plane_not_in_view_plane`；
斜切 0.707 → `cut_plane_oblique`；错位 offset → `section_suspicious`；offset 0 → ok。
6/6 通过。

## 十八、COM 桥接自愈与 outline 单一真源（0.8.22，P0-1 / P2-13）

### 桥接：缓存的 dispatcher 会变成死指针

症状：连续约 50 分钟后所有调用报 `Not connected`，持续 120 s 以上不恢复，而 SolidWorks
本体完全正常（`GetActiveObject` 拿到的 `RevisionNumber` 仍是 32.4.0）。
根因：`get_sw()` 缓存 dispatcher 后**永不验证**，死指针一路复用。

三条实测支撑的实现决策：

1. **存活探针 18 ms**：`RevisionNumber` 一次往返实测 17.97 ms（20 次平均）。
   这个成本不能每次 `get_sw()` 都付（`open_document` 在一个工具里会调好几次），
   所以节流到 `LIVENESS_PROBE_INTERVAL_S = 2.0` 秒；探针失败立即丢缓存重新
   `GetActiveObject`。
2. **`pythoncom.CoRegisterMessageFilter` 在本机不存在**（`hasattr` 为 False），
   所以「SW 忙 → 拒绝调用」不能靠 COM message filter 自动重试，只能手工分类：
   - 忙（`RPC_E_CALL_REJECTED` 0x80010001 等）→ 退避 0.25 s / 0.5 s 重试，最多 2 次
   - 断线（`CO_E_OBJNOTCONNECTED` 0x800401FD 等）→ **不重试**，丢缓存 + 抛 `BridgeLost`
   - `DISP_E_MEMBERNOTFOUND`（0x80020003）两者都不属：既不重试也不重连
3. **不做线程级硬超时**：pywin32 在拥有对象的 STA 线程上派发，把调用丢到工作线程
   再「超时放弃」会破坏 COM 公寓，而不是取消调用。所以覆盖的是「失败后永不恢复」
   这个真实症状，而不是「单次调用挂死」（后者是 SolidWorks 侧状态，客户端无法安全中断）。

顺带修掉一个更早的坑：`_LAST_CALL` 只在第一次 `note_last_call()` 之后才存在，而
`log_failure()` 会读它——**最早那次失败**（正是断线时刻）会让日志函数自己抛
`NameError`，写进 `_logger_error.log` 而不是 `mcp-*.log`。现在模块级有默认值。

### outline：11.9 mm 是唯一真源，逐顶点路全部堵死

`IView.GetOutline` 宽高各多报 11.9 mm（固定偏移）。曾写过 `ink_outline()` 想用可见边
顶点算精确 bbox 做交叉校验，2026-10-31 实测**四条路全断**：

| 路 | 结果 |
|---|---|
| `IView.GetVisibleEntities2(comp, 1)` | 返回**轮廓边**，`GetStartVertex` → None，只有 `GetCurve` |
| `IBody2.GetVertices` | 该阀体 31 面 / 32 边全是闭合圆边，返回 **None** |
| `IEdge.GetStartVertex`（走 call_method） | 仍返回 None（32 条边、0 个顶点） |
| `ICurve.GetEndPoint` / `GetTessPoints` | 该版本**没有**这些成员（AttributeError / UNKNOWNNAME） |

所以 `ink_outline()` 从来只会返回 None——一个静默失效的「交叉校验」，已删除。
`precise_outline()` 常数内缩不是捷径，而是唯一可行解，精度实测三视图
59.0×28.0 / 59.0×80.0 / 80.0×28.0（真值 59×28 / 59×80 / 80×28，误差 <0.1 mm），
detail 视图差值 11.92 mm。

注意 `ISheet` / `GetFirstView.GetOutline` 是**精确的**（实测 420×297 mm），不受该偏移
影响——所以像素映射（`px = x / sheet_w * W`）不用改。

## 十九、标注删除的成功判定与尺寸显示文本（0.8.23，P0 / P1-1 / P3-3）

### 删除成功不能只看图纸视图的注解数

`doc.GetFirstView` 是**图纸**（Type 1），它的 `GetAnnotations` 数量在删除**尺寸**时
**不变**——尺寸属于模型视图。实测（2026-10-31，灌装阀阀体）：

| 时机 | 图纸1 | 工程图视图1 | 全视图合计 |
|---|---|---|---|
| 删除前 | 83 | 8 | 122 |
| 删掉一个 D1 后 | **83（不变）** | **7** | **121** |

所以 `removed = after < before` 若只比图纸，会把**成功的删除报成 `delete_failed`**——
比失败更糟：调用方会重试，然后误以为"这工具删不掉"。现在改为跨全部视图求和，
并在回执里写 `counted_scope: all_views`。

附带一个更隐蔽的事实：`IAnnotation.Select(False)` 对尺寸**是有效的**
（`GetSelectedObjectCount2` 返回 1），`EditDelete` **也确实删掉了**。v0.8.14 之所以
"四类全失效"，根因与类型无关——`_find_annotation_by_name` 里写了未导入的
`call_method`，`NameError` 被外层 `except` 吞成"没有候选"，于是对**刚被列出过的名字**
回 `annotation_not_found`。

**重名**：SolidWorks 按草图复用 `D1`/`D2`/`D3`，一张图实测有 **19 个 `D1`**。
`delete_annotation` / `move_annotation` 现在接受 `position_m`，用
`list_drawing_annotations` 报告的位置精确定位具体那一个。

### 尺寸的显示文本在槽 1，不是 `GetText()`

| 调用 | Ø20 直径尺寸 | 线性尺寸 |
|---|---|---|
| `IDisplayDimension.GetText()` | `''` | `''` |
| `GetText(0)` | `''` | `''` |
| `GetText(1)` | **`'<MOD-DIAM>'`** | `''` |
| `GetText(2)` | `''` | `''` |
| `GetType` | **5** | 2 |

`GetText()` 不带槽号对尺寸返回空串——尺寸的值与格式码都在**槽 1**。
`delivery_check` 之前用"文本里有没有 `Ø` / `MOD-DIAM`"判直径，于是恒为 0 处，
把一张印着 Ø28/Ø23/Ø12/Ø18/Ø13 的图判成「仅供参考/未完成」。

现在两条路并用：`dimension_type`（`IDisplayDimension.GetType`，5 = 直径）作主判据，
`raw_text`（槽 1 原文）+ 解析后 `text`（`<MOD-DIAM>`→`Ø`、`<DIM>`→数值）作兜底。
实测 38 个尺寸里 18 个 `dimension_type == 5`，`delivery_check` 从「0 处 Ø 尺寸」
变为「17 处 Ø 尺寸」。

### 0.0 不是长度

`reconcile_ledger` 会把值为 0.0 的台账项当成"图纸缺项"，但 0.0 是草图几何关系/约束值，
任何图纸都标不出来。现在分离到 `skipped_zero_value`，覆盖率分母改用
`comparable_dimensions`（实测 45 个台账项里 2 个是 0.0 → 43 个可比项）。

## 二十、圆边枚举、标注碰版与注记避让（0.8.24，P1-2 / P2-4 / P3-1 / P3-2 / P3-4）

### 视图里的"边"没有顶点，`GetStartVertex` 恒为 None

`list_circular_edges` 早期用 `edge.GetStartVertex` / `GetEndVertex` 算弦长做预筛，
在阀体上返回 **0 条**——因为视图可见实体是**轮廓边（silhouette edge）**，没有顶点
（与第十八节同源）。改用 `ICurve.CircleParams`（一次调用给出圆心/轴/半径）后同一张图
返回 **23 条**，坐标由 `_view_axis_map` 的尺寸匹配投影；映射不可用时给 `x: null`
而不是编一个坐标。

### `AddHoleCallout2` 的失败与选择方式无关

同一张图、同一批圆边，三种选择方式全部返回 null：

| 方式 | 选中 | `AddHoleCallout2` |
|---|---|---|
| `SelectByID2("", "EDGE", 边位置)` | ✅ | **null** |
| `edge.Select2(False, 0)` | ✅ | **null** |
| 相邻面 `face.Select2(False, 0)` | ✅ | **null** |

而同一调用在钣金拉伸孔（门板 Ø8/Ø15）上成功。所以判据不是"选没选中"，而是**几何类型**。
0.8.24 起返回 `verdict: no_native_callout` 并附 `native_callout_note`，
让调用方不要再换 offset 反复重试。

### 布局避让：标签 note 是障碍，不是尺寸

两个独立 bug 叠在一起：

1. `int(annotation.GetType())` —— `GetType` 在晚绑定下是**属性**，加括号必抛
   `TypeError`，而外层 `except: pass` 把它吞了 → **类型过滤从未生效**，note 被当成尺寸搬动。
2. 即便 note 不被移动，sweep 也只避让尺寸，把 Ø8 推进了 "2:1" 标签，渲染成 `2Ø8¹`。

修法：`annotation_type()` 属性优先 / 方法回退；`kind != 4` 才当尺寸；note（`kind == 6`）
作为 **fixed** 障碍塞进 `placed`，撞上时尺寸**向下**让（向上会撞回视图）。

### 注记避让的估计口径

`add_note(avoid_collisions=True)` 按「行数 × 字高 × 1.6 行距」估高，
按「最长行 × 字高 × 0.75」估宽（中文近方、拉丁约 0.55，取中偏保守）。
障碍用 `precise_outline`，所以 11.9 mm 虚胖不会把注记无故推走。
实测：注记放在视图中心会被下移 3 次（`0.253 → 0.2074`）落到视图下方。

### 版本号不能说明"跑的是哪一份代码"

一份反馈花了一整轮去追已经修好的问题：MCP 宿主指向的是**旧副本**的 editable 安装
（0.8.14），而工作区里已经是新版本；机器上同时存在三份源码目录。
版本号无法澄清这一点，**路径可以**。所以 `ping` 的 `runtime` 与 `check_capabilities`
都回显 `server_path` / `python_executable` / `python_version` / `pid`——
一次调用就能确认口径，不必再去翻 `.pth` 文件。

## 二十一、裸名字 NameError 被 `except` 吞掉：一个藏了五个版本的 P0（0.8.25）

### 现象

`auto_dimension(eliminate_duplicates=True)` 从 0.8.19 起一直报
`duplicates_removed: 0`，看起来像"这张图本来就没有重复尺寸"。
实际上每一次删除都抛 `NameError`，被外层 `except` 吞掉，一条也没删成。

### 根因

`drawing.py` 里有 5 处 `_call_method(doc, "EditDelete")`，但**模块级从未定义 `_call_method`**——
只有一处 `from ..sw_session import call_method as _call_method_sf`（别名不同）。
Python 在运行时才解析裸名字，于是：

- 删除 → `NameError: name '_call_method' is not defined`
- 外层 `try/except Exception` → 静默跳过
- 汇总 → `removed=0`，**看起来完全正常**

同一个坑在 `layout.py` 的 `_find_annotation_by_name` 上出现过一次（v0.8.14 的"四类标注全失效"），
这次是第二次。

### 为什么两道防线都没抓到

| 防线 | 为什么漏 |
|---|---|
| `static_check.py` | 只校验 `module.attribute` 形式（`name.attr` 的 `name` 有没有 import），**裸函数名不在范围内** |
| `regressions.py` 的 smoke 契约 | 断言写的是 `isinstance(duplicates_removed, int)` —— `isinstance(0, int)` 也是 True |

### 修法与新增防线

1. `drawing.py` 模块级 `_call_method = call_method`（并在 import 列表里加 `call_method`）。
2. `auto_dimension` 补 `dedupe_failed_count`（原来只有截断到 10 条的 `dedupe_failed` 列表，
   "删 0 失败 23"和"干净的空操作"无法区分）。
3. **新回归断言 `assert_no_undefined_bare_calls`**：遍历 `src/` 全部 21 个模块，收集每个
   `ast.Call` 的裸函数名，要求它能解析到内置函数 / import / 定义 / 参数 / 赋值 / `global`。
   这是**整类 bug 的通用防线**，不再依赖"记得给某个函数加断言"。
4. smoke 的断言从 `isinstance(...)` 改成「`removed == 0` 且 `failed > 0` 即失败」，
   把那次的签名直接写成判据。

### 同类第二次：定义被整段删掉

同一份 v0.8.24 补丁在插入两个新函数时，把 `_EDGE_CACHE = {}` / `EDGE_CACHE_LIMIT = 32`
当上下文行一起删了，`list_drawing_edges` 立刻报 `_EDGE_CACHE is not defined`，
整条边坐标读取失效。所以通用检查必须覆盖 **`ast.Name`（Load）**，不能只看 `Call`——
变量引用不会出现在 `Call` 里，定义被删时也一样静默到运行期才炸。

### 一条通用规则

**任何 `except` 吞掉的分支，都要有一个"结果不可能为 0"的断言。**
失败的删除和没有可删的对象，返回值必须能区分——否则静默失败会伪装成正常空操作。

## 二十二、尺寸摆位：~10 mm 是绝对值，不随零件大小变（0.8.27）

`InsertModelAnnotations3` 把每条尺寸放在离几何约 **10 mm** 的地方——这是**绝对距离**，
与零件尺寸无关。Ø46 的圆盘上塞 25 条尺寸，全部落在轮廓内互相压字；换成 4:1 只是把
拥挤稀释掉，改的是图纸对尺寸的表达，不是排版。

`auto_dimension(place_outside=True)` 的做法：

1. 只取**本次调用插入**的尺寸（`before_names` 快照之后出现的），手放的永不移动；
2. 按 `precise_outline` 判断属于哪个视图（在框内即归属，否则取中心最近的）；
3. 取**最近的一条边**作为方向（left / right / top / bottom）；
4. 同一侧按另一轴排序后依次外扩：距离 = `gap + slot × step`，
   所以同侧两条尺寸**不会共用一个距离**；
5. 结束时 `EditRebuild3` + 落盘——`auto_dimension` 其余分支都不保存，
   这个分支必须自己存，否则图纸上什么都没变。

气缸盖实测（25 条尺寸）：

| 指标 | 默认摆位 | `place_outside=True` |
|---|---|---|
| 落在视图框内 | 21 / 25 | **0 / 25** |
| 文本重叠对 | 17 对 | **6 对** |
| 四侧分布 | — | 上 8 / 下 6 / 左 7 / 右 4 |

剩下 6 对重叠说明分层只解决了"出框"，不解决"同侧相邻两条太近"——
需要在外面再跑一次 `layout_dimensions` 收尾。

## 二十三、标题栏一字段多格，视图标签藏在 `<VLLABEL>` 里（0.8.28）

### 同一个 marker 会命中多个单元格

GB 图框在**框外**还有一条 180° 倒置的「图样代号」（在 `[0.053, 0.285]`，图纸左上角），
标题栏里的正主在 `[0.391, 0.022]`。原实现 `marker in text` 就 `planned.append`，
于是两处都写——图上多出一个倒着的代号。

修法：收集候选，按 **(是否倒置, 距 `TITLE_SLOTS` 的距离)** 排序取第一个，
其余写进 `skipped_extra_matches` 供核对。

### 空单元格没有 note，位置兜底也找不到

`_note_at(sheet, slot)` 只能找"已经存在的 note"。GB 模板那几格是空表格（只有线），
根本没有 note 对象，于是 `planned: []`。第三级兜底改为 **`InsertNote` 在槽位新建**，
注意位置必须 `call_method(ann, "SetPosition2", ...)`——属性派发会抛 MEMBERNOTFOUND。

### 视图标签的 `GetText` 是 `<VLLABEL>/<VLSCALEV>`

v0.8.24 曾判定它"复现不出来"，因为当时只看 `list_drawing_annotations`（那条路径没返回它）。
直接读 `IAnnotation.GetSpecificAnnotation.GetText` 就能拿到。渲染后是「A—A」叠「2.5 : 1」，
所以**估算宽度必须按可见文本**：19 个码字符直接乘字宽会得到 40 mm 的禁区，
把无关尺寸推得很远。

另外：**标签 note 在改比例后才出现在 `view.GetAnnotations` 里**——所以避让必须挂在
`set_view` 改比例之后，而不是等 `layout_dimensions` 事后处理。
实测：停在标签 `[0.33, 0.6371]` 的尺寸被移到 `[0.33, 0.6497]`。

## 二十四、剖视对齐与切割线：一个能解除，一个只能告警（0.8.29）

### 解除对齐有正经 API

`IView.RemoveAlignment()` 就是 UI 的「解除对齐」，必须在写 `Position` **之前**调用——
顺序反了 SolidWorks 会把坐标改回去。文档明确：它只清当前视图的对齐，不影响子视图
继承来的对齐。

实测（垂直剖）：请求 `y = 父视图 y + 0.05`，不解除 → 回 `0.2793`（= 父视图 y）、
`aligned_to_parent: true`；`break_alignment=True` → 落到 `0.3293`。

### 切割线跟随：`IDrSection` 可读，写入方式未探明

`IView.GetSection` **只在剖视图上**返回 `IDrSection`（父视图返回 None）：
- `GetLineInfo` 可用，返回 12 个 double（切割线两端在模型坐标的点）
- `GetLabel` → 'A'、`GetView` → 父视图
- **`SetLineInfo` 存在但所有尝试都失败**：直接调用返回 False，
  按 6/7/8/10/12 个参数传都报「无效的参数数目」——多半需要 VARIANT 数组或
  额外的 line-count 参数，本轮没探明

所以 v0.8.29 仍然只是「移动父视图时告警」（`derived_views` + `cut_lines_detached_warning`，
0.8.26 已做），没有真正让线跟随。要真修，下一步得先确定 `SetLineInfo` 的签名，
或者改走「重建剖面」路线（代价是该视图上的尺寸会丢）。

## 二十五、驱动尺寸不能"补插到第二个视图"，但能在那重新标（0.8.30）

### 参数改不出"重复插入"

`InsertModelAnnotations3(source, types, allViews, duplicateDims, hidden, ...)` 的第 4 个参数
看着像"允许插重复"，实际**不是决定因素**——这个调用是**文档级幂等**的：同一模型尺寸
一旦插入过，第二次调用直接返回 `None`。

实测（履带轮 P05，两视图）：

| 组合 | 返回 | 视图1 | 视图2 |
|---|---|---|---|
| `all_views=True` 首次 | 12 | 12 | 0 |
| 激活视图2 + `all_views=False` + `duplicateDims=True` | **None** | 12 | 0 |
| 激活视图2 + `all_views=False` + `duplicateDims=False` | **None** | 12 | 0 |
| `all_views=True` + `duplicateDims=True` | **None** | 12 | 0 |

所以「局部放大视图永远 inserted: 0」不是 bug，是 API 的语义。

### 能做的：在同一几何上重新标一条

`add_drawing_dimension` 引用的是**视图里的边**，不是"某条模型驱动尺寸的副本"，
所以它在第二个视图里能正常建：

```
list_drawing_edges(view_name="工程图视图2")        # 147 条边
add_drawing_dimension(view_name="工程图视图2", edge_index=0)
→ {"created": true, "value_m": 0.11, "selected": "edge_index"}
```

视图2 的尺寸数从 0 变成 1。数值必然与主视图一致（同一个模型几何），
代价是它是一条**新标注**（`FullName` 不同），不是原尺寸的副本——
所以 `reconcile_ledger` 会看到"图上有两条同值尺寸"，这属于正常。

结论写进 `check_capabilities` 的 `duplicate_dimension_to_view`（`unsupported`），
避免下一个使用者再花一轮去试参数。

## 二十六、剖面切割线：三个接口都够不着那张图上的线（0.8.31）

想让"移动父视图时切割线跟着走"，试完四组接口，都不行：

| 接口 | 结果 |
|---|---|
| `IDrSection.GetLineInfo` | 12 个 double，**模型坐标**（±1.03 m，而 A3 只有 420 mm）。移动视图前后**完全相同** |
| `IDrSection.SetLineInfo(VARIANT 数组)` | 能写，返回 `True`——但改的是**模型空间的剖切平面**，等于把剖面切到别的地方，不是移动图上那条线 |
| `IView.GetSectionLineInfo2` | 只读，且同样返回模型坐标（值 -1.0347 与 `GetLineInfo` 一致） |
| `IView.GetSketchCount` / `GetSketches` | 晚绑定下**不存在**（`DISP_E_UNKNOWNNAME`），触达不了图纸上的草图 |

`SetLineInfo` 的正确签名在文档里：**只接受一个参数**，是 VARIANT 数组，
每段 6 个 double（起点 xyz + 终点 xyz），多段的公共点必须保持一致。
之前按 6/7/8/10/12 个散参数传，全都报「无效的参数数目」——那是参数形式错，不是方法不存在。

### 顺带发现更基本的问题

渲染核对（移动前后各导一次 PDF）显示：`add_section_view` 建出来的切割线
**在图纸上的位置与父视图并不重合**——父视图在左上角，线飘在图纸中部。
也就是说「移动后脱离」之前，它就已经没压对地方。

原因是切割线用 `ISketchManager.CreateLine` 画在草图里，而坐标基准（图纸绝对 vs 视图局部）
在晚绑定下无法验证（草图对象拿不到）。修它需要重定坐标基准，属独立一轮。

### 现在的建议流程

1. **先把视图摆到位**（`place_views` / `set_view`），最后一步才 `add_section_view`
2. 已经移过 → `delete_view` 删剖面重建（代价：该视图上的尺寸会丢）
3. 每次改完 `export_drawing` + `render_drawing` 目视核对线是否压在父视图上

## 二十七、切割线画错位置：一个表达式里两个错（0.8.32）

`add_section_view` 原来这样算半长：

```python
half = (max(width, height) * 0.62) / scale     # width 来自 GetOutline
sm.CreateLine(offset, -half, 0.0, offset, half, 0.0)
```

两处都错：

### ① `width` 已经是图纸米，不该再除 scale

`GetOutline` 返回的是**图纸坐标**，`width = outline[2] - outline[0]` 因此已经是图纸米。
除 scale 后：`0.1669 × 0.62 / 0.1 = 1.0347 m` —— **A3 只有 420 mm**，线长 2 米多。

这个 1.0347 就是 `GetLineInfo` 报出的 ± 值，之前误判成"模型坐标"，
其实是**自己画进去的错误长度**。

### ② `CreateLine` 用绝对坐标，没加视图位置

草图在**图纸层**，坐标是图纸绝对坐标。写成 `(offset, ±half)` 就是 `x = 0`
（图纸左边缘），只有线中段恰好扫过视图，看起来像"线在视图附近但错位"。

### 修法

```python
precise_box = precise_outline(target) or outline
half = max(precise_box[2] - precise_box[0], precise_box[3] - precise_box[1]) * 0.62
...
sm.CreateLine(pos[0] + offset, pos[1] - half, 0.0,
              pos[0] + offset, pos[1] + half, 0.0)      # 竖剖
```

用 `precise_outline` 而不是 `GetOutline`，顺带把 11.9 mm 虚胖排除掉（否则线再长 24 mm）。

实测（履带轮，A3，视图 scale 0.1）：

| 指标 | 修前 | 修后 |
|---|---|---|
| 半长 `half` | 1.0347 m | 0.1035 m |
| 线 x | 0（图纸边缘） | 0.3133（= 视图中心） |
| 线 y | -1.03 ~ +1.03（图纸外） | 0.1832 ~ 0.3754（视图 0.2596~0.299，上下各伸出 ~80 mm） |
| `sanity` | — | `ratio 1.0, verdict ok` |

**「移动父视图后跟随」仍不支持**：`SetLineInfo` 会把多段独立短线重排成一条折返线
（实测两段被连成往返），所以只能保证**初始位置正确**，移动后仍需重建。

## 二十八、公差单位 ×1000 与"读不到文字就删标注"（0.8.33）

### `SetToleranceValues` 收米，代码传了毫米

`add_dimension_tolerance(value_mm=50, tolerance_type="symmetric", upper=0.1)` 渲染成
**"50 ±100"**——`0.1` 被当成 **0.1 米**。而回执一切正常：`applied: true`、
`tolerance_type_readback: 4`。**数值对了、类型对了、只有图是错的。**

这类错误只有两条防线：知道 API 的单位契约，或者在返回值里给**渲染出来的文本**。
现在两条都有：传参 `/1000`，并新增 `rendered_text`（PDF 实测 `±0.10`）。

教训：`applied: true` 只证明"调用没抛异常"，不证明"图对了"。
凡是**写入数值**的工具，回执都该带一个从文档读回来的字符串。

### SW2026 上 `GetText` 对所有孔标注返回 None

v0.8.26 用「文本里有没有 Ø」作为孔标注的准入门槛。在 SW2026 上
`GetText(0/1/2)` 对每一条 callout 都返回 `None`，于是 `has_diameter=False` →
走 `not has_diameter` 分支 → `_discard()` 删除。

结果：全新图纸 8 个孔，`created: 0 / failed: 8 / no_native_callout`——
**能力从"标注不理想"退化成"标注不存在"**，而这恰好是 README 头部第一项能力。

修法不是"再加一层校验"，而是换判据：

```python
text_unreadable = not str(text or "").strip()
if radius_only:                                  # 文本明确以 R 开头才拒
    _discard(...); continue
if not has_diameter and not text_unreadable:     # 有文本但没 Ø 才拒
    _discard(...); continue
# 其余一律保留，created 行标 text_unreadable=True
```

**原则：缺证据 ≠ 反证据。** 读不到就打 `text_unreadable` 让调用方自己决定，
不要替用户删掉已经建出来的东西。

### 顺带：`SystemValue` 对 callout 不统一

同一次调用里，印着 `Ø8` 的 callout `SystemValue` 是 **4.0（半径）**，
印着 `Ø15` 的却是 **15.0（已经是直径）**。所以 `list_drawing_annotations` 现在
**以文本里的数字为准**，取不到才回退 `SystemValue` 并标 `diameter_source`。

## 二十九、`_view_axis_map` 漏乘比例：一条 bug 卡住三个工具（0.8.34）

```python
scored.append((abs(extent[a] - width) + abs(extent[b] - height), (a, b)))
```

`extent` 来自 `GetPartBox`，是**模型米**；`width/height` 来自 `precise_outline`，是**图纸米**。
没乘视图比例，所以只要视图不是 1:1 就匹配不上，直接返回 `axes_mismatch`。

连锁后果：

| 工具 | 症状 |
|---|---|
| `list_circular_edges` | 全长返回 `x: null / position_unavailable`（实测门板 0/9，修复后 9/9） |
| `add_section_view(cut_plane_origin=…)` | 无法把模型坐标映射进视图，返回 `cut_plane_unmapped` |
| `add_hole_callout`（单点） | 没有可匹配的圆，一切坐标都 `edge_not_found` |

一行 `extent * view_scale` 三个都通。**教训**：跨坐标系的比较，先问一句"这两边单位一样吗"。

### 单点孔标注：匹配修好了，调用还没通

改用「枚举圆 + 投影取最近」后：匹配距离 **0.0 m**、`Select2` 后选择计数 **1**，
但 `AddHoleCallout2` 仍返回 `None`。批量版在同一个视图能成功，差异未定位。

已把证据写进错误体（`matched_circle` + `selected_before_call`），下轮不必重跑盲猜。

另外两个已经踩过的坑，记在这里省得再踩：
- `_view_axis_map` 读零件包围盒会 `open_document(model)`，**把活动文档切成模型**，
  之前取的 `view` / `edge` 全部失效——`Select2` 在死句柄上会"成功"，`AddHoleCallout2` 随即返回 None；
- 孔有**两面同心圆**，只有近面可标，批量版按 `z` 最小取，单点版也照此。

## 三十、单点孔标注：匹配对了就通；`Ø9→Θ99` 是文字槽不是数值（0.8.36）

### 单点 `add_hole_callout` 为什么之前必失败

三个原因叠在一起：

1. **`_view_axis_map` 漏乘视图比例**（0.8.34 修）→ 拿不到圆心的图纸坐标；
2. **按坐标射线选边命不中**（v0.8.23 就发现过）→ 改成枚举圆 + 投影取最近；
3. **按直径匹配会选错孔**——同径多孔时取到的是别的孔，于是「选边成功、`AddHoleCallout2` 返回 None」。

改用**精确模型圆心**（`CircleParams[0..2]`）匹配后：门板 **Ø8×2 + Ø15×3 = 5/5 成功**。

### 那 4 个 Ø6 圆本来就标不出来

最初验证时取 `rows[:3]`，正好是 4 个 Ø6 —— **批量版对它们也是 0 产出**（非孔或背面）。
所以那不是 bug。错误提示已改成「这个圆可能不是可标注孔（外圆弧或背面）」，而不是
「找不到边」——后者会让人一直重试坐标。

### `Ø9 → Θ99`：文字槽冲突

渲染图实测：Ø8 正常（`Ø8`），**Ø15 印成 `Ø1515`**。`GetText` 回读始终是 `Ø15`，说明值没错。

原因是 `AddHoleCallout2` 的标注文字分槽：**直径一个槽、深度另一个槽**。
`SetText(slot, "Ø15")` 只替换直径槽，深度（15）仍在另一槽 → 拼成 `Ø1515`。
反馈里的 `Ø9 → Θ99` 是同一件事（`Ø` 在 PDF 里被提取成 `Θ`，纯字形映射）。

三条路都试过：

| 做法 | 结果 |
|---|---|
| 只写直径 | 深度粘连（`Ø1515`） |
| 写直径后清空其余槽 | **标注整条消失** |
| 完全不写 | 印出字面量 `<MOD-DIAM>` |

所以这不是「一行能修」的问题，登记为 `hole_callout_text_slots`（`partial`）：
通孔正常，带深度的孔会粘连，workaround 是改用 `add_drawing_dimension` 或人工改一次文字。
回执里 `text_source=SetText_depth_may_trail` 就是这条限制的标记。

## 三十一、假脏标记、假尺寸、真空隙（0.8.37）

### 工程图操作会给模型打 `NeedsRebuild2`

删一条标注，模型没动，`model.Extension.NeedsRebuild2` 却变成 `True`。
`delivery_check` 把它记成 `needs_rebuild`，verdict 从「待设计审核」掉到「仅供参考/未完成」——
一个调用方在这个工具集里**永远清不掉**的阻塞项。

修法：读到脏就**先强制重建一次再读**。真改动重建后仍在，工程图引起的会消失。
实测：删标注后 `needs_rebuild` 消失，剩下的 `suppressed_features` 是门板真有的压缩特征。

### 错的边也会给出"合理"的尺寸

`AddDimension2` 只要求"当前有选中的实体"，不校验**选中的是不是你以为的那条**。
反馈里 `edge_index=71`（Ø36 分度圆弧段）量出了 Ø46 外圆的 `R23` —— 一个看起来完全正常的错误值。

便宜的判据是**关联实体**：建完读 `IDisplayDimension.GetAnnotation` 的
`GetAttachedEntityTypes`，为空说明这次选择根本没落上，撤销并报 `edge_mismatch`。
比"比对数值"便宜，也不会被"巧合相等"骗过。

### 去重保留"首次出现"，剖视图因此永远是空的

`eliminate_duplicates` 按模型 `FullName` 保留第一条，而插入顺序是**主视图在前**，
所以两个剖视图上的副本全被删掉（实测 83 插入 → 删 44 → 剖视 0 尺寸）。
而深度、内腔这类尺寸**只有剖视能承载**，`ledger_missing` 随即报 7 条缺料。

现在的 advisory 把这件事写清楚了，并给出两条出路：
`eliminate_duplicates=False`（全留，自己删），或 `dimension_filter`（精确保留）。

## 三十二、仓库卫生：四轮挂账一次清（0.8.38）

### `.gitignore` 只挡 CAD 扩展名，挡不住 zip/png/jpg/csv

规则里列了 `*.SLDPRT` / `*.STEP` 等，于是 `测试用例/` 下的 **zip、PNG、JPG、CSV** 一路绿灯，
7 个文件 4.1 MB 进了公开仓库——而 README 写着"仓库不含测试模型"。

改成**整树排除** `测试用例/`，再用 `git rm --cached -r` 移出索引。
**不改写历史**（`filter-repo` 要 force push），本地文件一个没动。

### smoke 缺模型时只说"missing"

clone 下来跑 `python tests/smoke.py` 得到一行 `missing test model: ...` 和 `exit 2`，
没有任何下一步。现在会列出所需的三样东西（零件 / 装配体 / 钣金件）并给出
`--part` / `--assembly` 的用法示例——仓库本来就不带模型，能指到本机模型就够了。

### `call_method` 有两份

`bom_native.py` 自己写了一份，`sw_session.py` 一份。四轮反馈后才统一——
代价是**重试与错误分类逻辑只长在 `sw_session` 那份上**，`bom_native` 的那条路径
永远拿不到桥接自愈。现在只留一处。

### `health_check` 不说自己在跑哪个包

多份 editable 安装并存时，"MCP 连的是哪个版本"是唯一重要的信息。补
`package_path` / `package_version` / `python_executable`。

### `static_check` 只看 `name.attribute`

裸名字（模块级常量、缓存）不在它的视野里。`_EDGE_CACHE` 被一次合并吞掉后，
模块仍然 import 成功，只有跑那条路径才会炸——**秒级防线当时给的是绿灯**。

现在 `check_bare_names()` 遍历每个模块的顶层函数，把 `ast.Name`（Load）逐个对
`imports ∪ 模块级定义 ∪ 参数 ∪ 局部 ∪ 内建 ∪ dunder` 求解。第一版漏了**嵌套函数的参数**，
99 个误报；补上 `ast.arg` 与 `global` 后归零。

## 三十三、可标注性可以算出来，不用试出来（0.8.39）

反馈：外形视图 13 条圆边，`add_hole_callout` 3/3 全拒；端视图 2/10、剖视 4/9 可标。
规律很明显，但**清单里没有任何可标注性信号**，调用方只能逐条试。

### 判据是几何的

一个圆能不能打孔标注，取决于**它的法向是否沿视图的观察方向**：

```python
axis_vec = CircleParams[3:6]                    # 圆的法向
depth_axis = 3 - axes[0] - axes[1]              # 视图显示两个轴，剩下那个就是观察方向
dominant = max(range(3), key=lambda i: abs(axis_vec[i]))
faces_viewer = dominant == depth_axis and abs(axis_vec[depth_axis]) >= 0.9
```

`axes` 来自既有的 `_view_axis_map`，所以**零额外 API 成本**，不碰 COM。

### 实测完全对齐

门板 `door.SLDPRT`：

| 圆 | `facing` | `annotatable_hint` | 实际 |
|---|---|---|---|
| Ø6 × 4 | `side` | `false` | 单点与批量**都标不出** |
| Ø8 × 2 | `front` | `true` | 5/5 成功 |
| Ø15 × 3 | `front` | `true` | 5/5 成功 |

9 条圆边，预判 9 条全对。调用方现在可以在清单阶段直接过滤。

### 顺带：`IView.Type` 的 7 是标准模型视图

初版 `_view_kind` 按文档猜了 `{1: sheet, 2: section, 3: detail, 4: projected}`，
结果每条都显示 `type_7`——测下来**图纸是 1、标准模型视图是 7**。补上后正常。

猜枚举值之前先打一次日志，比事后看 `type_7` 猜半天快。

## 三十四、孤儿尺寸：`GetAttachedEntities2` 为空不等于没引线（0.8.40）

`SetPosition` 移过的尺寸，`GetAttachedEntities2` 返回空——被记成 orphan。
`prune_orphan_dimensions` 的 caveat 早就写明"这些其实还有引线，照单全删会缺料"，
但 `delivery_check` 照旧把它们计为 blocking，verdict 永远到不了「可制造」。

**同一张图，两个工具给出相反的处理指引**，用户无所适从。这是第三次报告。

### 用几何做仲裁

```python
gap_x = max(box[0] - position[0], 0.0, position[0] - box[2])
gap_y = max(box[1] - position[1], 0.0, position[1] - box[3])
near_leader = ((gap_x ** 2 + gap_y ** 2) ** 0.5) <= 0.05     # 50 mm
```

活着尺寸离它标注的视图不远（`layout_dimensions` 就停放在轮廓外侧），
真死的会飘得很远。50 mm 是经验阈值：够容纳 `outside_gap + slot × step` 的堆叠，
又远小于跨图纸的漂移。

实测门板：5 个 orphan **全部** `orphan_with_leader: true`，
`dimension_attachments ok:true`、`blockers: []`、verdict 回到「待设计审核」。

回执里会写明"另有 5 条移动过的尺寸仍有引线，不计入"——**豁免要说出来**，
否则下次没人知道为什么 5 个 orphan 没扣分。

## 三十五、一次"修好了"的自我推翻：50 mm 不等于有引线（0.8.41）

0.8.40 我加了「orphan 到视图 outline 距离 ≤50 mm 就算还有引线，不计入 blocker」。
理由看似充分：`prune_orphan_dimensions` 的 caveat 说这些尺寸其实还有引线。

**错在哪**：从 v0.8.27 起，尺寸**默认就被摆到视图外侧**（"place dimensions outside
their view"）。也就是说，一张排得越整齐的图，每个尺寸离 outline 都越近——
这个判据实际衡量的是「**尺寸放在哪**」，而不是「**尺寸有没有挂到几何**」。

结果在履带轮那张图上：**15/15 条真孤儿全部被豁免**，
`dimension_attachments` 从 fail 翻成 ok，而同一个返回体里还写着
「23 个尺寸均已关联」——实际是 8 条关联 + 15 条被豁免。

### 回退成诚实的计数

```python
orphan_near_view = ...          # 只作观测，不再叫 "with_leader"
orphan = len(all_orphans)       # 全部计入
suffix = "（其中 %d 条位置在视图附近，但引线未验证）" % near_view
```

`checks` 里同时给 `attached_count` 和 `excused_count(=0)`，让"N 个已关联"不再
和"另有 5 条不计入"打架。

实测门板：`ok: false`、`detail: 5/17 个尺寸无关联实体（其中 5 条位置在视图附近，但引线未验证）`、
`attached_count: 12`。

### 教训

**判据必须回答它声称回答的问题。**「位置在哪」和「挂没挂上几何」是两件事，
用一个便宜的相关量去替代，就会把失败的判定变成通过——这恰好是项目第一条原则
（"工具返回成功 ≠ 图纸正确"）要防的事。

真正的判据在 `IAnnotation.GetLeader` 那一侧（引线指向哪个实体），下一轮探。
探不到就老实标 `unknown`，**不要当成"有"**。
