# 开发约定

这个项目的坑大多不是"不懂 API"，而是**改动没有立刻验证**。下面几条是从真实事故里长出来的。

---

## 一、改完必须跑一次真实调用

### 规则

改了任何工具函数后，**提交前**至少跑一次真实调用，不要只看 `compileall` 和 `static_check`。

```bash
python tests/check_all.py     # 秒级，无需 SolidWorks（三层）
python tests/smoke.py         # 约 18 分钟，需 SolidWorks（行为层，32 项）
```

`check_all.py` 依次跑：

| 层 | 文件 | 挡什么 |
|---|---|---|
| 1 语法 | `compileall` | 文件根本编不过 |
| 2 结构 | `static_check.py` | 未导入名字、README 漂移、重复 def、属性陷阱、路线图过期、包装器签名 |
| 3 **历史回归** | `regressions.py` | **每个已修复的 bug 一条断言（58 条，含 smoke 调用签名契约、残留文档自愈与 0.8.26 同名尺寸消歧 / 标注文本校验）** |

只改了一个工具时，跑它对应的那一项就够（照 `tests/smoke.py` 里的写法抄一个临时脚本）。

### 为什么

`static_check.py` 只验证结构：名字有没有导入、README 是否同步、有没有重复 `def`、属性派发陷阱、路线图过期。**它不知道函数能不能跑。**

三次真实事故，静态检查全绿而代码是坏的：

| # | 事故 | 静态检查为什么放行 |
|---|---|---|
| 1 | `row["text"]` 写在 `row = {...}` 之前 | 语法合法；`row` 在别的分支有赋值 |
| 2 | 同一文件两个 `def add_bom_table`，旧的覆盖新的 | 已加第四道防线，但当时还没加 |
| 3 | `seen_dims` 去重块插进了**错误的函数** | `seen_dims` 在另一个函数里存在，"名字有定义"检查通过 |

第 3 条是最典型的：**字符串替换命中了第一个匹配**（`rows, counts = [], {}` 在两个函数里都出现），结果把要修的 P0 又复现了一遍，花了三轮才改对。

**如果改完立刻跑一次 `list_drawing_annotations`，5 秒就能发现。**

---

## 二、改代码不要用字符串替换定位

### 规则

用 `s.replace(...)` 批量改代码时，锚点字符串**必须在文件里唯一**。

```python
# 危险：rows, counts = [], {} 在两个函数里都出现
s = s.replace('        rows, counts = [], {}\n        orphans = []', ...)

# 安全：先切函数再改
idx = s.find("def verify_dimension_attachments")
head, tail = s[:idx], s[idx:]
tail = tail.replace(...)
```

改完用 `grep -n <关键标识>` 确认它只在预期位置出现。

### 为什么

这是事故 #3 的直接原因，也是事故 #2（重复函数）的成因。两次都发生在同一天。

---

## 三、宁可 `not_checked`，不要错数

### 规则

一项检查如果没真正执行，返回体里必须写 `not_checked` 并在 `reason` 里说明**试过什么、卡在哪**。不许因为"看起来应该没问题"就给绿勾。

```python
add("min_flange_length", "not_checked",
    "需测量折弯边长度（4t+R）。圆柱面几何能取到，但折弯圆角与孔的分离不可靠"
    "（半径差=板厚会把 Ø12/Ø14 同心孔误判为折弯），故不做判定")
```

### 为什么

`sheet_metal_qc` 最早只有板厚弯折数，五项几何检查全部 `not_checked` 并写清原因。用户反馈的评价是：

> 五项 `not_checked` 的说明写得很清楚（含"试过面相邻性判据"、"遍历 80 个圆柱面需 200 s"），这部分体验是好的。

**一个诚实的"没检查"比一个假的"通过"有用得多。**

---

## 四、工具返回成功 ≠ 图纸正确

### 规则

每个写操作都要**回读验证**，不能靠返回值判定。

```python
# ❌ RunMacro 成功时也返回 False
ok = sw.RunMacro(macro, "Macro1", "main")

# ✅ 回读表格数
before = len(list(doc.GetFirstView.GetTableAnnotations or ()))
sw.RunMacro(macro, "Macro1", "main")
after = len(list(doc.GetFirstView.GetTableAnnotations or ()))
created = len(after) > len(before)
```

已知会骗人的返回值：

| 调用 | 骗法 |
|---|---|
| `ISldWorks.RunMacro` | 成功时返回 `False` |
| `IView.InsertBomTable4` | 返回 `None` 有时是成功、有时是失败 |
| `doc.EditRebuild3` | 是**属性**，取值不触发重建 |
| `sw.GetExecutablePath` | 是**属性**，加括号报 `'str' object is not callable` |

---

## 五、COM 方法必须走 `call_method`

### 规则

调用 SolidWorks 方法一律用 `sw_session.call_method(obj, "Name", ...)`，不要直接 `obj.Name(...)`。

```python
from ..sw_session import call_method

# ✅
call_method(doc, "ForceRebuild3", False)
call_method(edge, "GetStartVertex")

# ❌ 可能被当成属性派发，报 DISP_E_MEMBERNOTFOUND
doc.ForceRebuild3(False)
```

`static_check` 有一道防线会拦住清单里的直接调用。清单在 `sw_session.PROPERTY_TRAP_METHODS`。

### 性能上的额外收益

`GetStartVertex` 直接访问 **335 s**（1094 条边），走 `call_method` **179 s**。同样一次调用，差一倍。

---

## 六、能力表跟着代码走

### 规则

新增/修复工具后，同步三处：`README.md` 工具表、`tools/capabilities.py`、`docs/ROADMAP.md`。

`static_check` 会校验 README 的工具计数和工具名覆盖，但**抓不到能力描述是否过期**——那条靠自觉。

---

## 七、文档里的数字必须是实测值

### 规则

写"335 s → 47 s"就要能指出测量方法。写"10 行 4 列"就要有回读输出。

不确定的标注不确定：

```markdown
> 说明：该枚举的官方页面在 WebFetch 时只返回目录，上述数值来自枚举命名惯例与检索结果，
> **未从官方页面逐字确认**。
```

### 为什么

BOM 那次我一度在反馈文档里写「3 类实现，本工具只碰第一类末端」，这种"看起来专业"的表述如果没有实测支撑，比不写更糟——**它让读者以为已经查过**。

---

## 现有防线一览

| # | 检查 | 挡住什么 | 抓不到什么 |
|---|---|---|---|
| 1 | `check()` 未导入名字 | `NameError` | 运行期逻辑错 |
| 2 | `check_readme()` | 工具计数/名称漂移 | 能力描述过期 |
| 3 | `check_property_traps()` | 属性派发陷阱直接调用 | 写在 `call_method` 里的错 |
| 4 | `check_duplicate_defs()` | 同作用域重复 `def` | 插错函数的代码 |
| 5 | `check_stale_roadmap()` | 「vX.Y 待做」而版本已过 | 其他文档矛盾 |
| 6 | **`tests/smoke.py`** | **所有运行期错误** | 需要 SolidWorks，不能进 CI |

**前五道是静态的，第六道才是真的。改完代码请跑第六道。**
