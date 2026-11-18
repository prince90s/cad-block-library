# AutoCAD COM 操作实测踩坑清单

全部结论来自 Windows + AutoCAD 2023 (24.2s) + Python 3.13 + pywin32 312 的实测，
处理对象是一张 40 图层 / 180 块定义 / 25 MB 的弱电施工图母库。

## 1. 环境：pywin32 装不上

**现象**：`pip install pywin32` 报 `Could not find a version that satisfies the requirement pywin32 (from versions: none)`。

**原因**：pip 全局源指向清华镜像 `pypi.tuna.tsinghua.edu.cn`，该镜像没有 pywin32 这个包（不是限流，是索引里确实没有）。

**解法**：换阿里云或腾讯云镜像。

```bash
python -m pip install pywin32==312 -i https://mirrors.aliyun.com/pypi/simple --trusted-host mirrors.aliyun.com --no-deps
```

直连 `pypi.org` 在沙箱环境会被 SIGTERM，不要用。Python 3.13 对应 pywin32 312。

## 2. 连接：`GetActiveObject` 报"操作无法使用"

**现象**：`GetActiveObject('AutoCAD.Application')` 抛 `(-2147221021, '操作无法使用')`。

**含义**：**没有运行中的 AutoCAD 实例**，不是权限被拒，也不是 COM 不可用。

**解法**：先打开 AutoCAD（不必打开图纸也能连，但多数操作要有文档）。运行实例起来后，
裸名 `AutoCAD.Application` 即可连上，即使注册表里只登记了带版本号的
`AutoCAD.Application.24.1 / 24.2 / 25.1`。

**附加发现**：AutoCAD 允许零文档运行，此时 `app.ActiveDocument` 是空对象，
任何 `cur.FullName` 访问都会崩。脚本必须先判空再决定是否 `Documents.Open`。

## 3. 最容易踩的坑：改块基点，`entmod` 是假的

**错误做法**（看起来成功，实际无效）：

```lisp
(setq br (tblobjname "BLOCK" name))
(setq ed (entget br))
(setq ed (subst (list 10 cx cy 0.0) (assoc 10 ed) ed))
(entmod ed)   ; 返回正常，13 个块全报 OK
```

改完插入 `(0,0)`，图形中心仍在原坐标（如 `985733, -324184`），**基点纹丝不动**。
BLOCK_RECORD 的 10 组码在 AutoCAD 里并不等价于块基点。

**正确做法**：基点保持 `(0,0,0)`，把**块内几何整体平移到包围盒居中**。

```python
mn, mx = entity.GetBoundingBox()          # 合并所有块内实体求包围盒
cx, cy = (mn[0]+mx[0])/2, (mn[1]+mx[1])/2
for e in block:
    e.Move(point(0,0), point(-cx,-cy))    # ActiveX Move
```

效果等价：插入点即图例中心。实测 13 个块 815 个实体，0 失败，偏差全部 < 1。

**注意**：缩放过的块要迭代收敛（缩放会引入中心偏移），脚本用最多 3 轮
「算包围盒 → 若未居中再平移」保证归零。

## 4. 陷阱：`-PURGE` 会卡死 CAD 主线程

**现象**：通过 `SendCommand` 发 `_.-PURGE\n_A\n*\n_Y\n_N\n`，命令提示序列只要有一项对不上，
AutoCAD 就停在等待输入状态。此后**主线程被阻塞，所有 COM 调用失效**：

```
pywintypes.com_error: (-2147417848, '被调用的对象已与其客户端断开连接')
AttributeError: <unknown>.Name
```

只能人工按 ESC 才能救回来。

**替代方案**：清块定义直接走 COM。

```python
for name in kill_list:
    doc.Blocks.Item(name).Delete()
```

实测 163 个块一次过。**任何需要交互提示的 CAD 命令都不要用 `SendCommand` 发**，
改用 ActiveX 方法。真要发命令，只发无交互的单条表达式。

## 5. 集合索引：`Item(int)` 会抛异常

**现象**：`doc.Layers.Item(1)` 抛 `(-2147352567, '发生意外。')`。

**解法**：一律迭代，不按整数索引。

```python
layers = [L.Name for L in doc.Layers]      # 正确
# doc.Layers.Item(0) / Item(1)             # 别用
```

按名字查找也不可靠地依赖 `Item(name)` 对某些集合，遍历匹配最稳。
`doc.Blocks.Item(name)` 对具名块可用，对 `*U` 匿名块也可用。

## 6. 时序：删除报"对象被参照"，但其实没被引用

**现象**：图层/块已经确认无任何对象引用（模型空间、图纸空间、所有块内都扫过），
`Delete()` 仍报 `'对象被其他对象参照'`。等几分钟再删就成功了。

**原因**：AutoCAD 删除操作后，内部数据库引用释放有延迟。

**解法**：多轮复清 + 中途保存刷库。

- 单次删除失败：`doc.Save()` 后 sleep 1.5~2 s 重试，最多 3 轮。
- 整体流程：`all` 在 `slim` 之后额外跑一次 `clean + slim`（等 3 s），
  实测能把首轮漏掉的 1 个图层和 1 个块补删掉。

## 7. RPC_E_CALL_REJECTED：大文件加载期间 COM 被拒

**现象**：`(-2147418111, '被呼叫方拒绝接收呼叫')`。多发生在 `Documents.Open` 一张 25 MB 母库之后。

**解法**：两层防护。

1. **就绪等待**：打开后轮询 `app.ActiveDocument.Name` 是否匹配目标，且能成功读到图层表，
   超时 120 s。大文件加载要几十秒。
2. **调用重试**：把所有可能抛 `-2147418111` / "拒绝接收" / "RPC" 的调用包一层重试。

## 8. `ZoomExtents` 属于 Application，不属于 Document

```python
app.ZoomExtents()      # 正确
doc.ZoomExtents()      # AttributeError: <unknown>.ZoomExtents
```

且调用后手里那个 `doc` 引用会失效，**必须重新 `app.ActiveDocument` 取一次**。
同类问题：`doc.Close()` 之后访问 `doc.Name` 也会断开，要提前把名字存成字符串。

## 9. 嵌套块不能当垃圾删

`8+2口POE交换机` 内部引用了 `8口交换机`。删子块会报「对象被参照」。
必须递归解析目标块的嵌套依赖，把子块一并纳入保留名单。

```python
# 判断块内是否有块参照
if e.ObjectName == "AcDbBlockReference" and e.Name not in seen:
    ...
```

## 10. 模型空间必须先清空，否则块永远删不掉

模型空间里只要有引用某块的 `INSERT`，该块定义就是"被参照"，`Delete()` 必然失败。
**正确顺序：先清空模型空间 → 再精简块定义**。

这一点决定了 `all` 的动作顺序，也是最容易设计错的地方——初次实现时把 `layout`
（负责清空）放在 `slim` 后面，导致 163 个块全部删除失败。

## 11. 改图层会让 ByLayer 颜色变掉

实体颜色是 `256`（ByLayer）时，`e.Layer = "0"` 后颜色跟随 0 层变成白色/黑色。
若原图层是青色(4)、绿色(3)，图例观感就变了。

**解法**：改层前先读原图层 ACI 并固化到实体。

```python
if e.color == 256:
    e.color = layer_color_aci[old_layer]
e.Layer = "0"
```

若实体本身就是显式 ACI（如 3、7），直接改层即可，颜色不变。
判断依据：`e.color == 256` 才是 ByLayer。

## 12. `Defpoints` 并非绝对不可删

普遍说法是 `Defpoints` 是系统层不能删。实测：**图上没有标注点关联时，它可以删掉**。
只有存在尺寸标注引用它时才删不掉。所以不必写死保护，让它走"有引用才留"的正常判定。

`0` 层才是真正删不掉的（任何图纸至少有一个图层）。
