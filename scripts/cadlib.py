# -*- coding: utf-8 -*-
"""CAD 图块库提炼工具（AutoCAD COM 通道）

把一张杂乱母库（几十上百个图层、成百上千个块定义）提炼成干净可复用的图块图库。

用法：
    python cadlib.py <母库.dwg> audit
    python cadlib.py <母库.dwg> all --keep "块A,块B,块C" --out "成品.dwg"
    python cadlib.py <母库.dwg> clean
    python cadlib.py <母库.dwg> merge --merge-layers all
    python cadlib.py <母库.dwg> rebase --keep "块A,块B"
    python cadlib.py <母库.dwg> slim --keep "块A,块B"
    python cadlib.py <母库.dwg> layout --keep "块A,块B" --cols 4

动作：
    audit    只读盘点：图层、实体分布、块清单、图层引用关系。不改文件。
    clean    删除零引用图层（保留 0 / Defpoints 及被引用的层）。
    merge    把指定图层的实体并入 0 层后删层。ByLayer 颜色会先固化为显式 ACI，观感不变。
    rebase   把每个目标块内的几何整体平移到包围盒居中，使插入点即图例中心。
    slim     只保留目标块及其递归嵌套依赖，删除其余块定义。
    layout   清空模型空间，把目标块按网格排列展示。
    all      依次 clean -> merge -> rebase -> slim -> layout。

安全约定：
    默认先 SaveAs 到 --out（缺省 <原名>_lib.dwg），全部操作在副本上，原件不动。
    加 --inplace 才直接改当前打开的文档（自己负责备份）。
"""

import argparse
import json
import os
import sys
import time

SYSTEM_BLOCKS = {"*Model_Space", "*Paper_Space", "*Paper_Space0"}
# 只死保 0 层。Defpoints 在无标注点时是可以删掉的，
# 所以让它走正常判定：有对象引用就留，没人用就删。
ALWAYS_KEEP_LAYERS = {"0"}


# ---------------------------------------------------------------- COM 基础

def get_com():
    try:
        import win32com.client as w
    except ImportError:
        sys.exit("[FAIL] 缺 pywin32，先装：\n"
                 "  python -m pip install pywin32==312 "
                 "-i https://mirrors.aliyun.com/pypi/simple "
                 "--trusted-host mirrors.aliyun.com --no-deps")
    try:
        import pythoncom
        from win32com.client import VARIANT
    except ImportError:
        sys.exit("[FAIL] pywin32 不完整，重装一次")
    return w, pythoncom, VARIANT


def get_app(w):
    try:
        return w.GetActiveObject("AutoCAD.Application")
    except Exception:
        sys.exit("[FAIL] 没有运行中的 AutoCAD。请先打开一扇 AutoCAD 窗口，"
                 "载入任意图纸后再跑本脚本（COM 只挂已开实例，不另起进程）。")


def point(VARIANT, pythoncom, x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def safe(fn, tries=40, delay=1.0):
    """AutoCAD 忙时任何 COM 调用都可能抛 RPC_E_CALL_REJECTED(-2147418111)。

    这类错误是暂时性的，重试即可；其它异常照常抛出。
    """
    last = None
    for _ in range(tries):
        try:
            return fn()
        except Exception as ex:
            s = str(ex)
            if ("-2147418111" in s or "拒绝接收" in s or "被呼叫方拒绝" in s
                    or "RPC" in s):
                last = ex
                time.sleep(delay)
                continue
            raise
    raise last


def wait_ready(app, name, timeout=120):
    """等待刚打开的文档真正就绪（大文件加载要几十秒），期间 COM 会拒答。"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            d = app.ActiveDocument
            if d is not None and d.Name and d.Name.upper() == name.upper():
                safe(lambda: [L.Name for L in d.Layers], tries=3, delay=0.3)
                return d
        except Exception:
            pass
        time.sleep(1.5)
    raise RuntimeError("AutoCAD 在 %ds 内未就绪，请确认图纸已加载完成" % timeout)


def ents_of(blk):
    """取块内实体列表；CAD 忙时整体重试。"""
    return safe(lambda: list(blk))



def block_entry(doc, name):
    """按名字取块定义。必须迭代，doc.Blocks.Item(int) 会抛 COM 异常。"""
    for b in safe(lambda: list(doc.Blocks)):
        if b.Name == name:
            return b
    return None


def block_names(doc):
    return [b.Name for b in safe(lambda: list(doc.Blocks))]


def block_bbox(blk):
    """块内实体合并包围盒；部分实体的 GetBoundingBox 会抛错，跳过。"""
    xs, ys = [], []
    for e in ents_of(blk):
        try:
            mn, mx = e.GetBoundingBox()
            xs += [mn[0], mx[0]]
            ys += [mn[1], mx[1]]
        except Exception:
            pass
    if not xs:
        return None
    return min(xs), max(xs), min(ys), max(ys)


def layer_usage(doc):
    """统计每个图层被哪些容器引用：模型空间 / 图纸空间 / 具体块。"""
    used = {}

    def scan(coll, tag):
        for e in ents_of(coll):
            try:
                used.setdefault(e.Layer, set()).add(tag)
            except Exception:
                pass

    scan(doc.ModelSpace, "模型空间")
    scan(doc.PaperSpace, "图纸空间")
    for b in safe(lambda: list(doc.Blocks)):
        if b.Name.startswith("*"):
            continue
        scan(b, "块:" + b.Name)
    return used



# ---------------------------------------------------------------- 动作实现

def act_audit(doc, args):
    layers = [(L.Name, L.color) for L in doc.Layers]
    used = layer_usage(doc)
    blocks = []
    for b in safe(lambda: list(doc.Blocks)):
        if b.Name.startswith("*"):
            continue
        bb = block_bbox(b)
        blocks.append({
            "name": b.Name,
            "count": b.Count,
            "w": round(bb[1] - bb[0], 2) if bb else None,
            "h": round(bb[3] - bb[2], 2) if bb else None,
            "cx": round((bb[0] + bb[1]) / 2, 2) if bb else None,
            "cy": round((bb[2] + bb[3]) / 2, 2) if bb else None,
        })
    free = [n for n, _ in layers
            if n not in used and n.lower() not in ALWAYS_KEEP_LAYERS]
    rep = {
        "doc": doc.Name,
        "layers_total": len(layers),
        "layers_used": {k: sorted(v) for k, v in used.items()},
        "layers_free": free,
        "blocks_total": len(blocks),
        "blocks": blocks,
        "modelspace_count": doc.ModelSpace.Count,
    }
    print("=" * 66)
    print("文档: %s" % doc.Name)
    print("图层 %d 个，其中 %d 个零引用（可删）" % (len(layers), len(free)))
    for name, color in layers:
        who = used.get(name)
        tag = "  空" if not who else "  被引用: " + ",".join(sorted(who)[:3])
        print("  %-22s 色%-4s%s" % (name, color, tag))
    print("块定义 %d 个（不含系统块），模型空间实体 %d"
          % (len(blocks), doc.ModelSpace.Count))
    print("=" * 66)
    return rep


def act_clean(doc, args):
    used = layer_usage(doc)
    removed, kept = [], []
    for L in safe(lambda: list(doc.Layers)):
        name = L.Name
        if name.lower() in ALWAYS_KEEP_LAYERS or name in used:
            kept.append(name)
            continue
        try:
            safe(lambda: L.Delete(), tries=5, delay=0.3)
            removed.append(name)
        except Exception:
            kept.append(name)
    print("[clean] 删除零引用图层 %d 个，保留 %d 个" % (len(removed), len(kept)))
    if removed:
        print("        已删: " + ", ".join(removed))
    return {"removed": removed, "kept": kept}


def act_merge(doc, args):
    wanted = args.merge_layers
    if wanted == "all":
        targets = [L.Name for L in doc.Layers
                   if L.Name.lower() not in ALWAYS_KEEP_LAYERS]
    else:
        targets = [x.strip() for x in (wanted or "").split(",") if x.strip()]
    if not targets:
        print("[merge] 未指定图层，跳过")
        return {"moved": 0}

    aci = {}
    for L in doc.Layers:
        aci[L.Name] = L.color

    moved = fixed_color = 0

    def move_entities(coll):
        nonlocal moved, fixed_color
        for e in ents_of(coll):
            try:
                lay = e.Layer
            except Exception:
                continue
            if lay not in targets:
                continue
            # ByLayer(256) 直接改层会跟着 0 层变色，先固化为原层显式色
            try:
                if e.color == 256:
                    e.color = aci.get(lay, 7)
                    fixed_color += 1
                e.Layer = "0"
                moved += 1
            except Exception:
                pass

    move_entities(doc.ModelSpace)
    move_entities(doc.PaperSpace)
    for b in safe(lambda: list(doc.Blocks)):
        if b.Name.startswith("*"):
            continue
        move_entities(b)

    removed, failed = [], []
    for nm in targets:
        ok_del = False
        last_err = ""
        for _ in range(3):
            try:
                safe(lambda: doc.Layers.Item(nm).Delete(), tries=8, delay=0.4)
                ok_del = True
                break
            except Exception as ex:
                last_err = str(ex)[:60]
                # 数据库引用尚未释放，保存一次刷库再来
                try:
                    doc.Save()
                    time.sleep(1.5)
                except Exception:
                    pass
        (removed if ok_del else failed).append(nm if ok_del else (nm, last_err))
    print("[merge] 实体并入 0 层 %d 个（其中固化 ByLayer 色 %d 个），删除图层 %d 个"
          % (moved, fixed_color, len(removed)))
    for nm, why in failed:
        print("        图层未删: %s -> %s" % (nm, why))
    print("        剩余图层: " + ", ".join(L.Name for L in doc.Layers))
    return {"moved": moved, "fixed_color": fixed_color,
            "removed": removed, "failed": failed}


def act_rebase(doc, args):
    keep = parse_keep(args.keep)
    if not keep:
        print("[rebase] 未指定 --keep，跳过")
        return {}
    VARIANT, pythoncom = args._VARIANT, args._pythoncom
    o = point(VARIANT, pythoncom, 0, 0)
    done, skipped = [], []
    for name in keep:
        b = block_entry(doc, name)
        if b is None:
            skipped.append(name)
            continue
        moved = 0
        for _ in range(3):  # 迭代收敛，缩放/嵌套块可能一次不到位
            bb = block_bbox(b)
            if bb is None:
                break
            cx, cy = (bb[0] + bb[1]) / 2, (bb[2] + bb[3]) / 2
            if abs(cx) < 0.5 and abs(cy) < 0.5:
                break
            p = point(VARIANT, pythoncom, -cx, -cy)
            for e in ents_of(b):
                try:
                    e.Move(o, p)
                    moved += 1
                except Exception:
                    pass
        done.append({"name": name, "moved": moved})
        print("  [rebase] %-18s 平移 %d 个实体" % (name, moved))
    print("[rebase] 完成 %d 个块，缺失 %d 个" % (len(done), len(skipped)))
    return {"done": done, "missing": skipped}


def resolve_deps(doc, keep):
    """目标块 + 递归嵌套依赖，这些必须保留，否则父块会残缺。"""
    keepset = list(keep)
    seen = set(keepset)
    idx = 0
    while idx < len(keepset):
        b = block_entry(doc, keepset[idx])
        idx += 1
        if b is None:
            continue
        for e in ents_of(b):
            try:
                if e.ObjectName == "AcDbBlockReference" and e.Name not in seen:
                    seen.add(e.Name)
                    keepset.append(e.Name)
            except Exception:
                pass
    return keepset


def act_slim(doc, args):
    keep = parse_keep(args.keep)
    if not keep:
        print("[slim] 未指定 --keep，跳过")
        return {}
    ms = safe(lambda: doc.ModelSpace.Count)
    if ms:
        print("[slim] 注意：模型空间还有 %d 个实体，其中的块参照会让对应块定义删不掉；"
              "先清空模型空间（all 会自动做）" % ms)
    needed = resolve_deps(doc, keep)
    nested = [x for x in needed if x not in keep]
    if nested:
        print("[slim] 检测到嵌套依赖，一并保留: " + ", ".join(nested))
    kill = [n for n in block_names(doc)
            if n not in needed and n not in SYSTEM_BLOCKS]
    ok, fail = 0, []
    pending = [(n, "") for n in kill]
    for _ in range(3):
        if not pending:
            break
        still = []
        for nm, _msg in pending:
            try:
                safe(lambda: doc.Blocks.Item(nm).Delete(), tries=8, delay=0.4)
                ok += 1
            except Exception as ex:
                still.append((nm, str(ex)[:60]))
        if len(still) == len(pending):
            # 一个都没删动，说明数据库引用未释放，保存刷库后重来
            try:
                doc.Save()
                time.sleep(2.0)
            except Exception:
                break
        pending = still
    fail = pending
    print("[slim] 删除块定义 %d 个，失败 %d 个" % (ok, len(fail)))
    for nm, why in fail[:3]:
        print("        未删: %s -> %s" % (nm, why))
    if len(fail) > 3:
        print("        ...另有 %d 个" % (len(fail) - 3))
    print("        保留: " + ", ".join(needed))
    return {"deleted": ok,
            "failed": [f[0] for f in fail],
            "kept": needed, "nested": nested}


def act_layout(doc, args):
    keep = parse_keep(args.keep)
    if not keep:
        print("[layout] 未指定 --keep，跳过")
        return {}
    VARIANT, pythoncom = args._VARIANT, args._pythoncom
    for e in ents_of(doc.ModelSpace):
        e.Delete()
    info = []
    for name in keep:
        b = block_entry(doc, name)
        bb = block_bbox(b) if b is not None else None
        if bb is None:
            continue
        info.append((name, bb[1] - bb[0], bb[3] - bb[2]))
    if not info:
        return {}
    cw = max(i[1] for i in info) * 1.30
    ch = max(i[2] for i in info) * 1.55
    cols = max(1, args.cols)
    for idx, (name, w, h) in enumerate(info):
        x = (idx % cols) * cw
        y = -(idx // cols) * ch
        try:
            doc.ModelSpace.InsertBlock(point(VARIANT, pythoncom, x, y),
                                      name, 1.0, 1.0, 1.0, 0.0)
        except Exception as ex:
            print("  [layout] 插入失败 %s: %s" % (name, str(ex)[:50]))
    try:
        args._app.ZoomExtents()
    except Exception:
        pass
    print("[layout] 已排列 %d 个块（%d 列，格 %.0f x %.0f）"
          % (len(info), cols, cw, ch))
    return {"count": len(info), "cols": cols, "cell": [round(cw), round(ch)]}


def clear_modelspace(doc):
    """清空模型空间。图库提炼场景下模型空间内容是要丢弃的平面图；
    更关键的是块参照会占住块定义，不清空就永远删不掉多余块。"""
    n = 0
    for e in ents_of(doc.ModelSpace):
        try:
            e.Delete()
            n += 1
        except Exception:
            pass
    return n


def parse_keep(s):
    if not s:
        return []
    return [x.strip() for x in s.replace("，", ",").split(",") if x.strip()]


# ---------------------------------------------------------------- 主流程

ACTIONS = {
    "audit": act_audit,
    "clean": act_clean,
    "merge": act_merge,
    "rebase": act_rebase,
    "slim": act_slim,
    "layout": act_layout,
}


def main():
    ap = argparse.ArgumentParser(description="CAD 图块库提炼工具")
    ap.add_argument("doc", help="母库 DWG 完整路径")
    ap.add_argument("action", choices=list(ACTIONS) + ["all"])
    ap.add_argument("--keep", help="要保留的块名，逗号分隔")
    ap.add_argument("--merge-layers", default="",
                    help="要并入 0 层的图层，逗号分隔；填 all 表示除 0/Defpoints 外全部")
    ap.add_argument("--out", help="输出 DWG 路径，缺省 <原名>_lib.dwg")
    ap.add_argument("--inplace", action="store_true",
                    help="直接改当前文档，不另存副本（自己负责备份）")
    ap.add_argument("--cols", type=int, default=4, help="layout 列数，默认 4")
    ap.add_argument("--keep-ms", action="store_true",
                    help="保留模型空间原有内容（默认 all 会先清空，否则块参照占住块定义删不掉）")
    ap.add_argument("--report", help="把结果 JSON 写到该路径")
    args = ap.parse_args()

    w, pythoncom, VARIANT = get_com()
    args._VARIANT, args._pythoncom = VARIANT, pythoncom
    app = get_app(w)
    args._app = app

    src = os.path.abspath(args.doc)
    if not os.path.exists(src):
        sys.exit("[FAIL] 文件不存在: %s" % src)

    # AutoCAD 允许零文档（全关掉的空实例），此时 ActiveDocument 是空对象
    cur = None
    try:
        cur = safe(lambda: app.ActiveDocument, tries=3, delay=0.3)
    except Exception:
        cur = None

    need_open = True
    if cur is not None:
        try:
            need_open = (os.path.normcase(os.path.abspath(cur.FullName))
                         != os.path.normcase(src))
        except Exception:
            need_open = True

    if need_open:
        print("[open] 打开 %s" % src)
        safe(lambda: app.Documents.Open(src))
        wait_ready(app, os.path.basename(src))
        print("[open] 加载完成")
    else:
        wait_ready(app, os.path.basename(src), timeout=30)

    if not args.inplace and args.action != "audit":
        out = args.out or os.path.splitext(src)[0] + "_lib.dwg"
        out = os.path.abspath(out)
        safe(lambda: app.ActiveDocument.SaveAs(out))
        wait_ready(app, os.path.basename(out))
        print("[copy] 已另存副本: %s（原件未改动）" % out)

    doc = safe(lambda: app.ActiveDocument)
    result = {}

    if args.action == "all":
        result["audit"] = act_audit(doc, args)
        if not args.keep_ms:
            n = clear_modelspace(doc)
            print("[clear] 清空模型空间 %d 个实体"
                  "（释放块参照，否则块定义永远删不掉）" % n)
        result["clean"] = act_clean(doc, args)
        result["merge"] = act_merge(doc, args)
        result["rebase"] = act_rebase(doc, args)
        result["slim"] = act_slim(doc, args)
        # AutoCAD 删除图层/块后，数据库引用要过几秒才释放，
        # 首次跑常留下个别"N 个对象被参照"。等一下再复清一遍。
        print("[收尾] 等待数据库引用释放后复清残留...")
        time.sleep(3.0)
        doc = safe(lambda: app.ActiveDocument)
        result["clean2"] = act_clean(doc, args)
        result["slim2"] = act_slim(doc, args)
        doc = safe(lambda: app.ActiveDocument)
        result["layout"] = act_layout(doc, args)
    else:
        result = ACTIONS[args.action](doc, args)

    if args.action != "audit":
        try:
            safe(lambda: app.ActiveDocument.Save())
            print("[save] 已保存: %s" % safe(lambda: app.ActiveDocument.FullName))
        except Exception as ex:
            print("[save] 保存失败: %s" % str(ex)[:80])

    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=1)
        print("[report] %s" % args.report)


if __name__ == "__main__":
    main()
