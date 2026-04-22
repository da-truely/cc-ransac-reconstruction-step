# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import math
import os
import re
import sys
from pathlib import Path

EXPORT_DIR = r""  # leer => Desktop, sonst Zielordner
EXPORT_NAME = "cc_ransac_export.json"
PRIMS = {"cone", "cylinder", "plane", "sphere", "torus"}


def emit(msg: str, prefix: str = "cc_ransac_to_json") -> None:
    msg = " ".join(str(msg).split())
    if msg:
        print(f"[{prefix}] {msg}", flush=True)


def warn(msg: str) -> None:
    emit(f"warn: {msg}")


def desktop_dir() -> Path:
    for p in (Path.home() / "Desktop", Path.home() / "Schreibtisch"):
        if p.exists():
            return p
    return Path.home()


def writable_dir(p) -> Path | None:
    if not p:
        return None
    try:
        p = Path(p).expanduser().resolve()
        if str(p).strip() in {"", "."}:
            return None
        p.mkdir(parents=True, exist_ok=True)
        t = p / ".__cc_write_test__"
        t.write_text("1", encoding="utf-8")
        t.unlink(missing_ok=True)
        return p
    except Exception:
        return None


def load_cc_module():
    for n in ("cloudComPy", "pycc", "cc", "cloudcompare"):
        try:
            return __import__(n)
        except Exception:
            pass
    raise SystemExit("CloudCompare-Modul nicht gefunden. Skript aus CloudCompare starten.")


def get_cc_app():
    for mod_name in ("pycc", "cc", "cloudcompare", "cloudComPy"):
        try:
            mod = __import__(mod_name)
        except Exception:
            continue
        for attr in ("GetInstance", "getInstance", "instance", "app", "Application", "CCApp"):
            cand = getattr(mod, attr, None)
            if cand is None:
                continue
            try:
                app = cand() if callable(cand) else cand
                if app is not None:
                    return app
            except Exception:
                pass
    for mod in list(sys.modules.values()):
        if mod is None:
            continue
        for attr in ("GetInstance", "getInstance", "instance", "app"):
            cand = getattr(mod, attr, None)
            if cand is None:
                continue
            try:
                app = cand() if callable(cand) else cand
                if app is not None:
                    return app
            except Exception:
                pass
    raise RuntimeError("Could not access the CloudCompare application instance.")


def as_list(v):
    return [float(v[0]), float(v[1]), float(v[2])]


def safe_name(o):
    for a in ("getName", "name", "getDisplayName", "displayName", "getFullPathName"):
        obj = getattr(o, a, None)
        if obj is None:
            continue
        try:
            v = obj() if callable(obj) else obj
        except Exception:
            continue
        if isinstance(v, str) and v.strip():
            return v.strip()
    return "<unnamed>"


def type_name(o):
    for a in ("getTypeName", "getClassName", "className", "typeName"):
        obj = getattr(o, a, None)
        if obj is None:
            continue
        try:
            v = obj() if callable(obj) else obj
        except Exception:
            continue
        if isinstance(v, str) and v.strip():
            s = v.strip()
            if any(k.lower() in s.lower() for k in PRIMS):
                return s
    try:
        if hasattr(cc, "CC_TYPES"):
            for k in ("PLANE", "SPHERE", "CYLINDER", "CONE", "TORUS"):
                if hasattr(cc.CC_TYPES, k) and o.isA(getattr(cc.CC_TYPES, k)):
                    return k.title()
    except Exception:
        pass
    n = o.__class__.__name__
    return n[2:] if n.lower().startswith("cc") else n


def norm_type(o):
    return type_name(o).strip().lower()


def is_primitive(o):
    return norm_type(o) in PRIMS


def unique_id(o):
    for a in ("getUniqueID", "getIndex"):
        obj = getattr(o, a, None)
        if obj is None:
            continue
        try:
            v = obj() if callable(obj) else obj
            if v is not None:
                return int(v)
        except Exception:
            pass
    try:
        return id(o)
    except Exception:
        return hash(repr(o))


def bbox_dims(o):
    bb = None
    for fn in ("getDisplayBB_recursive", "getOwnBB", "getBoundingBox"):
        if hasattr(o, fn):
            try:
                bb = getattr(o, fn)()
                if bb is not None:
                    break
            except Exception:
                bb = None
    if bb is None:
        return None
    try:
        if hasattr(bb, "minCorner"):
            mn, mx = bb.minCorner(), bb.maxCorner()
        else:
            mn, mx = bb[0], bb[1]
        return {"X": float(mx[0]) - float(mn[0]), "Y": float(mx[1]) - float(mn[1]), "Z": float(mx[2]) - float(mn[2])}
    except Exception:
        return None


def max_bbox_dims(o):
    dims = []
    seen = set()

    def add_dims(obj):
        if obj is None:
            return
        uid = unique_id(obj)
        if uid in seen:
            return
        seen.add(uid)
        d = bbox_dims(obj)
        if d:
            dims.append(d)

    cur = o
    for _ in range(64):
        add_dims(cur)
        if hasattr(cur, "getAssociatedCloud"):
            try:
                add_dims(cur.getAssociatedCloud())
            except Exception:
                pass
        for ch in iter_children(cur):
            add_dims(ch)
            if hasattr(ch, "getAssociatedCloud"):
                try:
                    add_dims(ch.getAssociatedCloud())
                except Exception:
                    pass
        parent = get_parent(cur)
        if parent is None:
            break
        cur = parent
    if not dims:
        return None
    return {"X": max(float(d.get("X", 0.0)) for d in dims), "Y": max(float(d.get("Y", 0.0)) for d in dims), "Z": max(float(d.get("Z", 0.0)) for d in dims)}


def torus_radii_from_name(o):
    s = safe_name(o)
    m = re.search(r"r\s*=\s*([0-9.+\-eE]+)\s*/\s*R\s*=\s*([0-9.+\-eE]+)", s)
    if not m:
        return None, None
    return float(m.group(1)), float(m.group(2))


def cone_half_angle_deg(bottom_r, top_r, height):
    if height in (None, 0):
        return None
    return math.degrees(math.atan(abs(float(bottom_r) - float(top_r)) / float(height)))


def get_parent(o):
    for a in ("getParent", "parent", "getFather", "father"):
        obj = getattr(o, a, None)
        if obj is None:
            continue
        try:
            v = obj() if callable(obj) else obj
        except Exception:
            continue
        if v is not None:
            return v
    return None


def iter_children(o):
    if o is None:
        return
    yielded = set()

    def add(ch):
        if ch is None:
            return False
        uid = unique_id(ch)
        if uid in yielded:
            return False
        yielded.add(uid)
        return True

    direct_count = 0
    for count_attr, child_attr in (("getChildrenNumber", "getChild"), ("getChildCount", "getChild"), ("childCount", "child")):
        if hasattr(o, count_attr) and hasattr(o, child_attr):
            try:
                count_obj = getattr(o, count_attr)
                count = int(count_obj()) if callable(count_obj) else int(count_obj)
                direct_count = max(direct_count, max(0, count))
                for i in range(max(0, count)):
                    child_obj = getattr(o, child_attr)
                    ch = child_obj(i) if callable(child_obj) else child_obj[i]
                    if add(ch):
                        yield ch
            except Exception:
                pass

    for fn in ("getFirstChild", "getLastChild"):
        if hasattr(o, fn):
            try:
                ch = getattr(o, fn)()
                if add(ch):
                    yield ch
            except Exception:
                pass

    try:
        rec = int(o.getChildCountRecursive()) if hasattr(o, "getChildCountRecursive") else 0
    except Exception:
        rec = 0

    if rec > direct_count and hasattr(o, "getChild"):
        try:
            for i in range(rec):
                ch = o.getChild(i)
                if add(ch):
                    yield ch
        except Exception:
            pass


def walk(o, seen, out, path):
    if o is None:
        return
    oid = unique_id(o)
    if oid in seen:
        return
    seen.add(oid)
    cur = path + [safe_name(o)]
    if is_primitive(o):
        out.append((o, cur))
    for ch in iter_children(o):
        walk(ch, seen, out, cur)


def mul4(a, b):
    return [[sum(a[i][k] * b[k][j] for k in range(4)) for j in range(4)] for i in range(4)]


def identity4():
    return [[1.0, 0.0, 0.0, 0.0], [0.0, 1.0, 0.0, 0.0], [0.0, 0.0, 1.0, 0.0], [0.0, 0.0, 0.0, 1.0]]


def matrix4x4(m):
    if m is None:
        return None
    if isinstance(m, (list, tuple)):
        if len(m) == 4 and all(hasattr(r, "__len__") and len(r) >= 4 for r in m):
            return [[float(x) for x in row[:4]] for row in m[:4]]
        if len(m) == 16:
            d = [float(x) for x in m]
            return [[d[c * 4 + r] for c in range(4)] for r in range(4)]
    if hasattr(m, "getColumnAsVec3D"):
        try:
            cols = [as_list(m.getColumnAsVec3D(i)) for i in range(4)]
            return [[cols[0][r], cols[1][r], cols[2][r], cols[3][r]] for r in range(3)] + [[0.0, 0.0, 0.0, 1.0]]
        except Exception:
            pass
    if hasattr(m, "getTranslationAsVec3D"):
        try:
            t = as_list(m.getTranslationAsVec3D())
            return [[1.0, 0.0, 0.0, t[0]], [0.0, 1.0, 0.0, t[1]], [0.0, 0.0, 1.0, t[2]], [0.0, 0.0, 0.0, 1.0]]
        except Exception:
            pass
    if hasattr(m, "data") and callable(m.data):
        try:
            d = list(m.data())
            if len(d) == 16:
                return [[float(d[c * 4 + r]) for c in range(4)] for r in range(4)]
        except Exception:
            pass
    if hasattr(m, "toString") and callable(m.toString):
        try:
            nums = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", str(m.toString()))]
            if len(nums) == 16:
                return [[nums[c * 4 + r] for c in range(4)] for r in range(4)]
        except Exception:
            pass
    return None


def local_matrix(o):
    for a in ("getGLTransformationHistory", "getGLTransformation", "getTransformationHistory", "getTransformation", "getTransform", "getMatrix"):
        obj = getattr(o, a, None)
        if obj is None:
            continue
        try:
            v = obj() if callable(obj) else obj
        except Exception:
            continue
        m = matrix4x4(v)
        if m is not None:
            return m
    return None


def world_matrix(o):
    mats = []
    cur = o
    for _ in range(32):
        m = local_matrix(cur)
        if m is not None:
            mats.append(m)
        parent = get_parent(cur)
        if parent is None:
            break
        cur = parent
    if not mats:
        return identity4()
    w = identity4()
    for m in reversed(mats):
        w = mul4(w, m)
    return w


def primitive_payload(o, path):
    t = norm_type(o)
    m = world_matrix(o)
    out = {"Id": None, "Type": type_name(o), "Name": safe_name(o), "Path": path, "Transformation history": m, "Primitive": {}, "CC Object": {}}
    if t == "cone":
        br = trr = h = apex = None
        try:
            br = float(o.getBottomRadius())
            trr = float(o.getTopRadius())
            h = float(o.getHeight())
            bottom, top = as_list(o.getBottomCenter()), as_list(o.getTopCenter())
            apex = top if trr <= br else bottom
            if abs(br - trr) < 1e-12:
                apex = [0.5 * (bottom[i] + top[i]) for i in range(3)]
        except Exception as e:
            warn(f"{safe_name(o)}: cone fields incomplete ({e})")
        out["Primitive"] = {"Height": h, "Bottom radius": br, "Top radius": trr, "Apex": apex, "Half angle (in deg.)": cone_half_angle_deg(br, trr, h)}
    elif t == "cylinder":
        r = h = apex = None
        try:
            r = float(o.getRadius()) if hasattr(o, "getRadius") else float(o.getLargeRadius())
            h = float(o.getHeight())
            b, t0 = as_list(o.getBottomCenter()), as_list(o.getTopCenter())
            apex = [0.5 * (b[i] + t0[i]) for i in range(3)]
        except Exception as e:
            warn(f"{safe_name(o)}: cylinder fields incomplete ({e})")
        out["Primitive"] = {"Height": h, "Radius": r, "Apex": apex, "Half angle (in deg.)": 0.0 if h is not None else None}
    elif t == "plane":
        normal = None
        try:
            normal = as_list(o.getNormal())
        except Exception as e:
            warn(f"{safe_name(o)}: plane normal unavailable ({e})")
        out["Primitive"] = {"Normal": normal}
        out["CC Object"] = {"Local box dimensions": max_bbox_dims(o)}
    elif t == "sphere":
        radius = None
        try:
            radius = float(o.getRadius())
        except Exception as e:
            warn(f"{safe_name(o)}: sphere radius unavailable ({e})")
        out["Primitive"] = {"Radius": radius}
    elif t == "torus":
        ri, ro = torus_radii_from_name(o)
        out["Primitive"] = {"Inner radius": ri, "Outer radius": ro}
    return out


def export_path():
    name = EXPORT_NAME if EXPORT_NAME.lower().endswith(".json") else f"{EXPORT_NAME}.json"
    cfg = writable_dir(EXPORT_DIR) if str(EXPORT_DIR).strip() else None
    if cfg is not None:
        return cfg / name
    d = writable_dir(desktop_dir()) or desktop_dir()
    return d / name


def get_selection():
    for k in ("selectedEntities", "selected_entities", "selection"):
        v = globals().get(k)
        if v:
            return list(v) if isinstance(v, (list, tuple)) else [v]
    app = None
    try:
        app = get_cc_app()
    except Exception:
        pass
    sources = [app, cc]
    for src in sources:
        if src is None:
            continue
        for fn in ("getSelectedEntities", "getSelectedEntitiesList", "selectedEntities", "selection"):
            f = getattr(src, fn, None)
            if callable(f):
                try:
                    v = f()
                    if v:
                        return list(v) if isinstance(v, (list, tuple)) else [v]
                except Exception:
                    pass
            elif f is not None:
                try:
                    return list(f) if isinstance(f, (list, tuple)) else [f]
                except Exception:
                    pass
    return []


def main():
    global cc
    cc = load_cc_module()
    if hasattr(cc, "initCC"):
        try:
            cc.initCC()
        except Exception:
            pass
    roots = get_selection()
    if not roots:
        raise SystemExit("no selection found. Select one or more entities in CloudCompare and start again.")
    items, seen = [], set()
    for r in roots:
        walk(r, seen, items, [])
    emit(f"roots={len(roots)} entities={len(seen)} primitives={len(items)}")
    by_type, out_items = {}, []
    seen_ids = set()
    for o, path in items:
        t = type_name(o)
        by_type[t] = by_type.get(t, 0) + 1
        data = primitive_payload(o, path)
        src = safe_name(o).strip()
        name = src if src and src != "<unnamed>" else f"{t}_{by_type[t]:03d}"
        if name in seen_ids:
            name = f"{name}#{by_type[t]:03d}"
        seen_ids.add(name)
        data["Id"] = name
        out_items.append(data)
        emit(f"add {name} path={'/'.join(path)}")
    out = {"schema": "cc_ransac_to_json.v3", "count": len(out_items), "items": out_items}
    p = export_path()
    try:
        with p.open("w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        emit(f"{len(out_items)} Elements saved -> {p}")
    except Exception as e:
        raise SystemExit(f"write failed: {p}: {e}")


if __name__ == "__main__":
    main()
