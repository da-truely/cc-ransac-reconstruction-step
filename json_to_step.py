#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import math
import re
import sys
from collections import defaultdict
from pathlib import Path

STEP_SCHEMA = "AP214"
PLANE_PATCH_HALF_SIZE_DEFAULT = 1.0


RESET = "[0m"
GREEN = "[32m"
RED = "[31m"
YELLOW = "[33m"
USE_COLOR = sys.stdout.isatty()
DEFAULT_OUTPUT_NAME = "json_to_step_reconstruction.step"


def emit(msg: str, prefix: str = "json_to_step", color: str | None = None) -> None:
    msg = str(msg).strip()
    if not msg:
        return
    line = f"[{prefix}] {msg}"
    if color and USE_COLOR:
        print(f"{color}{line}{RESET}", flush=True)
    else:
        print(line, flush=True)


def warn(msg: str) -> None:
    emit(f"warn: {msg}", color=YELLOW)


def desktop_dir() -> Path:
    for p in (Path.home() / "Desktop", Path.home() / "Schreibtisch"):
        if p.exists():
            return p
    return Path.home()


def writable_dir(p) -> Path | None:
    if p is None:
        return None
    try:
        p = Path(p).expanduser().resolve()
        if str(p).strip() in {"", "."}:
            return None
        p.mkdir(parents=True, exist_ok=True)
        t = p / ".__json2step_write_test__"
        t.write_text("1", encoding="utf-8")
        t.unlink(missing_ok=True)
        return p
    except Exception:
        return None


def fnum(v, default=None):
    try:
        return float(v)
    except Exception:
        return default


def load_json(path: Path):
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_items(data):
    if isinstance(data, dict):
        if isinstance(data.get("items"), list):
            return data["items"]
        if isinstance(data.get("records"), list):
            return data["records"]
    raise SystemExit("JSON layout not supported: expected top-level 'items' or 'records'.")


def load_occ():
    backends = []

    try:
        from OCC.Core.BRepBuilderAPI import (
            BRepBuilderAPI_MakeEdge,
            BRepBuilderAPI_MakeFace,
            BRepBuilderAPI_MakeWire,
            BRepBuilderAPI_Transform,
        )
        from OCC.Core.BRepPrimAPI import (
            BRepPrimAPI_MakeCone,
            BRepPrimAPI_MakeCylinder,
            BRepPrimAPI_MakeSphere,
            BRepPrimAPI_MakeTorus,
        )
        from OCC.Core.IFSelect import IFSelect_RetDone
        from OCC.Core.Interface import Interface_Static
        from OCC.Core.STEPCAFControl import STEPCAFControl_Writer
        from OCC.Core.STEPControl import STEPControl_AsIs
        from OCC.Core.TCollection import TCollection_ExtendedString
        from OCC.Core.TDataStd import TDataStd_Name
        from OCC.Core.TDocStd import TDocStd_Document
        from OCC.Core.XCAFDoc import XCAFDoc_DocumentTool
        from OCC.Core.gp import gp_Dir, gp_Pln, gp_Pnt, gp_Trsf
        backends.append("OCC")
    except Exception as e1:
        try:
            import cadquery as cq  # noqa: F401
            from OCP.BRepBuilderAPI import (
                BRepBuilderAPI_MakeEdge,
                BRepBuilderAPI_MakeFace,
                BRepBuilderAPI_MakeWire,
                BRepBuilderAPI_Transform,
            )
            from OCP.BRepPrimAPI import (
                BRepPrimAPI_MakeCone,
                BRepPrimAPI_MakeCylinder,
                BRepPrimAPI_MakeSphere,
                BRepPrimAPI_MakeTorus,
            )
            from OCP.IFSelect import IFSelect_RetDone
            from OCP.Interface import Interface_Static
            from OCP.STEPCAFControl import STEPCAFControl_Writer
            from OCP.STEPControl import STEPControl_AsIs
            from OCP.TCollection import TCollection_ExtendedString
            from OCP.TDataStd import TDataStd_Name
            from OCP.TDocStd import TDocStd_Document
            from OCP.XCAFDoc import XCAFDoc_DocumentTool
            from OCP.gp import gp_Dir, gp_Pln, gp_Pnt, gp_Trsf
            backends.append("cadquery-ocp")
        except Exception as e2:
            raise SystemExit(
                f"STEP backend fehlt.\n"
                f"OCP error: {e1}\n"
                f"cadquery-ocp error: {e2}\n"
            )

    return {
        "BRepBuilderAPI_MakeEdge": BRepBuilderAPI_MakeEdge,
        "BRepBuilderAPI_MakeFace": BRepBuilderAPI_MakeFace,
        "BRepBuilderAPI_MakeWire": BRepBuilderAPI_MakeWire,
        "BRepBuilderAPI_Transform": BRepBuilderAPI_Transform,
        "BRepPrimAPI_MakeCone": BRepPrimAPI_MakeCone,
        "BRepPrimAPI_MakeCylinder": BRepPrimAPI_MakeCylinder,
        "BRepPrimAPI_MakeSphere": BRepPrimAPI_MakeSphere,
        "BRepPrimAPI_MakeTorus": BRepPrimAPI_MakeTorus,
        "IFSelect_RetDone": IFSelect_RetDone,
        "Interface_Static": Interface_Static,
        "STEPCAFControl_Writer": STEPCAFControl_Writer,
        "STEPControl_AsIs": STEPControl_AsIs,
        "TCollection_ExtendedString": TCollection_ExtendedString,
        "TDataStd_Name": TDataStd_Name,
        "TDocStd_Document": TDocStd_Document,
        "XCAFDoc_DocumentTool": XCAFDoc_DocumentTool,
        "gp_Dir": gp_Dir,
        "gp_Pln": gp_Pln,
        "gp_Pnt": gp_Pnt,
        "gp_Trsf": gp_Trsf,
        "backend": backends[0] if backends else "none",
    }


occ = load_occ()


def set_static_str(name, value):
    for fn in ("SetCVal_s", "SetCVal"):
        f = getattr(occ["Interface_Static"], fn, None)
        if callable(f):
            return f(name, value)
    raise RuntimeError(f"Interface_Static has no string setter for {name}")


def set_static_int(name, value):
    for fn in ("SetIVal_s", "SetIVal"):
        f = getattr(occ["Interface_Static"], fn, None)
        if callable(f):
            return f(name, int(value))
    raise RuntimeError(f"Interface_Static has no int setter for {name}")


def first_item_key(item):
    for k in ("Type", "type"):
        v = item.get(k)
        if isinstance(v, str) and v.strip():
            return v.strip().lower()
    return ""


def get_matrix(item):
    m = item.get("Transformation history", item.get("Matrix"))
    if isinstance(m, dict):
        m = m.get("matrix", m.get("Matrix"))
    if not m:
        return None
    try:
        if len(m) == 4 and all(len(r) == 4 for r in m):
            return [[float(x) for x in r] for r in m]
    except Exception:
        pass
    return None


def trsf_from_matrix(m):
    if not m:
        return None
    t = occ["gp_Trsf"]()
    t.SetValues(
        m[0][0], m[0][1], m[0][2], m[0][3],
        m[1][0], m[1][1], m[1][2], m[1][3],
        m[2][0], m[2][1], m[2][2], m[2][3],
    )
    return t


def apply_trsf(shape, m):
    if shape is None or not m:
        return shape
    tr = trsf_from_matrix(m)
    if tr is None:
        return shape
    return occ["BRepBuilderAPI_Transform"](shape, tr, True).Shape()


def translate_z(shape, dz):
    if shape is None or not dz:
        return shape
    tr = occ["gp_Trsf"]()
    tr.SetTranslation(occ["gp_Pnt"](0, 0, 0), occ["gp_Pnt"](0, 0, dz))
    return occ["BRepBuilderAPI_Transform"](shape, tr, True).Shape()


def parse_torus_radii(name):
    m = re.search(r"r\s*=\s*([0-9.+\-eE]+)\s*/\s*R\s*=\s*([0-9.+\-eE]+)", str(name))
    if not m:
        return None, None
    return float(m.group(1)), float(m.group(2))


def rect_face(x, y):
    x = abs(float(x)) if x is not None else 1.0
    y = abs(float(y)) if y is not None else 1.0
    p = [occ["gp_Pnt"](-x / 2, -y / 2, 0), occ["gp_Pnt"](x / 2, -y / 2, 0), occ["gp_Pnt"](x / 2, y / 2, 0), occ["gp_Pnt"](-x / 2, y / 2, 0)]
    e = [occ["BRepBuilderAPI_MakeEdge"](p[i], p[(i + 1) % 4]).Edge() for i in range(4)]
    w = occ["BRepBuilderAPI_MakeWire"]()
    for edge in e:
        w.Add(edge)
    return occ["BRepBuilderAPI_MakeFace"](w.Wire()).Face()


def dims_xy(dims):
    if isinstance(dims, (list, tuple)) and len(dims) >= 2:
        return dims[0], dims[1]
    if isinstance(dims, dict):
        vals = [dims.get(k) for k in ("X", "Y", "Z") if dims.get(k) is not None]
        if len(vals) >= 2:
            vals = sorted(vals, key=lambda v: abs(float(v)), reverse=True)
            return vals[0], vals[1]
    return None, None


def shape_from_item(item):
    t = first_item_key(item)
    prim = item.get("Primitive", {}) or item.get("params", {}) or {}
    obj = item.get("CC Object", {}) or {}
    mat = get_matrix(item)

    if t == "cone":
        h = fnum(prim.get("Height"))
        br = fnum(prim.get("Bottom radius"))
        tr = fnum(prim.get("Top radius"))
        if None in (h, br, tr):
            return None
        shp = occ["BRepPrimAPI_MakeCone"](br, tr, h).Shape()
        return apply_trsf(translate_z(shp, -h / 2), mat)

    if t == "cylinder":
        h = fnum(prim.get("Height"))
        r = fnum(prim.get("Radius"))
        if None in (h, r):
            return None
        shp = occ["BRepPrimAPI_MakeCylinder"](r, h).Shape()
        return apply_trsf(translate_z(shp, -h / 2), mat)

    if t == "plane":
        dims = obj.get("Local box dimensions", obj.get("local box dimensions", {})) or {}
        x, y = dims_xy(dims)
        shp = rect_face(x, y)
        return apply_trsf(shp, mat)

    if t == "sphere":
        r = fnum(prim.get("Radius"))
        if r is None:
            return None
        return apply_trsf(occ["BRepPrimAPI_MakeSphere"](r).Shape(), mat)

    if t == "torus":
        r = fnum(prim.get("Inner radius"))
        R = fnum(prim.get("Outer radius"))
        if r is None or R is None:
            r2, R2 = parse_torus_radii(item.get("Name", item.get("name", "")))
            r = r if r is not None else r2
            R = R if R is not None else R2
        if None in (r, R):
            return None
        return apply_trsf(occ["BRepPrimAPI_MakeTorus"](R, r).Shape(), mat)

    return None


def out_files(argv):
    if len(argv) > 1:
        p = Path(str(argv[1]).strip().strip('"').strip("'")).expanduser()
        if p.is_dir():
            return sorted(p.glob("*.json")), p
        if p.is_file():
            return [p], p.parent
    d = desktop_dir()
    return sorted(d.glob("*.json")), d


def resolve_output(argv, input_json):
    if len(argv) <= 2:
        return Path.cwd() / DEFAULT_OUTPUT_NAME
    raw = str(argv[2]).strip().strip('"').strip("'")
    if not raw:
        return Path.cwd() / DEFAULT_OUTPUT_NAME
    has_sep = ("\\" in raw) or ("/" in raw)
    p = Path(raw).expanduser()
    if p.suffix.lower() == ".step":
        if p.is_absolute():
            return p
        if has_sep:
            return Path.cwd() / Path(raw.lstrip('\\/'))
        return Path.cwd() / p.name
    if p.is_absolute():
        return p / DEFAULT_OUTPUT_NAME
    if has_sep:
        return Path.cwd() / Path(raw.lstrip('\\/')) / DEFAULT_OUTPUT_NAME
    return Path.cwd() / f"{raw}.step"


def shape_tool_for(doc):
    for fn in ("ShapeTool_s", "ShapeTool"):
        f = getattr(occ["XCAFDoc_DocumentTool"], fn, None)
        if callable(f):
            try:
                return f(doc.Main())
            except Exception:
                pass
    raise RuntimeError("Could not create XCAF shape tool")


def write_step(items, out_path):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = occ["TDocStd_Document"](occ["TCollection_ExtendedString"]("json2step"))
    shape_tool = shape_tool_for(doc)
    count = defaultdict(int)
    total = 0

    for src, data in items:
        t = first_item_key(data)
        if t not in {"cone", "cylinder", "plane", "sphere", "torus"}:
            warn(f"skip {src.name}: unsupported type={data.get('Type', data.get('type'))}")
            continue
        shp = shape_from_item(data)
        if shp is None:
            warn(f"skip {src.name}: shape build failed")
            continue
        count[t] += 1
        name = f"{data.get('Type', data.get('type', t)).strip().title()}_{count[t]:03d}"
        lab = shape_tool.AddShape(shp, False)
        try:
            occ["TDataStd_Name"].Set_s(lab, occ["TCollection_ExtendedString"](name))
        except Exception:
            pass
        total += 1
        emit(f"added {name} from {src.name}")

    if total == 0:
        raise SystemExit("no exportable shapes found")

    set_static_str("write.step.schema", STEP_SCHEMA)
    set_static_int("write.stepcaf.subshapes.name", 1)
    set_static_str("write.step.product.name", out_path.stem)

    writer = occ["STEPCAFControl_Writer"]()
    try:
        status = writer.Transfer(doc, occ["STEPControl_AsIs"])
    except TypeError:
        status = writer.Transfer(doc)
    if status is False:
        warn("writer.Transfer returned False")
    ret = writer.Write(str(out_path))
    if ret != occ["IFSelect_RetDone"]:
        raise RuntimeError(f"STEP write failed: {ret}")
    emit(f"{total} shapes saved -> {out_path}", color=GREEN)


def main(argv):
    files, base = out_files(argv)
    if not files:
        raise SystemExit(f"no JSON files found in {base}")
    emit(f"input_base={base}")
    emit(f"json_files={len(files)}")
    items = []
    for p in files:
        try:
            data = load_json(p)
            vals = load_items(data)
            for v in vals:
                items.append((p, v))
        except Exception as e:
            warn(f"cannot read {p.name}: {e}")
    input_json = files[0] if len(files) == 1 else None
    out_path = resolve_output(argv, input_json)
    write_step(items, out_path)


if __name__ == "__main__":
    main(sys.argv)
