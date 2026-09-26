# -*- coding: utf-8 -*-
"""Atrapa API HyperMesha do testów makra poza HyperMeshem.

Rejestruje w sys.modules moduły ``hm`` (Model, Collection, entities, ...),
``hm.mdi.apis`` (HmModelDebug) i ``hw`` (evalTcl, CaptureImageTool). Model jest
prostym słownikiem w pamięci wczytywanym z pliku JSON o rozszerzeniu .hm
(patrz make_models.py). Obsługiwany jest podzbiór API, którego używa
HM_Quality_Studio.py, w tym mini-dyspozytor poleceń Tcl (odczyt hurtowy
::hmqs::*, color_rgb, przezroczystość, zaznaczenie węzłów "by geoms").

Przełączniki symulujące różne wersje HM:
    CONFIG["transp_mode"]  = "setvalue" | "mark" | "pyapi" | "none"
    CONFIG["transp_read"]  = True / False   (czy hm_getvalue transparency działa)
    CONFIG["rgb"]          = True / False   (czy działa color_rgb)
    CONFIG["by_geoms"]     = True / False   (czy działa *createmark nodes "by geoms")
    CONFIG["interactive"]  = lista ID zwracana przez CollectionByInteractiveSelection
"""
import io
import json
import math
import os
import re
import sys
import types

CONFIG = {"transp_mode": "setvalue", "transp_read": True, "rgb": True, "by_geoms": True,
          "interactive": [1], "mark_surfs": [1], "bulk": True}

PALETTE = [
    (0, 0, 0), (255, 255, 255), (255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0), (0, 255, 255), (255, 0, 255),
    (204, 204, 204), (185, 185, 185), (163, 163, 163), (140, 140, 140), (118, 118, 118), (96, 96, 96), (73, 73, 73), (51, 51, 51),
    (254, 56, 23), (235, 49, 23), (200, 40, 23), (130, 5, 23), (49, 111, 255), (44, 102, 236), (35, 85, 199), (21, 49, 126),
    (255, 131, 88), (238, 117, 81), (201, 99, 65), (130, 56, 23), (95, 181, 255), (89, 167, 236), (74, 140, 199), (44, 85, 126),
    (254, 53, 138), (235, 49, 127), (199, 40, 105), (129, 5, 65), (163, 124, 255), (150, 115, 236), (126, 94, 199), (80, 56, 126),
    (255, 189, 185), (240, 175, 170), (203, 146, 142), (131, 31, 88), (252, 62, 255), (233, 56, 236), (198, 49, 199), (129, 27, 126),
    (255, 179, 23), (240, 165, 23), (203, 139, 23), (131, 83, 23), (97, 254, 110), (90, 236, 100), (78, 200, 82), (53, 125, 44),
    (255, 235, 124), (244, 217, 114), (207, 183, 96), (133, 116, 57), (201, 254, 23), (187, 236, 23), (158, 200, 23), (100, 125, 23)]


class Status(object):
    def __init__(self, status=0, message=""):
        self.status, self.message = status, message


class _State(object):
    """Model w pamięci (jeden na sesję)."""

    def __init__(self):
        self.clear()

    def clear(self):
        self.file = ""
        self.comps = {}      # id -> {"name", "color", "rgb", "shown", "transp"}
        self.elems = {}      # id -> {"cfg", "comp", "nodes", metryki...}
        self.nodes = {}      # id -> [x, y, z]
        self.sets = {}       # id -> {"name", "ids"}
        self.tags = {}       # id -> {"label", "body", "elem"}
        self.surfs = {}      # id -> {"nodes": [ids]}
        self.next_comp = 1
        self.next_set = 1
        self.next_tag = 1
        self.marks = {"elems": [], "nodes": [], "comps": [], "surfs": list(CONFIG["mark_surfs"])}
        self.hmqs_ids = []
        self.transp_value = 0
        self.view = [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, -1, -1, 1, 1]
        self.bg = "#ffffff"
        self.log = []

    def load(self, path):
        with io.open(path, "r", encoding="utf-8") as fh:
            d = json.load(fh)
        self.clear()
        self.file = os.path.normpath(path)
        for c in d["comps"]:
            self.comps[int(c["id"])] = {"name": c["name"], "color": int(c.get("color", 5)), "rgb": None, "shown": True, "transp": 0}
        self.next_comp = max(self.comps) + 1 if self.comps else 1
        for k, v in d["nodes"].items():
            self.nodes[int(k)] = [float(x) for x in v]
        for k, v in d["elems"].items():
            e = dict(v)
            e["nodes"] = [int(n) for n in e["nodes"]]
            self.elems[int(k)] = e
        for k, v in d.get("surfs", {}).items():
            self.surfs[int(k)] = {"nodes": [int(n) for n in v["nodes"]]}
        self.marks["surfs"] = list(CONFIG["mark_surfs"])
        return True

    def comp_by_name(self, name):
        for cid, c in self.comps.items():
            if c["name"] == name:
                return cid
        return 0

    def new_comp(self, name=""):
        cid = self.next_comp
        self.next_comp += 1
        self.comps[cid] = {"name": name or "component%d" % cid, "color": 5, "rgb": None, "shown": True, "transp": 0}
        return cid

    def comp_rgb(self, cid):
        c = self.comps[cid]
        if c["rgb"]:
            return tuple(c["rgb"])
        return PALETTE[max(1, min(64, c["color"])) - 1]

    def elems_of_comps(self, cids):
        cids = set(cids)
        return [eid for eid, e in self.elems.items() if e["comp"] in cids]

    def displayed_elems(self):
        shown = set(cid for cid, c in self.comps.items() if c["shown"])
        return [eid for eid, e in self.elems.items() if e["comp"] in shown]

    def centroid(self, eid):
        pts = [self.nodes[n] for n in self.elems[eid]["nodes"] if n in self.nodes]
        if not pts:
            return None
        k = float(len(pts))
        return tuple(sum(p[i] for p in pts) / k for i in range(3))


STATE = _State()


# ----------------------------------------------------------------- encje
class _Entity(object):
    kind = ""

    def __init__(self, model=None, id=None):
        self._model = model
        self.id = int(id) if id is not None else None

    def __repr__(self):
        return "<%s %s>" % (self.__class__.__name__, self.id)


class Node(_Entity):
    kind = "nodes"

    def __getattr__(self, name):
        if name in ("x", "y", "z"):
            p = STATE.nodes.get(self.id)
            if p is None:
                raise AttributeError(name)
            return p["xyz".index(name)]
        raise AttributeError(name)


class Component(_Entity):
    kind = "comps"

    def __init__(self, model=None, id=None):
        if id is None:
            id = STATE.new_comp()
        _Entity.__init__(self, model, id)

    def _rec(self):
        return STATE.comps[self.id]

    def __getattr__(self, name):
        if name.startswith("_") or name in ("id",):
            raise AttributeError(name)
        c = STATE.comps.get(self.id)
        if c is None:
            raise AttributeError(name)
        if name == "name":
            return c["name"]
        if name == "color":
            return c["color"]
        if name == "color_rgb":
            if not CONFIG["rgb"]:
                raise AttributeError(name)
            return list(STATE.comp_rgb(self.id))
        if name == "elements":
            return [Element(self._model, e) for e in STATE.elems_of_comps([self.id])]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name in ("name", "color", "color_rgb"):
            c = STATE.comps[self.id]
            if name == "name":
                c["name"] = "%s" % value
            elif name == "color":
                c["color"] = int(value)
                c["rgb"] = None
            else:
                if not CONFIG["rgb"]:
                    raise AttributeError(name)
                c["rgb"] = [int(v) for v in value][:3]
        else:
            object.__setattr__(self, name, value)


class Element(_Entity):
    kind = "elems"

    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        e = STATE.elems.get(self.id)
        if e is None:
            raise AttributeError(name)
        if name == "config":
            return e["cfg"]
        if name == "nodes":
            return [Node(self._model, n) for n in e["nodes"]]
        if name == "collector":
            return Component(self._model, e["comp"])
        if name in ("centerx", "centery", "centerz"):
            c = STATE.centroid(self.id)
            return None if c is None else c["xyz".index(name[-1])]
        if name in e:
            return e[name]
        return None


class Set(_Entity):
    kind = "sets"

    def __getattr__(self, name):
        s = STATE.sets.get(self.id)
        if s is None or name.startswith("_"):
            raise AttributeError(name)
        if name == "name":
            return s["name"]
        if name == "elements":
            return [Element(self._model, e) for e in s["ids"]]
        raise AttributeError(name)


class Tag(_Entity):
    kind = "tags"

    def __getattr__(self, name):
        t = STATE.tags.get(self.id)
        if t is None or name.startswith("_"):
            raise AttributeError(name)
        return t.get(name)


class Surface(_Entity):
    kind = "surfs"


class Line(_Entity):
    kind = "lines"


class Material(_Entity):
    kind = "mats"


class Property(_Entity):
    kind = "props"


ENT_KINDS = {"Node": Node, "Component": Component, "Element": Element, "Set": Set, "Tag": Tag, "Surface": Surface,
             "Line": Line, "Material": Material, "Property": Property}


def _all_ids(cls):
    return {Node: STATE.nodes, Component: STATE.comps, Element: STATE.elems, Set: STATE.sets, Tag: STATE.tags,
            Surface: STATE.surfs}.get(cls, {})


class FilterByCollection(object):
    def __init__(self, cls, by_cls):
        self.cls, self.by_cls = cls, by_cls


class Collection(object):
    def __init__(self, model, cls, ids=None):
        self.model = model
        if isinstance(cls, FilterByCollection):
            src = ids
            self.cls = cls.cls
            if cls.cls is Element and cls.by_cls is Component:
                self.ids = STATE.elems_of_comps([c.id for c in src])
            else:
                self.ids = []
            return
        self.cls = cls
        if ids is None:
            self.ids = sorted(_all_ids(cls).keys())
        else:
            self.ids = [int(i) for i in ids]
            if not self.ids:
                raise RuntimeError("empty collection")
            self.ids = [i for i in self.ids if i in _all_ids(cls)]

    def __iter__(self):
        return iter([self.cls(self.model, i) for i in self.ids])

    def __len__(self):
        return len(self.ids)


class CollectionByDisplayed(Collection):
    def __init__(self, model, cls):
        Collection.__init__(self, model, cls, None)
        if cls is Element:
            self.ids = sorted(STATE.displayed_elems())


class CollectionByInteractiveSelection(Collection):
    def __init__(self, model, cls):
        Collection.__init__(self, model, cls, None)
        self.ids = list(CONFIG["interactive"])


# ----------------------------------------------------------------- Model
class Model(object):
    def readfile(self, path, flag=0):
        p = path.replace("/", os.sep)
        if not os.path.isfile(p):
            return Status(1, "no such file")
        STATE.load(p)
        return Status(0, "")

    def mergefile2(self, path):
        p = path.replace("/", os.sep)
        with io.open(p, "r", encoding="utf-8") as fh:
            d = json.load(fh)
        noff = max(STATE.nodes) if STATE.nodes else 0
        eoff = max(STATE.elems) if STATE.elems else 0
        cmap = {}
        for c in d["comps"]:
            cid = STATE.new_comp(c["name"] + "_merged")
            cmap[int(c["id"])] = cid
        for k, v in d["nodes"].items():
            STATE.nodes[int(k) + noff] = [float(x) for x in v]
        for k, v in d["elems"].items():
            e = dict(v)
            e["nodes"] = [int(n) + noff for n in e["nodes"]]
            e["comp"] = cmap[int(e["comp"])]
            STATE.elems[int(k) + eoff] = e
        return Status(0, "")

    def mergefile(self, path, a, b):
        return self.mergefile2(path)

    def hm_answernext(self, s):
        return Status(0, "")

    def hm_redraw(self):
        return Status(0, "")

    def movemark(self, col, comp_name):
        cid = STATE.comp_by_name(comp_name)
        if not cid:
            return Status(1, "no component %s" % comp_name)
        for i in col.ids:
            if col.cls is Element and i in STATE.elems:
                STATE.elems[i]["comp"] = cid
        return Status(0, "")

    def deletemark(self, col):
        for i in col.ids:
            if col.cls is Component:
                for eid in STATE.elems_of_comps([i]):
                    del STATE.elems[eid]
                STATE.comps.pop(i, None)
            elif col.cls is Element:
                STATE.elems.pop(i, None)
            elif col.cls is Node:
                STATE.nodes.pop(i, None)
            elif col.cls is Set:
                STATE.sets.pop(i, None)
            elif col.cls is Tag:
                STATE.tags.pop(i, None)
            elif col.cls is Surface:
                STATE.surfs.pop(i, None)
        return Status(0, "")

    def displaycollectorsbymark(self, col, onoff, a, b):
        for i in col.ids:
            if i in STATE.comps:
                STATE.comps[i]["shown"] = onoff == "on"
        return Status(0, "")

    def transparencyvalue(self, v):
        STATE.transp_value = int(v)
        return Status(0, "")

    def transparencymark(self, col):
        if CONFIG["transp_mode"] not in ("pyapi", "mark"):
            return Status(1, "unknown command")
        for i in col.ids:
            if i in STATE.comps:
                STATE.comps[i]["transp"] = STATE.transp_value
        return Status(0, "")

    def viewset(self, *nums):
        STATE.view = [float(x) for x in nums[:20]]
        return Status(0, "")

    def hm_getcurrentview(self):
        vm = types.SimpleNamespace(viewMatrix=[STATE.view[i * 4:(i + 1) * 4] for i in range(4)])
        return (Status(0, ""), vm)

    def jpegfilenamed(self, path):
        _render(path, "jpg")
        return Status(0, "")


class HmModelDebug(object):
    def entitysetcreate(self, name, col):
        sid = STATE.next_set
        STATE.next_set += 1
        STATE.sets[sid] = {"name": name, "ids": list(col.ids)}
        return Status(0, "")

    def tagcreate(self, elem, label, body, color):
        tid = STATE.next_tag
        STATE.next_tag += 1
        STATE.tags[tid] = {"label": label, "body": body, "elem": elem.id, "color": color}
        return Status(0, "")

    def setbackgroundcolor(self, r, g, b):
        STATE.bg = "#%02x%02x%02x" % (r, g, b)
        return Status(0, "")


# ----------------------------------------------------------------- zrzut
def _render(path, fmt="png"):
    """Rzut z góry (x, y) elementów widocznych komponentów, kolory komponentów,
    przezroczystość jako mieszanie z białym."""
    from PIL import Image, ImageDraw
    W, H = 640, 420
    im = Image.new("RGB", (W, H), (255, 255, 255))
    dr = ImageDraw.Draw(im)
    if STATE.nodes:
        xs = [p[0] for p in STATE.nodes.values()]
        ys = [p[1] for p in STATE.nodes.values()]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        sx = (W - 60) / max(1e-9, x1 - x0)
        sy = (H - 60) / max(1e-9, y1 - y0)
        s = min(sx, sy)
        for eid, e in sorted(STATE.elems.items()):
            c = STATE.comps.get(e["comp"])
            if c is None or not c["shown"]:
                continue
            pts = [STATE.nodes[n] for n in e["nodes"][:4] if n in STATE.nodes]
            if len(pts) < 3:
                continue
            rgb = STATE.comp_rgb(e["comp"])
            a = 1.0 - c["transp"] / 100.0
            col = tuple(int(255 + (v - 255) * a) for v in rgb)
            poly = [(30 + (p[0] - x0) * s + p[2] * 2.0, H - 30 - (p[1] - y0) * s - p[2] * 2.0) for p in pts]
            dr.polygon(poly, fill=col, outline=(90, 90, 90))
    dr.text((8, 4), os.path.basename(STATE.file), fill=(0, 0, 0))
    if fmt == "jpg":
        im.save(path, "JPEG", quality=90)
    else:
        im.save(path, "PNG")


class CaptureImageTool(object):
    def __init__(self):
        self.type = "png"
        self.file = ""

    def capture(self):
        _render(self.file, "jpg" if self.type == "jpg" else "png")


# ----------------------------------------------------------------- Tcl
_DATANAMES = {"config": "cfg", "aspect": "aspect", "jacobian": "jacobian", "skew": "skew", "warpage": "warpage",
              "taper": "taper", "minangle": "minangle", "maxangle": "maxangle", "shortestside": "shortestside",
              "length": "length", "volume": "volume", "area": "area", "collector.id": "comp"}
_BULK = {"loaded": False}


def _fmt(v):
    if v is None:
        return "NA"
    if isinstance(v, float):
        return "%.10g" % v
    return "%s" % v


def _elem_value(eid, dn):
    e = STATE.elems.get(eid)
    if e is None:
        raise RuntimeError("no element")
    if dn == "id":
        return eid
    key = _DATANAMES.get(dn)
    if key is None:
        raise RuntimeError("unknown dataname %s" % dn)
    v = e.get(key)
    if v is None:
        raise RuntimeError("no value")
    return v


def evalTcl(cmd):
    c = cmd.strip()
    STATE.log.append(c[:80])
    if c.startswith("namespace eval ::hmqs"):
        if not CONFIG["bulk"]:
            raise RuntimeError("bulk disabled")
        _BULK["loaded"] = True
        return ""
    if c == "::hmqs::ping":
        if not _BULK["loaded"]:
            raise RuntimeError("no proc")
        return "hmqs-ok"
    if c == "expr 1":
        return "1"
    if c == "hm_winfo viewmatrix":
        return " ".join("%g" % x for x in STATE.view)
    if c == "hm_info currentfile":
        return STATE.file
    if c == "hm_winfo entitycolors":
        return " ".join("%d %d %d" % p for p in PALETTE)
    if c.startswith("hm_blockbrowserupdate") or c.startswith("hm_blockredraw"):
        return ""
    if c.startswith("hwf::setbackgroundcolor"):
        STATE.bg = c.split()[-1]
        return ""
    m = re.match(r"^set ::hmqs::ids \{(.*)\}$", c, re.S)
    if m:
        STATE.hmqs_ids = [int(x) for x in m.group(1).split()]
        return ""
    m = re.match(r"^::hmqs::mark (\w+) (\w+)$", c)
    if m:
        t, sel = m.group(1), m.group(2)
        pool = {"elems": STATE.elems, "nodes": STATE.nodes, "comps": STATE.comps}[t]
        if sel == "add":
            STATE.marks[t] += [i for i in STATE.hmqs_ids if i in pool]
            return "%d" % len(STATE.marks[t])
        if sel == "all":
            STATE.marks[t] = sorted(pool.keys())
        elif sel == "displayed":
            STATE.marks[t] = sorted(STATE.displayed_elems()) if t == "elems" else sorted(pool.keys())
        elif sel == "ids":
            STATE.marks[t] = [i for i in STATE.hmqs_ids if i in pool]
        else:
            STATE.marks[t] = []
        return " ".join("%d" % i for i in STATE.marks[t])
    m = re.match(r"^::hmqs::getmark (\w+)$", c)
    if m:
        return " ".join("%d" % i for i in STATE.marks[m.group(1)])
    m = re.match(r"^::hmqs::vals (\w+) ([\w.]+)$", c)
    if m:
        t, dn = m.group(1), m.group(2)
        out = []
        for i in STATE.marks[t]:
            if t == "nodes":
                if dn == "id":
                    out.append("%d" % i)
                elif dn in ("x", "y", "z"):
                    out.append(_fmt(STATE.nodes[i]["xyz".index(dn)]))
                else:
                    out.append("NA")
            else:
                try:
                    out.append(_fmt(_elem_value(i, dn)))
                except RuntimeError:
                    out.append("NA")
        return " ".join(out)
    if c == "::hmqs::nodes":
        return " ".join("{%s}" % " ".join("%d" % n for n in STATE.elems[i]["nodes"]) for i in STATE.marks["elems"])
    if c.startswith("foreach t {elems nodes comps}"):
        for k in ("elems", "nodes", "comps"):
            STATE.marks[k] = []
        return ""
    m = re.match(r"^\*clearmark (\w+) \d$", c)
    if m:
        STATE.marks[m.group(1)] = []
        return ""
    m = re.match(r'^\*createmark nodes 1 "by geoms" surfs \{?([\d ]+)\}?$', c)
    if m:
        if not CONFIG["by_geoms"]:
            raise RuntimeError("bad selection")
        ids = [int(x) for x in m.group(1).split()]
        out = []
        for s in ids:
            out += STATE.surfs.get(s, {}).get("nodes", [])
        STATE.marks["nodes"] = sorted(set(out))
        return ""
    m = re.match(r"^\*createmark comps 1 ([\d ]+)$", c)
    if m:
        STATE.marks["comps"] = [int(x) for x in m.group(1).split()]
        return ""
    if c.startswith("*createmark"):
        raise RuntimeError("unsupported selection: %s" % c)
    m = re.match(r"^hm_getmark (\w+) 1$", c)
    if m:
        return " ".join("%d" % i for i in STATE.marks[m.group(1)])
    m = re.match(r"^\*transparencyvalue (\d+)$", c)
    if m:
        if CONFIG["transp_mode"] != "mark":
            raise RuntimeError("unknown command")
        STATE.transp_value = int(m.group(1))
        return ""
    if c == "*transparencymark comps 1":
        if CONFIG["transp_mode"] != "mark":
            raise RuntimeError("unknown command")
        for i in STATE.marks["comps"]:
            STATE.comps[i]["transp"] = STATE.transp_value
        return ""
    m = re.match(r"^\*setvalue comps id=(\d+) color_rgb=\{(\d+) (\d+) (\d+)\}$", c)
    if m:
        if not CONFIG["rgb"]:
            raise RuntimeError("unknown dataname")
        STATE.comps[int(m.group(1))]["rgb"] = [int(m.group(2)), int(m.group(3)), int(m.group(4))]
        return ""
    m = re.match(r"^\*setvalue comps id=(\d+) transparency=(\d+)$", c)
    if m:
        if CONFIG["transp_mode"] != "setvalue":
            raise RuntimeError("unknown dataname")
        STATE.comps[int(m.group(1))]["transp"] = int(m.group(2))
        return ""
    m = re.match(r"^hm_getvalue comps id=(\d+) dataname=(\w+)$", c)
    if m:
        cid, dn = int(m.group(1)), m.group(2)
        if dn == "color_rgb":
            if not CONFIG["rgb"]:
                raise RuntimeError("unknown dataname")
            return "{%d %d %d}" % STATE.comp_rgb(cid)
        if dn == "transparency":
            if not CONFIG["transp_read"]:
                raise RuntimeError("unknown dataname")
            return "%d" % STATE.comps[cid]["transp"]
        raise RuntimeError("unknown dataname")
    m = re.match(r"^hm_getvalue sets id=(\d+) dataname=ids$", c)
    if m:
        return " ".join("%d" % i for i in STATE.sets[int(m.group(1))]["ids"])
    m = re.match(r"^if \{\[catch \{hm_getvalue elems id=(\d+) dataname=([\w.]+)\} v\]\} \{return NA\} else \{return \$v\}$", c)
    if m:
        try:
            return _fmt(_elem_value(int(m.group(1)), m.group(2)))
        except RuntimeError:
            return "NA"
    m = re.match(r"^hm_getvalue elems id=(\d+) dataname=([\w.]+)$", c)
    if m:
        return _fmt(_elem_value(int(m.group(1)), m.group(2)))
    raise RuntimeError("unknown Tcl: %s" % c[:60])


# ----------------------------------------------------------------- rejestracja
def install():
    hm = types.ModuleType("hm")
    hm.Model = Model
    hm.Collection = Collection
    hm.CollectionByDisplayed = CollectionByDisplayed
    hm.CollectionByInteractiveSelection = CollectionByInteractiveSelection
    hm.FilterByCollection = FilterByCollection
    ent = types.ModuleType("hm.entities")
    for nm, cls in ENT_KINDS.items():
        setattr(ent, nm, cls)
    hm.entities = ent
    mdi = types.ModuleType("hm.mdi")
    apis = types.ModuleType("hm.mdi.apis")
    apis.HmModelDebug = HmModelDebug
    mdi.apis = apis
    hm.mdi = mdi
    hw = types.ModuleType("hw")
    hw.evalTcl = evalTcl
    hw.CaptureImageTool = CaptureImageTool
    for name, mod in (("hm", hm), ("hm.entities", ent), ("hm.mdi", mdi), ("hm.mdi.apis", apis), ("hw", hw)):
        sys.modules[name] = mod
    return hm, hw
