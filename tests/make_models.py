# -*- coding: utf-8 -*-
"""Syntetyczne modele REF / INF (pliki JSON z rozszerzeniem .hm) dla atrapy
API HyperMesha: blok hexa8 (3D) nx x ny x nz + "skórka" quad4 (2D) na górnej
ścianie w osobnym komponencie + powierzchnia nr 1 = górna ściana.
INF = REF z przesuniętymi węzłami i pogorszonymi metrykami części elementów."""
import io
import json
import math
import os
import random


def make_model(path, nx=8, ny=6, nz=3, seed=1, morph=0.0, worsen=0.0, drop_last=0, extra=0):
    rnd = random.Random(seed)
    nodes, elems, comps = {}, {}, [{"id": 1, "name": "solid_block", "color": 21}, {"id": 2, "name": "skin_top", "color": 53},
                                  {"id": 3, "name": "beams", "color": 6}]
    nid = {}
    k = 1
    for iz in range(nz + 1):
        for iy in range(ny + 1):
            for ix in range(nx + 1):
                x, y, z = float(ix), float(iy), float(iz)
                if morph:
                    x += morph * math.sin(iy * 0.7) * iz / float(nz)
                    y += morph * math.cos(ix * 0.5) * iz / float(nz)
                    z += 0.3 * morph * math.sin(ix + iy)
                nodes[k] = [x, y, z]
                nid[(ix, iy, iz)] = k
                k += 1
    eid = 1
    for iz in range(nz):
        for iy in range(ny):
            for ix in range(nx):
                n = [nid[(ix, iy, iz)], nid[(ix + 1, iy, iz)], nid[(ix + 1, iy + 1, iz)], nid[(ix, iy + 1, iz)],
                     nid[(ix, iy, iz + 1)], nid[(ix + 1, iy, iz + 1)], nid[(ix + 1, iy + 1, iz + 1)], nid[(ix, iy + 1, iz + 1)]]
                base = 1.0 + 0.3 * rnd.random() + (0.8 if (ix + iy) % 5 == 0 else 0.0)
                w = worsen * (1.0 if (ix * iy + iz) % 4 == 0 else 0.0) * (0.5 + rnd.random())
                elems[eid] = {"cfg": 208, "comp": 1, "nodes": n, "aspect": round(base + w, 4),
                              "jacobian": round(max(0.05, min(1.0, 0.95 - 0.3 * rnd.random() - 0.5 * w)), 4),
                              "skew": round(10 + 30 * rnd.random() + 25 * w, 3), "warpage": round(2 * rnd.random(), 3),
                              "volume": 1.0, "shortestside": 1.0}
                eid += 1
    for iy in range(ny):
        for ix in range(nx):
            n = [nid[(ix, iy, nz)], nid[(ix + 1, iy, nz)], nid[(ix + 1, iy + 1, nz)], nid[(ix, iy + 1, nz)]]
            w = worsen * (1.0 if (ix + iy) % 3 == 0 else 0.0)
            elems[eid] = {"cfg": 104, "comp": 2, "nodes": n, "aspect": round(1.1 + 0.4 * rnd.random() + w, 4),
                          "jacobian": round(max(0.1, 0.9 - 0.2 * rnd.random() - 0.3 * w), 4), "skew": round(5 + 20 * rnd.random() + 20 * w, 3),
                          "warpage": round(rnd.random(), 3), "taper": round(0.1 * rnd.random(), 3), "minangle": round(60 + 25 * rnd.random(), 2),
                          "maxangle": round(95 + 20 * rnd.random(), 2), "area": 1.0, "shortestside": 1.0}
            eid += 1
    for ix in range(3):
        elems[eid] = {"cfg": 11, "comp": 3, "nodes": [nid[(ix, 0, 0)], nid[(ix + 1, 0, 0)]], "length": 1.0}
        eid += 1
    for i in range(extra):
        # elementy tylko w tym modelu (bez odpowiednika)
        elems[eid] = {"cfg": 208, "comp": 1, "nodes": elems[1]["nodes"], "aspect": 2.0, "jacobian": 0.7, "skew": 20.0, "volume": 1.0, "shortestside": 1.0}
        eid += 1
    for i in range(drop_last):
        elems.pop(max(elems))
    top = [nid[(ix, iy, nz)] for iy in range(ny + 1) for ix in range(nx + 1)]
    d = {"comps": comps, "nodes": nodes, "elems": elems, "surfs": {"1": {"nodes": top}, "2": {"nodes": [nid[(0, iy, iz)] for iy in range(ny + 1) for iz in range(nz + 1)]}}}
    with io.open(path, "w", encoding="utf-8") as fh:
        json.dump(d, fh)
    return path


def make_pair(folder):
    os.makedirs(folder, exist_ok=True)
    ref = make_model(os.path.join(folder, "model_REF.hm"), seed=1)
    inf = make_model(os.path.join(folder, "model_INF.hm"), seed=1, morph=0.25, worsen=1.2, extra=2)
    return ref, inf


if __name__ == "__main__":
    import sys
    print(make_pair(sys.argv[1] if len(sys.argv) > 1 else "models"))
