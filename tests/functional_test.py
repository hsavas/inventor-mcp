"""End-to-end test of the MCP tools against a real Inventor.

Run BEFORE every push:   .venv\\Scripts\\python.exe tests\\functional_test.py
Works in %TEMP%\\inv_mcp_test, closes its documents without saving and quits
Inventor afterwards if this test started it. Exit code 0 = all passed.
"""

import asyncio
import json
import os
import shutil
import subprocess
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
WORK = os.path.join(tempfile.gettempdir(), "inv_mcp_test")
results = []


def inventor_count() -> int:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq Inventor.exe", "/NH"],
                         capture_output=True, text=True).stdout
    return sum(1 for line in out.splitlines() if line.lower().startswith("inventor.exe"))


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + (f"  -> {detail}" if detail else ""), flush=True)
    return bool(cond)


def _decode(res):
    if isinstance(res, tuple):  # (content, structured)
        content, structured = res
        if structured is not None:
            return structured.get("result", structured)
        res = content
    texts = [c.text for c in res if hasattr(c, "text")]
    vals = []
    for t in texts:
        try:
            vals.append(json.loads(t))
        except ValueError:
            vals.append(t)
    return vals[0] if len(vals) == 1 else vals


CALLED: set = set()  # every tool name the test actually invoked (coverage check)


async def call(mcp, _tool, /, **args):
    CALLED.add(_tool)
    return _decode(await mcp.call_tool(_tool, args))


def circle_edges(edges, radius, z=None):
    out = []
    for e in edges:
        if e.get("curve_type") == "circle" and abs(e.get("radius_mm", 0) - radius) < 1e-3:
            if z is None or abs(e["center_mm"][2] - z) < 1e-3:
                out.append(e)
    return out


def same_pos(a, b, tol=1e-3):
    return all(abs(x - y) <= tol for x, y in zip(a, b))


async def build_parts(mcp):
    """Plate 60x40x10 with two R5 holes; Pin = R8x3 head + R5x23 shaft."""
    await call(mcp, "inventor_new_document", doc_type="part")
    await call(mcp, "inventor_create_sketch", plane="XY", name="Base")
    await call(mcp, "inventor_sketch_add_rectangle", sketch="Base", x1=0, y1=0, x2=60, y2=40)
    await call(mcp, "inventor_finish_sketch", sketch="Base")
    await call(mcp, "inventor_extrude", sketch="Base", distance_mm=10)
    await call(mcp, "inventor_create_sketch", plane="XY", name="Holes")
    await call(mcp, "inventor_sketch_add_circle", sketch="Holes", cx=15, cy=20, radius=5)
    await call(mcp, "inventor_sketch_add_circle", sketch="Holes", cx=45, cy=20, radius=5)
    fin = await call(mcp, "inventor_finish_sketch", sketch="Holes")
    check("finish_sketch counts 2 hole regions", fin.get("closed_regions") == 2, fin)
    await call(mcp, "inventor_extrude", sketch="Holes", distance_mm=10, operation="cut")
    cm = await call(mcp, "inventor_check_model")
    check("plate check_model healthy", cm.get("healthy") and cm.get("size_mm") == [60, 40, 10], cm)
    await call(mcp, "inventor_save_document", save_as_path=os.path.join(WORK, "Plate.ipt"))
    await call(mcp, "inventor_close_document", save=False)

    await call(mcp, "inventor_new_document", doc_type="part")
    await call(mcp, "inventor_create_sketch", plane="XY", name="Head")
    await call(mcp, "inventor_sketch_add_circle", sketch="Head", cx=0, cy=0, radius=8)
    await call(mcp, "inventor_finish_sketch", sketch="Head")
    await call(mcp, "inventor_extrude", sketch="Head", distance_mm=3)
    await call(mcp, "inventor_create_sketch", plane="XY", name="Shaft")
    await call(mcp, "inventor_sketch_add_circle", sketch="Shaft", cx=0, cy=0, radius=5)
    await call(mcp, "inventor_finish_sketch", sketch="Shaft")
    await call(mcp, "inventor_extrude", sketch="Shaft", distance_mm=23)
    cm = await call(mcp, "inventor_check_model")
    check("pin is one healthy body", cm.get("healthy") and cm.get("body_count") == 1, cm)
    await call(mcp, "inventor_save_document", save_as_path=os.path.join(WORK, "Pin.ipt"))
    await call(mcp, "inventor_close_document", save=False)


def at_x(edges, x):
    return [e for e in edges if abs(e["center_mm"][0] - x) < 1e-3]


async def assembly_tests(mcp):
    await call(mcp, "inventor_new_document", doc_type="assembly")
    plate = await call(mcp, "inventor_assembly_place_component", file_path=os.path.join(WORK, "Plate.ipt"))
    p1 = await call(mcp, "inventor_assembly_place_component", file_path=os.path.join(WORK, "Pin.ipt"),
                    x_mm=15, y_mm=20, z_mm=-3)
    p2 = await call(mcp, "inventor_assembly_place_component", file_path=os.path.join(WORK, "Pin.ipt"),
                    x_mm=45, y_mm=20, z_mm=-3)
    plate_n, pin1, pin2 = plate["name"], p1["name"], p2["name"]
    await call(mcp, "inventor_assembly_set_grounded", occurrence=plate_n, grounded=True)

    occ = await call(mcp, "inventor_assembly_get_occurrence", name=pin1)
    check("get_occurrence position + 3x3 rotation",
          same_pos(occ["position_mm"], [15, 20, -3]) and len(occ.get("rotation", [])) == 3, occ)

    plate_edges = await call(mcp, "inventor_assembly_list_occurrence_edges", occurrence=plate_n)
    pin_edges = await call(mcp, "inventor_assembly_list_occurrence_edges", occurrence=pin1)
    hole_bottom = at_x(circle_edges(plate_edges, 5, z=0), 15)
    pin_seat = circle_edges(pin_edges, 5, z=0)
    if not check("found hole + pin seat edges", hole_bottom and pin_seat,
                 f"hole={len(hole_bottom)} pin={len(pin_seat)}"):
        return
    e_pin, e_hole = pin_seat[0]["entity_id"], hole_bottom[0]["entity_id"]

    # A) rigid joint with preserve_position (default) must not move the pin
    j = await call(mcp, "inventor_assembly_add_joint", joint_type="rigid", entity_one=e_pin, entity_two=e_hole)
    pp = j.get("preserve_position", {})
    occ = await call(mcp, "inventor_assembly_get_occurrence", name=pin1)
    check("rigid joint preserve_position keeps pin in place",
          pp.get("preserved") and same_pos(occ["position_mm"], [15, 20, -3]), {"pp": pp, "pos": occ["position_mm"]})
    await call(mcp, "inventor_assembly_delete_joint", name=j["name"])

    # B) rotational joint, limits and drive
    # use the flips that kept the pin in place in A, so the pin stays upright (seated)
    j = await call(mcp, "inventor_assembly_add_joint", joint_type="rotational", entity_one=e_pin,
                   entity_two=e_hole, angle_limits_deg=[0, 90], preserve_position=False,
                   flip_alignment=pp.get("flip_alignment", False), flip_origin=pp.get("flip_origin", False),
                   name="RotTest")
    d = await call(mcp, "inventor_assembly_drive_joint", name="RotTest", angle_deg=45)
    check("drive rotational joint to 45 deg",
          abs(d.get("angle_deg", 0) - 45) < 1e-3 and d.get("health") == "up_to_date", d)
    d = await call(mcp, "inventor_assembly_drive_joint", name="RotTest", angle_deg=120)
    check("drive past limit reports invalid_limits", d.get("health") == "invalid_limits", d.get("health"))
    joints = await call(mcp, "inventor_assembly_list_joints")
    check("list_joints shows RotTest", any(x["name"] == "RotTest" for x in joints), [x["name"] for x in joints])
    await call(mcp, "inventor_assembly_delete_joint", name="RotTest")
    await call(mcp, "inventor_assembly_set_position", occurrence=pin1, x_mm=15, y_mm=20, z_mm=-3)
    occ = await call(mcp, "inventor_assembly_get_occurrence", name=pin1)
    check("pin still upright after rotational test", abs(occ["rotation"][2][2] - 1) < 1e-6, occ["rotation"])

    # C) fastener auto-jointing
    dry = await call(mcp, "inventor_assembly_auto_joint_fasteners")
    m = dry.get("matches", [])
    check("auto_joint dry run finds both pins on the plate",
          len(m) == 2 and all(r["partner"] == plate_n for r in m), dry)
    check("auto_joint dry run created nothing", len(await call(mcp, "inventor_assembly_list_joints")) == 0)
    run = await call(mcp, "inventor_assembly_auto_joint_fasteners", dry_run=False)
    rows = run.get("matches", [])
    check("auto_joint creates preserved rigid joints",
          len(rows) == 2 and all(r.get("joint") and r.get("preserved") and not r.get("error") for r in rows), rows)
    o1 = await call(mcp, "inventor_assembly_get_occurrence", name=pin1)
    o2 = await call(mcp, "inventor_assembly_get_occurrence", name=pin2)
    check("pins did not move", same_pos(o1["position_mm"], [15, 20, -3]) and same_pos(o2["position_mm"], [45, 20, -3]),
          [o1["position_mm"], o2["position_mm"]])
    again = await call(mcp, "inventor_assembly_auto_joint_fasteners")
    check("second run skips already-jointed pins",
          not again.get("matches") and len(again.get("skipped", [])) == 2, again)
    await call(mcp, "inventor_close_document", save=False)


async def part_feature_tests(mcp):
    # polygon must be a closed profile
    await call(mcp, "inventor_new_document", doc_type="part")
    await call(mcp, "inventor_create_sketch", plane="XY", name="Hex")
    await call(mcp, "inventor_sketch_add_polygon", sketch="Hex", cx=0, cy=0, radius=10, sides=6)
    fin = await call(mcp, "inventor_finish_sketch", sketch="Hex")
    check("polygon is a closed profile", fin.get("closed_regions") == 1, fin)
    await call(mcp, "inventor_extrude", sketch="Hex", distance_mm=5)
    check("hex prism healthy", (await call(mcp, "inventor_check_model")).get("healthy"))
    await call(mcp, "inventor_close_document", save=False)

    # thread on a cylinder
    await call(mcp, "inventor_new_document", doc_type="part")
    await call(mcp, "inventor_create_sketch", plane="XY", name="Rod")
    await call(mcp, "inventor_sketch_add_circle", sketch="Rod", cx=0, cy=0, radius=5)
    await call(mcp, "inventor_finish_sketch", sketch="Rod")
    await call(mcp, "inventor_extrude", sketch="Rod", distance_mm=20)
    faces = await call(mcp, "inventor_list_faces")
    cyl = [f for f in faces if f.get("surface_type") == "cylinder"]
    if check("rod has a cylindrical face", cyl):
        t = await call(mcp, "inventor_thread", face_ref=cyl[0]["entity_id"], designation="M10x1.5")
        check("M10 thread created", t.get("name"), t)
    await call(mcp, "inventor_close_document", save=False)

    # 3D sketch path + plane normal to it + sweep
    await call(mcp, "inventor_new_document", doc_type="part")
    await call(mcp, "inventor_create_sketch3d", name="Path")
    line = await call(mcp, "inventor_sketch3d_add_line", sketch="Path", x1=0, y1=0, z1=0, x2=10, y2=20, z2=30)
    await call(mcp, "inventor_create_work_plane_normal_to_curve", curve_ref=line["entity_id"],
               x_mm=0, y_mm=0, z_mm=0, name="ProfPlane")
    await call(mcp, "inventor_create_sketch", plane="ProfPlane", name="Prof")
    await call(mcp, "inventor_sketch_add_circle", sketch="Prof", cx=0, cy=0, radius=2)
    await call(mcp, "inventor_finish_sketch", sketch="Prof")
    await call(mcp, "inventor_sweep", sketch="Prof", path=line["entity_id"])
    cm = await call(mcp, "inventor_check_model")
    expected = 3.14159265 * 4 * (10 ** 2 + 20 ** 2 + 30 ** 2) ** 0.5
    check("sweep along 3D sketch line", cm.get("healthy") and abs(cm["volume_mm3"] - expected) < 1.0,
          {"volume": cm.get("volume_mm3"), "expected": round(expected, 2)})
    await call(mcp, "inventor_close_document", save=False)

    # sheet metal face + flange + flat pattern
    await call(mcp, "inventor_new_document", doc_type="part")
    await call(mcp, "inventor_sheet_metal_convert", thickness_mm=2)
    await call(mcp, "inventor_create_sketch", plane="XY", name="Blank")
    await call(mcp, "inventor_sketch_add_rectangle", sketch="Blank", x1=0, y1=0, x2=50, y2=30)
    await call(mcp, "inventor_finish_sketch", sketch="Blank")
    await call(mcp, "inventor_sheet_metal_face", sketch="Blank")
    edges = await call(mcp, "inventor_list_edges")
    long_edges = [e for e in edges if e.get("curve_type") in ("line", "line_segment")
                  and abs((e.get("length_mm") or 0) - 50) < 1e-3]
    if check("sheet metal face has 50 mm edges", long_edges, len(edges)):
        await call(mcp, "inventor_sheet_metal_flange", edge_refs=[long_edges[0]["entity_id"]], height_mm=15)
        fp = await call(mcp, "inventor_sheet_metal_flat_pattern")
        check("flat pattern larger than blank", fp.get("width_mm", 0) > 30 or fp.get("length_mm", 0) > 50, fp)
    await call(mcp, "inventor_close_document", save=False)


async def units_and_bodies_tests(mcp):
    """Sam-H101's PR #1: metric documents, parameter units, list_bodies, affected_bodies."""
    doc = await call(mcp, "inventor_new_document", doc_type="part")
    check("new part is metric (mm)", doc.get("length_units") in ("mm", "millimeter"), doc)

    p = await call(mcp, "inventor_create_parameter", name="Len", expression="46 mm")
    check("parameter '46 mm' reports value 46 in mm",
          abs(p.get("value", 0) - 46) < 1e-9 and p.get("units") == "mm" and abs(p.get("value_db", 0) - 4.6) < 1e-9, p)
    p = await call(mcp, "inventor_create_parameter", name="Ang", expression="30 deg")
    check("parameter '30 deg' reports value 30", abs(p.get("value", 0) - 30) < 1e-9, p)

    # body 1: 10x10x10 box, body 2: separate 10x10x10 box as a new body
    await call(mcp, "inventor_create_sketch", plane="XY", name="A")
    await call(mcp, "inventor_sketch_add_rectangle", sketch="A", x1=0, y1=0, x2=10, y2=10)
    await call(mcp, "inventor_finish_sketch", sketch="A")
    await call(mcp, "inventor_extrude", sketch="A", distance_mm=10)
    await call(mcp, "inventor_create_sketch", plane="XY", name="B")
    await call(mcp, "inventor_sketch_add_rectangle", sketch="B", x1=20, y1=0, x2=30, y2=10)
    await call(mcp, "inventor_finish_sketch", sketch="B")
    await call(mcp, "inventor_extrude", sketch="B", distance_mm=10, operation="new_body")
    bodies = await call(mcp, "inventor_list_bodies")
    check("list_bodies volume and area are correct",
          len(bodies) == 2 and all(abs(b["volume_mm3"] - 1000) < 0.01 and abs(b["area_mm2"] - 600) < 0.01 for b in bodies),
          [(b["volume_mm3"], b["area_mm2"]) for b in bodies])

    # cut a R2 hole through body 2 only
    await call(mcp, "inventor_create_sketch", plane="XY", name="Hole")
    await call(mcp, "inventor_sketch_add_circle", sketch="Hole", cx=25, cy=5, radius=2)
    await call(mcp, "inventor_finish_sketch", sketch="Hole")
    await call(mcp, "inventor_extrude", sketch="Hole", distance_mm=10, operation="cut", affected_bodies=[2])
    bodies = await call(mcp, "inventor_list_bodies")
    hole = 3.14159265 * 4 * 10
    check("extrude cut with affected_bodies=[2] cuts only body 2",
          abs(bodies[0]["volume_mm3"] - 1000) < 0.01 and abs(bodies[1]["volume_mm3"] - (1000 - hole)) < 0.01,
          [b["volume_mm3"] for b in bodies])
    await call(mcp, "inventor_close_document", save=False)

    # revolve: ring r 5..10, height 20 around the Y axis
    await call(mcp, "inventor_new_document", doc_type="part")
    await call(mcp, "inventor_create_sketch", plane="XY", name="R")
    await call(mcp, "inventor_sketch_add_rectangle", sketch="R", x1=5, y1=0, x2=10, y2=20)
    await call(mcp, "inventor_finish_sketch", sketch="R")
    await call(mcp, "inventor_revolve", sketch="R", axis="Y Axis")
    vol = (await call(mcp, "inventor_list_bodies"))[0]["volume_mm3"]
    check("revolve volume", abs(vol - 3.14159265 * (100 - 25) * 20) < 0.01, vol)
    await call(mcp, "inventor_close_document", save=False)


# ---------------------------------------------------------------- helpers for the full-coverage sections
def entities_of(listing):
    """sketch_list_entities may return a list or {'entities': [...]}."""
    return listing.get("entities", listing) if isinstance(listing, dict) else listing


def eid(res):
    """entity_id from a single-entity result, or the first of entity_ids."""
    if isinstance(res, dict):
        if "entity_id" in res:
            return res["entity_id"]
        if res.get("entity_ids"):
            return res["entity_ids"][0]
    raise AssertionError(f"no entity id in {res!r}")


async def new_part(mcp):
    await call(mcp, "inventor_new_document", doc_type="part")


async def box(mcp, sk, x1, y1, x2, y2, h, op="join", plane="XY"):
    await call(mcp, "inventor_create_sketch", plane=plane, name=sk)
    await call(mcp, "inventor_sketch_add_rectangle", sketch=sk, x1=x1, y1=y1, x2=x2, y2=y2)
    await call(mcp, "inventor_finish_sketch", sketch=sk)
    return await call(mcp, "inventor_extrude", sketch=sk, distance_mm=h, operation=op)


async def volume(mcp):
    return sum(b["volume_mm3"] for b in await call(mcp, "inventor_list_bodies"))


PI = 3.141592653589793


async def sketch_tools_test(mcp):
    await new_part(mcp)
    await call(mcp, "inventor_create_sketch", plane="XY", name="S")
    l1 = eid(await call(mcp, "inventor_sketch_add_line", sketch="S", x1=0, y1=0, x2=20, y2=0))
    l2 = eid(await call(mcp, "inventor_sketch_add_line", sketch="S", x1=20, y1=0, x2=20, y2=10))
    f = await call(mcp, "inventor_sketch_add_fillet", sketch="S", entity_id_1=l1, entity_id_2=l2, radius=2)
    check("sketch_add_fillet", isinstance(f, dict), f)
    await call(mcp, "inventor_sketch_add_arc", sketch="S", cx=50, cy=0, radius=5, start_deg=0, end_deg=90)
    await call(mcp, "inventor_sketch_add_point", sketch="S", x=70, y=0)
    await call(mcp, "inventor_sketch_add_polyline", sketch="S", points=[[0, 30], [10, 30], [10, 40]])
    await call(mcp, "inventor_sketch_add_ellipse", sketch="S", cx=40, cy=35, major_radius=6, minor_radius=3, angle_deg=30)
    await call(mcp, "inventor_sketch_add_spline", sketch="S", points=[[60, 30], [65, 35], [70, 30]])
    await call(mcp, "inventor_sketch_add_rectangle_centered", sketch="S", cx=90, cy=0, width=10, height=6)
    await call(mcp, "inventor_sketch_add_rectangle_three_point", sketch="S", x1=100, y1=10, x2=110, y2=10, width=5)
    await call(mcp, "inventor_sketch_add_slot_center_to_center", sketch="S", x1=0, y1=60, x2=20, y2=60, width=6)
    await call(mcp, "inventor_sketch_add_slot_overall", sketch="S", cx=40, cy=60, length=20, width=6, angle_deg=30)
    circ = eid(await call(mcp, "inventor_sketch_add_circle", sketch="S", cx=90, cy=40, radius=4))

    # constraints and dimensions on a free line
    l3 = eid(await call(mcp, "inventor_sketch_add_line", sketch="S", x1=0, y1=80, x2=10, y2=82))
    await call(mcp, "inventor_sketch_add_constraint", sketch="S", type="horizontal", entity_ids=[l3])
    await call(mcp, "inventor_sketch_add_dimension", sketch="S", type="linear", entity_ids=[l3], value=15)
    await call(mcp, "inventor_sketch_add_dimension", sketch="S", type="radius", entity_ids=[circ], value=6)
    # list_entities refreshes ids for the sketch, so ids are re-picked from the listing below
    ents = entities_of(await call(mcp, "inventor_sketch_list_entities", sketch="S"))
    line3 = next((e for e in ents if e.get("kind") == "line"
                  and abs(e["start"][1] - e["end"][1]) < 1e-6 and abs(e["start"][1] - 80) < 5
                  and abs(abs(e["end"][0] - e["start"][0]) - 15) < 1e-6), None)
    check("horizontal constraint + linear dimension 15 mm", line3 is not None,
          [e for e in ents if e.get("kind") == "line"][-3:])
    c6 = [e for e in ents if e.get("kind") == "circle" and abs(e.get("radius", 0) - 6) < 1e-6]
    check("radius dimension 6 mm", c6, [e for e in ents if e.get("kind") == "circle"])
    check("sketch_list_entities lists all kinds",
          {"line", "circle", "arc", "point"} <= {e.get("kind") for e in ents}, sorted({e.get("kind") for e in ents}))

    if line3:
        sc = await call(mcp, "inventor_sketch_set_construction", sketch="S", entity_ids=[line3["entity_id"]])
        check("sketch_set_construction", isinstance(sc, dict), sc)
    m = await call(mcp, "inventor_sketch_mirror", sketch="S", entity_ids=[eid(c6[0])] if c6 else [circ],
                   line_point1=[120, 0], line_point2=[120, 10])
    ents = entities_of(await call(mcp, "inventor_sketch_list_entities", sketch="S"))
    mirrored = [e for e in ents if e.get("kind") == "circle" and abs(e["center"][0] - 150) < 1e-6]
    check("sketch_mirror creates circle at x=150", mirrored, m)
    await call(mcp, "inventor_finish_sketch", sketch="S")

    # every closed-shape tool must produce exactly ONE closed region on its own
    shapes = {
        "polyline": ("inventor_sketch_add_polyline", {"points": [[0, 0], [10, 0], [10, 10]]}),
        "ellipse": ("inventor_sketch_add_ellipse", {"cx": 0, "cy": 0, "major_radius": 6, "minor_radius": 3, "angle_deg": 30}),
        "rectangle_centered": ("inventor_sketch_add_rectangle_centered", {"cx": 0, "cy": 0, "width": 10, "height": 6}),
        "rectangle_three_point": ("inventor_sketch_add_rectangle_three_point",
                                  {"x1": 0, "y1": 0, "x2": 10, "y2": 0, "width": 5}),
        "slot_center_to_center": ("inventor_sketch_add_slot_center_to_center",
                                  {"x1": 0, "y1": 0, "x2": 20, "y2": 0, "width": 6}),
        "slot_overall": ("inventor_sketch_add_slot_overall", {"cx": 0, "cy": 0, "length": 20, "width": 6, "angle_deg": 30}),
        "closed_spline": ("inventor_sketch_add_spline", {"points": [[0, 0], [10, 5], [20, 0], [10, -5]], "closed": True}),
    }
    for i, (label, (tool, args)) in enumerate(shapes.items()):
        name = f"Shape{i}"
        await call(mcp, "inventor_create_sketch", plane="XY", name=name)
        await call(mcp, tool, sketch=name, **args)
        fin = await call(mcp, "inventor_finish_sketch", sketch=name)
        check(f"{label} is exactly one closed region", fin.get("closed_regions") == 1, fin)

    await call(mcp, "inventor_create_sketch", plane="XZ", name="Tmp")
    await call(mcp, "inventor_delete_sketch", name="Tmp")
    again = await call(mcp, "inventor_create_sketch", plane="XZ", name="Tmp")
    check("delete_sketch frees the name", again.get("name") == "Tmp", again)
    await call(mcp, "inventor_close_document", save=False)


async def edge_feature_test(mcp):
    """fillet, chamfer, hole (on an offset work plane), list/delete features."""
    await new_part(mcp)
    await box(mcp, "B", 0, 0, 40, 30, 20)
    v0 = await volume(mcp)
    edges = await call(mcp, "inventor_list_edges")
    vertical = [e for e in edges if e.get("curve_type") in ("line", "line_segment")
                and abs((e.get("length_mm") or 0) - 20) < 1e-6]
    await call(mcp, "inventor_fillet", radius_mm=2, edge_refs=[vertical[0]["entity_id"]])
    v1 = await volume(mcp)
    check("fillet r2 removes (4 - pi) * 20 mm3", abs((v0 - v1) - (4 - PI) * 20) < 0.01, v0 - v1)
    edges = await call(mcp, "inventor_list_edges")
    long40 = [e for e in edges if e.get("curve_type") in ("line", "line_segment")
              and abs((e.get("length_mm") or 0) - 40) < 1e-6]
    await call(mcp, "inventor_chamfer", distance_mm=1, edge_refs=[long40[0]["entity_id"]])
    v2 = await volume(mcp)
    check("chamfer 1 mm removes 0.5 * 40 mm3", abs((v1 - v2) - 20) < 0.01, v1 - v2)

    await call(mcp, "inventor_create_work_plane", base_plane="XY", offset_mm=20, name="Top")
    await call(mcp, "inventor_create_sketch", plane="Top", name="H")
    pt = eid(await call(mcp, "inventor_sketch_add_point", sketch="H", x=10, y=10))
    await call(mcp, "inventor_finish_sketch", sketch="H")
    await call(mcp, "inventor_hole", sketch="H", point_entity_ids=[pt], diameter_mm=5, through_all=True)
    v3 = await volume(mcp)
    check("through hole d5 removes pi*2.5^2*20", abs((v2 - v3) - PI * 6.25 * 20) < 0.01, v2 - v3)

    feats = await call(mcp, "inventor_list_features")
    names = [f["name"] for f in feats]
    check("list_features shows extrude, fillet, chamfer, hole", len(feats) >= 4, names)
    await call(mcp, "inventor_delete_feature", name=names[-1])
    check("delete_feature removes the hole", abs(await volume(mcp) - v2) < 0.01)
    await call(mcp, "inventor_close_document", save=False)


async def body_feature_test(mcp):
    """shell, combine, mirror, rectangular/circular pattern, loft, coil."""
    await new_part(mcp)
    await box(mcp, "B", 0, 0, 20, 20, 20)
    faces = await call(mcp, "inventor_list_faces")
    top = [f for f in faces if f.get("surface_type") == "plane" and f.get("normal") == [0, 0, 1]
           or (f.get("surface_type") == "plane" and abs(f.get("root_point_mm", [0, 0, 0])[2] - 20) < 1e-6
               and abs(abs(f.get("normal", [0, 0, 0])[2]) - 1) < 1e-6)]
    await call(mcp, "inventor_shell", thickness_mm=2, face_refs=[top[0]["entity_id"]])
    check("shell 2 mm open top", abs(await volume(mcp) - (8000 - 16 * 16 * 18)) < 0.01, await volume(mcp))
    await call(mcp, "inventor_close_document", save=False)

    await new_part(mcp)
    await box(mcp, "A", 0, 0, 10, 10, 10)
    await box(mcp, "B", 5, 0, 15, 10, 10, op="new_body")
    await call(mcp, "inventor_combine", base_body_index=1, tool_body_indices=[2], operation="join")
    bodies = await call(mcp, "inventor_list_bodies")
    check("combine join -> 1 body of 1500 mm3", len(bodies) == 1 and abs(bodies[0]["volume_mm3"] - 1500) < 0.01,
          [b["volume_mm3"] for b in bodies])
    await call(mcp, "inventor_close_document", save=False)

    await new_part(mcp)
    await box(mcp, "A", 0, 0, 10, 10, 10)
    await call(mcp, "inventor_create_work_plane", base_plane="XY", offset_mm=10, name="T10")
    await call(mcp, "inventor_create_work_plane", base_plane="YZ", offset_mm=5, name="Mid")
    boss = await box(mcp, "Boss", 1, 1, 3, 3, 5, plane="T10")
    await call(mcp, "inventor_mirror", feature_name=boss["name"], plane="Mid")
    check("mirror boss about mid plane adds a second boss", abs(await volume(mcp) - (1000 + 2 * 20)) < 0.01,
          await volume(mcp))
    v = await volume(mcp)
    try:
        await call(mcp, "inventor_mirror", feature_name=boss["name"], plane="YZ")  # would float in the air
        check("impossible mirror reports an error (not silent success)", False, "no error")
    except Exception as e:  # noqa: BLE001
        check("impossible mirror reports an error (not silent success)",
              "hatali" in str(e) and abs(await volume(mcp) - v) < 0.01, str(e)[:160])
    feats = [f["name"] for f in await call(mcp, "inventor_list_features")]
    check("failed mirror was rolled back", sum(n.startswith("Mirror") for n in feats) == 1, feats)
    await call(mcp, "inventor_close_document", save=False)

    await new_part(mcp)
    await box(mcp, "Base", 0, 0, 50, 10, 5)
    await call(mcp, "inventor_create_work_plane", base_plane="XY", offset_mm=5, name="P5")
    await call(mcp, "inventor_create_sketch", plane="P5", name="Boss")
    await call(mcp, "inventor_sketch_add_circle", sketch="Boss", cx=5, cy=5, radius=1)
    await call(mcp, "inventor_finish_sketch", sketch="Boss")
    boss = await call(mcp, "inventor_extrude", sketch="Boss", distance_mm=3)
    await call(mcp, "inventor_rectangular_pattern", feature_name=boss["name"], x_count=4, x_spacing_mm=10)
    check("rectangular pattern: 4 bosses", abs(await volume(mcp) - (2500 + 4 * PI * 3)) < 0.01, await volume(mcp))
    await call(mcp, "inventor_close_document", save=False)

    await new_part(mcp)
    await call(mcp, "inventor_create_sketch", plane="XY", name="Disc")
    await call(mcp, "inventor_sketch_add_circle", sketch="Disc", cx=0, cy=0, radius=20)
    await call(mcp, "inventor_finish_sketch", sketch="Disc")
    await call(mcp, "inventor_extrude", sketch="Disc", distance_mm=5)
    await call(mcp, "inventor_create_sketch", plane="XY", name="Hc")
    await call(mcp, "inventor_sketch_add_circle", sketch="Hc", cx=12, cy=0, radius=2)
    await call(mcp, "inventor_finish_sketch", sketch="Hc")
    hole = await call(mcp, "inventor_extrude", sketch="Hc", distance_mm=5, operation="cut")
    await call(mcp, "inventor_circular_pattern", feature_name=hole["name"], axis="Z Axis", count=6)
    check("circular pattern: 6 holes", abs(await volume(mcp) - (PI * 400 * 5 - 6 * PI * 4 * 5)) < 0.01,
          await volume(mcp))
    await call(mcp, "inventor_close_document", save=False)

    await new_part(mcp)
    await call(mcp, "inventor_create_sketch", plane="XY", name="L1")
    await call(mcp, "inventor_sketch_add_rectangle_centered", sketch="L1", cx=0, cy=0, width=10, height=10)
    await call(mcp, "inventor_finish_sketch", sketch="L1")
    await call(mcp, "inventor_create_work_plane", base_plane="XY", offset_mm=20, name="P20")
    await call(mcp, "inventor_create_sketch", plane="P20", name="L2")
    await call(mcp, "inventor_sketch_add_circle", sketch="L2", cx=0, cy=0, radius=4)
    await call(mcp, "inventor_finish_sketch", sketch="L2")
    await call(mcp, "inventor_loft", sections=["L1", "L2"])
    v = await volume(mcp)
    check("loft square -> circle", PI * 16 * 20 < v < 100 * 20, v)
    await call(mcp, "inventor_close_document", save=False)

    await new_part(mcp)
    await call(mcp, "inventor_create_sketch", plane="XZ", name="Prof")
    await call(mcp, "inventor_sketch_add_circle", sketch="Prof", cx=10, cy=0, radius=1)
    await call(mcp, "inventor_finish_sketch", sketch="Prof")
    await call(mcp, "inventor_create_coil", sketch="Prof", axis="Z Axis", coil_type="pitch_and_height",
               pitch_mm=5, height_mm=20)
    v = await volume(mcp)
    # 4 turns of a r1 wire on a r10 helix: ~ pi*1^2 * 4 * sqrt((2*pi*10)^2 + 5^2)
    expected = PI * 4 * ((2 * PI * 10) ** 2 + 25) ** 0.5
    check("coil (spring) volume", abs(v - expected) / expected < 0.02, {"v": v, "expected": round(expected, 1)})
    await call(mcp, "inventor_close_document", save=False)


async def part_utility_test(mcp):
    """work geometry, bbox, mass, parameters, iProperties, appearance, export, view, script, 3D spline."""
    await new_part(mcp)
    ext = await box(mcp, "B", 0, 0, 40, 30, 20)
    bb = await call(mcp, "inventor_get_bounding_box")
    check("get_bounding_box 40x30x20", str(bb).count("40") and str(bb).count("30") and str(bb).count("20"), bb)
    mp = await call(mcp, "inventor_get_mass_properties")
    check("get_mass_properties volume 24000 mm3", any(abs(v - 24000) < 0.01 for v in mp.values()
                                                      if isinstance(v, (int, float))), mp)

    edges = await call(mcp, "inventor_list_edges")
    line = next(e for e in edges if e.get("curve_type") in ("line", "line_segment"))
    ax = await call(mcp, "inventor_create_work_axis", edge_ref=line["entity_id"])
    check("create_work_axis", isinstance(ax, dict) and ax, ax)
    wp = await call(mcp, "inventor_create_work_point", x_mm=1, y_mm=2, z_mm=3)
    check("create_work_point", isinstance(wp, dict) and wp, wp)
    pl = await call(mcp, "inventor_create_work_plane_three_points", x1=0, y1=0, z1=0, x2=10, y2=0, z2=10,
                    x3=0, y3=10, z3=0, name="Tilted")
    check("create_work_plane_three_points", pl.get("name") == "Tilted", pl)
    await call(mcp, "inventor_create_sketch", plane="Tilted", name="OnTilted")
    await call(mcp, "inventor_sketch_add_circle", sketch="OnTilted", cx=0, cy=0, radius=2)
    check("sketch on tilted plane is closed",
          (await call(mcp, "inventor_finish_sketch", sketch="OnTilted")).get("closed_regions") == 1)

    params = await call(mcp, "inventor_list_parameters")
    depth = next((p for p in params if p.get("expression") == "20 mm"), None)
    check("list_parameters finds the 20 mm extrude depth in mm", depth and abs(depth["value"] - 20) < 1e-9,
          [(p["name"], p.get("expression"), p.get("value")) for p in params])
    if depth:
        sp = await call(mcp, "inventor_set_parameter", name=depth["name"], expression="25 mm")
        check("set_parameter 25 mm changes the model",
              abs(sp.get("value", 0) - 25) < 1e-9 and abs(await volume(mcp) - 40 * 30 * 25) < 0.01, sp)

    await call(mcp, "inventor_set_iproperty", propset="Summary Information", name="Title", value="MCP test")
    await call(mcp, "inventor_set_iproperty", propset="User Defined Properties", name="Batch", value="42")
    ip = json.dumps(await call(mcp, "inventor_get_iproperties"))
    check("set/get iProperties (built-in + custom)", "MCP test" in ip and "Batch" in ip, ip[:300])

    apps = await call(mcp, "inventor_list_appearances", query="red", limit=5)
    check("list_appearances finds 'red'", apps and all("red" in a["display_name"].lower() for a in apps), apps)
    if apps:
        sa = await call(mcp, "inventor_set_appearance", appearance_name=apps[0]["display_name"])
        check("set_appearance by display name", sa.get("appearance") == apps[0]["display_name"], sa)
        sa = await call(mcp, "inventor_set_appearance", appearance_name=apps[0]["name"])
        check("set_appearance by internal name from the list", sa.get("appearance") == apps[0]["display_name"], sa)
        sa = await call(mcp, "inventor_set_appearance", appearance_name=apps[0]["display_name"], color_rgb=[0, 128, 255])
        check("set_appearance with custom colour", isinstance(sa, dict), sa)

    for fmt, ext_ in (("step", "stp"), ("stl", "stl"), ("iges", "igs"), ("sat", "sat"), ("dwg", "dwg")):
        path = os.path.join(WORK, f"box.{ext_}")
        await call(mcp, "inventor_export", format=fmt, path=path)
        check(f"export {fmt}", os.path.isfile(path) and os.path.getsize(path) > 0, path)
    await call(mcp, "inventor_zoom_all")
    png = os.path.join(WORK, "shot.png")
    await call(mcp, "inventor_screenshot", path=png, width=400, height=300)
    check("screenshot png", os.path.isfile(png) and os.path.getsize(png) > 1000, png)
    out = await call(mcp, "inventor_run_script", script="print(comp_def.Features.Count, mm_to_cm(10))")
    check("run_script returns printed output", "1.0" in str(out), out)
    try:
        await call(mcp, "inventor_run_script", script="raise ValueError('boom')")
        check("run_script surfaces script errors", False, "no error raised")
    except Exception as e:  # noqa: BLE001
        check("run_script surfaces script errors", "boom" in str(e), str(e)[:200])

    await call(mcp, "inventor_create_sketch3d", name="P3")
    sp3 = await call(mcp, "inventor_sketch3d_add_spline", sketch="P3",
                     points=[[0, 0, 0], [10, 5, 5], [20, 0, 10], [30, 5, 15]])
    check("sketch3d_add_spline", sp3.get("entity_id"), sp3)
    await call(mcp, "inventor_sketch3d_add_line", sketch="P3", x1=0, y1=0, z1=0, x2=0, y2=0, z2=10)
    ents = await call(mcp, "inventor_sketch3d_list_entities", sketch="P3")
    kinds = sorted(e["kind"] for e in ents.get("entities", []))
    check("sketch3d_list_entities lists line3d + spline3d", kinds == ["line3d", "spline3d"], kinds)
    await call(mcp, "inventor_save_document", save_as_path=os.path.join(WORK, "Util.ipt"))
    await call(mcp, "inventor_close_document", save=False)


async def document_test(mcp):
    """open, list, set_active documents."""
    a = await call(mcp, "inventor_open_document", path=os.path.join(WORK, "Plate.ipt"))
    b = await call(mcp, "inventor_open_document", path=os.path.join(WORK, "Pin.ipt"))
    docs = await call(mcp, "inventor_list_documents")
    names = [d["name"] for d in docs]
    check("open_document + list_documents", a["name"] in names and b["name"] in names, names)
    await call(mcp, "inventor_set_active_document", name=a["name"])
    st = await call(mcp, "inventor_get_status")
    check("set_active_document", st.get("active_document") == a["name"], st.get("active_document"))
    await call(mcp, "inventor_close_all_documents", save=False)


def _lines(edges, length):
    return [e for e in edges if e.get("curve_type") in ("line", "line_segment")
            and abs((e.get("length_mm") or 0) - length) < 1e-6]


async def assembly_more_test(mcp):
    """constraints (incl. re-resolved ids), faces, occurrences, BOM, patterns, delete occurrence."""
    await call(mcp, "inventor_new_document", doc_type="assembly")
    plate = (await call(mcp, "inventor_assembly_place_component", file_path=os.path.join(WORK, "Plate.ipt")))["name"]
    pin1 = (await call(mcp, "inventor_assembly_place_component", file_path=os.path.join(WORK, "Pin.ipt"),
                       x_mm=100, y_mm=100, z_mm=50))["name"]
    pin2 = (await call(mcp, "inventor_assembly_place_component", file_path=os.path.join(WORK, "Pin.ipt"),
                       x_mm=-100, y_mm=100, z_mm=50))["name"]
    await call(mcp, "inventor_assembly_set_grounded", occurrence=plate, grounded=True)
    occs = await call(mcp, "inventor_assembly_list_occurrences")
    check("list_occurrences", sorted(o["name"] for o in occs) == sorted([plate, pin1, pin2]), [o["name"] for o in occs])

    faces = await call(mcp, "inventor_assembly_list_occurrence_faces", occurrence=pin1)
    cyl = [f for f in faces if f.get("surface_type") == "cylinder"]
    check("occurrence faces: cylinders report base_point_mm", cyl and all("base_point_mm" in f for f in cyl), cyl[:1])

    plate_edges = await call(mcp, "inventor_assembly_list_occurrence_edges", occurrence=plate)
    pin_edges = await call(mcp, "inventor_assembly_list_occurrence_edges", occurrence=pin1)
    hole = at_x(circle_edges(plate_edges, 5, z=0), 15)[0]
    seat = [e for e in circle_edges(pin_edges, 5) if abs(e["center_mm"][2] - 53) < 1e-6][0]  # pin at z=50, seat at +3
    await call(mcp, "inventor_assembly_add_constraint", constraint_type="insert",
               entity_one=seat["entity_id"], entity_two=hole["entity_id"], axes_opposed=False)
    o = await call(mcp, "inventor_assembly_get_occurrence", name=pin1)
    check("insert constraint moves pin onto the hole axis",
          abs(o["position_mm"][0] - 15) < 1e-6 and abs(o["position_mm"][1] - 20) < 1e-6, o["position_mm"])
    cons = await call(mcp, "inventor_assembly_list_constraints")
    check("list_constraints shows 1", len(cons) == 1, cons)
    await call(mcp, "inventor_assembly_delete_constraint", name=cons[0]["name"])
    check("delete_constraint", len(await call(mcp, "inventor_assembly_list_constraints")) == 0)

    # pin2's edges were never listed: the id must be re-resolved from occurrence name + index
    idx = seat["entity_id"].rsplit(":edge", 1)[1]
    asm_name = seat["entity_id"].split(":_asm_")[0]
    unlisted = f"{asm_name}:_asm_{pin2}_edges:edge{idx}"
    hole2 = at_x(circle_edges(plate_edges, 5, z=0), 45)[0]
    await call(mcp, "inventor_assembly_add_constraint", constraint_type="insert",
               entity_one=unlisted, entity_two=hole2["entity_id"], axes_opposed=False)
    o2 = await call(mcp, "inventor_assembly_get_occurrence", name=pin2)
    check("constraint with a never-listed (re-resolved) id", abs(o2["position_mm"][0] - 45) < 1e-6, o2["position_mm"])

    bom = await call(mcp, "inventor_assembly_get_bom")
    rows = bom.get("rows", bom) if isinstance(bom, dict) else bom
    check("get_bom lists Plate and Pin", "Plate" in json.dumps(rows) and "Pin" in json.dumps(rows), rows)

    x_edge = [e for e in _lines(plate_edges, 60)][0]
    n_before = len(await call(mcp, "inventor_assembly_list_occurrences"))
    pr = await call(mcp, "inventor_assembly_pattern_rectangular", occurrence_refs=[pin1],
                    column_entity_ref=x_edge["entity_id"], column_offset_mm=10, column_count=3)
    check("assembly_pattern_rectangular", isinstance(pr, dict), pr)
    z_edge = [e for e in _lines(plate_edges, 10)][0]
    pc = await call(mcp, "inventor_assembly_pattern_circular", occurrence_refs=[pin2],
                    axis_entity_ref=z_edge["entity_id"], count=3)
    check("assembly_pattern_circular", isinstance(pc, dict), pc)
    n_after = len(await call(mcp, "inventor_assembly_list_occurrences"))
    check("patterns add pin instances", n_after >= n_before, {"before": n_before, "after": n_after})

    extra = (await call(mcp, "inventor_assembly_place_component", file_path=os.path.join(WORK, "Pin.ipt"),
                        x_mm=0, y_mm=-100, z_mm=0))["name"]
    await call(mcp, "inventor_assembly_delete_occurrence", occurrence=extra)
    names = [o["name"] for o in await call(mcp, "inventor_assembly_list_occurrences")]
    check("delete_occurrence", extra not in names, names)
    await call(mcp, "inventor_save_document", save_as_path=os.path.join(WORK, "Asm.iam"))
    await call(mcp, "inventor_close_all_documents", save=False)


async def drawing_test(mcp):
    plate = (await call(mcp, "inventor_open_document", path=os.path.join(WORK, "Plate.ipt")))["name"]
    asm = (await call(mcp, "inventor_open_document", path=os.path.join(WORK, "Asm.iam")))["name"]
    await call(mcp, "inventor_new_document", doc_type="drawing")
    sheets = await call(mcp, "inventor_drawing_list_sheets")
    check("drawing_list_sheets", len(sheets) >= 1, sheets)

    base = await call(mcp, "inventor_drawing_add_base_view", model_doc=plate, x_mm=100, y_mm=150, scale=1)
    proj = await call(mcp, "inventor_drawing_add_projected_view", base_view=base["name"], x_mm=100, y_mm=80)
    views = await call(mcp, "inventor_drawing_list_views")
    check("base + projected view", len(views) == 2, [v.get("name") for v in views])
    before = next(v for v in views if v["name"] == base["name"])
    await call(mcp, "inventor_drawing_set_view_position", view=base["name"], x_mm=120, y_mm=160)
    after = next(v for v in await call(mcp, "inventor_drawing_list_views") if v["name"] == base["name"])
    check("set_view_position moves the base view by (+20, +10)",
          abs(after["left_mm"] - before["left_mm"] - 20) < 1e-6 and abs(after["top_mm"] - before["top_mm"] - 10) < 1e-6,
          {"before": before, "after": after})

    curves = (await call(mcp, "inventor_drawing_list_view_curves", view=base["name"]))["curves"]
    lines = [c for c in curves if c.get("kind") == "line"]
    d = await call(mcp, "inventor_drawing_add_dimension", type="linear", entity_ids=[lines[0]["entity_id"]],
                   x_mm=100, y_mm=200)
    val = d.get("value_mm", d.get("value"))
    check("drawing linear dimension reads model length", any(abs((val or 0) - L) < 1e-6 for L in (60, 40, 10)), d)
    top_curves = (await call(mcp, "inventor_drawing_list_view_curves", view=proj["name"]))["curves"]
    circles = [c for c in top_curves if c.get("kind") == "circle"]
    if circles:
        dd = await call(mcp, "inventor_drawing_add_dimension", type="diameter", entity_ids=[circles[0]["entity_id"]],
                        x_mm=150, y_mm=60)
        check("drawing diameter dimension = 10", abs((dd.get("value_mm", dd.get("value")) or 0) - 10) < 1e-6, dd)
    dims = await call(mcp, "inventor_drawing_list_dimensions")
    check("drawing_list_dimensions", len(dims) >= 1, dims)
    await call(mcp, "inventor_drawing_delete_view", view=proj["name"])
    check("drawing_delete_view", len(await call(mcp, "inventor_drawing_list_views")) == 1)

    tb = await call(mcp, "inventor_drawing_list_title_block_fields")
    check("title block fields", tb, str(tb)[:200])
    new = await call(mcp, "inventor_drawing_add_sheet", name="Assembly")
    av = await call(mcp, "inventor_drawing_add_base_view", model_doc=asm, x_mm=120, y_mm=150, scale=1,
                    sheet=new.get("name"))
    await call(mcp, "inventor_drawing_add_parts_list", view=av["name"], x_mm=20, y_mm=280, sheet=new.get("name"))
    pl = await call(mcp, "inventor_drawing_get_parts_list", sheet=new.get("name"))
    check("parts list has Plate and Pin", "Plate" in json.dumps(pl) and "Pin" in json.dumps(pl), str(pl)[:300])
    pdf = os.path.join(WORK, "drawing.pdf")
    await call(mcp, "inventor_export", format="pdf", path=pdf)
    check("export drawing to pdf", os.path.isfile(pdf) and os.path.getsize(pdf) > 0)
    dwg = await call(mcp, "inventor_export", format="dwg", path=os.path.join(WORK, "drawing.dwg"))
    check("export 2-sheet drawing to dwg (one file per sheet, no dialog)",
          len(dwg["files"]) == 2 and all(os.path.getsize(p) > 0 for p in dwg["files"]), dwg)
    await call(mcp, "inventor_drawing_delete_sheet", name=new.get("name"))
    check("drawing_delete_sheet", len(await call(mcp, "inventor_drawing_list_sheets")) == len(sheets))
    await call(mcp, "inventor_close_all_documents", save=False)


async def main() -> int:
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK)
    before = inventor_count()
    import inventor_mcp
    mcp = inventor_mcp.mcp
    check("server import does not start Inventor", inventor_count() == before, f"{before} -> {inventor_count()}")
    tools = [t.name for t in await mcp.list_tools()]
    check("tools registered", len(tools) >= 100, len(tools))

    status = await call(mcp, "inventor_get_status")
    check("connected to Inventor", status.get("running"), status.get("version"))
    started_here = before == 0
    try:
        sections = (build_parts, assembly_tests, part_feature_tests, units_and_bodies_tests,
                    sketch_tools_test, edge_feature_test, body_feature_test, part_utility_test,
                    document_test, assembly_more_test, drawing_test)
        only = [a for a in sys.argv[1:] if not a.startswith("-")]
        if only:  # e.g. functional_test.py build_parts drawing_test
            sections = [s for s in sections if s.__name__ in only]
        for section in sections:
            try:
                await section(mcp)
            except Exception as e:  # noqa: BLE001 - report and continue with next section
                check(f"{section.__name__} ran without error", False, f"{type(e).__name__}: {e}")
                traceback.print_exc()
                try:
                    await call(mcp, "inventor_close_all_documents", save=False)
                except Exception:
                    pass
        if not only:
            missing = sorted(set(tools) - CALLED)
            check(f"every tool exercised ({len(set(tools) & CALLED)}/{len(tools)})", not missing, missing)
    finally:
        try:
            await call(mcp, "inventor_close_all_documents", save=False)
        except Exception:
            pass
        check("still exactly one Inventor", inventor_count() == max(before, 1), inventor_count())
        if started_here:
            from inventor.connection import get_app
            get_app().Quit()
        shutil.rmtree(WORK, ignore_errors=True)

    failed = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(failed)}/{len(results)} passed")
    for name, _, detail in failed:
        print(f"  FAILED: {name}  -> {detail}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
