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


async def call(mcp, _tool, /, **args):
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
        for section in (build_parts, assembly_tests, part_feature_tests, units_and_bodies_tests):
            try:
                await section(mcp)
            except Exception as e:  # noqa: BLE001 - report and continue with next section
                check(f"{section.__name__} ran without error", False, f"{type(e).__name__}: {e}")
                traceback.print_exc()
                try:
                    await call(mcp, "inventor_close_all_documents", save=False)
                except Exception:
                    pass
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
