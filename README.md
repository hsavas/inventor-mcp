# Autodesk Inventor MCP

An [MCP](https://modelcontextprotocol.io) server that bridges Claude (or any MCP-compatible client) to **Autodesk Inventor** via COM automation. It exposes Inventor's part modeling, assembly, drawing, and document tools as callable MCP tools, so an AI assistant can sketch, extrude, fillet, place assembly components, generate drawing views, and more — directly inside a running Inventor session.

> **Platform note:** Autodesk Inventor's COM API is Windows-only. This server must run on Windows with Inventor installed, using a Python environment with `pywin32`.

## Features

- **Sketching** — lines, rectangles, circles, arcs, polygons, polylines, slots, ellipses, splines, fillets, construction geometry, mirroring, constraints, and dimensions.
- **3D part features** — extrude, revolve, sweep, loft, coil (helical threads/springs), hole, fillet, chamfer, shell, rectangular/circular pattern, mirror, combine (boolean union/subtract/intersect).
- **3D sketches** — free-orientation 3D lines and splines (usable as sweep paths), plus work planes through three points or normal to a curve.
- **Sheet metal** — convert a part to sheet metal, face, flange, and flat pattern with unfolded size.
- **Threads** — standard (cosmetic) ISO metric threads on cylindrical faces.
- **Model checks** — `inventor_check_model` health report (bodies, volume, size, failed features, open sketches); `inventor_finish_sketch` reports closed-region count.
- **Scripting** — `inventor_run_script` runs a Python snippet against the Inventor COM API in a single call (see the security note below).
- **Work geometry** — work planes, axes, and points.
- **Parameters** — list, read, set, and create model parameters.
- **Properties** — iProperties and mass properties (volume, mass, center of gravity, etc.).
- **Geometry inspection** — list bodies, faces, edges, and bounding boxes.
- **Assemblies** — place/ground/position components, inspect a single occurrence's position and rotation, add and inspect mate constraints, rectangular/circular component patterns, bill of materials.
- **Joints** — rigid, rotational, slider, cylindrical, planar, and ball assembly joints with optional angular/linear limits; joints keep already-placed parts where they are by default; list, delete, and drive joints to test motion; automatically joint screws/pins/washers to the parts they sit on.
- **Drawings** — sheets, base/projected views, dimensions, parts lists, title block fields.
- **Appearances** — list and apply visual styles/materials.
- **Document & app control** — new/open/save/close documents, active document switching, status checks.
- **Visualization** — zoom-to-fit and screenshot capture for visual verification.

Tool implementations live under `inventor/tools/`, organized by domain (`sketch_tools.py`, `feature_tools.py`, `assembly_tools.py`, `joint_tools.py`, `drawing_tools.py`, `geometry_tools.py`, `parameter_tools.py`, `property_tools.py`, `appearance_tools.py`, `export_tools.py`, `view_tools.py`, `app_tools.py`, `sketch3d_tools.py`, `extra_tools.py`, `script_tools.py`).

## Requirements

- Windows with **Autodesk Inventor** installed and licensed.
- Python 3.10+ (a virtual environment is recommended).
- Dependencies listed in `requirements.txt`, notably:
  - [`mcp`](https://pypi.org/project/mcp/) (FastMCP server framework)
  - [`pywin32`](https://pypi.org/project/pywin32/) (COM automation)

## Installation

```powershell
# from the repository root
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Running the server

The server communicates over stdio, as expected by MCP clients such as Claude Desktop:

```powershell
.venv\Scripts\python.exe inventor_mcp.py
```

Starting the server does **not** start Inventor. On the first tool call it attaches to an already running Inventor (including one you opened yourself); only if none is running does it launch a single instance, guarded by a machine-wide lock so that several server processes or threads never open duplicate copies.

### Environment variables

| Variable | Effect |
|---|---|
| `INVENTOR_MCP_DISABLE_SCRIPT=1` | Do not register `inventor_run_script`. |
| `INVENTOR_MCP_TRACE=<path.jsonl>` | Append every tool call (arguments, result, duration) to a JSONL file. Long string arguments and results are truncated. |
| `INVENTOR_MCP_SESSION=<tag>` | Optional session label written into each trace record. |

> **Security note:** `inventor_run_script` executes arbitrary Python with the server's permissions, i.e. full access to your machine, not just Inventor. Only connect this server to clients you trust, or set `INVENTOR_MCP_DISABLE_SCRIPT=1`.

### Using with Claude Desktop / Cowork

Add an entry to your MCP client's configuration pointing at the venv interpreter and `inventor_mcp.py`, for example:

```json
{
  "mcpServers": {
    "inventor": {
      "command": "C:\\path\\to\\repo\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\repo\\inventor_mcp.py"]
    }
  }
}
```

A ready-to-customize Claude Desktop Extension (`.mcpb`) manifest template is provided at `inventor-mcp-ext/manifest.example.json` — copy it to `inventor-mcp-ext/manifest.json`, replace the `<ABSOLUTE_PATH_TO_REPO>` placeholders with your local checkout path, then package or load it.

### Example: joints

Joints connect a face or edge on one occurrence to a face or edge on another. Get `entity_id`s from `inventor_assembly_list_occurrence_faces` / `inventor_assembly_list_occurrence_edges`. Inventor moves the occurrence of `entity_one` onto `entity_two`, so pass the moving part first and the grounded part second.

```text
inventor_assembly_list_occurrence_edges(occurrence="pin:1")
inventor_assembly_list_occurrence_edges(occurrence="plate:1")

inventor_assembly_add_joint(
    joint_type="rotational",
    entity_one="Assembly1:_asm_pin:1_edges:edge0",    # circular edge of the pin
    entity_two="Assembly1:_asm_plate:1_edges:edge4",  # circular edge of the hole
    angle_limits_deg=[0, 90],
)
inventor_assembly_drive_joint(name="Rotational:1", angle_deg=45)
```

`point_one`/`point_two` (`start`, `mid`, `end`, `center`) pick the joint origin: a point on an edge, a point on the axis of a cylindrical/conical face, or the center of a planar face (`center` only). Angular limits and `angle_deg` apply to rotational and cylindrical joints; linear limits and `linear_mm` apply to slider and cylindrical joints. A joint driven outside its limits reports `health: "invalid_limits"`.

By default (`preserve_position=True`) `inventor_assembly_add_joint` keeps already-placed parts where they are. If a rigid joint would move the part, it tries the other `flip_alignment`/`flip_origin` combinations. If only an in-plane rotation remains (common with circular edges), it creates a rotational joint locked at that angle instead, named `<name>_locked`. The result's `preserve_position` field reports the final joint type, the flips used and the residual error in mm/deg. Pass `preserve_position=False` to apply the requested flips as-is.

To joint fasteners in bulk, preview first, then apply:

```text
inventor_assembly_auto_joint_fasteners()                     # dry run: fastener, partner, edge ids, radius, seat
inventor_assembly_auto_joint_fasteners(skip=["pin:3"], dry_run=False)
```

Occurrences whose names match `name_filter` (regex, default `screw|pin|washer`) are matched to a coaxial, coplanar circular edge on another occurrence. A head-underside-on-seat contact is preferred. Fasteners that already have joints or constraints are skipped, and duplicates (same file, identical position) are reported.

## Architecture notes

- `inventor/connection.py` manages the connection to `Inventor.Application`, re-initializing COM per call to stay safe across the multi-threaded calls an MCP server can receive. Inventor's COM server is *single-use*: every `CreateObject`/`Dispatch("Inventor.Application")` starts a **new** Inventor process. The connection layer therefore always attaches through the Running Object Table first (`GetActiveObject`) and creates an instance only when none exists, under a named mutex. At import time it only loads the type library so `win32com.client.constants` is populated, without starting Inventor.
- `inventor/registry.py` tracks sketch entities and other COM object references by an opaque `entity_id` string, so tool calls can refer back to objects created in earlier calls without re-querying geometry.
- `inventor/helpers.py` and `inventor/units.py` provide shared geometry helpers and millimeter/centimeter unit conversions (Inventor's API is internally in centimeters).
- Several feature/sketch collections in Inventor's COM type library are wrapped with `win32com.client.dynamic.Dispatch` rather than the early-bound gencache proxy, working around a `gen_py` cache staleness issue observed with certain collections (e.g. `CoilFeatures`, `SweepFeatures`, `LoftFeatures`, `CombineFeatures`).
- Closed sketch profiles are built by chaining each new line from the previous line's own endpoint object, rather than via `AddCoincident` on independent points — this avoids a COM error Inventor raises when merging two line-endpoint `SketchPoint`s directly.

## Credits

- **[Andrzej Kalinowski](https://github.com/AndrzejKalinowski)** — assembly joint tools (`joint_tools.py`) with position-preserving joints and fastener auto-jointing, `inventor_assembly_get_occurrence`, robust face/edge listing for threaded and imported parts, re-resolving occurrence entity IDs after a server restart, and registry-based type library lookup ([PR #2](https://github.com/hsavas/inventor-mcp/pull/2)).

See [CHANGELOG.md](CHANGELOG.md) for the history of changes.

## License

Free for non-commercial use. Commercial use requires prior written permission from the author — see [LICENSE](LICENSE).
