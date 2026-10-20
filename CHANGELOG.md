# Changelog

## 2026-10-20

### Added
- `inventor_assembly_get_occurrence` — one occurrence's position, full rotation
  matrix, and grounded/suppressed/visible state. *(Andrzej Kalinowski)*
- `preserve_position` option for `inventor_assembly_add_joint` (default on): keeps
  already-placed parts where they are by trying other flip combinations, or a
  rotational joint locked at the restoring angle. *(Andrzej Kalinowski)*
- `inventor_assembly_auto_joint_fasteners` — matches screws/pins/washers to their
  seats by coaxial circular edges and joins them with rigid joints; dry run by
  default. *(Andrzej Kalinowski)*

### Changed
- The Inventor type library version is read from the registry instead of being
  hard-coded (approach from Andrzej Kalinowski's PR #2).

## 2026-10-17

### Fixed
- **Two Inventor instances opened when Claude Desktop started.** Inventor's COM
  server is single-use, so every `Dispatch("Inventor.Application")` launches a new
  `Inventor.exe`, and the server connected at import time. The server no longer
  starts Inventor on startup; on the first tool call it attaches to a running
  Inventor through the Running Object Table and launches a single instance only if
  none exists, under a machine-wide lock (`inventor/connection.py`).
- Polygons now form a closed profile (`inventor_sketch_add_polygon`).
- `inventor_finish_sketch` reports the real number of closed regions
  (`closed_regions`) and warns when it is 0; the temporary probe profile is deleted.
- Face/edge listing no longer fails on threaded screws or imported bodies whose
  geometry cannot be evaluated; such entities are reported as `"unknown"`.
  *(Andrzej Kalinowski)*

### Added
- **Assembly joints** — `inventor_assembly_add_joint`, `_list_joints`,
  `_delete_joint`, `_drive_joint` (rigid, rotational, slider, cylindrical, planar,
  ball; angular/linear limits). *(Andrzej Kalinowski)*
- Occurrence face/edge `entity_id`s are re-resolved after a server restart for
  constraints and joints; cylindrical faces report `base_point_mm`.
  *(Andrzej Kalinowski)*
- 3D sketches (`inventor_create_sketch3d`, `inventor_sketch3d_add_line`,
  `_add_spline`, `_list_entities`); 3D sketch curves usable as sweep paths.
- Work planes through three points and normal to a curve.
- Sheet metal: convert, face, flange, flat pattern.
- Standard threads (`inventor_thread`, optional `start_edge_ref`).
- `inventor_check_model` health report and `inventor_close_all_documents`.
- `inventor_run_script` for running Python against the Inventor COM API
  (disable with `INVENTOR_MCP_DISABLE_SCRIPT=1`).
- Optional JSONL call tracing via `INVENTOR_MCP_TRACE`.
