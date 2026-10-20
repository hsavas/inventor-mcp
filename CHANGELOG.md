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
- `units` argument for `inventor_new_document` (default `"metric"`): uses the
  metric template and forces mm/gram/degree. *(Sam-H101)*
- `affected_bodies` argument for `inventor_extrude` and `inventor_revolve`, so
  join/cut/intersect can target specific bodies in multi-body parts. *(Sam-H101)*
- `units` argument for `inventor_create_parameter`. *(Sam-H101)*
- Automated tests tracked in the repo: `tests/functional_test.py` (end-to-end,
  all 109 tools) and `tests/check_single_instance.py`; see "Testing" in the README.
- `preserve_position` result reports `flips_changed` when the requested flips
  were changed to keep parts in place.

### Fixed
- User parameters are created in the document's units ("46 mm" no longer shows
  as cm); parameter `value` is reported in the parameter's own units, with the
  raw database value in `value_db`. *(Sam-H101)*
- `inventor_list_bodies` works on Inventor versions without
  `SurfaceBody.MassProperties` (before 2023), using `Volume(0.0001)` and face
  areas as a fallback. *(Sam-H101; fallback precision corrected)*
- `kFinishParameter` and joint health constants are looked up safely, so a
  missing constant in an older type library no longer stops the server from
  starting.
- A failed position-preserving joint attempt no longer leaves a temporary
  joint behind in the assembly.
- **Feature tools no longer report success for broken features.** Inventor can
  create a feature (hole, mirror, pattern, ...) that it marks as failed and that
  changes nothing; all 16 feature-creating tools now check the result, undo a
  failed feature (keeping its sketches) and return a clear error.
- `inventor_hole` drilled in a fixed direction, so a hole sketched on a plane on
  top of the part always failed. New `direction` argument (default `"auto"`
  finds the direction into the material; `"negative"`/`"positive"` force one).
- `inventor_set_parameter` now rebuilds the model; before, the value changed but
  the geometry did not.
- `inventor_sketch_add_dimension` now re-solves the sketch, so the geometry
  actually takes the dimension value; `linear` also accepts a single line (its
  length).
- `inventor_sketch_add_rectangle_three_point` now creates a closed profile.
- `inventor_sketch_mirror` (and other sketch helpers) failed on circles with
  "Member not found": entities were cast to the line interface regardless of
  their real type.
- `inventor_set_iproperty` creates a custom (User Defined) property that does
  not exist yet.
- `inventor_set_appearance` accepts the internal names returned by
  `inventor_list_appearances`, and a custom colour works for an appearance that
  is already used in the document.
- `inventor_export`: never shows a dialog (Inventor's SilentOperation), so an
  export can no longer block the server; DWG export works (translator is loaded
  and ZIP packaging is turned off), and multi-sheet drawings return one file
  per sheet in `files`.
- Test suite now exercises all 109 tools (108 checks) and fails if any tool is
  not called.

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
