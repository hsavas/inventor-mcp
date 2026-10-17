"""Montaj (assembly) araclari: parca yerlestirme, kisitlamalar, pattern, BOM."""

import re

import pywintypes
from win32com.client import CastTo, constants

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import get_assembly_component_def, get_document, inventor_tool
from inventor.tools.geometry_tools import edge_summary, face_summary
from inventor.units import cm_to_mm, deg_to_rad, mm_to_cm, rad_to_deg

CONSTRAINT_TYPE_NAMES = {
    constants.kMateConstraintObject: "mate",
    constants.kFlushConstraintObject: "flush",
    constants.kAngleConstraintObject: "angle",
    constants.kInsertConstraintObject: "insert",
    constants.kTangentConstraintObject: "tangent",
}

# Constraint turune gore deger (offset/aci) tasiyan ozelligi okumak icin CastTo hedefi.
_CONSTRAINT_VALUE_CAST = {
    constants.kMateConstraintObject: ("MateConstraint", "Offset", "offset_mm", cm_to_mm),
    constants.kFlushConstraintObject: ("FlushConstraint", "Offset", "offset_mm", cm_to_mm),
    constants.kAngleConstraintObject: ("AngleConstraint", "Angle", "angle_deg", rad_to_deg),
    constants.kInsertConstraintObject: ("InsertConstraint", "Distance", "offset_mm", cm_to_mm),
    constants.kTangentConstraintObject: ("TangentConstraint", "Offset", "offset_mm", cm_to_mm),
}


def _resolve_occurrence(comp_def, ref: str):
    try:
        return comp_def.Occurrences.ItemByName(ref)
    except pywintypes.com_error:
        raise LookupError(f"'{ref}' adli occurrence bulunamadi.")


# inventor_assembly_list_occurrence_faces/_edges'in urettigi id bicimi:
# "<dokuman>:_asm_<occurrence>_faces:face<n>" (n, 0'dan baslayan yuz/kenar sirasi).
_ASM_ENTITY_ID = re.compile(r"^(?P<doc>.+?):_asm_(?P<occ>.+)_(?P<group>faces|edges):(?:face|edge)(?P<idx>\d+)$")


def resolve_entity(document, comp_def, entity_id: str):
    """Bir occurrence yuz/kenar entity_id'sini COM nesnesine cevirir.

    Id bu oturumda listelenmemisse (orn. sunucu yeniden basladiysa) id'deki
    occurrence adi ve sira numarasi ile yuz/kenar dogrudan govdeden alinir.
    """
    try:
        return registry.resolve(entity_id)
    except LookupError:
        m = _ASM_ENTITY_ID.match(entity_id)
        if not m or m["doc"] != document.DisplayName:
            raise LookupError(
                f"Bilinmeyen entity_id: '{entity_id}'. Once 'inventor_assembly_list_occurrence_faces' "
                f"veya '_edges' ile '{document.DisplayName}' dokumaninda listeleyin."
            )
    body = _resolve_occurrence(comp_def, m["occ"]).SurfaceBodies.Item(1)
    items = body.Faces if m["group"] == "faces" else body.Edges
    idx = int(m["idx"]) + 1
    if idx > items.Count:
        raise LookupError(f"Bilinmeyen entity_id: '{entity_id}' (occurrence'ta {items.Count} oge var).")
    return items.Item(idx)


def _object_collection(app, items):
    coll = app.TransientObjects.CreateObjectCollection()
    for item in items:
        coll.Add(item)
    return coll


def _occurrence_summary(occ) -> dict:
    t = occ.Transformation.Translation
    return {
        "name": occ.Name,
        "file": occ.ReferencedFileDescriptor.FullFileName,
        "grounded": bool(occ.Grounded),
        "suppressed": bool(occ.Suppressed),
        "visible": bool(occ.Visible),
        "position_mm": [cm_to_mm(t.X), cm_to_mm(t.Y), cm_to_mm(t.Z)],
    }


def _constraint_summary(c) -> dict:
    entry = {
        "name": c.Name,
        "type": CONSTRAINT_TYPE_NAMES.get(c.Type, str(c.Type)),
        "suppressed": bool(c.Suppressed),
        "occurrence_one": c.OccurrenceOne.Name if c.OccurrenceOne else None,
        "occurrence_two": c.OccurrenceTwo.Name if c.OccurrenceTwo else None,
    }
    cast_info = _CONSTRAINT_VALUE_CAST.get(c.Type)
    if cast_info:
        iface_name, prop_name, key, convert = cast_info
        specific = CastTo(c, iface_name)
        entry[key] = convert(getattr(specific, prop_name).Value)
    return entry


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_place_component(
        file_path: str,
        x_mm: float = 0.0,
        y_mm: float = 0.0,
        z_mm: float = 0.0,
        doc: str | None = None,
    ) -> dict:
        """Bir parca/alt-montaj dosyasini (.ipt/.iam) montaja yeni bir occurrence olarak yerlestirir.

        x_mm/y_mm/z_mm: montaj orijinine gore yerlestirme konumu (mm).
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        tg = app.TransientGeometry
        matrix = tg.CreateMatrix()
        if x_mm or y_mm or z_mm:
            matrix.SetTranslation(tg.CreateVector(mm_to_cm(x_mm), mm_to_cm(y_mm), mm_to_cm(z_mm)))
        occ = comp_def.Occurrences.Add(file_path, matrix)
        return _occurrence_summary(occ)

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_list_occurrences(doc: str | None = None) -> list[dict]:
        """Montajdaki en ust seviye occurrence'lari (yerlestirilmis parca/alt-montajlar) listeler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        return [_occurrence_summary(comp_def.Occurrences.Item(i)) for i in range(1, comp_def.Occurrences.Count + 1)]

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_get_occurrence(name: str, doc: str | None = None) -> dict:
        """Tek bir occurrence'in konumunu (mm), tam rotasyonunu, grounded/suppressed durumunu dondurur.

        rotation: occ.Transformation'in 3x3 rotasyon kismi, satir satir; sutunlar
        occurrence'in X/Y/Z eksenlerinin montaj uzayindaki yonleridir.
        Tum listeyi cekmeden tek parcayi kontrol etmek icin `inventor_assembly_list_occurrences`
        yerine kullanilabilir.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        occ = _resolve_occurrence(comp_def, name)
        entry = _occurrence_summary(occ)
        m = occ.Transformation
        entry["rotation"] = [[m.Cell(r, c) for c in range(1, 4)] for r in range(1, 4)]
        return entry

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_set_grounded(occurrence: str, grounded: bool = True, doc: str | None = None) -> dict:
        """Bir occurrence'in sabitlenmis (grounded) durumunu ayarlar."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        occ = _resolve_occurrence(comp_def, occurrence)
        occ.Grounded = grounded
        return _occurrence_summary(occ)

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_set_position(
        occurrence: str,
        x_mm: float,
        y_mm: float,
        z_mm: float,
        doc: str | None = None,
    ) -> dict:
        """Bir occurrence'in konumunu (montaj orijinine gore, mm) degistirir; rotasyonu korur."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        tg = app.TransientGeometry
        occ = _resolve_occurrence(comp_def, occurrence)
        matrix = occ.Transformation
        matrix.SetTranslation(tg.CreateVector(mm_to_cm(x_mm), mm_to_cm(y_mm), mm_to_cm(z_mm)))
        occ.Transformation = matrix
        return _occurrence_summary(occ)

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_delete_occurrence(occurrence: str, doc: str | None = None) -> dict:
        """Bir occurrence'i (ve varsa onunla iliskili kisitlamalari) montajdan siler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        occ = _resolve_occurrence(comp_def, occurrence)
        name = occ.Name
        occ.Delete()
        return {"deleted": name}

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_list_occurrence_faces(occurrence: str, doc: str | None = None) -> list[dict]:
        """Bir occurrence'in (montaj uzayindaki) govdesinin tum yuzeylerini listeler.

        Donen entity_id'ler `inventor_assembly_add_constraint` ve
        `inventor_assembly_add_joint` icin kullanilabilir. Silindirik yuzeyler icin
        eksen uzerindeki base_point_mm, ayni yaricapli delikleri ayirt etmeye yarar.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        occ = _resolve_occurrence(comp_def, occurrence)
        body = occ.SurfaceBodies.Item(1)

        group = f"_asm_{occurrence}_faces"
        registry.clear_sketch(document.DisplayName, group)

        result = []
        for i in range(1, body.Faces.Count + 1):
            face = body.Faces.Item(i)
            entity_id = registry.register(document.DisplayName, group, "face", face)
            entry = {"entity_id": entity_id, "index": i}
            entry.update(face_summary(face))
            result.append(entry)
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_list_occurrence_edges(occurrence: str, doc: str | None = None) -> list[dict]:
        """Bir occurrence'in (montaj uzayindaki) govdesinin tum kenarlarini listeler.

        Donen entity_id'ler pattern araclarinda eksen/kenar referansi olarak ve
        `inventor_assembly_add_joint` / `inventor_assembly_add_constraint` icin
        kullanilabilir.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        occ = _resolve_occurrence(comp_def, occurrence)
        body = occ.SurfaceBodies.Item(1)

        group = f"_asm_{occurrence}_edges"
        registry.clear_sketch(document.DisplayName, group)

        result = []
        for i in range(1, body.Edges.Count + 1):
            edge = body.Edges.Item(i)
            entity_id = registry.register(document.DisplayName, group, "edge", edge)
            entry = {"entity_id": entity_id, "index": i}
            entry.update(edge_summary(edge))
            result.append(entry)
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_add_constraint(
        constraint_type: str,
        entity_one: str,
        entity_two: str,
        offset_mm: float = 0.0,
        angle_deg: float = 0.0,
        axes_opposed: bool = True,
        inside_tangency: bool = True,
        doc: str | None = None,
    ) -> dict:
        """Iki yuz/kenar arasina montaj kisitlamasi ekler.

        constraint_type: 'mate', 'flush', 'angle', 'insert', 'tangent'.
        entity_one/entity_two: `inventor_assembly_list_occurrence_faces` (veya edges)
        ile alinan entity_id'ler. Bu oturumda listelenmemis id'ler de (orn. sunucu
        yeniden basladiktan sonra) occurrence govdesinden yeniden cozulur.
        offset_mm: mate/flush/insert/tangent icin ofset (mm).
        angle_deg: angle kisitlamasi icin aci (derece).
        axes_opposed: insert kisitlamasi icin eksenlerin zit yonlu olup olmadigi.
        inside_tangency: tangent kisitlamasi icin ic teget secimi.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        ent1 = resolve_entity(document, comp_def, entity_one)
        ent2 = resolve_entity(document, comp_def, entity_two)
        constraints = comp_def.Constraints

        kind = constraint_type.lower()
        if kind == "mate":
            c = constraints.AddMateConstraint(ent1, ent2, mm_to_cm(offset_mm))
        elif kind == "flush":
            c = constraints.AddFlushConstraint(ent1, ent2, mm_to_cm(offset_mm))
        elif kind == "angle":
            c = constraints.AddAngleConstraint(ent1, ent2, deg_to_rad(angle_deg))
        elif kind == "insert":
            c = constraints.AddInsertConstraint(ent1, ent2, axes_opposed, mm_to_cm(offset_mm))
        elif kind == "tangent":
            c = constraints.AddTangentConstraint(ent1, ent2, inside_tangency, mm_to_cm(offset_mm))
        else:
            raise LookupError(
                f"Gecersiz constraint_type: '{constraint_type}'. Gecerli degerler: mate, flush, angle, insert, tangent."
            )
        return _constraint_summary(c)

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_list_constraints(doc: str | None = None) -> list[dict]:
        """Montajdaki tum kisitlamalari listeler (ad, tur, occurrence'lar, offset/aci)."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        constraints = comp_def.Constraints
        return [_constraint_summary(constraints.Item(i)) for i in range(1, constraints.Count + 1)]

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_delete_constraint(name: str, doc: str | None = None) -> dict:
        """Adi verilen montaj kisitlamasini siler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        try:
            c = comp_def.Constraints.Item(name)
        except pywintypes.com_error:
            raise LookupError(f"'{name}' adli kisitlama bulunamadi.")
        c.Delete()
        return {"deleted": name}

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_get_bom(doc: str | None = None) -> list[dict]:
        """Montajin malzeme listesini (BOM, yapilandirilmis gorunum) satir satir dondurur."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        structured = comp_def.BOM.BOMViews.Item(1)
        rows = structured.BOMRows
        result = []
        for i in range(1, rows.Count + 1):
            row = rows.Item(i)
            cdefs = row.ComponentDefinitions
            result.append(
                {
                    "item_number": row.ItemNumber,
                    "quantity": row.ItemQuantity,
                    "document": cdefs.Item(1).Document.DisplayName,
                    "occurrences": [o.Name for o in row.ComponentOccurrences],
                }
            )
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_pattern_rectangular(
        occurrence_refs: list[str],
        column_entity_ref: str,
        column_offset_mm: float,
        column_count: int,
        row_entity_ref: str | None = None,
        row_offset_mm: float | None = None,
        row_count: int | None = None,
        doc: str | None = None,
    ) -> dict:
        """Bir veya daha fazla occurrence'i dikdortgen duzende cogaltir.

        column_entity_ref/row_entity_ref: `inventor_assembly_list_occurrence_edges` ile
        alinan dogrusal kenar entity_id'leri (yon belirler).
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        occs = [_resolve_occurrence(comp_def, ref) for ref in occurrence_refs]
        parents = _object_collection(app, occs)
        column_entity = registry.resolve(column_entity_ref)

        if row_entity_ref:
            row_entity = registry.resolve(row_entity_ref)
            pattern = comp_def.OccurrencePatterns.AddRectangularPattern(
                parents,
                column_entity,
                True,
                mm_to_cm(column_offset_mm),
                column_count,
                row_entity,
                True,
                mm_to_cm(row_offset_mm),
                row_count,
            )
        else:
            pattern = comp_def.OccurrencePatterns.AddRectangularPattern(
                parents, column_entity, True, mm_to_cm(column_offset_mm), column_count
            )
        return {"name": pattern.Name, "type": "rectangular_pattern"}

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_pattern_circular(
        occurrence_refs: list[str],
        axis_entity_ref: str,
        count: int,
        angle_deg: float = 360.0,
        doc: str | None = None,
    ) -> dict:
        """Bir veya daha fazla occurrence'i bir eksen etrafinda dairesel duzende cogaltir.

        axis_entity_ref: `inventor_assembly_list_occurrence_edges` ile alinan dogrusal
        bir kenarin entity_id'si (donus ekseni olarak kullanilir).
        angle_deg: pattern elemanlari arasindaki toplam aci (360 ise esit araliklarla tam tur).
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        occs = [_resolve_occurrence(comp_def, ref) for ref in occurrence_refs]
        parents = _object_collection(app, occs)
        axis_entity = registry.resolve(axis_entity_ref)

        pattern = comp_def.OccurrencePatterns.AddCircularPattern(
            parents, axis_entity, True, deg_to_rad(angle_deg), count
        )
        return {"name": pattern.Name, "type": "circular_pattern"}
