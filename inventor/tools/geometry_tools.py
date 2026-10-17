"""Govde/yuz/kenar sorgulama araclari."""

import pywintypes
from win32com.client import CastTo, constants

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import get_document, get_part_component_def, inventor_tool
from inventor.units import cm2_to_mm2, cm3_to_mm3, cm_to_mm

SURFACE_TYPE_NAMES = {
    constants.kPlaneSurface: "plane",
    constants.kCylinderSurface: "cylinder",
    constants.kEllipticalCylinderSurface: "elliptical_cylinder",
    constants.kConeSurface: "cone",
    constants.kEllipticalConeSurface: "elliptical_cone",
    constants.kTorusSurface: "torus",
    constants.kSphereSurface: "sphere",
    constants.kBSplineSurface: "bspline",
    constants.kUnknownSurface: "unknown",
}

CURVE_TYPE_NAMES = {
    constants.kLineCurve: "line",
    constants.kLineSegmentCurve: "line_segment",
    constants.kCircleCurve: "circle",
    constants.kCircularArcCurve: "circular_arc",
    constants.kEllipseFullCurve: "ellipse",
    constants.kEllipticalArcCurve: "elliptical_arc",
    constants.kBSplineCurve: "bspline",
    constants.kPolylineCurve: "polyline",
    constants.kUnknownCurve: "unknown",
}


def _point_mm(p) -> list:
    return [cm_to_mm(p.X), cm_to_mm(p.Y), cm_to_mm(p.Z)]


def _vector_xyz(v) -> list:
    return [v.X, v.Y, v.Z]


def face_summary(face) -> dict:
    """Bir Face nesnesi icin yuzey turu, alan (mm^2) ve turune ozgu geometriyi dondurur.

    Evaluator/Geometry'si okunamayan yuzeyler icin hata firlatmak yerine
    surface_type 'unknown' ve area_mm2 None dondurur.
    """
    try:
        surface_type = face.SurfaceType
        evaluator = face.Evaluator
        geo = face.Geometry
    except pywintypes.com_error:
        return {"surface_type": "unknown", "area_mm2": None}
    if evaluator is None or geo is None:
        return {"surface_type": "unknown", "area_mm2": None}
    entry = {
        "surface_type": SURFACE_TYPE_NAMES.get(surface_type, str(surface_type)),
        "area_mm2": cm2_to_mm2(evaluator.Area),
    }
    if surface_type == constants.kPlaneSurface:
        pl = CastTo(geo, "Plane")
        entry["normal"] = _vector_xyz(pl.Normal)
        entry["root_point_mm"] = _point_mm(pl.RootPoint)
    elif surface_type == constants.kCylinderSurface:
        cy = CastTo(geo, "Cylinder")
        entry["radius_mm"] = cm_to_mm(cy.Radius)
        entry["axis"] = _vector_xyz(cy.AxisVector)
        entry["base_point_mm"] = _point_mm(cy.BasePoint)
    return entry


def edge_summary(edge) -> dict:
    """Bir Edge nesnesi icin egri turu, uzunluk (mm) ve turune ozgu geometriyi dondurur.

    Bazi kenarlarda (dis/helis iceren vidalar, bazi ice aktarilmis SolidWorks
    govdeleri) Evaluator veya Geometry None donebiliyor; tum listeyi bozmak
    yerine bu kenarlar curve_type 'unknown' ve length_mm None ile raporlanir.
    """
    try:
        curve_type = edge.GeometryType
        evaluator = edge.Evaluator
        geo = edge.Geometry
        if evaluator is None or geo is None:
            return {"curve_type": "unknown", "length_mm": None}
        min_param, max_param = evaluator.GetParamExtents()
        length_cm = evaluator.GetLengthAtParam(min_param, max_param)
    except pywintypes.com_error:
        return {"curve_type": "unknown", "length_mm": None}
    entry = {
        "curve_type": CURVE_TYPE_NAMES.get(curve_type, str(curve_type)),
        "length_mm": cm_to_mm(length_cm),
    }
    # geo zaten win32com tarafindan dogru somut tipte (LineSegment3d/Circle3d/Arc3d
    # vb.) dondurulur; CastTo cagirmak farkli Inventor surumlerinde arayuz adi
    # uyusmazligi yuzunden hata verebiliyor (bkz. slot insasinda da karsilasilan
    # ayni sinif sorun) - dogrudan geo uzerinden erisiyoruz.
    if curve_type == constants.kLineSegmentCurve:
        entry["start_mm"] = _point_mm(geo.StartPoint)
        entry["end_mm"] = _point_mm(geo.EndPoint)
    elif curve_type == constants.kCircleCurve:
        entry["center_mm"] = _point_mm(geo.Center)
        entry["radius_mm"] = cm_to_mm(geo.Radius)
    elif curve_type == constants.kCircularArcCurve:
        entry["center_mm"] = _point_mm(geo.Center)
        entry["radius_mm"] = cm_to_mm(geo.Radius)
    return entry


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_list_bodies(doc: str | None = None) -> list[dict]:
        """Part dokumanindaki tum katı govdeleri (SurfaceBodies) listeler: ad, yuz/kenar
        sayisi, hacim (mm^3), yuzey alani (mm^2) ve sinir kutusu (mm)."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)

        result = []
        for i in range(1, comp_def.SurfaceBodies.Count + 1):
            body = comp_def.SurfaceBodies.Item(i)
            rb = body.RangeBox
            mp = body.MassProperties
            result.append(
                {
                    "index": i,
                    "name": body.Name,
                    "face_count": body.Faces.Count,
                    "edge_count": body.Edges.Count,
                    "volume_mm3": cm3_to_mm3(mp.Volume),
                    "area_mm2": cm2_to_mm2(mp.Area),
                    "bounding_box_mm": {
                        "min": _point_mm(rb.MinPoint),
                        "max": _point_mm(rb.MaxPoint),
                    },
                }
            )
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_list_faces(body_index: int = 1, doc: str | None = None) -> list[dict]:
        """Bir govdenin tum yuzeylerini listeler: entity_id, yuzey turu, alan (mm^2) ve
        turune ozgu geometri (duzlem icin normal+kok nokta, silindir icin yaricap+eksen+
        eksen uzerindeki base_point_mm)."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        body = comp_def.SurfaceBodies.Item(body_index)

        registry.clear_sketch(document.DisplayName, "_faces")

        result = []
        for i in range(1, body.Faces.Count + 1):
            face = body.Faces.Item(i)
            entity_id = registry.register(document.DisplayName, "_faces", "face", face)
            entry = {"entity_id": entity_id, "index": i}
            entry.update(face_summary(face))
            result.append(entry)
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_list_edges(body_index: int = 1, doc: str | None = None) -> list[dict]:
        """Bir govdenin tum kenarlarini listeler: entity_id, egri turu, uzunluk (mm) ve
        turune ozgu geometri (dogru icin uc noktalar, daire/yay icin merkez+yaricap)."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        body = comp_def.SurfaceBodies.Item(body_index)

        registry.clear_sketch(document.DisplayName, "_edges")

        result = []
        for i in range(1, body.Edges.Count + 1):
            edge = body.Edges.Item(i)
            entity_id = registry.register(document.DisplayName, "_edges", "edge", edge)
            entry = {"entity_id": entity_id, "index": i}
            entry.update(edge_summary(edge))
            result.append(entry)
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_get_bounding_box(doc: str | None = None) -> dict:
        """Tum govdelerin birlesik sinir kutusunu (mm) dondurur: min, max ve boyutlar (dx,dy,dz)."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)

        if comp_def.SurfaceBodies.Count == 0:
            raise LookupError("Dokumanda govde (SurfaceBody) yok.")

        min_pt, max_pt = None, None
        for body in comp_def.SurfaceBodies:
            rb = body.RangeBox
            b_min = (rb.MinPoint.X, rb.MinPoint.Y, rb.MinPoint.Z)
            b_max = (rb.MaxPoint.X, rb.MaxPoint.Y, rb.MaxPoint.Z)
            if min_pt is None:
                min_pt, max_pt = b_min, b_max
            else:
                min_pt = tuple(min(a, b) for a, b in zip(min_pt, b_min))
                max_pt = tuple(max(a, b) for a, b in zip(max_pt, b_max))

        min_mm = [cm_to_mm(v) for v in min_pt]
        max_mm = [cm_to_mm(v) for v in max_pt]
        size_mm = [b - a for a, b in zip(min_mm, max_mm)]
        return {"min_mm": min_mm, "max_mm": max_mm, "size_mm": size_mm}
