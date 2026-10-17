"""2D sketch araclari: cizim, kisitlama, boyutlandirma."""

import math

import pywintypes
import win32com.client.dynamic
from win32com.client import CastTo, constants

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import (
    count_closed_regions,
    get_document,
    get_part_component_def,
    get_sketch,
    inventor_tool,
    point2d,
)
from inventor.units import cm_to_mm, deg_to_rad, mm_to_cm, rad_to_deg

PLANE_NAMES = {"XY": "XY Plane", "XZ": "XZ Plane", "YZ": "YZ Plane"}


def _resolve_plane(comp_def, plane: str):
    plane_name = PLANE_NAMES.get(plane.upper(), plane)
    try:
        return comp_def.WorkPlanes.Item(plane_name)
    except pywintypes.com_error:
        raise LookupError(f"'{plane}' adli bir duzlem/work plane bulunamadi.")


def _line_summary(line) -> dict:
    sp = line.StartSketchPoint.Geometry
    ep = line.EndSketchPoint.Geometry
    return {"start": [cm_to_mm(sp.X), cm_to_mm(sp.Y)], "end": [cm_to_mm(ep.X), cm_to_mm(ep.Y)]}


def _circle_summary(circle) -> dict:
    c = circle.CenterSketchPoint.Geometry
    return {"center": [cm_to_mm(c.X), cm_to_mm(c.Y)], "radius": cm_to_mm(circle.Radius)}


def _arc_summary(arc) -> dict:
    c = arc.CenterSketchPoint.Geometry
    return {
        "center": [cm_to_mm(c.X), cm_to_mm(c.Y)],
        "radius": cm_to_mm(arc.Radius),
        "start_deg": rad_to_deg(arc.StartAngle),
        "sweep_deg": rad_to_deg(arc.SweepAngle),
    }


def _point_summary(pt) -> dict:
    g = pt.Geometry
    return {"x": cm_to_mm(g.X), "y": cm_to_mm(g.Y)}


CONSTRAINT_ARITY = {
    "coincident": 2,
    "horizontal": 1,
    "vertical": 1,
    "parallel": 2,
    "perpendicular": 2,
    "tangent": 2,
    "equal": 2,
    "symmetric": 3,
}


def _add_constraint(geo_constraints, kind: str, entities: list):
    # win32com'un gen_py onbellegi GeometricConstraints icin de eksik/bozuk
    # olabilir (CoilFeatures'da gorulen ayni sorun). Dinamik (late-bound)
    # dispatch ile sariyoruz; bu, onbellekteki metot tanimini atlayip canli
    # COM nesnesindeki gercek metotlari isme gore cagirir.
    geo_constraints = win32com.client.dynamic.Dispatch(geo_constraints._oleobj_)
    if kind == "coincident":
        return geo_constraints.AddCoincident(entities[0], entities[1])
    if kind == "horizontal":
        return geo_constraints.AddHorizontal(entities[0])
    if kind == "vertical":
        return geo_constraints.AddVertical(entities[0])
    if kind == "parallel":
        return geo_constraints.AddParallel(entities[0], entities[1])
    if kind == "perpendicular":
        return geo_constraints.AddPerpendicular(entities[0], entities[1])
    if kind == "tangent":
        return geo_constraints.AddTangent(entities[0], entities[1])
    if kind == "equal":
        return geo_constraints.AddEqual(entities[0], entities[1])
    if kind == "symmetric":
        return geo_constraints.AddSymmetry(entities[0], entities[1], entities[2])
    raise LookupError(
        f"Gecersiz kisitlama turu: '{kind}'. Gecerli turler: {', '.join(CONSTRAINT_ARITY)}."
    )


def _entity_point(entity):
    """Bir entity'den bir SketchPoint dondurur (SketchLine ise baslangic noktasi)."""
    if hasattr(entity, "StartSketchPoint"):
        return entity.StartSketchPoint
    return entity


def _midpoint_text_point(app, p1, p2, offset_mm: float = 5.0):
    """Iki sketch noktasi arasindaki orta noktanin biraz ustunde bir TextPoint dondurur."""
    g1, g2 = p1.Geometry, p2.Geometry
    x_mm = cm_to_mm((g1.X + g2.X) / 2)
    y_mm = cm_to_mm((g1.Y + g2.Y) / 2) + offset_mm
    return point2d(app, x_mm, y_mm)


def _radial_text_point(app, entity, offset_mm: float = 5.0):
    """Bir daire/yayin merkezinden 45 derece yonde, yaricapin disinda bir TextPoint dondurur."""
    center = entity.CenterSketchPoint.Geometry
    radius_mm = cm_to_mm(entity.Radius)
    cx_mm, cy_mm = cm_to_mm(center.X), cm_to_mm(center.Y)
    x_mm = cx_mm + (radius_mm + offset_mm) * math.cos(math.radians(45))
    y_mm = cy_mm + (radius_mm + offset_mm) * math.sin(math.radians(45))
    return point2d(app, x_mm, y_mm)


def _add_slot_loop(app, document, sk, x1: float, y1: float, x2: float, y2: float, width: float) -> list[str]:
    """Iki merkez (mm) ve genislik (mm) verilen kapali bir 'stadium' slot profili
    olusturur: 2 paralel dogru + 2 yari-daire teget yay. Native Inventor 'Slot'
    komutunun urettigi geometriyle birebir aynidir (sadece dogru/yay ilkellerinden
    kurulu, ayrik bir API nesnesi degil)."""
    dx, dy = x2 - x1, y2 - y1
    center_span = math.hypot(dx, dy)
    if center_span < 1e-9:
        raise LookupError("Slot'un iki merkezi cakisik olamaz.")
    r = width / 2
    theta = math.atan2(dy, dx)

    p1a = (x1 + r * math.cos(theta + math.pi / 2), y1 + r * math.sin(theta + math.pi / 2))
    p1b = (x1 + r * math.cos(theta - math.pi / 2), y1 + r * math.sin(theta - math.pi / 2))
    p2a = (x2 + r * math.cos(theta + math.pi / 2), y2 + r * math.sin(theta + math.pi / 2))
    p2b = (x2 + r * math.cos(theta - math.pi / 2), y2 + r * math.sin(theta - math.pi / 2))

    # Once-merkez/iki-uc-nokta (AddByCenterStartEndPoint) kullanarak yaylari ONCE
    # olustur; sonra dogrulari yaylarin KENDI SketchPoint'lerinden insa et (yeni
    # Point2d degil). Boylece dogru/yay uc noktalari basindan itibaren AYNI
    # SketchPoint nesnesini paylasir ve ayrica bir AddCoincident'e gerek kalmaz
    # (bkz. Autodesk'in resmi MiniToolbarSlotSample'i: ayni desen kullanir).
    cap1 = sk.SketchArcs.AddByCenterStartEndPoint(
        point2d(app, x1, y1), point2d(app, *p1a), point2d(app, *p1b)
    )
    cap2 = sk.SketchArcs.AddByCenterStartEndPoint(
        point2d(app, x2, y2), point2d(app, *p2b), point2d(app, *p2a)
    )
    line_top = sk.SketchLines.AddByTwoPoints(cap1.StartSketchPoint, cap2.EndSketchPoint)
    line_bottom = sk.SketchLines.AddByTwoPoints(cap1.EndSketchPoint, cap2.StartSketchPoint)

    geo = sk.GeometricConstraints
    geo.AddEqualRadius(cap1, cap2)
    geo.AddTangent(line_top, cap1)
    geo.AddTangent(line_top, cap2)
    geo.AddTangent(line_bottom, cap1)
    geo.AddTangent(line_bottom, cap2)

    entities = [
        (line_top, "line"),
        (line_bottom, "line"),
        (cap1, "arc"),
        (cap2, "arc"),
    ]
    return [registry.register(document.DisplayName, sk.Name, kind, obj) for obj, kind in entities]


def _add_closed_polyline(app, sk, points: list[tuple[float, float]]) -> list:
    """N kose noktasi (mm) verilen kapali bir poligon olusturur.

    _add_slot_loop'taki ayni desen: her yeni cizgi, onceki cizginin KENDI
    SketchPoint'inden (yeni bir Point2d degil) baslar, boylece uc noktalar
    basindan itibaren ayni SketchPoint nesnesini paylasir ve ayrica
    AddCoincident'e gerek kalmaz (bu API'de iki cizgi-uc-noktasini birbirine
    AddCoincident ile sonradan birlestirmek guvenilir calismiyor).
    """
    n = len(points)
    if n < 3:
        raise LookupError("Kapali poligon en az 3 nokta gerektirir.")
    first_line = sk.SketchLines.AddByTwoPoints(point2d(app, *points[0]), point2d(app, *points[1]))
    lines = [first_line]
    prev_point = first_line.EndSketchPoint
    for i in range(2, n):
        line = sk.SketchLines.AddByTwoPoints(prev_point, point2d(app, *points[i]))
        lines.append(line)
        prev_point = line.EndSketchPoint
    closing_line = sk.SketchLines.AddByTwoPoints(prev_point, first_line.StartSketchPoint)
    lines.append(closing_line)
    return lines


def _ellipse_summary(ellipse) -> dict:
    c = ellipse.CenterSketchPoint.Geometry
    return {
        "center": [cm_to_mm(c.X), cm_to_mm(c.Y)],
        "major_radius": cm_to_mm(ellipse.MajorRadius),
        "minor_radius": cm_to_mm(ellipse.MinorRadius),
    }


def _reflect_point(px: float, py: float, x1: float, y1: float, x2: float, y2: float) -> tuple[float, float]:
    """(px,py) noktasini (x1,y1)-(x2,y2) dogrusuna gore yansitir (ayni birim, mm)."""
    dx, dy = x2 - x1, y2 - y1
    length_sq = dx * dx + dy * dy
    if length_sq < 1e-12:
        raise LookupError("Ayna dogrusunun iki noktasi cakisik olamaz.")
    vx, vy = px - x1, py - y1
    t = (vx * dx + vy * dy) / length_sq
    proj_x, proj_y = x1 + t * dx, y1 + t * dy
    return 2 * proj_x - px, 2 * proj_y - py


def _as_concrete_entity(entity):
    """AddAsTwoPointRectangle gibi metotlar generic 'SketchEntity' tipinde nesneler
    dondurebilir (StartSketchPoint gibi alt-tip ozelliklerini gostermez); somut
    SketchLine/SketchArc arayuzune CastTo ile gecis yapar."""
    for type_name in ("SketchLine", "SketchArc"):
        try:
            return CastTo(entity, type_name)
        except (pywintypes.com_error, TypeError):
            continue
    return entity


def _entity_endpoints(entity):
    """Bir line/arc entity'sinin uc SketchPoint'lerini dondurur."""
    entity = _as_concrete_entity(entity)
    if hasattr(entity, "StartSketchPoint"):
        return [entity.StartSketchPoint, entity.EndSketchPoint]
    return [entity]


def _entity_midpoint(app, entity):
    """Bir line/arc entity'sinin orta noktasinda bir Point2d dondurur.

    AddByFillet, fillet'in hangi taraftan uygulanacagini belirlemek icin her
    entity uzerinde (ona ait) bir 'near point' ister. Paylasilan kose noktasini
    vermek (tam kirpilma sinirinda, belirsiz) basarisiz olur; entity'nin orta
    noktasi (kesinlikle 'korunacak' tarafta) guvenilir sonuc verir.
    """
    entity = _as_concrete_entity(entity)
    if hasattr(entity, "StartSketchPoint"):
        g1 = entity.StartSketchPoint.Geometry
        g2 = entity.EndSketchPoint.Geometry
        return point2d(app, cm_to_mm((g1.X + g2.X) / 2), cm_to_mm((g1.Y + g2.Y) / 2))
    # Arc: orta aciya karsilik gelen nokta.
    c = entity.CenterSketchPoint.Geometry
    mid_angle = entity.StartAngle + entity.SweepAngle / 2
    x_cm = c.X + entity.Radius * math.cos(mid_angle)
    y_cm = c.Y + entity.Radius * math.sin(mid_angle)
    return point2d(app, cm_to_mm(x_cm), cm_to_mm(y_cm))


def _angle_text_point(app, line1, line2, offset_mm: float = 5.0):
    """Iki cizginin uc noktalarinin ortalamasinin biraz disinda bir TextPoint dondurur."""
    points = []
    for line in (line1, line2):
        for sketch_point in (line.StartSketchPoint, line.EndSketchPoint):
            g = sketch_point.Geometry
            points.append((g.X, g.Y))
    cx = sum(p[0] for p in points) / len(points)
    cy = sum(p[1] for p in points) / len(points)
    return point2d(app, cm_to_mm(cx) + offset_mm, cm_to_mm(cy) + offset_mm)


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_create_sketch(plane: str = "XY", name: str | None = None, doc: str | None = None) -> dict:
        """Aktif Part dokumaninda yeni bir 2D sketch (PlanarSketch) olusturur.

        plane: 'XY', 'XZ', 'YZ' varsayilan duzlemlerden biri ya da var olan bir
        work plane'in adi olabilir.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)

        plane_obj = _resolve_plane(comp_def, plane)
        sketch = comp_def.Sketches.Add(plane_obj)
        if name:
            sketch.Name = name
        return {"name": sketch.Name, "plane": plane_obj.Name}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_line(
        sketch: str, x1: float, y1: float, x2: float, y2: float, doc: str | None = None
    ) -> dict:
        """Sketch'e iki nokta arasinda bir cizgi ekler (koordinatlar mm cinsinden)."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        line = sk.SketchLines.AddByTwoPoints(point2d(app, x1, y1), point2d(app, x2, y2))
        entity_id = registry.register(document.DisplayName, sk.Name, "line", line)
        return {"entity_id": entity_id, **_line_summary(line)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_rectangle(
        sketch: str, x1: float, y1: float, x2: float, y2: float, doc: str | None = None
    ) -> dict:
        """Sketch'e kosegen kose noktalari (mm) verilen iki-nokta dikdortgen ekler (4 cizgi)."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        lines = sk.SketchLines.AddAsTwoPointRectangle(point2d(app, x1, y1), point2d(app, x2, y2))
        entity_ids = [registry.register(document.DisplayName, sk.Name, "line", line) for line in lines]
        return {"entity_ids": entity_ids, "corner1": [x1, y1], "corner2": [x2, y2]}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_circle(
        sketch: str, cx: float, cy: float, radius: float, doc: str | None = None
    ) -> dict:
        """Sketch'e merkez (mm) ve yaricap (mm) ile bir daire ekler."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        circle = sk.SketchCircles.AddByCenterRadius(point2d(app, cx, cy), mm_to_cm(radius))
        entity_id = registry.register(document.DisplayName, sk.Name, "circle", circle)
        return {"entity_id": entity_id, **_circle_summary(circle)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_arc(
        sketch: str,
        cx: float,
        cy: float,
        radius: float,
        start_deg: float,
        end_deg: float,
        doc: str | None = None,
    ) -> dict:
        """Sketch'e merkez (mm), yaricap (mm) ve baslangic/bitis acilariyla (derece, +X
        ekseninden itibaren saat yonu tersine) bir yay ekler."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        center = point2d(app, cx, cy)
        sweep_rad = deg_to_rad(end_deg - start_deg)
        arc = sk.SketchArcs.AddByCenterStartSweepAngle(center, mm_to_cm(radius), deg_to_rad(start_deg), sweep_rad)
        entity_id = registry.register(document.DisplayName, sk.Name, "arc", arc)
        return {"entity_id": entity_id, **_arc_summary(arc)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_point(sketch: str, x: float, y: float, doc: str | None = None) -> dict:
        """Sketch'e bir referans noktasi ekler (mm). Delik gibi feature'larin yerini
        belirlemek icin kullanilabilir."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        pt = sk.SketchPoints.Add(point2d(app, x, y))
        entity_id = registry.register(document.DisplayName, sk.Name, "point", pt)
        return {"entity_id": entity_id, **_point_summary(pt)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_rectangle_centered(
        sketch: str, cx: float, cy: float, width: float, height: float, doc: str | None = None
    ) -> dict:
        """Sketch'e merkez nokta (mm) ve genislik/yukseklik (mm) verilen bir dikdortgen ekler
        (UI'daki 'Rectangle Two Point Center' komutuna esdeger)."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        x1, y1 = cx - width / 2, cy - height / 2
        x2, y2 = cx + width / 2, cy + height / 2
        lines = sk.SketchLines.AddAsTwoPointRectangle(point2d(app, x1, y1), point2d(app, x2, y2))
        entity_ids = [registry.register(document.DisplayName, sk.Name, "line", line) for line in lines]
        return {"entity_ids": entity_ids, "center": [cx, cy], "width": width, "height": height}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_rectangle_three_point(
        sketch: str, x1: float, y1: float, x2: float, y2: float, width: float, doc: str | None = None
    ) -> dict:
        """Sketch'e ilk kenari (x1,y1)-(x2,y2) olan, bu kenara dik yonde 'width' (mm)
        genislikte (esit isaretli kenarin solunda) dondurulmus bir dikdortgen ekler
        (UI'daki 'Rectangle Three Point' komutuna esdeger)."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)

        dx, dy = x2 - x1, y2 - y1
        length = math.hypot(dx, dy)
        if length < 1e-9:
            raise LookupError("Ilk kenarin iki ucu cakisik olamaz.")
        nx, ny = -dy / length * width, dx / length * width

        p1 = (x1, y1)
        p2 = (x2, y2)
        p3 = (x2 + nx, y2 + ny)
        p4 = (x1 + nx, y1 + ny)
        corners = [p1, p2, p3, p4]

        lines = []
        for i in range(4):
            ax, ay = corners[i]
            bx, by = corners[(i + 1) % 4]
            line = sk.SketchLines.AddByTwoPoints(point2d(app, ax, ay), point2d(app, bx, by))
            lines.append(line)
        entity_ids = [registry.register(document.DisplayName, sk.Name, "line", line) for line in lines]
        return {"entity_ids": entity_ids, "corners": [list(p1), list(p2), list(p3), list(p4)]}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_polygon(
        sketch: str,
        cx: float,
        cy: float,
        radius: float,
        sides: int,
        angle_deg: float = 0,
        inscribed: bool = True,
        doc: str | None = None,
    ) -> dict:
        """Sketch'e merkez (mm), yaricap (mm) ve kenar sayisi verilen bir duzgun
        cokgen ekler (UI'daki 'Polygon' komutuna esdeger).

        inscribed=True: 'radius' merkezden koselere olan uzaklik (cevrel cember).
        inscribed=False: 'radius' merkezden kenar ortalarina olan uzaklik (ic teget cember),
        gercek kose yaricapi radius/cos(pi/sides) olarak hesaplanir.
        angle_deg: ilk kosenin +X ekseninden itibaren acisi (derece).
        """
        if sides < 3:
            raise LookupError("Bir cokgenin en az 3 kenari olmalidir.")
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)

        r = radius if inscribed else radius / math.cos(math.pi / sides)
        verts = []
        for i in range(sides):
            theta = math.radians(angle_deg) + i * 2 * math.pi / sides
            verts.append((cx + r * math.cos(theta), cy + r * math.sin(theta)))

        # Bagimsiz AddByTwoPoints cizgileri uc noktalarini paylasmaz ve kapali
        # profil olusturmaz (finish_sketch -> closed_regions: 0). Polyline ile
        # ayni zincirleme yontemi kullaniliyor.
        lines = _add_closed_polyline(app, sk, verts)
        entity_ids = [registry.register(document.DisplayName, sk.Name, "line", line) for line in lines]
        return {"entity_ids": entity_ids, "center": [cx, cy], "sides": sides, "vertex_radius": r}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_polyline(
        sketch: str, points: list[list[float]], doc: str | None = None
    ) -> dict:
        """Sketch'e N kose noktasi (mm, [[x,y], ...] en az 3 nokta) verilen KAPALI
        bir poligon (dogru zinciri) ekler; son nokta ilk noktaya otomatik baglanir.

        Elle cizilen dogrularda nokta paylasimi olmadigindan (her AddByTwoPoints
        cagrisi kendi bagimsiz uc noktalarini yaratir) ve bu API'de iki ayri
        dogrunun uc noktalarini sonradan AddCoincident ile birlestirmek guvenilir
        calismadigindan, bu arac her yeni dogruyu ONCEKI dogrunun KENDI
        EndSketchPoint'inden baslatarak insa eder (bkz. _add_slot_loop). Sonuc:
        tek kapali bir profil (inventor_finish_sketch -> profile_count: 1).
        Asimetrik/duzensiz sekiller (orn. V-disi profili) icin kullanin; duzgun
        cokgenler icin inventor_sketch_add_polygon yeterlidir.
        """
        if len(points) < 3:
            raise LookupError("Kapali poligon en az 3 nokta gerektirir.")
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        coords = [(p[0], p[1]) for p in points]
        lines = _add_closed_polyline(app, sk, coords)
        entity_ids = [registry.register(document.DisplayName, sk.Name, "line", line) for line in lines]
        return {"entity_ids": entity_ids, "points": points}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_slot_center_to_center(
        sketch: str, x1: float, y1: float, x2: float, y2: float, width: float, doc: str | None = None
    ) -> dict:
        """Sketch'e iki yuvarlatma merkezi (mm) ve genislik (mm) verilen, tek kapali
        profilli bir 'stadium' yuva (slot) ekler: 2 dogru + 2 yari-daire yay, teget
        kisitlamalariyla baglanir (UI'daki 'Slot Center to Center' komutuna esdeger).
        """
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        entity_ids = _add_slot_loop(app, document, sk, x1, y1, x2, y2, width)
        return {
            "entity_ids": entity_ids,
            "center1": [x1, y1],
            "center2": [x2, y2],
            "width": width,
        }

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_slot_overall(
        sketch: str,
        cx: float,
        cy: float,
        length: float,
        width: float,
        angle_deg: float = 0,
        doc: str | None = None,
    ) -> dict:
        """Sketch'e merkez nokta (mm), toplam uzunluk + genislik (mm) ve aci (derece)
        verilen bir 'stadium' yuva (slot) ekler (UI'daki 'Slot Overall' komutuna
        esdeger). 'length' yuvanin iki yari-daire merkezi arasindaki mesafe DEGIL,
        ucdan-uca toplam uzunluktur (merkez araligi = length - width)."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)

        center_span = length - width
        if center_span <= 0:
            raise LookupError("'length' degeri 'width'ten buyuk olmalidir.")
        half = center_span / 2
        theta = math.radians(angle_deg)
        ux, uy = math.cos(theta), math.sin(theta)
        x1, y1 = cx - ux * half, cy - uy * half
        x2, y2 = cx + ux * half, cy + uy * half

        entity_ids = _add_slot_loop(app, document, sk, x1, y1, x2, y2, width)
        return {
            "entity_ids": entity_ids,
            "center": [cx, cy],
            "length": length,
            "width": width,
            "angle_deg": angle_deg,
        }

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_fillet(
        sketch: str, entity_id_1: str, entity_id_2: str, radius: float, doc: str | None = None
    ) -> dict:
        """Iki sketch dogrusu (veya bir dogru/yay cifti) arasina gercek bir sketch
        fillet ekler: Inventor'in kendi 'SketchArcs.AddByFillet' API'sini cagirir, bu
        da UI'daki Sketch > Fillet komutuyla AYNI seydir (iki entity otomatik
        kirpilir/trim edilir ve aralarina teget bir yay yerlestirilir). Fillet'in
        hangi taraftan uygulanacagini belirten 'near point' ciftini, her entity'nin
        orta noktasini kullanarak otomatik hesaplar.

        entity_id_1/2: inventor_sketch_add_line/_rectangle veya
        inventor_sketch_list_entities'den alinan id'ler.
        """
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)

        e1 = _as_concrete_entity(registry.resolve(entity_id_1))
        e2 = _as_concrete_entity(registry.resolve(entity_id_2))
        near1 = _entity_midpoint(app, e1)
        near2 = _entity_midpoint(app, e2)
        arc = sk.SketchArcs.AddByFillet(e1, e2, mm_to_cm(radius), near1, near2)
        entity_id = registry.register(document.DisplayName, sk.Name, "arc", arc)
        return {"entity_id": entity_id, **_arc_summary(arc)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_list_entities(sketch: str, doc: str | None = None) -> dict:
        """Sketch'teki tum cizgi/daire/yay/nokta nesnelerini yeniden tarar, kayit defterini
        tazeler ve her biri icin entity_id ile ozet geometriyi (mm/derece) dondurur."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        registry.clear_sketch(document.DisplayName, sk.Name)

        entities = []
        for line in sk.SketchLines:
            entity_id = registry.register(document.DisplayName, sk.Name, "line", line)
            entities.append({"entity_id": entity_id, "kind": "line", **_line_summary(line)})
        for circle in sk.SketchCircles:
            entity_id = registry.register(document.DisplayName, sk.Name, "circle", circle)
            entities.append({"entity_id": entity_id, "kind": "circle", **_circle_summary(circle)})
        for arc in sk.SketchArcs:
            entity_id = registry.register(document.DisplayName, sk.Name, "arc", arc)
            entities.append({"entity_id": entity_id, "kind": "arc", **_arc_summary(arc)})
        for ellipse in sk.SketchEllipses:
            entity_id = registry.register(document.DisplayName, sk.Name, "ellipse", ellipse)
            entities.append({"entity_id": entity_id, "kind": "ellipse", **_ellipse_summary(ellipse)})
        for spline in sk.SketchSplines:
            entity_id = registry.register(document.DisplayName, sk.Name, "spline", spline)
            entities.append({"entity_id": entity_id, "kind": "spline"})
        for pt in sk.SketchPoints:
            entity_id = registry.register(document.DisplayName, sk.Name, "point", pt)
            entities.append({"entity_id": entity_id, "kind": "point", **_point_summary(pt)})

        return {"sketch": sk.Name, "entities": entities}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_constraint(sketch: str, type: str, entity_ids: list[str], doc: str | None = None) -> dict:
        """Sketch entity'leri arasina geometrik kisitlama ekler.

        type: 'coincident' (2), 'horizontal' (1), 'vertical' (1), 'parallel' (2),
        'perpendicular' (2), 'tangent' (2), 'equal' (2), 'symmetric' (3: e1, e2, simetri-cizgisi).
        entity_ids: inventor_sketch_add_* veya inventor_sketch_list_entities'den alinan id'ler.
        """
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)

        kind = type.lower()
        arity = CONSTRAINT_ARITY.get(kind)
        if arity is None:
            raise LookupError(f"Gecersiz kisitlama turu: '{type}'. Gecerli turler: {', '.join(CONSTRAINT_ARITY)}.")
        if len(entity_ids) != arity:
            raise LookupError(f"'{kind}' kisitlamasi {arity} entity gerektirir, {len(entity_ids)} verildi.")

        entities = [registry.resolve(eid) for eid in entity_ids]
        _add_constraint(sk.GeometricConstraints, kind, entities)
        return {"sketch": sk.Name, "constraint": kind, "entity_ids": entity_ids}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_dimension(
        sketch: str, type: str, entity_ids: list[str], value: float, doc: str | None = None
    ) -> dict:
        """Sketch entity'lerine boyut kisitlamasi ekler ve degerini ayarlar.

        type: 'linear' (2 nokta/cizgi arasi mesafe, mm), 'radius' (1 daire/yay, mm),
        'diameter' (1 daire, mm), 'angle' (2 cizgi arasi aci, derece).
        """
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        dims = sk.DimensionConstraints

        kind = type.lower()
        entities = [registry.resolve(eid) for eid in entity_ids]

        if kind == "linear":
            if len(entities) != 2:
                raise LookupError("'linear' boyutu 2 entity (nokta veya cizgi) gerektirir.")
            p1, p2 = _entity_point(entities[0]), _entity_point(entities[1])
            text_point = _midpoint_text_point(app, p1, p2)
            constraint = dims.AddTwoPointDistance(p1, p2, constants.kAlignedDim, text_point)
            constraint.Parameter.Value = mm_to_cm(value)
        elif kind in ("radius", "diameter"):
            if len(entities) != 1:
                raise LookupError(f"'{kind}' boyutu 1 entity (daire veya yay) gerektirir.")
            text_point = _radial_text_point(app, entities[0])
            if kind == "diameter":
                constraint = dims.AddDiameter(entities[0], text_point)
            else:
                constraint = dims.AddRadius(entities[0], text_point)
            constraint.Parameter.Value = mm_to_cm(value)
        elif kind == "angle":
            if len(entities) != 2:
                raise LookupError("'angle' boyutu 2 cizgi gerektirir.")
            text_point = _angle_text_point(app, entities[0], entities[1])
            constraint = dims.AddTwoLineAngle(entities[0], entities[1], text_point)
            constraint.Parameter.Value = deg_to_rad(value)
        else:
            raise LookupError(f"Gecersiz boyut turu: '{type}'. Gecerli turler: linear, radius, diameter, angle.")

        return {"sketch": sk.Name, "dimension": kind, "value": value, "parameter": constraint.Parameter.Name}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_ellipse(
        sketch: str,
        cx: float,
        cy: float,
        major_radius: float,
        minor_radius: float,
        angle_deg: float = 0,
        doc: str | None = None,
    ) -> dict:
        """Sketch'e merkez (mm), buyuk/kucuk eksen yaricaplari (mm) verilen bir
        elips ekler (UI'daki 'Sketch > Create > Ellipse' komutuna esdeger).
        angle_deg: buyuk eksenin +X ekseninden itibaren acisi (derece).
        """
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        tg = app.TransientGeometry
        center = point2d(app, cx, cy)
        direction = tg.CreateUnitVector2d(math.cos(math.radians(angle_deg)), math.sin(math.radians(angle_deg)))
        ellipse = sk.SketchEllipses.Add(center, direction, mm_to_cm(major_radius), mm_to_cm(minor_radius))
        entity_id = registry.register(document.DisplayName, sk.Name, "ellipse", ellipse)
        return {"entity_id": entity_id, **_ellipse_summary(ellipse)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_add_spline(
        sketch: str, points: list[list[float]], closed: bool = False, doc: str | None = None
    ) -> dict:
        """Sketch'e verilen N gecis noktasindan (mm, [[x,y], ...] en az 3 nokta)
        gecen interpole edilmis bir spline ekler (UI'daki Line aracinin
        'Spline' alt-secenegine esdeger). closed=True ise spline kapatilir
        (son nokta ilk noktaya teget/surekli baglanir).
        """
        if len(points) < 3:
            raise LookupError("Bir spline en az 3 gecis noktasi gerektirir.")
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        fit_points = app.TransientObjects.CreateObjectCollection()
        for x, y in points:
            fit_points.Add(point2d(app, x, y))
        spline = sk.SketchSplines.Add(fit_points)
        if closed:
            try:
                spline.Closed = True
            except (pywintypes.com_error, AttributeError):
                raise LookupError("Bu Inventor surumunde spline'i API uzerinden kapatmak desteklenmiyor.")
        entity_id = registry.register(document.DisplayName, sk.Name, "spline", spline)
        return {"entity_id": entity_id, "points": points, "closed": bool(closed)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_set_construction(
        sketch: str, entity_ids: list[str], construction: bool = True, doc: str | None = None
    ) -> dict:
        """Verilen sketch entity'lerini yapi (construction) geometrisine cevirir
        (ya da geri normal geometriye dondurur). Yapi geometrisi profil/kesit
        algilamasina dahil edilmez; sadece referans/kisitlama amaclidir
        (UI'daki Sketch > Format > Construction komutuna esdeger)."""
        for eid in entity_ids:
            entity = registry.resolve(eid)
            entity.Construction = construction
        return {"entity_ids": entity_ids, "construction": construction}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch_mirror(
        sketch: str,
        entity_ids: list[str],
        line_point1: list[float],
        line_point2: list[float],
        doc: str | None = None,
    ) -> dict:
        """Verilen entity'lerin (cizgi/daire/yay), iki nokta (mm) ile tanimlanan
        bir ayna dogrusuna gore YANSITILMIS kopyalarini olusturur.

        Not: Resmi Inventor API'sinde sketch-seviyesinde dogrudan bir 'Mirror'
        metodu YOK (Autodesk'in kendi forumlarinda dogrulanmis bir kisitlama);
        UI'daki Mirror komutunun ayni sonucunu, her entity'nin tanimlayici
        noktalarini matematiksel olarak yansitip yeni geometri olusturarak
        elde eder (orijinaller silinmez).
        """
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        x1, y1 = line_point1
        x2, y2 = line_point2

        new_ids = []
        for eid in entity_ids:
            entity = _as_concrete_entity(registry.resolve(eid))
            if hasattr(entity, "StartSketchPoint") and hasattr(entity, "CenterSketchPoint"):
                # Arc: merkez + uc noktalari yansit; yon tersine donecegi icin
                # baslangic/bitis noktalarini takas ederek doğru sarmal/sweep
                # yonunu koru.
                c = entity.CenterSketchPoint.Geometry
                sp = entity.StartSketchPoint.Geometry
                ep = entity.EndSketchPoint.Geometry
                rc = _reflect_point(cm_to_mm(c.X), cm_to_mm(c.Y), x1, y1, x2, y2)
                rs = _reflect_point(cm_to_mm(sp.X), cm_to_mm(sp.Y), x1, y1, x2, y2)
                re = _reflect_point(cm_to_mm(ep.X), cm_to_mm(ep.Y), x1, y1, x2, y2)
                new_entity = sk.SketchArcs.AddByCenterStartEndPoint(
                    point2d(app, *rc), point2d(app, *re), point2d(app, *rs)
                )
                kind = "arc"
            elif hasattr(entity, "StartSketchPoint"):
                # Line
                sp = entity.StartSketchPoint.Geometry
                ep = entity.EndSketchPoint.Geometry
                rs = _reflect_point(cm_to_mm(sp.X), cm_to_mm(sp.Y), x1, y1, x2, y2)
                re = _reflect_point(cm_to_mm(ep.X), cm_to_mm(ep.Y), x1, y1, x2, y2)
                new_entity = sk.SketchLines.AddByTwoPoints(point2d(app, *rs), point2d(app, *re))
                kind = "line"
            elif hasattr(entity, "CenterSketchPoint"):
                # Circle
                c = entity.CenterSketchPoint.Geometry
                rc = _reflect_point(cm_to_mm(c.X), cm_to_mm(c.Y), x1, y1, x2, y2)
                new_entity = sk.SketchCircles.AddByCenterRadius(point2d(app, *rc), entity.Radius)
                kind = "circle"
            else:
                raise LookupError(f"'{eid}' yansitma icin desteklenmeyen bir entity turu.")
            new_ids.append(registry.register(document.DisplayName, sk.Name, kind, new_entity))

        return {"entity_ids": new_ids, "mirrored_from": entity_ids, "line_point1": line_point1, "line_point2": line_point2}

    @mcp.tool()
    @inventor_tool
    async def inventor_finish_sketch(sketch: str, doc: str | None = None) -> dict:
        """Sketch duzenlemeyi sonlandirir ve sketch'teki kapali bolge (profil yolu)
        sayisini dondurur. closed_regions=0 ise sketch extrude/revolve edilemez:
        profil kapali degil (acik uc, cakisan/kopuk cizgi) demektir."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch(document, sketch)
        try:
            sk.ExitEdit()
        except (pywintypes.com_error, AttributeError):
            pass
        closed = count_closed_regions(sk)
        result = {"sketch": sk.Name, "closed_regions": closed, "profile_count": closed}
        if closed == 0:
            result["warning"] = "Kapali profil yok: extrude/revolve basarisiz olur. Cizgi uclarini kontrol edin."
        return result
