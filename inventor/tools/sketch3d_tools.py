"""3D sketch araclari: keyfi yonelimli 3B cizgi/spline (gercek geodezik/serbest yol icin).

Standart PlanarSketch (inventor_create_sketch) sadece XY/XZ/YZ ve bunlarin
paralel-ofset duzlemlerinde calisir; bu, orijin merkezli ama eksen-hizali
OLMAYAN bir duzlemde yatan bir egri (orn. bir kure uzerindeki iki keyfi
nokta arasindaki buyuk-cember yayi) cizmeyi imkansiz kilar. Inventor'in
Sketch3D nesnesi (PartComponentDefinition.Sketches3D) herhangi bir duzleme
bagli olmadan dogrudan 3B koordinatlarda cizgi/spline olusturabilir; bu da
Sweep feature'in path'i olarak kullanilabilir - boylece keyfi yonelimli
kavisli/duz 3B yollar (orn. kure yuzeyini takip eden cubuklar) gercek
Inventor geometrisi olarak insa edilebilir.
"""

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import get_document, get_part_component_def, get_sketch3d, inventor_tool, point3d, has_member
from inventor.units import cm_to_mm


def _point3d_summary(pt) -> list:
    g = pt.Geometry if has_member(pt, "Geometry") else pt
    return [cm_to_mm(g.X), cm_to_mm(g.Y), cm_to_mm(g.Z)]


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_create_sketch3d(name: str | None = None, doc: str | None = None) -> dict:
        """Aktif Part dokumaninda yeni bir 3D sketch (Sketch3D) olusturur.

        PlanarSketch'in aksine belirli bir duzleme bagli degildir; icine
        eklenen cizgi/spline'lar herhangi bir 3B yonelimde olabilir. Ozellikle
        Sweep feature'i icin keyfi yonelimli path (yol) egrileri olusturmak
        icin kullanilir (inventor_sketch3d_add_line / _add_spline).
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        sketch3d = comp_def.Sketches3D.Add()
        if name:
            sketch3d.Name = name
        return {"name": sketch3d.Name}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch3d_add_line(
        sketch: str,
        x1: float, y1: float, z1: float,
        x2: float, y2: float, z2: float,
        doc: str | None = None,
    ) -> dict:
        """3D sketch'e iki keyfi 3B nokta (mm) arasinda bir dogru ekler.

        Duz (kiris) cubuklar veya hizalama eksenleri icin kullanilir. Bu
        dogru daha sonra inventor_sweep'e 'path' olarak verilebilir.
        """
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch3d(document, sketch)
        p1 = point3d(app, x1, y1, z1)
        p2 = point3d(app, x2, y2, z2)
        line = sk.SketchLines3D.AddByTwoPoints(p1, p2)
        entity_id = registry.register(document.DisplayName, sk.Name, "line3d", line)
        return {"entity_id": entity_id, "start_mm": [x1, y1, z1], "end_mm": [x2, y2, z2]}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch3d_add_spline(
        sketch: str,
        points: list[list[float]],
        closed: bool = False,
        doc: str | None = None,
    ) -> dict:
        """3D sketch'e verilen N gecis noktasindan (mm, [[x,y,z], ...] en az 2 nokta)
        gecen interpole edilmis bir 3B spline ekler.

        Keyfi yonelimli kavisli yollar icin kullanilir (orn. bir kurenin
        yuzeyini takip eden bir buyuk-cember/geodezik yay: cagiran taraf,
        arc'i onceden N noktaya orneklemeli ve bu noktalari buraya vermelidir).
        Sonuc, inventor_sweep'e 'path' olarak verilebilir.
        """
        if len(points) < 2:
            raise ValueError("Bir 3B spline en az 2 gecis noktasi gerektirir.")
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch3d(document, sketch)
        fit_points = app.TransientObjects.CreateObjectCollection()
        for x, y, z in points:
            fit_points.Add(point3d(app, x, y, z))
        spline = sk.SketchSplines3D.Add(fit_points)
        if closed:
            try:
                spline.Closed = True
            except Exception:
                raise ValueError("Bu Inventor surumunde 3B spline'i API uzerinden kapatmak desteklenmiyor.")
        entity_id = registry.register(document.DisplayName, sk.Name, "spline3d", spline)
        return {"entity_id": entity_id, "points": points, "closed": bool(closed)}

    @mcp.tool()
    @inventor_tool
    async def inventor_sketch3d_list_entities(sketch: str, doc: str | None = None) -> dict:
        """3D sketch'teki tum cizgi/spline nesnelerini yeniden tarar, kayit defterini
        tazeler ve her biri icin entity_id dondurur."""
        app = get_app()
        document = get_document(app, doc)
        sk = get_sketch3d(document, sketch)
        registry.clear_sketch(document.DisplayName, sk.Name)

        entities = []
        for line in sk.SketchLines3D:
            entity_id = registry.register(document.DisplayName, sk.Name, "line3d", line)
            entities.append(
                {
                    "entity_id": entity_id,
                    "kind": "line3d",
                    "start_mm": _point3d_summary(line.StartSketchPoint),
                    "end_mm": _point3d_summary(line.EndSketchPoint),
                }
            )
        for spline in sk.SketchSplines3D:
            entity_id = registry.register(document.DisplayName, sk.Name, "spline3d", spline)
            entities.append({"entity_id": entity_id, "kind": "spline3d"})

        return {"sketch": sk.Name, "entities": entities}
