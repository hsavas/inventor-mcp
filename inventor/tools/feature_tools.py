"""3D feature araclari: extrude/revolve/hole/fillet/chamfer/shell/pattern/mirror/work features."""

import pythoncom
import pywintypes
import win32com.client.dynamic
from win32com.client import constants

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import get_document, get_feature, get_part_component_def, get_sketch, inventor_tool
from inventor.units import deg_to_rad, mm_to_cm

OPERATION_MAP = {
    "join": constants.kJoinOperation,
    "cut": constants.kCutOperation,
    "intersect": constants.kIntersectOperation,
    "new_body": constants.kNewBodyOperation,
}

DIRECTION_MAP = {
    "positive": constants.kPositiveExtentDirection,
    "negative": constants.kNegativeExtentDirection,
    "symmetric": constants.kSymmetricExtentDirection,
}



def _operation(name: str):
    op = OPERATION_MAP.get(name.lower())
    if op is None:
        raise LookupError(f"Gecersiz operation: '{name}'. Gecerli degerler: {', '.join(OPERATION_MAP)}.")
    return op


def _direction(name: str):
    d = DIRECTION_MAP.get(name.lower())
    if d is None:
        raise LookupError(f"Gecersiz direction: '{name}'. Gecerli degerler: {', '.join(DIRECTION_MAP)}.")
    return d


def _resolve_axis(comp_def, axis_ref: str):
    try:
        return registry.resolve(axis_ref)
    except LookupError:
        pass
    try:
        return comp_def.WorkAxes.Item(axis_ref)
    except pywintypes.com_error:
        raise LookupError(f"'{axis_ref}' bir entity_id veya work axis adi olarak bulunamadi.")


def _resolve_plane(comp_def, plane_ref: str):
    try:
        return registry.resolve(plane_ref)
    except LookupError:
        pass
    try:
        return comp_def.WorkPlanes.Item(plane_ref)
    except pywintypes.com_error:
        raise LookupError(f"'{plane_ref}' bir entity_id veya work plane adi olarak bulunamadi.")


def _object_collection(app, items):
    coll = app.TransientObjects.CreateObjectCollection()
    for item in items:
        coll.Add(item)
    return coll


def _edge_collection(app, items):
    coll = app.TransientObjects.CreateEdgeCollection()
    for item in items:
        coll.Add(item)
    return coll


def _face_collection(app, items):
    coll = app.TransientObjects.CreateFaceCollection()
    for item in items:
        coll.Add(item)
    return coll


def _resolve_edges(app, comp_def, edge_refs: list[str] | None):
    """edge_refs verilmezse ilk govdedeki tum kenarlari, verilirse registry'den
    cozulen kenar nesnelerini bir EdgeCollection olarak dondurur."""
    if not edge_refs:
        body = comp_def.SurfaceBodies.Item(1)
        return _edge_collection(app, list(body.Edges))
    return _edge_collection(app, [registry.resolve(ref) for ref in edge_refs])


def _apply_affected_bodies(app, comp_def, definition, affected_bodies: list[int] | None):
    """join/cut/intersect islemlerinin hangi govdeleri etkileyecegini belirler.

    Inventor'un varsayilani SADECE ILK govdedir. Cok govdeli bir partta bu sessizce
    yanlis sonuc verir: yeni malzeme, gercekte kestigi/dokundugu govdelerle
    birlesmek yerine 1. govdeye eklenir ve geri kalanlarla ic ice gecmis
    (interference) ayri govdeler olusur. Bu yuzden hedef govdeler acikca verilebilir.
    """
    if not affected_bodies:
        return
    bodies = [comp_def.SurfaceBodies.Item(i) for i in affected_bodies]
    definition.AffectedBodies = _object_collection(app, bodies)


def _combine_into(app, comp_def, targets, tool_body, op):
    """tool_body'yi hedef govdelere Combine ile uygular.

    join: tum hedefler + tool tek govdede birlesir.
    cut/intersect: her hedefe ayri ayri uygulanir (tool son adima kadar korunur).
    """
    combine_features = win32com.client.dynamic.Dispatch(comp_def.Features.CombineFeatures._oleobj_)
    if op == constants.kJoinOperation:
        tools = _object_collection(app, list(targets[1:]) + [tool_body])
        combine_features.Add(targets[0], tools, op, False)
        return
    for idx, target in enumerate(targets):
        keep = idx < len(targets) - 1
        combine_features.Add(target, _object_collection(app, [tool_body]), op, keep)


def _new_body_then_combine(app, comp_def, affected_bodies, op, make_feature):
    """Definition nesnesi olmayan feature'lar icin AffectedBodies esdegeri.

    Inventor 2020'de RevolveFeatures'in CreateRevolveDefinition'i (ve dolayisiyla
    AffectedBodies'i) yoktur; AddFull/AddByAngle sadece bir operation alir ve
    join/cut yine SADECE ilk govdeyi etkiler. Bu yuzden feature once AYRI bir
    govde olarak olusturulur, sonra Combine ile istenen govdelere uygulanir.
    """
    targets = [comp_def.SurfaceBodies.Item(i) for i in affected_bodies]
    names_before = {b.Name for b in comp_def.SurfaceBodies}

    feature = make_feature(constants.kNewBodyOperation)

    tool_body = None
    for b in comp_def.SurfaceBodies:
        if b.Name not in names_before:
            tool_body = b
            break
    if tool_body is None:
        tool_body = comp_def.SurfaceBodies.Item(comp_def.SurfaceBodies.Count)

    _combine_into(app, comp_def, targets, tool_body, op)
    return feature


def _feature_summary(feature) -> dict:
    return {"name": feature.Name, "type": feature.Type if hasattr(feature, "Type") else feature.__class__.__name__}


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_extrude(
        sketch: str,
        distance_mm: float,
        operation: str = "join",
        direction: str = "positive",
        taper_deg: float = 0.0,
        affected_bodies: list[int] | None = None,
        doc: str | None = None,
    ) -> dict:
        """Sketch profilini extrude ederek bir feature olusturur.

        operation: 'join', 'cut', 'intersect', 'new_body'.
        direction: 'positive', 'negative', 'symmetric'.
        taper_deg: opsiyonel taslama acisi (derece).
        affected_bodies: join/cut/intersect'in etkileyecegi govdelerin 1 tabanli
        indeksleri (inventor_list_bodies). Verilmezse Inventor SADECE ILK govdeyi
        etkiler; cok govdeli partlarda bu neredeyse her zaman istenmeyen sonuc
        verir, bu yuzden govde sayisi birden fazlaysa acikca belirtin.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        sk = get_sketch(document, sketch)

        op = _operation(operation)
        dir_enum = _direction(direction)

        profile = sk.Profiles.AddForSolid()
        ext_features = comp_def.Features.ExtrudeFeatures
        ext_def = ext_features.CreateExtrudeDefinition(profile, op)
        if taper_deg:
            ext_def.TaperAngle = deg_to_rad(taper_deg)
        ext_def.SetDistanceExtent(mm_to_cm(distance_mm), dir_enum)
        _apply_affected_bodies(app, comp_def, ext_def, affected_bodies)
        feature = ext_features.Add(ext_def)
        return {"name": feature.Name, "type": "extrude"}

    @mcp.tool()
    @inventor_tool
    async def inventor_revolve(
        sketch: str,
        axis: str,
        angle_deg: float = 360.0,
        operation: str = "join",
        affected_bodies: list[int] | None = None,
        doc: str | None = None,
    ) -> dict:
        """Sketch profilini bir eksen etrafinda dondurerek (revolve) feature olusturur.

        axis: bir entity_id (sketch cizgisi) veya work axis adi (orn. 'Y Axis').
        angle_deg: 360 ise tam tur, aksi halde belirtilen aci kadar.
        affected_bodies: bkz. inventor_extrude - verilmezse Inventor sadece ilk
        govdeyi etkiler. Revolve'da bu, feature'i once ayri bir govde olarak
        olusturup Combine ile uygulayarak saglanir (Inventor 2020'de revolve'un
        AffectedBodies destegi yoktur), yani feature agacinda ek bir Combine
        gorunur.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        sk = get_sketch(document, sketch)

        op = _operation(operation)
        axis_entity = _resolve_axis(comp_def, axis)

        profile = sk.Profiles.AddForSolid()
        rev_features = comp_def.Features.RevolveFeatures

        def _make(operation_enum):
            if angle_deg >= 360.0:
                return rev_features.AddFull(profile, axis_entity, operation_enum)
            return rev_features.AddByAngle(
                profile, axis_entity, deg_to_rad(angle_deg),
                constants.kPositiveExtentDirection, operation_enum,
            )

        if affected_bodies and op != constants.kNewBodyOperation:
            feature = _new_body_then_combine(app, comp_def, affected_bodies, op, _make)
        else:
            feature = _make(op)
        return {"name": feature.Name, "type": "revolve"}

    @mcp.tool()
    @inventor_tool
    async def inventor_create_coil(
        sketch: str,
        axis: str,
        coil_type: str = "pitch_and_height",
        pitch_mm: float | None = None,
        revolutions: float | None = None,
        height_mm: float | None = None,
        taper_deg: float = 0.0,
        rotation: str = "right",
        operation: str = "new_body",
        doc: str | None = None,
    ) -> dict:
        """Sketch profilini bir eksen etrafinda helis (coil/sarmal) yolunda supurerek
        gercek helisel bir feature olusturur. Vida/somun disleri gibi gercek 3B
        helisel geometriler bu yontemle modellenir (UI'daki 'Coil' komutuna esdeger).

        Resmi Inventor API'sinde CoilFeatures.CreateCoilDefinition diye bir metot
        YOK (bircok eski ornekte gecse de); gercek API direkt
        AddByPitchAndHeight / AddByPitchAndRevolution / AddByRevolutionAndHeight /
        AddSpiral metotlarini kullanir - bu fonksiyon bunlari sarmalar.

        sketch: profili iceren PlanarSketch'in adi. Bu sketch, EKSENI ICEREN bir
        duzlemde (eksene DIK degil) cizilmis olmalidir - orn. dis profili icin
        XZ veya YZ duzleminde, eksenden dis buyuk capi kadar uzaklikta kucuk bir
        ucgen/kapali kesit. Profil tek bir helis donusu icin gerekli kesittir;
        Coil ozelligi bu kesiti otomatik olarak tum helis boyunca tekrarlar.
        axis: entity_id (sketch cizgisi) veya work axis adi (orn. 'Z Axis').
        coil_type:
            'pitch_and_revolution' -> pitch_mm + revolutions
            'revolution_and_height' -> revolutions + height_mm
            'pitch_and_height' -> pitch_mm + height_mm (vida disi icin tipik secim)
            'spiral' -> pitch_mm + revolutions (sabit yukseklikte, yaricap degisen spiral)
        taper_deg: opsiyonel helis konisi acisi (derece).
        rotation: 'right' (sag-el, standart vida) veya 'left' (sol-el).
        operation: 'join', 'cut' (govdeye dis acmak icin), 'intersect', 'new_body'.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        sk = get_sketch(document, sketch)

        op = _operation(operation)
        axis_entity = _resolve_axis(comp_def, axis)
        profile = sk.Profiles.AddForSolid()

        kind = coil_type.lower()
        clockwise = rotation.lower() == "right"
        if rotation.lower() not in ("right", "left"):
            raise LookupError(f"Gecersiz rotation: '{rotation}'. 'right' veya 'left' olmali.")
        taper_rad = deg_to_rad(taper_deg) if taper_deg else 0.0

        # win32com'un gen_py onbellegi bu Inventor kurulumunda CoilFeatures'in
        # gercek Add* metotlarini icermiyor olabilir (eski/eksik onbellek).
        # Bu yuzden CoilFeatures'i dinamik (late-bound) dispatch ile sariyoruz;
        # bu, onbellekteki metot listesini atlayip canli COM nesnesindeki
        # gercek metotlari isme gore cagirir.
        coil_features_raw = comp_def.Features.CoilFeatures
        coil_features = win32com.client.dynamic.Dispatch(coil_features_raw._oleobj_)

        try:
            if kind == "pitch_and_height":
                if pitch_mm is None or height_mm is None:
                    raise LookupError("pitch_and_height icin pitch_mm ve height_mm gerekli.")
                feature = coil_features.AddByPitchAndHeight(
                    profile, axis_entity, mm_to_cm(pitch_mm), mm_to_cm(height_mm), op,
                    False, clockwise, taper_rad,
                )
            elif kind == "pitch_and_revolution":
                if pitch_mm is None or revolutions is None:
                    raise LookupError("pitch_and_revolution icin pitch_mm ve revolutions gerekli.")
                feature = coil_features.AddByPitchAndRevolution(
                    profile, axis_entity, mm_to_cm(pitch_mm), revolutions, op,
                    False, clockwise, taper_rad,
                )
            elif kind == "revolution_and_height":
                if revolutions is None or height_mm is None:
                    raise LookupError("revolution_and_height icin revolutions ve height_mm gerekli.")
                feature = coil_features.AddByRevolutionAndHeight(
                    profile, axis_entity, revolutions, mm_to_cm(height_mm), op,
                    False, clockwise, taper_rad,
                )
            elif kind == "spiral":
                if pitch_mm is None or revolutions is None:
                    raise LookupError("spiral icin pitch_mm ve revolutions gerekli.")
                feature = coil_features.AddSpiral(
                    profile, axis_entity, mm_to_cm(pitch_mm), revolutions, op,
                    False, clockwise,
                )
            else:
                raise LookupError(
                    f"Gecersiz coil_type: '{coil_type}'. Gecerli degerler: "
                    "pitch_and_height, pitch_and_revolution, revolution_and_height, spiral."
                )
        except pywintypes.com_error as exc:
            raise RuntimeError(f"Coil feature olusturulamadi: {exc}")

        return {"name": feature.Name, "type": "coil"}

    def _resolve_path_curve(document, comp_def, path_ref: str):
        """path_ref bir entity_id (kayitli cizgi/yay/spline) ise onu, degilse bir
        sketch adi olarak yorumlayip o sketch'teki ilk egriyi dondurur."""
        try:
            return registry.resolve(path_ref)
        except LookupError:
            pass
        sk = get_sketch(document, path_ref)
        for coll_name in ("SketchLines", "SketchArcs", "SketchSplines"):
            coll = getattr(sk, coll_name)
            if coll.Count > 0:
                return coll.Item(1)
        raise LookupError(f"'{path_ref}' bir entity_id veya icinde egrisi olan bir sketch adi olarak bulunamadi.")

    @mcp.tool()
    @inventor_tool
    async def inventor_sweep(
        sketch: str,
        path: str,
        operation: str = "join",
        taper_deg: float = 0.0,
        doc: str | None = None,
    ) -> dict:
        """Sketch profilini bir yol (path) boyunca supurerek (sweep) feature olusturur
        (UI'daki 'Sweep' komutuna esdeger). Profil ve yol AYRI sketch'lerde olmalidir.

        sketch: profili iceren sketch'in adi.
        path: yol egrisi - bir entity_id (cizgi/yay/spline, inventor_sketch_list_entities'den)
        veya icinde tek bir baglantili egri zinciri olan bir sketch adi (o sketch'teki
        ilk egri otomatik secilir).
        taper_deg: opsiyonel taslama acisi (derece, sadece duz/dogrusal yollarda gecerlidir).
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        sk = get_sketch(document, sketch)

        op = _operation(operation)
        profile = sk.Profiles.AddForSolid()
        path_curve = _resolve_path_curve(document, comp_def, path)
        path_obj = comp_def.Features.CreatePath(path_curve)

        # CoilFeatures'da gorulen ayni gen_py onbellek sorunuyla karsilasmamak
        # icin SweepFeatures'i de dinamik (late-bound) dispatch ile sariyoruz.
        sweep_features = win32com.client.dynamic.Dispatch(comp_def.Features.SweepFeatures._oleobj_)
        try:
            if taper_deg:
                feature = sweep_features.AddUsingPath(profile, path_obj, op, deg_to_rad(taper_deg))
            else:
                feature = sweep_features.AddUsingPath(profile, path_obj, op)
        except pywintypes.com_error as exc:
            raise RuntimeError(f"Sweep feature olusturulamadi: {exc}")
        return {"name": feature.Name, "type": "sweep"}

    @mcp.tool()
    @inventor_tool
    async def inventor_loft(
        sections: list[str],
        operation: str = "join",
        closed: bool = False,
        doc: str | None = None,
    ) -> dict:
        """Birden fazla sketch profilini ('sections', en az 2 sketch adi) birbirine
        yumusak gecisle baglayan bir loft feature olusturur (UI'daki 'Loft'
        komutuna esdeger). Profiller sirayla, dis dis bir zincir olarak harmanlanir.

        closed=True ise son kesit ilk kesite de baglanarak kapali (halka seklinde)
        bir govde olusturulur.
        """
        if len(sections) < 2:
            raise LookupError("Loft icin en az 2 kesit (sketch) gerekir.")
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)

        op = _operation(operation)
        profiles = [get_sketch(document, name).Profiles.AddForSolid() for name in sections]
        profile_coll = _object_collection(app, profiles)

        loft_features = win32com.client.dynamic.Dispatch(comp_def.Features.LoftFeatures._oleobj_)
        try:
            loft_def = loft_features.CreateLoftDefinition(profile_coll, op)
            if closed:
                try:
                    loft_def.Closed = True
                except (pywintypes.com_error, AttributeError):
                    raise LookupError("Bu Inventor surumunde loft'u API uzerinden kapatmak desteklenmiyor.")
            feature = loft_features.Add(loft_def)
        except pywintypes.com_error as exc:
            raise RuntimeError(f"Loft feature olusturulamadi: {exc}")
        return {"name": feature.Name, "type": "loft", "sections": sections}

    @mcp.tool()
    @inventor_tool
    async def inventor_combine(
        base_body_index: int,
        tool_body_indices: list[int],
        operation: str = "join",
        keep_tool_bodies: bool = False,
        doc: str | None = None,
    ) -> dict:
        """Birden fazla katı govdeyi (SurfaceBody) boolean islemle birlestirir
        (UI'daki 'Combine' komutuna esdeger). operation: 'join' (birlestir/union),
        'cut' (cikar/subtract), 'intersect' (kesisim). Govde indeksleri (1 tabanli)
        icin inventor_list_bodies kullanin. keep_tool_bodies=True ise tool govdeler
        ayri govdeler olarak da saklanir (varsayilan: birlestirme sonrasi silinir).
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)

        op = _operation(operation)
        base_body = comp_def.SurfaceBodies.Item(base_body_index)
        tool_bodies = _object_collection(app, [comp_def.SurfaceBodies.Item(i) for i in tool_body_indices])

        combine_features = win32com.client.dynamic.Dispatch(comp_def.Features.CombineFeatures._oleobj_)
        try:
            feature = combine_features.Add(base_body, tool_bodies, op, keep_tool_bodies)
        except pywintypes.com_error as exc:
            raise RuntimeError(f"Combine feature olusturulamadi: {exc}")
        return {"name": feature.Name, "type": "combine", "operation": operation}

    @mcp.tool()
    @inventor_tool
    async def inventor_hole(
        sketch: str,
        point_entity_ids: list[str],
        diameter_mm: float,
        depth_mm: float | None = None,
        through_all: bool = False,
        hole_type: str = "simple",
        counterbore_diameter_mm: float | None = None,
        counterbore_depth_mm: float | None = None,
        countersink_diameter_mm: float | None = None,
        countersink_angle_deg: float | None = None,
        doc: str | None = None,
    ) -> dict:
        """Sketch'teki nokta(lar)dan delik (hole) feature'i olusturur.

        hole_type: 'simple', 'counterbore', 'countersink'.
        depth_mm verilmezse ve through_all=False ise delik 1mm derinliginde acilir.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        sk = get_sketch(document, sketch)

        points = _object_collection(app, [registry.resolve(eid) for eid in point_entity_ids])

        hole_features = comp_def.Features.HoleFeatures
        placement = hole_features.CreateSketchPlacementDefinition(points)

        diameter_cm = mm_to_cm(diameter_mm)
        direction = constants.kNegativeExtentDirection
        kind = hole_type.lower()

        if kind == "simple":
            if through_all:
                feature = hole_features.AddDrilledByThroughAllExtent(placement, diameter_cm, direction)
            else:
                depth_cm = mm_to_cm(depth_mm if depth_mm is not None else 1.0)
                feature = hole_features.AddDrilledByDistanceExtent(placement, diameter_cm, depth_cm, direction)
        elif kind == "counterbore":
            if counterbore_diameter_mm is None or counterbore_depth_mm is None:
                raise LookupError("counterbore icin counterbore_diameter_mm ve counterbore_depth_mm gerekli.")
            cbore_diameter_cm = mm_to_cm(counterbore_diameter_mm)
            cbore_depth_cm = mm_to_cm(counterbore_depth_mm)
            if through_all:
                feature = hole_features.AddCBoreByThroughAllExtent(placement, diameter_cm, direction, cbore_diameter_cm, cbore_depth_cm)
            else:
                depth_cm = mm_to_cm(depth_mm if depth_mm is not None else 1.0)
                feature = hole_features.AddCBoreByDistanceExtent(placement, diameter_cm, depth_cm, direction, cbore_diameter_cm, cbore_depth_cm)
        elif kind == "countersink":
            if countersink_diameter_mm is None or countersink_angle_deg is None:
                raise LookupError("countersink icin countersink_diameter_mm ve countersink_angle_deg gerekli.")
            csink_diameter_cm = mm_to_cm(countersink_diameter_mm)
            csink_angle_rad = deg_to_rad(countersink_angle_deg)
            if through_all:
                feature = hole_features.AddCSinkByThroughAllExtent(placement, diameter_cm, direction, csink_diameter_cm, csink_angle_rad)
            else:
                depth_cm = mm_to_cm(depth_mm if depth_mm is not None else 1.0)
                feature = hole_features.AddCSinkByDistanceExtent(placement, diameter_cm, depth_cm, direction, csink_diameter_cm, csink_angle_rad)
        else:
            raise LookupError(f"Gecersiz hole_type: '{hole_type}'. Gecerli degerler: simple, counterbore, countersink.")

        return {"name": feature.Name, "type": "hole"}

    @mcp.tool()
    @inventor_tool
    async def inventor_fillet(
        radius_mm: float,
        edge_refs: list[str] | None = None,
        feature_name: str | None = None,
        doc: str | None = None,
    ) -> dict:
        """Kenarlara sabit yaricapli fillet (yuvarlatma) ekler.

        edge_refs verilmezse ilk govdenin TUM kenarlarina uygulanir.
        feature_name verilirse, o feature'in (orn. bir kesme extrude'unun)
        olusturdugu TUM kenarlar tek seferde fillete eklenir (Inventor API'si
        EdgeCollection icine Feature nesnesi eklenmesine izin verir; bu, yuzlerce
        kenari tek tek entity_id ile belirtmek zorunda kalmadan tum slot kenarlari
        gibi buyuk kenar setlerini filtelemeyi saglar). edge_refs ve feature_name
        birlikte de verilebilir.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        if feature_name:
            edges = _resolve_edges(app, comp_def, edge_refs) if edge_refs else app.TransientObjects.CreateEdgeCollection()
            edges.Add(get_feature(document, feature_name))
        else:
            edges = _resolve_edges(app, comp_def, edge_refs)
        feature = comp_def.Features.FilletFeatures.AddSimple(edges, mm_to_cm(radius_mm))
        return {"name": feature.Name, "type": "fillet"}

    @mcp.tool()
    @inventor_tool
    async def inventor_chamfer(distance_mm: float, edge_refs: list[str] | None = None, doc: str | None = None) -> dict:
        """Kenarlara esit-mesafeli pah (chamfer) ekler.

        edge_refs verilmezse ilk govdenin TUM kenarlarina uygulanir.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        edges = _resolve_edges(app, comp_def, edge_refs)
        feature = comp_def.Features.ChamferFeatures.AddUsingDistance(edges, mm_to_cm(distance_mm))
        return {"name": feature.Name, "type": "chamfer"}

    @mcp.tool()
    @inventor_tool
    async def inventor_shell(thickness_mm: float, face_refs: list[str] | None = None, doc: str | None = None) -> dict:
        """Govdeyi icini bosaltir (shell). face_refs ile kaldirilacak yuzeyler belirtilebilir."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        faces = _face_collection(app, [registry.resolve(ref) for ref in face_refs]) if face_refs else app.TransientObjects.CreateFaceCollection()
        shell_def = comp_def.Features.ShellFeatures.CreateShellDefinition(faces, mm_to_cm(thickness_mm), constants.kInsideShellDirection)
        feature = comp_def.Features.ShellFeatures.Add(shell_def)
        return {"name": feature.Name, "type": "shell"}

    @mcp.tool()
    @inventor_tool
    async def inventor_rectangular_pattern(
        feature_name: str,
        x_count: int,
        x_spacing_mm: float,
        y_count: int = 1,
        y_spacing_mm: float = 0.0,
        x_axis: str = "X Axis",
        y_axis: str = "Y Axis",
        doc: str | None = None,
    ) -> dict:
        """Bir feature'i dikdortgen (lineer) duzende cogaltir.

        x_axis/y_axis: yon belirleyen work axis adlari (varsayilan parca eksenleri).
        y_count=1 ise sadece X yonunde tek-yonlu pattern olusturulur.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        feature = get_feature(document, feature_name)
        parents = _object_collection(app, [feature])

        dir1 = _resolve_axis(comp_def, x_axis)
        if y_count and y_count > 1:
            dir2 = _resolve_axis(comp_def, y_axis)
            rect_feature = comp_def.Features.RectangularPatternFeatures.Add(
                ParentFeatures=parents,
                XDirectionEntity=dir1,
                NaturalXDirection=True,
                XCount=x_count,
                XSpacing=mm_to_cm(x_spacing_mm),
                YDirectionEntity=dir2,
                NaturalYDirection=True,
                YCount=y_count,
                YSpacing=mm_to_cm(y_spacing_mm),
                ComputeType=constants.kIdenticalCompute,
            )
        else:
            rect_feature = comp_def.Features.RectangularPatternFeatures.Add(
                ParentFeatures=parents,
                XDirectionEntity=dir1,
                NaturalXDirection=True,
                XCount=x_count,
                XSpacing=mm_to_cm(x_spacing_mm),
                ComputeType=constants.kIdenticalCompute,
            )
        return {"name": rect_feature.Name, "type": "rectangular_pattern"}

    @mcp.tool()
    @inventor_tool
    async def inventor_circular_pattern(
        feature_name: str,
        axis: str,
        count: int,
        angle_deg: float = 360.0,
        doc: str | None = None,
    ) -> dict:
        """Bir feature'i bir eksen etrafinda dairesel duzende cogaltir.

        axis: entity_id veya work axis adi (orn. 'Z Axis').
        angle_deg=360 ise esit araliklarla tam tur.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        feature = get_feature(document, feature_name)
        parents = _object_collection(app, [feature])
        axis_entity = _resolve_axis(comp_def, axis)

        circ_feature = comp_def.Features.CircularPatternFeatures.Add(
            parents, axis_entity, True, count, deg_to_rad(angle_deg), True, constants.kIdenticalCompute
        )
        return {"name": circ_feature.Name, "type": "circular_pattern"}

    @mcp.tool()
    @inventor_tool
    async def inventor_mirror(feature_name: str, plane: str, doc: str | None = None) -> dict:
        """Bir feature'i bir duzleme gore aynalar.

        plane: 'XY', 'XZ', 'YZ' varsayilan duzlemlerden biri ya da work plane adi.
        """
        from inventor.tools.sketch_tools import PLANE_NAMES

        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        feature = get_feature(document, feature_name)
        parents = _object_collection(app, [feature])

        plane_name = PLANE_NAMES.get(plane.upper(), plane)
        plane_obj = _resolve_plane(comp_def, plane_name)

        mirror_feature = comp_def.Features.MirrorFeatures.Add(parents, plane_obj)
        return {"name": mirror_feature.Name, "type": "mirror"}

    @mcp.tool()
    @inventor_tool
    async def inventor_create_work_plane(base_plane: str, offset_mm: float, name: str | None = None, doc: str | None = None) -> dict:
        """Var olan bir duzlemden belirli bir mesafede ofsetli yeni bir work plane olusturur.

        base_plane: 'XY', 'XZ', 'YZ' veya var olan bir work plane adi.
        """
        from inventor.tools.sketch_tools import PLANE_NAMES

        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)

        plane_name = PLANE_NAMES.get(base_plane.upper(), base_plane)
        plane_obj = _resolve_plane(comp_def, plane_name)

        new_plane = comp_def.WorkPlanes.AddByPlaneAndOffset(plane_obj, mm_to_cm(offset_mm))
        if name:
            new_plane.Name = name
        entity_id = registry.register(document.DisplayName, "_workfeatures", "plane", new_plane)
        return {"entity_id": entity_id, "name": new_plane.Name}

    @mcp.tool()
    @inventor_tool
    async def inventor_create_work_axis(edge_ref: str, doc: str | None = None) -> dict:
        """Bir dogrusal kenardan/cizgiden gecen yeni bir work axis olusturur."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        edge = registry.resolve(edge_ref)
        new_axis = comp_def.WorkAxes.AddByLine(edge)
        entity_id = registry.register(document.DisplayName, "_workfeatures", "axis", new_axis)
        return {"entity_id": entity_id, "name": new_axis.Name}

    @mcp.tool()
    @inventor_tool
    async def inventor_create_work_point(x_mm: float, y_mm: float, z_mm: float, doc: str | None = None) -> dict:
        """Belirtilen 3D koordinatta (mm, parca orijinine gore) yeni bir work point olusturur."""
        from inventor.helpers import point3d

        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        p = point3d(app, x_mm, y_mm, z_mm)
        new_point = comp_def.WorkPoints.AddFixed(p)
        entity_id = registry.register(document.DisplayName, "_workfeatures", "point", new_point)
        return {"entity_id": entity_id, "name": new_point.Name}

    @mcp.tool()
    @inventor_tool
    async def inventor_list_features(doc: str | None = None) -> list[dict]:
        """Part dokumanindaki tum feature'lari sirayla listeler (ad, tur, baski/sonme durumu)."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        result = []
        for feature in comp_def.Features:
            result.append(
                {
                    "name": feature.Name,
                    "type": feature.Type,
                    "suppressed": bool(feature.Suppressed),
                }
            )
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_delete_feature(name: str, doc: str | None = None) -> dict:
        """Adi verilen bir feature'i (extrude, fillet, pattern, vb.) siler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        for feature in comp_def.Features:
            if feature.Name == name:
                feature.Delete()
                return {"deleted": name}
        raise LookupError(f"'{name}' adli bir feature bulunamadi.")

    @mcp.tool()
    @inventor_tool
    async def inventor_delete_sketch(name: str, doc: str | None = None) -> dict:
        """Adi verilen bir sketch'i siler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        for sk in comp_def.Sketches:
            if sk.Name == name:
                sk.Delete()
                return {"deleted": name}
        raise LookupError(f"'{name}' adli bir sketch bulunamadi.")
