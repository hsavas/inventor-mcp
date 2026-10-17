"""Montaj eklem (joint) araclari: rigid/rotational/slider/cylindrical/planar/ball.

Kisitlamalardan (constraint) farkli olarak bir joint, iki bilesen arasindaki
serbestlik derecelerini tek seferde tanimlar. Inventor, entity_one'in ait oldugu
occurrence'i entity_two'ya tasir (UI'daki "ilk secim ikinciye tasinir" davranisi).
"""

import math

import pywintypes
from win32com.client import CastTo, constants

from inventor.connection import get_app
from inventor.helpers import get_assembly_component_def, get_document, inventor_tool
from inventor.tools.assembly_tools import resolve_entity
from inventor.units import cm_to_mm, deg_to_rad, mm_to_cm, rad_to_deg

# AssemblyJointTypeEnum. Not: slider icin tip kutuphanesindeki ad kSlideJointType'dir.
JOINT_TYPES = {
    "rigid": constants.kRigidJointType,
    "rotational": constants.kRotationalJointType,
    "slider": constants.kSlideJointType,
    "cylindrical": constants.kCylindricalJointType,
    "planar": constants.kPlanarJointType,
    "ball": constants.kBallJointType,
}
JOINT_TYPE_NAMES = {v: k for k, v in JOINT_TYPES.items()}

# Pozisyon/limit ozelliklerini destekleyen turler. Desteklemeyen bir turde limit
# atamak hata vermiyor, Joints.Add ise genel bir COM hatasiyla basarisiz oluyor;
# bu yuzden limitler eklemeden once burada kontrol edilir.
_ANGULAR_TYPES = ("rotational", "cylindrical")
_LINEAR_TYPES = ("slider", "cylindrical")

HEALTH_NAMES = {
    constants.kUpToDateHealth: "up_to_date",
    constants.kOutOfDateHealth: "out_of_date",
    constants.kDriverLostHealth: "driver_lost",
    constants.kInErrorHealth: "in_error",
    constants.kDeletedHealth: "deleted",
    constants.kCannotComputeHealth: "cannot_compute",
    constants.kSuppressedHealth: "suppressed",
    constants.kBeyondStopNodeHealth: "beyond_stop_node",
    constants.kInconsistentHealth: "inconsistent",
    constants.kRedundantHealth: "redundant",
    constants.kNewlyAddedHealth: "newly_added",
    constants.kInvalidLimitsHealth: "invalid_limits",
    constants.kJointDOFLockedHealth: "dof_locked",
    constants.kUnknownHealth: "unknown",
}

# Kenarlar icin PointIntentEnum.
_EDGE_POINT_INTENTS = {
    "start": constants.kStartPointIntent,
    "mid": constants.kMidPointIntent,
    "end": constants.kEndPointIntent,
    "center": constants.kCenterPointIntent,
}
# Silindirik/konik yuzeyler sadece eksen uzerindeki noktalarla calisir
# (kCenterPointIntent/kMidPointIntent ile Joints.Add hata veriyor).
_AXIS_POINT_INTENTS = {
    "start": constants.kAxisStartPointIntent,
    "mid": constants.kAxisMidPointIntent,
    "center": constants.kAxisMidPointIntent,
    "end": constants.kAxisEndPointIntent,
}


def _point_intent(entity, point: str):
    """Entity turune gore 'start'/'mid'/'end'/'center' secimini PointIntentEnum'a cevirir."""
    p = point.lower()
    if p not in _EDGE_POINT_INTENTS:
        raise LookupError(f"Gecersiz point: '{point}'. Gecerli degerler: start, mid, end, center.")
    if entity.Type not in (constants.kFaceObject, constants.kFaceProxyObject):
        return _EDGE_POINT_INTENTS[p]
    surface_type = entity.SurfaceType
    if surface_type == constants.kPlaneSurface:
        if p != "center":
            raise LookupError("Duzlemsel yuzeyler icin sadece point='center' (yuz merkezi) gecerlidir.")
        return constants.kPlanarFaceCenterPointIntent
    if surface_type in (constants.kCylinderSurface, constants.kConeSurface):
        return _AXIS_POINT_INTENTS[p]
    raise LookupError("Joint icin sadece duzlemsel, silindirik veya konik yuzeyler (ya da kenarlar) kullanilabilir.")


def _param_value(definition, prop: str, convert):
    """Joint tanimindaki ModelParameter degerini dondurur; joint turu desteklemiyorsa None."""
    try:
        return convert(getattr(definition, prop).Value)
    except pywintypes.com_error:
        return None


def _joint_summary(j) -> dict:
    d = j.Definition
    entry = {
        "name": j.Name,
        "type": JOINT_TYPE_NAMES.get(d.JointType, str(d.JointType)),
        "occurrence_one": j.OccurrenceOne.Name if j.OccurrenceOne else None,
        "occurrence_two": j.OccurrenceTwo.Name if j.OccurrenceTwo else None,
        "suppressed": bool(j.Suppressed),
        "health": HEALTH_NAMES.get(j.HealthStatus, str(j.HealthStatus)),
        "flip_alignment": bool(d.FlipAlignmentDirection),
        "flip_origin": bool(d.FlipOriginDirection),
    }
    angle = _param_value(d, "AngularPosition", rad_to_deg)
    if angle is not None:
        entry["angle_deg"] = angle
        entry["angle_limits_deg"] = (
            [
                _param_value(d, "AngularPositionStartLimit", rad_to_deg),
                _param_value(d, "AngularPositionEndLimit", rad_to_deg),
            ]
            if d.HasAngularPositionLimits
            else None
        )
    linear = _param_value(d, "LinearPosition", cm_to_mm)
    if linear is not None:
        entry["linear_mm"] = linear
        has_start, has_end = d.HasLinearPositionStartLimit, d.HasLinearPositionEndLimit
        entry["linear_limits_mm"] = (
            [
                _param_value(d, "LinearPositionStartLimit", cm_to_mm) if has_start else None,
                _param_value(d, "LinearPositionEndLimit", cm_to_mm) if has_end else None,
            ]
            if has_start or has_end
            else None
        )
    return entry


def _get_joint(comp_def, name: str):
    try:
        return comp_def.Joints.Item(name)
    except pywintypes.com_error:
        raise LookupError(f"'{name}' adli joint bulunamadi.")


def _validate_joint_args(kind: str, angle_limits_deg, linear_limits_mm) -> None:
    if kind not in JOINT_TYPES:
        raise LookupError(f"Gecersiz joint_type: '{kind}'. Gecerli degerler: {', '.join(JOINT_TYPES)}.")
    if angle_limits_deg is not None:
        if kind not in _ANGULAR_TYPES:
            raise LookupError(f"'{kind}' joint turu aci limitlerini desteklemiyor (rotational, cylindrical).")
        if len(angle_limits_deg) != 2:
            raise LookupError("angle_limits_deg [baslangic, bitis] seklinde 2 elemanli olmali.")
    if linear_limits_mm is not None:
        if kind not in _LINEAR_TYPES:
            raise LookupError(f"'{kind}' joint turu dogrusal limitleri desteklemiyor (slider, cylindrical).")
        if len(linear_limits_mm) != 2:
            raise LookupError("linear_limits_mm [baslangic, bitis] seklinde 2 elemanli olmali.")


def _create_joint(
    comp_def, kind, ent1, ent2, point_one, point_two, flip_alignment, flip_origin,
    angle_limits_deg=None, linear_limits_mm=None, angle_position_deg=None,
):
    """Joint tanimini olusturup ekler (document.Update cagirmaz)."""
    intent1 = comp_def.CreateGeometryIntent(ent1, _point_intent(ent1, point_one))
    intent2 = comp_def.CreateGeometryIntent(ent2, _point_intent(ent2, point_two))
    joints = comp_def.Joints
    jd = joints.CreateAssemblyJointDefinition(JOINT_TYPES[kind], intent1, intent2)
    if flip_alignment:
        jd.FlipAlignmentDirection = True
    if flip_origin:
        jd.FlipOriginDirection = True

    # Yeni tanimda pozisyon/limit degerleri dahili birimlerde (rad, cm) sayi olarak atanir.
    if angle_position_deg is not None:
        jd.AngularPosition = deg_to_rad(angle_position_deg)
    if angle_limits_deg is not None:
        jd.HasAngularPositionLimits = True
        jd.AngularPositionStartLimit = deg_to_rad(angle_limits_deg[0])
        jd.AngularPositionEndLimit = deg_to_rad(angle_limits_deg[1])
    if linear_limits_mm is not None:
        start, end = linear_limits_mm
        if start is not None:
            jd.HasLinearPositionStartLimit = True
            jd.LinearPositionStartLimit = mm_to_cm(start)
        if end is not None:
            jd.HasLinearPositionEndLimit = True
            jd.LinearPositionEndLimit = mm_to_cm(end)
    return joints.Add(jd)


# preserve_position toleranslari: konum 0.001 mm, rotasyon matris hucresi basina 1e-6.
_POS_TOL_MM = 0.001
_ROT_TOL = 1e-6
_FLIP_COMBOS = ((False, False), (True, False), (False, True), (True, True))


def _containing_occurrence(entity):
    """Montaj uzayindaki bir kenar/yuz proxy'sinin ait oldugu occurrence'i dondurur."""
    if entity.Type == constants.kEdgeProxyObject:
        return CastTo(entity, "EdgeProxy").ContainingOccurrence
    if entity.Type == constants.kFaceProxyObject:
        return CastTo(entity, "FaceProxy").ContainingOccurrence
    return None


def _snapshot(occurrences) -> list:
    return [(occ, occ.Transformation) for occ in occurrences if occ is not None]


def _residual(snapshot) -> tuple[float, float, bool]:
    """Snapshot'tan bu yana en cok kayan occurrence icin (konum mm, aci derece, tolerans icinde mi)."""
    pos_mm, rot_deg, ok = 0.0, 0.0, True
    for occ, before in snapshot:
        after = occ.Transformation
        tb, ta = before.Translation, after.Translation
        dist_mm = cm_to_mm(math.dist((tb.X, tb.Y, tb.Z), (ta.X, ta.Y, ta.Z)))
        diffs = [before.Cell(r, c) - after.Cell(r, c) for r in range(1, 4) for c in range(1, 4)]
        # ||R1 - R2||_F = 2*sqrt(2)*sin(theta/2)
        fro = math.sqrt(sum(d * d for d in diffs))
        angle = math.degrees(2 * math.asin(min(1.0, fro / (2 * math.sqrt(2)))))
        pos_mm, rot_deg = max(pos_mm, dist_mm), max(rot_deg, angle)
        ok = ok and dist_mm <= _POS_TOL_MM and max(abs(d) for d in diffs) <= _ROT_TOL
    return pos_mm, rot_deg, ok


def _restore(snapshot) -> None:
    for occ, before in snapshot:
        occ.Transformation = before


def _add_joint(
    document, comp_def, kind, ent1, ent2, point_one="center", point_two="center",
    flip_alignment=False, flip_origin=False, angle_limits_deg=None, linear_limits_mm=None,
    name=None, preserve_position=True,
):
    """Joint ekler; preserve_position ise occurrence'larin konumunu korumaya calisir.

    (joint, preserve_bilgisi | None) dondurur. Rigid joint parcayi kaydirirsa diger
    flip kombinasyonlari denenir; hicbiri tutmazsa (orn. dairesel kenarlarda duzlem ici
    aci farki) - rigid tanimda aci secenegi olmadigi icin (AngularPosition E_FAIL,
    AlignmentOne/Two aci degil geometri referansi) - ayni orijinlerde rotational joint
    olusturulur, eski konumu veren aciya surulur ve [a, a] limitleriyle kilitlenir.
    """
    if not preserve_position:
        j = _create_joint(comp_def, kind, ent1, ent2, point_one, point_two, flip_alignment, flip_origin,
                          angle_limits_deg, linear_limits_mm)
        if name:
            j.Name = name
        document.Update()
        return j, None

    snap = _snapshot([_containing_occurrence(ent1), _containing_occurrence(ent2)])
    attempts = 0

    def attempt(k, fa, fo, **kwargs):
        nonlocal attempts
        attempts += 1
        j = _create_joint(comp_def, k, ent1, ent2, point_one, point_two, fa, fo, **kwargs)
        document.Update()
        return j, _residual(snap)

    def finish(j, final_type, fa, fo, res, fallback=None, locked_angle=None):
        info = {
            "preserved": res[2],
            "requested_type": kind,
            "final_type": final_type,
            "flip_alignment": fa,
            "flip_origin": fo,
            "fallback": fallback,
            "attempts": attempts,
            "residual_mm": res[0],
            "residual_deg": res[1],
        }
        if locked_angle is not None:
            info["locked_angle_deg"] = locked_angle
        return j, info

    j, res = attempt(kind, flip_alignment, flip_origin,
                     angle_limits_deg=angle_limits_deg, linear_limits_mm=linear_limits_mm)
    if res[2] or kind != "rigid":
        if name:
            j.Name = name
        return finish(j, kind, flip_alignment, flip_origin, res)

    best = (res[0], res[1], flip_alignment, flip_origin)
    j.Delete()
    _restore(snap)
    for fa, fo in _FLIP_COMBOS:
        if (fa, fo) == (flip_alignment, flip_origin):
            continue
        j, res = attempt("rigid", fa, fo)
        if res[2]:
            if name:
                j.Name = name
            return finish(j, "rigid", fa, fo, res)
        best = min(best, (res[0], res[1], fa, fo))
        j.Delete()
        _restore(snap)

    _, _, bfa, bfo = best
    if best[0] <= _POS_TOL_MM:
        # Sadece eksen etrafinda aci farki kaldi: rotational joint'i dogru aciya sur.
        j, res = attempt("rotational", bfa, bfo)
        found = None
        if res[0] <= _POS_TOL_MM:
            for a in (res[1], -res[1]):
                j.Definition.AngularPosition.Value = deg_to_rad(a)
                document.Update()
                if _residual(snap)[2]:
                    found = a
                    break
        j.Delete()
        _restore(snap)
        if found is not None:
            j, res = attempt("rotational", bfa, bfo, angle_limits_deg=[found, found], angle_position_deg=found)
            j.Name = f"{name or j.Name}_locked"
            return finish(j, "rotational", bfa, bfo, res, fallback="rotational_locked", locked_angle=found)

    # Konum korunamadi: en az kaydiran rigid kombinasyonu birak ve kalan hatayi raporla.
    j, res = attempt("rigid", bfa, bfo)
    if name:
        j.Name = name
    return finish(j, "rigid", bfa, bfo, res)


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_add_joint(
        joint_type: str,
        entity_one: str,
        entity_two: str,
        point_one: str = "center",
        point_two: str = "center",
        flip_alignment: bool = False,
        flip_origin: bool = False,
        angle_limits_deg: list[float] | None = None,
        linear_limits_mm: list[float | None] | None = None,
        preserve_position: bool = True,
        name: str | None = None,
        doc: str | None = None,
    ) -> dict:
        """Iki occurrence arasina montaj eklemi (joint) ekler.

        joint_type: 'rigid', 'rotational', 'slider', 'cylindrical', 'planar', 'ball'.
        entity_one/entity_two: `inventor_assembly_list_occurrence_edges` / `_faces` ile
        alinan entity_id'ler (dairesel/dogrusal kenar, silindirik/konik/duzlemsel yuz).
        Bu oturumda listelenmemis id'ler de occurrence govdesinden yeniden cozulur.
        Inventor entity_one'in occurrence'ini entity_two'ya tasir: hareketli parcayi
        entity_one, sabit (grounded) parcayi entity_two olarak verin.
        point_one/point_two: joint orijini - 'start', 'mid', 'end', 'center'. Kenarlarda
        kenarin noktasi; silindirik/konik yuzeylerde eksenin baslangic/orta/bitis noktasi
        ('center' = orta); duzlemsel yuzeylerde sadece 'center' (yuz merkezi).
        flip_alignment / flip_origin: hizalama / orijin yonunu ters cevirir.
        angle_limits_deg: [baslangic, bitis] aci limitleri (derece) - rotational, cylindrical.
        linear_limits_mm: [baslangic, bitis] dogrusal limitler (mm) - slider, cylindrical;
        tek tarafli limit icin digerini null verin.
        preserve_position: True ise (varsayilan) occurrence'larin mevcut konumu korunmaya
        calisilir. Rigid joint parcayi kaydirirsa diger flip_alignment/flip_origin
        kombinasyonlari denenir; hicbiri tutmazsa (orn. dairesel kenarlarda duzlem ici aci
        farki) ayni orijinlerde rotational joint olusturulup eski aciya kilitlenir
        (limitler [a, a], adi '<ad>_locked'). Konum hic korunamiyorsa en az kaydiran
        flip kombinasyonu kalir - bu, istenen flip_* degerlerini degistirebilir; flip'leri
        aynen uygulamak icin preserve_position=False verin. Diger turlerde sadece kayma
        raporlanir.
        Sonuctaki 'preserve_position' alani: final_type, kullanilan flip'ler, fallback,
        kalan hata (residual_mm / residual_deg).
        name: opsiyonel joint adi (verilmezse Inventor 'Rotational:1' gibi ad verir).
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)

        kind = joint_type.lower()
        _validate_joint_args(kind, angle_limits_deg, linear_limits_mm)
        ent1 = resolve_entity(document, comp_def, entity_one)
        ent2 = resolve_entity(document, comp_def, entity_two)
        j, info = _add_joint(
            document, comp_def, kind, ent1, ent2, point_one, point_two, flip_alignment, flip_origin,
            angle_limits_deg, linear_limits_mm, name, preserve_position,
        )
        entry = _joint_summary(j)
        if info is not None:
            entry["preserve_position"] = info
        return entry

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_list_joints(doc: str | None = None) -> list[dict]:
        """Montajdaki tum joint'leri listeler: ad, tur, occurrence'lar, suppressed, saglik
        durumu, flip durumu ve (tura gore) aci/dogrusal pozisyon ve limitler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        joints = comp_def.Joints
        return [_joint_summary(joints.Item(i)) for i in range(1, joints.Count + 1)]

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_delete_joint(name: str, doc: str | None = None) -> dict:
        """Adi verilen joint'i siler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        _get_joint(comp_def, name).Delete()
        return {"deleted": name}

    @mcp.tool()
    @inventor_tool
    async def inventor_assembly_drive_joint(
        name: str,
        angle_deg: float | None = None,
        linear_mm: float | None = None,
        doc: str | None = None,
    ) -> dict:
        """Bir joint'in pozisyonunu ayarlar (hareketi test etmek icin).

        angle_deg: aci pozisyonu (derece) - rotational, cylindrical.
        linear_mm: dogrusal pozisyon (mm) - slider, cylindrical.
        Limit disi bir deger verilirse joint'in health degeri 'invalid_limits' olur.
        """
        if angle_deg is None and linear_mm is None:
            raise LookupError("angle_deg veya linear_mm'den en az biri verilmeli.")
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_assembly_component_def(document)
        j = _get_joint(comp_def, name)
        d = j.Definition
        kind = JOINT_TYPE_NAMES.get(d.JointType, str(d.JointType))
        if angle_deg is not None:
            try:
                d.AngularPosition.Value = deg_to_rad(angle_deg)
            except pywintypes.com_error:
                raise LookupError(f"'{kind}' joint'i aci pozisyonu desteklemiyor (rotational, cylindrical).")
        if linear_mm is not None:
            try:
                d.LinearPosition.Value = mm_to_cm(linear_mm)
            except pywintypes.com_error:
                raise LookupError(f"'{kind}' joint'i dogrusal pozisyon desteklemiyor (slider, cylindrical).")
        document.Update()
        return _joint_summary(j)
