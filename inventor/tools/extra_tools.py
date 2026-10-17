"""Ek araclar: oturum sifirlama, model saglik kontrolu, sac metal ve dis (thread).

Forumlarda en cok sorulan ama MCP'de eksik olan islemler (sac metal flans/
acinim, standart dis) ile yerel model egitimi icin gereken yardimci araclar
(tum dokumanlari kapatma, modelin saglikli olup olmadigini dogrulama).
"""

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
)
from inventor.units import cm_to_mm, cm3_to_mm3, mm_to_cm

SHEET_METAL_SUBTYPE = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
PART_SUBTYPE = "{4D29B490-49B2-11D0-93C3-7E0706000000}"

# Inventor HealthStatusEnum degerleri, surume gore constants'tan okunur
# (or. kUpToDateHealth -> "up_to_date", kInErrorHealth -> "in_error").
def _health_names() -> dict:
    import re

    out = {}
    for d in constants.__dicts__:
        for n, v in d.items():
            if n.startswith("k") and n.endswith("Health"):
                out[v] = re.sub(r"(?<!^)(?=[A-Z])", "_", n[1:-6]).lower()
    return out


_HEALTH = _health_names()
_OK_HEALTH = {"up_to_date", "suppressed", "beyond_stop_node", "newly_added"}


def _dyn(obj):
    return win32com.client.dynamic.Dispatch(obj._oleobj_)


def _sheet_metal_def(document):
    if document.DocumentType != constants.kPartDocumentObject:
        raise LookupError("Sac metal islemleri sadece Part (.ipt) dokumanlarinda yapilir.")
    if document.SubType != SHEET_METAL_SUBTYPE:
        raise LookupError(
            "Dokuman sac metal degil. Once 'inventor_sheet_metal_convert' ile donusturun."
        )
    return _dyn(CastTo(document, "PartDocument").ComponentDefinition)


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_close_all_documents(save: bool = False) -> dict:
        """Acik TUM dokumanlari kapatir (temiz bir calisma ortami icin).
        save=False ise kaydedilmemis degisiklikler atilir."""
        app = get_app()
        names = [d.DisplayName for d in app.Documents]
        if save:
            for d in list(app.Documents):
                if d.Dirty and d.FullFileName:
                    d.Save()
        app.Documents.CloseAll(False)
        return {"closed": names, "count": len(names)}

    @mcp.tool()
    @inventor_tool
    async def inventor_check_model(doc: str | None = None) -> dict:
        """Parcanin saglik raporunu verir: govde sayisi, toplam hacim, sinir kutusu,
        hatali/uyarili feature'lar ve kapali profili olmayan sketch'ler.
        Bir modelleme isini bitirdiginizi beyan etmeden once bunu cagirin."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        try:
            document.Update()
        except pywintypes.com_error:
            pass

        bodies = comp_def.SurfaceBodies
        volume = 0.0
        size = None
        if bodies.Count:
            mins, maxs = None, None
            for b in bodies:
                rb = b.RangeBox
                bmin = (rb.MinPoint.X, rb.MinPoint.Y, rb.MinPoint.Z)
                bmax = (rb.MaxPoint.X, rb.MaxPoint.Y, rb.MaxPoint.Z)
                mins = bmin if mins is None else tuple(map(min, mins, bmin))
                maxs = bmax if maxs is None else tuple(map(max, maxs, bmax))
            size = [round(cm_to_mm(b - a), 4) for a, b in zip(mins, maxs)]
            try:
                volume = comp_def.MassProperties.Volume
            except pywintypes.com_error:
                pass

        problems = []
        for f in comp_def.Features:
            try:
                hs = f.HealthStatus
            except pywintypes.com_error:
                continue
            name = _HEALTH.get(hs, str(hs))
            if name not in _OK_HEALTH:
                problems.append({"feature": f.Name, "health": name})

        open_sketches = []
        for sk in comp_def.Sketches:
            if sk.Consumed:
                continue
            n = count_closed_regions(sk)
            if n == 0 and sk.SketchEntities.Count > 0:
                open_sketches.append(sk.Name)

        return {
            "document": document.DisplayName,
            "is_sheet_metal": document.SubType == SHEET_METAL_SUBTYPE,
            "body_count": bodies.Count,
            "volume_mm3": round(cm3_to_mm3(volume), 4),
            "size_mm": size,
            "feature_count": comp_def.Features.Count,
            "problem_features": problems,
            "unconsumed_sketches_without_closed_profile": open_sketches,
            "healthy": bodies.Count > 0 and not problems,
        }

    # ------------------------------------------------------------------ sac metal
    @mcp.tool()
    @inventor_tool
    async def inventor_sheet_metal_convert(thickness_mm: float, doc: str | None = None) -> dict:
        """Aktif parcayi sac metal parcaya donusturur ve sac kalinligini ayarlar.
        Bos bir parcada da cagrilabilir (sonra 'inventor_sheet_metal_face' ile baslanir)."""
        app = get_app()
        document = get_document(app, doc)
        if document.DocumentType != constants.kPartDocumentObject:
            raise LookupError("Sadece Part dokumani sac metale donusturulebilir.")
        if document.SubType != SHEET_METAL_SUBTYPE:
            document.SubType = SHEET_METAL_SUBTYPE
        sm = _sheet_metal_def(document)
        sm.UseSheetMetalStyleThickness = False
        sm.Thickness.Expression = f"{thickness_mm} mm"
        return {"document": document.DisplayName, "sheet_metal": True, "thickness_mm": thickness_mm}

    @mcp.tool()
    @inventor_tool
    async def inventor_sheet_metal_face(sketch: str, doc: str | None = None) -> dict:
        """Kapali bir sketch profilinden sac metal 'Face' (duz sac) olusturur.
        Kalinlik sac metal kurallarindan gelir (inventor_sheet_metal_convert)."""
        app = get_app()
        document = get_document(app, doc)
        sm = _sheet_metal_def(document)
        sk = get_sketch(document, sketch)
        profile = sk.Profiles.AddForSolid()
        faces = sm.Features.FaceFeatures
        fdef = faces.CreateFaceFeatureDefinition(profile)
        feat = faces.Add(fdef)
        return {"name": feat.Name, "type": "sheet_metal_face"}

    @mcp.tool()
    @inventor_tool
    async def inventor_sheet_metal_flange(
        edge_refs: list[str], height_mm: float, angle_deg: float = 90.0, doc: str | None = None
    ) -> dict:
        """Secilen dogrusal kenar(lar)a flans (bukum + duvar) ekler.
        edge_refs: 'inventor_list_edges' ile alinan entity_id'ler (sacin ust/alt yuz kenarlari)."""
        app = get_app()
        document = get_document(app, doc)
        sm = _sheet_metal_def(document)
        edges = app.TransientObjects.CreateEdgeCollection()
        for ref in edge_refs:
            edges.Add(registry.resolve(ref))
        flanges = sm.Features.FlangeFeatures
        fdef = flanges.CreateFlangeDefinition(edges, f"{angle_deg} deg", f"{height_mm} mm")
        feat = flanges.Add(fdef)
        return {"name": feat.Name, "type": "flange", "height_mm": height_mm, "angle_deg": angle_deg}

    @mcp.tool()
    @inventor_tool
    async def inventor_sheet_metal_flat_pattern(doc: str | None = None) -> dict:
        """Sac metal parcanin acinimini (flat pattern) olusturur/gunceller ve
        acinim olculerini (mm) dondurur. Sonra katlanmis gorunume geri doner."""
        app = get_app()
        document = get_document(app, doc)
        sm = _sheet_metal_def(document)
        if not sm.HasFlatPattern:
            sm.Unfold()
        else:
            sm.FlatPattern.Edit()
        fp = sm.FlatPattern
        result = {"length_mm": round(cm_to_mm(fp.Length), 4), "width_mm": round(cm_to_mm(fp.Width), 4)}
        try:
            fp.ExitEdit()
        except pywintypes.com_error:
            pass
        return result

    # ------------------------------------------------------------------ dis
    @mcp.tool()
    @inventor_tool
    async def inventor_thread(
        face_ref: str,
        designation: str = "M10x1.5",
        internal: bool = False,
        thread_type: str = "ISO Metric profile",
        thread_class: str | None = None,
        full_length: bool = True,
        depth_mm: float | None = None,
        right_handed: bool = True,
        start_edge_ref: str | None = None,
        doc: str | None = None,
    ) -> dict:
        """Silindirik bir yuzeye standart (kozmetik) dis ekler.

        face_ref: 'inventor_list_faces' ile alinan silindir yuzeyin entity_id'si.
        internal: delik (ic dis) ise True, mil (dis dis) ise False.
        designation: ornek 'M10x1.5', 'M6x1'. Cap, yuzey capiyla uyumlu olmali.
        thread_class: verilmezse ic dis icin '6H', dis dis icin '6g'.
        start_edge_ref: disin baslayacagi dairesel kenar ('inventor_list_edges'
          entity_id'si). Verilmezse yuzeyin ilk kenari kullanilir; bu silindirin
          yanlis ucu olabilir, kismi uzunluklu dislerde (full_length=False) verin.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        face = registry.resolve(face_ref)
        cls = thread_class or ("6H" if internal else "6g")
        threads = _dyn(comp_def.Features.ThreadFeatures)
        info = threads.CreateStandardThreadInfo(internal, right_handed, thread_type, designation, cls)
        start_edge = registry.resolve(start_edge_ref) if start_edge_ref else face.Edges.Item(1)
        if full_length:
            feat = threads.Add(face, start_edge, info, False, True)
        else:
            if depth_mm is None:
                raise ValueError("full_length=False ise depth_mm verilmeli.")
            feat = threads.Add(face, start_edge, info, False, False, f"{depth_mm} mm")
        return {"name": feat.Name, "type": "thread", "designation": designation, "class": cls, "internal": internal}
