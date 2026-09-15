"""Uygulama ve dokuman yonetimi araclari."""

import os

from inventor.connection import get_app
from inventor.helpers import DOC_TYPE_FROM_NAME, doc_type_name, get_document, inventor_tool

# Varsayilan sablon kurulumun diline gore inch tabanli olabilir. Metrik istendiginde
# once sablon klasorunun 'Metric' alt klasorundeki karsiligi aranir.
_METRIC_TEMPLATE_NAMES = {
    "part": "Standard (mm).ipt",
    "assembly": "Standard (mm).iam",
    "drawing": "ISO.idw",
}


def _resolve_template(app, doc_type: str, type_const: str, metric: bool) -> str:
    """Dokuman sablonunu bulur.

    FileManager.GetTemplateFile bir 'SystemsOfMeasure' argumani kabul eder ama
    bazi kurulumlarda (orn. Inventor 2020 en-US) bunu YOK SAYIP her zaman ayni
    inch tabanli Standard.ipt yolunu dondurur. Bu yuzden metrik sablon dosyasi
    dogrudan diskten aranir; bulunamazsa varsayilan sablona dusulur ve birimler
    dokuman olusturulduktan sonra zorlanir.
    """
    default_template = app.FileManager.GetTemplateFile(type_const)
    if not metric:
        return default_template
    name = _METRIC_TEMPLATE_NAMES.get(doc_type)
    if name:
        candidate = os.path.join(os.path.dirname(default_template), "Metric", name)
        if os.path.isfile(candidate):
            return candidate
    return default_template


def _force_metric_units(document) -> None:
    """Dokumanin goruntu birimlerini mm/gram/derece yapar.

    Sablon inch tabanli kaldiysa tek guvenilir duzeltme budur. Birimi
    desteklemeyen dokuman turlerinde (bazi drawing'ler) sessizce gecilir.
    """
    try:
        uom = document.UnitsOfMeasure
    except Exception:
        return
    for attr, unit in (("LengthUnits", "mm"), ("MassUnits", "gram"), ("AngleUnits", "deg")):
        try:
            setattr(uom, attr, uom.GetTypeFromString(unit))
        except Exception:
            pass


def _doc_summary(d) -> dict:
    return {
        "name": d.DisplayName,
        "type": doc_type_name(d),
        "full_path": d.FullFileName,
        "dirty": bool(d.Dirty),
    }


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_get_status() -> dict:
        """Inventor uygulamasinin durumunu dondurur: calisiyor mu, surum, gorunurluk,
        acik dokumanlar ve aktif dokuman."""
        app = get_app()
        active_doc = app.ActiveDocument
        return {
            "running": True,
            "visible": bool(app.Visible),
            "version": app.SoftwareVersion.DisplayVersion,
            "documents": [_doc_summary(d) for d in app.Documents],
            "active_document": active_doc.DisplayName if active_doc else None,
        }

    @mcp.tool()
    @inventor_tool
    async def inventor_list_documents() -> list[dict]:
        """Acik tum Inventor dokumanlarini listeler (ad, tur, dosya yolu, kaydedilmemis degisiklik)."""
        app = get_app()
        return [_doc_summary(d) for d in app.Documents]

    @mcp.tool()
    @inventor_tool
    async def inventor_new_document(doc_type: str = "part", units: str = "metric") -> dict:
        """Yeni bir Inventor dokumani olusturur.

        doc_type: 'part' (.ipt), 'assembly' (.iam) veya 'drawing' (.idw/.dwg).
        units: 'metric' (varsayilan, mm) veya 'english' (inch). Bu MCP'nin tum
        araclari mm ile calistigi icin varsayilan metriktir; kurulumun varsayilan
        sablonu inch tabanli olsa bile birimler mm'ye ayarlanir.
        """
        app = get_app()
        key = doc_type.lower()
        const = DOC_TYPE_FROM_NAME.get(key)
        if const is None:
            raise LookupError(
                f"Gecersiz doc_type: '{doc_type}'. 'part', 'assembly' veya 'drawing' olmali."
            )
        unit_key = units.lower()
        if unit_key not in ("metric", "english"):
            raise LookupError(f"Gecersiz units: '{units}'. 'metric' veya 'english' olmali.")

        metric = unit_key == "metric"
        template = _resolve_template(app, key, const, metric)
        doc = app.Documents.Add(const, template, True)
        if metric:
            _force_metric_units(doc)
        summary = _doc_summary(doc)
        summary["template"] = template
        try:
            uom = doc.UnitsOfMeasure
            summary["length_units"] = uom.GetStringFromType(uom.LengthUnits)
        except Exception:
            pass
        return summary

    @mcp.tool()
    @inventor_tool
    async def inventor_open_document(path: str) -> dict:
        """Var olan bir Inventor dosyasini (.ipt/.iam/.idw/.dwg) acar."""
        app = get_app()
        doc = app.Documents.Open(path, True)
        return _doc_summary(doc)

    @mcp.tool()
    @inventor_tool
    async def inventor_save_document(name: str | None = None, save_as_path: str | None = None) -> dict:
        """Dokumani kaydeder.

        name verilmezse aktif dokuman kaydedilir. save_as_path verilirse o yola
        farkli kaydedilir (Save As).
        """
        app = get_app()
        doc = get_document(app, name)
        if save_as_path:
            doc.SaveAs(save_as_path, False)
        else:
            doc.Save()
        return _doc_summary(doc)

    @mcp.tool()
    @inventor_tool
    async def inventor_close_document(name: str | None = None, save: bool = True) -> dict:
        """Dokumani kapatir. name verilmezse aktif dokuman kapatilir.

        save=True ise kapatmadan once kaydedilmemis degisiklikler kaydedilir.
        """
        app = get_app()
        doc = get_document(app, name)
        display_name = doc.DisplayName
        if save and doc.Dirty:
            doc.Save()
        doc.Close(True)
        return {"closed": display_name}

    @mcp.tool()
    @inventor_tool
    async def inventor_set_active_document(name: str) -> dict:
        """Belirtilen acik dokumani aktif (on plandaki) dokuman yapar."""
        app = get_app()
        doc = get_document(app, name)
        doc.Activate()
        return _doc_summary(doc)
