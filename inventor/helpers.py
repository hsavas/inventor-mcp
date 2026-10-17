"""Ortak yardimcilar: dokuman/sketch/feature lookup, hata sarmalama, transient geometry."""

import functools
import json
import os
import threading
import time

import pywintypes
from win32com.client import CastTo, constants

# Inventor COM nesnesi tek-threaded (STA). FastMCP eszamanli/paralel tool
# cagrilarini farkli thread'lerde calistirabilir; bu, "yanlis dokuman
# kapandi" gibi yarış (race condition) hatalarina yol acabilir (gercekten
# yasandi - bkz. inventor_close_document olayı). Bu kilit, ayni anda
# sadece BIR Inventor COM cagrisinin yapilmasini garanti eder.
_com_lock = threading.Lock()

from inventor.units import mm_to_cm

DOC_TYPE_NAMES = {
    constants.kPartDocumentObject: "part",
    constants.kAssemblyDocumentObject: "assembly",
    constants.kDrawingDocumentObject: "drawing",
    constants.kPresentationDocumentObject: "presentation",
}

DOC_TYPE_FROM_NAME = {
    "part": constants.kPartDocumentObject,
    "assembly": constants.kAssemblyDocumentObject,
    "drawing": constants.kDrawingDocumentObject,
}


def format_com_error(e: pywintypes.com_error) -> str:
    """pywintypes.com_error nesnesini okunabilir bir mesaja cevirir."""
    hresult = e.args[0]
    msg = e.args[1] if len(e.args) > 1 else ""
    excepinfo = e.args[2] if len(e.args) > 2 else None
    detail = ""
    if excepinfo:
        source = excepinfo[1] if len(excepinfo) > 1 else None
        description = excepinfo[2] if len(excepinfo) > 2 else None
        if description:
            detail = f" -- {source}: {description}" if source else f" -- {description}"
    return f"Inventor COM hatasi 0x{hresult & 0xFFFFFFFF:08X}: {msg}{detail}"


_TRACE_PATH = os.environ.get("INVENTOR_MCP_TRACE")  # JSONL dosya yolu; bossa kayit yok


def _trace(tool: str, kwargs: dict, ok: bool, payload, started: float) -> None:
    """Her tool cagrisini (egitim verisi / hata analizi icin) JSONL'e yazar."""
    if not _TRACE_PATH:
        return
    try:
        text = json.dumps(payload, ensure_ascii=False, default=str)
        rec = {
            "ts": time.time(),
            "session": os.environ.get("INVENTOR_MCP_SESSION", ""),
            "tool": tool,
            # Buyuk argumanlar (orn. inventor_run_script betikleri) logu sisirmesin.
            "args": {
                k: (v[:2000] + f"...<{len(v)} karakter>" if isinstance(v, str) and len(v) > 2000 else v)
                for k, v in kwargs.items()
            },
            "ok": ok,
            "result": text[:4000],
            "ms": round((time.time() - started) * 1000, 1),
        }
        with open(_TRACE_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
    except Exception:
        pass  # kayit hatasi asla tool'u bozmamali


def inventor_tool(func):
    """COM ve lookup hatalarini okunabilir RuntimeError'a ceviren decorator.

    INVENTOR_MCP_TRACE ortam degiskeni tanimliysa her cagri JSONL olarak kaydedilir.
    """

    @functools.wraps(func)
    async def wrapper(*args, **kwargs):
        started = time.time()
        with _com_lock:
            try:
                result = await func(*args, **kwargs)
            except pywintypes.com_error as e:
                msg = format_com_error(e)
                _trace(func.__name__, kwargs, False, msg, started)
                raise RuntimeError(msg) from None
            except (LookupError, ValueError) as e:
                _trace(func.__name__, kwargs, False, str(e), started)
                raise RuntimeError(str(e)) from None
            except Exception as e:
                _trace(func.__name__, kwargs, False, f"{type(e).__name__}: {e}", started)
                raise
            _trace(func.__name__, kwargs, True, result, started)
            return result

    return wrapper


def doc_type_name(doc) -> str:
    return DOC_TYPE_NAMES.get(doc.DocumentType, str(doc.DocumentType))


def get_document(app, name: str | None = None):
    """Doküman adına veya tam dosya yoluna göre açık bir dokümanı bulur.

    name verilmezse aktif dokümanı döndürür.
    """
    if name is None:
        doc = app.ActiveDocument
        if doc is None:
            raise LookupError("Acik bir dokuman yok ve 'name' belirtilmedi.")
        return doc

    for d in app.Documents:
        if d.DisplayName == name or d.FullFileName == name:
            return d
    raise LookupError(f"'{name}' adli acik bir dokuman bulunamadi.")


def get_part_component_def(doc):
    """Bir Part dokumaninin ComponentDefinition'ini dondurur.

    `Documents.Add`/`Documents.Item` gibi genel donus tipli COM cagrilari erken-bagli
    generic 'Document' arayuzunu dondurur; bu arayuzde 'ComponentDefinition' yoktur.
    Bu yuzden somut 'PartDocument' arayuzune CastTo ile gecis yapilir.
    """
    if doc.DocumentType != constants.kPartDocumentObject:
        raise LookupError(
            f"Bu islem sadece Part (.ipt) dokumanlarinda calisir, '{doc.DisplayName}' bir {doc_type_name(doc)} dokumani."
        )
    part_doc = CastTo(doc, "PartDocument")
    return part_doc.ComponentDefinition


def get_assembly_component_def(doc):
    """Bir Assembly dokumaninin AssemblyComponentDefinition'ini dondurur."""
    if doc.DocumentType != constants.kAssemblyDocumentObject:
        raise LookupError(
            f"Bu islem sadece Assembly (.iam) dokumanlarinda calisir, '{doc.DisplayName}' bir {doc_type_name(doc)} dokumani."
        )
    asm_doc = CastTo(doc, "AssemblyDocument")
    return asm_doc.ComponentDefinition


def get_drawing_document(doc):
    """Bir Drawing dokumanini somut 'DrawingDocument' arayuzune cevirir."""
    if doc.DocumentType != constants.kDrawingDocumentObject:
        raise LookupError(
            f"Bu islem sadece Drawing (.idw/.dwg) dokumanlarinda calisir, '{doc.DisplayName}' bir {doc_type_name(doc)} dokumani."
        )
    return CastTo(doc, "DrawingDocument")


def get_sheet(dwg_doc, name: str | None = None):
    """Bir Drawing dokumanindaki sheet'i adina gore bulur; verilmezse aktif sheet'i dondurur."""
    if name is None:
        return dwg_doc.ActiveSheet
    try:
        return dwg_doc.Sheets.Item(name)
    except pywintypes.com_error:
        raise LookupError(f"'{name}' adli sheet bulunamadi.")


def get_sketch(doc, name: str):
    """Part dokumanindaki bir PlanarSketch'i adina gore bulur."""
    comp_def = get_part_component_def(doc)
    for sk in comp_def.Sketches:
        if sk.Name == name:
            return sk
    raise LookupError(f"'{name}' adli sketch bulunamadi.")


def count_closed_regions(sk) -> int:
    """Sketch'teki kapali bolge (profil yolu) sayisini dondurur.

    Profiles.Count sadece daha once olusturulmus Profile nesnelerini sayar
    (yeni sketch'te hep 0). Bu yuzden gecici bir solid profili olusturup yol
    sayisina bakiyor, sonra o gecici profili siliyoruz ki her cagri sketch'te
    artik bir Profile nesnesi birakmasin.
    """
    try:
        probe = sk.Profiles.AddForSolid()
    except pywintypes.com_error:
        return 0
    try:
        return probe.Count
    finally:
        try:
            probe.Delete()
        except (pywintypes.com_error, AttributeError):
            pass


def get_sketch3d(doc, name: str):
    """Part dokumanindaki bir Sketch3D'yi adina gore bulur."""
    comp_def = get_part_component_def(doc)
    for sk in comp_def.Sketches3D:
        if sk.Name == name:
            return sk
    raise LookupError(f"'{name}' adli 3D sketch bulunamadi.")


def get_feature(doc, name: str):
    """Part dokumanindaki bir feature'i adina gore bulur."""
    comp_def = get_part_component_def(doc)
    try:
        return comp_def.Features.Item(name)
    except pywintypes.com_error:
        raise LookupError(f"'{name}' adli feature bulunamadi.")


def point2d(app, x_mm: float, y_mm: float):
    """mm cinsinden 2D nokta olusturur (Inventor icinde cm'e cevrilir)."""
    tg = app.TransientGeometry
    return tg.CreatePoint2d(mm_to_cm(x_mm), mm_to_cm(y_mm))


def point3d(app, x_mm: float, y_mm: float, z_mm: float):
    """mm cinsinden 3D nokta olusturur (Inventor icinde cm'e cevrilir)."""
    tg = app.TransientGeometry
    return tg.CreatePoint(mm_to_cm(x_mm), mm_to_cm(y_mm), mm_to_cm(z_mm))
