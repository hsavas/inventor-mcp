"""STEP/IGES/STL/SAT/DWG/PDF disa aktarim araclari.

Tum disa aktarimlar Inventor'un SilentOperation modunda yapilir: aksi halde bazi
ceviriciler (orn. DWG) Inventor icinde modal bir dialog acar ve MCP sunucusu
kullanici dialogu kapatana kadar kilitlenir.
"""

import glob
import os
import re
import tempfile
import time

from win32com.client import CastTo, constants

from inventor.connection import get_app
from inventor.helpers import get_document, inventor_tool

TRANSLATOR_IDS = {
    "step": "{90AF7F40-0C01-11D5-8E83-0010B541CD80}",
    "iges": "{90AF7F44-0C01-11D5-8E83-0010B541CD80}",
    "stl": "{533E9A98-FC3B-11D4-8E7E-0010B541CD80}",
    "sat": "{89162634-02B6-11D5-8E80-0010B541CD80}",
    "dwg": "{C24E3AC2-122E-11D5-8E91-0010B541CD80}",
    "pdf": "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}",
}

# Inventor kurulumundaki exportdwg.ini bulunamazsa kullanilan asgari ayarlar.
_MIN_DWG_INI = """[EXPORT SELECT OPTIONS]
AUTOCAD VERSION=AutoCAD 2018
CREATE AUTOCAD MECHANICAL=No
USE TRANSMITTAL=No
USE CUSTOMIZE=No
CREATE LAYER GROUP=No
PARTS ONLY=No
[EXPORT PROPERTIES]
SELECTED PROPERTIES=
[EXPORT DESTINATION]
SPACE=Model
SCALING=Geometry
MAPPING=MapsBest
MODEL GEOMETRY ONLY=No
SHEET CONTENT ONLY=No
EXPLODE DIMENSIONS=No
SYMBOLS ARE BLOCKED=Yes
AUTOCAD TEMPLATE=
DESTINATION DXF=No
"""


def _dwg_ini(app) -> str:
    """Inventor'un kendi exportdwg.ini'sinin, ZIP paketlemesi (USE TRANSMITTAL)
    kapatilmis bir kopyasini olusturur ve yolunu dondurur.

    Varsayilan ayar 'USE TRANSMITTAL=Yes' oldugunda cikti .dwg yerine .zip olur."""
    text = None
    try:
        src = os.path.join(app.FileLocations.DesignDataPath, "DWG-DXF", "exportdwg.ini")
        with open(src, encoding="mbcs", errors="replace") as f:
            text = f.read()
    except Exception:
        pass
    if not text:
        text = _MIN_DWG_INI
    text = re.sub(r"(?im)^USE TRANSMITTAL=.*$", "USE TRANSMITTAL=No", text)
    dst = os.path.join(tempfile.gettempdir(), "inventor_mcp_exportdwg.ini")
    with open(dst, "w", encoding="mbcs", errors="replace") as f:
        f.write(text)
    return dst


def _written_files(path: str, started: float) -> list[str]:
    """Istenen yol veya (cok sayfali DWG gibi) ayni adla baslayan yeni dosyalar."""
    if os.path.exists(path):
        return [path]
    stem, ext = os.path.splitext(path)
    return sorted(p for p in glob.glob(glob.escape(stem) + "*" + ext) if os.path.getmtime(p) >= started - 1)


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_export(format: str, path: str, doc: str | None = None) -> dict:
        """Dokumani belirtilen formatta disa aktarir (Save Copy As).

        format: 'step', 'iges', 'stl', 'sat', 'dwg', 'pdf'.
        path: cikti dosyasinin tam yolu. Cok sayfali bir drawing DWG'ye aktarilirsa
        Inventor her sayfa icin ayri dosya yazabilir; hepsi 'files' alaninda doner.
        Inventor hicbir dialog gostermez (SilentOperation).
        """
        app = get_app()
        document = get_document(app, doc)

        fmt = format.lower()
        clsid = TRANSLATOR_IDS.get(fmt)
        if clsid is None:
            raise LookupError(f"Gecersiz format: '{format}'. Gecerli degerler: {', '.join(TRANSLATOR_IDS)}.")
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)

        addin = app.ApplicationAddIns.ItemById(clsid)
        # Bazi ceviriciler (orn. DWG) Inventor acilisinda yuklenmez; yuklenmemis bir
        # translator'da SaveCopyAs genel bir COM hatasi (0x80020009) verir.
        if not addin.Activated:
            addin.Activate()
        translator = CastTo(addin, "TranslatorAddIn")

        context = app.TransientObjects.CreateTranslationContext()
        context.Type = constants.kFileBrowseIOMechanism
        options = app.TransientObjects.CreateNameValueMap()
        data_medium = app.TransientObjects.CreateDataMedium()
        data_medium.FileName = path

        started = time.time()
        previous_silent = app.SilentOperation
        app.SilentOperation = True
        try:
            translator.HasSaveCopyAsOptions(document, context, options)
            if fmt == "dwg":
                options.SetValue("Export_Acad_IniFile", _dwg_ini(app))
            translator.SaveCopyAs(document, context, options, data_medium)
        finally:
            app.SilentOperation = previous_silent

        files = _written_files(path, started)
        if not files:
            raise RuntimeError(f"{fmt} disa aktarimi hata vermedi ama '{path}' olusmadi.")
        return {"format": fmt, "path": files[0], "files": files, "exists": True,
                "size_bytes": sum(os.path.getsize(p) for p in files)}
