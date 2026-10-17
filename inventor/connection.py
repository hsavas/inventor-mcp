"""Autodesk Inventor.Application baglanti yonetimi.

NEDEN IKI INVENTOR ACILIYORDU
-----------------------------
* Inventor COM sunucusu "single-use"dir: ProgID uzerinden yapilan HER
  CoCreateInstance (Dispatch / EnsureDispatch / CreateObject) calisan bir
  Inventor olsa bile YENI bir Inventor.exe (/Automation -Embedding) baslatir.
* Otomasyonla baslatilan Inventor kendini Running Object Table'a (ROT)
  kaydetmez, yani GetActiveObject onu bulamaz.
* Eski kod modul import edilirken get_app() -> EnsureDispatch yapiyordu.
  Claude Desktop MCP sunucusunu iki surec olarak baslattiginda her biri
  acilista kendi Inventor'unu aciyordu.

YENI DAVRANIS
-------------
1. Import sirasinda Inventor BASLATILMAZ; sadece tip kutuphanesi (gen_py)
   yuklenir ki win32com.client.constants dolsun.
2. Ilk tool cagrisinda once ROT'taki calisan Inventor'a baglanilir
   (kullanicinin kendi actigi Inventor dahil).
3. Hic yoksa, makine genelinde bir mutex altinda TEK bir Inventor baslatilir
   ve ROT'a biz kaydederiz. Boylece ayni surecin diger thread'leri ve diger
   MCP surecleri ayni Inventor'a baglanir.
"""

import threading

import pythoncom
import pywintypes
import win32api
import win32com.client
import win32com.client.gencache
import win32event

PROG_ID = "Inventor.Application"
# Inventor Object Library tip kutuphanesi (gen_py: D98A091D-...x0x1x0)
TYPELIB_GUID = "{D98A091D-3A0F-4C3E-B36E-61F62068D488}"
TYPELIB_LCID, TYPELIB_MAJOR, TYPELIB_MINOR = 0, 1, 0

MUTEX_NAME = "Local\\inventor_mcp_launch"
MUTEX_TIMEOUT_MS = 180_000  # baska surec Inventor'u acarken en fazla bu kadar bekle

_thread_lock = threading.Lock()
_rot_cookie = None  # biz kaydettiysek ROT kaydinin cookie'si
_launched_app = None  # baslattigimiz instance (ROT kaydini canli tutmak icin)


def _co_init() -> None:
    try:
        pythoncom.CoInitialize()
    except pythoncom.com_error:
        pass  # zaten bu thread'de initialize edilmis (S_FALSE)


def _ensure_typelib() -> None:
    """Inventor'u BASLATMADAN gen_py modulunu yukler -> constants dolar."""
    try:
        win32com.client.gencache.EnsureModule(TYPELIB_GUID, TYPELIB_LCID, TYPELIB_MAJOR, TYPELIB_MINOR)
    except Exception:
        # Tip kutuphanesi bulunamazsa constants ilk baglantida (EnsureDispatch) dolar.
        pass


def _get_running():
    """ROT'taki calisan Inventor'a bu thread icin early-bound proxy; yoksa None."""
    try:
        obj = pythoncom.GetActiveObject(PROG_ID)
    except pywintypes.com_error:
        return None
    try:
        disp = obj.QueryInterface(pythoncom.IID_IDispatch)
        app = win32com.client.gencache.EnsureDispatch(disp)
        _ = app.SoftwareVersion  # olu (kapatilmis) kayitlari ele
        return app
    except pywintypes.com_error:
        return None


def _launch_and_register():
    """Yeni bir Inventor baslatir, gorunur yapar ve ROT'a kaydeder."""
    global _rot_cookie, _launched_app

    # DispatchEx her zaman yeni surec acar - burada istenen tam olarak bu.
    raw = win32com.client.DispatchEx(PROG_ID)
    app = win32com.client.gencache.EnsureDispatch(raw._oleobj_)
    app.Visible = True

    _launched_app = app

    # Inventor genelde acilisi bitince kendini ROT'a kaydeder. Kaydetmediyse
    # (otomasyon modu) biz kaydederiz; basarisiz olursa da baglanti calisir,
    # sadece diger sureclerin ayni instance'i bulmasi o kayda baglidir.
    if _get_running() is None:
        if _rot_cookie is not None:
            try:
                pythoncom.RevokeActiveObject(_rot_cookie)
            except pywintypes.com_error:
                pass
            _rot_cookie = None
        try:
            unk = app._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
            _rot_cookie = pythoncom.RegisterActiveObject(
                unk, pywintypes.IID(PROG_ID), pythoncom.ACTIVEOBJECT_WEAK
            )
        except pywintypes.com_error:
            pass
    return app


def get_app():
    """Calisan Inventor.Application nesnesini dondurur; hic yoksa BIR KEZ baslatir.

    FastMCP her tool cagrisini farkli bir thread'de calistirabilir. COM
    pointer'larini thread'ler arasi paylasmiyoruz: her cagrida bu thread icin
    CoInitialize yapip ROT'tan taze bir proxy aliyoruz.
    """
    _co_init()

    app = _get_running()
    if app is not None:
        return app

    with _thread_lock:
        mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
        try:
            win32event.WaitForSingleObject(mutex, MUTEX_TIMEOUT_MS)
            try:
                # Beklerken baska thread/surec baslatip kaydetmis olabilir.
                app = _get_running()
                if app is None:
                    app = _launch_and_register()
                return app
            finally:
                win32event.ReleaseMutex(mutex)
        finally:
            win32api.CloseHandle(mutex)


# Import'ta Inventor BASLATILMAZ. Sadece tip kutuphanesi yuklenir ki diger
# modullerdeki ust-seviye 'from win32com.client import constants' kullanimlari
# (OPERATION_MAP vb.) calissin.
_ensure_typelib()
