"""Rastgele Python betigi calistirma araci (Fusion MCP'deki 'script' feature'inin
Inventor esdegeri).

Cok sayida kucuk adimdan olusan (yuzlerce dugum/cubuk gibi) buyuk montajlari,
tek tek MCP tool cagrisi yapmak yerine TEK bir cagrida, dogrudan Inventor COM
API'sini (win32com uzerinden) kullanan bir Python betigiyle insa etmek icin
kullanilir. Betik, bu modulun sagladigi hazir degiskenlere (app, document,
comp_def, constants, tg, mm_to_cm, deg_to_rad, ...) erisebilir.
"""

import contextlib
import io
import math
import os

import pywintypes
import win32com.client.dynamic
from win32com.client import CastTo, constants

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import (
    format_com_error,
    get_document,
    get_part_component_def,
    inventor_tool,
    point2d,
    point3d,
)
from inventor.units import cm_to_mm, cm2_to_mm2, cm3_to_mm3, deg_to_rad, mm_to_cm, rad_to_deg


def register(mcp):
    # GUVENLIK: bu arac MCP istemcisine bu makinede tam Python erisimi verir.
    # Kapatmak icin ortam degiskeni: INVENTOR_MCP_DISABLE_SCRIPT=1
    if os.environ.get("INVENTOR_MCP_DISABLE_SCRIPT", "").strip().lower() in ("1", "true", "yes"):
        return

    @mcp.tool()
    @inventor_tool
    async def inventor_run_script(script: str, doc: str | None = None) -> dict:
        """Verilen Python kodunu, Inventor COM API'sine dogrudan erisimi olan bir
        ortamda calistirir (win32com/pywin32 uzerinden - Fusion MCP'deki
        'script' feature'inin ayni mantikla Inventor karsiligidir).

        Betik, asagidaki hazir degiskenleri kullanabilir:
          - app: Inventor.Application COM nesnesi
          - document: aktif (veya 'doc' parametresiyle belirtilen) Document
          - comp_def: eger 'document' bir Part ise onun PartComponentDefinition'i
            (degilse None)
          - tg: app.TransientGeometry (CreatePoint, CreatePoint2d, CreateVector, ...)
          - constants: win32com Inventor enum sabitleri (kJoinOperation, ...)
          - CastTo: win32com.client.CastTo (somut arayuze gecis icin)
          - dyn(obj): obj'yi dinamik (late-bound) COM dispatch'e cevirir - bazi
            Inventor surumlerinde CombineFeatures/SweepFeatures/LoftFeatures/
            CoilFeatures gibi nesnelerin gen_py onbellegi eksik/bozuk olabiliyor,
            bu fonksiyon o sorunu asar (bkz. feature_tools.py'deki ayni desen)
          - math: Python standart math modulu
          - mm_to_cm/cm_to_mm/deg_to_rad/rad_to_deg/cm2_to_mm2/cm3_to_mm3: birim
            donusum fonksiyonlari (Inventor COM API'si HER ZAMAN cm/radyan
            kullanir; kullanicidan/Claude'dan mm/derece beklenir)
          - point2d(app, x_mm, y_mm) / point3d(app, x_mm, y_mm, z_mm): mm
            cinsinden 2D/3D Point nesnesi olusturur

        Betik icinde print() ile yazilan her sey tool ciktisi olarak dondurulur.
        Betik bir Exception firlatirsa (yakalanmadan), hata mesaji COM
        detaylariyla birlikte tool hatasi olarak yukari tasinir - bu, hata
        ayiklamak icin ONEMLIDIR (bkz. Fusion MCP'nin ayni prensibi: script
        icinde exception yutmayin).
        """
        app = get_app()
        document = get_document(app, doc)
        try:
            comp_def = get_part_component_def(document)
        except LookupError:
            comp_def = None

        def dyn(obj):
            return win32com.client.dynamic.Dispatch(obj._oleobj_)

        namespace = {
            "app": app,
            "document": document,
            "comp_def": comp_def,
            "tg": app.TransientGeometry,
            "constants": constants,
            "CastTo": CastTo,
            "dyn": dyn,
            "math": math,
            "mm_to_cm": mm_to_cm,
            "cm_to_mm": cm_to_mm,
            "deg_to_rad": deg_to_rad,
            "rad_to_deg": rad_to_deg,
            "cm2_to_mm2": cm2_to_mm2,
            "cm3_to_mm3": cm3_to_mm3,
            "point2d": point2d,
            "point3d": point3d,
            "registry": registry,
        }

        stdout_buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(stdout_buf):
                exec(compile(script, "<inventor_run_script>", "exec"), namespace)
        except pywintypes.com_error as e:
            output = stdout_buf.getvalue()
            raise RuntimeError(
                f"{format_com_error(e)}\n--- betik ciktisi (hataya kadar) ---\n{output}"
            ) from None
        except Exception as e:
            output = stdout_buf.getvalue()
            raise RuntimeError(
                f"{type(e).__name__}: {e}\n--- betik ciktisi (hataya kadar) ---\n{output}"
            ) from None

        return {"output": stdout_buf.getvalue()}
