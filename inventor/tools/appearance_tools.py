"""Gorunum (appearance/renk) araclari: govde/yuz renklendirme."""

from win32com.client import CastTo, constants, dynamic

from inventor import registry
from inventor.connection import get_app
from inventor.helpers import get_document, get_part_component_def, inventor_tool


def _local_appearance_assets(document):
    if document.DocumentType == constants.kPartDocumentObject:
        return CastTo(document, "PartDocument").AppearanceAssets
    if document.DocumentType == constants.kAssemblyDocumentObject:
        return CastTo(document, "AssemblyDocument").AppearanceAssets
    raise LookupError("Bu islem sadece Part (.ipt) veya Assembly (.iam) dokumanlarinda calisir.")


def _iter_libraries(app, library: str | None):
    for i in range(1, app.AssetLibraries.Count + 1):
        lib = app.AssetLibraries.Item(i)
        if library and library.lower() not in lib.DisplayName.lower():
            continue
        yield lib


def _find_library_asset(app, name: str, library: str | None = None):
    """Asset kutuphanelerinde isme gore bir AppearanceAsset arar.

    Once tam (case-insensitive) eslesme, bulunamazsa ilk alt-dize eslesmesi donulur.
    """
    name_lower = name.lower()
    fallback = None
    for lib in _iter_libraries(app, library):
        assets = lib.AppearanceAssets
        for j in range(1, assets.Count + 1):
            a = assets.Item(j)
            # inventor_list_appearances hem ic adi ('name', orn. 'ACADGen-209') hem
            # gorunen adi ('display_name') dondurur; ikisi de kabul edilir.
            if a.DisplayName.lower() == name_lower or a.Name.lower() == name_lower:
                return a
            if fallback is None and name_lower in a.DisplayName.lower():
                fallback = a
    if fallback is not None:
        return fallback
    raise LookupError(f"'{name}' adli bir appearance (gorunum) bulunamadi.")


def _get_or_copy_local_asset(document, source_asset):
    local_assets = _local_appearance_assets(document)
    for i in range(1, local_assets.Count + 1):
        a = local_assets.Item(i)
        if a.DisplayName == source_asset.DisplayName:
            return a
    return source_asset.CopyTo(document)


def _diffuse_color(app, local_asset, color_rgb: list[int] | None, opacity: float | None):
    """generic_diffuse rengini (ve/veya opacity'sini) degistirir, [r,g,b,opacity] dondurur."""
    prop = local_asset.Item("generic_diffuse")
    dyn = dynamic.Dispatch(prop._oleobj_)
    current = dyn.Value
    r, g, b = color_rgb if color_rgb is not None else (current.Red, current.Green, current.Blue)
    a = opacity if opacity is not None else current.Opacity
    dyn.Value = app.TransientObjects.CreateColor(int(r), int(g), int(b), float(a))
    return [int(r), int(g), int(b), float(a)]


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_list_appearances(query: str | None = None, library: str | None = None, limit: int = 30) -> list[dict]:
        """Asset kutuphanelerindeki gorunum (appearance/renk/malzeme) isimlerini listeler.

        query: isimde aranan alt-dize (case-insensitive, orn. 'black', 'gold', 'aluminum').
        library: kutuphane adinda aranan alt-dize (orn. 'Appearance', 'Material').
        limit: donulecek maksimum sonuc sayisi.
        """
        app = get_app()
        query_lower = query.lower() if query else None

        result = []
        for lib in _iter_libraries(app, library):
            assets = lib.AppearanceAssets
            for j in range(1, assets.Count + 1):
                a = assets.Item(j)
                if query_lower and query_lower not in a.DisplayName.lower():
                    continue
                result.append({"library": lib.DisplayName, "name": a.Name, "display_name": a.DisplayName})
                if len(result) >= limit:
                    return result
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_set_appearance(
        appearance_name: str,
        target: str = "body",
        body_index: int = 1,
        entity_id: str | None = None,
        library: str | None = None,
        color_rgb: list[int] | None = None,
        opacity: float | None = None,
        doc: str | None = None,
    ) -> dict:
        """Bir govdeye veya yuze appearance (gorunum/renk) atar.

        target: 'body' (body_index ile) veya 'face' (entity_id ile, inventor_list_faces'tan).
        appearance_name: asset kutuphanesindeki gorunum adi (orn. 'Black', 'Nickel', 'Red');
        inventor_list_appearances'in dondurdugu 'name' veya 'display_name' olabilir.
        Tam isim eslesmezse alt-dize eslesmesiyle en yakin gorunum kullanilir
        (`inventor_list_appearances` ile arama yapilabilir).
        library: arama yapilacak kutuphane adinda alt-dize (orn. 'Appearance').
        color_rgb: [r,g,b] (0-255) verilirse, secilen gorunumun rengi bu degerle
        ozellestirilmis YENI bir kopya olarak atanir (orijinal kutuphane girdisi degismez).
        opacity: 0.0-1.0 araliginda saydamlik; verilirse rengin opacity'si de ayarlanir.
        """
        app = get_app()
        document = get_document(app, doc)

        if target == "body":
            comp_def = get_part_component_def(document)
            target_obj = comp_def.SurfaceBodies.Item(body_index)
            target_desc = f"body[{body_index}]"
        elif target == "face":
            if entity_id is None:
                raise LookupError("target='face' icin entity_id gerekli (inventor_list_faces ile alinabilir).")
            target_obj = registry.resolve(entity_id)
            target_desc = entity_id
        else:
            raise LookupError("target 'body' veya 'face' olmalidir.")

        source_asset = _find_library_asset(app, appearance_name, library)

        color_result = None
        if color_rgb is not None or opacity is not None:
            # Renk ozellestirmesi icin YENI bir kopya gerekir (ayni gorunumu kullanan diger
            # govdeler degismesin). Kutuphaneden ayni asset'i ikinci kez CopyTo ile almak
            # isim cakismasi yuzunden hata verir; dokumandaki kopyanin Duplicate'i ise
            # benzersiz adli ('Cored(1)' gibi) yeni bir asset olusturur.
            base = _get_or_copy_local_asset(document, source_asset)
            local = dynamic.Dispatch(base._oleobj_).Duplicate()
            color_result = _diffuse_color(app, local, color_rgb, opacity)
        else:
            local = _get_or_copy_local_asset(document, source_asset)

        target_obj.Appearance = local

        result = {"target": target_desc, "appearance": local.DisplayName}
        if color_result is not None:
            result["color_rgb"] = color_result[:3]
            result["opacity"] = color_result[3]
        return result
