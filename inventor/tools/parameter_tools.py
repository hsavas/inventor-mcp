"""Parametre araclari: model/kullanici parametrelerini listeleme, guncelleme, olusturma."""

import pywintypes
from win32com.client import constants

from inventor.connection import get_app
from inventor.helpers import get_document, get_part_component_def, inventor_tool

PARAMETER_TYPE_NAMES = {
    constants.kModelParameter: "model",
    constants.kReferenceParameter: "reference",
    constants.kTableParameter: "table",
    constants.kUserParameter: "user",
    constants.kDerivedParameter: "derived",
}

# kFinishParameter bazi Inventor surumlerinin typelib'inde bulunmuyor; varsa ekle.
_finish_parameter = getattr(constants, "kFinishParameter", None)
if _finish_parameter is not None:
    PARAMETER_TYPE_NAMES[_finish_parameter] = "finish"

def _units_specifier_candidates(document):
    """AddByExpression icin denenecek UnitsSpecifier sirasi: uzunluk -> aci -> birimsiz.

    DIKKAT: burada kDatabaseLengthUnits KULLANILMAZ. Database birimleri her zaman
    santimetredir; onu UnitsSpecifier olarak vermek, dokuman milimetre olsa bile
    parametreyi cm cinsinden olusturur ("46 mm" dogru degerlendirilir ama
    parametre listesinde cm olarak gorunur). Bunun yerine dokumanin kendi
    goruntu birimleri kullanilir.
    """
    uom = document.UnitsOfMeasure
    return [uom.LengthUnits, uom.AngleUnits, constants.kUnitlessUnits]


def _display_value(uom, value_db, units):
    """Database birimindeki bir degeri parametrenin kendi birimine cevirir.

    ConvertUnits bos kaynak birimi kabul etmedigi icin once uzunluk, sonra aci
    database birimi denenir; ikisi de uymuyorsa (birimsiz parametreler) ham
    deger zaten dogrudur.
    """
    for source in (constants.kDatabaseLengthUnits, constants.kDatabaseAngleUnits):
        try:
            return uom.ConvertUnits(value_db, source, units)
        except Exception:
            continue
    return value_db


def _parameter_summary(p, document=None) -> dict:
    """Parametre ozeti.

    'value' parametrenin KENDI biriminde dondurulur (orn. '46 mm' -> 46.0).
    p.Value her zaman database birimindedir (uzunluklar icin cm), ki bu mm
    tabanli bu API'de yaniltici olur; ham deger 'value_db' alaninda kalir.
    """
    value_db = p.Value
    value = value_db
    if document is not None:
        value = _display_value(document.UnitsOfMeasure, value_db, p.Units)
    return {
        "name": p.Name,
        "value": value,
        "value_db": value_db,
        "expression": p.Expression,
        "units": p.Units,
        "comment": p.Comment,
        "kind": PARAMETER_TYPE_NAMES.get(p.ParameterType, str(p.ParameterType)),
    }


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_list_parameters(doc: str | None = None) -> list[dict]:
        """Part dokumanindaki tum parametreleri (model + kullanici + referans) listeler."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        return [_parameter_summary(p, document) for p in comp_def.Parameters]

    @mcp.tool()
    @inventor_tool
    async def inventor_set_parameter(name: str, expression: str, doc: str | None = None) -> dict:
        """Var olan bir parametrenin ifadesini (expression) degistirir.

        expression: '12 mm', '45 deg', '2 * d0' gibi birim/ifade iceren bir string.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        try:
            param = comp_def.Parameters.Item(name)
        except pywintypes.com_error:
            raise LookupError(f"'{name}' adli parametre bulunamadi.")
        param.Expression = expression
        # Parametre degisikligi modeli kendiliginden yeniden hesaplamaz; Update olmadan
        # govde eski olculerde kalir (hacim/sinir kutusu degismez).
        try:
            document.Update()
        except pywintypes.com_error:
            pass
        return _parameter_summary(param, document)

    @mcp.tool()
    @inventor_tool
    async def inventor_create_parameter(
        name: str,
        expression: str,
        comment: str | None = None,
        units: str | None = None,
        doc: str | None = None,
    ) -> dict:
        """Yeni bir kullanici parametresi (UserParameter) olusturur.

        expression: '12 mm', '45 deg', '2' gibi birim/ifade iceren bir string.
        units: parametrenin birimi ('mm', 'deg', 'ul', ...). Verilmezse dokumanin
        kendi birimleri (uzunluk -> aci -> birimsiz sirasiyla) denenir.
        """
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        user_params = comp_def.Parameters.UserParameters

        if units:
            candidates = [units]
        else:
            candidates = _units_specifier_candidates(document)

        param = None
        last_error = None
        for units_specifier in candidates:
            try:
                param = user_params.AddByExpression(name, expression, units_specifier)
                break
            except pywintypes.com_error as e:
                last_error = e
        if param is None:
            raise last_error

        if comment:
            param.Comment = comment
        return _parameter_summary(param, document)
