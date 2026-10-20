"""iProperty ve kutle ozellikleri araclari."""

import datetime

import pywintypes

from inventor.connection import get_app
from inventor.helpers import get_document, get_part_component_def, inventor_tool
from inventor.units import cm2_to_mm2, cm3_to_mm3, cm_to_mm


def _json_safe(value):
    if isinstance(value, datetime.datetime):
        return value.isoformat()
    return value


USER_DEFINED_PROPSET_ID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"  # "Inventor User Defined Properties"


def _resolve_property_set(doc, propset: str):
    propset_lower = propset.lower()
    for ps in doc.PropertySets:
        if ps.Name.lower() == propset_lower or ps.DisplayName.lower() == propset_lower:
            return ps
    raise LookupError(f"'{propset}' adli bir property set bulunamadi.")


def register(mcp):
    @mcp.tool()
    @inventor_tool
    async def inventor_get_iproperties(doc: str | None = None) -> dict:
        """Dokumanin tum iProperty setlerini (Summary/Document Summary/Design Tracking/
        User Defined) ve icindeki ozellik degerlerini dondurur."""
        app = get_app()
        document = get_document(app, doc)

        result = {}
        for ps in document.PropertySets:
            props = {}
            for prop in ps:
                try:
                    props[prop.Name] = _json_safe(prop.Value)
                except pywintypes.com_error:
                    continue
            result[ps.DisplayName] = props
        return result

    @mcp.tool()
    @inventor_tool
    async def inventor_set_iproperty(
        propset: str, name: str, value: str | float | int, doc: str | None = None
    ) -> dict:
        """Bir iProperty degerini ayarlar.

        propset: 'Summary Information', 'Document Summary Information',
        'Design Tracking Properties' veya 'User Defined Properties' (DisplayName veya
        internal Name kabul edilir). 'User Defined Properties' icinde olmayan bir ad
        verilirse yeni ozel ozellik olusturulur (sonucta created=True).
        """
        app = get_app()
        document = get_document(app, doc)
        ps = _resolve_property_set(document, propset)
        created = False
        try:
            prop = ps.Item(name)
            prop.Value = value
        except pywintypes.com_error:
            # Ozel (User Defined) ozellikler ilk kez ayarlandiginda olusturulur - UI'daki
            # iProperties > Custom ile ayni davranis. Yerlesik setlere yeni ozellik eklenemez.
            if ps.InternalName != USER_DEFINED_PROPSET_ID:
                raise LookupError(f"'{name}' adli ozellik '{ps.DisplayName}' icinde bulunamadi.") from None
            prop = ps.Add(value, name)
            created = True
        return {"propset": ps.DisplayName, "name": prop.Name, "value": _json_safe(prop.Value), "created": created}

    @mcp.tool()
    @inventor_tool
    async def inventor_get_mass_properties(doc: str | None = None) -> dict:
        """Part dokumaninin kutle, hacim, yuzey alani, agirlik merkezi ve atalet
        ozelliklerini dondurur (kg, mm, mm^2, mm^3, kg*mm^2)."""
        app = get_app()
        document = get_document(app, doc)
        comp_def = get_part_component_def(document)
        mp = comp_def.MassProperties

        com = mp.CenterOfMass
        i1, i2, i3 = mp.PrincipalMomentsOfInertia()
        kx, ky, kz = mp.RadiusOfGyration()
        ixx, iyy, izz, ixy, iyz, ixz = mp.XYZMomentsOfInertia()

        return {
            "mass_kg": mp.Mass,
            "volume_mm3": cm3_to_mm3(mp.Volume),
            "area_mm2": cm2_to_mm2(mp.Area),
            "center_of_mass_mm": [cm_to_mm(com.X), cm_to_mm(com.Y), cm_to_mm(com.Z)],
            "principal_moments_of_inertia_kg_mm2": [cm2_to_mm2(i1), cm2_to_mm2(i2), cm2_to_mm2(i3)],
            "radius_of_gyration_mm": [cm_to_mm(kx), cm_to_mm(ky), cm_to_mm(kz)],
            "xyz_moments_of_inertia_kg_mm2": {
                "ixx": cm2_to_mm2(ixx),
                "iyy": cm2_to_mm2(iyy),
                "izz": cm2_to_mm2(izz),
                "ixy": cm2_to_mm2(ixy),
                "iyz": cm2_to_mm2(iyz),
                "ixz": cm2_to_mm2(ixz),
            },
        }
