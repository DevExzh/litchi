# Native package schema-validation limitation

The authoritative RNG validation run was attempted for the API-mutated native candidate and for both LibreOffice outputs (`save1` and `save2`). All three validation attempts failed. The native outputs retain extension markup, so this validator run does not establish that any of the three complete packages conform to the selected normative RNG schema.

The LibreOffice evidence in [native-roundtrip-final-freeze-0202.log](native-roundtrip-final-freeze-0202.log) therefore records successful headless open/save/reopen and Litchi typed readback only. It must not be read as a blanket normative schema-validity claim. The field-level semantic losses observed after each save remain independent of that schema-validation limitation. The owner-level RNG result for the retained API-mutated input is recorded separately in [native-owner-schema.json](native-owner-schema.json).
