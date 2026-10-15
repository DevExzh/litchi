//! Golden snapshots of the property registry export (`litchi_ops::schema`).
//!
//! Each registered format's export is compared with
//! `tests/golden/registry/<format>.json`; a registry change therefore shows
//! as a reviewed diff. Regenerate with `LITCHI_BLESS=1` after an intended
//! change, then rerun `tools/officecli_capability_delta.py`, which reads
//! these snapshots.
//!
//! As with the capability goldens, a build with `encryption` writes the DOCX
//! and XLSX snapshots as `<format>-encryption.json`: that feature adds
//! operation gaps to those formats.
use litchi_core::FileFormat;

#[test]
fn registry_goldens_follow_enabled_formats() {
    for (format, name, enabled) in [
        (FileFormat::Doc, "doc", cfg!(feature = "doc")),
        (FileFormat::Docx, "docx", cfg!(feature = "docx")),
        (FileFormat::Ppt, "ppt", cfg!(feature = "ppt")),
        (FileFormat::Pptx, "pptx", cfg!(feature = "pptx")),
        (FileFormat::Rtf, "rtf", cfg!(feature = "rtf")),
        (FileFormat::Odt, "odt", cfg!(feature = "odt")),
        (FileFormat::Xls, "xls", cfg!(feature = "xls")),
        (FileFormat::Xlsb, "xlsb", cfg!(feature = "xlsb")),
        (FileFormat::Xlsx, "xlsx", cfg!(feature = "xlsx")),
        (FileFormat::Ods, "ods", cfg!(feature = "ods")),
        (FileFormat::Odp, "odp", cfg!(feature = "odp")),
        (FileFormat::Oth, "oth", cfg!(feature = "oth")),
        (FileFormat::Odm, "odm", cfg!(feature = "odm")),
        (FileFormat::Odg, "odg", cfg!(feature = "odg")),
        (FileFormat::Odc, "odc", cfg!(feature = "odc")),
        (FileFormat::Odi, "odi", cfg!(feature = "odi")),
        (FileFormat::Odb, "odb", cfg!(feature = "odb")),
        (FileFormat::Odf, "odf", cfg!(feature = "odf-formula")),
        (FileFormat::Pages, "pages", cfg!(feature = "pages")),
        (FileFormat::Keynote, "keynote", cfg!(feature = "keynote")),
        (FileFormat::Numbers, "numbers", cfg!(feature = "numbers")),
    ] {
        let schema = litchi_ops::schema(format);
        assert_eq!(schema.registered(), enabled, "{name}");
        if !enabled {
            continue;
        }
        let actual = format!("{}\n", serde_json::to_string_pretty(&schema).unwrap());
        let golden_name = match (name, cfg!(feature = "encryption")) {
            ("docx" | "xlsx", true) => format!("{name}-encryption"),
            _ => name.to_owned(),
        };
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/golden/registry")
            .join(format!("{golden_name}.json"));
        if std::env::var_os("LITCHI_BLESS").is_some() {
            std::fs::create_dir_all(path.parent().unwrap()).unwrap();
            std::fs::write(&path, &actual).unwrap();
        }
        let expected = std::fs::read_to_string(&path).unwrap_or_else(|error| {
            panic!("{}: {error}; bless with LITCHI_BLESS=1", path.display())
        });
        assert_eq!(expected, actual, "{golden_name}");
    }
}

/// The export agrees with the capability report on every element and on
/// every property's flags: both read the same registry declarations.
#[test]
fn registry_export_matches_capabilities() {
    use litchi_ops::capabilities::Capabilities;

    for format in [
        FileFormat::Doc,
        FileFormat::Docx,
        FileFormat::Ppt,
        FileFormat::Pptx,
        FileFormat::Rtf,
        FileFormat::Xls,
        FileFormat::Xlsb,
        FileFormat::Xlsx,
        FileFormat::Ods,
        FileFormat::Odp,
    ] {
        let schema = litchi_ops::schema(format);
        let capabilities = Capabilities::of(format);
        assert_eq!(
            schema.elements().count(),
            capabilities.elements().count(),
            "{format:?}"
        );
        for (exported, reported) in schema.elements().zip(capabilities.elements()) {
            assert_eq!(exported.name(), reported.name());
            for (a, b) in exported.properties().zip(reported.properties()) {
                assert_eq!(a.name(), b.name());
                assert_eq!((a.get(), a.set(), a.add()), (b.get(), b.set(), b.add()));
                // An alias never shadows a property of the same element.
                for alias in a.aliases() {
                    assert!(
                        exported.property(alias).is_none(),
                        "{format:?} {}.{alias}",
                        exported.name()
                    );
                }
            }
        }
    }
}
