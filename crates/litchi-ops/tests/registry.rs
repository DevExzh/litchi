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

/// Every format this build enables.
fn enabled_formats() -> impl Iterator<Item = FileFormat> {
    [
        (FileFormat::Doc, cfg!(feature = "doc")),
        (FileFormat::Docx, cfg!(feature = "docx")),
        (FileFormat::Ppt, cfg!(feature = "ppt")),
        (FileFormat::Pptx, cfg!(feature = "pptx")),
        (FileFormat::Rtf, cfg!(feature = "rtf")),
        (FileFormat::Odt, cfg!(feature = "odt")),
        (FileFormat::Xls, cfg!(feature = "xls")),
        (FileFormat::Xlsb, cfg!(feature = "xlsb")),
        (FileFormat::Xlsx, cfg!(feature = "xlsx")),
        (FileFormat::Ods, cfg!(feature = "ods")),
        (FileFormat::Odp, cfg!(feature = "odp")),
        (FileFormat::Oth, cfg!(feature = "oth")),
        (FileFormat::Odm, cfg!(feature = "odm")),
        (FileFormat::Odg, cfg!(feature = "odg")),
        (FileFormat::Odc, cfg!(feature = "odc")),
        (FileFormat::Odi, cfg!(feature = "odi")),
        (FileFormat::Odb, cfg!(feature = "odb")),
        (FileFormat::Odf, cfg!(feature = "odf-formula")),
        (FileFormat::Pages, cfg!(feature = "pages")),
        (FileFormat::Keynote, cfg!(feature = "keynote")),
        (FileFormat::Numbers, cfg!(feature = "numbers")),
    ]
    .into_iter()
    .filter_map(|(format, enabled)| enabled.then_some(format))
}

/// An element's operation claims agree with its properties' flags (record
/// 0330). Both come from the same spec tables, but through different
/// lists, so they can drift; this is the rule that keeps them together:
///
/// 1. an element claims `set` exactly when at least one of its properties
///    is settable (a `set` with no settable key could only be refused, and a
///    settable key on an element without `set` is never offered);
/// 2. an element with an addable property claims `add` (an element may
///    claim `add` without any: a slide or a row is added bare);
/// 3. an element that claims `add`, `set` or `remove` also claims `get`,
///    and an element with a gap claims nothing and declares no writable
///    property.
///
/// Contextual refusals (a selection, an owner state, one property) are
/// operation or property gaps; they never turn an element claim off. The
/// report checked is the format's default one ([`Capabilities::of`]); a
/// file's report only narrows it by owner state.
#[test]
fn element_operations_agree_with_property_flags() {
    use litchi_ops::capabilities::Capabilities;

    let mut contradictions = Vec::new();
    for format in enabled_formats() {
        let capabilities = Capabilities::of(format);
        for element in capabilities.elements() {
            let operations = element.operations();
            let settable: Vec<_> = element
                .properties()
                .filter(|p| p.set())
                .map(|p| p.name())
                .collect();
            let addable: Vec<_> = element
                .properties()
                .filter(|p| p.add())
                .map(|p| p.name())
                .collect();
            let name = element.name();
            if operations.set && settable.is_empty() {
                contradictions.push(format!(
                    "{format:?} {name}: claims set without a settable property"
                ));
            }
            if !operations.set && !settable.is_empty() {
                contradictions.push(format!(
                    "{format:?} {name}: settable {settable:?} without a set claim"
                ));
            }
            if !operations.add && !addable.is_empty() {
                contradictions.push(format!(
                    "{format:?} {name}: addable {addable:?} without an add claim"
                ));
            }
            if (operations.add || operations.set || operations.remove) && !operations.get {
                contradictions.push(format!("{format:?} {name}: writes without get"));
            }
            if element.gap().is_some()
                && (operations.add
                    || operations.set
                    || operations.remove
                    || !settable.is_empty()
                    || !addable.is_empty())
            {
                contradictions.push(format!("{format:?} {name}: a gap element claims writes"));
            }
        }
    }
    assert!(
        contradictions.is_empty(),
        "element claims contradict property flags:\n{}",
        contradictions.join("\n")
    );
}
