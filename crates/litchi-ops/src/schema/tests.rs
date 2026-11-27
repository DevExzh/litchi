use super::*;

#[cfg(not(feature = "pages"))]
#[test]
fn unregistered_formats_export_an_empty_pending_schema() {
    let pages = schema(FileFormat::Pages);
    assert!(!pages.registered());
    assert_eq!(pages.elements().count(), 0);
    let json = serde_json::to_value(pages).unwrap();
    assert_eq!(json["format"], "pages");
    assert_eq!(json["pending"]["supported"], false);
    assert_eq!(json["elements"], serde_json::json!([]));
}

#[cfg(feature = "xlsx")]
#[test]
fn aliased_tables_are_exported() {
    use crate::format::xlsx::{CELL_WRITE_TABLE, COLUMN_WRITE_TABLE, ROW_WRITE_TABLE};

    // Every aliased registry table, with an element that writes through it.
    // A table declared with aliases must add its arm to `prop_aliases` and
    // its row here.
    let tables: [(&str, &[PropAlias]); 4] = [
        ("cell", &CELL_WRITE_TABLE.prop_aliases()),
        ("range", &CELL_WRITE_TABLE.prop_aliases()),
        ("row", &ROW_WRITE_TABLE.prop_aliases()),
        ("col", &COLUMN_WRITE_TABLE.prop_aliases()),
    ];
    let xlsx = schema(FileFormat::Xlsx);
    for (element, aliases) in tables {
        assert!(!aliases.is_empty(), "{element}");
        let element = xlsx.element(element).unwrap();
        for alias in aliases {
            let property = element
                .property(alias.name)
                .unwrap_or_else(|| panic!("{} has no {}", element.name(), alias.name));
            assert!(
                property.aliases().any(|spelling| spelling == alias.alias),
                "{}.{} lacks alias {}",
                element.name(),
                alias.name,
                alias.alias
            );
        }
    }
    let row = xlsx.element("row").unwrap();
    let outline = row.property("outline").unwrap();
    assert_eq!(outline.constraint(), Constraint::Int { min: 0, max: 7 });
    assert!(outline.clearable());
    let json = serde_json::to_value(xlsx).unwrap();
    let row = json["elements"]
        .as_array()
        .unwrap()
        .iter()
        .find(|element| element["element"] == "row")
        .unwrap();
    assert_eq!(
        row["properties"]["outline"]["aliases"],
        serde_json::json!(["group", "outlinelevel"])
    );
    assert_eq!(
        row["properties"]["outline"]["constraint"],
        serde_json::json!({"min": 0, "max": 7})
    );
    assert_eq!(
        row["properties"]["height"]["constraint"]["maxEmu"],
        serde_json::json!(409 * 12_700)
    );
}

#[cfg(feature = "docx")]
#[test]
fn docx_export_reports_units_tokens_and_flags() {
    let docx = schema(FileFormat::Docx);
    assert!(docx.registered());
    let json = serde_json::to_value(docx).unwrap();
    let run = json["elements"]
        .as_array()
        .unwrap()
        .iter()
        .find(|element| element["element"] == "run")
        .unwrap();
    let size = &run["properties"]["size"];
    assert_eq!(size["type"], "font-size");
    assert_eq!(size["unit"], "pt");
    let underline = &run["properties"]["underline"];
    assert_eq!(underline["type"], "enum");
    assert!(
        underline["values"]
            .as_array()
            .is_some_and(|v| !v.is_empty())
    );
    assert_eq!(run["properties"]["color"]["colorContext"], "word");
    for (_, property) in run["properties"].as_object().unwrap() {
        for flag in ["add", "set", "get"] {
            assert!(property[flag].is_boolean());
        }
    }
}

#[test]
fn property_gaps_attach_to_their_element() {
    for format in [
        FileFormat::Doc,
        FileFormat::Docx,
        FileFormat::Ppt,
        FileFormat::Pptx,
        FileFormat::Rtf,
        FileFormat::Xls,
        FileFormat::Xlsb,
        FileFormat::Xlsx,
    ] {
        let export = schema(format);
        let json = serde_json::to_value(export).unwrap();
        let attached: usize = json["elements"]
            .as_array()
            .unwrap()
            .iter()
            .map(|element| {
                element["propertyGaps"]
                    .as_object()
                    .map_or(0, |gaps| gaps.len())
            })
            .sum();
        let declared = export
            .gaps()
            .filter(|gap| matches!(gap.subject(), Subject::Property { .. }))
            .count();
        assert_eq!(attached, declared, "{format:?}");
        // No gap is lost: each is listed with its element or at the top.
        let listed = json["gaps"].as_array().unwrap().len()
            + attached
            + json["elements"]
                .as_array()
                .unwrap()
                .iter()
                .filter(|element| element.get("gap").is_some())
                .count();
        assert_eq!(listed, export.gaps().count(), "{format:?}");
    }
}
