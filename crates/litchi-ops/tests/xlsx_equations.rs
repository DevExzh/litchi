//! XLSX equations through the unified API: read, add, set and remove in
//! the worksheet drawing, for source and eager owners, with read-back,
//! untouched equations kept, and typed refusals that leave the input
//! intact.
#![cfg(feature = "xlsx")]

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

/// An ONLYOFFICE-written workbook whose first sheet holds an equation in a
/// text box paragraph (see `test-data/poi/PROVENANCE.md`).
const ONLYOFFICE: &[u8] = include_bytes!(
    "../../../test-data/poi/test-data/spreadsheet/LIBRE_OFFICE-128382-0-small-media.xlsx"
);

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

/// Both owners of `bytes`: source-backed and eager.
fn owners(bytes: &[u8]) -> [File; 2] {
    [SourceMode::Auto, SourceMode::Eager].map(|mode| {
        litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode)).unwrap()
    })
}

#[cfg(feature = "equation")]
fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[cfg(feature = "equation")]
fn compact(text: Option<String>) -> Option<String> {
    text.map(|text| text.chars().filter(|ch| !ch.is_whitespace()).collect())
}

fn equation_paths(file: &File) -> Vec<String> {
    file.query("equation")
        .unwrap()
        .map(|node| node.unwrap().view().path().to_string())
        .collect()
}

/// A new workbook with one empty sheet.
#[cfg(feature = "equation")]
fn blank() -> File {
    litchi_ops::create(litchi_core::FileFormat::Xlsx).unwrap()
}

#[test]
fn equations_are_listed_with_their_cell_and_omml() {
    for file in owners(ONLYOFFICE) {
        assert_eq!(equation_paths(&file), ["/sheet[1]/equation[1]"]);
        let path = "/Basic features/equation[1]";
        assert_eq!(prop(&file, path, "anchor").as_deref(), Some("A22"));
        assert_eq!(prop(&file, path, "name").as_deref(), Some("TextBox 3"));
        let omml = prop(&file, path, "omml").unwrap();
        assert!(omml.starts_with("<m:oMathPara>"), "{omml}");
        #[cfg(feature = "equation")]
        assert!(
            prop(&file, path, "latex").unwrap().starts_with("e^x= 1+"),
            "{:?}",
            prop(&file, path, "latex")
        );
        #[cfg(not(feature = "equation"))]
        assert_eq!(prop(&file, path, "latex"), None);
        // The positional sheet path reaches it too.
        assert!(file.get("/sheet[1]/math[1]").is_ok());
    }
}

#[cfg(not(feature = "equation"))]
#[test]
fn writes_need_the_equation_feature() {
    for file in owners(ONLYOFFICE) {
        let error = file
            .set("/Basic features/equation[1]", props! { "latex" => "x" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file.remove("/Basic features/equation[1]").unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .add(
                "/Basic features",
                NewNode::new("equation").prop("latex", "x"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }
}

#[cfg(feature = "equation")]
#[test]
fn equations_are_added_at_a_cell_and_round_trip() {
    let file = blank();
    let added = file
        .add(
            "/Sheet1/C4",
            NewNode::new("equation")
                .prop("latex", r"\int_0^1 x\,dx")
                .prop("size", "14pt"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/sheet[1]/equation[1]");
    let read = reopen(&added.file);
    assert_eq!(
        compact(prop(&read, "/Sheet1/equation[1]", "latex")).as_deref(),
        Some(r"\int_0^1x\,dx")
    );
    assert_eq!(
        prop(&read, "/Sheet1/equation[1]", "anchor").as_deref(),
        Some("C4")
    );
    let omml = prop(&read, "/Sheet1/equation[1]", "omml").unwrap();
    assert!(omml.starts_with("<m:oMathPara "), "{omml}");
    assert!(omml.contains(r#"sz="1400""#), "{omml}");
    // The text box is not one of the sheet's shapes.
    assert!(read.get("/Sheet1/shape[1]").is_err());

    // OMML is written as given, at the `anchor` cell.
    let raw = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>q</m:t></m:r></m:oMath>"#;
    let added = read
        .add(
            "/Sheet1",
            NewNode::new("math")
                .prop("omml", raw)
                .prop("anchor", "$B$9"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/sheet[1]/equation[2]");
    let read = reopen(&added.file);
    assert_eq!(
        prop(&read, "/Sheet1/equation[2]", "omml").as_deref(),
        Some(raw)
    );
    assert_eq!(
        prop(&read, "/Sheet1/equation[2]", "anchor").as_deref(),
        Some("B9")
    );
}

#[cfg(feature = "equation")]
#[test]
fn an_equation_is_replaced_and_removed_and_the_others_keep_their_bytes() {
    let file = blank();
    let file = file
        .add("/Sheet1/A1", NewNode::new("equation").prop("latex", "a+b"))
        .unwrap()
        .file;
    let file = file
        .add(
            "/Sheet1/A8",
            NewNode::new("equation").prop("latex", r"\sqrt{x}"),
        )
        .unwrap()
        .file;
    let bytes = file.to_bytes().unwrap();
    for file in owners(&bytes) {
        let other = prop(&file, "/Sheet1/equation[2]", "omml");
        let changed = file
            .set("/Sheet1/equation[1]", props! { "latex" => r"\frac{a}{b}" })
            .unwrap();
        assert_eq!(file.to_bytes().unwrap(), bytes, "input unchanged");
        let read = reopen(&changed);
        assert_eq!(
            compact(prop(&read, "/Sheet1/equation[1]", "latex")).as_deref(),
            Some(r"\frac{a}{b}")
        );
        assert_eq!(prop(&read, "/Sheet1/equation[2]", "omml"), other);

        let removed = read.remove("/Sheet1/equation[1]").unwrap();
        let read = reopen(&removed);
        assert_eq!(equation_paths(&read), ["/sheet[1]/equation[1]"]);
        assert_eq!(prop(&read, "/Sheet1/equation[1]", "omml"), other);

        // The same content again is an exact no-op.
        let raw = other.unwrap();
        let bytes = read.to_bytes().unwrap();
        let same = read
            .set("/Sheet1/equation[1]", props! { "omml" => raw.as_str() })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
    }
}

#[cfg(feature = "equation")]
#[test]
fn bad_input_and_unsupported_targets_are_typed_refusals() {
    for file in owners(ONLYOFFICE) {
        let before = file.to_bytes().unwrap();
        // The ONLYOFFICE equation sits inside a paragraph: read only.
        let error = file
            .set("/Basic features/equation[1]", props! { "latex" => "y" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        let error = file.remove("/Basic features/equation[1]").unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        let error = file
            .add(
                "/Basic features",
                NewNode::new("equation").prop("latex", r"\frac{"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue, "{error}");
        let error = file
            .add(
                "/Basic features",
                NewNode::new("equation")
                    .prop("latex", "x")
                    .prop("omml", "<m:oMath/>"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument, "{error}");
        let error = file
            .add("/Basic features", NewNode::new("equation"))
            .unwrap_err();
        assert_eq!(error.code(), Code::MissingProperty, "{error}");
        let error = file
            .add(
                "/Basic features",
                NewNode::new("equation").prop("omml", "<m:r/>"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue, "{error}");
        let error = file
            .add(
                "/Basic features",
                NewNode::new("equation")
                    .prop("latex", "x")
                    .prop("anchor", "nowhere"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue, "{error}");
        assert!(file.get("/Basic features/equation[2]").is_err());
        assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
    }
}
