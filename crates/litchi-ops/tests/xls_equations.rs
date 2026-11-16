//! XLS OLE equations through the unified API: add, get, query, set and
//! remove on both owners, with read-back, untouched objects kept, and typed
//! refusals that leave the input intact.
#![cfg(feature = "xls")]

use litchi_core::FileFormat;
#[cfg(feature = "equation")]
use litchi_ops::props;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode};

#[cfg(feature = "equation")]
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

fn embedded() -> Vec<u8> {
    std::fs::read(format!(
        "{}/../../test-data/ole/xls/WithEmbeddedObjects.xls",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap()
}

fn blank() -> Vec<u8> {
    litchi_ops::create(FileFormat::Xls)
        .unwrap()
        .to_bytes()
        .unwrap()
}

/// An Equation Editor 3.0 `Equation Native` stream holding the digit 2,
/// base64-encoded.
#[cfg(feature = "equation")]
const EQUATION3: &str = "HAAAAAIAbcENAAAAAAAAAAAAAAAAAAAAAAAAAAMBAQMACgECiDIAAAA=";

#[cfg(not(feature = "equation"))]
#[test]
fn writes_need_the_equation_feature() {
    for file in owners(&blank()) {
        let error = file
            .add("/Sheet1", NewNode::new("equation").prop("latex", "x"))
            .unwrap_err();
        // Without the feature the content keys are read-only.
        assert_eq!(error.code(), Code::UnsupportedProperty);
        let capabilities = litchi_ops::capabilities::Capabilities::of(FileFormat::Xls);
        let json = serde_json::to_value(capabilities).unwrap();
        let equation = json["elements"]
            .as_array()
            .unwrap()
            .iter()
            .find(|element| element["element"] == "equation")
            .unwrap();
        assert_eq!(equation["operations"]["add"], false);
        assert_eq!(equation["operations"]["set"], false);
        assert_eq!(equation["operations"]["get"], true);
    }
}

#[test]
fn a_workbook_without_equations_lists_none() {
    for file in owners(&embedded()) {
        assert_eq!(file.query("equation").unwrap().count(), 0);
        // The Excel-authored embedded documents stay pictures.
        assert!(file.query("picture").unwrap().count() > 0);
    }
}

#[cfg(feature = "equation")]
#[test]
fn equations_are_added_read_set_and_removed() {
    for file in owners(&blank()) {
        let before = file.to_bytes().unwrap();
        let added = file
            .add(
                "/Sheet1",
                NewNode::new("equation")
                    .prop("latex", r"\frac{a}{b}")
                    .prop("anchor", "C3"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/sheet[1]/equation[1]");
        assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
        for read in owners(&added.file.to_bytes().unwrap()) {
            let path = "/Sheet1/equation[1]";
            assert_eq!(prop(&read, path, "latex").as_deref(), Some(r"\frac{a}{b}"));
            assert_eq!(
                prop(&read, path, "progId").as_deref(),
                Some("Equation.DSMT4")
            );
            assert_eq!(prop(&read, path, "anchor").as_deref(), Some("C3"));
            assert!(prop(&read, path, "mtef").is_some());
            assert_eq!(read.query("equation").unwrap().count(), 1);
            assert_eq!(read.query("picture").unwrap().count(), 0);
        }
        // A second equation, from MTEF (Equation Editor 3.0 keeps its kind).
        let second = added
            .file
            .add("/Sheet1/B10", NewNode::new("math").prop("mtef", EQUATION3))
            .unwrap();
        assert_eq!(second.path.to_string(), "/sheet[1]/equation[2]");
        let file = second.file;
        assert_eq!(
            prop(&file, "/Sheet1/equation[2]", "progId").as_deref(),
            Some("Equation.3")
        );
        assert_eq!(
            prop(&file, "/Sheet1/equation[2]", "anchor").as_deref(),
            Some("B10")
        );
        let first = prop(&file, "/Sheet1/equation[1]", "mtef");

        // Set keeps the kind and leaves the other equation alone.
        let set = file
            .set("/Sheet1/equation[2]", props! { "latex" => r"\sqrt{x}" })
            .unwrap();
        let file = set;
        assert_eq!(
            prop(&file, "/Sheet1/equation[2]", "latex").as_deref(),
            Some(r"\sqrt{x}")
        );
        assert_eq!(
            prop(&file, "/Sheet1/equation[2]", "progId").as_deref(),
            Some("Equation.3")
        );
        assert_eq!(prop(&file, "/Sheet1/equation[1]", "mtef"), first);
        let set = file
            .set(
                "/Sheet1/equation[1]",
                props! { "omml" => "<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>" },
            )
            .unwrap();
        let file = set;
        assert_eq!(
            prop(&file, "/Sheet1/equation[1]", "latex").as_deref(),
            Some("y")
        );

        // Remove one, then the other.
        let removed = file.remove("/Sheet1/equation[1]").unwrap();
        let file = removed;
        assert_eq!(file.query("equation").unwrap().count(), 1);
        assert_eq!(
            prop(&file, "/Sheet1/equation[1]", "progId").as_deref(),
            Some("Equation.3")
        );
        let removed = file.remove("/Sheet1/equation[1]").unwrap();
        assert_eq!(removed.query("equation").unwrap().count(), 0);
    }
}

#[cfg(feature = "equation")]
#[test]
fn equations_join_excel_embedded_objects() {
    for file in owners(&embedded()) {
        let pictures = file.query("picture").unwrap().count();
        let added = file
            .add("/Sheet1", NewNode::new("equation").prop("latex", "E=mc^2"))
            .unwrap();
        let read = owners(&added.file.to_bytes().unwrap());
        for read in read {
            assert_eq!(read.query("picture").unwrap().count(), pictures);
            assert_eq!(read.query("equation").unwrap().count(), 1);
            let latex = prop(&read, "/Sheet1/equation[1]", "latex").unwrap();
            let latex = latex.replace([' ', '{', '}'], "");
            assert_eq!(latex, "E=mc^2");
        }
    }
}

#[cfg(feature = "equation")]
#[test]
fn refusals_leave_the_workbook_intact() {
    for file in owners(&blank()) {
        let before = file.to_bytes().unwrap();
        let error = file.add("/Sheet1", NewNode::new("equation")).unwrap_err();
        assert_eq!(error.code(), Code::MissingProperty);
        let error = file
            .add(
                "/Sheet1",
                NewNode::new("equation")
                    .prop("latex", "x")
                    .prop("omml", "<m:oMath/>"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        let error = file
            .add("/Sheet1", NewNode::new("equation").prop("latex", r"\frac{"))
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .add(
                "/Sheet1",
                NewNode::new("equation").prop("mtef", "not base64!"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .add(
                "/Sheet1",
                NewNode::new("equation").prop("progId", "Equation.3"),
            )
            .unwrap_err();
        assert_ne!(error.code(), Code::InternalError);
        assert!(file.get("/Sheet1/equation[1]").is_err());
        assert_eq!(file.to_bytes().unwrap(), before);
    }
}
