//! PPT equations through the unified API: read (`latex`, `mtef`, `progId`
//! and the frame), add, set and remove for source and eager owners, with
//! read-back, untouched equations kept, and typed refusals that leave the
//! input intact.
#![cfg(feature = "ppt")]

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

/// An Equation Editor 3.0 object PowerPoint saved, on slide 16 (Apache
/// POI `slideshow/42520.ppt`).
const FIXTURE: &[u8] = include_bytes!("../../../test-data/poi/test-data/slideshow/42520.ppt");
const EQUATION: &str = "/slide[16]/equation[1]";

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

fn equation_paths(file: &File) -> Vec<String> {
    file.query("equation")
        .unwrap()
        .map(|node| node.unwrap().view().path().to_string())
        .collect()
}

#[test]
fn the_powerpoint_equation_is_read() {
    for file in owners(FIXTURE) {
        assert_eq!(equation_paths(&file), [EQUATION]);
        assert_eq!(
            prop(&file, EQUATION, "progId").as_deref(),
            Some("Equation.3")
        );
        let mtef = prop(&file, EQUATION, "mtef").unwrap();
        assert!(mtef.starts_with("HAAAAAIA"), "{mtef}");
        assert!(prop(&file, EQUATION, "x").is_some());
        #[cfg(feature = "equation")]
        assert_eq!(
            prop(&file, EQUATION, "latex").as_deref(),
            Some("y= 0.317+ (- .008•DSO)")
        );
        #[cfg(not(feature = "equation"))]
        assert_eq!(prop(&file, EQUATION, "latex"), None);
        // The frame is still an OLE node, after the slide's chart.
        assert_eq!(
            prop(&file, "/slide[16]/ole[2]", "progId").as_deref(),
            Some("Equation.3")
        );
    }
}

#[cfg(not(feature = "equation"))]
#[test]
fn writes_need_the_equation_feature() {
    for file in owners(FIXTURE) {
        let error = file.set(EQUATION, props! { "latex" => "x" }).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file.remove(EQUATION).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .add("/slide[1]", NewNode::new("equation").prop("latex", "x"))
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }
}

#[cfg(feature = "equation")]
fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[cfg(feature = "equation")]
#[test]
fn an_equation_is_added_to_a_slide() {
    for file in owners(FIXTURE) {
        let before = file.to_bytes().unwrap();
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("equation")
                    .prop("latex", r"\sqrt{x^2+1}")
                    .prop("x", "1in")
                    .prop("y", "2in")
                    .prop("size", "28pt"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/equation[1]");
        assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
        let read = reopen(&added.file);
        assert_eq!(
            prop(&read, "/slide[1]/equation[1]", "latex").as_deref(),
            Some(r"\sqrt{x^{2}+ 1}")
        );
        assert_eq!(
            prop(&read, "/slide[1]/equation[1]", "progId").as_deref(),
            Some("Equation.DSMT4")
        );
        assert_eq!(
            prop(&read, "/slide[1]/equation[1]", "x").as_deref(),
            Some("2.54cm")
        );
        assert_eq!(
            prop(&read, "/slide[1]/equation[1]", "y").as_deref(),
            Some("5.08cm")
        );
        // The PowerPoint equation keeps its bytes.
        assert_eq!(prop(&read, EQUATION, "mtef"), prop(&file, EQUATION, "mtef"));

        // Raw MTEF 3 makes an Equation Editor 3.0 object, stored as given.
        let mtef = prop(&file, EQUATION, "mtef").unwrap();
        let added = read
            .add(
                "/slide[2]",
                NewNode::new("math").prop("mtef", mtef.as_str()),
            )
            .unwrap();
        let read = reopen(&added.file);
        assert_eq!(prop(&read, "/slide[2]/equation[1]", "mtef"), Some(mtef));
        assert_eq!(
            prop(&read, "/slide[2]/equation[1]", "progId").as_deref(),
            Some("Equation.3")
        );
        // OMML converts too.
        let omml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>q</m:t></m:r></m:oMath>"#;
        let added = read
            .add("/slide[2]", NewNode::new("equation").prop("omml", omml))
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[2]/equation[2]");
        assert_eq!(
            prop(&reopen(&added.file), "/slide[2]/equation[2]", "latex").as_deref(),
            Some("q")
        );
    }
}

#[cfg(feature = "equation")]
#[test]
fn the_powerpoint_equation_is_replaced_keeping_its_kind() {
    for file in owners(FIXTURE) {
        let x = prop(&file, EQUATION, "x");
        let changed = file
            .set(EQUATION, props! { "latex" => r"\frac{1}{2}" })
            .unwrap();
        let read = reopen(&changed);
        assert_eq!(
            prop(&read, EQUATION, "latex").as_deref(),
            Some(r"\frac{1}{2}")
        );
        assert_eq!(
            prop(&read, EQUATION, "progId").as_deref(),
            Some("Equation.3")
        );
        assert_eq!(prop(&read, EQUATION, "x"), x);
        // Raw MTEF of another version is refused.
        let mathtype = {
            let added = read
                .add("/slide[1]", NewNode::new("equation").prop("latex", "m"))
                .unwrap();
            prop(&reopen(&added.file), "/slide[1]/equation[1]", "mtef").unwrap()
        };
        let error = read
            .set(EQUATION, props! { "mtef" => mathtype.as_str() })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
    }
}

#[cfg(feature = "equation")]
#[test]
fn the_powerpoint_equation_is_removed() {
    for file in owners(FIXTURE) {
        let before = file.to_bytes().unwrap();
        let removed = file.remove(EQUATION).unwrap();
        assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
        let read = reopen(&removed);
        assert!(equation_paths(&read).is_empty());
        assert!(read.get("/slide[16]/ole[2]").is_err());
        assert_eq!(
            prop(&read, "/slide[16]/ole[1]", "progId").as_deref(),
            Some("MSGraph.Chart.8")
        );
        assert_eq!(
            read.get("/slide[1]").unwrap().view().path().to_string(),
            "/slide[1]"
        );
    }
}

#[cfg(feature = "equation")]
#[test]
fn bad_equations_are_refused_with_typed_errors() {
    for file in owners(FIXTURE) {
        let before = file.to_bytes().unwrap();
        let malformed = file
            .set(EQUATION, props! { "latex" => r"\frac{" })
            .unwrap_err();
        assert_eq!(malformed.code(), Code::InvalidValue);
        assert_eq!(malformed.property(), Some("latex"));
        let not_base64 = file.set(EQUATION, props! { "mtef" => "%%" }).unwrap_err();
        assert_eq!(not_base64.code(), Code::InvalidValue);
        let not_mtef = file.set(EQUATION, props! { "mtef" => "AAAA" }).unwrap_err();
        assert_eq!(not_mtef.code(), Code::InvalidValue);
        let two = file
            .set(EQUATION, props! { "latex" => "x", "omml" => "<m:oMath/>" })
            .unwrap_err();
        assert_eq!(two.code(), Code::InvalidArgument);
        let none = file
            .add("/slide[1]", NewNode::new("equation").prop("x", "1in"))
            .unwrap_err();
        assert_eq!(none.code(), Code::MissingProperty);
        let wrong_parent = file
            .add("/", NewNode::new("equation").prop("latex", "x"))
            .unwrap_err();
        assert_ne!(wrong_parent.code(), Code::InternalError);
        let set_position = file.set(EQUATION, props! { "x" => "1in" }).unwrap_err();
        assert_ne!(set_position.code(), Code::InternalError);
        assert_eq!(file.to_bytes().unwrap(), before);
    }
}

/// The fixture with bytes of its equation storage's zlib stream flipped.
fn corrupt_fixture() -> Vec<u8> {
    let mut file =
        litchi_cfb::OleFile::open(std::io::Cursor::new(FIXTURE), Default::default()).unwrap();
    let document = file.open_stream(&["PowerPoint Document"]).unwrap();
    // The fixture's equation storage is the `ExOleObjStg` at 62126.
    let window = &document[62126 + 60..62126 + 92];
    let mut bytes = FIXTURE.to_vec();
    let at = bytes
        .windows(window.len())
        .position(|candidate| candidate == window)
        .unwrap();
    for byte in &mut bytes[at..at + window.len()] {
        *byte ^= 0x5A;
    }
    bytes
}

#[test]
fn a_corrupt_equation_storage_fails_only_its_equation() {
    let names = |file: &File| -> Vec<String> {
        file.query("/slide[16]/*")
            .unwrap()
            .map(|node| node.unwrap().view().path().to_string())
            .collect()
    };
    for (file, corrupt) in owners(FIXTURE).into_iter().zip(owners(&corrupt_fixture())) {
        assert_eq!(names(&corrupt), names(&file));
        assert_eq!(
            prop(&corrupt, EQUATION, "progId").as_deref(),
            Some("Equation.3")
        );
        assert_eq!(prop(&corrupt, EQUATION, "mtef"), None);
        assert_eq!(prop(&corrupt, EQUATION, "latex"), None);
        assert!(prop(&corrupt, EQUATION, "error").is_some());
        assert_eq!(prop(&file, EQUATION, "error"), None);
        assert_eq!(equation_paths(&corrupt), [EQUATION]);
    }
}
