//! PPTX OLE equation objects (MathType, Equation Editor 3.0) through the
//! unified API: listed as `equation` nodes after a slide's Office Math
//! equations, added, replaced and removed for source and eager owners, with
//! read-back through the leaf reader and MTEF, untouched parts kept, and
//! typed refusals that leave the input intact.
#![cfg(feature = "pptx")]

#[cfg(feature = "equation")]
use std::collections::BTreeMap;
#[cfg(feature = "equation")]
use std::io::{Cursor, Read as _, Write as _};

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

/// PowerPoint 2013: a MathType frame in the VML-choice form with a WMF
/// preview, a legacy VML drawing and an animation of the frame.
const FIXTURE: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf169088.pptx");
#[cfg(feature = "equation")]
const HANDOUT: &[u8] = include_bytes!(
    "../../../test-data/microsoft-office/powerpoint/powerpoint-fonts-math-handout.pptx"
);
#[cfg(feature = "equation")]
const SLIDE: &str = "ppt/slides/slide1.xml";

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
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open PPTX ZIP");
    (0..archive.len())
        .map(|index| {
            let mut member = archive.by_index(index).expect("member");
            let mut contents = Vec::new();
            member.read_to_end(&mut contents).expect("read member");
            (member.name().expect("member name").into_owned(), contents)
        })
        .collect()
}

#[cfg(feature = "equation")]
/// `bytes` with the member `name` passed through `change`.
fn rewritten(bytes: &[u8], name: &str, change: impl FnOnce(String) -> String) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let mut change = Some(change);
    for (member, contents) in members(bytes) {
        let contents = if member == name {
            let text = String::from_utf8(contents).expect("UTF-8");
            (change.take().expect("one member"))(text).into_bytes()
        } else {
            contents
        };
        zip.start_file(member, options).expect("member");
        zip.write_all(&contents).expect("write member");
    }
    zip.finish().expect("finish").into_inner()
}

#[cfg(feature = "equation")]
/// The fixture without its animation, which refuses the frame's removal.
fn without_timing() -> Vec<u8> {
    rewritten(FIXTURE, SLIDE, |xml| {
        let start = xml.find("<p:timing>").unwrap();
        let end = xml.find("</p:timing>").unwrap() + "</p:timing>".len();
        format!("{}{}", &xml[..start], &xml[end..])
    })
}

fn equation_paths(file: &File) -> Vec<String> {
    file.query("equation")
        .unwrap()
        .map(|node| node.unwrap().view().path().to_string())
        .collect()
}

#[cfg(feature = "equation")]
fn compact(text: Option<String>) -> Option<String> {
    text.map(|text| text.chars().filter(|ch| !ch.is_whitespace()).collect())
}

#[cfg(feature = "equation")]
fn decode(mtef: &str) -> Vec<u8> {
    use base64::Engine as _;
    base64::engine::general_purpose::STANDARD
        .decode(mtef)
        .unwrap()
}

/// The OLE equations the leaf reader finds on slide 1 of `bytes`.
#[cfg(feature = "equation")]
fn leaf(bytes: &[u8]) -> Vec<litchi_pptx::shape::math::OleEquation> {
    let package = litchi_pptx::Package::from_bytes(bytes).unwrap();
    let presentation = package.presentation().unwrap();
    presentation
        .slide(0)
        .unwrap()
        .unwrap()
        .ole_equations()
        .unwrap()
}

#[test]
fn powerpoints_equation_object_is_an_equation_node() {
    for file in owners(FIXTURE) {
        assert_eq!(equation_paths(&file), ["/slide[1]/equation[1]"]);
        let path = "/slide[1]/equation[1]";
        assert_eq!(
            prop(&file, path, "progId").as_deref(),
            Some("Equation.DSMT4")
        );
        assert_eq!(prop(&file, path, "omml"), None);
        assert_eq!(prop(&file, path, "error"), None);
        // The frame is an OLE object as well.
        assert!(file.get("/slide[1]/ole[1]").is_ok());
        #[cfg(feature = "equation")]
        {
            let native = decode(&prop(&file, path, "mtef").unwrap());
            let read =
                litchi_equation::read_mtef(&native, &litchi_equation::Limits::DEFAULT).unwrap();
            assert_eq!(read.info().unwrap().version, 5);
            let latex = prop(&file, path, "latex").unwrap();
            assert_eq!(latex, read.latex());
        }
    }
}

#[cfg(feature = "equation")]
#[test]
fn an_unreadable_embedding_is_listed_with_its_error() {
    let garbled = {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (member, contents) in members(FIXTURE) {
            let contents = if member == "ppt/embeddings/oleObject1.bin" {
                b"garbage".to_vec()
            } else {
                contents
            };
            zip.start_file(member, options).unwrap();
            zip.write_all(&contents).unwrap();
        }
        zip.finish().unwrap().into_inner()
    };
    for file in owners(&garbled) {
        let path = "/slide[1]/equation[1]";
        assert_eq!(equation_paths(&file), [path]);
        assert!(prop(&file, path, "error").is_some());
        assert_eq!(prop(&file, path, "mtef"), None);
        assert_eq!(prop(&file, path, "latex"), None);
        assert_eq!(
            prop(&file, path, "progId").as_deref(),
            Some("Equation.DSMT4")
        );
    }
}

#[cfg(feature = "equation")]
#[test]
fn a_mathtype_object_is_added_after_the_slides_equations() {
    for file in owners(FIXTURE) {
        let before = file.to_bytes().unwrap();
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("equation")
                    .prop("progId", "Equation.DSMT4")
                    .prop("latex", r"\sqrt{x+1}")
                    .prop("x", "1in")
                    .prop("y", "1in")
                    .prop("size", "32pt"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/equation[2]");
        assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
        let bytes = added.file.to_bytes().unwrap();
        let read = litchi_ops::open_bytes(bytes.clone()).unwrap();
        let path = "/slide[1]/equation[2]";
        assert_eq!(
            compact(prop(&read, path, "latex")).as_deref(),
            Some(r"\sqrt{x+1}")
        );
        assert_eq!(
            prop(&read, path, "progId").as_deref(),
            Some("Equation.DSMT4")
        );
        let native = decode(&prop(&read, path, "mtef").unwrap());
        let mtef = litchi_equation::read_mtef(&native, &litchi_equation::Limits::DEFAULT).unwrap();
        assert_eq!(mtef.info().unwrap().version, 5);
        // The leaf reader sees the same object, in the plain form.
        let objects = leaf(&bytes);
        assert_eq!(objects.len(), 2);
        assert!(!objects[1].has_legacy_drawing());
        assert_eq!(&*objects[1].equation_native().unwrap(), native.as_slice());
        assert!(objects[1].preview_part().unwrap().ends_with(".png"));
        // Every member but the slide, its relationships, the manifest and
        // the two new parts keeps its bytes.
        let (a, b) = (members(&before), members(&bytes));
        let new: Vec<_> = b.keys().filter(|name| !a.contains_key(*name)).collect();
        assert_eq!(new.len(), 2, "{new:?}");
        for (name, contents) in &a {
            if ![
                SLIDE,
                "ppt/slides/_rels/slide1.xml.rels",
                "[Content_Types].xml",
            ]
            .contains(&name.as_str())
            {
                assert!(b.get(name) == Some(contents), "{name}");
            }
        }

        // Raw MTEF makes another object of the same kind.
        let copy = read
            .add(
                "/slide[1]",
                NewNode::new("equation").prop("mtef", prop(&read, path, "mtef").unwrap()),
            )
            .unwrap();
        assert_eq!(copy.path.to_string(), "/slide[1]/equation[3]");
        assert_eq!(
            prop(&copy.file, "/slide[1]/equation[3]", "mtef"),
            prop(&read, path, "mtef")
        );
    }
}

#[cfg(feature = "equation")]
#[test]
fn an_ole_equation_is_replaced_keeping_its_kind() {
    for file in owners(FIXTURE) {
        let path = "/slide[1]/equation[1]";
        let changed = file
            .set(path, props! { "latex" => r"\frac{a}{b}" })
            .unwrap();
        let bytes = changed.to_bytes().unwrap();
        let read = litchi_ops::open_bytes(bytes.clone()).unwrap();
        assert_eq!(
            compact(prop(&read, path, "latex")).as_deref(),
            Some(r"\frac{a}{b}")
        );
        assert_eq!(
            prop(&read, path, "progId").as_deref(),
            Some("Equation.DSMT4")
        );
        let saved = members(&bytes);
        // The old WMF preview and the VML drawing that showed it are gone.
        assert!(!saved.contains_key("ppt/media/image1.wmf"));
        assert!(!saved.contains_key("ppt/drawings/vmlDrawing1.vml"));
        assert!(saved.contains_key("ppt/embeddings/oleObject1.bin"));

        // OMML and raw MTEF replace it too.
        let omml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>k</m:t></m:r></m:oMath>"#;
        let changed = read.set(path, props! { "omml" => omml }).unwrap();
        assert_eq!(prop(&changed, path, "latex").as_deref(), Some("k"));
        assert_eq!(prop(&changed, path, "omml"), None);
        let mtef = prop(&file, path, "mtef").unwrap();
        let changed = changed
            .set(path, props! { "mtef" => mtef.as_str() })
            .unwrap();
        assert_eq!(prop(&changed, path, "mtef").as_deref(), Some(mtef.as_str()));
        assert_eq!(prop(&changed, path, "latex"), prop(&file, path, "latex"));
    }
}

#[cfg(feature = "equation")]
#[test]
fn a_removed_object_leaves_nothing_behind() {
    let input = without_timing();
    for file in owners(&input) {
        let native = decode(&prop(&file, "/slide[1]/equation[1]", "mtef").unwrap());
        let changed = file.remove("/slide[1]/equation[1]").unwrap();
        assert!(equation_paths(&changed).is_empty());
        let bytes = changed.to_bytes().unwrap();
        let saved = members(&bytes);
        for gone in [
            "ppt/embeddings/oleObject1.bin",
            "ppt/media/image1.wmf",
            "ppt/drawings/vmlDrawing1.vml",
            "ppt/drawings/_rels/vmlDrawing1.vml.rels",
        ] {
            assert!(!saved.contains_key(gone), "{gone}");
        }
        let types = String::from_utf8(saved["[Content_Types].xml"].clone()).unwrap();
        for extension in ["bin", "wmf", "vml"] {
            assert!(
                !types.contains(&format!(r#"Extension="{extension}""#)),
                "{extension}: {types}"
            );
        }
        let mtef = &native[28..];
        assert!(
            saved
                .values()
                .all(|member| !member.windows(mtef.len()).any(|window| window == mtef)),
            "the removed equation's MTEF survives"
        );
        // Everything else keeps its bytes.
        let before = members(&input);
        for (name, contents) in &saved {
            if ![
                SLIDE,
                "ppt/slides/_rels/slide1.xml.rels",
                "[Content_Types].xml",
            ]
            .contains(&name.as_str())
            {
                assert!(before.get(name) == Some(contents), "{name}");
            }
        }
    }
}

#[cfg(feature = "equation")]
#[test]
fn bad_ole_equation_writes_are_typed_and_leave_the_input_intact() {
    for file in owners(FIXTURE) {
        let before = file.to_bytes().unwrap();
        let path = "/slide[1]/equation[1]";
        // An animation targets the frame.
        assert!(file.remove(path).is_err());
        let error = file
            .add(
                "/slide[1]",
                NewNode::new("equation")
                    .prop("progId", "Equation.3")
                    .prop("latex", "x"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("progId"));
        let error = file
            .add(
                "/slide[1]",
                NewNode::new("equation").prop("mtef", "not base64!"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .add(
                "/slide[1]",
                NewNode::new("equation")
                    .prop("mtef", "AAAA")
                    .prop("latex", "x"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        assert!(
            file.add("/slide[1]", NewNode::new("equation").prop("mtef", "AAAA"))
                .is_err()
        );
        assert!(
            file.set(path, props! { "progId" => "Equation.DSMT4" })
                .is_err()
        );
        assert!(file.set(path, props! { "latex" => r"\frac{" }).is_err());
        assert_eq!(file.to_bytes().unwrap(), before);
    }
    // MTEF does not belong to an Office Math equation.
    for file in owners(HANDOUT) {
        let error = file
            .set("/slide[2]/equation[1]", props! { "mtef" => "AAAA" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        assert_eq!(prop(&file, "/slide[2]/equation[1]", "progId"), None);
        assert_eq!(prop(&file, "/slide[2]/equation[1]", "mtef"), None);
    }
}

#[cfg(not(feature = "equation"))]
#[test]
fn ole_equation_writes_need_the_equation_feature() {
    for file in owners(FIXTURE) {
        let path = "/slide[1]/equation[1]";
        let error = file.set(path, props! { "latex" => "x" }).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file.remove(path).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .add(
                "/slide[1]",
                NewNode::new("equation")
                    .prop("progId", "Equation.DSMT4")
                    .prop("latex", "x"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }
}

#[test]
fn capabilities_report_ole_equation_properties() {
    let capabilities = litchi_ops::capabilities::Capabilities::of(litchi_core::FileFormat::Pptx);
    let equation = capabilities
        .elements()
        .find(|element| element.name() == "equation")
        .expect("equation element");
    let mtef = equation.property("mtef").expect("mtef");
    assert!(mtef.get() && mtef.set() && mtef.add());
    let prog_id = equation.property("progId").expect("progId");
    assert!(prog_id.get() && !prog_id.set() && prog_id.add());
    let error = equation.property("error").expect("error");
    assert!(error.get() && !error.set() && !error.add());
}
