//! PPTX equations through the unified API: add, set and remove in
//! PowerPoint's form, for source and eager owners, with read-back, untouched
//! equations kept, and typed refusals that leave the input intact.
#![cfg(feature = "pptx")]

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

const HANDOUT: &[u8] = include_bytes!(
    "../../../test-data/microsoft-office/powerpoint/powerpoint-fonts-math-handout.pptx"
);
#[cfg(feature = "equation")]
const MCE: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/oox/qa/unit/data/import-mce.pptx");

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

#[cfg(feature = "equation")]
/// The OMML of the `index`-th equation of slide `slide` (both from 1).
fn omml(file: &File, slide: usize, index: usize) -> String {
    prop(file, &format!("/slide[{slide}]/equation[{index}]"), "omml").unwrap_or_default()
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

#[cfg(feature = "equation")]
fn equation_paths(file: &File) -> Vec<String> {
    file.query("equation")
        .unwrap()
        .map(|node| node.unwrap().view().path().to_string())
        .collect()
}

#[cfg(not(feature = "equation"))]
#[test]
fn writes_need_the_equation_feature() {
    for file in owners(HANDOUT) {
        let error = file
            .set("/slide[2]/equation[1]", props! { "latex" => "x" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file.remove("/slide[2]/equation[1]").unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .add("/slide[1]", NewNode::new("equation").prop("latex", "x"))
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }
}

#[cfg(feature = "equation")]
#[test]
fn equations_are_added_to_a_slide_as_text_boxes() {
    for file in owners(HANDOUT) {
        let before = file.to_bytes().unwrap();
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("equation")
                    .prop("latex", r"\int_0^1 x\,dx")
                    .prop("x", "1in")
                    .prop("y", "2in")
                    .prop("size", "24pt"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/equation[1]");
        assert_eq!(file.to_bytes().unwrap(), before, "input unchanged");
        let read = reopen(&added.file);
        assert_eq!(
            compact(prop(&read, "/slide[1]/equation[1]", "latex")).as_deref(),
            Some(r"\int_0^1x\,dx")
        );
        let written = omml(&read, 1, 1);
        assert!(written.starts_with("<m:oMathPara "), "{written}");
        assert!(written.contains(r#"sz="2400""#), "{written}");
        // The equation is reachable through its new shape as well.
        let paths = equation_paths(&read);
        assert_eq!(paths.len(), 4, "{paths:?}");
        assert!(paths[0].starts_with("/slide[1]/shape["), "{paths:?}");
        // OMML is written as given.
        let raw = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>q</m:t></m:r></m:oMath>"#;
        let added = read
            .add("/slide[1]", NewNode::new("math").prop("omml", raw))
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/equation[2]");
        let read = reopen(&added.file);
        assert_eq!(omml(&read, 1, 2), raw);
    }
}

#[cfg(feature = "equation")]
#[test]
fn a_powerpoint_equation_is_replaced_and_the_others_keep_their_bytes() {
    for file in owners(HANDOUT) {
        let others = [omml(&file, 3, 1), omml(&file, 4, 1)];
        let changed = file
            .set(
                "/slide[2]/equation[1]",
                props! { "latex" => r"\frac{1}{2}" },
            )
            .unwrap();
        let read = reopen(&changed);
        assert_eq!(
            compact(prop(&read, "/slide[2]/equation[1]", "latex")).as_deref(),
            Some(r"\frac{1}{2}")
        );
        assert_eq!([omml(&read, 3, 1), omml(&read, 4, 1),], others);
        // Through the shape path, with OMML.
        let raw = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>w</m:t></m:r></m:oMath>"#;
        let shape = equation_paths(&read)[0].clone();
        assert!(shape.starts_with("/slide[2]/shape["), "{shape}");
        let changed = read.set(shape.as_str(), props! { "omml" => raw }).unwrap();
        let read = reopen(&changed);
        assert_eq!(omml(&read, 2, 1), raw);
        assert_eq!(
            prop(&read, "/slide[2]/equation[1]", "latex").as_deref(),
            Some("w")
        );
        // The same content again is an exact no-op.
        let bytes = read.to_bytes().unwrap();
        let same = read
            .set("/slide[2]/equation[1]", props! { "omml" => raw })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
    }
}

#[cfg(feature = "equation")]
#[test]
fn an_equation_is_removed_with_its_shape() {
    for file in owners(HANDOUT) {
        let before = equation_paths(&file);
        let removed = file.remove("/slide[3]/equation[1]").unwrap();
        let read = reopen(&removed);
        let after = equation_paths(&read);
        assert_eq!(after.len() + 1, before.len(), "{after:?}");
        assert!(after.iter().all(|path| !path.starts_with("/slide[3]/")));
        assert!(read.get("/slide[3]/equation[1]").is_err());
        assert_eq!(omml(&read, 4, 1), omml(&file, 4, 1));
    }
}

#[cfg(feature = "equation")]
#[test]
fn an_autofit_text_box_equation_round_trips() {
    for file in owners(MCE) {
        let changed = file
            .set(
                "/slide[1]/equation[1]",
                props! { "latex" => r"a^2+b^2=c^2" },
            )
            .unwrap();
        let read = reopen(&changed);
        assert_eq!(
            compact(prop(&read, "/slide[1]/shape[1]/equation[1]", "latex")).as_deref(),
            Some("a^{2}+b^{2}=c^{2}")
        );
        let removed = read.remove("/slide[1]/shape[1]/equation[1]").unwrap();
        assert!(equation_paths(&reopen(&removed)).is_empty());
    }
}

#[cfg(feature = "equation")]
#[test]
fn bad_equation_writes_are_typed_and_leave_the_input_intact() {
    for file in owners(HANDOUT) {
        let bytes = file.to_bytes().unwrap();
        let code = |result: litchi_ops::Result<litchi_ops::Added>| result.unwrap_err().code();
        assert_eq!(
            file.set(
                "/slide[2]/equation[1]",
                props! { "latex" => "x", "omml" => "<m:oMath/>" }
            )
            .unwrap_err()
            .code(),
            Code::InvalidArgument
        );
        assert_eq!(
            file.set("/slide[2]/equation[1]", props! { "latex" => r"\frac{" })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        assert_eq!(
            file.set("/slide[2]/equation[1]", props! { "latex" => "" })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        assert!(
            file.set(
                "/slide[2]/equation[1]",
                props! { "omml" => "<m:oMath><m:r>" }
            )
            .is_err()
        );
        assert!(
            file.set("/slide[2]/equation[9]", props! { "latex" => "x" })
                .is_err()
        );
        assert_eq!(
            code(file.add(
                "/slide[1]/shape[1]",
                NewNode::new("equation").prop("latex", "x").prop("x", "1in")
            )),
            Code::UnsupportedProperty
        );
        assert_eq!(
            code(file.add("/", NewNode::new("equation").prop("latex", "x"))),
            Code::UnsupportedType
        );
        // Caller OMML cannot refer to relationships the slide does not have.
        let linked = r#"<m:oMath><m:r><a:rPr><a:hlinkClick r:id="rId77"/></a:rPr><m:t>x</m:t></m:r></m:oMath>"#;
        assert!(
            file.set("/slide[2]/equation[1]", props! { "omml" => linked })
                .is_err()
        );
        assert_eq!(
            code(file.add("/slide[1]", NewNode::new("equation"))),
            Code::MissingProperty
        );
        assert_eq!(
            code(
                file.add(
                    "/slide[1]",
                    NewNode::new("equation")
                        .prop("latex", "x")
                        .prop("size", "0pt")
                )
            ),
            Code::InvalidValue
        );
        assert_eq!(file.to_bytes().unwrap(), bytes);
    }
}

#[cfg(feature = "equation")]
#[test]
fn an_equation_beside_text_is_edited_with_a_text_fallback() {
    // A text box holding text and an equation, in PowerPoint's block.
    let slide = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>where </a:t></a:r><a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></a14:m></a:p></p:txBody></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#;
    let mut package =
        litchi_opc::OpcPackage::from_bytes(HANDOUT, litchi_opc::SourceLimits::default()).unwrap();
    package
        .get_part_mut(&litchi_opc::PackURI::new("/ppt/slides/slide2.xml").unwrap())
        .unwrap()
        .set_blob(slide.as_bytes().to_vec());
    let mut bytes = Vec::new();
    package.write_to(std::io::Cursor::new(&mut bytes)).unwrap();
    for file in owners(&bytes) {
        let changed = file
            .set("/slide[2]/equation[1]", props! { "latex" => "y" })
            .unwrap();
        let read = reopen(&changed);
        assert_eq!(
            prop(&read, "/slide[2]/equation[1]", "latex").as_deref(),
            Some("y")
        );
        let removed = read.remove("/slide[2]/equation[1]").unwrap();
        let read = reopen(&removed);
        assert!(read.get("/slide[2]/equation[1]").is_err());
        // The shape stays with its text.
        assert!(
            read.query("shape").unwrap().any(|node| node
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref()
                == Some("where "))
        );
    }
}

#[test]
fn capabilities_report_equation_writes() {
    let capabilities = litchi_ops::capabilities::Capabilities::of(FileFormat::Pptx);
    let equation = capabilities
        .elements()
        .find(|element| element.name() == "equation")
        .expect("equation element");
    let latex = equation.property("latex").expect("latex");
    assert!(latex.get() && latex.set() && latex.add());
    for key in ["x", "y", "size"] {
        let property = equation.property(key).expect(key);
        assert!(
            property.add() && !property.set() && !property.get(),
            "{key}"
        );
    }
}

#[cfg(feature = "equation")]
#[test]
fn an_equation_is_added_to_a_shapes_text_and_edited_there() {
    for file in owners(HANDOUT) {
        // The subtitle of slide 1 is ordinary text.
        let added = file
            .add(
                "/slide[1]/shape[2]",
                NewNode::new("equation").prop("latex", r"e^{i\pi}"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/equation[1]");
        let read = reopen(&added.file);
        let paths = equation_paths(&read);
        assert!(
            paths[0].starts_with("/slide[1]/shape[2]/equation[1]"),
            "{paths:?}"
        );
        assert!(omml(&read, 1, 1).starts_with("<m:oMath "));
        // Replace it, then remove it: the shape keeps its text.
        let changed = read
            .set("/slide[1]/equation[1]", props! { "latex" => "x" })
            .unwrap();
        let read = reopen(&changed);
        assert_eq!(
            prop(&read, "/slide[1]/equation[1]", "latex").as_deref(),
            Some("x")
        );
        let removed = read.remove("/slide[1]/equation[1]").unwrap();
        let read = reopen(&removed);
        assert!(read.get("/slide[1]/equation[1]").is_err());
        assert!(read.get("/slide[1]/shape[2]").is_ok());
    }
}
