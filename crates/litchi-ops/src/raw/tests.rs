//! Raw-layer unit tests (record 0336).

#![allow(clippy::unwrap_used, clippy::expect_used, reason = "tests")]

#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "odt"))]
use super::Action;
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx"))]
use super::NewPart;
use crate::Code;
use litchi_core::FileFormat;

#[cfg(feature = "docx")]
#[test]
fn docx_parts_and_aliases() {
    let file = crate::create(FileFormat::Docx).unwrap();
    let parts = file.parts().unwrap();
    assert!(parts.iter().any(|part| part.name() == "/word/document.xml"));
    assert!(parts.windows(2).all(|pair| pair[0].name() < pair[1].name()));
    let document = file.raw("/document").unwrap();
    assert_eq!(document.name(), "/word/document.xml");
    assert!(document.content_type().contains("document.main+xml"));
    assert!(document.is_xml());
    assert_eq!(file.raw("/word/document.xml").unwrap(), document);
    assert_eq!(file.raw("/styles").unwrap().name(), "/word/styles.xml");
    assert_eq!(file.raw("/nothing").unwrap_err().code(), Code::PathNotFound);
    assert_eq!(file.raw("document").unwrap_err().code(), Code::InvalidPath);
    assert_eq!(file.raw("/").unwrap_err().code(), Code::InvalidPath);
}

#[cfg(feature = "docx")]
#[test]
fn docx_raw_set_actions() {
    let file = crate::create(FileFormat::Docx).unwrap();
    let edited = file
        .raw_set(
            "/document",
            "//w:body",
            Action::Prepend,
            Some("\n  <w:p>\n    <w:r><w:t xml:space=\"preserve\">Hello </w:t></w:r>\n  </w:p>\n"),
        )
        .unwrap();
    assert_eq!(edited.matched, 1);
    assert_eq!(edited.file.text().unwrap().trim_end(), "Hello");
    // The original handle is unchanged.
    assert_eq!(file.text().unwrap().trim(), "");
    let xml = edited.file.raw("/document").unwrap();
    assert!(
        xml.text()
            .unwrap()
            .contains("<w:body><w:p><w:r><w:t xml:space=\"preserve\">Hello </w:t></w:r></w:p>")
    );

    let second = edited
        .file
        .raw_set(
            "/document",
            "/w:document/w:body/w:p[1]",
            Action::InsertAfter,
            Some("<w:p><w:r><w:t>World</w:t></w:r></w:p>"),
        )
        .unwrap();
    let text = second.file.text().unwrap();
    assert!(text.contains("Hello") && text.contains("World"), "{text}");

    let last = second
        .file
        .raw_set("/document", "//w:body/w:p[last()]", Action::Remove, None)
        .unwrap();
    assert!(!last.file.text().unwrap().contains("World"));

    let styled = edited
        .file
        .raw_set(
            "/document",
            "//w:p",
            Action::Prepend,
            Some("<w:pPr><w:jc w:val=\"center\"/></w:pPr>"),
        )
        .unwrap();
    let centered = styled
        .file
        .raw_set(
            "/document",
            "//w:jc[@w:val='center']",
            Action::SetAttr,
            Some("w:val=right"),
        )
        .unwrap();
    assert_eq!(centered.matched, 1);
    let xml = centered.file.raw("/document").unwrap();
    assert!(xml.text().unwrap().contains("<w:jc w:val=\"right\"/>"));
    // Setting the same value again is an exact no-op on the same handle.
    let same = centered
        .file
        .raw_set("/document", "//w:jc", Action::SetAttr, Some("w:val=right"))
        .unwrap();
    assert_eq!(
        same.file.to_bytes().unwrap(),
        centered.file.to_bytes().unwrap()
    );
    assert!(centered.file.validate().unwrap().is_valid());
}

#[cfg(feature = "docx")]
#[test]
fn docx_raw_set_refusals() {
    let file = crate::create(FileFormat::Docx).unwrap();
    let code = |xpath: &str, action: Action, xml: Option<&str>| {
        file.raw_set("/document", xpath, action, xml)
            .unwrap_err()
            .code()
    };
    assert_eq!(code("//w:tbl", Action::Remove, None), Code::ZeroMatches);
    assert_eq!(code("w:body", Action::Remove, None), Code::InvalidXpath);
    assert_eq!(code("//q:body", Action::Remove, None), Code::InvalidXpath);
    assert_eq!(
        code("//w:body[text()]", Action::Remove, None),
        Code::InvalidXpath
    );
    assert_eq!(
        code("/w:document", Action::Remove, None),
        Code::InvalidArgument
    );
    assert_eq!(code("//*", Action::Remove, None), Code::InvalidArgument);
    assert_eq!(
        code("//w:body", Action::Append, None),
        Code::MissingArgument
    );
    assert_eq!(
        code("//w:body", Action::Append, Some("<w:p>")),
        Code::InvalidArgument
    );
    assert_eq!(
        code("//w:body", Action::Append, Some("text")),
        Code::InvalidArgument
    );
    assert_eq!(
        code("//w:body", Action::Append, Some("<zz:p/>")),
        Code::InvalidArgument
    );
    assert_eq!(
        code("//w:body", Action::SetAttr, Some("novalue")),
        Code::InvalidArgument
    );
    assert_eq!("delete".parse::<Action>().unwrap(), Action::Remove);
    assert_eq!("After".parse::<Action>().unwrap(), Action::InsertAfter);
    assert_eq!(
        "frob".parse::<Action>().unwrap_err().code(),
        Code::InvalidArgument
    );
}

#[cfg(feature = "docx")]
#[test]
fn docx_add_part() {
    let file = crate::create(FileFormat::Docx).unwrap();
    let added = file
        .add_part("/document", NewPart::of_type("header").unwrap())
        .unwrap();
    assert_eq!(added.part, "/word/header1.xml");
    let header = added.file.raw("/header[1]").unwrap();
    assert_eq!(header.name(), "/word/header1.xml");
    assert!(header.content_type().ends_with("header+xml"));
    let again = added
        .file
        .add_part("/document", NewPart::of_type("header").unwrap())
        .unwrap();
    assert_eq!(again.part, "/word/header2.xml");
    assert_ne!(again.relationship_id, added.relationship_id);
    assert!(again.file.validate().unwrap().is_valid());

    let custom = file
        .add_part(
            "/",
            NewPart::new(
                "application/xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
                "<root/>",
            )
            .named("/customXml/item9.xml"),
        )
        .unwrap();
    assert_eq!(
        custom.file.raw("/customXml/item9.xml").unwrap().bytes(),
        b"<root/>"
    );
    let taken = custom.file.add_part(
        "/",
        NewPart::new("application/xml", "urn:x", "<root/>").named("/customXml/item9.xml"),
    );
    assert_eq!(taken.unwrap_err().code(), Code::FileExists);
    let unnamed = file.add_part("/", NewPart::new("application/xml", "urn:x", "<root/>"));
    assert_eq!(unnamed.unwrap_err().code(), Code::MissingArgument);
    let broken = file.add_part(
        "/",
        NewPart::new("application/xml", "urn:x", "<root>").named("/customXml/item8.xml"),
    );
    assert!(broken.is_err());
    assert_eq!(
        NewPart::of_type("table").unwrap_err().code(),
        Code::UnsupportedType
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_sheet_aliases() {
    let file = crate::create(FileFormat::Xlsx).unwrap();
    let sheet = file.raw("/Sheet1").unwrap();
    assert!(sheet.content_type().ends_with("worksheet+xml"));
    assert_eq!(file.raw("/sheet1").unwrap().name(), sheet.name());
    assert_eq!(file.raw("/sheet[1]").unwrap().name(), sheet.name());
    assert_eq!(file.raw("/workbook").unwrap().name(), "/xl/workbook.xml");
    // A sheet without a drawing has nowhere to relate a chart from.
    let chart = file.add_part("/Sheet1", NewPart::of_type("chart").unwrap());
    assert_eq!(chart.unwrap_err().code(), Code::UnsupportedCapability);
    let header = file.add_part("/Sheet1", NewPart::of_type("header").unwrap());
    assert_eq!(header.unwrap_err().code(), Code::UnsupportedType);
    let edited = file
        .raw_set("/Sheet1", "//x:sheetData", Action::Append, Some(r#"<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>raw</x:t></x:is></x:c></x:row>"#))
        .unwrap();
    assert_eq!(
        edited
            .file
            .get("/Sheet1/A1")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("raw")
    );
    assert!(edited.file.validate().unwrap().is_valid());
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_slide_aliases_and_append() {
    let file = crate::create(FileFormat::Pptx).unwrap();
    let file = file.set_command_free_slide().unwrap();
    let slide = file.raw("/slide[1]").unwrap();
    assert!(slide.content_type().ends_with("slide+xml"));
    assert!(file.raw("/slideMaster[1]").is_ok());
    assert!(file.raw("/slideLayout[1]").is_ok());
    assert!(file.raw("/presentation").is_ok());
    let shapes = |file: &crate::File| file.get("/slide[1]").unwrap().children().count();
    let before = shapes(&file);
    let edited = file
        .raw_set(
            "/slide[1]",
            "//p:cSld/p:spTree",
            Action::Append,
            Some(r#"<p:sp><p:nvSpPr><p:cNvPr id="99" name="Raw"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>raw shape</a:t></a:r></a:p></p:txBody></p:sp>"#),
        )
        .unwrap();
    assert_eq!(shapes(&edited.file), before + 1);
    assert!(edited.file.text().unwrap().contains("raw shape"));
    assert!(edited.file.validate().unwrap().is_valid());
    let chart = edited
        .file
        .add_part("/slide[1]", NewPart::of_type("chart").unwrap())
        .unwrap();
    assert_eq!(chart.part, "/ppt/charts/chart1.xml");
    assert_eq!(
        chart.file.raw("/slide[1]/chart[1]").unwrap().name(),
        chart.part
    );
}

#[cfg(feature = "pptx")]
impl crate::File {
    /// A deck with one blank slide.
    fn set_command_free_slide(&self) -> crate::Result<Self> {
        self.add(
            "/",
            crate::NewNode::new("slide").with_props(crate::props! { "layout" => "Blank" })?,
        )
        .map(|added| added.file)
    }
}

#[cfg(feature = "odt")]
#[test]
fn odf_raw_reads_and_refused_writes() {
    let file = crate::create(FileFormat::Odt).unwrap();
    let content = file.raw("/content").unwrap();
    assert_eq!(content.name(), "/content.xml");
    assert_eq!(content.content_type(), "text/xml");
    assert!(file.raw("/META-INF/manifest.xml").is_ok());
    assert!(
        file.parts()
            .unwrap()
            .iter()
            .any(|part| part.name() == "/mimetype")
    );
    let refused = file
        .raw_set("/content", "//*", Action::Remove, None)
        .unwrap_err();
    assert_eq!(refused.code(), Code::UnsupportedCapability);
    assert_eq!(refused.format(), Some(FileFormat::Odt));
    assert!(file.validate().unwrap().is_valid());
}

#[cfg(feature = "rtf")]
#[test]
fn rtf_has_no_parts_but_validates() {
    let file = crate::create(FileFormat::Rtf).unwrap();
    let refused = file.raw("/document").unwrap_err();
    assert_eq!(refused.code(), Code::UnsupportedCapability);
    assert_eq!(refused.format(), Some(FileFormat::Rtf));
    assert!(file.validate().unwrap().is_valid());
}

#[cfg(all(feature = "doc", feature = "xls", feature = "ppt"))]
#[test]
fn legacy_formats_validate() {
    for format in [FileFormat::Doc, FileFormat::Xls, FileFormat::Ppt] {
        let file = crate::create(format).unwrap();
        let validation = file.validate().unwrap();
        assert_eq!(validation.format(), format);
        assert!(
            validation.is_valid(),
            "{format:?}: {:?}",
            validation.report()
        );
    }
}

#[test]
fn raw_layer_verbs_answer_by_package() {
    use crate::capabilities::{Capabilities, Verb};
    for format in [
        FileFormat::Docx,
        FileFormat::Xlsx,
        FileFormat::Pptx,
        FileFormat::Xlsb,
        FileFormat::Odt,
        FileFormat::Ods,
        FileFormat::Odp,
        FileFormat::Doc,
        FileFormat::Xls,
        FileFormat::Ppt,
        FileFormat::Rtf,
    ] {
        let capabilities = Capabilities::of(format);
        if !capabilities.verb(Verb::Get).is_supported() {
            continue;
        }
        let ooxml = matches!(
            format,
            FileFormat::Docx | FileFormat::Xlsx | FileFormat::Pptx | FileFormat::Xlsb
        );
        let odf = matches!(format, FileFormat::Odt | FileFormat::Ods | FileFormat::Odp);
        assert_eq!(
            capabilities.verb(Verb::Raw).is_supported(),
            ooxml || odf,
            "{format:?}"
        );
        assert_eq!(
            capabilities.verb(Verb::AddPart).is_supported(),
            ooxml,
            "{format:?}"
        );
        assert_eq!(
            capabilities.verb(Verb::RawSet).is_supported(),
            ooxml,
            "{format:?}"
        );
        assert!(
            capabilities.verb(Verb::Validate).is_supported(),
            "{format:?}"
        );
        assert_eq!(
            capabilities.verb(Verb::Dump).is_supported(),
            !matches!(format, FileFormat::Odp | FileFormat::Ppt),
            "{format:?}"
        );
        for verb in [
            Verb::Dump,
            Verb::Raw,
            Verb::Validate,
            Verb::AddPart,
            Verb::RawSet,
        ] {
            if let Some(code) = capabilities.verb(verb).code() {
                assert_eq!(code, Code::UnsupportedCapability, "{format:?} {verb:?}");
            }
        }
    }
}

/// A chart's embedded workbook carries the workbook's own content type, as
/// Word writes it; validation accepts it (record 0336).
#[cfg(feature = "docx")]
#[test]
fn docx_with_a_chart_validates() {
    let file = crate::create(FileFormat::Docx).unwrap();
    let file = file
        .add("/body", crate::NewNode::new("paragraph"))
        .unwrap()
        .file;
    let chart = crate::NewNode::new("chart")
        .with_props(crate::props! {
            "chartType" => "column",
            "data" => "East:1,2;West:3,4",
            "categories" => "Q1,Q2",
        })
        .unwrap();
    let file = file.add("/body/p[1]", chart).unwrap().file;
    let validation = file.validate().unwrap();
    assert!(validation.is_valid(), "{:?}", validation.report());
}
