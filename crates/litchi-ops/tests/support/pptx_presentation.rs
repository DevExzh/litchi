//! A synthetic presentation whose every element the tests know.
//!
//! Slide 1 has a centred-title placeholder, a body placeholder, a text box
//! with formatted runs, a line break and a field, a picture with
//! alternative text, a two-by-two table, a group holding a shape and a
//! nested group, and a connector; it has speaker notes and two comments.
//! Slide 2 repeats a shape id, names a shape with a `!!` morph prefix,
//! holds typed SmartArt and OLE frames, and has two placeholders sharing
//! index 1. Slide 3 is empty and hidden. The second slide also repeats a
//! semantic shape name so ordinary name selection has an ambiguity case.
//! A separate minimal fixture retains an unknown graphic frame to verify its
//! explicit capability error.

use std::io::{Cursor, Write as _};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT: &str = "application/vnd.openxmlformats-officedocument.presentationml";
const DIAGRAM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const TABLE: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";
const OLE: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";

/// A one-pixel PNG.
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0xF0,
    0x1F, 0x00, 0x05, 0x00, 0x01, 0xFF, 0x89, 0x99, 0x3D, 0x1D, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn non_visual(tag: &str, id: u32, name: &str, extra: &str) -> String {
    format!(
        r#"<p:nv{tag}Pr><p:cNvPr id="{id}" name="{name}"{extra}/><p:cNv{tag}Pr/><p:nvPr/></p:nv{tag}Pr>"#
    )
}

fn placeholder(id: u32, name: &str, ph: &str, text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{text}</p:txBody></p:sp>"#
    )
}

fn shape(id: u32, name: &str, xfrm: &str, body: &str) -> String {
    format!(
        r#"<p:sp>{}<p:spPr>{xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>{body}</p:sp>"#,
        non_visual("Sp", id, name, "")
    )
}

fn xfrm(x: i64, y: i64, cx: i64, cy: i64, extra: &str) -> String {
    format!(r#"<a:xfrm{extra}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>"#)
}

fn slide(shapes: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"#
    )
}

/// The first slide: one element of every listed kind.
fn slide1() -> String {
    let title = placeholder(
        2,
        "Title 1",
        r#"<p:ph type="ctrTitle"/>"#,
        r#"<a:p><a:r><a:rPr lang="en-US"/><a:t>Quarterly Review</a:t></a:r></a:p>"#,
    );
    let body = placeholder(
        3,
        "Content Placeholder 2",
        r#"<p:ph idx="1"/>"#,
        r#"<a:p><a:pPr lvl="1"/><a:r><a:t>First point</a:t></a:r></a:p><a:p><a:r><a:t>Second point</a:t></a:r></a:p>"#,
    );
    let text_box = shape(
        4,
        "TextBox 3",
        &xfrm(
            914_400,
            1_828_800,
            3_810_000,
            914_400,
            r#" rot="5400000" flipH="1""#,
        ),
        concat!(
            r#"<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>"#,
            r#"<a:p><a:pPr algn="ctr" marL="342900" indent="-342900" rtl="0"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="1200"/></a:spcBef></a:pPr>"#,
            r#"<a:r><a:rPr lang="en-US" sz="2400" b="1" i="0" u="sng" strike="noStrike" spc="150" kern="1200" cap="all" baseline="30000"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:highlight><a:srgbClr val="FFFF00"/></a:highlight><a:latin typeface="Calibri"/><a:ea typeface="MS Gothic"/><a:cs typeface="Arial"/></a:rPr><a:t>Revenue grew</a:t></a:r>"#,
            r#"<a:br><a:rPr lang="en-US"/></a:br>"#,
            r#"<a:r><a:rPr lang="en-US" sz="2400" b="1"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t> 25%</a:t></a:r>"#,
            r#"<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>"#,
            r#"</a:p><a:p><a:pPr algn="r"/><a:r><a:t>Follow-up</a:t></a:r><a:r><a:t> detail</a:t></a:r><a:endParaRPr lang="en-US" sz="1000"/></a:p></p:txBody>"#
        ),
    );
    let picture = format!(
        r#"<p:pic>{}<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
        non_visual("Pic", 5, "Picture 4", r#" descr="A chart of revenue""#),
        xfrm(5_000_000, 1_000_000, 1_828_800, 1_828_800, "")
    );
    let cell = |text: &str| {
        format!(
            r#"<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#
        )
    };
    let table = format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Table 5"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="3657600"/><a:ext cx="3657600" cy="741680"/></p:xfrm><a:graphic><a:graphicData uri="{TABLE}"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid><a:tr h="370840">{}{}</a:tr><a:tr h="370840">{}{}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
        cell("Region"),
        cell("Sales"),
        cell("West"),
        cell("1200")
    );
    let inner = shape(
        9,
        "Inner Shape",
        &xfrm(100, 100, 200, 200, ""),
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Deep text</a:t></a:r></a:p></p:txBody>"#,
    );
    let nested = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Nested Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="200" cy="200"/><a:chOff x="100" y="100"/><a:chExt cx="200" cy="200"/></a:xfrm></p:grpSpPr>{inner}</p:grpSp>"#
    );
    let grouped = shape(
        10,
        "Grouped Shape",
        &xfrm(0, 0, 100, 100, ""),
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Grouped text</a:t></a:r></a:p></p:txBody>"#,
    );
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="Group 6"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="600000"><a:off x="6000000" y="4000000"/><a:ext cx="1000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm></p:grpSpPr>{grouped}{nested}</p:grpSp>"#
    );
    let connector = format!(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="11" name="Straight Connector 10"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>{}<a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>"#,
        xfrm(914_400, 5_000_000, 3_000_000, 0, r#" flipV="1""#)
    );
    slide(&format!(
        "{title}{body}{text_box}{picture}{table}{group}{connector}"
    ))
}

/// The second slide: a repeated id, a morph-prefixed name, typed SmartArt
/// and OLE frames, and placeholders.
fn slide2() -> String {
    let first = shape(
        5,
        "Repeated A",
        &xfrm(0, 0, 3_600_000, 360_000, ""),
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody>"#,
    );
    let second = shape(
        5,
        "Repeated A",
        &xfrm(0, 400_000, 360_000, 360_000, ""),
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Beta</a:t></a:r></a:p></p:txBody>"#,
    );
    let morph = shape(
        6,
        "!!Morph Box",
        &xfrm(0, 800_000, 360_000, 360_000, ""),
        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Gamma</a:t></a:r></a:p></p:txBody>"#,
    );
    let smartart = format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Diagram 6"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="2000000" y="2000000"/><a:ext cx="2640000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="{DIAGRAM}"><dgm:relIds xmlns:dgm="{DIAGRAM}" r:dm="rId7" r:lo="rId8" r:qs="rId9" r:cs="rId10"/></a:graphicData></a:graphic></p:graphicFrame>"#
    );
    let ole = format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Object 7"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="5000000" y="2000000"/><a:ext cx="1000000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="{OLE}"><p:oleObj progId="Package" r:id="rId11"/></a:graphicData></a:graphic></p:graphicFrame>"#
    );
    let body = placeholder(
        9,
        "Text Placeholder 8",
        r#"<p:ph type="body" idx="1"/>"#,
        r#"<a:p><a:r><a:t>Delta</a:t></a:r></a:p>"#,
    );
    let picture = placeholder(
        10,
        "Picture Placeholder 9",
        r#"<p:ph type="pic" idx="1"/>"#,
        "<a:p/>",
    );
    slide(&format!(
        "{first}{second}{morph}{smartart}{ole}{body}{picture}"
    ))
}

/// A slide with an opaque graphic frame whose payload has no typed adapter.
fn unknown_frame_slide() -> String {
    let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Opaque Frame"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="urn:opaque"><opaque:item xmlns:opaque="urn:opaque"/></a:graphicData></a:graphic></p:graphicFrame>"#;
    slide(frame)
}

fn slide3() -> String {
    slide("").replacen("<p:sld ", r#"<p:sld show="0" "#, 1)
}

const NOTES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Remember the forecast</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"#;

const COMMENTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="0" dt="2024-03-01T10:00:00.000" idx="1"><p:pos x="10" y="20"/><p:text>Check the numbers</p:text></p:cm><p:cm authorId="1" dt="2024-03-02T11:30:00.000" idx="1"><p:pos x="30" y="40"/><p:text>Looks good</p:text></p:cm></p:cmLst>"#;

const AUTHORS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="1" clrIdx="0"/><p:cmAuthor id="1" name="Alan Turing" initials="AT" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>"#;

const MODERN_COMMENT_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.comments+xml";
const MODERN_COMMENT_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/comments";
const MODERN_AUTHOR_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.authors+xml";
const MODERN_AUTHOR_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/authors";

const MODERN_AUTHORS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{A1B2C3D4-0000-4000-8000-000000000001}" name="Ada Lovelace" initials="AL" userId="ada@example.com" providerId="None"/><p188:author id="{A1B2C3D4-0000-4000-8000-000000000002}" name="Grace Hopper" initials="GH" userId="grace@example.com" providerId="None"/></p188:authorLst>"#;

const MODERN_COMMENTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:cm id="{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C01}" authorId="{A1B2C3D4-0000-4000-8000-000000000001}" status="resolved" created="2026-10-01T09:30:00.000" complete="50%" title="Forecast"><p188:pos x="10" y="20"/><p188:replyLst><p188:reply id="{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C11}" authorId="{A1B2C3D4-0000-4000-8000-000000000002}" created="2026-10-01T10:00:00.000"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Updated the model</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Please review the forecast</a:t></a:r></a:p></p188:txBody></p188:cm><p188:cm id="{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C02}" authorId="{A1B2C3D4-0000-4000-8000-000000000002}" created="2026-10-01T11:00:00.000"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Looks ready</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#;

const LAYOUT: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1"><p:cSld name="Title Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;

const MASTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1"><p:cSld name="Corporate Master"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst></p:sldMaster>"#;

const NOTES_MASTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/><p:clrMap/></p:notesMaster>"#;

const THEME: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Synthetic Notes"><a:themeElements/></a:theme>"#;

const CORE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Synthetic Deck</dc:title><dc:creator>Litchi Tests</dc:creator></cp:coreProperties>"#;

const THEME_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.theme+xml";

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut out = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{REL}">"#
    );
    for (id, kind, target) in entries {
        out.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{kind}" Target="{target}"/>"#
        ));
    }
    out.push_str("</Relationships>");
    out
}

/// Writes a package of `parts` (name, bytes).
fn package(parts: &[(String, Vec<u8>)]) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (name, bytes) in parts {
        writer
            .start_file(name.as_str(), options)
            .expect("zip entry");
        writer.write_all(bytes).expect("zip write");
    }
    writer.finish().expect("zip finish").into_inner()
}

const OFFICE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A deck of the given slide XML documents; the first slide gets notes,
/// legacy and modern comments, and a picture, every slide the title layout.
fn deck_of(slides: &[String], extras: bool) -> Vec<u8> {
    let mut types = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="{CT}.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="{CT}.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="{CT}.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="{THEME_CONTENT_TYPE}"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>"#
    );
    if extras {
        types.push_str(&format!(
            r#"<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="{CT}.notesMaster+xml"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="{CT}.notesSlide+xml"/><Override PartName="/ppt/comments/comment1.xml" ContentType="{CT}.comments+xml"/><Override PartName="/ppt/commentAuthors.xml" ContentType="{CT}.commentAuthors+xml"/><Override PartName="/ppt/comments/modernComment1.xml" ContentType="{MODERN_COMMENT_CONTENT_TYPE}"/><Override PartName="/ppt/modernAuthors.xml" ContentType="{MODERN_AUTHOR_CONTENT_TYPE}"/>"#
        ));
    }
    let mut ids = String::new();
    let mut presentation_rels = Vec::new();
    for index in 1..=slides.len() {
        types.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{index}.xml" ContentType="{CT}.slide+xml"/>"#
        ));
        ids.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{index}"/>"#,
            255 + index
        ));
        presentation_rels.push((
            format!("rId{index}"),
            format!("{OFFICE}/slide"),
            format!("slides/slide{index}.xml"),
        ));
    }
    types.push_str("</Types>");
    presentation_rels.push((
        "rIdMaster".to_owned(),
        format!("{OFFICE}/slideMaster"),
        "slideMasters/slideMaster1.xml".to_owned(),
    ));
    if extras {
        presentation_rels.push((
            "rIdNotesMaster".to_owned(),
            format!("{OFFICE}/notesMaster"),
            "notesMasters/notesMaster1.xml".to_owned(),
        ));
        presentation_rels.push((
            "rIdAuthors".to_owned(),
            format!("{OFFICE}/commentAuthors"),
            "commentAuthors.xml".to_owned(),
        ));
        presentation_rels.push((
            "rIdModernAuthors".to_owned(),
            MODERN_AUTHOR_RELATIONSHIP.to_owned(),
            "modernAuthors.xml".to_owned(),
        ));
    }
    let notes_master = if extras {
        r#"<p:notesMasterIdLst><p:notesMasterId r:id="rIdNotesMaster"/></p:notesMasterIdLst>"#
    } else {
        ""
    };
    let presentation = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>{notes_master}<p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    let borrowed: Vec<(&str, &str, &str)> = presentation_rels
        .iter()
        .map(|(id, kind, target)| (id.as_str(), kind.as_str(), target.as_str()))
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), types.into_bytes()),
        (
            "_rels/.rels".to_owned(),
            relationships(&[
                ("rId1", &format!("{OFFICE}/officeDocument"), "ppt/presentation.xml"),
                (
                    "rId2",
                    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                    "docProps/core.xml",
                ),
            ])
            .into_bytes(),
        ),
        ("docProps/core.xml".to_owned(), CORE.as_bytes().to_vec()),
        ("ppt/presentation.xml".to_owned(), presentation.into_bytes()),
        (
            "ppt/_rels/presentation.xml.rels".to_owned(),
            relationships(&borrowed).into_bytes(),
        ),
        (
            "ppt/slideMasters/slideMaster1.xml".to_owned(),
            MASTER.as_bytes().to_vec(),
        ),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels".to_owned(),
            relationships(&[
                (
                    "rIdLayout",
                    &format!("{OFFICE}/slideLayout"),
                    "../slideLayouts/slideLayout1.xml",
                ),
                (
                    "rIdTheme",
                    &format!("{OFFICE}/theme"),
                    "../theme/theme1.xml",
                ),
            ])
            .into_bytes(),
        ),
        (
            "ppt/slideLayouts/slideLayout1.xml".to_owned(),
            LAYOUT.as_bytes().to_vec(),
        ),
        (
            "ppt/theme/theme1.xml".to_owned(),
            THEME.as_bytes().to_vec(),
        ),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels".to_owned(),
            relationships(&[(
                "rIdMaster",
                &format!("{OFFICE}/slideMaster"),
                "../slideMasters/slideMaster1.xml",
            )])
            .into_bytes(),
        ),
    ];
    for (index, xml) in slides.iter().enumerate() {
        let number = index + 1;
        parts.push((
            format!("ppt/slides/slide{number}.xml"),
            xml.clone().into_bytes(),
        ));
        let layout = format!("{OFFICE}/slideLayout");
        let mut rels = vec![("rId1", layout.as_str(), "../slideLayouts/slideLayout1.xml")];
        let image = format!("{OFFICE}/image");
        let notes = format!("{OFFICE}/notesSlide");
        let comments = format!("{OFFICE}/comments");
        let modern_comments = MODERN_COMMENT_RELATIONSHIP;
        if extras && number == 1 {
            rels.push(("rId2", image.as_str(), "../media/image1.png"));
            rels.push(("rId3", notes.as_str(), "../notesSlides/notesSlide1.xml"));
            rels.push(("rId4", comments.as_str(), "../comments/comment1.xml"));
            rels.push(("rId5", modern_comments, "../comments/modernComment1.xml"));
        }
        parts.push((
            format!("ppt/slides/_rels/slide{number}.xml.rels"),
            relationships(&rels).into_bytes(),
        ));
    }
    if extras {
        parts.push((
            "ppt/notesMasters/notesMaster1.xml".to_owned(),
            NOTES_MASTER.as_bytes().to_vec(),
        ));
        parts.push((
            "ppt/notesMasters/_rels/notesMaster1.xml.rels".to_owned(),
            relationships(&[(
                "rIdTheme",
                &format!("{OFFICE}/theme"),
                "../theme/theme1.xml",
            )])
            .into_bytes(),
        ));
        parts.push(("ppt/media/image1.png".to_owned(), PNG.to_vec()));
        parts.push((
            "ppt/notesSlides/notesSlide1.xml".to_owned(),
            NOTES.as_bytes().to_vec(),
        ));
        parts.push((
            "ppt/notesSlides/_rels/notesSlide1.xml.rels".to_owned(),
            relationships(&[
                ("rId1", &format!("{OFFICE}/slide"), "../slides/slide1.xml"),
                (
                    "rId2",
                    &format!("{OFFICE}/notesMaster"),
                    "../notesMasters/notesMaster1.xml",
                ),
            ])
            .into_bytes(),
        ));
        parts.push((
            "ppt/comments/comment1.xml".to_owned(),
            COMMENTS.as_bytes().to_vec(),
        ));
        parts.push((
            "ppt/commentAuthors.xml".to_owned(),
            AUTHORS.as_bytes().to_vec(),
        ));
        parts.push((
            "ppt/comments/modernComment1.xml".to_owned(),
            MODERN_COMMENTS.as_bytes().to_vec(),
        ));
        parts.push((
            "ppt/modernAuthors.xml".to_owned(),
            MODERN_AUTHORS.as_bytes().to_vec(),
        ));
    }
    package(&parts)
}

/// The synthetic deck of the module documentation.
pub(super) fn bytes() -> Vec<u8> {
    deck_of(&[slide1(), slide2(), slide3()], true)
}

/// A minimal deck whose selected slide contains one unknown graphic frame.
/// The unified facade must reject the slide with `unsupported_capability`
/// before exposing any frame as an ordinary node.
pub(super) fn bytes_with_generic_frame() -> Vec<u8> {
    deck_of(&[unknown_frame_slide()], false)
}

/// The same shape-rich first slide followed by a deliberately malformed
/// second slide. The package catalog remains valid, so a lazy query that
/// stops on the first match must never decode the malformed tail.
pub(super) fn bytes_with_malformed_tail() -> Vec<u8> {
    deck_of(
        &[
            slide1(),
            format!(r#"<p:sld xmlns:p="{P}"><p:cSld><p:spTree><p:sp>"#),
        ],
        true,
    )
}

/// A deck of `slides` slides, each with a title and `boxes` text boxes, the
/// first of which is 11 cm wide and says "Revenue".
#[allow(dead_code, reason = "fixture helper reserved for allocation tests")]
pub(super) fn generated(slides: usize, boxes: usize) -> Vec<u8> {
    let slides: Vec<String> = (0..slides)
        .map(|index| {
            let mut shapes = placeholder(
                2,
                "Title 1",
                r#"<p:ph type="title"/>"#,
                &format!(r#"<a:p><a:r><a:t>Slide {}</a:t></a:r></a:p>"#, index + 1),
            );
            for number in 0..boxes {
                let id = u32::try_from(number).unwrap_or(0) + 3;
                let width = if number == 0 { 3_960_000 } else { 1_800_000 };
                let text = if number == 0 { "Revenue grew" } else { "Notes" };
                shapes.push_str(&shape(
                    id,
                    &format!("TextBox {}", id - 1),
                    &xfrm(360_000, 360_000 * i64::from(id), width, 360_000, ""),
                    &format!(
                        r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>{text} {number}</a:t></a:r></a:p></p:txBody>"#
                    ),
                ));
            }
            slide(&shapes)
        })
        .collect();
    deck_of(&slides, false)
}
