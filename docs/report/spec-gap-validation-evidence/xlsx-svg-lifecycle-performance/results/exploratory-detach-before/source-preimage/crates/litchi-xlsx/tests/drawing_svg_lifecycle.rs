//! Ordinary worksheet SVG lifecycle integration coverage.
//!
//! The fixtures are assembled at the OPC boundary so the assertions exercise
//! `Workbook::edit` and its worksheet transaction rather than a private XML
//! splice.  The source drawing deliberately contains all three anchor forms,
//! inherited namespace bindings, opaque extension payload, and relationship
//! edges which make SVG cleanup observable.

#![allow(
    clippy::unwrap_used,
    reason = "focused integration fixtures use panic-on-failure assertions"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::drawing::{
    DrawingAnchor, DrawingDialect, EditAs, PictureSelector, RelationshipDialect, ScanLimits,
    SourceDrawing, SvgInput, SvgOwnerState, parse,
};
use litchi_xlsx::{Error, Package, ReadLimits, Workbook};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WORKBOOK: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const SHEET2: &str = "/xl/worksheets/sheet2.xml";
const DRAWING: &str = "/xl/drawings/drawing1.xml";
const DRAWING2: &str = "/xl/drawings/drawing2.xml";
const ORPHAN_DRAWING: &str = "/xl/drawings/orphan.xml";
const RASTER: &str = "/xl/media/image1.png";
const RASTER2: &str = "/xl/media/image3.png";
const SVG: &str = "/xl/media/image2.svg";
const OPAQUE: &str = "/xl/media/untouched.bin";
const OTHER: &str = "/xl/opaque-owner.xml";

const NEW_SVG: &[u8] = br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0 L4 4"/></svg>"#;
const PNG: &[u8] = b"synthetic-png-payload";
const SVG_PAYLOAD: &[u8] =
    br#"<svg xmlns="http://www.w3.org/2000/svg"><circle cx="3" cy="4" r="2"/></svg>"#;

#[derive(Clone, Copy, Debug)]
struct DrawingOptions {
    strict: bool,
    anchors: bool,
    admitted_svg: bool,
    unknown_extension: bool,
    duplicate_svg: bool,
    mce_svg: bool,
    linked_svg: bool,
    foreign_descendant: bool,
    incoming_svg_edge: bool,
    external_raster: bool,
    wrong_raster_type: bool,
}

impl Default for DrawingOptions {
    fn default() -> Self {
        Self {
            strict: false,
            anchors: true,
            admitted_svg: false,
            unknown_extension: true,
            duplicate_svg: false,
            mce_svg: false,
            linked_svg: false,
            foreign_descendant: false,
            incoming_svg_edge: false,
            external_raster: false,
            wrong_raster_type: false,
        }
    }
}

fn q(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

fn marker(prefix: &str, name: &str, col: i64, col_off: i64, row: i64, row_off: i64) -> String {
    let marker = q(prefix, name);
    let col_name = q(prefix, "col");
    let col_off_name = q(prefix, "colOff");
    let row_name = q(prefix, "row");
    let row_off_name = q(prefix, "rowOff");
    format!(
        "<{marker}><{col_name}>{col}</{col_name}><{col_off_name}>{col_off}</{col_off_name}><{row_name}>{row}</{row_name}><{row_off_name}>{row_off}</{row_off_name}></{marker}>"
    )
}

fn picture(xdr: &str, a: &str, rel: &str, id: usize, body: &str) -> String {
    let pic = q(xdr, "pic");
    let nv = q(xdr, "nvPicPr");
    let c_nv_pr = q(xdr, "cNvPr");
    let c_nv_pic_pr = q(xdr, "cNvPicPr");
    let blip_fill = q(xdr, "blipFill");
    let blip = q(a, "blip");
    let sp_pr = q(xdr, "spPr");
    format!(
        r#"<{pic}><{nv}><{c_nv_pr} id="{id}" name="picture-{id}"/><{c_nv_pic_pr}/></{nv}><{blip_fill}><{blip} {rel}:embed="rIdRaster">{body}</{blip}></{blip_fill}><{sp_pr}/></{pic}>"#
    )
}

fn owner_body(options: DrawingOptions, a: &str, rel: &str, index: usize) -> String {
    let opaque = if options.unknown_extension {
        r#"<a:ext uri="urn:future-extension"><!--future-comment--><future:payload xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future-payload>]]><?future-processing?></future:payload></a:ext>"#
            .replace("a:", &format!("{a}:"))
    } else {
        String::new()
    };
    let foreign = if options.foreign_descendant {
        format!(
            r#"<{a}:ext uri="urn:foreign"><foreign:wrapper xmlns:foreign="urn:foreign"><foreign:svgBlip {rel}:embed="rIdForeign"/></foreign:wrapper></{a}:ext>"#,
            a = a,
            rel = rel,
        )
    } else {
        String::new()
    };
    let admitted = if options.admitted_svg {
        let attr_prefix = if options.strict { "trans" } else { rel };
        let attr = if options.linked_svg { "link" } else { "embed" };
        format!(
            r#"<{a}:ext uri="{SVG_URI}"><asvg:svgBlip {attr_prefix}:{attr}="rIdSvg"/></{a}:ext>"#,
            a = a,
            attr_prefix = attr_prefix,
            attr = attr,
        )
    } else {
        String::new()
    };
    let duplicate = if options.duplicate_svg {
        format!(
            r#"<{a}:ext uri="{SVG_URI}"><asvg:svgBlip {rel}:embed="rIdSvg"/></{a}:ext>"#,
            a = a,
            rel = rel,
        )
    } else {
        String::new()
    };
    let extensions = if options.mce_svg {
        format!(
            r#"<mc:AlternateContent><mc:Choice Requires="asvg"><{a}:extLst><{a}:ext uri="{SVG_URI}"><asvg:svgBlip {rel}:embed="rIdSvg"/></{a}:ext></{a}:extLst></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
            a = a,
            rel = rel,
        )
    } else if admitted.is_empty() && duplicate.is_empty() && opaque.is_empty() && foreign.is_empty()
    {
        String::new()
    } else {
        format!(
            r#"<{a}:extLst>{admitted}{duplicate}{opaque}{foreign}</{a}:extLst>"#,
            a = a,
        )
    };
    // Keep the source owner lexical spelling distinct on one picture so the
    // scanner has to use ancestry and namespace bindings rather than text.
    let extensions = if index == 1 && options.admitted_svg && !options.strict {
        extensions.replace("<asvg:svgBlip", "<svg:svgBlip")
    } else {
        extensions
    };
    extensions
}

fn anchor(xdr: &str, a: &str, rel: &str, index: usize, body: &str) -> String {
    let pic = picture(xdr, a, rel, index + 1, body);
    let client_data = q(xdr, "clientData");
    match index {
        0 => format!(
            r#"<{xdr}:twoCellAnchor>{from}{to}{pic}<{client_data}/></{xdr}:twoCellAnchor>"#,
            xdr = xdr,
            from = marker(xdr, "from", 1, 2, 3, 4),
            to = marker(xdr, "to", 5, 6, 7, 8),
        ),
        1 => {
            let ext = q(xdr, "ext");
            format!(
                r#"<{xdr}:oneCellAnchor>{from}<{ext} cx="123456" cy="654321"/>{pic}<{client_data}/></{xdr}:oneCellAnchor>"#,
                xdr = xdr,
                from = marker(xdr, "from", 9, 19, 10, 20),
            )
        },
        _ => {
            let pos = q(xdr, "pos");
            let ext = q(xdr, "ext");
            format!(
                r#"<{xdr}:absoluteAnchor><{pos} x="-900" y="456"/><{ext} cx="777888" cy="999000"/>{pic}<{client_data}/></{xdr}:absoluteAnchor>"#,
                xdr = xdr,
            )
        },
    }
}

fn drawing_xml(options: DrawingOptions) -> Vec<u8> {
    let (xdr, a, rel, root, close) = if options.strict {
        (
            "x",
            "a",
            "r",
            format!(
                r#"<x:wsDr xmlns:x="{STRICT_XDR}" xmlns:a="{STRICT_A}" xmlns:r="{STRICT_REL}" xmlns:trans="{REL}" xmlns:asvg="{SVG_NS}" xmlns:mc="{MCE}" xmlns:future="urn:litchi:future" xmlns:svg="{SVG_NS}">"#
            ),
            "</x:wsDr>",
        )
    } else {
        (
            "x",
            "a",
            "r",
            format!(
                r#"<x:wsDr xmlns:x="{XDR}" xmlns:a="{A}" xmlns:r="{REL}" xmlns:asvg="{SVG_NS}" xmlns:mc="{MCE}" xmlns:future="urn:litchi:future" xmlns:svg="{SVG_NS}">"#
            ),
            "</x:wsDr>",
        )
    };
    let count = if options.anchors { 3 } else { 1 };
    let body = (0..count)
        .map(|index| anchor(xdr, a, rel, index, &owner_body(options, a, rel, index)))
        .collect::<String>();
    format!("{root}{body}{close}").into_bytes()
}

fn worksheet_xml(strict: bool) -> Vec<u8> {
    let (main, rel) = if strict {
        ("http://purl.oclc.org/ooxml/spreadsheetml/main", STRICT_REL)
    } else {
        (SML, REL)
    };
    format!(
        r#"<worksheet xmlns="{main}" xmlns:r="{rel}"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#
    )
    .into_bytes()
}

fn package_bytes(options: DrawingOptions) -> Vec<u8> {
    let mut package = Package::create().unwrap().into_plain_opc();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(worksheet_xml(options.strict));
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(DRAWING).unwrap(),
            ct::OFC_DRAWING.to_owned(),
            drawing_xml(options),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(RASTER).unwrap(),
            if options.wrong_raster_type {
                "image/jpeg".to_owned()
            } else {
                ct::PNG.to_owned()
            },
            PNG.to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(OPAQUE).unwrap(),
            "application/octet-stream".to_owned(),
            b"untouched opaque member".to_vec(),
        )))
        .unwrap();
    if (options.admitted_svg || options.duplicate_svg || options.mce_svg) && !options.linked_svg {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SVG).unwrap(),
                "image/svg+xml".to_owned(),
                SVG_PAYLOAD.to_vec(),
            )))
            .unwrap();
    }
    if options.incoming_svg_edge {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(OTHER).unwrap(),
                "application/xml".to_owned(),
                b"<opaque-owner/>".to_vec(),
            )))
            .unwrap();
    }
    let sheet = package.get_part_mut(&PackURI::new(SHEET).unwrap()).unwrap();
    sheet
        .rels_mut()
        .try_add_relationship(
            if options.strict {
                rt::STRICT_DRAWING
            } else {
                rt::DRAWING
            }
            .to_owned(),
            "../drawings/drawing1.xml".to_owned(),
            "rIdDrawing".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let drawing = package
        .get_part_mut(&PackURI::new(DRAWING).unwrap())
        .unwrap();
    drawing
        .rels_mut()
        .try_add_relationship(
            if options.strict {
                rt::STRICT_IMAGE
            } else {
                rt::IMAGE
            }
            .to_owned(),
            if options.external_raster {
                "https://example.invalid/fallback.png".to_owned()
            } else {
                "../media/image1.png".to_owned()
            },
            "rIdRaster".to_owned(),
            if options.external_raster {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )
        .unwrap();
    if options.admitted_svg || options.linked_svg || options.duplicate_svg || options.mce_svg {
        drawing
            .rels_mut()
            .try_add_relationship(
                if options.strict {
                    rt::STRICT_IMAGE
                } else {
                    rt::IMAGE
                }
                .to_owned(),
                if options.linked_svg {
                    "https://example.invalid/vector.svg".to_owned()
                } else {
                    "../media/image2.svg".to_owned()
                },
                "rIdSvg".to_owned(),
                if options.linked_svg {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )
            .unwrap();
    }
    if options.incoming_svg_edge {
        package
            .get_part_mut(&PackURI::new(OTHER).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "media/image2.svg".to_owned(),
                "rIdIncomingSvg".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn raster_fixture() -> Vec<u8> {
    package_bytes(DrawingOptions::default())
}

fn svg_fixture() -> Vec<u8> {
    package_bytes(DrawingOptions {
        admitted_svg: true,
        anchors: false,
        unknown_extension: true,
        ..DrawingOptions::default()
    })
}

fn shared_svg_fixture(incoming_svg_edge: bool) -> Vec<u8> {
    package_bytes(DrawingOptions {
        admitted_svg: true,
        anchors: true,
        incoming_svg_edge,
        unknown_extension: true,
        ..DrawingOptions::default()
    })
}

fn empty_drawing_xml() -> Vec<u8> {
    format!(r#"<x:wsDr xmlns:x="{XDR}" xmlns:a="{A}" xmlns:r="{REL}"/>"#).into_bytes()
}

fn two_worksheet_fixture(shared_svg: bool) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&raster_fixture()).unwrap();
    let options = DrawingOptions {
        anchors: false,
        admitted_svg: shared_svg,
        unknown_extension: true,
        ..DrawingOptions::default()
    };
    package
        .get_part_mut(&PackURI::new(DRAWING).unwrap())
        .unwrap()
        .set_blob(drawing_xml(options));
    if shared_svg {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SVG).unwrap(),
                "image/svg+xml".to_owned(),
                SVG_PAYLOAD.to_vec(),
            )))
            .unwrap();
        package
            .get_part_mut(&PackURI::new(DRAWING).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/image2.svg".to_owned(),
                "rIdSvg".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }

    let sheet2_xml = String::from_utf8(worksheet_xml(false))
        .unwrap()
        .replace("rIdDrawing", "rIdDrawing2")
        .into_bytes();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SHEET2).unwrap(),
            ct::SML_WORKSHEET.to_owned(),
            sheet2_xml.clone(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(DRAWING2).unwrap(),
            ct::OFC_DRAWING.to_owned(),
            drawing_xml(options),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(RASTER2).unwrap(),
            ct::PNG.to_owned(),
            PNG.to_vec(),
        )))
        .unwrap();
    let sheet2 = package
        .get_part_mut(&PackURI::new(SHEET2).unwrap())
        .unwrap();
    sheet2.set_blob(sheet2_xml);
    sheet2
        .rels_mut()
        .try_add_relationship(
            rt::DRAWING.to_owned(),
            "../drawings/drawing2.xml".to_owned(),
            "rIdDrawing2".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let drawing2 = package
        .get_part_mut(&PackURI::new(DRAWING2).unwrap())
        .unwrap();
    drawing2
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/image3.png".to_owned(),
            "rIdRaster".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    if shared_svg {
        drawing2
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/image2.svg".to_owned(),
                "rIdSvg".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    let workbook = package
        .get_part_mut(&PackURI::new(WORKBOOK).unwrap())
        .unwrap();
    let workbook_xml = String::from_utf8(workbook.blob().to_vec()).unwrap();
    workbook.set_blob(
        workbook_xml
            .replace(
                "</sheets>",
                r#"<sheet name="Sheet2" sheetId="2" r:id="rId3"/></sheets>"#,
            )
            .into_bytes(),
    );
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet2.xml".to_owned(),
            "rId3".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn orphan_relationship_fixture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&raster_fixture()).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(ORPHAN_DRAWING).unwrap(),
            ct::OFC_DRAWING.to_owned(),
            empty_drawing_xml(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::DRAWING.to_owned(),
            "../drawings/orphan.xml".to_owned(),
            "rIdAOrphan".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn has_part(bytes: &[u8], name: &str) -> bool {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(name).unwrap())
        .is_ok()
}

fn content_type(bytes: &[u8], name: &str) -> String {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .content_type()
        .to_owned()
}

fn part_name_with_content_type(bytes: &[u8], wanted: &str) -> String {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .iter_parts()
        .find(|part| part.content_type() == wanted)
        .unwrap()
        .partname()
        .as_str()
        .to_owned()
}

fn part_names_with_content_type(bytes: &[u8], wanted: &str) -> Vec<String> {
    let mut names = OpcPackage::from_bytes(bytes)
        .unwrap()
        .iter_parts()
        .filter(|part| part.content_type() == wanted)
        .map(|part| part.partname().as_str().to_owned())
        .collect::<Vec<_>>();
    names.sort();
    names
}

fn has_content_type(bytes: &[u8], wanted: &str) -> bool {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .iter_parts()
        .any(|part| part.content_type() == wanted)
}

fn rels(bytes: &[u8], name: &str) -> Vec<u8> {
    let name = name.trim_start_matches('/');
    let (directory, file) = name.rsplit_once('/').unwrap();
    let member = format!("{directory}/_rels/{file}.rels");
    PhysPkgReader::new(bytes)
        .unwrap()
        .read_member(&member)
        .unwrap()
}

fn content_types_member(bytes: &[u8]) -> Vec<u8> {
    PhysPkgReader::new(bytes)
        .unwrap()
        .read_member("[Content_Types].xml")
        .unwrap()
}

fn rewrite_content_types_member(bytes: &[u8], rewrite: impl FnOnce(Vec<u8>) -> Vec<u8>) -> Vec<u8> {
    let reader = PhysPkgReader::new(bytes).unwrap();
    let members = reader.member_names().unwrap();
    let mut writer = PhysPkgWriter::new();
    let mut rewrite = Some(rewrite);
    for member in members {
        let payload = reader.read_member(&member).unwrap();
        let payload = if member == "[Content_Types].xml" {
            rewrite.take().unwrap()(payload)
        } else {
            payload
        };
        let uri = PackURI::new(format!("/{member}")).unwrap();
        writer.write(&uri, &payload).unwrap();
    }
    writer.finish().unwrap()
}

fn unusual_manifest_fixture() -> Vec<u8> {
    rewrite_content_types_member(&raster_fixture(), |manifest| {
        let closing = manifest
            .windows(b"</Types>".len())
            .position(|window| window == b"</Types>")
            .unwrap();
        let mut output = Vec::with_capacity(manifest.len() + 37);
        output.extend_from_slice(&manifest[..closing]);
        output.extend_from_slice(b"<!-- source-manifest-comment -->");
        output.extend_from_slice(&manifest[closing..]);
        output
    })
}

fn rewrite_drawing(bytes: &[u8], rewrite: impl FnOnce(String) -> String) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    let uri = PackURI::new(DRAWING).unwrap();
    let drawing = package.get_part_mut(&uri).unwrap();
    let source = String::from_utf8(drawing.blob().to_vec()).unwrap();
    drawing.set_blob(rewrite(source).into_bytes());
    PackageWriter::to_bytes(&package).unwrap()
}

fn rewrite_sheet(bytes: &[u8], rewrite: impl FnOnce(String) -> String) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    let uri = PackURI::new(SHEET).unwrap();
    let sheet = package.get_part_mut(&uri).unwrap();
    let source = String::from_utf8(sheet.blob().to_vec()).unwrap();
    sheet.set_blob(rewrite(source).into_bytes());
    PackageWriter::to_bytes(&package).unwrap()
}

fn missing_direct_drawing_fixture() -> Vec<u8> {
    rewrite_sheet(&raster_fixture(), |xml| {
        let drawing = r#"<drawing r:id="rIdDrawing"/>"#;
        assert!(xml.contains(drawing));
        xml.replace(drawing, "")
    })
}

fn hidden_mce_drawing_fixture() -> Vec<u8> {
    rewrite_sheet(&raster_fixture(), |xml| {
        let drawing = r#"<drawing r:id="rIdDrawing"/>"#;
        let hidden = format!(
            r#"<mc:AlternateContent><mc:Choice Requires="x"><drawing r:id="rIdDrawing"/></mc:Choice><mc:Fallback><drawing r:id="rIdDrawing"/></mc:Fallback></mc:AlternateContent>"#
        );
        assert!(xml.contains(drawing));
        let old_ns = format!("xmlns:r=\"{REL}\"");
        let new_ns = format!("xmlns:r=\"{REL}\" xmlns:mc=\"{MCE}\" xmlns:x=\"urn:litchi:choice\"");
        xml.replace(drawing, &hidden).replace(&old_ns, &new_ns)
    })
}

fn worksheet_namespace_alias_fixture() -> Vec<u8> {
    rewrite_sheet(&raster_fixture(), |_| {
        r#"<ws:worksheet xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/mai&#110;" xmlns:rr="http://schemas.openxmlformats.org/officeDocument/2006/relationship&#115;"><ws:dimension ref="A1:C3"/><ws:sheetData/><ws:drawing rr:id="rIdDrawing"/></ws:worksheet>"#
            .to_owned()
    })
}

fn duplicate_direct_drawing_fixture() -> Vec<u8> {
    rewrite_sheet(&raster_fixture(), |xml| {
        let drawing = r#"<drawing r:id="rIdDrawing"/>"#;
        let duplicate = r#"<drawing r:id="rIdDrawing"/><drawing r:id="rIdDrawing"/>"#;
        assert_eq!(xml.matches(drawing).count(), 1);
        xml.replace(drawing, duplicate)
    })
}

fn mce_and_direct_drawing_fixture() -> Vec<u8> {
    rewrite_sheet(&raster_fixture(), |xml| {
        let drawing = r#"<drawing r:id="rIdDrawing"/>"#;
        let hidden_and_direct = r#"<mc:AlternateContent><mc:Choice Requires="x"><drawing r:id="rIdDrawing"/></mc:Choice><mc:Fallback/></mc:AlternateContent><drawing r:id="rIdDrawing"/>"#;
        assert_eq!(xml.matches(drawing).count(), 1);
        let old_root = format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}">"#);
        let new_root = format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}" xmlns:mc="{MCE}" xmlns:x="urn:litchi:mce">"#
        );
        xml.replace(&old_root, &new_root)
            .replace(drawing, hidden_and_direct)
    })
}

fn opaque_worksheet_extension_fixture() -> Vec<u8> {
    rewrite_sheet(&raster_fixture(), |_| {
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}" xmlns:mc="{MCE}" xmlns:s="{SML}" xmlns:ext="urn:litchi:worksheet-extension"><dimension ref="A1:C3"/><sheetData/><drawing r:id="rIdDrawing"/><extLst><ext uri="urn:litchi:opaque"><ext:payload><s:drawing r:id="rIdDrawing"/><mc:AlternateContent><mc:Choice Requires="x"><s:drawing r:id="rIdDrawing"/></mc:Choice><mc:Fallback/></mc:AlternateContent></ext:payload></ext></extLst></worksheet>"#
        )
    })
}

fn namespace_only_fixture() -> Vec<u8> {
    rewrite_drawing(&raster_fixture(), |xml| {
        let old = r#"<a:ext uri="urn:future-extension"><!--future-comment--><future:payload xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future-payload>]]><?future-processing?></future:payload></a:ext>"#;
        let replacement = r#"<a:ext uri="urn:future-extension"><future:empty xmlns:future="urn:litchi:future"/></a:ext>"#;
        assert!(xml.contains(old));
        xml.replace(old, replacement)
    })
}

fn edit_attach(bytes: &[u8], selector: PictureSelector, svg: &[u8]) -> Result<Vec<u8>, Error> {
    let workbook = Workbook::from_bytes(bytes.to_vec())?;
    let mut edit = workbook.edit()?;
    edit.sheet("Sheet1")?
        .ok_or_else(|| Error::Invalid("Sheet1 is missing".into()))?
        .attach_svg(selector, SvgInput::borrowed(svg))?;
    edit.commit()?.into_workbook().to_plain_bytes()
}

fn edit_attach_with_limits(
    bytes: &[u8],
    selector: PictureSelector,
    svg: &[u8],
    limits: ReadLimits,
) -> Result<Vec<u8>, Error> {
    let workbook = Workbook::from_bytes_with_limits(bytes.to_vec(), limits)?;
    let mut edit = workbook.edit()?;
    edit.sheet("Sheet1")?
        .ok_or_else(|| Error::Invalid("Sheet1 is missing".into()))?
        .attach_svg(selector, SvgInput::borrowed(svg))?;
    edit.commit()?.into_workbook().to_plain_bytes()
}

fn edit_detach(bytes: &[u8], selector: PictureSelector) -> Result<Vec<u8>, Error> {
    let workbook = Workbook::from_bytes(bytes.to_vec())?;
    let mut edit = workbook.edit()?;
    edit.sheet("Sheet1")?
        .ok_or_else(|| Error::Invalid("Sheet1 is missing".into()))?
        .detach_svg(selector)?;
    edit.commit()?.into_workbook().to_plain_bytes()
}

fn selector(picture: usize) -> PictureSelector {
    selector_at(0, picture)
}

fn selector_at(drawing: usize, picture: usize) -> PictureSelector {
    PictureSelector::new(drawing, picture)
}

fn geometry_kind(anchor: &DrawingAnchor) -> &'static str {
    match anchor {
        DrawingAnchor::TwoCell { edit_as, .. } => match edit_as {
            EditAs::TwoCell => "twoCell",
            EditAs::OneCell => "twoCell-oneCell",
            EditAs::Absolute => "twoCell-absolute",
        },
        DrawingAnchor::OneCell { .. } => "oneCell",
        DrawingAnchor::Absolute { .. } => "absolute",
    }
}

#[test]
fn scanner_matches_typed_inventory_for_all_anchor_forms_and_direct_svg_owner() {
    // The typed DrawingML reader intentionally rejects processing instructions
    // and opaque payloads. Keep this comparison fixture structural so it
    // exercises source-order ancestry against the semantic inventory; the raw
    // preservation assertions below retain the full opaque fixture.
    let bytes = package_bytes(DrawingOptions {
        unknown_extension: false,
        ..DrawingOptions::default()
    });
    let drawing_xml = part(&bytes, DRAWING);
    let source = SourceDrawing::scan_with_ordinal(&drawing_xml, 4).unwrap();
    assert_eq!(source.drawing_ordinal(), 4);
    assert_eq!(source.dialect(), DrawingDialect::Transitional);
    assert_eq!(
        source.relationship_dialect(),
        RelationshipDialect::Transitional
    );
    assert_eq!(source.pictures().len(), 3);
    let text = String::from_utf8(drawing_xml.clone()).unwrap();
    let typed = parse(&text).unwrap().unwrap();
    let typed_anchors = typed
        .pictures()
        .map(|picture| *picture.drawing_anchor())
        .collect::<Vec<_>>();
    assert_eq!(typed_anchors.len(), 3);
    match source.picture(0).unwrap().anchor() {
        DrawingAnchor::TwoCell {
            from,
            to,
            edit_as: EditAs::TwoCell,
        } => {
            assert_eq!(
                (
                    from.column,
                    from.column_offset.0,
                    from.row,
                    from.row_offset.0
                ),
                (1, 2, 3, 4)
            );
            assert_eq!(
                (to.column, to.column_offset.0, to.row, to.row_offset.0),
                (5, 6, 7, 8)
            );
        },
        other => panic!("unexpected first anchor geometry: {other:?}"),
    }
    match source.picture(1).unwrap().anchor() {
        DrawingAnchor::OneCell { from, extent } => {
            assert_eq!(
                (
                    from.column,
                    from.column_offset.0,
                    from.row,
                    from.row_offset.0
                ),
                (9, 19, 10, 20)
            );
            assert_eq!((extent.width.0, extent.height.0), (123_456, 654_321));
        },
        other => panic!("unexpected second anchor geometry: {other:?}"),
    }
    match source.picture(2).unwrap().anchor() {
        DrawingAnchor::Absolute { position, extent } => {
            assert_eq!((position.x.0, position.y.0), (-900, 456));
            assert_eq!((extent.width.0, extent.height.0), (777_888, 999_000));
        },
        other => panic!("unexpected third anchor geometry: {other:?}"),
    }
    for (index, picture) in source.pictures().iter().enumerate() {
        assert_eq!(picture.picture_ordinal(), index);
        assert_eq!(picture.drawing_ordinal(), 4);
        assert_eq!(picture.anchor(), &typed_anchors[index]);
        assert!(picture.anchor_range().start < picture.anchor_range().end);
        assert!(picture.picture_range().start < picture.picture_range().end);
        assert_eq!(
            picture.picture_bytes(&drawing_xml).unwrap(),
            &drawing_xml[picture.picture_range().start..picture.picture_range().end]
        );
        assert_eq!(
            geometry_kind(picture.anchor()),
            geometry_kind(&typed_anchors[index])
        );
        assert_eq!(picture.raster_relationship_id(), "rIdRaster");
        assert!(
            picture
                .relationship_references()
                .iter()
                .any(|reference| reference.id() == "rIdRaster")
        );
    }
    assert!(matches!(
        source.pictures()[0].svg_owner(),
        SvgOwnerState::None | SvgOwnerState::Opaque
    ));
    assert!(
        source
            .pictures()
            .iter()
            .all(|picture| !picture.is_direct_embedded_svg())
    );
    assert_eq!(source.relationship_references().len(), 3);
}

#[test]
fn scanner_preserves_direct_owner_ranges_inherited_prefixes_and_opaque_payload() {
    let bytes = svg_fixture();
    let drawing_xml = part(&bytes, DRAWING);
    let source = SourceDrawing::scan(&drawing_xml).unwrap();
    let picture = source.picture(0).unwrap();
    let owner = picture.svg_owner().owner().unwrap();
    assert_eq!(owner.embedded_relationship_id(), Some("rIdSvg"));
    assert_eq!(owner.uri_lexical(), SVG_URI.as_bytes());
    assert!(owner.extension_range().start > picture.blip_range().start_end());
    assert!(owner.extension_range().end <= picture.blip_range().range().end);
    assert!(owner.svg_blip_range().start < owner.svg_blip_range().end);
    assert!(
        owner
            .parsed_source()
            .unwrap()
            .windows(b"svgBlip".len())
            .any(|w| w == b"svgBlip")
    );
    let completed = owner.namespace_complete(&drawing_xml, 16 * 1024).unwrap();
    assert!(
        completed
            .windows(SVG_NS.len())
            .any(|w| w == SVG_NS.as_bytes())
    );
    assert!(
        picture
            .picture_bytes(&drawing_xml)
            .unwrap()
            .windows(b"future:keep=\"yes\"".len())
            .any(|w| w == b"future:keep=\"yes\"")
    );
    assert!(
        picture
            .picture_bytes(&drawing_xml)
            .unwrap()
            .windows(b"<![CDATA[<?future-payload>]]>".len())
            .any(|w| w == b"<![CDATA[<?future-payload>]]>")
    );
}

#[test]
fn attach_save_reopen_detach_preserves_geometry_source_and_opaque_member() {
    let before = unusual_manifest_fixture();
    let untouched = part(&before, OPAQUE);
    let before_content_types = content_types_member(&before);
    assert!(
        !before_content_types
            .windows(b"image/svg+xml".len())
            .any(|window| window == b"image/svg+xml")
    );
    let attached = edit_attach(&before, selector(0), NEW_SVG).unwrap();
    assert_ne!(attached, before);
    assert_eq!(part(&attached, OPAQUE), untouched);
    let attached_svg = part_name_with_content_type(&attached, "image/svg+xml");
    assert_eq!(part(&attached, &attached_svg), NEW_SVG);
    assert_eq!(content_type(&attached, &attached_svg), "image/svg+xml");
    assert_ne!(content_types_member(&attached), before_content_types);
    assert!(
        content_types_member(&attached)
            .windows(b"source-manifest-comment".len())
            .any(|window| window == b"source-manifest-comment")
    );
    let attached_drawing = part(&attached, DRAWING);
    let attached_source = SourceDrawing::scan(&attached_drawing).unwrap();
    let before_drawing = part(&before, DRAWING);
    let before_source = SourceDrawing::scan(&before_drawing).unwrap();
    assert_eq!(
        attached_source.picture(0).unwrap().anchor(),
        before_source.picture(0).unwrap().anchor()
    );
    assert!(attached_source.picture(0).unwrap().is_direct_embedded_svg());
    assert!(
        String::from_utf8(part(&attached, DRAWING))
            .unwrap()
            .contains("asvg:svgBlip")
    );
    let attached_xml = String::from_utf8(part(&attached, DRAWING)).unwrap();
    assert!(attached_xml.contains("future:keep=\"yes\""));
    assert!(attached_xml.contains("<!--future-comment-->"));
    assert!(attached_xml.contains("<![CDATA[<?future-payload>]]>"));
    assert!(attached_xml.contains("<?future-processing?>"));
    let reopened = Workbook::from_bytes(attached.clone()).unwrap();
    assert_eq!(reopened.to_plain_bytes().unwrap(), attached);
    let detached = edit_detach(&attached, selector(0)).unwrap();
    assert_eq!(part(&detached, DRAWING), part(&before, DRAWING));
    assert_eq!(part(&detached, OPAQUE), untouched);
    assert_eq!(content_types_member(&detached), before_content_types);
    assert!(!has_content_type(&detached, "image/svg+xml"));
}

#[test]
fn namespace_only_and_opaque_extension_wrappers_survive_attach() {
    let before = namespace_only_fixture();
    let attached = edit_attach(&before, selector(0), NEW_SVG).unwrap();
    let xml = String::from_utf8(part(&attached, DRAWING)).unwrap();
    assert!(xml.contains(r#"<future:empty xmlns:future="urn:litchi:future"/>"#));
    assert!(xml.contains(&format!(r#"uri="{SVG_URI}""#)));
    assert!(xml.contains("asvg:svgBlip"));
}

#[test]
fn attach_each_selector_keeps_two_cell_one_cell_and_absolute_geometry() {
    let before = raster_fixture();
    let before_drawing = part(&before, DRAWING);
    let before_source = SourceDrawing::scan(&before_drawing).unwrap();
    for index in 0..3 {
        let attached = edit_attach(&before, selector(index), NEW_SVG).unwrap();
        let attached_drawing = part(&attached, DRAWING);
        let after_source = SourceDrawing::scan(&attached_drawing).unwrap();
        assert_eq!(
            after_source.picture(index).unwrap().anchor(),
            before_source.picture(index).unwrap().anchor()
        );
        assert!(
            after_source
                .picture(index)
                .unwrap()
                .is_direct_embedded_svg()
        );
    }
}

#[test]
fn multiple_pictures_attach_and_detach_in_one_atomic_transaction() {
    let before = raster_fixture();
    let workbook = Workbook::from_bytes(before.clone()).unwrap();
    let mut edit = workbook.edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        for picture in 0..3 {
            sheet
                .attach_svg(selector(picture), SvgInput::borrowed(NEW_SVG))
                .unwrap();
        }
    }
    assert_eq!(edit.len(), 3);
    let attached = edit.commit().unwrap().into_workbook();
    let attached_bytes = attached.to_plain_bytes().unwrap();
    let svg_parts = part_names_with_content_type(&attached_bytes, "image/svg+xml");
    assert_eq!(svg_parts.len(), 3);
    let drawing = part(&attached_bytes, DRAWING);
    let source = SourceDrawing::scan(&drawing).unwrap();
    assert!(
        source
            .pictures()
            .iter()
            .all(|picture| picture.is_direct_embedded_svg())
    );

    let mut detach_edit = attached.edit().unwrap();
    {
        let mut sheet = detach_edit.sheet("Sheet1").unwrap().unwrap();
        for picture in 0..3 {
            sheet.detach_svg(selector(picture)).unwrap();
        }
    }
    assert_eq!(detach_edit.len(), 3);
    let restored = detach_edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    assert_eq!(part(&restored, DRAWING), part(&before, DRAWING));
    assert_eq!(
        content_types_member(&restored),
        content_types_member(&before)
    );
    assert!(part_names_with_content_type(&restored, "image/svg+xml").is_empty());
}

#[test]
fn independent_worksheets_share_manifest_and_allocate_distinct_svg_parts() {
    let before = two_worksheet_fixture(false);
    let workbook = Workbook::from_bytes(before.clone()).unwrap();
    let mut edit = workbook.edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet
            .attach_svg(selector(0), SvgInput::borrowed(NEW_SVG))
            .unwrap();
    }
    {
        let mut sheet = edit.sheet("Sheet2").unwrap().unwrap();
        sheet
            .attach_svg(selector(0), SvgInput::borrowed(SVG_PAYLOAD))
            .unwrap();
    }
    assert_eq!(edit.len(), 2);
    let attached = edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    let svg_parts = part_names_with_content_type(&attached, "image/svg+xml");
    assert_eq!(svg_parts.len(), 2);
    assert_ne!(svg_parts[0], svg_parts[1]);
    assert_eq!(part(&attached, &svg_parts[0]).len(), NEW_SVG.len());
    assert_eq!(part(&attached, &svg_parts[1]).len(), SVG_PAYLOAD.len());
    let drawing1_rels = String::from_utf8(rels(&attached, DRAWING)).unwrap();
    let drawing2_rels = String::from_utf8(rels(&attached, DRAWING2)).unwrap();
    assert!(drawing1_rels.contains(&svg_parts[0].trim_start_matches("/xl/media/")));
    assert!(drawing2_rels.contains(&svg_parts[1].trim_start_matches("/xl/media/")));
    assert!(
        String::from_utf8(content_types_member(&attached))
            .unwrap()
            .matches("image/svg+xml")
            .count()
            >= 2
    );

    let mut detach_edit = Workbook::from_bytes(attached).unwrap().edit().unwrap();
    {
        let mut sheet = detach_edit.sheet("Sheet1").unwrap().unwrap();
        sheet.detach_svg(selector(0)).unwrap();
    }
    {
        let mut sheet = detach_edit.sheet("Sheet2").unwrap().unwrap();
        sheet.detach_svg(selector(0)).unwrap();
    }
    let restored = detach_edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    assert_eq!(part(&restored, DRAWING), part(&before, DRAWING));
    assert_eq!(part(&restored, DRAWING2), part(&before, DRAWING2));
    assert_eq!(
        content_types_member(&restored),
        content_types_member(&before)
    );
    assert!(part_names_with_content_type(&restored, "image/svg+xml").is_empty());
}

#[test]
fn shared_svg_target_across_worksheets_is_removed_after_last_same_transaction_owner() {
    let before = two_worksheet_fixture(true);
    let baseline = two_worksheet_fixture(false);
    let mut edit = Workbook::from_bytes(before.clone())
        .unwrap()
        .edit()
        .unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet.detach_svg(selector(0)).unwrap();
    }
    {
        let mut sheet = edit.sheet("Sheet2").unwrap().unwrap();
        sheet.detach_svg(selector(0)).unwrap();
    }
    let final_bytes = edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    assert!(!has_part(&final_bytes, SVG));
    assert!(part_names_with_content_type(&final_bytes, "image/svg+xml").is_empty());
    let final_manifest = String::from_utf8(content_types_member(&final_bytes)).unwrap();
    assert!(!final_manifest.contains("image/svg+xml"));
    assert_eq!(
        content_types_member(&final_bytes),
        content_types_member(&baseline)
    );
    assert!(
        !String::from_utf8(rels(&final_bytes, DRAWING))
            .unwrap()
            .contains("rIdSvg")
    );
    assert!(
        !String::from_utf8(rels(&final_bytes, DRAWING2))
            .unwrap()
            .contains("rIdSvg")
    );
}

#[test]
fn mixed_cell_and_svg_changes_publish_as_one_ordinary_transaction() {
    let before = raster_fixture();
    let mut edit = Workbook::from_bytes(before).unwrap().edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet
            .attach_svg(selector(0), SvgInput::borrowed(NEW_SVG))
            .unwrap()
            .set("A1", "mixed")
            .unwrap();
    }
    let committed = edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    assert!(
        String::from_utf8(part(&committed, SHEET))
            .unwrap()
            .contains("mixed")
    );
    assert!(
        String::from_utf8(part(&committed, DRAWING))
            .unwrap()
            .contains("asvg:svgBlip")
    );
}

#[test]
fn worksheet_direct_drawing_relationship_selection_and_hidden_owners_refuse_safely() {
    let orphan = orphan_relationship_fixture();
    let attached = edit_attach(&orphan, selector(0), NEW_SVG).unwrap();
    assert!(
        String::from_utf8(part(&attached, DRAWING))
            .unwrap()
            .contains("asvg:svgBlip")
    );
    assert!(
        !String::from_utf8(part(&attached, ORPHAN_DRAWING))
            .unwrap()
            .contains("svgBlip")
    );

    let missing = missing_direct_drawing_fixture();
    assert!(edit_attach(&missing, selector(0), NEW_SVG).is_err());

    let hidden = hidden_mce_drawing_fixture();
    assert!(edit_attach(&hidden, selector(0), NEW_SVG).is_err());
}

#[test]
fn worksheet_direct_drawing_accepts_entity_escaped_namespace_aliases() {
    let before = worksheet_namespace_alias_fixture();
    let before_sheet = part(&before, SHEET);
    let attached = edit_attach(&before, selector(0), NEW_SVG).unwrap();
    assert_eq!(part(&attached, SHEET), before_sheet);
    assert!(
        String::from_utf8(part(&attached, DRAWING))
            .unwrap()
            .contains("asvg:svgBlip")
    );
}

#[test]
fn duplicate_direct_worksheet_drawing_children_are_refused() {
    let duplicate = duplicate_direct_drawing_fixture();
    assert!(edit_attach(&duplicate, selector(0), NEW_SVG).is_err());
    assert!(edit_detach(&duplicate, selector(0)).is_err());
}

#[test]
fn hidden_mce_drawing_plus_direct_reference_is_refused() {
    let hidden_and_direct = mce_and_direct_drawing_fixture();
    assert!(edit_attach(&hidden_and_direct, selector(0), NEW_SVG).is_err());
    assert!(edit_detach(&hidden_and_direct, selector(0)).is_err());
}

#[test]
fn worksheet_extension_drawing_and_mce_lookalikes_remain_opaque() {
    let before = opaque_worksheet_extension_fixture();
    let before_sheet = part(&before, SHEET);
    let attached = edit_attach(&before, selector(0), NEW_SVG).unwrap();
    assert_eq!(part(&attached, SHEET), before_sheet);
    let sheet = String::from_utf8(part(&attached, SHEET)).unwrap();
    assert!(sheet.contains("urn:litchi:opaque"));
    assert!(sheet.contains("<s:drawing r:id=\"rIdDrawing\"/>"));
    assert!(sheet.contains("<mc:AlternateContent>"));
    assert!(
        String::from_utf8(part(&attached, DRAWING))
            .unwrap()
            .contains("asvg:svgBlip")
    );
}

#[test]
fn same_picture_attach_then_detach_projects_back_to_an_exact_noop() {
    let before = raster_fixture();
    let workbook = Workbook::from_bytes(before.clone()).unwrap();
    let mut edit = workbook.edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet
            .attach_svg(selector(0), SvgInput::borrowed(NEW_SVG))
            .unwrap()
            .detach_svg(selector(0))
            .unwrap();
    }
    assert_eq!(edit.len(), 2);
    let after = edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    assert_eq!(after, before);
}

#[test]
fn same_picture_detach_then_attach_projects_to_the_new_payload() {
    let before = svg_fixture();
    let workbook = Workbook::from_bytes(before.clone()).unwrap();
    let mut edit = workbook.edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet
            .detach_svg(selector(0))
            .unwrap()
            .attach_svg(selector(0), SvgInput::borrowed(NEW_SVG))
            .unwrap();
    }
    assert_eq!(edit.len(), 2);
    let after = edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    let svg_parts = part_names_with_content_type(&after, "image/svg+xml");
    assert_eq!(svg_parts.len(), 1);
    assert_eq!(part(&after, &svg_parts[0]), NEW_SVG);
    assert_ne!(part(&after, &svg_parts[0]), SVG_PAYLOAD);
    assert_eq!(part(&after, RASTER), part(&before, RASTER));
}

#[test]
fn strict_host_uses_strict_core_and_transitional_svg_attribute_namespace() {
    let before = package_bytes(DrawingOptions {
        strict: true,
        anchors: true,
        unknown_extension: true,
        ..DrawingOptions::default()
    });
    let attached = edit_attach(&before, selector(0), NEW_SVG).unwrap();
    let xml = part(&attached, DRAWING);
    let text = String::from_utf8(xml.clone()).unwrap();
    assert!(text.contains(STRICT_XDR));
    assert!(text.contains(STRICT_A));
    assert!(text.contains(STRICT_REL));
    assert!(text.contains(REL));
    let source = SourceDrawing::scan(&xml).unwrap();
    assert_eq!(source.dialect(), DrawingDialect::Strict);
    assert_eq!(source.relationship_dialect(), RelationshipDialect::Strict);
    assert_eq!(
        source
            .picture(0)
            .unwrap()
            .svg_owner()
            .owner()
            .unwrap()
            .relationship_dialect(),
        RelationshipDialect::Transitional
    );
}

#[test]
fn shared_svg_detach_removes_media_only_after_last_direct_owner() {
    let before = shared_svg_fixture(false);
    let first = edit_detach(&before, selector(0)).unwrap();
    assert!(has_part(&first, SVG));
    assert!(
        String::from_utf8(part(&first, DRAWING))
            .unwrap()
            .contains("rIdSvg")
    );
    let second = edit_detach(&first, selector(1)).unwrap();
    assert!(has_part(&second, SVG));
    let final_bytes = edit_detach(&second, selector(2)).unwrap();
    assert!(!has_part(&final_bytes, SVG));
    assert!(
        !String::from_utf8(part(&final_bytes, DRAWING))
            .unwrap()
            .contains("rIdSvg")
    );
    assert!(
        !String::from_utf8(rels(&final_bytes, DRAWING))
            .unwrap()
            .contains("rIdSvg")
    );
}

#[test]
fn incoming_other_part_edge_keeps_svg_media_after_last_picture_detach() {
    let before = shared_svg_fixture(true);
    let first = edit_detach(&before, selector(0)).unwrap();
    let second = edit_detach(&first, selector(1)).unwrap();
    let final_bytes = edit_detach(&second, selector(2)).unwrap();
    assert!(has_part(&final_bytes, SVG));
    assert_eq!(part(&final_bytes, SVG), SVG_PAYLOAD);
    assert!(
        !String::from_utf8(part(&final_bytes, DRAWING))
            .unwrap()
            .contains("rIdSvg")
    );
    assert!(
        String::from_utf8(rels(&final_bytes, OTHER))
            .unwrap()
            .contains("rIdIncomingSvg")
    );
}

#[test]
fn unknown_duplicate_mce_external_and_foreign_owners_are_safe_refusals() {
    let unknown = package_bytes(DrawingOptions {
        anchors: false,
        unknown_extension: true,
        foreign_descendant: true,
        ..DrawingOptions::default()
    });
    let unknown_drawing = part(&unknown, DRAWING);
    let unknown_source = SourceDrawing::scan(&unknown_drawing).unwrap();
    assert!(matches!(
        unknown_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Opaque | SvgOwnerState::None
    ));
    assert!(!unknown_source.picture(0).unwrap().is_direct_embedded_svg());
    assert!(
        unknown_source
            .picture(0)
            .unwrap()
            .relationship_references()
            .iter()
            .any(|reference| reference.id() == "rIdForeign")
    );
    let duplicate = package_bytes(DrawingOptions {
        anchors: false,
        admitted_svg: true,
        duplicate_svg: true,
        ..DrawingOptions::default()
    });
    let duplicate_drawing = part(&duplicate, DRAWING);
    let duplicate_source = SourceDrawing::scan(&duplicate_drawing).unwrap();
    assert!(matches!(
        duplicate_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Ambiguous
    ));
    assert!(edit_detach(&duplicate, selector(0)).is_err());
    let mce = package_bytes(DrawingOptions {
        anchors: false,
        mce_svg: true,
        ..DrawingOptions::default()
    });
    let mce_drawing = part(&mce, DRAWING);
    let mce_source = SourceDrawing::scan(&mce_drawing).unwrap();
    assert!(matches!(
        mce_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));
    assert!(edit_detach(&mce, selector(0)).is_err());
    let linked = package_bytes(DrawingOptions {
        anchors: false,
        admitted_svg: true,
        linked_svg: true,
        ..DrawingOptions::default()
    });
    let linked_drawing = part(&linked, DRAWING);
    let linked_source = SourceDrawing::scan(&linked_drawing).unwrap();
    assert!(matches!(
        linked_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Linked(_)
    ));
    assert!(edit_detach(&linked, selector(0)).is_err());
    assert!(edit_attach(&linked, selector(0), NEW_SVG).is_err());
}

#[test]
fn no_op_detach_exact_inverse_and_stale_patch_are_checked() {
    let before = raster_fixture();
    let detached = edit_detach(&before, selector(0)).unwrap();
    assert_eq!(detached, before);
    let mut noop_edit = Workbook::from_bytes(before.clone())
        .unwrap()
        .edit()
        .unwrap();
    {
        noop_edit
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .detach_svg(selector(0))
            .unwrap();
    }
    assert_eq!(noop_edit.len(), 0);
    assert!(noop_edit.is_empty());
    assert_eq!(
        noop_edit
            .commit()
            .unwrap()
            .into_workbook()
            .to_plain_bytes()
            .unwrap(),
        before
    );
    let workbook = Workbook::from_bytes(before.clone()).unwrap();
    let mut edit = workbook.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .attach_svg(selector(0), SvgInput::borrowed(NEW_SVG))
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.workbook().to_plain_bytes().unwrap();
    let durable = commit.patch().durable().unwrap();
    let inverse = commit.patch().inverse().durable().unwrap();
    assert_eq!(
        inverse
            .apply(&Workbook::from_bytes(after).unwrap())
            .unwrap()
            .to_plain_bytes()
            .unwrap(),
        before
    );
    let mut divergent = Workbook::from_bytes(before).unwrap().edit().unwrap();
    divergent
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", "stale")
        .unwrap();
    let divergent = divergent.commit().unwrap().into_workbook();
    assert!(matches!(
        durable.apply(&divergent),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn invalid_input_selector_and_custom_scan_limits_fail_before_publication() {
    let before = raster_fixture();
    let workbook = Workbook::from_bytes(before.clone()).unwrap();
    let mut edit = workbook.edit().unwrap();
    assert!(
        edit.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .attach_svg(selector(99), SvgInput::borrowed(NEW_SVG))
            .is_err()
    );
    assert!(
        edit.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .attach_svg(selector(0), SvgInput::borrowed(b""))
            .is_err()
    );
    let unchanged = edit
        .commit()
        .unwrap()
        .into_workbook()
        .to_plain_bytes()
        .unwrap();
    assert_eq!(unchanged, before);
    let drawing = part(&before, DRAWING);
    let limited = ScanLimits {
        max_xml_bytes: drawing.len() - 1,
        ..ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(&drawing, 0, limited).is_err());
    let limited_nodes = ScanLimits {
        max_nodes: 2,
        ..ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(&drawing, 0, limited_nodes).is_err());
    let source = SourceDrawing::scan(&drawing).unwrap();
    let owner = source.picture(0).unwrap();
    assert!(owner.namespace_complete_picture(&drawing, 1).is_err());
    let foreign_source = drawing.clone();
    assert!(owner.picture_bytes(&foreign_source).is_err());

    // Leave enough input budget to open the source member exactly, then make
    // the authored extension exceed that same per-part output ceiling. The
    // transaction must reject the publication before it can expose a partial
    // drawing, relationship, or content-types change.
    let output_cap = ReadLimits::builder()
        .max_part_bytes(drawing.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(edit_attach_with_limits(&before, selector(0), NEW_SVG, output_cap).is_err());
    assert_eq!(
        Workbook::from_bytes(before.clone())
            .unwrap()
            .to_plain_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn invalid_external_svg_and_png_fallbacks_fail_before_staging() {
    let linked_svg = package_bytes(DrawingOptions {
        anchors: false,
        admitted_svg: true,
        linked_svg: true,
        ..DrawingOptions::default()
    });
    assert!(edit_detach(&linked_svg, selector(0)).is_err());
    assert!(edit_attach(&linked_svg, selector(0), NEW_SVG).is_err());

    let external_raster = package_bytes(DrawingOptions {
        anchors: false,
        external_raster: true,
        ..DrawingOptions::default()
    });
    assert!(edit_attach(&external_raster, selector(0), NEW_SVG).is_err());
    assert!(edit_detach(&external_raster, selector(0)).is_err());

    let wrong_raster_type = package_bytes(DrawingOptions {
        anchors: false,
        wrong_raster_type: true,
        ..DrawingOptions::default()
    });
    assert!(edit_attach(&wrong_raster_type, selector(0), NEW_SVG).is_err());
    assert!(edit_detach(&wrong_raster_type, selector(0)).is_err());
}

#[test]
fn source_relationship_ranges_and_content_type_are_stable_across_inverse() {
    let before = raster_fixture();
    let before_content_types = content_types_member(&before);
    let before_drawing = part(&before, DRAWING);
    let source = SourceDrawing::scan(&before_drawing).unwrap();
    let picture = source.picture(0).unwrap();
    assert_eq!(picture.blip_range().prefix(), b"a");
    assert!(picture.blip_range().start_end() <= picture.blip_range().range().end);
    assert!(picture.ext_list_range().is_some());
    assert_eq!(content_type(&before, RASTER), "image/png");
    let workbook = Workbook::from_bytes(before.clone()).unwrap();
    let mut edit = workbook.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .attach_svg(selector(0), SvgInput::borrowed(NEW_SVG))
        .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let restored = patch
        .inverse()
        .durable()
        .unwrap()
        .apply(&after)
        .unwrap()
        .to_plain_bytes()
        .unwrap();
    assert_eq!(part(&restored, DRAWING), part(&before, DRAWING));
    assert_eq!(part(&restored, SHEET), part(&before, SHEET));
    assert_eq!(rels(&restored, DRAWING), rels(&before, DRAWING));
    assert_eq!(rels(&restored, SHEET), rels(&before, SHEET));
    assert_eq!(content_type(&restored, RASTER), "image/png");
    assert_eq!(content_types_member(&restored), before_content_types);
}
