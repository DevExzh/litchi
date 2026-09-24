//! Ordinary worksheet drawing/SVG read integration coverage.
//!
//! This test lane deliberately exercises the public worksheet facade and the
//! public source scanner. The older lifecycle lane owns mutation coverage;
//! this file is kept read-only so source/typed pairing, package graph
//! validation, and bounded XML ingress remain independently reviewable.

#![allow(
    clippy::unwrap_used,
    reason = "focused integration fixtures use panic-on-failure assertions"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::drawing::{
    DrawingAnchor, DrawingDialect, PictureSelector, RelationshipDialect, SourceDrawing,
    SvgOwnerState, WorksheetSourceLimits, WorksheetSourceScan,
};
use litchi_xlsx::{Package, ReadLimits, Workbook};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FOREIGN_MCE: &str = "http://purl.oclc.org/ooxml/markup-compatibility/2006";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const DRAWING: &str = "/xl/drawings/drawing1.xml";
const RASTER: &str = "/xl/media/image1.png";
const SVG: &str = "/xl/media/image2.svg";
const OTHER_MEDIA: &str = "/xl/media/other.bin";
const NEW_SVG: &[u8] = br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0 L4 4"/></svg>"#;
const PNG: &[u8] = b"synthetic-png-payload";

const NATIVE_XLSX: &[u8] = include_bytes!(
    "../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx"
);

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MediaFault {
    None,
    MissingRaster,
    WrongRasterType,
    WrongRasterContentType,
    ExternalRaster,
    RasterOutbound,
    MissingSvg,
    WrongSvgType,
    ExternalSvg,
    WrongSvgContentType,
    SvgOutbound,
}

fn marker(prefix: &str, name: &str, col: i64, col_off: i64, row: i64, row_off: i64) -> String {
    let marker = format!("{prefix}:{name}");
    let col_name = format!("{prefix}:col");
    let col_off_name = format!("{prefix}:colOff");
    let row_name = format!("{prefix}:row");
    let row_off_name = format!("{prefix}:rowOff");
    format!(
        "<{marker}><{col_name}>{col}</{col_name}><{col_off_name}>{col_off}</{col_off_name}><{row_name}>{row}</{row_name}><{row_off_name}>{row_off}</{row_off_name}></{marker}>"
    )
}

fn picture(strict: bool, index: usize, svg: bool, opaque: bool) -> String {
    let svg_rel = if strict { "trans" } else { "r" };
    let owner = if svg {
        format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip {svg_rel}:embed="rIdSvg"/></a:ext>"#)
    } else {
        String::new()
    };
    let opaque = if opaque {
        r#"<a:ext uri="urn:litchi:opaque"><!--opaque-comment--><![CDATA[opaque-cdata]]><?opaque-pi?><future:payload xmlns:future="urn:litchi:future" future:keep="yes"/></a:ext>"#.to_owned()
    } else {
        String::new()
    };
    let extensions = if owner.is_empty() && opaque.is_empty() {
        String::new()
    } else {
        format!(r#"<a:extLst>{owner}{opaque}</a:extLst>"#)
    };
    format!(
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="picture-{}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdRaster">{extensions}</a:blip></xdr:blipFill><xdr:spPr/></xdr:pic>"#,
        index + 1,
        index + 1,
    )
}

fn anchor(strict: bool, index: usize, svg: bool, opaque: bool) -> String {
    let picture = picture(strict, index, svg, opaque);
    match index {
        0 => format!(
            r#"<xdr:twoCellAnchor editAs="oneCell">{}{}{}<xdr:clientData/></xdr:twoCellAnchor>"#,
            marker("xdr", "from", 1, 2, 3, 4),
            marker("xdr", "to", 5, 6, 7, 8),
            picture,
        ),
        1 => format!(
            r#"<xdr:oneCellAnchor>{}<xdr:ext cx="123456" cy="654321"/>{}<xdr:clientData/></xdr:oneCellAnchor>"#,
            marker("xdr", "from", 9, 19, 10, 20),
            picture,
        ),
        _ => format!(
            r#"<xdr:absoluteAnchor><xdr:pos x="-900" y="456"/><xdr:ext cx="777888" cy="999000"/>{}<xdr:clientData/></xdr:absoluteAnchor>"#,
            picture,
        ),
    }
}

fn drawing_xml(strict: bool, picture_count: usize, svg: bool, opaque: bool) -> Vec<u8> {
    let root = if strict {
        format!(
            r#"<xdr:wsDr xmlns:xdr="{STRICT_XDR}" xmlns:a="{STRICT_A}" xmlns:r="{STRICT_REL}" xmlns:trans="{REL}" xmlns:asvg="{SVG_NS}" xmlns:mc="{MCE}" xmlns:future="urn:litchi:future">"#
        )
    } else {
        format!(
            r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{REL}" xmlns:asvg="{SVG_NS}" xmlns:mc="{MCE}" xmlns:future="urn:litchi:future">"#
        )
    };
    let body = (0..picture_count)
        .map(|index| {
            anchor(
                strict,
                index,
                svg && (index == 0 || picture_count == 2),
                opaque,
            )
        })
        .collect::<String>();
    format!("{root}{body}</xdr:wsDr>").into_bytes()
}

fn drawing_with_mce_wrapped_svg(strict: bool, namespace: &str, prefix: &str) -> Vec<u8> {
    let drawing = String::from_utf8(drawing_xml(strict, 1, true, false)).unwrap();
    let svg_owner = if strict {
        format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip trans:embed="rIdSvg"/></a:ext>"#)
    } else {
        format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#)
    };
    let wrapped = format!(
        r#"<{prefix}:AlternateContent xmlns:{prefix}="{namespace}" xmlns:future="urn:litchi:future" {prefix}:Ignorable="future"><{prefix}:Choice Requires="future">{svg_owner}</{prefix}:Choice><{prefix}:Fallback/></{prefix}:AlternateContent>"#
    );
    assert!(drawing.contains(&svg_owner));
    drawing.replace(&svg_owner, &wrapped).into_bytes()
}

fn worksheet_xml(strict: bool, direct: &str) -> Vec<u8> {
    let (main, rel) = if strict {
        (STRICT_SML, STRICT_REL)
    } else {
        (SML, REL)
    };
    format!(
        r#"<ws:worksheet xmlns:ws="{main}" xmlns:r="{rel}"><ws:dimension ref="A1:C3"/><ws:sheetData/>{direct}</ws:worksheet>"#
    )
    .into_bytes()
}

fn direct_drawing(strict: bool) -> String {
    let rel = if strict { STRICT_REL } else { REL };
    format!(r#"<ws:drawing xmlns:r="{rel}" r:id="rIdDrawing"/>"#)
}

fn package_bytes(
    strict: bool,
    picture_count: usize,
    svg: bool,
    worksheet: Option<Vec<u8>>,
    fault: MediaFault,
) -> Vec<u8> {
    let mut package = Package::create().unwrap().into_plain_opc();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(worksheet.unwrap_or_else(|| worksheet_xml(strict, &direct_drawing(strict))));
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(DRAWING).unwrap(),
            ct::OFC_DRAWING.to_owned(),
            drawing_xml(strict, picture_count, svg, false),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(RASTER).unwrap(),
            if fault == MediaFault::WrongRasterContentType {
                "image/jpeg".to_owned()
            } else {
                ct::PNG.to_owned()
            },
            PNG.to_vec(),
        )))
        .unwrap();
    if svg {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SVG).unwrap(),
                if fault == MediaFault::WrongSvgContentType {
                    "image/png".to_owned()
                } else {
                    "image/svg+xml".to_owned()
                },
                NEW_SVG.to_vec(),
            )))
            .unwrap();
    }
    if matches!(fault, MediaFault::RasterOutbound | MediaFault::SvgOutbound) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(OTHER_MEDIA).unwrap(),
                "application/octet-stream".to_owned(),
                b"outbound-target".to_vec(),
            )))
            .unwrap();
    }

    let sheet = package.get_part_mut(&PackURI::new(SHEET).unwrap()).unwrap();
    sheet
        .rels_mut()
        .try_add_relationship(
            if strict {
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
    if fault != MediaFault::MissingRaster {
        let raster_type = if fault == MediaFault::WrongRasterType {
            "urn:litchi:wrong-image".to_owned()
        } else if fault == MediaFault::ExternalRaster {
            "urn:litchi:external-image".to_owned()
        } else if strict {
            rt::STRICT_IMAGE.to_owned()
        } else {
            rt::IMAGE.to_owned()
        };
        let (target, mode) = if fault == MediaFault::ExternalRaster {
            (
                "https://example.invalid/raster.png".to_owned(),
                TargetMode::External,
            )
        } else {
            ("../media/image1.png".to_owned(), TargetMode::Internal)
        };
        drawing
            .rels_mut()
            .try_add_relationship(raster_type, target, "rIdRaster".to_owned(), mode)
            .unwrap();
    }
    if svg && fault != MediaFault::MissingSvg {
        let relation_type = if fault == MediaFault::WrongSvgType {
            "urn:litchi:wrong-svg".to_owned()
        } else if fault == MediaFault::ExternalSvg {
            "urn:litchi:external-svg".to_owned()
        } else if strict {
            rt::STRICT_IMAGE.to_owned()
        } else {
            rt::IMAGE.to_owned()
        };
        let (target, mode) = if fault == MediaFault::ExternalSvg {
            (
                "https://example.invalid/image.svg".to_owned(),
                TargetMode::External,
            )
        } else {
            ("../media/image2.svg".to_owned(), TargetMode::Internal)
        };
        drawing
            .rels_mut()
            .try_add_relationship(relation_type, target, "rIdSvg".to_owned(), mode)
            .unwrap();
    }
    if fault == MediaFault::RasterOutbound {
        package
            .get_part_mut(&PackURI::new(RASTER).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "other.bin".to_owned(),
                "rIdOutbound".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    if fault == MediaFault::SvgOutbound {
        package
            .get_part_mut(&PackURI::new(SVG).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "other.bin".to_owned(),
                "rIdOutbound".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn synthetic_fixture(strict: bool, picture_count: usize, svg: bool) -> Vec<u8> {
    package_bytes(strict, picture_count, svg, None, MediaFault::None)
}

fn rewrite_sheet(bytes: &[u8], source: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(source);
    PackageWriter::to_bytes(&package).unwrap()
}

fn rewrite_physical_member(
    bytes: &[u8],
    wanted: &str,
    rewrite: impl FnOnce(Vec<u8>) -> Vec<u8>,
) -> Vec<u8> {
    let reader = PhysPkgReader::new(bytes).unwrap();
    let members = reader.member_names().unwrap();
    let mut writer = PhysPkgWriter::new();
    let mut rewrite = Some(rewrite);
    for member in members {
        let payload = reader.read_member(&member).unwrap();
        let payload = if member == wanted {
            rewrite.take().unwrap()(payload)
        } else {
            payload
        };
        writer
            .write(&PackURI::new(format!("/{member}")).unwrap(), &payload)
            .unwrap();
    }
    writer.finish().unwrap()
}

fn case_variant_svg_target_fixture(bytes: &[u8]) -> Vec<u8> {
    rewrite_physical_member(
        bytes,
        "xl/drawings/_rels/drawing1.xml.rels",
        |relationships| {
            let relationships = String::from_utf8(relationships).unwrap();
            assert!(relationships.contains(r#"Target="../media/image2.svg""#));
            relationships
                .replace(
                    r#"Target="../media/image2.svg""#,
                    r#"Target="../media/IMAGE2.SVG""#,
                )
                .into_bytes()
        },
    )
}

fn encoded_drawing_namespace_fixture(bytes: &[u8]) -> Vec<u8> {
    rewrite_physical_member(bytes, "xl/drawings/drawing1.xml", |drawing| {
        let drawing = String::from_utf8(drawing).unwrap();
        let encoded_xdr = XDR.replace('/', "&#x2F;");
        let encoded_a = A.replace('/', "&#x2F;");
        let encoded_rel = REL.replace('/', "&#x2F;");
        drawing
            .replace(
                &format!(r#"xmlns:xdr="{XDR}""#),
                &format!(r#"xmlns:xdr="{encoded_xdr}""#),
            )
            .replace(
                &format!(r#"xmlns:a="{A}""#),
                &format!(r#"xmlns:a="{encoded_a}""#),
            )
            .replace(
                &format!(r#"xmlns:r="{REL}""#),
                &format!(r#"xmlns:r="{encoded_rel}""#),
            )
            .into_bytes()
    })
}

fn workbook(bytes: &[u8]) -> Workbook {
    Workbook::from_bytes(bytes.to_vec()).unwrap()
}

fn worksheet_source(inner: &str) -> Vec<u8> {
    format!(
        r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}"><ws:dimension ref="A1"/><ws:sheetData/>{inner}</ws:worksheet>"#
    )
    .into_bytes()
}

fn nested_worksheet_source() -> Vec<u8> {
    worksheet_source(
        r#"<ws:sheetData><ws:row><ws:c/></ws:row></ws:sheetData><ws:drawing r:id="rIdDrawing"/>"#,
    )
}

fn source_limits_reject(source: &[u8], mutate: impl FnOnce(&mut WorksheetSourceLimits)) {
    let mut limits = WorksheetSourceLimits::default();
    mutate(&mut limits);
    assert!(WorksheetSourceScan::scan_with_limits(source, limits).is_err());
}

#[test]
fn native_fixture_pairs_typed_source_and_borrows_svg_payload() {
    let workbook = workbook(NATIVE_XLSX);
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();

    assert_eq!(drawing.source().dialect(), DrawingDialect::Transitional);
    assert_eq!(
        drawing.source().relationship_dialect(),
        RelationshipDialect::Transitional
    );
    assert_eq!(drawing.picture_count(), 2);
    assert_eq!(drawing.typed().pictures().count(), 2);
    assert!(std::ptr::eq(
        drawing.source().source().as_ptr(),
        drawing.source_xml().as_ptr()
    ));

    for position in 0..2 {
        let picture = drawing.picture(position).unwrap();
        let typed = drawing
            .source()
            .match_typed_picture(drawing.typed(), position)
            .unwrap();
        assert_eq!(picture.source().picture_ordinal(), position);
        assert_eq!(
            picture.source().raster_relationship_id(),
            typed.relationship_id
        );
        assert_eq!(picture.anchor(), typed.drawing_anchor());
        assert!(matches!(picture.anchor(), DrawingAnchor::TwoCell { .. }));
        let descriptor = picture.svg().unwrap().unwrap();
        assert_eq!(descriptor.relationship_id(), "rId2");
        assert_eq!(descriptor.part_uri().as_str(), SVG);
        assert_eq!(descriptor.content_type(), "image/svg+xml");
        assert_eq!(descriptor.owner().embedded_relationship_id(), Some("rId2"));
        assert_eq!(
            descriptor.reference().embedded.as_ref().unwrap().as_str(),
            "rId2"
        );
        let image = picture.read_svg_image().unwrap();
        assert_eq!(image.bytes().len(), 313);
    }

    let first = worksheet
        .read_svg_image(PictureSelector::new(0, 0))
        .unwrap();
    let first_ptr = first.bytes().as_ptr();
    let second = worksheet
        .read_svg_image(PictureSelector::new(0, 0))
        .unwrap();
    assert_eq!(first.bytes(), second.bytes());
    assert_eq!(first_ptr, second.bytes().as_ptr());
    let workbook_read = workbook
        .read_svg_image("Sheet1", PictureSelector::new(0, 1))
        .unwrap();
    assert_eq!(workbook_read.bytes(), first.bytes());
}

#[test]
fn synthetic_all_anchor_forms_pair_source_and_typed_geometry() {
    let workbook = workbook(&synthetic_fixture(false, 3, true));
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    assert_eq!(drawing.picture_count(), 3);
    assert!(matches!(
        drawing.picture(0).unwrap().anchor(),
        DrawingAnchor::TwoCell { .. }
    ));
    assert!(matches!(
        drawing.picture(1).unwrap().anchor(),
        DrawingAnchor::OneCell { .. }
    ));
    assert!(matches!(
        drawing.picture(2).unwrap().anchor(),
        DrawingAnchor::Absolute { .. }
    ));
    for position in 0..3 {
        let picture = drawing.picture(position).unwrap();
        let typed = drawing
            .source()
            .match_typed_picture(drawing.typed(), position)
            .unwrap();
        assert_eq!(picture.source().picture_ordinal(), position);
        assert_eq!(picture.anchor(), typed.drawing_anchor());
        assert_eq!(picture.source().raster_relationship_id(), "rIdRaster");
    }
    let svg = drawing.picture(0).unwrap().read_svg_image().unwrap();
    assert_eq!(svg.bytes(), NEW_SVG);
    assert!(matches!(
        drawing.picture(1).unwrap().svg_owner(),
        SvgOwnerState::None
    ));
}

#[test]
fn strict_core_with_transitional_svg_attribute_is_read_as_strict() {
    let workbook = workbook(&synthetic_fixture(true, 1, true));
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    assert_eq!(drawing.source().dialect(), DrawingDialect::Strict);
    assert_eq!(
        drawing.source().relationship_dialect(),
        RelationshipDialect::Strict
    );
    let picture = drawing.picture(0).unwrap();
    let owner = picture.svg_owner().owner().unwrap();
    assert_eq!(owner.embedded_relationship_id(), Some("rIdSvg"));
    assert_eq!(
        owner.relationship_dialect(),
        RelationshipDialect::Transitional
    );
    assert_eq!(picture.read_svg_image().unwrap().bytes(), NEW_SVG);
}

#[test]
fn strict_mce_is_canonical_and_foreign_purl_wrappers_stay_opaque() {
    let canonical = drawing_with_mce_wrapped_svg(true, MCE, "mc");
    let canonical_source = SourceDrawing::scan(&canonical).unwrap();
    assert_eq!(canonical_source.dialect(), DrawingDialect::Strict);
    assert_eq!(canonical_source.pictures().len(), 1);
    assert!(matches!(
        canonical_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));

    let foreign = drawing_with_mce_wrapped_svg(true, FOREIGN_MCE, "pmc");
    let foreign_source = SourceDrawing::scan(&foreign).unwrap();
    assert_eq!(foreign_source.dialect(), DrawingDialect::Strict);
    assert_eq!(foreign_source.pictures().len(), 1);
    let foreign_picture = foreign_source.picture(0).unwrap();
    assert!(matches!(foreign_picture.svg_owner(), SvgOwnerState::None));
    let picture_bytes = foreign_picture.picture_bytes(&foreign).unwrap();
    assert!(
        picture_bytes
            .windows(b"pmc:AlternateContent".len())
            .any(|window| window == b"pmc:AlternateContent")
    );
    assert!(
        picture_bytes
            .windows(b"pmc:Ignorable=\"future\"".len())
            .any(|window| window == b"pmc:Ignorable=\"future\"")
    );
}

#[test]
fn worksheet_namespace_aliases_and_opaque_lookalikes_keep_direct_ownership() {
    let alias = String::from(
        r#"<ws:worksheet xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/mai&#110;" xmlns:rr="http://schemas.openxmlformats.org/officeDocument/2006/relationship&#115;"><ws:sheetData/><ws:drawing rr:id="rIdDrawing"/></ws:worksheet>"#,
    );
    let scan = WorksheetSourceScan::scan(alias.as_bytes()).unwrap();
    assert_eq!(scan.relationship_ids().collect::<Vec<_>>(), ["rIdDrawing"]);
    let bytes = rewrite_sheet(&synthetic_fixture(false, 1, true), alias.into_bytes());
    let worksheet = workbook(&bytes).sheet("Sheet1").unwrap().unwrap();
    assert_eq!(worksheet.drawing(0).unwrap().picture_count(), 1);

    let opaque = format!(
        r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}" xmlns:mc="{MCE}" xmlns:foreign="urn:litchi:foreign"><ws:sheetData/><ws:drawing r:id="rIdDrawing"/><ws:extLst><foreign:payload><ws:drawing r:id="rIdWrong"/><mc:AlternateContent><mc:Choice Requires="foreign"><ws:drawing r:id="rIdWrong"/></mc:Choice><mc:Fallback/></mc:AlternateContent></foreign:payload></ws:extLst></ws:worksheet>"#
    );
    let scan = WorksheetSourceScan::scan(opaque.as_bytes()).unwrap();
    assert_eq!(scan.drawings().len(), 1);
    let bytes = rewrite_sheet(&synthetic_fixture(false, 1, true), opaque.into_bytes());
    let worksheet = workbook(&bytes).sheet("Sheet1").unwrap().unwrap();
    assert_eq!(worksheet.drawing(0).unwrap().picture_count(), 1);
}

#[test]
fn case_variant_svg_media_target_resolves_the_canonical_package_part() {
    let bytes = case_variant_svg_target_fixture(&synthetic_fixture(false, 1, true));
    let worksheet = workbook(&bytes).sheet("Sheet1").unwrap().unwrap();
    let image = worksheet
        .read_svg_image(PictureSelector::new(0, 0))
        .unwrap();
    assert_eq!(image.bytes(), NEW_SVG);
    assert_eq!(image.descriptor().part_uri().as_str(), SVG);
}

#[test]
fn entity_escaped_drawing_namespace_uris_pair_source_and_typed_reads() {
    let bytes = encoded_drawing_namespace_fixture(&synthetic_fixture(false, 1, true));
    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let drawing_bytes = package
        .get_part(&PackURI::new(DRAWING).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let source = SourceDrawing::scan(&drawing_bytes).unwrap();
    assert_eq!(source.dialect(), DrawingDialect::Transitional);
    assert_eq!(source.pictures().len(), 1);
    assert_eq!(
        source.picture(0).unwrap().raster_relationship_id(),
        "rIdRaster"
    );

    let workbook = workbook(&bytes);
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    assert_eq!(drawing.picture_count(), 1);
    assert_eq!(drawing.typed().pictures().count(), 1);
    assert_eq!(
        worksheet
            .read_svg_image(PictureSelector::new(0, 0))
            .unwrap()
            .bytes(),
        NEW_SVG
    );
}

#[test]
fn mce_foreign_wrappers_are_opaque_but_direct_choice_drawings_refuse() {
    let foreign_wrapper = format!(
        r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}" xmlns:mc="{MCE}" xmlns:foreign="urn:litchi:foreign"><ws:sheetData/><mc:AlternateContent><mc:Choice Requires="foreign"><foreign:wrapper><ws:drawing r:id="rIdHidden"/></foreign:wrapper></mc:Choice><mc:Fallback/></mc:AlternateContent></ws:worksheet>"#
    );
    let scan = WorksheetSourceScan::scan(foreign_wrapper.as_bytes()).unwrap();
    assert!(scan.drawings().is_empty());
    let bytes = rewrite_sheet(
        &synthetic_fixture(false, 1, true),
        foreign_wrapper.into_bytes(),
    );
    let worksheet = workbook(&bytes).sheet("Sheet1").unwrap().unwrap();
    assert!(worksheet.drawing(0).is_err());

    let direct_choice = format!(
        r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}" xmlns:mc="{MCE}" xmlns:x="urn:litchi:choice"><ws:sheetData/><mc:AlternateContent><mc:Choice Requires="x"><ws:drawing r:id="rIdHidden"/></mc:Choice><mc:Fallback/></mc:AlternateContent></ws:worksheet>"#
    );
    assert!(WorksheetSourceScan::scan(direct_choice.as_bytes()).is_err());
}

#[test]
fn duplicate_direct_and_hidden_mce_drawing_owners_are_refused() {
    let duplicate = worksheet_source(r#"<ws:drawing r:id="rIdOne"/><ws:drawing r:id="rIdTwo"/>"#);
    assert!(WorksheetSourceScan::scan(&duplicate).is_err());

    let canonical = format!(
        r#"<ws:worksheet xmlns:ws="{STRICT_SML}" xmlns:r="{STRICT_REL}" xmlns:mc="{MCE}" xmlns:x="urn:litchi:choice"><ws:sheetData/><mc:AlternateContent><mc:Choice Requires="x"><ws:drawing r:id="rIdHidden"/></mc:Choice><mc:Fallback/></mc:AlternateContent><ws:drawing r:id="rIdDirect"/></ws:worksheet>"#
    );
    assert!(WorksheetSourceScan::scan(canonical.as_bytes()).is_err());
    let bytes = rewrite_sheet(&synthetic_fixture(true, 1, true), canonical.into_bytes());
    let worksheet = workbook(&bytes).sheet("Sheet1").unwrap().unwrap();
    assert!(worksheet.drawing(0).is_err());

    let foreign = format!(
        r#"<ws:worksheet xmlns:ws="{STRICT_SML}" xmlns:r="{STRICT_REL}" xmlns:pmc="{FOREIGN_MCE}" xmlns:x="urn:litchi:choice"><ws:sheetData/><pmc:AlternateContent pmc:Ignorable="x"><pmc:Choice Requires="x"><ws:drawing r:id="rIdHidden"/></pmc:Choice><pmc:Fallback/></pmc:AlternateContent><ws:drawing r:id="rIdDrawing"/></ws:worksheet>"#
    );
    let scan = WorksheetSourceScan::scan(foreign.as_bytes()).unwrap();
    assert_eq!(scan.relationship_ids().collect::<Vec<_>>(), ["rIdDrawing"]);
    let bytes = rewrite_sheet(&synthetic_fixture(true, 1, true), foreign.into_bytes());
    let worksheet = workbook(&bytes).sheet("Sheet1").unwrap().unwrap();
    assert_eq!(worksheet.drawing(0).unwrap().picture_count(), 1);

    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let retained = package
        .get_part(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .blob();
    assert!(
        retained
            .windows(b"pmc:AlternateContent".len())
            .any(|window| window == b"pmc:AlternateContent")
    );
    assert!(
        retained
            .windows(b"pmc:Ignorable=\"x\"".len())
            .any(|window| window == b"pmc:Ignorable=\"x\"")
    );
}

#[test]
fn worksheet_source_preserves_legal_opaque_events_and_rejects_xml_seam_errors() {
    let legal = format!(
        r#"<?keep?><ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}"><!--comment--><ws:sheetData><![CDATA[opaque]]></ws:sheetData><?after?><ws:drawing r:id="rIdDrawing"/></ws:worksheet>"#
    );
    assert_eq!(
        WorksheetSourceScan::scan(legal.as_bytes())
            .unwrap()
            .relationship_ids()
            .collect::<Vec<_>>(),
        ["rIdDrawing"]
    );

    let malformed = [
        worksheet_source(r#"<ws:sheetData title="&foo;"/><ws:drawing r:id="rIdDrawing"/>"#),
        worksheet_source(r#"<ws:sheetData title="1<2"/><ws:drawing r:id="rIdDrawing"/>"#),
        worksheet_source(r#"<ws:sheetData>]]></ws:sheetData><ws:drawing r:id="rIdDrawing"/>"#),
        format!(
            "junk{}",
            String::from_utf8(worksheet_source(r#"<ws:drawing r:id="rIdDrawing"/>"#)).unwrap()
        )
        .into_bytes(),
        format!(
            "{}tail",
            String::from_utf8(worksheet_source(r#"<ws:drawing r:id="rIdDrawing"/>"#)).unwrap()
        )
        .into_bytes(),
        format!(
            r#"<?xml version="2.0"?>{}"#,
            String::from_utf8(worksheet_source(r#"<ws:drawing r:id="rIdDrawing"/>"#)).unwrap()
        )
        .into_bytes(),
        format!(
            r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}" xmlns:p=""><ws:drawing r:id="rIdDrawing"/></ws:worksheet>"#
        )
        .into_bytes(),
        worksheet_source(
            r#"<ws:extLst><ws:ext uri="urn:opaque" data="&unknown;"/></ws:extLst><ws:drawing r:id="rIdDrawing"/>"#,
        ),
    ];
    for source in malformed {
        assert!(
            WorksheetSourceScan::scan(&source).is_err(),
            "malformed worksheet source was admitted: {:?}",
            String::from_utf8_lossy(&source)
        );
    }
}

#[test]
fn source_drawing_retains_opaque_comment_cdata_and_pi_bytes() {
    let source_bytes = drawing_xml(false, 1, true, true);
    let source = SourceDrawing::scan(&source_bytes).unwrap();
    assert_eq!(source.source(), source_bytes.as_slice());
    let picture = source.picture(0).unwrap();
    let bytes = picture.picture_bytes(&source_bytes).unwrap();
    assert!(
        bytes
            .windows(b"opaque-comment".len())
            .any(|w| w == b"opaque-comment")
    );
    assert!(
        bytes
            .windows(b"opaque-cdata".len())
            .any(|w| w == b"opaque-cdata")
    );
    assert!(bytes.windows(b"opaque-pi".len()).any(|w| w == b"opaque-pi"));
}

#[test]
fn worksheet_source_caller_caps_are_enforced_before_reference_publication() {
    let source = worksheet_source(r#"<ws:drawing r:id="rIdDrawing"/>"#);
    source_limits_reject(&source, |limits| limits.max_xml_bytes = source.len() - 1);
    source_limits_reject(&source, |limits| limits.max_nodes = 1);
    source_limits_reject(&nested_worksheet_source(), |limits| limits.max_depth = 1);
    source_limits_reject(&source, |limits| limits.max_drawing_references = 0);
    source_limits_reject(&source, |limits| limits.max_attributes = 1);
    source_limits_reject(&source, |limits| limits.max_namespace_declarations = 1);
    source_limits_reject(&source, |limits| limits.max_active_namespace_bindings = 1);
    source_limits_reject(&source, |limits| limits.max_namespace_bytes = 4);
    source_limits_reject(&source, |limits| limits.max_attribute_value_bytes = 1);
    source_limits_reject(&source, |limits| limits.max_relationship_id_bytes = 2);
}

#[test]
fn worksheet_source_empty_and_exact_limit_boundaries_are_deterministic() {
    let empty = format!(r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}"/>"#);
    let empty_scan = WorksheetSourceScan::scan(empty.as_bytes()).unwrap();
    assert!(empty_scan.drawings().is_empty());

    let root_depth_zero = WorksheetSourceLimits {
        max_depth: 0,
        ..WorksheetSourceLimits::default()
    };
    assert!(WorksheetSourceScan::scan_with_limits(empty.as_bytes(), root_depth_zero).is_err());
    let empty_child =
        format!(r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}"><ws:sheetData/></ws:worksheet>"#);
    let child_depth_one = WorksheetSourceLimits {
        max_depth: 1,
        ..WorksheetSourceLimits::default()
    };
    assert!(
        WorksheetSourceScan::scan_with_limits(empty_child.as_bytes(), child_depth_one).is_err()
    );

    let second_root = format!(r#"{empty}<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}"/>"#);
    assert!(WorksheetSourceScan::scan(second_root.as_bytes()).is_err());

    let source = worksheet_source(r#"<ws:drawing r:id="rIdDrawing"/>"#);
    let exact = WorksheetSourceLimits {
        max_xml_bytes: source.len(),
        max_namespace_bytes: REL.len(),
        max_namespace_declarations: 2,
        max_active_namespace_bindings: 2,
        max_attributes: 2,
        max_attribute_value_bytes: REL.len(),
        max_relationship_id_bytes: "rIdDrawing".len(),
        ..WorksheetSourceLimits::default()
    };
    assert_eq!(
        WorksheetSourceScan::scan_with_limits(&source, exact)
            .unwrap()
            .relationship_ids()
            .collect::<Vec<_>>(),
        ["rIdDrawing"]
    );
}

#[test]
fn caller_part_xml_cap_refuses_mce_heavy_worksheet_before_read_projection() {
    let choices = (0..24)
        .map(|index| {
            format!(
                r#"<mc:Choice Requires="x"><foreign:payload data="{index}"><foreign:node/></foreign:payload></mc:Choice>"#
            )
        })
        .collect::<String>();
    let source = format!(
        r#"<ws:worksheet xmlns:ws="{SML}" xmlns:r="{REL}" xmlns:mc="{MCE}" xmlns:foreign="urn:litchi:foreign" xmlns:x="urn:litchi:choice"><ws:sheetData/><mc:AlternateContent>{choices}<mc:Fallback/></mc:AlternateContent><ws:drawing r:id="rIdDrawing"/></ws:worksheet>"#
    );
    let bytes = rewrite_sheet(
        &synthetic_fixture(false, 1, true),
        source.clone().into_bytes(),
    );
    let limits = ReadLimits::builder()
        .max_part_bytes((source.len() - 1) as u64)
        .unwrap();
    assert!(Workbook::from_bytes_with_limits(bytes, limits.build().unwrap()).is_err());
}

#[test]
fn media_relationship_graph_faults_are_refused_without_payload_guessing() {
    for fault in [
        MediaFault::MissingRaster,
        MediaFault::WrongRasterType,
        MediaFault::WrongRasterContentType,
        MediaFault::ExternalRaster,
        MediaFault::RasterOutbound,
    ] {
        let workbook = workbook(&package_bytes(false, 1, true, None, fault));
        let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
        assert!(
            worksheet.picture(PictureSelector::new(0, 0)).is_err(),
            "raster graph fault {fault:?} was admitted"
        );
    }
    for fault in [
        MediaFault::MissingSvg,
        MediaFault::WrongSvgType,
        MediaFault::ExternalSvg,
        MediaFault::WrongSvgContentType,
        MediaFault::SvgOutbound,
    ] {
        let workbook = workbook(&package_bytes(false, 1, true, None, fault));
        let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
        let picture = worksheet.picture(PictureSelector::new(0, 0)).unwrap();
        assert!(
            picture.svg().is_err(),
            "SVG graph fault {fault:?} was admitted"
        );
        assert!(picture.read_svg_image().is_err());
    }
}
