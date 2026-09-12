//! Source-backed SpreadsheetDrawing `xdr:contentPart` coverage.
//!
//! The fixtures are public OPC graphs assembled from the direct core
//! transitional and strict SpreadsheetDrawing grammars.  The Office 2010
//! `xdr14:contentPart` group profile remains deliberately out of this test
//! owner until its relationship contract is closed by the design record.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused XML/OPC fixtures use panic-on-failure assertions"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::drawing::{
    ContentPartProfile, DrawingAnchor, DrawingDialect, EditAs, Object, RelationshipDialect,
    SourceDrawing, UnknownKind,
};
use litchi_xlsx::{Package, ReadLimits, Workbook};

const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FOREIGN_MCE: &str = "urn:litchi:foreign-mce";
const XDR14: &str = "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const DRAWING: &str = "/xl/drawings/drawing1.xml";
const TARGET_PREFIX: &str = "/xl/customXml/item";
const UNRELATED: &str = "/xl/customXml/unrelated.bin";
const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";

fn marker(prefix: &str, name: &str, col: i64, col_off: i64, row: i64, row_off: i64) -> String {
    format!(
        r#"<{prefix}:{name}><{prefix}:col>{col}</{prefix}:col><{prefix}:colOff>{col_off}</{prefix}:colOff><{prefix}:row>{row}</{prefix}:row><{prefix}:rowOff>{row_off}</{prefix}:rowOff></{prefix}:{name}>"#
    )
}

fn core_content_part(prefix: &str, rel_prefix: &str, id: &str) -> String {
    format!(r#"<{prefix}:contentPart {rel_prefix}:id="{id}"/>"#)
}

fn core_picture(prefix: &str, rel_prefix: &str, a_prefix: &str, index: usize) -> String {
    let picture_id = index + 1;
    format!(
        r#"<{prefix}:pic><{prefix}:nvPicPr><{prefix}:cNvPr id="{picture_id}" name="picture-{picture_id}"/><{prefix}:cNvPicPr/></{prefix}:nvPicPr><{prefix}:blipFill><{a_prefix}:blip {rel_prefix}:embed="rIdPicture{index}"/></{prefix}:blipFill><{prefix}:spPr/></{prefix}:pic>"#
    )
}

fn core_picture_anchor(prefix: &str, rel_prefix: &str, a_prefix: &str, index: usize) -> String {
    let template = core_anchor(prefix, rel_prefix, index);
    let owner = core_content_part(prefix, rel_prefix, &format!("rIdContent{index}"));
    template.replace(&owner, &core_picture(prefix, rel_prefix, a_prefix, index))
}

fn core_anchor(prefix: &str, rel_prefix: &str, index: usize) -> String {
    let owner = core_content_part(prefix, rel_prefix, &format!("rIdContent{index}"));
    match index {
        0 => format!(
            r#"<{prefix}:twoCellAnchor editAs="oneCell">{}{}{owner}<{prefix}:clientData/></{prefix}:twoCellAnchor>"#,
            marker(prefix, "from", 1, 2, 3, 4),
            marker(prefix, "to", 5, 6, 7, 8),
        ),
        1 => format!(
            r#"<{prefix}:oneCellAnchor>{}<{}:ext cx="123456" cy="654321"/>{owner}<{prefix}:clientData/></{prefix}:oneCellAnchor>"#,
            marker(prefix, "from", 9, 19, 10, 20),
            prefix,
        ),
        _ => format!(
            r#"<{prefix}:absoluteAnchor><{prefix}:pos x="-900" y="456"/><{prefix}:ext cx="777888" cy="999000"/>{owner}<{prefix}:clientData/></{prefix}:absoluteAnchor>"#
        ),
    }
}

fn direct_core_drawing(strict: bool) -> Vec<u8> {
    direct_core_drawing_count(strict, 3)
}

fn direct_core_drawing_count(strict: bool, count: usize) -> Vec<u8> {
    let (xdr, rel, prefix) = if strict {
        (STRICT_XDR, STRICT_REL, "xdr")
    } else {
        (XDR, REL, "xdr")
    };
    let anchors = (0..count)
        .map(|index| core_anchor(prefix, "r", index))
        .collect::<String>();
    format!(r#"<{prefix}:wsDr xmlns:{prefix}="{xdr}" xmlns:r="{rel}">{anchors}</{prefix}:wsDr>"#)
        .into_bytes()
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PackageFault {
    None,
    MissingRelation,
    WrongRelationType,
    ExternalRelation,
    DanglingTarget,
    WrongTargetContentType,
    MissingTargetContentType,
    MalformedTarget,
    InvalidQNameTarget,
    UnknownEntityTarget,
    RawAttributeTarget,
    RawCdataCloseTarget,
    ControlCharacterTarget,
    BadXmlDeclarationTarget,
    BadXmlEncodingTarget,
    BadXmlVersionTarget,
    TextBeforeRootTarget,
    TextAfterRootTarget,
    CdataOutsideRootTarget,
    UnboundNamespaceTarget,
    EmptyNamespaceTarget,
    ReservedXmlNamespaceTarget,
    ReservedXmlnsNamespaceTarget,
    EmptyTarget,
    EmptyRootTarget,
    NestedEmptyTarget,
    NestedDepthTarget,
    LegalOpaqueTarget,
    SharedTarget,
    OutboundGraph,
}

fn worksheet_xml(strict: bool) -> Vec<u8> {
    let (sml, rel) = if strict {
        (STRICT_SML, STRICT_REL)
    } else {
        (SML, REL)
    };
    format!(
        r#"<ws:worksheet xmlns:ws="{sml}" xmlns:r="{rel}"><ws:dimension ref="A1"/><ws:sheetData/><ws:drawing r:id="rIdDrawing"/></ws:worksheet>"#
    )
    .into_bytes()
}

fn target_path(index: usize, fault: PackageFault) -> String {
    if index == 0
        && (fault == PackageFault::MissingTargetContentType || invalid_payload_fault(fault))
    {
        let stem = if fault == PackageFault::MissingTargetContentType {
            "item-missing"
        } else {
            "item0"
        };
        format!("/xl/customXml/{stem}.custom")
    } else {
        format!("{TARGET_PREFIX}{index}.xml")
    }
}

fn target_content_type(index: usize, fault: PackageFault) -> String {
    if index == 0 && invalid_payload_fault(fault) {
        "application/octet-stream".to_owned()
    } else if fault == PackageFault::WrongTargetContentType && index == 0 {
        "application/octet-stream".to_owned()
    } else if index == 1 {
        "application/litchi+xml".to_owned()
    } else {
        "text/xml".to_owned()
    }
}

fn invalid_payload_fault(fault: PackageFault) -> bool {
    matches!(
        fault,
        PackageFault::MalformedTarget
            | PackageFault::InvalidQNameTarget
            | PackageFault::UnknownEntityTarget
            | PackageFault::RawAttributeTarget
            | PackageFault::RawCdataCloseTarget
            | PackageFault::ControlCharacterTarget
            | PackageFault::BadXmlDeclarationTarget
            | PackageFault::BadXmlEncodingTarget
            | PackageFault::BadXmlVersionTarget
            | PackageFault::TextBeforeRootTarget
            | PackageFault::TextAfterRootTarget
            | PackageFault::CdataOutsideRootTarget
            | PackageFault::UnboundNamespaceTarget
            | PackageFault::EmptyNamespaceTarget
            | PackageFault::ReservedXmlNamespaceTarget
            | PackageFault::ReservedXmlnsNamespaceTarget
            | PackageFault::EmptyTarget
    )
}

fn target_payload(index: usize, fault: PackageFault) -> Vec<u8> {
    if index == 0 {
        let malformed = match fault {
            PackageFault::MalformedTarget => {
                Some(br#"<payload xmlns="urn:litchi:content"><value>unterminated"#.to_vec())
            },
            PackageFault::InvalidQNameTarget => Some(br#"<1bad/>"#.to_vec()),
            PackageFault::UnknownEntityTarget => Some(br#"<payload>&foo;</payload>"#.to_vec()),
            PackageFault::RawAttributeTarget => Some(br#"<payload attr="<inattr"/>"#.to_vec()),
            PackageFault::RawCdataCloseTarget => Some(br#"<payload/>]]>text"#.to_vec()),
            PackageFault::ControlCharacterTarget => Some(b"<payload>\x01</payload>".to_vec()),
            PackageFault::BadXmlDeclarationTarget => {
                Some(br#"<?xml version="1.0" standalone="maybe"?><payload/>"#.to_vec())
            },
            PackageFault::BadXmlEncodingTarget => {
                Some(br#"<?xml encoding="definitely-not-an-encoding"?><payload/>"#.to_vec())
            },
            PackageFault::BadXmlVersionTarget => {
                Some(br#"<?xml version="2.0"?><payload/>"#.to_vec())
            },
            PackageFault::TextBeforeRootTarget => Some(b"text<payload/>".to_vec()),
            PackageFault::TextAfterRootTarget => Some(b"<payload/>text".to_vec()),
            PackageFault::CdataOutsideRootTarget => Some(b"<![CDATA[before]]><payload/>".to_vec()),
            PackageFault::UnboundNamespaceTarget => Some(br#"<p:payload/>"#.to_vec()),
            PackageFault::EmptyNamespaceTarget => Some(br#"<p:payload xmlns:p=""/>"#.to_vec()),
            PackageFault::ReservedXmlNamespaceTarget => {
                Some(br#"<payload xmlns:xml="urn:litchi:wrong-xml"/>"#.to_vec())
            },
            PackageFault::ReservedXmlnsNamespaceTarget => {
                Some(br#"<payload xmlns:xmlns="urn:litchi:wrong-xmlns"/>"#.to_vec())
            },
            PackageFault::EmptyTarget => Some(Vec::new()),
            PackageFault::EmptyRootTarget => Some(br#"<payload/>"#.to_vec()),
            PackageFault::NestedEmptyTarget => {
                Some(
                    br#"<payload><empty><empty><empty><empty/></empty></empty></empty></payload>"#
                        .to_vec(),
                )
            },
            PackageFault::NestedDepthTarget => Some(
                br#"<payload><a><b><c><d/></c></b></a></payload>"#.to_vec(),
            ),
            PackageFault::LegalOpaqueTarget => Some(
                br#"<?xml version="1.0"?><payload><!--comment--><?opaque data?><![CDATA[opaque & <]]></payload>"#.to_vec(),
            ),
            _ => None,
        };
        if let Some(malformed) = malformed {
            return malformed;
        }
    }
    format!(
        r#"<payload xmlns="urn:litchi:content" index="{index}"><value>content-{index}</value></payload>"#
    )
    .into_bytes()
}

fn content_part_package(strict: bool, fault: PackageFault, owner_count: usize) -> Vec<u8> {
    let mut package = Package::create().unwrap().into_plain_opc();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(worksheet_xml(strict));
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(DRAWING).unwrap(),
            ct::OFC_DRAWING.to_owned(),
            direct_core_drawing_count(strict, owner_count),
        )))
        .unwrap();

    for index in 0..owner_count {
        let path = target_path(index, fault);
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(path).unwrap(),
                target_content_type(index, fault),
                target_payload(index, fault),
            )))
            .unwrap();
    }
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(UNRELATED).unwrap(),
            "application/octet-stream".to_owned(),
            b"unrelated-package-member".to_vec(),
        )))
        .unwrap();
    if fault == PackageFault::OutboundGraph {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/customXml/outbound.xml").unwrap(),
                "text/xml".to_owned(),
                br#"<outbound xmlns="urn:litchi:outbound"/>"#.to_vec(),
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
    for index in 0..owner_count {
        if fault == PackageFault::MissingRelation && index == 0 {
            continue;
        }
        let target = if fault == PackageFault::ExternalRelation && index == 0 {
            "https://example.invalid/content.xml".to_owned()
        } else if fault == PackageFault::DanglingTarget && index == 0 {
            "../customXml/does-not-exist.xml".to_owned()
        } else if fault == PackageFault::SharedTarget && index == 1 {
            "../customXml/item0.xml".to_owned()
        } else {
            let path = target_path(index, fault);
            path.trim_start_matches("/xl/")
                .to_owned()
                .replace("customXml/", "../customXml/")
        };
        let relation_type = if fault == PackageFault::WrongRelationType && index == 0 {
            "urn:litchi:wrong-content-part".to_owned()
        } else if strict {
            STRICT_CUSTOM_XML.to_owned()
        } else {
            rt::CUSTOM_XML.to_owned()
        };
        drawing
            .rels_mut()
            .try_add_relationship(
                relation_type,
                target,
                format!("rIdContent{index}"),
                if fault == PackageFault::ExternalRelation && index == 0 {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )
            .unwrap();
    }
    if fault == PackageFault::OutboundGraph {
        package
            .get_part_mut(&PackURI::new("/xl/customXml/item0.xml").unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::CUSTOM_XML.to_owned(),
                "outbound.xml".to_owned(),
                "rIdOutbound".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }

    let bytes = PackageWriter::to_bytes(&package).unwrap();
    if fault == PackageFault::MissingTargetContentType {
        rewrite_content_type_override(&bytes, &target_path(0, fault), None)
    } else if invalid_payload_fault(fault) {
        rewrite_content_type_override(&bytes, &target_path(0, fault), Some("text/xml"))
    } else {
        bytes
    }
}

fn rewrite_content_type_override(
    bytes: &[u8],
    part_name: &str,
    content_type: Option<&str>,
) -> Vec<u8> {
    rewrite_physical_member(bytes, "[Content_Types].xml", |content_types| {
        let mut content_types = String::from_utf8(content_types).unwrap();
        let needle = format!(r#"PartName="{part_name}""#);
        let marker = content_types.find(&needle).unwrap();
        let start = content_types[..marker].rfind("<Override").unwrap();
        let end = content_types[marker..].find("/>").unwrap() + marker + 2;
        if let Some(content_type) = content_type {
            let override_xml = content_types[start..end].replace(
                r#"ContentType="application/octet-stream""#,
                &format!(r#"ContentType="{content_type}""#),
            );
            content_types.replace_range(start..end, &override_xml);
        } else {
            content_types.replace_range(start..end, "");
        }
        content_types.into_bytes()
    })
}

fn rewrite_physical_member(
    bytes: &[u8],
    wanted: &str,
    rewrite: impl FnOnce(Vec<u8>) -> Vec<u8>,
) -> Vec<u8> {
    let reader = PhysPkgReader::new(bytes).unwrap();
    let mut writer = PhysPkgWriter::new();
    let mut rewrite = Some(rewrite);
    for member in reader.member_names().unwrap() {
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

#[test]
fn source_core_content_parts_pair_all_anchor_geometry_and_ranges() {
    for (strict, dialect, relationship_dialect) in [
        (
            false,
            DrawingDialect::Transitional,
            RelationshipDialect::Transitional,
        ),
        (true, DrawingDialect::Strict, RelationshipDialect::Strict),
    ] {
        let xml = direct_core_drawing(strict);
        let source = SourceDrawing::scan(&xml).expect("core content-part source must scan");
        assert_eq!(source.dialect(), dialect);
        assert_eq!(source.relationship_dialect(), relationship_dialect);
        assert_eq!(source.content_parts().len(), 3);

        for (index, owner) in source.content_parts().iter().enumerate() {
            assert_eq!(owner.content_part_ordinal(), index);
            assert_eq!(owner.profile(), ContentPartProfile::CoreAnchor);
            assert_eq!(owner.relationship_id(), format!("rIdContent{index}"));
            assert_eq!(owner.relationship_dialect(), relationship_dialect);
            assert!(!owner.has_mce_ancestor());
            assert!(std::ptr::eq(owner.source().as_ptr(), xml.as_ptr()));
            assert_eq!(owner.source().len(), xml.len());

            let owner_bytes = owner
                .owner_bytes(&xml)
                .expect("owner range must borrow from original source");
            assert!(owner_bytes.starts_with(b"<xdr:contentPart"));
            assert!(owner_bytes.ends_with(b"/>"));
            assert_eq!(
                owner_bytes,
                &xml[owner.owner_range().start..owner.owner_range().end]
            );

            let anchor_bytes = owner
                .anchor_bytes(&xml)
                .expect("anchor range must borrow from original source");
            assert!(anchor_bytes.contains(&b'x'));
            assert_eq!(
                anchor_bytes,
                &xml[owner.anchor_range().start..owner.anchor_range().end]
            );

            match (index, owner.anchor()) {
                (0, DrawingAnchor::TwoCell { from, to, edit_as }) => {
                    assert_eq!(from.column, 1);
                    assert_eq!(from.column_offset.emu(), 2);
                    assert_eq!(from.row, 3);
                    assert_eq!(from.row_offset.emu(), 4);
                    assert_eq!(to.column, 5);
                    assert_eq!(to.column_offset.emu(), 6);
                    assert_eq!(to.row, 7);
                    assert_eq!(to.row_offset.emu(), 8);
                    assert_eq!(*edit_as, EditAs::OneCell);
                },
                (1, DrawingAnchor::OneCell { from, extent }) => {
                    assert_eq!(from.column, 9);
                    assert_eq!(from.column_offset.emu(), 19);
                    assert_eq!(from.row, 10);
                    assert_eq!(from.row_offset.emu(), 20);
                    assert_eq!(extent.width.emu(), 123456);
                    assert_eq!(extent.height.emu(), 654321);
                },
                (2, DrawingAnchor::Absolute { position, extent }) => {
                    assert_eq!(position.x.emu(), -900);
                    assert_eq!(position.y.emu(), 456);
                    assert_eq!(extent.width.emu(), 777888);
                    assert_eq!(extent.height.emu(), 999000);
                },
                _ => panic!("content-part geometry did not pair with source order {index}"),
            }
        }
    }
}

#[test]
fn alias_default_namespace_and_mixed_picture_content_part_order_stay_paired() {
    let anchors = format!(
        "{}{}{}{}",
        core_picture_anchor("x", "rel", "a", 0),
        core_anchor("x", "rel", 0),
        core_picture_anchor("x", "rel", "a", 1),
        core_anchor("x", "rel", 1),
    );
    let xml =
        format!(r#"<x:wsDr xmlns:x="{XDR}" xmlns:rel="{REL}" xmlns:a="{A}">{anchors}</x:wsDr>"#)
            .replace("x:", "")
            .replace("xmlns:x=", "xmlns=");

    let source = SourceDrawing::scan(xml.as_bytes()).expect("default namespace drawing scans");
    assert_eq!(source.pictures().len(), 2);
    assert_eq!(source.content_parts().len(), 2);

    let typed = litchi_xlsx::drawing::parse(&xml).unwrap().unwrap();
    assert_eq!(typed.objects().len(), 4);
    assert!(matches!(&typed.objects()[0], Object::Picture(_)));
    assert!(matches!(
        &typed.objects()[1],
        Object::Unknown(value) if value.kind == UnknownKind::ContentPart
    ));
    assert!(matches!(&typed.objects()[2], Object::Picture(_)));
    assert!(matches!(
        &typed.objects()[3],
        Object::Unknown(value) if value.kind == UnknownKind::ContentPart
    ));

    for picture_ordinal in 0..2 {
        source
            .match_typed_picture(&typed, picture_ordinal)
            .expect("mixed picture source pairs with typed picture");
    }
    for (content_part_ordinal, object_index) in [(0, 1), (1, 3)] {
        let typed_anchor = match &typed.objects()[object_index] {
            Object::Unknown(value) => value.drawing_anchor(),
            _ => panic!("mixed content part did not remain an unknown object"),
        };
        assert_eq!(
            source.content_part(content_part_ordinal).unwrap().anchor(),
            typed_anchor
        );
        assert_eq!(
            source
                .content_part(content_part_ordinal)
                .unwrap()
                .relationship_id(),
            format!("rIdContent{content_part_ordinal}")
        );
    }
}

#[test]
fn source_content_part_requires_childless_core_element_and_canonical_qname() {
    let valid = String::from_utf8(direct_core_drawing(false)).unwrap();
    let malformed = [
        valid.replace("/>", "><xdr:future/></xdr:contentPart>"),
        valid.replace("rIdContent0", ""),
        valid.replace("xdr:contentPart", "contentPart"),
        valid.replace(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
            "xmlns:r=\"urn:litchi:foreign-rel\"",
        ),
        valid.replace(
            "xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"",
            "xmlns:xdr=\"urn:litchi:foreign-xdr\"",
        ),
        valid.replace(
            "<xdr:contentPart r:id=\"rIdContent0\"/>",
            "<xdr:contentPart r:id=\"rIdContent0\">text</xdr:contentPart>",
        ),
        valid
            .replace("r:id=\"rIdContent0\"", "r:id=\"rIdContent0\"><xdr:future/>")
            .replace("<xdr:clientData/>", "</xdr:contentPart><xdr:clientData/>"),
    ];
    for candidate in malformed {
        assert!(
            SourceDrawing::scan(candidate.as_bytes()).is_err(),
            "malformed core contentPart was accepted: {candidate}"
        );
    }
}

#[test]
fn strict_core_content_part_rejects_unknown_attributes_and_pairs_valid_typed_owners() {
    let valid = String::from_utf8(direct_core_drawing(true)).unwrap();
    let unknown_unqualified = valid.replace(
        "<xdr:contentPart r:id=\"rIdContent0\"/>",
        "<xdr:contentPart r:id=\"rIdContent0\" unexpected=\"yes\"/>",
    );
    let foreign_namespace = valid
        .replace(
            &format!(r#"xmlns:xdr="{STRICT_XDR}""#),
            &format!(r#"xmlns:xdr="{STRICT_XDR}" xmlns:foreign="urn:litchi:foreign""#),
        )
        .replace(
            "<xdr:contentPart r:id=\"rIdContent0\"/>",
            "<xdr:contentPart r:id=\"rIdContent0\" foreign:attr=\"yes\"/>",
        );
    for candidate in [unknown_unqualified, foreign_namespace] {
        assert!(
            SourceDrawing::scan(candidate.as_bytes()).is_err(),
            "strict contentPart accepted an unknown attribute: {candidate}"
        );
        assert!(
            litchi_xlsx::drawing::parse(&candidate).is_err(),
            "typed strict parser accepted an unknown attribute: {candidate}"
        );
    }

    let source = SourceDrawing::scan(valid.as_bytes()).unwrap();
    assert_eq!(source.dialect(), DrawingDialect::Strict);
    assert_eq!(source.relationship_dialect(), RelationshipDialect::Strict);
    let typed = litchi_xlsx::drawing::parse(&valid).unwrap().unwrap();
    assert_eq!(typed.objects().len(), source.content_parts().len());
    for (index, owner) in source.content_parts().iter().enumerate() {
        let typed_anchor = match &typed.objects()[index] {
            Object::Unknown(value) if value.kind == UnknownKind::ContentPart => {
                value.drawing_anchor()
            },
            _ => panic!("strict typed object {index} is not a content part"),
        };
        assert_eq!(owner.relationship_id(), format!("rIdContent{index}"));
        assert_eq!(owner.anchor(), typed_anchor);
    }
}

#[test]
fn source_content_part_refuses_mce_owners_and_keeps_foreign_opaque_markup_out() {
    let active_choice = format!(
        r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:r="{REL}" xmlns:mc="{MCE}" xmlns:x="urn:litchi:choice"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent><mc:Choice Requires="x">{owner}</mc:Choice><mc:Fallback><xdr:sp/></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#,
        owner = core_content_part("xdr", "r", "rIdMce"),
    );
    assert!(SourceDrawing::scan(active_choice.as_bytes()).is_err());

    let foreign_wrapper = format!(
        r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:r="{REL}" xmlns:pmc="{FOREIGN_MCE}" xmlns:opaque="urn:litchi:opaque">{}<pmc:AlternateContent pmc:Ignorable="opaque"><pmc:Choice Requires="opaque"><xdr:contentPart r:id="rIdOpaque"/></pmc:Choice><pmc:Fallback><opaque:payload/></pmc:Fallback></pmc:AlternateContent></xdr:wsDr>"#,
        core_anchor("xdr", "r", 0),
    );
    let source =
        SourceDrawing::scan(foreign_wrapper.as_bytes()).expect("foreign wrapper stays opaque");
    assert_eq!(source.content_parts().len(), 1);
    assert_eq!(source.content_parts()[0].relationship_id(), "rIdContent0");
    assert!(
        foreign_wrapper
            .as_bytes()
            .windows(b"pmc:AlternateContent".len())
            .any(|window| window == b"pmc:AlternateContent")
    );

    let duplicate = String::from_utf8(direct_core_drawing(false))
        .unwrap()
        .replace(
            "<xdr:clientData/>",
            "<xdr:contentPart r:id=\"rIdDuplicate\"/><xdr:clientData/>",
        );
    assert!(
        SourceDrawing::scan(duplicate.as_bytes()).is_err(),
        "two direct core owners in one anchor must be refused"
    );
}

#[test]
fn source_content_part_limit_accepts_exact_and_rejects_one_under() {
    let xml = direct_core_drawing(false);
    let exact = litchi_xlsx::drawing::ScanLimits {
        max_content_parts: 3,
        ..litchi_xlsx::drawing::ScanLimits::default()
    };
    assert_eq!(
        SourceDrawing::scan_with_limits(&xml, 0, exact)
            .unwrap()
            .content_parts()
            .len(),
        3
    );
    let one_under = litchi_xlsx::drawing::ScanLimits {
        max_content_parts: 2,
        ..litchi_xlsx::drawing::ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(&xml, 0, one_under).is_err());
}

#[test]
fn nested_xdr14_group_content_part_is_not_inferred_as_core_owner() {
    let xml = format!(
        r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:r="{REL}" xmlns:xdr14="{XDR14}"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:nvGrpSpPr/><xdr:grpSpPr/><xdr14:contentPart r:id="rIdGroup"/></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#
    );
    let source = SourceDrawing::scan(xml.as_bytes()).expect("unresolved extension remains opaque");
    assert!(source.content_parts().is_empty());
}

#[test]
fn worksheet_content_part_facade_pairs_source_typed_geometry_and_borrowed_payload() {
    for strict in [false, true] {
        let bytes = content_part_package(strict, PackageFault::None, 3);
        let before = OpcPackage::from_bytes(&bytes).unwrap();
        let unrelated_before = before
            .get_part(&PackURI::new(UNRELATED).unwrap())
            .unwrap()
            .blob()
            .to_vec();

        let workbook = Workbook::from_slice(&bytes).expect("content-part package must open");
        let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
        let drawing = worksheet
            .drawing(0)
            .expect("worksheet drawing must resolve");
        assert_eq!(drawing.content_part_count(), 3);
        assert_eq!(drawing.source().content_parts().len(), 3);
        let batch = drawing
            .content_parts()
            .expect("batch content-part resolution must succeed");
        assert_eq!(batch.len(), 3);

        let typed = drawing
            .typed()
            .objects()
            .iter()
            .filter_map(|object| match object {
                Object::Unknown(value) if value.kind == UnknownKind::ContentPart => Some(value),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(typed.len(), 3);

        for index in 0..3 {
            let part = drawing
                .content_part(index)
                .expect("core content-part relationship must resolve");
            let batched = &batch[index];
            let source = drawing.source().content_part(index).unwrap();
            assert_eq!(part.profile(), ContentPartProfile::CoreAnchor);
            assert_eq!(
                batched.relationship().target_uri(),
                part.relationship().target_uri()
            );
            assert_eq!(part.source().content_part_ordinal(), index);
            assert_eq!(
                part.source().relationship_id(),
                format!("rIdContent{index}")
            );
            assert_eq!(part.anchor(), source.anchor());
            assert_eq!(typed[index].drawing_anchor(), part.anchor());
            assert_eq!(
                part.owner_bytes().unwrap(),
                source.owner_bytes(drawing.source_xml()).unwrap()
            );

            let relationship = part.relationship();
            assert_eq!(relationship.relationship_id(), format!("rIdContent{index}"));
            assert_eq!(
                relationship.relationship_type(),
                if strict {
                    STRICT_CUSTOM_XML
                } else {
                    rt::CUSTOM_XML
                }
            );
            assert_eq!(
                relationship.target_uri().as_str(),
                format!("{TARGET_PREFIX}{index}.xml")
            );
            assert_eq!(
                relationship.content_type(),
                if index == 1 {
                    "application/litchi+xml"
                } else {
                    "text/xml"
                }
            );

            let payload = part.payload().expect("target XML must be read lazily");
            assert_eq!(
                payload.bytes(),
                target_payload(index, PackageFault::None).as_slice()
            );
            assert!(payload.outbound_relationships().is_empty());
            let first_payload_ptr = payload.bytes().as_ptr();
            let second_payload = part.payload().unwrap();
            assert_eq!(first_payload_ptr, second_payload.bytes().as_ptr());
        }

        let after_bytes = workbook.to_plain_bytes().unwrap();
        let after = OpcPackage::from_bytes(&after_bytes).unwrap();
        assert_eq!(
            after
                .get_part(&PackURI::new(UNRELATED).unwrap())
                .unwrap()
                .blob(),
            unrelated_before.as_slice()
        );
    }
}

#[test]
fn worksheet_content_part_relation_and_target_graph_failures_are_typed() {
    for fault in [
        PackageFault::MissingRelation,
        PackageFault::WrongRelationType,
        PackageFault::ExternalRelation,
        PackageFault::DanglingTarget,
        PackageFault::WrongTargetContentType,
    ] {
        for strict in [false, true] {
            let bytes = content_part_package(strict, fault, 3);
            let workbook = Workbook::from_slice(&bytes).expect("graph fixture package must open");
            let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
            let drawing = worksheet
                .drawing(0)
                .expect("drawing source must remain readable");
            assert!(
                drawing.content_part(0).is_err(),
                "fault {fault:?} was accepted for {} dialect",
                if strict { "strict" } else { "transitional" }
            );
        }
    }

    let missing_content_type =
        content_part_package(false, PackageFault::MissingTargetContentType, 1);
    assert!(
        Workbook::from_slice(&missing_content_type).is_err(),
        "an XML target without a content-type mapping must fail package ingress"
    );

    let malformed_target = content_part_package(false, PackageFault::MalformedTarget, 1);
    let workbook = Workbook::from_slice(&malformed_target).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    assert!(drawing.content_part(0).unwrap().payload().is_err());
}

#[test]
fn worksheet_content_part_payload_xml_contract_is_bounded_and_preserving() {
    for fault in [
        PackageFault::MalformedTarget,
        PackageFault::InvalidQNameTarget,
        PackageFault::UnknownEntityTarget,
        PackageFault::RawAttributeTarget,
        PackageFault::RawCdataCloseTarget,
        PackageFault::ControlCharacterTarget,
        PackageFault::BadXmlDeclarationTarget,
        PackageFault::BadXmlEncodingTarget,
        PackageFault::BadXmlVersionTarget,
        PackageFault::TextBeforeRootTarget,
        PackageFault::TextAfterRootTarget,
        PackageFault::CdataOutsideRootTarget,
        PackageFault::UnboundNamespaceTarget,
        PackageFault::EmptyNamespaceTarget,
        PackageFault::ReservedXmlNamespaceTarget,
        PackageFault::ReservedXmlnsNamespaceTarget,
        PackageFault::EmptyTarget,
    ] {
        let bytes = content_part_package(false, fault, 1);
        let workbook = Workbook::from_slice(&bytes).expect("payload fixture package must open");
        let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
        let drawing = worksheet.drawing(0).unwrap();
        assert!(
            drawing.content_part(0).unwrap().payload().is_err(),
            "payload fault {fault:?} was accepted"
        );
    }

    let legal = content_part_package(false, PackageFault::LegalOpaqueTarget, 1);
    let workbook = Workbook::from_slice(&legal).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    let payload = drawing.content_part(0).unwrap().payload().unwrap();
    assert!(
        payload
            .bytes()
            .windows(b"<!--comment-->".len())
            .any(|window| { window == b"<!--comment-->" })
    );
    assert!(
        payload
            .bytes()
            .windows(b"<?opaque data?>".len())
            .any(|window| { window == b"<?opaque data?>" })
    );
    assert!(
        payload
            .bytes()
            .windows(b"<![CDATA[opaque & <]]>".len())
            .any(|window| window == b"<![CDATA[opaque & <]]>")
    );

    let empty_root = content_part_package(false, PackageFault::EmptyRootTarget, 1);
    let root_limit = ReadLimits::builder()
        // The worksheet/drawing source itself reaches depth four; the empty
        // payload root must still count as an XML element at that boundary.
        .max_xml_depth(4)
        .unwrap()
        .build()
        .unwrap();
    let workbook = Workbook::from_slice_with_limits(&empty_root, root_limit).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    assert_eq!(
        drawing.content_part(0).unwrap().payload().unwrap().bytes(),
        b"<payload/>"
    );

    let nested_empty = content_part_package(false, PackageFault::NestedEmptyTarget, 1);
    let one_under = ReadLimits::builder()
        // The payload's fifth empty-element depth is one over the shared cap.
        .max_xml_depth(4)
        .unwrap()
        .build()
        .unwrap();
    let workbook = Workbook::from_slice_with_limits(&nested_empty, one_under).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    assert!(drawing.content_part(0).unwrap().payload().is_err());

    let nested = content_part_package(false, PackageFault::NestedDepthTarget, 1);
    let limits = ReadLimits::builder()
        .max_xml_depth(4)
        .unwrap()
        .build()
        .unwrap();
    let workbook = Workbook::from_slice_with_limits(&nested, limits).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    assert!(drawing.content_part(0).unwrap().payload().is_err());
}

#[test]
fn worksheet_content_part_shared_target_and_outbound_graph_are_retained_and_bounded() {
    let shared = content_part_package(false, PackageFault::SharedTarget, 3);
    let workbook = Workbook::from_slice(&shared).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    let batch = drawing.content_parts().unwrap();
    assert_eq!(batch.len(), 3);
    assert_eq!(
        batch[0].relationship().target_uri(),
        batch[1].relationship().target_uri()
    );
    let first = drawing.content_part(0).unwrap();
    let second = drawing.content_part(1).unwrap();
    assert_eq!(
        first.relationship().target_uri(),
        second.relationship().target_uri()
    );
    assert_eq!(
        first.payload().unwrap().bytes(),
        second.payload().unwrap().bytes()
    );

    let outbound = content_part_package(false, PackageFault::OutboundGraph, 1);
    let workbook = Workbook::from_slice(&outbound).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = worksheet.drawing(0).unwrap();
    let payload = drawing.content_part(0).unwrap().payload().unwrap();
    assert_eq!(payload.outbound_relationships().len(), 1);
    assert_eq!(
        payload.outbound_relationships()[0].relationship_id(),
        "rIdOutbound"
    );
    assert_eq!(
        payload.outbound_relationships()[0].relationship_type(),
        rt::CUSTOM_XML
    );
    assert_eq!(
        payload.outbound_relationships()[0].target_ref(),
        "outbound.xml"
    );
    assert_eq!(
        payload.outbound_relationships()[0].target_mode(),
        TargetMode::Internal
    );
}

#[test]
fn worksheet_content_part_limit_accepts_exact_and_rejects_one_under() {
    // The worksheet owner maps max_relationships_per_part into the drawing
    // scanner's max_content_parts ceiling as well as its relationship budget.
    let bytes = content_part_package(false, PackageFault::None, 3);
    let exact = ReadLimits::builder()
        .max_relationships_per_part(3)
        .unwrap()
        .build()
        .unwrap();
    let workbook = Workbook::from_slice_with_limits(&bytes, exact).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    assert_eq!(worksheet.drawing(0).unwrap().content_part_count(), 3);

    let one_under = ReadLimits::builder()
        .max_relationships_per_part(2)
        .unwrap()
        .build()
        .unwrap();
    // Keep the physical drawing relationship member within the caller's
    // relationship ceiling so the failure is charged to the source owner
    // cap, rather than to OPC ingress itself.
    let bytes_under = content_part_package(false, PackageFault::MissingRelation, 3);
    let workbook = Workbook::from_slice_with_limits(&bytes_under, one_under).unwrap();
    let worksheet = workbook.sheet("Sheet1").unwrap().unwrap();
    assert!(worksheet.drawing(0).is_err());
}
