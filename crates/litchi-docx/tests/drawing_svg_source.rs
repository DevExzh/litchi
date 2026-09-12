use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::Path;

use litchi_docx::Package;
use litchi_docx::drawing::{
    DrawingDialect, DrawingPlacement, RelationshipDialect, ScanLimits, SourceDrawing, SvgOwnerState,
};
use litchi_opc::PackURI;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ASVG: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_WP: &str = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_PIC: &str = "http://purl.oclc.org/ooxml/drawingml/picture";
const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_MCE: &str = "http://purl.oclc.org/ooxml/markup-compatibility/2006";
const PICTURE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const STRICT_PICTURE_URI: &str = "http://purl.oclc.org/ooxml/drawingml/picture";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

const INLINE_FIXTURE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/svg.docx"
);
const FLOATING_FIXTURE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf164835_nonDummyLineHeight.docx"
);

#[derive(Debug, Eq, PartialEq)]
struct RelationshipSummary {
    id: String,
    target: String,
    external: bool,
}

fn main_story(path: &str) -> (Vec<u8>, Vec<RelationshipSummary>) {
    assert!(
        Path::new(path).is_file(),
        "native fixture is missing: {path}"
    );
    let package = Package::from_reader(Cursor::new(
        std::fs::read(path).expect("read native DOCX fixture"),
    ))
    .expect("open native DOCX fixture through the OPC package API");
    let document = package
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").expect("document part URI"))
        .expect("main document part")
        .blob()
        .to_vec();
    let part = package
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").expect("document part URI"))
        .expect("main document part");
    let mut relationships = part
        .rels()
        .iter()
        .map(|relationship| RelationshipSummary {
            id: relationship.r_id().to_owned(),
            target: relationship
                .target_partname()
                .expect("native image relationship is internal")
                .as_str()
                .to_owned(),
            external: relationship.is_external(),
        })
        .collect::<Vec<_>>();
    relationships.sort_by(|left, right| left.id.cmp(&right.id));
    (document, relationships)
}

fn assert_source_ranges(source: &SourceDrawing<'_>) {
    let bytes = source.source();
    for picture in source.pictures() {
        let drawing = picture.drawing_range();
        let anchor = picture.anchor_range();
        let graphic = picture.graphic_range();
        let graphic_data = picture.graphic_data_range();
        let picture_range = picture.picture_range();
        for range in [drawing, anchor, graphic, graphic_data, picture_range] {
            assert!(range.start < range.end, "source range must be non-empty");
            assert!(range.end <= bytes.len(), "source range must stay in source");
            assert_eq!(range.slice(bytes).unwrap(), &bytes[range.start..range.end]);
        }
        assert!(drawing.start <= anchor.start);
        assert!(anchor.start <= graphic.start);
        assert!(graphic.start <= graphic_data.start);
        assert!(graphic_data.start <= picture_range.start);
        assert!(picture_range.end <= graphic_data.end);
        assert!(graphic_data.end <= graphic.end);
        assert!(graphic.end <= anchor.end);
        assert!(anchor.end <= drawing.end);
        assert!(std::ptr::eq(picture.source().as_ptr(), bytes.as_ptr()));
        assert_eq!(picture.source().len(), bytes.len());

        if let Some(blip) = picture.blip_range() {
            assert_eq!(
                blip.bytes().unwrap(),
                picture.source_bytes(blip.range()).unwrap()
            );
            assert!(picture_range.start <= blip.range().start);
            assert!(blip.range().end <= picture_range.end);
        }
        if let Some(ext_list) = picture.ext_list_range() {
            assert_eq!(
                ext_list.bytes().unwrap(),
                picture.source_bytes(ext_list.range()).unwrap()
            );
            assert!(picture_range.start <= ext_list.range().start);
            assert!(ext_list.range().end <= picture_range.end);
        }
        if let SvgOwnerState::Embedded(owner) | SvgOwnerState::Linked(owner) = picture.svg_owner() {
            assert!(std::ptr::eq(owner.source().as_ptr(), bytes.as_ptr()));
            assert_eq!(owner.source().len(), bytes.len());
            assert_eq!(
                owner.extension_element().bytes().unwrap(),
                picture
                    .source_bytes(owner.extension_range())
                    .expect("owner extension range")
            );
            assert_eq!(
                owner.svg_blip_element().bytes().unwrap(),
                picture
                    .source_bytes(owner.svg_blip_element().range())
                    .expect("SVG owner range")
            );
            assert!(picture_range.start <= owner.extension_range().start);
            assert!(owner.extension_range().end <= picture_range.end);
            assert!(owner.extension_range().start <= owner.svg_blip_element().range().start);
            assert!(owner.svg_blip_element().range().end <= owner.extension_range().end);
        }
        for reference in picture.relationship_references() {
            assert!(std::ptr::eq(reference.source().as_ptr(), bytes.as_ptr()));
            assert_eq!(reference.source().len(), bytes.len());
            assert!(picture_range.start <= reference.range().start);
            assert!(reference.range().end <= picture_range.end);
            assert!(!reference.id().is_empty());
        }
    }
    for reference in source.relationship_references() {
        assert!(std::ptr::eq(reference.source().as_ptr(), bytes.as_ptr()));
        assert_eq!(reference.source().len(), bytes.len());
        assert!(reference.range().end <= bytes.len());
    }
}

fn assert_standalone_svg_fragment(fragment: &[u8], svg_prefix: &str, relationship_prefix: &str) {
    let root = format!("<{svg_prefix}:svgBlip");
    assert!(fragment.starts_with(root.as_bytes()));
    assert!(fragment.ends_with(b"/>"));
    let declaration = format!("xmlns:{relationship_prefix}=");
    assert!(
        fragment
            .windows(declaration.len())
            .any(|window| window == declaration.as_bytes()),
        "standalone SVG child must carry its relationship namespace"
    );
    let svg_declaration = format!("xmlns:{svg_prefix}=");
    assert!(
        fragment
            .windows(svg_declaration.len())
            .any(|window| window == svg_declaration.as_bytes())
    );

    let mut reader = Reader::from_reader(fragment);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut roots = 0;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .expect("standalone SVG XML")
        {
            Event::Empty(element) => {
                assert_eq!(element.name().local_name().as_ref(), b"svgBlip");
                roots += 1;
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Text(_) | Event::Comment(_) => {},
            event => panic!("unexpected standalone SVG event: {event:?}"),
        }
        buffer.clear();
    }
    assert_eq!(roots, 1);
}

#[test]
fn openxml_sdk_inline_fixture_has_borrowed_source_owner_and_graph() {
    let (document, relationships) = main_story(INLINE_FIXTURE);
    assert_eq!(document.len(), 3_850);
    assert_eq!(
        relationships,
        vec![
            RelationshipSummary {
                id: "rId1".into(),
                target: "/word/styles.xml".into(),
                external: false,
            },
            RelationshipSummary {
                id: "rId2".into(),
                target: "/word/settings.xml".into(),
                external: false,
            },
            RelationshipSummary {
                id: "rId3".into(),
                target: "/word/webSettings.xml".into(),
                external: false,
            },
            RelationshipSummary {
                id: "rId4".into(),
                target: "/word/media/image1.png".into(),
                external: false,
            },
            RelationshipSummary {
                id: "rId5".into(),
                target: "/word/media/image2.svg".into(),
                external: false,
            },
            RelationshipSummary {
                id: "rId6".into(),
                target: "/word/fontTable.xml".into(),
                external: false,
            },
            RelationshipSummary {
                id: "rId7".into(),
                target: "/word/theme/theme1.xml".into(),
                external: false,
            },
        ]
    );

    let source = SourceDrawing::scan_with_ordinal(&document, 12).expect("scan inline fixture");
    assert_eq!(source.drawing_ordinal(), 12);
    assert_eq!(source.dialect(), DrawingDialect::Transitional);
    assert_eq!(
        source.relationship_dialect(),
        RelationshipDialect::Transitional
    );
    assert_eq!(source.pictures().len(), 1);
    assert_eq!(source.relationship_references().len(), 2);
    assert_source_ranges(&source);

    let picture = source.picture(0).unwrap();
    assert_eq!(picture.placement(), DrawingPlacement::Inline);
    assert_eq!(picture.c_nv_pr_id(), Some("1"));
    assert_eq!(picture.raster_relationship_id(), Some("rId4"));
    assert_eq!(picture.relationship_references().len(), 2);
    assert!(picture.anchor_bytes().unwrap().starts_with(b"<wp:inline"));
    assert!(
        picture
            .anchor_bytes()
            .unwrap()
            .windows(b"cx=\"2895600\" cy=\"2762250\"".len())
            .any(|window| window == b"cx=\"2895600\" cy=\"2762250\"")
    );

    let owner = picture
        .svg_owner()
        .owner()
        .expect("native inline fixture SVG owner");
    assert_eq!(owner.embedded_relationship_id(), Some("rId5"));
    assert_eq!(owner.linked_relationship_id(), None);
    assert_eq!(
        owner.relationship_dialect(),
        RelationshipDialect::Transitional
    );
    assert_eq!(owner.uri_lexical(), SVG_URI.as_bytes());
    let standalone = owner.namespace_complete(16 * 1024).unwrap();
    assert_standalone_svg_fragment(&standalone, "asvg", "r");
}

#[test]
fn libreoffice_floating_fixture_preserves_anchor_geometry_and_source_provenance() {
    let (document, relationships) = main_story(FLOATING_FIXTURE);
    assert_eq!(document.len(), 5_599);
    assert_eq!(
        relationships
            .iter()
            .filter(|relationship| relationship.id == "rId4" || relationship.id == "rId5")
            .collect::<Vec<_>>(),
        vec![
            &RelationshipSummary {
                id: "rId4".into(),
                target: "/word/media/image1.png".into(),
                external: false,
            },
            &RelationshipSummary {
                id: "rId5".into(),
                target: "/word/media/image2.svg".into(),
                external: false,
            },
        ]
    );

    let source = SourceDrawing::scan(&document).expect("scan floating fixture");
    assert_eq!(source.dialect(), DrawingDialect::Transitional);
    assert_eq!(
        source.relationship_dialect(),
        RelationshipDialect::Transitional
    );
    assert_eq!(source.pictures().len(), 1);
    assert_source_ranges(&source);

    let picture = source.picture(0).unwrap();
    assert_eq!(picture.placement(), DrawingPlacement::Floating);
    assert_eq!(picture.c_nv_pr_id(), Some("197641587"));
    assert_eq!(picture.raster_relationship_id(), Some("rId4"));
    let anchor = picture.anchor_bytes().unwrap();
    assert!(anchor.starts_with(b"<wp:anchor"));
    for expected in [
        b"<wp:simplePos x=\"0\" y=\"0\"/>".as_slice(),
        b"<wp:positionH relativeFrom=\"margin\"><wp:align>left</wp:align></wp:positionH>"
            .as_slice(),
        b"<wp:wrapSquare wrapText=\"bothSides\"/>".as_slice(),
        b"cx=\"2524125\" cy=\"2524125\"".as_slice(),
        b"<wp14:sizeRelH relativeFrom=\"page\">".as_slice(),
        b"<wp14:sizeRelV relativeFrom=\"page\">".as_slice(),
    ] {
        assert!(
            anchor
                .windows(expected.len())
                .any(|window| window == expected)
        );
    }
    let owner = picture.svg_owner().owner().expect("floating SVG owner");
    assert_eq!(owner.embedded_relationship_id(), Some("rId5"));
    let standalone = owner.namespace_complete(16 * 1024).unwrap();
    assert_standalone_svg_fragment(&standalone, "asvg", "r");

    // The view and every range retain identity with the OPC-extracted member;
    // a same-content copy is a different source provenance token.
    assert!(std::ptr::eq(source.source().as_ptr(), document.as_ptr()));
    let copied = document.clone();
    let copied_source = SourceDrawing::scan(&copied).expect("scan copied source");
    assert!(std::ptr::eq(
        copied_source.source().as_ptr(),
        copied.as_ptr()
    ));
    assert!(!std::ptr::eq(
        source.source().as_ptr(),
        copied_source.source().as_ptr()
    ));
    assert_ne!(source, copied_source);
}

fn synthetic_picture(extension: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="{PICTURE_URI}"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="synthetic"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdRaster"><a:extLst>{extension}</a:extLst></a:blip></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

fn synthetic_document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R}" xmlns:asvg="{ASVG}" xmlns:mc="{MC}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn strict_document(body: &str) -> Vec<u8> {
    let mut document = String::from_utf8(synthetic_document(body)).expect("synthetic XML");
    for (transitional, strict) in [
        (W, STRICT_W),
        (WP, STRICT_WP),
        (A, STRICT_A),
        (PIC, STRICT_PIC),
        (R, STRICT_R),
        (PICTURE_URI, STRICT_PICTURE_URI),
    ] {
        document = document.replace(transitional, strict);
    }
    document.into_bytes()
}

#[test]
fn alias_unknown_duplicate_and_mce_owners_keep_direct_policy() {
    let aliased = synthetic_document(
        &synthetic_picture(
            &r#"<a:ext uri="{SVG_URI}"><svg:svgBlip xmlns:svg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:embed="rIdSvg"/></a:ext>"#
                .replace("{SVG_URI}", SVG_URI),
        ),
    );
    let aliased_source = SourceDrawing::scan(&aliased).expect("scan aliased SVG owner");
    let aliased_owner = aliased_source
        .picture(0)
        .unwrap()
        .svg_owner()
        .owner()
        .expect("aliased owner is admitted by expanded name");
    assert_eq!(aliased_owner.embedded_relationship_id(), Some("rIdSvg"));
    assert_eq!(
        aliased_owner.relationship_dialect(),
        RelationshipDialect::Transitional
    );
    let standalone = aliased_owner.namespace_complete(16 * 1024).unwrap();
    assert_standalone_svg_fragment(&standalone, "svg", "rel");
    assert!(
        standalone
            .windows(b"rel:embed=\"rIdSvg\"".len())
            .any(|window| { window == b"rel:embed=\"rIdSvg\"" })
    );

    let unknown = synthetic_document(&synthetic_picture(
        r#"<a:ext uri="urn:future"><future:svgBlip xmlns:future="urn:future" r:embed="rIdForeign"/></a:ext>"#,
    ));
    assert!(matches!(
        SourceDrawing::scan(&unknown)
            .unwrap()
            .picture(0)
            .unwrap()
            .svg_owner(),
        SvgOwnerState::Opaque
    ));

    let duplicate = synthetic_document(&synthetic_picture(&format!(
        r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdOne"/></a:ext><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdTwo"/></a:ext>"#
    )));
    assert!(matches!(
        SourceDrawing::scan(&duplicate)
            .unwrap()
            .picture(0)
            .unwrap()
            .svg_owner(),
        SvgOwnerState::Ambiguous
    ));

    let mce = synthetic_document(&format!(
        r#"<mc:AlternateContent><mc:Choice Requires="w14">{}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        synthetic_picture(&format!(
            r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdHidden"/></a:ext>"#
        )),
    ));
    assert!(matches!(
        SourceDrawing::scan(&mce)
            .unwrap()
            .picture(0)
            .unwrap()
            .svg_owner(),
        SvgOwnerState::Refused
    ));
}

#[test]
fn strict_host_keeps_canonical_mc_semantics_and_ignores_strict_mc_alias() {
    let extension =
        format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdStrictSvg"/></a:ext>"#);
    let picture = synthetic_picture(&extension);

    let strict_xml = strict_document(&picture);
    let strict = SourceDrawing::scan(&strict_xml).expect("strict SVG source");
    assert_eq!(strict.dialect(), DrawingDialect::Strict);
    assert!(matches!(
        strict.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Embedded(_)
    ));

    let canonical_mce = strict_document(&format!(
        r#"<mc:AlternateContent><mc:Choice Requires="w14">{picture}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#
    ));
    let canonical = SourceDrawing::scan(&canonical_mce).expect("canonical MC source");
    assert!(matches!(
        canonical.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));

    let foreign_mce = strict_document(&format!(
        r#"<strictMc:AlternateContent xmlns:strictMc="{STRICT_MCE}"><strictMc:Choice Requires="w14">{picture}</strictMc:Choice><strictMc:Fallback/></strictMc:AlternateContent>"#
    ));
    let foreign = SourceDrawing::scan(&foreign_mce).expect("foreign strict MC source");
    assert!(
        foreign.pictures().is_empty(),
        "the strict MC URI is foreign and must remain opaque/unselected"
    );
}

#[test]
fn source_scan_and_namespace_output_honor_explicit_bounds() {
    let (document, _) = main_story(INLINE_FIXTURE);
    let mut limits = ScanLimits {
        max_xml_bytes: document.len() - 1,
        ..ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(&document, 0, limits).is_err());

    limits = ScanLimits {
        max_global_relationship_references: 1,
        ..ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(&document, 0, limits).is_err());

    let source = SourceDrawing::scan(&document).unwrap();
    let owner = source.picture(0).unwrap().svg_owner().owner().unwrap();
    let complete = owner.namespace_complete(16 * 1024).unwrap();
    assert!(complete.len() > 1);
    assert!(owner.namespace_complete(complete.len() - 1).is_err());

    let one_picture = ScanLimits {
        max_pictures: 1,
        ..ScanLimits::default()
    };
    let bounded = SourceDrawing::scan_with_limits(&document, 0, one_picture).unwrap();
    assert_eq!(bounded.pictures().len(), 1);
}

#[test]
fn native_relationship_ids_match_the_global_source_references() {
    let (document, _) = main_story(INLINE_FIXTURE);
    let source = SourceDrawing::scan(&document).unwrap();
    let mut global = source
        .relationship_references()
        .iter()
        .map(|reference| (reference.local_name().to_owned(), reference.id().to_owned()))
        .collect::<Vec<_>>();
    global.sort();
    assert_eq!(
        global,
        vec![
            ("embed".to_owned(), "rId4".to_owned()),
            ("embed".to_owned(), "rId5".to_owned()),
        ]
    );

    let mut per_picture = source
        .picture(0)
        .unwrap()
        .relationship_references()
        .iter()
        .map(|reference| reference.id().to_owned())
        .collect::<Vec<_>>();
    per_picture.sort();
    assert_eq!(per_picture, vec!["rId4", "rId5"]);

    let mut by_id = BTreeMap::new();
    for reference in source.relationship_references() {
        by_id.insert(reference.id(), reference.dialect());
    }
    assert_eq!(by_id["rId4"], RelationshipDialect::Transitional);
    assert_eq!(by_id["rId5"], RelationshipDialect::Transitional);
}
