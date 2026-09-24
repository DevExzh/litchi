use super::source::{DrawingDialect, DrawingPlacement, ScanLimits, SourceDrawing, SvgOwnerState};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ASVG: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PICTURE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

fn document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R}" xmlns:asvg="{ASVG}" xmlns:mc="{MC}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn drawing(placement: &str, extension: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:{placement}><a:graphic><a:graphicData uri="{PICTURE_URI}"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdRaster"><a:extLst>{extension}</a:extLst></a:blip></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:{placement}></w:drawing></w:r>"#
    )
}

fn admitted(id: &str) -> String {
    format!(r#"<a:ext uri="  {SVG_URI}  "><asvg:svgBlip r:embed="{id}"/></a:ext>"#)
}

fn encoded_slashes(uri: &str) -> String {
    uri.replace('/', "&#x2F;")
}

#[test]
fn scans_inline_and_anchor_with_borrowed_source_and_ranges() {
    let xml = document(&format!(
        "{}{}",
        drawing("inline", &admitted("rIdSvgInline")),
        drawing("anchor", &admitted("rIdSvgAnchor")),
    ));
    let source = SourceDrawing::scan_with_ordinal(&xml, 4).expect("scan direct pictures");
    assert_eq!(source.drawing_ordinal(), 4);
    assert_eq!(source.dialect(), DrawingDialect::Transitional);
    assert_eq!(source.pictures().len(), 2);

    let inline = source.picture(0).expect("inline picture");
    assert_eq!(inline.placement(), DrawingPlacement::Inline);
    assert_eq!(inline.raster_relationship_id(), Some("rIdRaster"));
    assert_eq!(inline.c_nv_pr_id(), Some("7"));
    assert!(
        inline
            .c_nv_pr_range()
            .unwrap()
            .bytes()
            .unwrap()
            .ends_with(b"/>")
    );
    assert_eq!(inline.source(), xml.as_slice());
    assert!(std::ptr::eq(inline.source().as_ptr(), xml.as_ptr()));
    assert!(inline.picture_bytes().unwrap().starts_with(b"<pic:pic"));
    let owner = inline.svg_owner().owner().expect("embedded owner");
    assert_eq!(owner.embedded_relationship_id(), Some("rIdSvgInline"));
    assert_eq!(owner.uri_lexical(), format!("  {SVG_URI}  ").as_bytes());
    assert!(
        owner
            .value()
            .namespace_context()
            .is_some_and(|context| context.shares_storage(owner.namespace_context()))
    );
    assert_eq!(
        owner.value().raw_source(),
        Some(owner.svg_blip_element().bytes().unwrap())
    );
    let complete = owner.namespace_complete(16 * 1024).unwrap();
    assert!(complete.contains(&b'='));
    assert!(
        !complete
            .windows(b"xmlns:xmlns".len())
            .any(|w| w == b"xmlns:xmlns")
    );

    let floating = source.picture(1).expect("floating picture");
    assert_eq!(floating.placement(), DrawingPlacement::Floating);
    assert_eq!(
        floating
            .svg_owner()
            .owner()
            .unwrap()
            .embedded_relationship_id(),
        Some("rIdSvgAnchor")
    );
    assert!(floating.anchor_bytes().unwrap().starts_with(b"<wp:anchor"));
    assert!(floating.anchor_range().start < floating.picture_range().start);
}

#[test]
fn encoded_namespace_uris_resolve_across_owner_profiles_and_export_canonically() {
    let mut xml =
        String::from_utf8(document(&drawing("inline", &admitted("rIdSvgEncoded")))).unwrap();
    for uri in [W, WP, A, PIC, R, ASVG] {
        let encoded = encoded_slashes(uri);
        xml = xml.replace(uri, &encoded);
    }

    let source = SourceDrawing::scan(xml.as_bytes()).expect("encoded namespace aliases scan");
    let owner = source
        .picture(0)
        .expect("encoded picture")
        .svg_owner()
        .owner()
        .expect("encoded SVG owner");
    assert_eq!(owner.embedded_relationship_id(), Some("rIdSvgEncoded"));
    let standalone = owner.namespace_complete(16 * 1024).unwrap();
    assert!(
        standalone
            .windows(ASVG.len())
            .any(|window| window == ASVG.as_bytes())
    );
    assert!(
        standalone
            .windows(R.len())
            .any(|window| window == R.as_bytes())
    );
    assert!(
        !standalone
            .windows(b"&#x2F;".len())
            .any(|window| window == b"&#x2F;")
    );
}

#[test]
fn unknown_duplicate_and_mce_owners_are_inert_or_refused() {
    let unknown = document(&drawing(
        "inline",
        r#"<a:ext uri="urn:future"><future:payload xmlns:future="urn:future"/></a:ext>"#,
    ));
    let unknown = SourceDrawing::scan(&unknown).unwrap();
    assert!(matches!(
        unknown.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Opaque
    ));

    let duplicate = document(&drawing(
        "inline",
        &format!("{}{}", admitted("rIdOne"), admitted("rIdTwo")),
    ));
    let duplicate = SourceDrawing::scan(&duplicate).unwrap();
    assert!(matches!(
        duplicate.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Ambiguous
    ));

    let mce = document(&format!(
        r#"<mc:AlternateContent><mc:Choice Requires="w14">{}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        drawing("inline", &admitted("rIdHidden")),
    ));
    let mce = SourceDrawing::scan(&mce).unwrap();
    assert!(matches!(
        mce.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));
}

#[test]
fn strict_core_uses_strict_picture_uri_and_transitional_svg_relationship() {
    let body = r#"<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip trans:embed="rIdSvg"/></a:ext></a:extLst></a:blip></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
        .replace("{SVG_URI}", SVG_URI);
    let xml = format!(
        r#"<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:trans="{R}" xmlns:asvg="{ASVG}"><w:body>{}</w:body></w:document>"#,
        body
    );
    let source = SourceDrawing::scan(xml.as_bytes()).unwrap();
    assert_eq!(source.dialect(), DrawingDialect::Strict);
    assert_eq!(
        source
            .picture(0)
            .unwrap()
            .svg_owner()
            .owner()
            .unwrap()
            .embedded_relationship_id(),
        Some("rIdSvg")
    );
}

#[test]
fn recognized_container_text_and_missing_raster_fail_closed() {
    let text = document(&drawing(
        "inline",
        &format!(r#"<a:ext uri="{SVG_URI}">bad<asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#),
    ));
    assert!(SourceDrawing::scan(&text).is_err());

    let foreign_ext_list_child = document(&drawing("inline", r#"<a:foreign/>"#));
    assert!(SourceDrawing::scan(&foreign_ext_list_child).is_err());

    let missing_raster = document(&format!(
        r#"<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="{PICTURE_URI}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip><a:extLst>{}</a:extLst></a:blip></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
        admitted("rIdSvg"),
    ));
    let source = SourceDrawing::scan(&missing_raster).unwrap();
    assert!(matches!(
        source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));
}

#[test]
fn invalid_xml_namespace_and_declaration_forms_fail_closed() {
    let empty_prefix = document(r#"<w:p xmlns:x="" x:value="bad"/>"#);
    assert!(SourceDrawing::scan(&empty_prefix).is_err());

    let invalid_default = document(r#"<w:p xmlns="http://www.w3.org/XML/1998/namespace"/>"#);
    assert!(SourceDrawing::scan(&invalid_default).is_err());

    let mut invalid_version = b"<?xml version=".to_vec();
    invalid_version.extend_from_slice(br#"1.1"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#);
    assert!(SourceDrawing::scan(&invalid_version).is_err());
}

#[test]
fn namespace_context_budget_is_precharged_across_scope_nodes() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p xmlns:x="urn:one"><w:r/></w:p><w:p xmlns:y="urn:two"><w:r/></w:p></w:body></w:document>"#
    );
    let root_bytes = 1 + W.len();
    let first_scope_bytes = 1 + "urn:one".len();
    let second_scope_bytes = 1 + "urn:two".len();
    let below_aggregate = root_bytes + first_scope_bytes + second_scope_bytes - 1;
    let limits = ScanLimits {
        max_namespace_bytes: below_aggregate,
        ..ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(xml.as_bytes(), 0, limits).is_err());
    let exact = ScanLimits {
        max_namespace_bytes: below_aggregate + 1,
        ..limits
    };
    assert!(SourceDrawing::scan_with_limits(xml.as_bytes(), 0, exact).is_ok());
}

#[test]
fn ordinary_owner_rejects_foreign_root_and_ignores_legacy_picture_payload() {
    let foreign_before_root = format!(
        r#"<f:pre xmlns:f="urn:foreign"/><w:document xmlns:w="{W}"><w:body/></w:document>"#
    );
    assert!(SourceDrawing::scan(foreign_before_root.as_bytes()).is_err());

    let foreign_root = format!(
        r#"<f:outer xmlns:f="urn:foreign" xmlns:w="{W}"><w:document><w:body/></w:document></f:outer>"#
    );
    assert!(SourceDrawing::scan(foreign_root.as_bytes()).is_err());

    let legacy = document(&format!(
        r#"<w:pict>{}<w:object>{}</w:object></w:pict>"#,
        drawing("inline", &admitted("rIdHidden")),
        drawing("anchor", &admitted("rIdHidden2")),
    ));
    let source = SourceDrawing::scan(&legacy).expect("legacy payload remains opaque");
    assert!(source.pictures().is_empty());

    let opaque = String::from_utf8(document(&format!(
        r#"<f:opaque>{}</f:opaque>"#,
        drawing("inline", &admitted("rIdOpaque")),
    )))
    .unwrap()
        .replace(
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"#,
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:foreign"#,
        )
        .into_bytes();
    assert!(SourceDrawing::scan(&opaque).unwrap().pictures().is_empty());

    let nested_run = document(&format!(r#"<w:r>{}</w:r>"#, drawing("inline", "")));
    assert!(
        SourceDrawing::scan(&nested_run)
            .unwrap()
            .pictures()
            .is_empty()
    );

    let group = document(&format!(
        r#"<a:grpSp>{}</a:grpSp>"#,
        drawing("inline", &admitted("rIdGroup")),
    ));
    assert!(SourceDrawing::scan(&group).unwrap().pictures().is_empty());

    let unbound = document(r#"<x:opaque/>"#);
    assert!(SourceDrawing::scan(&unbound).is_err());
}

#[test]
fn direct_picture_shape_and_raster_choice_are_checked() {
    let out_of_order = document(
        &drawing(
            "inline",
            &format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#),
        )
        .replace("<pic:blipFill>", "<pic:spPr/><pic:blipFill>"),
    );
    assert!(SourceDrawing::scan(&out_of_order).is_err());

    let invalid_id = document(&drawing("inline", "").replace("id=\"7\"", "id=\"4294967296\""));
    assert!(SourceDrawing::scan(&invalid_id).is_err());

    let both_raster_relations = document(&drawing("inline", "").replace(
        "r:embed=\"rIdRaster\"",
        "r:embed=\"rIdRaster\" r:link=\"rIdOther\"",
    ));
    assert!(SourceDrawing::scan(&both_raster_relations).is_err());

    let unknown_ext_attribute = document(&drawing(
        "inline",
        &format!(
            r#"<a:ext uri="{SVG_URI}" future="opaque"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#
        ),
    ));
    let source = SourceDrawing::scan(&unknown_ext_attribute).unwrap();
    assert!(matches!(
        source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));

    let c_nv_pic_pr_first = document(&drawing("inline", "").replace(
        "<pic:cNvPr id=\"7\" name=\"picture\"/><pic:cNvPicPr/>",
        "<pic:cNvPicPr/><pic:cNvPr id=\"7\" name=\"picture\"/>",
    ));
    assert!(SourceDrawing::scan(&c_nv_pic_pr_first).is_err());

    let missing_name = document(&drawing("inline", "").replace(" name=\"picture\"", ""));
    assert!(SourceDrawing::scan(&missing_name).is_err());

    let style_after_ext_list = document(&drawing("inline", "").replace(
        "<pic:spPr/></pic:pic>",
        "<pic:spPr/><pic:extLst/><pic:style/></pic:pic>",
    ));
    assert!(SourceDrawing::scan(&style_after_ext_list).is_err());

    let duplicate_expanded_attribute = document(&drawing("inline", "").replace(
        "name=\"picture\"",
        &format!(
            "name=\"picture\" pic:future=\"one\" xmlns:p2=\"{}\" p2:future=\"two\"",
            encoded_slashes(PIC)
        ),
    ));
    assert!(SourceDrawing::scan(&duplicate_expanded_attribute).is_err());

    let unknown_graphic_data_child =
        document(&drawing("inline", "").replace("<pic:pic>", "<a:foreign/><pic:pic>"));
    assert!(SourceDrawing::scan(&unknown_graphic_data_child).is_err());

    let raster_link =
        document(&drawing("inline", "").replace("r:embed=\"rIdRaster\"", "r:link=\"rIdRaster\""));
    let raster_link = SourceDrawing::scan(&raster_link).unwrap();
    let link_picture = raster_link.picture(0).unwrap();
    assert!(link_picture.raster_relationship_is_link());
    assert!(matches!(link_picture.svg_owner(), SvgOwnerState::Refused));
}

#[test]
fn one_anchor_cannot_publish_multiple_direct_pictures_and_limits_are_clamped() {
    let one = drawing("inline", "");
    let pic_start = one.find("<pic:pic").unwrap();
    let pic_end = one.find("</pic:pic>").unwrap() + "</pic:pic>".len();
    let pic = &one[pic_start..pic_end];
    let insert_at = one.find("</a:graphicData>").unwrap();
    let duplicate = format!("{}{}{}", &one[..insert_at], pic, &one[insert_at..]);
    let duplicate = document(&duplicate);
    assert!(SourceDrawing::scan(&duplicate).is_err());

    let two = document(&format!(
        "{}{}",
        drawing("inline", ""),
        drawing("anchor", "")
    ));
    let limits = ScanLimits {
        max_pictures: 1,
        ..ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(&two, 0, limits).is_err());
}

#[test]
fn relationship_dialect_comes_from_host_profile_not_first_opaque_reference() {
    let xml = format!(
        r#"<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:trans="{R}" xmlns:asvg="{ASVG}"><w:body><w:hyperlink trans:id="rIdOpaque"/>{}</w:body></w:document>"#,
        drawing("inline", &admitted("rIdSvg")),
    );
    let source = SourceDrawing::scan(xml.as_bytes()).unwrap();
    assert_eq!(
        source.relationship_dialect(),
        super::source::RelationshipDialect::Strict
    );
}
