use super::source::{MAX_PICTURES, MAX_RELATIONSHIP_REFERENCES, MAX_XML_NODES};
use super::{
    DrawingAnchor, DrawingDialect, RelationshipDialect, ScanLimits, SourceDrawing, SvgOwnerState,
};

const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SVG: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn marker(name: &str, column: i64, row: i64) -> String {
    format!(
        "<xdr:{name}><xdr:col>{column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{name}>"
    )
}

fn picture(id: usize, extensions: &str, extra_blip_content: &str) -> String {
    format!(
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id}" name="picture-{id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdRaster">{extra_blip_content}<a:extLst>{extensions}</a:extLst></a:blip></xdr:blipFill><xdr:spPr/></xdr:pic>"#
    )
}

fn two_cell(id: usize, extensions: &str, extra_blip_content: &str) -> String {
    format!(
        "<xdr:twoCellAnchor>{}{}{}<xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 0, 0),
        marker("to", 2, 2),
        picture(id, extensions, extra_blip_content)
    )
}

fn one_cell(id: usize) -> String {
    format!(
        r#"<xdr:oneCellAnchor>{}<xdr:ext cx="100" cy="200">{}</xdr:ext>{}<xdr:clientData/></xdr:oneCellAnchor>"#,
        marker("from", 3, 4),
        "",
        picture(id, "", "")
    )
}

fn absolute(id: usize) -> String {
    format!(
        r#"<xdr:absoluteAnchor><xdr:pos x="-10" y="20"/><xdr:ext cx="300" cy="400"/>{}<xdr:clientData/></xdr:absoluteAnchor>"#,
        picture(id, "", "")
    )
}

fn transitional_drawing(body: &str) -> Vec<u8> {
    format!(
        r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{REL}" xmlns:asvg="{SVG}" xmlns:mc="{MCE}" xmlns:future="urn:future">{body}</xdr:wsDr>"#
    )
    .into_bytes()
}

fn long_scope_drawing(body: &str) -> Vec<u8> {
    let mut root = format!(
        r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{REL}" xmlns:asvg="{SVG}" xmlns:mc="{MCE}" xmlns:future="urn:future""#
    );
    let uri_suffix = "x".repeat(128);
    for index in 0..128 {
        root.push_str(&format!(
            r#" xmlns:q{index}="urn:long-scope-{index}-{uri_suffix}""#
        ));
    }
    root.push('>');
    root.push_str(body);
    root.push_str("</xdr:wsDr>");
    root.into_bytes()
}

fn strict_drawing(body: &str) -> Vec<u8> {
    format!(
        r#"<xdr:wsDr xmlns:xdr="{STRICT_XDR}" xmlns:a="{STRICT_A}" xmlns:r="{STRICT_REL}" xmlns:trans="{REL}" xmlns:asvg="{SVG}">{body}</xdr:wsDr>"#
    )
    .into_bytes()
}

#[test]
fn scans_direct_pictures_in_all_anchor_forms_and_preserves_ranges() {
    let owner = format!(r#"<a:ext uri="  {SVG_URI}  "><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#);
    let opaque = r#"<a:ext uri="urn:future"><!--future ]] > payload--><![CDATA[opaque]]><?future payload?><future:payload r:link="rIdOpaque"/></a:ext>"#
        .replace("]] >", "]]>" );
    let body = format!(
        "<future:scope xmlns:r=\"urn:shadow\"/>{}{}{}",
        two_cell(1, &format!("{owner}{opaque}"), "<![CDATA[ ]]>&#x20;"),
        one_cell(2),
        absolute(3)
    );
    let source_bytes = transitional_drawing(&body);
    let source = SourceDrawing::scan_with_ordinal(&source_bytes, 7).unwrap();

    assert_eq!(source.drawing_ordinal(), 7);
    assert_eq!(source.dialect(), DrawingDialect::Transitional);
    assert_eq!(
        source.relationship_dialect(),
        RelationshipDialect::Transitional
    );
    assert_eq!(source.pictures().len(), 3);
    assert!(matches!(
        source.picture(0).unwrap().anchor(),
        DrawingAnchor::TwoCell { .. }
    ));
    assert!(matches!(
        source.picture(1).unwrap().anchor(),
        DrawingAnchor::OneCell { .. }
    ));
    assert!(matches!(
        source.picture(2).unwrap().anchor(),
        DrawingAnchor::Absolute { .. }
    ));

    let first = source.picture(0).unwrap();
    assert_eq!(first.c_nv_pr_id(), Some("1"));
    assert_eq!(first.raster_relationship_id(), "rIdRaster");
    assert!(first.blip_range().range().start < first.blip_range().range().end);
    assert!(!first.blip_range().is_empty());
    assert!(first.ext_list_range().is_some());
    assert!(!first.ext_list_range().unwrap().is_empty());
    assert!(first.picture_bytes(&source_bytes).unwrap().contains(&b'!'));
    assert!(
        first
            .relationship_references()
            .iter()
            .any(|reference| reference.id() == "rIdOpaque")
    );
    assert_eq!(source.relationship_references().len(), 5);

    let owner = first.svg_owner().owner().unwrap();
    assert_eq!(owner.uri_lexical(), format!("  {SVG_URI}  ").as_bytes());
    assert_eq!(owner.embedded_relationship_id(), Some("rIdSvg"));
    let complete = owner.namespace_complete(&source_bytes, 16 * 1024).unwrap();
    assert!(
        complete
            .windows(SVG.len())
            .any(|window| window == SVG.as_bytes())
    );
    assert!(
        first
            .namespace_complete_blip(&source_bytes, 16 * 1024)
            .unwrap()
            .starts_with(b"<a:blip")
    );
    assert!(
        first
            .namespace_complete_ext_list(&source_bytes, 16 * 1024)
            .unwrap()
            .is_some()
    );
}

#[test]
fn contextual_svg_values_share_scope_without_retaining_completed_fragments() {
    let owner = format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#);
    let body = (0..32)
        .map(|index| two_cell(index + 1, &owner, ""))
        .collect::<String>();
    let source_bytes = long_scope_drawing(&body);
    let source = SourceDrawing::scan(&source_bytes).unwrap();

    assert_eq!(source.pictures().len(), 32);
    let owners = source
        .pictures()
        .iter()
        .map(|picture| picture.svg_owner().owner().unwrap())
        .collect::<Vec<_>>();
    let first_context = owners[0].value().namespace_context().unwrap();
    assert!(first_context.binding_count() >= 128);
    assert!(owners.iter().all(|owner| owner.value().source().is_none()));
    assert!(owners.iter().all(|owner| {
        owner
            .value()
            .namespace_context()
            .is_some_and(|context| first_context.shares_storage(context))
    }));

    let mut raw_bytes = 0usize;
    for owner in owners {
        let raw = owner.value().raw_source().unwrap();
        let range = owner.svg_blip_range();
        assert_eq!(raw, &source_bytes[range.start..range.end]);
        assert!(
            !raw.windows(b"xmlns:q127".len())
                .any(|window| { window == b"xmlns:q127" })
        );
        raw_bytes += raw.len();
    }
    assert!(raw_bytes < source_bytes.len() / 2);
}

#[test]
fn contextual_projection_uses_raw_fragment_limit_boundary() {
    let raw = r#"<asvg:svgBlip r:embed="rIdSvg"/>"#;
    let owner = format!(r#"<a:ext uri="{SVG_URI}">{raw}</a:ext>"#);
    let source_bytes = transitional_drawing(&two_cell(1, &owner, ""));

    let exact = SourceDrawing::scan_with_limits(
        &source_bytes,
        0,
        ScanLimits {
            max_fragment_bytes: raw.len(),
            ..ScanLimits::default()
        },
    )
    .unwrap();
    assert_eq!(
        exact
            .picture(0)
            .unwrap()
            .svg_owner()
            .owner()
            .unwrap()
            .value()
            .raw_source(),
        Some(raw.as_bytes())
    );

    assert!(
        SourceDrawing::scan_with_limits(
            &source_bytes,
            0,
            ScanLimits {
                max_fragment_bytes: raw.len() - 1,
                ..ScanLimits::default()
            },
        )
        .is_err()
    );
}

#[test]
fn strict_host_keeps_transitional_svg_attribute_dialect() {
    let owner = format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip trans:embed="rIdSvg"/></a:ext>"#);
    let source_bytes = strict_drawing(&two_cell(1, &owner, ""));
    let source = SourceDrawing::scan(&source_bytes).unwrap();

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
fn strict_mixed_namespace_profile_keeps_shared_svg_relationships_typed() {
    let owner = format!(
        r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip xmlns="" xmlns:ns0="{REL}" ns0:embed="rId2"/></a:ext>"#
    );
    let body = format!("{}{}", two_cell(1, &owner, ""), two_cell(2, &owner, ""));
    let source_bytes = strict_drawing(&body);
    let source = SourceDrawing::scan(&source_bytes).unwrap();

    assert_eq!(source.pictures().len(), 2);
    assert_eq!(source.relationship_references().len(), 4);
    assert!(
        source
            .pictures()
            .iter()
            .all(|picture| picture.raster_relationship_id() == "rIdRaster")
    );
    assert!(
        source
            .pictures()
            .iter()
            .all(|picture| picture.svg_owner().owner().is_some())
    );
    assert!(source.pictures().iter().all(|picture| {
        picture
            .svg_owner()
            .owner()
            .unwrap()
            .embedded_relationship_id()
            == Some("rId2")
    }));
}

#[test]
fn duplicate_and_mce_owner_branches_are_refused_without_guessing() {
    let duplicate = format!(
        r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdOne"/></a:ext><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdTwo"/></a:ext>"#
    );
    let duplicate_bytes = transitional_drawing(&two_cell(1, &duplicate, ""));
    let duplicate_source = SourceDrawing::scan(&duplicate_bytes).unwrap();
    assert!(matches!(
        duplicate_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Ambiguous
    ));

    let mce = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="asvg"><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext></mc:Choice><mc:Fallback/></mc:AlternateContent>"#
    );
    let mce_bytes = transitional_drawing(&two_cell(1, &mce, ""));
    let mce_source = SourceDrawing::scan(&mce_bytes).unwrap();
    assert!(matches!(
        mce_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));

    let unknown = r#"<a:ext uri="urn:future"><future:payload xmlns:future="urn:future" r:link="rIdOpaque"/></a:ext>"#;
    let unknown_bytes = transitional_drawing(&two_cell(1, unknown, ""));
    let unknown_source = SourceDrawing::scan(&unknown_bytes).unwrap();
    assert!(matches!(
        unknown_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Opaque
    ));

    let malformed = format!(
        r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/><future:payload/></a:ext>"#
    );
    let malformed_bytes = transitional_drawing(&two_cell(1, &malformed, ""));
    let malformed_source = SourceDrawing::scan(&malformed_bytes).unwrap();
    assert!(matches!(
        malformed_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));
}

#[test]
fn source_limits_and_provenance_reject_unbounded_or_unrelated_inputs() {
    let source_bytes = transitional_drawing(&two_cell(1, "", ""));
    let limited = ScanLimits {
        max_nodes: 2,
        ..ScanLimits::default()
    };
    assert!(SourceDrawing::scan_with_limits(&source_bytes, 0, limited).is_err());
    assert!(
        SourceDrawing::scan_with_limits(
            &source_bytes,
            0,
            ScanLimits {
                max_nodes: MAX_XML_NODES + 1,
                ..ScanLimits::default()
            }
        )
        .is_err()
    );
    assert!(
        SourceDrawing::scan_with_limits(
            &source_bytes,
            0,
            ScanLimits {
                max_pictures: MAX_PICTURES + 1,
                ..ScanLimits::default()
            }
        )
        .is_err()
    );
    assert!(
        SourceDrawing::scan_with_limits(
            &source_bytes,
            0,
            ScanLimits {
                max_relationship_references: MAX_RELATIONSHIP_REFERENCES + 1,
                ..ScanLimits::default()
            }
        )
        .is_err()
    );

    let source = SourceDrawing::scan(&source_bytes).unwrap();
    let picture = source.picture(0).unwrap();
    let unrelated = source_bytes.clone();
    assert!(
        source
            .source_range(&unrelated, picture.picture_range())
            .is_err()
    );
    assert!(picture.picture_bytes(&unrelated).is_err());
    assert!(
        picture
            .namespace_complete_picture(&source_bytes, 1)
            .is_err()
    );

    let foreign_root = format!(
        r#"<future:root xmlns:future="urn:future"><xdr:wsDr xmlns:xdr="{XDR}"/></future:root>"#
    );
    assert!(SourceDrawing::scan(foreign_root.as_bytes()).is_err());
}

#[test]
fn cloned_source_records_keep_the_member_borrowed() {
    let source_bytes = transitional_drawing(&two_cell(1, "", ""));
    let drawing = SourceDrawing::scan(&source_bytes).unwrap();
    let picture = drawing.picture(0).unwrap().clone();
    let owner = picture.svg_owner().clone();
    drop(drawing);

    // The cloned projections still carry the original member borrow.  The
    // source Vec cannot be mutably changed while these values are alive, and
    // the records remain usable after the index itself is dropped.
    assert!(
        picture
            .namespace_complete_picture(&source_bytes, 16 * 1024)
            .is_ok()
    );
    assert!(owner.is_refused() || owner.owner().is_some() || matches!(owner, SvgOwnerState::None));
}
