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
    let source = SourceDrawing::scan(&source_bytes, 7, ScanLimits::default()).unwrap();

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
    let source = SourceDrawing::scan(&source_bytes, 0, ScanLimits::default()).unwrap();

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

    let exact = SourceDrawing::scan(
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
        SourceDrawing::scan(
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
    let source = SourceDrawing::scan(&source_bytes, 0, ScanLimits::default()).unwrap();

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
    let source = SourceDrawing::scan(&source_bytes, 0, ScanLimits::default()).unwrap();

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
    let duplicate_source = SourceDrawing::scan(&duplicate_bytes, 0, ScanLimits::default()).unwrap();
    assert!(matches!(
        duplicate_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Ambiguous
    ));

    let mce = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="asvg"><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext></mc:Choice><mc:Fallback/></mc:AlternateContent>"#
    );
    let mce_bytes = transitional_drawing(&two_cell(1, &mce, ""));
    let mce_source = SourceDrawing::scan(&mce_bytes, 0, ScanLimits::default()).unwrap();
    assert!(matches!(
        mce_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Refused
    ));

    let unknown = r#"<a:ext uri="urn:future"><future:payload xmlns:future="urn:future" r:link="rIdOpaque"/></a:ext>"#;
    let unknown_bytes = transitional_drawing(&two_cell(1, unknown, ""));
    let unknown_source = SourceDrawing::scan(&unknown_bytes, 0, ScanLimits::default()).unwrap();
    assert!(matches!(
        unknown_source.picture(0).unwrap().svg_owner(),
        SvgOwnerState::Opaque
    ));

    let malformed = format!(
        r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/><future:payload/></a:ext>"#
    );
    let malformed_bytes = transitional_drawing(&two_cell(1, &malformed, ""));
    let malformed_source = SourceDrawing::scan(&malformed_bytes, 0, ScanLimits::default()).unwrap();
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
    assert!(SourceDrawing::scan(&source_bytes, 0, limited).is_err());
    assert!(
        SourceDrawing::scan(
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
        SourceDrawing::scan(
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
        SourceDrawing::scan(
            &source_bytes,
            0,
            ScanLimits {
                max_relationship_references: MAX_RELATIONSHIP_REFERENCES + 1,
                ..ScanLimits::default()
            }
        )
        .is_err()
    );

    let source = SourceDrawing::scan(&source_bytes, 0, ScanLimits::default()).unwrap();
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
    assert!(SourceDrawing::scan(foreign_root.as_bytes(), 0, ScanLimits::default()).is_err());
}

#[test]
fn cloned_source_records_keep_the_member_borrowed() {
    let source_bytes = transitional_drawing(&two_cell(1, "", ""));
    let drawing = SourceDrawing::scan(&source_bytes, 0, ScanLimits::default()).unwrap();
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

#[test]
fn non_picture_anchor_objects_do_not_hide_or_capture_neighbor_pictures() {
    for strict in [false, true] {
        for object in ["sp", "cxnSp", "graphicFrame"] {
            for anchored in [two_cell(8, "", ""), one_cell(8), absolute(8)] {
                let nested = picture(99, "", "");
                let replacement = format!("<xdr:{object}>{nested}</xdr:{object}>");
                let other = anchored.replace(&picture(8, "", ""), &replacement);
                let body = format!("{}{other}{}", two_cell(1, "", ""), two_cell(2, "", ""));
                let bytes = if strict {
                    strict_drawing(&body)
                } else {
                    transitional_drawing(&body)
                };
                let scan = SourceDrawing::scan(&bytes, 0, ScanLimits::default()).unwrap();
                assert_eq!(scan.pictures().len(), 2);
                assert_eq!(scan.pictures()[0].c_nv_pr_id(), Some("1"));
                assert_eq!(scan.pictures()[1].c_nv_pr_id(), Some("2"));
            }
        }
    }
}

#[test]
fn non_picture_objects_still_require_the_unique_ordered_anchor_slot() {
    for object in ["sp", "cxnSp", "graphicFrame"] {
        let object = format!("<xdr:{object}/>");
        for children in [
            format!(
                "{object}{}{}<xdr:clientData/>",
                marker("from", 0, 0),
                marker("to", 2, 2)
            ),
            format!(
                "{}{}{object}{object}<xdr:clientData/>",
                marker("from", 0, 0),
                marker("to", 2, 2)
            ),
            format!("{}{}{object}", marker("from", 0, 0), marker("to", 2, 2)),
            format!(
                "{}{}<xdr:clientData/>{object}",
                marker("from", 0, 0),
                marker("to", 2, 2)
            ),
        ] {
            let bytes = transitional_drawing(&format!(
                "<xdr:twoCellAnchor>{children}</xdr:twoCellAnchor>"
            ));
            assert!(SourceDrawing::scan(&bytes, 0, ScanLimits::default()).is_err());
        }
    }
}

/// One anchor child per letter of an anchor-order row.
fn anchor_child(letter: char) -> String {
    match letter {
        'F' => marker("from", 0, 0),
        'T' => marker("to", 2, 2),
        'P' => r#"<xdr:pos x="1" y="2"/>"#.to_owned(),
        'E' => r#"<xdr:ext cx="100" cy="200"/>"#.to_owned(),
        'O' => picture(1, "", ""),
        'C' => "<xdr:clientData/>".to_owned(),
        other => panic!("unknown anchor child letter {other}"),
    }
}

/// What the schema's anchor child sequence answers for `order`: the
/// children must be exactly `required`, in that order.
fn anchor_order_refusal(required: &str, order: &str) -> Option<&'static str> {
    for (index, child) in order.chars().enumerate() {
        match required.chars().nth(index) {
            None => return Some("drawing anchor has duplicate or trailing children"),
            Some(expected) if expected != child => {
                return Some("drawing anchor children are out of order");
            },
            Some(_) => {},
        }
    }
    (order.len() < required.len()).then_some("drawing anchor is missing required children")
}

/// Every order of up to five anchor children, for each anchor kind, is
/// accepted or refused by the schema sequence in the typed parser and in the
/// source scan. The source scan does not classify a `to`, `pos` or `ext`
/// under an anchor kind that has no such child, so it skips those.
#[test]
fn anchor_child_order_is_the_schema_sequence_in_both_readers() {
    const LETTERS: [char; 6] = ['F', 'T', 'P', 'E', 'O', 'C'];
    let mut orders = vec![String::new()];
    let mut start = 0;
    for _ in 0..5 {
        let end = orders.len();
        for index in start..end {
            for letter in LETTERS {
                let mut next = orders[index].clone();
                next.push(letter);
                orders.push(next);
            }
        }
        start = end;
    }
    let mut accepted = 0;
    for (anchor, required, source_skips) in [
        ("twoCellAnchor", "FTOC", "PE"),
        ("oneCellAnchor", "FEOC", "TP"),
        ("absoluteAnchor", "PEOC", "T"),
    ] {
        for order in &orders {
            let children = order.chars().map(anchor_child).collect::<String>();
            let bytes = transitional_drawing(&format!("<xdr:{anchor}>{children}</xdr:{anchor}>"));
            let xml = std::str::from_utf8(&bytes).expect("UTF-8 fixture");
            let typed = super::parse(xml, &litchi_opc::SourceLimits::default())
                .map(|_| ())
                .map_err(|error| error.to_string());
            let source = SourceDrawing::scan(&bytes, 0, ScanLimits::default())
                .map(|_| ())
                .map_err(|error| error.to_string());
            let expected_typed = anchor_order_refusal(required, order).map_or(Ok(()), |refusal| {
                Err(format!("invalid XLSX structure: {refusal}"))
            });
            assert_eq!(typed, expected_typed, "{anchor} {order}");
            let source_order = order
                .chars()
                .filter(|letter| !source_skips.contains(*letter))
                .collect::<String>();
            let expected_source =
                anchor_order_refusal(required, &source_order).map_or(Ok(()), |refusal| {
                    Err(format!(
                        "invalid XLSX structure: source-backed XLSX drawing: {refusal}"
                    ))
                });
            assert_eq!(source, expected_source, "{anchor} {order}");
            accepted += usize::from(typed.is_ok());
        }
    }
    assert_eq!(accepted, 3);
}

/// Freezes how the source drawing scanner treats what precedes, follows or replaces the
/// `xdr:wsDr` root, one row per document shape.
#[test]
fn source_drawing_envelope_matrix() {
    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
    let open = format!(r#"<xdr:wsDr xmlns:xdr="{XDR}">"#);
    let close = "</xdr:wsDr>";
    let empty = format!(r#"<xdr:wsDr xmlns:xdr="{XDR}"/>"#);
    let root = format!("{open}{close}");
    let rows: [(&str, String); 30] = [
        ("empty input", String::new()),
        ("whitespace only", " \n".to_owned()),
        ("declaration only", DECLARATION.to_owned()),
        ("empty-element root", empty.clone()),
        ("start and end root", root.clone()),
        ("root then whitespace", format!("{root}\n ")),
        ("root then comment", format!("{root}<!--c-->")),
        (
            "root then processing instruction",
            format!("{root}<?pi x?>"),
        ),
        ("root then text", format!("{root}text")),
        ("root then a second identical root", format!("{root}{root}")),
        (
            "root then a foreign element, start form",
            format!("{root}<foreign></foreign>"),
        ),
        (
            "root then a foreign element, empty form",
            format!("{root}<foreign/>"),
        ),
        ("declaration after the root", format!("{root}{DECLARATION}")),
        (
            "two declarations",
            format!("{DECLARATION}{DECLARATION}{root}"),
        ),
        (
            "comment before the declaration",
            format!("<!--c-->{DECLARATION}{root}"),
        ),
        ("text before the root", format!("text{root}")),
        ("unterminated root", open.clone()),
        ("stray end tag", format!("{root}{close}")),
        ("DOCTYPE", format!("<!DOCTYPE x>{root}")),
        (
            "top-level CDATA before the root",
            format!("<![CDATA[x]]>{root}"),
        ),
        (
            "top-level CDATA after the root",
            format!("{root}<![CDATA[x]]>"),
        ),
        (
            "one declaration then the root",
            format!("{DECLARATION}{root}"),
        ),
        (
            "whitespace then the declaration",
            format!(" {DECLARATION}{root}"),
        ),
        (
            "declaration inside the root",
            format!("{open}{DECLARATION}{close}"),
        ),
        ("empty-element root then text", format!("{empty}text")),
        (
            "empty-element root then a second root",
            format!("{empty}{empty}"),
        ),
        ("text inside the root", format!("{open}text{close}")),
        ("whitespace before the root", format!(" \n{root}")),
        (
            "nested wsDr inside the root",
            format!("{open}<xdr:wsDr></xdr:wsDr>{close}"),
        ),
        ("comment then the root", format!("<!--c-->{root}")),
    ];
    let expected: &[&str] = &[
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML has an incomplete xdr:wsDr root\"))",
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML has an incomplete xdr:wsDr root\"))",
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML has an incomplete xdr:wsDr root\"))",
        "Ok(pictures: 0)",
        "Ok(pictures: 0)",
        "Ok(pictures: 0)",
        "Ok(pictures: 0)",
        "Ok(pictures: 0)",
        "Err(Invalid(\"source-backed XLSX drawing: non-whitespace text appears outside xdr:wsDr\"))",
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML contains a second root element\"))",
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML contains a second root element\"))",
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML contains a second root element\"))",
        "Err(Invalid(\"source-backed XLSX drawing: XML declaration is not at the beginning of drawing XML\"))",
        "Err(Invalid(\"source-backed XLSX drawing: XML declaration is not at the beginning of drawing XML\"))",
        "Ok(pictures: 0)",
        "Err(Invalid(\"source-backed XLSX drawing: non-whitespace text appears outside xdr:wsDr\"))",
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML has an incomplete xdr:wsDr root\"))",
        "Err(Invalid(\"source-backed XLSX drawing: source-backed drawing XML: ill-formed document: close tag `</xdr:wsDr>` does not match any open tag\"))",
        "Err(Invalid(\"source-backed XLSX drawing: DOCTYPE is forbidden in drawing XML\"))",
        "Err(Invalid(\"source-backed XLSX drawing: CDATA appears outside drawing XML\"))",
        "Err(Invalid(\"source-backed XLSX drawing: CDATA appears outside drawing XML\"))",
        "Ok(pictures: 0)",
        "Ok(pictures: 0)",
        "Err(Invalid(\"source-backed XLSX drawing: XML declaration is not at the beginning of drawing XML\"))",
        "Err(Invalid(\"source-backed XLSX drawing: non-whitespace text appears outside xdr:wsDr\"))",
        "Err(Invalid(\"source-backed XLSX drawing: drawing XML contains a second root element\"))",
        "Err(Invalid(\"source-backed XLSX drawing: known drawing container contains non-whitespace text\"))",
        "Ok(pictures: 0)",
        "Ok(pictures: 0)",
        "Ok(pictures: 0)",
    ];
    let actual = rows
        .iter()
        .map(
            |(_, input)| match SourceDrawing::scan(input.as_bytes(), 0, ScanLimits::default()) {
                Ok(source) => format!("Ok(pictures: {})", source.pictures().len()),
                Err(error) => format!("Err({error:?})"),
            },
        )
        .collect::<Vec<_>>();
    let listing = actual
        .iter()
        .map(|value| format!("    {value:?},\n"))
        .collect::<String>();
    assert_eq!(actual.len(), expected.len(), "actual rows:\n{listing}");
    for (((label, input), actual), expected) in rows.iter().zip(&actual).zip(expected) {
        assert_eq!(
            actual, expected,
            "{label}: {input:?}\nactual rows:\n{listing}"
        );
    }
}
