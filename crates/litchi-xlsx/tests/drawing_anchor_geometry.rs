//! Independent `SpreadsheetDrawing` anchor geometry checks.
//!
//! The fixtures in this file follow the ECMA-376 strict/transitional
//! `dml-spreadsheetDrawing` shape: a two-cell anchor carries two markers, a
//! one-cell anchor carries a marker and a positive extent, and an absolute
//! anchor carries a position and a positive extent.  They intentionally use
//! source-distinct geometry so an absolute anchor cannot pass by projecting
//! to a zero cell anchor.

use std::fmt::Write as _;

use litchi_xlsx::drawing::{DrawingAnchor, EditAs, Object, UnknownKind, parse};
use litchi_xlsx::shapes::Emu;

const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

fn q(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

fn marker_named(
    xdr_prefix: &str,
    marker_name: &str,
    col: i64,
    col_off: i64,
    row: i64,
    row_off: i64,
) -> String {
    let marker = q(xdr_prefix, marker_name);
    let col_name = q(xdr_prefix, "col");
    let col_off_name = q(xdr_prefix, "colOff");
    let row_name = q(xdr_prefix, "row");
    let row_off_name = q(xdr_prefix, "rowOff");
    format!(
        "<{marker}><{col_name}>{col}</{col_name}><{col_off_name}>{col_off}</{col_off_name}><{row_name}>{row}</{row_name}><{row_off_name}>{row_off}</{row_off_name}></{marker}>",
    )
}

fn marker(xdr_prefix: &str, col: i64, col_off: i64, row: i64, row_off: i64) -> String {
    marker_named(xdr_prefix, "from", col, col_off, row, row_off)
}

fn to_marker(xdr_prefix: &str, col: i64, col_off: i64, row: i64, row_off: i64) -> String {
    marker_named(xdr_prefix, "to", col, col_off, row, row_off)
}

fn picture(xdr_prefix: &str, a_prefix: &str, r_prefix: &str, relationship_id: &str) -> String {
    let pic = q(xdr_prefix, "pic");
    let nv_pic_pr = q(xdr_prefix, "nvPicPr");
    let c_nv_pr = q(xdr_prefix, "cNvPr");
    let c_nv_pic_pr = q(xdr_prefix, "cNvPicPr");
    let blip_fill = q(xdr_prefix, "blipFill");
    let blip = q(a_prefix, "blip");
    let sp_pr = q(xdr_prefix, "spPr");
    format!(
        r#"<{pic}><{nv_pic_pr}><{c_nv_pr} id="1" name="picture"/><{c_nv_pic_pr}/></{nv_pic_pr}><{blip_fill}><{blip} {r_prefix}:embed="{relationship_id}"/></{blip_fill}><{sp_pr}/></{pic}>"#,
    )
}

fn direct_picture_anchor(
    xdr_prefix: &str,
    a_prefix: &str,
    r_prefix: &str,
    relationship_id: &str,
    edit_as: Option<&str>,
) -> String {
    let anchor = q(xdr_prefix, "twoCellAnchor");
    let client_data = q(xdr_prefix, "clientData");
    let edit_as = edit_as.map_or_else(String::new, |value| format!(r#" editAs="{value}""#));
    format!(
        r#"<{anchor}{edit_as}>{from}{to}{picture}<{client_data}/></{anchor}>"#,
        from = marker(xdr_prefix, 1, 2, 3, 4),
        to = to_marker(xdr_prefix, 5, 6, 7, 8),
        picture = picture(xdr_prefix, a_prefix, r_prefix, relationship_id),
        client_data = client_data,
    )
}

fn anchor_with_body(xdr_prefix: &str, body: &str) -> String {
    let anchor = q(xdr_prefix, "twoCellAnchor");
    let client_data = q(xdr_prefix, "clientData");
    let markers = format!(
        "{}{}",
        marker(xdr_prefix, 1, 2, 3, 4),
        to_marker(xdr_prefix, 5, 6, 7, 8)
    );
    format!("<{anchor}>{markers}{body}<{client_data}/></{anchor}>",)
}

fn transitional_root(body: &str) -> String {
    format!(
        r#"<x:wsDr xmlns:x="{XDR}" xmlns:d="{A}" xmlns:rel="{R}">{body}</x:wsDr>"#,
        XDR = XDR,
        A = A,
        R = R,
    )
}

fn strict_default_root(body: &str) -> String {
    format!(
        r#"<wsDr xmlns="{XDR}" xmlns:d="{A}" xmlns:rel="{R}">{body}</wsDr>"#,
        XDR = STRICT_XDR,
        A = STRICT_A,
        R = STRICT_R,
    )
}

fn is_picture(object: &Object) -> bool {
    matches!(object, Object::Picture(_))
}

fn only_pictures(xml: &str) -> litchi_xlsx::drawing::Drawing {
    parse(xml)
        .expect("drawing must parse")
        .expect("drawing root must be recognized")
}

fn picture_anchors(xml: &str) -> Vec<(String, DrawingAnchor)> {
    let drawing = only_pictures(xml);
    drawing
        .pictures()
        .map(|picture| (picture.relationship_id.clone(), *picture.drawing_anchor()))
        .collect()
}

fn assert_emu(value: Emu, expected: i64) {
    assert_eq!(value.emu(), expected);
}

#[test]
fn projects_all_anchor_forms_with_actual_geometry_and_edit_as() {
    let body = format!(
        "{}{}{}{}",
        direct_picture_anchor("x", "d", "rel", "rIdTwoDefault", None),
        direct_picture_anchor("x", "d", "rel", "rIdTwoOneCell", Some("oneCell")),
        {
            let anchor = q("x", "oneCellAnchor");
            let extent = q("x", "ext");
            let client_data = q("x", "clientData");
            format!(
                "<{anchor}>{from}<{extent} cx=\" 123456 \" cy=\"\n654321\t\"/>{picture}<{client_data}/></{anchor}>",
                from = marker("x", 9, 19, 10, 20),
                picture = picture("x", "d", "rel", "rIdOne"),
            )
        },
        {
            let anchor = q("x", "absoluteAnchor");
            let pos = q("x", "pos");
            let extent = q("x", "ext");
            let client_data = q("x", "clientData");
            format!(
                "<{anchor}><{pos} x=\" -900 \" y=\" 456 \"/><{extent} cx=\" 777888 \" cy=\"999000\t\"/>{picture}<{client_data}/></{anchor}>",
                picture = picture("x", "d", "rel", "rIdAbsolute"),
            )
        },
    );
    let pictures = picture_anchors(&transitional_root(&body));
    assert_eq!(pictures.len(), 4);
    assert_eq!(pictures[0].0, "rIdTwoDefault");
    assert_eq!(pictures[1].0, "rIdTwoOneCell");
    assert_eq!(pictures[2].0, "rIdOne");
    assert_eq!(pictures[3].0, "rIdAbsolute");

    match &pictures[0].1 {
        DrawingAnchor::TwoCell { from, to, edit_as } => {
            assert_eq!(*edit_as, EditAs::TwoCell);
            assert_eq!(from.column, 1);
            assert_emu(from.column_offset, 2);
            assert_eq!(from.row, 3);
            assert_emu(from.row_offset, 4);
            assert_eq!(to.column, 5);
            assert_emu(to.column_offset, 6);
            assert_eq!(to.row, 7);
            assert_emu(to.row_offset, 8);
        },
        other => panic!("expected default two-cell anchor, got {other:?}"),
    }

    match &pictures[1].1 {
        DrawingAnchor::TwoCell { edit_as, .. } => assert_eq!(*edit_as, EditAs::OneCell),
        other => panic!("expected editAs=oneCell two-cell anchor, got {other:?}"),
    }

    match &pictures[2].1 {
        DrawingAnchor::OneCell { from, extent } => {
            assert_eq!(from.column, 9);
            assert_emu(from.column_offset, 19);
            assert_eq!(from.row, 10);
            assert_emu(from.row_offset, 20);
            assert_emu(extent.width, 123_456);
            assert_emu(extent.height, 654_321);
        },
        other => panic!("expected one-cell anchor, got {other:?}"),
    }

    match &pictures[3].1 {
        DrawingAnchor::Absolute { position, extent } => {
            assert_emu(position.x, -900);
            assert_emu(position.y, 456);
            assert_emu(extent.width, 777_888);
            assert_emu(extent.height, 999_000);
        },
        other => panic!("expected absolute anchor with source geometry, got {other:?}"),
    }
}

#[test]
fn resolves_strict_default_and_transitional_custom_namespaces() {
    let transitional = transitional_root(&direct_picture_anchor(
        "x",
        "d",
        "rel",
        "rIdTransitional",
        Some("absolute"),
    ));
    let strict = strict_default_root(&direct_picture_anchor(
        "",
        "d",
        "rel",
        "rIdStrict",
        Some("twoCell"),
    ));

    let transitional = only_pictures(&transitional);
    let strict = only_pictures(&strict);
    assert_eq!(transitional.pictures().count(), 1);
    assert_eq!(strict.pictures().count(), 1);
    assert_eq!(
        transitional.pictures().next().unwrap().relationship_id,
        "rIdTransitional"
    );
    assert_eq!(
        strict.pictures().next().unwrap().relationship_id,
        "rIdStrict"
    );
    assert!(matches!(
        transitional.pictures().next().unwrap().drawing_anchor(),
        DrawingAnchor::TwoCell {
            edit_as: EditAs::Absolute,
            ..
        }
    ));
}

#[test]
fn only_direct_picture_blips_are_picture_owners() {
    let group = r#"<x:grpSp><x:pic><x:blipFill><d:blip rel:embed="rIdNested"/></x:blipFill></x:pic></x:grpSp>"#;
    let body = format!(
        "{}{}",
        direct_picture_anchor("x", "d", "rel", "rIdDirect", None),
        anchor_with_body("x", group),
    );
    let drawing = only_pictures(&transitional_root(&body));
    assert_eq!(drawing.pictures().count(), 1);
    assert_eq!(
        drawing.pictures().next().unwrap().relationship_id,
        "rIdDirect"
    );
    assert_eq!(
        drawing
            .unknown()
            .filter(|object| object.kind == UnknownKind::Group)
            .count(),
        1
    );
    assert_eq!(
        drawing
            .objects()
            .iter()
            .filter(|object| is_picture(object))
            .count(),
        1
    );
}

#[test]
fn unknown_payload_blip_does_not_replace_direct_picture_relationship() {
    let anchor = q("x", "twoCellAnchor");
    let client_data = q("x", "clientData");
    let pic = q("x", "pic");
    let nv_pic_pr = q("x", "nvPicPr");
    let c_nv_pr = q("x", "cNvPr");
    let c_nv_pic_pr = q("x", "cNvPicPr");
    let blip_fill = q("x", "blipFill");
    let blip = q("d", "blip");
    let sp_pr = q("x", "spPr");
    let unknown = r#"<future:opaque xmlns:future="urn:litchi:future"><d:blip rel:embed="rIdLookalike"/></future:opaque>"#;
    let body = format!(
        r#"<{anchor}>{from}{to}<{pic}><{nv_pic_pr}><{c_nv_pr} id="2" name="payload"/><{c_nv_pic_pr}/></{nv_pic_pr}><{blip_fill}><{blip} rel:embed="rIdRaster">{unknown}</{blip}></{blip_fill}><{sp_pr}/></{pic}><{client_data}/></{anchor}>"#,
        from = marker("x", 1, 2, 3, 4),
        to = to_marker("x", 5, 6, 7, 8),
    );
    let drawing = only_pictures(&transitional_root(&body));
    let picture = drawing.pictures().next().expect("direct picture");
    assert_eq!(picture.relationship_id, "rIdRaster");
    assert_eq!(drawing.pictures().count(), 1);
}

fn valid_picture_with_anchor(anchor_body: &str) -> String {
    format!(
        r#"<x:wsDr xmlns:x="{XDR}" xmlns:d="{A}" xmlns:rel="{R}">{anchor_body}</x:wsDr>"#,
        XDR = XDR,
        A = A,
        R = R,
    )
}

fn malformed_two_cell(body: &str) -> String {
    valid_picture_with_anchor(&format!(
        r#"<x:twoCellAnchor>{body}<x:pic><x:blipFill><d:blip rel:embed="rId"/></x:blipFill></x:pic><x:clientData/></x:twoCellAnchor>"#,
    ))
}

fn malformed_one_cell(body: &str) -> String {
    valid_picture_with_anchor(&format!(
        r#"<x:oneCellAnchor>{body}<x:pic><x:blipFill><d:blip rel:embed="rId"/></x:blipFill></x:pic><x:clientData/></x:oneCellAnchor>"#,
    ))
}

fn malformed_absolute(body: &str) -> String {
    valid_picture_with_anchor(&format!(
        r#"<x:absoluteAnchor>{body}<x:pic><x:blipFill><d:blip rel:embed="rId"/></x:blipFill></x:pic><x:clientData/></x:absoluteAnchor>"#,
    ))
}

#[test]
fn rejects_missing_duplicate_and_out_of_order_geometry() {
    let complete_from = marker("x", 1, 2, 3, 4);
    let complete_to = to_marker("x", 5, 6, 7, 8);

    for malformed in [
        malformed_two_cell(&complete_from),
        malformed_two_cell(&format!("{complete_to}{complete_to}")),
        malformed_two_cell(&format!("{complete_from}{complete_from}{complete_to}")),
        malformed_two_cell(&format!("{complete_to}{complete_from}")),
        malformed_two_cell(&format!(
            "<x:from><x:col>1</x:col><x:colOff>2</x:colOff><x:row>3</x:row></x:from>{complete_to}"
        )),
        malformed_one_cell(&format!(
            "{}<x:ext cx=\"12\" cy=\"13\"/><x:ext cx=\"14\" cy=\"15\"/>",
            complete_from
        )),
        malformed_one_cell(&format!("<x:ext cx=\"12\" cy=\"13\"/>{complete_from}")),
        malformed_absolute(r#"<x:ext cx="12" cy="13"/><x:pos x="14" y="15"/>"#),
        malformed_absolute(r#"<x:pos x="14" y="15"/>"#),
        malformed_absolute(
            r#"<x:pos x="14" y="15"/><x:ext cx="12" cy="13"/><x:ext cx="14" cy="15"/>"#,
        ),
    ] {
        assert!(
            parse(&malformed).is_err(),
            "accepted malformed drawing: {malformed}"
        );
    }
}

#[test]
fn rejects_invalid_edit_as_and_schema_bounds() {
    let bad_edit_as = valid_picture_with_anchor(&direct_picture_anchor(
        "x",
        "d",
        "rel",
        "rId",
        Some("not-an-edit-mode"),
    ));
    let negative_index = malformed_two_cell(&format!(
        "{}{}",
        marker("x", -1, 0, 0, 0),
        marker("x", 0, 0, 0, 0),
    ));
    let negative_extent = malformed_one_cell(&format!(
        "{}<x:ext cx=\"-1\" cy=\"12\"/>",
        marker("x", 0, 0, 0, 0)
    ));
    let nbsp = '\u{00a0}';
    let nbsp_extent = malformed_one_cell(&format!(
        "{}<x:ext cx=\"{nbsp}12\" cy=\"12\"/>",
        marker("x", 0, 0, 0, 0)
    ));
    let nbsp_marker_from = format!(
        "<x:from><x:col>{nbsp}1</x:col><x:colOff>0</x:colOff><x:row>0</x:row><x:rowOff>0</x:rowOff></x:from>"
    );
    let nbsp_marker =
        malformed_two_cell(&format!("{nbsp_marker_from}{}", to_marker("x", 1, 0, 1, 0)));

    for malformed in [
        bad_edit_as,
        negative_index,
        negative_extent,
        nbsp_extent,
        nbsp_marker,
    ] {
        assert!(
            parse(&malformed).is_err(),
            "accepted invalid drawing: {malformed}"
        );
    }
}

#[test]
fn applies_drawing_size_and_marker_text_limits_before_projection() {
    let oversized_marker = "7".repeat(65);
    let marker_limit = valid_picture_with_anchor(&format!(
        r#"<x:twoCellAnchor><x:from><x:col>{oversized_marker}</x:col><x:colOff>0</x:colOff><x:row>0</x:row><x:rowOff>0</x:rowOff></x:from><x:to><x:col>1</x:col><x:colOff>0</x:colOff><x:row>1</x:row><x:rowOff>0</x:rowOff></x:to><x:pic><x:blipFill><d:blip rel:embed="rId"/></x:blipFill></x:pic><x:clientData/></x:twoCellAnchor>"#
    ));
    assert!(parse(&marker_limit).is_err());

    let huge_comment = "x".repeat(32 * 1024 * 1024);
    let oversized_source = format!(
        r#"<x:wsDr xmlns:x="{XDR}" xmlns:d="{A}" xmlns:rel="{R}"><!--{huge_comment}--></x:wsDr>"#,
        XDR = XDR,
        A = A,
        R = R,
    );
    assert!(parse(&oversized_source).is_err());
}

#[test]
fn rejects_mce_output_that_exceeds_the_drawing_budget_after_expansion() {
    // MCE declares each namespace once (change 0653), except that the
    // declarations a dropped `mc:AlternateContent` wrapper made are hoisted
    // onto every child it emits. A large but bounded URI declared on the
    // wrapper therefore makes the processed XML exceed the 32 MiB drawing
    // limit while the caller input remains small. Each hoisted declaration
    // stays under the parser's 1 MiB attribute limit, so the MCE output bound
    // is what refuses the expansion.
    let padding = "x".repeat(768 * 1024);
    let source = |count: usize| {
        let mut children = String::new();
        for index in 0..count {
            write!(children, "<p:node{index}/>").expect("small fixture formatting cannot fail");
        }
        format!(
            r#"<x:wsDr xmlns:x="{XDR}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent xmlns:p="urn:{padding}"><mc:Choice Requires="p">{children}</mc:Choice><mc:Fallback>{children}</mc:Fallback></mc:AlternateContent></x:wsDr>"#,
            XDR = XDR,
        )
    };
    // One hoisted declaration fits, so the shape itself is admitted.
    assert!(parse(&source(1)).is_ok());
    let expanded = source(48);
    assert!(expanded.len() < 32 * 1024 * 1024);
    assert!(matches!(
        parse(&expanded),
        Err(litchi_xlsx::Error::MarkupCompatibility(
            litchi_ooxml_common::mce::Error::LimitExceeded(resource)
        )) if resource == "output bytes"
    ));
}
