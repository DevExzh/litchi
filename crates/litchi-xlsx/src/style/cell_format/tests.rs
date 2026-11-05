#![allow(
    clippy::float_cmp,
    reason = "parsed lexical values are compared with the exact values they spell"
)]

use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::sync::Arc;

use super::cache::MAX_RETAINED_WEIGHT;
use super::*;
use crate::error::Error;

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn parse_table(xml: &str) -> Arc<FormatTable> {
    Arc::new(FormatTable::parse(xml.as_bytes()).expect("styles parse"))
}

fn format(table: &Arc<FormatTable>, index: u32) -> Format {
    Format::new(Arc::clone(table), index).expect("format in range")
}

fn rgb(value: &str) -> Color {
    Color::from_rgb(Rgb::from_hex(value).expect("hex"))
}

fn tint(value: f64) -> Tint {
    Tint::new(value).expect("tint")
}

fn complete_sheet() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="{S}" xmlns:x14ac="urn:x14ac" xmlns:f="urn:foreign">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="0.00&amp; &quot;units&quot;"/>
    <numFmt numFmtId="14" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="3" x14ac:knownFonts="1">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
    <font>
      <b/><i val="0"/><strike val="true"/><outline/><shadow val="1"/><condense val="0"/><extend/>
      <u/><vertAlign val="superscript"/><sz val="10.5"/><color rgb="FFFF0000" tint="-0.25"/>
      <name val="A &amp; B"/><family val="1"/><charset val="128"/><scheme val="major"/>
      <f:unknown val="ignored"/><unknownChild/>
    </font>
    <font><u val="doubleAccounting"/><color indexed="12"/></font>
    <font><color auto="1"/></font>
  </fonts>
  <fills count="99">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="4" tint="0.39997558519241921"/><bgColor indexed="64"/></patternFill></fill>
    <fill><gradientFill type="path" degree="90" left="0.25" right="0.75" top="0.5" bottom="0.5">
      <stop position="0"><color rgb="FF000000"/></stop>
      <stop position="1"><color theme="0" tint="0.5"/></stop>
    </gradientFill></fill>
    <fill/>
    <fill><patternFill><fgColor rgb="00FF00"/></patternFill></fill>
    <fill><gradientFill><stop position="0.5"><color rgb="FF123456"/></stop></gradientFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border diagonalUp="1" diagonalDown="true" outline="0">
      <left style="thin"><color auto="1"/></left>
      <right style="none"><color rgb="FF00FF00"/></right>
      <top style="double"/>
      <bottom style="mediumDashDot"><color theme="3" tint="0.5"/></bottom>
      <diagonal style="hair"><color indexed="10"/></diagonal>
      <vertical style="dotted"/>
      <horizontal style="dashed"/>
    </border>
    <border><start style="thick"/><end style="medium"/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="1" fillId="2" borderId="1"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" quotePrefix="1">
      <alignment horizontal="centerContinuous" vertical="distributed" textRotation=" 255 " wrapText="true"
        indent="3" relativeIndent="-2" justifyLastLine="1" shrinkToFit="0" readingOrder="2"
        mergeCell="1" f:extra="ignored"/>
      <protection locked="0" hidden="1" f:extra="ignored"/>
      <extLst><ext uri="urn:test"><alignment horizontal="left"/></ext></extLst>
    </xf>
    <xf numFmtId="14" fontId="2" fillId="3" borderId="2"/>
    <xf numFmtId="49" fontId="3" fillId="4"/>
    <xf numFmtId="5" fontId="99" fillId="99" borderId="99" quotePrefix="0"/>
    <xf fillId="5"><protection/></xf>
    <xf numFmtId="170" fillId="6"><alignment/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="1"><dxf><font><b/></font></dxf></dxfs>
  <f:cellXfs><f:xf/></f:cellXfs>
</styleSheet>"#
    )
}

#[test]
fn reads_every_font_property_with_tri_state_booleans() {
    let table = parse_table(&complete_sheet());
    assert_eq!(table.len(), 7);

    let base = format(&table, 0);
    assert!(base.is_base());
    assert!(base.has_default_font());
    let font = base.font().expect("default font");
    assert_eq!(font.name(), Some("Calibri"));
    assert_eq!(font.size(), Some(11.0));
    assert_eq!(font.color(), Some(Color::theme(1)));
    assert_eq!(font.family(), Some(2));
    assert_eq!(font.scheme(), Some(Scheme::Minor));
    assert_eq!(font.bold(), None);
    assert_eq!(font.underline(), None);

    let styled = format(&table, 1);
    assert!(!styled.is_base());
    assert!(!styled.has_default_font());
    let font = styled.font().expect("font 1");
    assert_eq!(font.bold(), Some(true));
    assert_eq!(font.italic(), Some(false));
    assert_eq!(font.strike(), Some(true));
    assert_eq!(font.outline(), Some(true));
    assert_eq!(font.shadow(), Some(true));
    assert_eq!(font.condense(), Some(false));
    assert_eq!(font.extend(), Some(true));
    assert_eq!(font.underline(), Some(Underline::Single));
    assert_eq!(font.script(), Some(Script::Superscript));
    assert_eq!(font.size(), Some(10.5));
    assert_eq!(font.color(), Some(rgb("FFFF0000").with_tint(tint(-0.25))));
    assert_eq!(font.name(), Some("A & B"));
    assert_eq!(font.family(), Some(1));
    assert_eq!(font.charset(), Some(128));
    assert_eq!(font.scheme(), Some(Scheme::Major));

    let font = format(&table, 2).font().expect("font 2").clone();
    assert_eq!(font.underline(), Some(Underline::DoubleAccounting));
    assert_eq!(font.color(), Some(Color::indexed(12)));
    assert_eq!(font.name(), None);
    assert_eq!(font.size(), None);
    assert_eq!(
        format(&table, 3).font().and_then(Font::color),
        Some(Color::auto())
    );
}

#[test]
fn reads_number_formats_declared_built_in_and_locale_dependent() {
    let table = parse_table(&complete_sheet());
    let general_format = format(&table, 0);
    let general = general_format.number_format();
    assert!(general.is_general());
    assert_eq!(general.code(), Some("General"));
    assert!(!general.is_declared());

    let custom_format = format(&table, 1);
    let custom = custom_format.number_format();
    assert_eq!(custom.id(), 164);
    assert_eq!(custom.code(), Some("0.00& \"units\""));
    assert!(custom.is_declared());
    assert!(!custom.is_date_time());

    // A workbook declaration of a built-in identifier wins over the table.
    let redefined_format = format(&table, 2);
    let redefined = redefined_format.number_format();
    assert_eq!(redefined.id(), 14);
    assert_eq!(redefined.code(), Some("yyyy-mm-dd"));
    assert!(redefined.is_declared());
    assert!(redefined.is_date_time());

    let text_format = format(&table, 3);
    let text = text_format.number_format();
    assert_eq!(
        (text.id(), text.code(), text.is_declared()),
        (49, Some("@"), false)
    );

    let currency_format = format(&table, 4);
    let currency = currency_format.number_format();
    assert_eq!((currency.id(), currency.code()), (5, None));
    assert!(!currency.is_date_time());

    // An undeclared custom identifier has no code.
    assert_eq!(format(&table, 6).number_format().code(), None);
    // A record without numFmtId is General.
    assert!(format(&table, 5).number_format().is_general());
}

#[test]
fn reads_pattern_gradient_and_empty_fills_with_typed_colors() {
    let table = parse_table(&complete_sheet());
    assert!(format(&table, 0).has_default_fill());
    assert_eq!(
        format(&table, 0).fill(),
        Some(&Fill::Pattern(PatternFill {
            pattern: Some(Pattern::None),
            foreground: None,
            background: None,
        }))
    );

    let solid = format(&table, 1);
    assert!(!solid.has_default_fill());
    let fill = solid.fill().and_then(Fill::pattern).expect("pattern fill");
    assert_eq!(fill.pattern(), Some(Pattern::Solid));
    assert_eq!(
        fill.foreground(),
        Some(Color::theme(4).with_tint(tint(0.399_975_585_192_419_2)))
    );
    assert_eq!(fill.background(), Some(Color::indexed(64)));

    let gradient = format(&table, 2)
        .fill()
        .and_then(Fill::gradient)
        .expect("gradient fill")
        .clone();
    assert!(gradient.validate().is_ok());
    assert_eq!(gradient.kind(), GradientKind::Path);
    assert_eq!(
        (
            gradient.degree(),
            gradient.left(),
            gradient.right(),
            gradient.top(),
            gradient.bottom()
        ),
        (90.0, 0.25, 0.75, 0.5, 0.5)
    );
    assert_eq!(gradient.stops().len(), 2);
    assert_eq!(gradient.stops()[0].position(), 0.0);
    assert_eq!(gradient.stops()[0].color(), rgb("FF000000"));
    assert_eq!(
        gradient.stops()[1].color(),
        Color::theme(0).with_tint(tint(0.5))
    );

    // `<fill/>` authors nothing.
    assert_eq!(
        format(&table, 3).fill(),
        Some(&Fill::Pattern(PatternFill::default()))
    );
    // A pattern fill without a pattern keeps its colour; six digits are opaque.
    let unpatterned = format(&table, 5).fill().and_then(Fill::pattern).copied();
    assert_eq!(
        unpatterned,
        Some(PatternFill {
            pattern: None,
            foreground: Some(rgb("FF00FF00")),
            background: None,
        })
    );
    let defaults = format(&table, 6)
        .fill()
        .and_then(Fill::gradient)
        .expect("gradient")
        .clone();
    assert_eq!(defaults.kind(), GradientKind::Linear);
    assert_eq!(defaults.degree(), 0.0);
    assert_eq!(defaults.stops()[0].position(), 0.5);
}

#[test]
fn authored_pattern_and_gradient_constructors_preserve_typed_fields() {
    const EMPTY_PATTERN: PatternFill = PatternFill::new(None, None, None);
    assert_eq!(EMPTY_PATTERN, PatternFill::default());

    let pattern = PatternFill::new(
        Some(Pattern::Solid),
        Some(Color::rgb(0x11, 0x22, 0x33)),
        Some(Color::theme(4)),
    );
    assert_eq!(pattern.pattern(), Some(Pattern::Solid));
    assert_eq!(pattern.foreground(), Some(Color::rgb(0x11, 0x22, 0x33)));
    assert_eq!(pattern.background(), Some(Color::theme(4)));

    let stops = vec![
        GradientStop::new(-0.0, Color::rgb(0, 0, 0)).expect("first stop"),
        GradientStop::new(1.0, Color::theme(1)).expect("second stop"),
    ]
    .into_boxed_slice();
    let linear = GradientFill::linear(-0.0, stops).expect("linear gradient");
    assert_eq!(linear.kind(), GradientKind::Linear);
    assert_eq!(linear.degree(), 0.0);
    assert_eq!(
        (linear.left(), linear.right(), linear.top(), linear.bottom()),
        (0.0, 0.0, 0.0, 0.0)
    );
    assert_eq!(linear.stops()[0].position(), 0.0);
    assert_eq!(linear.stops()[1].position(), 1.0);

    let path = GradientFill::path(
        -0.0,
        1.0,
        0.25,
        0.75,
        linear.stops().to_vec().into_boxed_slice(),
    )
    .expect("path gradient");
    assert_eq!(path.kind(), GradientKind::Path);
    assert_eq!(path.degree(), 0.0);
    assert_eq!(
        (path.left(), path.right(), path.top(), path.bottom()),
        (0.0, 1.0, 0.25, 0.75)
    );
    assert_eq!(path.stops(), linear.stops());
}

#[test]
fn gradient_constructors_reject_non_finite_and_out_of_range_values() {
    let color = Color::rgb(0, 0, 0);
    for position in [f64::NAN, f64::NEG_INFINITY, f64::INFINITY, -0.1, 1.1] {
        assert!(GradientStop::new(position, color).is_err());
    }

    let stops = || vec![GradientStop::new(0.0, color).expect("stop")].into_boxed_slice();
    for degree in [f64::NAN, f64::NEG_INFINITY, f64::INFINITY, -0.1, 360.1] {
        assert!(GradientFill::linear(degree, stops()).is_err());
    }

    assert!(GradientFill::path(-0.1, 0.0, 0.0, 0.0, stops()).is_err());
    assert!(GradientFill::path(0.0, 1.1, 0.0, 0.0, stops()).is_err());
    assert!(GradientFill::path(0.0, 0.0, f64::NAN, 0.0, stops()).is_err());
    assert!(GradientFill::path(0.0, 0.0, 0.0, f64::INFINITY, stops()).is_err());
}

#[test]
fn gradient_constructor_bounds_stop_collection_and_hashes_signed_zero_canonically() {
    let stop = GradientStop::new(-0.0, Color::rgb(0, 0, 0)).expect("stop");
    let at_limit = vec![stop; 1].into_boxed_slice();
    assert!(GradientFill::linear(360.0, at_limit).is_ok());
    assert!(validate_stop_count(MAX_RECORDS).is_ok());
    assert!(matches!(
        validate_stop_count(MAX_RECORDS + 1),
        Err(Error::ResourceLimit(limit)) if limit.observed == 1_000_001 && limit.limit == 1_000_000
    ));

    let positive_stop = GradientStop {
        position: 0.0,
        color: Color::rgb(0, 0, 0),
    };
    let negative_stop = GradientStop {
        position: -0.0,
        color: Color::rgb(0, 0, 0),
    };
    assert_eq!(positive_stop, negative_stop);
    assert_eq!(hash_value(&positive_stop), hash_value(&negative_stop));

    let positive = GradientFill {
        kind: GradientKind::Linear,
        degree: 0.0,
        left: 0.0,
        right: 0.0,
        top: 0.0,
        bottom: 0.0,
        stops: vec![positive_stop].into_boxed_slice(),
    };
    let negative = GradientFill {
        kind: GradientKind::Linear,
        degree: -0.0,
        left: -0.0,
        right: -0.0,
        top: -0.0,
        bottom: -0.0,
        stops: vec![negative_stop].into_boxed_slice(),
    };
    assert_eq!(positive, negative);
    assert_eq!(hash_value(&positive), hash_value(&negative));
    assert_eq!(
        hash_value(&Fill::Gradient(positive)),
        hash_value(&Fill::Gradient(negative))
    );

    let read_created_outside_authoring_profile = GradientFill {
        kind: GradientKind::Linear,
        degree: 361.0,
        left: 0.0,
        right: 0.0,
        top: 0.0,
        bottom: 0.0,
        stops: Box::new([]),
    };
    assert!(read_created_outside_authoring_profile.validate().is_err());
    assert!(
        Fill::Gradient(read_created_outside_authoring_profile)
            .validate()
            .is_err()
    );
}

fn hash_value<T: Hash>(value: &T) -> u64 {
    let mut hasher = DefaultHasher::new();
    value.hash(&mut hasher);
    hasher.finish()
}

#[test]
fn reads_borders_alignment_protection_and_quote_prefix() {
    let table = parse_table(&complete_sheet());
    let base = format(&table, 0);
    assert!(base.has_default_border());
    assert!(!base.border().expect("default border").has_borders());
    assert_eq!(base.alignment(), None);
    assert_eq!(base.protection(), None);
    assert!(!base.quote_prefix());

    let styled = format(&table, 1);
    assert!(!styled.has_default_border());
    let border = styled.border().expect("border 1");
    assert_eq!(border.outline, Some(false));
    assert_eq!(
        border.left,
        Some(Side::new(Line::Thin).with_color(Color::auto()))
    );
    assert_eq!(
        border.right, None,
        "style none is no side, even with a color"
    );
    assert_eq!(border.top, Some(Side::new(Line::Double)));
    assert_eq!(
        border.bottom,
        Some(Side::new(Line::MediumDashDot).with_color(Color::theme(3).with_tint(tint(0.5))))
    );
    let diagonal = border.diagonal.as_ref().expect("diagonal");
    assert_eq!(diagonal.dir(), Some(Dir::Both));
    assert_eq!(
        diagonal.side(),
        Some(&Side::new(Line::Hair).with_color(Color::indexed(10)))
    );
    assert_eq!(border.vertical, Some(Side::new(Line::Dotted)));
    assert_eq!(border.horizontal, Some(Side::new(Line::Dashed)));

    let strict_edges = format(&table, 2).border().expect("border 2").clone();
    assert_eq!(strict_edges.start, Some(Side::new(Line::Thick)));
    assert_eq!(strict_edges.end, Some(Side::new(Line::Medium)));

    let alignment = styled.alignment().expect("alignment");
    assert_eq!(alignment.horizontal, Some(Horizontal::CenterContinuous));
    assert_eq!(alignment.vertical, Some(Vertical::Distributed));
    assert_eq!(alignment.text_rotation, Some(Rotation::stacked()));
    assert!(alignment.wrap_text);
    assert_eq!(alignment.indent, Some(Indent::new(3)));
    assert_eq!(alignment.relative_indent, Some(-2));
    assert_eq!(alignment.justify_last_line, Some(true));
    assert!(!alignment.shrink_to_fit);
    assert_eq!(alignment.reading_order, Some(Reading::RightToLeft));

    let protection = styled.protection().expect("protection");
    assert_eq!(
        (protection.locked(), protection.hidden()),
        (Some(false), Some(true))
    );
    assert!(styled.quote_prefix());

    let empty = format(&table, 5);
    assert_eq!(empty.protection(), Some(Protection::default()));
    assert_eq!(format(&table, 6).alignment(), Some(&Alignment::new()));
    assert!(!format(&table, 4).quote_prefix());
}

#[test]
fn dangling_component_references_resolve_to_none() {
    let table = parse_table(&complete_sheet());
    let dangling = format(&table, 4);
    assert_eq!(dangling.font(), None);
    assert_eq!(dangling.fill(), None);
    assert_eq!(dangling.border(), None);
    assert!(!dangling.has_default_font());

    let without_components = parse_table(&format!(
        r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs></styleSheet>"#
    ));
    let only = format(&without_components, 0);
    assert!(only.has_default_font());
    assert_eq!(only.font(), None);
    assert_eq!(only.fill(), None);
    assert_eq!(only.border(), None);
}

#[test]
fn format_handles_check_their_index_and_share_their_table() {
    let table = parse_table(&complete_sheet());
    assert!(Format::new(Arc::clone(&table), 7).is_err());
    assert!(Format::new(Arc::clone(&table), u32::MAX).is_err());
    let first = format(&table, 1);
    let again = format(&table, 1);
    assert!(first.same(&again));
    assert!(!first.same(&format(&table, 2)));
    let debug = format!("{first:?}");
    assert!(debug.contains("is_base: false"), "{debug}");
    assert!(!debug.contains("Superscript"), "{debug}");
}

#[test]
fn reads_strict_namespaces_and_markup_compatibility_like_the_catalog() {
    let strict = format!(
        r#"<x:styleSheet xmlns:x="{STRICT}"><x:fonts><x:font><x:b/></x:font></x:fonts>
           <x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
    );
    let table = parse_table(&strict);
    assert_eq!(table.len(), 1);
    assert_eq!(format(&table, 0).font().and_then(Font::bold), Some(true));

    let alternate = format!(
        r#"<styleSheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x="urn:future">
           <fonts><font><mc:AlternateContent><mc:Choice Requires="x"><x:b/></mc:Choice>
             <mc:Fallback><i/></mc:Fallback></mc:AlternateContent></font></fonts>
           <mc:AlternateContent><mc:Choice Requires="x"><cellXfs count="1"><xf/></cellXfs></mc:Choice>
             <mc:Fallback><cellXfs count="2"><xf fontId="0"/><xf/></cellXfs></mc:Fallback>
           </mc:AlternateContent></styleSheet>"#
    );
    let table = parse_table(&alternate);
    assert_eq!(table.len(), 2);
    let catalog = crate::raw::styles::parse(alternate.as_bytes()).expect("catalog");
    assert_eq!(u32::try_from(table.len()), Ok(catalog.len()));
    let font = format(&table, 0).font().expect("font").clone();
    assert_eq!((font.bold(), font.italic()), (None, Some(true)));
}

#[test]
fn counts_cell_formats_exactly_as_the_shared_style_catalog() {
    let cases = [
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs count="2"><xf/><xf><alignment/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><unrelated><xf/><cellXfs><xf/></cellXfs></unrelated><cellXfs count="2"><xf><nested><xf/></nested></xf><unrelated><xf/></unrelated><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}" xmlns:f="urn:f"><cellXfs><xf/><f:xf/></cellXfs><f:cellXfs><f:xf/></f:cellXfs></styleSheet>"#
        ),
        format!(r#"<styleSheet xmlns="{S}"/>"#),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs/></styleSheet>"#),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs/><cellXfs><xf/></cellXfs></styleSheet>"#),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs count="2"><xf/></cellXfs></styleSheet>"#),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs count="x"><xf/></cellXfs></styleSheet>"#),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs count="65431"><xf/></cellXfs></styleSheet>"#),
        r#"<styleSheet xmlns="urn:other"><cellXfs><xf/></cellXfs></styleSheet>"#.to_owned(),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs></styleSheet><tail/>"#),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs>"#),
    ];
    for xml in cases {
        let catalog = crate::raw::styles::parse(xml.as_bytes());
        let parsed = FormatTable::parse(xml.as_bytes());
        match (&catalog, &parsed) {
            (Ok(catalog), Ok(parsed)) => {
                assert_eq!(u32::try_from(parsed.len()), Ok(catalog.len()), "{xml}");
            },
            (Err(_), Err(_)) => {},
            _ => panic!("catalog {catalog:?} and format table {parsed:?} disagree on {xml}"),
        }
    }
}

#[test]
fn refuses_values_outside_their_schema_domain() {
    let wrap = |body: &str| {
        format!(r#"<styleSheet xmlns="{S}">{body}<cellXfs><xf/></cellXfs></styleSheet>"#)
    };
    let cases = [
        wrap(r#"<fonts><font><u val="Single"/></font></fonts>"#),
        wrap(r#"<fonts><font><b val="maybe"/></font></fonts>"#),
        wrap(r#"<fonts><font><b/><b/></font></fonts>"#),
        wrap(r#"<fonts><font><sz val="0"/></font></fonts>"#),
        wrap(r#"<fonts><font><sz val="NaN"/></font></fonts>"#),
        wrap(r#"<fonts><font><sz/></font></fonts>"#),
        wrap(r#"<fonts><font><vertAlign val="raised"/></font></fonts>"#),
        wrap(r#"<fonts><font><scheme val="body"/></font></fonts>"#),
        wrap(r#"<fonts><font><color rgb="FF000000" theme="1"/></font></fonts>"#),
        wrap(r#"<fonts><font><color rgb="F00"/></font></fonts>"#),
        wrap(r#"<fonts><font><color theme="1" tint="1.5"/></font></fonts>"#),
        wrap(r#"<fonts><font><family val="-1"/></font></fonts>"#),
        wrap(r#"<numFmts><numFmt formatCode="0"/></numFmts>"#),
        wrap(r#"<numFmts><numFmt numFmtId="164"/></numFmts>"#),
        wrap(
            r#"<numFmts><numFmt numFmtId="164" formatCode="0"/><numFmt numFmtId="164" formatCode="1"/></numFmts>"#,
        ),
        wrap(r#"<fills><fill><patternFill patternType="dotted"/></fill></fills>"#),
        wrap(
            r#"<fills><fill><patternFill><fgColor rgb="FF000000"/><fgColor rgb="FF000000"/></patternFill></fill></fills>"#,
        ),
        wrap(r#"<fills><fill><patternFill/><gradientFill/></fill></fills>"#),
        wrap(r#"<fills><fill><gradientFill type="radial"/></fill></fills>"#),
        wrap(r#"<fills><fill><gradientFill degree="x"/></fill></fills>"#),
        wrap(
            r#"<fills><fill><gradientFill><stop position="2"><color rgb="FF000000"/></stop></gradientFill></fill></fills>"#,
        ),
        wrap(r#"<fills><fill><gradientFill><stop position="0"/></gradientFill></fill></fills>"#),
        wrap(
            r#"<fills><fill><gradientFill><stop><color rgb="FF000000"/></stop></gradientFill></fill></fills>"#,
        ),
        wrap(r#"<borders><border><left style="wavy"/></border></borders>"#),
        wrap(r#"<borders><border><left/><left/></border></borders>"#),
        wrap(r#"<borders><border diagonalUp="2"/></borders>"#),
        wrap(
            r#"<borders><border><top style="thin"><color auto="1"/><color auto="1"/></top></border></borders>"#,
        ),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf fontId="-1"/></cellXfs></styleSheet>"#),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf quotePrefix="yes"/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment horizontal="middle"/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment textRotation="200"/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment indent="256"/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment readingOrder="3"/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment/><alignment/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><protection locked="2"/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><protection/><protection/></xf></cellXfs></styleSheet>"#
        ),
        format!(r#"<styleSheet xmlns="{S}"><fonts/><fonts/><cellXfs><xf/></cellXfs></styleSheet>"#),
        format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf a="1" a="2"/></cellXfs></styleSheet>"#),
    ];
    for xml in cases {
        let error = FormatTable::parse(xml.as_bytes()).expect_err(&xml);
        assert!(
            matches!(error, Error::Invalid(_) | Error::MarkupCompatibility(_)),
            "{xml}: {error:?}"
        );
    }
}

#[test]
fn parser_errors_do_not_echo_authored_values() {
    const MARKER: &str = "authored-cell-format-marker";
    let cases = [
        format!(
            r#"<styleSheet xmlns="{S}"><fonts><font><u val="{MARKER}"/></font></fonts><cellXfs><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><fonts><font><b val="{MARKER}"/></font></fonts><cellXfs><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><fills><fill><patternFill patternType="{MARKER}"/></fill></fills><cellXfs><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><fills><fill><gradientFill type="{MARKER}"/></fill></fills><cellXfs><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><fills><fill><gradientFill><stop position="{MARKER}"/></gradientFill></fill></fills><cellXfs><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><borders><border><left style="{MARKER}"/></border></borders><cellXfs><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf numFmtId="{MARKER}"/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><fonts><font><color rgb="{MARKER}"/></font></fonts><cellXfs><xf/></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment horizontal="{MARKER}"/></xf></cellXfs></styleSheet>"#
        ),
        format!(
            r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment relativeIndent="{MARKER}"/></xf></cellXfs></styleSheet>"#
        ),
    ];
    for xml in cases {
        let error = FormatTable::parse(xml.as_bytes()).expect_err(&xml);
        let rendered = error.to_string();
        assert!(
            !rendered.contains(MARKER),
            "parser error echoed authored input: {rendered}"
        );
    }
}

#[test]
fn ignores_extension_markup_and_collection_counts() {
    let xml = format!(
        r#"<styleSheet xmlns="{S}" xmlns:f="urn:foreign">
           <fonts count="9"><font><f:b val="nonsense"/><mystery val="x"/><b/></font><f:font/></fonts>
           <fills count="0"><fill><patternFill patternType="solid" f:x="1"><f:fgColor rgb="nonsense"/></patternFill></fill></fills>
           <borders><border f:x="1"><f:left style="nonsense"/><unknown style="nonsense"/></border></borders>
           <cellXfs><xf fontId="0" f:fontId="x" applyFont="maybe"><alignment f:horizontal="nonsense" unknown="x"/></xf></cellXfs>
         </styleSheet>"#
    );
    let table = parse_table(&xml);
    assert_eq!(table.len(), 1);
    let only = format(&table, 0);
    assert_eq!(only.font().and_then(Font::bold), Some(true));
    assert_eq!(
        only.fill()
            .and_then(Fill::pattern)
            .map(PatternFill::foreground),
        Some(None)
    );
    assert!(!only.border().expect("border").has_borders());
    assert_eq!(only.alignment(), Some(&Alignment::new()));
}

#[test]
fn pattern_tokens_round_trip() {
    for pattern in Pattern::ALL {
        assert_eq!(Pattern::from_token(pattern.as_str()), Some(pattern));
        assert_eq!(pattern.to_string(), pattern.as_str());
    }
    assert_eq!(Pattern::from_token("Solid"), None);
    assert_eq!(GradientKind::from_token("path"), Some(GradientKind::Path));
    assert_eq!(GradientKind::from_token("Path"), None);
    assert_eq!(GradientKind::default().as_str(), "linear");
}

#[test]
fn weight_counts_retained_records_and_strings() {
    let small = parse_table(&format!(
        r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs></styleSheet>"#
    ));
    let large = parse_table(&complete_sheet());
    assert!(small.weight() >= size_of::<FormatTable>());
    assert!(large.weight() > small.weight());

    // The format's maximum record count stays below the retention ceiling.
    let mut xml = format!(r#"<styleSheet xmlns="{S}"><cellXfs count="65430">"#);
    for _ in 0..65_430 {
        xml.push_str(
            r#"<xf numFmtId="1" fontId="1" fillId="1" borderId="1" quotePrefix="1"><alignment horizontal="left" relativeIndent="1"/><protection locked="1"/></xf>"#,
        );
    }
    xml.push_str("</cellXfs></styleSheet>");
    let maximal = parse_table(&xml);
    assert_eq!(maximal.len(), 65_430);
    // The retention ceiling's documentation quotes these figures.
    #[cfg(target_pointer_width = "64")]
    {
        assert_eq!(size_of::<CellXf>(), 56);
        assert_eq!(size_of::<Font>(), 88);
        assert_eq!(size_of::<Fill>(), 64);
        assert_eq!(size_of::<Border>(), 304);
        assert_eq!(maximal.weight(), 3_664_168);
    }
    assert!(
        maximal.weight() < MAX_RETAINED_WEIGHT,
        "maximal cellXfs table weighs {} bytes",
        maximal.weight()
    );
}

#[test]
fn cache_hits_one_generation_and_replaces_it_on_another() {
    let cache = FormatCache::new();
    let first = Arc::new(complete_sheet().into_bytes());
    let mut builds = 0;
    let read = |payload: &Arc<Vec<u8>>, builds: &mut i32| {
        cache
            .get_or_build(FormatKey::allocation(payload), || {
                *builds += 1;
                FormatTable::parse(payload)
            })
            .expect("table")
    };
    let one = read(&first, &mut builds);
    let two = read(&first, &mut builds);
    assert!(Arc::ptr_eq(&one, &two));
    assert_eq!(builds, 1);

    // Equal bytes in another allocation are another generation.
    let copy = Arc::new(first.as_ref().clone());
    let three = read(&copy, &mut builds);
    assert!(!Arc::ptr_eq(&one, &three));
    assert_eq!(*one, *three);
    assert_eq!(builds, 2);
    drop((one, two));
    let diagnostics = cache.diagnostics();
    assert_eq!(
        (diagnostics.hits, diagnostics.misses, diagnostics.builds),
        (1, 2, 2)
    );
    assert_eq!(diagnostics.pinned_evictions, 1);
    assert!(diagnostics.retained);
    assert_eq!(diagnostics.retained_weight, three.weight());
}

#[test]
fn cache_key_does_not_keep_the_payload_alive() {
    let cache = FormatCache::new();
    let payload = Arc::new(complete_sheet().into_bytes());
    let weak = Arc::downgrade(&payload);
    let table = cache
        .get_or_build(FormatKey::allocation(&payload), || {
            FormatTable::parse(&payload)
        })
        .expect("table");
    drop(payload);
    assert_eq!(
        weak.strong_count(),
        0,
        "the cache pins only the parsed table"
    );
    assert!(table.len() > 0);
}

#[test]
fn cache_trims_only_clean_tables_and_never_caches_failures() {
    let cache = FormatCache::new();
    let version = litchi_core::SourceVersion::new(7, 1);
    let failed = cache.get_or_build(FormatKey::source(version), || {
        FormatTable::parse(b"<styleSheet/>")
    });
    assert!(failed.is_err());
    assert!(!cache.diagnostics().retained);

    let pinned = cache
        .get_or_build(FormatKey::source(version), || {
            FormatTable::parse(complete_sheet().as_bytes())
        })
        .expect("table");
    let handle = Format::new(Arc::clone(&pinned), 1).expect("format");
    drop(pinned);
    assert!(cache.diagnostics().pinned);
    assert!(!cache.trim_clean(), "a pinned table stays retained");
    drop(handle);
    assert!(cache.trim_clean());
    let diagnostics = cache.diagnostics();
    assert!(!diagnostics.retained);
    assert_eq!(diagnostics.clean_evictions, 1);
    assert!(!cache.trim_clean());

    // Another revision of the same source identity is another generation.
    let other = litchi_core::SourceVersion::new(7, 2);
    cache
        .get_or_build(FormatKey::source(version), || {
            FormatTable::parse(complete_sheet().as_bytes())
        })
        .expect("table");
    let mut rebuilt = false;
    cache
        .get_or_build(FormatKey::source(other), || {
            rebuilt = true;
            FormatTable::parse(complete_sheet().as_bytes())
        })
        .expect("table");
    assert!(rebuilt);
}

#[test]
fn cache_returns_oversized_tables_without_retaining_them() {
    let cache = FormatCache::new();
    let mut xml = format!(r#"<styleSheet xmlns="{S}"><numFmts>"#);
    // One declared code of 64 KiB per record pushes the table past the
    // ceiling with a few hundred records.
    let code = "0".repeat(64 * 1024);
    let count = MAX_RETAINED_WEIGHT / code.len() + 1;
    for id in 0..count {
        xml.push_str(&format!(
            r#"<numFmt numFmtId="{}" formatCode="{code}"/>"#,
            164 + id
        ));
    }
    xml.push_str("</numFmts><cellXfs><xf/></cellXfs></styleSheet>");
    let version = litchi_core::SourceVersion::new(9, 0);
    let mut builds = 0;
    for _ in 0..2 {
        let table = cache
            .get_or_build(FormatKey::source(version), || {
                builds += 1;
                FormatTable::parse(xml.as_bytes())
            })
            .expect("oversized table");
        assert!(table.weight() > MAX_RETAINED_WEIGHT);
    }
    assert_eq!(builds, 2, "an oversized table is parsed again");
    let diagnostics = cache.diagnostics();
    assert_eq!(diagnostics.oversized_bypasses, 2);
    assert!(!diagnostics.retained);
}

#[test]
fn every_repository_stylesheet_reads_with_the_catalog_count() {
    let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data");
    let mut files = Vec::new();
    collect_workbooks(std::path::Path::new(root), &mut files);
    assert!(!files.is_empty(), "no XLSX fixtures under {root}");
    let mut read = 0usize;
    let mut failures = Vec::new();
    for path in files {
        let Ok(bytes) = std::fs::read(&path) else {
            continue;
        };
        let Ok(package) = litchi_opc::OpcPackage::from_bytes(&bytes, SourceLimits::default())
        else {
            continue;
        };
        let Ok(workbook) = package.main_document_part() else {
            continue;
        };
        let Some(relationship) = workbook.rels().iter().find(|relationship| {
            matches!(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::STYLES
                    | litchi_opc::constants::relationship_type::STRICT_STYLES
            )
        }) else {
            continue;
        };
        let Ok(uri) = relationship.target_partname() else {
            continue;
        };
        let Ok(part) = package.get_part(&uri) else {
            continue;
        };
        let Ok(catalog) = crate::raw::styles::parse(part.blob()) else {
            continue;
        };
        match FormatTable::parse(part.blob()) {
            Ok(table) if u32::try_from(table.len()) == Ok(catalog.len()) => {
                for index in 0..catalog.len() {
                    let format = Format::new(
                        Arc::new(FormatTable::parse(part.blob()).expect("reparse")),
                        index,
                    )
                    .expect("every catalog index has a format");
                    let _ = format.number_format();
                }
                read += 1;
            },
            Ok(table) => failures.push(format!(
                "{}: {} formats, catalog {}",
                path.display(),
                table.len(),
                catalog.len()
            )),
            Err(error) => failures.push(format!("{}: {error}", path.display())),
        }
    }
    assert!(failures.is_empty(), "{failures:#?}");
    assert!(read > 100, "only {read} stylesheets were read");
}

fn collect_workbooks(directory: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_workbooks(&path, files);
        } else if path
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| {
                matches!(
                    value.to_ascii_lowercase().as_str(),
                    "xlsx" | "xlsm" | "xltx" | "xltm"
                )
            })
        {
            files.push(path);
        }
    }
}

#[test]
fn alignment_justification_retains_false_and_absence() {
    let xml = format!(
        r#"<styleSheet xmlns="{S}"><cellXfs count="3"><xf><alignment justifyLastLine="0" relativeIndent="0"/></xf><xf><alignment justifyLastLine="1"/></xf><xf><alignment/></xf></cellXfs></styleSheet>"#
    );
    let table = parse_table(&xml);
    let explicit_false = format(&table, 0);
    let alignment = explicit_false.alignment().unwrap();
    assert_eq!(alignment.justify_last_line, Some(false));
    assert_eq!(alignment.relative_indent, Some(0));
    assert_eq!(
        format(&table, 1).alignment().unwrap().justify_last_line,
        Some(true)
    );
    assert_eq!(
        format(&table, 2).alignment().unwrap().justify_last_line,
        None
    );
}
