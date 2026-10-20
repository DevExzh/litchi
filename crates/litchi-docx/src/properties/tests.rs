//! Tests for the direct-property readers.

use super::value::{
    Sign, Unit, hex_color, integer, measure, number, on_off, percent_fiftieths, text_scale,
};
use super::*;
use crate::color::Theme;
use crate::format::UnderlineStyle;
use crate::paragraph::{LineSpacingRule, RunUnderlineColor};
use crate::table::VMergeState;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

fn run(inner: &str) -> crate::Run {
    crate::Run::new(format!(r#"<w:r xmlns:w="{W}">{inner}</w:r>"#).into_bytes())
}

fn paragraph(inner: &str) -> crate::Paragraph {
    crate::Paragraph::new(format!(r#"<w:p xmlns:w="{W}">{inner}</w:p>"#).into_bytes())
}

fn invalid<'a>(names: impl Iterator<Item = &'a str>) -> Vec<&'a str> {
    names.collect()
}

#[test]
fn on_off_accepts_exactly_the_schema_tokens() {
    for (raw, value) in [
        ("true", true),
        ("1", true),
        ("on", true),
        (" on ", true),
        ("false", false),
        ("0", false),
        ("off", false),
    ] {
        assert_eq!(on_off(raw), Some(value), "{raw}");
    }
    for raw in ["TRUE", "yes", "", "2", "onn"] {
        assert_eq!(on_off(raw), None, "{raw}");
    }
}

#[test]
fn integers_follow_xsd_integer() {
    assert_eq!(integer("42"), Some(42));
    assert_eq!(integer("+42"), Some(42));
    assert_eq!(integer("-42"), Some(-42));
    assert_eq!(integer("007"), Some(7));
    assert_eq!(integer(" 5 "), Some(5));
    for raw in ["", "-", "4.0", "1e3", "0x10", "99999999999999999999"] {
        assert_eq!(integer(raw), None, "{raw}");
    }
}

#[test]
fn measures_convert_universal_units_exactly_and_round_half_away_from_zero() {
    let twips = |raw| measure(raw, Unit::Twips, Sign::Signed);
    assert_eq!(twips("720"), Some(720));
    assert_eq!(twips("-720"), Some(-720));
    assert_eq!(twips("+720"), Some(720));
    assert_eq!(twips(" 720 "), Some(720));
    assert_eq!(twips("720.5"), Some(721));
    assert_eq!(twips("-720.5"), Some(-721));
    assert_eq!(twips("720.4"), Some(720));
    assert_eq!(twips("1in"), Some(1440));
    assert_eq!(twips("2.54cm"), Some(1440));
    assert_eq!(twips("25.4mm"), Some(1440));
    assert_eq!(twips("12pt"), Some(240));
    assert_eq!(twips("1pc"), Some(240));
    assert_eq!(twips("1pi"), Some(240));
    assert_eq!(twips("1mm"), Some(57));
    assert_eq!(twips("-1in"), Some(-1440));
    for raw in [
        "", "in", "+1in", "1.in", ".5in", "1 in", "1e3", "1inch", "12px",
    ] {
        assert_eq!(twips(raw), None, "{raw}");
    }
    assert_eq!(measure("-1", Unit::Twips, Sign::Unsigned), None);
    assert_eq!(measure("-0", Unit::Twips, Sign::Unsigned), Some(0));
    assert_eq!(measure("12pt", Unit::HalfPoints, Sign::Unsigned), Some(24));
    assert_eq!(measure("0.25pt", Unit::HalfPoints, Sign::Unsigned), Some(1));
    assert_eq!(
        measure("0.75pt", Unit::EighthPoints, Sign::Unsigned),
        Some(6)
    );
    assert_eq!(measure("1in", Unit::Points, Sign::Unsigned), Some(72));
    assert_eq!(
        measure("1234567890123456789012345678901", Unit::Twips, Sign::Signed),
        None
    );
    assert_eq!(
        measure("9223372036854775808", Unit::Twips, Sign::Signed),
        None
    );
    assert_eq!(
        measure("9223372036854775807", Unit::Twips, Sign::Signed),
        Some(i64::MAX)
    );
}

#[test]
fn percentages_and_plain_numbers() {
    assert_eq!(percent_fiftieths("50%"), Some(2500));
    assert_eq!(percent_fiftieths("100%"), Some(5000));
    assert_eq!(percent_fiftieths("33.33%"), Some(1667));
    assert_eq!(percent_fiftieths("-10%"), Some(-500));
    assert_eq!(percent_fiftieths("50"), None);
    assert_eq!(percent_fiftieths("+50%"), None);
    assert_eq!(number("5000"), Some(5000));
    assert_eq!(number("9440.0"), Some(9440));
    assert_eq!(number("1in"), None);
}

#[test]
fn text_scale_accepts_both_spellings_up_to_six_hundred() {
    assert_eq!(text_scale("100"), Some(100));
    assert_eq!(text_scale("100%"), Some(100));
    assert_eq!(text_scale("0600%"), Some(600));
    assert_eq!(text_scale("000%"), Some(0));
    assert_eq!(text_scale("601"), None);
    assert_eq!(text_scale("-1"), None);
    assert_eq!(text_scale("%"), None);
}

#[test]
fn hex_colors() {
    assert_eq!(hex_color("auto"), Some(HexColor::Auto));
    assert_eq!(hex_color("FF0080"), Some(HexColor::Rgb([0xFF, 0x00, 0x80])));
    assert_eq!(hex_color("ff0080"), Some(HexColor::Rgb([0xFF, 0x00, 0x80])));
    for raw in ["FFF", "#FF0000", "GG0000", "Auto", "FF00000"] {
        assert_eq!(hex_color(raw), None, "{raw}");
    }
}

#[test]
fn vocabularies_round_trip_their_tokens() {
    fn check<T: Copy + PartialEq + std::fmt::Debug>(
        all: &[T],
        as_str: fn(T) -> &'static str,
        from: fn(&str) -> Option<T>,
    ) {
        for value in all {
            assert_eq!(from(as_str(*value)), Some(*value));
        }
    }
    check(Highlight::ALL, Highlight::as_str, Highlight::from_xml);
    check(FontHint::ALL, FontHint::as_str, FontHint::from_xml);
    check(ThemeFont::ALL, ThemeFont::as_str, ThemeFont::from_xml);
    check(
        VerticalAlign::ALL,
        VerticalAlign::as_str,
        VerticalAlign::from_xml,
    );
    check(Emphasis::ALL, Emphasis::as_str, Emphasis::from_xml);
    check(TextEffect::ALL, TextEffect::as_str, TextEffect::from_xml);
    check(
        ShadingPattern::ALL,
        ShadingPattern::as_str,
        ShadingPattern::from_xml,
    );
    check(
        Justification::ALL,
        Justification::as_str,
        Justification::from_xml,
    );
    check(
        TextAlignment::ALL,
        TextAlignment::as_str,
        TextAlignment::from_xml,
    );
    check(
        TextDirection::ALL,
        TextDirection::as_str,
        TextDirection::from_xml,
    );
    check(
        TextboxTightWrap::ALL,
        TextboxTightWrap::as_str,
        TextboxTightWrap::from_xml,
    );
    check(TabKind::ALL, TabKind::as_str, TabKind::from_xml);
    check(TabLeader::ALL, TabLeader::as_str, TabLeader::from_xml);
    check(
        TableJustification::ALL,
        TableJustification::as_str,
        TableJustification::from_xml,
    );
    check(TableLayout::ALL, TableLayout::as_str, TableLayout::from_xml);
    check(VerticalJc::ALL, VerticalJc::as_str, VerticalJc::from_xml);
    check(Merge::ALL, Merge::as_str, Merge::from_xml);
    check(HeightRule::ALL, HeightRule::as_str, HeightRule::from_xml);
    assert_eq!(Highlight::ALL.len(), 17);
    assert_eq!(ShadingPattern::ALL.len(), 38);
    assert_eq!(Justification::from_xml("Center"), None);
}

#[test]
fn border_styles_cover_the_schema_vocabulary() {
    let mut count = 0;
    for token in [
        "nil",
        "none",
        "single",
        "thick",
        "double",
        "dotted",
        "dashed",
        "dotDash",
        "dotDotDash",
        "triple",
        "thinThickSmallGap",
        "wave",
        "doubleWave",
        "dashSmallGap",
        "dashDotStroked",
        "threeDEmboss",
        "threeDEngrave",
        "outset",
        "inset",
        "apples",
        "zigZagStitch",
        "custom",
    ] {
        let style = BorderStyle::from_xml(token).unwrap_or_else(|| panic!("{token}"));
        assert_eq!(style.as_str(), token);
        count += 1;
    }
    assert_eq!(count, 22);
    assert!(BorderStyle::from_xml("nil").is_some_and(BorderStyle::is_none));
    assert!(BorderStyle::from_xml("none").is_some_and(BorderStyle::is_none));
    assert!(!BorderStyle::from_xml("single").is_some_and(BorderStyle::is_none));
    for token in ["", "Single", "solid", "zzz"] {
        assert_eq!(BorderStyle::from_xml(token), None, "{token}");
    }
}

#[test]
fn run_properties_are_typed() {
    let properties = run(concat!(
        r#"<w:rPr><w:rStyle w:val="Strong"/>"#,
        r#"<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun" w:cs="Mangal" w:hint="eastAsia" w:asciiTheme="minorHAnsi" w:cstheme="minorBidi"/>"#,
        r#"<w:b/><w:bCs w:val="0"/><w:i w:val="true"/><w:iCs w:val="off"/><w:caps/><w:smallCaps w:val="false"/>"#,
        r#"<w:strike/><w:dstrike w:val="0"/><w:outline/><w:shadow/><w:emboss/><w:imprint/><w:noProof/>"#,
        r#"<w:snapToGrid w:val="0"/><w:vanish/><w:webHidden/>"#,
        r#"<w:color w:val="1F4E79" w:themeColor="accent1" w:themeShade="BF"/>"#,
        r#"<w:spacing w:val="-20"/><w:w w:val="150"/><w:kern w:val="28"/><w:position w:val="-6"/>"#,
        r#"<w:sz w:val="28"/><w:szCs w:val="12pt"/><w:highlight w:val="darkCyan"/>"#,
        r#"<w:u w:val="double" w:color="FF0000" w:themeColor="accent2" w:themeTint="99"/>"#,
        r#"<w:effect w:val="shimmer"/><w:bdr w:val="single" w:sz="4" w:space="1" w:color="auto"/>"#,
        r#"<w:shd w:val="pct10" w:color="auto" w:fill="FFFF00" w:themeFill="accent6"/>"#,
        r#"<w:vertAlign w:val="superscript"/><w:rtl/><w:cs/><w:em w:val="dot"/>"#,
        r#"<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/><w:specVanish/><w:oMath/>"#,
        r#"</w:rPr><w:t>text</w:t>"#
    ))
    .direct_properties()
    .unwrap();

    assert_eq!(properties.style.as_deref(), Some("Strong"));
    let fonts = properties.fonts.as_ref().unwrap();
    assert_eq!(fonts.ascii.as_deref(), Some("Arial"));
    assert_eq!(fonts.h_ansi.as_deref(), Some("Arial"));
    assert_eq!(fonts.east_asia.as_deref(), Some("SimSun"));
    assert_eq!(fonts.complex_script.as_deref(), Some("Mangal"));
    assert_eq!(fonts.hint, Some(FontHint::EastAsia));
    assert_eq!(fonts.ascii_theme, Some(ThemeFont::MinorHAnsi));
    assert_eq!(fonts.h_ansi_theme, None);
    assert_eq!(fonts.complex_script_theme, Some(ThemeFont::MinorBidi));
    assert_eq!(properties.bold, Some(true));
    assert_eq!(properties.bold_complex_script, Some(false));
    assert_eq!(properties.italic, Some(true));
    assert_eq!(properties.italic_complex_script, Some(false));
    assert_eq!(properties.caps, Some(true));
    assert_eq!(properties.small_caps, Some(false));
    assert_eq!(properties.strike, Some(true));
    assert_eq!(properties.double_strike, Some(false));
    assert_eq!(properties.outline, Some(true));
    assert_eq!(properties.shadow, Some(true));
    assert_eq!(properties.emboss, Some(true));
    assert_eq!(properties.imprint, Some(true));
    assert_eq!(properties.no_proof, Some(true));
    assert_eq!(properties.snap_to_grid, Some(false));
    assert_eq!(properties.vanish, Some(true));
    assert_eq!(properties.web_hidden, Some(true));
    assert_eq!(
        properties.color,
        Some(Color {
            value: HexColor::Rgb([0x1F, 0x4E, 0x79]),
            theme: Some(Theme::Accent1),
            theme_tint: None,
            theme_shade: Some(0xBF),
        })
    );
    assert_eq!(properties.character_spacing, Some(Twips(-20)));
    assert_eq!(properties.scale, Some(150));
    assert_eq!(properties.kerning, Some(HalfPoints(28)));
    assert_eq!(properties.position, Some(HalfPoints(-6)));
    assert_eq!(properties.size, Some(HalfPoints(28)));
    assert_eq!(properties.size_complex_script, Some(HalfPoints(24)));
    assert_eq!(properties.highlight, Some(Highlight::DarkCyan));
    let underline = properties.underline.unwrap();
    assert_eq!(underline.style, UnderlineStyle::Double);
    assert_eq!(underline.color, Some(RunUnderlineColor::Rgb([0xFF, 0, 0])));
    assert_eq!(underline.theme_color, Some(Theme::Accent2));
    assert_eq!(underline.theme_tint, Some(0x99));
    assert_eq!(properties.effect, Some(TextEffect::Shimmer));
    let border = properties.border.unwrap();
    assert_eq!(border.style.as_str(), "single");
    assert_eq!(border.size, Some(EighthPoints(4)));
    assert_eq!(border.space, Some(Points(1)));
    assert_eq!(border.color, Some(HexColor::Auto));
    let shading = properties.shading.unwrap();
    assert_eq!(shading.pattern, ShadingPattern::Percent10);
    assert_eq!(shading.fill, Some(HexColor::Rgb([0xFF, 0xFF, 0])));
    assert_eq!(shading.theme_fill, Some(Theme::Accent6));
    assert_eq!(properties.vertical_align, Some(VerticalAlign::Superscript));
    assert_eq!(properties.right_to_left, Some(true));
    assert_eq!(properties.complex_script, Some(true));
    assert_eq!(properties.emphasis, Some(Emphasis::Dot));
    let language = properties.language.as_ref().unwrap();
    assert_eq!(language.latin.as_deref(), Some("en-US"));
    assert_eq!(language.east_asia.as_deref(), Some("zh-CN"));
    assert_eq!(language.bidi.as_deref(), Some("ar-SA"));
    assert_eq!(properties.special_vanish, Some(true));
    assert_eq!(properties.office_math, Some(true));
    assert_eq!(invalid(properties.invalid_properties()), Vec::<&str>::new());
}

#[test]
fn an_empty_or_missing_rpr_states_nothing() {
    assert_eq!(
        run("<w:t>x</w:t>").direct_properties().unwrap(),
        Run::default()
    );
    assert_eq!(
        run("<w:rPr/><w:t>x</w:t>").direct_properties().unwrap(),
        Run::default()
    );
    let empty = crate::Run::new(format!(r#"<w:r xmlns:w="{W}"/>"#).into_bytes());
    assert_eq!(empty.direct_properties().unwrap(), Run::default());
}

#[test]
fn run_properties_ignore_revision_snapshots_and_nested_runs() {
    // F-8: the previous formatting in w:rPrChange and the runs of a text box
    // or phonetic guide inside the run are not the run's own properties.
    let properties = run(concat!(
        r#"<w:rPr><w:i/><w:rPrChange w:id="1" w:author="A"><w:rPr><w:b/><w:sz w:val="40"/><w:highlight w:val="red"/></w:rPr></w:rPrChange></w:rPr>"#,
        r#"<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:txbxContent><w:p><w:r><w:rPr><w:strike/><w:caps/></w:rPr><w:t>box</w:t></w:r></w:p></w:txbxContent></wp:inline></w:drawing>"#,
        r#"<w:ruby><w:rt><w:r><w:rPr><w:vanish/></w:rPr><w:t>a</w:t></w:r></w:rt></w:ruby>"#
    ))
    .direct_properties()
    .unwrap();
    assert_eq!(properties.italic, Some(true));
    assert_eq!(properties.bold, None);
    assert_eq!(properties.size, None);
    assert_eq!(properties.highlight, None);
    assert_eq!(properties.strike, None);
    assert_eq!(properties.caps, None);
    assert_eq!(properties.vanish, None);

    let nested_only = run(
        r#"<w:pict><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>box</w:t></w:r></w:p></w:txbxContent></w:pict>"#,
    )
    .direct_properties()
    .unwrap();
    assert_eq!(nested_only, Run::default());
}

#[test]
fn run_properties_are_namespace_checked() {
    let properties = run(concat!(
        r#"<w:rPr xmlns:x="urn:other" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">"#,
        r#"<x:b/><x:sz x:val="40"/><w14:shadow/><w:i x:val="0"/><w:color w:val="FF0000" x:themeColor="accent1"/>"#,
        r#"</w:rPr>"#
    ))
    .direct_properties()
    .unwrap();
    assert_eq!(properties.bold, None);
    assert_eq!(properties.size, None);
    assert_eq!(properties.shadow, None);
    // A foreign `val` is not the property's value, so `<w:i>` has none.
    assert_eq!(properties.italic, Some(true));
    assert_eq!(properties.color.unwrap().theme, None);

    // A foreign container is not the run's properties.
    let foreign = crate::Run::new(
        format!(r#"<w:r xmlns:w="{W}" xmlns:x="urn:x"><x:rPr><w:b/></x:rPr></w:r>"#).into_bytes(),
    );
    assert_eq!(foreign.direct_properties().unwrap().bold, None);
}

#[test]
fn run_properties_resolve_strict_unprefixed_and_cut_fragments() {
    let strict = crate::Run::new(
        format!(r#"<s:r xmlns:s="{STRICT}"><s:rPr><s:b/><s:sz s:val="20"/></s:rPr></s:r>"#)
            .into_bytes(),
    );
    let properties = strict.direct_properties().unwrap();
    assert_eq!(properties.bold, Some(true));
    assert_eq!(properties.size, Some(HalfPoints(20)));

    let default_namespace = crate::Run::new(
        format!(r#"<r xmlns="{W}"><rPr><b/><sz val="20"/></rPr></r>"#).into_bytes(),
    );
    // Unprefixed attributes are in no namespace, so `val` is not Word's.
    let properties = default_namespace.direct_properties().unwrap();
    assert_eq!(properties.bold, Some(true));
    assert_eq!(properties.size, None);
    assert_eq!(invalid(properties.invalid_properties()), ["sz"]);

    // A fragment cut from a part: the `w` prefix is declared on an ancestor.
    let cut = crate::Run::new(br#"<w:r><w:rPr><w:b w:val="0"/><x:i/></w:rPr></w:r>"#.to_vec());
    let properties = cut.direct_properties().unwrap();
    assert_eq!(properties.bold, Some(false));
    assert_eq!(properties.italic, None);

    let unprefixed_cut = crate::Run::new(br#"<r><rPr><b/><sz val="20"/></rPr></r>"#.to_vec());
    let properties = unprefixed_cut.direct_properties().unwrap();
    assert_eq!(properties.bold, Some(true));
    assert_eq!(properties.size, Some(HalfPoints(20)));
}

#[test]
fn invalid_run_properties_are_listed_without_hiding_the_others() {
    let properties = run(concat!(
        r##"<w:rPr><w:b w:val="yes"/><w:i/><w:highlight w:val="purple"/><w:sz/><w:color w:val="#FF0000"/>"##,
        r#"<w:u w:val="squiggly"/><w:rFonts w:ascii="Arial" w:hint="latin"/><w:w w:val="700"/><w:kern w:val="-2"/>"#,
        r#"<w:rStyle/><w:shd w:val="pct11"/><w:bdr w:val="single" w:sz="thin"/></w:rPr>"#
    ))
    .direct_properties()
    .unwrap();
    assert_eq!(properties.bold, None);
    assert_eq!(properties.italic, Some(true));
    assert_eq!(properties.highlight, None);
    assert_eq!(properties.size, None);
    assert_eq!(properties.color, None);
    assert_eq!(properties.underline, None);
    assert_eq!(properties.scale, None);
    assert_eq!(properties.kerning, None);
    assert_eq!(properties.style, None);
    assert_eq!(properties.shading, None);
    assert_eq!(properties.border, None);
    // Attributes of w:rFonts are read one by one.
    let fonts = properties.fonts.as_ref().unwrap();
    assert_eq!(fonts.ascii.as_deref(), Some("Arial"));
    assert_eq!(fonts.hint, None);
    assert_eq!(
        invalid(properties.invalid_properties()),
        [
            "rStyle",
            "rFonts",
            "b",
            "color",
            "w",
            "kern",
            "sz",
            "highlight",
            "u",
            "bdr",
            "shd"
        ]
    );
}

#[test]
fn repeated_run_properties_read_their_first_occurrence() {
    let properties = run(concat!(
        r#"<w:rPr><w:b/><w:b w:val="0"/><w:sz w:val="20"/><w:sz w:val="99"/>"#,
        r#"<w:rFonts w:ascii="First"/><w:rFonts w:ascii="Second" w:hAnsi="Merged"/>"#,
        r#"<w:lang w:val="en-US"/><w:lang w:val="fr-FR" w:bidi="he-IL"/></w:rPr>"#
    ))
    .direct_properties()
    .unwrap();
    assert_eq!(properties.bold, Some(true));
    assert_eq!(properties.size, Some(HalfPoints(20)));
    let fonts = properties.fonts.as_ref().unwrap();
    assert_eq!(fonts.ascii.as_deref(), Some("First"));
    assert_eq!(fonts.h_ansi.as_deref(), Some("Merged"));
    let language = properties.language.as_ref().unwrap();
    assert_eq!(language.latin.as_deref(), Some("en-US"));
    assert_eq!(language.bidi.as_deref(), Some("he-IL"));
    assert_eq!(properties.invalid_properties().count(), 0);
}

#[test]
fn run_property_defaults_follow_the_schema() {
    let properties = run(r#"<w:rPr><w:u/><w:w/><w:vertAlign w:val="baseline"/></w:rPr>"#)
        .direct_properties()
        .unwrap();
    assert_eq!(properties.underline.unwrap().style, UnderlineStyle::Single);
    assert_eq!(properties.scale, Some(100));
    assert_eq!(properties.vertical_align, Some(VerticalAlign::Baseline));
}

#[test]
fn run_structure_errors_are_refused() {
    let not_a_run = crate::Run::new(format!(r#"<w:p xmlns:w="{W}"/>"#).into_bytes());
    assert!(not_a_run.direct_properties().is_err());
    let foreign_run = crate::Run::new(br#"<x:r xmlns:x="urn:x"><x:rPr/></x:r>"#.to_vec());
    assert!(foreign_run.direct_properties().is_err());
    assert!(run("<w:rPr/><w:rPr/>").direct_properties().is_err());
    assert!(run("<w:rPr><w:b></w:rPr>").direct_properties().is_err());
    assert!(
        run(r#"<w:rPr><w:b w:val="1" w:val="0"/></w:rPr>"#)
            .direct_properties()
            .is_err()
    );
    assert!(
        crate::Run::new(format!(r#"<w:r xmlns:w="{W}"><w:rPr>"#).into_bytes())
            .direct_properties()
            .is_err()
    );
    assert!(
        crate::Run::new(format!(r#"<w:r xmlns:w="{W}"/><w:r xmlns:w="{W}"/>"#).into_bytes())
            .direct_properties()
            .is_err()
    );
    assert!(crate::Run::new(Vec::new()).direct_properties().is_err());
}

#[test]
fn paragraph_properties_are_typed() {
    let properties = paragraph(concat!(
        r#"<w:pPr><w:pStyle w:val="ListParagraph"/><w:keepNext/><w:keepLines w:val="0"/><w:pageBreakBefore/>"#,
        r#"<w:widowControl w:val="0"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr>"#,
        r#"<w:suppressLineNumbers/><w:pBdr><w:top w:val="single" w:sz="6"/><w:between w:val="dotted"/></w:pBdr>"#,
        r#"<w:shd w:val="clear" w:fill="EEEEEE"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/><w:tab w:val="clear" w:pos="720"/></w:tabs>"#,
        r#"<w:suppressAutoHyphens/><w:kinsoku w:val="0"/><w:wordWrap/><w:overflowPunct/><w:topLinePunct/>"#,
        r#"<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:bidi/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>"#,
        r#"<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto" w:beforeAutospacing="1"/>"#,
        r#"<w:ind w:start="720" w:hanging="360" w:firstLineChars="200"/>"#,
        r#"<w:contextualSpacing/><w:mirrorIndents/><w:suppressOverlap/><w:jc w:val="both"/>"#,
        r#"<w:textDirection w:val="tbRl"/><w:textAlignment w:val="center"/><w:textboxTightWrap w:val="allLines"/>"#,
        r#"<w:outlineLvl w:val="2"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:pPr>"#,
        r#"<w:r><w:t>text</w:t></w:r>"#
    ))
    .direct_properties()
    .unwrap();

    assert_eq!(properties.style.as_deref(), Some("ListParagraph"));
    assert_eq!(properties.keep_next, Some(true));
    assert_eq!(properties.keep_lines, Some(false));
    assert_eq!(properties.page_break_before, Some(true));
    assert_eq!(properties.widow_control, Some(false));
    let numbering = properties.numbering.unwrap();
    assert_eq!((numbering.id, numbering.level), (Some(7), Some(1)));
    assert_eq!(properties.suppress_line_numbers, Some(true));
    let borders = properties.borders.unwrap();
    assert_eq!(borders.top.unwrap().size, Some(EighthPoints(6)));
    assert_eq!(borders.between.unwrap().style.as_str(), "dotted");
    assert_eq!(borders.left, None);
    assert_eq!(
        properties.shading.unwrap().fill,
        Some(HexColor::Rgb([0xEE, 0xEE, 0xEE]))
    );
    let tabs = properties.tabs.as_ref().unwrap();
    assert_eq!(tabs.len(), 2);
    assert_eq!(tabs[0].kind, TabKind::Right);
    assert_eq!(tabs[0].leader, Some(TabLeader::Dot));
    assert_eq!(tabs[0].position, Twips(9350));
    assert_eq!(tabs[1].kind, TabKind::Clear);
    assert_eq!(tabs[1].leader, None);
    assert_eq!(properties.suppress_auto_hyphens, Some(true));
    assert_eq!(properties.kinsoku, Some(false));
    assert_eq!(properties.word_wrap, Some(true));
    assert_eq!(properties.overflow_punctuation, Some(true));
    assert_eq!(properties.top_line_punctuation, Some(true));
    assert_eq!(properties.auto_space_east_asian_latin, Some(false));
    assert_eq!(properties.auto_space_east_asian_numbers, Some(false));
    assert_eq!(properties.bidi, Some(true));
    assert_eq!(properties.adjust_right_indent, Some(false));
    assert_eq!(properties.snap_to_grid, Some(false));
    let spacing = properties.spacing.unwrap();
    assert_eq!(spacing.before, Some(240));
    assert_eq!(spacing.after, Some(120));
    assert_eq!(spacing.line, Some(360));
    assert_eq!(spacing.line_rule, Some(LineSpacingRule::Auto));
    assert_eq!(spacing.before_auto_spacing, Some(true));
    let indentation = properties.indentation.unwrap();
    assert_eq!(indentation.start, Some(Twips(720)));
    assert_eq!(indentation.leading(), Some(Twips(720)));
    assert_eq!(indentation.trailing(), None);
    assert_eq!(indentation.hanging, Some(Twips(360)));
    assert_eq!(indentation.first_line_chars, Some(200));
    assert_eq!(properties.contextual_spacing, Some(true));
    assert_eq!(properties.mirror_indents, Some(true));
    assert_eq!(properties.suppress_overlap, Some(true));
    assert_eq!(properties.justification, Some(Justification::Both));
    assert_eq!(properties.text_direction, Some(TextDirection::TbRl));
    assert_eq!(properties.text_alignment, Some(TextAlignment::Center));
    assert_eq!(
        properties.textbox_tight_wrap,
        Some(TextboxTightWrap::AllLines)
    );
    assert_eq!(properties.outline_level, Some(2));
    let mark = properties.mark.as_deref().unwrap();
    assert_eq!(mark.bold, Some(true));
    assert_eq!(mark.size, Some(HalfPoints(32)));
    assert_eq!(properties.invalid_properties().count(), 0);
}

#[test]
fn paragraph_properties_ignore_snapshots_section_breaks_and_content() {
    let properties = paragraph(concat!(
        r#"<w:pPr><w:jc w:val="center"/>"#,
        r#"<w:rPr><w:rPrChange w:id="2" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>"#,
        r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:space="720"/></w:sectPr>"#,
        r#"<w:pPrChange w:id="1" w:author="A"><w:pPr><w:jc w:val="right"/><w:keepNext/><w:ind w:left="1440"/></w:pPr></w:pPrChange></w:pPr>"#,
        r#"<w:r><w:pict><w:txbxContent><w:p><w:pPr><w:keepLines/></w:pPr></w:p></w:txbxContent></w:pict></w:r>"#
    ))
    .direct_properties()
    .unwrap();
    assert_eq!(properties.justification, Some(Justification::Center));
    assert_eq!(properties.keep_next, None);
    assert_eq!(properties.keep_lines, None);
    assert_eq!(properties.indentation, None);
    assert_eq!(properties.mark.as_deref(), Some(&Run::default()));
}

#[test]
fn paragraph_numbering_and_indentation_edge_cases() {
    let properties = paragraph(r#"<w:pPr><w:numPr/></w:pPr>"#)
        .direct_properties()
        .unwrap();
    assert_eq!(properties.numbering, Some(Numbering::default()));

    let properties = paragraph(r#"<w:pPr><w:numPr><w:ilvl w:val="12"/></w:numPr></w:pPr>"#)
        .direct_properties()
        .unwrap();
    // Levels are reported as authored; list resolution clamps them.
    assert_eq!(properties.numbering.unwrap().level, Some(12));
    assert_eq!(properties.numbering.unwrap().id, None);

    // Word merges repeated w:ind elements attribute by attribute.
    let properties = paragraph(
        r#"<w:pPr><w:ind w:left="360"/><w:ind w:left="9999" w:firstLine="360"/></w:pPr>"#,
    )
    .direct_properties()
    .unwrap();
    let indentation = properties.indentation.unwrap();
    assert_eq!(indentation.left, Some(Twips(360)));
    assert_eq!(indentation.first_line, Some(Twips(360)));
    assert_eq!(indentation.leading(), Some(Twips(360)));

    let properties = paragraph(r#"<w:pPr><w:ind w:right="1in" w:end="10"/></w:pPr>"#)
        .direct_properties()
        .unwrap();
    assert_eq!(
        properties.indentation.unwrap().trailing(),
        Some(Twips(1440))
    );
}

#[test]
fn invalid_paragraph_properties_are_listed() {
    let properties = paragraph(concat!(
        r#"<w:pPr><w:jc w:val="middle"/><w:outlineLvl w:val="10"/><w:keepNext w:val="always"/>"#,
        r#"<w:numPr><w:numId w:val="-1"/><w:ilvl w:val="2"/></w:numPr><w:ind w:hanging="-360" w:left="720"/>"#,
        r#"<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="sideways" w:pos="1440"/><w:tab w:val="left"/></w:tabs>"#,
        r#"<w:pBdr><w:top w:val="hairline"/><w:bottom w:val="single"/></w:pBdr><w:spacing w:before="x"/>"#,
        r#"<w:rPr><w:highlight w:val="neon"/></w:rPr><w:pStyle/></w:pPr>"#
    ))
    .direct_properties()
    .unwrap();
    assert_eq!(properties.justification, None);
    assert_eq!(properties.outline_level, None);
    assert_eq!(properties.keep_next, None);
    let numbering = properties.numbering.unwrap();
    assert_eq!((numbering.id, numbering.level), (None, Some(2)));
    let indentation = properties.indentation.unwrap();
    assert_eq!(indentation.hanging, None);
    assert_eq!(indentation.left, Some(Twips(720)));
    assert_eq!(properties.tabs.as_deref().map(<[TabStop]>::len), Some(1));
    let borders = properties.borders.unwrap();
    assert_eq!(borders.top, None);
    assert!(borders.bottom.is_some());
    assert_eq!(properties.spacing, None);
    assert_eq!(properties.style, None);
    assert_eq!(
        invalid(properties.invalid_properties()),
        [
            "pStyle",
            "keepNext",
            "numPr",
            "pBdr",
            "tabs",
            "spacing",
            "ind",
            "jc",
            "outlineLvl"
        ]
    );
    let mark = properties.mark.as_deref().unwrap();
    assert_eq!(invalid(mark.invalid_properties()), ["highlight"]);
}

#[test]
fn repeated_paragraph_containers_read_their_first_occurrence() {
    let properties = paragraph(concat!(
        r#"<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:numPr><w:numId w:val="2"/><w:ilvl w:val="3"/></w:numPr>"#,
        r#"<w:rPr><w:b/></w:rPr><w:rPr><w:i/></w:rPr><w:tabs/><w:tabs><w:tab w:val="left" w:pos="1"/></w:tabs></w:pPr>"#
    ))
    .direct_properties()
    .unwrap();
    let numbering = properties.numbering.unwrap();
    assert_eq!((numbering.id, numbering.level), (Some(1), None));
    let mark = properties.mark.as_deref().unwrap();
    assert_eq!((mark.bold, mark.italic), (Some(true), None));
    assert_eq!(properties.tabs, Some(Vec::new()));
}

#[test]
fn paragraph_structure_errors_are_refused() {
    let not_a_paragraph = crate::Paragraph::new(format!(r#"<w:r xmlns:w="{W}"/>"#).into_bytes());
    assert!(not_a_paragraph.direct_properties().is_err());
    assert!(paragraph("<w:pPr/><w:pPr/>").direct_properties().is_err());
    assert!(
        paragraph("<w:pPr><w:jc></w:pPr>")
            .direct_properties()
            .is_err()
    );
    let no_properties = paragraph("<w:r><w:t>x</w:t></w:r>")
        .direct_properties()
        .unwrap();
    assert_eq!(no_properties, Paragraph::default());
}

fn table(inner: &str) -> crate::Table {
    crate::Table::new(format!(r#"<w:tbl xmlns:w="{W}">{inner}</w:tbl>"#).into_bytes())
}

#[test]
fn table_properties_are_typed() {
    let properties = table(concat!(
        r#"<w:tblPr><w:tblStyle w:val="TableGrid"/><w:bidiVisual/><w:tblStyleRowBandSize w:val="2"/>"#,
        r#"<w:tblStyleColBandSize w:val="1"/><w:tblW w:w="50%" w:type="pct"/><w:jc w:val="center"/>"#,
        r#"<w:tblCellSpacing w:w="20" w:type="dxa"/><w:tblInd w:w="-108" w:type="dxa"/>"#,
        r#"<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:start w:val="double"/><w:left w:val="dashed"/><w:insideH w:val="nil"/></w:tblBorders>"#,
        r#"<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/><w:tblLayout w:type="fixed"/>"#,
        r#"<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:end w:w="5%" w:type="pct"/></w:tblCellMar>"#,
        r#"<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>"#,
        r#"<w:tblCaption w:val="Results"/><w:tblDescription w:val="Quarterly results"/></w:tblPr>"#,
        r#"<w:tblGrid><w:gridCol w:w="4675"/><w:gridCol/><w:gridCol w:w="2in"/></w:tblGrid>"#,
        r#"<w:tr><w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>"#
    ))
    .properties()
    .unwrap();
    assert_eq!(properties.style.as_deref(), Some("TableGrid"));
    assert_eq!(properties.bidi_visual, Some(true));
    assert_eq!(properties.row_band_size, Some(2));
    assert_eq!(properties.column_band_size, Some(1));
    assert_eq!(properties.width, Some(Width::Percent(2500)));
    assert_eq!(properties.justification, Some(TableJustification::Center));
    assert_eq!(properties.cell_spacing, Some(Width::Twips(Twips(20))));
    assert_eq!(properties.indent, Some(Width::Twips(Twips(-108))));
    let borders = properties.borders.unwrap();
    assert_eq!(borders.top.unwrap().size, Some(EighthPoints(4)));
    assert_eq!(borders.start.unwrap().style.as_str(), "double");
    assert_eq!(borders.left.unwrap().style.as_str(), "dashed");
    assert!(borders.inside_horizontal.unwrap().style.is_none());
    assert_eq!(borders.bottom, None);
    assert_eq!(
        properties.shading.unwrap().fill,
        Some(HexColor::Rgb([0xD9, 0xE2, 0xF3]))
    );
    assert_eq!(properties.layout, Some(TableLayout::Fixed));
    let margins = properties.cell_margins.unwrap();
    assert_eq!(margins.top, Some(Width::Twips(Twips(0))));
    assert_eq!(margins.left, Some(Width::Twips(Twips(108))));
    assert_eq!(margins.end, Some(Width::Percent(250)));
    let look = properties.look.unwrap();
    assert_eq!(look.value, Some(0x04A0));
    assert_eq!(look.mask(), 0x04A0);
    assert_eq!(look.first_row, Some(true));
    assert_eq!(look.no_vertical_band, Some(true));
    assert_eq!(properties.caption.as_deref(), Some("Results"));
    assert_eq!(properties.description.as_deref(), Some("Quarterly results"));
    assert_eq!(
        properties.grid,
        Some(vec![Some(Twips(4675)), None, Some(Twips(2880))])
    );
    assert_eq!(properties.invalid_properties().count(), 0);
}

#[test]
fn table_widths_follow_their_type() {
    for (attributes, expected) in [
        (r#"w:w="5000" w:type="pct""#, Some(Width::Percent(5000))),
        (r#"w:w="100%" w:type="pct""#, Some(Width::Percent(5000))),
        (r#"w:type="pct""#, Some(Width::Percent(0))),
        (
            r#"w:w="9440.0" w:type="dxa""#,
            Some(Width::Twips(Twips(9440))),
        ),
        (r#"w:w="2880""#, Some(Width::Twips(Twips(2880)))),
        (r#"w:w="1in" w:type="dxa""#, Some(Width::Twips(Twips(1440)))),
        (r#"w:w="0" w:type="auto""#, Some(Width::Auto)),
        (r#"w:type="auto""#, Some(Width::Auto)),
        (r#"w:type="nil""#, Some(Width::Nil)),
        (r#"w:w="50%" w:type="dxa""#, None),
        (r#"w:w="1in" w:type="pct""#, None),
        (r#"w:w="10" w:type="em""#, None),
    ] {
        let properties = table(&format!(r#"<w:tblPr><w:tblW {attributes}/></w:tblPr>"#))
            .properties()
            .unwrap();
        assert_eq!(properties.width, expected, "{attributes}");
        assert_eq!(
            properties.invalid_properties().count(),
            usize::from(expected.is_none()),
            "{attributes}"
        );
    }
}

#[test]
fn table_look_mask_is_assembled_from_flags() {
    let properties = table(
        r#"<w:tblPr><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="on" w:noHBand="true"/></w:tblPr>"#,
    )
    .properties()
    .unwrap();
    assert_eq!(properties.look.unwrap().mask(), 0x0020 | 0x0080 | 0x0200);
}

#[test]
fn table_properties_ignore_nested_tables_and_snapshots() {
    let properties = table(concat!(
        r#"<w:tblPr><w:tblStyle w:val="Outer"/><w:tblPrChange w:id="1" w:author="A"><w:tblPr><w:jc w:val="right"/></w:tblPr></w:tblPrChange></w:tblPr>"#,
        r#"<w:tblGrid><w:gridCol w:w="100"/><w:tblGridChange w:id="2"><w:tblGrid><w:gridCol w:w="7"/></w:tblGrid></w:tblGridChange></w:tblGrid>"#,
        r#"<w:tr><w:tblPrEx><w:jc w:val="left"/></w:tblPrEx><w:tc><w:tbl><w:tblPr><w:tblStyle w:val="Inner"/><w:bidiVisual/></w:tblPr>"#,
        r#"<w:tblGrid><w:gridCol w:w="50"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr>"#
    ))
    .properties()
    .unwrap();
    assert_eq!(properties.style.as_deref(), Some("Outer"));
    assert_eq!(properties.justification, None);
    assert_eq!(properties.bidi_visual, None);
    assert_eq!(properties.grid, Some(vec![Some(Twips(100))]));
}

#[test]
fn table_structure_errors_are_refused() {
    assert!(table("<w:tblPr/><w:tblPr/>").properties().is_err());
    assert!(table("<w:tblGrid/><w:tblGrid/>").properties().is_err());
    let not_a_table = crate::Table::new(format!(r#"<w:tc xmlns:w="{W}"/>"#).into_bytes());
    assert!(not_a_table.properties().is_err());
    assert_eq!(table("").properties().unwrap(), Table::default());
}

#[test]
fn row_properties_are_typed_and_scoped() {
    let row = crate::Row::new(
        format!(
            concat!(
                r#"<w:tr xmlns:w="{}"><w:tblPrEx><w:tblW w:w="1" w:type="dxa"/></w:tblPrEx>"#,
                r#"<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="2"/><w:wBefore w:w="720" w:type="dxa"/>"#,
                r#"<w:wAfter w:w="0" w:type="auto"/><w:cantSplit/><w:trHeight w:val="567" w:hRule="atLeast"/>"#,
                r#"<w:tblHeader w:val="0"/><w:tblCellSpacing w:w="15" w:type="dxa"/><w:jc w:val="end"/><w:hidden/>"#,
                r#"<w:trPrChange w:id="1" w:author="A"><w:trPr><w:cantSplit w:val="0"/></w:trPr></w:trPrChange></w:trPr>"#,
                r#"<w:tc><w:tbl><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr>"#
            ),
            W
        )
        .into_bytes(),
    );
    let properties = row.properties().unwrap();
    assert_eq!(properties.grid_before, Some(1));
    assert_eq!(properties.grid_after, Some(2));
    assert_eq!(properties.width_before, Some(Width::Twips(Twips(720))));
    assert_eq!(properties.width_after, Some(Width::Auto));
    assert_eq!(properties.cant_split, Some(true));
    let height = properties.height.unwrap();
    assert_eq!(height.value, Twips(567));
    assert_eq!(height.rule, Some(HeightRule::AtLeast));
    assert_eq!(properties.header, Some(false));
    assert_eq!(properties.cell_spacing, Some(Width::Twips(Twips(15))));
    assert_eq!(properties.justification, Some(TableJustification::End));
    assert_eq!(properties.hidden, Some(true));
    assert_eq!(properties.invalid_properties().count(), 0);

    let invalid_row = crate::Row::new(
        format!(r#"<w:tr xmlns:w="{W}"><w:trPr><w:trHeight w:hRule="tall"/><w:gridBefore w:val="-1"/></w:trPr></w:tr>"#)
            .into_bytes(),
    );
    let properties = invalid_row.properties().unwrap();
    assert_eq!(properties.height, None);
    assert_eq!(properties.grid_before, None);
    assert_eq!(
        invalid(properties.invalid_properties()),
        ["gridBefore", "trHeight"]
    );
}

fn cell(inner: &str) -> crate::Cell {
    crate::Cell::new(format!(r#"<w:tc xmlns:w="{W}">{inner}</w:tc>"#).into_bytes())
}

#[test]
fn cell_properties_are_typed() {
    let properties = cell(concat!(
        r#"<w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:gridSpan w:val="3"/><w:hMerge w:val="restart"/><w:vMerge/>"#,
        r#"<w:tcBorders><w:top w:val="single" w:sz="12" w:color="FF0000" w:themeColor="accent2"/><w:tl2br w:val="single"/><w:end w:val="nil"/></w:tcBorders>"#,
        r#"<w:shd w:val="solid" w:color="auto" w:fill="auto"/><w:noWrap/>"#,
        r#"<w:tcMar><w:top w:w="72" w:type="dxa"/><w:start w:w="0" w:type="nil"/><w:right w:w="115" w:type="dxa"/></w:tcMar>"#,
        r#"<w:textDirection w:val="btLr"/><w:tcFitText w:val="0"/><w:vAlign w:val="bottom"/><w:hideMark/></w:tcPr><w:p/>"#
    ))
    .properties()
    .unwrap();
    assert_eq!(properties.width, Some(Width::Twips(Twips(2880))));
    assert_eq!(properties.grid_span, Some(3));
    assert_eq!(properties.horizontal_merge, Some(Merge::Restart));
    assert_eq!(properties.vertical_merge, Some(Merge::Continue));
    let borders = properties.borders.unwrap();
    let top = borders.top.unwrap();
    assert_eq!(top.size, Some(EighthPoints(12)));
    assert_eq!(top.theme_color, Some(Theme::Accent2));
    assert!(borders.top_left_to_bottom_right.is_some());
    assert!(borders.end.is_some_and(|border| border.style.is_none()));
    assert_eq!(properties.shading.unwrap().pattern, ShadingPattern::Solid);
    assert_eq!(properties.no_wrap, Some(true));
    let margins = properties.margins.unwrap();
    assert_eq!(margins.top, Some(Width::Twips(Twips(72))));
    assert_eq!(margins.start, Some(Width::Nil));
    assert_eq!(margins.right, Some(Width::Twips(Twips(115))));
    assert_eq!(properties.text_direction, Some(TextDirection::BtLr));
    assert_eq!(properties.fit_text, Some(false));
    assert_eq!(properties.vertical_align, Some(VerticalJc::Bottom));
    assert_eq!(properties.hide_mark, Some(true));
    assert_eq!(properties.invalid_properties().count(), 0);
}

#[test]
fn cell_properties_read_only_the_cells_own_tcpr() {
    // F-7: a nested table's cell properties are not the outer cell's.
    let nested = concat!(
        r#"<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="4"/><w:vMerge w:val="restart"/><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr>"#,
        r#"<w:p/></w:tc></w:tr></w:tbl><w:p/>"#
    );
    let outer_without_properties = cell(nested);
    assert_eq!(outer_without_properties.grid_span().unwrap(), 1);
    assert_eq!(outer_without_properties.v_merge().unwrap(), None);
    assert_eq!(
        outer_without_properties.properties().unwrap(),
        Cell::default()
    );

    let outer_with_properties = cell(&format!(r#"<w:tcPr><w:tcW w:w="100"/></w:tcPr>{nested}"#));
    assert_eq!(outer_with_properties.grid_span().unwrap(), 1);
    assert_eq!(outer_with_properties.v_merge().unwrap(), None);
    let properties = outer_with_properties.properties().unwrap();
    assert_eq!(properties.grid_span, None);
    assert_eq!(properties.vertical_merge, None);
    assert_eq!(properties.shading, None);

    let own = cell(&format!(
        r#"<w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="continue"/></w:tcPr>{nested}"#
    ));
    assert_eq!(own.grid_span().unwrap(), 2);
    assert_eq!(own.v_merge().unwrap(), Some(VMergeState::Continue));
    let properties = own.properties().unwrap();
    assert_eq!(properties.grid_span, Some(2));
    assert_eq!(properties.vertical_merge, Some(Merge::Continue));
}

#[test]
fn cell_merge_readers_skip_revision_snapshots_and_keep_their_refusals() {
    let snapshot = cell(
        r#"<w:tcPr><w:tcPrChange w:id="1" w:author="A"><w:tcPr><w:gridSpan w:val="3"/><w:vMerge w:val="restart"/></w:tcPr></w:tcPrChange></w:tcPr><w:p/>"#,
    );
    assert_eq!(snapshot.grid_span().unwrap(), 1);
    assert_eq!(snapshot.v_merge().unwrap(), None);
    assert_eq!(snapshot.properties().unwrap(), Cell::default());

    let zero = cell(r#"<w:tcPr><w:gridSpan w:val="0"/></w:tcPr>"#);
    assert!(zero.grid_span().is_err());
    let properties = zero.properties().unwrap();
    assert_eq!(properties.grid_span, None);
    assert_eq!(invalid(properties.invalid_properties()), ["gridSpan"]);

    let sideways = cell(r#"<w:tcPr><w:vMerge w:val="sideways"/></w:tcPr>"#);
    assert!(sideways.v_merge().is_err());
    assert_eq!(
        invalid(sideways.properties().unwrap().invalid_properties()),
        ["vMerge"]
    );

    // A repeated property keeps its first value, as before.
    let repeated = cell(r#"<w:tcPr><w:gridSpan w:val="2"/><w:gridSpan w:val="5"/></w:tcPr>"#);
    assert_eq!(repeated.grid_span().unwrap(), 2);
    assert_eq!(repeated.properties().unwrap().grid_span, Some(2));

    // An attribute-free span is one column.
    assert_eq!(
        cell(r#"<w:tcPr><w:gridSpan/></w:tcPr>"#)
            .grid_span()
            .unwrap(),
        1
    );
    // Malformed XML anywhere in the cell is still refused.
    assert!(cell(r#"<w:tcPr/><w:p><w:r></w:p>"#).grid_span().is_err());
}

#[test]
fn cell_structure_errors_are_refused() {
    assert!(cell("<w:tcPr/><w:tcPr/>").properties().is_err());
    let not_a_cell = crate::Cell::new(format!(r#"<w:tr xmlns:w="{W}"/>"#).into_bytes());
    assert!(not_a_cell.properties().is_err());
    // The legacy readers keep accepting any root.
    assert_eq!(not_a_cell.grid_span().unwrap(), 1);
}

#[test]
fn malformed_attributes_and_root_content_are_refused_without_echoing_values() {
    for xml in [
        format!(r#"<w:r xmlns:w="{W}" a="PRIVATE" a="PRIVATE"><w:rPr/></w:r>"#),
        format!(
            r#"<w:r xmlns:w="{W}" xmlns:q="{W}"><w:rPr><w:b w:val="1" q:val="0"/></w:rPr></w:r>"#
        ),
        format!(r#"<w:r xmlns:w="{W}"><w:rPr><w:unknown a="PRIVATE" a="PRIVATE"/></w:rPr></w:r>"#),
        format!(r#"<w:r xmlns:w="{W}"><w:rPr><w:unknown a="&undefined;"/></w:rPr></w:r>"#),
        format!(r#"<!DOCTYPE r><w:r xmlns:w="{W}"/>"#),
        format!(r#"PRIVATE<w:r xmlns:w="{W}"/>"#),
    ] {
        let error = crate::Run::new(xml.into_bytes())
            .direct_properties()
            .unwrap_err();
        assert!(!format!("{error:?}").contains("PRIVATE"));
        assert!(!error.to_string().contains("PRIVATE"));
    }
}

#[test]
fn authored_property_strings_are_not_debug_output() {
    let run = run(r#"<w:rPr><w:rStyle w:val="PRIVATE"/><w:rFonts w:ascii="PRIVATE"/><w:lang w:val="PRIVATE"/></w:rPr>"#).direct_properties().unwrap();
    assert!(!format!("{run:?}").contains("PRIVATE"));
    assert!(!format!("{:?}", run.fonts).contains("PRIVATE"));
    assert!(!format!("{:?}", run.language).contains("PRIVATE"));
    let paragraph = paragraph(r#"<w:pPr><w:pStyle w:val="PRIVATE"/></w:pPr>"#)
        .direct_properties()
        .unwrap();
    assert!(!format!("{paragraph:?}").contains("PRIVATE"));
    let table = crate::Table::new(format!(r#"<w:tbl xmlns:w="{W}"><w:tblPr><w:tblCaption w:val="PRIVATE"/><w:tblDescription w:val="PRIVATE"/></w:tblPr></w:tbl>"#).into_bytes()).properties().unwrap();
    assert!(!format!("{table:?}").contains("PRIVATE"));
}

#[test]
fn retained_property_views_resolve_ancestor_aliases_and_foreign_prefixes() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:q="{W}" xmlns:f="urn:foreign"><w:body><w:p><w:pPr><q:jc q:val="center"/></w:pPr><q:r><q:rPr><q:b/><f:i/></q:rPr><q:t>x</q:t></q:r><f:r/><q:hyperlink q:anchor="target"><q:r><q:t>link</q:t></q:r></q:hyperlink></w:p><w:tbl><w:tblPr><q:tblW q:w="100" q:type="dxa"/></w:tblPr><w:tr><w:tc><w:tcPr><q:gridSpan q:val="2"/><q:vMerge q:val="restart"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
    );
    let mut package = crate::Package::new().unwrap();
    package
        .edit_opc_in_place(|opc| {
            opc.get_part_mut(&litchi_opc::PackURI::new("/word/document.xml").unwrap())?
                .set_blob(xml.into_bytes());
            Ok::<_, crate::Error>(())
        })
        .unwrap();
    let document = package.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(
        paragraphs[0].direct_properties().unwrap().justification,
        Some(Justification::Center)
    );
    let runs = paragraphs[0].addressable_runs().unwrap();
    assert_eq!(runs.len(), 2);
    let resolved = paragraphs[0]
        .inlines_with_relationships(package.opc_package().main_document_part().unwrap().rels())
        .unwrap();
    assert!(matches!(resolved.last(), Some(crate::Inline::Hyperlink(_))));
    let inlines = paragraphs[0].inlines().unwrap();
    assert!(matches!(inlines[0], crate::Inline::Run(_)));
    assert_eq!(runs[0].direct_properties().unwrap().bold, Some(true));
    assert_eq!(runs[0].direct_properties().unwrap().italic, None);
    let tables = document.tables().unwrap();
    assert!(tables[0].properties().unwrap().width.is_some());
    assert_eq!(
        tables[0].rows().unwrap()[0].cells().unwrap()[0]
            .properties()
            .unwrap()
            .grid_span,
        Some(2)
    );
    let rows = tables[0].rows().unwrap();
    let cells = rows[0].cells().unwrap();
    assert_eq!(cells[0].grid_span().unwrap(), 2);
    assert_eq!(cells[0].v_merge().unwrap(), Some(VMergeState::Restart));
    let source = crate::source_backed::Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
    let paragraphs = source.document().unwrap().paragraphs().unwrap();
    assert_eq!(
        paragraphs[0].direct_properties().unwrap().justification,
        Some(Justification::Center)
    );
    assert_eq!(
        paragraphs[0].addressable_runs().unwrap()[0]
            .direct_properties()
            .unwrap()
            .bold,
        Some(true)
    );
}

#[test]
fn positional_runs_refuse_malformed_roots_without_echoing_content() {
    for xml in [
        "PRIVATE<w:p/>",
        "<!DOCTYPE p><w:p/>",
        "<![CDATA[PRIVATE]]><w:p/>",
        "<w:p a='PRIVATE' a='PRIVATE'/>",
        "<w:p><PRIVATE></w:p>",
    ] {
        let error = crate::Paragraph::new(xml.as_bytes().to_vec())
            .addressable_runs()
            .unwrap_err();
        assert!(!format!("{error:?}").contains("PRIVATE"));
    }
    let paragraph = paragraph(
        "<w:r><w:t>PRIVATE</w:t></w:r><w:fldSimple w:instr='PRIVATE'/><w:bookmarkStart w:id='0' w:name='PRIVATE'/>",
    );
    assert!(!format!("{paragraph:?}").contains("PRIVATE"));
    for inline in paragraph.inlines().unwrap() {
        assert!(!format!("{inline:?}").contains("PRIVATE"));
        match inline {
            crate::Inline::Unknown(opaque) => assert!(!format!("{opaque:?}").contains("PRIVATE")),
            crate::Inline::Run(run) => assert!(!format!("{run:?}").contains("PRIVATE")),
            _ => {},
        }
    }
}

#[test]
fn cell_value_errors_do_not_echo_authored_content() {
    let cell = cell(r#"<w:tcPr><w:gridSpan w:val="PRIVATE"/><w:vMerge w:val="PRIVATE"/></w:tcPr>"#);
    assert!(
        !cell
            .grid_span()
            .unwrap_err()
            .to_string()
            .contains("PRIVATE")
    );
    assert!(!cell.v_merge().unwrap_err().to_string().contains("PRIVATE"));
}
