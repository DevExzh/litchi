//! Reader tests for the shared `DrawingML` text body.

use quick_xml::events::BytesStart;
use quick_xml::name::NamespaceResolver;

use super::reader::{self, Limits};
use super::{
    Body, Bullet, BulletColor, BulletFont, BulletSize, Field, Hyperlink, RunKind, Spacing, TabStop,
    TextFill,
};
use crate::color::{Rgb, Scheme, Value};
use crate::text::{
    Alignment, Anchor, AutoNumber, Autofit, Caps, Coordinate32, FontAlignment, Strike,
    TabAlignment, TextPoint, TextSize, Underline,
};

const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn body(content: &str) -> String {
    format!(r#"<p:txBody xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">{content}</p:txBody>"#)
}

fn read(xml: &[u8]) -> crate::Result<Body> {
    reader::read(xml, NamespaceResolver::default(), Limits::DEFAULT)
}

fn read_body(content: &str) -> Body {
    read(body(content).as_bytes()).unwrap()
}

#[test]
fn reads_paragraph_properties_spacing_tabs_and_bullets() {
    let body = read_body(
        r#"<a:bodyPr anchor="ctr" lIns="0" wrap="none"><a:normAutofit fontScale="90000"/></a:bodyPr><a:lstStyle/>
        <a:p><a:pPr marL="342900" marR="10" indent="-342900" lvl="1" algn="just" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="true" fontAlgn="base">
            <a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPct val="50%"/></a:spcAft>
            <a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="75000"/>
            <a:buFont typeface="Wingdings" pitchFamily="2" charset="2" panose="05000000000000000000"/>
            <a:buChar char="&#x2022;"/>
            <a:tabLst><a:tab pos="914400" algn="ctr"/><a:tab pos="1828800"/></a:tabLst>
            <a:defRPr sz="1800" b="1"/>
        </a:pPr><a:r><a:t>Bulleted</a:t></a:r><a:endParaRPr lang="en-US" sz="1200"/></a:p>
        <a:p><a:pPr><a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr></a:p>
        <a:p><a:pPr><a:buSzPts val="1400"/><a:buNone/></a:pPr></a:p>
        <a:p><a:pPr><a:buBlip><a:blip r:embed="rId7"/></a:buBlip></a:pPr></a:p>"#,
    );
    assert_eq!(body.properties.vertical_anchor, Anchor::Center);
    assert_eq!(body.properties.insets.left.as_emu(), Some(0));
    assert_eq!(body.properties.autofit, Autofit::Normal);
    assert_eq!(body.paragraphs.len(), 4);

    let first = &body.paragraphs[0].properties;
    assert_eq!(first.margin_left, Some(342_900));
    assert_eq!(first.margin_right, Some(10));
    assert_eq!(first.indent, Some(-342_900));
    assert_eq!(first.level, Some(1));
    assert_eq!(first.alignment, Some(Alignment::Justified));
    assert_eq!(
        first
            .default_tab_size
            .as_ref()
            .and_then(Coordinate32::as_emu),
        Some(914_400)
    );
    assert_eq!(first.right_to_left, Some(false));
    assert_eq!(first.east_asian_line_break, Some(true));
    assert_eq!(first.latin_line_break, Some(false));
    assert_eq!(first.hanging_punctuation, Some(true));
    assert_eq!(first.font_alignment, Some(FontAlignment::Baseline));
    assert_eq!(first.line_spacing, Some(Spacing::Percent(150_000)));
    assert_eq!(first.space_before, Some(Spacing::Points(600)));
    assert_eq!(first.space_after, Some(Spacing::Percent(50_000)));
    assert_eq!(
        first.bullet_color,
        Some(BulletColor::Color(Value::rgb(Rgb::new(0xFF, 0, 0))))
    );
    assert_eq!(first.bullet_size, Some(BulletSize::Percent(75_000)));
    let Some(BulletFont::Font(font)) = &first.bullet_font else {
        panic!("expected a bullet font");
    };
    assert_eq!(font.typeface, "Wingdings");
    assert_eq!((font.pitch_family, font.charset), (Some(2), Some(2)));
    assert_eq!(font.panose.as_deref(), Some("05000000000000000000"));
    assert_eq!(first.bullet, Some(Bullet::Character("\u{2022}".to_owned())));
    assert_eq!(
        first.tab_stops.as_deref(),
        Some(
            &[
                TabStop {
                    position: Some(Coordinate32::from(914_400)),
                    alignment: Some(TabAlignment::Center),
                },
                TabStop {
                    position: Some(Coordinate32::from(1_828_800)),
                    alignment: None,
                },
            ][..]
        )
    );
    let default_run = first.default_run.as_deref().unwrap();
    assert_eq!(default_run.size, TextSize::new(1800).ok());
    assert_eq!(default_run.bold, Some(true));
    let end = body.paragraphs[0].end_properties.as_deref().unwrap();
    assert_eq!(end.language.as_deref(), Some("en-US"));
    assert_eq!(end.size, TextSize::new(1200).ok());

    let second = &body.paragraphs[1].properties;
    assert_eq!(second.bullet_color, Some(BulletColor::FollowText));
    assert_eq!(second.bullet_size, Some(BulletSize::FollowText));
    assert_eq!(second.bullet_font, Some(BulletFont::FollowText));
    assert_eq!(
        second.bullet,
        Some(Bullet::AutoNumber {
            scheme: AutoNumber::ArabicPeriod,
            start_at: Some(3),
        })
    );
    let third = &body.paragraphs[2].properties;
    assert_eq!(
        third.bullet_size,
        Some(BulletSize::Points(TextSize::new(1400).unwrap()))
    );
    assert_eq!(third.bullet, Some(Bullet::None));
    assert_eq!(
        body.paragraphs[3].properties.bullet,
        Some(Bullet::Picture(super::Picture {
            relationship_id: Some("rId7".to_owned())
        }))
    );
    assert!(body.paragraphs[1].runs.is_empty());
}

#[test]
fn reads_run_properties_fonts_colors_links_breaks_and_fields() {
    let body = read_body(
        r#"<a:bodyPr/><a:p>
        <a:r><a:rPr lang="en-US" altLang="zh-CN" sz="2400" b="1" i="0" u="dbl" strike="sngStrike" kern="1200" cap="small" spc="-50" baseline="30000">
            <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>
            <a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
            <a:latin typeface="+mn-lt"/><a:ea typeface="MS Gothic"/><a:cs typeface="Arial"/><a:sym typeface="Symbol"/>
            <a:hlinkClick r:id="rId2" tooltip="Go &amp; see" action="ppaction://hlinkfile"/>
        </a:rPr><a:t>Styled</a:t></a:r>
        <a:br><a:rPr sz="1000"/></a:br>
        <a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:pPr/><a:t>&lt;#&gt;</a:t></a:fld>
        <a:r><a:rPr><a:noFill/></a:rPr><a:t xml:space="preserve"> tail </a:t></a:r>
        <a:r><a:rPr><a:gradFill><a:gsLst/></a:gradFill></a:rPr><a:t><![CDATA[x<y]]></a:t></a:r>
        </a:p>"#,
    );
    let runs = &body.paragraphs[0].runs;
    assert_eq!(runs.len(), 5);

    let styled = &runs[0];
    assert_eq!(styled.kind, RunKind::Text);
    assert_eq!(styled.text, "Styled");
    let properties = &styled.properties;
    assert_eq!(properties.language.as_deref(), Some("en-US"));
    assert_eq!(properties.alternate_language.as_deref(), Some("zh-CN"));
    assert_eq!(properties.size, TextSize::new(2400).ok());
    assert_eq!(properties.bold, Some(true));
    assert_eq!(properties.italic, Some(false));
    assert_eq!(properties.underline, Some(Underline::Double));
    assert_eq!(properties.strike, Some(Strike::Single));
    assert_eq!(properties.kerning, TextPoint::new(1200).ok());
    assert_eq!(properties.caps, Some(Caps::Small));
    assert_eq!(properties.spacing, TextPoint::new(-50).ok());
    assert_eq!(properties.baseline, Some(30_000));
    let color = properties.color().unwrap();
    let transformed = color.as_transformed().unwrap();
    assert_eq!(
        transformed.base(),
        &crate::color::Base::Scheme(Scheme::Accent1)
    );
    assert_eq!(transformed.transforms().len(), 1);
    assert_eq!(
        properties.highlight,
        Some(Value::rgb(Rgb::new(0xFF, 0xFF, 0)))
    );
    let latin = properties.latin.as_ref().unwrap();
    assert_eq!(latin.typeface, "+mn-lt");
    assert_eq!(latin.theme_reference(), Some(("mn", "lt")));
    assert_eq!(
        properties.east_asian.as_ref().unwrap().typeface,
        "MS Gothic"
    );
    assert_eq!(
        properties.complex_script.as_ref().unwrap().typeface,
        "Arial"
    );
    assert_eq!(properties.symbol.as_ref().unwrap().typeface, "Symbol");
    assert_eq!(
        properties.hyperlink,
        Some(Hyperlink {
            relationship_id: Some("rId2".to_owned()),
            action: Some("ppaction://hlinkfile".to_owned()),
            tooltip: Some("Go & see".to_owned()),
        })
    );

    assert_eq!(runs[1].kind, RunKind::Break);
    assert_eq!(runs[1].text, "\n");
    assert_eq!(runs[1].properties.size, TextSize::new(1000).ok());

    assert_eq!(
        runs[2].kind,
        RunKind::Field(Field {
            id: "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}".to_owned(),
            kind: Some("slidenum".to_owned()),
        })
    );
    assert_eq!(runs[2].text, "<#>");
    assert_eq!(runs[2].properties.language.as_deref(), Some("en-US"));

    assert_eq!(runs[3].text, " tail ");
    assert_eq!(runs[3].properties.fill, Some(TextFill::None));
    assert_eq!(runs[4].text, "x<y");
    assert_eq!(runs[4].properties.fill, Some(TextFill::Gradient));
    assert_eq!(body.paragraphs[0].text(), "Styled\n<#> tail x<y");
}

#[test]
fn resolves_prefixes_inherited_from_the_owner_and_strict_names() {
    // The fragment uses prefixes its ancestors declare.
    let fragment = br#"<p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId9"/></a:rPr><a:t>Inherited</a:t></a:r></a:p></p:txBody>"#;
    let owner = format!(r#"<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"/>"#);
    let mut reader = quick_xml::Reader::from_reader(owner.as_bytes());
    let quick_xml::events::Event::Empty(root) = reader.read_event().unwrap() else {
        panic!("expected the owner root");
    };
    let mut scope = NamespaceResolver::default();
    scope.push(&root).unwrap();
    let body = reader::read(fragment, scope, Limits::DEFAULT).unwrap();
    assert_eq!(body.text(), "Inherited");
    assert_eq!(
        body.paragraphs[0].runs[0]
            .properties
            .hyperlink
            .as_ref()
            .and_then(|link| link.relationship_id.as_deref()),
        Some("rId9")
    );

    // Without the inherited scope nothing is a DrawingML element.
    let unbound = read(fragment).unwrap();
    assert!(unbound.paragraphs.is_empty());

    let strict = format!(
        r#"<x:txBody xmlns:x="urn:host" xmlns:s="{STRICT_A}"><s:p><s:pPr algn="r"/><s:r><s:t>S</s:t></s:r></s:p></x:txBody>"#
    );
    let body = read(strict.as_bytes()).unwrap();
    assert_eq!(body.text(), "S");
    assert_eq!(
        body.paragraphs[0].properties.alignment,
        Some(Alignment::Right)
    );
}

#[test]
fn skips_unmodeled_markup_without_interpreting_it() {
    let body = read_body(
        r#"<a:bodyPr><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></a:bodyPr>
        <a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="9999999"/></a:lvl1pPr></a:lstStyle>
        <future:element xmlns:future="urn:future"><a:p><a:r><a:t>hidden</a:t></a:r></a:p></future:element>
        <a:p>stray text<a:r><a:rPr><a:ln w="1"/><a:effectLst><a:glow/></a:effectLst><a:extLst><a:ext uri="x"/></a:extLst></a:rPr><a:t>A<a:unknown>skipped</a:unknown>B</a:t></a:r>
        <a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a:r><a:t>math</a:t></a:r></a14:m></a:p>"#,
    );
    assert_eq!(body.paragraphs.len(), 1);
    assert_eq!(body.text(), "AB");
    assert_eq!(body.paragraphs[0].properties.alignment, None);
}

#[test]
fn duplicate_or_competing_properties_are_refused() {
    for content in [
        r#"<a:p><a:pPr algn="l"/><a:pPr lvl="2"/></a:p>"#,
        r#"<a:p><a:r><a:rPr b="1"/><a:rPr i="1"/><a:t>x</a:t></a:r></a:p>"#,
        r#"<a:bodyPr><a:noAutofit/><a:spAutoFit/></a:bodyPr>"#,
        r#"<a:p><a:pPr><a:buNone/><a:buChar char="x"/></a:pPr></a:p>"#,
        r#"<a:p><a:pPr><a:lnSpc><a:spcPts val="100"/><a:spcPct val="100000"/></a:lnSpc></a:pPr></a:p>"#,
        r#"<a:p><a:r><a:rPr><a:noFill/><a:solidFill/></a:rPr></a:r></a:p>"#,
    ] {
        assert!(read(body(content).as_bytes()).is_err(), "{content}");
    }
}

#[test]
fn keeps_colors_the_color_codec_cannot_type_losslessly() {
    let body = read_body(
        r#"<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent9"/></a:solidFill></a:rPr><a:t>x</a:t></a:r>
        <a:r><a:rPr><a:solidFill/></a:rPr><a:t>y</a:t></a:r>
        <a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:satMod val="150000"/></a:schemeClr></a:solidFill></a:rPr><a:t>z</a:t></a:r>
        <a:r><a:rPr><a:highlight><a:srgbClr val="XYZ"/></a:highlight></a:rPr><a:t>w</a:t></a:r></a:p>"#,
    );
    let runs = &body.paragraphs[0].runs;
    assert!(runs[0].properties.color().unwrap().is_unknown());
    assert_eq!(runs[1].properties.fill, Some(TextFill::Solid(None)));
    let Some(Value::Unknown(saturated)) = runs[2].properties.color() else {
        panic!("expected an untyped color");
    };
    assert_eq!(
        saturated.as_xml(),
        br#"<a:schemeClr val="accent1"><a:satMod val="150000"/></a:schemeClr>"#
    );
    assert!(runs[3].properties.highlight.as_ref().unwrap().is_unknown());
}

#[test]
fn strict_profile_refuses_invalid_modeled_values() {
    for content in [
        r#"<a:p><a:pPr algn="middle"/></a:p>"#,
        r#"<a:p><a:pPr lvl="9"/></a:p>"#,
        r#"<a:p><a:pPr marL="-1"/></a:p>"#,
        r#"<a:p><a:pPr><a:lnSpc><a:spcPts val="158401"/></a:lnSpc></a:pPr></a:p>"#,
        r#"<a:p><a:pPr><a:buSzPct val="20000"/></a:pPr></a:p>"#,
        r#"<a:p><a:pPr><a:buAutoNum type="arabicDash"/></a:pPr></a:p>"#,
        r#"<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="0"/></a:pPr></a:p>"#,
        r#"<a:p><a:pPr><a:buChar/></a:pPr></a:p>"#,
        r#"<a:p><a:r><a:rPr sz="99"/><a:t>x</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:rPr strike="yes"/><a:t>x</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:rPr kern="-1"/><a:t>x</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:rPr cap="upper"/><a:t>x</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:rPr><a:latin panose="00"/></a:rPr><a:t>x</a:t></a:r></a:p>"#,
        r#"<a:p><a:fld type="slidenum"><a:t>1</a:t></a:fld></a:p>"#,
        r#"<a:p><a:br><a:rPr b="maybe"/></a:br></a:p>"#,
        r#"<a:bodyPr anchor="middle"/>"#,
    ] {
        assert!(
            read(body(content).as_bytes()).is_err(),
            "accepted {content}"
        );
    }
    let alternate = body(&format!(
        r#"<a:p><mc:AlternateContent xmlns:mc="{MC}"><mc:Fallback><a:r><a:t>x</a:t></a:r></mc:Fallback></mc:AlternateContent></a:p>"#
    ));
    assert!(read(alternate.as_bytes()).is_err());
}

#[test]
fn strict_reader_refuses_previously_ignored_invalid_properties() {
    for content in [
        r#"<a:p><a:pPr algn="middle"/></a:p>"#,
        r#"<a:p><a:r><a:rPr strike="yes"/></a:r></a:p>"#,
        r#"<a:p><a:br><a:rPr sz="1"/></a:br></a:p>"#,
        r#"<a:p><a:endParaRPr sz="7"/></a:p>"#,
    ] {
        assert!(read(body(content).as_bytes()).is_err(), "{content}");
    }
}

#[test]
fn limits_bound_text_nodes_depth_and_input() {
    let xml = body(r#"<a:p><a:r><a:t>0123456789</a:t></a:r></a:p>"#);
    let text_limit = Limits::new(1 << 20, 64, 64, 9).unwrap();
    assert!(matches!(
        reader::read(xml.as_bytes(), NamespaceResolver::default(), text_limit),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::TextBytes,
            ..
        })
    ));
    let node_limit = Limits::new(1 << 20, 64, 2, 64).unwrap();
    assert!(matches!(
        reader::read(xml.as_bytes(), NamespaceResolver::default(), node_limit),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::TextBodyElements,
            ..
        })
    ));
    let depth_limit = Limits::new(1 << 20, 2, 64, 64).unwrap();
    assert!(matches!(
        reader::read(xml.as_bytes(), NamespaceResolver::default(), depth_limit),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::TextBodyNestingDepth,
            ..
        })
    ));
    let input_limit = Limits::new(8, 64, 64, 64).unwrap();
    assert!(matches!(
        reader::read(xml.as_bytes(), NamespaceResolver::default(), input_limit),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::InputBytes,
            ..
        })
    ));
    assert!(Limits::new(0, 1, 1, 1).is_none());
    assert_eq!(Limits::default(), Limits::DEFAULT);
}

#[test]
fn text_budget_preflights_decoded_entity_bytes() {
    let xml = body(r#"<a:p><a:r><a:t>&#x1F600;</a:t></a:r></a:p>"#);
    let exact = Limits::new(1 << 20, 64, 64, 4).unwrap();
    let parsed = reader::read(xml.as_bytes(), NamespaceResolver::default(), exact)
        .expect("decoded entity fits its byte budget");
    assert_eq!(parsed.text(), "😀");

    let short = Limits::new(1 << 20, 64, 64, 3).unwrap();
    assert!(matches!(
        reader::read(xml.as_bytes(), NamespaceResolver::default(), short),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::TextBytes,
            ..
        })
    ));
}

#[test]
fn rejects_fragments_that_are_not_one_element() {
    for xml in [
        String::new(),
        format!("{}{}", body(""), body("")),
        format!(r#"<p:txBody xmlns:p="{P}"><a:p>"#),
        format!(r#"<!DOCTYPE x><p:txBody xmlns:p="{P}"/>"#),
    ] {
        assert!(read(xml.as_bytes()).is_err(), "accepted {xml:?}");
    }
    let empty = format!(r#"<p:txBody xmlns:p="{P}"/>"#);
    assert_eq!(read(empty.as_bytes()).unwrap(), Body::default());
}

#[test]
fn builder_and_body_property_helpers_are_usable_by_streaming_hosts() {
    let mut properties = super::Properties::default();
    let element = BytesStart::from_content(r#"a:bodyPr numCol="2" wrap="none""#, 8);
    let decoder = quick_xml::Reader::from_reader(&b""[..]).decoder();
    reader::apply_body_properties(&mut properties, &element, decoder).unwrap();
    assert_eq!(properties.column_count.get(), 2);
    assert_eq!(reader::autofit(b"spAutoFit"), Some(Autofit::Shape));
    assert_eq!(reader::autofit(b"bodyPr"), None);

    let builder = reader::Builder::new(b"", Limits::DEFAULT).unwrap();
    assert_eq!(builder.finish().unwrap(), Body::default());
}

#[test]
fn tab_stops_are_not_text_and_run_content_tabs_are() {
    let content = r#"<a:lstStyle><a:lvl1pPr><a:tabLst><a:tab pos="457200"/></a:tabLst></a:lvl1pPr></a:lstStyle>
        <a:p><a:pPr><a:tabLst><a:tab pos="914400" algn="ctr"/><a:tab pos="1828800"></a:tab></a:tabLst></a:pPr>
        <a:r><a:t>x</a:t></a:r><a:r><a:tab/><a:t>y</a:t><a:br/></a:r><a:tab/></a:p>"#;
    let strict = read_body(content);
    let paragraph = &strict.paragraphs[0];
    assert_eq!(
        paragraph.properties.tab_stops.as_ref().map(Vec::len),
        Some(2)
    );
    assert_eq!(paragraph.runs.len(), 2);
    assert_eq!(paragraph.runs[1].text, "\ty\n");
    assert_eq!(strict.text(), "x\ty\n");
}

#[test]
fn strict_relationships_and_debug_are_content_free() {
    for attr in [
        r#"r:id="private-one" s:id="private-two""#,
        r#"unknown:id="private-one""#,
    ] {
        let xml = body(&format!(
            r#"<a:p><a:r><a:rPr><a:hlinkClick xmlns:s="{R}" {attr}/></a:rPr></a:r></a:p>"#
        ));
        let error = read(xml.as_bytes()).unwrap_err();
        assert!(!format!("{error:?} {error}").contains("private"));
    }
    let parsed = read_body(
        r#"<a:p><a:r><a:rPr lang="private-language"><a:hlinkClick r:id="private-relationship" tooltip="private-tooltip"/></a:rPr><a:t>private-text</a:t></a:r></a:p>"#,
    );
    let debug = format!(
        "{parsed:?} {:?} {:?} {:?}",
        parsed.paragraphs[0], parsed.paragraphs[0].runs[0], parsed.paragraphs[0].runs[0].properties
    );
    assert!(!debug.contains("private"));
    let too_small = Limits::new(1, 4, 4, 4).unwrap();
    assert!(matches!(
        reader::Builder::new(b"too long", too_small),
        Err(crate::Error::Limit { .. })
    ));
}

#[test]
fn generated_break_and_tab_text_respects_the_budget() {
    for content in [
        r#"<a:p><a:br/><a:br/></a:p>"#,
        r#"<a:p><a:r><a:tab/><a:br/></a:r></a:p>"#,
    ] {
        let xml = body(content);
        let limits = Limits::new(4096, 64, 64, 1).unwrap();
        assert!(matches!(
            reader::read(xml.as_bytes(), NamespaceResolver::default(), limits),
            Err(crate::Error::Limit {
                kind: crate::LimitKind::TextBytes,
                ..
            })
        ));
    }
}

#[test]
fn field_ids_and_relationship_dialects_are_checked() {
    for field_id in [
        "",
        "private-invalid",
        "00000000-0000-0000-0000-000000000000",
    ] {
        let xml = body(&format!(
            r#"<a:p><a:fld id="{field_id}"><a:t>x</a:t></a:fld></a:p>"#
        ));
        let error = read(xml.as_bytes()).unwrap_err();
        assert!(!format!("{error:?} {error}").contains("private-invalid"));
    }
    let xml = body(r#"<a:p><a:r><a:rPr><a:hlinkClick r:id="payload"/></a:rPr></a:r></a:p>"#)
        .replace(A, STRICT_A)
        .replace(P, "http://purl.oclc.org/ooxml/presentationml/main");
    assert!(read(xml.as_bytes()).is_err());
    let xml = xml.replace(R, "http://purl.oclc.org/ooxml/officeDocument/relationships");
    assert!(read(xml.as_bytes()).is_ok());
}

#[test]
fn unmodeled_autofit_scale_is_not_silently_reauthored() {
    let parsed = read_body(r#"<a:bodyPr><a:normAutofit fontScale="85000"/></a:bodyPr><a:p/>"#);
    assert!(parsed.properties.has_unmodeled_content);
    let mut output = "unchanged".to_owned();
    assert!(super::writer::write(&mut output, &parsed, super::writer::Limits::DEFAULT,).is_err());
    assert_eq!(output, "unchanged");
}

#[test]
fn action_only_links_are_preserved_and_empty_color_choices_are_not_written() {
    let parsed = read_body(
        r#"<a:p><a:r><a:rPr><a:hlinkClick action="ppaction://hlinkshowjump?jump=nextslide" tooltip="private-tooltip"/></a:rPr><a:t>next</a:t></a:r></a:p>"#,
    );
    let link = parsed.paragraphs[0].runs[0]
        .properties
        .hyperlink
        .as_ref()
        .unwrap();
    assert!(super::raw::hyperlink_relationship_id(link).is_none());
    assert_eq!(
        link.action.as_deref(),
        Some("ppaction://hlinkshowjump?jump=nextslide")
    );
    assert_eq!(link.tooltip.as_deref(), Some("private-tooltip"));
    for content in [
        r#"<a:p><a:r><a:rPr><a:highlight/></a:rPr></a:r></a:p>"#,
        r#"<a:p><a:pPr><a:buClr/></a:pPr></a:p>"#,
    ] {
        let parsed = read_body(content);
        assert!(parsed.properties.has_unmodeled_content);
        let mut output = "untouched".to_owned();
        assert!(
            super::writer::write(&mut output, &parsed, super::writer::Limits::DEFAULT,).is_err()
        );
        assert_eq!(output, "untouched");
    }
}
