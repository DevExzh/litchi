//! The single-property run readers read only the run's own `w:rPr`
//! (finding F-8 of record 0010).

use crate::Run;
use crate::format::UnderlineStyle;
use litchi_core::VerticalPosition;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn run(inner: &str) -> Run {
    Run::new(format!(r#"<w:r xmlns:w="{W}">{inner}</w:r>"#).into_bytes())
}

/// A previous-formatting snapshot holding every property the readers know.
const CHANGE: &str = concat!(
    r#"<w:rPrChange w:id="1" w:author="A"><w:rPr><w:rFonts w:ascii="Old"/><w:b/><w:i/><w:strike/>"#,
    r#"<w:sz w:val="40"/><w:u w:val="double"/><w:vertAlign w:val="superscript"/></w:rPr></w:rPrChange>"#
);

/// A text box inside the run whose own run states every property.
const TEXT_BOX: &str = concat!(
    r#"<w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:r>"#,
    r#"<w:rPr><w:rFonts w:ascii="Box"/><w:b/><w:i/><w:strike/><w:sz w:val="40"/><w:u w:val="double"/>"#,
    r#"<w:vertAlign w:val="subscript"/></w:rPr><w:t>box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>"#
);

fn assert_reads_nothing(run: &Run) {
    assert_eq!(run.bold().unwrap(), None);
    assert_eq!(run.italic().unwrap(), None);
    assert_eq!(run.strikethrough().unwrap(), None);
    assert_eq!(run.vertical_position().unwrap(), None);
    assert_eq!(run.font_name().unwrap(), None);
    assert_eq!(run.font_size().unwrap(), None);
    let properties = run.get_properties().unwrap();
    assert_eq!(properties.bold, None);
    assert_eq!(properties.italic, None);
    assert_eq!(properties.strikethrough, None);
    assert_eq!(properties.underline, None);
    assert_eq!(properties.vertical_position, None);
    let (_, properties) = run.get_text_and_properties().unwrap();
    assert_eq!(properties.bold, None);
    assert_eq!(properties.italic, None);
    assert_eq!(properties.strikethrough, None);
    assert_eq!(properties.underline, None);
    assert_eq!(properties.vertical_position, None);
}

#[test]
fn readers_skip_the_previous_formatting_of_a_tracked_change() {
    assert_reads_nothing(&run(&format!(
        r#"<w:rPr><w:color w:val="FF0000"/>{CHANGE}</w:rPr><w:t>x</w:t>"#
    )));
}

#[test]
fn readers_skip_the_runs_of_a_text_box_inside_the_run() {
    assert_reads_nothing(&run(TEXT_BOX));
    assert_reads_nothing(&run(&format!(r#"<w:rPr/>{TEXT_BOX}"#)));
    assert_reads_nothing(&run(&format!(
        r#"<w:rPr><w:color w:val="FF0000"/></w:rPr>{TEXT_BOX}"#
    )));
}

#[test]
fn readers_keep_reading_the_runs_own_properties() {
    let run = run(&format!(
        concat!(
            r#"<w:rPr><w:rFonts w:ascii="Arial"/><w:b w:val="0"/><w:i/><w:strike w:val="1"/><w:sz w:val="24"/>"#,
            r#"<w:u w:val="single"/><w:vertAlign w:val="subscript"/>{}</w:rPr><w:t>text</w:t>{}"#
        ),
        CHANGE, TEXT_BOX
    ));
    assert_eq!(run.bold().unwrap(), Some(false));
    assert_eq!(run.italic().unwrap(), Some(true));
    assert_eq!(run.strikethrough().unwrap(), Some(true));
    assert_eq!(
        run.vertical_position().unwrap(),
        Some(VerticalPosition::Subscript)
    );
    assert_eq!(run.font_name().unwrap().as_deref(), Some("Arial"));
    assert_eq!(run.font_size().unwrap(), Some(24));
    let properties = run.get_properties().unwrap();
    assert_eq!(properties.bold, Some(false));
    assert_eq!(properties.italic, Some(true));
    assert_eq!(properties.strikethrough, Some(true));
    assert_eq!(properties.underline, Some(UnderlineStyle::Single));
    assert_eq!(
        properties.vertical_position,
        Some(VerticalPosition::Subscript)
    );
    let (text, properties) = run.get_text_and_properties().unwrap();
    // The text reader is unchanged: it still includes the text box's text.
    assert_eq!(text, "textbox");
    assert_eq!(properties.bold, Some(false));
    assert_eq!(properties.underline, Some(UnderlineStyle::Single));
}

#[test]
fn readers_keep_their_lenient_value_and_name_matching() {
    // Values other than the on tokens read as off, and names are matched by
    // local name, as before; `Run::direct_properties` is the checked reader.
    let run = run(r#"<w:rPr xmlns:x="urn:x"><w:b w:val="yes"/><x:i/></w:rPr>"#);
    assert_eq!(run.bold().unwrap(), Some(false));
    assert_eq!(run.italic().unwrap(), Some(true));
    let checked = run.direct_properties().unwrap();
    assert_eq!(checked.bold, None);
    assert_eq!(checked.italic, None);
}

#[test]
fn readers_accept_an_rpr_fragment_as_before() {
    let properties =
        Run::new(format!(r#"<w:rPr xmlns:w="{W}"><w:b/><w:sz w:val="22"/></w:rPr>"#).into_bytes());
    assert_eq!(properties.bold().unwrap(), Some(true));
    assert_eq!(properties.font_size().unwrap(), Some(22));
}

#[test]
fn readers_still_refuse_malformed_xml() {
    let malformed = run("<w:rPr><w:i/></w:rPr><w:t>x</w:r>");
    assert!(malformed.bold().is_err());
    assert!(malformed.vertical_position().is_err());
    // As before, the font readers stop at the end of the run's `w:rPr`.
    assert_eq!(malformed.font_name().unwrap(), None);
    assert_eq!(malformed.italic().unwrap(), Some(true));
}

#[test]
fn effects_skip_the_runs_of_a_text_box_inside_the_run() {
    let glow = concat!(
        r#"<w14:glow xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:rad="63500">"#,
        r#"<w14:srgbClr w14:val="FF0000"/></w14:glow>"#
    );
    let nested = run(&format!(
        r#"<w:rPr><w:b/></w:rPr><w:pict><w:txbxContent><w:p><w:r><w:rPr>{glow}</w:rPr></w:r></w:p></w:txbxContent></w:pict>"#
    ));
    assert!(
        nested
            .effects()
            .unwrap()
            .iter()
            .all(|effect| !matches!(effect, crate::Effect::Glow(_)))
    );
    let own = run(&format!(r#"<w:rPr>{glow}</w:rPr>"#));
    assert!(
        own.effects()
            .unwrap()
            .iter()
            .any(|effect| matches!(effect, crate::Effect::Glow(_)))
    );
}
