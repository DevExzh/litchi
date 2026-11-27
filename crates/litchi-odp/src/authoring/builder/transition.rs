//! Drawing-page transition style codec.

use super::xml::push_optional_attribute;
use crate::Slide;
use litchi_core::xml::escape_xml;

pub(crate) const DEFAULT_DRAWING_PAGE_STYLE_NAME: &str = "dp1";
pub(crate) const DEFAULT_DRAWING_PAGE_STYLE: &str = r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties/></style:style>"#;

pub(crate) fn slide_style_name(slide: &Slide, index: usize) -> String {
    if slide
        .transition
        .as_ref()
        .is_some_and(|value| !value.is_empty())
    {
        format!("dpTransition{}", index + 1)
    } else {
        DEFAULT_DRAWING_PAGE_STYLE_NAME.to_string()
    }
}

pub(crate) fn generate_transition_styles(slides: &[Slide]) -> String {
    let mut output = String::from(DEFAULT_DRAWING_PAGE_STYLE);
    for (index, slide) in slides.iter().enumerate() {
        push_transition_style(&mut output, slide, index);
    }
    output
}

pub(crate) fn push_transition_style(target: &mut String, slide: &Slide, index: usize) {
    let Some(transition) = slide.transition.as_ref().filter(|value| !value.is_empty()) else {
        return;
    };
    let mut output = String::new();
    output.push_str(r#"<style:style style:name=""#);
    output.push_str(&slide_style_name(slide, index));
    output.push_str(r#"" style:family="drawing-page"><style:drawing-page-properties"#);
    push_optional_attribute(
        &mut output,
        "presentation:transition-type",
        transition
            .transition_type
            .map(crate::model::transition::Type::as_str),
    );
    push_optional_attribute(
        &mut output,
        "presentation:transition-style",
        transition
            .style
            .as_ref()
            .map(crate::model::transition::Style::as_str),
    );
    push_optional_attribute(
        &mut output,
        "presentation:transition-speed",
        transition
            .speed
            .map(crate::model::transition::Speed::as_str),
    );
    push_optional_attribute(&mut output, "smil:type", transition.smil_type.as_deref());
    push_optional_attribute(
        &mut output,
        "smil:subtype",
        transition.smil_subtype.as_deref(),
    );
    push_optional_attribute(
        &mut output,
        "smil:direction",
        transition
            .direction
            .map(crate::model::transition::Direction::as_str),
    );
    push_optional_attribute(
        &mut output,
        "smil:fadeColor",
        transition.fade_color.as_deref(),
    );
    push_optional_attribute(
        &mut output,
        "presentation:duration",
        transition.duration.as_deref(),
    );
    if let Some(sound) = transition.sound.as_ref() {
        output.push('>');
        output.push_str(r#"<presentation:sound xlink:type="simple" xlink:href=""#);
        output.push_str(&escape_xml(&sound.href));
        output.push('"');
        if sound.actuate_on_request {
            output.push_str(r#" xlink:actuate="onRequest""#);
        }
        push_optional_attribute(
            &mut output,
            "xlink:show",
            sound.show.map(crate::model::transition::SoundShow::as_str),
        );
        push_optional_attribute(&mut output, "xml:id", sound.xml_id.as_deref());
        push_optional_attribute(
            &mut output,
            "presentation:play-full",
            sound
                .play_full
                .map(|value| if value { "true" } else { "false" }),
        );
        output.push_str("/></style:drawing-page-properties>");
    } else {
        output.push_str("/>");
    }
    output.push_str("</style:style>");
    target.push_str(&output);
}
