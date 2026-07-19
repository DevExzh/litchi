//! Inert ODF presentation settings and custom-show metadata.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_CUSTOM_SHOWS: usize = 65_536;
const MAX_CUSTOM_SHOW_PAGES: usize = 65_536;

/// Schema-defined on/off state used by presentation features.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresentationFeatureState {
    /// The feature is enabled.
    Enabled,
    /// The feature is disabled.
    Disabled,
}

impl PresentationFeatureState {
    fn parse(attribute: &str, value: &str) -> Result<Self> {
        match value {
            "enabled" => Ok(Self::Enabled),
            "disabled" => Ok(Self::Disabled),
            _ => Err(invalid(format!(
                "presentation:{attribute} must be 'enabled' or 'disabled'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Enabled => "enabled",
            Self::Disabled => "disabled",
        }
    }
}

/// One named custom slide show in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomPresentationShow {
    /// Unique custom-show name.
    pub name: String,
    /// Ordered drawing-page names.
    pub pages: Vec<String>,
}

impl CustomPresentationShow {
    /// Create a validated custom show.
    pub fn new(name: impl Into<String>, pages: Vec<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            pages,
        };
        validate_custom_show(&value)?;
        Ok(value)
    }
}

/// Static `presentation:settings` metadata.
///
/// These values are retained and written but are never used to launch or control
/// a slide show.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PresentationSettings {
    pub animations: Option<PresentationFeatureState>,
    pub endless: Option<bool>,
    pub force_manual: Option<bool>,
    pub full_screen: Option<bool>,
    pub mouse_as_pen: Option<bool>,
    pub mouse_visible: Option<bool>,
    /// Validated XML Schema duration lexical value.
    pub pause: Option<String>,
    /// Name of the custom show selected for playback.
    pub show: Option<String>,
    pub show_end_of_presentation_slide: Option<bool>,
    pub show_logo: Option<bool>,
    /// Drawing-page name selected as the first slide.
    pub start_page: Option<String>,
    pub start_with_navigator: Option<bool>,
    pub stay_on_top: Option<bool>,
    pub transition_on_click: Option<PresentationFeatureState>,
    /// Named custom shows in document order.
    pub custom_shows: Vec<CustomPresentationShow>,
}

impl PresentationSettings {
    /// Validate all settings and cross-references.
    pub fn validate(&self) -> Result<()> {
        if let Some(pause) = &self.pause {
            validate_duration(pause)?;
        }
        if let Some(show) = &self.show {
            validate_text(show, "presentation:show", false)?;
        }
        if let Some(start_page) = &self.start_page {
            validate_text(start_page, "presentation:start-page", false)?;
        }
        if self.custom_shows.len() > MAX_CUSTOM_SHOWS {
            return Err(invalid("presentation settings exceed 65536 custom shows"));
        }
        let mut names = HashSet::with_capacity(self.custom_shows.len());
        for show in &self.custom_shows {
            validate_custom_show(show)?;
            if !names.insert(show.name.as_str()) {
                return Err(invalid(format!(
                    "duplicate custom presentation show '{}'",
                    show.name
                )));
            }
        }
        if let Some(selected) = &self.show
            && !names.contains(selected.as_str())
        {
            return Err(invalid(format!(
                "presentation:show references missing custom show '{selected}'"
            )));
        }
        Ok(())
    }

    /// Return whether no settings or custom shows are explicitly present.
    pub fn is_empty(&self) -> bool {
        self.animations.is_none()
            && self.endless.is_none()
            && self.force_manual.is_none()
            && self.full_screen.is_none()
            && self.mouse_as_pen.is_none()
            && self.mouse_visible.is_none()
            && self.pause.is_none()
            && self.show.is_none()
            && self.show_end_of_presentation_slide.is_none()
            && self.show_logo.is_none()
            && self.start_page.is_none()
            && self.start_with_navigator.is_none()
            && self.stay_on_top.is_none()
            && self.transition_on_click.is_none()
            && self.custom_shows.is_empty()
    }
}

/// Validate page-name references against the names that will be emitted on
/// direct `draw:page` children.
pub(crate) fn validate_presentation_page_references(
    settings: Option<&PresentationSettings>,
    page_names: &[String],
) -> Result<()> {
    let Some(settings) = settings else {
        return Ok(());
    };
    settings.validate()?;

    let mut unique_names = HashSet::with_capacity(page_names.len());
    let mut ambiguous_names = HashSet::new();
    for name in page_names {
        if !unique_names.insert(name.as_str()) {
            ambiguous_names.insert(name.as_str());
        }
    }

    let validate_reference = |name: &str, description: &str| -> Result<()> {
        if !unique_names.contains(name) {
            return Err(invalid(format!(
                "{description} references missing presentation page '{name}'"
            )));
        }
        if ambiguous_names.contains(name) {
            return Err(invalid(format!(
                "{description} references ambiguous presentation page '{name}'"
            )));
        }
        Ok(())
    };

    if let Some(name) = &settings.start_page {
        validate_reference(name, "presentation:start-page")?;
    }
    for show in &settings.custom_shows {
        for name in &show.pages {
            validate_reference(
                name,
                &format!("custom presentation show '{}'", show.name),
            )?;
        }
    }
    Ok(())
}

pub(crate) fn settings_reference_page(
    settings: Option<&PresentationSettings>,
    page_name: &str,
) -> bool {
    settings.is_some_and(|settings| {
        settings.start_page.as_deref() == Some(page_name)
            || settings
                .custom_shows
                .iter()
                .any(|show| show.pages.iter().any(|name| name == page_name))
    })
}

/// Parse the single direct `presentation:settings` child, if present.
pub fn parse_presentation_settings(xml: &str) -> Result<Option<PresentationSettings>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("presentation settings XML exceeds 8 MiB"));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut presentation_depth = None;
    let mut settings_depth = None;
    let mut show_depth = None;
    let mut found_presentation = false;
    let mut found_settings = false;
    let mut settings = PresentationSettings::default();

    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML nesting overflow"))?;
                if element_is(&reader, &element, OFFICE_NAMESPACE, b"presentation") {
                    if found_presentation {
                        return Err(invalid("duplicate office:presentation element"));
                    }
                    found_presentation = true;
                    presentation_depth = Some(depth);
                } else if element_is(&reader, &element, PRESENTATION_NAMESPACE, b"settings") {
                    if presentation_depth != Some(depth - 1) {
                        return Err(invalid(
                            "presentation:settings must be a direct office:presentation child",
                        ));
                    }
                    if found_settings {
                        return Err(invalid("duplicate presentation:settings element"));
                    }
                    found_settings = true;
                    parse_settings_attributes(&reader, &element, &mut settings)?;
                    settings_depth = Some(depth);
                } else if settings_depth == Some(depth - 1)
                    && element_is(&reader, &element, PRESENTATION_NAMESPACE, b"show")
                {
                    settings
                        .custom_shows
                        .push(parse_custom_show(&reader, &element)?);
                    show_depth = Some(depth);
                } else if settings_depth.is_some() {
                    return Err(invalid(
                        "presentation:settings may contain only presentation:show elements",
                    ));
                }
            },
            Event::Empty(element) => {
                if element_is(&reader, &element, PRESENTATION_NAMESPACE, b"settings") {
                    if presentation_depth != Some(depth) {
                        return Err(invalid(
                            "presentation:settings must be a direct office:presentation child",
                        ));
                    }
                    if found_settings {
                        return Err(invalid("duplicate presentation:settings element"));
                    }
                    found_settings = true;
                    parse_settings_attributes(&reader, &element, &mut settings)?;
                } else if settings_depth == Some(depth)
                    && element_is(&reader, &element, PRESENTATION_NAMESPACE, b"show")
                {
                    settings
                        .custom_shows
                        .push(parse_custom_show(&reader, &element)?);
                } else if settings_depth.is_some() {
                    return Err(invalid(
                        "presentation:settings may contain only presentation:show elements",
                    ));
                }
            },
            Event::End(element) => {
                if show_depth == Some(depth) {
                    if !end_is(&reader, &element, PRESENTATION_NAMESPACE, b"show") {
                        return Err(invalid("unexpected element inside presentation:show"));
                    }
                    show_depth = None;
                } else if settings_depth == Some(depth) {
                    if !end_is(&reader, &element, PRESENTATION_NAMESPACE, b"settings") {
                        return Err(invalid("unexpected presentation settings end element"));
                    }
                    settings_depth = None;
                } else if presentation_depth == Some(depth)
                    && end_is(&reader, &element, OFFICE_NAMESPACE, b"presentation")
                {
                    presentation_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced XML end element"))?;
            },
            Event::Text(text) if settings_depth.is_some() => {
                let bytes: &[u8] = text.as_ref();
                if bytes.iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(invalid("presentation settings cannot contain text"));
                }
            },
            Event::CData(cdata) if settings_depth.is_some() => {
                let bytes: &[u8] = cdata.as_ref();
                if bytes.iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(invalid("presentation settings cannot contain CDATA"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("active XML declarations are prohibited"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if settings_depth.is_some() || show_depth.is_some() {
        return Err(invalid("unterminated presentation settings"));
    }
    if !found_settings {
        return Ok(None);
    }
    settings.validate()?;
    Ok(Some(settings))
}

/// Serialize validated presentation settings in schema order.
pub(crate) fn write_presentation_settings(
    settings: Option<&PresentationSettings>,
) -> Result<String> {
    let Some(settings) = settings else {
        return Ok(String::new());
    };
    settings.validate()?;
    if settings.is_empty() {
        return Ok(String::new());
    }
    let mut output = String::with_capacity(256);
    output.push_str("<presentation:settings");
    write_state(&mut output, "animations", settings.animations);
    write_bool(&mut output, "endless", settings.endless);
    write_bool(&mut output, "force-manual", settings.force_manual);
    write_bool(&mut output, "full-screen", settings.full_screen);
    write_bool(&mut output, "mouse-as-pen", settings.mouse_as_pen);
    write_bool(&mut output, "mouse-visible", settings.mouse_visible);
    write_text(&mut output, "pause", settings.pause.as_deref());
    write_text(&mut output, "show", settings.show.as_deref());
    write_bool(
        &mut output,
        "show-end-of-presentation-slide",
        settings.show_end_of_presentation_slide,
    );
    write_bool(&mut output, "show-logo", settings.show_logo);
    write_text(&mut output, "start-page", settings.start_page.as_deref());
    write_bool(
        &mut output,
        "start-with-navigator",
        settings.start_with_navigator,
    );
    write_bool(&mut output, "stay-on-top", settings.stay_on_top);
    write_state(
        &mut output,
        "transition-on-click",
        settings.transition_on_click,
    );
    if settings.custom_shows.is_empty() {
        output.push_str("/>");
        return Ok(output);
    }
    output.push('>');
    for show in &settings.custom_shows {
        output.push_str("<presentation:show presentation:name=\"");
        output.push_str(&escape_xml(&show.name));
        output.push_str("\" presentation:pages=\"");
        output.push_str(&escape_xml(&show.pages.join(",")));
        output.push_str("\"/>");
    }
    output.push_str("</presentation:settings>");
    Ok(output)
}

fn parse_settings_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    settings: &mut PresentationSettings,
) -> Result<()> {
    let mut seen = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(found) if found.as_ref() == PRESENTATION_NAMESPACE)
        {
            return Err(invalid(
                "unsupported presentation:settings attribute namespace",
            ));
        }
        let local = std::str::from_utf8(local_name.as_ref())
            .map_err(|_| invalid("presentation setting attribute name is not UTF-8"))?;
        if !seen.insert(local.to_string()) {
            return Err(invalid(format!("duplicate presentation:{local} attribute")));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        validate_text(&value, local, true)?;
        match local {
            "animations" => {
                settings.animations = Some(PresentationFeatureState::parse(local, &value)?)
            },
            "endless" => settings.endless = Some(parse_bool(local, &value)?),
            "force-manual" => settings.force_manual = Some(parse_bool(local, &value)?),
            "full-screen" => settings.full_screen = Some(parse_bool(local, &value)?),
            "mouse-as-pen" => settings.mouse_as_pen = Some(parse_bool(local, &value)?),
            "mouse-visible" => settings.mouse_visible = Some(parse_bool(local, &value)?),
            "pause" => {
                validate_duration(&value)?;
                settings.pause = Some(value);
            },
            "show" => settings.show = Some(value),
            "show-end-of-presentation-slide" => {
                settings.show_end_of_presentation_slide = Some(parse_bool(local, &value)?)
            },
            "show-logo" => settings.show_logo = Some(parse_bool(local, &value)?),
            "start-page" => settings.start_page = Some(value),
            "start-with-navigator" => {
                settings.start_with_navigator = Some(parse_bool(local, &value)?)
            },
            "stay-on-top" => settings.stay_on_top = Some(parse_bool(local, &value)?),
            "transition-on-click" => {
                settings.transition_on_click = Some(PresentationFeatureState::parse(local, &value)?)
            },
            _ => {
                return Err(invalid(format!(
                    "unsupported presentation:{local} attribute"
                )));
            },
        }
    }
    Ok(())
}

fn parse_custom_show(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<CustomPresentationShow> {
    let mut name = None;
    let mut pages = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(found) if found.as_ref() == PRESENTATION_NAMESPACE)
        {
            return Err(invalid("unsupported presentation:show attribute namespace"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        match local_name.as_ref() {
            b"name" if name.is_none() => name = Some(value),
            b"pages" if pages.is_none() => pages = Some(parse_pages(&value)?),
            b"name" | b"pages" => {
                return Err(invalid("duplicate presentation:show attribute"));
            },
            _ => return Err(invalid("unsupported presentation:show attribute")),
        }
    }
    CustomPresentationShow::new(
        name.ok_or_else(|| invalid("presentation:show requires presentation:name"))?,
        pages.ok_or_else(|| invalid("presentation:show requires presentation:pages"))?,
    )
}

fn parse_pages(value: &str) -> Result<Vec<String>> {
    validate_text(value, "presentation:pages", false)?;
    let pages = value
        .split(',')
        .map(str::trim)
        .map(str::to_string)
        .collect::<Vec<_>>();
    if pages.is_empty() || pages.iter().any(String::is_empty) {
        return Err(invalid("presentation:pages contains an empty page name"));
    }
    if pages.len() > MAX_CUSTOM_SHOW_PAGES {
        return Err(invalid("custom presentation show exceeds 65536 pages"));
    }
    for page in &pages {
        validate_text(page, "presentation:pages item", false)?;
    }
    Ok(pages)
}

fn validate_custom_show(show: &CustomPresentationShow) -> Result<()> {
    validate_text(&show.name, "presentation:name", false)?;
    if show.pages.is_empty() {
        return Err(invalid(
            "custom presentation show requires at least one page",
        ));
    }
    if show.pages.len() > MAX_CUSTOM_SHOW_PAGES {
        return Err(invalid("custom presentation show exceeds 65536 pages"));
    }
    for page in &show.pages {
        validate_text(page, "presentation:pages item", false)?;
        if page.contains(',') {
            return Err(invalid(
                "custom presentation page names cannot contain commas",
            ));
        }
    }
    Ok(())
}

fn validate_duration(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > 128 || !value.starts_with('P') {
        return Err(invalid("presentation:pause is not an XML Schema duration"));
    }
    let bytes = value.as_bytes();
    let mut index = 1usize;
    let mut time = false;
    let mut saw_time_value = false;
    let mut components = 0usize;
    let mut last_order = 0u8;
    while index < bytes.len() {
        if bytes[index] == b'T' {
            if time {
                return Err(invalid("presentation:pause contains duplicate 'T'"));
            }
            time = true;
            index += 1;
            continue;
        }
        let start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if index == start {
            return Err(invalid(
                "presentation:pause has an invalid duration component",
            ));
        }
        let mut decimal = false;
        if index < bytes.len() && bytes[index] == b'.' {
            decimal = true;
            index += 1;
            let fraction = index;
            while index < bytes.len() && bytes[index].is_ascii_digit() {
                index += 1;
            }
            if fraction == index {
                return Err(invalid("presentation:pause has an empty seconds fraction"));
            }
        }
        let Some(&designator) = bytes.get(index) else {
            return Err(invalid(
                "presentation:pause duration component lacks a designator",
            ));
        };
        index += 1;
        let order = match (time, designator) {
            (false, b'Y') => 1,
            (false, b'M') => 2,
            (false, b'D') => 3,
            (true, b'H') => 4,
            (true, b'M') => 5,
            (true, b'S') => 6,
            _ => {
                return Err(invalid(
                    "presentation:pause has an invalid duration designator",
                ));
            },
        };
        if decimal && designator != b'S' {
            return Err(invalid("only presentation:pause seconds may be fractional"));
        }
        if order <= last_order {
            return Err(invalid(
                "presentation:pause duration components are out of order",
            ));
        }
        last_order = order;
        components += 1;
        saw_time_value |= time;
    }
    if components == 0 || (time && !saw_time_value) {
        return Err(invalid("presentation:pause is an incomplete duration"));
    }
    Ok(())
}

fn parse_bool(attribute: &str, value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!(
            "presentation:{attribute} is not an XML Schema boolean"
        ))),
    }
}

fn validate_text(value: &str, description: &str, allow_empty: bool) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("{description} exceeds 1 MiB")));
    }
    if !allow_empty && value.is_empty() {
        return Err(invalid(format!("{description} cannot be empty")));
    }
    if value.chars().any(|character| {
        character == '\0' || (character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Err(invalid(format!(
            "{description} contains invalid XML characters"
        )));
    }
    Ok(())
}

fn element_is(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn end_is(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesEnd<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn write_bool(output: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        write_text(output, name, Some(if value { "true" } else { "false" }));
    }
}

fn write_state(output: &mut String, name: &str, value: Option<PresentationFeatureState>) {
    if let Some(value) = value {
        write_text(output, name, Some(value.as_str()));
    }
}

fn write_text(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push_str(" presentation:");
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!("presentation settings XML parsing error: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{MutablePresentation, Presentation, PresentationBuilder};

    const PREFIX: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><o:body><o:presentation>"#;
    const SUFFIX: &str = "</o:presentation></o:body></o:document-content>";

    #[test]
    fn parses_and_writes_all_presentation_settings() {
        let xml = format!(
            r#"{PREFIX}<p:settings p:animations="disabled" p:endless="1" p:force-manual="false" p:full-screen="true" p:mouse-as-pen="false" p:mouse-visible="true" p:pause="PT1M2.5S" p:show="Short" p:show-end-of-presentation-slide="true" p:show-logo="false" p:start-page="page1" p:start-with-navigator="true" p:stay-on-top="false" p:transition-on-click="enabled"><p:show p:name="Short" p:pages="page1,page3"/></p:settings>{SUFFIX}"#
        );
        let settings = parse_presentation_settings(&xml).unwrap().unwrap();
        assert_eq!(
            settings.animations,
            Some(PresentationFeatureState::Disabled)
        );
        assert_eq!(settings.pause.as_deref(), Some("PT1M2.5S"));
        assert_eq!(settings.custom_shows[0].pages, ["page1", "page3"]);
        let written = write_presentation_settings(Some(&settings)).unwrap();
        assert!(written.contains("presentation:animations=\"disabled\""));
        assert!(written.contains("presentation:pages=\"page1,page3\""));
        assert_eq!(
            parse_presentation_settings(&format!("{PREFIX}{written}{SUFFIX}"))
                .unwrap()
                .unwrap(),
            settings
        );
    }

    #[test]
    fn builder_and_mutable_round_trip_inert_settings() {
        let mut settings = PresentationSettings {
            endless: Some(true),
            pause: Some("PT15S".to_string()),
            show: Some("Executive".to_string()),
            ..PresentationSettings::default()
        };
        settings
            .custom_shows
            .push(CustomPresentationShow::new("Executive", vec!["page1".to_string()]).unwrap());
        let mut builder = PresentationBuilder::new();
        builder.set_settings(Some(settings.clone())).unwrap();
        builder.add_slide_with_title("Title", "Body").unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(presentation.settings().unwrap(), Some(settings.clone()));

        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        assert_eq!(mutable.settings(), Some(&settings));
        let reparsed = Presentation::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.settings().unwrap(), Some(settings));
    }

    #[test]
    fn rejects_invalid_placement_values_structure_and_references() {
        for xml in [
            format!(r#"{PREFIX}<p:settings p:endless="yes"/>{SUFFIX}"#),
            format!(r#"{PREFIX}<p:settings p:pause="P"/>{SUFFIX}"#),
            format!(r#"{PREFIX}<p:settings><p:show p:name="x"/></p:settings>{SUFFIX}"#),
            format!(r#"{PREFIX}<p:settings p:show="missing"/>{SUFFIX}"#),
            format!(
                r#"{PREFIX}<p:settings><p:show p:name="x" p:pages="page1"/><p:show p:name="x" p:pages="page2"/></p:settings>{SUFFIX}"#
            ),
            format!(
                r#"{PREFIX}<p:settings><p:show p:name="x" p:pages="page1"><p:show p:name="nested" p:pages="page2"/></p:show></p:settings>{SUFFIX}"#
            ),
        ] {
            assert!(parse_presentation_settings(&xml).is_err(), "accepted {xml}");
        }
        let outside = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><p:settings/>{PREFIX}{SUFFIX}</o:document-content>"#
        );
        assert!(parse_presentation_settings(&outside).is_err());
        let active = format!(r#"{PREFIX}<!DOCTYPE x><p:settings/>{SUFFIX}"#);
        assert!(parse_presentation_settings(&active).is_err());
    }
}
