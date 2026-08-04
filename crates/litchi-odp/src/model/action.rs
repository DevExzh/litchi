//! Inert hyperlinks and event bindings attached to presentation shapes.

use super::media::{validate_bounded_xml_value, validate_href, validate_ncname};
use super::{TransitionSound, TransitionSpeed};
use litchi_core::{Error, Result, xml::escape_xml};

const PRESENTATION_ACTIONS: &[&str] = &[
    "none",
    "previous-page",
    "next-page",
    "first-page",
    "last-page",
    "hide",
    "stop",
    "execute",
    "show",
    "verb",
    "fade-out",
    "sound",
    "last-visited-page",
];

const PRESENTATION_EFFECTS: &[&str] = &[
    "none",
    "fade",
    "move",
    "stripes",
    "open",
    "close",
    "dissolve",
    "wavyline",
    "random",
    "lines",
    "laser",
    "appear",
    "hide",
    "move-short",
    "checkerboard",
    "rotate",
    "stretch",
];

const PRESENTATION_DIRECTIONS: &[&str] = &[
    "none",
    "from-left",
    "from-top",
    "from-right",
    "from-bottom",
    "from-center",
    "from-upper-left",
    "from-upper-right",
    "from-lower-left",
    "from-lower-right",
    "to-left",
    "to-top",
    "to-right",
    "to-bottom",
    "to-upper-left",
    "to-upper-right",
    "to-lower-right",
    "to-lower-left",
    "path",
    "spiral-inward-left",
    "spiral-inward-right",
    "spiral-outward-left",
    "spiral-outward-right",
    "vertical",
    "horizontal",
    "to-center",
    "clockwise",
    "counter-clockwise",
];

fn invalid(name: &str, value: &str) -> Error {
    Error::InvalidFormat(format!("invalid {name} value '{value}'"))
}

/// Action requested by an ODF presentation event listener.
///
/// Litchi preserves these requests as metadata and never invokes them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Action {
    None,
    PreviousPage,
    NextPage,
    FirstPage,
    LastPage,
    Hide,
    Stop,
    Execute,
    Show,
    Verb,
    FadeOut,
    Sound,
    LastVisitedPage,
}

impl Action {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "previous-page" => Ok(Self::PreviousPage),
            "next-page" => Ok(Self::NextPage),
            "first-page" => Ok(Self::FirstPage),
            "last-page" => Ok(Self::LastPage),
            "hide" => Ok(Self::Hide),
            "stop" => Ok(Self::Stop),
            "execute" => Ok(Self::Execute),
            "show" => Ok(Self::Show),
            "verb" => Ok(Self::Verb),
            "fade-out" => Ok(Self::FadeOut),
            "sound" => Ok(Self::Sound),
            "last-visited-page" => Ok(Self::LastVisitedPage),
            _ => Err(invalid("presentation:action", value)),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::PreviousPage => "previous-page",
            Self::NextPage => "next-page",
            Self::FirstPage => "first-page",
            Self::LastPage => "last-page",
            Self::Hide => "hide",
            Self::Stop => "stop",
            Self::Execute => "execute",
            Self::Show => "show",
            Self::Verb => "verb",
            Self::FadeOut => "fade-out",
            Self::Sound => "sound",
            Self::LastVisitedPage => "last-visited-page",
        }
    }

    /// Return every schema-defined action value.
    pub fn supported_values() -> &'static [&'static str] {
        PRESENTATION_ACTIONS
    }
}

/// Visual effect used by a presentation event action.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Effect(String);

impl Effect {
    /// Create a schema-defined presentation effect.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if PRESENTATION_EFFECTS.contains(&value.as_str()) {
            Ok(Self(value))
        } else {
            Err(invalid("presentation:effect", &value))
        }
    }

    /// Return the ODF lexical value.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Return every schema-defined value.
    pub fn supported_values() -> &'static [&'static str] {
        PRESENTATION_EFFECTS
    }
}

/// Direction used by a presentation event effect.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EffectDirection(String);

impl EffectDirection {
    /// Create a schema-defined presentation effect direction.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if PRESENTATION_DIRECTIONS.contains(&value.as_str()) {
            Ok(Self(value))
        } else {
            Err(invalid("presentation:direction", &value))
        }
    }

    /// Return the ODF lexical value.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Return every schema-defined value.
    pub fn supported_values() -> &'static [&'static str] {
        PRESENTATION_DIRECTIONS
    }
}

/// Inert presentation action metadata attached to a shape event.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EventListener {
    pub event_name: String,
    pub action: Action,
    pub effect: Option<Effect>,
    pub direction: Option<EffectDirection>,
    pub speed: Option<TransitionSpeed>,
    pub start_scale: Option<String>,
    /// Optional action target. Its serialized XLink type is always `simple`.
    pub href: Option<String>,
    pub show_embed: bool,
    pub actuate_on_request: bool,
    pub verb: Option<u64>,
    pub sound: Option<TransitionSound>,
}

impl EventListener {
    /// Create a presentation action binding.
    pub fn new(event_name: impl Into<String>, action: Action) -> Result<Self> {
        let event_name = event_name.into();
        validate_bounded_xml_value(&event_name, "presentation event name")?;
        Ok(Self {
            event_name,
            action,
            effect: None,
            direction: None,
            speed: None,
            start_scale: None,
            href: None,
            show_embed: false,
            actuate_on_request: false,
            verb: None,
            sound: None,
        })
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_bounded_xml_value(&self.event_name, "presentation event name")?;
        if let Some(start_scale) = &self.start_scale {
            validate_percent(start_scale)?;
        }
        if let Some(href) = &self.href {
            validate_href(href)?;
        } else if self.show_embed || self.actuate_on_request {
            return Err(Error::InvalidFormat(
                "presentation action XLink behavior requires an href".to_string(),
            ));
        }
        if let Some(sound) = &self.sound {
            validate_transition_sound(sound, "presentation action sound")?;
        }
        Ok(())
    }
}

/// Inert script binding attached to a shape event.
///
/// Exactly one of `macro_name` and `href` must be present. Litchi never executes
/// either form.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScriptEventListener {
    pub event_name: String,
    pub language: String,
    pub macro_name: Option<String>,
    pub href: Option<String>,
    pub actuate_on_request: bool,
}

impl ScriptEventListener {
    /// Create an inert macro-name binding.
    pub fn macro_binding(
        event_name: impl Into<String>,
        language: impl Into<String>,
        macro_name: impl Into<String>,
    ) -> Result<Self> {
        let listener = Self {
            event_name: event_name.into(),
            language: language.into(),
            macro_name: Some(macro_name.into()),
            href: None,
            actuate_on_request: false,
        };
        listener.validate()?;
        Ok(listener)
    }

    /// Create an inert external script reference.
    pub fn external_binding(
        event_name: impl Into<String>,
        language: impl Into<String>,
        href: impl Into<String>,
    ) -> Result<Self> {
        let listener = Self {
            event_name: event_name.into(),
            language: language.into(),
            macro_name: None,
            href: Some(href.into()),
            actuate_on_request: false,
        };
        listener.validate()?;
        Ok(listener)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_bounded_xml_value(&self.event_name, "script event name")?;
        validate_bounded_xml_value(&self.language, "script language")?;
        if self.macro_name.is_some() == self.href.is_some() {
            return Err(Error::InvalidFormat(
                "script event listener requires exactly one macro name or href".to_string(),
            ));
        }
        if let Some(macro_name) = &self.macro_name {
            validate_bounded_xml_value(macro_name, "script macro name")?;
        }
        if let Some(href) = &self.href {
            validate_href(href)?;
        } else if self.actuate_on_request {
            return Err(Error::InvalidFormat(
                "script listener xlink:actuate requires an href".to_string(),
            ));
        }
        Ok(())
    }
}

/// An inert event listener attached to a drawing shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ShapeEventListener {
    Action(Box<EventListener>),
    Script(ScriptEventListener),
}

/// `xlink:show` behavior for a drawing hyperlink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyperlinkShow {
    New,
    Replace,
}

impl HyperlinkShow {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "new" => Ok(Self::New),
            "replace" => Ok(Self::Replace),
            _ => Err(invalid("draw:a xlink:show", value)),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::New => "new",
            Self::Replace => "replace",
        }
    }
}

/// Hyperlink metadata wrapping exactly one drawing shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingHyperlink {
    href: String,
    actuate_on_request: bool,
    show: Option<HyperlinkShow>,
    target_frame_name: Option<String>,
    name: Option<String>,
    title: Option<String>,
    server_map: Option<bool>,
    xml_id: Option<String>,
}

impl DrawingHyperlink {
    /// Create a hyperlink. The serialized XLink type is always `simple`.
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let href = href.into();
        validate_href(&href)?;
        Ok(Self {
            href,
            actuate_on_request: false,
            show: None,
            target_frame_name: None,
            name: None,
            title: None,
            server_map: None,
            xml_id: None,
        })
    }

    pub fn href(&self) -> &str {
        &self.href
    }

    pub fn set_href(&mut self, href: impl Into<String>) -> Result<()> {
        let href = href.into();
        validate_href(&href)?;
        self.href = href;
        Ok(())
    }

    pub fn actuate_on_request(&self) -> bool {
        self.actuate_on_request
    }

    pub fn set_actuate_on_request(&mut self, value: bool) {
        self.actuate_on_request = value;
    }

    pub fn show(&self) -> Option<HyperlinkShow> {
        self.show
    }

    pub fn set_show(&mut self, value: Option<HyperlinkShow>) {
        self.show = value;
    }

    pub fn target_frame_name(&self) -> Option<&str> {
        self.target_frame_name.as_deref()
    }

    pub fn set_target_frame_name(&mut self, value: Option<String>) -> Result<()> {
        validate_optional(&value, "hyperlink target frame name")?;
        self.target_frame_name = value;
        Ok(())
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn set_name(&mut self, value: Option<String>) -> Result<()> {
        validate_optional(&value, "hyperlink name")?;
        self.name = value;
        Ok(())
    }

    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    pub fn set_title(&mut self, value: Option<String>) -> Result<()> {
        validate_optional(&value, "hyperlink title")?;
        self.title = value;
        Ok(())
    }

    pub fn server_map(&self) -> Option<bool> {
        self.server_map
    }

    pub fn set_server_map(&mut self, value: Option<bool>) {
        self.server_map = value;
    }

    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    pub fn set_xml_id(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = &value {
            validate_ncname(value, "hyperlink XML ID")?;
        }
        self.xml_id = value;
        Ok(())
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_href(&self.href)?;
        validate_optional(&self.target_frame_name, "hyperlink target frame name")?;
        validate_optional(&self.name, "hyperlink name")?;
        validate_optional(&self.title, "hyperlink title")?;
        if let Some(xml_id) = &self.xml_id {
            validate_ncname(xml_id, "hyperlink XML ID")?;
        }
        Ok(())
    }

    pub(crate) fn write_open_xml(&self, output: &mut String) -> Result<()> {
        self.validate()?;
        output.push_str("<draw:a xlink:type=\"simple\" xlink:href=\"");
        output.push_str(&escape_xml(&self.href));
        output.push('"');
        if self.actuate_on_request {
            output.push_str(" xlink:actuate=\"onRequest\"");
        }
        if let Some(show) = self.show {
            push_attribute(output, "xlink:show", show.as_str());
        }
        if let Some(value) = &self.target_frame_name {
            push_attribute(output, "office:target-frame-name", value);
        }
        if let Some(value) = &self.name {
            push_attribute(output, "office:name", value);
        }
        if let Some(value) = &self.title {
            push_attribute(output, "office:title", value);
        }
        if let Some(value) = self.server_map {
            push_attribute(
                output,
                "office:server-map",
                if value { "true" } else { "false" },
            );
        }
        if let Some(value) = &self.xml_id {
            push_attribute(output, "xml:id", value);
        }
        output.push('>');
        Ok(())
    }
}

pub(crate) fn validate_event_listeners(listeners: &[ShapeEventListener]) -> Result<()> {
    if listeners.len() > 4096 {
        return Err(Error::InvalidFormat(
            "ODP shape exceeds 4096 event listeners".to_string(),
        ));
    }
    for listener in listeners {
        match listener {
            ShapeEventListener::Action(listener) => listener.validate()?,
            ShapeEventListener::Script(listener) => listener.validate()?,
        }
    }
    Ok(())
}

pub(crate) fn write_event_listeners(
    output: &mut String,
    listeners: &[ShapeEventListener],
) -> Result<()> {
    validate_event_listeners(listeners)?;
    if listeners.is_empty() {
        return Ok(());
    }
    output.push_str("<office:event-listeners>");
    for listener in listeners {
        match listener {
            ShapeEventListener::Script(listener) => {
                output.push_str("<script:event-listener");
                push_attribute(output, "script:event-name", &listener.event_name);
                push_attribute(output, "script:language", &listener.language);
                if let Some(value) = &listener.macro_name {
                    push_attribute(output, "script:macro-name", value);
                }
                if let Some(value) = &listener.href {
                    push_attribute(output, "xlink:type", "simple");
                    push_attribute(output, "xlink:href", value);
                    if listener.actuate_on_request {
                        push_attribute(output, "xlink:actuate", "onRequest");
                    }
                }
                output.push_str("/>");
            },
            ShapeEventListener::Action(listener) => {
                output.push_str("<presentation:event-listener");
                push_attribute(output, "script:event-name", &listener.event_name);
                push_attribute(output, "presentation:action", listener.action.as_str());
                if let Some(value) = &listener.effect {
                    push_attribute(output, "presentation:effect", value.as_str());
                }
                if let Some(value) = &listener.direction {
                    push_attribute(output, "presentation:direction", value.as_str());
                }
                if let Some(value) = listener.speed {
                    push_attribute(output, "presentation:speed", value.as_str());
                }
                if let Some(value) = &listener.start_scale {
                    push_attribute(output, "presentation:start-scale", value);
                }
                if let Some(value) = &listener.href {
                    push_attribute(output, "xlink:type", "simple");
                    push_attribute(output, "xlink:href", value);
                    if listener.show_embed {
                        push_attribute(output, "xlink:show", "embed");
                    }
                    if listener.actuate_on_request {
                        push_attribute(output, "xlink:actuate", "onRequest");
                    }
                }
                if let Some(value) = listener.verb {
                    push_attribute(output, "presentation:verb", &value.to_string());
                }
                if let Some(sound) = &listener.sound {
                    output.push('>');
                    write_sound(output, sound)?;
                    output.push_str("</presentation:event-listener>");
                } else {
                    output.push_str("/>");
                }
            },
        }
    }
    output.push_str("</office:event-listeners>");
    Ok(())
}

fn write_sound(output: &mut String, sound: &TransitionSound) -> Result<()> {
    validate_transition_sound(sound, "presentation action sound")?;
    output.push_str("<presentation:sound xlink:type=\"simple\"");
    push_attribute(output, "xlink:href", &sound.href);
    if sound.actuate_on_request {
        push_attribute(output, "xlink:actuate", "onRequest");
    }
    if let Some(show) = sound.show {
        push_attribute(output, "xlink:show", show.as_str());
    }
    if let Some(play_full) = sound.play_full {
        push_attribute(
            output,
            "presentation:play-full",
            if play_full { "true" } else { "false" },
        );
    }
    if let Some(xml_id) = &sound.xml_id {
        push_attribute(output, "xml:id", xml_id);
    }
    output.push_str("/>");
    Ok(())
}

fn validate_transition_sound(sound: &TransitionSound, description: &str) -> Result<()> {
    validate_href(&sound.href)?;
    if let Some(xml_id) = &sound.xml_id {
        validate_ncname(xml_id, &format!("{description} XML ID"))?;
    }
    Ok(())
}

fn validate_optional(value: &Option<String>, description: &str) -> Result<()> {
    if let Some(value) = value {
        validate_bounded_xml_value(value, description)?;
    }
    Ok(())
}

fn validate_percent(value: &str) -> Result<()> {
    validate_bounded_xml_value(value, "presentation start scale")?;
    let mut number = value
        .strip_suffix('%')
        .filter(|number| !number.is_empty())
        .ok_or_else(|| invalid("presentation:start-scale", value))?;
    if let Some(unsigned) = number.strip_prefix('-') {
        number = unsigned;
    }
    let mut parts = number.split('.');
    let integer = parts.next().expect("split always yields one part");
    let fraction = parts.next();
    if parts.next().is_some()
        || (integer.is_empty() && fraction.is_none_or(str::is_empty))
        || !integer.bytes().all(|byte| byte.is_ascii_digit())
        || fraction.is_some_and(|part| !part.bytes().all(|byte| byte.is_ascii_digit()))
    {
        return Err(invalid("presentation:start-scale", value));
    }
    Ok(())
}

fn push_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exposes_all_schema_effect_values() {
        for value in Action::supported_values() {
            assert_eq!(Action::parse(value).unwrap().as_str(), *value);
        }
        for value in Effect::supported_values() {
            assert_eq!(Effect::new(*value).unwrap().as_str(), *value);
        }
        for value in EffectDirection::supported_values() {
            assert_eq!(EffectDirection::new(*value).unwrap().as_str(), *value);
        }
    }

    #[test]
    fn rejects_ambiguous_script_bindings_and_invalid_percentages() {
        let mut listener =
            ScriptEventListener::macro_binding("dom:click", "ooo:script", "Standard.Module1.Main")
                .unwrap();
        listener.href = Some("Scripts/main.js".to_string());
        assert!(listener.validate().is_err());

        let mut action = EventListener::new("dom:click", Action::FadeOut).unwrap();
        action.start_scale = Some("fifty".to_string());
        assert!(action.validate().is_err());
    }

    #[test]
    fn serializes_event_metadata_without_executing_it() {
        let mut action = EventListener::new("dom:click", Action::Execute).unwrap();
        action.href = Some("https://example.invalid/app".to_string());
        action.show_embed = true;
        let listeners = vec![ShapeEventListener::Action(Box::new(action))];
        let mut xml = String::new();
        write_event_listeners(&mut xml, &listeners).unwrap();
        assert!(xml.contains("presentation:action=\"execute\""));
        assert!(xml.contains("xlink:href=\"https://example.invalid/app\""));
    }
}
