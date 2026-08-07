//! ODF presentation slide transition properties.

use litchi_core::{Error, Result};

const TRANSITION_STYLES: &[&str] = &[
    "none",
    "fade-from-left",
    "fade-from-top",
    "fade-from-right",
    "fade-from-bottom",
    "fade-from-upperleft",
    "fade-from-upperright",
    "fade-from-lowerleft",
    "fade-from-lowerright",
    "move-from-left",
    "move-from-top",
    "move-from-right",
    "move-from-bottom",
    "move-from-upperleft",
    "move-from-upperright",
    "move-from-lowerleft",
    "move-from-lowerright",
    "uncover-to-left",
    "uncover-to-top",
    "uncover-to-right",
    "uncover-to-bottom",
    "uncover-to-upperleft",
    "uncover-to-upperright",
    "uncover-to-lowerleft",
    "uncover-to-lowerright",
    "fade-to-center",
    "fade-from-center",
    "vertical-stripes",
    "horizontal-stripes",
    "clockwise",
    "counterclockwise",
    "open-vertical",
    "open-horizontal",
    "close-vertical",
    "close-horizontal",
    "wavyline-from-left",
    "wavyline-from-top",
    "wavyline-from-right",
    "wavyline-from-bottom",
    "spiralin-left",
    "spiralin-right",
    "spiralout-left",
    "spiralout-right",
    "roll-from-top",
    "roll-from-left",
    "roll-from-right",
    "roll-from-bottom",
    "stretch-from-left",
    "stretch-from-top",
    "stretch-from-right",
    "stretch-from-bottom",
    "vertical-lines",
    "horizontal-lines",
    "dissolve",
    "random",
    "vertical-checkerboard",
    "horizontal-checkerboard",
    "interlocking-horizontal-left",
    "interlocking-horizontal-right",
    "interlocking-vertical-top",
    "interlocking-vertical-bottom",
    "fly-away",
    "open",
    "close",
    "melt",
];

/// Legacy ODF transition trigger behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Type {
    /// Advance only when requested by the presenter.
    Manual,
    /// Advance after the slide duration.
    Automatic,
    /// Allow both automatic and manual advancement.
    SemiAutomatic,
}

impl Type {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "manual" => Ok(Self::Manual),
            "automatic" => Ok(Self::Automatic),
            "semi-automatic" => Ok(Self::SemiAutomatic),
            _ => Err(invalid("presentation:transition-type", value)),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::Manual => "manual",
            Self::Automatic => "automatic",
            Self::SemiAutomatic => "semi-automatic",
        }
    }
}

/// ODF transition speed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Speed {
    /// Slow transition.
    Slow,
    /// Medium transition.
    Medium,
    /// Fast transition.
    Fast,
}

impl Speed {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "slow" => Ok(Self::Slow),
            "medium" => Ok(Self::Medium),
            "fast" => Ok(Self::Fast),
            _ => Err(invalid("presentation:transition-speed", value)),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::Slow => "slow",
            Self::Medium => "medium",
            Self::Fast => "fast",
        }
    }
}

/// SMIL transition direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Direction {
    /// Play the transition in its normal direction.
    Forward,
    /// Play the transition in reverse.
    Reverse,
}

impl Direction {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "forward" => Ok(Self::Forward),
            "reverse" => Ok(Self::Reverse),
            _ => Err(invalid("smil:direction", value)),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::Forward => "forward",
            Self::Reverse => "reverse",
        }
    }
}

/// A schema-defined legacy `presentation:transition-style` value.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Style(String);

impl Style {
    /// Parse a transition style defined by ODF 1.0 through 1.2.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if TRANSITION_STYLES.contains(&value.as_str()) {
            Ok(Self(value))
        } else {
            Err(invalid("presentation:transition-style", &value))
        }
    }

    /// Return the ODF lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Return every transition style allowed by the ODF schema.
    #[must_use]
    pub fn supported_values() -> &'static [&'static str] {
        TRANSITION_STYLES
    }
}

/// How a transition sound link should be presented.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SoundShow {
    /// Open the linked resource in a new context.
    New,
    /// Replace the current context.
    Replace,
}

impl SoundShow {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "new" => Ok(Self::New),
            "replace" => Ok(Self::Replace),
            _ => Err(invalid("xlink:show", value)),
        }
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::New => "new",
            Self::Replace => "replace",
        }
    }
}

/// Sound played with a slide transition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sound {
    /// Package-relative or external sound URI.
    pub href: String,
    /// Whether to play the complete sound.
    pub play_full: Option<bool>,
    /// Whether `xlink:actuate="onRequest"` is explicitly present.
    pub actuate_on_request: bool,
    /// Optional `XLink` presentation behavior.
    pub show: Option<SoundShow>,
    /// Optional XML identifier.
    pub xml_id: Option<String>,
}

impl Sound {
    /// Create a transition sound link.
    pub fn new(href: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            play_full: None,
            actuate_on_request: false,
            show: None,
            xml_id: None,
        }
    }
}

/// Complete drawing-page transition configuration for a slide.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Transition {
    pub(crate) transition_type: Option<Type>,
    pub(crate) style: Option<Style>,
    pub(crate) speed: Option<Speed>,
    pub(crate) smil_type: Option<String>,
    pub(crate) smil_subtype: Option<String>,
    pub(crate) direction: Option<Direction>,
    pub(crate) fade_color: Option<String>,
    pub(crate) duration: Option<String>,
    pub(crate) sound: Option<Sound>,
}

impl Transition {
    /// Create an empty transition configuration.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Legacy transition trigger behavior.
    #[must_use]
    pub fn transition_type(&self) -> Option<Type> {
        self.transition_type
    }

    /// Set the legacy transition trigger behavior.
    pub fn set_transition_type(&mut self, value: Option<Type>) -> &mut Self {
        self.transition_type = value;
        self
    }

    /// Legacy ODF transition style.
    #[must_use]
    pub fn style(&self) -> Option<&Style> {
        self.style.as_ref()
    }

    /// Set the legacy ODF transition style.
    pub fn set_style(&mut self, value: Option<Style>) -> &mut Self {
        self.style = value;
        self
    }

    /// Transition speed.
    #[must_use]
    pub fn speed(&self) -> Option<Speed> {
        self.speed
    }

    /// Set the transition speed.
    pub fn set_speed(&mut self, value: Option<Speed>) -> &mut Self {
        self.speed = value;
        self
    }

    /// SMIL transition type.
    #[must_use]
    pub fn smil_type(&self) -> Option<&str> {
        self.smil_type.as_deref()
    }

    /// Set the SMIL transition type.
    pub fn set_smil_type(&mut self, value: Option<impl Into<String>>) -> &mut Self {
        self.smil_type = value.map(Into::into);
        self
    }

    /// SMIL transition subtype.
    #[must_use]
    pub fn smil_subtype(&self) -> Option<&str> {
        self.smil_subtype.as_deref()
    }

    /// Set the SMIL transition subtype.
    pub fn set_smil_subtype(&mut self, value: Option<impl Into<String>>) -> &mut Self {
        self.smil_subtype = value.map(Into::into);
        self
    }

    /// SMIL transition direction.
    #[must_use]
    pub fn direction(&self) -> Option<Direction> {
        self.direction
    }

    /// Set the SMIL transition direction.
    pub fn set_direction(&mut self, value: Option<Direction>) -> &mut Self {
        self.direction = value;
        self
    }

    /// SMIL fade color as `#RRGGBB`.
    #[must_use]
    pub fn fade_color(&self) -> Option<&str> {
        self.fade_color.as_deref()
    }

    /// Set the SMIL fade color, validating the ODF color grammar.
    pub fn set_fade_color(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        let value = value.map(Into::into);
        if let Some(color) = value.as_deref()
            && !is_color(color)
        {
            return Err(invalid("smil:fadeColor", color));
        }
        self.fade_color = value;
        Ok(self)
    }

    /// Automatic slide duration as an XML Schema duration.
    #[must_use]
    pub fn duration(&self) -> Option<&str> {
        self.duration.as_deref()
    }

    /// Set the automatic slide duration, validating its complete lexical form.
    pub fn set_duration(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        let value = value.map(Into::into);
        if let Some(duration) = value.as_deref()
            && !is_xsd_duration(duration)
        {
            return Err(invalid("presentation:duration", duration));
        }
        self.duration = value;
        Ok(self)
    }

    /// Transition sound.
    #[must_use]
    pub fn sound(&self) -> Option<&Sound> {
        self.sound.as_ref()
    }

    /// Set or clear the transition sound.
    pub fn set_sound(&mut self, value: Option<Sound>) -> &mut Self {
        self.sound = value;
        self
    }

    pub(crate) fn is_empty(&self) -> bool {
        self == &Self::default()
    }

    pub(crate) fn inherit_from(&mut self, parent: &Self) {
        if self.transition_type.is_none() {
            self.transition_type = parent.transition_type;
        }
        if self.style.is_none() {
            self.style.clone_from(&parent.style);
        }
        if self.speed.is_none() {
            self.speed = parent.speed;
        }
        if self.smil_type.is_none() {
            self.smil_type.clone_from(&parent.smil_type);
        }
        if self.smil_subtype.is_none() {
            self.smil_subtype.clone_from(&parent.smil_subtype);
        }
        if self.direction.is_none() {
            self.direction = parent.direction;
        }
        if self.fade_color.is_none() {
            self.fade_color.clone_from(&parent.fade_color);
        }
        if self.duration.is_none() {
            self.duration.clone_from(&parent.duration);
        }
        if self.sound.is_none() {
            self.sound.clone_from(&parent.sound);
        }
    }
}

pub(crate) fn is_color(value: &str) -> bool {
    value.len() == 7
        && value.starts_with('#')
        && value.as_bytes()[1..].iter().all(u8::is_ascii_hexdigit)
}

pub(crate) fn is_xsd_duration(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'-'));
    if bytes.get(index) != Some(&b'P') {
        return false;
    }
    index += 1;
    let mut any = false;
    any |= consume_integer_component(bytes, &mut index, b'Y');
    any |= consume_integer_component(bytes, &mut index, b'M');
    any |= consume_integer_component(bytes, &mut index, b'D');
    if bytes.get(index) == Some(&b'T') {
        index += 1;
        let mut any_time = false;
        any_time |= consume_integer_component(bytes, &mut index, b'H');
        any_time |= consume_integer_component(bytes, &mut index, b'M');
        any_time |= consume_seconds(bytes, &mut index);
        if !any_time {
            return false;
        }
        any = true;
    }
    any && index == bytes.len()
}

fn consume_integer_component(bytes: &[u8], index: &mut usize, suffix: u8) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end > *index && bytes.get(end) == Some(&suffix) {
        *index = end + 1;
        true
    } else {
        false
    }
}

fn consume_seconds(bytes: &[u8], index: &mut usize) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end == *index {
        return false;
    }
    if bytes.get(end) == Some(&b'.') {
        end += 1;
        let start = end;
        while bytes.get(end).is_some_and(u8::is_ascii_digit) {
            end += 1;
        }
        if end == start {
            return false;
        }
    }
    if bytes.get(end) == Some(&b'S') {
        *index = end + 1;
        true
    } else {
        false
    }
}

fn invalid(attribute: &str, value: &str) -> Error {
    Error::InvalidFormat(format!("invalid {attribute} value '{value}'"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_every_schema_transition_style() {
        for value in Style::supported_values() {
            assert_eq!(Style::new(*value).unwrap().as_str(), *value);
        }
        assert!(Style::new("not-a-transition").is_err());
    }

    #[test]
    fn validates_color_and_duration_without_partial_parses() {
        let mut transition = Transition::new();
        transition.set_fade_color(Some("#aB09fF")).unwrap();
        transition.set_duration(Some("P1Y2M3DT4H5M6.25S")).unwrap();
        assert!(transition.set_fade_color(Some("red")).is_err());
        assert!(transition.set_duration(Some("PT1.S")).is_err());
        assert_eq!(transition.fade_color(), Some("#aB09fF"));
        assert_eq!(transition.duration(), Some("P1Y2M3DT4H5M6.25S"));
    }

    #[test]
    fn child_values_override_inherited_transition_values() {
        let mut parent = Transition::new();
        parent
            .set_transition_type(Some(Type::Automatic))
            .set_speed(Some(Speed::Slow));
        parent.set_duration(Some("PT5S")).unwrap();

        let mut child = Transition::new();
        child.set_speed(Some(Speed::Fast));
        child.inherit_from(&parent);
        assert_eq!(child.transition_type(), Some(Type::Automatic));
        assert_eq!(child.speed(), Some(Speed::Fast));
        assert_eq!(child.duration(), Some("PT5S"));
    }
}
