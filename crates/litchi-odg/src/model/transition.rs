//! Inert drawing-page transition metadata.
//!
//! ODF stores page transitions on the drawing-page style rather than directly
//! on `draw:page`.  The value in this module is a small semantic projection;
//! the original style XML remains authoritative for publication, so unknown
//! attributes and producer spelling stay untouched.

use litchi_core::{Error, Result};

const MAX_VALUE_BYTES: usize = 1 << 20;
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

/// A bounded inert transition sound link.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Sound {
    href: String,
    play_full: Option<bool>,
    actuate_on_request: bool,
    show: Option<String>,
    xml_id: Option<String>,
}

impl Sound {
    /// Creates a transition sound reference. The URI is never fetched.
    pub fn new(href: impl Into<String>) -> Result<Self> {
        Ok(Self {
            // `xlink:href` is an ODF anyIRI/anyURI value.  Empty is a valid
            // same-document reference; package resource closure is checked
            // by the owning transaction when publishing.
            href: bounded_string(href.into(), "transition sound URI")?,
            play_full: None,
            actuate_on_request: false,
            show: None,
            xml_id: None,
        })
    }

    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    #[must_use]
    pub const fn play_full(&self) -> Option<bool> {
        self.play_full
    }

    #[must_use]
    pub const fn with_play_full(mut self, value: Option<bool>) -> Self {
        self.play_full = value;
        self
    }

    #[must_use]
    pub const fn actuate_on_request(&self) -> bool {
        self.actuate_on_request
    }

    #[must_use]
    pub const fn with_actuate_on_request(mut self, value: bool) -> Self {
        self.actuate_on_request = value;
        self
    }

    #[must_use]
    pub fn show(&self) -> Option<&str> {
        self.show.as_deref()
    }

    pub fn with_show(mut self, value: Option<impl Into<String>>) -> Result<Self> {
        self.show = value
            .map(Into::into)
            .map(|value| bounded(value, "xlink:show"))
            .transpose()?;
        if self
            .show
            .as_deref()
            .is_some_and(|value| !matches!(value, "new" | "replace"))
        {
            return Err(invalid("xlink:show"));
        }
        Ok(self)
    }

    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    pub fn with_xml_id(mut self, value: Option<impl Into<String>>) -> Result<Self> {
        self.xml_id = value
            .map(Into::into)
            .map(|value| bounded(value, "xml:id"))
            .transpose()?;
        if self
            .xml_id
            .as_deref()
            .is_some_and(|value| !is_xml_id(value))
        {
            return Err(invalid("xml:id"));
        }
        Ok(self)
    }
}

/// Inert ODF drawing-page transition metadata.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Transition {
    transition_type: Option<String>,
    style: Option<String>,
    speed: Option<String>,
    smil_type: Option<String>,
    smil_subtype: Option<String>,
    direction: Option<String>,
    fade_color: Option<String>,
    duration: Option<String>,
    sound: Option<Sound>,
}

impl Transition {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            transition_type: None,
            style: None,
            speed: None,
            smil_type: None,
            smil_subtype: None,
            direction: None,
            fade_color: None,
            duration: None,
            sound: None,
        }
    }

    #[must_use]
    pub fn transition_type(&self) -> Option<&str> {
        self.transition_type.as_deref()
    }

    pub fn set_transition_type(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        self.transition_type = checked_choice(
            value,
            "presentation:transition-type",
            &["manual", "automatic", "semi-automatic"],
        )?;
        Ok(self)
    }

    #[must_use]
    pub fn style(&self) -> Option<&str> {
        self.style.as_deref()
    }

    pub fn set_style(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        let value = value.map(Into::into);
        if value
            .as_deref()
            .is_some_and(|value| !TRANSITION_STYLES.contains(&value))
        {
            return Err(invalid("presentation:transition-style"));
        }
        self.style = value;
        Ok(self)
    }

    #[must_use]
    pub fn supported_styles() -> &'static [&'static str] {
        TRANSITION_STYLES
    }

    #[must_use]
    pub fn speed(&self) -> Option<&str> {
        self.speed.as_deref()
    }

    pub fn set_speed(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        self.speed = checked_choice(
            value,
            "presentation:transition-speed",
            &["slow", "medium", "fast"],
        )?;
        Ok(self)
    }

    #[must_use]
    pub fn smil_type(&self) -> Option<&str> {
        self.smil_type.as_deref()
    }

    pub fn set_smil_type(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        self.smil_type = optional_string(value, "smil:type")?;
        Ok(self)
    }

    #[must_use]
    pub fn smil_subtype(&self) -> Option<&str> {
        self.smil_subtype.as_deref()
    }

    pub fn set_smil_subtype(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        self.smil_subtype = optional_string(value, "smil:subtype")?;
        Ok(self)
    }

    #[must_use]
    pub fn direction(&self) -> Option<&str> {
        self.direction.as_deref()
    }

    pub fn set_direction(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        self.direction = checked_choice(value, "smil:direction", &["forward", "reverse"])?;
        Ok(self)
    }

    #[must_use]
    pub fn fade_color(&self) -> Option<&str> {
        self.fade_color.as_deref()
    }

    pub fn set_fade_color(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        let value = optional_value(value, "smil:fadeColor")?;
        if value.as_deref().is_some_and(|value| {
            !(value.len() == 7
                && value.starts_with('#')
                && value.as_bytes()[1..].iter().all(u8::is_ascii_hexdigit))
        }) {
            return Err(invalid("smil:fadeColor"));
        }
        self.fade_color = value;
        Ok(self)
    }

    #[must_use]
    pub fn duration(&self) -> Option<&str> {
        self.duration.as_deref()
    }

    pub fn set_duration(&mut self, value: Option<impl Into<String>>) -> Result<&mut Self> {
        let value = optional_value(value, "presentation:duration")?;
        if value
            .as_deref()
            .is_some_and(|value| !is_xsd_duration(value))
        {
            return Err(invalid("presentation:duration"));
        }
        self.duration = value;
        Ok(self)
    }

    #[must_use]
    pub fn sound(&self) -> Option<&Sound> {
        self.sound.as_ref()
    }

    pub fn set_sound(&mut self, value: Option<Sound>) -> &mut Self {
        self.sound = value;
        self
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self == &Self::new()
    }

    pub(crate) fn from_parts(
        transition_type: Option<String>,
        style: Option<String>,
        speed: Option<String>,
        smil_type: Option<String>,
        smil_subtype: Option<String>,
        direction: Option<String>,
        fade_color: Option<String>,
        duration: Option<String>,
        sound: Option<Sound>,
    ) -> Result<Self> {
        let mut value = Self::new();
        value.set_transition_type(transition_type)?;
        value.set_style(style)?;
        value.set_speed(speed)?;
        value.set_smil_type(smil_type)?;
        value.set_smil_subtype(smil_subtype)?;
        value.set_direction(direction)?;
        value.set_fade_color(fade_color)?;
        value.set_duration(duration)?;
        value.set_sound(sound);
        Ok(value)
    }

    pub(crate) fn inherit_from(&mut self, parent: &Self) {
        if self.transition_type.is_none() {
            self.transition_type = parent.transition_type.clone();
        }
        if self.style.is_none() {
            self.style = parent.style.clone();
        }
        if self.speed.is_none() {
            self.speed = parent.speed.clone();
        }
        if self.smil_type.is_none() {
            self.smil_type = parent.smil_type.clone();
        }
        if self.smil_subtype.is_none() {
            self.smil_subtype = parent.smil_subtype.clone();
        }
        if self.direction.is_none() {
            self.direction = parent.direction.clone();
        }
        if self.fade_color.is_none() {
            self.fade_color = parent.fade_color.clone();
        }
        if self.duration.is_none() {
            self.duration = parent.duration.clone();
        }
        if self.sound.is_none() {
            self.sound = parent.sound.clone();
        }
    }
}

fn optional_value<T>(value: Option<T>, field: &str) -> Result<Option<String>>
where
    T: Into<String>,
{
    value
        .map(Into::into)
        .map(|value| bounded(value, field))
        .transpose()
}

fn optional_string<T>(value: Option<T>, field: &str) -> Result<Option<String>>
where
    T: Into<String>,
{
    value
        .map(Into::into)
        .map(|value| bounded_string(value, field))
        .transpose()
}

fn checked_choice<T>(value: Option<T>, field: &str, choices: &[&str]) -> Result<Option<String>>
where
    T: Into<String>,
{
    let value = optional_value(value, field)?;
    if value
        .as_deref()
        .is_some_and(|value| !choices.contains(&value))
    {
        return Err(invalid(field));
    }
    Ok(value)
}

fn bounded(value: String, field: &str) -> Result<String> {
    if value.is_empty() || value.len() > MAX_VALUE_BYTES || value.bytes().any(|byte| byte == 0) {
        return Err(invalid(field));
    }
    Ok(value)
}

fn bounded_string(value: String, field: &str) -> Result<String> {
    if value.len() > MAX_VALUE_BYTES || value.bytes().any(|byte| byte == 0) {
        return Err(invalid(field));
    }
    Ok(value)
}

fn is_xml_id(value: &str) -> bool {
    let mut characters = value.chars();
    let Some(first) = characters.next() else {
        return false;
    };
    is_ncname_start(first) && characters.all(is_ncname_char)
}

fn is_ncname_start(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00c0}'..='\u{00d6}'
            | '\u{00d8}'..='\u{00f6}'
            | '\u{00f8}'..='\u{02ff}'
            | '\u{0370}'..='\u{037d}'
            | '\u{037f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

fn is_xsd_duration(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'-'));
    if bytes.get(index) != Some(&b'P') {
        return false;
    }
    index += 1;
    let mut any = false;
    any |= consume_integer(bytes, &mut index, b'Y');
    any |= consume_integer(bytes, &mut index, b'M');
    any |= consume_integer(bytes, &mut index, b'D');
    if bytes.get(index) == Some(&b'T') {
        index += 1;
        let mut time = false;
        time |= consume_integer(bytes, &mut index, b'H');
        time |= consume_integer(bytes, &mut index, b'M');
        time |= consume_seconds(bytes, &mut index);
        if !time {
            return false;
        }
        any = true;
    }
    // `xsd:duration` has no timezone suffix.  `Z` and numeric offsets are
    // dateTime lexical forms and must remain invalid here.
    any && index == bytes.len()
}

fn consume_integer(bytes: &[u8], index: &mut usize, suffix: u8) -> bool {
    let start = *index;
    while bytes.get(*index).is_some_and(u8::is_ascii_digit) {
        *index += 1;
    }
    if *index > start && bytes.get(*index) == Some(&suffix) {
        *index += 1;
        true
    } else {
        *index = start;
        false
    }
}

fn consume_seconds(bytes: &[u8], index: &mut usize) -> bool {
    let start = *index;
    while bytes.get(*index).is_some_and(u8::is_ascii_digit) {
        *index += 1;
    }
    let whole = *index != start;
    let mut fraction = false;
    if bytes.get(*index) == Some(&b'.') {
        *index += 1;
        let fraction_start = *index;
        while bytes.get(*index).is_some_and(u8::is_ascii_digit) {
            *index += 1;
        }
        fraction = *index != fraction_start;
    }
    if (whole || fraction) && bytes.get(*index) == Some(&b'S') {
        *index += 1;
        true
    } else {
        *index = start;
        false
    }
}

fn invalid(field: &str) -> Error {
    Error::InvalidFormat(format!("invalid ODG {field} value"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_choices_color_duration_and_bounds() {
        let mut value = Transition::new();
        value.set_transition_type(Some("automatic")).unwrap();
        value.set_style(Some("dissolve")).unwrap();
        value.set_speed(Some("fast")).unwrap();
        value.set_direction(Some("reverse")).unwrap();
        value.set_fade_color(Some("#aB09fF")).unwrap();
        value.set_duration(Some("PT2.5S")).unwrap();
        assert!(value.set_direction(Some("sideways")).is_err());
        assert!(value.set_fade_color(Some("red")).is_err());
        assert!(value.set_duration(Some("PT1.S")).is_ok());
        assert!(value.set_duration(Some("PT.5S")).is_ok());
        assert!(value.set_duration(Some("PT1S+05:30")).is_err());
        assert!(value.set_duration(Some("PT1SZ")).is_err());
        assert!(value.set_style(Some("missing")).is_err());
        assert!(value.set_smil_type(Some("")).is_ok());

        assert!(
            Sound::new("media/sound.wav")
                .unwrap()
                .with_xml_id(Some("é-2"))
                .is_ok()
        );
        assert!(
            Sound::new("media/sound.wav")
                .unwrap()
                .with_xml_id(Some("2bad"))
                .is_err()
        );
        assert_eq!(Sound::new("").unwrap().href(), "");
    }
}
