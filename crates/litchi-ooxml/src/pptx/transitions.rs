//! Slide transition effects for PowerPoint presentations.
//!
//! This module provides types and functionality for working with slide transitions,
//! including transition types, speeds, and directions.

use crate::common::xml::unqualified_attribute_value;
use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

const P14_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_NAMESPACE_BYTES: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const MARKUP_COMPATIBILITY_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Slide transition type.
///
/// Represents the various transition effects available in PowerPoint.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TransitionType {
    /// No transition
    None,
    /// Cut transition (instant change)
    Cut,
    /// Fade through black
    Fade,
    /// Push transition
    Push { direction: TransitionDirection },
    /// Wipe transition
    Wipe { direction: TransitionDirection },
    /// Split transition
    Split { direction: TransitionDirection },
    /// Reveal transition
    Reveal { direction: TransitionDirection },
    /// Random bars
    RandomBars { direction: TransitionDirection },
    /// Shape (circle, diamond, plus)
    Shape { shape_type: ShapeTransitionType },
    /// Cover transition
    Cover { direction: TransitionDirection },
    /// Uncover transition
    Uncover { direction: TransitionDirection },
    /// Dissolve transition
    Dissolve,
    /// Checkerboard
    Checker { direction: TransitionDirection },
    /// Blinds
    Blinds { direction: TransitionDirection },
    /// Clock (clockwise sweep)
    Clock { direction: ClockDirection },
    /// Zoom (in/out)
    Zoom { direction: ZoomDirection },
    /// Random transition (PowerPoint picks)
    Random,
    /// Wheel (spokes)
    Wheel { spokes: u8 },
    /// Circle transition
    Circle,
    /// Diamond transition
    Diamond,
    /// Plus transition
    Plus,
    /// Wedge transition
    Wedge,
    /// Newsflash transition
    Newsflash,
    /// Flash transition
    Flash,
    /// PowerPoint 2010 ripple transition with a fade compatibility fallback.
    Ripple { direction: RippleDirection },
    /// Strips transition
    Strips { direction: TransitionDirection },
    /// Comb transition
    Comb { direction: TransitionDirection },
    /// Other/Unknown transition
    Other(String),
}

/// Transition direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransitionDirection {
    /// Left to right
    Left,
    /// Right to left
    Right,
    /// Top to bottom
    Up,
    /// Bottom to top
    Down,
    /// From the upper-left corner
    LeftUp,
    /// From the upper-right corner
    RightUp,
    /// From the lower-left corner
    LeftDown,
    /// From the lower-right corner
    RightDown,
    /// Horizontal (left and right)
    Horizontal,
    /// Vertical (up and down)
    Vertical,
    /// From all corners inward
    In,
    /// From center outward
    Out,
}

/// Shape transition type for shape-based transitions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeTransitionType {
    /// Circle shape
    Circle,
    /// Diamond shape
    Diamond,
    /// Plus shape
    Plus,
}

/// Clock transition direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ClockDirection {
    /// Clockwise
    Clockwise,
    /// Counterclockwise
    Counterclockwise,
}

/// Zoom transition direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ZoomDirection {
    /// Zoom in
    In,
    /// Zoom out
    Out,
}

/// In/out direction used by a split transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SplitDirection {
    /// Move toward the center.
    In,
    /// Move away from the center.
    Out,
}

/// Direction used by a PowerPoint 2010 ripple transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RippleDirection {
    /// Start at the center of the slide.
    Center,
    /// Start at the upper-left corner.
    LeftUp,
    /// Start at the upper-right corner.
    RightUp,
    /// Start at the lower-left corner.
    LeftDown,
    /// Start at the lower-right corner.
    RightDown,
}

impl RippleDirection {
    fn from_xml_value(value: &str) -> Result<Self> {
        match value {
            "center" => Ok(Self::Center),
            "lu" => Ok(Self::LeftUp),
            "ru" => Ok(Self::RightUp),
            "ld" => Ok(Self::LeftDown),
            "rd" => Ok(Self::RightDown),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid PowerPoint 2010 ripple direction '{value}'"
            ))),
        }
    }

    fn to_xml_value(self) -> &'static str {
        match self {
            Self::Center => "center",
            Self::LeftUp => "lu",
            Self::RightUp => "ru",
            Self::LeftDown => "ld",
            Self::RightDown => "rd",
        }
    }
}

/// Transition speed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransitionSpeed {
    /// Slow transition (1500ms)
    Slow,
    /// Medium transition (1000ms)
    Medium,
    /// Fast transition (500ms)
    Fast,
}

impl TransitionSpeed {
    /// Get the duration in milliseconds.
    pub fn duration_ms(&self) -> u32 {
        match self {
            TransitionSpeed::Slow => 1500,
            TransitionSpeed::Medium => 1000,
            TransitionSpeed::Fast => 500,
        }
    }

    /// Create from duration in milliseconds.
    pub fn from_duration_ms(ms: u32) -> Self {
        if ms <= 700 {
            TransitionSpeed::Fast
        } else if ms <= 1200 {
            TransitionSpeed::Medium
        } else {
            TransitionSpeed::Slow
        }
    }

    /// Convert to OOXML speed value.
    pub(crate) fn to_xml_value(self) -> &'static str {
        match self {
            TransitionSpeed::Slow => "slow",
            TransitionSpeed::Medium => "med",
            TransitionSpeed::Fast => "fast",
        }
    }

    /// Parse from OOXML speed value.
    pub(crate) fn from_xml_value(value: &str) -> Self {
        match value {
            "slow" => TransitionSpeed::Slow,
            "fast" => TransitionSpeed::Fast,
            _ => TransitionSpeed::Medium,
        }
    }
}

/// Complete slide transition configuration.
///
/// Includes the transition type, speed, and timing settings.
#[derive(Debug, Clone, PartialEq)]
pub struct SlideTransition {
    /// Type of transition effect
    pub transition_type: TransitionType,
    /// Speed of the transition
    pub speed: TransitionSpeed,
    /// Duration in milliseconds (optional, overrides speed)
    pub duration_ms: Option<u32>,
    /// Whether a cut or fade transition goes through black.
    ///
    /// `None` preserves an omitted attribute and therefore the OOXML default.
    pub through_black: Option<bool>,
    /// Explicit in/out direction for a split transition.
    ///
    /// `None` preserves an omitted attribute and therefore the OOXML default.
    pub split_direction: Option<SplitDirection>,
    /// Whether to advance slide on mouse click
    pub advance_on_click: bool,
    /// Auto-advance after delay in milliseconds (None = no auto-advance)
    pub advance_after_ms: Option<u32>,
    /// Whether sound should play during transition
    pub sound: Option<TransitionSound>,
}

/// Transition sound configuration.
#[derive(Debug, Clone, PartialEq)]
pub struct TransitionSound {
    /// Sound name or built-in sound identifier
    pub name: String,
    /// Whether to loop the sound
    pub loop_sound: bool,
}

#[derive(Debug)]
struct ParsedTransitionEffect {
    transition_type: TransitionType,
    through_black: Option<bool>,
    split_direction: Option<SplitDirection>,
}

impl ParsedTransitionEffect {
    fn new(transition_type: TransitionType) -> Self {
        Self {
            transition_type,
            through_black: None,
            split_direction: None,
        }
    }
}

impl Default for SlideTransition {
    fn default() -> Self {
        Self {
            transition_type: TransitionType::None,
            speed: TransitionSpeed::Medium,
            duration_ms: None,
            through_black: None,
            split_direction: None,
            advance_on_click: true,
            advance_after_ms: None,
            sound: None,
        }
    }
}

impl SlideTransition {
    /// Create a new transition with default settings.
    pub fn new(transition_type: TransitionType) -> Self {
        Self {
            transition_type,
            ..Default::default()
        }
    }

    /// Set the transition speed.
    pub fn with_speed(mut self, speed: TransitionSpeed) -> Self {
        self.speed = speed;
        self
    }

    /// Set a custom duration in milliseconds.
    pub fn with_duration_ms(mut self, duration_ms: u32) -> Self {
        self.duration_ms = Some(duration_ms);
        self
    }

    /// Set whether a cut or fade transition goes through black.
    pub fn with_through_black(mut self, through_black: bool) -> Self {
        self.through_black = Some(through_black);
        self
    }

    /// Set the in/out direction of a split transition.
    pub fn with_split_direction(mut self, direction: SplitDirection) -> Self {
        self.split_direction = Some(direction);
        self
    }

    /// Set whether to advance on mouse click.
    pub fn with_advance_on_click(mut self, advance: bool) -> Self {
        self.advance_on_click = advance;
        self
    }

    /// Set auto-advance delay in milliseconds.
    pub fn with_advance_after_ms(mut self, delay_ms: u32) -> Self {
        self.advance_after_ms = Some(delay_ms);
        self
    }

    /// Add a sound to the transition.
    pub fn with_sound(mut self, name: String, loop_sound: bool) -> Self {
        self.sound = Some(TransitionSound { name, loop_sound });
        self
    }

    /// Get the effective duration in milliseconds.
    pub fn effective_duration_ms(&self) -> u32 {
        self.duration_ms.unwrap_or_else(|| self.speed.duration_ms())
    }

    /// Parse transition from slide XML.
    pub(crate) fn from_xml(xml: &[u8]) -> Result<Option<Self>> {
        let mut capabilities = MceCapabilities::ooxml_baseline();
        capabilities.understand_namespace(P14_NAMESPACE);
        let xml = process_markup_compatibility(xml, &capabilities, &MceLimits::default())?.xml;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut stack = Vec::new();
        let mut transition: Option<SlideTransition> = None;
        let mut selected_transition_depth = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    let depth = stack.len().checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("transition XML nesting is too deep".to_string())
                    })?;
                    let is_transition =
                        is_presentationml_name(&namespace, element.name(), b"transition");
                    if is_transition && transition.is_none() {
                        transition = Some(Self::parse_transition_attributes(
                            &element, decoder, &resolver,
                        )?);
                        selected_transition_depth = Some(depth);
                    } else if selected_transition_depth
                        .and_then(|transition_depth| transition_depth.checked_add(1))
                        == Some(depth)
                    {
                        if let Some(effect) =
                            Self::parse_transition_type(&namespace, &element, decoder)?
                            && let Some(transition) = transition.as_mut()
                        {
                            transition.apply_effect(effect);
                        }
                    }
                    stack.push(is_transition);
                },
                Event::Empty(element) => {
                    let depth = stack.len().checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("transition XML nesting is too deep".to_string())
                    })?;
                    let is_transition =
                        is_presentationml_name(&namespace, element.name(), b"transition");
                    if is_transition && transition.is_none() {
                        transition = Some(Self::parse_transition_attributes(
                            &element, decoder, &resolver,
                        )?);
                    } else if selected_transition_depth
                        .and_then(|transition_depth| transition_depth.checked_add(1))
                        == Some(depth)
                    {
                        if let Some(effect) =
                            Self::parse_transition_type(&namespace, &element, decoder)?
                            && let Some(transition) = transition.as_mut()
                        {
                            transition.apply_effect(effect);
                        }
                    }
                },
                Event::End(_) => {
                    let is_transition = stack.pop().ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid transition XML nesting".to_string())
                    })?;
                    let depth = stack.len().checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("transition XML nesting is too deep".to_string())
                    })?;
                    if is_transition && selected_transition_depth == Some(depth) {
                        selected_transition_depth = None;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(transition)
    }

    fn parse_transition_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Self> {
        let mut speed = TransitionSpeed::Medium;
        let mut legacy_duration_ms = None;
        let mut extended_duration_ms = None;
        let mut advance_on_click = true;
        let mut advance_after_ms = None;

        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
            let key = attribute.key;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            let value = value.as_ref();

            if key.prefix().is_none() {
                match key.local_name().as_ref() {
                    b"spd" => speed = TransitionSpeed::from_xml_value(value),
                    b"dur" => legacy_duration_ms = parse_duration_ms(value),
                    b"advClick" => advance_on_click = value == "1" || value == "true",
                    b"advTm" => advance_after_ms = value.parse::<u32>().ok(),
                    _ => {},
                }
            } else {
                let (namespace, _) = resolver.resolve_attribute(key);
                if is_p14_namespace(&namespace) && key.local_name().as_ref() == b"dur" {
                    extended_duration_ms = parse_duration_ms(value);
                }
            }
        }

        Ok(Self {
            transition_type: TransitionType::None,
            speed,
            duration_ms: extended_duration_ms.or(legacy_duration_ms),
            through_black: None,
            split_direction: None,
            advance_on_click,
            advance_after_ms,
            sound: None,
        })
    }

    fn apply_effect(&mut self, effect: ParsedTransitionEffect) {
        self.transition_type = effect.transition_type;
        self.through_black = effect.through_black;
        self.split_direction = effect.split_direction;
    }

    /// Parse a transition effect from a direct child of `<p:transition>`.
    fn parse_transition_type(
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<ParsedTransitionEffect>> {
        if is_p14_name(namespace, element.name(), b"ripple") {
            let direction = unqualified_attribute_value(element, b"dir", decoder)?
                .map(|value| RippleDirection::from_xml_value(&value))
                .transpose()?
                .unwrap_or(RippleDirection::Center);
            return Ok(Some(ParsedTransitionEffect::new(TransitionType::Ripple {
                direction,
            })));
        }

        let tag_name = element.local_name();
        if !is_presentationml_name(namespace, element.name(), tag_name.as_ref()) {
            return Ok(None);
        }

        let transition_type = match tag_name.as_ref() {
            b"cut" => Some(TransitionType::Cut),
            b"fade" => Some(TransitionType::Fade),
            b"push" => Some(TransitionType::Push {
                direction: parse_side_direction(element, decoder)?,
            }),
            b"wipe" => Some(TransitionType::Wipe {
                direction: parse_side_direction(element, decoder)?,
            }),
            b"split" => Some(TransitionType::Split {
                direction: parse_orientation(element, b"orient", decoder)?,
            }),
            b"pull" => Some(TransitionType::Uncover {
                direction: parse_eight_direction(element, decoder)?,
            }),
            b"cover" => Some(TransitionType::Cover {
                direction: parse_eight_direction(element, decoder)?,
            }),
            b"dissolve" => Some(TransitionType::Dissolve),
            b"blinds" => Some(TransitionType::Blinds {
                direction: parse_orientation(element, b"dir", decoder)?,
            }),
            b"checker" => Some(TransitionType::Checker {
                direction: parse_orientation(element, b"dir", decoder)?,
            }),
            b"randomBar" => Some(TransitionType::RandomBars {
                direction: parse_orientation(element, b"dir", decoder)?,
            }),
            b"strips" => Some(TransitionType::Strips {
                direction: parse_corner_direction(element, decoder)?,
            }),
            b"comb" => Some(TransitionType::Comb {
                direction: parse_orientation(element, b"dir", decoder)?,
            }),
            b"circle" => Some(TransitionType::Circle),
            b"diamond" => Some(TransitionType::Diamond),
            b"plus" => Some(TransitionType::Plus),
            b"wedge" => Some(TransitionType::Wedge),
            b"zoom" => Some(TransitionType::Zoom {
                direction: parse_zoom_direction(element, decoder)?,
            }),
            b"wheel" => Some(TransitionType::Wheel {
                spokes: parse_wheel_spokes(element, decoder)?,
            }),
            b"random" => Some(TransitionType::Random),
            b"newsflash" => Some(TransitionType::Newsflash),
            _ => None,
        };

        let Some(transition_type) = transition_type else {
            return Ok(None);
        };
        let through_black =
            if matches!(&transition_type, TransitionType::Cut | TransitionType::Fade) {
                parse_optional_boolean(element, b"thruBlk", decoder)?
            } else {
                None
            };
        let split_direction = if matches!(&transition_type, TransitionType::Split { .. }) {
            parse_split_direction(element, decoder)?
        } else {
            None
        };

        Ok(Some(ParsedTransitionEffect {
            transition_type,
            through_black,
            split_direction,
        }))
    }

    /// Generate XML for this transition.
    pub(crate) fn to_xml(&self) -> Result<String> {
        self.validate_effect_options()?;

        if let TransitionType::Ripple { direction } = &self.transition_type {
            return Ok(self.ripple_to_xml(*direction));
        }

        let mut xml = String::with_capacity(512);
        self.write_transition_start(&mut xml, Some("dur"));

        self.write_transition_type_xml(&mut xml)?;
        xml.push_str("</p:transition>");

        Ok(xml)
    }

    fn write_transition_start(&self, xml: &mut String, duration_attribute: Option<&str>) {
        xml.push_str(r#"<p:transition"#);
        xml.push_str(r#" spd=""#);
        xml.push_str(self.speed.to_xml_value());
        xml.push('"');

        if let (Some(duration_attribute), Some(dur)) = (duration_attribute, self.duration_ms) {
            xml.push(' ');
            xml.push_str(duration_attribute);
            xml.push_str(r#"=""#);
            xml.push_str(&dur.to_string());
            xml.push('"');
        }

        if !self.advance_on_click {
            xml.push_str(r#" advClick="0""#);
        }

        if let Some(adv) = self.advance_after_ms {
            xml.push_str(r#" advTm=""#);
            xml.push_str(&adv.to_string());
            xml.push('"');
        }

        xml.push('>');
    }

    fn ripple_to_xml(&self, direction: RippleDirection) -> String {
        let mut xml = String::with_capacity(512);
        xml.push_str(r#"<mc:AlternateContent xmlns:mc=""#);
        xml.push_str(MARKUP_COMPATIBILITY_NAMESPACE);
        xml.push_str(r#"" xmlns:p14=""#);
        xml.push_str(P14_NAMESPACE);
        xml.push_str(r#""><mc:Choice Requires="p14">"#);
        self.write_transition_start(&mut xml, Some("p14:dur"));
        xml.push_str(r#"<p14:ripple dir=""#);
        xml.push_str(direction.to_xml_value());
        xml.push_str(r#""/></p:transition></mc:Choice><mc:Fallback>"#);
        self.write_transition_start(&mut xml, None);
        xml.push_str("<p:fade/></p:transition></mc:Fallback></mc:AlternateContent>");

        xml
    }

    fn validate_effect_options(&self) -> Result<()> {
        if self.through_black.is_some()
            && !matches!(
                &self.transition_type,
                TransitionType::Cut | TransitionType::Fade
            )
        {
            return Err(OoxmlError::InvalidFormat(
                "through-black is only valid for cut and fade transitions".to_string(),
            ));
        }
        if self.split_direction.is_some()
            && !matches!(&self.transition_type, TransitionType::Split { .. })
        {
            return Err(OoxmlError::InvalidFormat(
                "split direction is only valid for split transitions".to_string(),
            ));
        }

        Ok(())
    }

    /// Write the transition type-specific XML.
    fn write_transition_type_xml(&self, xml: &mut String) -> Result<()> {
        match &self.transition_type {
            TransitionType::None => {
                // No transition element
            },
            TransitionType::Cut => {
                xml.push_str("<p:cut");
                if let Some(through_black) = self.through_black {
                    xml.push_str(" thruBlk=\"");
                    xml.push_str(boolean_to_xml(through_black));
                    xml.push('"');
                }
                xml.push_str("/>");
            },
            TransitionType::Fade => {
                xml.push_str("<p:fade");
                if let Some(through_black) = self.through_black {
                    xml.push_str(" thruBlk=\"");
                    xml.push_str(boolean_to_xml(through_black));
                    xml.push('"');
                }
                xml.push_str("/>");
            },
            TransitionType::Push { direction } => {
                xml.push_str("<p:push dir=\"");
                xml.push_str(Self::side_direction_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Wipe { direction } => {
                xml.push_str("<p:wipe dir=\"");
                xml.push_str(Self::side_direction_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Split { direction } => {
                // Split uses "orient" not "dir" per OOXML spec
                xml.push_str("<p:split orient=\"");
                xml.push_str(Self::orientation_to_xml(*direction)?);
                xml.push('"');
                if let Some(split_direction) = self.split_direction {
                    xml.push_str(" dir=\"");
                    xml.push_str(split_direction_to_xml(split_direction));
                    xml.push('"');
                }
                xml.push_str("/>");
            },
            TransitionType::Uncover { direction } => {
                xml.push_str("<p:pull dir=\"");
                xml.push_str(Self::eight_direction_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Cover { direction } => {
                xml.push_str("<p:cover dir=\"");
                xml.push_str(Self::eight_direction_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Dissolve => {
                xml.push_str("<p:dissolve/>");
            },
            TransitionType::Blinds { direction } => {
                xml.push_str("<p:blinds dir=\"");
                xml.push_str(Self::orientation_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Checker { direction } => {
                xml.push_str("<p:checker dir=\"");
                xml.push_str(Self::orientation_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::RandomBars { direction } => {
                xml.push_str("<p:randomBar dir=\"");
                xml.push_str(Self::orientation_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Strips { direction } => {
                xml.push_str("<p:strips dir=\"");
                xml.push_str(Self::corner_direction_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Comb { direction } => {
                xml.push_str("<p:comb dir=\"");
                xml.push_str(Self::orientation_to_xml(*direction)?);
                xml.push_str("\"/>");
            },
            TransitionType::Circle => {
                xml.push_str("<p:circle/>");
            },
            TransitionType::Diamond => {
                xml.push_str("<p:diamond/>");
            },
            TransitionType::Plus => {
                xml.push_str("<p:plus/>");
            },
            TransitionType::Shape { shape_type } => match shape_type {
                ShapeTransitionType::Circle => xml.push_str("<p:circle/>"),
                ShapeTransitionType::Diamond => xml.push_str("<p:diamond/>"),
                ShapeTransitionType::Plus => xml.push_str("<p:plus/>"),
            },
            TransitionType::Wedge => {
                xml.push_str("<p:wedge/>");
            },
            TransitionType::Zoom { direction } => {
                let dir_str = match direction {
                    ZoomDirection::In => "in",
                    ZoomDirection::Out => "out",
                };
                xml.push_str("<p:zoom dir=\"");
                xml.push_str(dir_str);
                xml.push_str("\"/>");
            },
            TransitionType::Random => {
                xml.push_str("<p:random/>");
            },
            TransitionType::Wheel { spokes } => {
                xml.push_str("<p:wheel spokes=\"");
                xml.push_str(&spokes.to_string());
                xml.push_str("\"/>");
            },
            TransitionType::Newsflash => {
                xml.push_str("<p:newsflash/>");
            },
            TransitionType::Ripple { .. } => {
                return Err(OoxmlError::Other(
                    "ripple transitions must be emitted through compatibility markup".to_string(),
                ));
            },
            TransitionType::Reveal { .. }
            | TransitionType::Clock { .. }
            | TransitionType::Flash => {
                return Err(OoxmlError::Other(
                    "this transition type does not have a standard PresentationML writer"
                        .to_string(),
                ));
            },
            TransitionType::Other(name) => {
                return Err(OoxmlError::Other(format!(
                    "cannot serialize unknown transition type '{name}'"
                )));
            },
        }

        Ok(())
    }

    fn side_direction_to_xml(direction: TransitionDirection) -> Result<&'static str> {
        match direction {
            TransitionDirection::Left => Ok("l"),
            TransitionDirection::Right => Ok("r"),
            TransitionDirection::Up => Ok("u"),
            TransitionDirection::Down => Ok("d"),
            _ => Err(invalid_writer_direction("side", direction)),
        }
    }

    fn orientation_to_xml(direction: TransitionDirection) -> Result<&'static str> {
        match direction {
            TransitionDirection::Horizontal => Ok("horz"),
            TransitionDirection::Vertical => Ok("vert"),
            _ => Err(invalid_writer_direction("orientation", direction)),
        }
    }

    fn corner_direction_to_xml(direction: TransitionDirection) -> Result<&'static str> {
        match direction {
            TransitionDirection::LeftUp => Ok("lu"),
            TransitionDirection::RightUp => Ok("ru"),
            TransitionDirection::LeftDown => Ok("ld"),
            TransitionDirection::RightDown => Ok("rd"),
            _ => Err(invalid_writer_direction("corner", direction)),
        }
    }

    fn eight_direction_to_xml(direction: TransitionDirection) -> Result<&'static str> {
        match direction {
            TransitionDirection::Left => Ok("l"),
            TransitionDirection::Right => Ok("r"),
            TransitionDirection::Up => Ok("u"),
            TransitionDirection::Down => Ok("d"),
            TransitionDirection::LeftUp => Ok("lu"),
            TransitionDirection::RightUp => Ok("ru"),
            TransitionDirection::LeftDown => Ok("ld"),
            TransitionDirection::RightDown => Ok("rd"),
            _ => Err(invalid_writer_direction("eight-way", direction)),
        }
    }
}

fn parse_duration_ms(value: &str) -> Option<u32> {
    value
        .strip_suffix("ms")
        .unwrap_or(value)
        .parse::<u32>()
        .ok()
}

fn parse_side_direction(element: &BytesStart<'_>, decoder: Decoder) -> Result<TransitionDirection> {
    let value =
        unqualified_attribute_value(element, b"dir", decoder)?.unwrap_or_else(|| "l".to_string());
    match value.as_str() {
        "l" => Ok(TransitionDirection::Left),
        "r" => Ok(TransitionDirection::Right),
        "u" => Ok(TransitionDirection::Up),
        "d" => Ok(TransitionDirection::Down),
        _ => Err(invalid_xml_direction("side", &value)),
    }
}

fn parse_orientation(
    element: &BytesStart<'_>,
    attribute: &[u8],
    decoder: Decoder,
) -> Result<TransitionDirection> {
    let value = unqualified_attribute_value(element, attribute, decoder)?
        .unwrap_or_else(|| "horz".to_string());
    match value.as_str() {
        "horz" => Ok(TransitionDirection::Horizontal),
        "vert" => Ok(TransitionDirection::Vertical),
        _ => Err(invalid_xml_direction("orientation", &value)),
    }
}

fn parse_corner_direction(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<TransitionDirection> {
    let value =
        unqualified_attribute_value(element, b"dir", decoder)?.unwrap_or_else(|| "lu".to_string());
    match value.as_str() {
        "lu" => Ok(TransitionDirection::LeftUp),
        "ru" => Ok(TransitionDirection::RightUp),
        "ld" => Ok(TransitionDirection::LeftDown),
        "rd" => Ok(TransitionDirection::RightDown),
        _ => Err(invalid_xml_direction("corner", &value)),
    }
}

fn parse_eight_direction(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<TransitionDirection> {
    let value =
        unqualified_attribute_value(element, b"dir", decoder)?.unwrap_or_else(|| "l".to_string());
    match value.as_str() {
        "l" => Ok(TransitionDirection::Left),
        "r" => Ok(TransitionDirection::Right),
        "u" => Ok(TransitionDirection::Up),
        "d" => Ok(TransitionDirection::Down),
        "lu" => Ok(TransitionDirection::LeftUp),
        "ru" => Ok(TransitionDirection::RightUp),
        "ld" => Ok(TransitionDirection::LeftDown),
        "rd" => Ok(TransitionDirection::RightDown),
        _ => Err(invalid_xml_direction("eight-way", &value)),
    }
}

fn parse_zoom_direction(element: &BytesStart<'_>, decoder: Decoder) -> Result<ZoomDirection> {
    let value =
        unqualified_attribute_value(element, b"dir", decoder)?.unwrap_or_else(|| "in".to_string());
    match value.as_str() {
        "in" => Ok(ZoomDirection::In),
        "out" => Ok(ZoomDirection::Out),
        _ => Err(invalid_xml_direction("zoom", &value)),
    }
}

fn parse_split_direction(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Option<SplitDirection>> {
    let Some(value) = unqualified_attribute_value(element, b"dir", decoder)? else {
        return Ok(None);
    };
    match value.as_str() {
        "in" => Ok(Some(SplitDirection::In)),
        "out" => Ok(Some(SplitDirection::Out)),
        _ => Err(invalid_xml_direction("split", &value)),
    }
}

fn parse_optional_boolean(
    element: &BytesStart<'_>,
    attribute: &[u8],
    decoder: Decoder,
) -> Result<Option<bool>> {
    let Some(value) = unqualified_attribute_value(element, attribute, decoder)? else {
        return Ok(None);
    };
    match value.as_str() {
        "1" | "true" => Ok(Some(true)),
        "0" | "false" => Ok(Some(false)),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "invalid boolean transition attribute '{}'",
            value
        ))),
    }
}

fn parse_wheel_spokes(element: &BytesStart<'_>, decoder: Decoder) -> Result<u8> {
    let Some(value) = unqualified_attribute_value(element, b"spokes", decoder)? else {
        return Ok(4);
    };
    value.parse::<u8>().map_err(|_| {
        OoxmlError::InvalidFormat(format!("invalid wheel transition spoke count '{value}'"))
    })
}

fn invalid_xml_direction(kind: &str, value: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("invalid {kind} transition direction '{value}'"))
}

fn invalid_writer_direction(kind: &str, direction: TransitionDirection) -> OoxmlError {
    OoxmlError::InvalidFormat(format!(
        "{direction:?} is not valid for a {kind} transition direction"
    ))
}

fn boolean_to_xml(value: bool) -> &'static str {
    if value { "1" } else { "0" }
}

fn split_direction_to_xml(direction: SplitDirection) -> &'static str {
    match direction {
        SplitDirection::In => "in",
        SplitDirection::Out => "out",
    }
}

fn is_p14_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == P14_NAMESPACE_BYTES
    )
}

fn is_p14_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name && is_p14_namespace(namespace)
}

#[cfg(test)]
mod tests {
    use super::*;

    const STANDARD_COVER: &[u8] =
        include_bytes!("../../../../test-data/ooxml/pptx/transitions/standard_cover.xml");
    const STANDARD_EFFECT_OPTIONS: &[u8] =
        include_bytes!("../../../../test-data/ooxml/pptx/transitions/standard_effect_options.xml");

    #[test]
    fn test_transition_speed() {
        assert_eq!(TransitionSpeed::Fast.duration_ms(), 500);
        assert_eq!(TransitionSpeed::Medium.duration_ms(), 1000);
        assert_eq!(TransitionSpeed::Slow.duration_ms(), 1500);
    }

    #[test]
    fn test_transition_builder() {
        let trans = SlideTransition::new(TransitionType::Fade)
            .with_speed(TransitionSpeed::Fast)
            .with_advance_after_ms(3000);

        assert_eq!(trans.transition_type, TransitionType::Fade);
        assert_eq!(trans.speed, TransitionSpeed::Fast);
        assert_eq!(trans.advance_after_ms, Some(3000));
    }

    #[test]
    fn test_transition_xml_generation() {
        let trans = SlideTransition::new(TransitionType::Fade).with_speed(TransitionSpeed::Fast);

        let xml = trans.to_xml().unwrap();
        assert!(xml.contains("spd=\"fast\""));
        assert!(xml.contains("<p:fade"));
    }

    #[test]
    fn ripple_transition_writes_a_compatibility_choice_and_round_trips() {
        let transition = SlideTransition::new(TransitionType::Ripple {
            direction: RippleDirection::LeftDown,
        })
        .with_speed(TransitionSpeed::Slow)
        .with_duration_ms(1500)
        .with_advance_on_click(false)
        .with_advance_after_ms(4250);

        let xml = transition.to_xml().unwrap();
        assert!(xml.contains(r#"<mc:AlternateContent"#));
        assert!(xml.contains(r#"<mc:Choice Requires="p14">"#));
        assert!(xml.contains(r#"p14:dur="1500""#));
        assert!(xml.contains(r#"<p14:ripple dir="ld"/>"#));
        assert!(xml.contains(
            r#"<mc:Fallback><p:transition spd="slow" advClick="0" advTm="4250"><p:fade/>"#
        ));
        assert_eq!(parse_serialized_transition(&xml), transition);
    }

    #[test]
    fn parses_local_standard_cover_fixture() {
        let transition = SlideTransition::from_xml(STANDARD_COVER).unwrap().unwrap();

        assert_eq!(transition.speed, TransitionSpeed::Fast);
        assert_eq!(transition.advance_on_click, false);
        assert_eq!(transition.advance_after_ms, Some(750));
        assert_eq!(
            transition.transition_type,
            TransitionType::Cover {
                direction: TransitionDirection::RightDown,
            }
        );
    }

    #[test]
    fn parses_local_standard_effect_options_fixture() {
        let transition = SlideTransition::from_xml(STANDARD_EFFECT_OPTIONS)
            .unwrap()
            .unwrap();

        assert_eq!(transition.transition_type, TransitionType::Fade);
        assert_eq!(transition.through_black, Some(true));
        assert_eq!(transition.split_direction, None);
    }

    #[test]
    fn parses_standard_transition_effects_and_directions() {
        assert_eq!(
            parse_effect(r#"<p:push dir="d"/>"#).transition_type,
            TransitionType::Push {
                direction: TransitionDirection::Down,
            }
        );
        let split = parse_effect(r#"<p:split orient="vert" dir="in"/>"#);
        assert_eq!(
            split.transition_type,
            TransitionType::Split {
                direction: TransitionDirection::Vertical,
            }
        );
        assert_eq!(split.split_direction, Some(SplitDirection::In));
        assert_eq!(
            parse_effect(r#"<p:pull dir="lu"/>"#).transition_type,
            TransitionType::Uncover {
                direction: TransitionDirection::LeftUp,
            }
        );
        assert_eq!(
            parse_effect(r#"<p:blinds dir="vert"/>"#).transition_type,
            TransitionType::Blinds {
                direction: TransitionDirection::Vertical,
            }
        );
        assert_eq!(
            parse_effect(r#"<p:randomBar dir="vert"/>"#).transition_type,
            TransitionType::RandomBars {
                direction: TransitionDirection::Vertical,
            }
        );
        assert_eq!(
            parse_effect(r#"<p:strips dir="ld"/>"#).transition_type,
            TransitionType::Strips {
                direction: TransitionDirection::LeftDown,
            }
        );
        assert_eq!(
            parse_effect(r#"<p:comb dir="vert"/>"#).transition_type,
            TransitionType::Comb {
                direction: TransitionDirection::Vertical,
            }
        );
        assert_eq!(
            parse_effect(r#"<p:wheel spokes="6"/>"#).transition_type,
            TransitionType::Wheel { spokes: 6 }
        );
        assert_eq!(
            parse_effect(r#"<p:zoom dir="out"/>"#).transition_type,
            TransitionType::Zoom {
                direction: ZoomDirection::Out,
            }
        );
        assert_eq!(
            parse_effect(r#"<p:newsflash/>"#).transition_type,
            TransitionType::Newsflash
        );
    }

    #[test]
    fn writes_standard_transition_effects_without_fade_fallbacks() {
        let cases = [
            (
                TransitionType::Push {
                    direction: TransitionDirection::Down,
                },
                r#"<p:push dir="d"/>"#,
            ),
            (
                TransitionType::Split {
                    direction: TransitionDirection::Vertical,
                },
                r#"<p:split orient="vert"/>"#,
            ),
            (
                TransitionType::Uncover {
                    direction: TransitionDirection::LeftUp,
                },
                r#"<p:pull dir="lu"/>"#,
            ),
            (
                TransitionType::Cover {
                    direction: TransitionDirection::RightDown,
                },
                r#"<p:cover dir="rd"/>"#,
            ),
            (
                TransitionType::Blinds {
                    direction: TransitionDirection::Vertical,
                },
                r#"<p:blinds dir="vert"/>"#,
            ),
            (
                TransitionType::RandomBars {
                    direction: TransitionDirection::Vertical,
                },
                r#"<p:randomBar dir="vert"/>"#,
            ),
            (
                TransitionType::Strips {
                    direction: TransitionDirection::LeftDown,
                },
                r#"<p:strips dir="ld"/>"#,
            ),
            (
                TransitionType::Comb {
                    direction: TransitionDirection::Vertical,
                },
                r#"<p:comb dir="vert"/>"#,
            ),
            (
                TransitionType::Wheel { spokes: 6 },
                r#"<p:wheel spokes="6"/>"#,
            ),
            (TransitionType::Newsflash, "<p:newsflash/>"),
            (
                TransitionType::Shape {
                    shape_type: ShapeTransitionType::Plus,
                },
                "<p:plus/>",
            ),
        ];

        for (transition_type, expected_effect) in cases {
            let xml = SlideTransition::new(transition_type).to_xml().unwrap();
            assert!(
                xml.contains(expected_effect),
                "expected {expected_effect:?} in {xml:?}"
            );
            assert!(!xml.contains("<p:fade"), "unexpected fallback in {xml:?}");
        }
    }

    #[test]
    fn preserves_through_black_and_split_direction_options() {
        let fade = SlideTransition::new(TransitionType::Fade).with_through_black(true);
        let fade_xml = fade.to_xml().unwrap();
        assert!(fade_xml.contains(r#"<p:fade thruBlk="1"/>"#));
        assert_eq!(
            parse_serialized_transition(&fade_xml).through_black,
            Some(true)
        );

        let cut = SlideTransition::new(TransitionType::Cut).with_through_black(false);
        assert!(cut.to_xml().unwrap().contains(r#"<p:cut thruBlk="0"/>"#));

        let split = SlideTransition::new(TransitionType::Split {
            direction: TransitionDirection::Vertical,
        })
        .with_split_direction(SplitDirection::In);
        let split_xml = split.to_xml().unwrap();
        assert!(split_xml.contains(r#"<p:split orient="vert" dir="in"/>"#));
        let parsed = parse_serialized_transition(&split_xml);
        assert_eq!(parsed.split_direction, Some(SplitDirection::In));
    }

    #[test]
    fn rejects_invalid_standard_transition_directions() {
        let xml = transition_xml(r#"<p:push dir="horz"/>"#);
        assert!(matches!(
            SlideTransition::from_xml(xml.as_bytes()),
            Err(OoxmlError::InvalidFormat(message)) if message.contains("side transition direction")
        ));

        let transition = SlideTransition::new(TransitionType::Wipe {
            direction: TransitionDirection::Horizontal,
        });
        assert!(matches!(
            transition.to_xml(),
            Err(OoxmlError::InvalidFormat(message)) if message.contains("not valid")
        ));
    }

    #[test]
    fn rejects_effect_options_on_incompatible_transition_types() {
        let through_black = SlideTransition::new(TransitionType::Push {
            direction: TransitionDirection::Left,
        })
        .with_through_black(true);
        assert!(matches!(
            through_black.to_xml(),
            Err(OoxmlError::InvalidFormat(message)) if message.contains("through-black")
        ));

        let split_direction =
            SlideTransition::new(TransitionType::Fade).with_split_direction(SplitDirection::Out);
        assert!(matches!(
            split_direction.to_xml(),
            Err(OoxmlError::InvalidFormat(message)) if message.contains("split direction")
        ));
    }

    fn parse_effect(effect: &str) -> SlideTransition {
        let xml = transition_xml(effect);
        SlideTransition::from_xml(xml.as_bytes()).unwrap().unwrap()
    }

    fn transition_xml(effect: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:transition>{effect}</p:transition></p:sld>"#
        )
    }

    fn parse_serialized_transition(xml: &str) -> SlideTransition {
        let xml = format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{xml}</p:sld>"#
        );
        SlideTransition::from_xml(xml.as_bytes()).unwrap().unwrap()
    }
}
