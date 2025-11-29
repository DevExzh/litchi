//! Slide transition effects for PowerPoint presentations.
//!
//! This module provides types and functionality for working with slide transitions,
//! including transition types, speeds, and directions.

use crate::ooxml::error::{OoxmlError, Result};

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

impl Default for SlideTransition {
    fn default() -> Self {
        Self {
            transition_type: TransitionType::None,
            speed: TransitionSpeed::Medium,
            duration_ms: None,
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
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut transition: Option<SlideTransition> = None;
        let mut advance_on_click = true;
        let mut advance_after_ms: Option<u32> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    if tag_name.as_ref() == b"transition" {
                        let mut speed = TransitionSpeed::Medium;
                        let mut duration_ms = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"spd" => {
                                    let val = std::str::from_utf8(&attr.value).unwrap_or("med");
                                    speed = TransitionSpeed::from_xml_value(val);
                                },
                                b"dur" => {
                                    if let Ok(val_str) = std::str::from_utf8(&attr.value)
                                        && let Ok(val) = val_str.parse::<u32>()
                                    {
                                        duration_ms = Some(val);
                                    }
                                },
                                b"advClick" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        advance_on_click = val == "1" || val == "true";
                                    }
                                },
                                b"advTm" => {
                                    if let Ok(val_str) = std::str::from_utf8(&attr.value)
                                        && let Ok(val) = val_str.parse::<u32>()
                                    {
                                        advance_after_ms = Some(val);
                                    }
                                },
                                _ => {},
                            }
                        }

                        transition = Some(SlideTransition {
                            transition_type: TransitionType::None,
                            speed,
                            duration_ms,
                            advance_on_click,
                            advance_after_ms,
                            sound: None,
                        });
                    }

                    // Parse specific transition types
                    if transition.is_some() {
                        let t_type = Self::parse_transition_type(tag_name.as_ref(), e)?;
                        if let Some(t) = t_type
                            && let Some(ref mut trans) = transition
                        {
                            trans.transition_type = t;
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(transition)
    }

    /// Parse transition type from XML element.
    fn parse_transition_type(
        tag_name: &[u8],
        _element: &quick_xml::events::BytesStart<'_>,
    ) -> Result<Option<TransitionType>> {
        let t = match tag_name {
            b"cut" => Some(TransitionType::Cut),
            b"fade" => Some(TransitionType::Fade),
            b"push" => Some(TransitionType::Push {
                direction: TransitionDirection::Left,
            }),
            b"wipe" => Some(TransitionType::Wipe {
                direction: TransitionDirection::Left,
            }),
            b"split" => Some(TransitionType::Split {
                direction: TransitionDirection::Horizontal,
            }),
            b"dissolve" => Some(TransitionType::Dissolve),
            b"blinds" => Some(TransitionType::Blinds {
                direction: TransitionDirection::Vertical,
            }),
            b"checker" => Some(TransitionType::Checker {
                direction: TransitionDirection::Horizontal,
            }),
            b"circle" => Some(TransitionType::Circle),
            b"diamond" => Some(TransitionType::Diamond),
            b"plus" => Some(TransitionType::Plus),
            b"wedge" => Some(TransitionType::Wedge),
            b"zoom" => Some(TransitionType::Zoom {
                direction: ZoomDirection::In,
            }),
            b"random" => Some(TransitionType::Random),
            _ => None,
        };

        Ok(t)
    }

    /// Generate XML for this transition.
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(512);

        xml.push_str(r#"<p:transition"#);
        xml.push_str(r#" spd=""#);
        xml.push_str(self.speed.to_xml_value());
        xml.push('"');

        if let Some(dur) = self.duration_ms {
            xml.push_str(r#" dur=""#);
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

        // Add transition type XML
        self.write_transition_type_xml(&mut xml)?;

        xml.push_str("</p:transition>");

        Ok(xml)
    }

    /// Write the transition type-specific XML.
    fn write_transition_type_xml(&self, xml: &mut String) -> Result<()> {
        match &self.transition_type {
            TransitionType::None => {
                // No transition element
            },
            TransitionType::Cut => {
                xml.push_str("<p:cut/>");
            },
            TransitionType::Fade => {
                xml.push_str("<p:fade thruBlk=\"false\"/>");
            },
            TransitionType::Push { direction } => {
                xml.push_str("<p:push dir=\"");
                xml.push_str(Self::direction_to_xml(*direction));
                xml.push_str("\"/>");
            },
            TransitionType::Wipe { direction } => {
                xml.push_str("<p:wipe dir=\"");
                xml.push_str(Self::direction_to_xml(*direction));
                xml.push_str("\"/>");
            },
            TransitionType::Split { direction } => {
                // Split uses "orient" not "dir" per OOXML spec
                xml.push_str("<p:split orient=\"");
                xml.push_str(Self::direction_to_xml(*direction));
                xml.push_str("\"/>");
            },
            TransitionType::Dissolve => {
                xml.push_str("<p:dissolve/>");
            },
            TransitionType::Blinds { direction } => {
                // Blinds uses "orient" not "dir" per OOXML spec
                xml.push_str("<p:blinds orient=\"");
                xml.push_str(Self::direction_to_xml(*direction));
                xml.push_str("\"/>");
            },
            TransitionType::Checker { direction } => {
                xml.push_str("<p:checker dir=\"");
                xml.push_str(Self::direction_to_xml(*direction));
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
            _ => {
                // For other types, use fade as fallback
                xml.push_str("<p:fade thruBlk=\"false\"/>");
            },
        }

        Ok(())
    }

    /// Convert direction to XML attribute value.
    fn direction_to_xml(direction: TransitionDirection) -> &'static str {
        match direction {
            TransitionDirection::Left => "l",
            TransitionDirection::Right => "r",
            TransitionDirection::Up => "u",
            TransitionDirection::Down => "d",
            TransitionDirection::Horizontal => "horz",
            TransitionDirection::Vertical => "vert",
            TransitionDirection::In => "in",
            TransitionDirection::Out => "out",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
