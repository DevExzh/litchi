//! Animation support for PowerPoint presentations.
//!
//! This module provides read/write support for slide animations and timing.

use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

/// Animation effect type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationEffect {
    /// Appear effect
    Appear,
    /// Fade effect
    Fade,
    /// Fly in effect
    FlyIn,
    /// Float in effect
    FloatIn,
    /// Split effect
    Split,
    /// Wipe effect
    Wipe,
    /// Zoom effect
    Zoom,
    /// Bounce effect
    Bounce,
    /// Spin effect
    Spin,
    /// Grow/Shrink effect
    GrowShrink,
    /// Custom/Unknown effect
    Custom(String),
}

impl AnimationEffect {
    /// Parse from preset ID.
    pub fn from_preset_id(id: u32) -> Self {
        match id {
            1 => AnimationEffect::Appear,
            2 => AnimationEffect::FlyIn,
            10 => AnimationEffect::Fade,
            16 => AnimationEffect::Split,
            22 => AnimationEffect::Wipe,
            23 => AnimationEffect::Zoom,
            24 => AnimationEffect::Bounce,
            _ => AnimationEffect::Custom(format!("preset_{}", id)),
        }
    }

    /// Parse from preset class string (for backwards compatibility).
    pub fn from_preset(preset: &str) -> Self {
        match preset.to_lowercase().as_str() {
            "entr" | "appear" => AnimationEffect::Appear,
            "fade" => AnimationEffect::Fade,
            "fly" | "flyin" => AnimationEffect::FlyIn,
            "float" | "floatin" => AnimationEffect::FloatIn,
            "split" => AnimationEffect::Split,
            "wipe" => AnimationEffect::Wipe,
            "zoom" => AnimationEffect::Zoom,
            "bounce" => AnimationEffect::Bounce,
            "spin" => AnimationEffect::Spin,
            "grow" | "growshrink" => AnimationEffect::GrowShrink,
            other => AnimationEffect::Custom(other.to_string()),
        }
    }

    /// Get the preset ID for this effect.
    /// These are defined in ECMA-376 Part 1.
    pub fn preset_id(&self) -> u32 {
        match self {
            AnimationEffect::Appear => 1,
            AnimationEffect::FlyIn => 2,
            AnimationEffect::FloatIn => 42,
            AnimationEffect::Split => 16,
            AnimationEffect::Fade => 10,
            AnimationEffect::Wipe => 22,
            AnimationEffect::Zoom => 23,
            AnimationEffect::Bounce => 24,
            AnimationEffect::Spin => 8, // Spin is emphasis, but using ID 8
            AnimationEffect::GrowShrink => 6, // GrowShrink is emphasis
            AnimationEffect::Custom(_) => 1, // Default to Appear
        }
    }

    /// Get the preset class for this effect.
    /// Valid values: "entr" (entrance), "exit", "emph" (emphasis), "path", "verb", "mediacall"
    pub fn preset_class(&self) -> &str {
        match self {
            // Entrance effects
            AnimationEffect::Appear => "entr",
            AnimationEffect::FlyIn => "entr",
            AnimationEffect::FloatIn => "entr",
            AnimationEffect::Split => "entr",
            AnimationEffect::Fade => "entr",
            AnimationEffect::Wipe => "entr",
            AnimationEffect::Zoom => "entr",
            AnimationEffect::Bounce => "entr",
            // Emphasis effects
            AnimationEffect::Spin => "emph",
            AnimationEffect::GrowShrink => "emph",
            // Default to entrance
            AnimationEffect::Custom(_) => "entr",
        }
    }

    /// Get the preset class string (deprecated, use preset_class instead).
    #[deprecated(note = "Use preset_class() and preset_id() instead")]
    pub fn to_preset(&self) -> &str {
        self.preset_class()
    }
}

/// Animation trigger type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnimationTrigger {
    /// Start on click
    #[default]
    OnClick,
    /// Start with previous animation
    WithPrevious,
    /// Start after previous animation
    AfterPrevious,
}

/// Animation direction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationDirection {
    Up,
    Down,
    Left,
    Right,
    UpLeft,
    UpRight,
    DownLeft,
    DownRight,
}

/// Duration in milliseconds.
pub type Duration = u32;

/// An animation applied to a shape.
#[derive(Debug, Clone)]
pub struct Animation {
    /// Target shape ID
    pub shape_id: u32,
    /// Animation effect
    pub effect: AnimationEffect,
    /// Trigger type
    pub trigger: AnimationTrigger,
    /// Duration in milliseconds
    pub duration: Duration,
    /// Delay before starting (ms)
    pub delay: Duration,
    /// Direction (for directional effects)
    pub direction: Option<AnimationDirection>,
    /// Sequence order (1-based)
    pub order: u32,
}

impl Animation {
    /// Create a new animation.
    pub fn new(shape_id: u32, effect: AnimationEffect) -> Self {
        Self {
            shape_id,
            effect,
            trigger: AnimationTrigger::OnClick,
            duration: 500,
            delay: 0,
            direction: None,
            order: 1,
        }
    }

    /// Set the trigger type.
    pub fn with_trigger(mut self, trigger: AnimationTrigger) -> Self {
        self.trigger = trigger;
        self
    }

    /// Set the duration.
    pub fn with_duration(mut self, duration: Duration) -> Self {
        self.duration = duration;
        self
    }

    /// Set the delay.
    pub fn with_delay(mut self, delay: Duration) -> Self {
        self.delay = delay;
        self
    }

    /// Set the direction.
    pub fn with_direction(mut self, direction: AnimationDirection) -> Self {
        self.direction = Some(direction);
        self
    }
}

/// Animation sequence for a slide.
#[derive(Debug, Clone, Default)]
pub struct AnimationSequence {
    /// List of animations in order
    pub animations: Vec<Animation>,
}

impl AnimationSequence {
    /// Create a new empty animation sequence.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add an animation to the sequence.
    pub fn add(&mut self, animation: Animation) {
        self.animations.push(animation);
    }

    /// Get the number of animations.
    pub fn len(&self) -> usize {
        self.animations.len()
    }

    /// Check if the sequence is empty.
    pub fn is_empty(&self) -> bool {
        self.animations.is_empty()
    }

    /// Parse timing XML from a slide.
    pub fn parse_timing_xml(xml: &str) -> Result<Self> {
        let mut sequence = Self::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut current_shape_id: Option<u32> = None;
        let mut order = 1u32;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                    b"spTgt" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"spid"
                                && let Ok(id_str) = std::str::from_utf8(&attr.value)
                            {
                                current_shape_id = id_str.parse().ok();
                            }
                        }
                    },
                    b"anim" | b"set" | b"animEffect" => {
                        if let Some(shape_id) = current_shape_id {
                            let mut anim = Animation::new(shape_id, AnimationEffect::Appear);
                            anim.order = order;
                            order += 1;
                            sequence.add(anim);
                        }
                    },
                    _ => {},
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(sequence)
    }

    /// Generate timing XML for a slide.
    pub fn to_xml(&self) -> String {
        if self.is_empty() {
            return String::new();
        }

        let mut xml = String::with_capacity(2048);
        xml.push_str("<p:timing>");
        xml.push_str("<p:tnLst>");
        xml.push_str(r#"<p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">"#);
        xml.push_str(r#"<p:childTnLst><p:seq concurrent="1" nextAc="seek">"#);
        xml.push_str(r#"<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>"#);

        let mut tn_id = 3u32;
        for anim in &self.animations {
            xml.push_str(&format!(
                r#"<p:par><p:cTn id="{}" fill="hold"><p:stCondLst><p:cond delay="{}"/></p:stCondLst>"#,
                tn_id,
                if anim.trigger == AnimationTrigger::OnClick { "indefinite" } else { "0" }
            ));
            tn_id += 1;

            xml.push_str("<p:childTnLst><p:par>");
            xml.push_str(&format!(
                r#"<p:cTn id="{}" fill="hold"><p:stCondLst><p:cond delay="{}"/></p:stCondLst>"#,
                tn_id, anim.delay
            ));
            tn_id += 1;

            xml.push_str("<p:childTnLst><p:par>");
            xml.push_str(&format!(r#"<p:cTn id="{}" presetID="{}" presetClass="{}" presetSubtype="0" fill="hold" nodeType="clickEffect">"#, 
                tn_id, anim.effect.preset_id(), anim.effect.preset_class()));
            tn_id += 1;

            xml.push_str("<p:childTnLst>");
            xml.push_str(&format!(r#"<p:set><p:cBhvr><p:cTn id="{}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>"#, tn_id));
            tn_id += 1;
            xml.push_str(&format!(
                r#"<p:tgtEl><p:spTgt spid="{}"/></p:tgtEl>"#,
                anim.shape_id
            ));
            xml.push_str(r#"<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>"#);
            xml.push_str("</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>");
        }

        xml.push_str("</p:childTnLst></p:cTn>");
        xml.push_str(r#"<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>"#);
        xml.push_str(r#"<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>"#);
        xml.push_str("</p:seq></p:childTnLst></p:cTn></p:par>");
        xml.push_str("</p:tnLst>");
        xml.push_str("</p:timing>");

        xml
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_animation_effect_preset() {
        assert_eq!(AnimationEffect::Fade.preset_class(), "entr");
        assert_eq!(AnimationEffect::Fade.preset_id(), 10);
        assert_eq!(AnimationEffect::from_preset("fade"), AnimationEffect::Fade);
    }

    #[test]
    fn test_animation_sequence() {
        let mut seq = AnimationSequence::new();
        seq.add(Animation::new(1, AnimationEffect::Fade).with_duration(1000));
        seq.add(
            Animation::new(2, AnimationEffect::FlyIn).with_trigger(AnimationTrigger::AfterPrevious),
        );

        assert_eq!(seq.len(), 2);
        assert!(!seq.to_xml().is_empty());
    }
}
