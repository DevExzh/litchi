//! Bounded, inert PowerPoint click and hover action-setting discovery.
//!
//! Action settings are returned strictly as stored metadata. This module never
//! follows hyperlinks, opens files or presentations, runs macros or programs,
//! plays media, or ends, navigates, or otherwise controls a slide show.

use crate::common::xml::{is_drawingml_name, unqualified_attribute_value};
use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{is_presentationml_name, relationship_attribute_value};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_ACTION_SETTINGS: usize = 4_096;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;
const MAX_ATTRIBUTE_BYTES: usize = 4_096;

/// The user interaction that activates an action setting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptxActionTrigger {
    /// The action is configured for a click.
    Click,
    /// The action is configured for a pointer hover.
    Hover,
}

/// A reserved PowerPoint slide-show jump target.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptxSlideShowJump {
    /// End the slide show.
    EndShow,
    /// Jump to the first slide.
    FirstSlide,
    /// Jump to the last slide.
    LastSlide,
    /// Jump to the most recently viewed slide.
    LastSlideViewed,
    /// Jump to the next slide.
    NextSlide,
    /// Jump to the previous slide.
    PreviousSlide,
}

/// The recognized meaning of a stored action string.
///
/// The original string remains available from PptxActionSetting::action.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptxActionKind {
    /// A regular hyperlink relationship without a PowerPoint action string.
    Hyperlink,
    /// A relationship-targeted jump to a slide in this presentation.
    SlideJump,
    /// A reserved presentation-relative slide-show jump.
    SlideShowJump(PptxSlideShowJump),
    /// A named custom show identified by its stored numeric ID.
    CustomShow { id: u32 },
    /// An external file action.
    File,
    /// An external presentation action with its stored starting slide index.
    Presentation { start_slide_index: u32 },
    /// A macro action. The stored macro name remains inert in action().
    Macro,
    /// An external-program action.
    Program,
    /// A media-playback action.
    Media,
    /// A setting without an action string or relationship reference.
    None,
    /// An action string outside the bounded recognized PowerPoint vocabulary.
    Unknown,
}

/// A declared relationship target attached to an action setting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PptxActionTarget {
    /// An internal OPC target part.
    Internal {
        /// Absolute package part name.
        part_name: PackURI,
        /// Declared relationship type URI.
        relationship_type: String,
    },
    /// An external target retained as an inert string.
    External {
        /// Declared target URI or path.
        target: String,
        /// Declared relationship type URI.
        relationship_type: String,
    },
}

impl PptxActionTarget {
    /// Return the declared relationship type URI.
    #[inline]
    pub fn relationship_type(&self) -> &str {
        match self {
            Self::Internal {
                relationship_type, ..
            }
            | Self::External {
                relationship_type, ..
            } => relationship_type,
        }
    }

    /// Return the target part name for an internal relationship.
    #[inline]
    pub fn part_name(&self) -> Option<&PackURI> {
        match self {
            Self::Internal { part_name, .. } => Some(part_name),
            Self::External { .. } => None,
        }
    }

    /// Return the stored target string for an external relationship.
    #[inline]
    pub fn external_target(&self) -> Option<&str> {
        match self {
            Self::Internal { .. } => None,
            Self::External { target, .. } => Some(target),
        }
    }
}

/// An inert PowerPoint click or hover action setting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxActionSetting {
    slide_index: usize,
    action_index: usize,
    trigger: PptxActionTrigger,
    kind: PptxActionKind,
    action: Option<String>,
    relationship_id: Option<String>,
    target: Option<PptxActionTarget>,
    tooltip: Option<String>,
    target_frame: Option<String>,
}

impl PptxActionSetting {
    /// Return the zero-based index of the slide that owns this setting.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this setting on the slide.
    #[inline]
    pub fn action_index(&self) -> usize {
        self.action_index
    }

    /// Return whether this setting is activated by click or hover.
    #[inline]
    pub fn trigger(&self) -> PptxActionTrigger {
        self.trigger
    }

    /// Return the recognized action kind.
    #[inline]
    pub fn kind(&self) -> PptxActionKind {
        self.kind
    }

    /// Return the original stored action string, when present.
    #[inline]
    pub fn action(&self) -> Option<&str> {
        self.action.as_deref()
    }

    /// Return the optional relationship ID from the owning slide.
    #[inline]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// Return the declared relationship target, when present.
    #[inline]
    pub fn target(&self) -> Option<&PptxActionTarget> {
        self.target.as_ref()
    }

    /// Return the optional stored tooltip.
    #[inline]
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }

    /// Return the optional stored target frame.
    #[inline]
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }
}

#[derive(Default)]
pub(crate) struct ActionLoadLimits {
    total_slide_xml_bytes: usize,
    action_count: usize,
}

struct ParsedAction {
    trigger: PptxActionTrigger,
    action: Option<String>,
    relationship_id: Option<String>,
    tooltip: Option<String>,
    target_frame: Option<String>,
}

/// Load bounded, inert action settings from one PresentationML slide.
pub(crate) fn load_slide_action_settings(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut ActionLoadLimits,
) -> Result<Vec<PptxActionSetting>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "action-setting discovery requires a PresentationML slide part",
        ));
    }
    limits.add_slide_xml(slide.blob().len())?;

    scan_action_settings(slide.blob(), limits)?
        .into_iter()
        .enumerate()
        .map(|(action_index, parsed)| {
            let target = parsed
                .relationship_id
                .as_deref()
                .map(|relationship_id| resolve_target(package, slide_index, slide, relationship_id))
                .transpose()?;
            let kind = classify_action(parsed.action.as_deref(), parsed.relationship_id.is_some());
            Ok(PptxActionSetting {
                slide_index,
                action_index,
                trigger: parsed.trigger,
                kind,
                action: parsed.action,
                relationship_id: parsed.relationship_id,
                target,
                tooltip: parsed.tooltip,
                target_frame: parsed.target_frame,
            })
        })
        .collect()
}

impl ActionLoadLimits {
    fn add_slide_xml(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_SLIDE_XML_BYTES {
            return Err(limit("slide XML bytes"));
        }
        self.total_slide_xml_bytes = self
            .total_slide_xml_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total slide XML bytes"))?;
        if self.total_slide_xml_bytes > MAX_TOTAL_SLIDE_XML_BYTES {
            return Err(limit("total slide XML bytes"));
        }
        Ok(())
    }

    fn add_action(&mut self) -> Result<()> {
        self.action_count = self
            .action_count
            .checked_add(1)
            .ok_or_else(|| limit("slide action-setting count"))?;
        if self.action_count > MAX_ACTION_SETTINGS {
            return Err(limit("slide action-setting count"));
        }
        Ok(())
    }
}

fn scan_action_settings(
    xml_bytes: &[u8],
    limits: &mut ActionLoadLimits,
) -> Result<Vec<ParsedAction>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }

    let capabilities = MceCapabilities::ooxml_baseline();
    let mce_limits = MceLimits {
        max_input_bytes: MAX_SLIDE_XML_BYTES,
        max_output_bytes: MAX_SLIDE_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4_096,
        max_directive_tokens: 4_096,
        max_choices_per_alternate: 1_024,
    };
    let xml = process_markup_compatibility(xml_bytes, &capabilities, &mce_limits)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut actions = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;

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
                increment_nodes(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth"));
                }
                if depth == 1 {
                    validate_slide_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                }
                maybe_push_action(
                    &mut actions,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    limits,
                )?;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth"))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth"));
                }
                if child_depth == 1 {
                    validate_slide_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                    closed_root = true;
                }
                maybe_push_action(
                    &mut actions,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    limits,
                )?;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid slide XML nesting"));
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must close with a PresentationML sld element",
                        ));
                    }
                    closed_root = true;
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            Event::Eof => {
                if !saw_root || !closed_root || depth != 0 {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(actions)
}

fn maybe_push_action(
    actions: &mut Vec<ParsedAction>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    limits: &mut ActionLoadLimits,
) -> Result<()> {
    let trigger = if is_drawingml_name(namespace, element.name(), b"hlinkClick") {
        PptxActionTrigger::Click
    } else if is_drawingml_name(namespace, element.name(), b"hlinkHover") {
        PptxActionTrigger::Hover
    } else {
        return Ok(());
    };

    limits.add_action()?;
    let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?;
    let relationship_id =
        bounded_optional(relationship_id, "relationship ID")?.filter(|value| !value.is_empty());
    let action = bounded_optional(
        unqualified_attribute_value(element, b"action", decoder)?,
        "action string",
    )?;
    let tooltip = bounded_optional(
        unqualified_attribute_value(element, b"tooltip", decoder)?,
        "tooltip",
    )?;
    let target_frame = bounded_optional(
        unqualified_attribute_value(element, b"tgtFrame", decoder)?,
        "target frame",
    )?;
    actions.push(ParsedAction {
        trigger,
        action,
        relationship_id,
        tooltip,
        target_frame,
    });
    Ok(())
}

fn resolve_target(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    relationship_id: &str,
) -> Result<PptxActionTarget> {
    let relationship = slide.rels().get(relationship_id).ok_or_else(|| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} action setting references missing relationship '{relationship_id}'"
        ))
    })?;
    let relationship_type = relationship.reltype().to_owned();
    if relationship.is_external() {
        let target = relationship.target_ref().to_owned();
        if target.is_empty() {
            return Err(OoxmlError::InvalidRelationship(format!(
                "slide {slide_index} action relationship '{relationship_id}' has an empty external target"
            )));
        }
        bounded(&target, "external action target")?;
        return Ok(PptxActionTarget::External {
            target,
            relationship_type,
        });
    }

    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} action relationship '{relationship_id}' has an invalid target: {error}"
        ))
    })?;
    package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "slide {slide_index} action relationship '{relationship_id}' targets missing part '{}': {error}",
            part_name.as_str()
        ))
    })?;
    Ok(PptxActionTarget::Internal {
        part_name,
        relationship_type,
    })
}

fn classify_action(action: Option<&str>, has_relationship: bool) -> PptxActionKind {
    let Some(action) = action else {
        return if has_relationship { PptxActionKind::Hyperlink } else { PptxActionKind::None };
    };
    match action {
        "ppaction://hlinkfile" => PptxActionKind::File,
        "ppaction://hlinksldjump" => PptxActionKind::SlideJump,
        "ppaction://hlinkshowjump?jump=endshow" => {
            PptxActionKind::SlideShowJump(PptxSlideShowJump::EndShow)
        },
        "ppaction://hlinkshowjump?jump=firstslide" => {
            PptxActionKind::SlideShowJump(PptxSlideShowJump::FirstSlide)
        },
        "ppaction://hlinkshowjump?jump=lastslide" => {
            PptxActionKind::SlideShowJump(PptxSlideShowJump::LastSlide)
        },
        "ppaction://hlinkshowjump?jump=lastslideviewed" => {
            PptxActionKind::SlideShowJump(PptxSlideShowJump::LastSlideViewed)
        },
        "ppaction://hlinkshowjump?jump=nextslide" => {
            PptxActionKind::SlideShowJump(PptxSlideShowJump::NextSlide)
        },
        "ppaction://hlinkshowjump?jump=previousslide" => {
            PptxActionKind::SlideShowJump(PptxSlideShowJump::PreviousSlide)
        },
        "ppaction://program" => PptxActionKind::Program,
        "ppaction://media" => PptxActionKind::Media,
        action if action.starts_with("ppaction://customshow?id=") => action
            .strip_prefix("ppaction://customshow?id=")
            .and_then(|value| value.parse().ok())
            .map_or(PptxActionKind::Unknown, |id| PptxActionKind::CustomShow {
                id,
            }),
        action if action.starts_with("ppaction://hlinkpres?slideindex=") => action
            .strip_prefix("ppaction://hlinkpres?slideindex=")
            .and_then(|value| value.parse().ok())
            .map_or(PptxActionKind::Unknown, |start_slide_index| {
                PptxActionKind::Presentation { start_slide_index }
            }),
        action
            if action.starts_with("ppaction://macro?name=")
                && action
                    .strip_prefix("ppaction://macro?name=")
                    .is_some_and(|name| !name.is_empty()) =>
        {
            PptxActionKind::Macro
        },
        _ => PptxActionKind::Unknown,
    }
}

fn validate_slide_root(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    root_seen: bool,
) -> Result<()> {
    if root_seen || !is_presentationml_name(namespace, name, b"sld") {
        return Err(invalid(
            "slide XML must have one PresentationML sld root element",
        ));
    }
    Ok(())
}

fn bounded_optional(value: Option<String>, what: &str) -> Result<Option<String>> {
    if let Some(value) = &value {
        bounded(value, what)?;
    }
    Ok(value)
}

fn bounded(value: &str, what: &str) -> Result<()> {
    if value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit(what));
    }
    Ok(())
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("slide XML node count"))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("slide XML node count"));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn classifies_reserved_powerpoint_action_values() {
        assert_eq!(
            classify_action(Some("ppaction://hlinkshowjump?jump=nextslide"), false),
            PptxActionKind::SlideShowJump(PptxSlideShowJump::NextSlide)
        );
        assert_eq!(
            classify_action(Some("ppaction://customshow?id=42"), false),
            PptxActionKind::CustomShow { id: 42 }
        );
        assert_eq!(
            classify_action(Some("ppaction://hlinkpres?slideindex=7"), true),
            PptxActionKind::Presentation {
                start_slide_index: 7
            }
        );
        assert_eq!(
            classify_action(Some("ppaction://macro?name=Module1.Run"), false),
            PptxActionKind::Macro
        );
        assert_eq!(
            classify_action(Some("urn:vendor:custom"), false),
            PptxActionKind::Unknown
        );
        assert_eq!(classify_action(None, true), PptxActionKind::Hyperlink);
        assert_eq!(classify_action(None, false), PptxActionKind::None);
    }
}
