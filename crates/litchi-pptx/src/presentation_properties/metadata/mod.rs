//! Layered PresentationML document and presentation metadata.
//!
//! Each child owns one bounded semantic capability. Its `model` layer keeps
//! package-independent values, `codec` owns XML, and `package` owns OPC graph
//! traversal and mutation. The module is attached below the existing
//! presentation-properties facade until the crate-level facade integrator wires
//! the final public path.

mod slide_patch;

use std::borrow::Cow;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part, SourceLimits};

pub mod changes;
pub mod color_map;
pub mod custom_show;
pub mod designer_tags;
pub mod events;
pub mod guides;
pub mod handout;
pub mod protection;
pub mod revision;
pub use super::read_only_recommended;
pub mod sections;
pub mod slide_sync;
pub mod structure;
pub mod template;
pub mod tracks;

/// Caller-selected resource policy shared by the presentation metadata
/// codecs.
///
/// The metadata owners are deliberately small, but their XML is still
/// attacker-controlled.  Keeping one policy value for the family means that
/// a package opened with a tighter source profile carries the same bounds
/// through detached parsing, source-preserving edits and patch replay.  The
/// format defaults below are only the historical maxima; an operation always
/// receives an explicit policy and never consults a process-wide ceiling.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_depth: usize,
    max_nodes: usize,
    max_string_bytes: usize,
    max_extensions: usize,
    max_custom_shows: usize,
    max_guides: usize,
    max_sections: usize,
    max_slide_references: usize,
    max_attributes: usize,
    max_relationships: usize,
}

impl Limits {
    /// Historical upper bounds used by the metadata codecs.
    pub const DEFAULT: Self = Self {
        max_input_bytes: 8 * 1024 * 1024,
        max_output_bytes: 8 * 1024 * 1024,
        max_depth: 128,
        max_nodes: 100_000,
        max_string_bytes: 1024 * 1024,
        max_extensions: 1_024,
        max_custom_shows: 4_096,
        max_guides: 16_384,
        max_sections: 4_096,
        max_slide_references: 100_000,
        max_attributes: 20_000,
        max_relationships: 100_000,
    };

    /// Derive metadata bounds from the retained physical source policy.
    ///
    /// The semantic maxima are never widened when a caller selected a tighter
    /// source profile.  This conversion is cheap and copy-only, so callers can
    /// derive it at an edit or replay boundary without cloning an owner.
    #[must_use = "retain the derived metadata policy for the operation"]
    pub fn from_source_limits(source: SourceLimits) -> Self {
        let default = Self::DEFAULT;
        let part_bytes = usize::try_from(source.max_part_bytes()).unwrap_or(usize::MAX);
        let events = source.max_events();
        let depth = source.max_depth();
        let attributes = source.max_xml_attribute_bytes();
        Self {
            max_input_bytes: default.max_input_bytes.min(part_bytes),
            max_output_bytes: default.max_output_bytes.min(part_bytes),
            max_depth: default.max_depth.min(depth),
            max_nodes: default.max_nodes.min(events),
            max_string_bytes: default.max_string_bytes.min(attributes),
            max_extensions: default.max_extensions.min(events),
            max_custom_shows: default.max_custom_shows.min(events),
            max_guides: default.max_guides.min(events),
            max_sections: default.max_sections.min(events),
            max_slide_references: default.max_slide_references.min(events),
            max_attributes: default.max_attributes.min(events),
            max_relationships: default.max_relationships.min(events),
        }
    }

    /// Construct a policy with explicit dimensions.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_depth: usize,
        max_nodes: usize,
        max_string_bytes: usize,
        max_extensions: usize,
        max_custom_shows: usize,
        max_guides: usize,
        max_sections: usize,
        max_slide_references: usize,
        max_attributes: usize,
        max_relationships: usize,
    ) -> Option<Self> {
        let default = Self::DEFAULT;
        if max_input_bytes == 0
            || max_input_bytes > default.max_input_bytes
            || max_output_bytes == 0
            || max_output_bytes > default.max_output_bytes
            || max_depth == 0
            || max_depth > default.max_depth
            || max_nodes == 0
            || max_nodes > default.max_nodes
            || max_string_bytes == 0
            || max_string_bytes > default.max_string_bytes
            || max_extensions == 0
            || max_extensions > default.max_extensions
            || max_custom_shows == 0
            || max_custom_shows > default.max_custom_shows
            || max_guides == 0
            || max_guides > default.max_guides
            || max_sections == 0
            || max_sections > default.max_sections
            || max_slide_references == 0
            || max_slide_references > default.max_slide_references
            || max_attributes == 0
            || max_attributes > default.max_attributes
            || max_relationships == 0
            || max_relationships > default.max_relationships
        {
            return None;
        }
        Some(Self {
            max_input_bytes,
            max_output_bytes,
            max_depth,
            max_nodes,
            max_string_bytes,
            max_extensions,
            max_custom_shows,
            max_guides,
            max_sections,
            max_slide_references,
            max_attributes,
            max_relationships,
        })
    }

    #[inline]
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }
    #[inline]
    #[must_use]
    pub const fn with_input_bytes(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_input_bytes).then_some(Self {
            max_input_bytes: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
    #[inline]
    #[must_use]
    pub const fn with_output_bytes(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_output_bytes).then_some(Self {
            max_output_bytes: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }
    #[inline]
    #[must_use]
    pub const fn with_depth(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_depth).then_some(Self {
            max_depth: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_nodes(self) -> usize {
        self.max_nodes
    }
    #[inline]
    #[must_use]
    pub const fn with_nodes(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_nodes).then_some(Self {
            max_nodes: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_string_bytes(self) -> usize {
        self.max_string_bytes
    }
    #[inline]
    #[must_use]
    pub const fn with_string_bytes(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_string_bytes).then_some(Self {
            max_string_bytes: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_extensions(self) -> usize {
        self.max_extensions
    }
    #[inline]
    #[must_use]
    pub const fn with_extensions(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_extensions).then_some(Self {
            max_extensions: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_custom_shows(self) -> usize {
        self.max_custom_shows
    }
    #[inline]
    #[must_use]
    pub const fn with_custom_shows(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_custom_shows).then_some(Self {
            max_custom_shows: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_guides(self) -> usize {
        self.max_guides
    }
    #[inline]
    #[must_use]
    pub const fn with_guides(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_guides).then_some(Self {
            max_guides: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_sections(self) -> usize {
        self.max_sections
    }
    #[inline]
    #[must_use]
    pub const fn with_sections(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_sections).then_some(Self {
            max_sections: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_slide_references(self) -> usize {
        self.max_slide_references
    }
    #[inline]
    #[must_use]
    pub const fn with_slide_references(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_slide_references).then_some(Self {
            max_slide_references: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_attributes(self) -> usize {
        self.max_attributes
    }
    #[inline]
    #[must_use]
    pub const fn with_attributes(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_attributes).then_some(Self {
            max_attributes: value,
            ..self
        })
    }
    #[inline]
    #[must_use]
    pub const fn max_relationships(self) -> usize {
        self.max_relationships
    }
    #[inline]
    #[must_use]
    pub const fn with_relationships(self, value: usize) -> Option<Self> {
        (value > 0 && value <= Self::DEFAULT.max_relationships).then_some(Self {
            max_relationships: value,
            ..self
        })
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Derive the selected metadata profile at a package boundary.
#[inline]
pub(crate) fn limits_for(package: &OpcPackage) -> Limits {
    Limits::from_source_limits(package.read_limits())
}

/// Metadata owner whose changed payload receives family-specific replay
/// admission.  The common patch engine already checks the exact before image,
/// relationship graph, and recipient physical source limits; these families
/// add the inexpensive semantic check for the changed owner only.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum ReplayFamily {
    Changes,
    ColorMaps,
    CustomShows,
    DesignerTags,
    Events,
    Guides,
    Revision,
    SlideSync,
    Structure,
    Tracks,
}

/// Admit one changed metadata payload under the recipient's retained source
/// policy before the generic package patch is published.
///
/// The physical scan is deliberately performed only for changed XML.  A
/// no-op patch never calls this hook, and relationship-only patches do not
/// allocate or rescan unchanged parts.  Family codecs then validate the
/// changed owner where they expose a detached parser; package-aware families
/// inspect only the candidate owner needed by their existing read path.
pub(crate) fn validate_replay_part(
    package: &OpcPackage,
    family: ReplayFamily,
    part_name: &PackURI,
    part: &dyn Part,
) -> crate::Result<()> {
    let source_limits = package.read_limits();
    if part.content_type() == crate::presentation_properties::CONTENT_TYPE {
        return crate::presentation_properties::validate_replay_part(
            part,
            crate::presentation_properties::Limits::from_source_limits(source_limits),
        );
    }

    if part.content_type().ends_with("+xml") {
        crate::source_limits::validate_xml_limits(part.blob(), source_limits)?;
    }

    match family {
        ReplayFamily::Changes if part.content_type() == changes::CONTENT_TYPE => {
            changes::Info::parse(part.blob()).map(drop)
        },
        ReplayFamily::Revision if part.content_type() == revision::CONTENT_TYPE => {
            revision::Info::parse(part.blob()).map(drop)
        },
        ReplayFamily::SlideSync if part.content_type() == slide_sync::CONTENT_TYPE => {
            slide_sync::Properties::parse(part.blob()).map(drop)
        },
        ReplayFamily::Tracks if part.content_type() == tracks::CONTENT_TYPE => {
            tracks::File::parse(part.blob()).map(drop)
        },
        ReplayFamily::CustomShows if is_presentation_content_type(part.content_type()) => {
            let xml = std::str::from_utf8(part.blob())
                .map_err(|_| crate::Error::Invalid("PresentationML is not UTF-8".into()))?;
            custom_show::List::parse_xml(xml, Limits::from_source_limits(source_limits)).map(drop)
        },
        ReplayFamily::ColorMaps => {
            validate_color_map(part, Limits::from_source_limits(source_limits))
        },
        ReplayFamily::Events if part.content_type() == ct::PML_SLIDE => {
            events::load(package, part_name).map(drop)
        },
        ReplayFamily::Guides if is_presentation_content_type(part.content_type()) => {
            guides::load(package, guides::Limits::from_source_limits(source_limits)).map(drop)
        },
        ReplayFamily::Structure if is_presentation_content_type(part.content_type()) => {
            structure::load(
                package,
                structure::Limits::from_source_limits(source_limits),
            )
            .map(drop)
        },
        // Designer-tag ownership is selected by a stable slide ID inside the
        // presentation owner.  Its changed-payload source scan above still
        // enforces the recipient physical XML policy; the edit's source-bound
        // transaction remains the authoritative selected-owner validator.
        _ => Ok(()),
    }
}

/// Process one metadata XML source with the caller's output and tree policy.
///
/// The generic `process_ooxml` convenience function uses the shared, much
/// larger MCE defaults. Metadata operations bind MCE admission to the
/// selected owner profile before the processor allocates a marked or selected
/// branch, so all eight raw metadata families use this adapter instead.
pub(crate) fn process_ooxml<'a>(xml: &'a [u8], limits: Limits) -> crate::Result<Cow<'a, [u8]>> {
    use litchi_ooxml_common::mce::{
        Capabilities, Limits as MceLimits, process_markup_compatibility,
    };

    let defaults = MceLimits::default();
    let mce_limits = MceLimits {
        max_input_bytes: defaults.max_input_bytes.min(limits.max_input_bytes()),
        max_output_bytes: defaults.max_output_bytes.min(limits.max_output_bytes()),
        max_depth: defaults.max_depth.min(limits.max_depth()),
        max_namespace_bindings: defaults.max_namespace_bindings.min(limits.max_nodes()),
        max_directive_tokens: defaults.max_directive_tokens.min(limits.max_nodes()),
        max_choices_per_alternate: defaults.max_choices_per_alternate.min(limits.max_nodes()),
        max_attributes_per_element: defaults
            .max_attributes_per_element
            .min(limits.max_attributes()),
    };
    Ok(process_markup_compatibility(xml, &Capabilities::ooxml_baseline(), &mce_limits)?.xml)
}

/// String form of [`process_ooxml`], retaining the borrowed fast path.
pub(crate) fn process_str<'a>(xml: &'a str, limits: Limits) -> crate::Result<Cow<'a, str>> {
    match process_ooxml(xml.as_bytes(), limits)? {
        Cow::Borrowed(_) => Ok(Cow::Borrowed(xml)),
        Cow::Owned(value) => String::from_utf8(value)
            .map(Cow::Owned)
            .map_err(|error| crate::Error::Xml(format!("MCE output is not UTF-8: {error}"))),
    }
}

fn validate_color_map(part: &dyn Part, limits: Limits) -> crate::Result<()> {
    match part.content_type() {
        ct::PML_SLIDE_MASTER => color_map::parse_master(part.blob(), limits).map(drop),
        ct::PML_SLIDE_LAYOUT => {
            color_map::parse_override(part.blob(), b"sldLayout", "slide layout", limits).map(drop)
        },
        ct::PML_SLIDE => color_map::parse_override(part.blob(), b"sld", "slide", limits).map(drop),
        _ => Ok(()),
    }
}

fn is_presentation_content_type(content_type: &str) -> bool {
    matches!(
        content_type,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    )
}

pub(crate) fn escape_xml(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

pub(crate) fn new_guid() -> String {
    use std::sync::atomic::{AtomicU64, Ordering};
    use std::time::{SystemTime, UNIX_EPOCH};

    static NEXT: AtomicU64 = AtomicU64::new(1);
    let sequence = NEXT.fetch_add(1, Ordering::Relaxed);
    let ticks = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_or(0, |duration| duration.as_nanos() as u64);
    let mixed = ticks ^ sequence.rotate_left(17);
    format!(
        "{{{mixed:08X}-{part2:04X}-{part3:04X}-{part4:04X}-{part5:012X}}}",
        part2 = (mixed >> 32) as u16,
        part3 = (mixed >> 16) as u16,
        part4 = mixed as u16,
        part5 = mixed.wrapping_mul(0x9E37_79B9_7F4A_7C15) & 0x0000_FFFF_FFFF_FFFF,
    )
}

pub(crate) fn is_presentationml_name(
    namespace: &quick_xml::name::ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local_name: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            *value == b"http://schemas.openxmlformats.org/presentationml/2006/main"
                || *value == b"http://purl.oclc.org/ooxml/presentationml/main"
        },
        // Element identity is determined by the resolved namespace URI, never
        // by the producer's chosen prefix. An unresolved conventional prefix
        // is malformed context for a PresentationML element.
        quick_xml::name::ResolveResult::Unknown(_) | quick_xml::name::ResolveResult::Unbound => {
            false
        },
    }
}
