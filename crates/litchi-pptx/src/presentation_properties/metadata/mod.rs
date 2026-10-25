//! Layered PresentationML document and presentation metadata.
//!
//! Each child owns one bounded semantic capability. Its `model` layer keeps
//! package-independent values, `codec` owns XML, and `package` owns OPC graph
//! traversal and mutation. The module is attached below the existing
//! presentation-properties facade until the crate-level facade integrator wires
//! the final public path.

mod slide_patch;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};

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
    validate_replay_part_policy(
        package,
        family,
        part_name,
        part,
        crate::open_options::SemanticLimits::default(),
    )
}

pub(crate) fn validate_replay_part_policy(
    package: &OpcPackage,
    family: ReplayFamily,
    _part_name: &PackURI,
    part: &dyn Part,
    semantic_limits: crate::open_options::SemanticLimits,
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
            changes::Info::parse(
                part.blob(),
                semantic_limits
                    .changes()
                    .constrained_by_source(source_limits),
            )
            .map(drop)
        },
        ReplayFamily::Revision if part.content_type() == revision::CONTENT_TYPE => {
            revision::Info::parse(
                part.blob(),
                semantic_limits
                    .revision()
                    .constrained_by_source(source_limits),
            )
            .map(drop)
        },
        ReplayFamily::SlideSync if part.content_type() == slide_sync::CONTENT_TYPE => {
            slide_sync::Properties::parse(
                part.blob(),
                semantic_limits
                    .slide_sync()
                    .constrained_by_source(source_limits),
            )
            .map(drop)
        },
        ReplayFamily::Tracks if part.content_type() == ct::PML_SLIDE => tracks::validate_slide_xml(
            part.blob(),
            semantic_limits
                .tracks()
                .constrained_by_source(source_limits),
        ),
        ReplayFamily::Tracks if part.content_type() == tracks::CONTENT_TYPE => tracks::File::parse(
            part.blob(),
            semantic_limits
                .tracks()
                .constrained_by_source(source_limits),
        )
        .map(drop),
        ReplayFamily::CustomShows if is_presentation_content_type(part.content_type()) => {
            let xml = std::str::from_utf8(part.blob())
                .map_err(|_| crate::Error::Invalid("PresentationML is not UTF-8".into()))?;
            custom_show::List::parse_xml(xml).map(drop)
        },
        ReplayFamily::ColorMaps => validate_color_map(part),
        ReplayFamily::Events if part.content_type() == ct::PML_SLIDE => events::validate_slide_xml(
            part,
            semantic_limits
                .events()
                .constrained_by_source(source_limits),
        ),
        ReplayFamily::Guides if is_presentation_content_type(part.content_type()) => {
            guides::load(package).map(drop)
        },
        ReplayFamily::Structure if is_presentation_content_type(part.content_type()) => {
            structure::load(package).map(drop)
        },
        // Designer-tag ownership is selected by a stable slide ID inside the
        // presentation owner.  Its changed-payload source scan above still
        // enforces the recipient physical XML policy; the edit's source-bound
        // transaction remains the authoritative selected-owner validator.
        _ => Ok(()),
    }
}

fn validate_color_map(part: &dyn Part) -> crate::Result<()> {
    match part.content_type() {
        ct::PML_SLIDE_MASTER => color_map::parse_master(part.blob()).map(drop),
        ct::PML_SLIDE_LAYOUT => {
            color_map::parse_override(part.blob(), b"sldLayout", "slide layout").map(drop)
        },
        ct::PML_SLIDE => color_map::parse_override(part.blob(), b"sld", "slide").map(drop),
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
