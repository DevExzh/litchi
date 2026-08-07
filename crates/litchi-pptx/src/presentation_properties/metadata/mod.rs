//! Layered PresentationML document and presentation metadata.
//!
//! Each child owns one bounded semantic capability. Its `model` layer keeps
//! package-independent values, `codec` owns XML, and `package` owns OPC graph
//! traversal and mutation. The module is attached below the existing
//! presentation-properties facade until the crate-level facade integrator wires
//! the final public path.

mod slide_patch;

pub mod changes;
pub mod color_map;
pub mod custom_show;
pub mod designer_tags;
pub mod events;
pub mod guides;
pub mod handout;
pub mod protection;
pub mod revision;
pub mod sections;
pub mod slide_sync;
pub mod structure;
pub mod template;
pub mod tracks;

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
