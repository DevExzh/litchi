//! Checked requested-storage envelope for the byte-buffer MCE settings pass.
//!
//! The legacy MCE codec keeps its namespace and directive context behind
//! `Arc` links, but the layers themselves own vectors, strings, and hash
//! tables.  This module describes those owners without constructing any
//! collection.  The caller can populate [`Profile`] from a guarded source
//! preflight and reserve the returned number before invoking the codec.

use std::{collections::HashSet, mem::size_of, ops::Range, sync::Arc};

use litchi_ooxml_common::mce::{Capabilities, Name, Report};
use quick_xml::Reader;

use crate::settings::{WORD_2010_NAMESPACE, WORD_2012_NAMESPACE};

const INITIAL_NAMESPACE_BYTES: usize = 73;
const ARC_HEADER_WORDS: usize = 2;
const FRAME_LAYOUT_WORDS: usize = 12;
const DIAGNOSTIC_OVERHEAD_BYTES: usize = 128;

// This is the exact fixed namespace set installed by
// `Capabilities::ooxml_baseline`, followed by the two settings extension
// namespaces installed by `settings::extensions::process_bytes_with_limits`.
const SETTINGS_CAPABILITIES: [&str; 19] = [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://purl.oclc.org/ooxml/presentationml/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://purl.oclc.org/ooxml/drawingml/chart",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://purl.oclc.org/ooxml/officeDocument/math",
    "urn:schemas-microsoft-com:vml",
    "urn:schemas-microsoft-com:office:office",
    "http://www.w3.org/XML/1998/namespace",
    WORD_2010_NAMESPACE,
    WORD_2012_NAMESPACE,
];

/// Source-derived inputs to [`memory_requirement`].
///
/// `namespace_bytes` is the maximum live namespace-context byte count,
/// including the fixed resolver floor. `namespace_declarations` and
/// `directive_tokens` are source totals, so they safely overbound the number
/// of entries that can be distributed among the simultaneously live layers.
/// `directive_owned_bytes` is the total expanded URI/local-name payload
/// observed for directive targets.  `max_attributes_per_event` must include
/// `xmlns` attributes; a caller whose preflight excludes them should pass its
/// event-byte ceiling instead.
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct Profile {
    pub(super) source_bytes: usize,
    pub(super) output_bytes: usize,
    pub(super) max_token_bytes: usize,
    pub(super) max_depth: usize,
    pub(super) max_namespace_bindings: usize,
    pub(super) max_directive_tokens_per_event: usize,
    pub(super) max_choices_per_alternate: usize,
    pub(super) namespace_bytes: usize,
    pub(super) namespace_declarations: usize,
    pub(super) directive_tokens: usize,
    pub(super) directive_owned_bytes: usize,
    pub(super) max_attributes_per_event: usize,
}

/// Output storage retained after the MCE working state has been dropped.
pub(super) fn output_memory_requirement(bytes: usize) -> Option<u64> {
    vec_capacity(u64::try_from(bytes).ok()?, 1)
}

/// Return a checked requested-storage envelope for one legacy MCE pass.
///
/// This function only performs integer arithmetic over `Profile`; it does not
/// allocate, inspect input, or construct a collection.  The `V`, `Str`, and
/// `H` factors are pinned geometric capacity allowances.  They intentionally
/// cover the requested table/vector growth envelope rather than asserting a
/// portable process-RSS layout for the standard library allocator.
#[must_use]
pub(super) fn memory_requirement(profile: Profile) -> Option<u64> {
    // The scalar limits do not own a dynamic collection, but checking their
    // successor keeps this helper valid if a caller derives a profile term
    // from either limit later.
    let _ = profile.max_namespace_bindings.checked_add(1)?;
    let _ = profile.max_choices_per_alternate.checked_add(1)?;

    let source = u64::try_from(profile.source_bytes).ok()?;
    let output = u64::try_from(profile.output_bytes).ok()?;
    let token_bytes_usize = profile.max_token_bytes.min(profile.source_bytes);
    let token_bytes = u64::try_from(token_bytes_usize).ok()?;
    let token = token_bytes.checked_add(1)?;
    let depth = u64::try_from(profile.max_depth).ok()?;
    let levels = depth.checked_add(1)?;
    let attributes = u64::try_from(
        profile
            .max_attributes_per_event
            .max(1)
            .min(profile.source_bytes.max(1)),
    )
    .ok()?;
    let namespace_bytes =
        u64::try_from(profile.namespace_bytes.max(INITIAL_NAMESPACE_BYTES)).ok()?;
    let namespace_declarations =
        u64::try_from(profile.namespace_declarations.min(profile.source_bytes)).ok()?;
    let directive_tokens =
        u64::try_from(profile.directive_tokens.min(profile.source_bytes)).ok()?;
    let directive_tokens_per_event = u64::try_from(
        profile
            .max_directive_tokens_per_event
            .min(profile.directive_tokens)
            .min(profile.source_bytes),
    )
    .ok()?;
    let directive_owned_bytes = u64::try_from(profile.directive_owned_bytes).ok()?;
    let qualified_name_bytes = token_bytes.max(1);

    let usize_bytes = u64::try_from(size_of::<usize>()).ok()?;
    let string_bytes = u64::try_from(size_of::<String>()).ok()?;
    let name_bytes = u64::try_from(size_of::<Name>()).ok()?;
    let pair_bytes = u64::try_from(size_of::<(String, String)>()).ok()?;
    let directive_pair_bytes = u64::try_from(size_of::<(String, &str)>()).ok()?;
    let range_bytes = u64::try_from(size_of::<Range<usize>>()).ok()?;
    let reference_bytes = u64::try_from(size_of::<&str>()).ok()?;
    let hash_string_wrapper = u64::try_from(size_of::<HashSet<String>>()).ok()?;
    let hash_name_wrapper = u64::try_from(size_of::<HashSet<Name>>()).ok()?;
    let option_arc_bytes = u64::try_from(size_of::<Option<Arc<()>>>()).ok()?;
    let vector_pair_bytes = u64::try_from(size_of::<Vec<(String, String)>>()).ok()?;
    let reader_bytes = u64::try_from(size_of::<Reader<&[u8]>>()).ok()?;
    let report_bytes = u64::try_from(size_of::<Report>()).ok()?;

    let namespace_layers = namespace_declarations.min(depth);
    let directive_layers = directive_tokens.min(depth);
    let current_directive_entries = directive_tokens_per_event;

    // A NamePattern is private to the common crate.  An exact pattern owns a
    // public Name (two Strings), while a wildcard owns one String; the extra
    // usize covers the enum discriminant/alignment under the pinned layout.
    let pattern_slot = name_bytes
        .checked_add(string_bytes)?
        .checked_add(usize_bytes)?;

    let arc_header = usize_bytes.checked_mul(ARC_HEADER_WORDS as u64)?;
    let namespace_layer_object = option_arc_bytes.checked_add(vector_pair_bytes)?;
    let directive_layer_object =
        option_arc_bytes.checked_add(hash_string_wrapper.max(hash_name_wrapper).checked_mul(4)?)?;
    let namespace_layer_owner = arc_header.checked_add(namespace_layer_object)?;
    let directive_layer_owner = arc_header.checked_add(directive_layer_object)?;

    let event_buffer = vec_capacity(token, 1)?;
    let open_name_payload = source.min(levels.checked_mul(token)?);
    let reader_open_names =
        vec_capacity(open_name_payload, 1)?.checked_add(vec_capacity(levels, usize_bytes)?)?;

    let frame_item = usize_bytes.checked_mul(FRAME_LAYOUT_WORDS as u64)?;
    let mce_frames = vec_capacity(depth, frame_item)?.checked_add(string_capacity(
        source.checked_add(qualified_name_bytes.checked_mul(2)?)?,
        depth.checked_add(2)?,
    )?)?;

    let raw_attributes = string_capacity(token, attributes.checked_mul(2)?)?
        .checked_add(vec_capacity(attributes, pair_bytes)?)?;
    let quickxml_attributes = vec_capacity(attributes, range_bytes)?
        .checked_add(hash_capacity(attributes, size_of::<u64>() as u64)?)?;

    let namespace_storage =
        string_capacity(namespace_bytes, namespace_declarations.checked_mul(2)?)?
            .checked_add(vec_sum_capacity(
                namespace_declarations,
                namespace_layers,
                pair_bytes,
            )?)?
            .checked_add(namespace_layer_owner.checked_mul(namespace_layers)?)?;

    let directive_vector = string_capacity(token, attributes)?
        .checked_add(vec_capacity(attributes, directive_pair_bytes)?)?;

    // The preflight's total target bytes already includes every expanded
    // target URI/local pair.  Its extra lexical token bytes make this an
    // upper bound for retained NamePattern payloads.  Distinct ignorable URI
    // values are bounded by the live namespace bytes, while local_ignorable
    // retains another current-event copy.
    let directive_payload = string_capacity(namespace_bytes, directive_tokens)?
        .checked_add(string_capacity(namespace_bytes, current_directive_entries)?)?
        .checked_add(string_capacity(
            directive_owned_bytes.max(directive_tokens),
            directive_tokens.checked_mul(2)?,
        )?)?;

    let directive_tables = hash_capacity(current_directive_entries, string_bytes)?
        .checked_add(hash_capacity(current_directive_entries, reference_bytes)?)?
        .checked_add(hash_sum_capacity(
            directive_tokens,
            directive_layers,
            string_bytes,
        )?)?
        .checked_add(
            hash_sum_capacity(directive_tokens, directive_layers, pattern_slot)?.checked_mul(3)?,
        )?
        .checked_add(directive_layer_owner.checked_mul(directive_layers)?)?;

    let transient_names = string_capacity(namespace_bytes.checked_add(qualified_name_bytes)?, 2)?
        .checked_add(string_capacity(
            namespace_bytes.checked_add(qualified_name_bytes)?,
            2,
        )?)?
        .checked_add(string_capacity(
            source.checked_add(DIAGNOSTIC_OVERHEAD_BYTES as u64)?,
            2,
        )?)?;

    let output_owner = vec_capacity(output, 1)?;
    let capabilities_owner = capabilities_owner(string_bytes, pattern_slot)?;
    let fixed_state = reader_bytes
        .checked_add(usize_bytes.checked_mul(4)?)?
        .checked_add(report_bytes)?
        .checked_add(capabilities_owner)?;

    [
        event_buffer,
        reader_open_names,
        mce_frames,
        raw_attributes,
        quickxml_attributes,
        namespace_storage,
        directive_vector,
        directive_payload,
        directive_tables,
        transient_names,
        output_owner,
        fixed_state,
    ]
    .into_iter()
    .try_fold(0_u64, u64::checked_add)
}

fn string_capacity(bytes: u64, count: u64) -> Option<u64> {
    if count == 0 {
        return Some(0);
    }
    bytes.checked_mul(2)?.checked_add(count.checked_mul(8)?)
}

fn vec_capacity(count: u64, item_bytes: u64) -> Option<u64> {
    if count == 0 {
        return Some(0);
    }
    let floor = item_bytes.checked_mul(8)?;
    let geometric = count.checked_mul(item_bytes)?.checked_mul(2)?;
    Some(floor.max(geometric))
}

/// Sum `V<T>(entries_l)` over at most `layers` non-empty vectors whose total
/// entry count is at most `entries`.  The geometric payload and every
/// materialized vector's eight-entry floor are charged separately.
fn vec_sum_capacity(entries: u64, layers: u64, item_bytes: u64) -> Option<u64> {
    if entries == 0 || layers == 0 {
        return Some(0);
    }
    entries
        .checked_mul(item_bytes)?
        .checked_mul(2)?
        .checked_add(layers.checked_mul(item_bytes)?.checked_mul(8)?)
}

fn hash_capacity(count: u64, key_bytes: u64) -> Option<u64> {
    if count == 0 {
        return Some(0);
    }
    let entry_bytes = key_bytes.checked_add(1)?;
    let floor = entry_bytes.checked_mul(8)?;
    let geometric = count
        .checked_add(1)?
        .checked_mul(entry_bytes)?
        .checked_mul(4)?;
    Some(floor.max(geometric))
}

/// Sum `H<T>(entries_l)` over at most `layers` non-empty sets whose total
/// entry count is at most `entries`.  The sum charges four bytes per entry
/// slot plus the explicit eight-entry floor for every materialized table, so
/// it does not need the private per-layer counts.
fn hash_sum_capacity(entries: u64, layers: u64, key_bytes: u64) -> Option<u64> {
    if entries == 0 || layers == 0 {
        return Some(0);
    }
    let entry_bytes = key_bytes.checked_add(1)?;
    entries
        .checked_mul(entry_bytes)?
        .checked_mul(4)?
        .checked_add(layers.checked_mul(entry_bytes)?.checked_mul(8)?)
}

fn capabilities_owner(string_bytes: u64, pattern_slot: u64) -> Option<u64> {
    let mut payload_bytes = 0_u64;
    for namespace in &SETTINGS_CAPABILITIES {
        payload_bytes = payload_bytes.checked_add(u64::try_from(namespace.len()).ok()?)?;
    }
    let count = u64::try_from(SETTINGS_CAPABILITIES.len()).ok()?;
    let string_payload = string_capacity(payload_bytes, count)?;
    let understood_table = hash_capacity(count, string_bytes)?;
    let extension_table = hash_capacity(0, pattern_slot)?;
    u64::try_from(size_of::<Capabilities>())
        .ok()?
        .checked_add(string_payload)?
        .checked_add(understood_table)?
        .checked_add(extension_table)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn profile() -> Profile {
        Profile {
            source_bytes: 512,
            output_bytes: 512,
            max_token_bytes: 64,
            max_depth: 4,
            max_namespace_bindings: 8,
            max_directive_tokens_per_event: 16,
            max_choices_per_alternate: 8,
            namespace_bytes: 96,
            namespace_declarations: 3,
            directive_tokens: 2,
            directive_owned_bytes: 32,
            max_attributes_per_event: 6,
        }
    }

    #[test]
    fn small_profile_has_a_checked_owner_envelope() {
        let requirement = memory_requirement(profile());
        assert!(requirement.is_some_and(|bytes| bytes > 512));
    }

    #[test]
    fn expanded_directive_payload_increases_the_envelope() {
        let small = memory_requirement(profile()).expect("small profile fits");
        let mut expanded = profile();
        expanded.directive_owned_bytes = 4 * 1024;
        let large = memory_requirement(expanded).expect("expanded profile fits");
        assert!(large > small);
    }

    #[test]
    fn checked_arithmetic_rejects_unrepresentable_profiles() {
        let huge = Profile {
            source_bytes: usize::MAX,
            output_bytes: usize::MAX,
            max_token_bytes: usize::MAX,
            max_depth: usize::MAX,
            max_namespace_bindings: usize::MAX - 1,
            max_directive_tokens_per_event: usize::MAX,
            max_choices_per_alternate: usize::MAX - 1,
            namespace_bytes: usize::MAX,
            namespace_declarations: usize::MAX,
            directive_tokens: usize::MAX,
            directive_owned_bytes: usize::MAX,
            max_attributes_per_event: usize::MAX,
        };
        assert!(memory_requirement(huge).is_none());
    }
}
