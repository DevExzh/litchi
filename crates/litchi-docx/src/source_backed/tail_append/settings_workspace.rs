//! Requested-storage arithmetic for the bounded settings model path.
//!
//! This module is intentionally allocation free.  `tail_append` performs the
//! source preflight first, then uses this checked envelope before running the
//! full borrowed settings/mail-merge codecs.  The parent module owns the
//! source-derived facts; this helper only turns them into a finite byte term.

use std::mem::size_of;

use crate::mail_merge::{DataSourceObject, FieldMap, Settings as MailMergeSettings};
use crate::settings::{
    CompatibilityOption, CompatibilitySetting, DocumentSettings, Extension, Extensions,
    SmartTagType,
};
use quick_xml::reader::NsReader;

const MAX_FIELD_MAPS: usize = 16_384;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1_024;
const MAX_ATTACHED_TEMPLATE_TARGET_BYTES: usize = 32 * 1024;
const MAX_EXTENSIONS: usize = 128;
const MAX_OPAQUE_BYTES: usize = 4 * 1024 * 1024;

/// Return a checked requested-storage envelope for the full settings parser.
///
/// The mail-merge codec materializes a complete private `Node` tree, whose
/// layout is intentionally reproduced below from its three `String` fields,
/// two `Vec` fields, and one boolean.  The extension codec retains opaque
/// ranges and may build one self-contained replacement while the old range is
/// still live.  The direct settings model then retains its own vectors and
/// decoded strings.  Each phase is charged independently and the terms are
/// summed by the caller's conservative admission policy.
pub(super) fn model_memory_requirement(
    source_bytes: usize,
    facts: &super::SettingsGuardFacts,
) -> Option<u64> {
    let nodes = facts.node_count.max(1);
    let depth = facts.max_depth.max(1);
    let attributes = facts.attribute_count;
    let attribute_nodes = facts.attribute_bearing_nodes;
    let token = facts.max_token.max(1);
    let namespace_bytes = facts
        .max_namespace_buffer
        .max(super::INITIAL_QUICK_XML_NAMESPACE_BYTES);
    let namespace_bindings = facts.max_namespace_bindings.max(2);
    // Lossy UTF-8 and decoded String construction may grow geometrically.
    // Charge capacity for the counted name/namespace/value fields, including
    // the small nonempty String floor, rather than just their final lengths.
    let semantic_strings = nodes.checked_add(attributes)?.checked_mul(3)?;
    let semantic_bytes = facts
        .semantic_owned_bytes
        .checked_mul(2)?
        .checked_add(semantic_strings.checked_mul(8)?)?;

    let node_header = rounded_node_header()?;
    // `OwnedNamespace` has two String-bearing variants plus a unit variant;
    // charge an explicit discriminant-sized word in addition to the three
    // String payloads so this remains an upper bound on targets where the
    // enum does not fit in a String's niche.
    let attribute_header = size_of::<String>()
        .checked_mul(3)?
        .checked_add(size_of::<usize>())?;
    let child_slots = nodes.checked_mul(4)?;
    let stack_slots = depth.checked_mul(2)?;
    let attribute_slots = attributes
        .checked_mul(2)?
        .max(attribute_nodes.checked_mul(4)?);
    let reader = reader_storage(depth, token)?;

    // `mail_merge::codec::Node` and `Attribute` are private to that module.
    // These are conservative upper bounds for their actual headers.  The
    // geometric floors mirror the current Vec growth shape: four child slots
    // per observed node, two stack slots per admitted depth, and at least four
    // attribute slots for each attribute-bearing node.
    let tree_headers = nodes
        .checked_mul(node_header)?
        .checked_add(child_slots.checked_mul(node_header)?)?
        .checked_add(stack_slots.checked_mul(node_header)?)?
        .checked_add(attributes.checked_mul(attribute_header)?)?
        .checked_add(attribute_slots.checked_mul(attribute_header)?)?;
    let resolver = resolver_storage(namespace_bytes, namespace_bindings)?;
    let tree = tree_headers
        .checked_add(semantic_bytes)?
        .checked_add(token)?
        .checked_add(reader)?
        .checked_add(resolver)?;

    // `parse_mail_merge` clones selected Node attribute values while the tree
    // remains live.  The selected bytes are bounded by the complete semantic
    // byte total; relationship IDs have the mail-merge model's explicit
    // 1024-byte ceiling.  Field maps currently grow through Vec::push, hence
    // the two-slot geometric capacity term.
    let field_map_count = nodes.min(MAX_FIELD_MAPS);
    let field_map_slots = field_map_count.checked_mul(2)?;
    let mail_merge_fixed =
        size_of::<MailMergeSettings>().checked_add(size_of::<DataSourceObject>())?;
    let mail_merge_model = mail_merge_fixed
        .checked_add(field_map_slots.checked_mul(size_of::<FieldMap>())?)?
        .checked_add(semantic_bytes)?
        .checked_add(MAX_RELATIONSHIP_ID_BYTES.checked_mul(4)?)?
        .checked_add(token)?
        .checked_add(reader)?;

    // `Extensions::parse` retains copied unknown children.  `R <= source` is
    // the disjoint copied source range total.  Without an exact Q fact, bound
    // repeated self-contained declarations by every possible direct child,
    // the largest active binding byte layer, and a syntax/header term.  Each
    // unknown child is still subject to MAX_OPAQUE_BYTES in the model codec.
    let declaration_syntax = nodes
        .checked_mul(namespace_bindings)?
        .checked_mul(10usize.checked_add(size_of::<(Option<Vec<u8>>, Vec<u8>)>())?)?;
    let added_namespace_bytes = nodes
        .checked_mul(namespace_bytes)?
        .checked_mul(6)?
        .checked_add(declaration_syntax)?;
    let opaque_hard_bound = MAX_EXTENSIONS.checked_mul(MAX_OPAQUE_BYTES)?;
    let opaque_final = source_bytes
        .checked_add(added_namespace_bytes)?
        .min(opaque_hard_bound);
    let opaque_rewrite_overlap = source_bytes;
    let declared_prefixes = facts
        .max_attributes_per_event
        .checked_add(namespace_bindings)?
        .max(1);
    let active_binding_slots = namespace_bindings
        .checked_mul(size_of::<(Option<Vec<u8>>, Vec<u8>)>())?
        .checked_mul(2)?;
    let additions_slots = namespace_bindings
        .checked_mul(size_of::<&(Option<Vec<u8>>, Vec<u8>)>())?
        .checked_mul(2)?;
    let declared_slots = declared_prefixes
        .checked_mul(size_of::<Option<Vec<u8>>>())?
        .checked_mul(4)?;
    let extension_binding_scratch = namespace_bytes
        .checked_mul(2)?
        .checked_add(active_binding_slots)?
        .checked_add(additions_slots)?
        .checked_add(declared_slots)?
        .checked_add(size_of::<std::collections::HashSet<Option<Vec<u8>>>>())?
        .checked_add(size_of::<Vec<&(Option<Vec<u8>>, Vec<u8>)>>())?;
    let extension_values = MAX_EXTENSIONS
        .checked_mul(2)?
        .checked_mul(size_of::<Extension>())?;
    let extension_fixed = size_of::<Extensions>();
    // Opaque validation calls `into_owned()` on one event while its
    // self-contained output is retained.  Charge that event-sized transient
    // and the resolver that parses generated namespace declarations too.
    let opaque_event_transient = source_bytes
        .checked_add(added_namespace_bytes)?
        .min(MAX_OPAQUE_BYTES);
    let extension_validation_resolver = resolver_storage(
        opaque_event_transient.max(super::INITIAL_QUICK_XML_NAMESPACE_BYTES),
        declared_prefixes.max(2),
    )?;
    let extension = opaque_final
        .checked_add(opaque_rewrite_overlap)?
        .checked_add(extension_binding_scratch)?
        .checked_add(extension_values)?
        .checked_add(extension_fixed)?
        .checked_add(opaque_event_transient)?
        .checked_add(token)?
        .checked_add(reader)?
        .checked_add(resolver)?
        .checked_add(extension_validation_resolver)?;

    // The direct model is scalar-heavy, but its independently growing vectors
    // and decoded strings are real owners.  Counts are bounded by the guarded
    // element count when no more specific per-schema count is available.
    let model_slots = nodes.checked_mul(2)?;
    let direct_vectors = model_slots
        .checked_mul(size_of::<CompatibilityOption>())?
        .checked_add(model_slots.checked_mul(size_of::<CompatibilitySetting>())?)?
        .checked_add(model_slots.checked_mul(size_of::<SmartTagType>())?)?;
    let direct_model = size_of::<DocumentSettings>()
        .checked_add(direct_vectors)?
        .checked_add(semantic_bytes)?
        // attachedTemplate r:id has no 1024-byte codec ceiling. Its decoded
        // String is bounded by the guarded token, with a capacity allowance.
        .checked_add(token.checked_mul(2)?.checked_add(8)?)?
        .checked_add(MAX_ATTACHED_TEMPLATE_TARGET_BYTES)?
        .checked_add(token)?
        .checked_add(reader)?
        .checked_add(resolver)?;

    // The source bytes and MCE output are accounted by the caller's separate
    // MCE/source terms.  This function covers only the complete model phases;
    // all additions below are checked before conversion to the policy's u64.
    let total = tree
        .checked_add(mail_merge_model)?
        .checked_add(extension)?
        .checked_add(direct_model)?;
    u64::try_from(total).ok()
}

fn resolver_storage(namespace_bytes: usize, namespace_bindings: usize) -> Option<usize> {
    let binding_bytes = super::QUICK_XML_NAMESPACE_BINDING_BYTES;
    namespace_bytes
        .checked_mul(4)?
        // The reader retains geometric capacity while each model event
        // holds an exact-length clone of the resolver's binding vector.
        .checked_add(namespace_bindings.checked_mul(binding_bytes.checked_mul(4)?)?)
}

fn reader_storage(depth: usize, token: usize) -> Option<usize> {
    // `NsReader` contains the fixed Reader/config/resolver state.  The reader
    // also grows `opened_buffer` and `opened_starts` independently of the
    // resolver's namespace buffer/binding vectors; retain those capacities as
    // separate terms so resolver charging cannot hide this owner.
    let levels = depth.checked_add(1)?;
    let opened_names = levels.checked_mul(token)?.checked_mul(2)?;
    let opened_indexes = levels.checked_mul(size_of::<usize>())?.checked_mul(2)?;
    size_of::<NsReader<&[u8]>>()
        .checked_add(opened_names)?
        .checked_add(opened_indexes)
}

fn rounded_node_header() -> Option<usize> {
    let raw = size_of::<String>()
        .checked_mul(3)?
        .checked_add(size_of::<Vec<usize>>().checked_mul(2)?)?
        .checked_add(size_of::<bool>())?;
    let alignment = size_of::<usize>();
    raw.checked_add(alignment.saturating_sub(1))
        .and_then(|value| value.checked_div(alignment))
        .and_then(|value| value.checked_mul(alignment))
}
