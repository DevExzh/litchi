//! Member- and object-level comparison of an input package and the host and
//! leaf outputs of one operation.
//!
//! The host output defines what the operation touched: a ZIP member or IWA
//! object whose content the host left equal to the input is *untouched*,
//! and the leaf must keep it byte-identical to the input. Everything the
//! host changed must be equal in the leaf output on every known protobuf
//! field (through [`KnownFields`]). Violations fail Tier B; notes (ZIP
//! layout, unknown fields, object order) are reported only.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;

use litchi_iwa_archive::iwa::{Archive, ArchiveInfo, RawMessage, SnappyStream};
use litchi_iwa_archive::package::Catalog;

use crate::harness::TestResult;
use crate::known_fields::KnownFields;

/// One ZIP member: decoded content and exact physical records.
pub struct Member {
    /// Member name in source order.
    pub name: String,
    /// Decoded (uncompressed) payload.
    pub data: Vec<u8>,
    /// Local record and compressed payload, followed by the central
    /// header's metadata (the central record itself is not compared: it
    /// carries the local-header offset, which moves when an earlier member
    /// changes size).
    pub physical: Vec<u8>,
    /// Compression method, sizes and CRC, for readable notes.
    pub layout: String,
}

/// Read every member of a package, in order.
pub fn members(bytes: &[u8]) -> TestResult<Vec<Member>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut members = Vec::new();
    for entry in catalog.iter() {
        let record = entry.raw_record();
        let metadata = entry.metadata();
        let mut physical = record.local_record().to_vec();
        physical.extend_from_slice(record.compressed_data());
        physical.extend_from_slice(format!("{:?}", metadata.central()).as_bytes());
        members.push(Member {
            name: entry.name().to_owned(),
            data: entry.data().to_vec(),
            physical,
            layout: format!(
                "method {} compressed {} uncompressed {} crc {:08x}",
                metadata.central().compression_method(),
                metadata.compressed_size(),
                metadata.uncompressed_size(),
                metadata.crc32()
            ),
        });
    }
    Ok(members)
}

/// One IWA object of a component.
pub struct Object {
    /// Object identifier (`ArchiveInfo.identifier`).
    pub identifier: u64,
    /// Decoded header.
    pub info: ArchiveInfo,
    /// Payloads in order.
    pub messages: Vec<RawMessage>,
    /// Exact bytes of the object (length prefix, header and payloads).
    pub raw: Vec<u8>,
}

/// Parse the objects of one Snappy-framed IWA component, in order.
pub fn objects(component: &[u8]) -> TestResult<Vec<Object>> {
    let stream = SnappyStream::decompress(component, Default::default())?;
    let bytes = stream.as_bytes();
    let archive = Archive::parse(bytes, Default::default())?;
    let mut objects = Vec::new();
    for object in archive.objects {
        let identifier = object
            .archive_info
            .identifier
            .ok_or("IWA object without an identifier")?;
        let start = usize::try_from(object.header_offset)?;
        let end = usize::try_from(object.data_offset + object.data_length)?;
        let raw = bytes
            .get(start..end)
            .ok_or("IWA object span out of range")?
            .to_vec();
        objects.push(Object {
            identifier,
            info: object.archive_info,
            messages: object.messages,
            raw,
        });
    }
    Ok(objects)
}

/// The outcome of one three-way comparison.
#[derive(Default)]
pub struct Diff {
    /// Differences that fail Tier B.
    pub violations: Vec<String>,
    /// Differences that are reported only.
    pub notes: Vec<String>,
    /// Members whose content the host changed, added or removed.
    pub host_changed_members: Vec<String>,
    /// Objects whose bytes the host changed, added or removed.
    pub host_changed_objects: usize,
    /// Objects whose bytes the leaf changed, added or removed.
    pub leaf_changed_objects: usize,
}

impl Diff {
    /// Multi-line rendering for failure output.
    pub fn render(&self) -> String {
        let mut text = String::new();
        let _ = writeln!(
            text,
            "  members changed by host: {:?}; objects changed: host {}, leaf {}",
            self.host_changed_members, self.host_changed_objects, self.leaf_changed_objects
        );
        for violation in &self.violations {
            let _ = writeln!(text, "  VIOLATION {violation}");
        }
        for note in &self.notes {
            let _ = writeln!(text, "  note {note}");
        }
        text
    }
}

/// Compare `input`, the host output and the leaf output member by member
/// and object by object.
///
/// `leaf_drops` names members the leaf is documented to delete where the
/// host keeps them (Keynote's root previews, see
/// [`crate::harness::Case::leaf_invalidates_root_previews`]): when such a
/// member is in the input and the host output and absent from the leaf
/// output, the difference is a note instead of a violation.
pub fn compare(input: &[u8], host: &[u8], leaf: &[u8], known: &KnownFields) -> TestResult<Diff> {
    compare_with_policies(input, host, leaf, known, &[], MetadataPolicy::default())
}

/// [`compare`] with the documented leaf policies a case accepts.
pub fn compare_with_policies(
    input: &[u8],
    host: &[u8],
    leaf: &[u8],
    known: &KnownFields,
    leaf_drops: &[&str],
    policy: MetadataPolicy,
) -> TestResult<Diff> {
    compare_with_policies_except_objects(
        input,
        host,
        leaf,
        known,
        leaf_drops,
        policy,
        &BTreeSet::new(),
    )
}

/// Compare with a narrowly scoped reference-order policy.
pub fn compare_with_reference_policy(
    input: &[u8],
    host: &[u8],
    leaf: &[u8],
    known: &KnownFields,
    leaf_drops: &[&str],
    policy: MetadataPolicy,
    reference_policy: ReferenceOrderPolicy,
) -> TestResult<Diff> {
    compare_with_policies_except_objects_and_members_with_reference_policy(
        input,
        host,
        leaf,
        known,
        leaf_drops,
        policy,
        &BTreeSet::new(),
        &[],
        reference_policy,
    )
}

/// [`compare_with_policies`] with a documented host-only object island.
///
/// The excluded identifiers are still subject to the caller's graph proof;
/// this only permits the host to retain those objects while the leaf removes
/// them. Every surviving object continues through the normal raw, reference,
/// and known-field checks, and untouched ZIP members remain byte-exact.
pub fn compare_with_policies_except_objects(
    input: &[u8],
    host: &[u8],
    leaf: &[u8],
    known: &KnownFields,
    leaf_drops: &[&str],
    policy: MetadataPolicy,
    excluded_object_ids: &BTreeSet<u64>,
) -> TestResult<Diff> {
    compare_with_policies_except_objects_and_members_with_reference_policy(
        input,
        host,
        leaf,
        known,
        leaf_drops,
        policy,
        excluded_object_ids,
        &[],
        ReferenceOrderPolicy::Exact,
    )
}

/// [`compare_with_policies_except_objects`] with an explicitly documented set
/// of host-only members belonging to the same allowed orphan island.
pub fn compare_with_policies_except_objects_and_members(
    input: &[u8],
    host: &[u8],
    leaf: &[u8],
    known: &KnownFields,
    leaf_drops: &[&str],
    policy: MetadataPolicy,
    excluded_object_ids: &BTreeSet<u64>,
    allowed_host_members: &[&str],
) -> TestResult<Diff> {
    compare_with_policies_except_objects_and_members_with_reference_policy(
        input,
        host,
        leaf,
        known,
        leaf_drops,
        policy,
        excluded_object_ids,
        allowed_host_members,
        ReferenceOrderPolicy::Exact,
    )
}

fn compare_with_policies_except_objects_and_members_with_reference_policy(
    input: &[u8],
    host: &[u8],
    leaf: &[u8],
    known: &KnownFields,
    leaf_drops: &[&str],
    policy: MetadataPolicy,
    excluded_object_ids: &BTreeSet<u64>,
    allowed_host_members: &[&str],
    reference_policy: ReferenceOrderPolicy,
) -> TestResult<Diff> {
    let input = members(input)?;
    let mut host = members(host)?;
    let leaf = members(leaf)?;
    let mut diff = Diff::default();
    if !leaf_drops.is_empty() {
        let in_leaf: BTreeSet<_> = leaf.iter().map(|m| m.name.clone()).collect();
        let in_input: BTreeSet<_> = input.iter().map(|m| m.name.clone()).collect();
        let before = host.len();
        host.retain(|member| {
            !(leaf_drops.contains(&member.name.as_str())
                && in_input.contains(&member.name)
                && !in_leaf.contains(&member.name))
        });
        if host.len() != before {
            diff.notes.push(format!(
                "the leaf deleted {} root preview member(s) the host kept (documented rendering invalidation)",
                before - host.len()
            ));
        }
    }

    let names = |members: &[Member]| members.iter().map(|m| m.name.clone()).collect::<Vec<_>>();
    let (input_names, host_names, leaf_names) = (names(&input), names(&host), names(&leaf));
    if host_names != leaf_names {
        let host_set: BTreeSet<_> = host_names.iter().collect();
        let leaf_set: BTreeSet<_> = leaf_names.iter().collect();
        let only_host: Vec<_> = host_set
            .difference(&leaf_set)
            .filter(|name| !allowed_host_members.contains(&name.as_str()))
            .collect();
        let only_leaf: Vec<_> = leaf_set.difference(&host_set).collect();
        if only_host.is_empty() && only_leaf.is_empty() {
            diff.notes.push(format!(
                "ZIP member order differs: host {host_names:?}, leaf {leaf_names:?}"
            ));
        } else {
            diff.violations.push(format!(
                "ZIP member sets differ: only in host {only_host:?}, only in leaf {only_leaf:?}"
            ));
        }
    }
    if host_names != input_names {
        diff.notes.push(format!(
            "host changed the ZIP member list: input {} members, host {}",
            input_names.len(),
            host_names.len()
        ));
    }

    let by_name = |members: &[Member]| -> BTreeMap<String, usize> {
        members
            .iter()
            .enumerate()
            .map(|(index, member)| (member.name.clone(), index))
            .collect()
    };
    let (input_at, host_at, leaf_at) = (by_name(&input), by_name(&host), by_name(&leaf));
    for name in &host_names {
        let Some(&host_index) = host_at.get(name) else {
            continue;
        };
        let Some(&leaf_index) = leaf_at.get(name) else {
            continue;
        };
        let source = input_at.get(name).map(|&index| &input[index]);
        let (h, l) = (&host[host_index], &leaf[leaf_index]);
        let host_untouched = source.is_some_and(|source| source.data == h.data);
        if !host_untouched {
            diff.host_changed_members.push(name.clone());
        }
        if h.data == l.data {
            if host_untouched {
                if let Some(source) = source
                    && source.physical != l.physical
                {
                    diff.violations.push(format!(
                        "{name}: untouched member re-written by the leaf (content equal; input {}, leaf {})",
                        source.layout, l.layout
                    ));
                }
            } else if h.physical != l.physical {
                diff.notes.push(format!(
                    "{name}: ZIP layout differs, content equal (host {}, leaf {})",
                    h.layout, l.layout
                ));
            }
            if name.ends_with(".iwa") && !host_untouched {
                let host_objects = objects(&h.data)?;
                let source_objects = match source {
                    Some(source) => objects(&source.data)?,
                    None => Vec::new(),
                };
                let changed = count_changed(&source_objects, &host_objects, excluded_object_ids);
                diff.host_changed_objects += changed;
                diff.leaf_changed_objects += changed;
            }
            continue;
        }
        if !name.ends_with(".iwa") {
            diff.violations.push(format!(
                "{name}: non-IWA member differs from the host ({} vs {} bytes, first difference at {:?}){}",
                h.data.len(),
                l.data.len(),
                first_difference(&h.data, &l.data),
                if host_untouched {
                    "; the host left it untouched"
                } else {
                    ""
                }
            ));
            continue;
        }
        let source_objects = match source {
            Some(source) => objects(&source.data)?,
            None => Vec::new(),
        };
        compare_objects(
            name,
            &source_objects,
            &objects(&h.data)?,
            &objects(&l.data)?,
            known,
            policy,
            reference_policy,
            &mut diff,
            excluded_object_ids,
            allowed_host_members,
        );
    }
    Ok(diff)
}

fn count_changed(
    source: &[Object],
    target: &[Object],
    excluded_object_ids: &BTreeSet<u64>,
) -> usize {
    let source: BTreeMap<_, _> = source.iter().map(|o| (o.identifier, &o.raw)).collect();
    let target: BTreeMap<_, _> = target.iter().map(|o| (o.identifier, &o.raw)).collect();
    source
        .keys()
        .chain(target.keys())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .filter(|id| !excluded_object_ids.contains(*id))
        .filter(|id| source.get(id) != target.get(id))
        .count()
}

fn compare_objects(
    member: &str,
    input: &[Object],
    host: &[Object],
    leaf: &[Object],
    known: &KnownFields,
    policy: MetadataPolicy,
    reference_policy: ReferenceOrderPolicy,
    diff: &mut Diff,
    excluded_object_ids: &BTreeSet<u64>,
    allowed_host_members: &[&str],
) {
    diff.host_changed_objects += count_changed(input, host, excluded_object_ids);
    diff.leaf_changed_objects += count_changed(input, leaf, excluded_object_ids);
    let index = |objects: &[Object]| -> BTreeMap<u64, usize> {
        objects
            .iter()
            .enumerate()
            .map(|(position, object)| (object.identifier, position))
            .collect()
    };
    let (input_at, host_at, leaf_at) = (index(input), index(host), index(leaf));
    let host_ids: BTreeSet<_> = host_at
        .keys()
        .filter(|id| !excluded_object_ids.contains(*id))
        .copied()
        .collect();
    let leaf_ids: BTreeSet<_> = leaf_at
        .keys()
        .filter(|id| !excluded_object_ids.contains(*id))
        .copied()
        .collect();
    if host_ids != leaf_ids {
        diff.violations.push(format!(
            "{member}: object identifiers differ: only in host {:?}, only in leaf {:?}",
            host_ids.difference(&leaf_ids).collect::<Vec<_>>(),
            leaf_ids.difference(&host_ids).collect::<Vec<_>>()
        ));
    }
    let host_order: Vec<_> = host
        .iter()
        .map(|o| o.identifier)
        .filter(|id| leaf_ids.contains(id))
        .collect();
    let leaf_order: Vec<_> = leaf
        .iter()
        .map(|o| o.identifier)
        .filter(|id| host_ids.contains(id))
        .collect();
    if host_order != leaf_order {
        diff.notes.push(format!(
            "{member}: object order differs between host and leaf"
        ));
    }
    for id in host_ids.intersection(&leaf_ids) {
        let (h, l) = (&host[host_at[id]], &leaf[leaf_at[id]]);
        if h.raw == l.raw {
            continue;
        }
        let source = input_at.get(id).map(|&position| &input[position]);
        let host_untouched = source.is_some_and(|source| source.raw == h.raw);
        let leaf_untouched = source.is_some_and(|source| source.raw == l.raw);
        let label = format!("{member} object {id} ({})", types(h));
        // A documented metadata policy also covers routes whose host
        // never rewrites the registry (for example a text-only rewrite):
        // the leaf's advanced save token is normalised below.
        let policy_metadata =
            policy.any() && h.messages.iter().any(|message| message.type_ == 11_006);
        if host_untouched && !policy_metadata {
            diff.violations.push(format!(
                "{label}: the leaf changed an object the host left untouched"
            ));
            continue;
        }
        if leaf_untouched {
            diff.notes.push(format!(
                "{label}: the host re-encoded an object the leaf kept byte-identical"
            ));
        }
        compare_known(
            &label,
            source,
            h,
            l,
            known,
            policy,
            reference_policy,
            diff,
            excluded_object_ids,
            allowed_host_members,
        );
    }
}

const CHART_MESSAGE_TYPE: u32 = 5_021;
const REFERENCE_LINE_NON_STYLES_MAP: u32 = 1;
const REFERENCE_LINE_STYLES_MAP: u32 = 2;
const REFERENCE_LINE_AXIS_ITEMS: u32 = 2;
const REFERENCE_LINE_AXIS_STYLES: u32 = 2;
const REFERENCE_LINE_STYLE_ENTRIES: u32 = 2;
const REFERENCE_LINE_NON_STYLE: u32 = 1;
const REFERENCE_LINE_ITEM_UUID: u32 = 2;
const REFERENCE_LINE_STYLE_REFERENCE: u32 = 2;
const REFERENCE_LINE_IDENTIFIER: u32 = 1;
const REFERENCE_LINE_UUID_UPPER: u32 = 2;

fn compare_archive_references(
    source: Option<&Object>,
    host_archive_info: &ArchiveInfo,
    leaf_archive_info: &ArchiveInfo,
    host: &Object,
    leaf: &Object,
    reference_policy: ReferenceOrderPolicy,
) -> Result<Vec<String>, String> {
    let references = |info: &ArchiveInfo| {
        info.message_infos
            .iter()
            .map(|message| {
                (
                    message.type_,
                    message.object_references.clone(),
                    message.data_references.clone(),
                )
            })
            .collect::<Vec<_>>()
    };
    if host_archive_info.message_infos.len() != leaf_archive_info.message_infos.len() {
        return Err(format!(
            "message types or object/data references differ (host {:?}, leaf {:?})",
            references(host_archive_info),
            references(leaf_archive_info)
        ));
    }
    if host.messages.len() != leaf.messages.len() {
        return Err(format!(
            "message types or object/data references differ (host {:?}, leaf {:?})",
            references(host_archive_info),
            references(leaf_archive_info)
        ));
    }

    let mut notes = Vec::new();
    for (position, ((host_message_info, leaf_message_info), (host_message, leaf_message))) in
        host_archive_info
            .message_infos
            .iter()
            .zip(&leaf_archive_info.message_infos)
            .zip(host.messages.iter().zip(&leaf.messages))
            .enumerate()
    {
        if host_message_info.type_ != leaf_message_info.type_
            || host_message_info.data_references != leaf_message_info.data_references
        {
            return Err(format!(
                "message types or object/data references differ (host {:?}, leaf {:?})",
                references(host_archive_info),
                references(leaf_archive_info)
            ));
        }
        let selected_reference_line_chart = host.identifier == leaf.identifier
            && reference_line_policy_selected(reference_policy, host.identifier);
        if host_message_info.type_ == CHART_MESSAGE_TYPE && selected_reference_line_chart {
            match compare_reference_line_metadata(
                &host_message.data,
                &host_message_info.object_references,
                &leaf_message.data,
                &leaf_message_info.object_references,
            ) {
                Ok(true) => notes.push(
                    "chart reference-line-owned dependency references differ only in order; all unrelated references remain ordered and exact"
                        .to_owned(),
                ),
                Ok(false) => {},
                Err(reason) => {
                    return Err(format!(
                        "chart reference-line metadata references differ: {reason} (host {:?}, leaf {:?})",
                        references(host_archive_info),
                        references(leaf_archive_info)
                    ));
                },
            }
            let Some(source_message) = source
                .and_then(|object| object.info.message_infos.get(position))
                .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
                .and_then(|_| source.and_then(|object| object.messages.get(position)))
            else {
                return Err(
                    "selected chart reference-line message is absent from the source snapshot"
                        .to_owned(),
                );
            };
            if let Err(reason) = compare_reference_line_payloads(
                &source_message.data,
                &host_message.data,
                &leaf_message.data,
            ) {
                return Err(format!(
                    "chart reference-line payload differs outside allocation UUIDs: {reason}"
                ));
            }
        } else if host_message_info.object_references != leaf_message_info.object_references {
            return Err(format!(
                "message types or object/data references differ (host {:?}, leaf {:?})",
                references(host_archive_info),
                references(leaf_archive_info)
            ));
        }
    }
    Ok(notes)
}

fn compare_reference_line_metadata(
    host_payload: &[u8],
    host_references: &[u64],
    leaf_payload: &[u8],
    leaf_references: &[u64],
) -> Result<bool, String> {
    let host_owned = reference_line_object_ids(host_payload)?;
    let leaf_owned = reference_line_object_ids(leaf_payload)?;
    if host_owned != leaf_owned {
        return Err(format!(
            "extension-owned IDs differ (host {host_owned:?}, leaf {leaf_owned:?})"
        ));
    }
    let (host_owned_refs, host_other_refs) =
        split_reference_line_references(host_references, &host_owned, "host")?;
    let (leaf_owned_refs, leaf_other_refs) =
        split_reference_line_references(leaf_references, &leaf_owned, "leaf")?;
    if host_other_refs != leaf_other_refs {
        return Err(format!(
            "unowned reference order or multiplicity differs (host {host_other_refs:?}, leaf {leaf_other_refs:?})"
        ));
    }
    let mut host_owned_sorted = host_owned_refs.clone();
    let mut leaf_owned_sorted = leaf_owned_refs.clone();
    host_owned_sorted.sort_unstable();
    leaf_owned_sorted.sort_unstable();
    if host_owned_sorted != leaf_owned_sorted {
        return Err(format!(
            "owned reference membership or multiplicity differs (host {host_owned_refs:?}, leaf {leaf_owned_refs:?})"
        ));
    }
    Ok(host_owned_refs != leaf_owned_refs)
}

fn compare_reference_line_payloads(source: &[u8], host: &[u8], leaf: &[u8]) -> Result<(), String> {
    let source_uuids = reference_line_item_uuid_payloads(source)?;
    validate_existing_reference_line_item_uuids(&source_uuids, host, "host")?;
    validate_existing_reference_line_item_uuids(&source_uuids, leaf, "leaf")?;
    let host = reference_line_payload_without_item_uuids(host, &source_uuids)?;
    let leaf = reference_line_payload_without_item_uuids(leaf, &source_uuids)?;
    if host == leaf {
        return Ok(());
    }
    Err(format!(
        "first difference at {:?}",
        first_difference(&host, &leaf)
    ))
}

fn reference_line_payload_without_item_uuids(
    payload: &[u8],
    source_uuids: &BTreeMap<u64, Vec<u8>>,
) -> Result<Vec<u8>, String> {
    use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};

    let Some(extension) = codec::reference_lines(payload, Limits::DEFAULT)
        .map_err(|error| format!("chart reference-lines extension: {error}"))?
    else {
        return Ok(payload.to_vec());
    };
    let extension = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        extension,
        &[REFERENCE_LINE_NON_STYLES_MAP, REFERENCE_LINE_AXIS_ITEMS],
        |item| -> Result<Vec<u8>, litchi_iwa_common::Error> {
            let reference = required_length_delimited(
                item,
                REFERENCE_LINE_NON_STYLE,
                "reference-line non-style reference",
            )
            .map_err(litchi_iwa_common::Error::InvalidFormat)?;
            let identifier = required_varint(
                reference,
                REFERENCE_LINE_IDENTIFIER,
                "reference-line non-style identifier",
            )
            .map_err(litchi_iwa_common::Error::InvalidFormat)?;
            if source_uuids.contains_key(&identifier) {
                return Ok(item.to_vec());
            }
            litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
                item,
                &[REFERENCE_LINE_ITEM_UUID],
                |uuid| {
                    let uuid = litchi_iwa_common::wire::patch_varint_field(
                        uuid,
                        REFERENCE_LINE_IDENTIFIER,
                        true,
                        Some(0),
                    )?;
                    litchi_iwa_common::wire::patch_varint_field(
                        &uuid,
                        REFERENCE_LINE_UUID_UPPER,
                        true,
                        Some(0),
                    )
                },
            )
        },
    )
    .map_err(|error| format!("reference-line item UUIDs: {error}"))?;
    codec::rewrite_reference_lines(payload, Some(&extension), Limits::DEFAULT)
        .map(|rewrite| rewrite.into_bytes().into_owned())
        .map_err(|error| format!("chart reference-lines rewrite: {error}"))
}

fn validate_existing_reference_line_item_uuids(
    source_uuids: &BTreeMap<u64, Vec<u8>>,
    payload: &[u8],
    side: &str,
) -> Result<(), String> {
    let output_uuids = reference_line_item_uuid_payloads(payload)?;
    for (identifier, expected) in source_uuids {
        let Some(actual) = output_uuids.get(identifier) else {
            return Err(format!(
                "{side} output is missing existing reference-line item UUID for object {identifier}"
            ));
        };
        if actual != expected {
            return Err(format!(
                "{side} output changed existing reference-line item UUID for object {identifier}"
            ));
        }
    }
    Ok(())
}

fn reference_line_item_uuid_payloads(payload: &[u8]) -> Result<BTreeMap<u64, Vec<u8>>, String> {
    use litchi_iwa_common::wire::repeated_length_delimited_payloads;
    use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};

    let extension = codec::reference_lines(payload, Limits::DEFAULT)
        .map_err(|error| format!("chart reference-lines extension: {error}"))?;
    let Some(extension) = extension else {
        return Ok(BTreeMap::new());
    };
    let mut output = BTreeMap::new();
    for axis in repeated_length_delimited_payloads(extension, REFERENCE_LINE_NON_STYLES_MAP)
        .map_err(|error| format!("reference-line non-style map: {error}"))?
    {
        for item in repeated_length_delimited_payloads(axis, REFERENCE_LINE_AXIS_ITEMS)
            .map_err(|error| format!("reference-line items: {error}"))?
        {
            let reference = required_length_delimited(
                item,
                REFERENCE_LINE_NON_STYLE,
                "reference-line non-style reference",
            )?;
            let uuid = required_length_delimited(
                item,
                REFERENCE_LINE_ITEM_UUID,
                "reference-line non-style UUID",
            )?;
            let identifier = required_varint(
                reference,
                REFERENCE_LINE_IDENTIFIER,
                "reference-line non-style identifier",
            )?;
            required_varint(
                uuid,
                REFERENCE_LINE_IDENTIFIER,
                "reference-line non-style UUID lower half",
            )?;
            required_varint(
                uuid,
                REFERENCE_LINE_UUID_UPPER,
                "reference-line non-style UUID upper half",
            )?;
            if output.insert(identifier, uuid.to_vec()).is_some() {
                return Err(format!(
                    "extension-owned reference-line object ID {identifier} is repeated"
                ));
            }
        }
    }
    Ok(output)
}

fn split_reference_line_references(
    references: &[u64],
    owned: &[u64],
    side: &str,
) -> Result<(Vec<u64>, Vec<u64>), String> {
    for identifier in owned {
        let count = references
            .iter()
            .filter(|candidate| *candidate == identifier)
            .count();
        if count != 1 {
            return Err(format!(
                "{side} metadata references owned line object {identifier} {count} times"
            ));
        }
    }
    let mut owned_references = Vec::new();
    let mut other_references = Vec::new();
    for identifier in references {
        if owned.contains(identifier) {
            owned_references.push(*identifier);
        } else {
            other_references.push(*identifier);
        }
    }
    Ok((owned_references, other_references))
}

fn reference_line_object_ids(payload: &[u8]) -> Result<Vec<u64>, String> {
    use litchi_iwa_common::wire::repeated_length_delimited_payloads;
    use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};

    let extension = codec::reference_lines(payload, Limits::DEFAULT)
        .map_err(|error| format!("chart reference-lines extension: {error}"))?;
    let Some(extension) = extension else {
        return Ok(Vec::new());
    };
    let mut identifiers = Vec::new();
    for axis in repeated_length_delimited_payloads(extension, REFERENCE_LINE_NON_STYLES_MAP)
        .map_err(|error| format!("reference-line non-style map: {error}"))?
    {
        for item in repeated_length_delimited_payloads(axis, REFERENCE_LINE_AXIS_ITEMS)
            .map_err(|error| format!("reference-line items: {error}"))?
        {
            let reference = required_length_delimited(
                item,
                REFERENCE_LINE_NON_STYLE,
                "reference-line non-style reference",
            )?;
            let uuid = required_length_delimited(
                item,
                REFERENCE_LINE_ITEM_UUID,
                "reference-line non-style UUID",
            )?;
            identifiers.push(required_varint(
                reference,
                REFERENCE_LINE_IDENTIFIER,
                "reference-line non-style identifier",
            )?);
            required_varint(
                uuid,
                REFERENCE_LINE_IDENTIFIER,
                "reference-line non-style UUID lower half",
            )?;
            required_varint(
                uuid,
                REFERENCE_LINE_UUID_UPPER,
                "reference-line non-style UUID upper half",
            )?;
        }
    }
    for axis in repeated_length_delimited_payloads(extension, REFERENCE_LINE_STYLES_MAP)
        .map_err(|error| format!("reference-line style map: {error}"))?
    {
        let Some(sparse) =
            optional_length_delimited(axis, REFERENCE_LINE_AXIS_STYLES, "reference-line styles")?
        else {
            continue;
        };
        for entry in repeated_length_delimited_payloads(sparse, REFERENCE_LINE_STYLE_ENTRIES)
            .map_err(|error| format!("reference-line style entries: {error}"))?
        {
            let reference = required_length_delimited(
                entry,
                REFERENCE_LINE_STYLE_REFERENCE,
                "reference-line style reference",
            )?;
            required_varint(
                entry,
                REFERENCE_LINE_IDENTIFIER,
                "reference-line style index",
            )?;
            identifiers.push(required_varint(
                reference,
                REFERENCE_LINE_IDENTIFIER,
                "reference-line style identifier",
            )?);
        }
    }
    let mut seen = BTreeSet::new();
    for identifier in &identifiers {
        if *identifier == 0 || !seen.insert(*identifier) {
            return Err(format!(
                "extension-owned reference-line object ID {identifier} is zero or repeated"
            ));
        }
    }
    Ok(identifiers)
}

fn optional_length_delimited<'a>(
    data: &'a [u8],
    field: u32,
    label: &str,
) -> Result<Option<&'a [u8]>, String> {
    use litchi_iwa_common::wire::repeated_length_delimited_payloads;

    let payloads = repeated_length_delimited_payloads(data, field)
        .map_err(|error| format!("{label}: {error}"))?;
    match payloads.as_slice() {
        [] => Ok(None),
        [payload] => Ok(Some(*payload)),
        _ => Err(format!("{label} field {field} is repeated")),
    }
}

fn required_length_delimited<'a>(
    data: &'a [u8],
    field: u32,
    label: &str,
) -> Result<&'a [u8], String> {
    optional_length_delimited(data, field, label)?
        .ok_or_else(|| format!("{label} field {field} is missing"))
}

fn required_varint(data: &[u8], field: u32, label: &str) -> Result<u64, String> {
    use litchi_iwa_common::wire::repeated_varint_values;

    let values =
        repeated_varint_values(data, field).map_err(|error| format!("{label}: {error}"))?;
    match values.as_slice() {
        [value] => Ok(*value),
        [] => Err(format!("{label} field {field} is missing")),
        _ => Err(format!("{label} field {field} is repeated")),
    }
}

fn types(object: &Object) -> String {
    object
        .messages
        .iter()
        .map(|message| message.type_.to_string())
        .collect::<Vec<_>>()
        .join("+")
}

fn compare_known(
    label: &str,
    source: Option<&Object>,
    host: &Object,
    leaf: &Object,
    known: &KnownFields,
    policy: MetadataPolicy,
    reference_policy: ReferenceOrderPolicy,
    diff: &mut Diff,
    excluded_object_ids: &BTreeSet<u64>,
    allowed_host_members: &[&str],
) {
    if host.messages.len() != leaf.messages.len() {
        diff.violations.push(format!(
            "{label}: payload count differs (host {}, leaf {})",
            host.messages.len(),
            leaf.messages.len()
        ));
        return;
    }
    match compare_archive_references(source, &host.info, &leaf.info, host, leaf, reference_policy) {
        Ok(notes) => diff
            .notes
            .extend(notes.into_iter().map(|note| format!("{label}: {note}"))),
        Err(detail) => diff.violations.push(format!("{label}: {detail}")),
    }
    for (position, (h, l)) in host.messages.iter().zip(&leaf.messages).enumerate() {
        if h.type_ != l.type_ {
            continue; // already reported through the references above
        }
        if h.data == l.data {
            continue;
        }
        let Some((name, decode)) = known.get(h.type_) else {
            diff.violations.push(format!(
                "{label} payload {position}: bytes differ (host {} bytes, leaf {} bytes, first difference at {:?}) and no known-field schema is registered for message type {}; register it in known_fields.rs or with KnownFields::with",
                h.data.len(),
                l.data.len(),
                first_difference(&h.data, &l.data),
                h.type_
            ));
            continue;
        };
        let normalised = if h.type_ == PACKAGE_METADATA && policy.any() {
            let input = source.and_then(|source| {
                source
                    .messages
                    .iter()
                    .find(|message| message.type_ == PACKAGE_METADATA)
            });
            match normalise_metadata(
                input.map(|m| m.data.as_slice()),
                &h.data,
                &l.data,
                policy,
                excluded_object_ids,
                allowed_host_members,
            ) {
                Ok((host_data, data, notes)) => {
                    for note in notes {
                        diff.notes
                            .push(format!("{label} payload {position}: {note}"));
                    }
                    Some((host_data, data))
                },
                Err(violation) => {
                    diff.violations
                        .push(format!("{label} payload {position}: {violation}"));
                    continue;
                },
            }
        } else {
            None
        };
        let host_data = normalised
            .as_ref()
            .map_or(h.data.as_slice(), |(data, _)| data.as_slice());
        let leaf_data = normalised
            .as_ref()
            .map_or(l.data.as_slice(), |(_, data)| data.as_slice());
        if host_data == leaf_data {
            continue;
        }
        match (decode(host_data), decode(leaf_data)) {
            (Ok(hd), Ok(ld)) => {
                if hd.canonical != ld.canonical {
                    diff.violations.push(format!(
                        "{label} payload {position} ({name}): known fields differ\n{}",
                        debug_diff(&hd.debug, &ld.debug)
                    ));
                } else {
                    diff.notes.push(format!(
                        "{label} payload {position} ({name}): equal on known fields, bytes differ (unknown fields host {:?}, leaf {:?})",
                        hd.unknown_fields, ld.unknown_fields
                    ));
                }
            },
            (host_result, leaf_result) => diff.violations.push(format!(
                "{label} payload {position} ({name}): schema decode failed (host {:?}, leaf {:?})",
                host_result.err(),
                leaf_result.err()
            )),
        }
    }
}

const PACKAGE_METADATA: u32 = 11_006;

/// Reference-order rules that a parity case may opt into.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ReferenceOrderPolicy {
    /// Require every message reference list to match in source order.
    #[default]
    Exact,
    /// For one selected chart's message 5021 only, permit the host's
    /// reference-line-owned dependency IDs to be reordered while preserving
    /// every other reference and its order.
    ReferenceLineAllocation { chart_object_id: u64 },
}

fn reference_line_policy_selected(policy: ReferenceOrderPolicy, object_id: u64) -> bool {
    matches!(
        policy,
        ReferenceOrderPolicy::ReferenceLineAllocation { chart_object_id }
            if chart_object_id == object_id
    )
}

/// Documented leaf metadata policies a case accepts. Each one normalises the
/// compared metadata values for exactly the difference it names (after
/// checking the leaf value has the documented shape), so every other field is
/// still compared. A caller may additionally provide a pinned host-only
/// component/member island; those records are removed from the host projection
/// only after the member-to-component closure is validated.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct MetadataPolicy {
    /// The leaf advances the document save token by one and stamps every
    /// component it rewrites with the new token (record 0298 §2.1); the
    /// host keeps the source token, or advances it but leaves a component
    /// it rewrote unstamped (record 0376b).
    pub save_token: bool,
    /// The leaf derives the UUIDs of objects it creates deterministically
    /// (record 0298 §2.1); the host draws random ones. The identifiers must
    /// match; only the UUIDs of objects absent from the input may differ,
    /// and the leaf's must be unique.
    pub fresh_uuids: bool,
}

impl MetadataPolicy {
    fn any(self) -> bool {
        self.save_token || self.fresh_uuids
    }
}

fn normalise_metadata(
    input: Option<&[u8]>,
    host: &[u8],
    leaf: &[u8],
    policy: MetadataPolicy,
    excluded_object_ids: &BTreeSet<u64>,
    allowed_host_members: &[&str],
) -> Result<(Vec<u8>, Vec<u8>, Vec<String>), String> {
    use litchi_iwa_protos::tsp::PackageMetadata;
    use prost::Message as _;
    let decode = |data: &[u8]| PackageMetadata::decode(data).map_err(|e| e.to_string());
    let mut host = decode(host)?;
    let mut leaf = decode(leaf)?;
    let input = input.map(decode).transpose()?;
    let mut notes = Vec::new();

    // A table removal has one documented host-only island.  Its member names
    // and object identifiers are pinned by the case, so this normalisation
    // cannot grow to cover an arbitrary new deletion.  Validate the member ->
    // component relation before removing anything from the metadata graph.
    let mut orphan_component_ids = BTreeSet::new();
    for member in allowed_host_members {
        let locator = member
            .strip_prefix("Index/")
            .and_then(|name| name.strip_suffix(".iwa"))
            .ok_or_else(|| format!("invalid allowed host member {member:?}"))?;
        let matches = host
            .components
            .iter()
            .filter(|component| {
                component.locator.as_deref() == Some(locator)
                    || component.preferred_locator == locator
            })
            .map(|component| component.identifier)
            .collect::<Vec<_>>();
        let [identifier] = matches.as_slice() else {
            return Err(format!(
                "allowed host member {member:?} maps to {} metadata components, expected one",
                matches.len()
            ));
        };
        orphan_component_ids.insert(*identifier);
    }

    fn remove_orphan_links(
        component: &mut litchi_iwa_protos::tsp::ComponentInfo,
        excluded_object_ids: &BTreeSet<u64>,
        orphan_component_ids: &BTreeSet<u64>,
    ) {
        component
            .object_uuid_map_entries
            .retain(|entry| !excluded_object_ids.contains(&entry.identifier));
        component
            .ambiguous_object_identifiers
            .retain(|identifier| !excluded_object_ids.contains(identifier));
        component.external_references.retain(|reference| {
            !orphan_component_ids.contains(&reference.component_identifier)
                && reference
                    .object_identifier
                    .is_none_or(|identifier| !excluded_object_ids.contains(&identifier))
        });
        component.versioned_external_references.retain(|reference| {
            !orphan_component_ids.contains(&reference.component_identifier)
                && reference
                    .object_identifier
                    .is_none_or(|identifier| !excluded_object_ids.contains(&identifier))
        });
        component.data_references.retain_mut(|reference| {
            let before = reference.object_reference_list.len();
            reference
                .object_reference_list
                .retain(|object| !excluded_object_ids.contains(&object.object_identifier));
            before == reference.object_reference_list.len()
                || !reference.object_reference_list.is_empty()
        });
    }

    if !orphan_component_ids.is_empty() {
        host.components
            .retain(|component| !orphan_component_ids.contains(&component.identifier));
        host.versioned_components
            .retain(|component| !orphan_component_ids.contains(&component.identifier));
        notes.push(format!(
            "removed {} host-only metadata component(s) belonging to the pinned orphan island",
            orphan_component_ids.len()
        ));
    }
    for component in &mut host.components {
        remove_orphan_links(component, excluded_object_ids, &orphan_component_ids);
    }
    for component in &mut host.versioned_components {
        remove_orphan_links(component, excluded_object_ids, &orphan_component_ids);
    }
    for component in &mut leaf.components {
        remove_orphan_links(component, excluded_object_ids, &orphan_component_ids);
    }
    for component in &mut leaf.versioned_components {
        remove_orphan_links(component, excluded_object_ids, &orphan_component_ids);
    }

    if policy.save_token && leaf.save_token != host.save_token {
        let source = input.as_ref().and_then(|input| input.save_token);
        let advanced = source.zip(leaf.save_token).is_some_and(|(s, l)| l == s + 1);
        if host.save_token != source || !advanced {
            return Err(format!(
                "save token: input {source:?}, host {:?}, leaf {:?} (the leaf must advance the input token by one)",
                host.save_token, leaf.save_token
            ));
        }
        let token = leaf.save_token;
        let mut stamped = 0;
        for component in &mut leaf.components {
            let Some(theirs) = host
                .components
                .iter()
                .find(|c| c.identifier == component.identifier)
            else {
                continue;
            };
            if component.save_token != theirs.save_token {
                if component.save_token != token {
                    return Err(format!(
                        "component {} save token {:?} is neither the host's {:?} nor the new document token {token:?}",
                        component.identifier, component.save_token, theirs.save_token
                    ));
                }
                component.save_token = theirs.save_token;
                stamped += 1;
            }
        }
        leaf.save_token = host.save_token;
        notes.push(format!(
            "the leaf advanced the save token to {token:?} and stamped {stamped} component(s) (documented policy)"
        ));
    } else if policy.save_token {
        // Both advanced the document token: the leaf also stamps a component
        // it rewrites that the host leaves at its input token (the host's
        // volatile conditional-highlight family rewrites the calculation
        // engine without stamping it, record 0376b).
        let token = leaf.save_token;
        let mut stamped = 0;
        for component in &mut leaf.components {
            let Some(theirs) = host
                .components
                .iter()
                .find(|c| c.identifier == component.identifier)
            else {
                continue;
            };
            if component.save_token == theirs.save_token {
                continue;
            }
            let source = input
                .as_ref()
                .and_then(|input| {
                    input
                        .components
                        .iter()
                        .find(|c| c.identifier == component.identifier)
                })
                .and_then(|c| c.save_token);
            if component.save_token != token || theirs.save_token != source {
                return Err(format!(
                    "component {} save token: input {source:?}, host {:?}, leaf {:?} (the leaf may only stamp the new document token {token:?} on a component the host left unstamped)",
                    component.identifier, theirs.save_token, component.save_token
                ));
            }
            component.save_token = theirs.save_token;
            stamped += 1;
        }
        if stamped > 0 {
            notes.push(format!(
                "the leaf stamped {stamped} component(s) the host rewrote without stamping (documented policy)"
            ));
        }
    }
    if policy.fresh_uuids {
        let known: BTreeSet<u64> = input
            .iter()
            .flat_map(|input| &input.components)
            .flat_map(|c| &c.object_uuid_map_entries)
            .map(|entry| entry.identifier)
            .collect();
        let host_uuid: BTreeMap<u64, _> = host
            .components
            .iter()
            .flat_map(|c| &c.object_uuid_map_entries)
            .map(|entry| (entry.identifier, entry.uuid))
            .collect();
        let mut seen = BTreeSet::new();
        let mut replaced = 0;
        for entry in leaf
            .components
            .iter_mut()
            .flat_map(|c| &mut c.object_uuid_map_entries)
        {
            if !seen.insert((entry.uuid.lower, entry.uuid.upper)) {
                return Err(format!(
                    "the leaf reuses a UUID for object {}",
                    entry.identifier
                ));
            }
            if known.contains(&entry.identifier) {
                continue;
            }
            if let Some(uuid) = host_uuid.get(&entry.identifier)
                && *uuid != entry.uuid
            {
                entry.uuid = *uuid;
                replaced += 1;
            }
        }
        if replaced > 0 {
            notes.push(format!(
                "{replaced} new object UUID(s) differ (leaf deterministic, host random; documented policy)"
            ));
        }
    }
    Ok((host.encode_to_vec(), leaf.encode_to_vec(), notes))
}

/// Offset of the first differing byte, or the shorter length.
pub fn first_difference(left: &[u8], right: &[u8]) -> Option<usize> {
    left.iter()
        .zip(right)
        .position(|(a, b)| a != b)
        .or_else(|| (left.len() != right.len()).then(|| left.len().min(right.len())))
}

/// The differing middle of two multi-line renderings (common prefix and
/// suffix trimmed), at most 30 lines per side.
pub fn debug_diff(host: &str, leaf: &str) -> String {
    let host: Vec<_> = host.lines().collect();
    let leaf: Vec<_> = leaf.lines().collect();
    let prefix = host.iter().zip(&leaf).take_while(|(a, b)| a == b).count();
    let suffix = host[prefix..]
        .iter()
        .rev()
        .zip(leaf[prefix..].iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let mut text = format!("    (from line {})\n", prefix + 1);
    for line in host[prefix..host.len() - suffix].iter().take(30) {
        let _ = writeln!(text, "    - host {line}");
    }
    for line in leaf[prefix..leaf.len() - suffix].iter().take(30) {
        let _ = writeln!(text, "    + leaf {line}");
    }
    text
}

#[cfg(test)]
mod tests {
    use super::*;

    fn varint(mut value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return output;
            }
        }
    }

    fn varint_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = varint(u64::from(field) << 3);
        output.extend(varint(value));
        output
    }

    fn length_delimited(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = varint((u64::from(field) << 3) | 2);
        output.extend(varint(payload.len() as u64));
        output.extend_from_slice(payload);
        output
    }

    fn chart_reference_line_payload(identifiers: &[u64]) -> Vec<u8> {
        chart_reference_line_payload_with_uuid_unknown(identifiers, None)
    }

    fn chart_reference_line_payload_with_uuid_unknown(
        identifiers: &[u64],
        uuid_unknown: Option<(u64, u64)>,
    ) -> Vec<u8> {
        let mut axis_id = varint_field(1, 2);
        axis_id.extend(varint_field(2, 0));
        let mut axis = length_delimited(1, &axis_id);
        for identifier in identifiers {
            let reference = varint_field(1, *identifier);
            let mut uuid = varint_field(1, *identifier);
            uuid.extend(varint_field(2, identifier.saturating_add(100)));
            if let Some((unknown_identifier, unknown_value)) = uuid_unknown
                && unknown_identifier == *identifier
            {
                uuid.extend(varint_field(12_345, unknown_value));
            }
            let mut item = length_delimited(1, &reference);
            item.extend(length_delimited(2, &uuid));
            axis.extend(length_delimited(2, &item));
        }
        let extension = length_delimited(1, &axis);
        let chart = length_delimited(10_005, &extension);
        length_delimited(10_000, &chart)
    }

    #[test]
    fn reference_line_policy_only_permits_owned_dependency_reordering() {
        let payload = chart_reference_line_payload(&[11, 22]);
        let reordered =
            compare_reference_line_metadata(&payload, &[7, 11, 8, 22], &payload, &[7, 22, 8, 11])
                .expect("owned references can reorder");
        assert!(reordered);

        let error =
            compare_reference_line_metadata(&payload, &[7, 11, 8, 22], &payload, &[8, 22, 7, 11])
                .expect_err("unrelated references cannot reorder");
        assert!(error.contains("unowned reference order"), "{error}");
    }

    #[test]
    fn reference_line_policy_is_opt_in_and_chart_bound() {
        assert!(!reference_line_policy_selected(
            ReferenceOrderPolicy::Exact,
            42
        ));
        assert!(!reference_line_policy_selected(
            ReferenceOrderPolicy::ReferenceLineAllocation { chart_object_id: 7 },
            42
        ));
        assert!(reference_line_policy_selected(
            ReferenceOrderPolicy::ReferenceLineAllocation {
                chart_object_id: 42
            },
            42
        ));
    }

    #[test]
    fn reference_line_policy_preserves_unchanged_unknown_payload_bytes() {
        let payload = chart_reference_line_payload(&[11, 22]);
        let mut with_unknown = payload.clone();
        litchi_iwa_common::wire::append_varint_field(&mut with_unknown, 12_345, 1)
            .expect("unknown field is bounded");
        assert!(compare_reference_line_payloads(&payload, &with_unknown, &with_unknown).is_ok());

        let mut changed = payload.clone();
        litchi_iwa_common::wire::append_varint_field(&mut changed, 12_345, 2)
            .expect("unknown field is bounded");
        assert!(compare_reference_line_payloads(&payload, &with_unknown, &changed).is_err());
    }

    #[test]
    fn reference_line_policy_preserves_unknown_fields_in_new_item_uuids() {
        let source = chart_reference_line_payload(&[11]);
        let host = chart_reference_line_payload_with_uuid_unknown(&[11, 22], Some((22, 17)));
        let same = chart_reference_line_payload_with_uuid_unknown(&[11, 22], Some((22, 17)));
        let changed = chart_reference_line_payload_with_uuid_unknown(&[11, 22], Some((22, 18)));
        assert!(compare_reference_line_payloads(&source, &host, &same).is_ok());
        assert!(compare_reference_line_payloads(&source, &host, &changed).is_err());
    }

    #[test]
    fn reference_line_policy_default_and_wrong_chart_reject_reordered_metadata() {
        let payload = chart_reference_line_payload(&[11, 22]);
        let host = synthetic_chart_object(42, &[7, 11, 8, 22], &payload);
        let leaf = synthetic_chart_object(42, &[7, 22, 8, 11], &payload);

        assert!(
            compare_archive_references(
                Some(&host),
                &host.info,
                &leaf.info,
                &host,
                &leaf,
                ReferenceOrderPolicy::Exact,
            )
            .is_err()
        );
        assert!(
            compare_archive_references(
                Some(&host),
                &host.info,
                &leaf.info,
                &host,
                &leaf,
                ReferenceOrderPolicy::ReferenceLineAllocation { chart_object_id: 7 },
            )
            .is_err()
        );
        assert!(
            compare_archive_references(
                Some(&host),
                &host.info,
                &leaf.info,
                &host,
                &leaf,
                ReferenceOrderPolicy::ReferenceLineAllocation {
                    chart_object_id: 42
                },
            )
            .is_ok()
        );
    }

    fn synthetic_chart_object(identifier: u64, references: &[u64], payload: &[u8]) -> Object {
        let mut message_info = litchi_iwa_archive::iwa::MessageInfo::new(
            CHART_MESSAGE_TYPE,
            payload.len(),
            litchi_iwa_archive::iwa::Limits::default(),
        )
        .expect("synthetic message metadata is bounded");
        message_info.object_references = references.to_vec();
        Object {
            identifier,
            info: ArchiveInfo {
                identifier: Some(identifier),
                message_infos: vec![message_info],
                should_merge: None,
            },
            messages: vec![RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: payload.to_vec(),
            }],
            raw: Vec::new(),
        }
    }

    #[test]
    fn reference_line_policy_rejects_missing_duplicate_wrong_and_malformed_refs() {
        let payload = chart_reference_line_payload(&[11, 22]);
        let missing =
            compare_reference_line_metadata(&payload, &[7, 11, 8, 22], &payload, &[7, 11, 8])
                .expect_err("missing owned reference must fail");
        assert!(
            missing.contains("owned line object 22 0 times"),
            "{missing}"
        );

        let duplicate = compare_reference_line_metadata(
            &payload,
            &[7, 11, 8, 22],
            &payload,
            &[7, 11, 11, 8, 22],
        )
        .expect_err("duplicate owned reference must fail");
        assert!(
            duplicate.contains("owned line object 11 2 times"),
            "{duplicate}"
        );

        let wrong_payload = chart_reference_line_payload(&[11, 33]);
        let wrong = compare_reference_line_metadata(
            &payload,
            &[7, 11, 8, 22],
            &wrong_payload,
            &[7, 11, 8, 33],
        )
        .expect_err("wrong extension-owned reference must fail");
        assert!(wrong.contains("extension-owned IDs differ"), "{wrong}");

        let duplicate_payload = chart_reference_line_payload(&[11, 11]);
        let duplicate_payload_error = reference_line_object_ids(&duplicate_payload)
            .expect_err("duplicate extension-owned ID must fail");
        assert!(
            duplicate_payload_error.contains("zero or repeated"),
            "{duplicate_payload_error}"
        );

        let malformed =
            reference_line_object_ids(&[0xff]).expect_err("malformed chart payload must fail");
        assert!(
            malformed.contains("chart reference-lines extension"),
            "{malformed}"
        );
    }
}
