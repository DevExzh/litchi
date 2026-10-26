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
pub fn compare(input: &[u8], host: &[u8], leaf: &[u8], known: &KnownFields) -> TestResult<Diff> {
    let input = members(input)?;
    let host = members(host)?;
    let leaf = members(leaf)?;
    let mut diff = Diff::default();

    let names = |members: &[Member]| members.iter().map(|m| m.name.clone()).collect::<Vec<_>>();
    let (input_names, host_names, leaf_names) = (names(&input), names(&host), names(&leaf));
    if host_names != leaf_names {
        let host_set: BTreeSet<_> = host_names.iter().collect();
        let leaf_set: BTreeSet<_> = leaf_names.iter().collect();
        let only_host: Vec<_> = host_set.difference(&leaf_set).collect();
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
                let changed = count_changed(&source_objects, &host_objects);
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
            &mut diff,
        );
    }
    Ok(diff)
}

fn count_changed(source: &[Object], target: &[Object]) -> usize {
    let source: BTreeMap<_, _> = source.iter().map(|o| (o.identifier, &o.raw)).collect();
    let target: BTreeMap<_, _> = target.iter().map(|o| (o.identifier, &o.raw)).collect();
    source
        .keys()
        .chain(target.keys())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .filter(|id| source.get(id) != target.get(id))
        .count()
}

fn compare_objects(
    member: &str,
    input: &[Object],
    host: &[Object],
    leaf: &[Object],
    known: &KnownFields,
    diff: &mut Diff,
) {
    diff.host_changed_objects += count_changed(input, host);
    diff.leaf_changed_objects += count_changed(input, leaf);
    let index = |objects: &[Object]| -> BTreeMap<u64, usize> {
        objects
            .iter()
            .enumerate()
            .map(|(position, object)| (object.identifier, position))
            .collect()
    };
    let (input_at, host_at, leaf_at) = (index(input), index(host), index(leaf));
    let host_ids: BTreeSet<_> = host_at.keys().copied().collect();
    let leaf_ids: BTreeSet<_> = leaf_at.keys().copied().collect();
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
        if host_untouched {
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
        compare_known(&label, h, l, known, diff);
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

fn compare_known(label: &str, host: &Object, leaf: &Object, known: &KnownFields, diff: &mut Diff) {
    if host.messages.len() != leaf.messages.len() {
        diff.violations.push(format!(
            "{label}: payload count differs (host {}, leaf {})",
            host.messages.len(),
            leaf.messages.len()
        ));
        return;
    }
    let references = |info: &ArchiveInfo| {
        info.message_infos
            .iter()
            .map(|m| {
                (
                    m.type_,
                    m.object_references.clone(),
                    m.data_references.clone(),
                )
            })
            .collect::<Vec<_>>()
    };
    if references(&host.info) != references(&leaf.info) {
        diff.violations.push(format!(
            "{label}: message types or object/data references differ (host {:?}, leaf {:?})",
            references(&host.info),
            references(&leaf.info)
        ));
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
        match (decode(&h.data), decode(&l.data)) {
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
