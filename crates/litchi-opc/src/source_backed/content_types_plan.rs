//! Bounded, source-preserving plans for `[Content_Types].xml` edits.
//!
//! A plan scans the source manifest once and records only the lexical ranges
//! that will be removed plus the exact insertion point and byte count for new
//! overrides.  It does not allocate the final XML or an intermediate
//! insertion fragment.  Callers can therefore admit aggregate package limits
//! using [`ContentTypesPlan::final_len`] before asking [`ContentTypesPlan::materialize`]
//! to allocate one output buffer.

use super::{
    DiagnosticCounter, append_relationship_xml_bytes, escaped_xml_attribute_len,
    map_execution_error, overlay_unavailable, push_xml_escaped,
};
use crate::constants::namespace;
use crate::content_type::{ContentType, ContentTypeMap, validate_content_type};
use crate::error::{OpcError, Result};
use crate::limits::{ReadLimits, ReadResource};
use crate::packuri::PackURI;
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::cmp::Ordering;
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

/// A validated content-type override staged for insertion.
pub(crate) type ContentTypeOverride = (PackURI, String);

/// A source-preserving content-types edit plan.
///
/// The source bytes are borrowed for the lifetime of the plan. Staged
/// overrides are retained in bounded owned storage while the final XML
/// remains unallocated until materialize.
#[derive(Debug)]
pub(crate) struct ContentTypesPlan<'source> {
    source: &'source [u8],
    removals: Vec<Range<usize>>,
    additions: Vec<ContentTypeOverride>,
    insertion_offset: Option<usize>,
    /// Prefix QName bytes from the source Types name, including the colon when
    /// the root is prefixed. Reused for generated Override elements and an
    /// expanded empty-root closing tag.
    root_prefix: Option<Range<usize>>,
    /// Slash and event end for a self-closing Types root. The slash is
    /// replaced lexically by generated children plus a matching close tag.
    empty_root_slash: Option<usize>,
    empty_root_end: Option<usize>,
    final_len: usize,
    mapping_count: usize,
    event_count: usize,
    /// Charges temporary indexes and retained ranges for the complete plan
    /// lifetime, including the materialization phase.
    planning_memory_reservation: Option<Arc<Reservation>>,
}

impl<'source> ContentTypesPlan<'source> {
    /// Package-level entry point for callers that have already performed
    /// their bounded input preflight. Taking ownership here avoids copying
    /// staged selectors a second time before the source scan.
    pub(crate) fn plan_public_owned(
        source: &'source [u8],
        additions: Vec<ContentTypeOverride>,
        removals: &[PackURI],
        default_removals: Vec<String>,
        limits: ReadLimits,
    ) -> Result<Self> {
        Self::plan_with_defaults_owned(
            source,
            additions,
            removals,
            default_removals,
            limits,
            None,
            None,
        )
    }

    pub(crate) fn materialize_public(self, limits: ReadLimits) -> Result<Vec<u8>> {
        self.materialize(limits, None, None).map(|(bytes, _)| bytes)
    }

    /// Scan and validate a source manifest and compute the exact candidate
    /// length and mapping count without constructing candidate XML.
    ///
    /// The caller must have already parsed `source` through
    /// `ContentTypeMap::from_xml`. This scanner repeats the root namespace,
    /// direct-child, event-count, depth, and text/DTD checks needed to make
    /// source ranges safe without allocating a second source map.
    pub(super) fn plan(
        source: &'source [u8],
        additions: &[ContentTypeOverride],
        removals: &[PackURI],
        limits: ReadLimits,
        context: Option<&ExecutionContext>,
        reservation_failures: Option<&DiagnosticCounter>,
    ) -> Result<Self> {
        Self::plan_with_defaults(
            source,
            additions,
            removals,
            &[],
            limits,
            context,
            reservation_failures,
        )
    }

    /// Source-backed topology entry point for staged additions that are
    /// already owned by the caller.  This preserves the single bounded copy
    /// performed by the topology planner while the borrowed compatibility
    /// wrapper remains useful to focused tests and legacy callers.
    pub(super) fn plan_owned(
        source: &'source [u8],
        additions: Vec<ContentTypeOverride>,
        removals: &[PackURI],
        limits: ReadLimits,
        context: Option<&ExecutionContext>,
        reservation_failures: Option<&DiagnosticCounter>,
    ) -> Result<Self> {
        Self::plan_with_defaults_owned(
            source,
            additions,
            removals,
            Vec::new(),
            limits,
            context,
            reservation_failures,
        )
    }

    /// Scan a source manifest with both explicit `<Override>` removals and
    /// `<Default Extension="...">` removals.  The final XML writer is shared
    /// with the ordinary topology planner; this extra selector family only
    /// changes which direct mapping ranges are removed during the scan.
    pub(super) fn plan_with_defaults(
        source: &'source [u8],
        additions: &[ContentTypeOverride],
        removals: &[PackURI],
        default_removals: &[String],
        limits: ReadLimits,
        context: Option<&ExecutionContext>,
        reservation_failures: Option<&DiagnosticCounter>,
    ) -> Result<Self> {
        let mut normalized_defaults = Vec::new();
        normalized_defaults
            .try_reserve_exact(default_removals.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC content-types default selectors",
                source,
            })?;
        for value in default_removals {
            normalized_defaults.push(collapse_xml_token(value)?);
        }
        Self::plan_with_defaults_owned(
            source,
            additions.to_vec(),
            removals,
            normalized_defaults,
            limits,
            context,
            reservation_failures,
        )
    }

    fn plan_with_defaults_owned(
        source: &'source [u8],
        additions: Vec<ContentTypeOverride>,
        removals: &[PackURI],
        default_removals: Vec<String>,
        limits: ReadLimits,
        context: Option<&ExecutionContext>,
        reservation_failures: Option<&DiagnosticCounter>,
    ) -> Result<Self> {
        limits.check(
            ReadResource::ContentTypesBytes,
            source.len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            source.len() as u64,
            limits.max_archive_entry_bytes(),
        )?;
        check_context(context)?;

        if additions.len() > limits.max_content_type_mappings() {
            return Err(OpcError::ReadLimit {
                resource: ReadResource::ContentTypeMappings,
                actual: additions.len() as u64,
                maximum: limits.max_content_type_mappings() as u64,
            });
        }
        let removal_selector_count =
            removals
                .len()
                .checked_add(default_removals.len())
                .ok_or(OpcError::ReadLimit {
                    resource: ReadResource::ContentTypeMappings,
                    actual: u64::MAX,
                    maximum: limits.max_content_type_mappings() as u64,
                })?;
        if removal_selector_count > limits.max_content_type_mappings() {
            return Err(OpcError::ReadLimit {
                resource: ReadResource::ContentTypeMappings,
                actual: removal_selector_count as u64,
                maximum: limits.max_content_type_mappings() as u64,
            });
        }

        // Validate staged MIME values without cloning on the success path.
        // An error owns a copy only so the diagnostic retains the rejected
        // value.
        for (_, content_type) in &additions {
            if let Err(reason) = validate_content_type(content_type) {
                return Err(OpcError::InvalidContentType {
                    value: content_type.clone(),
                    reason,
                });
            }
        }

        let planning_memory_reservation = reserve_memory(
            context,
            planning_memory_bound(additions.len(), removal_selector_count)?,
            reservation_failures,
            "source-backed OPC content-types planning metadata",
        )?;

        // The caller normally supplies already canonicalized additions, but
        // reject equivalent selectors here before the plan records any edit.
        // Sorting indexes avoids allocating folded copies of every PartName.
        let addition_order = sorted_override_indexes(&additions)?;
        if addition_order.windows(2).any(|window| {
            additions[window[0]]
                .0
                .is_equivalent_to(&additions[window[1]].0)
        }) {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-type override selectors contain duplicate part names".into(),
            ));
        }
        let removal_order = sorted_pack_uri_indexes(removals)?;
        if removal_order
            .windows(2)
            .any(|window| removals[window[0]].is_equivalent_to(&removals[window[1]]))
        {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-type removal selectors contain duplicate part names".into(),
            ));
        }
        let default_order = sorted_default_indexes(&default_removals)?;
        if default_order.windows(2).any(|window| {
            default_removals[window[0]].eq_ignore_ascii_case(default_removals[window[1]].as_str())
        }) {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-type default selectors contain duplicate extensions".into(),
            ));
        }
        let mut matched_removals = Vec::new();
        matched_removals
            .try_reserve_exact(removals.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC content-types removal matches",
                source,
            })?;
        matched_removals.resize(removals.len(), false);
        let mut matched_defaults = Vec::new();
        matched_defaults
            .try_reserve_exact(default_removals.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC content-types default matches",
                source,
            })?;
        matched_defaults.resize(default_removals.len(), false);

        let mut reader = NsReader::from_reader(source);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut root_seen = false;
        let mut root_is_empty = false;
        let mut root_close_start = None;
        let mut root_prefix = None;
        let mut empty_root_slash = None;
        let mut empty_root_end = None;
        let mut event_count = 0usize;
        let mut mapping_count = 0usize;
        let mut removals_ranges = Vec::new();
        removals_ranges
            .try_reserve_exact(removal_selector_count)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC content-types removal spans",
                source,
            })?;

        loop {
            check_context(context)?;
            event_count = event_count.checked_add(1).ok_or(OpcError::ReadLimit {
                resource: ReadResource::XmlEvents,
                actual: u64::MAX,
                maximum: limits.max_xml_events() as u64,
            })?;
            limits.check(
                ReadResource::XmlEvents,
                event_count as u64,
                limits.max_xml_events() as u64,
            )?;
            let event_start = usize::try_from(reader.buffer_position())
                .map_err(|_| overlay_unavailable("content-types XML position overflows usize"))?;
            let decoder = reader.decoder();
            let (matched_removal, reached_eof, root_start_event, root_empty_event) = {
                let (resolved_namespace, event) = reader.read_resolved_event()?;
                let in_content_types_namespace = matches!(
                    resolved_namespace,
                    ResolveResult::Bound(Namespace(value))
                        if value == namespace::OPC_CONTENT_TYPES.as_bytes()
                );
                let root_start_event = depth == 0 && matches!(&event, Event::Start(_));
                let root_empty_event = depth == 0 && matches!(&event, Event::Empty(_));
                let mut matched_removal = None;
                match &event {
                    Event::Start(element) if depth == 0 => {
                        if root_seen
                            || element.local_name().as_ref() != b"Types"
                            || !in_content_types_namespace
                        {
                            return Err(invalid_structure(
                                "content-types root must be Types in the OPC content-types namespace",
                            ));
                        }
                        root_seen = true;
                        depth = next_depth(depth, limits)?;
                    },
                    Event::Empty(element) if depth == 0 => {
                        if root_seen
                            || element.local_name().as_ref() != b"Types"
                            || !in_content_types_namespace
                        {
                            return Err(invalid_structure(
                                "content-types root must be Types in the OPC content-types namespace",
                            ));
                        }
                        root_seen = true;
                        root_is_empty = true;
                    },
                    Event::Start(element) if depth == 1 => {
                        ensure_direct_mapping(element, in_content_types_namespace)?;
                        mapping_count = checked_mapping_count(mapping_count, limits)?;
                        if (!removals.is_empty() && element.local_name().as_ref() == b"Override")
                            || (!default_removals.is_empty()
                                && element.local_name().as_ref() == b"Default")
                        {
                            return Err(invalid_structure(
                                "non-empty content-type mappings are unsupported for removal",
                            ));
                        }
                        depth = next_depth(depth, limits)?;
                    },
                    Event::Empty(element) if depth == 1 => {
                        ensure_direct_mapping(element, in_content_types_namespace)?;
                        mapping_count = checked_mapping_count(mapping_count, limits)?;
                        if element.local_name().as_ref() == b"Override" && !removals.is_empty() {
                            let part_name = override_part_name(element, decoder)?;
                            let Some(index) = find_pack_uri(removals, &removal_order, &part_name)
                            else {
                                continue;
                            };
                            if matched_removals[index] {
                                return Err(OpcError::InvalidContentTypesManifest(format!(
                                    "duplicate content-type Override for '{}'",
                                    removals[index]
                                )));
                            }
                            matched_removals[index] = true;
                            matched_removal = Some(index);
                        } else if element.local_name().as_ref() == b"Default"
                            && !default_removals.is_empty()
                        {
                            let extension = default_extension(element, decoder)?;
                            let Some(index) =
                                find_default(&default_removals, &default_order, &extension)
                            else {
                                continue;
                            };
                            if matched_defaults[index] {
                                return Err(OpcError::InvalidContentTypesManifest(format!(
                                    "duplicate content-type Default for extension '{}'",
                                    default_removals[index]
                                )));
                            }
                            matched_defaults[index] = true;
                            matched_removal = Some(removals.len() + index);
                        }
                    },
                    Event::Start(_) | Event::Empty(_) => {
                        return Err(invalid_structure(
                            "nested elements are not permitted in a content-type mapping",
                        ));
                    },
                    Event::End(element) => {
                        if depth == 0 {
                            return Err(invalid_structure(
                                "unmatched closing element in content-types manifest",
                            ));
                        }
                        if depth == 1 {
                            if element.local_name().as_ref() != b"Types" {
                                return Err(invalid_structure(
                                    "content-types root closing element is not Types",
                                ));
                            }
                            root_close_start = Some(event_start);
                        }
                        depth = depth.checked_sub(1).ok_or_else(|| {
                            invalid_structure("content-types XML depth underflows")
                        })?;
                    },
                    Event::Text(text) if !is_xml_whitespace(text.as_ref()) => {
                        return Err(invalid_structure(
                            "non-whitespace text is not permitted in the content types manifest",
                        ));
                    },
                    Event::CData(_) | Event::DocType(_) => {
                        return Err(invalid_structure(
                            "CDATA and DTDs are not permitted in the content types manifest",
                        ));
                    },
                    Event::Eof => break,
                    Event::Text(_)
                    | Event::Comment(_)
                    | Event::Decl(_)
                    | Event::PI(_)
                    | Event::GeneralRef(_) => {},
                }
                (
                    matched_removal,
                    matches!(&event, Event::Eof),
                    root_start_event,
                    root_empty_event,
                )
            };
            if reached_eof {
                break;
            }
            let event_end = usize::try_from(reader.buffer_position())
                .map_err(|_| overlay_unavailable("content-types XML position overflows usize"))?;
            if event_end < event_start || event_end > source.len() {
                return Err(invalid_structure(
                    "content-types XML event range is invalid",
                ));
            }
            if root_start_event || root_empty_event {
                root_prefix = root_prefix_range(source, event_start, event_end)?;
                if root_empty_event {
                    empty_root_slash = Some(self_closing_slash(source, event_start, event_end)?);
                    empty_root_end = Some(event_end);
                }
            }
            if matched_removal.is_some() {
                removals_ranges.push(event_start..event_end);
            }
        }

        if !root_seen || (!root_is_empty && depth != 0) {
            return Err(OpcError::InvalidContentTypesManifest(
                "missing or unclosed Types root".into(),
            ));
        }
        if let Some(index) = matched_removals.iter().position(|matched| !matched) {
            return Err(OpcError::InvalidContentTypesManifest(format!(
                "content-type Override for '{}' was not found lexically",
                removals[index]
            )));
        }
        if let Some(index) = matched_defaults.iter().position(|matched| !matched) {
            return Err(OpcError::InvalidContentTypesManifest(format!(
                "content-type Default for extension '{}' was not found lexically",
                default_removals[index]
            )));
        }

        if !additions.is_empty() {
            if !root_is_empty && root_close_start.is_none() {
                return Err(OpcError::InvalidContentTypesManifest(
                    "missing Types closing tag".into(),
                ));
            }
        }

        removals_ranges.sort_unstable_by_key(|range| range.start);
        let removed_bytes = removals_ranges.iter().try_fold(0usize, |total, range| {
            total.checked_add(range.end.checked_sub(range.start)?)
        });
        let removed_bytes = removed_bytes
            .ok_or_else(|| overlay_unavailable("removed content-types bytes overflow"))?;
        let inserted_len = additions.iter().try_fold(0usize, |total, (part, kind)| {
            let escaped_part = escaped_xml_attribute_len(part.as_str())?;
            let escaped_kind = escaped_xml_attribute_len(kind)?;
            total
                .checked_add(
                    root_prefix
                        .as_ref()
                        .map_or(0, |range| range.end.saturating_sub(range.start))
                        .checked_add(b"<Override PartName=\"\" ContentType=\"\"/>".len())
                        .ok_or_else(|| {
                            overlay_unavailable("content-type insertion length overflows")
                        })?,
                )
                .and_then(|value| value.checked_add(escaped_part))
                .and_then(|value| value.checked_add(escaped_kind))
                .ok_or_else(|| overlay_unavailable("content-type insertion length overflows"))
        })?;
        let final_len = source
            .len()
            .checked_sub(removed_bytes)
            .and_then(|value| value.checked_add(inserted_len))
            .and_then(|value| {
                if root_is_empty && !additions.is_empty() {
                    let prefix_len = root_prefix
                        .as_ref()
                        .map_or(0, |range| range.end.saturating_sub(range.start));
                    value
                        .checked_add(prefix_len)
                        .and_then(|value| value.checked_add(7))
                } else {
                    Some(value)
                }
            })
            .ok_or_else(|| overlay_unavailable("content-type output length overflows"))?;
        limits.check(
            ReadResource::ContentTypesBytes,
            final_len as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            ReadResource::PartBytes,
            final_len as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            final_len as u64,
            limits.max_archive_entry_bytes(),
        )?;
        let mapping_count = mapping_count
            .checked_sub(removal_selector_count)
            .and_then(|value| value.checked_add(additions.len()))
            .ok_or_else(|| overlay_unavailable("content-type mapping count underflows"))?;
        limits.check(
            ReadResource::ContentTypeMappings,
            mapping_count as u64,
            limits.max_content_type_mappings() as u64,
        )?;
        let event_count = event_count
            .checked_sub(removal_selector_count)
            .and_then(|value| value.checked_add(additions.len()))
            .and_then(|value| {
                value.checked_add(usize::from(root_is_empty && !additions.is_empty()))
            })
            .ok_or_else(|| overlay_unavailable("content-type XML event count underflows"))?;
        limits.check(
            ReadResource::XmlEvents,
            event_count as u64,
            limits.max_xml_events() as u64,
        )?;

        let insertion_offset = if additions.is_empty() || root_is_empty {
            None
        } else {
            Some(
                root_close_start
                    .ok_or_else(|| invalid_structure("missing Types closing tag for insertion"))?,
            )
        };
        Ok(Self {
            source,
            removals: removals_ranges,
            additions,
            insertion_offset,
            root_prefix,
            empty_root_slash,
            empty_root_end,
            final_len,
            mapping_count,
            event_count,
            planning_memory_reservation,
        })
    }

    /// Exact final serialized manifest length computed during planning.
    #[must_use]
    pub(crate) const fn final_len(&self) -> usize {
        self.final_len
    }

    /// Exact final number of Default/Override mappings.
    #[must_use]
    pub(crate) const fn mapping_count(&self) -> usize {
        self.mapping_count
    }

    /// Whether this plan retains every source byte and mapping unchanged.
    #[must_use]
    pub(crate) const fn is_noop(&self) -> bool {
        self.additions.is_empty() && self.removals.is_empty()
    }

    /// Exact final quick-xml event count, including `Event::Eof`.
    #[must_use]
    pub(crate) const fn event_count(&self) -> usize {
        self.event_count
    }

    /// Materialize one final output allocation after caller aggregate limits
    /// have been admitted. The reservation covers the output vector and the
    /// typed readback map, including decoded strings and hash tables.
    pub(super) fn materialize(
        self,
        limits: ReadLimits,
        context: Option<&ExecutionContext>,
        reservation_failures: Option<&DiagnosticCounter>,
    ) -> Result<(Vec<u8>, Option<Arc<Reservation>>)> {
        let Self {
            source,
            additions,
            removals,
            insertion_offset,
            root_prefix,
            empty_root_slash,
            empty_root_end,
            final_len,
            mapping_count,
            event_count,
            planning_memory_reservation,
        } = self;
        // Keep planning metadata charged until all ranges have been consumed.
        let _planning_memory_reservation = planning_memory_reservation;

        check_context(context)?;
        limits.check(
            ReadResource::ContentTypesBytes,
            final_len as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            final_len as u64,
            limits.max_archive_entry_bytes(),
        )?;
        limits.check(
            ReadResource::XmlEvents,
            event_count as u64,
            limits.max_xml_events() as u64,
        )?;
        let reservation = reserve_memory(
            context,
            materialization_memory_bound(final_len, mapping_count)?,
            reservation_failures,
            "source-backed OPC content-types materialization",
        )?;

        let mut output = Vec::new();
        output
            .try_reserve_exact(final_len)
            .map_err(|source| OpcError::Allocation {
                resource: "source-backed OPC content-types output",
                source,
            })?;
        let prefix = root_prefix
            .as_ref()
            .map_or(&[][..], |range| &source[range.clone()]);
        if !additions.is_empty() {
            if let Some(slash) = empty_root_slash {
                let root_end = empty_root_end.ok_or_else(|| {
                    overlay_unavailable("self-closing Types root end is unavailable")
                })?;
                if !removals.is_empty() || root_end < slash + 2 {
                    return Err(overlay_unavailable(
                        "self-closing Types edit has an invalid lexical range",
                    ));
                }
                output.extend_from_slice(&source[..slash]);
                output.push(b'>');
                append_overrides(&mut output, &additions, prefix, context)?;
                output.extend_from_slice(b"</");
                output.extend_from_slice(prefix);
                output.extend_from_slice(b"Types>");
                output.extend_from_slice(&source[root_end..]);
            } else {
                let mut cursor = 0usize;
                let mut inserted = false;
                for range in &removals {
                    if let Some(offset) =
                        insertion_offset.filter(|offset| !inserted && *offset <= range.start)
                    {
                        output.extend_from_slice(&source[cursor..offset]);
                        append_overrides(&mut output, &additions, prefix, context)?;
                        cursor = offset;
                        inserted = true;
                    }
                    output.extend_from_slice(&source[cursor..range.start]);
                    cursor = range.end;
                }
                if let Some(offset) = insertion_offset.filter(|_| !inserted) {
                    output.extend_from_slice(&source[cursor..offset]);
                    append_overrides(&mut output, &additions, prefix, context)?;
                    cursor = offset;
                }
                output.extend_from_slice(&source[cursor..]);
            }
        } else {
            let mut cursor = 0usize;
            for range in &removals {
                output.extend_from_slice(&source[cursor..range.start]);
                cursor = range.end;
            }
            output.extend_from_slice(&source[cursor..]);
        }
        if output.len() != final_len {
            return Err(overlay_unavailable(
                "content-types materialization length differs from plan",
            ));
        }
        let parsed = ContentTypeMap::from_xml(&output, limits)?;
        if parsed.mapping_count() != mapping_count {
            return Err(invalid_structure(
                "content-types mapping count differs from the source plan",
            ));
        }
        check_context(context)?;
        Ok((output, reservation))
    }
}

fn append_overrides(
    output: &mut Vec<u8>,
    additions: &[ContentTypeOverride],
    prefix: &[u8],
    context: Option<&ExecutionContext>,
) -> Result<()> {
    for (index, (partname, content_type)) in additions.iter().enumerate() {
        if index & 0x3f == 0 {
            check_context(context)?;
        }
        append_relationship_xml_bytes(output, b"<")?;
        output.extend_from_slice(prefix);
        append_relationship_xml_bytes(output, b"Override PartName=\"")?;
        push_xml_escaped(output, partname.as_str())?;
        append_relationship_xml_bytes(output, b"\" ContentType=\"")?;
        push_xml_escaped(output, content_type)?;
        append_relationship_xml_bytes(output, b"\"/>")?;
    }
    Ok(())
}

fn override_part_name(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<PackURI> {
    let mut part_name = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| OpcError::InvalidContentTypesManifest(error.to_string()))?;
        if attribute.key.as_ref() == b"PartName" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| OpcError::InvalidContentTypesManifest(error.to_string()))?;
            part_name = Some(PackURI::new(value.into_owned()).map_err(OpcError::InvalidPackUri)?);
        }
    }
    part_name.ok_or_else(|| {
        OpcError::InvalidContentTypesManifest(
            "content-type Override is missing PartName".to_string(),
        )
    })
}

fn default_extension(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<String> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| OpcError::InvalidContentTypesManifest(error.to_string()))?;
        if attribute.key.as_ref() == b"Extension" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| OpcError::InvalidContentTypesManifest(error.to_string()))?;
            return collapse_xml_token(value.as_ref());
        }
    }
    Err(OpcError::InvalidContentTypesManifest(
        "content-type Default is missing Extension".to_owned(),
    ))
}

/// Apply the XML Schema `xsd:token` whitespace rule: XML whitespace bytes
/// are collapsed to one U+0020 between non-whitespace characters, while
/// other Unicode whitespace such as NBSP remains lexical content.
pub(crate) fn collapse_xml_token(value: &str) -> Result<String> {
    let mut collapsed = String::new();
    collapsed
        .try_reserve_exact(value.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC XML token normalization",
            source,
        })?;
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, ' ' | '\t' | '\r' | '\n') {
            if !collapsed.is_empty() {
                pending_space = true;
            }
            continue;
        }
        if pending_space {
            collapsed.push(' ');
            pending_space = false;
        }
        collapsed.push(character);
    }
    Ok(collapsed)
}

fn ensure_direct_mapping(
    element: &quick_xml::events::BytesStart<'_>,
    in_content_types_namespace: bool,
) -> Result<()> {
    if !in_content_types_namespace {
        return Err(invalid_structure(
            "content-type mapping is outside the OPC content-types namespace",
        ));
    }
    match element.local_name().as_ref() {
        b"Default" | b"Override" => Ok(()),
        _ => Err(invalid_structure(
            "only Default and Override children are permitted",
        )),
    }
}

fn root_prefix_range(source: &[u8], start: usize, end: usize) -> Result<Option<Range<usize>>> {
    let name_start = start
        .checked_add(1)
        .ok_or_else(|| overlay_unavailable("content-types root name position overflows"))?;
    if name_start >= end || source.get(start) != Some(&b'<') {
        return Err(invalid_structure(
            "content-types root start range is invalid",
        ));
    }
    let mut name_end = name_start;
    while name_end < end && !is_xml_name_delimiter(source[name_end]) {
        name_end += 1;
    }
    let name = &source[name_start..name_end];
    if name == b"Types" {
        return Ok(None);
    }
    let Some(colon) = name.iter().position(|byte| *byte == b':') else {
        return Err(invalid_structure("content-types root QName is not Types"));
    };
    if colon == 0 || colon + 1 >= name.len() || &name[colon + 1..] != b"Types" {
        return Err(invalid_structure("content-types root QName is not Types"));
    }
    Ok(Some(name_start..name_start + colon + 1))
}

fn self_closing_slash(source: &[u8], start: usize, end: usize) -> Result<usize> {
    let mut cursor = end
        .checked_sub(1)
        .ok_or_else(|| invalid_structure("empty content-types root range is invalid"))?;
    if source.get(cursor) != Some(&b'>') {
        return Err(invalid_structure(
            "empty content-types root has no closing bracket",
        ));
    }
    cursor = cursor
        .checked_sub(1)
        .ok_or_else(|| invalid_structure("empty content-types root has no self-close slash"))?;
    if cursor <= start || source[cursor] != b'/' {
        return Err(invalid_structure(
            "Types empty root is missing its self-close slash",
        ));
    }
    Ok(cursor)
}

fn is_xml_name_delimiter(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>')
}

fn checked_mapping_count(current: usize, limits: ReadLimits) -> Result<usize> {
    let next = current.checked_add(1).ok_or(OpcError::ReadLimit {
        resource: ReadResource::ContentTypeMappings,
        actual: u64::MAX,
        maximum: limits.max_content_type_mappings() as u64,
    })?;
    limits.check(
        ReadResource::ContentTypeMappings,
        next as u64,
        limits.max_content_type_mappings() as u64,
    )?;
    Ok(next)
}

fn next_depth(depth: usize, limits: ReadLimits) -> Result<usize> {
    let next = depth.checked_add(1).ok_or(OpcError::ReadLimit {
        resource: ReadResource::XmlDepth,
        actual: u64::MAX,
        maximum: limits.max_xml_depth() as u64,
    })?;
    limits.check(
        ReadResource::XmlDepth,
        next as u64,
        limits.max_xml_depth() as u64,
    )?;
    Ok(next)
}

fn planning_memory_bound(additions: usize, removals: usize) -> Result<u64> {
    let index_items = additions
        .checked_add(removals)
        .and_then(|value| value.checked_mul(size_of::<usize>()))
        .ok_or_else(|| overlay_unavailable("content-types planning index size overflows"))?;
    let matched_bytes = removals
        .checked_mul(size_of::<bool>())
        .ok_or_else(|| overlay_unavailable("content-types planning match size overflows"))?;
    let range_bytes = removals
        .checked_mul(size_of::<Range<usize>>())
        .ok_or_else(|| overlay_unavailable("content-types planning range size overflows"))?;
    let total = index_items
        .checked_add(matched_bytes)
        .and_then(|value| value.checked_add(range_bytes))
        .and_then(|value| value.checked_add(4096))
        .ok_or_else(|| overlay_unavailable("content-types planning memory size overflows"))?;
    u64::try_from(total)
        .map_err(|_| overlay_unavailable("content-types planning memory exceeds u64"))
}

fn materialization_memory_bound(final_len: usize, mapping_count: usize) -> Result<u64> {
    // Charge the output Vec plus a conservative XML-sized readback-map term.
    // The per-entry term covers lower-case hash keys, PackURI/ContentType
    // strings, and hash-table allocator overhead retained by ContentTypeMap.
    let serialized = u64::try_from(final_len)
        .map_err(|_| overlay_unavailable("content-types output size exceeds u64"))?;
    let entry_overhead = u64::try_from(mapping_count)
        .ok()
        .and_then(|count| count.checked_mul(512))
        .ok_or_else(|| overlay_unavailable("content-types map overhead overflows"))?;
    serialized
        .checked_mul(3)
        .and_then(|value| value.checked_add(entry_overhead))
        .and_then(|value| value.checked_add(size_of::<ContentTypeMap>() as u64))
        .and_then(|value| value.checked_add(size_of::<ContentType>() as u64))
        .and_then(|value| value.checked_add(4096))
        .ok_or_else(|| overlay_unavailable("content-types materialization memory overflows"))
}

fn reserve_memory(
    context: Option<&ExecutionContext>,
    amount: u64,
    reservation_failures: Option<&DiagnosticCounter>,
    _resource: &'static str,
) -> Result<Option<Arc<Reservation>>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let reservation = context.reserve(Resource::Memory, amount).map_err(|error| {
        if matches!(error, ExecutionError::ResourceLimit(_)) {
            if let Some(counter) = reservation_failures {
                counter.increment();
            }
        }
        map_execution_error(error)
    })?;
    Ok(Some(Arc::new(reservation)))
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn invalid_structure(message: impl Into<String>) -> OpcError {
    OpcError::InvalidContentTypesManifest(message.into())
}

fn sorted_pack_uri_indexes(values: &[PackURI]) -> Result<Vec<usize>> {
    let mut indexes = Vec::new();
    indexes
        .try_reserve_exact(values.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC content-types selector order",
            source,
        })?;
    indexes.extend(0..values.len());
    indexes.sort_unstable_by(|left, right| {
        cmp_ascii_case_insensitive(values[*left].as_str(), values[*right].as_str())
    });
    Ok(indexes)
}

fn sorted_override_indexes(values: &[ContentTypeOverride]) -> Result<Vec<usize>> {
    let mut indexes = Vec::new();
    indexes
        .try_reserve_exact(values.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC content-types override order",
            source,
        })?;
    indexes.extend(0..values.len());
    indexes.sort_unstable_by(|left, right| {
        cmp_ascii_case_insensitive(values[*left].0.as_str(), values[*right].0.as_str())
    });
    Ok(indexes)
}

fn sorted_default_indexes(values: &[String]) -> Result<Vec<usize>> {
    let mut indexes = Vec::new();
    indexes
        .try_reserve_exact(values.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC content-types default order",
            source,
        })?;
    indexes.extend(0..values.len());
    indexes.sort_unstable_by(|left, right| {
        cmp_ascii_case_insensitive(&values[*left], &values[*right])
    });
    Ok(indexes)
}

fn find_pack_uri(values: &[PackURI], order: &[usize], candidate: &PackURI) -> Option<usize> {
    order
        .binary_search_by(|index| {
            cmp_ascii_case_insensitive(values[*index].as_str(), candidate.as_str())
        })
        .ok()
        .map(|position| order[position])
}

fn find_default(values: &[String], order: &[usize], candidate: &str) -> Option<usize> {
    order
        .binary_search_by(|index| cmp_ascii_case_insensitive(&values[*index], candidate))
        .ok()
        .map(|position| order[position])
}

fn cmp_ascii_case_insensitive(left: &str, right: &str) -> Ordering {
    let ordering = left
        .bytes()
        .zip(right.bytes())
        .map(|(left, right)| left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase()))
        .find(|ordering| *ordering != Ordering::Equal);
    ordering.unwrap_or_else(|| left.len().cmp(&right.len()))
}

fn check_context(context: Option<&ExecutionContext>) -> Result<()> {
    context
        .map(ExecutionContext::check)
        .transpose()
        .map_err(map_execution_error)
        .map(|_| ())
}

#[cfg(test)]
mod tests {
    use super::*;

    const SOURCE: &[u8] = br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<!-- preserve -->
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/drop.xml" ContentType="application/xml"/>
</Types>"#;

    #[test]
    fn plan_reports_exact_lexical_output_before_materialization() {
        let additions = [(
            PackURI::new("/word/new.bin").unwrap(),
            "application/octet-stream".to_string(),
        )];
        let removal = PackURI::new("/word/drop.xml").unwrap();
        let plan = ContentTypesPlan::plan(
            SOURCE,
            &additions,
            &[removal],
            ReadLimits::default(),
            None,
            None,
        )
        .unwrap();
        let expected_fragment =
            br#"<Override PartName="/word/new.bin" ContentType="application/octet-stream"/>"#;
        assert_eq!(
            plan.final_len(),
            SOURCE.len()
                - br#"<Override PartName="/word/drop.xml" ContentType="application/xml"/>"#.len()
                + expected_fragment.len()
        );
        assert_eq!(plan.mapping_count(), 2);
        let (output, _) = plan.materialize(ReadLimits::default(), None, None).unwrap();
        assert!(
            output
                .windows(b"<!-- preserve -->".len())
                .any(|window| { window == b"<!-- preserve -->" })
        );
        assert!(
            !output
                .windows(b"/word/drop.xml".len())
                .any(|window| { window == b"/word/drop.xml" })
        );
        assert!(
            output
                .windows(expected_fragment.len())
                .any(|window| { window == expected_fragment })
        );
        assert_eq!(
            ContentTypeMap::from_xml(&output, ReadLimits::default())
                .unwrap()
                .mapping_count(),
            2
        );
    }

    #[test]
    fn exact_final_byte_cap_is_admitted_before_materialization() {
        let additions = [(
            PackURI::new("/word/new.bin").unwrap(),
            "application/octet-stream".to_string(),
        )];
        let first =
            ContentTypesPlan::plan(SOURCE, &additions, &[], ReadLimits::default(), None, None)
                .unwrap();
        let final_len = first.final_len();
        let exact = ReadLimits::builder()
            .max_content_types_bytes(final_len)
            .unwrap()
            .max_archive_entry_bytes(final_len as u64)
            .unwrap()
            .build()
            .unwrap();
        let exact_plan =
            ContentTypesPlan::plan(SOURCE, &additions, &[], exact, None, None).unwrap();
        let (output, _) = exact_plan.materialize(exact, None, None).unwrap();
        assert_eq!(output.len(), final_len);

        let below = ReadLimits::builder()
            .max_content_types_bytes(final_len - 1)
            .unwrap()
            .max_archive_entry_bytes((final_len - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            ContentTypesPlan::plan(SOURCE, &additions, &[], below, None, None),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ContentTypesBytes,
                ..
            })
        ));
    }

    #[test]
    fn plan_rejects_unknown_direct_children_instead_of_counting_them() {
        let source = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Future/></Types>"#;
        let error = ContentTypesPlan::plan(source, &[], &[], ReadLimits::default(), None, None)
            .unwrap_err();
        assert!(matches!(error, OpcError::InvalidContentTypesManifest(_)));
    }

    #[test]
    fn plan_expands_self_closing_root_before_output_allocation() {
        let source =
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;
        let additions = [(
            PackURI::new("/word/new.bin").unwrap(),
            "application/octet-stream".to_string(),
        )];
        let plan =
            ContentTypesPlan::plan(source, &additions, &[], ReadLimits::default(), None, None)
                .unwrap();
        let (output, _) = plan.materialize(ReadLimits::default(), None, None).unwrap();
        assert_eq!(
            output,
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/new.bin" ContentType="application/octet-stream"/></Types>"#
        );
        assert_eq!(
            ContentTypeMap::from_xml(&output, ReadLimits::default())
                .unwrap()
                .mapping_count(),
            1
        );
    }

    #[test]
    fn empty_root_without_changes_is_an_exact_noop() {
        let source =
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;
        let plan =
            ContentTypesPlan::plan(source, &[], &[], ReadLimits::default(), None, None).unwrap();
        let final_len = plan.final_len();
        let (output, _) = plan.materialize(ReadLimits::default(), None, None).unwrap();
        assert_eq!(output, source);
        assert_eq!(final_len, source.len());
    }

    #[test]
    fn prefixed_root_uses_same_qname_for_inserted_override() {
        let source = br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><!-- keep --></ct:Types>"#;
        let additions = [(
            PackURI::new("/word/new.bin").unwrap(),
            "application/octet-stream".to_string(),
        )];
        let plan =
            ContentTypesPlan::plan(source, &additions, &[], ReadLimits::default(), None, None)
                .unwrap();
        let (output, _) = plan.materialize(ReadLimits::default(), None, None).unwrap();
        assert_eq!(
            output,
            br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><!-- keep --><ct:Override PartName="/word/new.bin" ContentType="application/octet-stream"/></ct:Types>"#
        );
        assert_eq!(
            ContentTypeMap::from_xml(&output, ReadLimits::default())
                .unwrap()
                .mapping_count(),
            1
        );
    }

    #[test]
    fn prefixed_self_closing_root_expands_with_prefixed_close() {
        let source =
            br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;
        let additions = [(
            PackURI::new("/word/new.bin").unwrap(),
            "application/octet-stream".to_string(),
        )];
        let plan =
            ContentTypesPlan::plan(source, &additions, &[], ReadLimits::default(), None, None)
                .unwrap();
        let (output, _) = plan.materialize(ReadLimits::default(), None, None).unwrap();
        assert_eq!(
            output,
            br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><ct:Override PartName="/word/new.bin" ContentType="application/octet-stream"/></ct:Types>"#
        );
        assert_eq!(
            ContentTypeMap::from_xml(&output, ReadLimits::default())
                .unwrap()
                .mapping_count(),
            1
        );
    }
}
