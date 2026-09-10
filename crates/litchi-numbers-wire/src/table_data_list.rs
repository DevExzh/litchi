//! Shared topology and publication coordinator for Numbers table-data lists.
//!
//! Format adapters retain ownership of their generated-free codecs, budgets,
//! and semantic converters. This module coordinates only the wire-independent
//! root/segment selection and the final key/value projection. Message payloads
//! remain borrowed from the adapter's package storage for the entire call.

use std::collections::HashSet;

/// Legacy root `TableDataList` message type.
pub const TABLE_DATA_LIST_MESSAGE_KIND: u32 = 6_005;
/// Native root `TableDataList` compatibility message type.
pub const NATIVE_TABLE_DATA_LIST_MESSAGE_KIND: u32 = 6_201;
/// Referenced `TableDataListSegment` message type.
pub const TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND: u32 = 6_011;

/// One borrowed message candidate supplied by a format adapter.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Message<'source> {
    /// Native message type used for topology routing.
    pub kind: u32,
    /// Message payload borrowed from the adapter's immutable package storage.
    pub data: &'source [u8],
}

impl<'source> Message<'source> {
    /// Construct one borrowed routing candidate.
    #[must_use]
    pub const fn new(kind: u32, data: &'source [u8]) -> Self {
        Self { kind, data }
    }
}

/// Which list envelope a decoder should parse.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RootOrSegment {
    /// A root `TableDataList` archive.
    Root,
    /// A referenced `TableDataListSegment` archive.
    Segment,
}

/// A segment's source-declared key range.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct KeyRange {
    /// First key admitted by the segment range.
    pub location: u32,
    /// Number of keys covered by the segment range.
    pub length: u32,
}

/// Minimum and maximum keys observed while decoding one candidate.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EntryBounds {
    /// Smallest entry key, when the candidate contains entries.
    pub minimum: u32,
    /// Largest entry key, when the candidate contains entries.
    pub maximum: u32,
}

/// A fully decoded list candidate staged by an adapter codec.
///
/// The vectors and key set are owned by the candidate so the coordinator can
/// publish them only after all selected roots and referenced segments have
/// been checked. `structural_error` takes precedence over `semantic_error` at
/// publication. A decoder should leave staged prefixes in the candidate when
/// it records a deferred visitor error; the coordinator will still walk the
/// remaining topology before returning that error.
#[derive(Debug)]
pub struct Candidate<T, E> {
    /// List type observed in the decoded envelope.
    pub list_type: i32,
    /// Converted values that can be published after topology validation.
    /// Each key must belong to `keys`; duplicate candidate keys must be
    /// reported through `structural_error` by the decoder.
    pub values: Vec<(u32, T)>,
    /// All keys visited by the candidate, including keys whose conversion was
    /// deferred or failed and therefore has no value in `values`.
    pub keys: HashSet<u32>,
    /// Referenced segment object identifiers from a root candidate. The
    /// decoder owns the bounded identity set and must keep this vector unique;
    /// a repeated reference is reported through `structural_error` before the
    /// candidate is returned.
    pub segment_refs: Vec<u64>,
    /// Source key range from a segment candidate.
    pub key_range: Option<KeyRange>,
    /// Bounds over all entry keys observed by the candidate.
    pub entry_bounds: Option<EntryBounds>,
    /// First deferred structural failure observed by the adapter visitor.
    pub structural_error: Option<E>,
    /// First deferred semantic or allocation failure observed by the adapter
    /// visitor.
    pub semantic_error: Option<E>,
}

/// Allocation target reported by the shared publication pass.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AllocationTarget {
    /// The merged key set.
    Keys,
    /// The merged value vector.
    Values,
}

/// A topology, shape, or bounded-publication issue independent of an adapter's
/// concrete error type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CoordinatorIssue {
    /// No root candidate with the requested list type was found.
    MissingRoot {
        /// Object containing the root candidates.
        object_id: u64,
        /// Requested list type.
        expected_type: i32,
    },
    /// More than one matching root candidate was found.
    DuplicateRoot {
        /// Object containing the duplicate roots.
        object_id: u64,
        /// Requested list type.
        expected_type: i32,
    },
    /// A root-referenced segment object could not be resolved.
    MissingSegment {
        /// Root object containing the reference.
        object_id: u64,
        /// Missing segment identifier.
        segment_id: u64,
    },
    /// A segment object contains no segment archive candidate.
    MissingSegmentPayload {
        /// Segment object inspected by the coordinator.
        segment_id: u64,
    },
    /// A segment object contains more than one segment archive candidate.
    DuplicateSegmentPayload {
        /// Segment object containing duplicate archives.
        segment_id: u64,
    },
    /// A segment archive has a different list type than the selected root.
    WrongSegmentType {
        /// Segment object with the wrong type.
        segment_id: u64,
        /// Requested root list type.
        expected_type: i32,
        /// Type decoded from the segment archive.
        actual_type: i32,
    },
    /// A segment range's location plus length overflowed `u32`.
    KeyRangeOverflow {
        /// Segment object with the invalid range.
        segment_id: u64,
    },
    /// A decoded segment did not provide its source key range.
    MissingKeyRange {
        /// Segment object with the incomplete archive.
        segment_id: u64,
    },
    /// A decoded segment entry lies outside its declared key range.
    EntryOutsideKeyRange {
        /// Segment object with the invalid entry.
        segment_id: u64,
    },
    /// A value key was already admitted by the root or an earlier segment.
    DuplicateEntryKey {
        /// Repeated key.
        key: u32,
    },
    /// The selected projection exceeded its entry ceiling.
    EntryLimit { observed: usize, maximum: usize },
    /// A final merge allocation could not be reserved.
    Allocation {
        /// Collection that could not be reserved.
        target: AllocationTarget,
        /// Requested element count.
        amount: usize,
    },
}

/// Controls bounded value publication after candidate decoding.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OverflowPolicy {
    /// Return the mapped issue as soon as the configured limit or range check
    /// overflows.
    Immediate,
    /// Retain the first issue, continue structural traversal, and return it
    /// after structural errors have had precedence.
    Deferred,
}

/// Entry ceiling and overflow semantics for one coordinated list read.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ListReadPolicy {
    /// Maximum number of admitted keys in the merged projection.
    pub max_entries: usize,
    /// Whether an overflow returns immediately or is deferred.
    pub overflow: OverflowPolicy,
    /// Whether a segment key-range arithmetic overflow returns immediately or
    /// is deferred until the coordinator has finished walking the topology.
    /// Host-side compatibility keeps the historical immediate failure, while
    /// the focused reader uses the deferred default.
    pub range_overflow: OverflowPolicy,
}

impl Default for ListReadPolicy {
    fn default() -> Self {
        Self {
            max_entries: usize::MAX,
            overflow: OverflowPolicy::Deferred,
            range_overflow: OverflowPolicy::Deferred,
        }
    }
}

/// Adapter hooks for strict list envelope decoding and native error mapping.
pub trait ListDecoder<'source> {
    /// Semantic value produced by an admitted entry.
    type Value;
    /// Adapter-owned error type used for codec, conversion, and coordinator
    /// failures.
    type Error;

    /// Probe the list type without staging semantic values.
    fn probe(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        object_id: u64,
    ) -> Result<i32, Self::Error>;

    /// Decode and strictly validate one root or segment candidate.
    ///
    /// `admit` is false for wrong-type and duplicate candidates. Such calls
    /// must still validate their complete wire shape while avoiding semantic
    /// sidecar resolution and retained-value allocation.
    fn decode(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        object_id: u64,
        admit: bool,
    ) -> Result<Candidate<Self::Value, Self::Error>, Self::Error>;

    /// Map a neutral coordinator issue into the adapter's error vocabulary.
    fn map_issue(&mut self, issue: CoordinatorIssue) -> Self::Error;
}

/// Coordinate root/segment selection, topology validation, and sorted value
/// publication for one table-data-list object.
///
/// The root iterator and every resolver iterator must yield [`Message`] values
/// borrowing from the same source lifetime. The coordinator never stores raw
/// message objects or native IDs after the call; only converted `(key, value)`
/// pairs are returned.
pub fn read_list<'source, Root, Segments, Resolve, Decoder>(
    object_id: u64,
    expected_type: i32,
    root_messages: Root,
    mut resolve_segment: Resolve,
    decoder: &mut Decoder,
    policy: ListReadPolicy,
) -> Result<Vec<(u32, Decoder::Value)>, Decoder::Error>
where
    Root: IntoIterator<Item = Message<'source>>,
    Segments: IntoIterator<Item = Message<'source>>,
    Resolve: FnMut(u64) -> Result<Option<Segments>, Decoder::Error>,
    Decoder: ListDecoder<'source>,
{
    let mut selected = None;
    let mut structural_error = None;
    let mut semantic_error = None;

    for message in root_messages {
        if !matches!(
            message.kind,
            TABLE_DATA_LIST_MESSAGE_KIND | NATIVE_TABLE_DATA_LIST_MESSAGE_KIND
        ) {
            continue;
        }
        let probed_type = decoder.probe(message.data, RootOrSegment::Root, object_id)?;
        let duplicate = selected.is_some();
        let admit = !duplicate && probed_type == expected_type;
        let candidate = decoder.decode(message.data, RootOrSegment::Root, object_id, admit)?;
        if candidate.list_type != expected_type {
            continue;
        }
        if duplicate {
            record_issue(
                decoder,
                &mut structural_error,
                CoordinatorIssue::DuplicateRoot {
                    object_id,
                    expected_type,
                },
            );
            continue;
        }
        let Candidate {
            list_type: _,
            values,
            keys,
            segment_refs,
            key_range: _,
            entry_bounds: _,
            structural_error: candidate_structural,
            semantic_error: candidate_semantic,
        } = candidate;
        if let Some(error) = candidate_structural {
            record_error(&mut structural_error, error);
        }
        if let Some(error) = candidate_semantic {
            record_error(&mut semantic_error, error);
        }
        if keys.len() > policy.max_entries {
            let issue = CoordinatorIssue::EntryLimit {
                observed: keys.len(),
                maximum: policy.max_entries,
            };
            if matches!(policy.overflow, OverflowPolicy::Immediate) {
                return Err(decoder.map_issue(issue));
            }
            record_issue(decoder, &mut semantic_error, issue);
        }
        selected = Some((values, keys, segment_refs));
    }

    let Some((mut values, mut keys, mut segment_refs)) = selected else {
        return Err(decoder.map_issue(CoordinatorIssue::MissingRoot {
            object_id,
            expected_type,
        }));
    };

    for segment_id in segment_refs.drain(..) {
        let Some(segment_messages) = (match resolve_segment(segment_id) {
            Ok(messages) => messages,
            Err(error) => {
                record_error(&mut structural_error, error);
                continue;
            },
        }) else {
            record_issue(
                decoder,
                &mut structural_error,
                CoordinatorIssue::MissingSegment {
                    object_id,
                    segment_id,
                },
            );
            continue;
        };

        let mut segment_count = 0usize;
        for message in segment_messages {
            if message.kind != TABLE_DATA_LIST_SEGMENT_MESSAGE_KIND {
                continue;
            }
            segment_count = segment_count.saturating_add(1);
            let probed_type = decoder.probe(message.data, RootOrSegment::Segment, segment_id)?;
            let admit = segment_count == 1 && probed_type == expected_type;
            let candidate =
                decoder.decode(message.data, RootOrSegment::Segment, segment_id, admit)?;
            if segment_count > 1 {
                record_issue(
                    decoder,
                    &mut structural_error,
                    CoordinatorIssue::DuplicateSegmentPayload { segment_id },
                );
                continue;
            }
            if candidate.list_type != expected_type {
                record_issue(
                    decoder,
                    &mut structural_error,
                    CoordinatorIssue::WrongSegmentType {
                        segment_id,
                        expected_type,
                        actual_type: candidate.list_type,
                    },
                );
                continue;
            }
            let Candidate {
                list_type: _,
                values: segment_values,
                keys: _,
                segment_refs: _,
                key_range,
                entry_bounds,
                structural_error: candidate_structural,
                semantic_error: candidate_semantic,
            } = candidate;
            if let Some(error) = candidate_structural {
                record_error(&mut structural_error, error);
            }
            if let Some(error) = candidate_semantic {
                record_error(&mut semantic_error, error);
            }

            let Some(key_range) = key_range else {
                record_issue(
                    decoder,
                    &mut structural_error,
                    CoordinatorIssue::MissingKeyRange { segment_id },
                );
                continue;
            };
            let Some(range_end) = key_range.location.checked_add(key_range.length) else {
                if matches!(policy.range_overflow, OverflowPolicy::Immediate) {
                    return Err(
                        decoder.map_issue(CoordinatorIssue::KeyRangeOverflow { segment_id })
                    );
                }
                record_issue(
                    decoder,
                    &mut structural_error,
                    CoordinatorIssue::KeyRangeOverflow { segment_id },
                );
                continue;
            };
            if let Some(bounds) = entry_bounds
                && (bounds.minimum < key_range.location || bounds.maximum >= range_end)
            {
                record_issue(
                    decoder,
                    &mut structural_error,
                    CoordinatorIssue::EntryOutsideKeyRange { segment_id },
                );
                continue;
            }

            for (key, value) in segment_values {
                if keys.contains(&key) {
                    record_issue(
                        decoder,
                        &mut structural_error,
                        CoordinatorIssue::DuplicateEntryKey { key },
                    );
                    continue;
                }
                if keys.len() >= policy.max_entries {
                    let issue = CoordinatorIssue::EntryLimit {
                        observed: keys.len().saturating_add(1),
                        maximum: policy.max_entries,
                    };
                    if matches!(policy.overflow, OverflowPolicy::Immediate) {
                        return Err(decoder.map_issue(issue));
                    }
                    record_issue(decoder, &mut semantic_error, issue);
                    continue;
                }
                if keys.try_reserve(1).is_err() {
                    record_issue(
                        decoder,
                        &mut semantic_error,
                        CoordinatorIssue::Allocation {
                            target: AllocationTarget::Keys,
                            amount: keys.len().saturating_add(1),
                        },
                    );
                    continue;
                }
                keys.insert(key);
                if values.try_reserve(1).is_err() {
                    record_issue(
                        decoder,
                        &mut semantic_error,
                        CoordinatorIssue::Allocation {
                            target: AllocationTarget::Values,
                            amount: values.len().saturating_add(1),
                        },
                    );
                    continue;
                }
                values.push((key, value));
            }
        }
        if segment_count == 0 {
            record_issue(
                decoder,
                &mut structural_error,
                CoordinatorIssue::MissingSegmentPayload { segment_id },
            );
        }
    }

    if let Some(error) = structural_error {
        return Err(error);
    }
    if let Some(error) = semantic_error {
        return Err(error);
    }
    values.sort_unstable_by_key(|(key, _)| *key);
    for pair in values.windows(2) {
        if pair[0].0 == pair[1].0 {
            return Err(decoder.map_issue(CoordinatorIssue::DuplicateEntryKey { key: pair[0].0 }));
        }
    }
    Ok(values)
}

fn record_error<E>(slot: &mut Option<E>, error: E) {
    if slot.is_none() {
        *slot = Some(error);
    }
}

fn record_issue<'source, Decoder>(
    decoder: &mut Decoder,
    slot: &mut Option<Decoder::Error>,
    issue: CoordinatorIssue,
) where
    Decoder: ListDecoder<'source>,
{
    if slot.is_none() {
        *slot = Some(decoder.map_issue(issue));
    }
}
