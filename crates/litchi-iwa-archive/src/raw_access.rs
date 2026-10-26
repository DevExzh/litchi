//! Raw physical member access and message-type statistics (record 0363 D5).
//!
//! These functions are the physical inspection vocabulary that the
//! `litchi-iwa` migration host used to own on its `Package` and `Document`
//! types: ordered member names, IWA component names, raw member bytes, a
//! bounded cached parse of one IWA component, and message-type statistics.
//! They operate on the archive-owned [`PackageState`] and on parsed
//! [`Archive`] values only. They publish no generated messages, native object
//! identifiers or format semantic values, and they are re-exported only from
//! [`crate::unstable`], never from a format facade.

use std::collections::HashMap;
use std::sync::Arc;

use litchi_iwa_core::{Archive, Error as CoreError, Limits as ArchiveLimits};
use litchi_iwa_core::{SnappyLimits, SnappyStream};

use crate::package::{
    GetOrInsertError, PackageEntry as Entry, PackageState, ParseError, is_legacy_operation_storage,
};

/// Normalize a caller-supplied member name to the stored form.
///
/// Stored names never start with `/`; one leading separator is accepted and
/// removed so `"/Index/Document.iwa"` and `"Index/Document.iwa"` address the
/// same member.
#[must_use]
pub fn normalize_entry_name(name: &str) -> &str {
    name.strip_prefix('/').unwrap_or(name)
}

/// Enumerate every package member name in preserved order.
pub fn entry_names(state: &PackageState) -> impl Iterator<Item = &str> {
    state.iter().map(Entry::name)
}

/// Enumerate package members that contain IWA object archives.
///
/// A legacy `OperationStorage.iwa` member whose payload is a `bvxn`
/// operation log is retained as a raw member but excluded from this list.
pub fn iwa_entry_names(state: &PackageState) -> impl Iterator<Item = &str> {
    state.iwa_entry_names()
}

/// Borrow one member's stored bytes by (normalized) name.
#[must_use]
pub fn entry<'state>(state: &'state PackageState, name: &str) -> Option<&'state [u8]> {
    let position = state.position(normalize_entry_name(name))?;
    Some(state.get_at(position)?.data())
}

/// Why one IWA component could not be read.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum ComponentReadError {
    /// The member name does not end in `.iwa`.
    #[error("package entry {name} is not an IWA component")]
    NotComponent {
        /// The normalized member name.
        name: String,
    },
    /// No member with this name exists.
    #[error("IWA package entry not found: {name}")]
    MissingEntry {
        /// The normalized member name.
        name: String,
    },
    /// The member is a legacy `bvxn` operation log, not an object archive.
    #[error("package entry {name} is a legacy operation log, not an IWA object archive")]
    OperationLog {
        /// The normalized member name.
        name: String,
    },
    /// Snappy framing or IWA framing refused the member under the limits.
    #[error(transparent)]
    Codec(#[from] CoreError),
}

/// Decompress and parse one IWA component under explicit limits.
///
/// Returns the parsed archive and its cache weight (the decompressed stream
/// length, at least one byte). Callers that want the shared bounded cache
/// use [`cached_component`] instead.
///
/// # Errors
///
/// Returns [`ComponentReadError`] when the member is missing, is not an IWA
/// component, is a legacy operation log, or is refused by the Snappy or IWA
/// limits.
pub fn parse_component(
    state: &PackageState,
    name: &str,
    snappy: SnappyLimits,
    limits: ArchiveLimits,
) -> Result<(Archive, usize), ComponentReadError> {
    let name = normalize_entry_name(name);
    if !name.ends_with(".iwa") {
        return Err(ComponentReadError::NotComponent {
            name: name.to_owned(),
        });
    }
    let compressed = entry(state, name).ok_or_else(|| ComponentReadError::MissingEntry {
        name: name.to_owned(),
    })?;
    if is_legacy_operation_storage(name, compressed) {
        return Err(ComponentReadError::OperationLog {
            name: name.to_owned(),
        });
    }
    let stream = SnappyStream::decompress(compressed, snappy)?;
    let weight = stream.as_bytes().len().max(1);
    let archive = Archive::parse(stream.as_bytes(), limits)?;
    Ok((archive, weight))
}

/// Return one parsed IWA component through the state's bounded cache.
///
/// Repeated reads of one snapshot share decompression and parsing; a parser
/// failure is returned as [`GetOrInsertError::Parse`] wrapping the
/// [`ComponentReadError`] and remains retryable.
///
/// # Errors
///
/// Returns the cache error, whose parse variant downcasts to
/// [`ComponentReadError`].
pub fn cached_component(
    state: &PackageState,
    name: &str,
    snappy: SnappyLimits,
) -> Result<Arc<Archive>, GetOrInsertError> {
    let name = normalize_entry_name(name);
    state.get_or_parse_archive(name, |limits| {
        parse_component(state, name, snappy, limits)
            .map_err(|error| -> ParseError { Box::new(error) })
    })
}

/// Count each payload message type across objects.
///
/// The iterator yields one message type per payload; repeated types are
/// counted once per payload.
#[must_use]
pub fn count_message_types(types: impl IntoIterator<Item = u32>) -> HashMap<u32, usize> {
    let mut counts = HashMap::new();
    for type_ in types {
        *counts.entry(type_).or_insert(0) += 1;
    }
    counts
}

/// Return the most frequent message type and its count.
///
/// Ties resolve to the smallest type so the answer is deterministic.
#[must_use]
pub fn most_common_message_type(counts: &HashMap<u32, usize>) -> Option<(u32, usize)> {
    counts
        .iter()
        .max_by(|(left_type, left_count), (right_type, right_count)| {
            left_count
                .cmp(right_count)
                .then_with(|| right_type.cmp(left_type))
        })
        .map(|(&type_, &count)| (type_, count))
}

/// Summarize the five most frequent message types deterministically.
///
/// Entries are ordered by descending count, then ascending type, formatted as
/// `"type: count"` and joined by `", "`; when more types exist the summary
/// ends with `" (and N more)"`.
#[must_use]
pub fn message_type_summary(counts: &HashMap<u32, usize>) -> String {
    let mut types: Vec<_> = counts.iter().collect();
    types.sort_by(|(left_type, left_count), (right_type, right_count)| {
        right_count
            .cmp(left_count)
            .then_with(|| left_type.cmp(right_type))
    });

    let top_types: Vec<String> = types
        .into_iter()
        .take(5)
        .map(|(type_, count)| format!("{type_}: {count}"))
        .collect();

    if top_types.len() < counts.len() {
        format!(
            "{} (and {} more)",
            top_types.join(", "),
            counts.len() - top_types.len()
        )
    } else {
        top_types.join(", ")
    }
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{ArchiveObject, RawMessage};

    use super::*;

    fn component(objects: &[(u64, &[u32])]) -> Vec<u8> {
        let limits = ArchiveLimits::default();
        let archive = Archive {
            objects: objects
                .iter()
                .map(|(identifier, types)| {
                    ArchiveObject::new(
                        *identifier,
                        types
                            .iter()
                            .map(|&type_| RawMessage {
                                type_,
                                data: vec![1, 2],
                            })
                            .collect(),
                        limits,
                    )
                    .expect("fixture object")
                })
                .collect(),
        };
        let stream = archive.to_bytes(limits).expect("fixture stream");
        SnappyStream::compress(&stream, SnappyLimits::default()).expect("fixture snappy")
    }

    fn state() -> PackageState {
        PackageState::from_entries(
            vec![
                Entry::new("Index/Document.iwa".to_owned(), component(&[(1, &[1, 3])])),
                Entry::new(
                    "Index/OperationStorage.iwa".to_owned(),
                    b"bvxn....".to_vec(),
                ),
                Entry::new("Metadata/Properties.plist".to_owned(), b"plist".to_vec()),
            ],
            ArchiveLimits::default(),
        )
        .expect("fixture state")
    }

    #[test]
    fn names_entries_and_components_follow_the_stored_order() {
        let state = state();
        assert_eq!(
            entry_names(&state).collect::<Vec<_>>(),
            [
                "Index/Document.iwa",
                "Index/OperationStorage.iwa",
                "Metadata/Properties.plist"
            ]
        );
        assert_eq!(
            iwa_entry_names(&state).collect::<Vec<_>>(),
            ["Index/Document.iwa"]
        );
        assert_eq!(
            entry(&state, "/Metadata/Properties.plist"),
            Some(&b"plist"[..])
        );
        assert_eq!(entry(&state, "missing"), None);
    }

    #[test]
    fn component_reads_are_typed_and_cached() {
        let state = state();
        let snappy = SnappyLimits::default();
        let first = cached_component(&state, "/Index/Document.iwa", snappy).expect("component");
        let second = cached_component(&state, "Index/Document.iwa", snappy).expect("component");
        assert!(Arc::ptr_eq(&first, &second));
        assert_eq!(first.objects[0].messages.len(), 2);

        let limits = ArchiveLimits::default();
        assert!(matches!(
            parse_component(&state, "Metadata/Properties.plist", snappy, limits),
            Err(ComponentReadError::NotComponent { .. })
        ));
        assert!(matches!(
            parse_component(&state, "Index/Missing.iwa", snappy, limits),
            Err(ComponentReadError::MissingEntry { .. })
        ));
        assert!(matches!(
            parse_component(&state, "Index/OperationStorage.iwa", snappy, limits),
            Err(ComponentReadError::OperationLog { .. })
        ));
        match cached_component(&state, "Index/OperationStorage.iwa", snappy) {
            Err(GetOrInsertError::Parse(error)) => assert!(matches!(
                error.downcast_ref::<ComponentReadError>(),
                Some(ComponentReadError::OperationLog { .. })
            )),
            result => panic!("expected a typed parse failure: {result:?}"),
        }
    }

    #[test]
    fn message_type_statistics_are_deterministic() {
        let counts = count_message_types([3, 1, 3, 2, 3, 1, 9, 8, 7, 6]);
        assert_eq!(most_common_message_type(&counts), Some((3, 3)));
        assert_eq!(
            message_type_summary(&counts),
            "3: 3, 1: 2, 2: 1, 6: 1, 7: 1 (and 2 more)"
        );
        let tie = count_message_types([5, 4]);
        assert_eq!(most_common_message_type(&tie), Some((4, 1)));
        assert_eq!(message_type_summary(&tie), "4: 1, 5: 1");
        assert_eq!(most_common_message_type(&HashMap::new()), None);
        assert_eq!(message_type_summary(&HashMap::new()), "");
    }
}
