//! Fully validated ingress for immutable, already-decoded package entries.

use std::fmt;

use litchi_iwa_package::FrozenEntryStore;

use crate::zip::is_iwa_name;
use crate::{ComponentCatalog, Error, LimitKind, Limits, Result};

/// One immutable logical-entry snapshot and the components decoded from it.
///
/// This type does not claim a physical ZIP representation or exact-save
/// provenance. Construction consumes a frozen copy-on-write entry store,
/// validates every entry before decoding any IWA component, and binds the
/// resulting component catalog to the same immutable state. The complete
/// logical payload is governed by the expanded-byte ceiling; the physical
/// input-byte ceiling is deliberately not reinterpreted as a logical payload
/// ceiling.
pub struct LogicalSourceCatalog {
    entries: FrozenEntryStore,
    components: ComponentCatalog,
    limits: Limits,
}

impl fmt::Debug for LogicalSourceCatalog {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("LogicalSourceCatalog")
            .field("entries", &self.entries.len())
            .field("components", &self.components.len())
            .field("limits", &self.limits)
            .finish()
    }
}

impl LogicalSourceCatalog {
    /// Admit frozen logical entries under the default physical profile.
    ///
    /// # Errors
    ///
    /// Returns an error when an entry name or payload is unsafe, ambiguous,
    /// encrypted, malformed, or exceeds a configured resource ceiling.
    pub fn from_frozen_entries(entries: FrozenEntryStore) -> Result<Self> {
        Self::from_frozen_entries_with_limits(entries, Limits::default())
    }

    /// Admit frozen logical entries under explicit physical limits.
    ///
    /// Names must already be exact, portable slash-separated member names.
    /// A nested `Index.zip` is rejected because this route accepts normalized
    /// logical entries rather than an unexpanded physical container.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_frozen_entries`].
    pub fn from_frozen_entries_with_limits(
        entries: FrozenEntryStore,
        limits: Limits,
    ) -> Result<Self> {
        let checked_limits = limits.validate()?;
        let component_capacity = admit_entries(&entries, checked_limits)?;
        let components = ComponentCatalog::from_validated_logical_entries(
            entries.iter().map(|entry| (entry.name(), entry.data())),
            component_capacity,
            checked_limits,
        )?;
        Ok(Self {
            entries,
            components,
            limits: checked_limits,
        })
    }

    /// Return the retained physical profile used for logical admission and
    /// neutral IWA decoding.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Borrow the immutable logical entries that authorized this snapshot.
    #[must_use]
    pub const fn entries(&self) -> &FrozenEntryStore {
        &self.entries
    }

    /// Borrow parsed components in deterministic normalized-name order.
    #[must_use]
    pub const fn components(&self) -> &ComponentCatalog {
        &self.components
    }

    /// Consume the bound snapshot without cloning entries or components.
    #[must_use]
    pub fn into_parts(self) -> (FrozenEntryStore, ComponentCatalog) {
        (self.entries, self.components)
    }

    /// Consume the snapshot and release all retained logical entry payloads.
    #[must_use]
    pub fn into_components(self) -> ComponentCatalog {
        self.components
    }
}

fn admit_entries(entries: &FrozenEntryStore, limits: Limits) -> Result<usize> {
    check_entry_count(entries.len(), limits)?;
    let zip_limits = limits.zip_limits();
    let mut metadata_bytes = 0_u64;
    let mut total_bytes = 0_u64;
    let mut component_capacity = 0_usize;

    for entry in entries.iter() {
        validate_exact_name(entry.name())?;
        let name_bytes = usize_u64(entry.name().len());
        check_limit(
            LimitKind::MemberNameBytes,
            name_bytes,
            zip_limits.max_member_name_bytes,
        )?;
        metadata_bytes = metadata_bytes.checked_add(name_bytes).ok_or_else(|| {
            Error::InvalidBundle("logical entry metadata byte count overflowed u64".to_owned())
        })?;
        check_limit(
            LimitKind::MetadataBytes,
            metadata_bytes,
            zip_limits.max_metadata_bytes,
        )?;

        let data_bytes = usize_u64(entry.data().len());
        check_limit(LimitKind::EntryBytes, data_bytes, limits.max_entry_bytes())?;
        total_bytes = total_bytes.checked_add(data_bytes).ok_or_else(|| {
            Error::InvalidBundle("logical entry payload byte count overflowed u64".to_owned())
        })?;
        check_limit(LimitKind::TotalBytes, total_bytes, limits.max_total_bytes())?;

        let basename = entry.name().rsplit('/').next().unwrap_or(entry.name());
        if matches!(basename, ".iwpv2" | ".iwph") {
            return Err(Error::Encrypted);
        }
        if basename == "Index.zip" {
            return Err(Error::InvalidBundle(
                "logical entry snapshot contains an unexpanded Index.zip".to_owned(),
            ));
        }
        if is_iwa_name(entry.name()) {
            component_capacity = component_capacity.checked_add(1).ok_or_else(|| {
                Error::InvalidBundle("logical IWA component count overflowed usize".to_owned())
            })?;
        }
    }
    Ok(component_capacity)
}

fn validate_exact_name(name: &str) -> Result<()> {
    let invalid = name.is_empty()
        || name.starts_with('/')
        || name.ends_with('/')
        || name.contains(['\0', '\\'])
        || name.chars().any(char::is_control)
        || name.split('/').any(|component| {
            component.is_empty() || matches!(component, "." | "..") || component.contains(':')
        });
    if invalid {
        return Err(Error::InvalidBundle(format!(
            "logical entry name is not an exact portable member name: {name:?}"
        )));
    }
    Ok(())
}

fn check_entry_count(observed: usize, limits: Limits) -> Result<()> {
    check_limit(
        LimitKind::Entries,
        usize_u64(observed),
        usize_u64(limits.max_entries()),
    )
}

fn check_limit(kind: LimitKind, observed: u64, maximum: u64) -> Result<()> {
    if observed > maximum {
        return Err(Error::Limit {
            kind,
            observed,
            maximum,
        });
    }
    Ok(())
}

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::shadow_unrelated,
    reason = "logical ingress tests use fixed fallible fixtures and compare independent failures"
)]
mod tests {
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_package::{Entry, EntryStore};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    fn iwa(identifier: u64, message_type: u32) -> Vec<u8> {
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    identifier,
                    vec![RawMessage {
                        type_: message_type,
                        data: Vec::new(),
                    }],
                )
                .expect("valid test archive object"),
            ],
        };
        SnappyStream::compress(&archive.to_bytes().expect("encode test archive"))
            .expect("compress test archive")
    }

    fn frozen(entries: Vec<(&str, Vec<u8>)>) -> FrozenEntryStore {
        EntryStore::try_from_entries(
            entries
                .into_iter()
                .map(|(name, data)| Entry::new(name.to_owned(), data))
                .collect(),
        )
        .expect("unique test entries")
        .freeze()
    }

    fn limits(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_total_bytes: u64,
    ) -> Limits {
        Limits::new(
            max_input_bytes,
            max_entries,
            max_entry_bytes,
            max_total_bytes,
            1024 * 1024,
        )
        .expect("valid test limits")
    }

    #[test]
    fn logical_components_match_direct_zip_and_skip_operation_storage() {
        let document = iwa(1, 6_000);
        let calculation = iwa(2, 6_001);
        let operation = b"bvxn operation log".to_vec();
        let entries = frozen(vec![
            ("Index/CalculationEngine.iwa", calculation.clone()),
            ("Index/OperationStorage.iwa", operation.clone()),
            ("Index/Document.iwa", document.clone()),
            ("Data/image.png", vec![1, 2, 3]),
        ]);
        let logical =
            LogicalSourceCatalog::from_frozen_entries(entries).expect("valid logical snapshot");

        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("Index/CalculationEngine.iwa", &calculation)
            .expect("write calculation component");
        writer
            .write_stored("Index/OperationStorage.iwa", &operation)
            .expect("write operation storage");
        writer
            .write_stored("Index/Document.iwa", &document)
            .expect("write document component");
        writer
            .write_stored("Data/image.png", &[1, 2, 3])
            .expect("write sidecar");
        let direct =
            ComponentCatalog::from_bytes(&writer.finish_to_bytes().expect("finish test ZIP"))
                .expect("parse direct ZIP");

        let logical_names = logical
            .components()
            .iter()
            .map(crate::Component::name)
            .collect::<Vec<_>>();
        let direct_names = direct
            .iter()
            .map(crate::Component::name)
            .collect::<Vec<_>>();
        assert_eq!(logical_names, direct_names);
        assert_eq!(logical_names.len(), 2);
    }

    #[test]
    fn frozen_snapshot_is_cow_isolated_without_payload_copy() {
        let document = iwa(1, 6_000);
        let mut mutable = EntryStore::try_from_entries(vec![Entry::new(
            "Index/Document.iwa".to_owned(),
            document,
        )])
        .expect("valid entry store");
        let snapshot = mutable.snapshot();
        let source_pointer = snapshot
            .get("Index/Document.iwa")
            .expect("frozen document")
            .data()
            .as_ptr();
        mutable
            .replace_data(0, b"mutated after freeze".to_vec())
            .expect("replace mutable entry");

        let logical = LogicalSourceCatalog::from_frozen_entries(snapshot)
            .expect("frozen source remains valid");
        let retained_pointer = logical
            .entries()
            .get("Index/Document.iwa")
            .expect("retained document")
            .data()
            .as_ptr();
        assert_eq!(source_pointer, retained_pointer);
        assert_eq!(logical.components().len(), 1);
    }

    #[test]
    fn validates_exact_and_one_over_entry_payload_and_metadata_limits() {
        let one = frozen(vec![("Index/sidecar", vec![1, 2, 3, 4])]);
        LogicalSourceCatalog::from_frozen_entries_with_limits(one.clone(), limits(1024, 1, 4, 4))
            .expect("exact logical limits");

        let error = LogicalSourceCatalog::from_frozen_entries_with_limits(
            one.clone(),
            limits(1024, 1, 3, 4),
        )
        .expect_err("one-over entry limit");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: 4,
                maximum: 3
            }
        ));

        let error =
            LogicalSourceCatalog::from_frozen_entries_with_limits(one, limits(1024, 1, 4, 3))
                .expect_err("one-over total limit");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::TotalBytes,
                observed: 4,
                maximum: 3
            }
        ));

        let two = frozen(vec![("a", Vec::new()), ("b", Vec::new())]);
        LogicalSourceCatalog::from_frozen_entries_with_limits(two.clone(), limits(2, 2, 1, 1))
            .expect("exact entry-count and metadata limits");
        let error = LogicalSourceCatalog::from_frozen_entries_with_limits(
            two.clone(),
            limits(1024, 1, 1, 1),
        )
        .expect_err("one-over entry count");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Entries,
                observed: 2,
                maximum: 1
            }
        ));

        let error = LogicalSourceCatalog::from_frozen_entries_with_limits(two, limits(1, 2, 1, 1))
            .expect_err("aggregate name metadata exceeds profile");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed: 2,
                maximum: 1
            }
        ));

        let exact_name = frozen(vec![("ab", Vec::new())]);
        LogicalSourceCatalog::from_frozen_entries_with_limits(
            exact_name.clone(),
            limits(2, 1, 1, 1),
        )
        .expect("exact member-name limit");
        let error =
            LogicalSourceCatalog::from_frozen_entries_with_limits(exact_name, limits(1, 1, 1, 1))
                .expect_err("one-over member-name limit");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MemberNameBytes,
                observed: 2,
                maximum: 1
            }
        ));
    }

    #[test]
    fn does_not_reinterpret_input_limit_as_logical_payload_limit() {
        let entries = frozen(vec![("a", vec![0; 32])]);
        LogicalSourceCatalog::from_frozen_entries_with_limits(entries, limits(1, 1, 32, 32))
            .expect("logical payload is governed by expanded limits");
    }

    #[test]
    fn rejects_unsafe_ambiguous_and_encrypted_entries() {
        for name in [
            "",
            "/Index/Document.iwa",
            "Index//Document.iwa",
            "Index/./Document.iwa",
            "Index/../Document.iwa",
            "Index\\Document.iwa",
            "C:/Index/Document.iwa",
        ] {
            let error = LogicalSourceCatalog::from_frozen_entries(frozen(vec![(name, Vec::new())]))
                .expect_err("unsafe logical name");
            assert!(matches!(error, Error::InvalidBundle(_)), "{name:?}");
        }

        for name in ["Index.zip", "nested/Index.zip"] {
            let error = LogicalSourceCatalog::from_frozen_entries(frozen(vec![(name, Vec::new())]))
                .expect_err("unexpanded nested index");
            assert!(matches!(error, Error::InvalidBundle(_)));
        }

        for name in [".iwpv2", "Metadata/.iwph"] {
            let error = LogicalSourceCatalog::from_frozen_entries(frozen(vec![(name, Vec::new())]))
                .expect_err("encrypted marker");
            assert!(matches!(error, Error::Encrypted));
        }
    }
}
