//! Compact package-local object lookup for the Numbers adapter.

use std::collections::HashSet;

use litchi_iwa_core::RawMessage;

use super::{Components, Error, Result};

#[derive(Debug, Clone, Copy)]
pub(super) struct Entry {
    identifier: u64,
    message_type: u32,
}

impl Entry {
    pub(super) const fn id(self) -> u64 {
        self.identifier
    }
}

/// A borrowed native object selected through the package-local index.
#[derive(Debug, Clone, Copy)]
pub(super) struct Resolved<'a> {
    pub(super) messages: &'a [RawMessage],
}

#[derive(Debug)]
pub(super) struct Index {
    entries: Box<[Entry]>,
    object_count: usize,
}

impl Index {
    pub(super) fn from_components(components: &Components) -> Result<Self> {
        let mut identifiers = HashSet::new();
        let mut entries = Vec::new();
        let mut object_count = 0usize;
        for object in components.iter_objects() {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat("Numbers IWA object has no identifier".to_owned())
            })?;
            if !identifiers.insert(identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers package contains duplicate object identifier {identifier}"
                )));
            }
            object_count = object_count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Numbers object count overflows usize".to_owned())
            })?;
            entries
                .try_reserve(object.messages.len())
                .map_err(|_error| {
                    Error::Common(litchi_iwa_common::Error::Allocation {
                        resource: "Numbers object index entries",
                        amount: object.messages.len(),
                    })
                })?;
            entries.extend(object.messages.iter().map(|message| Entry {
                identifier,
                message_type: message.type_,
            }));
        }
        entries.sort_unstable_by_key(|entry| (entry.message_type, entry.identifier));
        Ok(Self {
            entries: entries.into_boxed_slice(),
            object_count,
        })
    }

    pub(super) fn object_count(&self) -> usize {
        self.object_count
    }

    pub(super) fn iter_entries_by_type(
        &self,
        message_type: u32,
    ) -> impl Iterator<Item = Entry> + '_ {
        let start = self
            .entries
            .partition_point(|entry| entry.message_type < message_type);
        let end = self
            .entries
            .partition_point(|entry| entry.message_type <= message_type);
        self.entries[start..end].iter().copied()
    }

    pub(super) fn resolve_ref<'a>(
        &self,
        components: &'a Components,
        identifier: u64,
    ) -> Result<Option<Resolved<'a>>> {
        self.resolve_ref_id(components, identifier)
    }

    pub(super) fn resolve_ref_id<'a>(
        &self,
        components: &'a Components,
        identifier: u64,
    ) -> Result<Option<Resolved<'a>>> {
        Ok(components.find_object(identifier).map(|object| Resolved {
            messages: &object.messages,
        }))
    }
}
