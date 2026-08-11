//! Compact package-local object lookup for the Numbers adapter.

use std::mem::size_of;

use litchi_iwa_core::RawMessage;

use super::{Components, Error, Result, SemanticLimitKind, SemanticPath};

#[derive(Debug, Clone, Copy)]
pub(super) struct Entry {
    identifier: u64,
    primary_message_type: u32,
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
    pub(super) component_index: usize,
    pub(super) object_index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ObjectLocator {
    identifier: u64,
    component: u32,
    object: u32,
}

#[derive(Debug)]
pub(super) struct Index {
    locators: Box<[ObjectLocator]>,
    primary_entries: Box<[Entry]>,
}

impl Index {
    pub(super) fn from_components(components: &Components, max_objects: usize) -> Result<Self> {
        let object_count = components
            .iter_objects()
            .try_fold(0usize, |count, _object| {
                count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Numbers object count overflows usize".to_owned())
                })
            })?;
        if object_count > max_objects {
            return Err(Error::SemanticLimit {
                kind: SemanticLimitKind::Objects,
                observed: object_count,
                maximum: max_objects,
                path: SemanticPath::Package,
            });
        }

        let mut locators = Vec::new();
        locators.try_reserve_exact(object_count).map_err(|_error| {
            Error::Common(litchi_iwa_common::Error::Allocation {
                resource: "Numbers object locators",
                amount: object_count,
            })
        })?;
        let mut primary_entries = Vec::new();
        primary_entries
            .try_reserve_exact(object_count)
            .map_err(|_error| {
                Error::Common(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers primary object types",
                    amount: object_count,
                })
            })?;

        for (component_index, component_entry) in components.catalog().iter().enumerate() {
            let component = u32::try_from(component_index).map_err(|_error| {
                Error::InvalidFormat("Numbers component count exceeds compact indexing".to_owned())
            })?;
            for (object_index, object_entry) in component_entry.archive().objects.iter().enumerate()
            {
                let identifier = object_entry.archive_info.identifier.ok_or_else(|| {
                    Error::InvalidFormat("Numbers IWA object has no identifier".to_owned())
                })?;
                if identifier == 0 {
                    return Err(Error::InvalidFormat(
                        "Numbers IWA object uses the null identifier".to_owned(),
                    ));
                }
                let object = u32::try_from(object_index).map_err(|_error| {
                    Error::InvalidFormat(
                        "Numbers component object count exceeds compact indexing".to_owned(),
                    )
                })?;
                locators.push(ObjectLocator {
                    identifier,
                    component,
                    object,
                });
                if let Some(message) = object_entry.messages.first() {
                    primary_entries.push(Entry {
                        identifier,
                        primary_message_type: message.type_,
                    });
                }
            }
        }

        locators.sort_unstable_by_key(|locator| locator.identifier);
        if locators
            .windows(2)
            .any(|pair| pair[0].identifier == pair[1].identifier)
        {
            return Err(Error::InvalidFormat(
                "Numbers package contains duplicate object identities".to_owned(),
            ));
        }
        primary_entries
            .sort_unstable_by_key(|entry| (entry.primary_message_type, entry.identifier));
        Ok(Self {
            locators: locators.into_boxed_slice(),
            primary_entries: primary_entries.into_boxed_slice(),
        })
    }

    pub(super) fn object_count(&self) -> usize {
        self.locators.len()
    }

    /// Conservative comparison work for one binary object-identifier lookup.
    pub(super) const fn lookup_work(&self) -> usize {
        let object_count = self.locators.len();
        if object_count <= 1 {
            return 1;
        }
        usize::BITS
            .saturating_sub((object_count - 1).leading_zeros())
            .saturating_add(1) as usize
    }

    /// Conservative allocation/population/sort cost for rebuilding this index.
    pub(super) fn rebuild_work(&self) -> usize {
        let object_count = self.locators.len();
        let allocation = object_count
            .saturating_mul(size_of::<ObjectLocator>().saturating_add(size_of::<Entry>()));
        let populate_and_sort = object_count
            .saturating_mul(self.lookup_work())
            .saturating_mul(2)
            .saturating_add(object_count);
        allocation.saturating_add(populate_and_sort)
    }

    pub(super) fn iter_entries_by_type(
        &self,
        message_type: u32,
    ) -> impl Iterator<Item = Entry> + '_ {
        let start = self
            .primary_entries
            .partition_point(|entry| entry.primary_message_type < message_type);
        let end = self
            .primary_entries
            .partition_point(|entry| entry.primary_message_type <= message_type);
        self.primary_entries[start..end].iter().copied()
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
        let Ok(position) = self
            .locators
            .binary_search_by_key(&identifier, |locator| locator.identifier)
        else {
            return Ok(None);
        };
        let locator = self.locators[position];
        let component_index = usize::try_from(locator.component).map_err(|_error| {
            Error::InvalidFormat("Numbers object locator is invalid".to_owned())
        })?;
        let object_index = usize::try_from(locator.object).map_err(|_error| {
            Error::InvalidFormat("Numbers object locator is invalid".to_owned())
        })?;
        let object = components
            .catalog()
            .get_index(component_index)
            .and_then(|component| component.archive().objects.get(object_index))
            .ok_or_else(|| Error::InvalidFormat("Numbers object locator is invalid".to_owned()))?;
        Ok(Some(Resolved {
            messages: &object.messages,
            component_index,
            object_index,
        }))
    }
}
