//! Bounded migration ingress for immutable, already-parsed cache records.

use super::{Component, ComponentArchive, ComponentCatalog, charge_semantic_iwa_objects};
use crate::{Error, LimitKind, Limits, Result};
use litchi_iwa_core::Archive;
use std::sync::Arc;

impl ComponentCatalog {
    /// Retain parsed component owners without copying their payloads.
    ///
    /// This unstable migration handoff accepts canonical component authorities
    /// from an already-validated source. It validates the parsed archive profile,
    /// cumulative encoded IWA bytes, names, entry count, and aggregate object
    /// count. It cannot prove ZIP provenance or compressed-source limits; the
    /// originating package remains responsible for those checks. No serialized
    /// source or full semantic document is constructed.
    ///
    /// # Errors
    ///
    /// Returns a typed error for invalid names, duplicate authorities, malformed
    /// archives, exceeded resource limits, or failed reservations.
    #[doc(hidden)]
    pub fn __from_shared_archives<'a>(
        archives: impl IntoIterator<Item = (&'a str, Arc<Archive>)>,
        limits: Limits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        let mut components = Vec::new();
        let mut metadata_bytes = 0_u64;
        let mut iwa_bytes = 0_u64;
        let mut objects = 0;
        for (name, archive) in archives {
            check(
                components.len().saturating_add(1) as u64,
                limits.max_entries() as u64,
                LimitKind::Entries,
            )?;
            check(
                name.len() as u64,
                Limits::MAX_MEMBER_NAME_BYTES,
                LimitKind::MemberNameBytes,
            )?;
            metadata_bytes = metadata_bytes.saturating_add(name.len() as u64);
            check(
                metadata_bytes,
                limits.max_metadata_bytes(),
                LimitKind::MetadataBytes,
            )?;
            if name.starts_with('/')
                || !name.ends_with(".iwa")
                || name.chars().any(char::is_control)
                || name.contains('\\')
                || name
                    .split('/')
                    .any(|part| part.is_empty() || matches!(part, "." | "..") || part.contains(':'))
            {
                return Err(Error::InvalidBundle(
                    "invalid shared IWA component authority".to_owned(),
                ));
            }
            objects = charge_semantic_iwa_objects(objects, archive.objects.len())?;
            let encoded = archive.encoded_len_with_limits(limits.archive_limits())?;
            iwa_bytes = limits.charge_iwa_total_bytes(iwa_bytes, encoded as u64)?;
            let mut owned_name = String::new();
            owned_name
                .try_reserve_exact(name.len())
                .map_err(|_| Error::Allocation {
                    resource: "shared IWA component name",
                    amount: name.len(),
                })?;
            owned_name.push_str(name);
            components.try_reserve(1).map_err(|_| Error::Allocation {
                resource: "shared IWA component catalog",
                amount: 1,
            })?;
            components.push(Component {
                name: owned_name.into_boxed_str(),
                archive: ComponentArchive::Shared(archive),
            });
        }
        components.sort_unstable_by(|left, right| left.name().cmp(right.name()));
        if components
            .windows(2)
            .any(|pair| pair[0].name() == pair[1].name())
        {
            return Err(Error::InvalidBundle(
                "duplicate shared IWA component authority".to_owned(),
            ));
        }
        Ok(Self {
            components: components.into_boxed_slice(),
        })
    }
}

fn check(observed: u64, maximum: u64, kind: LimitKind) -> Result<()> {
    if observed > maximum {
        return Err(Error::Limit {
            kind,
            observed,
            maximum,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_core::{ArchiveLimits, ArchiveObject, RawMessage};

    fn archive_with_objects(object_count: usize) -> Result<Arc<Archive>> {
        let mut objects = Vec::with_capacity(object_count);
        for identifier in 0..object_count {
            objects.push(ArchiveObject::new(
                u64::try_from(identifier).map_err(|_error| {
                    Error::InvalidBundle("test object identifier overflowed u64".to_owned())
                })? + 1,
                vec![RawMessage {
                    type_: 6_000,
                    data: vec![1, 2, 3],
                }],
            )?);
        }
        Ok(Arc::new(Archive { objects }))
    }

    fn ingress_limits(max_entries: usize, max_total_bytes: u64) -> Result<Limits> {
        Limits::new(
            64 * 1024,
            max_entries,
            64 * 1024,
            max_total_bytes,
            64 * 1024,
        )
    }

    #[test]
    fn shared_ingress_retains_archive_allocation_and_owns_component_name() -> Result<()> {
        let archive = archive_with_objects(1)?;
        let archive_pointer = Arc::as_ptr(&archive);
        let archive_lifetime = Arc::downgrade(&archive);
        let limits = ingress_limits(1, 64 * 1024)?;
        let catalog = {
            let name = String::from("Index/Document.iwa");
            ComponentCatalog::__from_shared_archives(
                [(name.as_str(), Arc::clone(&archive))],
                limits,
            )?
        };

        let component = catalog
            .get("Index/Document.iwa")
            .ok_or_else(|| Error::InvalidBundle("shared component was not retained".to_owned()))?;
        let component_pointer: *const Archive = component.archive();
        assert_eq!(component_pointer, archive_pointer);
        assert_eq!(component.archive().objects.len(), 1);

        drop(archive);
        assert!(archive_lifetime.upgrade().is_some());
        assert_eq!(component.name(), "Index/Document.iwa");

        drop(catalog);
        assert!(archive_lifetime.upgrade().is_none());
        Ok(())
    }

    #[test]
    fn shared_ingress_rejects_duplicate_and_noncanonical_names() -> Result<()> {
        let archive = archive_with_objects(1)?;
        let duplicate = ComponentCatalog::__from_shared_archives(
            [
                ("Index/Document.iwa", Arc::clone(&archive)),
                ("Index/Document.iwa", Arc::clone(&archive)),
            ],
            ingress_limits(2, 64 * 1024)?,
        );
        assert!(matches!(
            duplicate,
            Err(Error::InvalidBundle(message))
                if message.contains("duplicate shared IWA component authority")
        ));

        // Physical component catalogs admit portable auxiliary IWA members
        // outside Index/, just like the ordinary archive ingress.
        for name in ["Document.iwa", "Data/Document.iwa"] {
            let catalog = ComponentCatalog::__from_shared_archives(
                [(name, Arc::clone(&archive))],
                ingress_limits(1, 64 * 1024)?,
            )?;
            assert!(catalog.get(name).is_some());
        }
        for name in [
            "/Index/Document.iwa",
            "Index/Document",
            "Index/Document.zip",
            "Index//Document.iwa",
            "Index/./Document.iwa",
            "Index/../Document.iwa",
            "Index/Document:iwa",
            "Index/Document\\iwa",
            "Index/Doc\nument.iwa",
        ] {
            let result = ComponentCatalog::__from_shared_archives(
                [(name, Arc::clone(&archive))],
                ingress_limits(1, 64 * 1024)?,
            );
            assert!(
                matches!(
                    result,
                    Err(Error::InvalidBundle(message))
                        if message.contains("invalid shared IWA component authority")
                ),
                "name should be refused: {name:?}"
            );
        }
        Ok(())
    }

    #[test]
    fn shared_ingress_accepts_exact_entry_budget_and_rejects_one_below() -> Result<()> {
        let alpha = archive_with_objects(1)?;
        let bravo = archive_with_objects(1)?;
        let exact = ComponentCatalog::__from_shared_archives(
            [
                ("Index/Alpha.iwa", Arc::clone(&alpha)),
                ("Index/Bravo.iwa", Arc::clone(&bravo)),
            ],
            ingress_limits(2, 64 * 1024)?,
        )?;
        assert_eq!(exact.len(), 2);

        let one_below = ComponentCatalog::__from_shared_archives(
            [
                ("Index/Alpha.iwa", Arc::clone(&alpha)),
                ("Index/Bravo.iwa", Arc::clone(&bravo)),
            ],
            ingress_limits(1, 64 * 1024)?,
        );
        assert!(matches!(
            one_below,
            Err(Error::Limit {
                kind: LimitKind::Entries,
                observed: 2,
                maximum: 1,
            })
        ));
        Ok(())
    }

    #[test]
    fn shared_ingress_accepts_exact_aggregate_iwa_budget_and_rejects_one_below() -> Result<()> {
        let alpha = archive_with_objects(1)?;
        let bravo = archive_with_objects(1)?;
        let alpha_bytes = u64::try_from(alpha.encoded_len()?).map_err(|_error| {
            Error::InvalidBundle("test IWA length does not fit u64".to_owned())
        })?;
        let bravo_bytes = u64::try_from(bravo.encoded_len()?).map_err(|_error| {
            Error::InvalidBundle("test IWA length does not fit u64".to_owned())
        })?;
        let exact_total = alpha_bytes.checked_add(bravo_bytes).ok_or_else(|| {
            Error::InvalidBundle("test IWA aggregate length overflowed u64".to_owned())
        })?;

        let exact = ComponentCatalog::__from_shared_archives(
            [
                ("Index/Alpha.iwa", Arc::clone(&alpha)),
                ("Index/Bravo.iwa", Arc::clone(&bravo)),
            ],
            ingress_limits(2, exact_total)?,
        )?;
        assert_eq!(exact.len(), 2);

        let one_below = ComponentCatalog::__from_shared_archives(
            [
                ("Index/Alpha.iwa", Arc::clone(&alpha)),
                ("Index/Bravo.iwa", Arc::clone(&bravo)),
            ],
            ingress_limits(2, exact_total - 1)?,
        );
        assert!(matches!(
            one_below,
            Err(Error::Limit {
                kind: LimitKind::IwaTotalBytes,
                observed,
                maximum,
            }) if observed == exact_total && maximum == exact_total - 1
        ));
        Ok(())
    }

    #[test]
    fn shared_ingress_accepts_exact_archive_profile_and_rejects_one_below() -> Result<()> {
        let archive = archive_with_objects(2)?;
        let encoded = archive.encoded_len()?;
        let exact_profile = ArchiveLimits::default()
            .with_archive_bytes(encoded)?
            .with_objects(2)?;
        let exact_limits = ingress_limits(1, 64 * 1024)?.with_archive_limits(exact_profile)?;
        let exact = ComponentCatalog::__from_shared_archives(
            [("Index/Document.iwa", Arc::clone(&archive))],
            exact_limits,
        )?;
        assert_eq!(exact.len(), 1);

        let below_archive_profile = ArchiveLimits::default().with_archive_bytes(encoded - 1)?;
        let below_archive_limits =
            ingress_limits(1, 64 * 1024)?.with_archive_limits(below_archive_profile)?;
        let below_archive = ComponentCatalog::__from_shared_archives(
            [("Index/Document.iwa", Arc::clone(&archive))],
            below_archive_limits,
        );
        assert!(matches!(
            below_archive,
            Err(Error::Iwa(litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::ArchiveBytes,
                observed,
                maximum,
            })) if observed == encoded && maximum == encoded - 1
        ));

        let below_objects_profile = ArchiveLimits::default()
            .with_archive_bytes(encoded)?
            .with_objects(1)?;
        let below_objects_limits =
            ingress_limits(1, 64 * 1024)?.with_archive_limits(below_objects_profile)?;
        let below_objects = ComponentCatalog::__from_shared_archives(
            [("Index/Document.iwa", Arc::clone(&archive))],
            below_objects_limits,
        );
        assert!(matches!(
            below_objects,
            Err(Error::Iwa(litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::Objects,
                observed: 2,
                maximum: 1,
            }))
        ));
        Ok(())
    }

    #[test]
    fn shared_component_consumption_unwraps_the_retained_archive() -> Result<()> {
        let archive = archive_with_objects(1)?;
        let objects_pointer = archive.objects.as_ptr();
        let catalog = ComponentCatalog::__from_shared_archives(
            [("Index/Document.iwa", Arc::clone(&archive))],
            ingress_limits(1, 64 * 1024)?,
        )?;
        drop(archive);

        let component = catalog
            .into_iter()
            .next()
            .ok_or_else(|| Error::InvalidBundle("shared component was not retained".to_owned()))?;
        let (name, owned_archive) = component.into_parts();
        assert_eq!(name, "Index/Document.iwa");
        assert_eq!(owned_archive.objects.len(), 1);
        assert_eq!(owned_archive.objects.as_ptr(), objects_pointer);
        Ok(())
    }
}
