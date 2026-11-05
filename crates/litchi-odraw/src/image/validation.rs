//! `OfficeArt` image invariants and cross-record validation.

use super::LimitKind;
use super::model::{Blip, Block, Context, Entry, Kind, Limits, Offset, Storage, Store};
use crate::{Error, Record, RecordKind, Result};

impl<'data> Entry<'data> {
    /// Returns the physical storage selected by this entry.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedImage` when no embedded BLIP matches the
    /// admitted instance, or a nonempty entry has neither an embedded nor
    /// a delayed BLIP.
    pub fn storage(&self) -> Result<Storage<'data>> {
        if !self.embedded.is_empty() {
            let instance = self.admitted_instance();
            for blip in self.embedded() {
                let embedded_blip = blip?;
                if matches_instance(&embedded_blip, instance) {
                    return Ok(Storage::Embedded(embedded_blip));
                }
            }
            return Err(Error::MalformedImage {
                reason: "FBSE has no embedded BLIP matching its instance",
            });
        }
        if self.refs == 0 {
            return Ok(Storage::Empty);
        }
        if self.delay == u32::MAX {
            return Err(Error::MalformedImage {
                reason: "nonempty FBSE has no embedded or delayed BLIP",
            });
        }
        Ok(Storage::Delay(Offset(self.delay)))
    }

    /// Resolves this entry using a host-supplied context.
    ///
    /// # Errors
    ///
    /// Returns an error from [`Entry::storage`] when the entry is malformed,
    /// `Error::MissingDelayStore` when a delayed BLIP lacks a delay store, or
    /// `Error::MalformedImage` when the delay offset does not point to a BLIP
    /// or the resolved BLIP fails cross-record validation. A delayed BLIP must
    /// satisfy both the entry's and the context's BLIP byte limits; otherwise this
    /// returns `Error::ImageLimitExceeded`.
    pub fn resolve(&self, context: Context<'data>) -> Result<Option<Blip<'data>>> {
        match self.storage()? {
            Storage::Embedded(blip) => Ok(Some(blip)),
            Storage::Empty => Ok(None),
            Storage::Delay(offset) => {
                let mut delay = context.delay.ok_or(Error::MissingDelayStore)?;
                delay.limits = delay.limits.constrained_by(self.limits);
                match delay.at(offset)? {
                    Block::Blip(blip) => {
                        self.validate_selected(&blip, true)?;
                        Ok(Some(blip))
                    },
                    Block::Entry(_) => Err(Error::MalformedImage {
                        reason: "FBSE delay offset does not point to a BLIP",
                    }),
                }
            },
        }
    }

    pub(super) fn validate_selected(&self, blip: &Blip<'_>, delayed: bool) -> Result<()> {
        let instance = self.admitted_instance();
        if !matches_instance(blip, instance) {
            return Err(Error::MalformedImage {
                reason: "resolved BLIP kind does not match the FBSE instance",
            });
        }
        if delayed {
            let actual = blip
                .record()
                .len()
                .checked_add(8)
                .ok_or(Error::ArithmeticOverflow {
                    context: "resolved BLIP wire length",
                })?;
            if actual != self.size {
                let actual_size =
                    usize::try_from(actual).map_err(|_err| Error::ArithmeticOverflow {
                        context: "resolved BLIP wire length",
                    })?;
                return Err(Error::ImageSizeMismatch {
                    field: "FBSE size",
                    declared: u64::from(self.size),
                    actual: actual_size,
                });
            }
        }
        // FBSE.rgbUid identifies pixel data (MS-ODRAW 2.2.32), while a
        // BLIP's rgbUid identifies its uncompressed file data. The format
        // does not require those fields to be equal, and native PowerPoint
        // stores distinct values. The slot's foDelay, kind and size select
        // the BLIP; both identifiers remain available without rewriting.
        Ok(())
    }

    fn admitted_instance(&self) -> u8 {
        // Entry::parse is the sole constructor and checks that the instance
        // fits in a byte before publishing the immutable entry. Reuse that
        // admission fact when selecting or resolving a BLIP.
        debug_assert!(u8::try_from(self.record.instance()).is_ok());
        self.record.instance().to_le_bytes()[0]
    }
}

impl<'data> Store<'data> {
    /// Validates a previously parsed `BStore` record under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedImage` when the record is not an `OfficeArt`
    /// `BStore` container, contains an invalid child, or its file-block count
    /// does not match `recInstance`, and `Error::ImageLimitExceeded` when the
    /// raw record or entry count exceeds its configured image limit.
    pub fn from_record(record: Record<'data>, limits: Limits) -> Result<Self> {
        let effective_records = record.limits().constrained_by(limits.records);
        if record.len() > effective_records.max_record_bytes {
            return Err(Error::ImageLimitExceeded {
                limit: LimitKind::RecordBytes,
                observed: u64::from(record.len()),
                maximum: u64::from(effective_records.max_record_bytes),
            });
        }
        if record.kind() != RecordKind::BStoreContainer
            || record.version() != 0x0F
            || !record.is_container()
        {
            return Err(Error::MalformedImage {
                reason: "record is not an OfficeArt BStore container",
            });
        }
        let count = record.instance();
        if count > limits.max_store_entries {
            return Err(Error::ImageLimitExceeded {
                limit: LimitKind::StoreEntries,
                observed: u64::from(count),
                maximum: u64::from(limits.max_store_entries),
            });
        }
        // Keep the image validator in charge of classifying its raw record
        // budget. The generic iterator still enforces that budget for direct
        // callers, but this admission path reports the image-specific kind.
        let parser_limits = crate::Limits {
            max_records: u32::MAX,
            ..effective_records
        };
        let mut visited = 0_u32;
        let mut actual = 0u16;
        for child in crate::Children::new(record.data(), parser_limits) {
            child.map_err(super::map_traversal_limit)?;
            visited = visited.checked_add(1).ok_or(Error::ImageLimitExceeded {
                limit: LimitKind::Records,
                observed: u64::from(visited) + 1,
                maximum: u64::from(effective_records.max_records),
            })?;
            if visited > effective_records.max_records {
                return Err(Error::ImageLimitExceeded {
                    limit: LimitKind::Records,
                    observed: u64::from(visited),
                    maximum: u64::from(effective_records.max_records),
                });
            }
            actual = actual.checked_add(1).ok_or(Error::MalformedImage {
                reason: "BStore file-block count exceeds 4095",
            })?;
        }
        if actual != count {
            return Err(Error::MalformedImage {
                reason: "BStore file-block count does not match recInstance",
            });
        }
        // A caller may validate an already admitted record under a different
        // policy. Rebind the borrowed view after admission to the exact
        // intersection: neither the origin admission nor the caller's
        // narrower dimensions may be weakened by the compact Store state.
        let record = record.with_limits(effective_records);
        Ok(Self {
            record,
            count,
            max_blip_bytes: limits.max_blip_bytes,
            max_store_entries: limits.max_store_entries,
            max_files: limits.max_files,
        })
    }
}

pub(super) fn matches_instance(blip: &Blip<'_>, instance: u8) -> bool {
    blip.kind().raw() == instance
        || matches!(blip, Blip::Opaque(_))
            && matches!(Kind::from_raw(instance), Kind::Unknown | Kind::Other(_))
}

pub(super) fn validate_embedded(entry: &Entry<'_>, instance: u8) -> Result<()> {
    if entry.embedded.is_empty() {
        return Ok(());
    }

    let mut selected = false;
    for record in crate::Children::new(entry.embedded, entry.limits.record_limits()) {
        let blip = record
            .map_err(super::map_traversal_limit)
            .and_then(|child| Blip::from_record(child, entry.limits))?;
        if matches_instance(&blip, instance) {
            entry.validate_selected(&blip, false)?;
            selected = true;
        }
    }
    if !selected {
        return Err(Error::MalformedImage {
            reason: "FBSE has no embedded BLIP matching its instance",
        });
    }
    Ok(())
}
