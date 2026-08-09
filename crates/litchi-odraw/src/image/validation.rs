//! `OfficeArt` image invariants and cross-record validation.

use super::model::{Blip, Block, Context, Entry, Kind, Limits, Offset, Storage, Store};
use crate::{Error, ImageLimit, Record, RecordKind, Result};

impl<'data> Entry<'data> {
    /// Returns the physical storage selected by this entry.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedImage` when the FBSE instance is not a valid
    /// `MSOBLIPTYPE` value, no embedded BLIP matches the instance, or a
    /// nonempty entry has neither an embedded nor a delayed BLIP.
    pub fn storage(&self) -> Result<Storage<'data>> {
        if !self.embedded.is_empty() {
            let instance =
                u8::try_from(self.record.instance()).map_err(|_err| Error::MalformedImage {
                    reason: "FBSE instance is not an MSOBLIPTYPE value",
                })?;
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
    /// or the resolved BLIP fails cross-record validation.
    pub fn resolve(&self, context: Context<'data>) -> Result<Option<Blip<'data>>> {
        match self.storage()? {
            Storage::Embedded(blip) => Ok(Some(blip)),
            Storage::Empty => Ok(None),
            Storage::Delay(offset) => {
                let delay = context.delay.ok_or(Error::MissingDelayStore)?;
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
        let instance =
            u8::try_from(self.record.instance()).map_err(|_err| Error::MalformedImage {
                reason: "FBSE instance is not an MSOBLIPTYPE value",
            })?;
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
        if let Some(uids) = blip.uids()
            && uids.effective() != self.uid
        {
            return Err(Error::MalformedImage {
                reason: "resolved BLIP UID does not match the FBSE UID",
            });
        }
        Ok(())
    }
}

impl<'data> Store<'data> {
    /// Validates a previously parsed `BStore` record.
    ///
    /// # Errors
    ///
    /// Returns an error from [`Store::from_record_with`] under the default
    /// limits.
    pub fn from_record(record: Record<'data>) -> Result<Self> {
        Self::from_record_with(record, Limits::default())
    }

    /// Validates a previously parsed `BStore` record under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedImage` when the record is not an `OfficeArt`
    /// `BStore` container, contains an invalid child, or its file-block count
    /// does not match `recInstance`, and `Error::ImageLimitExceeded` when the
    /// entry count exceeds `limits.max_store_entries`.
    pub fn from_record_with(record: Record<'data>, limits: Limits) -> Result<Self> {
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
                limit: ImageLimit::StoreEntries,
                maximum: u64::from(limits.max_store_entries),
            });
        }
        let mut actual = 0u16;
        for child in crate::Children::new(record.data()) {
            child?;
            actual = actual.checked_add(1).ok_or(Error::MalformedImage {
                reason: "BStore file-block count exceeds 4095",
            })?;
        }
        if actual != count {
            return Err(Error::MalformedImage {
                reason: "BStore file-block count does not match recInstance",
            });
        }
        Ok(Self {
            record,
            count,
            limits,
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
    for record in crate::Children::new(entry.embedded) {
        let blip = record.and_then(|child| Blip::from_record_with(child, entry.limits))?;
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
