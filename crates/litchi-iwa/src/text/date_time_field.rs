//! Native Date & Time field identities and aggregates.
//!
//! The semantic formatter payload lives in `litchi_iwa_text::date_time`; this
//! module retains only the IWA object identifier and text-range association.

use crate::{Error, Result};

use litchi_iwa_text::date_time::Settings;
use litchi_iwa_text::position::TextRange;

/// Identifier of a native Date & Time smart-field object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextDateTimeFieldId(u64);

impl TextDateTimeFieldId {
    /// Construct an identifier obtained from a previously read field.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork Date & Time field object identifier cannot be zero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the underlying package object identifier.
    #[must_use]
    pub const fn object_id(self) -> u64 {
        self.0
    }

    pub(crate) const fn from_native(identifier: u64) -> Self {
        Self(identifier)
    }
}

/// One native Date & Time field attached to a nonempty UTF-16 range.
#[derive(Debug, Clone, PartialEq)]
pub struct TextDateTimeField {
    /// Native smart-field object identifier.
    pub id: TextDateTimeFieldId,
    /// Nonempty UTF-16 text range covered by the field.
    pub range: TextRange,
    /// Losslessly decoded semantic formatter payload.
    pub settings: Settings,
}

impl TextDateTimeField {
    pub(crate) fn new(id: TextDateTimeFieldId, range: TextRange, settings: Settings) -> Self {
        Self {
            id,
            range,
            settings,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::TextDateTimeFieldId;

    #[test]
    fn native_field_ids_are_nonzero_and_lossless() {
        assert!(TextDateTimeFieldId::from_object_id(0).is_err());
        let id = TextDateTimeFieldId::from_object_id(u64::MAX).unwrap();
        assert_eq!(id.object_id(), u64::MAX);
    }
}
