use std::borrow::Cow;

use crate::{RtfError, RtfResult};

pub const MAX_WRITE_RESERVATION_BYTES: usize = 65_536;

/// Opaque deprecated `writereservation` payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyWriteReservation<'a> {
    pub data: Cow<'a, str>,
}

impl<'a> LegacyWriteReservation<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(data: Cow<'a, str>) -> RtfResult<Self> {
        let value = Self { data };
        value.validate()?;
        Ok(value)
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.data.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF legacy write-reservation payload must not be empty".to_string(),
            ));
        }
        if self.data.len() > MAX_WRITE_RESERVATION_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF legacy write-reservation payload exceeds the resource limit".to_string(),
            ));
        }
        if self
            .data
            .chars()
            .any(|character| matches!(character, '\0' | '\r' | '\n'))
        {
            return Err(RtfError::MalformedDocument(
                "RTF legacy write-reservation payload contains a forbidden character".to_string(),
            ));
        }
        Ok(())
    }

    fn into_owned(self) -> LegacyWriteReservation<'static> {
        LegacyWriteReservation {
            data: Cow::Owned(self.data.into_owned()),
        }
    }
}

/// Opaque bytes from the modern `writereservhash` destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteReservationHash<'a> {
    pub data: Cow<'a, [u8]>,
}

impl<'a> WriteReservationHash<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(data: Cow<'a, [u8]>) -> RtfResult<Self> {
        let value = Self { data };
        value.validate()?;
        Ok(value)
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.data.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF write-reservation hash must not be empty".to_string(),
            ));
        }
        if self.data.len() > MAX_WRITE_RESERVATION_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF write-reservation hash exceeds the resource limit".to_string(),
            ));
        }
        Ok(())
    }

    fn into_owned(self) -> WriteReservationHash<'static> {
        WriteReservationHash {
            data: Cow::Owned(self.data.into_owned()),
        }
    }
}

/// Passive write-reservation metadata. Values are never authenticated or decrypted.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentWriteReservations<'a> {
    pub legacy: Option<LegacyWriteReservation<'a>>,
    pub hash: Option<WriteReservationHash<'a>>,
}

impl DocumentWriteReservations<'_> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if let Some(legacy) = &self.legacy {
            legacy.validate()?;
        }
        if let Some(hash) = &self.hash {
            hash.validate()?;
        }
        Ok(())
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.legacy.is_none() && self.hash.is_none()
    }

    pub(crate) fn into_owned(self) -> DocumentWriteReservations<'static> {
        DocumentWriteReservations {
            legacy: self.legacy.map(LegacyWriteReservation::into_owned),
            hash: self.hash.map(WriteReservationHash::into_owned),
        }
    }
}
