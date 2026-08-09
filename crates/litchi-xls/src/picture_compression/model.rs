//! Semantic settings and lossless record containers.

use crate::{Error, Result};

use super::RECORD_TYPE;

pub(crate) const HEADER_LEN: usize = 12;
pub(crate) const MIN_PAYLOAD_LEN: usize = HEADER_LEN + 4;
pub(crate) const MAX_RECORD_PAYLOAD: usize = 8_224;
pub(crate) const MAX_RECORDS: usize = 1_024;
pub(crate) const MAX_STREAM_BYTES: usize = 1_048_576;

/// Typed `CompressPictures` payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Settings {
    /// The recommendation encoded by `fAutoCompressPictures`.
    recommend_compression: bool,
    /// The complete future-record header, retained so unknown header bytes
    /// survive an untouched round trip.
    header: [u8; HEADER_LEN],
    /// Producer extension bytes after the specified payload.
    opaque_tail: Box<[u8]>,
}

impl Settings {
    /// Creates a canonical record with a zeroed `FrtHeader` except for `rt`.
    #[must_use]
    pub fn new(recommend_compression: bool) -> Self {
        let mut header = [0; HEADER_LEN];
        header[..2].copy_from_slice(&RECORD_TYPE.to_le_bytes());
        Self {
            recommend_compression,
            header,
            opaque_tail: Box::default(),
        }
    }

    /// Parses one `CompressPictures` payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(payload: &[u8]) -> Result<Self> {
        super::codec::parse_settings(payload)
    }

    #[must_use]
    pub const fn recommends_compression(&self) -> bool {
        self.recommend_compression
    }

    /// Bytes of the future-record header, including the validated `rt`.
    #[must_use]
    pub const fn header(&self) -> &[u8; HEADER_LEN] {
        &self.header
    }

    /// Retained bytes not modeled by the current MS-XLS owner.
    #[must_use]
    pub fn opaque_tail(&self) -> &[u8] {
        &self.opaque_tail
    }

    pub(crate) fn from_wire(
        recommend_compression: bool,
        header: [u8; HEADER_LEN],
        opaque_tail: Box<[u8]>,
    ) -> Self {
        Self {
            recommend_compression,
            header,
            opaque_tail,
        }
    }

    pub(crate) const fn payload_len(&self) -> usize {
        MIN_PAYLOAD_LEN + self.opaque_tail.len()
    }

    pub(crate) fn payload(&self) -> Result<Vec<u8>> {
        let mut payload = Vec::with_capacity(self.payload_len());
        payload.extend_from_slice(&self.header);
        payload.extend_from_slice(&u32::from(self.recommend_compression).to_le_bytes());
        payload.extend_from_slice(&self.opaque_tail);
        if payload.len() > MAX_RECORD_PAYLOAD {
            return Err(invalid("CompressPictures payload exceeds 8224 bytes"));
        }
        Ok(payload)
    }
}

/// A BIFF record retained without interpretation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unknown {
    record_type: u16,
    payload: Box<[u8]>,
}

impl Unknown {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(record_type: u16, payload: impl Into<Vec<u8>>) -> Result<Self> {
        if record_type == RECORD_TYPE {
            return Err(invalid("CompressPictures must use Record::Settings"));
        }
        let payload = payload.into();
        if payload.len() > MAX_RECORD_PAYLOAD {
            return Err(invalid("unknown BIFF payload exceeds 8224 bytes"));
        }
        Ok(Self {
            record_type,
            payload: payload.into_boxed_slice(),
        })
    }

    #[must_use]
    pub const fn record_type(&self) -> u16 {
        self.record_type
    }

    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }
}

/// One record in a detached `CompressPictures` snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Record {
    Settings(Settings),
    Unknown(Unknown),
}

impl Record {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn unknown(record_type: u16, payload: impl Into<Vec<u8>>) -> Result<Self> {
        Unknown::try_new(record_type, payload).map(Self::Unknown)
    }
}

/// Bounded, record-order snapshot for the owner.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) records: Vec<Record>,
}

impl Snapshot {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(records: Vec<Record>) -> Result<Self> {
        let value = Self { records };
        super::validation::validate(&value)?;
        Ok(value)
    }

    #[must_use]
    pub const fn empty() -> Self {
        Self {
            records: Vec::new(),
        }
    }

    #[must_use]
    pub fn records(&self) -> &[Record] {
        &self.records
    }

    #[must_use]
    pub fn settings(&self) -> Option<&Settings> {
        self.records.iter().find_map(|record| match record {
            Record::Settings(value) => Some(value),
            Record::Unknown(_) => None,
        })
    }

    pub fn unknown_records(&self) -> impl Iterator<Item = &Unknown> {
        self.records.iter().filter_map(|record| match record {
            Record::Unknown(value) => Some(value),
            Record::Settings(_) => None,
        })
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn encode(&self) -> Result<Vec<u8>> {
        super::codec::write(self)
    }

    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    pub(crate) fn validate(&self) -> Result<()> {
        super::validation::validate(self)
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: RECORD_TYPE,
        message: message.into(),
    }
}
