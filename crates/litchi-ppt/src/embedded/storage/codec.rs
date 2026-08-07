use std::borrow::Cow;
use std::io::Read;

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

use super::model::{Compression, Kind, MAX_DECLARED_BYTES, MAX_STORED_BYTES, Ref, Storage};
use super::snapshot::{Metadata, Snapshot};

#[allow(dead_code)]
impl<'a> Ref<'a> {
    /// Resolve one strict `ExOleObjStg` record directly from presentation bytes.
    pub(crate) fn parse_at(document: &'a [u8], offset: usize, kind: Kind) -> Result<Self> {
        let header_end = offset
            .checked_add(8)
            .ok_or_else(|| Error::Corrupted("ExOleObjStg header offset overflows".into()))?;
        let header: &[u8; 8] = document
            .get(offset..header_end)
            .and_then(|bytes| bytes.try_into().ok())
            .ok_or_else(|| Error::Corrupted("truncated ExOleObjStg header".into()))?;
        let version_instance = u16::from_le_bytes([header[0], header[1]]);
        let version = version_instance & 0x000f;
        let instance = (version_instance >> 4) & 0x0fff;
        let record_type = u16::from_le_bytes([header[2], header[3]]);
        let stored_len = usize::try_from(u32::from_le_bytes([
            header[4], header[5], header[6], header[7],
        ]))
        .map_err(|_| Error::Corrupted("ExOleObjStg size exceeds usize".into()))?;
        if version != 0 || record_type != RecordType::ExternalOleObjectStg.as_u16() {
            return corrupted("persisted storage is not a strict ExOleObjStg record");
        }
        if stored_len > MAX_STORED_BYTES {
            return corrupted("ExOleObjStg exceeds 128 MiB stored data");
        }
        let payload_end = header_end
            .checked_add(stored_len)
            .ok_or_else(|| Error::Corrupted("ExOleObjStg payload size overflows".into()))?;
        let payload = document
            .get(header_end..payload_end)
            .ok_or_else(|| Error::Corrupted("truncated ExOleObjStg payload".into()))?;
        match instance {
            0 => Ok(Self {
                kind,
                compression: Compression::Uncompressed,
                declared_uncompressed_len: None,
                data: payload,
            }),
            1 => {
                let prefix: &[u8; 4] = payload
                    .get(..4)
                    .and_then(|bytes| bytes.try_into().ok())
                    .ok_or_else(|| {
                        Error::Corrupted("compressed ExOleObjStg is missing its size prefix".into())
                    })?;
                let declared_uncompressed_len = u32::from_le_bytes(*prefix);
                if declared_uncompressed_len > MAX_DECLARED_BYTES {
                    return corrupted("compressed ExOleObjStg declares more than 256 MiB");
                }
                Ok(Self {
                    kind,
                    compression: Compression::Zlib,
                    declared_uncompressed_len: Some(declared_uncompressed_len),
                    data: payload.get(4..).ok_or_else(|| {
                        Error::Corrupted("compressed ExOleObjStg is missing its payload".into())
                    })?,
                })
            },
            _ => corrupted("ExOleObjStg has an invalid storage instance"),
        }
    }

    pub(crate) fn metadata(self) -> Metadata {
        self.snapshot().metadata()
    }

    /// Reject a caller-specific stored-size ceiling before payload allocation.
    pub(crate) fn check_stored_limit(self, maximum: usize) -> Result<Self> {
        if self.data.len() > maximum {
            return corrupted(format!(
                "VbaProjectStg stores {} bytes above the {maximum}-byte limit",
                self.data.len()
            ));
        }
        Ok(self)
    }

    /// Return bounded CFB bytes, borrowing the uncompressed representation.
    pub(crate) fn decompressed_bytes(self, maximum: usize) -> Result<Cow<'a, [u8]>> {
        match self.compression {
            Compression::Uncompressed => {
                if self.data.len() > maximum {
                    return corrupted(format!(
                        "uncompressed ExOleObjStg exceeds the {maximum}-byte output limit"
                    ));
                }
                Ok(Cow::Borrowed(self.data))
            },
            Compression::Zlib => {
                let declared = self.declared_uncompressed_len.ok_or_else(|| {
                    Error::Corrupted("compressed ExOleObjStg is missing its size".into())
                })?;
                let declared = usize::try_from(declared).map_err(|_| {
                    Error::Corrupted(
                        "compressed ExOleObjStg size does not fit in memory".to_string(),
                    )
                })?;
                if declared > maximum {
                    return corrupted(format!(
                        "compressed ExOleObjStg declares {declared} bytes above the {maximum}-byte output limit"
                    ));
                }
                let read_limit = maximum.checked_add(1).ok_or_else(|| {
                    Error::Corrupted("ExOleObjStg output limit overflows".to_string())
                })?;
                let read_limit = u64::try_from(read_limit)
                    .map_err(|_| Error::Corrupted("ExOleObjStg output limit exceeds u64".into()))?;
                let decoder = flate2::read::ZlibDecoder::new(self.data);
                let mut limited = decoder.take(read_limit);
                let mut output = Vec::with_capacity(declared.min(64 * 1024));
                limited.read_to_end(&mut output).map_err(|error| {
                    Error::Corrupted(format!("invalid ExOleObjStg zlib payload: {error}"))
                })?;
                let decoder = limited.into_inner();
                if output.len() > maximum {
                    return corrupted(format!(
                        "decompressed ExOleObjStg exceeds the {maximum}-byte output limit"
                    ));
                }
                if output.len() != declared {
                    return corrupted(format!(
                        "ExOleObjStg decompressed to {} bytes instead of declared {declared}",
                        output.len()
                    ));
                }
                let stored_len = u64::try_from(self.data.len())
                    .map_err(|_| Error::Corrupted("ExOleObjStg stored size exceeds u64".into()))?;
                if decoder.total_in() != stored_len {
                    return corrupted("ExOleObjStg zlib payload has trailing bytes");
                }
                Ok(Cow::Owned(output))
            },
        }
    }
}

impl Storage {
    /// Parse an embedded OLE-object storage.
    pub fn parse(record: &Record) -> Result<Self> {
        Self::parse_as(record, Kind::OleObject)
    }

    /// Parse a persisted storage using its referencing record's kind.
    pub fn parse_as(record: &Record, kind: Kind) -> Result<Self> {
        if record.version != 0
            || record.record_type_raw != RecordType::ExternalOleObjectStg.as_u16()
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
            || record.data.len() > MAX_STORED_BYTES
        {
            return corrupted("ExOleObjStg has an invalid header or stored size");
        }
        match record.instance {
            0 => Self::uncompressed(kind, record.data.clone()),
            1 => {
                let prefix: &[u8; 4] = record
                    .data
                    .get(..4)
                    .and_then(|bytes| bytes.try_into().ok())
                    .ok_or_else(|| {
                        Error::Corrupted("compressed ExOleObjStg is missing its size prefix".into())
                    })?;
                let declared = u32::from_le_bytes(*prefix);
                Self::compressed(
                    kind,
                    declared,
                    record
                        .data
                        .get(4..)
                        .ok_or_else(|| {
                            Error::Corrupted("compressed ExOleObjStg is missing its payload".into())
                        })?
                        .to_vec(),
                )
            },
            _ => corrupted("ExOleObjStg has an invalid storage instance"),
        }
    }

    /// Return the uncompressed structured-storage bytes with an explicit cap.
    pub fn decompressed_bytes(&self, maximum: usize) -> Result<Vec<u8>> {
        Ref {
            kind: self.kind,
            compression: self.compression,
            declared_uncompressed_len: self.declared_uncompressed_len,
            data: &self.data,
        }
        .decompressed_bytes(maximum)
        .map(Cow::into_owned)
    }

    /// Encode the complete `ExOleObjStg` record.
    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    /// Encode the complete record without an intermediate parsed view.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.snapshot().to_record_bytes()
    }
}

impl<'a> Snapshot<'a> {
    /// Encode the complete `ExOleObjStg` record represented by this view.
    pub fn to_record_bytes(self) -> Result<Vec<u8>> {
        let (instance, declared) = match self.compression() {
            Compression::Uncompressed => (0u16, None),
            Compression::Zlib => (
                1u16,
                Some(self.declared_uncompressed_len().ok_or_else(|| {
                    Error::Corrupted("compressed ExOleObjStg is missing its size".into())
                })?),
            ),
        };
        let data_len = self.record_payload_len();
        if data_len > MAX_STORED_BYTES {
            return corrupted("ExOleObjStg exceeds 128 MiB stored data");
        }
        if let Some(declared) = declared {
            if declared > MAX_DECLARED_BYTES {
                return corrupted("compressed ExOleObjStg declares more than 256 MiB");
            }
        }
        let length = u32::try_from(data_len)
            .map_err(|_| Error::Corrupted("ExOleObjStg payload exceeds u32".into()))?;
        let mut bytes = Vec::with_capacity(self.record_len());
        bytes.extend_from_slice(&(instance << 4).to_le_bytes());
        bytes.extend_from_slice(&RecordType::ExternalOleObjectStg.as_u16().to_le_bytes());
        bytes.extend_from_slice(&length.to_le_bytes());
        if let Some(declared) = declared {
            bytes.extend_from_slice(&declared.to_le_bytes());
        }
        bytes.extend_from_slice(self.stored_bytes());
        Ok(bytes)
    }

    /// Parse the encoded view as a complete PowerPoint record.
    pub fn to_record(self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
