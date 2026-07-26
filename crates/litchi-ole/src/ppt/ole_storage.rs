//! Bounded PowerPoint OLE, VBA, and ActiveX persisted storage records.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;
use std::io::Read;

const MAX_STORED_BYTES: usize = 128 * 1_048_576;
const MAX_DECLARED_BYTES: u32 = 256 * 1_048_576;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointOleStorageKind {
    OleObject,
    VbaProject,
    ActiveXControl,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointOleStorageCompression {
    Uncompressed,
    Zlib { uncompressed_len: u32 },
}

/// One persisted storage payload with byte-exact inert contents.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointOleStorage {
    pub kind: PowerPointOleStorageKind,
    pub compression: PowerPointOleStorageCompression,
    /// Raw uncompressed bytes, or raw zlib bytes without the four-byte size prefix.
    pub data: Vec<u8>,
}

impl PowerPointOleStorage {
    /// Parse an embedded OLE-object storage.
    ///
    /// MS-PPT uses the referencing record—not `recInstance`—to distinguish
    /// OLE objects, VBA projects, and ActiveX controls. Call [`Self::parse_as`]
    /// when the reference context identifies another kind.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        Self::parse_as(record, PowerPointOleStorageKind::OleObject)
    }

    /// Parse a persisted storage using its reference-context kind.
    pub fn parse_as(record: &PptRecord, kind: PowerPointOleStorageKind) -> Result<Self> {
        if record.version != 0
            || record.record_type_raw != PptRecordType::ExternalOleObjectStg.as_u16()
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
            || record.data.len() > MAX_STORED_BYTES
        {
            return corrupted("ExOleObjStg has an invalid header or stored size");
        }
        let compressed = match record.instance {
            0 => false,
            1 => true,
            _ => return corrupted("ExOleObjStg has an invalid storage instance"),
        };
        if compressed {
            if record.data.len() < 4 {
                return corrupted("compressed ExOleObjStg is missing its size prefix");
            }
            let uncompressed_len =
                u32::from_le_bytes(record.data[..4].try_into().expect("fixed slice"));
            if uncompressed_len > MAX_DECLARED_BYTES {
                return corrupted("compressed ExOleObjStg declares more than 256 MiB");
            }
            Ok(Self {
                kind,
                compression: PowerPointOleStorageCompression::Zlib { uncompressed_len },
                data: record.data[4..].to_vec(),
            })
        } else {
            Ok(Self {
                kind,
                compression: PowerPointOleStorageCompression::Uncompressed,
                data: record.data.clone(),
            })
        }
    }

    /// Return the uncompressed structured-storage bytes with an explicit cap.
    ///
    /// Zlib streams must consume the complete stored payload and produce
    /// exactly the declared byte count. The returned CFB bytes are inert.
    pub fn decompressed_bytes(&self, maximum: usize) -> Result<Vec<u8>> {
        match self.compression {
            PowerPointOleStorageCompression::Uncompressed => {
                if self.data.len() > maximum {
                    return corrupted(format!(
                        "uncompressed ExOleObjStg exceeds the {maximum}-byte output limit"
                    ));
                }
                Ok(self.data.clone())
            },
            PowerPointOleStorageCompression::Zlib { uncompressed_len } => {
                let declared = usize::try_from(uncompressed_len).map_err(|_| {
                    PptError::Corrupted(
                        "compressed ExOleObjStg size does not fit in memory".to_string(),
                    )
                })?;
                if declared > maximum {
                    return corrupted(format!(
                        "compressed ExOleObjStg declares {declared} bytes above the {maximum}-byte output limit"
                    ));
                }
                let read_limit = maximum.checked_add(1).ok_or_else(|| {
                    PptError::Corrupted("ExOleObjStg output limit overflows".to_string())
                })?;
                let decoder = flate2::read::ZlibDecoder::new(self.data.as_slice());
                let mut limited = decoder.take(u64::try_from(read_limit).unwrap_or(u64::MAX));
                let mut output = Vec::with_capacity(declared.min(64 * 1024));
                limited.read_to_end(&mut output).map_err(|error| {
                    PptError::Corrupted(format!("invalid ExOleObjStg zlib payload: {error}"))
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
                if decoder.total_in() != self.data.len() as u64 {
                    return corrupted("ExOleObjStg zlib payload has trailing bytes");
                }
                Ok(output)
            },
        }
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let compressed = matches!(
            self.compression,
            PowerPointOleStorageCompression::Zlib { .. }
        );
        let instance = u16::from(compressed);
        let prefix_len = if compressed { 4 } else { 0 };
        if self.data.len().saturating_add(prefix_len) > MAX_STORED_BYTES {
            return corrupted("ExOleObjStg exceeds 128 MiB stored data");
        }
        let mut data = Vec::with_capacity(self.data.len().saturating_add(prefix_len));
        if let PowerPointOleStorageCompression::Zlib { uncompressed_len } = self.compression {
            if uncompressed_len > MAX_DECLARED_BYTES {
                return corrupted("compressed ExOleObjStg declares more than 256 MiB");
            }
            data.extend_from_slice(&uncompressed_len.to_le_bytes());
        }
        data.extend_from_slice(&self.data);
        let length = u32::try_from(data.len())
            .map_err(|_| PptError::Corrupted("ExOleObjStg payload exceeds u32".into()))?;
        let mut bytes = Vec::with_capacity(data.len().saturating_add(8));
        bytes.extend_from_slice(&(instance << 4).to_le_bytes());
        bytes.extend_from_slice(&PptRecordType::ExternalOleObjectStg.as_u16().to_le_bytes());
        bytes.extend_from_slice(&length.to_le_bytes());
        bytes.extend_from_slice(&data);
        Ok(bytes)
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_storage_kind_and_compression_roundtrips_exactly() {
        for kind in [
            PowerPointOleStorageKind::OleObject,
            PowerPointOleStorageKind::VbaProject,
            PowerPointOleStorageKind::ActiveXControl,
        ] {
            for compression in [
                PowerPointOleStorageCompression::Uncompressed,
                PowerPointOleStorageCompression::Zlib {
                    uncompressed_len: 4096,
                },
            ] {
                let expected = PowerPointOleStorage {
                    kind,
                    compression,
                    data: vec![0x78, 0x9c, 1, 2, 3, 4],
                };
                let record = expected.to_record().unwrap();
                assert_eq!(
                    record.instance,
                    u16::from(matches!(
                        compression,
                        PowerPointOleStorageCompression::Zlib { .. }
                    ))
                );
                let parsed = PowerPointOleStorage::parse_as(&record, kind).unwrap();
                assert_eq!(parsed, expected);
                assert_eq!(
                    parsed.to_record_bytes().unwrap(),
                    expected.to_record_bytes().unwrap()
                );
            }
        }
    }

    #[test]
    fn storage_rejects_invalid_instance_and_truncated_compressed_header() {
        let value = PowerPointOleStorage {
            kind: PowerPointOleStorageKind::OleObject,
            compression: PowerPointOleStorageCompression::Uncompressed,
            data: Vec::new(),
        };
        let mut bytes = value.to_record_bytes().unwrap();
        bytes[0..2].copy_from_slice(&(2u16 << 4).to_le_bytes());
        assert!(PowerPointOleStorage::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
        bytes.truncate(8);
        bytes[0..2].copy_from_slice(&(1u16 << 4).to_le_bytes());
        assert!(PowerPointOleStorage::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    #[test]
    fn storage_rejects_declared_decompression_bomb_size() {
        let value = PowerPointOleStorage {
            kind: PowerPointOleStorageKind::VbaProject,
            compression: PowerPointOleStorageCompression::Zlib {
                uncompressed_len: MAX_DECLARED_BYTES + 1,
            },
            data: vec![0x78, 0x9c],
        };
        assert!(value.to_record_bytes().is_err());
    }

    #[test]
    fn bounded_zlib_decompression_requires_exact_size_and_no_trailing_data() {
        use std::io::Write;

        let original = b"compound storage bytes".repeat(100);
        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder.write_all(&original).unwrap();
        let compressed = encoder.finish().unwrap();
        let storage = PowerPointOleStorage {
            kind: PowerPointOleStorageKind::VbaProject,
            compression: PowerPointOleStorageCompression::Zlib {
                uncompressed_len: original.len() as u32,
            },
            data: compressed.clone(),
        };
        assert_eq!(
            storage.decompressed_bytes(original.len()).unwrap(),
            original
        );
        assert!(storage.decompressed_bytes(original.len() - 1).is_err());

        let mut wrong_size = storage.clone();
        wrong_size.compression = PowerPointOleStorageCompression::Zlib {
            uncompressed_len: original.len() as u32 + 1,
        };
        assert!(wrong_size.decompressed_bytes(original.len() + 1).is_err());

        let mut trailing = storage;
        trailing.data.push(0);
        assert!(trailing.decompressed_bytes(original.len()).is_err());
    }
}
