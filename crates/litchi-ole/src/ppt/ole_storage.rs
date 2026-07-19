//! Bounded opaque PowerPoint OLE, VBA, and ActiveX persisted storage records.
//!
//! Payloads are never decompressed, parsed as compound files, activated, or executed.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

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
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0
            || record.record_type_raw != PptRecordType::ExternalOleObjectStg.as_u16()
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
            || record.data.len() > MAX_STORED_BYTES
        {
            return corrupted("ExOleObjStg has an invalid header or stored size");
        }
        let (kind, compressed) = match record.instance {
            0 => (PowerPointOleStorageKind::OleObject, false),
            1 => (PowerPointOleStorageKind::OleObject, true),
            2 => (PowerPointOleStorageKind::VbaProject, false),
            3 => (PowerPointOleStorageKind::VbaProject, true),
            4 => (PowerPointOleStorageKind::ActiveXControl, false),
            5 => (PowerPointOleStorageKind::ActiveXControl, true),
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

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let compressed = matches!(
            self.compression,
            PowerPointOleStorageCompression::Zlib { .. }
        );
        let instance: u16 = match (self.kind, compressed) {
            (PowerPointOleStorageKind::OleObject, false) => 0,
            (PowerPointOleStorageKind::OleObject, true) => 1,
            (PowerPointOleStorageKind::VbaProject, false) => 2,
            (PowerPointOleStorageKind::VbaProject, true) => 3,
            (PowerPointOleStorageKind::ActiveXControl, false) => 4,
            (PowerPointOleStorageKind::ActiveXControl, true) => 5,
        };
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
                let parsed = PowerPointOleStorage::parse(&expected.to_record().unwrap()).unwrap();
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
        bytes[0..2].copy_from_slice(&(6u16 << 4).to_le_bytes());
        assert!(PowerPointOleStorage::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
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
}
