//! Inert VBA project metadata from MS-PPT 2.4.10 and 2.4.11.

use crate::consts::PptRecordType;

use super::ole_storage::{
    PowerPointOleStorage, PowerPointOleStorageCompression, PowerPointOleStorageKind,
};
use super::package::{PptError, Result};
use super::records::PptRecord;

/// Strictly validated metadata pointing at a persisted VBA project storage.
///
/// This type never opens, decompresses, interprets, or executes the referenced
/// project. The storage remains available only through the bounded opaque
/// storage API.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointVbaInfo {
    /// Persist-directory identifier of the `VbaProjectStg` record.
    pub persist_id_ref: u32,
    /// Whether the referenced VBA storage contains data.
    pub has_macros: bool,
    /// VBA runtime version. MS-PPT requires this to be `2`.
    pub runtime_version: u32,
}

/// Payload-free metadata for a PowerPoint `VbaProjectStg` persist object.
///
/// This describes the outer MS-PPT storage record only. It deliberately does
/// not expose, decompress, parse as CFB, inspect, or execute the embedded
/// MS-OVBA project payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointVbaProjectStorage {
    info: PowerPointVbaInfo,
    compression: Option<PowerPointOleStorageCompression>,
    stored_payload_len: Option<usize>,
}

impl PowerPointVbaProjectStorage {
    pub(crate) fn from_info_and_storage(
        info: PowerPointVbaInfo,
        storage: Option<&PowerPointOleStorage>,
    ) -> Result<Self> {
        info.validate()?;
        if info.persist_id_ref == 0 && storage.is_some() {
            return Err(PptError::Corrupted(
                "VBAInfoAtom has a null persist reference with storage metadata".to_string(),
            ));
        }
        if info.persist_id_ref != 0 && storage.is_none() {
            return Err(PptError::Corrupted(format!(
                "VBAInfoAtom persist ID {} has no storage record",
                info.persist_id_ref
            )));
        }
        if let Some(storage) = storage
            && storage.kind != PowerPointOleStorageKind::VbaProject
        {
            return Err(PptError::Corrupted(format!(
                "VBAInfoAtom persist ID {} does not reference VBA project storage",
                info.persist_id_ref
            )));
        }
        Ok(Self {
            info,
            compression: storage.map(|storage| storage.compression),
            stored_payload_len: storage.map(|storage| storage.data.len()),
        })
    }

    /// Return the `VBAInfoAtom` metadata that points at this storage.
    pub fn info(&self) -> PowerPointVbaInfo {
        self.info
    }

    /// Return the persisted VBA storage identifier.
    pub fn persist_id_ref(&self) -> u32 {
        self.info.persist_id_ref
    }

    /// Whether `VBAInfoAtom` declares that the project storage has data.
    pub fn has_macros(&self) -> bool {
        self.info.has_macros
    }

    /// Return the VBA runtime version recorded by PowerPoint.
    pub fn runtime_version(&self) -> u32 {
        self.info.runtime_version
    }

    /// Whether a non-null VBA project storage record is present.
    pub fn has_persisted_storage(&self) -> bool {
        self.compression.is_some()
    }

    /// Return outer-record compression metadata without decompressing data.
    pub fn compression(&self) -> Option<PowerPointOleStorageCompression> {
        self.compression
    }

    /// Return the stored opaque payload length, excluding a compressed size prefix.
    pub fn stored_payload_len(&self) -> Option<usize> {
        self.stored_payload_len
    }

    /// Return the declared decompressed payload length, if the storage is compressed.
    ///
    /// This is metadata from the outer record. The library never attempts the
    /// corresponding decompression.
    pub fn declared_uncompressed_len(&self) -> Option<u32> {
        match self.compression {
            Some(PowerPointOleStorageCompression::Zlib { uncompressed_len }) => {
                Some(uncompressed_len)
            },
            Some(PowerPointOleStorageCompression::Uncompressed) | None => None,
        }
    }

    /// Whether outer metadata declares a compressed VBA project storage.
    pub fn is_compressed(&self) -> bool {
        matches!(
            self.compression,
            Some(PowerPointOleStorageCompression::Zlib { .. })
        )
    }

    /// Whether persisted metadata conservatively indicates macro data.
    ///
    /// This does not inspect project, module, or source-code bytes.
    pub fn may_contain_macro_code(&self) -> bool {
        self.info.has_macros && self.has_persisted_storage()
    }
}

impl PowerPointVbaInfo {
    /// Parse one complete `VBAInfoContainer`.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::VBAInfo
            || record.version != 0x0f
            || record.instance != 1
            || record.data.len() != 20
        {
            return Err(PptError::Corrupted(
                "VBAInfoContainer has an invalid record header or size".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "VBAInfoContainer")?;
        if children.len() != 1 {
            return Err(PptError::Corrupted(
                "VBAInfoContainer must contain exactly one VBAInfoAtom".to_string(),
            ));
        }
        let atom = &children[0];
        if atom.record_type != PptRecordType::VBAInfoAtom
            || atom.version != 2
            || atom.instance != 0
            || atom.data.len() != 12
        {
            return Err(PptError::Corrupted(
                "VBAInfoAtom has an invalid record header or size".to_string(),
            ));
        }
        let persist_id_ref = u32::from_le_bytes(atom.data[0..4].try_into().unwrap());
        let has_macros = match u32::from_le_bytes(atom.data[4..8].try_into().unwrap()) {
            0 => false,
            1 => true,
            _ => {
                return Err(PptError::Corrupted(
                    "VBAInfoAtom has an invalid fHasMacros value".to_string(),
                ));
            },
        };
        let runtime_version = u32::from_le_bytes(atom.data[8..12].try_into().unwrap());
        let result = Self {
            persist_id_ref,
            has_macros,
            runtime_version,
        };
        result.validate()?;
        Ok(result)
    }

    /// Discover the single document-level VBA metadata container, if present.
    pub(crate) fn parse_records(records: &[&PptRecord]) -> Result<Option<Self>> {
        let mut matches = records
            .iter()
            .copied()
            .filter(|record| record.record_type == PptRecordType::VBAInfo);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(PptError::Corrupted(
                "Record tree contains multiple VBAInfoContainer records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }

    /// Encode the exact container and atom headers required by MS-PPT.
    pub fn to_record(self) -> Result<PptRecord> {
        self.validate()?;
        let mut atom_data = Vec::with_capacity(12);
        atom_data.extend_from_slice(&self.persist_id_ref.to_le_bytes());
        atom_data.extend_from_slice(&u32::from(self.has_macros).to_le_bytes());
        atom_data.extend_from_slice(&self.runtime_version.to_le_bytes());
        let atom = PptRecord {
            record_type: PptRecordType::VBAInfoAtom,
            record_type_raw: 1024,
            version: 2,
            instance: 0,
            data_length: 12,
            data: atom_data,
            children: Vec::new(),
        };
        let mut data = Vec::with_capacity(20);
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&1024u16.to_le_bytes());
        data.extend_from_slice(&12u32.to_le_bytes());
        data.extend_from_slice(&atom.data);
        Ok(PptRecord {
            record_type: PptRecordType::VBAInfo,
            record_type_raw: 1023,
            version: 0x0f,
            instance: 1,
            data_length: 20,
            data,
            children: vec![atom],
        })
    }

    fn validate(self) -> Result<()> {
        if self.runtime_version != 2 {
            return Err(PptError::Corrupted(
                "VBAInfoAtom has an invalid runtime version".to_string(),
            ));
        }
        if self.has_macros && self.persist_id_ref == 0 {
            return Err(PptError::Corrupted(
                "VBAInfoAtom declares macro data without a persist reference".to_string(),
            ));
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_inert_vba_metadata() {
        let expected = PowerPointVbaInfo {
            persist_id_ref: 41,
            has_macros: true,
            runtime_version: 2,
        };
        let record = expected.to_record().unwrap();
        assert_eq!(record.version, 0x0f);
        assert_eq!(record.instance, 1);
        assert_eq!(record.data.len(), 20);
        assert_eq!(PowerPointVbaInfo::parse(&record).unwrap(), expected);
    }

    #[test]
    fn rejects_invalid_flags_versions_and_missing_persist_reference() {
        let mut record = PowerPointVbaInfo {
            persist_id_ref: 41,
            has_macros: true,
            runtime_version: 2,
        }
        .to_record()
        .unwrap();
        record.data[12..16].copy_from_slice(&2u32.to_le_bytes());
        assert!(PowerPointVbaInfo::parse(&record).is_err());

        record.data[12..16].copy_from_slice(&1u32.to_le_bytes());
        record.data[16..20].copy_from_slice(&3u32.to_le_bytes());
        assert!(PowerPointVbaInfo::parse(&record).is_err());

        assert!(
            PowerPointVbaInfo {
                persist_id_ref: 0,
                has_macros: true,
                runtime_version: 2,
            }
            .to_record()
            .is_err()
        );
    }

    #[test]
    fn rejects_duplicate_document_metadata() {
        let record = PowerPointVbaInfo {
            persist_id_ref: 0,
            has_macros: false,
            runtime_version: 2,
        }
        .to_record()
        .unwrap();
        assert!(PowerPointVbaInfo::parse_records(&[&record, &record]).is_err());
    }

    #[test]
    fn summarizes_vba_storage_without_exposing_or_decompressing_payload_data() {
        let info = PowerPointVbaInfo {
            persist_id_ref: 41,
            has_macros: true,
            runtime_version: 2,
        };
        let storage = PowerPointOleStorage {
            kind: PowerPointOleStorageKind::VbaProject,
            compression: PowerPointOleStorageCompression::Zlib {
                uncompressed_len: 4096,
            },
            data: vec![0x78, 0x9c, 1, 2, 3],
        };
        let summary =
            PowerPointVbaProjectStorage::from_info_and_storage(info, Some(&storage)).unwrap();

        assert_eq!(summary.info(), info);
        assert_eq!(summary.persist_id_ref(), 41);
        assert!(summary.has_macros());
        assert!(summary.has_persisted_storage());
        assert!(summary.is_compressed());
        assert_eq!(summary.stored_payload_len(), Some(5));
        assert_eq!(summary.declared_uncompressed_len(), Some(4096));
        assert!(summary.may_contain_macro_code());
    }

    #[test]
    fn rejects_missing_or_wrong_vba_storage_without_touching_payload_contents() {
        let info = PowerPointVbaInfo {
            persist_id_ref: 41,
            has_macros: false,
            runtime_version: 2,
        };
        assert!(PowerPointVbaProjectStorage::from_info_and_storage(info, None).is_err());

        let wrong_storage = PowerPointOleStorage {
            kind: PowerPointOleStorageKind::OleObject,
            compression: PowerPointOleStorageCompression::Uncompressed,
            data: vec![0x01],
        };
        assert!(
            PowerPointVbaProjectStorage::from_info_and_storage(info, Some(&wrong_storage)).is_err()
        );

        let empty = PowerPointVbaInfo {
            persist_id_ref: 0,
            has_macros: false,
            runtime_version: 2,
        };
        let summary = PowerPointVbaProjectStorage::from_info_and_storage(empty, None).unwrap();
        assert!(!summary.has_persisted_storage());
        assert!(!summary.may_contain_macro_code());
    }
}
