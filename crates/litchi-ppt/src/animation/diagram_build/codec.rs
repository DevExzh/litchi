//! Wire parsing and snapshot serialization for diagram-build records.

use super::model::{Atom, Build, BuildType, Container, Kind, Limits};
use super::validation::{
    ATOM_PAYLOAD_LEN, BUILD_PAYLOAD_LEN, CONTAINER_PAYLOAD_LEN, HEADER_LEN, encode_header,
    encode_record, parse_bool, validate_atom, validate_container,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

impl Build {
    pub const RECORD_LEN: usize = HEADER_LEN + BUILD_PAYLOAD_LEN;

    /// Parse a shared `BuildAtom` child.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_record(record: &Record) -> Result<Self> {
        validate_atom(
            record,
            RecordType::BuildAtom,
            BUILD_PAYLOAD_LEN,
            "BuildAtom",
        )?;
        let kind = Kind::from_raw(read_u32(&record.data, 0));
        Ok(Self::from_parts(
            kind,
            read_u32(&record.data, 4),
            read_u32(&record.data, 8),
            parse_bool(record.data[12], "BuildAtom.fExpanded")?,
            parse_bool(record.data[13], "BuildAtom.fUIExpanded")?,
            [record.data[14], record.data[15]],
        ))
    }

    /// Parse one exact serialized `BuildAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != Self::RECORD_LEN {
            return Err(Error::Corrupted(format!(
                "BuildAtom record must be exactly {} bytes, got {}",
                Self::RECORD_LEN,
                bytes.len()
            )));
        }
        let (record, consumed) = Record::parse_strict(bytes, 0)?;
        if consumed != bytes.len() {
            return Err(Error::Corrupted("BuildAtom has trailing bytes".to_string()));
        }
        Self::parse_record(&record)
    }

    /// Serialize the fixed `BuildAtom` payload.
    #[must_use]
    pub const fn to_payload(self) -> [u8; BUILD_PAYLOAD_LEN] {
        let kind = self.kind().raw().to_le_bytes();
        let build_id = self.build_id.to_le_bytes();
        let shape_id_ref = self.shape_id_ref.to_le_bytes();
        [
            kind[0],
            kind[1],
            kind[2],
            kind[3],
            build_id[0],
            build_id[1],
            build_id[2],
            build_id[3],
            shape_id_ref[0],
            shape_id_ref[1],
            shape_id_ref[2],
            shape_id_ref[3],
            self.expanded as u8,
            self.ui_expanded as u8,
            self.reserved()[0],
            self.reserved()[1],
        ]
    }

    /// Convert to the generic PPT record representation.
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "`BUILD_PAYLOAD_LEN` is a compile-time 16-byte constant, so it always fits in `u32`"
    )]
    pub fn to_record(self) -> Record {
        let data = self.to_payload().to_vec();
        Record {
            record_type: RecordType::BuildAtom,
            record_type_raw: RecordType::BuildAtom.as_u16(),
            version: 0,
            instance: 0,
            data_length: BUILD_PAYLOAD_LEN as u32,
            data,
            children: Vec::new(),
        }
    }

    /// Serialize the complete `BuildAtom` record.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(Self::RECORD_LEN);
        bytes.extend_from_slice(&encode_header(
            0,
            0,
            RecordType::BuildAtom,
            BUILD_PAYLOAD_LEN,
        ));
        bytes.extend_from_slice(&self.to_payload());
        bytes
    }
}

impl Atom {
    pub const RECORD_LEN: usize = HEADER_LEN + ATOM_PAYLOAD_LEN;

    /// Parse a `DiagramBuildAtom` child.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_record(record: &Record) -> Result<Self> {
        validate_atom(
            record,
            RecordType::DiagramBuildAtom,
            ATOM_PAYLOAD_LEN,
            "DiagramBuildAtom",
        )?;
        Ok(Self::new(BuildType::from_raw(read_u32(&record.data, 0))))
    }

    /// Parse one exact serialized `DiagramBuildAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != Self::RECORD_LEN {
            return Err(Error::Corrupted(format!(
                "DiagramBuildAtom record must be exactly {} bytes, got {}",
                Self::RECORD_LEN,
                bytes.len()
            )));
        }
        let (record, consumed) = Record::parse_strict(bytes, 0)?;
        if consumed != bytes.len() {
            return Err(Error::Corrupted(
                "DiagramBuildAtom has trailing bytes".to_string(),
            ));
        }
        Self::parse_record(&record)
    }

    /// Serialize the fixed `DiagramBuildAtom` payload.
    #[must_use]
    pub const fn to_payload(self) -> [u8; ATOM_PAYLOAD_LEN] {
        self.build_type.raw().to_le_bytes()
    }

    /// Convert to the generic PPT record representation.
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "`ATOM_PAYLOAD_LEN` is a compile-time 4-byte constant, so it always fits in `u32`"
    )]
    pub fn to_record(self) -> Record {
        let data = self.to_payload().to_vec();
        Record {
            record_type: RecordType::DiagramBuildAtom,
            record_type_raw: RecordType::DiagramBuildAtom.as_u16(),
            version: 0,
            instance: 0,
            data_length: ATOM_PAYLOAD_LEN as u32,
            data,
            children: Vec::new(),
        }
    }

    /// Serialize the complete `DiagramBuildAtom` record.
    #[must_use]
    pub fn to_bytes(self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(Self::RECORD_LEN);
        bytes.extend_from_slice(&encode_header(
            0,
            0,
            RecordType::DiagramBuildAtom,
            ATOM_PAYLOAD_LEN,
        ));
        bytes.extend_from_slice(&self.to_payload());
        bytes
    }
}

impl Container {
    pub const RECORD_LEN: usize = HEADER_LEN + CONTAINER_PAYLOAD_LEN;

    /// Parse one typed diagram-build container from a generic record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_record(record: &Record) -> Result<Self> {
        validate_container(record)?;
        let build = Build::parse_record(&record.children[0])?;
        let atom = Atom::parse_record(&record.children[1])?;
        Self::new(build, atom)
    }

    /// Parse one exact serialized diagram-build container.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse one exact container subject to an explicit allocation/size bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        if bytes.len() > limits.max_record_bytes {
            return Err(Error::InvalidFormat(
                "DiagramBuild exceeds the configured record-size limit".to_string(),
            ));
        }
        if bytes.len() != Self::RECORD_LEN {
            return Err(Error::Corrupted(format!(
                "DiagramBuild record must be exactly {} bytes, got {}",
                Self::RECORD_LEN,
                bytes.len()
            )));
        }
        let (record, consumed) = Record::parse_strict(bytes, 0)?;
        if consumed != bytes.len() {
            return Err(Error::Corrupted(
                "DiagramBuild has trailing bytes".to_string(),
            ));
        }
        Self::parse_record(&record)
    }

    /// Serialize the complete container, including both fixed child records.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let build = self.build().to_bytes();
        let atom = self.atom().to_bytes();
        let mut bytes = Vec::with_capacity(Self::RECORD_LEN);
        bytes.extend_from_slice(&encode_header(
            0x0F,
            0,
            RecordType::DiagramBuild,
            CONTAINER_PAYLOAD_LEN,
        ));
        bytes.extend_from_slice(&build);
        bytes.extend_from_slice(&atom);
        bytes
    }

    /// Convert to the generic PPT record representation.
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "`CONTAINER_PAYLOAD_LEN` is a compile-time 36-byte constant, so it always fits in `u32`"
    )]
    pub fn to_record(self) -> Record {
        let build = self.build().to_record();
        let atom = self.atom().to_record();
        let mut data = Vec::with_capacity(CONTAINER_PAYLOAD_LEN);
        data.extend_from_slice(&encode_record(&build));
        data.extend_from_slice(&encode_record(&atom));
        Record {
            record_type: RecordType::DiagramBuild,
            record_type_raw: RecordType::DiagramBuild.as_u16(),
            version: 0x0F,
            instance: 0,
            data_length: CONTAINER_PAYLOAD_LEN as u32,
            data,
            children: vec![build, atom],
        }
    }
}

/// Parse the typed diagram-build container from a generic PPT record.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_record(record: &Record) -> Result<Container> {
    Container::parse_record(record)
}

/// Parse one exact serialized diagram-build container.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_bytes(bytes: &[u8]) -> Result<Container> {
    Container::parse_bytes(bytes)
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}
