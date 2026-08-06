//! OfficeArt record wire primitives and zero-copy inspection.
//!
//! The encoder uses the shared `litchi-odraw` header vocabulary, while this
//! module keeps the PPT writer's ergonomic builder and the exact source bytes
//! needed for lossless inspection of producer-specific records.

#![allow(dead_code)]

use std::collections::HashSet;
use std::io::{Error, ErrorKind};

use litchi_odraw::{Container, Parser, Record, RecordKind};

use super::super::{Error as EscherError, EscherHeader};

/// Escher record builder used by PPT record families.
#[derive(Debug, Clone)]
pub(crate) struct EscherBuilder {
    header: EscherHeader,
    data: Vec<u8>,
}

impl EscherBuilder {
    /// Creates an empty record with the supplied OfficeArt header fields.
    pub(crate) fn new(version: u8, instance: u16, record_type: u16) -> Self {
        Self {
            header: EscherHeader::new(version, instance, record_type, 0),
            data: Vec::new(),
        }
    }

    /// Appends a wire payload to the record.
    pub(crate) fn add_data(&mut self, data: &[u8]) {
        self.data.extend_from_slice(data);
        self.header.length = self.data.len() as u32;
    }

    /// Finalizes the complete record, including its eight-byte header.
    pub(crate) fn build(&self) -> Result<Vec<u8>, EscherError> {
        let mut record = Vec::with_capacity(8 + self.data.len());
        self.header.write(&mut record)?;
        record.extend_from_slice(&self.data);
        Ok(record)
    }
}

/// A borrowed unknown OfficeArt record.
///
/// `bytes` includes the eight-byte record header. Keeping the complete source
/// slice, instead of only its payload, means an extension record can be copied
/// back without normalizing its header or losing its original identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct UnknownRecord<'data> {
    raw_kind: u16,
    version: u8,
    instance: u16,
    bytes: &'data [u8],
}

impl<'data> UnknownRecord<'data> {
    /// Returns the exact unknown record type value from the wire.
    pub(crate) const fn raw_kind(self) -> u16 {
        self.raw_kind
    }

    /// Returns the record version nibble.
    pub(crate) const fn version(self) -> u8 {
        self.version
    }

    /// Returns the record instance, preserving extension identifiers.
    pub(crate) const fn instance(self) -> u16 {
        self.instance
    }

    /// Returns the complete record bytes without copying.
    pub(crate) const fn bytes(self) -> &'data [u8] {
        self.bytes
    }
}

/// A checked, zero-copy PPT drawing view.
///
/// The view intentionally borrows the input stream. Known shape topology is
/// validated by the format-neutral OfficeArt substrate; unknown records are
/// surfaced separately and remain byte-exact for a future snapshot editor.
#[derive(Debug, Clone)]
pub(crate) struct Drawing<'data> {
    data: &'data [u8],
    root: Container<'data>,
}

impl<'data> Drawing<'data> {
    /// Parses one complete OfficeArt container without copying its payloads.
    pub(crate) fn parse(data: &'data [u8]) -> Result<Self, Error> {
        let root = Parser::new(data)
            .root()
            .map_err(map_wire_error)?
            .ok_or_else(|| invalid_data("OfficeArt stream has no container root"))?;
        Ok(Self { data, root })
    }

    /// Returns the root record kind.
    pub(crate) const fn root_kind(&self) -> RecordKind {
        self.root.record().kind()
    }

    /// Returns the borrowed root container.
    pub(crate) const fn root(&self) -> &Container<'data> {
        &self.root
    }

    /// Returns the original complete stream without normalization.
    pub(crate) const fn bytes(&self) -> &'data [u8] {
        self.data
    }

    /// Validates the group/shape/container topology and shape atoms.
    pub(crate) fn validate_shapes(&self) -> Result<(), Error> {
        litchi_odraw::shape::parse(self.data)
            .map_err(map_wire_error)
            .map(|_| ())?;

        let ids = self.shape_ids()?;
        let mut seen = HashSet::with_capacity(ids.len());
        for id in ids {
            if id == 0 {
                return Err(invalid_data("OfficeArt shape identifier must be non-zero"));
            }
            if !seen.insert(id) {
                return Err(invalid_data("OfficeArt shape identifiers must be unique"));
            }
        }
        Ok(())
    }

    /// Returns every shape identifier in wire traversal order.
    pub(crate) fn shape_ids(&self) -> Result<Vec<u32>, Error> {
        let mut ids = Vec::new();
        collect_shape_ids(&self.root, &mut ids)?;
        Ok(ids)
    }

    /// Retains every unknown direct or nested record in source order.
    pub(crate) fn unknown_records(&self) -> Result<Vec<UnknownRecord<'data>>, Error> {
        let mut records = Vec::new();
        collect_unknown_records(&self.root, self.data, &mut records)?;
        Ok(records)
    }
}

fn collect_shape_ids(container: &Container<'_>, ids: &mut Vec<u32>) -> Result<(), Error> {
    for child in container.children() {
        let child = child.map_err(map_wire_error)?;
        if child.kind() == RecordKind::Sp {
            let payload = child.data();
            let id = payload
                .get(..4)
                .and_then(|bytes| bytes.try_into().ok())
                .map(u32::from_le_bytes)
                .ok_or_else(|| invalid_data("OfficeArt Sp atom has no shape identifier"))?;
            ids.push(id);
        }
        if child.kind().is_container() {
            let nested = Container::try_new(child).map_err(map_wire_error)?;
            collect_shape_ids(&nested, ids)?;
        }
    }
    Ok(())
}

fn collect_unknown_records<'data>(
    container: &Container<'data>,
    source: &'data [u8],
    records: &mut Vec<UnknownRecord<'data>>,
) -> Result<(), Error> {
    for child in container.children() {
        let child = child.map_err(map_wire_error)?;
        if matches!(child.kind(), RecordKind::Unknown(_)) {
            records.push(UnknownRecord {
                raw_kind: child.raw_kind(),
                version: child.version(),
                instance: child.instance(),
                bytes: raw_record_bytes(&child, source)?,
            });
            // Unknown container grammars are intentionally opaque. Retaining
            // the complete record avoids guessing at extension child framing.
            continue;
        }
        if child.kind().is_container() {
            let nested = Container::try_new(child).map_err(map_wire_error)?;
            collect_unknown_records(&nested, source, records)?;
        }
    }
    Ok(())
}

fn raw_record_bytes<'data>(
    record: &Record<'data>,
    source: &'data [u8],
) -> Result<&'data [u8], Error> {
    let body_start = record
        .data_offset(source)
        .ok_or_else(|| invalid_data("OfficeArt record is not backed by its source stream"))?;
    let start = body_start
        .checked_sub(8)
        .ok_or_else(|| invalid_data("OfficeArt record header offset underflow"))?;
    let end = start
        .checked_add(8)
        .and_then(|offset| offset.checked_add(record.data().len()))
        .ok_or_else(|| invalid_data("OfficeArt record extent overflows"))?;
    source
        .get(start..end)
        .ok_or_else(|| invalid_data("OfficeArt record extent exceeds its source stream"))
}

fn map_wire_error(error: impl std::fmt::Display) -> Error {
    invalid_data(error.to_string())
}

fn invalid_data(message: impl Into<String>) -> Error {
    Error::new(ErrorKind::InvalidData, message.into())
}
