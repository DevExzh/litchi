//! Shape-level Boolean flags stored in OfficeArt `ClientData`.
//!
//! Implements MS-PPT sections 2.7.3, 2.7.5, and 2.7.6. All retained
//! client-data records remain inert and are serialized byte-for-byte.

use crate::consts::PptRecordType;

use super::package::{PptError, Result};
use super::records::PptRecord;

const OFFICEART_CLIENT_DATA_TYPE: u16 = 0xf011;

/// Resource limits for shape-flag client-data projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointShapeFlagLimits {
    /// Maximum OfficeArt `ClientData` payload size.
    pub max_client_data_bytes: usize,
    /// Maximum number of direct PPT records in `ClientData`.
    pub max_client_data_records: usize,
    /// Maximum bytes retained after the optional shape-flag prefix.
    pub max_trailing_bytes: usize,
}

impl Default for PowerPointShapeFlagLimits {
    fn default() -> Self {
        Self {
            max_client_data_bytes: 4 * 1024 * 1024,
            max_client_data_records: 4096,
            max_trailing_bytes: 4 * 1024 * 1024,
        }
    }
}

/// MS-PPT 2.7.5 `ShapeFlagsAtom`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointShapeFlags {
    /// Whether the shape is rendered on top of other shapes.
    pub always_on_top: bool,
}

impl PowerPointShapeFlags {
    /// Parse a complete `RT_ShapeAtom` record.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        validate_atom_record(record, PptRecordType::ShapeAtom, "ShapeFlagsAtom")?;
        Self::parse_payload(&record.data)
    }

    /// Parse the one-byte `ShapeFlagsAtom` payload.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != 1 {
            return corrupted("ShapeFlagsAtom payload must be exactly one byte");
        }
        if data[0] & 0xfe != 0 {
            return corrupted("ShapeFlagsAtom reserved bits must be zero");
        }
        Ok(Self {
            always_on_top: data[0] & 0x01 != 0,
        })
    }

    /// Serialize the one-byte atom payload.
    pub fn to_payload(self) -> [u8; 1] {
        [u8::from(self.always_on_top)]
    }

    /// Build a generic PPT atom record.
    pub fn to_record(self) -> PptRecord {
        let data = self.to_payload().to_vec();
        PptRecord {
            record_type: PptRecordType::ShapeAtom,
            record_type_raw: PptRecordType::ShapeAtom.as_u16(),
            version: 0,
            instance: 0,
            data_length: 1,
            data,
            children: Vec::new(),
        }
    }

    /// Serialize a complete `ShapeFlagsAtom` record.
    pub fn to_bytes(self) -> Vec<u8> {
        encode_record(0, 0, PptRecordType::ShapeAtom.as_u16(), &self.to_payload())
    }
}

/// MS-PPT 2.7.6 `ShapeFlags10Atom`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointShapeFlags10 {
    /// Whether the shape is a picture in the presentation's photo album.
    pub is_photo_album_picture: bool,
}

impl PowerPointShapeFlags10 {
    /// Parse a complete `RT_ShapeFlags10Atom` record.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        validate_atom_record(record, PptRecordType::ShapeFlags10Atom, "ShapeFlags10Atom")?;
        Self::parse_payload(&record.data)
    }

    /// Parse the one-byte `ShapeFlags10Atom` payload.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != 1 {
            return corrupted("ShapeFlags10Atom payload must be exactly one byte");
        }
        if data[0] & !0x04 != 0 {
            return corrupted("ShapeFlags10Atom reserved bits must be zero");
        }
        Ok(Self {
            is_photo_album_picture: data[0] & 0x04 != 0,
        })
    }

    /// Serialize the one-byte atom payload.
    pub fn to_payload(self) -> [u8; 1] {
        [if self.is_photo_album_picture { 0x04 } else { 0 }]
    }

    /// Build a generic PPT atom record.
    pub fn to_record(self) -> PptRecord {
        let data = self.to_payload().to_vec();
        PptRecord {
            record_type: PptRecordType::ShapeFlags10Atom,
            record_type_raw: PptRecordType::ShapeFlags10Atom.as_u16(),
            version: 0,
            instance: 0,
            data_length: 1,
            data,
            children: Vec::new(),
        }
    }

    /// Serialize a complete `ShapeFlags10Atom` record.
    pub fn to_bytes(self) -> Vec<u8> {
        encode_record(
            0,
            0,
            PptRecordType::ShapeFlags10Atom.as_u16(),
            &self.to_payload(),
        )
    }
}

/// Typed optional flag prefix projected from one OfficeArt `ClientData` record.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointShapeFlagProjection {
    /// PowerPoint 97 shape flags.
    pub flags: Option<PowerPointShapeFlags>,
    /// PowerPoint 2002 shape flags.
    pub flags10: Option<PowerPointShapeFlags10>,
    trailing_records: Vec<Vec<u8>>,
}

/// Slide-level shape flag result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointShapeFlagEntry {
    /// OfficeArt shape identifier.
    pub shape_id: u32,
    /// Typed flag projection for the shape.
    pub projection: PowerPointShapeFlagProjection,
}

/// Presentation-level shape flag result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointPresentationShapeFlagEntry {
    /// One-based slide number.
    pub slide_number: usize,
    /// OfficeArt shape identifier.
    pub shape_id: u32,
    /// Typed flag projection for the shape.
    pub projection: PowerPointShapeFlagProjection,
}

impl PowerPointShapeFlagProjection {
    /// Parse a complete OfficeArt `ClientData` record, including its header.
    pub fn parse_officeart_client_data(
        data: &[u8],
        limits: PowerPointShapeFlagLimits,
    ) -> Result<Self> {
        if data.len() < 8 {
            return corrupted("Truncated OfficeArt ClientData record header");
        }
        let version_instance = u16::from_le_bytes([data[0], data[1]]);
        let record_type = u16::from_le_bytes([data[2], data[3]]);
        let length = usize::try_from(u32::from_le_bytes([data[4], data[5], data[6], data[7]]))
            .map_err(|_| PptError::Corrupted("OfficeArt ClientData size overflow".into()))?;
        if version_instance & 0x000f != 0x000f
            || version_instance >> 4 != 0
            || record_type != OFFICEART_CLIENT_DATA_TYPE
        {
            return corrupted("Invalid OfficeArt ClientData record header");
        }
        if length.checked_add(8) != Some(data.len()) {
            return corrupted("OfficeArt ClientData length does not match its payload");
        }
        Self::parse_client_data_payload(&data[8..], limits)
    }

    /// Parse the direct PPT-record sequence inside OfficeArt `ClientData`.
    pub fn parse_client_data_payload(
        data: &[u8],
        limits: PowerPointShapeFlagLimits,
    ) -> Result<Self> {
        check_limit(
            data.len(),
            limits.max_client_data_bytes,
            "OfficeArt ClientData payload",
        )?;
        let records = split_records(data, limits.max_client_data_records)?;
        let mut flags = None;
        let mut flags10 = None;
        let mut trailing_records = Vec::new();
        let mut trailing = false;

        for bytes in records {
            let record_type = u16::from_le_bytes([bytes[2], bytes[3]]);
            if record_type == PptRecordType::ShapeAtom.as_u16() {
                if trailing || flags.is_some() || flags10.is_some() {
                    return corrupted(
                        "ShapeFlagsAtom is duplicated or appears outside its ClientData slot",
                    );
                }
                let (record, consumed) = PptRecord::parse_strict(&bytes, 0)?;
                if consumed != bytes.len() {
                    return corrupted("ShapeFlagsAtom was only partially parsed");
                }
                flags = Some(PowerPointShapeFlags::parse(&record)?);
            } else if record_type == PptRecordType::ShapeFlags10Atom.as_u16() {
                if trailing || flags10.is_some() {
                    return corrupted(
                        "ShapeFlags10Atom is duplicated or appears outside its ClientData slot",
                    );
                }
                let (record, consumed) = PptRecord::parse_strict(&bytes, 0)?;
                if consumed != bytes.len() {
                    return corrupted("ShapeFlags10Atom was only partially parsed");
                }
                flags10 = Some(PowerPointShapeFlags10::parse(&record)?);
            } else {
                trailing = true;
                trailing_records.push(bytes);
            }
        }

        let trailing_bytes = trailing_records.iter().try_fold(0usize, |total, record| {
            total.checked_add(record.len()).ok_or_else(|| {
                PptError::Corrupted("OfficeArt ClientData trailing size overflow".into())
            })
        })?;
        check_limit(
            trailing_bytes,
            limits.max_trailing_bytes,
            "OfficeArt ClientData trailing records",
        )?;
        Ok(Self {
            flags,
            flags10,
            trailing_records,
        })
    }

    /// Whether either defined shape-flag atom is present.
    pub fn has_flags(&self) -> bool {
        self.flags.is_some() || self.flags10.is_some()
    }

    /// Raw later client-data records retained for lossless serialization.
    pub fn trailing_records(&self) -> &[Vec<u8>] {
        &self.trailing_records
    }

    /// Serialize the PPT-record payload of OfficeArt `ClientData`.
    pub fn to_client_data_payload(&self, limits: PowerPointShapeFlagLimits) -> Result<Vec<u8>> {
        let record_count = usize::from(self.flags.is_some())
            .checked_add(usize::from(self.flags10.is_some()))
            .and_then(|count| count.checked_add(self.trailing_records.len()))
            .ok_or_else(|| {
                PptError::Corrupted("OfficeArt ClientData record count overflow".into())
            })?;
        check_limit(
            record_count,
            limits.max_client_data_records,
            "OfficeArt ClientData record count",
        )?;
        let mut output = Vec::new();
        if let Some(flags) = self.flags {
            output.extend_from_slice(&flags.to_bytes());
        }
        if let Some(flags10) = self.flags10 {
            output.extend_from_slice(&flags10.to_bytes());
        }
        for record in &self.trailing_records {
            output.extend_from_slice(record);
        }
        check_limit(
            output.len(),
            limits.max_client_data_bytes,
            "OfficeArt ClientData payload",
        )?;
        // Reparse so malformed public mutations or retained-record construction
        // cannot bypass ordering, structural validation, or limits.
        Self::parse_client_data_payload(&output, limits)?;
        Ok(output)
    }

    /// Serialize a complete OfficeArt `ClientData` record.
    pub fn to_officeart_client_data(&self, limits: PowerPointShapeFlagLimits) -> Result<Vec<u8>> {
        let payload = self.to_client_data_payload(limits)?;
        Ok(encode_record(0x0f, 0, OFFICEART_CLIENT_DATA_TYPE, &payload))
    }
}

fn validate_atom_record(record: &PptRecord, kind: PptRecordType, name: &str) -> Result<()> {
    if record.record_type != kind
        || record.record_type_raw != kind.as_u16()
        || record.version != 0
        || record.instance != 0
        || record.data_length != 1
        || record.data.len() != 1
    {
        return corrupted(format!("Invalid {name} record header or length"));
    }
    Ok(())
}

fn split_records(data: &[u8], max_records: usize) -> Result<Vec<Vec<u8>>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        if records.len() >= max_records {
            return corrupted("OfficeArt ClientData exceeds its record-count limit");
        }
        let header_end = offset
            .checked_add(8)
            .ok_or_else(|| PptError::Corrupted("ClientData record header overflow".into()))?;
        if header_end > data.len() {
            return corrupted("Truncated PPT record header in OfficeArt ClientData");
        }
        let length = usize::try_from(u32::from_le_bytes([
            data[offset + 4],
            data[offset + 5],
            data[offset + 6],
            data[offset + 7],
        ]))
        .map_err(|_| PptError::Corrupted("ClientData record size overflow".into()))?;
        let end = header_end
            .checked_add(length)
            .ok_or_else(|| PptError::Corrupted("ClientData record end overflow".into()))?;
        if end > data.len() {
            return corrupted("PPT record extends beyond OfficeArt ClientData");
        }
        records.push(data[offset..end].to_vec());
        offset = end;
    }
    Ok(records)
}

fn encode_record(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(8usize.saturating_add(data.len()));
    output.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&(data.len() as u32).to_le_bytes());
    output.extend_from_slice(data);
    output
}

fn check_limit(actual: usize, limit: usize, field: &str) -> Result<()> {
    if actual > limit {
        corrupted(format!("{field} exceeds its configured limit"))
    } else {
        Ok(())
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        encode_record(version, instance, kind, payload)
    }

    #[test]
    fn parses_and_writes_both_typed_atoms() {
        let flags = PowerPointShapeFlags::parse_payload(&[1]).unwrap();
        assert!(flags.always_on_top);
        assert_eq!(flags.to_bytes(), record(0, 0, 0x0bdb, &[1]));
        assert_eq!(
            PowerPointShapeFlags::parse(&flags.to_record()).unwrap(),
            flags
        );

        let flags10 = PowerPointShapeFlags10::parse_payload(&[4]).unwrap();
        assert!(flags10.is_photo_album_picture);
        assert_eq!(flags10.to_bytes(), record(0, 0, 0x0bdc, &[4]));
        assert_eq!(
            PowerPointShapeFlags10::parse(&flags10.to_record()).unwrap(),
            flags10
        );
    }

    #[test]
    fn projects_owned_prefix_and_round_trips_client_data_exactly() {
        let mut payload = record(0, 0, 0x0bdb, &[1]);
        payload.extend_from_slice(&record(0, 0, 0x0bdc, &[4]));
        payload.extend_from_slice(&record(0, 0, 0x0bc1, &37u32.to_le_bytes()));
        let complete = record(0x0f, 0, OFFICEART_CLIENT_DATA_TYPE, &payload);

        let projection = PowerPointShapeFlagProjection::parse_officeart_client_data(
            &complete,
            PowerPointShapeFlagLimits::default(),
        )
        .unwrap();
        assert_eq!(projection.flags.unwrap().always_on_top, true);
        assert_eq!(projection.flags10.unwrap().is_photo_album_picture, true);
        assert_eq!(projection.trailing_records().len(), 1);
        assert_eq!(
            projection
                .to_officeart_client_data(PowerPointShapeFlagLimits::default())
                .unwrap(),
            complete
        );
    }

    #[test]
    fn rejects_reserved_bits_headers_duplicates_late_atoms_and_truncation() {
        assert!(PowerPointShapeFlags::parse_payload(&[2]).is_err());
        assert!(PowerPointShapeFlags10::parse_payload(&[1]).is_err());
        assert!(PowerPointShapeFlags::parse_payload(&[]).is_err());

        let limits = PowerPointShapeFlagLimits::default();
        let bad_version = record(1, 0, 0x0bdb, &[0]);
        assert!(
            PowerPointShapeFlagProjection::parse_client_data_payload(&bad_version, limits).is_err()
        );

        let duplicate = [record(0, 0, 0x0bdb, &[0]), record(0, 0, 0x0bdb, &[1])].concat();
        assert!(
            PowerPointShapeFlagProjection::parse_client_data_payload(&duplicate, limits).is_err()
        );

        let reversed = [record(0, 0, 0x0bdc, &[0]), record(0, 0, 0x0bdb, &[0])].concat();
        assert!(
            PowerPointShapeFlagProjection::parse_client_data_payload(&reversed, limits).is_err()
        );

        let late = [record(0, 0, 0x0bc1, &[0; 4]), record(0, 0, 0x0bdc, &[0])].concat();
        assert!(PowerPointShapeFlagProjection::parse_client_data_payload(&late, limits).is_err());

        let mut truncated = record(0, 0, 0x0bc1, &[0; 4]);
        truncated.pop();
        assert!(
            PowerPointShapeFlagProjection::parse_client_data_payload(&truncated, limits).is_err()
        );

        let bad_client_header = record(0, 0, OFFICEART_CLIENT_DATA_TYPE, &[]);
        assert!(
            PowerPointShapeFlagProjection::parse_officeart_client_data(&bad_client_header, limits)
                .is_err()
        );
    }

    #[test]
    fn enforces_all_projection_limits() {
        let trailing = record(0, 0, 0x0bc1, &[0; 4]);
        let defaults = PowerPointShapeFlagLimits::default();
        let cases = [
            PowerPointShapeFlagLimits {
                max_client_data_bytes: trailing.len() - 1,
                ..defaults
            },
            PowerPointShapeFlagLimits {
                max_client_data_records: 0,
                ..defaults
            },
            PowerPointShapeFlagLimits {
                max_trailing_bytes: trailing.len() - 1,
                ..defaults
            },
        ];
        for limits in cases {
            assert!(
                PowerPointShapeFlagProjection::parse_client_data_payload(&trailing, limits)
                    .is_err()
            );
        }
    }
}
