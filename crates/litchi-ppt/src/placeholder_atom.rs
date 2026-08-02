//! Strict MS-PPT 2.7.8 `PlaceholderAtom` support.
//!
//! Placeholder metadata is inert. Undefined payload bytes and unrelated
//! OfficeArt client-data records are retained for byte-exact serialization.

use crate::consts::PptRecordType;

use super::package::{PptError, Result};
use super::records::PptRecord;

const OFFICEART_CLIENT_DATA_TYPE: u16 = 0xf011;

/// Slide/master owner used to validate `PlaceholderEnum` constraints.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointPlaceholderContext {
    /// A main master slide.
    MainMaster,
    /// A title master slide.
    TitleMaster,
    /// A notes master slide.
    NotesMaster,
    /// A handout master slide.
    HandoutMaster,
    /// A normal presentation slide.
    PresentationSlide,
    /// A notes slide.
    NotesSlide,
}

/// Exact nonzero `PlaceholderEnum` values accepted by `PlaceholderAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum PowerPointPlaceholderKind {
    MasterTitle = 0x01,
    MasterBody = 0x02,
    MasterCenterTitle = 0x03,
    MasterSubTitle = 0x04,
    MasterNotesSlideImage = 0x05,
    MasterNotesBody = 0x06,
    MasterDate = 0x07,
    MasterSlideNumber = 0x08,
    MasterFooter = 0x09,
    MasterHeader = 0x0a,
    NotesSlideImage = 0x0b,
    NotesBody = 0x0c,
    Title = 0x0d,
    Body = 0x0e,
    CenterTitle = 0x0f,
    SubTitle = 0x10,
    VerticalTitle = 0x11,
    VerticalBody = 0x12,
    Object = 0x13,
    Graph = 0x14,
    Table = 0x15,
    ClipArt = 0x16,
    OrgChart = 0x17,
    Media = 0x18,
    VerticalObject = 0x19,
    Picture = 0x1a,
}

impl TryFrom<u8> for PowerPointPlaceholderKind {
    type Error = PptError;

    fn try_from(value: u8) -> Result<Self> {
        Ok(match value {
            0x01 => Self::MasterTitle,
            0x02 => Self::MasterBody,
            0x03 => Self::MasterCenterTitle,
            0x04 => Self::MasterSubTitle,
            0x05 => Self::MasterNotesSlideImage,
            0x06 => Self::MasterNotesBody,
            0x07 => Self::MasterDate,
            0x08 => Self::MasterSlideNumber,
            0x09 => Self::MasterFooter,
            0x0a => Self::MasterHeader,
            0x0b => Self::NotesSlideImage,
            0x0c => Self::NotesBody,
            0x0d => Self::Title,
            0x0e => Self::Body,
            0x0f => Self::CenterTitle,
            0x10 => Self::SubTitle,
            0x11 => Self::VerticalTitle,
            0x12 => Self::VerticalBody,
            0x13 => Self::Object,
            0x14 => Self::Graph,
            0x15 => Self::Table,
            0x16 => Self::ClipArt,
            0x17 => Self::OrgChart,
            0x18 => Self::Media,
            0x19 => Self::VerticalObject,
            0x1a => Self::Picture,
            0 => return corrupted("PT_None MUST NOT be used in PlaceholderAtom"),
            _ => return corrupted("PlaceholderAtom has an unknown PlaceholderEnum value"),
        })
    }
}

/// Exact `PlaceholderSize` discriminants.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum PowerPointPlaceholderSize {
    Full = 0,
    Half = 1,
    Quarter = 2,
}

impl TryFrom<u8> for PowerPointPlaceholderSize {
    type Error = PptError;

    fn try_from(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::Full),
            1 => Ok(Self::Half),
            2 => Ok(Self::Quarter),
            _ => corrupted("PlaceholderAtom has an unknown PlaceholderSize value"),
        }
    }
}

/// Typed `PlaceholderAtom` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointPlaceholderAtom {
    /// Placeholder position identifier; `-1` means this is not a placeholder.
    pub position: i32,
    /// Exact placeholder kind.
    pub kind: PowerPointPlaceholderKind,
    /// Preferred placeholder size.
    pub size: PowerPointPlaceholderSize,
    /// Spec-undefined bytes retained for lossless round trips.
    pub unused: [u8; 2],
}

impl PowerPointPlaceholderAtom {
    /// Parse a complete `RT_PlaceholderAtom` record for its owning slide context.
    pub fn parse(record: &PptRecord, context: PowerPointPlaceholderContext) -> Result<Self> {
        if record.record_type != PptRecordType::OEPlaceholderAtom
            || record.record_type_raw != PptRecordType::OEPlaceholderAtom.as_u16()
            || record.version != 0
            || record.instance != 0
            || record.data_length != 8
            || record.data.len() != 8
        {
            return corrupted("Invalid PlaceholderAtom record header or length");
        }
        Self::parse_payload(&record.data, context)
    }

    /// Parse the exact eight-byte payload for its owning slide context.
    pub fn parse_payload(data: &[u8], context: PowerPointPlaceholderContext) -> Result<Self> {
        if data.len() != 8 {
            return corrupted("PlaceholderAtom payload must be exactly eight bytes");
        }
        let kind = PowerPointPlaceholderKind::try_from(data[4])?;
        validate_context(kind, context)?;
        Ok(Self {
            position: i32::from_le_bytes([data[0], data[1], data[2], data[3]]),
            kind,
            size: PowerPointPlaceholderSize::try_from(data[5])?,
            unused: [data[6], data[7]],
        })
    }

    /// Serialize the exact eight-byte payload after revalidating its context.
    pub fn to_payload(self, context: PowerPointPlaceholderContext) -> Result<[u8; 8]> {
        validate_context(self.kind, context)?;
        let mut data = [0u8; 8];
        data[..4].copy_from_slice(&self.position.to_le_bytes());
        data[4] = self.kind as u8;
        data[5] = self.size as u8;
        data[6..].copy_from_slice(&self.unused);
        Ok(data)
    }

    /// Build a generic PPT record.
    pub fn to_record(self, context: PowerPointPlaceholderContext) -> Result<PptRecord> {
        Ok(PptRecord {
            record_type: PptRecordType::OEPlaceholderAtom,
            record_type_raw: PptRecordType::OEPlaceholderAtom.as_u16(),
            version: 0,
            instance: 0,
            data_length: 8,
            data: self.to_payload(context)?.to_vec(),
            children: Vec::new(),
        })
    }

    /// Serialize a complete PPT atom record.
    pub fn to_bytes(self, context: PowerPointPlaceholderContext) -> Result<Vec<u8>> {
        Ok(encode_record(
            0,
            0,
            PptRecordType::OEPlaceholderAtom.as_u16(),
            &self.to_payload(context)?,
        ))
    }
}

/// Resource limits for client-data placeholder projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointPlaceholderLimits {
    pub max_client_data_bytes: usize,
    pub max_client_data_records: usize,
    pub max_retained_bytes: usize,
}

impl Default for PowerPointPlaceholderLimits {
    fn default() -> Self {
        Self {
            max_client_data_bytes: 4 * 1024 * 1024,
            max_client_data_records: 4096,
            max_retained_bytes: 4 * 1024 * 1024,
        }
    }
}

/// Context-validated placeholder projected from OfficeArt `ClientData`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointPlaceholderProjection {
    pub placeholder: Option<PowerPointPlaceholderAtom>,
    before_records: Vec<Vec<u8>>,
    after_records: Vec<Vec<u8>>,
}

/// Slide-level placeholder result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointPlaceholderEntry {
    pub shape_id: u32,
    pub placeholder: PowerPointPlaceholderAtom,
}

/// Presentation-level placeholder result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointPresentationPlaceholderEntry {
    pub slide_number: usize,
    pub shape_id: u32,
    pub placeholder: PowerPointPlaceholderAtom,
}

impl PowerPointPlaceholderProjection {
    /// Parse a complete OfficeArt `ClientData` record.
    pub fn parse_officeart_client_data(
        data: &[u8],
        context: PowerPointPlaceholderContext,
        limits: PowerPointPlaceholderLimits,
    ) -> Result<Self> {
        if data.len() < 8 {
            return corrupted("Truncated OfficeArt ClientData record header");
        }
        let version_instance = u16::from_le_bytes([data[0], data[1]]);
        let record_type = u16::from_le_bytes([data[2], data[3]]);
        let length = usize::try_from(u32::from_le_bytes([data[4], data[5], data[6], data[7]]))
            .map_err(|_| PptError::Corrupted("OfficeArt ClientData size overflow".into()))?;
        if version_instance != 0x000f || record_type != OFFICEART_CLIENT_DATA_TYPE {
            return corrupted("Invalid OfficeArt ClientData record header");
        }
        if length.checked_add(8) != Some(data.len()) {
            return corrupted("OfficeArt ClientData length does not match its payload");
        }
        Self::parse_client_data_payload(&data[8..], context, limits)
    }

    /// Parse the direct PPT-record sequence inside OfficeArt `ClientData`.
    pub fn parse_client_data_payload(
        data: &[u8],
        context: PowerPointPlaceholderContext,
        limits: PowerPointPlaceholderLimits,
    ) -> Result<Self> {
        check_limit(
            data.len(),
            limits.max_client_data_bytes,
            "ClientData payload",
        )?;
        let records = split_records(data, limits.max_client_data_records)?;
        let mut placeholder = None;
        let mut before_records = Vec::new();
        let mut after_records = Vec::new();
        let mut after_placeholder = false;
        let mut saw_later_slot = false;

        for bytes in records {
            let kind = PptRecordType::from(u16::from_le_bytes([bytes[2], bytes[3]]));
            if kind == PptRecordType::OEPlaceholderAtom {
                if placeholder.is_some() || saw_later_slot {
                    return corrupted(
                        "PlaceholderAtom is duplicated or appears outside its ClientData slot",
                    );
                }
                let (record, consumed) = PptRecord::parse_strict(&bytes, 0)?;
                if consumed != bytes.len() {
                    return corrupted("PlaceholderAtom was only partially parsed");
                }
                placeholder = Some(PowerPointPlaceholderAtom::parse(&record, context)?);
                after_placeholder = true;
                continue;
            }

            if is_later_client_data_slot(kind) {
                saw_later_slot = true;
            }
            if after_placeholder && is_earlier_client_data_slot(kind) {
                return corrupted(
                    "A pre-placeholder ClientData record appears after PlaceholderAtom",
                );
            }
            if after_placeholder {
                after_records.push(bytes);
            } else {
                before_records.push(bytes);
            }
        }

        let retained =
            before_records
                .iter()
                .chain(&after_records)
                .try_fold(0usize, |total, record| {
                    total.checked_add(record.len()).ok_or_else(|| {
                        PptError::Corrupted("Retained ClientData size overflow".into())
                    })
                })?;
        check_limit(
            retained,
            limits.max_retained_bytes,
            "retained ClientData records",
        )?;
        Ok(Self {
            placeholder,
            before_records,
            after_records,
        })
    }

    pub fn before_records(&self) -> &[Vec<u8>] {
        &self.before_records
    }

    pub fn after_records(&self) -> &[Vec<u8>] {
        &self.after_records
    }

    /// Serialize the PPT-record payload while retaining unrelated data exactly.
    pub fn to_client_data_payload(
        &self,
        context: PowerPointPlaceholderContext,
        limits: PowerPointPlaceholderLimits,
    ) -> Result<Vec<u8>> {
        let count = self
            .before_records
            .len()
            .checked_add(self.after_records.len())
            .and_then(|value| value.checked_add(usize::from(self.placeholder.is_some())))
            .ok_or_else(|| PptError::Corrupted("ClientData record count overflow".into()))?;
        check_limit(
            count,
            limits.max_client_data_records,
            "ClientData record count",
        )?;
        let mut data = Vec::new();
        for record in &self.before_records {
            data.extend_from_slice(record);
        }
        if let Some(placeholder) = self.placeholder {
            data.extend_from_slice(&placeholder.to_bytes(context)?);
        }
        for record in &self.after_records {
            data.extend_from_slice(record);
        }
        check_limit(
            data.len(),
            limits.max_client_data_bytes,
            "ClientData payload",
        )?;
        Self::parse_client_data_payload(&data, context, limits)?;
        Ok(data)
    }

    /// Serialize a complete OfficeArt `ClientData` record.
    pub fn to_officeart_client_data(
        &self,
        context: PowerPointPlaceholderContext,
        limits: PowerPointPlaceholderLimits,
    ) -> Result<Vec<u8>> {
        let data = self.to_client_data_payload(context, limits)?;
        Ok(encode_record(0x0f, 0, OFFICEART_CLIENT_DATA_TYPE, &data))
    }
}

fn validate_context(
    kind: PowerPointPlaceholderKind,
    context: PowerPointPlaceholderContext,
) -> Result<()> {
    use PowerPointPlaceholderContext as C;
    use PowerPointPlaceholderKind as K;
    let valid = match kind {
        K::MasterTitle | K::MasterBody => context == C::MainMaster,
        K::MasterCenterTitle | K::MasterSubTitle => context == C::TitleMaster,
        K::MasterNotesSlideImage | K::MasterNotesBody => context == C::NotesMaster,
        K::MasterDate | K::MasterSlideNumber | K::MasterFooter => matches!(
            context,
            C::MainMaster | C::TitleMaster | C::NotesMaster | C::HandoutMaster
        ),
        K::MasterHeader => matches!(context, C::NotesMaster | C::HandoutMaster),
        K::NotesSlideImage | K::NotesBody => context == C::NotesSlide,
        K::Title
        | K::Body
        | K::CenterTitle
        | K::SubTitle
        | K::VerticalTitle
        | K::VerticalBody
        | K::Object
        | K::Graph
        | K::Table
        | K::ClipArt
        | K::OrgChart
        | K::Media
        | K::VerticalObject
        | K::Picture => context == C::PresentationSlide,
    };
    if valid {
        Ok(())
    } else {
        corrupted("PlaceholderAtom kind is invalid for its owning slide context")
    }
}

fn is_earlier_client_data_slot(kind: PptRecordType) -> bool {
    matches!(
        kind,
        PptRecordType::ShapeAtom
            | PptRecordType::ShapeFlags10Atom
            | PptRecordType::ExternalObjectRefAtom
            | PptRecordType::AnimationInfo
            | PptRecordType::InteractiveInfo
    )
}

fn is_later_client_data_slot(kind: PptRecordType) -> bool {
    matches!(
        kind,
        PptRecordType::RecolorInfoAtom
            | PptRecordType::ProgTags
            | PptRecordType::RoundTripNewPlaceholderId12Atom
            | PptRecordType::RoundTripShapeId12Atom
            | PptRecordType::RoundTripHFPlaceholder12Atom
            | PptRecordType::RoundTripShapeCheckSumForCustomLayouts12Atom
    )
}

fn split_records(data: &[u8], max_records: usize) -> Result<Vec<Vec<u8>>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        if records.len() >= max_records {
            return corrupted("ClientData exceeds its record-count limit");
        }
        let header_end = offset
            .checked_add(8)
            .ok_or_else(|| PptError::Corrupted("ClientData header offset overflow".into()))?;
        if header_end > data.len() {
            return corrupted("Truncated PPT record header in ClientData");
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
            return corrupted("PPT record extends beyond ClientData");
        }
        records.push(data[offset..end].to_vec());
        offset = end;
    }
    Ok(records)
}

fn encode_record(version: u16, instance: u16, kind: u16, data: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(8usize.saturating_add(data.len()));
    output.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    output.extend_from_slice(&kind.to_le_bytes());
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

    fn atom(kind: u8, size: u8, unused: [u8; 2]) -> Vec<u8> {
        let mut data = 7i32.to_le_bytes().to_vec();
        data.extend_from_slice(&[kind, size, unused[0], unused[1]]);
        record(0, 0, PptRecordType::OEPlaceholderAtom.as_u16(), &data)
    }

    #[test]
    fn round_trips_every_discriminant_and_undefined_bytes() {
        for raw in 1u8..=0x1a {
            let context = match raw {
                1 | 2 => PowerPointPlaceholderContext::MainMaster,
                3 | 4 => PowerPointPlaceholderContext::TitleMaster,
                5 | 6 | 10 => PowerPointPlaceholderContext::NotesMaster,
                7..=9 => PowerPointPlaceholderContext::HandoutMaster,
                11 | 12 => PowerPointPlaceholderContext::NotesSlide,
                _ => PowerPointPlaceholderContext::PresentationSlide,
            };
            let bytes = atom(raw, raw % 3, [0xa5, 0x5a]);
            let (record, consumed) = PptRecord::parse_strict(&bytes, 0).unwrap();
            assert_eq!(consumed, bytes.len());
            let parsed = PowerPointPlaceholderAtom::parse(&record, context).unwrap();
            assert_eq!(parsed.unused, [0xa5, 0x5a]);
            assert_eq!(parsed.to_bytes(context).unwrap(), bytes);
        }
    }

    #[test]
    fn validates_context_none_unknown_size_headers_and_lengths() {
        assert!(
            PowerPointPlaceholderAtom::parse_payload(
                &[0, 0, 0, 0, 0, 0, 0, 0],
                PowerPointPlaceholderContext::PresentationSlide,
            )
            .is_err()
        );
        assert!(
            PowerPointPlaceholderAtom::parse_payload(
                &[0, 0, 0, 0, 0x1b, 0, 0, 0],
                PowerPointPlaceholderContext::PresentationSlide,
            )
            .is_err()
        );
        assert!(
            PowerPointPlaceholderAtom::parse_payload(
                &[0, 0, 0, 0, 0x0d, 3, 0, 0],
                PowerPointPlaceholderContext::PresentationSlide,
            )
            .is_err()
        );
        assert!(
            PowerPointPlaceholderAtom::parse_payload(
                &[0, 0, 0, 0, 1, 0, 0, 0],
                PowerPointPlaceholderContext::PresentationSlide,
            )
            .is_err()
        );
        assert!(
            PowerPointPlaceholderAtom::parse_payload(
                &[0; 7],
                PowerPointPlaceholderContext::MainMaster,
            )
            .is_err()
        );

        let bad_header = record(1, 0, PptRecordType::OEPlaceholderAtom.as_u16(), &[0; 8]);
        let (record, _) = PptRecord::parse_strict(&bad_header, 0).unwrap();
        assert!(
            PowerPointPlaceholderAtom::parse(
                &record,
                PowerPointPlaceholderContext::PresentationSlide,
            )
            .is_err()
        );
    }

    #[test]
    fn enforces_client_data_ownership_order_and_round_trip() {
        let before = record(0, 0, PptRecordType::ExternalObjectRefAtom.as_u16(), &[0; 4]);
        let placeholder = atom(0x0d, 0, [7, 9]);
        let after = record(0, 0, PptRecordType::RecolorInfoAtom.as_u16(), &[]);
        let payload = [before.clone(), placeholder, after.clone()].concat();
        let complete = record(0x0f, 0, OFFICEART_CLIENT_DATA_TYPE, &payload);
        let limits = PowerPointPlaceholderLimits::default();
        let projection = PowerPointPlaceholderProjection::parse_officeart_client_data(
            &complete,
            PowerPointPlaceholderContext::PresentationSlide,
            limits,
        )
        .unwrap();
        assert_eq!(projection.before_records(), &[before]);
        assert_eq!(projection.after_records(), &[after]);
        assert_eq!(
            projection
                .to_officeart_client_data(PowerPointPlaceholderContext::PresentationSlide, limits)
                .unwrap(),
            complete
        );

        let duplicate = [atom(0x0d, 0, [0; 2]), atom(0x0e, 0, [0; 2])].concat();
        assert!(
            PowerPointPlaceholderProjection::parse_client_data_payload(
                &duplicate,
                PowerPointPlaceholderContext::PresentationSlide,
                limits,
            )
            .is_err()
        );
        let late = [
            record(0, 0, PptRecordType::RecolorInfoAtom.as_u16(), &[]),
            atom(0x0d, 0, [0; 2]),
        ]
        .concat();
        assert!(
            PowerPointPlaceholderProjection::parse_client_data_payload(
                &late,
                PowerPointPlaceholderContext::PresentationSlide,
                limits,
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_truncation_and_all_limits() {
        let mut data = record(0, 0, PptRecordType::ExternalObjectRefAtom.as_u16(), &[0; 4]);
        let defaults = PowerPointPlaceholderLimits::default();
        let cases = [
            PowerPointPlaceholderLimits {
                max_client_data_bytes: data.len() - 1,
                ..defaults
            },
            PowerPointPlaceholderLimits {
                max_client_data_records: 0,
                ..defaults
            },
            PowerPointPlaceholderLimits {
                max_retained_bytes: data.len() - 1,
                ..defaults
            },
        ];
        for limits in cases {
            assert!(
                PowerPointPlaceholderProjection::parse_client_data_payload(
                    &data,
                    PowerPointPlaceholderContext::PresentationSlide,
                    limits,
                )
                .is_err()
            );
        }
        data.pop();
        assert!(
            PowerPointPlaceholderProjection::parse_client_data_payload(
                &data,
                PowerPointPlaceholderContext::PresentationSlide,
                defaults,
            )
            .is_err()
        );
    }
}
