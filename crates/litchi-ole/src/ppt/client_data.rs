//! Typed grammar for MS-PPT OfficeArtClientData containers.

use super::package::{PptError, Result};

/// OfficeArt record type for OfficeArtClientData.
pub const OFFICE_ART_CLIENT_DATA_RECORD_TYPE: u16 = 0xF011;

const HEADER_LEN: usize = 8;
const MAX_DEFINED_CHILDREN: usize = 13;

const RT_ROUND_TRIP_SHAPE_ID_12: u16 = 0x041F;
const RT_ROUND_TRIP_HF_PLACEHOLDER_12: u16 = 0x0420;
const RT_ROUND_TRIP_SHAPE_CHECKSUM_FOR_CL_12: u16 = 0x0426;
const RT_EXTERNAL_OBJECT_REF: u16 = 0x0BC1;
const RT_PLACEHOLDER: u16 = 0x0BC3;
const RT_SHAPE_FLAGS: u16 = 0x0BDB;
const RT_SHAPE_FLAGS_10: u16 = 0x0BDC;
const RT_ROUND_TRIP_NEW_PLACEHOLDER_ID_12: u16 = 0x0BDD;
const RT_RECOLOR_INFO: u16 = 0x0FE7;
const RT_INTERACTIVE_INFO: u16 = 0x0FF2;
const RT_ANIMATION_INFO: u16 = 0x1014;
const RT_PROG_TAGS: u16 = 0x1388;

/// Resource limits for a shape client-data container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointClientDataLimits {
    /// Maximum enclosing payload size.
    pub max_payload_bytes: usize,
    /// Maximum payload size of any individual child.
    pub max_child_payload_bytes: usize,
    /// Maximum number of child records.
    pub max_child_records: usize,
}

impl Default for PowerPointClientDataLimits {
    fn default() -> Self {
        Self {
            max_payload_bytes: 16 * 1024 * 1024,
            max_child_payload_bytes: 16 * 1024 * 1024,
            max_child_records: MAX_DEFINED_CHILDREN,
        }
    }
}

/// Every record alternative permitted by MS-PPT §§2.7.3-2.7.4.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PowerPointClientDataChildKind {
    ShapeFlags,
    ShapeFlags10,
    ExternalObjectReference,
    AnimationInfo,
    MouseClickInteractiveInfo,
    MouseOverInteractiveInfo,
    Placeholder,
    RecolorInfo,
    ProgrammableTags,
    RoundTripNewPlaceholderId12,
    RoundTripShapeId12,
    RoundTripHeaderFooterPlaceholder12,
    RoundTripShapeChecksumForCustomLayouts12,
}

impl PowerPointClientDataChildKind {
    fn record_type(self) -> u16 {
        match self {
            Self::ShapeFlags => RT_SHAPE_FLAGS,
            Self::ShapeFlags10 => RT_SHAPE_FLAGS_10,
            Self::ExternalObjectReference => RT_EXTERNAL_OBJECT_REF,
            Self::AnimationInfo => RT_ANIMATION_INFO,
            Self::MouseClickInteractiveInfo | Self::MouseOverInteractiveInfo => RT_INTERACTIVE_INFO,
            Self::Placeholder => RT_PLACEHOLDER,
            Self::RecolorInfo => RT_RECOLOR_INFO,
            Self::ProgrammableTags => RT_PROG_TAGS,
            Self::RoundTripNewPlaceholderId12 => RT_ROUND_TRIP_NEW_PLACEHOLDER_ID_12,
            Self::RoundTripShapeId12 => RT_ROUND_TRIP_SHAPE_ID_12,
            Self::RoundTripHeaderFooterPlaceholder12 => RT_ROUND_TRIP_HF_PLACEHOLDER_12,
            Self::RoundTripShapeChecksumForCustomLayouts12 => {
                RT_ROUND_TRIP_SHAPE_CHECKSUM_FOR_CL_12
            },
        }
    }

    fn canonical_header(self) -> (u16, u16) {
        match self {
            Self::AnimationInfo => (0x0F, 0),
            Self::MouseClickInteractiveInfo => (0x0F, 0),
            Self::MouseOverInteractiveInfo => (0x0F, 1),
            Self::ProgrammableTags => (0x0F, 0),
            _ => (0, 0),
        }
    }

    fn slot(self) -> u8 {
        match self {
            Self::ShapeFlags => 0,
            Self::ShapeFlags10 => 1,
            Self::ExternalObjectReference => 2,
            Self::AnimationInfo => 3,
            Self::MouseClickInteractiveInfo => 4,
            Self::MouseOverInteractiveInfo => 5,
            Self::Placeholder => 6,
            Self::RecolorInfo => 7,
            _ => 8,
        }
    }
}

/// One validated child record, retaining its exact payload and advisory instance.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointClientDataChild {
    kind: PowerPointClientDataChildKind,
    version: u16,
    instance: u16,
    payload: Vec<u8>,
}

impl PowerPointClientDataChild {
    /// Construct a canonical child record for the selected alternative.
    pub fn new(kind: PowerPointClientDataChildKind, payload: Vec<u8>) -> Result<Self> {
        let (version, instance) = kind.canonical_header();
        let child = Self {
            kind,
            version,
            instance,
            payload,
        };
        child.validate()?;
        Ok(child)
    }

    /// Classified child kind.
    pub fn kind(&self) -> PowerPointClientDataChildKind {
        self.kind
    }

    /// Record version retained from the input.
    pub fn version(&self) -> u16 {
        self.version
    }

    /// Record instance retained from the input.
    pub fn instance(&self) -> u16 {
        self.instance
    }

    /// Exact child payload.
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// External object ID when this is an ExObjRefAtom.
    pub fn external_object_id(&self) -> Option<u32> {
        (self.kind == PowerPointClientDataChildKind::ExternalObjectReference)
            .then(|| u32_at(&self.payload, 0))
    }

    /// PowerPoint 12 shape ID when this is the corresponding round-trip atom.
    pub fn round_trip_shape_id(&self) -> Option<u32> {
        (self.kind == PowerPointClientDataChildKind::RoundTripShapeId12)
            .then(|| u32_at(&self.payload, 0))
    }

    /// Shape and text checksums when this is the checksum round-trip atom.
    pub fn round_trip_checksums(&self) -> Option<(u32, u32)> {
        (self.kind == PowerPointClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12)
            .then(|| (u32_at(&self.payload, 0), u32_at(&self.payload, 4)))
    }

    /// Placeholder ID carried by either one-byte placeholder round-trip atom.
    pub fn round_trip_placeholder_id(&self) -> Option<u8> {
        matches!(
            self.kind,
            PowerPointClientDataChildKind::RoundTripNewPlaceholderId12
                | PowerPointClientDataChildKind::RoundTripHeaderFooterPlaceholder12
        )
        .then(|| self.payload[0])
    }

    /// Serialize this complete child record.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        encode_record(
            self.version,
            self.instance,
            self.kind.record_type(),
            &self.payload,
        )
    }

    fn validate(&self) -> Result<()> {
        let (expected_version, expected_instance) = self.kind.canonical_header();
        if self.version != expected_version {
            return corrupted(format!("{:?} recVer must be {expected_version}", self.kind));
        }
        if self.kind == PowerPointClientDataChildKind::ProgrammableTags {
            // MS-PPT only recommends zero here. Preserve nonzero producer values inertly.
        } else if self.instance != expected_instance {
            return corrupted(format!(
                "{:?} recInstance must be {expected_instance}",
                self.kind
            ));
        }

        match self.kind {
            PowerPointClientDataChildKind::ShapeFlags => {
                require_len(&self.payload, 1, "ShapeFlagsAtom")?;
                if self.payload[0] & 0xFE != 0 {
                    return corrupted("ShapeFlagsAtom reserved bits must be zero");
                }
            },
            PowerPointClientDataChildKind::ShapeFlags10 => {
                require_len(&self.payload, 1, "ShapeFlags10Atom")?;
                if self.payload[0] & !0x04 != 0 {
                    return corrupted("ShapeFlags10Atom reserved bits must be zero");
                }
            },
            PowerPointClientDataChildKind::ExternalObjectReference => {
                require_len(&self.payload, 4, "ExObjRefAtom")?;
            },
            PowerPointClientDataChildKind::Placeholder => {
                require_len(&self.payload, 8, "PlaceholderAtom")?;
                if self.payload[4] == 0 || self.payload[4] > 26 {
                    return corrupted("PlaceholderAtom has an invalid placementId");
                }
                if self.payload[5] > 2 {
                    return corrupted("PlaceholderAtom has an invalid size");
                }
                if self.payload[6] != 0 || self.payload[7] != 0 {
                    return corrupted("PlaceholderAtom unused bytes must be zero");
                }
            },
            PowerPointClientDataChildKind::RoundTripNewPlaceholderId12 => {
                require_len(&self.payload, 1, "RoundTripNewPlaceholderId12Atom")?;
                if !matches!(self.payload[0], 25 | 26) {
                    return corrupted(
                        "RoundTripNewPlaceholderId12Atom has an invalid placeholder ID",
                    );
                }
            },
            PowerPointClientDataChildKind::RoundTripShapeId12 => {
                require_len(&self.payload, 4, "RoundTripShapeId12Atom")?;
            },
            PowerPointClientDataChildKind::RoundTripHeaderFooterPlaceholder12 => {
                require_len(&self.payload, 1, "RoundTripHFPlaceholder12Atom")?;
                if !matches!(self.payload[0], 7..=10) {
                    return corrupted("RoundTripHFPlaceholder12Atom has an invalid placeholder ID");
                }
            },
            PowerPointClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12 => {
                require_len(
                    &self.payload,
                    8,
                    "RoundTripShapeCheckSumForCustomLayouts12Atom",
                )?;
            },
            PowerPointClientDataChildKind::AnimationInfo
            | PowerPointClientDataChildKind::MouseClickInteractiveInfo
            | PowerPointClientDataChildKind::MouseOverInteractiveInfo
            | PowerPointClientDataChildKind::RecolorInfo
            | PowerPointClientDataChildKind::ProgrammableTags => {},
        }
        Ok(())
    }
}

/// A complete, ordered OfficeArtClientData container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointClientData {
    children: Vec<PowerPointClientDataChild>,
}

impl PowerPointClientData {
    /// Construct and validate a container from ordered typed children.
    pub fn new(children: Vec<PowerPointClientDataChild>) -> Result<Self> {
        validate_sequence(&children, PowerPointClientDataLimits::default())?;
        Ok(Self { children })
    }

    /// Parse one exact complete OfficeArtClientData record.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, PowerPointClientDataLimits::default())
    }

    /// Parse one exact complete record with explicit resource limits.
    pub fn parse_with_limits(bytes: &[u8], limits: PowerPointClientDataLimits) -> Result<Self> {
        if bytes.len() < HEADER_LEN {
            return corrupted("OfficeArtClientData header is truncated");
        }
        let version_instance = u16::from_le_bytes([bytes[0], bytes[1]]);
        let version = version_instance & 0x0F;
        let instance = version_instance >> 4;
        if version != 0x0F {
            return corrupted(format!(
                "OfficeArtClientData recVer must be 15, got {version}"
            ));
        }
        if instance != 0 {
            return corrupted(format!(
                "OfficeArtClientData recInstance must be 0, got {instance}"
            ));
        }
        let record_type = u16::from_le_bytes([bytes[2], bytes[3]]);
        if record_type != OFFICE_ART_CLIENT_DATA_RECORD_TYPE {
            return corrupted(format!(
                "OfficeArtClientData recType must be 0xF011, got 0x{record_type:04X}"
            ));
        }
        let payload_len = u32_at(bytes, 4) as usize;
        if payload_len > limits.max_payload_bytes {
            return corrupted("OfficeArtClientData payload exceeds the configured limit");
        }
        let expected_len = HEADER_LEN
            .checked_add(payload_len)
            .ok_or_else(|| PptError::Corrupted("OfficeArtClientData length overflows".into()))?;
        if bytes.len() != expected_len {
            return corrupted(format!(
                "OfficeArtClientData record length is {}, expected {expected_len}",
                bytes.len()
            ));
        }

        let mut children = Vec::new();
        let mut offset = HEADER_LEN;
        while offset < bytes.len() {
            if children.len() >= limits.max_child_records {
                return corrupted("OfficeArtClientData exceeds the configured child count");
            }
            if bytes.len() - offset < HEADER_LEN {
                return corrupted("OfficeArtClientData child header is truncated");
            }
            let header = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
            let child_version = header & 0x0F;
            let child_instance = header >> 4;
            let child_type = u16::from_le_bytes([bytes[offset + 2], bytes[offset + 3]]);
            let child_len = u32_at(bytes, offset + 4) as usize;
            if child_len > limits.max_child_payload_bytes {
                return corrupted("OfficeArtClientData child exceeds the configured size limit");
            }
            let end = offset
                .checked_add(HEADER_LEN)
                .and_then(|value| value.checked_add(child_len))
                .ok_or_else(|| PptError::Corrupted("client-data child length overflows".into()))?;
            if end > bytes.len() {
                return corrupted("OfficeArtClientData child payload is truncated");
            }
            let kind = classify(child_type, child_instance)?;
            let child = PowerPointClientDataChild {
                kind,
                version: child_version,
                instance: child_instance,
                payload: bytes[offset + HEADER_LEN..end].to_vec(),
            };
            child.validate()?;
            children.push(child);
            offset = end;
        }
        validate_sequence(&children, limits)?;
        Ok(Self { children })
    }

    /// Ordered child records.
    pub fn children(&self) -> &[PowerPointClientDataChild] {
        &self.children
    }

    /// Return the unique record of a particular kind.
    pub fn child(&self, kind: PowerPointClientDataChildKind) -> Option<&PowerPointClientDataChild> {
        self.children.iter().find(|child| child.kind == kind)
    }

    pub fn shape_flags(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::ShapeFlags)
    }

    pub fn shape_flags10(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::ShapeFlags10)
    }

    pub fn external_object_reference(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::ExternalObjectReference)
    }

    pub fn animation_info(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::AnimationInfo)
    }

    pub fn mouse_click_interactive_info(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::MouseClickInteractiveInfo)
    }

    pub fn mouse_over_interactive_info(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::MouseOverInteractiveInfo)
    }

    pub fn placeholder(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::Placeholder)
    }

    pub fn recolor_info(&self) -> Option<&PowerPointClientDataChild> {
        self.child(PowerPointClientDataChildKind::RecolorInfo)
    }

    /// Iterate over the §2.7.4 tail in its original record order.
    pub fn round_trip_records(&self) -> impl Iterator<Item = &PowerPointClientDataChild> {
        self.children.iter().filter(|child| child.kind.slot() == 8)
    }

    /// Serialize the complete container and every child byte-exactly.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_sequence(&self.children, PowerPointClientDataLimits::default())?;
        let mut payload = Vec::new();
        for child in &self.children {
            payload.extend_from_slice(&child.to_bytes()?);
        }
        encode_record(0x0F, 0, OFFICE_ART_CLIENT_DATA_RECORD_TYPE, &payload)
    }
}

fn classify(record_type: u16, instance: u16) -> Result<PowerPointClientDataChildKind> {
    let kind = match record_type {
        RT_SHAPE_FLAGS => PowerPointClientDataChildKind::ShapeFlags,
        RT_SHAPE_FLAGS_10 => PowerPointClientDataChildKind::ShapeFlags10,
        RT_EXTERNAL_OBJECT_REF => PowerPointClientDataChildKind::ExternalObjectReference,
        RT_ANIMATION_INFO => PowerPointClientDataChildKind::AnimationInfo,
        RT_INTERACTIVE_INFO if instance == 0 => {
            PowerPointClientDataChildKind::MouseClickInteractiveInfo
        },
        RT_INTERACTIVE_INFO if instance == 1 => {
            PowerPointClientDataChildKind::MouseOverInteractiveInfo
        },
        RT_INTERACTIVE_INFO => {
            return corrupted(format!(
                "InteractiveInfo recInstance must be 0 or 1, got {instance}"
            ));
        },
        RT_PLACEHOLDER => PowerPointClientDataChildKind::Placeholder,
        RT_RECOLOR_INFO => PowerPointClientDataChildKind::RecolorInfo,
        RT_PROG_TAGS => PowerPointClientDataChildKind::ProgrammableTags,
        RT_ROUND_TRIP_NEW_PLACEHOLDER_ID_12 => {
            PowerPointClientDataChildKind::RoundTripNewPlaceholderId12
        },
        RT_ROUND_TRIP_SHAPE_ID_12 => PowerPointClientDataChildKind::RoundTripShapeId12,
        RT_ROUND_TRIP_HF_PLACEHOLDER_12 => {
            PowerPointClientDataChildKind::RoundTripHeaderFooterPlaceholder12
        },
        RT_ROUND_TRIP_SHAPE_CHECKSUM_FOR_CL_12 => {
            PowerPointClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12
        },
        _ => {
            return corrupted(format!(
                "record type 0x{record_type:04X} is not valid in OfficeArtClientData"
            ));
        },
    };
    Ok(kind)
}

fn validate_sequence(
    children: &[PowerPointClientDataChild],
    limits: PowerPointClientDataLimits,
) -> Result<()> {
    if children.len() > limits.max_child_records || children.len() > MAX_DEFINED_CHILDREN {
        return corrupted("OfficeArtClientData has too many child records");
    }
    let mut seen = [false; MAX_DEFINED_CHILDREN];
    let mut last_slot = 0;
    for (index, child) in children.iter().enumerate() {
        child.validate()?;
        let slot = child.kind.slot();
        if index != 0 && slot < last_slot {
            return corrupted(format!(
                "{:?} appears outside its OfficeArtClientData slot",
                child.kind
            ));
        }
        last_slot = slot;
        let identity = child.kind as usize;
        if seen[identity] {
            return corrupted(format!(
                "OfficeArtClientData contains duplicate {:?} records",
                child.kind
            ));
        }
        seen[identity] = true;
    }
    Ok(())
}

fn require_len(payload: &[u8], expected: usize, name: &str) -> Result<()> {
    if payload.len() != expected {
        return corrupted(format!("{name} payload must be exactly {expected} bytes"));
    }
    Ok(())
}

fn u32_at(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}

fn encode_record(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(HEADER_LEN.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn child(
        version: u16,
        instance: u16,
        kind: PowerPointClientDataChildKind,
        data: &[u8],
    ) -> Vec<u8> {
        encode_record(version, instance, kind.record_type(), data).unwrap()
    }

    fn container(payload: &[u8]) -> Vec<u8> {
        encode_record(0x0F, 0, OFFICE_ART_CLIENT_DATA_RECORD_TYPE, payload).unwrap()
    }

    #[test]
    fn parses_accesses_and_round_trips_the_complete_ordered_grammar() {
        let mut payload = Vec::new();
        payload.extend(child(0, 0, PowerPointClientDataChildKind::ShapeFlags, &[1]));
        payload.extend(child(
            0,
            0,
            PowerPointClientDataChildKind::ShapeFlags10,
            &[4],
        ));
        payload.extend(child(
            0,
            0,
            PowerPointClientDataChildKind::ExternalObjectReference,
            &77u32.to_le_bytes(),
        ));
        payload.extend(child(
            0,
            0,
            PowerPointClientDataChildKind::Placeholder,
            &[3, 0, 0, 0, 13, 1, 0, 0],
        ));
        payload.extend(child(
            0x0F,
            7,
            PowerPointClientDataChildKind::ProgrammableTags,
            &[0xAA, 0xBB],
        ));
        payload.extend(child(
            0,
            0,
            PowerPointClientDataChildKind::RoundTripShapeId12,
            &0x1020u32.to_le_bytes(),
        ));
        payload.extend(child(
            0,
            0,
            PowerPointClientDataChildKind::RoundTripNewPlaceholderId12,
            &[26],
        ));
        payload.extend(child(
            0,
            0,
            PowerPointClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12,
            &[1, 0, 0, 0, 2, 0, 0, 0],
        ));
        let bytes = container(&payload);

        let parsed = PowerPointClientData::parse(&bytes).unwrap();
        assert!(parsed.shape_flags().is_some());
        assert_eq!(
            parsed
                .external_object_reference()
                .unwrap()
                .external_object_id(),
            Some(77)
        );
        assert_eq!(
            parsed
                .child(PowerPointClientDataChildKind::ProgrammableTags)
                .unwrap()
                .instance(),
            7
        );
        assert_eq!(
            parsed
                .child(PowerPointClientDataChildKind::RoundTripShapeId12)
                .unwrap()
                .round_trip_shape_id(),
            Some(0x1020)
        );
        assert_eq!(parsed.round_trip_records().count(), 4);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn canonical_constructor_writes_valid_records() {
        let records = vec![
            PowerPointClientDataChild::new(
                PowerPointClientDataChildKind::ExternalObjectReference,
                9u32.to_le_bytes().to_vec(),
            )
            .unwrap(),
            PowerPointClientDataChild::new(
                PowerPointClientDataChildKind::RoundTripHeaderFooterPlaceholder12,
                vec![7],
            )
            .unwrap(),
        ];
        let value = PowerPointClientData::new(records).unwrap();
        assert_eq!(
            PowerPointClientData::parse(&value.to_bytes().unwrap()).unwrap(),
            value
        );
    }

    #[test]
    fn rejects_outer_header_length_truncation_and_trailing_data() {
        let valid = container(&[]);
        for index in [0usize, 1, 2] {
            let mut bad = valid.clone();
            bad[index] ^= 1;
            assert!(PowerPointClientData::parse(&bad).is_err());
        }
        let mut bad_length = valid.clone();
        bad_length[4..8].copy_from_slice(&1u32.to_le_bytes());
        assert!(PowerPointClientData::parse(&bad_length).is_err());
        let mut trailing = valid.clone();
        trailing.push(0);
        assert!(PowerPointClientData::parse(&trailing).is_err());
        assert!(PowerPointClientData::parse(&valid[..7]).is_err());
    }

    #[test]
    fn rejects_unknown_duplicate_and_out_of_order_children() {
        let flags = child(0, 0, PowerPointClientDataChildKind::ShapeFlags, &[0]);
        let placeholder = child(
            0,
            0,
            PowerPointClientDataChildKind::Placeholder,
            &[0, 0, 0, 0, 13, 0, 0, 0],
        );
        assert!(
            PowerPointClientData::parse(&container(&[flags.clone(), flags.clone()].concat()))
                .is_err()
        );
        assert!(PowerPointClientData::parse(&container(&[placeholder, flags].concat())).is_err());
        assert!(
            PowerPointClientData::parse(&container(&encode_record(0, 0, 0x2222, &[]).unwrap()))
                .is_err()
        );

        let shape_id = child(
            0,
            0,
            PowerPointClientDataChildKind::RoundTripShapeId12,
            &[0; 4],
        );
        assert!(
            PowerPointClientData::parse(&container(&[shape_id.clone(), shape_id].concat()))
                .is_err()
        );
    }

    #[test]
    fn rejects_bad_child_headers_reserved_values_and_boundaries() {
        assert!(
            PowerPointClientData::parse(&container(&child(
                1,
                0,
                PowerPointClientDataChildKind::ShapeFlags,
                &[0]
            )))
            .is_err()
        );
        assert!(
            PowerPointClientData::parse(&container(&child(
                0,
                1,
                PowerPointClientDataChildKind::ShapeFlags,
                &[0]
            )))
            .is_err()
        );
        assert!(
            PowerPointClientData::parse(&container(&child(
                0,
                0,
                PowerPointClientDataChildKind::ShapeFlags,
                &[2]
            )))
            .is_err()
        );
        assert!(
            PowerPointClientData::parse(&container(&child(
                0,
                0,
                PowerPointClientDataChildKind::Placeholder,
                &[0, 0, 0, 0, 0, 0, 0, 0]
            )))
            .is_err()
        );
        assert!(
            PowerPointClientData::parse(&container(&child(
                0,
                0,
                PowerPointClientDataChildKind::RoundTripNewPlaceholderId12,
                &[24]
            )))
            .is_err()
        );
        assert!(
            PowerPointClientData::parse(&container(&child(
                0,
                2,
                PowerPointClientDataChildKind::MouseClickInteractiveInfo,
                &[]
            )))
            .is_err()
        );

        let mut truncated = child(
            0,
            0,
            PowerPointClientDataChildKind::ExternalObjectReference,
            &[0; 4],
        );
        truncated.pop();
        assert!(PowerPointClientData::parse(&container(&truncated)).is_err());
    }

    #[test]
    fn enforces_payload_child_size_and_count_limits() {
        let record = container(&child(
            0x0F,
            0,
            PowerPointClientDataChildKind::ProgrammableTags,
            &[1, 2],
        ));
        assert!(
            PowerPointClientData::parse_with_limits(
                &record,
                PowerPointClientDataLimits {
                    max_payload_bytes: 1,
                    ..Default::default()
                }
            )
            .is_err()
        );
        assert!(
            PowerPointClientData::parse_with_limits(
                &record,
                PowerPointClientDataLimits {
                    max_child_payload_bytes: 1,
                    ..Default::default()
                }
            )
            .is_err()
        );
        assert!(
            PowerPointClientData::parse_with_limits(
                &record,
                PowerPointClientDataLimits {
                    max_child_records: 0,
                    ..Default::default()
                }
            )
            .is_err()
        );
    }
}
