//! `OfficeArtClientData` binary codec and grammar validation.

use std::sync::Arc;

use super::model::{
    ClientData, ClientDataChild, ClientDataChildKind, ClientDataLimits, MAX_KNOWN_CHILDREN, u32_at,
};
use crate::embedded::reference::Reference;
use crate::package::{Error, Result};

/// `OfficeArt` record type for `OfficeArtClientData`.
pub const OFFICE_ART_CLIENT_DATA_RECORD_TYPE: u16 = 0xF011;

const HEADER_LEN: usize = 8;

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

impl ClientDataChildKind {
    pub(super) fn record_type(self) -> u16 {
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
            Self::Unknown => 0,
        }
    }

    fn known_record_type(self) -> Option<u16> {
        (!self.is_unknown()).then(|| self.record_type())
    }

    pub(super) fn canonical_header(self) -> (u16, u16) {
        match self {
            Self::AnimationInfo | Self::MouseClickInteractiveInfo | Self::ProgrammableTags => {
                (0x0F, 0)
            },
            Self::MouseOverInteractiveInfo => (0x0F, 1),
            Self::ShapeFlags
            | Self::ShapeFlags10
            | Self::ExternalObjectReference
            | Self::Placeholder
            | Self::RecolorInfo
            | Self::RoundTripNewPlaceholderId12
            | Self::RoundTripShapeId12
            | Self::RoundTripHeaderFooterPlaceholder12
            | Self::RoundTripShapeChecksumForCustomLayouts12
            | Self::Unknown => (0, 0),
        }
    }

    pub(super) fn slot(self) -> Option<u8> {
        match self {
            Self::ShapeFlags => Some(0),
            Self::ShapeFlags10 => Some(1),
            Self::ExternalObjectReference => Some(2),
            Self::AnimationInfo => Some(3),
            Self::MouseClickInteractiveInfo => Some(4),
            Self::MouseOverInteractiveInfo => Some(5),
            Self::Placeholder => Some(6),
            Self::RecolorInfo => Some(7),
            Self::ProgrammableTags
            | Self::RoundTripNewPlaceholderId12
            | Self::RoundTripShapeId12
            | Self::RoundTripHeaderFooterPlaceholder12
            | Self::RoundTripShapeChecksumForCustomLayouts12 => Some(8),
            Self::Unknown => None,
        }
    }
}

impl ClientDataChild {
    /// Construct a canonical child record for the selected alternative.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(kind: ClientDataChildKind, payload: Vec<u8>) -> Result<Self> {
        let record_type = kind.known_record_type().ok_or_else(|| {
            Error::InvalidFormat("opaque client-data children require a raw record type".into())
        })?;
        let (version, instance) = kind.canonical_header();
        let child = Self {
            kind,
            record_type,
            version,
            instance,
            payload: Arc::from(payload.into_boxed_slice()),
        };
        child.validate()?;
        Ok(child)
    }

    /// Construct an inert producer-defined child without interpreting it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn opaque(version: u16, instance: u16, record_type: u16, payload: Vec<u8>) -> Result<Self> {
        let child = Self {
            kind: ClientDataChildKind::Unknown,
            record_type,
            version,
            instance,
            payload: Arc::from(payload.into_boxed_slice()),
        };
        child.validate()?;
        Ok(child)
    }

    /// Serialize this complete child record.
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        encode_record(self.version, self.instance, self.record_type, &self.payload)
    }

    fn validate(&self) -> Result<()> {
        if self.version > 0x000F {
            return corrupted(format!(
                "client-data recVer is out of range: {}",
                self.version
            ));
        }
        if self.instance > 0x0FFF {
            return corrupted(format!(
                "client-data recInstance is out of range: {}",
                self.instance
            ));
        }
        if let Some(expected_record_type) = self.kind.known_record_type() {
            if self.record_type != expected_record_type {
                return corrupted(format!(
                    "{:?} recType must be 0x{expected_record_type:04X}",
                    self.kind
                ));
            }
        } else if is_known_record_type(self.record_type) {
            return corrupted(format!(
                "opaque client-data record type 0x{:04X} is a defined child",
                self.record_type
            ));
        }

        if self.kind.is_unknown() {
            return Ok(());
        }
        let (expected_version, expected_instance) = self.kind.canonical_header();
        if self.version != expected_version {
            return corrupted(format!("{:?} recVer must be {expected_version}", self.kind));
        }
        if self.kind == ClientDataChildKind::ProgrammableTags {
            // MS-PPT only recommends zero here. Preserve nonzero producer values inertly.
        } else if self.instance != expected_instance {
            return corrupted(format!(
                "{:?} recInstance must be {expected_instance}",
                self.kind
            ));
        }

        match self.kind {
            ClientDataChildKind::ShapeFlags => {
                require_len(&self.payload, 1, "ShapeFlagsAtom")?;
                if self.payload[0] & 0xFE != 0 {
                    return corrupted("ShapeFlagsAtom reserved bits must be zero");
                }
            },
            ClientDataChildKind::ShapeFlags10 => {
                require_len(&self.payload, 1, "ShapeFlags10Atom")?;
                if self.payload[0] & !0x04 != 0 {
                    return corrupted("ShapeFlags10Atom reserved bits must be zero");
                }
            },
            ClientDataChildKind::ExternalObjectReference => {
                Reference::parse_payload(&self.payload)?;
            },
            ClientDataChildKind::Placeholder => {
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
            ClientDataChildKind::RoundTripNewPlaceholderId12 => {
                require_len(&self.payload, 1, "RoundTripNewPlaceholderId12Atom")?;
                if !matches!(self.payload[0], 25 | 26) {
                    return corrupted(
                        "RoundTripNewPlaceholderId12Atom has an invalid placeholder ID",
                    );
                }
            },
            ClientDataChildKind::RoundTripShapeId12 => {
                require_len(&self.payload, 4, "RoundTripShapeId12Atom")?;
            },
            ClientDataChildKind::RoundTripHeaderFooterPlaceholder12 => {
                require_len(&self.payload, 1, "RoundTripHFPlaceholder12Atom")?;
                if !matches!(self.payload[0], 7..=10) {
                    return corrupted("RoundTripHFPlaceholder12Atom has an invalid placeholder ID");
                }
            },
            ClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12 => {
                require_len(
                    &self.payload,
                    8,
                    "RoundTripShapeCheckSumForCustomLayouts12Atom",
                )?;
            },
            ClientDataChildKind::AnimationInfo
            | ClientDataChildKind::MouseClickInteractiveInfo
            | ClientDataChildKind::MouseOverInteractiveInfo
            | ClientDataChildKind::RecolorInfo
            | ClientDataChildKind::ProgrammableTags => {},
            ClientDataChildKind::Unknown => unreachable!("unknown children return above"),
        }
        Ok(())
    }
}

impl ClientData {
    /// Construct and validate a container from ordered typed children.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(children: Vec<ClientDataChild>) -> Result<Self> {
        validate_sequence(&children, ClientDataLimits::default())?;
        Ok(Self { children })
    }

    /// Parse one exact complete `OfficeArtClientData` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, ClientDataLimits::default())
    }

    /// Parse one exact complete record with explicit resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: &[u8], limits: ClientDataLimits) -> Result<Self> {
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
            .ok_or_else(|| Error::Corrupted("OfficeArtClientData length overflows".into()))?;
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
                .ok_or_else(|| Error::Corrupted("client-data child length overflows".into()))?;
            if end > bytes.len() {
                return corrupted("OfficeArtClientData child payload is truncated");
            }
            let kind = classify(child_type, child_instance)?;
            let child = ClientDataChild {
                kind,
                record_type: child_type,
                version: child_version,
                instance: child_instance,
                payload: Arc::from(bytes[offset + HEADER_LEN..end].to_vec().into_boxed_slice()),
            };
            child.validate()?;
            children.push(child);
            offset = end;
        }
        validate_sequence(&children, limits)?;
        Ok(Self { children })
    }

    /// Iterate over the §2.7.4 tail in its original record order.
    pub fn round_trip_records(&self) -> impl Iterator<Item = &ClientDataChild> {
        self.children
            .iter()
            .filter(|child| child.kind.slot() == Some(8))
    }

    /// Serialize the complete container and every child byte-exactly.
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limits(ClientDataLimits::default())
    }

    /// Serialize with the same resource bounds used by a source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub fn to_bytes_with_limits(&self, limits: ClientDataLimits) -> Result<Vec<u8>> {
        validate_sequence(&self.children, limits)?;
        let payload_capacity = self
            .children
            .iter()
            .try_fold(0usize, |total, child| {
                total
                    .checked_add(HEADER_LEN)
                    .and_then(|with_header| with_header.checked_add(child.payload.len()))
            })
            .ok_or_else(|| {
                Error::Corrupted("OfficeArtClientData payload length overflows".into())
            })?;
        let mut payload = Vec::with_capacity(payload_capacity);
        for child in &self.children {
            payload.extend_from_slice(&child.to_bytes()?);
        }
        encode_record(0x0F, 0, OFFICE_ART_CLIENT_DATA_RECORD_TYPE, &payload)
    }

    pub(super) fn validate_with_limits(&self, limits: ClientDataLimits) -> Result<()> {
        validate_sequence(&self.children, limits)
    }
}

fn classify(record_type: u16, instance: u16) -> Result<ClientDataChildKind> {
    let kind = match record_type {
        RT_SHAPE_FLAGS => ClientDataChildKind::ShapeFlags,
        RT_SHAPE_FLAGS_10 => ClientDataChildKind::ShapeFlags10,
        RT_EXTERNAL_OBJECT_REF => ClientDataChildKind::ExternalObjectReference,
        RT_ANIMATION_INFO => ClientDataChildKind::AnimationInfo,
        RT_INTERACTIVE_INFO if instance == 0 => ClientDataChildKind::MouseClickInteractiveInfo,
        RT_INTERACTIVE_INFO if instance == 1 => ClientDataChildKind::MouseOverInteractiveInfo,
        RT_INTERACTIVE_INFO => {
            return corrupted(format!(
                "InteractiveInfo recInstance must be 0 or 1, got {instance}"
            ));
        },
        RT_PLACEHOLDER => ClientDataChildKind::Placeholder,
        RT_RECOLOR_INFO => ClientDataChildKind::RecolorInfo,
        RT_PROG_TAGS => ClientDataChildKind::ProgrammableTags,
        RT_ROUND_TRIP_NEW_PLACEHOLDER_ID_12 => ClientDataChildKind::RoundTripNewPlaceholderId12,
        RT_ROUND_TRIP_SHAPE_ID_12 => ClientDataChildKind::RoundTripShapeId12,
        RT_ROUND_TRIP_HF_PLACEHOLDER_12 => ClientDataChildKind::RoundTripHeaderFooterPlaceholder12,
        RT_ROUND_TRIP_SHAPE_CHECKSUM_FOR_CL_12 => {
            ClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12
        },
        _ => ClientDataChildKind::Unknown,
    };
    Ok(kind)
}

fn validate_sequence(children: &[ClientDataChild], limits: ClientDataLimits) -> Result<()> {
    if children.len() > limits.max_child_records || children.len() > MAX_KNOWN_CHILDREN {
        return corrupted("OfficeArtClientData has too many child records");
    }
    let mut seen = [false; MAX_KNOWN_CHILDREN];
    let mut last_slot = None;
    let mut payload_len = 0usize;
    for child in children {
        child.validate()?;
        if child.payload.len() > limits.max_child_payload_bytes {
            return corrupted("OfficeArtClientData child exceeds the configured size limit");
        }
        payload_len = payload_len
            .checked_add(HEADER_LEN)
            .and_then(|value| value.checked_add(child.payload.len()))
            .ok_or_else(|| {
                Error::Corrupted("OfficeArtClientData payload length overflows".into())
            })?;
        if payload_len > limits.max_payload_bytes {
            return corrupted("OfficeArtClientData payload exceeds the configured limit");
        }

        let Some(slot) = child.kind.slot() else {
            continue;
        };
        if last_slot.is_some_and(|previous| slot < previous) {
            return corrupted(format!(
                "{:?} appears outside its OfficeArtClientData slot",
                child.kind
            ));
        }
        last_slot = Some(slot);
        let identity = child.kind as usize;
        if identity >= seen.len() {
            return corrupted("OfficeArtClientData child kind exceeds the known grammar");
        }
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

fn is_known_record_type(record_type: u16) -> bool {
    matches!(
        record_type,
        RT_ROUND_TRIP_SHAPE_ID_12
            | RT_ROUND_TRIP_HF_PLACEHOLDER_12
            | RT_ROUND_TRIP_SHAPE_CHECKSUM_FOR_CL_12
            | RT_EXTERNAL_OBJECT_REF
            | RT_PLACEHOLDER
            | RT_SHAPE_FLAGS
            | RT_SHAPE_FLAGS_10
            | RT_ROUND_TRIP_NEW_PLACEHOLDER_ID_12
            | RT_RECOLOR_INFO
            | RT_INTERACTIVE_INFO
            | RT_ANIMATION_INFO
            | RT_PROG_TAGS
    )
}

fn require_len(payload: &[u8], expected: usize, name: &str) -> Result<()> {
    if payload.len() != expected {
        return corrupted(format!("{name} payload must be exactly {expected} bytes"));
    }
    Ok(())
}

pub(super) fn encode_record(
    version: u16,
    instance: u16,
    record_type: u16,
    data: &[u8],
) -> Result<Vec<u8>> {
    if version > 0x000F {
        return Err(Error::InvalidFormat(
            "PowerPoint record version exceeds its four-bit field".into(),
        ));
    }
    if instance > 0x0FFF {
        return Err(Error::InvalidFormat(
            "PowerPoint record instance exceeds its twelve-bit field".into(),
        ));
    }
    let length = u32::try_from(data.len())
        .map_err(|_err| Error::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(HEADER_LEN.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
