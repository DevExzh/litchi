//! Bounded, lossless record framing for one master layout.

use super::model::Limits;
use crate::package::{Error, Result};
use crate::records::Record;

struct Encoder {
    limits: Limits,
    records: usize,
}

/// Parse exactly one complete PPT record without resynchronizing over corrupt
/// bytes. The parent remains free to retain a source snapshot when this fails.
pub(super) fn parse(bytes: &[u8], limits: Limits) -> Result<Record> {
    if bytes.len() > limits.max_bytes {
        return invalid("master-layout record exceeds the configured byte limit");
    }
    let (record, consumed) = Record::parse_strict(bytes, 0)?;
    if consumed != bytes.len() {
        return corrupted("master-layout input contains trailing record bytes");
    }
    Ok(record)
}

/// Encode one complete record tree, rebuilding only nodes represented by
/// parsed children. Opaque records retain their payload verbatim.
pub(super) fn encode(root: &Record, limits: Limits) -> Result<Vec<u8>> {
    let mut encoder = Encoder { limits, records: 0 };
    encoder.record(root, 1)
}

impl Encoder {
    fn record(&mut self, record: &Record, depth: usize) -> Result<Vec<u8>> {
        if depth > self.limits.max_depth {
            return invalid("master-layout record nesting exceeds the configured depth limit");
        }
        self.records = self
            .records
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("master-layout record count overflow".into()))?;
        if self.records > self.limits.max_records {
            return invalid("master-layout record count exceeds the configured limit");
        }
        if record.version > 0x0f || record.instance > 0x0fff {
            return invalid("PPT record header exceeds its version or instance bit field");
        }

        if !record.children.is_empty() && !is_container(record.record_type_raw) {
            return invalid("non-container PPT records cannot own child records");
        }

        let payload = if record.children.is_empty() {
            record.data.clone()
        } else {
            let mut payload = Vec::new();
            for child in &record.children {
                let encoded = self.record(child, depth + 1)?;
                let total = payload.len().checked_add(encoded.len()).ok_or_else(|| {
                    Error::InvalidFormat("master-layout payload size overflow".into())
                })?;
                if total > self.limits.max_bytes {
                    return invalid("master-layout payload exceeds the configured byte limit");
                }
                payload.extend_from_slice(&encoded);
            }
            if !record.data.is_empty() && record.data != payload {
                return corrupted(
                    "master-layout container payload is not represented by its child tree",
                );
            }
            payload
        };
        let length = u32::try_from(payload.len())
            .map_err(|_err| Error::InvalidFormat("PPT record payload exceeds u32".into()))?;
        let total = payload
            .len()
            .checked_add(8)
            .ok_or_else(|| Error::InvalidFormat("PPT record size overflow".into()))?;
        if total > self.limits.max_bytes {
            return invalid("master-layout record exceeds the configured byte limit");
        }

        let mut bytes = Vec::new();
        bytes.try_reserve_exact(total).map_err(|_err| {
            Error::InvalidFormat("master-layout record allocation failed".into())
        })?;
        bytes.extend_from_slice(&(record.version | (record.instance << 4)).to_le_bytes());
        bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
        bytes.extend_from_slice(&length.to_le_bytes());
        bytes.extend_from_slice(&payload);
        Ok(bytes)
    }
}

/// Materialize a supported empty container before a nested edit.
pub(super) fn materialize(record: &mut Record, limits: Limits) -> Result<()> {
    if !is_container(record.record_type_raw) {
        return invalid("the selected record is not an editable PPT container");
    }
    if record.children.is_empty() && !record.data.is_empty() {
        let children = Record::parse_sequence_strict(&record.data, "master-layout container")?;
        if children.len() > limits.max_records {
            return invalid("master-layout container has too many child records");
        }
        record.children = children;
    }
    Ok(())
}

/// Keep the denormalized `data` field coherent for transaction-local views.
pub(super) fn sync(record: &mut Record, limits: Limits) -> Result<()> {
    if record.children.is_empty() {
        record.data_length = u32::try_from(record.data.len())
            .map_err(|_err| Error::InvalidFormat("PPT record payload exceeds u32".into()))?;
        return Ok(());
    }
    for child in &mut record.children {
        sync(child, limits)?;
    }
    let payload = encode_children(&record.children, limits)?;
    record.data_length = u32::try_from(payload.len())
        .map_err(|_err| Error::InvalidFormat("PPT container payload exceeds u32".into()))?;
    record.data = payload;
    Ok(())
}

fn encode_children(children: &[Record], limits: Limits) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for child in children {
        let encoded = encode(child, limits)?;
        let length = output
            .len()
            .checked_add(encoded.len())
            .ok_or_else(|| Error::InvalidFormat("PPT child payload size overflow".into()))?;
        if length > limits.max_bytes {
            return invalid("PPT child payload exceeds the configured byte limit");
        }
        output.extend_from_slice(&encoded);
    }
    Ok(output)
}

/// The subset of container records that can contain a navigable child path.
/// `OfficeArt` drawing payloads are deliberately opaque here; their own owner
/// remains responsible for `[MS-ODRAW]` editing.
pub(super) const fn is_container(raw: u16) -> bool {
    matches!(
        raw,
        1000      // Document
            | 1006  // Slide
            | 1008  // Notes
            | 1016  // MainMaster
            | 1018  // SlideViewInfo
            | 1023  // VBAInfo
            | 1033  // ExObjList
            | 2000  // DocInfoList
            | 2005  // FontCollection
            | 2006  // FontCollection10
            | 4040  // Kinsoku
            | 4055  // ExternalHyperlink
            | 4057  // HeadersFooters
            | 4068  // ExternalHyperlink9
            | 4080  // SlideListWithText
            | 4082  // InteractiveInfo
            | 5000  // ProgTags
            | 5001  // ProgStringTag
            | 5002  // ProgBinaryTag
            | 0x0fc9 // Handout
            | 0x2b00 // AnimationInfo
            | 0x2b02 // BuildList
            | 0x2b03 // ParaBuild
            | 0x2b04 // ChartBuild
            | 0x2b05 // DiagramBuild
            | 0x2b07 // ExtTimeNode
            | 0x2b08 // TimeConditionContainer
            | 0x2b09 // TimeBehaviorContainer
            | 0x2b0a // TimeAnimateBehaviorContainer
            | 0x2b0c // TimeEffectBehaviorContainer
            | 0x2b0e // TimeMotionBehaviorContainer
            | 0x2b0f // TimeRotationBehaviorContainer
            | 0x2b10 // TimeScaleBehaviorContainer
            | 0x2b11 // TimeSetBehaviorContainer
            | 0x2b12 // TimeCommandBehaviorContainer
            | 0x2b13 // TimeClientVisualElement
            | 0x2b14 // TimePropertyList
            | 0x2b15 // TimeVariantList
            | 0x2b16 // TimeAnimationValueList
    )
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
