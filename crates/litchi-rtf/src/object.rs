//! Embedded and linked RTF object metadata.

use std::borrow::Cow;

const COMPOUND_FILE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Storage/link mode declared by an RTF `object` destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ObjectKind {
    /// An object embedded in the RTF file (`objemb`)
    Embedded,
    /// A link to external content (`objlink`)
    Link,
    /// A link that the producing application can update automatically (`objautlink`)
    AutoLink,
    /// An HTML object (`objhtml`)
    Html,
    /// No recognized kind control was present
    #[default]
    Unknown,
}

/// Bounds-checked view of the OLE ObjectHeader stored in `objdata`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OleObjectHeader<'a> {
    /// OLE version from the ObjectHeader
    pub ole_version: u32,
    /// OLE format identifier
    pub format_id: u32,
    /// Null-terminated class-name payload, without the terminator
    pub class_name: &'a [u8],
    /// Null-terminated topic-name payload, without the terminator
    pub topic_name: &'a [u8],
    /// Null-terminated item-name payload, without the terminator
    pub item_name: &'a [u8],
    /// Native object bytes
    pub native_data: &'a [u8],
}

impl OleObjectHeader<'_> {
    /// Whether the native payload starts with the OLE2 Compound File signature.
    #[inline]
    pub fn is_compound_file(&self) -> bool {
        self.native_data.starts_with(&COMPOUND_FILE_SIGNATURE)
    }
}

/// An RTF embedded or linked object.
#[derive(Debug, Clone)]
pub struct EmbeddedObject<'a> {
    /// Storage/link mode
    pub kind: ObjectKind,
    /// Programmatic class name (`objclass`)
    pub class_name: Cow<'a, str>,
    /// User-visible object name (`objname`)
    pub name: Cow<'a, str>,
    /// Object width in twips
    pub width: i32,
    /// Object height in twips
    pub height: i32,
    /// Whether the object is locked
    pub locked: bool,
    /// Whether the producer requested an external-link update
    pub update_requested: bool,
    /// Whether the producer requested size synchronization
    pub set_size: bool,
    /// Decoded `objdata`, including the OLE ObjectHeader
    pub data: Vec<u8>,
}

impl<'a> EmbeddedObject<'a> {
    /// Construct an empty object record.
    #[inline]
    pub fn new() -> Self {
        Self {
            kind: ObjectKind::Unknown,
            class_name: Cow::Borrowed(""),
            name: Cow::Borrowed(""),
            width: 0,
            height: 0,
            locked: false,
            update_requested: false,
            set_size: false,
            data: Vec::new(),
        }
    }

    /// Parse the decoded ObjectHeader without copying its strings or native payload.
    pub fn ole_header(&self) -> Option<OleObjectHeader<'_>> {
        let mut offset = 0usize;
        let ole_version = read_u32(&self.data, &mut offset)?;
        let format_id = read_u32(&self.data, &mut offset)?;
        let class_name = read_counted_bytes(&self.data, &mut offset)?;
        let topic_name = read_counted_bytes(&self.data, &mut offset)?;
        let item_name = read_counted_bytes(&self.data, &mut offset)?;
        let native_size = usize::try_from(read_u32(&self.data, &mut offset)?).ok()?;
        let native_end = offset.checked_add(native_size)?;
        let native_data = self.data.get(offset..native_end)?;
        Some(OleObjectHeader {
            ole_version,
            format_id,
            class_name,
            topic_name,
            item_name,
            native_data,
        })
    }
}

impl Default for EmbeddedObject<'_> {
    fn default() -> Self {
        Self::new()
    }
}

fn read_u32(data: &[u8], offset: &mut usize) -> Option<u32> {
    let end = offset.checked_add(4)?;
    let bytes: [u8; 4] = data.get(*offset..end)?.try_into().ok()?;
    *offset = end;
    Some(u32::from_le_bytes(bytes))
}

fn read_counted_bytes<'a>(data: &'a [u8], offset: &mut usize) -> Option<&'a [u8]> {
    let length = usize::try_from(read_u32(data, offset)?).ok()?;
    let end = offset.checked_add(length)?;
    let value = data.get(*offset..end)?;
    *offset = end;
    Some(value.strip_suffix(&[0]).unwrap_or(value))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_object_header_and_native_compound_payload() {
        let mut object = EmbeddedObject::new();
        object.data.extend_from_slice(&0x501_u32.to_le_bytes());
        object.data.extend_from_slice(&2_u32.to_le_bytes());
        object.data.extend_from_slice(&8_u32.to_le_bytes());
        object.data.extend_from_slice(b"Package\0");
        object.data.extend_from_slice(&0_u32.to_le_bytes());
        object.data.extend_from_slice(&0_u32.to_le_bytes());
        object.data.extend_from_slice(&8_u32.to_le_bytes());
        object.data.extend_from_slice(&COMPOUND_FILE_SIGNATURE);

        let header = object.ole_header().unwrap();
        assert_eq!(header.ole_version, 0x501);
        assert_eq!(header.format_id, 2);
        assert_eq!(header.class_name, b"Package");
        assert!(header.topic_name.is_empty());
        assert!(header.item_name.is_empty());
        assert!(header.is_compound_file());
    }

    #[test]
    fn rejects_truncated_object_header() {
        let mut object = EmbeddedObject::new();
        object.data = vec![1, 2, 3, 4];
        assert!(object.ole_header().is_none());
    }
}
