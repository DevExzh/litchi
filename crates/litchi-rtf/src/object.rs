//! Embedded and linked RTF object metadata.

use std::borrow::Cow;

use crate::{RtfError, RtfResult};

const COMPOUND_FILE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Maximum number of inert object destinations retained from one document.
pub const MAX_EMBEDDED_OBJECTS: usize = 65_536;
/// Maximum aggregate text metadata accepted for one object.
pub const MAX_OBJECT_METADATA_BYTES: usize = 1024 * 1024;
/// Maximum decoded `objdata` payload accepted for one object.
pub const MAX_OBJECT_DATA_BYTES: usize = 64 * 1024 * 1024;

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
    /// A Macintosh Edition Manager subscriber (`objsub`)
    Subscriber,
    /// A Macintosh Edition Manager publisher (`objpub`)
    Publisher,
    /// A Macintosh Installable Command embedder (`objicemb`)
    InstallableCommand,
    /// An OLE control (`objocx`)
    OleControl,
    /// No recognized kind control was present
    #[default]
    Unknown,
}

/// Requested representation of an object's `result` destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ObjectResultKind {
    /// Standard RTF (`rsltrtf`)
    Rtf,
    /// Plain text (`rslttxt`)
    Text,
    /// Windows metafile or MacPict (`rsltpict`)
    Picture,
    /// Bitmap (`rsltbmp`)
    Bitmap,
    /// HTML (`rslthtml`)
    Html,
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
    /// UTF-8 byte offset in the visible document body
    pub position: usize,
    /// Storage/link mode
    pub kind: ObjectKind,
    /// Whether the link targets another part of this document (`linkself`)
    pub link_self: bool,
    /// Programmatic class name (`objclass`)
    pub class_name: Cow<'a, str>,
    /// User-visible object name (`objname`)
    pub name: Cow<'a, str>,
    /// Optional alias for the object (`objalias`)
    pub alias: Option<Cow<'a, str>>,
    /// Optional name of the linked document section (`objsect`)
    pub section: Option<Cow<'a, str>>,
    /// Optional time the object was last updated (`objtime`)
    pub time: Option<crate::RtfTimestamp>,
    /// Optional original CLSID from the `oleclsid` destination
    pub class_id: Cow<'a, str>,
    /// Object width in twips
    pub width: i32,
    /// Object height in twips
    pub height: i32,
    /// Optional tab-stop alignment distance in twips (`objalign`)
    pub alignment: Option<i32>,
    /// Optional vertical baseline translation in twips (`objtransy`)
    pub translation_y: Option<i32>,
    /// Optional top crop in twips (`objcropt`)
    pub crop_top: Option<i32>,
    /// Optional bottom crop in twips (`objcropb`)
    pub crop_bottom: Option<i32>,
    /// Optional left crop in twips (`objcropl`)
    pub crop_left: Option<i32>,
    /// Optional right crop in twips (`objcropr`)
    pub crop_right: Option<i32>,
    /// Optional horizontal scale percentage (`objscalex`)
    pub scale_x: Option<i32>,
    /// Optional vertical scale percentage (`objscaley`)
    pub scale_y: Option<i32>,
    /// Whether the object is locked
    pub locked: bool,
    /// Whether the producer requested an external-link update
    pub update_requested: bool,
    /// Whether the producer requested size synchronization
    pub set_size: bool,
    /// Whether formatting from the current result should be retained (`rsltmerge`)
    pub merge_result: bool,
    /// Requested representation for the `result` destination
    pub result_kind: Option<ObjectResultKind>,
    /// Plain-text rendered fallback from the `result` destination
    pub result_text: Cow<'a, str>,
    /// Indices into `RtfDocument::pictures()` for rendered fallback images
    pub result_picture_indices: Vec<usize>,
    /// Decoded `objdata`, including the OLE ObjectHeader
    pub data: Vec<u8>,
}

impl<'a> EmbeddedObject<'a> {
    /// Construct an empty object record.
    #[inline]
    pub fn new() -> Self {
        Self {
            position: 0,
            kind: ObjectKind::Unknown,
            link_self: false,
            class_name: Cow::Borrowed(""),
            name: Cow::Borrowed(""),
            alias: None,
            section: None,
            time: None,
            class_id: Cow::Borrowed(""),
            width: 0,
            height: 0,
            alignment: None,
            translation_y: None,
            crop_top: None,
            crop_bottom: None,
            crop_left: None,
            crop_right: None,
            scale_x: None,
            scale_y: None,
            locked: false,
            update_requested: false,
            set_size: false,
            merge_result: false,
            result_kind: None,
            result_text: Cow::Borrowed(""),
            result_picture_indices: Vec::new(),
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

    /// Validate positional metadata and references into the shared picture store.
    pub fn validate(&self, body: &str, picture_count: usize) -> RtfResult<()> {
        if body.get(..self.position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF embedded object position is not a UTF-8 body boundary".to_string(),
            ));
        }
        let metadata_bytes = self
            .class_name
            .len()
            .checked_add(self.name.len())
            .and_then(|size| size.checked_add(self.class_id.len()))
            .and_then(|size| {
                size.checked_add(
                    self.alias.as_ref().map_or(0, |alias| alias.len()),
                )
            })
            .and_then(|size| {
                size.checked_add(
                    self.section.as_ref().map_or(0, |section| section.len()),
                )
            })
            .and_then(|size| size.checked_add(self.result_text.len()))
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF embedded object metadata size overflow".to_string(),
                )
            })?;
        if metadata_bytes > MAX_OBJECT_METADATA_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF embedded object metadata exceeds the safety limit".to_string(),
            ));
        }
        if self.data.len() > MAX_OBJECT_DATA_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF embedded object data exceeds the safety limit".to_string(),
            ));
        }
        if self
            .result_picture_indices
            .iter()
            .any(|index| *index >= picture_count)
        {
            return Err(RtfError::MalformedDocument(
                "RTF embedded object references a missing result picture".to_string(),
            ));
        }
        Ok(())
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
