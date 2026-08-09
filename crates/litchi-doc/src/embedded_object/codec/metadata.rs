//! `[MS-OLEDS]` metadata codecs for DOC `ObjectPool` storages.

use super::super::model::{Clipboard, CompObj, Info, Kind, Metadata, Ole, Unknown};
use super::{array_at, corrupted, u32_at};
use crate::package::Result;
use litchi_codepage::Mbcs;
use litchi_ole_common::object::Object;
use std::ops::Range;
use std::sync::Arc;

const COMP_OBJ_STREAM: &str = "\u{1}CompObj";
const OLE_STREAM: &str = "\u{1}Ole";
const UNICODE_MARKER: u32 = 0x71B2_39F4;
const MAX_STRING_UNITS: usize = 1_048_576;
const MAX_RESERVED_ANSI: usize = 0x28;
const MAX_CLIPBOARD_UNITS: u32 = 0x190;

#[derive(Clone, Copy)]
struct Cursor<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Cursor<'a> {
    const fn new(data: &'a [u8], offset: usize) -> Self {
        Self { data, offset }
    }

    const fn position(self) -> usize {
        self.offset
    }

    const fn remaining(self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }

    fn peek_u32(self, name: &str) -> Result<u32> {
        u32_at(self.data, self.offset).map_err(|_| corrupted(format!("truncated {name}")))
    }

    fn u32(&mut self, name: &str) -> Result<u32> {
        let value = self.peek_u32(name)?;
        self.offset = self
            .offset
            .checked_add(4)
            .ok_or_else(|| corrupted(format!("{name} offset overflow")))?;
        Ok(value)
    }

    fn u64(&mut self, name: &str) -> Result<u64> {
        let bytes = self.take(8, name)?;
        Ok(u64::from_le_bytes(
            bytes
                .try_into()
                .map_err(|_| corrupted(format!("truncated {name}")))?,
        ))
    }

    fn take(&mut self, length: usize, name: &str) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| corrupted(format!("{name} range overflow")))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| corrupted(format!("truncated {name}")))?;
        self.offset = end;
        Ok(bytes)
    }
}

impl CompObj {
    pub(in crate::embedded_object) fn read(data: Arc<[u8]>) -> Result<Self> {
        if data.len() < 28 {
            return Err(corrupted("CompObj header is truncated"));
        }
        let raw = data;
        let data = raw.as_ref();
        let reserved1 = u32_at(data, 0)?;
        let version = u32_at(data, 4)?;
        let reserved2 = array_at(data, 8, "CompObj reserved header")?;
        let mut cursor = Cursor::new(data, 28);
        let ansi_user_type = read_ansi_string(&mut cursor, "CompObj ANSI user type")?;
        let ansi_clipboard = read_ansi_clipboard(&mut cursor)?;

        let mut reserved_ansi_present = false;
        let mut unicode_marker = None;
        let mut unicode_user_type = None;
        let mut unicode_clipboard = None;
        let mut reserved_unicode_present = false;
        let trailing_start;

        if cursor.remaining() == 0 {
            trailing_start = cursor.position();
        } else {
            // Reserved1 is optional. A marker immediately after the ANSI
            // clipboard means that the producer omitted it.
            if cursor.remaining() < 4 {
                trailing_start = cursor.position();
            } else if cursor.peek_u32("CompObj field")? == UNICODE_MARKER {
                let (marker, user_type, clipboard, reserved, trailing) =
                    read_unicode_section(&mut cursor)?;
                unicode_marker = Some(marker);
                unicode_user_type = Some(user_type);
                unicode_clipboard = Some(clipboard);
                reserved_unicode_present = reserved;
                trailing_start = trailing;
            } else {
                let reserved_start = cursor.position();
                let mut probe = cursor;
                if read_ansi_reserved(&mut probe).is_ok() {
                    cursor = probe;
                    reserved_ansi_present = true;
                    if cursor.remaining() == 0 {
                        trailing_start = cursor.position();
                    } else if cursor.remaining() < 4 {
                        trailing_start = cursor.position();
                    } else if cursor.peek_u32("CompObj Unicode marker")? != UNICODE_MARKER {
                        trailing_start = cursor.position();
                    } else {
                        let (marker, user_type, clipboard, reserved, trailing) =
                            read_unicode_section(&mut cursor)?;
                        unicode_marker = Some(marker);
                        unicode_user_type = Some(user_type);
                        unicode_clipboard = Some(clipboard);
                        reserved_unicode_present = reserved;
                        trailing_start = trailing;
                    }
                } else {
                    trailing_start = reserved_start;
                }
            }
        }

        Ok(Self::from_parts(super::super::model::comp_obj::Parts {
            reserved1,
            version,
            reserved2,
            ansi_user_type,
            ansi_clipboard,
            reserved_ansi_present,
            unicode_marker,
            unicode_user_type,
            unicode_clipboard,
            reserved_unicode_present,
            raw,
            trailing_start,
        }))
    }
}

impl Ole {
    pub(in crate::embedded_object) fn read(data: Arc<[u8]>) -> Result<Self> {
        if data.len() < 16 {
            return Err(corrupted("Ole stream header is truncated"));
        }
        let raw = data;
        let data = raw.as_ref();
        let mut cursor = Cursor::new(data, 0);
        let version = cursor.u32("Ole version")?;
        if version != 0x0200_0001 {
            return Err(corrupted("Ole stream version is unsupported"));
        }
        let flags = cursor.u32("Ole flags")?;
        let kind = if flags & 1 == 0 {
            Kind::Embedded
        } else {
            Kind::Linked
        };
        let cache_storage = flags & 0x1000 != 0;
        let update_option = cursor.u32("Ole link update option")?;
        let reserved = cursor.u32("Ole reserved field")?;
        if reserved != 0 {
            return Err(corrupted("Ole reserved field is non-zero"));
        }

        let reserved_moniker = read_moniker(&mut cursor, "Ole reserved moniker")?;
        let mut relative_moniker = None;
        let mut absolute_moniker = None;
        let mut class_id = None;
        let mut reserved_display_name = None;
        let mut reserved2 = None;
        let mut local_update_time = None;
        let mut local_check_update_time = None;
        let mut remote_update_time = None;

        if kind.is_linked() {
            relative_moniker = read_moniker(&mut cursor, "Ole relative moniker")?;
            absolute_moniker = Some(
                read_required_moniker(&mut cursor, "Ole absolute moniker")?
                    .ok_or_else(|| corrupted("Ole absolute moniker is missing"))?,
            );
            if cursor.remaining() != 0 {
                let indicator = cursor.u32("Ole CLSID indicator")?;
                if indicator != u32::MAX {
                    return Err(corrupted("Ole CLSID indicator is not -1"));
                }
                class_id = Some(array_at(cursor.data, cursor.position(), "Ole CLSID")?);
                cursor.offset = cursor
                    .offset
                    .checked_add(16)
                    .ok_or_else(|| corrupted("Ole CLSID offset overflow"))?;
                reserved_display_name = Some(read_unicode_range(
                    &mut cursor,
                    "Ole reserved display name",
                )?);
                if cursor.remaining() >= 4 {
                    reserved2 = Some(cursor.u32("Ole reserved2")?);
                }
                if cursor.remaining() >= 24 {
                    local_update_time = Some(cursor.u64("Ole local update time")?);
                    local_check_update_time = Some(cursor.u64("Ole local check time")?);
                    remote_update_time = Some(cursor.u64("Ole remote update time")?);
                }
            }
        }

        let trailing_start = cursor.position();
        Ok(Self::from_parts(super::super::model::ole::Parts {
            version,
            flags,
            kind,
            cache_storage,
            update_option,
            reserved,
            reserved_moniker,
            relative_moniker,
            absolute_moniker,
            class_id,
            reserved_display_name,
            reserved2,
            local_update_time,
            local_check_update_time,
            remote_update_time,
            raw,
            trailing_start,
        }))
    }
}

impl Metadata {
    /// Reads one selected `ObjectPool` storage without activating its payload.
    pub fn of(object: &Object) -> Result<Self> {
        let mut comp_obj = None;
        let mut ole = None;
        let mut obj_info = None;
        let mut unknown = Vec::new();

        for stream in object.streams() {
            let known = stream.path().len() == 1;
            let name = stream.name();
            let bytes = stream.bytes_shared();
            match (known, name) {
                (true, Some(COMP_OBJ_STREAM)) if comp_obj.is_none() => {
                    match CompObj::read(Arc::clone(&bytes)) {
                        Ok(value) => comp_obj = Some(value),
                        Err(_) => unknown.push(Unknown::from_parts(stream.path(), bytes)),
                    }
                },
                (true, Some(OLE_STREAM)) if ole.is_none() => match Ole::read(Arc::clone(&bytes)) {
                    Ok(value) => ole = Some(value),
                    Err(_) => unknown.push(Unknown::from_parts(stream.path(), bytes)),
                },
                (true, Some(super::OBJ_INFO_STREAM)) if obj_info.is_none() => {
                    match Info::read(&bytes) {
                        Ok(value) => obj_info = Some(value),
                        Err(_) => unknown.push(Unknown::from_parts(stream.path(), bytes)),
                    }
                },
                _ => unknown.push(Unknown::from_parts(stream.path(), bytes)),
            }
        }

        Ok(Self::from_parts(
            object.storage().clsid().map(str::to_owned),
            comp_obj,
            ole,
            obj_info,
            unknown,
        ))
    }
}

fn read_ansi_string(cursor: &mut Cursor<'_>, name: &str) -> Result<String> {
    let range = read_ansi_range(cursor, name, MAX_STRING_UNITS)?;
    decode_ansi(&cursor.data[range], name)
}

fn read_ansi_range(cursor: &mut Cursor<'_>, name: &str, max_length: usize) -> Result<Range<usize>> {
    let length = cursor.u32(name)? as usize;
    if length > max_length {
        return Err(corrupted(format!("{name} exceeds metadata limit")));
    }
    if length == 0 {
        return Ok(cursor.position()..cursor.position());
    }
    let start = cursor.position();
    let bytes = cursor.take(length, name)?;
    if bytes.last() != Some(&0) {
        return Err(corrupted(format!("{name} is not null terminated")));
    }
    Ok(start..start + length - 1)
}

fn read_ansi_reserved(cursor: &mut Cursor<'_>) -> Result<()> {
    let _ = read_ansi_range(cursor, "CompObj ANSI reserved string", MAX_RESERVED_ANSI)?;
    Ok(())
}

fn read_unicode_string(cursor: &mut Cursor<'_>, name: &str) -> Result<String> {
    let range = read_unicode_range(cursor, name)?;
    let bytes = cursor
        .data
        .get(range)
        .ok_or_else(|| corrupted("Unicode string range is invalid"))?;
    let units = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).map_err(|_| corrupted(format!("{name} is not valid UTF-16")))
}

fn read_unicode_range(cursor: &mut Cursor<'_>, name: &str) -> Result<Range<usize>> {
    let units = cursor.u32(name)? as usize;
    if units > MAX_STRING_UNITS {
        return Err(corrupted(format!("{name} exceeds metadata limit")));
    }
    if units == 0 {
        return Ok(cursor.position()..cursor.position());
    }
    let length = units
        .checked_mul(2)
        .ok_or_else(|| corrupted(format!("{name} byte length overflows")))?;
    let start = cursor.position();
    let bytes = cursor.take(length, name)?;
    if bytes.len() < 2 || bytes[bytes.len() - 2..] != [0, 0] {
        return Err(corrupted(format!("{name} is not null terminated")));
    }
    Ok(start..start + length - 2)
}

fn read_unicode_reserved(cursor: &mut Cursor<'_>) -> Result<()> {
    let _ = read_unicode_range(cursor, "CompObj Unicode reserved string")?;
    Ok(())
}

fn read_unicode_section(cursor: &mut Cursor<'_>) -> Result<(u32, String, Clipboard, bool, usize)> {
    let marker = cursor.u32("CompObj Unicode marker")?;
    if marker != UNICODE_MARKER {
        return Err(corrupted("CompObj Unicode marker is invalid"));
    }
    let user_type = read_unicode_string(cursor, "CompObj Unicode user type")?;
    let clipboard = read_unicode_clipboard(cursor)?;
    let (reserved, trailing) = if cursor.remaining() == 0 {
        (false, cursor.position())
    } else {
        let reserved_start = cursor.position();
        let mut probe = *cursor;
        if read_unicode_reserved(&mut probe).is_ok() {
            *cursor = probe;
            (true, cursor.position())
        } else {
            (false, reserved_start)
        }
    };
    Ok((marker, user_type, clipboard, reserved, trailing))
}

fn read_ansi_clipboard(cursor: &mut Cursor<'_>) -> Result<Clipboard> {
    let marker = cursor.u32("CompObj ANSI clipboard marker")?;
    match marker {
        0 => Ok(Clipboard::None),
        0xffff_fffe | 0xffff_ffff => Ok(Clipboard::Standard(
            cursor.u32("CompObj ANSI clipboard format")?,
        )),
        1..=MAX_CLIPBOARD_UNITS => {
            let start = cursor.position();
            let bytes = cursor.take(marker as usize, "CompObj ANSI clipboard name")?;
            if bytes.last() != Some(&0) {
                return Err(corrupted(
                    "CompObj ANSI clipboard name is not null terminated",
                ));
            }
            Ok(Clipboard::Registered(decode_ansi(
                &cursor.data[start..start + marker as usize - 1],
                "CompObj ANSI clipboard name",
            )?))
        },
        _ => Err(corrupted("CompObj ANSI clipboard marker is invalid")),
    }
}

fn read_unicode_clipboard(cursor: &mut Cursor<'_>) -> Result<Clipboard> {
    let marker = cursor.u32("CompObj Unicode clipboard marker")?;
    match marker {
        0 => Ok(Clipboard::None),
        0xffff_fffe | 0xffff_ffff => Ok(Clipboard::Standard(
            cursor.u32("CompObj Unicode clipboard format")?,
        )),
        1..=MAX_CLIPBOARD_UNITS => {
            let range = read_unicode_range(cursor, "CompObj Unicode clipboard name")?;
            let bytes = cursor
                .data
                .get(range)
                .ok_or_else(|| corrupted("CompObj Unicode clipboard range is invalid"))?;
            let units = bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect::<Vec<_>>();
            Ok(Clipboard::Registered(String::from_utf16(&units).map_err(
                |_| corrupted("CompObj Unicode clipboard name is invalid"),
            )?))
        },
        _ => Err(corrupted("CompObj Unicode clipboard marker is invalid")),
    }
}

fn decode_ansi(bytes: &[u8], name: &str) -> Result<String> {
    Ok(Mbcs::WINDOWS_1252
        .decode(bytes)
        .map_err(|_| corrupted(format!("{name} is not valid ANSI")))?
        .into_owned())
}

fn read_moniker(cursor: &mut Cursor<'_>, name: &str) -> Result<Option<Range<usize>>> {
    let size = cursor.u32(&format!("{name} size"))? as usize;
    if size == 0 {
        return Ok(None);
    }
    if size < 4 {
        return Err(corrupted(format!(
            "{name} size is smaller than its size field"
        )));
    }
    let length = size - 4;
    let start = cursor.position();
    cursor.take(length, name)?;
    Ok(Some(start..start + length))
}

fn read_required_moniker(cursor: &mut Cursor<'_>, name: &str) -> Result<Option<Range<usize>>> {
    let value = read_moniker(cursor, name)?;
    if value.is_none() {
        return Err(corrupted(format!("{name} must be present")));
    }
    Ok(value)
}
