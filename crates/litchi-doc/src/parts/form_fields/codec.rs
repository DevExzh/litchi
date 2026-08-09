//! MS-DOC form-field wire codec.

use crate::package::Result;

use super::model::{FormFieldData, NilPicfAndBinData};
use super::validation;

/// NilPICFAndBinData.cbHeader: the mandated offset of binData.
pub(super) const CB_HEADER: u16 = 0x0044;
/// Size in bytes of the `NilPICFAndBinData` header (lcb + cbHeader + ignored).
pub(super) const HEADER_LEN: usize = 68;
/// FFData.version: the mandated version marker.
pub(super) const FF_DATA_VERSION: u32 = 0xFFFF_FFFF;
/// STTB fExtend marker for Unicode strings.
pub(super) const STTB_EXTENDED: u16 = 0xFFFF;

impl NilPicfAndBinData {
    /// Parse one `NilPICFAndBinData` from the front of data.
    ///
    /// data may extend past the structure, for example when it is the
    /// remainder of the Data stream; exactly lcb bytes are consumed.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < HEADER_LEN {
            return Err(validation::corrupted(
                "NilPICFAndBinData header is truncated",
            ));
        }
        let lcb = i32::from_le_bytes(data[0..4].try_into().expect("length checked"));
        if lcb < HEADER_LEN as i32 {
            return Err(validation::corrupted(
                "NilPICFAndBinData lcb is smaller than its header",
            ));
        }
        let lcb = usize::try_from(lcb)
            .map_err(|_| validation::corrupted("NilPICFAndBinData lcb is negative"))?;
        if lcb > data.len() {
            return Err(validation::corrupted(
                "NilPICFAndBinData extends past its containing data",
            ));
        }
        let cb_header = u16::from_le_bytes(data[4..6].try_into().expect("length checked"));
        if cb_header != CB_HEADER {
            return Err(validation::corrupted(
                "NilPICFAndBinData cbHeader is not 0x0044",
            ));
        }
        Ok(Self {
            bin_data: data[HEADER_LEN..lcb].to_vec(),
        })
    }

    /// Parse one `NilPICFAndBinData` at offset within the Data stream.
    pub fn parse_at(data_stream: &[u8], offset: u32) -> Result<Self> {
        let offset = usize::try_from(offset).map_err(|_| {
            validation::corrupted("NilPICFAndBinData offset does not fit in memory")
        })?;
        let data = data_stream.get(offset..).ok_or_else(|| {
            validation::corrupted("NilPICFAndBinData offset is past the Data stream")
        })?;
        Self::parse(data)
    }

    /// Re-encode the structure, reproducing the original bytes for
    /// well-formed input.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let lcb = (HEADER_LEN + self.bin_data.len()) as u32;
        let mut out = Vec::with_capacity(lcb as usize);
        out.extend_from_slice(&lcb.to_le_bytes());
        out.extend_from_slice(&CB_HEADER.to_le_bytes());
        out.extend_from_slice(&[0; HEADER_LEN - 6]);
        out.extend_from_slice(&self.bin_data);
        out
    }
}

impl FormFieldData {
    /// Parse one `FFData`, the binData of a `NilPICFAndBinData` whose cbHeader is
    /// 0x0044. Trailing bytes are rejected.
    pub fn parse(bin_data: &[u8]) -> Result<Self> {
        let mut cursor = Cursor::new(bin_data);
        let version = cursor.read_u32("FFData version")?;
        if version != FF_DATA_VERSION {
            return Err(validation::corrupted("FFData version is not 0xFFFFFFFF"));
        }
        let bits = cursor.read_u16("FFDataBits")?;
        let max_length = cursor.read_u16("FFData cch")?;
        let size_half_points = cursor.read_u16("FFData hps")?;
        let decoded = validation::decode_bits(bits, max_length, size_half_points)?;

        let name = cursor.read_xstz("xstzName", validation::MAX_NAME_CHARS)?;
        let mut default_text = None;
        let mut default_state = None;
        match decoded.kind {
            super::model::FormFieldDataKind::Text => {
                let text = cursor.read_xstz("xstzTextDef", validation::MAX_DEFAULT_TEXT_CHARS)?;
                validation::validate_text_default(decoded.text_kind, &text)?;
                default_text = Some(text);
            },
            super::model::FormFieldDataKind::CheckBox
            | super::model::FormFieldDataKind::DropDown => {
                let w_def = cursor.read_u16("FFData wDef")?;
                validation::validate_default_state(decoded.kind, w_def)?;
                default_state = Some(w_def);
            },
        }
        let text_format = cursor.read_xstz("xstzTextFormat", validation::MAX_TEXT_FORMAT_CHARS)?;
        validation::validate_text_format(decoded.kind, &text_format)?;
        let help_text = cursor.read_xstz("xstzHelpText", validation::MAX_HELP_TEXT_CHARS)?;
        let status_text = cursor.read_xstz("xstzStatText", validation::MAX_STATUS_TEXT_CHARS)?;
        let entry_macro = cursor.read_xstz("xstzEntryMcr", validation::MAX_MACRO_NAME_CHARS)?;
        let exit_macro = cursor.read_xstz("xstzExitMcr", validation::MAX_MACRO_NAME_CHARS)?;
        let items = if decoded.kind == super::model::FormFieldDataKind::DropDown {
            let items = cursor.read_sttb("hsttbDropList")?;
            let w_def = default_state.expect("dropdown reads wDef");
            validation::validate_dropdown_default(w_def, items.len())?;
            Some(items)
        } else {
            None
        };
        if !cursor.is_finished() {
            return Err(validation::corrupted("FFData has trailing bytes"));
        }

        Ok(Self {
            kind: decoded.kind,
            state: decoded.state,
            text_kind: decoded.text_kind,
            own_help: decoded.own_help,
            own_status: decoded.own_status,
            protected: decoded.protected,
            auto_size: decoded.auto_size,
            recalculate: decoded.recalculate,
            has_list_box: decoded.has_list_box,
            max_length,
            size_half_points,
            name,
            default_text,
            default_state,
            text_format,
            help_text,
            status_text,
            entry_macro,
            exit_macro,
            items,
        })
    }

    /// Parse the `FFData` of the `NilPICFAndBinData` stored at offset in the Data
    /// stream.
    pub fn parse_at(data_stream: &[u8], offset: u32) -> Result<Self> {
        Self::parse(NilPicfAndBinData::parse_at(data_stream, offset)?.bin_data())
    }

    /// Re-encode the `FFData`, reproducing the original bytes for well-formed
    /// input.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut bits = match self.kind {
            super::model::FormFieldDataKind::Text => 0,
            super::model::FormFieldDataKind::CheckBox => 1,
            super::model::FormFieldDataKind::DropDown => 2,
        };
        bits |= u16::from(self.state) << validation::IRES_SHIFT;
        if self.own_help {
            bits |= validation::F_OWN_HELP;
        }
        if self.own_status {
            bits |= validation::F_OWN_STAT;
        }
        if self.protected {
            bits |= validation::F_PROT;
        }
        if self.auto_size {
            bits |= validation::F_ISIZE;
        }
        if self.kind == super::model::FormFieldDataKind::Text {
            let text_kind = match self.text_kind {
                super::model::FormFieldTextKind::Regular => 0,
                super::model::FormFieldTextKind::Number => 1,
                super::model::FormFieldTextKind::Date => 2,
                super::model::FormFieldTextKind::CurrentDate => 3,
                super::model::FormFieldTextKind::CurrentTime => 4,
                super::model::FormFieldTextKind::Calculation => 5,
            };
            bits |= text_kind << validation::ITYPE_TXT_SHIFT;
        }
        if self.recalculate {
            bits |= validation::F_RECALC;
        }
        if self.has_list_box {
            bits |= validation::F_HAS_LISTBOX;
        }

        let mut out = Vec::new();
        out.extend_from_slice(&FF_DATA_VERSION.to_le_bytes());
        out.extend_from_slice(&bits.to_le_bytes());
        out.extend_from_slice(&self.max_length.to_le_bytes());
        out.extend_from_slice(&self.size_half_points.to_le_bytes());
        write_xstz(&mut out, &self.name);
        if let Some(default_text) = &self.default_text {
            write_xstz(&mut out, default_text);
        }
        if let Some(default_state) = self.default_state {
            out.extend_from_slice(&default_state.to_le_bytes());
        }
        write_xstz(&mut out, &self.text_format);
        write_xstz(&mut out, &self.help_text);
        write_xstz(&mut out, &self.status_text);
        write_xstz(&mut out, &self.entry_macro);
        write_xstz(&mut out, &self.exit_macro);
        if let Some(items) = &self.items {
            out.extend_from_slice(&STTB_EXTENDED.to_le_bytes());
            out.extend_from_slice(&(items.len() as u16).to_le_bytes());
            out.extend_from_slice(&0u16.to_le_bytes());
            for item in items {
                let units: Vec<u16> = item.encode_utf16().collect();
                out.extend_from_slice(&(units.len() as u16).to_le_bytes());
                for unit in units {
                    out.extend_from_slice(&unit.to_le_bytes());
                }
            }
        }
        out
    }
}

/// Little-endian byte cursor with exact-consumption tracking.
struct Cursor<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Cursor<'a> {
    const fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn is_finished(&self) -> bool {
        self.offset == self.data.len()
    }

    fn read_u16(&mut self, what: &str) -> Result<u16> {
        let bytes = self
            .data
            .get(self.offset..self.offset + 2)
            .ok_or_else(|| validation::corrupted(format!("{what} is truncated")))?;
        self.offset += 2;
        Ok(u16::from_le_bytes(
            bytes.try_into().expect("length checked"),
        ))
    }

    fn read_u32(&mut self, what: &str) -> Result<u32> {
        let bytes = self
            .data
            .get(self.offset..self.offset + 4)
            .ok_or_else(|| validation::corrupted(format!("{what} is truncated")))?;
        self.offset += 4;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("length checked"),
        ))
    }

    /// Read an Xstz (MS-DOC 2.9.354): cch UTF-16 code units plus a mandated
    /// zero terminator.
    fn read_xstz(&mut self, what: &str, max_chars: u16) -> Result<String> {
        let cch = self.read_u16(what)?;
        if cch > max_chars {
            return Err(validation::corrupted(format!(
                "{what} exceeds its length cap"
            )));
        }
        let byte_len = usize::from(cch) * 2;
        let bytes = self
            .data
            .get(self.offset..self.offset + byte_len)
            .ok_or_else(|| validation::corrupted(format!("{what} is truncated")))?;
        self.offset += byte_len;
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes(chunk.try_into().expect("length checked")))
            .collect();
        let text = String::from_utf16(&units)
            .map_err(|_| validation::corrupted(format!("{what} is not valid UTF-16")))?;
        let terminator = self.read_u16(what)?;
        if terminator != 0 {
            return Err(validation::corrupted(format!(
                "{what} is not null-terminated"
            )));
        }
        Ok(text)
    }

    /// Read an extended STTB of Unicode strings with no extra data, the
    /// hsttbDropList layout from MS-DOC 2.9.78.
    fn read_sttb(&mut self, what: &str) -> Result<Vec<String>> {
        if self.read_u16(what)? != STTB_EXTENDED {
            return Err(validation::corrupted(format!(
                "{what} is not an extended STTB"
            )));
        }
        let count = self.read_u16(what)?;
        if count > validation::MAX_DROPDOWN_ITEMS {
            return Err(validation::corrupted(format!("{what} exceeds 25 elements")));
        }
        if self.read_u16(what)? != 0 {
            return Err(validation::corrupted(format!(
                "{what} carries unexpected extra data"
            )));
        }
        let mut items = Vec::with_capacity(usize::from(count));
        for _ in 0..count {
            let cch = usize::from(self.read_u16(what)?);
            let byte_len = cch * 2;
            let bytes = self
                .data
                .get(self.offset..self.offset + byte_len)
                .ok_or_else(|| validation::corrupted(format!("{what} is truncated")))?;
            self.offset += byte_len;
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes(chunk.try_into().expect("length checked")))
                .collect();
            items.push(
                String::from_utf16(&units)
                    .map_err(|_| validation::corrupted(format!("{what} is not valid UTF-16")))?,
            );
        }
        Ok(items)
    }
}

/// Write an Xstz: cch, the UTF-16 code units, and a zero terminator.
pub(super) fn write_xstz(out: &mut Vec<u8>, text: &str) {
    let units: Vec<u16> = text.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    for unit in units {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out.extend_from_slice(&0u16.to_le_bytes());
}
