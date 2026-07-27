//! Legacy Word form-field binary data (`NilPICFAndBinData`/`FFData`).
//!
//! [MS-DOC] §2.9.158 defines `NilPICFAndBinData`, the binary payload stored in
//! the Data stream for the picture character (U+0001) of a `FORMTEXT`,
//! `FORMCHECKBOX`, or `FORMDROPDOWN` field. When `cbHeader` is 0x0044 the
//! payload is an `FFData` ([MS-DOC] §2.9.78) describing the stored state of
//! the form field: its kind, default value, current checkbox or selection
//! state, help and status texts, and entry/exit macro names.
//!
//! Everything in this module is inert metadata. Stored macro names are kept
//! verbatim and are never resolved or invoked, forms are never filled,
//! checkbox and selection states are never changed, and fields are never
//! refreshed.
//!
//! The drop-down item list is always the inline `hsttbDropList` STTB inside
//! the `FFData`. The FIB's `fcFormFldSttbs`/`lcbFormFldSttbs` pair (table
//! pointer index 45) is defined as "undefined and MUST be ignored" with a
//! mandated length of zero ([MS-DOC] §2.5), so no table-stream fallback path
//! exists to parse.

use crate::doc::package::{DocError, Result};

/// `NilPICFAndBinData.cbHeader`: the mandated offset of `binData`.
const CB_HEADER: u16 = 0x0044;
/// Size in bytes of the `NilPICFAndBinData` header (lcb + cbHeader + ignored).
const HEADER_LEN: usize = 68;
/// `FFData.version`: the mandated version marker.
const FF_DATA_VERSION: u32 = 0xFFFF_FFFF;

// FFDataBits layout (MS-DOC 2.9.79), least-significant bit first.
const ITYPE_MASK: u16 = 0x0003;
const IRES_SHIFT: u16 = 2;
const IRES_MASK: u16 = 0x001F;
const F_OWN_HELP: u16 = 0x0080;
const F_OWN_STAT: u16 = 0x0100;
const F_PROT: u16 = 0x0200;
const F_ISIZE: u16 = 0x0400;
const ITYPE_TXT_SHIFT: u16 = 11;
const ITYPE_TXT_MASK: u16 = 0x0007;
const F_RECALC: u16 = 0x4000;
const F_HAS_LISTBOX: u16 = 0x8000;

/// `iRes` value marking an undefined checkbox state or drop-down selection.
const UNDEFINED_STATE: u8 = 25;
/// `FFData.cch` upper bound: the maximum text-box value length.
const MAX_TEXT_LENGTH: u16 = 32767;
/// `FFData.hps` bounds for a check box (in half-points).
const MIN_CHECKBOX_HPS: u16 = 2;
const MAX_CHECKBOX_HPS: u16 = 3168;
/// `xstzName.cch` upper bound.
const MAX_NAME_CHARS: u16 = 20;
/// `xstzTextDef.cch` upper bound.
const MAX_DEFAULT_TEXT_CHARS: u16 = 255;
/// `xstzTextFormat.cch` upper bound.
const MAX_TEXT_FORMAT_CHARS: u16 = 64;
/// `xstzHelpText.cch` upper bound.
const MAX_HELP_TEXT_CHARS: u16 = 255;
/// `xstzStatText.cch` upper bound.
const MAX_STATUS_TEXT_CHARS: u16 = 138;
/// `xstzEntryMcr.cch` and `xstzExitMcr.cch` upper bound.
const MAX_MACRO_NAME_CHARS: u16 = 32;
/// `hsttbDropList` element count upper bound.
const MAX_DROPDOWN_ITEMS: u16 = 25;
/// STTB `fExtend` marker for Unicode strings.
const STTB_EXTENDED: u16 = 0xFFFF;

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

/// The stored kind of a legacy form field (`FFDataBits.iType`, MS-DOC 2.9.79).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormFieldDataKind {
    /// `iTypeText` (0): a text box.
    Text,
    /// `iTypeChck` (1): a check box.
    CheckBox,
    /// `iTypeDrop` (2): a drop-down list box.
    DropDown,
}

/// The stored text-box value kind (`FFDataBits.iTypeTxt`, MS-DOC 2.9.79).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormFieldTextKind {
    /// `iTypeTxtReg` (0): regular text.
    Regular,
    /// `iTypeTxtNum` (1): a number.
    Number,
    /// `iTypeTxtDate` (2): a date or time.
    Date,
    /// `iTypeTxtCurDate` (3): the current date.
    CurrentDate,
    /// `iTypeTxtCurTime` (4): the current time.
    CurrentTime,
    /// `iTypeTxtCalc` (5): calculated from the stored default-text expression.
    Calculation,
}

/// The stored state of a check-box form field.
///
/// Undefined check boxes are treated as unchecked by Word; this type only
/// retains the stored state and never changes it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CheckBoxState {
    /// The check box is stored as unchecked (`iRes`/`wDef` 0).
    Unchecked,
    /// The check box is stored as checked (`iRes`/`wDef` 1).
    Checked,
    /// The stored checkbox state is undefined (`iRes` 25).
    Undefined,
}

/// A parsed `NilPICFAndBinData` wrapper (MS-DOC 2.9.158).
///
/// Only the `cbHeader == 0x0044` layout is accepted. The 62 ignored header
/// bytes are not retained; [`Self::to_bytes`] re-emits them as zero, which
/// reproduces the original bytes for well-formed input (where they MUST be
/// zero).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NilPicfAndBinData {
    bin_data: Vec<u8>,
}

impl NilPicfAndBinData {
    /// Parse one `NilPICFAndBinData` from the front of `data`.
    ///
    /// `data` may extend past the structure (for example, when it is the
    /// remainder of the Data stream); exactly `lcb` bytes are consumed.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < HEADER_LEN {
            return Err(corrupted("NilPICFAndBinData header is truncated"));
        }
        let lcb = i32::from_le_bytes(data[0..4].try_into().expect("length checked"));
        if lcb < HEADER_LEN as i32 {
            return Err(corrupted(
                "NilPICFAndBinData lcb is smaller than its header",
            ));
        }
        let lcb =
            usize::try_from(lcb).map_err(|_| corrupted("NilPICFAndBinData lcb is negative"))?;
        if lcb > data.len() {
            return Err(corrupted(
                "NilPICFAndBinData extends past its containing data",
            ));
        }
        let cb_header = u16::from_le_bytes(data[4..6].try_into().expect("length checked"));
        if cb_header != CB_HEADER {
            return Err(corrupted("NilPICFAndBinData cbHeader is not 0x0044"));
        }
        Ok(Self {
            bin_data: data[HEADER_LEN..lcb].to_vec(),
        })
    }

    /// Parse one `NilPICFAndBinData` at `offset` within the Data stream.
    pub fn parse_at(data_stream: &[u8], offset: u32) -> Result<Self> {
        let offset = usize::try_from(offset)
            .map_err(|_| corrupted("NilPICFAndBinData offset does not fit in memory"))?;
        let data = data_stream
            .get(offset..)
            .ok_or_else(|| corrupted("NilPICFAndBinData offset is past the Data stream"))?;
        Self::parse(data)
    }

    /// The stored binary payload (`binData`).
    pub fn bin_data(&self) -> &[u8] {
        &self.bin_data
    }

    /// Re-encode the structure, reproducing the original bytes for
    /// well-formed input.
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

/// Typed, inert form-field data (`FFData`, MS-DOC 2.9.78).
///
/// All values are stored state only: entry and exit macro names are retained
/// verbatim and never invoked, the form is never filled, checkbox and
/// selection states are never changed, and no field is refreshed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormFieldData {
    kind: FormFieldDataKind,
    /// Raw 5-bit `iRes`: checkbox state or drop-down selection (zero when the
    /// kind is text, where the spec mandates 0).
    state: u8,
    text_kind: FormFieldTextKind,
    own_help: bool,
    own_status: bool,
    protected: bool,
    auto_size: bool,
    recalculate: bool,
    has_list_box: bool,
    /// `FFData.cch`: maximum text-box value length (0 means unlimited; always
    /// 0 for non-text kinds).
    max_length: u16,
    /// `FFData.hps`: check-box size in half-points (meaningful for check boxes
    /// only).
    size_half_points: u16,
    name: String,
    /// `xstzTextDef`: stored default text (text boxes only).
    default_text: Option<String>,
    /// `wDef`: stored default checkbox state (0/1) or drop-down index.
    default_state: Option<u16>,
    text_format: String,
    help_text: String,
    status_text: String,
    entry_macro: String,
    exit_macro: String,
    /// `hsttbDropList` entries (drop-down lists only).
    items: Option<Vec<String>>,
}

impl FormFieldData {
    /// Parse one `FFData` (the `binData` of a `NilPICFAndBinData` whose
    /// `cbHeader` is 0x0044). Trailing bytes are rejected.
    pub fn parse(bin_data: &[u8]) -> Result<Self> {
        let mut cursor = Cursor::new(bin_data);
        let version = cursor.read_u32("FFData version")?;
        if version != FF_DATA_VERSION {
            return Err(corrupted("FFData version is not 0xFFFFFFFF"));
        }
        let bits = cursor.read_u16("FFDataBits")?;
        let max_length = cursor.read_u16("FFData cch")?;
        let size_half_points = cursor.read_u16("FFData hps")?;

        let kind = match bits & ITYPE_MASK {
            0 => FormFieldDataKind::Text,
            1 => FormFieldDataKind::CheckBox,
            2 => FormFieldDataKind::DropDown,
            _ => return Err(corrupted("FFDataBits iType is reserved")),
        };
        let state = u8::try_from((bits >> IRES_SHIFT) & IRES_MASK).expect("mask fits in u8");
        let own_help = bits & F_OWN_HELP != 0;
        let own_status = bits & F_OWN_STAT != 0;
        let protected = bits & F_PROT != 0;
        let auto_size = bits & F_ISIZE != 0;
        let text_kind_raw = (bits >> ITYPE_TXT_SHIFT) & ITYPE_TXT_MASK;
        let recalculate = bits & F_RECALC != 0;
        let has_list_box = bits & F_HAS_LISTBOX != 0;

        // iType-reserved relationships (MS-DOC 2.9.79).
        match kind {
            FormFieldDataKind::Text => {
                if state != 0 {
                    return Err(corrupted("FFDataBits iRes is not 0 for a text box"));
                }
            },
            FormFieldDataKind::CheckBox => {
                if !matches!(state, 0 | 1 | UNDEFINED_STATE) {
                    return Err(corrupted("FFDataBits iRes is not a checkbox state"));
                }
            },
            FormFieldDataKind::DropDown => {
                if state > UNDEFINED_STATE {
                    return Err(corrupted("FFDataBits iRes is not a drop-down selection"));
                }
            },
        }
        if auto_size && kind != FormFieldDataKind::CheckBox {
            return Err(corrupted(
                "FFDataBits iSize is set for a non-checkbox field",
            ));
        }
        if has_list_box != (kind == FormFieldDataKind::DropDown) {
            return Err(corrupted("FFDataBits fHasListBox disagrees with iType"));
        }
        let text_kind = match kind {
            FormFieldDataKind::Text => match text_kind_raw {
                0 => FormFieldTextKind::Regular,
                1 => FormFieldTextKind::Number,
                2 => FormFieldTextKind::Date,
                3 => FormFieldTextKind::CurrentDate,
                4 => FormFieldTextKind::CurrentTime,
                5 => FormFieldTextKind::Calculation,
                _ => return Err(corrupted("FFDataBits iTypeTxt is reserved")),
            },
            // iTypeTxt MUST be 0 and MUST be ignored for non-text kinds.
            _ => FormFieldTextKind::Regular,
        };
        if kind != FormFieldDataKind::Text && max_length != 0 {
            return Err(corrupted("FFData cch is not 0 for a non-text field"));
        }
        if max_length > MAX_TEXT_LENGTH {
            return Err(corrupted("FFData cch exceeds 32767"));
        }
        if kind == FormFieldDataKind::CheckBox
            && !(MIN_CHECKBOX_HPS..=MAX_CHECKBOX_HPS).contains(&size_half_points)
        {
            return Err(corrupted("FFData hps is outside the checkbox range"));
        }

        let name = cursor.read_xstz("xstzName", MAX_NAME_CHARS)?;
        let mut default_text = None;
        let mut default_state = None;
        match kind {
            FormFieldDataKind::Text => {
                let text = cursor.read_xstz("xstzTextDef", MAX_DEFAULT_TEXT_CHARS)?;
                if matches!(
                    text_kind,
                    FormFieldTextKind::CurrentDate | FormFieldTextKind::CurrentTime
                ) && !text.is_empty()
                {
                    return Err(corrupted(
                        "xstzTextDef is not empty for a current date/time text box",
                    ));
                }
                default_text = Some(text);
            },
            FormFieldDataKind::CheckBox | FormFieldDataKind::DropDown => {
                let w_def = cursor.read_u16("FFData wDef")?;
                if kind == FormFieldDataKind::CheckBox && w_def > 1 {
                    return Err(corrupted("FFData wDef is not a checkbox state"));
                }
                default_state = Some(w_def);
            },
        }
        let text_format = cursor.read_xstz("xstzTextFormat", MAX_TEXT_FORMAT_CHARS)?;
        if kind != FormFieldDataKind::Text && !text_format.is_empty() {
            return Err(corrupted(
                "xstzTextFormat is not empty for a non-text field",
            ));
        }
        let help_text = cursor.read_xstz("xstzHelpText", MAX_HELP_TEXT_CHARS)?;
        let status_text = cursor.read_xstz("xstzStatText", MAX_STATUS_TEXT_CHARS)?;
        let entry_macro = cursor.read_xstz("xstzEntryMcr", MAX_MACRO_NAME_CHARS)?;
        let exit_macro = cursor.read_xstz("xstzExitMcr", MAX_MACRO_NAME_CHARS)?;
        let items = if kind == FormFieldDataKind::DropDown {
            let items = cursor.read_sttb("hsttbDropList")?;
            let w_def = default_state.expect("dropdown reads wDef");
            if u32::from(w_def) >= items.len() as u32 {
                return Err(corrupted("FFData wDef is past the drop-down item list"));
            }
            Some(items)
        } else {
            None
        };
        if !cursor.is_finished() {
            return Err(corrupted("FFData has trailing bytes"));
        }

        Ok(Self {
            kind,
            state,
            text_kind,
            own_help,
            own_status,
            protected,
            auto_size,
            recalculate,
            has_list_box,
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

    /// Parse the `FFData` of the `NilPICFAndBinData` stored at `offset` in
    /// the Data stream.
    pub fn parse_at(data_stream: &[u8], offset: u32) -> Result<Self> {
        Self::parse(NilPicfAndBinData::parse_at(data_stream, offset)?.bin_data())
    }

    /// Re-encode the `FFData`, reproducing the original bytes for well-formed
    /// input.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut bits = match self.kind {
            FormFieldDataKind::Text => 0,
            FormFieldDataKind::CheckBox => 1,
            FormFieldDataKind::DropDown => 2,
        };
        bits |= u16::from(self.state) << IRES_SHIFT;
        if self.own_help {
            bits |= F_OWN_HELP;
        }
        if self.own_status {
            bits |= F_OWN_STAT;
        }
        if self.protected {
            bits |= F_PROT;
        }
        if self.auto_size {
            bits |= F_ISIZE;
        }
        if self.kind == FormFieldDataKind::Text {
            let text_kind = match self.text_kind {
                FormFieldTextKind::Regular => 0,
                FormFieldTextKind::Number => 1,
                FormFieldTextKind::Date => 2,
                FormFieldTextKind::CurrentDate => 3,
                FormFieldTextKind::CurrentTime => 4,
                FormFieldTextKind::Calculation => 5,
            };
            bits |= text_kind << ITYPE_TXT_SHIFT;
        }
        if self.recalculate {
            bits |= F_RECALC;
        }
        if self.has_list_box {
            bits |= F_HAS_LISTBOX;
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

    /// The stored kind of the form field (`FFDataBits.iType`).
    pub const fn kind(&self) -> FormFieldDataKind {
        self.kind
    }

    /// The stored text-box value kind (`FFDataBits.iTypeTxt`), or `None` when
    /// this is not a text box.
    pub const fn text_kind(&self) -> Option<FormFieldTextKind> {
        match self.kind {
            FormFieldDataKind::Text => Some(self.text_kind),
            _ => None,
        }
    }

    /// The stored maximum length, in characters, of the text-box value
    /// (`FFData.cch`). Zero means unlimited. Always `None` for non-text kinds.
    pub const fn max_length(&self) -> Option<u16> {
        match self.kind {
            FormFieldDataKind::Text => Some(self.max_length),
            _ => None,
        }
    }

    /// The stored checkbox state, or `None` when this is not a check box.
    pub const fn checkbox_state(&self) -> Option<CheckBoxState> {
        match self.kind {
            FormFieldDataKind::CheckBox => Some(match self.state {
                0 => CheckBoxState::Unchecked,
                1 => CheckBoxState::Checked,
                _ => CheckBoxState::Undefined,
            }),
            _ => None,
        }
    }

    /// The stored default checkbox state (`wDef`), or `None` when this is not
    /// a check box.
    pub const fn is_checked_by_default(&self) -> Option<bool> {
        match self.kind {
            FormFieldDataKind::CheckBox => match self.default_state {
                Some(1) => Some(true),
                _ => Some(false),
            },
            _ => None,
        }
    }

    /// The stored zero-based default selected item (`wDef`), or `None` when
    /// this is not a drop-down list.
    pub const fn default_item_index(&self) -> Option<u16> {
        match self.kind {
            FormFieldDataKind::DropDown => self.default_state,
            _ => None,
        }
    }

    /// The stored zero-based selected item (`iRes`), or `None` when this is
    /// not a drop-down list or the stored selection is undefined.
    pub const fn selected_item_index(&self) -> Option<u8> {
        match self.kind {
            FormFieldDataKind::DropDown => match self.state {
                UNDEFINED_STATE => None,
                index => Some(index),
            },
            _ => None,
        }
    }

    /// The stored check-box size in half-points (`FFData.hps`), or `None`
    /// when this is not a check box.
    pub const fn checkbox_size_half_points(&self) -> Option<u16> {
        match self.kind {
            FormFieldDataKind::CheckBox => Some(self.size_half_points),
            _ => None,
        }
    }

    /// Whether the stored properties size the check box from the surrounding
    /// text size (`FFDataBits.iSize`). Always `false` for non-checkbox kinds.
    pub const fn is_checkbox_auto_sized(&self) -> bool {
        self.auto_size
    }

    /// The stored drop-down list entries (`hsttbDropList`). Empty when this
    /// is not a drop-down list.
    pub fn dropdown_items(&self) -> &[String] {
        self.items.as_deref().unwrap_or(&[])
    }

    /// The stored name of the form field (`xstzName`).
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The stored default text of the text box (`xstzTextDef`), or `None`
    /// when this is not a text box. For a calculated text box this is the
    /// stored expression; it is inert and never evaluated.
    pub fn default_text(&self) -> Option<&str> {
        self.default_text.as_deref()
    }

    /// The stored text-box format string (`xstzTextFormat`). Empty for
    /// non-text kinds.
    pub fn text_format(&self) -> &str {
        &self.text_format
    }

    /// The stored help text (`xstzHelpText`).
    pub fn help_text(&self) -> &str {
        &self.help_text
    }

    /// The stored status bar text (`xstzStatText`).
    pub fn status_text(&self) -> &str {
        &self.status_text
    }

    /// The stored entry macro name (`xstzEntryMcr`).
    ///
    /// This name is inert metadata: it is never resolved, loaded, or invoked.
    pub fn entry_macro(&self) -> &str {
        &self.entry_macro
    }

    /// The stored exit macro name (`xstzExitMcr`).
    ///
    /// This name is inert metadata: it is never resolved, loaded, or invoked.
    pub fn exit_macro(&self) -> &str {
        &self.exit_macro
    }

    /// Whether the stored properties mark the help text as custom
    /// (`FFDataBits.fOwnHelp`).
    pub const fn has_own_help_text(&self) -> bool {
        self.own_help
    }

    /// Whether the stored properties mark the status bar text as custom
    /// (`FFDataBits.fOwnStat`).
    pub const fn has_own_status_text(&self) -> bool {
        self.own_status
    }

    /// Whether the stored properties protect the field value from changes
    /// (`FFDataBits.fProt`).
    pub const fn is_protected(&self) -> bool {
        self.protected
    }

    /// Whether the stored properties mark the value for automatic
    /// recalculation (`FFDataBits.fRecalc`). This flag is retained verbatim;
    /// nothing is ever recalculated.
    pub const fn is_marked_for_recalculation(&self) -> bool {
        self.recalculate
    }

    /// Whether the stored properties mark the field as having a list box
    /// (`FFDataBits.fHasListBox`).
    pub const fn has_list_box(&self) -> bool {
        self.has_list_box
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
            .ok_or_else(|| corrupted(format!("{what} is truncated")))?;
        self.offset += 2;
        Ok(u16::from_le_bytes(
            bytes.try_into().expect("length checked"),
        ))
    }

    fn read_u32(&mut self, what: &str) -> Result<u32> {
        let bytes = self
            .data
            .get(self.offset..self.offset + 4)
            .ok_or_else(|| corrupted(format!("{what} is truncated")))?;
        self.offset += 4;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("length checked"),
        ))
    }

    /// Read an `Xstz` (MS-DOC 2.9.354): `cch` UTF-16 code units plus a
    /// mandated zero terminator.
    fn read_xstz(&mut self, what: &str, max_chars: u16) -> Result<String> {
        let cch = self.read_u16(what)?;
        if cch > max_chars {
            return Err(corrupted(format!("{what} exceeds its length cap")));
        }
        let byte_len = usize::from(cch) * 2;
        let bytes = self
            .data
            .get(self.offset..self.offset + byte_len)
            .ok_or_else(|| corrupted(format!("{what} is truncated")))?;
        self.offset += byte_len;
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes(chunk.try_into().expect("length checked")))
            .collect();
        let text = String::from_utf16(&units)
            .map_err(|_| corrupted(format!("{what} is not valid UTF-16")))?;
        let terminator = self.read_u16(what)?;
        if terminator != 0 {
            return Err(corrupted(format!("{what} is not null-terminated")));
        }
        Ok(text)
    }

    /// Read an extended `STTB` of Unicode strings with no extra data (the
    /// `hsttbDropList` layout, MS-DOC 2.9.78).
    fn read_sttb(&mut self, what: &str) -> Result<Vec<String>> {
        if self.read_u16(what)? != STTB_EXTENDED {
            return Err(corrupted(format!("{what} is not an extended STTB")));
        }
        let count = self.read_u16(what)?;
        if count > MAX_DROPDOWN_ITEMS {
            return Err(corrupted(format!("{what} exceeds 25 elements")));
        }
        if self.read_u16(what)? != 0 {
            return Err(corrupted(format!("{what} carries unexpected extra data")));
        }
        let mut items = Vec::with_capacity(usize::from(count));
        for _ in 0..count {
            let cch = usize::from(self.read_u16(what)?);
            let byte_len = cch * 2;
            let bytes = self
                .data
                .get(self.offset..self.offset + byte_len)
                .ok_or_else(|| corrupted(format!("{what} is truncated")))?;
            self.offset += byte_len;
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes(chunk.try_into().expect("length checked")))
                .collect();
            items.push(
                String::from_utf16(&units)
                    .map_err(|_| corrupted(format!("{what} is not valid UTF-16")))?,
            );
        }
        Ok(items)
    }
}

/// Write an `Xstz`: `cch`, the UTF-16 code units, and a zero terminator.
fn write_xstz(out: &mut Vec<u8>, text: &str) {
    let units: Vec<u16> = text.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    for unit in units {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out.extend_from_slice(&0u16.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    fn xstz(text: &str) -> Vec<u8> {
        let mut out = Vec::new();
        write_xstz(&mut out, text);
        out
    }

    fn nil_picf(bin_data: &[u8]) -> Vec<u8> {
        let lcb = (HEADER_LEN + bin_data.len()) as u32;
        let mut out = Vec::new();
        out.extend_from_slice(&lcb.to_le_bytes());
        out.extend_from_slice(&CB_HEADER.to_le_bytes());
        out.extend_from_slice(&[0; HEADER_LEN - 6]);
        out.extend_from_slice(bin_data);
        out
    }

    fn text_ff_data(bits: u16, name: &str, default: &str, format: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&FF_DATA_VERSION.to_le_bytes());
        out.extend_from_slice(&bits.to_le_bytes());
        out.extend_from_slice(&10u16.to_le_bytes()); // cch
        out.extend_from_slice(&0u16.to_le_bytes()); // hps
        out.extend_from_slice(&xstz(name));
        out.extend_from_slice(&xstz(default));
        out.extend_from_slice(&xstz(format));
        out.extend_from_slice(&xstz("help"));
        out.extend_from_slice(&xstz("status"));
        out.extend_from_slice(&xstz("EntryMacro"));
        out.extend_from_slice(&xstz("ExitMacro"));
        out
    }

    fn checkbox_ff_data(state: u16, w_def: u16) -> Vec<u8> {
        let bits = 1 | (state << IRES_SHIFT) | F_ISIZE;
        let mut out = Vec::new();
        out.extend_from_slice(&FF_DATA_VERSION.to_le_bytes());
        out.extend_from_slice(&bits.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // cch MUST be 0
        out.extend_from_slice(&20u16.to_le_bytes()); // hps
        out.extend_from_slice(&xstz("Check1"));
        out.extend_from_slice(&w_def.to_le_bytes());
        for text in ["", "help", "status", "", ""] {
            out.extend_from_slice(&xstz(text));
        }
        out
    }

    fn dropdown_ff_data(selection: u16, w_def: u16, items: &[&str]) -> Vec<u8> {
        let bits = 2 | (selection << IRES_SHIFT) | F_HAS_LISTBOX;
        let mut out = Vec::new();
        out.extend_from_slice(&FF_DATA_VERSION.to_le_bytes());
        out.extend_from_slice(&bits.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(&xstz("Drop1"));
        out.extend_from_slice(&w_def.to_le_bytes());
        for text in ["", "", "", "", ""] {
            out.extend_from_slice(&xstz(text));
        }
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
        out
    }

    #[test]
    fn parses_text_field() {
        // iType text, iTypeTxt date, fRecalc.
        let bytes = text_ff_data(
            2 << ITYPE_TXT_SHIFT | F_RECALC,
            "Text1",
            "1.1.2000",
            "d.M.yyyy",
        );
        let data = FormFieldData::parse(&bytes).unwrap();
        assert_eq!(data.kind(), FormFieldDataKind::Text);
        assert_eq!(data.text_kind(), Some(FormFieldTextKind::Date));
        assert_eq!(data.max_length(), Some(10));
        assert_eq!(data.name(), "Text1");
        assert_eq!(data.default_text(), Some("1.1.2000"));
        assert_eq!(data.text_format(), "d.M.yyyy");
        assert_eq!(data.help_text(), "help");
        assert_eq!(data.status_text(), "status");
        assert_eq!(data.entry_macro(), "EntryMacro");
        assert_eq!(data.exit_macro(), "ExitMacro");
        assert!(data.is_marked_for_recalculation());
        assert!(!data.is_protected());
        assert_eq!(data.checkbox_state(), None);
        assert_eq!(data.selected_item_index(), None);
        assert!(data.dropdown_items().is_empty());
        assert_eq!(data.to_bytes(), bytes);
    }

    #[test]
    fn parses_checkbox_field() {
        let bytes = checkbox_ff_data(1, 1);
        let data = FormFieldData::parse(&bytes).unwrap();
        assert_eq!(data.kind(), FormFieldDataKind::CheckBox);
        assert_eq!(data.checkbox_state(), Some(CheckBoxState::Checked));
        assert_eq!(data.is_checked_by_default(), Some(true));
        assert_eq!(data.checkbox_size_half_points(), Some(20));
        assert!(data.is_checkbox_auto_sized());
        assert_eq!(data.text_kind(), None);
        assert_eq!(data.default_text(), None);
        assert_eq!(data.to_bytes(), bytes);

        let undefined = FormFieldData::parse(&checkbox_ff_data(UNDEFINED_STATE.into(), 0)).unwrap();
        assert_eq!(undefined.checkbox_state(), Some(CheckBoxState::Undefined));
        assert_eq!(undefined.is_checked_by_default(), Some(false));
    }

    #[test]
    fn parses_dropdown_field() {
        let bytes = dropdown_ff_data(2, 0, &["one", "two", "three"]);
        let data = FormFieldData::parse(&bytes).unwrap();
        assert_eq!(data.kind(), FormFieldDataKind::DropDown);
        assert!(data.has_list_box());
        assert_eq!(data.selected_item_index(), Some(2));
        assert_eq!(data.default_item_index(), Some(0));
        assert_eq!(data.dropdown_items(), &["one", "two", "three"]);
        assert_eq!(data.name(), "Drop1");
        assert_eq!(data.to_bytes(), bytes);

        let undefined =
            FormFieldData::parse(&dropdown_ff_data(UNDEFINED_STATE.into(), 1, &["a", "b"]))
                .unwrap();
        assert_eq!(undefined.selected_item_index(), None);
        assert_eq!(undefined.default_item_index(), Some(1));
    }

    #[test]
    fn rejects_invalid_text_kinds() {
        // Reserved iType (3).
        assert!(FormFieldData::parse(&text_ff_data(3, "n", "", "")).is_err());
        // iRes not 0 for a text box.
        assert!(FormFieldData::parse(&text_ff_data(1 << IRES_SHIFT, "n", "", "")).is_err());
        // Reserved iTypeTxt (6) for a text box.
        assert!(FormFieldData::parse(&text_ff_data(6 << ITYPE_TXT_SHIFT, "n", "", "")).is_err());
        // fHasListBox set on a text box.
        assert!(FormFieldData::parse(&text_ff_data(F_HAS_LISTBOX, "n", "", "")).is_err());
        // iSize set on a text box.
        assert!(FormFieldData::parse(&text_ff_data(F_ISIZE, "n", "", "")).is_err());
        // Non-empty default text for a current-date text box.
        assert!(FormFieldData::parse(&text_ff_data(3 << ITYPE_TXT_SHIFT, "n", "x", "")).is_err());
    }

    #[test]
    fn rejects_invalid_checkbox_and_dropdown_states() {
        // iRes 2 is not a checkbox state.
        assert!(FormFieldData::parse(&checkbox_ff_data(2, 0)).is_err());
        // wDef 2 is not a checkbox state.
        assert!(FormFieldData::parse(&checkbox_ff_data(0, 2)).is_err());
        // hps below the checkbox range.
        let mut bad_hps = checkbox_ff_data(0, 0);
        bad_hps[8..10].copy_from_slice(&1u16.to_le_bytes());
        assert!(FormFieldData::parse(&bad_hps).is_err());
        // iRes 26 exceeds the undefined-selection marker.
        assert!(FormFieldData::parse(&dropdown_ff_data(26, 0, &["a"])).is_err());
        // wDef past the item list.
        assert!(FormFieldData::parse(&dropdown_ff_data(0, 3, &["a"])).is_err());
        // Missing fHasListBox on a drop-down.
        let mut no_listbox = dropdown_ff_data(0, 0, &["a"]);
        no_listbox[4..6].copy_from_slice(&0u16.to_le_bytes());
        assert!(FormFieldData::parse(&no_listbox).is_err());
        // More than 25 items.
        let items: Vec<&str> = vec!["x"; 26];
        assert!(FormFieldData::parse(&dropdown_ff_data(0, 0, &items)).is_err());
        // Non-zero cch for a non-text field.
        let mut bad_cch = checkbox_ff_data(0, 0);
        bad_cch[6..8].copy_from_slice(&5u16.to_le_bytes());
        assert!(FormFieldData::parse(&bad_cch).is_err());
    }

    #[test]
    fn rejects_malformed_payloads() {
        let good = text_ff_data(0, "Text1", "", "");
        // Wrong version.
        let mut bad_version = good.clone();
        bad_version[0..4].copy_from_slice(&0u32.to_le_bytes());
        assert!(FormFieldData::parse(&bad_version).is_err());
        // Truncated.
        assert!(FormFieldData::parse(&good[..good.len() - 1]).is_err());
        assert!(FormFieldData::parse(&good[..4]).is_err());
        // Trailing bytes.
        let mut trailing = good.clone();
        trailing.push(0);
        assert!(FormFieldData::parse(&trailing).is_err());
        // xstzName beyond its 20-character cap.
        let long_name = text_ff_data(0, &"n".repeat(21), "", "");
        assert!(FormFieldData::parse(&long_name).is_err());
        // Non-zero Xstz terminator.
        let mut bad_terminator = good.clone();
        // xstzName "Text1": cch at 10, units at 12..22, terminator at 22.
        bad_terminator[22..24].copy_from_slice(&1u16.to_le_bytes());
        assert!(FormFieldData::parse(&bad_terminator).is_err());
        // Lone surrogate in an Xstz.
        let mut bad_utf16 = good.clone();
        bad_utf16[12..14].copy_from_slice(&0xD800u16.to_le_bytes());
        assert!(FormFieldData::parse(&bad_utf16).is_err());
        // Non-text field with a non-empty format string.
        let mut bad_format = checkbox_ff_data(0, 0);
        // name "Check1" ends at 10+2+12+2=26, wDef 26..28, format cch at 28.
        bad_format[28..30].copy_from_slice(&1u16.to_le_bytes());
        assert!(FormFieldData::parse(&bad_format).is_err());
    }

    #[test]
    fn parses_nil_picf_and_bin_data() {
        let ff_data = text_ff_data(0, "Text1", "", "");
        let bytes = nil_picf(&ff_data);
        let parsed = NilPicfAndBinData::parse(&bytes).unwrap();
        assert_eq!(parsed.bin_data(), ff_data.as_slice());
        assert_eq!(parsed.to_bytes(), bytes);

        // Parsing from a longer buffer consumes exactly lcb bytes.
        let mut padded = bytes.clone();
        padded.extend_from_slice(&[0xAA; 16]);
        let parsed = NilPicfAndBinData::parse(&padded).unwrap();
        assert_eq!(parsed.bin_data(), ff_data.as_slice());

        // Parsing at an offset of a Data-stream-like buffer.
        let mut stream = vec![0u8; 7];
        stream.extend_from_slice(&bytes);
        let parsed = NilPicfAndBinData::parse_at(&stream, 7).unwrap();
        assert_eq!(parsed.bin_data(), ff_data.as_slice());
        let data = FormFieldData::parse_at(&stream, 7).unwrap();
        assert_eq!(data.name(), "Text1");
    }

    #[test]
    fn rejects_malformed_nil_picf_and_bin_data() {
        let ff_data = text_ff_data(0, "Text1", "", "");
        let good = nil_picf(&ff_data);
        // Truncated header.
        assert!(NilPicfAndBinData::parse(&good[..HEADER_LEN - 1]).is_err());
        // Wrong cbHeader.
        let mut bad_header = good.clone();
        bad_header[4..6].copy_from_slice(&0x0042u16.to_le_bytes());
        assert!(NilPicfAndBinData::parse(&bad_header).is_err());
        // lcb smaller than the header.
        let mut small_lcb = good.clone();
        small_lcb[0..4].copy_from_slice(&10i32.to_le_bytes());
        assert!(NilPicfAndBinData::parse(&small_lcb).is_err());
        // lcb past the containing data.
        let mut huge_lcb = good.clone();
        huge_lcb[0..4].copy_from_slice(&1_000_000i32.to_le_bytes());
        assert!(NilPicfAndBinData::parse(&huge_lcb).is_err());
        // Negative lcb.
        let mut negative_lcb = good.clone();
        negative_lcb[0..4].copy_from_slice(&(-1i32).to_le_bytes());
        assert!(NilPicfAndBinData::parse(&negative_lcb).is_err());
        // Offset past the stream.
        assert!(NilPicfAndBinData::parse_at(&good, 1_000_000).is_err());
    }
}
