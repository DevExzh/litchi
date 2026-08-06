//! MS-DOC form-field invariants and bounded semantic validation.

use crate::package::{Error as PackageError, Result};

use super::model::{FormFieldDataKind, FormFieldTextKind};

// FFDataBits layout (MS-DOC 2.9.79), least-significant bit first.
pub(crate) const ITYPE_MASK: u16 = 0x0003;
pub(crate) const IRES_SHIFT: u16 = 2;
const IRES_MASK: u16 = 0x001F;
pub(crate) const F_OWN_HELP: u16 = 0x0080;
pub(crate) const F_OWN_STAT: u16 = 0x0100;
pub(crate) const F_PROT: u16 = 0x0200;
pub(crate) const F_ISIZE: u16 = 0x0400;
pub(crate) const ITYPE_TXT_SHIFT: u16 = 11;
const ITYPE_TXT_MASK: u16 = 0x0007;
pub(crate) const F_RECALC: u16 = 0x4000;
pub(crate) const F_HAS_LISTBOX: u16 = 0x8000;

/// FFData.cch upper bound: the maximum text-box value length.
pub(crate) const MAX_TEXT_LENGTH: u16 = 32767;
/// FFData.hps bounds for a check box (in half-points).
const MIN_CHECKBOX_HPS: u16 = 2;
const MAX_CHECKBOX_HPS: u16 = 3168;
/// xstzName.cch upper bound.
pub(crate) const MAX_NAME_CHARS: u16 = 20;
/// xstzTextDef.cch upper bound.
pub(crate) const MAX_DEFAULT_TEXT_CHARS: u16 = 255;
/// xstzTextFormat.cch upper bound.
pub(crate) const MAX_TEXT_FORMAT_CHARS: u16 = 64;
/// xstzHelpText.cch upper bound.
pub(crate) const MAX_HELP_TEXT_CHARS: u16 = 255;
/// xstzStatText.cch upper bound.
pub(crate) const MAX_STATUS_TEXT_CHARS: u16 = 138;
/// xstzEntryMcr.cch and xstzExitMcr.cch upper bound.
pub(crate) const MAX_MACRO_NAME_CHARS: u16 = 32;
/// hsttbDropList element count upper bound.
pub(crate) const MAX_DROPDOWN_ITEMS: u16 = 25;

pub(crate) fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

/// The semantically decoded FFDataBits fields used by the wire codec.
#[derive(Debug, Clone, Copy)]
pub(crate) struct DecodedBits {
    pub(crate) kind: FormFieldDataKind,
    pub(crate) state: u8,
    pub(crate) text_kind: FormFieldTextKind,
    pub(crate) own_help: bool,
    pub(crate) own_status: bool,
    pub(crate) protected: bool,
    pub(crate) auto_size: bool,
    pub(crate) recalculate: bool,
    pub(crate) has_list_box: bool,
}

/// Decode and validate the FFDataBits relationships mandated by MS-DOC.
pub(crate) fn decode_bits(
    bits: u16,
    max_length: u16,
    size_half_points: u16,
) -> Result<DecodedBits> {
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

    match kind {
        FormFieldDataKind::Text => {
            if state != 0 {
                return Err(corrupted("FFDataBits iRes is not 0 for a text box"));
            }
        },
        FormFieldDataKind::CheckBox => {
            if !matches!(state, 0 | 1 | super::model::UNDEFINED_STATE) {
                return Err(corrupted("FFDataBits iRes is not a checkbox state"));
            }
        },
        FormFieldDataKind::DropDown => {
            if state > super::model::UNDEFINED_STATE {
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

    Ok(DecodedBits {
        kind,
        state,
        text_kind,
        own_help,
        own_status,
        protected,
        auto_size,
        recalculate,
        has_list_box,
    })
}

/// Validate the kind-specific stored default value.
pub(crate) fn validate_default_state(kind: FormFieldDataKind, w_def: u16) -> Result<()> {
    if kind == FormFieldDataKind::CheckBox && w_def > 1 {
        return Err(corrupted("FFData wDef is not a checkbox state"));
    }
    Ok(())
}

/// Validate the text-box default against its date/time semantics.
pub(crate) fn validate_text_default(text_kind: FormFieldTextKind, text: &str) -> Result<()> {
    if matches!(
        text_kind,
        FormFieldTextKind::CurrentDate | FormFieldTextKind::CurrentTime
    ) && !text.is_empty()
    {
        return Err(corrupted(
            "xstzTextDef is not empty for a current date/time text box",
        ));
    }
    Ok(())
}

/// Validate the format string, which is empty for non-text fields.
pub(crate) fn validate_text_format(kind: FormFieldDataKind, text_format: &str) -> Result<()> {
    if kind != FormFieldDataKind::Text && !text_format.is_empty() {
        return Err(corrupted(
            "xstzTextFormat is not empty for a non-text field",
        ));
    }
    Ok(())
}

/// Validate a drop-down default against the inline item list.
pub(crate) fn validate_dropdown_default(w_def: u16, item_count: usize) -> Result<()> {
    if u32::from(w_def) >= item_count as u32 {
        return Err(corrupted("FFData wDef is past the drop-down item list"));
    }
    Ok(())
}
