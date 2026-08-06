//! Typed, contextualized models for legacy MS-DOC form-field metadata.

/// The stored kind of a legacy form field (FFDataBits.iType, MS-DOC 2.9.79).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormFieldDataKind {
    /// iTypeText (0): a text box.
    Text,
    /// iTypeChck (1): a check box.
    CheckBox,
    /// iTypeDrop (2): a drop-down list box.
    DropDown,
}

/// The stored text-box value kind (FFDataBits.iTypeTxt, MS-DOC 2.9.79).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormFieldTextKind {
    /// iTypeTxtReg (0): regular text.
    Regular,
    /// iTypeTxtNum (1): a number.
    Number,
    /// iTypeTxtDate (2): a date or time.
    Date,
    /// iTypeTxtCurDate (3): the current date.
    CurrentDate,
    /// iTypeTxtCurTime (4): the current time.
    CurrentTime,
    /// iTypeTxtCalc (5): calculated from the stored default-text expression.
    Calculation,
}

/// The stored state of a check-box form field.
///
/// Undefined check boxes are treated as unchecked by Word; this type only
/// retains the stored state and never changes it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CheckBoxState {
    /// The check box is stored as unchecked (iRes/wDef 0).
    Unchecked,
    /// The check box is stored as checked (iRes/wDef 1).
    Checked,
    /// The stored checkbox state is undefined (iRes 25).
    Undefined,
}

/// iRes value marking an undefined checkbox state or drop-down selection.
pub(crate) const UNDEFINED_STATE: u8 = 25;

/// A parsed NilPICFAndBinData wrapper (MS-DOC 2.9.158).
///
/// Only the cbHeader == 0x0044 layout is accepted. The 62 ignored header
/// bytes are not retained; to_bytes re-emits them as zero, reproducing the
/// original bytes for well-formed input where they MUST be zero.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NilPicfAndBinData {
    pub(crate) bin_data: Vec<u8>,
}

impl NilPicfAndBinData {
    /// The stored binary payload (binData).
    pub fn bin_data(&self) -> &[u8] {
        &self.bin_data
    }
}

/// Typed, inert form-field data (FFData, MS-DOC 2.9.78).
///
/// All values are stored state only: entry and exit macro names are retained
/// verbatim and never invoked, the form is never filled, checkbox and
/// selection states are never changed, and no field is refreshed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormFieldData {
    pub(crate) kind: FormFieldDataKind,
    /// Raw 5-bit iRes: checkbox state or drop-down selection (zero when the
    /// kind is text, where the spec mandates 0).
    pub(crate) state: u8,
    pub(crate) text_kind: FormFieldTextKind,
    pub(crate) own_help: bool,
    pub(crate) own_status: bool,
    pub(crate) protected: bool,
    pub(crate) auto_size: bool,
    pub(crate) recalculate: bool,
    pub(crate) has_list_box: bool,
    /// FFData.cch: maximum text-box value length (0 means unlimited; always
    /// 0 for non-text kinds).
    pub(crate) max_length: u16,
    /// FFData.hps: check-box size in half-points (meaningful for check boxes
    /// only).
    pub(crate) size_half_points: u16,
    pub(crate) name: String,
    /// xstzTextDef: stored default text (text boxes only).
    pub(crate) default_text: Option<String>,
    /// wDef: stored default checkbox state (0/1) or drop-down index.
    pub(crate) default_state: Option<u16>,
    pub(crate) text_format: String,
    pub(crate) help_text: String,
    pub(crate) status_text: String,
    pub(crate) entry_macro: String,
    pub(crate) exit_macro: String,
    /// hsttbDropList entries (drop-down lists only).
    pub(crate) items: Option<Vec<String>>,
}

impl FormFieldData {
    /// The stored kind of the form field (FFDataBits.iType).
    pub const fn kind(&self) -> FormFieldDataKind {
        self.kind
    }

    /// The stored text-box value kind (FFDataBits.iTypeTxt), or None when
    /// this is not a text box.
    pub const fn text_kind(&self) -> Option<FormFieldTextKind> {
        match self.kind {
            FormFieldDataKind::Text => Some(self.text_kind),
            _ => None,
        }
    }

    /// The stored maximum length, in characters, of the text-box value
    /// (FFData.cch). Zero means unlimited. Always None for non-text kinds.
    pub const fn max_length(&self) -> Option<u16> {
        match self.kind {
            FormFieldDataKind::Text => Some(self.max_length),
            _ => None,
        }
    }

    /// The stored checkbox state, or None when this is not a check box.
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

    /// The stored default checkbox state (wDef), or None when this is not a
    /// check box.
    pub const fn is_checked_by_default(&self) -> Option<bool> {
        match self.kind {
            FormFieldDataKind::CheckBox => match self.default_state {
                Some(1) => Some(true),
                _ => Some(false),
            },
            _ => None,
        }
    }

    /// The stored zero-based default selected item (wDef), or None when this
    /// is not a drop-down list.
    pub const fn default_item_index(&self) -> Option<u16> {
        match self.kind {
            FormFieldDataKind::DropDown => self.default_state,
            _ => None,
        }
    }

    /// The stored zero-based selected item (iRes), or None when this is not a
    /// drop-down list or the stored selection is undefined.
    pub const fn selected_item_index(&self) -> Option<u8> {
        match self.kind {
            FormFieldDataKind::DropDown => match self.state {
                UNDEFINED_STATE => None,
                index => Some(index),
            },
            _ => None,
        }
    }

    /// The stored check-box size in half-points (FFData.hps), or None when
    /// this is not a check box.
    pub const fn checkbox_size_half_points(&self) -> Option<u16> {
        match self.kind {
            FormFieldDataKind::CheckBox => Some(self.size_half_points),
            _ => None,
        }
    }

    /// Whether the stored properties size the check box from the surrounding
    /// text size (FFDataBits.iSize). Always false for non-checkbox kinds.
    pub const fn is_checkbox_auto_sized(&self) -> bool {
        self.auto_size
    }

    /// The stored drop-down list entries (hsttbDropList). Empty when this is
    /// not a drop-down list.
    pub fn dropdown_items(&self) -> &[String] {
        self.items.as_deref().unwrap_or(&[])
    }

    /// The stored name of the form field (xstzName).
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The stored default text of the text box (xstzTextDef), or None when
    /// this is not a text box. For a calculated text box this is the stored
    /// expression; it is inert and never evaluated.
    pub fn default_text(&self) -> Option<&str> {
        self.default_text.as_deref()
    }

    /// The stored text-box format string (xstzTextFormat). Empty for
    /// non-text kinds.
    pub fn text_format(&self) -> &str {
        &self.text_format
    }

    /// The stored help text (xstzHelpText).
    pub fn help_text(&self) -> &str {
        &self.help_text
    }

    /// The stored status bar text (xstzStatText).
    pub fn status_text(&self) -> &str {
        &self.status_text
    }

    /// The stored entry macro name (xstzEntryMcr).
    ///
    /// This name is inert metadata: it is never resolved, loaded, or invoked.
    pub fn entry_macro(&self) -> &str {
        &self.entry_macro
    }

    /// The stored exit macro name (xstzExitMcr).
    ///
    /// This name is inert metadata: it is never resolved, loaded, or invoked.
    pub fn exit_macro(&self) -> &str {
        &self.exit_macro
    }

    /// Whether the stored properties mark the help text as custom
    /// (FFDataBits.fOwnHelp).
    pub const fn has_own_help_text(&self) -> bool {
        self.own_help
    }

    /// Whether the stored properties mark the status bar text as custom
    /// (FFDataBits.fOwnStat).
    pub const fn has_own_status_text(&self) -> bool {
        self.own_status
    }

    /// Whether the stored properties protect the field value from changes
    /// (FFDataBits.fProt).
    pub const fn is_protected(&self) -> bool {
        self.protected
    }

    /// Whether the stored properties mark the value for automatic
    /// recalculation (FFDataBits.fRecalc). This flag is retained verbatim;
    /// nothing is ever recalculated.
    pub const fn is_marked_for_recalculation(&self) -> bool {
        self.recalculate
    }

    /// Whether the stored properties mark the field as having a list box
    /// (FFDataBits.fHasListBox).
    pub const fn has_list_box(&self) -> bool {
        self.has_list_box
    }
}
