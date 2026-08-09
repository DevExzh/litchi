//! Typed BIFF8 defined-name values and the internal `Lbl` slot model.

/// A non-macro internal defined name from the workbook globals substream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedName {
    /// One-based position in the complete `Lbl` collection, including macro slots.
    pub record_index: u32,
    pub name: String,
    pub scope: NameScope,
    pub hidden: bool,
    pub function: bool,
    pub vba_procedure: bool,
    pub procedure: bool,
    pub calculated_expression: bool,
    pub function_group: u8,
    pub published: bool,
    pub workbook_parameter: bool,
    pub shortcut_key: Option<u8>,
    pub kind: DefinedNameKind,
    /// Rendered formula using the same leading-`=` convention as cell formulas.
    pub formula: Option<String>,
    /// Original `NameParsedFormula.rgce` bytes.
    pub formula_tokens: Vec<u8>,
    pub formula_extra: Vec<u8>,
    pub continuation_chunks: Vec<Vec<u8>>,
    pub custom_menu: String,
    pub description: String,
    pub help_topic: String,
    pub status_bar: String,
    pub comment: Option<String>,
    pub future_records: DefinedNameFutureRecords,
}

/// Optional BIFF8 future records associated with one immediately preceding `Lbl`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DefinedNameFutureRecords {
    pub function_group: Option<NameFnGrp12>,
    pub publication: Option<NamePublish>,
}

/// Extended function-category metadata from `NameFnGrp12`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NameFnGrp12 {
    pub function_name: String,
    /// Raw BIFF category number in the inclusive range 32..=255.
    pub category: u8,
}

impl NameFnGrp12 {
    #[must_use]
    pub fn category_index(&self) -> usize {
        usize::from(self.category - 32)
    }
}

/// Server publication metadata from `NamePublish`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamePublish {
    pub published: bool,
    pub workbook_parameter: bool,
    pub name: String,
}

impl DefinedName {
    #[must_use]
    pub fn is_macro(&self) -> bool {
        self.function || self.vba_procedure || self.procedure
    }

    /// Whether the rendered definition contains a deleted reference.
    #[must_use]
    pub fn is_deleted(&self) -> bool {
        self.formula
            .as_deref()
            .is_some_and(|formula| formula.contains("#REF!"))
    }
}

/// Scope of a defined name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NameScope {
    Workbook,
    Worksheet(usize),
}

/// User-defined or reserved built-in name kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DefinedNameKind {
    User,
    BuiltIn(BuiltInName),
}

/// Built-in BIFF8 defined-name identifiers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuiltInName {
    ConsolidateArea,
    AutoOpen,
    AutoClose,
    Extract,
    Database,
    Criteria,
    PrintArea,
    PrintTitles,
    Recorder,
    DataForm,
    AutoActivate,
    AutoDeactivate,
    SheetTitle,
    FilterDatabase,
}

impl BuiltInName {
    pub(super) fn from_code(code: u8) -> Option<Self> {
        Some(match code {
            0x00 => Self::ConsolidateArea,
            0x01 => Self::AutoOpen,
            0x02 => Self::AutoClose,
            0x03 => Self::Extract,
            0x04 => Self::Database,
            0x05 => Self::Criteria,
            0x06 => Self::PrintArea,
            0x07 => Self::PrintTitles,
            0x08 => Self::Recorder,
            0x09 => Self::DataForm,
            0x0A => Self::AutoActivate,
            0x0B => Self::AutoDeactivate,
            0x0C => Self::SheetTitle,
            0x0D => Self::FilterDatabase,
            _ => return None,
        })
    }

    pub(crate) fn canonical_name(self) -> &'static str {
        match self {
            Self::ConsolidateArea => "Consolidate_Area",
            Self::AutoOpen => "Auto_Open",
            Self::AutoClose => "Auto_Close",
            Self::Extract => "Extract",
            Self::Database => "Database",
            Self::Criteria => "Criteria",
            Self::PrintArea => "Print_Area",
            Self::PrintTitles => "Print_Titles",
            Self::Recorder => "Recorder",
            Self::DataForm => "Data_Form",
            Self::AutoActivate => "Auto_Activate",
            Self::AutoDeactivate => "Auto_Deactivate",
            Self::SheetTitle => "Sheet_Title",
            Self::FilterDatabase => "_FilterDatabase",
        }
    }

    pub(crate) fn code(self) -> u8 {
        match self {
            Self::ConsolidateArea => 0x00,
            Self::AutoOpen => 0x01,
            Self::AutoClose => 0x02,
            Self::Extract => 0x03,
            Self::Database => 0x04,
            Self::Criteria => 0x05,
            Self::PrintArea => 0x06,
            Self::PrintTitles => 0x07,
            Self::Recorder => 0x08,
            Self::DataForm => 0x09,
            Self::AutoActivate => 0x0a,
            Self::AutoDeactivate => 0x0b,
            Self::SheetTitle => 0x0c,
            Self::FilterDatabase => 0x0d,
        }
    }
}

/// One parsed `Lbl` slot. Macro slots remain here so `PtgName` indices do not shift.
#[derive(Debug, Clone)]
pub(crate) struct DefinedNameSlot {
    pub(super) record_index: u32,
    pub(super) name: String,
    pub(super) itab: u16,
    pub(super) hidden: bool,
    pub(super) function: bool,
    pub(super) vba_procedure: bool,
    pub(super) procedure: bool,
    pub(super) calculated_expression: bool,
    pub(super) function_group: u8,
    pub(super) published: bool,
    pub(super) workbook_parameter: bool,
    pub(super) shortcut_key: Option<u8>,
    pub(super) kind: DefinedNameKind,
    pub(super) formula_tokens: Vec<u8>,
    pub(super) formula_extra: Vec<u8>,
    pub(super) continuation_chunks: Vec<Vec<u8>>,
    pub(super) custom_menu: String,
    pub(super) description: String,
    pub(super) help_topic: String,
    pub(super) status_bar: String,
    pub(super) comment: Option<String>,
    pub(super) future_records: DefinedNameFutureRecords,
}
