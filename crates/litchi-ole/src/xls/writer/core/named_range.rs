//! Named range core types for XLS writer.
//!
//! This module defines workbook-level named ranges for the legacy XLS
//! writer. The actual BIFF8 `NAME` (Lbl) record emission is handled by
//! the `biff` module; this module only models the logical structure and
//! provides helpers to convert range references into BIFF formula bytes.

use crate::xls::XlsResult;
use crate::xls::writer::formula::{Ptg, encode_ptg_tokens, parse_cell_ref};
use crate::xls::{XlsBuiltInName, XlsDefinedNameKind, XlsError, XlsNameScope};

/// Complete inert BIFF8 `Lbl` metadata for names beyond simple ranges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsDefinedNameRecordOptions {
    pub name: String,
    pub kind: XlsDefinedNameKind,
    pub scope: XlsNameScope,
    pub hidden: bool,
    pub function: bool,
    pub vba_procedure: bool,
    pub procedure: bool,
    pub calculated_expression: bool,
    pub function_group: u8,
    pub published: bool,
    pub workbook_parameter: bool,
    pub shortcut_key: Option<u8>,
    pub formula_tokens: Vec<u8>,
    pub formula_extra: Vec<u8>,
    pub custom_menu: String,
    pub description: String,
    pub help_topic: String,
    pub status_bar: String,
    pub comment: Option<String>,
}

impl XlsDefinedNameRecordOptions {
    pub(super) fn validate(&self, sheet_count: usize) -> XlsResult<()> {
        let name_len = self.name.encode_utf16().count();
        match self.kind {
            XlsDefinedNameKind::User if !(1..=255).contains(&name_len) || self.name.contains('\0') => {
                return Err(XlsError::InvalidData("defined name must contain 1..=255 non-NUL UTF-16 units".to_string()));
            },
            XlsDefinedNameKind::BuiltIn(_) => {},
            _ => {},
        }
        if matches!(self.scope, XlsNameScope::Worksheet(index) if index >= sheet_count) {
            return Err(XlsError::InvalidData("defined name scope is outside worksheet collection".to_string()));
        }
        if (self.function || self.vba_procedure) && !self.procedure {
            return Err(XlsError::InvalidData("function/VBA name flags require procedure".to_string()));
        }
        if self.function_group > 31 {
            return Err(XlsError::InvalidData("defined name function group must be at most 31".to_string()));
        }
        if self.shortcut_key.is_some_and(|key| {
            self.function || !self.procedure || !key.is_ascii_alphabetic()
        }) {
            return Err(XlsError::InvalidData("invalid defined-name macro shortcut".to_string()));
        }
        if self.formula_tokens.len() > u16::MAX as usize
            || self.formula_tokens.len().checked_add(self.formula_extra.len()).is_none_or(|len| len > 1_048_576)
        {
            return Err(XlsError::InvalidData("defined-name formula bytes exceed resource bound".to_string()));
        }
        for value in [&self.custom_menu, &self.description, &self.help_topic, &self.status_bar] {
            if value.chars().count() > 255 || value.chars().any(|character| u32::from(character) > 0xff) {
                return Err(XlsError::InvalidData("legacy defined-name UI strings must be <=255 compressed characters".to_string()));
            }
        }
        if self.comment.as_ref().is_some_and(|comment| comment.encode_utf16().count() > 255) {
            return Err(XlsError::InvalidData("defined-name comment exceeds 255 UTF-16 units".to_string()));
        }
        Ok(())
    }

    pub(crate) fn serialized_name(&self) -> &str {
        match self.kind {
            XlsDefinedNameKind::User => &self.name,
            XlsDefinedNameKind::BuiltIn(name) => name.canonical_name(),
        }
    }

    pub(crate) fn built_in(&self) -> Option<XlsBuiltInName> {
        match self.kind { XlsDefinedNameKind::BuiltIn(name) => Some(name), _ => None }
    }
}

/// Workbook-level defined name (named range).
///
/// This mirrors the high-level structure of OOXML named ranges but is
/// tailored for BIFF8 `NAME` (Lbl) records.
#[derive(Debug, Clone)]
pub struct XlsDefinedName {
    /// Name of the defined range (e.g. "TaxRate", "SalesData").
    pub name: String,
    /// Reference text for the name.
    ///
    /// For the initial implementation this supports the following
    /// syntax forms:
    /// - Single cell: `"A1"`
    /// - Cell area: `"A1:B10"`
    ///
    /// More complex formulas are intentionally rejected so that the
    /// writer never produces syntactically invalid `rgce` payloads.
    pub reference: String,
    /// Optional user-visible comment/description.
    pub comment: Option<String>,
    /// One-based sheet index for a sheet-local name.
    ///
    /// When `None`, the name is workbook-scoped. When `Some(itab)`,
    /// the value corresponds to the `itab` field of the Lbl record
    /// and is a one-based index into the BoundSheet8 collection.
    pub local_sheet: Option<u16>,
    /// Zero-based sheet index used when encoding PtgArea3d tokens.
    ///
    /// This is the sheet whose cells the range refers to. For
    /// workbook-scoped names that still point to a single sheet
    /// (common in practice), this holds the 0-based sheet index as
    /// well.
    pub target_sheet: Option<u16>,
    /// Whether the name is hidden from the UI.
    pub hidden: bool,
    /// Whether this name represents a macro/function (not yet used).
    pub is_function: bool,
    /// Whether this name is a built-in name such as `_FilterDatabase`.
    pub is_built_in: bool,
    /// Optional built-in code for `fBuiltin` names (e.g. 13 for `_FilterDatabase`).
    pub built_in_code: Option<u8>,
}

impl XlsDefinedName {
    /// Convert this defined name's reference to a BIFF8 `rgce` payload.
    ///
    /// This currently supports only simple A1-style references as
    /// documented on [`XlsDefinedName::reference`].
    pub fn to_biff_formula(&self) -> XlsResult<Vec<u8>> {
        let trimmed = self.reference.trim();

        if let Some(colon_pos) = trimmed.find(':') {
            // Area reference like "A1:B10".
            let first_ref = trimmed[..colon_pos].trim();
            let second_ref = trimmed[colon_pos + 1..].trim();

            let start = parse_cell_ref(first_ref)?;
            let end = parse_cell_ref(second_ref)?;

            let (row_first, row_last, col_first, col_last) = match (start, end) {
                (Ptg::PtgRef(r1, c1, ..), Ptg::PtgRef(r2, c2, ..)) => {
                    let row_first = r1.min(r2);
                    let row_last = r1.max(r2);
                    let col_first = c1.min(c2);
                    let col_last = c1.max(c2);
                    (row_first, row_last, col_first, col_last)
                },
                _ => {
                    return Err(crate::xls::XlsError::InvalidData(
                        "Named range must reference cell addresses (A1-style)".to_string(),
                    ));
                },
            };

            // Prefer a 3D area reference when we know the target sheet,
            // since NameParsedFormula forbids plain PtgArea/PtgRef in
            // BIFF8. Fall back to 2D if no sheet context is available
            // (future enhancement: support multi-sheet / external refs
            // via SupBook/ExternSheet).
            if let Some(sheet_index) = self.target_sheet {
                let tokens = [Ptg::PtgArea3d(
                    sheet_index,
                    row_first,
                    row_last,
                    col_first,
                    col_last,
                )];
                Ok(encode_ptg_tokens(&tokens))
            } else {
                let tokens = [Ptg::PtgArea(row_first, row_last, col_first, col_last)];
                Ok(encode_ptg_tokens(&tokens))
            }
        } else {
            // Single-cell reference like "A1".
            let token = parse_cell_ref(trimmed)?;
            match token {
                Ptg::PtgRef(row, col, ..) => {
                    let row_first = row;
                    let row_last = row;
                    let col_first = col;
                    let col_last = col;

                    if let Some(sheet_index) = self.target_sheet {
                        let tokens = [Ptg::PtgArea3d(
                            sheet_index,
                            row_first,
                            row_last,
                            col_first,
                            col_last,
                        )];
                        Ok(encode_ptg_tokens(&tokens))
                    } else {
                        Ok(encode_ptg_tokens(&[Ptg::PtgArea(
                            row_first, row_last, col_first, col_last,
                        )]))
                    }
                },
                _ => Err(crate::xls::XlsError::InvalidData(
                    "Named range must reference a cell or cell area".to_string(),
                )),
            }
        }
    }
}
