//! Named range core types for XLS writer.
//!
//! This module defines workbook-level named ranges for the legacy XLS
//! writer. The actual BIFF8 `NAME` (Lbl) record emission is handled by
//! the `biff` module; this module only models the logical structure and
//! provides helpers to convert range references into BIFF formula bytes.

use crate::ole::xls::XlsResult;
use crate::ole::xls::writer::formula::{Ptg, encode_ptg_tokens, parse_cell_ref};

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
                    return Err(crate::ole::xls::XlsError::InvalidData(
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
                _ => Err(crate::ole::xls::XlsError::InvalidData(
                    "Named range must reference a cell or cell area".to_string(),
                )),
            }
        }
    }
}
