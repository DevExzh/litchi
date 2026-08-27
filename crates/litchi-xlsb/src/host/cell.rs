//! Cell representation for XLSB files

use crate::package::error::{Error, Result};
use crate::package::formula::{Compiler, Context, Group, GroupKind, ParsedFormula, Parser};
use crate::package::records::CellRecord;
use crate::package::shared_strings::SharedString;
use litchi_core::sheet::{Cell as SheetCell, CellValue};
use std::sync::Arc;

/// Why a formula cell has an opaque formula rather than a resolved expression.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FormulaOpacityReason {
    /// The formula uses a feature that this XLSB implementation cannot render.
    Unsupported,
    /// The formula depends on workbook metadata that is not available.
    Unresolved,
    /// The formula bytes were authored or received without validation.
    Unvalidated,
}

/// Resolution state for a formula cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FormulaResolutionStatus {
    /// A usable formula text representation is present; this does not prove
    /// independent syntax validation, evaluation, or cached-value freshness.
    Resolved,
    /// The formula is preserved, but its expression is intentionally opaque.
    Opaque(FormulaOpacityReason),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellHeader {
    pub col: u32,
    pub style_id: u32,
    pub show_phonetic: bool,
}

/// XLSB cell implementation
///
/// Dynamically sized values precede compact coordinate and flag metadata.
#[derive(Debug, Clone)]
pub struct Cell {
    /// Cell value (largest field, aligned first)
    value: CellValue,
    /// Row index (0-based)
    row: u32,
    /// Column index (0-based)
    col: u32,
    /// Zero-based index into the workbook's cell XF table.
    style_id: u32,
    /// Whether phonetic information should be shown for this cell.
    show_phonetic: bool,
    /// Resolution state for a formula cell, or `None` for an ordinary cell.
    formula_resolution_status: Option<FormulaResolutionStatus>,
    /// Source-stored cached formula result, when one was present.
    cached_value: Option<Box<CellValue>>,
    /// Original binary formula, retained even if a token is not understood.
    formula: Option<ParsedFormula>,
    /// Exact formula bytes received from a legacy `CellRecord` without parsing.
    raw_formula: Option<Vec<u8>>,
    /// Inline rich-text value from a `BrtCellRString` record.
    rich_string: Option<SharedString>,
    /// Shared array/shared-formula definition.
    formula_group: Option<Arc<Group>>,
    /// Formula calculation flags from `GrbitFmla`.
    formula_flags: u16,
}

impl Cell {
    fn authored_formula_metadata(
        value: &CellValue,
    ) -> (Option<FormulaResolutionStatus>, Option<Box<CellValue>>) {
        match value {
            CellValue::Formula {
                formula,
                cached_value,
                ..
            } => {
                let status = if formula.is_empty() {
                    FormulaResolutionStatus::Opaque(FormulaOpacityReason::Unvalidated)
                } else {
                    FormulaResolutionStatus::Resolved
                };
                (Some(status), cached_value.clone())
            },
            CellValue::Empty
            | CellValue::Bool(_)
            | CellValue::Int(_)
            | CellValue::Float(_)
            | CellValue::String(_)
            | CellValue::DateTime(_)
            | CellValue::Error(_) => (None, None),
        }
    }

    /// Create a new XLSB cell
    pub fn new(row: u32, col: u32, value: CellValue) -> Self {
        let (formula_resolution_status, cached_value) = Self::authored_formula_metadata(&value);
        Self {
            row,
            col,
            style_id: 0,
            show_phonetic: false,
            value,
            formula_resolution_status,
            cached_value,
            formula: None,
            raw_formula: None,
            rich_string: None,
            formula_group: None,
            formula_flags: 0,
        }
    }

    /// Create a new XLSB cell from an authored formula value.
    ///
    /// A [`CellValue::Formula`] with non-empty formula text is marked as
    /// [`FormulaResolutionStatus::Resolved`]. A scalar value, or a formula
    /// value with no text, is retained as opaque and marked
    /// [`FormulaOpacityReason::Unvalidated`].
    pub fn new_formula(row: u32, col: u32, value: CellValue) -> Self {
        let (formula_resolution_status, cached_value) = match &value {
            CellValue::Formula { .. } => Self::authored_formula_metadata(&value),
            CellValue::Empty
            | CellValue::Bool(_)
            | CellValue::Int(_)
            | CellValue::Float(_)
            | CellValue::String(_)
            | CellValue::DateTime(_)
            | CellValue::Error(_) => (
                Some(FormulaResolutionStatus::Opaque(
                    FormulaOpacityReason::Unvalidated,
                )),
                Some(Box::new(value.clone())),
            ),
        };
        Self {
            row,
            col,
            style_id: 0,
            show_phonetic: false,
            value,
            formula_resolution_status,
            cached_value,
            formula: None,
            raw_formula: None,
            rich_string: None,
            formula_group: None,
            formula_flags: 0,
        }
    }

    /// Create a formula cell while preserving the exact XLSB token and
    /// ancillary streams.
    pub(crate) fn new_formula_binary(
        row: u32,
        header: CellHeader,
        cached_value: CellValue,
        formula: ParsedFormula,
        formula_flags: u16,
        formula_context: &Context,
    ) -> Result<Self> {
        let resolution = Parser::with_extra(&formula.rgce, &formula.rgcb)
            .parse()
            .map_err(Error::from)
            .and_then(|tokens| {
                Compiler::try_tokens_to_string_with_resolution(&tokens, formula_context)
                    .map_err(Error::from)
            });
        let (value, formula_resolution_status) = match resolution {
            Ok(formula_text) => (
                CellValue::Formula {
                    formula: formula_text,
                    cached_value: Some(Box::new(cached_value.clone())),
                    is_array: false,
                    array_range: None,
                },
                FormulaResolutionStatus::Resolved,
            ),
            Err(Error::UnsupportedFeature(_)) => (
                cached_value.clone(),
                FormulaResolutionStatus::Opaque(FormulaOpacityReason::Unsupported),
            ),
            Err(Error::UnresolvedDependency(_)) => (
                cached_value.clone(),
                FormulaResolutionStatus::Opaque(FormulaOpacityReason::Unresolved),
            ),
            Err(error) => return Err(error),
        };

        Ok(Self {
            value,
            row,
            col: header.col,
            style_id: header.style_id,
            show_phonetic: header.show_phonetic,
            formula_resolution_status: Some(formula_resolution_status),
            cached_value: Some(Box::new(cached_value)),
            formula: Some(formula),
            raw_formula: None,
            rich_string: None,
            formula_group: None,
            formula_flags,
        })
    }

    pub(crate) fn new_grouped_formula(
        row: u32,
        header: CellHeader,
        cached_value: CellValue,
        placeholder: ParsedFormula,
        formula_flags: u16,
        group: Arc<Group>,
        formula_context: &Context,
    ) -> Result<Self> {
        let resolution = match group.kind {
            GroupKind::Array => {
                Parser::with_extra(&group.formula.rgce, &group.formula.rgcb).parse()
            },
            GroupKind::Shared => Parser::with_base_cell_and_extra(
                &group.formula.rgce,
                &group.formula.rgcb,
                row,
                header.col,
            )
            .parse(),
        }
        .map_err(Error::from)
        .and_then(|tokens| {
            Compiler::try_tokens_to_string_with_resolution(&tokens, formula_context)
                .map_err(Error::from)
        });
        let is_array = group.kind == GroupKind::Array;
        let (value, formula_resolution_status) = match resolution {
            Ok(formula_text) => (
                CellValue::Formula {
                    formula: formula_text,
                    cached_value: Some(Box::new(cached_value.clone())),
                    is_array,
                    array_range: is_array.then(|| group.range.to_a1()),
                },
                FormulaResolutionStatus::Resolved,
            ),
            Err(Error::UnsupportedFeature(_)) => (
                cached_value.clone(),
                FormulaResolutionStatus::Opaque(FormulaOpacityReason::Unsupported),
            ),
            Err(Error::UnresolvedDependency(_)) => (
                cached_value.clone(),
                FormulaResolutionStatus::Opaque(FormulaOpacityReason::Unresolved),
            ),
            Err(error) => return Err(error),
        };
        Ok(Self {
            value,
            row,
            col: header.col,
            style_id: header.style_id,
            show_phonetic: header.show_phonetic,
            formula_resolution_status: Some(formula_resolution_status),
            cached_value: Some(Box::new(cached_value)),
            formula: Some(placeholder),
            raw_formula: None,
            rich_string: None,
            formula_group: Some(group),
            formula_flags,
        })
    }

    /// Raw XLSB formula RPN token stream (`rgce`).
    pub fn formula_bytes(&self) -> Option<&[u8]> {
        self.formula.as_ref().map(|formula| formula.rgce.as_slice())
    }

    /// Raw XLSB ancillary formula stream (`rgcb`).
    pub fn formula_extra_bytes(&self) -> Option<&[u8]> {
        self.formula.as_ref().map(|formula| formula.rgcb.as_slice())
    }

    /// Exact formula bytes received from a legacy cell record without parsing.
    pub fn raw_formula_bytes(&self) -> Option<&[u8]> {
        self.raw_formula.as_deref()
    }

    /// Formula resolution state, or `None` for a non-formula cell.
    pub fn formula_resolution_status(&self) -> Option<FormulaResolutionStatus> {
        self.formula_resolution_status
    }

    /// Source-stored cached result for a formula, never an evaluated result.
    pub fn cached_value(&self) -> Option<&CellValue> {
        self.cached_value.as_deref()
    }

    /// Actual formula definition for an array/shared formula cell.
    pub fn formula_definition_bytes(&self) -> Option<&[u8]> {
        self.formula_group
            .as_ref()
            .map(|group| group.formula.rgce.as_slice())
    }

    /// Ancillary token data for an array/shared formula definition.
    pub fn formula_definition_extra_bytes(&self) -> Option<&[u8]> {
        self.formula_group
            .as_ref()
            .map(|group| group.formula.rgcb.as_slice())
    }

    /// Whether the grouped definition requests unconditional recalculation.
    pub fn formula_group_always_calculates(&self) -> Option<bool> {
        self.formula_group
            .as_ref()
            .map(|group| group.always_calculate)
    }

    /// Whether this cell belongs to an XLSB shared-formula group.
    pub fn is_shared_formula(&self) -> bool {
        self.formula_group
            .as_ref()
            .is_some_and(|group| group.kind == GroupKind::Shared)
    }

    /// Array/shared formula range in A1 notation.
    pub fn formula_range(&self) -> Option<String> {
        self.formula_group.as_ref().map(|group| group.range.to_a1())
    }

    /// Formula recalculation flags from `GrbitFmla`.
    pub fn formula_flags(&self) -> Option<u16> {
        self.is_formula().then_some(self.formula_flags)
    }

    /// Zero-based index into
    /// [`StylesTable::cell_xfs`](crate::package::styles_table::StylesTable::cell_xfs).
    pub fn style_id(&self) -> u32 {
        self.style_id
    }

    /// Whether the XLSB cell header requests display of phonetic information.
    pub fn show_phonetic(&self) -> bool {
        self.show_phonetic
    }

    /// Complete inline rich-text value, when this cell came from `BrtCellRString`.
    pub fn rich_string(&self) -> Option<&SharedString> {
        self.rich_string.as_ref()
    }

    pub(crate) fn new_styled(row: u32, header: CellHeader, value: CellValue) -> Self {
        Self {
            row,
            col: header.col,
            style_id: header.style_id,
            show_phonetic: header.show_phonetic,
            value,
            formula_resolution_status: None,
            cached_value: None,
            formula: None,
            raw_formula: None,
            rich_string: None,
            formula_group: None,
            formula_flags: 0,
        }
    }

    pub(crate) fn new_rich_string(row: u32, header: CellHeader, rich_string: SharedString) -> Self {
        Self {
            value: CellValue::String(rich_string.text.clone()),
            row,
            col: header.col,
            style_id: header.style_id,
            show_phonetic: header.show_phonetic,
            formula_resolution_status: None,
            cached_value: None,
            formula: None,
            raw_formula: None,
            rich_string: Some(rich_string),
            formula_group: None,
            formula_flags: 0,
        }
    }

    /// Create cell from XLSB record
    pub fn from_record(record: &CellRecord, shared_strings: Option<&Vec<String>>) -> Option<Self> {
        let (value, formula_resolution_status, raw_formula, cached_value) = match &record.value {
            crate::package::records::CellValue::Blank => (CellValue::Empty, None, None, None),
            crate::package::records::CellValue::Bool(b) => (CellValue::Bool(*b), None, None, None),
            crate::package::records::CellValue::Error(e) => {
                // Convert error code to Excel error string
                let error_str = match e {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0F => "#VALUE!",
                    0x17 => "#REF!",
                    0x1D => "#NAME?",
                    0x24 => "#NUM!",
                    0x2A => "#N/A",
                    0x2B => "#GETTING_DATA",
                    _ => "#ERR!",
                };
                (CellValue::Error(error_str.to_string()), None, None, None)
            },
            crate::package::records::CellValue::Real(f) => (CellValue::Float(*f), None, None, None),
            crate::package::records::CellValue::String(s) => {
                (CellValue::String(s.clone()), None, None, None)
            },
            crate::package::records::CellValue::Isst(idx) => {
                let val = if let Some(sst) = shared_strings {
                    if let Some(s) = sst.get(*idx as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::Error("Invalid SST index".to_string())
                    }
                } else {
                    CellValue::Error("SST not available".to_string())
                };
                (val, None, None, None)
            },
            crate::package::records::CellValue::Formula { value, formula } => {
                // Extract the cached value from the formula
                // Formula parsing from bytes is complex and formula bytes are in a binary RPN format
                // For now, we use the cached value which is sufficient for most use cases
                let cached = Self::extract_formula_value(value, shared_strings);
                (
                    cached.clone(),
                    Some(FormulaResolutionStatus::Opaque(
                        FormulaOpacityReason::Unvalidated,
                    )),
                    formula.clone(),
                    Some(Box::new(cached)),
                )
            },
        };

        Some(Cell {
            row: record.row,
            col: u32::from(record.col),
            style_id: 0,
            show_phonetic: false,
            value,
            formula_resolution_status,
            cached_value,
            formula: None,
            raw_formula,
            rich_string: None,
            formula_group: None,
            formula_flags: 0,
        })
    }

    /// Extract value from formula cached value
    fn extract_formula_value(
        formula_value: &crate::package::records::CellValue,
        shared_strings: Option<&Vec<String>>,
    ) -> CellValue {
        match formula_value {
            crate::package::records::CellValue::Blank => CellValue::Empty,
            crate::package::records::CellValue::Bool(b) => CellValue::Bool(*b),
            crate::package::records::CellValue::Error(e) => {
                let error_str = match e {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0F => "#VALUE!",
                    0x17 => "#REF!",
                    0x1D => "#NAME?",
                    0x24 => "#NUM!",
                    0x2A => "#N/A",
                    0x2B => "#GETTING_DATA",
                    _ => "#ERR!",
                };
                CellValue::Error(error_str.to_string())
            },
            crate::package::records::CellValue::Real(f) => CellValue::Float(*f),
            crate::package::records::CellValue::String(s) => CellValue::String(s.clone()),
            crate::package::records::CellValue::Isst(idx) => {
                if let Some(sst) = shared_strings {
                    if let Some(s) = sst.get(*idx as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::Error("Invalid SST index".to_string())
                    }
                } else {
                    CellValue::Error("SST not available".to_string())
                }
            },
            crate::package::records::CellValue::Formula { value, formula: _ } => {
                // Recursive formula values (shouldn't happen, but handle it)
                Self::extract_formula_value(value, shared_strings)
            },
        }
    }
}

impl SheetCell for Cell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.col
    }

    fn coordinate(&self) -> String {
        crate::package::utils::cell_reference(self.row, self.col)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        self.formula_resolution_status.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::{Cell, CellHeader, FormulaOpacityReason, FormulaResolutionStatus};
    use crate::package::error::Error;
    use crate::package::formula::{
        Context, Group, GroupKind, ParsedFormula, Range, ptg_types, text,
    };
    use crate::package::records::{CellRecord, CellValue as RecordCellValue};
    use litchi_core::sheet::{Cell as SheetCell, CellValue};
    use std::sync::Arc;

    fn header(col: u32) -> CellHeader {
        CellHeader {
            col,
            style_id: 3,
            show_phonetic: true,
        }
    }

    #[test]
    fn ordinary_cells_have_no_formula_resolution_status() {
        let cell = Cell::new(2, 4, CellValue::String("text".to_string()));

        assert_eq!(cell.formula_resolution_status(), None);
        assert_eq!(cell.cached_value(), None);
        assert!(!cell.is_formula());
    }

    #[test]
    fn authored_formula_status_distinguishes_text_empty_and_scalar_values() {
        let cached = CellValue::Int(3);
        let authored = Cell::new_formula(
            0,
            0,
            CellValue::Formula {
                formula: "1+2".to_string(),
                cached_value: Some(Box::new(cached.clone())),
                is_array: false,
                array_range: None,
            },
        );
        assert_eq!(
            authored.formula_resolution_status(),
            Some(FormulaResolutionStatus::Resolved)
        );
        assert_eq!(authored.cached_value(), Some(&cached));

        let empty = Cell::new_formula(
            0,
            1,
            CellValue::Formula {
                formula: String::new(),
                cached_value: None,
                is_array: false,
                array_range: None,
            },
        );
        assert_eq!(
            empty.formula_resolution_status(),
            Some(FormulaResolutionStatus::Opaque(
                FormulaOpacityReason::Unvalidated
            ))
        );

        let scalar = Cell::new_formula(0, 2, CellValue::String("source".to_string()));
        assert_eq!(
            scalar.formula_resolution_status(),
            Some(FormulaResolutionStatus::Opaque(
                FormulaOpacityReason::Unvalidated
            ))
        );
        assert_eq!(
            scalar.cached_value(),
            Some(&CellValue::String("source".to_string()))
        );
    }

    #[test]
    fn binary_arithmetic_resolves_and_keeps_the_source_cache() {
        let formula = text::Compiler::compile("1+2").unwrap();
        let cached = CellValue::Int(99);
        let cell = Cell::new_formula_binary(
            4,
            header(5),
            cached.clone(),
            formula,
            0x0002,
            &Context::default(),
        )
        .unwrap();

        assert_eq!(
            cell.formula_resolution_status(),
            Some(FormulaResolutionStatus::Resolved)
        );
        assert_eq!(cell.cached_value(), Some(&cached));
        assert_eq!(cell.formula_flags(), Some(0x0002));
        match cell.value() {
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                array_range,
            } => {
                assert_eq!(formula, "(1+2)");
                assert_eq!(cached_value.as_deref(), Some(&cached));
                assert!(!is_array);
                assert_eq!(*array_range, None);
            },
            value @ (CellValue::Empty
            | CellValue::Bool(_)
            | CellValue::Int(_)
            | CellValue::Float(_)
            | CellValue::String(_)
            | CellValue::DateTime(_)
            | CellValue::Error(_)) => panic!("expected resolved formula, got {value:?}"),
        }
    }

    #[test]
    fn binary_missing_name_is_opaque_but_preserves_cached_value() {
        let formula = ParsedFormula {
            rgce: vec![ptg_types::PTG_NAME, 2, 0, 0, 0],
            rgcb: Vec::new(),
        };
        let cached = CellValue::String("cached".to_string());
        let rgce = formula.rgce.clone();
        let cell = Cell::new_formula_binary(
            0,
            header(0),
            cached.clone(),
            formula,
            0,
            &Context::default(),
        )
        .unwrap();

        assert_eq!(
            cell.formula_resolution_status(),
            Some(FormulaResolutionStatus::Opaque(
                FormulaOpacityReason::Unresolved
            ))
        );
        assert_eq!(cell.value(), &cached);
        assert_eq!(cell.cached_value(), Some(&cached));
        assert_eq!(cell.formula_bytes(), Some(rgce.as_slice()));
    }

    #[test]
    fn known_unsupported_formula_source_is_opaque() {
        let formula = ParsedFormula {
            rgce: vec![ptg_types::PTG_FUNC, 0xFF, 0xFF],
            rgcb: Vec::new(),
        };
        let cached = CellValue::Int(17);
        let cell = Cell::new_formula_binary(
            0,
            header(0),
            cached.clone(),
            formula,
            0,
            &Context::default(),
        )
        .unwrap();

        assert_eq!(
            cell.formula_resolution_status(),
            Some(FormulaResolutionStatus::Opaque(
                FormulaOpacityReason::Unsupported
            ))
        );
        assert_eq!(cell.value(), &cached);
        assert_eq!(cell.cached_value(), Some(&cached));
    }

    #[test]
    fn malformed_known_formula_token_remains_a_typed_error() {
        let result = Cell::new_formula_binary(
            0,
            header(0),
            CellValue::Int(1),
            ParsedFormula {
                rgce: vec![ptg_types::PTG_NAME, 0, 0, 0, 0],
                rgcb: Vec::new(),
            },
            0,
            &Context::default(),
        );

        assert!(matches!(result, Err(Error::InvalidFormula(_))));
    }

    #[test]
    fn binary_cell_preserves_exact_parsed_rgce_and_rgcb() {
        let formula = text::Compiler::compile("{1,2}").unwrap();
        assert!(!formula.rgcb.is_empty());
        let rgce = formula.rgce.clone();
        let rgcb = formula.rgcb.clone();
        let cell = Cell::new_formula_binary(
            0,
            header(0),
            CellValue::String("source cache".to_string()),
            formula,
            0,
            &Context::default(),
        )
        .unwrap();

        assert_eq!(cell.formula_bytes(), Some(rgce.as_slice()));
        assert_eq!(cell.formula_extra_bytes(), Some(rgcb.as_slice()));
    }

    #[test]
    fn from_record_keeps_raw_formula_and_marks_unvalidated_cache() {
        let raw = vec![0xCA, 0xFE, 0x01, 0x02];
        let record = CellRecord {
            row: 5,
            col: 7,
            value: RecordCellValue::Formula {
                value: Box::new(RecordCellValue::String("cached".to_string())),
                formula: Some(raw.clone()),
            },
        };
        let cell = Cell::from_record(&record, None).unwrap();
        let cached = CellValue::String("cached".to_string());

        assert_eq!(cell.raw_formula_bytes(), Some(raw.as_slice()));
        assert_eq!(cell.value(), &cached);
        assert_eq!(cell.cached_value(), Some(&cached));
        assert_eq!(
            cell.formula_resolution_status(),
            Some(FormulaResolutionStatus::Opaque(
                FormulaOpacityReason::Unvalidated
            ))
        );
    }

    fn assert_group_streams(cell: &Cell, placeholder: &ParsedFormula, group: &Group) {
        assert_eq!(cell.formula_bytes(), Some(placeholder.rgce.as_slice()));
        assert_eq!(
            cell.formula_extra_bytes(),
            Some(placeholder.rgcb.as_slice())
        );
        assert_eq!(
            cell.formula_definition_bytes(),
            Some(group.formula.rgce.as_slice())
        );
        assert_eq!(
            cell.formula_definition_extra_bytes(),
            Some(group.formula.rgcb.as_slice())
        );
        assert_eq!(cell.formula_range(), Some(group.range.to_a1()));
    }

    #[test]
    fn grouped_cells_preserve_placeholders_definitions_and_ranges_for_both_states() {
        let context = Context::default();
        let resolved_group = Group {
            kind: GroupKind::Array,
            range: Range::new(1, 2, 3, 4).unwrap(),
            formula: text::Compiler::compile("1+1").unwrap(),
            always_calculate: true,
        };
        let resolved_placeholder = ParsedFormula::exp(1, 3).unwrap();
        let resolved_cached = CellValue::Int(2);
        let resolved = Cell::new_grouped_formula(
            1,
            header(3),
            resolved_cached.clone(),
            resolved_placeholder.clone(),
            0x0002,
            Arc::new(resolved_group.clone()),
            &context,
        )
        .unwrap();
        assert_eq!(
            resolved.formula_resolution_status(),
            Some(FormulaResolutionStatus::Resolved)
        );
        assert_eq!(resolved.cached_value(), Some(&resolved_cached));
        assert!(!resolved.is_shared_formula());
        assert_eq!(resolved.formula_group_always_calculates(), Some(true));
        assert_group_streams(&resolved, &resolved_placeholder, &resolved_group);
        match resolved.value() {
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                array_range,
            } => {
                assert_eq!(formula, "(1+1)");
                assert_eq!(cached_value.as_deref(), Some(&resolved_cached));
                assert!(*is_array);
                assert_eq!(array_range.as_deref(), Some("D2:E3"));
            },
            value @ (CellValue::Empty
            | CellValue::Bool(_)
            | CellValue::Int(_)
            | CellValue::Float(_)
            | CellValue::String(_)
            | CellValue::DateTime(_)
            | CellValue::Error(_)) => {
                panic!("expected resolved grouped formula, got {value:?}")
            },
        }

        let opaque_group = Group {
            kind: GroupKind::Shared,
            range: Range::new(4, 5, 0, 0).unwrap(),
            formula: ParsedFormula {
                rgce: vec![ptg_types::PTG_NAME, 2, 0, 0, 0],
                rgcb: Vec::new(),
            },
            always_calculate: false,
        };
        let opaque_placeholder = ParsedFormula::exp(4, 0).unwrap();
        let opaque_cached = CellValue::String("group cache".to_string());
        let opaque = Cell::new_grouped_formula(
            4,
            header(0),
            opaque_cached.clone(),
            opaque_placeholder.clone(),
            0,
            Arc::new(opaque_group.clone()),
            &context,
        )
        .unwrap();
        assert_eq!(
            opaque.formula_resolution_status(),
            Some(FormulaResolutionStatus::Opaque(
                FormulaOpacityReason::Unresolved
            ))
        );
        assert_eq!(opaque.value(), &opaque_cached);
        assert_eq!(opaque.cached_value(), Some(&opaque_cached));
        assert!(opaque.is_shared_formula());
        assert_eq!(opaque.formula_group_always_calculates(), Some(false));
        assert_group_streams(&opaque, &opaque_placeholder, &opaque_group);
    }
}
