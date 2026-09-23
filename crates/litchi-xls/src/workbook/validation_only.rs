//! The validation-only mode of the complete XLS reader.
//!
//! Several edit owners open a complete [`Workbook`] only to validate a package
//! and to read a handful of facts from it: sheet coverage, protection, macro
//! markers, the shared-string table, and at most the few cells an edit
//! changed. [`Workbook::new`] decodes every cell of every worksheet into a
//! per-cell map with cloned strings and rendered formulas, which those owners
//! then drop unread.
//!
//! This mode runs the same parser. Every CFB, globals, worksheet and package
//! record is framed, decoded and checked by the same code, in the same order,
//! under the same limits and the same default [`OpenOptions`], so it accepts
//! exactly the packages [`Workbook::new`] accepts and refuses the others with
//! the same first error. The one difference is inside each worksheet: a
//! validated cell record is counted in a compact occupancy map, which answers
//! the duplicate-cell and `Array` ownership checks, and is decoded into a cell
//! only when its position was asked for.

use super::codec::{DecodeEveryCell, Position, ValidateCells};
use super::model::{OpenOptions, Workbook};
use crate::cell::Cell;
use crate::error::{Error, Result};
use crate::protection::{SheetProtection, WorkbookProtection};
use crate::records::SharedStringProperties;
use crate::sheet_metadata::SheetMetadata;
use litchi_cfb::OleFile;
use std::io::{Read, Seek};
use std::sync::Arc;

/// Cells a validation-only open decodes, by `BoundSheet8` tab position.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct KeptCells {
    /// Tab positions in ascending order, each with its sorted, deduplicated
    /// positions.
    tabs: Vec<(usize, Vec<Position>)>,
}

impl KeptCells {
    /// Keeps no cell: the open validates and decodes none.
    pub(crate) const fn none() -> Self {
        Self { tabs: Vec::new() }
    }

    /// Keeps the listed `(tab position, row, column)` cells.
    ///
    /// # Errors
    ///
    /// Returns an allocation error when the list cannot be retained.
    pub(crate) fn from_cells(cells: impl IntoIterator<Item = (usize, u16, u16)>) -> Result<Self> {
        let mut flat = Vec::new();
        for cell in cells {
            flat.try_reserve(1)
                .map_err(|_error| Error::Allocation("listing kept validation cells"))?;
            flat.push(cell);
        }
        flat.sort_unstable();
        flat.dedup();
        let mut tabs: Vec<(usize, Vec<Position>)> = Vec::new();
        for (tab, row, column) in flat {
            match tabs.last_mut() {
                Some((last, positions)) if *last == tab => {
                    positions
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation("listing kept validation cells"))?;
                    positions.push((row, column));
                },
                _ => {
                    let mut positions = Vec::new();
                    positions
                        .try_reserve(1)
                        .map_err(|_error| Error::Allocation("listing kept validation cells"))?;
                    positions.push((row, column));
                    tabs.try_reserve(1)
                        .map_err(|_error| Error::Allocation("listing kept validation cells"))?;
                    tabs.push((tab, positions));
                },
            }
        }
        Ok(Self { tabs })
    }

    /// The sorted positions kept on one tab.
    pub(crate) fn for_tab(&self, tab: usize) -> &[Position] {
        self.tabs
            .binary_search_by_key(&tab, |(kept_tab, _)| *kept_tab)
            .ok()
            .and_then(|index| self.tabs.get(index))
            .map_or(&[], |(_, positions)| positions.as_slice())
    }

    fn contains(&self, tab: usize, position: Position) -> bool {
        self.for_tab(tab).binary_search(&position).is_ok()
    }
}

/// Which worksheet cells one complete open decodes.
#[derive(Clone, Copy)]
pub(super) enum CellDecoding<'a> {
    /// Every cell: the public reader.
    Every,
    /// Only the kept cells; every other cell record is validated identically
    /// and counted, not decoded.
    Kept(&'a KeptCells),
}

impl<R: Read + Seek> Workbook<R> {
    /// Runs the complete validation of [`Self::new`] and keeps every fact of
    /// the workbook except the decoded cells, of which it keeps only `kept`.
    ///
    /// # Errors
    ///
    /// Returns exactly the error [`Self::new`] returns for the same reader,
    /// and accepts every package [`Self::new`] accepts. The one addition is
    /// resource exhaustion: this mode reserves its occupancy map fallibly and
    /// returns a typed allocation error where the complete reader's
    /// infallible cell map would abort the process.
    pub(crate) fn validation_only(reader: R, kept: KeptCells) -> Result<ValidationWorkbook<R>> {
        let mut workbook = Self::empty(OleFile::open(reader)?);
        workbook.xml_map = crate::xml_map::parse_stream_if_present(&mut workbook.ole_file)?;
        workbook.parse_workbook(&OpenOptions::default(), CellDecoding::Kept(&kept))?;
        Ok(ValidationWorkbook { workbook, kept })
    }

    /// Parses one worksheet with the store `cells` selects for its tab.
    pub(super) fn parse_worksheet_decoding(
        &self,
        workbook_data: &[u8],
        tab: usize,
        bound_sheet: &crate::records::BoundSheetRecord,
        encoding: &crate::records::Encoding,
        compatibility_profile: crate::CompatibilityProfile,
        cells: CellDecoding<'_>,
    ) -> Result<crate::worksheet::Worksheet> {
        match cells {
            CellDecoding::Every => self.parse_worksheet_from_position(
                workbook_data,
                bound_sheet,
                encoding,
                compatibility_profile,
                &mut DecodeEveryCell,
            ),
            CellDecoding::Kept(kept) => self.parse_worksheet_from_position(
                workbook_data,
                bound_sheet,
                encoding,
                compatibility_profile,
                &mut ValidateCells::new(kept.for_tab(tab)),
            ),
        }
    }
}

/// A complete, independent validation open of an XLS package.
///
/// It holds everything [`Workbook::new`] would hold except the decoded cells:
/// the same sheet directory, protection, macro markers, shared-string table,
/// worksheet collectors and package facts, reached through accessors of the
/// same names. It is deliberately not a [`Workbook`]: its worksheets carry
/// only the cells it was asked to keep, so the one cell accessor it offers
/// refuses any other position rather than answering it as empty.
#[derive(Debug)]
pub(crate) struct ValidationWorkbook<R: Read + Seek> {
    workbook: Workbook<R>,
    kept: KeptCells,
}

impl<R: Read + Seek> ValidationWorkbook<R> {
    /// All workbook sheet directory entries in tab order.
    pub(crate) fn sheets(&self) -> &[SheetMetadata] {
        self.workbook.sheets()
    }

    /// Sheet directory entry at a workbook tab index.
    pub(crate) fn sheet(&self, index: usize) -> Option<&SheetMetadata> {
        self.workbook.sheet(index)
    }

    /// Workbook-level protection records.
    pub(crate) fn protection(&self) -> &WorkbookProtection {
        self.workbook.protection()
    }

    /// Protection records of the parsed worksheet at `worksheet_index`.
    ///
    /// # Errors
    ///
    /// Returns the [`Workbook::xls_worksheet`] error for an index that names
    /// no parsed worksheet.
    pub(crate) fn worksheet_protection(&self, worksheet_index: usize) -> Result<&SheetProtection> {
        Ok(self.workbook.xls_worksheet(worksheet_index)?.protection())
    }

    /// Comments of the parsed worksheet at `worksheet_index`.
    ///
    /// # Errors
    ///
    /// Returns the [`Workbook::xls_worksheet`] error for an index that names
    /// no parsed worksheet.
    pub(crate) fn worksheet_comments(
        &self,
        worksheet_index: usize,
    ) -> Result<&[crate::comments::Comment]> {
        Ok(self.workbook.xls_worksheet(worksheet_index)?.comments())
    }

    /// Inert VBA project metadata, as [`Workbook::vba_metadata`] reports it.
    pub(crate) fn vba_metadata(&self) -> crate::vba::VbaMetadata {
        self.workbook.vba_metadata()
    }

    /// The `_VBA_PROJECT_CUR` storage, as [`Workbook::vba_project_storage`]
    /// discovers it.
    pub(crate) fn vba_project_storage(&self) -> Option<crate::VbaProjectStorage> {
        self.workbook.vba_project_storage()
    }

    /// The shared-string table exactly as this open built it.
    pub(crate) fn shared_strings_shared(&self) -> Arc<Vec<String>> {
        self.workbook.shared_strings_shared()
    }

    /// The rich-text and phonetic property table exactly as this open built it.
    pub(crate) fn shared_string_properties_shared(
        &self,
    ) -> Option<Arc<Vec<Option<Box<SharedStringProperties>>>>> {
        self.workbook.shared_string_properties_shared()
    }

    /// Rich-text and phonetic properties for a shared-string index.
    pub(crate) fn shared_string_properties(&self, index: u32) -> Option<&SharedStringProperties> {
        self.workbook.shared_string_properties(index)
    }

    /// The cell the complete reader keeps at a kept position of the parsed
    /// worksheet at `worksheet_index`, or `None` when that reader keeps none
    /// there.
    ///
    /// # Errors
    ///
    /// Returns the [`Workbook::xls_worksheet`] error for an index that names
    /// no parsed worksheet, and an unsafe-edit refusal for a position this
    /// open was not asked to keep, which it cannot answer.
    pub(crate) fn kept_cell(
        &self,
        worksheet_index: usize,
        row: u16,
        column: u16,
    ) -> Result<Option<&Cell>> {
        let worksheet = self.workbook.xls_worksheet(worksheet_index)?;
        let kept = self
            .workbook
            .sheets()
            .iter()
            .position(|sheet| sheet.parsed_worksheet_index() == Some(worksheet_index))
            .is_some_and(|tab| self.kept.contains(tab, (row, column)));
        if !kept {
            return Err(Error::UnsafeEdit(
                "the validation-only XLS open was not asked to keep the requested cell".into(),
            ));
        }
        Ok(worksheet.get_cell(u32::from(row), u32::from(column)))
    }

    /// The complete reader's workbook, for differential tests that compare
    /// every non-cell fact of the two modes.
    #[cfg(test)]
    pub(crate) fn as_workbook_for_tests(&self) -> &Workbook<R> {
        &self.workbook
    }
}
