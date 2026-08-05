//! Typed workbook state and semantic XLS accessors.
//!
//! This module owns the in-memory workbook model exposed by the crate. BIFF
//! record decoding lives in the sibling codec module, while CFB stream and
//! transaction operations live in the sibling package module.

use crate::cell::Cell;
use crate::defined_names::{BuiltInName, DefinedName, DefinedNameKind, NameScope};
use crate::error::{Error, Result};
use crate::formula::FormulaContext;
use crate::leniency::{Leniency, ToleranceReport};
use crate::number_format::{DateSystem, ExtendedFormat, Formatting, NumberFormat};
use crate::pivot_table;
use crate::protection;
use crate::records::{BiffVersion, SharedStringProperties};
use crate::sheet_metadata::SheetMetadata;
use crate::worksheet::Worksheet;
use litchi_cfb::OleFile;
use litchi_core::sheet::{
    Result as SheetResult, Worksheet as SheetTrait, WorksheetIterator as WorksheetIteratorTrait,
};
use std::io::{Read, Seek};
use std::sync::Arc;

/// XLS workbook implementation
#[derive(Debug)]
pub struct Workbook<R: Read + Seek> {
    pub(super) ole_file: OleFile<R>,
    pub(super) worksheets: Vec<Worksheet>,
    pub(super) worksheet_names: Vec<String>,
    pub(super) sheets: Vec<SheetMetadata>,
    /// Shared string table (Arc for zero-copy sharing across worksheets)
    pub(super) shared_strings: Option<Arc<Vec<String>>>,
    /// Sparse rich-text and phonetic properties parallel to `shared_strings`.
    pub(super) shared_string_properties: Option<Arc<Vec<Option<Box<SharedStringProperties>>>>>,
    pub(super) shared_string_reference_count: u32,
    pub(super) palette: crate::palette::Palette,
    pub(super) fonts: Vec<crate::font::Font>,
    pub(super) biff_version: BiffVersion,
    pub(super) is_1904_date_system: bool,
    pub(super) formula_context: FormulaContext,
    pub(super) defined_names: Vec<DefinedName>,
    pub(super) defined_name_records: Vec<DefinedName>,
    pub(super) formatting: Arc<Formatting>,
    pub(super) protection: protection::WorkbookProtection,
    pub(super) calculation: crate::calculation::WorkbookCalculation,
    pub(super) vba_metadata: crate::vba::VbaMetadata,
    pub(super) environment: crate::environment::WorkbookEnvironment,
    pub(super) book_ext: Option<crate::book_ext::BookExt>,
    pub(super) style_extensions: Vec<crate::style_ext::StyleExt>,
    pub(super) theme: Option<crate::theme::Theme>,
    pub(super) write_access: Result<Option<crate::access::WriteAccess>>,
    pub(super) table_styles: Option<crate::table_styles::TableStyles>,
    pub(super) shared_string_index: Result<Option<crate::shared_string_index::SharedStringIndex>>,
    pub(super) workbook_view: crate::workbook_view::WorkbookView,
    /// Workbook-wide custom views (`UserBView` records), in record order.
    pub(super) custom_views: Vec<crate::custom_view::WorkbookCustomView>,
    /// Real-time data (RTD) topics (`RealTimeData` records), in record order.
    pub(super) real_time_data: Vec<crate::real_time_data::RealTimeData>,
    /// MDX (OLAP cube) metadata from the workbook globals `METADATA` production.
    pub(super) mdx_metadata: crate::mdx_metadata::MdxMetadata,
    /// Published Web pages (`WebPub` records), in record order.
    pub(super) web_publications: Vec<crate::web_pub::WebPub>,
    pub(super) function_groups: Option<crate::function_group::FunctionGroups>,
    pub(super) external_links: crate::external_link::Links,
    pub(super) pivot_caches: Vec<crate::PivotCache>,
    /// SXStreamID values in global PivotCache ordinal order.
    pub(super) pivot_cache_stream_ids: Vec<u16>,
    /// Formatting defects repaired while opening; always empty in strict mode.
    pub(super) tolerance: ToleranceReport,
}

/// Options for opening a legacy XLS workbook.
#[derive(Debug, Clone, Copy, Default)]
pub struct OpenOptions<'a> {
    /// Password used for BIFF8 password-to-open encryption.
    pub password: Option<&'a str>,
    /// How non-structural formatting defects are treated.
    ///
    /// Defaults to [`Leniency::Strict`], which rejects any deviation from
    /// MS-XLS. Set [`Leniency::TolerateFormattingDefects`] to open the
    /// widespread real-world workbooks whose cosmetic formatting metadata is
    /// self-contradictory; everything repaired is then enumerable through
    /// [`Workbook::tolerance_report`]. Structural defects — record framing,
    /// stream grammar, and encryption — remain hard errors either way.
    pub leniency: Leniency,
}

impl<R: Read + Seek> Workbook<R> {
    /// Formatting defects repaired while opening this workbook.
    ///
    /// Always clean under [`Leniency::Strict`], because a strict open either
    /// rejects the defect or never encounters one. Under
    /// [`Leniency::TolerateFormattingDefects`] every repair the reader made
    /// is enumerable here; see [`crate::FormattingDefect`] for the
    /// closed set of defects that can appear and the substitute value each one
    /// produced.
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.tolerance
    }

    /// Optional `ExtSST` shared-string lookup index.
    ///
    /// A malformed optional index is reported here without preventing workbook content parsing.
    pub fn shared_string_index(
        &self,
    ) -> std::result::Result<Option<&crate::shared_string_index::SharedStringIndex>, &Error> {
        match &self.shared_string_index {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(error),
        }
    }

    /// Workbook window state and stable sheet identifiers.
    pub fn workbook_view(&self) -> &crate::workbook_view::WorkbookView {
        &self.workbook_view
    }

    /// Workbook-wide custom views (`UserBView` records), in record order.
    ///
    /// The views are inert: applying one is a UI operation this reader never
    /// performs.
    pub fn custom_views(&self) -> &[crate::custom_view::WorkbookCustomView] {
        &self.custom_views
    }

    /// Real-time data (RTD) topics declared in the workbook globals, in
    /// record order.
    ///
    /// The topics are inert: this reader never locates, launches, or queries
    /// an RTD server; each entry only reports the topic, the last cached
    /// value, and the subscribed cells.
    pub fn real_time_data(&self) -> &[crate::real_time_data::RealTimeData] {
        &self.real_time_data
    }

    /// MDX (OLAP cube) metadata collected from the workbook globals
    /// `METADATA` production; empty when the workbook carries none.
    ///
    /// The metadata is inert: connection names and MDX unique names are stored
    /// verbatim and no OLAP server is ever contacted.
    pub fn mdx_metadata(&self) -> &crate::mdx_metadata::MdxMetadata {
        &self.mdx_metadata
    }

    /// Web pages published from the workbook globals, in record order.
    ///
    /// The records are inert: destination URLs and paths are never opened,
    /// resolved, or fetched.
    pub fn web_publications(&self) -> &[crate::web_pub::WebPub] {
        &self.web_publications
    }

    /// Built-in and custom function categories, when the FNGROUPS collection exists.
    pub fn function_groups(&self) -> Option<&crate::function_group::FunctionGroups> {
        self.function_groups.as_ref()
    }

    /// Inert supporting-book links and cached external cell values.
    pub fn external_links(&self) -> &crate::external_link::Links {
        &self.external_links
    }

    /// Parsed workbook PivotCache streams ordered by their one-based stream ID.
    pub fn pivot_caches(&self) -> &[crate::PivotCache] {
        &self.pivot_caches
    }

    /// Global PivotCache ordinal-to-storage-stream map from SXStreamID records.
    pub fn pivot_cache_stream_ids(&self) -> &[u16] {
        &self.pivot_cache_stream_ids
    }

    /// Resolves a worksheet PivotTable's global cache link.
    pub fn pivot_cache_for_table(
        &self,
        table: &pivot_table::PivotTable,
    ) -> Result<&crate::PivotCache> {
        let stream_id = *self
            .pivot_cache_stream_ids
            .get(usize::from(table.cache_index()))
            .ok_or_else(|| Error::InvalidRecord {
                record_type: pivot_table::SXVIEW_TYPE,
                message: "PivotTable global cache index is out of range".to_string(),
            })?;
        self.pivot_caches
            .iter()
            .find(|cache| cache.stream_id() == stream_id)
            .ok_or_else(|| Error::InvalidRecord {
                record_type: pivot_table::SXVIEW_TYPE,
                message: "PivotTable SXStreamID has no matching cache storage".to_string(),
            })
    }

    /// Parsed PivotTables on one worksheet.
    pub fn worksheet_pivot_tables(&self, index: usize) -> Result<&[pivot_table::PivotTable]> {
        Ok(self.xls_worksheet(index)?.pivot_tables())
    }

    /// Access the typed `Worksheet` at the given index.
    ///
    /// This provides access to XLS-specific data (protection, comments,
    /// autofilter, pivot tables) that is not exposed through the generic
    /// `WorkbookTrait` / `Worksheet` trait.
    pub fn xls_worksheet(&self, index: usize) -> Result<&Worksheet> {
        self.worksheets
            .get(index)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet index {}", index)))
    }

    /// All workbook sheet directory entries in tab order.
    pub fn sheets(&self) -> &[SheetMetadata] {
        &self.sheets
    }

    /// Sheet directory entry at a workbook tab index.
    pub fn sheet(&self, index: usize) -> Option<&SheetMetadata> {
        self.sheets.get(index)
    }

    /// Case-insensitive sheet directory lookup.
    pub fn sheet_by_name(&self, name: &str) -> Option<&SheetMetadata> {
        self.sheets
            .iter()
            .find(|sheet| sheet.name().eq_ignore_ascii_case(name))
    }

    /// Total number of cell references represented by the workbook SST.
    pub fn shared_string_reference_count(&self) -> u32 {
        self.shared_string_reference_count
    }

    pub fn formatting(&self) -> &Formatting {
        &self.formatting
    }

    pub fn date_system(&self) -> DateSystem {
        self.formatting.date_system()
    }

    pub fn protection(&self) -> &protection::WorkbookProtection {
        &self.protection
    }

    pub fn calculation(&self) -> &crate::calculation::WorkbookCalculation {
        &self.calculation
    }

    pub fn environment(&self) -> &crate::environment::WorkbookEnvironment {
        &self.environment
    }

    /// Workbook extension flags from the `BookExt` record, when present.
    pub fn book_ext(&self) -> Option<&crate::book_ext::BookExt> {
        self.book_ext.as_ref()
    }

    /// Cell-style extensions from `StyleExt` records, in record order.
    pub fn style_extensions(&self) -> &[crate::style_ext::StyleExt] {
        &self.style_extensions
    }

    /// The document theme from the `Theme` record, when present.
    pub fn theme(&self) -> Option<&crate::theme::Theme> {
        self.theme.as_ref()
    }

    /// Strictly access the user recorded as last creating, opening, or modifying the workbook.
    ///
    /// Noncanonical legacy producer variants are deferred until this metadata is requested.
    pub fn write_access(&self) -> Result<Option<&crate::access::WriteAccess>> {
        match &self.write_access {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(Error::InvalidRecord {
                record_type: crate::access::WRITE_ACCESS_RECORD_TYPE,
                message: format!("invalid WriteAccess metadata: {error}"),
            }),
        }
    }

    /// Default table and PivotTable style catalog, when present.
    pub fn table_styles(&self) -> Option<&crate::table_styles::TableStyles> {
        self.table_styles.as_ref()
    }

    pub fn number_formats(&self) -> &[NumberFormat] {
        self.formatting.number_formats()
    }

    pub fn extended_formats(&self) -> &[ExtendedFormat] {
        self.formatting.extended_formats()
    }

    /// Global differential formats referenced by custom table-style elements.
    pub fn differential_formats(&self) -> &[crate::differential_format::DifferentialFormat] {
        self.formatting.differential_formats()
    }

    pub fn differential_format(
        &self,
        id: crate::table_styles::DifferentialFormatId,
    ) -> Option<&crate::differential_format::DifferentialFormat> {
        self.formatting.differential_format(id)
    }

    /// Resolves an XF's effective property families through its parent StyleXF.
    pub fn effective_extended_format(
        &self,
        index: u16,
    ) -> Option<crate::number_format::EffectiveExtendedFormat<'_>> {
        self.formatting.effective_extended_format(index)
    }

    /// Workbook color palette, using BIFF8 defaults when no `Palette` record exists.
    pub fn palette(&self) -> &crate::palette::Palette {
        &self.palette
    }

    /// Font records in physical workbook order.
    pub fn fonts(&self) -> &[crate::font::Font] {
        &self.fonts
    }

    /// Resolve a BIFF8 logical font index. Index 4 is reserved and returns `None`.
    pub fn font(&self, index: u16) -> Option<&crate::font::Font> {
        self.fonts.iter().find(|font| font.index() == index)
    }

    /// Resolve a font's color through the workbook palette.
    pub fn font_color(&self, index: u16) -> Option<crate::palette::Color> {
        self.font(index)
            .and_then(|font| self.palette.color(font.color_index()))
    }

    /// Resolves the global Font record referenced by an XF record.
    pub fn extended_format_font(&self, format: &ExtendedFormat) -> Option<&crate::font::Font> {
        self.font(format.font_index())
    }

    /// Rich-text and phonetic properties for a shared-string index.
    ///
    /// Returns `None` for an out-of-range index and for an ordinary string
    /// without either optional BIFF8 payload.
    pub fn shared_string_properties(&self, index: u32) -> Option<&SharedStringProperties> {
        self.shared_string_properties
            .as_ref()?
            .get(index as usize)?
            .as_deref()
    }

    /// Rich-text and phonetic properties for a cell backed by `LabelSst`.
    pub fn shared_string_properties_for_cell(
        &self,
        cell: &Cell,
    ) -> Option<&SharedStringProperties> {
        self.shared_string_properties(cell.shared_string_index()?)
    }

    /// Non-macro internal defined names in `Lbl` record order.
    pub fn defined_names(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Every internal `Lbl`, including inert macro and procedure metadata.
    pub fn defined_name_records(&self) -> &[DefinedName] {
        &self.defined_name_records
    }

    /// The built-in `_FilterDatabase` defined name scoped to a zero-based
    /// sheet index, if present.
    ///
    /// Its rendered formula describes the AutoFilter cell range.
    pub fn filter_database_name(&self, sheet_index: usize) -> Option<&DefinedName> {
        self.defined_names.iter().find(|defined_name| {
            defined_name.kind == DefinedNameKind::BuiltIn(BuiltInName::FilterDatabase)
                && defined_name.scope == NameScope::Worksheet(sheet_index)
        })
    }

    /// Case-insensitive name lookup with sheet-local-before-workbook precedence.
    /// Duplicate definitions use the last matching `Lbl` record.
    pub fn defined_name(&self, name: &str, sheet_index: Option<usize>) -> Option<&DefinedName> {
        if let Some(sheet_index) = sheet_index
            && let Some(local) = self.defined_names.iter().rev().find(|defined_name| {
                defined_name.scope == NameScope::Worksheet(sheet_index)
                    && names_equal(&defined_name.name, name)
            })
        {
            return Some(local);
        }
        self.defined_names.iter().rev().find(|defined_name| {
            defined_name.scope == NameScope::Workbook && names_equal(&defined_name.name, name)
        })
    }

    /// Built-in print area for a worksheet, if present.
    pub fn print_area(&self, sheet_index: usize) -> Option<&DefinedName> {
        self.built_in_sheet_name(sheet_index, BuiltInName::PrintArea)
    }

    /// Built-in print-title rows/columns for a worksheet, if present.
    pub fn print_titles(&self, sheet_index: usize) -> Option<&DefinedName> {
        self.built_in_sheet_name(sheet_index, BuiltInName::PrintTitles)
    }

    fn built_in_sheet_name(
        &self,
        sheet_index: usize,
        built_in: BuiltInName,
    ) -> Option<&DefinedName> {
        self.defined_names.iter().rev().find(|defined_name| {
            defined_name.scope == NameScope::Worksheet(sheet_index)
                && defined_name.kind == DefinedNameKind::BuiltIn(built_in)
        })
    }
}

impl<R: Read + Seek + std::fmt::Debug + Send + Sync> litchi_core::sheet::WorkbookTrait
    for Workbook<R>
{
    fn active_worksheet(&self) -> SheetResult<Box<dyn SheetTrait + '_>> {
        if self.worksheets.is_empty() {
            return Err(Box::new(Error::WorksheetNotFound(
                "No worksheets found".to_string(),
            )));
        }
        // Return reference instead of clone - zero-copy!
        Ok(Box::new(&self.worksheets[0]))
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.worksheet_names
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn SheetTrait + '_>> {
        for worksheet in &self.worksheets {
            if worksheet.name() == name {
                // Return reference instead of clone - zero-copy!
                return Ok(Box::new(worksheet));
            }
        }
        Err(Box::new(Error::WorksheetNotFound(name.to_string())))
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn SheetTrait + '_>> {
        if index >= self.worksheets.len() {
            return Err(Box::new(Error::WorksheetNotFound(format!(
                "Index {} out of bounds",
                index
            ))));
        }
        // Return reference instead of clone - zero-copy!
        Ok(Box::new(&self.worksheets[index]))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIteratorTrait<'_> + '_> {
        Box::new(WorksheetIterator {
            worksheets: self.worksheets.iter().collect(),
            index: 0,
        })
    }

    fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        0 // Default to first sheet
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904_date_system
    }
}

/// Worksheet iterator for XLS workbooks
struct WorksheetIterator<'a> {
    worksheets: Vec<&'a Worksheet>,
    index: usize,
}

impl<'a> WorksheetIteratorTrait<'a> for WorksheetIterator<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn SheetTrait + 'a>>> {
        if self.index >= self.worksheets.len() {
            None
        } else {
            let worksheet = self.worksheets[self.index];
            self.index += 1;
            // Return reference instead of clone - zero-copy!
            Some(Ok(Box::new(worksheet)))
        }
    }
}

fn names_equal(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}
