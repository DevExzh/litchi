#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! XLSB workbook stream and record encoding.

use super::model::{SheetSlot, WorkbookWriter};
use crate::calc;
use crate::named_ranges::validate_name;
use crate::package::error::{Error, Result};
use crate::package::formula::ParsedFormula;
use crate::raw::{Writer, kind};
use std::io::Write;

impl WorkbookWriter {
    /// Write workbook-level defined names (BrtName records).
    fn write_named_ranges<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        for named_range in &self.named_ranges {
            if named_range.function {
                return Err(Error::UnsupportedFeature(format!(
                    "macro defined name {} cannot be emitted",
                    named_range.name
                )));
            }
            validate_name(&named_range.name)?;
            if let Some(sheet_id) = named_range.sheet_id
                && usize::try_from(sheet_id)
                    .ok()
                    .is_none_or(|index| index >= self.sheet_order.len())
            {
                return Err(Error::InvalidFormula(format!(
                    "defined name {} has invalid sheet scope {sheet_id}",
                    named_range.name
                )));
            }
            let formula = named_range.formula.as_ref().ok_or_else(|| {
                Error::InvalidFormula(format!("defined name {} has no formula", named_range.name))
            })?;
            let parsed_formula = ParsedFormula {
                rgce: formula.clone(),
                rgcb: Vec::new(),
            };
            let mut data = Vec::new();
            let mut temp_writer = Writer::new(&mut data);

            let mut flags = 0u32;
            if named_range.hidden {
                flags |= 0x0001;
            }
            temp_writer.write_u32(flags)?;
            temp_writer.write_u8(0)?; // chKey; zero for non-macro names

            temp_writer.write_u32(named_range.sheet_id.unwrap_or(u32::MAX))?;
            temp_writer.write_wide_string(&named_range.name)?;
            for byte in parsed_formula.to_bytes()? {
                temp_writer.write_u8(byte)?;
            }
            temp_writer.write_u32(u32::MAX)?; // NULL comment

            writer.write_record(kind::NAME, &data)?;
        }

        Ok(())
    }

    /// Write workbook structure.
    ///
    /// The record order is based on the minimal SheetJS `write_wb_bin`
    /// implementation and [MS-XLSB] examples:
    ///
    /// ```text
    /// BrtBeginBook (0x0083)
    /// BrtFileVersion (0x0080)
    /// BrtWbProp (0x0099)
    /// [BrtBeginBookViews/BrtBookView/BrtEndBookViews]
    /// BrtBeginBundleShs / BrtBundleSh / BrtEndBundleShs (0x008F / 0x009C / 0x0090)
    /// [BrtBeginPivotCacheIDs / BrtBeginPivotCacheID / BrtEndPivotCacheID / BrtEndPivotCacheIDs]
    /// BrtBeginExternals / BrtSupSelf / BrtExternSheet / BrtEndExternals
    /// [BrtCalcProp]
    /// BrtEndBook (0x0084)
    /// ```
    ///
    /// The book views and calculation properties are currently written with a
    /// single default view and sensible defaults for calculation settings.
    pub(super) fn write_workbook<W: Write>(
        &self,
        writer: &mut Writer<W>,
        formula_sheet_ranges: &[(u32, u32)],
        pivot_cache_rel_ids: &[(u32, String)],
        external_link_rel_ids: &[String],
    ) -> Result<()> {
        // BrtBeginBook
        writer.write_record(kind::BEGIN_BOOK, &[])?;

        // BrtFileVersion - required by Excel
        self.write_file_version(writer)?;

        // BrtWbProp - basic workbook properties
        self.write_workbook_properties(writer)?;

        // Optional book views. We currently always emit a single default view
        // similar to SheetJS. This is small and helps some consumers which
        // expect explicit book view records.
        self.write_book_views(writer)?;

        // BrtBeginBundleShs / BrtBundleSh / BrtEndBundleShs - sheet metadata
        self.write_bundle_sheets(writer)?;

        // PivotCache identifiers, if any caches were attached.
        Self::write_pivot_cache_ids(writer, pivot_cache_rel_ids)?;

        // EXTERNALS block with self-references, mirroring SheetJS and
        // [MS-XLSB] examples. This creates a minimal but fully valid
        // extern sheet table for the workbook.
        self.write_externals(writer, formula_sheet_ranges, external_link_rel_ids)?;

        // Defined names (named ranges), if any.
        self.write_named_ranges(writer)?;

        // Basic calculation properties describing recalc behavior and
        // numerical tolerance. This is tiny and follows the spec example
        // values, so we emit it unconditionally.
        self.write_calc(writer)?;

        // BrtEndBook
        writer.write_record(kind::END_BOOK, &[])?;

        Ok(())
    }

    /// Write the PivotCache ID collection (BrtBeginPivotCacheIDs,
    /// MS-XLSB 2.4.170): one BrtBeginPivotCacheID record per attached cache,
    /// pairing the workbook cache identifier (`idSx`) with the relationship
    /// ID of its PivotCache Definition part.
    fn write_pivot_cache_ids<W: Write>(
        writer: &mut Writer<W>,
        pivot_cache_rel_ids: &[(u32, String)],
    ) -> Result<()> {
        if pivot_cache_rel_ids.is_empty() {
            return Ok(());
        }
        writer.write_record(kind::BEGIN_PIVOT_CACHE_IDS, &[])?;
        for (cache_id, rel_id) in pivot_cache_rel_ids {
            let mut data = Vec::with_capacity(rel_id.len() * 2 + 8);
            let mut temp_writer = Writer::new(&mut data);
            temp_writer.write_u32(*cache_id)?;
            temp_writer.write_wide_string(rel_id)?;
            writer.write_record(kind::BEGIN_PIVOT_CACHE_ID, &data)?;
            writer.write_record(kind::END_PIVOT_CACHE_ID, &[])?;
        }
        writer.write_record(kind::END_PIVOT_CACHE_IDS, &[])?;
        Ok(())
    }

    /// Write file version record (BrtFileVersion)
    /// This is REQUIRED for Excel to open the file
    fn write_file_version<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // Build structure per spec example (48 bytes total):
        // guidCodeName (16 zero bytes), stAppName ("xl"), stLastEdited ("4"),
        // stLowestEdited ("4"), stRupBuild ("4505")
        let mut data = Vec::with_capacity(48);
        let mut w = Writer::new(&mut data);

        // GUID (16 bytes of zeros)
        w.write_u32(0)?;
        w.write_u32(0)?;
        w.write_u32(0)?;
        w.write_u32(0)?;

        // stAppName: "xl"
        w.write_wide_string("xl")?;
        // stLastEdited: "4"
        w.write_wide_string("4")?;
        // stLowestEdited: "4"
        w.write_wide_string("4")?;
        // stRupBuild: "4505"
        w.write_wide_string("4505")?;

        writer.write_record(kind::FILE_VERSION, &data)?;
        Ok(())
    }

    /// Write workbook properties (BrtWbProp)
    fn write_workbook_properties<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Flags (4 bytes). We currently only support the 1904 date system
        // bit, mirroring the minimal SheetJS implementation:
        //   bit 0 (0x0000_0001) = f1904 (date1904)
        let mut flags: u32 = 0;
        if self.is_1904 {
            flags |= 0x0000_0001;
        }
        temp_writer.write_u32(flags)?;

        // Reserved/unused DWORD (4 bytes), set to 0.
        temp_writer.write_u32(0)?;

        // Code name (XLWideString). Use the standard VBA code name
        // "ThisWorkbook" as SheetJS and Excel commonly do.
        temp_writer.write_wide_string("ThisWorkbook")?;

        writer.write_record(kind::WORKBOOK_PROP, &data)?;
        Ok(())
    }

    /// Write book views (REQUIRED by Excel)
    fn write_book_views<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        if self.sheet_order.is_empty() {
            return Ok(());
        }
        writer.write_record(kind::BEGIN_BOOK_VIEWS, &[])?;

        // Write one default book view
        let mut view_data = Vec::new();
        let mut temp_writer = Writer::new(&mut view_data);

        // xWn (4), yWn (4), dxWn (4), dyWn (4)
        temp_writer.write_u32(0)?; // xWn
        temp_writer.write_u32(0)?; // yWn
        temp_writer.write_u32(0x00004E20)?; // dxWn (width)
        temp_writer.write_u32(0x00002710)?; // dyWn (height)

        // iTabRatio (4): 0 means auto
        temp_writer.write_u32(0)?;
        // itabFirst (4): first visible bundle sheet index
        temp_writer.write_u32(0)?;
        // itabCur (4): active sheet index. The public workbook API exposes
        // worksheets, so choose the first worksheet when a chart sheet was
        // inserted before it in the complete workbook tab order.
        let active_tab = self
            .sheet_order
            .iter()
            .position(|slot| matches!(slot, SheetSlot::Worksheet(_)))
            .unwrap_or(0);
        temp_writer.write_u32(u32::try_from(active_tab).map_err(|_| {
            Error::InvalidFormula("active workbook sheet index overflow".to_string())
        })?)?;

        // Flags (1 byte) - D/E/F bits set for scrollbars and tabs
        temp_writer.write_u8(0x78)?; // Total: 7*4 + 1 = 29 bytes

        writer.write_record(kind::BOOK_VIEW, &view_data)?;

        writer.write_record(kind::END_BOOK_VIEWS, &[])?;
        Ok(())
    }

    /// Write bundle sheets in workbook order.
    fn write_bundle_sheets<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_record(kind::BEGIN_BUNDLE_SHS, &[])?;

        for (i, slot) in self.sheet_order.iter().copied().enumerate() {
            let mut sheet_data = Vec::new();
            let mut temp_writer = Writer::new(&mut sheet_data);

            let state = match slot {
                SheetSlot::Worksheet(_) => 0,
                SheetSlot::ChartSheet(index) => match self.chart_sheets[index].metadata().state {
                    crate::package::chartsheet::State::Visible => 0,
                    crate::package::chartsheet::State::Hidden => 1,
                    crate::package::chartsheet::State::VeryHidden => 2,
                },
            };
            temp_writer.write_u32(state)?;
            // itabID (u32): unique sheet id (1-based)
            temp_writer
                .write_u32(u32::try_from(i + 1).map_err(|_| {
                    Error::InvalidFormula("sheet identifier overflow".to_string())
                })?)?;
            // RelID (XLWideString): rIdN
            temp_writer.write_wide_string(&format!("rId{}", i + 1))?;
            // strName (XLWideString): sheet name
            temp_writer.write_wide_string(self.sheet_name(slot))?;

            writer.write_record(kind::BUNDLE_SH, &sheet_data)?;
        }

        writer.write_record(kind::END_BUNDLE_SHS, &[])?;
        Ok(())
    }

    /// Write calculation properties (CALC_PROP, 0x009D)
    ///
    /// Spec example fields and order
    fn write_calc<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_header(kind::CALC_PROP, calc::LEN)?;
        calc::write(&self.calc, writer)?;
        Ok(())
    }

    /// Write externals section (self-references)
    ///
    /// Based on SheetJS implementation: always writes BrtSupSelf with BrtExternSheet
    /// This creates self-references for the workbook and all sheets.
    fn write_externals<W: Write>(
        &self,
        writer: &mut Writer<W>,
        formula_sheet_ranges: &[(u32, u32)],
        external_link_rel_ids: &[String],
    ) -> Result<()> {
        // BrtBeginExternals - no data
        writer.write_record(kind::BEGIN_EXTERNALS, &[])?;

        // BrtSupSelf - no data
        writer.write_record(kind::SUP_SELF, &[])?;

        for relationship_id in external_link_rel_ids {
            let mut data = Vec::with_capacity(4 + relationship_id.len() * 2);
            Writer::new(&mut data).write_wide_string(relationship_id)?;
            writer.write_record(kind::SUP_BOOK_SRC, &data)?;
        }

        // BrtExternSheet - self-references data
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        let sheet_count = self.sheet_order.len();

        // Total count: workbook and #REF entries, single-sheet entries, then
        // the distinct multi-sheet ranges referenced by formulas.
        let entry_count = sheet_count
            .checked_add(2)
            .and_then(|count| count.checked_add(formula_sheet_ranges.len()))
            .ok_or_else(|| {
                Error::InvalidFormula("BrtExternSheet entry count overflow".to_string())
            })?;
        if entry_count >= 65_536 {
            return Err(Error::InvalidFormula(format!(
                "BrtExternSheet entry count {entry_count} exceeds 65,535"
            )));
        }
        temp_writer.write_u32(u32::try_from(entry_count).map_err(|_| {
            Error::InvalidFormula("BrtExternSheet entry count overflow".to_string())
        })?)?;

        // First entry: workbook-level reference (0, -2, -2)
        temp_writer.write_u32(0)?;
        temp_writer.write_i32(-2)?;
        temp_writer.write_i32(-2)?;

        // Second entry: #REF! (0, -1, -1)
        temp_writer.write_u32(0)?;
        temp_writer.write_i32(-1)?;
        temp_writer.write_i32(-1)?;

        // Then for each sheet: (0, sheet_index, sheet_index)
        for i in 0..sheet_count {
            temp_writer.write_u32(0)?;
            temp_writer.write_i32(i as i32)?;
            temp_writer.write_i32(i as i32)?;
        }

        for &(first_sheet, last_sheet) in formula_sheet_ranges {
            if last_sheet < first_sheet
                || usize::try_from(last_sheet)
                    .ok()
                    .is_none_or(|last_sheet| last_sheet >= sheet_count)
            {
                return Err(Error::InvalidFormula(format!(
                    "invalid formula sheet range {first_sheet}..={last_sheet}"
                )));
            }
            temp_writer.write_u32(0)?;
            temp_writer.write_i32(i32::try_from(first_sheet).map_err(|_| {
                Error::InvalidFormula("first formula sheet index overflow".to_string())
            })?)?;
            temp_writer.write_i32(i32::try_from(last_sheet).map_err(|_| {
                Error::InvalidFormula("last formula sheet index overflow".to_string())
            })?)?;
        }

        writer.write_record(kind::EXTERN_SHEET, &data)?;

        // BrtEndExternals - no data
        writer.write_record(kind::END_EXTERNALS, &[])?;

        Ok(())
    }
}
