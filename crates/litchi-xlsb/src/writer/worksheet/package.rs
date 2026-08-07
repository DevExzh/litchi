//! XLSB worksheet package/stream assembly.

use crate::package::error::{Error, Result};
use crate::raw::{Writer, kind};
use std::io::Write;

use super::model::MutableWorksheet;

impl MutableWorksheet {
    /// Validate and fully encode the optional sparkline block without
    /// publishing any worksheet or relationship state.
    pub(crate) fn stage_sparkline_block(&self) -> Result<Option<Vec<u8>>> {
        self.sparkline_groups
            .as_ref()
            .map(|groups| {
                crate::sparkline::encode_block(groups, self.sparkline_limits).map_err(|error| {
                    Error::InvalidFormula(format!("unable to encode sparkline groups: {error}"))
                })
            })
            .transpose()
    }

    /// Write worksheet to binary format
    ///
    /// Following Excel's required structure
    #[cfg(test)]
    pub(crate) fn write<W: Write>(
        &self,
        writer: &mut Writer<W>,
        shared_strings: &mut crate::writer::MutableSharedStringsWriter,
    ) -> Result<()> {
        let sparkline_block = self.stage_sparkline_block()?;
        self.write_with_sparkline_block(writer, shared_strings, sparkline_block.as_deref())
    }

    /// Serialize with a block staged during workbook-wide validation.
    pub(crate) fn write_with_sparkline_block<W: Write>(
        &self,
        writer: &mut Writer<W>,
        shared_strings: &mut crate::writer::MutableSharedStringsWriter,
        sparkline_block: Option<&[u8]>,
    ) -> Result<()> {
        // Write BrtBeginSheet
        writer.write_record(kind::BEGIN_SHEET, &[])?;

        // Write worksheet properties and basic formatting information.
        self.write_ws_properties(writer)?;

        // Write worksheet dimensions
        self.write_dimensions(writer)?;

        // Write worksheet views (minimal SheetJS-style layout)
        self.write_ws_views(writer)?;

        // Write sheet formatting properties (BrtSheetFormatPr)
        self.write_sheet_format_pr(writer)?;

        // Column information (BrtBeginColInfos / BrtColInfo / BrtEndColInfos)
        self.write_col_infos(writer)?;

        // Write sheet data
        writer.write_record(kind::BEGIN_SHEET_DATA, &[])?;
        self.write_cells(writer, shared_strings)?;
        writer.write_record(kind::END_SHEET_DATA, &[])?;

        // Sheet protection (BrtSheetProtection) - minimal skeleton mirroring
        // SheetJS and [MS-XLSB] examples.
        self.write_sheet_protection(writer)?;

        // AutoFilter skeleton (BrtBeginAFilter / BrtEndAFilter).
        self.write_auto_filter(writer)?;

        // Write merged cells if present
        if !self.merged_cells.is_empty() {
            self.write_merged_cells(writer)?;
        }

        // Write hyperlinks if present
        if !self.hyperlinks.is_empty() {
            self.write_hyperlinks(writer)?;
        }

        // Write data validations if present
        if !self.data_validations.is_empty() {
            crate::writer::data_validation::write_data_validations(
                writer,
                &self.data_validations,
                self.data_validation_settings,
                self.data_validation14_settings,
            )?;
        }

        // Write conditional formatting if present
        if !self.conditional_formattings.is_empty() {
            crate::conditional_formatting::write_conditional_formattings(
                writer,
                &self.conditional_formattings,
            )?;
        }

        if !self.web_extension_bindings.is_empty() {
            writer.write_record(kind::BEGIN_WEB_EXTENSIONS, &[])?;
            for binding in &self.web_extension_bindings {
                writer.write_record(kind::WEB_EXTENSION, &binding.to_payload()?)?;
            }
            writer.write_record(kind::END_WEB_EXTENSIONS, &[])?;
        }

        if self.has_drawing_objects() {
            let rel_id = self.drawing_rel_id.as_deref().ok_or_else(|| {
                Error::InvalidFormula(
                    "worksheet drawing objects lack a Drawings relationship ID".to_string(),
                )
            })?;
            let mut payload = Vec::with_capacity(4 + rel_id.len() * 2);
            Writer::new(&mut payload).write_wide_string(rel_id)?;
            writer.write_record(kind::DRAWING, &payload)?;
        }

        // Write table references (BrtBeginListParts / BrtListPart /
        // BrtEndListParts) after all other sheet features.
        if !self.tables.is_empty() {
            if self.table_rel_ids.len() != self.tables.len() {
                return Err(Error::InvalidFormula(
                    "worksheet tables lack relationship IDs from the workbook writer".to_string(),
                ));
            }
            crate::package::table::write::write_list_parts(writer, &self.table_rel_ids)?;
        }

        if let Some(block) = sparkline_block {
            writer.get_mut().write_all(block)?;
        }

        // Write BrtEndSheet
        writer.write_record(kind::END_SHEET, &[])?;

        Ok(())
    }
}
