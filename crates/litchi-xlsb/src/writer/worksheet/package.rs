//! XLSB worksheet package/stream assembly.

use crate::package::error::{Error, Result};
use crate::raw::{Encoder, kind};
use std::io::Write;

use super::model::Worksheet;

impl Worksheet {
    /// Validate and fully encode the optional sparkline block without
    /// publishing any worksheet or relationship state.
    pub(crate) fn stage_sparkline_block(&self) -> Result<Option<Vec<u8>>> {
        self.sparkline_groups
            .as_ref()
            .map(|groups| {
                crate::sparkline::encode_block(groups, self.sparkline_limits)
                    .map_err(|_error| Error::InvalidFormula("unable to encode sparkline groups"))
            })
            .transpose()
    }

    /// Write worksheet to binary format
    ///
    /// Following Excel's required structure
    #[cfg(test)]
    pub(crate) fn write<W: Write>(
        &self,
        writer: &mut Encoder<W>,
        shared_strings: &mut crate::writer::SharedStringsWriter,
    ) -> Result<()> {
        let sparkline_block = self.stage_sparkline_block()?;
        self.write_with_sparkline_block(writer, shared_strings, sparkline_block.as_deref())
    }

    /// Serialize with a block staged during workbook-wide validation.
    pub(crate) fn write_with_sparkline_block<W: Write>(
        &self,
        writer: &mut Encoder<W>,
        shared_strings: &mut crate::writer::SharedStringsWriter,
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
                writer.limits(),
            )?;
        }

        // Write conditional formatting if present
        if !self.conditional_formattings.is_empty() {
            crate::conditional_formatting::write_conditional_formattings(
                writer,
                &self.conditional_formattings,
            )
            .map_err(|error| match error {
                crate::conditional_formatting::Error::UnsupportedFeature(_)
                | crate::conditional_formatting::Error::Unrecognized { .. } => Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "build_xlsb",
                    "the conditional format contains unsupported authored content",
                ),
                error @ (crate::conditional_formatting::Error::InvalidLength { .. }
                | crate::conditional_formatting::Error::InvalidFormula(_)
                | crate::conditional_formatting::Error::InvalidCellReference(_)
                | crate::conditional_formatting::Error::Encoding(_)
                | crate::conditional_formatting::Error::Wire(_)
                | crate::conditional_formatting::Error::Formula(_)
                | crate::conditional_formatting::Error::Io(_)) => error.into(),
            })?;
        }

        if !self.web_extension_bindings.is_empty() {
            writer.write_record(kind::BEGIN_WEB_EXTENSIONS, &[])?;
            for binding in &self.web_extension_bindings {
                writer.write_record(
                    kind::WEB_EXTENSION,
                    &binding.to_payload(crate::raw::Limits::DEFAULT)?,
                )?;
            }
            writer.write_record(kind::END_WEB_EXTENSIONS, &[])?;
        }

        if self.has_drawing_objects() {
            let rel_id = self.drawing_rel_id.as_deref().ok_or(Error::InvalidFormula(
                "worksheet drawing objects lack a Drawings relationship ID",
            ))?;
            let mut payload = Vec::with_capacity(4 + rel_id.len() * 2);
            Encoder::new(&mut payload, crate::raw::Limits::DEFAULT).write_wide_string(rel_id)?;
            writer.write_record(kind::DRAWING, &payload)?;
        }

        // Write table references (BrtBeginListParts / BrtListPart /
        // BrtEndListParts) after all other sheet features.
        if !self.tables.is_empty() {
            if self.table_rel_ids.len() != self.tables.len() {
                return Err(Error::InvalidFormula(
                    "worksheet tables lack relationship IDs from the workbook writer",
                ));
            }
            crate::package::table::write::write_list_parts(writer, &self.table_rel_ids)?;
        }

        if let Some(block) = sparkline_block {
            writer.write_record(kind::FRT_BEGIN, &[2, 14, 0, 0])?;
            writer.get_mut().write_all(block)?;
            writer.write_record(kind::FRT_END, &[])?;
        }

        // Write BrtEndSheet
        writer.write_record(kind::END_SHEET, &[])?;

        Ok(())
    }
}
