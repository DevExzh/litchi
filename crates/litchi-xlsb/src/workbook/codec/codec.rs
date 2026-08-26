//! Worksheet and shared-string BIFF12 stream materialization.

use super::super::model::Workbook;
use crate::package::error::Result;
use crate::package::formula::Context;
use crate::package::shared_strings::SharedString;
use crate::raw::{Records, kind};
use crate::sheet::Worksheet;
use litchi_core::binary;
use litchi_opc::constants::relationship_type;
use std::io::Cursor;

impl Workbook {
    pub fn worksheet(&self, index: usize) -> Result<Worksheet> {
        if index >= self.formula_context.worksheet_names.len() {
            return Err(crate::package::error::Error::InvalidFormat(format!(
                "Worksheet index {} out of bounds",
                index
            ))
            .into());
        }

        let name = &self.formula_context.worksheet_names[index];
        let rel_id = self
            .worksheet_rel_ids
            .get(index)
            .and_then(Option::as_deref)
            .ok_or_else(|| {
                crate::package::error::Error::UnsupportedFeature(format!(
                    "sheet {name:?} has no worksheet relationship"
                ))
            })?;
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;
        let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
            crate::package::error::Error::FileNotFound(format!(
                "relationship {rel_id:?} for sheet {name:?}"
            ))
        })?;
        if relationship.is_external() {
            return Err(crate::package::error::Error::UnsupportedFeature(format!(
                "sheet {name:?} has an external worksheet relationship"
            )));
        }
        let sheet_uri = relationship.target_partname()?;

        let sheet_part = self.package.get_part(&sheet_uri)?;
        let comments_uri = {
            let mut relationships = sheet_part
                .rels()
                .iter()
                .filter(|rel| rel.reltype() == relationship_type::COMMENTS);
            let first = relationships.next();
            if relationships.next().is_some() {
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "worksheet comments relationship".to_string(),
                    val: "multiple relationships".to_string(),
                });
            }
            match first {
                Some(rel) if rel.is_external() => {
                    return Err(crate::package::error::Error::UnsupportedFeature(
                        "external XLSB comments part".to_string(),
                    ));
                },
                Some(rel) => Some(rel.target_partname()?),
                None => None,
            }
        };
        let blob = sheet_part.blob();
        let cursor = Cursor::new(blob);
        let mut worksheet = Self::read_worksheet(
            cursor,
            name.clone(),
            &self.shared_strings,
            &self.formula_context,
            index,
            self.styles.cell_xfs.len(),
        )?;
        worksheet.set_scenarios(crate::package::scenarios::parse_worksheet(blob)?);
        worksheet.set_slicers(
            crate::slicer::package::load_views(&self.package, &sheet_uri)?.map(|part| part.views),
        );
        worksheet.set_timelines(
            crate::timeline::package::load_views(&self.package, &sheet_uri)?.map(|part| part.views),
        );
        worksheet.set_sparkline_groups(self.sparklines(index)?.groups().cloned());
        if let Some(uri) = comments_uri {
            let part = self.package.get_part(&uri)?;
            if !part.rels().is_empty() {
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "Comments part".to_string(),
                    val: "relationships are not permitted".to_string(),
                });
            }
            for comment in crate::comments::read(part.blob())? {
                worksheet.add_comment(comment);
            }
        }
        Ok(worksheet)
    }

    /// Read shared strings from SST
    pub(in crate::workbook) fn read_shared_strings(
        iter: &mut Records<'_>,
        strings: &mut Vec<SharedString>,
    ) -> Result<()> {
        let initial_count = strings.len();
        let mut expected_unique = None;
        let mut ended = false;
        for record in iter.by_ref() {
            let record = record?;
            match record.kind() {
                kind::BEGIN_SST => {
                    if expected_unique.is_some() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginSst".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    if record.payload().len() != 8 {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 8,
                            found: record.payload().len(),
                        });
                    }
                    let total = binary::read_u32_le_at(record.payload(), 0)?;
                    let unique = binary::read_u32_le_at(record.payload(), 4)?;
                    if total > 0x7FFF_FFFF || unique > total {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginSst counts".to_string(),
                            val: format!("total={total}, unique={unique}"),
                        });
                    }
                    expected_unique = Some(unique as usize);
                },
                kind::SST_ITEM => {
                    let expected = expected_unique.ok_or_else(|| {
                        crate::package::error::Error::Unrecognized {
                            typ: "BrtSSTItem".to_string(),
                            val: "record before BrtBeginSst".to_string(),
                        }
                    })?;
                    let found = strings.len() - initial_count;
                    if found >= expected {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtSSTItem count".to_string(),
                            val: format!("more than declared {expected}"),
                        });
                    }
                    strings.push(SharedString::parse(record.payload())?);
                },
                kind::END_SST => {
                    if !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 0,
                            found: record.payload().len(),
                        });
                    }
                    let expected = expected_unique.ok_or_else(|| {
                        crate::package::error::Error::Unrecognized {
                            typ: "BrtEndSst".to_string(),
                            val: "record before BrtBeginSst".to_string(),
                        }
                    })?;
                    let found = strings.len() - initial_count;
                    if found != expected {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtSSTItem count".to_string(),
                            val: format!("declared {expected}, found {found}"),
                        });
                    }
                    ended = true;
                    break;
                },
                _ => {
                    // Skip other records
                },
            }
        }
        if expected_unique.is_none() {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "SST stream".to_string(),
                val: "missing BrtBeginSst".to_string(),
            });
        }
        if !ended {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "SST stream".to_string(),
                val: "missing BrtEndSst".to_string(),
            });
        }
        Ok(())
    }

    pub(in crate::workbook) fn read_worksheet(
        cursor: Cursor<&[u8]>,
        name: String,
        shared_strings: &[SharedString],
        formula_context: &Context,
        sheet_index: usize,
        cell_xf_count: usize,
    ) -> Result<Worksheet> {
        let mut worksheet = Worksheet::new(name);
        let iter = crate::package::records::Stream::new(cursor);
        let formula_context = formula_context.for_sheet(sheet_index);
        let mut cells_reader = crate::package::cells_reader::CellsReader::new(
            iter,
            shared_strings,
            &formula_context,
            cell_xf_count,
        )?;

        // Read all cells
        while let Some(cell) = cells_reader.next_cell()? {
            worksheet.add_cell(cell);
        }

        // Transfer advanced features from reader to worksheet
        for merged in cells_reader.merged_cells {
            worksheet.add_merged_cell(merged);
        }
        for hyperlink in cells_reader.hyperlinks {
            worksheet.add_hyperlink(hyperlink);
        }
        worksheet.set_column_infos(cells_reader.column_infos);
        worksheet.set_row_infos(cells_reader.row_infos);
        worksheet.set_auto_filter(cells_reader.auto_filter);
        worksheet.set_sheet_protection(cells_reader.sheet_protection);
        worksheet.set_strong_sheet_protection(cells_reader.strong_sheet_protection);
        worksheet.set_data_validations(
            cells_reader.data_validation_settings,
            cells_reader.data_validation14_settings,
            cells_reader.data_validations,
        );
        worksheet.set_conditional_formattings(cells_reader.conditional_formattings);
        worksheet.set_web_extension_bindings(cells_reader.web_extension_bindings);
        worksheet.set_views(cells_reader.views);

        Ok(worksheet)
    }
}
