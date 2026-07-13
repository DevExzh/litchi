//! Workbook implementation for XLSB files

use crate::xlsb::error::XlsbResult;
use crate::xlsb::formula::{
    FormulaExternalBook, FormulaExternalSheet, FormulaResolutionContext, FormulaSupportingLink,
    excel_name_eq,
};
use crate::xlsb::named_ranges::{NamedRange, validate_defined_name};
use crate::xlsb::records::{XlsbRecordIter, record_types};
use crate::xlsb::worksheet::XlsbWorksheet;
use litchi_core::binary;
use litchi_core::sheet::{Result, Worksheet as SheetTrait, WorksheetIterator};
use litchi_opc::OpcPackage;
use std::io::{BufReader, Cursor, Read, Seek};

/// XLSB workbook implementation
#[allow(dead_code)]
pub struct XlsbWorkbook {
    package: OpcPackage,
    worksheets: Vec<XlsbWorksheet>,
    formula_context: FormulaResolutionContext,
    shared_strings: Vec<String>,
    is_1904: bool,
}

impl std::fmt::Debug for XlsbWorkbook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("XlsbWorkbook")
            .field("worksheet_names", &self.formula_context.worksheet_names)
            .field("shared_strings_count", &self.shared_strings.len())
            .field("is_1904", &self.is_1904)
            .finish()
    }
}

impl XlsbWorkbook {
    /// Workbook and sheet-scoped defined names in `PtgName` index order.
    pub fn defined_names(&self) -> &[String] {
        &self.formula_context.defined_names
    }

    /// Open an XLSB workbook from a reader
    pub fn new<R: Read + Seek>(reader: R) -> XlsbResult<Self> {
        let package = OpcPackage::from_reader(reader)?;
        let mut workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            is_1904: false,
        };

        workbook.load_workbook_info()?;
        workbook.load_shared_strings()?;

        Ok(workbook)
    }

    /// Create an XLSB workbook from an already-parsed OPC package.
    ///
    /// This is used for single-pass parsing where the OPC package has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `package` - An already-parsed OPC package
    pub fn from_opc_package(package: OpcPackage) -> XlsbResult<Self> {
        let mut workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            is_1904: false,
        };

        workbook.load_workbook_info()?;
        workbook.load_shared_strings()?;

        Ok(workbook)
    }

    /// Load workbook information from workbook.bin
    fn load_workbook_info(&mut self) -> XlsbResult<()> {
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;

        let blob = workbook_part.blob();
        let mut iter = XlsbRecordIter::new(BufReader::new(blob));
        let mut worksheet_names = Vec::new();
        let mut supporting_links = Vec::new();
        let mut external_sheets = Vec::new();
        let mut external_link_rel_ids = Vec::new();
        let mut defined_names = Vec::new();
        Self::read_workbook(
            &mut iter,
            &mut worksheet_names,
            &mut supporting_links,
            &mut external_sheets,
            &mut external_link_rel_ids,
            &mut defined_names,
            &mut self.is_1904,
        )?;
        let external_link_uris = external_link_rel_ids
            .iter()
            .map(|rel_id| {
                let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
                    crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "BrtSupBookSrc relationship {rel_id:?} is missing"
                    ))
                })?;
                if relationship.is_external() {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "BrtSupBookSrc relationship {rel_id:?} is external"
                    )));
                }
                relationship.target_partname().map_err(Into::into)
            })
            .collect::<XlsbResult<Vec<_>>>()?;
        let external_books = external_link_uris
            .iter()
            .map(|uri| self.load_external_book(uri))
            .collect::<XlsbResult<Vec<_>>>()?;
        self.formula_context = FormulaResolutionContext {
            worksheet_names: worksheet_names.into(),
            supporting_links: supporting_links.into(),
            external_sheets: external_sheets.into(),
            external_books: external_books.into(),
            defined_names: defined_names.into(),
            current_sheet: None,
        };

        Ok(())
    }

    /// Load shared strings from xl/sharedStrings.bin
    fn load_shared_strings(&mut self) -> XlsbResult<()> {
        let shared_strings_uri = litchi_opc::PackURI::new("/xl/sharedStrings.bin")?;
        if let Ok(shared_strings_part) = self.package.get_part(&shared_strings_uri) {
            let blob = shared_strings_part.blob();
            let mut iter = XlsbRecordIter::new(BufReader::new(blob));
            Self::read_shared_strings(&mut iter, &mut self.shared_strings)?;
        }

        Ok(())
    }

    /// Get a worksheet by index (lazy loading)
    fn get_worksheet(&self, index: usize) -> XlsbResult<XlsbWorksheet> {
        if index >= self.formula_context.worksheet_names.len() {
            return Err(crate::error::OoxmlError::InvalidFormat(format!(
                "Worksheet index {} out of bounds",
                index
            ))
            .into());
        }

        let name = &self.formula_context.worksheet_names[index];
        // For now, assume worksheets are at xl/worksheets/sheet1.bin, sheet2.bin, etc.
        let sheet_path = format!("/xl/worksheets/sheet{}.bin", index + 1);
        let sheet_uri = litchi_opc::PackURI::new(&sheet_path)?;

        let sheet_part = self.package.get_part(&sheet_uri)?;
        let blob = sheet_part.blob();
        let cursor = Cursor::new(blob);
        Self::read_worksheet(
            cursor,
            name.clone(),
            &self.shared_strings,
            &self.formula_context,
            index,
        )
    }

    /// Read shared strings from SST
    fn read_shared_strings(
        iter: &mut XlsbRecordIter<impl Read>,
        strings: &mut Vec<String>,
    ) -> XlsbResult<()> {
        for record in iter.by_ref() {
            let record = record?;
            match record.header.record_type {
                record_types::BEGIN_SST => {
                    // SST header, continue reading
                },
                record_types::SST_ITEM => {
                    if let Ok(sst_item) = crate::xlsb::records::SstItemRecord::parse(&record.data) {
                        strings.push(sst_item.string);
                    }
                },
                record_types::END_SST => {
                    break;
                },
                _ => {
                    // Skip other records
                },
            }
        }
        Ok(())
    }

    /// Read workbook structure
    fn read_workbook(
        iter: &mut XlsbRecordIter<impl Read>,
        worksheet_names: &mut Vec<String>,
        supporting_links: &mut Vec<FormulaSupportingLink>,
        external_sheets: &mut Vec<FormulaExternalSheet>,
        external_link_rel_ids: &mut Vec<String>,
        defined_names: &mut Vec<String>,
        is_1904: &mut bool,
    ) -> XlsbResult<()> {
        for record in iter.by_ref() {
            let record = record?;
            match record.header.record_type {
                record_types::WORKBOOK_PROP => {
                    if let Ok(prop) = crate::xlsb::records::WorkbookPropRecord::parse(&record.data)
                    {
                        *is_1904 = prop.is_date1904;
                    }
                },
                record_types::BUNDLE_SH => {
                    let bundle_sh = crate::xlsb::records::BundleSheetRecord::parse(&record.data)?;
                    worksheet_names.push(bundle_sh.name);
                },
                record_types::SUP_SELF => {
                    supporting_links.push(FormulaSupportingLink::SelfWorkbook);
                },
                record_types::SUP_SAME => {
                    supporting_links.push(FormulaSupportingLink::SameSheet);
                },
                record_types::SUP_BOOK_SRC => {
                    let (rel_id, consumed) = crate::xlsb::records::wide_str_with_len(&record.data)?;
                    if rel_id.is_empty() || consumed != record.data.len() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupBookSrc has an invalid relationship ID".to_string(),
                        ));
                    }
                    let book_index = u32::try_from(external_link_rel_ids.len()).map_err(|_| {
                        crate::xlsb::error::XlsbError::InvalidFormula(
                            "external-link count overflow".to_string(),
                        )
                    })?;
                    external_link_rel_ids.push(rel_id);
                    supporting_links.push(FormulaSupportingLink::ExternalWorkbook(book_index));
                },
                record_types::SUP_ADDIN => {
                    supporting_links.push(FormulaSupportingLink::AddIn);
                },
                record_types::EXTERN_SHEET => {
                    Self::parse_extern_sheet(&record.data, external_sheets)?;
                },
                record_types::NAME => {
                    let named_range = NamedRange::parse(&record.data)?;
                    if named_range
                        .sheet_id
                        .is_some_and(|index| index as usize >= worksheet_names.len())
                    {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                            "BrtName {} has invalid sheet scope {:?}",
                            named_range.name, named_range.sheet_id
                        )));
                    }
                    defined_names.push(named_range.name);
                },
                _ => {
                    // Skip other records
                },
            }
        }
        Ok(())
    }

    /// Read a worksheet
    fn read_worksheet(
        cursor: Cursor<&[u8]>,
        name: String,
        shared_strings: &[String],
        formula_context: &FormulaResolutionContext,
        sheet_index: usize,
    ) -> XlsbResult<XlsbWorksheet> {
        let mut worksheet = XlsbWorksheet::new(name);
        let iter = crate::xlsb::records::RecordIter::<std::io::Cursor<&[u8]>>::from_cursor(cursor);
        let formula_context = formula_context.for_sheet(sheet_index);
        let mut cells_reader = crate::xlsb::cells_reader::XlsbCellsReader::new(
            iter,
            shared_strings,
            &formula_context,
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

        Ok(worksheet)
    }

    fn parse_extern_sheet(
        data: &[u8],
        external_sheets: &mut Vec<FormulaExternalSheet>,
    ) -> XlsbResult<()> {
        if data.len() < 4 {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let count = usize::try_from(binary::read_u32_le_at(data, 0)?).map_err(|_| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "BrtExternSheet count overflow".to_string(),
            )
        })?;
        if count >= 65_536 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtExternSheet count {count} exceeds 65,535"
            )));
        }
        let expected = 4usize
            .checked_add(count.checked_mul(12).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?)
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?;
        if data.len() != expected {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        external_sheets.reserve(count);
        for chunk in data[4..].chunks_exact(12) {
            external_sheets.push(FormulaExternalSheet {
                external_link: binary::read_u32_le_at(chunk, 0)?,
                first_sheet: binary::read_u32_le_at(chunk, 4)? as i32,
                last_sheet: binary::read_u32_le_at(chunk, 8)? as i32,
            });
        }
        Ok(())
    }

    fn load_external_book(&self, uri: &litchi_opc::PackURI) -> XlsbResult<FormulaExternalBook> {
        let part = self.package.get_part(uri)?;
        let mut iter = XlsbRecordIter::new(BufReader::new(part.blob()));
        let mut link_type = None;
        let mut target_key = String::new();
        let mut target_detail = String::new();
        let mut sheet_names = Vec::new();
        let mut defined_names = Vec::new();
        let mut saw_sup_tabs = false;
        // 0 = outside a name, 1 = expect formula, 2 = expect bits,
        // 3 = expect end (or optional DDE/OLE cached values).
        let mut sup_name_state = 0u8;
        let mut saw_end = false;

        for record in &mut iter {
            let record = record?;
            if saw_end {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "external link has records after BrtEndSupBook".to_string(),
                ));
            }
            if link_type.is_none() && record.header.record_type != record_types::BEGIN_SUP_BOOK {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "external link does not start with BrtBeginSupBook".to_string(),
                ));
            }
            match record.header.record_type {
                record_types::BEGIN_SUP_BOOK => {
                    if link_type.is_some() || record.data.len() < 10 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtBeginSupBook framing".to_string(),
                        ));
                    }
                    let kind = binary::read_u16_le_at(&record.data, 0)?;
                    let (first, consumed) =
                        crate::xlsb::records::wide_str_with_len(&record.data[2..])?;
                    let mut offset = 2 + consumed;
                    let (second, consumed) = if kind == 0 {
                        Self::parse_nullable_wide_string(&record.data[offset..])?
                    } else {
                        let (value, consumed) =
                            crate::xlsb::records::wide_str_with_len(&record.data[offset..])?;
                        (Some(value), consumed)
                    };
                    offset += consumed;
                    if offset != record.data.len() || kind > 2 || first.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtBeginSupBook payload".to_string(),
                        ));
                    }
                    if kind == 0 && second.is_some() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "external workbook BrtBeginSupBook string2 is not NULL".to_string(),
                        ));
                    }
                    link_type = Some(kind);
                    target_key = first;
                    target_detail = second.unwrap_or_default();
                },
                record_types::SUP_TABS => {
                    if link_type != Some(0) || saw_sup_tabs || sup_name_state != 0 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupTabs".to_string(),
                        ));
                    }
                    sheet_names = Self::parse_external_sheet_names(&record.data)?;
                    saw_sup_tabs = true;
                },
                record_types::SUP_NAME_START => {
                    let kind = link_type.ok_or_else(|| {
                        crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupNameStart precedes BrtBeginSupBook".to_string(),
                        )
                    })?;
                    if sup_name_state != 0 || (kind == 0 && !saw_sup_tabs) {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupNameStart".to_string(),
                        ));
                    }
                    let (name, consumed) = crate::xlsb::records::wide_str_with_len(&record.data)?;
                    if consumed != record.data.len() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupNameStart has trailing bytes".to_string(),
                        ));
                    }
                    validate_defined_name(&name)?;
                    if kind == 0 {
                        defined_names.push(name);
                        sup_name_state = 1;
                    } else {
                        sup_name_state = 2;
                    }
                },
                record_types::SUP_NAME_FORMULA => {
                    if link_type != Some(0) || sup_name_state != 1 || record.data.len() < 4 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupNameFmla".to_string(),
                        ));
                    }
                    let formula_len = usize::try_from(binary::read_u32_le_at(&record.data, 0)?)
                        .map_err(|_| {
                            crate::xlsb::error::XlsbError::InvalidFormula(
                                "BrtSupNameFmla size overflow".to_string(),
                            )
                        })?;
                    let expected = formula_len.checked_add(4).ok_or_else(|| {
                        crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupNameFmla size overflow".to_string(),
                        )
                    })?;
                    if record.data.len() != expected {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected,
                            found: record.data.len(),
                        });
                    }
                    sup_name_state = 2;
                },
                record_types::SUP_NAME_BITS => {
                    if sup_name_state != 2 || record.data.len() != 7 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupNameBits".to_string(),
                        ));
                    }
                    sup_name_state = 3;
                },
                record_types::SUP_NAME_END => {
                    if sup_name_state != 3 || !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtSupNameEnd".to_string(),
                        ));
                    }
                    sup_name_state = 0;
                },
                record_types::END_SUP_BOOK => {
                    if !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected: 0,
                            found: record.data.len(),
                        });
                    }
                    if sup_name_state != 0 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtEndSupBook occurs inside an external-name block".to_string(),
                        ));
                    }
                    saw_end = true;
                },
                _ => {
                    if link_type == Some(0) && sup_name_state != 0 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected record inside an external defined name".to_string(),
                        ));
                    }
                },
            }
        }
        let kind = link_type.ok_or_else(|| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "external link has no BrtBeginSupBook".to_string(),
            )
        })?;
        if !saw_end {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "external link has no BrtEndSupBook".to_string(),
            ));
        }
        if kind == 0 && !saw_sup_tabs {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "external workbook link has no BrtSupTabs".to_string(),
            ));
        }
        let target = if kind == 1 {
            format!("{target_key}:{target_detail}")
        } else {
            let relationship = part.rels().get(&target_key).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "external data relationship {target_key:?} is missing"
                ))
            })?;
            if !relationship.is_external() {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "external data relationship {target_key:?} is internal"
                )));
            }
            relationship.target_ref().to_string()
        };
        Ok(FormulaExternalBook {
            target,
            sheet_names: sheet_names.into(),
            defined_names: defined_names.into(),
            is_workbook: kind == 0,
        })
    }

    fn parse_external_sheet_names(data: &[u8]) -> XlsbResult<Vec<String>> {
        if data.len() < 4 {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let count = usize::try_from(binary::read_u32_le_at(data, 0)?).map_err(|_| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "external sheet-name count overflow".to_string(),
            )
        })?;
        if count >= 65_535 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "external sheet-name count {count} exceeds 65,534"
            )));
        }
        let mut names = Vec::with_capacity(count);
        let mut offset = 4;
        for _ in 0..count {
            let (name, consumed) = crate::xlsb::records::wide_str_with_len(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "external sheet-name size overflow".to_string(),
                )
            })?;
            let name_len = name.encode_utf16().count();
            if name_len == 0
                || name_len > 31
                || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
                || name.starts_with('\'')
                || name.ends_with('\'')
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "external sheet name {name:?} does not follow sheet-name grammar"
                )));
            }
            if names
                .iter()
                .any(|existing: &String| excel_name_eq(existing, &name))
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "duplicate external sheet name {name:?}"
                )));
            }
            names.push(name);
        }
        if offset != data.len() {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtSupTabs has {} trailing bytes",
                data.len() - offset
            )));
        }
        Ok(names)
    }

    fn parse_nullable_wide_string(data: &[u8]) -> XlsbResult<(Option<String>, usize)> {
        if data.len() < 4 {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        if binary::read_u32_le_at(data, 0)? == u32::MAX {
            Ok((None, 4))
        } else {
            let (value, consumed) = crate::xlsb::records::wide_str_with_len(data)?;
            Ok((Some(value), consumed))
        }
    }
}

impl litchi_core::sheet::WorkbookTrait for XlsbWorkbook {
    fn active_sheet_index(&self) -> usize {
        0
    }

    fn active_worksheet(&self) -> Result<Box<dyn SheetTrait + '_>> {
        self.worksheet_by_index(0)
    }

    fn worksheet_count(&self) -> usize {
        self.formula_context.worksheet_names.len()
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.formula_context.worksheet_names
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn SheetTrait + '_>> {
        let worksheet = self.get_worksheet(index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn SheetTrait + '_>> {
        for (i, ws_name) in self.formula_context.worksheet_names.iter().enumerate() {
            if ws_name == name {
                return self.worksheet_by_index(i);
            }
        }
        Err(Box::new(crate::error::OoxmlError::InvalidFormat(format!(
            "Worksheet '{}' not found",
            name
        ))))
    }

    fn worksheets<'a>(&'a self) -> Box<dyn WorksheetIterator<'a> + 'a> {
        Box::new(XlsbWorksheetIterator {
            workbook: self,
            index: 0,
        })
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904
    }
}

pub struct XlsbWorksheetIterator<'a> {
    workbook: &'a XlsbWorkbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for XlsbWorksheetIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetTrait + 'a>>> {
        if self.index < self.workbook.formula_context.worksheet_names.len() {
            match self.workbook.get_worksheet(self.index) {
                Ok(worksheet) => {
                    self.index += 1;
                    Some(Ok(Box::new(worksheet)))
                },
                Err(e) => {
                    self.index += 1; // Continue to next worksheet even on error
                    Some(Err(Box::new(e)))
                },
            }
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::formula::{FormulaConverter, FormulaParser};
    use crate::xlsb::writer::RecordWriter;
    use litchi_core::sheet::{Cell, Worksheet};
    use litchi_opc::part::Part;
    use litchi_opc::{BlobPart, PackURI};
    use std::fs::File;

    fn wide_string(value: &str) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
        for unit in value.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    fn external_link_records(records: &[(u16, Vec<u8>)]) -> Vec<u8> {
        let mut data = Vec::new();
        let mut writer = RecordWriter::new(&mut data);
        for (record_type, payload) in records {
            writer.write_record(*record_type, payload).unwrap();
        }
        data
    }

    fn parse_external_link(records: &[(u16, Vec<u8>)]) -> XlsbResult<FormulaExternalBook> {
        let uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
        let mut part = BlobPart::new(
            uri.clone(),
            "application/vnd.ms-excel.externalLink".to_string(),
            external_link_records(records),
        );
        part.rels_mut().add_relationship(
            "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath".to_string(),
            "Book.xlsx".to_string(),
            "rIdPath".to_string(),
            true,
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(part));
        let workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            is_1904: false,
        };
        workbook.load_external_book(&uri)
    }

    fn external_workbook_records() -> Vec<(u16, Vec<u8>)> {
        let mut begin = 0u16.to_le_bytes().to_vec();
        begin.extend_from_slice(&wide_string("rIdPath"));
        begin.extend_from_slice(&u32::MAX.to_le_bytes());
        let mut tabs = 1u32.to_le_bytes().to_vec();
        tabs.extend_from_slice(&wide_string("Data Sheet"));
        vec![
            (record_types::BEGIN_SUP_BOOK, begin),
            (record_types::SUP_TABS, tabs),
            (record_types::SUP_NAME_START, wide_string("Rate")),
            (record_types::SUP_NAME_FORMULA, 0u32.to_le_bytes().to_vec()),
            (record_types::SUP_NAME_BITS, vec![0; 7]),
            (record_types::SUP_NAME_END, Vec::new()),
            (record_types::END_SUP_BOOK, Vec::new()),
        ]
    }

    #[test]
    fn reads_formula_records_from_real_workbook_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/universal-content.xlsb"
        );
        let workbook = XlsbWorkbook::new(File::open(path).unwrap()).unwrap();
        let mut formula_cells = Vec::new();
        for index in 0..workbook.formula_context.worksheet_names.len() {
            let worksheet = workbook.get_worksheet(index).unwrap();
            if let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() {
                for row in min_row..=max_row {
                    for col in min_col..=max_col {
                        let Some(cell) = worksheet.get_cell(row, col) else {
                            continue;
                        };
                        if cell.is_formula() {
                            formula_cells.push((
                                worksheet.name().to_string(),
                                cell.coordinate(),
                                cell.value().clone(),
                                cell.formula_bytes().unwrap().to_vec(),
                            ));
                        }
                    }
                }
            }
        }
        assert_eq!(formula_cells.len(), 4);
        let formulas: Vec<_> = formula_cells
            .iter()
            .map(|cell| match &cell.2 {
                litchi_core::sheet::CellValue::Formula {
                    formula,
                    cached_value,
                    ..
                } => (cell.1.as_str(), formula.as_str(), cached_value.as_deref()),
                value => panic!("expected decoded formula, found {value:?}"),
            })
            .collect();
        assert_eq!(formulas[0].0, "C1");
        assert_eq!(formulas[0].1, "(2*3)");
        assert_eq!(formulas[1].1, "(2+3)");
        assert_eq!(formulas[2].1, "(2-3)");
        assert_eq!(formulas[3].1, "(C1+C2)");
        assert!(matches!(
            formulas[3].2,
            Some(litchi_core::sheet::CellValue::Float(11.0))
        ));
        assert!(formula_cells.iter().all(|cell| !cell.3.is_empty()));
    }

    #[test]
    fn reads_external_book_metadata_from_poi_corpus_when_available() {
        let path = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/poi/test-data/spreadsheet/bug66682.xlsb"
        ));
        if !path.exists() {
            return;
        }

        let package = OpcPackage::open(path).unwrap();
        let workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            is_1904: false,
        };
        let uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
        let book = workbook.load_external_book(&uri).unwrap();
        assert!(book.is_workbook);
        assert_eq!(book.target, "ab");
        assert_eq!(&*book.sheet_names, &["ab"]);
    }

    #[test]
    fn parses_external_workbook_sheet_and_name_metadata() {
        let book = parse_external_link(&external_workbook_records()).unwrap();
        assert!(book.is_workbook);
        assert_eq!(book.target, "Book.xlsx");
        assert_eq!(&*book.sheet_names, &["Data Sheet"]);
        assert_eq!(&*book.defined_names, &["Rate"]);
    }

    #[test]
    fn resolves_external_formula_tokens_from_package_relationships() {
        let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
        let mut workbook_data = Vec::new();
        {
            let mut writer = RecordWriter::new(&mut workbook_data);
            writer
                .write_record(record_types::SUP_BOOK_SRC, &wide_string("rIdExternal"))
                .unwrap();
            let mut extern_sheet = 1u32.to_le_bytes().to_vec();
            extern_sheet.extend_from_slice(&0u32.to_le_bytes());
            extern_sheet.extend_from_slice(&0u32.to_le_bytes());
            extern_sheet.extend_from_slice(&0u32.to_le_bytes());
            writer
                .write_record(record_types::EXTERN_SHEET, &extern_sheet)
                .unwrap();
        }
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
            workbook_data,
        );
        workbook_part.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
                .to_string(),
            "externalLinks/externalLink1.bin".to_string(),
            "rIdExternal".to_string(),
            false,
        );

        let external_uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
        let mut external_part = BlobPart::new(
            external_uri,
            "application/vnd.ms-excel.externalLink".to_string(),
            external_link_records(&external_workbook_records()),
        );
        external_part.rels_mut().add_relationship(
            "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath".to_string(),
            "Book.xlsx".to_string(),
            "rIdPath".to_string(),
            true,
        );

        let mut package = OpcPackage::new();
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(external_part));
        let workbook = XlsbWorkbook::from_opc_package(package).unwrap();

        let reference = FormulaParser::new(&[0x5A, 0, 0, 0, 0, 0, 0, 0, 0])
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(
                &reference,
                &workbook.formula_context
            )
            .unwrap(),
            "'[Book.xlsx]Data Sheet'!$A$1"
        );
        let name = FormulaParser::new(&[0x59, 0, 0, 1, 0, 0, 0])
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&name, &workbook.formula_context)
                .unwrap(),
            "'[Book.xlsx]'!Rate"
        );
    }

    #[test]
    fn rejects_malformed_external_workbook_record_sequences() {
        let mut duplicate_tabs = external_workbook_records();
        duplicate_tabs.insert(2, duplicate_tabs[1].clone());
        assert!(matches!(
            parse_external_link(&duplicate_tabs),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));

        let mut unclosed_name = external_workbook_records();
        unclosed_name.remove(5);
        assert!(matches!(
            parse_external_link(&unclosed_name),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));

        let mut trailing_record = external_workbook_records();
        trailing_record.push((record_types::SUP_NAME_END, Vec::new()));
        assert!(matches!(
            parse_external_link(&trailing_record),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));
    }
}
