#![allow(
    clippy::cast_possible_wrap,
    clippy::expect_used,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error to this codec boundary"
)]

//! Typed workbook, external-link, PivotTable, and table record codecs.

use super::super::model::Workbook;
use crate::calc::{self, Props};
use crate::named_ranges::Definition;
use crate::package::error::Result;
use crate::package::formula::{
    ExternalSheet, SupportingLink, View, excel_name_eq, table::Definition as TableDefinition,
};
use crate::raw::{Records, kind};
use litchi_core::binary;

#[derive(Default)]
pub(in crate::workbook) struct ParsedWorkbookInfo {
    pub(in crate::workbook) worksheet_names: Vec<String>,
    pub(in crate::workbook) worksheet_rel_ids: Vec<Option<String>>,
    pub(in crate::workbook) worksheet_states: Vec<u32>,
    pub(in crate::workbook) active_catalog_position: Option<usize>,
    pub(in crate::workbook) supporting_links: Vec<SupportingLink>,
    pub(in crate::workbook) external_sheets: Vec<ExternalSheet>,
    pub(in crate::workbook) external_link_rel_ids: Vec<String>,
    pub(in crate::workbook) defined_names: Vec<String>,
    pub(in crate::workbook) is_1904: bool,
    pub(in crate::workbook) calc: Option<Props>,
}

impl Workbook {
    pub(in crate::workbook) fn read_workbook(iter: &mut Records<'_>) -> Result<ParsedWorkbookInfo> {
        let mut info = ParsedWorkbookInfo::default();
        let worksheet_names = &mut info.worksheet_names;
        let worksheet_rel_ids = &mut info.worksheet_rel_ids;
        let worksheet_states = &mut info.worksheet_states;
        let supporting_links = &mut info.supporting_links;
        let external_sheets = &mut info.external_sheets;
        let external_link_rel_ids = &mut info.external_link_rel_ids;
        let defined_names = &mut info.defined_names;
        let is_1904 = &mut info.is_1904;
        let mut book_views_open = false;
        let mut book_views_seen = false;
        let mut max_book_view_catalog_position: Option<usize> = None;
        for record in iter.by_ref() {
            let record = record?;
            match record.kind() {
                kind::BEGIN_BOOK_VIEWS => {
                    if !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 0,
                            found: record.payload().len(),
                        });
                    }
                    if book_views_seen || book_views_open {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginBookViews".to_string(),
                            val: "duplicate or nested book-view block".to_string(),
                        });
                    }
                    book_views_seen = true;
                    book_views_open = true;
                },
                kind::END_BOOK_VIEWS => {
                    if !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 0,
                            found: record.payload().len(),
                        });
                    }
                    if !book_views_open {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtEndBookViews".to_string(),
                            val: "book-view block has no matching begin record".to_string(),
                        });
                    }
                    book_views_open = false;
                },
                kind::BOOK_VIEW => {
                    if !book_views_open {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBookView".to_string(),
                            val: "record is outside the book-view block".to_string(),
                        });
                    }
                    const BOOK_VIEW_PAYLOAD_LEN: usize = 29;
                    let data = record.payload();
                    if data.len() != BOOK_VIEW_PAYLOAD_LEN {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: BOOK_VIEW_PAYLOAD_LEN,
                            found: data.len(),
                        });
                    }
                    let first =
                        usize::try_from(binary::read_u32_le_at(data, 20)?).map_err(|_| {
                            crate::package::error::Error::InvalidFormula(
                                "BrtBookView itabFirst does not fit usize".to_string(),
                            )
                        })?;
                    let active =
                        usize::try_from(binary::read_u32_le_at(data, 24)?).map_err(|_| {
                            crate::package::error::Error::InvalidFormula(
                                "BrtBookView itabCur does not fit usize".to_string(),
                            )
                        })?;
                    let maximum = first.max(active);
                    max_book_view_catalog_position = Some(
                        max_book_view_catalog_position
                            .map_or(maximum, |current| current.max(maximum)),
                    );
                    if info.active_catalog_position.is_none() {
                        info.active_catalog_position = Some(active);
                    }
                },
                kind::WORKBOOK_PROP => {
                    if let Ok(prop) =
                        crate::package::records::WorkbookPropRecord::parse(record.payload())
                    {
                        *is_1904 = prop.is_date1904;
                    }
                },
                kind::CALC_PROP => {
                    if info.calc.is_some() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtCalcProp".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    info.calc = Some(calc::read(record.payload())?);
                },
                kind::BUNDLE_SH => {
                    let bundle_sh =
                        crate::package::records::BundleSheetRecord::parse(record.payload())?;
                    if worksheet_names
                        .iter()
                        .any(|name| excel_name_eq(name, &bundle_sh.name))
                    {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBundleSh strName".to_string(),
                            val: format!("duplicate sheet name {:?}", bundle_sh.name),
                        });
                    }
                    worksheet_names.push(bundle_sh.name);
                    worksheet_rel_ids.push(bundle_sh.rel_id);
                    worksheet_states.push(bundle_sh.state);
                },
                kind::SUP_SELF => {
                    supporting_links.push(SupportingLink::SelfWorkbook);
                },
                kind::SUP_SAME => {
                    supporting_links.push(SupportingLink::SameSheet);
                },
                kind::SUP_BOOK_SRC => {
                    let (rel_id, consumed) =
                        crate::package::records::decode_string(record.payload())?;
                    if rel_id.is_empty() || consumed != record.payload().len() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtSupBookSrc has an invalid relationship ID".to_string(),
                        ));
                    }
                    let book_index = u32::try_from(external_link_rel_ids.len()).map_err(|_| {
                        crate::package::error::Error::InvalidFormula(
                            "external-link count overflow".to_string(),
                        )
                    })?;
                    external_link_rel_ids.push(rel_id);
                    supporting_links.push(SupportingLink::ExternalWorkbook(book_index));
                },
                kind::SUP_ADDIN => {
                    supporting_links.push(SupportingLink::AddIn);
                },
                kind::EXTERN_SHEET => {
                    Self::parse_extern_sheet(record.payload(), external_sheets)?;
                },
                kind::NAME => {
                    let named_range = Definition::parse(record.payload())?;
                    if named_range
                        .sheet_id
                        .is_some_and(|index| index as usize >= worksheet_names.len())
                    {
                        return Err(crate::package::error::Error::InvalidFormula(format!(
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
        if book_views_open {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "BrtBeginBookViews".to_string(),
                val: "missing BrtEndBookViews record".to_string(),
            });
        }
        if let Some(maximum) = max_book_view_catalog_position
            && maximum >= worksheet_names.len()
        {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtBookView sheet position {maximum} exceeds the {} available sheets",
                worksheet_names.len()
            )));
        }
        Ok(info)
    }

    fn parse_extern_sheet(data: &[u8], external_sheets: &mut Vec<ExternalSheet>) -> Result<()> {
        if data.len() < 4 {
            return Err(crate::package::error::Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let count = usize::try_from(binary::read_u32_le_at(data, 0)?).map_err(|_| {
            crate::package::error::Error::InvalidFormula(
                "BrtExternSheet count overflow".to_string(),
            )
        })?;
        if count >= 65_536 {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtExternSheet count {count} exceeds 65,535"
            )));
        }
        let expected = 4usize
            .checked_add(count.checked_mul(12).ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?)
            .ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?;
        if data.len() != expected {
            return Err(crate::package::error::Error::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        external_sheets.reserve(count);
        for chunk in data[4..].chunks_exact(12) {
            external_sheets.push(ExternalSheet {
                external_link: binary::read_u32_le_at(chunk, 0)?,
                first_sheet: binary::read_u32_le_at(chunk, 4)? as i32,
                last_sheet: binary::read_u32_le_at(chunk, 8)? as i32,
            });
        }
        Ok(())
    }

    pub(in crate::workbook) fn parse_nullable_wide_string(
        data: &[u8],
    ) -> Result<(Option<String>, usize)> {
        if data.len() < 4 {
            return Err(crate::package::error::Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        if binary::read_u32_le_at(data, 0)? == u32::MAX {
            Ok((None, 4))
        } else {
            let (value, consumed) = crate::package::records::decode_string(data)?;
            Ok((Some(value), consumed))
        }
    }

    pub(in crate::workbook) fn parse_pivot_cache_ids(data: &[u8]) -> Result<Vec<(u32, String)>> {
        let mut in_collection = false;
        let mut open_cache = false;
        let mut ended = false;
        let mut caches = Vec::new();
        for record in Records::new(data) {
            let record = record?;
            match record.kind() {
                kind::BEGIN_PIVOT_CACHE_IDS => {
                    if in_collection || ended {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "duplicate BrtBeginPivotCacheIDs collection".to_string(),
                        ));
                    }
                    in_collection = true;
                },
                kind::BEGIN_PIVOT_CACHE_ID => {
                    if !in_collection || open_cache || record.payload().len() < 8 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "malformed BrtBeginPivotCacheID nesting or payload".to_string(),
                        ));
                    }
                    let cache_id = binary::read_u32_le_at(record.payload(), 0)?;
                    let (rel_id, consumed) =
                        crate::package::records::decode_string(&record.payload()[4..])?;
                    if 4 + consumed != record.payload().len()
                        || rel_id.is_empty()
                        || rel_id.encode_utf16().count() > 255
                        || caches
                            .iter()
                            .any(|(existing, _): &(u32, String)| *existing == cache_id)
                    {
                        return Err(crate::package::error::Error::InvalidFormula(format!(
                            "invalid or duplicate PivotCache ID {cache_id}"
                        )));
                    }
                    caches.push((cache_id, rel_id));
                    open_cache = true;
                },
                kind::END_PIVOT_CACHE_ID => {
                    if !open_cache || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unbalanced BrtEndPivotCacheID".to_string(),
                        ));
                    }
                    open_cache = false;
                },
                kind::END_PIVOT_CACHE_IDS => {
                    if !in_collection || open_cache || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unbalanced BrtEndPivotCacheIDs".to_string(),
                        ));
                    }
                    in_collection = false;
                    ended = true;
                },
                _ => {},
            }
        }
        if in_collection || open_cache {
            return Err(crate::package::error::Error::InvalidFormula(
                "unterminated PivotCache ID collection".to_string(),
            ));
        }
        Ok(caches)
    }

    pub(in crate::workbook) fn parse_pivot_view(data: &[u8], sheet_index: usize) -> Result<View> {
        let mut view = None;
        for record in Records::new(data) {
            let record = record?;
            if record.kind() != kind::BEGIN_SX_VIEW {
                continue;
            }
            if view.is_some() || record.payload().len() < 36 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "PivotTable part has duplicate or truncated BrtBeginSXView".to_string(),
                ));
            }
            let cache_id = binary::read_u32_le_at(record.payload(), 28)?;
            let (name, consumed) = crate::package::records::decode_string(&record.payload()[32..])?;
            if consumed > record.payload().len() - 32 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "PivotTable view name overruns BrtBeginSXView".to_string(),
                ));
            }
            view = Some(View::try_new(cache_id, sheet_index, name)?);
        }
        view.ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "PivotTable part omits BrtBeginSXView".to_string(),
            )
        })
    }

    pub(in crate::workbook) fn parse_table_definition(
        data: &[u8],
        sheet_index: usize,
    ) -> Result<TableDefinition> {
        // Extension content is inert and source-owned by the XML Maps codec.
        // A fixed stack bounds adversarial nesting without allocating in this
        // formula-context discovery pass.
        const MAX_XML_EXTENSION_DEPTH: usize = 64;

        let mut table_header: Option<(u32, String, usize)> = None;
        let mut expected_columns = None;
        let mut columns = Vec::new();
        let mut in_column = false;
        let mut in_xml_properties = false;
        let mut saw_xml_properties = false;
        let mut xml_extension_stack = [kind::FRT_BEGIN; MAX_XML_EXTENSION_DEPTH];
        let mut xml_extension_depth = 0usize;
        let mut ended_columns = false;
        let mut ended_table = false;
        let mut iter = Records::new(data);
        for record in iter.by_ref() {
            let record = record?;
            if ended_table {
                return Err(crate::package::error::Error::InvalidFormula(
                    "XLSB table part contains records after BrtEndList".to_string(),
                ));
            }
            if in_xml_properties {
                match record.kind() {
                    kind::FRT_BEGIN | kind::AC_BEGIN => {
                        if xml_extension_depth == MAX_XML_EXTENSION_DEPTH {
                            return Err(crate::package::error::Error::InvalidFormula(
                                "BrtBeginListXmlCPr extension nesting exceeds 64".to_string(),
                            ));
                        }
                        xml_extension_stack[xml_extension_depth] = record.kind();
                        xml_extension_depth += 1;
                    },
                    kind::FRT_END | kind::AC_END => {
                        let Some(depth) = xml_extension_depth.checked_sub(1) else {
                            return Err(crate::package::error::Error::InvalidFormula(
                                "unmatched extension end in BrtBeginListXmlCPr".to_string(),
                            ));
                        };
                        let begin = xml_extension_stack[depth];
                        let expected = if begin == kind::FRT_BEGIN {
                            kind::FRT_END
                        } else {
                            kind::AC_END
                        };
                        if record.kind() != expected {
                            return Err(crate::package::error::Error::InvalidFormula(
                                "mismatched extension wrapper in BrtBeginListXmlCPr".to_string(),
                            ));
                        }
                        xml_extension_depth = depth;
                    },
                    kind::END_LIST_XML_CPR if xml_extension_depth == 0 => {
                        if !record.payload().is_empty() {
                            return Err(crate::package::error::Error::InvalidFormula(
                                "nonempty BrtEndListXmlCPr".to_string(),
                            ));
                        }
                        in_xml_properties = false;
                    },
                    kind::BEGIN_LIST
                    | kind::END_LIST
                    | kind::BEGIN_LIST_COLS
                    | kind::END_LIST_COLS
                    | kind::BEGIN_LIST_COL
                    | kind::END_LIST_COL
                    | kind::BEGIN_LIST_XML_CPR
                        if xml_extension_depth == 0 =>
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "table structure occurs inside BrtBeginListXmlCPr".to_string(),
                        ));
                    },
                    _ => {
                        // Unknown records and all records inside balanced FRT/AC
                        // wrappers are deliberately left to the source-owning
                        // XML Maps codec. This pass neither executes nor drops
                        // their bytes.
                    },
                }
                continue;
            }
            match record.kind() {
                kind::BEGIN_LIST => {
                    if table_header.is_some() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "XLSB table part contains duplicate BrtBeginList".to_string(),
                        ));
                    }
                    table_header = Some(Self::parse_table_header(record.payload())?);
                },
                kind::BEGIN_LIST_COLS => {
                    let (_, _, range_columns) = table_header.as_ref().ok_or_else(|| {
                        crate::package::error::Error::InvalidFormula(
                            "BrtBeginListCols precedes BrtBeginList".to_string(),
                        )
                    })?;
                    if expected_columns.is_some() || record.payload().len() != 4 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid or duplicate BrtBeginListCols".to_string(),
                        ));
                    }
                    let count = usize::try_from(binary::read_u32_le_at(record.payload(), 0)?)
                        .map_err(|_| {
                            crate::package::error::Error::InvalidFormula(
                                "table column count overflow".to_string(),
                            )
                        })?;
                    if count == 0 || count > 16_384 || count != *range_columns {
                        return Err(crate::package::error::Error::InvalidFormula(format!(
                            "table column count {count} disagrees with range width {range_columns}"
                        )));
                    }
                    expected_columns = Some(count);
                },
                kind::BEGIN_LIST_COL => {
                    if expected_columns.is_none() || ended_columns || in_column {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtBeginListCol occurs outside its column collection".to_string(),
                        ));
                    }
                    columns.push(Self::parse_table_column(record.payload(), columns.len())?);
                    in_column = true;
                    saw_xml_properties = false;
                },
                kind::BEGIN_LIST_XML_CPR => {
                    if !in_column || saw_xml_properties {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtBeginListXmlCPr occurs outside a column or is duplicated"
                                .to_string(),
                        ));
                    }
                    in_xml_properties = true;
                    saw_xml_properties = true;
                },
                kind::END_LIST_XML_CPR => {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "unmatched BrtEndListXmlCPr".to_string(),
                    ));
                },
                kind::END_LIST_COL => {
                    if !in_column || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unmatched or nonempty BrtEndListCol".to_string(),
                        ));
                    }
                    in_column = false;
                },
                kind::END_LIST_COLS => {
                    if expected_columns.is_none()
                        || in_column
                        || ended_columns
                        || !record.payload().is_empty()
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtEndListCols".to_string(),
                        ));
                    }
                    ended_columns = true;
                },
                kind::END_LIST => {
                    if !ended_columns || in_column || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtEndList".to_string(),
                        ));
                    }
                    ended_table = true;
                },
                _ => {},
            }
        }
        let (table_id, display_name, _) = table_header.ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "XLSB table part omits BrtBeginList".to_string(),
            )
        })?;
        let expected = expected_columns.ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "XLSB table part omits BrtBeginListCols".to_string(),
            )
        })?;
        if !ended_table || columns.len() != expected {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "XLSB table contains {} of {expected} declared columns or is unterminated",
                columns.len()
            )));
        }
        TableDefinition::try_new(table_id, sheet_index, display_name, columns)
    }

    fn parse_table_header(data: &[u8]) -> Result<(u32, String, usize)> {
        if data.len() < 64 {
            return Err(crate::package::error::Error::InvalidLength {
                expected: 64,
                found: data.len(),
            });
        }
        let row_first = binary::read_u32_le_at(data, 0)?;
        let row_last = binary::read_u32_le_at(data, 4)?;
        let col_first = binary::read_u32_le_at(data, 8)?;
        let col_last = binary::read_u32_le_at(data, 12)?;
        if row_first > row_last
            || row_last >= 1_048_576
            || col_first > col_last
            || col_last >= 16_384
        {
            return Err(crate::package::error::Error::InvalidFormula(
                "BrtBeginList contains an invalid table range".to_string(),
            ));
        }
        for offset in [24, 28] {
            if binary::read_u32_le_at(data, offset)? > 1 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "BrtBeginList contains a non-Boolean row flag".to_string(),
                ));
            }
        }
        let table_id = binary::read_u32_le_at(data, 20)?;
        let mut offset = 64;
        let mut strings = Vec::with_capacity(6);
        for _ in 0..6 {
            let (value, consumed) = Self::parse_nullable_wide_string(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
                    "BrtBeginList string size overflow".to_string(),
                )
            })?;
            strings.push(value);
        }
        if offset != data.len() {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtBeginList has {} trailing bytes",
                data.len() - offset
            )));
        }
        let display_name = strings[1].clone().ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "BrtBeginList has a NULL display name".to_string(),
            )
        })?;
        Ok((
            table_id,
            display_name,
            usize::try_from(col_last - col_first + 1).expect("bounded table width"),
        ))
    }

    fn parse_table_column(data: &[u8], index: usize) -> Result<String> {
        if data.len() < 24 || binary::read_u32_le_at(data, 0)? == 0 {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtBeginListCol {index} has an invalid header"
            )));
        }
        let mut offset = 24;
        let mut strings = Vec::with_capacity(6);
        for _ in 0..6 {
            let (value, consumed) = Self::parse_nullable_wide_string(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
                    "BrtBeginListCol string size overflow".to_string(),
                )
            })?;
            strings.push(value);
        }
        if offset != data.len() {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtBeginListCol has {} trailing bytes",
                data.len() - offset
            )));
        }
        strings[0]
            .clone()
            .or_else(|| strings[1].clone())
            .ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(format!(
                    "BrtBeginListCol {index} has neither a name nor caption"
                ))
            })
    }
}
