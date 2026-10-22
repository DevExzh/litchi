#![allow(
    clippy::cast_possible_wrap,
    clippy::expect_used,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error to this codec boundary"
)]

//! Typed workbook, external-link, PivotTable, and table record codecs.

use std::collections::HashSet;

use super::super::model::Workbook;
use crate::calc::{self, Props};
use crate::leniency::{Tolerance, ToleranceReport};
use crate::named_ranges::Definition;
use crate::package::error::Result;
use crate::package::formula::{
    ExternalSheet, SupportingLink, View, table::Definition as TableDefinition,
};
use crate::raw::{Records, kind};
use litchi_core::binary;

fn worksheet_name_key(name: &str) -> Result<String> {
    let key_len =
        name.chars()
            .flat_map(char::to_lowercase)
            .try_fold(0usize, |length, character| {
                length.checked_add(character.len_utf8()).ok_or(
                    crate::package::error::Error::CapacityOverflow {
                        resource: "XLSB worksheet-name equivalence key",
                    },
                )
            })?;
    let mut key = String::new();
    key.try_reserve_exact(key_len)
        .map_err(|source| crate::package::error::Error::Allocation {
            resource: "XLSB worksheet-name equivalence key",
            source,
        })?;
    key.extend(name.chars().flat_map(char::to_lowercase));
    Ok(key)
}

/// A record's byte offset in its part, saturated for the tolerance report.
fn record_location(offset: usize) -> u32 {
    u32::try_from(offset).unwrap_or(u32::MAX)
}

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
    /// Parse the workbook stream, recording every Excel-compatible repair
    /// into `report` (see [`crate::leniency`]).
    pub(in crate::workbook) fn read_workbook(
        iter: &mut Records<'_>,
        report: &mut ToleranceReport,
        raw_limits: crate::raw::Limits,
    ) -> Result<ParsedWorkbookInfo> {
        let mut info = ParsedWorkbookInfo::default();
        let worksheet_names = &mut info.worksheet_names;
        let worksheet_rel_ids = &mut info.worksheet_rel_ids;
        let worksheet_states = &mut info.worksheet_states;
        let supporting_links = &mut info.supporting_links;
        let external_sheets = &mut info.external_sheets;
        let external_link_rel_ids = &mut info.external_link_rel_ids;
        let defined_names = &mut info.defined_names;
        let mut worksheet_name_keys = HashSet::new();
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
                            typ: "BrtBeginBookViews",
                            val: "duplicate or nested book-view block",
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
                            typ: "BrtEndBookViews",
                            val: "book-view block has no matching begin record",
                        });
                    }
                    book_views_open = false;
                },
                kind::BOOK_VIEW => {
                    if !book_views_open {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBookView",
                            val: "record is outside the book-view block",
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
                                "BrtBookView itabFirst does not fit usize",
                            )
                        })?;
                    let active =
                        usize::try_from(binary::read_u32_le_at(data, 24)?).map_err(|_| {
                            crate::package::error::Error::InvalidFormula(
                                "BrtBookView itabCur does not fit usize",
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
                    if let Ok(prop) = crate::package::records::WorkbookPropRecord::parse(
                        record.payload(),
                        raw_limits,
                    ) {
                        *is_1904 = prop.is_date1904;
                    }
                },
                kind::CALC_PROP => {
                    if info.calc.is_some() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtCalcProp",
                            val: "duplicate record",
                        });
                    }
                    info.calc = Some(calc::read_with(
                        record.payload(),
                        raw_limits,
                        Tolerance::recording(report),
                    )?);
                },
                kind::BUNDLE_SH => {
                    let bundle_sh = crate::package::records::BundleSheetRecord::parse_with(
                        record.payload(),
                        u32::try_from(worksheet_names.len()).unwrap_or(u32::MAX),
                        Tolerance::recording(report),
                        raw_limits,
                    )?;
                    let name_key = worksheet_name_key(&bundle_sh.name)?;
                    if worksheet_name_keys.contains(&name_key) {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBundleSh strName",
                            val: "duplicate sheet name",
                        });
                    }
                    worksheet_name_keys.try_reserve(1).map_err(|source| {
                        crate::package::error::Error::Allocation {
                            resource: "XLSB worksheet-name equivalence index",
                            source,
                        }
                    })?;
                    worksheet_name_keys.insert(name_key);
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
                        crate::package::records::decode_string(record.payload(), raw_limits)?;
                    if rel_id.is_empty() || consumed != record.payload().len() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtSupBookSrc has an invalid relationship ID",
                        ));
                    }
                    external_link_rel_ids.try_reserve(1).map_err(|source| {
                        crate::package::error::Error::Allocation {
                            resource: "eager XLSB external-link relationship IDs",
                            source,
                        }
                    })?;
                    supporting_links.try_reserve(1).map_err(|source| {
                        crate::package::error::Error::Allocation {
                            resource: "eager XLSB supporting-link results",
                            source,
                        }
                    })?;
                    let book_index = u32::try_from(external_link_rel_ids.len()).map_err(|_| {
                        crate::package::error::Error::InvalidFormula("external-link count overflow")
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
                    let named_range = Definition::parse_with(
                        record.payload(),
                        u32::try_from(defined_names.len()).unwrap_or(u32::MAX),
                        Tolerance::recording(report),
                        raw_limits,
                    )?;
                    if named_range
                        .sheet_id
                        .is_some_and(|index| index as usize >= worksheet_names.len())
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtName has invalid sheet scope",
                        ));
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
                typ: "BrtBeginBookViews",
                val: "missing BrtEndBookViews record",
            });
        }
        if let Some(maximum) = max_book_view_catalog_position
            && maximum >= worksheet_names.len()
        {
            return Err(crate::package::error::Error::InvalidFormula(
                "BrtBookView sheet position exceeds the available sheets",
            ));
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
            crate::package::error::Error::InvalidFormula("BrtExternSheet count overflow")
        })?;
        if count >= 65_536 {
            return Err(crate::package::error::Error::InvalidFormula(
                "BrtExternSheet count exceeds 65,535",
            ));
        }
        let expected = 4usize
            .checked_add(count.checked_mul(12).ok_or(
                crate::package::error::Error::InvalidFormula("BrtExternSheet size overflow"),
            )?)
            .ok_or(crate::package::error::Error::InvalidFormula(
                "BrtExternSheet size overflow",
            ))?;
        if data.len() != expected {
            return Err(crate::package::error::Error::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        external_sheets.reserve(count);
        for chunk in data[4..].as_chunks::<12>().0.iter() {
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
        raw_limits: crate::raw::Limits,
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
            let (value, consumed) = crate::package::records::decode_string(data, raw_limits)?;
            Ok((Some(value), consumed))
        }
    }

    pub(in crate::workbook) fn parse_pivot_cache_ids(
        data: &[u8],
        raw_limits: crate::raw::Limits,
    ) -> Result<Vec<(u32, String)>> {
        let mut in_collection = false;
        let mut open_cache = false;
        let mut ended = false;
        let mut caches = Vec::new();
        for record in Records::new(data, raw_limits) {
            let record = record?;
            match record.kind() {
                kind::BEGIN_PIVOT_CACHE_IDS => {
                    if in_collection || ended {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "duplicate BrtBeginPivotCacheIDs collection",
                        ));
                    }
                    in_collection = true;
                },
                kind::BEGIN_PIVOT_CACHE_ID => {
                    if !in_collection || open_cache || record.payload().len() < 8 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "malformed BrtBeginPivotCacheID nesting or payload",
                        ));
                    }
                    let cache_id = binary::read_u32_le_at(record.payload(), 0)?;
                    let (rel_id, consumed) =
                        crate::package::records::decode_string(&record.payload()[4..], raw_limits)?;
                    if 4 + consumed != record.payload().len()
                        || rel_id.is_empty()
                        || rel_id.encode_utf16().count() > 255
                        || caches
                            .iter()
                            .any(|(existing, _): &(u32, String)| *existing == cache_id)
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid or duplicate PivotCache ID",
                        ));
                    }
                    caches.push((cache_id, rel_id));
                    open_cache = true;
                },
                kind::END_PIVOT_CACHE_ID => {
                    if !open_cache || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unbalanced BrtEndPivotCacheID",
                        ));
                    }
                    open_cache = false;
                },
                kind::END_PIVOT_CACHE_IDS => {
                    if !in_collection || open_cache || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unbalanced BrtEndPivotCacheIDs",
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
                "unterminated PivotCache ID collection",
            ));
        }
        Ok(caches)
    }

    pub(in crate::workbook) fn parse_pivot_view(
        data: &[u8],
        sheet_index: usize,
        raw_limits: crate::raw::Limits,
    ) -> Result<View> {
        let mut view = None;
        for record in Records::new(data, raw_limits) {
            let record = record?;
            if record.kind() != kind::BEGIN_SX_VIEW {
                continue;
            }
            if view.is_some() || record.payload().len() < 36 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "PivotTable part has duplicate or truncated BrtBeginSXView",
                ));
            }
            let cache_id = binary::read_u32_le_at(record.payload(), 28)?;
            let (name, consumed) =
                crate::package::records::decode_string(&record.payload()[32..], raw_limits)?;
            if consumed > record.payload().len() - 32 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "PivotTable view name overruns BrtBeginSXView",
                ));
            }
            view = Some(View::try_new(cache_id, sheet_index, name)?);
        }
        view.ok_or(crate::package::error::Error::InvalidFormula(
            "PivotTable part omits BrtBeginSXView",
        ))
    }

    /// Parse the formula-context view of one table part, ignoring and
    /// recording trailing payload bytes through `tolerance`.
    pub(in crate::workbook) fn parse_table_definition(
        data: &[u8],
        sheet_index: usize,
        tolerance: &mut Tolerance<'_>,
        raw_limits: crate::raw::Limits,
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
        let mut iter = Records::new(data, raw_limits);
        for record in iter.by_ref() {
            let record = record?;
            if ended_table {
                return Err(crate::package::error::Error::InvalidFormula(
                    "XLSB table part contains records after BrtEndList",
                ));
            }
            if in_xml_properties {
                match record.kind() {
                    kind::FRT_BEGIN | kind::AC_BEGIN => {
                        if xml_extension_depth == MAX_XML_EXTENSION_DEPTH {
                            return Err(crate::package::error::Error::InvalidFormula(
                                "BrtBeginListXmlCPr extension nesting exceeds 64",
                            ));
                        }
                        xml_extension_stack[xml_extension_depth] = record.kind();
                        xml_extension_depth += 1;
                    },
                    kind::FRT_END | kind::AC_END => {
                        let Some(depth) = xml_extension_depth.checked_sub(1) else {
                            return Err(crate::package::error::Error::InvalidFormula(
                                "unmatched extension end in BrtBeginListXmlCPr",
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
                                "mismatched extension wrapper in BrtBeginListXmlCPr",
                            ));
                        }
                        xml_extension_depth = depth;
                    },
                    kind::END_LIST_XML_CPR if xml_extension_depth == 0 => {
                        if !record.payload().is_empty() {
                            return Err(crate::package::error::Error::InvalidFormula(
                                "nonempty BrtEndListXmlCPr",
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
                            "table structure occurs inside BrtBeginListXmlCPr",
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
                            "XLSB table part contains duplicate BrtBeginList",
                        ));
                    }
                    table_header = Some(Self::parse_table_header(
                        record.payload(),
                        record_location(record.offset()),
                        tolerance,
                        raw_limits,
                    )?);
                },
                kind::BEGIN_LIST_COLS => {
                    let (_, _, range_columns) = table_header.as_ref().ok_or({
                        crate::package::error::Error::InvalidFormula(
                            "BrtBeginListCols precedes BrtBeginList",
                        )
                    })?;
                    if expected_columns.is_some() || record.payload().len() != 4 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid or duplicate BrtBeginListCols",
                        ));
                    }
                    let count = usize::try_from(binary::read_u32_le_at(record.payload(), 0)?)
                        .map_err(|_| {
                            crate::package::error::Error::InvalidFormula(
                                "table column count overflow",
                            )
                        })?;
                    if count == 0 || count > 16_384 || count != *range_columns {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "table column count disagrees with range width",
                        ));
                    }
                    expected_columns = Some(count);
                },
                kind::BEGIN_LIST_COL => {
                    if expected_columns.is_none() || ended_columns || in_column {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtBeginListCol occurs outside its column collection",
                        ));
                    }
                    columns.push(Self::parse_table_column(
                        record.payload(),
                        columns.len(),
                        record_location(record.offset()),
                        tolerance,
                        raw_limits,
                    )?);
                    in_column = true;
                    saw_xml_properties = false;
                },
                kind::BEGIN_LIST_XML_CPR => {
                    if !in_column || saw_xml_properties {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtBeginListXmlCPr occurs outside a column or is duplicated",
                        ));
                    }
                    in_xml_properties = true;
                    saw_xml_properties = true;
                },
                kind::END_LIST_XML_CPR => {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "unmatched BrtEndListXmlCPr",
                    ));
                },
                kind::END_LIST_COL => {
                    if !in_column || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unmatched or nonempty BrtEndListCol",
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
                            "invalid BrtEndListCols",
                        ));
                    }
                    ended_columns = true;
                },
                kind::END_LIST => {
                    if !ended_columns || in_column || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtEndList",
                        ));
                    }
                    ended_table = true;
                },
                _ => {},
            }
        }
        let (table_id, display_name, _) = table_header.ok_or(
            crate::package::error::Error::InvalidFormula("XLSB table part omits BrtBeginList"),
        )?;
        let expected = expected_columns.ok_or(crate::package::error::Error::InvalidFormula(
            "XLSB table part omits BrtBeginListCols",
        ))?;
        if !ended_table || columns.len() != expected {
            return Err(crate::package::error::Error::InvalidFormula(
                "XLSB table contains of declared columns or is unterminated",
            ));
        }
        TableDefinition::try_new(table_id, sheet_index, display_name, columns)
    }

    fn parse_table_header(
        data: &[u8],
        location: u32,
        tolerance: &mut Tolerance<'_>,
        raw_limits: crate::raw::Limits,
    ) -> Result<(u32, String, usize)> {
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
                "BrtBeginList contains an invalid table range",
            ));
        }
        for offset in [24, 28] {
            if binary::read_u32_le_at(data, offset)? > 1 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "BrtBeginList contains a non-Boolean row flag",
                ));
            }
        }
        let table_id = binary::read_u32_le_at(data, 20)?;
        let mut offset = 64;
        let mut strings = Vec::with_capacity(6);
        for _ in 0..6 {
            let (value, consumed) = Self::parse_nullable_wide_string(&data[offset..], raw_limits)?;
            offset = offset.checked_add(consumed).ok_or(
                crate::package::error::Error::InvalidFormula("BrtBeginList string size overflow"),
            )?;
            strings.push(value);
        }
        let trailing = data.len().saturating_sub(offset);
        tolerance.tolerate_trailing(kind::BEGIN_LIST, location, trailing, || {
            crate::package::error::Error::InvalidFormula("BrtBeginList has trailing bytes")
        })?;
        let display_name =
            strings[1]
                .clone()
                .ok_or(crate::package::error::Error::InvalidFormula(
                    "BrtBeginList has a NULL display name",
                ))?;
        Ok((
            table_id,
            display_name,
            usize::try_from(col_last - col_first + 1).expect("bounded table width"),
        ))
    }

    fn parse_table_column(
        data: &[u8],
        _index: usize,
        location: u32,
        tolerance: &mut Tolerance<'_>,
        raw_limits: crate::raw::Limits,
    ) -> Result<String> {
        if data.len() < 24 || binary::read_u32_le_at(data, 0)? == 0 {
            return Err(crate::package::error::Error::InvalidFormula(
                "BrtBeginListCol has an invalid header",
            ));
        }
        let mut offset = 24;
        let mut strings = Vec::with_capacity(6);
        for _ in 0..6 {
            let (value, consumed) = Self::parse_nullable_wide_string(&data[offset..], raw_limits)?;
            offset = offset.checked_add(consumed).ok_or(
                crate::package::error::Error::InvalidFormula(
                    "BrtBeginListCol string size overflow",
                ),
            )?;
            strings.push(value);
        }
        let trailing = data.len().saturating_sub(offset);
        tolerance.tolerate_trailing(kind::BEGIN_LIST_COL, location, trailing, || {
            crate::package::error::Error::InvalidFormula("BrtBeginListCol has trailing bytes")
        })?;
        strings[0].clone().or_else(|| strings[1].clone()).ok_or({
            crate::package::error::Error::InvalidFormula(
                "BrtBeginListCol has neither a name nor caption",
            )
        })
    }
}
