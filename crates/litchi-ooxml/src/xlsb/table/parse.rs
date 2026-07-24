//! Record-walking parser for the XLSB Table (ListObject) stream
//! (MS-XLSB 2.1.7.51).
//!
//! The parser is strict about record payloads it fully understands and
//! tolerant about everything else: unknown record types are ignored, and
//! known begin/end record pairs that carry no modelled data (XML column
//! properties, FRT wrappers, ...) are skipped as balanced collections.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::records::{XlsbRecord, XlsbRecordIter, record_types as rt, wide_str_with_len};
use crate::xlsb::table::model::*;
use litchi_core::binary;

// `BrtBeginList` flags word (MS-XLSB 2.4.100).
const LIST_SHOWN_TOTAL_ROW: u32 = 1 << 0;
const LIST_SINGLE_CELL: u32 = 1 << 1;
const LIST_INSERT_ROW_VISIBLE: u32 = 1 << 2;
const LIST_INSERT_ROW_INSERTED_CELLS: u32 = 1 << 3;
const LIST_PUBLISHED: u32 = 1 << 4;

/// `DXFId` value meaning no differential formatting (MS-XLSB 2.5.38).
const NO_DXF: u32 = u32::MAX;

/// `dwConnID` value meaning no external connection (MS-XLSB 2.4.100).
const NO_CONNECTION: u32 = 0;

// `BrtListCCFmla` / `BrtListTrFmla` flags byte (MS-XLSB 2.4.706, 2.4.708).
const FORMULA_ARRAY: u8 = 1 << 1;

// `BrtTableStyleClient` flags word (MS-XLSB 2.4.847).
const STYLE_FIRST_COLUMN: u16 = 1 << 0;
const STYLE_LAST_COLUMN: u16 = 1 << 1;
const STYLE_ROW_STRIPES: u16 = 1 << 2;
const STYLE_COLUMN_STRIPES: u16 = 1 << 3;

/// Byte length of an `FRTBlank` header (MS-XLSB 2.5.55).
const FRT_BLANK_LEN: usize = 4;

/// Parse a Table part (`tables/table*.bin`) into a typed [`XlsbTable`].
///
/// The stream must start with `BrtBeginList`. Records after `BrtEndList` are
/// ignored. Unknown record types anywhere in the stream are skipped without
/// failing.
pub fn parse_table_part(data: &[u8]) -> XlsbResult<XlsbTable> {
    let mut walker = RecordWalker::new(data);
    let first = walker.required("BrtBeginList")?;
    if first.header.record_type != rt::BEGIN_LIST {
        return Err(XlsbError::UnexpectedRecord {
            expected: rt::BEGIN_LIST,
            found: first.header.record_type,
        });
    }
    let mut table = parse_list_payload(&first.data)?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_LIST => return Ok(table),
            rt::BEGIN_LIST_COLS => parse_columns(&mut walker, &record.data, &mut table)?,
            rt::TABLE_STYLE_CLIENT => {
                table.style_info = Some(parse_style_client(&record.data)?);
            },
            rt::LIST14 => parse_list14(&record.data, &mut table)?,
            other => walker.skip_unhandled(other, "Table stream")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream("BrtEndList".to_string()))
}

/// Extract the Table part relationship identifiers a worksheet declares in
/// its `BrtBeginListParts` collection (MS-XLSB 2.4.103, 2.4.707).
///
/// Records outside the collection are ignored; the identifiers are returned
/// verbatim and are never dereferenced.
pub fn parse_table_part_rel_ids(data: &[u8]) -> XlsbResult<Vec<String>> {
    let mut walker = RecordWalker::new(data);
    let mut rel_ids = Vec::new();
    while let Some(record) = walker.next()? {
        if record.header.record_type == rt::BEGIN_LIST_PARTS {
            parse_list_parts(&mut walker, &record.data, &mut rel_ids)?;
        } else {
            walker.skip_unhandled(record.header.record_type, "Worksheet stream")?;
        }
    }
    Ok(rel_ids)
}

/// `BrtBeginListParts` collection (MS-XLSB 2.4.103).
fn parse_list_parts(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    rel_ids: &mut Vec<String>,
) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginListParts");
    let declared = cursor.read_u32()?;
    cursor.finish()?;
    let mut found = 0u32;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_LIST_PARTS => {
                if found != declared {
                    return Err(malformed(
                        "BrtBeginListParts",
                        format!("declared {declared} BrtListPart records, found {found}"),
                    ));
                }
                return Ok(());
            },
            rt::LIST_PART => {
                found = found
                    .checked_add(1)
                    .ok_or_else(|| malformed("BrtBeginListParts", "BrtListPart count overflow"))?;
                let mut cursor = PayloadCursor::new(&record.data, "BrtListPart");
                rel_ids.push(cursor.read_wide_string()?);
                cursor.finish()?;
            },
            other => walker.skip_unhandled(other, "BrtBeginListParts collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndListParts".to_string(),
    ))
}

/// Wraps the shared record iterator with the collection helpers this parser needs.
struct RecordWalker<'a> {
    iter: XlsbRecordIter<&'a [u8]>,
}

impl<'a> RecordWalker<'a> {
    fn new(data: &'a [u8]) -> Self {
        RecordWalker {
            iter: XlsbRecordIter::new(data),
        }
    }

    fn next(&mut self) -> XlsbResult<Option<XlsbRecord>> {
        self.iter.next().transpose()
    }

    fn required(&mut self, context: &'static str) -> XlsbResult<XlsbRecord> {
        self.next()?
            .ok_or_else(|| XlsbError::UnexpectedEndOfStream(context.to_string()))
    }

    /// Consume records up to and including `end_type`, tolerating nested
    /// collections of the same record pair.
    fn skip_collection(
        &mut self,
        begin_type: u16,
        end_type: u16,
        context: &'static str,
    ) -> XlsbResult<()> {
        let mut depth = 1u32;
        while let Some(record) = self.next()? {
            if record.header.record_type == begin_type {
                depth += 1;
            } else if record.header.record_type == end_type {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
            }
        }
        Err(XlsbError::UnexpectedEndOfStream(context.to_string()))
    }

    /// Skip a record the parser does not handle: a balanced collection when
    /// the type is a known begin record, a single record otherwise.
    fn skip_unhandled(&mut self, record_type: u16, context: &'static str) -> XlsbResult<()> {
        if let Some(end_type) = paired_end(record_type) {
            self.skip_collection(record_type, end_type, context)?;
        }
        Ok(())
    }
}

/// Map a known begin record type to its matching end record type.
///
/// Returns `None` for standalone records and unknown types, which the parser
/// then skips as single records.
fn paired_end(record_type: u16) -> Option<u16> {
    Some(match record_type {
        rt::BEGIN_LIST_COLS => rt::END_LIST_COLS,
        rt::BEGIN_LIST_COL => rt::END_LIST_COL,
        rt::BEGIN_LIST_XML_CPR => rt::END_LIST_XML_CPR,
        rt::BEGIN_LIST_PARTS => rt::END_LIST_PARTS,
        rt::FRT_BEGIN => rt::FRT_END,
        rt::AC_BEGIN => rt::AC_END,
        _ => return None,
    })
}

/// Bounds-checked cursor over one record payload.
struct PayloadCursor<'a> {
    data: &'a [u8],
    offset: usize,
    context: &'static str,
}

impl<'a> PayloadCursor<'a> {
    fn new(data: &'a [u8], context: &'static str) -> Self {
        PayloadCursor {
            data,
            offset: 0,
            context,
        }
    }

    fn remaining(&self) -> usize {
        self.data.len() - self.offset
    }

    fn guard(&self, needed: usize) -> XlsbResult<()> {
        if self.remaining() < needed {
            return Err(XlsbError::InvalidLength {
                expected: self.offset + needed,
                found: self.data.len(),
            });
        }
        Ok(())
    }

    fn read_u8(&mut self) -> XlsbResult<u8> {
        self.guard(1)?;
        let value = self.data[self.offset];
        self.offset += 1;
        Ok(value)
    }

    fn read_u16(&mut self) -> XlsbResult<u16> {
        self.guard(2)?;
        let value = binary::read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;
        Ok(value)
    }

    fn read_u32(&mut self) -> XlsbResult<u32> {
        self.guard(4)?;
        let value = binary::read_u32_le_at(self.data, self.offset)?;
        self.offset += 4;
        Ok(value)
    }

    /// Read a `Boolean` encoded as a 32-bit integer (MS-XLSB 2.5.98.3).
    fn read_bool32(&mut self) -> XlsbResult<u32> {
        let value = self.read_u32()?;
        if value > 1 {
            return Err(malformed(self.context, "non-Boolean 32-bit flag"));
        }
        Ok(value)
    }

    /// Read a `DXFId`, mapping `0xFFFFFFFF` to `None` (MS-XLSB 2.5.38).
    fn read_dxf_id(&mut self) -> XlsbResult<Option<u32>> {
        Ok(match self.read_u32()? {
            NO_DXF => None,
            id => Some(id),
        })
    }

    /// Read an `XLWideString` (MS-XLSB 2.5.169).
    fn read_wide_string(&mut self) -> XlsbResult<String> {
        let (value, consumed) = wide_str_with_len(&self.data[self.offset..])?;
        self.offset += consumed;
        Ok(value)
    }

    /// Read an `XLNullableWideString` (MS-XLSB 2.5.167).
    fn read_nullable_wide_string(&mut self) -> XlsbResult<Option<String>> {
        self.guard(4)?;
        if binary::read_u32_le_at(self.data, self.offset)? == u32::MAX {
            self.offset += 4;
            return Ok(None);
        }
        self.read_wide_string().map(Some)
    }

    /// Read a length-prefixed byte blob (`cce`/`cb` prefixed formula parts).
    fn read_blob(&mut self) -> XlsbResult<Vec<u8>> {
        let len = usize::try_from(self.read_u32()?)
            .map_err(|_| malformed(self.context, "byte blob length overflow"))?;
        self.guard(len)?;
        let blob = self.data[self.offset..self.offset + len].to_vec();
        self.offset += len;
        Ok(blob)
    }

    /// Reject payloads with unparsed trailing bytes.
    fn finish(&self) -> XlsbResult<()> {
        if self.remaining() != 0 {
            return Err(malformed(
                self.context,
                format!("{} trailing bytes", self.remaining()),
            ));
        }
        Ok(())
    }
}

fn malformed(context: &str, detail: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}

/// `BrtBeginList` payload (MS-XLSB 2.4.100).
fn parse_list_payload(data: &[u8]) -> XlsbResult<XlsbTable> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginList");
    let range = XlsbTableRange {
        first_row: cursor.read_u32()?,
        last_row: cursor.read_u32()?,
        first_column: cursor.read_u32()?,
        last_column: cursor.read_u32()?,
    };
    let table_type = XlsbTableType::try_from(cursor.read_u32()?)?;
    let id = cursor.read_u32()?;
    let header_row_count = cursor.read_bool32()?;
    let totals_row_count = cursor.read_bool32()?;
    let flags = cursor.read_u32()?;
    let header_dxf_id = cursor.read_dxf_id()?;
    let data_dxf_id = cursor.read_dxf_id()?;
    let totals_dxf_id = cursor.read_dxf_id()?;
    let border_dxf_id = cursor.read_dxf_id()?;
    let header_border_dxf_id = cursor.read_dxf_id()?;
    let totals_border_dxf_id = cursor.read_dxf_id()?;
    let connection_id = match cursor.read_u32()? {
        NO_CONNECTION => None,
        id => Some(id),
    };
    let name = cursor.read_nullable_wide_string()?;
    let display_name = cursor.read_nullable_wide_string()?;
    let comment = cursor.read_nullable_wide_string()?;
    let header_style = cursor.read_nullable_wide_string()?;
    let data_style = cursor.read_nullable_wide_string()?;
    let totals_style = cursor.read_nullable_wide_string()?;
    cursor.finish()?;
    Ok(XlsbTable {
        id,
        name,
        display_name,
        comment,
        range,
        table_type,
        header_row_count,
        totals_row_count,
        totals_row_shown: flags & LIST_SHOWN_TOTAL_ROW != 0,
        single_cell: flags & LIST_SINGLE_CELL != 0,
        insert_row_visible: flags & LIST_INSERT_ROW_VISIBLE != 0,
        insert_row_inserted_cells: flags & LIST_INSERT_ROW_INSERTED_CELLS != 0,
        published: flags & LIST_PUBLISHED != 0,
        header_dxf_id,
        data_dxf_id,
        totals_dxf_id,
        border_dxf_id,
        header_border_dxf_id,
        totals_border_dxf_id,
        connection_id,
        header_style,
        data_style,
        totals_style,
        ..XlsbTable::default()
    })
}

/// `BrtBeginListCols` collection (MS-XLSB 2.4.102).
fn parse_columns(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    table: &mut XlsbTable,
) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginListCols");
    let declared = cursor.read_u32()?;
    cursor.finish()?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_LIST_COLS => {
                let found = u32::try_from(table.columns.len())
                    .map_err(|_| malformed("BrtBeginListCols", "column count overflow"))?;
                if found != declared {
                    return Err(malformed(
                        "BrtBeginListCols",
                        format!("declared {declared} columns, found {found}"),
                    ));
                }
                return Ok(());
            },
            rt::BEGIN_LIST_COL => {
                table.columns.push(parse_column(walker, &record.data)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginListCols collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndListCols".to_string(),
    ))
}

/// `BrtBeginListCol` collection (MS-XLSB 2.4.101).
fn parse_column(walker: &mut RecordWalker<'_>, data: &[u8]) -> XlsbResult<XlsbTableColumn> {
    let mut column = parse_column_payload(data)?;
    while let Some(record) = walker.next()? {
        match record.header.record_type {
            rt::END_LIST_COL => return Ok(column),
            rt::LIST_CC_FMLA => {
                column.calculated_column_formula = Some(parse_formula(&record.data)?);
            },
            rt::LIST_TR_FMLA => {
                column.totals_row_formula = Some(parse_formula(&record.data)?);
            },
            other => walker.skip_unhandled(other, "BrtBeginListCol collection")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream(
        "BrtEndListCol".to_string(),
    ))
}

/// `BrtBeginListCol` payload (MS-XLSB 2.4.101).
fn parse_column_payload(data: &[u8]) -> XlsbResult<XlsbTableColumn> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginListCol");
    let id = cursor.read_u32()?;
    let totals_row_function = XlsbTableTotalsRowFunction::try_from(cursor.read_u32()?)?;
    let header_dxf_id = cursor.read_dxf_id()?;
    let insert_row_dxf_id = cursor.read_dxf_id()?;
    let totals_dxf_id = cursor.read_dxf_id()?;
    let query_table_field_id = cursor.read_u32()?;
    let name = cursor.read_nullable_wide_string()?;
    let caption = cursor.read_nullable_wide_string()?;
    let totals_row_label = cursor.read_nullable_wide_string()?;
    let header_style = cursor.read_nullable_wide_string()?;
    let insert_row_style = cursor.read_nullable_wide_string()?;
    let totals_style = cursor.read_nullable_wide_string()?;
    cursor.finish()?;
    Ok(XlsbTableColumn {
        id,
        totals_row_function,
        header_dxf_id,
        insert_row_dxf_id,
        totals_dxf_id,
        query_table_field_id,
        name,
        caption,
        totals_row_label,
        header_style,
        insert_row_style,
        totals_style,
        ..XlsbTableColumn::default()
    })
}

/// `BrtListCCFmla`/`BrtListTrFmla` payload: a flag byte followed by a
/// `ListParsedFormula` (MS-XLSB 2.4.706, 2.4.708, 2.5.98.11), stored verbatim.
fn parse_formula(data: &[u8]) -> XlsbResult<XlsbTableFormula> {
    let mut cursor = PayloadCursor::new(data, "BrtList formula");
    let flags = cursor.read_u8()?;
    let tokens = cursor.read_blob()?;
    let extra = cursor.read_blob()?;
    cursor.finish()?;
    Ok(XlsbTableFormula {
        array: flags & FORMULA_ARRAY != 0,
        tokens,
        extra,
    })
}

/// `BrtTableStyleClient` payload (MS-XLSB 2.4.847).
fn parse_style_client(data: &[u8]) -> XlsbResult<XlsbTableStyleInfo> {
    let mut cursor = PayloadCursor::new(data, "BrtTableStyleClient");
    let flags = cursor.read_u16()?;
    let name = cursor.read_nullable_wide_string()?;
    cursor.finish()?;
    Ok(XlsbTableStyleInfo {
        name,
        show_first_column: flags & STYLE_FIRST_COLUMN != 0,
        show_last_column: flags & STYLE_LAST_COLUMN != 0,
        show_row_stripes: flags & STYLE_ROW_STRIPES != 0,
        show_column_stripes: flags & STYLE_COLUMN_STRIPES != 0,
    })
}

/// `BrtList14` payload (MS-XLSB 2.4.705): table alternate text.
fn parse_list14(data: &[u8], table: &mut XlsbTable) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtList14");
    cursor.guard(FRT_BLANK_LEN)?;
    cursor.offset += FRT_BLANK_LEN;
    table.alternate_text = cursor.read_nullable_wide_string()?;
    table.alternate_text_summary = cursor.read_nullable_wide_string()?;
    cursor.finish()
}
