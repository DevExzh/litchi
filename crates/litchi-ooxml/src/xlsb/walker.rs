//! Shared record-walking and payload-cursor helpers for XLSB binary part
//! parsers (pivot caches, tables, chart sheets, connections, ...).
//!
//! Parsers built on these helpers are strict about record payloads they
//! fully understand and tolerant about everything else: unknown record
//! types are ignored, and known begin/end record pairs that carry no
//! modelled data are skipped as balanced collections.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::records::{XlsbRecord, XlsbRecordIter, record_types as rt, wide_str_with_len};

/// Maximum number of leading future-record wrapper blocks skipped while
/// looking for the record that opens a part stream.
const MAX_LEADING_WRAPPER_BLOCKS: usize = 16;

/// `BrtBeginPRule` / `BrtEndPRule` (MS-XLSB 2.4.186): unmodelled pivot rule.
const BEGIN_PRULE: u16 = 247;
const END_PRULE: u16 = 248;
/// `BrtBeginPCDKPIs` / `BrtEndPCDKPIs` (MS-XLSB 2.4.149): unmodelled KPI collection.
const BEGIN_PCD_KPIS: u16 = 269;
const END_PCD_KPIS: u16 = 270;
/// `BrtBeginPCDKPI` / `BrtEndPCDKPI` (MS-XLSB 2.4.148): unmodelled KPI.
const BEGIN_PCD_KPI: u16 = 271;
const END_PCD_KPI: u16 = 272;

pub(crate) fn malformed(context: &str, detail: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}

/// Wraps the shared record iterator with collection helpers.
pub(crate) struct RecordWalker<'a> {
    iter: XlsbRecordIter<&'a [u8]>,
}

impl<'a> RecordWalker<'a> {
    pub(crate) fn new(data: &'a [u8]) -> Self {
        RecordWalker {
            iter: XlsbRecordIter::new(data),
        }
    }

    pub(crate) fn next(&mut self) -> XlsbResult<Option<XlsbRecord>> {
        self.iter.next().transpose()
    }

    pub(crate) fn required(&mut self, context: &'static str) -> XlsbResult<XlsbRecord> {
        self.next()?
            .ok_or_else(|| XlsbError::UnexpectedEndOfStream(context.to_string()))
    }

    /// Read the record that opens a part stream, requiring it to be
    /// `begin_type`.
    ///
    /// Excel may prefix a part with future-record wrapper blocks
    /// (`BrtACBegin`/`BrtACEnd` and `BrtFRTBegin`/`BrtFRTEnd`) that carry
    /// application content a given reader version does not model. MS-XLSB
    /// requires readers to skip wrapper blocks they do not understand, so any
    /// leading balanced wrappers are consumed before the expected begin record
    /// is matched. At most [`MAX_LEADING_WRAPPER_BLOCKS`] are skipped so a
    /// crafted stream cannot loop unboundedly.
    pub(crate) fn required_begin(
        &mut self,
        begin_type: u16,
        context: &'static str,
    ) -> XlsbResult<XlsbRecord> {
        for _ in 0..MAX_LEADING_WRAPPER_BLOCKS {
            let record = self.required(context)?;
            let record_type = record.header.record_type;
            if record_type == begin_type {
                return Ok(record);
            }
            if !matches!(record_type, rt::AC_BEGIN | rt::FRT_BEGIN) {
                return Err(XlsbError::UnexpectedRecord {
                    expected: begin_type,
                    found: record_type,
                });
            }
            self.skip_unhandled(record_type, context)?;
        }
        Err(XlsbError::UnexpectedRecord {
            expected: begin_type,
            found: rt::AC_BEGIN,
        })
    }

    /// Consume records up to and including `end_type`, tolerating nested
    /// collections of the same record pair.
    pub(crate) fn skip_collection(
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
    pub(crate) fn skip_unhandled(
        &mut self,
        record_type: u16,
        context: &'static str,
    ) -> XlsbResult<()> {
        if let Some(end_type) = paired_end(record_type) {
            self.skip_collection(record_type, end_type, context)?;
        }
        Ok(())
    }

    /// Consume everything up to the matching end record of a collection that
    /// is expected to contain no modelled children.
    pub(crate) fn expect_end(&mut self, end_type: u16, context: &'static str) -> XlsbResult<()> {
        while let Some(record) = self.next()? {
            let record_type = record.header.record_type;
            if record_type == end_type {
                return Ok(());
            }
            self.skip_unhandled(record_type, context)?;
        }
        Err(XlsbError::UnexpectedEndOfStream(context.to_string()))
    }
}

/// Map a known begin record type to its matching end record type.
///
/// Returns `None` for standalone records and unknown types, which the parser
/// then skips as single records. This is the union of the record families
/// the XLSB part parsers understand; begin types are unique across families,
/// so the union cannot mis-pair a record.
fn paired_end(record_type: u16) -> Option<u16> {
    Some(match record_type {
        // PivotCache definition stream (2.1.7.38).
        rt::BEGIN_PCD_SOURCE => rt::END_PCD_SOURCE,
        rt::BEGIN_PCDS_RANGE => rt::END_PCDS_RANGE,
        rt::BEGIN_PCDS_CONSOL => rt::END_PCDS_CONSOL,
        rt::BEGIN_PCDSC_PAGES => rt::END_PCDSC_PAGES,
        rt::BEGIN_PCDSC_PAGE => rt::END_PCDSC_PAGE,
        rt::BEGIN_PCDSCP_ITEM => rt::END_PCDSCP_ITEM,
        rt::BEGIN_PCDSC_SETS => rt::END_PCDSC_SETS,
        rt::BEGIN_PCDSC_SET => rt::END_PCDSC_SET,
        rt::BEGIN_PCD_FIELDS => rt::END_PCD_FIELDS,
        rt::BEGIN_PCD_FIELD => rt::END_PCD_FIELD,
        rt::BEGIN_PCDF_ATBL => rt::END_PCDF_ATBL,
        rt::BEGIN_PCDI_RUN => rt::END_PCDI_RUN,
        rt::BEGIN_PCDF_GROUP => rt::END_PCDF_GROUP,
        rt::BEGIN_PCDFG_ITEMS => rt::END_PCDFG_ITEMS,
        rt::BEGIN_PCDFG_RANGE => rt::END_PCDFG_RANGE,
        rt::BEGIN_PCDFG_DISCRETE => rt::END_PCDFG_DISCRETE,
        rt::BEGIN_PCD_HIERARCHIES => rt::END_PCD_HIERARCHIES,
        rt::BEGIN_PCD_HIERARCHY => rt::END_PCD_HIERARCHY,
        rt::BEGIN_PCDH_FIELDS_USAGE => rt::END_PCDH_FIELDS_USAGE,
        rt::BEGIN_PCDHG_LEVELS => rt::END_PCDHG_LEVELS,
        rt::BEGIN_PCDHG_LEVEL => rt::END_PCDHG_LEVEL,
        rt::BEGIN_PCDHGL_GROUPS => rt::END_PCDHGL_GROUPS,
        rt::BEGIN_PCDHGL_GROUP => rt::END_PCDHGL_GROUP,
        rt::BEGIN_PCDHGLG_MEMBERS => rt::END_PCDHGLG_MEMBERS,
        rt::BEGIN_PCDHGLG_MEMBER => rt::END_PCDHGLG_MEMBER,
        rt::BEGIN_PCDSD_TUPLE_CACHE => rt::END_PCDSD_TUPLE_CACHE,
        rt::BEGIN_PCDSDTC_ENTRIES => rt::END_PCDSDTC_ENTRIES,
        rt::BEGIN_PCDSDTC_MEMBERS => rt::END_PCDSDTC_MEMBERS,
        rt::BEGIN_PCDSDTC_MEMBER => rt::END_PCDSDTC_MEMBER,
        rt::BEGIN_PCDSDTC_QUERIES => rt::END_PCDSDTC_QUERIES,
        rt::BEGIN_PCDSDTC_QUERY => rt::END_PCDSDTC_QUERY,
        rt::BEGIN_PCDSDTC_SETS => rt::END_PCDSDTC_SETS,
        rt::BEGIN_PCDSDTC_SET => rt::END_PCDSDTC_SET,
        rt::BEGIN_PCDSDTC_MEMBERS_SORT_BY => rt::END_PCDSDTC_MEMBERS_SORT_BY,
        rt::BEGIN_PCD_SFCI_ENTRIES => rt::END_PCD_SFCI_ENTRIES,
        rt::BEGIN_PCD_CALC_ITEMS => rt::END_PCD_CALC_ITEMS,
        rt::BEGIN_PCD_CALC_ITEM => rt::END_PCD_CALC_ITEM,
        rt::BEGIN_PCD_CALC_MEMS => rt::END_PCD_CALC_MEMS,
        rt::BEGIN_PCD_CALC_MEM => rt::END_PCD_CALC_MEM,
        rt::BEGIN_PCD_CALC_MEM14 => rt::END_PCD_CALC_MEM14,
        rt::BEGIN_PCD_CALC_MEM_EXT => rt::END_PCD_CALC_MEM_EXT,
        rt::BEGIN_PCD_CALC_MEMS_EXT => rt::END_PCD_CALC_MEMS_EXT,
        rt::BEGIN_PCD14 => rt::END_PCD14,
        rt::BEGIN_PR_FILTERS => rt::END_PR_FILTERS,
        rt::BEGIN_PR_FILTER => rt::END_PR_FILTER,
        rt::BEGIN_PRF_ITEM => rt::END_PRF_ITEM,
        rt::BEGIN_PR_FILTERS14 => rt::END_PR_FILTERS14,
        rt::BEGIN_PR_FILTER14 => rt::END_PR_FILTER14,
        rt::BEGIN_PRF_ITEM14 => rt::END_PRF_ITEM14,
        rt::BEGIN_P_NAMES => rt::END_P_NAMES,
        rt::BEGIN_P_NAME => rt::END_P_NAME,
        rt::BEGIN_PN_PAIRS => rt::END_PN_PAIRS,
        rt::BEGIN_PN_PAIR => rt::END_PN_PAIR,
        rt::BEGIN_ITEM_UNIQUE_NAMES => rt::END_ITEM_UNIQUE_NAMES,
        // Table stream (2.1.7.51).
        rt::BEGIN_LIST_COLS => rt::END_LIST_COLS,
        rt::BEGIN_LIST_COL => rt::END_LIST_COL,
        rt::BEGIN_LIST_XML_CPR => rt::END_LIST_XML_CPR,
        rt::BEGIN_LIST_PARTS => rt::END_LIST_PARTS,
        // Chart sheet stream (2.1.7.7).
        rt::BEGIN_CS_VIEWS => rt::END_CS_VIEWS,
        rt::BEGIN_CS_VIEW => rt::END_CS_VIEW,
        // Worksheet stream (2.1.7.62) sheet views.
        rt::BEGIN_WS_VIEWS => rt::END_WS_VIEWS,
        rt::BEGIN_WS_VIEW => rt::END_WS_VIEW,
        rt::BEGIN_HEADER_FOOTER => rt::END_HEADER_FOOTER,
        // External Data Connections part (2.1.7.24).
        rt::BEGIN_EC_DB_PROPS => rt::END_EC_DB_PROPS,
        rt::BEGIN_EC_OLAP_PROPS => rt::END_EC_OLAP_PROPS,
        rt::BEGIN_EC_WEB_PROPS => rt::END_EC_WEB_PROPS,
        rt::BEGIN_EC_WP_TABLES => rt::END_EC_WP_TABLES,
        rt::BEGIN_EC_PARAMS => rt::END_EC_PARAMS,
        rt::BEGIN_EC_PARAM => rt::END_EC_PARAM,
        rt::BEGIN_EC_TXT_WIZ => rt::END_EC_TXT_WIZ,
        rt::BEGIN_EXT_CONN14 => rt::END_EXT_CONN14,
        rt::BEGIN_EXT_CONN15 => rt::END_EXT_CONN15,
        rt::BEGIN_EC_TXT_WIZ15 => rt::END_EC_TXT_WIZ15,
        rt::BEGIN_DATA_FEED_PR15 => rt::END_DATA_FEED_PR15,
        rt::BEGIN_DB_TABLES15 => rt::END_DB_TABLES15,
        // Shared wrapper records.
        rt::FRT_BEGIN => rt::FRT_END,
        rt::AC_BEGIN => rt::AC_END,
        // Record families without constants in `records::record_types`.
        BEGIN_PRULE => END_PRULE,
        BEGIN_PCD_KPIS => END_PCD_KPIS,
        BEGIN_PCD_KPI => END_PCD_KPI,
        _ => return None,
    })
}

/// Bounds-checked cursor over one record payload.
pub(crate) struct PayloadCursor<'a> {
    data: &'a [u8],
    offset: usize,
    context: &'static str,
}

impl<'a> PayloadCursor<'a> {
    pub(crate) fn new(data: &'a [u8], context: &'static str) -> Self {
        PayloadCursor {
            data,
            offset: 0,
            context,
        }
    }

    pub(crate) fn remaining(&self) -> usize {
        self.data.len() - self.offset
    }

    pub(crate) fn guard(&self, needed: usize) -> XlsbResult<()> {
        if self.remaining() < needed {
            return Err(XlsbError::InvalidLength {
                expected: self.offset + needed,
                found: self.data.len(),
            });
        }
        Ok(())
    }

    /// Advance over `len` bytes without reading them.
    pub(crate) fn skip(&mut self, len: usize) -> XlsbResult<()> {
        self.guard(len)?;
        self.offset += len;
        Ok(())
    }

    pub(crate) fn read_bytes(&mut self, len: usize) -> XlsbResult<&'a [u8]> {
        self.guard(len)?;
        let bytes = &self.data[self.offset..self.offset + len];
        self.offset += len;
        Ok(bytes)
    }

    pub(crate) fn read_u8(&mut self) -> XlsbResult<u8> {
        self.guard(1)?;
        let value = self.data[self.offset];
        self.offset += 1;
        Ok(value)
    }

    pub(crate) fn read_u16(&mut self) -> XlsbResult<u16> {
        self.guard(2)?;
        let value = u16::from_le_bytes([self.data[self.offset], self.data[self.offset + 1]]);
        self.offset += 2;
        Ok(value)
    }

    pub(crate) fn read_i16(&mut self) -> XlsbResult<i16> {
        Ok(self.read_u16()? as i16)
    }

    pub(crate) fn read_u32(&mut self) -> XlsbResult<u32> {
        self.guard(4)?;
        let value = u32::from_le_bytes(
            self.data[self.offset..self.offset + 4]
                .try_into()
                .expect("slice length is guarded"),
        );
        self.offset += 4;
        Ok(value)
    }

    pub(crate) fn read_i32(&mut self) -> XlsbResult<i32> {
        Ok(self.read_u32()? as i32)
    }

    pub(crate) fn read_f64(&mut self) -> XlsbResult<f64> {
        self.guard(8)?;
        let value = f64::from_le_bytes(
            self.data[self.offset..self.offset + 8]
                .try_into()
                .expect("slice length is guarded"),
        );
        self.offset += 8;
        Ok(value)
    }

    /// Read a full-width byte Boolean (any nonzero value is `true`).
    pub(crate) fn read_bool8(&mut self) -> XlsbResult<bool> {
        Ok(self.read_u8()? != 0)
    }

    /// Read a `Boolean` encoded as a 32-bit integer (MS-XLSB 2.5.98.3).
    pub(crate) fn read_bool32(&mut self) -> XlsbResult<u32> {
        let value = self.read_u32()?;
        if value > 1 {
            return Err(malformed(self.context, "non-Boolean 32-bit flag"));
        }
        Ok(value)
    }

    /// Read an `XLWideString` (MS-XLSB 2.5.169).
    pub(crate) fn read_wide_string(&mut self) -> XlsbResult<String> {
        let (value, consumed) = wide_str_with_len(&self.data[self.offset..])?;
        self.offset += consumed;
        Ok(value)
    }

    /// Read an `XLNullableWideString` (MS-XLSB 2.5.167).
    pub(crate) fn read_nullable_wide_string(&mut self) -> XlsbResult<Option<String>> {
        self.guard(4)?;
        if u32::from_le_bytes(
            self.data[self.offset..self.offset + 4]
                .try_into()
                .expect("slice length is guarded"),
        ) == u32::MAX
        {
            self.offset += 4;
            return Ok(None);
        }
        self.read_wide_string().map(Some)
    }

    /// Read a length-prefixed byte blob (`LPByteBuf`, MS-XLSB 2.5.91).
    pub(crate) fn read_blob(&mut self) -> XlsbResult<Vec<u8>> {
        let len = usize::try_from(self.read_u32()?)
            .map_err(|_| malformed(self.context, "byte blob length overflow"))?;
        self.guard(len)?;
        let blob = self.data[self.offset..self.offset + len].to_vec();
        self.offset += len;
        Ok(blob)
    }

    /// Reject payloads with unparsed trailing bytes.
    pub(crate) fn finish(&self) -> XlsbResult<()> {
        if self.remaining() != 0 {
            return Err(malformed(
                self.context,
                format!("{} trailing bytes", self.remaining()),
            ));
        }
        Ok(())
    }
}
