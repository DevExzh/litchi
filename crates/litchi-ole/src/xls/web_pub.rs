//! BIFF8 `WebPub` record (MS-XLS 2.4.344): a single published Web page.
//!
//! `WebPub` records appear in the workbook globals substream and in
//! worksheet substreams (`*WebPub` in both `WORKBOOKCONTENT` and
//! `WORKSHEETCONTENT`, MS-XLS 2.1). Each record describes one published
//! item: what was published (`tws`), the kind of page produced (`twd`),
//! where it was published to, and the destination bookmark and title.
//!
//! Everything in this module is INERT: URLs and file paths are stored
//! verbatim and are never opened, resolved, or fetched.

use super::{XlsError, XlsResult};

/// Record type of the `WebPub` record (MS-XLS 2.4.344).
pub(crate) const WEB_PUB_RECORD_TYPE: u16 = 0x0801;

/// Size in bytes of the fixed record part: `FrtRefHeaderU`, `tws`, `twd`,
/// the flag word, `reserved3`, `unused2`, `nStyleId`, and `cb`.
const FIXED_LEN: usize = 28;
/// Size in bytes of the trailing `unused3` field.
const TRAILING_UNUSED_LEN: usize = 2;
/// Maximum character count of a `WebPubString` (MS-XLS 2.5.278).
const MAX_WEB_PUB_STRING_CHARS: usize = 255;

// `FrtRefHeaderU.grbitFrt` bits (MS-XLS 2.5.135 FrtFlags).
const FRT_REF: u16 = 0x0001;

// Flag word bits following `tws`/`twd`.
const AUTO_REPUBLISH: u16 = 0x0002;
const MHTML: u16 = 0x0008;

/// `fHighByte` bit of a BIFF8 string option byte.
const HIGH_BYTE: u8 = 0x01;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: WEB_PUB_RECORD_TYPE,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

/// The kind of Web source that was published (`WebPub.tws`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsWebSourceType {
    /// The source is undefined.
    Undefined,
    /// The whole workbook.
    Workbook,
    /// An entire sheet.
    Sheet,
    /// A print area.
    PrintArea,
    /// An AutoFilter range.
    AutoFilter,
    /// A range of cells; the record's `FrtRefHeaderU.ref8` applies.
    Range,
    /// A chart; the record carries the chart's shape identifier.
    Chart,
    /// A PivotTable report.
    PivotTable,
    /// A query table (external data range).
    QueryTable,
    /// A named range.
    NamedRange,
}

impl XlsWebSourceType {
    fn from_code(code: u8) -> XlsResult<Self> {
        Ok(match code {
            0xFF => Self::Undefined,
            0x00 => Self::Workbook,
            0x01 => Self::Sheet,
            0x02 => Self::PrintArea,
            0x03 => Self::AutoFilter,
            0x04 => Self::Range,
            0x05 => Self::Chart,
            0x06 => Self::PivotTable,
            0x07 => Self::QueryTable,
            0x08 => Self::NamedRange,
            other => return Err(invalid(format!("unknown WebPub tws value 0x{other:02X}"))),
        })
    }

    /// Raw `tws` code; governs the conditional `srcName`/`crtID` fields.
    fn code(self) -> u8 {
        match self {
            Self::Undefined => 0xFF,
            Self::Workbook => 0x00,
            Self::Sheet => 0x01,
            Self::PrintArea => 0x02,
            Self::AutoFilter => 0x03,
            Self::Range => 0x04,
            Self::Chart => 0x05,
            Self::PivotTable => 0x06,
            Self::QueryTable => 0x07,
            Self::NamedRange => 0x08,
        }
    }
}

/// The kind of Web page created for a published item (`WebPub.twd`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsWebPageType {
    /// A non-interactive page, only for viewing.
    ViewOnly,
    /// An interactive page using workbook functionality.
    WorkbookFunctionality,
    /// An interactive page using PivotTable functionality.
    PivotTableFunctionality,
    /// An interactive page using chart functionality.
    ChartFunctionality,
}

impl XlsWebPageType {
    fn from_code(code: u8) -> XlsResult<Self> {
        Ok(match code {
            0x00 => Self::ViewOnly,
            0x01 => Self::WorkbookFunctionality,
            0x02 => Self::PivotTableFunctionality,
            0x03 => Self::ChartFunctionality,
            other => return Err(invalid(format!("unknown WebPub twd value 0x{other:02X}"))),
        })
    }

    /// Raw `twd` code.
    fn code(self) -> u8 {
        match self {
            Self::ViewOnly => 0x00,
            Self::WorkbookFunctionality => 0x01,
            Self::PivotTableFunctionality => 0x02,
            Self::ChartFunctionality => 0x03,
        }
    }
}

/// The cell range a `WebPub` record publishes (`FrtRefHeaderU.ref8`),
/// present only when the source type is [`XlsWebSourceType::Range`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsWebPubRange {
    /// First row of the range.
    pub first_row: u16,
    /// Last row of the range.
    pub last_row: u16,
    /// First column of the range.
    pub first_column: u16,
    /// Last column of the range.
    pub last_column: u16,
}

/// Typed `WebPub` record content (MS-XLS 2.4.344).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsWebPub {
    /// What was published (`tws`).
    pub source: XlsWebSourceType,
    /// The kind of Web page created (`twd`).
    pub page_type: XlsWebPageType,
    /// The published cell range, present iff `source` is
    /// [`XlsWebSourceType::Range`].
    pub range: Option<XlsWebPubRange>,
    /// Whether the page is republished when the workbook is saved
    /// (`fAutoRepublish`).
    pub auto_republish: bool,
    /// Whether the page is published as a single Web page (MHTML) rather
    /// than a page with references to other files (`fMhtml`).
    pub single_file: bool,
    /// Unique identifier of the published content (`nStyleId`).
    pub style_id: u32,
    /// The named range to publish (`srcName`), present iff the `tws` code
    /// is greater than 4.
    pub source_name: Option<String>,
    /// URL or path of the published page (`stFileDest`).
    pub file_destination: String,
    /// Destination bookmark of the published page (`stDivId`).
    pub div_id: String,
    /// Title of the published item (`stTitle`).
    pub title: String,
    /// Shape identifier of the published chart object (`crtID`), present
    /// iff `source` is [`XlsWebSourceType::Chart`].
    pub chart_shape_id: Option<u32>,
    /// Bytes reserved for future use (`frtRgb`), preserved verbatim.
    pub reserved: Vec<u8>,
}

impl XlsWebPub {
    /// Parse a `WebPub` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < FIXED_LEN + TRAILING_UNUSED_LEN {
            return Err(XlsError::InvalidLength {
                expected: FIXED_LEN + TRAILING_UNUSED_LEN,
                found: data.len(),
            });
        }
        if read_u16(data, 0) != WEB_PUB_RECORD_TYPE {
            return Err(invalid("WebPub FrtRefHeaderU.rt mismatch"));
        }
        let has_ref = read_u16(data, 2) & FRT_REF != 0;
        let range_ref = XlsWebPubRange {
            first_row: read_u16(data, 4),
            last_row: read_u16(data, 6),
            first_column: read_u16(data, 8),
            last_column: read_u16(data, 10),
        };

        let source = XlsWebSourceType::from_code(data[12])?;
        let page_type = XlsWebPageType::from_code(data[13])?;
        let flags = read_u16(data, 14);
        let style_id = read_u32(data, 20);
        let declared_size = read_u32(data, 24);

        // cb counts everything after the fixed part (MS-XLS 2.4.344).
        let tail_size = data.len() - FIXED_LEN;
        if usize::try_from(declared_size) != Ok(tail_size) {
            return Err(invalid("WebPub cb does not match the record size"));
        }
        // Per MS-XLS 2.4.344 the ref8 range applies iff tws is 0x04, and
        // fFrtRef MUST be set exactly in that case.
        if (source == XlsWebSourceType::Range) != has_ref {
            return Err(invalid("WebPub fFrtRef does not match the tws source type"));
        }

        let mut offset = FIXED_LEN;
        let source_name = if source.code() > XlsWebSourceType::Range.code() {
            let (name, used) = parse_web_pub_string(&data[offset..])?;
            offset += used;
            Some(name)
        } else {
            None
        };
        let (file_destination, used) = parse_web_pub_string(&data[offset..])?;
        offset += used;
        let (div_id, used) = parse_web_pub_string(&data[offset..])?;
        offset += used;
        let (title, used) = parse_web_pub_string(&data[offset..])?;
        offset += used;
        let chart_shape_id = if source == XlsWebSourceType::Chart {
            let raw = data
                .get(offset..offset + 4)
                .ok_or(XlsError::InvalidLength {
                    expected: offset + 4,
                    found: data.len(),
                })?;
            offset += 4;
            Some(read_u32(raw, 0))
        } else {
            None
        };

        // frtRgb fills the record up to the trailing 2-byte unused3 field.
        let reserved_end = data.len() - TRAILING_UNUSED_LEN;
        if offset > reserved_end {
            return Err(invalid("WebPub strings overrun the record"));
        }
        let reserved = data[offset..reserved_end].to_vec();

        Ok(XlsWebPub {
            source,
            page_type,
            range: (source == XlsWebSourceType::Range).then_some(range_ref),
            auto_republish: flags & AUTO_REPUBLISH != 0,
            single_file: flags & MHTML != 0,
            style_id,
            source_name,
            file_destination,
            div_id,
            title,
            chart_shape_id,
            reserved,
        })
    }

    /// Serialize back to a complete `WebPub` record payload.
    ///
    /// The conditional fields must agree with [`XlsWebPub::source`]:
    /// `range` is required exactly for [`XlsWebSourceType::Range`],
    /// `source_name` is required exactly when the source code is greater
    /// than 4, and `chart_shape_id` exactly for [`XlsWebSourceType::Chart`].
    pub(crate) fn to_payload(&self) -> XlsResult<Vec<u8>> {
        let wants_range = self.source == XlsWebSourceType::Range;
        if wants_range != self.range.is_some() {
            return Err(XlsError::InvalidData(
                "WebPub range must be present iff the source type is Range".to_string(),
            ));
        }
        let wants_name = self.source.code() > XlsWebSourceType::Range.code();
        if wants_name != self.source_name.is_some() {
            return Err(XlsError::InvalidData(
                "WebPub source_name must be present iff the tws code is greater than 4".to_string(),
            ));
        }
        let wants_shape = self.source == XlsWebSourceType::Chart;
        if wants_shape != self.chart_shape_id.is_some() {
            return Err(XlsError::InvalidData(
                "WebPub chart_shape_id must be present iff the source type is Chart".to_string(),
            ));
        }

        let mut tail = Vec::new();
        if let Some(name) = &self.source_name {
            write_web_pub_string(&mut tail, name)?;
        }
        write_web_pub_string(&mut tail, &self.file_destination)?;
        write_web_pub_string(&mut tail, &self.div_id)?;
        write_web_pub_string(&mut tail, &self.title)?;
        if let Some(shape_id) = self.chart_shape_id {
            tail.extend_from_slice(&shape_id.to_le_bytes());
        }
        tail.extend_from_slice(&self.reserved);
        tail.extend_from_slice(&[0u8; TRAILING_UNUSED_LEN]); // unused3

        let mut payload = Vec::with_capacity(FIXED_LEN + tail.len());
        payload.extend_from_slice(&WEB_PUB_RECORD_TYPE.to_le_bytes());
        let mut grbit_frt = 0u16;
        if wants_range {
            grbit_frt |= FRT_REF;
        }
        payload.extend_from_slice(&grbit_frt.to_le_bytes());
        let range = self.range.unwrap_or(XlsWebPubRange {
            first_row: 0,
            last_row: 0,
            first_column: 0,
            last_column: 0,
        });
        payload.extend_from_slice(&range.first_row.to_le_bytes());
        payload.extend_from_slice(&range.last_row.to_le_bytes());
        payload.extend_from_slice(&range.first_column.to_le_bytes());
        payload.extend_from_slice(&range.last_column.to_le_bytes());
        payload.push(self.source.code());
        payload.push(self.page_type.code());
        let mut flags = 0u16;
        if self.auto_republish {
            flags |= AUTO_REPUBLISH;
        }
        if self.single_file {
            flags |= MHTML;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&[0u8; 4]); // reserved3 + unused2
        payload.extend_from_slice(&self.style_id.to_le_bytes());
        payload.extend_from_slice(&(tail.len() as u32).to_le_bytes());
        payload.extend_from_slice(&tail);
        Ok(payload)
    }
}

/// Serialize a `WebPubString` (MS-XLS 2.5.278), compressed when every
/// character is in U+0000..=U+00FF and wide otherwise.
fn write_web_pub_string(out: &mut Vec<u8>, text: &str) -> XlsResult<()> {
    let compressible = text.chars().all(|ch| u32::from(ch) <= 0xFF);
    let char_count = if compressible {
        text.len()
    } else {
        text.encode_utf16().count()
    };
    if char_count > MAX_WEB_PUB_STRING_CHARS {
        return Err(XlsError::InvalidData(
            "WebPubString exceeds 255 characters".to_string(),
        ));
    }
    out.extend_from_slice(&(char_count as u16).to_le_bytes());
    if compressible {
        out.push(0u8); // fHighByte = 0
        out.extend(text.chars().map(|ch| ch as u8));
    } else {
        out.push(HIGH_BYTE);
        for unit in text.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(())
}

/// Parse a `WebPubString` (MS-XLS 2.5.278): a 2-byte character count
/// followed by an `XLUnicodeStringNoCch`. Returns the string and the number
/// of bytes consumed.
fn parse_web_pub_string(data: &[u8]) -> XlsResult<(String, usize)> {
    if data.len() < 3 {
        return Err(XlsError::InvalidLength {
            expected: 3,
            found: data.len(),
        });
    }
    let char_count = usize::from(read_u16(data, 0));
    if char_count > MAX_WEB_PUB_STRING_CHARS {
        return Err(invalid("WebPubString exceeds 255 characters"));
    }
    let wide = data[2] & HIGH_BYTE != 0;
    let byte_len = if wide { char_count * 2 } else { char_count };
    let bytes = data.get(3..3 + byte_len).ok_or(XlsError::InvalidLength {
        expected: 3 + byte_len,
        found: data.len(),
    })?;
    let text = if wide {
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16(&units).map_err(|_| invalid("WebPubString is not valid UTF-16LE"))?
    } else {
        bytes.iter().map(|&byte| char::from(byte)).collect()
    };
    Ok((text, 3 + byte_len))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a compressed WebPubString.
    fn web_pub_string(text: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&(text.len() as u16).to_le_bytes());
        out.push(0u8); // fHighByte = 0
        out.extend_from_slice(text.as_bytes());
        out
    }

    struct WebPubBuilder {
        tws: u8,
        twd: u8,
        has_ref: bool,
        flags: u16,
        style_id: u32,
        source_name: Option<String>,
        file_destination: String,
        div_id: String,
        title: String,
        chart_shape_id: Option<u32>,
        reserved: Vec<u8>,
    }

    impl WebPubBuilder {
        fn new(tws: u8) -> Self {
            WebPubBuilder {
                tws,
                twd: 0,
                has_ref: tws == 0x04,
                flags: 0,
                style_id: 1,
                source_name: None,
                file_destination: String::new(),
                div_id: String::new(),
                title: String::new(),
                chart_shape_id: None,
                reserved: Vec::new(),
            }
        }

        fn build(self) -> Vec<u8> {
            let mut tail = Vec::new();
            if self.tws > 0x04 && self.tws != 0xFF {
                tail.extend_from_slice(&web_pub_string(self.source_name.as_deref().unwrap_or("")));
            }
            tail.extend_from_slice(&web_pub_string(&self.file_destination));
            tail.extend_from_slice(&web_pub_string(&self.div_id));
            tail.extend_from_slice(&web_pub_string(&self.title));
            if let Some(shape_id) = self.chart_shape_id {
                tail.extend_from_slice(&shape_id.to_le_bytes());
            }
            tail.extend_from_slice(&self.reserved);
            tail.extend_from_slice(&[0u8; 2]); // unused3

            let mut payload = Vec::new();
            payload.extend_from_slice(&WEB_PUB_RECORD_TYPE.to_le_bytes());
            payload.extend_from_slice(&(if self.has_ref { FRT_REF } else { 0 }).to_le_bytes());
            payload.extend_from_slice(&1u16.to_le_bytes()); // rwFirst
            payload.extend_from_slice(&9u16.to_le_bytes()); // rwLast
            payload.extend_from_slice(&2u16.to_le_bytes()); // colFirst
            payload.extend_from_slice(&5u16.to_le_bytes()); // colLast
            payload.push(self.tws);
            payload.push(self.twd);
            payload.extend_from_slice(&self.flags.to_le_bytes());
            payload.extend_from_slice(&[0u8; 4]); // reserved3 + unused2
            payload.extend_from_slice(&self.style_id.to_le_bytes());
            payload.extend_from_slice(&(tail.len() as u32).to_le_bytes());
            payload.extend_from_slice(&tail);
            payload
        }
    }

    #[test]
    fn parses_workbook_publication() {
        let mut builder = WebPubBuilder::new(0x00);
        builder.twd = 0x01;
        builder.flags = AUTO_REPUBLISH | MHTML;
        builder.style_id = 0x1122_3344;
        builder.file_destination = "https://example.com/report.mht".to_string();
        builder.div_id = "bookmarks".to_string();
        builder.title = "Quarterly report".to_string();
        let pub_record = XlsWebPub::parse(&builder.build()).expect("parse");
        assert_eq!(pub_record.source, XlsWebSourceType::Workbook);
        assert_eq!(pub_record.page_type, XlsWebPageType::WorkbookFunctionality);
        assert_eq!(pub_record.range, None);
        assert!(pub_record.auto_republish);
        assert!(pub_record.single_file);
        assert_eq!(pub_record.style_id, 0x1122_3344);
        assert_eq!(pub_record.source_name, None);
        assert_eq!(
            pub_record.file_destination,
            "https://example.com/report.mht"
        );
        assert_eq!(pub_record.div_id, "bookmarks");
        assert_eq!(pub_record.title, "Quarterly report");
        assert_eq!(pub_record.chart_shape_id, None);
        assert!(pub_record.reserved.is_empty());
    }

    #[test]
    fn parses_range_publication_with_ref8() {
        let mut builder = WebPubBuilder::new(0x04);
        builder.file_destination = "C:\\pub\\range.htm".to_string();
        builder.title = "Range".to_string();
        let pub_record = XlsWebPub::parse(&builder.build()).expect("parse");
        assert_eq!(pub_record.source, XlsWebSourceType::Range);
        assert_eq!(
            pub_record.range,
            Some(XlsWebPubRange {
                first_row: 1,
                last_row: 9,
                first_column: 2,
                last_column: 5,
            })
        );
        assert_eq!(pub_record.source_name, None);
    }

    #[test]
    fn parses_chart_publication_with_source_name_and_shape_id() {
        let mut builder = WebPubBuilder::new(0x05);
        builder.source_name = Some("Chart 1".to_string());
        builder.chart_shape_id = Some(0x0401);
        builder.file_destination = "chart.htm".to_string();
        builder.title = "Chart".to_string();
        builder.reserved = vec![0xAA, 0xBB, 0xCC];
        let pub_record = XlsWebPub::parse(&builder.build()).expect("parse");
        assert_eq!(pub_record.source, XlsWebSourceType::Chart);
        assert_eq!(pub_record.source_name.as_deref(), Some("Chart 1"));
        assert_eq!(pub_record.chart_shape_id, Some(0x0401));
        assert_eq!(pub_record.reserved, vec![0xAA, 0xBB, 0xCC]);
    }

    #[test]
    fn parses_named_range_publication() {
        let mut builder = WebPubBuilder::new(0x08);
        builder.source_name = Some("Sales".to_string());
        builder.file_destination = "sales.htm".to_string();
        builder.title = "Sales".to_string();
        let pub_record = XlsWebPub::parse(&builder.build()).expect("parse");
        assert_eq!(pub_record.source, XlsWebSourceType::NamedRange);
        assert_eq!(pub_record.source_name.as_deref(), Some("Sales"));
        assert_eq!(pub_record.chart_shape_id, None);
    }

    #[test]
    fn rejects_ref_flag_mismatch() {
        // tws = Range without fFrtRef.
        let mut builder = WebPubBuilder::new(0x04);
        builder.has_ref = false;
        builder.file_destination = "x.htm".to_string();
        builder.title = "x".to_string();
        assert!(XlsWebPub::parse(&builder.build()).is_err());

        // tws = Workbook with fFrtRef set.
        let mut builder = WebPubBuilder::new(0x00);
        builder.has_ref = true;
        builder.file_destination = "x.htm".to_string();
        builder.title = "x".to_string();
        assert!(XlsWebPub::parse(&builder.build()).is_err());
    }

    #[test]
    fn rejects_cb_mismatch() {
        let mut builder = WebPubBuilder::new(0x00);
        builder.file_destination = "x.htm".to_string();
        builder.title = "x".to_string();
        let mut payload = builder.build();
        // Corrupt the cb field (offset 24).
        payload[24] = payload[24].wrapping_add(1);
        assert!(XlsWebPub::parse(&payload).is_err());
    }

    #[test]
    fn rejects_unknown_type_codes_and_rt_mismatch() {
        let mut builder = WebPubBuilder::new(0x09); // not in the tws table
        builder.file_destination = "x.htm".to_string();
        builder.title = "x".to_string();
        assert!(XlsWebPub::parse(&builder.build()).is_err());

        let mut builder = WebPubBuilder::new(0x00);
        builder.twd = 0x04; // not in the twd table
        builder.file_destination = "x.htm".to_string();
        builder.title = "x".to_string();
        assert!(XlsWebPub::parse(&builder.build()).is_err());

        let mut builder = WebPubBuilder::new(0x00);
        builder.file_destination = "x.htm".to_string();
        builder.title = "x".to_string();
        let mut payload = builder.build();
        payload[0] = 0x02; // corrupt rt to 0x0902... -> mismatch
        payload[1] = 0x09;
        assert!(XlsWebPub::parse(&payload).is_err());
    }

    #[test]
    fn rejects_truncated_and_oversized_strings() {
        assert!(XlsWebPub::parse(&[]).is_err());
        assert!(XlsWebPub::parse(&[0u8; 20]).is_err());

        // A WebPubString longer than 255 characters is illegal.
        let mut builder = WebPubBuilder::new(0x00);
        builder.file_destination = "x".repeat(300);
        builder.title = "x".to_string();
        assert!(XlsWebPub::parse(&builder.build()).is_err());
    }

    #[test]
    fn payload_round_trips() {
        let values = [
            XlsWebPub {
                source: XlsWebSourceType::Workbook,
                page_type: XlsWebPageType::WorkbookFunctionality,
                range: None,
                auto_republish: true,
                single_file: true,
                style_id: 0x1122_3344,
                source_name: None,
                file_destination: "https://example.com/report.mht".to_string(),
                div_id: "top".to_string(),
                title: "Quarterly report".to_string(),
                chart_shape_id: None,
                reserved: Vec::new(),
            },
            XlsWebPub {
                source: XlsWebSourceType::Range,
                page_type: XlsWebPageType::ViewOnly,
                range: Some(XlsWebPubRange {
                    first_row: 1,
                    last_row: 9,
                    first_column: 2,
                    last_column: 5,
                }),
                auto_republish: false,
                single_file: false,
                style_id: 7,
                source_name: None,
                file_destination: "C:\\pub\\range.htm".to_string(),
                div_id: String::new(),
                title: "Range €".to_string(),
                chart_shape_id: None,
                reserved: vec![0xAA, 0xBB],
            },
            XlsWebPub {
                source: XlsWebSourceType::Chart,
                page_type: XlsWebPageType::ChartFunctionality,
                range: None,
                auto_republish: true,
                single_file: false,
                style_id: 9,
                source_name: Some("Chart 1".to_string()),
                file_destination: "chart.htm".to_string(),
                div_id: "c1".to_string(),
                title: "Chart".to_string(),
                chart_shape_id: Some(0x0401),
                reserved: Vec::new(),
            },
        ];
        for value in values {
            let payload = value.to_payload().expect("serialize");
            let parsed = XlsWebPub::parse(&payload).expect("re-parse");
            assert_eq!(parsed, value);
        }
    }

    #[test]
    fn serialize_rejects_inconsistent_conditional_fields() {
        let mut value = XlsWebPub {
            source: XlsWebSourceType::Workbook,
            page_type: XlsWebPageType::ViewOnly,
            range: Some(XlsWebPubRange {
                first_row: 0,
                last_row: 1,
                first_column: 0,
                last_column: 1,
            }),
            auto_republish: false,
            single_file: false,
            style_id: 1,
            source_name: None,
            file_destination: "x.htm".to_string(),
            div_id: String::new(),
            title: "x".to_string(),
            chart_shape_id: None,
            reserved: Vec::new(),
        };
        assert!(value.to_payload().is_err());
        value.range = None;
        value.source_name = Some("unexpected".to_string());
        assert!(value.to_payload().is_err());
        value.source_name = None;
        value.chart_shape_id = Some(1);
        assert!(value.to_payload().is_err());

        // A WebPubString longer than 255 characters cannot be written.
        value.chart_shape_id = None;
        value.title = "x".repeat(300);
        assert!(value.to_payload().is_err());
    }
}
