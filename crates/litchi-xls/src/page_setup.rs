//! BIFF8 worksheet print and page setup records.

use super::number_format::{COMPRESSED_CHAR_BYTES, UTF16_CHAR_BYTES, XL_UNICODE_STRING_HIGH_BYTE};
use crate::error::{Error, Result};

const HEADER_RECORD_TYPE: u16 = 0x0014;
const FOOTER_RECORD_TYPE: u16 = 0x0015;
const VERTICAL_PAGE_BREAKS_RECORD_TYPE: u16 = 0x001a;
const HORIZONTAL_PAGE_BREAKS_RECORD_TYPE: u16 = 0x001b;
const PRINT_HEADERS_RECORD_TYPE: u16 = 0x002a;
const PRINT_GRIDLINES_RECORD_TYPE: u16 = 0x002b;
const LEFT_MARGIN_RECORD_TYPE: u16 = 0x0026;
const RIGHT_MARGIN_RECORD_TYPE: u16 = 0x0027;
const TOP_MARGIN_RECORD_TYPE: u16 = 0x0028;
const BOTTOM_MARGIN_RECORD_TYPE: u16 = 0x0029;
const PLS_RECORD_TYPE: u16 = 0x004d;
const CONTINUE_RECORD_TYPE: u16 = 0x003c;
const HCENTER_RECORD_TYPE: u16 = 0x0083;
const VCENTER_RECORD_TYPE: u16 = 0x0084;
const SETUP_RECORD_TYPE: u16 = 0x00a1;
const HEADER_FOOTER_RECORD_TYPE: u16 = 0x089c;
use super::custom_view::{
    USER_S_VIEW_BEGIN_RECORD_TYPE as USER_SVIEW_BEGIN_RECORD_TYPE,
    USER_S_VIEW_END_RECORD_TYPE as USER_SVIEW_END_RECORD_TYPE,
};

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_f64(data: &[u8], offset: usize) -> f64 {
    f64::from_le_bytes(data[offset..offset + 8].try_into().unwrap())
}

/// Order in which a multi-page worksheet is printed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PrintOrder {
    DownThenOver,
    OverThenDown,
}

/// Explicit paper orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PrintOrientation {
    Landscape,
    Portrait,
}

/// Representation used for cell errors when printing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PrintErrors {
    Displayed,
    Blank,
    Dashes,
    NotAvailable,
}

/// How cell comments are printed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PrintComments {
    None,
    AsDisplayed,
    AtEnd,
}

/// One explicit page break and the inclusive perpendicular page span.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageBreak {
    position: u16,
    range_start: u16,
    range_end: u16,
}

impl PageBreak {
    /// First row below a horizontal break, or first column right of a vertical break.
    pub const fn position(&self) -> u16 {
        self.position
    }
    pub const fn range_start(&self) -> u16 {
        self.range_start
    }
    pub const fn range_end(&self) -> u16 {
        self.range_end
    }
}

/// Fixed `SETUP` print configuration.
#[derive(Debug, Clone, PartialEq)]
pub struct PrintSetup {
    printer_settings_available: bool,
    paper_size: Option<u16>,
    scale_percent: Option<u16>,
    starting_page_number: Option<i16>,
    fit_width_pages: u16,
    fit_height_pages: u16,
    print_order: PrintOrder,
    orientation: Option<PrintOrientation>,
    black_and_white: bool,
    draft_quality: bool,
    comments: PrintComments,
    errors: PrintErrors,
    horizontal_resolution_dpi: Option<u16>,
    vertical_resolution_dpi: Option<u16>,
    header_margin_inches: f64,
    footer_margin_inches: f64,
    copies: Option<u16>,
}

impl PrintSetup {
    pub fn printer_settings_available(&self) -> bool {
        self.printer_settings_available
    }
    pub fn paper_size(&self) -> Option<u16> {
        self.paper_size
    }
    pub fn scale_percent(&self) -> Option<u16> {
        self.scale_percent
    }
    pub fn starting_page_number(&self) -> Option<i16> {
        self.starting_page_number
    }
    pub fn fit_width_pages(&self) -> u16 {
        self.fit_width_pages
    }
    pub fn fit_height_pages(&self) -> u16 {
        self.fit_height_pages
    }
    pub fn print_order(&self) -> PrintOrder {
        self.print_order
    }
    pub fn orientation(&self) -> Option<PrintOrientation> {
        self.orientation
    }
    pub fn is_black_and_white(&self) -> bool {
        self.black_and_white
    }
    pub fn is_draft_quality(&self) -> bool {
        self.draft_quality
    }
    pub fn comments(&self) -> PrintComments {
        self.comments
    }
    pub fn errors(&self) -> PrintErrors {
        self.errors
    }
    pub fn horizontal_resolution_dpi(&self) -> Option<u16> {
        self.horizontal_resolution_dpi
    }
    pub fn vertical_resolution_dpi(&self) -> Option<u16> {
        self.vertical_resolution_dpi
    }
    pub fn header_margin_inches(&self) -> f64 {
        self.header_margin_inches
    }
    pub fn footer_margin_inches(&self) -> f64 {
        self.footer_margin_inches
    }
    pub fn copies(&self) -> Option<u16> {
        self.copies
    }
}

impl Default for PrintSetup {
    fn default() -> Self {
        Self {
            printer_settings_available: false,
            paper_size: None,
            scale_percent: None,
            starting_page_number: None,
            fit_width_pages: 0,
            fit_height_pages: 0,
            print_order: PrintOrder::DownThenOver,
            orientation: None,
            black_and_white: false,
            draft_quality: false,
            comments: PrintComments::None,
            errors: PrintErrors::Displayed,
            horizontal_resolution_dpi: None,
            vertical_resolution_dpi: None,
            header_margin_inches: 0.5,
            footer_margin_inches: 0.5,
            copies: None,
        }
    }
}

/// Print configuration associated with a worksheet.
#[derive(Debug, Clone, PartialEq)]
pub struct PageSetup {
    print_headers: bool,
    print_gridlines: bool,
    header: String,
    footer: String,
    horizontally_centered: bool,
    vertically_centered: bool,
    left_margin_inches: Option<f64>,
    right_margin_inches: Option<f64>,
    top_margin_inches: Option<f64>,
    bottom_margin_inches: Option<f64>,
    horizontal_page_breaks: Vec<PageBreak>,
    vertical_page_breaks: Vec<PageBreak>,
    printer_driver_data: Vec<Vec<u8>>,
    print_setup: PrintSetup,
    header_footer: Option<HeaderFooter>,
}

impl PageSetup {
    pub const fn print_headers(&self) -> bool {
        self.print_headers
    }
    pub const fn print_gridlines(&self) -> bool {
        self.print_gridlines
    }
    /// Raw header text, including `&L`, `&C`, `&R`, and formatting commands.
    pub fn header(&self) -> &str {
        &self.header
    }
    /// Raw footer text, including `&L`, `&C`, `&R`, and formatting commands.
    pub fn footer(&self) -> &str {
        &self.footer
    }
    pub fn is_horizontally_centered(&self) -> bool {
        self.horizontally_centered
    }
    pub fn is_vertically_centered(&self) -> bool {
        self.vertically_centered
    }
    pub fn left_margin_inches(&self) -> Option<f64> {
        self.left_margin_inches
    }
    pub fn right_margin_inches(&self) -> Option<f64> {
        self.right_margin_inches
    }
    pub fn top_margin_inches(&self) -> Option<f64> {
        self.top_margin_inches
    }
    pub fn bottom_margin_inches(&self) -> Option<f64> {
        self.bottom_margin_inches
    }
    pub fn horizontal_page_breaks(&self) -> &[PageBreak] {
        &self.horizontal_page_breaks
    }
    pub fn vertical_page_breaks(&self) -> &[PageBreak] {
        &self.vertical_page_breaks
    }
    /// Opaque DEVMODE payloads from `PLS`; these bytes are never executed.
    pub fn printer_driver_data(&self) -> &[Vec<u8>] {
        &self.printer_driver_data
    }
    pub fn print_setup(&self) -> &PrintSetup {
        &self.print_setup
    }
    /// Even-page and first-page header/footer text and display flags from the
    /// `HeaderFooter` record, when present.
    pub fn header_footer(&self) -> Option<&HeaderFooter> {
        self.header_footer.as_ref()
    }
}

fn parse_header_footer(data: &[u8], record_type: u16) -> Result<String> {
    if data.is_empty() {
        return Ok(String::new());
    }
    if data.len() < 3 {
        return Err(invalid(
            record_type,
            "header/footer string has a truncated XLUnicodeString header",
        ));
    }
    let count = usize::from(read_u16(data, 0));
    let flags = data[2];
    if count > 255 {
        return Err(invalid(
            record_type,
            "header/footer text exceeds 255 UTF-16 code units",
        ));
    }
    // MS-XLS 2.5.293: every option bit other than `fHighByte` is reserved and
    // "MUST be zero, and MUST be ignored", so a writer that leaves one set does
    // not make the record unreadable.
    let width = if flags & XL_UNICODE_STRING_HIGH_BYTE != 0 {
        UTF16_CHAR_BYTES
    } else {
        COMPRESSED_CHAR_BYTES
    };
    let byte_count = count
        .checked_mul(width)
        .ok_or_else(|| invalid(record_type, "header/footer string length overflow"))?;
    if data.len() != 3 + byte_count {
        return Err(invalid(
            record_type,
            "header/footer character count does not match its payload",
        ));
    }
    if width == 1 {
        Ok(data[3..].iter().map(|&byte| char::from(byte)).collect())
    } else {
        let units = data[3..]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_| invalid(record_type, "header/footer contains invalid UTF-16"))
    }
}

/// Size in bytes of the fixed `HeaderFooter` portion before the four strings:
/// `FrtHeader` (12) + `guidSView` (16) + flags (2) + four character counts (8).
const HEADER_FOOTER_FIXED_LEN: usize = 38;
/// `HeaderFooter` flag: odd and even pages use different headers/footers.
const HF_DIFF_ODD_EVEN: u16 = 0x0001;
/// `HeaderFooter` flag: the first page uses a different header/footer.
const HF_DIFF_FIRST: u16 = 0x0002;
/// `HeaderFooter` flag: the header/footer is scaled with the sheet.
const HF_SCALE_WITH_DOC: u16 = 0x0004;
/// `HeaderFooter` flag: header/footer edges align with the page margins.
const HF_ALIGN_MARGINS: u16 = 0x0008;
/// Maximum length of a header/footer string in UTF-16 code units.
const MAX_HEADER_FOOTER_CHARS: usize = 255;

/// Typed `HeaderFooter` record (MS-XLS 2.4.137): even-page and first-page
/// header/footer text plus header/footer display flags.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooter {
    /// GUID of the sheet view the record applies to; `None` for the sheet
    /// itself (an all-zero `guidSView`).
    sheet_view_guid: Option<[u8; 16]>,
    /// Whether odd and even pages use different headers/footers.
    diff_odd_even: bool,
    /// Whether the first page uses a different header/footer.
    diff_first: bool,
    /// Whether the header/footer is scaled with the sheet.
    scale_with_doc: bool,
    /// Whether header/footer edges align with the page margins.
    align_margins: bool,
    /// Even-page header text (`&L`/`&C`/`&R` commands included verbatim).
    even_header: String,
    /// Even-page footer text.
    even_footer: String,
    /// First-page header text.
    first_header: String,
    /// First-page footer text.
    first_footer: String,
}

impl HeaderFooter {
    pub fn sheet_view_guid(&self) -> Option<&[u8; 16]> {
        self.sheet_view_guid.as_ref()
    }
    pub const fn diff_odd_even(&self) -> bool {
        self.diff_odd_even
    }
    pub const fn diff_first(&self) -> bool {
        self.diff_first
    }
    pub const fn scale_with_doc(&self) -> bool {
        self.scale_with_doc
    }
    pub const fn align_margins(&self) -> bool {
        self.align_margins
    }
    pub fn even_header(&self) -> &str {
        &self.even_header
    }
    pub fn even_footer(&self) -> &str {
        &self.even_footer
    }
    pub fn first_header(&self) -> &str {
        &self.first_header
    }
    pub fn first_footer(&self) -> &str {
        &self.first_footer
    }

    /// Even-page header/footer text; setting either marks the record as
    /// differentiating odd and even pages.
    pub fn set_even(&mut self, header: String, footer: String) -> Result<()> {
        validate_header_footer_text(&header)?;
        validate_header_footer_text(&footer)?;
        self.diff_odd_even = true;
        self.even_header = header;
        self.even_footer = footer;
        Ok(())
    }

    /// First-page header/footer text; setting either marks the record as
    /// differentiating the first page.
    pub fn set_first(&mut self, header: String, footer: String) -> Result<()> {
        validate_header_footer_text(&header)?;
        validate_header_footer_text(&footer)?;
        self.diff_first = true;
        self.first_header = header;
        self.first_footer = footer;
        Ok(())
    }

    pub fn set_scale_with_doc(&mut self, scale_with_doc: bool) {
        self.scale_with_doc = scale_with_doc;
    }

    pub fn set_align_margins(&mut self, align_margins: bool) {
        self.align_margins = align_margins;
    }

    /// Parse a `HeaderFooter` record payload.
    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < HEADER_FOOTER_FIXED_LEN {
            return Err(Error::InvalidLength {
                expected: HEADER_FOOTER_FIXED_LEN,
                found: data.len(),
            });
        }
        if read_u16(data, 0) != HEADER_FOOTER_RECORD_TYPE {
            return Err(invalid(
                HEADER_FOOTER_RECORD_TYPE,
                "HeaderFooter FrtHeader.rt mismatch",
            ));
        }
        let guid: [u8; 16] = data[12..28].try_into().expect("length checked");
        let flags = read_u16(data, 28);
        let diff_odd_even = flags & HF_DIFF_ODD_EVEN != 0;
        let diff_first = flags & HF_DIFF_FIRST != 0;
        let counts = [
            usize::from(read_u16(data, 30)),
            usize::from(read_u16(data, 32)),
            usize::from(read_u16(data, 34)),
            usize::from(read_u16(data, 36)),
        ];
        for (index, &count) in counts.iter().enumerate() {
            if count > MAX_HEADER_FOOTER_CHARS {
                return Err(invalid(
                    HEADER_FOOTER_RECORD_TYPE,
                    "header/footer text exceeds 255 UTF-16 code units",
                ));
            }
            let gated = match index {
                0 | 1 => diff_odd_even,
                _ => diff_first,
            };
            if count > 0 && !gated {
                return Err(invalid(
                    HEADER_FOOTER_RECORD_TYPE,
                    "header/footer text present without its difference flag",
                ));
            }
        }
        let mut offset = HEADER_FOOTER_FIXED_LEN;
        let mut strings = [String::new(), String::new(), String::new(), String::new()];
        for (string, &count) in strings.iter_mut().zip(&counts) {
            if count == 0 {
                continue;
            }
            *string = parse_no_cch_string(data, &mut offset, count)?;
        }
        if offset != data.len() {
            return Err(invalid(
                HEADER_FOOTER_RECORD_TYPE,
                "header/footer character counts do not match the payload",
            ));
        }
        let [even_header, even_footer, first_header, first_footer] = strings;
        Ok(Self {
            sheet_view_guid: (guid != [0; 16]).then_some(guid),
            diff_odd_even,
            diff_first,
            scale_with_doc: flags & HF_SCALE_WITH_DOC != 0,
            align_margins: flags & HF_ALIGN_MARGINS != 0,
            even_header,
            even_footer,
            first_header,
            first_footer,
        })
    }

    /// Serialize back to a complete `HeaderFooter` record payload.
    pub(crate) fn to_payload(&self) -> Result<Vec<u8>> {
        for text in [
            &self.even_header,
            &self.even_footer,
            &self.first_header,
            &self.first_footer,
        ] {
            validate_header_footer_text(text)?;
        }
        if (!self.diff_odd_even && (!self.even_header.is_empty() || !self.even_footer.is_empty()))
            || (!self.diff_first
                && (!self.first_header.is_empty() || !self.first_footer.is_empty()))
        {
            return Err(invalid(
                HEADER_FOOTER_RECORD_TYPE,
                "header/footer text requires its difference flag",
            ));
        }
        let mut flags = 0u16;
        if self.diff_odd_even {
            flags |= HF_DIFF_ODD_EVEN;
        }
        if self.diff_first {
            flags |= HF_DIFF_FIRST;
        }
        if self.scale_with_doc {
            flags |= HF_SCALE_WITH_DOC;
        }
        if self.align_margins {
            flags |= HF_ALIGN_MARGINS;
        }
        let mut payload = Vec::with_capacity(HEADER_FOOTER_FIXED_LEN);
        payload.extend_from_slice(&HEADER_FOOTER_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; 10]);
        payload.extend_from_slice(&self.sheet_view_guid.unwrap_or([0; 16]));
        payload.extend_from_slice(&flags.to_le_bytes());
        for text in [
            &self.even_header,
            &self.even_footer,
            &self.first_header,
            &self.first_footer,
        ] {
            payload.extend_from_slice(&(text.encode_utf16().count() as u16).to_le_bytes());
        }
        for text in [
            &self.even_header,
            &self.even_footer,
            &self.first_header,
            &self.first_footer,
        ] {
            write_no_cch_string(&mut payload, text);
        }
        Ok(payload)
    }
}

fn validate_header_footer_text(text: &str) -> Result<()> {
    if text.encode_utf16().count() > MAX_HEADER_FOOTER_CHARS {
        return Err(invalid(
            HEADER_FOOTER_RECORD_TYPE,
            "header/footer text exceeds 255 UTF-16 code units",
        ));
    }
    Ok(())
}

/// Parse an `XLUnicodeStringNoCch` (MS-XLS 2.5.296) of `count` characters,
/// advancing `offset` past the consumed bytes.
fn parse_no_cch_string(data: &[u8], offset: &mut usize, count: usize) -> Result<String> {
    let flags = *data.get(*offset).ok_or(Error::InvalidLength {
        expected: *offset + 1,
        found: data.len(),
    })?;
    *offset += 1;
    let width = if flags & XL_UNICODE_STRING_HIGH_BYTE != 0 {
        UTF16_CHAR_BYTES
    } else {
        COMPRESSED_CHAR_BYTES
    };
    let byte_count = count
        .checked_mul(width)
        .ok_or_else(|| invalid(HEADER_FOOTER_RECORD_TYPE, "header/footer length overflow"))?;
    let text = data
        .get(*offset..*offset + byte_count)
        .ok_or(Error::InvalidLength {
            expected: *offset + byte_count,
            found: data.len(),
        })?;
    *offset += byte_count;
    if width == COMPRESSED_CHAR_BYTES {
        Ok(text.iter().map(|&byte| char::from(byte)).collect())
    } else {
        let units = text
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_| invalid(HEADER_FOOTER_RECORD_TYPE, "header/footer invalid UTF-16"))
    }
}

/// Append an `XLUnicodeStringNoCch`, compressing to one byte per character
/// when every code unit fits.
fn write_no_cch_string(out: &mut Vec<u8>, text: &str) {
    let units = text.encode_utf16().collect::<Vec<_>>();
    if units.is_empty() {
        return;
    }
    let compressed = units.iter().all(|&unit| unit <= 0x00FF);
    out.push(u8::from(!compressed));
    for unit in units {
        if compressed {
            out.push(unit as u8);
        } else {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
}

fn parse_bool(data: &[u8], record_type: u16) -> Result<bool> {
    if data.len() != 2 {
        return Err(invalid(
            record_type,
            format!("Boolean payload must be 2 bytes, found {}", data.len()),
        ));
    }
    match read_u16(data, 0) {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(invalid(record_type, "Boolean value must be zero or one")),
    }
}

fn parse_margin(data: &[u8], record_type: u16) -> Result<f64> {
    if data.len() != 8 {
        return Err(invalid(
            record_type,
            format!("margin payload must be 8 bytes, found {}", data.len()),
        ));
    }
    let margin = read_f64(data, 0);
    if !margin.is_finite() || !(0.0..=49.0).contains(&margin) {
        return Err(invalid(
            record_type,
            "page margin must be finite and between 0 and 49 inches",
        ));
    }
    Ok(margin)
}

fn parse_page_breaks(data: &[u8], record_type: u16) -> Result<Vec<PageBreak>> {
    if data.len() < 2 {
        return Err(invalid(
            record_type,
            "page-break record is missing its count",
        ));
    }
    let count = usize::from(read_u16(data, 0));
    let maximum = if record_type == HORIZONTAL_PAGE_BREAKS_RECORD_TYPE {
        1026
    } else {
        255
    };
    if count > maximum {
        return Err(invalid(
            record_type,
            format!("page-break count exceeds {maximum}"),
        ));
    }
    let expected = 2 + count * 6;
    if data.len() != expected {
        return Err(invalid(
            record_type,
            format!(
                "page-break count requires {expected} bytes, found {}",
                data.len()
            ),
        ));
    }
    let mut breaks: Vec<PageBreak> = Vec::with_capacity(count);
    for chunk in data[2..].chunks_exact(6) {
        let page_break = PageBreak {
            position: read_u16(chunk, 0),
            range_start: read_u16(chunk, 2),
            range_end: read_u16(chunk, 4),
        };
        if page_break.range_end <= page_break.range_start {
            return Err(invalid(
                record_type,
                "page-break range end must be greater than its start",
            ));
        }
        if record_type == HORIZONTAL_PAGE_BREAKS_RECORD_TYPE && page_break.range_end > 16383 {
            return Err(invalid(
                record_type,
                "horizontal page-break column exceeds 16383",
            ));
        }
        if record_type == VERTICAL_PAGE_BREAKS_RECORD_TYPE && page_break.position > 255 {
            return Err(invalid(
                record_type,
                "vertical page-break column exceeds 255",
            ));
        }
        if let Some(previous) = breaks.last() {
            if (page_break.position, page_break.range_start)
                < (previous.position, previous.range_start)
            {
                return Err(invalid(record_type, "page breaks are not sorted"));
            }
            if page_break.position == previous.position
                && page_break.range_start <= previous.range_end
            {
                return Err(invalid(record_type, "page-break ranges overlap"));
            }
        }
        breaks.push(page_break);
    }
    Ok(breaks)
}

fn parse_pls(data: &[u8]) -> Result<Vec<u8>> {
    if data.len() < 2 {
        return Err(invalid(
            PLS_RECORD_TYPE,
            "PLS is missing its reserved field",
        ));
    }
    if read_u16(data, 0) != 0 {
        return Err(invalid(PLS_RECORD_TYPE, "PLS reserved field must be zero"));
    }
    Ok(data[2..].to_vec())
}

fn parse_setup(data: &[u8]) -> Result<PrintSetup> {
    if data.len() != 34 {
        return Err(invalid(
            SETUP_RECORD_TYPE,
            format!("SETUP payload must be 34 bytes, found {}", data.len()),
        ));
    }
    let paper_size = read_u16(data, 0);
    let flags = read_u16(data, 10);
    let fit_width_pages = read_u16(data, 6);
    let fit_height_pages = read_u16(data, 8);
    let no_printer_settings = flags & 0x0004 != 0;
    if flags & 0xf000 != 0 {
        return Err(invalid(
            SETUP_RECORD_TYPE,
            "SETUP reserved flag bits must be zero",
        ));
    }
    if fit_width_pages > 32767 || fit_height_pages > 32767 {
        return Err(invalid(
            SETUP_RECORD_TYPE,
            "SETUP fit dimensions must not exceed 32767 pages",
        ));
    }
    if !no_printer_settings && (118..=255).contains(&paper_size) {
        return Err(invalid(
            SETUP_RECORD_TYPE,
            "SETUP paper size uses a reserved identifier",
        ));
    }
    let header_margin_inches = read_f64(data, 16);
    let footer_margin_inches = read_f64(data, 24);
    if !header_margin_inches.is_finite()
        || !(0.0..49.0).contains(&header_margin_inches)
        || !footer_margin_inches.is_finite()
        || !(0.0..49.0).contains(&footer_margin_inches)
    {
        return Err(invalid(
            SETUP_RECORD_TYPE,
            "SETUP header/footer margins must be finite and less than 49 inches",
        ));
    }
    let printer_value = |value| (!no_printer_settings).then_some(value);
    let orientation = if no_printer_settings || flags & 0x0040 != 0 {
        None
    } else if flags & 0x0002 != 0 {
        Some(PrintOrientation::Portrait)
    } else {
        Some(PrintOrientation::Landscape)
    };
    let comments = if flags & 0x0020 == 0 {
        PrintComments::None
    } else if flags & 0x0200 != 0 {
        PrintComments::AtEnd
    } else {
        PrintComments::AsDisplayed
    };
    let errors = match (flags >> 10) & 3 {
        0 => PrintErrors::Displayed,
        1 => PrintErrors::Blank,
        2 => PrintErrors::Dashes,
        _ => PrintErrors::NotAvailable,
    };

    Ok(PrintSetup {
        printer_settings_available: !no_printer_settings,
        paper_size: printer_value(paper_size),
        scale_percent: printer_value(read_u16(data, 2)),
        starting_page_number: (flags & 0x0080 != 0).then_some(read_u16(data, 4) as i16),
        fit_width_pages,
        fit_height_pages,
        print_order: if flags & 0x0001 != 0 {
            PrintOrder::OverThenDown
        } else {
            PrintOrder::DownThenOver
        },
        orientation,
        black_and_white: flags & 0x0008 != 0,
        draft_quality: flags & 0x0010 != 0,
        comments,
        errors,
        horizontal_resolution_dpi: printer_value(read_u16(data, 12)),
        vertical_resolution_dpi: printer_value(read_u16(data, 14)),
        header_margin_inches,
        footer_margin_inches,
        copies: printer_value(read_u16(data, 32)),
    })
}

#[derive(Default)]
struct PartialPageSetup {
    print_headers: Option<bool>,
    print_gridlines: Option<bool>,
    header: Option<String>,
    footer: Option<String>,
    horizontally_centered: Option<bool>,
    vertically_centered: Option<bool>,
    left_margin_inches: Option<f64>,
    right_margin_inches: Option<f64>,
    top_margin_inches: Option<f64>,
    bottom_margin_inches: Option<f64>,
    horizontal_page_breaks: Option<Vec<PageBreak>>,
    vertical_page_breaks: Option<Vec<PageBreak>>,
    printer_driver_data: Vec<Vec<u8>>,
    print_setup: Option<PrintSetup>,
    header_footer: Option<HeaderFooter>,
}

/// Collects primary worksheet page records while excluding custom views.
pub(crate) struct PageSetupCollector {
    page: PartialPageSetup,
    in_custom_view: bool,
    collecting_pls: bool,
}

impl PageSetupCollector {
    pub(crate) fn new() -> Self {
        Self {
            page: PartialPageSetup::default(),
            in_custom_view: false,
            collecting_pls: false,
        }
    }

    fn duplicate(record_type: u16) -> Error {
        invalid(
            record_type,
            "worksheet contains a duplicate primary page-setup record",
        )
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        if self.collecting_pls {
            if record_type == CONTINUE_RECORD_TYPE {
                self.page
                    .printer_driver_data
                    .last_mut()
                    .unwrap()
                    .extend_from_slice(data);
                return Ok(());
            }
            self.collecting_pls = false;
        }
        if record_type == USER_SVIEW_BEGIN_RECORD_TYPE {
            self.in_custom_view = true;
            return Ok(());
        }
        if record_type == USER_SVIEW_END_RECORD_TYPE {
            self.in_custom_view = false;
            return Ok(());
        }
        if self.in_custom_view {
            return Ok(());
        }
        match record_type {
            PRINT_HEADERS_RECORD_TYPE => {
                if self.page.print_headers.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.print_headers = Some(parse_bool(data, record_type)?);
            },
            PRINT_GRIDLINES_RECORD_TYPE => {
                if self.page.print_gridlines.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                if data.len() != 2 {
                    return Err(invalid(
                        record_type,
                        "PRINTGRIDLINES payload must be 2 bytes",
                    ));
                }
                self.page.print_gridlines = Some(read_u16(data, 0) & 1 != 0);
            },
            HORIZONTAL_PAGE_BREAKS_RECORD_TYPE => {
                if self.page.horizontal_page_breaks.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.horizontal_page_breaks = Some(parse_page_breaks(data, record_type)?);
            },
            VERTICAL_PAGE_BREAKS_RECORD_TYPE => {
                if self.page.vertical_page_breaks.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.vertical_page_breaks = Some(parse_page_breaks(data, record_type)?);
            },
            HEADER_RECORD_TYPE => {
                if self.page.header.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.header = Some(parse_header_footer(data, record_type)?);
            },
            FOOTER_RECORD_TYPE => {
                if self.page.footer.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.footer = Some(parse_header_footer(data, record_type)?);
            },
            HCENTER_RECORD_TYPE => {
                if self.page.horizontally_centered.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.horizontally_centered = Some(parse_bool(data, record_type)?);
            },
            VCENTER_RECORD_TYPE => {
                if self.page.vertically_centered.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.vertically_centered = Some(parse_bool(data, record_type)?);
            },
            LEFT_MARGIN_RECORD_TYPE => {
                if self.page.left_margin_inches.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.left_margin_inches = Some(parse_margin(data, record_type)?);
            },
            RIGHT_MARGIN_RECORD_TYPE => {
                if self.page.right_margin_inches.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.right_margin_inches = Some(parse_margin(data, record_type)?);
            },
            TOP_MARGIN_RECORD_TYPE => {
                if self.page.top_margin_inches.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.top_margin_inches = Some(parse_margin(data, record_type)?);
            },
            BOTTOM_MARGIN_RECORD_TYPE => {
                if self.page.bottom_margin_inches.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.bottom_margin_inches = Some(parse_margin(data, record_type)?);
            },
            HEADER_FOOTER_RECORD_TYPE => {
                if self.page.header_footer.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.header_footer = Some(HeaderFooter::parse(data)?);
            },
            PLS_RECORD_TYPE => {
                self.page.printer_driver_data.push(parse_pls(data)?);
                self.collecting_pls = true;
            },
            SETUP_RECORD_TYPE => {
                if self.page.print_setup.is_some() {
                    return Err(Self::duplicate(record_type));
                }
                self.page.print_setup = Some(parse_setup(data)?);
            },
            _ => {},
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> Result<Option<PageSetup>> {
        let has_partial = self.page.print_headers.is_some()
            || self.page.print_gridlines.is_some()
            || self.page.horizontal_page_breaks.is_some()
            || self.page.vertical_page_breaks.is_some()
            || !self.page.printer_driver_data.is_empty()
            || self.page.header.is_some()
            || self.page.footer.is_some()
            || self.page.horizontally_centered.is_some()
            || self.page.vertically_centered.is_some()
            || self.page.left_margin_inches.is_some()
            || self.page.right_margin_inches.is_some()
            || self.page.top_margin_inches.is_some()
            || self.page.bottom_margin_inches.is_some()
            || self.page.header_footer.is_some();
        if !has_partial && self.page.print_setup.is_none() {
            return Ok(None);
        }
        Ok(Some(PageSetup {
            print_headers: self.page.print_headers.unwrap_or(false),
            print_gridlines: self.page.print_gridlines.unwrap_or(false),
            header: self.page.header.unwrap_or_default(),
            footer: self.page.footer.unwrap_or_default(),
            horizontally_centered: self.page.horizontally_centered.unwrap_or(false),
            vertically_centered: self.page.vertically_centered.unwrap_or(false),
            left_margin_inches: self.page.left_margin_inches,
            right_margin_inches: self.page.right_margin_inches,
            top_margin_inches: self.page.top_margin_inches,
            bottom_margin_inches: self.page.bottom_margin_inches,
            horizontal_page_breaks: self.page.horizontal_page_breaks.unwrap_or_default(),
            vertical_page_breaks: self.page.vertical_page_breaks.unwrap_or_default(),
            printer_driver_data: self.page.printer_driver_data,
            print_setup: self.page.print_setup.unwrap_or_default(),
            header_footer: self.page.header_footer,
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn setup(flags: u16) -> [u8; 34] {
        let mut data = [0u8; 34];
        data[0..2].copy_from_slice(&1u16.to_le_bytes());
        data[2..4].copy_from_slice(&100u16.to_le_bytes());
        data[6..8].copy_from_slice(&1u16.to_le_bytes());
        data[8..10].copy_from_slice(&1u16.to_le_bytes());
        data[10..12].copy_from_slice(&flags.to_le_bytes());
        data[12..14].copy_from_slice(&600u16.to_le_bytes());
        data[14..16].copy_from_slice(&600u16.to_le_bytes());
        data[16..24].copy_from_slice(&0.5f64.to_le_bytes());
        data[24..32].copy_from_slice(&0.5f64.to_le_bytes());
        data[32..34].copy_from_slice(&1u16.to_le_bytes());
        data
    }

    fn compressed(text: &str) -> Vec<u8> {
        let mut data = Vec::with_capacity(3 + text.len());
        data.extend_from_slice(&(text.len() as u16).to_le_bytes());
        data.push(0);
        data.extend_from_slice(text.as_bytes());
        data
    }

    #[test]
    fn parses_typed_page_setup() {
        let mut collector = PageSetupCollector::new();
        collector
            .feed_record(HEADER_RECORD_TYPE, &compressed("&Ctitle"))
            .unwrap();
        collector.feed_record(FOOTER_RECORD_TYPE, &[]).unwrap();
        collector.feed_record(HCENTER_RECORD_TYPE, &[1, 0]).unwrap();
        collector.feed_record(VCENTER_RECORD_TYPE, &[0, 0]).unwrap();
        collector
            .feed_record(LEFT_MARGIN_RECORD_TYPE, &0.7f64.to_le_bytes())
            .unwrap();
        collector
            .feed_record(SETUP_RECORD_TYPE, &setup(0x0002))
            .unwrap();
        let page = collector.finish().unwrap().unwrap();
        assert_eq!(page.header(), "&Ctitle");
        assert!(page.is_horizontally_centered());
        assert_eq!(
            page.print_setup().orientation(),
            Some(PrintOrientation::Portrait)
        );
        assert_eq!(page.print_setup().horizontal_resolution_dpi(), Some(600));
        assert_eq!(
            parse_header_footer(&compressed("&Cone && two &&&&"), HEADER_RECORD_TYPE).unwrap(),
            "&Cone && two &&&&"
        );
    }

    #[test]
    fn rejects_malformed_and_duplicate_page_records() {
        assert!(parse_header_footer(&[1, 0], HEADER_RECORD_TYPE).is_err());
        assert!(parse_bool(&[2, 0], HCENTER_RECORD_TYPE).is_err());
        assert!(parse_margin(&f64::NAN.to_le_bytes(), LEFT_MARGIN_RECORD_TYPE).is_err());
        assert!(parse_setup(&[0; 33]).is_err());

        let mut collector = PageSetupCollector::new();
        collector.feed_record(HEADER_RECORD_TYPE, &[]).unwrap();
        assert!(collector.feed_record(HEADER_RECORD_TYPE, &[]).is_err());

        assert!(
            parse_page_breaks(
                &[1, 0, 4, 0, 2, 0, 2, 0],
                HORIZONTAL_PAGE_BREAKS_RECORD_TYPE
            )
            .is_err()
        );
        let overlapping = [2, 0, 4, 0, 0, 0, 10, 0, 4, 0, 9, 0, 12, 0];
        assert!(parse_page_breaks(&overlapping, HORIZONTAL_PAGE_BREAKS_RECORD_TYPE).is_err());
        assert!(parse_pls(&[1, 0]).is_err());
    }

    #[test]
    fn partial_page_block_uses_tolerant_defaults() {
        let mut collector = PageSetupCollector::new();
        collector
            .feed_record(PRINT_HEADERS_RECORD_TYPE, &[1, 0])
            .unwrap();
        let page = collector.finish().unwrap().unwrap();
        assert!(page.print_headers());
        assert!(!page.print_gridlines());
        assert!(!page.print_setup().printer_settings_available());
    }

    #[test]
    fn ignores_custom_view_page_records() {
        let mut collector = PageSetupCollector::new();
        collector
            .feed_record(HEADER_RECORD_TYPE, &compressed("primary"))
            .unwrap();
        collector
            .feed_record(USER_SVIEW_BEGIN_RECORD_TYPE, &[])
            .unwrap();
        collector
            .feed_record(HEADER_RECORD_TYPE, &compressed("custom"))
            .unwrap();
        collector
            .feed_record(USER_SVIEW_END_RECORD_TYPE, &[])
            .unwrap();
        collector
            .feed_record(SETUP_RECORD_TYPE, &setup(0x0002))
            .unwrap();
        assert_eq!(collector.finish().unwrap().unwrap().header(), "primary");

        let mut omitted = PageSetupCollector::new();
        omitted
            .feed_record(SETUP_RECORD_TYPE, &setup(0x0002))
            .unwrap();
        let page = omitted.finish().unwrap().unwrap();
        assert_eq!(page.header(), "");
        assert_eq!(page.footer(), "");
    }

    #[test]
    fn reads_poi_page_setup_fixtures() {
        use crate::Workbook;
        use std::fs::File;
        use std::path::Path;

        let fixture = |name: &str| {
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/spreadsheet")
                .join(name)
        };

        let dbcs = Workbook::new(File::open(fixture("DBCSHeader.xls")).unwrap()).unwrap();
        let page = dbcs.xls_worksheet(0).unwrap().page_setup().unwrap();
        assert_eq!(
            page.header(),
            "&L\u{090f}\u{0915}&C\u{0939}\u{094b}\u{0917}\u{093e}&R\u{091c}\u{093e}"
        );
        assert_eq!(
            page.footer(),
            "&L\u{091c}\u{093e}&C\u{091c}\u{093e}&R\u{091c}\u{093e}"
        );
        assert_eq!(
            page.print_setup().orientation(),
            Some(PrintOrientation::Portrait)
        );
        assert_eq!(page.print_setup().paper_size(), Some(1));

        let breaks =
            Workbook::new(File::open(fixture("SimpleWithPageBreaks.xls")).unwrap()).unwrap();
        let page = breaks.xls_worksheet(0).unwrap().page_setup().unwrap();
        assert!(
            !page.horizontal_page_breaks().is_empty() || !page.vertical_page_breaks().is_empty()
        );
    }

    fn header_footer_record(flags: u16, counts: [u16; 4], strings: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&HEADER_FOOTER_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 26]);
        data.extend_from_slice(&flags.to_le_bytes());
        for count in counts {
            data.extend_from_slice(&count.to_le_bytes());
        }
        data.extend_from_slice(strings);
        data
    }

    #[test]
    fn parses_header_footer_even_and_first_pages() {
        // Even header "EH" (compressed), even footer empty, first header
        // "文" (UTF-16), first footer empty; scale and align flags set.
        let mut strings = Vec::new();
        strings.push(0);
        strings.extend_from_slice(b"EH");
        strings.push(1);
        strings.extend_from_slice(&0x6587u16.to_le_bytes());
        let data = header_footer_record(0x000F, [2, 0, 1, 0], &strings);

        let parsed = HeaderFooter::parse(&data).unwrap();
        assert!(parsed.diff_odd_even());
        assert!(parsed.diff_first());
        assert!(parsed.scale_with_doc());
        assert!(parsed.align_margins());
        assert_eq!(parsed.even_header(), "EH");
        assert_eq!(parsed.even_footer(), "");
        assert_eq!(parsed.first_header(), "文");
        assert_eq!(parsed.first_footer(), "");
        assert!(parsed.sheet_view_guid().is_none());
        assert_eq!(parsed.to_payload().unwrap(), data);
    }

    #[test]
    fn parses_header_footer_without_strings() {
        let data = header_footer_record(0x000C, [0, 0, 0, 0], &[]);
        let parsed = HeaderFooter::parse(&data).unwrap();
        assert!(!parsed.diff_odd_even());
        assert!(!parsed.diff_first());
        assert!(parsed.scale_with_doc());
        assert!(parsed.align_margins());
        assert_eq!(parsed.to_payload().unwrap(), data);
    }

    #[test]
    fn rejects_malformed_header_footer_records() {
        // Truncated fixed portion.
        assert!(HeaderFooter::parse(&[0; 30]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = header_footer_record(0x000C, [0, 0, 0, 0], &[]);
        wrong_rt[0..2].copy_from_slice(&0x0862u16.to_le_bytes());
        assert!(HeaderFooter::parse(&wrong_rt).is_err());
        // Text without its difference flag.
        let mut strings = Vec::new();
        strings.push(0);
        strings.extend_from_slice(b"EH");
        assert!(
            HeaderFooter::parse(&header_footer_record(0x000C, [2, 0, 0, 0], &strings)).is_err()
        );
        // Character count above the 255-unit limit.
        assert!(HeaderFooter::parse(&header_footer_record(0x000F, [256, 0, 0, 0], &[])).is_err());
        // Trailing garbage after the declared strings.
        assert!(HeaderFooter::parse(&header_footer_record(0x000C, [0, 0, 0, 0], &[0])).is_err());
        // Truncated string payload.
        assert!(
            HeaderFooter::parse(&header_footer_record(0x0001, [3, 0, 0, 0], &[0, b'E'])).is_err()
        );
    }

    #[test]
    fn header_footer_builder_validates_and_serializes() {
        let mut value = HeaderFooter::default();
        value
            .set_even("&LEven".to_string(), "&CFooter".to_string())
            .unwrap();
        value.set_first("First".to_string(), String::new()).unwrap();
        value.set_scale_with_doc(true);
        value.set_align_margins(true);
        let parsed = HeaderFooter::parse(&value.to_payload().unwrap()).unwrap();
        assert_eq!(parsed, value);

        assert!(value.set_even("x".repeat(256), String::new()).is_err());
    }
}
