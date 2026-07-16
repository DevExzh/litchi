//! BIFF8 worksheet print and page setup records.

use crate::xls::error::{XlsError, XlsResult};

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
const USER_SVIEW_BEGIN_RECORD_TYPE: u16 = 0x01aa;
const USER_SVIEW_END_RECORD_TYPE: u16 = 0x01ab;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
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
pub enum XlsPrintOrder {
    DownThenOver,
    OverThenDown,
}

/// Explicit paper orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsPrintOrientation {
    Landscape,
    Portrait,
}

/// Representation used for cell errors when printing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsPrintErrors {
    Displayed,
    Blank,
    Dashes,
    NotAvailable,
}

/// How cell comments are printed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsPrintComments {
    None,
    AsDisplayed,
    AtEnd,
}

/// One explicit page break and the inclusive perpendicular page span.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsPageBreak {
    position: u16,
    range_start: u16,
    range_end: u16,
}

impl XlsPageBreak {
    /// First row below a horizontal break, or first column right of a vertical break.
    pub const fn position(&self) -> u16 { self.position }
    pub const fn range_start(&self) -> u16 { self.range_start }
    pub const fn range_end(&self) -> u16 { self.range_end }
}

/// Fixed `SETUP` print configuration.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsPrintSetup {
    printer_settings_available: bool,
    paper_size: Option<u16>,
    scale_percent: Option<u16>,
    starting_page_number: Option<i16>,
    fit_width_pages: u16,
    fit_height_pages: u16,
    print_order: XlsPrintOrder,
    orientation: Option<XlsPrintOrientation>,
    black_and_white: bool,
    draft_quality: bool,
    comments: XlsPrintComments,
    errors: XlsPrintErrors,
    horizontal_resolution_dpi: Option<u16>,
    vertical_resolution_dpi: Option<u16>,
    header_margin_inches: f64,
    footer_margin_inches: f64,
    copies: Option<u16>,
}

impl XlsPrintSetup {
    pub fn printer_settings_available(&self) -> bool { self.printer_settings_available }
    pub fn paper_size(&self) -> Option<u16> { self.paper_size }
    pub fn scale_percent(&self) -> Option<u16> { self.scale_percent }
    pub fn starting_page_number(&self) -> Option<i16> { self.starting_page_number }
    pub fn fit_width_pages(&self) -> u16 { self.fit_width_pages }
    pub fn fit_height_pages(&self) -> u16 { self.fit_height_pages }
    pub fn print_order(&self) -> XlsPrintOrder { self.print_order }
    pub fn orientation(&self) -> Option<XlsPrintOrientation> { self.orientation }
    pub fn is_black_and_white(&self) -> bool { self.black_and_white }
    pub fn is_draft_quality(&self) -> bool { self.draft_quality }
    pub fn comments(&self) -> XlsPrintComments { self.comments }
    pub fn errors(&self) -> XlsPrintErrors { self.errors }
    pub fn horizontal_resolution_dpi(&self) -> Option<u16> { self.horizontal_resolution_dpi }
    pub fn vertical_resolution_dpi(&self) -> Option<u16> { self.vertical_resolution_dpi }
    pub fn header_margin_inches(&self) -> f64 { self.header_margin_inches }
    pub fn footer_margin_inches(&self) -> f64 { self.footer_margin_inches }
    pub fn copies(&self) -> Option<u16> { self.copies }
}

impl Default for XlsPrintSetup {
    fn default() -> Self {
        Self {
            printer_settings_available: false,
            paper_size: None,
            scale_percent: None,
            starting_page_number: None,
            fit_width_pages: 0,
            fit_height_pages: 0,
            print_order: XlsPrintOrder::DownThenOver,
            orientation: None,
            black_and_white: false,
            draft_quality: false,
            comments: XlsPrintComments::None,
            errors: XlsPrintErrors::Displayed,
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
pub struct XlsPageSetup {
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
    horizontal_page_breaks: Vec<XlsPageBreak>,
    vertical_page_breaks: Vec<XlsPageBreak>,
    printer_driver_data: Vec<Vec<u8>>,
    print_setup: XlsPrintSetup,
}

impl XlsPageSetup {
    pub const fn print_headers(&self) -> bool { self.print_headers }
    pub const fn print_gridlines(&self) -> bool { self.print_gridlines }
    /// Raw header text, including `&L`, `&C`, `&R`, and formatting commands.
    pub fn header(&self) -> &str { &self.header }
    /// Raw footer text, including `&L`, `&C`, `&R`, and formatting commands.
    pub fn footer(&self) -> &str { &self.footer }
    pub fn is_horizontally_centered(&self) -> bool { self.horizontally_centered }
    pub fn is_vertically_centered(&self) -> bool { self.vertically_centered }
    pub fn left_margin_inches(&self) -> Option<f64> { self.left_margin_inches }
    pub fn right_margin_inches(&self) -> Option<f64> { self.right_margin_inches }
    pub fn top_margin_inches(&self) -> Option<f64> { self.top_margin_inches }
    pub fn bottom_margin_inches(&self) -> Option<f64> { self.bottom_margin_inches }
    pub fn horizontal_page_breaks(&self) -> &[XlsPageBreak] { &self.horizontal_page_breaks }
    pub fn vertical_page_breaks(&self) -> &[XlsPageBreak] { &self.vertical_page_breaks }
    /// Opaque DEVMODE payloads from `PLS`; these bytes are never executed.
    pub fn printer_driver_data(&self) -> &[Vec<u8>] { &self.printer_driver_data }
    pub fn print_setup(&self) -> &XlsPrintSetup { &self.print_setup }
}

fn parse_header_footer(data: &[u8], record_type: u16) -> XlsResult<String> {
    if data.is_empty() {
        return Ok(String::new());
    }
    if data.len() < 3 {
        return Err(invalid(record_type, "header/footer string has a truncated XLUnicodeString header"));
    }
    let count = usize::from(read_u16(data, 0));
    let flags = data[2];
    if count > 255 {
        return Err(invalid(record_type, "header/footer text exceeds 255 UTF-16 code units"));
    }
    if flags & 0xfe != 0 {
        return Err(invalid(record_type, "header/footer XLUnicodeString has reserved option bits"));
    }
    let width = if flags & 1 != 0 { 2 } else { 1 };
    let byte_count = count
        .checked_mul(width)
        .ok_or_else(|| invalid(record_type, "header/footer string length overflow"))?;
    if data.len() != 3 + byte_count {
        return Err(invalid(record_type, "header/footer character count does not match its payload"));
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

fn parse_bool(data: &[u8], record_type: u16) -> XlsResult<bool> {
    if data.len() != 2 {
        return Err(invalid(record_type, format!("Boolean payload must be 2 bytes, found {}", data.len())));
    }
    match read_u16(data, 0) {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(invalid(record_type, "Boolean value must be zero or one")),
    }
}

fn parse_margin(data: &[u8], record_type: u16) -> XlsResult<f64> {
    if data.len() != 8 {
        return Err(invalid(record_type, format!("margin payload must be 8 bytes, found {}", data.len())));
    }
    let margin = read_f64(data, 0);
    if !margin.is_finite() || !(0.0..=49.0).contains(&margin) {
        return Err(invalid(record_type, "page margin must be finite and between 0 and 49 inches"));
    }
    Ok(margin)
}

fn parse_page_breaks(data: &[u8], record_type: u16) -> XlsResult<Vec<XlsPageBreak>> {
    if data.len() < 2 {
        return Err(invalid(record_type, "page-break record is missing its count"));
    }
    let count = usize::from(read_u16(data, 0));
    let maximum = if record_type == HORIZONTAL_PAGE_BREAKS_RECORD_TYPE { 1026 } else { 255 };
    if count > maximum {
        return Err(invalid(record_type, format!("page-break count exceeds {maximum}")));
    }
    let expected = 2 + count * 6;
    if data.len() != expected {
        return Err(invalid(record_type, format!("page-break count requires {expected} bytes, found {}", data.len())));
    }
    let mut breaks: Vec<XlsPageBreak> = Vec::with_capacity(count);
    for chunk in data[2..].chunks_exact(6) {
        let page_break = XlsPageBreak {
            position: read_u16(chunk, 0),
            range_start: read_u16(chunk, 2),
            range_end: read_u16(chunk, 4),
        };
        if page_break.range_end <= page_break.range_start {
            return Err(invalid(record_type, "page-break range end must be greater than its start"));
        }
        if record_type == HORIZONTAL_PAGE_BREAKS_RECORD_TYPE && page_break.range_end > 16383 {
            return Err(invalid(record_type, "horizontal page-break column exceeds 16383"));
        }
        if record_type == VERTICAL_PAGE_BREAKS_RECORD_TYPE && page_break.position > 255 {
            return Err(invalid(record_type, "vertical page-break column exceeds 255"));
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

fn parse_pls(data: &[u8]) -> XlsResult<Vec<u8>> {
    if data.len() < 2 {
        return Err(invalid(PLS_RECORD_TYPE, "PLS is missing its reserved field"));
    }
    if read_u16(data, 0) != 0 {
        return Err(invalid(PLS_RECORD_TYPE, "PLS reserved field must be zero"));
    }
    Ok(data[2..].to_vec())
}

fn parse_setup(data: &[u8]) -> XlsResult<XlsPrintSetup> {
    if data.len() != 34 {
        return Err(invalid(SETUP_RECORD_TYPE, format!("SETUP payload must be 34 bytes, found {}", data.len())));
    }
    let paper_size = read_u16(data, 0);
    let flags = read_u16(data, 10);
    let fit_width_pages = read_u16(data, 6);
    let fit_height_pages = read_u16(data, 8);
    let no_printer_settings = flags & 0x0004 != 0;
    if flags & 0xf000 != 0 {
        return Err(invalid(SETUP_RECORD_TYPE, "SETUP reserved flag bits must be zero"));
    }
    if fit_width_pages > 32767 || fit_height_pages > 32767 {
        return Err(invalid(SETUP_RECORD_TYPE, "SETUP fit dimensions must not exceed 32767 pages"));
    }
    if !no_printer_settings && (118..=255).contains(&paper_size) {
        return Err(invalid(SETUP_RECORD_TYPE, "SETUP paper size uses a reserved identifier"));
    }
    let header_margin_inches = read_f64(data, 16);
    let footer_margin_inches = read_f64(data, 24);
    if !header_margin_inches.is_finite()
        || !(0.0..49.0).contains(&header_margin_inches)
        || !footer_margin_inches.is_finite()
        || !(0.0..49.0).contains(&footer_margin_inches)
    {
        return Err(invalid(SETUP_RECORD_TYPE, "SETUP header/footer margins must be finite and less than 49 inches"));
    }
    let printer_value = |value| (!no_printer_settings).then_some(value);
    let orientation = if no_printer_settings || flags & 0x0040 != 0 {
        None
    } else if flags & 0x0002 != 0 {
        Some(XlsPrintOrientation::Portrait)
    } else {
        Some(XlsPrintOrientation::Landscape)
    };
    let comments = if flags & 0x0020 == 0 {
        XlsPrintComments::None
    } else if flags & 0x0200 != 0 {
        XlsPrintComments::AtEnd
    } else {
        XlsPrintComments::AsDisplayed
    };
    let errors = match (flags >> 10) & 3 {
        0 => XlsPrintErrors::Displayed,
        1 => XlsPrintErrors::Blank,
        2 => XlsPrintErrors::Dashes,
        _ => XlsPrintErrors::NotAvailable,
    };

    Ok(XlsPrintSetup {
        printer_settings_available: !no_printer_settings,
        paper_size: printer_value(paper_size),
        scale_percent: printer_value(read_u16(data, 2)),
        starting_page_number: (flags & 0x0080 != 0).then_some(read_u16(data, 4) as i16),
        fit_width_pages,
        fit_height_pages,
        print_order: if flags & 0x0001 != 0 { XlsPrintOrder::OverThenDown } else { XlsPrintOrder::DownThenOver },
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
    horizontal_page_breaks: Option<Vec<XlsPageBreak>>,
    vertical_page_breaks: Option<Vec<XlsPageBreak>>,
    printer_driver_data: Vec<Vec<u8>>,
    print_setup: Option<XlsPrintSetup>,
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

    fn duplicate(record_type: u16) -> XlsError {
        invalid(record_type, "worksheet contains a duplicate primary page-setup record")
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if self.collecting_pls {
            if record_type == CONTINUE_RECORD_TYPE {
                self.page.printer_driver_data.last_mut().unwrap().extend_from_slice(data);
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
                if self.page.print_headers.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.print_headers = Some(parse_bool(data, record_type)?);
            }
            PRINT_GRIDLINES_RECORD_TYPE => {
                if self.page.print_gridlines.is_some() { return Err(Self::duplicate(record_type)); }
                if data.len() != 2 {
                    return Err(invalid(record_type, "PRINTGRIDLINES payload must be 2 bytes"));
                }
                self.page.print_gridlines = Some(read_u16(data, 0) & 1 != 0);
            }
            HORIZONTAL_PAGE_BREAKS_RECORD_TYPE => {
                if self.page.horizontal_page_breaks.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.horizontal_page_breaks = Some(parse_page_breaks(data, record_type)?);
            }
            VERTICAL_PAGE_BREAKS_RECORD_TYPE => {
                if self.page.vertical_page_breaks.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.vertical_page_breaks = Some(parse_page_breaks(data, record_type)?);
            }
            HEADER_RECORD_TYPE => {
                if self.page.header.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.header = Some(parse_header_footer(data, record_type)?);
            }
            FOOTER_RECORD_TYPE => {
                if self.page.footer.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.footer = Some(parse_header_footer(data, record_type)?);
            }
            HCENTER_RECORD_TYPE => {
                if self.page.horizontally_centered.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.horizontally_centered = Some(parse_bool(data, record_type)?);
            }
            VCENTER_RECORD_TYPE => {
                if self.page.vertically_centered.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.vertically_centered = Some(parse_bool(data, record_type)?);
            }
            LEFT_MARGIN_RECORD_TYPE => {
                if self.page.left_margin_inches.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.left_margin_inches = Some(parse_margin(data, record_type)?);
            }
            RIGHT_MARGIN_RECORD_TYPE => {
                if self.page.right_margin_inches.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.right_margin_inches = Some(parse_margin(data, record_type)?);
            }
            TOP_MARGIN_RECORD_TYPE => {
                if self.page.top_margin_inches.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.top_margin_inches = Some(parse_margin(data, record_type)?);
            }
            BOTTOM_MARGIN_RECORD_TYPE => {
                if self.page.bottom_margin_inches.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.bottom_margin_inches = Some(parse_margin(data, record_type)?);
            }
            PLS_RECORD_TYPE => {
                self.page.printer_driver_data.push(parse_pls(data)?);
                self.collecting_pls = true;
            }
            SETUP_RECORD_TYPE => {
                if self.page.print_setup.is_some() { return Err(Self::duplicate(record_type)); }
                self.page.print_setup = Some(parse_setup(data)?);
            }
            _ => {}
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> XlsResult<Option<XlsPageSetup>> {
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
                || self.page.bottom_margin_inches.is_some();
        if !has_partial && self.page.print_setup.is_none() {
            return Ok(None);
        }
        Ok(Some(XlsPageSetup {
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
        collector.feed_record(HEADER_RECORD_TYPE, &compressed("&Ctitle")).unwrap();
        collector.feed_record(FOOTER_RECORD_TYPE, &[]).unwrap();
        collector.feed_record(HCENTER_RECORD_TYPE, &[1, 0]).unwrap();
        collector.feed_record(VCENTER_RECORD_TYPE, &[0, 0]).unwrap();
        collector.feed_record(LEFT_MARGIN_RECORD_TYPE, &0.7f64.to_le_bytes()).unwrap();
        collector.feed_record(SETUP_RECORD_TYPE, &setup(0x0002)).unwrap();
        let page = collector.finish().unwrap().unwrap();
        assert_eq!(page.header(), "&Ctitle");
        assert!(page.is_horizontally_centered());
        assert_eq!(page.print_setup().orientation(), Some(XlsPrintOrientation::Portrait));
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

        assert!(parse_page_breaks(&[1, 0, 4, 0, 2, 0, 2, 0], HORIZONTAL_PAGE_BREAKS_RECORD_TYPE).is_err());
        let overlapping = [2, 0, 4, 0, 0, 0, 10, 0, 4, 0, 9, 0, 12, 0];
        assert!(parse_page_breaks(&overlapping, HORIZONTAL_PAGE_BREAKS_RECORD_TYPE).is_err());
        assert!(parse_pls(&[1, 0]).is_err());
    }

    #[test]
    fn partial_page_block_uses_tolerant_defaults() {
        let mut collector = PageSetupCollector::new();
        collector.feed_record(PRINT_HEADERS_RECORD_TYPE, &[1, 0]).unwrap();
        let page = collector.finish().unwrap().unwrap();
        assert!(page.print_headers());
        assert!(!page.print_gridlines());
        assert!(!page.print_setup().printer_settings_available());
    }

    #[test]
    fn ignores_custom_view_page_records() {
        let mut collector = PageSetupCollector::new();
        collector.feed_record(HEADER_RECORD_TYPE, &compressed("primary")).unwrap();
        collector.feed_record(USER_SVIEW_BEGIN_RECORD_TYPE, &[]).unwrap();
        collector.feed_record(HEADER_RECORD_TYPE, &compressed("custom")).unwrap();
        collector.feed_record(USER_SVIEW_END_RECORD_TYPE, &[]).unwrap();
        collector.feed_record(SETUP_RECORD_TYPE, &setup(0x0002)).unwrap();
        assert_eq!(collector.finish().unwrap().unwrap().header(), "primary");

        let mut omitted = PageSetupCollector::new();
        omitted.feed_record(SETUP_RECORD_TYPE, &setup(0x0002)).unwrap();
        let page = omitted.finish().unwrap().unwrap();
        assert_eq!(page.header(), "");
        assert_eq!(page.footer(), "");
    }

    #[test]
    fn reads_poi_page_setup_fixtures() {
        use crate::xls::XlsWorkbook;
        use std::fs::File;
        use std::path::Path;

        let fixture = |name: &str| {
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../3rdparty/poi/test-data/spreadsheet")
                .join(name)
        };

        let dbcs = XlsWorkbook::new(File::open(fixture("DBCSHeader.xls")).unwrap()).unwrap();
        let page = dbcs.xls_worksheet(0).unwrap().page_setup().unwrap();
        assert_eq!(page.header(), "&L\u{090f}\u{0915}&C\u{0939}\u{094b}\u{0917}\u{093e}&R\u{091c}\u{093e}");
        assert_eq!(page.footer(), "&L\u{091c}\u{093e}&C\u{091c}\u{093e}&R\u{091c}\u{093e}");
        assert_eq!(page.print_setup().orientation(), Some(XlsPrintOrientation::Portrait));
        assert_eq!(page.print_setup().paper_size(), Some(1));

        let breaks = XlsWorkbook::new(
            File::open(fixture("SimpleWithPageBreaks.xls")).unwrap(),
        )
        .unwrap();
        let page = breaks.xls_worksheet(0).unwrap().page_setup().unwrap();
        assert!(
            !page.horizontal_page_breaks().is_empty()
                || !page.vertical_page_breaks().is_empty()
        );

    }
}
