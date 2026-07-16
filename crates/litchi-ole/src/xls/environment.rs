//! Typed BIFF8 workbook-global environment and behavioral options.

use super::{XlsError, XlsResult};

pub(crate) const BACKUP_RECORD_TYPE: u16 = 0x0040;
pub(crate) const TEMPLATE_RECORD_TYPE: u16 = 0x0060;
pub(crate) const COUNTRY_RECORD_TYPE: u16 = 0x008C;
pub(crate) const HIDE_OBJ_RECORD_TYPE: u16 = 0x008D;
pub(crate) const BOOK_BOOL_RECORD_TYPE: u16 = 0x00DA;
pub(crate) const USES_ELFS_RECORD_TYPE: u16 = 0x0160;
pub(crate) const DSF_RECORD_TYPE: u16 = 0x0161;
pub(crate) const REFRESH_ALL_RECORD_TYPE: u16 = 0x01B7;
pub(crate) const EXCEL9_FILE_RECORD_TYPE: u16 = 0x01C0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsObjectDisplayMode { ShowAll, ShowPlaceholders, HideAll }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsLinkUpdateMode { Prompt, Never, Silent }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsWorkbookEnvironment {
    template: bool,
    has_biff5_stream: bool,
    excel9_file_marker: bool,
    create_backup_copy: bool,
    object_display_mode: XlsObjectDisplayMode,
    refresh_external_data_on_load: bool,
    save_external_link_values: bool,
    has_envelope: bool,
    envelope_visible: bool,
    envelope_initialized: bool,
    link_update_mode: XlsLinkUpdateMode,
    hide_unselected_table_borders: bool,
    supports_natural_language_formulas: bool,
    default_country_code: u16,
    current_country_code: u16,
}

impl Default for XlsWorkbookEnvironment {
    fn default() -> Self {
        Self {
            template: false, has_biff5_stream: false,
            excel9_file_marker: false, create_backup_copy: false,
            object_display_mode: XlsObjectDisplayMode::ShowAll,
            refresh_external_data_on_load: false, save_external_link_values: true,
            has_envelope: false, envelope_visible: false, envelope_initialized: false,
            link_update_mode: XlsLinkUpdateMode::Prompt,
            hide_unselected_table_borders: false,
            supports_natural_language_formulas: false,
            default_country_code: 1, current_country_code: 1,
        }
    }
}

impl XlsWorkbookEnvironment {
    pub fn is_template(&self) -> bool { self.template }
    pub fn has_biff5_stream(&self) -> bool { self.has_biff5_stream }
    pub fn has_excel9_file_marker(&self) -> bool { self.excel9_file_marker }
    pub fn create_backup_copy(&self) -> bool { self.create_backup_copy }
    pub fn object_display_mode(&self) -> XlsObjectDisplayMode { self.object_display_mode }
    /// Metadata only: the reader never refreshes or opens external data.
    pub fn refresh_external_data_on_load(&self) -> bool { self.refresh_external_data_on_load }
    pub fn save_external_link_values(&self) -> bool { self.save_external_link_values }
    pub fn has_envelope(&self) -> bool { self.has_envelope }
    pub fn envelope_visible(&self) -> bool { self.envelope_visible }
    pub fn envelope_initialized(&self) -> bool { self.envelope_initialized }
    pub fn link_update_mode(&self) -> XlsLinkUpdateMode { self.link_update_mode }
    pub fn hide_unselected_table_borders(&self) -> bool { self.hide_unselected_table_borders }
    pub fn supports_natural_language_formulas(&self) -> bool { self.supports_natural_language_formulas }
    pub fn default_country_code(&self) -> u16 { self.default_country_code }
    pub fn current_country_code(&self) -> u16 { self.current_country_code }
}

pub(crate) struct EnvironmentCollector {
    value: XlsWorkbookEnvironment,
    seen: u16,
    last_rank: Option<u8>,
}

impl EnvironmentCollector {
    pub(crate) fn new() -> Self {
        Self { value: XlsWorkbookEnvironment::default(), seen: 0, last_rank: None }
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        let (rank, bit) = match record_type {
            TEMPLATE_RECORD_TYPE => (0, 1 << 0), DSF_RECORD_TYPE => (1, 1 << 1),
            EXCEL9_FILE_RECORD_TYPE => (2, 1 << 2), BACKUP_RECORD_TYPE => (3, 1 << 3),
            HIDE_OBJ_RECORD_TYPE => (4, 1 << 4), REFRESH_ALL_RECORD_TYPE => (5, 1 << 5),
            BOOK_BOOL_RECORD_TYPE => (6, 1 << 6), USES_ELFS_RECORD_TYPE => (7, 1 << 7),
            COUNTRY_RECORD_TYPE => (8, 1 << 8), _ => return Ok(()),
        };
        if self.seen & bit != 0 { return invalid(record_type, "duplicate workbook environment record"); }
        if self.last_rank.is_some_and(|previous| rank < previous) {
            return invalid(record_type, "workbook environment record is out of BIFF8 order");
        }
        self.seen |= bit;
        self.last_rank = Some(rank);
        match record_type {
            TEMPLATE_RECORD_TYPE => {
                require_length(record_type, data, 0)?;
                self.value.template = true;
            },
            DSF_RECORD_TYPE => {
                self.value.has_biff5_stream = parse_bool(record_type, data)?;
            },
            EXCEL9_FILE_RECORD_TYPE => {
                require_length(record_type, data, 0)?;
                self.value.excel9_file_marker = true;
            },
            BACKUP_RECORD_TYPE => self.value.create_backup_copy = parse_bool(record_type, data)?,
            HIDE_OBJ_RECORD_TYPE => {
                require_length(record_type, data, 2)?;
                self.value.object_display_mode = match read_u16(data, 0) {
                    0 => XlsObjectDisplayMode::ShowAll,
                    1 => XlsObjectDisplayMode::ShowPlaceholders,
                    2 => XlsObjectDisplayMode::HideAll,
                    value => return invalid(record_type, format!("invalid HideObj value {value}")),
                };
            },
            REFRESH_ALL_RECORD_TYPE => self.value.refresh_external_data_on_load = parse_bool(record_type, data)?,
            BOOK_BOOL_RECORD_TYPE => self.parse_book_bool(data)?,
            USES_ELFS_RECORD_TYPE => self.value.supports_natural_language_formulas = parse_bool(record_type, data)?,
            COUNTRY_RECORD_TYPE => {
                require_length(record_type, data, 4)?;
                let default = read_u16(data, 0);
                let current = read_u16(data, 2);
                if !(1..=981).contains(&default) || !(1..=981).contains(&current) {
                    return invalid(record_type, "Country codes must be 1..=981");
                }
                self.value.default_country_code = default;
                self.value.current_country_code = current;
            },
            _ => unreachable!(),
        }
        Ok(())
    }
    fn parse_book_bool(&mut self, data: &[u8]) -> XlsResult<()> {
        require_length(BOOK_BOOL_RECORD_TYPE, data, 2)?;
        let bits = read_u16(data, 0);
        if bits & 0xFE02 != 0 { return invalid(BOOK_BOOL_RECORD_TYPE, "BookBool contains reserved bits"); }
        let has_envelope = bits & 0x0004 != 0;
        let visible = bits & 0x0008 != 0;
        let initialized = bits & 0x0010 != 0;
        if (visible || initialized) && !has_envelope {
            return invalid(BOOK_BOOL_RECORD_TYPE, "BookBool envelope flags require fHasEnvelope");
        }
        self.value.save_external_link_values = bits & 1 == 0;
        self.value.has_envelope = has_envelope;
        self.value.envelope_visible = visible;
        self.value.envelope_initialized = initialized;
        self.value.link_update_mode = match (bits >> 5) & 3 {
            0 => XlsLinkUpdateMode::Prompt, 1 => XlsLinkUpdateMode::Never,
            2 => XlsLinkUpdateMode::Silent,
            _ => return invalid(BOOK_BOOL_RECORD_TYPE, "invalid BookBool link update mode"),
        };
        self.value.hide_unselected_table_borders = bits & 0x0100 != 0;
        Ok(())
    }
    pub(crate) fn finish(self) -> XlsResult<XlsWorkbookEnvironment> {
        if self.value.refresh_external_data_on_load && !self.value.template {
            return invalid(REFRESH_ALL_RECORD_TYPE, "RefreshAll must be zero for a non-template workbook");
        }
        Ok(self.value)
    }
}

fn parse_bool(record_type: u16, data: &[u8]) -> XlsResult<bool> {
    require_length(record_type, data, 2)?;
    match read_u16(data, 0) { 0 => Ok(false), 1 => Ok(true), value => invalid(record_type, format!("Boolean must be 0 or 1, got {value}")) }
}
fn require_length(record_type: u16, data: &[u8], expected: usize) -> XlsResult<()> {
    if data.len() != expected { return invalid(record_type, format!("payload must be exactly {expected} bytes, got {}", data.len())); }
    Ok(())
}
fn read_u16(data: &[u8], offset: usize) -> u16 { u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap()) }
fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord { record_type, message: message.into() })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn rejects_bad_lengths_reserved_bits_values_and_order() {
        let mut collector = EnvironmentCollector::new();
        assert!(collector.feed_record(DSF_RECORD_TYPE, &[2, 0]).is_err());
        let mut collector = EnvironmentCollector::new();
        assert!(collector.feed_record(HIDE_OBJ_RECORD_TYPE, &[3, 0]).is_err());
        let mut collector = EnvironmentCollector::new();
        assert!(collector.feed_record(BOOK_BOOL_RECORD_TYPE, &[2, 0]).is_err());
        let mut collector = EnvironmentCollector::new();
        collector.feed_record(COUNTRY_RECORD_TYPE, &[1, 0, 1, 0]).unwrap();
        assert!(collector.feed_record(BACKUP_RECORD_TYPE, &[0, 0]).is_err());
    }
    #[test]
    fn rejects_refresh_for_non_template() {
        let mut collector = EnvironmentCollector::new();
        collector.feed_record(REFRESH_ALL_RECORD_TYPE, &[1, 0]).unwrap();
        assert!(collector.finish().is_err());
    }
    #[test]
    fn accepts_dual_stream_dsf_value() {
        let mut collector = EnvironmentCollector::new();
        collector.feed_record(DSF_RECORD_TYPE, &[1, 0]).unwrap();
        assert!(collector.finish().unwrap().has_biff5_stream());
    }
}
