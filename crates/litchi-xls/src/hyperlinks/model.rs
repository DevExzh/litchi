//! Typed, inert values for BIFF8/MS-XLS hyperlinks.

pub const RECORD_TYPE: u16 = 0x01B8;
pub const TOOLTIP_RECORD_TYPE: u16 = 0x0800;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsHyperlinkRange {
    pub(crate) first_row: u16,
    pub(crate) last_row: u16,
    pub(crate) first_column: u8,
    pub(crate) last_column: u8,
}
impl XlsHyperlinkRange {
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsUrlMoniker {
    pub(crate) url: String,
    pub(crate) serialization_uri_flags: Option<u32>,
}
impl XlsUrlMoniker {
    pub fn url(&self) -> &str {
        &self.url
    }
    pub fn serialization_uri_flags(&self) -> Option<u32> {
        self.serialization_uri_flags
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsFileMoniker {
    pub(crate) parent_directory_count: u16,
    pub(crate) ansi_path: String,
    pub(crate) unicode_path: Option<String>,
    pub(crate) unc_server_character_count: Option<u16>,
}
impl XlsFileMoniker {
    pub fn parent_directory_count(&self) -> u16 {
        self.parent_directory_count
    }
    pub fn ansi_path(&self) -> &str {
        &self.ansi_path
    }
    pub fn unicode_path(&self) -> Option<&str> {
        self.unicode_path.as_deref()
    }
    pub fn path(&self) -> &str {
        self.unicode_path.as_deref().unwrap_or(&self.ansi_path)
    }
    pub fn unc_server_character_count(&self) -> Option<u16> {
        self.unc_server_character_count
    }
    pub fn is_unc(&self) -> bool {
        self.unc_server_character_count.is_some()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsItemMoniker {
    pub(crate) delimiter_ansi: String,
    pub(crate) delimiter_unicode: Option<String>,
    pub(crate) item_ansi: String,
    pub(crate) item_unicode: Option<String>,
}
impl XlsItemMoniker {
    pub fn delimiter(&self) -> &str {
        self.delimiter_unicode
            .as_deref()
            .unwrap_or(&self.delimiter_ansi)
    }
    pub fn item(&self) -> &str {
        self.item_unicode.as_deref().unwrap_or(&self.item_ansi)
    }
    pub fn delimiter_ansi(&self) -> &str {
        &self.delimiter_ansi
    }
    pub fn item_ansi(&self) -> &str {
        &self.item_ansi
    }
}

/// Serialized moniker data retained without activation or resolution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsHyperlinkMoniker {
    String(String),
    Url(XlsUrlMoniker),
    File(XlsFileMoniker),
    Composite(Vec<XlsHyperlinkMoniker>),
    Anti { count: u32 },
    Item(XlsItemMoniker),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsHyperlinkTargetKind {
    Document,
    Url,
    Email,
    File,
    Unc,
    StringMoniker,
    Composite,
    Anti,
    Item,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsHyperlink {
    pub(crate) range: XlsHyperlinkRange,
    pub(crate) class_id: [u8; 16],
    pub(crate) absolute: bool,
    pub(crate) site_gave_display_name: bool,
    pub(crate) absolute_from_relative: bool,
    pub(crate) display_name: Option<String>,
    pub(crate) target_frame: Option<String>,
    pub(crate) moniker: Option<XlsHyperlinkMoniker>,
    pub(crate) location: Option<String>,
    pub(crate) hyperlink_guid: Option<[u8; 16]>,
    pub(crate) creation_time: Option<u64>,
    pub(crate) tooltip: Option<String>,
}
impl XlsHyperlink {
    pub fn range(&self) -> XlsHyperlinkRange {
        self.range
    }
    pub fn class_id(&self) -> &[u8; 16] {
        &self.class_id
    }
    pub fn absolute(&self) -> bool {
        self.absolute
    }
    pub fn site_gave_display_name(&self) -> bool {
        self.site_gave_display_name
    }
    pub fn absolute_from_relative(&self) -> bool {
        self.absolute_from_relative
    }
    pub fn display_name(&self) -> Option<&str> {
        self.display_name.as_deref()
    }
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }
    pub fn moniker(&self) -> Option<&XlsHyperlinkMoniker> {
        self.moniker.as_ref()
    }
    pub fn location(&self) -> Option<&str> {
        self.location.as_deref()
    }
    pub fn hyperlink_guid(&self) -> Option<&[u8; 16]> {
        self.hyperlink_guid.as_ref()
    }
    pub fn creation_time(&self) -> Option<u64> {
        self.creation_time
    }
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }
    pub fn target_kind(&self) -> XlsHyperlinkTargetKind {
        match self.moniker.as_ref() {
            None => XlsHyperlinkTargetKind::Document,
            Some(XlsHyperlinkMoniker::Url(url))
                if starts_ascii_case_insensitive(url.url(), "mailto:") =>
            {
                XlsHyperlinkTargetKind::Email
            },
            Some(XlsHyperlinkMoniker::Url(_)) => XlsHyperlinkTargetKind::Url,
            Some(XlsHyperlinkMoniker::File(file)) if file.is_unc() => XlsHyperlinkTargetKind::Unc,
            Some(XlsHyperlinkMoniker::File(_)) => XlsHyperlinkTargetKind::File,
            Some(XlsHyperlinkMoniker::String(value)) if value.starts_with("\\\\") => {
                XlsHyperlinkTargetKind::Unc
            },
            Some(XlsHyperlinkMoniker::String(value))
                if starts_ascii_case_insensitive(value, "mailto:") =>
            {
                XlsHyperlinkTargetKind::Email
            },
            Some(XlsHyperlinkMoniker::String(_)) => XlsHyperlinkTargetKind::StringMoniker,
            Some(XlsHyperlinkMoniker::Composite(_)) => XlsHyperlinkTargetKind::Composite,
            Some(XlsHyperlinkMoniker::Anti { .. }) => XlsHyperlinkTargetKind::Anti,
            Some(XlsHyperlinkMoniker::Item(_)) => XlsHyperlinkTargetKind::Item,
        }
    }
    /// Serialized base address, without filesystem or network resolution.
    pub fn address(&self) -> Option<&str> {
        match self.moniker.as_ref() {
            Some(XlsHyperlinkMoniker::String(value)) => Some(value),
            Some(XlsHyperlinkMoniker::Url(url)) => Some(url.url()),
            Some(XlsHyperlinkMoniker::File(file)) => Some(file.path()),
            Some(XlsHyperlinkMoniker::Item(item)) => Some(item.item()),
            Some(XlsHyperlinkMoniker::Composite(_) | XlsHyperlinkMoniker::Anti { .. }) => None,
            None => self.location(),
        }
    }
}

fn starts_ascii_case_insensitive(value: &str, prefix: &str) -> bool {
    value
        .get(..prefix.len())
        .is_some_and(|start| start.eq_ignore_ascii_case(prefix))
}
