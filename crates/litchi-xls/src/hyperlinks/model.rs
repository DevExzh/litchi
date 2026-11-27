//! Typed, inert values for BIFF8/MS-XLS hyperlinks.

pub const RECORD_TYPE: u16 = 0x01B8;
pub const TOOLTIP_RECORD_TYPE: u16 = 0x0800;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct HyperlinkRange {
    pub(crate) first_row: u16,
    pub(crate) last_row: u16,
    pub(crate) first_column: u8,
    pub(crate) last_column: u8,
}
impl HyperlinkRange {
    #[must_use]
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    #[must_use]
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    #[must_use]
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    #[must_use]
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UrlMoniker {
    pub(crate) url: String,
    pub(crate) serialization_uri_flags: Option<u32>,
}
impl UrlMoniker {
    #[must_use]
    pub fn url(&self) -> &str {
        &self.url
    }
    #[must_use]
    pub fn serialization_uri_flags(&self) -> Option<u32> {
        self.serialization_uri_flags
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileMoniker {
    pub(crate) parent_directory_count: u16,
    pub(crate) ansi_path: String,
    pub(crate) unicode_path: Option<String>,
    pub(crate) unc_server_character_count: Option<u16>,
}
impl FileMoniker {
    #[must_use]
    pub fn parent_directory_count(&self) -> u16 {
        self.parent_directory_count
    }
    #[must_use]
    pub fn ansi_path(&self) -> &str {
        &self.ansi_path
    }
    #[must_use]
    pub fn unicode_path(&self) -> Option<&str> {
        self.unicode_path.as_deref()
    }
    #[must_use]
    pub fn path(&self) -> &str {
        self.unicode_path.as_deref().unwrap_or(&self.ansi_path)
    }
    #[must_use]
    pub fn unc_server_character_count(&self) -> Option<u16> {
        self.unc_server_character_count
    }
    #[must_use]
    pub fn is_unc(&self) -> bool {
        self.unc_server_character_count.is_some()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ItemMoniker {
    pub(crate) delimiter_ansi: String,
    pub(crate) delimiter_unicode: Option<String>,
    pub(crate) item_ansi: String,
    pub(crate) item_unicode: Option<String>,
}
impl ItemMoniker {
    #[must_use]
    pub fn delimiter(&self) -> &str {
        self.delimiter_unicode
            .as_deref()
            .unwrap_or(&self.delimiter_ansi)
    }
    #[must_use]
    pub fn item(&self) -> &str {
        self.item_unicode.as_deref().unwrap_or(&self.item_ansi)
    }
    #[must_use]
    pub fn delimiter_ansi(&self) -> &str {
        &self.delimiter_ansi
    }
    #[must_use]
    pub fn item_ansi(&self) -> &str {
        &self.item_ansi
    }
}

/// Serialized moniker data retained without activation or resolution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum HyperlinkMoniker {
    String(String),
    Url(UrlMoniker),
    File(FileMoniker),
    Composite(Vec<HyperlinkMoniker>),
    Anti { count: u32 },
    Item(ItemMoniker),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyperlinkTargetKind {
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
pub struct Hyperlink {
    pub(crate) range: HyperlinkRange,
    pub(crate) class_id: [u8; 16],
    pub(crate) absolute: bool,
    pub(crate) site_gave_display_name: bool,
    pub(crate) absolute_from_relative: bool,
    pub(crate) display_name: Option<String>,
    pub(crate) target_frame: Option<String>,
    pub(crate) moniker: Option<HyperlinkMoniker>,
    pub(crate) location: Option<String>,
    pub(crate) hyperlink_guid: Option<[u8; 16]>,
    pub(crate) creation_time: Option<u64>,
    pub(crate) tooltip: Option<String>,
}
impl Hyperlink {
    #[must_use]
    pub fn range(&self) -> HyperlinkRange {
        self.range
    }
    #[must_use]
    pub fn class_id(&self) -> &[u8; 16] {
        &self.class_id
    }
    #[must_use]
    pub fn absolute(&self) -> bool {
        self.absolute
    }
    #[must_use]
    pub fn site_gave_display_name(&self) -> bool {
        self.site_gave_display_name
    }
    #[must_use]
    pub fn absolute_from_relative(&self) -> bool {
        self.absolute_from_relative
    }
    #[must_use]
    pub fn display_name(&self) -> Option<&str> {
        self.display_name.as_deref()
    }
    #[must_use]
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }
    #[must_use]
    pub fn moniker(&self) -> Option<&HyperlinkMoniker> {
        self.moniker.as_ref()
    }
    #[must_use]
    pub fn location(&self) -> Option<&str> {
        self.location.as_deref()
    }
    #[must_use]
    pub fn hyperlink_guid(&self) -> Option<&[u8; 16]> {
        self.hyperlink_guid.as_ref()
    }
    #[must_use]
    pub fn creation_time(&self) -> Option<u64> {
        self.creation_time
    }
    #[must_use]
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }
    #[must_use]
    pub fn target_kind(&self) -> HyperlinkTargetKind {
        match self.moniker.as_ref() {
            None => HyperlinkTargetKind::Document,
            Some(HyperlinkMoniker::Url(url))
                if starts_ascii_case_insensitive(url.url(), "mailto:") =>
            {
                HyperlinkTargetKind::Email
            },
            Some(HyperlinkMoniker::Url(_)) => HyperlinkTargetKind::Url,
            Some(HyperlinkMoniker::File(file)) if file.is_unc() => HyperlinkTargetKind::Unc,
            Some(HyperlinkMoniker::File(_)) => HyperlinkTargetKind::File,
            Some(HyperlinkMoniker::String(value)) if value.starts_with("\\\\") => {
                HyperlinkTargetKind::Unc
            },
            Some(HyperlinkMoniker::String(value))
                if starts_ascii_case_insensitive(value, "mailto:") =>
            {
                HyperlinkTargetKind::Email
            },
            Some(HyperlinkMoniker::String(_)) => HyperlinkTargetKind::StringMoniker,
            Some(HyperlinkMoniker::Composite(_)) => HyperlinkTargetKind::Composite,
            Some(HyperlinkMoniker::Anti { .. }) => HyperlinkTargetKind::Anti,
            Some(HyperlinkMoniker::Item(_)) => HyperlinkTargetKind::Item,
        }
    }
    /// Serialized base address, without filesystem or network resolution.
    #[must_use]
    pub fn address(&self) -> Option<&str> {
        match self.moniker.as_ref() {
            Some(HyperlinkMoniker::String(value)) => Some(value),
            Some(HyperlinkMoniker::Url(url)) => Some(url.url()),
            Some(HyperlinkMoniker::File(file)) => Some(file.path()),
            Some(HyperlinkMoniker::Item(item)) => Some(item.item()),
            Some(HyperlinkMoniker::Composite(_) | HyperlinkMoniker::Anti { .. }) => None,
            None => self.location(),
        }
    }
}

fn starts_ascii_case_insensitive(value: &str, prefix: &str) -> bool {
    value
        .get(..prefix.len())
        .is_some_and(|start| start.eq_ignore_ascii_case(prefix))
}
