/// Passive file and template preferences from document-formatting flags.
///
/// These values record requested metadata only. They never create backups,
/// save files, or change the document's storage format.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentFileSettings {
    pub automatic_backup: bool,
    pub default_save_format_rtf: bool,
    pub template_or_stationery: bool,
}

impl DocumentFileSettings {
    pub fn is_empty(self) -> bool {
        self == Self::default()
    }
}
