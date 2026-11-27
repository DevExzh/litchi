//! Chartsheet metadata children from `SpreadsheetML`.

#[derive(Debug, Clone, PartialEq)]
pub struct Color {
    pub automatic: Option<bool>,
    pub indexed: Option<u32>,
    pub rgb: Option<String>,
    pub theme: Option<u32>,
    pub tint: Option<f64>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Properties {
    pub published: Option<bool>,
    pub code_name: Option<String>,
    pub tab_color: Option<Color>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Protection {
    pub password_hash: Option<String>,
    pub content: Option<bool>,
    pub objects: Option<bool>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Margins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageSetup {
    pub paper_size: Option<u32>,
    pub first_page_number: Option<u32>,
    pub orientation: Option<super::PageOrientation>,
    pub use_printer_defaults: Option<bool>,
    pub black_and_white: Option<bool>,
    pub draft: Option<bool>,
    pub use_first_page_number: Option<bool>,
    pub horizontal_dpi: Option<u32>,
    pub vertical_dpi: Option<u32>,
    pub copies: Option<u32>,
    /// Inert relationship reference to a binary Printer Settings part.
    pub printer_settings_relationship_id: Option<String>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooter {
    pub different_odd_even: Option<bool>,
    pub different_first: Option<bool>,
    pub scale_with_document: Option<bool>,
    pub align_with_margins: Option<bool>,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
}
