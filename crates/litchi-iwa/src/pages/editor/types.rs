//! Public semantic value types used by the Pages editor.

use crate::protobuf::tp::DocumentArchive;
use crate::text::TextStorageInfo;

/// Which page variant owns a Pages header/footer storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesTemplateKind {
    First,
    Even,
    Odd,
}

/// Whether a reachable Pages text region is a header or a footer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PagesHeaderFooterKind {
    Header,
    Footer,
}

/// A reachable header/footer slot and its current writable text storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesHeaderFooterInfo {
    pub section_id: u64,
    pub section_name: Option<String>,
    /// UTF-16 position where the section begins in the body storage.
    pub section_character_index: u32,
    pub template_id: u64,
    pub template: PagesTemplateKind,
    pub kind: PagesHeaderFooterKind,
    /// Archive order within the header/footer list, normally left/center/right.
    pub slot: usize,
    pub storage: TextStorageInfo,
}

/// A writable text storage owned by a drawable reachable from a Pages document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesDrawableTextInfo {
    pub drawable_object_id: u64,
    pub storage: TextStorageInfo,
}

/// Result of removing a body-anchored Pages text box.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedPagesTextBox {
    pub text: PagesDrawableTextInfo,
    /// UTF-16 body position formerly occupied by the object-replacement character.
    pub anchor_character_index: u32,
}

/// A section boundary reachable from the main Pages body storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PagesSectionInfo {
    pub object_id: u64,
    /// UTF-16 position where the section begins in the body storage.
    pub character_index: u32,
    pub name: Option<String>,
    pub first_template_id: Option<u64>,
    pub even_template_id: Option<u64>,
    pub odd_template_id: Option<u64>,
}

/// Writable settings stored directly on a Pages section.
///
/// Numeric kinds remain raw so newer iWork values can round-trip without an
/// artificial enum rejecting them. `background_fill_payload`, when present,
/// is the exact encoded `TSD.FillArchive` payload.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PagesSectionSettings {
    pub name: Option<String>,
    pub inherit_previous_header_footer: Option<bool>,
    pub first_page_different: Option<bool>,
    pub even_odd_pages_different: Option<bool>,
    pub start_kind: Option<u32>,
    pub page_number_kind: Option<u32>,
    pub page_number_start: Option<u32>,
    pub first_page_hides_header_footer: Option<bool>,
    pub background_fill_payload: Option<Vec<u8>>,
}

/// RGB color space used by a semantic Pages section background.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PagesRgbColorSpace {
    Srgb,
    DisplayP3,
}

/// Normalized RGB color components in the inclusive `0.0..=1.0` range.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PagesRgbaColor {
    pub red: f32,
    pub green: f32,
    pub blue: f32,
    pub alpha: f32,
    pub color_space: PagesRgbColorSpace,
}

/// Semantic Pages section background.
///
/// Gradient, image, extension, and future fills are exposed as `Opaque` so
/// callers can round-trip them losslessly through the same API.
#[derive(Debug, Clone, PartialEq)]
pub enum PagesSectionBackground {
    None,
    Solid(PagesRgbaColor),
    Opaque(Vec<u8>),
}

/// Writable page geometry stored on the Pages document root.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesPageLayout {
    pub page_width: Option<f32>,
    pub page_height: Option<f32>,
    pub left_margin: Option<f32>,
    pub right_margin: Option<f32>,
    pub top_margin: Option<f32>,
    pub bottom_margin: Option<f32>,
    pub header_margin: Option<f32>,
    pub footer_margin: Option<f32>,
    pub page_scale: Option<f32>,
    /// Raw Pages orientation value; `0` is the default used by portrait files.
    pub orientation: Option<u32>,
    pub lays_out_body_vertically: Option<bool>,
}

impl From<&DocumentArchive> for PagesPageLayout {
    fn from(document: &DocumentArchive) -> Self {
        Self {
            page_width: document.page_width,
            page_height: document.page_height,
            left_margin: document.left_margin,
            right_margin: document.right_margin,
            top_margin: document.top_margin,
            bottom_margin: document.bottom_margin,
            header_margin: document.header_margin,
            footer_margin: document.footer_margin,
            page_scale: document.page_scale,
            orientation: document.orientation,
            lays_out_body_vertically: document.lays_out_body_vertically,
        }
    }
}
