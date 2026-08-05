//! Typed SpreadsheetML chartsheet semantic model.

use crate::error::{Error, Result};

pub(super) const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(super) const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
pub(super) const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
pub(super) const CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub(super) const STRICT_CHART: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
pub(super) const DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
pub(super) const STRICT_DRAWING_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing";
pub(super) const CHART_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
pub(super) const STRICT_CHART_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chart";
pub(super) const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
pub(super) const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
pub(super) const IMAGE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub(super) const STRICT_IMAGE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
pub(super) const CHART_USER_SHAPES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
pub(super) const STRICT_CHART_USER_SHAPES_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes";
pub(super) const THEME_OVERRIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";
pub(super) const STRICT_THEME_OVERRIDE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride";
pub(super) const PACKAGE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
pub(super) const STRICT_PACKAGE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/package";
pub(super) const VML_DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
pub(super) const STRICT_VML_DRAWING_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/vmlDrawing";
pub(super) const PRINTER_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
pub(super) const STRICT_PRINTER_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub fn sml(self) -> &'static str {
        match self {
            Self::Transitional => SML,
            Self::Strict => STRICT_SML,
        }
    }
    pub fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }
    pub fn xdr(self) -> &'static str {
        match self {
            Self::Transitional => XDR,
            Self::Strict => STRICT_XDR,
        }
    }
    pub fn chart(self) -> &'static str {
        match self {
            Self::Transitional => CHART,
            Self::Strict => STRICT_CHART,
        }
    }
    pub fn chartsheet_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHARTSHEET_REL,
            Self::Strict => STRICT_CHARTSHEET_REL,
        }
    }
    pub fn drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => DRAWING_REL,
            Self::Strict => STRICT_DRAWING_REL,
        }
    }
    pub fn chart_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHART_REL,
            Self::Strict => STRICT_CHART_REL,
        }
    }
    pub fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => IMAGE_REL,
            Self::Strict => STRICT_IMAGE_REL,
        }
    }
    pub fn chart_user_shapes_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHART_USER_SHAPES_REL,
            Self::Strict => STRICT_CHART_USER_SHAPES_REL,
        }
    }
    pub fn theme_override_rel(self) -> &'static str {
        match self {
            Self::Transitional => THEME_OVERRIDE_REL,
            Self::Strict => STRICT_THEME_OVERRIDE_REL,
        }
    }
    pub fn package_rel(self) -> &'static str {
        match self {
            Self::Transitional => PACKAGE_REL,
            Self::Strict => STRICT_PACKAGE_REL,
        }
    }
    pub fn vml_drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => VML_DRAWING_REL,
            Self::Strict => STRICT_VML_DRAWING_REL,
        }
    }
    pub fn printer_rel(self) -> &'static str {
        match self {
            Self::Transitional => PRINTER_REL,
            Self::Strict => STRICT_PRINTER_REL,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum State {
    Visible,
    Hidden,
    VeryHidden,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Default,
    Portrait,
    Landscape,
}

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
pub struct View {
    pub tab_selected: Option<bool>,
    pub zoom_scale: Option<u32>,
    pub workbook_view_id: u32,
    pub zoom_to_fit: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Protection {
    pub password_hash: Option<String>,
    pub content: Option<bool>,
    pub objects: Option<bool>,
}

/// One saved chartsheet view from `CT_CustomChartsheetView`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomView {
    /// Braced UUID lexical form required by SpreadsheetML `ST_Guid`.
    pub guid: String,
    pub scale: Option<u32>,
    pub state: Option<State>,
    pub zoom_to_fit: Option<bool>,
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
    pub orientation: Option<PageOrientation>,
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

/// Schema-complete `ST_WebSourceType` values for inert web-publishing metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebSourceType {
    Sheet,
    PrintArea,
    AutoFilter,
    Range,
    Chart,
    PivotTable,
    Query,
    Label,
}

impl WebSourceType {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Sheet => "sheet",
            Self::PrintArea => "printArea",
            Self::AutoFilter => "autoFilter",
            Self::Range => "range",
            Self::Chart => "chart",
            Self::PivotTable => "pivotTable",
            Self::Query => "query",
            Self::Label => "label",
        }
    }
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "sheet" => Ok(Self::Sheet),
            "printArea" => Ok(Self::PrintArea),
            "autoFilter" => Ok(Self::AutoFilter),
            "range" => Ok(Self::Range),
            "chart" => Ok(Self::Chart),
            "pivotTable" => Ok(Self::PivotTable),
            "query" => Ok(Self::Query),
            "label" => Ok(Self::Label),
            _ => Err(Error::Invalid(format!(
                "invalid web publish sourceType '{value}'"
            ))),
        }
    }
}

/// One `webPublishItem`; all paths, names, and references are opaque inert strings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebPublishItem {
    pub id: u32,
    pub div_id: String,
    pub source_type: WebSourceType,
    pub source_ref: Option<String>,
    pub source_object: Option<String>,
    pub destination_file: String,
    pub title: Option<String>,
    /// `None` preserves the schema default; no publishing action is ever performed.
    pub auto_republish: Option<bool>,
}

/// A present `webPublishItems` collection is schema-required to be non-empty.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebPublishItems {
    /// Preserves explicit `count`; when present it must equal `items.len()`.
    pub count: Option<u32>,
    pub items: Vec<WebPublishItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    pub properties: Option<Properties>,
    pub views: Vec<View>,
    pub protection: Option<Protection>,
    /// `None` preserves absence; a present collection must be non-empty.
    pub custom_views: Option<Vec<CustomView>>,
    pub margins: Option<Margins>,
    pub page_setup: Option<PageSetup>,
    pub header_footer: Option<HeaderFooter>,
    pub drawing_relationship_id: String,
    pub legacy_drawing_relationship_id: Option<String>,
    pub legacy_header_footer_drawing_relationship_id: Option<String>,
    /// Relationship for the optional tiled chartsheet background image.
    pub background_picture_relationship_id: Option<String>,
    /// Inert web-publishing metadata. Destinations and sources are never resolved or accessed.
    pub web_publish_items: Option<WebPublishItems>,
    /// Inert, canonicalized wildcard extension markup. Payload semantics are never interpreted.
    pub extension_list: Option<ExtensionList>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub uri: String,
    /// Canonical, namespace-aware XML for the single wildcard child of `ext`.
    pub payload_xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    pub extensions: Vec<Extension>,
}
