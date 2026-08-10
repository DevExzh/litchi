//! Typed inert resources in a chartsheet OPC graph.

use super::super::{Chart, State};
use super::invalid;
use crate::Result;
use crate::package::printer_settings::PrinterSettingsResource;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExtensionRelationshipTarget {
    Internal { part_name: String },
    External { target: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionRelationship {
    pub relationship_id: String,
    pub relationship_type: String,
    pub target: ExtensionRelationshipTarget,
}

/// Supported inert image media types for chartsheet backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BackgroundImageContentType {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    Emf,
    Wmf,
}

impl BackgroundImageContentType {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Bmp => "image/bmp",
            Self::Tiff => "image/tiff",
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
        }
    }
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "image/png" => Ok(Self::Png),
            "image/jpeg" => Ok(Self::Jpeg),
            "image/gif" => Ok(Self::Gif),
            "image/bmp" => Ok(Self::Bmp),
            "image/tiff" => Ok(Self::Tiff),
            "image/x-emf" => Ok(Self::Emf),
            "image/x-wmf" => Ok(Self::Wmf),
            _ => Err(invalid(format!(
                "unsupported chartsheet background image content type '{value}'"
            ))),
        }
    }
}

/// Opaque internal package resource referenced by `chartsheet/picture`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BackgroundPicture {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: BackgroundImageContentType,
    /// Preserved without decoding, rendering, metadata inspection, or external fetches.
    pub data: Vec<u8>,
}

/// Chartsheet `pageSetup` relationship paired with a generic inert DEVMODE resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrinterSettings {
    pub relationship_id: String,
    pub resource: PrinterSettingsResource,
}

/// Opaque internal VML drawing bytes; VML semantics are intentionally not interpreted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VmlDrawingResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without shape rendering, script, hyperlink, external fetch, or macro execution.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without chart evaluation, external-data loading, or macro execution.
    pub data: Vec<u8>,
    pub kind: ChartResourceKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartCompanionResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without rendering, theme application, formula evaluation, or external access.
    pub data: Vec<u8>,
}

/// Supported inert image payload types for a chartEx chart-user-shapes drawing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageContentType {
    Bmp,
    Gif,
    Png,
    Tif,
    Tiff,
    Icon,
    Pcx,
    Jpeg,
    Jp2,
    Emf,
    Wmf,
    Svg,
}

impl ImageContentType {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Bmp => "image/bmp",
            Self::Gif => "image/gif",
            Self::Png => "image/png",
            Self::Tif => "image/tif",
            Self::Tiff => "image/tiff",
            Self::Icon => "image/x-icon",
            Self::Pcx => "image/x-pcx",
            Self::Jpeg => "image/jpeg",
            Self::Jp2 => "image/jp2",
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
            Self::Svg => "image/svg+xml",
        }
    }
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "image/bmp" => Ok(Self::Bmp),
            "image/gif" => Ok(Self::Gif),
            "image/png" => Ok(Self::Png),
            "image/tif" => Ok(Self::Tif),
            "image/tiff" => Ok(Self::Tiff),
            "image/x-icon" => Ok(Self::Icon),
            "image/x-pcx" => Ok(Self::Pcx),
            "image/jpeg" => Ok(Self::Jpeg),
            "image/jp2" => Ok(Self::Jp2),
            "image/x-emf" => Ok(Self::Emf),
            "image/x-wmf" => Ok(Self::Wmf),
            "image/svg+xml" => Ok(Self::Svg),
            _ => Err(invalid(format!(
                "unsupported chart user-shape image content type '{value}'"
            ))),
        }
    }
    pub(super) fn validates_part_name(self, value: &str) -> bool {
        let lower = value.to_ascii_lowercase();
        match self {
            Self::Bmp => lower.ends_with(".bmp"),
            Self::Gif => lower.ends_with(".gif"),
            Self::Png => lower.ends_with(".png"),
            Self::Tif => lower.ends_with(".tif"),
            Self::Tiff => lower.ends_with(".tiff"),
            Self::Icon => lower.ends_with(".ico"),
            Self::Pcx => lower.ends_with(".pcx"),
            Self::Jpeg => lower.ends_with(".jpg") || lower.ends_with(".jpeg"),
            Self::Jp2 => lower.ends_with(".jp2"),
            Self::Emf => lower.ends_with(".emf"),
            Self::Wmf => lower.ends_with(".wmf"),
            Self::Svg => lower.ends_with(".svg"),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: ImageContentType,
    /// Preserved without decoding, rendering, metadata inspection, or external fetches.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartUserShapesResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without interpreting shapes, actions, hyperlinks, or embedded markup.
    pub data: Vec<u8>,
    pub images: Vec<ImageResource>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartEmbeddedPackageContentType {
    Docx,
    Dotx,
    Potx,
    Ppsx,
    Pptx,
    Sldx,
    Thmx,
    Xlsx,
    Xltx,
}
impl ChartEmbeddedPackageContentType {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Docx => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            Self::Dotx => "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
            Self::Potx => "application/vnd.openxmlformats-officedocument.presentationml.template",
            Self::Ppsx => "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
            Self::Pptx => {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            },
            Self::Sldx => "application/vnd.openxmlformats-officedocument.presentationml.slide",
            Self::Thmx => "application/vnd.ms-officetheme",
            Self::Xlsx => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            Self::Xltx => "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
        }
    }
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => {
                Ok(Self::Docx)
            },
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template" => {
                Ok(Self::Dotx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.template" => {
                Ok(Self::Potx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow" => {
                Ok(Self::Ppsx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" => {
                Ok(Self::Pptx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.slide" => Ok(Self::Sldx),
            "application/vnd.ms-officetheme" => Ok(Self::Thmx),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => Ok(Self::Xlsx),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template" => {
                Ok(Self::Xltx)
            },
            _ => Err(invalid(format!(
                "unsupported or active chartEx embedded-package content type '{value}'"
            ))),
        }
    }
    pub(super) fn validates_part_name(self, value: &str) -> bool {
        let suffix = match self {
            Self::Docx => ".docx",
            Self::Dotx => ".dotx",
            Self::Potx => ".potx",
            Self::Ppsx => ".ppsx",
            Self::Pptx => ".pptx",
            Self::Sldx => ".sldx",
            Self::Thmx => ".thmx",
            Self::Xlsx => ".xlsx",
            Self::Xltx => ".xltx",
        };
        value.to_ascii_lowercase().ends_with(suffix)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartThemeOverrideResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without theme application, image rendering, or external access.
    pub data: Vec<u8>,
    pub images: Vec<ImageResource>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartEmbeddedPackageResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: ChartEmbeddedPackageContentType,
    /// Preserved without opening, parsing, activation, macro execution, or external access.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExternalResource {
    pub relationship_id: String,
    pub relationship_type: String,
    /// Inert external target retained verbatim and never fetched.
    pub target: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartOpaqueResource {
    pub relationship_id: String,
    pub relationship_type: String,
    pub part_name: String,
    pub content_type: String,
    /// Relationship-free inert payload retained without interpretation.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartOutboundResource {
    Image(ImageResource),
    ThemeOverride(ChartThemeOverrideResource),
    EmbeddedPackage(ChartEmbeddedPackageResource),
    External(ChartExternalResource),
    Opaque(ChartOpaqueResource),
}
impl ChartOutboundResource {
    pub(super) fn relationship_id(&self) -> &str {
        match self {
            Self::Image(value) => &value.relationship_id,
            Self::ThemeOverride(value) => &value.relationship_id,
            Self::EmbeddedPackage(value) => &value.relationship_id,
            Self::External(value) => &value.relationship_id,
            Self::Opaque(value) => &value.relationship_id,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartResourceKind {
    Classic,
    ClassicWithRelationships {
        user_shapes: Option<ChartUserShapesResource>,
        outbound_resources: Vec<ChartOutboundResource>,
    },
    Extended {
        styles: Vec<ChartCompanionResource>,
        color_styles: Vec<ChartCompanionResource>,
        user_shapes: Option<ChartUserShapesResource>,
        outbound_resources: Vec<ChartOutboundResource>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingResource {
    pub part_name: String,
    pub content_type: String,
    /// Preserved without rendering or interpreting drawing actions.
    pub data: Vec<u8>,
    pub charts: Vec<ChartResource>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    pub name: String,
    pub sheet_id: u32,
    pub state: State,
    pub workbook_relationship_id: String,
    pub part_name: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Package {
    pub entry: Entry,
    pub chartsheet: Chart,
    pub drawing: DrawingResource,
    pub legacy_drawing: Option<VmlDrawingResource>,
    pub legacy_header_footer_drawing: Option<VmlDrawingResource>,
    pub background_picture: Option<BackgroundPicture>,
    pub printer_settings: Option<PrinterSettings>,
    /// Metadata for otherwise-unknown relationships referenced by extension markup.
    pub extension_relationships: Vec<ExtensionRelationship>,
}
