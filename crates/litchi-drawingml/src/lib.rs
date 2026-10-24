//! Shared `DrawingML` grammar independent of DOCX, PPTX, XLSX, and XLSB.
//!
//! [`chart`] owns the chart model and XML codec, while [`diagram`] owns the
//! host-neutral `SmartArt` model and part grammar. Concrete formats retain only
//! their package relationships and anchoring semantics.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]
// Public DrawingML types intentionally mirror ECMA-376 names, builder
// ownership, and boolean-rich schema records. Renaming or reshaping them here
// would break every host format without changing wire correctness.
#![allow(
    clippy::double_must_use,
    clippy::module_name_repetitions,
    clippy::ref_option,
    clippy::return_self_not_must_use,
    clippy::struct_excessive_bools,
    reason = "shared DrawingML APIs retain specification-shaped names, ownership, and centrally documented failure contracts"
)]
// Streaming chart, theme, geometry, and transform parsers reuse event-local
// names as raw XML becomes typed vocabulary.
#![allow(
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    reason = "short-lived parser bindings track successive XML events and semantic projections"
)]
// Codec declarations and string emitters follow schema traversal order. These
// bounded serializers favor direct append operations and explicit cases so the
// output can be audited against the standard.
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::doc_markdown,
    clippy::format_push_string,
    clippy::items_after_statements,
    clippy::unreadable_literal,
    reason = "DrawingML codecs retain schema order and bounded direct XML emission for specification auditability"
)]

pub mod blip;
pub mod chart;
pub mod color;
pub mod coordinate;
pub mod diagram;
pub mod ext;
pub mod fill;
pub mod geom;
pub mod geometry;
pub mod ink;
pub mod model3d;
pub mod svg_blip;
pub mod text;
pub mod theme;
pub mod transform;

use std::fmt;

use thiserror::Error;

/// Preflight the decoded XML 1.0 attribute-value length without constructing
/// quick-xml's owned normalization buffer. Callers still perform their
/// module-specific raw-byte guard before decoding; this pass accounts for
/// entity expansion and XML whitespace normalization so escaped values are not
/// rejected merely because their lexical form is longer.
pub(crate) fn preflight_attribute_value_bytes(value: &[u8], maximum: usize) -> Result<()> {
    preflight_attribute_value_bytes_with_kind(value, maximum, LimitKind::AttributeValueBytes)
}

/// Preflight a decoded XML attribute under a caller-owned typed dimension.
/// InkML uses a distinct public dimension from the generic DrawingML XML
/// codecs, while sharing the same entity/whitespace normalization pass.
pub(crate) fn preflight_attribute_value_bytes_with_kind(
    value: &[u8],
    maximum: usize,
    kind: LimitKind,
) -> Result<()> {
    let value = std::str::from_utf8(value)
        .map_err(|error| Error::Xml(format!("attribute value is not UTF-8: {error}")))?;
    let mut length = 0usize;
    let mut cursor = 0usize;
    while cursor < value.len() {
        let character = value[cursor..]
            .chars()
            .next()
            .ok_or_else(|| Error::Xml("attribute value has an invalid UTF-8 boundary".into()))?;
        match character {
            '&' => {
                let end = value[cursor + 1..]
                    .find(';')
                    .map(|offset| cursor + 1 + offset)
                    .ok_or_else(|| Error::Xml("attribute entity is unterminated".into()))?;
                let entity = &value[cursor + 1..end];
                let decoded = if let Some(number) = entity.strip_prefix("#x") {
                    u32::from_str_radix(number, 16)
                        .ok()
                        .and_then(char::from_u32)
                        .ok_or_else(|| {
                            Error::Xml("attribute character reference is invalid".into())
                        })?
                        .len_utf8()
                } else if let Some(number) = entity.strip_prefix("#X") {
                    u32::from_str_radix(number, 16)
                        .ok()
                        .and_then(char::from_u32)
                        .ok_or_else(|| {
                            Error::Xml("attribute character reference is invalid".into())
                        })?
                        .len_utf8()
                } else if let Some(number) = entity.strip_prefix('#') {
                    number
                        .parse::<u32>()
                        .ok()
                        .and_then(char::from_u32)
                        .ok_or_else(|| {
                            Error::Xml("attribute character reference is invalid".into())
                        })?
                        .len_utf8()
                } else {
                    quick_xml::escape::resolve_xml_entity(entity)
                        .map(str::len)
                        .ok_or_else(|| Error::Xml("attribute entity is unknown".into()))?
                };
                length = length
                    .checked_add(decoded)
                    .ok_or_else(|| Error::Xml("attribute value length overflows usize".into()))?;
                cursor = end + 1;
            },
            '\r' => {
                length = length
                    .checked_add(1)
                    .ok_or_else(|| Error::Xml("attribute value length overflows usize".into()))?;
                cursor += 1;
                if value.as_bytes().get(cursor) == Some(&b'\n') {
                    cursor += 1;
                }
            },
            '\n' | '\t' => {
                length = length
                    .checked_add(1)
                    .ok_or_else(|| Error::Xml("attribute value length overflows usize".into()))?;
                cursor += character.len_utf8();
            },
            _ => {
                length = length
                    .checked_add(character.len_utf8())
                    .ok_or_else(|| Error::Xml("attribute value length overflows usize".into()))?;
                cursor += character.len_utf8();
            },
        }
    }
    if length > maximum {
        return Err(Error::limit(kind, length, maximum));
    }
    Ok(())
}

/// Result of a `DrawingML` read operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Typed resource dimensions used by every bounded DrawingML reader and
/// editor.
///
/// The enum is intentionally crate-wide because [`enum@Error`] is the shared
/// boundary for the chart, transform, theme, diagram, text, color, model3d,
/// InkML, and action families.  Family modules may re-export the vocabulary
/// they own, but a returned error always carries this nominal type rather
/// than a resource string which callers would have to reinterpret.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete serialized output.
    OutputBytes,
    /// Retained source bytes.
    SourceBytes,
    /// One opaque payload.
    PayloadBytes,
    /// XML bytes for one family value.
    XmlBytes,
    /// XML node count.
    XmlNodes,
    /// XML nesting depth.
    XmlDepth,
    /// XML attribute count.
    XmlAttributes,
    /// Decoded XML attribute or scalar value bytes.
    AttributeValueBytes,
    /// Namespace prefix bytes.
    NamespacePrefixBytes,
    /// Namespace URI bytes.
    NamespaceBytes,
    /// Generic structural nodes.
    Nodes,
    /// Generic structural depth.
    Depth,
    /// Typed scalar bytes.
    ScalarBytes,
    /// Text bytes.
    TextBytes,
    /// Lexical (pre-decoding) value bytes.
    LexicalValueBytes,
    /// Source input bytes.
    InputBytes,
    /// Action records.
    Actions,
    /// Action groups.
    ActionGroups,
    /// Queued edit operations.
    EditOperations,
    /// Action element-name bytes.
    ActionElementNameBytes,
    /// Action namespace-prefix bytes.
    ActionNamespacePrefixBytes,
    /// Action opaque-attribute value bytes.
    ActionOpaqueAttributeValueBytes,
    /// Action reference value bytes.
    ActionReferenceValueBytes,
    /// InkML context nodes.
    InkMlContextNodes,
    /// InkML trace records.
    InkMlTraces,
    /// InkML brush properties.
    InkMlBrushProperties,
    /// InkML brush records.
    InkMlBrushRecords,
    /// InkML source bytes.
    InkMlSourceBytes,
    /// InkML output bytes.
    InkMlOutputBytes,
    /// InkML XML nodes.
    InkMlXmlNodes,
    /// InkML XML depth.
    InkMlXmlDepth,
    /// InkML identifier bytes.
    InkMlIdentifierBytes,
    /// InkML reference bytes.
    InkMlReferenceBytes,
    /// InkML trace text bytes.
    InkMlTraceTextBytes,
    /// InkML brush text bytes.
    InkMlBrushTextBytes,
    /// InkML retained authoring text.
    InkMlRetainedAuthoringText,
    /// InkML attributes per element.
    InkMlAttributesPerElement,
    /// InkML attribute value bytes.
    InkMlAttributeValueBytes,
    /// Chart series.
    ChartSeries,
    /// Chart categories.
    ChartCategories,
    /// Chart points.
    ChartPoints,
    /// Chart text bytes.
    ChartTextBytes,
    /// Chart-style color entries.
    ChartStyleColors,
    /// Chart-style variations.
    ChartStyleVariations,
    /// Chart-style color transforms.
    ChartStyleTransforms,
    /// Chart-style modifier records.
    ChartStyleModifiers,
    /// One chart-style modifier's bytes.
    ChartStyleModifierBytes,
    /// Chart-style retained string bytes.
    ChartStyleStrings,
    /// Extended-chart part bytes.
    ChartexPartBytes,
    /// Extended-chart XML nodes.
    ChartexXmlNodes,
    /// Extended-chart XML nesting depth.
    ChartexXmlDepth,
    /// Color transform count.
    ColorTransforms,
    /// Diagram data-model bytes.
    DiagramDataModelBytes,
    /// Diagram drawing bytes.
    DiagramDrawingBytes,
    /// Diagram points.
    DiagramPoints,
    /// Diagram connections.
    DiagramConnections,
    /// Diagram data-model text bytes.
    DiagramDataModelTextBytes,
    /// Diagram data-model identifier bytes.
    DiagramDataModelIdentifierBytes,
    /// Diagram nodes.
    DiagramNodes,
    /// Diagram node depth.
    DiagramNodeDepth,
    /// Diagram node text bytes.
    DiagramNodeTextBytes,
    /// Custom-geometry guides.
    GeometryGuides,
    /// Custom-geometry handles or connection sites.
    GeometrySites,
    /// Custom-geometry paths.
    GeometryPaths,
    /// Custom-geometry path commands.
    GeometryPathCommands,
    /// SmartArt nodes.
    SmartArtNodes,
    /// SmartArt node depth.
    SmartArtNodeDepth,
    /// SmartArt node text bytes.
    SmartArtNodeTextBytes,
    /// Theme-family name bytes.
    ThemeFamilyNameBytes,
    /// Text-body elements.
    TextBodyElements,
    /// Text-body nesting depth.
    TextBodyNestingDepth,
    /// Model3D child records.
    ModelChildren,
    /// Model3D namespace declarations.
    NamespaceDeclarations,
    /// Model3D raster children.
    RasterChildren,
    /// Model3D renderer text.
    RendererText,
    /// Model3D inert fragment bytes.
    InertFragmentBytes,
    /// Model3D inert fragment nodes.
    InertFragmentNodes,
    /// Model3D inert fragment depth.
    InertFragmentDepth,
}

impl LimitKind {
    /// Stable diagnostic name for this dimension.
    #[must_use]
    pub const fn resource(self) -> &'static str {
        match self {
            Self::OutputBytes => "DrawingML output bytes",
            Self::SourceBytes => "DrawingML source bytes",
            Self::PayloadBytes => "DrawingML payload bytes",
            Self::XmlBytes => "DrawingML XML bytes",
            Self::XmlNodes => "DrawingML XML nodes",
            Self::XmlDepth => "DrawingML XML depth",
            Self::XmlAttributes => "DrawingML XML attributes",
            Self::AttributeValueBytes => "DrawingML XML attribute value bytes",
            Self::NamespacePrefixBytes => "DrawingML namespace prefix bytes",
            Self::NamespaceBytes => "DrawingML namespace bytes",
            Self::Nodes => "DrawingML nodes",
            Self::Depth => "DrawingML depth",
            Self::ScalarBytes => "DrawingML scalar bytes",
            Self::TextBytes => "DrawingML text bytes",
            Self::LexicalValueBytes => "DrawingML lexical value bytes",
            Self::InputBytes => "DrawingML input bytes",
            Self::Actions => "ink action records",
            Self::ActionGroups => "ink action groups",
            Self::EditOperations => "ink action edit operations",
            Self::ActionElementNameBytes => "ink action element name bytes",
            Self::ActionNamespacePrefixBytes => "ink action namespace prefix bytes",
            Self::ActionOpaqueAttributeValueBytes => "ink action opaque attribute value bytes",
            Self::ActionReferenceValueBytes => "ink action reference value bytes",
            Self::InkMlContextNodes => "InkML context nodes",
            Self::InkMlTraces => "InkML traces",
            Self::InkMlBrushProperties => "InkML brush properties",
            Self::InkMlBrushRecords => "InkML brush records",
            Self::InkMlSourceBytes => "InkML source bytes",
            Self::InkMlOutputBytes => "InkML output bytes",
            Self::InkMlXmlNodes => "InkML XML nodes",
            Self::InkMlXmlDepth => "InkML XML depth",
            Self::InkMlIdentifierBytes => "InkML identifier bytes",
            Self::InkMlReferenceBytes => "InkML reference bytes",
            Self::InkMlTraceTextBytes => "InkML trace text bytes",
            Self::InkMlBrushTextBytes => "InkML brush text bytes",
            Self::InkMlRetainedAuthoringText => "InkML retained authoring text",
            Self::InkMlAttributesPerElement => "InkML attributes per element",
            Self::InkMlAttributeValueBytes => "InkML attribute value bytes",
            Self::ChartSeries => "authored chart series",
            Self::ChartCategories => "authored chart categories",
            Self::ChartPoints => "authored chart points",
            Self::ChartTextBytes => "authored chart text bytes",
            Self::ChartStyleColors => "chart-style colors",
            Self::ChartStyleVariations => "chart-style variations",
            Self::ChartStyleTransforms => "chart-style color transforms",
            Self::ChartStyleModifiers => "chart-style modifiers",
            Self::ChartStyleModifierBytes => "chart-style modifier bytes",
            Self::ChartStyleStrings => "chart-style strings",
            Self::ChartexPartBytes => "chartex part bytes",
            Self::ChartexXmlNodes => "chartex XML nodes",
            Self::ChartexXmlDepth => "chartex XML depth",
            Self::ColorTransforms => "DrawingML color transforms",
            Self::DiagramDataModelBytes => "diagram data-model bytes",
            Self::DiagramDrawingBytes => "diagram drawing bytes",
            Self::DiagramPoints => "diagram points",
            Self::DiagramConnections => "diagram connections",
            Self::DiagramDataModelTextBytes => "diagram data-model text bytes",
            Self::DiagramDataModelIdentifierBytes => "diagram data-model identifier bytes",
            Self::DiagramNodes => "diagram nodes",
            Self::DiagramNodeDepth => "diagram node depth",
            Self::DiagramNodeTextBytes => "diagram node text bytes",
            Self::GeometryGuides => "custom geometry guides",
            Self::GeometrySites => "custom geometry sites",
            Self::GeometryPaths => "custom geometry paths",
            Self::GeometryPathCommands => "custom geometry path commands",
            Self::SmartArtNodes => "SmartArt nodes",
            Self::SmartArtNodeDepth => "SmartArt node depth",
            Self::SmartArtNodeTextBytes => "SmartArt node text bytes",
            Self::ThemeFamilyNameBytes => "theme family name bytes",
            Self::TextBodyElements => "text body elements",
            Self::TextBodyNestingDepth => "text body nesting depth",
            Self::ModelChildren => "children",
            Self::NamespaceDeclarations => "namespace declarations",
            Self::RasterChildren => "raster children",
            Self::RendererText => "renderer text",
            Self::InertFragmentBytes => "inert fragment bytes",
            Self::InertFragmentNodes => "inert fragment nodes",
            Self::InertFragmentDepth => "inert fragment depth",
        }
    }
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.resource())
    }
}

/// Failure to parse a shared `DrawingML` primitive.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The XML is malformed or contains an invalid encoded value.
    #[error("invalid DrawingML XML: {0}")]
    Xml(String),

    /// The document violates `DrawingML` structural or value constraints.
    #[error("invalid DrawingML: {0}")]
    Invalid(String),

    /// Reading a caller-provided chart stream failed.
    #[error("DrawingML input error: {0}")]
    Io(#[from] std::io::Error),

    /// Markup-compatibility preprocessing failed.
    #[error("DrawingML markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::mce::Error),

    /// Shared OOXML entity or namespace decoding failed.
    #[error("DrawingML shared XML error: {0}")]
    Common(#[from] litchi_ooxml_common::XmlError),

    /// A bounded `DrawingML` codec resource was exhausted.
    #[error("DrawingML {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Typed resource dimension whose configured bound was exceeded.
        kind: LimitKind,
        /// First observed value beyond the selected bound.
        observed: usize,
        /// Active upper bound.
        maximum: usize,
    },

    /// A caller supplied an invalid bounded-resource policy.
    #[error("invalid DrawingML {kind} limit {value} (allowed 1..={maximum})")]
    InvalidLimit {
        /// Typed resource dimension being configured.
        kind: LimitKind,
        /// Rejected caller value.
        value: usize,
        /// Hard maximum accepted for this dimension.
        maximum: usize,
    },

    /// A patch was made against a different immutable source owner.
    #[error("DrawingML patch source does not match the recipient")]
    PatchConflict,
}

impl Error {
    /// Construct an observed typed limit failure.
    #[must_use]
    pub(crate) const fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Self {
        Self::Limit {
            kind,
            observed,
            maximum,
        }
    }

    /// Construct an invalid typed limit failure.
    #[must_use]
    pub(crate) const fn invalid_limit(kind: LimitKind, value: usize, maximum: usize) -> Self {
        Self::InvalidLimit {
            kind,
            value,
            maximum,
        }
    }
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
