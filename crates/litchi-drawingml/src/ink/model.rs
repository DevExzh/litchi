//! Semantic values for source-backed InkML and DrawingML ink metadata.

use std::{borrow::Cow, fmt, ops::Range, sync::Arc};

use thiserror::Error;

use super::{
    MAX_BRUSH_PROPERTIES, MAX_CONTEXTS, MAX_GUID_BYTES, MAX_SOURCE_BYTES, MAX_TOKEN_BYTES,
    MAX_TRACES,
};

/// A source byte span into an Ink document source buffer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[must_use]
pub struct SourceSpan {
    start: usize,
    end: usize,
}

impl SourceSpan {
    /// Construct an unchecked half-open source span.
    ///
    /// A source-backed reader validates the span before exposing it. Callers
    /// that select semantic projections must use spans from the same source;
    /// unmatched selections are rejected by the filtered readers.
    pub const fn new(start: usize, end: usize) -> Self {
        Self { start, end }
    }
    /// Start offset, inclusive.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }
    /// End offset, exclusive.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }
    /// Return the span as a standard half-open range.
    #[must_use]
    pub const fn range(self) -> Range<usize> {
        self.start..self.end
    }
}

/// A checked InkML GUID lexical value.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[must_use]
pub struct Guid(Box<str>);

impl Guid {
    /// Construct a GUID in the uppercase brace form required by MS-ODRAWXML.
    /// # Errors
    ///
    /// Returns an error for any non-canonical value.
    pub fn new(value: impl AsRef<str>) -> Result<Self, ValueError> {
        let value = value.as_ref();
        if value.len() > MAX_GUID_BYTES {
            return Err(ValueError::TooLong {
                field: "GUID",
                limit: MAX_GUID_BYTES,
            });
        }
        if !is_guid(value) {
            return Err(ValueError::Guid {
                value: value.to_owned(),
            });
        }
        Ok(Self(value.into()))
    }
    /// Borrow the exact GUID lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Guid {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for Guid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for Guid {
    type Error = ValueError;
    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

/// Known or future InkML context-node type.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ContextKind {
    Root,
    UnclassifiedInk,
    WritingRegion,
    AnalysisHint,
    Object,
    InkDrawing,
    Image,
    Paragraph,
    Line,
    InkBullet,
    InkWord,
    TextWord,
    CustomRecognizer,
    MathRegion,
    MathEquation,
    MathStruct,
    MathSymbol,
    MathIdentifier,
    MathOperator,
    MathNumber,
    NonInkDrawing,
    GroupNode,
    MixedDrawing,
    Custom(Guid),
}

impl ContextKind {
    pub(crate) fn parse(value: &str) -> Result<Self, ValueError> {
        if value.len() > MAX_TOKEN_BYTES {
            return Err(ValueError::TooLong {
                field: "semantic type",
                limit: MAX_TOKEN_BYTES,
            });
        }
        Ok(match value {
            "root" => Self::Root,
            "unclassifiedInk" => Self::UnclassifiedInk,
            "writingRegion" => Self::WritingRegion,
            "analysisHint" => Self::AnalysisHint,
            "object" => Self::Object,
            "inkDrawing" => Self::InkDrawing,
            "image" => Self::Image,
            "paragraph" => Self::Paragraph,
            "line" => Self::Line,
            "inkBullet" => Self::InkBullet,
            "inkWord" => Self::InkWord,
            "textWord" => Self::TextWord,
            "customRecognizer" => Self::CustomRecognizer,
            "mathRegion" => Self::MathRegion,
            "mathEquation" => Self::MathEquation,
            "mathStruct" => Self::MathStruct,
            "mathSymbol" => Self::MathSymbol,
            "mathIdentifier" => Self::MathIdentifier,
            "mathOperator" => Self::MathOperator,
            "mathNumber" => Self::MathNumber,
            "nonInkDrawing" => Self::NonInkDrawing,
            "groupNode" => Self::GroupNode,
            "mixedDrawing" => Self::MixedDrawing,
            _ => Self::Custom(Guid::new(value)?),
        })
    }
    /// Return the exact schema lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Root => "root",
            Self::UnclassifiedInk => "unclassifiedInk",
            Self::WritingRegion => "writingRegion",
            Self::AnalysisHint => "analysisHint",
            Self::Object => "object",
            Self::InkDrawing => "inkDrawing",
            Self::Image => "image",
            Self::Paragraph => "paragraph",
            Self::Line => "line",
            Self::InkBullet => "inkBullet",
            Self::InkWord => "inkWord",
            Self::TextWord => "textWord",
            Self::CustomRecognizer => "customRecognizer",
            Self::MathRegion => "mathRegion",
            Self::MathEquation => "mathEquation",
            Self::MathStruct => "mathStruct",
            Self::MathSymbol => "mathSymbol",
            Self::MathIdentifier => "mathIdentifier",
            Self::MathOperator => "mathOperator",
            Self::MathNumber => "mathNumber",
            Self::NonInkDrawing => "nonInkDrawing",
            Self::GroupNode => "groupNode",
            Self::MixedDrawing => "mixedDrawing",
            Self::Custom(value) => value.as_str(),
        }
    }
}

/// Known or future semantic classification for a context node.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SemanticType {
    None,
    Underline,
    Strikethrough,
    Highlight,
    ScratchOut,
    VerticalRange,
    Callout,
    Enclosure,
    Comment,
    Container,
    Connector,
    Custom(u32),
}

impl SemanticType {
    pub(crate) fn parse(value: &str) -> Result<Self, ValueError> {
        if value.len() > MAX_TOKEN_BYTES {
            return Err(ValueError::TooLong {
                field: "semantic type",
                limit: MAX_TOKEN_BYTES,
            });
        }
        Ok(match value {
            "none" => Self::None,
            "underline" => Self::Underline,
            "strikethrough" => Self::Strikethrough,
            "highlight" => Self::Highlight,
            "scratchOut" => Self::ScratchOut,
            "verticalRange" => Self::VerticalRange,
            "callout" => Self::Callout,
            "enclosure" => Self::Enclosure,
            "comment" => Self::Comment,
            "container" => Self::Container,
            "connector" => Self::Connector,
            _ => Self::Custom(value.parse().map_err(|_| ValueError::SemanticType {
                value: value.to_owned(),
            })?),
        })
    }
    /// Return the schema lexical value.
    #[must_use]
    pub fn as_string(&self) -> String {
        match self {
            Self::None => "none".into(),
            Self::Underline => "underline".into(),
            Self::Strikethrough => "strikethrough".into(),
            Self::Highlight => "highlight".into(),
            Self::ScratchOut => "scratchOut".into(),
            Self::VerticalRange => "verticalRange".into(),
            Self::Callout => "callout".into(),
            Self::Enclosure => "enclosure".into(),
            Self::Comment => "comment".into(),
            Self::Container => "container".into(),
            Self::Connector => "connector".into(),
            Self::Custom(value) => value.to_string(),
        }
    }
}

/// Known InkML brush property or a future/custom property name.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BrushPropertyName {
    Width,
    Height,
    Color,
    Transparency,
    Tip,
    RasterOp,
    AntiAliased,
    FitToCurve,
    IgnorePressure,
    InkEffects,
    AnchorX,
    AnchorY,
    ScaleFactor,
    Custom(Box<str>),
}

impl BrushPropertyName {
    pub(crate) fn parse(value: &str) -> Self {
        match value {
            "width" => Self::Width,
            "height" => Self::Height,
            "color" => Self::Color,
            "transparency" => Self::Transparency,
            "tip" => Self::Tip,
            "rasterOp" => Self::RasterOp,
            "antiAliased" => Self::AntiAliased,
            "fitToCurve" => Self::FitToCurve,
            "ignorePressure" => Self::IgnorePressure,
            "inkEffects" => Self::InkEffects,
            "anchorX" => Self::AnchorX,
            "anchorY" => Self::AnchorY,
            "scaleFactor" => Self::ScaleFactor,
            _ => Self::Custom(value.into()),
        }
    }
    /// Return the source lexical property name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Width => "width",
            Self::Height => "height",
            Self::Color => "color",
            Self::Transparency => "transparency",
            Self::Tip => "tip",
            Self::RasterOp => "rasterOp",
            Self::AntiAliased => "antiAliased",
            Self::FitToCurve => "fitToCurve",
            Self::IgnorePressure => "ignorePressure",
            Self::InkEffects => "inkEffects",
            Self::AnchorX => "anchorX",
            Self::AnchorY => "anchorY",
            Self::ScaleFactor => "scaleFactor",
            Self::Custom(value) => value,
        }
    }
}

/// Known or future DrawingML 2016 ink effect texture.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum InkEffect {
    None,
    Pencil,
    Rainbow,
    Galaxy,
    Gold,
    Silver,
    Lava,
    Ocean,
    Rosegold,
    Bronze,
    Custom(Box<str>),
}

impl InkEffect {
    fn parse(value: &str) -> Self {
        match value {
            "none" => Self::None,
            "pencil" => Self::Pencil,
            "rainbow" => Self::Rainbow,
            "galaxy" => Self::Galaxy,
            "gold" => Self::Gold,
            "silver" => Self::Silver,
            "lava" => Self::Lava,
            "ocean" => Self::Ocean,
            "rosegold" => Self::Rosegold,
            "bronze" => Self::Bronze,
            _ => Self::Custom(value.into()),
        }
    }

    /// Return the exact schema lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::None => "none",
            Self::Pencil => "pencil",
            Self::Rainbow => "rainbow",
            Self::Galaxy => "galaxy",
            Self::Gold => "gold",
            Self::Silver => "silver",
            Self::Lava => "lava",
            Self::Ocean => "ocean",
            Self::Rosegold => "rosegold",
            Self::Bronze => "bronze",
            Self::Custom(value) => value,
        }
    }
}

/// A typed context node discovered inside InkML annotation metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Context {
    pub(crate) kind: ContextKind,
    pub(crate) id: Option<Guid>,
    pub(crate) semantic_type: Option<SemanticType>,
    pub(crate) alignment_level: Option<i32>,
    pub(crate) content_type: Option<i32>,
    pub(crate) rotation_angle: Option<i32>,
    pub(crate) source: SourceSpan,
}

impl Context {
    /// Context-node type.
    #[must_use]
    pub const fn kind(&self) -> &ContextKind {
        &self.kind
    }
    /// Optional schema GUID.
    #[must_use]
    pub const fn id(&self) -> Option<&Guid> {
        self.id.as_ref()
    }
    /// Optional semantic classification.
    #[must_use]
    pub const fn semantic_type(&self) -> Option<&SemanticType> {
        self.semantic_type.as_ref()
    }
    /// Optional paragraph alignment level.
    #[must_use]
    pub const fn alignment_level(&self) -> Option<i32> {
        self.alignment_level
    }
    /// Optional paragraph content type.
    #[must_use]
    pub const fn content_type(&self) -> Option<i32> {
        self.content_type
    }
    /// Optional drawing rotation angle.
    #[must_use]
    pub const fn rotation_angle(&self) -> Option<i32> {
        self.rotation_angle
    }
    /// Source span of the complete context element.
    #[must_use]
    pub const fn source_span(&self) -> SourceSpan {
        self.source
    }
    /// Borrow the exact context element from its owning document.
    #[must_use]
    pub fn xml<'a>(&self, document: &'a Document) -> &'a [u8] {
        document.fragment(self.source)
    }
}

/// A typed InkML trace locator. Payload is borrowed from the source buffer.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Trace {
    pub(crate) context_ref: Option<Box<str>>,
    pub(crate) brush_ref: Option<Box<str>>,
    pub(crate) source: SourceSpan,
    pub(crate) data: SourceSpan,
}

impl Trace {
    /// Optional contextRef value.
    #[must_use]
    pub fn context_ref(&self) -> Option<&str> {
        self.context_ref.as_deref()
    }
    /// Optional brushRef value.
    #[must_use]
    pub fn brush_ref(&self) -> Option<&str> {
        self.brush_ref.as_deref()
    }
    /// Complete source span of the trace element.
    #[must_use]
    pub const fn source_span(&self) -> SourceSpan {
        self.source
    }
    /// Borrow the exact trace coordinate payload without allocating.
    #[must_use]
    pub fn data<'a>(&self, document: &'a Document) -> &'a [u8] {
        document.fragment(self.data)
    }
    /// Borrow the exact trace element XML.
    #[must_use]
    pub fn xml<'a>(&self, document: &'a Document) -> &'a [u8] {
        document.fragment(self.source)
    }
}

/// A typed inkml:brushProperty, including DrawingML 2016 effect names.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct BrushProperty {
    pub(crate) name: BrushPropertyName,
    pub(crate) value: Box<str>,
    pub(crate) units: Option<Box<str>>,
    pub(crate) source: SourceSpan,
}

impl BrushProperty {
    /// Typed property name.
    #[must_use]
    pub const fn name(&self) -> &BrushPropertyName {
        &self.name
    }
    /// Exact decoded property value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
    /// Optional decoded units.
    #[must_use]
    pub fn units(&self) -> Option<&str> {
        self.units.as_deref()
    }
    /// Interpret the DrawingML 2016 `inkEffects` value when this is that
    /// property. Unknown future values remain available losslessly.
    #[must_use]
    pub fn ink_effect(&self) -> Option<InkEffect> {
        if matches!(self.name, BrushPropertyName::InkEffects) {
            Some(InkEffect::parse(&self.value))
        } else {
            None
        }
    }
    /// Return the effective MS-ODRAWXML brush value.
    ///
    /// The source [`Self::value`] and [`Self::units`] accessors always expose
    /// the decoded source lexicals. This method applies the profile defaults
    /// when a recognized base or DrawingML 2016 property has a missing or
    /// invalid value/unit pair. Unknown generic InkML properties return
    /// `None`; the generic reader still retains them in [`Self`].
    #[must_use]
    pub fn effective(&self) -> Option<EffectiveBrushProperty> {
        let normalized_value = collapse_xsd_whitespace(&self.value);
        let normalized_value = normalized_value.as_ref();
        let (value, units, defaulted) = match &self.name {
            BrushPropertyName::Width => {
                if is_xsd_decimal(normalized_value)
                    && self.units.as_deref().is_some_and(is_length_unit)
                {
                    (
                        normalized_value.to_owned().into_boxed_str(),
                        self.units.clone(),
                        false,
                    )
                } else {
                    (".053".into(), Some("cm".into()), true)
                }
            },
            BrushPropertyName::Height => {
                if is_xsd_decimal(normalized_value)
                    && self.units.as_deref().is_some_and(is_length_unit)
                {
                    (
                        normalized_value.to_owned().into_boxed_str(),
                        self.units.clone(),
                        false,
                    )
                } else {
                    (".001".into(), Some("cm".into()), true)
                }
            },
            BrushPropertyName::Color => {
                if is_rgb_hex(&self.value) && self.units.is_none() {
                    (self.value.clone(), None, false)
                } else {
                    ("#000000".into(), None, true)
                }
            },
            BrushPropertyName::Transparency => {
                if is_bounded_integer(normalized_value, 0, 255) && self.units.is_none() {
                    (normalized_value.to_owned().into_boxed_str(), None, false)
                } else {
                    ("0".into(), None, true)
                }
            },
            BrushPropertyName::Tip => {
                if matches!(self.value.as_ref(), "ellipse" | "rectangle") && self.units.is_none() {
                    (self.value.clone(), None, false)
                } else {
                    ("ellipse".into(), None, true)
                }
            },
            BrushPropertyName::RasterOp => {
                if is_raster_operation(&self.value) && self.units.is_none() {
                    (self.value.clone(), None, false)
                } else {
                    ("copyPen".into(), None, true)
                }
            },
            BrushPropertyName::AntiAliased => {
                if is_xsd_boolean(normalized_value) && self.units.is_none() {
                    (canonical_boolean(normalized_value).into(), None, false)
                } else {
                    ("true".into(), None, true)
                }
            },
            BrushPropertyName::FitToCurve | BrushPropertyName::IgnorePressure => {
                if is_xsd_boolean(normalized_value) && self.units.is_none() {
                    (canonical_boolean(normalized_value).into(), None, false)
                } else {
                    ("false".into(), None, true)
                }
            },
            BrushPropertyName::InkEffects => {
                if is_ink_effect(&self.value) && self.units.is_none() {
                    (self.value.clone(), None, false)
                } else {
                    ("none".into(), None, true)
                }
            },
            BrushPropertyName::AnchorX | BrushPropertyName::AnchorY => {
                if is_xsd_decimal(normalized_value) && self.units.is_none() {
                    (normalized_value.to_owned().into_boxed_str(), None, false)
                } else {
                    ("0".into(), None, true)
                }
            },
            BrushPropertyName::ScaleFactor => {
                if is_xsd_decimal(normalized_value) && self.units.is_none() {
                    (normalized_value.to_owned().into_boxed_str(), None, false)
                } else {
                    ("0.5".into(), None, true)
                }
            },
            BrushPropertyName::Custom(_) => return None,
        };
        Some(EffectiveBrushProperty {
            name: self.name.clone(),
            value,
            units,
            defaulted,
        })
    }
    /// Source span of the complete brushProperty element.
    #[must_use]
    pub const fn source_span(&self) -> SourceSpan {
        self.source
    }
    /// Borrow the exact source XML.
    #[must_use]
    pub fn xml<'a>(&self, document: &'a Document) -> &'a [u8] {
        document.fragment(self.source)
    }
}

/// The effective value of a recognized MS-ODRAWXML brush property.
///
/// This value is derived from source lexicals; it does not replace the exact
/// source retained by [`BrushProperty`]. `defaulted` is true when the profile
/// rule selected the specification default because the source value or units
/// were absent or invalid.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct EffectiveBrushProperty {
    name: BrushPropertyName,
    value: Box<str>,
    units: Option<Box<str>>,
    defaulted: bool,
}

impl EffectiveBrushProperty {
    /// Property name.
    #[must_use]
    pub const fn name(&self) -> &BrushPropertyName {
        &self.name
    }
    /// Effective value lexical form.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
    /// Effective optional units.
    #[must_use]
    pub fn units(&self) -> Option<&str> {
        self.units.as_deref()
    }
    /// Whether the profile substituted a specification default.
    #[must_use]
    pub const fn defaulted(&self) -> bool {
        self.defaulted
    }
}

fn collapse_xsd_whitespace(value: &str) -> Cow<'_, str> {
    if !value.chars().any(is_xsd_whitespace) {
        return Cow::Borrowed(value);
    }

    let mut collapsed = String::with_capacity(value.len());
    let mut pending_space = false;
    for character in value.chars() {
        if is_xsd_whitespace(character) {
            pending_space = true;
            continue;
        }
        if pending_space && !collapsed.is_empty() {
            collapsed.push(' ');
        }
        collapsed.push(character);
        pending_space = false;
    }
    Cow::Owned(collapsed)
}

const fn is_xsd_whitespace(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\r' | '\n')
}

fn is_xsd_decimal(value: &str) -> bool {
    let collapsed = collapse_xsd_whitespace(value);
    let value = collapsed.as_ref();
    let value = value.strip_prefix(['+', '-']).unwrap_or(value);
    if value.is_empty() {
        return false;
    }
    let mut parts = value.split('.');
    let integer = parts.next().unwrap_or_default();
    let fraction = parts.next();
    if parts.next().is_some() {
        return false;
    }
    integer.chars().all(|character| character.is_ascii_digit())
        && fraction.is_none_or(|part| part.chars().all(|character| character.is_ascii_digit()))
        && (fraction.is_some() || !integer.is_empty())
        && (fraction.is_none_or(|part| !part.is_empty()) || !integer.is_empty())
}

fn is_xsd_boolean(value: &str) -> bool {
    let collapsed = collapse_xsd_whitespace(value);
    matches!(collapsed.as_ref(), "true" | "false" | "1" | "0")
}

fn canonical_boolean(value: &str) -> &'static str {
    match value {
        "true" | "1" => "true",
        _ => "false",
    }
}

fn is_bounded_integer(value: &str, minimum: i32, maximum: i32) -> bool {
    let collapsed = collapse_xsd_whitespace(value);
    let value = collapsed.as_ref();
    let digits = value.strip_prefix(['+', '-']).unwrap_or(value);
    !digits.is_empty()
        && digits.chars().all(|character| character.is_ascii_digit())
        && value
            .parse::<i32>()
            .is_ok_and(|value| (minimum..=maximum).contains(&value))
}

fn is_rgb_hex(value: &str) -> bool {
    value.len() == 7
        && value.starts_with('#')
        && value[1..]
            .chars()
            .all(|character| character.is_ascii_hexdigit())
}

fn is_length_unit(value: &str) -> bool {
    matches!(value, "m" | "cm" | "mm" | "in" | "pt" | "pc" | "em" | "ex")
}

fn is_raster_operation(value: &str) -> bool {
    matches!(
        value,
        "black"
            | "copyPen"
            | "maskNotPen"
            | "maskPenNot"
            | "maskPen"
            | "mergeNotPen"
            | "mergePen"
            | "mergePenNot"
            | "noOperation"
            | "not"
            | "notCopyPen"
            | "notMaskPen"
            | "notMergePen"
            | "notXOrPen"
            | "white"
            | "xOrPen"
    )
}

fn is_ink_effect(value: &str) -> bool {
    matches!(
        value,
        "none"
            | "pencil"
            | "rainbow"
            | "galaxy"
            | "gold"
            | "silver"
            | "lava"
            | "ocean"
            | "rosegold"
            | "bronze"
    )
}

/// Immutable InkML metadata projection independent of the source buffer.
///
/// The projection owns only typed scalar metadata and source offsets. It does
/// not retain the input XML allocation; callers that need XML spans must keep
/// the corresponding source bytes separately and use a [`Document`]. Cloning
/// this value shares its three immutable metadata arrays.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Metadata {
    contexts: Arc<[Context]>,
    traces: Arc<[Trace]>,
    brush_properties: Arc<[BrushProperty]>,
}

impl Metadata {
    /// Borrow typed context nodes in source order.
    #[must_use]
    pub fn contexts(&self) -> &[Context] {
        self.contexts.as_ref()
    }
    /// Borrow typed trace locators in source order.
    #[must_use]
    pub fn traces(&self) -> &[Trace] {
        self.traces.as_ref()
    }
    /// Borrow typed brush properties in source order.
    #[must_use]
    pub fn brush_properties(&self) -> &[BrushProperty] {
        self.brush_properties.as_ref()
    }
    /// Number of discovered context nodes.
    #[must_use]
    pub fn context_count(&self) -> usize {
        self.contexts().len()
    }
    /// Number of discovered traces.
    #[must_use]
    pub fn trace_count(&self) -> usize {
        self.traces().len()
    }
    /// Number of discovered brush properties.
    #[must_use]
    pub fn brush_property_count(&self) -> usize {
        self.brush_properties().len()
    }

    pub(crate) fn from_parts(
        contexts: Vec<Context>,
        traces: Vec<Trace>,
        brush_properties: Vec<BrushProperty>,
    ) -> Self {
        Self {
            contexts: Arc::from(contexts.into_boxed_slice()),
            traces: Arc::from(traces.into_boxed_slice()),
            brush_properties: Arc::from(brush_properties.into_boxed_slice()),
        }
    }

    pub(crate) fn validate(&self, source_len: Option<usize>) -> Result<(), ValueError> {
        if self.contexts.len() > MAX_CONTEXTS {
            return Err(ValueError::Limit {
                resource: "InkML context nodes",
                limit: MAX_CONTEXTS,
            });
        }
        if self.traces.len() > MAX_TRACES {
            return Err(ValueError::Limit {
                resource: "InkML traces",
                limit: MAX_TRACES,
            });
        }
        if self.brush_properties.len() > MAX_BRUSH_PROPERTIES {
            return Err(ValueError::Limit {
                resource: "InkML brush properties",
                limit: MAX_BRUSH_PROPERTIES,
            });
        }
        if let Some(source_len) = source_len {
            for context in self.contexts() {
                validate_span(context.source, source_len)?;
            }
            for trace in self.traces() {
                validate_span(trace.source, source_len)?;
                validate_span(trace.data, source_len)?;
            }
            for property in self.brush_properties() {
                validate_span(property.source, source_len)?;
            }
        }
        Ok(())
    }
}

/// Immutable, source-backed InkML document view.
#[derive(Debug, Clone)]
#[must_use]
pub struct Document {
    pub(crate) source: Arc<Vec<u8>>,
    pub(crate) metadata: Metadata,
}

impl PartialEq for Document {
    fn eq(&self, other: &Self) -> bool {
        self.source.as_slice() == other.source.as_slice() && self.metadata == other.metadata
    }
}
impl Eq for Document {}

impl Document {
    /// Borrow the complete source bytes, including an optional XML declaration.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        self.source.as_slice()
    }
    /// Borrow the typed source-independent metadata projection.
    #[must_use]
    pub const fn metadata(&self) -> &Metadata {
        &self.metadata
    }
    /// Borrow typed context nodes in source order.
    #[must_use]
    pub fn contexts(&self) -> &[Context] {
        self.metadata.contexts()
    }
    /// Borrow typed trace locators in source order.
    #[must_use]
    pub fn traces(&self) -> &[Trace] {
        self.metadata.traces()
    }
    /// Borrow typed brush properties in source order.
    #[must_use]
    pub fn brush_properties(&self) -> &[BrushProperty] {
        self.metadata.brush_properties()
    }
    /// Return the source fragment for a checked span.
    #[must_use]
    pub fn fragment(&self, span: SourceSpan) -> &[u8] {
        self.source.get(span.range()).unwrap_or_default()
    }
    /// Number of discovered context nodes.
    #[must_use]
    pub fn context_count(&self) -> usize {
        self.contexts().len()
    }
    /// Number of discovered traces.
    #[must_use]
    pub fn trace_count(&self) -> usize {
        self.traces().len()
    }
    /// Number of discovered brush properties.
    #[must_use]
    pub fn brush_property_count(&self) -> usize {
        self.brush_properties().len()
    }

    pub(crate) fn new(source: Arc<Vec<u8>>, metadata: Metadata) -> Result<Self, ValueError> {
        let value = Self { source, metadata };
        value.validate()?;
        Ok(value)
    }
    pub(crate) fn validate(&self) -> Result<(), ValueError> {
        if self.source.len() > MAX_SOURCE_BYTES {
            return Err(ValueError::Limit {
                resource: "InkML source bytes",
                limit: MAX_SOURCE_BYTES,
            });
        }
        self.metadata.validate(Some(self.source.len()))
    }
}

/// Construction or validation error for a typed ink scalar.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ValueError {
    #[error("InkML {resource} exceeds the limit of {limit}")]
    Limit {
        resource: &'static str,
        limit: usize,
    },
    #[error("InkML {field} exceeds the limit of {limit} bytes")]
    TooLong { field: &'static str, limit: usize },
    #[error("invalid InkML GUID '{value}'")]
    Guid { value: String },
    #[error("invalid InkML semantic type '{value}'")]
    SemanticType { value: String },
    #[error("InkML attribute '{name}' is missing")]
    MissingAttribute { name: &'static str },
    #[error("invalid InkML integer attribute '{name}'")]
    Integer { name: &'static str },
    #[error("invalid InkML points attribute '{name}'")]
    Points { name: &'static str },
    #[error("InkML source span is outside the source")]
    Span,
}

pub(crate) fn validate_span(span: SourceSpan, source_len: usize) -> Result<(), ValueError> {
    if span.start <= span.end && span.end <= source_len {
        Ok(())
    } else {
        Err(ValueError::Span)
    }
}

pub(crate) fn is_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 38
        && bytes[0] == b'{'
        && bytes[9] == b'-'
        && bytes[14] == b'-'
        && bytes[19] == b'-'
        && bytes[24] == b'-'
        && bytes[37] == b'}'
        && bytes
            .iter()
            .enumerate()
            .filter(|(index, _)| !matches!(*index, 0 | 9 | 14 | 19 | 24 | 37))
            .all(|(_, byte)| byte.is_ascii_hexdigit() && !byte.is_ascii_lowercase())
}
