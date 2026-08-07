//! Safe SVG emission primitives for EMF+ playback.
//!
//! This module intentionally has no knowledge of EMF+ record decoding.  The
//! playback layer supplies typed geometry and paint values; this module only
//! serializes the SVG subset it owns.  In particular, it never accepts SVG/XML
//! fragments, script URLs, external images, or `foreignObject` content.

use base64::Engine;
use litchi_core::{Error, Result};

const SVG_NAMESPACE: &str = "http://www.w3.org/2000/svg";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";

/// Resource limits for one SVG document.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SvgLimits {
    /// Maximum UTF-8 bytes in the completed document.
    pub max_output_bytes: usize,
    /// Maximum number of body elements.
    pub max_elements: usize,
    /// Maximum number of definitions.
    pub max_definitions: usize,
    /// Maximum path commands in a single path.
    pub max_path_commands: usize,
    /// Maximum input bytes accepted for one embedded image.
    pub max_image_bytes: usize,
    /// Maximum diagnostics retained in the output.
    pub max_diagnostics: usize,
}

impl SvgLimits {
    /// Reject unusable limit configurations before accepting any output.
    pub fn validate(self) -> Result<Self> {
        if self.max_output_bytes == 0 {
            return Err(invalid("max_output_bytes must be greater than zero"));
        }
        if self.max_elements == 0 {
            return Err(invalid("max_elements must be greater than zero"));
        }
        if self.max_definitions == 0 {
            return Err(invalid("max_definitions must be greater than zero"));
        }
        if self.max_path_commands == 0 {
            return Err(invalid("max_path_commands must be greater than zero"));
        }
        Ok(self)
    }
}

impl Default for SvgLimits {
    fn default() -> Self {
        Self {
            max_output_bytes: 32 * 1024 * 1024,
            max_elements: 1_000_000,
            max_definitions: 65_536,
            max_path_commands: 1_000_000,
            max_image_bytes: 16 * 1024 * 1024,
            max_diagnostics: 1_024,
        }
    }
}

/// A non-user-controlled ID generated for an SVG definition.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct SvgId(String);

impl SvgId {
    /// Return the identifier for diagnostic or caller-side bookkeeping.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A warning produced while safely degrading an unsupported operation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SvgDiagnostic {
    /// The unsupported feature that was skipped or approximated.
    pub feature: SvgUnsupportedFeature,
}

/// Inputs deliberately not represented by the safe SVG subset.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SvgUnsupportedFeature {
    /// SVG source text or markup was not accepted.
    RawMarkup,
    /// An image URI was rejected because only embedded bytes are accepted.
    ExternalImage,
    /// Source-copy cannot be faithfully represented by ordinary SVG painting.
    SourceCopyCompositing,
}

/// A safely produced SVG fragment, suitable for composition into a pattern.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct SvgFragment {
    defs: String,
    body: String,
    diagnostics: Vec<SvgDiagnostic>,
}

impl SvgFragment {
    /// The serialized definitions, without a surrounding `<defs>` element.
    #[must_use]
    pub fn defs(&self) -> &str {
        &self.defs
    }

    /// The serialized body elements, without an outer `<svg>` element.
    #[must_use]
    pub fn body(&self) -> &str {
        &self.body
    }

    /// Any safe-degradation notices accumulated while producing this fragment.
    #[must_use]
    pub fn diagnostics(&self) -> &[SvgDiagnostic] {
        &self.diagnostics
    }
}

/// Completed standalone SVG output and its component fragments.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SvgDocument {
    svg: String,
    fragment: SvgFragment,
}

impl SvgDocument {
    /// The complete standalone XML SVG document.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.svg
    }

    /// The safely generated component fragments.
    #[must_use]
    pub fn fragment(&self) -> &SvgFragment {
        &self.fragment
    }

    /// Consume the document and return its complete SVG source.
    #[must_use]
    pub fn into_string(self) -> String {
        self.svg
    }
}

/// An RGBA colour, with alpha expressed as an 8-bit channel.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SvgColor {
    /// Red channel.
    pub red: u8,
    /// Green channel.
    pub green: u8,
    /// Blue channel.
    pub blue: u8,
    /// Alpha channel.
    pub alpha: u8,
}

impl SvgColor {
    /// Construct an opaque RGB colour.
    #[must_use]
    pub const fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self::rgba(red, green, blue, u8::MAX)
    }

    /// Construct an RGBA colour.
    #[must_use]
    pub const fn rgba(red: u8, green: u8, blue: u8, alpha: u8) -> Self {
        Self {
            red,
            green,
            blue,
            alpha,
        }
    }
}

/// A coordinate pair.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct SvgPoint {
    /// Horizontal coordinate.
    pub x: f64,
    /// Vertical coordinate.
    pub y: f64,
}

/// An axis-aligned rectangle.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct SvgRect {
    /// Left coordinate.
    pub x: f64,
    /// Top coordinate.
    pub y: f64,
    /// Width (must be non-negative).
    pub width: f64,
    /// Height (must be non-negative).
    pub height: f64,
}

/// An affine transform in SVG's `matrix(a b c d e f)` order.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct SvgTransform {
    /// Matrix coefficient.
    pub a: f64,
    /// Matrix coefficient.
    pub b: f64,
    /// Matrix coefficient.
    pub c: f64,
    /// Matrix coefficient.
    pub d: f64,
    /// Translation X.
    pub e: f64,
    /// Translation Y.
    pub f: f64,
}

impl SvgTransform {
    /// The identity transform.
    #[must_use]
    pub const fn identity() -> Self {
        Self {
            a: 1.0,
            b: 0.0,
            c: 0.0,
            d: 1.0,
            e: 0.0,
            f: 0.0,
        }
    }
}

/// Typed path commands; unlike a raw `d` attribute, these cannot inject XML.
#[derive(Clone, Debug, PartialEq)]
pub enum SvgPathCommand {
    /// Begin a subpath.
    MoveTo(SvgPoint),
    /// Add a line.
    LineTo(SvgPoint),
    /// Add a cubic Bezier segment.
    CubicTo {
        /// First control point.
        first: SvgPoint,
        /// Second control point.
        second: SvgPoint,
        /// Destination.
        to: SvgPoint,
    },
    /// Add a quadratic Bezier segment.
    QuadraticTo {
        /// Control point.
        control: SvgPoint,
        /// Destination.
        to: SvgPoint,
    },
    /// Add an elliptical arc.
    ArcTo {
        /// Horizontal radius.
        rx: f64,
        /// Vertical radius.
        ry: f64,
        /// Ellipse-axis rotation in degrees.
        rotation: f64,
        /// Select the larger arc.
        large_arc: bool,
        /// Select positive-angle sweep.
        sweep: bool,
        /// Destination.
        to: SvgPoint,
    },
    /// Close the current subpath.
    Close,
}

/// A typed SVG path.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct SvgPath {
    commands: Vec<SvgPathCommand>,
}

impl SvgPath {
    /// Create an empty path.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            commands: Vec::new(),
        }
    }

    /// Append one path command.
    pub fn push(&mut self, command: SvgPathCommand) {
        self.commands.push(command);
    }

    /// Return the typed commands.
    #[must_use]
    pub fn commands(&self) -> &[SvgPathCommand] {
        &self.commands
    }
}

/// A fill or stroke paint.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum SvgPaint {
    /// Do not paint this channel.
    None,
    /// Paint with a solid colour.
    Solid(SvgColor),
    /// Paint with a definition created by this builder.
    Reference(SvgId),
}

/// SVG line cap setting.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SvgLineCap {
    /// Butt cap.
    Butt,
    /// Round cap.
    Round,
    /// Square cap.
    Square,
}

/// SVG line join setting.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SvgLineJoin {
    /// Miter join.
    Miter,
    /// Round join.
    Round,
    /// Bevel join.
    Bevel,
}

/// Stroke configuration.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgStroke {
    /// Stroke paint.
    pub paint: SvgPaint,
    /// Stroke width.
    pub width: f64,
    /// Line cap.
    pub line_cap: SvgLineCap,
    /// Line join.
    pub line_join: SvgLineJoin,
    /// Optional dash lengths.
    pub dashes: Vec<f64>,
    /// Dash phase.
    pub dash_offset: f64,
}

/// Paint and common element attributes.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgStyle {
    /// Fill paint.
    pub fill: SvgPaint,
    /// Optional stroke.
    pub stroke: Option<SvgStroke>,
    /// Combined opacity in the inclusive range 0..=1.
    pub opacity: f64,
    /// EMF+ compositing request.
    pub compositing: SvgCompositingMode,
}

impl Default for SvgStyle {
    fn default() -> Self {
        Self {
            fill: SvgPaint::Solid(SvgColor::rgb(0, 0, 0)),
            stroke: None,
            opacity: 1.0,
            compositing: SvgCompositingMode::SourceOver,
        }
    }
}

/// EMF+ compositing modes supported by the playback model.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SvgCompositingMode {
    /// Standard alpha compositing, represented by normal SVG painting.
    SourceOver,
    /// Replace destination pixels; SVG has no safe equivalent in this emitter.
    SourceCopy,
}

/// One stop in a linear gradient.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct SvgGradientStop {
    /// Offset from 0 through 1.
    pub offset: f64,
    /// Stop colour.
    pub color: SvgColor,
}

/// Typed linear gradient definition.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgLinearGradient {
    /// Start point.
    pub start: SvgPoint,
    /// End point.
    pub end: SvgPoint,
    /// Gradient stops, emitted in supplied order.
    pub stops: Vec<SvgGradientStop>,
    /// Optional gradient transform.
    pub transform: Option<SvgTransform>,
}

/// Safe pattern content.  There is intentionally no raw XML variant.
#[derive(Clone, Debug, PartialEq)]
pub enum SvgPatternContent {
    /// A rectangle painted inside the pattern cell.
    Rect { rect: SvgRect, style: SvgStyle },
    /// A path painted inside the pattern cell.
    Path { path: SvgPath, style: SvgStyle },
}

/// Typed SVG pattern definition.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgPattern {
    /// Pattern cell bounds.
    pub bounds: SvgRect,
    /// Pattern-local content.
    pub content: Vec<SvgPatternContent>,
    /// Optional transform applied to the pattern.
    pub transform: Option<SvgTransform>,
}

/// A supported encoded image MIME type.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SvgImageMime {
    /// PNG image bytes.
    Png,
    /// JPEG image bytes.
    Jpeg,
    /// GIF image bytes.
    Gif,
    /// WebP image bytes.
    Webp,
}

impl SvgImageMime {
    fn as_str(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Webp => "image/webp",
        }
    }
}

/// Image source accepted by the emitter.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum SvgImageSource {
    /// Bytes encoded into a `data:` URI.
    Embedded {
        /// Image MIME type.
        mime: SvgImageMime,
        /// Encoded image bytes.
        bytes: Vec<u8>,
    },
    /// An external URI, intentionally rejected with a diagnostic.
    ExternalUri,
}

/// An image draw operation.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgImage {
    /// Image placement.
    pub rect: SvgRect,
    /// Image source.
    pub source: SvgImageSource,
    /// Optional transform.
    pub transform: Option<SvgTransform>,
    /// Image opacity in the inclusive range 0..=1.
    pub opacity: f64,
}

/// A text draw operation.  Text is always emitted as XML character data.
#[derive(Clone, Debug, PartialEq)]
pub struct SvgText {
    /// Text content.
    pub value: String,
    /// Text baseline position.
    pub origin: SvgPoint,
    /// Font family name, emitted as an escaped attribute.
    pub font_family: String,
    /// Font size.
    pub font_size: f64,
    /// Text style.
    pub style: SvgStyle,
    /// Optional transform.
    pub transform: Option<SvgTransform>,
}

/// State-preserving SVG emitter for typed EMF+ playback operations.
#[derive(Debug)]
pub struct SvgBuilder {
    limits: SvgLimits,
    width: f64,
    height: f64,
    fragment: SvgFragment,
    elements: usize,
    definitions: usize,
    next_id: usize,
}

impl SvgBuilder {
    /// Create a document emitter with explicit dimensions and resource limits.
    pub fn new(width: f64, height: f64, limits: SvgLimits) -> Result<Self> {
        limits.validate()?;
        validate_non_negative(width, "width")?;
        validate_non_negative(height, "height")?;
        Ok(Self {
            limits,
            width,
            height,
            fragment: SvgFragment::default(),
            elements: 0,
            definitions: 0,
            next_id: 0,
        })
    }

    /// Emit a path.
    pub fn path(
        &mut self,
        path: &SvgPath,
        style: &SvgStyle,
        transform: Option<SvgTransform>,
        clip: Option<&SvgId>,
    ) -> Result<()> {
        validate_path(path, self.limits)?;
        self.ensure_space(estimated_path_bytes(path.commands.len())?)?;
        let mut element = String::new();
        element.push_str("<path d=\"");
        write_path_data(&mut element, path)?;
        element.push('"');
        write_style(
            &mut element,
            style,
            &mut self.fragment.diagnostics,
            self.limits,
        )?;
        write_transform(&mut element, transform)?;
        write_clip(&mut element, clip);
        element.push_str("/>");
        self.push_body(element)
    }

    /// Emit a rectangle.
    pub fn rect(
        &mut self,
        rect: SvgRect,
        style: &SvgStyle,
        transform: Option<SvgTransform>,
        clip: Option<&SvgId>,
    ) -> Result<()> {
        let mut element = String::from("<rect");
        write_rect_attributes(&mut element, rect)?;
        write_style(
            &mut element,
            style,
            &mut self.fragment.diagnostics,
            self.limits,
        )?;
        write_transform(&mut element, transform)?;
        write_clip(&mut element, clip);
        element.push_str("/>");
        self.push_body(element)
    }

    /// Emit an ellipse bounded by `rect`.
    pub fn ellipse(
        &mut self,
        rect: SvgRect,
        style: &SvgStyle,
        transform: Option<SvgTransform>,
        clip: Option<&SvgId>,
    ) -> Result<()> {
        validate_rect(rect)?;
        let mut element = String::from("<ellipse cx=\"");
        write_number(&mut element, rect.x + rect.width / 2.0)?;
        element.push_str("\" cy=\"");
        write_number(&mut element, rect.y + rect.height / 2.0)?;
        element.push_str("\" rx=\"");
        write_number(&mut element, rect.width / 2.0)?;
        element.push_str("\" ry=\"");
        write_number(&mut element, rect.height / 2.0)?;
        element.push('"');
        write_style(
            &mut element,
            style,
            &mut self.fragment.diagnostics,
            self.limits,
        )?;
        write_transform(&mut element, transform)?;
        write_clip(&mut element, clip);
        element.push_str("/>");
        self.push_body(element)
    }

    /// Emit text as escaped character data.
    pub fn text(&mut self, text: &SvgText, clip: Option<&SvgId>) -> Result<()> {
        validate_point(text.origin)?;
        validate_non_negative(text.font_size, "font size")?;
        self.ensure_space(estimated_escaped_bytes(&text.value, &text.font_family)?)?;
        let mut element = String::from("<text x=\"");
        write_number(&mut element, text.origin.x)?;
        element.push_str("\" y=\"");
        write_number(&mut element, text.origin.y)?;
        element.push_str("\" font-family=\"");
        write_xml_escaped(&mut element, &text.font_family);
        element.push_str("\" font-size=\"");
        write_number(&mut element, text.font_size)?;
        element.push('"');
        write_style(
            &mut element,
            &text.style,
            &mut self.fragment.diagnostics,
            self.limits,
        )?;
        write_transform(&mut element, text.transform)?;
        write_clip(&mut element, clip);
        element.push('>');
        write_xml_escaped(&mut element, &text.value);
        element.push_str("</text>");
        self.push_body(element)
    }

    /// Emit a base64-embedded raster image.  External image sources are skipped.
    pub fn image(&mut self, image: &SvgImage, clip: Option<&SvgId>) -> Result<()> {
        validate_rect(image.rect)?;
        validate_opacity(image.opacity)?;
        let SvgImageSource::Embedded { mime, bytes } = &image.source else {
            self.push_diagnostic(SvgUnsupportedFeature::ExternalImage)?;
            return Ok(());
        };
        if bytes.len() > self.limits.max_image_bytes {
            return Err(limit("embedded image bytes"));
        }
        let encoded_len = base64_len(bytes.len())?;
        self.ensure_space(
            encoded_len
                .checked_add(128)
                .ok_or_else(|| limit("embedded image encoding"))?,
        )?;
        let mut element = String::new();
        element.push_str("<image");
        write_rect_attributes(&mut element, image.rect)?;
        element.push_str(" href=\"data:");
        element.push_str(mime.as_str());
        element.push_str(";base64,");
        base64::engine::general_purpose::STANDARD.encode_string(bytes, &mut element);
        element.push('"');
        write_opacity(&mut element, image.opacity)?;
        write_transform(&mut element, image.transform)?;
        write_clip(&mut element, clip);
        element.push_str("/>");
        self.push_body(element)
    }

    /// Define a linear gradient and return an ID usable as `SvgPaint::Reference`.
    pub fn define_linear_gradient(&mut self, gradient: &SvgLinearGradient) -> Result<SvgId> {
        validate_point(gradient.start)?;
        validate_point(gradient.end)?;
        if gradient.stops.is_empty() {
            return Err(invalid("linear gradient needs at least one stop"));
        }
        if gradient.stops.len() > self.limits.max_path_commands {
            return Err(limit("gradient stop count"));
        }
        self.ensure_space(
            gradient
                .stops
                .len()
                .checked_mul(128)
                .ok_or_else(|| limit("gradient stop count"))?,
        )?;
        let id = self.next_definition_id("gradient")?;
        let mut definition = String::from("<linearGradient id=\"");
        definition.push_str(id.as_str());
        definition.push_str("\" x1=\"");
        write_number(&mut definition, gradient.start.x)?;
        definition.push_str("\" y1=\"");
        write_number(&mut definition, gradient.start.y)?;
        definition.push_str("\" x2=\"");
        write_number(&mut definition, gradient.end.x)?;
        definition.push_str("\" y2=\"");
        write_number(&mut definition, gradient.end.y)?;
        definition.push('"');
        write_transform_attribute(&mut definition, "gradientTransform", gradient.transform)?;
        definition.push('>');
        for stop in &gradient.stops {
            validate_unit_interval(stop.offset, "gradient stop offset")?;
            definition.push_str("<stop offset=\"");
            write_number(&mut definition, stop.offset)?;
            definition.push_str("\" stop-color=\"");
            write_color(&mut definition, stop.color);
            definition.push('"');
            write_alpha(&mut definition, "stop-opacity", stop.color.alpha);
            definition.push_str("/>");
        }
        definition.push_str("</linearGradient>");
        self.push_definition(definition)?;
        Ok(id)
    }

    /// Define a rectangular clip region and return its generated ID.
    pub fn define_clip_rect(
        &mut self,
        rect: SvgRect,
        transform: Option<SvgTransform>,
    ) -> Result<SvgId> {
        let id = self.next_definition_id("clip")?;
        let mut definition = String::from("<clipPath id=\"");
        definition.push_str(id.as_str());
        definition.push_str("\"><rect");
        write_rect_attributes(&mut definition, rect)?;
        write_transform(&mut definition, transform)?;
        definition.push_str("/></clipPath>");
        self.push_definition(definition)?;
        Ok(id)
    }

    /// Define a path clip region and return its generated ID.
    pub fn define_clip_path(
        &mut self,
        path: &SvgPath,
        transform: Option<SvgTransform>,
    ) -> Result<SvgId> {
        validate_path(path, self.limits)?;
        self.ensure_space(estimated_path_bytes(path.commands.len())?)?;
        let id = self.next_definition_id("clip")?;
        let mut definition = String::from("<clipPath id=\"");
        definition.push_str(id.as_str());
        definition.push_str("\"><path d=\"");
        write_path_data(&mut definition, path)?;
        definition.push('"');
        write_transform(&mut definition, transform)?;
        definition.push_str("/></clipPath>");
        self.push_definition(definition)?;
        Ok(id)
    }

    /// Define a pattern containing only typed, safe geometry.
    pub fn define_pattern(&mut self, pattern: &SvgPattern) -> Result<SvgId> {
        validate_rect(pattern.bounds)?;
        if pattern.content.len() > self.limits.max_elements {
            return Err(limit("pattern content count"));
        }
        self.ensure_space(
            pattern
                .content
                .len()
                .checked_mul(192)
                .ok_or_else(|| limit("pattern content count"))?,
        )?;
        let id = self.next_definition_id("pattern")?;
        let mut definition = String::from("<pattern id=\"");
        definition.push_str(id.as_str());
        definition.push_str("\" patternUnits=\"userSpaceOnUse\"");
        write_rect_attributes(&mut definition, pattern.bounds)?;
        write_transform_attribute(&mut definition, "patternTransform", pattern.transform)?;
        definition.push('>');
        for content in &pattern.content {
            write_pattern_content(
                &mut definition,
                content,
                &mut self.fragment.diagnostics,
                self.limits,
            )?;
        }
        definition.push_str("</pattern>");
        self.push_definition(definition)?;
        Ok(id)
    }

    /// Consume the builder and create a standalone SVG document.
    pub fn finish(self) -> Result<SvgDocument> {
        let mut svg = String::new();
        checked_push(&mut svg, "<svg xmlns=\"", self.limits.max_output_bytes)?;
        checked_push(&mut svg, SVG_NAMESPACE, self.limits.max_output_bytes)?;
        checked_push(&mut svg, "\" xmlns:xlink=\"", self.limits.max_output_bytes)?;
        checked_push(&mut svg, XLINK_NAMESPACE, self.limits.max_output_bytes)?;
        checked_push(&mut svg, "\" width=\"", self.limits.max_output_bytes)?;
        write_number(&mut svg, self.width)?;
        checked_push(&mut svg, "\" height=\"", self.limits.max_output_bytes)?;
        write_number(&mut svg, self.height)?;
        checked_push(&mut svg, "\" viewBox=\"0 0 ", self.limits.max_output_bytes)?;
        write_number(&mut svg, self.width)?;
        svg.push(' ');
        write_number(&mut svg, self.height)?;
        checked_push(&mut svg, "\">", self.limits.max_output_bytes)?;
        if !self.fragment.defs.is_empty() {
            checked_push(&mut svg, "<defs>", self.limits.max_output_bytes)?;
            checked_push(&mut svg, &self.fragment.defs, self.limits.max_output_bytes)?;
            checked_push(&mut svg, "</defs>", self.limits.max_output_bytes)?;
        }
        checked_push(&mut svg, &self.fragment.body, self.limits.max_output_bytes)?;
        checked_push(&mut svg, "</svg>", self.limits.max_output_bytes)?;
        Ok(SvgDocument {
            svg,
            fragment: self.fragment,
        })
    }

    fn next_definition_id(&mut self, kind: &str) -> Result<SvgId> {
        if self.definitions >= self.limits.max_definitions {
            return Err(limit("definition count"));
        }
        self.next_id = self
            .next_id
            .checked_add(1)
            .ok_or_else(|| limit("definition ID"))?;
        Ok(SvgId(format!("emfplus-{kind}-{}", self.next_id)))
    }

    #[allow(
        clippy::needless_pass_by_value,
        reason = "the completed element is consumed as one bounded emission unit"
    )]
    fn push_body(&mut self, element: String) -> Result<()> {
        if self.elements >= self.limits.max_elements {
            return Err(limit("element count"));
        }
        self.ensure_space(element.len())?;
        checked_push(
            &mut self.fragment.body,
            &element,
            self.limits.max_output_bytes,
        )?;
        self.elements += 1;
        Ok(())
    }

    #[allow(
        clippy::needless_pass_by_value,
        reason = "the completed definition is consumed as one bounded emission unit"
    )]
    fn push_definition(&mut self, definition: String) -> Result<()> {
        self.ensure_space(definition.len())?;
        checked_push(
            &mut self.fragment.defs,
            &definition,
            self.limits.max_output_bytes,
        )?;
        self.definitions += 1;
        Ok(())
    }

    fn push_diagnostic(&mut self, feature: SvgUnsupportedFeature) -> Result<()> {
        if self.fragment.diagnostics.len() >= self.limits.max_diagnostics {
            return Err(limit("diagnostic count"));
        }
        self.fragment
            .diagnostics
            .try_reserve(1)
            .map_err(allocation)?;
        self.fragment.diagnostics.push(SvgDiagnostic { feature });
        Ok(())
    }

    fn ensure_space(&self, additional: usize) -> Result<()> {
        let content = self
            .fragment
            .defs
            .len()
            .checked_add(self.fragment.body.len())
            .and_then(|value| value.checked_add(additional))
            .ok_or_else(|| limit("output bytes"))?;
        if content > self.limits.max_output_bytes {
            return Err(limit("output bytes"));
        }
        Ok(())
    }
}

fn write_pattern_content(
    output: &mut String,
    content: &SvgPatternContent,
    diagnostics: &mut Vec<SvgDiagnostic>,
    limits: SvgLimits,
) -> Result<()> {
    match content {
        SvgPatternContent::Rect { rect, style } => {
            output.push_str("<rect");
            write_rect_attributes(output, *rect)?;
            write_style(output, style, diagnostics, limits)?;
            output.push_str("/>");
        },
        SvgPatternContent::Path { path, style } => {
            validate_path(path, limits)?;
            output.push_str("<path d=\"");
            write_path_data(output, path)?;
            output.push('"');
            write_style(output, style, diagnostics, limits)?;
            output.push_str("/>");
        },
    }
    Ok(())
}

fn write_path_data(output: &mut String, path: &SvgPath) -> Result<()> {
    for command in &path.commands {
        match command {
            SvgPathCommand::MoveTo(point) => write_point_command(output, 'M', *point)?,
            SvgPathCommand::LineTo(point) => write_point_command(output, 'L', *point)?,
            SvgPathCommand::CubicTo { first, second, to } => {
                output.push('C');
                write_point(output, *first)?;
                output.push(' ');
                write_point(output, *second)?;
                output.push(' ');
                write_point(output, *to)?;
            },
            SvgPathCommand::QuadraticTo { control, to } => {
                output.push('Q');
                write_point(output, *control)?;
                output.push(' ');
                write_point(output, *to)?;
            },
            SvgPathCommand::ArcTo {
                rx,
                ry,
                rotation,
                large_arc,
                sweep,
                to,
            } => {
                validate_non_negative(*rx, "arc horizontal radius")?;
                validate_non_negative(*ry, "arc vertical radius")?;
                validate_finite(*rotation, "arc rotation")?;
                output.push('A');
                write_number(output, *rx)?;
                output.push(' ');
                write_number(output, *ry)?;
                output.push(' ');
                write_number(output, *rotation)?;
                output.push(' ');
                output.push(if *large_arc { '1' } else { '0' });
                output.push(' ');
                output.push(if *sweep { '1' } else { '0' });
                output.push(' ');
                write_point(output, *to)?;
            },
            SvgPathCommand::Close => output.push('Z'),
        }
    }
    Ok(())
}

fn write_point_command(output: &mut String, command: char, point: SvgPoint) -> Result<()> {
    output.push(command);
    write_point(output, point)
}

fn write_point(output: &mut String, point: SvgPoint) -> Result<()> {
    validate_point(point)?;
    write_number(output, point.x)?;
    output.push(' ');
    write_number(output, point.y)
}

fn write_rect_attributes(output: &mut String, rect: SvgRect) -> Result<()> {
    validate_rect(rect)?;
    output.push_str(" x=\"");
    write_number(output, rect.x)?;
    output.push_str("\" y=\"");
    write_number(output, rect.y)?;
    output.push_str("\" width=\"");
    write_number(output, rect.width)?;
    output.push_str("\" height=\"");
    write_number(output, rect.height)?;
    output.push('"');
    Ok(())
}

fn write_style(
    output: &mut String,
    style: &SvgStyle,
    diagnostics: &mut Vec<SvgDiagnostic>,
    limits: SvgLimits,
) -> Result<()> {
    let fill_alpha = write_paint(output, "fill", &style.fill);
    if let Some(alpha) = fill_alpha {
        write_alpha(output, "fill-opacity", alpha);
    }
    validate_opacity(style.opacity)?;
    write_opacity(output, style.opacity)?;
    if let Some(stroke) = &style.stroke {
        if stroke.dashes.len() > limits.max_path_commands {
            return Err(limit("dash count"));
        }
        validate_non_negative(stroke.width, "stroke width")?;
        let stroke_alpha = write_paint(output, "stroke", &stroke.paint);
        if let Some(alpha) = stroke_alpha {
            write_alpha(output, "stroke-opacity", alpha);
        }
        output.push_str(" stroke-width=\"");
        write_number(output, stroke.width)?;
        output.push('"');
        output.push_str(" stroke-linecap=\"");
        output.push_str(line_cap_name(stroke.line_cap));
        output.push_str("\" stroke-linejoin=\"");
        output.push_str(line_join_name(stroke.line_join));
        output.push('"');
        if !stroke.dashes.is_empty() {
            output.push_str(" stroke-dasharray=\"");
            for (index, dash) in stroke.dashes.iter().enumerate() {
                validate_non_negative(*dash, "dash length")?;
                if index != 0 {
                    output.push(' ');
                }
                write_number(output, *dash)?;
            }
            output.push('"');
            validate_finite(stroke.dash_offset, "dash offset")?;
            output.push_str(" stroke-dashoffset=\"");
            write_number(output, stroke.dash_offset)?;
            output.push('"');
        }
    }
    if style.compositing == SvgCompositingMode::SourceCopy {
        push_diagnostic(
            diagnostics,
            limits,
            SvgUnsupportedFeature::SourceCopyCompositing,
        )?;
    }
    Ok(())
}

fn write_paint(output: &mut String, name: &str, paint: &SvgPaint) -> Option<u8> {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    match paint {
        SvgPaint::None => output.push_str("none"),
        SvgPaint::Solid(color) => write_color(output, *color),
        SvgPaint::Reference(id) => {
            output.push_str("url(#");
            output.push_str(id.as_str());
            output.push(')');
        },
    }
    output.push('"');
    match paint {
        SvgPaint::Solid(color) => Some(color.alpha),
        SvgPaint::None | SvgPaint::Reference(_) => None,
    }
}

fn write_color(output: &mut String, color: SvgColor) {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    output.push('#');
    for channel in [color.red, color.green, color.blue] {
        output.push(char::from(HEX[usize::from(channel >> 4)]));
        output.push(char::from(HEX[usize::from(channel & 0x0f)]));
    }
}

fn write_alpha(output: &mut String, name: &str, alpha: u8) {
    if alpha != u8::MAX {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        let mut buffer = ryu::Buffer::new();
        output.push_str(buffer.format(f64::from(alpha) / f64::from(u8::MAX)));
        output.push('"');
    }
}

fn write_opacity(output: &mut String, opacity: f64) -> Result<()> {
    validate_opacity(opacity)?;
    if opacity < 1.0 {
        output.push_str(" opacity=\"");
        write_number(output, opacity)?;
        output.push('"');
    }
    Ok(())
}

fn write_transform(output: &mut String, transform: Option<SvgTransform>) -> Result<()> {
    write_transform_attribute(output, "transform", transform)
}

fn write_transform_attribute(
    output: &mut String,
    name: &str,
    transform: Option<SvgTransform>,
) -> Result<()> {
    let Some(transform) = transform else {
        return Ok(());
    };
    for value in [
        transform.a,
        transform.b,
        transform.c,
        transform.d,
        transform.e,
        transform.f,
    ] {
        validate_finite(value, "transform coefficient")?;
    }
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"matrix(");
    for (index, value) in [
        transform.a,
        transform.b,
        transform.c,
        transform.d,
        transform.e,
        transform.f,
    ]
    .iter()
    .enumerate()
    {
        if index != 0 {
            output.push(' ');
        }
        write_number(output, *value)?;
    }
    output.push_str(")\"");
    Ok(())
}

fn write_clip(output: &mut String, clip: Option<&SvgId>) {
    if let Some(clip) = clip {
        output.push_str(" clip-path=\"url(#");
        output.push_str(clip.as_str());
        output.push_str(")\"");
    }
}

fn line_cap_name(line_cap: SvgLineCap) -> &'static str {
    match line_cap {
        SvgLineCap::Butt => "butt",
        SvgLineCap::Round => "round",
        SvgLineCap::Square => "square",
    }
}

fn line_join_name(line_join: SvgLineJoin) -> &'static str {
    match line_join {
        SvgLineJoin::Miter => "miter",
        SvgLineJoin::Round => "round",
        SvgLineJoin::Bevel => "bevel",
    }
}

fn write_xml_escaped(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{10ffff}' => {
                output.push(character);
            },
            _ => output.push('\u{fffd}'),
        }
    }
}

fn write_number(output: &mut String, value: f64) -> Result<()> {
    validate_finite(value, "number")?;
    let canonical = if value == 0.0 { 0.0 } else { value };
    let mut buffer = ryu::Buffer::new();
    let formatted = buffer.format(canonical);
    output.push_str(formatted.strip_suffix(".0").unwrap_or(formatted));
    Ok(())
}

fn validate_rect(rect: SvgRect) -> Result<()> {
    validate_finite(rect.x, "rectangle x")?;
    validate_finite(rect.y, "rectangle y")?;
    validate_non_negative(rect.width, "rectangle width")?;
    validate_non_negative(rect.height, "rectangle height")
}

fn validate_path(path: &SvgPath, limits: SvgLimits) -> Result<()> {
    if path.commands.len() > limits.max_path_commands {
        return Err(limit("path command count"));
    }
    Ok(())
}

fn validate_point(point: SvgPoint) -> Result<()> {
    validate_finite(point.x, "point x")?;
    validate_finite(point.y, "point y")
}

fn validate_opacity(opacity: f64) -> Result<()> {
    validate_unit_interval(opacity, "opacity")
}

fn validate_unit_interval(value: f64, name: &str) -> Result<()> {
    validate_finite(value, name)?;
    if !(0.0..=1.0).contains(&value) {
        return Err(invalid(&format!("{name} must be in 0..=1")));
    }
    Ok(())
}

fn validate_non_negative(value: f64, name: &str) -> Result<()> {
    validate_finite(value, name)?;
    if value < 0.0 {
        return Err(invalid(&format!("{name} must be non-negative")));
    }
    Ok(())
}

fn validate_finite(value: f64, name: &str) -> Result<()> {
    if !value.is_finite() {
        return Err(invalid(&format!("{name} must be finite")));
    }
    Ok(())
}

fn checked_push(output: &mut String, value: &str, maximum: usize) -> Result<()> {
    let length = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| limit("output bytes"))?;
    if length > maximum {
        return Err(limit("output bytes"));
    }
    output.try_reserve(value.len()).map_err(allocation)?;
    output.push_str(value);
    Ok(())
}

fn base64_len(bytes: usize) -> Result<usize> {
    bytes
        .checked_add(2)
        .and_then(|value| value.checked_div(3))
        .and_then(|value| value.checked_mul(4))
        .ok_or_else(|| limit("embedded image encoding"))
}

fn estimated_path_bytes(commands: usize) -> Result<usize> {
    commands
        .checked_mul(192)
        .ok_or_else(|| limit("path command count"))
}

fn estimated_escaped_bytes(value: &str, family: &str) -> Result<usize> {
    value
        .len()
        .checked_add(family.len())
        .and_then(|length| length.checked_mul(6))
        .and_then(|length| length.checked_add(256))
        .ok_or_else(|| limit("text bytes"))
}

fn push_diagnostic(
    diagnostics: &mut Vec<SvgDiagnostic>,
    limits: SvgLimits,
    feature: SvgUnsupportedFeature,
) -> Result<()> {
    if diagnostics.len() >= limits.max_diagnostics {
        return Err(limit("diagnostic count"));
    }
    diagnostics.try_reserve(1).map_err(allocation)?;
    diagnostics.push(SvgDiagnostic { feature });
    Ok(())
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}

fn limit(resource: &str) -> Error {
    Error::CorruptedFile(format!("EMF+ SVG {resource} exceeds configured limit"))
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "EMF+ SVG output",
        source,
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "unit tests use unwrap for concise assertions"
)]
mod tests {
    use super::{
        SvgBuilder, SvgColor, SvgCompositingMode, SvgImage, SvgImageSource, SvgLimits, SvgPaint,
        SvgPath, SvgPathCommand, SvgPoint, SvgRect, SvgStyle, SvgText, SvgTransform,
        SvgUnsupportedFeature,
    };

    fn builder() -> SvgBuilder {
        SvgBuilder::new(100.0, 50.0, SvgLimits::default()).unwrap()
    }

    #[test]
    fn text_and_attributes_are_escaped() {
        let mut svg = builder();
        svg.text(
            &SvgText {
                value: "<script>&'\"\u{1}".to_owned(),
                origin: SvgPoint { x: 1.0, y: 2.0 },
                font_family: "a\" onclick=\"evil".to_owned(),
                font_size: 12.0,
                style: SvgStyle::default(),
                transform: None,
            },
            None,
        )
        .unwrap();
        let result = svg.finish().unwrap().into_string();
        assert!(result.contains("&lt;script&gt;&amp;&apos;&quot;�"));
        assert!(result.contains("font-family=\"a&quot; onclick=&quot;evil\""));
        assert!(!result.contains("<script>"));
    }

    #[test]
    fn non_finite_numbers_are_rejected_and_negative_zero_is_canonical() {
        let mut svg = builder();
        let error = svg
            .rect(
                SvgRect {
                    x: f64::NAN,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0,
                },
                &SvgStyle::default(),
                None,
                None,
            )
            .unwrap_err();
        assert!(error.to_string().contains("finite"));
        svg.rect(
            SvgRect {
                x: -0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            &SvgStyle::default(),
            None,
            None,
        )
        .unwrap();
        assert!(svg.finish().unwrap().as_str().contains("x=\"0\""));
    }

    #[test]
    fn output_and_element_limits_are_enforced() {
        let limits = SvgLimits {
            max_output_bytes: 300,
            max_elements: 1,
            ..SvgLimits::default()
        };
        let mut svg = SvgBuilder::new(1.0, 1.0, limits).unwrap();
        svg.rect(
            SvgRect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            &SvgStyle::default(),
            None,
            None,
        )
        .unwrap();
        assert!(
            svg.rect(
                SvgRect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0
                },
                &SvgStyle::default(),
                None,
                None,
            )
            .is_err()
        );
        assert!(svg.finish().is_ok());
    }

    #[test]
    fn generated_ids_are_valid_and_definition_references_are_safe() {
        let mut svg = builder();
        let first = svg
            .define_clip_rect(
                SvgRect {
                    x: 0.0,
                    y: 0.0,
                    width: 10.0,
                    height: 10.0,
                },
                Some(SvgTransform::identity()),
            )
            .unwrap();
        let second = svg
            .define_clip_rect(
                SvgRect {
                    x: 1.0,
                    y: 1.0,
                    width: 8.0,
                    height: 8.0,
                },
                None,
            )
            .unwrap();
        assert_eq!(first.as_str(), "emfplus-clip-1");
        assert_eq!(second.as_str(), "emfplus-clip-2");
        svg.rect(
            SvgRect {
                x: 0.0,
                y: 0.0,
                width: 2.0,
                height: 2.0,
            },
            &SvgStyle::default(),
            None,
            Some(&first),
        )
        .unwrap();
        let result = svg.finish().unwrap().into_string();
        assert!(result.contains("id=\"emfplus-clip-1\""));
        assert!(result.contains("clip-path=\"url(#emfplus-clip-1)\""));
    }

    #[test]
    fn paths_emit_only_typed_commands_and_validate_arc_radii() {
        let mut path = SvgPath::new();
        path.push(SvgPathCommand::MoveTo(SvgPoint { x: 1.0, y: 2.0 }));
        path.push(SvgPathCommand::LineTo(SvgPoint { x: 3.0, y: 4.0 }));
        path.push(SvgPathCommand::Close);
        let mut svg = builder();
        svg.path(&path, &SvgStyle::default(), None, None).unwrap();
        assert!(svg.finish().unwrap().as_str().contains("d=\"M1 2L3 4Z\""));
        let mut invalid_path = SvgPath::new();
        invalid_path.push(SvgPathCommand::ArcTo {
            rx: -1.0,
            ry: 1.0,
            rotation: 0.0,
            large_arc: false,
            sweep: false,
            to: SvgPoint { x: 0.0, y: 0.0 },
        });
        assert!(
            builder()
                .path(&invalid_path, &SvgStyle::default(), None, None)
                .is_err()
        );
    }

    #[test]
    fn unsupported_external_images_and_source_copy_are_diagnostic_only() {
        let mut svg = builder();
        svg.image(
            &SvgImage {
                rect: SvgRect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0,
                },
                source: SvgImageSource::ExternalUri,
                transform: None,
                opacity: 1.0,
            },
            None,
        )
        .unwrap();
        let style = SvgStyle {
            compositing: SvgCompositingMode::SourceCopy,
            ..SvgStyle::default()
        };
        svg.rect(
            SvgRect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            &style,
            None,
            None,
        )
        .unwrap();
        let document = svg.finish().unwrap();
        assert_eq!(document.fragment().diagnostics().len(), 2);
        assert_eq!(
            document.fragment().diagnostics()[0].feature,
            SvgUnsupportedFeature::ExternalImage
        );
        assert_eq!(
            document.fragment().diagnostics()[1].feature,
            SvgUnsupportedFeature::SourceCopyCompositing
        );
    }

    #[test]
    fn image_data_is_embedded_without_external_uri() {
        let mut svg = builder();
        svg.image(
            &SvgImage {
                rect: SvgRect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0,
                },
                source: SvgImageSource::Embedded {
                    mime: super::SvgImageMime::Png,
                    bytes: vec![0, 1, 2],
                },
                transform: None,
                opacity: 1.0,
            },
            None,
        )
        .unwrap();
        assert!(
            svg.finish()
                .unwrap()
                .as_str()
                .contains("data:image/png;base64,AAEC")
        );
    }

    #[test]
    fn definition_and_path_limits_are_enforced() {
        let limits = SvgLimits {
            max_definitions: 1,
            max_path_commands: 1,
            ..SvgLimits::default()
        };
        let mut svg = SvgBuilder::new(1.0, 1.0, limits).unwrap();
        svg.define_clip_rect(
            SvgRect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            None,
        )
        .unwrap();
        assert!(
            svg.define_clip_rect(
                SvgRect {
                    x: 0.0,
                    y: 0.0,
                    width: 1.0,
                    height: 1.0
                },
                None
            )
            .is_err()
        );
        let mut path = SvgPath::new();
        path.push(SvgPathCommand::MoveTo(SvgPoint { x: 0.0, y: 0.0 }));
        path.push(SvgPathCommand::Close);
        assert!(svg.path(&path, &SvgStyle::default(), None, None).is_err());
    }

    #[test]
    fn paint_is_not_user_supplied_xml() {
        let mut svg = builder();
        let style = SvgStyle {
            fill: SvgPaint::Solid(SvgColor::rgba(1, 2, 3, 4)),
            ..SvgStyle::default()
        };
        svg.rect(
            SvgRect {
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
            },
            &style,
            None,
            None,
        )
        .unwrap();
        let output = svg.finish().unwrap().into_string();
        assert!(output.contains("fill=\"#010203\" fill-opacity=\""));
    }
}
