//! Watermark support for DOCX documents.
//!
//! Based on Apache POI's XWPFHeaderFooterPolicy watermark implementation.
use crate::error::{Error, Result};
use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";

/// Rotation applied to diagonal text watermarks (`rotation:315`).
const DIAGONAL_ROTATION: &str = "315";
/// Fill opacity applied to semitransparent text watermarks.
const SEMITRANSPARENT_FILL_OPACITY: &str = ".5";
/// Z-index of image watermark shapes (behind document content).
const IMAGE_WATERMARK_Z_INDEX: i64 = -251654144;

/// Layout orientation of a text watermark.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum WatermarkLayout {
    /// Diagonal text (`rotation:315`, the Word default).
    #[default]
    Diagonal,
    /// Horizontal text (no rotation).
    Horizontal,
}

/// A watermark for a Word document.
///
/// Watermarks are typically displayed as diagonal text behind the document content.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi_docx::writer::Watermark;
///
/// let watermark = Watermark::text("CONFIDENTIAL");
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Watermark {
    /// Watermark text
    text: String,
    /// Font family (default: Cambria)
    font: String,
    /// Font size in points (default: 1pt for shape-based watermarks)
    font_size: f64,
    /// Color (RGB hex, default: "000000" black)
    color: String,
    /// Layout orientation (default: diagonal)
    layout: WatermarkLayout,
    /// Whether the fill is semitransparent (default: true)
    semitransparent: bool,
}

impl Watermark {
    /// Create a text watermark.
    ///
    /// # Arguments
    ///
    /// * `text` - Text to display as watermark
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let wm = Watermark::text("CONFIDENTIAL");
    /// ```
    pub fn text(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            font: "Cambria".to_string(),
            font_size: 1.0,
            color: "black".to_string(),
            layout: WatermarkLayout::Diagonal,
            semitransparent: true,
        }
    }

    /// Set the font family.
    pub fn set_font(&mut self, font: impl Into<String>) {
        self.font = font.into();
    }

    /// Set the font size in points.
    pub fn set_font_size(&mut self, size: u32) {
        self.font_size = f64::from(size);
    }

    /// Set a potentially fractional VML text-path font size in points.
    pub fn set_font_size_points(&mut self, size: f64) -> Result<()> {
        if !size.is_finite() || size < 0.0 {
            return Err(Error::InvalidFormat(
                "watermark font size must be finite and non-negative".into(),
            ));
        }
        self.font_size = size;
        Ok(())
    }

    /// Set the color.
    pub fn set_color(&mut self, color: impl Into<String>) {
        self.color = color.into();
    }

    /// Set the layout orientation (diagonal or horizontal).
    pub fn set_layout(&mut self, layout: WatermarkLayout) {
        self.layout = layout;
    }

    /// Set whether the watermark fill is semitransparent.
    pub fn set_semitransparent(&mut self, semitransparent: bool) {
        self.semitransparent = semitransparent;
    }

    /// Get the watermark text.
    #[inline]
    pub fn get_text(&self) -> &str {
        &self.text
    }

    /// Get the watermark font family.
    #[inline]
    pub fn font(&self) -> &str {
        &self.font
    }

    /// Get the VML text-path font size in points.
    #[inline]
    pub fn font_size(&self) -> f64 {
        self.font_size
    }

    /// Get the VML fill color.
    #[inline]
    pub fn color(&self) -> &str {
        &self.color
    }

    /// Get the layout orientation.
    #[inline]
    pub fn layout(&self) -> WatermarkLayout {
        self.layout
    }

    /// Get whether the watermark fill is semitransparent.
    #[inline]
    pub fn semitransparent(&self) -> bool {
        self.semitransparent
    }

    /// Parse text watermarks from a Word header part.
    pub(crate) fn from_header_xml(xml: &[u8]) -> Result<Vec<Self>> {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut shapes: Vec<(usize, Option<WatermarkCandidate>)> = Vec::new();
        let mut watermarks = Vec::new();

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word header XML nesting is too deep".into())
                    })?;
                    if depth > 256 {
                        return Err(Error::InvalidFormat(
                            "Word header XML nesting exceeds 256 elements".into(),
                        ));
                    }
                    if is_vml(&namespace) && element.local_name().as_ref() == b"shape" {
                        shapes.push((
                            depth,
                            WatermarkCandidate::from_shape(&element, decoder, &resolver)?,
                        ));
                    } else if is_vml(&namespace) && element.local_name().as_ref() == b"textpath" {
                        apply_text_path(&mut shapes, &element, decoder, &resolver)?;
                    } else if is_vml(&namespace) && element.local_name().as_ref() == b"fill" {
                        apply_fill(&mut shapes, &element, decoder, &resolver)?;
                    }
                },
                Event::Empty(element)
                    if is_vml(&namespace) && element.local_name().as_ref() == b"textpath" =>
                {
                    apply_text_path(&mut shapes, &element, decoder, &resolver)?;
                },
                Event::Empty(element)
                    if is_vml(&namespace) && element.local_name().as_ref() == b"fill" =>
                {
                    apply_fill(&mut shapes, &element, decoder, &resolver)?;
                },
                Event::End(element)
                    if is_vml(&namespace) && element.local_name().as_ref() == b"shape" =>
                {
                    let Some((shape_depth, candidate)) = shapes.pop() else {
                        return Err(Error::InvalidFormat(
                            "VML shape end has no matching start".into(),
                        ));
                    };
                    if shape_depth != depth {
                        return Err(Error::InvalidFormat(
                            "invalid nested VML shape structure".into(),
                        ));
                    }
                    if let Some(watermark) = candidate.and_then(WatermarkCandidate::finish) {
                        watermarks.push(watermark);
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word header XML nesting".into())
                    })?;
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word header XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 || !shapes.is_empty() => {
                    return Err(Error::InvalidFormat("unterminated Word header XML".into()));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(watermarks)
    }

    /// Generate watermark XML for a header paragraph.
    ///
    /// This creates a VML shape-based watermark that appears behind document content.
    /// The watermark is embedded in a header paragraph.
    ///
    /// # Arguments
    ///
    /// * `idx` - Index for unique ID generation (1, 2, or 3 for default/first/even)
    pub(crate) fn to_header_xml(&self, idx: u32) -> Result<String> {
        let mut xml = String::with_capacity(2048);

        // Start paragraph with Header style
        xml.push_str(r#"<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>"#);

        // Start run with picture
        xml.push_str(r#"<w:r><w:rPr><w:noProof/></w:rPr><w:pict>"#);

        // Shapetype for diagonal text (NO v:group wrapper for watermarks!)
        xml.push_str(r#"<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">"#);

        // Formulas for shape
        xml.push_str(r#"<v:formulas>"#);
        xml.push_str(r#"<v:f eqn="sum #0 0 10800"/>"#);
        xml.push_str(r#"<v:f eqn="prod #0 2 1"/>"#);
        xml.push_str(r#"<v:f eqn="sum 21600 0 @1"/>"#);
        xml.push_str(r#"<v:f eqn="sum 0 0 @2"/>"#);
        xml.push_str(r#"<v:f eqn="sum 21600 0 @3"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 @3 0"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 21600 @1"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 0 @2"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 @4 21600"/>"#);
        xml.push_str(r#"<v:f eqn="mid @5 @6"/>"#);
        xml.push_str(r#"<v:f eqn="mid @8 @5"/>"#);
        xml.push_str(r#"<v:f eqn="mid @7 @8"/>"#);
        xml.push_str(r#"<v:f eqn="mid @6 @7"/>"#);
        xml.push_str(r#"<v:f eqn="sum @6 0 @5"/>"#);
        xml.push_str(r#"</v:formulas>"#);

        xml.push_str(r#"<v:path textpathok="t" o:connecttype="custom" "#);
        xml.push_str(r#"o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" "#);
        xml.push_str(r#"o:connectangles="270,180,90,0"/>"#);

        xml.push_str(r#"<v:textpath on="t" fitshape="t"/>"#);

        xml.push_str(r#"<v:handles>"#);
        xml.push_str(r##"<v:h position="#0,bottomRight" xrange="6629,14971"/>"##);
        xml.push_str(r#"</v:handles>"#);

        xml.push_str(r#"<o:lock v:ext="edit" text="t" shapetype="t"/>"#);

        xml.push_str(r#"</v:shapetype>"#);

        // Main shape
        xml.push_str(&format!(
            r#"<v:shape id="PowerPlusWaterMarkObject{}" "#,
            idx
        ));
        xml.push_str(&format!(r##"o:spid="_x0000_s102{}" "##, 4 + idx));
        xml.push_str(r##"type="#_x0000_t136" "##);
        xml.push_str(
            r#"style="position:absolute;margin-left:0;margin-top:0;width:439.9pt;height:219.95pt;"#,
        );
        if self.layout == WatermarkLayout::Diagonal {
            xml.push_str("rotation:");
            xml.push_str(DIAGONAL_ROTATION);
            xml.push(';');
        }
        xml.push_str("z-index:-251655168;");
        xml.push_str(r#"mso-position-horizontal:center;mso-position-horizontal-relative:margin;"#);
        xml.push_str(r#"mso-position-vertical:center;mso-position-vertical-relative:margin" "#);
        xml.push_str(r#"o:allowincell="f" "#);
        // Add # prefix to color if not already present
        let color_with_hash = if self.color.starts_with('#')
            || !self.color.bytes().all(|byte| byte.is_ascii_hexdigit())
        {
            self.color.clone()
        } else {
            format!("#{}", self.color)
        };
        xml.push_str(&format!(r#"fillcolor="{}" "#, color_with_hash));
        xml.push_str(r#"stroked="f">"#);

        // Fill with opacity (separate element, not attribute!)
        if self.semitransparent {
            xml.push_str(&format!(
                "<v:fill opacity=\"{SEMITRANSPARENT_FILL_OPACITY}\"/>"
            ));
        }

        // Text path
        // Shape dimensions control the visual scale; the text-path point size
        // is still preserved for producer compatibility.
        xml.push_str(&format!(
            r#"<v:textpath style="font-family:&quot;{}&quot;;font-size:{}pt" "#,
            self.font, self.font_size
        ));
        xml.push_str(&format!(r#"string="{}"/>"#, escape_xml(&self.text)));

        xml.push_str(r#"</v:shape>"#);

        // End picture and run
        xml.push_str(r#"</w:pict></w:r></w:p>"#);

        Ok(xml)
    }
}

#[derive(Default)]
struct WatermarkCandidate {
    color: Option<String>,
    text: Option<String>,
    font: Option<String>,
    font_size: Option<f64>,
    rotation: Option<f64>,
    fill_opacity: Option<f64>,
}

impl WatermarkCandidate {
    fn from_shape(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Option<Self>> {
        let id = unqualified_attribute(element, b"id", decoder, resolver)?;
        let shape_type = unqualified_attribute(element, b"type", decoder, resolver)?;
        let style = unqualified_attribute(element, b"style", decoder, resolver)?;
        let canonical_id = id
            .as_deref()
            .is_some_and(|value| value.starts_with("PowerPlusWaterMarkObject"));
        let background_word_art = shape_type.as_deref() == Some("#_x0000_t136")
            && style.as_deref().is_some_and(|value| {
                value.contains("z-index:-")
                    && value.contains("mso-position-horizontal:center")
                    && value.contains("mso-position-vertical:center")
            });
        if !canonical_id && !background_word_art {
            return Ok(None);
        }
        let rotation = style.as_deref().and_then(|value| {
            value.split(';').find_map(|declaration| {
                let (name, rotation) = declaration.split_once(':')?;
                if name.trim().eq_ignore_ascii_case("rotation") {
                    rotation.trim().parse::<f64>().ok()
                } else {
                    None
                }
            })
        });
        Ok(Some(Self {
            color: unqualified_attribute(element, b"fillcolor", decoder, resolver)?,
            rotation,
            ..Self::default()
        }))
    }

    fn finish(self) -> Option<Watermark> {
        // Word writes rotation:315 for diagonal watermarks and omits the
        // rotation declaration for horizontal ones.
        let layout = match self.rotation {
            Some(315.0) => WatermarkLayout::Diagonal,
            _ => WatermarkLayout::Horizontal,
        };
        Some(Watermark {
            text: self.text?,
            font: self.font.unwrap_or_else(|| "Cambria".to_string()),
            font_size: self.font_size.unwrap_or(1.0),
            color: self.color.unwrap_or_else(|| "black".to_string()),
            layout,
            semitransparent: self.fill_opacity.is_some_and(|opacity| opacity < 1.0),
        })
    }
}

fn is_vml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == VML_NAMESPACE)
}

fn apply_fill(
    shapes: &mut [(usize, Option<WatermarkCandidate>)],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let Some((_, Some(candidate))) = shapes.last_mut() else {
        return Ok(());
    };
    if let Some(opacity) = unqualified_attribute(element, b"opacity", decoder, resolver)?
        && let Ok(opacity) = opacity.parse::<f64>()
        && opacity.is_finite()
    {
        candidate.fill_opacity = Some(opacity);
    }
    Ok(())
}

fn apply_text_path(
    shapes: &mut [(usize, Option<WatermarkCandidate>)],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let Some((_, Some(candidate))) = shapes.last_mut() else {
        return Ok(());
    };
    candidate.text = unqualified_attribute(element, b"string", decoder, resolver)?;
    if let Some(style) = unqualified_attribute(element, b"style", decoder, resolver)? {
        for declaration in style.split(';') {
            let Some((name, value)) = declaration.split_once(':') else {
                continue;
            };
            match name.trim().to_ascii_lowercase().as_str() {
                "font-family" => {
                    let value = value.trim().trim_matches(['\'', '"']);
                    if !value.is_empty() {
                        candidate.font = Some(value.to_string());
                    }
                },
                "font-size" => {
                    let value = value.trim().strip_suffix("pt").unwrap_or(value.trim());
                    if let Ok(size) = value.parse::<f64>()
                        && size.is_finite()
                        && size >= 0.0
                    {
                        candidate.font_size = Some(size);
                    }
                },
                _ => {},
            }
        }
    }
    Ok(())
}

fn unqualified_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name
            || !matches!(
                resolver.resolve_attribute(attribute.key).0,
                ResolveResult::Unbound
            )
        {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate VML attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

/// Maximum image watermark payload size accepted by the authoring API.
pub const MAX_WATERMARK_IMAGE_BYTES: usize = 16 * 1024 * 1024;
/// Maximum image watermark scale factor accepted by the authoring API.
const MAX_WATERMARK_SCALE: f64 = 100.0;
/// Screen resolution used to convert image pixels to points.
const PIXELS_PER_INCH: f64 = 96.0;
/// Points per inch, converting pixel dimensions to point dimensions.
const POINTS_PER_INCH: f64 = 72.0;

/// A picture watermark being authored in a document.
///
/// The image is stored verbatim as an ordinary media part and referenced
/// from VML watermark shapes in the headers with centered default geometry.
/// The payload is never interpreted beyond format/dimension sniffing.
#[derive(Debug, Clone)]
pub struct ImageWatermark {
    /// Image bytes, stored verbatim.
    data: Vec<u8>,
    /// Sniffed image format.
    format: crate::format::ImageFormat,
    /// Width of the watermark shape in points.
    width_pt: f64,
    /// Height of the watermark shape in points.
    height_pt: f64,
}

impl ImageWatermark {
    /// Create an image watermark from image bytes at natural size (96 DPI).
    ///
    /// The format (PNG, JPEG, and the other formats supported by
    /// [`crate::format::ImageFormat`]) and pixel dimensions are sniffed
    /// from the bytes; when dimensions cannot be determined the shape
    /// defaults to one square inch before scaling.
    pub fn new(data: Vec<u8>) -> Result<Self> {
        if data.len() > MAX_WATERMARK_IMAGE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "watermark image exceeds {MAX_WATERMARK_IMAGE_BYTES} bytes"
            )));
        }
        let format = crate::format::ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| Error::InvalidFormat("Unknown watermark image format".to_string()))?;
        let (width_pt, height_pt) =
            sniff_dimensions_pt(&data, format).unwrap_or((POINTS_PER_INCH, POINTS_PER_INCH));
        Ok(Self {
            data,
            format,
            width_pt,
            height_pt,
        })
    }

    /// Scale the watermark shape relative to its natural size.
    pub fn set_scale(&mut self, scale: f64) -> Result<&mut Self> {
        if !scale.is_finite() || scale <= 0.0 || scale > MAX_WATERMARK_SCALE {
            return Err(Error::InvalidFormat(format!(
                "watermark scale must be finite and within (0, {MAX_WATERMARK_SCALE}]"
            )));
        }
        self.width_pt *= scale;
        self.height_pt *= scale;
        Ok(self)
    }

    /// Get the image bytes.
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get the sniffed image format.
    pub fn format(&self) -> crate::format::ImageFormat {
        self.format
    }

    /// Get the watermark shape dimensions in points.
    pub fn dimensions_pt(&self) -> (f64, f64) {
        (self.width_pt, self.height_pt)
    }

    /// Generate the image watermark XML for a header paragraph.
    ///
    /// * `idx` - Index for unique shape ID generation (1, 2, or 3)
    /// * `rel_id` - Relationship ID of the media part on the header part
    pub(crate) fn to_header_xml(&self, idx: u32, rel_id: &str) -> Result<String> {
        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>"#);
        xml.push_str(r#"<w:r><w:rPr><w:noProof/></w:rPr><w:pict>"#);
        xml.push_str(r#"<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">"#);
        xml.push_str(r#"<v:stroke joinstyle="miter"/>"#);
        xml.push_str(r#"<v:formulas>"#);
        xml.push_str(r#"<v:f eqn="if lineDrawn pixelLineWidth 0"/>"#);
        xml.push_str(r#"<v:f eqn="sum @0 1 0"/>"#);
        xml.push_str(r#"<v:f eqn="sum 0 0 @1"/>"#);
        xml.push_str(r#"<v:f eqn="prod @2 1 2"/>"#);
        xml.push_str(r#"<v:f eqn="prod @3 21600 pixelWidth"/>"#);
        xml.push_str(r#"<v:f eqn="prod @3 21600 pixelHeight"/>"#);
        xml.push_str(r#"<v:f eqn="sum @0 0 1"/>"#);
        xml.push_str(r#"<v:f eqn="prod @6 1 2"/>"#);
        xml.push_str(r#"<v:f eqn="prod @7 21600 pixelWidth"/>"#);
        xml.push_str(r#"<v:f eqn="sum @8 21600 0"/>"#);
        xml.push_str(r#"<v:f eqn="prod @7 21600 pixelHeight"/>"#);
        xml.push_str(r#"<v:f eqn="sum @10 21600 0"/>"#);
        xml.push_str(r#"</v:formulas>"#);
        xml.push_str(r#"<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>"#);
        xml.push_str(r#"<o:lock v:ext="edit" aspectratio="t"/>"#);
        xml.push_str(r#"</v:shapetype>"#);
        xml.push_str(&format!(
            r##"<v:shape id="PowerPlusWaterMarkObject{idx}" o:spid="_x0000_s102{}" type="#_x0000_t75" "##,
            4 + idx
        ));
        xml.push_str(&format!(
            r#"style="position:absolute;margin-left:0;margin-top:0;width:{}pt;height:{}pt;z-index:{IMAGE_WATERMARK_Z_INDEX};mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" "#,
            self.width_pt, self.height_pt
        ));
        xml.push_str(r#"o:allowincell="f" stroked="f">"#);
        xml.push_str(&format!(r#"<v:imagedata r:id="{}" o:title=""/>"#, rel_id));
        xml.push_str(r#"</v:shape>"#);
        xml.push_str(r#"</w:pict></w:r></w:p>"#);
        Ok(xml)
    }
}

/// Sniff pixel dimensions and convert them to points at 96 DPI.
fn sniff_dimensions_pt(data: &[u8], format: crate::format::ImageFormat) -> Option<(f64, f64)> {
    let (width_px, height_px) = match format {
        crate::format::ImageFormat::Png => png_dimensions(data)?,
        crate::format::ImageFormat::Jpeg => jpeg_dimensions(data)?,
        _ => return None,
    };
    if width_px == 0 || height_px == 0 {
        return None;
    }
    Some((
        f64::from(width_px) * POINTS_PER_INCH / PIXELS_PER_INCH,
        f64::from(height_px) * POINTS_PER_INCH / PIXELS_PER_INCH,
    ))
}

/// Read the width and height from a PNG IHDR chunk.
fn png_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    // 8-byte signature, 4-byte length, 4-byte "IHDR", then width/height.
    let ihdr = data.get(12..16)?;
    if ihdr != b"IHDR" {
        return None;
    }
    let width = u32::from_be_bytes(data.get(16..20)?.try_into().ok()?);
    let height = u32::from_be_bytes(data.get(20..24)?.try_into().ok()?);
    Some((width, height))
}

/// Scan JPEG segments for a start-of-frame marker carrying the dimensions.
fn jpeg_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    if data.get(0..2)? != [0xFF, 0xD8] {
        return None;
    }
    let mut offset = 2usize;
    while let Some(marker) = data.get(offset..offset + 2) {
        if marker[0] != 0xFF {
            return None;
        }
        let code = marker[1];
        // Standalone markers without a length field.
        if code == 0xD8 || code == 0x01 || (0xD0..=0xD7).contains(&code) {
            offset += 2;
            continue;
        }
        let length =
            u16::from_be_bytes(data.get(offset + 2..offset + 4)?.try_into().ok()?) as usize;
        if length < 2 {
            return None;
        }
        // Start-of-frame markers (excluding DHT/DAC/RST/TEM) carry dimensions.
        if matches!(
            code,
            0xC0 | 0xC1
                | 0xC2
                | 0xC3
                | 0xC5
                | 0xC6
                | 0xC7
                | 0xC9
                | 0xCA
                | 0xCB
                | 0xCD
                | 0xCE
                | 0xCF
        ) {
            let height = u16::from_be_bytes(data.get(offset + 5..offset + 7)?.try_into().ok()?);
            let width = u16::from_be_bytes(data.get(offset + 7..offset + 9)?.try_into().ok()?);
            return Some((u32::from(width), u32::from(height)));
        }
        offset = offset.checked_add(2 + length)?;
    }
    None
}

/// A picture watermark anchor discovered in a Word header part.
///
/// The anchor carries the relationship ID referencing the media part from
/// the header part, plus the shape geometry in points. Resolving the media
/// part bytes requires the header part's relationships; see
/// [`crate::Document::image_watermarks`].
#[derive(Debug, Clone, PartialEq)]
pub struct ImageWatermarkAnchor {
    /// Relationship ID of the `v:imagedata` reference.
    relationship_id: String,
    /// Shape width in points.
    width_pt: f64,
    /// Shape height in points.
    height_pt: f64,
}

impl ImageWatermarkAnchor {
    /// Get the relationship ID of the image reference.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Get the shape dimensions in points.
    pub fn dimensions_pt(&self) -> (f64, f64) {
        (self.width_pt, self.height_pt)
    }

    /// Discover picture watermark anchors in a Word header part.
    pub(crate) fn from_header_xml(xml: &[u8]) -> Result<Vec<Self>> {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut shapes: Vec<(usize, Option<ImageAnchorCandidate>)> = Vec::new();
        let mut anchors = Vec::new();

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word header XML nesting is too deep".into())
                    })?;
                    if depth > 256 {
                        return Err(Error::InvalidFormat(
                            "Word header XML nesting exceeds 256 elements".into(),
                        ));
                    }
                    if is_vml(&namespace) && element.local_name().as_ref() == b"shape" {
                        shapes.push((
                            depth,
                            ImageAnchorCandidate::from_shape(&element, decoder, &resolver)?,
                        ));
                    } else if is_vml(&namespace) && element.local_name().as_ref() == b"imagedata" {
                        apply_image_data(&mut shapes, &element, &resolver)?;
                    }
                },
                Event::Empty(element)
                    if is_vml(&namespace) && element.local_name().as_ref() == b"imagedata" =>
                {
                    apply_image_data(&mut shapes, &element, &resolver)?;
                },
                Event::End(element)
                    if is_vml(&namespace) && element.local_name().as_ref() == b"shape" =>
                {
                    let Some((shape_depth, candidate)) = shapes.pop() else {
                        return Err(Error::InvalidFormat(
                            "VML shape end has no matching start".into(),
                        ));
                    };
                    if shape_depth != depth {
                        return Err(Error::InvalidFormat(
                            "invalid nested VML shape structure".into(),
                        ));
                    }
                    if let Some(anchor) = candidate.and_then(ImageAnchorCandidate::finish) {
                        anchors.push(anchor);
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word header XML nesting".into())
                    })?;
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word header XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 || !shapes.is_empty() => {
                    return Err(Error::InvalidFormat("unterminated Word header XML".into()));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(anchors)
    }
}

struct ImageAnchorCandidate {
    width_pt: f64,
    height_pt: f64,
    relationship_id: Option<String>,
}

impl ImageAnchorCandidate {
    fn from_shape(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Option<Self>> {
        let id = unqualified_attribute(element, b"id", decoder, resolver)?;
        let style = unqualified_attribute(element, b"style", decoder, resolver)?;
        let canonical_id = id
            .as_deref()
            .is_some_and(|value| value.starts_with("PowerPlusWaterMarkObject"));
        let background = style.as_deref().is_some_and(|value| {
            value.contains("z-index:-")
                && value.contains("mso-position-horizontal:center")
                && value.contains("mso-position-vertical:center")
        });
        if !canonical_id && !background {
            return Ok(None);
        }
        let mut width_pt = 0.0;
        let mut height_pt = 0.0;
        if let Some(style) = style.as_deref() {
            for declaration in style.split(';') {
                let Some((name, value)) = declaration.split_once(':') else {
                    continue;
                };
                let value = value.trim().strip_suffix("pt").unwrap_or(value.trim());
                let Ok(points) = value.parse::<f64>() else {
                    continue;
                };
                if !points.is_finite() {
                    continue;
                }
                match name.trim().to_ascii_lowercase().as_str() {
                    "width" => width_pt = points,
                    "height" => height_pt = points,
                    _ => {},
                }
            }
        }
        Ok(Some(Self {
            width_pt,
            height_pt,
            relationship_id: None,
        }))
    }

    fn finish(self) -> Option<ImageWatermarkAnchor> {
        Some(ImageWatermarkAnchor {
            relationship_id: self.relationship_id?,
            width_pt: self.width_pt,
            height_pt: self.height_pt,
        })
    }
}

fn apply_image_data(
    shapes: &mut [(usize, Option<ImageAnchorCandidate>)],
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let Some((_, Some(candidate))) = shapes.last_mut() else {
        return Ok(());
    };
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == b"http://schemas.openxmlformats.org/officeDocument/2006/relationships".as_slice()
                    || value == b"http://purl.oclc.org/ooxml/officeDocument/relationships".as_slice()
        );
        if is_relationship {
            candidate.relationship_id =
                Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned());
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_watermark() {
        let wm = Watermark::text("CONFIDENTIAL");
        assert_eq!(wm.get_text(), "CONFIDENTIAL");
    }

    #[test]
    fn test_watermark_customization() {
        let mut wm = Watermark::text("DRAFT");
        wm.set_font("Arial");
        wm.set_color("red");
        wm.set_font_size(2);

        let xml = wm.to_header_xml(1).unwrap();
        assert!(xml.contains("DRAFT"));
        assert!(xml.contains("Arial"));
    }

    #[test]
    fn test_watermark_xml() {
        let wm = Watermark::text("TEST");
        let xml = wm.to_header_xml(1).unwrap();

        assert!(xml.contains("<w:p>"));
        assert!(xml.contains("<w:pict>"));
        assert!(xml.contains("PowerPlusWaterMarkObject"));
        assert!(xml.contains("TEST"));
    }

    #[test]
    fn parses_generated_watermark_from_header_xml() {
        let mut expected = Watermark::text("DRAFT & REVIEW");
        expected.set_font("Arial");
        expected.set_font_size_points(2.5).unwrap();
        expected.set_color("C0FFEE");
        let body = expected.to_header_xml(1).unwrap();
        let xml = format!(
            r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{body}</w:hdr>"#
        );

        let parsed = Watermark::from_header_xml(xml.as_bytes()).unwrap();

        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].get_text(), "DRAFT & REVIEW");
        assert_eq!(parsed[0].font(), "Arial");
        assert_eq!(parsed[0].font_size(), 2.5);
        assert_eq!(parsed[0].color(), "#C0FFEE");
    }

    #[test]
    fn recognizes_aliased_vml_and_ignores_ordinary_or_spoofed_shapes() {
        let xml = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:q="urn:schemas-microsoft-com:vml" xmlns:false="urn:not-vml">
            <q:shape id="ordinary"><q:textpath string="ordinary"/></q:shape>
            <false:shape id="PowerPlusWaterMarkObject0"><false:textpath string="spoofed"/></false:shape>
            <q:shape id="PowerPlusWaterMarkObject1" fillcolor="navy">
                <q:textpath style="font-family:'Aptos';font-size:3pt" string="REAL"/>
            </q:shape>
        </w:hdr>"#;

        let parsed = Watermark::from_header_xml(xml).unwrap();

        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].get_text(), "REAL");
        assert_eq!(parsed[0].font(), "Aptos");
        assert_eq!(parsed[0].font_size(), 3.0);
        assert_eq!(parsed[0].color(), "navy");
    }

    #[test]
    fn rejects_unterminated_watermark_header_xml() {
        let xml = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="PowerPlusWaterMarkObject1">"#;
        assert!(Watermark::from_header_xml(xml).is_err());
    }

    #[test]
    fn rejects_invalid_fractional_font_sizes() {
        let mut watermark = Watermark::text("DRAFT");
        assert!(watermark.set_font_size_points(f64::NAN).is_err());
        assert!(watermark.set_font_size_points(-1.0).is_err());
        assert_eq!(watermark.font_size(), 1.0);
    }

    /// Minimal PNG with a real IHDR chunk carrying the given dimensions.
    fn minimal_png(width: u32, height: u32) -> Vec<u8> {
        let mut data = b"\x89PNG\r\n\x1a\n".to_vec();
        data.extend_from_slice(&13u32.to_be_bytes());
        data.extend_from_slice(b"IHDR");
        data.extend_from_slice(&width.to_be_bytes());
        data.extend_from_slice(&height.to_be_bytes());
        data.extend_from_slice(&[8, 6, 0, 0, 0]);
        data.extend_from_slice(&0u32.to_be_bytes());
        data
    }

    /// Minimal JPEG with a start-of-frame segment carrying the dimensions.
    fn minimal_jpeg(width: u16, height: u16) -> Vec<u8> {
        let mut data = vec![0xFF, 0xD8, 0xFF, 0xC0];
        data.extend_from_slice(&17u16.to_be_bytes());
        data.push(8);
        data.extend_from_slice(&height.to_be_bytes());
        data.extend_from_slice(&width.to_be_bytes());
        data.extend_from_slice(&[3; 10]);
        data
    }

    #[test]
    fn serializes_arbitrary_text_watermark_variants() {
        // Defaults keep the generated diagonal, semitransparent contract.
        let diagonal = Watermark::text("DIAGONAL");
        let xml = diagonal.to_header_xml(1).unwrap();
        assert!(xml.contains("rotation:315;"));
        assert!(xml.contains(r#"<v:fill opacity=".5"/>"#));
        let parsed = Watermark::from_header_xml(
            format!(r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{xml}</w:hdr>"#).as_bytes(),
        )
        .unwrap();
        assert_eq!(parsed[0].layout(), WatermarkLayout::Diagonal);
        assert!(parsed[0].semitransparent());

        // Horizontal, opaque, caller-styled variant.
        let mut horizontal = Watermark::text("TOP SECRET");
        horizontal.set_layout(WatermarkLayout::Horizontal);
        horizontal.set_semitransparent(false);
        horizontal.set_font("Consolas");
        horizontal.set_font_size_points(4.5).unwrap();
        horizontal.set_color("#112233");
        let xml = horizontal.to_header_xml(2).unwrap();
        assert!(!xml.contains("rotation"));
        assert!(!xml.contains("<v:fill"));
        let parsed = Watermark::from_header_xml(
            format!(r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{xml}</w:hdr>"#).as_bytes(),
        )
        .unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0], horizontal);
    }

    #[test]
    fn sniffs_png_and_jpeg_dimensions() {
        let watermark = ImageWatermark::new(minimal_png(200, 100)).unwrap();
        assert_eq!(watermark.dimensions_pt(), (150.0, 75.0));

        let watermark = ImageWatermark::new(minimal_jpeg(96, 48)).unwrap();
        assert_eq!(watermark.dimensions_pt(), (72.0, 36.0));

        let mut scaled = ImageWatermark::new(minimal_png(96, 96)).unwrap();
        scaled.set_scale(2.0).unwrap();
        assert_eq!(scaled.dimensions_pt(), (144.0, 144.0));
    }

    #[test]
    fn validates_image_watermark_inputs() {
        assert!(ImageWatermark::new(vec![0, 1, 2, 3]).is_err());
        assert!(ImageWatermark::new(vec![0x89; MAX_WATERMARK_IMAGE_BYTES + 1]).is_err());
        let mut watermark = ImageWatermark::new(minimal_png(10, 10)).unwrap();
        assert!(watermark.set_scale(0.0).is_err());
        assert!(watermark.set_scale(f64::NAN).is_err());
        assert!(watermark.set_scale(101.0).is_err());
        assert!(watermark.set_scale(0.5).is_ok());
    }

    #[test]
    fn serializes_image_watermark_and_discovers_anchor() {
        let mut watermark = ImageWatermark::new(minimal_png(96, 48)).unwrap();
        watermark.set_scale(0.5).unwrap();
        let body = watermark.to_header_xml(1, "rId7").unwrap();
        assert!(body.contains("PowerPlusWaterMarkObject1"));
        assert!(body.contains(r#"<v:imagedata r:id="rId7" o:title=""/>"#));
        assert!(body.contains("width:36pt;height:18pt"));
        assert!(body.contains("z-index:-251654144"));

        let xml = format!(
            r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{body}</w:hdr>"#
        );
        let anchors = ImageWatermarkAnchor::from_header_xml(xml.as_bytes()).unwrap();
        assert_eq!(anchors.len(), 1);
        assert_eq!(anchors[0].relationship_id(), "rId7");
        assert_eq!(anchors[0].dimensions_pt(), (36.0, 18.0));

        // Text watermarks are not image anchors and vice versa.
        assert!(
            Watermark::from_header_xml(xml.as_bytes())
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn round_trips_arbitrary_text_watermark_through_saved_package() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        let mut watermark = Watermark::text("NEED TO KNOW");
        watermark.set_font("Consolas");
        watermark.set_font_size_points(3.5).unwrap();
        watermark.set_color("#445566");
        watermark.set_layout(WatermarkLayout::Horizontal);
        watermark.set_semitransparent(false);
        package
            .document_mut()
            .unwrap()
            .set_watermark(watermark.clone());
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let watermarks = reopened.document().unwrap().watermarks().unwrap();
        assert_eq!(watermarks, vec![watermark]);
    }

    #[test]
    fn round_trips_image_watermark_through_saved_package() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let payload = minimal_png(120, 60);
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        let mut watermark = ImageWatermark::new(payload.clone()).unwrap();
        watermark.set_scale(2.0).unwrap();
        package
            .document_mut()
            .unwrap()
            .set_image_watermark(watermark);
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        // Text discovery is unaffected by the image watermark.
        assert!(document.watermarks().unwrap().is_empty());

        let parts = document.image_watermarks().unwrap();
        assert_eq!(parts.len(), 3, "one anchor per generated header");
        for part in &parts {
            assert_eq!(part.part_name, "/word/media/watermarkImage1.png");
            assert_eq!(part.content_type, "image/png");
            assert_eq!(part.bytes, payload.as_slice());
        }
        let anchors = document
            .headers()
            .unwrap()
            .into_iter()
            .map(|header| header.image_watermarks().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(anchors.len(), 3);
        assert_eq!(anchors[0][0].dimensions_pt(), (180.0, 90.0));
    }

    #[test]
    fn removal_still_works_for_both_variants() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.set_watermark(Watermark::text("DRAFT"));
            document.set_image_watermark(ImageWatermark::new(minimal_png(10, 10)).unwrap());
            document.remove_watermark();
            document.remove_image_watermark();
            assert!(!document.has_watermark());
            assert!(!document.has_image_watermark());
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        assert!(document.watermarks().unwrap().is_empty());
        assert!(document.image_watermarks().unwrap().is_empty());
    }

    #[test]
    fn text_and_image_watermarks_coexist() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let payload = minimal_png(48, 48);
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.add_paragraph_with_text("content page");
            document.set_watermark(Watermark::text("CONFIDENTIAL"));
            document.set_image_watermark(ImageWatermark::new(payload.clone()).unwrap());
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        assert!(document.text().unwrap().contains("content page"));

        let watermarks = document.watermarks().unwrap();
        assert_eq!(watermarks.len(), 1);
        assert_eq!(watermarks[0].get_text(), "CONFIDENTIAL");

        let parts = document.image_watermarks().unwrap();
        assert_eq!(parts.len(), 3);
        assert!(parts.iter().all(|part| part.bytes == payload.as_slice()));
    }
}
