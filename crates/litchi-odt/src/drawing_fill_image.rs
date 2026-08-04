//! Typed, inert ODF drawing fill-image resources.

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK_NS: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_IMAGES: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_ENCODED_BYTES: usize = 24 * 1_048_576;
const MAX_INLINE_BYTES: usize = 16 * 1_048_576;
const MAX_TOTAL_INLINE_BYTES: usize = 64 * 1_048_576;
const MAX_AGGREGATE_BYTES: usize = 96 * 1_048_576;

/// Unit for an optional fill-image intrinsic size.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FillImageLengthUnit {
    Centimeter,
    Millimeter,
    Inch,
    Point,
    Pica,
    Pixel,
}

impl FillImageLengthUnit {
    const fn suffix(self) -> &'static str {
        match self {
            Self::Centimeter => "cm",
            Self::Millimeter => "mm",
            Self::Inch => "in",
            Self::Point => "pt",
            Self::Pica => "pc",
            Self::Pixel => "px",
        }
    }
}

/// A finite, nonnegative ODF length.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct FillImageLength {
    value: f64,
    unit: FillImageLengthUnit,
}

impl FillImageLength {
    pub fn new(value: f64, unit: FillImageLengthUnit) -> Result<Self> {
        if !value.is_finite() || value < 0.0 {
            return invalid("fill-image length must be finite and nonnegative");
        }
        Ok(Self { value, unit })
    }

    pub const fn value(self) -> f64 {
        self.value
    }

    pub const fn unit(self) -> FillImageLengthUnit {
        self.unit
    }
}

impl FromStr for FillImageLength {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let (number, unit) = split_length(value)?;
        validate_decimal(number, value)?;
        let number = number
            .parse::<f64>()
            .map_err(|_| make_error(format!("invalid fill-image length '{value}'")))?;
        Self::new(number, unit)
    }
}

impl fmt::Display for FillImageLength {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{}{}",
            canonical_number(self.value),
            self.unit.suffix()
        )
    }
}

/// Whether an inert href is a safe relative package path.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FillImageLinkKind {
    PackagePart,
    InertExternal,
}

/// A retained link which is never automatically dereferenced.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct FillImageLink {
    href: String,
    kind: FillImageLinkKind,
}

impl FillImageLink {
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let href = href.into();
        validate_text(&href, "xlink:href", true, MAX_VALUE_BYTES)?;
        let kind = if safe_package_path(&href) {
            FillImageLinkKind::PackagePart
        } else {
            FillImageLinkKind::InertExternal
        };
        Ok(Self { href, kind })
    }

    pub fn href(&self) -> &str {
        &self.href
    }

    pub const fn kind(&self) -> FillImageLinkKind {
        self.kind
    }

    pub fn package_path(&self) -> Option<&str> {
        (self.kind == FillImageLinkKind::PackagePart).then_some(&self.href)
    }
}

/// The source of a fill-image resource.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum FillImageSource {
    Linked(FillImageLink),
    Inline {
        bytes: Vec<u8>,
        /// ODF consumers ignore this link when inline data is present.
        ignored_link: Option<FillImageLink>,
    },
}

impl FillImageSource {
    pub fn inline_bytes(&self) -> Option<&[u8]> {
        match self {
            Self::Inline { bytes, .. } => Some(bytes),
            Self::Linked(_) => None,
        }
    }

    pub fn link(&self) -> Option<&FillImageLink> {
        match self {
            Self::Linked(link) => Some(link),
            Self::Inline { ignored_link, .. } => ignored_link.as_ref(),
        }
    }
}

/// The only schema-defined `xlink:show` mode for fill images.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FillImageShow {
    Embed,
}

/// The only schema-defined `xlink:actuate` mode for fill images.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FillImageActuate {
    OnLoad,
}

/// One named `draw:fill-image` resource.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawingFillImage {
    pub name: String,
    pub display_name: Option<String>,
    pub width: Option<FillImageLength>,
    pub height: Option<FillImageLength>,
    pub source: FillImageSource,
    pub show: Option<FillImageShow>,
    pub actuate: Option<FillImageActuate>,
}

impl DrawingFillImage {
    pub fn validate(&self) -> Result<()> {
        validate_text(&self.name, "draw:name", false, MAX_VALUE_BYTES)?;
        if let Some(display_name) = &self.display_name {
            validate_text(display_name, "draw:display-name", true, MAX_VALUE_BYTES)?;
        }
        for length in [self.width, self.height].into_iter().flatten() {
            FillImageLength::new(length.value, length.unit)?;
        }
        match &self.source {
            FillImageSource::Linked(link) => validate_link(link)?,
            FillImageSource::Inline {
                bytes,
                ignored_link,
            } => {
                if bytes.len() > MAX_INLINE_BYTES {
                    return invalid("inline fill image exceeds 16 MiB");
                }
                if let Some(link) = ignored_link {
                    validate_link(link)?;
                } else if self.show.is_some() || self.actuate.is_some() {
                    return invalid("XLink modes require an xlink:href");
                }
            },
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(256 + encoded_size(&self.source));
        write_fill_image(&mut output, self, true);
        Ok(output)
    }
}

/// Ordered fill-image resources from `office:styles`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct DrawingFillImages {
    pub images: Vec<DrawingFillImage>,
}

impl DrawingFillImages {
    pub fn get(&self, name: &str) -> Option<&DrawingFillImage> {
        self.images.iter().find(|image| image.name == name)
    }

    pub fn validate(&self) -> Result<()> {
        if self.images.len() > MAX_IMAGES {
            return invalid(format!("drawing styles exceed {MAX_IMAGES} fill images"));
        }
        let mut names = HashSet::with_capacity(self.images.len());
        let mut aggregate = 0usize;
        let mut inline_total = 0usize;
        for image in &self.images {
            image.validate()?;
            if !names.insert(image.name.as_str()) {
                return invalid(format!(
                    "duplicate drawing fill-image name '{}'",
                    image.name
                ));
            }
            aggregate = aggregate
                .checked_add(image.name.len())
                .and_then(|size| {
                    size.checked_add(image.display_name.as_ref().map_or(0, String::len))
                })
                .and_then(|size| {
                    size.checked_add(image.source.link().map_or(0, |link| link.href().len()))
                })
                .ok_or_else(|| make_error("fill-image size overflow"))?;
            if let Some(bytes) = image.source.inline_bytes() {
                inline_total = inline_total
                    .checked_add(bytes.len())
                    .ok_or_else(|| make_error("inline fill-image size overflow"))?;
            }
            if aggregate > MAX_AGGREGATE_BYTES || inline_total > MAX_TOTAL_INLINE_BYTES {
                return invalid("fill-image resources exceed aggregate limits");
            }
        }
        Ok(())
    }

    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let capacity = self.images.iter().fold(256usize, |size, image| {
            size.saturating_add(256 + encoded_size(&image.source))
        });
        let mut output = String::with_capacity(capacity);
        output.push_str(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">"#,
        );
        for image in &self.images {
            write_fill_image(&mut output, image, false);
        }
        output.push_str("</office:styles>");
        Ok(output)
    }
}

impl crate::OpenDocumentPackage {
    pub fn drawing_fill_images(&self) -> Result<DrawingFillImages> {
        let styles = self.styles_xml()?;
        parse_drawing_fill_images(styles.as_deref().unwrap_or_default())
    }

    /// Load a safe package image or borrow inline bytes without copying.
    pub fn drawing_fill_image_bytes<'a>(
        &self,
        image: &'a DrawingFillImage,
    ) -> Result<Option<Cow<'a, [u8]>>> {
        match &image.source {
            FillImageSource::Inline { bytes, .. } => Ok(Some(Cow::Borrowed(bytes))),
            FillImageSource::Linked(link) => {
                let Some(path) = link.package_path() else {
                    return Ok(None);
                };
                if self.has_file(path)? {
                    self.get_file(path).map(Cow::Owned).map(Some)
                } else {
                    Ok(None)
                }
            },
        }
    }
}

impl crate::FlatOpenDocument {
    pub fn drawing_fill_images(&self) -> Result<DrawingFillImages> {
        parse_drawing_fill_images(self.xml())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum NamespaceKind {
    None,
    Office,
    Draw,
    Svg,
    Xlink,
    Other,
}

#[derive(Clone)]
struct Frame {
    namespace: NamespaceKind,
    local: String,
}

struct FillBuilder {
    parent_depth: usize,
    name: String,
    display_name: Option<String>,
    width: Option<FillImageLength>,
    height: Option<FillImageLength>,
    link: Option<FillImageLink>,
    show: Option<FillImageShow>,
    actuate: Option<FillImageActuate>,
    binary_present: bool,
    binary_parent_depth: Option<usize>,
    encoded: String,
}

type Attributes = HashMap<(NamespaceKind, String), String>;

pub fn parse_drawing_fill_images(xml: &str) -> Result<DrawingFillImages> {
    if !xml.contains("fill-image") {
        return Ok(DrawingFillImages::default());
    }
    if xml.len() > MAX_XML_BYTES {
        return invalid("drawing fill-image XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<FillBuilder> = None;
    let mut result = DrawingFillImages::default();
    let mut aggregate = 0usize;
    let mut inline_total = 0usize;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid drawing fill-image XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if let Some(fill) = active.as_mut() {
                    if fill.binary_parent_depth.is_some()
                        || namespace != NamespaceKind::Office
                        || local != "binary-data"
                        || stack.len() != fill.parent_depth + 1
                    {
                        return invalid("draw:fill-image contains an unsupported child element");
                    }
                    if fill.binary_present || element.attributes().next().is_some() {
                        return invalid("invalid or duplicate office:binary-data element");
                    }
                    fill.binary_present = true;
                    fill.binary_parent_depth = Some(stack.len());
                } else if namespace == NamespaceKind::Draw && local == "fill-image" {
                    ensure_location(&stack)?;
                    ensure_count(result.images.len())?;
                    active = Some(parse_fill_start(
                        &reader,
                        element,
                        stack.len(),
                        &mut aggregate,
                    )?);
                }
                stack.push(Frame { namespace, local });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("fill-image XML exceeds {MAX_DEPTH} levels"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                if let Some(fill) = active.as_mut() {
                    if fill.binary_parent_depth.is_some()
                        || namespace != NamespaceKind::Office
                        || local != "binary-data"
                        || stack.len() != fill.parent_depth + 1
                        || fill.binary_present
                        || element.attributes().next().is_some()
                    {
                        return invalid("draw:fill-image contains an unsupported child element");
                    }
                    fill.binary_present = true;
                } else if namespace == NamespaceKind::Draw && local == "fill-image" {
                    ensure_location(&stack)?;
                    ensure_count(result.images.len())?;
                    let builder = parse_fill_start(&reader, element, stack.len(), &mut aggregate)?;
                    result.images.push(finish_fill(builder, &mut inline_total)?);
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("fill-image XML depth underflow"))?;
                if let Some(fill) = active.as_mut()
                    && fill.binary_parent_depth == Some(stack.len())
                {
                    if frame.namespace != NamespaceKind::Office || frame.local != "binary-data" {
                        return invalid("unexpected office:binary-data end element");
                    }
                    fill.binary_parent_depth = None;
                } else if active
                    .as_ref()
                    .is_some_and(|fill| fill.parent_depth == stack.len())
                {
                    if frame.namespace != NamespaceKind::Draw || frame.local != "fill-image" {
                        return invalid("unexpected draw:fill-image end element");
                    }
                    result.images.push(finish_fill(
                        active.take().expect("active fill image checked"),
                        &mut inline_total,
                    )?);
                }
            },
            Event::Text(ref text) if active.is_some() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| make_error(format!("invalid fill-image text: {error}")))?;
                if active
                    .as_ref()
                    .and_then(|fill| fill.binary_parent_depth)
                    .is_some()
                {
                    append_base64(
                        active.as_mut().expect("active binary fill image"),
                        &value,
                        &mut aggregate,
                    )?;
                } else if !value.chars().all(char::is_whitespace) {
                    return invalid("draw:fill-image may contain only office:binary-data");
                }
            },
            Event::CData(ref value)
                if active
                    .as_ref()
                    .and_then(|fill| fill.binary_parent_depth)
                    .is_some() =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| make_error(format!("invalid fill-image CDATA: {error}")))?;
                append_base64(
                    active.as_mut().expect("active binary fill image"),
                    &value,
                    &mut aggregate,
                )?;
            },
            Event::CData(_) | Event::GeneralRef(_) if active.is_some() => {
                return invalid("draw:fill-image contains unsupported character data");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are prohibited in fill images");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() {
        return invalid("unterminated drawing fill-image XML");
    }
    result.validate()?;
    Ok(result)
}

fn parse_fill_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    parent_depth: usize,
    aggregate: &mut usize,
) -> Result<FillBuilder> {
    let mut values = attributes(reader, element, aggregate)?;
    let name = required(&mut values, NamespaceKind::Draw, "name", "draw:name")?;
    let display_name = take(&mut values, NamespaceKind::Draw, "display-name");
    let width = take(&mut values, NamespaceKind::Svg, "width")
        .map(|value| value.parse())
        .transpose()?;
    let height = take(&mut values, NamespaceKind::Svg, "height")
        .map(|value| value.parse())
        .transpose()?;
    let href = take(&mut values, NamespaceKind::Xlink, "href");
    let link_type = take(&mut values, NamespaceKind::Xlink, "type");
    let show = take(&mut values, NamespaceKind::Xlink, "show")
        .map(|value| match value.as_str() {
            "embed" => Ok(FillImageShow::Embed),
            _ => invalid(format!("unsupported xlink:show '{value}'")),
        })
        .transpose()?;
    let actuate = take(&mut values, NamespaceKind::Xlink, "actuate")
        .map(|value| match value.as_str() {
            "onLoad" => Ok(FillImageActuate::OnLoad),
            _ => invalid(format!("unsupported xlink:actuate '{value}'")),
        })
        .transpose()?;
    reject_attributes(&values)?;
    let link = href.map(FillImageLink::new).transpose()?;
    match (&link, link_type.as_deref()) {
        (Some(_), Some("simple")) => {},
        (Some(_), _) => return invalid("linked fill image requires xlink:type='simple'"),
        (None, None) if show.is_none() && actuate.is_none() => {},
        (None, _) => return invalid("XLink attributes require xlink:href"),
    }
    Ok(FillBuilder {
        parent_depth,
        name,
        display_name,
        width,
        height,
        link,
        show,
        actuate,
        binary_present: false,
        binary_parent_depth: None,
        encoded: String::new(),
    })
}

fn finish_fill(builder: FillBuilder, inline_total: &mut usize) -> Result<DrawingFillImage> {
    let source = if builder.binary_present {
        let bytes = BASE64_STANDARD
            .decode(builder.encoded.as_bytes())
            .map_err(|error| make_error(format!("invalid fill-image base64 data: {error}")))?;
        if bytes.len() > MAX_INLINE_BYTES {
            return invalid("inline fill image exceeds 16 MiB");
        }
        *inline_total = inline_total
            .checked_add(bytes.len())
            .ok_or_else(|| make_error("inline fill-image size overflow"))?;
        if *inline_total > MAX_TOTAL_INLINE_BYTES {
            return invalid("inline fill images exceed 64 MiB");
        }
        FillImageSource::Inline {
            bytes,
            ignored_link: builder.link,
        }
    } else {
        FillImageSource::Linked(
            builder
                .link
                .ok_or_else(|| make_error("fill image requires xlink:href or binary data"))?,
        )
    };
    let image = DrawingFillImage {
        name: builder.name,
        display_name: builder.display_name,
        width: builder.width,
        height: builder.height,
        source,
        show: builder.show,
        actuate: builder.actuate,
    };
    image.validate()?;
    Ok(image)
}

fn append_base64(fill: &mut FillBuilder, value: &str, aggregate: &mut usize) -> Result<()> {
    for byte in value.bytes() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        if !(byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/' | b'=')) {
            return invalid("inline fill image contains a non-base64 character");
        }
        fill.encoded.push(char::from(byte));
    }
    if fill.encoded.len() > MAX_ENCODED_BYTES {
        return invalid("encoded inline fill image exceeds 24 MiB");
    }
    *aggregate = aggregate
        .checked_add(value.len())
        .ok_or_else(|| make_error("fill-image aggregate size overflow"))?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return invalid("fill-image values exceed aggregate limit");
    }
    Ok(())
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut values = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| make_error(format!("invalid fill-image attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid fill-image attribute: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("fill-image attribute exceeds 64 KiB");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("fill-image aggregate size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("fill-image values exceed aggregate limit");
        }
        if values.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded fill-image attribute");
        }
    }
    Ok(values)
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    match resolved {
        ResolveResult::Unbound => Ok(NamespaceKind::None),
        ResolveResult::Bound(namespace) => {
            let bytes: &[u8] = namespace.as_ref();
            Ok(if bytes == OFFICE_NS {
                NamespaceKind::Office
            } else if bytes == DRAW_NS {
                NamespaceKind::Draw
            } else if bytes == SVG_NS {
                NamespaceKind::Svg
            } else if bytes == XLINK_NS {
                NamespaceKind::Xlink
            } else {
                NamespaceKind::Other
            })
        },
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn ensure_location(stack: &[Frame]) -> Result<()> {
    if stack
        .last()
        .is_some_and(|frame| frame.namespace == NamespaceKind::Office && frame.local == "styles")
    {
        Ok(())
    } else {
        invalid("draw:fill-image must be a direct child of office:styles")
    }
}

fn ensure_count(count: usize) -> Result<()> {
    if count >= MAX_IMAGES {
        invalid(format!("drawing styles exceed {MAX_IMAGES} fill images"))
    } else {
        Ok(())
    }
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if local == "fill-image" && namespace != NamespaceKind::Draw {
        return invalid("fill-image element uses the wrong namespace");
    }
    if local == "binary-data" && namespace != NamespaceKind::Office {
        return invalid("binary-data element uses the wrong namespace");
    }
    Ok(())
}

fn take(values: &mut Attributes, namespace: NamespaceKind, local: &str) -> Option<String> {
    values.remove(&(namespace, local.to_owned()))
}

fn required(
    values: &mut Attributes,
    namespace: NamespaceKind,
    local: &str,
    qualified: &str,
) -> Result<String> {
    take(values, namespace, local)
        .ok_or_else(|| make_error(format!("missing required {qualified} attribute")))
}

fn reject_attributes(values: &Attributes) -> Result<()> {
    if let Some(((namespace, local), _)) = values.iter().next() {
        return invalid(format!(
            "unsupported fill-image attribute {namespace:?}:{local}"
        ));
    }
    Ok(())
}

fn validate_link(link: &FillImageLink) -> Result<()> {
    validate_text(link.href(), "xlink:href", true, MAX_VALUE_BYTES)?;
    if link.kind
        != if safe_package_path(link.href()) {
            FillImageLinkKind::PackagePart
        } else {
            FillImageLinkKind::InertExternal
        }
    {
        return invalid("fill-image link classification is inconsistent");
    }
    Ok(())
}

fn safe_package_path(value: &str) -> bool {
    !value.is_empty()
        && !value.starts_with('/')
        && !value.starts_with('\\')
        && !value.contains(['\\', '?', '#', '%', ':'])
        && !value.chars().any(char::is_control)
        && value
            .split('/')
            .all(|segment| !segment.is_empty() && segment != "." && segment != "..")
}

fn split_length(value: &str) -> Result<(&str, FillImageLengthUnit)> {
    for (suffix, unit) in [
        ("cm", FillImageLengthUnit::Centimeter),
        ("mm", FillImageLengthUnit::Millimeter),
        ("in", FillImageLengthUnit::Inch),
        ("pt", FillImageLengthUnit::Point),
        ("pc", FillImageLengthUnit::Pica),
        ("px", FillImageLengthUnit::Pixel),
    ] {
        if let Some(number) = value.strip_suffix(suffix) {
            return Ok((number, unit));
        }
    }
    invalid(format!(
        "fill-image length lacks a supported unit: '{value}'"
    ))
}

fn validate_decimal(number: &str, original: &str) -> Result<()> {
    if number.is_empty()
        || number.starts_with('+')
        || number.starts_with('-')
        || number.chars().any(char::is_whitespace)
        || number.contains(['e', 'E'])
    {
        return invalid(format!("invalid fill-image length '{original}'"));
    }
    let mut parts = number.split('.');
    let integer = parts.next().unwrap_or_default();
    let fraction = parts.next();
    if parts.next().is_some()
        || integer.is_empty()
        || !integer.bytes().all(|byte| byte.is_ascii_digit())
        || fraction
            .is_some_and(|part| part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit()))
    {
        return invalid(format!("invalid fill-image length '{original}'"));
    }
    Ok(())
}

fn validate_text(value: &str, name: &str, allow_empty: bool, limit: usize) -> Result<()> {
    if (!allow_empty && value.is_empty()) || value.len() > limit {
        return invalid(format!("invalid {name} length"));
    }
    if value
        .chars()
        .any(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    {
        return invalid(format!("{name} contains prohibited control characters"));
    }
    Ok(())
}

fn write_fill_image(output: &mut String, image: &DrawingFillImage, standalone: bool) {
    output.push_str("<draw:fill-image");
    if standalone {
        output.push_str(
            r#" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink""#,
        );
    }
    write_attribute(output, "draw:name", &image.name);
    if let Some(display_name) = &image.display_name {
        write_attribute(output, "draw:display-name", display_name);
    }
    if let Some(width) = image.width {
        write_attribute(output, "svg:width", &width.to_string());
    }
    if let Some(height) = image.height {
        write_attribute(output, "svg:height", &height.to_string());
    }
    if let Some(link) = image.source.link() {
        write_attribute(output, "xlink:type", "simple");
        write_attribute(output, "xlink:href", link.href());
        if image.show.is_some() {
            write_attribute(output, "xlink:show", "embed");
        }
        if image.actuate.is_some() {
            write_attribute(output, "xlink:actuate", "onLoad");
        }
    }
    match &image.source {
        FillImageSource::Linked(_) => output.push_str("/>"),
        FillImageSource::Inline { bytes, .. } => {
            output.push_str("><office:binary-data>");
            BASE64_STANDARD.encode_string(bytes, output);
            output.push_str("</office:binary-data></draw:fill-image>");
        },
    }
}

fn encoded_size(source: &FillImageSource) -> usize {
    source
        .inline_bytes()
        .map_or(0, |bytes| bytes.len().saturating_add(2) / 3 * 4)
}

fn write_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape_xml(output, value);
    output.push('"');
}

fn escape_xml(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn canonical_number(value: f64) -> String {
    if value == 0.0 {
        "0".to_owned()
    } else {
        value.to_string()
    }
}

fn decode(value: &[u8], what: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| make_error(format!("invalid UTF-8 in fill-image {what}")))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    const XLINK: &str = "http://www.w3.org/1999/xlink";

    fn wrap(body: &str) -> String {
        format!(
            r#"<office:styles xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:svg="{SVG}" xmlns:xlink="{XLINK}">{body}</office:styles>"#
        )
    }

    #[test]
    fn parses_and_round_trips_linked_and_inline_images() {
        let xml = wrap(
            r#"<draw:fill-image draw:name="package" draw:display-name="Package" svg:width="2.5cm" svg:height="30px" xlink:type="simple" xlink:href="Pictures/fill.png" xlink:show="embed" xlink:actuate="onLoad"/><draw:fill-image draw:name="inline"><office:binary-data>AAEC/w==</office:binary-data></draw:fill-image><draw:fill-image draw:name="remote" xlink:type="simple" xlink:href="https://example.invalid/image.png"/>"#,
        );
        let parsed = parse_drawing_fill_images(&xml).unwrap();
        assert_eq!(parsed.images.len(), 3);
        assert_eq!(
            parsed.get("inline").unwrap().source.inline_bytes(),
            Some([0, 1, 2, 255].as_slice())
        );
        assert_eq!(
            parsed.get("package").unwrap().source.link().unwrap().kind(),
            FillImageLinkKind::PackagePart
        );
        assert_eq!(
            parsed.get("remote").unwrap().source.link().unwrap().kind(),
            FillImageLinkKind::InertExternal
        );
        let serialized = parsed.to_xml().unwrap();
        assert_eq!(parse_drawing_fill_images(&serialized).unwrap(), parsed);
    }

    #[test]
    fn rejects_malformed_or_unsafe_structure() {
        for xml in [
            wrap(r#"<draw:fill-image draw:name="x"/>"#),
            wrap(r#"<draw:fill-image draw:name="x" xlink:href="Pictures/x.png"/>"#),
            wrap(
                r#"<draw:fill-image draw:name="x" xlink:type="extended" xlink:href="Pictures/x.png"/>"#,
            ),
            wrap(
                r#"<draw:fill-image draw:name="x" xlink:show="embed"><office:binary-data>AA==</office:binary-data></draw:fill-image>"#,
            ),
            wrap(
                r#"<draw:fill-image draw:name="x"><office:binary-data>!!!</office:binary-data></draw:fill-image>"#,
            ),
            wrap(
                r#"<draw:fill-image draw:name="x"><office:binary-data>AA==</office:binary-data><office:binary-data>AA==</office:binary-data></draw:fill-image>"#,
            ),
            wrap(r#"<draw:fill-image draw:name="x"><draw:image/></draw:fill-image>"#),
            wrap(
                r#"<draw:fill-image draw:name="x"><office:binary-data>AA==</office:binary-data></draw:fill-image><draw:fill-image draw:name="x" xlink:type="simple" xlink:href="x"/>"#,
            ),
            format!(
                r#"<office:document xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:xlink="{XLINK}"><draw:fill-image draw:name="x" xlink:type="simple" xlink:href="x"/></office:document>"#
            ),
            format!(
                r#"<!DOCTYPE x><office:styles xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:xlink="{XLINK}"><draw:fill-image draw:name="x" xlink:type="simple" xlink:href="x"/></office:styles>"#
            ),
        ] {
            assert!(parse_drawing_fill_images(&xml).is_err(), "accepted {xml}");
        }
        assert_eq!(
            FillImageLink::new("../Pictures/x.png").unwrap().kind(),
            FillImageLinkKind::InertExternal
        );
    }

    #[test]
    fn parses_local_linked_and_inline_resources() {
        let linked_xml = include_str!("../../../test-data/odf/drawing/fill-image-linked.fodp");
        let linked = crate::FlatOpenDocument::from_bytes(linked_xml.as_bytes().to_vec()).unwrap();
        let images = linked.drawing_fill_images().unwrap();
        assert_eq!(
            images
                .get("remote_bg")
                .unwrap()
                .source
                .link()
                .unwrap()
                .kind(),
            FillImageLinkKind::InertExternal
        );

        let inline_xml = include_str!("../../../test-data/odf/drawing/fill-image-inline.fodg");
        let inline = crate::FlatOpenDocument::from_bytes(inline_xml.as_bytes().to_vec()).unwrap();
        let images = inline.drawing_fill_images().unwrap();
        let bytes = images
            .get("libreoffice_5f_0")
            .unwrap()
            .source
            .inline_bytes()
            .unwrap();
        assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n");
    }
}
