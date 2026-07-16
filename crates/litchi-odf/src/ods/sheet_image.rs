//! Inert sheet-level image frames from `table:shapes`.

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result, xml::escape_xml};

pub(crate) const MAX_IMAGES_PER_SHEET: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_INLINE_BYTES: usize = 16 * 1_048_576;

pub(crate) fn normalize_sheet_image(
    mut image: crate::OdfImage,
    sheet_name: &str,
) -> Result<crate::OdfImage> {
    image.part = crate::OdfImagePart::Content;
    image.alternative_index = 0;
    let frame = image.frame.as_mut().ok_or_else(|| {
        Error::InvalidFormat("sheet image requires a draw:frame".to_string())
    })?;
    frame.sheet_shape = true;
    frame.sheet_name = Some(sheet_name.to_string());
    frame.page_name = None;
    validate_sheet_image(&image)?;
    Ok(image)
}

pub(crate) fn validate_sheet_image(image: &crate::OdfImage) -> Result<()> {
    if image.part != crate::OdfImagePart::Content {
        return invalid("sheet image must belong to content.xml");
    }
    if image.alternative_index != 0 {
        return invalid("sheet image alternatives are unsupported");
    }
    let frame = image.frame.as_ref().ok_or_else(|| {
        Error::InvalidFormat("sheet image requires a draw:frame".to_string())
    })?;
    if !frame.sheet_shape {
        return invalid("sheet image frame is not anchored in table:shapes");
    }
    for (name, value) in [
        ("draw:name", frame.name.as_deref()),
        ("xml:id", frame.xml_id.as_deref()),
        ("svg:title", frame.title.as_deref()),
        ("svg:desc", frame.description.as_deref()),
        ("text:anchor-type", frame.anchor_type.as_deref()),
        ("table:name", frame.sheet_name.as_deref()),
        ("draw:image xml:id", image.xml_id.as_deref()),
        ("draw:filter-name", image.filter_name.as_deref()),
        ("draw:mime-type", image.declared_media_type.as_deref()),
    ] {
        if let Some(value) = value {
            validate_text(value, name, true)?;
        }
    }
    for (name, value, nonnegative) in [
        ("svg:x", frame.x.as_deref(), false),
        ("svg:y", frame.y.as_deref(), false),
        ("svg:width", frame.width.as_deref(), true),
        ("svg:height", frame.height.as_deref(), true),
    ] {
        if let Some(value) = value {
            validate_length(value, name, nonnegative)?;
        }
    }
    if !matches!(image.link_type.as_deref(), None | Some("simple")) {
        return invalid("sheet draw:image xlink:type must be simple");
    }
    if !matches!(image.show.as_deref(), None | Some("embed")) {
        return invalid("sheet draw:image xlink:show must be embed");
    }
    if !matches!(image.actuate.as_deref(), None | Some("onLoad")) {
        return invalid("sheet draw:image xlink:actuate must be onLoad");
    }
    match &image.source {
        crate::OdfImageSource::Inline { bytes, ignored_href } => {
            if bytes.len() > MAX_INLINE_BYTES {
                return invalid("inline sheet image exceeds 16 MiB");
            }
            if let Some(href) = ignored_href {
                validate_text(href, "ignored xlink:href", true)?;
            }
        },
        crate::OdfImageSource::PackagePart { href, .. }
        | crate::OdfImageSource::MissingPackagePart { href, .. }
        | crate::OdfImageSource::Linked { href } => {
            validate_text(href, "xlink:href", false)?;
        },
        crate::OdfImageSource::Missing => return invalid("sheet image requires a source"),
    }
    Ok(())
}

pub(crate) fn write_sheet_images(out: &mut String, images: &[crate::OdfImage]) -> Result<()> {
    if images.is_empty() {
        return Ok(());
    }
    if images.len() > MAX_IMAGES_PER_SHEET {
        return invalid(format!("sheet exceeds {MAX_IMAGES_PER_SHEET} images"));
    }
    out.push_str("<table:shapes>");
    for image in images {
        validate_sheet_image(image)?;
        let frame = image.frame.as_ref().expect("validated frame");
        out.push_str("<draw:frame");
        attribute(out, "draw:name", frame.name.as_deref());
        attribute(out, "xml:id", frame.xml_id.as_deref());
        attribute(out, "text:anchor-type", frame.anchor_type.as_deref());
        attribute(out, "svg:x", frame.x.as_deref());
        attribute(out, "svg:y", frame.y.as_deref());
        attribute(out, "svg:width", frame.width.as_deref());
        attribute(out, "svg:height", frame.height.as_deref());
        out.push('>');
        if let Some(title) = &frame.title {
            out.push_str("<svg:title>");
            out.push_str(&escape_xml(title));
            out.push_str("</svg:title>");
        }
        if let Some(description) = &frame.description {
            out.push_str("<svg:desc>");
            out.push_str(&escape_xml(description));
            out.push_str("</svg:desc>");
        }
        out.push_str("<draw:image");
        attribute(out, "xml:id", image.xml_id.as_deref());
        attribute(out, "draw:filter-name", image.filter_name.as_deref());
        attribute(out, "draw:mime-type", image.declared_media_type.as_deref());
        let href = match &image.source {
            crate::OdfImageSource::Inline { ignored_href, .. } => ignored_href.as_deref(),
            crate::OdfImageSource::PackagePart { href, .. }
            | crate::OdfImageSource::MissingPackagePart { href, .. }
            | crate::OdfImageSource::Linked { href } => Some(href.as_str()),
            crate::OdfImageSource::Missing => None,
        };
        attribute(out, "xlink:href", href);
        attribute(out, "xlink:type", image.link_type.as_deref());
        attribute(out, "xlink:show", image.show.as_deref());
        attribute(out, "xlink:actuate", image.actuate.as_deref());
        match &image.source {
            crate::OdfImageSource::Inline { bytes, .. } => {
                out.push_str("><office:binary-data>");
                out.push_str(&BASE64_STANDARD.encode(bytes));
                out.push_str("</office:binary-data></draw:image>");
            },
            _ => out.push_str("/>"),
        }
        out.push_str("</draw:frame>");
    }
    out.push_str("</table:shapes>");
    Ok(())
}

fn attribute(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&escape_xml(value));
        out.push('"');
    }
}

fn validate_length(value: &str, name: &str, nonnegative: bool) -> Result<()> {
    validate_text(value, name, false)?;
    let number = ["cm", "mm", "in", "pt", "pc", "px"]
        .into_iter()
        .find_map(|suffix| value.strip_suffix(suffix))
        .ok_or_else(|| Error::InvalidFormat(format!("invalid {name} length '{value}'")))?;
    let number = number
        .parse::<f64>()
        .map_err(|_| Error::InvalidFormat(format!("invalid {name} length '{value}'")))?;
    if !number.is_finite() || (nonnegative && number < 0.0) {
        return invalid(format!("invalid {name} length '{value}'"));
    }
    Ok(())
}

fn validate_text(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds {MAX_VALUE_BYTES} bytes"));
    }
    if value
        .chars()
        .any(|character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'))
    {
        return invalid(format!("{name} contains invalid XML characters"));
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
