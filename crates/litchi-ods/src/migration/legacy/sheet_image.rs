//! Inert sheet-level image frames from `table:shapes`.

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result, xml::escape_xml};

pub(crate) const MAX_IMAGES_PER_SHEET: usize = 65_536;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_INLINE_BYTES: usize = 16 * 1_048_576;
const MAX_TOTAL_INLINE_BYTES: usize = 64 * 1_048_576;

pub(crate) fn normalize_sheet_image(
    mut image: crate::Image,
    sheet_name: &str,
) -> Result<crate::Image> {
    image.part = crate::ImagePart::Content;
    image.alternative_index = 0;
    let frame = image
        .frame
        .as_mut()
        .ok_or_else(|| Error::InvalidFormat("sheet image requires a draw:frame".to_string()))?;
    frame.sheet_shape = true;
    frame.sheet_name = Some(sheet_name.to_string());
    frame.page_name = None;
    validate_sheet_image(&image)?;
    Ok(image)
}

pub(crate) fn validate_sheet_image(image: &crate::Image) -> Result<()> {
    if image.part != crate::ImagePart::Content {
        return invalid("sheet image must belong to content.xml");
    }
    let frame = image
        .frame
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("sheet image requires a draw:frame".to_string()))?;
    if !frame.sheet_shape {
        return invalid("sheet image frame is not anchored in table:shapes");
    }
    for (name, value) in [
        ("draw:name", frame.name.as_deref()),
        ("xml:id", frame.xml_id.as_deref()),
        ("svg:title", frame.title.as_deref()),
        ("svg:desc", frame.description.as_deref()),
        ("text:anchor-type", frame.anchor_type.as_deref()),
        ("table:end-cell-address", frame.end_cell_address.as_deref()),
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
        crate::ImageSource::Inline {
            bytes,
            ignored_href,
        } => {
            if bytes.len() > MAX_INLINE_BYTES {
                return invalid("inline sheet image exceeds 16 MiB");
            }
            if let Some(href) = ignored_href {
                validate_text(href, "ignored xlink:href", true)?;
            }
        },
        crate::ImageSource::PackagePart { href, .. }
        | crate::ImageSource::MissingPackagePart { href, .. }
        | crate::ImageSource::Linked { href } => {
            validate_text(href, "xlink:href", false)?;
        },
        crate::ImageSource::Missing => return invalid("sheet image requires a source"),
    }
    Ok(())
}

pub(crate) fn validate_sheet_images(images: &[crate::Image]) -> Result<()> {
    if images.len() > MAX_IMAGES_PER_SHEET {
        return invalid(format!("sheet exceeds {MAX_IMAGES_PER_SHEET} images"));
    }
    let mut previous: Option<&crate::Image> = None;
    let mut total_inline_bytes = 0usize;
    for image in images {
        validate_sheet_image(image)?;
        if let crate::ImageSource::Inline { bytes, .. } = &image.source {
            total_inline_bytes = total_inline_bytes.checked_add(bytes.len()).ok_or_else(|| {
                Error::InvalidFormat("total inline sheet image data size overflow".to_string())
            })?;
            if total_inline_bytes > MAX_TOTAL_INLINE_BYTES {
                return invalid(format!(
                    "total inline sheet image data exceeds {MAX_TOTAL_INLINE_BYTES} bytes"
                ));
            }
        }
        match image.alternative_index {
            0 => {},
            index => {
                let prior = previous.ok_or_else(|| {
                    Error::InvalidFormat(
                        "sheet image alternative has no preceding primary image".to_string(),
                    )
                })?;
                if prior.alternative_index.checked_add(1) != Some(index) {
                    return invalid("sheet image alternative indices must be contiguous");
                }
                if prior.frame != image.frame {
                    return invalid("sheet image alternatives must share one draw:frame");
                }
            },
        }
        previous = Some(image);
    }
    Ok(())
}

fn validate_sheet_images_for_sheet(images: &[crate::Image], sheet_name: &str) -> Result<()> {
    validate_sheet_images(images)?;
    if images.iter().any(|image| {
        image
            .frame
            .as_ref()
            .and_then(|frame| frame.sheet_name.as_deref())
            != Some(sheet_name)
    }) {
        return invalid("sheet image frame belongs to a different sheet");
    }
    Ok(())
}

fn frame_group_end(images: &[crate::Image], primary_image_index: usize) -> Result<usize> {
    let primary = images.get(primary_image_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "sheet image group index {primary_image_index} out of bounds"
        ))
    })?;
    if primary.alternative_index != 0 {
        return invalid(format!(
            "sheet image index {primary_image_index} is not a frame-group primary"
        ));
    }
    Ok(images[primary_image_index + 1..]
        .iter()
        .position(|image| image.alternative_index == 0)
        .map_or(images.len(), |offset| primary_image_index + 1 + offset))
}

pub(crate) fn insert_sheet_image_alternative(
    images: &mut Vec<crate::Image>,
    sheet_name: &str,
    primary_image_index: usize,
    alternative_index: usize,
    image: crate::Image,
) -> Result<()> {
    validate_sheet_images_for_sheet(images, sheet_name)?;
    if images.len() >= MAX_IMAGES_PER_SHEET {
        return invalid(format!("sheet exceeds {MAX_IMAGES_PER_SHEET} images"));
    }
    let group_end = frame_group_end(images, primary_image_index)?;
    let next_alternative_index = group_end - primary_image_index;
    if !(1..=next_alternative_index).contains(&alternative_index) {
        return invalid(format!(
            "sheet image alternative index {alternative_index} out of bounds for frame group"
        ));
    }

    let mut image = normalize_sheet_image(image, sheet_name)?;
    if image.frame != images[primary_image_index].frame {
        return invalid("sheet image alternative frame does not match its primary frame");
    }
    image.alternative_index = alternative_index;

    let mut candidate = images.clone();
    candidate.insert(primary_image_index + alternative_index, image);
    let new_group_end = group_end + 1;
    for (index, image) in candidate[primary_image_index..new_group_end]
        .iter_mut()
        .enumerate()
    {
        image.alternative_index = index;
    }
    validate_sheet_images_for_sheet(&candidate, sheet_name)?;
    *images = candidate;
    Ok(())
}

pub(crate) fn append_sheet_image_alternative(
    images: &mut Vec<crate::Image>,
    sheet_name: &str,
    primary_image_index: usize,
    image: crate::Image,
) -> Result<()> {
    validate_sheet_images_for_sheet(images, sheet_name)?;
    let group_end = frame_group_end(images, primary_image_index)?;
    insert_sheet_image_alternative(
        images,
        sheet_name,
        primary_image_index,
        group_end - primary_image_index,
        image,
    )
}

pub(crate) fn remove_sheet_image_alternative(
    images: &mut Vec<crate::Image>,
    sheet_name: &str,
    primary_image_index: usize,
    alternative_index: usize,
) -> Result<crate::Image> {
    validate_sheet_images_for_sheet(images, sheet_name)?;
    let group_end = frame_group_end(images, primary_image_index)?;
    let group_len = group_end - primary_image_index;
    if alternative_index == 0 || alternative_index >= group_len {
        return invalid(format!(
            "sheet image alternative index {alternative_index} out of bounds for frame group"
        ));
    }

    let mut candidate = images.clone();
    let removed = candidate.remove(primary_image_index + alternative_index);
    for (index, image) in candidate[primary_image_index..group_end - 1]
        .iter_mut()
        .enumerate()
    {
        image.alternative_index = index;
    }
    validate_sheet_images_for_sheet(&candidate, sheet_name)?;
    *images = candidate;
    Ok(removed)
}

pub(crate) fn write_sheet_images_content(out: &mut String, images: &[crate::Image]) -> Result<()> {
    if images.is_empty() {
        return Ok(());
    }
    validate_sheet_images(images)?;
    for (index, image) in images.iter().enumerate() {
        let frame = image.frame.as_ref().expect("validated frame");
        if image.alternative_index == 0 {
            out.push_str("<draw:frame");
            attribute(out, "draw:name", frame.name.as_deref());
            attribute(out, "xml:id", frame.xml_id.as_deref());
            attribute(out, "text:anchor-type", frame.anchor_type.as_deref());
            attribute(out, "svg:x", frame.x.as_deref());
            attribute(out, "svg:y", frame.y.as_deref());
            attribute(out, "svg:width", frame.width.as_deref());
            attribute(out, "svg:height", frame.height.as_deref());
            attribute(
                out,
                "table:end-cell-address",
                frame.end_cell_address.as_deref(),
            );
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
        }
        out.push_str("<draw:image");
        attribute(out, "xml:id", image.xml_id.as_deref());
        attribute(out, "draw:filter-name", image.filter_name.as_deref());
        attribute(out, "draw:mime-type", image.declared_media_type.as_deref());
        let href = match &image.source {
            crate::ImageSource::Inline { ignored_href, .. } => ignored_href.as_deref(),
            crate::ImageSource::PackagePart { href, .. }
            | crate::ImageSource::MissingPackagePart { href, .. }
            | crate::ImageSource::Linked { href } => Some(href.as_str()),
            crate::ImageSource::Missing => None,
        };
        attribute(out, "xlink:href", href);
        attribute(out, "xlink:type", image.link_type.as_deref());
        attribute(out, "xlink:show", image.show.as_deref());
        attribute(out, "xlink:actuate", image.actuate.as_deref());
        match &image.source {
            crate::ImageSource::Inline { bytes, .. } => {
                out.push_str("><office:binary-data>");
                out.push_str(&BASE64_STANDARD.encode(bytes));
                out.push_str("</office:binary-data></draw:image>");
            },
            _ => out.push_str("/>"),
        }
        if images
            .get(index + 1)
            .is_none_or(|next| next.alternative_index == 0)
        {
            out.push_str("</draw:frame>");
        }
    }
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

pub(crate) fn validate_length(value: &str, name: &str, nonnegative: bool) -> Result<()> {
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

pub(crate) fn validate_text(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds {MAX_VALUE_BYTES} bytes"));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return invalid(format!("{name} contains invalid XML characters"));
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn image(alternative_index: usize, href: &str) -> crate::Image {
        crate::Image {
            part: crate::ImagePart::Content,
            source: crate::ImageSource::Linked {
                href: href.to_string(),
            },
            frame: Some(crate::ImageFrame {
                name: Some("alternatives".to_string()),
                sheet_name: Some("Sheet1".to_string()),
                sheet_shape: true,
                width: Some("2cm".to_string()),
                height: Some("1cm".to_string()),
                ..Default::default()
            }),
            xml_id: None,
            filter_name: None,
            declared_media_type: None,
            link_type: Some("simple".to_string()),
            show: Some("embed".to_string()),
            actuate: Some("onLoad".to_string()),
            alternative_index,
        }
    }

    #[test]
    fn writes_ordered_alternatives_in_one_frame() {
        let images = [
            image(0, "Pictures/vector.svg"),
            image(1, "Pictures/fallback.png"),
        ];
        let mut xml = String::new();
        write_sheet_images_content(&mut xml, &images).unwrap();
        assert_eq!(xml.matches("<draw:frame").count(), 1);
        assert_eq!(xml.matches("<draw:image").count(), 2);
        assert!(xml.find("vector.svg").unwrap() < xml.find("fallback.png").unwrap());
        assert!(xml.ends_with("</draw:frame>"));
    }

    #[test]
    fn rejects_broken_alternative_groups() {
        assert!(validate_sheet_images(&[image(1, "fallback.png")]).is_err());
        assert!(
            validate_sheet_images(&[image(0, "primary.svg"), image(2, "fallback.png")]).is_err()
        );
        let primary = image(0, "primary.svg");
        let mut alternative = image(1, "fallback.png");
        alternative.frame.as_mut().unwrap().name = Some("other".to_string());
        assert!(validate_sheet_images(&[primary, alternative]).is_err());
    }
}
