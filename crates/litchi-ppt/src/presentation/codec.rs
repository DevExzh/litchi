use crate::consts::RecordType;
use crate::header_footer::HeaderFooterDisplayText;
use crate::package::Result;
use crate::records::Record;
use crate::slide::Slide;

/// Collect text boxes from a shared OfficeArt shape tree without owning the
/// OfficeArt model in this format-specific crate.
pub(super) fn drawing_textboxes(data: &[u8]) -> Result<Vec<litchi_odraw::Record<'_>>> {
    fn collect<'data>(
        shape: &litchi_odraw::shape::Shape<'data>,
        records: &mut Vec<litchi_odraw::Record<'data>>,
    ) -> Result<()> {
        if let Some(textbox) = crate::odraw::textbox(shape)? {
            records.push(textbox);
        }
        for child in shape.children() {
            collect(child, records)?;
        }
        Ok(())
    }

    let shapes = crate::odraw::parse(data)?;
    let mut records = Vec::new();
    for shape in &shapes {
        collect(shape, &mut records)?;
    }
    Ok(records)
}

pub(super) fn placeholder_display_from_record(
    record: &Record,
) -> Result<Option<HeaderFooterDisplayText>> {
    let Some(drawing) = record.find_child(RecordType::PPDrawing) else {
        return Ok(None);
    };
    let parsed = crate::odraw::parse(&drawing.data)?;
    let mut shapes = Vec::with_capacity(parsed.len());
    for shape in &parsed {
        if let Some(shape) = Slide::<'static>::convert_odraw_to_shape_enum(shape)? {
            shapes.push(shape);
        }
    }
    placeholder_display_from_shapes(&shapes)
}

pub(super) fn placeholder_display_from_shapes(
    shapes: &[crate::shapes::ShapeEnum<'static>],
) -> Result<Option<HeaderFooterDisplayText>> {
    use crate::shapes::PlaceholderType;

    let mut display = HeaderFooterDisplayText::default();
    for shape in shapes {
        let Some(placeholder) = shape.as_placeholder() else {
            continue;
        };
        let target = match placeholder.placeholder_type() {
            PlaceholderType::DateAndTime => &mut display.user_date,
            PlaceholderType::Header => &mut display.header,
            PlaceholderType::Footer => &mut display.footer,
            _ => continue,
        };
        if target.is_some() {
            continue;
        }
        let text = shape.text()?;
        if !text.is_empty() && text != "*" {
            *target = Some(text);
        }
    }
    if display == HeaderFooterDisplayText::default() {
        Ok(None)
    } else {
        Ok(Some(display))
    }
}
