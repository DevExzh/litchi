use crate::consts::RecordType;
use crate::header_footer::HeaderFooterDisplayText;
use crate::package::Result;
use crate::records::Record;
use crate::slide::Slide;

/// Collect text boxes from a shared `OfficeArt` shape tree without owning the
/// `OfficeArt` model in this format-specific crate.
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
    // The header/footer display-text projection re-parses drawings whose
    // text-style and shape repairs (see `crate::leniency`) are enumerated by
    // `Presentation::collect_tolerance_report`: slide drawings through each
    // slide's shapes, main-master drawings through
    // `record_drawing_repairs`. They are not recorded a second time here.
    let Some(shapes) = drawing_shapes(record, &mut crate::leniency::ToleranceReport::default())?
    else {
        return Ok(None);
    };
    placeholder_display_from_shapes(&shapes)
}

/// Parse the shapes of `record`'s `PPDrawing`, recording every text-style
/// and shape repair into `report`. `None` when the record has no drawing.
fn drawing_shapes(
    record: &Record,
    report: &mut crate::leniency::ToleranceReport,
) -> Result<Option<Vec<crate::shapes::ShapeEnum<'static>>>> {
    let Some(drawing) = record.find_child(RecordType::PPDrawing) else {
        return Ok(None);
    };
    let parsed = crate::odraw::parse(&drawing.data)?;
    let mut shapes = Vec::with_capacity(parsed.len());
    for (index, shape) in parsed.iter().enumerate() {
        if let Some(converted) =
            Slide::<'static>::convert_odraw_to_shape_enum(shape, index, report)?
        {
            shapes.push(converted);
        }
    }
    Ok(Some(shapes))
}

/// Record the text-style and shape repairs of `record`'s drawing (a main
/// master) into `report`.
pub(super) fn record_drawing_repairs(
    record: &Record,
    report: &mut crate::leniency::ToleranceReport,
) -> Result<()> {
    drawing_shapes(record, report).map(|_| ())
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
        let placeholder_type = placeholder.placeholder_type();
        let target = if placeholder_type == PlaceholderType::DateAndTime {
            &mut display.user_date
        } else if placeholder_type == PlaceholderType::Header {
            &mut display.header
        } else if placeholder_type == PlaceholderType::Footer {
            &mut display.footer
        } else {
            continue;
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
