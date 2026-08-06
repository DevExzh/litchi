use super::model::*;

pub(crate) fn validate_story_events(
    text: &str,
    shapes: &[crate::Shape<'_>],
    shape_groups: &[crate::ShapeGroup<'_>],
    drawing_order: &[crate::StoryDrawing],
    events: &[StoryEvent],
    label: &str,
) -> crate::RtfResult<()> {
    crate::shape::validate_story_drawings(text, shapes, shape_groups, drawing_order, label)?;
    let mut drawings = Vec::with_capacity(drawing_order.len());
    let mut fields = std::collections::BTreeSet::new();
    let mut previous = None;
    for event in events {
        let position = match *event {
            StoryEvent::PageBreak(value) => {
                if text.get(value.position..value.position).is_none() {
                    return Err(crate::RtfError::MalformedDocument(format!(
                        "RTF {label} page break is not at a UTF-8 boundary"
                    )));
                }
                value.position
            },
            StoryEvent::Drawing(drawing) => {
                drawings.push(drawing);
                match drawing {
                    crate::StoryDrawing::Shape(index) => {
                        shapes.get(index).map(|shape| shape.position)
                    },
                    crate::StoryDrawing::ShapeGroup(index) => {
                        shape_groups.get(index).map(|group| group.position)
                    },
                }
                .ok_or_else(|| {
                    crate::RtfError::MalformedDocument(format!(
                        "RTF {label} story order has an invalid drawing reference"
                    ))
                })?
            },
            StoryEvent::Field(field) => {
                if !fields.insert(field.field_index)
                    || text.get(field.position..field.position).is_none()
                {
                    return Err(crate::RtfError::MalformedDocument(format!(
                        "RTF {label} story order has an invalid or duplicate field reference"
                    )));
                }
                field.position
            },
        };
        if previous.is_some_and(|value| value > position) {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF {label} story order moves backwards"
            )));
        }
        previous = Some(position);
    }
    if drawings != drawing_order {
        return Err(crate::RtfError::MalformedDocument(format!(
            "RTF {label} story order is incomplete or changes drawing order"
        )));
    }
    Ok(())
}

pub(crate) fn push_story_page_break(
    events: &mut Vec<StoryEvent>,
    text: &str,
    position: usize,
    label: &str,
) -> crate::RtfResult<()> {
    if text.get(position..position).is_none() {
        return Err(crate::RtfError::MalformedDocument(format!(
            "RTF {label} page break is not at a UTF-8 boundary"
        )));
    }
    events.push(StoryEvent::PageBreak(PageBreak::new(position)));
    Ok(())
}
