//! List slide-owned Keynote text targets and ordinary text-box CLI indexes.

use std::env;

use litchi_iwa::keynote::{KeynoteEditor, KeynoteSlideTextRole};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_keynote_text_boxes <input.key>")?;
    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        println!(
            "slide={} node={} object={} name={:?}",
            slide.index, slide.node_id, slide.slide_id, slide.name
        );
        let mut text_box_index = 0usize;
        for (text_index, text) in editor
            .slide_text_storages(slide.index)?
            .into_iter()
            .enumerate()
        {
            let ordinary_index = (text.role == KeynoteSlideTextRole::TextBox).then(|| {
                let index = text_box_index;
                text_box_index += 1;
                index
            });
            let geometry = if text.role == KeynoteSlideTextRole::TextBox {
                Some(editor.slide_text_box_geometry(slide.index, text.drawable_object_id)?)
            } else {
                None
            };
            let properties = if text.role == KeynoteSlideTextRole::TextBox {
                Some(editor.slide_text_box_properties(slide.index, text.drawable_object_id)?)
            } else {
                None
            };
            let columns = if text.role == KeynoteSlideTextRole::TextBox {
                Some(editor.slide_text_box_columns(slide.index, text.drawable_object_id)?)
            } else {
                None
            };
            let text_layout = if text.role == KeynoteSlideTextRole::TextBox {
                Some(editor.slide_text_box_text_layout(slide.index, text.drawable_object_id)?)
            } else {
                None
            };
            println!(
                "  text_index={text_index} text_box_index={ordinary_index:?} role={:?} drawable={} storage={} text={:?} geometry={geometry:?} properties={properties:?} columns={columns:?} text_layout={text_layout:?}",
                text.role,
                text.drawable_object_id,
                text.storage.object_id,
                text.storage.storage.text()
            );
        }
    }
    Ok(())
}
