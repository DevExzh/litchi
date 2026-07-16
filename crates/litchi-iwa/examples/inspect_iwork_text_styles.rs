//! List effective uniform character formatting in every text storage.

use std::env;

use litchi_iwa::text::IWorkTextEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: inspect_iwork_text_styles <input.pages|input.numbers|input.key>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let editor = IWorkTextEditor::open(input)?;
    for storage in editor.storages()? {
        match editor.text_style(storage.object_id) {
            Ok(style) => println!(
                "storage={} points={} bold={} italic={}",
                storage.object_id,
                style.point_size.points(),
                style.bold,
                style.italic
            ),
            Err(error) => println!("storage={} unavailable={error}", storage.object_id),
        }
    }
    Ok(())
}
