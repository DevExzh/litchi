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
        if let Ok(languages) = editor.text_languages(storage.object_id) {
            println!("storage={} languages={languages:?}", storage.object_id);
        }
        if let Ok(levels) = editor.paragraph_list_levels(storage.object_id) {
            println!("storage={} list_levels={levels:?}", storage.object_id);
        }
        match (
            editor.text_style(storage.object_id),
            editor.text_font(storage.object_id),
            editor.text_decorations(storage.object_id),
            editor.text_color(storage.object_id),
            editor.text_capitalization(storage.object_id),
            editor.text_script(storage.object_id),
            editor.text_baseline_shift(storage.object_id),
            editor.text_character_spacing(storage.object_id),
            editor.text_ligatures(storage.object_id),
            editor.text_outline(storage.object_id),
            editor.text_shadow(storage.object_id),
            editor.text_background(storage.object_id),
            editor.paragraph_list(storage.object_id),
        ) {
            (
                Ok(style),
                Ok(font),
                Ok(decorations),
                Ok(color),
                Ok(capitalization),
                Ok(script),
                Ok(baseline_shift),
                Ok(character_spacing),
                Ok(ligatures),
                Ok(outline),
                Ok(shadow),
                Ok(background),
                Ok(list),
            ) => println!(
                "storage={} font={font:?} points={} bold={} italic={} underline={:?} strikethrough={:?} rgba=({},{},{},{}) color_space={:?} capitalization={capitalization:?} script={script:?} baseline_shift_points={} character_spacing_percent={} ligatures={ligatures:?} outline={outline:?} shadow={shadow:?} background={background:?} list={list:?}",
                storage.object_id,
                style.point_size.points(),
                style.bold,
                style.italic,
                decorations.underline,
                decorations.strikethrough,
                color.red(),
                color.green(),
                color.blue(),
                color.alpha(),
                color.color_space(),
                baseline_shift.points(),
                character_spacing.percent()
            ),
            (Err(error), _, _, _, _, _, _, _, _, _, _, _, _)
            | (_, Err(error), _, _, _, _, _, _, _, _, _, _, _)
            | (_, _, Err(error), _, _, _, _, _, _, _, _, _, _)
            | (_, _, _, Err(error), _, _, _, _, _, _, _, _, _)
            | (_, _, _, _, Err(error), _, _, _, _, _, _, _, _)
            | (_, _, _, _, _, Err(error), _, _, _, _, _, _, _)
            | (_, _, _, _, _, _, Err(error), _, _, _, _, _, _)
            | (_, _, _, _, _, _, _, Err(error), _, _, _, _, _)
            | (_, _, _, _, _, _, _, _, Err(error), _, _, _, _)
            | (_, _, _, _, _, _, _, _, _, Err(error), _, _, _)
            | (_, _, _, _, _, _, _, _, _, _, Err(error), _, _)
            | (_, _, _, _, _, _, _, _, _, _, _, Err(error), _)
            | (_, _, _, _, _, _, _, _, _, _, _, _, Err(error)) => {
                println!("storage={} unavailable={error}", storage.object_id)
            },
        }
    }
    Ok(())
}
