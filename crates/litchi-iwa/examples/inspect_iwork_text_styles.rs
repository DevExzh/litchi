//! List effective uniform character formatting in every text storage.

use std::env;

use litchi_iwa::text::IWorkTextEditor;
use litchi_iwa_text::comment::raw::comment_id_value;

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
        if let Ok(languages) = editor.text_languages(storage.id) {
            println!("storage={} languages={languages:?}", storage.id);
        }
        if let Ok(hyperlinks) = editor.text_hyperlinks(storage.id) {
            println!("storage={} hyperlinks={hyperlinks:?}", storage.id);
        }
        if let Ok(highlights) = editor.text_highlights(storage.id) {
            println!("storage={} highlights={highlights:?}", storage.id);
        }
        if let Ok(comments) = editor.text_comments(storage.id) {
            println!("storage={} comments={comments:?}", storage.id);
            for comment in comments {
                if let Ok(replies) = editor.text_comment_replies(storage.id, comment.id()) {
                    println!(
                        "storage={} comment={} replies={replies:?}",
                        storage.id,
                        comment_id_value(comment.id())
                    );
                }
            }
        }
        if let Ok(levels) = editor.paragraph_list_levels(storage.id) {
            println!("storage={} list_levels={levels:?}", storage.id);
        }
        if let Ok(flow) = editor.paragraph_flow(storage.id) {
            println!("storage={} paragraph_flow={flow:?}", storage.id);
        }
        match (
            editor.text_style(storage.id),
            editor.text_font(storage.id),
            editor.text_decorations(storage.id),
            editor.text_color(storage.id),
            editor.text_capitalization(storage.id),
            editor.text_script(storage.id),
            editor.text_baseline_shift(storage.id),
            editor.text_character_spacing(storage.id),
            editor.text_ligatures(storage.id),
            editor.text_outline(storage.id),
            editor.text_shadow(storage.id),
            editor.text_background(storage.id),
            editor.paragraph_list(storage.id),
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
                storage.id,
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
                println!("storage={} unavailable={error}", storage.id)
            },
        }
    }
    Ok(())
}
