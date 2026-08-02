//! Duplicate an ordinary file-backed Pages body movie.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_pages_movie <input.pages> <output.pages> <movie-index> <utf16-anchor>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
    let anchor: usize = arguments.next().ok_or("missing UTF-16 anchor")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let source = editor
        .body_movies()?
        .into_iter()
        .nth(movie_index)
        .ok_or("ordinary movie index is out of bounds")?;
    let created = editor.duplicate_body_movie(source.drawable_object_id, anchor)?;
    editor.save(output)?;
    println!(
        "anchor={} drawable={} source={} video={} poster={}",
        created.anchor_character_index,
        created.drawable_object_id,
        source.drawable_object_id,
        created.movie_data_identifier,
        created.poster_image_data_identifier,
    );
    Ok(())
}
