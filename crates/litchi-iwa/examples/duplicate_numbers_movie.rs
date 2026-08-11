//! Duplicate an ordinary file-backed Numbers sheet movie.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_numbers_movie <input.numbers> <output.numbers> <sheet-index> <movie-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index is out of bounds")?;
    let source = editor
        .sheet_movies(sheet.id())?
        .into_iter()
        .nth(movie_index)
        .ok_or("ordinary movie index is out of bounds")?;
    let created = editor.duplicate_sheet_movie(sheet.id(), source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "sheet={} drawable={} source={} video={} poster={}",
        sheet.id(),
        created.drawable_object_id,
        source.drawable_object_id,
        created.movie_data_identifier,
        created.poster_image_data_identifier,
    );
    Ok(())
}
