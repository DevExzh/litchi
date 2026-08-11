//! Inspect ordinary file-backed movies owned by Numbers sheets.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_numbers_movies <input.numbers>")?;
    let editor = NumbersEditor::open(input)?;
    for sheet in editor.sheets()? {
        for (movie_index, movie) in editor.sheet_movies(sheet.id())?.iter().enumerate() {
            println!(
                "sheet={} movie_index={movie_index} drawable={} movie_data={} poster_data={} duration={:?} geometry={:?} original_size={:?} natural_size={:?}",
                sheet.id(),
                movie.drawable_object_id,
                movie.movie_data_identifier,
                movie.poster_image_data_identifier,
                movie.duration,
                movie.geometry,
                movie.original_size,
                movie.natural_size
            );
        }
    }
    Ok(())
}
