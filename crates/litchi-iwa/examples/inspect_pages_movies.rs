//! Inspect ordinary file-backed movies anchored to the Pages body.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_movies <input.pages>")?;
    let editor = PagesEditor::open(input)?;
    for (movie_index, movie) in editor.body_movies()?.iter().enumerate() {
        println!(
            "movie_index={movie_index} anchor={} drawable={} movie_data={} poster_data={} duration={:?} geometry={:?} original_size={:?} natural_size={:?}",
            movie.anchor_character_index,
            movie.drawable_object_id,
            movie.movie_data_identifier,
            movie.poster_image_data_identifier,
            movie.duration,
            movie.geometry,
            movie.original_size,
            movie.natural_size
        );
    }
    Ok(())
}
