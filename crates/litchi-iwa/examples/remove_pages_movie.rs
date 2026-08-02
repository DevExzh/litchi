//! Remove one body-anchored Pages movie and its private graph.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_pages_movie <input.pages> <output.pages> <movie-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let movie = editor
        .body_movies()?
        .get(movie_index)
        .cloned()
        .ok_or("movie index is out of bounds")?;
    let removed = editor.remove_body_movie(movie.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "drawable={} removed_data={:?}",
        removed.movie.drawable_object_id, removed.removed_data_identifiers
    );
    Ok(())
}
