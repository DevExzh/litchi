//! Duplicate an ordinary file-backed Keynote slide movie.

use std::env;

use litchi_iwa::keynote::{KeynoteEditor, KeynoteSlideMovieKind};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_keynote_movie <input.key> <output.key> <slide-index> <movie-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let source = editor
        .slide_movies(slide_index)?
        .into_iter()
        .filter(|movie| movie.kind == KeynoteSlideMovieKind::File)
        .nth(movie_index)
        .ok_or("ordinary movie index is out of bounds")?;
    let created = editor.duplicate_slide_movie(slide_index, source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} source={} video={:?} poster={:?}",
        created.drawable_object_id,
        source.drawable_object_id,
        created.movie_data_identifier,
        created.poster_image_data_identifier,
    );
    Ok(())
}
