//! Create a Pages movie and restore its typed original dimensions from scratch.

use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::pages::{PagesEditor, PagesMovieOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const BODY_TEXT: &str = "This movie was restored to its original size by litchi-iwa.";
const MOVIE_DURATION: Duration = Duration::from_secs(8);
const MOVIE_POSITION: DrawablePoint = DrawablePoint { x: 64.0, y: 128.0 };
const DISPLAYED_MOVIE_SIZE: DrawableSize = DrawableSize {
    width: 240.0,
    height: 135.0,
};
const ORIGINAL_MOVIE_SIZE: DrawableSize = DrawableSize {
    width: 480.0,
    height: 270.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_original_size_movie <output.pages> <movie> <poster-image>")?;
    let movie_path = arguments
        .next()
        .ok_or("usage: create_pages_original_size_movie <output.pages> <movie> <poster-image>")?;
    let poster_path = arguments
        .next()
        .ok_or("usage: create_pages_original_size_movie <output.pages> <movie> <poster-image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let movie_filename = filename(&movie_path, "movie")?;
    let poster_filename = filename(&poster_path, "poster")?;
    let movie = fs::read(&movie_path)?;
    let poster = fs::read(&poster_path)?;

    let mut editor = PagesEditor::create_with_text(BODY_TEXT)?;
    let created = editor.add_body_movie(
        BODY_TEXT.encode_utf16().count(),
        movie_filename,
        &movie,
        poster_filename,
        &poster,
        PagesMovieOptions::new(MOVIE_POSITION, DISPLAYED_MOVIE_SIZE, MOVIE_DURATION)
            .with_natural_size(ORIGINAL_MOVIE_SIZE),
    )?;
    editor.restore_body_movie_original_size(created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Pages movie {} at its original size",
        created.drawable_object_id
    );
    Ok(())
}

fn filename<'a>(path: &'a str, kind: &str) -> Result<&'a str, Box<dyn std::error::Error>> {
    Path::new(path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| format!("{kind} path must end in a UTF-8 file name").into())
}
