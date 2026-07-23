//! Create a Keynote movie and apply a native horizontal Arrange flip from scratch.

use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteSlideMovieOptions};
use litchi_iwa::shapes::{DrawableFlipAxis, DrawablePoint, DrawableSize};

const MOVIE_DURATION: Duration = Duration::from_secs(8);
const MOVIE_POSITION: DrawablePoint = DrawablePoint { x: 800.0, y: 450.0 };
const MOVIE_SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 180.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_flipped_movie <output.key> <movie> <poster-image>")?;
    let movie_path = arguments
        .next()
        .ok_or("usage: create_keynote_flipped_movie <output.key> <movie> <poster-image>")?;
    let poster_path = arguments
        .next()
        .ok_or("usage: create_keynote_flipped_movie <output.key> <movie> <poster-image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let movie_filename = filename(&movie_path, "movie")?;
    let poster_filename = filename(&poster_path, "poster")?;
    let movie = fs::read(&movie_path)?;
    let poster = fs::read(&poster_path)?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Flipped Movie")
        .subtitle("Typed native Arrange flip")
        .build()?;
    let created = editor.add_slide_movie(
        0,
        movie_filename,
        &movie,
        poster_filename,
        &poster,
        KeynoteSlideMovieOptions::new(MOVIE_POSITION, MOVIE_SIZE, MOVIE_DURATION),
    )?;
    editor.flip_slide_movie(0, created.drawable_object_id, DrawableFlipAxis::Horizontal)?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Keynote movie {}",
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
