//! Create a Keynote movie and restore its typed original dimensions from scratch.

use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::slide::movie::Options as SlideMovieOptions;

const MOVIE_DURATION: Duration = Duration::from_secs(8);
const MOVIE_POSITION: Point = Point { x: 720.0, y: 405.0 };
const DISPLAYED_MOVIE_SIZE: Size = Size {
    width: 240.0,
    height: 135.0,
};
const ORIGINAL_MOVIE_SIZE: Size = Size {
    width: 480.0,
    height: 270.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_original_size_movie <output.key> <movie> <poster-image>")?;
    let movie_path = arguments
        .next()
        .ok_or("usage: create_keynote_original_size_movie <output.key> <movie> <poster-image>")?;
    let poster_path = arguments
        .next()
        .ok_or("usage: create_keynote_original_size_movie <output.key> <movie> <poster-image>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let movie_filename = filename(&movie_path, "movie")?;
    let poster_filename = filename(&poster_path, "poster")?;
    let movie = fs::read(&movie_path)?;
    let poster = fs::read(&poster_path)?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Original Size Movie")
        .subtitle("Typed native media dimensions")
        .build()?;
    let created = editor.add_slide_movie(
        0,
        movie_filename,
        &movie,
        poster_filename,
        &poster,
        SlideMovieOptions::new(MOVIE_POSITION, DISPLAYED_MOVIE_SIZE, MOVIE_DURATION)?
            .with_natural_size(ORIGINAL_MOVIE_SIZE)?,
    )?;
    editor.restore_slide_movie_original_size(0, created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Keynote movie {} at its original size",
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
