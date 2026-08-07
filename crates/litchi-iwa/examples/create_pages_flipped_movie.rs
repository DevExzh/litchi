//! Create a Pages movie and apply a native horizontal Arrange flip from scratch.

use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::DrawableFlipAxis;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_pages::movie::Options as PagesMovieOptions;

const BODY_TEXT: &str = "This mirrored movie was created entirely by litchi-iwa.";
const MOVIE_DURATION: Duration = Duration::from_secs(8);
const MOVIE_POSITION: Point = Point { x: 96.0, y: 144.0 };
const MOVIE_SIZE: Size = Size {
    width: 320.0,
    height: 180.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_flipped_movie <output.pages> <movie> <poster-image>")?;
    let movie_path = arguments
        .next()
        .ok_or("usage: create_pages_flipped_movie <output.pages> <movie> <poster-image>")?;
    let poster_path = arguments
        .next()
        .ok_or("usage: create_pages_flipped_movie <output.pages> <movie> <poster-image>")?;
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
        PagesMovieOptions::new(MOVIE_POSITION, MOVIE_SIZE, MOVIE_DURATION)?,
    )?;
    editor.flip_body_movie(created.drawable_object_id, DrawableFlipAxis::Horizontal)?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Pages movie {}",
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
