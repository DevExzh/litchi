//! Edit geometry of an ordinary file-backed Keynote slide movie.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_keynote::slide::media::MovieKind;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_movie_geometry <input.key> <output.key> <slide-index> <movie-index> <x> <y> <width> <height> <angle>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
    let x: f32 = arguments.next().ok_or("missing x")?.parse()?;
    let y: f32 = arguments.next().ok_or("missing y")?.parse()?;
    let width: f32 = arguments.next().ok_or("missing width")?.parse()?;
    let height: f32 = arguments.next().ok_or("missing height")?.parse()?;
    let angle: f32 = arguments.next().ok_or("missing angle")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let source = editor
        .slide_movies(slide_index)?
        .into_iter()
        .filter(|movie| movie.kind == MovieKind::File)
        .nth(movie_index)
        .ok_or("ordinary movie index is out of bounds")?;
    let mut geometry = source.geometry;
    geometry.position = Some(DrawablePoint { x, y });
    geometry.size = Some(DrawableSize { width, height });
    geometry.angle = Some(angle);
    editor.set_slide_movie_geometry(slide_index, source.drawable_object_id, geometry)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} geometry={geometry:?}",
        source.drawable_object_id
    );
    Ok(())
}
