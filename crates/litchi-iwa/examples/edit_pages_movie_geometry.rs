//! Update one body-anchored Pages movie's position and displayed size.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_movie_geometry <input.pages> <output.pages> <movie-index> <x> <y> <width> <height>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
    let position = DrawablePoint {
        x: arguments.next().ok_or("missing x")?.parse()?,
        y: arguments.next().ok_or("missing y")?.parse()?,
    };
    let size = DrawableSize {
        width: arguments.next().ok_or("missing width")?.parse()?,
        height: arguments.next().ok_or("missing height")?.parse()?,
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let movie = editor
        .body_movies()?
        .get(movie_index)
        .cloned()
        .ok_or("movie index is out of bounds")?;
    let geometry = litchi_iwa::shapes::DrawableGeometry {
        position: Some(position),
        size: Some(size),
        ..movie.geometry
    };
    editor.set_body_movie_geometry(movie.drawable_object_id, geometry)?;
    editor.save(output)?;
    println!(
        "drawable={} geometry={geometry:?}",
        movie.drawable_object_id
    );
    Ok(())
}
