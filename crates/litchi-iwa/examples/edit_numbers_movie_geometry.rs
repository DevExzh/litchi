//! Update one Numbers movie's position and displayed size.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_movie_geometry <input.numbers> <output.numbers> <sheet-index> <movie-index> <x> <y> <width> <height>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
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

    let mut editor = NumbersEditor::open(input)?;
    let sheet_id = editor
        .sheets()?
        .get(sheet_index)
        .ok_or("sheet index is out of bounds")?
        .id();
    let movie = editor
        .sheet_movies(sheet_id)?
        .get(movie_index)
        .cloned()
        .ok_or("movie index is out of bounds")?;
    let geometry = litchi_iwa::shapes::DrawableGeometry {
        position: Some(position),
        size: Some(size),
        ..movie.geometry
    };
    editor.set_sheet_movie_geometry(sheet_id, movie.drawable_object_id, geometry)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_id} drawable={} geometry={geometry:?}",
        movie.drawable_object_id
    );
    Ok(())
}
