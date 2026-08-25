//! Edit geometry of an ordinary file-backed Keynote slide movie.

use std::{env, fs::File};

use litchi_keynote::slide::media::{Point, Size, geometry::MovieGeometry};
use litchi_keynote::{MovieSelector, Package, SlideSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_movie_geometry <input.key> <output.key> <slide-index> <movie-index> <x> <y> <width> <height>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
    let x: f32 = arguments.next().ok_or("missing x")?.parse()?;
    let y: f32 = arguments.next().ok_or("missing y")?.parse()?;
    let width: f32 = arguments.next().ok_or("missing width")?.parse()?;
    let height: f32 = arguments.next().ok_or("missing height")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let package = Package::open(input)?;
    let geometry = MovieGeometry::new(Point { x, y }, Size { width, height })?;
    let commit = package
        .edit_slide_movie_geometry(
            SlideSelector::index(slide_index),
            MovieSelector::index(movie_index),
        )?
        .set(geometry)?
        .commit()?;
    let mut output = File::create(output)?;
    commit.package().write_to(&mut output)?;
    println!("slide={slide_index} movie={movie_index} geometry={geometry:?}",);
    Ok(())
}
