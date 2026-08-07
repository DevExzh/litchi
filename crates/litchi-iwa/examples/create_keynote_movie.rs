//! Create a Keynote presentation with a movie and no input package.

use std::env;
use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa_common::media::playback::{MediaLoopMode, MediaVolume};
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::slide::movie::Options as SlideMovieOptions;

const SLIDE_WIDTH_POINTS: f32 = 1_920.0;
const SLIDE_HEIGHT_POINTS: f32 = 1_080.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments.next().ok_or(
        "usage: create_keynote_movie <output.key> <movie> <poster> <width> <height> <duration-seconds>",
    )?;
    let movie_path = arguments.next().ok_or("missing movie path")?;
    let poster_path = arguments.next().ok_or("missing poster path")?;
    let width: f32 = arguments.next().ok_or("missing movie width")?.parse()?;
    let height: f32 = arguments.next().ok_or("missing movie height")?.parse()?;
    let duration_seconds: f64 = arguments.next().ok_or("missing movie duration")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let movie_filename = filename(&movie_path)?;
    let poster_filename = filename(&poster_path)?;
    let movie = fs::read(&movie_path)?;
    let poster = fs::read(&poster_path)?;
    let size = Size { width, height };
    let position = Point {
        x: (SLIDE_WIDTH_POINTS - width) / 2.0,
        y: (SLIDE_HEIGHT_POINTS - height) / 2.0,
    };

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Movie built from typed IWA objects")
        .build()?;
    let created = editor.add_slide_movie(
        0,
        movie_filename,
        &movie,
        poster_filename,
        &poster,
        SlideMovieOptions::new(
            position,
            size,
            Duration::try_from_secs_f64(duration_seconds)?,
        )?,
    )?;
    let mut properties = editor.slide_movie_properties(0, created.drawable_object_id)?;
    properties.accessibility_description = Some(format!("Embedded movie: {movie_filename}"));
    editor.set_slide_movie_properties(0, created.drawable_object_id, properties)?;
    let playback = created
        .playback
        .ok_or("created Keynote movie has no playback settings")?;
    editor.set_slide_movie_playback_settings(
        0,
        created.drawable_object_id,
        playback
            .with_loop_mode(Some(MediaLoopMode::Repeat))
            .with_volume(Some(MediaVolume::new(0.75)?)),
    )?;
    editor.set_slide_movie_title(0, created.drawable_object_id, "Source-built Keynote movie")?;
    editor.set_slide_movie_caption(
        0,
        created.drawable_object_id,
        &format!("Native title and caption for {movie_filename}"),
    )?;
    let labels = editor.slide_movie_title_caption(0, created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Keynote movie {} backed by video {:?} and poster {:?} with labels {labels:?}",
        created.drawable_object_id,
        created.movie_data_identifier,
        created.poster_image_data_identifier,
    );
    Ok(())
}

fn filename(path: &str) -> Result<&str, Box<dyn std::error::Error>> {
    Path::new(path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "media path must end in a UTF-8 file name".into())
}
