//! Create a Pages document and body-anchored movie without an input package.

use std::env;
use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa_common::media::playback::{MediaLoopMode, MediaVolume};
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_pages::movie::Options as PagesMovieOptions;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments.next().ok_or(
        "usage: create_pages_movie <output.pages> <movie> <poster-image> [duration-seconds]",
    )?;
    let movie_path = arguments.next().ok_or("missing movie path")?;
    let poster_path = arguments.next().ok_or("missing poster-image path")?;
    let duration_seconds = arguments
        .next()
        .map(|value| value.parse::<f64>())
        .transpose()?
        .unwrap_or(8.0);
    if !duration_seconds.is_finite() || duration_seconds <= 0.0 {
        return Err("duration must be finite and greater than zero".into());
    }
    let duration = Duration::try_from_secs_f64(duration_seconds)
        .map_err(|_| "duration is too large for std::time::Duration")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let movie = fs::read(&movie_path)?;
    let poster = fs::read(&poster_path)?;
    let movie_filename = Path::new(&movie_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("movie path has no UTF-8 basename")?;
    let poster_filename = Path::new(&poster_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("poster path has no UTF-8 basename")?;
    let body = "Source-built Pages movie";
    let anchor = body.encode_utf16().count();
    let mut editor = PagesEditor::create_with_text(body)?;
    let created = editor.add_body_movie(
        anchor,
        movie_filename,
        &movie,
        poster_filename,
        &poster,
        PagesMovieOptions::new(
            Point { x: 96.0, y: 144.0 },
            Size {
                width: 320.0,
                height: 180.0,
            },
            duration,
        )?,
    )?;
    let mut properties = editor.body_movie_properties(created.drawable_object_id)?;
    properties.accessibility_description = Some(format!("Embedded movie: {movie_filename}"));
    editor.set_body_movie_properties(created.drawable_object_id, properties)?;
    editor.set_body_movie_playback_settings(
        created.drawable_object_id,
        created
            .playback
            .with_loop_mode(Some(MediaLoopMode::Repeat))
            .with_volume(Some(MediaVolume::new(0.75)?)),
    )?;
    editor.set_body_movie_title(created.drawable_object_id, "Source-built Pages movie")?;
    editor.set_body_movie_caption(
        created.drawable_object_id,
        &format!("Native title and caption for {movie_filename}"),
    )?;
    let labels = editor.body_movie_title_caption(created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "anchor={} drawable={} movie_data={} poster_data={} duration={:?} labels={labels:?}",
        created.anchor_character_index,
        created.drawable_object_id,
        created.movie_data_identifier,
        created.poster_image_data_identifier,
        created.duration
    );
    Ok(())
}
