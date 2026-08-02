//! List standalone movie drawables owned by Keynote slides.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: inspect_keynote_movies <input.key>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        let builds = editor.slide_builds(slide.index)?;
        for movie in editor.slide_movies(slide.index)? {
            let movie_builds = builds
                .iter()
                .filter(|build| build.drawable_object_id == movie.drawable_object_id)
                .collect::<Vec<_>>();
            println!(
                "slide={} drawable={} kind={:?} video={:?} poster={:?} geometry={:?} builds={movie_builds:?}",
                slide.index,
                movie.drawable_object_id,
                movie.kind,
                movie.movie_data_identifier,
                movie.poster_image_data_identifier,
                movie.geometry,
            );
        }
    }
    Ok(())
}
