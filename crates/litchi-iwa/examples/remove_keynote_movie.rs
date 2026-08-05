//! Remove an ordinary file-backed Keynote slide movie and its private graph.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_keynote::slide::media::MovieKind;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: remove_keynote_movie <input.key> <output.key> <slide-index> <movie-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let movie_index: usize = arguments.next().ok_or("missing movie index")?.parse()?;
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
    let removed = editor.remove_slide_movie(slide_index, source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} culled_data={:?}",
        removed.movie.drawable_object_id, removed.removed_data_identifiers,
    );
    Ok(())
}
