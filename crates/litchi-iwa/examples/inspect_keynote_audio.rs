//! List independently positioned audio clips and playback builds on Keynote slides.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_keynote_audio <input.key>")?;
    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        let builds = editor.slide_builds(slide.index)?;
        for audio in editor.slide_audio(slide.index)? {
            let audio_builds = builds
                .iter()
                .filter(|build| build.drawable_object_id == audio.drawable_object_id)
                .collect::<Vec<_>>();
            println!("{audio:?} builds={audio_builds:?}");
        }
    }
    Ok(())
}
