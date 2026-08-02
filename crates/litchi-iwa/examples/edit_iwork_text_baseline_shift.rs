//! Set one uniform typed baseline shift in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextBaselineShift};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_iwork_text_baseline_shift <input> <output> <storage-id> <points>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let shift = TextBaselineShift::from_points(
        arguments
            .next()
            .ok_or("missing signed baseline shift in points")?
            .parse()?,
    )?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_baseline_shift(storage_id, shift)?;
    editor.save(output)?;
    Ok(())
}
