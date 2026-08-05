use std::env;

use litchi_iwa::IWorkDrawableCommentEditor;
use litchi_iwa::comments::DrawableObjectId;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_drawable_comment <input> <output> <drawable-id> <text|--clear>")?;
    let output = arguments
        .next()
        .ok_or("usage: edit_drawable_comment <input> <output> <drawable-id> <text|--clear>")?;
    let drawable_id = arguments
        .next()
        .ok_or("missing drawable object identifier")?
        .parse::<u64>()?;
    let drawable_id =
        DrawableObjectId::new(drawable_id).ok_or("drawable object identifier must be non-zero")?;
    let replacement = arguments
        .next()
        .ok_or("missing replacement text or --clear")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkDrawableCommentEditor::open(&input)?;
    let old = editor.comment(drawable_id)?;
    if replacement == "--clear" {
        editor.clear_comment(drawable_id)?;
    } else {
        editor.set_comment(drawable_id, replacement)?;
    }
    editor.save(&output)?;
    let new = editor.comment(drawable_id)?;
    println!(
        "application={:?} drawable={} old={:?} new={:?}",
        editor.application(),
        drawable_id,
        old.as_ref().map(|value| value.comment.text.as_str()),
        new.as_ref().map(|value| value.comment.text.as_str())
    );
    Ok(())
}
