use std::env;

use litchi_iwa::IWorkDrawableCommentEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_drawable_comments <input>")?;
    let editor = IWorkDrawableCommentEditor::open(input)?;
    println!("application={:?}", editor.application());
    for drawable in editor.drawables()? {
        let comment = editor.comment(drawable.id)?;
        println!(
            "drawable={} type={} storage={:?} comment={:?}",
            drawable.id,
            drawable.message_type,
            drawable.comment_id,
            comment.as_ref().map(|value| &value.comment)
        );
        for reply in editor.replies(drawable.id)? {
            println!(
                "  reply={} root={} comment={:?}",
                reply.storage_id, reply.root_storage_id, reply.comment
            );
        }
    }
    Ok(())
}
