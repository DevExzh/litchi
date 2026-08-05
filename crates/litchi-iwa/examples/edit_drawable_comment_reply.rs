use std::env;

use litchi_iwa::IWorkDrawableCommentEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_drawable_comment_reply <input> <output> <drawable-id> <add|set|remove> [reply-id] [text]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let drawable_id = arguments
        .next()
        .ok_or("missing drawable object identifier")?
        .parse::<u64>()?;
    let operation = arguments.next().ok_or("missing add, set, or remove")?;

    let mut editor = IWorkDrawableCommentEditor::open(&input)?;
    match operation.as_str() {
        "add" => {
            let text = arguments.next().ok_or("missing reply text")?;
            if arguments.next().is_some() {
                return Err("unexpected extra arguments".into());
            }
            let reply_id = editor.add_reply(drawable_id, text)?;
            println!("added reply={reply_id}");
        },
        "set" => {
            let reply_id = arguments
                .next()
                .ok_or("missing reply object identifier")?
                .parse::<u64>()?;
            let text = arguments.next().ok_or("missing replacement text")?;
            if arguments.next().is_some() {
                return Err("unexpected extra arguments".into());
            }
            let current_id = editor.set_reply(drawable_id, reply_id, text)?;
            println!("updated reply={reply_id} current_reply={current_id}");
        },
        "remove" => {
            let reply_id = arguments
                .next()
                .ok_or("missing reply object identifier")?
                .parse::<u64>()?;
            if arguments.next().is_some() {
                return Err("unexpected extra arguments".into());
            }
            editor.remove_reply(drawable_id, reply_id)?;
            println!("removed reply={reply_id}");
        },
        _ => return Err("operation must be add, set, or remove".into()),
    }
    editor.save(output)?;
    for reply in editor.replies(drawable_id)? {
        println!(
            "reply={} author={:?} text={:?}",
            reply.storage_id, reply.comment.author_id, reply.comment.text
        );
    }
    Ok(())
}
