use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let input = args.next().ok_or(
        "usage: set_keynote_slide_layout <input.key> <output.key> \
         <slide-index> <layout-id|exact-layout-name>",
    )?;
    let output = args.next().ok_or("missing output path")?;
    let slide_index = args.next().ok_or("missing slide index")?.parse()?;
    let requested = args.next().ok_or("missing layout identifier or name")?;
    if args.next().is_some() {
        return Err("too many arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let layouts = editor.slide_layouts()?;
    let matches = layouts
        .iter()
        .filter(|layout| {
            requested
                .parse::<u64>()
                .map_or_else(|_| layout.name == requested, |id| layout.id.as_u64() == id)
        })
        .collect::<Vec<_>>();
    let [layout] = matches.as_slice() else {
        return Err(format!(
            "layout {requested:?} matched {} theme layouts; expected exactly one",
            matches.len()
        )
        .into());
    };
    editor.set_slide_layout(slide_index, layout.id)?;
    editor.save(output)?;
    Ok(())
}
