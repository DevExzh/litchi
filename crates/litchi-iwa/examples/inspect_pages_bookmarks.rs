use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = std::env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_bookmarks <input.pages>")?;
    let pages = PagesEditor::open(input)?;
    for bookmark in pages.body_bookmarks()? {
        println!(
            "id={} range={}..{} name={:?} visibility={:?}",
            bookmark.id.get(),
            bookmark.range.start().utf16_index(),
            bookmark.range.end().utf16_index(),
            bookmark.settings.name().map(|name| name.as_str()),
            bookmark.settings.visibility(),
        );
    }
    Ok(())
}
