use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_pages <output.pages> [body text]")?;
    let body_text = std::env::args()
        .nth(2)
        .unwrap_or_else(|| "Created from scratch with litchi-iwa".to_owned());

    PagesEditor::create_with_text(body_text)?.save(output)?;
    Ok(())
}
