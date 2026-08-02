use std::env;
use std::path::Path;

use litchi_iwa::EmbeddedMediaAsset;
use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_reachable_media <iwork-file>")?;
    match Path::new(&input)
        .extension()
        .and_then(|extension| extension.to_str())
    {
        Some("key") => {
            let editor = KeynoteEditor::open(&input)?;
            print_assets("presentation", &editor.media_assets()?);
            for slide in editor.slides()? {
                print_assets(
                    &format!("slide[{}] id={}", slide.index, slide.slide_id),
                    &editor.slide_media_assets(slide.index)?,
                );
            }
        },
        Some("numbers") => {
            let editor = NumbersEditor::open(&input)?;
            print_assets("spreadsheet", &editor.media_assets()?);
            for sheet in editor.sheets()? {
                print_assets(
                    &format!("sheet[{}] id={}", sheet.index, sheet.object_id),
                    &editor.sheet_media_assets(sheet.object_id)?,
                );
            }
        },
        Some("pages") => {
            let editor = PagesEditor::open(&input)?;
            print_assets("document", &editor.media_assets()?);
            for section in editor.sections() {
                print_assets(
                    &format!("section id={}", section.object_id),
                    &editor.section_media_assets(section.object_id)?,
                );
            }
        },
        extension => return Err(format!("unsupported iWork extension: {extension:?}").into()),
    }
    Ok(())
}

fn print_assets(label: &str, assets: &[EmbeddedMediaAsset]) {
    let identifiers = assets
        .iter()
        .map(|asset| asset.data_identifier)
        .collect::<Vec<_>>();
    println!("{label}: {identifiers:?}");
}
