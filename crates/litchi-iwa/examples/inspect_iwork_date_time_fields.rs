use litchi_iwa::IWorkPackage;
use litchi_iwa::text::IWorkTextEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = std::env::args()
        .nth(1)
        .ok_or("usage: inspect_iwork_date_time_fields <input.pages|input.numbers|input.key>")?;
    let editor = IWorkTextEditor::from_package(IWorkPackage::open(input)?);
    for storage in editor.storages()? {
        for field in editor.text_date_time_fields(storage.object_id)? {
            println!(
                "storage={} id={} range={}..{} settings={:?}",
                storage.object_id,
                field.id.object_id(),
                field.range.start().utf16_index(),
                field.range.end().utf16_index(),
                field.settings,
            );
        }
    }
    Ok(())
}
