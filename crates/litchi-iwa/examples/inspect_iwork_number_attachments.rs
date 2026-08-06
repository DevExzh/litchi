use litchi_iwa::raw::package::IWorkPackage;
use litchi_iwa::text::IWorkTextEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = std::env::args()
        .nth(1)
        .ok_or("usage: inspect_iwork_number_attachments <input.pages|input.numbers|input.key>")?;
    let editor = IWorkTextEditor::from_package(IWorkPackage::open(input)?);
    for storage in editor.storages()? {
        for attachment in editor.text_number_attachments(storage.object_id)? {
            println!(
                "storage={} id={} position={} settings={:?}",
                storage.object_id,
                attachment.id.object_id(),
                attachment.position.utf16_index(),
                attachment.settings,
            );
        }
    }
    Ok(())
}
