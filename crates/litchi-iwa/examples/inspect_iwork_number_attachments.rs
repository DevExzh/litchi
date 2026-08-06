use litchi_iwa::text::IWorkTextEditor;
use litchi_iwa_text::number_attachment::raw::object_id as native_object_id;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = std::env::args()
        .nth(1)
        .ok_or("usage: inspect_iwork_number_attachments <input.pages|input.numbers|input.key>")?;
    let editor = IWorkTextEditor::open(input)?;
    for storage in editor.storages()? {
        for attachment in editor.text_number_attachments(storage.object_id)? {
            println!(
                "storage={} id={} position={} settings={:?}",
                storage.object_id,
                native_object_id(attachment.id),
                attachment.position.utf16_index(),
                attachment.settings,
            );
        }
    }
    Ok(())
}
