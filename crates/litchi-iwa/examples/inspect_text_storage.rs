use std::env;

use litchi_iwa::raw::package::IWorkPackage;
use litchi_iwa_protos::tswp::StorageArchive;
use prost::Message;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: inspect_text_storage <file>")?;
    let package = IWorkPackage::open(path)?;
    for name in package.entry_names().filter(|name| name.ends_with(".iwa")) {
        let archive = package.archive(name)?;
        for object in archive.objects {
            let id = object.archive_info.identifier.unwrap_or_default();
            for message in object.messages {
                if message.type_ != 2001 && message.type_ != 2022 {
                    continue;
                }
                let Ok(storage) = StorageArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                if storage.text.is_empty() {
                    continue;
                }
                println!(
                    "id={id} type={} kind={:?} archive={name} text={:?}",
                    message.type_, storage.kind, storage.text
                );
            }
        }
    }
    Ok(())
}
