//! Inspect drawable ownership and protobuf payloads for Pages text boxes.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::raw::package::IWorkPackage;
use litchi_iwa_protos::tp::{DocumentArchive, DrawablesZOrderArchive, FloatingDrawablesArchive};
use litchi_iwa_protos::tswp::{DrawableAttachmentArchive, ShapeInfoArchive, StorageArchive};
use prost::Message;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_text_boxes <input.pages>")?;
    let package = IWorkPackage::open(&input)?;
    let editor = PagesEditor::from_package(package.clone())?;
    let document_archive = package.archive("Index/Document.iwa")?;
    let root = document_archive.object(1).ok_or("missing Pages root")?;
    let document = root
        .messages
        .iter()
        .find(|message| message.type_ == 10000)
        .ok_or("missing Pages document payload")?;
    let document = DocumentArchive::decode(document.data.as_slice())?;
    println!(
        "body={:?} floating={:?} zorder={:?}",
        document.body_storage.map(|reference| reference.identifier),
        document
            .floating_drawables
            .map(|reference| reference.identifier),
        document
            .drawables_zorder
            .map(|reference| reference.identifier)
    );
    if let Some(reference) = document.body_storage {
        print_decoded::<StorageArchive>(&package, reference.identifier, 2001, "body")?;
    }
    if let Some(reference) = document.floating_drawables {
        print_decoded::<FloatingDrawablesArchive>(
            &package,
            reference.identifier,
            10010,
            "floating",
        )?;
    }
    if let Some(reference) = document.drawables_zorder {
        print_decoded::<DrawablesZOrderArchive>(&package, reference.identifier, 10015, "zorder")?;
    }
    for text in editor.drawable_text_storages()? {
        let geometry = editor.text_box_geometry(text.drawable_object_id).ok();
        let properties = editor.text_box_properties(text.drawable_object_id).ok();
        let columns = editor.text_box_columns(text.drawable_object_id).ok();
        let text_layout = editor.text_box_text_layout(text.drawable_object_id).ok();
        println!(
            "drawable={} storage={} text={:?} geometry={geometry:?} properties={properties:?} columns={columns:?} text_layout={text_layout:?}",
            text.drawable_object_id,
            text.storage.id,
            text.storage.storage.text(),
        );
        for name in package.iwa_entry_names() {
            let archive = package.archive(name)?;
            let Some(object) = archive.object(text.drawable_object_id) else {
                continue;
            };
            for message in &object.messages {
                if message.type_ == 2011 {
                    println!(
                        "  member={name} shape={:#?}",
                        ShapeInfoArchive::decode(message.data.as_slice())?
                    );
                }
            }
        }
    }
    Ok(())
}

fn print_decoded<T: Message + Default + std::fmt::Debug>(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
    label: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        let message = object
            .messages
            .iter()
            .find(|message| message.type_ == message_type)
            .ok_or("object has no expected message type")?;
        println!(
            "  member={name} {label}={:#?}",
            T::decode(message.data.as_slice())?
        );
        for referenced in &object.archive_info.message_infos[0].object_references {
            for candidate in package.iwa_entry_names() {
                let referenced_archive = package.archive(candidate)?;
                let Some(referenced_object) = referenced_archive.object(*referenced) else {
                    continue;
                };
                if let Some(attachment) = referenced_object
                    .messages
                    .iter()
                    .find(|message| message.type_ == 2003)
                {
                    println!(
                        "    attachment={} value={:#?}",
                        referenced,
                        DrawableAttachmentArchive::decode(attachment.data.as_slice())?
                    );
                }
            }
        }
    }
    Ok(())
}
