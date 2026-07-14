use std::collections::HashMap;
use std::env;

use litchi_iwa::IWorkPackage;
use litchi_iwa::archive::ArchiveObject;
use litchi_iwa::protobuf::kn::{
    BuildArchive, BuildChunkArchive, DocumentArchive, PlaceholderArchive, ShowArchive,
    SlideArchive, SlideNodeArchive,
};
use litchi_iwa::protobuf::tswp::{ShapeInfoArchive, StorageArchive};
use prost::Message;

#[allow(deprecated)]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: inspect_keynote_structure <file>")?;
    let package = IWorkPackage::open(path)?;
    let mut objects: HashMap<u64, (String, ArchiveObject)> = HashMap::new();
    for name in package.entry_names().filter(|name| name.ends_with(".iwa")) {
        for object in package.archive(name)?.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            objects.insert(identifier, (name.to_owned(), object));
        }
    }
    let (_, root) = objects.get(&1).ok_or("document object 1 is missing")?;
    let document = decode::<DocumentArchive>(root).ok_or("no document payload")?;
    println!("document -> show {}", document.show.identifier);
    let (_, show_object) = objects
        .get(&document.show.identifier)
        .ok_or("show object is missing")?;
    let show = decode::<ShowArchive>(show_object).ok_or("no show payload")?;
    println!(
        "slide tree refs: {:?} size={}x{} slide_numbers={:?} loop={:?} mode={:?} autoplay=({:?},{:?},{:?}) idle=({:?},{:?}) slide_list={:?} recording={:?}",
        show.slide_tree
            .slides
            .iter()
            .map(|r| r.identifier)
            .collect::<Vec<_>>(),
        show.size.width,
        show.size.height,
        show.slide_numbers_visible,
        show.loop_presentation,
        show.mode,
        show.autoplay_transition_delay,
        show.autoplay_build_delay,
        show.automatically_plays_upon_open,
        show.idle_timer_active,
        show.idle_timer_delay,
        show.slide_list.as_ref().map(|r| r.identifier),
        show.recording.as_ref().map(|r| r.identifier)
    );
    for reference in show.slide_tree.slides {
        let (name, object) = objects
            .get(&reference.identifier)
            .ok_or("slide node object is missing")?;
        println!(
            "ref={} archive={} types={:?}",
            reference.identifier,
            name,
            object.messages.iter().map(|m| m.type_).collect::<Vec<_>>()
        );
        if let Some(node) = decode::<SlideNodeArchive>(object) {
            println!(
                " node -> slide {:?} skipped={} slide_number={:?} unique={:?} copy_from={:?} thumbnails={} dirty={:?} template_id={:?} build_cache=({:?},{:?},{:?},{:?})",
                node.slide.as_ref().map(|r| r.identifier),
                node.is_skipped,
                node.is_slide_number_visible,
                node.unique_identifier,
                node.copy_from_slide_identifier,
                node.thumbnails.len(),
                node.thumbnails_are_dirty,
                node.template_slide_id,
                node.build_event_count,
                node.build_event_count_cache_version,
                node.has_explicit_builds,
                node.has_explicit_builds_cache_version,
            );
            if let Some(slide_ref) = node.slide {
                let (name, object) = objects
                    .get(&slide_ref.identifier)
                    .ok_or("slide object is missing")?;
                let slide = decode::<SlideArchive>(object).ok_or("no slide payload")?;
                let note_id = slide.note.as_ref().map(|reference| reference.identifier);
                println!(
                    " slide={} archive={} name={:?} owned={:?} title={:?} body={:?} note={:?}",
                    slide_ref.identifier,
                    name,
                    slide.name,
                    slide
                        .owned_drawables
                        .iter()
                        .map(|r| r.identifier)
                        .collect::<Vec<_>>(),
                    slide.title_placeholder.as_ref().map(|r| r.identifier),
                    slide.body_placeholder.as_ref().map(|r| r.identifier),
                    note_id,
                );
                let transition = &slide.transition.attributes;
                println!(
                    "  transition modern={:?} legacy=({:?},{:?},{:?},{:?},{:?})",
                    transition.animation_attributes,
                    transition.database_animation_type,
                    transition.database_effect,
                    transition.database_duration,
                    transition.database_delay,
                    transition.database_is_automatic,
                );
                for build_reference in &slide.builds {
                    let (_, build_object) = objects
                        .get(&build_reference.identifier)
                        .ok_or("build object is missing")?;
                    let build = decode::<BuildArchive>(build_object).ok_or("no build payload")?;
                    let chunks = slide
                        .build_chunks
                        .iter()
                        .filter_map(|reference| {
                            let (_, object) = objects.get(&reference.identifier)?;
                            let chunk = decode::<BuildChunkArchive>(object)?;
                            (chunk.build.as_ref().map(|build| build.identifier)
                                == Some(build_reference.identifier))
                            .then_some((reference.identifier, chunk))
                        })
                        .collect::<Vec<_>>();
                    println!(
                        "  build={} drawable={:?} delivery={:?} attributes={:?} chunks={:?}",
                        build_reference.identifier,
                        build
                            .drawable
                            .as_ref()
                            .map(|reference| reference.identifier),
                        build.delivery,
                        build.attributes,
                        chunks,
                    );
                }
                for drawable in slide.owned_drawables {
                    let (name, object) = objects
                        .get(&drawable.identifier)
                        .ok_or("drawable object is missing")?;
                    let storage = decode::<ShapeInfoArchive>(object)
                        .and_then(|shape| shape.owned_storage)
                        .or_else(|| {
                            decode::<PlaceholderArchive>(object)
                                .and_then(|placeholder| placeholder.super_.owned_storage)
                        })
                        .map(|reference| reference.identifier);
                    println!(
                        "  drawable={} archive={} types={:?} storage={:?}",
                        drawable.identifier,
                        name,
                        object.messages.iter().map(|m| m.type_).collect::<Vec<_>>(),
                        storage
                    );
                    if let Some(storage_id) = storage {
                        let (name, object) = objects
                            .get(&storage_id)
                            .ok_or("text storage object is missing")?;
                        println!(
                            "   storage={} archive={} types={:?} text={:?}",
                            storage_id,
                            name,
                            object.messages.iter().map(|m| m.type_).collect::<Vec<_>>(),
                            decode::<StorageArchive>(object).map(|storage| storage.text)
                        );
                    }
                }
                if let Some(note_id) = note_id {
                    let (name, object) = objects.get(&note_id).ok_or("note object is missing")?;
                    let note = decode::<litchi_iwa::protobuf::kn::NoteArchive>(object)
                        .ok_or("no note payload")?;
                    let storage_id = note.contained_storage.identifier;
                    let (storage_name, storage_object) = objects
                        .get(&storage_id)
                        .ok_or("note storage object is missing")?;
                    println!(
                        "  note={} archive={} storage={} storage_archive={} text={:?}",
                        note_id,
                        name,
                        storage_id,
                        storage_name,
                        decode::<StorageArchive>(storage_object).map(|storage| storage.text)
                    );
                }
            }
        }
    }
    Ok(())
}

fn decode<T: Message + Default>(object: &ArchiveObject) -> Option<T> {
    object
        .messages
        .iter()
        .find_map(|message| T::decode(message.data.as_slice()).ok())
}
