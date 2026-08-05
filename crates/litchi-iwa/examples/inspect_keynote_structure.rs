use std::collections::HashMap;
use std::env;

use litchi_iwa::IWorkPackage;
use litchi_iwa::archive::ArchiveObject;
use litchi_iwa::keynote::{Acceleration, KeynoteEditor, KeynoteShowMode, TextDelivery};
use litchi_iwa::protobuf::kn::{
    BuildArchive, BuildChunkArchive, DocumentArchive, PlaceholderArchive, ShowArchive,
    SlideArchive, SlideNodeArchive, Soundtrack, ThemeArchive,
};
use litchi_iwa::protobuf::tsd::{ImageArchive, MovieArchive};
use litchi_iwa::protobuf::tsp::PackageMetadata;
use litchi_iwa::protobuf::tswp::{ShapeInfoArchive, StorageArchive};
use prost::Message;

const STORAGELESS_PLACEHOLDER_STORAGE_ID: u64 = 0;

const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

#[allow(deprecated)]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: inspect_keynote_structure <file>")?;
    let editor = KeynoteEditor::open(&path)?;
    for slide in editor.slides()? {
        println!(
            "slide {} current layout: {:?}",
            slide.index + 1,
            slide.layout
        );
    }
    println!("soundtrack settings: {:?}", editor.soundtrack_settings()?);
    println!("soundtrack items: {:?}", editor.soundtrack_items()?);
    let media_assets = editor.media_assets()?;
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
    let package_metadata = objects
        .values()
        .find_map(|(_, object)| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        })
        .map(|message| PackageMetadata::decode(message.data.as_slice()))
        .transpose()?
        .ok_or("package metadata payload is missing")?;
    for component in package_metadata
        .components
        .iter()
        .filter(|component| component.preferred_locator == "Slide")
    {
        println!(
            "component {} locator={:?}/{:?} uuids={:?} external={:?} data={:?} ambiguous={:?}",
            component.identifier,
            component.preferred_locator,
            component.locator,
            component
                .object_uuid_map_entries
                .iter()
                .map(|entry| entry.identifier)
                .collect::<Vec<_>>(),
            component.external_references,
            component.data_references,
            component.ambiguous_object_identifiers,
        );
    }
    let (_, root) = objects.get(&1).ok_or("document object 1 is missing")?;
    let document = decode::<DocumentArchive>(root).ok_or("no document payload")?;
    println!("document -> show {}", document.show.identifier);
    let (_, show_object) = objects
        .get(&document.show.identifier)
        .ok_or("show object is missing")?;
    let show = decode::<ShowArchive>(show_object).ok_or("no show payload")?;
    println!(
        "slide tree refs: {:?} size={}x{} slide_numbers={:?} loop={:?} mode={:?} autoplay=({:?},{:?},{:?}) idle=({:?},{:?}) slide_list={:?} recording={:?} soundtrack={:?}",
        show.slide_tree
            .slides
            .iter()
            .map(|r| r.identifier)
            .collect::<Vec<_>>(),
        show.size.width,
        show.size.height,
        show.slide_numbers_visible,
        show.loop_presentation,
        show.mode.map(KeynoteShowMode::from_raw),
        show.autoplay_transition_delay,
        show.autoplay_build_delay,
        show.automatically_plays_upon_open,
        show.idle_timer_active,
        show.idle_timer_delay,
        show.slide_list.as_ref().map(|r| r.identifier),
        show.recording.as_ref().map(|r| r.identifier),
        show.soundtrack.as_ref().map(|r| r.identifier),
    );
    let (_, theme_object) = objects
        .get(&show.theme.identifier)
        .ok_or("theme object is missing")?;
    let theme = decode::<ThemeArchive>(theme_object).ok_or("no theme payload")?;
    for layout_reference in &theme.templates {
        let (_, node_object) = objects
            .get(&layout_reference.identifier)
            .ok_or("layout node is missing")?;
        let node = decode::<SlideNodeArchive>(node_object).ok_or("no layout node payload")?;
        let slide_id = node
            .slide
            .as_ref()
            .ok_or("layout node has no slide")?
            .identifier;
        let (_, slide_object) = objects.get(&slide_id).ok_or("layout slide is missing")?;
        let slide = decode::<SlideArchive>(slide_object).ok_or("no layout slide payload")?;
        println!(
            "layout node={} slide={} name={:?} uuid={:?} style={} placeholders=({:?},{:?},{:?},{:?}) owned={:?}",
            layout_reference.identifier,
            slide_id,
            slide.name,
            node.template_slide_id,
            slide.style.identifier,
            slide.title_placeholder.as_ref().map(|r| r.identifier),
            slide.body_placeholder.as_ref().map(|r| r.identifier),
            slide.object_placeholder.as_ref().map(|r| r.identifier),
            slide
                .slide_number_placeholder
                .as_ref()
                .map(|r| r.identifier),
            slide
                .owned_drawables
                .iter()
                .map(|r| r.identifier)
                .collect::<Vec<_>>(),
        );
        for drawable in &slide.owned_drawables {
            let (archive_name, object) = objects
                .get(&drawable.identifier)
                .ok_or("layout drawable is missing")?;
            if let Some(image) = decode::<ImageArchive>(object) {
                println!(
                    " layout image={} archive={} parent={:?} style={:?} mask={:?} data={:?}/{:?} flags={:?} metadata={:?}",
                    drawable.identifier,
                    archive_name,
                    image.super_.parent.map(|reference| reference.identifier),
                    image.style.map(|reference| reference.identifier),
                    image.mask.map(|reference| reference.identifier),
                    image.data.map(|reference| reference.identifier),
                    image.thumbnail_data.map(|reference| reference.identifier),
                    image.flags,
                    object.archive_info.message_infos,
                );
            }
            if let Some(movie) = decode::<MovieArchive>(object) {
                print_movie(
                    " layout movie",
                    drawable.identifier,
                    archive_name,
                    object,
                    movie,
                );
            }
        }
    }
    if let Some(reference) = &show.soundtrack {
        let (archive_name, object) = objects
            .get(&reference.identifier)
            .ok_or("soundtrack object is missing")?;
        let soundtrack = decode::<Soundtrack>(object).ok_or("no soundtrack payload")?;
        println!(
            "soundtrack {} archive={} info={:?}: {soundtrack:?}",
            reference.identifier, archive_name, object.archive_info.message_infos
        );
        for media in &soundtrack.movie_media {
            println!(
                " soundtrack media {}: {:?}",
                media.identifier,
                media_assets
                    .iter()
                    .find(|asset| asset.data_identifier == media.identifier)
            );
        }
        for component in package_metadata
            .components
            .iter()
            .chain(&package_metadata.versioned_components)
            .filter(|component| {
                component.data_references.iter().any(|data| {
                    soundtrack
                        .movie_media
                        .iter()
                        .any(|media| media.identifier == data.data_identifier)
                })
            })
        {
            println!(
                " soundtrack component {} locator={:?}/{:?} data={:?}",
                component.identifier,
                component.preferred_locator,
                component.locator,
                component.data_references
            );
        }
    }
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
                    " slide={} archive={} name={:?} owned={:?} title={:?} body={:?} object={:?} slide_number={:?} note={:?}",
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
                    slide.object_placeholder.as_ref().map(|r| r.identifier),
                    slide
                        .slide_number_placeholder
                        .as_ref()
                        .map(|r| r.identifier),
                    note_id,
                );
                println!(
                    "  layout style={} template={:?} z_order={:?} layer_with_template={:?} \
                     title=(geometry={:?},shape_style={:?},text_style={:?},layout={:?}) \
                     body=(geometry={:?},shape_style={:?},text_style={:?},layout={:?})",
                    slide.style.identifier,
                    slide.template_slide.as_ref().map(|r| r.identifier),
                    slide
                        .drawables_z_order
                        .iter()
                        .map(|r| r.identifier)
                        .collect::<Vec<_>>(),
                    slide.slide_objects_layer_with_template,
                    slide.title_placeholder_geometry,
                    slide.title_placeholder_shape_style_index,
                    slide.title_placeholder_text_style_index,
                    slide.title_layout_properties,
                    slide.body_placeholder_geometry,
                    slide.body_placeholder_shape_style_index,
                    slide.body_placeholder_text_style_index,
                    slide.body_layout_properties,
                );
                println!("  layout metadata={:?}", object.archive_info.message_infos);
                let transition = &slide.transition.attributes;
                println!(
                    "  transition modern={:?} custom=(timing={:?},delivery={:?},mosaic={:?}) legacy=({:?},{:?},{:?},{:?},{:?})",
                    transition.animation_attributes,
                    transition
                        .custom_timing_curve
                        .map(Acceleration::from_native),
                    transition
                        .custom_text_delivery_type
                        .map(TextDelivery::from_native),
                    transition.custom_mosaic_type,
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
                let retained_placeholders = [
                    slide.title_placeholder.as_ref(),
                    slide.body_placeholder.as_ref(),
                ]
                .into_iter()
                .flatten()
                .filter(|placeholder| {
                    !slide
                        .owned_drawables
                        .iter()
                        .any(|owned| owned.identifier == placeholder.identifier)
                })
                .cloned()
                .collect::<Vec<_>>();
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
                        .map(|reference| reference.identifier)
                        .filter(|identifier| *identifier != STORAGELESS_PLACEHOLDER_STORAGE_ID);
                    println!(
                        "  drawable={} archive={} types={:?} storage={:?}",
                        drawable.identifier,
                        name,
                        object.messages.iter().map(|m| m.type_).collect::<Vec<_>>(),
                        storage
                    );
                    if let Some(placeholder) = decode::<PlaceholderArchive>(object) {
                        println!(
                            "   placeholder kind={:?} geometry={:?} style={:?} metadata={:?}",
                            placeholder.kind,
                            placeholder.super_.super_.super_.geometry,
                            placeholder.super_.super_.style,
                            object.archive_info.message_infos,
                        );
                    }
                    if let Some(image) = decode::<ImageArchive>(object) {
                        println!(
                            "   image parent={:?} style={:?} mask={:?} data={:?}/{:?} flags={:?} metadata={:?}",
                            image.super_.parent.map(|reference| reference.identifier),
                            image.style.map(|reference| reference.identifier),
                            image.mask.map(|reference| reference.identifier),
                            image.data.map(|reference| reference.identifier),
                            image.thumbnail_data.map(|reference| reference.identifier),
                            image.flags,
                            object.archive_info.message_infos,
                        );
                    }
                    if let Some(movie) = decode::<MovieArchive>(object) {
                        print_movie("   movie", drawable.identifier, name, object, movie);
                    }
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
                for placeholder in retained_placeholders {
                    let (_, object) = objects
                        .get(&placeholder.identifier)
                        .ok_or("retained placeholder object is missing")?;
                    let placeholder_archive = decode::<PlaceholderArchive>(object)
                        .ok_or("retained placeholder has no placeholder payload")?;
                    println!(
                        "  retained placeholder={} kind={:?} geometry={:?} style={:?} metadata={:?}",
                        placeholder.identifier,
                        placeholder_archive.kind,
                        placeholder_archive.super_.super_.super_.geometry,
                        placeholder_archive.super_.super_.style,
                        object.archive_info.message_infos,
                    );
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

fn print_movie(
    prefix: &str,
    identifier: u64,
    archive_name: &str,
    object: &ArchiveObject,
    movie: MovieArchive,
) {
    println!(
        "{prefix}={identifier} archive={archive_name} parent={:?} style={:?} data={:?} poster={:?} database={:?}/{:?} flags={:?} live={:?} size={:?}/{:?} metadata={:?}",
        movie.super_.parent.map(|reference| reference.identifier),
        movie.style.map(|reference| reference.identifier),
        movie.movie_data.map(|reference| reference.identifier),
        movie
            .poster_image_data
            .map(|reference| reference.identifier),
        movie
            .database_movie_data
            .map(|reference| reference.identifier),
        movie
            .database_poster_image_data
            .map(|reference| reference.identifier),
        movie.flags,
        movie.is_live_video,
        movie.original_size,
        movie.natural_size,
        object.archive_info.message_infos,
    );
}
