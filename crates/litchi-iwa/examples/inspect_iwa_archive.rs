//! Print object IDs, payload types, and indexed references for one IWA member.

use std::env;

use litchi_iwa::IWorkPackage;
use litchi_iwa::protobuf::kn;
use litchi_iwa::protobuf::tn;
use litchi_iwa::protobuf::tp::{
    DocumentArchive, SectionArchive, SectionTemplateArchive, UserDefinedGuideMapArchive,
};
use litchi_iwa::protobuf::tsce::CalculationEngineArchive;
use litchi_iwa::protobuf::tsd::CommentStorageArchive;
use litchi_iwa::protobuf::tsd::GuideStorageArchive;
use litchi_iwa::protobuf::tsk::{AnnotationAuthorArchive, AnnotationAuthorStorageArchive};
use litchi_iwa::protobuf::tsp::PackageMetadata;
use litchi_iwa::protobuf::tst::TableDataList;
use litchi_iwa::protobuf::tswp::{ShapeInfoArchive, StorageArchive};
use prost::Message;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: inspect_iwa_archive <input> [Index/member.iwa]")?;
    let member = arguments.next();
    let object_id = arguments
        .next()
        .map(|value| value.parse::<u64>())
        .transpose()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let package = IWorkPackage::open(input)?;
    let members = member.map_or_else(
        || package.iwa_entry_names().map(str::to_owned).collect(),
        |member| vec![member],
    );
    for member in members {
        let archive = package.archive(&member)?;
        println!("member={member}");
        print_archive(archive, object_id);
    }
    Ok(())
}

fn print_archive(archive: litchi_iwa::archive::Archive, object_id: Option<u64>) {
    for object in archive.objects {
        let identifier = object.archive_info.identifier.unwrap_or_default();
        if object_id.is_some_and(|expected| expected != identifier) {
            continue;
        }
        println!(
            "object={identifier} merge={:?} types={:?}",
            object.archive_info.should_merge,
            object
                .messages
                .iter()
                .map(|message| message.type_)
                .collect::<Vec<_>>()
        );
        for (index, info) in object.archive_info.message_infos.iter().enumerate() {
            println!(
                "  message={index} versions={:?} length={} refs={:?} data_refs={:?} fields={}",
                info.versions,
                info.length,
                info.object_references,
                info.data_references,
                info.field_infos.len()
            );
            for field in &info.field_infos {
                println!("    field={field:#?}");
            }
        }
        for message in &object.messages {
            if matches!(message.type_, 0 | 8 | 153 | 212 | 213 | 3_056) {
                let hex = message
                    .data
                    .iter()
                    .map(|byte| format!("{byte:02x}"))
                    .collect::<String>();
                println!("  payload_hex={hex}");
            }
            if message.type_ == 8
                && let Ok(build) = kn::BuildArchive::decode(message.data.as_slice())
            {
                println!("  keynote_build={build:#?}");
            }
            if message.type_ == 10
                && let Ok(theme) = kn::ThemeArchive::decode(message.data.as_slice())
            {
                println!("  keynote_theme={theme:#?}");
            }
            if message.type_ == 2_043
                && let Ok(attachment) =
                    kn::SlideNumberAttachmentArchive::decode(message.data.as_slice())
            {
                println!("  keynote_slide_number_attachment={attachment:#?}");
            }
            if message.type_ == 153
                && let Ok(chunk) = kn::BuildChunkArchive::decode(message.data.as_slice())
            {
                println!("  keynote_build_chunk={chunk:#?}");
            }
            if message.type_ == 212
                && let Ok(author) = AnnotationAuthorArchive::decode(message.data.as_slice())
            {
                println!("  annotation_author={author:#?}");
            }
            if message.type_ == 213
                && let Ok(storage) = AnnotationAuthorStorageArchive::decode(message.data.as_slice())
            {
                println!("  annotation_author_storage={storage:#?}");
            }
            if message.type_ == 3_056
                && let Ok(comment) = CommentStorageArchive::decode(message.data.as_slice())
            {
                println!("  comment_storage={comment:#?}");
            }
            if message.type_ == 6_005
                && let Ok(list) = TableDataList::decode(message.data.as_slice())
            {
                println!("  table_data_list={list:#?}");
            }
            if message.type_ == 1
                && let Ok(document) = kn::DocumentArchive::decode(message.data.as_slice())
            {
                println!("  keynote_document={document:#?}");
            }
            if message.type_ == 1
                && let Ok(document) = tn::DocumentArchive::decode(message.data.as_slice())
            {
                println!("  numbers_document={document:#?}");
            }
            if message.type_ == 4
                && let Ok(node) = kn::SlideNodeArchive::decode(message.data.as_slice())
            {
                println!("  keynote_slide_node={node:#?}");
            }
            if message.type_ == 5
                && let Ok(slide) = kn::SlideArchive::decode(message.data.as_slice())
            {
                println!("  keynote_slide={slide:#?}");
            }
            if message.type_ == 2_011
                && let Ok(shape) = ShapeInfoArchive::decode(message.data.as_slice())
            {
                println!("  text_shape={shape:#?}");
            }
            if matches!(message.type_, 2_001 | 2_022)
                && let Ok(storage) = StorageArchive::decode(message.data.as_slice())
            {
                println!("  text_storage={storage:#?}");
            }
            if message.type_ == 4_000
                && let Ok(engine) = CalculationEngineArchive::decode(message.data.as_slice())
            {
                println!("  calculation_engine={engine:#?}");
            }
            if message.type_ == 10_000
                && let Ok(document) = DocumentArchive::decode(message.data.as_slice())
            {
                println!("  pages_document={document:#?}");
            }
            if message.type_ == 10_011
                && let Ok(section) = SectionArchive::decode(message.data.as_slice())
            {
                println!("  pages_section={section:#?}");
            }
            if message.type_ == 10_143
                && let Ok(template) = SectionTemplateArchive::decode(message.data.as_slice())
            {
                println!("  pages_section_template={template:#?}");
            }
            if message.type_ == 10_016
                && let Ok(guides) = UserDefinedGuideMapArchive::decode(message.data.as_slice())
            {
                println!("  pages_guide_map={guides:#?}");
            }
            if message.type_ == 3_047
                && let Ok(guides) = GuideStorageArchive::decode(message.data.as_slice())
            {
                println!("  guide_storage={guides:#?}");
            }
            if message.type_ == 11_006
                && let Ok(metadata) = PackageMetadata::decode(message.data.as_slice())
            {
                println!(
                    "  package_last_object_identifier={} save_token={:?} revision={:?}",
                    metadata.last_object_identifier, metadata.save_token, metadata.revision
                );
                for component in metadata.components {
                    println!(
                        "  component={} locator={:?} uuid_objects={:?} external_refs={:?}",
                        component.identifier,
                        component
                            .locator
                            .as_deref()
                            .unwrap_or(&component.preferred_locator),
                        component
                            .object_uuid_map_entries
                            .iter()
                            .map(|entry| (
                                entry.identifier,
                                format!("{:016x}{:016x}", entry.uuid.upper, entry.uuid.lower),
                            ))
                            .collect::<Vec<_>>(),
                        component
                            .external_references
                            .iter()
                            .map(|reference| (
                                reference.component_identifier,
                                reference.object_identifier,
                                reference.is_weak
                            ))
                            .collect::<Vec<_>>()
                    );
                }
            }
        }
    }
}
