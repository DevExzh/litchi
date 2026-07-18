//! Print object IDs, payload types, and indexed references for one IWA member.

use std::env;

use litchi_iwa::IWorkPackage;
use litchi_iwa::IWorkThemeArchive;
use litchi_iwa::charts::IWorkChartArchive;
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
use litchi_iwa::protobuf::tss::StylesheetArchive;
use litchi_iwa::protobuf::tst::TableDataList;
use litchi_iwa::protobuf::tst::{TableInfoArchive, TableModelArchive};
use litchi_iwa::protobuf::tswp::{
    BookmarkFieldArchive, CharacterStyleArchive, ColumnStyleArchive, DateTimeSmartFieldArchive,
    DropCapStyleArchive, HighlightArchive, HyperlinkFieldArchive, ListStyleArchive,
    ParagraphStyleArchive, ShapeInfoArchive, ShapeStyleArchive, StorageArchive,
};
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
            if message.type_ == 5_021
                && let Ok(chart) = IWorkChartArchive::decode(message.data.as_slice())
            {
                println!("  chart_drawable={chart:#?}");
            }
            if message.type_ == 12_006
                && let Ok(mediator) = tn::ChartMediatorArchive::decode(message.data.as_slice())
            {
                println!("  numbers_chart_mediator={mediator:#?}");
            }
            if matches!(message.type_, 0 | 8 | 153 | 205 | 212 | 213 | 3_056 | 5_021) {
                let hex = message
                    .data
                    .iter()
                    .map(|byte| format!("{byte:02x}"))
                    .collect::<String>();
                println!("  payload_hex={hex}");
            }
            if message.type_ == 2_032
                && let Ok(hyperlink) = HyperlinkFieldArchive::decode(message.data.as_slice())
            {
                println!("  hyperlink={hyperlink:#?}");
            }
            if message.type_ == 2_034
                && let Ok(date_time) = DateTimeSmartFieldArchive::decode(message.data.as_slice())
            {
                println!("  date_time={date_time:#?}");
            }
            if message.type_ == 2_035
                && let Ok(bookmark) = BookmarkFieldArchive::decode(message.data.as_slice())
            {
                println!("  bookmark={bookmark:#?}");
            }
            if message.type_ == 2_013
                && let Ok(highlight) = HighlightArchive::decode(message.data.as_slice())
            {
                println!("  highlight={highlight:#?}");
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
            if message.type_ == 10
                && let Ok(theme) = IWorkThemeArchive::decode(message.data.as_slice())
            {
                println!(
                    "  keynote_theme_presets=drawing={:?} text={:?} chart={:?} table={:?} application={:?}",
                    theme.extensions.drawing.as_ref().map(|presets| (
                        presets.gradient_fill_presets.len(),
                        presets.image_fill_presets.len(),
                        presets.shadow_presets.len(),
                        presets.line_style_presets.len(),
                        presets.shape_style_presets.len(),
                        presets.textbox_style_presets.len(),
                        presets.image_style_presets.len(),
                        presets.movie_style_presets.len(),
                        presets.drawing_line_style_presets.len(),
                    )),
                    theme.extensions.text.as_ref().map(|presets| (
                        presets.list_style_presets.len(),
                        presets.character_style_presets.len(),
                        presets.paragraph_style_presets.len(),
                        presets.dropcap_style_presets.len(),
                    )),
                    theme
                        .extensions
                        .chart
                        .as_ref()
                        .map(|presets| presets.chart_presets.len()),
                    theme
                        .extensions
                        .table
                        .as_ref()
                        .map(|presets| presets.table_style_presets.len()),
                    theme.extensions.application.as_ref().map(|presets| (
                        presets.caption_style_presets.len(),
                        presets.svg_import_style_presets.len(),
                    )),
                );
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
            if message.type_ == 2
                && let Ok(sheet) = tn::SheetArchive::decode(message.data.as_slice())
            {
                println!("  numbers_sheet={sheet:#?}");
            }
            if message.type_ == 2
                && let Ok(show) = kn::ShowArchive::decode(message.data.as_slice())
            {
                println!("  keynote_show={show:#?}");
            }
            if message.type_ == 601
                && let Ok(state) = litchi_iwa::protobuf::tsa::FunctionBrowserStateArchive::decode(
                    message.data.as_slice(),
                )
            {
                println!("  function_browser_state={state:#?}");
            }
            if message.type_ == 12_009
                && let Ok(theme) = tn::ThemeArchive::decode(message.data.as_slice())
            {
                println!("  numbers_theme={theme:#?}");
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
            if message.type_ == 7
                && let Ok(placeholder) = kn::PlaceholderArchive::decode(message.data.as_slice())
            {
                println!("  keynote_placeholder={placeholder:#?}");
            }
            if message.type_ == 9
                && let Ok(style) = kn::SlideStyleArchive::decode(message.data.as_slice())
            {
                println!("  keynote_slide_style={style:#?}");
            }
            if message.type_ == 15
                && let Ok(note) = kn::NoteArchive::decode(message.data.as_slice())
            {
                println!("  keynote_note={note:#?}");
            }
            if message.type_ == 21
                && let Ok(soundtrack) = kn::Soundtrack::decode(message.data.as_slice())
            {
                println!("  keynote_soundtrack={soundtrack:#?}");
            }
            if message.type_ == 184
                && let Ok(source) = kn::LiveVideoSource::decode(message.data.as_slice())
            {
                println!("  keynote_live_video_source={source:#?}");
            }
            if message.type_ == 185
                && let Ok(collection) =
                    kn::LiveVideoSourceCollection::decode(message.data.as_slice())
            {
                println!("  keynote_live_video_collection={collection:#?}");
            }
            if message.type_ == 401
                && let Ok(stylesheet) = StylesheetArchive::decode(message.data.as_slice())
            {
                println!("  stylesheet={stylesheet:#?}");
            }
            if message.type_ == 2_011
                && let Ok(shape) = ShapeInfoArchive::decode(message.data.as_slice())
            {
                println!("  text_shape={shape:#?}");
            }
            if message.type_ == 2_021
                && let Ok(style) = CharacterStyleArchive::decode(message.data.as_slice())
            {
                println!("  character_style={style:#?}");
            }
            if message.type_ == 2_022
                && let Ok(style) = ParagraphStyleArchive::decode(message.data.as_slice())
            {
                println!("  paragraph_style={style:#?}");
            }
            if message.type_ == 2_023
                && let Ok(style) = ListStyleArchive::decode(message.data.as_slice())
            {
                println!("  list_style={style:#?}");
            }
            if message.type_ == 2_024
                && let Ok(style) = ColumnStyleArchive::decode(message.data.as_slice())
            {
                println!("  column_style={style:#?}");
            }
            if message.type_ == 2_025
                && let Ok(style) = ShapeStyleArchive::decode(message.data.as_slice())
            {
                println!("  shape_style={style:#?}");
            }
            if message.type_ == 10_024
                && let Ok(style) = DropCapStyleArchive::decode(message.data.as_slice())
            {
                println!("  drop_cap_style={style:#?}");
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
            if message.type_ == 4_008
                && let Ok(owner) =
                    litchi_iwa::protobuf::tsce::FormulaOwnerDependenciesArchive::decode(
                        message.data.as_slice(),
                    )
            {
                println!("  formula_owner_dependencies={owner:#?}");
            }
            if matches!(message.type_, 6_000 | 6_001)
                && let Ok(model) = TableModelArchive::decode(message.data.as_slice())
            {
                println!("  table_model={model:#?}");
            }
            if message.type_ == 6_000
                && let Ok(info) = TableInfoArchive::decode(message.data.as_slice())
            {
                println!("  table_info={info:#?}");
            }
            if message.type_ == 6_220
                && let Ok(filter) =
                    litchi_iwa::protobuf::tst::FilterSetArchive::decode(message.data.as_slice())
            {
                println!("  filter_set={filter:#?}");
            }
            if message.type_ == 6_373
                && let Ok(group) =
                    litchi_iwa::protobuf::tst::GroupByArchive::decode(message.data.as_slice())
            {
                println!("  group_by={group:#?}");
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
                    "  package_last_object_identifier={} save_token={:?} revision={:?} package_type={:?} versions={:?}/{:?}/{:?}",
                    metadata.last_object_identifier,
                    metadata.save_token,
                    metadata.revision,
                    metadata.preferred_package_type,
                    metadata.read_version,
                    metadata.write_version,
                    metadata.file_format_version,
                );
                for component in metadata.components {
                    println!(
                        "  component={} locator={:?} versions={:?}/{:?}/{:?} uuid_objects={:?} external_refs={:?}",
                        component.identifier,
                        component
                            .locator
                            .as_deref()
                            .unwrap_or(&component.preferred_locator),
                        component.document_read_version,
                        component.document_write_version,
                        component.component_read_version,
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
