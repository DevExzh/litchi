use std::{
    env,
    error::Error,
    fs,
    path::{Path, PathBuf},
};

use sha2::{Digest, Sha256};

#[path = "src/production_codec_guard.rs"]
mod production_codec_guard;

use production_codec_guard::{
    FORBIDDEN_BUFFA_OWNERSHIP_MARKERS, FORBIDDEN_PROST_CODEC_MARKERS, production_codec_source,
};

fn main() -> Result<(), Box<dyn Error>> {
    const PROTO_DIRECTORY: &str = "src/protos";
    const BUFFA_PROJECTION_DIRECTORY: &str = "src/buffa-projections";
    let proto_directory = Path::new(PROTO_DIRECTORY);
    let buffa_projection_directory = Path::new(BUFFA_PROJECTION_DIRECTORY);

    println!("cargo:rerun-if-changed={PROTO_DIRECTORY}");
    println!("cargo:rerun-if-changed={BUFFA_PROJECTION_DIRECTORY}");
    println!("cargo:rerun-if-changed=src/archive_codec.rs");
    println!("cargo:rerun-if-changed=src/group_node_category_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_document_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_chart_caption_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_chart_title_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_show_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_placeholder_text_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_speaker_notes_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_slide_number_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_soundtrack_settings_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_media_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_slide_transition_codec.rs");
    println!("cargo:rerun-if-changed=src/hyperlink_codec.rs");
    println!("cargo:rerun-if-changed=src/comment_storage_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_names_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_sheet_order_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_table_header_settings_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_table_title_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_table_cell_storage_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_table_cell_dependency_codec.rs");
    // Keep the native Numbers and Pages message-ID routes tied to their schema
    // projections when this crate is built from the workspace. Published
    // standalone copies do not contain these sibling sources, so the
    // provenance checks below are intentionally conditional on their presence.
    for path in [
        "../litchi-iwa/src/protobuf.rs",
        "../litchi-numbers/src/package/extractor.rs",
        "../litchi-iwa/src/pages/editor.rs",
        "../litchi-iwa/src/pages/editor/movies/graph.rs",
        "../litchi-iwa/src/pages/editor/movies/caption.rs",
        "../litchi-iwa/src/pages/editor/audio/graph.rs",
        "../litchi-iwa/src/pages/editor/body_shapes/caption.rs",
        "../litchi-iwa/src/image_caption.rs",
        "../litchi-iwa/src/pages/editor/footnotes.rs",
        "../litchi-iwa/src/pages/creation.rs",
        "../litchi-pages/src/package.rs",
        "../litchi-pages/src/package/footnote_text.rs",
        "../litchi-pages/src/package/document_settings.rs",
        "../litchi-pages/src/package/page_layout.rs",
        "../litchi-pages/src/package/section_background.rs",
        "../litchi-pages/src/package/section_pagination.rs",
        "../litchi-pages/src/package/section_settings.rs",
        "../litchi-pages/src/package/section_text.rs",
        "../litchi-pages/src/package/section_transaction.rs",
        "../litchi-pages/src/package/table_lock.rs",
    ] {
        if Path::new(path).is_file() {
            println!("cargo:rerun-if-changed={path}");
        }
    }
    println!("cargo:rerun-if-changed=src/package_metadata_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_formula_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_body_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_media_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_movie_caption_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_footnote_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_footnote_marker_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_section_background_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_document_settings_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_page_layout_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_section_codec.rs");
    println!("cargo:rerun-if-changed=src/production_codec_guard.rs");
    println!("cargo:rerun-if-changed=src/table_info_codec.rs");
    println!("cargo:rerun-if-changed=src/text_storage_codec.rs");
    println!("cargo:rerun-if-changed=src/lib.rs");

    let mut proto_files = fs::read_dir(proto_directory)?
        .map(|directory_entry| directory_entry.map(|entry| entry.path()))
        .collect::<Result<Vec<PathBuf>, _>>()?;
    proto_files.retain(|path| {
        path.extension()
            .is_some_and(|extension| extension == "proto")
    });
    proto_files.sort_unstable();

    if proto_files.is_empty() {
        return Err(format!("no Protocol Buffer schemas found in {PROTO_DIRECTORY}").into());
    }
    enforce_projection_schema_ratchets(buffa_projection_directory)?;
    enforce_production_ingress_ratchets()?;
    enforce_text_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_group_node_category_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_keynote_document_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_keynote_chart_caption_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_keynote_chart_title_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_keynote_show_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_keynote_placeholder_text_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_keynote_speaker_notes_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_keynote_slide_number_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_keynote_soundtrack_settings_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_keynote_slide_transition_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_comment_storage_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_numbers_names_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_numbers_sheet_order_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_table_header_settings_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_table_title_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_table_cell_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_numbers_table_data_list_provenance(proto_directory, buffa_projection_directory)?;
    enforce_package_metadata_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_formula_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_pages_native_message_provenance(proto_directory)?;
    enforce_pages_body_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_pages_media_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_pages_movie_caption_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_pages_footnote_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_pages_footnote_marker_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_pages_section_background_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_pages_section_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_table_info_projection_provenance(proto_directory, buffa_projection_directory)?;

    prost_build::Config::new()
        .include_file("iwa_protos.rs")
        .compile_protos(&proto_files, &[proto_directory])?;

    // Keep the archive-header sidecar isolated from format projections. Prost
    // remains the full-corpus compatibility generator during migration.
    let buffa_proto_files = [
        proto_directory.join("TSPMessages.proto"),
        proto_directory.join("TSPArchiveMessages.proto"),
    ];
    let buffa_out_directory = PathBuf::from(env::var("OUT_DIR")?).join("buffa");
    buffa_build::Config::new()
        .files(&buffa_proto_files)
        .includes(&[proto_directory])
        .out_dir(&buffa_out_directory)
        .include_file("iwa_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(true)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_full_buffa_projection_budget(&buffa_out_directory)?;

    // The text decoder never encodes or preserves from its view: caller-owned
    // source bytes remain authoritative. Generate the tiny derived projection
    // separately with unknown retention disabled so unrelated native fields
    // consume neither generated closure nor unknown-span storage.
    let buffa_text_out_directory = PathBuf::from(env::var("OUT_DIR")?).join("buffa-text-storage");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSWPStorageArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_text_out_directory)
        .include_file("iwa_text_storage_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;

    enforce_text_projection_budget(&buffa_text_out_directory)?;

    // Numbers comment storage needs only the optional text/date/author/UUID
    // facts.  The repeated replies field remains on the strict handwritten
    // source router so no generated input-width vector can be materialized.
    let buffa_comment_storage_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-numbers-comment-storage");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSDCommentStorageArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_comment_storage_out_directory)
        .include_file("iwa_comment_storage_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_comment_storage_projection_budget(&buffa_comment_storage_out_directory)?;

    // Group-by category labels need only a zero-field GroupNode envelope plus
    // UUID and four scalar wrappers. The streaming adapter routes recursive
    // children and CellValue branches without a generated repeated-field
    // vector. Keep this format-specific read-only projection separate from the
    // full TST/TSCE schema closure.
    let buffa_group_node_category_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-group-node-category");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSTGroupNodeCategoryArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_group_node_category_out_directory)
        .include_file("iwa_group_node_category_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_group_node_category_projection_budget(&buffa_group_node_category_out_directory)?;

    // Keynote consumes only the show reference from its root document. Keep
    // the TSA/TSK base archive opaque so opening a presentation cannot
    // materialize unrelated generated metadata through this projection.
    let buffa_keynote_document_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-document");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("KNDocumentArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_document_out_directory)
        .include_file("iwa_keynote_document_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_document_projection_budget(&buffa_keynote_document_out_directory)?;

    // Keynote chart caption reads need only the optional drawable super,
    // optional caption reference, and required nested identifier. Keep the
    // chart extension closure out of generated code.
    let buffa_keynote_chart_caption_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-chart-caption");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSCHChartCaptionArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_chart_caption_out_directory)
        .include_file("iwa_keynote_chart_caption_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_chart_caption_projection_budget(&buffa_keynote_chart_caption_out_directory)?;

    // Keynote chart-title reads need only the two scalar fields from the
    // generated ChartNonStyleArchive extension. Keep the outer non-style
    // envelope and every unrelated generated field caller-owned.
    let buffa_keynote_chart_title_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-chart-title");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSCHChartTitleArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_chart_title_out_directory)
        .include_file("iwa_keynote_chart_title_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_chart_title_projection_budget(&buffa_keynote_chart_title_out_directory)?;

    // Numbers reaches a table model through field 2 of TableInfo. Keep the
    // drawable base archive and all display metadata out of generated code;
    // the format adapter owns strict source validation and raw preservation.
    let buffa_table_info_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-table-info");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSTTableInfoArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_table_info_out_directory)
        .include_file("iwa_table_info_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_table_info_projection_budget(&buffa_table_info_out_directory)?;

    // Numbers name reads need only the direct sheet name, the form-sheet
    // inheritance envelope, and the table model's identity/display strings.
    // Keep repeated drawable and model metadata outside generated code; the
    // strict borrowed codec owns all traversal and resource limits.
    let buffa_numbers_names_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-numbers-names");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TNNumbersNamesArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_numbers_names_out_directory)
        .include_file("iwa_numbers_names_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_numbers_names_projection_budget(&buffa_numbers_names_out_directory)?;

    let buffa_numbers_sheet_order_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-numbers-sheet-order");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TNNumbersSheetReferenceArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_numbers_sheet_order_out_directory)
        .include_file("iwa_numbers_sheet_order_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_numbers_sheet_order_projection_budget(&buffa_numbers_sheet_order_out_directory)?;

    // Numbers table-header settings require only dimensions and nine scalar
    // header/footer/freeze/repetition facts. Keep required style/data-store
    // references and all repeated table content outside generated code.
    let buffa_table_header_settings_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-numbers-table-header-settings");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSTTableHeaderSettingsArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_table_header_settings_out_directory)
        .include_file("iwa_numbers_table_header_settings_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_table_header_settings_projection_budget(&buffa_table_header_settings_out_directory)?;

    let buffa_table_title_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-numbers-table-title");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSTTableTitleSettingsArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_table_title_out_directory)
        .include_file("iwa_numbers_table_title_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_table_title_projection_budget(&buffa_table_title_out_directory)?;

    // Scalar-cell edits traverse several collection-heavy native archives.
    // Project only singular envelopes and scalars; the generated-free codec
    // streams repeated tiles, rows, headers, strings, and dependency records.
    let buffa_table_cell_storage_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-numbers-table-cell-storage");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSTTableCellStorageArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_table_cell_storage_out_directory)
        .include_file("iwa_numbers_table_cell_storage_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_table_cell_storage_projection_budget(&buffa_table_cell_storage_out_directory)?;

    let buffa_table_cell_dependency_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-numbers-table-cell-dependency");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSCETableCellDependenciesArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_table_cell_dependency_out_directory)
        .include_file("iwa_numbers_table_cell_dependency_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_table_cell_dependency_projection_budget(&buffa_table_cell_dependency_out_directory)?;

    // PackageMetadata publication edits only one scalar plus selected nested
    // registry records. Repeated components and records stay on the strict
    // handwritten streaming path to keep the generated closure width-free.
    let buffa_package_metadata_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-package-metadata");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSPPackageMetadataArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_package_metadata_out_directory)
        .include_file("iwa_package_metadata_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_package_metadata_projection_budget(&buffa_package_metadata_out_directory)?;

    let buffa_formula_out_directory = PathBuf::from(env::var("OUT_DIR")?).join("buffa-formula");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSCEFormulaArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_formula_out_directory)
        .include_file("iwa_formula_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_formula_projection_budget(&buffa_formula_out_directory)?;

    // Keynote's show reader projects only scalar settings, required direct
    // references, and presentation size. The repeated slide tree is routed by
    // a bounded handwritten iterator so generated code never owns an
    // input-width reference vector.
    let buffa_keynote_show_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-show");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("KNShowArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_show_out_directory)
        .include_file("iwa_keynote_show_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_show_projection_budget(&buffa_keynote_show_out_directory)?;

    // A semantic Keynote title/body edge ends at one placeholder's owned text
    // storage through three required inheritance envelopes. Generate only
    // that singular chain; source records remain the preservation authority.
    let buffa_keynote_placeholder_text_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-placeholder-text");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("KNPlaceholderTextOwnerArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_placeholder_text_out_directory)
        .include_file("iwa_keynote_placeholder_text_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_placeholder_text_projection_budget(
        &buffa_keynote_placeholder_text_out_directory,
    )?;

    // Focused semantic slide ownership needs the note/title/body edges, the
    // slide's scalar selector fields, and the required transition envelope.
    // Unknown content remains byte-authoritative in caller-owned IWA.
    let buffa_keynote_speaker_notes_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-speaker-notes");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("KNSpeakerNotesArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_speaker_notes_out_directory)
        .include_file("iwa_keynote_speaker_notes_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_speaker_notes_projection_budget(&buffa_keynote_speaker_notes_out_directory)?;

    // Slide numbers need one visibility bit plus a small scalar storage and
    // textual-attachment chain.  The repeated attachment table is raw bytes
    // here and receives one bounded handwritten strict pass in the codec.
    let buffa_keynote_slide_number_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-slide-number");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("KNSlideNumberArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_slide_number_out_directory)
        .include_file("iwa_keynote_slide_number_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_slide_number_projection_budget(&buffa_keynote_slide_number_out_directory)?;

    let buffa_keynote_soundtrack_settings_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-soundtrack-settings");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("KNSoundtrackSettingsArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_soundtrack_settings_out_directory)
        .include_file("iwa_keynote_soundtrack_settings_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_soundtrack_settings_projection_budget(
        &buffa_keynote_soundtrack_settings_out_directory,
    )?;

    // Keynote slide transitions use only a small nested scalar path.  The
    // source archive remains authoritative for preservation; Buffa supplies a
    // borrowed semantic cross-check after strict wire preflight.
    let buffa_keynote_slide_transition_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-keynote-slide-transition");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("KNSlideTransitionArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_keynote_slide_transition_out_directory)
        .include_file("iwa_keynote_slide_transition_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_keynote_slide_transition_projection_budget(
        &buffa_keynote_slide_transition_out_directory,
    )?;

    // Pages section pagination is three optional scalar values. Keep all
    // template, name, and fill data outside generated code and decode the
    // selected values through a borrowed lazy view.
    let buffa_pages_section_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-pages-section");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TPSectionArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_pages_section_out_directory)
        .include_file("iwa_pages_section_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_pages_section_projection_budget(&buffa_pages_section_out_directory)?;

    // Pages section backgrounds need only field 30 and the nested solid-color
    // discriminants/components. Strict handwritten routing remains the raw
    // preservation and publication authority.
    let buffa_pages_section_background_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-pages-section-background");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TPSectionBackgroundArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_pages_section_background_out_directory)
        .include_file("iwa_pages_section_background_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_pages_section_background_projection_budget(
        &buffa_pages_section_background_out_directory,
    )?;

    // Pages root/body traversal needs only three root references, scalar page
    // layout/document settings, and one streamed section-boundary entry. The
    // enclosing section table stays out of generated code, and strict
    // preflight owns every ingress limit.
    let buffa_pages_body_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-pages-body");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TPDocumentBodyArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_pages_body_out_directory)
        .include_file("iwa_pages_body_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_pages_body_projection_budget(&buffa_pages_body_out_directory)?;

    // Pages media discovery needs only the audio-only discriminator from the
    // shared MovieArchive. Keep the complete media graph and every unrelated
    // field on the caller-owned raw/prost compatibility path.
    let buffa_pages_media_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-pages-media");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSDMovieAudioFlagArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_pages_media_out_directory)
        .include_file("iwa_pages_media_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_pages_media_projection_budget(&buffa_pages_media_out_directory)?;

    // Pages movie captions need only the bounded caption-info inheritance
    // chain, its private graph references, and the native kind/text-box
    // scalars. Strict raw routing remains the source-preservation authority.
    let buffa_pages_movie_caption_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-pages-movie-caption");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TPMovieCaptionArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_pages_movie_caption_out_directory)
        .include_file("iwa_pages_movie_caption_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_pages_movie_caption_projection_budget(&buffa_pages_movie_caption_out_directory)?;

    // Pages footnote references need only the textual-attachment envelope,
    // contained-storage reference, and custom marker. The strict handwritten
    // codec validates these selected fields before forcing this private lazy
    // view; unknown source bytes remain outside generated code.
    let buffa_pages_footnote_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-pages-footnote");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSWPFootnoteReferenceAttachmentArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_pages_footnote_out_directory)
        .include_file("iwa_pages_footnote_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_pages_footnote_projection_budget(&buffa_pages_footnote_out_directory)?;

    // Pages footnote markers need only the two scalar fields from one
    // TextualAttachmentArchive. Keep this projection separate from the
    // footnote-reference closure so marker validation cannot materialize
    // references or any unrelated TSWP archive fields.
    let buffa_pages_footnote_marker_out_directory =
        PathBuf::from(env::var("OUT_DIR")?).join("buffa-pages-footnote-marker");
    buffa_build::Config::new()
        .files(&[buffa_projection_directory.join("TSWPTextualAttachmentArchive.proto")])
        .includes(&[buffa_projection_directory])
        .out_dir(&buffa_pages_footnote_marker_out_directory)
        .include_file("iwa_pages_footnote_marker_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(false)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;
    enforce_pages_footnote_marker_projection_budget(&buffa_pages_footnote_marker_out_directory)?;

    Ok(())
}

fn enforce_projection_schema_ratchets(projection_directory: &Path) -> Result<(), Box<dyn Error>> {
    // Every derived schema is compiled in isolation.  Keep the complete
    // source inventory, byte width, and digest explicit so a new field,
    // imported closure, or unreviewed projection file cannot silently widen a
    // lazy ingress boundary while remaining under a per-generator budget.
    const EXPECTED_PROJECTIONS: &[(&str, usize, &str)] = &[
        (
            "KNDocumentArchive.proto",
            844,
            "d4dba9f6a73a35531e9c8bb9731504891000d415bb981b74f783710998630236",
        ),
        (
            "KNPlaceholderTextOwnerArchive.proto",
            1108,
            "2f076952a2f963ab9fa410f2625f3eac7f5ee1f26f5b1c49f833266016f13d8c",
        ),
        (
            "KNShowArchive.proto",
            1682,
            "06641b66f5bc7c137578a29ff16e43167a6e307fe13159b07538a99e49c96d56",
        ),
        (
            "KNSlideNumberArchive.proto",
            767,
            "0467ffa978c763fb10bf9de2097eaad354eedef61714ee4707a2e1355bc4b607",
        ),
        (
            "KNSlideTransitionArchive.proto",
            2347,
            "7a74d790563b72453833a73fa7352b0a796ba11a8b87f9c7665897b9ce3a28e0",
        ),
        (
            "KNSoundtrackSettingsArchive.proto",
            338,
            "7c64e558e49c485272e1c878aa93571e516a4a7730e6de945f72998829e4f2d2",
        ),
        (
            "KNSpeakerNotesArchive.proto",
            1402,
            "20500304a527d8c0531148217c3302b36eb2cee7a48d7a90533130f9a70acd77",
        ),
        (
            "TNNumbersNamesArchive.proto",
            910,
            "348c2e554240f1f2800fb283bb62e025897ed5c45c720550844a83d796126e93",
        ),
        (
            "TNNumbersSheetReferenceArchive.proto",
            280,
            "5ea19e0ad657c4367b1974d0730d1ea0a75602d4e38d2da4dfeb67b1ac753436",
        ),
        (
            "TPDocumentBodyArchive.proto",
            2600,
            "3461ea3c1165a3fd82fba1aebbcd239a604dd56e8ea1052014e59900a3db0d7b",
        ),
        (
            "TPMovieCaptionArchive.proto",
            1055,
            "c1c8f7131f4362794811e0ce396769a37c61175d10befb63886b4a76109c6e7c",
        ),
        (
            "TPSectionArchive.proto",
            1653,
            "4f284cd2403ade092ae8e8105924e15c34f69e203307fd0aabc90e5d705041d7",
        ),
        (
            "TPSectionBackgroundArchive.proto",
            783,
            "0a6f03a7046c285e431953b8752096a1f0117206724b561da294c64092aa9cfc",
        ),
        (
            "TSCEFormulaArchive.proto",
            2343,
            "3c477f4610fedd8fc563ffd83122984d3042cfa8b9d756e3e1c60a719e8d5ba8",
        ),
        (
            "TSCETableCellDependenciesArchive.proto",
            3257,
            "b1ec4313ee2c0012f0441829de19662567457ea9b445dc557fc6354c7abfe533",
        ),
        (
            "TSCHChartCaptionArchive.proto",
            546,
            "caa27f6d9eaab23e7eb33d48744c205cf38629765be97b7f135ceb9edab04d12",
        ),
        (
            "TSCHChartTitleArchive.proto",
            526,
            "8faaf5b5fd30a2a73e34e73aa003cff591c4f2aa3e32dfca58cc1c5fce09f10a",
        ),
        (
            "TSDCommentStorageArchive.proto",
            1173,
            "396d98fd78f6a417a57af4a1e7f3830362e3174753687aef2fe49aaf7a88087d",
        ),
        (
            "TSDMovieAudioFlagArchive.proto",
            251,
            "e4306fa9440c13f2f77587f2edfb25c81eab639087d217bf88f5738ce416c2a9",
        ),
        (
            "TSPPackageMetadataArchive.proto",
            858,
            "f33fc54b7382231d9b8ece89390928cf634bd9a8108e5a929a30563e2d693a60",
        ),
        (
            "TSTGroupNodeCategoryArchive.proto",
            1196,
            "3f2a9a2d2f53d6cd9f6496899f356f53cb507122c71da8dfc82acf57b5735f40",
        ),
        (
            "TSTTableCellStorageArchive.proto",
            3641,
            "17d1cd1afd6f59c46d29f2c481744ead27568ffe937a2b9d9633ce376cb754c9",
        ),
        (
            "TSTTableHeaderSettingsArchive.proto",
            930,
            "1236d9a9d0116885c7140683e5de2d33b6a083435bf3d3cfbebf91172c856d24",
        ),
        (
            "TSTTableInfoArchive.proto",
            1010,
            "93d7d29b24f2e279e5d62a900142890e99417ce0caa098cbf65d3dbb47088c3c",
        ),
        (
            "TSTTableTitleSettingsArchive.proto",
            746,
            "66e04d6d4049bd2bdaa79da79f52e01431cd1579890c903c16ac1764e1476715",
        ),
        (
            "TSWPStorageArchive.proto",
            587,
            "54be1aea50f7e6a211ccb2e19b4abbf9b7ab9c9748a99941dad3e68b7cfa37ba",
        ),
        (
            "TSWPFootnoteReferenceAttachmentArchive.proto",
            954,
            "6a8b19d679e9cb331764f537b9e342943f1e08860784284ad176016b4fcddde4",
        ),
        (
            "TSWPTextualAttachmentArchive.proto",
            470,
            "12cd2d1186d8c0241c6439085c5d0b911fde6ecd5dec759626a2d94dceca3c15",
        ),
    ];

    let mut actual_names = fs::read_dir(projection_directory)?
        .map(|entry| {
            let entry = entry?;
            if !entry.file_type()?.is_file() {
                return Ok(None);
            }
            let name = entry
                .file_name()
                .into_string()
                .map_err(|_| "projection filename is not UTF-8")?;
            if !name.ends_with(".proto") {
                return Err(format!("unexpected non-proto projection file {name}").into());
            }
            Ok(Some(name))
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?
        .into_iter()
        .flatten()
        .collect::<Vec<_>>();
    actual_names.sort_unstable();
    let mut expected_names = EXPECTED_PROJECTIONS
        .iter()
        .map(|(name, _length, _digest)| (*name).to_owned())
        .collect::<Vec<_>>();
    expected_names.sort_unstable();
    if actual_names != expected_names {
        return Err(format!(
            "derived projection inventory drifted: found {actual_names:?}, expected {expected_names:?}"
        )
        .into());
    }

    for (name, expected_length, expected_digest) in EXPECTED_PROJECTIONS {
        let path = projection_directory.join(name);
        let source = fs::read(&path)?;
        if source.len() != *expected_length {
            return Err(format!(
                "derived projection {name} changed size to {} bytes; expected {expected_length}",
                source.len()
            )
            .into());
        }
        let digest = Sha256::digest(&source)
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect::<String>();
        if digest != *expected_digest {
            return Err(format!(
                "derived projection {name} digest {digest} does not match reviewed digest {expected_digest}"
            )
            .into());
        }
        if source
            .split(|byte| *byte == b'\n')
            .map(|line| line.iter().copied().skip_while(u8::is_ascii_whitespace))
            .any(|line| line.collect::<Vec<_>>().starts_with(b"import "))
        {
            return Err(format!(
                "derived projection {name} introduced an unreviewed imported dependency"
            )
            .into());
        }
    }
    Ok(())
}

fn enforce_production_ingress_ratchets() -> Result<(), Box<dyn Error>> {
    // Keep every production Buffa ingress tied to one private generated
    // module.  Prost remains the compatibility type generator, but no
    // production codec may decode untrusted bytes through a Prost-owned
    // message or expose a generated view boundary to downstream crates.
    const CODECS: &[(&str, &str, &str)] = &[
        (
            "src/archive_codec.rs",
            "buffa_generated::TSP",
            "mod buffa_generated {",
        ),
        (
            "src/text_storage_codec.rs",
            "crate::buffa_text_storage_generated::",
            "mod buffa_text_storage_generated {",
        ),
        (
            "src/comment_storage_codec.rs",
            "crate::buffa_comment_storage_generated::",
            "mod buffa_comment_storage_generated {",
        ),
        (
            "src/group_node_category_codec.rs",
            "crate::buffa_group_node_category_generated::",
            "mod buffa_group_node_category_generated {",
        ),
        (
            "src/keynote_document_codec.rs",
            "crate::buffa_keynote_document_generated::",
            "mod buffa_keynote_document_generated {",
        ),
        (
            "src/keynote_chart_caption_codec.rs",
            "crate::buffa_keynote_chart_caption_generated::",
            "mod buffa_keynote_chart_caption_generated {",
        ),
        (
            "src/keynote_chart_title_codec.rs",
            "crate::buffa_keynote_chart_title_generated::",
            "mod buffa_keynote_chart_title_generated {",
        ),
        (
            "src/keynote_placeholder_text_codec.rs",
            "crate::buffa_keynote_placeholder_text_generated::",
            "mod buffa_keynote_placeholder_text_generated {",
        ),
        (
            "src/keynote_speaker_notes_codec.rs",
            "crate::buffa_keynote_speaker_notes_generated::",
            "mod buffa_keynote_speaker_notes_generated {",
        ),
        (
            "src/keynote_slide_number_codec.rs",
            "crate::buffa_keynote_slide_number_generated::",
            "mod buffa_keynote_slide_number_generated {",
        ),
        (
            "src/keynote_soundtrack_settings_codec.rs",
            "crate::buffa_keynote_soundtrack_settings_generated::",
            "mod buffa_keynote_soundtrack_settings_generated {",
        ),
        (
            "src/keynote_media_codec.rs",
            "crate::buffa_generated::TSP::",
            "mod buffa_generated {",
        ),
        (
            "src/keynote_slide_transition_codec.rs",
            "crate::buffa_keynote_slide_transition_generated::",
            "mod buffa_keynote_slide_transition_generated {",
        ),
        (
            "src/keynote_show_codec.rs",
            "crate::buffa_keynote_show_generated::",
            "mod buffa_keynote_show_generated {",
        ),
        (
            "src/numbers_names_codec.rs",
            "crate::buffa_numbers_names_generated::",
            "mod buffa_numbers_names_generated {",
        ),
        (
            "src/numbers_sheet_order_codec.rs",
            "crate::buffa_numbers_sheet_order_generated::",
            "mod buffa_numbers_sheet_order_generated {",
        ),
        (
            "src/numbers_table_header_settings_codec.rs",
            "crate::buffa_numbers_table_header_settings_generated::",
            "mod buffa_numbers_table_header_settings_generated {",
        ),
        (
            "src/numbers_table_title_codec.rs",
            "crate::buffa_numbers_table_title_generated::",
            "mod buffa_numbers_table_title_generated {",
        ),
        (
            "src/numbers_table_cell_storage_codec.rs",
            "crate::buffa_numbers_table_cell_storage_generated::",
            "mod buffa_numbers_table_cell_storage_generated {",
        ),
        (
            "src/numbers_table_cell_dependency_codec.rs",
            "crate::buffa_numbers_table_cell_dependency_generated::",
            "mod buffa_numbers_table_cell_dependency_generated {",
        ),
        (
            "src/package_metadata_codec.rs",
            "crate::buffa_package_metadata_generated::",
            "mod buffa_package_metadata_generated {",
        ),
        (
            "src/numbers_formula_codec.rs",
            "crate::buffa_formula_generated::",
            "mod buffa_formula_generated {",
        ),
        (
            "src/table_info_codec.rs",
            "crate::buffa_table_info_generated::",
            "mod buffa_table_info_generated {",
        ),
        (
            "src/pages_section_codec.rs",
            "crate::buffa_pages_section_generated::",
            "mod buffa_pages_section_generated {",
        ),
        (
            "src/pages_section_background_codec.rs",
            "crate::buffa_pages_section_background_generated::",
            "mod buffa_pages_section_background_generated {",
        ),
        (
            "src/pages_body_codec.rs",
            "crate::buffa_pages_body_generated::",
            "mod buffa_pages_body_generated {",
        ),
        (
            "src/pages_media_codec.rs",
            "crate::buffa_pages_media_generated::",
            "mod buffa_pages_media_generated {",
        ),
        (
            "src/pages_movie_caption_codec.rs",
            "crate::buffa_pages_movie_caption_generated::",
            "mod buffa_pages_movie_caption_generated {",
        ),
        (
            "src/pages_footnote_codec.rs",
            "crate::buffa_pages_footnote_generated::",
            "mod buffa_pages_footnote_generated {",
        ),
        (
            "src/pages_footnote_marker_codec.rs",
            "crate::buffa_pages_footnote_marker_generated::",
            "mod buffa_pages_footnote_marker_generated {",
        ),
        (
            "src/pages_page_layout_codec.rs",
            "crate::buffa_pages_body_generated::",
            "mod buffa_pages_body_generated {",
        ),
        (
            "src/pages_document_settings_codec.rs",
            "crate::buffa_pages_body_generated::",
            "mod buffa_pages_body_generated {",
        ),
    ];
    // Not every focused codec needs a generated view.  Keep those raw-only
    // paths explicit as well, so adding a new `*_codec.rs` cannot silently
    // bypass the production ingress review.  Raw entries are checked below
    // for the same forbidden Prost/owned-view markers and for accidental
    // Buffa usage.
    const RAW_CODECS: &[&str] = &["src/hyperlink_codec.rs"];

    let mut expected_paths = CODECS
        .iter()
        .map(|(path, _generated_marker, _private_module_marker)| *path)
        .chain(RAW_CODECS.iter().copied())
        .map(str::to_owned)
        .collect::<Vec<_>>();
    expected_paths.sort_unstable();
    let mut actual_paths = fs::read_dir("src")?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .filter(|path| {
            path.is_file()
                && path.extension().is_some_and(|extension| extension == "rs")
                && path
                    .file_stem()
                    .and_then(|stem| stem.to_str())
                    .is_some_and(|stem| stem.ends_with("_codec"))
        })
        .map(|path| path.to_string_lossy().replace('\\', "/"))
        .collect::<Vec<_>>();
    actual_paths.sort_unstable();
    if actual_paths != expected_paths {
        return Err(format!(
            "production codec inventory drifted: found {actual_paths:?}, expected {expected_paths:?}"
        )
        .into());
    }

    let lib = fs::read_to_string("src/lib.rs")?;
    for (path, generated_marker, private_module_marker) in CODECS {
        let source = fs::read_to_string(path)?;
        // Some codecs have cfg(test) allocation probes near their imports;
        // the shared source slicer removes every test-only item without
        // truncating production at the first such probe.
        let production = production_codec_source(&source);
        if !production.contains("decode_lazy_view")
            || !production.contains(generated_marker)
            || !has_exact_private_module_declaration(&lib, private_module_marker)
            || FORBIDDEN_PROST_CODEC_MARKERS
                .iter()
                .any(|fragment| production.contains(fragment))
            || FORBIDDEN_BUFFA_OWNERSHIP_MARKERS
                .iter()
                .any(|fragment| production.contains(fragment))
        {
            return Err(format!(
                "production Buffa ingress ratchet failed for {path}: expected private {generated_marker} lazy decode and no Prost decode"
            )
            .into());
        }
    }

    for path in RAW_CODECS {
        let source = fs::read_to_string(path)?;
        let production = production_codec_source(&source);
        if production.contains("decode_lazy_view")
            || production.contains("buffa::")
            || production.contains("buffa_generated")
            || FORBIDDEN_PROST_CODEC_MARKERS
                .iter()
                .any(|fragment| production.contains(fragment))
            || FORBIDDEN_BUFFA_OWNERSHIP_MARKERS
                .iter()
                .any(|fragment| production.contains(fragment))
        {
            return Err(format!(
                "raw production codec {path} unexpectedly uses Buffa or a forbidden generated ingress marker"
            )
            .into());
        }
    }
    Ok(())
}

/// Return whether `source` contains exactly one private declaration for the
/// generated module and no published declaration for the same module.
///
/// A plain substring search is insufficient here: `pub mod foo {` contains
/// `mod foo {` and would otherwise satisfy the generated-boundary ratchet.
fn has_exact_private_module_declaration(source: &str, declaration: &str) -> bool {
    let private_count = source
        .lines()
        .filter(|line| line.trim() == declaration)
        .count();
    let published_count = source
        .lines()
        .filter(|line| {
            let line = line.trim();
            (line.starts_with("pub ") || line.starts_with("pub(")) && line.ends_with(declaration)
        })
        .count();
    private_count == 1 && published_count == 0
}

/// Count an exact private or crate-visible route declaration.
///
/// Route constants are implementation details. Accept a private declaration
/// or restricted `pub(...)` visibility, but reject a public item; checking the
/// complete trimmed line avoids accepting `const NAME ...` as a substring of
/// `pub const NAME ...`. The lexical count also keeps declaration text in
/// comments and string literals from satisfying the marker.
fn rust_non_public_declaration_count(source: &str, declaration: &str) -> usize {
    let Some(expected) = rust_declaration_item(declaration) else {
        return 0;
    };
    let production = production_codec_source(source);
    let candidates = production
        .lines()
        .filter_map(rust_declaration_item)
        .filter(|line| *line == expected)
        .count();
    if candidates == 1 && rust_code_marker_count(source, expected) == 1 {
        1
    } else {
        0
    }
}

/// Strip a Rust item's visibility while rejecting an unrestricted `pub`
/// declaration.  Restricted forms such as `pub(crate)`, `pub(super)`, and
/// `pub(in crate::pages)` remain implementation details and are accepted by
/// the numeric-route ratchet.
fn rust_declaration_item(declaration: &str) -> Option<&str> {
    let declaration = declaration.trim();
    if declaration.starts_with("pub ") {
        return None;
    }
    if let Some(rest) = declaration.strip_prefix("pub(") {
        let end = rest.find(") ")?;
        return Some(&rest[end + 2..]);
    }
    Some(declaration)
}

/// Count a route marker only when it starts in production Rust code and does
/// not cross a comment. String literals remain searchable because a few route
/// markers intentionally include a literal type name; a marker that starts
/// inside a literal is rejected. `production_codec_source` removes complete
/// `#[cfg(test)]` items before the lexical pass.
fn rust_code_marker_count(source: &str, marker: &str) -> usize {
    if marker.is_empty() {
        return 0;
    }

    let production = production_codec_source(source);
    let source = production.as_ref();
    let bytes = source.as_bytes();
    let mut code = vec![false; bytes.len()];
    let mut marker_start = vec![false; bytes.len()];
    let mut cursor = 0usize;

    while cursor < bytes.len() {
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'/') {
            cursor += 2;
            while cursor < bytes.len() && bytes[cursor] != b'\n' {
                cursor += 1;
            }
            continue;
        }
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
            let mut depth = 1usize;
            cursor += 2;
            while cursor < bytes.len() && depth != 0 {
                if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
                    depth = depth.saturating_add(1);
                    cursor += 2;
                } else if bytes.get(cursor) == Some(&b'*') && bytes.get(cursor + 1) == Some(&b'/') {
                    depth = depth.saturating_sub(1);
                    cursor += 2;
                } else {
                    cursor += 1;
                }
            }
            continue;
        }

        if let Some(end) = rust_route_literal_end(bytes, cursor) {
            for position in cursor..end {
                code[position] = true;
            }
            cursor = end;
            continue;
        }

        code[cursor] = true;
        marker_start[cursor] = true;
        cursor += 1;
    }

    source
        .match_indices(marker)
        .filter(|(start, _)| {
            let Some(end) = start.checked_add(marker.len()) else {
                return false;
            };
            marker_start.get(*start).copied().unwrap_or(false)
                && code
                    .get(*start..end)
                    .is_some_and(|span| span.iter().all(|is_code| *is_code))
        })
        .count()
}

/// Return the end of a Rust string, byte string, raw string, character, or
/// byte-character literal beginning at `cursor`. Lifetimes are left in normal
/// code so a `'name` token cannot hide a route marker.
fn rust_route_literal_end(bytes: &[u8], cursor: usize) -> Option<usize> {
    let (quote, raw_hashes) = if bytes.get(cursor) == Some(&b'r')
        || (bytes.get(cursor) == Some(&b'b') && bytes.get(cursor + 1) == Some(&b'r'))
    {
        let prefix = if bytes[cursor] == b'b' {
            cursor + 2
        } else {
            cursor + 1
        };
        let mut quote = prefix;
        while bytes.get(quote) == Some(&b'#') {
            quote += 1;
        }
        (bytes.get(quote) == Some(&b'"')).then_some((quote, quote - prefix))?
    } else if bytes.get(cursor) == Some(&b'"') {
        (cursor, 0)
    } else if bytes.get(cursor) == Some(&b'b') && bytes.get(cursor + 1) == Some(&b'"') {
        (cursor + 1, 0)
    } else if bytes.get(cursor) == Some(&b'\'') {
        let end = rust_route_char_end(bytes, cursor)?;
        return Some(end);
    } else if bytes.get(cursor) == Some(&b'b') && bytes.get(cursor + 1) == Some(&b'\'') {
        let end = rust_route_char_end(bytes, cursor + 1)?;
        return Some(end);
    } else {
        return None;
    };

    if raw_hashes != 0 || bytes.get(quote) == Some(&b'"') && cursor != quote {
        let mut probe = quote + 1;
        while probe < bytes.len() {
            if bytes[probe] == b'"'
                && bytes
                    .get(probe + 1..probe + 1 + raw_hashes)
                    .is_some_and(|tail| tail.iter().all(|byte| *byte == b'#'))
            {
                return Some(probe + 1 + raw_hashes);
            }
            probe += 1;
        }
        return Some(bytes.len());
    }

    let mut probe = quote + 1;
    while probe < bytes.len() {
        if bytes[probe] == b'\\' {
            probe = probe.saturating_add(2);
        } else if bytes[probe] == b'"' {
            return Some(probe + 1);
        } else {
            probe += 1;
        }
    }
    Some(bytes.len())
}

fn rust_route_char_end(bytes: &[u8], quote: usize) -> Option<usize> {
    let first = *bytes.get(quote + 1)?;
    if (first.is_ascii_alphanumeric() || first == b'_') && bytes.get(quote + 2) != Some(&b'\'') {
        return None;
    }

    let mut probe = quote + 1;
    while probe < bytes.len() {
        if bytes[probe] == b'\n' || bytes[probe] == b'\r' {
            return None;
        }
        if bytes[probe] == b'\\' {
            probe = probe.saturating_add(2);
        } else if bytes[probe] == b'\'' {
            return Some(probe + 1);
        } else {
            probe += 1;
        }
    }
    None
}

fn proto_identifier_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_'
}

fn skip_proto_space(bytes: &[u8], mut index: usize) -> usize {
    loop {
        while bytes
            .get(index)
            .is_some_and(|byte| byte.is_ascii_whitespace())
        {
            index += 1;
        }
        let Some(next) = skip_proto_comment(bytes, index) else {
            break;
        };
        index = next;
    }
    index
}

fn skip_proto_comment(bytes: &[u8], index: usize) -> Option<usize> {
    if bytes.get(index) == Some(&b'/') && bytes.get(index.saturating_add(1)) == Some(&b'/') {
        let mut cursor = index.saturating_add(2);
        while cursor < bytes.len() && bytes[cursor] != b'\n' {
            cursor += 1;
        }
        return Some(cursor);
    }
    if bytes.get(index) == Some(&b'/') && bytes.get(index.saturating_add(1)) == Some(&b'*') {
        let mut cursor = index.saturating_add(2);
        while cursor.saturating_add(1) < bytes.len() {
            if bytes[cursor] == b'*' && bytes[cursor.saturating_add(1)] == b'/' {
                return Some(cursor.saturating_add(2));
            }
            cursor += 1;
        }
        return None;
    }
    None
}

fn skip_proto_string(bytes: &[u8], index: usize) -> Option<usize> {
    let quote = *bytes.get(index)?;
    if quote != b'"' && quote != b'\'' {
        return None;
    }
    let mut cursor = index.saturating_add(1);
    while cursor < bytes.len() {
        if bytes[cursor] == b'\\' {
            cursor = cursor.saturating_add(2);
        } else if bytes[cursor] == quote {
            return Some(cursor.saturating_add(1));
        } else {
            cursor += 1;
        }
    }
    None
}

fn proto_keyword_at(bytes: &[u8], index: usize, keyword: &[u8]) -> bool {
    bytes
        .get(index..index.saturating_add(keyword.len()))
        .is_some_and(|tail| tail == keyword)
        && (index == 0 || !proto_identifier_byte(bytes[index - 1]))
        && bytes
            .get(index.saturating_add(keyword.len()))
            .is_none_or(|byte| !proto_identifier_byte(*byte))
}

fn proto_balanced_end(bytes: &[u8], opening_brace: usize) -> Option<usize> {
    let mut depth = 0usize;
    let mut index = opening_brace;
    while index < bytes.len() {
        if let Some(next) = skip_proto_comment(bytes, index) {
            index = next;
            continue;
        }
        if let Some(next) = skip_proto_string(bytes, index) {
            index = next;
            continue;
        }
        match bytes[index] {
            b'{' => depth = depth.checked_add(1)?,
            b'}' => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return index.checked_add(1);
                }
            },
            _ => {},
        }
        index += 1;
    }
    None
}

/// Return one complete proto message at exactly `target_depth`, including
/// nested declarations. Comments, strings, and messages at another nesting
/// depth are lexical context rather than declarations for this lookup; this
/// prevents a commented or unrelated nested message from satisfying a
/// provenance ratchet.
fn proto_message_block_at_depth<'source>(
    source: &'source str,
    name: &str,
    target_depth: usize,
) -> Option<&'source str> {
    let bytes = source.as_bytes();
    let name_bytes = name.as_bytes();
    let mut depth = 0usize;
    let mut index = 0usize;
    let mut found = None;
    while index < bytes.len() {
        if let Some(next) = skip_proto_comment(bytes, index) {
            index = next;
            continue;
        }
        if let Some(next) = skip_proto_string(bytes, index) {
            index = next;
            continue;
        }
        if depth == target_depth && proto_keyword_at(bytes, index, b"message") {
            let mut cursor = skip_proto_space(bytes, index.saturating_add(b"message".len()));
            if bytes
                .get(cursor..cursor.saturating_add(name_bytes.len()))
                .is_some_and(|candidate| candidate == name_bytes)
                && (cursor == 0 || !proto_identifier_byte(bytes[cursor - 1]))
                && bytes
                    .get(cursor.saturating_add(name_bytes.len()))
                    .is_none_or(|byte| !proto_identifier_byte(*byte))
            {
                cursor = skip_proto_space(bytes, cursor.saturating_add(name_bytes.len()));
                if bytes.get(cursor) == Some(&b'{') {
                    let end = proto_balanced_end(bytes, cursor)?;
                    if found.is_some() {
                        return None;
                    }
                    found = Some((index, end));
                    index = end;
                    continue;
                }
            }
        }
        match bytes[index] {
            b'{' => depth = depth.checked_add(1)?,
            b'}' => depth = depth.checked_sub(1)?,
            _ => {},
        }
        index += 1;
    }
    found.and_then(|(start, end)| source.get(start..end))
}

fn proto_message_block<'source>(source: &'source str, name: &str) -> Option<&'source str> {
    proto_message_block_at_depth(source, name, 0)
}

fn proto_nested_message_block<'source>(source: &'source str, name: &str) -> Option<&'source str> {
    proto_message_block_at_depth(source, name, 1)
}

/// Count one direct field declaration in a message block.
///
/// The search ignores comments and strings and only accepts declarations at
/// the message body's own brace depth. Nested enums/messages and similarly
/// named fields elsewhere therefore cannot satisfy the check.
fn proto_field(message_block: &str, declaration: &str) -> usize {
    let bytes = message_block.as_bytes();
    let declaration_bytes = declaration.as_bytes();
    let mut opening_brace = None;
    let mut index = 0usize;
    while index < bytes.len() {
        if let Some(next) = skip_proto_comment(bytes, index) {
            index = next;
            continue;
        }
        if let Some(next) = skip_proto_string(bytes, index) {
            index = next;
            continue;
        }
        if bytes[index] == b'{' {
            opening_brace = Some(index);
            break;
        }
        index += 1;
    }
    let Some(opening_brace) = opening_brace else {
        return 0;
    };

    let mut nested_depth = 0usize;
    let mut matches = 0usize;
    index = opening_brace.saturating_add(1);
    while index < bytes.len() {
        if let Some(next) = skip_proto_comment(bytes, index) {
            index = next;
            continue;
        }
        if let Some(next) = skip_proto_string(bytes, index) {
            index = next;
            continue;
        }
        match bytes[index] {
            b'{' => {
                nested_depth = match nested_depth.checked_add(1) {
                    Some(depth) => depth,
                    None => return matches,
                };
                index += 1;
                continue;
            },
            b'}' => {
                if nested_depth == 0 {
                    break;
                }
                nested_depth -= 1;
                index += 1;
                continue;
            },
            _ => {},
        }
        if nested_depth == 0
            && bytes
                .get(index..index.saturating_add(declaration_bytes.len()))
                .is_some_and(|candidate| candidate == declaration_bytes)
            && (index == 0 || !proto_identifier_byte(bytes[index - 1]))
            && bytes
                .get(index.saturating_add(declaration_bytes.len()))
                .is_none_or(|byte| !proto_identifier_byte(*byte))
        {
            matches = matches.saturating_add(1);
        }
        index += 1;
    }
    matches
}

fn sha256_hex(source: &str) -> String {
    Sha256::digest(source.as_bytes())
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn enforce_text_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TEXT_DECLARATION: &str = "repeated string text = 3;";
    // Pin the complete canonical StorageArchive declaration, not merely a
    // globally matching field spelling.  A duplicate `text = 3` in another
    // message must never authorize this ingress projection.
    const STORAGE_ARCHIVE_DIGEST: &str =
        "280fa513b7bda90cd2866d50ef250129afc885e7296fa391937dd93e9a143936";

    let canonical = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TSWPStorageArchive.proto"))?;
    let Some(storage_archive) = proto_message_block(&canonical, "StorageArchive") else {
        return Err("TSWP text provenance lost the canonical StorageArchive message".into());
    };
    if proto_field(storage_archive, TEXT_DECLARATION) != 1
        || sha256_hex(storage_archive) != STORAGE_ARCHIVE_DIGEST
        || projection.matches(TEXT_DECLARATION).count() != 1
    {
        return Err(
            "derived TSWP text projection is out of sync with StorageArchive field 3".into(),
        );
    }
    Ok(())
}

fn enforce_comment_storage_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_PROJECTION_DIGEST: &str =
        "396d98fd78f6a417a57af4a1e7f3830362e3174753687aef2fe49aaf7a88087d";
    const TSD_FIELDS: [&str; 5] = [
        "optional string text = 1;",
        "optional .TSP.Date creation_date = 2;",
        "optional .TSP.Reference author = 3;",
        "repeated .TSP.Reference replies = 4;",
        "optional .TSP.UUID storage_uuid = 5;",
    ];
    const TSP_DATE: &str = "message Date {\n  required double seconds = 1;\n}";
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const TSP_UUID: &str =
        "message UUID {\n  required uint64 lower = 1;\n  required uint64 upper = 2;\n}";
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaCommentStorageProjection;\n\
message Date {\n\
required double seconds = 1;\n\
}\n\
message Reference {\n\
required uint64 identifier = 1;\n\
optional int32 deprecated_type = 2;\n\
optional bool deprecated_is_external = 3;\n\
}\n\
message Uuid {\n\
required uint64 lower = 1;\n\
required uint64 upper = 2;\n\
}\n\
message CommentStorageArchive {\n\
optional string text = 1;\n\
optional .LitchiIwaCommentStorageProjection.Date creation_date = 2;\n\
optional .LitchiIwaCommentStorageProjection.Reference author = 3;\n\
optional .LitchiIwaCommentStorageProjection.Uuid storage_uuid = 5;\n\
}";
    const ROUTER_DECLARATIONS: [&str; 11] = [
        "const TEXT_FIELD: u32 = 1;",
        "const CREATION_DATE_FIELD: u32 = 2;",
        "const AUTHOR_FIELD: u32 = 3;",
        "const REPLIES_FIELD: u32 = 4;",
        "const STORAGE_UUID_FIELD: u32 = 5;",
        "pub trait CommentStorageVisitor",
        "pub fn decode_comment_storage_archive(",
        "pub fn decode_comment_storage_archive_with_report(",
        "pub fn decode_comment_storage_archive_with_visitor<'source>(",
        "pub fn decode_comment_storage_with_visitor<'source>(",
        "pub fn visit_comment_storage_replies<'source, F>(",
    ];
    const FORBIDDEN_PUBLIC_FUNCTION_FRAGMENTS: [&str; 4] =
        ["encode", "serialize", "to_owned", "write"];
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_comment_storage_generated {",
        "\"/buffa-numbers-comment-storage/iwa_comment_storage_buffa_protos.rs\"",
    ];

    let tsd = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSDCommentStorageArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let projection_digest = Sha256::digest(projection.as_bytes())
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let expected_projection_schema = PROJECTION_SCHEMA
        .lines()
        .map(str::trim)
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/comment_storage_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if !TSD_FIELDS
        .iter()
        .all(|declaration| tsd.matches(declaration).count() == 1)
        || tsp.matches(TSP_DATE).count() != 1
        || tsp.matches(TSP_REFERENCE).count() != 1
        || tsp.matches(TSP_UUID).count() != 1
        || projection_schema != expected_projection_schema
        || projection_digest != EXPECTED_PROJECTION_DIGEST
        || projection.len() > 3 * 1024
        || projection_schema.contains("repeated ")
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| production_codec.matches(declaration).count() == 1)
        || production_codec.contains("prost")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
        || production_codec.contains("pub fn encode_comment_storage_archive")
        || production_codec.contains("pub fn to_owned_comment_storage")
        || production_codec_has_forbidden_public_function(
            production_codec.as_ref(),
            &FORBIDDEN_PUBLIC_FUNCTION_FRAGMENTS,
        )
    {
        return Err(
            "derived Numbers comment-storage projection/router drifted from canonical TSD/TSP fields, exposed repeated generated replies, or introduced Prost/generated/production encoding"
                .into(),
        );
    }
    Ok(())
}

fn production_codec_has_forbidden_public_function(
    source: &str,
    forbidden_name_fragments: &[&str],
) -> bool {
    source.lines().any(|line| {
        let declaration = line.trim_start();
        let Some(signature) = declaration.strip_prefix("pub fn ") else {
            return false;
        };
        let name = signature
            .split(|character: char| !character.is_ascii_alphanumeric() && character != '_')
            .next()
            .unwrap_or_default();
        forbidden_name_fragments
            .iter()
            .any(|fragment| name.contains(fragment))
    })
}

fn enforce_group_node_category_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_DECLARATIONS: [&str; 2] =
        ["required uint64 lower = 1;", "required uint64 upper = 2;"];
    const TSCE_DECLARATIONS: [&str; 8] = [
        "required bool value = 1;",
        "required double value = 1;",
        "optional double value = 1;",
        "required string value = 1;",
        "optional .TSCE.BooleanCellValueArchive boolean_value = 2;",
        "optional .TSCE.DateCellValueArchive date_value = 3;",
        "optional .TSCE.NumberCellValueArchive number_value = 4;",
        "optional .TSCE.StringCellValueArchive string_value = 5;",
    ];
    const TST_DECLARATIONS: [&str; 3] = [
        "required .TSP.UUID group_uid = 1;",
        "repeated .TST.GroupByArchive.GroupNodeArchive child = 3;",
        "optional .TSCE.CellValueArchive group_cell_value = 7;",
    ];
    const PROJECTION_DECLARATIONS: [&str; 7] = [
        "required uint64 lower = 1;",
        "required uint64 upper = 2;",
        "required bool value = 1;",
        "required double value = 1;",
        "optional double value = 1;",
        "required string value = 1;",
        "message GroupNodeCategory {}",
    ];
    const ROUTER_DECLARATIONS: [&str; 7] = [
        "const GROUP_UID_FIELD: u32 = 1;",
        "const GROUP_CHILD_FIELD: u32 = 3;",
        "const GROUP_CELL_VALUE_FIELD: u32 = 7;",
        "const BOOLEAN_VALUE_FIELD: u32 = 2;",
        "const DATE_VALUE_FIELD: u32 = 3;",
        "const NUMBER_VALUE_FIELD: u32 = 4;",
        "const STRING_VALUE_FIELD: u32 = 5;",
    ];

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let tsce = fs::read_to_string(proto_directory.join("TSCEArchives.proto"))?;
    let tst = fs::read_to_string(proto_directory.join("TSTArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSTGroupNodeCategoryArchive.proto"))?;
    let router = fs::read_to_string("src/group_node_category_codec.rs")?;
    if !TSP_DECLARATIONS
        .iter()
        .all(|declaration| tsp.matches(declaration).count() == 1)
        || !TSCE_DECLARATIONS
            .iter()
            .all(|declaration| tsce.matches(declaration).count() == 1)
        || !TST_DECLARATIONS
            .iter()
            .all(|declaration| tst.matches(declaration).count() == 1)
        || !PROJECTION_DECLARATIONS
            .iter()
            .all(|declaration| projection.matches(declaration).count() == 1)
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| router.matches(declaration).count() == 1)
    {
        return Err(
            "derived GroupNode category projection is out of sync with its canonical TSP/TSCE/TST fields"
                .into(),
        );
    }
    Ok(())
}

fn enforce_keynote_document_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const KN_DOCUMENT: &str = "message DocumentArchive {\n  required .TSA.DocumentArchive super = 3;\n  required .TSP.Reference show = 2;\n  optional .TSP.Reference tables_custom_format_list = 4;\n}";
    const PROJECTION_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const PROJECTION_DOCUMENT: &str =
        "message KeynoteDocumentArchive {\n  required .LitchiIwaProjection.Reference show = 2;\n}";
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_keynote_document_generated {",
        "\"/buffa-keynote-document/iwa_keynote_document_buffa_protos.rs\"",
    ];

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let keynote = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("KNDocumentArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let expected_projection_schema = [
        "syntax = \"proto2\";",
        "package LitchiIwaProjection;",
        PROJECTION_REFERENCE,
        PROJECTION_DOCUMENT,
    ]
    .join("\n")
    .lines()
    .map(str::trim)
    .collect::<Vec<_>>()
    .join("\n");
    let codec = fs::read_to_string("src/keynote_document_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if tsp.matches(TSP_REFERENCE).count() != 1
        || keynote.matches(KN_DOCUMENT).count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_DOCUMENT).count() != 1
        || projection_schema != expected_projection_schema
        || projection.len() > 1024
        || projection.contains("repeated ")
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Keynote document/root codec drifted from KN.DocumentArchive.show or TSP.Reference.identifier, exceeded its 1 KiB source budget, exposed generated code, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_keynote_chart_caption_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const TSCH_DRAWABLE: &str = "message ChartDrawableArchive {\n  optional .TSD.DrawableArchive super = 1;\n  extensions 10000 to 536870911;\n}";
    const TSD_CAPTION: &str = "optional .TSP.Reference caption = 11;";
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message Reference {\n\
required uint64 identifier = 1;\n\
}\n\
message DrawableArchive {\n\
optional .LitchiIwaProjection.Reference caption = 11;\n\
}\n\
message ChartDrawableArchive {\n\
optional .LitchiIwaProjection.DrawableArchive super = 1;\n\
}";
    const ROUTER_DECLARATIONS: [&str; 7] = [
        "const CHART_DRAWABLE_SUPER_FIELD: u32 = 1;",
        "const DRAWABLE_CAPTION_FIELD: u32 = 11;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const MAX_RECURSION_LIMIT: u32 = 64;",
        "pub fn decode_chart_caption(",
        "fn preflight_chart_caption(",
        "fn next_strict_field<",
    ];
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_keynote_chart_caption_generated {",
        "\"/buffa-keynote-chart-caption/iwa_keynote_chart_caption_buffa_protos.rs\"",
    ];
    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let tsch = fs::read_to_string(proto_directory.join("TSCHArchives.proto"))?;
    let tsd = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSCHChartCaptionArchive.proto"))?;
    let normalize = |source: &str| {
        source
            .lines()
            .map(str::trim)
            .filter(|line| !line.is_empty() && !line.starts_with("//"))
            .collect::<Vec<_>>()
            .join("\n")
    };
    let codec = fs::read_to_string("src/keynote_chart_caption_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if tsp.matches(TSP_REFERENCE).count() != 1
        || tsch.matches(TSCH_DRAWABLE).count() != 1
        || tsd.matches(TSD_CAPTION).count() != 1
        || normalize(&projection) != normalize(PROJECTION_SCHEMA)
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || projection.contains("message ChartArchive")
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| production_codec.matches(declaration).count() == 1)
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || production_codec.contains("prost")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Keynote chart-caption projection/router drifted from canonical TSCH/TSD/TSP fields, exposed chart extensions, or introduced generated/production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_keynote_chart_title_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSCH_NON_STYLE_FIELDS: [&str; 3] = [
        "message ChartNonStyleArchive {",
        "optional bool tschchartinfodefaultshowtitle = 21;",
        "optional string tschchartinfodefaulttitle = 23;",
    ];
    const TSCH_NON_STYLE_EXTENSION: &str =
        "optional .TSCH.Generated.ChartNonStyleArchive current = 10000;";
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message ChartTitleArchive {\n\
optional bool tschchartinfodefaultshowtitle = 21;\n\
optional string tschchartinfodefaulttitle = 23;\n\
}";
    const ROUTER_DECLARATIONS: [&str; 8] = [
        "const CHART_TITLE_VISIBLE_FIELD: u32 = 21;",
        "const CHART_TITLE_TEXT_FIELD: u32 = 23;",
        "const MAX_RECURSION_LIMIT: u32 = 64;",
        "pub fn decode_chart_title<'source>(",
        "pub fn decode_chart_title_with_report<'source>(",
        "fn preflight_chart_title<'source>(",
        "fn next_strict_field<",
        "fn require_canonical_bool(",
    ];
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_keynote_chart_title_generated {",
        "\"/buffa-keynote-chart-title/iwa_keynote_chart_title_buffa_protos.rs\"",
    ];

    let tsch = fs::read_to_string(proto_directory.join("TSCHArchives.GEN.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TSCHChartTitleArchive.proto"))?;
    let normalize = |source: &str| {
        source
            .lines()
            .map(str::trim)
            .filter(|line| !line.is_empty() && !line.starts_with("//"))
            .collect::<Vec<_>>()
            .join("\n")
    };
    let expected_projection = normalize(PROJECTION_SCHEMA);
    let codec = fs::read_to_string("src/keynote_chart_title_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if !TSCH_NON_STYLE_FIELDS
        .iter()
        .all(|declaration| tsch.matches(declaration).count() == 1)
        || tsch.matches(TSCH_NON_STYLE_EXTENSION).count() != 1
        || normalize(&projection) != expected_projection
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| production_codec.matches(declaration).count() == 1)
        || production_codec.contains("prost")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
        || production_codec.contains("chart_graph")
        || production_codec.contains("IWorkPackage")
    {
        return Err(
            "derived Keynote chart-title projection/router drifted from TSCH.Generated.ChartNonStyleArchive fields 21/23, exposed the outer chart transaction, or introduced production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_table_info_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const TSD_DRAWABLE: &str = "message DrawableArchive {";
    const TSD_DRAWABLE_LOCKED: &str = "optional bool locked = 5;";
    const TST_TABLE_INFO: &str = "message TableInfoArchive {\n  required .TSD.DrawableArchive super = 1;\n  required .TSP.Reference tableModel = 2;\n  optional .TSP.Reference editing_state = 3 [deprecated = true];\n  optional .TSP.Reference summary_model = 4;\n  optional .TSP.Reference category_order = 5;\n  optional .TSP.Reference view_column_row_uids = 6;\n  optional .TSP.UUID group_by_uuid = 7;\n  optional .TSP.UUID hidden_states_uuid = 8;\n  optional uint32 formula_coord_space_in_pre40 = 9 [deprecated = true];\n  optional uint32 formula_coord_space = 10;\n  optional .TSCE.CoordMapperArchive pasteboard_coord_mapper = 13;\n  optional .TST.LayoutEngineArchive layout_engine = 14;\n  optional .TSP.Reference pivot_data_model = 15;\n  optional bool is_a_pivot_table = 16;\n  optional .TSP.Reference pivot_order = 17;\n}";
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message TableModelReference {\n\
required uint64 identifier = 1;\n\
}\n\
message DrawableArchive {\n\
optional bool locked = 5;\n\
}\n\
message TableInfoArchive {\n\
required .LitchiIwaProjection.DrawableArchive super = 1;\n\
required .LitchiIwaProjection.TableModelReference table_model = 2;\n\
}";

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let tsd = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let tst = fs::read_to_string(proto_directory.join("TSTArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TSTTableInfoArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/table_info_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if tsp.matches(TSP_REFERENCE).count() != 1
        || tsd.matches(TSD_DRAWABLE).count() != 1
        || tsd.matches(TSD_DRAWABLE_LOCKED).count() != 1
        || tst.matches(TST_TABLE_INFO).count() != 1
        || projection_schema != PROJECTION_SCHEMA
        || projection.len() > 1024
        || projection.contains("repeated ")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Numbers TableInfo projection drifted from TST.TableInfoArchive.super/tableModel, TSD.DrawableArchive.locked, or TSP.Reference.identifier; it may have exceeded its 1 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_numbers_names_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TN_SHEET_NAME: &str = "required string name = 1;";
    const TN_FORM_SHEET_SUPER: &str = "required .TN.SheetArchive super = 1;";
    const TST_TABLE_MODEL_FIELDS: [&str; 2] = [
        "required string table_id = 1;",
        "required string table_name = 8;",
    ];
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message NumbersSheetArchive {\n\
required string name = 1;\n\
}\n\
message NumbersFormBasedSheetArchive {\n\
required .LitchiIwaProjection.NumbersSheetArchive super = 1;\n\
}\n\
message NumbersTableModelArchive {\n\
required string table_id = 1;\n\
required string table_name = 8;\n\
}";
    const ROUTER_DECLARATIONS: [&str; 5] = [
        "const SHEET_NAME_FIELD: u32 = 1;",
        "const FORM_SHEET_SUPER_FIELD: u32 = 1;",
        "const TABLE_MODEL_ID_FIELD: u32 = 1;",
        "const TABLE_MODEL_NAME_FIELD: u32 = 8;",
        "const MAX_RECURSION: u32 = 64;",
    ];
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_numbers_names_generated {",
        "\"/buffa-numbers-names/iwa_numbers_names_buffa_protos.rs\"",
    ];

    let numbers = fs::read_to_string(proto_directory.join("TNArchives.proto"))?;
    let tables = fs::read_to_string(proto_directory.join("TSTArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TNNumbersNamesArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/numbers_names_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if numbers.matches(TN_SHEET_NAME).count() != 1
        || numbers.matches(TN_FORM_SHEET_SUPER).count() != 1
        || !TST_TABLE_MODEL_FIELDS
            .iter()
            .all(|declaration| tables.matches(declaration).count() == 1)
        || projection_schema != PROJECTION_SCHEMA
        || projection.len() > 1024
        || projection.contains("repeated ")
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| codec.matches(declaration).count() == 1)
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Numbers names projection/codec drifted from TN sheet/form or TST table-model fields, exceeded its 1 KiB source budget, exposed generated code, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_numbers_sheet_order_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message NumbersSheetReferenceArchive {\n\
required uint64 identifier = 1;\n\
optional int32 deprecated_type = 2;\n\
optional bool deprecated_is_external = 3;\n\
}";
    const ROUTER_DECLARATIONS: [&str; 12] = [
        "const DOCUMENT_SHEETS_FIELD: u32 = 1;",
        "const DOCUMENT_SIDEBAR_ORDER_FIELD: u32 = 5;",
        "const TREE_NODE_CHILDREN_FIELD: u32 = 2;",
        "const TREE_NODE_OBJECT_FIELD: u32 = 3;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;",
        "const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;",
        "const MAX_RECURSION: u32 = 64;",
        "pub fn decode_document_sheet_order(",
        "pub fn decode_document_sheet_order_with_report(",
        "pub fn decode_tree_node(",
        "pub fn decode_tree_node_with_report(",
    ];
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_numbers_sheet_order_generated {",
        "\"/buffa-numbers-sheet-order/iwa_numbers_sheet_order_buffa_protos.rs\"",
    ];
    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let tn = fs::read_to_string(proto_directory.join("TNArchives.proto"))?;
    let tsk = fs::read_to_string(proto_directory.join("TSKArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TNNumbersSheetReferenceArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/numbers_sheet_order_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if tsp.matches(REFERENCE).count() != 1
        || tn.matches("repeated .TSP.Reference sheets = 1;").count() != 1
        || tn
            .matches("required .TSP.Reference sidebar_order = 5;")
            .count()
            != 1
        || tsk.matches("repeated .TSP.Reference children = 2;").count() != 1
        || tsk.matches("optional .TSP.Reference object = 3;").count() != 1
        || projection_schema != PROJECTION_SCHEMA
        || projection.contains("repeated ")
        || projection.len() > 1024
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| codec.matches(declaration).count() == 1)
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || production_codec.contains("prost")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err("Numbers sheet-order projection/codec drifted from the exact TN/TSK/TSP reference routes, exposed generated code, introduced repeated storage, or added Prost/production encoding".into());
    }
    Ok(())
}

fn enforce_table_header_settings_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TST_FIELDS: [&str; 9] = [
        "required uint32 number_of_rows = 6;",
        "required uint32 number_of_columns = 7;",
        "optional uint32 number_of_header_rows = 9;",
        "optional uint32 number_of_header_columns = 10;",
        "optional uint32 number_of_footer_rows = 11;",
        "optional bool header_rows_frozen = 12;",
        "optional bool header_columns_frozen = 13;",
        "optional bool repeating_header_rows_enabled = 29;",
        "optional bool repeating_header_columns_enabled = 32;",
    ];
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message NumbersTableHeaderSettingsArchive {\n\
required uint32 number_of_rows = 6;\n\
required uint32 number_of_columns = 7;\n\
optional uint32 number_of_header_rows = 9;\n\
optional uint32 number_of_header_columns = 10;\n\
optional uint32 number_of_footer_rows = 11;\n\
optional bool header_rows_frozen = 12;\n\
optional bool header_columns_frozen = 13;\n\
optional bool repeating_header_rows_enabled = 29;\n\
optional bool repeating_header_columns_enabled = 32;\n\
}";
    const ROUTER_DECLARATIONS: [&str; 10] = [
        "const TABLE_ROWS_FIELD: u32 = 6;",
        "const TABLE_COLUMNS_FIELD: u32 = 7;",
        "const HEADER_ROWS_FIELD: u32 = 9;",
        "const HEADER_COLUMNS_FIELD: u32 = 10;",
        "const FOOTER_ROWS_FIELD: u32 = 11;",
        "const HEADER_ROWS_FROZEN_FIELD: u32 = 12;",
        "const HEADER_COLUMNS_FROZEN_FIELD: u32 = 13;",
        "const REPEATING_HEADER_ROWS_FIELD: u32 = 29;",
        "const REPEATING_HEADER_COLUMNS_FIELD: u32 = 32;",
        "const MAX_RECURSION: u32 = 64;",
    ];
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_numbers_table_header_settings_generated {",
        "\"/buffa-numbers-table-header-settings/iwa_numbers_table_header_settings_buffa_protos.rs\"",
    ];

    let tables = fs::read_to_string(proto_directory.join("TSTArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSTTableHeaderSettingsArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/numbers_table_header_settings_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if !TST_FIELDS
        .iter()
        .all(|declaration| tables.matches(declaration).count() == 1)
        || projection_schema != PROJECTION_SCHEMA
        || projection.len() > 1024
        || projection_schema.contains("repeated ")
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| codec.matches(declaration).count() == 1)
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Numbers table-header settings projection/codec drifted from TST.TableModelArchive scalar fields, exceeded its 1 KiB source budget, exposed generated code, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_table_title_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const TST_FIELDS: [&str; 5] = [
        "optional bool table_name_enabled = 22;",
        "optional .TSP.Reference table_name_style = 30;",
        "optional double table_name_height = 33;",
        "optional .TSP.Reference table_name_shape_style = 36;",
        "optional bool table_name_border_enabled = 37;",
    ];
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message TableTitleSettingsArchive {\n\
optional bool table_name_enabled = 22;\n\
optional fixed64 table_name_height_bits = 33;\n\
optional bool table_name_border_enabled = 37;\n\
}";
    const ROUTER_DECLARATIONS: [&str; 14] = [
        "const TABLE_NAME_ENABLED_FIELD: u32 = 22;",
        "const TABLE_NAME_STYLE_FIELD: u32 = 30;",
        "const TABLE_NAME_HEIGHT_FIELD: u32 = 33;",
        "const TABLE_NAME_SHAPE_STYLE_FIELD: u32 = 36;",
        "const TABLE_NAME_BORDER_ENABLED_FIELD: u32 = 37;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;",
        "const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;",
        "const MAX_RECURSION: u32 = 64;",
        "const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;",
        "const MIN_SIGN_EXTENDED_I32: u64 = 0xffff_ffff_8000_0000;",
        "pub fn decode_table_title_settings(",
        "pub fn decode_table_title_settings_with_report(",
        "reference_projection::NumbersSheetReferenceArchiveLazyView<'_>",
    ];
    const PRIVATE_DECLARATIONS: [&str; 4] = [
        "#[doc(hidden)]\nmod buffa_numbers_table_title_generated {",
        "\"/buffa-numbers-table-title/iwa_numbers_table_title_buffa_protos.rs\"",
        "pub mod numbers_table_title_codec;",
        "mod buffa_numbers_sheet_order_generated {",
    ];
    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let canonical = fs::read_to_string(proto_directory.join("TSTArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSTTableTitleSettingsArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/numbers_table_title_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if tsp.matches(TSP_REFERENCE).count() != 1
        || !TST_FIELDS
            .iter()
            .all(|field| canonical.matches(field).count() == 1)
        || projection_schema != PROJECTION_SCHEMA
        || projection.len() > 1024
        || projection_schema.contains("repeated ")
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| production_codec.matches(declaration).count() == 1)
        || !PRIVATE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || production_codec
            .contains("buffa_numbers_sheet_order_generated::LitchiIwaProjection as projection")
        || production_codec.contains("RepeatedView")
        || production_codec.contains("LazyRepeatedView")
        || production_codec.contains("prost::")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err("Numbers table-title projection/codec drifted from the exact TST/TSP scalar routes, lost its private generated boundary or shared reference lazy view, introduced generated repeated storage, or added Prost/production encoding".into());
    }
    Ok(())
}

fn enforce_table_cell_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const REQUIRED_CANONICAL: [&str; 18] = [
        "message TileRowInfo {",
        "message TileStorage {",
        "message TableDataList {",
        "message TableDataListSegment {",
        "message HeaderStorageBucket {",
        "message HeaderStorage {",
        "message DataStore {",
        "message TableModelArchive {",
        "message CellRecordExpandedArchive {",
        "message CellRecordTileArchive {",
        "message RangeBackDependencyArchive {",
        "message RangePrecedentsTileArchive {",
        "message FormulaOwnerDependenciesArchive {",
        "message DependencyTrackerArchive {",
        "message CalculationEngineArchive {",
        "repeated .TST.TileRowInfo rowInfos = 5;",
        "repeated .TSP.Reference formula_owner_dependencies = 6;",
        "required .TSCE.DependencyTrackerArchive dependency_tracker = 2;",
    ];
    const STORAGE_CANONICAL_FIELDS: &[&str] = &[
        "required string table_id = 1;",
        "required .TST.DataStore base_data_store = 4;",
        "required uint32 number_of_rows = 6;",
        "required uint32 number_of_columns = 7;",
        "required string table_name = 8;",
        "optional .TSP.Reference hidden_state_formula_owner_for_columns = 34;",
        "optional .TSP.Reference hidden_state_formula_owner_for_rows = 35;",
        "optional .TSP.CFUUIDArchive conditional_style_formula_owner_id = 39;",
        "optional .TSP.Reference pivot_owner = 85;",
        "optional .TSP.Reference category_owner = 86;",
        "optional .TSCE.SpillOwnerArchive spill_owner = 93;",
        "required .TST.HeaderStorage rowHeaders = 1;",
        "required .TSP.Reference columnHeaders = 2;",
        "required .TST.TileStorage tiles = 3;",
        "required .TSP.Reference stringTable = 4;",
        "required .TSP.Reference styleTable = 5;",
        "required .TSP.Reference formula_table = 6;",
        "required uint32 nextRowStripID = 7;",
        "required uint32 nextColumnStripID = 8;",
        "required .TST.TableRBTree rowTileTree = 9;",
        "required .TST.TableRBTree columnTileTree = 10;",
        "required .TSP.Reference format_table_pre_bnc = 11;",
        "optional .TSP.Reference formulaErrorTable = 12;",
        "optional .TSP.Reference merge_region_map = 13;",
        "optional uint32 storage_version_pre_bnc = 14;",
        "optional .TSP.Reference deprecated_custom_format_table = 15;",
        "optional .TSP.Reference multipleChoiceListFormatTable = 16;",
        "optional .TSP.Reference rich_text_table = 17;",
        "optional .TSP.Reference conditionalstyletable = 18;",
        "optional .TSP.Reference commentStorageTable = 19;",
        "optional .TSP.Reference importWarningSetTable = 20;",
        "optional .TSP.Reference control_cell_spec_table = 21;",
        "optional .TSP.Reference format_table = 22;",
        "repeated .TST.TileStorage.Tile tiles = 1;",
        "required uint32 tileid = 1;",
        "required .TSP.Reference tile = 2;",
        "optional uint32 tile_size = 2;",
        "optional bool should_use_wide_rows = 3;",
        "required uint32 maxColumn = 1;",
        "required uint32 maxRow = 2;",
        "required uint32 numCells = 3;",
        "required uint32 numrows = 4;",
        "repeated .TST.TileRowInfo rowInfos = 5;",
        "optional uint32 storage_version = 6;",
        "optional bool last_saved_in_BNC = 7;",
        "optional bool should_use_wide_rows = 8;",
        "required uint32 tile_row_index = 1;",
        "required uint32 cell_count = 2;",
        "required bytes cell_storage_buffer_pre_bnc = 3;",
        "required bytes cell_offsets_pre_bnc = 4;",
        "optional uint32 storage_version = 5;",
        "optional bytes cell_storage_buffer = 6;",
        "optional bytes cell_offsets = 7;",
        "optional bool has_wide_offsets = 8;",
        "required uint32 bucketHashFunction = 1;",
        "repeated .TSP.Reference buckets = 2;",
        "repeated .TST.HeaderStorageBucket.Header headers = 2;",
        "required uint32 index = 1;",
        "required float size = 2;",
        "required uint32 hidingState = 3;",
        "required uint32 numberOfCells = 4;",
        "optional .TSP.Reference cell_style = 5;",
        "optional .TSP.Reference text_style = 6;",
        "required .TST.TableDataList.ListType listType = 1;",
        "required uint32 nextListID = 2;",
        "repeated .TST.TableDataList.ListEntry entries = 3;",
        "repeated .TSP.Reference segments = 4;",
        "optional bool is_new_for_bnc = 5;",
        "required uint32 key = 1;",
        "required uint32 refcount = 2;",
        "optional string string = 3;",
        "optional .TSP.Reference reference = 4;",
        "optional .TSCE.FormulaArchive formula = 5;",
        "optional .TSK.FormatStructArchive format = 6;",
        "optional .TSK.CustomFormatArchive custom_format = 8;",
        "optional .TSP.Reference rich_text_payload = 9;",
        "optional .TSP.Reference comment_storage = 10;",
        "optional .TST.ImportWarningSetArchive import_warning_set = 11;",
        "optional .TST.CellSpecArchive cell_spec = 12;",
        "required .TST.TableDataList.ListType list_type = 1;",
        "required .TSP.Range key_range = 2;",
    ];
    const DEPENDENCY_CANONICAL_FIELDS: &[&str] = &[
        "optional bool base_date_1904 = 1;",
        "required .TSCE.DependencyTrackerArchive dependency_tracker = 2;",
        "optional .TSP.Reference named_reference_manager = 3;",
        "optional .TSP.Reference remote_data_store = 12;",
        "optional .TSP.Reference header_name_manager = 14;",
        "optional .TSP.Reference refs_to_dirty = 15;",
        "optional .TSCE.OwnerIDMapArchive owner_id_map = 3;",
        "optional uint64 number_of_formulas = 5;",
        "repeated .TSP.Reference formula_owner_dependencies = 6;",
        "required .TSP.UUID formula_owner_uid = 1;",
        "required uint32 internal_formula_owner_id = 2;",
        "optional uint32 owner_kind = 3 [default = 0];",
        "optional .TSCE.CellDependenciesExpandedArchive cell_dependencies = 4;",
        "optional .TSCE.RangeDependenciesArchive range_dependencies = 5;",
        "optional .TSCE.VolatileDependenciesExpandedArchive volatile_dependencies = 6;",
        "optional .TSCE.SpanningDependenciesExpandedArchive spanning_column_dependencies = 7;",
        "optional .TSCE.SpanningDependenciesExpandedArchive spanning_row_dependencies = 8;",
        "optional .TSCE.WholeOwnerDependenciesExpandedArchive whole_owner_dependencies = 9;",
        "optional .TSCE.CellErrorsArchive cell_errors = 10;",
        "optional .TSP.Reference formula_owner = 11;",
        "optional .TSP.UUID base_owner_uid = 12;",
        "optional .TSCE.CellDependenciesTiledArchive tiled_cell_dependencies = 13;",
        "optional .TSCE.UuidReferencesArchive uuid_references = 14;",
        "optional .TSCE.RangeDependenciesTiledArchive tiled_range_dependencies = 15;",
        "optional .TSCE.CellSpillSizesArchive spill_range_sizes = 16;",
        "required uint32 column = 1;",
        "required uint32 row = 2;",
        "optional uint64 dirty_self_plus_precedents_count = 3 [default = 0];",
        "optional bool is_in_a_cycle = 4 [default = false];",
        "optional bool has_calculated_precedents = 5 [default = false];",
        "optional .TSCE.ExpandedEdgesArchive expanded_edges = 6;",
        "repeated .TSCE.CellRecordExpandedArchive cell_record = 1;",
        "repeated .TSCE.RangeBackDependencyArchive back_dependency = 2;",
        "required uint32 internal_owner_id = 1;",
        "required uint32 tile_column_begin = 2;",
        "required uint32 tile_row_begin = 3;",
        "repeated .TSCE.CellRecordExpandedArchive cell_records = 4;",
        "required uint32 cell_coord_row = 1;",
        "required uint32 cell_coord_column = 2;",
        "optional .TSCE.RangeReferenceArchive range_reference = 3;",
        "optional .TSCE.InternalRangeReferenceArchive internal_range_reference = 4;",
        "required uint32 to_owner_id = 1;",
        "repeated .TSCE.RangePrecedentsTileArchive.FromToRangeArchive from_to_range = 2;",
        "required .TSCE.CellCoordinateArchive from_coord = 1;",
        "required .TSCE.CellRectArchive refers_to_rect = 2;",
        "repeated .TSP.Reference cell_record_tiles = 1;",
        "repeated .TSP.Reference range_precedents_tile = 1;",
        "repeated uint32 edge_without_owner_rows = 1;",
        "repeated uint32 edge_without_owner_columns = 2;",
        "repeated uint32 edge_with_owner_rows = 3;",
        "repeated uint32 edge_with_owner_columns = 4;",
        "repeated uint32 internal_owner_id_for_edge = 5;",
        "optional fixed32 packedData = 1;",
        "required .TSCE.CellCoordinateArchive origin = 1;",
        "required .TSCE.ColumnRowSize size = 2;",
        "required .TSP.CFUUIDArchive table_id = 1;",
        "required .TSCE.RangeCoordinateArchive range = 2;",
    ];
    const TSP_CANONICAL_FIELDS: &[&str] = &[
        "required uint64 identifier = 1;",
        "optional int32 deprecated_type = 2;",
        "optional bool deprecated_is_external = 3;",
        "required uint32 location = 1;",
        "required uint32 length = 2;",
        "required uint64 lower = 1;",
        "required uint64 upper = 2;",
        "optional bytes uuid_bytes = 1;",
        "optional uint32 uuid_w0 = 2;",
        "optional uint32 uuid_w1 = 3;",
        "optional uint32 uuid_w2 = 4;",
        "optional uint32 uuid_w3 = 5;",
    ];
    const DEPENDENCY_PROJECTION_MESSAGES: &[&str] = &[
        "message ExpandedEdgesArchive {}",
        "message CellCoordinateArchive {\n  optional fixed32 packed_data = 1;\n  optional uint32 column = 2;\n  optional uint32 row = 3;\n}",
        "message ColumnRowSize {\n  optional uint32 num_columns = 1;\n  optional uint32 num_rows = 2;\n}",
        "message CellRectArchive {\n  required bytes origin = 1;\n  required bytes size = 2;\n}",
        "message CFUUIDArchive {\n  optional bytes uuid_bytes = 1;\n  optional uint32 uuid_w0 = 2;\n  optional uint32 uuid_w1 = 3;\n  optional uint32 uuid_w2 = 4;\n  optional uint32 uuid_w3 = 5;\n}",
        "message RangeReferenceArchive {\n  required bytes table_id = 1;\n  required uint32 top_left_column = 2;\n  required uint32 top_left_row = 3;\n  required uint32 bottom_right_column = 4;\n  required uint32 bottom_right_row = 5;\n}",
        "message RangeCoordinateArchive {\n  required uint32 top_left_column = 1;\n  required uint32 top_left_row = 2;\n  required uint32 bottom_right_column = 3;\n  required uint32 bottom_right_row = 4;\n}",
        "message InternalRangeReferenceArchive {\n  required uint32 owner_id = 1;\n  required bytes range = 2;\n}",
    ];
    const ROUTER_SYMBOLS: &[&str] = &[
        "pub fn decode_table_model(",
        "pub fn decode_data_store(",
        "pub fn decode_tile_storage(",
        "pub fn decode_tile(",
        "pub fn decode_tile_row_info(",
        "pub fn decode_header_storage(",
        "pub fn decode_header_storage_bucket(",
        "pub fn decode_header(",
        "pub fn rewrite_header_storage_bucket_sizes(",
        "pub fn plan_header_storage_bucket_sizes<'source>(",
        "pub fn execute_header_storage_bucket_size_plan(",
        "pub fn decode_table_data_list(",
        "pub fn decode_table_data_list_entry(",
        "pub fn decode_table_data_list_segment(",
        "pub fn decode_calculation_engine(",
        "pub fn decode_dependency_tracker(",
        "pub fn decode_formula_owner_dependencies(",
        "pub fn decode_cell_record(",
        "pub fn decode_cell_record_tile(",
        "pub fn decode_range_back_dependency(",
        "pub fn decode_range_precedents_tile(",
        "pub fn decode_from_to_range(",
        "pub fn decode_expanded_edges(",
        "pub fn decode_cell_coordinate(",
        "pub fn decode_range_reference(",
        "pub fn decode_internal_range_reference(",
        "pub trait StorageVisitor",
        "pub struct HeaderRecord<'source>",
        "fn visit_header_record(&mut self, record: HeaderRecord<'_>)",
        "pub trait DependencyVisitor",
        "if snapshot.deprecated_is_external == Some(true)",
        "payloads: <redacted>",
    ];
    let tst = fs::read_to_string(proto_directory.join("TSTArchives.proto"))?;
    let tsce = fs::read_to_string(proto_directory.join("TSCEArchives.proto"))?;
    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let canonical = format!("{tst}\n{tsce}");
    let storage =
        fs::read_to_string(projection_directory.join("TSTTableCellStorageArchive.proto"))?;
    let dependency =
        fs::read_to_string(projection_directory.join("TSCETableCellDependenciesArchive.proto"))?;
    let storage_codec = fs::read_to_string("src/numbers_table_cell_storage_codec.rs")?;
    let dependency_codec = fs::read_to_string("src/numbers_table_cell_dependency_codec.rs")?;
    let lib = fs::read_to_string("src/lib.rs")?;
    let production = [storage_codec.as_str(), dependency_codec.as_str()]
        .into_iter()
        .map(production_codec_source)
        .collect::<Vec<_>>()
        .join("\n");
    if !REQUIRED_CANONICAL
        .iter()
        .all(|declaration| canonical.contains(declaration))
        || !STORAGE_CANONICAL_FIELDS
            .iter()
            .all(|field| tst.contains(field))
        || !DEPENDENCY_CANONICAL_FIELDS
            .iter()
            .all(|field| tsce.contains(field))
        || !TSP_CANONICAL_FIELDS.iter().all(|field| tsp.contains(field))
        || !DEPENDENCY_PROJECTION_MESSAGES
            .iter()
            .all(|message| dependency.matches(message).count() == 1)
        || !ROUTER_SYMBOLS
            .iter()
            .all(|symbol| production.contains(symbol))
        || storage.contains("repeated ")
        || dependency.contains("repeated ")
        || storage.len() > 5 * 1024
        || dependency.len() > 5 * 1024
        || !lib.contains("mod buffa_numbers_table_cell_storage_generated {")
        || !lib.contains("mod buffa_numbers_table_cell_dependency_generated {")
        || !lib.contains("pub mod numbers_table_cell_storage_codec;")
        || !lib.contains("pub mod numbers_table_cell_dependency_codec;")
        || production.contains("RepeatedView")
        || production.contains("LazyRepeatedView")
        || production.contains("prost::")
        || production.contains("to_owned_message")
        || production.contains("encode_to_vec")
        || production.contains("try_encode")
        || production.contains(".encode(")
    {
        return Err("Numbers table-cell projections/codecs drifted from canonical TST/TSCE envelopes, exceeded source budgets, exposed repeated generated storage, or introduced production encoding".into());
    }
    Ok(())
}

fn enforce_numbers_table_data_list_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    // These are the native TST routes that reach the two strict list codecs:
    // 6005/6201 carry TableDataList and 6011 carries TableDataListSegment.
    // Keep the IDs explicit here even though the wire schema itself has no
    // archive-type field; the workspace route checks below tie those IDs to
    // the application adapters when their sibling sources are present.
    const TABLE_DATA_LIST_NATIVE_IDS: [u32; 2] = [6005, 6201];
    const TABLE_DATA_LIST_SEGMENT_NATIVE_ID: u32 = 6011;

    const TST_TABLE_DATA_LIST_DIGEST: &str =
        "679dfddf38cfaf6164586de3fa1730e9c32793158db9cc5ccab882393bc4d26b";
    const TST_TABLE_DATA_LIST_SEGMENT_DIGEST: &str =
        "55134c3152ac35abbbd4b52bd737708bd23c9ea663fe8fd8edf25ada9585d8b0";
    const TSP_RANGE_DIGEST: &str =
        "98a79cc93228d486c83088884c21bcb2180ec18d58feba907516dc12f81bca24";
    const PROJECTION_TABLE_DATA_LIST_DIGEST: &str =
        "2aeff27e1e4284310c2bd88566eac5fdb754e62584f1f81df47a9741197225cc";
    const PROJECTION_TABLE_DATA_LIST_ENTRY_DIGEST: &str =
        "4d5af02897609de02ecaa66b48518d35467cb3f4209e195779530eb19102e0b7";
    const PROJECTION_TABLE_DATA_LIST_SEGMENT_DIGEST: &str =
        "81c497d1902a3f07684d753c28760c7ac5faf82ca0c70efb4eed78db483a7af6";

    const TST_TABLE_DATA_LIST_FIELDS: [&str; 5] = [
        "required .TST.TableDataList.ListType listType = 1;",
        "required uint32 nextListID = 2;",
        "repeated .TST.TableDataList.ListEntry entries = 3;",
        "repeated .TSP.Reference segments = 4;",
        "optional bool is_new_for_bnc = 5;",
    ];
    const TST_TABLE_DATA_LIST_ENTRY_FIELDS: [&str; 11] = [
        "required uint32 key = 1;",
        "required uint32 refcount = 2;",
        "optional string string = 3;",
        "optional .TSP.Reference reference = 4;",
        "optional .TSCE.FormulaArchive formula = 5;",
        "optional .TSK.FormatStructArchive format = 6;",
        "optional .TSK.CustomFormatArchive custom_format = 8;",
        "optional .TSP.Reference rich_text_payload = 9;",
        "optional .TSP.Reference comment_storage = 10;",
        "optional .TST.ImportWarningSetArchive import_warning_set = 11;",
        "optional .TST.CellSpecArchive cell_spec = 12;",
    ];
    const TST_TABLE_DATA_LIST_SEGMENT_FIELDS: [&str; 3] = [
        "required .TST.TableDataList.ListType list_type = 1;",
        "required .TSP.Range key_range = 2;",
        "repeated .TST.TableDataList.ListEntry entries = 3;",
    ];
    const TSP_RANGE_FIELDS: [&str; 2] = [
        "required uint32 location = 1;",
        "required uint32 length = 2;",
    ];
    const PROJECTION_TABLE_DATA_LIST_FIELDS: [&str; 3] = [
        "required int32 list_type = 1;",
        "required uint32 next_list_id = 2;",
        "optional bool is_new_for_bnc = 5;",
    ];
    const PROJECTION_TABLE_DATA_LIST_ENTRY_FIELDS: [&str; 11] = [
        "required uint32 key = 1;",
        "required uint32 ref_count = 2;",
        "optional string string_value = 3;",
        "optional bytes reference = 4;",
        "optional bytes formula = 5;",
        "optional bytes format = 6;",
        "optional bytes custom_format = 8;",
        "optional bytes rich_text_payload = 9;",
        "optional bytes comment_storage = 10;",
        "optional bytes import_warning_set = 11;",
        "optional bytes cell_spec = 12;",
    ];
    const PROJECTION_TABLE_DATA_LIST_SEGMENT_FIELDS: [&str; 2] = [
        "required int32 list_type = 1;",
        "required bytes key_range = 2;",
    ];
    const SEGMENT_CODEC_FIELDS: [&str; 13] = [
        "pub struct TableDataListSegmentSnapshot<'source> {",
        "key_range_location: u32,",
        "key_range_length: u32,",
        "pub fn decode_table_data_list_segment_with_visitor<'source>(",
        "fn decode_table_data_list_segment_in<'source>(",
        "let mut key_range = None;",
        "let mut key_range_location = None;",
        "let mut key_range_length = None;",
        "let (location, length) = decode_range(raw, budget, child_depth)?;",
        "key_range = Some(raw);",
        "3 => visitor.visit_list_entry(decode_table_data_list_entry_in(",
        "key_range_location: key_range_location.ok_or_else(DecodeError::invalid)?,",
        "key_range_length: key_range_length.ok_or_else(DecodeError::invalid)?,",
    ];
    const SEGMENT_CODEC_PARITY: &str =
        "if view.list_type != snapshot.list_type || view.key_range != snapshot.key_range {";

    let tst = fs::read_to_string(proto_directory.join("TSTArchives.proto"))?;
    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSTTableCellStorageArchive.proto"))?;
    let codec = fs::read_to_string("src/numbers_table_cell_storage_codec.rs")?;
    let Some(table_data_list) = proto_message_block(&tst, "TableDataList") else {
        return Err("Numbers TableDataList provenance lost its canonical message block".into());
    };
    let Some(table_data_list_entry) = proto_nested_message_block(table_data_list, "ListEntry")
    else {
        return Err("Numbers TableDataList provenance lost its nested ListEntry block".into());
    };
    let Some(table_data_list_segment) = proto_message_block(&tst, "TableDataListSegment") else {
        return Err(
            "Numbers TableDataListSegment provenance lost its canonical message block".into(),
        );
    };
    let Some(range) = proto_message_block(&tsp, "Range") else {
        return Err("Numbers TableDataList provenance lost TSP.Range".into());
    };
    let Some(projection_table_data_list) = proto_message_block(&projection, "TableDataListArchive")
    else {
        return Err("Numbers TableDataList provenance lost its projection block".into());
    };
    let Some(projection_table_data_list_entry) =
        proto_message_block(&projection, "TableDataListEntryArchive")
    else {
        return Err("Numbers TableDataList provenance lost its entry projection block".into());
    };
    let Some(projection_table_data_list_segment) =
        proto_message_block(&projection, "TableDataListSegmentArchive")
    else {
        return Err("Numbers TableDataListSegment provenance lost its projection block".into());
    };

    let canonical_scope_ok = TST_TABLE_DATA_LIST_FIELDS
        .iter()
        .all(|field| proto_field(table_data_list, field) == 1)
        && TST_TABLE_DATA_LIST_ENTRY_FIELDS
            .iter()
            .all(|field| proto_field(table_data_list_entry, field) == 1)
        && TST_TABLE_DATA_LIST_SEGMENT_FIELDS
            .iter()
            .all(|field| proto_field(table_data_list_segment, field) == 1)
        && TSP_RANGE_FIELDS
            .iter()
            .all(|field| proto_field(range, field) == 1)
        && sha256_hex(table_data_list) == TST_TABLE_DATA_LIST_DIGEST
        && sha256_hex(table_data_list_segment) == TST_TABLE_DATA_LIST_SEGMENT_DIGEST
        && sha256_hex(range) == TSP_RANGE_DIGEST;
    let projection_scope_ok = PROJECTION_TABLE_DATA_LIST_FIELDS
        .iter()
        .all(|field| proto_field(projection_table_data_list, field) == 1)
        && PROJECTION_TABLE_DATA_LIST_ENTRY_FIELDS
            .iter()
            .all(|field| proto_field(projection_table_data_list_entry, field) == 1)
        && PROJECTION_TABLE_DATA_LIST_SEGMENT_FIELDS
            .iter()
            .all(|field| proto_field(projection_table_data_list_segment, field) == 1)
        && !projection_table_data_list_segment.contains("repeated ")
        && sha256_hex(projection_table_data_list) == PROJECTION_TABLE_DATA_LIST_DIGEST
        && sha256_hex(projection_table_data_list_entry) == PROJECTION_TABLE_DATA_LIST_ENTRY_DIGEST
        && sha256_hex(projection_table_data_list_segment)
            == PROJECTION_TABLE_DATA_LIST_SEGMENT_DIGEST;
    let codec_scope_ok = SEGMENT_CODEC_FIELDS
        .iter()
        .all(|field| codec.contains(field))
        && codec.contains(SEGMENT_CODEC_PARITY);

    let route_paths = [
        Path::new("../litchi-iwa/src/protobuf.rs"),
        Path::new("../litchi-numbers/src/package/extractor.rs"),
    ];
    let registry_table_data_list_markers = [
        format!(
            "{}u32 => decode_table_data_list,",
            TABLE_DATA_LIST_NATIVE_IDS[0]
        ),
        format!(
            "{}u32 => decode_table_data_list,",
            TABLE_DATA_LIST_NATIVE_IDS[1]
        ),
    ];
    let registry_table_data_list_segment_marker = format!(
        "{}u32 => decode_table_data_list_segment,",
        TABLE_DATA_LIST_SEGMENT_NATIVE_ID
    );
    let extractor_table_data_list_marker = format!(
        ".filter(|message| message.type_ == {} || message.type_ == {})",
        TABLE_DATA_LIST_NATIVE_IDS[0], TABLE_DATA_LIST_NATIVE_IDS[1]
    );
    let extractor_table_data_list_segment_marker = format!(
        ".filter(|message| message.type_ == {})",
        TABLE_DATA_LIST_SEGMENT_NATIVE_ID
    );
    let route_sources = route_paths
        .iter()
        .filter(|path| path.is_file())
        .map(fs::read_to_string)
        .collect::<Result<Vec<_>, _>>()?;
    let route_scope_ok = if route_sources.is_empty() {
        true
    } else if route_sources.len() != route_paths.len() {
        false
    } else {
        route_sources[0]
            .matches(registry_table_data_list_markers[0].as_str())
            .count()
            == 1
            && route_sources[0]
                .matches(registry_table_data_list_markers[1].as_str())
                .count()
                == 1
            && route_sources[0]
                .matches(registry_table_data_list_segment_marker.as_str())
                .count()
                == 1
            && route_sources[1]
                .matches(extractor_table_data_list_marker.as_str())
                .count()
                == 1
            && route_sources[1]
                .matches(extractor_table_data_list_segment_marker.as_str())
                .count()
                == 1
    };

    if !canonical_scope_ok || !projection_scope_ok || !codec_scope_ok || !route_scope_ok {
        return Err(
            "Numbers TableDataList/TableDataListSegment provenance drifted: native 6005/6201/6011 routes, message-scoped fields/digests, entries=3, key_range, or parsed segment fields no longer match"
                .into(),
        );
    }
    Ok(())
}

fn enforce_package_metadata_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const CANONICAL: &[&str] = &[
        "message PackageMetadata {",
        "required uint64 last_object_identifier = 1;",
        "repeated .TSP.ComponentInfo components = 3;",
        "repeated .TSP.ComponentInfo versioned_components = 11;",
        "message ComponentInfo {",
        "required uint64 identifier = 1;",
        "required string preferred_locator = 2;",
        "optional string locator = 3;",
        "repeated .TSP.ComponentExternalReference external_references = 6;",
        "repeated .TSP.ComponentDataReference data_references = 7;",
        "repeated .TSP.ObjectUUIDMapEntry object_uuid_map_entries = 11;",
        "repeated .TSP.ComponentExternalReference versioned_external_references = 18;",
        "repeated uint64 ambiguous_object_identifiers = 20 [packed = true];",
        "message ComponentExternalReference {",
        "required uint64 component_identifier = 1;",
        "optional uint64 object_identifier = 2;",
        "optional bool is_weak = 3;",
        "message ComponentDataReference {",
        "message ObjectReference {",
        "required uint64 object_identifier = 1;",
        "required uint32 count = 2;",
        "required uint64 data_identifier = 1;",
        "repeated .TSP.ComponentDataReference.ObjectReference object_reference_list = 2;",
        "message ObjectUUIDMapEntry {",
        "required uint64 identifier = 1;",
        "required .TSP.UUID uuid = 2;",
        "message UUID {",
        "required uint64 lower = 1;",
        "required uint64 upper = 2;",
    ];
    let canonical = fs::read_to_string(proto_directory.join("TSPArchiveMessages.proto"))?;
    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSPPackageMetadataArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    const EXPECTED_PROJECTION_SCHEMA: &str = r#"syntax = "proto2";
package LitchiIwaPackageMetadataProjection;
message PackageMetadataArchive {
required uint64 last_object_identifier = 1;
}
message ComponentInfoArchive {
required uint64 identifier = 1;
required string preferred_locator = 2;
optional string locator = 3;
}
message ComponentExternalReferenceArchive {
required uint64 component_identifier = 1;
optional uint64 object_identifier = 2;
optional bool is_weak = 3;
}
message ObjectUUIDMapEntryArchive {
required uint64 identifier = 1;
required bytes uuid = 2;
}
message UUIDArchive {
required uint64 lower = 1;
required uint64 upper = 2;
}"#;
    const CODEC_SYMBOLS: &[&str] = &[
        "pub struct RewriteOptions",
        "pub struct ComponentSelector",
        "pub struct ObjectUuidAddition",
        "pub struct ExternalReferenceAddition",
        "pub struct Batch",
        "pub struct ObjectUuidRemoval",
        "pub struct ExternalReferenceRemoval",
        "pub struct DataReferenceOwnerRemoval",
        "pub struct RemovalBatch",
        "pub struct ComponentDescriptor",
        "pub struct ObjectUuidDescriptor",
        "pub struct ExternalReferenceDescriptor",
        "pub trait PackageMetadataVisitor",
        "pub fn inspect_package_metadata_with_visitor",
        "pub fn rewrite_package_metadata",
        "pub fn remove_package_metadata",
        "precharge_rewrite_and_verification",
        "try_reserve_exact",
    ];
    let codec = fs::read_to_string("src/package_metadata_codec.rs")?;
    let production = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if !CANONICAL
        .iter()
        .all(|declaration| canonical.contains(declaration) || tsp.contains(declaration))
        || projection_schema != EXPECTED_PROJECTION_SCHEMA
        || projection.contains("repeated ")
        || projection.len() > 2 * 1024
        || lib
            .matches("mod buffa_package_metadata_generated {")
            .count()
            != 1
        || lib.matches("pub mod package_metadata_codec;").count() != 1
        || lib.contains("pub mod buffa_package_metadata_generated")
        || !CODEC_SYMBOLS
            .iter()
            .all(|symbol| production.contains(symbol))
        || production.contains("RepeatedView")
        || production.contains("LazyRepeatedView")
        || production.contains("prost::")
        || production.contains("to_owned_message")
        || production.contains("encode_to_vec")
        || production.contains("try_encode")
        || production.contains(".encode(")
    {
        return Err("PackageMetadata projection/codec drifted from canonical TSP registry fields, exceeded its source budget, exposed generated repeated storage, or introduced generated/Prost encoding".into());
    }
    Ok(())
}

fn enforce_formula_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const CANONICAL: &[&str] = &[
        "message FormulaArchive {",
        "required .TSCE.ASTNodeArrayArchive AST_node_array = 1;",
        "optional uint32 host_column = 2;",
        "optional uint32 host_row = 3;",
        "message ASTNodeArrayArchive {",
        "repeated .TSCE.ASTNodeArrayArchive.ASTNodeArchive AST_node = 1;",
        "message ASTNodeArchive {",
        "required .TSCE.ASTNodeArrayArchive.ASTNodeType AST_node_type = 1;",
        "optional uint32 AST_function_node_index = 2;",
        "optional uint32 AST_function_node_numArgs = 3;",
        "optional double AST_number_node_number = 4;",
        "optional uint64 AST_number_node_decimal_low = 42;",
        "optional uint64 AST_number_node_decimal_high = 43;",
        "optional bool AST_boolean_node_boolean = 5;",
        "optional string AST_string_node_string = 6;",
        "optional bool AST_token_node_boolean = 10;",
        "optional .TSCE.ASTNodeArrayArchive.ASTLocalCellReferenceNodeArchive AST_local_cell_reference_node_reference = 15;",
        "optional .TSCE.ASTNodeArrayArchive.ASTColumnCoordinateArchive AST_column = 26;",
        "optional .TSCE.ASTNodeArrayArchive.ASTRowCoordinateArchive AST_row = 27;",
        "optional .TSCE.ASTNodeArrayArchive.ASTCrossTableReferenceExtraInfoArchive AST_cross_table_reference_extra_info = 28;",
        "optional .TSCE.ASTNodeArrayArchive.ASTStickyBits AST_sticky_bits = 33;",
        "optional .TSCE.ASTNodeArrayArchive.ASTColonTractArchive AST_colon_tract = 40;",
        "message ASTCrossTableReferenceExtraInfoArchive {",
        "required .TSP.CFUUIDArchive table_id = 1;",
        "message ASTStickyBits {",
        "required bool begin_row_is_absolute = 1;",
        "required bool begin_column_is_absolute = 2;",
        "required bool end_row_is_absolute = 3;",
        "required bool end_column_is_absolute = 4;",
        "message ASTColonTractArchive {",
        "required int32 range_begin = 1;",
        "optional int32 range_end = 2;",
        "required uint32 range_begin = 1;",
        "optional uint32 range_end = 2;",
        "repeated .TSCE.ASTNodeArrayArchive.ASTColonTractArchive.ASTColonTractRelativeRangeArchive relative_column = 1;",
        "repeated .TSCE.ASTNodeArrayArchive.ASTColonTractArchive.ASTColonTractRelativeRangeArchive relative_row = 2;",
        "repeated .TSCE.ASTNodeArrayArchive.ASTColonTractArchive.ASTColonTractAbsoluteRangeArchive absolute_column = 3;",
        "repeated .TSCE.ASTNodeArrayArchive.ASTColonTractArchive.ASTColonTractAbsoluteRangeArchive absolute_row = 4;",
        "optional bool preserve_rectangular = 5 [default = true];",
        "message ASTLocalCellReferenceNodeArchive {",
        "required uint32 row_handle = 1;",
        "required uint32 column_handle = 2;",
        "message ASTColumnCoordinateArchive {",
        "required sint32 column = 1;",
        "message ASTRowCoordinateArchive {",
        "required sint32 row = 1;",
    ];
    const SYMBOLS: &[&str] = &[
        "pub struct DecodeOptions",
        "pub struct FormulaContext",
        "pub enum FormulaNode",
        "pub struct LocalPrecedent",
        "pub struct UnsupportedLocal",
        "pub trait FormulaVisitor",
        "pub trait FormulaDependencyVisitor",
        "pub struct DecodeReport",
        "pub fn inspect_formula_archive",
        "pub fn inspect_formula_dependencies_with_visitor",
        "pub fn decode_formula_archive_with_visitor",
        "pub enum FormulaWriteNode",
        "pub struct FormulaWritePlan",
        "pub struct FormulaWriteRequirements",
        "pub fn plan_formula_archive",
        "pub fn execute_formula_archive_plan",
        "preflight_callback_pass",
    ];
    let canonical = fs::read_to_string(proto_directory.join("TSCEArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TSCEFormulaArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    const EXPECTED_SCHEMA: &str = r#"syntax = "proto2";
package LitchiIwaFormulaProjection;
message FormulaArchive {
required bytes ast_node_array = 1;
optional uint32 host_column = 2;
optional uint32 host_row = 3;
optional bool host_column_is_negative = 4 [default = false];
optional bool host_row_is_negative = 5 [default = false];
optional bytes translation_flags = 6;
optional bytes host_table_uid = 7;
optional bytes host_column_uid = 8;
optional bytes host_row_uid = 9;
}
message ASTNodeArchive {
required int32 node_type = 1;
optional uint32 function_index = 2;
optional uint32 function_num_args = 3;
optional double number = 4;
optional bool boolean = 5;
optional string string = 6;
optional bool token_boolean = 10;
optional bytes thunk_array = 14;
optional bytes local_cell_reference = 15;
optional bytes cross_table_cell_reference = 16;
optional string whitespace = 25;
optional bytes column = 26;
optional bytes row = 27;
optional bytes cross_table_extra = 28;
optional bytes uid_coordinate = 30;
optional bytes sticky_bits = 33;
optional bytes tract_list = 38;
optional bytes colon_tract = 40;
optional uint64 decimal_low = 42;
optional uint64 decimal_high = 43;
}
message LocalCellReferenceArchive {
required uint32 row_handle = 1;
required uint32 column_handle = 2;
required uint32 row_is_sticky = 3;
required uint32 column_is_sticky = 4;
}
message ColumnCoordinateArchive {
required sint32 column = 1;
optional bool absolute = 2 [default = false];
}
message RowCoordinateArchive {
required sint32 row = 1;
optional bool absolute = 2 [default = false];
}
message CFUUIDArchive {
optional uint32 word0 = 2;
optional uint32 word1 = 3;
optional uint32 word2 = 4;
optional uint32 word3 = 5;
}
message CrossTableExtraArchive {
required bytes table_id = 1;
}
message StickyBitsArchive {
required bool begin_row_absolute = 1;
required bool begin_column_absolute = 2;
required bool end_row_absolute = 3;
required bool end_column_absolute = 4;
}
message RelativeRangeArchive {
required int32 begin = 1;
optional int32 end = 2;
}
message AbsoluteRangeArchive {
required uint32 begin = 1;
optional uint32 end = 2;
}"#;
    let codec = fs::read_to_string("src/numbers_formula_codec.rs")?;
    let production = production_codec_source(&codec);
    let lib = fs::read_to_string("src/lib.rs")?;
    if !CANONICAL.iter().all(|item| canonical.contains(item))
        || projection_schema != EXPECTED_SCHEMA
        || projection.contains("repeated ")
        || projection.len() > 4 * 1024
        || lib.matches("mod buffa_formula_generated {").count() != 1
        || lib.matches("pub mod numbers_formula_codec;").count() != 1
        || lib.contains("pub mod buffa_formula_generated")
        || !SYMBOLS.iter().all(|symbol| codec.contains(symbol))
        || production.contains("prost::")
        || production.contains("to_owned_message")
        || production.contains("encode_to_vec")
        || production.contains("try_encode")
        || production.contains("RepeatedView")
        || production.contains("LazyRepeatedView")
    {
        return Err("FormulaArchive projection/codec drifted from canonical TSCE fields, exposed generated repeated storage, or introduced owned/generated decoding".into());
    }
    Ok(())
}

fn enforce_keynote_show_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const TSP_SIZE: &str =
        "message Size {\n  required float width = 1;\n  required float height = 2;\n}";
    const KN_SLIDE_TREE: &str = "message SlideTreeArchive {\n  optional .TSP.Reference rootSlideNode = 1 [deprecated = true];\n  repeated .TSP.Reference slides = 2;\n}";
    const KN_SHOW: &str = "message ShowArchive {\n  enum KNShowMode {\n    kKNShowModeNormal = 0;\n    kKNShowModeAutoPlay = 1;\n    kKNShowModeHyperlinksOnly = 2;\n  }\n  optional .TSP.Reference uiState = 1;\n  required .TSP.Reference theme = 2;\n  required .KN.SlideTreeArchive slideTree = 3;\n  required .TSP.Size size = 4;\n  required .TSP.Reference stylesheet = 5;\n  optional bool slideNumbersVisible = 6;\n  optional .TSP.Reference recording = 7;\n  optional bool loop_presentation = 8;\n  optional .KN.ShowArchive.KNShowMode mode = 9 [default = kKNShowModeNormal];\n  optional double autoplay_transition_delay = 10 [default = 5];\n  optional double autoplay_build_delay = 11 [default = 2];\n  optional bool idle_timer_active = 15;\n  optional double idle_timer_delay = 16 [default = 900];\n  optional .TSP.Reference soundtrack = 17;\n  optional bool automatically_plays_upon_open = 18;\n  optional .TSP.Reference slideList = 19;\n}";
    const PROJECTION_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const PROJECTION_SIZE: &str =
        "message Size {\n  required float width = 1;\n  required float height = 2;\n}";
    const PROJECTION_SHOW: &str = "message KeynoteShowArchive {\n  optional .LitchiIwaProjection.Reference ui_state = 1;\n  required .LitchiIwaProjection.Reference theme = 2;\n  required .LitchiIwaProjection.Size size = 4;\n  required .LitchiIwaProjection.Reference stylesheet = 5;\n  optional bool slide_numbers_visible = 6;\n  optional .LitchiIwaProjection.Reference recording = 7;\n  optional bool loop_presentation = 8;\n  optional int32 mode = 9 [default = 0];\n  optional double autoplay_transition_delay = 10 [default = 5];\n  optional double autoplay_build_delay = 11 [default = 2];\n  optional bool idle_timer_active = 15;\n  optional double idle_timer_delay = 16 [default = 900];\n  optional .LitchiIwaProjection.Reference soundtrack = 17;\n  optional bool automatically_plays_upon_open = 18;\n  optional .LitchiIwaProjection.Reference slide_list = 19;\n}";
    const PRIVATE_MODULE_DECLARATIONS: [&str; 2] = [
        "#[doc(hidden)]\nmod buffa_keynote_show_generated {",
        "\"/buffa-keynote-show/iwa_keynote_show_buffa_protos.rs\"",
    ];
    const ROUTER_DECLARATIONS: [&str; 23] = [
        "const SHOW_UI_STATE_FIELD: u32 = 1;",
        "const SHOW_THEME_FIELD: u32 = 2;",
        "const SHOW_SLIDE_TREE_FIELD: u32 = 3;",
        "const SHOW_SIZE_FIELD: u32 = 4;",
        "const SHOW_STYLESHEET_FIELD: u32 = 5;",
        "const SHOW_SLIDE_NUMBERS_VISIBLE_FIELD: u32 = 6;",
        "const SHOW_RECORDING_FIELD: u32 = 7;",
        "const SHOW_LOOP_PRESENTATION_FIELD: u32 = 8;",
        "const SHOW_MODE_FIELD: u32 = 9;",
        "const SHOW_AUTOPLAY_TRANSITION_DELAY_FIELD: u32 = 10;",
        "const SHOW_AUTOPLAY_BUILD_DELAY_FIELD: u32 = 11;",
        "const SHOW_IDLE_TIMER_ACTIVE_FIELD: u32 = 15;",
        "const SHOW_IDLE_TIMER_DELAY_FIELD: u32 = 16;",
        "const SHOW_SOUNDTRACK_FIELD: u32 = 17;",
        "const SHOW_AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD: u32 = 18;",
        "const SHOW_SLIDE_LIST_FIELD: u32 = 19;",
        "const SLIDE_TREE_ROOT_FIELD: u32 = 1;",
        "const SLIDE_TREE_SLIDES_FIELD: u32 = 2;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;",
        "const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;",
        "const SIZE_WIDTH_FIELD: u32 = 1;",
        "const SIZE_HEIGHT_FIELD: u32 = 2;",
    ];

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let keynote = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("KNShowArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let expected_projection_schema = [
        "syntax = \"proto2\";",
        "package LitchiIwaProjection;",
        PROJECTION_REFERENCE,
        PROJECTION_SIZE,
        PROJECTION_SHOW,
    ]
    .join("\n")
    .lines()
    .map(str::trim)
    .collect::<Vec<_>>()
    .join("\n");
    let router = fs::read_to_string("src/keynote_show_codec.rs")?;
    let production_router = production_codec_source(&router);
    let lib = fs::read_to_string("src/lib.rs")?;
    if tsp.matches(TSP_REFERENCE).count() != 1
        || tsp.matches(TSP_SIZE).count() != 1
        || keynote.matches(KN_SLIDE_TREE).count() != 1
        || keynote.matches(KN_SHOW).count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_SIZE).count() != 1
        || projection.matches(PROJECTION_SHOW).count() != 1
        || projection_schema != expected_projection_schema
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| router.matches(declaration).count() == 1)
        || !PRIVATE_MODULE_DECLARATIONS
            .iter()
            .all(|declaration| lib.matches(declaration).count() == 1)
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || production_router.contains("to_owned_message")
        || production_router.contains("encode_to_vec")
        || production_router.contains("try_encode")
        || production_router.contains(".encode(")
    {
        return Err(
            "derived Keynote show projection/router drifted from canonical fields, exceeded its 2 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_pages_section_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const CANONICAL_FIELDS: [&str; 12] = [
        "optional bool inherit_previous_header_footer = 17;",
        "optional bool section_template_first_page_different = 18;",
        "optional bool section_template_even_odd_pages_different = 19;",
        "optional uint32 section_start_kind = 20;",
        "optional uint32 section_page_number_kind = 21;",
        "optional uint32 section_page_number_start = 22;",
        "optional .TSP.Reference first_section_template_page = 23;",
        "optional .TSP.Reference even_section_template_page = 24;",
        "optional .TSP.Reference odd_section_template_page = 25;",
        "optional string name = 26;",
        "optional bool section_template_first_page_hides_header_footer = 28;",
        "optional .TSP.Reference user_defined_guide_storage = 29;",
    ];
    const PAGINATION_PROJECTION: &str = "message PagesSectionPaginationArchive {\n  optional uint32 section_start_kind = 20;\n  optional uint32 section_page_number_kind = 21;\n  optional uint32 section_page_number_start = 22;\n}";
    const SETTINGS_PROJECTION: &str = "message PagesSectionSettingsArchive {\n  optional bool inherit_previous_header_footer = 17;\n  optional bool section_template_first_page_different = 18;\n  optional bool section_template_even_odd_pages_different = 19;\n  optional uint32 section_start_kind = 20;\n  optional uint32 section_page_number_kind = 21;\n  optional uint32 section_page_number_start = 22;\n  optional .LitchiIwaProjection.Reference first_section_template_page = 23;\n  optional .LitchiIwaProjection.Reference even_section_template_page = 24;\n  optional .LitchiIwaProjection.Reference odd_section_template_page = 25;\n  optional string name = 26;\n  optional bool section_template_first_page_hides_header_footer = 28;\n  optional .LitchiIwaProjection.Reference user_defined_guide_storage = 29;\n}";
    const REFERENCE_PROJECTION: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const ROUTER_DECLARATIONS: [&str; 17] = [
        "const INHERIT_HEADER_FOOTER_FIELD: u32 = 17;",
        "const FIRST_PAGE_DIFFERENT_FIELD: u32 = 18;",
        "const EVEN_ODD_PAGES_DIFFERENT_FIELD: u32 = 19;",
        "const SECTION_START_FIELD: u32 = 20;",
        "const PAGE_NUMBERING_FIELD: u32 = 21;",
        "const STARTING_PAGE_NUMBER_FIELD: u32 = 22;",
        "const FIRST_TEMPLATE_FIELD: u32 = 23;",
        "const EVEN_TEMPLATE_FIELD: u32 = 24;",
        "const ODD_TEMPLATE_FIELD: u32 = 25;",
        "const SECTION_NAME_FIELD: u32 = 26;",
        "const FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD: u32 = 28;",
        "const GUIDE_STORAGE_FIELD: u32 = 29;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;",
        "const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;",
        "const MAX_RECURSION: u32 = 64;",
        "const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;",
    ];

    let pages = fs::read_to_string(proto_directory.join("TPArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TPSectionArchive.proto"))?;
    let codec = fs::read_to_string("src/pages_section_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if !CANONICAL_FIELDS
        .iter()
        .all(|declaration| pages.matches(declaration).count() == 1)
        || projection.matches(PAGINATION_PROJECTION).count() != 1
        || projection.matches(SETTINGS_PROJECTION).count() != 1
        || projection.matches(REFERENCE_PROJECTION).count() != 1
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| codec.matches(declaration).count() == 1)
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || production_codec.contains("prost")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Pages section projections drifted from TP.SectionArchive fields 17--29 (excluding 27), exceeded their 2 KiB source budget, introduced generated repeated storage, or added Prost/generated production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_pages_section_background_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_PROJECTION_DIGEST: &str =
        "0a6f03a7046c285e431953b8752096a1f0117206724b561da294c64092aa9cfc";
    const SECTION_FIELD: &str = "optional .TSD.FillArchive background_fill = 30;";
    const FILL_FIELDS: [&str; 3] = [
        "optional .TSP.Color color = 1;",
        "optional .TSD.GradientArchive gradient = 2;",
        "optional .TSD.ImageFillArchive image = 3;",
    ];
    const COLOR_FIELDS: [&str; 6] = [
        "required .TSP.Color.ColorModel model = 1;",
        "optional float r = 3;",
        "optional float g = 4;",
        "optional float b = 5;",
        "optional float a = 6 [default = 1];",
        "optional .TSP.Color.RGBColorSpace rgbspace = 12;",
    ];
    const CODEC_MARKERS: [&str; 4] = [
        "const SECTION_BACKGROUND_FIELD: u32 = 30;",
        "const FILL_COLOR_FIELD: u32 = 1;",
        "const COLOR_MODEL_FIELD: u32 = 1;",
        "const COLOR_SPACE_FIELD: u32 = 12;",
    ];

    let pages = fs::read_to_string(proto_directory.join("TPArchives.proto"))?;
    let drawing = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let common = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TPSectionBackgroundArchive.proto"))?;
    let projection_digest = Sha256::digest(projection.as_bytes())
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let codec = fs::read_to_string("src/pages_section_background_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if pages.matches(SECTION_FIELD).count() != 1
        || !FILL_FIELDS.iter().all(|field| drawing.contains(field))
        || !COLOR_FIELDS
            .iter()
            .all(|field| common.matches(field).count() == 1)
        || !CODEC_MARKERS
            .iter()
            .all(|marker| codec.matches(marker).count() == 1)
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || projection_digest != EXPECTED_PROJECTION_DIGEST
        || production_codec.contains("prost::")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err("Pages section-background projection/codec drifted from TP field 30, TSD Fill fields 1--3, or TSP Color fields 1/3--6/12; exceeded its source budget, introduced repeated storage, Prost, or generated production encoding".into());
    }
    Ok(())
}

fn enforce_pages_body_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const TP_FIELDS: [&str; 15] = [
        "required .TSA.DocumentArchive super = 15;",
        "optional .TSP.Reference body_storage = 4;",
        "optional .TSP.Reference section = 5;",
        "optional .TSP.Reference settings = 7;",
        "optional float page_width = 30;",
        "optional float page_height = 31;",
        "optional float left_margin = 32;",
        "optional float right_margin = 33;",
        "optional float top_margin = 34;",
        "optional float bottom_margin = 35;",
        "optional float header_margin = 36;",
        "optional float footer_margin = 37;",
        "optional float page_scale = 38;",
        "optional bool lays_out_body_vertically = 39;",
        "optional uint32 orientation = 42 [default = 0];",
    ];
    const TP_SETTINGS_FIELDS: [&str; 10] = [
        "optional bool body = 1 [default = true];",
        "optional bool headers = 2 [default = true];",
        "optional bool footers = 3 [default = true];",
        "optional bool hyphenation = 9 [default = false];",
        "optional bool use_ligatures = 10 [default = false];",
        "optional .TP.SettingsArchive.FootnoteKind footnote_kind = 30;",
        "optional .TP.SettingsArchive.FootnoteFormat footnote_format = 31;",
        "optional .TP.SettingsArchive.FootnoteNumbering footnote_numbering = 32;",
        "optional int32 footnote_gap = 33;",
        "optional bool facing_pages = 34 [default = false];",
    ];
    const TSWP_BOUNDARY: &str = "message ObjectAttribute {\n    required uint32 character_index = 1;\n    optional .TSP.Reference object = 2;\n  }";
    const PROJECTION_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const PROJECTION_DOCUMENT: &str = "message PagesDocumentBodyArchive {\n  optional .LitchiIwaProjection.Reference body_storage = 4;\n  optional .LitchiIwaProjection.Reference initial_section = 5;\n  optional .LitchiIwaProjection.Reference settings = 7;\n  optional float page_width = 30;\n  optional float page_height = 31;\n  optional float left_margin = 32;\n  optional float right_margin = 33;\n  optional float top_margin = 34;\n  optional float bottom_margin = 35;\n  optional float header_margin = 36;\n  optional float footer_margin = 37;\n  optional float page_scale = 38;\n  optional bool lays_out_body_vertically = 39;\n  optional uint32 orientation = 42 [default = 0];\n}";
    const PROJECTION_SETTINGS: &str = "message PagesSettingsArchive {\n  optional bool body = 1 [default = true];\n  optional bool headers = 2 [default = true];\n  optional bool footers = 3 [default = true];\n  optional bool hyphenation = 9 [default = false];\n  optional bool use_ligatures = 10 [default = false];\n  optional int32 footnote_kind = 30;\n  optional int32 footnote_format = 31;\n  optional int32 footnote_numbering = 32;\n  optional int32 footnote_gap = 33;\n  optional bool facing_pages = 34 [default = false];\n}";
    const PROJECTION_BOUNDARY: &str = "message PagesSectionBoundaryEntry {\n  required uint32 character_index = 1;\n  optional .LitchiIwaProjection.Reference section = 2;\n}";
    const BODY_ROUTER_DECLARATIONS: [&str; 9] = [
        "const DOCUMENT_BODY_STORAGE_FIELD: u32 = 4;",
        "const DOCUMENT_INITIAL_SECTION_FIELD: u32 = 5;",
        "const DOCUMENT_SUPER_FIELD: u32 = 15;",
        "const BOUNDARY_CHARACTER_INDEX_FIELD: u32 = 1;",
        "const BOUNDARY_SECTION_FIELD: u32 = 2;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;",
        "const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;",
        "const MAX_RECURSION_LIMIT: u32 = 64;",
    ];
    const LAYOUT_ROUTER_DECLARATIONS: [&str; 15] = [
        "const SUPER: u32 = 15;",
        "const BODY_STORAGE: u32 = 4;",
        "const INITIAL_SECTION: u32 = 5;",
        "const WIDTH: u32 = 30;",
        "const HEIGHT: u32 = 31;",
        "const LEFT: u32 = 32;",
        "const RIGHT: u32 = 33;",
        "const TOP: u32 = 34;",
        "const BOTTOM: u32 = 35;",
        "const HEADER: u32 = 36;",
        "const FOOTER: u32 = 37;",
        "const SCALE: u32 = 38;",
        "const VERTICAL: u32 = 39;",
        "const ORIENTATION: u32 = 42;",
        "const MAX_RECURSION: u32 = 64;",
    ];
    const SETTINGS_ROUTER_DECLARATIONS: [&str; 18] = [
        "const ROOT_SUPER: u32 = 15;",
        "const ROOT_SETTINGS: u32 = 7;",
        "const ROOT_BODY_STORAGE: u32 = 4;",
        "const ROOT_INITIAL_SECTION: u32 = 5;",
        "const REFERENCE_IDENTIFIER: u32 = 1;",
        "const REFERENCE_TYPE: u32 = 2;",
        "const REFERENCE_EXTERNAL: u32 = 3;",
        "const SETTINGS_BODY: u32 = 1;",
        "const SETTINGS_HEADERS: u32 = 2;",
        "const SETTINGS_FOOTERS: u32 = 3;",
        "const SETTINGS_HYPHENATION: u32 = 9;",
        "const SETTINGS_USE_LIGATURES: u32 = 10;",
        "const SETTINGS_FOOTNOTE_KIND: u32 = 30;",
        "const SETTINGS_FOOTNOTE_FORMAT: u32 = 31;",
        "const SETTINGS_FOOTNOTE_NUMBERING: u32 = 32;",
        "const SETTINGS_FOOTNOTE_GAP: u32 = 33;",
        "const SETTINGS_FACING_PAGES: u32 = 34;",
        "const MAX_RECURSION: u32 = 64;",
    ];

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let pages = fs::read_to_string(proto_directory.join("TPArchives.proto"))?;
    let text = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TPDocumentBodyArchive.proto"))?;
    let codec = fs::read_to_string("src/pages_body_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let settings_codec = fs::read_to_string("src/pages_document_settings_codec.rs")?;
    let production_settings_codec = production_codec_source(&settings_codec);
    let layout_codec = fs::read_to_string("src/pages_page_layout_codec.rs")?;
    let production_layout_codec = production_codec_source(&layout_codec);
    if tsp.matches(TSP_REFERENCE).count() != 1
        || !TP_FIELDS
            .iter()
            .all(|declaration| pages.matches(declaration).count() == 1)
        || !TP_SETTINGS_FIELDS
            .iter()
            .all(|declaration| pages.matches(declaration).count() == 1)
        || text.matches(TSWP_BOUNDARY).count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_DOCUMENT).count() != 1
        || projection.matches(PROJECTION_SETTINGS).count() != 1
        || projection.matches(PROJECTION_BOUNDARY).count() != 1
        || !BODY_ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| codec.matches(declaration).count() == 1)
        || !LAYOUT_ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| layout_codec.matches(declaration).count() == 1)
        || !SETTINGS_ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| settings_codec.matches(declaration).count() == 1)
        || projection.len() > 3 * 1024
        || projection.contains("repeated ")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
        || production_settings_codec.contains("to_owned_message")
        || production_settings_codec.contains("encode_to_vec")
        || production_settings_codec.contains("try_encode")
        || production_settings_codec.contains(".encode(")
        || production_layout_codec.contains("to_owned_message")
        || production_layout_codec.contains("encode_to_vec")
        || production_layout_codec.contains("try_encode")
        || production_layout_codec.contains(".encode(")
    {
        return Err(
            "derived Pages body/layout/settings projection or codec drifted from canonical TP/TSWP/TSP fields, exceeded its 3 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_pages_media_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const MOVIE_AUDIO_ONLY: &str = "optional bool audioOnly = 9;";
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\npackage LitchiIwaProjection;\nmessage MovieAudioFlagArchive {\noptional bool audio_only = 9;\n}";
    const CODEC_MARKERS: [&str; 7] = [
        "const AUDIO_ONLY_FIELD: u32 = 9;",
        "const MAX_RECURSION_LIMIT: u32 = 64;",
        "pub struct MovieAudioFlagSnapshot",
        "pub fn decode_movie_audio_only(",
        "decode_lazy_view",
        "crate::buffa_pages_media_generated::",
        "Unknown source fields are never materialized or",
    ];

    let tsd = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSDMovieAudioFlagArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/pages_media_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if tsd.matches(MOVIE_AUDIO_ONLY).count() != 1
        || projection_schema != PROJECTION_SCHEMA
        || projection.len() > 1024
        || projection.contains("repeated ")
        || !CODEC_MARKERS
            .iter()
            .all(|marker| production_codec.matches(marker).count() == 1)
        || FORBIDDEN_PROST_CODEC_MARKERS
            .iter()
            .any(|fragment| production_codec.contains(fragment))
    {
        return Err("Pages media projection/codec drifted from TSD.MovieArchive.audioOnly, exceeded its source budget, introduced repeated storage, or added production Prost/encoding".into());
    }
    Ok(())
}

fn enforce_pages_native_message_provenance(proto_directory: &Path) -> Result<(), Box<dyn Error>> {
    // Native IWA object type numbers are not part of the protobuf schemas.
    // Keep their workspace routes private and check them only as a build-time
    // provenance seam. The canonical message blocks below are the authority
    // for what each number means; no ID is added to the public codec API.
    const ROUTE_DECLARATIONS: [(&str, &str, &str, &str); 21] = [
        (
            "../litchi-iwa/src/pages/editor.rs",
            "const MOVIE_MESSAGE_TYPE: u32 = 3_007;",
            "../litchi-iwa/src/pages/editor.rs",
            "MOVIE_MESSAGE_TYPE => remap_pages_movie_wire",
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            "const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;",
            "../litchi-iwa/src/pages/editor.rs",
            "CAPTION_INFO_MESSAGE_TYPE => remap_pages_caption_info_wire",
        ),
        (
            "../litchi-iwa/src/pages/editor/movies/graph.rs",
            "const MOVIE_MESSAGE_TYPE: u32 = 3_007;",
            "../litchi-iwa/src/pages/editor/movies/graph.rs",
            "decode_typed_package_object(package, identifier, MOVIE_MESSAGE_TYPE",
        ),
        (
            "../litchi-iwa/src/pages/editor/movies/caption.rs",
            "const MOVIE_MESSAGE_TYPE: u32 = 3_007;",
            "../litchi-iwa/src/pages/editor/movies/caption.rs",
            "MOVIE_MESSAGE_TYPE,\n        \"TSD.MovieArchive\",",
        ),
        (
            "../litchi-iwa/src/pages/editor/audio/graph.rs",
            "const AUDIO_MESSAGE_TYPE: u32 = 3_007;",
            "../litchi-iwa/src/pages/editor/audio/graph.rs",
            "decode_typed_package_object(\n        editor.package(),\n        drawable_object_id,\n        AUDIO_MESSAGE_TYPE,\n",
        ),
        (
            "../litchi-iwa/src/pages/editor/body_shapes/caption.rs",
            "const THEME_MESSAGE_TYPE: u32 = 10_001;",
            "../litchi-iwa/src/pages/editor/body_shapes/caption.rs",
            ".filter(|message| message.type_ == THEME_MESSAGE_TYPE)",
        ),
        (
            "../litchi-iwa/src/image_caption.rs",
            "pub(crate) const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;",
            "../litchi-iwa/src/image_caption.rs",
            ".filter(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)",
        ),
        (
            "../litchi-iwa/src/pages/editor/footnotes.rs",
            "const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;",
            "../litchi-iwa/src/pages/editor/footnotes.rs",
            "pages_footnote_codec::decode_footnote_reference(",
        ),
        (
            "../litchi-iwa/src/pages/editor/footnotes.rs",
            "const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;",
            "../litchi-iwa/src/pages/editor/footnotes.rs",
            "pages_footnote_marker_codec::decode_textual_attachment(",
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            "const DOCUMENT_MESSAGE_TYPE: u32 = 10000;",
            "../litchi-iwa/src/pages/editor.rs",
            ".find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)\n        .map(|message| message.data.as_slice())",
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            "const SECTION_MESSAGE_TYPE: u32 = 10011;",
            "../litchi-iwa/src/pages/editor.rs",
            "SECTION_MESSAGE_TYPE,\n        \"TP.SectionArchive\",",
        ),
        (
            "../litchi-iwa/src/pages/creation.rs",
            "enum PagesMessageType {",
            "../litchi-iwa/src/pages/creation.rs",
            "PagesMessageType::Document,\n            document,",
        ),
        (
            "../litchi-pages/src/package.rs",
            "const SECTION_MESSAGE_TYPE: u32 = 10_011;",
            "../litchi-pages/src/package.rs",
            "unique_message_payload(&object.messages, SECTION_MESSAGE_TYPE",
        ),
        (
            "../litchi-pages/src/package.rs",
            "const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;",
            "../litchi-pages/src/package.rs",
            "FOOTNOTE_REFERENCE_MESSAGE_TYPE,\n        &format!(\"Pages footnote reference object",
        ),
        (
            "../litchi-pages/src/package.rs",
            "const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;",
            "../litchi-pages/src/package.rs",
            "TEXTUAL_ATTACHMENT_MESSAGE_TYPE,\n        &format!(\"Pages footnote marker object",
        ),
        (
            "../litchi-pages/src/package/document_settings.rs",
            "const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;",
            "../litchi-pages/src/package/document_settings.rs",
            "page_layout::unique_message(root, DOCUMENT_MESSAGE_TYPE)",
        ),
        (
            "../litchi-pages/src/package/document_settings.rs",
            "const SETTINGS_MESSAGE_TYPE: u32 = 10_012;",
            "../litchi-pages/src/package/document_settings.rs",
            "RawMessage {\n                    type_: SETTINGS_MESSAGE_TYPE,",
        ),
        (
            "../litchi-pages/src/package/page_layout.rs",
            "const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;",
            "../litchi-pages/src/package/page_layout.rs",
            "unique_message(document_object, DOCUMENT_MESSAGE_TYPE)?",
        ),
        (
            "../litchi-pages/src/package/section_transaction.rs",
            "pub(super) const TEMPLATE_MESSAGE_TYPE: u32 = 10_143;",
            "../litchi-pages/src/package/section_settings.rs",
            "transaction::unique_message(object, transaction::TEMPLATE_MESSAGE_TYPE, path)?",
        ),
        (
            "../litchi-pages/src/package/section_transaction.rs",
            "pub(super) const STORAGE_MESSAGE_TYPES: [u32; 2] = [2_001, 2_022];",
            "../litchi-pages/src/package/section_transaction.rs",
            "STORAGE_MESSAGE_TYPES.contains(&message.type_)",
        ),
        (
            "../litchi-pages/src/package/table_lock.rs",
            "const ROOT_MESSAGE_TYPE: u32 = 10_000;",
            "../litchi-pages/src/package/table_lock.rs",
            "ROOT_MESSAGE_TYPE,\n        budget,",
        ),
    ];
    // Keep the additional live consumers tied to their exact production
    // decode/remap/write sites. The declaration table above proves that the
    // private numeric routes exist; these markers prove that each route is
    // still used by the intended code path rather than a comment, string, or
    // test-only compatibility shim.
    const PRODUCTION_ROUTE_MARKERS: [(&str, &str, usize); 40] = [
        (
            "../litchi-iwa/src/pages/editor/movies/graph.rs",
            "decode_typed_package_object(\n        editor.package(),\n        drawable_object_id,\n        MOVIE_MESSAGE_TYPE,\n",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/movies/graph.rs",
            ".filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/movies/graph.rs",
            "pages_movie_object(\n            ids.drawable,\n            MOVIE_MESSAGE_TYPE,\n            movie,",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/movies/graph.rs",
            ".filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)",
            2,
        ),
        (
            "../litchi-iwa/src/pages/editor/movies/graph.rs",
            "type_: MOVIE_MESSAGE_TYPE,\n                data,",
            2,
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            ".find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)\n        .and_then(|message| DocumentArchive::decode(message.data.as_slice()).ok())",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            "SECTION_MESSAGE_TYPE => {\n                        if !section_objects.insert(identifier)",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            ".any(|message| message.type_ == SECTION_MESSAGE_TYPE)",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            "SECTION_MESSAGE_TYPE => {\n            const REFERENCE_PATHS:",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor.rs",
            "let mut cloned = clone_pages_object_metadata(\n        source,\n        new_identifier,\n        vec![RawMessage {\n            type_: message.type_,\n            data,\n        }],",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/audio/graph.rs",
            ".filter(|message| message.type_ == AUDIO_MESSAGE_TYPE)",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/audio/graph.rs",
            ".filter(|(_, message)| message.type_ == AUDIO_MESSAGE_TYPE)",
            2,
        ),
        (
            "../litchi-iwa/src/pages/editor/audio/graph.rs",
            "decode_typed_package_object(package, identifier, AUDIO_MESSAGE_TYPE, \"TSD.MovieArchive\")?",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/audio/graph.rs",
            "pages_audio_object(\n            ids.drawable,\n            AUDIO_MESSAGE_TYPE,\n            audio,",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/audio/graph.rs",
            "type_: AUDIO_MESSAGE_TYPE,\n                data,",
            2,
        ),
        (
            "../litchi-iwa/src/pages/creation.rs",
            "Document = 10_000,",
            1,
        ),
        (
            "../litchi-iwa/src/pages/creation.rs",
            "Section = 10_011,",
            1,
        ),
        (
            "../litchi-iwa/src/pages/creation.rs",
            "Settings = 10_012,",
            1,
        ),
        (
            "../litchi-iwa/src/pages/creation.rs",
            "PagesMessageType::Section,\n            tp::SectionArchive",
            1,
        ),
        (
            "../litchi-iwa/src/pages/creation.rs",
            "PagesMessageType::Settings,\n            tp::SettingsArchive",
            1,
        ),
        (
            "../litchi-iwa/src/pages/creation.rs",
            "type_: message_type.value(),\n            data,",
            1,
        ),
        (
            "../litchi-pages/src/package.rs",
            "let payload = unique_message_payload(&object.messages, 10_000, \"Pages root object 1\")?",
            1,
        ),
        (
            "../litchi-pages/src/package/footnote_text.rs",
            "let root = root_references_with_limits(components, package.state.source.limits())",
            1,
        ),
        (
            "../litchi-pages/src/package/footnote_text.rs",
            "let payload = unique_message_payload(\n            &reference.messages,\n            FOOTNOTE_REFERENCE_MESSAGE_TYPE,",
            1,
        ),
        (
            "../litchi-pages/src/package/footnote_text.rs",
            "let marker_payload = unique_message_payload(\n            &marker.messages,\n            TEXTUAL_ATTACHMENT_MESSAGE_TYPE,",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/body_shapes/caption.rs",
            ".filter(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)",
            1,
        ),
        (
            "../litchi-iwa/src/pages/editor/body_shapes/caption.rs",
            "pages_movie_caption_codec::decode_caption_info(",
            1,
        ),
        (
            "../litchi-pages/src/package/section_pagination.rs",
            ".filter(|(_index, message)| message.type_ == SECTION_MESSAGE_TYPE)",
            1,
        ),
        (
            "../litchi-pages/src/package/section_pagination.rs",
            "type_: SECTION_MESSAGE_TYPE,\n                data: rewritten,",
            1,
        ),
        (
            "../litchi-pages/src/package/section_pagination.rs",
            ".filter(|message| message.type_ == SECTION_MESSAGE_TYPE)",
            1,
        ),
        (
            "../litchi-pages/src/package/section_settings.rs",
            "transaction::resolve_target(source, position, &mut budget)?",
            1,
        ),
        (
            "../litchi-pages/src/package/section_settings.rs",
            "transaction::unique_message(object, transaction::TEMPLATE_MESSAGE_TYPE, path)?",
            1,
        ),
        (
            "../litchi-pages/src/package/section_settings.rs",
            "transaction::STORAGE_MESSAGE_TYPES.contains(&message.type_)",
            1,
        ),
        (
            "../litchi-pages/src/package/section_background.rs",
            "transaction::resolve_target(source, position, &mut budget).map_err(map_transaction)?",
            1,
        ),
        (
            "../litchi-pages/src/package/section_background.rs",
            "transaction::resolve_target(package, position, budget).map_err(map_transaction)?",
            1,
        ),
        (
            "../litchi-pages/src/package/section_text.rs",
            "transaction::resolve_body_target(source, position, &mut budget)",
            1,
        ),
        (
            "../litchi-pages/src/package/footnote_text.rs",
            "let body =\n        find_object(components, body_identifier.get())",
            1,
        ),
        (
            "../litchi-pages/src/package/footnote_text.rs",
            ".object_mut(native_footnote.reference_identifier.get())",
            1,
        ),
        (
            "../litchi-pages/src/package/footnote_text.rs",
            ".object_mut(native_footnote.storage_identifier.get())",
            1,
        ),
        (
            "../litchi-pages/src/package/footnote_text.rs",
            "let marker_identifier = footnote_marker_identifier(",
            1,
        ),
    ];
    const REGISTRY_DECLARATIONS: [&str; 5] = [
        "3007u32 => decode_shape_archive,",
        "2004u32 => decode_storage_archive,",
        "2008u32 => decode_storage_archive,",
        "2001u32 => decode_storage_archive,",
        "2022u32 => decode_storage_archive,",
    ];
    const MOVIE_FIELDS: [&str; 2] = [
        "required .TSD.DrawableArchive super = 1;",
        "optional bool audioOnly = 9;",
    ];
    const CAPTION_FIELDS: [&str; 3] = [
        "required .TSWP.ShapeInfoArchive super = 1;",
        "optional .TSP.Reference placement = 2;",
        "optional .TSD.CaptionOrTitleKind childInfoKind = 3;",
    ];
    const FOOTNOTE_FIELDS: [&str; 3] = [
        "optional .TSWP.TextualAttachmentArchive super = 1;",
        "optional .TSP.Reference contained_storage = 2;",
        "optional string custom_mark_string = 3;",
    ];
    const TEXTUAL_FIELDS: [&str; 2] = [
        "optional string string_equivalent = 1;",
        "optional .TSWP.TextualAttachmentArchive.Kind kind = 2;",
    ];
    const DOCUMENT_FIELDS: [&str; 17] = [
        "required .TSA.DocumentArchive super = 15;",
        "optional .TSP.Reference body_storage = 4;",
        "optional .TSP.Reference section = 5;",
        "optional .TSP.Reference settings = 7;",
        "optional .TSP.Reference deprecated_layout_state = 11;",
        "optional .TSP.Reference deprecated_view_state = 12;",
        "optional float page_width = 30;",
        "optional float page_height = 31;",
        "optional float left_margin = 32;",
        "optional float right_margin = 33;",
        "optional float top_margin = 34;",
        "optional float bottom_margin = 35;",
        "optional float header_margin = 36;",
        "optional float footer_margin = 37;",
        "optional float page_scale = 38;",
        "optional bool lays_out_body_vertically = 39;",
        "optional uint32 orientation = 42 [default = 0];",
    ];
    const SETTINGS_FIELDS: [&str; 10] = [
        "optional bool body = 1 [default = true];",
        "optional bool headers = 2 [default = true];",
        "optional bool footers = 3 [default = true];",
        "optional bool hyphenation = 9 [default = false];",
        "optional bool use_ligatures = 10 [default = false];",
        "optional .TP.SettingsArchive.FootnoteKind footnote_kind = 30;",
        "optional .TP.SettingsArchive.FootnoteFormat footnote_format = 31;",
        "optional .TP.SettingsArchive.FootnoteNumbering footnote_numbering = 32;",
        "optional int32 footnote_gap = 33;",
        "optional bool facing_pages = 34 [default = false];",
    ];
    const SECTION_FIELDS: [&str; 12] = [
        "optional bool inherit_previous_header_footer = 17;",
        "optional bool section_template_first_page_different = 18;",
        "optional bool section_template_even_odd_pages_different = 19;",
        "optional uint32 section_start_kind = 20;",
        "optional uint32 section_page_number_kind = 21;",
        "optional uint32 section_page_number_start = 22;",
        "optional .TSP.Reference first_section_template_page = 23;",
        "optional .TSP.Reference even_section_template_page = 24;",
        "optional .TSP.Reference odd_section_template_page = 25;",
        "optional string name = 26;",
        "optional bool section_template_first_page_hides_header_footer = 28;",
        "optional .TSP.Reference user_defined_guide_storage = 29;",
    ];

    let tsd = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let tsa = fs::read_to_string(proto_directory.join("TSAArchives.proto"))?;
    let tswp = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let tp = fs::read_to_string(proto_directory.join("TPArchives.proto"))?;
    let Some(movie) = proto_message_block(&tsd, "MovieArchive") else {
        return Err("Pages native-ID provenance lost TSD.MovieArchive".into());
    };
    let Some(caption) = proto_message_block(&tsa, "CaptionInfoArchive") else {
        return Err("Pages native-ID provenance lost TSA.CaptionInfoArchive".into());
    };
    let Some(footnote) = proto_message_block(&tswp, "FootnoteReferenceAttachmentArchive") else {
        return Err(
            "Pages native-ID provenance lost TSWP.FootnoteReferenceAttachmentArchive".into(),
        );
    };
    let Some(textual) = proto_message_block(&tswp, "TextualAttachmentArchive") else {
        return Err("Pages native-ID provenance lost TSWP.TextualAttachmentArchive".into());
    };
    let Some(document) = proto_message_block(&tp, "DocumentArchive") else {
        return Err("Pages native-ID provenance lost TP.DocumentArchive".into());
    };
    let Some(settings) = proto_message_block(&tp, "SettingsArchive") else {
        return Err("Pages native-ID provenance lost TP.SettingsArchive".into());
    };
    let Some(section) = proto_message_block(&tp, "SectionArchive") else {
        return Err("Pages native-ID provenance lost TP.SectionArchive".into());
    };

    let canonical_scope_ok = MOVIE_FIELDS
        .iter()
        .all(|field| proto_field(movie, field) == 1)
        && CAPTION_FIELDS
            .iter()
            .all(|field| proto_field(caption, field) == 1)
        && FOOTNOTE_FIELDS
            .iter()
            .all(|field| proto_field(footnote, field) == 1)
        && TEXTUAL_FIELDS
            .iter()
            .all(|field| proto_field(textual, field) == 1)
        && DOCUMENT_FIELDS
            .iter()
            .all(|field| proto_field(document, field) == 1)
        && SETTINGS_FIELDS
            .iter()
            .all(|field| proto_field(settings, field) == 1)
        && SECTION_FIELDS
            .iter()
            .all(|field| proto_field(section, field) == 1);

    // A standalone publication of litchi-iwa-protos has no sibling editor
    // sources. In that layout the schema-side guard still runs, while the
    // workspace-only numeric route check is intentionally skipped. A partial
    // sibling checkout is an error so one missing route cannot weaken this
    // seam silently.
    const REGISTRY_SOURCE: &str = "../litchi-iwa/src/protobuf.rs";
    const ADDITIONAL_ROUTE_SOURCES: [&str; 5] = [
        "../litchi-pages/src/package/footnote_text.rs",
        "../litchi-pages/src/package/section_background.rs",
        "../litchi-pages/src/package/section_pagination.rs",
        "../litchi-pages/src/package/section_settings.rs",
        "../litchi-pages/src/package/section_text.rs",
    ];
    let any_route_present = ROUTE_DECLARATIONS
        .iter()
        .any(|(declaration_path, _, use_path, _)| {
            Path::new(declaration_path).is_file() || Path::new(use_path).is_file()
        })
        || ADDITIONAL_ROUTE_SOURCES
            .iter()
            .any(|path| Path::new(path).is_file())
        || Path::new(REGISTRY_SOURCE).is_file();
    let all_routes_present = ROUTE_DECLARATIONS
        .iter()
        .all(|(declaration_path, _, use_path, _)| {
            Path::new(declaration_path).is_file() && Path::new(use_path).is_file()
        })
        && ADDITIONAL_ROUTE_SOURCES
            .iter()
            .all(|path| Path::new(path).is_file())
        && Path::new(REGISTRY_SOURCE).is_file();
    let route_scope_ok = if !any_route_present {
        true
    } else if !all_routes_present {
        false
    } else {
        let declarations_ok = ROUTE_DECLARATIONS
            .iter()
            .map(|(declaration_path, declaration, use_path, use_marker)| {
                fs::read_to_string(declaration_path).and_then(|declaration_source| {
                    fs::read_to_string(use_path).map(|use_source| {
                        rust_non_public_declaration_count(&declaration_source, declaration) == 1
                            && rust_code_marker_count(&use_source, use_marker) == 1
                    })
                })
            })
            .collect::<Result<Vec<_>, _>>()?
            .into_iter()
            .all(|matches| matches);
        let production_markers_ok = PRODUCTION_ROUTE_MARKERS
            .iter()
            .map(|(path, marker, expected_count)| {
                fs::read_to_string(path)
                    .map(|source| rust_code_marker_count(&source, marker) == *expected_count)
            })
            .collect::<Result<Vec<_>, _>>()?
            .into_iter()
            .all(|matches| matches);
        declarations_ok && production_markers_ok
    };

    let registry_scope_ok = if !any_route_present || !all_routes_present {
        true
    } else {
        let registry = fs::read_to_string(REGISTRY_SOURCE)?;
        REGISTRY_DECLARATIONS
            .iter()
            .all(|declaration| rust_code_marker_count(&registry, declaration) == 1)
    };

    if !canonical_scope_ok || !route_scope_ok || !registry_scope_ok {
        return Err(
            "Pages native message provenance drifted: private Movie/audio 3007, body-shape theme 10001, Caption 633, FootnoteReference 2008, TextualAttachment 2004, section template/storage (10143/2001/2022), legacy editor Document/Section (10000/10011), creation Document/Section/Settings (10000/10011/10012), package Document/Section/Settings (10000/10011/10012), or table-lock root (10000) routes no longer match their message-scoped canonical declarations and production decode/remap/write sites"
                .into(),
        );
    }
    Ok(())
}

fn enforce_pages_movie_caption_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = r#"message Reference {
  required uint64 identifier = 1;
  optional int32 deprecated_type = 2;
  optional bool deprecated_is_external = 3;
}"#;
    const DRAWABLE_FIELDS: [&str; 1] = ["optional .TSP.Reference parent = 2;"];
    const SHAPE_FIELDS: [&str; 2] = [
        "required .TSD.DrawableArchive super = 1;",
        "optional .TSP.Reference style = 2;",
    ];
    const SHAPE_INFO_FIELDS: [&str; 4] = [
        "required .TSD.ShapeArchive super = 1;",
        "optional .TSP.Reference deprecated_storage = 2 [deprecated = true];",
        "optional .TSP.Reference owned_storage = 4;",
        "optional bool is_text_box = 6;",
    ];
    const CAPTION_INFO_FIELDS: [&str; 3] = [
        "required .TSWP.ShapeInfoArchive super = 1;",
        "optional .TSP.Reference placement = 2;",
        "optional .TSD.CaptionOrTitleKind childInfoKind = 3;",
    ];
    const PROJECTION_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n}";
    const PROJECTION_DRAWABLE: &str =
        "message DrawableArchive {\n  optional .LitchiIwaProjection.Reference parent = 2;\n}";
    const PROJECTION_SHAPE: &str = "message ShapeArchive {\n  required .LitchiIwaProjection.DrawableArchive super = 1;\n  optional .LitchiIwaProjection.Reference style = 2;\n}";
    const PROJECTION_SHAPE_INFO: &str = "message ShapeInfoArchive {\n  required .LitchiIwaProjection.ShapeArchive super = 1;\n  optional .LitchiIwaProjection.Reference deprecated_storage = 2;\n  optional .LitchiIwaProjection.Reference owned_storage = 4;\n  optional bool is_text_box = 6;\n}";
    const PROJECTION_CAPTION_INFO: &str = "message CaptionInfoArchive {\n  required .LitchiIwaProjection.ShapeInfoArchive super = 1;\n  optional .LitchiIwaProjection.Reference placement = 2;\n  optional int32 child_info_kind = 3;\n}";
    const CODEC_MARKERS: [&str; 13] = [
        "const CAPTION_INFO_SUPER_FIELD: u32 = 1;",
        "const CAPTION_INFO_PLACEMENT_FIELD: u32 = 2;",
        "const CAPTION_INFO_KIND_FIELD: u32 = 3;",
        "const SHAPE_INFO_SUPER_FIELD: u32 = 1;",
        "const SHAPE_INFO_DEPRECATED_STORAGE_FIELD: u32 = 2;",
        "const SHAPE_INFO_OWNED_STORAGE_FIELD: u32 = 4;",
        "const SHAPE_INFO_IS_TEXT_BOX_FIELD: u32 = 6;",
        "const SHAPE_SUPER_FIELD: u32 = 1;",
        "const SHAPE_STYLE_FIELD: u32 = 2;",
        "const DRAWABLE_PARENT_FIELD: u32 = 2;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const MAX_RECURSION_LIMIT: u32 = 64;",
        "pub fn decode_caption_info(",
    ];

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let tsd = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let tsd_commands = fs::read_to_string(proto_directory.join("TSDCommandArchives.proto"))?;
    let tswp = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let tsa = fs::read_to_string(proto_directory.join("TSAArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TPMovieCaptionArchive.proto"))?;
    let drawable_block = tsd
        .split_once("message DrawableArchive {")
        .and_then(|(_, remainder)| remainder.split_once("\n}"))
        .map_or("", |(body, _)| body);
    let shape_block = tsd
        .split_once("message ShapeArchive {")
        .and_then(|(_, remainder)| remainder.split_once("\n}"))
        .map_or("", |(body, _)| body);
    let projection_digest = Sha256::digest(projection.as_bytes())
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let codec = fs::read_to_string("src/pages_movie_caption_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if tsp.matches(TSP_REFERENCE).count() != 1
        || tsd.matches("message DrawableArchive {").count() != 1
        || !DRAWABLE_FIELDS
            .iter()
            .all(|field| drawable_block.matches(field).count() == 1)
        || tsd.matches("message ShapeArchive {").count() != 1
        || !SHAPE_FIELDS
            .iter()
            .all(|field| shape_block.matches(field).count() == 1)
        || !SHAPE_INFO_FIELDS
            .iter()
            .all(|field| tswp.matches(field).count() == 1)
        || !CAPTION_INFO_FIELDS
            .iter()
            .all(|field| tsa.matches(field).count() == 1)
        || tsd_commands.matches("enum CaptionOrTitleKind {").count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_DRAWABLE).count() != 1
        || projection.matches(PROJECTION_SHAPE).count() != 1
        || projection.matches(PROJECTION_SHAPE_INFO).count() != 1
        || projection.matches(PROJECTION_CAPTION_INFO).count() != 1
        || projection.len() > 4 * 1024
        || projection.contains("repeated ")
        || projection_digest != "c1c8f7131f4362794811e0ce396769a37c61175d10befb63886b4a76109c6e7c"
        || !CODEC_MARKERS
            .iter()
            .all(|marker| production_codec.matches(marker).count() == 1)
        || production_codec.contains("prost::")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err("Pages movie-caption projection/codec drifted from TSA/TSWP/TSD/TSP fields, exceeded its source budget, introduced repeated storage, Prost, or generated production encoding".into());
    }
    Ok(())
}

fn enforce_pages_footnote_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_PROJECTION_DIGEST: &str =
        "6a8b19d679e9cb331764f537b9e342943f1e08860784284ad176016b4fcddde4";
    const TEXTUAL_FIELDS: [&str; 2] = [
        "optional string string_equivalent = 1;",
        "optional .TSWP.TextualAttachmentArchive.Kind kind = 2;",
    ];
    const TSP_REFERENCE: &str = r#"message Reference {
  required uint64 identifier = 1;
  optional int32 deprecated_type = 2;
  optional bool deprecated_is_external = 3;
}"#;
    const TSWP_TEXTUAL: &str = "message TextualAttachmentArchive {\n  enum Kind {\n    kKindPageNumber = 0;\n    kKindPageCount = 1;\n    kKindFootnoteMark = 2;\n  }\n  optional string string_equivalent = 1;\n  optional .TSWP.TextualAttachmentArchive.Kind kind = 2;\n}";
    const TSWP_FOOTNOTE: &str = "message FootnoteReferenceAttachmentArchive {\n  optional .TSWP.TextualAttachmentArchive super = 1;\n  optional .TSP.Reference contained_storage = 2;\n  optional string custom_mark_string = 3;\n}";
    const PROJECTION_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const PROJECTION_TEXTUAL: &str = "message TextualAttachmentArchive {\n  optional string string_equivalent = 1;\n  optional int32 kind = 2;\n}";
    const PROJECTION_FOOTNOTE: &str = "message FootnoteReferenceAttachmentArchive {\n  optional .LitchiIwaProjection.TextualAttachmentArchive super = 1;\n  optional .LitchiIwaProjection.Reference contained_storage = 2;\n  optional string custom_mark_string = 3;\n}";
    const CODEC_MARKERS: [&str; 11] = [
        "const FOOTNOTE_SUPER_FIELD: u32 = 1;",
        "const FOOTNOTE_CONTAINED_STORAGE_FIELD: u32 = 2;",
        "const FOOTNOTE_CUSTOM_MARK_FIELD: u32 = 3;",
        "const TEXTUAL_STRING_EQUIVALENT_FIELD: u32 = 1;",
        "const TEXTUAL_KIND_FIELD: u32 = 2;",
        "const REFERENCE_IDENTIFIER_FIELD: u32 = 1;",
        "const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;",
        "const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;",
        "const MAX_RECURSION_LIMIT: u32 = 64;",
        "pub fn decode_footnote_reference",
        "decode_lazy_view",
    ];

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let text = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let projection = fs::read_to_string(
        projection_directory.join("TSWPFootnoteReferenceAttachmentArchive.proto"),
    )?;
    let projection_digest = Sha256::digest(projection.as_bytes())
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let codec = fs::read_to_string("src/pages_footnote_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if tsp.matches(TSP_REFERENCE).count() != 1
        || text.matches(TSWP_TEXTUAL).count() != 1
        || !TEXTUAL_FIELDS
            .iter()
            .all(|field| text.matches(field).count() == 1)
        || text.matches(TSWP_FOOTNOTE).count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_TEXTUAL).count() != 1
        || projection.matches(PROJECTION_FOOTNOTE).count() != 1
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || projection_digest != EXPECTED_PROJECTION_DIGEST
        || !CODEC_MARKERS
            .iter()
            .all(|marker| production_codec.matches(marker).count() == 1)
        || production_codec.contains("prost::")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err("Pages footnote projection/codec drifted from TSWP textual-attachment/reference fields, exceeded its source budget, introduced repeated storage, Prost, or generated production encoding".into());
    }
    Ok(())
}

fn enforce_pages_footnote_marker_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_PROJECTION_DIGEST: &str =
        "12cd2d1186d8c0241c6439085c5d0b911fde6ecd5dec759626a2d94dceca3c15";
    const TSWP_TEXTUAL: &str = "message TextualAttachmentArchive {\n  enum Kind {\n    kKindPageNumber = 0;\n    kKindPageCount = 1;\n    kKindFootnoteMark = 2;\n  }\n  optional string string_equivalent = 1;\n  optional .TSWP.TextualAttachmentArchive.Kind kind = 2;\n}";
    const PROJECTION: &str = "syntax = \"proto2\";\n\npackage LitchiIwaProjection;\n\nmessage TextualAttachmentArchive {\n  optional string string_equivalent = 1;\n  optional int32 kind = 2;\n}";
    const CODEC_MARKERS: [&str; 6] = [
        "const TEXTUAL_STRING_EQUIVALENT_FIELD: u32 = 1;",
        "const TEXTUAL_KIND_FIELD: u32 = 2;",
        "const MAX_RECURSION_LIMIT: u32 = 64;",
        "pub fn decode_textual_attachment",
        "decode_lazy_view",
        "pub const fn raw(self)",
    ];

    let text = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("TSWPTextualAttachmentArchive.proto"))?;
    let projection_digest = Sha256::digest(projection.as_bytes())
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let codec = fs::read_to_string("src/pages_footnote_marker_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if text.matches(TSWP_TEXTUAL).count() != 1
        || projection.matches(PROJECTION).count() != 1
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || projection_digest != EXPECTED_PROJECTION_DIGEST
        || !CODEC_MARKERS
            .iter()
            .all(|marker| production_codec.matches(marker).count() == 1)
        || production_codec.contains("prost::")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err("Pages footnote-marker projection/codec drifted from TSWP TextualAttachmentArchive fields, exceeded its source budget, introduced repeated storage, Prost, or generated production encoding".into());
    }
    Ok(())
}

fn enforce_keynote_placeholder_text_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const TSD_DRAWABLE: &str = "message DrawableArchive {";
    const TSD_SHAPE_SUPER: &str = "required .TSD.DrawableArchive super = 1;";
    const TSWP_SHAPE_INFO_FIELDS: [&str; 2] = [
        "required .TSD.ShapeArchive super = 1;",
        "optional .TSP.Reference owned_storage = 4;",
    ];
    const KN_PLACEHOLDER_FIELDS: [&str; 2] = [
        "required .TSWP.ShapeInfoArchive super = 1;",
        "optional .KN.PlaceholderArchive.Kind kind = 2 [default = kKindPlaceholder];",
    ];
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message Reference {\n\
required uint64 identifier = 1;\n\
optional int32 deprecated_type = 2;\n\
optional bool deprecated_is_external = 3;\n\
}\n\
message DrawableArchive {}\n\
message ShapeArchive {\n\
required .LitchiIwaProjection.DrawableArchive super = 1;\n\
}\n\
message ShapeInfoArchive {\n\
required .LitchiIwaProjection.ShapeArchive super = 1;\n\
optional .LitchiIwaProjection.Reference owned_storage = 4;\n\
}\n\
message PlaceholderArchive {\n\
required .LitchiIwaProjection.ShapeInfoArchive super = 1;\n\
optional int32 kind = 2 [default = 0];\n\
}";

    let reference_schema = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let drawable_schema = fs::read_to_string(proto_directory.join("TSDArchives.proto"))?;
    let shape_info_schema = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let keynote_schema = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("KNPlaceholderTextOwnerArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let shape_block = drawable_schema
        .split_once("message ShapeArchive {")
        .and_then(|(_prefix, remainder)| {
            remainder.split_once("\n}\n\nmessage ConnectionLineArchive")
        })
        .map_or("", |(block, _suffix)| block);
    let codec = fs::read_to_string("src/keynote_placeholder_text_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if reference_schema.matches(TSP_REFERENCE).count() != 1
        || drawable_schema.matches(TSD_DRAWABLE).count() != 1
        || shape_block.matches(TSD_SHAPE_SUPER).count() != 1
        || !TSWP_SHAPE_INFO_FIELDS
            .iter()
            .all(|declaration| shape_info_schema.matches(declaration).count() == 1)
        || !KN_PLACEHOLDER_FIELDS
            .iter()
            .all(|declaration| keynote_schema.matches(declaration).count() == 1)
        || projection_schema != PROJECTION_SCHEMA
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Keynote placeholder-text projection drifted from the canonical KN/TSWP/TSD/TSP owner chain, exceeded its 2 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_keynote_speaker_notes_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TSP_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const KN_TRANSITION: &str = "message TransitionArchive {\n  required .KN.TransitionAttributesArchive attributes = 2;\n}";
    const KN_NOTE: &str =
        "message NoteArchive {\n  required .TSP.Reference containedStorage = 1;\n}";
    const KN_SLIDE_FIELDS: [&str; 8] = [
        "required .TSP.Reference style = 1;",
        "required .KN.TransitionArchive transition = 4;",
        "optional .TSP.Reference titlePlaceholder = 5;",
        "optional .TSP.Reference bodyPlaceholder = 6;",
        "optional string name = 10;",
        "required bool inDocument = 19;",
        "optional .TSP.Reference slideNumberPlaceholder = 20;",
        "optional .TSP.Reference note = 27;",
    ];
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message Reference {\n\
required uint64 identifier = 1;\n\
optional int32 deprecated_type = 2;\n\
optional bool deprecated_is_external = 3;\n\
}\n\
message TransitionAttributesArchive {}\n\
message TransitionArchive {\n\
required .LitchiIwaProjection.TransitionAttributesArchive attributes = 2;\n\
}\n\
message SlideArchive {\n\
required .LitchiIwaProjection.Reference style = 1;\n\
required .LitchiIwaProjection.TransitionArchive transition = 4;\n\
optional .LitchiIwaProjection.Reference title_placeholder = 5;\n\
optional .LitchiIwaProjection.Reference body_placeholder = 6;\n\
optional string name = 10;\n\
required bool in_document = 19;\n\
optional .LitchiIwaProjection.Reference slide_number_placeholder = 20;\n\
optional .LitchiIwaProjection.Reference note = 27;\n\
}\n\
message NoteArchive {\n\
required .LitchiIwaProjection.Reference contained_storage = 1;\n\
}";

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let keynote = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("KNSpeakerNotesArchive.proto"))?;
    let projection_schema = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let slide_block = keynote
        .split_once("message SlideArchive {")
        .and_then(|(_prefix, remainder)| remainder.split_once("\n}\n\nmessage SlideNodeArchive"))
        .map_or("", |(block, _suffix)| block);
    let codec = fs::read_to_string("src/keynote_speaker_notes_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    if tsp.matches(TSP_REFERENCE).count() != 1
        || keynote.matches(KN_TRANSITION).count() != 1
        || keynote.matches(KN_NOTE).count() != 1
        || !KN_SLIDE_FIELDS
            .iter()
            .all(|declaration| slide_block.matches(declaration).count() == 1)
        || projection_schema != PROJECTION_SCHEMA
        || projection.len() > 2 * 1024
        || projection.contains("repeated ")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Keynote speaker-notes projection drifted from TSP.Reference or the selected KN owner fields, exceeded its 2 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_keynote_slide_number_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const NODE: &str = "optional bool isSlideNumberVisible = 18 [default = false];";
    const STORAGE: [&str; 4] = [
        "optional .TSWP.StorageArchive.KindType kind = 1 [default = TEXTBOX];",
        "repeated string text = 3;",
        "optional .TSWP.ObjectAttributeTable table_attachment = 9;",
        "optional bool in_document = 10 [default = false];",
    ];
    const TEXTUAL: [&str; 2] = [
        "optional string string_equivalent = 1;",
        "optional .TSWP.TextualAttachmentArchive.Kind kind = 2;",
    ];
    const ATTACHMENT: &str = "required .TSWP.TextualAttachmentArchive super = 1;";
    const TABLE: &str = "repeated .TSWP.ObjectAttributeTable.ObjectAttribute entries = 1;";
    const ENTRY: [&str; 2] = [
        "required uint32 character_index = 1;",
        "optional .TSP.Reference object = 2;",
    ];
    const PROJECTION: &str = "syntax = \"proto2\";\npackage LitchiIwaProjection;\nmessage SlideNumberNodeArchive {\noptional bool is_slide_number_visible = 18 [default = false];\n}\nmessage SlideNumberStorageArchive {\noptional int32 kind = 1 [default = 3];\noptional bytes attachment_table = 9;\noptional bool in_document = 10 [default = false];\n}\nmessage TextualAttachmentArchive {\noptional string string_equivalent = 1;\noptional int32 kind = 2;\n}\nmessage SlideNumberAttachmentArchive {\nrequired .LitchiIwaProjection.TextualAttachmentArchive super = 1;\n}";
    let keynote = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let text = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("KNSlideNumberArchive.proto"))?;
    let normalized = projection
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("//"))
        .collect::<Vec<_>>()
        .join("\n");
    let codec = fs::read_to_string("src/keynote_slide_number_codec.rs")?;
    let production = production_codec_source(&codec);
    if keynote.matches(NODE).count() != 1
        || !STORAGE.iter().all(|field| text.matches(field).count() == 1)
        || !TEXTUAL.iter().all(|field| text.matches(field).count() == 1)
        || keynote.matches(ATTACHMENT).count() != 1
        || text.matches(TABLE).count() != 1
        || !ENTRY.iter().all(|field| text.contains(field))
        || normalized != PROJECTION
        || projection.len() > 2 * 1024
        || production.contains("RepeatedView")
        || production.contains("LazyRepeatedView")
        || production.contains("encode_to_vec")
        || production.contains("try_encode")
        || production.contains(".encode(")
        || !fs::read_to_string("src/lib.rs")?.contains("mod buffa_keynote_slide_number_generated")
    {
        return Err("derived Keynote slide-number projection drifted from the selected KN/TSWP fields, introduced generated repeated storage or production encoding, or lost its private generated boundary".into());
    }
    Ok(())
}

fn enforce_keynote_soundtrack_settings_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const SHOW: &str = "optional .TSP.Reference soundtrack = 17;";
    const SOUNDTRACK: [&str; 3] = [
        "optional double volume = 1;",
        "optional .KN.Soundtrack.SoundtrackMode mode = 2 [default = kKNSoundtrackModePlayOnce];",
        "repeated .TSP.DataReference movie_media = 3;",
    ];
    let keynote = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("KNSoundtrackSettingsArchive.proto"))?;
    let codec = fs::read_to_string("src/keynote_soundtrack_settings_codec.rs")?;
    let production = production_codec_source(&codec);
    if keynote.matches(SHOW).count() != 1
        || !SOUNDTRACK
            .iter()
            .all(|field| keynote.matches(field).count() == 1)
        || projection.contains("repeated ")
        || projection.len() > 2 * 1024
        || production.contains("RepeatedView")
        || production.contains("LazyRepeatedView")
        || production.contains("encode_to_vec")
        || production.contains("try_encode")
        || production.contains(".encode(")
        || !fs::read_to_string("src/lib.rs")?
            .contains("mod buffa_keynote_soundtrack_settings_generated")
    {
        return Err("derived Keynote soundtrack-settings projection drifted from Show/Soundtrack scalar routes, introduced generated repeated storage or production encoding, or lost its private boundary".into());
    }
    Ok(())
}

fn enforce_keynote_slide_transition_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const CANONICAL_SLIDE: &str = "required .KN.TransitionArchive transition = 4;";
    const CANONICAL_TRANSITION: &str = "required .KN.TransitionAttributesArchive attributes = 2;";
    const ROUTER_DECLARATIONS: [&str; 3] = [
        "const SLIDE_TRANSITION_FIELD: u32 = 4;",
        "const TRANSITION_ATTRIBUTES_FIELD: u32 = 2;",
        "const ATTRIBUTES_ANIMATION_FIELD: u32 = 8;",
    ];
    const CANONICAL_ANIMATION: [&str; 16] = [
        "optional string animation_type = 1;",
        "optional string effect = 2;",
        "optional double duration = 3;",
        "optional uint32 direction = 4;",
        "optional double delay = 5;",
        "optional bool is_automatic = 6;",
        "optional .TSP.Color color = 7;",
        "optional .TSD.PathSourceArchive custom_effect_timing_curve_1 = 8;",
        "optional .TSD.PathSourceArchive custom_effect_timing_curve_2 = 9;",
        "optional .TSD.PathSourceArchive custom_effect_timing_curve_3 = 10;",
        "optional uint32 random_number_seed = 11;",
        "optional double custom_detail = 12;",
        "optional string custom_effect_timing_curve_theme_name_1 = 13;",
        "optional string custom_effect_timing_curve_theme_name_2 = 14;",
        "optional string custom_effect_timing_curve_theme_name_3 = 15;",
        "optional bool writing_direction_is_rtl = 16;",
    ];
    const CANONICAL_ATTRIBUTES: [&str; 10] = [
        "optional .KN.AnimationAttributesArchive animationAttributes = 8;",
        "optional float custom_twist = 9;",
        "optional uint32 custom_mosaic_size = 10;",
        "optional uint32 custom_mosaic_type = 11;",
        "optional bool custom_bounce = 12;",
        "optional bool custom_magic_move_fade_unmatched_objects = 13;",
        "optional .KN.TransitionAttributesArchive.TransitionCustomAttributesTimingCurveType custom_timing_curve = 15;",
        "optional .KN.TransitionAttributesArchive.TransitionCustomAttributesTextDeliveryType custom_text_delivery_type = 16;",
        "optional bool custom_motion_blur = 17;",
        "optional float custom_travel_distance = 18;",
    ];
    const CANONICAL_SLIDE_NODE: &str = "required bool hasTransition = 7;";
    const PROJECTION_MESSAGES: [&str; 5] = [
        "message KeynoteAnimationAttributes {",
        "message KeynoteTransitionAttributes {",
        "message KeynoteTransitionArchive {",
        "message KeynoteSlideTransitionArchive {",
        "message KeynoteSlideNodeTransitionArchive {",
    ];
    const PROJECTION_FIELDS: [&str; 29] = [
        "optional string animation_type = 1;",
        "optional string effect = 2;",
        "optional double duration = 3;",
        "optional uint32 direction = 4;",
        "optional double delay = 5;",
        "optional bool is_automatic = 6;",
        "optional bytes color = 7;",
        "optional bytes custom_effect_timing_curve_1 = 8;",
        "optional bytes custom_effect_timing_curve_2 = 9;",
        "optional bytes custom_effect_timing_curve_3 = 10;",
        "optional uint32 random_number_seed = 11;",
        "optional double custom_detail = 12;",
        "optional string custom_effect_timing_curve_theme_name_1 = 13;",
        "optional string custom_effect_timing_curve_theme_name_2 = 14;",
        "optional string custom_effect_timing_curve_theme_name_3 = 15;",
        "optional bool writing_direction_is_rtl = 16;",
        "optional .LitchiIwaProjection.KeynoteAnimationAttributes animation_attributes = 8;",
        "optional float custom_twist = 9;",
        "optional uint32 custom_mosaic_size = 10;",
        "optional uint32 custom_mosaic_type = 11;",
        "optional bool custom_bounce = 12;",
        "optional bool custom_magic_move_fade_unmatched_objects = 13;",
        "optional int32 custom_timing_curve = 15;",
        "optional int32 custom_text_delivery_type = 16;",
        "optional bool custom_motion_blur = 17;",
        "optional float custom_travel_distance = 18;",
        "required .LitchiIwaProjection.KeynoteTransitionAttributes attributes = 2;",
        "required .LitchiIwaProjection.KeynoteTransitionArchive transition = 4;",
        "required bool has_transition = 7;",
    ];

    let keynote = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let projection =
        fs::read_to_string(projection_directory.join("KNSlideTransitionArchive.proto"))?;
    let codec = fs::read_to_string("src/keynote_slide_transition_codec.rs")?;
    let production_codec = production_codec_source(&codec);
    let animation_block = keynote
        .split_once("message AnimationAttributesArchive {")
        .and_then(|(_prefix, remainder)| {
            remainder.split_once("\n}\n\nmessage TransitionAttributesArchive")
        })
        .map_or("", |(block, _suffix)| block);
    let attributes_block = keynote
        .split_once("message TransitionAttributesArchive {")
        .and_then(|(_prefix, remainder)| remainder.split_once("\n}\n\nmessage TransitionArchive"))
        .map_or("", |(block, _suffix)| block);
    if keynote.matches(CANONICAL_SLIDE).count() != 1
        || keynote.matches(CANONICAL_TRANSITION).count() != 1
        || !CANONICAL_ANIMATION
            .iter()
            .all(|declaration| animation_block.matches(declaration).count() == 1)
        || !CANONICAL_ATTRIBUTES
            .iter()
            .all(|declaration| attributes_block.matches(declaration).count() == 1)
        || keynote.matches(CANONICAL_SLIDE_NODE).count() != 1
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| production_codec.matches(declaration).count() == 1)
        || !PROJECTION_MESSAGES
            .iter()
            .all(|declaration| projection.matches(declaration).count() == 1)
        || !PROJECTION_FIELDS
            .iter()
            .all(|declaration| projection.matches(declaration).count() == 1)
        || projection.len() > 4 * 1024
        || projection.contains("repeated ")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Keynote slide-transition projection/router drifted from canonical KN fields, exceeded its 4 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
    }
    Ok(())
}

fn enforce_text_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    const MAX_GENERATED_BYTES: u64 = 32 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
    }

    if files != EXPECTED_FILES || bytes > MAX_GENERATED_BYTES {
        return Err(format!(
            "TSWP text projection generated {files} files/{bytes} bytes; expected {EXPECTED_FILES} files and at most {MAX_GENERATED_BYTES} bytes"
        )
        .into());
    }
    Ok(())
}

fn enforce_full_buffa_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    // The archive/keynote sidecar intentionally covers the native TSP closure,
    // including its repeated lazy views. Pin the generated surface so a schema
    // or generator change cannot widen production's full-sidecar ingress
    // without an explicit review.
    const EXPECTED_FILES: &[&str] = &[
        "TSP.mod.rs",
        "TSPArchiveMessages.__lazy_view.rs",
        "TSPArchiveMessages.__view.rs",
        "TSPArchiveMessages.rs",
        "TSPMessages.__ext.rs",
        "TSPMessages.__lazy_view.rs",
        "TSPMessages.__view.rs",
        "TSPMessages.rs",
        "iwa_buffa_protos.rs",
    ];
    const EXPECTED_GENERATED_BYTES: u64 = 2_761_538;
    const EXPECTED_REPEATED_VIEWS: usize = 228;
    const EXPECTED_LAZY_REPEATED_VIEWS: usize = 49;
    const EXPECTED_DIGEST: &str =
        "06db03da3614be74f6802feba5a0e1b647b320aae80ad023e326052e9e912e06";

    let mut entries = fs::read_dir(directory)?
        .map(|result| result.map(|entry| (entry.file_name(), entry.path(), entry.file_type())))
        .collect::<Result<Vec<_>, _>>()?;
    entries.sort_by(|left, right| left.0.cmp(&right.0));
    let mut names = Vec::new();
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    let mut digest = Sha256::new();
    for (name, path, file_type) in entries {
        if !file_type?.is_file() {
            continue;
        }
        names.push(
            name.into_string()
                .map_err(|_name| "full Buffa sidecar generated a non-UTF-8 filename")?,
        );
        let generated = fs::read(path)?;
        bytes = bytes
            .checked_add(u64::try_from(generated.len())?)
            .ok_or("full Buffa sidecar generated-byte count overflow")?;
        let text = std::str::from_utf8(&generated)?;
        repeated_views = repeated_views
            .checked_add(text.matches("RepeatedView").count())
            .ok_or("full Buffa sidecar repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(text.matches("LazyRepeatedView").count())
            .ok_or("full Buffa sidecar lazy-repeated-view count overflow")?;
        digest.update(generated);
    }
    let aggregate_digest = digest
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    if names.iter().map(String::as_str).collect::<Vec<_>>() != EXPECTED_FILES
        || bytes != EXPECTED_GENERATED_BYTES
        || repeated_views != EXPECTED_REPEATED_VIEWS
        || lazy_repeated_views != EXPECTED_LAZY_REPEATED_VIEWS
        || aggregate_digest != EXPECTED_DIGEST
    {
        return Err(format!(
            "full Buffa sidecar generated {names:?}/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions/digest {aggregate_digest}; expected exactly {EXPECTED_FILES:?}/{EXPECTED_GENERATED_BYTES} bytes/{EXPECTED_REPEATED_VIEWS} RepeatedView mentions/{EXPECTED_LAZY_REPEATED_VIEWS} LazyRepeatedView mentions/digest {EXPECTED_DIGEST}"
        )
        .into());
    }
    Ok(())
}

fn enforce_comment_storage_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: [&str; 5] = [
        "LitchiIwaCommentStorageProjection.mod.rs",
        "TSDCommentStorageArchive.__lazy_view.rs",
        "TSDCommentStorageArchive.__view.rs",
        "TSDCommentStorageArchive.rs",
        "iwa_comment_storage_buffa_protos.rs",
    ];
    const MAX_GENERATED_BYTES: u64 = 122_000;

    let mut files = Vec::new();
    let mut bytes = 0u64;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files.push(
            entry
                .file_name()
                .to_str()
                .ok_or("generated filename is not UTF-8")?
                .to_owned(),
        );
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        if generated.contains("RepeatedView") || generated.contains("LazyRepeatedView") {
            return Err(
                "Numbers comment-storage projection generated repeated lazy storage".into(),
            );
        }
    }
    files.sort_unstable();
    let mut expected = EXPECTED_FILES
        .iter()
        .map(|name| (*name).to_owned())
        .collect::<Vec<_>>();
    expected.sort_unstable();

    if files != expected || bytes > MAX_GENERATED_BYTES {
        return Err(format!(
            "Numbers comment-storage projection generated {} files/{bytes} bytes; expected {} files and at most {MAX_GENERATED_BYTES} bytes",
            files.len(),
            EXPECTED_FILES.len(),
        )
        .into());
    }
    Ok(())
}

fn enforce_group_node_category_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    const MAX_GENERATED_BYTES: u64 = 160 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
    }

    if files != EXPECTED_FILES || bytes > MAX_GENERATED_BYTES {
        return Err(format!(
            "GroupNode category projection generated {files} files/{bytes} bytes; expected {EXPECTED_FILES} files and at most {MAX_GENERATED_BYTES} bytes"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_document_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 58,630 bytes for the singular show-reference path.
    // Keep the allowance narrow so an unreviewed closure cannot enter the
    // root projection.
    const MAX_GENERATED_BYTES: u64 = 60 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_views = 0usize;
    let mut generated_lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        generated_repeated_views = generated_repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        generated_lazy_repeated_views = generated_lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_views != 0
        || generated_lazy_repeated_views != 0
    {
        return Err(format!(
            "Keynote document projection generated {files} files/{bytes} bytes/{generated_repeated_views} RepeatedView mentions/{generated_lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_chart_caption_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    const MAX_GENERATED_BYTES: u64 = 96 * 1024;
    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut repeated = 0usize;
    let mut lazy_repeated = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        repeated = repeated
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        lazy_repeated = lazy_repeated
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }
    if files != EXPECTED_FILES || bytes > MAX_GENERATED_BYTES || repeated != 0 || lazy_repeated != 0
    {
        return Err(format!(
            "Keynote chart-caption projection generated {files} files/{bytes} bytes/{repeated} RepeatedView mentions/{lazy_repeated} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_chart_title_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    const MAX_GENERATED_BYTES: u64 = 52 * 1024;
    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut repeated = 0usize;
    let mut lazy_repeated = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        repeated = repeated
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        lazy_repeated = lazy_repeated
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }
    if files != EXPECTED_FILES || bytes > MAX_GENERATED_BYTES || repeated != 0 || lazy_repeated != 0
    {
        return Err(format!(
            "Keynote chart-title projection generated {files} files/{bytes} bytes/{repeated} RepeatedView mentions/{lazy_repeated} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_table_info_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 83,529 bytes for the table-model reference, required
    // drawable envelope, and scalar lock. Keep a narrow codegen allowance
    // without permitting an unreviewed schema closure.
    const MAX_GENERATED_BYTES: u64 = 84 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_view_mentions = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        generated_repeated_view_mentions = generated_repeated_view_mentions
            .checked_add(
                fs::read_to_string(entry.path())?
                    .matches("RepeatedView")
                    .count(),
            )
            .ok_or("generated repeated-view mention count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_view_mentions != 0
    {
        return Err(format!(
            "Numbers TableInfo/lock projection generated {files} files/{bytes} bytes/{generated_repeated_view_mentions} RepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_numbers_names_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 82,641 bytes for the three singular name shells.
    // Leave only a narrow generator/formatter allowance so another schema
    // closure cannot enter this read-only projection unnoticed.
    const MAX_GENERATED_BYTES: u64 = 84 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_views = 0usize;
    let mut generated_lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        generated_repeated_views = generated_repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        generated_lazy_repeated_views = generated_lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_views != 0
        || generated_lazy_repeated_views != 0
    {
        return Err(format!(
            "Numbers names projection generated {files} files/{bytes} bytes/{generated_repeated_views} RepeatedView mentions/{generated_lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_numbers_sheet_order_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    const EXPECTED_FILES: [&str; 5] = [
        "LitchiIwaProjection.mod.rs",
        "TNNumbersSheetReferenceArchive.__lazy_view.rs",
        "TNNumbersSheetReferenceArchive.__view.rs",
        "TNNumbersSheetReferenceArchive.rs",
        "iwa_numbers_sheet_order_buffa_protos.rs",
    ];
    // Buffa 0.9.1 emits 32,579 bytes for the isolated three-scalar reference.
    // Retain only a small formatter/codegen allowance without admitting a
    // second message or generated repeated-field machinery.
    const MAX_GENERATED_BYTES: u64 = 33 * 1024;
    const EXPECTED_DIGEST: &str =
        "2a0850fd82cfbf337ed48e582d4a998bd27e5046eb63c61f6939fa5ff1a09854";

    let mut entries = fs::read_dir(directory)?
        .map(|entry| entry.map(|value| value.path()))
        .collect::<Result<Vec<_>, _>>()?;
    entries.retain(|path| path.is_file());
    entries.sort_unstable_by(|left, right| left.file_name().cmp(&right.file_name()));
    let names = entries
        .iter()
        .map(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .map(str::to_owned)
                .ok_or_else(|| {
                    format!(
                        "generated Numbers sheet-order path is not UTF-8: {}",
                        path.display()
                    )
                })
        })
        .collect::<Result<Vec<_>, _>>()?;
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    let mut digest = Sha256::new();
    for path in &entries {
        let generated = fs::read(path)?;
        bytes = bytes
            .checked_add(u64::try_from(generated.len())?)
            .ok_or("Numbers sheet-order generated-byte count overflow")?;
        let text = std::str::from_utf8(&generated)?;
        repeated_views += text.matches("RepeatedView").count();
        lazy_repeated_views += text.matches("LazyRepeatedView").count();
        digest.update(generated);
    }
    let finalized = digest.finalize();
    let mut aggregate_digest = String::with_capacity(finalized.len() * 2);
    for byte in finalized {
        aggregate_digest.push(char::from(HEX[usize::from(byte >> 4)]));
        aggregate_digest.push(char::from(HEX[usize::from(byte & 0x0f)]));
    }
    if names != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || repeated_views != 0
        || lazy_repeated_views != 0
        || aggregate_digest != EXPECTED_DIGEST
    {
        return Err(format!(
            "Numbers sheet-order projection generated {names:?}/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions/digest {aggregate_digest}; expected {EXPECTED_FILES:?}, at most {MAX_GENERATED_BYTES} bytes, zero repeated views, and digest {EXPECTED_DIGEST}"
        )
        .into());
    }
    Ok(())
}

fn enforce_table_header_settings_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 51,480 bytes for the nine scalar settings. Leave only
    // a narrow codegen/formatter allowance without admitting table data.
    const MAX_GENERATED_BYTES: u64 = 52 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_views = 0usize;
    let mut generated_lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        generated_repeated_views = generated_repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        generated_lazy_repeated_views = generated_lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_views != 0
        || generated_lazy_repeated_views != 0
    {
        return Err(format!(
            "Numbers table-header settings projection generated {files} files/{bytes} bytes/{generated_repeated_views} RepeatedView mentions/{generated_lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_table_title_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: [&str; 5] = [
        "LitchiIwaProjection.mod.rs",
        "TSTTableTitleSettingsArchive.__lazy_view.rs",
        "TSTTableTitleSettingsArchive.__view.rs",
        "TSTTableTitleSettingsArchive.rs",
        "iwa_numbers_table_title_buffa_protos.rs",
    ];
    // Buffa 0.9.1 emits 32,332 bytes for the three scalar fields. Keep less
    // than 1.5 KiB of formatter/generator headroom; the digest below detects
    // even a within-cap change.
    const MAX_GENERATED_BYTES: u64 = 33 * 1024;
    const EXPECTED_DIGEST: &str =
        "56cfd70666ffa6079175bdab0a63a4ddd055099edf3c771ed3ad8b3051596ee1";

    let mut entries = fs::read_dir(directory)?
        .map(|result| result.map(|entry| (entry.file_name(), entry.path(), entry.file_type())))
        .collect::<Result<Vec<_>, _>>()?;
    entries.sort_by(|left, right| left.0.cmp(&right.0));
    let mut names = Vec::new();
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    let mut digest = Sha256::new();
    for (file_name, path, file_type_result) in entries {
        if !file_type_result?.is_file() {
            continue;
        }
        let name = file_name
            .into_string()
            .map_err(|_name| "Numbers table-title generated a non-UTF-8 filename")?;
        let generated = fs::read(&path)?;
        let text = std::str::from_utf8(&generated)?;
        bytes = bytes
            .checked_add(u64::try_from(generated.len())?)
            .ok_or("Numbers table-title generated byte count overflow")?;
        repeated_views = repeated_views
            .checked_add(text.matches("RepeatedView").count())
            .ok_or("Numbers table-title repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(text.matches("LazyRepeatedView").count())
            .ok_or("Numbers table-title lazy-repeated-view count overflow")?;
        digest.update(generated);
        names.push(name);
    }
    let aggregate_digest = digest
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    if names.as_slice() != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || repeated_views != 0
        || lazy_repeated_views != 0
        || aggregate_digest != EXPECTED_DIGEST
    {
        return Err(format!(
            "Numbers table-title projection generated {names:?}/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions/digest {aggregate_digest}; expected {EXPECTED_FILES:?}, at most {MAX_GENERATED_BYTES} bytes, zero repeated views, and digest {EXPECTED_DIGEST}"
        )
        .into());
    }
    Ok(())
}

fn enforce_table_cell_storage_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: &[&str] = &[
        "LitchiIwaTableCellProjection.mod.rs",
        "TSTTableCellStorageArchive.__lazy_view.rs",
        "TSTTableCellStorageArchive.__view.rs",
        "TSTTableCellStorageArchive.rs",
        "iwa_numbers_table_cell_storage_buffa_protos.rs",
    ];
    enforce_table_cell_exact_budget(
        directory,
        "Numbers table-cell storage",
        EXPECTED_FILES,
        469_001,
        "a4ad92afd34f6f276ad8fcd34e249a0738b8adc1b0477aa3e86fc447cf074776",
    )
}

fn enforce_table_cell_dependency_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: &[&str] = &[
        "LitchiIwaTableCellDependencyProjection.mod.rs",
        "TSCETableCellDependenciesArchive.__lazy_view.rs",
        "TSCETableCellDependenciesArchive.__view.rs",
        "TSCETableCellDependenciesArchive.rs",
        "iwa_numbers_table_cell_dependency_buffa_protos.rs",
    ];
    enforce_table_cell_exact_budget(
        directory,
        "Numbers table-cell dependency",
        EXPECTED_FILES,
        544_538,
        "2fba7c22aef58ed3cfe6eba1f77e5eaf79d2597dd79966e05d20e50c0e2b33b3",
    )
}

fn enforce_package_metadata_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: &[&str] = &[
        "LitchiIwaPackageMetadataProjection.mod.rs",
        "TSPPackageMetadataArchive.__lazy_view.rs",
        "TSPPackageMetadataArchive.__view.rs",
        "TSPPackageMetadataArchive.rs",
        "iwa_package_metadata_buffa_protos.rs",
    ];
    enforce_table_cell_exact_budget(
        directory,
        "PackageMetadata",
        EXPECTED_FILES,
        145_681,
        "ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31",
    )
}

fn enforce_formula_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: &[&str] = &[
        "LitchiIwaFormulaProjection.mod.rs",
        "TSCEFormulaArchive.__lazy_view.rs",
        "TSCEFormulaArchive.__view.rs",
        "TSCEFormulaArchive.rs",
        "iwa_formula_buffa_protos.rs",
    ];
    enforce_table_cell_exact_budget(
        directory,
        "FormulaArchive",
        EXPECTED_FILES,
        360_069,
        "e94549480102d09d181f89cdf82197c6d873959ac07446d7a67ec7bba9c06091",
    )
}

fn enforce_table_cell_exact_budget(
    directory: &Path,
    label: &str,
    expected_files: &[&str],
    expected_bytes: u64,
    expected_digest: &str,
) -> Result<(), Box<dyn Error>> {
    let mut entries = fs::read_dir(directory)?
        .map(|result| result.map(|entry| (entry.file_name(), entry.path(), entry.file_type())))
        .collect::<Result<Vec<_>, _>>()?;
    entries.sort_by(|left, right| left.0.cmp(&right.0));
    let mut files = Vec::new();
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    let mut digest = Sha256::new();
    for (name, path, file_type) in entries {
        if !file_type?.is_file() {
            continue;
        }
        files.push(
            name.into_string()
                .map_err(|_name| "table-cell projection generated a non-UTF-8 filename")?,
        );
        let generated = fs::read(path)?;
        bytes = bytes
            .checked_add(u64::try_from(generated.len())?)
            .ok_or("generated byte count overflow")?;
        let text = std::str::from_utf8(&generated)?;
        repeated_views = repeated_views
            .checked_add(text.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(text.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
        digest.update(generated);
    }
    let aggregate_digest = digest
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    if files.iter().map(String::as_str).collect::<Vec<_>>() != expected_files
        || bytes != expected_bytes
        || repeated_views != 0
        || lazy_repeated_views != 0
        || aggregate_digest != expected_digest
    {
        return Err(format!(
            "{label} projection generated {files:?}/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions/digest {aggregate_digest}; expected exactly {expected_files:?}/{expected_bytes} bytes/zero repeated views/digest {expected_digest}"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_show_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // The current Buffa 0.9.1 output is 138,661 bytes. Keep only a small
    // formatter/codegen patch allowance so an accidental schema expansion
    // fails at build time.
    const MAX_GENERATED_BYTES: u64 = 140 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_views = 0usize;
    let mut generated_lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        generated_repeated_views = generated_repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        generated_lazy_repeated_views = generated_lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_views != 0
        || generated_lazy_repeated_views != 0
    {
        return Err(format!(
            "Keynote show projection generated {files} files/{bytes} bytes/{generated_repeated_views} RepeatedView mentions/{generated_lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_placeholder_text_projection_budget(
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Five singular shells keep codegen compact while leaving formatter and
    // generator-version headroom. No input-width storage is generated.
    const MAX_GENERATED_BYTES: u64 = 144 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_views = 0usize;
    let mut generated_lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        generated_repeated_views = generated_repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        generated_lazy_repeated_views = generated_lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_views != 0
        || generated_lazy_repeated_views != 0
    {
        return Err(format!(
            "Keynote placeholder-text projection generated {files} files/{bytes} bytes/{generated_repeated_views} RepeatedView mentions/{generated_lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_speaker_notes_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 162,241 bytes with the semantic placeholder refs. The
    // 168-KiB ceiling leaves modest codegen/formatter headroom.
    const MAX_GENERATED_BYTES: u64 = 168 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_views = 0usize;
    let mut generated_lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        generated_repeated_views = generated_repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        generated_lazy_repeated_views = generated_lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_views != 0
        || generated_lazy_repeated_views != 0
    {
        return Err(format!(
            "Keynote speaker-notes projection generated {files} files/{bytes} bytes/{generated_repeated_views} RepeatedView mentions/{generated_lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_slide_number_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Measured after the first deterministic generation; the small cushion
    // detects accidental closure growth while permitting codegen metadata.
    // Buffa 0.9.1 deterministically emits 112,101 bytes for this five-file
    // closure. The 116-KiB cap preserves a narrow 4.3-KiB ratchet margin.
    const MAX_GENERATED_BYTES: u64 = 116 * 1024;
    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut repeated = 0usize;
    let mut lazy_repeated = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        repeated = repeated
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("generated repeated-view count overflow")?;
        lazy_repeated = lazy_repeated
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("generated lazy repeated-view count overflow")?;
    }
    if files != EXPECTED_FILES || bytes > MAX_GENERATED_BYTES || repeated != 0 || lazy_repeated != 0
    {
        return Err(format!("Keynote slide-number projection generated {files} files/{bytes} bytes/{repeated} RepeatedView mentions/{lazy_repeated} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views").into());
    }
    Ok(())
}

fn enforce_keynote_soundtrack_settings_projection_budget(
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: &[&str] = &[
        "KNSoundtrackSettingsArchive.__lazy_view.rs",
        "KNSoundtrackSettingsArchive.__view.rs",
        "KNSoundtrackSettingsArchive.rs",
        "LitchiIwaProjection.mod.rs",
        "iwa_keynote_soundtrack_settings_buffa_protos.rs",
    ];
    enforce_table_cell_exact_budget(
        directory,
        "Keynote soundtrack-settings",
        EXPECTED_FILES,
        27_753,
        "458206e0b57d8ec5ae4c3fc706bf793ccd385ab867b7e92ac30d66ab1858b4d3",
    )
}

fn enforce_keynote_slide_transition_projection_budget(
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 208,052 bytes for the five scalar-only message
    // shells. Leave a small codegen/formatter allowance without permitting a
    // second schema closure to slip in unnoticed.
    const MAX_GENERATED_BYTES: u64 = 224 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        generated_repeated_views = generated_repeated_views
            .checked_add(
                fs::read_to_string(entry.path())?
                    .matches("LazyRepeatedView")
                    .count(),
            )
            .ok_or("generated repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES || bytes > MAX_GENERATED_BYTES || generated_repeated_views != 0 {
        return Err(format!(
            "Keynote slide-transition projection generated {files} files/{bytes} bytes/{generated_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_pages_section_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: [&str; 5] = [
        "LitchiIwaProjection.mod.rs",
        "TPSectionArchive.__lazy_view.rs",
        "TPSectionArchive.__view.rs",
        "TPSectionArchive.rs",
        "iwa_pages_section_buffa_protos.rs",
    ];
    // Buffa 0.9.1 emits 132,318 bytes for the retained pagination projection
    // and the reference-aware aggregate projection. Keep a narrow allowance
    // for generator/formatter drift; the digest catches any within-cap drift.
    const MAX_GENERATED_BYTES: u64 = 136 * 1024;
    const EXPECTED_DIGEST: &str =
        "245050c1428bf926619ca05f01a5e83ef2b0edbaa0b880a3c050663556ab9440";

    let mut entries = fs::read_dir(directory)?
        .map(|result| result.map(|entry| (entry.file_name(), entry.path(), entry.file_type())))
        .collect::<Result<Vec<_>, _>>()?;
    entries.sort_by(|left, right| left.0.cmp(&right.0));
    let mut names = Vec::new();
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    let mut digest = Sha256::new();
    for (file_name, path, file_type_result) in entries {
        if !file_type_result?.is_file() {
            continue;
        }
        let name = file_name
            .into_string()
            .map_err(|_name| "Pages section projection generated a non-UTF-8 filename")?;
        let generated = fs::read(&path)?;
        let text = std::str::from_utf8(&generated)?;
        bytes = bytes
            .checked_add(u64::try_from(generated.len())?)
            .ok_or("Pages section generated-byte count overflow")?;
        repeated_views = repeated_views
            .checked_add(text.matches("RepeatedView").count())
            .ok_or("Pages section repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(text.matches("LazyRepeatedView").count())
            .ok_or("Pages section lazy-repeated-view count overflow")?;
        digest.update(generated);
        names.push(name);
    }
    let aggregate_digest = digest
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    if names.as_slice() != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || repeated_views != 0
        || lazy_repeated_views != 0
        || aggregate_digest != EXPECTED_DIGEST
    {
        return Err(format!(
            "Pages section projections generated {names:?}/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions/digest {aggregate_digest}; expected {EXPECTED_FILES:?}, at most {MAX_GENERATED_BYTES} bytes, zero repeated views, and digest {EXPECTED_DIGEST}"
        )
        .into());
    }
    Ok(())
}

fn enforce_pages_section_background_projection_budget(
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: [&str; 5] = [
        "LitchiIwaPagesBackgroundProjection.mod.rs",
        "TPSectionBackgroundArchive.__lazy_view.rs",
        "TPSectionBackgroundArchive.__view.rs",
        "TPSectionBackgroundArchive.rs",
        "iwa_pages_section_background_buffa_protos.rs",
    ];
    const EXPECTED_GENERATED_BYTES: u64 = 99_593;
    const EXPECTED_DIGEST: &str =
        "9abd261dfe79866b0718411e0da75e1001a1eeeda50770037400c9e309cbb9ca";
    let mut entries = fs::read_dir(directory)?
        .map(|result| result.map(|entry| (entry.file_name(), entry.path(), entry.file_type())))
        .collect::<Result<Vec<_>, _>>()?;
    entries.sort_by(|left, right| left.0.cmp(&right.0));
    let mut names = Vec::new();
    let mut bytes = 0u64;
    let mut repeated = 0usize;
    let mut lazy_repeated = 0usize;
    let mut digest = Sha256::new();
    for (file_name, path, file_type_result) in entries {
        if !file_type_result?.is_file() {
            continue;
        }
        let name = file_name
            .into_string()
            .map_err(|_name| "Pages background projection generated a non-UTF-8 filename")?;
        let generated = fs::read(&path)?;
        let text = std::str::from_utf8(&generated)?;
        bytes = bytes
            .checked_add(u64::try_from(generated.len())?)
            .ok_or("Pages background generated-byte overflow")?;
        repeated = repeated
            .checked_add(text.matches("RepeatedView").count())
            .ok_or("Pages background repeated-view count overflow")?;
        lazy_repeated = lazy_repeated
            .checked_add(text.matches("LazyRepeatedView").count())
            .ok_or("Pages background lazy-repeated-view count overflow")?;
        digest.update(generated);
        names.push(name);
    }
    let aggregate_digest = digest
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    if names.as_slice() != EXPECTED_FILES
        || bytes != EXPECTED_GENERATED_BYTES
        || repeated != 0
        || lazy_repeated != 0
        || aggregate_digest != EXPECTED_DIGEST
    {
        return Err(format!(
            "Pages section-background projection generated {names:?}/{bytes} bytes/{repeated} RepeatedView mentions/{lazy_repeated} LazyRepeatedView mentions/digest {aggregate_digest}; expected {EXPECTED_FILES:?}, exactly {EXPECTED_GENERATED_BYTES} bytes, zero repeated views, and digest {EXPECTED_DIGEST}"
        )
        .into());
    }
    Ok(())
}

fn enforce_pages_body_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 174,682 bytes for body/settings references, the
    // streamed section-boundary entry, and selected scalar settings/layout.
    // Leave only a small generator/formatter allowance so another schema
    // closure cannot enter this focused projection unnoticed.
    const MAX_GENERATED_BYTES: u64 = 176 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut generated_repeated_view_mentions = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("generated byte count overflow")?;
        generated_repeated_view_mentions = generated_repeated_view_mentions
            .checked_add(
                fs::read_to_string(entry.path())?
                    .matches("RepeatedView")
                    .count(),
            )
            .ok_or("generated repeated-view mention count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || generated_repeated_view_mentions != 0
    {
        return Err(format!(
            "Pages body/layout/settings projection generated {files} files/{bytes} bytes/{generated_repeated_view_mentions} RepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_pages_media_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // The selected media discriminator is one scalar with no generated
    // repeated closure. Keep a finite ceiling on generated output width.
    const MAX_GENERATED_BYTES: u64 = 64 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("Pages media generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("Pages media generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        repeated_views = repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("Pages media repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("Pages media lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || repeated_views != 0
        || lazy_repeated_views != 0
    {
        return Err(format!(
            "Pages media projection generated {files} files/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_pages_movie_caption_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // The caption-info inheritance chain is intentionally narrow and has no
    // repeated native fields. Keep a finite ceiling on generated closure
    // width so a future schema edit cannot silently widen this seam.
    const MAX_GENERATED_BYTES: u64 = 176 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("Pages movie-caption generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("Pages movie-caption generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        repeated_views = repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("Pages movie-caption repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("Pages movie-caption lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || repeated_views != 0
        || lazy_repeated_views != 0
    {
        return Err(format!(
            "Pages movie-caption projection generated {files} files/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_pages_footnote_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // The attachment envelope plus one Reference closure is intentionally
    // kept below this ceiling; a repeated native table must not enter this
    // private projection by accident.
    const MAX_GENERATED_BYTES: u64 = 128 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("Pages footnote generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("Pages footnote generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        repeated_views = repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("Pages footnote repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("Pages footnote lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || repeated_views != 0
        || lazy_repeated_views != 0
    {
        return Err(format!(
            "Pages footnote projection generated {files} files/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_pages_footnote_marker_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // A marker has only two scalar fields. Keep a finite ceiling on the
    // generated closure so a future projection edit cannot silently import
    // the wider footnote-reference graph.
    const MAX_GENERATED_BYTES: u64 = 64 * 1024;

    let mut files = 0usize;
    let mut bytes = 0u64;
    let mut repeated_views = 0usize;
    let mut lazy_repeated_views = 0usize;
    for entry_result in fs::read_dir(directory)? {
        let entry = entry_result?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        files = files
            .checked_add(1)
            .ok_or("Pages footnote-marker generated file count overflow")?;
        bytes = bytes
            .checked_add(entry.metadata()?.len())
            .ok_or("Pages footnote-marker generated byte count overflow")?;
        let generated = fs::read_to_string(entry.path())?;
        repeated_views = repeated_views
            .checked_add(generated.matches("RepeatedView").count())
            .ok_or("Pages footnote-marker repeated-view count overflow")?;
        lazy_repeated_views = lazy_repeated_views
            .checked_add(generated.matches("LazyRepeatedView").count())
            .ok_or("Pages footnote-marker lazy-repeated-view count overflow")?;
    }

    if files != EXPECTED_FILES
        || bytes > MAX_GENERATED_BYTES
        || repeated_views != 0
        || lazy_repeated_views != 0
    {
        return Err(format!(
            "Pages footnote-marker projection generated {files} files/{bytes} bytes/{repeated_views} RepeatedView mentions/{lazy_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}
