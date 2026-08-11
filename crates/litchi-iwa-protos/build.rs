use std::{
    env,
    error::Error,
    fs,
    path::{Path, PathBuf},
};

use sha2::{Digest, Sha256};

fn main() -> Result<(), Box<dyn Error>> {
    const PROTO_DIRECTORY: &str = "src/protos";
    const BUFFA_PROJECTION_DIRECTORY: &str = "src/buffa-projections";
    let proto_directory = Path::new(PROTO_DIRECTORY);
    let buffa_projection_directory = Path::new(BUFFA_PROJECTION_DIRECTORY);

    println!("cargo:rerun-if-changed={PROTO_DIRECTORY}");
    println!("cargo:rerun-if-changed={BUFFA_PROJECTION_DIRECTORY}");
    println!("cargo:rerun-if-changed=src/group_node_category_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_document_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_show_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_placeholder_text_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_speaker_notes_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_slide_number_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_soundtrack_settings_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_slide_transition_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_names_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_sheet_order_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_table_header_settings_codec.rs");
    println!("cargo:rerun-if-changed=src/numbers_table_title_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_body_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_document_settings_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_page_layout_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_section_codec.rs");
    println!("cargo:rerun-if-changed=src/table_info_codec.rs");
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
    enforce_text_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_group_node_category_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_keynote_document_projection_provenance(proto_directory, buffa_projection_directory)?;
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
    enforce_numbers_names_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_numbers_sheet_order_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_table_header_settings_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
    enforce_table_title_projection_provenance(proto_directory, buffa_projection_directory)?;
    enforce_pages_body_projection_provenance(proto_directory, buffa_projection_directory)?;
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
        .out_dir(buffa_out_directory)
        .include_file("iwa_buffa_protos.rs")
        .generate_views(true)
        .lazy_views(true)
        .preserve_unknown_fields(true)
        .generate_json(false)
        .generate_text(false)
        .reflect_mode(buffa_build::ReflectMode::Off)
        .idiomatic_field_names(true)
        .compile()?;

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

    Ok(())
}

fn enforce_text_projection_provenance(
    proto_directory: &Path,
    projection_directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const TEXT_DECLARATION: &str = "repeated string text = 3;";

    let canonical = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TSWPStorageArchive.proto"))?;
    if canonical.matches(TEXT_DECLARATION).count() != 1
        || projection.matches(TEXT_DECLARATION).count() != 1
    {
        return Err(
            "derived TSWP text projection is out of sync with StorageArchive field 3".into(),
        );
    }
    Ok(())
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production_router = router
        .split_once("#[cfg(test)]")
        .map_or(router.as_str(), |(production, _tests)| production);
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
    const CANONICAL_FIELDS: [&str; 3] = [
        "optional uint32 section_start_kind = 20;",
        "optional uint32 section_page_number_kind = 21;",
        "optional uint32 section_page_number_start = 22;",
    ];
    const PROJECTION_MESSAGE: &str = "message PagesSectionPaginationArchive {\n  optional uint32 section_start_kind = 20;\n  optional uint32 section_page_number_kind = 21;\n  optional uint32 section_page_number_start = 22;\n}";

    let pages = fs::read_to_string(proto_directory.join("TPArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TPSectionArchive.proto"))?;
    let codec = fs::read_to_string("src/pages_section_codec.rs")?;
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
    if !CANONICAL_FIELDS
        .iter()
        .all(|declaration| pages.matches(declaration).count() == 1)
        || projection.matches(PROJECTION_MESSAGE).count() != 1
        || projection.len() > 1024
        || projection.contains("repeated ")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Pages section projection drifted from TP.SectionArchive fields 20--22, exceeded its 1 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
    let settings_codec = fs::read_to_string("src/pages_document_settings_codec.rs")?;
    let production_settings_codec = settings_codec
        .split_once("#[cfg(test)]")
        .map_or(settings_codec.as_str(), |(production, _tests)| production);
    let layout_codec = fs::read_to_string("src/pages_page_layout_codec.rs")?;
    let production_layout_codec = layout_codec
        .split_once("#[cfg(test)]")
        .map_or(layout_codec.as_str(), |(production, _tests)| production);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    let production = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(body, _)| body);
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
    let production = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(body, _)| body);
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
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
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
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 27,753 bytes for this scalar-only five-file closure.
    const MAX_GENERATED_BYTES: u64 = 32 * 1024;
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
        return Err(format!("Keynote soundtrack-settings projection generated {files} files/{bytes} bytes/{repeated} RepeatedView mentions/{lazy_repeated} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views").into());
    }
    Ok(())
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
    const EXPECTED_FILES: usize = 5;
    const MAX_GENERATED_BYTES: u64 = 64 * 1024;

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
            "Pages section projection generated {files} files/{bytes} bytes/{generated_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
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
