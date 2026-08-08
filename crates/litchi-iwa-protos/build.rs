use std::{
    env,
    error::Error,
    fs,
    path::{Path, PathBuf},
};

fn main() -> Result<(), Box<dyn Error>> {
    const PROTO_DIRECTORY: &str = "src/protos";
    const BUFFA_PROJECTION_DIRECTORY: &str = "src/buffa-projections";
    let proto_directory = Path::new(PROTO_DIRECTORY);
    let buffa_projection_directory = Path::new(BUFFA_PROJECTION_DIRECTORY);

    println!("cargo:rerun-if-changed={PROTO_DIRECTORY}");
    println!("cargo:rerun-if-changed={BUFFA_PROJECTION_DIRECTORY}");
    println!("cargo:rerun-if-changed=src/group_node_category_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_show_codec.rs");
    println!("cargo:rerun-if-changed=src/keynote_slide_transition_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_body_codec.rs");
    println!("cargo:rerun-if-changed=src/pages_section_codec.rs");
    println!("cargo:rerun-if-changed=src/table_info_codec.rs");

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
    enforce_keynote_slide_transition_projection_provenance(
        proto_directory,
        buffa_projection_directory,
    )?;
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

    // Pages root/body traversal needs only two root references and one
    // streamed section-boundary entry. The enclosing section table stays out
    // of generated code, and strict preflight owns every ingress limit.
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

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let keynote = fs::read_to_string(proto_directory.join("KNArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("KNDocumentArchive.proto"))?;
    if tsp.matches(TSP_REFERENCE).count() != 1
        || keynote.matches(KN_DOCUMENT).count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_DOCUMENT).count() != 1
    {
        return Err(
            "derived Keynote document projection is out of sync with KN.DocumentArchive.show or TSP.Reference.identifier"
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
    const TST_TABLE_INFO: &str = "message TableInfoArchive {\n  required .TSD.DrawableArchive super = 1;\n  required .TSP.Reference tableModel = 2;\n  optional .TSP.Reference editing_state = 3 [deprecated = true];\n  optional .TSP.Reference summary_model = 4;\n  optional .TSP.Reference category_order = 5;\n  optional .TSP.Reference view_column_row_uids = 6;\n  optional .TSP.UUID group_by_uuid = 7;\n  optional .TSP.UUID hidden_states_uuid = 8;\n  optional uint32 formula_coord_space_in_pre40 = 9 [deprecated = true];\n  optional uint32 formula_coord_space = 10;\n  optional .TSCE.CoordMapperArchive pasteboard_coord_mapper = 13;\n  optional .TST.LayoutEngineArchive layout_engine = 14;\n  optional .TSP.Reference pivot_data_model = 15;\n  optional bool is_a_pivot_table = 16;\n  optional .TSP.Reference pivot_order = 17;\n}";
    const PROJECTION_SCHEMA: &str = "syntax = \"proto2\";\n\
package LitchiIwaProjection;\n\
message TableModelReference {\n\
required uint64 identifier = 1;\n\
}\n\
message TableInfoArchive {\n\
required .LitchiIwaProjection.TableModelReference table_model = 2;\n\
}";

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
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
            "derived Numbers TableInfo projection drifted from TST.TableInfoArchive.tableModel or TSP.Reference.identifier, exceeded its 1 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
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
    let router = fs::read_to_string("src/keynote_show_codec.rs")?;
    let production_router = router
        .split_once("#[cfg(test)]")
        .map_or(router.as_str(), |(production, _tests)| production);
    if tsp.matches(TSP_REFERENCE).count() != 1
        || tsp.matches(TSP_SIZE).count() != 1
        || keynote.matches(KN_SLIDE_TREE).count() != 1
        || keynote.matches(KN_SHOW).count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_SIZE).count() != 1
        || projection.matches(PROJECTION_SHOW).count() != 1
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| router.matches(declaration).count() == 1)
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
    const TP_FIELDS: [&str; 3] = [
        "required .TSA.DocumentArchive super = 15;",
        "optional .TSP.Reference body_storage = 4;",
        "optional .TSP.Reference section = 5;",
    ];
    const TSWP_BOUNDARY: &str = "message ObjectAttribute {\n    required uint32 character_index = 1;\n    optional .TSP.Reference object = 2;\n  }";
    const PROJECTION_REFERENCE: &str = "message Reference {\n  required uint64 identifier = 1;\n  optional int32 deprecated_type = 2;\n  optional bool deprecated_is_external = 3;\n}";
    const PROJECTION_DOCUMENT: &str = "message PagesDocumentBodyArchive {\n  optional .LitchiIwaProjection.Reference body_storage = 4;\n  optional .LitchiIwaProjection.Reference initial_section = 5;\n}";
    const PROJECTION_BOUNDARY: &str = "message PagesSectionBoundaryEntry {\n  required uint32 character_index = 1;\n  optional .LitchiIwaProjection.Reference section = 2;\n}";
    const ROUTER_DECLARATIONS: [&str; 9] = [
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

    let tsp = fs::read_to_string(proto_directory.join("TSPMessages.proto"))?;
    let pages = fs::read_to_string(proto_directory.join("TPArchives.proto"))?;
    let text = fs::read_to_string(proto_directory.join("TSWPArchives.proto"))?;
    let projection = fs::read_to_string(projection_directory.join("TPDocumentBodyArchive.proto"))?;
    let codec = fs::read_to_string("src/pages_body_codec.rs")?;
    let production_codec = codec
        .split_once("#[cfg(test)]")
        .map_or(codec.as_str(), |(production, _tests)| production);
    if tsp.matches(TSP_REFERENCE).count() != 1
        || !TP_FIELDS
            .iter()
            .all(|declaration| pages.matches(declaration).count() == 1)
        || text.matches(TSWP_BOUNDARY).count() != 1
        || projection.matches(PROJECTION_REFERENCE).count() != 1
        || projection.matches(PROJECTION_DOCUMENT).count() != 1
        || projection.matches(PROJECTION_BOUNDARY).count() != 1
        || !ROUTER_DECLARATIONS
            .iter()
            .all(|declaration| codec.matches(declaration).count() == 1)
        || projection.len() > 3 * 1024
        || projection.contains("repeated ")
        || production_codec.contains("to_owned_message")
        || production_codec.contains("encode_to_vec")
        || production_codec.contains("try_encode")
        || production_codec.contains(".encode(")
    {
        return Err(
            "derived Pages body projection/router drifted from canonical TP/TSWP/TSP fields, exceeded its 3 KiB source budget, introduced generated repeated storage, or added production encoding"
                .into(),
        );
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
    const MAX_GENERATED_BYTES: u64 = 64 * 1024;

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
            "Keynote document projection generated {files} files/{bytes} bytes; expected {EXPECTED_FILES} files and at most {MAX_GENERATED_BYTES} bytes"
        )
        .into());
    }
    Ok(())
}

fn enforce_table_info_projection_budget(directory: &Path) -> Result<(), Box<dyn Error>> {
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
            "Numbers TableInfo projection generated {files} files/{bytes} bytes/{generated_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
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
            "Keynote show projection generated {files} files/{bytes} bytes/{generated_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}

fn enforce_keynote_slide_transition_projection_budget(
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    const EXPECTED_FILES: usize = 5;
    // Buffa 0.9.1 emits 207,203 bytes for the five scalar-only message
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
    // Buffa 0.9.1 emits 93,867 bytes for the three singular message shells.
    // Leave only a small generator/formatter allowance so another schema
    // closure cannot enter this focused projection unnoticed.
    const MAX_GENERATED_BYTES: u64 = 96 * 1024;

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
            "Pages body projection generated {files} files/{bytes} bytes/{generated_repeated_views} LazyRepeatedView mentions; expected {EXPECTED_FILES} files, at most {MAX_GENERATED_BYTES} bytes, and no repeated views"
        )
        .into());
    }
    Ok(())
}
