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
