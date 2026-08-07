use std::{
    env,
    error::Error,
    fs,
    path::{Path, PathBuf},
};

fn main() -> Result<(), Box<dyn Error>> {
    const PROTO_DIRECTORY: &str = "src/protos";
    let proto_directory = Path::new(PROTO_DIRECTORY);

    println!("cargo:rerun-if-changed={PROTO_DIRECTORY}");

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

    prost_build::Config::new()
        .include_file("iwa_protos.rs")
        .compile_protos(&proto_files, &[proto_directory])?;

    // Start the Buffa sidecar at the archive-header seam. Expand it
    // format-by-format after bounded adapters land; Prost remains the
    // full-corpus raw-schema generator.
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

    Ok(())
}
