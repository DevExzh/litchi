use std::{
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

    Ok(())
}
