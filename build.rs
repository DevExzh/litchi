fn main() -> std::io::Result<()> {
    println!("cargo:rerun-if-changed=src/iwa/protos/");

    // Configure prost-build
    let mut config = prost_build::Config::new();

    // Collect all .proto files from the protos directory for complete compilation
    let all_proto_files = std::fs::read_dir("src/iwa/protos")
        .expect("Failed to read protos directory")
        .filter_map(|entry| {
            let entry = entry.ok()?;
            let path = entry.path();
            if path.extension()?.to_str()? == "proto" {
                Some(path.to_string_lossy().to_string())
            } else {
                None
            }
        })
        .collect::<Vec<_>>();

    println!("Compiling all {} protobuf files together for proper dependency resolution", all_proto_files.len());

    // Compile all protobuf files - will fail the build if any errors occur
    match config
        .enable_type_names()
        .include_file("iwa_protos.rs")
        .compile_protos(&all_proto_files, &["src/iwa/protos"]) {
            Ok(_) => println!("Successfully compiled all protobuf files"),
            Err(e) => {
                eprintln!("Failed to compile protobuf files: {}\n", e);
                panic!("Protobuf compilation failed - check for syntax errors in .proto files");
            }
        }

    Ok(())
}
