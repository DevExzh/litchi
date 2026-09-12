use std::env;
use std::fs;
use std::path::PathBuf;

fn main() {
    let manifest = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").expect("manifest directory"));
    let source = manifest.join("../../../../../crates/litchi-xlsb/tests/data_model_identity.rs");
    println!("cargo:rerun-if-changed={}", source.display());
    let contents = fs::read_to_string(&source).expect("read committed identity fixture");
    let mut body = String::new();
    let mut in_header = true;
    for line in contents.lines() {
        if in_header && (line.starts_with("//!") || line.starts_with("///")) {
            continue;
        }
        in_header = false;
        body.push_str(line);
        body.push('\n');
    }
    let output = PathBuf::from(env::var_os("OUT_DIR").expect("out directory"))
        .join("data_model_identity.rs");
    fs::write(output, body).expect("write fixture adapter");
}
