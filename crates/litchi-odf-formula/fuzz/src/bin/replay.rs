use std::path::{Path, PathBuf};

fn main() {
    let arguments: Vec<_> = std::env::args_os().skip(1).map(PathBuf::from).collect();
    let inputs = if arguments.is_empty() {
        vec![PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("corpus/formula")]
    } else {
        arguments
    };
    let mut corpus = Vec::new();
    for input in inputs {
        collect(&input, &mut corpus);
    }
    corpus.sort_unstable();
    assert!(!corpus.is_empty(), "Formula fuzz corpus is empty");
    for path in &corpus {
        let data = std::fs::read(path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        litchi_odf_formula_fuzz::exercise(&data);
    }
    println!("replayed {} Formula fuzz inputs", corpus.len());
}

fn collect(path: &Path, output: &mut Vec<PathBuf>) {
    if path.is_file() {
        output.push(path.to_path_buf());
        return;
    }
    let entries = std::fs::read_dir(path).unwrap_or_else(|error| {
        panic!(
            "failed to read corpus directory {}: {error}",
            path.display()
        )
    });
    for entry in entries {
        let path = entry
            .unwrap_or_else(|error| panic!("failed to read corpus entry: {error}"))
            .path();
        if path.is_file() {
            output.push(path);
        }
    }
}
