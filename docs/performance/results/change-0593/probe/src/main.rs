//! Scratch publication probe for change 0593.
//!
//! `tools/perf-baseline` has no ordinary-save (Path A) selector: it measures
//! source-backed publication only. This probe drives the documented
//! `OpcPackage` open / mutate / publish route directly so the publication plan
//! and the preservation writer can be counted and timed, and so the whole
//! OOXML fixture corpus can be checked for save-digest and refusal identity
//! between two builds.
//!
//! Scenarios (each opens the package fresh):
//!   noop        open, publish with no mutation (exact-source passthrough)
//!   pkgrels     take `rels_mut()` on the package, publish (every member pristine)
//!   reblob      rewrite the first XML part with an equal, freshly allocated
//!               payload, publish (one part compared by bytes, not by pointer)
//!   addrel      add one external relationship to the first part that already
//!               has relationships, publish (one `.rels` member regenerated)
//!
//! Usage:
//!   probe corpus <root>                       one line per fixture and scenario
//!   probe bench <file> <scenario> <iterations>  repeat open+publish
//!   probe savebench <file> <scenario> <iterations>  repeat publish only
//!   probe time <file> <scenario> <warmup> <samples>  per-sample publish nanoseconds
//!   probe save <file> <scenario> <out>        publish through `save(path)`

use std::env;
use std::path::{Path, PathBuf};

use litchi_opc::package::OpcPackage;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::Part;
use litchi_opc::pkgwriter::PackageWriter;
use sha2::{Digest, Sha256};

const EXTENSIONS: &[&str] = &[
    "xlsx", "xlsm", "xltx", "xltm", "docx", "docm", "dotx", "dotm", "pptx", "pptm", "potx", "ppsx",
    "xlsb",
];

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let usage = "usage: probe corpus <root> | probe bench <file> <scenario> <iterations> | probe save <file> <scenario> <out>";
    match arguments.first().map(String::as_str) {
        Some("corpus") => {
            let root = arguments.get(1).expect(usage);
            let mut files = Vec::new();
            collect(Path::new(root), &mut files);
            files.sort();
            for file in files {
                for scenario in ["noop", "pkgrels", "reblob", "addrel"] {
                    let outcome = publish(&file, scenario);
                    println!("{} {scenario} {outcome}", file.display());
                }
            }
        },
        Some("bench") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let iterations: u32 = arguments.get(3).expect(usage).parse().expect("iterations");
            let mut mixed = 0_u64;
            for _ in 0..iterations {
                let outcome = publish(&file, scenario);
                mixed = mixed.wrapping_add(outcome.len() as u64);
            }
            println!("{mixed}");
        },
        Some("savebench") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let iterations: u32 = arguments.get(3).expect(usage).parse().expect("iterations");
            // Open and mutate once so an isolation pair over `iterations`
            // differences whole publications and nothing else.
            let mut package = OpcPackage::open(&file).expect("open package");
            mutate(&mut package, scenario);
            let mut sink = 0_u64;
            for _ in 0..iterations {
                let bytes = PackageWriter::to_bytes(&package).expect("publish");
                sink = sink.wrapping_add(bytes.len() as u64);
            }
            println!("{sink}");
        },
        Some("time") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let warmup: u32 = arguments.get(3).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(4).expect(usage).parse().expect("samples");
            // Open and mutate once, outside the timer: every publish from the
            // resulting package repeats the same plan and the same writer.
            let mut package = OpcPackage::open(&file).expect("open package");
            mutate(&mut package, scenario);
            let mut sink = 0_u64;
            for _ in 0..warmup {
                sink =
                    sink.wrapping_add(PackageWriter::to_bytes(&package).expect("publish").len() as u64);
            }
            let mut timings = Vec::with_capacity(samples as usize);
            for _ in 0..samples {
                let start = std::time::Instant::now();
                let bytes = PackageWriter::to_bytes(&package).expect("publish");
                let elapsed = start.elapsed();
                sink = sink.wrapping_add(bytes.len() as u64);
                timings.push(elapsed.as_nanos());
            }
            for timing in &timings {
                println!("{timing}");
            }
            eprintln!("sink {sink}");
        },
        Some("timesave") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let out = PathBuf::from(arguments.get(3).expect(usage));
            let warmup: u32 = arguments.get(4).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(5).expect(usage).parse().expect("samples");
            let mut package = OpcPackage::open(&file).expect("open package");
            mutate(&mut package, scenario);
            for _ in 0..warmup {
                PackageWriter::write(&out, &package).expect("publish");
            }
            let mut timings = Vec::with_capacity(samples as usize);
            for _ in 0..samples {
                let start = std::time::Instant::now();
                PackageWriter::write(&out, &package).expect("publish");
                timings.push(start.elapsed().as_nanos());
            }
            for timing in &timings {
                println!("{timing}");
            }
        },
        Some("save") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let out = PathBuf::from(arguments.get(3).expect(usage));
            let mut package = OpcPackage::open(&file).expect("open package");
            mutate(&mut package, scenario);
            PackageWriter::write(&out, &package).expect("publish package");
            println!("saved {}", out.display());
        },
        _ => {
            eprintln!("{usage}");
            std::process::exit(2);
        },
    }
}

fn collect(root: &Path, files: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect(&path, files);
        } else if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| EXTENSIONS.contains(&extension.to_ascii_lowercase().as_str()))
        {
            files.push(path);
        }
    }
}

/// Run one scenario and report either the published digest or the typed error.
fn publish(file: &Path, scenario: &str) -> String {
    let mut package = match OpcPackage::open(file) {
        Ok(package) => package,
        Err(error) => return format!("open-error {error:?}"),
    };
    mutate(&mut package, scenario);
    match PackageWriter::to_bytes(&package) {
        Ok(bytes) => {
            let mut hasher = Sha256::new();
            hasher.update(&bytes);
            format!("ok {} {:x}", bytes.len(), hasher.finalize())
        },
        Err(error) => format!("save-error {error:?}"),
    }
}

fn mutate(package: &mut OpcPackage, scenario: &str) {
    match scenario {
        "noop" => {},
        "pkgrels" => {
            let _relationships = package.rels_mut();
        },
        "reblob" => {
            let target = first_xml_part(package);
            if let Some(partname) = target
                && let Ok(part) = package.get_part_mut(&partname)
            {
                let copy = part.blob().to_vec();
                part.set_blob(copy);
            }
        },
        "addrel" => {
            let target = first_related_part(package);
            if let Some(partname) = target
                && let Ok(part) = package.get_part_mut(&partname)
            {
                let _relationship = part.rels_mut().add_relationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                        .to_owned(),
                    "https://example.invalid/probe".to_owned(),
                    "rIdProbe0593".to_owned(),
                    true,
                );
            }
        },
        other => panic!("unknown scenario {other}"),
    }
}

fn first_xml_part(package: &OpcPackage) -> Option<PackURI> {
    let mut names: Vec<&PackURI> = package
        .iter_parts()
        .filter(|part| part.partname().as_str().ends_with(".xml"))
        .map(Part::partname)
        .collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    names.first().map(|name| (*name).clone())
}

fn first_related_part(package: &OpcPackage) -> Option<PackURI> {
    let mut names: Vec<&PackURI> = package
        .iter_parts()
        .filter(|part| !part.rels().is_empty())
        .map(Part::partname)
        .collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    names.first().map(|name| (*name).clone())
}
