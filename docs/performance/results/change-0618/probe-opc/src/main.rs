//! Scratch publication probe for change 0618 (extends change 0593's probe).
//!
//! `tools/perf-baseline` has no ordinary-save (Path A) selector, so this probe
//! drives the documented `OpcPackage` open / mutate / publish route directly.
//! Change 0618 needs a *multi-member* regeneration to exercise Deflate-state
//! reuse in the preservation writer, so this copy adds the `addrelN` and
//! `addrelall` scenarios on top of 0593's four.
//!
//! Scenarios (each opens the package fresh):
//!   noop        open, publish with no mutation (exact-source passthrough)
//!   pkgrels     take `rels_mut()` on the package, publish (every member pristine)
//!   reblob      rewrite the first XML part with an equal, freshly allocated
//!               payload, publish (one part compared by bytes, not by pointer)
//!   addrel      add one external relationship to the first part that already
//!               has relationships (one `.rels` member regenerated)
//!   addrelN:<k> the same on the first `k` related parts (k regenerated members)
//!   addrelall   the same on every related part (all `.rels` members regenerated)
//!
//! Usage:
//!   probe corpus <root>                              one line per fixture and scenario
//!   probe corpusmulti <root>                         the multi-member scenarios only
//!   probe count <file> <scenario>                    regenerated-member census
//!   probe savebench <file> <scenario> <iterations>   publish only, N times
//!   probe time <file> <scenario> <warmup> <samples>  per-sample publish nanoseconds

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

const SCENARIOS: &[&str] = &["noop", "pkgrels", "reblob", "addrel"];
const MULTI_SCENARIOS: &[&str] = &["addrelN:2", "addrelN:8", "addrelall"];

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let usage = "usage: probe corpus <root> | probe corpusmulti <root> | probe count <file> <scenario> | probe savebench <file> <scenario> <iterations> | probe time <file> <scenario> <warmup> <samples>";
    match arguments.first().map(String::as_str) {
        Some(mode @ ("corpus" | "corpusmulti")) => {
            let root = arguments.get(1).expect(usage);
            let scenarios = if mode == "corpus" {
                SCENARIOS
            } else {
                MULTI_SCENARIOS
            };
            let mut files = Vec::new();
            collect(Path::new(root), &mut files);
            files.sort();
            for file in files {
                for scenario in scenarios {
                    let outcome = publish(&file, scenario);
                    println!("{} {scenario} {outcome}", file.display());
                }
            }
        },
        Some("count") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let mut package = OpcPackage::open(&file).expect("open package");
            let parts = package.iter_parts().count();
            let related = package.iter_parts().filter(|p| !p.rels().is_empty()).count();
            let touched = mutate(&mut package, scenario);
            println!("parts\t{parts}");
            println!("related_parts\t{related}");
            println!("mutated_parts\t{touched}");
            let bytes = PackageWriter::to_bytes(&package).expect("publish");
            println!("published_bytes\t{}", bytes.len());
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
            let mut package = OpcPackage::open(&file).expect("open package");
            mutate(&mut package, scenario);
            let mut sink = 0_u64;
            for _ in 0..warmup {
                sink = sink
                    .wrapping_add(PackageWriter::to_bytes(&package).expect("publish").len() as u64);
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
    let touched = mutate(&mut package, scenario);
    match PackageWriter::to_bytes(&package) {
        Ok(bytes) => {
            let mut hasher = Sha256::new();
            hasher.update(&bytes);
            format!("ok {touched} {} {:x}", bytes.len(), hasher.finalize())
        },
        Err(error) => format!("save-error {touched} {error:?}"),
    }
}

/// Apply the scenario; returns the number of parts actually mutated.
fn mutate(package: &mut OpcPackage, scenario: &str) -> usize {
    match scenario {
        "noop" => 0,
        "pkgrels" => {
            let _relationships = package.rels_mut();
            1
        },
        "reblob" => {
            let target = first_xml_part(package);
            if let Some(partname) = target
                && let Ok(part) = package.get_part_mut(&partname)
            {
                let copy = part.blob().to_vec();
                part.set_blob(copy);
                return 1;
            }
            0
        },
        "addrel" => add_relationships(package, 1),
        "addrelall" => add_relationships(package, usize::MAX),
        other => match other.strip_prefix("addrelN:") {
            Some(count) => add_relationships(package, count.parse().expect("addrelN count")),
            None => panic!("unknown scenario {other}"),
        },
    }
}

fn add_relationships(package: &mut OpcPackage, limit: usize) -> usize {
    let targets = related_parts(package, limit);
    let mut touched = 0;
    for (index, partname) in targets.iter().enumerate() {
        if let Ok(part) = package.get_part_mut(partname) {
            let _relationship = part.rels_mut().add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                    .to_owned(),
                "https://example.invalid/probe".to_owned(),
                format!("rIdProbe0618x{index}"),
                true,
            );
            touched += 1;
        }
    }
    touched
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

fn related_parts(package: &OpcPackage, limit: usize) -> Vec<PackURI> {
    let mut names: Vec<&PackURI> = package
        .iter_parts()
        .filter(|part| !part.rels().is_empty())
        .map(Part::partname)
        .collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    names
        .into_iter()
        .take(limit)
        .map(|name| name.clone())
        .collect()
}
