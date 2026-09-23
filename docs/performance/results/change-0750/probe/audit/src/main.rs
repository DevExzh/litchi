//! Change 0750 audit probe: times the `xml-minifier` audits it was built
//! against on real OOXML parts, one process per invocation.
//!
//! Usage: litchi-audit-probe-0750 <parts-dir> [case-filter] [--iterations N]
//! Without `--iterations`, each case runs 41 batches of about 20 ms and prints
//! one JSON line with the per-audit nanoseconds of every batch. With it, each
//! case runs exactly N audits untimed (for callgrind).

use std::hint::black_box;
use std::time::Instant;

use xml_minifier::audit::{self, Limits};

struct Case {
    name: String,
    run: Box<dyn Fn() -> bool>,
}

fn read(dir: &str, name: &str) -> Vec<u8> {
    std::fs::read(format!("{dir}/{name}")).unwrap_or_else(|error| panic!("{name}: {error}"))
}

/// `part` with the first byte of the value of the `<v>` or `<w:t>` element
/// nearest its middle changed, so the pair differs in one element.
fn edited(part: &[u8], open: &[u8]) -> Vec<u8> {
    let middle = part.len() / 2;
    let at = (middle..part.len())
        .find(|&index| part[index..].starts_with(open))
        .map(|index| index + open.len())
        .expect("an element to edit");
    let mut copy = part.to_vec();
    copy[at] = if copy[at] == b'7' { b'8' } else { b'7' };
    copy
}

fn corpus(bytes: &[u8]) -> Vec<Vec<u8>> {
    let mut members = Vec::new();
    let mut cursor = 0;
    while cursor + 4 <= bytes.len() {
        let length = u32::from_le_bytes(bytes[cursor..cursor + 4].try_into().unwrap()) as usize;
        members.push(bytes[cursor + 4..cursor + 4 + length].to_vec());
        cursor += 4 + length;
    }
    members
}

fn cases(dir: &str) -> Vec<Case> {
    let limits = Limits::default();
    let mut cases = Vec::new();
    for part in [
        "ws-structured-1.5MB.xml",
        "ws-patriarch-3.4MB.xml",
        "docx-drawing-288KB.xml",
        "docx-table-alignment-40KB.xml",
        "pptx-slide11-121KB.xml",
    ] {
        let bytes = read(dir, part);
        cases.push(Case {
            name: format!("source/{part}"),
            run: Box::new(move || audit::verify_source(black_box(&bytes), limits).is_ok()),
        });
    }
    for (part, open) in [
        ("ws-structured-1.5MB.xml", b"<v>".as_slice()),
        ("docx-drawing-288KB.xml", b"<w:t>".as_slice()),
    ] {
        let original = read(dir, part);
        let replacement = edited(&original, open);
        cases.push(Case {
            name: format!("pair/{part}"),
            run: Box::new(move || {
                matches!(
                    audit::verify_source_replacement(
                        black_box(&original),
                        black_box(&replacement),
                        limits
                    ),
                    Ok(audit::ReplacementProof::Window { .. })
                )
            }),
        });
    }
    let members = corpus(&read(dir, "corpus-accepted.bin"));
    cases.push(Case {
        name: "source/corpus-accepted".to_owned(),
        run: Box::new(move || {
            members
                .iter()
                .all(|member| audit::verify_source(black_box(member), limits).is_ok())
        }),
    });
    // Tiny documents, as fragment audits pass them, where a fixed cost per
    // audit would show.
    let tiny: &'static [u8] = b"<w:p><w:r><w:t>text</w:t></w:r></w:p>";
    cases.push(Case {
        name: "authored/tiny-fragment".to_owned(),
        run: Box::new(move || audit::verify_authored(black_box(tiny), limits).is_ok()),
    });
    let small: &'static [u8] =
        b"<w:p xmlns:w=\"urn:w\"><w:r><w:t>text</w:t></w:r></w:p>";
    cases.push(Case {
        name: "source/tiny-document".to_owned(),
        run: Box::new(move || audit::verify_source(black_box(small), limits).is_ok()),
    });
    let compact = read(dir, "ws-patriarch-3.4MB.xml");
    cases.push(Case {
        name: "authored/ws-patriarch-3.4MB.xml".to_owned(),
        run: Box::new(move || audit::verify_authored(black_box(&compact), limits).is_ok()),
    });
    cases
}

fn main() {
    let arguments: Vec<String> = std::env::args().collect();
    let dir = &arguments[1];
    let filter = arguments.get(2).filter(|value| !value.starts_with("--")).cloned();
    let fixed = arguments
        .iter()
        .position(|value| value == "--iterations")
        .map(|index| arguments[index + 1].parse::<usize>().unwrap());
    for case in cases(dir) {
        if filter.as_ref().is_some_and(|filter| !case.name.contains(filter.as_str())) {
            continue;
        }
        let verdict = (case.run)();
        if let Some(iterations) = fixed {
            for _ in 0..iterations {
                black_box((case.run)());
            }
            println!("{{\"case\":\"{}\",\"verdict\":{verdict},\"iterations\":{iterations}}}", case.name);
            continue;
        }
        // Calibrate a batch of about 20 ms.
        let start = Instant::now();
        let mut calibration = 0u64;
        while start.elapsed().as_millis() < 20 {
            black_box((case.run)());
            calibration += 1;
        }
        let per_batch = calibration.max(1);
        for _ in 0..3 {
            black_box((case.run)());
        }
        let mut batches = Vec::new();
        for _ in 0..41 {
            let start = Instant::now();
            for _ in 0..per_batch {
                black_box((case.run)());
            }
            batches.push(start.elapsed().as_nanos() as f64 / per_batch as f64);
        }
        let rendered: Vec<String> = batches.iter().map(|ns| format!("{ns:.1}")).collect();
        println!(
            "{{\"case\":\"{}\",\"verdict\":{verdict},\"per_batch\":{per_batch},\"ns_per_audit\":[{}]}}",
            case.name,
            rendered.join(",")
        );
    }
}
