//! Retained-policy owner probe (0361k/0362kb Snappy owners, 0361i footnotes).
//! The same source compiles against the base and candidate libraries; the
//! public calls are identical. Usage: owner_probe ROUTE ITERATIONS FIXTURE_DIR
use std::{hint::black_box, time::Instant};

fn digest_hex(bytes: &[u8]) -> String {
    // SipHash-1-3 with fixed zero keys: deterministic across processes.
    use std::hash::Hasher;
    let mut hasher = std::hash::DefaultHasher::new();
    hasher.write(bytes);
    format!("{:016x}", hasher.finish())
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let route = args[1].as_str();
    let iterations: u64 = args[2].parse().unwrap();
    let dir = std::path::PathBuf::from(&args[3]);
    let read = |name: &str| std::fs::read(dir.join(name)).unwrap();
    let mut output = Vec::new();
    let mut checksum = 0u64;
    let started;
    let elapsed;
    match route {
        "numbers_headers" => {
            let package = litchi_numbers::Package::from_bytes(&read("basic.numbers")).unwrap();
            let base = package.table_header_settings(0usize, 0usize).unwrap();
            let mut changed = base;
            changed.header_rows_frozen = Some(!base.header_rows_are_frozen());
            started = Instant::now();
            for _ in 0..iterations {
                let commit = black_box(&package)
                    .edit_table_headers(0usize, 0usize)
                    .unwrap()
                    .set(black_box(changed))
                    .commit()
                    .unwrap();
                checksum = checksum.wrapping_add(commit.diagnostics().touched_components() as u64);
                black_box(commit);
            }
            elapsed = started.elapsed();
            let commit = package.edit_table_headers(0usize, 0usize).unwrap().set(changed).commit().unwrap();
            commit.package().write_to(&mut output).unwrap();
        },
        "keynote_media" => {
            let package =
                litchi_keynote::Package::from_bytes(&read("media-replacement-native.key")).unwrap();
            let tone = read("tone.wav");
            let run = |package: &litchi_keynote::Package| {
                package
                    .edit_slide_media_data(
                        litchi_keynote::SlideSelector::index(0),
                        litchi_keynote::MovieSelector::index(0),
                        litchi_keynote::slide::media::data::MediaPart::Content,
                    )
                    .and_then(|edit| edit.set(&tone))
                    .and_then(|edit| edit.commit())
                    .unwrap()
            };
            started = Instant::now();
            for _ in 0..iterations {
                let commit = run(black_box(&package));
                checksum = checksum.wrapping_add(1);
                black_box(commit);
            }
            elapsed = started.elapsed();
            run(&package).package().write_to(&mut output).unwrap();
        },
        "pages_footnote_insert" => {
            let package = litchi_pages::Package::from_bytes(&read("footnote-fixture.pages")).unwrap();
            let position = litchi_pages::footnote::body::Position::from_utf16_index(1).unwrap();
            started = Instant::now();
            for _ in 0..iterations {
                let commit = black_box(&package)
                    .insert_body_footnote(position, black_box("Inserted"), Some("\u{2021}"))
                    .unwrap();
                checksum = checksum.wrapping_add(commit.diagnostics().touched_components() as u64);
                black_box(commit);
            }
            elapsed = started.elapsed();
            let commit = package.insert_body_footnote(position, "Inserted", Some("\u{2021}")).unwrap();
            commit.package().write_to(&mut output).unwrap();
        },
        "pages_section_name" => {
            let package = litchi_pages::Package::from_bytes(&read("basic.pages")).unwrap();
            let name = "Native Pages section name with a deliberately longer UTF-8 value";
            let run = |package: &litchi_pages::Package| {
                let mut edit = package
                    .edit_section_name(litchi_pages::SectionSelector::index(0))
                    .unwrap();
                edit.set_name(Some(black_box(name))).unwrap();
                edit.commit().unwrap()
            };
            started = Instant::now();
            for _ in 0..iterations {
                let commit = run(black_box(&package));
                checksum = checksum.wrapping_add(1);
                black_box(commit);
            }
            elapsed = started.elapsed();
            run(&package).package().write_to(&mut output).unwrap();
        },
        other => panic!("unknown route {other}"),
    }
    println!(
        "{{\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{},\"checksum\":{checksum},\"output_bytes\":{},\"output_digest\":\"{}\"}}",
        elapsed.as_nanos(),
        output.len(),
        digest_hex(&output)
    );
}
