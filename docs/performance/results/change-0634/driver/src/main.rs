//! Change 0606 measurement and oracle driver for the PPT record tree.
//!
//! Modes:
//!   tree <paths...>                      record-tree census (0587 `ppt-tree` shape)
//!   profile <mode> <path> <w> <s>        callgrind isolation loop
//!   time <mode> <path> <w> <s>           in-process timing, one ns per line
//!   oracle <root>                        per-fixture reader dump for byte-identity diffs
//!
//! profile/time modes: eager-open eager-slides eager-text source-open owned-edit-save

use std::fmt::Write as _;
use std::hint::black_box;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Instant;

use litchi_core::{FileSource, ReadAt};

type BoxError = Box<dyn std::error::Error>;

fn walk(root: &Path, ext: &str, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(root) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, ext, out);
            } else if path
                .extension()
                .and_then(|e| e.to_str())
                .map(|e| e.eq_ignore_ascii_case(ext))
                .unwrap_or(false)
            {
                out.push(path);
            }
        }
    }
}

fn short(err: impl std::fmt::Debug) -> String {
    let s = format!("{err:?}").replace('\n', " ");
    if s.len() > 300 {
        format!("{}...", &s[..300])
    } else {
        s
    }
}

fn file_source(path: &Path) -> Result<Arc<dyn ReadAt>, BoxError> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn tree_stats(records: &[litchi_ppt::Record], depth: usize, out: &mut (usize, usize, usize)) {
    for record in records {
        out.0 += 1;
        out.1 += record.data.len();
        out.2 = out.2.max(depth);
        tree_stats(&record.children, depth + 1, out);
    }
}

/// Copy factor of the eager PPT record tree: owned payload bytes over stream length.
fn ppt_tree(path: &Path) -> Result<(), BoxError> {
    let file = std::fs::File::open(path)?;
    let mut ole = litchi_cfb::OleFile::open(file)?;
    let data = ole.open_stream(&["PowerPoint Document"])?;
    let mut offset = 0usize;
    let mut top = Vec::new();
    while offset + 8 <= data.len() {
        match litchi_ppt::Record::parse(&data, offset) {
            Ok((record, consumed)) => {
                top.push(record);
                if consumed == 0 {
                    break;
                }
                offset += consumed;
            },
            Err(_) => offset += 1,
        }
    }
    let mut stats = (0usize, 0usize, 0usize);
    tree_stats(&top, 1, &mut stats);
    println!(
        "{}\tstream_len={}\ttop_records={}\ttotal_records={}\tpayload_bytes={}\tcopy_factor={:.2}\tmax_depth={}",
        path.display(),
        data.len(),
        top.len(),
        stats.0,
        stats.1,
        stats.1 as f64 / data.len() as f64,
        stats.2
    );
    Ok(())
}

fn one_operation(mode: &str, path: &Path, bytes: &[u8]) -> Result<usize, BoxError> {
    Ok(match mode {
        "eager-open" => {
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes.to_vec()))?;
            let presentation = package.presentation()?;
            let n = presentation.slide_count();
            black_box(&presentation);
            n
        },
        "eager-slides" => {
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes.to_vec()))?;
            let presentation = package.presentation()?;
            let slides = presentation.slides()?;
            let n = slides.len();
            black_box(&slides);
            n
        },
        "eager-text" => {
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes.to_vec()))?;
            let presentation = package.presentation()?;
            let text = presentation.text()?;
            let n = text.len();
            black_box(text);
            n
        },
        "eager-notes" => {
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes.to_vec()))?;
            let presentation = package.presentation()?;
            let mut n = 0usize;
            for slide in presentation.slides()? {
                if let Some(notes) = slide.speaker_notes()? {
                    n += notes.text()?.len();
                }
            }
            black_box(n);
            n
        },
        "source-open" => {
            let package = litchi_ppt::SourceBackedPackage::from_path(path)?;
            let presentation = package.presentation()?;
            let n = presentation.slide_count();
            black_box(&presentation);
            n
        },
        "owned-edit-save" => {
            let snapshot = litchi_ppt::text_edit::Snapshot::from_bytes(bytes.to_vec())?;
            let commit = snapshot
                .edit_text(litchi_ppt::text_edit::Target::new(litchi_ppt::text_edit::Position::new(0), litchi_ppt::text_edit::Position::new(0)))?
                .commit()?;
            let n = commit.snapshot().bytes().len();
            black_box(&commit);
            n
        },
        other => return Err(format!("unknown mode {other}").into()),
    })
}

fn profile(mode: &str, path: &Path, warmups: usize, samples: usize) -> Result<(), BoxError> {
    let bytes = std::fs::read(path)?;
    let mut checksum = 0u64;
    for _ in 0..(warmups + samples) {
        checksum = checksum.wrapping_add(black_box(one_operation(mode, path, &bytes)?) as u64);
    }
    println!(
        "{{\"mode\":\"{mode}\",\"iterations\":{},\"checksum\":{checksum}}}",
        warmups + samples
    );
    Ok(())
}

fn time(mode: &str, path: &Path, warmups: usize, samples: usize) -> Result<(), BoxError> {
    let bytes = std::fs::read(path)?;
    let mut checksum = 0u64;
    for _ in 0..warmups {
        checksum = checksum.wrapping_add(black_box(one_operation(mode, path, &bytes)?) as u64);
    }
    for _ in 0..samples {
        let started = Instant::now();
        let value = one_operation(mode, path, &bytes)?;
        let elapsed = started.elapsed().as_nanos();
        checksum = checksum.wrapping_add(black_box(value) as u64);
        println!("{elapsed}");
    }
    eprintln!("checksum={checksum}");
    Ok(())
}

fn dump_presentation(out: &mut String, presentation: &litchi_ppt::Presentation) {
    macro_rules! line {
        ($label:literal, $expr:expr) => {
            match $expr {
                Ok(value) => {
                    let _ = writeln!(out, "  {}: OK {:?}", $label, value);
                },
                Err(error) => {
                    let _ = writeln!(out, "  {}: ERR {}", $label, short(error));
                },
            }
        };
    }

    let _ = writeln!(out, "  slide_count: {}", presentation.slide_count());
    line!("text", presentation.text());
    line!("extract_text_fast", presentation.extract_text_fast());
    match presentation.slides() {
        Ok(slides) => {
            let _ = writeln!(out, "  slides: OK {}", slides.len());
            for slide in &slides {
                let _ = writeln!(
                    out,
                    "    slide {} persist={} id={} shapes={:?} text={:?}",
                    slide.slide_number(),
                    slide.persist_id(),
                    slide.slide_id(),
                    slide.shape_count(),
                    slide.text()
                );
                let _ = writeln!(out, "    shapes: {:?}", slide.shapes());
                let _ = writeln!(
                    out,
                    "    placeholders: {:?} flags: {:?} tags: {:?}",
                    slide.placeholder_atoms(),
                    slide.shape_flags(),
                    slide.shape_programmable_tags()
                );
                let _ = writeln!(
                    out,
                    "    interactions: {:?} / {:?} outline={:?}",
                    slide.shape_interactions(),
                    slide.shape_text_interactions(),
                    slide.outline_text_refs()
                );
                let _ = writeln!(out, "    notes: {:?}", slide.speaker_notes().map(|notes| notes.map(|value| value.text().map(str::to_string))));
            }
        },
        Err(error) => {
            let _ = writeln!(out, "  slides: ERR {}", short(error));
        },
    }
    line!("images", presentation.images());
    line!("has_pictures", Ok::<_, ()>(presentation.has_pictures()));
    line!("document_atom", presentation.document_atom());
    line!("document_structure", presentation.document_structure());
    line!("color_schemes", presentation.color_schemes());
    line!("fonts", presentation.fonts());
    line!("header_footers", presentation.header_footers());
    line!("hyperlinks", presentation.hyperlinks());
    line!("comments", presentation.comments());
    line!("comment_catalog", presentation.comment_catalog());
    line!("ole_objects", presentation.ole_objects());
    line!("embedded_sounds", presentation.embedded_sounds());
    line!("external_media", presentation.external_media());
    line!("programmable_tags", presentation.programmable_tags());
    line!(
        "shape_programmable_tags",
        presentation.shape_programmable_tags()
    );
    line!("smart_tags", presentation.smart_tags());
    line!("text_metachars", presentation.text_metachars());
    line!("outline_text_refs", presentation.outline_text_refs());
    line!("modify_password", presentation.modify_password());
    line!("privacy_settings", presentation.privacy_settings());
    line!(
        "presentation_advisor",
        presentation.presentation_advisor_settings()
    );
    line!("html_document", presentation.html_document_settings());
    line!("html_publish", presentation.html_publish_settings());
    line!("broadcasts", presentation.broadcasts());
    line!("envelope_data", presentation.envelope_data());
    line!("envelope_settings", presentation.envelope_settings());
    line!("routing_slip", presentation.routing_slip());
    line!("normal_view_set_info", presentation.normal_view_set_info());
    line!("notes_text_view_info", presentation.notes_text_view_info());
    line!("slide_view_information", presentation.slide_view_information());
    line!(
        "outline_sorter_view_information",
        presentation.outline_sorter_view_information()
    );
    line!("document_comparison", presentation.document_comparison());
    line!("shape_flags", presentation.shape_flags());
    line!(
        "text_special_info_defaults",
        presentation.text_special_info_defaults()
    );
    line!(
        "powerpoint12_document_properties",
        presentation.powerpoint12_document_properties()
    );
    line!(
        "powerpoint12_main_master_metadata",
        presentation.powerpoint12_main_master_metadata()
    );
    line!(
        "main_master_programmable_tags",
        presentation.main_master_programmable_tags()
    );
    line!(
        "validate_interaction_sound_references",
        presentation.validate_interaction_sound_references()
    );
    line!("charts", presentation.charts());
    let _ = writeln!(
        out,
        "  slide_directory: {:?}",
        presentation.slide_directory()
    );
}

fn oracle(root: &Path) -> Result<(), BoxError> {
    let mut files = Vec::new();
    walk(root, "ppt", &mut files);
    files.sort();
    let mut out = String::new();
    for path in files {
        let relative = path.strip_prefix(root).unwrap_or(&path);
        let _ = writeln!(out, "== {}", relative.display());
        let bytes = std::fs::read(&path)?;
        let _ = writeln!(out, "  bytes: {}", bytes.len());

        // Eager owned reader.
        match litchi_ppt::Package::from_reader(Cursor::new(bytes.clone())) {
            Ok(mut package) => match package.presentation() {
                Ok(presentation) => {
                    let _ = writeln!(out, "  eager: OK");
                    dump_presentation(&mut out, &presentation);
                },
                Err(error) => {
                    let _ = writeln!(out, "  eager: presentation ERR {}", short(error));
                },
            },
            Err(error) => {
                let _ = writeln!(out, "  eager: package ERR {}", short(error));
            },
        }

        // Source-backed reader.
        match litchi_ppt::SourceBackedPackage::from_path(&path) {
            Ok(package) => match package.presentation() {
                Ok(presentation) => {
                    let _ = writeln!(out, "  source: OK");
                    dump_presentation(&mut out, &presentation);
                },
                Err(error) => {
                    let _ = writeln!(out, "  source: presentation ERR {}", short(error));
                },
            },
            Err(error) => {
                let _ = writeln!(out, "  source: package ERR {}", short(error));
            },
        }

        // Owned text editor: parse, read, and an identity edit-and-save.
        match litchi_ppt::text_edit::Snapshot::from_bytes(bytes.clone()) {
            Ok(snapshot) => {
                let _ = writeln!(out, "  owned_edit: snapshot OK len={}", snapshot.bytes().len());
                for slide in 0..4usize {
                    for shape in 0..3usize {
                        let target = litchi_ppt::text_edit::Target::new(litchi_ppt::text_edit::Position::new(slide), litchi_ppt::text_edit::Position::new(shape));
                        match snapshot.edit_text(target) {
                            Ok(transaction) => {
                                let text = transaction.text().to_string();
                                match transaction.commit() {
                                    Ok(commit) => {
                                        let _ = writeln!(
                                            out,
                                            "    ({slide},{shape}) text={text:?} noop_len={} digest={:016x}",
                                            commit.snapshot().bytes().len(),
                                            fnv(commit.snapshot().bytes())
                                        );
                                    },
                                    Err(error) => {
                                        let _ = writeln!(
                                            out,
                                            "    ({slide},{shape}) text={text:?} commit ERR {}",
                                            short(error)
                                        );
                                    },
                                }
                                let mut mutated = snapshot.edit_text(target)?;
                                let replacement = format!("litchi-0606-{slide}-{shape}");
                                match mutated.set_text(replacement) {
                                    Ok(()) => match mutated.commit() {
                                        Ok(commit) => {
                                            let _ = writeln!(
                                                out,
                                                "    ({slide},{shape}) edited_len={} digest={:016x}",
                                                commit.snapshot().bytes().len(),
                                                fnv(commit.snapshot().bytes())
                                            );
                                        },
                                        Err(error) => {
                                            let _ = writeln!(
                                                out,
                                                "    ({slide},{shape}) edited commit ERR {}",
                                                short(error)
                                            );
                                        },
                                    },
                                    Err(error) => {
                                        let _ = writeln!(
                                            out,
                                            "    ({slide},{shape}) set_text ERR {}",
                                            short(error)
                                        );
                                    },
                                }
                            },
                            Err(error) => {
                                let _ = writeln!(
                                    out,
                                    "    ({slide},{shape}) edit ERR {}",
                                    short(error)
                                );
                            },
                        }
                    }
                }
            },
            Err(error) => {
                let _ = writeln!(out, "  owned_edit: snapshot ERR {}", short(error));
            },
        }

        // Source-backed text editor.
        match file_source(&path)
            .and_then(|source| Ok(litchi_ppt::text_edit::SourceSnapshot::open(source)?))
        {
            Ok(snapshot) => {
                let _ = writeln!(out, "  source_edit: snapshot OK len={}", snapshot.len());
                for slide in 0..3usize {
                    for shape in 0..2usize {
                        let target = litchi_ppt::text_edit::Target::new(litchi_ppt::text_edit::Position::new(slide), litchi_ppt::text_edit::Position::new(shape));
                        let _ = writeln!(
                            out,
                            "    ({slide},{shape}) read={:?}",
                            snapshot.read_text(target).map_err(short)
                        );
                    }
                }
            },
            Err(error) => {
                let _ = writeln!(out, "  source_edit: snapshot ERR {}", short(error));
            },
        }

        // Record-tree census over the document stream.
        match std::fs::File::open(&path)
            .map_err(BoxError::from)
            .and_then(|file| Ok(litchi_cfb::OleFile::open(file)?))
        {
            Ok(mut ole) => match ole.open_stream(&["PowerPoint Document"]) {
                Ok(data) => {
                    let mut offset = 0usize;
                    let mut top = Vec::new();
                    while offset + 8 <= data.len() {
                        match litchi_ppt::Record::parse(&data, offset) {
                            Ok((record, consumed)) => {
                                top.push(record);
                                if consumed == 0 {
                                    break;
                                }
                                offset += consumed;
                            },
                            Err(_) => offset += 1,
                        }
                    }
                    let mut stats = (0usize, 0usize, 0usize);
                    tree_stats(&top, 1, &mut stats);
                    let _ = writeln!(
                        out,
                        "  tree: stream={} top={} total={} payload={} depth={} digest={:016x}",
                        data.len(),
                        top.len(),
                        stats.0,
                        stats.1,
                        stats.2,
                        tree_digest(&top)
                    );
                },
                Err(error) => {
                    let _ = writeln!(out, "  tree: stream ERR {}", short(error));
                },
            },
            Err(error) => {
                let _ = writeln!(out, "  tree: ole ERR {}", short(error));
            },
        }
    }
    print!("{out}");
    Ok(())
}

fn fnv(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

fn tree_digest(records: &[litchi_ppt::Record]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    let mut pending: Vec<&litchi_ppt::Record> = records.iter().rev().collect();
    while let Some(record) = pending.pop() {
        for value in [
            u64::from(record.record_type_raw),
            u64::from(record.version),
            u64::from(record.instance),
            u64::from(record.data_length),
            record.data.len() as u64,
            record.children.len() as u64,
            fnv(&record.data),
        ] {
            hash ^= value;
            hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
        }
        pending.extend(record.children.iter().rev());
    }
    hash
}

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("tree") => {
            for path in &args[1..] {
                ppt_tree(Path::new(path))?;
            }
        },
        Some("profile") => profile(
            &args[1],
            Path::new(&args[2]),
            args[3].parse()?,
            args[4].parse()?,
        )?,
        Some("time") => time(
            &args[1],
            Path::new(&args[2]),
            args[3].parse()?,
            args[4].parse()?,
        )?,
        Some("oracle") => oracle(Path::new(&args[1]))?,
        _ => {
            return Err("usage: tree | profile | time | oracle".into());
        },
    }
    Ok(())
}
