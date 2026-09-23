//! Change 0754 differential probe. Built once against each leg's checkout
//! (Cargo.toml.template with <CHECKOUT> replaced); every line of output is a
//! deterministic function of the input bytes and the leg's library, so the two
//! legs' outputs are compared byte for byte.
use std::io::{Cursor, Write};
use std::sync::Arc;

use litchi_core::Position;
use litchi_docx::document::{CompactionPolicy, ParagraphTextReplacement, Snapshot};
use sha2::{Digest, Sha256};

/// A counting wrapper over the system allocator, so the `loop` mode can
/// report allocations and requested bytes per iteration. Probe-only.
struct Counting;

static ALLOCATIONS: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
static ALLOCATED_BYTES: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

unsafe impl std::alloc::GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: std::alloc::Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(layout.size() as u64, std::sync::atomic::Ordering::Relaxed);
        unsafe { std::alloc::System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: std::alloc::Layout) {
        unsafe { std::alloc::System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: std::alloc::Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(new_size as u64, std::sync::atomic::Ordering::Relaxed);
        unsafe { std::alloc::System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn allocation_counters() -> (u64, u64) {
    (
        ALLOCATIONS.load(std::sync::atomic::Ordering::Relaxed),
        ALLOCATED_BYTES.load(std::sync::atomic::Ordering::Relaxed),
    )
}

fn hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut out = String::with_capacity(64);
    for byte in digest.iter() {
        out.push_str(&format!("{byte:02x}"));
    }
    out
}

fn esc(value: &str) -> String {
    let mut out = String::with_capacity(value.len() + 2);
    out.push('"');
    for ch in value.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

struct Line {
    fields: Vec<(String, String)>,
}

impl Line {
    fn new(file: &str) -> Self {
        Self { fields: vec![("file".into(), esc(file))] }
    }
    fn put(&mut self, key: &str, value: String) {
        self.fields.push((key.into(), esc(&value)));
    }
    fn finish(self) -> String {
        let body: Vec<String> = self
            .fields
            .into_iter()
            .map(|(k, v)| format!("{}:{}", esc(&k), v))
            .collect();
        format!("{{{}}}", body.join(","))
    }
}

fn semantic_docx(paragraphs: usize) -> Vec<u8> {
    let mut package = litchi_docx::Package::new().expect("new package");
    let document = package.document_mut().expect("document_mut");
    for index in 0..paragraphs {
        document.add_paragraph_with_text(&format!(
            "litchi-perf-baseline-docx-semantic-v1-source-{index:05}"
        ));
    }
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).expect("to_stream");
    output.into_inner()
}

fn open(bytes: &[u8]) -> Result<litchi_docx::Package, String> {
    litchi_docx::Package::from_reader(Cursor::new(bytes.to_vec())).map_err(|e| format!("open:{e}"))
}

fn save(package: &mut litchi_docx::Package) -> String {
    let mut sink = Cursor::new(Vec::new());
    match package.to_stream(&mut sink) {
        Ok(()) => format!("sha:{}", hex(sink.get_ref())),
        Err(error) => format!("save-err:{error} | {error:?}"),
    }
}

/// One fresh package, one edit built by `stage`, publish and save.
fn edit_route(
    bytes: &[u8],
    policy: CompactionPolicy,
    stage: &dyn Fn(&mut litchi_docx::document::Edit) -> Result<(), String>,
) -> String {
    let mut package = match open(bytes) {
        Ok(package) => package,
        Err(error) => return error,
    };
    let edit = match package.edit_document() {
        Ok(edit) => edit,
        Err(error) => return format!("edit-err:{error} | {error:?}"),
    };
    let mut edit = edit.with_compaction_policy(policy);
    if let Err(error) = stage(&mut edit) {
        return error;
    }
    let commit = match package.publish_document_edit(edit) {
        Ok(commit) => commit,
        Err(error) => return format!("publish-err:{error} | {error:?}"),
    };
    let saved = save(&mut package);
    let summary = format!(
        "changed={} ops={} snap={} {}",
        commit.patch().changed(),
        commit.diagnostics().operations(),
        hex(commit.snapshot().xml_bytes()),
        saved
    );
    // Inverse publication on the same package, then save again.
    let inverse = commit.patch().inverse();
    let restored = match package.apply_document_patch(&inverse) {
        Ok(snapshot) => format!("inv-snap={}", hex(snapshot.xml_bytes())),
        Err(error) => format!("inv-err:{error} | {error:?}"),
    };
    let saved_again = save(&mut package);
    // A second edit on the published package.
    let second = match package.edit_document() {
        Ok(mut edit) => {
            let count = edit.projected().paragraph_count();
            if count == 0 {
                "second:no-paragraphs".to_string()
            } else {
                match edit.replace_paragraph_text(Position::new(count - 1), "second probe edit") {
                    Ok(_) => match package.publish_document_edit(edit) {
                        Ok(commit) => format!(
                            "second:changed={} {}",
                            commit.patch().changed(),
                            save(&mut package)
                        ),
                        Err(error) => format!("second-publish-err:{error} | {error:?}"),
                    },
                    Err(error) => format!("second-stage-err:{error} | {error:?}"),
                }
            }
        },
        Err(error) => format!("second-edit-err:{error} | {error:?}"),
    };
    format!("{summary} | {restored} {saved_again} | {second}")
}

fn source_backed_route(bytes: &[u8], target: usize) -> String {
    let source: Arc<dyn litchi_core::ReadAt> =
        Arc::new(litchi_core::OwnedSource::new(bytes.to_vec()));
    let package = match litchi_docx::source_backed::Package::from_read_at(source) {
        Ok(package) => package,
        Err(error) => return format!("sb-open-err:{error}"),
    };
    let mut edit = match package.edit_document() {
        Ok(edit) => edit,
        Err(error) => return format!("sb-edit-err:{error} | {error:?}"),
    };
    let count = edit.projected().paragraph_count();
    let snapshot = format!(
        "p={} t={} c={}",
        count,
        edit.projected().table_count(),
        edit.projected().block_content_control_count()
    );
    if count == 0 {
        return format!("{snapshot} no-paragraphs");
    }
    let target = target.min(count - 1);
    if let Err(error) = edit.replace_paragraph_text(Position::new(target), "source probe edit") {
        return format!("{snapshot} sb-stage-err:{error} | {error:?}");
    }
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => return format!("{snapshot} sb-commit-err:{error} | {error:?}"),
    };
    let mut out = Vec::new();
    match package.publish_document_commit_to_stream(&mut out, &commit) {
        Ok(published) => format!(
            "{snapshot} sb-sha:{} snap={}",
            hex(&out),
            hex(published.xml_bytes())
        ),
        Err(error) => format!("{snapshot} sb-publish-err:{error} | {error:?}"),
    }
}

fn docx(file: &str) -> String {
    let bytes = std::fs::read(file).expect("read input");
    let mut line = Line::new(file);
    let package = match open(&bytes) {
        Ok(package) => package,
        Err(error) => {
            line.put("open", error);
            return line.finish();
        },
    };
    line.put("open", "ok".into());
    match package.document() {
        Ok(document) => {
            line.put(
                "text",
                match document.text() {
                    Ok(text) => format!("len={} sha={}", text.len(), hex(text.as_bytes())),
                    Err(error) => format!("err:{error} | {error:?}"),
                },
            );
            let options = litchi_core::TextOutputOptions::new("\n", "", 1 << 30, 1 << 24);
            let mut sink = Vec::new();
            line.put(
                "sink_text",
                match document.write_text_to(&mut sink, options) {
                    Ok(report) => format!(
                        "bytes={} objects={} sha={}",
                        report.bytes_written(),
                        report.objects_written(),
                        hex(&sink)
                    ),
                    Err(error) => format!("err:{error} | {error:?} partial={}", hex(&sink)),
                },
            );
            line.put(
                "para_texts",
                match document.paragraphs() {
                    Ok(paragraphs) => {
                        let mut all = String::new();
                        let mut errors = 0usize;
                        for paragraph in &paragraphs {
                            match paragraph.text() {
                                Ok(text) => all.push_str(&text),
                                Err(error) => {
                                    errors += 1;
                                    all.push_str(&format!("<err:{error}>"));
                                },
                            }
                            all.push('\u{1}');
                        }
                        format!("n={} errors={} sha={}", paragraphs.len(), errors, hex(all.as_bytes()))
                    },
                    Err(error) => format!("err:{error} | {error:?}"),
                },
            );
        },
        Err(error) => line.put("document", format!("err:{error} | {error:?}")),
    }
    let count = match package.edit_document() {
        Ok(edit) => {
            let snapshot = edit.projected();
            line.put(
                "edit",
                format!(
                    "ok p={} t={} c={}",
                    snapshot.paragraph_count(),
                    snapshot.table_count(),
                    snapshot.block_content_control_count()
                ),
            );
            let mut ranges = String::new();
            for paragraph in snapshot.paragraphs() {
                match paragraph.text() {
                    Ok(text) => ranges.push_str(&text),
                    Err(error) => ranges.push_str(&format!("<err:{error}>")),
                }
                ranges.push('\u{1}');
            }
            line.put("edit_paragraphs", hex(ranges.as_bytes()));
            Some(snapshot.paragraph_count())
        },
        Err(error) => {
            line.put("edit", format!("err:{error} | {error:?}"));
            None
        },
    };
    drop(package);
    line.put(
        "noop",
        edit_route(&bytes, CompactionPolicy::default(), &|_edit| Ok(())),
    );
    if let Some(count) = count.filter(|count| *count > 0) {
        let mut targets = vec![0, count / 2, count - 1];
        targets.dedup();
        for target in targets {
            for (label, policy) in [
                ("one", CompactionPolicy::PreserveUnmodified),
                ("one_whole", CompactionPolicy::WholeDocument),
            ] {
                line.put(
                    &format!("{label}_{target}"),
                    edit_route(&bytes, policy, &|edit| {
                        edit.replace_paragraph_text(
                            Position::new(target),
                            format!("probe edit {target} <&> \u{e9}"),
                        )
                        .map(|_| ())
                        .map_err(|error| format!("stage-err:{error} | {error:?}"))
                    }),
                );
            }
        }
        let step = (count / 10).max(1);
        let replacements: Vec<ParagraphTextReplacement> = (0..count)
            .step_by(step)
            .map(|index| ParagraphTextReplacement::new(Position::new(index), format!("multi {index}")))
            .collect();
        line.put(
            "multi",
            edit_route(&bytes, CompactionPolicy::default(), &|edit| {
                edit.replace_body_paragraph_texts(&replacements)
                    .map(|_| ())
                    .map_err(|error| format!("stage-err:{error} | {error:?}"))
            }),
        );
        line.put("sb_mid", source_backed_route(&bytes, count / 2));
    }
    line.finish()
}

fn pptx(file: &str) -> String {
    let bytes = std::fs::read(file).expect("read input");
    let mut line = Line::new(file);
    match litchi_pptx::Package::from_bytes(&bytes) {
        Ok(package) => match package.presentation() {
            Ok(presentation) => line.put(
                "text",
                match presentation.text() {
                    Ok(text) => format!("len={} sha={}", text.len(), hex(text.as_bytes())),
                    Err(error) => format!("err:{error} | {error:?}"),
                },
            ),
            Err(error) => line.put("presentation", format!("err:{error}")),
        },
        Err(error) => line.put("open", format!("err:{error}")),
    }
    line.finish()
}

fn snapxml(file: &str) -> String {
    let bytes = std::fs::read(file).expect("read input");
    let mut line = Line::new(file);
    line.put(
        "snapshot",
        match Snapshot::from_xml(bytes) {
            Ok(snapshot) => {
                let mut texts = String::new();
                for paragraph in snapshot.paragraphs() {
                    match paragraph.text() {
                        Ok(text) => texts.push_str(&text),
                        Err(error) => texts.push_str(&format!("<err:{error}>")),
                    }
                    texts.push('\u{1}');
                }
                format!(
                    "ok p={} t={} c={} texts={}",
                    snapshot.paragraph_count(),
                    snapshot.table_count(),
                    snapshot.block_content_control_count(),
                    hex(texts.as_bytes())
                )
            },
            Err(error) => format!("err:{error} | {error:?}"),
        },
    );
    line.finish()
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let mode = args.get(1).map(String::as_str).unwrap_or("");
    let stdout = std::io::stdout();
    let mut out = stdout.lock();
    match mode {
        "gen" => {
            let dir = &args[2];
            for (name, count) in [("tiny", 24usize), ("medium", 200), ("large", 10_000)] {
                std::fs::write(format!("{dir}/semantic-{name}.docx"), semantic_docx(count))
                    .expect("write corpus");
            }
        },
        "docx" | "pptx" | "snapxml" => {
            // File names come one per line on stdin.
            let mut input = String::new();
            std::io::Read::read_to_string(&mut std::io::stdin(), &mut input).expect("stdin");
            for file in input.lines().filter(|line| !line.is_empty()) {
                let result = std::panic::catch_unwind(|| match mode {
                    "docx" => docx(file),
                    "pptx" => pptx(file),
                    _ => snapxml(file),
                });
                let text = match result {
                    Ok(text) => text,
                    Err(_) => {
                        let mut line = Line::new(file);
                        line.put("panic", "panicked".into());
                        line.finish()
                    },
                };
                writeln!(out, "{text}").expect("stdout");
            }
        },
        "loop" => {
            // loop KIND FILE N: repeat one timed operation for profiling.
            let kind = args[2].as_str();
            let file = &args[3];
            let count: usize = args[4].parse().expect("count");
            let bytes = std::fs::read(file).expect("read input");
            let started = std::time::Instant::now();
            let (allocations_before, bytes_before) = allocation_counters();
            let mut sink_len = 0usize;
            match kind {
                "scan" => {
                    for _ in 0..count {
                        let snapshot = Snapshot::from_xml(bytes.clone()).expect("scan");
                        sink_len += snapshot.paragraph_count();
                    }
                },
                "text" => {
                    let package = open(&bytes).expect("open");
                    let document = package.document().expect("document");
                    for _ in 0..count {
                        sink_len += document.text().expect("text").len();
                    }
                },
                "noop" | "one" => {
                    let mut package = open(&bytes).expect("open");
                    for iteration in 0..count {
                        let mut edit = package.edit_document().expect("edit");
                        if kind == "one" {
                            edit.replace_paragraph_text(
                                Position::new(0),
                                format!("litchi-perf-baseline-docx-semantic-v1-updated-{iteration:05}"),
                            )
                            .expect("stage");
                        }
                        let _commit = package.publish_document_edit(edit).expect("publish");
                        let mut sink = Cursor::new(Vec::new());
                        package.to_stream(&mut sink).expect("save");
                        sink_len += sink.get_ref().len();
                    }
                },
                "pptx-text" => {
                    let package = litchi_pptx::Package::from_bytes(&bytes).expect("open");
                    let presentation = package.presentation().expect("presentation");
                    for _ in 0..count {
                        sink_len += presentation.text().expect("text").len();
                    }
                },
                _ => panic!("unknown loop kind"),
            }
            let (allocations_after, bytes_after) = allocation_counters();
            eprintln!(
                "{kind}: {count} iterations, {:.3} ms each ({sink_len}); allocations {} bytes {} (whole loop)",
                started.elapsed().as_secs_f64() * 1000.0 / count as f64,
                allocations_after - allocations_before,
                bytes_after - bytes_before
            );
        },
        _ => panic!("usage: probe0754 gen DIR | docx|pptx|snapxml < files | loop KIND FILE N"),
    }
}
