//! Probe for record 0760 (ADR 0032: the snapshot slide-root memo re-applied).
//!
//! The timing modes loop exactly the harness's timed public calls over the
//! harness's own deterministic corpus (the generator is copied from
//! `tools/perf-baseline`, as change 0743's probe did), so `perf stat` over two
//! iteration counts isolates one region. The `diff` mode runs a fixed set of
//! public flows over every repository PPTX fixture and over deterministic
//! mutations of their slide XML, printing one outcome line per step; two
//! builds (base and branch) are compared line by line.
#![allow(clippy::all)]

use std::fmt::Debug;
use std::io::Write;
use std::time::Instant;

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::opened::{Commit, Limits, Snapshot};
use litchi_pptx::Package;
use sha2::{Digest, Sha256};

#[cfg(feature = "count-alloc")]
mod counting {
    //! Probe-only counting allocator: calls and requested bytes on this
    //! process, read around one region. Not used by any production code.
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicU64, Ordering};

    pub static CALLS: AtomicU64 = AtomicU64::new(0);
    pub static BYTES: AtomicU64 = AtomicU64::new(0);

    pub struct Counting;

    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            CALLS.fetch_add(1, Ordering::Relaxed);
            BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
            unsafe { System.alloc(layout) }
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            unsafe { System.dealloc(ptr, layout) }
        }

        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            CALLS.fetch_add(1, Ordering::Relaxed);
            BYTES.fetch_add(new_size as u64, Ordering::Relaxed);
            unsafe { System.realloc(ptr, layout, new_size) }
        }
    }

    #[global_allocator]
    static GLOBAL: Counting = Counting;

    pub fn read() -> (u64, u64) {
        (CALLS.load(Ordering::Relaxed), BYTES.load(Ordering::Relaxed))
    }
}

#[cfg(feature = "count-alloc")]
fn allocations() -> (u64, u64) {
    counting::read()
}

#[cfg(not(feature = "count-alloc"))]
fn allocations() -> (u64, u64) {
    (0, 0)
}

// ---------------------------------------------------------------------------
// The harness corpus (copied from tools/perf-baseline/src/lib.rs).
// ---------------------------------------------------------------------------

fn semantic_pptx_text(slide: usize, shape: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-pptx-semantic-v1-{state}-{slide:03}-{shape:03}")
}

fn shape_dims(name: &str) -> (usize, usize) {
    match name {
        "tiny" => (3, 4),
        "medium" => (12, 8),
        "large" => (100, 100),
        _ => panic!("unknown shape"),
    }
}

fn build(slides: usize, boxes: usize) -> Vec<u8> {
    let mut package = Package::new().unwrap();
    let presentation = package.presentation_mut().unwrap();
    for slide_index in 0..slides {
        let slide = presentation.add_slide().unwrap();
        for shape_index in 0..boxes {
            slide.add_text_box(
                &semantic_pptx_text(slide_index, shape_index, false),
                36 + i64::try_from(shape_index % 4).unwrap() * 180,
                36 + i64::try_from(shape_index / 4).unwrap() * 90,
                144,
                54,
            );
        }
    }
    package.to_bytes().unwrap()
}

fn update_indices(count: usize) -> Vec<usize> {
    let updates = (count + 99) / 100;
    (0..updates).map(|index| index * count / updates).collect()
}

fn median(mut values: Vec<u128>) -> u128 {
    values.sort_unstable();
    values[values.len() / 2]
}

// ---------------------------------------------------------------------------
// Differential outcome formatting.
// ---------------------------------------------------------------------------

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn digest(bytes: &[u8]) -> String {
    hex(&Sha256::digest(bytes))[..32].to_owned()
}

fn error_text<E: Debug>(error: &E) -> String {
    let text = format!("{error:?}");
    if text.chars().count() > 200 {
        let head: String = text.chars().take(120).collect();
        format!("ERR {head}..#{}", digest(text.as_bytes()))
    } else {
        format!("ERR {text}")
    }
}

fn staged_text<E: Debug>(result: &Result<bool, E>) -> String {
    match result {
        Ok(changed) => format!("staged={changed}"),
        Err(error) => error_text(error),
    }
}

fn snapshot_summary(snapshot: &Snapshot) -> String {
    let mut identities = String::new();
    for slide in snapshot.slides() {
        identities.push_str(&format!(
            "{}|{}|{};",
            slide.id(),
            slide.name(),
            slide.part_name().as_str()
        ));
    }
    format!(
        "rev={} slides={} ids={}",
        &hex(&snapshot.revision())[..32],
        snapshot.slides().len(),
        digest(identities.as_bytes())
    )
}

fn commit_summary(commit: &Commit) -> String {
    let patch = match commit.patch().to_bytes() {
        Ok(bytes) => digest(&bytes),
        Err(error) => error_text(&error),
    };
    format!(
        "changed={} patch={patch} {}",
        commit.is_changed(),
        snapshot_summary(commit.snapshot())
    )
}

fn bytes_text(package: &mut Package) -> String {
    match package.to_bytes() {
        Ok(bytes) => format!("bytes={}", digest(&bytes)),
        Err(error) => format!("bytes={}", error_text(&error)),
    }
}

type Make<'a> = &'a dyn Fn() -> litchi_pptx::Result<Package>;

/// Publish `commit` on a fresh facade made by `make`.
fn publish(make: Make<'_>, commit: Commit) -> (String, Option<(Package, Snapshot)>) {
    let mut facade = match make() {
        Ok(facade) => facade,
        Err(error) => return (error_text(&error), None),
    };
    match facade.apply_opened_presentation_commit(commit) {
        Ok(published) => {
            let bytes = bytes_text(&mut facade);
            (
                format!("{} {bytes}", snapshot_summary(&published)),
                Some((facade, published)),
            )
        },
        Err(error) => (error_text(&error), None),
    }
}

struct Emitter<'w, W: Write> {
    label: String,
    out: &'w mut W,
    lines: usize,
}

impl<W: Write> Emitter<'_, W> {
    fn emit(&mut self, step: &str, outcome: String) {
        writeln!(self.out, "{}|{step}|{outcome}", self.label).unwrap();
        self.lines += 1;
    }
}

/// Commit `edit`, emit its outcome, then publish it and run one more
/// transaction on the published snapshot (whose memo a rebind projected) and
/// one on the committed snapshot (whose memo the recapture built).
fn commit_and_follow<W: Write>(
    emitter: &mut Emitter<'_, W>,
    make: Make<'_>,
    step: &str,
    edit: litchi_pptx::opened::Transaction,
    follow_slide: usize,
) {
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            emitter.emit(&format!("{step}.commit"), error_text(&error));
            return;
        },
    };
    emitter.emit(&format!("{step}.commit"), commit_summary(&commit));
    let mut chained = commit.snapshot().edit();
    let staged = chained.set_shape_text(follow_slide, 0, format!("0760 chained {follow_slide}"));
    if matches!(staged, Ok(true)) {
        match chained.commit() {
            Ok(next) => emitter.emit(&format!("{step}.chained"), commit_summary(&next)),
            Err(error) => emitter.emit(&format!("{step}.chained"), error_text(&error)),
        }
    } else {
        emitter.emit(&format!("{step}.chained.stage"), staged_text(&staged));
    }
    let (text, published) = publish(make, commit);
    emitter.emit(&format!("{step}.publish"), text);
    let Some((mut facade, published)) = published else {
        return;
    };
    let mut next = published.edit();
    let staged = next.set_shape_text(follow_slide, 0, format!("0760 next {follow_slide}"));
    if !matches!(staged, Ok(true)) {
        emitter.emit(&format!("{step}.next.stage"), staged_text(&staged));
        return;
    }
    match next.commit() {
        Ok(commit) => {
            emitter.emit(&format!("{step}.next.commit"), commit_summary(&commit));
            match facade.apply_opened_presentation_commit(commit) {
                Ok(snapshot) => {
                    let bytes = bytes_text(&mut facade);
                    emitter.emit(
                        &format!("{step}.next.publish"),
                        format!("{} {bytes}", snapshot_summary(&snapshot)),
                    );
                },
                Err(error) => emitter.emit(&format!("{step}.next.publish"), error_text(&error)),
            }
        },
        Err(error) => emitter.emit(&format!("{step}.next.commit"), error_text(&error)),
    }
}

fn run_flows<W: Write>(emitter: &mut Emitter<'_, W>, make: Make<'_>) {
    let package = match make() {
        Ok(package) => {
            emitter.emit("open", "ok".into());
            package
        },
        Err(error) => {
            emitter.emit("open", error_text(&error));
            return;
        },
    };
    let snapshot = match package.opened_presentation() {
        Ok(snapshot) => {
            emitter.emit("capture", snapshot_summary(&snapshot));
            snapshot
        },
        Err(error) => {
            emitter.emit("capture", error_text(&error));
            return;
        },
    };
    let count = snapshot.slides().len();
    match snapshot.edit().commit() {
        Ok(commit) => {
            emitter.emit("noop.commit", commit_summary(&commit));
            let (text, _) = publish(make, commit);
            emitter.emit("noop.publish", text);
        },
        Err(error) => emitter.emit("noop.commit", error_text(&error)),
    }
    if count == 0 {
        return;
    }
    let mut positions = vec![0, count / 2, count - 1];
    positions.dedup();
    for slide in positions {
        let mut edit = snapshot.edit();
        let staged = edit.set_shape_text(slide, 0, format!("0760 diff {slide}"));
        if !matches!(staged, Ok(true)) {
            emitter.emit(&format!("edit{slide}.stage"), staged_text(&staged));
            continue;
        }
        commit_and_follow(emitter, make, &format!("edit{slide}"), edit, (slide + 1) % count);
    }
    // Every slide rewritten: no hits.
    let mut edit = snapshot.edit();
    let mut staged_all = true;
    for slide in 0..count.min(64) {
        let staged = edit.set_shape_text(slide, 0, format!("0760 all {slide}"));
        if !matches!(staged, Ok(true)) {
            emitter.emit(&format!("all{slide}.stage"), staged_text(&staged));
            staged_all = false;
            break;
        }
    }
    if staged_all {
        commit_and_follow(emitter, make, "all", edit, 0);
    }
    // Notes of the first slide: every slide payload is untouched.
    let mut edit = snapshot.edit();
    let staged = edit.set_notes_text(0, "0760 notes");
    if matches!(staged, Ok(true)) {
        commit_and_follow(emitter, make, "notes", edit, count / 2);
    } else {
        emitter.emit("notes.stage", staged_text(&staged));
    }
    // Removal of the middle slide's first shape.
    let mut edit = snapshot.edit();
    let staged = edit.remove_shape(count / 2, 0);
    if matches!(staged, Ok(true)) {
        commit_and_follow(emitter, make, "remove", edit, 0);
    } else {
        emitter.emit("remove.stage", staged_text(&staged));
    }
    // Reordering: slide payloads untouched, presentation part rewritten.
    if count > 1 {
        let mut edit = snapshot.edit();
        let staged = edit.move_slide(0, count - 1);
        if matches!(staged, Ok(true)) {
            commit_and_follow(emitter, make, "move", edit, 0);
        } else {
            emitter.emit("move.stage", staged_text(&staged));
        }
    }
    // MCE retention off: the slide-root memo is independent of it.
    match package.opened_presentation_with_limits(Limits::default().with_max_retained_mce_bytes(0)) {
        Ok(plain) => {
            emitter.emit("nomce.capture", snapshot_summary(&plain));
            let mut edit = plain.edit();
            let staged = edit.set_shape_text(count - 1, 0, "0760 nomce");
            if matches!(staged, Ok(true)) {
                commit_and_follow(emitter, make, "nomce", edit, 0);
            } else {
                emitter.emit("nomce.stage", staged_text(&staged));
            }
        },
        Err(error) => emitter.emit("nomce.capture", error_text(&error)),
    }
}

// ---------------------------------------------------------------------------
// Mutations.
// ---------------------------------------------------------------------------

const SLIDE_CT: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const P_TRANSITIONAL: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const P_STRICT: &str = "http://purl.oclc.org/ooxml/presentationml/main";

struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        let mut x = self.0;
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        self.0 = x;
        x
    }

    fn below(&mut self, bound: usize) -> usize {
        if bound == 0 { 0 } else { (self.next() % bound as u64) as usize }
    }
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).position(|window| window == needle)
}

fn rfind(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).rposition(|window| window == needle)
}

fn splice(xml: &[u8], start: usize, end: usize, with: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(xml.len() + with.len());
    out.extend_from_slice(&xml[..start]);
    out.extend_from_slice(with);
    out.extend_from_slice(&xml[end..]);
    out
}

/// End of the root start tag of a slide (`<p:sld ...>`), just after its `>`.
fn root_start_end(xml: &[u8]) -> Option<usize> {
    let root = find(xml, b"<p:sld")?;
    Some(root + find(&xml[root..], b">")? + 1)
}

fn replace_first_text(xml: &[u8], with: &[u8]) -> Option<Vec<u8>> {
    let start = find(xml, b"<a:t>")? + b"<a:t>".len();
    let end = start + find(&xml[start..], b"</a:t>")?;
    Some(splice(xml, start, end, with))
}

fn replace_all(xml: &[u8], from: &[u8], to: &[u8]) -> Option<Vec<u8>> {
    find(xml, from)?;
    let mut out = Vec::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(index) = find(rest, from) {
        out.extend_from_slice(&rest[..index]);
        out.extend_from_slice(to);
        rest = &rest[index + from.len()..];
    }
    out.extend_from_slice(rest);
    Some(out)
}

const STRUCTURED: &[&str] = &[
    "cdata-in-text",
    "unbound-prefix",
    "dtd",
    "pi-in-root",
    "comment-in-root",
    "strict-namespace",
    "wrong-root",
    "duplicate-attribute",
    "mce-alternate-content",
    "mce-ignorable-undeclared",
    "mce-ignorable-extension",
    "bom",
    "whitespace",
    "text-change",
    "empty",
    "truncate-half",
    "swap-with-next",
    "share-with-next",
];

fn structured(kind: &str, xml: &[u8]) -> Option<Vec<u8>> {
    match kind {
        "cdata-in-text" => replace_first_text(xml, b"<![CDATA[x]]>"),
        "unbound-prefix" => replace_first_text(xml, b"<q:x/>"),
        "dtd" => {
            let at = find(xml, b"?>").map_or(0, |index| index + 2);
            Some(splice(xml, at, at, b"<!DOCTYPE x>"))
        },
        "pi-in-root" => root_start_end(xml).map(|at| splice(xml, at, at, b"<?pi x?>")),
        "comment-in-root" => root_start_end(xml).map(|at| splice(xml, at, at, b"<!-- c -->")),
        "strict-namespace" => replace_all(xml, P_TRANSITIONAL.as_bytes(), P_STRICT.as_bytes()),
        "wrong-root" => {
            let open = find(xml, b"<p:sld ")?;
            let renamed = splice(xml, open, open + 7, b"<p:sldX ");
            let close = rfind(&renamed, b"</p:sld>")?;
            Some(splice(&renamed, close, close + 8, b"</p:sldX>"))
        },
        "duplicate-attribute" => {
            let start = find(xml, b"xmlns:a=\"")?;
            let end = start + 9 + find(&xml[start + 9..], b"\"")? + 1;
            let attribute = xml[start..end].to_vec();
            let mut with = attribute.clone();
            with.push(b' ');
            with.extend_from_slice(&attribute);
            Some(splice(xml, start, end, &with))
        },
        "mce-alternate-content" => {
            let start = find(xml, b"<p:sp>")?;
            let end = start + find(&xml[start..], b"</p:sp>")? + b"</p:sp>".len();
            let shape = &xml[start..end];
            let mut wrapped = format!(
                "<mc:AlternateContent xmlns:mc=\"{MC}\"><mc:Choice Requires=\"a\">"
            )
            .into_bytes();
            wrapped.extend_from_slice(shape);
            wrapped.extend_from_slice(b"</mc:Choice><mc:Fallback/></mc:AlternateContent>");
            Some(splice(xml, start, end, &wrapped))
        },
        "mce-ignorable-undeclared" => {
            let at = find(xml, b"<p:sld ")? + 7;
            Some(splice(
                xml,
                at,
                at,
                format!("xmlns:mc=\"{MC}\" mc:Ignorable=\"zz\" ").as_bytes(),
            ))
        },
        "mce-ignorable-extension" => {
            let at = find(xml, b"<p:sld ")? + 7;
            let declared = splice(
                xml,
                at,
                at,
                format!("xmlns:mc=\"{MC}\" xmlns:zz=\"urn:litchi:0760\" mc:Ignorable=\"zz\" ")
                    .as_bytes(),
            );
            let inner = root_start_end(&declared)?;
            Some(splice(&declared, inner, inner, b"<zz:extension zz:value=\"1\"/>"))
        },
        "bom" => {
            let mut out = vec![0xEF, 0xBB, 0xBF];
            out.extend_from_slice(xml);
            Some(out)
        },
        "whitespace" => {
            let mut out = Vec::with_capacity(xml.len() + 64);
            let mut inserted = 0;
            let mut depth_seen = false;
            for (index, &byte) in xml.iter().enumerate() {
                out.push(byte);
                if byte == b'>' && inserted < 16 && index + 1 < xml.len() && xml[index + 1] == b'<' {
                    if depth_seen {
                        out.extend_from_slice(b"\n  ");
                        inserted += 1;
                    }
                    depth_seen = true;
                }
            }
            Some(out)
        },
        "text-change" => replace_first_text(xml, b"mutated by 0760"),
        "empty" => Some(Vec::new()),
        "truncate-half" => Some(xml[..xml.len() / 2].to_vec()),
        _ => None,
    }
}

const SNIPPETS: &[&[u8]] = &[
    b"<",
    b">",
    b"&",
    b"\"",
    b"<![CDATA[",
    b"]]>",
    b"<?x?>",
    b"<!--",
    b"-->",
    b" xmlns:q=\"urn:q\"",
    b"<q:x/>",
    b"<a:t>",
    b"</a:t>",
    b" mc:Ignorable=\"q\"",
    b"<!DOCTYPE x>",
    b"&amp;",
    b"&#0;",
    b"\x00",
    b"\xFF",
    b"<p:sld>",
];

fn random_mutation(rng: &mut Rng, xml: &[u8]) -> (String, Vec<u8>) {
    let len = xml.len().max(1);
    match rng.below(6) {
        0 => {
            let at = rng.below(len).min(xml.len().saturating_sub(1));
            let mut out = xml.to_vec();
            if !out.is_empty() {
                out[at] ^= (rng.below(255) + 1) as u8;
            }
            (format!("flip@{at}"), out)
        },
        1 => {
            let at = rng.below(len).min(xml.len());
            let end = (at + 1 + rng.below(16)).min(xml.len());
            (format!("delete@{at}..{end}"), splice(xml, at, end, b""))
        },
        2 => {
            let at = rng.below(len).min(xml.len());
            let snippet = SNIPPETS[rng.below(SNIPPETS.len())];
            (format!("insert@{at}#{}", digest(snippet)), splice(xml, at, at, snippet))
        },
        3 => {
            let at = rng.below(len).min(xml.len());
            let end = (at + 1 + rng.below(64)).min(xml.len());
            let piece = xml[at..end].to_vec();
            let target = rng.below(len).min(xml.len());
            (format!("duplicate@{at}..{end}->{target}"), splice(xml, target, target, &piece))
        },
        4 => {
            let at = rng.below(len).min(xml.len());
            (format!("truncate@{at}"), xml[..at].to_vec())
        },
        _ => {
            let angles: Vec<usize> = xml
                .iter()
                .enumerate()
                .filter(|(_, byte)| **byte == b'<')
                .map(|(index, _)| index)
                .collect();
            if angles.is_empty() {
                return ("amp-none".into(), xml.to_vec());
            }
            let at = angles[rng.below(angles.len())];
            (format!("amp@{at}"), splice(xml, at, at + 1, b"&"))
        },
    }
}

fn slide_names(opc: &OpcPackage) -> Vec<PackURI> {
    let mut names: Vec<PackURI> = opc
        .iter_parts()
        .filter(|metadata| metadata.content_type() == SLIDE_CT)
        .map(|metadata| metadata.partname().clone())
        .collect();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    names
}

fn collect_pptx(directory: &std::path::Path, found: &mut Vec<std::path::PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_pptx(&path, found);
        } else if path.extension().is_some_and(|extension| extension == "pptx") {
            found.push(path);
        }
    }
}

fn diff_mode(root: &str, random_per_fixture: usize, seed: u64) {
    let root = std::path::Path::new(root);
    let mut fixtures = Vec::new();
    collect_pptx(root, &mut fixtures);
    fixtures.sort();
    let stdout = std::io::stdout();
    let mut out = std::io::BufWriter::new(stdout.lock());
    let mut packages = 0usize;
    let mut lines = 0usize;
    for (fixture_index, fixture) in fixtures.iter().enumerate() {
        let name = fixture
            .strip_prefix(root)
            .unwrap_or(fixture)
            .display()
            .to_string();
        let bytes = std::fs::read(fixture).unwrap();
        // The unmodified fixture through the ordinary byte ingress.
        {
            let make = || Package::from_bytes(&bytes);
            let mut emitter = Emitter { label: format!("{name}#original"), out: &mut out, lines: 0 };
            run_flows(&mut emitter, &make);
            lines += emitter.lines;
            packages += 1;
        }
        let Ok(opc) = OpcPackage::from_bytes(&bytes) else {
            continue;
        };
        let slides = slide_names(&opc);
        if slides.is_empty() {
            continue;
        }
        let mut positions = vec![0, slides.len() / 2, slides.len() - 1];
        positions.dedup();
        for &position in &positions {
            let target = &slides[position];
            let next = &slides[(position + 1) % slides.len()];
            let Ok(original) = opc.get_part(target).map(|part| part.blob().to_vec()) else {
                continue;
            };
            for kind in STRUCTURED {
                let mutate = |opc: &mut OpcPackage| -> bool {
                    match *kind {
                        "swap-with-next" => {
                            if target == next {
                                return false;
                            }
                            let first = opc.get_part(target).unwrap().blob_arc();
                            let second = opc.get_part(next).unwrap().blob_arc();
                            opc.get_part_mut(target).unwrap().set_blob_shared(second);
                            opc.get_part_mut(next).unwrap().set_blob_shared(first);
                            true
                        },
                        "share-with-next" => {
                            if target == next {
                                return false;
                            }
                            let shared = opc.get_part(target).unwrap().blob_arc();
                            opc.get_part_mut(next).unwrap().set_blob_shared(shared);
                            true
                        },
                        other => match structured(other, &original) {
                            Some(mutated) => {
                                opc.get_part_mut(target).unwrap().set_blob(mutated);
                                true
                            },
                            None => false,
                        },
                    }
                };
                let mut probe = opc.clone();
                if !mutate(&mut probe) {
                    continue;
                }
                let make = || {
                    let mut variant = opc.clone();
                    mutate(&mut variant);
                    Package::from_opc_package(variant)
                };
                let mut emitter = Emitter {
                    label: format!("{name}#s{position}#{kind}"),
                    out: &mut out,
                    lines: 0,
                };
                run_flows(&mut emitter, &make);
                lines += emitter.lines;
                packages += 1;
            }
        }
        let mut rng = Rng(seed ^ ((fixture_index as u64 + 1) * 0x9E37_79B9_7F4A_7C15));
        for _ in 0..random_per_fixture {
            let position = rng.below(slides.len());
            let target = slides[position].clone();
            let Ok(original) = opc.get_part(&target).map(|part| part.blob().to_vec()) else {
                continue;
            };
            let (kind, mutated) = random_mutation(&mut rng, &original);
            let make = || {
                let mut variant = opc.clone();
                variant.get_part_mut(&target)?.set_blob(mutated.clone());
                Package::from_opc_package(variant)
            };
            let mut emitter = Emitter {
                label: format!("{name}#s{position}#{kind}"),
                out: &mut out,
                lines: 0,
            };
            run_flows(&mut emitter, &make);
            lines += emitter.lines;
            packages += 1;
        }
    }
    out.flush().unwrap();
    eprintln!("diff fixtures={} packages={packages} lines={lines}", fixtures.len());
}

// ---------------------------------------------------------------------------
// Timing modes.
// ---------------------------------------------------------------------------

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let mode = args.get(1).map(String::as_str).unwrap_or("stats");
    if mode == "diff" {
        let root = args.get(2).expect("fixture root");
        let random: usize = args.get(3).map(|value| value.parse().unwrap()).unwrap_or(0);
        let seed: u64 = args.get(4).map(|value| value.parse().unwrap()).unwrap_or(760);
        diff_mode(root, random, seed);
        return;
    }
    let shape = args.get(2).map(String::as_str).unwrap_or("large");
    let iters: usize = args.get(3).map(|value| value.parse().unwrap()).unwrap_or(10);
    let edits = args.get(4).map(String::as_str).unwrap_or("one");
    let (slides, boxes) = shape_dims(shape);
    let bytes = build(slides, boxes);
    let total = slides * boxes;
    let selected: Vec<usize> = match edits {
        "noop" => Vec::new(),
        "one" => vec![update_indices(total)[0]],
        "pct" => update_indices(total),
        _ => Vec::new(),
    };
    match mode {
        "setup" => {
            // The per-iteration setup of the edit/save cycle, outside its
            // clocks: a fresh owned package from the archive bytes.
            for _ in 0..iters {
                let package = Package::from_vec(bytes.clone()).unwrap();
                std::hint::black_box(package);
            }
            println!("setup done");
        },
        "fulltext" => {
            let package = Package::from_bytes(&bytes).unwrap();
            let presentation = package.presentation().unwrap();
            let mut times = Vec::new();
            for _ in 0..iters {
                let started = Instant::now();
                let text = presentation.text().unwrap();
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(text);
            }
            println!("fulltext median_ns {}", median(times));
        },
        "capture" => {
            let package = Package::from_bytes(&bytes).unwrap();
            let mut times = Vec::new();
            for _ in 0..iters {
                let started = Instant::now();
                let snapshot = package.opened_presentation().unwrap();
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(snapshot);
            }
            println!("capture median_ns {}", median(times));
        },
        "commit" => {
            // Only `Transaction::commit`: the staged transaction is cloned
            // outside the clock each iteration.
            let package = Package::from_bytes(&bytes).unwrap();
            let snapshot = package.opened_presentation().unwrap();
            let mut edit = snapshot.edit();
            for linear in &selected {
                let slide = *linear / boxes;
                let object = *linear % boxes;
                edit.set_shape_text(slide, object, semantic_pptx_text(slide, object, true))
                    .unwrap();
            }
            let mut times = Vec::new();
            for _ in 0..iters {
                let staged = edit.clone();
                let started = Instant::now();
                let commit = staged.commit().unwrap();
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(commit);
            }
            println!("commit median_ns {}", median(times));
        },
        "cycle" => {
            let mut phase = [Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new()];
            let mut totals = Vec::new();
            let mut digest_bytes = None;
            for _ in 0..iters {
                let mut package = Package::from_vec(bytes.clone()).unwrap();
                let t0 = Instant::now();
                let snapshot = package.opened_presentation().unwrap();
                let t1 = Instant::now();
                let mut edit = snapshot.edit();
                let t2 = Instant::now();
                for linear in &selected {
                    let slide = *linear / boxes;
                    let object = *linear % boxes;
                    assert!(edit
                        .set_shape_text(slide, object, semantic_pptx_text(slide, object, true))
                        .unwrap());
                }
                let t3 = Instant::now();
                let commit = edit.commit().unwrap();
                let t4 = Instant::now();
                package.apply_opened_presentation_commit(commit).unwrap();
                let t5 = Instant::now();
                let out = package.to_bytes().unwrap();
                let t6 = Instant::now();
                phase[0].push((t1 - t0).as_nanos());
                phase[1].push((t2 - t1).as_nanos());
                phase[2].push((t3 - t2).as_nanos());
                phase[3].push((t4 - t3).as_nanos());
                phase[4].push((t5 - t4).as_nanos());
                phase[5].push((t6 - t5).as_nanos());
                totals.push((t6 - t0).as_nanos());
                match digest_bytes {
                    None => digest_bytes = Some(out),
                    Some(ref previous) => assert_eq!(previous, &out, "nondeterministic output"),
                }
            }
            let names = ["capture", "edit", "set_text", "commit", "apply", "to_bytes"];
            for (name, values) in names.iter().zip(phase) {
                println!("{name:10} median_ns {}", median(values));
            }
            println!("total      median_ns {}", median(totals));
            let out = digest_bytes.unwrap();
            let reopened = Package::from_bytes(&out).unwrap();
            let text = reopened.presentation().unwrap().text().unwrap();
            println!("output bytes {} text bytes {}", out.len(), text.len());
        },
        "alloc" => {
            // Allocation calls and requested bytes of exactly the timed
            // regions, on a warmed process.
            for (label, chosen) in [
                ("noop", Vec::new()),
                ("one", vec![update_indices(total)[0]]),
                ("pct", update_indices(total)),
            ] {
                let mut package = Package::from_vec(bytes.clone()).unwrap();
                let before = allocations();
                let snapshot = package.opened_presentation().unwrap();
                let captured = allocations();
                let mut edit = snapshot.edit();
                for linear in &chosen {
                    let slide = *linear / boxes;
                    let object = *linear % boxes;
                    edit.set_shape_text(slide, object, semantic_pptx_text(slide, object, true))
                        .unwrap();
                }
                let staged = allocations();
                let commit = edit.commit().unwrap();
                let committed = allocations();
                package.apply_opened_presentation_commit(commit).unwrap();
                let out = package.to_bytes().unwrap();
                let after = allocations();
                std::hint::black_box(out);
                println!(
                    "{label} total calls {} bytes {} | capture calls {} bytes {} | commit calls {} bytes {}",
                    after.0 - before.0,
                    after.1 - before.1,
                    captured.0 - before.0,
                    captured.1 - before.1,
                    committed.0 - staged.0,
                    committed.1 - staged.1,
                );
            }
        },
        "dump" => {
            // Published bytes, revisions, patches and full text, for a
            // byte-for-byte comparison of two builds.
            let directory = std::path::PathBuf::from(args.get(4).expect("dump directory"));
            std::fs::create_dir_all(&directory).unwrap();
            let package = Package::from_bytes(&bytes).unwrap();
            let text = package.presentation().unwrap().text().unwrap();
            std::fs::write(directory.join(format!("{shape}-fulltext.txt")), text).unwrap();
            for (label, chosen) in [
                ("noop", Vec::new()),
                ("one", vec![update_indices(total)[0]]),
                ("pct", update_indices(total)),
            ] {
                let mut package = Package::from_vec(bytes.clone()).unwrap();
                let mut edit = package.opened_presentation_transaction().unwrap();
                for linear in &chosen {
                    let slide = *linear / boxes;
                    let object = *linear % boxes;
                    edit.set_shape_text(slide, object, semantic_pptx_text(slide, object, true))
                        .unwrap();
                }
                let commit = edit.commit().unwrap();
                let patch = commit.patch().to_bytes().unwrap();
                let snapshot = package.apply_opened_presentation_commit(commit).unwrap();
                let out = package.to_bytes().unwrap();
                std::fs::write(directory.join(format!("{shape}-{label}.pptx")), &out).unwrap();
                std::fs::write(directory.join(format!("{shape}-{label}.patch")), &patch).unwrap();
                std::fs::write(
                    directory.join(format!("{shape}-{label}.revision")),
                    format!("{}\n", hex(&snapshot.revision())),
                )
                .unwrap();
            }
        },
        other => panic!("unknown mode {other}"),
    }
}
