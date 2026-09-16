//! `saveprobe` — corpus publication evidence for change 0653.
//!
//! One fixture per process. For each documented OOXML publication route this
//! binary knows about, it opens the fixture, optionally makes the one semantic
//! edit `tools/perf-baseline`'s `*_ordinary_save_edit` selectors make, publishes
//! to a fresh temporary file on disk, and then emits one TSV row per published
//! archive member plus one outcome row for the route.
//!
//! Nothing here is timed. The output is a deterministic function of the input
//! bytes, so the two legs can be compared row for row.
//!
//! Row kinds (first column):
//!
//! ```text
//! C  fixture  format  source-sha256  source-bytes  source-member-count
//! X  fixture  route   admitted|refused  detail
//! O  fixture  route   OK   published-bytes  published-member-count
//! O  fixture  route   ERR  debug  display
//! M  fixture  route   member  pub-sha256-uncompressed  pub-size-uncompressed
//!                            src-sha256-uncompressed-or-ABSENT  identical(0/1)
//!                            pub-sha256-compressed  pub-size-compressed
//! ```
//!
//! Setting `SAVEPROBE_KEEP=1` retains each route's published archive in the
//! scratch directory instead of deleting it, so a differing member's bytes can
//! be recovered for a canonical comparison.
//!
//! `identical(0/1)` compares the **uncompressed** member payload of the
//! published archive with the same-named member's **uncompressed** payload in
//! the source archive. `ABSENT` means the published archive has a member the
//! source archive does not, in which case `identical` is 0.

use std::collections::BTreeMap;
use std::fmt::Write as _;
use std::io::Read as _;
use std::panic::AssertUnwindSafe;
use std::path::{Path, PathBuf};

use sha2::{Digest, Sha256};

/// Text the edit routes write. Same marker string `tools/perf-baseline`'s
/// ordinary-save family uses, so the edit lands in the same place.
const EDIT_MARKER: &str = "litchi-perf-0638-ordinary-save";

/// Bounds on the PPTX shape search, copied from the harness.
const MAX_SLIDES: usize = 64;
const MAX_SHAPES: usize = 32;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() != 4 {
        eprintln!("usage: saveprobe <corpus-root> <fixture-rel-path> <scratch-dir>");
        std::process::exit(2);
    }
    let root = PathBuf::from(&args[1]);
    let rel = args[2].clone();
    let scratch = PathBuf::from(&args[3]);
    let fixture = root.join(&rel);

    let mut out = String::new();
    run_fixture(&fixture, &rel, &scratch, &mut out);
    print!("{out}");
}

// ---------------------------------------------------------------------------
// archive reading
// ---------------------------------------------------------------------------

#[derive(Clone)]
struct Member {
    /// sha256 of the uncompressed payload.
    sha_uncompressed: String,
    size_uncompressed: u64,
    /// sha256 of the raw stored (compressed) payload.
    sha_compressed: String,
    size_compressed: u64,
}

/// Reads every non-directory member of a zip archive held in memory.
fn read_members(bytes: &[u8]) -> Result<Vec<(String, Member)>, String> {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|error| error.to_string())?;
    let mut members = Vec::new();
    for index in 0..archive.len() {
        // The raw reader hands back the stored payload without inflating it.
        let (name, is_dir, sha_compressed, size_compressed) = {
            let mut raw = archive
                .by_index_raw(index)
                .map_err(|error| error.to_string())?;
            let name = raw.name().to_owned();
            let is_dir = raw.is_dir();
            let mut buffer = Vec::new();
            if !is_dir {
                raw.read_to_end(&mut buffer)
                    .map_err(|error| error.to_string())?;
            }
            let size = buffer.len() as u64;
            (name, is_dir, sha256_hex(&buffer), size)
        };
        if is_dir {
            continue;
        }
        let mut entry = archive
            .by_index(index)
            .map_err(|error| error.to_string())?;
        let mut buffer = Vec::new();
        entry
            .read_to_end(&mut buffer)
            .map_err(|error| error.to_string())?;
        members.push((
            name,
            Member {
                sha_uncompressed: sha256_hex(&buffer),
                size_uncompressed: buffer.len() as u64,
                sha_compressed,
                size_compressed,
            },
        ));
    }
    Ok(members)
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    let digest = hasher.finalize();
    let mut hex = String::with_capacity(64);
    for byte in digest {
        let _ = write!(hex, "{byte:02x}");
    }
    hex
}

/// Escapes a field so a row is always exactly one line with stable columns.
fn esc(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '\t' => out.push_str("\\t"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            _ => out.push(ch),
        }
    }
    out
}

// ---------------------------------------------------------------------------
// routes
// ---------------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq, Eq)]
enum Format {
    Docx,
    Xlsx,
    Pptx,
}

impl Format {
    fn as_str(self) -> &'static str {
        match self {
            Self::Docx => "DOCX",
            Self::Xlsx => "XLSX",
            Self::Pptx => "PPTX",
        }
    }

    fn main_part(self) -> &'static str {
        match self {
            Self::Docx => "word/document.xml",
            Self::Xlsx => "xl/workbook.xml",
            Self::Pptx => "ppt/presentation.xml",
        }
    }
}

/// A typed refusal, a panic, or any other non-success outcome of one route.
struct RouteError {
    debug: String,
    display: String,
}

impl RouteError {
    fn new(debug: String, display: String) -> Self {
        Self { debug, display }
    }
}

/// Runs `body`, converting a panic into a `RouteError` so that a route that
/// aborts on a hostile fixture is still recorded as an outcome on both legs.
fn guarded<T>(body: impl FnOnce() -> Result<T, RouteError>) -> Result<T, RouteError> {
    let hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(|_| {}));
    let outcome = std::panic::catch_unwind(AssertUnwindSafe(body));
    std::panic::set_hook(hook);
    match outcome {
        Ok(result) => result,
        Err(payload) => {
            let text = payload
                .downcast_ref::<&str>()
                .map(|s| (*s).to_owned())
                .or_else(|| payload.downcast_ref::<String>().cloned())
                .unwrap_or_else(|| "<non-string panic payload>".to_owned());
            Err(RouteError::new(
                format!("Panic({text:?})"),
                format!("panic: {text}"),
            ))
        }
    }
}

/// Whether one route's edit was admitted, and the detail either way.
enum EditOutcome {
    Admitted(String),
    Refused(String),
}

fn run_fixture(fixture: &Path, rel: &str, scratch: &Path, out: &mut String) {
    let source_bytes = match std::fs::read(fixture) {
        Ok(bytes) => bytes,
        Err(error) => {
            let _ = writeln!(out, "C\t{}\tUNREADABLE\t-\t0\t0\t{}", esc(rel), esc(&error.to_string()));
            return;
        }
    };
    let source_members = match read_members(&source_bytes) {
        Ok(members) => members,
        Err(error) => {
            let _ = writeln!(
                out,
                "C\t{}\tNOT-A-ZIP\t{}\t{}\t0\t{}",
                esc(rel),
                sha256_hex(&source_bytes),
                source_bytes.len(),
                esc(&error)
            );
            return;
        }
    };
    let index: BTreeMap<String, Member> = source_members.iter().cloned().collect();
    let format = [Format::Docx, Format::Xlsx, Format::Pptx]
        .into_iter()
        .find(|candidate| index.contains_key(candidate.main_part()));
    let Some(format) = format else {
        let _ = writeln!(
            out,
            "C\t{}\tUNCLASSIFIED\t{}\t{}\t{}\t-",
            esc(rel),
            sha256_hex(&source_bytes),
            source_bytes.len(),
            source_members.len()
        );
        return;
    };
    let _ = writeln!(
        out,
        "C\t{}\t{}\t{}\t{}\t{}\t-",
        esc(rel),
        format.as_str(),
        sha256_hex(&source_bytes),
        source_bytes.len(),
        source_members.len()
    );

    let routes: &[&str] = match format {
        Format::Docx => &[
            "docx_noop_save",
            "docx_edit_save",
            "docx_source_backed_document",
            "docx_source_backed_section_layout",
        ],
        Format::Xlsx => &[
            "xlsx_noop_save",
            "xlsx_edit_save",
            "xlsx_source_backed_cell_values",
            "xlsx_source_backed_tab_state",
        ],
        Format::Pptx => &[
            "pptx_noop_save",
            "pptx_edit_save",
            "pptx_source_backed_slide",
            "pptx_guides_publish",
        ],
    };

    for route in routes {
        let destination = scratch.join(format!(
            "{}-{}.out",
            sha256_hex(rel.as_bytes()).split_at(24).0,
            route
        ));
        let _ = std::fs::remove_file(&destination);
        let edit_and_publish = guarded(|| publish(route, fixture, &destination));
        match edit_and_publish {
            Ok(edit) => {
                match &edit {
                    EditOutcome::Admitted(detail) => {
                        let _ = writeln!(out, "X\t{}\t{route}\tadmitted\t{}", esc(rel), esc(detail));
                    }
                    EditOutcome::Refused(detail) => {
                        let _ = writeln!(out, "X\t{}\t{route}\trefused\t{}", esc(rel), esc(detail));
                    }
                }
                match std::fs::read(&destination) {
                    Ok(published) => {
                        match read_members(&published) {
                            Ok(members) => {
                                let _ = writeln!(
                                    out,
                                    "O\t{}\t{route}\tOK\t{}\t{}",
                                    esc(rel),
                                    published.len(),
                                    members.len()
                                );
                                for (name, member) in members {
                                    let source = index.get(&name);
                                    let (source_sha, identical) = match source {
                                        Some(source) => {
                                            let identical =
                                                usize::from(source.sha_uncompressed == member.sha_uncompressed);
                                            (source.sha_uncompressed.clone(), identical)
                                        }
                                        None => ("ABSENT".to_owned(), 0),
                                    };
                                    let _ = writeln!(
                                        out,
                                        "M\t{}\t{route}\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
                                        esc(rel),
                                        esc(&name),
                                        member.sha_uncompressed,
                                        member.size_uncompressed,
                                        source_sha,
                                        identical,
                                        member.sha_compressed,
                                        member.size_compressed
                                    );
                                }
                            }
                            Err(error) => {
                                let _ = writeln!(
                                    out,
                                    "O\t{}\t{route}\tERR\tPublishedArchiveUnreadable\t{}",
                                    esc(rel),
                                    esc(&error)
                                );
                            }
                        }
                    }
                    Err(error) => {
                        let _ = writeln!(
                            out,
                            "O\t{}\t{route}\tERR\tPublishedArchiveMissing\t{}",
                            esc(rel),
                            esc(&error.to_string())
                        );
                    }
                }
            }
            Err(error) => {
                let _ = writeln!(
                    out,
                    "O\t{}\t{route}\tERR\t{}\t{}",
                    esc(rel),
                    esc(&error.debug),
                    esc(&error.display)
                );
            }
        }
        if std::env::var_os("SAVEPROBE_KEEP").is_none() {
            let _ = std::fs::remove_file(&destination);
        }
    }
}

fn publish(route: &str, fixture: &Path, destination: &Path) -> Result<EditOutcome, RouteError> {
    match route {
        "docx_noop_save" => docx_save(fixture, destination, false),
        "docx_edit_save" => docx_save(fixture, destination, true),
        "docx_source_backed_document" => docx_source_backed_document(fixture, destination),
        "docx_source_backed_section_layout" => {
            docx_source_backed_section_layout(fixture, destination)
        }
        "xlsx_noop_save" => xlsx_save(fixture, destination, false),
        "xlsx_edit_save" => xlsx_save(fixture, destination, true),
        "xlsx_source_backed_cell_values" => xlsx_source_backed_cell_values(fixture, destination),
        "xlsx_source_backed_tab_state" => xlsx_source_backed_tab_state(fixture, destination),
        "pptx_noop_save" => pptx_save(fixture, destination, false),
        "pptx_edit_save" => pptx_save(fixture, destination, true),
        "pptx_source_backed_slide" => pptx_source_backed_slide(fixture, destination),
        "pptx_guides_publish" => pptx_guides_publish(fixture, destination),
        other => Err(RouteError::new(
            format!("UnknownRoute({other:?})"),
            format!("unknown route {other}"),
        )),
    }
}

fn typed<E: std::fmt::Debug + std::fmt::Display>(error: E) -> RouteError {
    RouteError::new(format!("{error:?}"), format!("{error}"))
}

// --- DOCX ------------------------------------------------------------------

fn docx_save(fixture: &Path, destination: &Path, edit: bool) -> Result<EditOutcome, RouteError> {
    let mut package = litchi_docx::Package::open(fixture).map_err(typed)?;
    let outcome = if edit {
        match package.document_mut() {
            Ok(document) => {
                document.add_paragraph_with_text(EDIT_MARKER);
                EditOutcome::Admitted("document_mut().add_paragraph_with_text".to_owned())
            }
            Err(error) => EditOutcome::Refused(format!("{error}")),
        }
    } else {
        EditOutcome::Admitted("no-edit".to_owned())
    };
    package.save(destination).map_err(typed)?;
    Ok(outcome)
}

fn docx_source_backed_document(fixture: &Path, destination: &Path) -> Result<EditOutcome, RouteError> {
    let package = litchi_docx::source_backed::Package::open(fixture).map_err(typed)?;
    let edit = package.edit_document().map_err(typed)?;
    let commit = edit.commit().map_err(typed)?;
    let file = std::fs::File::create(destination).map_err(typed)?;
    let writer = std::io::BufWriter::new(file);
    package
        .publish_document_commit_to_stream(writer, &commit)
        .map_err(typed)?;
    Ok(EditOutcome::Admitted(
        "edit_document().commit() with no operation".to_owned(),
    ))
}

fn docx_source_backed_section_layout(
    fixture: &Path,
    destination: &Path,
) -> Result<EditOutcome, RouteError> {
    let package = litchi_docx::source_backed::Package::open(fixture).map_err(typed)?;
    // Section position 0; the publisher overlays only the main-document part
    // and copies every other ZIP member through the preservation plan.
    let edit = package.edit_section_layout(0usize).map_err(typed)?;
    let commit = edit.commit().map_err(typed)?;
    let file = std::fs::File::create(destination).map_err(typed)?;
    let writer = std::io::BufWriter::new(file);
    package
        .publish_section_layout_commit_to_stream(writer, &commit)
        .map_err(typed)?;
    Ok(EditOutcome::Admitted(
        "edit_section_layout(0).commit() with no operation".to_owned(),
    ))
}

// --- XLSX ------------------------------------------------------------------

fn xlsx_save(fixture: &Path, destination: &Path, edit: bool) -> Result<EditOutcome, RouteError> {
    let workbook = litchi_xlsx::Workbook::open(fixture).map_err(typed)?;
    let (workbook, outcome) = if edit {
        let Some(sheet_name) = workbook.sheets().next().map(|sheet| sheet.name().to_owned()) else {
            return Err(RouteError::new(
                "NoWorksheet".to_owned(),
                "the workbook declares no worksheet".to_owned(),
            ));
        };
        let attempt = (|| -> Result<litchi_xlsx::Workbook, String> {
            let mut edit = workbook.edit().map_err(|error| error.to_string())?;
            {
                let mut sheet = edit
                    .sheet(sheet_name.as_str())
                    .map_err(|error| error.to_string())?
                    .ok_or_else(|| "the selected worksheet is absent".to_owned())?;
                sheet
                    .set("A1", EDIT_MARKER)
                    .map_err(|error| error.to_string())?;
            }
            let commit = edit.commit().map_err(|error| error.to_string())?;
            if commit.patch().is_empty() {
                return Err("the edit produced an empty patch".to_owned());
            }
            Ok(commit.into_workbook())
        })();
        match attempt {
            Ok(edited) => (
                edited,
                EditOutcome::Admitted(format!("edit().sheet({sheet_name:?}).set(\"A1\")")),
            ),
            Err(error) => (workbook, EditOutcome::Refused(error)),
        }
    } else {
        (workbook, EditOutcome::Admitted("no-edit".to_owned()))
    };
    workbook.save(destination).map_err(typed)?;
    Ok(outcome)
}

fn xlsx_source_backed_cell_values(fixture: &Path, destination: &Path) -> Result<EditOutcome, RouteError> {
    let editor = litchi_xlsx::cell_values::SourceBackedEditor::open(fixture).map_err(typed)?;
    // Position 0 is the first worksheet the reader reports; the route is the
    // documented exact-source-checked value-only overlay publisher.
    let edit = editor.edit(0usize).map_err(typed)?;
    let commit = edit.commit().map_err(typed)?;
    let file = std::fs::File::create(destination).map_err(typed)?;
    let writer = std::io::BufWriter::new(file);
    editor
        .publish_commit_to_stream(writer, &commit)
        .map_err(typed)?;
    Ok(EditOutcome::Admitted(
        "cell_values::SourceBackedEditor::edit(0).commit() with no operation".to_owned(),
    ))
}

fn xlsx_source_backed_tab_state(
    fixture: &Path,
    destination: &Path,
) -> Result<EditOutcome, RouteError> {
    let package = litchi_opc::SourceBackedPackage::from_path(fixture).map_err(typed)?;
    let editor =
        litchi_xlsx::tab_state::SourceBackedEditor::from_source_backed_package(package)
            .map_err(typed)?;
    let commit = {
        let edit = editor.edit().map_err(typed)?;
        edit.commit().map_err(typed)?
    };
    let file = std::fs::File::create(destination).map_err(typed)?;
    let writer = std::io::BufWriter::new(file);
    editor
        .publish_commit_to_stream(writer, &commit)
        .map_err(typed)?;
    Ok(EditOutcome::Admitted(
        "tab_state::SourceBackedEditor::edit().commit() with no operation".to_owned(),
    ))
}

// --- PPTX ------------------------------------------------------------------

fn pptx_save(fixture: &Path, destination: &Path, edit: bool) -> Result<EditOutcome, RouteError> {
    let mut package = litchi_pptx::Package::open(fixture).map_err(typed)?;
    let outcome = if edit {
        match pptx_apply_edit(&mut package) {
            Ok(detail) => EditOutcome::Admitted(detail),
            Err(detail) => EditOutcome::Refused(detail),
        }
    } else {
        EditOutcome::Admitted("no-edit".to_owned())
    };
    package.save(destination).map_err(typed)?;
    Ok(outcome)
}

/// Finds the first `(slide, shape)` position the documented opened-presentation
/// transaction admits, then applies the marker there. This is the harness's own
/// derivation, bounded the same way.
fn pptx_apply_edit(package: &mut litchi_pptx::Package) -> Result<String, String> {
    let slide_count = match package.opened_presentation() {
        Ok(snapshot) => snapshot.slides().len().min(MAX_SLIDES),
        Err(error) => return Err(format!("{error}")),
    };
    let mut first_refusal: Option<String> = None;
    let mut target = None;
    'search: for slide in 0..slide_count {
        for shape in 0..MAX_SHAPES {
            let Ok(mut transaction) = package.opened_presentation_transaction() else {
                continue;
            };
            match transaction.set_shape_text(slide, shape, EDIT_MARKER) {
                Ok(true) => {
                    target = Some((slide, shape));
                    break 'search;
                }
                Ok(false) => continue,
                Err(error) => {
                    if first_refusal.is_none() {
                        first_refusal = Some(format!("{error}"));
                    }
                }
            }
        }
    }
    let Some((slide, shape)) = target else {
        return Err(first_refusal.unwrap_or_else(|| "no admitted shape position".to_owned()));
    };
    let mut transaction = package
        .opened_presentation_transaction()
        .map_err(|error| format!("{error}"))?;
    if !transaction
        .set_shape_text(slide, shape, EDIT_MARKER)
        .map_err(|error| format!("{error}"))?
    {
        return Err("the derived shape position reported no change".to_owned());
    }
    let commit = transaction.commit().map_err(|error| format!("{error}"))?;
    if !commit.is_changed() {
        return Err("the commit reports no change".to_owned());
    }
    package
        .apply_opened_presentation_commit(commit)
        .map_err(|error| format!("{error}"))?;
    Ok(format!("set_shape_text({slide}, {shape})"))
}

fn pptx_source_backed_slide(fixture: &Path, destination: &Path) -> Result<EditOutcome, RouteError> {
    let editor = litchi_pptx::SourceBackedPresentationEditor::open(fixture).map_err(typed)?;
    // Position 0 is the first slide; the route is the documented
    // exact-source-checked slide overlay publisher, whose own documentation
    // states that a no-op reproduces the whole source artifact byte for byte.
    let edit = editor.edit_slide(0).map_err(typed)?;
    let commit = edit.commit();
    let file = std::fs::File::create(destination).map_err(typed)?;
    let writer = std::io::BufWriter::new(file);
    editor
        .publish_slide_commit_to_stream(writer, &commit)
        .map_err(typed)?;
    Ok(EditOutcome::Admitted(
        "SourceBackedPresentationEditor::edit_slide(0).commit() with no operation".to_owned(),
    ))
}

/// The one documented route in this corpus's formats that publishes markup-
/// compatibility **codec output**: `guides::Snapshot::load` runs
/// `litchi_ooxml_common::mce::process_ooxml` over the whole
/// `ppt/presentation.xml`, and `Transaction::commit` runs it again and splices
/// the processed buffer, so the published presentation part carries exactly
/// what the codec emitted. (The other codec-output publisher named by change
/// 0653, `litchi-xlsb`'s drawing-anchor transfer, has no fixture in this
/// corpus, which is `.docx`/`.xlsx`/`.pptx` only.)
fn pptx_guides_publish(fixture: &Path, destination: &Path) -> Result<EditOutcome, RouteError> {
    use litchi_pptx::presentation_properties::metadata::guides;

    let mut package = litchi_opc::OpcPackage::open(fixture).map_err(typed)?;
    let snapshot = guides::Snapshot::load(&package).map_err(typed)?;
    let mut edit = snapshot.edit();
    let guide = guides::Guide {
        id: 900,
        name: Some("litchi-perf-0653".to_owned()),
        orientation: Some(guides::Orientation::Vertical),
        position: Some(12),
        user_drawn: Some(false),
        color: guides::Color {
            kind: guides::ColorKind::Srgb,
            xml: br#"<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="DDEEFF"/>"#
                .to_vec(),
        },
        extension_xml: None,
    };
    edit.push(guides::ListKind::Slide, guide).map_err(typed)?;
    let commit = edit.commit().map_err(typed)?;
    if !commit.changed() {
        return Err(RouteError::new(
            "GuidesCommitUnchanged".to_owned(),
            "the guides commit reports no change".to_owned(),
        ));
    }
    guides::apply_commit(&mut package, commit).map_err(typed)?;
    let file = std::fs::File::create(destination).map_err(typed)?;
    let writer = std::io::BufWriter::new(file);
    litchi_opc::PackageWriter::write_to_stream(writer, &package).map_err(typed)?;
    Ok(EditOutcome::Admitted(
        "guides::Snapshot::load().edit().push(Slide, guide).commit() then PackageWriter".to_owned(),
    ))
}
