//! Full-corpus fidelity regression (record 0289, plan step 26, design §14
//! layer 3).
//!
//! Every file under `test-data/` that the unified layer opens as a claimed
//! format (the iWork formats since record 0296) goes through `open` (file-backed, the retained source)
//! → `edit` with no mutation → `save_as`, and the saved bytes must equal
//! the input bytes. Signed files cannot be edited without consuming their
//! signature, so they go through `open` → `save_as` instead; locked
//! (encrypted) files need a password and are only counted. Where it is
//! cheap, a set-then-revert round trip follows: the first node (in a
//! bounded breadth-first walk) with a settable text property is set to a
//! new value and back, and the reopened value must equal the original.
//! Whether the reverted bytes also equal the input is reported, not
//! required: a revert is a second edit, and the leaves may normalize what
//! they rewrite. A structural round trip follows on XLS, PPTX and RTF
//! (record 0293) and Pages (record 0299): a row inserted and removed, a
//! slide or paragraph moved and moved back, or a Pages paragraph added and
//! removed, must read back with the same node texts.
//!
//! An empty edit the leaf editor refuses (typed) is reported with its
//! cause, and the file then goes through `open` → `save_as`, which must be
//! exact too. The whole corpus runs in about 20 s in a debug build, so no
//! subset is needed.
//!
//! The generated report is golden-tested:
//! `docs/unified-ops/results/0289/fidelity.{md,tsv}`; rewrite it with
//! `LITCHI_BLESS=1` after showing that a change is intended. A save that is
//! not byte-identical, an untyped failure (`internal_error`, a panic) or a
//! revert that does not restore the value fails the test, unless the format
//! is listed in [`NORMALIZING`] with its justification.

#![cfg(all(
    feature = "doc",
    feature = "docx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ppt",
    feature = "pptx",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "odg",
    feature = "odc",
    feature = "odi",
    feature = "odb",
    feature = "oth",
    feature = "odm",
    feature = "odf-formula",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]

use std::collections::BTreeMap;
use std::fmt::Write as _;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};

use litchi_core::FileFormat;
use litchi_ops::value::ValueType;
use litchi_ops::{File, NewNode, Placement, Props, open};

/// Formats whose no-op save is allowed to differ from the input, with the
/// documented normalization. Empty: every format must be exact.
const NORMALIZING: &[(&str, &str)] = &[];

fn test_data() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
}

fn name(format: FileFormat) -> &'static str {
    match format {
        FileFormat::Doc => "doc",
        FileFormat::Docx => "docx",
        FileFormat::Xls => "xls",
        FileFormat::Xlsx => "xlsx",
        FileFormat::Xlsb => "xlsb",
        FileFormat::Ppt => "ppt",
        FileFormat::Pptx => "pptx",
        FileFormat::Rtf => "rtf",
        FileFormat::Odt => "odt",
        FileFormat::Ods => "ods",
        FileFormat::Odp => "odp",
        FileFormat::Odg => "odg",
        FileFormat::Odc => "odc",
        FileFormat::Odi => "odi",
        FileFormat::Odb => "odb",
        FileFormat::Oth => "oth",
        FileFormat::Odm => "odm",
        FileFormat::Odf => "odf",
        FileFormat::Pages => "pages",
        FileFormat::Keynote => "keynote",
        FileFormat::Numbers => "numbers",
        _ => "other",
    }
}

fn guarded<T>(f: impl FnOnce() -> litchi_ops::Result<T>) -> Result<T, String> {
    match catch_unwind(AssertUnwindSafe(f)) {
        Ok(Ok(value)) => Ok(value),
        Ok(Err(error)) => Err(format!("{}: {}", error.code().as_str(), error.message())),
        Err(panic) => {
            let message = panic
                .downcast_ref::<&str>()
                .map(|text| (*text).to_owned())
                .or_else(|| panic.downcast_ref::<String>().cloned())
                .unwrap_or_default();
            Err(format!("panic: {message}"))
        },
    }
}

/// Keeps an evidence line short and free of table separators.
fn clean(text: &str) -> String {
    let text = text.replace(['\t', '\n', '|'], " ");
    if text.chars().count() > 140 {
        let mut short: String = text.chars().take(137).collect();
        short.push_str("...");
        short
    } else {
        text
    }
}

/// Every corpus file, sorted, with its size.
fn corpus() -> Vec<(String, u64)> {
    let root = test_data();
    let mut found = Vec::new();
    let mut stack = vec![root.clone()];
    while let Some(dir) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else if let Ok(meta) = entry.metadata() {
                let relative = path
                    .strip_prefix(&root)
                    .unwrap()
                    .to_string_lossy()
                    .into_owned();
                found.push((relative, meta.len()));
            }
        }
    }
    found.sort();
    found
}

/// The outcome of one file.
struct Outcome {
    format: &'static str,
    /// `exact`, `differs`, `refused` (typed edit refusal; `detail` holds
    /// the save-only result), `error` (untyped), `signed-exact`,
    /// `signed-differs`, `signed-error` or `locked`.
    noop: &'static str,
    detail: String,
    /// `byte-exact`, `value-exact`, `mismatch`, `one-way` (the set back is
    /// refused, typed), `refused`, `none` (no settable text within the
    /// walk) or `-` (not attempted).
    revert: &'static str,
    revert_detail: String,
    /// Structural round trip (record 0293): `byte-exact`, `value-exact`,
    /// `mismatch`, `one-way` (the inverse is refused, typed), `refused`,
    /// `none` (nothing to move) or `-` (not attempted for the format).
    structure: &'static str,
    structure_detail: String,
}

fn first_difference(a: &[u8], b: &[u8]) -> String {
    let offset = a
        .iter()
        .zip(b)
        .position(|(x, y)| x != y)
        .unwrap_or(a.len().min(b.len()));
    format!(
        "{} -> {} bytes, first difference at {offset}",
        a.len(),
        b.len()
    )
}

fn untyped(error: &str) -> bool {
    error.starts_with("internal_error") || error.starts_with("panic")
}

fn save_bytes(file: &File, dir: &Path, extension: &str) -> Result<Vec<u8>, String> {
    let path = dir.join(format!("saved.{extension}"));
    let _ = std::fs::remove_file(&path);
    guarded(|| file.save_as(&path))?;
    std::fs::read(&path).map_err(|error| error.to_string())
}

/// The value of `key` on the node at `path`: the node text for `text`,
/// otherwise the printed property.
fn current(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).ok()?;
    let view = node.view();
    if key == "text" {
        return view.text().ok()?.map(std::borrow::Cow::into_owned);
    }
    let format = view.format().ok()?;
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

/// Nodes in a bounded breadth-first walk with a settable, readable text
/// property and a non-empty value: path, key and value, at most six.
fn settable_texts(file: &File) -> Vec<(String, String, String)> {
    let caps = file.capabilities();
    let mut found = Vec::new();
    let Ok(root) = file.root() else {
        return found;
    };
    let mut queue = std::collections::VecDeque::from([root]);
    let mut seen = 0usize;
    while let Some(node) = queue.pop_front() {
        if let Some(element) = caps.element(node.kind()) {
            let path = node.path().to_string();
            for property in element.properties() {
                if !property.set()
                    || !property.get()
                    || !matches!(property.value_type(), ValueType::Text)
                {
                    continue;
                }
                if let Some(value) = current(file, &path, property.name())
                    .filter(|value| !value.is_empty() && value != "Litchi" && !value.contains('\n'))
                {
                    found.push((path.clone(), property.name().to_owned(), value));
                    if found.len() == 6 {
                        return found;
                    }
                    break;
                }
            }
        }
        for child in node.children() {
            let Ok(child) = child else { break };
            seen += 1;
            if seen > 400 {
                return found;
            }
            queue.push_back(child);
        }
    }
    found
}

/// Sets one text property to a new value and back on the first candidate
/// whose set is admitted.
fn revert(file: &File, input: &[u8]) -> (&'static str, String) {
    let candidates = catch_unwind(AssertUnwindSafe(|| settable_texts(file))).unwrap_or_default();
    if candidates.is_empty() {
        return ("none", String::new());
    }
    let mut refusals = Vec::new();
    for (path, key, original) in candidates {
        let changed =
            match guarded(|| file.set(path.as_str(), Props::new().with(key.as_str(), "Litchi"))) {
                Ok(changed) => changed,
                Err(error) => {
                    if untyped(&error) {
                        return ("mismatch", format!("{path} {key}: {error}"));
                    }
                    refusals.push(format!("{path} {key}: {error}"));
                    continue;
                },
            };
        let result = guarded(|| {
            let reverted = changed.set(
                path.as_str(),
                Props::new().with(key.as_str(), original.as_str()),
            )?;
            let bytes = reverted.to_bytes()?;
            let reopened = litchi_ops::open_bytes(bytes.clone())?;
            Ok((bytes, current(&reopened, &path, &key)))
        });
        return match result {
            Ok((bytes, Some(value))) if value == original => {
                if bytes == input {
                    ("byte-exact", format!("{path} {key}"))
                } else {
                    (
                        "value-exact",
                        format!("{path} {key}: {}", first_difference(input, &bytes)),
                    )
                }
            },
            Ok((_, value)) => (
                "mismatch",
                format!("{path} {key}: reopened {value:?}, expected {original:?}"),
            ),
            Err(error) if untyped(&error) => ("mismatch", format!("{path} {key}: revert {error}")),
            Err(error) => (
                "one-way",
                format!("{path} {key}: the set back is refused: {error}"),
            ),
        };
    }
    ("refused", refusals.first().cloned().unwrap_or_default())
}

/// Texts of the nodes at `paths`, for the value check of a structural round
/// trip.
fn texts(file: &File, paths: &[&str]) -> Vec<Option<String>> {
    paths
        .iter()
        .map(|path| current(file, path, "text"))
        .collect()
}

/// A structural edit and its inverse on the formats whose structure record
/// 0293 extended to producer files: XLS inserts a row at the top of the
/// first sheet and removes it again; PPTX moves the first slide after the
/// second and back; RTF moves the first body paragraph after the second and
/// back. The reverted file must read back with the same node texts; whether
/// its bytes equal the input is reported.
fn structure(file: &File, input: &[u8]) -> (&'static str, String) {
    type Forward = fn(&File) -> litchi_ops::Result<File>;
    type Inverse = fn(&File) -> litchi_ops::Result<File>;
    let count = |parent: &str, kind: &str| {
        file.get(parent).map_or(0, |node| {
            node.children()
                .filter_map(Result::ok)
                .filter(|child| child.kind() == kind)
                .count()
        })
    };
    let (label, checked, forward, inverse): (&str, &[&str], Forward, Inverse) = match file.format()
    {
        Some(FileFormat::Xls) => {
            if count("/", "sheet") == 0 {
                return ("none", String::new());
            }
            (
                "row add+remove",
                &["/sheet[1]/row[1]"],
                |file| {
                    file.add("/sheet[1]", NewNode::new("row").at(Placement::Start))
                        .map(|added| added.file)
                },
                |file| file.remove("/sheet[1]/row[1]"),
            )
        },
        Some(FileFormat::Pptx) => {
            if count("/", "slide") < 2 {
                return ("none", String::new());
            }
            (
                "slide move+back",
                &["/slide[1]", "/slide[2]"],
                |file| file.move_to("/slide[1]", Placement::after("/slide[2]")?),
                |file| file.move_to("/slide[2]", Placement::Start),
            )
        },
        Some(FileFormat::Rtf) => {
            if count("/body", "paragraph") < 2 {
                return ("none", String::new());
            }
            (
                "paragraph move+back",
                &["/body/p[1]", "/body/p[2]"],
                |file| file.move_to("/body/p[1]", Placement::after("/body/p[2]")?),
                |file| file.move_to("/body/p[2]", Placement::before("/body/p[1]")?),
            )
        },
        // Record 0299: a Pages paragraph added at the start of the body and
        // removed again, through the leaf paragraph transaction.
        Some(FileFormat::Pages) => (
            "paragraph add+remove",
            &["/body/p[1]"],
            |file| {
                file.add(
                    "/body",
                    NewNode::new("p")
                        .prop("text", "Structural probe")
                        .at(Placement::Start),
                )
                .map(|added| added.file)
            },
            |file| file.remove("/body/p[1]"),
        ),
        _ => return ("-", String::new()),
    };
    let original = catch_unwind(AssertUnwindSafe(|| texts(file, checked))).unwrap_or_default();
    let changed = match guarded(|| forward(file)) {
        Ok(changed) => changed,
        Err(error) if untyped(&error) => return ("mismatch", format!("{label}: {error}")),
        Err(error) => return ("refused", format!("{label}: {error}")),
    };
    let result = guarded(|| {
        let reverted = inverse(&changed)?;
        let bytes = reverted.to_bytes()?;
        let reopened = litchi_ops::open_bytes(bytes.clone())?;
        Ok((bytes, texts(&reopened, checked)))
    });
    match result {
        Ok((bytes, values)) if values == original => {
            if bytes == input {
                ("byte-exact", label.to_owned())
            } else {
                (
                    "value-exact",
                    format!("{label}: {}", first_difference(input, &bytes)),
                )
            }
        },
        Ok((_, values)) => (
            "mismatch",
            format!("{label}: reopened {values:?}, expected {original:?}"),
        ),
        Err(error) if untyped(&error) => ("mismatch", format!("{label}: inverse {error}")),
        Err(error) => (
            "one-way",
            format!("{label}: the inverse is refused: {error}"),
        ),
    }
}

/// The result of one file, or the open error for a file the unified layer
/// does not open as a claimed format.
fn run(relative: &str, dir: &Path) -> Result<Outcome, String> {
    let path = test_data().join(relative);
    let input = std::fs::read(&path).map_err(|error| error.to_string())?;
    let file = guarded(|| open(&path))?;
    let format = file.format().map_or("other", name);
    if format == "other" {
        return Err("not a claimed format".to_owned());
    }
    let extension = relative
        .rsplit('.')
        .next()
        .unwrap_or("bin")
        .to_ascii_lowercase();
    let outcome = |noop, detail, revert, revert_detail| Outcome {
        format,
        noop,
        detail,
        revert,
        revert_detail,
        structure: "-",
        structure_detail: String::new(),
    };
    if let File::Locked(_) = file {
        return Ok(outcome("locked", String::new(), "-", String::new()));
    }
    if file.is_signed() {
        let (noop, detail) = match save_bytes(&file, dir, &extension) {
            Ok(bytes) if bytes == input => ("signed-exact", String::new()),
            Ok(bytes) => ("signed-differs", first_difference(&input, &bytes)),
            Err(error) => ("signed-error", error),
        };
        return Ok(outcome(noop, detail, "-", String::new()));
    }
    let (noop, detail) = match guarded(|| file.edit(|_| Ok(()))) {
        Err(error) if untyped(&error) => ("error", format!("edit: {error}")),
        Err(error) => {
            let save = match save_bytes(&file, dir, &extension) {
                Ok(bytes) if bytes == input => "save-only exact".to_owned(),
                Ok(bytes) => format!("save-only DIFFERS {}", first_difference(&input, &bytes)),
                Err(save) => format!("save-only refused {save}"),
            };
            ("refused", format!("edit: {error}; {save}"))
        },
        Ok(commit) => match save_bytes(&commit.file, dir, &extension) {
            Ok(bytes) if bytes == input => ("exact", String::new()),
            Ok(bytes) => ("differs", first_difference(&input, &bytes)),
            Err(error) if untyped(&error) => ("error", format!("save: {error}")),
            Err(error) => ("refused", format!("save: {error}")),
        },
    };
    let (revert, revert_detail) = if noop == "exact" {
        revert(&file, &input)
    } else {
        ("-", String::new())
    };
    let (structure, structure_detail) = if noop == "exact" {
        structure(&file, &input)
    } else {
        ("-", String::new())
    };
    Ok(Outcome {
        structure,
        structure_detail,
        ..outcome(noop, detail, revert, revert_detail)
    })
}

/// Extensions of the claimed formats and their flat and template variants,
/// for the "not opened" table.
const OFFICE: &[&str] = &[
    "doc", "dot", "docx", "docm", "dotx", "dotm", "xls", "xlt", "xlsx", "xlsm", "xltx", "xltm",
    "xlsb", "ppt", "pot", "pptx", "pptm", "potx", "potm", "ppsx", "rtf", "odt", "ott", "fodt",
    "ods", "ots", "fods", "odp", "otp", "fodp", "odg", "otg", "fodg", "odc", "odi", "odb", "oth",
    "odm", "odf", "pages", "key", "numbers",
];

#[test]
#[allow(
    clippy::too_many_lines,
    reason = "runs the corpus and renders the report"
)]
fn open_empty_edit_save_is_byte_identical() {
    let dir = tempfile::tempdir().unwrap();
    let started = std::time::Instant::now();
    let mut outcomes: Vec<(String, Outcome)> = Vec::new();
    let mut not_opened: BTreeMap<String, BTreeMap<String, usize>> = BTreeMap::new();
    for (relative, _) in corpus() {
        match run(&relative, dir.path()) {
            Ok(outcome) => outcomes.push((relative, outcome)),
            Err(error) => {
                let extension = relative
                    .rsplit('.')
                    .next()
                    .unwrap_or("")
                    .to_ascii_lowercase();
                if OFFICE.contains(&extension.as_str()) {
                    let code = error.split(':').next().unwrap_or("").to_owned();
                    *not_opened
                        .entry(extension)
                        .or_default()
                        .entry(code)
                        .or_default() += 1;
                }
            },
        }
    }
    let elapsed = started.elapsed();

    let mut counts: BTreeMap<&str, BTreeMap<&str, usize>> = BTreeMap::new();
    for (_, outcome) in &outcomes {
        let entry = counts.entry(outcome.format).or_default();
        *entry.entry(outcome.noop).or_default() += 1;
        *entry
            .entry(match outcome.revert {
                "byte-exact" => "revert byte-exact",
                "value-exact" => "revert value-exact",
                "mismatch" => "revert mismatch",
                "one-way" => "revert one-way",
                "refused" => "revert refused",
                "none" => "revert none",
                _ => "revert -",
            })
            .or_default() += 1;
        if outcome.structure != "-" {
            *entry
                .entry(match outcome.structure {
                    "byte-exact" => "structure byte-exact",
                    "value-exact" => "structure value-exact",
                    "mismatch" => "structure mismatch",
                    "one-way" => "structure one-way",
                    "refused" => "structure refused",
                    _ => "structure none",
                })
                .or_default() += 1;
        }
    }
    let columns = [
        "exact",
        "differs",
        "refused",
        "error",
        "signed-exact",
        "signed-differs",
        "signed-error",
        "locked",
        "revert byte-exact",
        "revert value-exact",
        "revert mismatch",
        "revert one-way",
        "revert refused",
        "revert none",
        "structure byte-exact",
        "structure value-exact",
        "structure mismatch",
        "structure one-way",
        "structure refused",
        "structure none",
    ];
    let mut markdown = format!(
        "# Corpus fidelity (records 0289 and 0293, generated)\n\n\
         Generated by `crates/litchi-ops/tests/corpus_fidelity.rs` over every file\n\
         under `test-data/` that opens as a claimed format; rewrite with\n\
         `LITCHI_BLESS=1`. `exact`: open → empty edit → save is byte-identical;\n\
         `refused`: the leaf editor refuses even an empty edit (typed), and the\n\
         detail gives the open → save result; `signed-*`: open → save (signed\n\
         files refuse edits); `locked`: encrypted, not opened further. Revert:\n\
         set one text property and set it back; `byte-exact` also matches the\n\
         input bytes, `value-exact` only the value, `one-way` the set back was\n\
         refused (typed), `refused` every candidate set was refused (typed),\n\
         `none` no settable text within 400 nodes. Structure (record 0293,\n\
         XLS, PPTX and RTF only): insert a row at the top of the first sheet\n\
         and remove it, or move the first slide or body paragraph after the\n\
         second and back; `byte-exact` also matches the input bytes,\n\
         `value-exact` only the moved nodes' text, `refused` the forward edit\n\
         was refused (typed), `none` nothing to move.\n\n\
         | format | {} |\n|---|{}\n",
        columns.join(" | "),
        "---|".repeat(columns.len())
    );
    let mut totals = vec![0usize; columns.len()];
    for (format, entry) in &counts {
        let row: Vec<String> = columns
            .iter()
            .enumerate()
            .map(|(index, column)| {
                let count = entry.get(column).copied().unwrap_or(0);
                totals[index] += count;
                count.to_string()
            })
            .collect();
        let _ = writeln!(markdown, "| {format} | {} |", row.join(" | "));
    }
    let _ = writeln!(
        markdown,
        "| **all** | {} |",
        totals
            .iter()
            .map(ToString::to_string)
            .collect::<Vec<_>>()
            .join(" | ")
    );
    markdown.push_str(
        "\n## Office-extension files the unified layer does not open\n\n\
         | extension | open error codes |\n|---|---|\n",
    );
    for (extension, codes) in &not_opened {
        let codes: Vec<String> = codes
            .iter()
            .map(|(code, count)| format!("{code} ×{count}"))
            .collect();
        let _ = writeln!(markdown, "| {extension} | {} |", codes.join(", "));
    }
    markdown.push_str("\n## Every no-op that is not exact, and every revert that is not exact\n\n");
    let mut tsv = String::from(
        "file\tformat\tnoop\tdetail\trevert\trevert_detail\tstructure\tstructure_detail\n",
    );
    let mut unexpected = Vec::new();
    for (relative, outcome) in &outcomes {
        let _ = writeln!(
            tsv,
            "{relative}\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
            outcome.format,
            outcome.noop,
            clean(&outcome.detail),
            outcome.revert,
            clean(&outcome.revert_detail),
            outcome.structure,
            clean(&outcome.structure_detail)
        );
        let save_only_differs = outcome.detail.contains("save-only DIFFERS");
        let bad = matches!(
            outcome.noop,
            "differs" | "error" | "signed-differs" | "signed-error"
        ) || save_only_differs
            || outcome.revert == "mismatch"
            || outcome.structure == "mismatch";
        if bad
            || !matches!(outcome.noop, "exact" | "signed-exact" | "locked")
            || outcome.revert == "one-way"
            || outcome.structure == "one-way"
        {
            let _ = writeln!(
                markdown,
                "- `{relative}` ({}): {} {}{}",
                outcome.format,
                outcome.noop,
                clean(&outcome.detail),
                if matches!(outcome.revert, "mismatch" | "one-way") {
                    format!(
                        "; revert {}: {}",
                        outcome.revert,
                        clean(&outcome.revert_detail)
                    )
                } else if matches!(outcome.structure, "mismatch" | "one-way") {
                    format!(
                        "; structure {}: {}",
                        outcome.structure,
                        clean(&outcome.structure_detail)
                    )
                } else {
                    String::new()
                }
            );
        }
        if bad
            && !NORMALIZING
                .iter()
                .any(|(format, _)| *format == outcome.format)
        {
            unexpected.push(format!(
                "{relative} ({}): {} {} {} {}",
                outcome.format,
                outcome.noop,
                outcome.detail,
                outcome.revert_detail,
                outcome.structure_detail
            ));
        }
    }
    eprintln!(
        "{} files opened as claimed formats in {:.1} s",
        outcomes.len(),
        elapsed.as_secs_f64()
    );
    let out = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../docs/unified-ops/results/0289");
    if std::env::var_os("LITCHI_BLESS").is_some() {
        std::fs::create_dir_all(&out).unwrap();
        std::fs::write(out.join("fidelity.md"), &markdown).unwrap();
        std::fs::write(out.join("fidelity.tsv"), &tsv).unwrap();
    }
    assert!(
        unexpected.is_empty(),
        "fidelity failures ({}):\n{}",
        unexpected.len(),
        unexpected.join("\n")
    );
    if std::env::var_os("LITCHI_BLESS").is_none() {
        let golden = std::fs::read_to_string(out.join("fidelity.tsv"))
            .expect("fidelity report; bless with LITCHI_BLESS=1");
        assert!(
            golden == tsv,
            "the fidelity report changed; review and bless with LITCHI_BLESS=1"
        );
        let golden = std::fs::read_to_string(out.join("fidelity.md")).unwrap();
        assert!(
            golden == markdown,
            "fidelity.md is stale; bless with LITCHI_BLESS=1"
        );
    }
}
