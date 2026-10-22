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
//! (record 0293), Pages (record 0299), Keynote (record 0298), Numbers
//! (record 0301) and XLSB (record 0322): a row inserted and removed, a slide
//! or paragraph moved and moved back, a Pages paragraph added and removed,
//! a Keynote slide duplicated and a slide added and both removed, a Numbers
//! table added and removed, or an XLSB note added and removed, must read
//! back with the same node texts. DOC (record 0324) adds a bookmark and a
//! hyperlink to its first body paragraph and removes both; DOCX (record
//! 0335) adds a picture and a comment to its first body paragraph, sets
//! each once, and removes both; ODT, OTH and ODM
//! (record 0328) add a tab stop to the first body paragraph and remove it,
//! and ODP (record 0328) a line break to its first shape's first paragraph.
//!
//! An iWork package directory (the app-authored bundle form, record 0306)
//! is a corpus entry too: a copy of it goes through `open` → empty `edit`
//! → `save()` in place, and every file of the saved bundle must equal the
//! source file of the same name.
//!
//! An empty edit the leaf editor refuses (typed) is reported with its
//! cause, and the file then goes through `open` → `save_as`, which must be
//! exact too. The whole corpus runs in about 20 s in a debug build, so no
//! subset is needed.
//!
//! Every Office-extension file that does not open is listed in
//! [`NOT_OPENED`] with its expected typed code and class (invalid, fuzzed,
//! legacy version, opens leniently, or reader gap; record 0341): a new
//! refusal, a changed code or a listed file that opens fails the test, and
//! the lenient class is opened with its leniency here.
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
                if is_bundle(&path) {
                    let relative = path
                        .strip_prefix(&root)
                        .unwrap()
                        .to_string_lossy()
                        .into_owned();
                    found.push((relative, 0));
                }
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

/// Whether `path` is an iWork package directory: a `.pages`, `.key` or
/// `.numbers` directory holding `Index.zip`.
fn is_bundle(path: &Path) -> bool {
    path.extension()
        .is_some_and(|extension| matches!(extension.to_str(), Some("pages" | "key" | "numbers")))
        && path.join("Index.zip").is_file()
}

/// Every file below `root`, by `/`-separated relative name.
fn bundle_files(root: &Path) -> BTreeMap<String, Vec<u8>> {
    fn walk(directory: &Path, prefix: &str, files: &mut BTreeMap<String, Vec<u8>>) {
        for entry in std::fs::read_dir(directory).unwrap().flatten() {
            let name = entry.file_name().to_string_lossy().into_owned();
            let name = if prefix.is_empty() {
                name
            } else {
                format!("{prefix}/{name}")
            };
            if entry.path().is_dir() {
                walk(&entry.path(), &name, files);
            } else {
                files.insert(name, std::fs::read(entry.path()).unwrap());
            }
        }
    }
    let mut files = BTreeMap::new();
    walk(root, "", &mut files);
    files
}

fn copy_tree(from: &Path, to: &Path) {
    std::fs::create_dir_all(to).unwrap();
    for entry in std::fs::read_dir(from).unwrap().flatten() {
        let target = to.join(entry.file_name());
        if entry.path().is_dir() {
            copy_tree(&entry.path(), &target);
        } else {
            std::fs::copy(entry.path(), target).unwrap();
        }
    }
}

/// The no-op round trip of a package directory (record 0306): a copy is
/// opened, edited with no mutation and saved in place; every saved file must
/// equal its source.
fn run_bundle(relative: &str, dir: &Path) -> Result<Outcome, String> {
    let source = test_data().join(relative);
    let copy = dir.join(format!(
        "bundle.{}",
        relative.rsplit('.').next().unwrap_or("bin")
    ));
    let _ = std::fs::remove_dir_all(&copy);
    copy_tree(&source, &copy);
    let file = guarded(|| open(&copy))?;
    let format = file.format().map_or("other", name);
    let (noop, detail) = match guarded(|| file.edit(|_| Ok(()))) {
        Err(error) if untyped(&error) => ("error", format!("edit: {error}")),
        Err(error) => ("refused", format!("edit: {error}")),
        Ok(commit) => match guarded(|| commit.file().save()) {
            Ok(_) => {
                let (before, after) = (bundle_files(&source), bundle_files(&copy));
                if before == after {
                    ("exact", String::new())
                } else {
                    let changed: Vec<&String> = before
                        .keys()
                        .chain(after.keys())
                        .filter(|name| before.get(*name) != after.get(*name))
                        .collect();
                    ("differs", format!("bundle files differ: {changed:?}"))
                }
            },
            Err(error) if untyped(&error) => ("error", format!("save: {error}")),
            Err(error) => ("refused", format!("save: {error}")),
        },
    };
    let _ = std::fs::remove_dir_all(&copy);
    Ok(Outcome {
        format,
        noop,
        detail,
        revert: "-",
        revert_detail: String::new(),
        structure: "-",
        structure_detail: String::new(),
    })
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

/// A 1x1 PNG, the DOCX structural probe's image (record 0335).
const PROBE_PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

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
    type Forward<'a> = Box<dyn Fn(&File) -> litchi_ops::Result<File> + 'a>;
    type Inverse<'a> = Box<dyn Fn(&File) -> litchi_ops::Result<File> + 'a>;
    let count = |parent: &str, kind: &str| {
        file.get(parent).map_or(0, |node| {
            node.children()
                .filter_map(Result::ok)
                .filter(|child| child.kind() == kind)
                .count()
        })
    };
    let (label, checked, forward, inverse): (&str, &[&str], Forward<'_>, Inverse<'_>) = match file
        .format()
    {
        Some(FileFormat::Xls) => {
            if count("/", "sheet") == 0 {
                return ("none", String::new());
            }
            (
                "row add+remove",
                &["/sheet[1]/row[1]"],
                Box::new(|file: &File| {
                    file.add("/sheet[1]", NewNode::new("row").at(Placement::Start))
                        .map(|added| added.file)
                }),
                Box::new(|file: &File| file.remove("/sheet[1]/row[1]")),
            )
        },
        Some(FileFormat::Pptx) => {
            if count("/", "slide") < 2 {
                return ("none", String::new());
            }
            (
                "slide move+back",
                &["/slide[1]", "/slide[2]"],
                Box::new(|file: &File| file.move_to("/slide[1]", Placement::after("/slide[2]")?)),
                Box::new(|file: &File| file.move_to("/slide[2]", Placement::Start)),
            )
        },
        Some(FileFormat::Rtf) => {
            if count("/body", "paragraph") < 2 {
                return ("none", String::new());
            }
            (
                "paragraph move+back",
                &["/body/p[1]", "/body/p[2]"],
                Box::new(|file: &File| file.move_to("/body/p[1]", Placement::after("/body/p[2]")?)),
                Box::new(|file: &File| {
                    file.move_to("/body/p[2]", Placement::before("/body/p[1]")?)
                }),
            )
        },
        // Record 0324: a bookmark and a hyperlink added to the first body
        // paragraph and removed again, through the DOC bookmark-table and
        // field editors.
        Some(FileFormat::Doc) => {
            if count("/body", "paragraph") == 0 {
                return ("none", String::new());
            }
            (
                "bookmark+hyperlink add+remove",
                &["/body/p[1]"],
                Box::new(|file: &File| {
                    let marked = file
                        .add(
                            "/body/p[1]",
                            NewNode::new("bookmark").prop("name", "LitchiProbe"),
                        )?
                        .file;
                    marked
                        .add(
                            "/body/p[1]",
                            NewNode::new("hyperlink")
                                .prop("url", "http://example.com/")
                                .prop("text", "probe"),
                        )
                        .map(|added| added.file)
                }),
                Box::new(|file: &File| {
                    file.remove("/body/p[1]/hyperlink[last()]")?
                        .remove("/bookmark[@name=LitchiProbe]")
                }),
            )
        },
        // Record 0335: a picture and a comment added to the first body
        // paragraph, each set once, and both removed again, through the
        // leaf picture and comment edits; record 0345 resolves the comment.
        Some(FileFormat::Docx) => {
            if count("/body", "paragraph") == 0 {
                return ("none", String::new());
            }
            (
                "picture+comment add+set+remove",
                &["/body/p[1]"],
                Box::new(|file: &File| {
                    let pictured =
                        file.add("/body/p[1]", NewNode::new("picture").prop("src", PROBE_PNG))?;
                    let sized = pictured.file.set(
                        pictured.path.to_string().as_str(),
                        Props::new().with("width", "1in"),
                    )?;
                    let commented = sized.add(
                        "/body/p[1]",
                        NewNode::new("comment")
                            .prop("author", "Litchi")
                            .prop("text", "Structural probe"),
                    )?;
                    // Record 0345: resolving the comment writes its
                    // `commentsExtended` entry, which the removal strips.
                    commented.file.set(
                        commented.path.to_string().as_str(),
                        Props::new()
                            .with("text", "Structural probe, set")
                            .with("done", true),
                    )
                }),
                Box::new(|file: &File| {
                    let comments = file.query("/comments/comment")?.count();
                    let pictures = file.query("/body/p[1]/picture")?.count();
                    file.remove(format!("/comments/comment[{comments}]").as_str())?
                        .remove(format!("/body/p[1]/picture[{pictures}]").as_str())
                }),
            )
        },
        // Record 0328: a tab stop added to the first body paragraph and
        // removed again, through the leaf paragraph-property edit.
        Some(FileFormat::Odt | FileFormat::Oth | FileFormat::Odm) => {
            if count("/body", "paragraph") == 0 {
                return ("none", String::new());
            }
            (
                "tab add+remove",
                &["/body/p[1]"],
                Box::new(|file: &File| {
                    file.add("/body/p[1]", NewNode::new("tab").prop("pos", 7777))
                        .map(|added| added.file)
                }),
                Box::new(|file: &File| {
                    let tabs = file.query("/body/p[1]/tab")?.count();
                    let probe = (1..=tabs)
                        .find(|n| {
                            current(file, &format!("/body/p[1]/tab[{n}]"), "pos").as_deref()
                                == Some("7777")
                        })
                        .unwrap_or(tabs);
                    file.remove(format!("/body/p[1]/tab[{probe}]").as_str())
                }),
            )
        },
        // Record 0328: a line break appended to the first shape's first
        // paragraph and removed again, through the leaf scene splices.
        Some(FileFormat::Odp) => {
            if file.get("/slide[1]/shape[1]/paragraph[1]").is_err() {
                return ("none", String::new());
            }
            (
                "linebreak add+remove",
                &["/slide[1]/shape[1]/paragraph[1]", "/slide[1]/shape[1]"],
                Box::new(|file: &File| {
                    file.add("/slide[1]/shape[1]/paragraph[1]", NewNode::new("linebreak"))
                        .map(|added| added.file)
                }),
                Box::new(|file: &File| {
                    let breaks = file
                        .query("/slide[1]/shape[1]/paragraph[1]/linebreak")?
                        .count();
                    file.remove(
                        format!("/slide[1]/shape[1]/paragraph[1]/linebreak[{breaks}]").as_str(),
                    )
                }),
            )
        },
        // Record 0299: a Pages paragraph added at the start of the body and
        // removed again, through the leaf paragraph transaction.
        Some(FileFormat::Pages) => (
            "paragraph add+remove",
            &["/body/p[1]"],
            Box::new(|file: &File| {
                file.add(
                    "/body",
                    NewNode::new("p")
                        .prop("text", "Structural probe")
                        .at(Placement::Start),
                )
                .map(|added| added.file)
            }),
            Box::new(|file: &File| file.remove("/body/p[1]")),
        ),
        // Record 0298: the first Keynote slide duplicated after itself and a
        // slide added from the theme's default layout at the end, then both
        // removed, through the leaf slide-creation and deletion families.
        Some(FileFormat::Keynote) => (
            "slide duplicate+add+remove",
            &["/slide[1]/placeholder[1]", "/slide[1]/shape[1]"],
            Box::new(|file: &File| {
                let copied = file
                    .add(
                        "/",
                        NewNode::new("slide")
                            .prop("from", "/slide[1]")
                            .at(Placement::after("/slide[1]")?),
                    )?
                    .file;
                copied
                    .add("/", NewNode::new("slide"))
                    .map(|added| added.file)
            }),
            Box::new(|file: &File| {
                let last = file.query("slide")?.count();
                file.remove(format!("/slide[{last}]").as_str())?
                    .remove("/slide[2]")
            }),
        ),
        // Record 0301: a Numbers table added first in the first sheet and
        // removed again, through the leaf table-lifecycle and table-order
        // transactions. Record 0309: a row inserted at the start of the
        // original table and removed again, through the leaf topology
        // transaction (skipped where that table refuses it, typed).
        Some(FileFormat::Numbers) => {
            let rows = |file: &File, path: &str| current(file, path, "rowCount");
            let before = rows(file, "/sheet[1]/table[1]");
            (
                "table add+remove, row add+remove",
                &["/sheet[1]/table[1]/A1", "/sheet[1]/table[1]/B2"],
                Box::new(|file: &File| {
                    let added = file
                        .add("/sheet[1]", NewNode::new("table").at(Placement::Start))?
                        .file;
                    match added.add(
                        "/sheet[1]/table[2]",
                        NewNode::new("row").at(Placement::Start),
                    ) {
                        Ok(row) => Ok(row.file),
                        Err(error) if !untyped(&error.to_string()) => Ok(added),
                        Err(error) => Err(error),
                    }
                }),
                Box::new(move |file: &File| {
                    let file = if rows(file, "/sheet[1]/table[2]") != before {
                        file.remove("/sheet[1]/table[2]/row[1]")?
                    } else {
                        file.clone()
                    };
                    file.remove("/sheet[1]/table[1]")
                }),
            )
        },
        // Record 0322: an XLSB note added on the first worksheet's last cell
        // and removed again, through the leaf note transaction (the comments
        // part, its VML note shape and `BrtLegacyDrawing`).
        Some(FileFormat::Xlsb) => {
            if count("/", "sheet") == 0 {
                return ("none", String::new());
            }
            (
                "note add+remove",
                &["/sheet[1]/A1", "/sheet[1]/comment[1]"],
                Box::new(|file: &File| {
                    file.add(
                        "/sheet[1]",
                        NewNode::new("comment")
                            .prop("ref", "XFD1048576")
                            .prop("text", "Structural probe"),
                    )
                    .map(|added| added.file)
                }),
                Box::new(|file: &File| {
                    // The probe is the last note in row-major order.
                    let mut last = 1;
                    while file
                        .get(format!("/sheet[1]/comment[{}]", last + 1).as_str())
                        .is_ok()
                    {
                        last += 1;
                    }
                    file.remove(format!("/sheet[1]/comment[{last}]").as_str())
                }),
            )
        },
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
    if path.is_dir() {
        return run_bundle(relative, dir);
    }
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
        Ok(commit) => match save_bytes(commit.file(), dir, &extension) {
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

/// Why an Office-extension file of the corpus does not open (record 0341).
#[derive(Clone, Copy, PartialEq, Eq)]
enum Refusal {
    /// The input breaks its specification and the strict reader refuses it.
    Invalid,
    /// A fuzzer-made or garbage file.
    Fuzzed,
    /// An older format version Litchi does not read (Word 6.0 and 95).
    Legacy,
    /// Invalid input the reader opens under an opt-in leniency, checked here.
    Lenient,
    /// Valid input the reader does not support yet (a follow-up).
    Gap,
}

impl Refusal {
    const fn name(self) -> &'static str {
        match self {
            Self::Invalid => "invalid input",
            Self::Fuzzed => "fuzzed input",
            Self::Legacy => "legacy version",
            Self::Lenient => "invalid; opens leniently",
            Self::Gap => "reader gap",
        }
    }
}

/// Every Office-extension file of the corpus that does not open: the file,
/// its expected typed code, the class and why (record 0341). A file that
/// fails differently, a new file that does not open, and a listed file
/// that now opens all fail the test.
const NOT_OPENED: &[(&str, &str, Refusal, &str)] = &[
    (
        "libreoffice-core/chart2/qa/extras/data/fods/stacked-column-chart.fods",
        "corrupt_file",
        Refusal::Gap,
        "a chart object inlined as a nested office:document in draw:object (valid flat ODF); the flat ODS validator takes its office:body for a second body",
    ),
    (
        "libreoffice-core/sc/qa/unit/data/functions/mathematical/fods/aggregate.fods",
        "corrupt_file",
        Refusal::Invalid,
        "Sheet2 declares 1,048,578 logical rows, two past the 1,048,576-row grid",
    ),
    (
        "libreoffice-core/sw/qa/core/data/rtf/fail/forcepoint-4.rtf",
        "corrupt_file",
        Refusal::Fuzzed,
        "LibreOffice's must-fail fuzz case: a bitmap header without a DIB or bitmap picture",
    ),
    (
        "libreoffice-core/sw/qa/core/data/rtf/fail/forcepoint-5.rtf",
        "corrupt_file",
        Refusal::Fuzzed,
        "LibreOffice's must-fail fuzz case: a control parameter below i32",
    ),
    (
        "libreoffice-core/sw/qa/core/data/rtf/pass/forcepoint-1.rtf",
        "corrupt_file",
        Refusal::Fuzzed,
        "a fuzz case LibreOffice must survive: a control parameter past i32",
    ),
    (
        "libreoffice-core/sw/qa/core/data/rtf/pass/forcepoint-3.rtf",
        "corrupt_file",
        Refusal::Fuzzed,
        "a fuzz case LibreOffice must survive: a control parameter below i32",
    ),
    (
        "libreoffice-core/sw/qa/core/data/ww8/fail/redline-1.doc",
        "unsupported_format",
        Refusal::Fuzzed,
        "LibreOffice's must-fail case: a compound file with a corrupt DIFAT (first index out of range)",
    ),
    (
        "libreoffice-core/sw/qa/extras/rtfexport/data/all_gaps_word.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "a \\brdrw border width past the RTF 1.9.1 maximum of 75",
    ),
    (
        "libreoffice-core/sw/qa/extras/rtfexport/data/fdo77996.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "a bare \\rtf root without the required version parameter",
    ),
    (
        "libreoffice-core/sw/qa/extras/rtfexport/data/tdf108505_fieldCharFormat2.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "a \\datafield with an odd number of hexadecimal digits",
    ),
    (
        "libreoffice-core/sw/qa/extras/rtfexport/data/tdf167569-2.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "a revision root outside the revision-save table",
    ),
    (
        "libreoffice-core/sw/qa/extras/rtfimport/data/fdo53556.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "shape properties after \\shptxt instead of as direct \\shpinst children",
    ),
    (
        "libreoffice-core/sw/qa/extras/rtfimport/data/tdf117403.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "a bare \\rtf root without the required version parameter",
    ),
    (
        "libreoffice-core/sw/qa/extras/rtfimport/data/tdf91684.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "an unstarred \\shprslt destination in a shape group",
    ),
    (
        "libreoffice-core/sw/qa/writerfilter/rtftok/data/negative-page-border.rtf",
        "corrupt_file",
        Refusal::Invalid,
        "a negative page-border width (RTF admits 0..=75 twips)",
    ),
    (
        "odf/corpus/writer-minimal-nasty.odt",
        "corrupt_file",
        Refusal::Invalid,
        "content.xml references the undeclared entity &nasty;",
    ),
    (
        "odf/drawing/dashed-line.fodg",
        "corrupt_file",
        Refusal::Invalid,
        "a style fixture with an empty office:body (ODF requires office:drawing)",
    ),
    (
        "odf/drawing/hatch-angles.fodg",
        "corrupt_file",
        Refusal::Invalid,
        "a style fixture with an empty office:body (ODF requires office:drawing)",
    ),
    (
        "odf/drawing/multicolor-gradient.fodp",
        "corrupt_file",
        Refusal::Invalid,
        "a style fixture with an empty office:body (ODF requires office:presentation)",
    ),
    (
        "ole/doc/cfb-v3-uninitialized-size-high-word.doc",
        "unsupported_format",
        Refusal::Legacy,
        "a Word 95 document (nFib 0x0068)",
    ),
    (
        "ole/doc/duplicate-style-names.doc",
        "corrupt_file",
        Refusal::Lenient,
        "duplicate style names ([MS-DOC] 2.9 requires unique names); opens with Leniency::TolerateStylesheetDefects",
    ),
    (
        "ole/doc/footnote.doc",
        "corrupt_file",
        Refusal::Lenient,
        "duplicate style names; opens with Leniency::TolerateStylesheetDefects",
    ),
    (
        "ole/doc/lists-margins.doc",
        "corrupt_file",
        Refusal::Lenient,
        "duplicate style names; opens with Leniency::TolerateStylesheetDefects",
    ),
    (
        "ole/doc/picture.doc",
        "corrupt_file",
        Refusal::Lenient,
        "duplicate style names; opens with Leniency::TolerateStylesheetDefects",
    ),
    (
        "ole/doc/pictures_escher.doc",
        "corrupt_file",
        Refusal::Lenient,
        "duplicate style names; opens with Leniency::TolerateStylesheetDefects",
    ),
    (
        "ole/doc/word6-no-table-stream.doc",
        "unsupported_format",
        Refusal::Legacy,
        "a Word 6.0 document (nFib 0x0065)",
    ),
    (
        "poi/test-data/hpsf/TestBug52117.doc",
        "unsupported_format",
        Refusal::Legacy,
        "a Word 6.0 document (nFib 0x0065)",
    ),
    (
        "poi/test-data/hpsf/TestMickey.doc",
        "unsupported_format",
        Refusal::Legacy,
        "a Word 95 document (nFib 0x0068)",
    ),
    (
        "poi/test-data/openxml4j/OPCCompliance_CoreProperties_OnlyOneCorePropertiesPartFAIL.docx",
        "corrupt_file",
        Refusal::Invalid,
        "Apache POI's OPC-conformance negative case: two core-properties relationships",
    ),
    (
        "poi/test-data/openxml4j/OPCCompliance_DerivedPartNameFAIL.docx",
        "corrupt_file",
        Refusal::Invalid,
        "Apache POI's OPC-conformance negative case: a part name derived from another part's name",
    ),
    (
        "poi/test-data/spreadsheet/1900DateWindowing.xls",
        "unsupported_format",
        Refusal::Invalid,
        "a compound file whose FAT sectors are not marked FATSECT ([MS-CFB] requires it)",
    ),
    (
        "poi/test-data/spreadsheet/1904DateWindowing.xls",
        "unsupported_format",
        Refusal::Invalid,
        "a compound file whose FAT sectors are not marked FATSECT ([MS-CFB] requires it)",
    ),
    (
        "poi/test-data/spreadsheet/57456.xls",
        "corrupt_file",
        Refusal::Invalid,
        "an SST declaring 1,761 unique strings and 0 total",
    ),
    (
        "poi/test-data/spreadsheet/SharedFormulaTest.xls",
        "corrupt_file",
        Refusal::Invalid,
        "an XFCRC record declaring 72 XF records where 62 are present",
    ),
    (
        "rtf/binary-garbage-nested-groups.rtf",
        "unsupported_format",
        Refusal::Fuzzed,
        "binary garbage after an opening brace",
    ),
    (
        "rtf/test10001Encoding.rtf",
        "unsupported_capability",
        Refusal::Gap,
        "\\ansicpg10001 (Mac Japanese) has no exact codec; refused as unsupported (record 0347)",
    ),
    (
        "rtf/test10007Encoding.rtf",
        "unsupported_capability",
        Refusal::Gap,
        "\\ansicpg10007 (Mac Cyrillic) has no exact codec; refused as unsupported (record 0347)",
    ),
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
    let mut not_opened_files: Vec<(String, String)> = Vec::new();
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
                    not_opened_files.push((relative.clone(), error.clone()));
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
    // Record 0341: each file that does not open, classified.
    markdown.push_str(
        "\n### Each file, classified (record 0341)\n\n\
         | file | code | class | why |\n|---|---|---|---|\n",
    );
    let mut triage = Vec::new();
    for (file, error) in &not_opened_files {
        let code = error.split(':').next().unwrap_or("");
        match NOT_OPENED.iter().find(|(listed, ..)| listed == file) {
            Some((_, expected, class, why)) => {
                if code != *expected {
                    triage.push(format!(
                        "{file}: refused {code}, expected {expected}: {error}"
                    ));
                }
                if *class == Refusal::Lenient {
                    let options = litchi_ops::OpenOptions::new().doc(
                        litchi_ops::DocOptions::new()
                            .leniency(litchi_doc::Leniency::TolerateStylesheetDefects),
                    );
                    if let Err(error) =
                        guarded(|| litchi_ops::open_with(test_data().join(file), &options))
                    {
                        triage.push(format!("{file}: does not open leniently: {error}"));
                    }
                }
                let _ = writeln!(
                    markdown,
                    "| `{file}` | `{code}` | {} | {} |",
                    class.name(),
                    why.replace('|', "\\|")
                );
            },
            None => triage.push(format!("{file}: not classified: {error}")),
        }
    }
    for (listed, ..) in NOT_OPENED {
        if !not_opened_files.iter().any(|(file, _)| file == listed) {
            triage.push(format!("{listed}: listed as not opening, but it opens"));
        }
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
        triage.is_empty(),
        "files that do not open, against NOT_OPENED ({}):\n{}",
        triage.len(),
        triage.join("\n")
    );
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
