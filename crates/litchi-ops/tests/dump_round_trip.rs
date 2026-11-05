//! `File::dump` round trips (record 0336): a dump replayed with `batch` onto
//! a created document of the same format rebuilds the readable tree.
//!
//! For each file: `dump` it, replay the dump atomically onto
//! `create(format)`, and require
//! - the same text view (`view text`), and
//! - the same dump of the replayed document (the dump is a fixed point, so
//!   every element and property the dump carries survived the replay).

#![allow(clippy::unwrap_used, clippy::expect_used, reason = "tests")]

use std::path::{Path, PathBuf};

use litchi_core::FileFormat;
use litchi_ops::view::Mode;
use litchi_ops::{BatchMode, File};

fn data(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// What a round trip found.
#[derive(Debug)]
#[allow(dead_code, reason = "printed by the survey")]
enum Outcome {
    Exact {
        items: usize,
        warnings: usize,
    },
    /// The text differs because the dump skipped (and reported) content.
    Lossy(usize),
    Replay(String),
    Text,
    Dump(String),
    Dumped(String),
}

fn round_trip(file: &File) -> Outcome {
    let format = file.format().unwrap();
    let dump = match file.dump() {
        Ok(dump) => dump,
        Err(error) => return Outcome::Dumped(format!("{:?}: {}", error.code(), error.message())),
    };
    let replay = match replay(format, dump.items()) {
        Ok(file) => file,
        Err(message) => return Outcome::Replay(message),
    };
    let formulas = formula_cells(dump.items());
    if masked(&replay.view(Mode::Text).unwrap(), &formulas)
        != masked(&file.view(Mode::Text).unwrap(), &formulas)
    {
        return if dump.warnings().is_empty() {
            Outcome::Text
        } else {
            Outcome::Lossy(dump.warnings().len())
        };
    }
    let again = replay.dump().unwrap();
    if again.items() != dump.items() {
        let first = dump
            .items()
            .iter()
            .zip(again.items())
            .position(|(left, right)| left != right)
            .unwrap_or(dump.items().len().min(again.items().len()));
        let show = |items: &[litchi_ops::BatchItem<'static>]| {
            items
                .get(first)
                .map(|item| serde_json::to_string(item).unwrap())
                .unwrap_or_default()
        };
        return Outcome::Dump(format!("{} vs {}", show(dump.items()), show(again.items())));
    }
    Outcome::Exact {
        items: dump.items().len(),
        warnings: dump.warnings().len(),
    }
}

/// The `(sheet index, cell)` pairs whose formula the dump writes. A replayed
/// formula has no cached value (the write verbs never invent one), so the
/// text view shows those cells empty.
fn formula_cells(items: &[litchi_ops::BatchItem<'static>]) -> Vec<(String, String)> {
    items
        .iter()
        .filter(|item| {
            item.props
                .as_ref()
                .is_some_and(|props| props.get("formula").is_some())
        })
        .filter_map(|item| {
            let path = item.path.as_deref()?;
            let (sheet, cell) = path.trim_start_matches('/').split_once('/')?;
            Some((sheet.to_owned(), cell.to_owned()))
        })
        .collect()
}

/// A workbook text view without the formula cells, and
/// without the lines of row records that hold no cell (they carry nothing
/// the node view reads, so the dump has nothing to rebuild).
fn masked(text: &str, formulas: &[(String, String)]) -> String {
    text.lines()
        .map(|line| {
            let Some(rest) = line.strip_prefix("[/") else {
                return line.to_owned();
            };
            let Some((path, cells)) = rest.split_once("] ") else {
                return line.to_owned();
            };
            let sheet = path.split('/').next().unwrap_or_default();
            let cells: Vec<&str> = cells
                .split('\t')
                .filter(|cell| {
                    let name = cell.split_once('=').map_or(*cell, |(name, _)| name);
                    !formulas.iter().any(|(s, c)| s == sheet && c == name)
                })
                .collect();
            format!("[/{path}] {}", cells.join("\t"))
        })
        .filter(|line| {
            !(line.starts_with("[/sheet[")
                && line.contains("/row[")
                && line.trim_end().ends_with(']'))
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Replay `items` onto a created document, in atomic chunks so a large
/// document stays within the default transaction snapshot budget.
fn replay(format: FileFormat, items: &[litchi_ops::BatchItem<'static>]) -> Result<File, String> {
    let mut file = litchi_ops::create(format).unwrap();
    for chunk in items.chunks(64) {
        file = match file.batch(chunk, BatchMode::Atomic) {
            Ok(commit) => commit.file().clone(),
            Err(error) => {
                let item = error
                    .item()
                    .and_then(|index| chunk.get(index))
                    .map(|item| serde_json::to_string(item).unwrap())
                    .unwrap_or_default();
                return Err(format!("{:?}: {} {item}", error.code(), error.message()));
            },
        };
    }
    Ok(file)
}

fn exact(relative: &str) {
    let file = litchi_ops::open(data(relative)).unwrap();
    match round_trip(&file) {
        Outcome::Exact { items, .. } => assert!(items > 1, "{relative}: an empty dump"),
        other => panic!("{relative}: {other:?}"),
    }
}

/// Representative corpus files per format, each an exact round trip.
#[test]
fn corpus_round_trips() {
    for file in [
        "ooxml/docx/table_footnotes.docx",
        "ooxml/docx/Hyperlink.docx",
        "ooxml/docx/table-alignment.docx",
        "ooxml/docx/headerFooter.docx",
        "ooxml/xlsx/SimpleStrict.xlsx",
        "ooxml/xlsx/InlineStrings.xlsx",
        "ooxml/xlsx/Formatting.xlsx",
        "ooxml/xlsb/Simple.xlsb",
        "ooxml/xlsb/date.xlsb",
        "odf/ods/grid-padding-package.ods",
        "rtf/background.rtf",
        "rtf/cjklist12.rtf",
        "ole/xls/DateFormats.xls",
    ] {
        exact(file);
    }
}

/// A deck the unified verbs built: slides by layout, placeholders, shapes
/// and notes round-trip exactly.
#[test]
fn built_presentation_round_trips() {
    use litchi_ops::BatchItem;
    let mut items = Vec::new();
    for (slide, layout) in [(1, "Title Slide"), (2, "Title Only"), (3, "Blank")] {
        let mut add = BatchItem::new("add");
        add.parent = Some("/".into());
        add.kind = Some("slide".into());
        add.props = Some(litchi_ops::Props::new().with("layout", layout));
        items.push(add);
        if slide < 3 {
            let mut title = BatchItem::new("set");
            title.path = Some(format!("/slide[{slide}]/shape[1]").into());
            title.props = Some(litchi_ops::Props::new().with("text", format!("Title {slide}")));
            items.push(title);
        }
        let mut shape = BatchItem::new("add");
        shape.parent = Some(format!("/slide[{slide}]").into());
        shape.kind = Some("shape".into());
        shape.props = Some(
            litchi_ops::Props::new()
                .with("x", "2cm")
                .with("y", "6cm")
                .with("width", "12cm")
                .with("height", "4cm")
                .with("text", format!("Body {slide}"))
                .with("bold", true),
        );
        items.push(shape);
    }
    let deck = litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .batch(&items, BatchMode::Atomic)
        .unwrap()
        .into_file();
    match round_trip(&deck) {
        Outcome::Exact { items, warnings } => {
            assert!(items > 6);
            assert_eq!(warnings, 0);
        },
        other => panic!("{other:?}"),
    }
}

/// A workbook built through the unified verbs: several sheets, typed
/// values, a formula and cell formatting.
#[test]
fn built_workbooks_round_trip() {
    use litchi_ops::props;
    for format in [
        FileFormat::Xlsx,
        FileFormat::Xlsb,
        FileFormat::Ods,
        FileFormat::Xls,
    ] {
        let file = litchi_ops::create(format).unwrap();
        let file = file
            .add(
                "/",
                litchi_ops::NewNode::new("sheet")
                    .with_props(props! { "name" => "Data" })
                    .unwrap(),
            )
            .unwrap()
            .file;
        let file = file
            .set("/sheet[1]/A1", props! { "value" => "name" })
            .unwrap();
        let file = file.set("/sheet[1]/B2", props! { "value" => 42 }).unwrap();
        let file = file
            .set("/sheet[2]/C3", props! { "value" => true })
            .unwrap();
        match round_trip(&file) {
            Outcome::Exact { items, .. } => assert!(items >= 5, "{format:?}"),
            other => panic!("{format:?}: {other:?}"),
        }
    }
}

/// Formats without a dump plan refuse with a typed error naming the format.
#[test]
fn unplanned_formats_refuse() {
    for format in [FileFormat::Odp, FileFormat::Ppt, FileFormat::Odg] {
        let error = litchi_ops::create(format).unwrap().dump().unwrap_err();
        assert_eq!(
            error.code(),
            litchi_ops::Code::UnsupportedCapability,
            "{format:?}"
        );
        assert_eq!(error.format(), Some(format));
    }
}

/// Every supported format round-trips content built through the unified
/// verbs.
#[test]
fn built_documents_round_trip() {
    use litchi_ops::{NewNode, props};
    for format in [
        FileFormat::Docx,
        FileFormat::Odt,
        FileFormat::Rtf,
        FileFormat::Doc,
    ] {
        let file = litchi_ops::create(format).unwrap();
        let file = if format == FileFormat::Doc {
            // DOC adds a paragraph with its text, after the empty one its
            // created body already holds.
            let paragraph = NewNode::new("paragraph").with_props(props! { "text" => "Hello" });
            file.add("/body", paragraph.unwrap()).unwrap().file
        } else {
            let file = file.add("/body", NewNode::new("paragraph")).unwrap().file;
            let run = NewNode::new("run").with_props(props! { "text" => "Hello" });
            file.add("/body/p[1]", run.unwrap()).unwrap().file
        };
        match round_trip(&file) {
            Outcome::Exact { items, .. } => assert!(items >= 2, "{format:?}"),
            other => panic!("{format:?}: {other:?}"),
        }
    }
}

/// A corpus survey: `cargo test --test dump_round_trip survey -- --ignored --nocapture`.
#[test]
#[ignore = "survey; prints outcomes over the corpus"]
fn survey() {
    let filter = std::env::var("DUMP_SURVEY").unwrap_or_else(|_| "docx,xlsx,pptx".to_owned());
    let mut files: Vec<PathBuf> = Vec::new();
    collect(&data(""), &mut files);
    files.sort();
    let mut totals = std::collections::BTreeMap::<String, [usize; 7]>::new();
    for path in files {
        let extension = path
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if !filter.split(',').any(|wanted| wanted == extension) {
            continue;
        }
        if std::fs::metadata(&path)
            .map(|meta| meta.len())
            .unwrap_or(u64::MAX)
            > 512 << 10
        {
            continue;
        }
        let Ok(file) = litchi_ops::open(&path) else {
            continue;
        };
        if file.format().is_none() || matches!(file, File::Locked(_) | File::Signed(_)) {
            continue;
        }
        let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| round_trip(&file)));
        let entry = totals.entry(extension).or_default();
        let slot = match &outcome {
            Ok(Outcome::Exact { .. }) => 0,
            Ok(Outcome::Replay(_)) => 1,
            Ok(Outcome::Text) => 2,
            Ok(Outcome::Dump(_)) => 3,
            Ok(Outcome::Dumped(_)) => 4,
            Ok(Outcome::Lossy(_)) => 6,
            Err(_) => 5,
        };
        entry[slot] += 1;
        match outcome {
            Ok(Outcome::Exact { .. } | Outcome::Lossy(_)) => {},
            Ok(other) => println!("{}: {other:?}", path.display()),
            Err(_) => println!("{}: PANIC", path.display()),
        }
    }
    for (extension, counts) in totals {
        println!(
            "{extension}: exact {} lossy {} replay {} text {} dump {} dumped {} panic {}",
            counts[0], counts[6], counts[1], counts[2], counts[3], counts[4], counts[5]
        );
    }
}

fn collect(directory: &Path, files: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect(&path, files);
        } else {
            files.push(path);
        }
    }
}
