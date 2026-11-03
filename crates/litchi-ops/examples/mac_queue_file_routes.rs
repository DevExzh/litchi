//! Writes the Litchi-mutated packages of the `litchi_ops::File` rows of the
//! macOS verification queue (`docs/unified-ops/0363-mac-queue.md`, wave D7;
//! record 0430).
//!
//! ```text
//! cargo run -p litchi-ops --features pages,keynote,numbers \
//!     --example mac_queue_file_routes -- <output-directory>
//! ```
//!
//! For each format it writes the root cell-comment set and clear and the
//! chart point-size set and reset through `File::set`; for Numbers also an
//! empty-cell value write (the shared TST cell writer of the W1-B row); and
//! two Pages documents built with fresh `Identity::generate()` identities
//! (the `Identity::from_parts` row). It then writes one JSON line per
//! package to `<output-directory>/file-routes.jsonl` (stem, format, fixture,
//! selector, operation, the value read back after reopening).
//! `tools/mac_queue/generate.py` folds the lines into the queue manifest.

use std::io::Write as _;
use std::path::{Path, PathBuf};

use litchi_ops::{File, PropValue, open_bytes, props};

type Result<T = ()> = std::result::Result<T, Box<dyn std::error::Error>>;

/// The comment text the set route writes.
const COMMENT: &str = "Litchi 0430 native check";
/// The chart point size the set route writes.
const FONT_SIZE: &str = "17pt";

struct Target {
    /// Stem prefix of the written packages.
    format: &'static str,
    extension: &'static str,
    comment_fixture: &'static str,
    comment_cell: &'static str,
    chart_fixture: &'static str,
    chart: &'static str,
}

const TARGETS: [Target; 3] = [
    Target {
        format: "pages",
        extension: "pages",
        comment_fixture: "pages/body-table-cells-native.pages",
        comment_cell: "/body/tbl[1]/tr[2]/tc[2]",
        chart_fixture: "pages/chart-grid-authored-native.pages",
        chart: "/body/chart[1]",
    },
    Target {
        format: "keynote",
        extension: "key",
        comment_fixture: "keynote/slide-table-read-native.key",
        comment_cell: "/slide[1]/table[1]/tr[2]/tc[4]",
        chart_fixture: "keynote/chart-grid-authored-native.key",
        chart: "/slide[1]/chart[1]",
    },
    Target {
        format: "numbers",
        extension: "numbers",
        comment_fixture: "numbers/comment-reader-native-created.numbers",
        comment_cell: "/sheet[1]/table[1]/B2",
        chart_fixture: "numbers/chart-grid-authored-native.numbers",
        chart: "/sheet[1]/chart[1]",
    },
];

fn test_data() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork")
}

fn property(file: &File, path: &str, key: &str) -> Result<Option<String>> {
    Ok(file
        .get(path)?
        .view()
        .format()?
        .find(|property| property.key() == key)
        .map(|property| property.to_string()))
}

fn reopen(file: &File) -> Result<File> {
    Ok(open_bytes(file.to_bytes()?)?)
}

fn json_string(value: &str) -> String {
    let mut out = String::from("\"");
    for ch in value.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            ch if u32::from(ch) < 0x20 => out.push_str(&format!("\\u{:04x}", u32::from(ch))),
            ch => out.push(ch),
        }
    }
    out.push('"');
    out
}

struct Writer {
    directory: PathBuf,
    lines: Vec<String>,
}

impl Writer {
    #[allow(clippy::too_many_arguments, reason = "one JSON line, field by field")]
    fn emit(
        &mut self,
        target: &Target,
        operation: &str,
        fixture: &str,
        selector: &str,
        key: &str,
        file: &File,
    ) -> Result {
        let stem = format!("{}-{operation}", target.format);
        let bytes = file.to_bytes()?;
        std::fs::write(
            self.directory
                .join(format!("{stem}.file.{}", target.extension)),
            &bytes,
        )?;
        let read_back = property(&open_bytes(bytes)?, selector, key)?;
        let read_back = read_back.as_deref().map_or("null".to_owned(), json_string);
        self.lines.push(format!(
            "{{\"stem\": {}, \"format\": {}, \"extension\": {}, \"operation\": {}, \"fixture\": {}, \"selector\": {}, \"property\": {}, \"read_back\": {read_back}}}",
            json_string(&stem),
            json_string(target.format),
            json_string(target.extension),
            json_string(operation),
            json_string(&format!("test-data/iwork/{fixture}")),
            json_string(selector),
            json_string(key),
        ));
        Ok(())
    }
}

fn run(target: &Target, writer: &mut Writer) -> Result {
    let root = test_data();

    // Root cell comment: set (create or replace), then clear. A cell that
    // has no comment in the fixture is cleared after the set, as the queue
    // row says ("after a root-only comment").
    let source = open_bytes(std::fs::read(root.join(target.comment_fixture))?)?;
    let cell = target.comment_cell;
    let had_comment = property(&source, cell, "comment")?.is_some();
    let set = source.set(cell, props! { "comment" => COMMENT })?;
    writer.emit(
        target,
        "cell-comment-set",
        target.comment_fixture,
        cell,
        "comment",
        &set,
    )?;
    let clear_input = if had_comment { source } else { reopen(&set)? };
    let cleared = clear_input.set(cell, props! { "comment" => PropValue::Null })?;
    writer.emit(
        target,
        "cell-comment-clear",
        target.comment_fixture,
        cell,
        "comment",
        &cleared,
    )?;

    // Chart point size: set, then reset the set package.
    let source = open_bytes(std::fs::read(root.join(target.chart_fixture))?)?;
    let chart = target.chart;
    let set = source.set(chart, props! { "font.size" => FONT_SIZE })?;
    writer.emit(
        target,
        "chart-font-size-set",
        target.chart_fixture,
        chart,
        "font.size",
        &set,
    )?;
    let reset = reopen(&set)?.set(chart, props! { "font.size" => PropValue::Null })?;
    writer.emit(
        target,
        "chart-font-size-reset",
        target.chart_fixture,
        chart,
        "font.size",
        &reset,
    )?;

    if target.format == "numbers" {
        // An empty cell of table 1 (record 0371's W1-B cell-writer fixes).
        let fixture = "numbers/basic.numbers";
        let cell = "/sheet[1]/table[1]/G22";
        let source = open_bytes(std::fs::read(root.join(fixture))?)?;
        let set = source.set(cell, props! { "value" => COMMENT })?;
        writer.emit(target, "cell-value-set", fixture, cell, "value", &set)?;
    }
    Ok(())
}

/// Two Pages documents with fresh identities: Pages must treat them as
/// different documents.
fn builder_identities(writer: &mut Writer) -> Result {
    for suffix in ["a", "b"] {
        let package = litchi_pages::builder::Builder::new()
            .identity(litchi_pages::builder::Identity::generate())
            .body_text(format!("Litchi 0430 identity {suffix}"))
            .build()?;
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        let stem = format!("pages-builder-identity-{suffix}");
        std::fs::write(writer.directory.join(format!("{stem}.file.pages")), &bytes)?;
        // The unified reopen is the readback: the document must open.
        drop(open_bytes(bytes)?);
        writer.lines.push(format!(
            "{{\"stem\": {}, \"format\": \"pages\", \"extension\": \"pages\", \"operation\": \"builder-identity\", \"fixture\": \"builder\", \"selector\": \"-\", \"property\": \"-\", \"read_back\": null}}",
            json_string(&stem),
        ));
    }
    Ok(())
}

fn main() -> Result {
    let directory = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .ok_or("usage: mac_queue_file_routes <output-directory>")?;
    std::fs::create_dir_all(&directory)?;
    let mut writer = Writer {
        directory,
        lines: Vec::new(),
    };
    for target in &TARGETS {
        run(target, &mut writer).map_err(|error| format!("{}: {error}", target.format))?;
    }
    builder_identities(&mut writer).map_err(|error| format!("builder: {error}"))?;
    let mut out = std::fs::File::create(writer.directory.join("file-routes.jsonl"))?;
    for line in &writer.lines {
        writeln!(out, "{line}")?;
    }
    Ok(())
}
