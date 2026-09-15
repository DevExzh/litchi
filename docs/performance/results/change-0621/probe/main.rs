//! Per-call-site attribution of `ReadAt::version()` observations on the
//! source-backed XLS open, one-cell and full-text paths.
//!
//! Change 0621 evidence probe. Counts are deterministic and independent of the
//! build profile; the probe wraps a `FileSource` in a `ReadAt` that captures a
//! backtrace at every observation and buckets by the nearest `litchi_*` frames.
//!
//! Usage: xls-observation-sites <fixture.xls> <open|one-cell|all-cells|full-text>
use std::{
    backtrace::Backtrace,
    collections::BTreeMap,
    io,
    sync::{Arc, Mutex},
};

use litchi_core::{FileSource, ReadAt, SourceVersion};

struct SiteCountingSource {
    inner: FileSource,
    sites: Mutex<BTreeMap<String, u64>>,
    total: Mutex<u64>,
}

/// Frames that are the fence helper itself, not the site that called it.
const HELPERS: &[&str] = &[
    "SiteCountingSource",
    "xls_observation_sites",
    "check_source_version",
    "ensure_current_parts",
    "ensure_current",
    "source_version",
    "check_text_state",
    "SourceCheckedTextSink",
    "::check",
    "Write>::write",
];

fn interesting(frame: &str) -> bool {
    (frame.contains("litchi_xls") || frame.contains("litchi_cfb") || frame.contains("litchi::"))
        && !frame.contains("core::ops::function")
}

impl SiteCountingSource {
    fn note(&self) {
        *self.total.lock().expect("total") += 1;
        let text = Backtrace::force_capture().to_string();
        // Pair each frame symbol with the source location the next line gives,
        // then keep the innermost litchi frames. The fence helpers are recorded
        // separately from the site that called them.
        let mut frames: Vec<(String, String)> = Vec::new();
        let mut pending: Option<String> = None;
        for line in text.lines() {
            let line = line.trim();
            if let Some(rest) = line.strip_prefix("at ") {
                if let Some(symbol) = pending.take() {
                    frames.push((symbol, rest.to_owned()));
                }
                continue;
            }
            if let Some((_, rest)) = line.split_once(": ") {
                if let Some(symbol) = pending.take() {
                    frames.push((symbol, String::new()));
                }
                pending = Some(rest.trim().to_owned());
            }
        }
        if let Some(symbol) = pending.take() {
            frames.push((symbol, String::new()));
        }
        let mut key: Vec<String> = Vec::new();
        let mut helper: Vec<String> = Vec::new();
        for (symbol, location) in &frames {
            if !interesting(symbol) {
                continue;
            }
            let location = short(location);
            let entry = format!("{}@{location}", trim_symbol(symbol));
            if HELPERS.iter().any(|name| symbol.contains(name)) {
                if helper.len() < 3 {
                    helper.push(entry);
                }
                continue;
            }
            key.push(entry);
            if key.len() == 2 {
                break;
            }
        }
        let mut label = helper.join(" <- ");
        if !label.is_empty() {
            label.push_str("  ||  ");
        }
        label.push_str(&key.join(" <- "));
        *self.sites.lock().expect("sites").entry(label).or_insert(0) += 1;
    }
}

fn short(location: &str) -> String {
    location
        .rsplit_once("/crates/")
        .map_or_else(|| location.to_owned(), |(_, tail)| tail.to_owned())
}

fn trim_symbol(symbol: &str) -> String {
    let symbol = symbol.split("::{closure").next().unwrap_or(symbol);
    let symbol = symbol.split("::<").next().unwrap_or(symbol);
    symbol.rsplit_once("::").map_or_else(
        || symbol.to_owned(),
        |(head, tail)| {
            let head = head.rsplit("::").next().unwrap_or(head);
            format!("{head}::{tail}")
        },
    )
}

impl std::fmt::Debug for SiteCountingSource {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("SiteCountingSource")
    }
}

impl ReadAt for SiteCountingSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read_at(offset, output)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        self.note();
        self.inner.version()
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let path = args.next().expect("fixture path");
    let operation = args.next().unwrap_or_else(|| "open".to_owned());
    let counting = Arc::new(SiteCountingSource {
        inner: FileSource::open(&path)?,
        sites: Mutex::new(BTreeMap::new()),
        total: Mutex::new(0),
    });
    let source: Arc<dyn ReadAt> = counting.clone();
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(source)?;
    let outcome = match operation.as_str() {
        "open" => format!("worksheets={}", workbook.worksheet_count()?),
        "one-cell" => {
            let sheet = workbook.worksheet_by_index(0)?.expect("worksheet 0");
            format!("cell={:?}", sheet.cell_value(1, 0)?.is_some())
        },
        "all-cells" => {
            let sheet = workbook.worksheet_by_index(0)?.expect("worksheet 0");
            let mut count = 0_u64;
            sheet.visit_cells(|_cell| {
                count += 1;
                Ok(())
            })?;
            format!("cells={count}")
        },
        "full-text" => match workbook.text() {
            Ok(text) => format!("text_bytes={}", text.len()),
            Err(error) => format!("refused={error}"),
        },
        other => panic!("unknown operation {other}"),
    };
    let total = *counting.total.lock().expect("total");
    let sites = counting.sites.lock().expect("sites").clone();
    println!("# fixture={path} operation={operation} outcome={outcome}");
    println!("# total_observations={total}");
    let mut rows: Vec<_> = sites.into_iter().collect();
    rows.sort_by(|left, right| right.1.cmp(&left.1));
    for (site, count) in rows {
        println!("{count:8}  {site}");
    }
    Ok(())
}
