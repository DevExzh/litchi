//! Formula-evaluation parity against the cached results stored in the
//! spreadsheet corpus (record 0266).
//!
//! Every XLSX, XLSM, XLSB and XLS file under the corpus root is opened and
//! recalculated with `File::recalculate`; each formula's computed value is
//! compared with the value the producing application cached. The JSON
//! report on standard output has the totals by status and format (byte-identical
//! copies counted once), the
//! per-function counts (a formula counts once for every distinct function it
//! calls; formulas without a function count under `(operators)`), the files
//! that could not be evaluated, and the mismatches.
//!
//! ```text
//! cargo run --release -p litchi-ops --features xlsx,xlsb,xls,eval \
//!     --example eval_parity -- [corpus-root] > parity.json
//! ```
//!
//! The corpus root defaults to `LITCHI_TEST_DATA` or the workspace's
//! `test-data` directory.

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use litchi_ops::calc::{Options, Status, Value};
use serde_json::{Value as Json, json};

#[derive(Default, Clone, Copy)]
struct Tally {
    formulas: usize,
    matches: usize,
    mismatches: usize,
    uncached: usize,
    unsupported: usize,
    volatile: usize,
}

impl Tally {
    fn add(&mut self, status: Status) {
        self.formulas += 1;
        match status {
            Status::Match => self.matches += 1,
            Status::Mismatch => self.mismatches += 1,
            Status::Uncached => self.uncached += 1,
            Status::Unsupported => self.unsupported += 1,
            Status::Volatile => self.volatile += 1,
            _ => {},
        }
    }

    fn json(self) -> Json {
        let compared = self.matches + self.mismatches;
        json!({
            "formulas": self.formulas,
            "match": self.matches,
            "mismatch": self.mismatches,
            "uncached": self.uncached,
            "unsupported": self.unsupported,
            "volatile": self.volatile,
            "pass_rate_of_compared": rate(self.matches, compared),
            "pass_rate_of_cached": rate(self.matches, self.formulas - self.uncached),
        })
    }
}

fn rate(part: usize, whole: usize) -> Json {
    if whole == 0 {
        Json::Null
    } else {
        #[allow(clippy::cast_precision_loss, reason = "corpus counts are small")]
        let value = part as f64 / whole as f64;
        json!((value * 10_000.0).round() / 10_000.0)
    }
}

fn render(value: Option<&Value>) -> Json {
    match value {
        None => Json::Null,
        Some(Value::Empty) => json!(""),
        Some(Value::Bool(value)) => json!(value),
        Some(Value::Number(value)) => json!(value),
        Some(Value::Text(text)) => json!(text),
        Some(Value::Error(error)) => json!(error.as_str()),
        Some(_) => json!("?"),
    }
}

/// The distinct functions a formula calls, upper-cased, without the
/// `_xlfn.`/`_xlws.` prefixes, outside string literals.
fn functions(formula: &str) -> Vec<String> {
    let bytes = formula.as_bytes();
    let mut names = Vec::new();
    let mut quoted = false;
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'"' {
            quoted = !quoted;
            index += 1;
            continue;
        }
        if quoted || !(byte.is_ascii_alphabetic() || byte == b'_') {
            index += 1;
            continue;
        }
        let start = index;
        while index < bytes.len()
            && (bytes[index].is_ascii_alphanumeric() || matches!(bytes[index], b'_' | b'.'))
        {
            index += 1;
        }
        if bytes.get(index) == Some(&b'(') {
            let mut name = formula[start..index].to_ascii_uppercase();
            for prefix in ["_XLFN.", "_XLWS."] {
                if let Some(rest) = name.strip_prefix(prefix) {
                    name = rest.to_owned();
                }
            }
            if !names.contains(&name) {
                names.push(name);
            }
        }
    }
    if names.is_empty() {
        names.push("(operators)".to_owned());
    }
    names
}

fn corpus(root: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    let mut entries: Vec<_> = entries
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .collect();
    entries.sort();
    for path in entries {
        if path.is_dir() {
            corpus(&path, out);
        } else if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                ["xlsx", "xlsm", "xlsb", "xls"]
                    .iter()
                    .any(|known| known.eq_ignore_ascii_case(extension))
            })
        {
            out.push(path);
        }
    }
}

fn column_label(mut column: u32) -> String {
    let mut label = String::new();
    column += 1;
    while column > 0 {
        let digit = u8::try_from((column - 1) % 26).unwrap_or(0);
        label.insert(0, char::from(b'A' + digit));
        column = (column - 1) / 26;
    }
    label
}

fn main() {
    let root = std::env::args().nth(1).map_or_else(
        || {
            std::env::var_os("LITCHI_TEST_DATA").map_or_else(
                || Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data"),
                PathBuf::from,
            )
        },
        PathBuf::from,
    );
    let mut files = Vec::new();
    corpus(&root, &mut files);
    let options = Options::new();
    let mut total = Tally::default();
    let mut by_format: BTreeMap<String, Tally> = BTreeMap::new();
    let mut by_function: BTreeMap<String, Tally> = BTreeMap::new();
    let mut by_file: BTreeMap<String, Tally> = BTreeMap::new();
    let mut unsupported_samples: BTreeMap<String, Vec<String>> = BTreeMap::new();
    let mut failures = Vec::new();
    let mut mismatches = Vec::new();
    let mut evaluated_files = 0usize;
    let mut seen = std::collections::HashSet::new();
    let mut duplicates = Vec::new();
    for path in &files {
        let relative = path
            .strip_prefix(&root)
            .unwrap_or(path)
            .display()
            .to_string();
        // The corpus holds some byte-identical copies; count each once.
        let Ok(bytes) = std::fs::read(path) else {
            failures.push(json!({"file": relative, "stage": "read"}));
            continue;
        };
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        std::hash::Hash::hash(&bytes, &mut hasher);
        if !seen.insert((bytes.len(), std::hash::Hasher::finish(&hasher))) {
            duplicates.push(relative);
            continue;
        }
        let file = match litchi_ops::open(path) {
            Ok(file) => file,
            Err(error) => {
                failures.push(
                    json!({"file": relative, "stage": "open", "code": error.code().as_str()}),
                );
                continue;
            },
        };
        let format = file.format().map_or_else(
            || "unknown".to_owned(),
            |format| format!("{format:?}").to_lowercase(),
        );
        let report = match file.recalculate(&options) {
            Ok(report) => report,
            Err(error) => {
                failures.push(json!({
                    "file": relative,
                    "stage": "recalculate",
                    "code": error.code().as_str(),
                    "message": error.to_string(),
                }));
                continue;
            },
        };
        if !report.formulas().is_empty() {
            evaluated_files += 1;
        }
        let names = file
            .workbook()
            .and_then(|workbook| workbook.worksheet_names())
            .unwrap_or_default();
        for formula in report.formulas() {
            total.add(formula.status());
            by_file
                .entry(relative.clone())
                .or_default()
                .add(formula.status());
            if formula.status() == Status::Unsupported {
                for name in functions(formula.text()) {
                    let samples = unsupported_samples.entry(name).or_default();
                    if samples.len() < 3 && !samples.iter().any(|sample| sample == formula.text()) {
                        samples.push(formula.text().to_owned());
                    }
                }
            }
            by_format
                .entry(format.clone())
                .or_default()
                .add(formula.status());
            for name in functions(formula.text()) {
                by_function.entry(name).or_default().add(formula.status());
            }
            if formula.status() == Status::Mismatch {
                mismatches.push(json!({
                    "file": relative,
                    "sheet": names.get(formula.sheet()).cloned().unwrap_or_default(),
                    "cell": format!("{}{}", column_label(formula.column()), formula.row() + 1),
                    "formula": formula.text(),
                    "cached": render(formula.cached()),
                    "computed": render(formula.computed()),
                }));
            }
        }
    }
    let functions: BTreeMap<String, Json> = by_function
        .into_iter()
        .map(|(name, tally)| (name, tally.json()))
        .collect();
    let formats: BTreeMap<String, Json> = by_format
        .into_iter()
        .map(|(name, tally)| (name, tally.json()))
        .collect();
    let files_detail: BTreeMap<String, Json> = by_file
        .into_iter()
        .map(|(name, tally)| (name, tally.json()))
        .collect();
    let report = json!({
        "schema": "litchi.eval-parity.v1",
        "record": "0266",
        "tolerance": "relative 1e-9 (|a-b| <= 1e-9 * max(1,|a|,|b|)); text, booleans and errors exact; empty equals empty text",
        "files": files.len(),
        "duplicates_skipped": duplicates,
        "files_with_formulas": evaluated_files,
        "total": total.json(),
        "formats": formats,
        "functions": functions,
        "per_file": files_detail,
        "unsupported_samples": unsupported_samples,
        "failures": failures,
        "mismatches": mismatches,
    });
    println!(
        "{}",
        serde_json::to_string_pretty(&report).expect("serialize the parity report")
    );
}
