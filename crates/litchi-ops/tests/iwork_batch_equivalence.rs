//! Batch equivalence over the iWork corpus (record 0314): a batch of sets
//! publishes exactly what its items publish when applied one at a time.
//!
//! Coalesced runs (Pages paragraph texts, Keynote text, frame and
//! visibility sets, Numbers cell values) group items into one leaf
//! transaction; this property checks that the grouping is invisible. For
//! every Pages, Keynote and Numbers fixture under `test-data/`, seeded
//! pseudo-random sequences of sets are drawn from the elements the fixture
//! holds (repeated targets, values of every kind, items that refuse, and
//! items no run stages, which close the run). Each sequence runs three
//! ways: one `File::edit` closure, one atomic `File::batch`, and a chain of
//! single `File::set` calls on successive handles. The three must publish
//! the same bytes, or refuse with the same code at the same item.

#![cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]

use litchi_ops::{BatchItem, BatchMode, File, PropValue, Props, open_bytes};

fn fixtures(extension: &str) -> Vec<std::path::PathBuf> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut found = Vec::new();
    let mut stack = vec![root];
    while let Some(dir) = stack.pop() {
        for entry in std::fs::read_dir(&dir).expect("test-data") {
            let path = entry.expect("entry").path();
            if path.extension().and_then(|value| value.to_str()) == Some(extension) {
                if path.is_file() {
                    found.push(path);
                }
            } else if path.is_dir() {
                stack.push(path);
            }
        }
    }
    found.sort();
    found
}

/// A small deterministic generator (xorshift64*).
struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 >> 12;
        self.0 ^= self.0 << 25;
        self.0 ^= self.0 >> 27;
        self.0.wrapping_mul(0x2545_f491_4f6c_dd1d)
    }

    fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }
}

/// One set: a path and its properties.
#[derive(Clone)]
struct Set {
    path: String,
    props: Vec<(&'static str, PropValue<'static>)>,
}

impl std::fmt::Debug for Set {
    fn fmt(&self, out: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(out, "{}", self.path)?;
        for (key, value) in &self.props {
            write!(out, " {key}={value}")?;
        }
        Ok(())
    }
}

/// What a sequence published: the bytes, or the refusing item and code.
#[derive(Debug, PartialEq, Eq)]
enum Outcome {
    Published(Vec<u8>),
    Refused(Option<usize>, String),
}

impl std::fmt::Display for Outcome {
    fn fmt(&self, out: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Published(bytes) => write!(out, "{} bytes", bytes.len()),
            Self::Refused(item, code) => write!(out, "item {item:?} refused: {code}"),
        }
    }
}

fn refused(error: &litchi_ops::Error, fallback: Option<usize>) -> Outcome {
    Outcome::Refused(error.item().or(fallback), format!("{:?}", error.code()))
}

fn publish(file: &File) -> Outcome {
    match file.to_bytes() {
        Ok(bytes) => Outcome::Published(bytes),
        Err(error) => refused(&error, None),
    }
}

fn sequential(file: &File, sets: &[Set]) -> Outcome {
    let mut current = file.clone();
    for (index, set) in sets.iter().enumerate() {
        match current.set(set.path.as_str(), set.props.clone()) {
            Ok(next) => current = next,
            Err(error) => return Outcome::Refused(Some(index), format!("{:?}", error.code())),
        }
    }
    publish(&current)
}

fn closure(file: &File, sets: &[Set]) -> Outcome {
    let edited = file.edit(|tx| {
        for set in sets {
            tx.set(set.path.as_str(), set.props.clone())?;
        }
        Ok(())
    });
    match edited {
        Ok(commit) => publish(commit.file()),
        Err(error) => refused(&error, None),
    }
}

fn batch(file: &File, sets: &[Set]) -> Outcome {
    let items: Vec<BatchItem<'static>> = sets
        .iter()
        .map(|set| {
            let mut item = BatchItem::new("set");
            item.path = Some(set.path.clone().into());
            let mut props = Props::new();
            for (key, value) in &set.props {
                props = props.with(*key, value.clone());
            }
            item.props = Some(props);
            item
        })
        .collect();
    match file.batch(items, BatchMode::Atomic) {
        Ok(commit) => publish(commit.file()),
        Err(error) => refused(&error, None),
    }
}

/// The canonical paths of every element `kind` names.
fn paths(file: &File, kind: &str) -> Vec<String> {
    file.query(kind).map_or_else(
        |_| Vec::new(),
        |nodes| {
            nodes
                .filter_map(Result::ok)
                .map(|node| node.path().to_string())
                .collect()
        },
    )
}

/// One generator: the kind of element it targets and how it draws props.
type Draw = fn(&mut Rng, usize) -> Vec<(&'static str, PropValue<'static>)>;

#[cfg(any(feature = "pages", feature = "keynote"))]
fn text(rng: &mut Rng, item: usize) -> Vec<(&'static str, PropValue<'static>)> {
    let value = match rng.below(8) {
        0 => String::new(),
        1 => "two\nlines".to_owned(),
        2 => format!("Batch {item} \u{e9}\u{1f600}"),
        _ => format!("Batch {item}"),
    };
    vec![("text", PropValue::from(value))]
}

#[cfg(feature = "keynote")]
fn frame(rng: &mut Rng, item: usize) -> Vec<(&'static str, PropValue<'static>)> {
    let key = ["x", "y", "width", "height", "rotation"][rng.below(5)];
    let value = match (key, rng.below(6)) {
        ("width" | "height", 0) => "0pt".to_owned(),
        ("rotation", _) => format!("{}", item * 7 % 360),
        _ => format!("{}pt", 10 + item * 3),
    };
    let mut props = vec![(key, PropValue::from(value))];
    if rng.below(4) == 0 {
        props.push(("y", PropValue::from(format!("{}pt", 20 + item))));
    }
    props
}

#[cfg(feature = "keynote")]
fn hidden(rng: &mut Rng, _item: usize) -> Vec<(&'static str, PropValue<'static>)> {
    vec![("hidden", PropValue::from(rng.below(2) == 0))]
}

#[cfg(feature = "keynote")]
fn bold(rng: &mut Rng, _item: usize) -> Vec<(&'static str, PropValue<'static>)> {
    vec![("bold", PropValue::from(rng.below(2) == 0))]
}

#[cfg(feature = "pages")]
fn font_bold(rng: &mut Rng, _item: usize) -> Vec<(&'static str, PropValue<'static>)> {
    vec![("font.bold", PropValue::from(rng.below(2) == 0))]
}

#[cfg(feature = "numbers")]
fn value(rng: &mut Rng, item: usize) -> Vec<(&'static str, PropValue<'static>)> {
    let value = match rng.below(6) {
        0 => PropValue::Null,
        1 => PropValue::from(rng.below(2) == 0),
        2 => PropValue::from(format!("Batch {item}")),
        3 => PropValue::from(f64::NAN),
        _ => PropValue::from(item as f64 * 1.5),
    };
    vec![("value", value)]
}

#[cfg(feature = "numbers")]
fn formula(_rng: &mut Rng, item: usize) -> Vec<(&'static str, PropValue<'static>)> {
    vec![("formula", PropValue::from(format!("={item}+1")))]
}

/// Sequences drawn per fixture: 4, or `LITCHI_BATCH_EQUIVALENCE_CASES`.
fn cases() -> usize {
    std::env::var("LITCHI_BATCH_EQUIVALENCE_CASES")
        .ok()
        .and_then(|cases| cases.parse().ok())
        .unwrap_or(4)
}

/// Draw [`cases`] sequences per fixture and compare the three routes.
fn check(extension: &str, draws: &[(&str, Draw)], extra: fn(&File) -> Vec<(String, Draw)>) {
    let mut seed = 0x0314_u64;
    let mut sequences = 0;
    let mut published = 0;
    let mut failures = Vec::new();
    for fixture in fixtures(extension) {
        let Ok(bytes) = std::fs::read(&fixture) else {
            continue;
        };
        let Ok(file) = open_bytes(bytes) else {
            continue;
        };
        let mut targets: Vec<(String, Draw)> = Vec::new();
        for (kind, draw) in draws {
            // A few elements of each kind keep the sequences short and
            // their targets repeated.
            for path in paths(&file, kind).into_iter().take(6) {
                targets.push((path, *draw));
            }
        }
        targets.extend(extra(&file));
        if targets.is_empty() {
            continue;
        }
        for _ in 0..cases() {
            seed = seed.wrapping_add(0x9e37_79b9_7f4a_7c15);
            let mut rng = Rng(seed | 1);
            let length = 1 + rng.below(9);
            let sets: Vec<Set> = (0..length)
                .map(|item| {
                    let (path, draw) = &targets[rng.below(targets.len())];
                    Set {
                        path: path.clone(),
                        props: draw(&mut rng, item),
                    }
                })
                .collect();
            let one_by_one = sequential(&file, &sets);
            let edited = closure(&file, &sets);
            if edited != one_by_one {
                failures.push(format!(
                    "{}: edit {edited} against one by one {one_by_one}: {sets:?}",
                    fixture.display()
                ));
            }
            let batched = batch(&file, &sets);
            if batched != one_by_one {
                failures.push(format!(
                    "{}: batch {batched} against one by one {one_by_one}: {sets:?}",
                    fixture.display()
                ));
            }
            sequences += 1;
            published += usize::from(matches!(one_by_one, Outcome::Published(_)));
        }
    }
    eprintln!("{extension}: {sequences} sequences, {published} published");
    assert!(
        failures.is_empty(),
        "{} of {sequences} sequences differ:\n{}",
        failures.len(),
        failures.join("\n")
    );
    assert!(
        sequences >= 20 && published >= 10,
        "{sequences} {published}"
    );
}

#[cfg(feature = "keynote")]
fn none(_: &File) -> Vec<(String, Draw)> {
    Vec::new()
}

#[cfg(feature = "pages")]
#[test]
fn pages_batches_equal_their_items_one_at_a_time() {
    check(
        "pages",
        &[("p", text), ("p", font_bold), ("tc", text)],
        |file| {
            // The direct spelling the coalesced run takes, past the end too.
            let count = paths(file, "p").len();
            [1, 2, count, count + 1]
                .into_iter()
                .filter(|&index| index > 0)
                .map(|index| -> (String, Draw) { (format!("/body/p[{index}]"), text) })
                .collect()
        },
    );
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_batches_equal_their_items_one_at_a_time() {
    check(
        "key",
        &[
            ("placeholder", text),
            ("paragraph", text),
            ("shape", text),
            ("shape", frame),
            ("placeholder", frame),
            ("picture", frame),
            ("slide", hidden),
            ("run", bold),
            ("tc", text),
            ("notes", text),
        ],
        none,
    );
}

#[cfg(feature = "numbers")]
#[test]
fn numbers_batches_equal_their_items_one_at_a_time() {
    check("numbers", &[("cell", value), ("cell", formula)], |file| {
        // Bare A1 cells of the first table, stored or not, and one past
        // its grid.
        let Some(table) = paths(file, "table").into_iter().next() else {
            return Vec::new();
        };
        ["A1", "B2", "C3", "A2", "ZZ999"]
            .into_iter()
            .map(|cell| -> (String, Draw) { (format!("{table}/{cell}"), value) })
            .collect()
    });
}
