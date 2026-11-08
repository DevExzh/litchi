//! Batch equivalence for unsigned PPTX (record 0348, the rule of records
//! 0015 and 0314): a batch of sets publishes exactly what its items publish
//! when applied one at a time, or refuses with the same code at the same item.
//! A signed batch is an admission-level protected write: it refuses before
//! consuming its stream and therefore has no item scope.  Its behavior is
//! checked separately below rather than being forced into per-item equality.
//!
//! Coalesced PPTX runs plan table-cell and notes items against one-pass
//! indexes and take each item's change from the leaf; this checks that the
//! grouping is invisible. Unsigned fixtures under `test-data/`, and decks
//! made here whose cell and notes bodies hold values the text-body reader
//! refuses, run seeded pseudo-random sequences of sets (cells, cell
//! paragraphs and runs, notes paragraphs and runs, shape paragraphs;
//! repeated and missing targets) three ways: one `File::edit` closure, one
//! atomic `File::batch`, and a chain of single `File::set` calls. On both
//! owners the three must publish the same bytes, or refuse with the same
//! code at the same item. Signed corpus fixtures instead exercise the
//! owner-selected admission contract and are excluded from this equality.

#![cfg(feature = "pptx")]

use std::cell::Cell;
use std::io::{Read as _, Write as _};

use litchi_ops::{BatchItem, BatchMode, File, NewNode, OpenOptions, PropValue, Props, SourceMode};

fn fixtures() -> Vec<std::path::PathBuf> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut found = Vec::new();
    let mut stack = vec![root];
    while let Some(dir) = stack.pop() {
        for entry in std::fs::read_dir(&dir).expect("test-data") {
            let path = entry.expect("entry").path();
            if path.is_dir() {
                stack.push(path);
            } else if path.extension().and_then(|value| value.to_str()) == Some("pptx") {
                found.push(path);
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

#[derive(Clone, Debug)]
struct Set {
    path: String,
    props: Vec<(&'static str, PropValue<'static>)>,
}

#[derive(Debug, PartialEq, Eq)]
enum Outcome {
    Published(Vec<u8>),
    Refused(Option<usize>, String),
}

fn publish(file: &File) -> Outcome {
    match file.to_bytes() {
        Ok(bytes) => Outcome::Published(bytes),
        Err(error) => Outcome::Refused(None, format!("{:?}", error.code())),
    }
}

fn props(set: &Set) -> Props<'static> {
    let mut props = Props::new();
    for (key, value) in &set.props {
        props = props.with(*key, value.clone());
    }
    props
}

fn sequential(file: &File, sets: &[Set]) -> Outcome {
    let mut current = file.clone();
    for (index, set) in sets.iter().enumerate() {
        match current.set(set.path.as_str(), props(set)) {
            Ok(next) => current = next,
            Err(error) => return Outcome::Refused(Some(index), format!("{:?}", error.code())),
        }
    }
    publish(&current)
}

fn closure(file: &File, sets: &[Set]) -> Outcome {
    match file.edit(|tx| {
        for set in sets {
            tx.set(set.path.as_str(), props(set))?;
        }
        Ok(())
    }) {
        Ok(commit) => publish(commit.file()),
        Err(error) => admission(&error),
    }
}

/// Preserve the API's item scope.  A whole-batch admission refusal carries no
/// item; a refusal raised while reading an item carries that item's index.
fn admission(error: &litchi_ops::Error) -> Outcome {
    Outcome::Refused(error.item(), format!("{:?}", error.code()))
}

fn batch(file: &File, sets: &[Set]) -> Outcome {
    let items: Vec<BatchItem<'static>> = sets
        .iter()
        .map(|set| {
            let mut item = BatchItem::new("set");
            item.path = Some(set.path.clone().into());
            item.props = Some(props(set));
            item
        })
        .collect();
    match file.batch(items, BatchMode::Atomic) {
        Ok(commit) => publish(commit.file()),
        Err(error) => admission(&error),
    }
}

fn signed_batch_error_without_reading(
    file: &File,
    items: Vec<BatchItem<'static>>,
) -> litchi_ops::Error {
    let reads = Cell::new(0usize);
    let mut items = items.into_iter();
    let error = file
        .batch(
            std::iter::from_fn(|| {
                let item = items.next()?;
                reads.set(reads.get() + 1);
                Some(item)
            }),
            BatchMode::Atomic,
        )
        .expect_err("signed batch must refuse before consuming its stream");
    assert_eq!(reads.get(), 0, "signed preflight consumed a batch item");
    error
}

fn owners(bytes: &[u8]) -> Vec<File> {
    [SourceMode::Auto, SourceMode::Eager]
        .into_iter()
        .filter_map(|mode| {
            litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode)).ok()
        })
        .collect()
}

/// The set targets a file offers: cells (with paragraph and run steps),
/// notes paragraphs and runs, and shape paragraphs, plus missing ones.
fn targets(file: &File) -> Vec<String> {
    let paths = |kind: &str| -> Vec<String> {
        file.query(kind).map_or_else(
            |_| Vec::new(),
            |nodes| {
                nodes
                    .filter_map(Result::ok)
                    .map(|node| node.path().to_string())
                    .take(24)
                    .collect()
            },
        )
    };
    let mut out = Vec::new();
    for cell in paths("tc") {
        out.push(format!("{cell}/paragraph[1]"));
        out.push(format!("{cell}/paragraph[1]/run[1]"));
        out.push(format!("{cell}/paragraph[2]"));
        out.push(cell);
    }
    let slides = paths("slide").len();
    for slide in 1..=slides.min(4) {
        for paragraph in 1..=3 {
            out.push(format!("/slide[{slide}]/notes/paragraph[{paragraph}]"));
        }
        out.push(format!("/slide[{slide}]/notes/paragraph[1]/run[1]"));
        out.push(format!("/slide[{slide}]/shape[1]/paragraph[1]"));
        out.push(format!("/slide[{slide}]/table[1]/tr[1]/tc[9]/paragraph[1]"));
    }
    out
}

fn draw(rng: &mut Rng, item: usize, path: &str) -> Vec<(&'static str, PropValue<'static>)> {
    let cell = path.ends_with(']') && path.contains("/tc[") && !path.contains("paragraph");
    match rng.below(if cell { 6 } else { 5 }) {
        0 => vec![("bold", PropValue::from(rng.below(2) == 0))],
        1 => vec![("italic", PropValue::from(true))],
        2 => vec![(
            "align",
            PropValue::from(["left", "center", "right"][rng.below(3)]),
        )],
        3 if !path.contains("/run[") => vec![("text", PropValue::from(format!("Batch {item}")))],
        5 => vec![("fill", PropValue::from("FFEECC"))],
        _ => vec![("size", PropValue::from(12 + rng.below(20) as i64))],
    }
}

/// Every way agrees on `sets` for every owner of `bytes`.
fn check(name: &str, bytes: &[u8], sets_for: impl Fn(&File, u64) -> Vec<Set>) -> usize {
    let mut checked = 0;
    for file in owners(bytes) {
        for seed in 1..=3_u64 {
            let sets = sets_for(&file, seed);
            if sets.is_empty() {
                continue;
            }
            if file.is_signed() {
                let before = file.to_bytes().expect("publish signed source");
                assert_eq!(
                    batch(&file, &sets),
                    Outcome::Refused(None, "DocumentProtected".to_owned()),
                    "{name} seed {seed} signed batch admission: {sets:?}",
                );
                assert_eq!(
                    file.to_bytes().expect("republish signed source"),
                    before,
                    "{name} seed {seed} signed source changed",
                );
                checked += 1;
                continue;
            }
            let one = sequential(&file, &sets);
            assert_eq!(
                batch(&file, &sets),
                one,
                "{name} seed {seed} batch: {sets:?}"
            );
            assert_eq!(
                closure(&file, &sets),
                one,
                "{name} seed {seed} closure: {sets:?}"
            );
            checked += 1;
        }
    }
    checked
}

fn random_sets(file: &File, seed: u64) -> Vec<Set> {
    let targets = targets(file);
    if targets.is_empty() {
        return Vec::new();
    }
    let mut rng = Rng(seed.wrapping_mul(0x9e37_79b9_7f4a_7c15) | 1);
    (0..12)
        .map(|item| {
            let path = targets[rng.below(targets.len())].clone();
            let props = draw(&mut rng, item, &path);
            Set { path, props }
        })
        .collect()
}

#[test]
fn corpus_batches_keep_their_selected_contract() {
    let mut checked = 0;
    for path in fixtures() {
        let bytes = std::fs::read(&path).expect("fixture");
        checked += check(&path.display().to_string(), &bytes, random_sets);
    }
    assert!(checked >= 100, "{checked}");
}

#[test]
fn signed_batch_refuses_before_item_admission() {
    let bytes = std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx")
        .expect("signed PPTX fixture");
    let signed = litchi_ops::open_bytes(bytes.clone()).expect("open signed PPTX");
    assert!(signed.is_signed());

    let sets = vec![Set {
        path: "/slide[1]/shape[1]".to_owned(),
        props: vec![("text", PropValue::from("changed"))],
    }];

    let mut item = BatchItem::new("set");
    item.path = Some("/slide[1]/shape[1]".into());
    item.props = Some(Props::new().with("text", "changed"));
    let error = signed_batch_error_without_reading(&signed, vec![item]);
    assert_eq!(error.code(), litchi_ops::Code::DocumentProtected);
    assert_eq!(error.item(), None);

    assert_eq!(
        batch(&signed, &sets),
        Outcome::Refused(None, "DocumentProtected".to_owned()),
        "signed batch admission must retain its no-item scope",
    );
    assert_eq!(
        sequential(&signed, &sets),
        Outcome::Refused(Some(0), "DocumentProtected".to_owned()),
        "single-set refusal remains item-scoped",
    );
    assert_eq!(signed.to_bytes().expect("publish signed source"), bytes);
}

#[test]
fn signed_batch_admission_precedes_invalid_path_resolution() {
    let bytes = std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx")
        .expect("signed PPTX fixture");
    let sets = vec![Set {
        path: "/slide[1]/notes/paragraph[1]".to_owned(),
        props: vec![("align", PropValue::from("left"))],
    }];

    let mut owners_checked = 0;
    for file in owners(&bytes) {
        assert!(file.is_signed(), "fixture must remain signed");
        let before = file.to_bytes().expect("publish signed source");
        assert_eq!(
            batch(&file, &sets),
            Outcome::Refused(None, "DocumentProtected".to_owned()),
            "batch admission must protect the signed stream before path resolution",
        );
        let mut item = BatchItem::new("set");
        item.path = Some("/slide[1]/notes/paragraph[1]".into());
        item.props = Some(Props::new().with("align", "left"));
        let error = signed_batch_error_without_reading(&file, vec![item]);
        assert_eq!(error.code(), litchi_ops::Code::DocumentProtected);
        assert_eq!(error.item(), None);

        let single = Outcome::Refused(Some(0), "PathNotFound".to_owned());
        assert_eq!(sequential(&file, &sets), single);
        assert_eq!(closure(&file, &sets), single);
        assert_eq!(file.to_bytes().expect("republish signed source"), before);
        owners_checked += 1;
    }
    assert_eq!(owners_checked, 2, "Auto and Eager signed owners");
}

#[test]
fn signed_edit_resolves_a_later_invalid_path_after_staging() {
    let bytes = std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx")
        .expect("signed PPTX fixture");
    let sets = vec![
        Set {
            path: "/slide[1]/shape[1]".to_owned(),
            props: vec![("text", PropValue::from("signed batch changed"))],
        },
        Set {
            path: "/slide[1]/notes/paragraph[1]".to_owned(),
            props: vec![("align", PropValue::from("left"))],
        },
    ];

    let mut owners_checked = 0;
    for file in owners(&bytes) {
        assert!(file.is_signed(), "fixture must remain signed");
        let before = file.to_bytes().expect("publish signed source");
        assert_eq!(
            batch(&file, &sets),
            Outcome::Refused(None, "DocumentProtected".to_owned()),
            "batch admission must precede both target resolutions",
        );

        let mut first = BatchItem::new("set");
        first.path = Some("/slide[1]/shape[1]".into());
        first.props = Some(Props::new().with("text", "signed batch changed"));
        let mut second = BatchItem::new("set");
        second.path = Some("/slide[1]/notes/paragraph[1]".into());
        second.props = Some(Props::new().with("align", "left"));
        let error = signed_batch_error_without_reading(&file, vec![first, second]);
        assert_eq!(error.code(), litchi_ops::Code::DocumentProtected);
        assert_eq!(error.item(), None);

        assert_eq!(
            sequential(&file, &sets),
            Outcome::Refused(Some(0), "DocumentProtected".to_owned()),
            "single-set publication refuses the first signed change",
        );
        assert_eq!(
            closure(&file, &sets),
            Outcome::Refused(Some(1), "PathNotFound".to_owned()),
            "closure resolves the later invalid target before its final signed fence",
        );
        assert_eq!(file.to_bytes().expect("republish signed source"), before);
        owners_checked += 1;
    }
    assert_eq!(owners_checked, 2, "Auto and Eager signed owners");
}

/// `bytes` with the slide and notes parts rewritten by `edit`.
fn rewrite(bytes: &[u8], edit: impl Fn(&str, String) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes.to_vec())).expect("zip");
    let mut out = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut member = archive.by_index(index).expect("member");
        let name = member.name().expect("valid ZIP member name").into_owned();
        let mut data = Vec::new();
        member.read_to_end(&mut data).expect("read");
        if name.ends_with(".xml") && (name.contains("slides/") || name.contains("notesSlides/")) {
            data = edit(&name, String::from_utf8(data).expect("utf-8")).into_bytes();
        }
        out.start_file(name, zip::write::SimpleFileOptions::default())
            .expect("start");
        out.write_all(&data).expect("write");
    }
    out.finish().expect("finish").into_inner()
}

/// A deck with a 3×3 table and a three-paragraph notes page.
fn deck() -> Vec<u8> {
    let builder = litchi_pptx::builder::Builder::new().expect("builder");
    let base = builder.build().expect("build").to_bytes().expect("bytes");
    let file =
        litchi_ops::open_bytes_with(base, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("open");
    let mut file = file
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .expect("slide")
        .file;
    file = file
        .add(
            "/slide[1]",
            NewNode::new("table")
                .prop("rows", 3_i64)
                .prop("cols", 3_i64)
                .prop("data", "a1,b1,c1;a2,b2,c2;a3,b3,c3"),
        )
        .expect("table")
        .file;
    file = file
        .add("/slide[1]", NewNode::new("notes").prop("text", "n1"))
        .expect("notes")
        .file;
    for text in ["n2", "n3"] {
        file = file
            .add(
                "/slide[1]/notes",
                NewNode::new("paragraph").prop("text", text),
            )
            .expect("notes paragraph")
            .file;
    }
    file.to_bytes().expect("deck")
}

#[test]
fn refused_bodies_fail_the_batch_at_the_item_that_reads_them() {
    // Record 0348: a cell or notes body holding a value the text-body
    // reader refuses. Items that read that body refuse as their single set
    // does, at their own index; the others still coalesce.
    let base = deck();
    let refusals = [
        r#"<a:pPr algn="bogus"/>"#,
        r#"<a:pPr marL="wide"/>"#,
        r#"<a:pPr/><a:pPr/>"#,
    ];
    let cell = |row: usize, column: usize| format!("/slide[1]/table[1]/tr[{row}]/tc[{column}]");
    let notes = |paragraph: usize| format!("/slide[1]/notes/paragraph[{paragraph}]");
    let mut checked = 0;
    for refusal in refusals {
        // The second paragraph-opening of the slide is cell (1, 2); of the
        // notes part, its second body paragraph.
        let bytes = rewrite(&base, |name, xml| {
            let mut parts = xml.splitn(3, "<a:p>");
            let (Some(head), Some(first), Some(rest)) = (parts.next(), parts.next(), parts.next())
            else {
                return xml;
            };
            let _ = name;
            format!("{head}<a:p>{first}<a:p>{refusal}{rest}")
        });
        let sets = |_: &File, seed: u64| -> Vec<Set> {
            let bold = vec![("bold", PropValue::from(true))];
            let mut sets = vec![
                Set {
                    path: format!("{}/paragraph[1]", cell(2, 2)),
                    props: bold.clone(),
                },
                Set {
                    path: cell(1, 2),
                    props: bold.clone(),
                },
                Set {
                    path: format!("{}/paragraph[1]", cell(3, 3)),
                    props: bold.clone(),
                },
            ];
            let tail = match seed {
                1 => vec![format!("{}/paragraph[1]", cell(1, 2))],
                2 => vec![format!("{}/paragraph[1]/run[1]", cell(1, 2))],
                _ => vec![notes(1), notes(3)],
            };
            sets.extend(tail.into_iter().map(|path| Set {
                path,
                props: bold.clone(),
            }));
            sets
        };
        checked += check(refusal, &bytes, sets);
        // The refusal is the item's own.
        for file in owners(&bytes) {
            let probe = vec![
                Set {
                    path: format!("{}/paragraph[1]", cell(2, 2)),
                    props: vec![("bold", PropValue::from(true))],
                },
                Set {
                    path: format!("{}/paragraph[1]", cell(1, 2)),
                    props: vec![("bold", PropValue::from(true))],
                },
            ];
            assert!(
                matches!(sequential(&file, &probe), Outcome::Refused(Some(1), _)),
                "{refusal}"
            );
        }
    }
    assert!(checked >= 18, "{checked}");
}
