//! Path round trip over the XLSX corpus (record 0010 §2.6, §10.3): every
//! node a depth-2 walk lists (sheets, named ranges, rows), and the first
//! cells of the first rows, resolves from its canonical and positional
//! paths back to itself.

use std::path::{Path, PathBuf};

use super::{fixture, get, path_of};
use crate::Code;
use crate::node::Node;

/// Children listed per parent, so a large sheet costs a bounded walk.
const CHILDREN: usize = 48;
/// Rows whose cells are walked too, per sheet.
const ROWS_WITH_CELLS: usize = 4;
/// Cells walked per such row.
const CELLS: usize = 8;

fn corpus(directory: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            corpus(&path, out);
        } else if path
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
        {
            out.push(path);
        }
    }
}

#[derive(Default)]
struct Tally {
    files: usize,
    nodes: usize,
    /// Files the unified opener refuses or opens as another format.
    skipped: Vec<String>,
    /// Leaf read failures while listing children (malformed parts).
    unreadable: Vec<String>,
    failures: Vec<String>,
}

fn round_trip(file: &crate::format::xlsx::Xlsx, node: &Node<'_>, origin: &Path, tally: &mut Tally) {
    tally.nodes += 1;
    let canonical = path_of(node);
    for input in [canonical.clone(), node.positional_path().to_string()] {
        match get(file, &input) {
            Ok(again) if path_of(&again) == canonical && again.kind() == node.kind() => {},
            Ok(again) => tally.failures.push(format!(
                "{}: {input} resolved to {} ({}), not {canonical} ({})",
                origin.display(),
                path_of(&again),
                again.kind(),
                node.kind()
            )),
            Err(error) => tally
                .failures
                .push(format!("{}: {input}: {error}", origin.display())),
        }
    }
}

/// Round-trips `node`, then walks its children: every parent to depth 2
/// lists up to [`CHILDREN`], and the first [`ROWS_WITH_CELLS`] rows of each
/// sheet list their first [`CELLS`] cells.
fn walk(
    file: &crate::format::xlsx::Xlsx,
    node: &Node<'_>,
    (depth, position): (usize, usize),
    origin: &Path,
    tally: &mut Tally,
) {
    round_trip(file, node, origin, tally);
    let limit = match depth {
        0 | 1 => CHILDREN,
        2 if node.kind() == "row" && position < ROWS_WITH_CELLS => CELLS,
        _ => return,
    };
    for (index, child) in node.children().take(limit).enumerate() {
        match child {
            Ok(child) => walk(file, &child, (depth + 1, index), origin, tally),
            Err(error) => {
                assert!(
                    !matches!(error.code(), Code::InternalError | Code::InvalidPath),
                    "{}: listing {}: {error}",
                    origin.display(),
                    path_of(node)
                );
                tally
                    .unreadable
                    .push(format!("{}: {}: {error}", origin.display(), path_of(node)));
                break;
            },
        }
    }
}

#[test]
fn corpus_paths_round_trip() {
    let mut files = Vec::new();
    corpus(&fixture(""), &mut files);
    files.sort();
    let mut tally = Tally::default();
    for origin in &files {
        let file = match crate::open(origin) {
            Ok(crate::File::Xlsx(file)) => file,
            Ok(other) => {
                tally
                    .skipped
                    .push(format!("{}: {other:?}", origin.display()));
                continue;
            },
            Err(error) => {
                tally.skipped.push(format!("{}: {error}", origin.display()));
                continue;
            },
        };
        tally.files += 1;
        match get(&file, "/") {
            Ok(root) => walk(&file, &root, (0, 0), origin, &mut tally),
            Err(error) => tally
                .failures
                .push(format!("{}: /: {error}", origin.display())),
        }
    }
    assert!(
        tally.failures.is_empty(),
        "{} of {} nodes failed to round-trip:\n{}",
        tally.failures.len(),
        tally.nodes,
        tally.failures.join("\n")
    );
    assert!(
        tally.files >= files.len() * 9 / 10,
        "only {} of {} files opened:\n{}",
        tally.files,
        files.len(),
        tally.skipped.join("\n")
    );
    assert!(
        tally.unreadable.len() <= tally.files / 20,
        "{} listings failed:\n{}",
        tally.unreadable.len(),
        tally.unreadable.join("\n")
    );
    assert!(tally.nodes > 2_000, "walked only {} nodes", tally.nodes);
}
