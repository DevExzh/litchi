//! Direct table-cell writes (record 0312): a one-value set of a Keynote
//! slide-table cell or a Pages body-table cell named by its canonical path
//! (`/slide[N]/table[M]/tr[R]/tc[C]`, `/body/tbl[N]/tr[R]/tc[C]`) commits
//! without reading the table. These tests check, on every corpus fixture
//! holding a table, that it publishes exactly what the general route
//! publishes (the same cell spelled with the `row`/`cell` aliases resolves
//! through the object tree) and refuses with the same code where the
//! general route refuses: cells past the last row and column, a table past
//! the last, and every value kind.

#![cfg(any(feature = "pages", feature = "keynote"))]

use litchi_ops::{File, PropValue, open_bytes};

type Entry = (&'static str, PropValue<'static>);

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

/// The published bytes, or the refusal code.
fn outcome(file: &File, path: &str, entry: &Entry) -> Result<Vec<u8>, String> {
    file.set(path, [entry.clone()])
        .map(|file| file.to_bytes().expect("publish"))
        .map_err(|error| format!("{:?}", error.code()))
}

fn values() -> Vec<Entry> {
    vec![
        ("text", PropValue::from("Direct")),
        ("value", PropValue::from(42.5)),
        ("value", PropValue::from(true)),
        ("value", PropValue::Null),
    ]
}

/// Compare the direct and the general route on the cells of one table:
/// the first, the last and one past the last row and column.
fn compare(
    bytes: &[u8],
    direct: &dyn Fn(u64, u64) -> String,
    general: &dyn Fn(u64, u64) -> String,
    rows: u64,
    columns: u64,
) -> (usize, usize) {
    let mut compared = (0, 0);
    let mut positions = vec![1, rows, rows + 1];
    positions.dedup();
    let mut columns_at = vec![1, columns, columns + 1];
    columns_at.dedup();
    for (index, props) in values().iter().enumerate() {
        for &row in &positions {
            for &column in &columns_at {
                // Every value kind on the first cell; the text elsewhere.
                if index > 0 && (row, column) != (1, 1) {
                    continue;
                }
                let file = open_bytes(bytes.to_vec()).expect("open");
                let direct_path = direct(row, column);
                let left = outcome(&file, &direct_path, props);
                let file = open_bytes(bytes.to_vec()).expect("open");
                let right = outcome(&file, &general(row, column), props);
                assert!(
                    left == right,
                    "{direct_path}: {:?} against {:?}",
                    left.as_ref().map(Vec::len),
                    right.as_ref().map(Vec::len)
                );
                compared.0 += 1;
                compared.1 += usize::from(left.is_ok());
            }
        }
    }
    compared
}

fn count(file: &File, path: &str, key: &str) -> u64 {
    file.get(path)
        .expect("table")
        .view()
        .format()
        .expect("format")
        .find(|prop| prop.key() == key)
        .map_or(0, |prop| prop.to_string().parse().expect("count"))
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_direct_cells_match_the_general_route() {
    let mut compared = (0, 0);
    for path in fixtures("key") {
        let bytes = std::fs::read(&path).expect("fixture");
        let Ok(file) = open_bytes(bytes.clone()) else {
            continue;
        };
        let Ok(tables) = file.query("table") else {
            continue;
        };
        let tables: Vec<String> = tables
            .filter_map(Result::ok)
            .map(|node| node.path().to_string())
            .collect();
        for table in tables {
            let rows = count(&file, &table, "rows");
            let columns = count(&file, &table, "cols");
            let direct = |row: u64, column: u64| format!("{table}/tr[{row}]/tc[{column}]");
            let general = |row: u64, column: u64| format!("{table}/row[{row}]/cell[{column}]");
            let (cells, published) = compare(&bytes, &direct, &general, rows, columns);
            compared.0 += cells;
            compared.1 += published;
        }
        // A table past the slide's last.
        let file = open_bytes(bytes.clone()).expect("open");
        assert_eq!(
            outcome(&file, "/slide[1]/table[99]/tr[1]/tc[1]", &values()[0]),
            outcome(&file, "/slide[1]/table[99]/row[1]/cell[1]", &values()[0]),
        );
    }
    // (cells compared, writes published)
    eprintln!("{compared:?}");
    assert!(compared.0 >= 30 && compared.1 >= 10, "{compared:?}");
}

#[cfg(feature = "pages")]
#[test]
fn pages_direct_cells_match_the_general_route() {
    let mut compared = (0, 0);
    for path in fixtures("pages") {
        let bytes = std::fs::read(&path).expect("fixture");
        let Ok(file) = open_bytes(bytes.clone()) else {
            continue;
        };
        let Ok(tables) = file.query("tbl") else {
            continue;
        };
        let tables = tables.filter_map(Result::ok).count();
        for table in 1..=tables {
            let path = format!("/body/tbl[{table}]");
            let rows = count(&file, &path, "rows");
            let columns = count(&file, &path, "columns");
            let direct =
                |row: u64, column: u64| format!("/body/tbl[{table}]/tr[{row}]/tc[{column}]");
            let general =
                |row: u64, column: u64| format!("/body/tbl[{table}]/row[{row}]/cell[{column}]");
            let (cells, published) = compare(&bytes, &direct, &general, rows, columns);
            compared.0 += cells;
            compared.1 += published;
        }
        let file = open_bytes(bytes.clone()).expect("open");
        assert_eq!(
            outcome(&file, "/body/tbl[99]/tr[1]/tc[1]", &values()[0]),
            outcome(&file, "/body/tbl[99]/row[1]/cell[1]", &values()[0]),
        );
    }
    // (cells compared, writes published)
    eprintln!("{compared:?}");
    assert!(compared.0 >= 30 && compared.1 >= 10, "{compared:?}");
}
