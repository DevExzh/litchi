//! Record 0267: a query that can stop early (`first()`, `limit(n)`, or
//! alternatives that each select their `[N]`th match) reads a cold
//! worksheet through the leaf's cell cursor, which decodes only the leading
//! cells. These tests hold the streamed walks to the full walk: on generated
//! workbooks large enough that the cursors stream (XLSX prefixes, XLSB
//! batches, XLS resumed scans), every early-stopping query returns the full
//! query's leading matches, from memory and from a file, on cold and warm
//! handles, with each match's kind, path, text and properties.

#![cfg(all(feature = "xlsx", feature = "xlsb", feature = "xls"))]

use std::fmt::Write as _;

use litchi_ops::File;

#[allow(dead_code, reason = "shared test fixture")]
#[path = "support/xlsx_workbook.rs"]
mod xlsx_workbook;

const ROWS: usize = 3_000;

/// Selectors whose matches start early, late, or nowhere.
const SELECTORS: &[&str] = &[
    "cell",
    "cell[value>1000]",
    "cell[value<500]",
    "cell:contains(clo)",
    "cell[type=Number]",
    "cell[value=999999]",
    "B",
    "A1:C20",
    "cell[5]",
    "cell[2500]",
    "cell[last()]",
    "sheet cell[value>1000]",
    "cell[bold=true]",
];

/// The XLSB workbook of the XLSX fixture's shape: a header and `rows` rows
/// of region, sales, status.
fn xlsb_workbook(rows: usize) -> Vec<u8> {
    let mut sheet = litchi_xlsb::builder::Worksheet::new("Data");
    for (column, header) in ["Region", "Sales", "Status"].into_iter().enumerate() {
        sheet.set_cell(0, column as u32, header.to_string());
    }
    for index in 0..rows {
        let row = index as u32 + 1;
        sheet.set_cell(
            row,
            0,
            ["West", "East", "North", "South"][index % 4].to_string(),
        );
        sheet.set_cell(row, 1, [1200.0, 800.0, 450.5, 2000.0][index % 4]);
        sheet.set_cell(row, 2, ["open", "closed"][index % 2].to_string());
    }
    let mut builder = litchi_xlsb::Builder::new();
    builder.add_worksheet(sheet);
    builder.build().unwrap().to_bytes().unwrap()
}

/// The XLS workbook of the same shape.
fn xls_workbook(rows: usize) -> Vec<u8> {
    let mut builder = litchi_xls::writer::Builder::new();
    let sheet = builder.add_worksheet("Data").unwrap();
    for (column, header) in ["Region", "Sales", "Status"].into_iter().enumerate() {
        builder
            .write_string(sheet, 0, column as u16, header)
            .unwrap();
    }
    for index in 0..rows {
        let row = index as u32 + 1;
        builder
            .write_string(sheet, row, 0, ["West", "East", "North", "South"][index % 4])
            .unwrap();
        builder
            .write_number(sheet, row, 1, [1200.0, 800.0, 450.5, 2000.0][index % 4])
            .unwrap();
        builder
            .write_string(sheet, row, 2, ["open", "closed"][index % 2])
            .unwrap();
    }
    builder.build().unwrap().to_bytes().unwrap()
}

/// Every match's kind, path, text and properties, or the error.
fn describe<'f>(
    matches: impl Iterator<Item = litchi_ops::Result<litchi_ops::node::Node<'f>>>,
) -> Vec<String> {
    matches
        .map(|node| match node {
            Ok(node) => {
                let mut line = format!(
                    "{} {} {:?}",
                    node.kind(),
                    node.path(),
                    node.view().text().ok().flatten()
                );
                if let Ok(format) = node.view().format() {
                    let _ = write!(line, " {format:?}");
                }
                line
            },
            Err(error) => format!("failed {:?}", error.code()),
        })
        .collect()
}

/// The handles a check opens: from memory, and from a file.
fn handles(bytes: &[u8], name: &str, directory: &std::path::Path) -> Vec<File> {
    let path = directory.join(name);
    std::fs::write(&path, bytes).unwrap();
    vec![
        litchi_ops::open_bytes(bytes.to_vec()).unwrap(),
        litchi_ops::open(&path).unwrap(),
    ]
}

fn check(bytes: &[u8], name: &str) {
    let directory = tempfile::tempdir().unwrap();
    for selector in SELECTORS {
        // The reference: a full walk on its own handle.
        let full = {
            let file = &handles(bytes, name, directory.path())[0];
            describe(file.query(*selector).unwrap())
        };
        for limit in [1, 3, 200, 5_000] {
            for file in handles(bytes, name, directory.path()) {
                // Cold, then warm (a retained prefix or batch, a kept
                // validation), on one handle.
                for _ in 0..2 {
                    let early = describe(file.query(*selector).unwrap().limit(limit));
                    let expected: Vec<String> = full.iter().take(limit).cloned().collect();
                    assert_eq!(early, expected, "{name} {selector} limit {limit}");
                }
            }
        }
        for file in handles(bytes, name, directory.path()) {
            let first = file
                .query(*selector)
                .unwrap()
                .first()
                .map(|node| describe(node.into_iter().map(Ok)));
            match first {
                Ok(first) => assert_eq!(first, full.iter().take(1).cloned().collect::<Vec<_>>()),
                Err(error) => assert_eq!(
                    vec![format!("failed {:?}", error.code())],
                    full.iter().take(1).cloned().collect::<Vec<_>>()
                ),
            }
        }
    }
}

#[test]
fn xlsx_early_stopping_queries_return_the_full_querys_leading_matches() {
    check(&xlsx_workbook::workbook(ROWS), "early.xlsx");
}

#[test]
fn xlsb_early_stopping_queries_return_the_full_querys_leading_matches() {
    check(&xlsb_workbook(ROWS), "early.xlsb");
}

#[test]
fn xls_early_stopping_queries_return_the_full_querys_leading_matches() {
    check(&xls_workbook(ROWS), "early.xls");
}
