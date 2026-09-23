//! Change 0744 output-identity probe.
//!
//! Replicates the perf-baseline XLSX corpus (`build_xlsx_corpus`) and its
//! update coordinates, then writes every artifact the six measured cases and
//! their verification produce, so two builds of `litchi-xlsx` can be compared
//! byte for byte: the generated archive, the first-cell view, a full stored
//! cell dump, and the no-op, one-cell and one-percent commit+save outputs with
//! a full semantic dump of each reopened output.

use std::fmt::Write as _;
use std::path::Path;

use litchi_xlsx::Workbook;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

fn sheet_name(index: usize) -> String {
    if index == 0 {
        "Sheet1".to_owned()
    } else {
        format!("Bench{index:02}")
    }
}

fn address(row: usize, column: usize) -> String {
    let mut value = column + 1;
    let mut label = String::new();
    while value != 0 {
        let remainder = (value - 1) % 26;
        label.insert(0, char::from(b'A' + u8::try_from(remainder).unwrap()));
        value = (value - 1) / 26;
    }
    format!("{label}{}", row + 1)
}

fn value(sheet: usize, row: usize, column: usize) -> i32 {
    i32::try_from(sheet * 1_000_000 + row * 1_000 + column).unwrap()
}

fn build(sheets: usize, rows: usize, columns: usize) -> Result<Vec<u8>> {
    let workbook = Workbook::new()?;
    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit.sheet("Sheet1")?.ok_or("Sheet1")?;
        for row in 0..rows {
            for column in 0..columns {
                sheet.set(address(row, column).as_str(), value(0, row, column))?;
            }
        }
    }
    for index in 1..sheets {
        let mut sheet = edit.add(sheet_name(index))?;
        for row in 0..rows {
            for column in 0..columns {
                sheet.set(address(row, column).as_str(), value(index, row, column))?;
            }
        }
    }
    Ok(edit.commit()?.workbook().to_bytes()?)
}

fn updates(sheets: usize, rows: usize, columns: usize) -> Vec<(usize, usize, usize)> {
    let total = sheets * rows * columns;
    let count = total.div_ceil(100);
    (0..count)
        .map(|index| {
            let linear = index * total / count;
            let within = linear % (rows * columns);
            (linear / (rows * columns), within / columns, within % columns)
        })
        .collect()
}

fn dump(workbook: &Workbook) -> Result<String> {
    let mut text = String::new();
    for index in 0..workbook.len() {
        let sheet = workbook.sheet(index)?.ok_or("sheet")?;
        writeln!(text, "sheet {index} {}", sheet.name())?;
        for (address, cell) in sheet.cells("A1:XFD1048576")? {
            writeln!(text, "{address:?}={cell:?}")?;
        }
    }
    Ok(text)
}

fn commit_save(archive: &[u8], changes: &[(usize, usize, usize)]) -> Result<Vec<u8>> {
    let workbook = Workbook::from_bytes(archive.to_vec())?;
    let mut edit = workbook.edit()?;
    for &(sheet, row, column) in changes {
        let name = sheet_name(sheet);
        let mut target = edit.sheet(name.as_str())?.ok_or("sheet")?;
        target.set(address(row, column).as_str(), value(sheet, row, column) + 1)?;
    }
    let commit = edit.commit()?;
    let mut output = Vec::new();
    commit.workbook().write_to(&mut output)?;
    Ok(output)
}

fn main() -> Result<()> {
    let directory = std::env::args().nth(1).ok_or("output directory")?;
    let directory = Path::new(&directory);
    std::fs::create_dir_all(directory)?;
    for (shape, sheets, rows, columns) in [("dense-wide", 2, 256, 256), ("medium", 4, 32, 32)] {
        let archive = build(sheets, rows, columns)?;
        std::fs::write(directory.join(format!("{shape}-archive.xlsx")), &archive)?;
        let workbook = Workbook::from_bytes(archive.clone())?;
        let first = workbook.sheet("Sheet1")?.ok_or("Sheet1")?;
        std::fs::write(
            directory.join(format!("{shape}-first-cell.txt")),
            format!("{:?}", first.cell("A1")?),
        )?;
        std::fs::write(directory.join(format!("{shape}-cells.txt")), dump(&workbook)?)?;
        let all = updates(sheets, rows, columns);
        for (kind, changes) in [("noop", &all[..0]), ("one-cell", &all[..1]), ("one-percent", &all[..])] {
            let output = commit_save(&archive, changes)?;
            let reopened = Workbook::from_bytes(output.clone())?;
            std::fs::write(directory.join(format!("{shape}-{kind}.xlsx")), &output)?;
            std::fs::write(
                directory.join(format!("{shape}-{kind}-cells.txt")),
                dump(&reopened)?,
            )?;
        }
    }
    Ok(())
}
