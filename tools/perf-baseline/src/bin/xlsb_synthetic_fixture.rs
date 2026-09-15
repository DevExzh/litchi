//! Generate a deterministic, larger-than-corpus synthetic XLSB workbook.
//!
//! Every `.xlsb` fixture in this repository is at most 22,715 bytes with a
//! single worksheet and 48 stored cells (change 0587, `survey/xlsb.md`), so
//! every XLSB scaling question is buried under the fixed open cost. This
//! generator writes a workbook of a caller-chosen shape through the public
//! `litchi_xlsb::writer` API so that `xlsb_crud --fixture` has something to
//! scale against. The output is a scratch fixture: it is reproducible from
//! this binary and its arguments, and is not checked in.

use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use sha2::{Digest, Sha256};
use std::env;
use std::fs;
use std::io::Cursor;
use std::path::PathBuf;

type Error = Box<dyn std::error::Error + Send + Sync>;
type Result<T> = std::result::Result<T, Error>;

const DEFAULT_SHEETS: u32 = 4;
const DEFAULT_ROWS: u32 = 500;
const DEFAULT_COLUMNS: u32 = 8;

fn usage() -> &'static str {
    "usage: xlsb_synthetic_fixture --out PATH [--sheets N] [--rows N] [--columns N]"
}

fn main() -> Result<()> {
    let mut sheets = DEFAULT_SHEETS;
    let mut rows = DEFAULT_ROWS;
    let mut columns = DEFAULT_COLUMNS;
    let mut out: Option<PathBuf> = None;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        let mut value = || -> Result<String> {
            arguments
                .next()
                .ok_or_else(|| format!("missing value for {argument}").into())
        };
        match argument.as_str() {
            "--sheets" => sheets = value()?.parse()?,
            "--rows" => rows = value()?.parse()?,
            "--columns" => columns = value()?.parse()?,
            "--out" => out = Some(PathBuf::from(value()?)),
            "--help" | "-h" => {
                println!("{}", usage());
                return Ok(());
            },
            other => return Err(format!("unknown argument {other:?}\n\n{}", usage()).into()),
        }
    }
    let out = out.ok_or_else(|| format!("--out is required\n\n{}", usage()))?;
    if sheets == 0 || rows == 0 || columns == 0 {
        return Err("--sheets, --rows and --columns must all be greater than zero".into());
    }

    let mut writer = WorkbookWriter::new();
    for sheet in 0..sheets {
        let mut worksheet = MutableWorksheet::new(format!("Sheet{}", sheet + 1));
        for row in 0..rows {
            for column in 0..columns {
                // Leave a deterministic hole in one cell of every sixteen so
                // the sheet stays sparse: the CRUD harness refuses a fixture
                // whose stored cells fill their own bounding rectangle,
                // because that would hide rectangular expansion.
                if (row.wrapping_mul(columns).wrapping_add(column)) % 16 == 3 {
                    continue;
                }
                // A deterministic finite double per coordinate, never exactly
                // 1.0 so that the CRUD harness always has a distinct
                // replacement value available.
                let value = f64::from(sheet).mul_add(
                    1_000_000.0,
                    f64::from(row).mul_add(1_000.0, f64::from(column)),
                ) + 0.5;
                worksheet.set_cell(row, column, value);
            }
        }
        writer.add_worksheet(worksheet);
    }

    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes)?;
    let bytes = bytes.into_inner();

    // Prove the fixture opens through the same eager door the harness uses
    // before it is written anywhere.
    let reopened = litchi_xlsb::Workbook::new(Cursor::new(bytes.clone()))?;
    let stored = reopened.cell_values(0)?.cells().count();

    fs::write(&out, &bytes)?;
    let mut hasher = Sha256::new();
    hasher.update(&bytes);
    let digest = hasher.finalize();
    let mut sha256 = String::with_capacity(digest.len() * 2);
    for byte in digest {
        sha256.push_str(&format!("{byte:02x}"));
    }
    println!(
        "{{\"path\":{:?},\"bytes\":{},\"sha256\":\"{sha256}\",\"sheets\":{sheets},\"rows\":{rows},\"columns\":{columns},\"stored_cells_sheet0\":{stored}}}",
        out.display().to_string(),
        bytes.len()
    );
    Ok(())
}
