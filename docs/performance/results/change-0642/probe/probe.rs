//! Change 0642 evidence probe: `SourceBackedWorksheet::visit_cells`.
//!
//! One binary, several modes. Every mode is deterministic except `bench`.
//!
//! * `corpus OUT.xlsx SHEETS ROWS COLUMNS` writes a dense probe workbook whose
//!   shape follows the harness `dense-wide` XLSX corpus (2 sheets of
//!   256 x 256 integer cells).
//! * `diff FILE...` runs the four-way differential for every worksheet of
//!   every named file: cold `visit_cells`, cold `cells`, warm-store
//!   `visit_cells`, and the eager `litchi_xlsx::Workbook` store. It prints one
//!   tab-separated line per worksheet with each leg's visited count and a
//!   digest over its `(address, cell)` sequence, or the leg's refusal text.
//! * `visit FILE SHEET REPEAT` and `cells FILE SHEET REPEAT` run one whole
//!   sheet operation REPEAT times for a callgrind isolation pair.
//! * `bench visit|cells FILE SHEET WARMUP SAMPLES` prints one elapsed
//!   nanosecond count per line, one line per sample.
//!
//! The whole-sheet area is the same `A1:XFD1048576` the harness
//! `xlsx_full_cell_scan` selector uses.

use std::error::Error;
use std::fmt::Write as _;
use std::hint::black_box;
use std::path::Path;
use std::time::Instant;

type BoxError = Box<dyn Error>;

/// The whole-sheet area used by every mode.
const WHOLE_SHEET: &str = "A1:XFD1048576";

/// FNV-1a over the canonical rendering of a visited sequence.
struct Digest(u64);

impl Digest {
    fn new() -> Self {
        Self(0xcbf2_9ce4_8422_2325)
    }

    fn write(&mut self, bytes: &[u8]) {
        for byte in bytes {
            self.0 ^= u64::from(*byte);
            self.0 = self.0.wrapping_mul(0x0000_0100_0000_01b3);
        }
    }

    fn cell(&mut self, address: litchi_xlsx::Address, cell: &litchi_xlsx::cell::Cell) {
        let mut text = String::new();
        // `Cell` and `Value` derive `Debug`, so this rendering is exact for
        // every variant, including the formula cache and the empty cell.
        let _ = write!(&mut text, "{address}={cell:?}\n");
        self.write(text.as_bytes());
    }

    fn finish(&self) -> String {
        format!("{:016x}", self.0)
    }
}

/// One leg's outcome: a visited count and digest, or a refusal.
enum Leg {
    Visited(usize, String),
    Refused(String),
}

impl Leg {
    fn render(&self) -> String {
        match self {
            Self::Visited(count, digest) => format!("ok:{count}:{digest}"),
            Self::Refused(text) => format!("refused:{text}"),
        }
    }
}

fn source_visit(path: &Path, sheet: &str, warm_store: bool) -> Leg {
    let run = || -> Result<(usize, String), litchi_xlsx::Error> {
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
        let worksheet = match workbook.sheet(sheet)? {
            Some(worksheet) => worksheet,
            None => return Ok((0, "missing-sheet".to_owned())),
        };
        if warm_store {
            // Publish the materialized store first, so `visit_cells` takes the
            // stored route instead of the bounded selected scan.
            let _extent = worksheet.stored_extent()?;
        }
        let mut digest = Digest::new();
        let visited = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
            digest.cell(address, cell);
            Ok(())
        })?;
        Ok((visited, digest.finish()))
    };
    match run() {
        Ok((count, digest)) => Leg::Visited(count, digest),
        Err(error) => Leg::Refused(error.to_string()),
    }
}

fn source_cells(path: &Path, sheet: &str) -> Leg {
    let run = || -> Result<(usize, String), litchi_xlsx::Error> {
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
        let worksheet = match workbook.sheet(sheet)? {
            Some(worksheet) => worksheet,
            None => return Ok((0, "missing-sheet".to_owned())),
        };
        let values = worksheet.cells(WHOLE_SHEET)?;
        let mut digest = Digest::new();
        for value in &values {
            digest.cell(value.address, &value.cell);
        }
        Ok((values.len(), digest.finish()))
    };
    match run() {
        Ok((count, digest)) => Leg::Visited(count, digest),
        Err(error) => Leg::Refused(error.to_string()),
    }
}

fn eager_cells(path: &Path, sheet: &str) -> Leg {
    let run = || -> Result<(usize, String), litchi_xlsx::Error> {
        let workbook = litchi_xlsx::Workbook::open(path)?;
        let worksheet = match workbook.sheet(sheet)? {
            Some(worksheet) => worksheet,
            None => return Ok((0, "missing-sheet".to_owned())),
        };
        let mut digest = Digest::new();
        let mut count = 0usize;
        for (address, cell) in worksheet.cells(WHOLE_SHEET)? {
            digest.cell(address, cell);
            count += 1;
        }
        Ok((count, digest.finish()))
    };
    match run() {
        Ok((count, digest)) => Leg::Visited(count, digest),
        Err(error) => Leg::Refused(error.to_string()),
    }
}

fn sheet_names(path: &Path) -> Result<Vec<String>, litchi_xlsx::Error> {
    let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
    Ok(workbook
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect())
}

fn differential(paths: &[String]) -> Result<(), BoxError> {
    println!("file\tsheet\tcold_visit\tcold_cells\twarm_visit\teager_cells\tagree");
    let mut disagreements = 0usize;
    let mut worksheets = 0usize;
    for path in paths {
        let path = Path::new(path);
        let names = match sheet_names(path) {
            Ok(names) => names,
            Err(error) => {
                println!(
                    "{}\t-\topen-refused:{error}\t-\t-\t-\tskipped",
                    path.display()
                );
                continue;
            },
        };
        for name in names {
            worksheets += 1;
            let cold_visit = source_visit(path, &name, false);
            let cold_cells = source_cells(path, &name);
            let warm_visit = source_visit(path, &name, true);
            let eager = eager_cells(path, &name);
            let rendered = [
                cold_visit.render(),
                cold_cells.render(),
                warm_visit.render(),
                eager.render(),
            ];
            let agree = rendered.iter().all(|value| *value == rendered[0]);
            if !agree {
                disagreements += 1;
            }
            println!(
                "{}\t{}\t{}\t{}\t{}\t{}\t{}",
                path.display(),
                name,
                rendered[0],
                rendered[1],
                rendered[2],
                rendered[3],
                if agree { "yes" } else { "NO" },
            );
        }
    }
    println!("# worksheets={worksheets} disagreements={disagreements}");
    if disagreements != 0 {
        return Err("differential disagreement".into());
    }
    Ok(())
}

/// Walk a worksheet whose store is already materialized.
///
/// The open and the store publication are outside the returned closure, so a
/// caller can time only the walk.
fn warm_worksheet(
    path: &Path,
    sheet: &str,
) -> Result<
    (
        litchi_xlsx::SourceBackedWorkbook,
        litchi_xlsx::SourceWorksheet,
    ),
    BoxError,
> {
    let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
    let worksheet = workbook
        .sheet(sheet)?
        .ok_or("probe worksheet is missing from this workbook")?;
    let _extent = worksheet.stored_extent()?;
    Ok((workbook, worksheet))
}

fn warm_visit(worksheet: &litchi_xlsx::SourceWorksheet) -> Result<usize, BoxError> {
    let mut sum = 0u64;
    let visited = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
        sum = sum.wrapping_add(u64::from(address.row().get()));
        black_box(cell);
        Ok(())
    })?;
    black_box(sum);
    Ok(visited)
}

fn warm_cells(worksheet: &litchi_xlsx::SourceWorksheet) -> Result<usize, BoxError> {
    let values = worksheet.cells(WHOLE_SHEET)?;
    let count = values.len();
    black_box(&values);
    Ok(count)
}

fn whole_sheet_visit(path: &Path, sheet: &str) -> Result<usize, BoxError> {
    let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
    let worksheet = workbook
        .sheet(sheet)?
        .ok_or("probe worksheet is missing from this workbook")?;
    let mut sum = 0u64;
    let visited = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
        sum = sum.wrapping_add(u64::from(address.row().get()));
        black_box(cell);
        Ok(())
    })?;
    black_box(sum);
    Ok(visited)
}

fn whole_sheet_cells(path: &Path, sheet: &str) -> Result<usize, BoxError> {
    let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
    let worksheet = workbook
        .sheet(sheet)?
        .ok_or("probe worksheet is missing from this workbook")?;
    let values = worksheet.cells(WHOLE_SHEET)?;
    let count = values.len();
    black_box(&values);
    Ok(count)
}

fn build_corpus(out: &Path, sheets: usize, rows: u32, columns: u32) -> Result<(), BoxError> {
    // `WorksheetEdit` and `NewSheet` are distinct handles that share `set`.
    macro_rules! fill {
        ($sheet:expr, $base:expr) => {{
            let sheet = $sheet;
            for row in 1..=rows {
                for column in 1..=columns {
                    let address = litchi_xlsx::Address::at(row, column)?;
                    let value = $base + i32::try_from(row)? * 1_000 + i32::try_from(column)?;
                    sheet.set(address, value)?;
                }
            }
        }};
    }

    let workbook = litchi_xlsx::Workbook::new()?;
    let mut edit = workbook.edit()?;
    for index in 0..sheets {
        let base = i32::try_from(index)? * 1_000_000;
        if index == 0 {
            let mut sheet = edit
                .sheet("Sheet1")?
                .ok_or("probe corpus worksheet is missing")?;
            fill!(&mut sheet, base);
        } else {
            let name = format!("Sheet{}", index + 1);
            let mut sheet = edit.add(name.as_str())?;
            fill!(&mut sheet, base);
        }
    }
    let commit = edit.commit()?;
    std::fs::write(out, commit.workbook().to_bytes()?)?;
    Ok(())
}

fn main() -> Result<(), BoxError> {
    let arguments: Vec<String> = std::env::args().skip(1).collect();
    let mode = arguments.first().map(String::as_str).unwrap_or("");
    match mode {
        "corpus" => build_corpus(
            Path::new(&arguments[1]),
            arguments[2].parse()?,
            arguments[3].parse()?,
            arguments[4].parse()?,
        )?,
        "diff" => differential(&arguments[1..])?,
        "visit" | "cells" => {
            let path = Path::new(&arguments[1]);
            let sheet = &arguments[2];
            let repeat: usize = arguments[3].parse()?;
            let mut total = 0usize;
            for _ in 0..repeat {
                total += if mode == "visit" {
                    whole_sheet_visit(path, sheet)?
                } else {
                    whole_sheet_cells(path, sheet)?
                };
            }
            println!("{mode}\t{}\t{sheet}\t{repeat}\t{total}", path.display());
        },
        "visit-warm" | "cells-warm" => {
            let path = Path::new(&arguments[1]);
            let sheet = &arguments[2];
            let repeat: usize = arguments[3].parse()?;
            let (workbook, worksheet) = warm_worksheet(path, sheet)?;
            let mut total = 0usize;
            for _ in 0..repeat {
                total += if mode == "visit-warm" {
                    warm_visit(&worksheet)?
                } else {
                    warm_cells(&worksheet)?
                };
            }
            black_box(&workbook);
            println!("{mode}\t{}\t{sheet}\t{repeat}\t{total}", path.display());
        },
        "bench" => {
            let operation = arguments[1].as_str();
            let path = Path::new(&arguments[2]);
            let sheet = &arguments[3];
            let warmup: usize = arguments[4].parse()?;
            let samples: usize = arguments[5].parse()?;
            // The warm legs publish the store once, outside every sample.
            let warm = match operation {
                "visit-warm" | "cells-warm" => Some(warm_worksheet(path, sheet)?),
                _ => None,
            };
            let mut count = 0usize;
            for iteration in 0..(warmup + samples) {
                let started = Instant::now();
                let visited = match operation {
                    "visit" => whole_sheet_visit(path, sheet)?,
                    "cells" => whole_sheet_cells(path, sheet)?,
                    "visit-warm" => warm_visit(&warm.as_ref().expect("warm worksheet").1)?,
                    "cells-warm" => warm_cells(&warm.as_ref().expect("warm worksheet").1)?,
                    other => return Err(format!("unknown bench operation {other}").into()),
                };
                let elapsed = started.elapsed();
                count = visited;
                if iteration >= warmup {
                    println!("{}", elapsed.as_nanos());
                }
            }
            black_box(&warm);
            eprintln!("visited={count}");
        },
        other => return Err(format!("unknown mode {other:?}").into()),
    }
    Ok(())
}
