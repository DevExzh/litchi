//! Reproducible measurement support for change 0683.
//!
//! The probe deliberately measures the public source-backed worksheet
//! boundary.  It does not reach into `raw::selected_worksheet`, so a future
//! `SelectedRecord` layout change is priced through the owning operation that
//! retains and resolves those records.

use std::error::Error;
use std::fmt::Write as _;
use std::hint::black_box;
use std::path::Path;
use std::time::Instant;

use litchi_xlsx::{Cell, Workbook};

pub type BoxError = Box<dyn Error>;

/// The whole-sheet area keeps selection semantics identical across cases.
pub const WHOLE_SHEET: &str = "A1:XFD1048576";

/// FNV-1a over the canonical rendering of an observed sequence.
#[derive(Debug, Clone, Copy)]
pub struct Digest(u64);

impl Digest {
    #[must_use]
    pub const fn new() -> Self {
        Self(0xcbf2_9ce4_8422_2325)
    }

    pub fn write(&mut self, bytes: &[u8]) {
        for byte in bytes {
            self.0 ^= u64::from(*byte);
            self.0 = self.0.wrapping_mul(0x0000_0100_0000_01b3);
        }
    }

    pub fn cell(&mut self, address: litchi_xlsx::Address, cell: &Cell) {
        let mut text = String::new();
        let _ = write!(&mut text, "{address}={cell:?}\n");
        self.write(text.as_bytes());
    }

    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    #[must_use]
    pub fn finish(self) -> String {
        format!("{:016x}", self.0)
    }
}

impl Default for Digest {
    fn default() -> Self {
        Self::new()
    }
}

/// Semantic result of one route. Refusals are retained so differential output
/// can prove that both legs made the same fallback decision.
#[derive(Debug, Clone)]
pub enum Observation {
    Ok { callbacks: usize, digest: String },
    Refused { error: String, callbacks: usize },
}

/// A source-backed workbook and worksheet whose setup is complete. Keeping
/// this owner alive lets the allocator companion place setup before its
/// measurement reset and sample the retained gauge before dropping it.
pub struct PreparedWorksheet {
    workbook: litchi_xlsx::SourceBackedWorkbook,
    worksheet: litchi_xlsx::SourceWorksheet,
}

/// One measured operation's count and optional returned vector. The allocator
/// companion keeps `retained_cells` alive through its gauge reads so retained
/// bytes describe the public `cells` result when that route returns one.
pub struct OperationResult {
    pub callbacks: usize,
    pub retained_cells: Option<Vec<litchi_xlsx::SourceCell>>,
}

impl PreparedWorksheet {
    pub fn open(path: &Path, sheet: &str, materialize: bool) -> Result<Self, BoxError> {
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
        let worksheet = workbook
            .sheet(sheet)?
            .ok_or("probe worksheet is missing from this workbook")?;
        if materialize {
            let _extent = worksheet.stored_extent()?;
        }
        Ok(Self {
            workbook,
            worksheet,
        })
    }

    pub fn count(&self, operation: &str) -> Result<usize, BoxError> {
        let count = self.measure(operation)?.callbacks;
        black_box(&self.workbook);
        Ok(count)
    }

    pub fn measure(&self, operation: &str) -> Result<OperationResult, BoxError> {
        match operation {
            "visit-selected" | "visit-warm" => Ok(OperationResult {
                callbacks: visit_count_on_worksheet(&self.worksheet)?,
                retained_cells: None,
            }),
            "cells-selected" | "cells-warm" => {
                let values = self.worksheet.cells(WHOLE_SHEET)?;
                Ok(OperationResult {
                    callbacks: values.len(),
                    retained_cells: Some(values),
                })
            },
            other => Err(format!("unknown prepared operation {other}").into()),
        }
    }
}

pub fn cold_count(operation: &str, path: &Path, sheet: &str) -> Result<usize, BoxError> {
    Ok(cold_measure(operation, path, sheet)?.callbacks)
}

pub fn cold_measure(
    operation: &str,
    path: &Path,
    sheet: &str,
) -> Result<OperationResult, BoxError> {
    let prepared = PreparedWorksheet::open(path, sheet, false)?;
    let count = match operation {
        "visit-cold" => OperationResult {
            callbacks: visit_count_on_worksheet(&prepared.worksheet)?,
            retained_cells: None,
        },
        "cells-cold" => {
            let values = prepared.worksheet.cells(WHOLE_SHEET)?;
            OperationResult {
                callbacks: values.len(),
                retained_cells: Some(values),
            }
        },
        other => return Err(format!("unknown cold operation {other}").into()),
    };
    Ok(count)
}

impl Observation {
    #[must_use]
    pub fn render(&self) -> String {
        match self {
            Self::Ok { callbacks, digest } => format!("ok:{callbacks}:{digest}"),
            Self::Refused { error, callbacks } => format!("refused:{callbacks}:{error}"),
        }
    }

    #[must_use]
    pub const fn callbacks(&self) -> usize {
        match self {
            Self::Ok { callbacks, .. } => *callbacks,
            Self::Refused { callbacks, .. } => *callbacks,
        }
    }
}

fn source_visit(path: &Path, sheet: &str) -> Observation {
    let run = || -> Result<Observation, litchi_xlsx::Error> {
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
        let worksheet = match workbook.sheet(sheet)? {
            Some(worksheet) => worksheet,
            None => {
                return Ok(Observation::Ok {
                    callbacks: 0,
                    digest: "missing-sheet".to_owned(),
                });
            },
        };
        let mut digest = Digest::new();
        let mut callback_count = 0usize;
        let result = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
            callback_count = callback_count.saturating_add(1);
            digest.cell(address, cell);
            Ok(())
        });
        Ok(match result {
            Ok(callbacks) => Observation::Ok {
                callbacks,
                digest: digest.finish(),
            },
            Err(error) => Observation::Refused {
                error: error.to_string(),
                callbacks: callback_count,
            },
        })
    };
    match run() {
        Ok(observation) => observation,
        Err(error) => Observation::Refused {
            error: error.to_string(),
            callbacks: 0,
        },
    }
}

fn source_cells(path: &Path, sheet: &str) -> Observation {
    let run = || -> Result<Observation, litchi_xlsx::Error> {
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_path(path)?;
        let worksheet = match workbook.sheet(sheet)? {
            Some(worksheet) => worksheet,
            None => {
                return Ok(Observation::Ok {
                    callbacks: 0,
                    digest: "missing-sheet".to_owned(),
                });
            },
        };
        let values = worksheet.cells(WHOLE_SHEET)?;
        let mut digest = Digest::new();
        for value in &values {
            digest.cell(value.address, &value.cell);
        }
        Ok(Observation::Ok {
            callbacks: values.len(),
            digest: digest.finish(),
        })
    };
    match run() {
        Ok(observation) => observation,
        Err(error) => Observation::Refused {
            error: error.to_string(),
            callbacks: 0,
        },
    }
}

fn source_visit_warm(path: &Path, sheet: &str) -> Result<Observation, BoxError> {
    let prepared = PreparedWorksheet::open(path, sheet, true)?;
    Ok(source_visit_on_worksheet(&prepared.worksheet))
}

fn source_cells_warm(path: &Path, sheet: &str) -> Result<Observation, BoxError> {
    let prepared = PreparedWorksheet::open(path, sheet, true)?;
    Ok(source_cells_on_worksheet(&prepared.worksheet))
}

fn source_visit_on_worksheet(worksheet: &litchi_xlsx::SourceWorksheet) -> Observation {
    let mut digest = Digest::new();
    let mut callback_count = 0usize;
    let result = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
        callback_count = callback_count.saturating_add(1);
        digest.cell(address, cell);
        Ok(())
    });
    match result {
        Ok(callbacks) => Observation::Ok {
            callbacks,
            digest: digest.finish(),
        },
        Err(error) => Observation::Refused {
            error: error.to_string(),
            callbacks: callback_count,
        },
    }
}

fn source_cells_on_worksheet(worksheet: &litchi_xlsx::SourceWorksheet) -> Observation {
    let values = match worksheet.cells(WHOLE_SHEET) {
        Ok(values) => values,
        Err(error) => {
            return Observation::Refused {
                error: error.to_string(),
                callbacks: 0,
            };
        },
    };
    let mut digest = Digest::new();
    for value in &values {
        digest.cell(value.address, &value.cell);
    }
    Observation::Ok {
        callbacks: values.len(),
        digest: digest.finish(),
    }
}

// The callback-counting variants below deliberately do not render a digest;
// they are used by timed and allocator runs.
#[derive(Debug)]
struct VisitFailure {
    error: litchi_xlsx::Error,
    callbacks: usize,
}
impl std::fmt::Display for VisitFailure {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.error.fmt(f)
    }
}
impl Error for VisitFailure {}
pub fn error_callbacks(error: &(dyn Error + 'static)) -> usize {
    error
        .downcast_ref::<VisitFailure>()
        .map_or(0, |failure| failure.callbacks)
}

fn visit_count_on_worksheet(worksheet: &litchi_xlsx::SourceWorksheet) -> Result<usize, BoxError> {
    let mut observed = 0usize;
    let mut checksum = 0u64;
    let result = worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
        observed += 1;
        checksum = checksum
            .wrapping_add(u64::from(address.row().get()))
            .wrapping_add(u64::from(address.column().get()));
        black_box(cell);
        Ok(())
    });
    black_box(checksum);
    result.map_err(|error| {
        Box::new(VisitFailure {
            error,
            callbacks: observed,
        }) as BoxError
    })
}

fn eager_cells(path: &Path, sheet: &str) -> Observation {
    let run = || -> Result<Observation, litchi_xlsx::Error> {
        let workbook = Workbook::open(path)?;
        let worksheet = match workbook.sheet(sheet)? {
            Some(worksheet) => worksheet,
            None => {
                return Ok(Observation::Ok {
                    callbacks: 0,
                    digest: "missing-sheet".to_owned(),
                });
            },
        };
        let values = worksheet.cells(WHOLE_SHEET)?;
        let mut digest = Digest::new();
        let mut callbacks = 0usize;
        for (address, cell) in values {
            digest.cell(address, cell);
            callbacks = callbacks.saturating_add(1);
        }
        Ok(Observation::Ok {
            callbacks,
            digest: digest.finish(),
        })
    };
    match run() {
        Ok(observation) => observation,
        Err(error) => Observation::Refused {
            error: error.to_string(),
            callbacks: 0,
        },
    }
}

/// Run all semantic controls for one generated or real workbook.
pub fn differential(path: &Path, sheet: &str) -> Result<(), BoxError> {
    let cold_visit = source_visit(path, sheet);
    let cold_cells = source_cells(path, sheet);
    let warm_visit = match source_visit_warm(path, sheet) {
        Ok(observation) => observation,
        Err(error) => Observation::Refused {
            error: error.to_string(),
            callbacks: 0,
        },
    };
    let warm_cells = match source_cells_warm(path, sheet) {
        Ok(observation) => observation,
        Err(error) => Observation::Refused {
            error: error.to_string(),
            callbacks: 0,
        },
    };
    let eager = eager_cells(path, sheet);
    let rendered = [
        cold_visit.render(),
        cold_cells.render(),
        warm_visit.render(),
        warm_cells.render(),
        eager.render(),
    ];
    let agree = rendered.iter().all(|value| *value == rendered[0]);
    println!(
        "{}\t{}\tcold_visit={}\tcold_cells={}\twarm_visit={}\twarm_cells={}\teager_cells={}\tagree={}",
        path.display(),
        sheet,
        rendered[0],
        rendered[1],
        rendered[2],
        rendered[3],
        rendered[4],
        if agree { "yes" } else { "NO" },
    );
    if !agree {
        return Err("differential disagreement".into());
    }
    Ok(())
}

/// One timed operation and its semantic observation.
#[derive(Debug, Clone)]
pub struct TimingSample {
    pub nanos: u128,
    pub callbacks: usize,
    pub succeeded: bool,
}

/// Time a route with package setup explicitly outside the interval for warm
/// and selected modes. Cold modes include package opening in every sample.
pub fn benchmark(
    operation: &str,
    path: &Path,
    sheet: &str,
    warmup: usize,
    samples: usize,
) -> Result<Vec<TimingSample>, BoxError> {
    let selected = matches!(operation, "visit-selected" | "cells-selected");
    let warm = matches!(operation, "visit-warm" | "cells-warm");
    let cold = matches!(operation, "visit-cold" | "cells-cold");
    if !(selected || warm || cold) {
        return Err(format!("unknown benchmark operation {operation:?}").into());
    }

    let setup = if selected || warm {
        Some(PreparedWorksheet::open(path, sheet, warm)?)
    } else {
        None
    };

    let mut results = Vec::new();
    for iteration in 0..warmup.saturating_add(samples) {
        let started = Instant::now();
        let (callbacks, succeeded) = if let Some(prepared) = setup.as_ref() {
            let result = prepared.count(operation);
            match result {
                Ok(callbacks) => (callbacks, true),
                Err(error) => {
                    let callbacks = error_callbacks(error.as_ref());
                    black_box(error);
                    (callbacks, false)
                },
            }
        } else {
            let result = cold_count(operation, path, sheet);
            match result {
                Ok(callbacks) => (callbacks, true),
                Err(error) => {
                    let callbacks = error_callbacks(error.as_ref());
                    black_box(error);
                    (callbacks, false)
                },
            }
        };
        let nanos = started.elapsed().as_nanos();
        if iteration >= warmup {
            results.push(TimingSample {
                nanos,
                callbacks,
                succeeded,
            });
        }
    }
    black_box(&setup);
    Ok(results)
}
