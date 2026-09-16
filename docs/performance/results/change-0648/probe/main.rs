//! Scratch attribution probe for change 0636.
//!
//! Counts logical positional reads, read bytes and source-version observations
//! taken by the public source-backed `litchi_xls` APIs over one in-process
//! `ReadAt`, and freezes each operation's outcome as a digest so a before/after
//! pair is also a differential. Modes:
//!
//!   cursor-probe detail  FILE...   one block per fixture with size histograms
//!   cursor-probe sweep   FILE...   one TSV line per fixture and operation
//!   cursor-probe time OP N FILE    N timed iterations of one operation
use litchi_core::{ReadAt, SourceVersion};
use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;
use std::io;
use std::sync::{Arc, Mutex};

thread_local! {
    static ALLOCATIONS: Cell<usize> = const { Cell::new(0) };
    static ALLOCATED_BYTES: Cell<usize> = const { Cell::new(0) };
}

fn record_alloc(bytes: usize) {
    let _ = ALLOCATIONS.try_with(|cell| cell.set(cell.get().wrapping_add(1)));
    let _ = ALLOCATED_BYTES.try_with(|cell| cell.set(cell.get().wrapping_add(bytes)));
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        record_alloc(layout.size());
        unsafe { System.alloc(layout) }
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        record_alloc(layout.size());
        unsafe { System.alloc_zeroed(layout) }
    }
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        record_alloc(new_size.saturating_sub(layout.size()));
        unsafe { System.realloc(pointer, layout, new_size) }
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
}

struct Trace {
    bytes: Arc<[u8]>,
    log: Mutex<Vec<(u64, usize)>>,
    versions: Mutex<u64>,
    lens: Mutex<u64>,
}

impl ReadAt for Trace {
    fn len(&self) -> io::Result<u64> {
        *self.lens.lock().unwrap() += 1;
        Ok(self.bytes.len() as u64)
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.log.lock().unwrap().push((offset, output.len()));
        let start = usize::try_from(offset)
            .unwrap_or(usize::MAX)
            .min(self.bytes.len());
        let end = start.saturating_add(output.len()).min(self.bytes.len());
        output[..end - start].copy_from_slice(&self.bytes[start..end]);
        Ok(end - start)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        *self.versions.lock().unwrap() += 1;
        Ok(SourceVersion::new(self.bytes.len() as u64, 7))
    }
}

/// A plain owned source with no counters, for the timing mode.
struct Plain(Arc<[u8]>);
impl ReadAt for Plain {
    fn len(&self) -> io::Result<u64> {
        Ok(self.0.len() as u64)
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX).min(self.0.len());
        let end = start.saturating_add(output.len()).min(self.0.len());
        output[..end - start].copy_from_slice(&self.0[start..end]);
        Ok(end - start)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(self.0.len() as u64, 7))
    }
}

fn fnv(items: &[String]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for item in items {
        for byte in item.as_bytes() {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(0x1_0000_0001_b3);
        }
        hash ^= 0xff;
        hash = hash.wrapping_mul(0x1_0000_0001_b3);
    }
    hash
}

struct Counts {
    reads: usize,
    bytes: u64,
    small: usize,
    versions: u64,
    lens: u64,
    outcome: String,
}

fn measure(bytes: &Arc<[u8]>, op: &str) -> Counts {
    let trace = Arc::new(Trace {
        bytes: Arc::clone(bytes),
        log: Mutex::new(Vec::new()),
        versions: Mutex::new(0),
        lens: Mutex::new(0),
    });
    let source = Arc::clone(&trace) as Arc<dyn ReadAt>;
    let mut from = 0usize;
    let mut version_base = 0u64;
    let outcome = match op {
        "validate" => match litchi_xls::validation::validate_source(source) {
            Ok(report) => format!("ok:{:016x}", fnv(&[format!("{report:?}")])),
            Err(error) => format!("err:{error}"),
        },
        _ => {
            let workbook = match litchi_xls::SourceBackedWorkbook::from_read_at(source) {
                Ok(workbook) => workbook,
                Err(error) => {
                    return finish(&trace, 0, 0, format!("open-err:{error}"));
                },
            };
            match op {
                "open" => "ok".to_owned(),
                "full-text" => {
                    from = trace.log.lock().unwrap().len();
                    version_base = *trace.versions.lock().unwrap();
                    match workbook.text() {
                        Ok(text) => format!("text:{}:{:016x}", text.len(), fnv(&[text])),
                        Err(error) => format!("refused:{error}"),
                    }
                },
                _ => {
                    let names = match workbook.worksheet_names() {
                        Ok(names) => names,
                        Err(error) => return finish(&trace, 0, 0, format!("names-err:{error}")),
                    };
                    if op == "list" {
                        format!("names:{}:{:016x}", names.len(), fnv(&names))
                    } else {
                        let mut chosen = None;
                        for index in 0..names.len() {
                            if let Ok(Some(sheet)) = workbook.worksheet(index) {
                                let mut seen = 0usize;
                                if sheet
                                    .visit_cells(|_cell| {
                                        seen += 1;
                                        Ok(())
                                    })
                                    .is_ok()
                                    && seen > 0
                                {
                                    chosen = Some(index);
                                    break;
                                }
                            }
                        }
                        let Some(chosen) = chosen else {
                            return finish(&trace, 0, 0, "no-cells".to_owned());
                        };
                        from = trace.log.lock().unwrap().len();
                        version_base = *trace.versions.lock().unwrap();
                        let sheet = workbook.worksheet(chosen).unwrap().unwrap();
                        let mut cells = Vec::new();
                        match sheet.visit_cells(|cell| {
                            cells.push(format!("{cell:?}"));
                            Ok(())
                        }) {
                            Ok(()) => {
                                format!("sheet{chosen}:cells:{}:{:016x}", cells.len(), fnv(&cells))
                            },
                            Err(error) => format!("refused:{error}"),
                        }
                    }
                },
            }
        },
    };
    finish(&trace, from, version_base, outcome)
}

fn finish(trace: &Trace, from: usize, version_base: u64, outcome: String) -> Counts {
    let log = trace.log.lock().unwrap();
    let segment = &log[from.min(log.len())..];
    Counts {
        reads: segment.len(),
        bytes: segment.iter().map(|(_, length)| *length as u64).sum(),
        small: segment.iter().filter(|(_, length)| *length <= 512).count(),
        versions: trace.versions.lock().unwrap().saturating_sub(version_base),
        lens: *trace.lens.lock().unwrap(),
        outcome,
    }
}

const OPS: [&str; 6] = ["validate", "open", "list", "all-cells", "full-text", "one-cell"];

fn main() {
    let mut args = std::env::args().skip(1);
    let mode = args.next().unwrap_or_else(|| "sweep".to_owned());
    if mode == "sst" {
        // The source extent of each fixture's SST record group, walked out of
        // the Workbook stream by hand: the payload of the `SST` record (0x00FC)
        // through the last payload of the `Continue` run (0x003C) behind it,
        // the four-byte Continue headers between them included.
        for path in args {
            let Ok(bytes) = std::fs::read(&path) else { continue };
            let Ok(file) = litchi_cfb::SharedOleFile::open(Arc::new(
                litchi_core::OwnedSource::new(bytes),
            )) else { continue };
            let stream = match file.open_stream(&["Workbook"]) {
                Ok(stream) => stream,
                Err(_) => match file.open_stream(&["Book"]) {
                    Ok(stream) => stream,
                    Err(_) => continue,
                },
            };
            let mut offset = 0usize;
            let mut start = None;
            let mut end = 0usize;
            while offset + 4 <= stream.len() {
                let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
                let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
                if kind == 0x00FC && start.is_none() {
                    start = Some(offset + 4);
                    end = offset + 4 + length;
                } else if kind == 0x003C && start.is_some() && offset == end {
                    end = offset + 4 + length;
                } else if start.is_some() {
                    break;
                }
                offset += 4 + length;
            }
            match start {
                Some(start) => println!("{path}\t{start}\t{end}\t{}", end - start),
                None => println!("{path}\t-\t-\t0"),
            }
        }
        return;
    }
    if mode == "alloc" {
        let op = args.next().expect("operation");
        let path = args.next().expect("path");
        let bytes: Arc<[u8]> = std::fs::read(&path).expect("read").into();
        // One untimed warm iteration so lazily-built process state is not
        // counted, then one measured lifecycle.
        let _ = run_alloc_op(&op, Arc::clone(&bytes));
        let before = (ALLOCATIONS.with(Cell::get), ALLOCATED_BYTES.with(Cell::get));
        let resolved = run_alloc_op(&op, Arc::clone(&bytes));
        let after = (ALLOCATIONS.with(Cell::get), ALLOCATED_BYTES.with(Cell::get));
        println!(
            "{path}\t{op}\tallocations={}\tallocated_bytes={}\tstring_cells={resolved}",
            after.0 - before.0,
            after.1 - before.1
        );
        return;
    }
    if mode == "time" {
        let kind = args.next().expect("source kind: owned or file");
        let op = args.next().expect("operation");
        let iterations: usize = args.next().expect("iterations").parse().expect("number");
        let path = args.next().expect("path");
        let bytes: Arc<[u8]> = std::fs::read(&path).expect("read").into();
        for _ in 0..iterations {
            let source: Arc<dyn ReadAt> = if kind == "file" {
                Arc::new(litchi_core::FileSource::open(&path).expect("open"))
            } else {
                Arc::new(Plain(Arc::clone(&bytes)))
            };
            let start = std::time::Instant::now();
            let ok = match op.as_str() {
                "validate" => litchi_xls::validation::validate_source(source).is_ok(),
                "all-cells" => run_all_cells(source),
                "full-text" => litchi_xls::SourceBackedWorkbook::from_read_at(source)
                    .map(|workbook| workbook.text().is_ok())
                    .unwrap_or(false),
                "open" => litchi_xls::SourceBackedWorkbook::from_read_at(source).is_ok(),
                other => panic!("unknown operation {other}"),
            };
            println!("{}\t{ok}", start.elapsed().as_nanos());
        }
        return;
    }
    let paths: Vec<String> = args.collect();
    for path in paths {
        let Ok(raw) = std::fs::read(&path) else {
            continue;
        };
        let bytes: Arc<[u8]> = raw.into();
        for op in OPS {
            if op == "one-cell" {
                continue;
            }
            let counts = measure(&bytes, op);
            println!(
                "{path}\t{op}\treads={}\tbytes={}\tsmall={}\tversions={}\tlens={}\t{}",
                counts.reads, counts.bytes, counts.small, counts.versions, counts.lens,
                counts.outcome
            );
        }
        if mode == "detail" {
            detail(&bytes, &path);
        }
    }
}

fn run_all_cells(source: Arc<dyn ReadAt>) -> bool {
    let Ok(workbook) = litchi_xls::SourceBackedWorkbook::from_read_at(source) else {
        return false;
    };
    let Ok(names) = workbook.worksheet_names() else {
        return false;
    };
    for index in 0..names.len() {
        if let Ok(Some(sheet)) = workbook.worksheet(index) {
            let mut seen = 0usize;
            if sheet
                .visit_cells(|_cell| {
                    seen += 1;
                    Ok(())
                })
                .is_ok()
                && seen > 0
            {
                return true;
            }
        }
    }
    false
}

fn detail(bytes: &Arc<[u8]>, path: &str) {
    let trace = Arc::new(Trace {
        bytes: Arc::clone(bytes),
        log: Mutex::new(Vec::new()),
        versions: Mutex::new(0),
        lens: Mutex::new(0),
    });
    let _ = litchi_xls::validation::validate_source(Arc::clone(&trace) as Arc<dyn ReadAt>);
    let log = trace.log.lock().unwrap();
    let mut histogram: std::collections::BTreeMap<usize, usize> = Default::default();
    for (_offset, length) in log.iter() {
        *histogram.entry(*length).or_default() += 1;
    }
    let mut sizes: Vec<_> = histogram.into_iter().collect();
    sizes.sort_by_key(|(_, count)| std::cmp::Reverse(*count));
    print!("  {path} validate sizes:");
    for (size, count) in sizes.iter().take(12) {
        print!(" {size}x{count}");
    }
    println!();
    print!("  {path} validate first16:");
    for (offset, length) in log.iter().take(16) {
        print!(" ({offset},{length})");
    }
    println!();
}

/// One complete lifecycle of `op`, returning the number of string-valued cells
/// it published, which is the resolve count for the walk operations.
fn run_alloc_op(op: &str, bytes: Arc<[u8]>) -> usize {
    let source: Arc<dyn ReadAt> = Arc::new(Plain(bytes));
    let Ok(workbook) = litchi_xls::SourceBackedWorkbook::from_read_at(source) else {
        return 0;
    };
    match op {
        "open" => 0,
        "full-text" => {
            let _ = workbook.text();
            0
        }
        _ => {
            let Ok(names) = workbook.worksheet_names() else {
                return 0;
            };
            let mut strings = 0usize;
            for index in 0..names.len() {
                if let Ok(Some(sheet)) = workbook.worksheet(index) {
                    let mut seen = 0usize;
                    let walked = sheet.visit_cells(|cell| {
                        seen += 1;
                        if matches!(cell.value(), litchi_core::sheet::CellValue::String(_)) {
                            strings += 1;
                        }
                        Ok(())
                    });
                    if walked.is_ok() && seen > 0 {
                        return strings;
                    }
                }
            }
            strings
        }
    }
}
