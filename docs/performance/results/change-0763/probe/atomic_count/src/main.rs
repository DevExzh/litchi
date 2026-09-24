//! Budget-update count probe for change 0763 (built against the candidate).
//!
//! Runs the harness's large streaming DOCX and XLSX scripts (131,072
//! paragraphs or rows) against a single root budget and reads every counter
//! after every API call. With one holder and one level, every claim or
//! release is one atomic read-modify-write that changes its counter, and a
//! charge covered by a lease changes nothing, so the changes observed between
//! calls count the lease claims exactly (no call here charges more than one
//! chunk of a resource). Output bytes are still reserved per sink write and
//! committed whole, one atomic each, counted with a counting sink.

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
};
use litchi_docx::{StreamingDocumentLimits, StreamingDocumentWriter};
use litchi_xlsx::{StreamingCell, StreamingCellValue, StreamingWorkbookLimits, StreamingWorkbookWriter};
use std::io::Write;
use std::num::{NonZeroU64, NonZeroUsize};

const LEASED: [Resource; 3] = [Resource::Objects, Resource::Work, Resource::InputBytes];

#[derive(Default)]
struct CountingSink {
    writes: u64,
    bytes: u64,
}

impl Write for CountingSink {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.writes += 1;
        self.bytes += buffer.len() as u64;
        Ok(buffer.len())
    }
    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn context(budget: Budget) -> ExecutionContext {
    let one = NonZeroUsize::new(1).unwrap();
    let (_source, token) = CancellationSource::pair();
    ExecutionContext::new(budget, token, ExecutionLimits::new(one, one, NonZeroU64::new(1 << 20).unwrap(), 0).unwrap())
}

struct Watch {
    last: [u64; 3],
    changes: [u64; 3],
}

impl Watch {
    fn new(budget: &Budget) -> Self {
        Self { last: LEASED.map(|r| budget.used(r)), changes: [0; 3] }
    }
    fn observe(&mut self, budget: &Budget) {
        for (index, &resource) in LEASED.iter().enumerate() {
            let now = budget.used(resource);
            if now != self.last[index] {
                self.changes[index] += 1;
                self.last[index] = now;
            }
        }
    }
}

fn main() {
    let paragraphs = 131_072_u64;
    let texts: Vec<String> = (0..paragraphs).map(|i| format!("litchi-perf-docx-streaming-v1-{i:06}-café-<&>")).collect();
    let input: u64 = texts.iter().map(|t| t.len() as u64).sum();
    let max_run = texts.iter().map(|t| t.len() as u64).max().unwrap();
    let document_xml = (max_run * 8 + 256) * paragraphs + 16 * 1024;
    let output = document_xml + 256 * 1024;
    let limits = StreamingDocumentLimits::new(input, output, paragraphs, paragraphs, 1, max_run, document_xml, document_xml, 64);
    let budget = Budget::root("count", Limits::new(64, input, output, paragraphs * 2 + 16, 32, paragraphs * 4 + input + 32));
    let mut writer = StreamingDocumentWriter::new(CountingSink::default(), context(budget.clone()), limits).unwrap();
    let mut watch = Watch::new(&budget);
    let mut calls = 0_u64;
    for text in &texts {
        writer.start_paragraph().unwrap();
        watch.observe(&budget);
        writer.start_run().unwrap();
        watch.observe(&budget);
        writer.write_text(text).unwrap();
        watch.observe(&budget);
        writer.finish_run().unwrap();
        watch.observe(&budget);
        writer.finish_paragraph().unwrap();
        watch.observe(&budget);
        calls += 5;
    }
    let sink = writer.finish().unwrap();
    // `finish` returns the leases: one release per resource that holds units.
    watch.observe(&budget);
    let exact_charges: u64 = texts
        .iter()
        .map(|t| 2 + 2 + (t.chars().count() as u64).div_ceil(64) + 1 + 1 + 1)
        .sum();
    let leased = watch.changes.iter().sum::<u64>();
    println!("docx large: {calls} calls; leased-resource counter changes (claims + final releases): objects {}, work {}, input {} = {leased}", watch.changes[0], watch.changes[1], watch.changes[2]);
    println!("docx large: output-byte reservations (one per sink write) {}; memory reserve+release 2; fixed new() charges 2", sink.writes);
    println!("docx large: budget atomics per iteration with leases = {}", leased + sink.writes + 2 + 2);
    println!("docx large: exact accounting made {exact_charges} Objects/Work/InputBytes charges (one atomic each) plus the same output and fixed ones = {}", exact_charges + sink.writes + 2 + 2);
    println!("docx large: settled counters objects {} work {} input {}", budget.used(Resource::Objects), budget.used(Resource::Work), budget.used(Resource::InputBytes));

    let rows = 131_072_u64;
    let cells = rows * 4;
    let max_sheet = rows * 512 + 4 * 1024;
    let max_output = max_sheet * 2 + 64 * 1024;
    let objects = rows + cells + 16;
    let limits = StreamingWorkbookLimits::new(u32::try_from(rows).unwrap(), cells, 256, 4 * 1024, max_sheet, max_output);
    let budget = Budget::root("count", Limits::new(4 * 1024, 0, max_output, objects, 32, objects * 2));
    let mut writer = StreamingWorkbookWriter::new(CountingSink::default(), context(budget.clone()), limits).unwrap();
    let mut watch = Watch::new(&budget);
    for row in 1..=rows {
        let text = format!("litchi-perf-streaming-xlsx-row-{row:06}-café-<&>");
        let number = u32::try_from(row).unwrap();
        writer
            .write_row(
                number,
                [
                    StreamingCell::new(1, StreamingCellValue::Number(f64::from(number))),
                    StreamingCell::new(2, StreamingCellValue::Text(&text)),
                    StreamingCell::new(3, StreamingCellValue::Bool(row % 2 == 0)),
                    StreamingCell::new(4, StreamingCellValue::Blank),
                ],
            )
            .unwrap();
        watch.observe(&budget);
    }
    let sink = writer.finish().unwrap();
    watch.observe(&budget);
    let leased = watch.changes.iter().sum::<u64>();
    println!("xlsx large: {rows} rows; leased-resource counter changes: objects {}, work {} = {leased}", watch.changes[0], watch.changes[1]);
    println!("xlsx large: output-byte reservations {}; budget atomics with leases = {} (+ fixed new(): 1 memory, 7 objects/work, and the memory release)", sink.writes, leased + sink.writes);
    println!("xlsx large: exact accounting made {} row charges (1 row Work + 4 cell Work + 1 Objects per row) + 1 finish Objects", rows * 6);
    println!("xlsx large: settled counters objects {} work {}", budget.used(Resource::Objects), budget.used(Resource::Work));
}
