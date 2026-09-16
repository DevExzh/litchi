//! INSTRUMENTATION (change 0641, removed before commit).
//!
//! Counts allocation-chain links for one XLS scenario under three legs:
//! `base` (the cold per-sheet cursor construction this change replaces),
//! `hint` (this change: one worksheet-region position across the document's
//! sheets, disjoint from the shared-string resolver's), and `shared` (the
//! counterfactual change 0585 argued against without measuring: one position
//! serving both). Prints the link count at the end of the open and at the end
//! of the scenario, plus an FNV-1a digest of the result, so a leg that
//! produced different output says so.
use std::sync::Arc;
use std::sync::atomic::Ordering;

fn fnv1a(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

fn main() {
    let mut args = std::env::args().skip(1);
    let leg = args.next().expect("leg");
    let operation = args.next().expect("operation");
    let sheet: usize = args.next().expect("sheet").parse().expect("sheet index");
    let path = args.next().expect("path");

    match leg.as_str() {
        "base" => {
            litchi_xls::workbook::source::WORKSHEET_COLD.store(true, Ordering::Relaxed);
            litchi_xls::workbook::source::SHARED_HINT.store(false, Ordering::Relaxed);
        },
        "hint" => {
            litchi_xls::workbook::source::WORKSHEET_COLD.store(false, Ordering::Relaxed);
            litchi_xls::workbook::source::SHARED_HINT.store(false, Ordering::Relaxed);
        },
        "shared" => {
            litchi_xls::workbook::source::WORKSHEET_COLD.store(false, Ordering::Relaxed);
            litchi_xls::workbook::source::SHARED_HINT.store(true, Ordering::Relaxed);
        },
        other => panic!("unknown leg {other}"),
    }

    let bytes = std::fs::read(&path).expect("fixture");
    litchi_cfb::CHAIN_STEPS.store(0, Ordering::Relaxed);
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(bytes),
    ))
    .expect("open");
    let open_links = litchi_cfb::CHAIN_STEPS.load(Ordering::Relaxed);

    let digest = match operation.as_str() {
        "full-text" => match workbook.text() {
            Ok(text) => format!("text:{}:{:016x}", text.len(), fnv1a(text.as_bytes())),
            Err(error) => format!("refused:{error}"),
        },
        "all-cells" => {
            let worksheet = workbook
                .worksheet_by_index(sheet)
                .expect("worksheet")
                .expect("worksheet present");
            let mut seen = String::new();
            match worksheet.visit_cells(|cell| {
                seen.push_str(&format!("{},{}={:?}\n", cell.row(), cell.column(), cell.value()));
                Ok(())
            }) {
                Ok(()) => format!("cells:{:016x}", fnv1a(seen.as_bytes())),
                Err(error) => format!("refused:{error}"),
            }
        },
        "one-cell" => {
            let value = workbook.cell_value_by_index(sheet, 1, 0);
            format!("cell:{value:?}")
        },
        other => panic!("unknown operation {other}"),
    };
    let total_links = litchi_cfb::CHAIN_STEPS.load(Ordering::Relaxed);

    println!(
        "{leg}\t{operation}\t{path}\t{sheet}\topen={open_links}\ttotal={total_links}\tscenario={}\t{digest}",
        total_links - open_links
    );
}
