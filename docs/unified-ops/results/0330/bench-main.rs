//! 1k coalesced XLSX cell sets in one batch (record 0330 cell_plan check).
use litchi_core::FileFormat;
use litchi_ops::{BatchItem, BatchMode, Props};
use std::time::Instant;

fn main() {
    let rounds: usize = std::env::args().nth(1).and_then(|a| a.parse().ok()).unwrap_or(30);
    let file = litchi_ops::create(FileFormat::Xlsx).unwrap();
    let paths: Vec<String> = (1..=1000).map(|i| format!("/Sheet1/A{i}")).collect();
    let items: Vec<BatchItem<'_>> = paths
        .iter()
        .enumerate()
        .map(|(i, path)| {
            let mut item = BatchItem::new("set");
            item.path = Some(path.as_str().into());
            item.props = Some(Props::new().with("value", i as i64));
            item
        })
        .collect();
    for _ in 0..5 {
        std::hint::black_box(file.batch(&items, BatchMode::Atomic).unwrap());
    }
    let mut samples = Vec::with_capacity(rounds);
    for _ in 0..rounds {
        let start = Instant::now();
        let commit = file.batch(&items, BatchMode::Atomic).unwrap();
        samples.push(start.elapsed().as_nanos() as u64);
        std::hint::black_box(commit);
    }
    samples.sort_unstable();
    println!("{}", samples[samples.len() / 2]);
}
