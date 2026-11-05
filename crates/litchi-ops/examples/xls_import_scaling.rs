//! XLS CSV-import scaling probe (record 0359fa): times the leaf edit route
//! (1,024-cell chunks, as `File::import` stages them) and `File::import` of
//! the same `rows` x 5 table for each requested row count, so the per-row
//! cost can be compared across sizes.
//!
//! ```text
//! cargo build --release -p litchi-ops --features xls --example xls_import_scaling
//! taskset -c 2 target/release/examples/xls_import_scaling [leaf|unified|both] [rounds] 100 1000 10000
//! ```
//!
//! Each line reports the median and minimum over `rounds` samples in
//! milliseconds, the per-row microseconds of the median, and the size and
//! FNV-1a digest of the last published workbook, so two builds can be
//! compared for byte-identical output.

use std::time::Instant;

use litchi_core::FileFormat;
use litchi_ops::import::Options;
use litchi_ops::transaction::SemanticLimits;

fn table(rows: u32) -> Vec<(u32, String, f64, bool, String)> {
    (0..rows)
        .map(|index| {
            (
                index + 1,
                format!("name {index}"),
                f64::from(index % 1000) / 8.0 + 0.25,
                index % 3 == 0,
                format!("2026-01-{:02}", index % 28 + 1),
            )
        })
        .collect()
}

fn csv(table: &[(u32, String, f64, bool, String)]) -> Vec<u8> {
    let mut out = String::new();
    for (id, name, score, active, day) in table {
        out.push_str(&format!("{id},{name},{score},{active},{day}\n"));
    }
    out.into_bytes()
}

fn fnv1a(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf2_9ce4_8422_2325, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x0100_0000_01b3)
    })
}

fn leaf(bytes: &[u8], table: &[(u32, String, f64, bool, String)]) -> Vec<u8> {
    use litchi_xls::cell_values::{Reference, Selector, Value as Cell};
    let package = litchi_xls::Package::from_bytes(bytes).expect("leaf open");
    let sheet = || Selector::Position(litchi_core::Position::new(0));
    let mut current = package;
    for (chunk_index, chunk) in table.chunks(1_024 / 5).enumerate() {
        let mut edit = current.edit().expect("edit");
        for (offset, (id, name, score, active, day)) in chunk.iter().enumerate() {
            let r = u32::try_from(chunk_index * (1_024 / 5) + offset).expect("row");
            let at = |column| Reference::new(r, column).expect("reference");
            edit.insert_cell(sheet(), at(0), Cell::Number(f64::from(*id)))
                .expect("id");
            edit.insert_cell(sheet(), at(1), Cell::Text(name.clone()))
                .expect("name");
            edit.insert_cell(sheet(), at(2), Cell::Number(*score))
                .expect("score");
            edit.insert_cell(sheet(), at(3), Cell::Boolean(*active))
                .expect("active");
            edit.insert_cell(sheet(), at(4), Cell::Text(day.clone()))
                .expect("joined");
        }
        current = edit.commit().expect("commit").into_package();
    }
    current.to_bytes().expect("publish")
}

fn unified(bytes: &[u8], source: &[u8], cells: usize) -> Vec<u8> {
    let file = litchi_ops::open_bytes(bytes.to_vec()).expect("unified open");
    let limits = SemanticLimits::builder()
        .max_expanded_targets(1 << 22)
        .and_then(|builder| builder.max_staged_bytes(1 << 30))
        .and_then(|builder| builder.max_snapshot_bytes(1 << 32))
        .and_then(litchi_ops::transaction::SemanticLimitsBuilder::build)
        .expect("probe limits");
    let commit = file
        .import_with_options(
            "/sheet[1]",
            source,
            &Options::csv(),
            litchi_ops::EditOptions::new().semantic_limits(limits),
        )
        .expect("import");
    assert_eq!(commit.value().cells(), cells);
    commit.into_file().to_bytes().expect("publish")
}

fn main() {
    let mut args = std::env::args().skip(1);
    let route = args.next().unwrap_or_else(|| "both".to_owned());
    let rounds: usize = args
        .next()
        .map_or(3, |value| value.parse().expect("rounds"));
    let sizes: Vec<u32> = args.map(|value| value.parse().expect("rows")).collect();
    let sizes = if sizes.is_empty() {
        vec![100, 1_000]
    } else {
        sizes
    };
    let bytes = litchi_ops::create(FileFormat::Xls)
        .expect("create")
        .to_bytes()
        .expect("publish");
    for rows in sizes {
        let table = table(rows);
        let source = csv(&table);
        for leg in ["leaf", "unified"] {
            if route != "both" && route != leg {
                continue;
            }
            let mut samples = Vec::with_capacity(rounds);
            let mut output = Vec::new();
            for _ in 0..rounds {
                let start = Instant::now();
                output = if leg == "leaf" {
                    leaf(&bytes, &table)
                } else {
                    unified(&bytes, &source, table.len() * 5)
                };
                samples.push(start.elapsed().as_secs_f64() * 1e3);
            }
            samples.sort_by(f64::total_cmp);
            let median = samples[samples.len() / 2];
            println!(
                "{leg} rows={rows} median_ms={median:.3} min_ms={:.3} us_per_row={:.2} out_bytes={} out_fnv={:016x}",
                samples[0],
                median * 1e3 / f64::from(rows),
                output.len(),
                fnv1a(&output)
            );
        }
    }
}
