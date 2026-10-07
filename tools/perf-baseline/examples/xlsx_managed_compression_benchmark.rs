//! Timed opt-in replay of one bounded XLSX managed-compression case.

fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline::benchmark_xlsx_managed_compression(std::env::args_os().skip(1))
}
