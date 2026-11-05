//! Entrypoint for the memo 0359a-F3 native XLS source-backed qualification.

fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline::xls_f3_qualification::run_from_args(std::env::args().skip(1))
}
