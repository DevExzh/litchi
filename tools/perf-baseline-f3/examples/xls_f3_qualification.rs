fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline_f3::xls_f3_qualification::run_from_args(std::env::args().skip(1))
}
