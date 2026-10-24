fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline::ppt_lazy_native::run_from_args(std::env::args().skip(1), false)
}
