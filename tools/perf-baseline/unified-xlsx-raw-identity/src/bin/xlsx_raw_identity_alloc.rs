#[path = "../../../src/bin/support/counting_allocator.rs"]
mod allocator;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline::allocation_metrics::enable();
    let mut args = std::env::args_os();
    let _executable = args.next();
    match args.next() {
        Some(selector) if selector == std::ffi::OsStr::new("xlsx-raw-identity") => {
            litchi_perf_baseline::xlsx_raw_identity::run_from_args(args)
        },
        _ => Err("usage: xlsx-raw-identity-alloc xlsx-raw-identity [options]".into()),
    }
}
