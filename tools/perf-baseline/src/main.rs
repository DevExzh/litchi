#![forbid(unsafe_code)]

use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = std::env::args_os();
    let _program = args.next();
    if args.next().as_deref() == Some(std::ffi::OsStr::new("cache-retention")) {
        return litchi_perf_baseline::pptx_cache_retention::run_from_args(args);
    }
    litchi_perf_baseline::run()
}
