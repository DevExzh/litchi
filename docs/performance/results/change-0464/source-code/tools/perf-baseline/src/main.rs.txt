#![forbid(unsafe_code)]

use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = std::env::args_os();
    let _program = args.next();
    match args.next() {
        Some(selector) if selector == std::ffi::OsStr::new("cache-retention") => {
            return litchi_perf_baseline::pptx_cache_retention::run_from_args(args);
        },
        Some(selector) if selector == std::ffi::OsStr::new("provider-lifecycle") => {
            return litchi_perf_baseline::pptx_provider_lifecycle::run_from_args(args);
        },
        Some(selector) if selector == std::ffi::OsStr::new("pptx-pair-lifecycle") => {
            return litchi_perf_baseline::pptx_pair_lifecycle::run_from_args(args);
        },
        Some(selector) if selector == std::ffi::OsStr::new("native-image-lifecycle") => {
            return litchi_perf_baseline::pptx_native_image::run_from_args(args);
        },
        Some(selector) if selector == std::ffi::OsStr::new("odp-append-attribution") => {
            return litchi_perf_baseline::odp_append_attribution::run_from_args(args);
        },
        _ => {},
    }
    litchi_perf_baseline::run()
}
