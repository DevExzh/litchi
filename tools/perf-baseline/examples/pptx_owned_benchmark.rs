//! Run one route of the paired PPTX owning-input benchmark control.

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args_os().skip(1);
    let route = args
        .next()
        .ok_or("expected ROUTE WARMUP SAMPLES NEW_OUTPUT_DIR (ROUTE is materialized or owned)")?;
    let warmup = args
        .next()
        .ok_or("expected WARMUP after ROUTE")?
        .to_string_lossy()
        .parse::<usize>()?;
    let samples = args
        .next()
        .ok_or("expected SAMPLES after WARMUP")?
        .to_string_lossy()
        .parse::<usize>()?;
    let output = args.next().ok_or("expected NEW_OUTPUT_DIR after SAMPLES")?;
    if args.next().is_some() {
        return Err(
            "expected ROUTE WARMUP SAMPLES NEW_OUTPUT_DIR (ROUTE is materialized or owned)".into(),
        );
    }
    litchi_perf_baseline::run_pptx_owned_benchmark(
        &route.to_string_lossy(),
        warmup,
        samples,
        std::path::Path::new(&output),
    )
}
