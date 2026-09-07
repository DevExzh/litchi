//! Export explicit ODP source/candidate tail-append fixture archives.
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args_os().skip(1);
    let path = args
        .next()
        .ok_or("expected a new fixture output directory")?;
    if args.next().is_some() {
        return Err("expected exactly one output directory".into());
    }
    litchi_perf_baseline::write_odp_source_tail_append_fixtures(std::path::Path::new(&path))
}
