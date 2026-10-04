//! Export the untimed owning-input PPTX preservation qualification fixtures.

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args_os().skip(1);
    let path = args.next().ok_or("expected a new output directory")?;
    if args.next().is_some() {
        return Err("expected exactly one output directory".into());
    }
    litchi_perf_baseline::write_pptx_owned_control_fixtures(std::path::Path::new(&path))
}
