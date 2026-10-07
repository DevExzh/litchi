//! Qualify explicit bounded cold XLSX compression contexts without timing.

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args_os().skip(1);
    let directory = args.next().ok_or("expected a new output directory")?;
    if args.next().is_some() {
        return Err("expected exactly one new output directory".into());
    }
    litchi_perf_baseline::qualify_xlsx_managed_compression(std::path::Path::new(&directory))
}
