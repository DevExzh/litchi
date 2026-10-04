//! Export source-backed XLSX XML attribute-cardinality fixtures.

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args_os().skip(1);
    let path = args
        .next()
        .ok_or("expected a new fixture output directory")?;
    let include_one_edit = match args.next() {
        None => false,
        Some(flag) if flag == "--include-one-edit" => true,
        Some(_) => {
            return Err("expected only --include-one-edit after the output directory".into());
        },
    };
    if args.next().is_some() {
        return Err("expected one output directory and an optional --include-one-edit".into());
    }
    litchi_perf_baseline::write_xlsx_attribute_fixtures(
        std::path::Path::new(&path),
        include_one_edit,
    )
}
