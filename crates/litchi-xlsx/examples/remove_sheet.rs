use std::error::Error;
use std::io;

fn main() -> Result<(), Box<dyn Error>> {
    let output = std::env::args_os().nth(1).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: cargo run -p litchi-xlsx --example remove_sheet -- OUTPUT.xlsx",
        )
    })?;

    let baseline = litchi_xlsx::Package::new()?;
    let mut create = baseline.edit()?;
    create.add("Scratch")?.set("A1", "temporary")?.activate();
    create
        .add("Results")?
        .set("A1", "retained")?
        .set("B2", 42_i32)?;
    let source = create.commit()?.into_package();

    let mut edit = source.edit()?;
    edit.remove("Scratch")?.ok_or_else(|| {
        io::Error::new(io::ErrorKind::NotFound, "workbook has no Scratch worksheet")
    })?;
    let committed = edit.commit()?;
    committed.package().save(output)?;

    let active = committed
        .package()
        .workbook_ref()
        .active_sheet()
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidData, "active sheet disappeared"))?;
    println!(
        "removed Scratch in {} semantic change(s); active sheet is {}",
        committed.patch().operation_count(),
        active.name()
    );
    Ok(())
}
