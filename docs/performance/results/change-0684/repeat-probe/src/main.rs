use litchi_core::{FileSource, OwnedSource, ReadAt};
use litchi_xls::SourceBackedWorkbook;
use std::{error::Error, hint::black_box, sync::Arc};

fn main() -> Result<(), Box<dyn Error>> {
    let a: Vec<String> = std::env::args().skip(1).collect();
    if a.len() != 6 {
        return Err("usage: xls0684-repeat MODE FILE SHEET ROW COLUMN REPETITIONS".into());
    }
    let source: Arc<dyn ReadAt> = match a[0].as_str() {
        "owned" => Arc::new(OwnedSource::new(std::fs::read(&a[1])?)),
        "file" => Arc::new(FileSource::open(&a[1])?),
        _ => return Err("mode must be owned or file".into()),
    };
    let owner = SourceBackedWorkbook::from_read_at(source)?;
    let sheet = a[2].parse()?;
    let row = a[3].parse()?;
    let column = a[4].parse()?;
    let repeats: usize = a[5].parse()?;
    for _ in 0..2 {
        black_box(owner.cell_value_by_index(sheet, row, column)?);
    }
    let started = std::time::Instant::now();
    let mut found = 0usize;
    for _ in 0..repeats {
        found += usize::from(black_box(owner.cell_value_by_index(sheet, row, column)?).is_some());
    }
    let elapsed = started.elapsed().as_nanos();
    println!("repeats={repeats}\tfound={found}\tnanos={elapsed}");
    black_box(owner);
    Ok(())
}
