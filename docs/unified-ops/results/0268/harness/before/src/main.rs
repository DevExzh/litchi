//! Baseline (trunk 4d1afe92bc) ODF open/edit/save timings.
#[path = "../../common.rs"]
mod common;
use common::*;
use litchi_core::Result;

fn odt_source(paragraphs: usize) -> Result<Vec<u8>> {
    let mut builder = litchi_odt::Builder::new();
    for index in 0..paragraphs {
        builder.add_paragraph(&format!("Paragraph {index} with some ordinary body text."))?;
    }
    builder.build()
}

fn ods_source(rows: usize, cols: usize) -> Result<Vec<u8>> {
    use litchi_ods::{Builder, Cell, CellValue, Row, Sheet};
    let mut sheet = Sheet::new("Data")?;
    for row in 0..rows {
        let mut cells = Row::new();
        for col in 0..cols {
            let text = format!("r{row}c{col}");
            cells.push_cell(Cell::new(CellValue::Text(text.clone()), text))?;
        }
        sheet.push_row(cells)?;
    }
    let mut builder = Builder::new();
    builder.add_sheet(sheet)?;
    builder.build()
}

fn odp_source(slides: usize) -> Result<Vec<u8>> {
    let mut builder = litchi_odp::Builder::new();
    for index in 0..slides {
        builder.add_slide_with_title(&format!("Slide {index}"), "Body text for the slide.")?;
    }
    builder.build()
}

fn main() -> Result<()> {
    let iters: usize = std::env::args().nth(1).map_or(30, |v| v.parse().unwrap());
    let dir = std::env::args().nth(2).expect("output directory");
    let path = std::path::Path::new(&dir).join("before.out");

    // ODT
    let odt = odt_source(2000)?;
    eprintln!("odt bytes {} digest {:x}", odt.len(), digest(&odt));
    report("odt.build", time(iters, || { odt_source(2000).unwrap(); }));
    report("odt.open_package", time(iters, || {
        std::hint::black_box(litchi_odt::transaction::Snapshot::from_bytes(odt.clone()).unwrap());
    }));
    report("odt.open_projection", time(iters, || {
        let s = litchi_odt::transaction::Snapshot::from_bytes(odt.clone()).unwrap();
        std::hint::black_box(s.document().unwrap().paragraphs().unwrap().len());
    }));
    let snapshot = litchi_odt::transaction::Snapshot::from_bytes(odt.clone())?;
    report("odt.edit_commit", time(iters, || {
        let mut edit = snapshot.edit();
        edit.replace_paragraph(litchi_odt::transaction::Position::new(0), "changed").unwrap();
        std::hint::black_box(edit.commit().unwrap().into_snapshot().as_bytes().len());
    }));
    let mut edit = snapshot.edit();
    edit.replace_paragraph(litchi_odt::transaction::Position::new(0), "changed")?;
    let committed = edit.commit()?.into_snapshot();
    report("odt.save", time(iters, || { std::fs::write(&path, committed.as_bytes()).unwrap(); }));
    report("odt.to_bytes", time(iters, || { std::hint::black_box(committed.as_bytes().to_vec()); }));

    // ODS
    let ods = ods_source(500, 10)?;
    eprintln!("ods bytes {} digest {:x}", ods.len(), digest(&ods));
    report("ods.build", time(iters, || { ods_source(500, 10).unwrap(); }));
    report("ods.open_package", time(iters, || {
        std::hint::black_box(litchi_ods::document::Snapshot::from_bytes(ods.clone()).unwrap());
    }));
    report("ods.open_projection", time(iters, || {
        let s = litchi_ods::document::Snapshot::from_bytes(ods.clone()).unwrap();
        let w = litchi_ods::Spreadsheet::from_bytes(s.as_bytes().to_vec()).unwrap();
        std::hint::black_box(w.sheet_count());
    }));
    let snapshot = litchi_ods::document::Snapshot::from_bytes(ods.clone())?;
    let ods_edit = |s: &litchi_ods::document::Snapshot| {
        let mut edit = s.edit();
        edit.worksheets(|w| { let _ = w.set_formula("Data", 0, 0, "of:=1+1")?; Ok(()) }).unwrap();
        edit.commit().unwrap().into_snapshot()
    };
    report("ods.edit_commit", time(iters, || { std::hint::black_box(ods_edit(&snapshot).as_bytes().len()); }));
    let committed = ods_edit(&snapshot);
    report("ods.save", time(iters, || { std::fs::write(&path, committed.as_bytes()).unwrap(); }));
    report("ods.to_bytes", time(iters, || { std::hint::black_box(committed.as_bytes().to_vec()); }));

    // ODP
    let odp = odp_source(200)?;
    eprintln!("odp bytes {} digest {:x}", odp.len(), digest(&odp));
    report("odp.build", time(iters, || { odp_source(200).unwrap(); }));
    report("odp.open_package", time(iters, || {
        let s = litchi_odp::edit::Snapshot::from_bytes(odp.clone()).unwrap();
        std::hint::black_box(s.slides().len());
    }));
    report("odp.open_projection", time(iters, || {
        let s = litchi_odp::edit::Snapshot::from_bytes(odp.clone()).unwrap();
        std::hint::black_box(s.to_presentation().unwrap());
    }));
    let snapshot = litchi_odp::edit::Snapshot::from_bytes(odp.clone())?;
    let odp_edit = |s: &litchi_odp::edit::Snapshot| {
        let mut edit = s.transaction().unwrap();
        edit.add("Added", "added text").unwrap();
        edit.commit().unwrap().into_snapshot()
    };
    report("odp.edit_commit", time(iters, || { std::hint::black_box(odp_edit(&snapshot).bytes().len()); }));
    let committed = odp_edit(&snapshot);
    report("odp.save", time(iters, || { std::fs::write(&path, committed.bytes()).unwrap(); }));
    report("odp.to_bytes", time(iters, || { std::hint::black_box(committed.bytes().to_vec()); }));
    Ok(())
}
