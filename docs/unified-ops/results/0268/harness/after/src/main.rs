//! Candidate (uo2/s22-odf-w3) ODF open/edit/save timings.
#[path = "../../common.rs"]
mod common;
use common::*;
type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

fn odt_source(paragraphs: usize) -> Result<litchi_odt::Package> {
    let mut builder = litchi_odt::Builder::new();
    for index in 0..paragraphs {
        builder.add_paragraph(&format!("Paragraph {index} with some ordinary body text."))?;
    }
    Ok(builder.build()?)
}

fn ods_source(rows: usize, cols: usize) -> Result<litchi_ods::Package> {
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
    Ok(builder.build()?)
}

fn odp_source(slides: usize) -> Result<litchi_odp::Package> {
    let mut builder = litchi_odp::Builder::new();
    for index in 0..slides {
        builder.add_slide_with_title(&format!("Slide {index}"), "Body text for the slide.")?;
    }
    Ok(builder.build()?)
}

fn main() -> Result<()> {
    let iters: usize = std::env::args().nth(1).map_or(30, |v| v.parse().unwrap());
    let dir = std::env::args().nth(2).expect("output directory");
    let path = std::path::Path::new(&dir).join("after.out");
    let nosync = litchi_core::Durability::NoSync;

    // ODT
    let odt = odt_source(2000)?.to_bytes()?;
    eprintln!("odt bytes {} digest {:x}", odt.len(), digest(&odt));
    report("odt.build", time(iters, || { odt_source(2000).unwrap(); }));
    report("odt.open_package", time(iters, || {
        std::hint::black_box(litchi_odt::Package::from_bytes(&odt).unwrap());
    }));
    report("odt.open_projection", time(iters, || {
        let s = litchi_odt::Package::from_bytes(&odt).unwrap();
        std::hint::black_box(s.document().unwrap().paragraphs().unwrap().len());
    }));
    let package = litchi_odt::Package::from_bytes(&odt)?;
    report("odt.edit_commit", time(iters, || {
        let mut edit = package.edit();
        edit.replace_paragraph(litchi_odt::transaction::Position::new(0), "changed").unwrap();
        std::hint::black_box(edit.commit().unwrap().into_package().as_bytes().len());
    }));
    let mut edit = package.edit();
    edit.replace_paragraph(litchi_odt::transaction::Position::new(0), "changed")?;
    let committed = edit.commit()?.into_package();
    report("odt.save", time(iters, || { committed.save_with_durability(&path, nosync).unwrap(); }));
    report("odt.to_bytes", time(iters, || { std::hint::black_box(committed.to_bytes().unwrap()); }));
    report("odt.save_full", time(iters.min(10), || { committed.save(&path).unwrap(); }));

    // ODS
    let ods = ods_source(500, 10)?.to_bytes()?;
    eprintln!("ods bytes {} digest {:x}", ods.len(), digest(&ods));
    report("ods.build", time(iters, || { ods_source(500, 10).unwrap(); }));
    report("ods.open_package", time(iters, || {
        std::hint::black_box(litchi_ods::Package::from_bytes(&ods).unwrap());
    }));
    report("ods.open_projection", time(iters, || {
        let s = litchi_ods::Package::from_bytes(&ods).unwrap();
        std::hint::black_box(s.workbook().unwrap().sheet_count());
    }));
    let package = litchi_ods::Package::from_bytes(&ods)?;
    let ods_edit = |s: &litchi_ods::Package| {
        let mut edit = s.edit().unwrap();
        edit.worksheets(|w| { let _ = w.set_formula("Data", 0, 0, "of:=1+1")?; Ok(()) }).unwrap();
        edit.commit().unwrap().into_package()
    };
    report("ods.edit_commit", time(iters, || { std::hint::black_box(ods_edit(&package).as_bytes().len()); }));
    let committed = ods_edit(&package);
    report("ods.save", time(iters, || { committed.save_with_durability(&path, nosync).unwrap(); }));
    report("ods.to_bytes", time(iters, || { std::hint::black_box(committed.to_bytes().unwrap()); }));
    report("ods.save_full", time(iters.min(10), || { committed.save(&path).unwrap(); }));

    // ODP
    let odp = odp_source(200)?.to_bytes()?;
    eprintln!("odp bytes {} digest {:x}", odp.len(), digest(&odp));
    report("odp.build", time(iters, || { odp_source(200).unwrap(); }));
    report("odp.open_package", time(iters, || {
        let s = litchi_odp::Package::from_bytes(&odp).unwrap();
        std::hint::black_box(s.slides().unwrap().len());
    }));
    report("odp.open_projection", time(iters, || {
        let s = litchi_odp::Package::from_bytes(&odp).unwrap();
        std::hint::black_box(s.presentation().unwrap());
    }));
    let package = litchi_odp::Package::from_bytes(&odp)?;
    let _ = package.slides()?;
    let odp_edit = |s: &litchi_odp::Package| {
        let mut edit = s.edit().unwrap();
        edit.add("Added", "added text").unwrap();
        edit.commit().unwrap().into_package()
    };
    report("odp.edit_commit", time(iters, || { std::hint::black_box(odp_edit(&package).as_bytes().len()); }));
    let committed = odp_edit(&package);
    report("odp.save", time(iters, || { committed.save_with_durability(&path, nosync).unwrap(); }));
    report("odp.to_bytes", time(iters, || { std::hint::black_box(committed.to_bytes().unwrap()); }));
    report("odp.save_full", time(iters.min(10), || { committed.save(&path).unwrap(); }));
    Ok(())
}
