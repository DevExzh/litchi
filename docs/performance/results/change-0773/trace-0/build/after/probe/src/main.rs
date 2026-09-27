//! Change 0773 syscall probe: one atomic save per window on every public save
//! route, delimited by `statx` calls on nonexistent marker paths so that an
//! `strace` of the process can be cut into windows.
//!
//! Default windows (`save`) exist in both legs and run first, in the same
//! order, so their preceding process state is identical. With the `levels`
//! feature (after leg only) each route also runs `save_with_durability` at
//! `full`, `file-only` and `no-sync`.
//!
//! usage: durability-probe DIR XLSB_FIXTURE

use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;

type Fallible = Result<(), Box<dyn std::error::Error>>;

fn marker(label: &str) {
    let _ = std::fs::metadata(format!("/litchi-0773-marker/{label}"));
}

/// Runs `save` between two markers, twice: once replacing an existing
/// destination (`ROUTE`), once creating an absent one (`ROUTE+create`).
fn window(directory: &Path, route: &str, level: &str, mut save: impl FnMut(&Path) -> Fallible) -> Fallible {
    replace_window(directory, route, level, &mut save)?;
    create_window(directory, &format!("{route}+create"), level, &mut save)
}

/// Seeds `destination` with old bytes (replacement of an existing file),
/// then runs `save` between two markers.
fn replace_window(directory: &Path, route: &str, level: &str, save: &mut dyn FnMut(&Path) -> Fallible) -> Fallible {
    let destination = directory.join(format!("{route}-{level}.out"));
    std::fs::write(&destination, b"old destination")?;
    run_window(route, level, &destination, save)
}

/// Makes sure `destination` does not exist, then runs `save` between two
/// markers.
fn create_window(directory: &Path, route: &str, level: &str, save: &mut dyn FnMut(&Path) -> Fallible) -> Fallible {
    let destination = directory.join(format!("{route}-{level}.out"));
    match std::fs::remove_file(&destination) {
        Ok(()) => {},
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {},
        Err(error) => return Err(error.into()),
    }
    run_window(route, level, &destination, save)
}

fn run_window(route: &str, level: &str, destination: &Path, save: &mut dyn FnMut(&Path) -> Fallible) -> Fallible {
    marker(&format!("{route}/{level}/begin"));
    let result = save(destination);
    marker(&format!("{route}/{level}/end"));
    result?;
    let published = std::fs::read(destination)?;
    println!("{route} {level} {} bytes", published.len());
    Ok(())
}

#[cfg(feature = "levels")]
const LEVELS: [(litchi_core::Durability, &str); 3] = [
    (litchi_core::Durability::Full, "full"),
    (litchi_core::Durability::FileOnly, "file-only"),
    (litchi_core::Durability::NoSync, "no-sync"),
];

fn cfb_source() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut writer = litchi_cfb::OleWriter::new();
    writer.create_stream(&["Document"], &vec![0x21; 5_003])?;
    writer.create_stream(&["Mini"], &[0x07; 97])?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn overlay_plan(file: &litchi_cfb::SharedOleFile) -> Result<litchi_cfb::ValidatedOverlayPlan, Box<dyn std::error::Error>> {
    Ok(file.plan_same_length_stream_overlays(
        vec![litchi_cfb::SameLengthStreamOverlay::new(
            vec!["Document".to_owned()],
            Arc::from(vec![0x5a; 5_003]),
        )],
        litchi_cfb::OverlayLimits::default(),
    )?)
}

fn sequential_writer(payload: &[u8]) -> Result<litchi_cfb::SequentialOleWriter<'_>, Box<dyn std::error::Error>> {
    let mut writer = litchi_cfb::SequentialOleWriter::new();
    writer.add_stream(&["Document"], payload.len() as u64, Cursor::new(payload))?;
    Ok(writer)
}

fn main() -> Fallible {
    let mut arguments = std::env::args().skip(1);
    let directory = PathBuf::from(arguments.next().ok_or("DIR")?);
    let xlsb_fixture = PathBuf::from(arguments.next().ok_or("XLSB_FIXTURE")?);
    std::fs::create_dir_all(&directory)?;

    // Inputs, built before any window.
    let mut docx = litchi_docx::Package::new()?;
    docx.document_mut()?.add_paragraph_with_text("save durability 0773");
    let xlsx = litchi_xlsx::Workbook::create()?;
    let mut pptx = litchi_pptx::Package::new()?;
    let xlsb = litchi_xlsb::Package::open(&xlsb_fixture)?;
    let mut docx_bytes = Cursor::new(Vec::new());
    litchi_docx::Package::new()?.to_stream(&mut docx_bytes)?;
    let opc = litchi_opc::OpcPackage::from_vec(docx_bytes.into_inner())?;
    let cfb_bytes = cfb_source()?;
    let mut ole = litchi_cfb::OleWriter::new();
    ole.create_stream(&["Document"], &vec![0x21; 5_003])?;
    ole.create_stream(&["Mini"], &[0x07; 97])?;
    let owned = litchi_cfb::SharedOleFile::open_owned(
        Arc::from(cfb_bytes.clone()),
        litchi_core::SourceVersion::new(0x761, 0),
    )?;
    let owned_plan = overlay_plan(&owned)?;
    let generic = litchi_cfb::SharedOleFile::open(Arc::new(litchi_core::OwnedSource::new(cfb_bytes)))?;
    let generic_plan = overlay_plan(&generic)?;
    let sequential_payload = vec![0x31; 5_003];
    let mut doc = litchi_doc::writer::Writer::new();
    doc.add_paragraph("save durability 0773")?;
    let mut xls = litchi_xls::writer::Writer::new();
    let sheet = xls.add_worksheet("Notes")?;
    xls.write_string(sheet, 0, 0, "save durability 0773")?;
    let mut ppt = litchi_ppt::writer::Writer::new();
    let slide = ppt.add_slide()?;
    ppt.add_textbox(slide, 40, 10, 300, 30, "save durability 0773")?;

    // Default windows: the documented `save` of every route.
    window(&directory, "opc", "save", |path| Ok(opc.save(path)?))?;
    window(&directory, "docx", "save", |path| Ok(docx.save(path)?))?;
    window(&directory, "xlsx", "save", |path| Ok(xlsx.save(path)?))?;
    window(&directory, "pptx", "save", |path| Ok(pptx.save(path)?))?;
    window(&directory, "xlsb", "save", |path| Ok(xlsb.save(path)?))?;
    window(&directory, "cfb-writer", "save", |path| Ok(ole.save(path)?))?;
    window(&directory, "cfb-sequential", "save", |path| {
        sequential_writer(&sequential_payload)?.save(path)?;
        Ok(())
    })?;
    window(&directory, "cfb-overlay-owned", "save", |path| {
        owned_plan.save(path)?;
        Ok(())
    })?;
    window(&directory, "cfb-overlay-generic", "save", |path| {
        generic_plan.save(path)?;
        Ok(())
    })?;
    window(&directory, "doc", "save", |path| Ok(doc.save(path)?))?;
    window(&directory, "xls", "save", |path| Ok(xls.save(path)?))?;
    window(&directory, "ppt", "save", |path| Ok(ppt.save(path)?))?;

    #[cfg(feature = "levels")]
    for (durability, level) in LEVELS {
        window(&directory, "opc", level, |path| Ok(opc.save_with_durability(path, durability)?))?;
        window(&directory, "docx", level, |path| Ok(docx.save_with_durability(path, durability)?))?;
        window(&directory, "xlsx", level, |path| Ok(xlsx.save_with_durability(path, durability)?))?;
        window(&directory, "pptx", level, |path| Ok(pptx.save_with_durability(path, durability)?))?;
        window(&directory, "xlsb", level, |path| Ok(xlsb.save_with_durability(path, durability)?))?;
        window(&directory, "cfb-writer", level, |path| {
            Ok(ole.save_with_durability(path, durability)?)
        })?;
        window(&directory, "cfb-sequential", level, |path| {
            sequential_writer(&sequential_payload)?.save_with_durability(path, durability)?;
            Ok(())
        })?;
        window(&directory, "cfb-overlay-owned", level, |path| {
            owned_plan.save_with_durability(path, durability)?;
            Ok(())
        })?;
        window(&directory, "cfb-overlay-generic", level, |path| {
            generic_plan.save_with_durability(path, durability)?;
            Ok(())
        })?;
        window(&directory, "doc", level, |path| Ok(doc.save_with_durability(path, durability)?))?;
        window(&directory, "xls", level, |path| Ok(xls.save_with_durability(path, durability)?))?;
        window(&directory, "ppt", level, |path| Ok(ppt.save_with_durability(path, durability)?))?;
    }
    Ok(())
}
