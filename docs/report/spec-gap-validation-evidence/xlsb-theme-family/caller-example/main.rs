//! Standalone downstream caller: public APIs only, save and reopen each edit.

use std::{error::Error, fs, io::Cursor, path::PathBuf};

use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::{Workbook, theme::Family};

fn save_reopen(workbook: &Workbook) -> Result<Workbook, Box<dyn Error>> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output)?;
    Ok(Workbook::new(Cursor::new(output.into_inner()))?)
}

fn main() -> Result<(), Box<dyn Error>> {
    let evidence = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("..");
    let root = evidence.join("../../../..");
    let input = fs::read(root.join("test-data/ooxml/xlsb/date.xlsb"))?;
    let mut workbook = Workbook::new(Cursor::new(input))?;
    let original = workbook.theme()?.ok_or("missing native Theme")?;
    let family = original.family().ok_or("missing native family")?;
    assert_eq!(family.name(), "Office Theme");

    let mut edit = original.edit();
    // A fresh typed value replaces scalars while retaining original opaque XML.
    edit.set_family(Family::new(
        "Caller\tTheme & <Family>\r\n",
        family.id().as_str(),
        "{00000000-0000-0000-0000-000000000001}",
    )?)?;
    let updated = edit.commit()?;
    workbook.apply_theme(&updated)?;
    workbook = save_reopen(&workbook)?;
    let snapshot = workbook.theme()?.ok_or("missing updated Theme")?;
    assert_eq!(
        snapshot.family().unwrap().name(),
        "Caller\tTheme & <Family>\r\n"
    );
    fs::write(evidence.join("caller-updated.xml"), snapshot.source_xml())?;
    // Source checks accept a separately reopened package with the same source.
    workbook.apply_theme_patch(&updated.patch().inverse())?;
    assert_eq!(
        workbook.theme()?.unwrap().source_xml(),
        original.source_xml()
    );

    let mut edit = original.edit();
    edit.remove_family()?;
    let removed = edit.commit()?;
    workbook.apply_theme(&removed)?;
    workbook = save_reopen(&workbook)?;
    let absent = workbook.theme()?.ok_or("missing retained Theme")?;
    assert!(absent.family().is_none());
    fs::write(evidence.join("caller-removed.xml"), absent.source_xml())?;

    let mut edit = absent.edit();
    edit.set_family(Family::new(
        "Caller authored family",
        family.id(),
        family.variant_id(),
    )?)?;
    let added = edit.commit()?;
    workbook.apply_theme(&added)?;
    workbook = save_reopen(&workbook)?;
    let present = workbook.theme()?.ok_or("missing added Theme")?;
    assert_eq!(present.family().unwrap().name(), "Caller authored family");
    fs::write(evidence.join("caller-added.xml"), present.source_xml())?;
    workbook.apply_theme_patch(&added.patch().inverse())?;
    workbook.apply_theme_patch(&removed.patch().inverse())?;
    assert_eq!(
        workbook.theme()?.unwrap().source_xml(),
        original.source_xml()
    );
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Family example"));
    writer.set_theme_family(Family::new(
        "Caller workbook family",
        family.id(),
        family.variant_id(),
    )?)?;
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output)?;
    let authored = Workbook::new(Cursor::new(output.into_inner()))?;
    let authored = authored.theme()?.ok_or("missing authored Theme")?;
    assert_eq!(authored.family().unwrap().name(), "Caller workbook family");
    fs::write(evidence.join("caller-authored.xml"), authored.source_xml())?;
    println!("native update/remove/add, save/reopen, exact inverses, and writer authoring passed");
    Ok(())
}
