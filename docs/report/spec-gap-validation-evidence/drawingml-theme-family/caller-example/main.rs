use std::{error::Error, fs, path::PathBuf};

use litchi_drawingml::theme::family::{self, Family, Snapshot};

const ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";

fn main() -> Result<(), Box<dyn Error>> {
    let evidence = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("..");
    let root = evidence.join("../../../../");
    let fixture = root.join("crates/litchi-drawingml/tests/fixtures/theme-family-native.xml");

    let authored = Family::new("Office\tFamily & <Variant>\r\n", ID, VID)?;
    fs::write(
        evidence.join("caller-authored.xml"),
        family::write(&authored)?,
    )?;

    let snapshot = Snapshot::from_xml(fs::read(fixture)?)?;
    let mut edit = snapshot.edit();
    edit.set_name("Office Theme (caller edit)")?;
    edit.set_variant_id("{00000000-0000-0000-0000-000000000001}")?;
    let commit = edit.commit()?;
    fs::write(
        evidence.join("caller-edited-native.xml"),
        commit.snapshot().xml_bytes(),
    )?;
    Ok(())
}
