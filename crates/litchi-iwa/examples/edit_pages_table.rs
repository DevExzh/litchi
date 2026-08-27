use std::env;

use litchi_iwa::pages::{PagesCellValue, PagesEditor};
use litchi_pages::{BodyTableSelector, Package as PagesPackage};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let input = args.next().ok_or("usage: edit_pages_table INPUT OUTPUT")?;
    let output = args.next().ok_or("usage: edit_pages_table INPUT OUTPUT")?;
    let mut editor = PagesEditor::open(input)?;
    let table = editor
        .tables()?
        .into_iter()
        .next()
        .ok_or("the Pages body contains no table")?;
    editor.set_table_cell(
        table.model_object_id,
        2,
        0,
        PagesCellValue::Text("Updated by litchi-iwa".to_owned()),
    )?;
    let table_position = editor
        .tables()?
        .iter()
        .position(|candidate| candidate.model_object_id == table.model_object_id)
        .ok_or("the Pages table disappeared before rename")?;
    let selector = BodyTableSelector::index(table_position);
    let focused = PagesPackage::from_bytes(&editor.to_bytes()?)?;
    let commit = focused
        .edit_body_table_name(selector)?
        .set_name("Edited Table")?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    let verified = PagesPackage::from_bytes(&bytes)?;
    if verified.body_table_name(selector)?.as_str() != "Edited Table" {
        return Err("Pages table rename failed focused-package verification".into());
    }
    editor = PagesEditor::from_bytes(&bytes)?;
    editor.resize_table(
        table.model_object_id,
        table.rows.max(5),
        table.columns.max(4),
    )?;
    editor.save(output)?;
    Ok(())
}
