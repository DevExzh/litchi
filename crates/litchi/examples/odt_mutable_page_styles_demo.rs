#[cfg(feature = "odt")]
use litchi::Result;
#[cfg(feature = "odt")]
use litchi::common::Position;
#[cfg(feature = "odt")]
use litchi::odt::metadata::Field;
#[cfg(feature = "odt")]
use litchi::odt::{Builder, Package};

#[cfg(feature = "odt")]
fn main() -> Result<()> {
    let base_file = "odt_mutable_page_styles_base.odt";
    let output_file = "odt_mutable_page_styles_demo.odt";

    let mut builder = Builder::new();
    builder.add_heading("ODT Edit Demo", 1)?;
    builder.add_paragraph("This seed file will be reopened and updated through Package::edit.")?;
    builder.build()?.save(base_file)?;

    let package = Package::open(base_file)?;
    let mut edit = package.edit();
    edit.replace_paragraph(Position::new(0), "ODT Edit Demo (reopened and updated)")?;
    edit.add_paragraph(
        "The reopened document now includes extra paragraphs added through Package::edit.",
    )?;
    edit.add_paragraph("It also demonstrates in-place paragraph updates and metadata edits.")?;
    let edited = edit.commit()?.into_package();

    let mut metadata = edited.edit_metadata();
    metadata.set(Field::Title, Some("ODT Edit Demo"))?;
    metadata.set(Field::Subject, Some("Package::edit example"))?;
    let edited = metadata.commit()?.into_package();

    println!(
        "Paragraphs after editing: {}",
        edited.document()?.paragraphs()?.len()
    );

    edited.save(output_file)?;
    println!("Created {base_file} and {output_file}");
    Ok(())
}

#[cfg(not(feature = "odt"))]
fn main() {}
