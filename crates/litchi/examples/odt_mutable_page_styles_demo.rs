#[cfg(feature = "odt")]
use litchi::Result;
#[cfg(feature = "odt")]
use litchi::common::Position;
#[cfg(feature = "odt")]
use litchi::odt::mutable::MutableDocument;
#[cfg(feature = "odt")]
use litchi::odt::{Builder, Document};

#[cfg(feature = "odt")]
fn main() -> Result<()> {
    let base_file = "odt_mutable_page_styles_base.odt";
    let output_file = "odt_mutable_page_styles_demo.odt";

    let mut builder = Builder::new();
    builder.add_heading("Mutable ODT Demo", 1)?;
    builder
        .add_paragraph("This seed file will be reopened and updated through MutableDocument.")?;
    builder.save(base_file)?;

    let document = Document::open(base_file)?;
    let mut mutable = MutableDocument::from_document(document)?;

    mutable.replace_paragraph_at(Position::new(0), "Mutable ODT Demo (reopened and updated)")?;

    mutable.add_paragraph(
        "The reopened document now includes extra paragraphs added through MutableDocument.",
    )?;
    mutable
        .add_paragraph("It also demonstrates in-place paragraph updates and metadata mutation.")?;

    mutable.metadata_mut().title = Some("Mutable ODT Demo".to_string());
    mutable.metadata_mut().subject = Some("MutableDocument example".to_string());

    println!("Paragraphs after mutation: {}", mutable.paragraphs().len());

    mutable.save(output_file)?;
    println!("Created {base_file} and {output_file}");
    Ok(())
}

#[cfg(not(feature = "odt"))]
fn main() {}
