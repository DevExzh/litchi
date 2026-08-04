#[cfg(feature = "odf")]
use litchi::Result;
#[cfg(feature = "odf")]
use litchi::common::Metadata;
#[cfg(feature = "odf")]
use litchi::odf::odt::Builder;
#[cfg(feature = "odf")]
use litchi::odf::odt::elements::text::Paragraph;

#[cfg(feature = "odf")]
fn main() -> Result<()> {
    let output_file = "odt_page_styles_demo.odt";

    let mut builder = Builder::new();

    let metadata = Metadata {
        title: Some("ODT Page Style Demo".to_string()),
        author: Some("Litchi Examples".to_string()),
        subject: Some("ODT builder demonstration".to_string()),
        description: Some("Demonstrates the currently supported ODT builder features.".to_string()),
        ..Metadata::default()
    };
    builder.set_metadata(metadata);

    builder.add_heading("ODT Page Style Demo", 1)?;
    builder.add_paragraph(
        "This document demonstrates the currently supported ODT builder APIs: metadata, headings, paragraphs, rich text, and lists.",
    )?;
    builder.add_heading("Overview", 2)?;
    builder.add_paragraph(
        "The public builder currently focuses on core document content rather than advanced page-style or index authoring.",
    )?;
    builder.add_heading("Rich Text", 2)?;
    builder.add_rich_paragraph(vec![
        ("This paragraph mixes ", None),
        ("unstyled", Some("Standard")),
        (" and ", None),
        ("styled", Some("Emphasis")),
        (" spans through the supported builder API.", None),
    ])?;

    builder.add_heading("List Support", 2)?;
    builder.add_bulleted_list(vec![
        "Headings for structure",
        "Paragraphs for narrative text",
        "Lists for grouped content",
    ])?;

    let mut custom_paragraph = Paragraph::new();
    custom_paragraph.set_text("This paragraph was added as a Paragraph element directly.");
    builder.add_paragraph_element(custom_paragraph)?;

    builder.add_heading("Appendix", 2)?;
    builder.add_paragraph(
        "Use MutableDocument if you need to reopen and edit existing ODT files after generation.",
    )?;

    builder.save(output_file)?;
    println!("Created {output_file}");
    Ok(())
}

#[cfg(not(feature = "odf"))]
fn main() {}
