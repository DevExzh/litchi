#![cfg(all(feature = "docx", feature = "markdown"))]

use std::path::PathBuf;

use litchi::{Document, markdown::ToMarkdown};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/docx")
        .join(name)
}

/// Compact golden corpus from authored DOCX fixtures.  The expected values
/// intentionally contain only CommonMark-required whitespace: paragraph
/// separation is two newlines and no renderer indentation is permitted.
#[test]
fn real_docx_plain_paragraphs_have_compact_golden_markdown() -> litchi_core::Result<()> {
    for (name, expected) in [
        ("documentProperties.docx", "Hello World\\!\n\n"),
        ("documentProtection_no_protection.docx", "Non protetto.\n\n"),
        (
            "documentProtection_comments_no_password.docx",
            "Comments senza password.\n\n",
        ),
    ] {
        assert_eq!(
            Document::open(fixture(name))?.to_markdown()?,
            expected,
            "{name}"
        );
    }
    Ok(())
}
