//! A 10,000-paragraph text fanout is one unified write (record 0251f).
//! Record 0089 measured it as a typed refusal: the native composite
//! replacement and the edit's operation ledger stopped at 4,096.
#![cfg(feature = "docx")]

use std::io::{Cursor, Write as _};

use litchi_ops::{OpenOptions, SourceMode, props};

fn fixture(paragraphs: usize) -> Vec<u8> {
    let body: String = (0..paragraphs)
        .map(|index| format!("<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"))
        .collect();
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, options).expect("start member");
        zip.write_all(contents.as_bytes()).expect("write member");
    }
    zip.finish().expect("finish fixture").into_inner()
}

#[test]
fn ten_thousand_paragraph_fanout_is_one_write_on_both_owners() {
    let bytes = fixture(10_000);
    for eager in [false, true] {
        let file = if eager {
            litchi_ops::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
        } else {
            litchi_ops::open_bytes(bytes.clone())
        }
        .expect("open the fixture");
        let changed = file
            .set("/body > p", props! { "text" => "Same\ttext" })
            .unwrap_or_else(|error| panic!("10k fanout (eager: {eager}): {error}"));
        let reopened =
            litchi_ops::open_bytes(changed.to_bytes().expect("publish")).expect("reopen");
        for path in ["/body/p[1]", "/body/p[5000]", "/body/p[10000]"] {
            let text = reopened
                .get(path)
                .expect("paragraph")
                .view()
                .text()
                .expect("text")
                .map(std::borrow::Cow::into_owned);
            assert_eq!(text.as_deref(), Some("Same\ttext"), "{path}");
        }
    }
}
