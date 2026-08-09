#![cfg(all(feature = "docx", feature = "markdown"))]

use std::io::{Cursor, Write};

use litchi::{Document, markdown::ToMarkdown};
use zip::{CompressionMethod, ZipWriter, write::SimpleFileOptions};

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>"#;
const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
const DOCUMENT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/a%20b" TargetMode="External"/><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.png"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>"#;
const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
const FOOTNOTES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>note &amp; *body*</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
const ENDNOTES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>end note</w:t></w:r></w:p></w:endnote></w:endnotes>"#;
const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";

fn docx(document_xml: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut archive = ZipWriter::new(cursor);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("word/document.xml", document_xml.as_bytes()),
        ("word/_rels/document.xml.rels", DOCUMENT_RELS.as_bytes()),
        ("word/styles.xml", STYLES.as_bytes()),
        ("word/footnotes.xml", FOOTNOTES.as_bytes()),
        ("word/endnotes.xml", ENDNOTES.as_bytes()),
        ("word/media/p.png", PNG),
    ] {
        archive.start_file(name, options).unwrap();
        archive.write_all(contents).unwrap();
    }
    archive.finish().unwrap().into_inner()
}

#[test]
fn package_semantics_have_compact_golden_markdown() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:hyperlink r:id="rIdLink" w:tooltip="tip &quot;q&quot;"><w:r><w:rPr><w:b/></w:rPr><w:t>site [x]</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="Picture &quot;one&quot;" descr="alt [image]"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:p><w:r><w:t>With note</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> and end</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p></w:body></w:document>"#;
    let document = Document::from_bytes(docx(document_xml)).unwrap();

    assert_eq!(
        document.to_markdown().unwrap(),
        "Before [**site \\[x\\]**](<https://example.test/a%20b> \"tip \\\"q\\\"\") after\n\n![alt \\[image\\]](<data:image/png;base64,iVBORw0KGgo=> \"Picture &quot;one&quot;\")\n\nWith note[^fn-1] and end[^en-2]\n\n[^fn-1]: note \\& \\*body\\*\n[^en-2]: end note\n"
    );
}

#[test]
fn unsupported_link_placement_metadata_is_refused() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rIdLink" w:tgtFrame="_blank"><w:r><w:t>site</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#;
    let document = Document::from_bytes(docx(document_xml)).unwrap();
    assert!(matches!(
        document.to_markdown(),
        Err(litchi_core::Error::Unsupported(message)) if message.contains("w:tgtFrame")
    ));
}

#[test]
fn mixed_text_and_drawing_run_is_refused() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:t>before</w:t><w:drawing><wp:inline><wp:docPr id="1" name="p" descr="a"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#;
    let document = Document::from_bytes(docx(document_xml)).unwrap();
    assert!(matches!(
        document.to_markdown(),
        Err(litchi_core::Error::Unsupported(message)) if message.contains("interleaved DOCX text and embedded semantics")
    ));
}
