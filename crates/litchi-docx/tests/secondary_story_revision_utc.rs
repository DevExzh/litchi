use litchi_docx::Package;
use litchi_docx::header_footer::Kind;
use litchi_docx::section::Start;
use litchi_docx::writer::SectionProperties;
use std::io::Cursor;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const DATE_UTC: &str = "http://schemas.microsoft.com/office/word/2023/wordml/word16du";

fn revision_story(root: &str, alias: &str, timestamp: &str) -> String {
    format!(
        r#"<{alias}:{root} xmlns:{alias}="{WORD}" xmlns:m="{MCE}" xmlns:du="{DATE_UTC}" m:Ignorable="du"><{alias}:p><{alias}:ins {alias}:id="1" {alias}:author="Alice" du:dateUtc="{timestamp}"><{alias}:r><{alias}:t>tracked</{alias}:t></{alias}:r></{alias}:ins></{alias}:p></{alias}:{root}>"#
    )
}

#[test]
fn header_and_footer_revision_utc_survive_package_save_and_reopen() {
    let mut package = Package::new().unwrap();
    let header_xml = revision_story("hdr", "h", "2026-07-17T00:00:00Z");
    let footer_xml = revision_story("ftr", "f", "2026-07-18T00:00:00Z");
    {
        let document = package.document_mut().unwrap();
        document.add_paragraph_with_text("body");
        let mut section = SectionProperties::default().with_start_type(Start::NewPage);
        section
            .set_header_part(Kind::Primary, "utc-header", header_xml.as_str())
            .unwrap();
        document.insert_section_break(0, section).unwrap();
        document
            .section_mut()
            .set_footer_part(Kind::Primary, "utc-footer", footer_xml.as_str())
            .unwrap();
    }

    let mut bytes = Cursor::new(Vec::new());
    package.to_stream(&mut bytes).unwrap();
    let original_bytes = bytes.into_inner();
    let mut reopened = Package::from_reader(Cursor::new(original_bytes.as_slice())).unwrap();
    let document = reopened.document().unwrap();

    let header = document.headers().unwrap().into_iter().next().unwrap();
    assert_eq!(header.xml_bytes(), header_xml.as_bytes());
    let header_revision = header.paragraphs().unwrap()[0].revisions().unwrap();
    assert_eq!(header_revision.len(), 1);
    assert_eq!(header_revision[0].date_utc(), Some("2026-07-17T00:00:00Z"));
    assert!(
        header
            .xml_bytes()
            .windows(b"m:Ignorable=\"du\"".len())
            .any(|window| { window == b"m:Ignorable=\"du\"" })
    );

    let footer = document.footers().unwrap().into_iter().next().unwrap();
    assert_eq!(footer.xml_bytes(), footer_xml.as_bytes());
    let footer_revision = footer.paragraphs().unwrap()[0].revisions().unwrap();
    assert_eq!(footer_revision.len(), 1);
    assert_eq!(footer_revision[0].date_utc(), Some("2026-07-18T00:00:00Z"));
    assert!(
        footer
            .xml_bytes()
            .windows(b"du:dateUtc=\"2026-07-18T00:00:00Z\"".len())
            .any(|window| window == b"du:dateUtc=\"2026-07-18T00:00:00Z\"")
    );
    let mut unchanged = Cursor::new(Vec::new());
    reopened.to_stream(&mut unchanged).unwrap();
    assert_eq!(unchanged.into_inner(), original_bytes);
}
