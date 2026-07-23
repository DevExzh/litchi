use litchi_odf::{
    Document, MutableDocument, OdfFootnotePosition, OdfLineNumberFormat, OdfNoteClass,
    OdfNoteNumberingScope, OdfNotesConfiguration, OdfNotesConfigurations,
};
use std::io::{Cursor, Write};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn document(styles_body: &str) -> Document {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text><t:p>Body</t:p></o:text></o:body></o:document-content>"#
    );
    let styles = format!(
        r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:t="{TEXT}" o:version="1.3"><o:styles>{styles_body}</o:styles><o:automatic-styles/><o:master-styles/></o:document-styles>"#
    );
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(MIMETYPE.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file("styles.xml", deflated).unwrap();
    zip.write_all(styles.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    write!(
        zip,
        r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.3"><m:file-entry m:full-path="/" m:media-type="{MIMETYPE}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="styles.xml" m:media-type="text/xml"/></m:manifest>"#
    )
    .unwrap();
    zip.finish().unwrap();
    Document::from_bytes(output.into_inner()).unwrap()
}

fn footnote() -> OdfNotesConfiguration {
    OdfNotesConfiguration {
        note_class: OdfNoteClass::Footnote,
        citation_style_name: Some("FootnoteCitation".to_string()),
        citation_body_style_name: Some("FootnoteAnchor".to_string()),
        default_style_name: Some("Footnote".to_string()),
        master_page_name: Some("Standard".to_string()),
        start_value: Some(2),
        number_prefix: Some("[".to_string()),
        number_suffix: Some("]".to_string()),
        number_format: Some(OdfLineNumberFormat::LowerAlpha),
        letter_sync: Some(true),
        start_numbering_at: Some(OdfNoteNumberingScope::Chapter),
        footnotes_position: Some(OdfFootnotePosition::Page),
        continuation_notice_forward: Some("Continued on next page".to_string()),
        continuation_notice_backward: Some("Continued from previous page".to_string()),
    }
}

fn endnote() -> OdfNotesConfiguration {
    let mut configuration = OdfNotesConfiguration::new(OdfNoteClass::Endnote);
    configuration.start_value = Some(7);
    configuration.number_format = Some(OdfLineNumberFormat::UpperRoman);
    configuration.start_numbering_at = Some(OdfNoteNumberingScope::Document);
    configuration
}

#[test]
fn document_reads_note_configurations_from_styles() {
    let footnote = footnote();
    let endnote = endnote();
    let source = document(
        r#"<t:notes-configuration t:note-class="footnote" t:citation-style-name="FootnoteCitation" t:citation-body-style-name="FootnoteAnchor" t:default-style-name="Footnote" t:master-page-name="Standard" t:start-value="2" s:num-prefix="[" s:num-suffix="]" s:num-format="a" s:num-letter-sync="true" t:start-numbering-at="chapter" t:footnotes-position="page"><t:note-continuation-notice-forward>Continued on next page</t:note-continuation-notice-forward><t:note-continuation-notice-backward>Continued from previous page</t:note-continuation-notice-backward></t:notes-configuration><t:notes-configuration t:note-class="endnote" t:start-value="7" s:num-format="I" t:start-numbering-at="document"/>"#,
    );

    assert_eq!(
        source.notes_configurations().unwrap(),
        OdfNotesConfigurations {
            footnote: Some(footnote),
            endnote: Some(endnote),
        }
    );
}

#[test]
fn mutable_document_updates_note_configurations_without_touching_unrelated_styles() {
    let footnote = footnote();
    let endnote = endnote();
    let mut mutable =
        MutableDocument::from_document(document(r#"<s:style s:name="Keep"/>"#)).unwrap();

    assert_eq!(mutable.notes_configurations().unwrap(), Default::default());
    assert_eq!(mutable.set_notes_configuration(&footnote).unwrap(), None);
    assert_eq!(
        mutable.notes_configurations().unwrap().footnote,
        Some(footnote.clone())
    );

    assert_eq!(
        mutable
            .set_notes_configurations(&OdfNotesConfigurations {
                footnote: None,
                endnote: Some(endnote.clone()),
            })
            .unwrap()
            .footnote,
        Some(footnote)
    );
    assert_eq!(
        mutable.notes_configurations().unwrap(),
        OdfNotesConfigurations {
            footnote: None,
            endnote: Some(endnote.clone()),
        }
    );

    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.notes_configurations().unwrap().endnote,
        Some(endnote.clone())
    );
    let styles = String::from_utf8(reopened.get_file("styles.xml").unwrap()).unwrap();
    assert!(styles.contains(r#"<s:style s:name="Keep"/>"#));

    assert_eq!(
        mutable
            .clear_notes_configuration(OdfNoteClass::Endnote)
            .unwrap(),
        Some(endnote)
    );
    assert_eq!(mutable.notes_configurations().unwrap(), Default::default());
}
