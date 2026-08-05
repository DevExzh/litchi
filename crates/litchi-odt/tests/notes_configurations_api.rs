use litchi_odt::Document;
use litchi_odt::line_numbering::LineNumberFormat;
use litchi_odt::notes_configuration::{
    Class, Configuration, Configurations, NumberingScope, Position,
};
mod support;

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
    Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", content.as_bytes()),
            ("styles.xml", styles.as_bytes()),
        ],
    ))
    .unwrap()
}

fn footnote() -> Configuration {
    Configuration {
        note_class: Class::Footnote,
        citation_style_name: Some("FootnoteCitation".to_string()),
        citation_body_style_name: Some("FootnoteAnchor".to_string()),
        default_style_name: Some("Footnote".to_string()),
        master_page_name: Some("Standard".to_string()),
        start_value: Some(2),
        number_prefix: Some("[".to_string()),
        number_suffix: Some("]".to_string()),
        number_format: Some(LineNumberFormat::LowerAlpha),
        letter_sync: Some(true),
        start_numbering_at: Some(NumberingScope::Chapter),
        footnotes_position: Some(Position::Page),
        continuation_notice_forward: Some("Continued on next page".to_string()),
        continuation_notice_backward: Some("Continued from previous page".to_string()),
    }
}

fn endnote() -> Configuration {
    let mut configuration = Configuration::new(Class::Endnote);
    configuration.start_value = Some(7);
    configuration.number_format = Some(LineNumberFormat::UpperRoman);
    configuration.start_numbering_at = Some(NumberingScope::Document);
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
        Configurations {
            footnote: Some(footnote),
            endnote: Some(endnote),
        }
    );
}

#[test]
fn mutable_document_updates_note_configurations_without_touching_unrelated_styles() {
    let footnote = footnote();
    let endnote = endnote();
    let mut mutable = litchi_odt::mutable::MutableDocument::from_document(document(
        r#"<s:style s:name="Keep"/>"#,
    ))
    .unwrap();

    assert_eq!(mutable.notes_configurations().unwrap(), Default::default());
    assert_eq!(mutable.set_notes_configuration(&footnote).unwrap(), None);
    assert_eq!(
        mutable.notes_configurations().unwrap().footnote,
        Some(footnote.clone())
    );

    assert_eq!(
        mutable
            .set_notes_configurations(&Configurations {
                footnote: None,
                endnote: Some(endnote.clone()),
            })
            .unwrap()
            .footnote,
        Some(footnote)
    );
    assert_eq!(
        mutable.notes_configurations().unwrap(),
        Configurations {
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
        mutable.clear_notes_configuration(Class::Endnote).unwrap(),
        Some(endnote)
    );
    assert_eq!(mutable.notes_configurations().unwrap(), Default::default());
}
