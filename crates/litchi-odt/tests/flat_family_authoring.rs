use litchi_odt::elements::parser::OrderElement;
use litchi_odt::flat::{Document, Limits};

const SETTINGS: &str = r#"<office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewLeft" config:type="long">7</config:config-item></config:config-item-set></office:settings>"#;
const FLAT_TEXT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewLeft" config:type="long">7</config:config-item></config:config-item-set></office:settings><office:styles><style:style style:name="keep-me" style:family="paragraph"/></office:styles><office:master-styles/><office:body><office:text><text:p>Original intro</text:p><text:p>Cras eu leo sed justo</text:p><text:p>Structured <text:span>content</text:span></text:p></office:text></office:body></office:document>"#;

#[test]
fn flat_text_paragraph_edit_commits_and_round_trips() {
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut edit = source.edit();
    edit.update_paragraph(0, "Rewritten & Intro")
        .unwrap()
        .unwrap();
    let commit = edit.commit().unwrap();

    assert!(commit.diagnostics().is_empty());
    assert_eq!(source.as_bytes(), FLAT_TEXT.as_bytes());
    assert!(
        commit
            .document()
            .xml()
            .contains("<text:p>Rewritten &amp; Intro</text:p>")
    );
    assert!(commit.document().xml().contains("Cras eu leo sed justo"));
    assert!(commit.document().xml().contains("style:name=\"keep-me\""));
    assert!(!commit.document().xml().contains('\n'));

    let reopened = Document::from_bytes(commit.document().to_bytes()).unwrap();
    let paragraphs: Vec<_> = reopened
        .elements()
        .unwrap()
        .into_iter()
        .filter_map(|element| match element {
            OrderElement::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();
    assert_eq!(paragraphs[0].text().unwrap(), "Rewritten & Intro");
    assert_eq!(paragraphs[1].text().unwrap(), "Cras eu leo sed justo");

    let reverted = commit.patch().inverse().apply(commit.document()).unwrap();
    assert_eq!(reverted.as_bytes(), FLAT_TEXT.as_bytes());

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("roundtrip.fodt");
    commit.document().save(&path).unwrap();
    let reopened_from_disk = Document::open(path).unwrap();
    assert_eq!(reopened_from_disk.as_bytes(), commit.document().as_bytes());
}

#[test]
fn flat_text_mutation_preserves_settings_bytes() {
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut edit = source.edit();
    edit.update_paragraph(1, "Unrelated edit").unwrap().unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.document().xml().contains(SETTINGS));
}

#[test]
fn flat_text_missing_selector_returns_none_without_mutating_source() {
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut edit = source.edit();
    assert!(
        edit.update_paragraph(99, "Out of bounds")
            .unwrap()
            .is_none()
    );
    assert_eq!(
        edit.commit().unwrap().document().as_bytes(),
        source.as_bytes()
    );
    assert_eq!(source.as_bytes(), FLAT_TEXT.as_bytes());
}

#[test]
fn flat_text_structured_paragraph_edit_is_refused_losslessly() {
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut edit = source.edit();
    assert!(edit.update_paragraph(2, "Would discard markup").is_err());
    assert_eq!(source.as_bytes(), FLAT_TEXT.as_bytes());
}

#[test]
fn flat_text_empty_commit_is_an_exact_no_op() {
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let commit = source.edit().commit().unwrap();
    assert_eq!(commit.document().as_bytes(), source.as_bytes());
    assert_eq!(
        commit.patch().apply(&source).unwrap().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn formatted_flat_text_no_op_is_exact_but_changed_commit_refuses() {
    let formatted = FLAT_TEXT.replacen("<office:styles>", "\n<office:styles>", 1);
    let source = Document::from_bytes(formatted.as_bytes().to_vec()).unwrap();
    let no_op = source.edit().commit().unwrap();
    assert_eq!(no_op.document().as_bytes(), formatted.as_bytes());

    let mut edit = source.edit();
    edit.update_paragraph(0, "Changed").unwrap().unwrap();
    assert!(matches!(
        edit.commit(),
        Err(litchi_core::Error::XmlCompactness { .. })
    ));
    assert_eq!(source.as_bytes(), formatted.as_bytes());
}

#[test]
fn flat_text_limits_accept_exact_caps_and_reject_cap_plus_one() {
    let defaults = Limits::new();
    assert_eq!(
        defaults.max_document_bytes(),
        Limits::DEFAULT_MAX_DOCUMENT_BYTES
    );
    assert_eq!(
        defaults.max_paragraph_text_bytes(),
        Limits::DEFAULT_MAX_PARAGRAPH_TEXT_BYTES
    );
    assert_eq!(defaults.max_edits(), Limits::DEFAULT_MAX_EDITS);
    assert_eq!(defaults.max_xml_depth(), Limits::DEFAULT_MAX_XML_DEPTH);
    assert!(defaults.max_document_bytes() < Limits::HARD_MAX_DOCUMENT_BYTES);
    assert!(defaults.max_paragraph_text_bytes() < Limits::HARD_MAX_PARAGRAPH_TEXT_BYTES);
    assert!(defaults.max_edits() < Limits::HARD_MAX_EDITS);
    assert!(defaults.max_xml_depth() < Limits::HARD_MAX_XML_DEPTH);
    assert_eq!(
        defaults
            .with_max_document_bytes(Limits::HARD_MAX_DOCUMENT_BYTES)
            .unwrap()
            .max_document_bytes(),
        Limits::HARD_MAX_DOCUMENT_BYTES
    );
    assert_eq!(
        defaults
            .with_max_paragraph_text_bytes(Limits::HARD_MAX_PARAGRAPH_TEXT_BYTES)
            .unwrap()
            .max_paragraph_text_bytes(),
        Limits::HARD_MAX_PARAGRAPH_TEXT_BYTES
    );
    assert_eq!(
        defaults
            .with_max_edits(Limits::HARD_MAX_EDITS)
            .unwrap()
            .max_edits(),
        Limits::HARD_MAX_EDITS
    );
    assert_eq!(
        defaults
            .with_max_xml_depth(Limits::HARD_MAX_XML_DEPTH)
            .unwrap()
            .max_xml_depth(),
        Limits::HARD_MAX_XML_DEPTH
    );
    for (exact, too_large) in [
        (
            Limits::new().with_max_document_bytes(Limits::HARD_MAX_DOCUMENT_BYTES),
            Limits::new().with_max_document_bytes(Limits::HARD_MAX_DOCUMENT_BYTES + 1),
        ),
        (
            Limits::new().with_max_paragraph_text_bytes(Limits::HARD_MAX_PARAGRAPH_TEXT_BYTES),
            Limits::new().with_max_paragraph_text_bytes(Limits::HARD_MAX_PARAGRAPH_TEXT_BYTES + 1),
        ),
        (
            Limits::new().with_max_edits(Limits::HARD_MAX_EDITS),
            Limits::new().with_max_edits(Limits::HARD_MAX_EDITS + 1),
        ),
        (
            Limits::new().with_max_xml_depth(Limits::HARD_MAX_XML_DEPTH),
            Limits::new().with_max_xml_depth(Limits::HARD_MAX_XML_DEPTH + 1),
        ),
    ] {
        assert!(exact.is_ok());
        assert!(matches!(
            too_large,
            Err(litchi_core::Error::ResourceLimit(_))
        ));
    }

    let exact_document = Limits::new()
        .with_max_document_bytes(FLAT_TEXT.len())
        .unwrap();
    assert!(
        Document::from_bytes_with_limits(FLAT_TEXT.as_bytes().to_vec(), exact_document).is_ok()
    );
    let short_document = Limits::new()
        .with_max_document_bytes(FLAT_TEXT.len() - 1)
        .unwrap();
    assert!(
        Document::from_bytes_with_limits(FLAT_TEXT.as_bytes().to_vec(), short_document).is_err()
    );

    let text_limit = Limits::new().with_max_paragraph_text_bytes(3).unwrap();
    let source =
        Document::from_bytes_with_limits(FLAT_TEXT.as_bytes().to_vec(), text_limit).unwrap();
    assert!(source.edit().update_paragraph(0, "abc").is_ok());
    assert!(source.edit().update_paragraph(0, "abcd").is_err());

    let edit_limit = Limits::new().with_max_edits(2).unwrap();
    let source =
        Document::from_bytes_with_limits(FLAT_TEXT.as_bytes().to_vec(), edit_limit).unwrap();
    let mut edit = source.edit();
    edit.update_paragraph(0, "a").unwrap().unwrap();
    edit.update_paragraph(1, "b").unwrap().unwrap();
    assert!(edit.update_paragraph(2, "c").is_err());

    let depth_xml = r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text><t:p>x</t:p></o:text></o:body></o:document>"#;
    let exact_depth = Limits::new().with_max_xml_depth(4).unwrap();
    let source =
        Document::from_bytes_with_limits(depth_xml.as_bytes().to_vec(), exact_depth).unwrap();
    let mut edit = source.edit();
    edit.update_paragraph(0, "y").unwrap().unwrap();
    assert!(edit.commit().is_ok());
    let short_depth = Limits::new().with_max_xml_depth(3).unwrap();
    let source =
        Document::from_bytes_with_limits(depth_xml.as_bytes().to_vec(), short_depth).unwrap();
    let mut edit = source.edit();
    assert!(edit.update_paragraph(0, "y").is_err());
}

#[test]
fn concurrent_flat_text_saves_publish_only_complete_documents() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("shared.fodt");
    let first = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut second_edit = first.edit();
    second_edit
        .update_paragraph(0, "Second writer")
        .unwrap()
        .unwrap();
    let second = second_edit.commit().unwrap().into_document();
    let first_bytes = first.to_bytes();
    let second_bytes = second.to_bytes();
    let barrier = std::sync::Arc::new(std::sync::Barrier::new(3));
    let first_thread = {
        let barrier = barrier.clone();
        let path = path.clone();
        std::thread::spawn(move || {
            barrier.wait();
            first.save(path)
        })
    };
    let second_thread = {
        let barrier = barrier.clone();
        let path = path.clone();
        std::thread::spawn(move || {
            barrier.wait();
            second.save(path)
        })
    };
    barrier.wait();
    first_thread.join().unwrap().unwrap();
    second_thread.join().unwrap().unwrap();
    let published = std::fs::read(path).unwrap();
    assert!(published == first_bytes || published == second_bytes);
    assert!(std::fs::read_dir(directory.path()).unwrap().all(|entry| {
        !entry
            .unwrap()
            .file_name()
            .to_string_lossy()
            .ends_with(".tmp")
    }));
}

#[test]
fn failed_flat_text_save_preserves_foreign_temp_and_leaves_no_owned_temp() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("blocked.fodt");
    std::fs::create_dir(&path).unwrap();
    let foreign = directory.path().join(".blocked.fodt.litchi.tmp");
    std::fs::write(&foreign, b"foreign").unwrap();
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    assert!(source.save(&path).is_err());
    assert_eq!(std::fs::read(&foreign).unwrap(), b"foreign");
    assert_eq!(
        std::fs::read_dir(directory.path())
            .unwrap()
            .filter(|entry| entry.as_ref().unwrap().file_type().unwrap().is_file())
            .count(),
        1
    );
}

#[cfg(unix)]
#[test]
fn flat_text_save_refuses_destination_symlinks() {
    use std::os::unix::fs::symlink;

    let directory = tempfile::tempdir().unwrap();
    let victim = directory.path().join("victim.fodt");
    std::fs::write(&victim, b"victim").unwrap();
    let destination = directory.path().join("link.fodt");
    symlink(&victim, &destination).unwrap();
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    assert!(source.save(destination).is_err());
    assert_eq!(std::fs::read(victim).unwrap(), b"victim");
}

#[test]
fn flat_text_selector_ignores_paragraphs_nested_in_sections() {
    let xml = r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text><t:section t:name="nested"><t:p>Nested paragraph</t:p></t:section><t:p>Top paragraph</t:p></o:text></o:body></o:document>"#;
    let source = Document::from_bytes(xml.as_bytes().to_vec()).unwrap();
    let mut edit = source.edit();
    edit.update_paragraph(0, "Updated top").unwrap().unwrap();
    let committed = edit.commit().unwrap().into_document();
    assert!(committed.xml().contains("<t:p>Nested paragraph</t:p>"));
    assert!(committed.xml().contains("<t:p>Updated top</t:p>"));

    let nested_only = r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text><t:section t:name="nested"><t:p>Nested paragraph</t:p></t:section></o:text></o:body></o:document>"#;
    let source = Document::from_bytes(nested_only.as_bytes().to_vec()).unwrap();
    let mut edit = source.edit();
    assert!(edit.update_paragraph(0, "Must refuse").unwrap().is_none());
    assert_eq!(source.as_bytes(), nested_only.as_bytes());
}

#[test]
fn flat_text_save_supports_a_bare_relative_destination() {
    let reservation = tempfile::Builder::new()
        .prefix("litchi-flat-bare-")
        .suffix(".fodt")
        .tempfile_in(".")
        .unwrap();
    let bare = std::path::PathBuf::from(reservation.path().file_name().unwrap());
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    source.save(&bare).unwrap();
    assert_eq!(std::fs::read(&bare).unwrap(), source.as_bytes());
    drop(reservation);
}

#[cfg(all(unix, not(target_vendor = "apple")))]
#[test]
fn flat_text_save_supports_non_utf8_destinations() {
    use std::ffi::OsString;
    use std::os::unix::ffi::OsStringExt;

    let directory = tempfile::tempdir().unwrap();
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let non_utf8 = directory
        .path()
        .join(OsString::from_vec(b"flat-\xff.fodt".to_vec()));
    source.save(&non_utf8).unwrap();
    assert_eq!(std::fs::read(non_utf8).unwrap(), source.as_bytes());
}

#[cfg(unix)]
#[test]
fn flat_text_save_supports_near_name_max_destinations() {
    use std::ffi::OsString;
    use std::os::unix::ffi::OsStringExt;

    let directory = tempfile::tempdir().unwrap();
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    let mut long_name = vec![b'n'; 245];
    long_name.extend_from_slice(b".fodt");
    let near_name_max = directory.path().join(OsString::from_vec(long_name));
    source.save(&near_name_max).unwrap();
    assert_eq!(std::fs::read(near_name_max).unwrap(), source.as_bytes());
}

#[cfg(windows)]
#[test]
fn flat_text_save_refuses_windows_publication_without_atomic_primitive() {
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("unsupported.fodt");
    let source = Document::from_bytes(FLAT_TEXT.as_bytes().to_vec()).unwrap();
    assert!(matches!(
        source.save(&destination),
        Err(litchi_core::Error::Unsupported(_))
    ));
    assert!(!destination.exists());
}
