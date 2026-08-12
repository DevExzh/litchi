#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_odf_common::package::raw_identical_members;
use litchi_odp::content::{
    MAX_TEXT_BOX_MODEL_REPLACEMENTS, Paragraph, TextBoxModel, TextBoxModelReplacement,
};
use litchi_odp::{Presentation, edit};
use soapberry_zip::office::StreamingArchiveWriter;

const MIME: &str = "application/vnd.oasis.opendocument.presentation";

#[derive(Clone, Copy, Default)]
struct FixtureOptions {
    protected: Option<(usize, usize)>,
    opaque: Option<(usize, usize)>,
    processing_instruction: Option<(usize, usize)>,
    media: bool,
}

fn fixture(slides: usize, boxes_per_slide: usize, options: FixtureOptions) -> Vec<u8> {
    let mut content = String::from(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:future="urn:litchi:test:future"><office:body><office:presentation>"#,
    );
    for slide in 0..slides {
        content.push_str(&format!(r#"<draw:page draw:name="Slide {slide}">"#));
        for object in 0..boxes_per_slide {
            content.push_str(&format!(r#"<draw:frame draw:name="Box {slide}-{object}""#));
            if options.protected == Some((slide, object)) {
                content.push_str(r#" draw:protect="content""#);
            }
            content.push('>');
            if options.processing_instruction == Some((slide, object)) {
                content.push_str("<?producer retained?>");
            }
            content.push_str(&format!(
                "<draw:text-box><text:p>before {slide}-{object}</text:p></draw:text-box>"
            ));
            if options.opaque == Some((slide, object)) {
                content.push_str(r#"<future:opaque future:value="retained"/>"#);
            }
            content.push_str("</draw:frame>");
        }
        content.push_str(r#"<future:unselected future:value="byte-exact"/></draw:page>"#);
    }
    content.push_str("</office:presentation></office:body></office:document-content>");

    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer
        .add_file(
            "styles.xml",
            br#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#,
        )
        .unwrap();
    writer
        .add_file(
            "meta.xml",
            br#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#,
        )
        .unwrap();
    if options.media {
        writer.add_manifest_directory("Pictures/", "").unwrap();
        writer
            .add_file_with_media_type(
                "Pictures/opaque.bin",
                &vec![0x5a; 1024 * 1024],
                "application/octet-stream",
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn content_fixture(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn model(source: &edit::Snapshot, page: usize, name: &str, text: &str) -> TextBoxModel {
    let mut value = source
        .rich_content()
        .unwrap()
        .text_boxes()
        .iter()
        .find(|candidate| candidate.page() == page && candidate.name() == name)
        .unwrap()
        .clone();
    value
        .replace_paragraph(0, &Paragraph::plain(text).unwrap())
        .unwrap();
    value
}

fn renamed_model(
    source: &edit::Snapshot,
    page: usize,
    name: &str,
    new_name: &str,
    text: &str,
) -> TextBoxModel {
    let mut value = model(source, page, name, text);
    let xml = value.xml().replace(
        &format!(r#"draw:name="{name}""#),
        &format!(r#"draw:name="{new_name}""#),
    );
    value.set_xml(xml).unwrap();
    value
}

#[test]
fn atomic_batch_updates_first_middle_last_and_cross_slide_owners_deterministically() {
    let source_bytes = fixture(
        3,
        3,
        FixtureOptions {
            media: true,
            ..FixtureOptions::default()
        },
    );
    let source = edit::Snapshot::from_bytes(source_bytes.clone()).unwrap();
    let changed = [
        model(&source, 0, "Box 0-0", "after first"),
        model(&source, 0, "Box 0-1", "after middle"),
        model(&source, 0, "Box 0-2", "after last"),
        model(&source, 1, "Box 1-1", "after cross one"),
        model(&source, 2, "Box 2-2", "after cross two"),
    ];
    let forward = [
        TextBoxModelReplacement::at(0, "Box 0-0", &changed[0]),
        TextBoxModelReplacement::at(0, "Box 0-1", &changed[1]),
        TextBoxModelReplacement::at(0, "Box 0-2", &changed[2]),
        TextBoxModelReplacement::on("Slide 1", "Box 1-1", &changed[3]),
        TextBoxModelReplacement::on("Slide 2", "Box 2-2", &changed[4]),
    ];
    let reverse = [forward[4], forward[3], forward[2], forward[1], forward[0]];

    let mut edit = source.transaction().unwrap();
    assert_eq!(edit.replace_text_box_models(&reverse).unwrap(), 5);
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.patch().domains(), &[edit::Domain::Content]);

    let mut deterministic = source.transaction().unwrap();
    assert_eq!(deterministic.replace_text_box_models(&forward).unwrap(), 5);
    let deterministic = deterministic.commit().unwrap();
    assert_eq!(deterministic.snapshot().bytes(), commit.snapshot().bytes());

    let reopened = edit::Snapshot::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let inventory = reopened.rich_content().unwrap();
    for expected in &changed {
        assert!(inventory.text_boxes().contains(expected));
    }
    let presentation = Presentation::from_bytes(reopened.bytes().to_vec()).unwrap();
    assert_eq!(presentation.slides().unwrap().len(), 3);
    assert!(
        presentation
            .content_xml()
            .contains(r#"<future:unselected future:value="byte-exact"/>"#)
    );

    let identical = raw_identical_members(&source_bytes, commit.snapshot().bytes()).unwrap();
    for member in [
        "mimetype",
        "styles.xml",
        "meta.xml",
        "META-INF/manifest.xml",
        "Pictures/opaque.bin",
    ] {
        assert!(identical.contains(member), "{member}");
    }
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    assert_eq!(
        package.get_file("Pictures/opaque.bin").unwrap(),
        vec![0x5a; 1024 * 1024]
    );

    let durable =
        edit::Patch::from_durable_bytes(&commit.patch().to_durable_bytes().unwrap()).unwrap();
    let applied = durable.apply(&source).unwrap();
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    assert_eq!(
        durable.inverse().apply(&applied).unwrap().bytes(),
        source.bytes()
    );
    let foreign = edit::Snapshot::from_bytes(fixture(1, 1, FixtureOptions::default())).unwrap();
    assert!(durable.apply(&foreign).is_err());
}

#[test]
fn duplicate_missing_collision_and_late_invalid_selectors_are_atomic() {
    let source = edit::Snapshot::from_bytes(fixture(2, 2, FixtureOptions::default())).unwrap();
    let first = model(&source, 0, "Box 0-0", "valid first");
    let second = model(&source, 0, "Box 0-1", "valid second");
    let wrong_page = model(&source, 1, "Box 1-0", "wrong owner page");

    for replacements in [
        vec![
            TextBoxModelReplacement::at(0, "Box 0-0", &first),
            TextBoxModelReplacement::at(0, "Box 0-0", &first),
        ],
        vec![
            TextBoxModelReplacement::at(0, "Box 0-0", &first),
            TextBoxModelReplacement::at(0, "missing late", &second),
        ],
        vec![
            TextBoxModelReplacement::at(0, "Box 0-0", &first),
            TextBoxModelReplacement::at(0, "Box 0-1", &wrong_page),
        ],
    ] {
        let mut transaction = source.transaction().unwrap();
        assert!(transaction.replace_text_box_models(&replacements).is_err());
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert!(commit.patch().is_noop());
        assert_eq!(commit.snapshot().bytes(), source.bytes());
    }

    let mut renamed = first.clone();
    renamed
        .set_xml(
            renamed
                .xml()
                .replace("draw:name=\"Box 0-0\"", "draw:name=\"Box 0-1\""),
        )
        .unwrap();
    let mut collision = source.transaction().unwrap();
    assert!(
        collision
            .replace_text_box_models(&[TextBoxModelReplacement::at(0, "Box 0-0", &renamed,)])
            .is_err()
    );
    assert!(!collision.commit().unwrap().changed());

    let retained_incumbent = model(&source, 0, "Box 0-1", "incumbent keeps its name");
    let mut selected_collision = source.transaction().unwrap();
    let error = selected_collision
        .replace_text_box_models(&[
            TextBoxModelReplacement::at(0, "Box 0-0", &renamed),
            TextBoxModelReplacement::at(0, "Box 0-1", &retained_incumbent),
        ])
        .unwrap_err();
    assert!(error.to_string().contains("duplicate destination names"));
    let commit = selected_collision.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source.bytes());

    let protected = edit::Snapshot::from_bytes(fixture(
        1,
        2,
        FixtureOptions {
            protected: Some((0, 1)),
            ..FixtureOptions::default()
        },
    ))
    .unwrap();
    let valid_first = model(&protected, 0, "Box 0-0", "valid changed owner");
    let refused_late = model(&protected, 0, "Box 0-1", "protected changed owner");
    let mut transaction = protected.transaction().unwrap();
    assert!(
        transaction
            .replace_text_box_models(&[
                TextBoxModelReplacement::at(0, "Box 0-0", &valid_first),
                TextBoxModelReplacement::at(0, "Box 0-1", &refused_late),
            ])
            .is_err()
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), protected.bytes());

    let protected_ancestor_xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page draw:name="Slide 0"><draw:frame draw:name="Outside"><draw:text-box><text:p>outside</text:p></draw:text-box></draw:frame><draw:g draw:name="Protected Group" draw:protect="content"><draw:frame draw:name="Inside"><draw:text-box><text:p>inside</text:p></draw:text-box></draw:frame></draw:g></draw:page></office:presentation></office:body></office:document-content>"#;
    let protected_ancestor =
        edit::Snapshot::from_bytes(content_fixture(protected_ancestor_xml)).unwrap();
    let canonical = edit::Snapshot::from_bytes(fixture(1, 1, FixtureOptions::default())).unwrap();
    let valid_first = renamed_model(&canonical, 0, "Box 0-0", "Outside", "outside changed first");
    let refused_late = renamed_model(&canonical, 0, "Box 0-0", "Inside", "inside refused late");
    let mut transaction = protected_ancestor.transaction().unwrap();
    let error = transaction
        .replace_text_box_models(&[
            TextBoxModelReplacement::at(0, "Outside", &valid_first),
            TextBoxModelReplacement::at(0, "Inside", &refused_late),
        ])
        .unwrap_err();
    assert!(error.to_string().contains("protected drawing owner"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().bytes(), protected_ancestor.bytes());
}

#[test]
fn rename_cycles_are_published_simultaneously_and_read_back_by_destination() {
    let source = edit::Snapshot::from_bytes(fixture(1, 3, FixtureOptions::default())).unwrap();
    let cycle = [
        renamed_model(&source, 0, "Box 0-0", "Box 0-1", "zero moved to one"),
        renamed_model(&source, 0, "Box 0-1", "Box 0-2", "one moved to two"),
        renamed_model(&source, 0, "Box 0-2", "Box 0-0", "two moved to zero"),
    ];
    let replacements = [
        TextBoxModelReplacement::at(0, "Box 0-0", &cycle[0]),
        TextBoxModelReplacement::at(0, "Box 0-1", &cycle[1]),
        TextBoxModelReplacement::at(0, "Box 0-2", &cycle[2]),
    ];

    let mut transaction = source.transaction().unwrap();
    assert_eq!(
        transaction.replace_text_box_models(&replacements).unwrap(),
        3
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let reopened = edit::Snapshot::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let actual = reopened.rich_content().unwrap();
    for expected in &cycle {
        assert!(actual.text_boxes().contains(expected));
    }
}

#[test]
fn ambiguous_page_names_and_overlapping_named_owners_fail_atomically() {
    let ambiguous_xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page draw:name="Duplicate"><draw:frame draw:name="First"><draw:text-box><text:p>first</text:p></draw:text-box></draw:frame></draw:page><draw:page draw:name="Duplicate"><draw:frame draw:name="Second"><draw:text-box><text:p>second</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;
    let ambiguous = edit::Snapshot::from_bytes(content_fixture(ambiguous_xml)).unwrap();
    let changed = model(&ambiguous, 0, "First", "changed");
    let mut transaction = ambiguous.transaction().unwrap();
    let error = transaction
        .replace_text_box_models(&[TextBoxModelReplacement::on("Duplicate", "First", &changed)])
        .unwrap_err();
    assert!(error.to_string().contains("page name is ambiguous"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), ambiguous.bytes());

    let canonical = edit::Snapshot::from_bytes(fixture(1, 2, FixtureOptions::default())).unwrap();
    let outer = renamed_model(&canonical, 0, "Box 0-0", "Outer", "outer replacement");
    let inner = renamed_model(&canonical, 0, "Box 0-1", "Inner", "inner replacement");
    let overlapping_xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page draw:name="Slide 0"><draw:g draw:name="Outer"><draw:frame draw:name="Inner"><draw:text-box><text:p>inner</text:p></draw:text-box></draw:frame></draw:g></draw:page></office:presentation></office:body></office:document-content>"#;
    let overlapping = edit::Snapshot::from_bytes(content_fixture(overlapping_xml)).unwrap();
    let mut transaction = overlapping.transaction().unwrap();
    let error = transaction
        .replace_text_box_models(&[
            TextBoxModelReplacement::at(0, "Inner", &inner),
            TextBoxModelReplacement::at(0, "Outer", &outer),
        ])
        .unwrap_err();
    assert!(error.to_string().contains("selections overlap"));
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), overlapping.bytes());
}

#[test]
fn exact_batch_limit_mixed_noops_and_all_noop_sharing_are_checked() {
    let source = edit::Snapshot::from_bytes(fixture(
        1,
        MAX_TEXT_BOX_MODEL_REPLACEMENTS,
        FixtureOptions::default(),
    ))
    .unwrap();
    let inventory = source.rich_content().unwrap();
    let mut models = inventory.text_boxes().to_vec();
    models
        .last_mut()
        .unwrap()
        .replace_paragraph(0, &Paragraph::plain("only changed owner").unwrap())
        .unwrap();
    let replacements = models
        .iter()
        .map(|model| TextBoxModelReplacement::at(0, model.name(), model))
        .collect::<Vec<_>>();
    let mut transaction = source.transaction().unwrap();
    assert_eq!(
        transaction.replace_text_box_models(&replacements).unwrap(),
        1
    );
    assert!(transaction.commit().unwrap().changed());

    let noop_models = inventory.text_boxes().to_vec();
    let noop = noop_models
        .iter()
        .map(|model| TextBoxModelReplacement::at(0, model.name(), model))
        .collect::<Vec<_>>();
    let mut transaction = source.transaction().unwrap();
    assert_eq!(transaction.replace_text_box_models(&noop).unwrap(), 0);
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes().as_ptr(), source.bytes().as_ptr());

    let above = vec![noop[0]; MAX_TEXT_BOX_MODEL_REPLACEMENTS + 1];
    let mut transaction = source.transaction().unwrap();
    assert!(transaction.replace_text_box_models(&above).is_err());
    assert!(!transaction.commit().unwrap().changed());
}

#[test]
fn protected_and_opaque_selected_owners_allow_noops_but_refuse_changes() {
    for options in [
        FixtureOptions {
            protected: Some((0, 0)),
            ..FixtureOptions::default()
        },
        FixtureOptions {
            opaque: Some((0, 0)),
            ..FixtureOptions::default()
        },
    ] {
        let source = edit::Snapshot::from_bytes(fixture(1, 1, options)).unwrap();
        let original = source.rich_content().unwrap().text_boxes()[0].clone();
        let mut transaction = source.transaction().unwrap();
        assert_eq!(
            transaction
                .replace_text_box_models(&[TextBoxModelReplacement::at(0, "Box 0-0", &original,)])
                .unwrap(),
            0
        );
        assert!(!transaction.commit().unwrap().changed());

        let changed = model(&source, 0, "Box 0-0", "forbidden change");
        let mut transaction = source.transaction().unwrap();
        assert!(
            transaction
                .replace_text_box_models(&[TextBoxModelReplacement::at(0, "Box 0-0", &changed,)])
                .is_err()
        );
        assert!(!transaction.commit().unwrap().changed());
    }

    let active_xml = fixture(
        1,
        1,
        FixtureOptions {
            processing_instruction: Some((0, 0)),
            ..FixtureOptions::default()
        },
    );
    let active = edit::Snapshot::from_bytes(active_xml).unwrap();
    assert!(active.transaction().is_err());
}

#[test]
fn signed_encrypted_and_malformed_sources_fail_closed() {
    let mut signed = PackageWriter::new();
    signed.set_mimetype(MIME).unwrap();
    signed
        .add_file(
            "content.xml",
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
        )
        .unwrap();
    signed
        .add_file(
            "META-INF/documentsignatures.xml",
            br#"<dsig:document-signatures xmlns:dsig="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
        )
        .unwrap();
    let signed = edit::Snapshot::from_bytes(signed.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(
        signed.security_policy().unwrap(),
        edit::SecurityPolicy::SignedReadOnly
    );
    assert!(signed.transaction().is_err());

    const ENCRYPTED_MANIFEST: &[u8] = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.presentation"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="secret.bin" m:media-type="application/octet-stream" m:size="1"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="AAAAAAAAAAAAAAAA"/><m:start-key-generation m:start-key-generation-name="SHA1" m:key-size="20"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="AQ==" m:iteration-count="1000" m:key-size="32"/></m:encryption-data></m:file-entry></m:manifest>"#;
    let mut encrypted = StreamingArchiveWriter::new();
    encrypted.write_stored("mimetype", MIME.as_bytes()).unwrap();
    encrypted
        .write_deflated(
            "content.xml",
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
        )
        .unwrap();
    encrypted.write_deflated("secret.bin", b"x").unwrap();
    encrypted
        .write_deflated("META-INF/manifest.xml", ENCRYPTED_MANIFEST)
        .unwrap();
    let encrypted = edit::Snapshot::from_bytes(encrypted.finish_to_bytes().unwrap()).unwrap();
    assert_eq!(
        encrypted.security_policy().unwrap(),
        edit::SecurityPolicy::EncryptedReadOnly
    );
    assert!(encrypted.transaction().is_err());

    let malformed = fixture(1, 1, FixtureOptions::default());
    let package = OwnedPackage::from_bytes(malformed).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    let malformed_content = format!("<!DOCTYPE office:document-content>{content}");
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    assert!(
        writer
            .add_file("content.xml", malformed_content.as_bytes())
            .is_err()
    );
}
