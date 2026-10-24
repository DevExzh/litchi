#![cfg(any(unix, windows))]
//! Record 0011: the markup-compatibility admission of DOCX document
//! transactions, on both owners and under an execution context.
//!
//! A transaction edits and publishes the main part's own bytes while readers
//! see the markup-compatibility projection. Record 0008 found every Word
//! fixture refused because the part merely mentions the namespace; 0011
//! admits the parts whose projection hides nothing (stages 1 and 2), applies
//! the same rule to the owning package and the managed route (stage 3), and
//! keeps every real hazard refused.

use std::io::Cursor;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Position,
    Resource,
};
use litchi_docx::document::{Edit, Error as TransactionError, raw};
use litchi_docx::{Error, Package, SourceLimits, source_backed};
use litchi_ooxml_common::mce::admission::{Admission, admit};
use litchi_ooxml_common::mce::{self, Capabilities};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const MAIN: &str = "word/document.xml";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MCE_ESCAPED: &str = "http://schemas.openxmlformats.org/markup-compatibility/200&#54;";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const W16DU: &str = "http://schemas.microsoft.com/office/word/2023/wordml/word16du";
const WP14: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

/// Record 0008 §1: the fixtures that declare the namespace and use nothing
/// in it (stage 1).
const DECLARED_ONLY: [&str; 13] = [
    "FancyFoot",
    "Headers",
    "documentProperties",
    "documentProtection_comments_no_password",
    "documentProtection_forms_no_password",
    "documentProtection_no_protection",
    "documentProtection_readonly_no_password",
    "drawing",
    "endnotes",
    "footnotes",
    "headerFooter",
    "headerPic",
    "table_footnotes",
];

/// The fixtures whose only construct is a root `mc:Ignorable` (stage 2).
const IGNORABLE: [&str; 6] = [
    "ComplexNumberedLists",
    "Hyperlink",
    "comment",
    "table-alignment",
    "table-indent",
    "testComment",
];

/// The fixtures that enforce comments, forms or read-only protection.
const PROTECTED: [&str; 3] = [
    "documentProtection_comments_no_password",
    "documentProtection_forms_no_password",
    "documentProtection_readonly_no_password",
];

fn all_fixtures() -> impl Iterator<Item = &'static str> {
    DECLARED_ONLY.into_iter().chain(IGNORABLE)
}

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(format!(
        "{}/../../test-data/ooxml/docx/{name}.docx",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap_or_else(|error| panic!("{name}: {error}"))
}

/// The profile the main document is read with: the OOXML baseline, `w14`
/// and `w16du`.
fn word_profile() -> Capabilities {
    let mut capabilities = Capabilities::default();
    capabilities.understand_namespace(W14);
    capabilities.understand_namespace(W16DU);
    capabilities
}

/// One member as written: name, payload, compressed size and method.
#[derive(Debug, PartialEq, Eq)]
struct Member {
    name: String,
    bytes: Vec<u8>,
    compressed: u64,
    stored: bool,
}

fn members(archive: &[u8]) -> Vec<Member> {
    let reader =
        ArchiveReader::new(archive, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let names: Vec<String> = reader.file_names().map(str::to_owned).collect();
    names
        .into_iter()
        .map(|name| Member {
            bytes: reader.read(&name).unwrap(),
            compressed: reader.metadata(&name).unwrap().compressed_size(),
            stored: reader.is_stored(&name).unwrap(),
            name,
        })
        .collect()
}

fn main_part(archive: &[u8]) -> Vec<u8> {
    ArchiveReader::new(archive, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .read(MAIN)
        .unwrap()
}

fn source(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn eager(bytes: &[u8]) -> Package {
    Package::from_reader(Cursor::new(bytes)).unwrap()
}

fn managed(bytes: &[u8]) -> (Budget, CancellationSource, source_backed::Package) {
    let memory = 256 * 1024 * 1024;
    let budget = Budget::root(
        "docx-mce-admission",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let package = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes.to_vec())),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(ExecutionContext::new(budget.clone(), cancellation, limits)),
    )
    .unwrap();
    (budget, cancellation_source, package)
}

/// The edit record 0008's probe made: replace the text of the first
/// paragraph the grammar accepts, or append one paragraph when none is.
fn first_edit<P: Clone>(edit: &mut Edit<P>) {
    let count = edit.source().unwrap().paragraph_count();
    for position in 0..count {
        if edit
            .replace_paragraph_text(Position::new(position), "UO-0011 replaced")
            .is_ok()
        {
            return;
        }
    }
    edit.insert_paragraph(Position::new(count), "UO-0011 appended")
        .unwrap();
}
fn is_protection_refusal(error: &TransactionError) -> bool {
    matches!(
            error,
            TransactionError::Document(Error::Refused(refusal)) if refusal.operation() == "changed document publication" && refusal.reason() == "document or write protection, or tracked revisions, is enabled"
    )
}

/// The root start tag of a part, as written.
fn root_tag(xml: &[u8]) -> &[u8] {
    let mut cursor = 0;
    loop {
        let start = cursor + xml[cursor..].iter().position(|byte| *byte == b'<').unwrap();
        let end = start + xml[start..].iter().position(|byte| *byte == b'>').unwrap();
        if !matches!(xml.get(start + 1), Some(b'?' | b'!')) {
            return &xml[start..=end];
        }
        cursor = end + 1;
    }
}

fn count(haystack: &[u8], needle: &[u8]) -> usize {
    haystack
        .windows(needle.len())
        .filter(|window| *window == needle)
        .count()
}

#[test]
fn every_word_fixture_falls_into_the_stage_record_0008_names() {
    let mut declared = Vec::new();
    let mut ignorable = Vec::new();
    for name in all_fixtures() {
        let raw = main_part(&fixture(name));
        assert!(mce::mentions_namespace(&raw), "{name}");
        match admit(&raw, &word_profile(), &mce::Limits::default()) {
            Ok(Admission::DeclaredOnly) => declared.push(name),
            Ok(Admission::Ignorable) => ignorable.push(name),
            other => panic!("{name}: {other:?}"),
        }
    }
    assert_eq!(declared, DECLARED_ONLY, "stage 1");
    assert_eq!(ignorable, IGNORABLE, "stage 2");
}

#[test]
fn source_backed_edits_publish_what_the_eager_owner_publishes() {
    for name in all_fixtures() {
        let original = fixture(name);
        let protected = PROTECTED.contains(&name);

        let package = source(&original);
        let mut edit = package
            .edit()
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        first_edit(&mut edit);
        let mut published = Vec::new();
        let source_result = edit.commit().and_then(|commit| {
            package.apply(commit.patch()).and_then(|committed| {
                let committed = committed.into_package();
                committed.write_to(&mut published)?;
                committed.document_snapshot()
            })
        });

        let owner = eager(&original);
        let mut edit = owner
            .edit()
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        first_edit(&mut edit);
        let eager_result = edit.commit();

        if protected {
            // One protection policy, whichever owner holds the document.
            assert!(
                source_result.as_ref().is_err_and(is_protection_refusal),
                "{name}: {source_result:?}"
            );
            assert!(published.is_empty(), "{name}: refused before output");
            assert!(
                eager_result.as_ref().is_err_and(is_protection_refusal),
                "{name}: {:?}",
                eager_result.as_ref().err()
            );
            assert_eq!(
                format!("{:?}", source_result.unwrap_err()),
                format!("{:?}", eager_result.unwrap_err()),
                "{name}: identical refusals"
            );
            continue;
        }
        source_result.unwrap_or_else(|error| panic!("{name}: {error}"));
        let owner = eager_result
            .unwrap_or_else(|error| panic!("{name}: {error}"))
            .into_package();
        let mut saved = Vec::new();
        owner.write_to(&mut saved).unwrap();

        // The two owners publish the same main part, byte for byte.
        let source_main = main_part(&published);
        assert_eq!(source_main, main_part(&saved), "{name}: eager parity");

        // Every markup-compatibility construct survives byte for byte.
        let raw = main_part(&original);
        assert_eq!(root_tag(&source_main), root_tag(&raw), "{name}: root tag");
        for token in [&b"mc:"[..], b"ve:", b"Ignorable", b"markup-compatibility"] {
            assert_eq!(
                count(&source_main, token),
                count(&raw, token),
                "{name}: {token:?}"
            );
        }

        // Every other member is copied as it was written.
        let before = members(&original);
        let after = members(&published);
        assert_eq!(before.len(), after.len(), "{name}");
        for (old, new) in before.iter().zip(&after) {
            assert_eq!(old.name, new.name, "{name}: member order");
            if old.name != MAIN {
                assert_eq!(old, new, "{name}: {} preserved", old.name);
            }
        }
    }
}

#[test]
fn exact_no_ops_republish_every_fixture_byte_for_byte() {
    for name in all_fixtures() {
        let original = fixture(name);
        let package = source(&original);
        let commit = package.edit().unwrap().commit().unwrap();
        assert!(commit.patch().is_noop());
        let mut published = Vec::new();
        package
            .apply(commit.patch())
            .and_then(|committed| {
                let committed = committed.into_package();
                committed.write_to(&mut published)?;
                committed.document_snapshot()
            })
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(published, original, "{name}");

        let owner = eager(&original);
        let before = owner.document_snapshot().unwrap();
        let commit = owner.edit().unwrap().commit().unwrap();
        let after = commit;
        assert_eq!(
            after.patch().source().xml_bytes(),
            before.xml_bytes(),
            "{name}"
        );
    }
}

/// The projection oracle: the transaction scan over the source bytes and
/// over the projection agree on layout and every paragraph's text.
#[test]
fn raw_and_processed_main_parts_agree_on_what_transactions_read() {
    for name in all_fixtures() {
        let raw = main_part(&fixture(name));
        let processed =
            mce::process_markup_compatibility(&raw, &word_profile(), &mce::Limits::default())
                .unwrap();
        assert!(
            matches!(processed.xml, std::borrow::Cow::Owned(_)),
            "{name}"
        );
        let source = raw::Tree::from_xml(raw.clone(), raw::Limits::default()).unwrap();
        let visible =
            raw::Tree::from_xml(processed.xml.into_owned(), raw::Limits::default()).unwrap();
        let source = source.projected();
        let visible = visible.projected();
        assert_eq!(
            source.paragraph_count(),
            visible.paragraph_count(),
            "{name}"
        );
        assert_eq!(source.table_count(), visible.table_count(), "{name}");
        assert_eq!(
            source.block_content_control_count(),
            visible.block_content_control_count(),
            "{name}"
        );
        for (left, right) in source.paragraphs().iter().zip(visible.paragraphs()) {
            assert_eq!(left.text().unwrap(), right.text().unwrap(), "{name}");
        }
    }
}

/// The text the unmanaged reader (projection) and the managed reader
/// (source bytes) extract.
fn reader_texts(bytes: &[u8]) -> (String, String) {
    let unmanaged = source(bytes).document().unwrap().extract_text().unwrap();
    let (budget, _cancellation, package) = managed(bytes);
    let managed = package.document().unwrap().extract_text().unwrap();
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    (unmanaged, managed)
}

#[test]
fn document_variable_transactions_admit_every_fixture() {
    for name in all_fixtures() {
        let package = source(&fixture(name));
        let snapshot = package
            .document_variables_snapshot()
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        assert!(
            litchi_docx::variables::raw::Tree::from_xml(
                snapshot.xml_bytes().to_vec(),
                litchi_docx::variables::raw::Limits::default(),
            )
            .is_ok(),
            "{name}"
        );
    }
}

#[test]
fn managed_route_reads_and_edits_every_fixture() {
    for name in all_fixtures() {
        let original = fixture(name);
        let (unmanaged_text, managed_text) = reader_texts(&original);
        assert_eq!(unmanaged_text, managed_text, "{name}");

        let (budget, _cancellation, package) = managed(&original);
        let snapshot = package
            .document_snapshot()
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(
            snapshot.paragraph_count(),
            source(&original)
                .document_snapshot()
                .unwrap()
                .paragraph_count(),
            "{name}"
        );
        assert!(
            budget.used(Resource::Memory) > 0,
            "{name}: the snapshot is charged"
        );
        drop(snapshot);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0, "{name}: released");
    }
}

/// A one-part package around `document`, with an optional settings part.
fn package(document: &str, settings: Option<&str>) -> Vec<u8> {
    let settings_override = settings.map_or(String::new(), |_| {
        format!(
            r#"<Override PartName="/word/settings.xml" ContentType="{}"/>"#,
            ct::WML_SETTINGS
        )
    });
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/{MAIN}" ContentType="{}"/>{settings_override}</Types>"#,
        ct::WML_DOCUMENT_MAIN
    );
    let relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{}" Target="{MAIN}"/></Relationships>"#,
        rt::OFFICE_DOCUMENT
    );
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .unwrap();
    writer.write_stored(MAIN, document.as_bytes()).unwrap();
    if let Some(settings) = settings {
        let main_relationships = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{}" Target="settings.xml"/></Relationships>"#,
            rt::SETTINGS
        );
        writer
            .write_stored(
                "word/_rels/document.xml.rels",
                main_relationships.as_bytes(),
            )
            .unwrap();
        writer
            .write_stored("word/settings.xml", settings.as_bytes())
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn document(declarations: &str, body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:w14="{W14}" xmlns:x="urn:x" {declarations}><w:body>{body}<w:sectPr/></w:body></w:document>"#
    )
}

const PARAGRAPHS: &str = r#"<w:p w14:paraId="1A2B3C4D"><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>"#;

/// What each route must do with one witness.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Expect {
    /// Both owners and the managed route admit it.
    Admitted,
    /// Both owners admit it, and so does a managed read. A managed
    /// transaction captures the part through the OPC source-XML seam, which
    /// refuses a namespace declaration whose written value differs from its
    /// decoded value, whatever the namespace.
    AdmittedExceptManagedCapture,
    /// Both owners admit it; the managed route, which has no projection to
    /// hide the dropped markup, refuses it.
    UnmanagedOnly,
    /// Every route refuses it as an unsafe edit.
    Refused,
    /// Both owners return the projection's `MustUnderstand` error; the
    /// managed route refuses it as an unsafe edit.
    MustUnderstand,
}

#[test]
fn witness_matrix_on_both_owners_and_the_managed_route() {
    let mc = format!(r#"xmlns:mc="{MCE}""#);
    let escaped = format!(r#"xmlns:mc="{MCE_ESCAPED}""#);
    let witnesses: Vec<(&str, String, Expect)> = vec![
        (
            "declared only (Word 2007 `ve`)",
            document(&format!(r#"xmlns:ve="{MCE}""#), PARAGRAPHS),
            Expect::Admitted,
        ),
        (
            "root Ignorable, understood and unused prefixes",
            document(
                &format!(r#"{mc} xmlns:wp14="{WP14}" mc:Ignorable="w14 wp14""#),
                PARAGRAPHS,
            ),
            Expect::Admitted,
        ),
        (
            "escaped namespace, declared only",
            document(&escaped, PARAGRAPHS),
            Expect::AdmittedExceptManagedCapture,
        ),
        (
            "nested Ignorable of an understood namespace (revision writer)",
            document(
                &format!(r#"xmlns:w16du="{W16DU}""#),
                &format!(
                    r#"<w:p><w:ins w:id="1" w:author="A" xmlns:mc="{MCE}" mc:Ignorable="w16du" w16du:dateUtc="2026-01-01T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>{PARAGRAPHS}"#
                ),
            ),
            Expect::Admitted,
        ),
        (
            "W6: ignorable attribute and leaf element",
            document(
                &format!(r#"{mc} mc:Ignorable="x""#),
                r#"<w:p x:id="1"><w:r><x:flag/><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>"#,
            ),
            Expect::UnmanagedOnly,
        ),
        (
            // Record 0307: run-level alternate content is an opaque inline
            // object; the projection keeps every block where it is.
            "W1: AlternateContent in a run",
            document(
                &mc,
                &format!(
                    r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>BOX</w:t></mc:Choice><mc:Fallback><w:t>BOX</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>{PARAGRAPHS}"#
                ),
            ),
            Expect::UnmanagedOnly,
        ),
        (
            "AlternateContent in a paragraph, outside its runs",
            document(
                &mc,
                r#"<w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>BOX</w:t></w:r></mc:Choice></mc:AlternateContent></w:p><w:p/>"#,
            ),
            Expect::Refused,
        ),
        (
            "W2: AlternateContent in the body",
            document(
                &mc,
                r#"<w:p><w:r><w:t>P0</w:t></w:r></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice><mc:Fallback/></mc:AlternateContent><w:p><w:r><w:t>P2</w:t></w:r></w:p>"#,
            ),
            Expect::Refused,
        ),
        (
            "W7: AlternateContent through an escaped namespace",
            document(
                &escaped,
                r#"<w:p><w:r><w:t>P0</w:t></w:r></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice><mc:Fallback/></mc:AlternateContent><w:p><w:r><w:t>P2</w:t></w:r></w:p>"#,
            ),
            Expect::Refused,
        ),
        (
            "W3: ProcessContent",
            document(
                &format!(r#"{mc} mc:Ignorable="x" mc:ProcessContent="x:wrap""#),
                r#"<w:p><x:wrap><w:r><w:t>WRAPPED</w:t></w:r></x:wrap><w:r><w:t>B</w:t></w:r></w:p>"#,
            ),
            Expect::Refused,
        ),
        (
            "W4: used ignorable prefix holding a run",
            document(
                &format!(r#"{mc} mc:Ignorable="x""#),
                r#"<w:p><x:ext><w:r><w:t>HIDDEN</w:t></w:r></x:ext><w:r><w:t>B</w:t></w:r></w:p>"#,
            ),
            Expect::Refused,
        ),
        (
            "nested Ignorable of a foreign namespace",
            document(
                &mc,
                r#"<w:p mc:Ignorable="x"><w:r><w:t>first</w:t></w:r></w:p>"#,
            ),
            Expect::Refused,
        ),
        (
            "W5: MustUnderstand of an unknown namespace",
            document(&format!(r#"{mc} mc:MustUnderstand="x""#), PARAGRAPHS),
            Expect::MustUnderstand,
        ),
    ];

    for (label, xml, expect) in witnesses {
        let bytes = package(&xml, None);
        let source_owner = source(&bytes);
        let source_edit = source_owner.edit();
        let mut owner = eager(&bytes);
        let eager_edit = owner.edit();
        match expect {
            Expect::Admitted | Expect::AdmittedExceptManagedCapture | Expect::UnmanagedOnly => {
                let mut source_edit =
                    source_edit.unwrap_or_else(|error| panic!("{label}: {error}"));
                let mut eager_edit = eager_edit.unwrap_or_else(|error| panic!("{label}: {error}"));
                // Both owners edit the same bytes and publish the same part.
                source_edit
                    .replace_paragraph_text(Position::new(1), "edited")
                    .unwrap_or_else(|error| panic!("{label}: {error}"));
                eager_edit
                    .replace_paragraph_text(Position::new(1), "edited")
                    .unwrap();
                let commit = source_edit
                    .commit()
                    .unwrap_or_else(|error| panic!("{label}: {error}"));
                let mut published = Vec::new();
                commit.package().write_to(&mut published).unwrap();
                owner = eager_edit.commit().unwrap().into_package();
                let mut saved = Vec::new();
                owner.write_to(&mut saved).unwrap();
                assert_eq!(main_part(&published), main_part(&saved), "{label}");
                assert_eq!(
                    root_tag(&main_part(&published)),
                    root_tag(xml.as_bytes()),
                    "{label}"
                );
                // The reader sees the edit in the projection.
                let text = source(&published)
                    .document()
                    .unwrap()
                    .extract_text()
                    .unwrap();
                assert!(text.contains("edited"), "{label}: {text}");
            },
            Expect::Refused => {
                let source_edit = source_edit.and_then(|mut edit| {
                    edit.replace_paragraph_text(Position::new(1), "edited")?;
                    edit.commit().map(|_| ())
                });
                let eager_edit = eager_edit.and_then(|mut edit| {
                    edit.replace_paragraph_text(Position::new(1), "edited")?;
                    edit.commit().map(|_| ())
                });
                assert!(
                    matches!(
                                            source_edit,
                                            Err(TransactionError::Document(Error::Refused(refusal))) if refusal.operation() == "edit" && refusal.reason() == "source-backed document transactions do not support markup-compatibility branch selection"
                    ),
                    "{label}: {source_edit:?}"
                );
                assert!(
                    matches!(
                                            eager_edit,
                                            Err(TransactionError::Document(Error::Refused(refusal))) if refusal.operation() == "document_snapshot" && refusal.reason() == "document transactions do not support markup-compatibility branch selection"
                    ),
                    "{label}: {:?}",
                    eager_edit.as_ref().err()
                );
            },
            Expect::MustUnderstand => {
                let source_result = source_edit.and_then(|mut edit| {
                    edit.replace_paragraph_text(Position::new(1), "edited")?;
                    edit.commit().map(drop)
                });
                let eager_result = eager_edit.and_then(|mut edit| {
                    edit.replace_paragraph_text(Position::new(1), "edited")?;
                    edit.commit().map(drop)
                });
                for result in [source_result, eager_result] {
                    assert!(
                        matches!(
                            result,
                            Err(TransactionError::Document(Error::Mce(
                                mce::Error::MustUnderstand(ref namespace),
                            ))) if namespace == "urn:x"
                        ),
                        "{label}: {result:?}"
                    );
                }
            },
        }

        // The managed route reads source bytes without a projection.
        let (budget, _cancellation, package) = managed(&bytes);
        let read = package.document().map(drop);
        let snapshot = package.document_snapshot().map(drop);
        let literal_binding_refusal = |result: &Result<(), TransactionError>| {
            matches!(
                result,
                Err(TransactionError::Document(Error::Opc(
                    litchi_opc::OpcError::SourceBackedOverlayUnavailable { reason },
                ))) if reason == "source XML requires literal namespace bindings"
            )
        };
        if expect == Expect::Admitted {
            read.unwrap_or_else(|error| panic!("{label}: {error}"));
            snapshot.unwrap_or_else(|error| panic!("{label}: {error}"));
        } else if expect == Expect::AdmittedExceptManagedCapture {
            read.unwrap_or_else(|error| panic!("{label}: {error}"));
            assert!(literal_binding_refusal(&snapshot), "{label}: {snapshot:?}");
        } else {
            assert!(
                matches!(
                                    read,
                                    Err(Error::Refused(refusal)) if refusal.operation() == "source-backed document read" && refusal.reason() == "markup-compatibility preprocessing would require an unbudgeted owned payload"
                ),
                "{label}: {read:?}"
            );
            // A binding spelled with references stops the managed capture
            // at the OPC seam before the admission runs.
            assert!(
                matches!(
                                    snapshot,
                                    Err(TransactionError::Document(Error::Refused(refusal))) if refusal.operation() == "source-backed document read"
                ) || (xml.contains(MCE_ESCAPED) && literal_binding_refusal(&snapshot)),
                "{label}: {snapshot:?}"
            );
        }
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0, "{label}");
    }
}

#[test]
fn managed_admission_is_charged_to_the_context() {
    // Equal lengths: the namespace and an unknown URI of the same size, so
    // the managed scan charges the same work for both and only the proof
    // differs.
    let unknown = "http://schemas.openxmlformats.org/markup-compatibility/2999";
    assert_eq!(unknown.len(), MCE.len());
    let with = document(
        &format!(r#"xmlns:mc="{MCE}" mc:Ignorable="w14""#),
        PARAGRAPHS,
    );
    let without = with.replace(MCE, unknown);
    let mut work = Vec::new();
    for xml in [&without, &with] {
        let (budget, _cancellation, package) = managed(&package(xml, None));
        package.document_snapshot().unwrap();
        drop(package);
        work.push(budget.used(Resource::Work));
        assert_eq!(budget.used(Resource::Memory), 0);
    }
    assert!(
        work[1] >= work[0] + u64::try_from(with.len()).unwrap(),
        "the proof charges at least one unit per byte: {work:?}"
    );

    // A cancelled context stops before any verdict, with its own identity.
    let (_budget, cancellation, package) = managed(&package(&with, None));
    cancellation.cancel();
    assert!(matches!(
        package.document_snapshot(),
        Err(TransactionError::Document(Error::Opc(
            litchi_opc::OpcError::Cancelled
        )))
    ));
}

#[test]
fn document_variables_follow_the_same_rule() {
    let settings = |root_attributes: &str, body: &str| {
        format!(
            r#"<w:settings xmlns:w="{W}" xmlns:w14="{W14}" xmlns:w15="urn:w15" {root_attributes}><w:docVars><w:docVar w:name="kept" w:val="1"/></w:docVars>{body}</w:settings>"#
        )
    };
    let main = document("", PARAGRAPHS);
    for (label, xml, admitted) in [
        (
            "Word 2013 leaves under a root Ignorable",
            settings(
                &format!(r#"xmlns:mc="{MCE}" mc:Ignorable="w14 w15""#),
                r#"<w14:docId w14:val="7F00AA10"/><w15:chartTrackingRefBased/>"#,
            ),
            true,
        ),
        (
            "variables hidden in an ignorable wrapper",
            settings(
                &format!(r#"xmlns:mc="{MCE}" mc:Ignorable="w15""#),
                r#"<w15:wrap><w:docVars/></w15:wrap>"#,
            ),
            false,
        ),
        (
            "variables in an alternative branch",
            settings(
                &format!(r#"xmlns:mc="{MCE}""#),
                r#"<mc:AlternateContent><mc:Choice Requires="w14"><w:defaultTabStop w:val="720"/></mc:Choice></mc:AlternateContent>"#,
            ),
            false,
        ),
    ] {
        let package = source(&package(&main, Some(&xml)));
        let snapshot = package.document_variables_snapshot();
        if admitted {
            assert_eq!(
                snapshot
                    .unwrap_or_else(|error| panic!("{label}: {error}"))
                    .variables()
                    .count(),
                1,
                "{label}"
            );
        } else {
            assert!(
                matches!(
                                    snapshot,
                                    Err(Error::Refused(refusal)) if refusal.operation() == "document_variables_snapshot" && refusal.reason() == "source-backed document-variable transactions do not support markup-compatibility branch selection"
                ),
                "{label}: {:?}",
                snapshot.as_ref().err()
            );
        }
    }
}

/// Scanners that recognize the markup-compatibility namespace on namespace
/// values as written cannot see a binding spelled with references, which
/// preprocessing now applies (record 0011). Each refuses such a part, or
/// reports it, instead of acting on bytes readers see differently.
#[test]
fn raw_namespace_scanners_refuse_a_binding_spelled_with_references() {
    let escaped = document(
        &format!(r#"xmlns:mc="{MCE_ESCAPED}" mc:Ignorable="w14""#),
        PARAGRAPHS,
    );
    let literal = document(&format!(r#"xmlns:mc="{MCE}""#), PARAGRAPHS);

    // Tracked-revision authoring.
    assert!(
        litchi_docx::revision::authoring::Snapshot::from_xml(escaped.clone(), Default::default())
            .is_err()
    );
    assert!(
        litchi_docx::revision::authoring::Snapshot::from_xml(literal.clone(), Default::default())
            .is_ok()
    );

    // External-hyperlink redaction, past the transaction admission (a root
    // `Ignorable` of an understood namespace is admitted).
    let error = source(&package(&escaped, None))
        .external_hyperlink_redaction_snapshot()
        .unwrap_err();
    assert!(
        matches!(
                    error,
                    Error::Refused(refusal) if refusal.operation() == "external_hyperlink_redaction" && refusal.reason() == "markup-compatibility owner syntax is outside the redaction closure"
        ),
        "{error:?}"
    );
    assert!(
        source(&package(&literal, None))
            .external_hyperlink_redaction_snapshot()
            .is_ok()
    );

    // Story-hyperlink diagnostics report the namespace as present.
    let declared = document(&format!(r#"xmlns:mc="{MCE_ESCAPED}""#), PARAGRAPHS);
    let snapshot = source(&package(&declared, None))
        .story_hyperlinks_only_snapshot()
        .unwrap();
    assert!(snapshot.diagnostics().iter().any(|diagnostic| matches!(
        diagnostic.class(),
        litchi_docx::story_hyperlinks::UnsupportedClass::MarkupCompatibility
    )));

    // The ink inventory classifies frames on values as written while active
    // anchors come from the processor.
    assert!(source(&package(&escaped, None)).ink().is_err());
    assert!(source(&package(&literal, None)).ink().is_ok());
}

/// A transfer moves donor bytes into another document. Markup the donor's
/// root `mc:Ignorable` drops is ignorable only there, so the owning package
/// transfers only from a donor whose projection drops nothing.
#[test]
fn paragraph_transfer_requires_a_donor_whose_projection_drops_nothing() {
    let receiver = eager(&package(&document("", PARAGRAPHS), None));
    let opaque = eager(&package(
        &document(
            &format!(r#"xmlns:mc="{MCE}" mc:Ignorable="x""#),
            r#"<w:p x:id="1"><w:r><x:flag/><w:t>moved</w:t></w:r></w:p>"#,
        ),
        None,
    ));
    let error = receiver
        .plan_paragraph_transfer_from(&opaque, Position::new(0))
        .unwrap_err();
    assert!(
        matches!(
            error,
            TransactionError::Transfer(
                litchi_docx::document::TransferRefusal::MarkupCompatibilityScope
            )
        ),
        "{error:?}"
    );

    // Record 0307: a donor admitted for run-level alternate content selects
    // a branch, which a receiver need not understand alike; even a paragraph
    // without a wrapper is refused, as for the opaque donor above.
    let alternate = eager(&package(
        &document(
            &format!(r#"xmlns:mc="{MCE}""#),
            &format!(
                r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>BOX</w:t></mc:Choice></mc:AlternateContent></w:r></w:p>{PARAGRAPHS}"#
            ),
        ),
        None,
    ));
    alternate.edit().unwrap();
    for paragraph in [0, 1] {
        let error = receiver
            .plan_paragraph_transfer_from(&alternate, Position::new(paragraph))
            .unwrap_err();
        assert!(
            matches!(
                error,
                TransactionError::Transfer(
                    litchi_docx::document::TransferRefusal::MarkupCompatibilityScope
                )
            ),
            "{error:?}"
        );
    }

    let lexical = eager(&package(
        &document(
            &format!(r#"xmlns:mc="{MCE}" mc:Ignorable="w14""#),
            PARAGRAPHS,
        ),
        None,
    ));
    receiver
        .plan_paragraph_transfer_from(&lexical, Position::new(0))
        .unwrap();
}
