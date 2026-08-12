use std::io::{self, Write};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_docx::{Error, Package, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, TargetMode};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_SETTINGS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
const MAIN: &str = "word/document.xml";
const SETTINGS: &str = "word/settings.xml";
const UNUSED: &str = "word/unused.bin";

#[derive(Clone, Copy)]
enum SettingsLink {
    Missing,
    Internal(&'static str),
    External,
    Multiple,
    OtherInbound,
}

struct Fixture<'a> {
    settings: &'a [u8],
    relationship: SettingsLink,
    content_type: &'static str,
    signed: bool,
}

impl<'a> Fixture<'a> {
    fn transitional(settings: &'a [u8]) -> Self {
        Self {
            settings,
            relationship: SettingsLink::Internal(rt::SETTINGS),
            content_type: ct::WML_SETTINGS,
            signed: false,
        }
    }
}

fn fixture(options: Fixture<'_>) -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new(format!("/{MAIN}")).unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>"#).into_bytes(),
    );
    match options.relationship {
        SettingsLink::Missing => {},
        SettingsLink::Internal(kind) => {
            main.rels_mut()
                .try_add_relationship(
                    kind.to_owned(),
                    "settings.xml".to_owned(),
                    "rSettings".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        },
        SettingsLink::External => {
            main.rels_mut()
                .try_add_relationship(
                    rt::SETTINGS.to_owned(),
                    "https://example.invalid/settings.xml".to_owned(),
                    "rSettings".to_owned(),
                    TargetMode::External,
                )
                .unwrap();
        },
        SettingsLink::Multiple => {
            main.rels_mut()
                .try_add_relationship(
                    rt::SETTINGS.to_owned(),
                    "settings.xml".to_owned(),
                    "rSettingsOne".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
            main.rels_mut()
                .try_add_relationship(
                    STRICT_SETTINGS.to_owned(),
                    "settings2.xml".to_owned(),
                    "rSettingsTwo".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        },
        SettingsLink::OtherInbound => {
            main.rels_mut()
                .try_add_relationship(
                    rt::SETTINGS.to_owned(),
                    "settings.xml".to_owned(),
                    "rSettings".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        },
    }
    package.try_add_part(Box::new(main)).unwrap();
    if !matches!(
        options.relationship,
        SettingsLink::Missing | SettingsLink::External
    ) {
        let mut settings = BlobPart::new(
            PackURI::new(format!("/{SETTINGS}")).unwrap(),
            options.content_type.to_owned(),
            options.settings.to_vec(),
        );
        settings
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:unrelated".to_owned(),
                "https://example.invalid/preserve".to_owned(),
                "rOpaque".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        package.try_add_part(Box::new(settings)).unwrap();
        if matches!(options.relationship, SettingsLink::Multiple) {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/word/settings2.xml").unwrap(),
                    ct::WML_SETTINGS.to_owned(),
                    options.settings.to_vec(),
                )))
                .unwrap();
        }
    }
    let mut unused = BlobPart::new(
        PackURI::new(format!("/{UNUSED}")).unwrap(),
        "application/octet-stream".to_owned(),
        opaque_payload(),
    );
    if matches!(options.relationship, SettingsLink::OtherInbound) {
        unused
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:other-owner".to_owned(),
                "settings.xml".to_owned(),
                "rOtherSettings".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.try_add_part(Box::new(unused)).unwrap();
    let main_relationship = if matches!(
        options.relationship,
        SettingsLink::Internal(STRICT_SETTINGS)
    ) {
        rt::STRICT_OFFICE_DOCUMENT
    } else {
        rt::OFFICE_DOCUMENT
    };
    package.relate_to(MAIN, main_relationship);
    if options.signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn opaque_payload() -> Vec<u8> {
    let mut state = 0x51A7_E123_u32;
    (0..32 * 1024)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            (state >> 24) as u8
        })
        .collect()
}

fn open_source(bytes: Vec<u8>) -> source_backed::Package {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    source_backed::Package::from_read_at(source).unwrap()
}

fn payload_range(zip: &[u8], name: &str) -> std::ops::Range<usize> {
    let name = name.as_bytes();
    for (offset, _) in zip
        .windows(4)
        .enumerate()
        .filter(|(_, signature)| *signature == b"PK\x01\x02")
    {
        if offset + 46 > zip.len() {
            continue;
        }
        let compressed =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_len =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        if offset + 46 + name_len > zip.len() || &zip[offset + 46..offset + 46 + name_len] != name {
            continue;
        }
        let local = u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
        let local_name =
            u16::from_le_bytes(zip[local + 26..local + 28].try_into().unwrap()) as usize;
        let local_extra =
            u16::from_le_bytes(zip[local + 28..local + 30].try_into().unwrap()) as usize;
        let start = local + 30 + local_name + local_extra;
        return start..start + compressed;
    }
    panic!(
        "ZIP member was not found: {}",
        String::from_utf8_lossy(name)
    );
}

fn variables_xml(namespace: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><q:settings xmlns:q="{namespace}" xmlns:x="urn:opaque"><!--before--><q:zoom q:percent="133"/><q:docVars><q:docVar q:name="first" q:val="one"/><q:docVar q:name="second" q:val="two"/></q:docVars><x:opaque x:value="keep"><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
    )
    .into_bytes()
}

fn change_unrelated_member(bytes: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .get_part_mut(&PackURI::new(format!("/{UNUSED}")).unwrap())
        .unwrap()
        .set_blob(b"different unrelated bytes".to_vec());
    PackageWriter::to_bytes(&package).unwrap()
}

#[test]
fn set_remove_clear_noop_and_inverse_reopen_complete_packages() {
    let source_bytes = fixture(Fixture::transitional(&variables_xml(W)));
    let untouched = payload_range(&source_bytes, UNUSED);
    let package = open_source(source_bytes.clone());
    let mut edit = package.edit_document_variables().unwrap();
    edit.set("first", "updated").unwrap();
    edit.set("third", "three").unwrap();
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();

    let mut foreign_output = Vec::new();
    assert!(matches!(
        open_source(change_unrelated_member(&source_bytes))
            .publish_document_variables_commit_to_stream(&mut foreign_output, &commit),
        Err(Error::DocumentVariablesConflict)
    ));
    assert!(foreign_output.is_empty());

    let mut set_output = Vec::new();
    let publication = package
        .publish_document_variables_commit_to_stream(&mut set_output, &commit)
        .unwrap();
    assert_eq!(
        &set_output[payload_range(&set_output, UNUSED)],
        &source_bytes[untouched]
    );
    let reopened = Package::from_reader(io::Cursor::new(&set_output)).unwrap();
    let variables = reopened.document_variables().unwrap().unwrap();
    assert_eq!(variables.get("first"), Some("updated"));
    assert_eq!(variables.get("second"), Some("two"));
    assert_eq!(variables.get("third"), Some("three"));
    let opc = OpcPackage::from_bytes(&set_output).unwrap();
    let settings = opc
        .get_part(&PackURI::new(format!("/{SETTINGS}")).unwrap())
        .unwrap();
    assert!(settings.rels().get("rOpaque").is_some());
    let settings_text = std::str::from_utf8(settings.blob()).unwrap();
    assert!(settings_text.contains("<!--before--><q:zoom q:percent=\"133\"/>"));
    assert!(settings_text.contains("<x:opaque x:value=\"keep\"><![CDATA[a < b]]></x:opaque>"));

    let mut inverse_output = Vec::new();
    open_source(set_output.clone())
        .publish_document_variables_inverse_to_stream(&mut inverse_output, &publication)
        .unwrap();
    assert_eq!(inverse_output, source_bytes);
    let mut foreign_inverse_output = Vec::new();
    assert!(matches!(
        open_source(change_unrelated_member(&set_output))
            .publish_document_variables_inverse_to_stream(
                &mut foreign_inverse_output,
                &publication,
            ),
        Err(Error::DocumentVariablesConflict)
    ));
    assert!(foreign_inverse_output.is_empty());

    assert_eq!(
        inverse.apply(commit.snapshot()).unwrap().xml_bytes(),
        variables_xml(W)
    );
    let foreign_target = open_source(change_unrelated_member(&set_output))
        .document_variables_snapshot()
        .unwrap();
    assert!(matches!(
        inverse.apply(&foreign_target),
        Err(Error::DocumentVariablesConflict)
    ));

    let package = open_source(set_output);
    let mut remove = package.edit_document_variables().unwrap();
    assert_eq!(remove.remove("second"), Some("two".into()));
    let remove = remove.commit().unwrap();
    let mut removed_output = Vec::new();
    package
        .publish_document_variables_commit_to_stream(&mut removed_output, &remove)
        .unwrap();
    let reopened = Package::from_reader(io::Cursor::new(&removed_output)).unwrap();
    assert!(
        !reopened
            .document_variables()
            .unwrap()
            .unwrap()
            .contains("second")
    );

    let package = open_source(removed_output);
    let mut clear = package.edit_document_variables().unwrap();
    clear.clear();
    let clear = clear.commit().unwrap();
    let mut cleared_output = Vec::new();
    package
        .publish_document_variables_commit_to_stream(&mut cleared_output, &clear)
        .unwrap();
    let reopened = Package::from_reader(io::Cursor::new(&cleared_output)).unwrap();
    assert!(reopened.document_variables().unwrap().unwrap().is_empty());

    let package = open_source(cleared_output.clone());
    let no_op = package.edit_document_variables().unwrap().commit().unwrap();
    assert!(!no_op.changed());
    let mut no_op_output = Vec::new();
    package
        .publish_document_variables_commit_to_stream(&mut no_op_output, &no_op)
        .unwrap();
    assert_eq!(no_op_output, cleared_output);
}

#[test]
fn strict_settings_and_source_checked_patch_application_are_supported() {
    let strict = variables_xml(STRICT_W);
    let source_bytes = fixture(Fixture {
        settings: &strict,
        relationship: SettingsLink::Internal(STRICT_SETTINGS),
        content_type: ct::WML_SETTINGS,
        signed: false,
    });
    let package = open_source(source_bytes.clone());
    let mut edit = package.edit_document_variables().unwrap();
    edit.set("strict", "yes").unwrap();
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();
    let mut output = Vec::new();
    package
        .publish_document_variables_commit_to_stream(&mut output, &commit)
        .unwrap();
    let reopened = Package::from_reader(io::Cursor::new(&output)).unwrap();
    assert_eq!(
        reopened
            .document_variables()
            .unwrap()
            .unwrap()
            .get("strict"),
        Some("yes")
    );
    assert_eq!(
        inverse.apply(commit.snapshot()).unwrap().xml_bytes(),
        strict
    );

    let foreign_xml = String::from_utf8(strict)
        .unwrap()
        .replace("urn:opaque", "urn:foreign");
    let foreign = fixture(Fixture::transitional(foreign_xml.as_bytes()));
    let mut rejected = Vec::new();
    assert!(
        open_source(foreign)
            .publish_document_variables_commit_to_stream(&mut rejected, &commit)
            .is_err()
    );
    assert!(rejected.is_empty());
}

#[test]
fn creates_first_owner_in_existing_transitional_and_strict_settings() {
    for (word, relationship) in [(W, rt::SETTINGS), (STRICT_W, STRICT_SETTINGS)] {
        let settings = format!(r#"<settings xmlns="{word}"/>"#);
        let source_bytes = fixture(Fixture {
            settings: settings.as_bytes(),
            relationship: SettingsLink::Internal(relationship),
            content_type: ct::WML_SETTINGS,
            signed: false,
        });
        let untouched = payload_range(&source_bytes, UNUSED);
        let package = open_source(source_bytes.clone());
        let mut edit = package.edit_document_variables().unwrap();
        edit.set("first", "created").unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        package
            .publish_document_variables_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(
            &output[payload_range(&output, UNUSED)],
            &source_bytes[untouched]
        );
        let reopened = Package::from_reader(io::Cursor::new(&output)).unwrap();
        assert_eq!(
            reopened.document_variables().unwrap().unwrap().get("first"),
            Some("created")
        );

        // The locally declared canonical owner emitted for a default-namespace
        // root remains editable on the next source-backed round trip.
        let package = open_source(output);
        let mut second = package.edit_document_variables().unwrap();
        second.set("first", "updated").unwrap();
        let second = second.commit().unwrap();
        let mut second_output = Vec::new();
        package
            .publish_document_variables_commit_to_stream(&mut second_output, &second)
            .unwrap();
        let reopened = Package::from_reader(io::Cursor::new(second_output)).unwrap();
        assert_eq!(
            reopened.document_variables().unwrap().unwrap().get("first"),
            Some("updated")
        );
    }
}

#[test]
fn absent_multiple_external_wrong_type_and_mce_settings_fail_closed() {
    let xml = variables_xml(W);
    for fixture_bytes in [
        fixture(Fixture {
            settings: &xml,
            relationship: SettingsLink::Missing,
            content_type: ct::WML_SETTINGS,
            signed: false,
        }),
        fixture(Fixture {
            settings: &xml,
            relationship: SettingsLink::Multiple,
            content_type: ct::WML_SETTINGS,
            signed: false,
        }),
        fixture(Fixture {
            settings: &xml,
            relationship: SettingsLink::External,
            content_type: ct::WML_SETTINGS,
            signed: false,
        }),
        fixture(Fixture {
            settings: &xml,
            relationship: SettingsLink::Internal(rt::SETTINGS),
            content_type: "application/xml",
            signed: false,
        }),
    ] {
        assert!(
            open_source(fixture_bytes)
                .document_variables_snapshot()
                .is_err()
        );
    }

    let mce = format!(
        r#"<w:settings xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><w:docVars><w:docVar w:name="choice" w:val="ignored"/></w:docVars></mc:Choice><mc:Fallback><w:docVars><w:docVar w:name="fallback" w:val="selected"/></w:docVars></mc:Fallback></mc:AlternateContent></w:settings>"#
    );
    assert!(matches!(
        open_source(fixture(Fixture::transitional(mce.as_bytes()))).document_variables_snapshot(),
        Err(Error::UnsafeEdit { .. })
    ));
}

#[test]
fn additional_inbound_settings_owner_is_noop_safe_but_changed_refused() {
    let xml = variables_xml(W);
    let source_bytes = fixture(Fixture {
        settings: &xml,
        relationship: SettingsLink::OtherInbound,
        content_type: ct::WML_SETTINGS,
        signed: false,
    });
    let package = open_source(source_bytes.clone());
    let no_op = package.edit_document_variables().unwrap().commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_document_variables_commit_to_stream(&mut output, &no_op)
        .unwrap();
    assert_eq!(output, source_bytes);

    let package = open_source(source_bytes);
    let mut edit = package.edit_document_variables().unwrap();
    edit.set("first", "changed").unwrap();
    let commit = edit.commit().unwrap();
    output.clear();
    assert!(matches!(
        package.publish_document_variables_commit_to_stream(&mut output, &commit),
        Err(Error::DocumentVariablesPreservation(_))
    ));
    assert!(output.is_empty());
}

#[test]
fn selected_unknown_and_protected_changes_are_refused_but_noops_are_exact() {
    let unknown_sources = [
        format!(
            r#"<w:settings xmlns:w="{W}" xmlns:x="urn:unknown"><w:docVars x:opaque="keep"><w:docVar w:name="old" w:val="value"/><x:future x:value="keep"/></w:docVars></w:settings>"#
        ),
        format!(
            r#"<w:settings xmlns:w="{W}"><w:docVars xmlns:x="urn:selected"><w:docVar w:name="old" w:val="value"/></w:docVars></w:settings>"#
        ),
    ];
    let mut output = Vec::new();
    for unknown in unknown_sources {
        let unknown_bytes = fixture(Fixture::transitional(unknown.as_bytes()));
        let package = open_source(unknown_bytes.clone());
        let no_op = package.edit_document_variables().unwrap().commit().unwrap();
        output.clear();
        package
            .publish_document_variables_commit_to_stream(&mut output, &no_op)
            .unwrap();
        assert_eq!(output, unknown_bytes);
        let package = open_source(unknown_bytes);
        let mut edit = package.edit_document_variables().unwrap();
        edit.set("old", "changed").unwrap();
        let commit = edit.commit().unwrap();
        output.clear();
        assert!(
            package
                .publish_document_variables_commit_to_stream(&mut output, &commit)
                .is_err()
        );
        assert!(output.is_empty());
    }

    for protection in ["documentProtection", "writeProtection"] {
        let protected = format!(
            r#"<w:settings xmlns:w="{W}"><w:{protection}/><w:docVars><w:docVar w:name="old" w:val="value"/></w:docVars></w:settings>"#
        );
        let protected_bytes = fixture(Fixture::transitional(protected.as_bytes()));
        let package = open_source(protected_bytes.clone());
        let no_op = package.edit_document_variables().unwrap().commit().unwrap();
        output.clear();
        package
            .publish_document_variables_commit_to_stream(&mut output, &no_op)
            .unwrap();
        assert_eq!(output, protected_bytes);

        let package = open_source(protected_bytes);
        let mut edit = package.edit_document_variables().unwrap();
        edit.set("old", "changed").unwrap();
        let commit = edit.commit().unwrap();
        output.clear();
        assert!(matches!(
            package.publish_document_variables_commit_to_stream(&mut output, &commit),
            Err(Error::UnsafeEdit { .. })
        ));
        assert!(output.is_empty());
    }

    let duplicate = format!(
        r#"<w:settings xmlns:w="{W}"><w:documentProtection w:enforcement="on"/><w:documentProtection w:enforcement="off"/><w:docVars><w:docVar w:name="old" w:val="value"/></w:docVars></w:settings>"#
    );
    assert!(
        open_source(fixture(Fixture::transitional(duplicate.as_bytes())))
            .document_variables_snapshot()
            .is_err()
    );
}

#[test]
fn signed_change_source_version_limit_and_partial_sink_fail_before_or_during_output() {
    let xml = variables_xml(W);
    let signed_bytes = fixture(Fixture {
        settings: &xml,
        relationship: SettingsLink::Internal(rt::SETTINGS),
        content_type: ct::WML_SETTINGS,
        signed: true,
    });
    let package = open_source(signed_bytes.clone());
    let no_op = package.edit_document_variables().unwrap().commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_document_variables_commit_to_stream(&mut output, &no_op)
        .unwrap();
    assert_eq!(output, signed_bytes);
    let package = open_source(signed_bytes);
    let mut edit = package.edit_document_variables().unwrap();
    edit.set("first", "changed").unwrap();
    let commit = edit.commit().unwrap();
    output.clear();
    assert!(matches!(
        package.publish_document_variables_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());

    let versioned = Arc::new(VersionedSource::new(fixture(Fixture::transitional(&xml))));
    let read_at: Arc<dyn ReadAt> = versioned.clone();
    let package = source_backed::Package::from_read_at(read_at).unwrap();
    let mut edit = package.edit_document_variables().unwrap();
    edit.set("first", "changed").unwrap();
    let commit = edit.commit().unwrap();
    versioned.change();
    output.clear();
    assert!(matches!(
        package.publish_document_variables_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let limited_bytes = fixture(Fixture::transitional(&xml));
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(limited_bytes));
    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        source_backed::Package::from_read_at_with_limits(source, limits),
        Err(Error::Opc(OpcError::ReadLimit { .. }))
    ));

    let package = open_source(fixture(Fixture::transitional(&xml)));
    let mut edit = package.edit_document_variables().unwrap();
    edit.set("first", "changed").unwrap();
    let commit = edit.commit().unwrap();
    let mut sink = FailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        package.publish_document_variables_commit_to_stream(&mut sink, &commit),
        Err(Error::Opc(OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);
}

#[test]
fn overreporting_sink_returns_invalid_data_without_false_progress_or_panic() {
    let xml = variables_xml(W);
    let package = open_source(fixture(Fixture::transitional(&xml)));
    let commit = package.edit_document_variables().unwrap().commit().unwrap();
    let mut sink = OverReportingSink {
        calls: 0,
        accepted: 0,
    };

    assert!(matches!(
        package.publish_document_variables_commit_to_stream(&mut sink, &commit),
        Err(Error::Opc(OpcError::IoError(error)))
            if error.kind() == io::ErrorKind::InvalidData
    ));
    assert_eq!(sink.calls, 1);
    assert_eq!(sink.accepted, 0);
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn change(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(91, self.revision.load(Ordering::SeqCst)))
    }
}

struct FailingSink {
    accepted: usize,
    limit: usize,
}

struct OverReportingSink {
    calls: usize,
    accepted: usize,
}

impl Write for OverReportingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.calls = self.calls.saturating_add(1);
        Ok(bytes.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let count = bytes.len().min(self.limit - self.accepted);
        self.accepted += count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
