//! Public-facade tests for source-backed cross-presentation slide copying.
//!
//! These tests intentionally exercise the bounded dependency-free closure only:
//! one source slide, one internal layout edge, and no slide-owned resources.
//! The publisher must stage and validate the complete candidate before writing
//! any output, then raw-copy every untouched destination ZIP record.

use std::collections::{BTreeMap, BTreeSet};
use std::io::{self, Write};
use std::sync::{Arc, RwLock};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::constants::content_type as ct;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part};
use litchi_pptx::{
    Error, Package, ReadLimits, SourceBackedPresentation, SourceBackedPresentationEditor,
};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use soapberry_zip::office::ArchiveReader;

const PRESENTATION_RELATIONSHIPS: &str = "ppt/_rels/presentation.xml.rels";
const PRESENTATION_XML: &str = "ppt/presentation.xml";
const CONTENT_TYPES_XML: &str = "[Content_Types].xml";

type TestResult<T = ()> = litchi_pptx::Result<T>;

#[derive(Clone, Copy)]
struct Entry<'a> {
    name: &'a [u8],
    data: &'a [u8],
    local_extra: &'a [u8],
    central_extra: &'a [u8],
    comment: &'a [u8],
}

fn push_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

/// Rebuild a valid stored ZIP while deliberately retaining producer-style
/// local/central extras and member comments. This gives the raw-preservation
/// assertions a physical record to distinguish from decoded payload equality.
fn stored_archive(entries: &[Entry<'_>], archive_comment: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut central = Vec::new();
    for entry in entries {
        let local_offset = u32::try_from(output.len()).expect("fixture offset fits ZIP32");
        let size = u32::try_from(entry.data.len()).expect("fixture payload fits ZIP32");
        let crc = soapberry_zip::crc32(entry.data);
        push_u32(&mut output, 0x0403_4b50);
        push_u16(&mut output, 20);
        push_u16(&mut output, 0);
        push_u16(&mut output, 0);
        push_u16(&mut output, 0);
        push_u16(&mut output, 0);
        push_u32(&mut output, crc);
        push_u32(&mut output, size);
        push_u32(&mut output, size);
        push_u16(
            &mut output,
            u16::try_from(entry.name.len()).expect("fixture name fits ZIP32"),
        );
        push_u16(
            &mut output,
            u16::try_from(entry.local_extra.len()).expect("fixture extra fits ZIP32"),
        );
        output.extend_from_slice(entry.name);
        output.extend_from_slice(entry.local_extra);
        output.extend_from_slice(entry.data);

        push_u32(&mut central, 0x0201_4b50);
        push_u16(&mut central, 20);
        push_u16(&mut central, 20);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, crc);
        push_u32(&mut central, size);
        push_u32(&mut central, size);
        push_u16(
            &mut central,
            u16::try_from(entry.name.len()).expect("fixture name fits ZIP32"),
        );
        push_u16(
            &mut central,
            u16::try_from(entry.central_extra.len()).expect("fixture extra fits ZIP32"),
        );
        push_u16(
            &mut central,
            u16::try_from(entry.comment.len()).expect("fixture comment fits ZIP32"),
        );
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, 0);
        push_u32(&mut central, local_offset);
        central.extend_from_slice(entry.name);
        central.extend_from_slice(entry.central_extra);
        central.extend_from_slice(entry.comment);
    }
    let central_offset = u32::try_from(output.len()).expect("fixture offset fits ZIP32");
    let central_size = u32::try_from(central.len()).expect("fixture central fits ZIP32");
    output.extend_from_slice(&central);
    push_u32(&mut output, 0x0605_4b50);
    push_u16(&mut output, 0);
    push_u16(&mut output, 0);
    let count = u16::try_from(entries.len()).expect("fixture entry count fits ZIP32");
    push_u16(&mut output, count);
    push_u16(&mut output, count);
    push_u32(&mut output, central_size);
    push_u32(&mut output, central_offset);
    push_u16(
        &mut output,
        u16::try_from(archive_comment.len()).expect("fixture comment fits ZIP32"),
    );
    output.extend_from_slice(archive_comment);
    output
}

fn add_zip_metadata(bytes: &[u8], archive_comment: &[u8]) -> TestResult<Vec<u8>> {
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| Error::Invalid(format!("cannot index fixture ZIP: {error}")))?;
    let mut names = Vec::new();
    let mut payloads = Vec::new();
    for member in archive.file_names() {
        names.push(member.to_owned());
        payloads.push(archive.read(member).map_err(|error| {
            Error::Invalid(format!("cannot read fixture member {member}: {error}"))
        })?);
    }
    let entries = names
        .iter()
        .zip(payloads.iter())
        .enumerate()
        .map(|(index, (name, payload))| Entry {
            name: name.as_bytes(),
            data: payload,
            local_extra: if index % 2 == 0 {
                b"\x99\x99\x04\x00meta"
            } else {
                b"\x99\x99\x04\x00keep"
            },
            central_extra: if index % 2 == 0 {
                b"\x88\x88\x02\x00ce"
            } else {
                b"\x88\x88\x02\x00ck"
            },
            comment: if index % 2 == 0 {
                b"member-comment"
            } else {
                b"untouched-comment"
            },
        })
        .collect::<Vec<_>>();
    Ok(stored_archive(&entries, archive_comment))
}

/// Builds an authored package with the writer's registered layout/master/theme
/// graph and dependency-free slide bodies. Names are changed in the ZIP after
/// authoring because the public writer intentionally derives names from IDs.
fn authored_package(slide_titles: &[&str]) -> TestResult<Vec<u8>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for title in slide_titles {
            let slide = presentation.add_slide()?;
            slide.set_title(title);
            slide.add_text_box(&format!("body:{title}"), 10, 20, 300, 400);
        }
    }
    package.to_bytes()
}

fn source_fixture() -> TestResult<Vec<u8>> {
    let bytes = authored_package(&["source-one", "source-two"])?;
    let renamed = rewrite_slide_names(
        &bytes,
        &[
            ("ppt/slides/slide1.xml", "Source One"),
            ("ppt/slides/slide2.xml", "Source Two"),
        ],
    )?;
    add_zip_metadata(&renamed, b"source-backed-cross-copy-source")
}

fn destination_fixture() -> TestResult<Vec<u8>> {
    let bytes = authored_package(&["destination-first", "destination-last"])?;
    let renamed = rewrite_slide_names(
        &bytes,
        &[
            ("ppt/slides/slide1.xml", "Destination First"),
            ("ppt/slides/slide2.xml", "Destination Last"),
        ],
    )?;
    // Keep an unrelated, registered media part in the destination. It is
    // deliberately outside the copied slide closure and must be carried
    // through byte-for-byte by the source-backed publisher.
    let mut package = OpcPackage::from_vec(renamed)?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/unrelated.png").map_err(Error::Uri)?,
        ct::PNG.to_owned(),
        (0_u8..=31).collect(),
    )))?;
    package
        .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Uri)?)?
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/unrelated.png".to_owned(),
            "rIdUnrelatedMedia".to_owned(),
            litchi_opc::TargetMode::Internal,
        )?;
    let renamed = PackageWriter::to_bytes(&package)?;
    add_zip_metadata(&renamed, b"source-backed-cross-copy-destination")
}

fn open_source(bytes: &[u8]) -> TestResult<SourceBackedPresentation> {
    SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
}

fn open_editor(bytes: &[u8]) -> TestResult<SourceBackedPresentationEditor> {
    SourceBackedPresentationEditor::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
}

fn publish(
    source_bytes: &[u8],
    destination_bytes: &[u8],
    source_slide: usize,
    destination_slide: usize,
    position: usize,
) -> TestResult<Vec<u8>> {
    let source = open_source(source_bytes)?;
    let editor = open_editor(destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, source_slide, destination_slide, position)?;
    let mut output = Vec::new();
    editor.publish_cross_slide_copy_to_stream(&mut output, &plan)?;
    Ok(output)
}

fn plan_refused(source_bytes: &[u8], destination_bytes: &[u8]) -> TestResult<bool> {
    let source = match open_source(source_bytes) {
        Ok(source) => source,
        Err(_) => return Ok(true),
    };
    let editor = match open_editor(destination_bytes) {
        Ok(editor) => editor,
        Err(_) => return Ok(true),
    };
    Ok(editor.plan_cross_slide_copy(&source, 0, 1, 1).is_err())
}

fn is_source_changed(error: &Error) -> bool {
    match error {
        Error::StaleSource => true,
        Error::Opc(OpcError::SourceChanged { .. }) => true,
        Error::Opc(OpcError::IncompleteOutput { source, .. }) => {
            matches!(source.as_ref(), OpcError::SourceChanged { .. })
        },
        _ => false,
    }
}

fn memory_limit_observed(error: &Error) -> Option<u64> {
    fn opc(error: &OpcError) -> Option<u64> {
        match error {
            OpcError::Execution(litchi_core::ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory =>
            {
                Some(limit.observed)
            },
            OpcError::IncompleteOutput { source, .. } => opc(source),
            _ => None,
        }
    }
    match error {
        Error::Opc(error) => opc(error),
        _ => None,
    }
}

fn slide_projection(bytes: &[u8]) -> TestResult<Vec<(String, String)>> {
    let presentation = open_source(bytes)?;
    presentation
        .slides()
        .map(|slide| slide.text_and_name())
        .collect()
}

#[test]
fn source_backed_cross_copy_supports_middle_and_append_positions() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;

    let middle = publish(&source_bytes, &destination_bytes, 0, 1, 1)?;
    let middle_projection = slide_projection(&middle)?;
    assert_eq!(middle_projection.len(), 3);
    assert!(middle_projection[0].0.contains("destination-first"));
    assert!(middle_projection[1].0.contains("source-one"));
    assert!(middle_projection[2].0.contains("destination-last"));
    assert_eq!(middle_projection[0].1, "Destination First");
    assert_eq!(middle_projection[1].1, "Source One");
    assert_eq!(middle_projection[2].1, "Destination Last");

    let appended = publish(&source_bytes, &destination_bytes, 1, 1, 2)?;
    let projection = slide_projection(&appended)?;
    assert_eq!(projection.len(), 3);
    assert_eq!(projection[0].1, "Destination First");
    assert_eq!(projection[1].1, "Destination Last");
    assert_eq!(projection[2].1, "Source Two");
    Ok(())
}

#[test]
fn source_backed_cross_copy_reopens_through_source_and_eager_facades() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let source = open_source(&source_bytes)?;
    let editor = open_editor(&destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut output = Vec::new();
    let _snapshot = editor.publish_cross_slide_copy_to_stream(&mut output, &plan)?;

    let source_backed = open_source(&output)?;
    assert_eq!(source_backed.slide_count(), 3);
    assert_eq!(
        source_backed.slide(1).expect("copied slide").name()?,
        "Source One"
    );

    let eager = Package::from_vec(output.clone())?;
    let opened = eager.opened_presentation()?;
    assert_eq!(opened.slides().len(), 3);
    assert_eq!(opened.slides()[1].name(), "Source One");
    let eager_presentation = eager.presentation()?;
    let eager_slides = eager_presentation.slides()?;
    assert!(eager_slides[1].text()?.contains("source-one"));
    Ok(())
}

#[test]
fn source_backed_cross_copy_keeps_inserted_namespace_bindings_self_contained() -> TestResult {
    let source_bytes = source_fixture()?;
    let relationship_namespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let destination_bytes = rewrite_member(&destination_fixture()?, PRESENTATION_XML, |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        let root_binding = format!(" xmlns:r=\"{relationship_namespace}\"");
        if !xml.contains(&root_binding) {
            return Err(Error::Invalid(
                "fixture presentation has no root relationship binding".into(),
            ));
        }
        let without_root = xml.replacen(&root_binding, "", 1);
        Ok(without_root
            .replace(
                " r:id=",
                &format!(" xmlns:r=\"{relationship_namespace}\" r:id="),
            )
            .into_bytes())
    })?;

    let output = publish(&source_bytes, &destination_bytes, 0, 1, 1)?;
    let projection = slide_projection(&output)?;
    assert_eq!(projection.len(), 3);
    let presentation = String::from_utf8(
        ArchiveReader::new(&output)
            .map_err(|error| Error::Invalid(format!("cannot index output: {error}")))?
            .read(PRESENTATION_XML)
            .map_err(|error| Error::Invalid(format!("cannot read presentation: {error}")))?,
    )
    .map_err(|error| Error::Invalid(format!("presentation is not UTF-8: {error}")))?;
    assert!(presentation.matches("xmlns:r=").count() >= 4);
    Ok(())
}

#[test]
fn source_backed_cross_copy_keeps_source_immutable_and_is_deterministic() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let source = open_source(&source_bytes)?;
    let source_before = source_bytes.clone();

    let first = {
        let editor = open_editor(&destination_bytes)?;
        let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
        let mut output = Vec::new();
        editor.publish_cross_slide_copy_to_stream(&mut output, &plan)?;
        output
    };
    let second = {
        let editor = open_editor(&destination_bytes)?;
        let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
        let mut output = Vec::new();
        editor.publish_cross_slide_copy_to_stream(&mut output, &plan)?;
        output
    };

    assert_eq!(first, second);
    assert_eq!(source_bytes, source_before);
    source.check_source()?;
    assert_eq!(source.slide(0).expect("source slide").name()?, "Source One");
    Ok(())
}

#[test]
fn source_backed_cross_copy_preserves_untouched_destination_zip_records() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let before = raw_archive(&destination_bytes)?;
    let output = publish(&source_bytes, &destination_bytes, 0, 1, 1)?;
    let after = raw_archive(&output)?;

    assert_eq!(after.comment, before.comment);
    let changed = BTreeSet::from([
        CONTENT_TYPES_XML.to_owned(),
        PRESENTATION_XML.to_owned(),
        PRESENTATION_RELATIONSHIPS.to_owned(),
    ]);
    let appended: BTreeSet<_> = after
        .members
        .keys()
        .filter(|name| !before.members.contains_key(*name))
        .cloned()
        .collect();
    assert_eq!(
        appended.len(),
        2,
        "one slide and one slide relationship part"
    );
    assert!(
        appended
            .iter()
            .any(|name| name.starts_with("ppt/slides/slide"))
    );
    assert!(
        appended
            .iter()
            .any(|name| name.contains("ppt/slides/_rels/slide"))
    );
    let appended_local_order: Vec<_> = after
        .local_order
        .iter()
        .filter(|name| !before.members.contains_key(*name))
        .cloned()
        .collect();
    let appended_central_order: Vec<_> = after
        .central_order
        .iter()
        .filter(|name| !before.members.contains_key(*name))
        .cloned()
        .collect();
    assert_eq!(appended_local_order, appended_central_order);
    assert_eq!(appended_local_order.len(), 2);

    for name in &before.local_order {
        if changed.contains(name) {
            continue;
        }
        let old = before.members.get(name).expect("old local member");
        let new = after.members.get(name).expect("preserved local member");
        assert_eq!(new.local, old.local, "local record changed for {name}");
        assert_eq!(new.payload, old.payload, "payload changed for {name}");
    }
    for name in &before.central_order {
        if changed.contains(name) {
            continue;
        }
        let old = before.members.get(name).expect("old central member");
        let new = after.members.get(name).expect("preserved central member");
        assert_eq!(
            new.central, old.central,
            "central record changed for {name}"
        );
    }
    Ok(())
}

#[test]
fn source_backed_cross_copy_skips_and_preserves_an_orphan_slide_relationship_member() -> TestResult
{
    let source_bytes = source_fixture()?;
    let orphan_name = "ppt/slides/_rels/slide3.xml.rels";
    let orphan_payload = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
    let destination_bytes =
        add_stored_member(&destination_fixture()?, orphan_name, orphan_payload)?;
    let before = raw_archive(&destination_bytes)?;
    let output = publish(&source_bytes, &destination_bytes, 0, 1, 1)?;
    let after = raw_archive(&output)?;

    let old = before
        .members
        .get(orphan_name)
        .expect("orphan relationship member before publication");
    let new = after
        .members
        .get(orphan_name)
        .expect("orphan relationship member after publication");
    assert_eq!(new.local, old.local);
    assert_eq!(new.central, old.central);
    assert_eq!(new.payload, old.payload);
    assert!(after.members.contains_key("ppt/slides/slide4.xml"));
    assert!(
        after
            .members
            .contains_key("ppt/slides/_rels/slide4.xml.rels")
    );
    Ok(())
}

#[test]
fn source_backed_cross_copy_retargets_layout_to_destination_graph() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let output = publish(&source_bytes, &destination_bytes, 0, 1, 1)?;
    let archive = ArchiveReader::new(&output)
        .map_err(|error| Error::Invalid(format!("cannot index output: {error}")))?;
    let source_archive = ArchiveReader::new(&source_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index source: {error}")))?;
    let slide_member = archive
        .file_names()
        .find(|name| {
            name.starts_with("ppt/slides/slide")
                && archive
                    .read(name)
                    .map(|bytes| {
                        bytes
                            .windows(b"Source One".len())
                            .any(|w| w == b"Source One")
                    })
                    .unwrap_or(false)
        })
        .ok_or_else(|| Error::Invalid("copied slide member not found".into()))?;
    let file = slide_member
        .strip_prefix("ppt/slides/")
        .ok_or_else(|| Error::Invalid("invalid copied slide member".into()))?;
    let stem = file
        .strip_suffix(".xml")
        .ok_or_else(|| Error::Invalid("invalid copied slide extension".into()))?;
    let rels_name = format!("ppt/slides/_rels/{stem}.xml.rels");
    let source_slide_member = source_archive
        .file_names()
        .find(|name| {
            name.starts_with("ppt/slides/slide")
                && source_archive
                    .read(name)
                    .map(|bytes| {
                        bytes
                            .windows(b"Source One".len())
                            .any(|window| window == b"Source One")
                    })
                    .unwrap_or(false)
        })
        .ok_or_else(|| Error::Invalid("source slide member not found".into()))?;
    let copied_slide = archive
        .read(slide_member)
        .map_err(|error| Error::Invalid(format!("cannot read copied slide: {error}")))?;
    let source_slide = source_archive
        .read(source_slide_member)
        .map_err(|error| Error::Invalid(format!("cannot read source slide: {error}")))?;
    assert_eq!(copied_slide, source_slide);
    let rels =
        String::from_utf8(archive.read(&rels_name).map_err(|error| {
            Error::Invalid(format!("cannot read copied relationships: {error}"))
        })?)
        .map_err(|error| Error::Invalid(format!("copied relationships are not UTF-8: {error}")))?;
    assert_eq!(rels.matches("Type=").count(), 1);
    assert!(rels.contains("/relationships/slideLayout"));
    assert!(rels.contains("../slideLayouts/slideLayout1.xml"));
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_stale_foreign_and_same_source_plans() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let other_destination = authored_package(&["other-destination"])?;

    let source = open_source(&source_bytes)?;
    let editor = open_editor(&destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut foreign_output = Vec::new();
    assert!(
        open_editor(&other_destination)?
            .publish_cross_slide_copy_to_stream(&mut foreign_output, &plan)
            .is_err()
    );
    assert!(foreign_output.is_empty());

    let shared: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(destination_bytes.clone()));
    let same_source = SourceBackedPresentation::from_read_at(Arc::clone(&shared))?;
    assert!(
        SourceBackedPresentationEditor::from_read_at(shared)?
            .plan_cross_slide_copy(&same_source, 0, 1, 1)
            .is_err()
    );

    let mutable_destination = MutableSource::new(destination_bytes.clone());
    let destination_source: Arc<dyn ReadAt> = Arc::new(mutable_destination.clone());
    let stale_editor = SourceBackedPresentationEditor::from_read_at(destination_source)?;
    let plan = stale_editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    mutable_destination.replace(source_fixture()?);
    let mut stale_output = Vec::new();
    let stale_error =
        match stale_editor.publish_cross_slide_copy_to_stream(&mut stale_output, &plan) {
            Ok(_) => panic!("destination mutation must invalidate publication"),
            Err(error) => error,
        };
    assert!(
        is_source_changed(&stale_error),
        "unexpected error: {stale_error:?}"
    );
    assert!(stale_output.is_empty());

    let mutable_source = MutableSource::new(source_fixture()?);
    let source_view = SourceBackedPresentation::from_read_at(Arc::new(mutable_source.clone()))?;
    let stale_source_editor = open_editor(&destination_bytes)?;
    let source_plan = stale_source_editor.plan_cross_slide_copy(&source_view, 0, 1, 1)?;
    mutable_source.replace(destination_fixture()?);
    let mut stale_source_output = Vec::new();
    let stale_source_error = match stale_source_editor
        .publish_cross_slide_copy_to_stream(&mut stale_source_output, &source_plan)
    {
        Ok(_) => panic!("source mutation must invalidate publication"),
        Err(error) => error,
    };
    assert!(
        is_source_changed(&stale_source_error),
        "unexpected error: {stale_source_error:?}"
    );
    assert!(stale_source_output.is_empty());
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_name_and_layout_mismatch() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let collision_source = rewrite_slide_names(
        &source_bytes,
        &[("ppt/slides/slide1.xml", "Destination First")],
    )?;
    assert!(plan_refused(&collision_source, &destination_bytes)?);

    let mismatched_destination = rewrite_member(
        &destination_bytes,
        "ppt/slideLayouts/slideLayout1.xml",
        |bytes| {
            let mut changed = bytes.to_vec();
            changed.extend_from_slice(b"<!-- destination-only layout marker -->");
            Ok(changed)
        },
    )?;
    assert!(plan_refused(&source_bytes, &mismatched_destination)?);

    let noncanonical_destination_rels = rewrite_member(
        &destination_bytes,
        PRESENTATION_RELATIONSHIPS,
        |bytes| {
            let mut xml = String::from_utf8(bytes.to_vec())
                .map_err(|error| Error::Invalid(error.to_string()))?;
            xml = xml.replacen(
                "</Relationships>",
                "<Relationship Id=\"rIdExternalRoot\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.invalid\" TargetMode=\"External\"/></Relationships>",
                1,
            );
            Ok(xml.into_bytes())
        },
    )?;
    assert!(plan_refused(&source_bytes, &noncanonical_destination_rels)?);
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_unsupported_relationships_and_markup() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;

    let extra_slide = add_unreferenced_slide_part(&source_bytes)?;
    assert!(plan_refused(&extra_slide, &destination_bytes)?);

    let extra_media = add_media_relationship(&source_bytes, "ppt/slides/slide1.xml")?;
    assert!(plan_refused(&extra_media, &destination_bytes)?);

    let external = rewrite_member(&source_bytes, "ppt/slides/_rels/slide1.xml.rels", |bytes| {
        let mut xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        xml = xml.replacen(
                "</Relationships>",
                "<Relationship Id=\"rIdExternal\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.invalid\" TargetMode=\"External\"/></Relationships>",
                1,
            );
        Ok(xml.into_bytes())
    })?;
    assert!(plan_refused(&external, &destination_bytes)?);

    let mce = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        Ok(xml
            .replacen(
                "<p:sld ",
                "<p:sld xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"mc\" ",
                1,
            )
            .into_bytes())
    })?;
    assert!(plan_refused(&mce, &destination_bytes)?);

    let relationship_qualified_hidden = rewrite_member(
        &source_bytes,
        "ppt/slides/slide1.xml",
        |bytes| {
            let xml = String::from_utf8(bytes.to_vec())
                .map_err(|error| Error::Invalid(error.to_string()))?;
            Ok(xml
                .replacen(
                    "<p:sld ",
                    "<p:sld xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdHidden\" ",
                    1,
                )
                .into_bytes())
        },
    )?;
    assert!(plan_refused(
        &relationship_qualified_hidden,
        &destination_bytes
    )?);

    let mixed_relationship_namespace = rewrite_member(&source_bytes, PRESENTATION_XML, |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        let rewritten = xml.replacen(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
            1,
        );
        if rewritten == xml {
            return Err(Error::Invalid(
                "fixture presentation has no relationship namespace declaration".into(),
            ));
        }
        Ok(rewritten.into_bytes())
    })?;
    assert!(plan_refused(
        &mixed_relationship_namespace,
        &destination_bytes
    )?);

    let drawingml_slide_root = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        let rewritten = xml
            .replacen("<p:sld ", "<a:sld ", 1)
            .replacen("</p:sld>", "</a:sld>", 1);
        if rewritten == xml {
            return Err(Error::Invalid(
                "fixture slide root cannot be rewritten".into(),
            ));
        }
        Ok(rewritten.into_bytes())
    })?;
    assert!(plan_refused(&drawingml_slide_root, &destination_bytes)?);

    let processing_instruction = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        Ok(xml
            .replacen("?>", "?>\n<?fixture processing-instruction?>", 1)
            .into_bytes())
    })?;
    assert!(plan_refused(&processing_instruction, &destination_bytes)?);

    let doctype = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        Ok(xml
            .replacen("?>", "?>\n<!DOCTYPE p:sld []>", 1)
            .into_bytes())
    })?;
    assert!(plan_refused(&doctype, &destination_bytes)?);

    let invalid_reference = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        Ok(xml
            .replacen("body:source-one", "body:source-one&#0;", 1)
            .into_bytes())
    })?;
    assert!(plan_refused(&invalid_reference, &destination_bytes)?);

    let multiple_roots = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let mut xml = bytes.to_vec();
        xml.extend_from_slice(
            b"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>",
        );
        Ok(xml)
    })?;
    assert!(plan_refused(&multiple_roots, &destination_bytes)?);

    let repeated_declaration = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        Ok(xml
            .replacen("?>", "?><?xml version=\"1.0\"?>", 1)
            .into_bytes())
    })?;
    assert!(plan_refused(&repeated_declaration, &destination_bytes)?);
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_protection_opaque_and_trailing_data() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;

    let protected = rewrite_member(&source_bytes, PRESENTATION_XML, |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        Ok(xml
            .replacen(
                "</p:presentation>",
                "<p:modifyVerifier cryptAlgorithmSid=\"14\" spinCount=\"1\" saltData=\"AA==\" hashData=\"AA==\"/></p:presentation>",
                1,
            )
            .into_bytes())
    })?;
    assert!(plan_refused(&protected, &destination_bytes)?);

    let opaque_source = add_stored_member(&source_bytes, "producer/opaque.bin", b"opaque")?;
    assert!(plan_refused(&opaque_source, &destination_bytes)?);

    let macro_source = add_stored_member(&source_bytes, "ppt/vbaProject.bin", b"macro")?;
    assert!(plan_refused(&macro_source, &destination_bytes)?);

    let signature_source = add_stored_member(
        &source_bytes,
        "_xmlsignatures/origin.sigs",
        b"<signatures/>",
    )?;
    assert!(plan_refused(&signature_source, &destination_bytes)?);

    let encrypted_source = mark_entry_encrypted(
        add_unreferenced_media_part(&source_bytes)?,
        "ppt/media/unused.png",
    )?;
    assert!(plan_refused(&encrypted_source, &destination_bytes)?);

    let encrypted_destination =
        mark_entry_encrypted(destination_bytes.clone(), "ppt/media/unrelated.png")?;
    assert!(plan_refused(&source_bytes, &encrypted_destination)?);

    let trailing = {
        let mut bytes = source_bytes.clone();
        bytes.extend_from_slice(b"trailing-source-data");
        bytes
    };
    assert!(plan_refused(&trailing, &destination_bytes)?);
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_limits_and_cancellation_before_output() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let too_small = ReadLimits::builder()
        .max_archive_entry_bytes(1)
        .map_err(|error| Error::Invalid(error.to_string()))?
        .build()?;
    assert!(
        SourceBackedPresentation::from_read_at_with_limits(
            Arc::new(OwnedSource::new(source_bytes.clone())),
            too_small,
        )
        .is_err()
    );

    let budget = Budget::root(
        "pptx-cross-copy-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("nonzero workers"),
        std::num::NonZeroUsize::new(1).expect("nonzero tasks"),
        std::num::NonZeroU64::new(u64::MAX).expect("nonzero memory"),
        0,
    )
    .map_err(|error| Error::Invalid(error.to_string()))?;
    let context = ExecutionContext::new(budget, cancellation, execution_limits);
    let source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(source_bytes)),
        ReadLimits::default(),
        context.clone(),
    )?;
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(destination_bytes)),
        ReadLimits::default(),
        context,
    )?;
    cancellation_source.cancel();
    assert!(editor.plan_cross_slide_copy(&source, 0, 1, 1).is_err());
    Ok(())
}

#[test]
fn source_backed_cross_copy_preserves_typed_cancellation_during_output() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let budget = Budget::root(
        "pptx-cross-copy-output-cancel",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("nonzero workers"),
        std::num::NonZeroUsize::new(1).expect("nonzero tasks"),
        std::num::NonZeroU64::new(u64::MAX).expect("nonzero memory"),
        0,
    )
    .map_err(|error| Error::Invalid(error.to_string()))?;
    let context = ExecutionContext::new(budget, cancellation, execution_limits);
    let source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(source_bytes)),
        ReadLimits::default(),
        context.clone(),
    )?;
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(destination_bytes)),
        ReadLimits::default(),
        context,
    )?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut writer = CancellingWriter::new(cancellation_source);
    let error = editor
        .publish_cross_slide_copy_to_stream(&mut writer, &plan)
        .expect_err("mid-output cancellation must stop publication");
    match error {
        Error::Opc(OpcError::IncompleteOutput { written, source }) => {
            assert_eq!(written, writer.bytes.len() as u64);
            assert!(written > 0);
            assert!(matches!(source.as_ref(), OpcError::Cancelled));
        },
        other => panic!("unexpected cancellation error: {other:?}"),
    }
    Ok(())
}

#[test]
fn source_backed_cross_copy_reports_source_change_during_output_with_exact_progress() -> TestResult
{
    let mutable_source = MutableSource::new(source_fixture()?);
    let source = SourceBackedPresentation::from_read_at(Arc::new(mutable_source.clone()))?;
    let editor = open_editor(&destination_fixture()?)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut writer = SourceMutatingWriter::new(mutable_source, destination_fixture()?);
    let error = editor
        .publish_cross_slide_copy_to_stream(&mut writer, &plan)
        .expect_err("mid-output source change must stop publication");
    match error {
        Error::Opc(OpcError::IncompleteOutput { written, source }) => {
            assert_eq!(written, writer.bytes.len() as u64);
            assert!(written > 0);
            assert!(matches!(source.as_ref(), OpcError::SourceChanged { .. }));
        },
        other => panic!("unexpected source-change error: {other:?}"),
    }
    Ok(())
}

#[test]
fn source_backed_cross_copy_managed_memory_reservation_is_exact_and_bounded() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let source_slide_bytes = ArchiveReader::new(&source_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index source: {error}")))?
        .read("ppt/slides/slide1.xml")
        .map_err(|error| Error::Invalid(format!("cannot read source slide: {error}")))?;
    let destination_presentation_bytes = ArchiveReader::new(&destination_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index destination: {error}")))?
        .read(PRESENTATION_XML)
        .map_err(|error| {
            Error::Invalid(format!("cannot read destination presentation: {error}"))
        })?;
    let publication_reservation = source_slide_bytes
        .len()
        .checked_add(destination_presentation_bytes.len())
        .and_then(|bytes| bytes.checked_add(256))
        .ok_or_else(|| Error::Invalid("fixture memory estimate overflow".into()))?;

    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("nonzero workers"),
        std::num::NonZeroUsize::new(1).expect("nonzero tasks"),
        std::num::NonZeroU64::new(u64::MAX).expect("nonzero memory"),
        0,
    )
    .map_err(|error| Error::Invalid(error.to_string()))?;
    let measure_budget = Budget::root(
        "pptx-cross-copy-measure",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (measure_cancel, measure_token) = CancellationSource::pair();
    let measure_context =
        ExecutionContext::new(measure_budget.clone(), measure_token, execution_limits);
    let measured_source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(source_bytes.clone())),
        ReadLimits::default(),
        measure_context.clone(),
    )?;
    let measured_editor =
        SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
            Arc::new(OwnedSource::new(destination_bytes.clone())),
            ReadLimits::default(),
            measure_context,
        )?;
    let measured_plan = measured_editor.plan_cross_slide_copy(&measured_source, 0, 1, 1)?;
    let base_memory = measure_budget.used(Resource::Memory);
    drop(measured_plan);
    drop(measured_editor);
    drop(measured_source);
    measure_cancel.cancel();
    let initial_memory = base_memory
        .checked_add(u64::try_from(publication_reservation).map_err(|error| {
            Error::Invalid(format!("fixture memory estimate does not fit u64: {error}"))
        })?)
        .ok_or_else(|| Error::Invalid("fixture memory budget overflow".into()))?;

    let run = |memory_limit: u64| -> TestResult<(Vec<u8>, Option<Error>)> {
        let budget = Budget::root(
            "pptx-cross-copy-run",
            Limits::new(
                memory_limit,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        );
        let (_cancel_source, token) = CancellationSource::pair();
        let context = ExecutionContext::new(budget, token, execution_limits);
        let source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
            Arc::new(OwnedSource::new(source_bytes.clone())),
            ReadLimits::default(),
            context.clone(),
        )?;
        let editor =
            SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
                Arc::new(OwnedSource::new(destination_bytes.clone())),
                ReadLimits::default(),
                context,
            )?;
        let plan = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
            Ok(plan) => plan,
            Err(error) => return Ok((Vec::new(), Some(error))),
        };
        let mut output = Vec::new();
        match editor.publish_cross_slide_copy_to_stream(&mut output, &plan) {
            Ok(_) => Ok((output, None)),
            Err(error) => Ok((output, Some(error))),
        }
    };

    // The PPTX staging reservation composes with OPC's finite physical-name,
    // relationship, and content-types working reservations. Follow typed
    // `observed` boundaries until the complete operation fits, then prove the
    // resulting total is the exact hierarchical limit for this artifact.
    let mut exact_memory = initial_memory;
    loop {
        let (output, error) = run(exact_memory)?;
        match error {
            None => {
                assert!(!output.is_empty());
                break;
            },
            Some(error) => {
                let observed = memory_limit_observed(&error).ok_or_else(|| {
                    Error::Invalid(format!(
                        "memory calibration returned a non-memory error: {error:?}"
                    ))
                })?;
                if observed <= exact_memory {
                    return Err(Error::Invalid(
                        "memory calibration did not advance the typed boundary".into(),
                    ));
                }
                exact_memory = observed;
            },
        }
    }

    let (exact_output, exact_error) = run(exact_memory)?;
    assert!(
        exact_error.is_none(),
        "exact reservation rejected: {exact_error:?}"
    );
    assert_eq!(
        exact_output,
        publish(&source_bytes, &destination_bytes, 0, 1, 1)?
    );

    let (under_output, under_error) = run(exact_memory.saturating_sub(1))?;
    assert!(under_output.is_empty());
    assert!(matches!(
        under_error,
        Some(Error::Opc(OpcError::Execution(_)))
            | Some(Error::Opc(OpcError::Cancelled))
            | Some(Error::Opc(OpcError::IncompleteOutput { .. }))
    ));
    Ok(())
}

#[test]
fn source_backed_cross_copy_reports_partial_short_and_zero_sink_progress() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let expected = publish(&source_bytes, &destination_bytes, 0, 1, 1)?;
    let source = open_source(&source_bytes)?;
    let editor = open_editor(&destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;

    let mut short = ShortWriter::new(1);
    editor.publish_cross_slide_copy_to_stream(&mut short, &plan)?;
    assert_eq!(short.bytes, expected);
    assert_eq!(slide_projection(&short.bytes)?.len(), 3);

    let editor = open_editor(&destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut failing = FailingWriter::new(128);
    let failing_error = match editor.publish_cross_slide_copy_to_stream(&mut failing, &plan) {
        Ok(_) => panic!("bounded sink must fail after its accepted prefix"),
        Err(error) => error,
    };
    match failing_error {
        Error::Opc(OpcError::IncompleteOutput { written, source }) => {
            assert_eq!(written, failing.bytes.len() as u64);
            assert_eq!(failing.bytes, expected[..failing.bytes.len()]);
            assert!(matches!(source.as_ref(), OpcError::IoError(_)));
        },
        other => panic!("unexpected bounded sink error: {other:?}"),
    }

    let editor = open_editor(&destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut zero = ZeroWriter;
    let zero_error = match editor.publish_cross_slide_copy_to_stream(&mut zero, &plan) {
        Ok(_) => panic!("zero-progress sink must fail"),
        Err(error) => error,
    };
    assert!(matches!(
        zero_error,
        Error::Opc(OpcError::IoError(_))
            | Error::Opc(OpcError::IncompleteOutput { written: 0, .. })
    ));

    let editor = open_editor(&destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut flush = FlushFailingWriter::default();
    let flush_error = match editor.publish_cross_slide_copy_to_stream(&mut flush, &plan) {
        Ok(_) => panic!("flush failure must fail publication"),
        Err(error) => error,
    };
    match flush_error {
        Error::Opc(OpcError::IncompleteOutput { written, source }) => {
            assert_eq!(written, expected.len() as u64);
            assert_eq!(flush.bytes, expected);
            assert!(matches!(source.as_ref(), OpcError::IoError(_)));
        },
        other => panic!("unexpected flush error: {other:?}"),
    }

    let editor = open_editor(&destination_bytes)?;
    let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
    let mut over = OverReportingWriter;
    let over_error = match editor.publish_cross_slide_copy_to_stream(&mut over, &plan) {
        Ok(_) => panic!("over-reporting sink must fail publication"),
        Err(error) => error,
    };
    assert!(matches!(
        over_error,
        Error::Opc(OpcError::IoError(error))
            if error.kind() == io::ErrorKind::InvalidData
    ));
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_missing_name_and_duplicate_ids() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let missing_name = rewrite_member(&source_bytes, "ppt/slides/slide1.xml", |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        let start = xml
            .find(" name=\"")
            .ok_or_else(|| Error::Invalid("missing fixture name attribute".into()))?;
        let end = xml[start..]
            .find('"')
            .and_then(|offset| {
                xml[start + offset + 1..]
                    .find('"')
                    .map(|next| start + offset + next + 2)
            })
            .ok_or_else(|| Error::Invalid("unterminated fixture name".into()))?;
        Ok(format!("{}{}", &xml[..start], &xml[end..]).into_bytes())
    })?;
    assert!(plan_refused(&missing_name, &destination_bytes)?);

    let duplicate_ids = rewrite_member(&source_bytes, PRESENTATION_XML, |bytes| {
        let xml =
            String::from_utf8(bytes.to_vec()).map_err(|error| Error::Invalid(error.to_string()))?;
        let marker = "<p:sldId id=\"";
        let first = xml
            .find(marker)
            .ok_or_else(|| Error::Invalid("missing first slide ID".into()))?;
        let first_start = first + marker.len();
        let first_end = xml[first_start..]
            .find('"')
            .ok_or_else(|| Error::Invalid("unterminated first slide ID".into()))?
            + first_start;
        let second = xml[first_end..]
            .find(marker)
            .ok_or_else(|| Error::Invalid("missing second slide ID".into()))?
            + first_end;
        let second_start = second + marker.len();
        let second_end = xml[second_start..]
            .find('"')
            .ok_or_else(|| Error::Invalid("unterminated second slide ID".into()))?
            + second_start;
        let id = &xml[first_start..first_end];
        Ok(format!("{}{}{}", &xml[..second_start], id, &xml[second_end..]).into_bytes())
    })?;
    assert!(plan_refused(&duplicate_ids, &destination_bytes)?);
    Ok(())
}

#[derive(Debug, Default)]
struct ShortWriter {
    bytes: Vec<u8>,
    max_per_write: usize,
}

#[derive(Debug)]
struct CancellingWriter {
    bytes: Vec<u8>,
    cancellation: CancellationSource,
    cancelled: bool,
}

struct SourceMutatingWriter {
    bytes: Vec<u8>,
    source: MutableSource,
    replacement: Option<Vec<u8>>,
}

impl SourceMutatingWriter {
    fn new(source: MutableSource, replacement: Vec<u8>) -> Self {
        Self {
            bytes: Vec::new(),
            source,
            replacement: Some(replacement),
        }
    }
}

impl Write for SourceMutatingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !bytes.is_empty()
            && let Some(replacement) = self.replacement.take()
        {
            self.source.replace(replacement);
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl CancellingWriter {
    fn new(cancellation: CancellationSource) -> Self {
        Self {
            bytes: Vec::new(),
            cancellation,
            cancelled: false,
        }
    }
}

impl Write for CancellingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !bytes.is_empty() && !self.cancelled {
            self.cancellation.cancel();
            self.cancelled = true;
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl ShortWriter {
    fn new(max_per_write: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max_per_write,
        }
    }
}

impl Write for ShortWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(self.max_per_write);
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct FailingWriter {
    bytes: Vec<u8>,
    limit: usize,
}

impl FailingWriter {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }
}

impl Write for FailingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.limit {
            return Err(io::Error::other("deliberate sink failure"));
        }
        let count = bytes.len().min(self.limit - self.bytes.len());
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ZeroWriter;

impl Write for ZeroWriter {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug, Default)]
struct FlushFailingWriter {
    bytes: Vec<u8>,
}

impl Write for FlushFailingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Err(io::Error::other("deliberate flush failure"))
    }
}

struct OverReportingWriter;

impl Write for OverReportingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        Ok(bytes.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Clone)]
struct MutableSource {
    bytes: Arc<RwLock<Vec<u8>>>,
    version: Arc<std::sync::atomic::AtomicU64>,
    id: u64,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(RwLock::new(bytes)),
            version: Arc::new(std::sync::atomic::AtomicU64::new(0)),
            id: 0x50505458,
        }
    }

    fn replace(&self, bytes: Vec<u8>) {
        *self.bytes.write().expect("mutable source write lock") = bytes;
        self.version
            .fetch_add(1, std::sync::atomic::Ordering::AcqRel);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.read().expect("mutable source read lock").len())
            .map_err(|error| io::Error::other(error.to_string()))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.read().expect("mutable source read lock");
        let start = usize::try_from(offset).unwrap_or(bytes.len());
        let input = bytes.get(start..).unwrap_or_default();
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            self.id,
            self.version.load(std::sync::atomic::Ordering::Acquire),
        ))
    }
}

#[derive(Debug, Clone)]
struct RawMember {
    local: Vec<u8>,
    central: Vec<u8>,
    payload: Vec<u8>,
}

#[derive(Debug)]
struct RawArchive {
    members: BTreeMap<String, RawMember>,
    local_order: Vec<String>,
    central_order: Vec<String>,
    comment: Vec<u8>,
}

fn raw_archive(data: &[u8]) -> TestResult<RawArchive> {
    let slice = soapberry_zip::ZipArchive::from_slice(data)
        .map_err(|error| Error::Invalid(format!("cannot index raw ZIP: {error}")))?;
    let comment = slice.comment().as_bytes().to_vec();
    let archive = slice.into_zip_archive();
    let mut scratch = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = soapberry_zip::PreservationIndex::new(&archive, &mut scratch)
        .map_err(|error| Error::Invalid(format!("cannot preserve raw ZIP: {error}")))?;
    let mut members = BTreeMap::new();
    let mut local_positions = Vec::new();
    let mut central_order = Vec::new();
    for entry in index.entries() {
        let name = std::str::from_utf8(entry.raw_name_bytes())
            .map_err(|error| Error::Invalid(format!("raw ZIP name is not UTF-8: {error}")))?
            .to_owned();
        let local = entry.local_span();
        let central = entry.central_record();
        let local_start = usize::try_from(local.start)
            .map_err(|error| Error::Invalid(format!("local offset overflow: {error}")))?;
        let local_end = usize::try_from(local.end)
            .map_err(|error| Error::Invalid(format!("local end overflow: {error}")))?;
        let central_start = usize::try_from(central.start)
            .map_err(|error| Error::Invalid(format!("central offset overflow: {error}")))?;
        let central_end = usize::try_from(central.end)
            .map_err(|error| Error::Invalid(format!("central end overflow: {error}")))?;
        let local_bytes = data
            .get(local_start..local_end)
            .ok_or_else(|| Error::Invalid("local span out of bounds".into()))?
            .to_vec();
        let mut central_bytes = data
            .get(central_start..central_end)
            .ok_or_else(|| Error::Invalid("central span out of bounds".into()))?
            .to_vec();
        if central_bytes.len() < 46 {
            return Err(Error::Invalid("central record is too short".into()));
        }
        central_bytes[42..46].fill(0);
        let payload = ArchiveReader::new(data)
            .map_err(|error| Error::Invalid(format!("cannot read ZIP payload: {error}")))?
            .read(&name)
            .map_err(|error| Error::Invalid(format!("cannot read {name}: {error}")))?;
        if members
            .insert(
                name.clone(),
                RawMember {
                    local: local_bytes,
                    central: central_bytes,
                    payload,
                },
            )
            .is_some()
        {
            return Err(Error::Invalid(format!("duplicate raw ZIP member {name}")));
        }
        local_positions.push((local.start, name.clone()));
        central_order.push(name);
    }
    local_positions.sort_unstable_by_key(|(offset, _)| *offset);
    Ok(RawArchive {
        members,
        local_order: local_positions.into_iter().map(|(_, name)| name).collect(),
        central_order,
        comment,
    })
}

fn rewrite_slide_names(bytes: &[u8], names: &[(&str, &str)]) -> TestResult<Vec<u8>> {
    let replacements: BTreeMap<_, _> = names.iter().copied().collect();
    rewrite_archive(bytes, |name, payload| {
        let Some(new_name) = replacements.get(name) else {
            return Ok(payload.to_vec());
        };
        let xml = String::from_utf8(payload.to_vec())
            .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
        let marker = r#"name="#;
        let start = xml
            .find(marker)
            .ok_or_else(|| Error::Invalid(format!("missing slide name in {name}")))?;
        let end = xml[start..]
            .find('"')
            .and_then(|offset| {
                xml[start + offset + 1..]
                    .find('"')
                    .map(|next| start + offset + next + 2)
            })
            .ok_or_else(|| Error::Invalid("unterminated slide name".into()))?;
        Ok(format!("{}name=\"{new_name}\"{}", &xml[..start], &xml[end..]).into_bytes())
    })
}

fn rewrite_member<F>(bytes: &[u8], member: &str, mut replace: F) -> TestResult<Vec<u8>>
where
    F: FnMut(&[u8]) -> TestResult<Vec<u8>>,
{
    rewrite_archive(bytes, |name, payload| {
        if name == member {
            replace(payload)
        } else {
            Ok(payload.to_vec())
        }
    })
}

fn add_stored_member(bytes: &[u8], name: &str, payload: &[u8]) -> TestResult<Vec<u8>> {
    let mut output = rewrite_archive(bytes, |_name, payload| Ok(payload.to_vec()))?;
    let archive = ArchiveReader::new(&output)
        .map_err(|error| Error::Invalid(format!("cannot index rewritten ZIP: {error}")))?;
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for member in archive.file_names() {
        let data = archive
            .read(member)
            .map_err(|error| Error::Invalid(format!("cannot read {member}: {error}")))?;
        writer
            .write_stored(member, &data)
            .map_err(|error| Error::Invalid(format!("cannot copy {member}: {error}")))?;
    }
    writer
        .write_stored(name, payload)
        .map_err(|error| Error::Invalid(format!("cannot add {name}: {error}")))?;
    output = writer
        .finish_to_bytes()
        .map_err(|error| Error::Invalid(format!("cannot finish ZIP: {error}")))?;
    Ok(output)
}

fn add_unreferenced_slide_part(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(bytes.to_vec())?;
    let slide_name = PackURI::new("/ppt/slides/slide3.xml").map_err(Error::Uri)?;
    let mut slide = BlobPart::new(
        slide_name,
        litchi_opc::constants::content_type::PML_SLIDE.to_owned(),
        br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#.to_vec(),
    );
    slide.rels_mut().try_add_relationship(
        rt::SLIDE_LAYOUT.to_owned(),
        "../slideLayouts/slideLayout1.xml".to_owned(),
        "rIdLayout".to_owned(),
        litchi_opc::TargetMode::Internal,
    )?;
    package.try_add_part(Box::new(slide))?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn add_unreferenced_media_part(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(bytes.to_vec())?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/unused.png").map_err(Error::Uri)?,
        ct::PNG.to_owned(),
        b"inert encrypted fixture payload".to_vec(),
    )))?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn mark_entry_encrypted(mut bytes: Vec<u8>, wanted: &str) -> TestResult<Vec<u8>> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .ok_or_else(|| Error::Invalid("fixture ZIP has no EOCD".into()))?;
    let count = u16::from_le_bytes(
        bytes[eocd + 10..eocd + 12]
            .try_into()
            .map_err(|_| Error::Invalid("fixture EOCD count is truncated".into()))?,
    ) as usize;
    let central_offset = u32::from_le_bytes(
        bytes[eocd + 16..eocd + 20]
            .try_into()
            .map_err(|_| Error::Invalid("fixture EOCD offset is truncated".into()))?,
    ) as usize;
    let mut cursor = central_offset;
    for _ in 0..count {
        let fixed = bytes
            .get(cursor..cursor + 46)
            .ok_or_else(|| Error::Invalid("fixture central record is truncated".into()))?;
        if fixed[..4] != 0x0201_4b50_u32.to_le_bytes() {
            return Err(Error::Invalid(
                "fixture central signature is invalid".into(),
            ));
        }
        let read_u16 =
            |offset: usize| u16::from_le_bytes([fixed[offset], fixed[offset + 1]]) as usize;
        let name_len = read_u16(28);
        let extra_len = read_u16(30);
        let comment_len = read_u16(32);
        let local_offset = u32::from_le_bytes(
            fixed[42..46]
                .try_into()
                .map_err(|_| Error::Invalid("fixture local offset is truncated".into()))?,
        ) as usize;
        let name = bytes
            .get(cursor + 46..cursor + 46 + name_len)
            .ok_or_else(|| Error::Invalid("fixture member name is truncated".into()))?;
        if name == wanted.as_bytes() {
            let central_flags = u16::from_le_bytes([bytes[cursor + 8], bytes[cursor + 9]]) | 1;
            bytes[cursor + 8..cursor + 10].copy_from_slice(&central_flags.to_le_bytes());
            let local_flags = u16::from_le_bytes([
                *bytes
                    .get(local_offset + 6)
                    .ok_or_else(|| Error::Invalid("fixture local flags are truncated".into()))?,
                *bytes
                    .get(local_offset + 7)
                    .ok_or_else(|| Error::Invalid("fixture local flags are truncated".into()))?,
            ]) | 1;
            bytes[local_offset + 6..local_offset + 8].copy_from_slice(&local_flags.to_le_bytes());
            return Ok(bytes);
        }
        cursor = cursor
            .checked_add(46 + name_len + extra_len + comment_len)
            .ok_or_else(|| Error::Invalid("fixture central cursor overflow".into()))?;
    }
    Err(Error::Invalid(format!(
        "fixture has no member named {wanted}"
    )))
}

fn add_media_relationship(bytes: &[u8], slide_member: &str) -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(bytes.to_vec())?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/source-only.png").map_err(Error::Uri)?,
        ct::PNG.to_owned(),
        b"source-only-media".to_vec(),
    )))?;
    package
        .get_part_mut(&PackURI::new(format!("/{slide_member}")).map_err(Error::Uri)?)?
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/source-only.png".to_owned(),
            "rIdSourceOnlyImage".to_owned(),
            litchi_opc::TargetMode::Internal,
        )?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn rewrite_archive<F>(bytes: &[u8], mut replace: F) -> TestResult<Vec<u8>>
where
    F: FnMut(&str, &[u8]) -> TestResult<Vec<u8>>,
{
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| Error::Invalid(format!("cannot index fixture ZIP: {error}")))?;
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for member in archive.file_names() {
        let payload = archive.read(member).map_err(|error| {
            Error::Invalid(format!("cannot read fixture member {member}: {error}"))
        })?;
        let changed = replace(member, &payload)?;
        writer.write_stored(member, &changed).map_err(|error| {
            Error::Invalid(format!("cannot rewrite fixture member {member}: {error}"))
        })?;
    }
    writer
        .finish_to_bytes()
        .map_err(|error| Error::Invalid(format!("cannot finish fixture ZIP: {error}")))
}

const SOURCE_PICTURE_MEDIA: &[u8] = b"source-only-media";
const DESTINATION_COLLISION_MEDIA: &[u8] = b"destination-collision-media";

#[test]
fn source_backed_cross_copy_copies_one_embedded_picture_closure() -> TestResult {
    let source_bytes = picture_source_fixture()?;
    let destination_bytes = destination_fixture_with_media_collision()?;
    let collision_uri = PackURI::new("/ppt/media/source-only.png").map_err(Error::Uri)?;
    let expected_unrelated: Vec<u8> = (0..=31).collect();

    let source = open_source(&source_bytes)?;
    let source_image = source
        .slide(0)
        .ok_or_else(|| Error::Invalid("picture source slide is missing".into()))?
        .read_image(0)?;
    assert_eq!(source_image.bytes(), SOURCE_PICTURE_MEDIA);
    assert_eq!(
        source_image.descriptor().target().content_type(),
        Some(ct::PNG)
    );

    for &(insertion_position, copied_position) in &[(1, 1), (2, 2)] {
        let output = publish(&source_bytes, &destination_bytes, 0, 1, insertion_position)?;
        let projection = slide_projection(&output)?;
        assert_eq!(projection.len(), 3);
        assert_eq!(projection[copied_position].1, "Source One");

        let (target_uri, payload, content_type) = copied_media(&output)?;
        assert_eq!(target_uri.as_str(), "/ppt/media/source-only-copy1.png");
        assert_eq!(payload, SOURCE_PICTURE_MEDIA);
        assert_eq!(content_type, ct::PNG);

        let reopened = open_source(&output)?;
        let copied_image = reopened
            .slide(copied_position)
            .ok_or_else(|| Error::Invalid("copied picture slide is missing".into()))?
            .read_image(0)?;
        assert_eq!(copied_image.bytes(), SOURCE_PICTURE_MEDIA);
        assert_eq!(
            copied_image.descriptor().target().content_type(),
            Some(ct::PNG)
        );
        assert_eq!(
            copied_image
                .descriptor()
                .target()
                .part_uri()
                .map(|uri| uri.as_str()),
            Some(target_uri.as_str())
        );

        let package = Package::from_vec(output.clone())?;
        let opc = package.opc()?;
        let collision = opc.get_part(&collision_uri)?;
        assert_eq!(collision.blob(), DESTINATION_COLLISION_MEDIA);
        assert_eq!(collision.content_type(), "image/jpeg");
        let unrelated_uri = PackURI::new("/ppt/media/unrelated.png").map_err(Error::Uri)?;
        let unrelated = opc.get_part(&unrelated_uri)?;
        assert_eq!(unrelated.blob(), expected_unrelated.as_slice());
        assert_eq!(unrelated.content_type(), ct::PNG);

        let repeat = publish(&source_bytes, &destination_bytes, 0, 1, insertion_position)?;
        assert_eq!(output, repeat);
    }
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_unsupported_picture_media_closures() -> TestResult {
    let valid = picture_source_fixture()?;
    let destination = destination_fixture()?;
    let external = external_picture_fixture(&valid)?;
    let linked = replace_text_member(
        &valid,
        "ppt/slides/slide1.xml",
        r#"r:embed="rIdSourceOnlyImage""#,
        r#"r:link="rIdSourceOnlyImage""#,
    )?;
    let missing_media = replace_text_member(
        &valid,
        "ppt/slides/_rels/slide1.xml.rels",
        r#"Target="../media/source-only.png""#,
        r#"Target="../media/missing.png""#,
    )?;
    let wrong_media_type = replace_text_member(
        &valid,
        CONTENT_TYPES_XML,
        ct::PNG,
        "application/octet-stream",
    )?;
    let outbound_media = add_stored_member(
        &valid,
        "ppt/media/_rels/source-only.png.rels",
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOutbound" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/outbound" TargetMode="External"/></Relationships>"#,
    )?;
    let multiple_pictures = append_direct_picture(&valid, "rIdSourceOnlyImage")?;
    let second_image = add_second_image_relationship(&valid)?;
    let multiple_images = replace_text_member(
        &second_image,
        "ppt/slides/slide1.xml",
        r#"<a:blip r:embed="rIdSourceOnlyImage"/>"#,
        r#"<a:blip r:embed="rIdSourceOnlyImage"/><a:blip r:embed="rIdSecondImage"/>"#,
    )?;
    let unreferenced_image = second_image;
    let mce = replace_text_member(
        &valid,
        "ppt/slides/slide1.xml",
        "<p:sld",
        "<p:sld xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"p14\"",
    )?;
    let chart = append_slide_element(&valid, "<p:graphicFrame/>")?;
    let table = append_slide_element(&valid, "<p:tbl/>")?;
    let notes = append_slide_element(&valid, "<p:notes/>")?;
    let ole = append_slide_element(&valid, "<p:oleObj/>")?;
    let opaque = add_stored_member(&valid, "ppt/opaque.bin", b"opaque")?;

    let _opened_source = open_source(&valid)?;
    let _opened_destination = open_editor(&destination)?;
    let cases = [
        (
            "external image",
            external,
            litchi_pptx::SlideCopyRefusal::UnsupportedRelationship,
        ),
        (
            "linked image",
            linked,
            litchi_pptx::SlideCopyRefusal::UnsupportedRelationship,
        ),
        (
            "missing media",
            missing_media,
            litchi_pptx::SlideCopyRefusal::AmbiguousTopology,
        ),
        (
            "wrong media type",
            wrong_media_type,
            litchi_pptx::SlideCopyRefusal::AmbiguousTopology,
        ),
        (
            "media outbound relationship",
            outbound_media,
            litchi_pptx::SlideCopyRefusal::UnsupportedRelationship,
        ),
        (
            "multiple images",
            multiple_images,
            litchi_pptx::SlideCopyRefusal::AmbiguousTopology,
        ),
        (
            "unreferenced image relationship",
            unreferenced_image,
            litchi_pptx::SlideCopyRefusal::UnsupportedRelationship,
        ),
        (
            "MCE",
            mce,
            litchi_pptx::SlideCopyRefusal::MarkupCompatibility,
        ),
        (
            "chart",
            chart,
            litchi_pptx::SlideCopyRefusal::UnknownSemanticSurface,
        ),
        (
            "table",
            table,
            litchi_pptx::SlideCopyRefusal::UnknownSemanticSurface,
        ),
        (
            "notes",
            notes,
            litchi_pptx::SlideCopyRefusal::UnknownSemanticSurface,
        ),
        (
            "OLE",
            ole,
            litchi_pptx::SlideCopyRefusal::UnknownSemanticSurface,
        ),
        (
            "opaque member",
            opaque,
            litchi_pptx::SlideCopyRefusal::UnknownPhysicalMember,
        ),
    ];
    for (label, candidate, expected) in cases {
        expect_plan_refusal(label, &candidate, &destination, expected)?;
    }
    publish(&multiple_pictures, &destination, 0, 1, 1)?;
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_duplicate_scene_trees_and_misplaced_picture_blips() -> TestResult
{
    let valid = picture_source_fixture()?;
    let destination = destination_fixture()?;
    let duplicate_scene_tree = replace_text_member(
        &valid,
        "ppt/slides/slide1.xml",
        "</p:spTree>",
        "</p:spTree><p:spTree/>",
    )?;
    expect_plan_refusal(
        "duplicate top-level shape trees",
        &duplicate_scene_tree,
        &destination,
        litchi_pptx::SlideCopyRefusal::AmbiguousTopology,
    )?;

    let without_direct_blip = replace_text_member(
        &valid,
        "ppt/slides/slide1.xml",
        r#"<p:blipFill><a:blip r:embed="rIdSourceOnlyImage"/>"#,
        "<p:blipFill>",
    )?;
    let misplaced_blip = replace_text_member(
        &without_direct_blip,
        "ppt/slides/slide1.xml",
        "</p:blipFill><p:spPr>",
        r#"</p:blipFill><a:blip r:embed="rIdSourceOnlyImage"/><p:spPr>"#,
    )?;
    expect_plan_refusal(
        "misplaced picture blip",
        &misplaced_blip,
        &destination,
        litchi_pptx::SlideCopyRefusal::UnknownSemanticSurface,
    )?;
    Ok(())
}

#[test]
fn source_backed_cross_copy_rewrites_image_layout_relationship_id_collision() -> TestResult {
    let destination = destination_fixture_with_media_collision()?;
    let destination_package = Package::from_vec(destination.clone())?;
    let destination_opc = destination_package.opc()?;
    let anchor_uri = PackURI::new("/ppt/slides/slide2.xml").map_err(Error::Uri)?;
    let anchor = destination_opc.get_part(&anchor_uri)?;
    let destination_layout_id = anchor
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::SLIDE_LAYOUT)
        .map(|relationship| relationship.r_id().to_owned())
        .ok_or_else(|| {
            Error::Invalid("destination anchor layout relationship is missing".into())
        })?;
    let source = picture_source_with_image_relationship_id(&destination_layout_id)?;

    let output = publish(&source, &destination, 0, 1, 1)?;
    assert_eq!(output, publish(&source, &destination, 0, 1, 1)?);
    let (_, slide_xml, rels_xml) = copied_slide_parts(&output)?;
    let embeds = xml_attribute_values(&slide_xml, "r:embed");
    assert_eq!(embeds.len(), 1);
    assert_ne!(embeds[0], destination_layout_id);
    assert!(rels_xml.contains(&format!("Id=\"{}\"", embeds[0])));
    let image_media = copied_image_media(&output)?;
    assert_eq!(image_media.len(), 1);
    assert_ne!(image_media[0].1.as_str(), "/ppt/media/source-only.png");
    assert_eq!(image_media[0].2, b"source-only-media");
    assert_eq!(image_media[0].3, ct::PNG);
    Package::from_vec(output.clone())?;
    open_source(&output)?;
    Ok(())
}

#[test]
fn source_backed_cross_copy_reserves_before_reading_picture_media() -> TestResult {
    let source_bytes = picture_source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let (media_start, media_end) =
        zip_member_data_range(&source_bytes, "ppt/media/source-only.png")?;
    let source_slide_bytes = ArchiveReader::new(&source_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index source: {error}")))?
        .read("ppt/slides/slide1.xml")
        .map_err(|error| Error::Invalid(format!("cannot read source slide: {error}")))?;
    let destination_presentation_bytes = ArchiveReader::new(&destination_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index destination: {error}")))?
        .read(PRESENTATION_XML)
        .map_err(|error| {
            Error::Invalid(format!("cannot read destination presentation: {error}"))
        })?;
    let media_bytes = ArchiveReader::new(&source_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index source media: {error}")))?
        .read("ppt/media/source-only.png")
        .map_err(|error| Error::Invalid(format!("cannot read source media: {error}")))?;
    let staged_bytes = source_slide_bytes
        .len()
        .checked_add(destination_presentation_bytes.len())
        .and_then(|bytes| bytes.checked_add(media_bytes.len()))
        .and_then(|bytes| bytes.checked_add(256))
        .ok_or_else(|| Error::Invalid("picture staging size overflow".into()))?;
    let execution_limits = single_execution_limits()?;

    let measure_budget = Budget::root(
        "pptx-picture-media-measure",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (measure_cancel, measure_token) = CancellationSource::pair();
    let measure_context =
        ExecutionContext::new(measure_budget.clone(), measure_token, execution_limits);
    let measured_source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(source_bytes.clone())),
        ReadLimits::default(),
        measure_context.clone(),
    )?;
    let measured_editor =
        SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
            Arc::new(OwnedSource::new(destination_bytes.clone())),
            ReadLimits::default(),
            measure_context,
        )?;
    let baseline = measure_budget.used(Resource::Memory);
    drop(measured_editor);
    drop(measured_source);
    measure_cancel.cancel();
    let memory_limit = baseline
        .checked_add(u64::try_from(staged_bytes).map_err(|error| {
            Error::Invalid(format!("picture staging size does not fit u64: {error}"))
        })?)
        .and_then(|bytes| bytes.checked_sub(1))
        .ok_or_else(|| Error::Invalid("picture memory limit overflow".into()))?;

    let overlaps = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let source_adapter = RangeCountingSource::new(
        source_bytes.clone(),
        media_start,
        media_end,
        Arc::clone(&overlaps),
    );
    let budget = Budget::root(
        "pptx-picture-media-budget",
        Limits::new(
            memory_limit,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_cancel, token) = CancellationSource::pair();
    let context = ExecutionContext::new(budget, token, execution_limits);
    let source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(source_adapter),
        ReadLimits::default(),
        context.clone(),
    )?;
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(destination_bytes)),
        ReadLimits::default(),
        context,
    )?;
    let error = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
        Ok(_) => panic!("operation reservation must fail before image payload read"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::Execution(
            litchi_core::ExecutionError::ResourceLimit(limit)
        )) if limit.resource == Resource::Memory
    ));
    assert_eq!(
        overlaps.load(std::sync::atomic::Ordering::SeqCst),
        0,
        "budget refusal must not request bytes overlapping the image payload"
    );
    Ok(())
}

fn expect_plan_refusal(
    label: &str,
    source_bytes: &[u8],
    destination_bytes: &[u8],
    expected: litchi_pptx::SlideCopyRefusal,
) -> TestResult {
    let source = open_source(source_bytes)?;
    let editor = open_editor(destination_bytes)?;
    let error = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
        Ok(_) => panic!("fixture must reach the cross-copy planner"),
        Err(error) => error,
    };
    match error {
        Error::SlideCopyPlan { kind, .. } => {
            assert_eq!(kind, expected, "unexpected refusal for {label}");
        },
        other => panic!("unexpected error for {label}: {other:?}"),
    }
    Ok(())
}

fn picture_source_with_image_relationship_id(relationship_id: &str) -> TestResult<Vec<u8>> {
    let bytes = picture_source_fixture()?;
    let package = OpcPackage::from_vec(bytes.clone())?;
    let slide_uri = PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Uri)?;
    let source_layout_id = package
        .get_part(&slide_uri)?
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::SLIDE_LAYOUT)
        .map(|relationship| relationship.r_id().to_owned())
        .ok_or_else(|| Error::Invalid("source layout relationship is missing".into()))?;
    let source_layout_id = format!(r#"Id="{source_layout_id}""#);
    let bytes = replace_text_member(
        &bytes,
        "ppt/slides/_rels/slide1.xml.rels",
        &source_layout_id,
        r#"Id="rIdSourceLayoutForTest""#,
    )?;
    let embedded = format!(r#"r:embed="{relationship_id}""#);
    let bytes = replace_text_member(
        &bytes,
        "ppt/slides/slide1.xml",
        r#"r:embed="rIdSourceOnlyImage""#,
        &embedded,
    )?;
    let relation_id = format!(r#"Id="{relationship_id}""#);
    replace_text_member(
        &bytes,
        "ppt/slides/_rels/slide1.xml.rels",
        r#"Id="rIdSourceOnlyImage""#,
        &relation_id,
    )
}

fn single_execution_limits() -> TestResult<ExecutionLimits> {
    ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("nonzero workers"),
        std::num::NonZeroUsize::new(1).expect("nonzero tasks"),
        std::num::NonZeroU64::new(u64::MAX).expect("nonzero in-flight bytes"),
        0,
    )
    .map_err(|error| Error::Invalid(error.to_string()))
}

#[derive(Clone)]
struct RangeCountingSource {
    source: OwnedSource,
    range_start: u64,
    range_end: u64,
    overlaps: Arc<std::sync::atomic::AtomicUsize>,
}

impl RangeCountingSource {
    fn new(
        bytes: Vec<u8>,
        range_start: u64,
        range_end: u64,
        overlaps: Arc<std::sync::atomic::AtomicUsize>,
    ) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            range_start,
            range_end,
            overlaps,
        }
    }
}

impl ReadAt for RangeCountingSource {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let request_end = offset.saturating_add(output.len() as u64);
        if offset < self.range_end && request_end > self.range_start {
            self.overlaps
                .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        }
        self.source.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.source.version()
    }
}

fn zip_member_data_range(bytes: &[u8], wanted: &str) -> TestResult<(u64, u64)> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .ok_or_else(|| Error::Invalid("fixture ZIP has no EOCD".into()))?;
    let count = u16::from_le_bytes(
        bytes
            .get(eocd + 10..eocd + 12)
            .ok_or_else(|| Error::Invalid("fixture EOCD count is truncated".into()))?
            .try_into()
            .map_err(|_| Error::Invalid("fixture EOCD count is truncated".into()))?,
    ) as usize;
    let central_offset = u32::from_le_bytes(
        bytes
            .get(eocd + 16..eocd + 20)
            .ok_or_else(|| Error::Invalid("fixture EOCD offset is truncated".into()))?
            .try_into()
            .map_err(|_| Error::Invalid("fixture EOCD offset is truncated".into()))?,
    ) as usize;
    let mut cursor = central_offset;
    for _ in 0..count {
        let fixed = bytes
            .get(cursor..cursor + 46)
            .ok_or_else(|| Error::Invalid("fixture central record is truncated".into()))?;
        if fixed[..4] != 0x0201_4b50_u32.to_le_bytes() {
            return Err(Error::Invalid(
                "fixture central signature is invalid".into(),
            ));
        }
        let read_u16 =
            |offset: usize| u16::from_le_bytes([fixed[offset], fixed[offset + 1]]) as usize;
        let compressed_size = u32::from_le_bytes(
            fixed[20..24]
                .try_into()
                .map_err(|_| Error::Invalid("fixture compressed size is truncated".into()))?,
        ) as u64;
        let name_len = read_u16(28);
        let extra_len = read_u16(30);
        let comment_len = read_u16(32);
        let local_offset = u32::from_le_bytes(
            fixed[42..46]
                .try_into()
                .map_err(|_| Error::Invalid("fixture local offset is truncated".into()))?,
        ) as usize;
        let name = bytes
            .get(cursor + 46..cursor + 46 + name_len)
            .ok_or_else(|| Error::Invalid("fixture member name is truncated".into()))?;
        if name == wanted.as_bytes() {
            let local = bytes
                .get(local_offset..local_offset + 30)
                .ok_or_else(|| Error::Invalid("fixture local record is truncated".into()))?;
            if local[..4] != 0x0403_4b50_u32.to_le_bytes() {
                return Err(Error::Invalid("fixture local signature is invalid".into()));
            }
            let local_name_len = u16::from_le_bytes([local[26], local[27]]) as usize;
            let local_extra_len = u16::from_le_bytes([local[28], local[29]]) as usize;
            let start = local_offset
                .checked_add(30)
                .and_then(|value| value.checked_add(local_name_len))
                .and_then(|value| value.checked_add(local_extra_len))
                .ok_or_else(|| Error::Invalid("fixture local payload offset overflow".into()))?;
            let end = start
                .checked_add(usize::try_from(compressed_size).map_err(|_| {
                    Error::Invalid("fixture compressed size does not fit usize".into())
                })?)
                .ok_or_else(|| Error::Invalid("fixture local payload range overflow".into()))?;
            if end > bytes.len() {
                return Err(Error::Invalid("fixture local payload is truncated".into()));
            }
            return Ok((
                u64::try_from(start).map_err(|_| {
                    Error::Invalid("fixture payload offset does not fit u64".into())
                })?,
                u64::try_from(end)
                    .map_err(|_| Error::Invalid("fixture payload end does not fit u64".into()))?,
            ));
        }
        cursor = cursor
            .checked_add(46 + name_len + extra_len + comment_len)
            .ok_or_else(|| Error::Invalid("fixture central cursor overflow".into()))?;
    }
    Err(Error::Invalid(format!(
        "fixture has no member named {wanted}"
    )))
}

fn misdeclare_zip_member_uncompressed_size(
    bytes: &[u8],
    wanted: &str,
    increment: u32,
) -> TestResult<Vec<u8>> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .ok_or_else(|| Error::Invalid("fixture ZIP has no EOCD".into()))?;
    let count = u16::from_le_bytes(
        bytes
            .get(eocd + 10..eocd + 12)
            .ok_or_else(|| Error::Invalid("fixture EOCD count is truncated".into()))?
            .try_into()
            .map_err(|_| Error::Invalid("fixture EOCD count is truncated".into()))?,
    ) as usize;
    let central_offset = u32::from_le_bytes(
        bytes
            .get(eocd + 16..eocd + 20)
            .ok_or_else(|| Error::Invalid("fixture EOCD offset is truncated".into()))?
            .try_into()
            .map_err(|_| Error::Invalid("fixture EOCD offset is truncated".into()))?,
    ) as usize;
    let mut cursor = central_offset;
    for _ in 0..count {
        let fixed = bytes
            .get(cursor..cursor + 46)
            .ok_or_else(|| Error::Invalid("fixture central record is truncated".into()))?;
        if fixed[..4] != 0x0201_4b50_u32.to_le_bytes() {
            return Err(Error::Invalid(
                "fixture central signature is invalid".into(),
            ));
        }
        let read_u16 =
            |offset: usize| u16::from_le_bytes([fixed[offset], fixed[offset + 1]]) as usize;
        let name_len = read_u16(28);
        let extra_len = read_u16(30);
        let comment_len = read_u16(32);
        let name = bytes
            .get(cursor + 46..cursor + 46 + name_len)
            .ok_or_else(|| Error::Invalid("fixture member name is truncated".into()))?;
        if name == wanted.as_bytes() {
            let declared = u32::from_le_bytes(
                fixed[24..28]
                    .try_into()
                    .map_err(|_| Error::Invalid("fixture size is truncated".into()))?,
            )
            .checked_add(increment)
            .ok_or_else(|| Error::Invalid("fixture declared size overflow".into()))?;
            let local_offset = u32::from_le_bytes(
                fixed[42..46]
                    .try_into()
                    .map_err(|_| Error::Invalid("fixture local offset is truncated".into()))?,
            ) as usize;
            let local = bytes
                .get(local_offset..local_offset + 30)
                .ok_or_else(|| Error::Invalid("fixture local record is truncated".into()))?;
            if local[..4] != 0x0403_4b50_u32.to_le_bytes() {
                return Err(Error::Invalid("fixture local signature is invalid".into()));
            }
            let mut output = bytes.to_vec();
            output[cursor + 24..cursor + 28].copy_from_slice(&declared.to_le_bytes());
            output[local_offset + 22..local_offset + 26].copy_from_slice(&declared.to_le_bytes());
            return Ok(output);
        }
        cursor = cursor
            .checked_add(46 + name_len + extra_len + comment_len)
            .ok_or_else(|| Error::Invalid("fixture central cursor overflow".into()))?;
    }
    Err(Error::Invalid(format!(
        "fixture has no member named {wanted}"
    )))
}

fn picture_source_fixture() -> TestResult<Vec<u8>> {
    let with_media = add_media_relationship(&source_fixture()?, "ppt/slides/slide1.xml")?;
    append_direct_picture(&with_media, "rIdSourceOnlyImage")
}

fn destination_fixture_with_media_collision() -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(destination_fixture()?)?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/source-only.png").map_err(Error::Uri)?,
        "image/jpeg".to_owned(),
        DESTINATION_COLLISION_MEDIA.to_vec(),
    )))?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn append_direct_picture(bytes: &[u8], relationship_id: &str) -> TestResult<Vec<u8>> {
    rewrite_member(bytes, "ppt/slides/slide1.xml", |payload| {
        let xml = std::str::from_utf8(payload)
            .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
        let marker = "</p:spTree>";
        let insertion = xml
            .rfind(marker)
            .ok_or_else(|| Error::Invalid("picture fixture has no shape tree".into()))?;
        let picture = format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr id="42" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
        );
        Ok(format!("{}{picture}{}", &xml[..insertion], &xml[insertion..]).into_bytes())
    })
}

fn add_second_image_relationship(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(bytes.to_vec())?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/second-source.png").map_err(Error::Uri)?,
        ct::PNG.to_owned(),
        b"second-source-media".to_vec(),
    )))?;
    package
        .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Uri)?)?
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/second-source.png".to_owned(),
            "rIdSecondImage".to_owned(),
            litchi_opc::TargetMode::Internal,
        )?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn external_picture_fixture(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let bytes = replace_text_member(
        bytes,
        "ppt/slides/slide1.xml",
        r#"r:embed="rIdSourceOnlyImage""#,
        r#"r:embed="rIdExternalImage""#,
    )?;
    let bytes = replace_text_member(
        &bytes,
        "ppt/slides/_rels/slide1.xml.rels",
        "rIdSourceOnlyImage",
        "rIdExternalImage",
    )?;
    replace_text_member(
        &bytes,
        "ppt/slides/_rels/slide1.xml.rels",
        r#"Target="../media/source-only.png""#,
        r#"Target="https://example.invalid/picture.png" TargetMode="External""#,
    )
}

fn append_slide_element(bytes: &[u8], element: &str) -> TestResult<Vec<u8>> {
    rewrite_member(bytes, "ppt/slides/slide1.xml", |payload| {
        let xml = std::str::from_utf8(payload)
            .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
        let marker = "</p:spTree>";
        let insertion = xml
            .rfind(marker)
            .ok_or_else(|| Error::Invalid("surface fixture has no shape tree".into()))?;
        Ok(format!("{}{element}{}", &xml[..insertion], &xml[insertion..]).into_bytes())
    })
}

fn replace_text_member(bytes: &[u8], member: &str, from: &str, to: &str) -> TestResult<Vec<u8>> {
    rewrite_member(bytes, member, |payload| {
        let text = std::str::from_utf8(payload)
            .map_err(|error| Error::Invalid(format!("fixture member is not UTF-8: {error}")))?;
        if !text.contains(from) {
            return Err(Error::Invalid(format!(
                "fixture member {member} lacks replacement text"
            )));
        }
        Ok(text.replace(from, to).into_bytes())
    })
}

fn copied_slide_uri(bytes: &[u8]) -> TestResult<PackURI> {
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| Error::Invalid(format!("cannot index copied ZIP: {error}")))?;
    let marker = b"name=\"Source One\"";
    let mut found = None;
    for member in archive.file_names() {
        if !member.starts_with("ppt/slides/slide") || !member.ends_with(".xml") {
            continue;
        }
        let payload = archive.read(member).map_err(|error| {
            Error::Invalid(format!("cannot read copied slide {member}: {error}"))
        })?;
        if payload.windows(marker.len()).any(|window| window == marker) {
            if found.is_some() {
                return Err(Error::Invalid(
                    "copied ZIP contains multiple Source One slides".into(),
                ));
            }
            found = Some(member.to_owned());
        }
    }
    let member =
        found.ok_or_else(|| Error::Invalid("copied ZIP has no Source One slide".into()))?;
    PackURI::new(format!("/{member}")).map_err(Error::Uri)
}

fn copied_media(bytes: &[u8]) -> TestResult<(PackURI, Vec<u8>, String)> {
    let slide_uri = copied_slide_uri(bytes)?;
    let package = Package::from_vec(bytes.to_vec())?;
    let opc = package.opc()?;
    let slide = opc.get_part(&slide_uri)?;
    let image_relationships = slide
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE))
        .collect::<Vec<_>>();
    if image_relationships.len() != 1 {
        return Err(Error::Invalid(format!(
            "copied slide has {} image relationships",
            image_relationships.len()
        )));
    }
    let target_uri = image_relationships[0].target_partname()?;
    if !target_uri.as_str().starts_with("/ppt/media/") {
        return Err(Error::Invalid(
            "copied image relationship leaves /ppt/media".into(),
        ));
    }
    let media = opc.get_part(&target_uri)?;
    if !media.rels().is_empty() {
        return Err(Error::Invalid(
            "copied image target has outbound relationships".into(),
        ));
    }
    Ok((
        target_uri,
        media.blob().to_vec(),
        media.content_type().to_owned(),
    ))
}

fn copied_image_media(bytes: &[u8]) -> TestResult<Vec<(String, PackURI, Vec<u8>, String)>> {
    let (slide_uri, slide_xml, rels_xml) = copied_slide_parts(bytes)?;
    let relationship_ids = xml_attribute_values(&slide_xml, "r:embed");
    let package = Package::from_vec(bytes.to_vec())?;
    let opc = package.opc()?;
    let slide = opc.get_part(&slide_uri)?;
    let mut images = Vec::with_capacity(relationship_ids.len());
    for relationship_id in relationship_ids {
        let relationship = slide
            .rels()
            .iter()
            .find(|relationship| relationship.r_id() == relationship_id.as_str())
            .ok_or_else(|| {
                Error::Invalid(format!(
                    "copied slide relationship {relationship_id} is missing"
                ))
            })?;
        if !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE) {
            return Err(Error::Invalid(format!(
                "copied embed {relationship_id} is not an image relationship"
            )));
        }
        let target_uri = relationship.target_partname()?;
        if !target_uri.as_str().starts_with("/ppt/media/") {
            return Err(Error::Invalid(
                "copied image relationship leaves /ppt/media".into(),
            ));
        }
        let target_name = target_uri
            .as_str()
            .rsplit('/')
            .next()
            .ok_or_else(|| Error::Invalid("copied image target has no leaf name".into()))?;
        let relationship_target = relationship_target_for_id(&rels_xml, &relationship_id)?;
        assert!(relationship_target.ends_with(target_name));
        let media = opc.get_part(&target_uri)?;
        if !media.rels().is_empty() {
            return Err(Error::Invalid(
                "copied image target has outbound relationships".into(),
            ));
        }
        images.push((
            relationship_id,
            target_uri,
            media.blob().to_vec(),
            media.content_type().to_owned(),
        ));
    }
    Ok(images)
}

fn image_relationship_mapping(
    slide_xml: &str,
    rels_xml: &str,
) -> TestResult<Vec<(String, String)>> {
    let mut mapping = Vec::new();
    for relationship_id in xml_attribute_values(slide_xml, "r:embed") {
        mapping.push((
            relationship_id.clone(),
            relationship_target_for_id(rels_xml, &relationship_id)?,
        ));
    }
    Ok(mapping)
}

#[test]
fn source_backed_cross_copy_copies_multiple_embedded_images_in_source_order() -> TestResult {
    let source = multi_picture_source_fixture(false)?;
    let destination = destination_fixture_with_media_collision()?;
    let output = publish(&source, &destination, 0, 1, 1)?;
    assert_eq!(output, publish(&source, &destination, 0, 1, 1)?);

    let (_, slide_xml, rels_xml) = copied_slide_parts(&output)?;
    assert_image_targets(
        &slide_xml,
        &rels_xml,
        &[
            "source-only-copy1.png",
            "second-source-copy1.png",
            "third-source-copy1.jpg",
        ],
    )?;

    let archive = ArchiveReader::new(&output)
        .map_err(|error| Error::Invalid(format!("cannot index copied ZIP: {error}")))?;
    assert_eq!(
        archive
            .read("ppt/media/source-only-copy1.png")
            .expect("multi-image output must contain the copied first image"),
        b"source-only-media"
    );
    assert_eq!(
        archive
            .read("ppt/media/second-source-copy1.png")
            .expect("multi-image output must contain the copied second image"),
        b"second-source-media"
    );
    assert_eq!(
        archive
            .read("ppt/media/third-source-copy1.jpg")
            .expect("multi-image output must contain the copied third image"),
        b"third-source-media"
    );
    assert_eq!(
        archive
            .read("ppt/media/unrelated.png")
            .expect("multi-image output must preserve unrelated media"),
        (0..=31).collect::<Vec<_>>()
    );

    let package = Package::from_vec(output.clone())?;
    let opc = package.opc()?;
    for (member, content_type) in [
        ("ppt/media/source-only-copy1.png", ct::PNG),
        ("ppt/media/second-source-copy1.png", ct::PNG),
        ("ppt/media/third-source-copy1.jpg", "image/jpeg"),
    ] {
        let uri = PackURI::new(format!("/{member}")).map_err(Error::Uri)?;
        let part = opc.get_part(&uri)?;
        assert_eq!(part.content_type(), content_type);
        assert!(part.rels().is_empty());
    }
    open_source(&output)?;
    Ok(())
}

#[test]
fn source_backed_cross_copy_copies_shared_image_once_with_coherent_relationships() -> TestResult {
    let source = shared_picture_source_fixture()?;
    let destination = destination_fixture_with_media_collision()?;
    let output = publish(&source, &destination, 0, 1, 1)?;

    let (_, slide_xml, rels_xml) = copied_slide_parts(&output)?;
    assert_image_targets(
        &slide_xml,
        &rels_xml,
        &["source-only-copy1.png", "source-only-copy1.png"],
    )?;
    let embeds = xml_attribute_values(&slide_xml, "r:embed");
    assert_eq!(embeds[0], embeds[1]);

    let image_media = copied_image_media(&output)?;
    assert_eq!(image_media.len(), 2);
    assert_eq!(image_media[0].1, image_media[1].1);
    assert_eq!(image_media[0].2, b"source-only-media");
    assert_eq!(output, publish(&source, &destination, 0, 1, 1)?);
    Package::from_vec(output.clone())?;
    open_source(&output)?;
    Ok(())
}

#[test]
fn source_backed_cross_copy_allocates_case_equivalent_media_collision() -> TestResult {
    let source = picture_source_fixture()?;
    let destination = destination_fixture_with_case_equivalent_media_collision()?;
    let output = publish(&source, &destination, 0, 1, 1)?;

    let image_media = copied_image_media(&output)?;
    assert_eq!(image_media.len(), 1);
    let target = image_media[0].1.as_str().to_ascii_lowercase();
    assert_ne!(target, "/ppt/media/source-only.png");
    assert_ne!(target, "/ppt/media/source-only-copy1.png");
    assert_eq!(image_media[0].2, b"source-only-media");
    assert_eq!(image_media[0].3, ct::PNG);

    let archive = ArchiveReader::new(&output)
        .map_err(|error| Error::Invalid(format!("cannot index copied ZIP: {error}")))?;
    assert_eq!(
        archive
            .read("ppt/media/SOURCE-ONLY-COPY1.PNG")
            .expect("case-equivalent collision member must be preserved"),
        b"case-equivalent-media"
    );
    assert_eq!(
        archive
            .read("ppt/media/source-only.png")
            .expect("original destination collision member must be preserved"),
        DESTINATION_COLLISION_MEDIA
    );
    assert_eq!(output, publish(&source, &destination, 0, 1, 1)?);
    Package::from_vec(output.clone())?;
    open_source(&output)?;
    Ok(())
}

#[test]
fn source_backed_cross_copy_relationship_id_rewrite_is_source_order_deterministic() -> TestResult {
    let destination = destination_fixture_with_media_collision()?;
    let first = publish(&multi_picture_source_fixture(false)?, &destination, 0, 1, 1)?;
    let reordered = publish(&multi_picture_source_fixture(true)?, &destination, 0, 1, 1)?;
    let (_, first_slide, first_rels) = copied_slide_parts(&first)?;
    let (_, reordered_slide, reordered_rels) = copied_slide_parts(&reordered)?;
    assert_eq!(
        image_relationship_mapping(&first_slide, &first_rels)?,
        image_relationship_mapping(&reordered_slide, &reordered_rels)?
    );
    Ok(())
}

#[test]
fn source_backed_cross_copy_aggregates_image_reservation_before_payload_reads() -> TestResult {
    let source_bytes = multi_picture_source_fixture(false)?;
    let destination_bytes = destination_fixture()?;
    let media_members = [
        "ppt/media/source-only.png",
        "ppt/media/second-source.png",
        "ppt/media/third-source.jpg",
    ];
    let source_archive = ArchiveReader::new(&source_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index source: {error}")))?;
    let mut ranges = Vec::new();
    let mut media_size = 0usize;
    for member in media_members {
        ranges.push(zip_member_data_range(&source_bytes, member)?);
        media_size = media_size
            .checked_add(
                source_archive
                    .read(member)
                    .expect("multi-image fixture media must be readable")
                    .len(),
            )
            .ok_or_else(|| Error::Invalid("aggregate media size overflow".into()))?;
    }
    let source_slide_size = source_archive
        .read("ppt/slides/slide1.xml")
        .expect("multi-image fixture slide must be readable")
        .len();
    let destination_size = ArchiveReader::new(&destination_bytes)
        .map_err(|error| Error::Invalid(format!("cannot index destination: {error}")))?
        .read(PRESENTATION_XML)
        .expect("destination presentation fixture must be readable")
        .len();
    let staged_size = source_slide_size
        .checked_add(destination_size)
        .and_then(|size| size.checked_add(media_size))
        .and_then(|size| size.checked_add(256))
        .ok_or_else(|| Error::Invalid("aggregate staging size overflow".into()))?;

    let measure_budget = Budget::root(
        "pptx-picture-aggregate-measure",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (measure_cancel, measure_token) = CancellationSource::pair();
    let measure_context = ExecutionContext::new(
        measure_budget.clone(),
        measure_token,
        single_execution_limits()?,
    );
    let measured_source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(source_bytes.clone())),
        ReadLimits::default(),
        measure_context.clone(),
    )?;
    let measured_editor =
        SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
            Arc::new(OwnedSource::new(destination_bytes.clone())),
            ReadLimits::default(),
            measure_context,
        )?;
    let baseline = measure_budget.used(Resource::Memory);
    drop(measured_editor);
    drop(measured_source);
    measure_cancel.cancel();

    let memory_limit =
        baseline
            .checked_add(u64::try_from(staged_size).map_err(|error| {
                Error::Invalid(format!("staging size does not fit u64: {error}"))
            })?)
            .and_then(|size| size.checked_sub(1))
            .ok_or_else(|| Error::Invalid("aggregate memory limit overflow".into()))?;
    let overlaps = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let source = MultiRangeCountingSource::new(source_bytes, ranges, Arc::clone(&overlaps));
    let budget = Budget::root(
        "pptx-picture-aggregate-budget",
        Limits::new(
            memory_limit,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_cancel, token) = CancellationSource::pair();
    let context = ExecutionContext::new(budget, token, single_execution_limits()?);
    let source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(source),
        ReadLimits::default(),
        context.clone(),
    )?;
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(destination_bytes)),
        ReadLimits::default(),
        context,
    )?;
    let error = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
        Ok(_) => panic!("aggregate image reservation must fail before payload reads"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::Execution(
            litchi_core::ExecutionError::ResourceLimit(limit)
        )) if limit.resource == Resource::Memory
    ));
    assert_eq!(
        overlaps.load(std::sync::atomic::Ordering::SeqCst),
        0,
        "reservation refusal must not request any image payload range"
    );
    Ok(())
}

#[test]
fn source_backed_cross_copy_preserves_media_read_errors() -> TestResult {
    let source_bytes = picture_source_fixture()?;
    let (range_start, range_end) =
        zip_member_data_range(&source_bytes, "ppt/media/source-only.png")?;
    let overlaps = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let armed = Arc::new(std::sync::atomic::AtomicBool::new(false));
    let source = FailingRangeSource::new(
        source_bytes,
        range_start,
        range_end,
        Arc::clone(&overlaps),
        Arc::clone(&armed),
    );
    let budget = Budget::root(
        "pptx-picture-read-error",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancel, token) = CancellationSource::pair();
    let context = ExecutionContext::new(budget, token, single_execution_limits()?);
    let source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(source),
        ReadLimits::default(),
        context.clone(),
    )?;
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(destination_fixture()?)),
        ReadLimits::default(),
        context,
    )?;
    armed.store(true, std::sync::atomic::Ordering::SeqCst);
    let error = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
        Ok(_) => panic!("media read failure must not produce a successful plan"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::IoError(error))
            if error.kind() == io::ErrorKind::UnexpectedEof
    ));
    assert!(overlaps.load(std::sync::atomic::Ordering::SeqCst) > 0);
    Ok(())
}

#[test]
fn source_backed_cross_copy_rejects_actual_source_size_mismatch() -> TestResult {
    let bytes = misdeclare_zip_member_uncompressed_size(
        &picture_source_fixture()?,
        "ppt/media/source-only.png",
        1,
    )?;
    let source = open_source(&bytes)?;
    let editor = open_editor(&destination_fixture()?)?;
    let error = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
        Ok(_) => panic!("declared media size mismatch must fail during deferred data read"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::ZipError(message)) if message.contains("size")
    ));
    Ok(())
}

fn multi_picture_source_fixture(reverse_relationship_order: bool) -> TestResult<Vec<u8>> {
    let mut bytes = picture_source_fixture()?;
    let entries: [(&str, &str, &str, &[u8]); 2] = if reverse_relationship_order {
        [
            (
                "third-source.jpg",
                "rIdThirdImage",
                "image/jpeg",
                b"third-source-media",
            ),
            (
                "second-source.png",
                "rIdSecondImage",
                ct::PNG,
                b"second-source-media",
            ),
        ]
    } else {
        [
            (
                "second-source.png",
                "rIdSecondImage",
                ct::PNG,
                b"second-source-media",
            ),
            (
                "third-source.jpg",
                "rIdThirdImage",
                "image/jpeg",
                b"third-source-media",
            ),
        ]
    };
    for (member, relationship_id, content_type, payload) in entries {
        bytes = add_embedded_image_relationship(
            &bytes,
            member,
            relationship_id,
            content_type,
            payload,
        )?;
    }
    bytes = append_picture_element(&bytes, "rIdSecondImage", 43, "Second Picture")?;
    append_picture_element(&bytes, "rIdThirdImage", 44, "Third Picture")
}

fn shared_picture_source_fixture() -> TestResult<Vec<u8>> {
    append_picture_element(
        &picture_source_fixture()?,
        "rIdSourceOnlyImage",
        43,
        "Shared Picture",
    )
}

fn add_embedded_image_relationship(
    bytes: &[u8],
    member: &str,
    relationship_id: &str,
    content_type: &str,
    payload: &[u8],
) -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(bytes.to_vec())?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(format!("/ppt/media/{member}")).map_err(Error::Uri)?,
        content_type.to_owned(),
        payload.to_vec(),
    )))?;
    package
        .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Uri)?)?
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            format!("../media/{member}"),
            relationship_id.to_owned(),
            litchi_opc::TargetMode::Internal,
        )?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn append_picture_element(
    bytes: &[u8],
    relationship_id: &str,
    shape_id: u32,
    name: &str,
) -> TestResult<Vec<u8>> {
    let element = format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
    );
    append_slide_element(bytes, &element)
}

fn destination_fixture_with_case_equivalent_media_collision() -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(destination_fixture_with_media_collision()?)?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/SOURCE-ONLY-COPY1.PNG").map_err(Error::Uri)?,
        ct::PNG.to_owned(),
        b"case-equivalent-media".to_vec(),
    )))?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn copied_slide_parts(bytes: &[u8]) -> TestResult<(PackURI, String, String)> {
    let slide_uri = copied_slide_uri(bytes)?;
    let member = slide_uri.as_str().trim_start_matches('/');
    let (directory, file_name) = member
        .rsplit_once('/')
        .ok_or_else(|| Error::Invalid("copied slide URI has no parent".into()))?;
    let relationships_member = format!("{directory}/_rels/{file_name}.rels");
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| Error::Invalid(format!("cannot index copied ZIP: {error}")))?;
    let slide_xml = String::from_utf8(
        archive
            .read(member)
            .map_err(|error| Error::Invalid(format!("cannot read copied slide: {error}")))?,
    )
    .map_err(|error| Error::Invalid(format!("copied slide is not UTF-8: {error}")))?;
    let rels_xml = String::from_utf8(archive.read(&relationships_member).map_err(|error| {
        Error::Invalid(format!("cannot read copied slide relationships: {error}"))
    })?)
    .map_err(|error| Error::Invalid(format!("copied relationships are not UTF-8: {error}")))?;
    Ok((slide_uri, slide_xml, rels_xml))
}

fn xml_attribute_values(xml: &str, attribute: &str) -> Vec<String> {
    let marker = format!("{attribute}=\"");
    let mut values = Vec::new();
    let mut remaining = xml;
    while let Some(start) = remaining.find(&marker) {
        let value_start = start + marker.len();
        let Some(end) = remaining[value_start..].find('"') else {
            break;
        };
        values.push(remaining[value_start..value_start + end].to_owned());
        remaining = &remaining[value_start + end + 1..];
    }
    values
}

fn relationship_target_for_id(rels_xml: &str, relationship_id: &str) -> TestResult<String> {
    let id_marker = format!("Id=\"{relationship_id}\"");
    let id_offset = rels_xml
        .find(&id_marker)
        .ok_or_else(|| Error::Invalid(format!("missing relationship {relationship_id}")))?;
    let record_start = rels_xml[..id_offset]
        .rfind("<Relationship")
        .ok_or_else(|| Error::Invalid("relationship record start is missing".into()))?;
    let record_end = id_offset
        + rels_xml[id_offset..]
            .find("/>")
            .ok_or_else(|| Error::Invalid("relationship record end is missing".into()))?;
    let record = &rels_xml[record_start..record_end];
    let target_marker = "Target=\"";
    let target_start = record
        .find(target_marker)
        .ok_or_else(|| Error::Invalid("relationship target is missing".into()))?
        + target_marker.len();
    let target_end = target_start
        + record[target_start..]
            .find('"')
            .ok_or_else(|| Error::Invalid("relationship target is unterminated".into()))?;
    Ok(record[target_start..target_end].to_owned())
}

fn assert_image_targets(slide_xml: &str, rels_xml: &str, expected_targets: &[&str]) -> TestResult {
    let relationship_ids = xml_attribute_values(slide_xml, "r:embed");
    assert_eq!(relationship_ids.len(), expected_targets.len());
    let mut targets = Vec::with_capacity(relationship_ids.len());
    for relationship_id in &relationship_ids {
        targets.push(relationship_target_for_id(rels_xml, relationship_id)?);
    }
    for (target, expected) in targets.iter().zip(expected_targets) {
        assert!(
            target.ends_with(expected),
            "image target {target:?} does not end with {expected:?}"
        );
    }
    Ok(())
}

#[derive(Clone)]
struct MultiRangeCountingSource {
    source: OwnedSource,
    ranges: Vec<(u64, u64)>,
    overlaps: Arc<std::sync::atomic::AtomicUsize>,
}

impl MultiRangeCountingSource {
    fn new(
        bytes: Vec<u8>,
        ranges: Vec<(u64, u64)>,
        overlaps: Arc<std::sync::atomic::AtomicUsize>,
    ) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            ranges,
            overlaps,
        }
    }
}

impl ReadAt for MultiRangeCountingSource {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let request_end = offset.saturating_add(output.len() as u64);
        if self
            .ranges
            .iter()
            .any(|(start, end)| offset < *end && request_end > *start)
        {
            self.overlaps
                .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        }
        self.source.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.source.version()
    }
}

#[derive(Clone)]
struct FailingRangeSource {
    source: OwnedSource,
    range_start: u64,
    range_end: u64,
    overlaps: Arc<std::sync::atomic::AtomicUsize>,
    armed: Arc<std::sync::atomic::AtomicBool>,
}

impl FailingRangeSource {
    fn new(
        bytes: Vec<u8>,
        range_start: u64,
        range_end: u64,
        overlaps: Arc<std::sync::atomic::AtomicUsize>,
        armed: Arc<std::sync::atomic::AtomicBool>,
    ) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            range_start,
            range_end,
            overlaps,
            armed,
        }
    }
}

impl ReadAt for FailingRangeSource {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let request_end = offset.saturating_add(output.len() as u64);
        if offset < self.range_end && request_end > self.range_start {
            self.overlaps
                .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
            if self.armed.load(std::sync::atomic::Ordering::SeqCst) {
                return Err(io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "synthetic source media read failure",
                ));
            }
        }
        self.source.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.source.version()
    }
}

#[test]
fn source_backed_cross_copy_rewrites_only_namespace_resolved_picture_embed() -> TestResult {
    let source = foreign_embed_picture_source_fixture()?;
    let destination = destination_fixture_with_media_collision()?;
    let output = publish(&source, &destination, 0, 1, 1)?;

    let (_, slide_xml, rels_xml) = copied_slide_parts(&output)?;
    let foreign = r#"x:embed="foreign-value""#;
    let foreign_offset = slide_xml.find(foreign).ok_or_else(|| {
        Error::Invalid("copied picture lost its foreign x:embed attribute".into())
    })?;
    let real_offset = slide_xml
        .find(r#"r:embed=""#)
        .ok_or_else(|| Error::Invalid("copied picture lost its real r:embed attribute".into()))?;
    assert!(foreign_offset < real_offset);
    assert!(slide_xml.contains(r#"cstate="print""#));
    assert_eq!(xml_attribute_values(&slide_xml, "cstate"), ["print"]);
    assert_eq!(
        xml_attribute_values(&slide_xml, "x:embed"),
        ["foreign-value"]
    );

    let embeds = xml_attribute_values(&slide_xml, "r:embed");
    assert_eq!(embeds.len(), 1);
    assert_ne!(embeds[0], "rIdSourceOnlyImage");
    let relationship_target = relationship_target_for_id(&rels_xml, &embeds[0])?;
    let image_media = copied_image_media(&output)?;
    assert_eq!(image_media.len(), 1);
    assert_eq!(image_media[0].2, SOURCE_PICTURE_MEDIA);
    assert_eq!(image_media[0].3, ct::PNG);
    let emitted_name = image_media[0]
        .1
        .as_str()
        .rsplit('/')
        .next()
        .ok_or_else(|| Error::Invalid("emitted image URI has no leaf name".into()))?;
    assert!(relationship_target.ends_with(emitted_name));

    let reopened = open_source(&output)?;
    let copied_image = reopened
        .slide(1)
        .ok_or_else(|| Error::Invalid("reopened copied picture slide is missing".into()))?
        .read_image(0)?;
    assert_eq!(copied_image.bytes(), SOURCE_PICTURE_MEDIA);
    assert_eq!(
        copied_image
            .descriptor()
            .target()
            .part_uri()
            .map(|uri| uri.as_str()),
        Some(image_media[0].1.as_str())
    );
    assert_eq!(
        copied_image.descriptor().target().content_type(),
        Some(ct::PNG)
    );
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_lexical_embed_without_relationship_namespace() -> TestResult {
    let valid = picture_source_fixture()?;
    let control = open_source(&valid)?;
    let control_slide = control
        .slide(0)
        .ok_or_else(|| Error::Invalid("valid picture source lacks slide 0".into()))?;
    assert_eq!(control_slide.images()?.len(), 1);

    let malformed = remove_root_relationship_namespace(&valid)?;
    let source = open_source(&malformed)?;
    let slide = source
        .slide(0)
        .ok_or_else(|| Error::Invalid("malformed picture source lacks slide 0".into()))?;
    let error = match slide.images() {
        Ok(_) => panic!("an unbound lexical picture embed must refuse image inventory"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Invalid(detail)
            if detail == "picture a:blip r prefix is not bound to the slide relationship namespace"
    ));
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_wrong_type_on_non_selected_source_slide_binding() -> TestResult
{
    let source = source_fixture()?;
    let destination = destination_fixture()?;
    open_source(&source)?;
    open_editor(&destination)?;
    let relationship_id = presentation_slide_relationship_id(&source, 1)?;
    let malformed = replace_relationship_type(
        &source,
        PRESENTATION_RELATIONSHIPS,
        &relationship_id,
        rt::SLIDE_LAYOUT,
    )?;
    assert_wrong_slide_binding_open_refused(&malformed, &destination, &relationship_id, true)
}

#[test]
fn source_backed_cross_copy_refuses_wrong_type_on_non_anchor_destination_slide_binding()
-> TestResult {
    let source = source_fixture()?;
    let destination = destination_fixture()?;
    open_source(&source)?;
    open_editor(&destination)?;
    let relationship_id = presentation_slide_relationship_id(&destination, 0)?;
    let malformed = replace_relationship_type(
        &destination,
        PRESENTATION_RELATIONSHIPS,
        &relationship_id,
        rt::SLIDE_LAYOUT,
    )?;
    assert_wrong_slide_binding_open_refused(&source, &malformed, &relationship_id, false)
}

#[test]
fn source_backed_cross_copy_observes_post_entry_planning_cancellation() -> TestResult {
    let source_bytes = picture_source_fixture()?;
    let destination_bytes = destination_fixture()?;
    let (range_start, range_end) = zip_member_data_range(&source_bytes, "ppt/slides/slide1.xml")?;
    let armed = Arc::new(std::sync::atomic::AtomicBool::new(false));
    let triggered = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let source_adapter = PlanningCancellationSource::new(
        source_bytes,
        range_start,
        range_end,
        Arc::clone(&armed),
        Arc::clone(&triggered),
        cancellation_source.clone(),
    );
    let budget = Budget::root(
        "pptx-cross-copy-planning-cancel",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let context = ExecutionContext::new(budget, cancellation, single_execution_limits()?);
    let source = SourceBackedPresentation::from_read_at_with_limits_and_execution_context(
        Arc::new(source_adapter),
        ReadLimits::default(),
        context.clone(),
    )?;
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits_and_execution_context(
        Arc::new(OwnedSource::new(destination_bytes)),
        ReadLimits::default(),
        context,
    )?;
    assert_eq!(triggered.load(std::sync::atomic::Ordering::SeqCst), 0);
    armed.store(true, std::sync::atomic::Ordering::SeqCst);

    let error = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
        Ok(_) => panic!("planning cancellation must not return a prepared plan"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Opc(OpcError::Cancelled)));
    assert!(cancellation_source.is_cancelled());
    assert!(triggered.load(std::sync::atomic::Ordering::SeqCst) > 0);
    Ok(())
}

fn foreign_embed_picture_source_fixture() -> TestResult<Vec<u8>> {
    replace_text_member(
        &picture_source_fixture()?,
        "ppt/slides/slide1.xml",
        r#"<a:blip r:embed="rIdSourceOnlyImage"/>"#,
        r#"<a:blip xmlns:x="urn:test:foreign" cstate="print" x:embed="foreign-value" r:embed="rIdSourceOnlyImage"/>"#,
    )
}

fn remove_root_relationship_namespace(bytes: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(bytes, "ppt/slides/slide1.xml", |payload| {
        let xml = std::str::from_utf8(payload)
            .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
        let root_start = xml
            .find("<p:sld ")
            .ok_or_else(|| Error::Invalid("picture source slide root is missing".into()))?;
        let root_end = root_start
            + xml[root_start..].find('>').ok_or_else(|| {
                Error::Invalid("picture source slide root is unterminated".into())
            })?;
        let root = &xml[root_start..=root_end];
        let namespace_markers = [
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
            r#"xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships""#,
        ];
        let (marker_start, marker_len) = namespace_markers
            .iter()
            .find_map(|marker| root.find(*marker).map(|start| (start, marker.len())))
            .ok_or_else(|| Error::Invalid("slide root relationship namespace is missing".into()))?;
        let attribute_start = root[..marker_start]
            .char_indices()
            .rev()
            .find(|(_, character)| character.is_ascii_whitespace())
            .map(|(index, _)| index)
            .ok_or_else(|| {
                Error::Invalid("slide root namespace has no attribute boundary".into())
            })?;
        let mut changed = xml.to_owned();
        changed.replace_range(
            root_start + attribute_start..root_start + marker_start + marker_len,
            "",
        );
        Ok(changed.into_bytes())
    })
}

const TEST_PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const TEST_STRICT_PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/presentationml/main";
const TEST_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const TEST_STRICT_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

fn presentation_slide_relationship_id(bytes: &[u8], position: usize) -> TestResult<String> {
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| Error::Invalid(format!("cannot index presentation fixture: {error}")))?;
    let xml =
        String::from_utf8(archive.read(PRESENTATION_XML).map_err(|error| {
            Error::Invalid(format!("cannot read presentation fixture: {error}"))
        })?)
        .map_err(|error| Error::Invalid(format!("presentation fixture is not UTF-8: {error}")))?;

    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::new();
    let mut root_prefix = None;
    let mut saw_root = false;
    let mut saw_list = false;
    let mut ids = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                let kind = presentation_element_kind(&namespace, element.name())?;
                let prefix = element.name().prefix().map(|value| value.as_ref().to_vec());
                if stack.is_empty() {
                    if saw_root || kind != PresentationElement::Presentation {
                        return Err(Error::Invalid(
                            "presentation XML has an invalid root element".into(),
                        ));
                    }
                    root_prefix = prefix;
                    if root_prefix.is_none() {
                        return Err(Error::Invalid(
                            "presentation root has no qualified namespace prefix".into(),
                        ));
                    }
                    saw_root = true;
                } else {
                    validate_presentation_position(
                        kind,
                        prefix.as_deref(),
                        &stack,
                        root_prefix.as_deref(),
                    )?;
                }
                if kind == PresentationElement::SlideIdList {
                    if stack.as_slice() != [PresentationElement::Presentation] || saw_list {
                        return Err(Error::Invalid(
                            "presentation has a non-direct or duplicate slide-ID list".into(),
                        ));
                    }
                    saw_list = true;
                }
                if let Some(relationship_id) =
                    direct_slide_relationship_id(kind, &stack, &element, &reader)?
                {
                    ids.push(relationship_id);
                }
                stack.push(kind);
            },
            Event::Empty(element) => {
                let kind = presentation_element_kind(&namespace, element.name())?;
                let prefix = element.name().prefix().map(|value| value.as_ref().to_vec());
                if stack.is_empty() {
                    if saw_root || kind != PresentationElement::Presentation {
                        return Err(Error::Invalid(
                            "presentation XML has an invalid root element".into(),
                        ));
                    }
                    root_prefix = prefix;
                    saw_root = true;
                } else {
                    validate_presentation_position(
                        kind,
                        prefix.as_deref(),
                        &stack,
                        root_prefix.as_deref(),
                    )?;
                }
                if kind == PresentationElement::SlideIdList {
                    if stack.as_slice() != [PresentationElement::Presentation] || saw_list {
                        return Err(Error::Invalid(
                            "presentation has a non-direct or duplicate slide-ID list".into(),
                        ));
                    }
                    saw_list = true;
                } else if let Some(relationship_id) =
                    direct_slide_relationship_id(kind, &stack, &element, &reader)?
                {
                    ids.push(relationship_id);
                }
            },
            Event::End(element) => {
                let kind = presentation_element_kind(&namespace, element.name())?;
                if stack.pop() != Some(kind) {
                    return Err(Error::Invalid(
                        "presentation XML has mismatched element nesting".into(),
                    ));
                }
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::Invalid(
                    "presentation XML has unexpected character data".into(),
                ));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "presentation XML contains a rejected declaration".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !saw_root || !stack.is_empty() || !saw_list {
        return Err(Error::Invalid(
            "presentation XML has no complete direct slide-ID list".into(),
        ));
    }
    ids.into_iter()
        .nth(position)
        .ok_or_else(|| Error::Invalid("presentation slide binding is missing".into()))
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PresentationElement {
    Presentation,
    SlideIdList,
    SlideId,
    Other,
}

fn presentation_element_kind(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
) -> TestResult<PresentationElement> {
    let local = name.local_name();
    let expected = match local.as_ref() {
        b"presentation" => PresentationElement::Presentation,
        b"sldIdLst" => PresentationElement::SlideIdList,
        b"sldId" => PresentationElement::SlideId,
        _ => PresentationElement::Other,
    };
    let is_presentation_namespace = match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == TEST_PRESENTATIONML_NAMESPACE
                || *value == TEST_STRICT_PRESENTATIONML_NAMESPACE
        },
        ResolveResult::Unbound => false,
        ResolveResult::Unknown(prefix) => {
            return Err(Error::Invalid(format!(
                "unresolved presentation namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        },
    };
    if !is_presentation_namespace && expected != PresentationElement::Other {
        return Err(Error::Invalid(
            "presentation XML contains a lookalike element in another namespace".into(),
        ));
    }
    Ok(if is_presentation_namespace {
        expected
    } else {
        PresentationElement::Other
    })
}

fn validate_presentation_position(
    kind: PresentationElement,
    prefix: Option<&[u8]>,
    stack: &[PresentationElement],
    root_prefix: Option<&[u8]>,
) -> TestResult<()> {
    if kind != PresentationElement::Other && prefix != root_prefix {
        return Err(Error::Invalid(
            "presentation XML has an ambiguous namespace alias or rebinding".into(),
        ));
    }
    if kind == PresentationElement::Presentation {
        return Err(Error::Invalid(
            "presentation XML contains a nested presentation root".into(),
        ));
    }
    if kind == PresentationElement::SlideIdList && stack != [PresentationElement::Presentation] {
        return Err(Error::Invalid(
            "presentation slide-ID list is not a direct root child".into(),
        ));
    }
    Ok(())
}

fn direct_slide_relationship_id(
    kind: PresentationElement,
    stack: &[PresentationElement],
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> TestResult<Option<String>> {
    if kind != PresentationElement::SlideId {
        return Ok(None);
    }
    if stack
        != [
            PresentationElement::Presentation,
            PresentationElement::SlideIdList,
        ]
    {
        return Err(Error::Invalid(
            "presentation slide-ID binding is not a direct list child".into(),
        ));
    }
    Ok(Some(presentation_relationship_id(element, reader)?))
}

fn presentation_relationship_id(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> TestResult<String> {
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let is_relationship_id = match namespace {
            ResolveResult::Bound(Namespace(value)) => {
                (value == TEST_RELATIONSHIP_NAMESPACE
                    || value == TEST_STRICT_RELATIONSHIP_NAMESPACE)
                    && local.as_ref() == b"id"
            },
            ResolveResult::Unbound => false,
            ResolveResult::Unknown(prefix) => {
                return Err(Error::Invalid(format!(
                    "unresolved presentation attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        if !is_relationship_id {
            continue;
        }
        if relationship_id.is_some() {
            return Err(Error::Invalid(
                "presentation slide-ID binding has duplicate relationship IDs".into(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        if value.is_empty() {
            return Err(Error::Invalid(
                "presentation slide-ID relationship ID is empty".into(),
            ));
        }
        relationship_id = Some(value);
    }
    relationship_id.ok_or_else(|| {
        Error::Invalid("presentation slide-ID binding lacks an r:id attribute".into())
    })
}

fn replace_relationship_type(
    bytes: &[u8],
    member: &str,
    relationship_id: &str,
    replacement_type: &str,
) -> TestResult<Vec<u8>> {
    rewrite_member(bytes, member, |payload| {
        let xml = std::str::from_utf8(payload)
            .map_err(|error| Error::Invalid(format!("relationships are not UTF-8: {error}")))?;
        let id_marker = format!(r#"Id="{relationship_id}""#);
        let id_offset = xml
            .find(&id_marker)
            .ok_or_else(|| Error::Invalid(format!("relationship {relationship_id} is missing")))?;
        let record_start = xml[..id_offset]
            .rfind("<Relationship ")
            .ok_or_else(|| Error::Invalid("relationship record start is missing".into()))?;
        let record_end = id_offset
            + xml[id_offset..]
                .find("/>")
                .ok_or_else(|| Error::Invalid("relationship record end is missing".into()))?;
        let record = &xml[record_start..record_end];
        let type_marker = "Type=\"";
        let type_start = record
            .find(type_marker)
            .ok_or_else(|| Error::Invalid("relationship type is missing".into()))?
            + type_marker.len();
        let type_end = type_start
            + record[type_start..]
                .find('"')
                .ok_or_else(|| Error::Invalid("relationship type is unterminated".into()))?;
        let mut changed = xml.to_owned();
        changed.replace_range(
            record_start + type_start..record_start + type_end,
            replacement_type,
        );
        Ok(changed.into_bytes())
    })
}

fn assert_wrong_slide_binding_open_refused(
    source_bytes: &[u8],
    destination_bytes: &[u8],
    relationship_id: &str,
    source_side: bool,
) -> TestResult {
    let error = if source_side {
        match open_source(source_bytes) {
            Ok(_) => panic!("wrong source slide binding must be refused while opening"),
            Err(error) => error,
        }
    } else {
        open_source(source_bytes)?;
        match open_editor(destination_bytes) {
            Ok(_) => panic!("wrong destination slide binding must be refused while opening"),
            Err(error) => error,
        }
    };
    assert!(matches!(
        &error,
        Error::Relationship(detail)
            if detail.contains("unexpected type") && detail.contains(relationship_id)
    ));
    Ok(())
}

struct PlanningCancellationSource {
    source: OwnedSource,
    range_start: u64,
    range_end: u64,
    armed: Arc<std::sync::atomic::AtomicBool>,
    triggered: Arc<std::sync::atomic::AtomicUsize>,
    cancellation: CancellationSource,
}

impl PlanningCancellationSource {
    fn new(
        bytes: Vec<u8>,
        range_start: u64,
        range_end: u64,
        armed: Arc<std::sync::atomic::AtomicBool>,
        triggered: Arc<std::sync::atomic::AtomicUsize>,
        cancellation: CancellationSource,
    ) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            range_start,
            range_end,
            armed,
            triggered,
            cancellation,
        }
    }
}

impl ReadAt for PlanningCancellationSource {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let request_end = offset.saturating_add(output.len() as u64);
        if self.armed.load(std::sync::atomic::Ordering::SeqCst)
            && offset < self.range_end
            && request_end > self.range_start
        {
            self.triggered
                .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
            self.cancellation.cancel();
        }
        self.source.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.source.version()
    }
}

#[test]
fn source_backed_cross_copy_accepts_picture_on_destination_anchor_and_preserves_anchor_closure()
-> TestResult {
    let source = picture_source_fixture()?;
    let destination = destination_fixture_with_anchor_picture()?;
    let anchor_slide = read_zip_member_bytes(&destination, "ppt/slides/slide2.xml")?;
    let anchor_relationships =
        read_zip_member_bytes(&destination, "ppt/slides/_rels/slide2.xml.rels")?;
    let anchor_media = read_zip_member_bytes(&destination, "ppt/media/anchor.png")?;
    let unrelated_media = read_zip_member_bytes(&destination, "ppt/media/unrelated.png")?;
    let anchor_relationships = String::from_utf8(anchor_relationships)
        .map_err(|error| Error::Invalid(format!("anchor relationships are not UTF-8: {error}")))?;
    let anchor_layout_target =
        relationship_target_for_type(&anchor_relationships, rt::SLIDE_LAYOUT)?;

    let output = publish(&source, &destination, 0, 1, 1)?;
    assert_eq!(
        read_zip_member_bytes(&output, "ppt/slides/slide2.xml")?,
        anchor_slide
    );
    assert_eq!(
        read_zip_member_bytes(&output, "ppt/slides/_rels/slide2.xml.rels")?,
        anchor_relationships.as_bytes()
    );
    assert_eq!(
        read_zip_member_bytes(&output, "ppt/media/anchor.png")?,
        anchor_media
    );
    assert_eq!(
        read_zip_member_bytes(&output, "ppt/media/unrelated.png")?,
        unrelated_media
    );

    let (_, _, copied_relationships) = copied_slide_parts(&output)?;
    assert!(copied_relationships.contains(anchor_layout_target.as_str()));
    let image_media = copied_image_media(&output)?;
    assert_eq!(image_media.len(), 1);
    assert_eq!(image_media[0].2, SOURCE_PICTURE_MEDIA);
    assert_eq!(image_media[0].3, ct::PNG);
    Package::from_vec(output.clone())?;
    open_source(&output)?;
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_mismatched_presentation_end_name() -> TestResult {
    let valid = source_fixture()?;
    open_source(&valid)?;
    let malformed = replace_text_member(
        &valid,
        PRESENTATION_XML,
        "</p:sldIdLst>",
        "</p:sldMasterIdLst>",
    )?;
    let error = match open_source(&malformed) {
        Ok(_) => panic!("mismatched presentation end name must be refused while opening"),
        Err(error) => error,
    };
    assert_mismatched_xml_refusal(&error);
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_mismatched_slide_end_name_before_planning() -> TestResult {
    let valid = picture_source_fixture()?;
    let malformed =
        replace_text_member(&valid, "ppt/slides/slide1.xml", "</p:spTree>", "</p:cSld>")?;
    let source = open_source(&malformed)?;
    let editor = open_editor(&destination_fixture()?)?;
    let error = match editor.plan_cross_slide_copy(&source, 0, 1, 1) {
        Ok(_) => panic!("mismatched slide end name must not produce a plan"),
        Err(error) => error,
    };
    assert_mismatched_xml_refusal(&error);
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_mismatched_relationship_end_name() -> TestResult {
    let valid = picture_source_fixture()?;
    open_source(&valid)?;
    let malformed = replace_text_member(
        &valid,
        "ppt/slides/_rels/slide1.xml.rels",
        "</Relationships>",
        "</WrongRelationships>",
    )?;
    let error = match open_source(&malformed) {
        Ok(_) => panic!("mismatched relationship end name must be refused while opening"),
        Err(error) => error,
    };
    assert_mismatched_xml_refusal(&error);
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_unbound_foreign_slide_attribute() -> TestResult {
    let valid = picture_source_fixture()?;
    let control = open_source(&valid)?;
    let control_slide = control
        .slide(0)
        .ok_or_else(|| Error::Invalid("valid picture source lacks slide 0".into()))?;
    assert_eq!(control_slide.images()?.len(), 1);

    let malformed = unbound_foreign_attribute_picture_source_fixture()?;
    let source = open_source(&malformed)?;
    let slide = source
        .slide(0)
        .ok_or_else(|| Error::Invalid("unbound-attribute source slide is missing".into()))?;
    let error = match slide.images() {
        Ok(_) => panic!("unbound foreign slide attribute must refuse image inventory"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Invalid(detail) if detail == "picture a:blip contains an unresolved attribute prefix"
    ));
    Ok(())
}

#[test]
fn source_backed_cross_copy_refuses_picture_inventory_when_r_is_rebound() -> TestResult {
    let valid = picture_source_fixture()?;
    let control = open_source(&valid)?;
    let control_slide = control
        .slide(0)
        .ok_or_else(|| Error::Invalid("valid picture source lacks slide 0".into()))?;
    assert_eq!(control_slide.images()?.len(), 1);

    let rebound = replace_text_member(
        &valid,
        "ppt/slides/slide1.xml",
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
        r#"xmlns:r="urn:test:foreign""#,
    )?;
    let source = open_source(&rebound)?;
    let slide = source
        .slide(0)
        .ok_or_else(|| Error::Invalid("rebound-r source slide is missing".into()))?;
    let error = match slide.images() {
        Ok(_) => panic!("rebound relationship namespace must refuse image inventory"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Invalid(detail)
            if detail == "picture a:blip r prefix is not bound to the slide relationship namespace"
    ));
    Ok(())
}

#[test]
fn source_backed_cross_copy_shared_image_has_exact_physical_media_closure() -> TestResult {
    let source = shared_picture_source_fixture()?;
    let destination = destination_fixture_with_media_collision()?;
    let output = publish(&source, &destination, 0, 1, 1)?;
    let actual = physical_media_members(&output)?;
    let expected: BTreeSet<String> = [
        "ppt/media/source-only.png",
        "ppt/media/source-only-copy1.png",
        "ppt/media/unrelated.png",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect();
    assert_eq!(actual, expected);
    assert_eq!(
        actual
            .iter()
            .filter(|member| member.contains("-copy"))
            .count(),
        1
    );
    open_source(&output)?;
    Ok(())
}

fn destination_fixture_with_anchor_picture() -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(destination_fixture()?)?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/media/anchor.png").map_err(Error::Uri)?,
        ct::PNG.to_owned(),
        b"destination-anchor-media".to_vec(),
    )))?;
    package
        .get_part_mut(&PackURI::new("/ppt/slides/slide2.xml").map_err(Error::Uri)?)?
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/anchor.png".to_owned(),
            "rIdAnchorImage".to_owned(),
            litchi_opc::TargetMode::Internal,
        )?;
    let bytes = PackageWriter::to_bytes(&package)?;
    append_slide_element_to_member(
        &bytes,
        "ppt/slides/slide2.xml",
        r#"<p:pic><p:nvPicPr><p:cNvPr id="43" name="Anchor Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdAnchorImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
    )
}

fn append_slide_element_to_member(
    bytes: &[u8],
    member: &str,
    element: &str,
) -> TestResult<Vec<u8>> {
    rewrite_member(bytes, member, |payload| {
        let xml = std::str::from_utf8(payload)
            .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
        let marker = "</p:spTree>";
        let insertion = xml
            .rfind(marker)
            .ok_or_else(|| Error::Invalid("slide shape tree is missing".into()))?;
        Ok(format!("{}{element}{}", &xml[..insertion], &xml[insertion..]).into_bytes())
    })
}

fn read_zip_member_bytes(bytes: &[u8], member: &str) -> TestResult<Vec<u8>> {
    ArchiveReader::new(bytes)
        .map_err(|error| Error::Invalid(format!("cannot index ZIP member {member}: {error}")))?
        .read(member)
        .map_err(|error| Error::Invalid(format!("cannot read ZIP member {member}: {error}")))
}

fn relationship_target_for_type(rels_xml: &str, relationship_type: &str) -> TestResult<String> {
    let type_marker = format!(r#"Type="{relationship_type}""#);
    let type_offset = rels_xml
        .find(&type_marker)
        .ok_or_else(|| Error::Invalid("requested relationship type is missing".into()))?;
    let record_start = rels_xml[..type_offset]
        .rfind("<Relationship ")
        .ok_or_else(|| Error::Invalid("relationship record start is missing".into()))?;
    let record_end = type_offset
        + rels_xml[type_offset..]
            .find("/>")
            .ok_or_else(|| Error::Invalid("relationship record end is missing".into()))?;
    let record = &rels_xml[record_start..record_end];
    let target_marker = "Target=\"";
    let target_start = record
        .find(target_marker)
        .ok_or_else(|| Error::Invalid("relationship target is missing".into()))?
        + target_marker.len();
    let target_end = target_start
        + record[target_start..]
            .find('"')
            .ok_or_else(|| Error::Invalid("relationship target is unterminated".into()))?;
    Ok(record[target_start..target_end].to_owned())
}

fn unbound_foreign_attribute_picture_source_fixture() -> TestResult<Vec<u8>> {
    replace_text_member(
        &picture_source_fixture()?,
        "ppt/slides/slide1.xml",
        r#"<a:blip r:embed="rIdSourceOnlyImage"/>"#,
        r#"<a:blip x:embed="unbound-foreign-value" r:embed="rIdSourceOnlyImage"/>"#,
    )
}

fn assert_mismatched_xml_refusal(error: &Error) {
    assert!(matches!(
        error,
        Error::Xml(_) | Error::Invalid(_) | Error::Opc(_) | Error::SlideCopyPlan { .. }
    ));
    let message = error.to_string().to_ascii_lowercase();
    assert!(
        message.contains("mismatch")
            || message.contains("expect")
            || message.contains("closing")
            || message.contains("element")
            || message.contains("xml"),
        "unexpected malformed XML refusal: {error:?}"
    );
}

fn physical_media_members(bytes: &[u8]) -> TestResult<BTreeSet<String>> {
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| Error::Invalid(format!("cannot index media members: {error}")))?;
    Ok(archive
        .file_names()
        .filter(|member| member.starts_with("ppt/media/"))
        .map(str::to_owned)
        .collect())
}
