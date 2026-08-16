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
