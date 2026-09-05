#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test fixture construction and assertions panic on failure by design"
)]

use std::sync::Arc;

use litchi_opc::{
    FontEmbedding, OpcPackage, PackURI, PackageWriter, Part, Relationships, SaveOptions,
};
use litchi_pptx::opened::Limits;
use litchi_pptx::{CrossSlideCopyPatch, Error, Package, SlideCopyRefusal};

type TestResult<T = ()> = litchi_pptx::Result<T>;

const SENTINEL_PART: &str = "/custom/cross-copy-sentinel.bin";
const SENTINEL_RELATIONSHIP_ID: &str = "cross-copy-sentinel";
const SENTINEL_RELATIONSHIP_COUNT: usize = 0x51_7e;
const SENTINEL_BYTES: &[u8] = b"custom-part-implementation";

#[test]
fn cross_copy_rejects_stale_semantic_and_physical_inputs_without_mutation() -> TestResult {
    let source_bytes = authored_package(&["Source A"])?;
    let destination_bytes = authored_package(&["Destination A", "Destination B"])?;

    let mut source = Package::from_vec(source_bytes.clone())?;
    let planned_destination = Package::from_vec(destination_bytes.clone())?;
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = planned_destination.opened_presentation()?;
    let plan = destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0_usize, 0_usize, 1)?;

    let stale_metadata_bytes = authored_package(&["Destination A", "Destination Changed"])?;
    let mut stale_metadata = Package::from_vec(stale_metadata_bytes)?;
    let stale_metadata_before = serialized(&mut stale_metadata)?;
    let error = stale_metadata
        .apply_cross_slide_copy_plan(&source, &plan)
        .expect_err("a changed destination graph must reject the old plan");
    assert_unsafe_edit(error);
    assert_eq!(serialized(&mut stale_metadata)?, stale_metadata_before);

    let stale_source_bytes = with_eocd_comment(source_bytes.clone(), b"changed source archive")?;
    let stale_source = Package::from_vec(stale_source_bytes)?;
    let stale_source_snapshot = stale_source.opened_presentation()?;
    assert_eq!(
        stale_source_snapshot.slides()[0].name(),
        source_snapshot.slides()[0].name()
    );

    let mut stale_source_destination = Package::from_vec(destination_bytes.clone())?;
    let stale_source_destination_before = serialized(&mut stale_source_destination)?;
    let error = stale_source_destination
        .apply_cross_slide_copy_plan(&stale_source, &plan)
        .expect_err("a changed source archive must reject the old plan");
    assert_unsafe_edit(error);
    assert_eq!(
        serialized(&mut stale_source_destination)?,
        stale_source_destination_before
    );
    assert_eq!(serialized(&mut source)?, source_bytes);

    let stale_destination_bytes =
        with_eocd_comment(destination_bytes.clone(), b"changed destination archive")?;
    let mut stale_destination = Package::from_vec(stale_destination_bytes)?;
    let stale_destination_before = serialized(&mut stale_destination)?;
    let error = stale_destination
        .apply_cross_slide_copy_plan(&source, &plan)
        .expect_err("a changed destination archive must reject the old plan");
    assert_unsafe_edit(error);
    assert_eq!(
        serialized(&mut stale_destination)?,
        stale_destination_before
    );
    Ok(())
}

#[test]
fn cross_copy_rejects_tampered_durable_patch_and_limits_before_mutation() -> TestResult {
    let source_bytes = authored_package(&["Source A"])?;
    let destination_bytes = authored_package(&["Destination A", "Destination B"])?;
    let source = Package::from_vec(source_bytes)?;
    let destination_for_plan = Package::from_vec(destination_bytes.clone())?;
    let plan = destination_for_plan
        .opened_presentation()?
        .plan_cross_slide_copy(&source.opened_presentation()?, 0_usize, 0_usize, 1)?;

    let mut tampered_bytes = plan.patch().to_bytes()?;
    // LPCP0002 stores the source physical revision after the three semantic
    // revisions. Changing it keeps the durable patch well-formed while
    // invalidating its source authorization.
    let source_physical_revision_offset = 8 + 32 * 3;
    tampered_bytes[source_physical_revision_offset] ^= 1;
    let tampered = CrossSlideCopyPatch::from_bytes(&tampered_bytes)?;

    let mut destination = Package::from_vec(destination_bytes.clone())?;
    let before = serialized(&mut destination)?;
    let error = destination
        .apply_cross_slide_copy_patch(&source, &tampered)
        .expect_err("a tampered durable source revision must be rejected");
    assert_unsafe_edit(error);
    assert_eq!(serialized(&mut destination)?, before);

    let tiny_limits = Limits::new(4_096, 1, 1_024, 1, 1_024)
        .ok_or_else(|| Error::Invalid("test limits are invalid".into()))?;
    let error = CrossSlideCopyPatch::from_bytes_with_limits(&plan.patch().to_bytes()?, tiny_limits)
        .expect_err("a restrictive durable-patch byte limit must be enforced");
    assert!(matches!(error, Error::Limit { .. }));
    Ok(())
}

#[test]
fn cross_copy_plan_and_durable_patch_are_byte_equivalent_and_inverse_exact() -> TestResult {
    let source_bytes = authored_package(&["Source A"])?;
    let destination_bytes = with_eocd_comment(
        authored_package(&["Destination A", "Destination B"])?,
        b"destination archive comment",
    )?;

    let mut source = Package::from_vec(source_bytes.clone())?;
    let destination_for_plan = Package::from_vec(destination_bytes.clone())?;
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = destination_for_plan.opened_presentation()?;
    let plan = destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0_usize, 0_usize, 1)?;
    let durable = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;

    let mut plan_destination = Package::from_vec(destination_bytes.clone())?;
    plan_destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let plan_output = serialized(&mut plan_destination)?;

    let mut patch_destination = Package::from_vec(destination_bytes.clone())?;
    patch_destination.apply_cross_slide_copy_patch(&source, &durable)?;
    let patch_output = serialized(&mut patch_destination)?;
    assert_eq!(plan_output, patch_output);
    assert_eq!(
        archive_comment(&plan_output)?,
        b"destination archive comment"
    );

    let inverse = CrossSlideCopyPatch::from_bytes(&durable.inverse().to_bytes()?)?;
    patch_destination.apply_cross_slide_copy_patch(&source, &inverse)?;
    assert_eq!(serialized(&mut patch_destination)?, destination_bytes);
    assert_eq!(serialized(&mut source)?, source_bytes);

    let second_source = Package::from_vec(source_bytes)?;
    let second_destination = Package::from_vec(destination_bytes)?;
    let second_plan = second_destination
        .opened_presentation()?
        .plan_cross_slide_copy(&second_source.opened_presentation()?, 0_usize, 0_usize, 1)?;
    assert_eq!(plan, second_plan);
    Ok(())
}

#[test]
fn cross_copy_replans_after_dirty_destination_for_sequential_copies() -> TestResult {
    let source_bytes = authored_package(&["Source One", "Source Two"])?;
    let destination_bytes = with_eocd_comment(
        authored_package(&["Destination First", "Destination Last"])?,
        b"dirty destination archive",
    )?;
    let mut source = Package::from_vec(source_bytes.clone())?;
    let source_snapshot = source.opened_presentation()?;
    let mut destination = Package::from_vec(destination_bytes.clone())?;

    let first_plan = destination.opened_presentation()?.plan_cross_slide_copy(
        &source_snapshot,
        0_usize,
        0_usize,
        1,
    )?;
    destination.apply_cross_slide_copy_plan(&source, &first_plan)?;
    let after_first = serialized(&mut destination)?;
    assert_eq!(archive_comment(&after_first)?, b"dirty destination archive");

    let before_stale_retry = after_first.clone();
    let error = destination
        .apply_cross_slide_copy_plan(&source, &first_plan)
        .expect_err("the first plan must be stale after its successful publication");
    assert_unsafe_edit(error);
    assert_eq!(serialized(&mut destination)?, before_stale_retry);

    let second_plan = destination.opened_presentation()?.plan_cross_slide_copy(
        &source_snapshot,
        1_usize,
        2_usize,
        3,
    )?;
    destination.apply_cross_slide_copy_plan(&source, &second_plan)?;
    let output = serialized(&mut destination)?;
    let reopened = Package::from_vec(output.clone())?;
    let reopened_snapshot = reopened.opened_presentation()?;
    let slides = reopened_snapshot.slides();
    assert_eq!(
        slides.iter().map(|slide| slide.name()).collect::<Vec<_>>(),
        vec![
            "Destination First",
            "Source One",
            "Destination Last",
            "Source Two"
        ]
    );
    assert_eq!(archive_comment(&output)?, b"dirty destination archive");

    let mut replay = Package::from_vec(destination_bytes)?;
    replay.apply_cross_slide_copy_plan(&source, &first_plan)?;
    replay.apply_cross_slide_copy_plan(&source, &second_plan)?;
    assert_eq!(serialized(&mut replay)?, output);
    assert_eq!(serialized(&mut source)?, source_bytes);
    Ok(())
}

#[test]
fn cross_copy_exact_archive_limit_accepts_and_one_short_rejects_before_publication() -> TestResult {
    let source_bytes = authored_package(&["Source A"])?;
    let destination_bytes = authored_package(&["Destination A", "Destination B"])?;

    let source = Package::from_vec(source_bytes.clone())?;
    let mut baseline_destination = Package::from_vec(destination_bytes.clone())?;
    let plan = baseline_destination
        .opened_presentation()?
        .plan_cross_slide_copy(&source.opened_presentation()?, 0_usize, 0_usize, 1)?;
    baseline_destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let expected = serialized(&mut baseline_destination)?;
    assert!(expected.len() > source_bytes.len());
    assert!(expected.len() > destination_bytes.len());

    let limits = |max_patch_bytes| {
        Limits::new(
            4_096,
            max_patch_bytes,
            8 * 1024 * 1024,
            64,
            256 * 1024 * 1024,
        )
        .ok_or_else(|| Error::Invalid("test limits are invalid".into()))
    };

    let exact_limits = limits(expected.len())?;
    let mut exact_source = Package::from_vec(source_bytes.clone())?;
    let mut exact_destination = Package::from_vec(destination_bytes.clone())?;
    let exact_source_before = source_bytes.clone();
    let exact_destination_before = destination_bytes.clone();
    let exact_plan = exact_destination
        .opened_presentation_with_limits(exact_limits)?
        .plan_cross_slide_copy(
            &exact_source.opened_presentation_with_limits(exact_limits)?,
            0_usize,
            0_usize,
            1,
        )?;
    exact_destination.apply_cross_slide_copy_plan(&exact_source, &exact_plan)?;
    assert_eq!(serialized(&mut exact_destination)?, expected);
    assert_eq!(serialized(&mut exact_source)?, exact_source_before);
    assert_ne!(
        serialized(&mut exact_destination)?,
        exact_destination_before
    );

    let one_short = expected
        .len()
        .checked_sub(1)
        .ok_or_else(|| Error::Invalid("cross-copy output is empty".into()))?;
    let short_limits = limits(one_short)?;
    let mut short_source = Package::from_vec(source_bytes.clone())?;
    let mut short_destination = Package::from_vec(destination_bytes.clone())?;
    let short_source_before = source_bytes.clone();
    let short_destination_before = destination_bytes.clone();
    let error = short_destination
        .opened_presentation_with_limits(short_limits)?
        .plan_cross_slide_copy(
            &short_source.opened_presentation_with_limits(short_limits)?,
            0_usize,
            0_usize,
            1,
        )
        .expect_err("one byte below the serialized candidate must fail");
    assert!(matches!(
        error,
        Error::Limit {
            resource: "cross-slide serialized archive bytes",
            limit,
        } if limit == one_short
    ));
    assert_eq!(
        serialized(&mut short_destination)?,
        short_destination_before
    );
    assert_eq!(serialized(&mut short_source)?, short_source_before);
    Ok(())
}

#[test]
fn borrowed_real_producer_fixture_refuses_cross_copy_authorization() -> TestResult {
    let source = Package::from_bytes(include_bytes!("../../../test-data/ooxml/pptx/shapes.pptx"))?;
    let mut destination = Package::from_vec(authored_package(&[
        "Destination First",
        "Destination Last",
    ])?)?;
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = destination.opened_presentation()?;
    let before = serialized(&mut destination)?;
    let error = destination_snapshot
        .plan_cross_slide_copy(&source_snapshot, 0_usize, 0_usize, 1)
        .expect_err("borrowed real-producer input cannot authorize physical copying");
    assert!(matches!(
        error,
        Error::SlideCopyPlan {
            kind: SlideCopyRefusal::UnknownPhysicalMember,
            ..
        }
    ));
    assert_eq!(serialized(&mut destination)?, before);
    Ok(())
}

#[test]
fn cross_copy_preserves_custom_parts_and_save_options_through_forward_and_inverse() -> TestResult {
    let source_bytes = authored_package(&["Source A"])?;
    let source = Package::from_vec(source_bytes)?;
    let sentinel_name = PackURI::new(SENTINEL_PART).map_err(Error::Uri)?;
    let mut plan_destination = custom_destination(&sentinel_name)?;
    let before = serialized(&mut plan_destination)?;
    assert_custom_destination_state(&plan_destination, &sentinel_name)?;

    let plan = plan_destination
        .opened_presentation()?
        .plan_cross_slide_copy(&source.opened_presentation()?, 0_usize, 0_usize, 1)?;
    let durable = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;
    let inverse = CrossSlideCopyPatch::from_bytes(&durable.inverse().to_bytes()?)?;

    plan_destination.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_custom_destination_state(&plan_destination, &sentinel_name)?;
    let plan_output = serialized(&mut plan_destination)?;

    plan_destination.apply_cross_slide_copy_patch(&source, &inverse)?;
    assert_custom_destination_state(&plan_destination, &sentinel_name)?;
    assert_eq!(serialized(&mut plan_destination)?, before);

    let mut patch_destination = custom_destination(&sentinel_name)?;
    assert_eq!(serialized(&mut patch_destination)?, before);
    patch_destination.apply_cross_slide_copy_patch(&source, &durable)?;
    assert_custom_destination_state(&patch_destination, &sentinel_name)?;
    assert_eq!(serialized(&mut patch_destination)?, plan_output);
    patch_destination.apply_cross_slide_copy_patch(&source, &inverse)?;
    assert_custom_destination_state(&patch_destination, &sentinel_name)?;
    assert_eq!(serialized(&mut patch_destination)?, before);
    Ok(())
}

fn authored_package(names: &[&str]) -> TestResult<Vec<u8>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for name in names {
            let slide = presentation.add_slide()?;
            slide.set_title(name);
            slide.add_text_box(&format!("body:{name}"), 10, 20, 300, 400);
        }
    }
    let bytes = package.to_bytes()?;
    let mut opc = OpcPackage::from_vec(bytes)?;
    for (index, name) in names.iter().enumerate() {
        let slide_name =
            PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1)).map_err(Error::Uri)?;
        let part = opc.get_part_mut(&slide_name)?;
        let xml = std::str::from_utf8(part.blob())
            .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
        let marker = "<p:cSld name=\"";
        let marker_start = xml
            .find(marker)
            .ok_or_else(|| Error::Invalid("authored slide has no cSld name".into()))?;
        let value_start = marker_start + marker.len();
        let value_end = value_start
            + xml[value_start..]
                .find('"')
                .ok_or_else(|| Error::Invalid("authored slide name is unterminated".into()))?;
        let updated = format!("{}{}{}", &xml[..value_start], name, &xml[value_end..]);
        part.set_blob(updated.into_bytes());
    }
    Ok(PackageWriter::to_bytes(&opc)?)
}

fn custom_destination(sentinel_name: &PackURI) -> TestResult<Package> {
    let mut package = Package::from_vec(authored_package(&["Destination A", "Destination B"])?)?;
    let sentinel_name = sentinel_name.clone();
    package.edit_opc(|opc| {
        opc.try_add_part(Box::new(SentinelPart::new(sentinel_name)))?;
        opc.set_save_options(SaveOptions {
            fonts: FontEmbedding::Full,
        });
        Ok(())
    })?;
    Ok(package)
}

fn assert_custom_destination_state(package: &Package, sentinel_name: &PackURI) -> TestResult {
    package.with_opc(|opc| {
        assert_eq!(opc.save_options().fonts, FontEmbedding::Full);
        let part = opc.get_part(sentinel_name)?;
        assert_eq!(part.blob(), SENTINEL_BYTES);
        assert_eq!(
            part.rel_ref_count(SENTINEL_RELATIONSHIP_ID),
            SENTINEL_RELATIONSHIP_COUNT
        );
        Ok(())
    })
}

fn serialized(package: &mut Package) -> TestResult<Vec<u8>> {
    package.to_bytes()
}

fn assert_unsafe_edit(error: Error) {
    assert!(matches!(error, Error::UnsafeEdit { .. }));
}

fn with_eocd_comment(mut bytes: Vec<u8>, comment: &[u8]) -> TestResult<Vec<u8>> {
    let eocd_offset = bytes
        .windows(4)
        .rposition(|window| window == b"PK\x05\x06")
        .ok_or_else(|| Error::Invalid("fixture has no ZIP32 end record".into()))?;
    let eocd_end = eocd_offset
        .checked_add(22)
        .ok_or_else(|| Error::Invalid("ZIP end record offset overflow".into()))?;
    let comment_length_end = eocd_offset
        .checked_add(22)
        .ok_or_else(|| Error::Invalid("ZIP comment offset overflow".into()))?;
    let comment_length_start = comment_length_end - 2;
    let old_length_bytes = bytes
        .get(comment_length_start..comment_length_end)
        .ok_or_else(|| Error::Invalid("ZIP end record is truncated".into()))?;
    let old_length = usize::from(u16::from_le_bytes([
        old_length_bytes[0],
        old_length_bytes[1],
    ]));
    let old_end = eocd_end
        .checked_add(old_length)
        .ok_or_else(|| Error::Invalid("ZIP comment length overflow".into()))?;
    if old_end != bytes.len() {
        return Err(Error::Invalid(
            "fixture ZIP has trailing bytes after the end record".into(),
        ));
    }
    let new_length = u16::try_from(comment.len())
        .map_err(|_| Error::Invalid("fixture ZIP comment exceeds ZIP32 bounds".into()))?;
    bytes.truncate(eocd_end);
    bytes[comment_length_start..comment_length_end].copy_from_slice(&new_length.to_le_bytes());
    bytes.extend_from_slice(comment);
    Ok(bytes)
}

fn archive_comment(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes)
        .map_err(|error| Error::Invalid(format!("cannot inspect ZIP comment: {error}")))?;
    Ok(archive.comment().as_bytes().to_vec())
}

#[derive(Clone, Debug)]
struct SentinelPart {
    partname: PackURI,
    content_type: String,
    blob: Arc<Vec<u8>>,
    rels: Relationships,
}

impl SentinelPart {
    fn new(partname: PackURI) -> Self {
        Self {
            content_type: "application/octet-stream".to_owned(),
            blob: Arc::new(SENTINEL_BYTES.to_vec()),
            rels: Relationships::new(partname.base_uri().to_owned()),
            partname,
        }
    }
}

impl Part for SentinelPart {
    fn blob(&self) -> &[u8] {
        self.blob.as_slice()
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.blob)
    }

    fn content_type(&self) -> &str {
        &self.content_type
    }

    fn partname(&self) -> &PackURI {
        &self.partname
    }

    fn rel_ref_count(&self, r_id: &str) -> usize {
        if r_id == SENTINEL_RELATIONSHIP_ID {
            SENTINEL_RELATIONSHIP_COUNT
        } else {
            0
        }
    }

    fn rels(&self) -> &Relationships {
        &self.rels
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.blob = Arc::new(blob);
    }

    fn set_content_type(&mut self, content_type: String) -> litchi_opc::Result<()> {
        self.content_type = content_type;
        Ok(())
    }
}
