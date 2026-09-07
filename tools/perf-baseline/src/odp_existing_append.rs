//! Opt-in lifecycle baseline for appending one slide to an existing ODP.
//!
//! This is deliberately an owned, fully materialized control.  The timed
//! operation opens owned bytes through `edit::Snapshot::from_bytes`, appends
//! one compact title/body slide, commits the transaction, and writes the
//! committed bytes sequentially to `HashingDiscardSink`.  It does not exercise
//! a source-backed reader, bounded-memory publication, or a streaming save.
//! All semantic, preservation, patch, digest, and member-identity oracles run
//! outside the timer.

use super::{
    Case, CaseResult, Corpus, CorpusManifest, HashingDiscardSink, SemanticShape, SourceSummary,
    allocation_metrics, deterministic_sink_summary, elapsed_ns, iteration_count,
    odp_buffered_create, operation_metrics, process_metrics, record_elapsed, sha256_hex,
    statistics,
};
use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use serde::Serialize;
use sha2::{Digest as _, Sha256};
use soapberry_zip::{CompressionMethod, ZipArchive};
use std::{error::Error, fmt::Write as FmtWrite, io::Write};

/// Stable identity for the owned existing-ODP append fixture.
pub(crate) const ODP_EXISTING_APPEND_CORPUS_GENERATOR: &str =
    "litchi-odp-existing-append-lifecycle-v1";
const ODP_MIME: &str = "application/vnd.oasis.opendocument.presentation";
const ODP_XML_MEDIA_TYPE: &str = "text/xml";
pub(crate) const OPAQUE_PATH: &str = "Opaque/litchi-perf-odp-existing-append-opaque.bin";
const OPAQUE_MEDIA_TYPE: &str = "application/octet-stream";
const OPAQUE_BYTES: usize = 64 * 1024;

/// Decoded and compressed evidence for one complete ZIP member.
///
/// `compressed_sha256` covers the member's compressed data span only.  It does
/// not claim that local-header offsets, central-directory offsets, or unrelated
/// archive framing are physically unchanged after a commit.
#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
pub(crate) struct OdpExistingAppendMemberIdentity {
    pub(crate) path: String,
    pub(crate) media_type: String,
    pub(crate) compression_method: String,
    pub(crate) data_descriptor: bool,
    pub(crate) crc32: u32,
    pub(crate) decoded_bytes: usize,
    pub(crate) decoded_sha256: String,
    pub(crate) compressed_bytes: u64,
    pub(crate) compressed_sha256: String,
}

/// Untimed correctness, preservation, and lifecycle evidence for the append
/// pilot.  The source and candidate are intentionally fully materialized.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct OdpExistingAppendSummary {
    pub(crate) role: &'static str,
    pub(crate) implementation: &'static str,
    pub(crate) timing_scope: &'static str,
    pub(crate) performance_claim: &'static str,
    pub(crate) corpus_generator: &'static str,
    pub(crate) shape: &'static str,
    pub(crate) source_archive_sha256: String,
    pub(crate) source_archive_bytes: usize,
    pub(crate) output_archive_sha256: String,
    pub(crate) output_archive_bytes: usize,
    pub(crate) source_content_xml_sha256: String,
    pub(crate) source_content_xml_bytes: usize,
    pub(crate) output_content_xml_sha256: String,
    pub(crate) output_content_xml_bytes: usize,
    pub(crate) source_slide_count: usize,
    pub(crate) output_slide_count: usize,
    pub(crate) append_count: usize,
    pub(crate) appended_title_sha256: String,
    pub(crate) appended_body_sha256: String,
    pub(crate) opaque_member_path: &'static str,
    pub(crate) opaque_bytes: usize,
    pub(crate) opaque_sha256: String,
    pub(crate) opaque_member_compressed_bytes: u64,
    pub(crate) opaque_member_compressed_sha256: String,
    pub(crate) source_members: Vec<OdpExistingAppendMemberIdentity>,
    pub(crate) output_members: Vec<OdpExistingAppendMemberIdentity>,
    pub(crate) source_member_count: usize,
    pub(crate) output_member_count: usize,
    pub(crate) manifest_entry_count: usize,
    pub(crate) source_manifest_bindings_verified: bool,
    pub(crate) output_manifest_bindings_verified: bool,
    pub(crate) untouched_members_verified: bool,
    pub(crate) opaque_member_compressed_identity_verified: bool,
    pub(crate) source_semantic_reopen_verified: bool,
    pub(crate) output_semantic_reopen_verified: bool,
    pub(crate) append_exactly_one_verified: bool,
    pub(crate) source_unchanged_verified: bool,
    pub(crate) patch_replay_verified: bool,
    pub(crate) inverse_patch_verified: bool,
    pub(crate) stale_source_refusal_verified: bool,
    pub(crate) exact_noop_verified: bool,
    pub(crate) text_contract: &'static str,
    source_semantic_sha256: String,
    pub(crate) output_semantic_sha256: String,
    source_order_sha256: String,
    pub(crate) output_order_sha256: String,
    source_text_projection_sha256: String,
    pub(crate) output_text_projection_sha256: String,
    source_text_projection_bytes: usize,
    pub(crate) output_text_projection_bytes: usize,
    pub(crate) runtime_output_digest_verified: bool,
    pub(crate) runtime_sink_length_verified: bool,
    pub(crate) lifecycle_ns: Vec<u64>,
    pub(crate) output_sha256: Vec<String>,
}

/// Source fixture plus its once-built expected append candidate and gates.
#[derive(Debug)]
pub(crate) struct OdpExistingAppendCorpus {
    pub(crate) corpus: Corpus,
    pub(crate) shape: SemanticShape,
    pub(crate) source_members: Vec<OdpExistingAppendMemberIdentity>,
    expected_output: Vec<u8>,
    expected_output_members: Vec<OdpExistingAppendMemberIdentity>,
    pub(crate) source_content_xml: Vec<u8>,
    output_content_xml: Vec<u8>,
    expected_output_sha256: String,
    source_semantic_sha256: String,
    output_semantic_sha256: String,
    source_order_sha256: String,
    output_order_sha256: String,
    source_text_projection_sha256: String,
    output_text_projection_sha256: String,
    source_text_projection_bytes: usize,
    output_text_projection_bytes: usize,
    source_manifest_bindings_verified: bool,
    output_manifest_bindings_verified: bool,
    untouched_members_verified: bool,
    opaque_member_compressed_identity_verified: bool,
    patch_replay_verified: bool,
    inverse_patch_verified: bool,
    stale_source_refusal_verified: bool,
    exact_noop_verified: bool,
}

fn opaque_payload(variant: u8) -> Vec<u8> {
    let mut payload = Vec::with_capacity(OPAQUE_BYTES);
    for index in 0..OPAQUE_BYTES {
        let index_byte = index as u8;
        let page_byte = (index / 256) as u8;
        payload.push(
            index_byte
                .wrapping_mul(37)
                .wrapping_add(page_byte)
                .wrapping_add(0x5b)
                .wrapping_add(variant),
        );
    }
    payload
}

fn source_archive_from_base(
    base_archive: &[u8],
    opaque_variant: u8,
) -> Result<Vec<u8>, Box<dyn Error>> {
    let base = OwnedPackage::from_bytes(base_archive.to_vec())?;
    let content = base.get_file("content.xml")?;
    let styles = base.get_file("styles.xml")?;
    let meta = base.get_file("meta.xml")?;
    let opaque = opaque_payload(opaque_variant);

    let mut writer = PackageWriter::new();
    writer.set_mimetype(ODP_MIME)?;
    writer.add_file_with_media_type("content.xml", &content, ODP_XML_MEDIA_TYPE)?;
    writer.add_file_with_media_type("styles.xml", &styles, ODP_XML_MEDIA_TYPE)?;
    writer.add_file_with_media_type("meta.xml", &meta, ODP_XML_MEDIA_TYPE)?;
    writer.add_file_with_media_type(OPAQUE_PATH, &opaque, OPAQUE_MEDIA_TYPE)?;
    Ok(writer.finish_to_bytes()?)
}

pub(crate) fn appended_title(shape: SemanticShape) -> String {
    odp_buffered_create::odp_buffered_title(odp_buffered_create::odp_buffered_slide_count(shape))
}

pub(crate) fn appended_body(shape: SemanticShape) -> String {
    odp_buffered_create::odp_buffered_body(odp_buffered_create::odp_buffered_slide_count(shape))
}

fn append_once(source_bytes: &[u8], shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let source = litchi_odp::authoring::edit::Snapshot::from_bytes(source_bytes.to_vec())?;
    let mut transaction = source.transaction()?;
    let title = appended_title(shape);
    let body = appended_body(shape);
    transaction.add(&title, &body)?;
    let commit = transaction.commit()?;
    if !commit.changed() || commit.patch().is_noop() {
        return Err("ODP append candidate reported an exact no-op".into());
    }
    Ok(commit.snapshot().bytes().to_vec())
}

pub(crate) fn member_identities(
    bytes: &[u8],
) -> Result<Vec<OdpExistingAppendMemberIdentity>, Box<dyn Error>> {
    let package = OwnedPackage::from_bytes(bytes.to_vec())?;
    let package_view = package.package()?;
    let manifest = package_view.manifest();
    let archive = ZipArchive::from_slice(bytes)?;
    let mut records = Vec::new();
    for entry_result in archive.entries() {
        let entry = entry_result?;
        if entry.is_dir() {
            continue;
        }
        let path = entry.file_path().try_normalize()?.as_str().to_owned();
        let decoded = package.get_file(&path)?;
        let wayfinder = entry.wayfinder();
        let zip_entry = archive.get_entry(wayfinder)?;
        let (start, end) = zip_entry.compressed_data_range();
        let start = usize::try_from(start)?;
        let end = usize::try_from(end)?;
        let compressed = bytes
            .get(start..end)
            .ok_or_else(|| format!("compressed range for '{path}' is outside the archive"))?;
        records.push(OdpExistingAppendMemberIdentity {
            path: path.clone(),
            media_type: manifest
                .get_media_type(&path)
                .unwrap_or("application/octet-stream")
                .to_owned(),
            compression_method: format!("{:?}", entry.compression_method()),
            data_descriptor: entry.has_data_descriptor(),
            crc32: entry.crc32(),
            decoded_bytes: decoded.len(),
            decoded_sha256: sha256_hex(&decoded),
            compressed_bytes: u64::try_from(compressed.len())?,
            compressed_sha256: sha256_hex(compressed),
        });
    }
    records.sort_by(|left, right| left.path.cmp(&right.path));
    Ok(records)
}

pub(crate) fn find_member<'a>(
    members: &'a [OdpExistingAppendMemberIdentity],
    path: &str,
) -> Result<&'a OdpExistingAppendMemberIdentity, Box<dyn Error>> {
    members
        .iter()
        .find(|member| member.path == path)
        .ok_or_else(|| format!("ODP member '{path}' is missing").into())
}

pub(crate) fn semantic_slides(
    bytes: &[u8],
) -> Result<Vec<(Option<String>, String)>, Box<dyn Error>> {
    let presentation = litchi_odp::Presentation::from_bytes(bytes.to_vec())?;
    let slides = presentation.slides()?;
    slides
        .into_iter()
        .map(|slide| Ok((slide.title()?.map(str::to_owned), slide.text()?.to_owned())))
        .collect()
}

pub(crate) fn semantic_digest(
    slides: &[(Option<String>, String)],
) -> Result<String, Box<dyn Error>> {
    let mut hasher = Sha256::new();
    hasher.update(b"litchi-odp-buffered-semantic-v1\0");
    hasher.update(u64::try_from(slides.len())?.to_le_bytes());
    for (title, body) in slides {
        let title = title.as_deref().ok_or("ODP slide has no title")?;
        hasher.update(u64::try_from(title.len())?.to_le_bytes());
        hasher.update(title.as_bytes());
        hasher.update(u64::try_from(body.len())?.to_le_bytes());
        hasher.update(body.as_bytes());
    }
    Ok(hex_digest(hasher.finalize().as_slice()))
}

pub(crate) fn order_digest(slides: &[(Option<String>, String)]) -> Result<String, Box<dyn Error>> {
    let mut hasher = Sha256::new();
    hasher.update(b"litchi-odp-existing-append-order-v1\0");
    hasher.update(u64::try_from(slides.len())?.to_le_bytes());
    for (index, (title, body)) in slides.iter().enumerate() {
        let title = title.as_deref().ok_or("ODP slide has no title")?;
        hasher.update(u64::try_from(index)?.to_le_bytes());
        hasher.update(u64::try_from(title.len())?.to_le_bytes());
        hasher.update(title.as_bytes());
        hasher.update(u64::try_from(body.len())?.to_le_bytes());
        hasher.update(body.as_bytes());
    }
    Ok(hex_digest(hasher.finalize().as_slice()))
}

pub(crate) fn text_projection(
    slides: &[(Option<String>, String)],
) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut output = Vec::new();
    for (index, (title, body)) in slides.iter().enumerate() {
        let title = title.as_deref().ok_or("ODP slide has no title")?;
        if index != 0 {
            output.extend_from_slice(b"\n\n");
        }
        output.extend_from_slice(title.as_bytes());
        output.push(b'\n');
        output.extend_from_slice(body.as_bytes());
    }
    Ok(output)
}

fn hex_digest(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

pub(crate) fn expected_manifest_bindings(package: &OwnedPackage) -> Result<bool, Box<dyn Error>> {
    let view = package.package()?;
    let manifest = view.manifest();
    let expected = [
        ("/", ODP_MIME),
        ("content.xml", ODP_XML_MEDIA_TYPE),
        ("styles.xml", ODP_XML_MEDIA_TYPE),
        ("meta.xml", ODP_XML_MEDIA_TYPE),
        (OPAQUE_PATH, OPAQUE_MEDIA_TYPE),
    ];
    Ok(manifest.entries.len() == expected.len()
        && manifest.mimetype == ODP_MIME
        && expected.iter().all(|(path, media_type)| {
            manifest.entries.get(*path).is_some_and(|entry| {
                entry.media_type == *media_type
                    && entry.size.is_none()
                    && entry.encryption.is_none()
            })
        }))
}

pub(crate) fn expected_archive_shape(bytes: &[u8]) -> Result<bool, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let mut entries = archive.entries();
    let first = entries.next().ok_or("ODP archive has no members")??;
    let first_path = first.file_path().try_normalize()?.as_str().to_owned();
    let first_is_mimetype =
        first_path == "mimetype" && first.compression_method() == CompressionMethod::Store;
    let mut names = vec![first_path];
    let mut methods_are_contract = true;
    for entry in entries {
        let entry = entry?;
        methods_are_contract &= entry.compression_method() == CompressionMethod::Deflate;
        names.push(entry.file_path().try_normalize()?.as_str().to_owned());
    }
    names.sort();
    let expected = [
        "META-INF/manifest.xml",
        "Opaque/litchi-perf-odp-existing-append-opaque.bin",
        "content.xml",
        "meta.xml",
        "mimetype",
        "styles.xml",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect::<Vec<_>>();
    Ok(first_is_mimetype && methods_are_contract && names == expected)
}

fn expected_buffered_base_shape(bytes: &[u8]) -> Result<bool, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let mut entries = archive.entries();
    let first = entries.next().ok_or("ODP buffered base has no members")??;
    let first_path = first.file_path().try_normalize()?.as_str().to_owned();
    let mut names = vec![first_path.clone()];
    let mut methods_are_contract =
        first_path == "mimetype" && first.compression_method() == CompressionMethod::Store;
    for entry in entries {
        let entry = entry?;
        methods_are_contract &= entry.compression_method() == CompressionMethod::Deflate;
        names.push(entry.file_path().try_normalize()?.as_str().to_owned());
    }
    names.sort();
    let expected = [
        "META-INF/manifest.xml",
        "content.xml",
        "meta.xml",
        "mimetype",
        "styles.xml",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect::<Vec<_>>();
    Ok(methods_are_contract && names == expected)
}

pub(crate) fn verify_append_output(
    shape: SemanticShape,
    source_bytes: &[u8],
    output_bytes: &[u8],
    source_members: &[OdpExistingAppendMemberIdentity],
    output_members: &[OdpExistingAppendMemberIdentity],
) -> Result<(), Box<dyn Error>> {
    let source_slides = semantic_slides(source_bytes)?;
    let output_slides = semantic_slides(output_bytes)?;
    let expected_source_count = odp_buffered_create::odp_buffered_slide_count(shape);
    let expected_title = appended_title(shape);
    let expected_body = appended_body(shape);
    let expected_tail = (Some(expected_title), expected_body);
    if source_slides.len() != expected_source_count
        || output_slides.len() != expected_source_count + 1
        || output_slides[..expected_source_count] != source_slides[..]
        || output_slides.last() != Some(&expected_tail)
    {
        return Err("ODP append semantic slide order or appended slide differs".into());
    }
    for (index, (title, body)) in source_slides.iter().enumerate() {
        let expected_title = odp_buffered_create::odp_buffered_title(index);
        let expected_body = odp_buffered_create::odp_buffered_body(index);
        if title.as_deref() != Some(expected_title.as_str()) || body != &expected_body {
            return Err(
                format!("ODP source slide {index} differs from the buffered fixture").into(),
            );
        }
    }

    let source_package = OwnedPackage::from_bytes(source_bytes.to_vec())?;
    let output_package = OwnedPackage::from_bytes(output_bytes.to_vec())?;
    for path in [
        "mimetype",
        "styles.xml",
        "meta.xml",
        "META-INF/manifest.xml",
        OPAQUE_PATH,
    ] {
        if output_package.get_file(path)? != source_package.get_file(path)? {
            return Err(format!("ODP untouched member '{path}' changed decoded bytes").into());
        }
    }
    if output_package
        .package()?
        .manifest()
        .get_media_type(OPAQUE_PATH)
        != Some(OPAQUE_MEDIA_TYPE)
    {
        return Err("ODP opaque member media type changed".into());
    }

    let source_opaque = find_member(source_members, OPAQUE_PATH)?;
    let output_opaque = find_member(output_members, OPAQUE_PATH)?;
    if source_opaque.decoded_sha256 != output_opaque.decoded_sha256
        || source_opaque.decoded_bytes != output_opaque.decoded_bytes
        || source_opaque.compressed_sha256 != output_opaque.compressed_sha256
        || source_opaque.compressed_bytes != output_opaque.compressed_bytes
        || source_opaque.crc32 != output_opaque.crc32
        || source_opaque.compression_method != output_opaque.compression_method
    {
        return Err("ODP opaque member compressed or decoded identity changed".into());
    }
    Ok(())
}

fn verify_patch_contract(
    source_bytes: &[u8],
    expected_output: &[u8],
    stale_bytes: &[u8],
    shape: SemanticShape,
) -> Result<(bool, bool, bool, bool), Box<dyn Error>> {
    let source = litchi_odp::authoring::edit::Snapshot::from_bytes(source_bytes.to_vec())?;
    let mut transaction = source.transaction()?;
    let title = appended_title(shape);
    let body = appended_body(shape);
    transaction.add(&title, &body)?;
    let commit = transaction.commit()?;
    let replayed = commit.patch().apply(&source)?;
    let replay_ok = replayed.bytes() == expected_output;
    let inverse_ok = commit.patch().inverse().apply(&replayed)?.bytes() == source.bytes();
    let stale = litchi_odp::authoring::edit::Snapshot::from_bytes(stale_bytes.to_vec())?;
    let stale_refused = matches!(
        commit.patch().apply(&stale),
        Err(litchi_core::Error::InvalidFormat(message))
            if message == "stale ODP presentation patch source"
    );

    let no_op = source.transaction()?.commit()?;
    let exact_noop = !no_op.changed()
        && no_op.patch().is_noop()
        && no_op.snapshot().bytes().len() == source.bytes().len()
        && no_op.snapshot().bytes().as_ptr() == source.bytes().as_ptr()
        && no_op.snapshot().bytes() == source.bytes();
    if !replay_ok || !inverse_ok || !stale_refused || !exact_noop {
        return Err("ODP append patch or exact-noop gate failed".into());
    }
    Ok((replay_ok, inverse_ok, stale_refused, exact_noop))
}

/// Build and gate one deterministic existing-ODP append corpus.
pub(crate) fn build_odp_existing_append_corpus(
    shape: SemanticShape,
) -> Result<OdpExistingAppendCorpus, Box<dyn Error>> {
    let base = odp_buffered_create::build_odp_buffered_corpus(shape)?;
    if base.manifest.archive_member_count != 5 || !expected_buffered_base_shape(&base.archive)? {
        return Err(
            "ODP append base fixture no longer has the buffered five-member topology".into(),
        );
    }
    let base_package = OwnedPackage::from_bytes(base.archive.clone())?;
    if sha256_hex(&base_package.get_file("styles.xml")?)
        != odp_buffered_create::ODP_BUFFERED_DEFAULT_STYLES_SHA256
        || sha256_hex(&base_package.get_file("meta.xml")?)
            != odp_buffered_create::ODP_BUFFERED_DEFAULT_META_SHA256
    {
        return Err("ODP append base fixture styles/meta identity changed".into());
    }
    let slide_count = base.manifest.entry_count;
    let source_archive = source_archive_from_base(&base.archive, 0)?;
    let source_package = OwnedPackage::from_bytes(source_archive.clone())?;
    let source_content_xml = source_package.get_file("content.xml")?;
    if source_package.get_file(OPAQUE_PATH)? != opaque_payload(0) {
        return Err("ODP append opaque source bytes differ from the byte formula".into());
    }
    let source_members = member_identities(&source_archive)?;
    if source_members.len() != 6 || !expected_archive_shape(&source_archive)? {
        return Err("ODP append source archive member shape differs from the contract".into());
    }

    let expected_output = append_once(&source_archive, shape)?;
    let expected_output_members = member_identities(&expected_output)?;
    let output_package = OwnedPackage::from_bytes(expected_output.clone())?;
    let output_content_xml = output_package.get_file("content.xml")?;
    if expected_output_members.len() != 6 || !expected_archive_shape(&expected_output)? {
        return Err("ODP append output archive member shape differs from the contract".into());
    }
    verify_append_output(
        shape,
        &source_archive,
        &expected_output,
        &source_members,
        &expected_output_members,
    )?;

    let stale_archive = source_archive_from_base(&base.archive, 1)?;
    let (
        patch_replay_verified,
        inverse_patch_verified,
        stale_source_refusal_verified,
        exact_noop_verified,
    ) = verify_patch_contract(&source_archive, &expected_output, &stale_archive, shape)?;

    let source_manifest_bindings_verified = expected_manifest_bindings(&source_package)?;
    let output_manifest_bindings_verified = expected_manifest_bindings(&output_package)?;
    if !source_manifest_bindings_verified || !output_manifest_bindings_verified {
        return Err("ODP append manifest bindings differ from the contract".into());
    }
    let mut untouched_members_verified = true;
    for path in [
        "mimetype",
        "styles.xml",
        "meta.xml",
        "META-INF/manifest.xml",
        OPAQUE_PATH,
    ] {
        untouched_members_verified &=
            output_package.get_file(path)? == source_package.get_file(path)?;
    }
    let source_opaque = find_member(&source_members, OPAQUE_PATH)?;
    let output_opaque = find_member(&expected_output_members, OPAQUE_PATH)?;
    let opaque_member_compressed_identity_verified = source_opaque == output_opaque;
    if !untouched_members_verified || !opaque_member_compressed_identity_verified {
        return Err("ODP append untouched member identity changed".into());
    }

    let source_slides = semantic_slides(&source_archive)?;
    let output_slides = semantic_slides(&expected_output)?;
    let source_semantic_sha256 = semantic_digest(&source_slides)?;
    let output_semantic_sha256 = semantic_digest(&output_slides)?;
    let source_order_sha256 = order_digest(&source_slides)?;
    let output_order_sha256 = order_digest(&output_slides)?;
    let source_text_projection = text_projection(&source_slides)?;
    let output_text_projection = text_projection(&output_slides)?;
    let source_text_projection_sha256 = sha256_hex(&source_text_projection);
    let output_text_projection_sha256 = sha256_hex(&output_text_projection);

    let target_payload = source_content_xml.clone();
    let first_title = odp_buffered_create::odp_buffered_title(0);
    let first_body = odp_buffered_create::odp_buffered_body(0);
    let source_manifest = CorpusManifest {
        name: format!("odp-existing-append-lifecycle-{}", shape.name()),
        generator: ODP_EXISTING_APPEND_CORPUS_GENERATOR,
        package_format: "ODP/ODF/ZIP",
        shape: shape.name(),
        payload_kind: "deterministic-buffered-plain-titled-slides-with-opaque-member",
        compression: "mimetype=stored;xml=deflate;opaque=deflate",
        entry_count: slide_count,
        archive_member_count: source_members.len(),
        entry_bytes: first_title.len() + 1 + first_body.len(),
        uncompressed_payload_bytes: source_text_projection.len(),
        archive_bytes: source_archive.len(),
        archive_sha256: sha256_hex(&source_archive),
        target_entry: "content.xml".to_owned(),
        target_payload_bytes: target_payload.len(),
        target_payload_sha256: sha256_hex(&target_payload),
        rtf_variant: None,
        xlsx: None,
    };
    Ok(OdpExistingAppendCorpus {
        corpus: Corpus {
            manifest: source_manifest,
            archive: source_archive,
            target_name: "content.xml".to_owned(),
            target_payload,
            xlsx: None,
        },
        shape,
        source_members,
        expected_output_sha256: sha256_hex(&expected_output),
        expected_output,
        expected_output_members,
        source_content_xml,
        output_content_xml,
        source_semantic_sha256,
        output_semantic_sha256,
        source_order_sha256,
        output_order_sha256,
        source_text_projection_sha256,
        output_text_projection_sha256,
        source_text_projection_bytes: source_text_projection.len(),
        output_text_projection_bytes: output_text_projection.len(),
        source_manifest_bindings_verified,
        output_manifest_bindings_verified,
        untouched_members_verified,
        opaque_member_compressed_identity_verified,
        patch_replay_verified,
        inverse_patch_verified,
        stale_source_refusal_verified,
        exact_noop_verified,
    })
}

fn sink_observation(sink: super::SinkSummary) -> operation_metrics::SinkObservation {
    operation_metrics::SinkObservation {
        accepted_bytes: sink.accepted_bytes,
        write_calls: sink.write_calls,
        largest_write: sink.largest_write,
        bytes_0: sink.write_size_buckets.bytes_0,
        bytes_1_to_512: sink.write_size_buckets.bytes_1_to_512,
        bytes_513_to_4096: sink.write_size_buckets.bytes_513_to_4096,
        bytes_4097_to_16384: sink.write_size_buckets.bytes_4097_to_16384,
        bytes_16385_to_65536: sink.write_size_buckets.bytes_16385_to_65536,
        bytes_over_65536: sink.write_size_buckets.bytes_over_65536,
    }
}

/// Run the fully materialized existing-ODP append lifecycle.
pub(crate) fn run_odp_existing_append_lifecycle(
    case: Case,
    corpus: &OdpExistingAppendCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::OdpExistingAppendLifecycle
        || corpus.corpus.manifest.generator != ODP_EXISTING_APPEND_CORPUS_GENERATOR
    {
        return Err("non-existing-append case passed to ODP append runner".into());
    }
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut output_digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    let iterations = iteration_count(warmup_iterations, samples)?;
    for iteration in 0..iterations {
        // The owned input clone, append strings, and sink are outside the
        // timer. Snapshot opening is intentionally inside this selector's
        // measured open-through-publication lifecycle. Keep source and
        // candidate alive through both endpoint snapshots so allocation/process
        // deltas include the retained end state instead of measuring a release
        // as a negative publication allocation.
        let input = corpus.corpus.archive.clone();
        let title = appended_title(corpus.shape);
        let body = appended_body(corpus.shape);
        let mut sink = HashingDiscardSink::without_authoring_window(u64::try_from(
            corpus.expected_output.len(),
        )?);
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let started = std::time::Instant::now();
        let source = litchi_odp::authoring::edit::Snapshot::from_bytes(input)?;
        let mut transaction = source.transaction()?;
        transaction.add(&title, &body)?;
        let commit = transaction.commit()?;
        sink.write_all(commit.snapshot().bytes())?;
        let duration = started.elapsed();
        let allocation_metrics = allocation_region.finish();
        let process_after = process_metrics::Snapshot::read().ok();

        let candidate = commit.snapshot().bytes();
        if candidate != corpus.expected_output || source.bytes() != corpus.corpus.archive {
            return Err("ODP append lifecycle candidate or source lineage changed".into());
        }
        let candidate_digest = sha256_hex(candidate);
        let (mut summary, sink_digest) = sink.finish();
        if sink_digest != candidate_digest
            || summary.accepted_bytes != u64::try_from(candidate.len())?
        {
            return Err("ODP append lifecycle sink digest or length differs from candidate".into());
        }
        // These generic sink fields describe the full committed semantic
        // projection and content.xml size, not physical input or add arguments.
        summary.input_bytes = Some(u64::try_from(corpus.output_text_projection_bytes)?);
        summary.authored_part_bytes = Some(u64::try_from(corpus.output_content_xml.len())?);
        std::hint::black_box(candidate_digest.as_str());
        drop(commit);
        drop(source);

        if iteration >= warmup_iterations {
            sink_summaries.push(summary);
            output_digests.push(candidate_digest);
            observations.push(operation_metrics::InProcessObservation {
                elapsed_ns: elapsed_ns(duration)?,
                process_metrics: process_before
                    .zip(process_after)
                    .map(|(before, after)| after.delta(before)),
                allocation_metrics,
            });
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, "ODP existing append lifecycle")?;
    if sink.retained_output_bytes != Some(0) || sink.retained_authoring_window_bytes.is_some() {
        return Err("ODP append lifecycle unexpectedly retained output/window bytes".into());
    }
    if output_digests
        .iter()
        .any(|digest| digest != &corpus.expected_output_sha256)
    {
        return Err("ODP append lifecycle output digest changed across samples".into());
    }
    let operation_metrics = Some(operation_metrics::from_in_process_observations(
        &observations,
        sink_observation(sink),
    )?);
    let source = SourceSummary {
        odp_append: Some(OdpExistingAppendSummary {
            role: "existing_append",
            implementation: "litchi_odp::authoring::edit::Snapshot::from_bytes + Transaction::add/commit",
            timing_scope: "owned input clone, append strings, and HashingDiscardSink construction outside; the clock includes Snapshot::from_bytes, Snapshot::transaction, exactly one transaction.add, transaction.commit, and sequential committed-snapshot write_all; digest finalization, output/source/patch/member oracles, result assembly, and dropping source/candidate occur after the clock and endpoint snapshots",
            performance_claim: "owned fully materialized existing-ODP lifecycle evidence only; no source-backed, bounded-memory, or semantic-streaming-save claim; allocation live delta intentionally retains source and candidate at endpoint snapshots",
            corpus_generator: ODP_EXISTING_APPEND_CORPUS_GENERATOR,
            shape: corpus.shape.name(),
            source_archive_sha256: corpus.corpus.manifest.archive_sha256.clone(),
            source_archive_bytes: corpus.corpus.archive.len(),
            output_archive_sha256: corpus.expected_output_sha256.clone(),
            output_archive_bytes: corpus.expected_output.len(),
            source_content_xml_sha256: sha256_hex(&corpus.source_content_xml),
            source_content_xml_bytes: corpus.source_content_xml.len(),
            output_content_xml_sha256: sha256_hex(&corpus.output_content_xml),
            output_content_xml_bytes: corpus.output_content_xml.len(),
            source_slide_count: odp_buffered_create::odp_buffered_slide_count(corpus.shape),
            output_slide_count: odp_buffered_create::odp_buffered_slide_count(corpus.shape) + 1,
            append_count: 1,
            appended_title_sha256: sha256_hex(appended_title(corpus.shape).as_bytes()),
            appended_body_sha256: sha256_hex(appended_body(corpus.shape).as_bytes()),
            opaque_member_path: OPAQUE_PATH,
            opaque_bytes: find_member(&corpus.source_members, OPAQUE_PATH)?.decoded_bytes,
            opaque_sha256: find_member(&corpus.source_members, OPAQUE_PATH)?
                .decoded_sha256
                .clone(),
            opaque_member_compressed_bytes: find_member(&corpus.source_members, OPAQUE_PATH)?
                .compressed_bytes,
            opaque_member_compressed_sha256: find_member(&corpus.source_members, OPAQUE_PATH)?
                .compressed_sha256
                .clone(),
            source_members: corpus.source_members.clone(),
            output_members: corpus.expected_output_members.clone(),
            source_member_count: corpus.source_members.len(),
            output_member_count: corpus.expected_output_members.len(),
            manifest_entry_count: 5,
            source_manifest_bindings_verified: corpus.source_manifest_bindings_verified,
            output_manifest_bindings_verified: corpus.output_manifest_bindings_verified,
            untouched_members_verified: corpus.untouched_members_verified,
            opaque_member_compressed_identity_verified: corpus
                .opaque_member_compressed_identity_verified,
            source_semantic_reopen_verified: true,
            output_semantic_reopen_verified: true,
            append_exactly_one_verified: true,
            source_unchanged_verified: true,
            patch_replay_verified: corpus.patch_replay_verified,
            inverse_patch_verified: corpus.inverse_patch_verified,
            stale_source_refusal_verified: corpus.stale_source_refusal_verified,
            exact_noop_verified: corpus.exact_noop_verified,
            text_contract: "UTF-8 mixed Unicode/entities/plain text; one interior ASCII space; no CR, LF, tab, edge space, or repeated interior spaces",
            source_semantic_sha256: corpus.source_semantic_sha256.clone(),
            output_semantic_sha256: corpus.output_semantic_sha256.clone(),
            source_order_sha256: corpus.source_order_sha256.clone(),
            output_order_sha256: corpus.output_order_sha256.clone(),
            source_text_projection_sha256: corpus.source_text_projection_sha256.clone(),
            output_text_projection_sha256: corpus.output_text_projection_sha256.clone(),
            source_text_projection_bytes: corpus.source_text_projection_bytes,
            output_text_projection_bytes: corpus.output_text_projection_bytes,
            runtime_output_digest_verified: true,
            runtime_sink_length_verified: true,
            lifecycle_ns: elapsed.clone(),
            output_sha256: output_digests.clone(),
        }),
        ..SourceSummary::default()
    };
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(Box::new(source)),
        execution: None,
        output_sha256: Some(corpus.expected_output_sha256.clone()),
        operation_metrics,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn opaque_fixture_uses_authoritative_formula() {
        let bytes = opaque_payload(0);
        assert_eq!(bytes.len(), OPAQUE_BYTES);
        for index in [0, 1, 255, 256, 65_535] {
            let expected = (index as u8)
                .wrapping_mul(37)
                .wrapping_add((index / 256) as u8)
                .wrapping_add(0x5b);
            assert_eq!(bytes[index], expected);
        }
        assert_ne!(bytes, opaque_payload(1));
    }

    #[test]
    fn tiny_append_fixture_has_exact_output_and_patch_gates() {
        let corpus = build_odp_existing_append_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(
            corpus.corpus.manifest.generator,
            ODP_EXISTING_APPEND_CORPUS_GENERATOR
        );
        assert_eq!(
            corpus.corpus.manifest.name,
            "odp-existing-append-lifecycle-tiny"
        );
        assert_eq!(corpus.corpus.manifest.entry_count, 64);
        assert_eq!(corpus.source_members.len(), 6);
        assert_eq!(corpus.expected_output_members.len(), 6);
        assert!(corpus.source_manifest_bindings_verified);
        assert!(corpus.output_manifest_bindings_verified);
        assert!(corpus.untouched_members_verified);
        assert!(corpus.opaque_member_compressed_identity_verified);
        assert!(corpus.patch_replay_verified);
        assert!(corpus.inverse_patch_verified);
        assert!(corpus.stale_source_refusal_verified);
        assert!(corpus.exact_noop_verified);
        assert_eq!(
            corpus.source_content_xml.len(),
            corpus.corpus.manifest.target_payload_bytes
        );
        assert_ne!(corpus.corpus.archive, corpus.expected_output);
    }

    #[test]
    fn runner_is_opt_in_and_binds_actual_commit_bytes_to_sink() {
        let corpus = build_odp_existing_append_corpus(SemanticShape::Tiny).unwrap();
        let result =
            run_odp_existing_append_lifecycle(Case::OdpExistingAppendLifecycle, &corpus, 0, 1)
                .unwrap();
        assert_eq!(
            result.output_sha256.as_deref(),
            Some(corpus.expected_output_sha256.as_str())
        );
        let sink = result.sink.unwrap();
        assert_eq!(sink.accepted_bytes as usize, corpus.expected_output.len());
        assert_eq!(
            sink.input_bytes,
            Some(corpus.output_text_projection_bytes as u64)
        );
        assert_eq!(
            sink.authored_part_bytes,
            Some(corpus.output_content_xml.len() as u64)
        );
        let summary = result.source.unwrap().odp_append.unwrap();
        assert!(summary.append_exactly_one_verified);
        assert!(summary.runtime_output_digest_verified);
        assert!(summary.runtime_sink_length_verified);
    }

    #[test]
    fn output_oracle_rejects_missing_append_and_changed_opaque_member() {
        let shape = SemanticShape::Tiny;
        let corpus = build_odp_existing_append_corpus(shape).unwrap();
        assert!(
            verify_append_output(
                shape,
                &corpus.corpus.archive,
                &corpus.corpus.archive,
                &corpus.source_members,
                &corpus.source_members,
            )
            .is_err()
        );
        let base = odp_buffered_create::build_odp_buffered_corpus(shape).unwrap();
        let changed_source = source_archive_from_base(&base.archive, 1).unwrap();
        let changed_output = append_once(&changed_source, shape).unwrap();
        let changed_members = member_identities(&changed_output).unwrap();
        let error = verify_append_output(
            shape,
            &corpus.corpus.archive,
            &changed_output,
            &corpus.source_members,
            &changed_members,
        )
        .unwrap_err();
        assert!(error.to_string().contains("untouched member"));
    }
}
