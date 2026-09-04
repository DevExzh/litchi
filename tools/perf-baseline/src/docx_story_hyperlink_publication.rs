//! Correctness-only end-to-end DOCX story-hyperlink publication evidence.
//!
//! This runner intentionally owns a corpus separate from the plan-only story
//! hyperlink case.  The corpus exercises every relationship-reachable Word
//! story kind, keeps one shared external target selected in every story, and
//! retains unrelated links, media, and opaque package members so publication
//! locality can be checked independently of the typed DOCX inventory.

use super::{Case, CaseResult, CorpusManifest, CountingSink, SinkSummary, SourceSummary};
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_docx::source_backed;
use litchi_docx::story_hyperlinks::Mode;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::reader::NsReader;
use serde::Serialize;
use soapberry_zip::office::ArchiveReader;
use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;
use std::io::{self, Write};
use std::sync::{Arc, atomic::AtomicU64, atomic::Ordering};

const CORPUS_GENERATOR: &str = "litchi-docx-story-hyperlink-publication-v1";
const CORPUS_SHAPE: &str = "7-story-kinds-14-links-7-media-1-opaque";
const PAYLOAD_KIND: &str = "deterministic-inert-media-and-opaque-members";
const SHARED_TARGET: &str = "https://litchi-perf.invalid/story-shared";
const OPAQUE_MEMBER: &str = "word/opaque/story-hyperlink-sentinel.bin";
const MEDIA_BYTES: usize = 4096;
const OPAQUE_BYTES: usize = 8192;
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OPC_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const GLOSSARY_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const HYPERLINK_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const IMAGE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HEADER_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const FOOTNOTES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const ENDNOTES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const COMMENTS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const RELATIONSHIP_MEMBER_PREFIX: &str = "word/_rels/";
const EXPECTED_ENTRY_COUNT: usize = 15;
const EXPECTED_ARCHIVE_MEMBER_COUNT: usize = 24;
const EXPECTED_ARCHIVE_BYTES: usize = 9_900;
const EXPECTED_ARCHIVE_SHA256: &str =
    "457421e8f86ec8eb52fbe181cebe7d0821ce1e794a08142ff01a4c4e03df0cac";

/// Predeclared candidate-minus-control allocator deltas for the seven
/// Deflate story relationship reads. These values are an expected model only;
/// this correctness-only runner does not observe or claim allocator results.
#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
struct StoryRelationshipAllocatorModel {
    comparison: &'static str,
    status: &'static str,
    allocation_calls: i64,
    deallocation_calls: i64,
    reallocation_calls: i64,
    failed_allocation_calls: i64,
    allocated_bytes: i64,
    deallocated_bytes: i64,
}

const STORY_RELATIONSHIP_ALLOCATOR_MODEL: StoryRelationshipAllocatorModel =
    StoryRelationshipAllocatorModel {
        comparison: "candidate-control",
        status: "expected_not_observed",
        allocation_calls: -12,
        deallocation_calls: -12,
        reallocation_calls: 0,
        failed_allocation_calls: 0,
        allocated_bytes: -481_920,
        deallocated_bytes: -481_920,
    };

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct StorySpec {
    kind: &'static str,
    file: &'static str,
    root: &'static str,
    content_type: &'static str,
    relationship_type: &'static str,
}

const STORIES: [StorySpec; 7] = [
    StorySpec {
        kind: "main",
        file: "document.xml",
        root: "document",
        content_type: ct::WML_DOCUMENT_MAIN,
        relationship_type: rt::OFFICE_DOCUMENT,
    },
    StorySpec {
        kind: "header",
        file: "header1.xml",
        root: "hdr",
        content_type: ct::WML_HEADER,
        relationship_type: rt::HEADER,
    },
    StorySpec {
        kind: "footer",
        file: "footer1.xml",
        root: "ftr",
        content_type: ct::WML_FOOTER,
        relationship_type: rt::FOOTER,
    },
    StorySpec {
        kind: "footnotes",
        file: "footnotes.xml",
        root: "footnotes",
        content_type: ct::WML_FOOTNOTES,
        relationship_type: rt::FOOTNOTES,
    },
    StorySpec {
        kind: "endnotes",
        file: "endnotes.xml",
        root: "endnotes",
        content_type: ct::WML_ENDNOTES,
        relationship_type: rt::ENDNOTES,
    },
    StorySpec {
        kind: "comments",
        file: "comments.xml",
        root: "comments",
        content_type: ct::WML_COMMENTS,
        relationship_type: rt::COMMENTS,
    },
    StorySpec {
        kind: "glossary",
        file: "glossary.xml",
        root: "glossaryDocument",
        content_type: ct::WML_DOCUMENT_GLOSSARY,
        relationship_type: GLOSSARY_REL,
    },
];

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
pub(super) struct DocxStoryHyperlinkPublicationSummary {
    implementation: &'static str,
    timing_scope: &'static str,
    performance_claim: &'static str,
    /// Candidate-minus-control allocator model when this case performs the
    /// seven Deflate relationship reads; no-op publication has no delta.
    predeclared_allocator_model: Option<StoryRelationshipAllocatorModel>,
    story_kinds: Vec<String>,
    selected_target: String,
    selected_relationship_count: usize,
    unselected_relationship_count: usize,
    source_archive_bytes: u64,
    source_archive_sha256: String,
    output_archive_bytes: u64,
    output_archive_sha256: String,
    end_to_end_ns: Vec<u64>,
    source_zip_oracle_verified: bool,
    source_hash_verified: bool,
    no_op_exact_bytes_verified: bool,
    changed_member_locality_verified: bool,
    relationship_oracle_verified: bool,
    xml_semantic_oracle_verified: bool,
    deterministic_output_verified: bool,
    source_immutability_verified: bool,
    stale_source_refusal_verified: bool,
    foreign_source_refusal_verified: bool,
    signed_refusal_verified: bool,
    unknown_owner_refusal_verified: bool,
    partial_sink_refusal_verified: bool,
    zero_sink_refusal_verified: bool,
}

#[derive(Debug)]
pub(super) struct Corpus {
    manifest: CorpusManifest,
    archive: Vec<u8>,
    source_members: BTreeMap<String, Vec<u8>>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct StoryOracle {
    shared_links: usize,
    unselected_links: usize,
    media_refs: usize,
    text_nodes: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RelationshipOracleEntry {
    id: String,
    relationship_type: String,
    target: String,
    target_mode: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceIdentity {
    version: SourceVersion,
    length: u64,
    sha256: String,
}

impl SourceIdentity {
    fn capture(source: &OwnedSource) -> Result<Self, Box<dyn std::error::Error>> {
        Ok(Self {
            version: source.version()?,
            length: source.len()?,
            sha256: super::sha256_hex(source.as_slice()),
        })
    }

    fn matches(&self, source: &OwnedSource) -> Result<bool, Box<dyn std::error::Error>> {
        Ok(self.version == source.version()?
            && self.length == source.len()?
            && self.sha256 == super::sha256_hex(source.as_slice()))
    }
}

#[derive(Debug)]
struct FailingSink {
    accepted: usize,
    fail_after: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.fail_after {
            return Err(io::Error::other("intentional story publication sink failure"));
        }
        let remaining = self.fail_after - self.accepted;
        let accepted = remaining.min(bytes.len());
        self.accepted = self
            .accepted
            .checked_add(accepted)
            .ok_or_else(|| io::Error::other("story publication sink progress overflow"))?;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Vec<u8>,
    revision: Arc<AtomicU64>,
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("story hyperlink source length overflow"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5148_5950,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn media_name(index: usize) -> String {
    format!("word/media/story-hyperlink-{index:02}.png")
}

fn story_part_name(file: &str) -> String {
    format!("word/{file}")
}

fn story_rels_name(file: &str) -> String {
    format!("word/_rels/{file}.rels")
}

fn media_payload(index: usize) -> Vec<u8> {
    (0..MEDIA_BYTES)
        .map(|offset| ((offset + index * 17) % 251) as u8)
        .collect()
}

fn opaque_payload() -> Vec<u8> {
    (0..OPAQUE_BYTES)
        .map(|offset| ((offset * 31 + 7) % 251) as u8)
        .collect()
}

fn story_xml(story: StorySpec, index: usize) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut xml = String::new();
    xml.try_reserve(1_024)
        .map_err(|error| format!("DOCX story XML allocation failed: {error}"))?;
    write!(
        xml,
        r#"<w:{root} xmlns:w="{W}" xmlns:r="{R}"><w:p><w:hyperlink r:id="rShared"><w:r><w:t>{kind}-selected</w:t></w:r></w:hyperlink><w:hyperlink r:id="rOther"><w:r><w:t>{kind}-unselected</w:t></w:r></w:hyperlink><w:r><w:drawing r:embed="rMedia"/></w:r></w:p></w:{root}>"#,
        root = story.root,
        kind = story.kind,
    )?;
    let _ = index;
    Ok(xml.into_bytes())
}

fn main_xml() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    story_xml(STORIES[0], 0)
}

fn add_story_relationships(
    story: &mut BlobPart,
    story_index: usize,
) -> Result<(), Box<dyn std::error::Error>> {
    story.rels_mut().try_add_relationship(
        rt::HYPERLINK.to_owned(),
        SHARED_TARGET.to_owned(),
        "rShared".to_owned(),
        TargetMode::External,
    )?;
    story.rels_mut().try_add_relationship(
        rt::HYPERLINK.to_owned(),
        format!(
            "https://litchi-perf.invalid/story-{}/unselected",
            STORIES[story_index].kind
        ),
        "rOther".to_owned(),
        TargetMode::External,
    )?;
    story.rels_mut().try_add_relationship(
        rt::IMAGE.to_owned(),
        format!("media/story-hyperlink-{story_index:02}.png"),
        "rMedia".to_owned(),
        TargetMode::Internal,
    )?;
    Ok(())
}

fn build_archive() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml")?,
        STORIES[0].content_type.to_owned(),
        main_xml()?,
    );
    add_story_relationships(&mut main, 0)?;

    for (story_index, story_spec) in STORIES.iter().enumerate().skip(1) {
        main.rels_mut().try_add_relationship(
            story_spec.relationship_type.to_owned(),
            story_spec.file.to_owned(),
            format!("rStory{story_index:02}"),
            TargetMode::Internal,
        )?;
        let mut story = BlobPart::new(
            PackURI::new(format!("/{}", story_part_name(story_spec.file)))?,
            story_spec.content_type.to_owned(),
            story_xml(*story_spec, story_index)?,
        );
        add_story_relationships(&mut story, story_index)?;
        package.try_add_part(Box::new(story))?;
    }

    for story_index in 0..STORIES.len() {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{}", media_name(story_index)))?,
            ct::PNG.to_owned(),
            media_payload(story_index),
        )))?;
    }
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(format!("/{OPAQUE_MEMBER}"))?,
        "application/octet-stream".to_owned(),
        opaque_payload(),
    )))?;
    package.try_add_part(Box::new(main))?;
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

fn archive_members(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>, Box<dyn std::error::Error>> {
    let archive = ArchiveReader::new(bytes)?;
    let mut members = BTreeMap::new();
    for name in archive.file_names() {
        members.insert(name.to_owned(), archive.read(name)?);
    }
    Ok(members)
}

fn parse_relationship_entry(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<RelationshipOracleEntry, Box<dyn std::error::Error>> {
    let mut id = None;
    let mut relationship_type = None;
    let mut target = None;
    let mut target_mode = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)?
            .into_owned();
        match attribute.key.as_ref() {
            b"Id" => id = Some(value),
            b"Type" => relationship_type = Some(value),
            b"Target" => target = Some(value),
            b"TargetMode" => target_mode = Some(value),
            _ => return Err("DOCX relationship oracle saw an unexpected attribute".into()),
        }
    }
    Ok(RelationshipOracleEntry {
        id: id.ok_or("DOCX relationship oracle found a relationship without Id")?,
        relationship_type: relationship_type
            .ok_or("DOCX relationship oracle found a relationship without Type")?,
        target: target.ok_or("DOCX relationship oracle found a relationship without Target")?,
        target_mode: target_mode.unwrap_or_else(|| "Internal".to_owned()),
    })
}

fn validate_relationship_root(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut default_namespace = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        if attribute.key.as_ref() != b"xmlns" {
            return Err("DOCX relationship oracle saw an unexpected root attribute".into());
        }
        if default_namespace.is_some() {
            return Err("DOCX relationship oracle saw duplicate default namespaces".into());
        }
        default_namespace = Some(
            attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)?
                .into_owned(),
        );
    }
    if default_namespace.as_deref() != Some(OPC_RELATIONSHIPS_NAMESPACE) {
        return Err("DOCX relationship oracle saw the wrong OPC namespace".into());
    }
    Ok(())
}

fn parse_relationship_oracle(
    xml: &[u8],
) -> Result<BTreeMap<String, RelationshipOracleEntry>, Box<dyn std::error::Error>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut relationships = BTreeMap::new();
    loop {
        let (namespace, event) = reader.read_resolved_event()?;
        let is_opc_namespace = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if value == OPC_RELATIONSHIPS_NAMESPACE.as_bytes()
        );
        match event {
            Event::Start(element) => {
                let local_name = element.local_name();
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || !is_opc_namespace
                        || local_name.as_ref() != b"Relationships"
                    {
                        return Err("DOCX relationship oracle found an invalid root".into());
                    }
                    validate_relationship_root(&element, reader.decoder())?;
                    root_seen = true;
                } else if !is_opc_namespace {
                    return Err("DOCX relationship oracle found a foreign namespace".into());
                }
                depth = depth
                    .checked_add(1)
                    .ok_or("DOCX relationship oracle depth overflow")?;
                if local_name.as_ref() == b"Relationship" {
                    if depth != 2 {
                        return Err("DOCX relationship oracle found a nested relationship".into());
                    }
                    let relationship = parse_relationship_entry(&element, reader.decoder())?;
                    if relationships
                        .insert(relationship.id.clone(), relationship)
                        .is_some()
                    {
                        return Err("DOCX relationship oracle found duplicate Id".into());
                    }
                } else if depth == 2 {
                    return Err("DOCX relationship oracle found an unexpected child".into());
                }
            },
            Event::Empty(element) => {
                let local_name = element.local_name();
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || !is_opc_namespace
                        || local_name.as_ref() != b"Relationships"
                    {
                        return Err("DOCX relationship oracle found an invalid root".into());
                    }
                    validate_relationship_root(&element, reader.decoder())?;
                    root_seen = true;
                    root_closed = true;
                } else {
                    if !is_opc_namespace {
                        return Err("DOCX relationship oracle found a foreign namespace".into());
                    }
                    if local_name.as_ref() == b"Relationship" {
                        if depth != 1 {
                            return Err(
                                "DOCX relationship oracle found a nested relationship".into(),
                            );
                        }
                        let relationship =
                            parse_relationship_entry(&element, reader.decoder())?;
                        if relationships
                            .insert(relationship.id.clone(), relationship)
                            .is_some()
                        {
                            return Err("DOCX relationship oracle found duplicate Id".into());
                        }
                    } else if depth == 1 {
                        return Err("DOCX relationship oracle found a nested relationship".into());
                    }
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err("DOCX relationship oracle found an unexpected end".into());
                }
                depth -= 1;
                if depth == 0 {
                    if !is_opc_namespace || element.local_name().as_ref() != b"Relationships" {
                        return Err("DOCX relationship oracle found an invalid root close".into());
                    }
                    root_closed = true;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err("DOCX relationship oracle found an incomplete relationship document".into());
    }
    Ok(relationships)
}

fn expected_story_relationship_members() -> BTreeSet<String> {
    STORIES
        .iter()
        .map(|story| story_rels_name(story.file))
        .collect()
}

fn expected_owner_relationship_type(story: StorySpec) -> &'static str {
    match story.kind {
        "header" => HEADER_RELATIONSHIP_TYPE,
        "footer" => FOOTER_RELATIONSHIP_TYPE,
        "footnotes" => FOOTNOTES_RELATIONSHIP_TYPE,
        "endnotes" => ENDNOTES_RELATIONSHIP_TYPE,
        "comments" => COMMENTS_RELATIONSHIP_TYPE,
        "glossary" => GLOSSARY_REL,
        _ => unreachable!("only non-main stories have owner relationship types"),
    }
}

fn expected_story_relationships(
    story: StorySpec,
    story_index: usize,
    selected: bool,
) -> BTreeMap<String, RelationshipOracleEntry> {
    let mut expected = BTreeMap::from([
        (
            "rOther".to_owned(),
            RelationshipOracleEntry {
                id: "rOther".to_owned(),
                relationship_type: HYPERLINK_RELATIONSHIP_TYPE.to_owned(),
                target: format!(
                    "https://litchi-perf.invalid/story-{}/unselected",
                    story.kind
                ),
                target_mode: "External".to_owned(),
            },
        ),
        (
            "rMedia".to_owned(),
            RelationshipOracleEntry {
                id: "rMedia".to_owned(),
                relationship_type: IMAGE_RELATIONSHIP_TYPE.to_owned(),
                target: format!("media/story-hyperlink-{story_index:02}.png"),
                target_mode: "Internal".to_owned(),
            },
        ),
    ]);
    if !selected {
        expected.insert(
            "rShared".to_owned(),
            RelationshipOracleEntry {
                id: "rShared".to_owned(),
                relationship_type: HYPERLINK_RELATIONSHIP_TYPE.to_owned(),
                target: SHARED_TARGET.to_owned(),
                target_mode: "External".to_owned(),
            },
        );
    }
    if story_index == 0 {
        for (index, owner_story) in STORIES.iter().enumerate().skip(1) {
            let id = format!("rStory{index:02}");
            expected.insert(
                id.clone(),
                RelationshipOracleEntry {
                    id,
                    relationship_type: expected_owner_relationship_type(*owner_story).to_owned(),
                    target: owner_story.file.to_owned(),
                    target_mode: "Internal".to_owned(),
                },
            );
        }
    }
    expected
}

fn verify_relationship_oracle(
    members: &BTreeMap<String, Vec<u8>>,
    selected: bool,
) -> Result<bool, Box<dyn std::error::Error>> {
    let actual_members = members
        .keys()
        .filter(|name| {
            name.starts_with(RELATIONSHIP_MEMBER_PREFIX) && name.ends_with(".rels")
        })
        .cloned()
        .collect::<BTreeSet<_>>();
    if actual_members != expected_story_relationship_members() {
        return Ok(false);
    }
    for (story_index, story) in STORIES.iter().enumerate() {
        let name = story_rels_name(story.file);
        let xml = members
            .get(&name)
            .ok_or_else(|| format!("DOCX output is missing relationship member {name}"))?;
        let actual = parse_relationship_oracle(xml)?;
        if actual != expected_story_relationships(*story, story_index, selected) {
            return Ok(false);
        }
        if !members.contains_key(&media_name(story_index)) {
            return Ok(false);
        }
        if story_index == 0 {
            for owner_story in STORIES.iter().skip(1) {
                if !members.contains_key(&story_part_name(owner_story.file)) {
                    return Ok(false);
                }
            }
        }
    }
    Ok(true)
}

pub(super) fn build_corpus() -> Result<Corpus, Box<dyn std::error::Error>> {
    let archive = build_archive()?;
    let source_members = archive_members(&archive)?;
    if source_members.is_empty() {
        return Err("DOCX story publication corpus has no ZIP members".into());
    }
    let package = OpcPackage::from_bytes(&archive)?;
    let uncompressed_payload_bytes = package.iter_parts().try_fold(0usize, |total, part| {
        total
            .checked_add(part.blob().len())
            .ok_or("DOCX story publication payload count overflow")
    })?;
    let target_payload = SHARED_TARGET.as_bytes();
    let main_bytes = source_members
        .get("word/document.xml")
        .ok_or("DOCX story publication corpus is missing its main story")?;
    let archive_sha256 = super::sha256_hex(&archive);
    if package.part_count() != EXPECTED_ENTRY_COUNT
        || source_members.len() != EXPECTED_ARCHIVE_MEMBER_COUNT
        || archive.len() != EXPECTED_ARCHIVE_BYTES
        || archive_sha256 != EXPECTED_ARCHIVE_SHA256
    {
        return Err("DOCX story publication corpus identity changed unexpectedly".into());
    }
    let manifest = CorpusManifest {
        name: "docx-story-hyperlink-publication".to_owned(),
        generator: CORPUS_GENERATOR,
        package_format: "DOCX/OPC/ZIP",
        shape: CORPUS_SHAPE,
        payload_kind: PAYLOAD_KIND,
        compression: "deflate",
        entry_count: EXPECTED_ENTRY_COUNT,
        archive_member_count: EXPECTED_ARCHIVE_MEMBER_COUNT,
        entry_bytes: main_bytes.len(),
        uncompressed_payload_bytes,
        archive_bytes: EXPECTED_ARCHIVE_BYTES,
        archive_sha256: EXPECTED_ARCHIVE_SHA256.to_owned(),
        target_entry: "shared-external-hyperlink-relationship".to_owned(),
        target_payload_bytes: target_payload.len(),
        target_payload_sha256: super::sha256_hex(target_payload),
        rtf_variant: None,
        xlsx: None,
    };
    Ok(Corpus {
        manifest,
        archive,
        source_members,
    })
}

fn source_package(
    bytes: Vec<u8>,
) -> Result<source_backed::Package, Box<dyn std::error::Error>> {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    Ok(source_backed::Package::from_read_at(source)?)
}

fn publish(
    corpus: &Corpus,
    target_urls: &[&str],
    sink: &mut impl Write,
) -> Result<(), Box<dyn std::error::Error>> {
    let package = source_package(corpus.archive.clone())?;
    let plan = package.plan_story_hyperlink_redaction(target_urls, Mode::Strict)?;
    let commit = plan.apply()?;
    package.publish_story_hyperlink_redaction_to_stream(sink, &commit)?;
    Ok(())
}

fn publish_to_vec(
    corpus: &Corpus,
    target_urls: &[&str],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut output = Vec::new();
    output.try_reserve_exact(corpus.archive.len())?;
    publish(corpus, target_urls, &mut output)?;
    Ok(output)
}

fn story_oracle(xml: &[u8], story: StorySpec) -> Result<StoryOracle, Box<dyn std::error::Error>> {
    let mut reader = NsReader::from_reader(xml);
    let mut oracle = StoryOracle::default();
    loop {
        let (namespace, event) = reader.read_resolved_event()?;
        let is_word = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == W.as_bytes());
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_word && element.local_name().as_ref() == b"hyperlink" =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    if attribute.key.as_ref() != b"r:id" {
                        continue;
                    }
                    let value = attribute.decoded_and_normalized_value(
                        quick_xml::XmlVersion::Explicit1_0,
                        reader.decoder(),
                    )?;
                    match value.as_ref() {
                        "rShared" => oracle.shared_links += 1,
                        "rOther" => oracle.unselected_links += 1,
                        _ => return Err("DOCX XML oracle saw an unexpected hyperlink ID".into()),
                    }
                }
            },
            Event::Start(element) | Event::Empty(element)
                if is_word && element.local_name().as_ref() == b"drawing" =>
            {
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    if attribute.key.as_ref() == b"r:embed"
                        && attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                reader.decoder(),
                            )?
                            .as_ref()
                            == "rMedia"
                    {
                        oracle.media_refs += 1;
                    }
                }
            },
            Event::Start(element) | Event::Empty(element)
                if is_word && element.local_name().as_ref() == b"t" =>
            {
                oracle.text_nodes += 1;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if oracle.shared_links > 1 || oracle.unselected_links != 1 || oracle.media_refs != 1 {
        return Err(format!(
            "DOCX XML oracle found unexpected {} story counts: {oracle:?}",
            story.kind
        )
        .into());
    }
    Ok(oracle)
}

fn verify_story_xml(
    members: &BTreeMap<String, Vec<u8>>,
    selected: bool,
) -> Result<bool, Box<dyn std::error::Error>> {
    for story in STORIES {
        let name = story_part_name(story.file);
        let xml = members
            .get(&name)
            .ok_or_else(|| format!("DOCX output is missing story member {name}"))?;
        let oracle = story_oracle(xml, story)?;
        let expected_shared = usize::from(!selected);
        if oracle.shared_links != expected_shared
            || oracle.unselected_links != 1
            || oracle.media_refs != 1
            || oracle.text_nodes != 2
        {
            return Ok(false);
        }
        let root = format!("<w:{}", story.root);
        if !String::from_utf8_lossy(xml).contains(&root) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn expected_changed_members() -> BTreeSet<String> {
    STORIES
        .iter()
        .flat_map(|story| {
            [
                story_part_name(story.file),
                story_rels_name(story.file),
            ]
        })
        .collect()
}

fn verify_noop_output(corpus: &Corpus, output: &[u8]) -> Result<bool, Box<dyn std::error::Error>> {
    if output != corpus.archive {
        return Ok(false);
    }
    let members = archive_members(output)?;
    Ok(members == corpus.source_members && verify_story_xml(&members, false)?)
}

fn verify_redaction_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<bool, Box<dyn std::error::Error>> {
    let output_members = archive_members(output)?;
    if output_members.keys().collect::<BTreeSet<_>>()
        != corpus.source_members.keys().collect::<BTreeSet<_>>()
    {
        return Ok(false);
    }
    if !verify_story_xml(&output_members, true)? {
        return Ok(false);
    }
    let source_raw = super::raw_zip_members(&corpus.archive)?;
    let output_raw = super::raw_zip_members(output)?;
    if source_raw.keys().collect::<BTreeSet<_>>()
        != output_raw.keys().collect::<BTreeSet<_>>()
    {
        return Ok(false);
    }
    let expected_changed = expected_changed_members();
    let changed = corpus
        .source_members
        .iter()
        .filter_map(|(name, source)| {
            output_members
                .get(name)
                .filter(|candidate| *candidate != source)
                .map(|_| name.clone())
        })
        .collect::<BTreeSet<_>>();
    if changed != expected_changed {
        return Ok(false);
    }
    for (name, source) in &corpus.source_members {
        if !expected_changed.contains(name) && output_members.get(name) != Some(source) {
            return Ok(false);
        }
    }
    for (name, source) in &source_raw {
        if !expected_changed.contains(name) && output_raw.get(name) != Some(source) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn signed_archive(_corpus: &Corpus) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml")?,
        STORIES[0].content_type.to_owned(),
        main_xml()?,
    );
    add_story_relationships(&mut main, 0)?;
    for (story_index, story_spec) in STORIES.iter().enumerate().skip(1) {
        main.rels_mut().try_add_relationship(
            story_spec.relationship_type.to_owned(),
            story_spec.file.to_owned(),
            format!("rStory{story_index:02}"),
            TargetMode::Internal,
        )?;
        let mut story = BlobPart::new(
            PackURI::new(format!("/{}", story_part_name(story_spec.file)))?,
            story_spec.content_type.to_owned(),
            story_xml(*story_spec, story_index)?,
        );
        add_story_relationships(&mut story, story_index)?;
        package.try_add_part(Box::new(story))?;
    }
    for story_index in 0..STORIES.len() {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{}", media_name(story_index)))?,
            ct::PNG.to_owned(),
            media_payload(story_index),
        )))?;
    }
    package.try_add_part(Box::new(main))?;
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/_xmlsignatures/origin.sigs")?,
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        b"<origin/>".to_vec(),
    )))?;
    package.rels_mut().try_add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        "_xmlsignatures/origin.sigs".to_owned(),
        "rSignature".to_owned(),
        TargetMode::Internal,
    )?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn unknown_owner_archive(corpus: &Corpus) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut package = OpcPackage::from_bytes(&corpus.archive)?;
    let mut custom = BlobPart::new(
        PackURI::new("/word/opaque/custom-owner.xml")?,
        "application/xml".to_owned(),
        b"<custom/>".to_vec(),
    );
    custom.rels_mut().try_add_relationship(
        rt::HEADER.to_owned(),
        "../header1.xml".to_owned(),
        "rLateHeader".to_owned(),
        TargetMode::Internal,
    )?;
    package.try_add_part(Box::new(custom))?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn verify_refusal_gates(corpus: &Corpus) -> Result<(bool, bool, bool, bool, bool, bool), Box<dyn std::error::Error>> {
    let revision = Arc::new(AtomicU64::new(0));
    let versioned: Arc<dyn ReadAt> = Arc::new(VersionedSource {
        bytes: corpus.archive.clone(),
        revision: Arc::clone(&revision),
    });
    let stale_package = source_backed::Package::from_read_at(Arc::clone(&versioned))?;
    let stale_commit = stale_package
        .plan_story_hyperlink_redaction(&[SHARED_TARGET], Mode::Strict)?
        .apply()?;
    revision.fetch_add(1, Ordering::SeqCst);
    let mut stale_output = Vec::new();
    let stale_refusal = stale_package
        .publish_story_hyperlink_redaction_to_stream(&mut stale_output, &stale_commit)
        .is_err()
        && stale_output.is_empty();

    let foreign_first = source_package(corpus.archive.clone())?;
    let foreign_second = source_package(corpus.archive.clone())?;
    let foreign_commit = foreign_first
        .plan_story_hyperlink_redaction(&[SHARED_TARGET], Mode::Strict)?
        .apply()?;
    let mut foreign_output = Vec::new();
    let foreign_refusal = foreign_second
        .publish_story_hyperlink_redaction_to_stream(&mut foreign_output, &foreign_commit)
        .is_err()
        && foreign_output.is_empty();

    let signed_refusal = match signed_archive(corpus) {
        Ok(signed) => match source_package(signed) {
            Ok(signed_package) => signed_package
                .plan_story_hyperlink_redaction(&[SHARED_TARGET], Mode::Strict)
                .is_err(),
            Err(_) => false,
        },
        Err(_) => false,
    };

    let unknown_refusal = match unknown_owner_archive(corpus) {
        Ok(unknown) => match source_package(unknown) {
            Ok(unknown_package) => unknown_package
                .plan_story_hyperlink_redaction(&[SHARED_TARGET], Mode::Strict)
                .is_err(),
            Err(_) => false,
        },
        Err(_) => false,
    };

    let package = source_package(corpus.archive.clone())?;
    let commit = package
        .plan_story_hyperlink_redaction(&[SHARED_TARGET], Mode::Strict)?
        .apply()?;
    let mut partial = FailingSink {
        accepted: 0,
        fail_after: 1,
    };
    let partial_refusal = package
        .publish_story_hyperlink_redaction_to_stream(&mut partial, &commit)
        .is_err()
        && partial.accepted > 0;
    let package = source_package(corpus.archive.clone())?;
    let commit = package
        .plan_story_hyperlink_redaction(&[SHARED_TARGET], Mode::Strict)?
        .apply()?;
    let mut zero = FailingSink {
        accepted: 0,
        fail_after: 0,
    };
    let zero_refusal = package
        .publish_story_hyperlink_redaction_to_stream(&mut zero, &commit)
        .is_err()
        && zero.accepted == 0;

    Ok((
        stale_refusal,
        foreign_refusal,
        signed_refusal,
        unknown_refusal,
        partial_refusal,
        zero_refusal,
    ))
}

fn prepared_output(
    corpus: &Corpus,
    case: Case,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let target_urls = match case {
        Case::DocxStoryHyperlinkNoopSave => &[][..],
        Case::DocxStoryHyperlinkRedactionSave => &[SHARED_TARGET][..],
        _ => return Err("DOCX story publication received an unrelated case".into()),
    };
    publish_to_vec(corpus, target_urls)
}

pub(super) fn run(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn std::error::Error>> {
    if !matches!(
        case,
        Case::DocxStoryHyperlinkNoopSave | Case::DocxStoryHyperlinkRedactionSave
    ) {
        return Err("DOCX story publication runner received an unrelated case".into());
    }
    let noop_output = publish_to_vec(corpus, &[])?;
    let redaction_output = publish_to_vec(corpus, &[SHARED_TARGET])?;
    let expected_output = match case {
        Case::DocxStoryHyperlinkNoopSave => noop_output.clone(),
        Case::DocxStoryHyperlinkRedactionSave => redaction_output.clone(),
        _ => unreachable!("case validated above"),
    };
    let source_hash_verified = super::sha256_hex(&corpus.archive) == EXPECTED_ARCHIVE_SHA256
        && corpus.manifest.archive_sha256 == EXPECTED_ARCHIVE_SHA256;
    let source_zip_oracle_verified = corpus.archive.len() == EXPECTED_ARCHIVE_BYTES
        && corpus.manifest.entry_count == EXPECTED_ENTRY_COUNT
        && corpus.manifest.archive_member_count == EXPECTED_ARCHIVE_MEMBER_COUNT
        && corpus.source_members.len() == EXPECTED_ARCHIVE_MEMBER_COUNT
        && corpus.source_members.contains_key(OPAQUE_MEMBER)
        && (0..STORIES.len()).all(|index| corpus.source_members.contains_key(&media_name(index)));
    let no_op_exact_bytes_verified = verify_noop_output(corpus, &noop_output)?;
    let changed_member_locality_verified = verify_redaction_output(corpus, &redaction_output)?;
    let noop_members = archive_members(&noop_output)?;
    let redaction_members = archive_members(&redaction_output)?;
    let relationship_oracle_verified = verify_relationship_oracle(&noop_members, false)?
        && verify_relationship_oracle(&redaction_members, true)?;
    let xml_semantic_oracle_verified = verify_story_xml(&noop_members, false)?
        && verify_story_xml(&redaction_members, true)?;
    let repeated_output = prepared_output(corpus, case)?;
    let deterministic_output_verified = repeated_output == expected_output;
    let (
        stale_source_refusal_verified,
        foreign_source_refusal_verified,
        signed_refusal_verified,
        unknown_owner_refusal_verified,
        partial_sink_refusal_verified,
        zero_sink_refusal_verified,
    ) = verify_refusal_gates(corpus)?;
    if !source_hash_verified
        || !source_zip_oracle_verified
        || !no_op_exact_bytes_verified
        || !changed_member_locality_verified
        || !relationship_oracle_verified
        || !xml_semantic_oracle_verified
        || !deterministic_output_verified
        || !stale_source_refusal_verified
        || !foreign_source_refusal_verified
        || !signed_refusal_verified
        || !unknown_owner_refusal_verified
        || !partial_sink_refusal_verified
        || !zero_sink_refusal_verified
    {
        return Err(format!(
            "DOCX story publication correctness gates failed: source_hash={source_hash_verified}, source_zip={source_zip_oracle_verified}, noop={no_op_exact_bytes_verified}, locality={changed_member_locality_verified}, rels={relationship_oracle_verified}, xml={xml_semantic_oracle_verified}, deterministic={deterministic_output_verified}, stale={stale_source_refusal_verified}, foreign={foreign_source_refusal_verified}, signed={signed_refusal_verified}, unknown={unknown_owner_refusal_verified}, partial={partial_sink_refusal_verified}, zero={zero_sink_refusal_verified}"
        )
        .into());
    }

    let maximum = u64::try_from(expected_output.len())?
        .checked_add(64 * 1024)
        .ok_or("DOCX story publication sink budget overflow")?;
    let mut elapsed = Vec::new();
    elapsed.try_reserve_exact(samples)?;
    let mut sink_summaries: Vec<SinkSummary> = Vec::new();
    sink_summaries.try_reserve_exact(samples)?;
    let mut source_immutability_verified = true;
    for iteration in 0..super::iteration_count(warmup_iterations, samples)? {
        let source_bytes = corpus.archive.clone();
        let measured_source = Arc::new(OwnedSource::new(source_bytes));
        let source_identity = SourceIdentity::capture(&measured_source)?;
        let source: Arc<dyn ReadAt> = measured_source.clone();
        let mut sink = CountingSink::bounded(maximum, u64::MAX);
        sink.reserve_budget()?;
        let target_urls = match case {
            Case::DocxStoryHyperlinkNoopSave => &[][..],
            Case::DocxStoryHyperlinkRedactionSave => &[SHARED_TARGET][..],
            _ => unreachable!("case validated above"),
        };
        let started = std::time::Instant::now();
        let package = source_backed::Package::from_read_at(source)?;
        let plan = package.plan_story_hyperlink_redaction(target_urls, Mode::Strict)?;
        let commit = plan.apply()?;
        package.publish_story_hyperlink_redaction_to_stream(&mut sink, &commit)?;
        let duration = started.elapsed();
        source_immutability_verified &= source_identity.matches(&measured_source)?;
        if sink.bytes != expected_output {
            return Err("DOCX story publication output differs from its preflight output".into());
        }
        std::hint::black_box(&sink.bytes);
        if iteration >= warmup_iterations {
            elapsed.push(super::elapsed_ns(duration)?);
            sink_summaries.push(sink.summary());
        }
    }
    if !source_immutability_verified {
        return Err("DOCX story publication source identity changed during publication".into());
    }
    let sink = super::deterministic_sink_summary(&sink_summaries, "DOCX story publication")?;
    let output_hash = super::sha256_hex(&expected_output);
    let summary = DocxStoryHyperlinkPublicationSummary {
        implementation: "litchi-docx::source_backed::Package + story_hyperlinks::Plan",
        timing_scope: "fresh source and reserved sequential sink prepared outside; open + strict target plan + commit + sequential publication inside",
        performance_claim: "none: correctness-only end-to-end publication evidence",
        predeclared_allocator_model: (case == Case::DocxStoryHyperlinkRedactionSave)
            .then_some(STORY_RELATIONSHIP_ALLOCATOR_MODEL),
        story_kinds: STORIES.iter().map(|story| story.kind.to_owned()).collect(),
        selected_target: SHARED_TARGET.to_owned(),
        selected_relationship_count: STORIES.len(),
        unselected_relationship_count: STORIES.len(),
        source_archive_bytes: u64::try_from(corpus.archive.len())?,
        source_archive_sha256: corpus.manifest.archive_sha256.clone(),
        output_archive_bytes: u64::try_from(expected_output.len())?,
        output_archive_sha256: output_hash.clone(),
        end_to_end_ns: elapsed.clone(),
        source_zip_oracle_verified,
        source_hash_verified,
        no_op_exact_bytes_verified,
        changed_member_locality_verified,
        relationship_oracle_verified,
        xml_semantic_oracle_verified,
        deterministic_output_verified,
        source_immutability_verified,
        stale_source_refusal_verified,
        foreign_source_refusal_verified,
        signed_refusal_verified,
        unknown_owner_refusal_verified,
        partial_sink_refusal_verified,
        zero_sink_refusal_verified,
    };
    let source = SourceSummary {
        docx_story_hyperlink_publication: Some(summary),
        ..SourceSummary::default()
    };
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: super::statistics(elapsed),
        sink: Some(sink),
        source: Some(Box::new(source)),
        execution: None,
        output_sha256: Some(output_hash),
        operation_metrics: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn corpus_is_deterministic_and_covers_every_story_kind() {
        let first = build_corpus().expect("first DOCX story publication corpus");
        let second = build_corpus().expect("second DOCX story publication corpus");
        assert_eq!(first.archive, second.archive);
        assert_eq!(first.source_members, second.source_members);
        assert_eq!(first.manifest.archive_sha256, second.manifest.archive_sha256);
        assert_eq!(STORIES.len(), 7);
        assert_eq!(first.manifest.archive_member_count, first.source_members.len());
        assert!(first.source_members.contains_key(OPAQUE_MEMBER));
        assert!(!Case::DEFAULT.contains(&Case::DocxStoryHyperlinkNoopSave));
        assert!(!Case::DEFAULT.contains(&Case::DocxStoryHyperlinkRedactionSave));
        assert_eq!(
            super::super::parse_case("docx_story_hyperlink_noop_save"),
            Some(Case::DocxStoryHyperlinkNoopSave)
        );
        assert_eq!(
            super::super::parse_case("docx_story_hyperlink_redaction_save"),
            Some(Case::DocxStoryHyperlinkRedactionSave)
        );
    }

    #[test]
    fn noop_and_redaction_publishers_emit_end_to_end_gated_evidence() {
        let corpus = build_corpus().expect("DOCX story publication corpus");
        for case in [
            Case::DocxStoryHyperlinkNoopSave,
            Case::DocxStoryHyperlinkRedactionSave,
        ] {
            let result = run(case, &corpus, 0, 1).expect("DOCX story publication run");
            assert_eq!(result.case, case.name());
            assert_eq!(result.elapsed_ns.samples.len(), 1);
            let summary = result
                .source
                .expect("DOCX story publication source evidence")
                .docx_story_hyperlink_publication
                .expect("DOCX story publication summary");
            assert_eq!(summary.story_kinds.len(), 7);
            assert_eq!(summary.selected_relationship_count, 7);
            assert_eq!(summary.unselected_relationship_count, 7);
            let expected_allocator_model =
                (case == Case::DocxStoryHyperlinkRedactionSave)
                    .then_some(STORY_RELATIONSHIP_ALLOCATOR_MODEL);
            assert_eq!(summary.predeclared_allocator_model, expected_allocator_model);
            assert!(summary.source_zip_oracle_verified);
            assert!(summary.source_hash_verified);
            assert!(summary.no_op_exact_bytes_verified);
            assert!(summary.changed_member_locality_verified);
            assert!(summary.relationship_oracle_verified);
            assert!(summary.xml_semantic_oracle_verified);
            assert!(summary.deterministic_output_verified);
            assert!(summary.source_immutability_verified);
            assert!(summary.stale_source_refusal_verified);
            assert!(summary.foreign_source_refusal_verified);
            assert!(summary.signed_refusal_verified);
            assert!(summary.unknown_owner_refusal_verified);
            assert!(summary.partial_sink_refusal_verified);
            assert!(summary.zero_sink_refusal_verified);
            assert_eq!(summary.performance_claim, "none: correctness-only end-to-end publication evidence");
        }
    }
}
