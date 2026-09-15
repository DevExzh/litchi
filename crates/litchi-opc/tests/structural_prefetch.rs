#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "these integration tests fail explicitly when their small OPC fixture is invalid"
)]

//! Public contract coverage for the run-coalesced structural prefetch.
//!
//! Every test here drives the source-backed package through its ordinary
//! public constructors and observes only what a caller-supplied positional
//! source can see: how many reads it was asked for, at which offsets, and what
//! the open then reported. The prefetch is a cache of source bytes, so the
//! contract it must keep is that nothing except the read count moves.

use std::num::{NonZeroU64, NonZeroUsize};
use std::ops::Range;
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};
use std::{io, iter};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, SourceVersion,
};
use litchi_opc::{
    OpcError, PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy,
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CHILD_REL: &str = "urn:litchi:test/child";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CHILD_CONTENT_TYPE: &str = "application/xml";

/// A positional source that records every read it is asked for.
struct Recording {
    bytes: Vec<u8>,
    reads: Mutex<Vec<(u64, usize)>>,
    versions: AtomicU64,
    /// A read beginning at this offset and at least this long fails, so a
    /// coalescing failure can be observed without disturbing any other read.
    refuse: Option<(u64, usize)>,
}

impl Recording {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes,
            reads: Mutex::new(Vec::new()),
            versions: AtomicU64::new(0),
            refuse: None,
        })
    }

    fn refusing(bytes: Vec<u8>, offset: u64, at_least: usize) -> Arc<Self> {
        Arc::new(Self {
            bytes,
            reads: Mutex::new(Vec::new()),
            versions: AtomicU64::new(0),
            refuse: Some((offset, at_least)),
        })
    }

    fn take(&self) -> Vec<(u64, usize)> {
        std::mem::take(&mut self.reads.lock().unwrap())
    }

    fn versions(&self) -> u64 {
        self.versions.load(Ordering::SeqCst)
    }
}

impl ReadAt for Recording {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        if self
            .refuse
            .is_some_and(|(at, least)| offset == at && output.len() >= least)
        {
            return Err(io::Error::other("recording source refuses a long read"));
        }
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let taken = if start >= self.bytes.len() {
            0
        } else {
            output.len().min(self.bytes.len() - start)
        };
        if taken > 0 {
            output[..taken].copy_from_slice(&self.bytes[start..start + taken]);
        }
        self.reads.lock().unwrap().push((offset, taken));
        Ok(taken)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::SeqCst);
        Ok(SourceVersion::new(0x0623, 1))
    }
}

fn content_types(child_parts: &[String]) -> String {
    let overrides: String = child_parts
        .iter()
        .map(|name| format!(r#"<Override PartName="/{name}" ContentType="{CHILD_CONTENT_TYPE}"/>"#))
        .collect();
    format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="{DOCUMENT_CONTENT_TYPE}"/>{overrides}</Types>"#
    )
}

fn package_rels() -> String {
    format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    )
}

fn rels_naming(targets: &[String]) -> String {
    let entries: String = targets
        .iter()
        .enumerate()
        .map(|(index, target)| {
            format!(
                r#"<Relationship Id="rId{}" Type="{CHILD_REL}" Target="/{target}"/>"#,
                index + 1
            )
        })
        .collect();
    format!(r#"<Relationships xmlns="{RELATIONSHIPS_NS}">{entries}</Relationships>"#)
}

const MALFORMED_RELS: &str = concat!(
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Ty"#
);

fn duplicate_id_rels() -> String {
    format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rDup" Type="urn:litchi:test/a" Target="a.xml"/><Relationship Id="rDup" Type="urn:litchi:test/b" Target="b.xml"/></Relationships>"#
    )
}

/// A package whose `child_count` leaf parts each carry their own relationship
/// part, all written consecutively so the structural members form one run.
///
/// `broken` replaces the relationship part at that index with `replacement`,
/// which is how a malformed member is placed at the start, the middle or the
/// end of the run.
fn run_package(child_count: usize, broken: Option<(usize, String)>) -> Vec<u8> {
    let children: Vec<String> = (0..child_count)
        .map(|index| format!("word/child{index}.xml"))
        .collect();
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types(&children).as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", package_rels().as_bytes())
        .unwrap();
    writer
        .write_stored(
            "word/_rels/document.xml.rels",
            rels_naming(&children).as_bytes(),
        )
        .unwrap();
    for (index, _) in children.iter().enumerate() {
        let body = match &broken {
            Some((position, replacement)) if *position == index => replacement.clone(),
            _ => rels_naming(&[]),
        };
        writer
            .write_stored(
                &format!("word/_rels/child{index}.xml.rels"),
                body.as_bytes(),
            )
            .unwrap();
    }
    writer
        .write_stored("word/document.xml", b"<document/>")
        .unwrap();
    for child in &children {
        writer.write_stored(child, b"<child/>").unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn open(source: Arc<Recording>) -> Result<SourceBackedPackage, OpcError> {
    let dynamic: Arc<dyn ReadAt> = source;
    SourceBackedPackage::from_read_at(dynamic)
}

fn open_with_policy(
    source: Arc<Recording>,
    policy: SourceReadPolicy,
) -> Result<SourceBackedPackage, OpcError> {
    let dynamic: Arc<dyn ReadAt> = source;
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
        dynamic,
        ReadLimits::default(),
        SourceCacheLimits::default(),
        policy,
    )
}

fn local_header_offset(bytes: &[u8], name: &str) -> u64 {
    let needle = name.as_bytes();
    let mut offset = 0usize;
    while offset + 30 <= bytes.len() {
        if bytes[offset..offset + 4] == [0x50, 0x4b, 0x03, 0x04] {
            let name_len = u16::from_le_bytes([bytes[offset + 26], bytes[offset + 27]]) as usize;
            if bytes.get(offset + 30..offset + 30 + name_len) == Some(needle) {
                return offset as u64;
            }
        }
        offset += 1;
    }
    panic!("member {name} has no local header");
}

fn payload_range(bytes: &[u8], name: &str) -> Range<u64> {
    let header = local_header_offset(bytes, name) as usize;
    let compressed = u32::from_le_bytes([
        bytes[header + 18],
        bytes[header + 19],
        bytes[header + 20],
        bytes[header + 21],
    ]) as u64;
    let name_len = u16::from_le_bytes([bytes[header + 26], bytes[header + 27]]) as u64;
    let extra_len = u16::from_le_bytes([bytes[header + 28], bytes[header + 29]]) as u64;
    let start = header as u64 + 30 + name_len + extra_len;
    start..start + compressed
}

/// Everything the open observed, in one comparable value.
#[derive(Debug, PartialEq, Eq)]
struct Verdict {
    outcome: Result<Vec<String>, String>,
    non_parts: Vec<String>,
}

fn verdict(source: Arc<Recording>) -> Verdict {
    verdict_with_limits(source, ReadLimits::default())
}

fn verdict_with_limits(source: Arc<Recording>, limits: ReadLimits) -> Verdict {
    let dynamic: Arc<dyn ReadAt> = source;
    match SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
        dynamic,
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
    ) {
        Err(error) => Verdict {
            outcome: Err(format!("{error:?}")),
            non_parts: Vec::new(),
        },
        Ok(package) => {
            let mut parts: Vec<String> = package
                .iter_parts()
                .map(|part| {
                    format!(
                        "{} {} rels={}",
                        part.partname(),
                        part.content_type(),
                        part.rels().len()
                    )
                })
                .collect();
            parts.sort();
            let mut non_parts: Vec<String> = package
                .non_part_members()
                .iter()
                .map(|member| format!("{} {}", member.name(), member.reason().as_str()))
                .collect();
            non_parts.sort();
            Verdict {
                outcome: Ok(parts),
                non_parts,
            }
        },
    }
}

#[test]
fn one_read_serves_a_contiguous_run_of_structural_members() {
    let bytes = run_package(8, None);
    let source = Recording::new(bytes.clone());
    let package = open(Arc::clone(&source)).expect("open");
    let reads = source.take();

    // Eleven structural members — the content-types member, the package
    // `_rels/.rels`, the document's relationship part and eight child
    // relationship parts — are written consecutively, so one read covers all
    // of them. The remaining reads are the archive locator's.
    assert_eq!(package.iter_parts().count(), 9);
    assert_eq!(
        reads.len(),
        4,
        "expected three locator reads and one run read, got {reads:?}"
    );
    let run = reads
        .iter()
        .max_by_key(|(_, taken)| *taken)
        .copied()
        .expect("a run read");
    assert_eq!(
        run.0,
        local_header_offset(&bytes, "[Content_Types].xml"),
        "the run read begins at the first structural member's local header"
    );
    let document = payload_range(&bytes, "word/document.xml");
    assert!(
        run.0 + run.1 as u64 <= document.start,
        "the run read stops before the first ordinary payload"
    );
}

#[test]
fn a_run_read_never_covers_an_ordinary_part_payload() {
    // The ordinary parts sit after the structural members, and one more
    // ordinary member is written between two relationship parts so a run has
    // to break rather than read across it.
    let mut writer = StreamingArchiveWriter::new();
    let children = vec!["word/child0.xml".to_string(), "word/child1.xml".to_string()];
    writer
        .write_stored("[Content_Types].xml", content_types(&children).as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", package_rels().as_bytes())
        .unwrap();
    writer
        .write_stored(
            "word/_rels/document.xml.rels",
            rels_naming(&children).as_bytes(),
        )
        .unwrap();
    writer
        .write_stored("word/_rels/child0.xml.rels", rels_naming(&[]).as_bytes())
        .unwrap();
    let wedge: Vec<u8> = iter::repeat_n(b'W', 4096).collect();
    writer.write_stored("word/child0.xml", &wedge).unwrap();
    writer
        .write_stored("word/_rels/child1.xml.rels", rels_naming(&[]).as_bytes())
        .unwrap();
    writer.write_stored("word/child1.xml", b"<child/>").unwrap();
    writer
        .write_stored("word/document.xml", b"<document/>")
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let source = Recording::new(bytes.clone());
    open(Arc::clone(&source)).expect("open");
    let reads = source.take();
    let wedge_payload = payload_range(&bytes, "word/child0.xml");
    for (offset, taken) in &reads {
        let end = offset + *taken as u64;
        assert!(
            end <= wedge_payload.start || *offset >= wedge_payload.end,
            "read {offset}+{taken} overlaps the wedged ordinary payload {wedge_payload:?}"
        );
    }
}

#[test]
fn a_malformed_relationship_part_keeps_its_error_anywhere_in_a_run() {
    let mut observed = Vec::new();
    for position in [0usize, 3, 7] {
        let bytes = run_package(8, Some((position, MALFORMED_RELS.to_string())));
        let source = Recording::new(bytes);
        let error = open(Arc::clone(&source))
            .map(drop)
            .expect_err("a malformed relationship part refuses");
        let reads = source.take();
        assert!(
            reads.iter().any(|(_, taken)| *taken > 1024),
            "position {position} must refuse with the run read taken, not through the fallback: {reads:?}"
        );
        observed.push(format!("{error:?}"));

        // The same package, with the coalesced fetch refused, must reach the
        // same refusal through the per-member grammar.
        let bytes = run_package(8, Some((position, MALFORMED_RELS.to_string())));
        let coalesced = verdict(Recording::new(bytes.clone()));
        let fallback = verdict(Recording::refusing(bytes, 0, 1024));
        assert_eq!(coalesced, fallback, "position {position}");
    }
    assert_eq!(
        observed[0], observed[1],
        "the start and the middle of a run report the same error"
    );
    assert_eq!(
        observed[1], observed[2],
        "the middle and the end of a run report the same error"
    );
    assert!(
        observed[0].contains("Xml") || observed[0].contains("syntax"),
        "unexpected error identity {}",
        observed[0]
    );
}

#[test]
fn a_duplicate_relationship_id_keeps_its_error_anywhere_in_a_run() {
    let mut observed = Vec::new();
    for position in [0usize, 3, 7] {
        let bytes = run_package(8, Some((position, duplicate_id_rels())));
        let source = Recording::new(bytes);
        let error = open(Arc::clone(&source))
            .map(drop)
            .expect_err("a duplicate relationship ID refuses");
        observed.push(format!("{error:?}"));

        let bytes = run_package(8, Some((position, duplicate_id_rels())));
        let coalesced = verdict(Recording::new(bytes.clone()));
        let fallback = verdict(Recording::refusing(bytes, 0, 1024));
        assert_eq!(coalesced, fallback, "position {position}");
    }
    assert_eq!(observed[0], observed[1]);
    assert_eq!(observed[1], observed[2]);
    assert!(
        observed[0].contains("rDup"),
        "unexpected error identity {}",
        observed[0]
    );
}

/// Change 0577's `e4`/`e5` pair: the same untyped member is tolerated archive
/// junk when nothing names it and a fatal `ContentTypeNotFound` when a deep
/// relationship does. The verdict is a property of the package's bytes, and
/// coalescing the fetch must not move it.
#[test]
fn the_untyped_member_admission_verdict_does_not_move() {
    let build = |referenced: bool| {
        let children = vec!["word/child0.xml".to_string()];
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("[Content_Types].xml", content_types(&children).as_bytes())
            .unwrap();
        writer
            .write_stored("_rels/.rels", package_rels().as_bytes())
            .unwrap();
        writer
            .write_stored(
                "word/_rels/document.xml.rels",
                rels_naming(&children).as_bytes(),
            )
            .unwrap();
        let child_rels = if referenced {
            rels_naming(&["junk/thing.bin".to_string()])
        } else {
            rels_naming(&[])
        };
        writer
            .write_stored("word/_rels/child0.xml.rels", child_rels.as_bytes())
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<document/>")
            .unwrap();
        writer.write_stored("word/child0.xml", b"<child/>").unwrap();
        writer.write_stored("junk/thing.bin", b"junk").unwrap();
        writer.finish_to_bytes().unwrap()
    };

    let unreferenced = verdict(Recording::new(build(false)));
    assert!(
        unreferenced.outcome.is_ok(),
        "an untyped, unreferenced member is tolerated junk: {unreferenced:?}"
    );
    assert_eq!(
        unreferenced.non_parts,
        vec![
            "junk/thing.bin ZIP item has no content type and no relationship refers to it"
                .to_string()
        ]
    );

    let referenced = verdict(Recording::new(build(true)));
    let error = referenced
        .outcome
        .expect_err("a referenced untyped member refuses");
    assert!(
        error.contains("ContentTypeNotFound") && error.contains("junk/thing.bin"),
        "unexpected error identity {error}"
    );
}

#[test]
fn a_refused_run_read_falls_back_to_per_member_reads() {
    let bytes = run_package(8, None);
    let baseline = Recording::new(bytes.clone());
    let coalesced = verdict(Arc::clone(&baseline));
    let coalesced_reads = baseline.take().len();

    // The run begins at the first member's local header and is far longer than
    // that member's own local record, so refusing a long read from that offset
    // refuses exactly the coalesced fetch and nothing else.
    let refusing = Recording::refusing(bytes, 0, 1024);
    let fallback = verdict(Arc::clone(&refusing));
    let fallback_reads = refusing.take().len();
    assert_eq!(
        coalesced, fallback,
        "a package whose run read fails must open exactly as it does without one"
    );
    assert!(
        fallback_reads > coalesced_reads,
        "the fallback must reach the source once per member: {fallback_reads} against {coalesced_reads}"
    );
}

#[test]
fn a_scattered_package_costs_no_more_reads() {
    // Every relationship part is separated from the next by an ordinary
    // member, so no run of two can form and the prefetch must fetch nothing.
    let children: Vec<String> = (0..4)
        .map(|index| format!("word/child{index}.xml"))
        .collect();
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types(&children).as_bytes())
        .unwrap();
    writer.write_stored("pad/a.bin", b"aaaaaaaa").unwrap();
    writer
        .write_stored("_rels/.rels", package_rels().as_bytes())
        .unwrap();
    writer.write_stored("pad/b.bin", b"bbbbbbbb").unwrap();
    writer
        .write_stored(
            "word/_rels/document.xml.rels",
            rels_naming(&children).as_bytes(),
        )
        .unwrap();
    for (index, child) in children.iter().enumerate() {
        writer.write_stored(child, b"<child/>").unwrap();
        writer
            .write_stored(
                &format!("word/_rels/child{index}.xml.rels"),
                rels_naming(&[]).as_bytes(),
            )
            .unwrap();
    }
    writer
        .write_stored("word/document.xml", b"<document/>")
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let source = Recording::new(bytes);
    open(Arc::clone(&source)).expect("open");
    let reads = source.take();
    // Three locator reads plus one read for each of the seven structural
    // members: exactly what the same package costs without this mechanism.
    assert_eq!(reads.len(), 10, "unexpected read shape {reads:?}");
}

#[test]
fn the_prefetch_is_released_when_the_catalog_is_built() {
    let bytes = run_package(4, None);
    let source = Recording::new(bytes);
    let package = open(Arc::clone(&source)).expect("open");
    let open_reads = source.take();
    assert!(!open_reads.is_empty());

    let uri = PackURI::new("/word/document.xml").expect("uri");
    let data = package.part(&uri).expect("part").data().expect("data");
    assert_eq!(data.as_bytes(), b"<document/>");
    let part_reads = source.take();
    assert!(
        !part_reads.is_empty(),
        "an ordinary part read after the open must still reach the source"
    );
}

#[test]
fn a_managed_open_keeps_the_exact_grammar() {
    let bytes = run_package(8, None);
    let unmanaged = Recording::new(bytes.clone());
    open(Arc::clone(&unmanaged)).expect("unmanaged open");
    let unmanaged_reads = unmanaged.take();

    let managed_source = Recording::new(bytes);
    let dynamic: Arc<dyn ReadAt> = managed_source.clone();
    let budget = Budget::root(
        "structural-prefetch-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker limit is nonzero"),
        NonZeroUsize::new(1).expect("one operation limit is nonzero"),
        NonZeroU64::new(u64::MAX).expect("memory limit is nonzero"),
        0,
    )
    .expect("execution limits must be valid");
    let context = ExecutionContext::new(budget, cancellation, execution_limits);
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
        dynamic,
        ReadLimits::default(),
        SourceCacheLimits::default(),
        context,
    )
    .expect("managed open");
    let managed_reads = managed_source.take();

    assert!(
        managed_reads.len() > unmanaged_reads.len(),
        "a managed open keeps one read per structural member: {} against {}",
        managed_reads.len(),
        unmanaged_reads.len()
    );
    assert_eq!(managed_reads.len(), 14, "unexpected managed read shape");
}

#[test]
fn a_forward_window_open_keeps_its_own_counters() {
    let bytes = run_package(8, None);
    let source = Recording::new(bytes);
    let policy = SourceReadPolicy::forward_start(4096).expect("policy");
    let package = open_with_policy(Arc::clone(&source), policy).expect("open");
    let diagnostics = package
        .source_read_diagnostics()
        .expect("diagnostics")
        .expect("a configured window reports diagnostics");
    assert!(diagnostics.enabled);
    assert!(
        diagnostics.requests > 0,
        "the forward window must still see every read it sees today"
    );
}

#[test]
fn source_observations_do_not_move() {
    let bytes = run_package(8, None);
    let source = Recording::new(bytes);
    open(Arc::clone(&source)).expect("open");
    assert_eq!(
        source.versions(),
        4,
        "an ordinary open observes the source version exactly as it does today"
    );
}

/// A relationship budget that the package exceeds part-way through a coalesced
/// run must refuse with the same typed error, and at the same member, whether
/// the bytes came from one run read or from eight member reads.
#[test]
fn a_relationship_budget_refuses_identically_inside_a_run() {
    let bytes = run_package(8, None);
    for ceiling in [1usize, 400, 900, 1_400, 2_000] {
        let limits = ReadLimits::builder()
            .max_total_relationship_xml_bytes(ceiling)
            .expect("ceiling")
            .build()
            .expect("limits");
        let coalesced = verdict_with_limits(Recording::new(bytes.clone()), limits);
        let fallback = verdict_with_limits(Recording::refusing(bytes.clone(), 0, 1024), limits);
        assert_eq!(
            coalesced, fallback,
            "a {ceiling}-byte relationship budget must refuse identically with and without the run read"
        );
        if let Err(error) = &coalesced.outcome {
            assert!(
                error.contains("TotalRelationshipXmlBytes"),
                "unexpected refusal identity {error}"
            );
        }
    }
}

/// Every run this package admits must be reached through the prefetch, not
/// through the fallback, or the tests above would be proving the fallback's
/// behaviour rather than the coalesced one's.
#[test]
fn the_fixtures_above_really_do_coalesce() {
    let bytes = run_package(8, None);
    let source = Recording::new(bytes);
    open(Arc::clone(&source)).expect("open");
    let reads = source.take();
    assert!(
        reads
            .iter()
            .any(|(offset, taken)| *offset == 0 && *taken > 1024),
        "no run read was issued: {reads:?}"
    );
}
