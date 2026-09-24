//! Shared bounded OPC exercise paths for libFuzzer and the deterministic smoke
//! binary. The parser receives only caller-owned byte slices; it never
//! resolves paths, opens resources, fetches relationship targets, or executes
//! embedded content.

use std::hint::black_box;
use std::io::Cursor;

use litchi_opc::{
    AuthoredXmlFragment, OpcError, OpcPackage, PackURI, ReadLimits, ReadResource,
    SourceBackedPackage, SourceTopologyPlan, probe_package_catalog_from_reader_with_limits,
};

/// Maximum fuzzer input retained by one exercise.
pub const MAX_INPUT_BYTES: usize = 1 << 20;
const MAX_ARCHIVE_MEMBERS: usize = 64;
const MAX_PARTS: usize = 64;
const MAX_RELATIONSHIP_PARTS: usize = 32;
const MAX_MEMBER_BYTES: u64 = 256 << 10;
const MAX_ARCHIVE_TOTAL_BYTES: u64 = 1 << 20;
const MAX_RELATIONSHIP_XML_BYTES: usize = 64 << 10;
const MAX_TOTAL_RELATIONSHIP_XML_BYTES: usize = 256 << 10;
const MAX_RELATIONSHIPS_PER_PART: usize = 64;
const MAX_TOTAL_RELATIONSHIPS: usize = 256;
const MAX_XML_EVENTS: usize = 512;
const MAX_TOTAL_XML_EVENTS: usize = 2048;
const MAX_XML_DEPTH: usize = 32;
const MAX_XML_ATTRIBUTE_BYTES: usize = 16 << 10;
const MAX_RELATIONSHIP_TARGET_BYTES: usize = 4 << 10;

/// Results of the three public ingress paths for one input.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ExerciseStats {
    /// Whether the source-backed constructor admitted the input.
    pub source_backed_ok: bool,
    /// Whether one source XML edit was accepted by the topology plan.
    pub source_xml_replacement_ok: bool,
    /// Whether the bounded source XML copy was accepted by the topology plan.
    pub source_xml_addition_ok: bool,
    /// Whether the resulting source topology was written successfully.
    pub source_topology_write_ok: bool,
    /// Whether all admitted internal source relationships resolved to PackURI values.
    pub source_relationship_resolution_ok: bool,
    /// Whether the admitted source package resolved its unique main document.
    pub source_main_document_ok: bool,
    /// Whether the written source topology reopened with its main document and appended edge.
    pub source_topology_reopen_ok: bool,
    /// Whether the eager constructor admitted the input.
    pub eager_ok: bool,
    /// Whether all admitted internal eager relationships resolved to PackURI values.
    pub eager_relationship_resolution_ok: bool,
    /// Whether the eager package resolved its unique main document.
    pub eager_main_document_ok: bool,
    /// Whether the metadata-only catalog probe admitted the input.
    pub probe_ok: bool,
}

/// Exact resource counts for the deterministic smoke package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BoundaryExpectations {
    /// Compressed package length.
    pub input_bytes: usize,
    /// Number of physical ZIP members.
    pub archive_members: usize,
    /// Number of admitted ordinary OPC parts.
    pub parts: usize,
    /// Number of relationship parts.
    pub relationship_parts: usize,
    /// `[Content_Types].xml` length.
    pub content_types_bytes: usize,
    /// Number of content-type defaults and overrides.
    pub content_type_mappings: usize,
    /// Largest relationship XML member.
    pub relationship_xml_bytes: usize,
    /// Sum of relationship XML member lengths.
    pub total_relationship_xml_bytes: usize,
    /// Largest relationship count in one part.
    pub relationships_per_part: usize,
    /// Total relationship count.
    pub total_relationships: usize,
    /// Maximum ZIP member-name length.
    pub archive_member_name_bytes: usize,
    /// Aggregate central-directory metadata bytes.
    pub archive_metadata_bytes: usize,
    /// Maximum declared compressed member bytes.
    pub archive_compressed_bytes: usize,
    /// Maximum declared uncompressed member bytes.
    pub archive_entry_bytes: usize,
    /// Aggregate declared uncompressed ZIP bytes.
    pub archive_total_bytes: usize,
    /// Maximum admitted ordinary Part bytes.
    pub part_bytes: usize,
    /// Aggregate admitted ordinary Part bytes.
    pub total_part_bytes: usize,
    /// Relationship graph nodes visited from package and Part relationships.
    pub relationship_graph_nodes: usize,
    /// Maximum XML event count in one content-types or relationships member.
    pub xml_events: usize,
    /// Aggregate XML events across relationships members.
    pub total_relationship_xml_events: usize,
    /// Maximum individual XML attribute value bytes.
    pub xml_attribute_bytes: usize,
    /// Maximum relationship Target attribute bytes.
    pub relationship_target_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
enum BoundaryKind {
    InputBytes,
    ArchiveMembers,
    Parts,
    RelationshipParts,
    ContentTypesBytes,
    ContentTypeMappings,
    RelationshipXmlBytes,
    TotalRelationshipXmlBytes,
    RelationshipsPerPart,
    TotalRelationships,
    ArchiveMemberNameBytes,
    ArchiveMetadataBytes,
    ArchiveCompressedBytes,
    ArchiveEntryBytes,
    ArchiveTotalBytes,
    PartBytes,
    TotalPartBytes,
    RelationshipGraphNodes,
    XmlEvents,
    TotalRelationshipXmlEvents,
    XmlAttributeBytes,
    RelationshipTargetBytes,
}

impl BoundaryKind {
    const fn label(self) -> &'static str {
        match self {
            Self::InputBytes => "input bytes",
            Self::ArchiveMembers => "archive members",
            Self::Parts => "parts",
            Self::RelationshipParts => "relationship parts",
            Self::ContentTypesBytes => "content types bytes",
            Self::ContentTypeMappings => "content type mappings",
            Self::RelationshipXmlBytes => "relationship XML bytes",
            Self::TotalRelationshipXmlBytes => "total relationship XML bytes",
            Self::RelationshipsPerPart => "relationships per part",
            Self::TotalRelationships => "total relationships",
            Self::ArchiveMemberNameBytes => "archive member-name bytes",
            Self::ArchiveMetadataBytes => "archive metadata bytes",
            Self::ArchiveCompressedBytes => "archive compressed bytes",
            Self::ArchiveEntryBytes => "archive entry bytes",
            Self::ArchiveTotalBytes => "archive total bytes",
            Self::PartBytes => "Part bytes",
            Self::TotalPartBytes => "total Part bytes",
            Self::RelationshipGraphNodes => "relationship graph nodes",
            Self::XmlEvents => "XML events",
            Self::TotalRelationshipXmlEvents => "total relationship XML events",
            Self::XmlAttributeBytes => "XML attribute bytes",
            Self::RelationshipTargetBytes => "relationship Target bytes",
        }
    }
}

/// The deliberately tight policy used by this harness.
#[must_use]
pub fn tight_limits() -> ReadLimits {
    limits_for_input(MAX_INPUT_BYTES as u64)
}

fn limits_for_input(max_input_bytes: u64) -> ReadLimits {
    ReadLimits::builder()
        .max_input_bytes(max_input_bytes.max(1))
        .expect("nonzero fuzzer input ceiling")
        .max_archive_members(MAX_ARCHIVE_MEMBERS)
        .expect("nonzero archive-member ceiling")
        .max_parts(MAX_PARTS)
        .expect("nonzero part ceiling")
        .max_relationship_parts(MAX_RELATIONSHIP_PARTS)
        .expect("nonzero relationship-part ceiling")
        .max_archive_member_name_bytes(4 << 10)
        .expect("nonzero member-name ceiling")
        .max_archive_metadata_bytes(64 << 10)
        .expect("nonzero archive-metadata ceiling")
        .max_archive_compressed_bytes(MAX_MEMBER_BYTES)
        .expect("nonzero compressed-byte ceiling")
        .max_archive_entry_bytes(MAX_MEMBER_BYTES)
        .expect("nonzero entry-byte ceiling")
        .max_archive_total_bytes(MAX_ARCHIVE_TOTAL_BYTES)
        .expect("nonzero archive-total ceiling")
        .max_part_bytes(MAX_MEMBER_BYTES)
        .expect("nonzero part-byte ceiling")
        .max_total_part_bytes(MAX_ARCHIVE_TOTAL_BYTES)
        .expect("nonzero total-part ceiling")
        .max_content_types_bytes(MAX_RELATIONSHIP_XML_BYTES)
        .expect("nonzero content-types ceiling")
        .max_content_type_mappings(MAX_RELATIONSHIPS_PER_PART)
        .expect("nonzero content-type mapping ceiling")
        .max_relationship_xml_bytes(MAX_RELATIONSHIP_XML_BYTES)
        .expect("nonzero relationship-XML ceiling")
        .max_total_relationship_xml_bytes(MAX_TOTAL_RELATIONSHIP_XML_BYTES)
        .expect("nonzero total relationship-XML ceiling")
        .max_relationships_per_part(MAX_RELATIONSHIPS_PER_PART)
        .expect("nonzero per-part relationship ceiling")
        .max_total_relationships(MAX_TOTAL_RELATIONSHIPS)
        .expect("nonzero total relationship ceiling")
        .max_relationship_graph_nodes(MAX_TOTAL_RELATIONSHIPS)
        .expect("nonzero graph-node ceiling")
        .max_xml_events(MAX_XML_EVENTS)
        .expect("nonzero XML-event ceiling")
        .max_total_relationship_xml_events(MAX_TOTAL_XML_EVENTS)
        .expect("nonzero total XML-event ceiling")
        .max_xml_depth(MAX_XML_DEPTH)
        .expect("nonzero XML-depth ceiling")
        .max_xml_attribute_bytes(MAX_XML_ATTRIBUTE_BYTES)
        .expect("nonzero XML-attribute ceiling")
        .max_relationship_target_bytes(MAX_RELATIONSHIP_TARGET_BYTES)
        .expect("nonzero relationship-target ceiling")
        .build()
        .expect("consistent bounded OPC profile")
}

fn source_xml_insertion(bytes: &[u8]) -> Option<usize> {
    bytes.windows(2).rposition(|window| window == b"</")
}

/// Exercise all bounded public ingress paths and safe read/probe operations.
///
/// This function intentionally ignores ordinary parse errors: malformed input
/// is a first-class fuzz case. Internal consistency checks remain assertions,
/// so a panic is a harness failure rather than an accepted parse result.
#[must_use]
pub fn exercise(data: &[u8]) -> ExerciseStats {
    if data.len() > MAX_INPUT_BYTES {
        return ExerciseStats::default();
    }

    let limits = tight_limits();
    let probe_ok = probe_package(data, limits);
    let source = exercise_source_backed(data, limits);
    let eager = exercise_eager(data, limits);
    ExerciseStats {
        source_backed_ok: source.admitted,
        source_xml_replacement_ok: source.xml_replacement_ok,
        source_xml_addition_ok: source.xml_addition_ok,
        source_topology_write_ok: source.topology_write_ok,
        source_relationship_resolution_ok: source.relationship_resolution_ok,
        source_main_document_ok: source.main_document_ok,
        source_topology_reopen_ok: source.topology_reopen_ok,
        eager_ok: eager.admitted,
        eager_relationship_resolution_ok: eager.relationship_resolution_ok,
        eager_main_document_ok: eager.main_document_ok,
        probe_ok,
    }
}

fn probe_package(data: &[u8], limits: ReadLimits) -> bool {
    let mut reader = Cursor::new(data);
    probe_package_catalog_from_reader_with_limits(&mut reader, limits).is_ok()
}

#[derive(Debug, Clone, Copy, Default)]
struct SourceBackedStats {
    admitted: bool,
    xml_replacement_ok: bool,
    xml_addition_ok: bool,
    topology_write_ok: bool,
    relationship_resolution_ok: bool,
    main_document_ok: bool,
    topology_reopen_ok: bool,
}

#[derive(Debug, Clone, Copy, Default)]
struct EagerStats {
    admitted: bool,
    relationship_resolution_ok: bool,
    main_document_ok: bool,
}

fn source_relationships_resolve(package: &SourceBackedPackage) -> bool {
    package.rels().iter().all(|relationship| {
        relationship.is_external()
            || relationship
                .target_partname()
                .is_ok_and(|target| package.part(&target).is_ok())
    }) && package.iter_parts().all(|part| {
        part.rels().iter().all(|relationship| {
            relationship.is_external()
                || relationship
                    .target_partname()
                    .is_ok_and(|target| package.part(&target).is_ok())
        })
    })
}

fn source_main_document_resolves(package: &SourceBackedPackage) -> bool {
    package
        .main_document_part()
        .is_ok_and(|part| part.partname().as_str() == "/word/document.xml")
}

fn eager_relationships_resolve(package: &OpcPackage) -> bool {
    package.rels().iter().all(|relationship| {
        relationship.is_external()
            || relationship
                .target_partname()
                .is_ok_and(|target| package.contains_part(&target))
    }) && package.iter_parts().all(|part| {
        part.rels().iter().all(|relationship| {
            relationship.is_external()
                || relationship
                    .target_partname()
                    .is_ok_and(|target| package.contains_part(&target))
        })
    })
}

fn eager_main_document_resolves(package: &OpcPackage) -> bool {
    package
        .main_document_part()
        .is_ok_and(|part| part.partname().as_str() == "/word/document.xml")
}

fn topology_reopens(data: &[u8], limits: ReadLimits) -> bool {
    let Ok(source) = SourceBackedPackage::from_vec_with_limits(data.to_vec(), limits) else {
        return false;
    };
    if !source_relationships_resolve(&source)
        || !source_main_document_resolves(&source)
        || !source
            .rels()
            .get("rIdLitchiFuzzAppend")
            .is_some_and(|relationship| relationship.is_external())
    {
        return false;
    }
    let Ok(eager) = OpcPackage::from_bytes_with_limits(data, limits) else {
        return false;
    };
    eager_relationships_resolve(&eager)
        && eager_main_document_resolves(&eager)
        && eager
            .rels()
            .get("rIdLitchiFuzzAppend")
            .is_some_and(|relationship| relationship.is_external())
}

fn exercise_source_backed(data: &[u8], limits: ReadLimits) -> SourceBackedStats {
    let Ok(package) = SourceBackedPackage::from_vec_with_limits(data.to_vec(), limits) else {
        return SourceBackedStats::default();
    };

    let mut stats = SourceBackedStats {
        admitted: true,
        relationship_resolution_ok: source_relationships_resolve(&package),
        main_document_ok: source_main_document_resolves(&package),
        ..SourceBackedStats::default()
    };
    let mut plan = SourceTopologyPlan::new();
    for part in package.iter_parts() {
        black_box(part.partname());
        black_box(part.content_type());
        let _ = black_box(part.declared_uncompressed_size());

        // Capture the source-authorized precompressed member through cold and
        // warm paths, then verify that both retained tokens can be published.
        if let Ok((cold, token)) = part.data_and_authorize_precompressed() {
            if let Ok((warm, warm_token)) = part.data_and_authorize_precompressed() {
                assert_eq!(cold.as_bytes(), warm.as_bytes());
                drop(warm_token);
            }
            let retained = token.into_retained();
            let copy = retained.clone();
            let first = retained.authorize_for_publication();
            let second = copy.authorize_for_publication();
            drop(first);
            drop(second);
        }

        // Exercise source XML proof/publication only when this admitted part
        // exposes XML source. Binary parts remain opaque and bounded.
        if !stats.xml_replacement_ok
            && let Ok(source_xml) = part.source_xml()
            && let Some(insertion) = source_xml_insertion(source_xml.bytes())
            && let Ok(proof) = source_xml.checked_range(insertion..insertion, &[])
            && let Ok(fragment) = AuthoredXmlFragment::markup(b"<litchi-fuzz/>".to_vec())
            && let Ok(mut publication) = source_xml.into_publication()
            && publication.replace(proof, fragment).is_ok()
            && let Ok(edited) = publication.finish()
        {
            let target = part.partname().clone();
            if plan
                .try_replace_source_xml_part(target, edited.clone())
                .is_ok()
            {
                stats.xml_replacement_ok = true;
                if let Ok(copy_name) = PackURI::new("/litchi-fuzz-source-copy.xml") {
                    stats.xml_addition_ok = plan.try_add_source_xml_part(copy_name, edited).is_ok();
                }
            }
        }
    }

    // This relationship is deliberately external and inert. No target is
    // resolved or fetched by the harness or by the package writer.
    if let Ok(owner) = PackURI::new("/")
        && plan
            .try_add_external_relationship(
                owner,
                "rIdLitchiFuzzAppend",
                "urn:litchi:fuzz:external",
                "https://example.invalid/fuzz",
            )
            .is_ok()
    {
        let mut output = Vec::new();
        if package.write_topology_to_stream(&mut output, plan).is_ok() {
            stats.topology_write_ok = true;
            stats.topology_reopen_ok = topology_reopens(&output, limits);
        }
    }
    stats
}

fn exercise_eager(data: &[u8], limits: ReadLimits) -> EagerStats {
    let Ok(package) = OpcPackage::from_bytes_with_limits(data, limits) else {
        return EagerStats::default();
    };

    for rel in package.rels().iter() {
        black_box(rel.r_id());
        black_box(rel.reltype());
        black_box(rel.target_ref());
        black_box(rel.is_external());
        if !rel.is_external() {
            let _ = black_box(rel.target_partname());
        }
    }
    for part in package.iter_parts() {
        black_box(part.partname());
        black_box(part.content_type());
        black_box(part.blob().len());
        for rel in part.rels().iter() {
            black_box(rel.r_id());
            black_box(rel.reltype());
            black_box(rel.target_ref());
            black_box(rel.is_external());
            if !rel.is_external() {
                let _ = black_box(rel.target_partname());
            }
        }
    }
    let relationship_resolution_ok = eager_relationships_resolve(&package);
    let main_document_ok = eager_main_document_resolves(&package);
    black_box(package.part_count());
    EagerStats {
        admitted: true,
        relationship_resolution_ok,
        main_document_ok,
    }
}

fn boundary_limits(kind: BoundaryKind, value: usize) -> ReadLimits {
    let builder = ReadLimits::builder();
    let builder = match kind {
        BoundaryKind::InputBytes => builder
            .max_input_bytes(u64::try_from(value).expect("bounded boundary"))
            .expect("nonzero input boundary"),
        BoundaryKind::ArchiveMembers => builder
            .max_archive_members(value)
            .expect("nonzero member boundary")
            .max_parts(value)
            .expect("nonzero dependent part boundary")
            .max_relationship_parts(value)
            .expect("nonzero dependent relationship boundary"),
        BoundaryKind::Parts => builder.max_parts(value).expect("nonzero part boundary"),
        BoundaryKind::RelationshipParts => builder
            .max_relationship_parts(value)
            .expect("nonzero relationship boundary"),
        BoundaryKind::ContentTypesBytes => builder
            .max_content_types_bytes(value)
            .expect("nonzero content-types boundary"),
        BoundaryKind::ContentTypeMappings => builder
            .max_content_type_mappings(value)
            .expect("nonzero mapping boundary"),
        BoundaryKind::RelationshipXmlBytes => builder
            .max_relationship_xml_bytes(value)
            .expect("nonzero relationship XML boundary"),
        BoundaryKind::TotalRelationshipXmlBytes => builder
            .max_total_relationship_xml_bytes(value)
            .expect("nonzero total relationship XML boundary"),
        BoundaryKind::RelationshipsPerPart => builder
            .max_relationships_per_part(value)
            .expect("nonzero per-part relationship boundary"),
        BoundaryKind::TotalRelationships => builder
            .max_total_relationships(value)
            .expect("nonzero total relationship boundary"),
        BoundaryKind::ArchiveMemberNameBytes => builder
            .max_archive_member_name_bytes(value as u64)
            .expect("nonzero member-name boundary"),
        BoundaryKind::ArchiveMetadataBytes => builder
            .max_archive_metadata_bytes(value as u64)
            .expect("nonzero metadata boundary"),
        BoundaryKind::ArchiveCompressedBytes => builder
            .max_archive_compressed_bytes(value as u64)
            .expect("nonzero compressed-byte boundary"),
        BoundaryKind::ArchiveEntryBytes => builder
            .max_archive_entry_bytes(value as u64)
            .expect("nonzero entry-byte boundary"),
        BoundaryKind::ArchiveTotalBytes => builder
            .max_archive_total_bytes(value as u64)
            .expect("nonzero archive-total boundary"),
        BoundaryKind::PartBytes => builder
            .max_part_bytes(value as u64)
            .expect("nonzero Part-byte boundary"),
        BoundaryKind::TotalPartBytes => builder
            .max_total_part_bytes(value as u64)
            .expect("nonzero total-Part-byte boundary"),
        BoundaryKind::RelationshipGraphNodes => builder
            .max_relationship_graph_nodes(value)
            .expect("nonzero graph-node boundary"),
        BoundaryKind::XmlEvents => builder
            .max_xml_events(value)
            .expect("nonzero XML-event boundary"),
        BoundaryKind::TotalRelationshipXmlEvents => builder
            .max_total_relationship_xml_events(value)
            .expect("nonzero total XML-event boundary"),
        BoundaryKind::XmlAttributeBytes => builder
            .max_xml_attribute_bytes(value)
            .expect("nonzero XML-attribute boundary")
            .max_relationship_target_bytes(value)
            .expect("nonzero dependent target boundary"),
        BoundaryKind::RelationshipTargetBytes => builder
            .max_relationship_target_bytes(value)
            .expect("nonzero target boundary"),
    };
    builder.build().expect("consistent boundary profile")
}

fn ingress_stats(data: &[u8], limits: ReadLimits) -> ExerciseStats {
    let mut reader = Cursor::new(data);
    ExerciseStats {
        source_backed_ok: SourceBackedPackage::from_vec_with_limits(data.to_vec(), limits).is_ok(),
        eager_ok: OpcPackage::from_bytes_with_limits(data, limits).is_ok(),
        probe_ok: probe_package_catalog_from_reader_with_limits(&mut reader, limits).is_ok(),
        ..ExerciseStats::default()
    }
}

fn assert_boundary_rejected(
    data: &[u8],
    kind: BoundaryKind,
    value: usize,
    expected_resource: ReadResource,
) {
    let limits = boundary_limits(kind, value);
    let source = SourceBackedPackage::from_vec_with_limits(data.to_vec(), limits);
    assert!(
        matches!(
            source,
            Err(OpcError::ReadLimit { resource, .. }) if resource == expected_resource
        ),
        "source-backed path must reject the {} boundary at {value}",
        kind.label()
    );
    let eager = OpcPackage::from_bytes_with_limits(data, limits);
    assert!(
        matches!(
            eager,
            Err(OpcError::ReadLimit { resource, .. }) if resource == expected_resource
        ),
        "eager path must reject the {} boundary at {value}",
        kind.label()
    );
    let mut reader = Cursor::new(data);
    let probe = probe_package_catalog_from_reader_with_limits(&mut reader, limits);
    assert!(
        matches!(
            probe,
            Err(OpcError::ReadLimit { resource, .. }) if resource == expected_resource
        ),
        "probe path must reject the {} boundary at {value}",
        kind.label()
    );
}

/// Check exact, one-over, and one-under values for deterministic resource
/// boundaries through all three public ingress paths.
pub fn exercise_boundaries(data: &[u8], expected: BoundaryExpectations) {
    let cases = [
        (
            BoundaryKind::InputBytes,
            expected.input_bytes,
            ReadResource::InputBytes,
        ),
        (
            BoundaryKind::ArchiveMembers,
            expected.archive_members,
            ReadResource::ArchiveMembers,
        ),
        (BoundaryKind::Parts, expected.parts, ReadResource::Parts),
        (
            BoundaryKind::RelationshipParts,
            expected.relationship_parts,
            ReadResource::RelationshipParts,
        ),
        (
            BoundaryKind::ContentTypesBytes,
            expected.content_types_bytes,
            ReadResource::ContentTypesBytes,
        ),
        (
            BoundaryKind::ContentTypeMappings,
            expected.content_type_mappings,
            ReadResource::ContentTypeMappings,
        ),
        (
            BoundaryKind::RelationshipXmlBytes,
            expected.relationship_xml_bytes,
            ReadResource::RelationshipXmlBytes,
        ),
        (
            BoundaryKind::TotalRelationshipXmlBytes,
            expected.total_relationship_xml_bytes,
            ReadResource::TotalRelationshipXmlBytes,
        ),
        (
            BoundaryKind::RelationshipsPerPart,
            expected.relationships_per_part,
            ReadResource::RelationshipsPerPart,
        ),
        (
            BoundaryKind::TotalRelationships,
            expected.total_relationships,
            ReadResource::TotalRelationships,
        ),
        (
            BoundaryKind::ArchiveMemberNameBytes,
            expected.archive_member_name_bytes,
            ReadResource::ArchiveMemberNameBytes,
        ),
        (
            BoundaryKind::ArchiveMetadataBytes,
            expected.archive_metadata_bytes,
            ReadResource::ArchiveMetadataBytes,
        ),
        (
            BoundaryKind::ArchiveCompressedBytes,
            expected.archive_compressed_bytes,
            ReadResource::ArchiveCompressedBytes,
        ),
        (
            BoundaryKind::ArchiveEntryBytes,
            expected.archive_entry_bytes,
            ReadResource::ArchiveEntryBytes,
        ),
        (
            BoundaryKind::ArchiveTotalBytes,
            expected.archive_total_bytes,
            ReadResource::ArchiveTotalBytes,
        ),
        (
            BoundaryKind::PartBytes,
            expected.part_bytes,
            ReadResource::PartBytes,
        ),
        (
            BoundaryKind::TotalPartBytes,
            expected.total_part_bytes,
            ReadResource::TotalPartBytes,
        ),
        (
            BoundaryKind::RelationshipGraphNodes,
            expected.relationship_graph_nodes,
            ReadResource::RelationshipGraphNodes,
        ),
        (
            BoundaryKind::XmlEvents,
            expected.xml_events,
            ReadResource::XmlEvents,
        ),
        (
            BoundaryKind::TotalRelationshipXmlEvents,
            expected.total_relationship_xml_events,
            ReadResource::TotalRelationshipXmlEvents,
        ),
        (
            BoundaryKind::XmlAttributeBytes,
            expected.xml_attribute_bytes,
            ReadResource::XmlAttributeBytes,
        ),
        (
            BoundaryKind::RelationshipTargetBytes,
            expected.relationship_target_bytes,
            ReadResource::RelationshipTargetBytes,
        ),
    ];
    for (kind, exact, resource) in cases {
        assert!(
            exact > 1,
            "boundary fixture must leave an under-limit value"
        );
        for value in [exact, exact + 1] {
            let stats = ingress_stats(data, boundary_limits(kind, value));
            assert!(
                stats.source_backed_ok && stats.eager_ok && stats.probe_ok,
                "all ingress paths must admit exact/over {} boundary at {value}: {stats:?}",
                kind.label()
            );
        }
        assert_boundary_rejected(data, kind, exact - 1, resource);
    }
}
