//! Source and package fixtures for the bounded `stylesWithEffects` owner.
//!
//! The fixture helpers deliberately operate on package bytes. They are test
//! setup only; production publication is exercised through the public DOCX
//! package API below the fixture matrix.

use std::collections::BTreeMap;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::result::Result;

use litchi_docx::styles::effects::{Owner, Resource, Snapshot};
use litchi_docx::{Package, ReadLimits};
use sha2::{Digest, Sha256};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const EFFECTS_CONTENT_TYPE: &[u8] = b"application/vnd.ms-word.stylesWithEffects+xml";
const EFFECTS_RELATIONSHIP: &[u8] =
    b"http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";

#[derive(Debug)]
struct FixtureError(String);

type FixtureResult<T> = Result<T, FixtureError>;

impl std::fmt::Display for FixtureError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl std::error::Error for FixtureError {}

fn fixture_path(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn fixture_bytes(relative: &str) -> FixtureResult<Vec<u8>> {
    let path = fixture_path(relative);
    fs::read(&path).map_err(|error| FixtureError(format!("read {}: {error}", path.display())))
}

fn archive_members(bytes: &[u8]) -> FixtureResult<BTreeMap<String, Vec<u8>>> {
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| FixtureError(format!("open fixture archive: {error:?}")))?;
    archive
        .file_names()
        .map(|name| {
            let content = archive
                .read(name)
                .map_err(|error| FixtureError(format!("read archive member {name}: {error:?}")))?;
            Ok((name.to_owned(), content))
        })
        .collect()
}

fn archive_member(bytes: &[u8], name: &str) -> FixtureResult<Vec<u8>> {
    archive_members(bytes)?
        .remove(name)
        .ok_or_else(|| FixtureError(format!("fixture has no member {name}")))
}

fn rewrite_archive<F>(bytes: &[u8], mut rewrite: F) -> FixtureResult<Vec<u8>>
where
    F: FnMut(&str, Vec<u8>) -> FixtureResult<Option<Vec<u8>>>,
{
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| FixtureError(format!("open fixture archive: {error:?}")))?;
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let content = archive
            .read(name)
            .map_err(|error| FixtureError(format!("read archive member {name}: {error:?}")))?;
        if let Some(content) = rewrite(name, content)? {
            writer
                .write_stored(name, &content)
                .map_err(|error| FixtureError(format!("write archive member {name}: {error:?}")))?;
        }
    }
    writer
        .finish_to_bytes()
        .map_err(|error| FixtureError(format!("finish fixture archive: {error:?}")))
}

fn replace_member(bytes: &[u8], selected: &str, replacement: Vec<u8>) -> FixtureResult<Vec<u8>> {
    rewrite_archive(bytes, |name, content| {
        Ok(Some(if name == selected {
            replacement.clone()
        } else {
            content
        }))
    })
}

fn rename_member(bytes: &[u8], selected: &str, replacement: &str) -> FixtureResult<Vec<u8>> {
    let archive = ArchiveReader::new(bytes)
        .map_err(|error| FixtureError(format!("open fixture archive: {error:?}")))?;
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let content = archive
            .read(name)
            .map_err(|error| FixtureError(format!("read archive member {name}: {error:?}")))?;
        let output_name = if name == selected { replacement } else { name };
        writer
            .write_stored(output_name, &content)
            .map_err(|error| {
                FixtureError(format!("write archive member {output_name}: {error:?}"))
            })?;
    }
    writer
        .finish_to_bytes()
        .map_err(|error| FixtureError(format!("finish fixture archive: {error:?}")))
}

fn remove_member(bytes: &[u8], selected: &str) -> FixtureResult<Vec<u8>> {
    rewrite_archive(bytes, |name, content| {
        Ok((name != selected).then_some(content))
    })
}

fn add_member(bytes: &[u8], added: &str, content: Vec<u8>) -> FixtureResult<Vec<u8>> {
    let members = archive_members(bytes)?;
    if members.contains_key(added) {
        return Err(FixtureError(format!("fixture already has member {added}")));
    }
    let mut writer = StreamingArchiveWriter::new();
    for (name, member) in members {
        writer
            .write_stored(&name, &member)
            .map_err(|error| FixtureError(format!("write archive member {name}: {error:?}")))?;
    }
    writer
        .write_stored(added, &content)
        .map_err(|error| FixtureError(format!("write archive member {added}: {error:?}")))?;
    writer
        .finish_to_bytes()
        .map_err(|error| FixtureError(format!("finish fixture archive: {error:?}")))
}

fn sha256(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().fold(String::new(), |mut output, byte| {
        output.push_str(&format!("{byte:02x}"));
        output
    })
}

fn replace_once(source: &[u8], marker: &[u8], replacement: &[u8]) -> FixtureResult<Vec<u8>> {
    let offset = source
        .windows(marker.len())
        .position(|window| window == marker)
        .ok_or_else(|| FixtureError(format!("fixture marker {:?} is absent", marker)))?;
    let mut output = Vec::with_capacity(
        source
            .len()
            .saturating_sub(marker.len())
            .saturating_add(replacement.len()),
    );
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[offset + marker.len()..]);
    Ok(output)
}

fn replace_all(source: &[u8], marker: &[u8], replacement: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(source.len());
    let mut cursor = 0;
    while let Some(relative) = source[cursor..]
        .windows(marker.len())
        .position(|window| window == marker)
    {
        let offset = cursor + relative;
        output.extend_from_slice(&source[cursor..offset]);
        output.extend_from_slice(replacement);
        cursor = offset + marker.len();
    }
    output.extend_from_slice(&source[cursor..]);
    output
}

fn package_bytes(package: &mut Package) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    package
        .to_stream(&mut output)
        .expect("serialize DOCX package in test");
    output.into_inner()
}

fn fixture_resource(relative: &str, member: &str) -> FixtureResult<Resource> {
    let bytes = fixture_bytes(relative)?;
    let xml = archive_member(&bytes, member)?;
    Resource::from_xml(xml)
        .map_err(|error| FixtureError(format!("parse fixture resource {member}: {error}")))
}

fn changed_resource(snapshot: &Snapshot) -> Resource {
    let mut xml = snapshot
        .resource()
        .expect("fixture snapshot has effects resource")
        .xml_bytes()
        .to_vec();
    xml = replace_once(
        &xml,
        b"</w:styles>",
        br#"<w:extensibilityMarker data="changed"/></w:styles>"#,
    )
    .expect("fixture effects root");
    let resource = Resource::from_xml(xml).expect("changed effects resource is valid");
    assert_eq!(resource.conformance(), snapshot.conformance());
    resource
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct PackageMetrics {
    parts: usize,
    total_part_bytes: u64,
    total_relationships: usize,
    relationship_parts: usize,
    relationship_graph_nodes: usize,
    relationship_xml_events: usize,
    relationship_xml_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
enum MutationCap {
    Parts,
    TotalPartBytes,
    TotalRelationships,
    RelationshipParts,
    RelationshipGraphNodes,
    RelationshipXmlEvents,
    RelationshipXmlBytes,
}

impl MutationCap {
    const ADDITION: [Self; 5] = [
        Self::Parts,
        Self::TotalPartBytes,
        Self::TotalRelationships,
        Self::RelationshipXmlEvents,
        Self::RelationshipXmlBytes,
    ];

    const TOPOLOGY: [Self; 2] = [Self::RelationshipParts, Self::RelationshipGraphNodes];

    const fn label(self) -> &'static str {
        match self {
            Self::Parts => "part count",
            Self::TotalPartBytes => "aggregate part bytes",
            Self::TotalRelationships => "aggregate relationships",
            Self::RelationshipParts => "relationship part count",
            Self::RelationshipGraphNodes => "relationship graph node count",
            Self::RelationshipXmlEvents => "aggregate relationship XML events",
            Self::RelationshipXmlBytes => "aggregate relationship XML bytes",
        }
    }

    fn set(self, value: u64) -> FixtureResult<ReadLimits> {
        let builder = ReadLimits::builder();
        let builder = match self {
            Self::Parts => builder.max_parts(
                usize::try_from(value)
                    .map_err(|error| FixtureError(format!("part cap conversion: {error}")))?,
            ),
            Self::TotalPartBytes => builder.max_total_part_bytes(value),
            Self::TotalRelationships => {
                builder.max_total_relationships(usize::try_from(value).map_err(|error| {
                    FixtureError(format!("relationship cap conversion: {error}"))
                })?)
            },
            Self::RelationshipParts => {
                builder.max_relationship_parts(usize::try_from(value).map_err(|error| {
                    FixtureError(format!("relationship-part cap conversion: {error}"))
                })?)
            },
            Self::RelationshipGraphNodes => builder
                .max_relationship_graph_nodes(usize::try_from(value).map_err(|error| {
                    FixtureError(format!("graph-node cap conversion: {error}"))
                })?),
            Self::RelationshipXmlEvents => builder.max_total_relationship_xml_events(
                usize::try_from(value)
                    .map_err(|error| FixtureError(format!("event cap conversion: {error}")))?,
            ),
            Self::RelationshipXmlBytes => builder.max_total_relationship_xml_bytes(
                usize::try_from(value)
                    .map_err(|error| FixtureError(format!("XML byte cap conversion: {error}")))?,
            ),
        }
        .map_err(|error| FixtureError(format!("set {} cap: {error:?}", self.label())))?;
        builder
            .build()
            .map_err(|error| FixtureError(format!("build {} cap: {error:?}", self.label())))
    }

    fn source_value(self, metrics: PackageMetrics) -> u64 {
        match self {
            Self::Parts => metrics.parts as u64,
            Self::TotalPartBytes => metrics.total_part_bytes,
            Self::TotalRelationships => metrics.total_relationships as u64,
            Self::RelationshipParts => metrics.relationship_parts as u64,
            Self::RelationshipGraphNodes => metrics.relationship_graph_nodes as u64,
            Self::RelationshipXmlEvents => metrics.relationship_xml_events as u64,
            Self::RelationshipXmlBytes => metrics.relationship_xml_bytes as u64,
        }
    }
}

fn is_relationship_member(name: &str) -> bool {
    name == "_rels/.rels" || (name.contains("/_rels/") && name.ends_with(".rels"))
}

fn relationship_xml_events(xml: &[u8], name: &str) -> FixtureResult<usize> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| FixtureError(format!("relationship event count overflow in {name}")))?;
        let event = reader
            .read_event()
            .map_err(|error| FixtureError(format!("parse relationship member {name}: {error}")))?;
        if matches!(event, quick_xml::events::Event::Eof) {
            return Ok(events);
        }
    }
}

fn package_metrics(bytes: &[u8]) -> FixtureResult<PackageMetrics> {
    let opc = litchi_opc::OpcPackage::from_vec(bytes.to_vec())
        .map_err(|error| FixtureError(format!("open OPC metrics source: {error:?}")))?;
    let total_part_bytes = opc.try_iter_parts().try_fold(0_u64, |total, part| {
        let part = part.map_err(|error| FixtureError(format!("decode part payload: {error:?}")))?;
        let bytes = u64::try_from(part.blob().len())
            .map_err(|error| FixtureError(format!("part byte conversion: {error}")))?;
        total
            .checked_add(bytes)
            .ok_or_else(|| FixtureError("part byte count overflow".into()))
    })?;
    let total_relationships = opc.rels().iter().count()
        + opc
            .iter_parts()
            .map(|part| part.rels().iter().count())
            .sum::<usize>();
    let relationship_parts = archive_members(bytes)?
        .keys()
        .filter(|name| is_relationship_member(name))
        .count();
    let relationship_graph_nodes = relationship_graph_nodes(&opc)?;
    let relationship_xml_bytes = archive_members(bytes)?
        .into_iter()
        .filter(|(name, _)| is_relationship_member(name))
        .map(|(_, xml)| xml.len())
        .sum();
    let relationship_xml_events = archive_members(bytes)?
        .into_iter()
        .filter(|(name, _)| is_relationship_member(name))
        .try_fold(0usize, |total, (name, xml)| {
            total
                .checked_add(relationship_xml_events(&xml, &name)?)
                .ok_or_else(|| FixtureError("relationship XML event count overflow".into()))
        })?;
    Ok(PackageMetrics {
        parts: opc.part_count(),
        total_part_bytes,
        total_relationships,
        relationship_parts,
        relationship_graph_nodes,
        relationship_xml_events,
        relationship_xml_bytes,
    })
}

fn relationship_graph_nodes(package: &litchi_opc::OpcPackage) -> FixtureResult<usize> {
    let mut visited = Vec::new();
    let mut work_queue = Vec::new();
    for relationship in package.rels().iter().filter(|value| !value.is_external()) {
        enqueue_graph_target(
            relationship.target_partname().map_err(|error| {
                FixtureError(format!("resolve package graph target: {error:?}"))
            })?,
            &mut visited,
            &mut work_queue,
        );
    }
    while let Some(owner) = work_queue.pop() {
        let Ok(part) = package.get_part(&owner) else {
            continue;
        };
        for relationship in part.rels().iter().filter(|value| !value.is_external()) {
            enqueue_graph_target(
                relationship.target_partname().map_err(|error| {
                    FixtureError(format!("resolve part graph target: {error:?}"))
                })?,
                &mut visited,
                &mut work_queue,
            );
        }
    }
    Ok(visited.len())
}

fn enqueue_graph_target(
    target: litchi_opc::PackURI,
    visited: &mut Vec<litchi_opc::PackURI>,
    work_queue: &mut Vec<litchi_opc::PackURI>,
) {
    if visited
        .iter()
        .any(|existing| existing.is_equivalent_to(&target))
    {
        return;
    }
    visited.push(target.clone());
    work_queue.push(target);
}

fn fixture_without_main_effects(relative: &str) -> FixtureResult<Vec<u8>> {
    let source = fixture_bytes(relative)?;
    let relationships_name = "word/_rels/document.xml.rels";
    let relationships = archive_member(&source, relationships_name)?;
    let relationships = replace_once(
        &relationships,
        br#"<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/>"#,
        b"",
    )?;
    let source = replace_member(&source, relationships_name, relationships)?;
    let content_types_name = "[Content_Types].xml";
    let content_types = archive_member(&source, content_types_name)?;
    let content_types = replace_once(
        &content_types,
        br#"<Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/>"#,
        b"",
    )?;
    let source = replace_member(&source, content_types_name, content_types)?;
    remove_member(&source, "word/stylesWithEffects.xml")
}

fn fixture_without_main_effects_or_owner_relationships(relative: &str) -> FixtureResult<Vec<u8>> {
    let source = fixture_without_main_effects(relative)?;
    remove_member(&source, "word/_rels/document.xml.rels")
}

fn assert_effects_add_rejected_under_cap(
    source: &[u8],
    resource: &Resource,
    limits: ReadLimits,
    label: &str,
    cap: MutationCap,
) -> FixtureResult<()> {
    let expected_resource = match cap {
        MutationCap::Parts => litchi_opc::ReadResource::Parts,
        MutationCap::TotalPartBytes => litchi_opc::ReadResource::TotalPartBytes,
        MutationCap::TotalRelationships => litchi_opc::ReadResource::TotalRelationships,
        MutationCap::RelationshipParts => litchi_opc::ReadResource::RelationshipParts,
        MutationCap::RelationshipGraphNodes => litchi_opc::ReadResource::RelationshipGraphNodes,
        MutationCap::RelationshipXmlEvents => litchi_opc::ReadResource::TotalRelationshipXmlEvents,
        MutationCap::RelationshipXmlBytes => litchi_opc::ReadResource::TotalRelationshipXmlBytes,
    };
    let mut package = Package::from_reader_with_limits(Cursor::new(source.to_vec()), limits)
        .map_err(|error| FixtureError(format!("open baseline under {label}: {error}")))?;
    let baseline = package_bytes(&mut package);
    assert_eq!(baseline, source, "baseline changed under {label} cap");
    assert!(
        package
            .styles_with_effects(Owner::MainDocument)
            .map_err(|error| FixtureError(format!("read absent owner under {label}: {error}")))?
            .is_empty()
    );

    assert!(
        package
            .put_styles_with_effects(Owner::MainDocument, resource.clone())
            .is_err(),
        "put unexpectedly passed the {label} cap"
    );
    assert_eq!(
        package_bytes(&mut package),
        baseline,
        "failed put changed bytes under {label} cap"
    );

    let snapshot = package
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| FixtureError(format!("reload baseline after {label} put: {error}")))?;
    assert!(snapshot.is_empty());
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(resource.clone()))
        .map_err(|error| FixtureError(format!("stage {label} patch: {error}")))?;
    assert!(
        matches!(
            edit.commit(),
            Err(litchi_docx::Error::Opc(litchi_opc::OpcError::ReadLimit {
                resource, actual, maximum,
            })) if resource == expected_resource && actual > maximum
        ),
        "commit must reject {label} before serializing projected metadata"
    );
    assert_eq!(package_bytes(&mut package), baseline);

    // A patch created under generous limits must also honor the destination
    // package's limits, independently of the snapshot commit-stage check.
    let generous = Package::from_reader(Cursor::new(source.to_vec()))
        .map_err(|error| FixtureError(format!("open generous {label} source: {error}")))?;
    let generous_snapshot = generous
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| FixtureError(format!("read generous {label} source: {error}")))?;
    let mut edit = generous_snapshot.edit();
    edit.replace_resource(Some(resource.clone()))
        .map_err(|error| FixtureError(format!("stage generous {label} patch: {error}")))?;
    let patch = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit generous {label} patch: {error}")))?
        .into_patch();
    assert!(
        package
            .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
            .is_err(),
        "patch unexpectedly passed the destination's {label} cap"
    );
    assert_eq!(
        package_bytes(&mut package),
        baseline,
        "failed patch changed bytes under {label} cap"
    );
    let reopened = Package::from_reader_with_limits(Cursor::new(baseline), limits)
        .map_err(|error| FixtureError(format!("reopen baseline under {label}: {error}")))?;
    assert!(
        reopened
            .styles_with_effects(Owner::MainDocument)
            .map_err(|error| FixtureError(format!(
                "read reopened baseline under {label}: {error}"
            )))?
            .is_empty()
    );
    Ok(())
}

fn assert_main_owner_rejected(bytes: Vec<u8>) {
    if let Ok(package) = Package::from_reader(Cursor::new(bytes)) {
        assert!(
            package.styles_with_effects(Owner::MainDocument).is_err(),
            "typed main effects owner accepted an invalid fixture"
        );
    }
}

fn assert_effects_member(
    relative: &str,
    package_sha256: &str,
    expected: &[(&str, usize, &str)],
) -> FixtureResult<()> {
    let bytes = fixture_bytes(relative)?;
    assert_eq!(
        sha256(&bytes),
        package_sha256,
        "package hash for {relative}"
    );
    let members = archive_members(&bytes)?;
    let content_types = members
        .get("[Content_Types].xml")
        .ok_or_else(|| FixtureError("fixture has no [Content_Types].xml".into()))?;
    assert!(
        content_types
            .windows(EFFECTS_CONTENT_TYPE.len())
            .any(|window| window == EFFECTS_CONTENT_TYPE),
        "missing effects content type in {relative}"
    );
    let relationship_members = members
        .iter()
        .filter(|(name, _)| name.ends_with("document.xml.rels"));
    let relationship_count = relationship_members
        .map(|(_, content)| {
            content
                .windows(EFFECTS_RELATIONSHIP.len())
                .filter(|window| *window == EFFECTS_RELATIONSHIP)
                .count()
        })
        .sum::<usize>();
    assert_eq!(
        relationship_count,
        expected.len(),
        "effects relationship count"
    );
    for (name, expected_len, expected_sha256) in expected {
        let content = members
            .get(*name)
            .ok_or_else(|| FixtureError(format!("fixture has no member {name}")))?;
        assert_eq!(content.len(), *expected_len, "member length for {name}");
        assert_eq!(sha256(content), *expected_sha256, "member hash for {name}");
        assert!(
            content
                .windows(b"<w:styles".len())
                .any(|window| window == b"<w:styles"),
            "effects member {name} has no w:styles root"
        );
    }
    Ok(())
}

#[test]
fn native_effects_fixture_hashes_and_relationships_match_design_evidence() -> FixtureResult<()> {
    assert_effects_member(
        "poi/test-data/document/Bug54849.docx",
        "f54182713ea5ce5d77b9593d3d9d24e645460043cec0b40ef59c932385f084d3",
        &[
            (
                "word/stylesWithEffects.xml",
                19_883,
                "799de1f7a4ce43f0ca101dc750e8a8bd6e75bcb721f7744d4787dad576cda3b1",
            ),
            (
                "word/glossary/stylesWithEffects.xml",
                16_138,
                "d27f6ced340ffa173b3b861b08a4006e76673a411687b4145f5e46dc6dae13eb",
            ),
        ],
    )?;
    assert_effects_member(
        "poi/test-data/xmldsign/ms-office-2010-signed.docx",
        "bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f",
        &[(
            "word/stylesWithEffects.xml",
            15_710,
            "00c5cda7671bf545a8c97312f14b2b8bc0ee7fa469b36c25ee158c8a5c1c1568",
        )],
    )?;
    assert_effects_member(
        "ooxml/docx/ComplexNumberedLists.docx",
        "297a085a7d433af2eeee7661e8db21539452cb585096484774a1e9f5f258b0b6",
        &[(
            "word/stylesWithEffects.xml",
            15_955,
            "b4bf5d355a45daf0a1085e73fe27041b5db22bfa23820f18ffaa7f9c8cb70f18",
        )],
    )?;
    assert_effects_member(
        "libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx",
        "8ccd581d8f0ae102b220228ad26b3974821a7ce8e3ff7df4b78f7da8a0d06ed9",
        &[
            (
                "word/stylesWithEffects.xml",
                20_117,
                "e72df38e71a351ebaaf7b102e8ab7862b7e5eb04cc86e187f8510746b70d3f54",
            ),
            (
                "word/glossary/stylesWithEffects.xml",
                16_244,
                "78112f02ff4e0b94a6c99f8688d86c6fa91d54d6c57c504001453b9c73e3dad1",
            ),
        ],
    )?;
    Ok(())
}

#[test]
fn native_main_and_glossary_effects_are_independent_resources() -> FixtureResult<()> {
    for relative in [
        "poi/test-data/document/Bug54849.docx",
        "libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx",
    ] {
        let bytes = fixture_bytes(relative)?;
        let main = archive_member(&bytes, "word/stylesWithEffects.xml")?;
        let glossary = archive_member(&bytes, "word/glossary/stylesWithEffects.xml")?;
        let ordinary_main = archive_member(&bytes, "word/styles.xml")?;
        let ordinary_glossary = archive_member(&bytes, "word/glossary/styles.xml")?;
        assert_ne!(main, glossary, "main/glossary effects were deduplicated");
        assert_ne!(main, ordinary_main, "effects/main styles were deduplicated");
        assert_ne!(
            glossary, ordinary_glossary,
            "effects/glossary styles were deduplicated"
        );
    }
    Ok(())
}

#[test]
fn native_fixture_matrix_loads_the_declared_main_and_glossary_owners() {
    let cases = [
        ("poi/test-data/document/Bug54849.docx", true, true),
        (
            "poi/test-data/xmldsign/ms-office-2010-signed.docx",
            true,
            false,
        ),
        ("ooxml/docx/ComplexNumberedLists.docx", true, false),
        (
            "libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx",
            true,
            true,
        ),
    ];
    for (relative, has_main, has_glossary) in cases {
        let bytes = fixture_bytes(relative).expect("read native effects fixture");
        let package =
            Package::from_reader(Cursor::new(bytes)).expect("open native effects fixture");
        let main = package
            .styles_with_effects(Owner::MainDocument)
            .expect("load main effects owner");
        assert_eq!(main.is_empty(), !has_main, "main owner in {relative}");
        let glossary = package
            .styles_with_effects(Owner::Glossary)
            .expect("load glossary effects owner");
        assert_eq!(
            glossary.is_empty(),
            !has_glossary,
            "glossary owner in {relative}"
        );
        assert_eq!(main.owner(), Owner::MainDocument);
        assert_eq!(glossary.owner(), Owner::Glossary);
        if let Some(resource) = main.resource() {
            assert_eq!(resource.styles().len(), resource.definitions().len());
        }
        if let Some(resource) = glossary.resource() {
            assert_eq!(resource.styles().len(), resource.projection().len());
        }
    }
}

#[test]
fn source_snapshots_retain_shared_payloads_and_exact_noop_commits() {
    let bytes = fixture_bytes("poi/test-data/document/Bug54849.docx").expect("read fixture");
    let package = Package::from_reader(Cursor::new(bytes)).expect("open fixture");
    let snapshot = package
        .styles_with_effects(Owner::MainDocument)
        .expect("load effects owner");
    let clone = snapshot.clone();
    let source = snapshot
        .resource()
        .expect("main effects resource")
        .xml_bytes();
    let cloned_source = clone
        .resource()
        .expect("cloned main effects resource")
        .xml_bytes();
    assert_eq!(source, cloned_source);
    assert_eq!(source.as_ptr(), cloned_source.as_ptr());

    let commit = snapshot
        .edit()
        .commit()
        .expect("commit no-op effects transaction");
    assert_eq!(commit.patch().before(), commit.patch().after());
    assert_eq!(
        commit
            .snapshot()
            .resource()
            .expect("no-op resource")
            .xml_bytes()
            .as_ptr(),
        source.as_ptr()
    );
}

#[test]
fn source_resource_replace_inverse_stale_and_save_reopen_are_owner_bound() {
    let original = fixture_bytes("poi/test-data/document/Bug54849.docx").expect("read fixture");
    let mut package = Package::from_reader(Cursor::new(original)).expect("open fixture");
    let base = package
        .styles_with_effects(Owner::MainDocument)
        .expect("load main effects");
    let replacement = changed_resource(&base);

    let mut edit = base.edit();
    edit.replace_resource(Some(replacement))
        .expect("stage effects replacement");
    let commit = edit.commit().expect("commit effects replacement");
    assert_ne!(commit.patch().before(), commit.patch().after());
    let patch = commit.patch().clone();
    let changed_bytes = package_bytes(&mut package);

    package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
        .expect("publish effects replacement");
    assert_ne!(package_bytes(&mut package), changed_bytes);
    let published = package
        .styles_with_effects(Owner::MainDocument)
        .expect("reload published effects");
    assert!(
        published
            .resource()
            .expect("published resource")
            .xml_bytes()
            .windows(b"data=\"changed\"".len())
            .any(|window| window == b"data=\"changed\"")
    );

    let after_failed_stale = package_bytes(&mut package);
    assert!(
        package
            .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
            .is_err()
    );
    assert_eq!(package_bytes(&mut package), after_failed_stale);

    let inverse = patch.inverse();
    package
        .apply_styles_with_effects_patch(Owner::MainDocument, &inverse)
        .expect("publish inverse effects patch");
    assert_eq!(
        package
            .styles_with_effects(Owner::MainDocument)
            .expect("reload inverse effects")
            .resource()
            .expect("inverse resource")
            .xml_bytes(),
        base.resource().expect("base resource").xml_bytes()
    );

    let mut output = Cursor::new(Vec::new());
    package
        .to_stream(&mut output)
        .expect("save effects package");
    let reopened =
        Package::from_reader(Cursor::new(output.into_inner())).expect("reopen effects package");
    assert_eq!(
        reopened
            .styles_with_effects(Owner::MainDocument)
            .expect("read reopened effects")
            .resource()
            .expect("reopened resource")
            .xml_bytes(),
        base.resource().expect("base resource").xml_bytes()
    );
}

#[test]
fn public_patch_rejects_cross_owner_and_cross_package_snapshots_atomically() -> FixtureResult<()> {
    let source = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let mut package = Package::from_reader(Cursor::new(source.clone()))
        .map_err(|error| FixtureError(format!("open cross-owner source: {error}")))?;
    let main = package
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| FixtureError(format!("load cross-owner main: {error}")))?;
    let glossary_before = archive_member(&source, "word/glossary/stylesWithEffects.xml")?;
    let replacement = changed_resource(&main);
    let mut edit = main.edit();
    edit.replace_resource(Some(replacement.clone()))
        .map_err(|error| FixtureError(format!("stage cross-owner replacement: {error}")))?;
    let patch = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit cross-owner replacement: {error}")))?
        .into_patch();

    let before_cross_owner = package_bytes(&mut package);
    assert!(
        package
            .apply_styles_with_effects_patch(Owner::Glossary, &patch)
            .is_err(),
        "a main-owner patch must reject a glossary snapshot"
    );
    assert_eq!(
        package_bytes(&mut package),
        before_cross_owner,
        "cross-owner rejection changed package bytes"
    );
    assert_eq!(
        archive_member(
            &package_bytes(&mut package),
            "word/glossary/stylesWithEffects.xml"
        )?,
        glossary_before,
        "cross-owner rejection changed the independent glossary part"
    );

    let other_source =
        fixture_bytes("libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx")?;
    let mut other = Package::from_reader(Cursor::new(other_source.clone()))
        .map_err(|error| FixtureError(format!("open cross-package source: {error}")))?;
    let other_before = package_bytes(&mut other);
    assert!(
        other
            .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
            .is_err(),
        "a patch from another package snapshot must be rejected"
    );
    assert_eq!(
        package_bytes(&mut other),
        other_before,
        "cross-package rejection changed package bytes"
    );
    Ok(())
}

#[test]
fn public_replace_inverse_restores_exact_package_bytes() -> FixtureResult<()> {
    let original = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let mut package = Package::from_reader(Cursor::new(original.clone()))
        .map_err(|error| FixtureError(format!("open exact-inverse source: {error}")))?;
    let snapshot = package
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| FixtureError(format!("load exact-inverse source: {error}")))?;
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(changed_resource(&snapshot)))
        .map_err(|error| FixtureError(format!("stage exact-inverse replacement: {error}")))?;
    let patch = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit exact-inverse replacement: {error}")))?
        .into_patch();
    package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
        .map_err(|error| FixtureError(format!("publish exact-inverse replacement: {error}")))?;
    assert_ne!(package_bytes(&mut package), original);

    package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch.inverse())
        .map_err(|error| FixtureError(format!("publish exact-inverse restoration: {error}")))?;
    assert_eq!(
        package_bytes(&mut package),
        original,
        "inverse publication must restore every package byte"
    );
    Ok(())
}

#[test]
fn source_bound_clear_and_inverse_restore_custom_target_and_relationship() -> FixtureResult<()> {
    let source = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let custom_member = "word/effects/stylesWithEffects-main.xml";
    let custom_target = "effects/stylesWithEffects-main.xml";
    let custom_relationship_id = "rIdEffectsMain";

    let synthetic = rename_member(&source, "word/stylesWithEffects.xml", custom_member)?;
    let relationships_name = "word/_rels/document.xml.rels";
    let relationships = archive_member(&synthetic, relationships_name)?;
    let relationships = replace_once(
        &relationships,
        br#"Id="rId3""#,
        format!(r#"Id="{custom_relationship_id}""#).as_bytes(),
    )?;
    let relationships = replace_once(
        &relationships,
        br#"Target="stylesWithEffects.xml""#,
        format!(r#"Target="{custom_target}""#).as_bytes(),
    )?;
    let synthetic = replace_member(&synthetic, relationships_name, relationships)?;
    let content_types_name = "[Content_Types].xml";
    let content_types = archive_member(&synthetic, content_types_name)?;
    let content_types = replace_once(
        &content_types,
        br#"PartName="/word/stylesWithEffects.xml""#,
        format!(r#"PartName="/{custom_member}""#).as_bytes(),
    )?;
    let synthetic = replace_member(&synthetic, content_types_name, content_types)?;
    let original_effects = archive_member(&synthetic, custom_member)?;
    let original_relationships = archive_member(&synthetic, relationships_name)?;
    let original_content_types = archive_member(&synthetic, content_types_name)?;
    let relationship_binding = format!(
        r#"Id="{custom_relationship_id}" Type="{}" Target="{custom_target}""#,
        std::str::from_utf8(EFFECTS_RELATIONSHIP).expect("effects relationship URI is UTF-8")
    );
    assert!(
        original_relationships
            .windows(relationship_binding.len())
            .any(|window| window == relationship_binding.as_bytes())
    );

    let mut package = Package::from_reader(Cursor::new(synthetic)).expect("open synthetic source");
    let base = package
        .styles_with_effects(Owner::MainDocument)
        .expect("load custom-target effects owner");
    assert_eq!(
        base.resource().expect("custom-target resource").xml_bytes(),
        original_effects
    );

    let mut edit = base.edit();
    assert!(!edit.is_changed());
    edit.clear_resource();
    assert!(edit.is_changed());
    let commit = edit.commit().expect("commit effects removal");
    assert!(commit.changed());
    assert_eq!(commit.patch().owner(), Owner::MainDocument);
    assert_eq!(
        commit.patch().before().expect("removal source").xml_bytes(),
        original_effects
    );
    assert!(commit.patch().after().is_none());
    let patch = commit.patch().clone();

    let removed = package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
        .expect("publish effects removal");
    assert!(removed.is_empty());
    let removed_bytes = package_bytes(&mut package);
    assert!(archive_member(&removed_bytes, custom_member).is_err());
    let removed_relationships = archive_member(&removed_bytes, relationships_name)?;
    assert!(
        !removed_relationships
            .windows(EFFECTS_RELATIONSHIP.len())
            .any(|window| window == EFFECTS_RELATIONSHIP)
    );

    let restored = package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch.inverse())
        .expect("publish effects inverse");
    assert_eq!(
        restored
            .resource()
            .expect("restored effects resource")
            .xml_bytes(),
        original_effects
    );
    let restored_bytes = package_bytes(&mut package);
    assert_eq!(
        archive_member(&restored_bytes, custom_member)?,
        original_effects
    );
    assert_eq!(
        archive_member(&restored_bytes, relationships_name)?,
        original_relationships
    );
    assert_eq!(
        archive_member(&restored_bytes, content_types_name)?,
        original_content_types
    );
    let restored_relationships = archive_member(&restored_bytes, relationships_name)?;
    assert!(
        restored_relationships
            .windows(relationship_binding.len())
            .any(|window| window == relationship_binding.as_bytes())
    );

    let reopened =
        Package::from_reader(Cursor::new(restored_bytes)).expect("reopen inverse source");
    assert_eq!(
        reopened
            .styles_with_effects(Owner::MainDocument)
            .expect("read reopened inverse effects")
            .resource()
            .expect("reopened effects resource")
            .xml_bytes(),
        original_effects
    );
    Ok(())
}

#[test]
fn package_absent_owner_edit_keeps_source_owner_precondition() -> FixtureResult<()> {
    let source = fixture_bytes("poi/test-data/xmldsign/ms-office-2010-signed.docx")?;
    let relationships_name = "word/_rels/document.xml.rels";
    let relationships = archive_member(&source, relationships_name)?;
    let relationships = replace_once(
        &relationships,
        br#"<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/>"#,
        b"",
    )?;
    let source = replace_member(&source, relationships_name, relationships)?;
    let content_types = archive_member(&source, "[Content_Types].xml")?;
    let content_types = replace_once(
        &content_types,
        br#"<Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/>"#,
        b"",
    )?;
    let source = replace_member(&source, "[Content_Types].xml", content_types)?;
    let source = remove_member(&source, "word/stylesWithEffects.xml")?;
    let package = Package::from_reader(Cursor::new(source)).expect("open absent-owner source");
    let snapshot = package
        .styles_with_effects(Owner::MainDocument)
        .expect("load absent main effects owner");
    assert!(snapshot.is_empty());
    let mut edit = snapshot.edit();
    edit.clear_resource();
    let commit = edit.commit().expect("commit absent-owner no-op");
    assert!(!commit.changed());
    assert!(
        commit
            .patch()
            .apply_to_snapshot(&Snapshot::empty(Owner::MainDocument))
            .is_err()
    );
    Ok(())
}

#[test]
fn explicit_resource_crud_does_not_create_missing_glossary_or_seed_a_fifth_part() {
    let resource = fixture_resource(
        "poi/test-data/xmldsign/ms-office-2010-signed.docx",
        "word/stylesWithEffects.xml",
    )
    .expect("parse effects resource");

    let mut no_glossary = Package::new().expect("new DOCX package");
    assert!(
        no_glossary
            .styles_with_effects(Owner::Glossary)
            .expect("read absent glossary owner")
            .is_empty()
    );
    let before = package_bytes(&mut no_glossary);
    assert!(
        !no_glossary
            .remove_styles_with_effects(Owner::MainDocument)
            .expect("missing main effects is a no-op")
    );
    assert_eq!(package_bytes(&mut no_glossary), before);
    assert!(
        no_glossary
            .put_styles_with_effects(Owner::Glossary, resource.clone())
            .is_err()
    );
    assert_eq!(package_bytes(&mut no_glossary), before);

    let mut with_glossary = Package::new().expect("new DOCX package");
    with_glossary
        .put_glossary(
            litchi_docx::glossary::Catalog::new(),
            litchi_docx::glossary::Conformance::Transitional,
        )
        .expect("create glossary without effects resource");
    assert!(
        with_glossary
            .styles_with_effects(Owner::Glossary)
            .expect("read absent glossary effects")
            .is_empty()
    );
    assert_eq!(
        with_glossary
            .glossary_graph()
            .expect("read glossary graph")
            .expect("glossary graph")
            .parts
            .len(),
        4,
        "glossary seed must remain four resources"
    );
    assert!(
        with_glossary
            .put_styles_with_effects(Owner::Glossary, resource)
            .expect("explicitly create glossary effects resource")
    );
    assert!(
        with_glossary
            .styles_with_effects(Owner::Glossary)
            .expect("read explicitly created glossary effects")
            .resource()
            .is_some()
    );
}

#[test]
fn signed_effects_noop_preserves_source_and_changed_publication_requires_unsign() {
    let original = fixture_bytes("poi/test-data/xmldsign/ms-office-2010-signed.docx")
        .expect("read signed effects fixture");
    let mut package =
        Package::from_reader(Cursor::new(original.clone())).expect("open signed fixture");
    assert!(package.is_signed());
    let snapshot = package
        .styles_with_effects(Owner::MainDocument)
        .expect("load signed effects");
    let same = snapshot
        .resource()
        .expect("signed effects resource")
        .clone();
    assert!(
        !package
            .put_styles_with_effects(Owner::MainDocument, same)
            .expect("signed exact effects no-op")
    );
    assert!(package.is_signed());
    assert_eq!(package_bytes(&mut package), original);

    let changed = changed_resource(&snapshot);
    let before = package_bytes(&mut package);
    assert!(
        package
            .put_styles_with_effects(Owner::MainDocument, changed.clone())
            .is_err()
    );
    assert!(package.is_signed());
    assert_eq!(package_bytes(&mut package), before);

    package.unsign();
    assert!(
        package
            .put_styles_with_effects(Owner::MainDocument, changed)
            .expect("changed effects after explicit unsign")
    );
    assert!(!package.is_signed());
}

#[test]
fn effects_resource_limits_and_malformed_grammar_fail_closed() -> FixtureResult<()> {
    let bytes = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let limit = ReadLimits::builder()
        .max_part_bytes(1)
        .map_err(|error| FixtureError(format!("part limit: {error:?}")))?
        .build()
        .map_err(|error| FixtureError(format!("build read limits: {error:?}")))?;
    assert!(Package::from_reader_with_limits(Cursor::new(bytes.clone()), limit).is_err());

    let malformed = b"<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style></w:styles>";
    assert!(Resource::from_xml(malformed.to_vec()).is_err());
    let wrong_root =
        b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>";
    assert!(Resource::from_xml(wrong_root.to_vec()).is_err());
    Ok(())
}

#[test]
fn effects_resource_rejects_malformed_opaque_xml() {
    let cases = [
        ("unbound descendant", "", "<x:future/>", ""),
        ("invalid QName", "", "<1bad/>", ""),
        ("raw attribute delimiter", " bad=\"<\"", "", ""),
        ("unknown entity", "", "&unknown;", ""),
        ("raw text delimiter", "", "]]>", ""),
        ("control character", "", "\u{1}", ""),
        ("invalid character reference", "", "&#x1;", ""),
        ("empty prefixed binding", " xmlns:x=\"\" x:y=\"z\"", "", ""),
        (
            "reserved default namespace",
            " xmlns=\"http://www.w3.org/XML/1998/namespace\"",
            "",
            "",
        ),
        ("invalid XML version", "", "", "<?xml version=\"2.0\"?>"),
    ];
    let mut accepted = Vec::new();
    for (label, attributes, body, prolog) in cases {
        let xml = format!(
            "{prolog}<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"{attributes}>{body}</w:styles>"
        );
        if Resource::from_xml(xml.into_bytes()).is_ok() {
            accepted.push(label);
        }
    }
    assert!(accepted.is_empty(), "accepted malformed XML: {accepted:?}");
}

#[test]
fn secondary_effects_projection_failures_refuse_public_main_read_and_edit() -> FixtureResult<()> {
    let source = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let main_xml = archive_member(&source, "word/stylesWithEffects.xml")?;
    let secondary_cases: [(&str, &[u8]); 2] = [
        (
            "unsupported MustUnderstand",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:u="urn:unsupported" mc:MustUnderstand="u">
  <w:style w:type="paragraph" w:styleId="secondary"/>
</w:styles>"#,
        ),
        (
            "duplicate typed numId",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="secondary">
    <w:pPr><w:numPr><w:numId w:val="7"/><w:numId w:val="8"/></w:numPr></w:pPr>
  </w:style>
</w:styles>"#,
        ),
    ];

    for (label, secondary_xml) in secondary_cases {
        let synthetic = replace_member(
            &source,
            "word/glossary/stylesWithEffects.xml",
            secondary_xml.to_vec(),
        )?;
        assert_eq!(
            archive_member(&synthetic, "word/stylesWithEffects.xml")?,
            main_xml,
            "secondary {label} fixture changed the main effects source"
        );

        let mut read_package = Package::from_reader(Cursor::new(synthetic.clone()))
            .map_err(|error| FixtureError(format!("open secondary {label} fixture: {error}")))?;
        let before_read = package_bytes(&mut read_package);
        assert!(
            read_package.styles_with_effects(Owner::Glossary).is_err(),
            "secondary {label} fixture was accepted by its direct owner"
        );
        assert!(
            read_package
                .styles_with_effects(Owner::MainDocument)
                .is_err(),
            "main read bypassed secondary {label} projection failure"
        );
        assert_eq!(
            package_bytes(&mut read_package),
            before_read,
            "main read refusal changed package bytes for secondary {label}"
        );

        let mut edit_package = Package::from_reader(Cursor::new(synthetic))
            .map_err(|error| FixtureError(format!("reopen secondary {label} fixture: {error}")))?;
        let before_edit = package_bytes(&mut edit_package);
        assert!(
            edit_package
                .remove_styles_with_effects(Owner::MainDocument)
                .is_err(),
            "main edit bypassed secondary {label} projection failure"
        );
        assert_eq!(
            package_bytes(&mut edit_package),
            before_edit,
            "main edit refusal changed package bytes for secondary {label}"
        );
    }
    Ok(())
}

#[test]
fn native_effects_read_honors_caller_event_and_depth_limits() -> FixtureResult<()> {
    let bytes = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let xml = archive_member(&bytes, "word/stylesWithEffects.xml")?;
    let mut reader = quick_xml::Reader::from_reader(xml.as_slice());
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut maximum_depth = 0usize;
    loop {
        events += 1;
        match reader.read_event().expect("native effects XML") {
            quick_xml::events::Event::Start(_) => {
                depth += 1;
                maximum_depth = maximum_depth.max(depth);
            },
            quick_xml::events::Event::End(_) => depth -= 1,
            quick_xml::events::Event::Eof => break,
            _ => {},
        }
    }
    assert!(maximum_depth > 2);
    let limits = [
        ReadLimits::builder()
            .max_xml_events(events - 1)
            .unwrap()
            .build()
            .unwrap(),
        ReadLimits::builder()
            .max_xml_depth(maximum_depth - 1)
            .unwrap()
            .build()
            .unwrap(),
    ];
    for limit in limits {
        let package = litchi_opc::OpcPackage::from_vec_with_limits(bytes.clone(), limit)
            .expect("native OPC package fits ingress limits");
        assert!(litchi_docx::styles::effects::load(&package, Owner::MainDocument).is_err());
    }
    Ok(())
}

#[test]
fn effects_owner_rejects_duplicate_relationship_and_third_orphan_part() -> FixtureResult<()> {
    let base = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    assert_main_owner_rejected(remove_member(&base, "word/stylesWithEffects.xml")?);
    let rels = archive_member(&base, "word/_rels/document.xml.rels")?;
    let duplicate_rel = br#"<Relationship Id="rIdEffectsDuplicate" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/>"#;
    let duplicate_rels = replace_once(&rels, b"</Relationships>", duplicate_rel)?;
    assert_main_owner_rejected(replace_member(
        &base,
        "word/_rels/document.xml.rels",
        duplicate_rels,
    )?);

    let orphan_xml = archive_member(&base, "word/stylesWithEffects.xml")?;
    let orphaned = add_member(&base, "word/orphanStylesWithEffects.xml", orphan_xml)?;
    let content_types = archive_member(&orphaned, "[Content_Types].xml")?;
    let orphan_content_types = replace_once(
        &content_types,
        b"</Types>",
        br#"<Override PartName="/word/orphanStylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/></Types>"#,
    )?;
    let orphaned = replace_member(&orphaned, "[Content_Types].xml", orphan_content_types)?;
    assert_main_owner_rejected(orphaned);
    Ok(())
}

#[test]
fn effects_owner_rejects_external_and_wrong_content_type_targets() -> FixtureResult<()> {
    let base = fixture_bytes("poi/test-data/xmldsign/ms-office-2010-signed.docx")?;
    let rels = archive_member(&base, "word/_rels/document.xml.rels")?;
    let external_rels = replace_once(
        &rels,
        b"Target=\"stylesWithEffects.xml\"",
        b"Target=\"https://example.invalid/stylesWithEffects.xml\" TargetMode=\"External\"",
    )?;
    assert_main_owner_rejected(replace_member(
        &base,
        "word/_rels/document.xml.rels",
        external_rels,
    )?);

    let content_types = archive_member(&base, "[Content_Types].xml")?;
    let wrong_content_type =
        replace_once(&content_types, EFFECTS_CONTENT_TYPE, b"application/xml")?;
    assert_main_owner_rejected(replace_member(
        &base,
        "[Content_Types].xml",
        wrong_content_type,
    )?);
    Ok(())
}

#[test]
fn effects_owner_rejects_leaf_outbound_and_shared_inbound_edges() -> FixtureResult<()> {
    let base = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let outbound_rels = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOutbound" Type="urn:test:outbound" Target="https://example.invalid/outbound" TargetMode="External"/></Relationships>"#;
    let outbound = add_member(
        &base,
        "word/_rels/stylesWithEffects.xml.rels",
        outbound_rels.to_vec(),
    )?;
    assert_main_owner_rejected(outbound);

    let glossary_rels = archive_member(&base, "word/glossary/_rels/document.xml.rels")?;
    let shared = br#"<Relationship Id="rIdSharedEffects" Type="urn:test:shared-inbound" Target="../stylesWithEffects.xml"/>"#;
    let shared_with_close = [shared.as_slice(), b"</Relationships>"].concat();
    let shared_rels = replace_once(&glossary_rels, b"</Relationships>", &shared_with_close)?;
    let shared_package =
        replace_member(&base, "word/glossary/_rels/document.xml.rels", shared_rels)?;
    let mut package =
        Package::from_reader(Cursor::new(shared_package)).expect("open shared inbound fixture");
    let before = package_bytes(&mut package);
    assert!(
        package
            .remove_styles_with_effects(Owner::MainDocument)
            .is_err()
    );
    assert_eq!(package_bytes(&mut package), before);
    Ok(())
}

#[test]
fn effects_namespace_prefix_and_unknown_xml_are_source_preserving() -> FixtureResult<()> {
    let base = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let source = archive_member(&base, "word/stylesWithEffects.xml")?;
    let prefixed = replace_once(
        &replace_all(&source, b"w:", b"x:"),
        b"xmlns:w=",
        b"xmlns:x=",
    )?;
    let parsed = Resource::from_xml(prefixed.clone())
        .map_err(|error| FixtureError(format!("namespace-prefix parse: {error}")))?;
    assert_eq!(parsed.xml_bytes(), prefixed);

    let unknown = replace_once(
        &source,
        b"</w:styles>",
        br#"<w:unknownExtension xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" data="preserve"/></w:styles>"#,
    )?;
    let resource = Resource::from_xml(unknown.clone())
        .map_err(|error| FixtureError(format!("unknown-extension parse: {error}")))?;
    let mut package = Package::from_reader(Cursor::new(base)).expect("open source fixture");
    package
        .put_styles_with_effects(Owner::MainDocument, resource)
        .expect("publish source-preserving unknown extension");
    let reopened = Package::from_reader(Cursor::new(package_bytes(&mut package)))
        .expect("reopen source-preserving fixture");
    assert_eq!(
        reopened
            .styles_with_effects(Owner::MainDocument)
            .expect("read source-preserving effects")
            .resource()
            .expect("source-preserving effects resource")
            .xml_bytes(),
        unknown
    );
    Ok(())
}

#[test]
fn effects_addition_honors_public_package_caps_atomically() -> FixtureResult<()> {
    let native = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let resource = Resource::from_xml(archive_member(&native, "word/stylesWithEffects.xml")?)
        .map_err(|error| FixtureError(format!("parse effects addition resource: {error}")))?;
    let source = fixture_without_main_effects("poi/test-data/document/Bug54849.docx")?;
    let source_metrics = package_metrics(&source)?;

    let mut generous = Package::from_reader(Cursor::new(source.clone()))
        .map_err(|error| FixtureError(format!("open generous baseline: {error}")))?;
    assert!(
        generous
            .put_styles_with_effects(Owner::MainDocument, resource.clone())
            .map_err(|error| FixtureError(format!("derive generous addition: {error}")))?
    );
    let projected = package_bytes(&mut generous);
    let projected_metrics = package_metrics(&projected)?;
    assert!(projected_metrics.parts > source_metrics.parts);
    assert!(projected_metrics.total_part_bytes > source_metrics.total_part_bytes);
    assert!(projected_metrics.total_relationships > source_metrics.total_relationships);
    assert!(projected_metrics.relationship_xml_events > source_metrics.relationship_xml_events);

    for cap in MutationCap::ADDITION {
        let source_value = cap.source_value(source_metrics);
        assert_effects_add_rejected_under_cap(
            &source,
            &resource,
            cap.set(source_value)?,
            cap.label(),
            cap,
        )?;
    }

    for cap in MutationCap::ADDITION {
        let source_value = cap.source_value(source_metrics);
        let projected_value = cap.source_value(projected_metrics);
        assert!(
            projected_value > source_value,
            "{} metric did not grow after addition",
            cap.label()
        );
        let one_under = projected_value - 1;
        if one_under >= source_value {
            assert_effects_add_rejected_under_cap(
                &source,
                &resource,
                cap.set(one_under)?,
                cap.label(),
                cap,
            )?;
        }
    }

    let exact_limits = ReadLimits::builder()
        .max_parts(projected_metrics.parts)
        .map_err(|error| FixtureError(format!("set exact part cap: {error:?}")))?
        .max_total_part_bytes(projected_metrics.total_part_bytes)
        .map_err(|error| FixtureError(format!("set exact part-byte cap: {error:?}")))?
        .max_total_relationships(projected_metrics.total_relationships)
        .map_err(|error| FixtureError(format!("set exact relationship cap: {error:?}")))?
        .max_total_relationship_xml_events(projected_metrics.relationship_xml_events)
        .map_err(|error| FixtureError(format!("set exact event cap: {error:?}")))?
        .max_total_relationship_xml_bytes(projected_metrics.relationship_xml_bytes)
        .map_err(|error| FixtureError(format!("set exact XML byte cap: {error:?}")))?
        .build()
        .map_err(|error| FixtureError(format!("build exact mutation caps: {error:?}")))?;
    let mut exact = Package::from_reader_with_limits(Cursor::new(source.clone()), exact_limits)
        .map_err(|error| FixtureError(format!("open exact-fit baseline: {error}")))?;
    assert_eq!(package_bytes(&mut exact), source);
    assert!(
        exact
            .put_styles_with_effects(Owner::MainDocument, resource.clone())
            .map_err(|error| FixtureError(format!("exact-fit put: {error}")))?
    );
    assert_eq!(
        package_metrics(&package_bytes(&mut exact))?,
        projected_metrics
    );

    let mut exact_patch =
        Package::from_reader_with_limits(Cursor::new(source.clone()), exact_limits)
            .map_err(|error| FixtureError(format!("open exact-fit patch baseline: {error}")))?;
    let snapshot = exact_patch
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| FixtureError(format!("read exact-fit patch baseline: {error}")))?;
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(resource))
        .map_err(|error| FixtureError(format!("stage exact-fit addition patch: {error}")))?;
    let patch = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit exact-fit addition patch: {error}")))?
        .into_patch();
    assert!(
        exact_patch
            .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
            .map_err(|error| FixtureError(format!("exact-fit addition patch: {error}")))?
            .resource()
            .is_some()
    );
    assert_eq!(
        package_metrics(&package_bytes(&mut exact_patch))?,
        projected_metrics
    );
    Ok(())
}

#[test]
fn effects_addition_honors_relationship_part_and_graph_caps_with_source_less_owner()
-> FixtureResult<()> {
    let native = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let resource = Resource::from_xml(archive_member(&native, "word/stylesWithEffects.xml")?)
        .map_err(|error| FixtureError(format!("parse effects source-less addition: {error}")))?;
    let source = fixture_without_main_effects_or_owner_relationships(
        "ooxml/docx/ComplexNumberedLists.docx",
    )?;
    let source_metrics = package_metrics(&source)?;
    let owner_relationships = "word/_rels/document.xml.rels";
    assert!(archive_member(&source, owner_relationships).is_err());

    let mut generous = Package::from_reader(Cursor::new(source.clone()))
        .map_err(|error| FixtureError(format!("open source-less owner baseline: {error}")))?;
    assert!(
        generous
            .put_styles_with_effects(Owner::MainDocument, resource.clone())
            .map_err(|error| FixtureError(format!("create source-less owner rels: {error}")))?
    );
    let projected = package_bytes(&mut generous);
    let projected_metrics = package_metrics(&projected)?;
    assert!(
        projected_metrics.relationship_parts > source_metrics.relationship_parts,
        "source-less owner creation must add a physical relationship member"
    );
    assert!(
        projected_metrics.relationship_graph_nodes > source_metrics.relationship_graph_nodes,
        "effects target must add one reachable relationship graph node"
    );
    let projected_owner_relationships = archive_member(&projected, owner_relationships)?;
    assert!(
        projected_owner_relationships
            .windows(EFFECTS_RELATIONSHIP.len())
            .any(|window| window == EFFECTS_RELATIONSHIP),
        "source-less owner relationship member must contain the effects edge"
    );

    for cap in MutationCap::TOPOLOGY {
        let source_value = cap.source_value(source_metrics);
        let projected_value = cap.source_value(projected_metrics);
        assert!(
            projected_value > source_value,
            "{} metric did not grow after source-less owner addition",
            cap.label()
        );
        assert_effects_add_rejected_under_cap(
            &source,
            &resource,
            cap.set(projected_value - 1)?,
            &format!("{} one-under", cap.label()),
            cap,
        )?;
    }

    let exact_limits = ReadLimits::builder()
        .max_relationship_parts(projected_metrics.relationship_parts)
        .map_err(|error| FixtureError(format!("set exact relationship-part cap: {error:?}")))?
        .max_relationship_graph_nodes(projected_metrics.relationship_graph_nodes)
        .map_err(|error| FixtureError(format!("set exact graph-node cap: {error:?}")))?
        .build()
        .map_err(|error| FixtureError(format!("build exact topology caps: {error:?}")))?;
    let mut exact = Package::from_reader_with_limits(Cursor::new(source.clone()), exact_limits)
        .map_err(|error| FixtureError(format!("open source-less exact-fit baseline: {error}")))?;
    assert_eq!(package_bytes(&mut exact), source);
    assert!(
        exact
            .put_styles_with_effects(Owner::MainDocument, resource.clone())
            .map_err(|error| FixtureError(format!("source-less exact-fit put: {error}")))?
    );
    assert_eq!(
        package_metrics(&package_bytes(&mut exact))?,
        projected_metrics
    );

    let mut exact_patch = Package::from_reader_with_limits(Cursor::new(source), exact_limits)
        .map_err(|error| {
            FixtureError(format!(
                "open source-less exact-fit patch baseline: {error}"
            ))
        })?;
    let snapshot = exact_patch
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| {
            FixtureError(format!(
                "read source-less exact-fit patch baseline: {error}"
            ))
        })?;
    let mut edit = snapshot.edit();
    edit.replace_resource(Some(resource))
        .map_err(|error| FixtureError(format!("stage source-less exact-fit patch: {error}")))?;
    let patch = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit source-less exact-fit patch: {error}")))?
        .into_patch();
    assert!(
        exact_patch
            .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
            .map_err(|error| FixtureError(format!("apply source-less exact-fit patch: {error}")))?
            .resource()
            .is_some()
    );
    assert_eq!(
        package_metrics(&package_bytes(&mut exact_patch))?,
        projected_metrics
    );
    Ok(())
}

fn assert_replacement_honors_total_part_bytes_cap(
    source: &[u8],
    owner: Owner,
) -> FixtureResult<()> {
    let source_metrics = package_metrics(source)?;
    let mut generous = Package::from_reader(Cursor::new(source.to_vec())).map_err(|error| {
        FixtureError(format!(
            "open generous {owner:?} replacement source: {error}"
        ))
    })?;
    let base = generous
        .styles_with_effects(owner)
        .map_err(|error| FixtureError(format!("load generous {owner:?} replacement: {error}")))?;
    let replacement = changed_resource(&base);
    assert!(
        replacement.xml_bytes().len()
            > base
                .resource()
                .expect("existing owner replacement resource")
                .xml_bytes()
                .len(),
        "replacement must grow the selected {owner:?} part"
    );
    let mut edit = base.edit();
    edit.replace_resource(Some(replacement.clone()))
        .map_err(|error| FixtureError(format!("stage generous {owner:?} replacement: {error}")))?;
    let commit = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit generous {owner:?} replacement: {error}")))?;
    let published = generous
        .apply_styles_with_effects_commit(owner, commit)
        .map_err(|error| {
            FixtureError(format!("publish generous {owner:?} replacement: {error}"))
        })?;
    assert_eq!(
        published
            .resource()
            .expect("published replacement resource")
            .xml_bytes(),
        replacement.xml_bytes()
    );
    let projected = package_bytes(&mut generous);
    let projected_metrics = package_metrics(&projected)?;
    assert!(projected_metrics.total_part_bytes > source_metrics.total_part_bytes);
    let exact_total = projected_metrics.total_part_bytes;
    let one_under = exact_total
        .checked_sub(1)
        .ok_or_else(|| FixtureError("replacement total-part metric underflow".into()))?;
    assert!(
        one_under > source_metrics.total_part_bytes,
        "one-under replacement cap must still admit the source package"
    );

    let exact_limits = ReadLimits::builder()
        .max_total_part_bytes(exact_total)
        .map_err(|error| FixtureError(format!("set exact {owner:?} total-part cap: {error:?}")))?
        .build()
        .map_err(|error| {
            FixtureError(format!("build exact {owner:?} total-part cap: {error:?}"))
        })?;
    let mut exact = Package::from_reader_with_limits(Cursor::new(source.to_vec()), exact_limits)
        .map_err(|error| {
            FixtureError(format!("open exact {owner:?} replacement source: {error}"))
        })?;
    assert_eq!(package_bytes(&mut exact), source);
    let exact_base = exact.styles_with_effects(owner).map_err(|error| {
        FixtureError(format!("load exact {owner:?} replacement source: {error}"))
    })?;
    let mut exact_edit = exact_base.edit();
    exact_edit
        .replace_resource(Some(replacement.clone()))
        .map_err(|error| FixtureError(format!("stage exact {owner:?} replacement: {error}")))?;
    let exact_commit = exact_edit
        .commit()
        .map_err(|error| FixtureError(format!("commit exact {owner:?} replacement: {error}")))?;
    let exact_published = exact
        .apply_styles_with_effects_commit(owner, exact_commit)
        .map_err(|error| FixtureError(format!("publish exact {owner:?} replacement: {error}")))?;
    assert_eq!(
        exact_published
            .resource()
            .expect("exact published replacement resource")
            .xml_bytes(),
        replacement.xml_bytes()
    );
    let exact_output = package_bytes(&mut exact);
    assert_eq!(
        package_metrics(&exact_output)?.total_part_bytes,
        exact_total,
        "exact {owner:?} cap changed the final part-byte metric"
    );
    let reopened = Package::from_reader_with_limits(Cursor::new(exact_output), exact_limits)
        .map_err(|error| FixtureError(format!("reopen exact {owner:?} replacement: {error}")))?;
    assert_eq!(
        reopened
            .styles_with_effects(owner)
            .map_err(|error| FixtureError(format!("read reopened exact {owner:?}: {error}")))?
            .resource()
            .expect("reopened exact replacement resource")
            .xml_bytes(),
        replacement.xml_bytes()
    );

    let one_under_limits = ReadLimits::builder()
        .max_total_part_bytes(one_under)
        .map_err(|error| {
            FixtureError(format!("set one-under {owner:?} total-part cap: {error:?}"))
        })?
        .build()
        .map_err(|error| {
            FixtureError(format!(
                "build one-under {owner:?} total-part cap: {error:?}"
            ))
        })?;
    let mut limited =
        Package::from_reader_with_limits(Cursor::new(source.to_vec()), one_under_limits).map_err(
            |error| {
                FixtureError(format!(
                    "open one-under {owner:?} replacement source: {error}"
                ))
            },
        )?;
    let baseline = package_bytes(&mut limited);
    assert_eq!(baseline, source);
    let limited_base = limited.styles_with_effects(owner).map_err(|error| {
        FixtureError(format!(
            "load one-under {owner:?} replacement source: {error}"
        ))
    })?;
    let mut limited_edit = limited_base.edit();
    limited_edit
        .replace_resource(Some(replacement))
        .map_err(|error| FixtureError(format!("stage one-under {owner:?} replacement: {error}")))?;
    let commit_error = limited_edit
        .commit()
        .expect_err("one-under replacement must fail at transaction commit");
    assert!(
        matches!(
            commit_error,
            litchi_docx::Error::Opc(litchi_opc::OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::TotalPartBytes,
                actual,
                maximum,
            }) if actual == exact_total && maximum == one_under
        ),
        "one-under {owner:?} replacement returned the wrong commit error: {commit_error}"
    );
    assert_eq!(
        package_bytes(&mut limited),
        baseline,
        "one-under {owner:?} commit changed package output"
    );
    let reopened_baseline =
        Package::from_reader_with_limits(Cursor::new(baseline), one_under_limits).map_err(
            |error| FixtureError(format!("reopen one-under {owner:?} baseline: {error}")),
        )?;
    assert_eq!(
        reopened_baseline
            .styles_with_effects(owner)
            .map_err(|error| {
                FixtureError(format!(
                    "read reopened one-under {owner:?} baseline: {error}"
                ))
            })?
            .resource()
            .expect("reopened one-under source resource")
            .xml_bytes(),
        base.resource()
            .expect("generous source resource")
            .xml_bytes()
    );
    Ok(())
}

#[test]
fn existing_owner_replacement_honors_exact_and_one_under_total_part_bytes_caps() -> FixtureResult<()>
{
    let source = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    for owner in [Owner::MainDocument, Owner::Glossary] {
        assert_replacement_honors_total_part_bytes_cap(&source, owner)?;
    }
    Ok(())
}

#[test]
fn absent_owner_inverse_rejects_candidate_with_different_relationship_id() -> FixtureResult<()> {
    let native = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let resource = Resource::from_xml(archive_member(&native, "word/stylesWithEffects.xml")?)
        .map_err(|error| FixtureError(format!("parse stale inverse resource: {error}")))?;
    let source = fixture_without_main_effects("ooxml/docx/ComplexNumberedLists.docx")?;
    let mut package = Package::from_reader(Cursor::new(source))
        .map_err(|error| FixtureError(format!("open stale inverse source: {error}")))?;
    let absent = package
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| FixtureError(format!("read stale inverse absent owner: {error}")))?;
    assert!(absent.is_empty());
    let mut edit = absent.edit();
    edit.replace_resource(Some(resource.clone()))
        .map_err(|error| FixtureError(format!("stage stale inverse addition: {error}")))?;
    let patch = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit stale inverse addition: {error}")))?
        .into_patch();
    package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
        .map_err(|error| FixtureError(format!("publish stale inverse addition: {error}")))?;
    let added = package_bytes(&mut package);
    let relationships_name = "word/_rels/document.xml.rels";
    let relationships = archive_member(&added, relationships_name)?;
    assert!(
        relationships
            .windows(EFFECTS_RELATIONSHIP.len())
            .any(|window| window == EFFECTS_RELATIONSHIP)
    );
    let relationships = replace_once(&relationships, br#"Id="rIdEffects""#, br#"Id="rIdOther""#)?;
    let candidate_bytes = replace_member(&added, relationships_name, relationships)?;
    let mut candidate = Package::from_reader(Cursor::new(candidate_bytes))
        .map_err(|error| FixtureError(format!("open stale inverse candidate: {error}")))?;
    assert_eq!(
        candidate
            .styles_with_effects(Owner::MainDocument)
            .map_err(|error| FixtureError(format!("read stale inverse candidate: {error}")))?
            .resource()
            .expect("candidate effects resource")
            .xml_bytes(),
        resource.xml_bytes()
    );
    let before = package_bytes(&mut candidate);
    assert!(
        candidate
            .apply_styles_with_effects_patch(Owner::MainDocument, &patch.inverse())
            .is_err(),
        "inverse must reject a candidate with a different relationship ID"
    );
    assert_eq!(package_bytes(&mut candidate), before);
    assert!(
        candidate
            .styles_with_effects(Owner::MainDocument)
            .map_err(|error| FixtureError(format!("reload stale inverse candidate: {error}")))?
            .resource()
            .is_some()
    );
    Ok(())
}

#[test]
fn removing_effects_removes_empty_target_relationship_member_and_inverse_restores_it()
-> FixtureResult<()> {
    let base = fixture_bytes("poi/test-data/document/Bug54849.docx")?;
    let target_relationships = "word/_rels/stylesWithEffects.xml.rels";
    let empty_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
    let source = add_member(&base, target_relationships, empty_relationships.to_vec())?;
    let mut package = Package::from_reader(Cursor::new(source.clone()))
        .map_err(|error| FixtureError(format!("open empty-target-rels fixture: {error}")))?;
    assert_eq!(
        archive_member(&source, target_relationships)?,
        empty_relationships,
        "synthetic empty relationship member bytes"
    );
    let original_effects = archive_member(&source, "word/stylesWithEffects.xml")?;
    let original_owner_relationships = archive_member(&source, "word/_rels/document.xml.rels")?;
    let original_content_types = archive_member(&source, "[Content_Types].xml")?;
    let snapshot = package
        .styles_with_effects(Owner::MainDocument)
        .map_err(|error| FixtureError(format!("load empty-target-rels effects: {error}")))?;
    let mut edit = snapshot.edit();
    edit.clear_resource();
    let patch = edit
        .commit()
        .map_err(|error| FixtureError(format!("commit empty-target-rels removal: {error}")))?
        .into_patch();
    let removed = package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch)
        .map_err(|error| FixtureError(format!("remove empty-target-rels effects: {error}")))?;
    assert!(removed.is_empty());
    let removed_bytes = package_bytes(&mut package);
    assert!(archive_member(&removed_bytes, "word/stylesWithEffects.xml").is_err());
    assert!(archive_member(&removed_bytes, target_relationships).is_err());

    let restored = package
        .apply_styles_with_effects_patch(Owner::MainDocument, &patch.inverse())
        .map_err(|error| FixtureError(format!("restore empty-target-rels effects: {error}")))?;
    assert_eq!(
        restored.resource().expect("inverse resource").xml_bytes(),
        original_effects
    );
    let restored_bytes = package_bytes(&mut package);
    assert_eq!(
        archive_member(&restored_bytes, target_relationships)?,
        empty_relationships
    );
    assert_eq!(
        archive_member(&restored_bytes, "word/_rels/document.xml.rels")?,
        original_owner_relationships
    );
    assert_eq!(
        archive_member(&restored_bytes, "[Content_Types].xml")?,
        original_content_types
    );
    let reopened = Package::from_reader(Cursor::new(restored_bytes))
        .map_err(|error| FixtureError(format!("reopen empty-target-rels inverse: {error}")))?;
    assert_eq!(
        reopened
            .styles_with_effects(Owner::MainDocument)
            .map_err(|error| FixtureError(format!("read empty-target-rels inverse: {error}")))?
            .resource()
            .expect("reopened inverse resource")
            .xml_bytes(),
        original_effects
    );
    Ok(())
}
