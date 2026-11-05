//! Guarded raw metadata restoration through `SourceTopologyPlan`.
//!
//! These tests keep the semantic topology operation and its lexical metadata
//! image separate.  The former describes the effective graph; the latter is
//! admitted only when its parsed graph is the same candidate.
#![allow(clippy::expect_used, clippy::unwrap_used)]

use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::ops::Range;
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcError, PackURI, SourceBackedPackage, SourceLimitKind, SourceLimits, SourceTopologyPlan,
};
use litchi_opc::{SourceCacheLimits, SourceReadPolicy};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const DROP_REL: &str = "urn:test:drop";
const DOCUMENT: &str = "/word/document.xml";
const EXISTING: &str = "/custom/existing.xml";
const UNTOUCHED: &str = "/custom/untouched.bin";

const ORIGINAL_CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- before content-type entries -->
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/octet-stream"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.example.document+xml"/>
  <Override PartName="/custom/existing.xml" ContentType="application/vnd.example.existing+xml"/>
  <Override PartName="/custom/_rels/existing.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <!-- after content-type entries -->
</Types>
"#;

const ORIGINAL_PACKAGE_RELATIONSHIPS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"#;

const ORIGINAL_DOCUMENT_RELATIONSHIPS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- retain this relationship and its surrounding trivia -->
  <r:Relationship Id="rIdKeep" Type="urn:test:keep" Target="../custom/untouched.bin"/>
  <!-- remove only this entry -->
  <r:Relationship Target="../custom/existing.xml" Type="urn:test:drop" Id="rIdDrop"/>
</r:Relationships>
"#;

const ORIGINAL_EXISTING_RELATIONSHIPS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- this empty owner member is physically present and must be restorable -->
</Relationships>
"#;

const DOCUMENT_DROP_ENTRY: &[u8] =
    br#"  <r:Relationship Target="../custom/existing.xml" Type="urn:test:drop" Id="rIdDrop"/>
"#;
const EXISTING_OVERRIDE: &[u8] = br#"  <Override PartName="/custom/existing.xml" ContentType="application/vnd.example.existing+xml"/>
"#;
const EXISTING_REL_OVERRIDE: &[u8] = br#"  <Override PartName="/custom/_rels/existing.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
"#;

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn archive_bytes(include_empty_owned_relationships: bool, include_large_part: bool) -> Vec<u8> {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    let content_types = if include_empty_owned_relationships {
        ORIGINAL_CONTENT_TYPES.to_vec()
    } else {
        remove_once(ORIGINAL_CONTENT_TYPES, EXISTING_REL_OVERRIDE)
    };
    writer
        .write_stored("[Content_Types].xml", &content_types)
        .unwrap();
    writer
        .write_stored("_rels/.rels", ORIGINAL_PACKAGE_RELATIONSHIPS)
        .unwrap();
    writer
        .write_stored("word/document.xml", b"<document/>")
        .unwrap();
    writer
        .write_stored(
            "word/_rels/document.xml.rels",
            ORIGINAL_DOCUMENT_RELATIONSHIPS,
        )
        .unwrap();
    writer
        .write_stored("custom/existing.xml", b"<existing/>")
        .unwrap();
    if include_empty_owned_relationships {
        writer
            .write_stored(
                "custom/_rels/existing.xml.rels",
                ORIGINAL_EXISTING_RELATIONSHIPS,
            )
            .unwrap();
    }
    if include_large_part {
        writer
            .write_stored(UNTOUCHED.trim_start_matches('/'), &vec![b'L'; 128 * 1024])
            .unwrap();
    } else {
        writer
            .write_stored("custom/untouched.bin", b"untouched")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn package_bytes(include_empty_owned_relationships: bool, include_large_part: bool) -> Vec<u8> {
    let bytes = archive_bytes(include_empty_owned_relationships, include_large_part);
    // The hand-authored archive intentionally carries a noncanonical lexical
    // content-types member.  Reopen it once so the source catalog validates
    // the physical graph before the source-backed tests use it.
    SourceBackedPackage::from_vec(
        bytes.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .expect("fixture source must be admitted");
    bytes
}

fn signed_package_bytes() -> Vec<u8> {
    let signature_override = format!(
        "  <Override PartName=\"/_xmlsignatures/origin.sigs\" ContentType=\"{}\"/>\n</Types>\n",
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
    );
    let content_types = replace_once(
        ORIGINAL_CONTENT_TYPES,
        b"</Types>\n",
        signature_override.as_bytes(),
    );
    let signature_relationship = format!(
        "  <Relationship Id=\"rIdSignatureOrigin\" Type=\"{}\" Target=\"_xmlsignatures/origin.sigs\"/>\n</Relationships>\n",
        rt::DIGITAL_SIGNATURE_ORIGIN,
    );
    let package_relationships = replace_once(
        ORIGINAL_PACKAGE_RELATIONSHIPS,
        b"</Relationships>\n",
        signature_relationship.as_bytes(),
    );

    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored("[Content_Types].xml", &content_types)
        .unwrap();
    writer
        .write_stored("_rels/.rels", &package_relationships)
        .unwrap();
    writer
        .write_stored("word/document.xml", b"<document/>")
        .unwrap();
    writer
        .write_stored(
            "word/_rels/document.xml.rels",
            ORIGINAL_DOCUMENT_RELATIONSHIPS,
        )
        .unwrap();
    writer
        .write_stored("custom/existing.xml", b"<existing/>")
        .unwrap();
    writer
        .write_stored(
            "custom/_rels/existing.xml.rels",
            ORIGINAL_EXISTING_RELATIONSHIPS,
        )
        .unwrap();
    writer
        .write_stored("custom/untouched.bin", b"untouched")
        .unwrap();
    writer
        .write_stored("_xmlsignatures/origin.sigs", b"<origin/>")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "opc-raw-metadata-restore-integration",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

fn publish(source: &[u8], plan: SourceTopologyPlan) -> litchi_opc::Result<Vec<u8>> {
    let package = SourceBackedPackage::from_vec(
        source.to_vec(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )?;
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan)?;
    Ok(output)
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .read(name)
        .unwrap()
}

fn optional_member(bytes: &[u8], name: &str) -> Option<Vec<u8>> {
    ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .file_names()
        .any(|candidate| candidate == name)
        .then(|| member(bytes, name))
}

fn rewrite_member(source: &[u8], name: &str, replacement: &[u8]) -> Vec<u8> {
    let names = ArchiveReader::new(source, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .file_names()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for member_name in names {
        let bytes = if member_name == name {
            replacement.to_vec()
        } else {
            member(source, &member_name)
        };
        writer.write_stored(&member_name, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn remove_once(bytes: &[u8], fragment: &[u8]) -> Vec<u8> {
    let start = bytes
        .windows(fragment.len())
        .position(|window| window == fragment)
        .expect("fixture fragment must exist");
    let mut result = Vec::with_capacity(bytes.len() - fragment.len());
    result.extend_from_slice(&bytes[..start]);
    result.extend_from_slice(&bytes[start + fragment.len()..]);
    result
}

fn replace_once(bytes: &[u8], before: &[u8], after: &[u8]) -> Vec<u8> {
    let start = bytes
        .windows(before.len())
        .position(|window| window == before)
        .expect("fixture fragment must exist");
    let mut result = Vec::with_capacity(bytes.len() - before.len() + after.len());
    result.extend_from_slice(&bytes[..start]);
    result.extend_from_slice(after);
    result.extend_from_slice(&bytes[start + before.len()..]);
    result
}

fn removal_plan(
    original_types: &[u8],
    removed_types: &[u8],
    original_document_relationships: &[u8],
    removed_document_relationships: &[u8],
    original_owned_relationships: Option<&[u8]>,
) -> SourceTopologyPlan {
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
        .unwrap();
    plan.try_remove_part(pack(EXISTING)).unwrap();
    plan.try_restore_relationships_xml(
        pack(DOCUMENT),
        Some(Arc::new(original_document_relationships.to_vec())),
        Some(Arc::new(removed_document_relationships.to_vec())),
    )
    .unwrap();
    plan.try_restore_relationships_xml(
        pack(EXISTING),
        original_owned_relationships.map(|bytes| Arc::new(bytes.to_vec())),
        None,
    )
    .unwrap();
    plan.try_restore_content_types_xml(
        Arc::new(original_types.to_vec()),
        Arc::new(removed_types.to_vec()),
    )
    .unwrap();
    plan
}

#[test]
fn raw_metadata_remove_reopen_and_restore_recovers_exact_xml() {
    let source = package_bytes(true, false);
    let original_types = member(&source, "[Content_Types].xml");
    let original_document_relationships = member(&source, "word/_rels/document.xml.rels");
    let original_owned_relationships = member(&source, "custom/_rels/existing.xml.rels");
    let removed_types = remove_once(
        &remove_once(&original_types, EXISTING_OVERRIDE),
        EXISTING_REL_OVERRIDE,
    );
    let removed_document_relationships =
        remove_once(&original_document_relationships, DOCUMENT_DROP_ENTRY);

    let removed = publish(
        &source,
        removal_plan(
            &original_types,
            &removed_types,
            &original_document_relationships,
            &removed_document_relationships,
            Some(&original_owned_relationships),
        ),
    )
    .expect("semantic removal with a guarded raw image must publish");
    assert_eq!(member(&removed, "[Content_Types].xml"), removed_types);
    assert_eq!(
        member(&removed, "word/_rels/document.xml.rels"),
        removed_document_relationships
    );
    assert!(optional_member(&removed, "custom/existing.xml").is_none());
    assert!(optional_member(&removed, "custom/_rels/existing.xml.rels").is_none());

    let reopened = SourceBackedPackage::from_vec(
        removed.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut inverse = SourceTopologyPlan::new();
    inverse
        .try_add_part(
            pack(EXISTING),
            "application/vnd.example.existing+xml",
            b"<existing/>".to_vec(),
        )
        .unwrap();
    inverse
        .try_add_internal_relationship(pack(DOCUMENT), "rIdDrop", DROP_REL, pack(EXISTING))
        .unwrap();
    inverse
        .try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(removed_document_relationships)),
            Some(Arc::new(original_document_relationships.clone())),
        )
        .unwrap();
    inverse
        .try_restore_relationships_xml(
            pack(EXISTING),
            None,
            Some(Arc::new(original_owned_relationships.clone())),
        )
        .unwrap();
    inverse
        .try_restore_content_types_xml(Arc::new(removed_types), Arc::new(original_types.clone()))
        .unwrap();
    let mut restored = Vec::new();
    reopened
        .write_topology_to_stream(&mut restored, inverse)
        .unwrap();
    assert_eq!(member(&restored, "[Content_Types].xml"), original_types);
    assert_eq!(
        member(&restored, "word/_rels/document.xml.rels"),
        original_document_relationships
    );
    assert_eq!(
        member(&restored, "custom/_rels/existing.xml.rels"),
        original_owned_relationships
    );
    assert_eq!(member(&restored, "custom/existing.xml"), b"<existing/>");
}

#[test]
fn raw_metadata_guards_reject_stale_and_mismatched_candidates_before_output() {
    let source = package_bytes(true, false);
    let original_types = member(&source, "[Content_Types].xml");
    let original_document_relationships = member(&source, "word/_rels/document.xml.rels");
    let removed_types = remove_once(
        &remove_once(&original_types, EXISTING_OVERRIDE),
        EXISTING_REL_OVERRIDE,
    );
    let removed_document_relationships =
        remove_once(&original_document_relationships, DOCUMENT_DROP_ENTRY);

    let mut stale = SourceTopologyPlan::new();
    stale
        .try_remove_relationship(pack(DOCUMENT), "rIdDrop")
        .unwrap();
    stale
        .try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(b"<stale/>".to_vec())),
            Some(Arc::new(removed_document_relationships.clone())),
        )
        .unwrap();
    let mut output = Vec::new();
    let error = SourceBackedPackage::from_vec(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, stale)
    .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
            | OpcError::InvalidRelationshipsManifest(_)
            | OpcError::SourceChanged { .. }
    ));
    assert!(output.is_empty());

    let mismatched_relationships = replace_once(
        &original_document_relationships,
        DOCUMENT_DROP_ENTRY,
        br#"  <r:Relationship Id="rIdOther" Type="urn:test:other" Target="../custom/existing.xml"/>
    "#,
    );
    let mut mismatch = SourceTopologyPlan::new();
    mismatch
        .try_remove_relationship(pack(DOCUMENT), "rIdDrop")
        .unwrap();
    mismatch
        .try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(original_document_relationships)),
            Some(Arc::new(mismatched_relationships)),
        )
        .unwrap();
    let mut output = Vec::new();
    let error = SourceBackedPackage::from_vec(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, mismatch)
    .unwrap_err();
    assert!(
        matches!(
            error,
            OpcError::SourceBackedOverlayUnavailable { .. }
                | OpcError::InvalidRelationship(_)
                | OpcError::InvalidRelationshipsManifest(_)
        ),
        "unexpected relationship graph mismatch error: {error:?}"
    );
    assert!(output.is_empty());

    let mut content_types_mismatch = SourceTopologyPlan::new();
    content_types_mismatch
        .try_remove_part(pack(EXISTING))
        .unwrap();
    content_types_mismatch
        .try_restore_content_types_xml(Arc::new(original_types.clone()), Arc::new(original_types))
        .unwrap();
    let mut output = Vec::new();
    let error = SourceBackedPackage::from_vec(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, content_types_mismatch)
    .unwrap_err();
    assert!(
        matches!(
            error,
            OpcError::SourceBackedOverlayUnavailable { .. }
                | OpcError::InvalidRelationship(_)
                | OpcError::InvalidContentTypesManifest(_)
        ),
        "unexpected content-types graph mismatch error: {error:?}"
    );
    assert!(output.is_empty());

    // Keep the helper inputs live in this test: a stale or graph-mismatched
    // raw image must never be accepted merely because its byte length fits.
    assert!(!removed_types.is_empty());
}

fn assert_rejected_without_output(source: &[u8], plan: SourceTopologyPlan) {
    let mut output = Vec::new();
    let error = SourceBackedPackage::from_vec(
        source.to_vec(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, plan)
    .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
            | OpcError::InvalidRelationship(_)
            | OpcError::InvalidRelationshipsManifest(_)
            | OpcError::InvalidContentTypesManifest(_)
    ));
    assert!(output.is_empty());
}

fn assert_relationship_manifest_rejected(source: &[u8], plan: SourceTopologyPlan) {
    let package = SourceBackedPackage::from_vec(
        source.to_vec(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(
        matches!(error, OpcError::InvalidRelationshipsManifest(_)),
        "unexpected hostile relationship source error: {error:?}"
    );
    assert!(output.is_empty());
}

fn assert_content_types_manifest_rejected(source: &[u8], plan: SourceTopologyPlan) {
    let package = SourceBackedPackage::from_vec(
        source.to_vec(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(
        matches!(error, OpcError::InvalidContentTypesManifest(_)),
        "unexpected hostile content-types source error: {error:?}"
    );
    assert!(output.is_empty());
}

#[test]
fn raw_metadata_rejects_hidden_relationship_and_content_type_markup() {
    let source = package_bytes(true, false);
    let original_types = member(&source, "[Content_Types].xml");
    let original_relationships = member(&source, "word/_rels/document.xml.rels");
    let removed_types = remove_once(
        &remove_once(&original_types, EXISTING_OVERRIDE),
        EXISTING_REL_OVERRIDE,
    );
    let removed_relationships = remove_once(&original_relationships, DOCUMENT_DROP_ENTRY);

    // The ordinary OPC readers intentionally skip unknown children and record
    // unknown attributes.  A raw restore cannot carry that skipped markup
    // through a typed graph comparison: it would make the candidate's lexical
    // image depend on information the graph does not represent.
    let relationship_hidden = replace_once(
        &removed_relationships,
        b"</r:Relationships>",
        br#"  <v:Hidden xmlns:v="urn:vendor"><r:Relationship Id="rIdHidden" Type="urn:hidden" Target="../custom/hidden.xml"/></v:Hidden>
</r:Relationships>"#,
    );
    let relationship_mce = replace_once(
        &removed_relationships,
        b"</r:Relationships>",
        br#"  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:vendor"><mc:Choice Requires="v"><r:Relationship Id="rIdHidden" Type="urn:hidden" Target="../custom/hidden.xml"/></mc:Choice><mc:Fallback/></mc:AlternateContent>
</r:Relationships>"#,
    );
    let relationship_unknown_attribute = replace_once(
        &removed_relationships,
        b"Id=\"rIdKeep\"",
        b"Id=\"rIdKeep\" v:hint=\"retained\" xmlns:v=\"urn:vendor\"",
    );

    let relationship_hostiles = [
        relationship_hidden,
        relationship_mce,
        relationship_unknown_attribute,
    ];
    for candidate in &relationship_hostiles {
        let mut plan = SourceTopologyPlan::new();
        plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
            .unwrap();
        plan.try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(original_relationships.clone())),
            Some(Arc::new(candidate.to_vec())),
        )
        .unwrap();
        assert_rejected_without_output(&source, plan);
    }
    for stale in &relationship_hostiles {
        let mut plan = SourceTopologyPlan::new();
        plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
            .unwrap();
        plan.try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(stale.to_vec())),
            Some(Arc::new(removed_relationships.clone())),
        )
        .unwrap();
        assert_rejected_without_output(&source, plan);
    }

    let types_hidden = replace_once(
        &removed_types,
        b"</Types>",
        br#"  <v:Hidden xmlns:v="urn:vendor"><Override PartName="/custom/hidden.xml" ContentType="application/hidden+xml"/></v:Hidden>
</Types>"#,
    );
    let types_mce = replace_once(
        &removed_types,
        b"</Types>",
        br#"  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:vendor"><mc:Choice Requires="v"><Override PartName="/custom/hidden.xml" ContentType="application/hidden+xml"/></mc:Choice><mc:Fallback/></mc:AlternateContent>
</Types>"#,
    );
    let types_unknown_attribute = replace_once(
        &removed_types,
        b"PartName=\"/word/document.xml\"",
        b"PartName=\"/word/document.xml\" v:hint=\"retained\" xmlns:v=\"urn:vendor\"",
    );

    let types_hostiles = [types_hidden, types_mce, types_unknown_attribute];
    for candidate in &types_hostiles {
        let mut plan = SourceTopologyPlan::new();
        plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
            .unwrap();
        plan.try_remove_part(pack(EXISTING)).unwrap();
        plan.try_restore_content_types_xml(
            Arc::new(original_types.clone()),
            Arc::new(candidate.to_vec()),
        )
        .unwrap();
        assert_rejected_without_output(&source, plan);
    }
    for stale in &types_hostiles {
        let mut plan = SourceTopologyPlan::new();
        plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
            .unwrap();
        plan.try_remove_part(pack(EXISTING)).unwrap();
        plan.try_restore_content_types_xml(
            Arc::new(stale.to_vec()),
            Arc::new(removed_types.clone()),
        )
        .unwrap();
        assert_rejected_without_output(&source, plan);
    }

    // Repeat the hostile images as the physical source members.  This keeps
    // the expected image byte-for-byte equal to the source, so the strict
    // raw parser rather than the stale-image guard must reject it.
    let relationship_source_hostiles = [
        replace_once(
            &original_relationships,
            b"</r:Relationships>",
            br#"  <v:Hidden xmlns:v="urn:vendor"><r:Relationship Id="rIdHidden" Type="urn:hidden" Target="../custom/hidden.xml"/></v:Hidden>
</r:Relationships>"#,
        ),
        replace_once(
            &original_relationships,
            b"</r:Relationships>",
            br#"  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:vendor"><mc:Choice Requires="v"><r:Relationship Id="rIdHidden" Type="urn:hidden" Target="../custom/hidden.xml"/></mc:Choice><mc:Fallback/></mc:AlternateContent>
</r:Relationships>"#,
        ),
        replace_once(
            &original_relationships,
            b"Id=\"rIdKeep\"",
            b"Id=\"rIdKeep\" v:hint=\"retained\" xmlns:v=\"urn:vendor\"",
        ),
    ];
    for hostile in &relationship_source_hostiles {
        let hostile_source = rewrite_member(&source, "word/_rels/document.xml.rels", hostile);
        let mut plan = SourceTopologyPlan::new();
        plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
            .unwrap();
        plan.try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(hostile.to_vec())),
            Some(Arc::new(removed_relationships.clone())),
        )
        .unwrap();
        assert_relationship_manifest_rejected(&hostile_source, plan);
    }

    let types_source_hostiles = [
        replace_once(
            &original_types,
            b"</Types>",
            br#"  <v:Hidden xmlns:v="urn:vendor"><Override PartName="/custom/hidden.xml" ContentType="application/hidden+xml"/></v:Hidden>
</Types>"#,
        ),
        replace_once(
            &original_types,
            b"</Types>",
            br#"  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:vendor"><mc:Choice Requires="v"><Override PartName="/custom/hidden.xml" ContentType="application/hidden+xml"/></mc:Choice><mc:Fallback/></mc:AlternateContent>
</Types>"#,
        ),
        replace_once(
            &original_types,
            b"PartName=\"/word/document.xml\"",
            b"PartName=\"/word/document.xml\" v:hint=\"retained\" xmlns:v=\"urn:vendor\"",
        ),
    ];
    for hostile in &types_source_hostiles {
        let hostile_source = rewrite_member(&source, "[Content_Types].xml", hostile);
        let mut plan = SourceTopologyPlan::new();
        plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
            .unwrap();
        plan.try_remove_part(pack(EXISTING)).unwrap();
        plan.try_restore_content_types_xml(
            Arc::new(hostile.to_vec()),
            Arc::new(removed_types.clone()),
        )
        .unwrap();
        assert_content_types_manifest_rejected(&hostile_source, plan);
    }
}

#[test]
fn raw_metadata_restores_absent_and_present_empty_relationship_members() {
    let absent_source = package_bytes(false, false);
    let empty = ORIGINAL_EXISTING_RELATIONSHIPS.to_vec();
    let mut add_member = SourceTopologyPlan::new();
    add_member
        .try_restore_relationships_xml(pack(EXISTING), None, Some(Arc::new(empty.clone())))
        .unwrap();
    let with_member = publish(&absent_source, add_member).unwrap();
    assert_eq!(
        member(&with_member, "custom/_rels/existing.xml.rels"),
        empty
    );

    let reopened = SourceBackedPackage::from_vec(
        with_member,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut remove_member = SourceTopologyPlan::new();
    remove_member
        .try_restore_relationships_xml(pack(EXISTING), Some(Arc::new(empty.clone())), None)
        .unwrap();
    let restored_absent = {
        let mut bytes = Vec::new();
        reopened
            .write_topology_to_stream(&mut bytes, remove_member)
            .unwrap();
        bytes
    };
    assert!(optional_member(&restored_absent, "custom/_rels/existing.xml.rels").is_none());

    let present_source = package_bytes(true, false);
    let mut remove_present = SourceTopologyPlan::new();
    remove_present
        .try_restore_relationships_xml(pack(EXISTING), Some(Arc::new(empty.clone())), None)
        .unwrap();
    let absent = publish(&present_source, remove_present).unwrap();
    assert!(optional_member(&absent, "custom/_rels/existing.xml.rels").is_none());
    let reopened = SourceBackedPackage::from_vec(
        absent,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut restore_present = SourceTopologyPlan::new();
    restore_present
        .try_restore_relationships_xml(pack(EXISTING), None, Some(Arc::new(empty.clone())))
        .unwrap();
    let round_trip = publish_from_package(reopened, restore_present);
    assert_eq!(member(&round_trip, "custom/_rels/existing.xml.rels"), empty);
}

#[test]
fn raw_metadata_managed_preparation_releases_reservations_after_cancellation() {
    let source = package_bytes(true, false);
    let original_types = member(&source, "[Content_Types].xml");
    let original_document_relationships = member(&source, "word/_rels/document.xml.rels");
    let original_owned_relationships = member(&source, "custom/_rels/existing.xml.rels");
    let removed_types = remove_once(
        &remove_once(&original_types, EXISTING_OVERRIDE),
        EXISTING_REL_OVERRIDE,
    );
    let removed_document_relationships =
        remove_once(&original_document_relationships, DOCUMENT_DROP_ENTRY);

    let (budget, cancellation_source, context) = managed_context(16 * 1024 * 1024);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let baseline = budget.used(Resource::Memory);
    let prepared = package
        .prepare_topology(removal_plan(
            &original_types,
            &removed_types,
            &original_document_relationships,
            &removed_document_relationships,
            Some(&original_owned_relationships),
        ))
        .unwrap();
    assert!(budget.used(Resource::Memory) > baseline);

    cancellation_source.cancel();
    assert!(matches!(
        prepared.with_candidate(|_| Ok::<(), OpcError>(())),
        Err(OpcError::Cancelled)
    ));
    drop(prepared);
    assert_eq!(budget.used(Resource::Memory), baseline);
}

#[test]
fn raw_metadata_restore_none_to_none_requests_are_bounded() {
    let mut plan = SourceTopologyPlan::new();
    for index in 0..4096 {
        let owner = PackURI::new(format!("/raw-restore-owner-{index}.xml")).unwrap();
        plan.try_restore_relationships_xml(owner, None, None)
            .unwrap();
    }
    let overflow = plan.try_restore_relationships_xml(
        PackURI::new("/raw-restore-owner-overflow.xml").unwrap(),
        None,
        None,
    );
    assert!(matches!(
        overflow,
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
}

fn publish_from_package(package: SourceBackedPackage, plan: SourceTopologyPlan) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    output
}

#[test]
fn signed_raw_metadata_noop_is_exact_but_changed_restore_is_refused() {
    let source = signed_package_bytes();
    let types = member(&source, "[Content_Types].xml");
    let relationships = member(&source, "word/_rels/document.xml.rels");

    let mut noop = SourceTopologyPlan::new();
    noop.try_restore_content_types_xml(Arc::new(types.clone()), Arc::new(types.clone()))
        .unwrap();
    noop.try_restore_relationships_xml(
        pack(DOCUMENT),
        Some(Arc::new(relationships.clone())),
        Some(Arc::new(relationships.clone())),
    )
    .unwrap();
    assert_eq!(publish(&source, noop).unwrap(), source);

    let changed_types = append_comment(&types, "changed");
    let mut changed = SourceTopologyPlan::new();
    changed
        .try_restore_content_types_xml(Arc::new(types), Arc::new(changed_types))
        .unwrap();
    let mut output = Vec::new();
    let error = SourceBackedPackage::from_vec(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, changed)
    .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));
    assert!(output.is_empty());
}

#[test]
fn raw_metadata_restore_obeys_member_and_aggregate_limits() {
    let source = package_bytes(true, false);
    let original_types = member(&source, "[Content_Types].xml");
    let original_relationships = member(&source, "word/_rels/document.xml.rels");
    let enlarged_types = append_comment(&original_types, "content-type growth");
    let enlarged_relationships = append_comment(&original_relationships, "relationship growth");

    let part_limits = SourceLimits::builder()
        .max_part_bytes(original_types.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_vec(
        source.clone(),
        part_limits,
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_restore_content_types_xml(
        Arc::new(original_types.clone()),
        Arc::new(enlarged_types.clone()),
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut output, plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            ..
        })
    ));
    assert!(output.is_empty());

    let member_limits = SourceLimits::builder()
        .max_content_types_bytes(original_types.len())
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_vec(
        source.clone(),
        member_limits,
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut content_types_plan = SourceTopologyPlan::new();
    content_types_plan
        .try_restore_content_types_xml(Arc::new(original_types.clone()), Arc::new(enlarged_types))
        .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut output, content_types_plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::ContentTypesBytes,
            ..
        })
    ));
    assert!(output.is_empty());

    let member_limits = SourceLimits::builder()
        .max_relationship_xml_bytes(original_relationships.len())
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_vec(
        source.clone(),
        member_limits,
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut relationship_plan = SourceTopologyPlan::new();
    relationship_plan
        .try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(original_relationships.clone())),
            Some(Arc::new(enlarged_relationships.clone())),
        )
        .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut output, relationship_plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::RelationshipXmlBytes,
            ..
        })
    ));
    assert!(output.is_empty());

    let total_relationship_bytes = [
        "_rels/.rels",
        "word/_rels/document.xml.rels",
        "custom/_rels/existing.xml.rels",
    ]
    .into_iter()
    .map(|name| member(&source, name).len())
    .sum::<usize>();
    let member_limits = SourceLimits::builder()
        .max_relationship_xml_bytes(enlarged_relationships.len())
        .unwrap()
        .max_total_relationship_xml_bytes(total_relationship_bytes)
        .unwrap()
        .build()
        .unwrap();
    let package =
        SourceBackedPackage::from_vec(source, member_limits, SourceCacheLimits::default(), None)
            .unwrap();
    let mut relationship_plan = SourceTopologyPlan::new();
    relationship_plan
        .try_restore_relationships_xml(
            pack(DOCUMENT),
            Some(Arc::new(original_relationships)),
            Some(Arc::new(enlarged_relationships)),
        )
        .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut output, relationship_plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::TotalRelationshipXmlBytes,
            ..
        })
    ));
    assert!(output.is_empty());
}

fn append_comment(xml: &[u8], text: &str) -> Vec<u8> {
    let close = xml
        .windows(2)
        .rposition(|window| window == b"</")
        .expect("fixture XML has a closing root");
    let mut output = Vec::with_capacity(xml.len() + text.len() + 16);
    output.extend_from_slice(&xml[..close]);
    output.extend_from_slice(b"<!-- ");
    output.extend_from_slice(text.as_bytes());
    output.extend_from_slice(b" -->\n");
    output.extend_from_slice(&xml[close..]);
    output
}

#[derive(Debug)]
struct DeniedSource {
    bytes: Arc<Vec<u8>>,
    denied: Range<u64>,
    deny_payload: AtomicU64,
    revision: AtomicU64,
    reads: AtomicU64,
}

impl DeniedSource {
    fn new(bytes: Vec<u8>, denied: Range<u64>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            denied,
            deny_payload: AtomicU64::new(0),
            revision: AtomicU64::new(0),
            reads: AtomicU64::new(0),
        }
    }

    fn enable_payload_denial(&self) {
        self.deny_payload.store(1, Ordering::SeqCst);
    }
}

impl ReadAt for DeniedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let end = offset.saturating_add(output.len() as u64);
        if self.deny_payload.load(Ordering::SeqCst) != 0
            && offset < self.denied.end
            && self.denied.start < end
        {
            return Err(io::Error::other("untouched payload was read"));
        }
        let offset = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5241_574d,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn stored_member_data_range(bytes: &[u8], wanted: &str) -> Range<u64> {
    let mut offset = 0usize;
    while offset + 30 <= bytes.len() && bytes[offset..offset + 4] == [0x50, 0x4b, 0x03, 0x04] {
        let name_len = u16::from_le_bytes([bytes[offset + 26], bytes[offset + 27]]) as usize;
        let extra_len = u16::from_le_bytes([bytes[offset + 28], bytes[offset + 29]]) as usize;
        let start = offset + 30 + name_len + extra_len;
        let size = u32::from_le_bytes([
            bytes[offset + 18],
            bytes[offset + 19],
            bytes[offset + 20],
            bytes[offset + 21],
        ]) as usize;
        let end = start + size;
        if std::str::from_utf8(&bytes[offset + 30..start - extra_len]).unwrap() == wanted {
            return start as u64..end as u64;
        }
        offset = end;
    }
    panic!("missing stored member {wanted}");
}

#[test]
fn raw_metadata_restore_keeps_untouched_payloads_lazy() {
    let source = package_bytes(true, true);
    let original_relationships = member(&source, "word/_rels/document.xml.rels");
    let removed_relationships = remove_once(&original_relationships, DOCUMENT_DROP_ENTRY);
    let denied = stored_member_data_range(&source, UNTOUCHED.trim_start_matches('/'));
    let source_handle = Arc::new(DeniedSource::new(source, denied));
    let package = SourceBackedPackage::from_read_at(
        source_handle.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    source_handle.enable_payload_denial();

    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
        .unwrap();
    plan.try_restore_relationships_xml(
        pack(DOCUMENT),
        Some(Arc::new(original_relationships)),
        Some(Arc::new(removed_relationships)),
    )
    .unwrap();
    package
        .with_prepared_topology(plan, |candidate| {
            assert_eq!(
                candidate
                    .relationships(&pack(DOCUMENT))?
                    .get("rIdKeep")
                    .unwrap()
                    .target_ref(),
                "../custom/untouched.bin"
            );
            Ok(())
        })
        .expect("candidate preparation must not read an untouched payload");
    assert!(source_handle.reads.load(Ordering::SeqCst) > 0);
}

#[test]
fn prepared_metadata_reads_exact_source_changed_and_omitted_images() {
    let source = package_bytes(true, false);
    let package = SourceBackedPackage::from_vec(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let unchanged = package
        .clone()
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    assert_eq!(
        unchanged
            .content_types_data_with_limit(usize::MAX)
            .unwrap()
            .as_bytes(),
        member(&source, "[Content_Types].xml")
    );
    for (owner, name) in [
        ("/", "_rels/.rels"),
        (DOCUMENT, "word/_rels/document.xml.rels"),
        (EXISTING, "custom/_rels/existing.xml.rels"),
    ] {
        assert_eq!(
            unchanged
                .relationships_data_for_with_limit(&pack(owner), usize::MAX)
                .unwrap()
                .unwrap()
                .as_bytes(),
            member(&source, name)
        );
    }
    assert!(
        unchanged
            .relationships_data_for_with_limit(&pack(UNTOUCHED), usize::MAX)
            .unwrap()
            .is_none()
    );
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
        .unwrap();
    plan.try_remove_part(pack(EXISTING)).unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    let rels = prepared
        .relationships_data_for_with_limit(&pack(DOCUMENT), usize::MAX)
        .unwrap()
        .unwrap();
    let types = prepared.content_types_data_with_limit(usize::MAX).unwrap();
    assert!(
        prepared
            .relationships_data_for_with_limit(&pack(EXISTING), usize::MAX)
            .unwrap()
            .is_none()
    );
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    drop(prepared);
    assert_eq!(
        rels.as_bytes(),
        member(&output, "word/_rels/document.xml.rels")
    );
    assert_eq!(types.as_bytes(), member(&output, "[Content_Types].xml"));
    assert_ne!(
        rels.as_bytes(),
        member(&source, "word/_rels/document.xml.rels")
    );
    assert_ne!(types.as_bytes(), member(&source, "[Content_Types].xml"));
}

#[test]
fn prepared_metadata_reads_appended_relationships_including_present_empty() {
    for empty in [false, true] {
        let source = package_bytes(false, false);
        let mut plan = SourceTopologyPlan::new();
        if empty {
            plan.try_restore_relationships_xml(
                pack(EXISTING),
                None,
                Some(Arc::new(ORIGINAL_EXISTING_RELATIONSHIPS.to_vec())),
            )
            .unwrap();
        } else {
            plan.try_add_external_relationship(
                pack(EXISTING),
                "rIdNew",
                "urn:test:new",
                "https://example.test/",
            )
            .unwrap();
        }
        let prepared = SourceBackedPackage::from_vec(
            source,
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
        .unwrap()
        .prepare_topology(plan)
        .unwrap();
        let data = prepared
            .relationships_data_for_with_limit(&pack(EXISTING), usize::MAX)
            .unwrap()
            .unwrap();
        let mut output = Vec::new();
        prepared.publish_to_stream(&mut output).unwrap();
        assert_eq!(
            data.as_bytes(),
            member(&output, "custom/_rels/existing.xml.rels")
        );
        if empty {
            assert_eq!(data.as_bytes(), ORIGINAL_EXISTING_RELATIONSHIPS);
        }
        assert!(matches!(
            prepared.relationships_data_for_with_limit(&pack(EXISTING), data.as_bytes().len() - 1),
            Err(OpcError::Limit {
                kind: SourceLimitKind::RelationshipXmlBytes,
                ..
            })
        ));
    }
}

#[test]
fn prepared_metadata_handles_retain_managed_admissions_and_check_cancellation() {
    let (budget, cancellation, context) = managed_context(16 * 1024 * 1024);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(package_bytes(false, false))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let baseline = budget.used(Resource::Memory);
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_external_relationship(
        pack(EXISTING),
        "rIdNew",
        "urn:test:new",
        "https://example.test/",
    )
    .unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    let generated = prepared
        .relationships_data_for_with_limit(&pack(EXISTING), usize::MAX)
        .unwrap()
        .unwrap();
    let source = prepared.content_types_data_with_limit(usize::MAX).unwrap();
    let retained = budget.used(Resource::Memory);
    assert!(retained > baseline);
    let clone = generated.clone();
    assert_eq!(budget.used(Resource::Memory), retained);
    assert!(matches!(
        prepared.content_types_data_with_limit(1),
        Err(OpcError::Limit {
            kind: SourceLimitKind::ContentTypesBytes,
            ..
        })
    ));
    cancellation.cancel();
    assert!(matches!(
        prepared.relationships_data_for_with_limit(&pack(EXISTING), usize::MAX),
        Err(OpcError::Cancelled)
    ));
    assert!(matches!(
        generated.check_execution(),
        Err(OpcError::Cancelled)
    ));
    assert!(!generated.as_bytes().is_empty());
    drop(prepared);
    drop(generated);
    drop(source);
    assert!(budget.used(Resource::Memory) > baseline);
    drop(clone);
    assert_eq!(budget.used(Resource::Memory), baseline);
}

#[test]
fn prepared_metadata_capture_does_not_read_unrelated_payloads() {
    let bytes = package_bytes(true, true);
    let range = stored_member_data_range(&bytes, UNTOUCHED.trim_start_matches('/'));
    let source = Arc::new(DeniedSource::new(bytes, range));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    source.enable_payload_denial();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
        .unwrap();
    plan.try_remove_part(pack(EXISTING)).unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    assert!(
        !prepared
            .content_types_data_with_limit(usize::MAX)
            .unwrap()
            .as_bytes()
            .is_empty()
    );
    assert!(
        !prepared
            .relationships_data_for_with_limit(&pack(DOCUMENT), usize::MAX)
            .unwrap()
            .unwrap()
            .as_bytes()
            .is_empty()
    );
    assert!(
        !prepared
            .relationships_data_for_with_limit(&pack("/"), usize::MAX)
            .unwrap()
            .unwrap()
            .as_bytes()
            .is_empty()
    );
}

#[test]
fn prepared_metadata_resolves_equivalent_owner_names() {
    let package = SourceBackedPackage::from_vec(
        package_bytes(true, false),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    for changed in [false, true] {
        let mut plan = SourceTopologyPlan::new();
        if changed {
            plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
                .unwrap();
        }
        let prepared = package.clone().prepare_topology(plan).unwrap();
        let canonical = prepared
            .relationships_data_for_with_limit(&pack(DOCUMENT), usize::MAX)
            .unwrap()
            .unwrap();
        let equivalent = prepared
            .relationships_data_for_with_limit(&pack("/WORD/DOCUMENT.XML"), usize::MAX)
            .unwrap()
            .unwrap();
        assert_eq!(canonical.as_bytes(), equivalent.as_bytes());
    }
}

#[test]
fn prepared_metadata_capture_rechecks_source_but_captured_bytes_remain_immutable() {
    for changed in [false, true] {
        let source = Arc::new(DeniedSource::new(package_bytes(true, false), 0..0));
        let package = SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let mut plan = SourceTopologyPlan::new();
        if changed {
            plan.try_remove_relationship(pack(DOCUMENT), "rIdDrop")
                .unwrap();
        }
        let prepared = package.prepare_topology(plan).unwrap();
        let captured = prepared
            .relationships_data_for_with_limit(&pack(DOCUMENT), usize::MAX)
            .unwrap()
            .unwrap();
        let before = captured.as_bytes().to_vec();
        source.revision.store(1, Ordering::SeqCst);
        assert!(matches!(
            prepared.relationships_data_for_with_limit(&pack(DOCUMENT), usize::MAX),
            Err(OpcError::SourceChanged { .. })
        ));
        assert!(matches!(
            prepared.content_types_data_with_limit(usize::MAX),
            Err(OpcError::SourceChanged { .. })
        ));
        assert_eq!(captured.as_bytes(), before);
        captured.check_execution().unwrap();
    }
}
