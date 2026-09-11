//! Source-bound eager relationship edits and exact restoration.

use super::{
    OpcPackage, RelationshipBinding, SourceRelationshipsXml,
    check_canonical_relationship_attribute_limits, check_canonical_relationship_structure,
    check_relationship_capture_limits,
};
use crate::{
    OpcError, OwnedElementEdit, OwnedElementUpdate, OwnedXmlPart, PackURI, ReadLimits,
    ReadResource, Result, TargetMode,
};
use quick_xml::{events::Event, reader::NsReader};
use std::collections::HashSet;
use std::fmt::Write as _;
use std::sync::Arc;

/// Immutable relationship XML bound to its package or part owner.
///
/// Capture with [`OpcPackage::source_relationships`]. Clones share source bytes;
/// an original token can restore its exact XML after a changed save/reopen.
/// `source_relationships_with_limits` applies the caller's bounded read
/// profile; the compatibility wrapper uses the default profile. Part-member
/// presence is retained, including explicit empty members. The package root
/// always has a relationship member in authored publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OwnedRelationships {
    owner: PackURI,
    xml: OwnedXmlPart,
    member_present: bool,
}

impl OwnedRelationships {
    /// Source part URI, or `/` for package relationships.
    #[must_use]
    pub fn owner(&self) -> &PackURI {
        &self.owner
    }

    /// Exact relationship XML; an absent part member has a canonical empty view.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.xml.bytes()
    }

    pub(crate) fn bytes_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml.bytes)
    }

    /// Whether publication includes this relationship member.
    #[must_use]
    pub fn member_present(&self) -> bool {
        self.member_present
    }

    /// Remove one edge while preserving surrounding XML byte for byte.
    ///
    /// A missing ID returns an unchanged token. The result must fit the caller
    /// limit and the default OPC relationship XML limit. Removal keeps an
    /// existing member, even when the collection becomes empty. This edits only
    /// the edge; the format owner must validate and edit its dependency closure.
    pub fn without_relationship(&self, id: &str, max_output_bytes: usize) -> Result<Self> {
        self.without_relationships(&[id], max_output_bytes)
    }

    /// Remove a bounded batch of edges in one source scan and one output
    /// allocation while preserving every unrelated source byte.
    ///
    /// Missing IDs and duplicate selectors are idempotent. The borrowed
    /// selectors are deduplicated before source ranges are collected, and the
    /// default relationship-count ceiling bounds the selector index. The
    /// result must fit both `max_output_bytes` and the default OPC relationship
    /// XML ceiling. Removal keeps an existing member, including when the
    /// collection becomes empty.
    pub fn without_relationships(&self, ids: &[&str], max_output_bytes: usize) -> Result<Self> {
        let maximum = max_output_bytes.min(ReadLimits::default().max_relationship_xml_bytes());
        if ids.is_empty() {
            if self.bytes().len() > maximum {
                return Err(invalid("relationship XML exceeds output limit"));
            }
            return Ok(self.clone());
        }

        let maximum_selectors = ReadLimits::default().max_relationships_per_part();
        let mut selected_ids: HashSet<&str> = HashSet::new();
        selected_ids
            .try_reserve(ids.len().min(maximum_selectors))
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship removal IDs",
                source,
            })?;
        for id in ids {
            if !selected_ids.contains(id) && selected_ids.len() >= maximum_selectors {
                return Err(invalid(
                    "relationship removal batch exceeds the package limit",
                ));
            }
            selected_ids.insert(*id);
        }

        let mut updates = Vec::new();
        updates
            .try_reserve_exact(selected_ids.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship removal updates",
                source,
            })?;
        let mut reader = NsReader::from_reader(self.bytes());
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        loop {
            let start = reader.buffer_position() as usize;
            let event = reader.read_event()?;
            let end = reader.buffer_position() as usize;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if depth == 1 && element.local_name().as_ref() == b"Relationship" {
                        let mut remove = false;
                        for attribute in element.attributes().with_checks(true) {
                            let attribute =
                                attribute.map_err(|error| invalid(error.to_string()))?;
                            if attribute.key.as_ref() == b"Id" {
                                let value = attribute.decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )?;
                                if selected_ids.contains(value.as_ref()) {
                                    remove = true;
                                    break;
                                }
                            }
                        }
                        if remove {
                            updates.push(OwnedElementUpdate {
                                start_tag: start..end,
                                edit: OwnedElementEdit::Remove,
                            });
                        }
                    }
                    if !empty {
                        depth += 1;
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("unbalanced relationship XML"))?;
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if updates.is_empty() {
            if self.bytes().len() > maximum {
                return Err(invalid("relationship XML exceeds output limit"));
            }
            return Ok(self.clone());
        }
        let xml = self.xml.update_elements(&updates, maximum)?;
        crate::pkgreader::PackageReader::parse_owned_relationships(xml.bytes(), &self.owner)?;
        Ok(Self {
            owner: self.owner.clone(),
            xml,
            member_present: self.member_present,
        })
    }

    /// Append one validated internal or external relationship while retaining
    /// every unrelated source byte in the owner `.rels` member.
    pub fn with_relationship(
        &self,
        reltype: &str,
        target: &str,
        id: &str,
        mode: TargetMode,
        max_output_bytes: usize,
    ) -> Result<Self> {
        self.with_relationships(&[(reltype, target, id, mode)], max_output_bytes)
    }

    /// Append a bounded batch in one source scan and one output allocation.
    /// This avoids repeatedly copying a growing owner `.rels` member when one
    /// host operation creates several fresh Ink/image targets.
    pub fn with_relationships(
        &self,
        additions: &[(&str, &str, &str, TargetMode)],
        max_output_bytes: usize,
    ) -> Result<Self> {
        let maximum = max_output_bytes.min(ReadLimits::default().max_relationship_xml_bytes());
        if additions.is_empty() {
            if self.bytes().len() > maximum {
                return Err(invalid("relationship XML exceeds output limit"));
            }
            return Ok(self.clone());
        }
        let parsed =
            crate::pkgreader::PackageReader::parse_owned_relationships(self.bytes(), &self.owner)?;
        let maximum_relationships = ReadLimits::default().max_relationships_per_part();
        let total_relationships = parsed
            .len()
            .checked_add(additions.len())
            .ok_or_else(|| invalid("relationship count overflows"))?;
        if total_relationships > maximum_relationships {
            return Err(invalid("relationship count exceeds the package limit"));
        }
        let mut ids = HashSet::new();
        ids.try_reserve(additions.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship batch IDs",
                source,
            })?;
        for (reltype, target, id, _) in additions {
            if id.is_empty() || reltype.is_empty() || target.is_empty() {
                return Err(invalid("relationship fields must not be empty"));
            }
            if parsed.get(id).is_some() || !ids.insert(id) {
                return Err(invalid("relationship ID already exists"));
            }
        }
        let mut reader = NsReader::from_reader(self.bytes());
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut root_tag = None;
        let mut root_name = None;
        let mut root_close = None;
        let mut root_empty = false;
        loop {
            let start = reader.buffer_position() as usize;
            let event = reader.read_event()?;
            let end = reader.buffer_position() as usize;
            match event {
                Event::Start(element) if depth == 0 => {
                    if element.local_name().as_ref() != b"Relationships" {
                        return Err(invalid("relationships root must be Relationships"));
                    }
                    root_tag = Some(start..end);
                    root_name = Some(element.name().as_ref().to_vec());
                    depth = 1;
                },
                Event::Empty(element) if depth == 0 => {
                    if element.local_name().as_ref() != b"Relationships" {
                        return Err(invalid("relationships root must be Relationships"));
                    }
                    root_tag = Some(start..end);
                    root_name = Some(element.name().as_ref().to_vec());
                    root_empty = true;
                    break;
                },
                Event::Start(_) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("relationship XML depth overflows"))?;
                },
                Event::End(_) => {
                    if depth == 1 {
                        root_close = Some(start..end);
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("unbalanced relationship XML"))?;
                    if depth == 0 {
                        break;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        let root_tag = root_tag.ok_or_else(|| invalid("relationships root is missing"))?;
        if !root_empty && root_close.is_none() {
            return Err(invalid("relationships root is unclosed"));
        }
        let root_name = root_name.ok_or_else(|| invalid("relationships root name is missing"))?;
        let relationship_name = root_name.iter().position(|byte| *byte == b':').map_or_else(
            || b"Relationship".to_vec(),
            |colon| {
                let mut name = root_name[..colon].to_vec();
                name.extend_from_slice(b":Relationship");
                name
            },
        );
        let mut fragment_bound = 0usize;
        for (reltype, target, id, mode) in additions {
            let fields = reltype
                .len()
                .checked_add(target.len())
                .and_then(|size| size.checked_add(id.len()))
                .and_then(|size| size.checked_mul(6))
                .and_then(|size| {
                    size.checked_add(if *mode == TargetMode::External { 32 } else { 4 })
                })
                .and_then(|size| size.checked_add(relationship_name.len().saturating_add(37)))
                .ok_or_else(|| invalid("relationship batch XML size overflows"))?;
            fragment_bound = fragment_bound
                .checked_add(fields)
                .ok_or_else(|| invalid("relationship batch XML size overflows"))?;
        }
        if self
            .bytes()
            .len()
            .checked_add(fragment_bound)
            .is_none_or(|size| size > maximum)
        {
            return Err(invalid("relationship XML exceeds output limit"));
        }
        let mut fragment = Vec::new();
        fragment
            .try_reserve_exact(fragment_bound)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship batch XML",
                source,
            })?;
        for (reltype, target, id, mode) in additions {
            let mut element = String::new();
            write!(
                element,
                "<{} Id=\"{}\" Type=\"{}\" Target=\"{}\"",
                String::from_utf8_lossy(&relationship_name),
                litchi_core::xml::escape_xml(id),
                litchi_core::xml::escape_xml(reltype),
                litchi_core::xml::escape_xml(target),
            )
            .map_err(|_| invalid("relationship formatting failed"))?;
            if *mode == TargetMode::External {
                element.push_str(" TargetMode=\"External\"");
            }
            element.push_str("/>");
            fragment.extend_from_slice(element.as_bytes());
        }
        let is_empty = root_empty;
        let (range, suffix) = if let Some(close) = root_close.as_ref() {
            (close.start..close.start, Vec::new())
        } else {
            let mut suffix = Vec::new();
            suffix
                .try_reserve(root_name.len().saturating_add(4))
                .map_err(|source| OpcError::Allocation {
                    resource: "OPC relationship root close",
                    source,
                })?;
            suffix.extend_from_slice(b">");
            suffix.extend_from_slice(b"</");
            suffix.extend_from_slice(&root_name);
            suffix.push(b'>');
            (root_tag.end.saturating_sub(2)..root_tag.end, suffix)
        };
        let size = self
            .bytes()
            .len()
            .checked_sub(range.len())
            .and_then(|size| size.checked_add(fragment.len()))
            .and_then(|size| size.checked_add(suffix.len()))
            .ok_or_else(|| invalid("relationship XML output size overflows"))?;
        if size > maximum {
            return Err(invalid("relationship XML exceeds output limit"));
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship batch output",
                source,
            })?;
        bytes.extend_from_slice(&self.bytes()[..range.start]);
        if is_empty {
            bytes.extend_from_slice(&suffix[..1]);
            bytes.extend_from_slice(&fragment);
            bytes.extend_from_slice(&suffix[1..]);
        } else {
            bytes.extend_from_slice(&fragment);
        }
        bytes.extend_from_slice(&self.bytes()[range.end..]);
        let xml = OwnedXmlPart::capture(
            self.xml.name.clone(),
            self.xml.content_type.clone(),
            Arc::new(bytes),
        )?;
        crate::pkgreader::PackageReader::parse_owned_relationships(xml.bytes(), &self.owner)?;
        Ok(Self {
            owner: self.owner.clone(),
            xml,
            member_present: true,
        })
    }
}

impl OpcPackage {
    /// Capture source-bound relationship XML without exposing mutable provenance.
    /// Newly authored or changed collections receive a validated canonical view.
    pub fn source_relationships(&self, owner: &PackURI) -> Result<OwnedRelationships> {
        self.source_relationships_with_limits(owner, ReadLimits::default())
    }

    /// Capture source-bound relationship XML under an explicit bounded read
    /// policy. Relationship counts and field lengths are checked before
    /// semantic binding clones; retained XML and canonical output are checked
    /// before parsing or allocation.
    pub fn source_relationships_with_limits(
        &self,
        owner: &PackURI,
        limits: ReadLimits,
    ) -> Result<OwnedRelationships> {
        let (owner_ref, relationships) = if owner.as_str() == "/" {
            (owner, self.rels())
        } else {
            let part = self.get_part(owner)?;
            (part.partname(), part.rels())
        };
        OwnedXmlPart::check_derived_capture_member_name(owner_ref, limits)?;
        let relationship_uri = owner_ref.rels_uri().map_err(OpcError::InvalidPackUri)?;
        check_relationship_capture_limits(relationships, limits)?;
        let source = self
            .source_relationships_xml
            .get(owner_ref)
            .filter(|source| source.binding.matches(relationships));
        let bytes = if let Some(source) = source {
            limits.check(
                ReadResource::RelationshipXmlBytes,
                source.bytes.len() as u64,
                limits.max_relationship_xml_bytes() as u64,
            )?;
            limits.check(
                ReadResource::TotalRelationshipXmlBytes,
                source.bytes.len() as u64,
                limits.max_total_relationship_xml_bytes() as u64,
            )?;
            limits.check(
                ReadResource::PartBytes,
                source.bytes.len() as u64,
                limits.max_part_bytes(),
            )?;
            OwnedXmlPart::check_capture_size(&relationship_uri, source.bytes.len(), limits)?;
            Arc::clone(&source.bytes)
        } else {
            let length = crate::source_backed::canonical_relationship_xml_len(relationships)?;
            limits.check(
                ReadResource::RelationshipXmlBytes,
                length as u64,
                limits.max_relationship_xml_bytes() as u64,
            )?;
            limits.check(
                ReadResource::PartBytes,
                length as u64,
                limits.max_part_bytes(),
            )?;
            check_canonical_relationship_attribute_limits(relationships, limits)?;
            check_canonical_relationship_structure(relationships, length, limits)?;
            OwnedXmlPart::check_capture_size(&relationship_uri, length, limits)?;
            Arc::new(relationships.try_to_xml_bytes()?)
        };
        crate::pkgreader::PackageReader::parse_owned_relationships_with_limits(
            &bytes, owner_ref, limits,
        )?;
        let member_present = owner_ref.as_str() == "/"
            || !relationships.is_empty()
            || self.source_relationships_member_present(owner_ref);
        let owner = owner_ref.clone();
        let xml = OwnedXmlPart::capture_with_limits(
            relationship_uri,
            crate::constants::content_type::OPC_RELATIONSHIPS.into(),
            bytes,
            limits,
        )?;
        Ok(OwnedRelationships {
            owner,
            xml,
            member_present,
        })
    }

    /// Replace an exact expected relationship snapshot with a source-derived token.
    ///
    /// Owner, XML bytes and member presence must still match. Stale snapshots,
    /// different owners and signed changes fail before mutation. The caller is
    /// responsible for the format-level dependency closure of changed edges.
    pub fn try_replace_relationships(
        &mut self,
        expected: &OwnedRelationships,
        replacement: &OwnedRelationships,
    ) -> Result<bool> {
        self.try_replace_relationships_with_limits(expected, replacement, ReadLimits::default())
    }

    /// Replace an exact expected relationship snapshot under an explicit
    /// bounded read policy.
    ///
    /// The current owner XML is checked before replacement parsing or package
    /// mutation. A changed signed package is refused before retaining parsed
    /// replacement fields, and the replacement is parsed under the same
    /// policy used for the current source check.
    pub fn try_replace_relationships_with_limits(
        &mut self,
        expected: &OwnedRelationships,
        replacement: &OwnedRelationships,
        limits: ReadLimits,
    ) -> Result<bool> {
        if expected.owner != replacement.owner {
            return Err(invalid("relationship replacement has a different owner"));
        }
        let current = self.source_relationships_with_limits(&expected.owner, limits)?;
        if current != *expected {
            return Err(invalid("stale relationship replacement"));
        }
        if expected == replacement {
            return Ok(false);
        }
        if self.is_signed() || self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        let relationships = crate::pkgreader::PackageReader::parse_owned_relationships_with_limits(
            replacement.bytes(),
            &replacement.owner,
            limits,
        )?;
        if !replacement.member_present && !relationships.is_empty() {
            return Err(invalid("absent relationship member has edges"));
        }
        let binding = RelationshipBinding::from_relationships(&relationships)?;
        self.source_relationships_xml
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "owned relationship provenance",
                source,
            })?;
        let source = Arc::new(SourceRelationshipsXml {
            bytes: Arc::clone(&replacement.xml.bytes),
            binding,
        });
        if replacement.owner.as_str() == "/" {
            self.revoke_exact_source();
            self.signature_graph_tracked = true;
            self.rels = relationships;
        } else {
            *self.get_part_mut(&replacement.owner)?.rels_mut() = relationships;
        }
        if replacement.member_present {
            self.source_relationships_xml
                .insert(replacement.owner.clone(), source);
        } else {
            self.source_relationships_xml.remove(&replacement.owner);
        }
        Ok(true)
    }
}

fn invalid(message: impl Into<String>) -> OpcError {
    OpcError::InvalidRelationship(message.into())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure"
    )]
    use super::*;
    use crate::{BlobPart, PackageWriter};

    const XML: &[u8] = br#"<?xml version='1.0'?>
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>
 <!-- before --> <r:Relationship TargetMode='External' Target='https://example.test/a' Type='urn:test' Id='rId1'></r:Relationship>
 <!-- middle --> <r:Relationship Id='rId2' Type='urn:test' Target='https://example.test/b' TargetMode='External'/>
 <!-- after -->
</r:Relationships>"#;

    fn source(xml: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        for (name, bytes) in [
            ("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#.as_slice()),
            ("_rels/.rels", xml),
            ("custom/item.bin", b"opaque".as_slice()),
            ("custom/_rels/item.bin.rels", xml),
        ] { writer.write_deflated(name, bytes).unwrap(); }
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn relationship_removal_round_trip_preserves_context_and_restores_source_xml() {
        let source = source(XML);
        for owned in [false, true] {
            for owner in ["/", "/custom/item.bin"] {
                let owner = PackURI::new(owner).unwrap();
                let mut package = if owned {
                    OpcPackage::from_vec(source.clone()).unwrap()
                } else {
                    OpcPackage::from_bytes(&source).unwrap()
                };
                let before = package.source_relationships(&owner).unwrap();
                assert_eq!(before.bytes(), XML);
                assert!(Arc::ptr_eq(&before.xml.bytes, &before.clone().xml.bytes));
                let after = before.without_relationship("rId1", XML.len()).unwrap();
                assert_eq!(after.bytes(), String::from_utf8(XML.to_vec()).unwrap().replace("<r:Relationship TargetMode='External' Target='https://example.test/a' Type='urn:test' Id='rId1'></r:Relationship>", "").as_bytes());
                assert!(
                    before
                        .without_relationship("rId1", after.bytes().len() - 1)
                        .is_err()
                );
                assert_eq!(
                    before
                        .without_relationship("rId1", after.bytes().len())
                        .unwrap(),
                    after
                );
                assert_eq!(
                    before.without_relationship("missing", XML.len()).unwrap(),
                    before
                );
                assert!(!package.try_replace_relationships(&before, &before).unwrap());
                assert!(package.try_replace_relationships(&before, &after).unwrap());
                let output = PackageWriter::to_bytes(&package).unwrap();
                let archive = soapberry_zip::office::ArchiveReader::new(&output).unwrap();
                assert_eq!(
                    archive
                        .read(owner.rels_uri().unwrap().membername())
                        .unwrap(),
                    after.bytes()
                );
                assert_eq!(archive.read("custom/item.bin").unwrap(), b"opaque");
                let mut reopened = OpcPackage::from_vec(output).unwrap();
                assert_eq!(reopened.source_relationships(&owner).unwrap(), after);
                let unchanged = PackageWriter::to_bytes(&reopened).unwrap();
                assert!(reopened.try_replace_relationships(&before, &after).is_err());
                assert_eq!(PackageWriter::to_bytes(&reopened).unwrap(), unchanged);
                assert!(reopened.try_replace_relationships(&after, &before).unwrap());
                let restored = PackageWriter::to_bytes(&reopened).unwrap();
                let archive = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
                assert_eq!(
                    archive
                        .read(owner.rels_uri().unwrap().membername())
                        .unwrap(),
                    XML
                );
                assert_eq!(archive.read("custom/item.bin").unwrap(), b"opaque");
            }
        }
    }

    #[test]
    fn relationship_batch_append_copies_source_once_and_keeps_comments() {
        let bytes = source(XML);
        let mut package = OpcPackage::from_vec(bytes).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before
            .with_relationships(
                &[
                    ("urn:ink", "custom/item.bin", "rId7", TargetMode::Internal),
                    (
                        "urn:image",
                        "custom/image.png",
                        "rId8",
                        TargetMode::External,
                    ),
                ],
                1024 * 1024,
            )
            .unwrap();
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("<!-- before -->")
        );
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("Id=\"rId7\"")
        );
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("Id=\"rId8\"")
        );
        package.try_replace_relationships(&before, &after).unwrap();
        assert_eq!(package.source_relationships(&owner).unwrap(), after);
    }

    #[test]
    fn relationship_batch_removal_preserves_pi_comments_namespaces_and_paired_tags() {
        let xml = br#"<?xml version='1.0'?>
<?before?>
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>
 <!-- before --> <r:Relationship TargetMode='External' Target='https://example.test/a' Type='urn:test' Id='rId1'></r:Relationship>
 <?middle?> <r:Relationship Id='rId2' Type='urn:test' Target='https://example.test/b' TargetMode='External'/>
 <!-- after -->
</r:Relationships>
<?after?>
"#;
        let package = OpcPackage::from_bytes(&source(xml)).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before
            .without_relationships(&["rId1", "missing", "rId2", "rId2"], xml.len())
            .unwrap();

        assert!(after.member_present());
        assert!(
            !after
                .bytes()
                .windows(b"rId1".len())
                .any(|window| window == b"rId1")
        );
        assert!(
            !after
                .bytes()
                .windows(b"rId2".len())
                .any(|window| window == b"rId2")
        );
        for marker in [
            b"<?before?>".as_slice(),
            b"<?middle?>",
            b"<?after?>",
            b"<!-- before -->",
            b"<!-- after -->",
        ] {
            assert!(
                after
                    .bytes()
                    .windows(marker.len())
                    .any(|window| window == marker)
            );
        }
        let sequential = before
            .without_relationship("rId1", xml.len())
            .unwrap()
            .without_relationship("rId2", xml.len())
            .unwrap();
        assert_eq!(after.bytes(), sequential.bytes());
        assert!(
            before
                .without_relationships(&["rId1", "rId2"], after.bytes().len() - 1)
                .is_err()
        );
    }

    #[test]
    fn relationship_batch_removal_missing_duplicate_and_empty_batches_are_bounded_noops() {
        let bytes = source(XML);
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();

        let missing = before
            .without_relationships(&["missing", "missing"], XML.len())
            .unwrap();
        assert_eq!(missing, before);
        assert!(Arc::ptr_eq(&before.xml.bytes, &missing.xml.bytes));
        assert!(
            before
                .without_relationships(&["missing"], before.bytes().len() - 1)
                .is_err()
        );
        assert!(
            before
                .without_relationships(&[], before.bytes().len() - 1)
                .is_err()
        );
        let exact_empty = before
            .without_relationships(&[], before.bytes().len())
            .unwrap();
        assert_eq!(exact_empty, before);
        assert!(Arc::ptr_eq(&before.xml.bytes, &exact_empty.xml.bytes));
    }

    #[test]
    fn relationship_batch_removal_keeps_absent_member_presence_on_noop() {
        let mut package = OpcPackage::new();
        let owner = PackURI::new("/custom/no-rels.bin").unwrap();
        package.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/octet-stream".into(),
            vec![],
        )));
        let before = package.source_relationships(&owner).unwrap();
        assert!(!before.member_present());
        let after = before
            .without_relationships(&["missing"], before.bytes().len())
            .unwrap();
        assert_eq!(after, before);
        assert!(!after.member_present());
        assert!(Arc::ptr_eq(&before.xml.bytes, &after.xml.bytes));
    }

    #[test]
    fn relationship_batch_expands_empty_root_without_dropping_trailing_members() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/><!--tail--><?keep?>
"#;
        let archive = source(xml);
        let package = OpcPackage::from_bytes(&archive).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before
            .with_relationships(
                &[(
                    crate::constants::relationship_type::OFFICE_DOCUMENT,
                    "word/document.xml",
                    "rId1",
                    TargetMode::Internal,
                )],
                4096,
            )
            .unwrap();
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .ends_with("<!--tail--><?keep?>\n")
        );
    }

    #[test]
    fn empty_relationship_batch_enforces_requested_output_limit() {
        let bytes = source(XML);
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();

        assert!(before.with_relationships(&[], 0).is_err());
        assert!(
            before
                .with_relationships(&[], before.bytes().len() - 1)
                .is_err()
        );
        assert_eq!(
            before
                .with_relationships(&[], before.bytes().len())
                .unwrap(),
            before
        );
    }

    #[test]
    fn relationship_tokens_reject_lexically_stale_sources_and_different_owners() {
        let bytes = source(XML);
        let original = OpcPackage::from_bytes(&bytes).unwrap();
        let owner = PackURI::new("/custom/item.bin").unwrap();
        let before = original.source_relationships(&owner).unwrap();
        let after = before.without_relationship("rId2", XML.len()).unwrap();
        let changed = String::from_utf8(XML.to_vec())
            .unwrap()
            .replace("<!-- middle -->", "<!-- other -->");
        let changed_bytes = source(changed.as_bytes());
        let mut package = OpcPackage::from_vec(changed_bytes).unwrap();
        let saved = PackageWriter::to_bytes(&package).unwrap();
        assert!(package.try_replace_relationships(&before, &after).is_err());
        let root = package
            .source_relationships(&PackURI::new("/").unwrap())
            .unwrap();
        assert!(package.try_replace_relationships(&before, &root).is_err());
        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), saved);
    }

    #[test]
    fn relationship_tokens_preserve_empty_member_presence_and_refuse_signed_changes() {
        let bytes = source(XML);
        let mut package = OpcPackage::from_vec(bytes).unwrap();
        let owner = PackURI::new("/custom/item.bin").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let empty = before
            .without_relationship("rId1", XML.len())
            .unwrap()
            .without_relationship("rId2", XML.len())
            .unwrap();
        assert!(package.try_replace_relationships(&before, &empty).unwrap());
        let output = PackageWriter::to_bytes(&package).unwrap();
        assert_eq!(
            soapberry_zip::office::ArchiveReader::new(&output)
                .unwrap()
                .read(owner.rels_uri().unwrap().membername())
                .unwrap(),
            empty.bytes()
        );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            "application/vnd.openxmlformats-package.digital-signature-origin".into(),
            vec![],
        )));
        assert!(!package.try_replace_relationships(&empty, &empty).unwrap());
        assert!(matches!(
            package.try_replace_relationships(&empty, &before),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert_eq!(package.source_relationships(&owner).unwrap(), empty);
    }

    #[test]
    fn bounded_relationship_replacement_checks_token_before_mutation() {
        let mut package = OpcPackage::from_vec(source(XML)).unwrap();
        let owner = PackURI::new("/custom/item.bin").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before.without_relationship("rId1", XML.len()).unwrap();
        assert!(package.try_replace_relationships(&before, &after).unwrap());
        let current = package.source_relationships(&owner).unwrap();
        let before_bytes = PackageWriter::to_bytes(&package).unwrap();

        let under = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .max_total_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.try_replace_relationships_with_limits(&current, &before, under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                actual,
                maximum,
            }) if actual == XML.len() as u64 && maximum == (XML.len() - 1) as u64
        ));
        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), before_bytes);

        let exact = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_total_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_part_bytes(XML.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .try_replace_relationships_with_limits(&current, &before, exact)
                .unwrap()
        );
        assert_eq!(
            package
                .source_relationships_with_limits(&owner, exact)
                .unwrap(),
            before
        );
    }

    #[test]
    fn bounded_relationship_capture_checks_retained_and_authored_quotas() {
        let archive = source(XML);
        let package = OpcPackage::from_bytes(&archive).unwrap();
        let root = PackURI::new("/").unwrap();
        let exact = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_part_bytes(XML.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            package
                .source_relationships_with_limits(&root, exact)
                .unwrap()
                .bytes(),
            XML
        );

        let root_member = root.rels_uri().unwrap().membername().len();
        let root_name_exact = ReadLimits::builder()
            .max_archive_member_name_bytes(root_member as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .source_relationships_with_limits(&root, root_name_exact)
                .is_ok()
        );
        let root_name_under = ReadLimits::builder()
            .max_archive_member_name_bytes((root_member - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_relationships_with_limits(&root, root_name_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveMemberNameBytes,
                actual,
                maximum,
            }) if actual == root_member as u64 && maximum == (root_member - 1) as u64
        ));

        let custom_owner = PackURI::new("/custom/item.bin").unwrap();
        let custom_member = custom_owner.rels_uri().unwrap().membername().len();
        let custom_name_exact = ReadLimits::builder()
            .max_archive_member_name_bytes(custom_member as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .source_relationships_with_limits(&custom_owner, custom_name_exact)
                .is_ok()
        );
        let custom_name_under = ReadLimits::builder()
            .max_archive_member_name_bytes((custom_member - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_relationships_with_limits(&custom_owner, custom_name_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveMemberNameBytes,
                actual,
                maximum,
            }) if actual == custom_member as u64 && maximum == (custom_member - 1) as u64
        ));

        let under = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_relationships_with_limits(&root, under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                ..
            })
        ));

        let over = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len() + 1)
            .unwrap()
            .max_part_bytes((XML.len() + 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .source_relationships_with_limits(&root, over)
                .is_ok()
        );

        let mut authored = OpcPackage::new();
        authored
            .rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "https://example.test/a&b".to_owned(),
                "rId1".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        let authored_token = authored.source_relationships(&root).unwrap();
        let authored_exact = ReadLimits::builder()
            .max_relationship_xml_bytes(authored_token.bytes().len())
            .unwrap()
            .max_part_bytes(authored_token.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            authored
                .source_relationships_with_limits(&root, authored_exact)
                .unwrap()
                .bytes(),
            authored_token.bytes()
        );

        let event_under = ReadLimits::builder()
            .max_xml_events(4)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_relationships_with_limits(&root, event_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlEvents,
                actual: 5,
                maximum: 4,
            })
        ));
        let depth_under = ReadLimits::builder()
            .max_xml_depth(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_relationships_with_limits(&root, depth_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlDepth,
                actual: 2,
                maximum: 1,
            })
        ));
        let aggregate_bytes_under = ReadLimits::builder()
            .max_total_relationship_xml_bytes(authored_token.bytes().len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_relationships_with_limits(&root, aggregate_bytes_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlBytes,
                ..
            })
        ));

        let mut too_many = authored.clone();
        too_many
            .rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "https://example.test/c".to_owned(),
                "rId2".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        let count_limit = ReadLimits::builder()
            .max_relationships_per_part(1)
            .unwrap()
            .max_total_relationships(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            too_many.source_relationships_with_limits(&root, count_limit),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipsPerPart,
                actual: 2,
                maximum: 1,
            })
        ));
    }
}
