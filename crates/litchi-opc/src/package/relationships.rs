//! Source-bound eager relationship edits and exact restoration.

use super::{OpcPackage, RelationshipBinding, SourceRelationshipsXml};
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
/// Tokens use the default OPC relationship read limits. Part-member presence
/// is retained, including explicit empty members. The package root always has
/// a relationship member in authored publication.
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
        let maximum = max_output_bytes.min(ReadLimits::default().max_relationship_xml_bytes());
        let mut reader = NsReader::from_reader(self.bytes());
        let mut depth = 0usize;
        loop {
            let start = reader.buffer_position() as usize;
            let event = reader.read_event()?;
            let end = reader.buffer_position() as usize;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if depth == 1 && element.local_name().as_ref() == b"Relationship" {
                        for attribute in element.attributes().with_checks(true) {
                            let attribute =
                                attribute.map_err(|error| invalid(error.to_string()))?;
                            if attribute.key.as_ref() == b"Id"
                                && attribute
                                    .decoded_and_normalized_value(
                                        quick_xml::XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )?
                                    .as_ref()
                                    == id
                            {
                                let xml = self.xml.update_elements(
                                    &[OwnedElementUpdate {
                                        start_tag: start..end,
                                        edit: OwnedElementEdit::Remove,
                                    }],
                                    maximum,
                                )?;
                                crate::pkgreader::PackageReader::parse_owned_relationships(
                                    xml.bytes(),
                                    &self.owner,
                                )?;
                                return Ok(Self {
                                    owner: self.owner.clone(),
                                    xml,
                                    member_present: self.member_present,
                                });
                            }
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
        if self.bytes().len() > maximum {
            return Err(invalid("relationship XML exceeds output limit"));
        }
        Ok(self.clone())
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
        let (owner, relationships) = if owner.as_str() == "/" {
            (owner.clone(), self.rels())
        } else {
            let part = self.get_part(owner)?;
            (part.partname().clone(), part.rels())
        };
        let limits = ReadLimits::default();
        limits.check(
            ReadResource::RelationshipsPerPart,
            relationships.len() as u64,
            limits.max_relationships_per_part() as u64,
        )?;
        let binding = RelationshipBinding::from_relationships(relationships)?;
        let source = self
            .source_relationships_xml
            .get(&owner)
            .filter(|source| source.binding == binding);
        let bytes = if let Some(source) = source {
            Arc::clone(&source.bytes)
        } else {
            let length = crate::source_backed::canonical_relationship_xml_len(relationships)?;
            limits.check(
                ReadResource::RelationshipXmlBytes,
                length as u64,
                limits.max_relationship_xml_bytes() as u64,
            )?;
            Arc::new(relationships.try_to_xml_bytes()?)
        };
        crate::pkgreader::PackageReader::parse_owned_relationships(&bytes, &owner)?;
        let member_present = owner.as_str() == "/"
            || !relationships.is_empty()
            || self.source_relationships_member_present(&owner);
        let xml = OwnedXmlPart::capture(
            owner.rels_uri().map_err(OpcError::InvalidPackUri)?,
            crate::constants::content_type::OPC_RELATIONSHIPS.into(),
            bytes,
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
        if expected.owner != replacement.owner {
            return Err(invalid("relationship replacement has a different owner"));
        }
        let current = self.source_relationships(&expected.owner)?;
        if current != *expected {
            return Err(invalid("stale relationship replacement"));
        }
        if expected == replacement {
            return Ok(false);
        }
        if self.is_signed() || self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        let relationships = crate::pkgreader::PackageReader::parse_owned_relationships(
            replacement.bytes(),
            &replacement.owner,
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
}
