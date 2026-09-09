use std::collections::HashMap;
use std::sync::Arc;

use litchi_core::Position;
use litchi_drawingml::ink as shared;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{
    OpcPackage, PackURI, Part, PartData, PartReadSession, PartView, Relationships,
    SourceBackedPackage,
};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::{Form, scan};
use super::model::Payload;
use super::{Annotation, CONTENT_TYPE, Limits, Location, Snapshot};
use crate::package::story::{StoryDialect, StoryKind, capture};
use crate::{Error, Package, Result};

const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";

impl Package {
    /// Inventory active InkML annotations across all reachable Word stories.
    ///
    /// Shared targets are parsed once. Payloads and annotation snapshots retain
    /// immutable source storage after the package is dropped. No links are
    /// fetched, no handwriting is interpreted, and package state is unchanged.
    /// The projection exposes context/brush metadata and trace counts. Generic
    /// non-Ink content is left to its own owner; Word's product-specific
    /// compatibility for those content parts is not validated here.
    ///
    /// # Errors
    ///
    /// Returns an error for dirty facade state, invalid ownership, malformed
    /// Ink anchors/payloads, or exhausted resource budgets.
    pub fn ink(&self) -> Result<Snapshot> {
        self.ink_with_limits(Limits::default())
    }

    /// Inventory InkML using an explicit finite resource policy.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid limits, dirty state, malformed package/XML
    /// content, or exceeded limits. A failed inventory never changes the package.
    pub fn ink_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        let limits = limits.validate()?;
        self.ensure_story_opc_current("ink_with_limits")?;
        load(self.opc_package(), limits)
    }
}

pub(crate) fn load(package: &OpcPackage, limits: Limits) -> Result<Snapshot> {
    let owner = PackageRef::Owned(package);
    check_catalog(owner, limits)?;
    let stories = capture(package, limits.stories)?;
    load_stories(
        owner,
        limits,
        stories.dialect(),
        None,
        stories
            .stories()
            .iter()
            .map(|story| (story.part(), story.kind(), story.source())),
    )
}

pub(crate) fn load_source(package: &SourceBackedPackage, limits: Limits) -> Result<Snapshot> {
    let limits = limits.validate()?;
    let owner = PackageRef::Pinned(package);
    check_catalog(owner, limits)?;
    let mut session = package.read_session();
    let stories = crate::package::story::capture_source(package, limits.stories, &mut session)?;
    load_stories(
        owner,
        limits,
        stories.dialect(),
        Some(&mut session),
        stories
            .stories()
            .iter()
            .map(|story| (story.part(), story.kind(), story.source())),
    )
}

fn check_catalog(package: PackageRef<'_>, limits: Limits) -> Result<()> {
    check(
        "package parts",
        package.part_count(),
        limits.stories.max_package_parts,
    )?;
    let mut relationships = package.rels().len();
    check("relationships", relationships, limits.max_relationships)?;
    for part in package.parts() {
        relationships = relationships
            .checked_add(part.rels().len())
            .ok_or_else(|| exceeded("relationships", usize::MAX, limits.max_relationships))?;
        check("relationships", relationships, limits.max_relationships)?;
    }
    Ok(())
}

fn load_stories<'story>(
    package: PackageRef<'_>,
    limits: Limits,
    dialect: StoryDialect,
    mut session: Option<&mut PartReadSession<'_>>,
    stories: impl Iterator<Item = (&'story PackURI, StoryKind, &'story [u8])>,
) -> Result<Snapshot> {
    let mut payloads: HashMap<PackURI, Option<Arc<Payload>>> = HashMap::new();
    let mut annotations = Vec::new();
    let mut total_payload_bytes = 0usize;
    let mut anchor_count = 0usize;
    let mut roles = [0usize; 7];
    for (story_part, story_kind, story_source) in stories {
        package.check()?;
        let role = role_index(story_kind);
        let location = Location {
            kind: story_kind,
            position: Position::new(roles[role]),
        };
        roles[role] += 1;
        let anchors = scan(
            story_source,
            dialect,
            limits.max_xml_nodes,
            limits.max_xml_depth,
            limits.max_annotations.saturating_sub(anchor_count),
        )?;
        anchor_count = anchor_count
            .checked_add(anchors.len())
            .ok_or_else(|| exceeded("annotations", usize::MAX, limits.max_annotations))?;
        check("annotations", anchor_count, limits.max_annotations)?;
        let owner = package.part(story_part)?;
        for anchor in anchors {
            package.check()?;
            let relationship = owner.rels().get(&anchor.relationship_id).ok_or_else(|| {
                Error::InvalidRelationship("DOCX ink anchor relationship is missing".into())
            })?;
            let expected = if matches!(anchor.form, Form::Base) && dialect == StoryDialect::Strict {
                STRICT_CUSTOM_XML
            } else {
                rt::CUSTOM_XML
            };
            if relationship.is_external() || relationship.reltype() != expected {
                return Err(Error::InvalidRelationship(
                    "DOCX ink content part must use the owning story's internal customXml relationship".into(),
                ));
            }
            check(
                "target reference bytes",
                relationship.target_ref().len(),
                4096,
            )?;
            let target = relationship.target_partname()?;
            let part = package.part(&target)?;
            let is_declared_ink = part.content_type() == CONTENT_TYPE;
            if matches!(anchor.form, Form::GenericDrawing) && !is_declared_ink {
                continue;
            }
            if matches!(anchor.form, Form::Drawing) && !is_declared_ink {
                return Err(Error::ContentType {
                    expected: CONTENT_TYPE.into(),
                    actual: part.content_type().into(),
                });
            }
            if !matches!(anchor.form, Form::Base) && dialect == StoryDialect::Strict {
                return Err(Error::Invalid(
                    "Strict Word drawing Ink requires an explicit extension conformance policy"
                        .into(),
                ));
            }
            // Base contentPart also hosts non-Ink XML, which is preserved but
            // does not become an annotation. Word's text/xml variation applies
            // only to this base form.
            if !is_declared_ink && part.content_type() != "text/xml" {
                continue;
            }
            let document = if let Some(document) = payloads.get(part.partname()) {
                document.clone()
            } else {
                // Central-directory size is a preflight only. PartData checks
                // decoded integrity/length before exposing immutable bytes.
                let length = part.length()?;
                check("payload bytes", length, limits.max_payload_bytes)?;
                total_payload_bytes = total_payload_bytes.checked_add(length).ok_or_else(|| {
                    exceeded(
                        "total payload bytes",
                        usize::MAX,
                        limits.max_total_payload_bytes,
                    )
                })?;
                check(
                    "total payload bytes",
                    total_payload_bytes,
                    limits.max_total_payload_bytes,
                )?;
                let bytes = part.data(session.as_deref_mut())?;
                check(
                    "payload bytes",
                    bytes.as_bytes().len(),
                    limits.max_payload_bytes,
                )?;
                if bytes.as_bytes().len() != length {
                    return Err(Error::Invalid(
                        "DOCX InkML payload length changed during inventory".into(),
                    ));
                }
                let document = if has_ink_root(bytes.as_bytes())? {
                    // A declared Ink Content part cannot carry outgoing
                    // relationships to other parts under ECMA §15.2.4.
                    if !part.rels().is_empty() {
                        return Err(Error::InvalidRelationship(
                            "DOCX InkML payload has unsupported outgoing relationships".into(),
                        ));
                    }
                    Some(Arc::new(match bytes {
                        PayloadBytes::Owned(source) => Payload::Owned(shared::read_shared(source)?),
                        PayloadBytes::Pinned(source) => Payload::Pinned {
                            metadata: shared::read_metadata(source.as_bytes())?,
                            _source: source,
                        },
                    }))
                } else if is_declared_ink {
                    return Err(Error::Invalid(
                        "DOCX InkML payload has the wrong root namespace".into(),
                    ));
                } else {
                    None
                };
                payloads
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX ink distinct payloads",
                        source,
                    })?;
                payloads.insert(part.partname().clone(), document.clone());
                document
            };
            if let Some(document) = document {
                annotations
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX ink annotations",
                        source,
                    })?;
                annotations.push(Annotation { location, document });
            }
        }
    }
    Ok(Snapshot {
        annotations: Arc::new(annotations),
        distinct_payloads: payloads
            .values()
            .filter(|document| document.is_some())
            .count(),
    })
}

#[derive(Clone, Copy)]
enum PackageRef<'a> {
    Owned(&'a OpcPackage),
    Pinned(&'a SourceBackedPackage),
}

impl<'a> PackageRef<'a> {
    fn check(self) -> Result<()> {
        if let Self::Pinned(package) = self {
            package.check_execution()?;
            package.source_version()?;
        }
        Ok(())
    }

    fn part_count(self) -> usize {
        match self {
            Self::Owned(package) => package.part_count(),
            Self::Pinned(package) => package.iter_parts().count(),
        }
    }

    fn rels(self) -> &'a Relationships {
        match self {
            Self::Owned(package) => package.rels(),
            Self::Pinned(package) => package.rels(),
        }
    }

    fn parts(self) -> impl Iterator<Item = PartRef<'a>> {
        let (owned, pinned) = match self {
            Self::Owned(package) => (Some(package), None),
            Self::Pinned(package) => (None, Some(package)),
        };
        owned
            .into_iter()
            .flat_map(OpcPackage::iter_parts)
            .map(PartRef::Owned)
            .chain(
                pinned
                    .into_iter()
                    .flat_map(SourceBackedPackage::iter_parts)
                    .map(PartRef::Pinned),
            )
    }

    fn part(self, name: &PackURI) -> Result<PartRef<'a>> {
        match self {
            Self::Owned(package) => Ok(PartRef::Owned(package.get_part(name)?)),
            Self::Pinned(package) => Ok(PartRef::Pinned(package.part(name)?)),
        }
    }
}

#[derive(Clone, Copy)]
enum PartRef<'a> {
    Owned(&'a dyn Part),
    Pinned(PartView<'a>),
}

impl<'a> PartRef<'a> {
    fn partname(self) -> &'a PackURI {
        match self {
            Self::Owned(part) => part.partname(),
            Self::Pinned(part) => part.partname(),
        }
    }

    fn content_type(self) -> &'a str {
        match self {
            Self::Owned(part) => part.content_type(),
            Self::Pinned(part) => part.content_type(),
        }
    }

    fn rels(self) -> &'a Relationships {
        match self {
            Self::Owned(part) => part.rels(),
            Self::Pinned(part) => part.rels(),
        }
    }

    fn length(self) -> Result<usize> {
        match self {
            Self::Owned(part) => Ok(part.blob().len()),
            Self::Pinned(part) => {
                usize::try_from(part.declared_uncompressed_size()?).map_err(|_| {
                    Error::Invalid("DOCX InkML declared size exceeds address space".into())
                })
            },
        }
    }

    fn data(self, session: Option<&mut PartReadSession<'_>>) -> Result<PayloadBytes> {
        match self {
            Self::Owned(part) => {
                let source = part.blob_arc();
                let visible = part.blob();
                if !std::ptr::eq(source.as_slice(), visible) && source.as_slice() != visible {
                    return Err(Error::Invalid(
                        "DOCX InkML source storage is inconsistent".into(),
                    ));
                }
                Ok(PayloadBytes::Owned(source))
            },
            Self::Pinned(part) => Ok(PayloadBytes::Pinned(match session {
                Some(session) => session.read(part)?,
                None => part.data()?,
            })),
        }
    }
}

enum PayloadBytes {
    Owned(Arc<Vec<u8>>),
    Pinned(PartData),
}

impl PayloadBytes {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(bytes) => bytes.as_slice(),
            Self::Pinned(bytes) => bytes.as_bytes(),
        }
    }
}

pub(crate) fn has_ink_root(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    reader.resolver_mut().set_max_declarations_per_element(256);
    for _ in 0..256 {
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                return Ok(element.local_name().as_ref() == b"ink"
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri))
                        if uri == shared::INKML_NAMESPACE.as_bytes()));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            _ => {
                return Err(Error::Invalid(
                    "DOCX content-part XML has an invalid prolog".into(),
                ));
            },
        }
    }
    Err(exceeded("payload prolog events", 257, 256))
}

const fn role_index(kind: StoryKind) -> usize {
    match kind {
        StoryKind::Main => 0,
        StoryKind::Header => 1,
        StoryKind::Footer => 2,
        StoryKind::Footnotes => 3,
        StoryKind::Endnotes => 4,
        StoryKind::Comments => 5,
        StoryKind::Glossary => 6,
    }
}

fn exceeded(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InkLimit {
        resource,
        actual,
        maximum,
    }
}

fn check(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        Err(exceeded(resource, actual, maximum))
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::BlobPart;
    use litchi_opc::constants::content_type as ct;

    #[test]
    fn repeated_anchors_share_projection_and_original_payload_allocation() {
        let mut package = OpcPackage::new();
        let mut main = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.into(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:contentPart r:id="ink"/><w:contentPart r:id="ink"/></w:r></w:p></w:body></w:document>"#.to_vec(),
        );
        Part::rels_mut(&mut main).add_relationship(
            rt::CUSTOM_XML.into(),
            "../ink.xml".into(),
            "ink".into(),
            false,
        );
        package.add_part(Box::new(main));
        package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        let payload = Arc::new(
            br#"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:trace>1 2</i:trace></i:ink>"#
                .to_vec(),
        );
        package.add_part(Box::new(BlobPart::new_shared(
            PackURI::new("/ink.xml").unwrap(),
            CONTENT_TYPE.into(),
            Arc::clone(&payload),
        )));
        let snapshot = load(&package, Limits::default()).unwrap();
        assert_eq!(snapshot.annotations().len(), 2);
        let first = &snapshot.annotations()[0].document;
        let second = &snapshot.annotations()[1].document;
        assert!(Arc::ptr_eq(first, second));
        assert_eq!(first.source().as_ptr(), payload.as_ptr());
        let clone = snapshot.clone();
        assert!(Arc::ptr_eq(&snapshot.annotations, &clone.annotations));
        drop(package);
        drop(payload);
        assert_eq!(clone.annotations()[0].trace_count(), 1);
    }
}
