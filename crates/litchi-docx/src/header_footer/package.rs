//! WordprocessingML header/footer relationship orchestration.

use crate::error::{Error, Result};
use crate::parts::DocumentPart;
use crate::section::Section;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::packuri::PackURI;
use std::collections::HashSet;

use super::model::{Kind, Role, Story};

#[derive(Debug, Clone)]
struct Link {
    relationship_id: String,
    kind: Kind,
    target: PackURI,
}

/// Resolve header stories in the section-reference order of the main part.
pub(crate) fn headers(package: &OpcPackage, document: &DocumentPart<'_>) -> Result<Vec<Story>> {
    load(package, document, Role::Header)
}

/// Resolve footer stories in the section-reference order of the main part.
pub(crate) fn footers(package: &OpcPackage, document: &DocumentPart<'_>) -> Result<Vec<Story>> {
    load(package, document, Role::Footer)
}

/// Resolve picture watermark media referenced by header stories.
pub(crate) fn image_watermarks<'a>(
    package: &'a OpcPackage,
    document: &DocumentPart<'a>,
) -> Result<Vec<Image<'a>>> {
    let links = links(document, Role::Header)?;
    let mut images = Vec::new();
    for link in links {
        let header_part = package.get_part(&link.target)?;
        let story = Story::from_part(header_part, link.kind)?;
        if story.role() != Role::Header {
            return Err(Error::InvalidFormat(format!(
                "header relationship '{}' targets a footer story",
                link.relationship_id
            )));
        }
        for anchor in story.image_watermarks()? {
            let image_relationship = header_part
                .rels()
                .get(anchor.relationship_id())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "watermark image relationship '{}' is missing from {}",
                        anchor.relationship_id(),
                        link.target.as_str()
                    ))
                })?;
            if image_relationship.is_external() {
                return Err(Error::InvalidFormat(
                    "external watermark image relationship is rejected".into(),
                ));
            }
            let image_target = image_relationship.target_partname().map_err(|error| {
                Error::InvalidFormat(format!("invalid watermark image target: {error}"))
            })?;
            let image_part = package.get_part(&image_target)?;
            images.push(Image {
                source_header_name: link.target.as_str().to_owned(),
                relationship_id: anchor.relationship_id().to_owned(),
                part_name: image_target.as_str().to_owned(),
                content_type: image_part.content_type(),
                bytes: image_part.blob(),
            });
        }
    }
    Ok(images)
}

/// Borrowed media resolved from a header's local relationship graph.
pub(crate) struct Image<'a> {
    pub(crate) source_header_name: String,
    pub(crate) relationship_id: String,
    pub(crate) part_name: String,
    pub(crate) content_type: &'a str,
    pub(crate) bytes: &'a [u8],
}

fn load(package: &OpcPackage, document: &DocumentPart<'_>, role: Role) -> Result<Vec<Story>> {
    let links = links(document, role)?;
    let content_type = match role {
        Role::Header => content_type::WML_HEADER,
        Role::Footer => content_type::WML_FOOTER,
    };
    links
        .into_iter()
        .map(|link| {
            let part = package.get_part(&link.target)?;
            if part.content_type() != content_type {
                return Err(Error::ContentType {
                    expected: content_type.into(),
                    actual: part.content_type().into(),
                });
            }
            let story = Story::from_part(part, link.kind)?;
            if story.role() != role {
                return Err(Error::InvalidFormat(format!(
                    "{} relationship '{}' targets a {} story",
                    role_label(role),
                    link.relationship_id,
                    role_label(story.role())
                )));
            }
            Ok(story)
        })
        .collect()
}

fn links(document: &DocumentPart<'_>, role: Role) -> Result<Vec<Link>> {
    let relationship_name = match role {
        Role::Header => relationship_type::HEADER,
        Role::Footer => relationship_type::FOOTER,
    };
    let mut ordered_ids = Vec::new();
    crate::namespace::scan_word_element_ranges(
        document.xml_bytes(),
        &[b"sectPr".as_slice()],
        |_, start, length| {
            let start = usize::try_from(start)
                .map_err(|_| Error::InvalidFormat("section offset overflow".into()))?;
            let length = usize::try_from(length)
                .map_err(|_| Error::InvalidFormat("section length overflow".into()))?;
            let end = start
                .checked_add(length)
                .ok_or_else(|| Error::InvalidFormat("section range overflow".into()))?;
            let xml = document.xml_bytes().get(start..end).ok_or_else(|| {
                Error::InvalidFormat("section range is outside document XML".into())
            })?;
            let mut section = Section::from_xml_bytes(xml.to_vec())?;
            let references = if role == Role::Header {
                section.headers()?
            } else {
                section.footers()?
            };
            for reference in references {
                if !ordered_ids
                    .iter()
                    .any(|(relationship_id, _)| relationship_id == &reference.relationship_id)
                {
                    ordered_ids.push((reference.relationship_id, reference.kind));
                }
            }
            Ok(())
        },
    )?;

    let document_part = document.part();
    let mut links = Vec::new();
    let mut seen = HashSet::new();
    for (relationship_id, kind) in ordered_ids {
        let relationship = document_part.rels().get(&relationship_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{} reference '{}' has no package relationship",
                role_label(role),
                relationship_id
            ))
        })?;
        let link = link_from_relationship(relationship, role, relationship_name, kind)?;
        seen.insert(relationship_id);
        links.push(link);
    }

    // Preserve readable orphaned header/footer parts as a deterministic
    // low-level inventory. They have no section-local `w:type`, so the schema
    // default is the only safe semantic kind.
    let mut orphan_ids = document_part
        .rels()
        .iter()
        .filter(|relationship| {
            relationship.reltype() == relationship_name && !seen.contains(relationship.r_id())
        })
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<Vec<_>>();
    orphan_ids.sort_unstable();
    for relationship_id in orphan_ids {
        let relationship = document_part
            .rels()
            .get(&relationship_id)
            .ok_or_else(|| Error::InvalidFormat("relationship disappeared during scan".into()))?;
        links.push(link_from_relationship(
            relationship,
            role,
            relationship_name,
            Kind::Primary,
        )?);
    }
    Ok(links)
}

fn link_from_relationship(
    relationship: &litchi_opc::rel::Relationship,
    role: Role,
    expected_relationship: &str,
    kind: Kind,
) -> Result<Link> {
    if relationship.reltype() != expected_relationship {
        return Err(Error::InvalidFormat(format!(
            "{} reference uses relationship type '{}'",
            role_label(role),
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(Error::InvalidFormat(format!(
            "{} relationship '{}' must be internal",
            role_label(role),
            relationship.r_id()
        )));
    }
    let target = relationship.target_partname().map_err(|error| {
        Error::InvalidFormat(format!(
            "invalid {} relationship target '{}': {error}",
            role_label(role),
            relationship.r_id()
        ))
    })?;
    Ok(Link {
        relationship_id: relationship.r_id().to_owned(),
        kind,
        target,
    })
}

fn role_label(role: Role) -> &'static str {
    match role {
        Role::Header => "header",
        Role::Footer => "footer",
    }
}
