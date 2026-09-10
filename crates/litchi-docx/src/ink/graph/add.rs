//! Bounded, source-checked additions to the Ink OPC closure.
//!
//! New resources are always allocated under a fresh package part name.  The
//! host editor owns the story/XML splice; this module only adds the target
//! parts, their owner-scoped relationship IDs, and the corresponding manifest
//! overrides.  The delta retains exact source tokens so a later inverse can
//! restore lexical relationship and content-type XML after save/reopen.

use std::sync::Arc;

use litchi_drawingml::ink as shared;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, Part,
    Relationships, TargetMode,
};

use super::super::Limits;
use super::{
    add_retained, canonical_owner, capture_relationships, check_payload,
    cmp_ascii_case_insensitive, is_image_content_type, is_ink_content_type, owner_relationships,
    preflight_graph, replace_content_types, same_relationships,
};
use crate::format::ImageFormat;
use crate::{Error, Result};

const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
const MAX_RELATIONSHIP_XML_BYTES: usize = 8 * 1024 * 1024;

/// Relationship family selected from the owning Word package.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RelationshipDialect {
    Transitional,
    Strict,
}

impl RelationshipDialect {
    const fn custom_xml(self) -> &'static str {
        match self {
            Self::Transitional => rt::CUSTOM_XML,
            Self::Strict => STRICT_CUSTOM_XML,
        }
    }

    const fn image(self) -> &'static str {
        match self {
            Self::Transitional => rt::IMAGE,
            Self::Strict => rt::STRICT_IMAGE,
        }
    }
}

/// A bounded target staged by a host edit.
///
/// The kind and content type stay crate-private so the public Ink API cannot
/// manufacture arbitrary OPC identities.  The payload allocation is shared
/// directly with the inserted part.
#[derive(Clone, Debug)]
pub(crate) struct NewTarget {
    payload: Arc<Vec<u8>>,
    kind: NewTargetKind,
}

#[derive(Clone, Debug)]
enum NewTargetKind {
    Ink { content_type: String },
    Image { content_type: String },
}

impl NewTarget {
    /// Stage an InkML/XML target. The package graph validates the profile and
    /// payload before changing the candidate.
    pub(crate) fn ink(payload: Arc<Vec<u8>>, content_type: impl Into<String>) -> Self {
        Self {
            payload,
            kind: NewTargetKind::Ink {
                content_type: content_type.into(),
            },
        }
    }

    /// Stage a PNG/JPEG fallback image target. Magic bytes and MIME agreement
    /// are checked by [`AddDelta::insert_many`].
    pub(crate) fn image(payload: Arc<Vec<u8>>, content_type: impl Into<String>) -> Self {
        Self {
            payload,
            kind: NewTargetKind::Image {
                content_type: content_type.into(),
            },
        }
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.payload.as_slice()
    }
}

/// Identity returned to the host editor for one newly added target.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct AddedTarget {
    part: PackURI,
    relationship_id: String,
}

impl AddedTarget {
    pub(crate) fn part(&self) -> &PackURI {
        &self.part
    }

    pub(crate) fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
}

#[derive(Clone)]
struct AddedTargetGuard {
    target: AddedTarget,
    content_type: String,
    payload: Arc<Vec<u8>>,
    relationships: Vec<super::RelationshipBinding>,
    relationship_token: OwnedRelationships,
}

#[derive(Clone)]
struct AddDeltaInner {
    owner: PackURI,
    owner_before: OwnedRelationships,
    owner_after: OwnedRelationships,
    content_types_before: OwnedContentTypes,
    content_types_after: OwnedContentTypes,
    targets: Vec<AddedTargetGuard>,
}

/// Exact, reversible graph closure for a batch of newly created targets.
#[derive(Clone)]
pub(crate) struct AddDelta {
    inner: Arc<AddDeltaInner>,
    reversed: bool,
}

impl AddDelta {
    /// Add a bounded batch in one graph census and one manifest publication.
    pub(crate) fn insert_many(
        package: &mut OpcPackage,
        owner: &PackURI,
        inputs: &[NewTarget],
        limits: Limits,
    ) -> Result<(Vec<AddedTarget>, Self)> {
        let limits = limits.validate()?;
        let relationship_count = preflight_graph(package, limits)?;
        let owner_name = canonical_owner(package, owner)?;
        let owner_relationships = owner_relationships(package, &owner_name)?;
        let owner_before = package.source_relationships(&owner_name)?;
        let content_types_before = package.source_content_types()?;
        if inputs.is_empty() {
            return Ok((
                Vec::new(),
                Self {
                    inner: Arc::new(AddDeltaInner {
                        owner: owner_name,
                        owner_before: owner_before.clone(),
                        owner_after: owner_before,
                        content_types_before: content_types_before.clone(),
                        content_types_after: content_types_before,
                        targets: Vec::new(),
                    }),
                    reversed: false,
                },
            ));
        }
        if inputs.len() > limits.max_relationships
            || owner_relationships
                .len()
                .checked_add(inputs.len())
                .is_none_or(|count| count > limits.max_relationships)
            || relationship_count
                .checked_add(inputs.len())
                .is_none_or(|count| count > limits.max_relationships)
        {
            return Err(Error::InkLimit {
                resource: "Ink owner relationships",
                actual: owner_relationships.len().saturating_add(inputs.len()),
                maximum: limits.max_relationships,
            });
        }
        let mut total_payload = 0usize;
        for input in inputs {
            check_payload(input.payload.len(), limits, &mut total_payload)?;
            validate_new_target(input)?;
        }

        let dialect = relationship_dialect(package, owner_relationships);
        let mut next_ink_name = 1usize;
        let mut next_png_name = 1usize;
        let mut next_jpeg_name = 1usize;
        let mut next_relationship_id = 1usize;
        let mut targets = Vec::new();
        targets
            .try_reserve_exact(inputs.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink added targets",
                source,
            })?;
        let mut override_specs = Vec::new();
        override_specs
            .try_reserve_exact(inputs.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink added content-type overrides",
                source,
            })?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(inputs.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink added parts",
                source,
            })?;
        let mut relationship_specs = Vec::new();
        relationship_specs
            .try_reserve_exact(inputs.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink added relationship specs",
                source,
            })?;

        for input in inputs {
            let (content_type, relationship_type, extension) = match &input.kind {
                NewTargetKind::Ink { content_type } => {
                    (content_type.as_str(), dialect.custom_xml(), "xml")
                },
                NewTargetKind::Image { content_type } => (
                    content_type.as_str(),
                    dialect.image(),
                    image_extension(content_type)?,
                ),
            };
            let next_name = match &input.kind {
                NewTargetKind::Ink { .. } => &mut next_ink_name,
                NewTargetKind::Image { content_type }
                    if content_type.eq_ignore_ascii_case("image/png") =>
                {
                    &mut next_png_name
                },
                NewTargetKind::Image { .. } => &mut next_jpeg_name,
            };
            let part = fresh_part_name(package, next_name, extension, input, limits)?;
            let relationship_id =
                fresh_relationship_id(owner_relationships, &mut next_relationship_id, limits)?;
            let target_ref = part.relative_ref(owner_name.base_uri());
            relationship_specs.push((relationship_type, target_ref, relationship_id.clone()));
            override_specs.push((part.clone(), content_type.to_owned()));
            let part_blob: Box<dyn Part + Send + Sync> = Box::new(BlobPart::new_shared(
                part.clone(),
                content_type.to_owned(),
                Arc::clone(&input.payload),
            ));
            parts.push(part_blob);
            targets.push(AddedTarget {
                part,
                relationship_id,
            });
        }
        let mut relationship_additions = Vec::new();
        relationship_additions
            .try_reserve_exact(inputs.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink relationship additions",
                source,
            })?;
        for (reltype, target, id) in &relationship_specs {
            relationship_additions.push((
                *reltype,
                target.as_str(),
                id.as_str(),
                TargetMode::Internal,
            ));
        }
        let owner_after =
            owner_before.with_relationships(&relationship_additions, MAX_RELATIONSHIP_XML_BYTES)?;
        let mut overrides = Vec::new();
        overrides
            .try_reserve_exact(inputs.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink content-type additions",
                source,
            })?;
        for (part, content_type) in &override_specs {
            overrides.push((part, content_type.as_str()));
        }
        let content_types_after = content_types_before
            .with_part_overrides(&overrides, limits.stories.max_topology_bytes)?;
        package.try_add_parts_with_source_tokens(
            content_types_before.bytes(),
            &content_types_after,
            &owner_before,
            &owner_after,
            parts,
        )?;

        let mut guards = Vec::new();
        guards
            .try_reserve_exact(targets.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink added target guards",
                source,
            })?;
        for target in &targets {
            let part = package.get_part(&target.part).map_err(Error::Opc)?;
            guards.push(AddedTargetGuard {
                target: target.clone(),
                content_type: part.content_type().to_owned(),
                payload: part.blob_arc(),
                relationships: capture_relationships(part.rels())?,
                relationship_token: package.source_relationships(&target.part)?,
            });
        }
        let owner_after_token = package.source_relationships(&owner_name)?;
        debug_assert_eq!(owner_after_token, owner_after);
        Ok((
            targets,
            Self {
                inner: Arc::new(AddDeltaInner {
                    owner: owner_name,
                    owner_before,
                    owner_after: owner_after_token,
                    content_types_before,
                    content_types_after,
                    targets: guards,
                }),
                reversed: false,
            },
        ))
    }

    /// Apply this direction to an exact source-matching package.
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        let (expected_content, replacement_content) = if self.reversed {
            (
                &self.inner.content_types_after,
                &self.inner.content_types_before,
            )
        } else {
            (
                &self.inner.content_types_before,
                &self.inner.content_types_after,
            )
        };
        if package.source_content_types()?.bytes() != expected_content.bytes() {
            return Err(Error::Invalid(
                "DOCX Ink add content-types source is stale".into(),
            ));
        }
        let (expected_owner, replacement_owner) = if self.reversed {
            (&self.inner.owner_after, &self.inner.owner_before)
        } else {
            (&self.inner.owner_before, &self.inner.owner_after)
        };
        if package.source_relationships(expected_owner.owner())? != *expected_owner {
            return Err(Error::Invalid(
                "DOCX Ink add relationship source is stale".into(),
            ));
        }
        self.validate_targets(package)?;
        if self.reversed {
            ensure_no_foreign_incoming_batch(package, &self.inner.targets, &self.inner.owner)?;
            package.try_replace_relationships(expected_owner, replacement_owner)?;
            for target in &self.inner.targets {
                if !package.remove_part(&target.target.part) {
                    return Err(Error::Invalid(
                        "DOCX Ink added target disappeared during inverse".into(),
                    ));
                }
            }
            replace_content_types(package, replacement_content)?;
        } else {
            let mut parts = Vec::new();
            parts
                .try_reserve_exact(self.inner.targets.len())
                .map_err(|source| Error::Allocation {
                    resource: "DOCX Ink added replay parts",
                    source,
                })?;
            for target in &self.inner.targets {
                let part: Box<dyn Part + Send + Sync> = Box::new(BlobPart::new_shared(
                    target.target.part.clone(),
                    target.content_type.clone(),
                    Arc::clone(&target.payload),
                ));
                parts.push(part);
            }
            package.try_add_parts_with_source_tokens(
                expected_content.bytes(),
                replacement_content,
                expected_owner,
                replacement_owner,
                parts,
            )?;
        }
        if package.source_content_types()?.bytes() != replacement_content.bytes()
            || package.source_relationships(replacement_owner.owner())? != *replacement_owner
        {
            return Err(Error::Invalid("DOCX Ink add result is stale".into()));
        }
        Ok(())
    }

    /// Return the exact inverse direction while sharing immutable guards.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            inner: Arc::clone(&self.inner),
            reversed: !self.reversed,
        }
    }

    /// Conservative retained bytes for an edit staging budget.
    pub(crate) fn retained_bytes(&self) -> Result<usize> {
        let mut total = 0usize;
        for bytes in [
            self.inner.owner_before.bytes().len(),
            self.inner.owner_after.bytes().len(),
            self.inner.content_types_before.bytes().len(),
            self.inner.content_types_after.bytes().len(),
        ] {
            add_retained(&mut total, bytes)?;
        }
        for target in &self.inner.targets {
            add_retained(&mut total, target.target.part.as_str().len())?;
            add_retained(&mut total, target.target.relationship_id.len())?;
            add_retained(&mut total, target.content_type.len())?;
            add_retained(&mut total, target.payload.len())?;
            add_retained(&mut total, target.relationship_token.bytes().len())?;
            for relationship in &target.relationships {
                add_retained(&mut total, relationship.id.len())?;
                add_retained(&mut total, relationship.reltype.len())?;
                add_retained(&mut total, relationship.target.len())?;
            }
        }
        Ok(total)
    }

    fn validate_targets(&self, package: &OpcPackage) -> Result<()> {
        for target in &self.inner.targets {
            let part = package.get_part(&target.target.part);
            if self.reversed {
                let part = match part {
                    Ok(part) => part,
                    Err(OpcError::PartNotFound(_)) => {
                        return Err(Error::Invalid(
                            "DOCX Ink add inverse target is missing".into(),
                        ));
                    },
                    Err(error) => return Err(Error::Opc(error)),
                };
                if part.content_type() != target.content_type
                    || (part.blob() != target.payload.as_slice()
                        && !std::ptr::eq(part.blob(), target.payload.as_slice()))
                    || !same_relationships(part.rels(), &target.relationships)
                    || package.source_relationships(part.partname())? != target.relationship_token
                {
                    return Err(Error::Invalid("DOCX Ink add target is stale".into()));
                }
            } else {
                match part {
                    Ok(_) => {
                        return Err(Error::Invalid("DOCX Ink add target already exists".into()));
                    },
                    Err(OpcError::PartNotFound(_)) => {},
                    Err(error) => return Err(Error::Opc(error)),
                }
            }
        }
        Ok(())
    }
}

fn validate_new_target(input: &NewTarget) -> Result<()> {
    match &input.kind {
        NewTargetKind::Ink { content_type } => {
            if content_type.len() > 4096
                || !is_ink_content_type(content_type)
                || litchi_opc::ContentType::new(content_type.clone()).is_err()
            {
                return Err(unsupported_add(
                    "new Ink target has an unsupported content type",
                ));
            }
            if !super::super::package::has_ink_root(input.payload.as_slice())? {
                return Err(unsupported_add(
                    "new Ink target is not a namespace-bound InkML part",
                ));
            }
            let projection =
                super::super::package::validate_content_part(input.payload.as_slice())?;
            let _ = shared::read_metadata_with_source_spans(
                input.payload.as_slice(),
                projection.contexts(),
                projection.traces(),
                projection.brush_properties(),
                projection.links(),
            )?;
        },
        NewTargetKind::Image { content_type } => {
            if content_type.len() > 4096
                || !is_image_content_type(content_type)
                || litchi_opc::ContentType::new(content_type.clone()).is_err()
                || image_extension(content_type).is_err()
            {
                return Err(unsupported_add(
                    "new fallback image has an unsupported content type",
                ));
            }
            let format =
                ImageFormat::detect_from_bytes(input.payload.as_slice()).ok_or_else(|| {
                    unsupported_add("new fallback image has an unsupported or malformed signature")
                })?;
            if (content_type.eq_ignore_ascii_case("image/png") && format != ImageFormat::Png)
                || (content_type.eq_ignore_ascii_case("image/jpeg") && format != ImageFormat::Jpeg)
            {
                return Err(unsupported_add(
                    "fallback image MIME does not match its signature",
                ));
            }
        },
    }
    Ok(())
}

fn image_extension(content_type: &str) -> Result<&'static str> {
    if content_type.eq_ignore_ascii_case("image/png") {
        Ok("png")
    } else if content_type.eq_ignore_ascii_case("image/jpeg") {
        Ok("jpg")
    } else {
        Err(unsupported_add(
            "only PNG and JPEG fallback images are supported",
        ))
    }
}

fn fresh_part_name(
    package: &OpcPackage,
    next: &mut usize,
    extension: &str,
    input: &NewTarget,
    limits: Limits,
) -> Result<PackURI> {
    let prefix = match &input.kind {
        NewTargetKind::Ink { .. } => "/word/ink/ink",
        NewTargetKind::Image { .. } => "/word/media/ink",
    };
    while *next <= limits.stories.max_package_parts {
        let index = *next;
        *next = (*next).saturating_add(1);
        let candidate = PackURI::new(format!("{prefix}{index}.{extension}")).map_err(Error::Uri)?;
        match package.validate_new_part_name(&candidate) {
            Ok(()) => return Ok(candidate),
            Err(OpcError::DuplicatePartName(_) | OpcError::EquivalentPartNames { .. }) => {},
            Err(error) => {
                if let OpcError::DerivedPartNames { existing, .. } = &error {
                    // A part occupying the resource directory (or one of its
                    // ancestors) conflicts with every possible suffix. Refuse
                    // immediately instead of exhausting the numeric namespace.
                    let directory = match &input.kind {
                        NewTargetKind::Ink { .. } => "/word/ink",
                        NewTargetKind::Image { .. } => "/word/media",
                    };
                    let blocks_directory = directory.eq_ignore_ascii_case(existing)
                        || directory
                            .get(..existing.len())
                            .is_some_and(|prefix| prefix.eq_ignore_ascii_case(existing))
                            && directory.as_bytes().get(existing.len()) == Some(&b'/');
                    if !blocks_directory {
                        continue;
                    }
                }
                return Err(Error::Opc(error));
            },
        }
    }
    Err(Error::InkLimit {
        resource: "fresh Ink target names",
        actual: limits.stories.max_package_parts,
        maximum: limits.stories.max_package_parts,
    })
}

fn fresh_relationship_id(
    relationships: &Relationships,
    next: &mut usize,
    limits: Limits,
) -> Result<String> {
    while *next <= limits.max_relationships {
        let index = *next;
        *next = (*next).saturating_add(1);
        let candidate = format!("rId{index}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::InkLimit {
        resource: "fresh Ink relationship IDs",
        actual: limits.max_relationships,
        maximum: limits.max_relationships,
    })
}

fn relationship_dialect(package: &OpcPackage, owner: &Relationships) -> RelationshipDialect {
    let mut strict = false;
    let mut transitional = false;
    for relationship in owner.iter() {
        if relationship.reltype() == STRICT_CUSTOM_XML {
            strict = true;
        } else if relationship.reltype() == rt::CUSTOM_XML {
            transitional = true;
        }
    }
    for relationship in package.rels().iter() {
        if relationship.reltype() == rt::STRICT_OFFICE_DOCUMENT {
            strict = true;
        } else if relationship.reltype() == rt::OFFICE_DOCUMENT {
            transitional = true;
        }
    }
    if strict && !transitional {
        RelationshipDialect::Strict
    } else {
        RelationshipDialect::Transitional
    }
}

fn ensure_no_foreign_incoming_batch(
    package: &OpcPackage,
    targets: &[AddedTargetGuard],
    owner: &PackURI,
) -> Result<()> {
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink inverse target census",
            source,
        })?;
    for target in targets {
        selected.push((&target.target.part, target.target.relationship_id.as_str()));
    }
    selected.sort_unstable_by(|left, right| {
        cmp_ascii_case_insensitive(left.0.as_str(), right.0.as_str())
            .then_with(|| left.1.cmp(right.1))
    });

    let check = |source: &PackURI, relationships: &Relationships| -> Result<()> {
        for relationship in relationships.iter() {
            if relationship.is_external() {
                continue;
            }
            let requested = relationship.target_partname()?;
            let Ok(index) = selected.binary_search_by(|(target, _)| {
                cmp_ascii_case_insensitive(target.as_str(), requested.as_str())
            }) else {
                continue;
            };
            let (target, relationship_id) = selected[index];
            if !source.is_equivalent_to(owner) || relationship.r_id() != relationship_id {
                return Err(unsupported_add(
                    "new Ink target gained a foreign incoming relationship",
                ));
            }
            debug_assert!(target.is_equivalent_to(&requested));
        }
        Ok(())
    };
    let root = PackURI::new("/").map_err(Error::Uri)?;
    check(&root, package.rels())?;
    for part in package.iter_parts() {
        check(part.partname(), part.rels())?;
    }
    Ok(())
}

fn unsupported_add(reason: &'static str) -> Error {
    Error::UnsafeEdit {
        format: "DOCX",
        operation: "add_ink",
        reason,
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "bounded graph tests use panic-on-failure assertions"
    )]

    use super::*;
    use crate::ink::CONTENT_TYPE;
    use litchi_opc::PackageWriter;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};

    const INK: &[u8] = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML"><i:definitions><i:context xml:id="ctx0"/><i:brush xml:id="br0"/></i:definitions><i:trace contextRef="#ctx0" brushRef="#br0">1 2, 3 4</i:trace></i:ink>"##;

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
        )));
        package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        package
    }

    #[test]
    fn batch_add_and_inverse_restore_the_exact_authored_tokens() {
        let mut package = package();
        let before_content = package.source_content_types().unwrap().bytes().to_vec();
        let before_owner = package
            .source_relationships(&PackURI::new("/word/document.xml").unwrap())
            .unwrap();
        let input = NewTarget::ink(Arc::new(INK.to_vec()), CONTENT_TYPE);
        let (targets, delta) = AddDelta::insert_many(
            &mut package,
            &PackURI::new("/word/document.xml").unwrap(),
            &[input],
            Limits::default(),
        )
        .unwrap();
        assert_eq!(targets.len(), 1);
        assert_eq!(targets[0].part().as_str(), "/word/ink/ink1.xml");
        assert_eq!(targets[0].relationship_id(), "rId1");
        assert!(
            package
                .source_content_types()
                .unwrap()
                .bytes()
                .windows(b"/word/ink/ink1.xml".len())
                .any(|window| window == b"/word/ink/ink1.xml")
        );
        delta.inverse().apply(&mut package).unwrap();
        assert!(
            package
                .get_part(&PackURI::new("/word/ink/ink1.xml").unwrap())
                .is_err()
        );
        assert_eq!(
            package.source_content_types().unwrap().bytes(),
            before_content
        );
        assert_eq!(
            package
                .source_relationships(&PackURI::new("/word/document.xml").unwrap())
                .unwrap(),
            before_owner
        );
    }

    #[test]
    fn inverse_after_save_and_reopen_restores_the_source_tokens() {
        let mut package = package();
        let before_content = package.source_content_types().unwrap().bytes().to_vec();
        let before_owner = package
            .source_relationships(&PackURI::new("/word/document.xml").unwrap())
            .unwrap();
        let (_, delta) = AddDelta::insert_many(
            &mut package,
            &PackURI::new("/word/document.xml").unwrap(),
            &[NewTarget::ink(Arc::new(INK.to_vec()), CONTENT_TYPE)],
            Limits::default(),
        )
        .unwrap();
        let saved = PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_bytes(&saved).unwrap();
        delta.inverse().apply(&mut reopened).unwrap();
        assert_eq!(
            reopened.source_content_types().unwrap().bytes(),
            before_content
        );
        assert_eq!(
            reopened
                .source_relationships(&PackURI::new("/word/document.xml").unwrap())
                .unwrap(),
            before_owner
        );
    }

    #[test]
    fn forward_rejects_a_stale_owner_without_adding_parts() {
        let mut package = package();
        let before = package.clone();
        let (_, delta) = AddDelta::insert_many(
            &mut package,
            &PackURI::new("/word/document.xml").unwrap(),
            &[NewTarget::ink(Arc::new(INK.to_vec()), CONTENT_TYPE)],
            Limits::default(),
        )
        .unwrap();
        let mut stale = before;
        stale
            .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::CUSTOM_XML.to_owned(),
                "https://example.test/foreign".to_owned(),
                "foreign".to_owned(),
                true,
            );
        assert!(delta.apply(&mut stale).is_err());
        assert!(
            stale
                .get_part(&PackURI::new("/word/ink/ink1.xml").unwrap())
                .is_err()
        );
    }

    #[test]
    fn batch_add_skips_case_insensitive_part_and_relationship_collisions() {
        let mut package = package();
        let existing = PackURI::new("/word/ink/ink1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            existing.clone(),
            CONTENT_TYPE.to_owned(),
            INK.to_vec(),
        )));
        package
            .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::CUSTOM_XML.to_owned(),
                "ink/ink1.xml".to_owned(),
                "rId1".to_owned(),
                false,
            );
        let (targets, _) = AddDelta::insert_many(
            &mut package,
            &PackURI::new("/word/document.xml").unwrap(),
            &[NewTarget::ink(Arc::new(INK.to_vec()), CONTENT_TYPE)],
            Limits::default(),
        )
        .unwrap();
        assert_eq!(targets[0].part().as_str(), "/word/ink/ink2.xml");
        assert_eq!(targets[0].relationship_id(), "rId2");
    }

    #[test]
    fn occupied_resource_directory_refuses_without_exhausting_suffixes() {
        let mut package = package();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/WORD/INK").unwrap(),
            "application/octet-stream".into(),
            Vec::new(),
        )));
        let input = NewTarget::ink(Arc::new(INK.to_vec()), CONTENT_TYPE);
        let mut next = 1;
        assert!(matches!(
            fresh_part_name(&package, &mut next, "xml", &input, Limits::default()),
            Err(Error::Opc(OpcError::DerivedPartNames { .. }))
        ));
        assert_eq!(next, 2);
    }

    #[test]
    fn inverse_refuses_a_foreign_edge_and_keeps_the_added_target() {
        let mut package = package();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/other.xml").unwrap(),
            ct::WML_HEADER.to_owned(),
            b"<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"
                .to_vec(),
        )));
        let (targets, delta) = AddDelta::insert_many(
            &mut package,
            &PackURI::new("/word/document.xml").unwrap(),
            &[NewTarget::ink(Arc::new(INK.to_vec()), CONTENT_TYPE)],
            Limits::default(),
        )
        .unwrap();
        let target = targets[0].part().clone();
        package
            .get_part_mut(&PackURI::new("/word/other.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::CUSTOM_XML.to_owned(),
                target.relative_ref("/word"),
                "foreign".to_owned(),
                false,
            );
        assert!(delta.inverse().apply(&mut package).is_err());
        assert!(package.get_part(&target).is_ok());
    }
}
