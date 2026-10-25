//! Targeted validation for one notes-slide relationship closure.
//!
//! The source and eager readers use the same validator. It follows only the
//! selected slide's notes edge: the notes slide, its owning slide, its notes
//! master, that master's theme, and the canonical presentation
//! `notesMasterId` edge. It does not scan the presentation's other slides or
//! materialize an unrelated notes graph.

use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part, Relationship};

use super::super::{SourceInner, SourcePart};
use crate::notes::{
    Conformance, Limits, MAX_MASTER_XML, MAX_NOTES_XML, MAX_PRESENTATION_XML, MAX_SLIDE_XML,
    MAX_THEME_XML, root_conformance_with_limits, scan_xml_with_limits, validate_id,
    validate_resource_xml_with_limits,
};
use crate::parts::is_relationship_type;
use crate::{Error, Result};

/// A bounded part loader used by both source-backed and eager notes reads.
///
/// The callback receives a part whose bytes have already passed the loader's
/// physical and markup-compatibility limits. Relationships and package
/// identity remain available through the same part view.
trait PartLoader {
    fn with_part<T>(
        &self,
        uri: &PackURI,
        content_type: &str,
        maximum: usize,
        operation: impl FnOnce(&dyn Part, usize) -> Result<T>,
    ) -> Result<T>;

    fn exists(&self, uri: &PackURI) -> Result<()>;
}

/// Charge the selected notes payload, notes master, and theme exactly once.
///
/// The owning slide and presentation root have independent related-root
/// ceilings and are intentionally not part of this notes-family aggregate.
/// The charge is measured from the raw part payload, before any
/// markup-compatibility promotion, so source-backed and eager reads apply the
/// same physical-byte policy.
struct AggregateBytes {
    observed: usize,
    maximum: usize,
}

impl AggregateBytes {
    fn new(maximum: usize) -> Self {
        Self {
            observed: 0,
            maximum,
        }
    }

    fn observe(&mut self, bytes: usize) -> Result<()> {
        self.observed = self
            .observed
            .checked_add(bytes)
            .ok_or_else(|| invalid("notes aggregate bytes overflow"))?;
        if self.observed > self.maximum {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                self.observed,
                self.maximum,
            ));
        }
        Ok(())
    }
}

struct SourceLoader<'a> {
    owner: &'a SourceInner,
}

impl PartLoader for SourceLoader<'_> {
    fn with_part<T>(
        &self,
        uri: &PackURI,
        content_type: &str,
        maximum: usize,
        operation: impl FnOnce(&dyn Part, usize) -> Result<T>,
    ) -> Result<T> {
        let part = super::load_part(self.owner, uri, content_type, maximum)?;
        // `load_part` may retain MCE-promoted XML in `Part::blob()`.  The
        // notes aggregate follows the physical source contract, so charge
        // the original managed payload rather than that processed view.
        let raw_bytes = part.source_data().as_bytes().len();
        operation(&part, raw_bytes)
    }

    fn exists(&self, uri: &PackURI) -> Result<()> {
        self.owner.package.part(uri).map(|_| ()).map_err(Error::Opc)
    }
}

struct EagerLoader<'a> {
    package: &'a OpcPackage,
}

impl PartLoader for EagerLoader<'_> {
    fn with_part<T>(
        &self,
        uri: &PackURI,
        content_type: &str,
        maximum: usize,
        operation: impl FnOnce(&dyn Part, usize) -> Result<T>,
    ) -> Result<T> {
        let part = self
            .package
            .get_part(uri)
            .map_err(|_| Error::Relationship("notes relationship target part is missing".into()))?;
        if part.content_type() != content_type {
            return Err(Error::ContentType {
                expected: content_type.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        let raw_bytes = part.blob().len();
        let processed = super::process_eager_part(part, maximum, self.package.read_limits())?;
        operation(&processed, raw_bytes)
    }

    fn exists(&self, uri: &PackURI) -> Result<()> {
        self.package
            .get_part(uri)
            .map(|_| ())
            .map_err(|_| Error::Relationship("notes relationship target part is missing".into()))
    }
}

pub(super) fn validate_with_limits(
    owner: &SourceInner,
    slide: &PackURI,
    part: &SourcePart,
    limits: Limits,
) -> Result<()> {
    let presentation_uri = owner._presentation.source_partname().clone();
    let presentation_content_type = owner._presentation.content_type().to_owned();
    let loader = SourceLoader { owner };
    validate_selected(
        &loader,
        slide,
        &presentation_uri,
        &presentation_content_type,
        part,
        part.source_data().as_bytes().len(),
        limits,
    )
}

pub(super) fn validate_eager_with_limits(
    package: &OpcPackage,
    slide: &dyn Part,
    part: &dyn Part,
    limits: Limits,
) -> Result<()> {
    let presentation_part = package.main_document_part()?;
    let presentation_uri = presentation_part.partname().clone();
    let presentation_content_type = presentation_part.content_type().to_owned();
    let notes_uri = part.partname().clone();
    let loader = EagerLoader { package };
    loader.with_part(
        &notes_uri,
        ct::PML_NOTES_SLIDE,
        MAX_NOTES_XML.min(limits.max_input_bytes),
        |selected, raw_bytes| {
            validate_selected(
                &loader,
                slide.partname(),
                &presentation_uri,
                &presentation_content_type,
                selected,
                raw_bytes,
                limits,
            )
        },
    )
}

pub(super) fn processed_eager_notes_with_limits(
    part: &dyn Part,
    source_limits: litchi_opc::SourceLimits,
    limits: Limits,
) -> Result<Arc<Vec<u8>>> {
    Ok(super::process_eager_part(
        part,
        MAX_NOTES_XML.min(limits.max_input_bytes),
        source_limits,
    )?
    .blob_arc())
}

fn validate_selected<L: PartLoader>(
    loader: &L,
    slide: &PackURI,
    presentation_uri: &PackURI,
    presentation_content_type: &str,
    part: &dyn Part,
    notes_bytes: usize,
    limits: Limits,
) -> Result<()> {
    require_content_type(part, ct::PML_NOTES_SLIDE)?;
    validate_leaf_path(part.partname(), "/ppt/notesSlides/")?;

    // This is deliberately the raw source payload length.  A source-backed
    // loader can expose MCE-promoted bytes through `Part::blob()`, while the
    // selected notes aggregate is the same physical notes/master/theme
    // budget constrained by `SourceLimits::max_total_part_bytes`.
    let mut aggregate = AggregateBytes::new(limits.max_total_bytes);
    aggregate.observe(notes_bytes)?;

    let conformance = validate_xml(
        part,
        MAX_NOTES_XML.min(limits.max_input_bytes),
        "notes",
        "notes slide",
        limits,
    )?;

    // The backlink target is checked against the selected slide, and the
    // selected slide itself is loaded so a stale relationship cannot validate
    // against a missing or non-slide part.
    let slide_conformance = loader.with_part(
        slide,
        ct::PML_SLIDE,
        // Slide XML has its own 16-MiB related-root ceiling.  The selected
        // notes policy still governs the notes, master, theme, and
        // presentation roots, but must not narrow this owning-slide proof to
        // the notes payload's default 8-MiB limit.
        MAX_SLIDE_XML,
        |slide_part, _raw_bytes| {
            root_conformance_with_limits(slide_part.blob(), MAX_SLIDE_XML, "sld", limits)
        },
    )?;
    if slide_conformance != conformance {
        return Err(invalid("notes slide and owning slide conformance differ"));
    }

    let (backlink_target, master_target) = required_notes_relationships(loader, part, conformance)?;
    if backlink_target != *slide {
        return Err(invalid("notes slide backlink targets the selected slide"));
    }

    validate_leaf_path(&master_target, "/ppt/notesMasters/")?;
    validate_presentation_notes_master(
        loader,
        presentation_uri,
        presentation_content_type,
        conformance,
        &master_target,
        limits,
    )?;

    let (master_conformance, theme_target) = loader.with_part(
        &master_target,
        ct::PML_NOTES_MASTER,
        MAX_MASTER_XML.min(limits.max_input_bytes),
        |master, raw_bytes| {
            aggregate.observe(raw_bytes)?;
            let master_conformance = validate_xml(
                master,
                MAX_MASTER_XML,
                "notesMaster",
                "notes master",
                limits,
            )?;
            let theme_target = required_theme(loader, master, conformance)?;
            Ok((master_conformance, theme_target))
        },
    )?;
    if master_conformance != conformance {
        return Err(invalid("notes master and notes slide conformance differ"));
    }

    validate_leaf_path(&theme_target, "/ppt/theme/")?;
    let theme_conformance = loader.with_part(
        &theme_target,
        crate::notes::THEME_CT,
        MAX_THEME_XML.min(limits.max_input_bytes),
        |theme, raw_bytes| {
            aggregate.observe(raw_bytes)?;
            let theme_conformance =
                validate_xml(theme, MAX_THEME_XML, "theme", "notes-master theme", limits)?;
            validate_opaque_relationships(loader, theme)?;
            Ok(theme_conformance)
        },
    )?;
    if theme_conformance != conformance {
        return Err(invalid("notes theme and notes slide conformance differ"));
    }
    Ok(())
}

/// Validate the selected notes XML and return the dialect its root uses.
fn validate_xml(
    part: &dyn Part,
    maximum: usize,
    root: &str,
    label: &str,
    limits: Limits,
) -> Result<Conformance> {
    let conformance = root_conformance_with_limits(part.blob(), maximum, root, limits)?;
    // `root_conformance` checks the root and namespace. The resource scan
    // additionally refuses authored relationship references, matching the
    // complete notes graph's inert-resource policy.
    validate_resource_xml_with_limits(part.blob(), maximum, conformance, root, label, limits)?;
    Ok(conformance)
}

/// Bind the selected notes master to the canonical `p:notesMasterId` edge in
/// the retained presentation root. The notes slide's relationship is an
/// owned closure edge; the presentation edge is the package's canonical
/// declaration and must resolve to the same physical part.
fn validate_presentation_notes_master<L: PartLoader>(
    loader: &L,
    presentation_uri: &PackURI,
    presentation_content_type: &str,
    conformance: Conformance,
    selected: &PackURI,
    limits: Limits,
) -> Result<()> {
    loader.with_part(
        presentation_uri,
        presentation_content_type,
        MAX_PRESENTATION_XML.min(limits.max_input_bytes),
        |presentation, _raw_bytes| {
            let presentation_conformance = root_conformance_with_limits(
                presentation.blob(),
                MAX_PRESENTATION_XML,
                "presentation",
                limits,
            )?;
            if presentation_conformance != conformance {
                return Err(invalid("presentation and notes slide conformance differ"));
            }

            let presentation_scan = scan_xml_with_limits(
                presentation.blob(),
                MAX_PRESENTATION_XML,
                conformance,
                "presentation",
                limits,
            )?;
            if presentation_scan.notes_master_ids.len() > 1 {
                return Err(invalid(
                    "presentation has multiple notesMasterId references",
                ));
            }
            let relationship_id = presentation_scan
                .notes_master_ids
                .first()
                .map(String::as_str)
                .ok_or_else(|| invalid("presentation lacks notesMasterId"))?;
            validate_id(relationship_id, "presentation notes-master ID")?;
            let relationship = presentation
                .rels()
                .get(relationship_id)
                .ok_or_else(|| invalid("presentation notes-master relationship is missing"))?;

            let mut equivalent_count = 0usize;
            let mut has_mixed_conformance = false;
            for candidate in presentation.rels().iter() {
                validate_id(candidate.r_id(), "presentation relationship ID")?;
                if is_relationship_type(
                    candidate.reltype(),
                    conformance.notes_master_rel(),
                    "notesMaster",
                ) {
                    equivalent_count += 1;
                    has_mixed_conformance |= candidate.reltype() != conformance.notes_master_rel();
                }
            }
            if equivalent_count != 1 {
                return Err(invalid(
                    "presentation has duplicate notes-master relationships",
                ));
            }
            if has_mixed_conformance || relationship.reltype() != conformance.notes_master_rel() {
                return Err(invalid(
                    "presentation notes-master relationship has mixed conformance",
                ));
            }
            let target = internal_target(relationship, "presentation notes-master relationship")?;
            if &target != selected {
                return Err(invalid(
                    "presentation notesMasterId targets a different notes master",
                ));
            }
            Ok(())
        },
    )
}

fn required_notes_relationships<L: PartLoader>(
    loader: &L,
    part: &dyn Part,
    conformance: Conformance,
) -> Result<(PackURI, PackURI)> {
    let mut backlink = None;
    let mut notes_master = None;

    for relationship in part.rels().iter() {
        validate_id(relationship.r_id(), "notes relationship ID")?;
        if relationship.is_external() {
            return Err(invalid("notes slide has an external relationship"));
        }
        if is_relationship_type(relationship.reltype(), conformance.slide_rel(), "slide") {
            if relationship.reltype() != conformance.slide_rel() {
                return Err(invalid("notes slide backlink has mixed conformance"));
            }
            let target = internal_target(relationship, "notes slide backlink")?;
            if backlink.replace(target).is_some() {
                return Err(invalid("notes slide has multiple slide backlinks"));
            }
        } else if is_relationship_type(
            relationship.reltype(),
            conformance.notes_master_rel(),
            "notesMaster",
        ) {
            if relationship.reltype() != conformance.notes_master_rel() {
                return Err(invalid(
                    "notes slide notes-master relationship has mixed conformance",
                ));
            }
            let target = internal_target(relationship, "notes-master relationship")?;
            if notes_master.replace(target).is_some() {
                return Err(invalid(
                    "notes slide has multiple notes-master relationships",
                ));
            }
        } else {
            validate_opaque_relationship(loader, relationship)?;
        }
    }

    let backlink = backlink.ok_or_else(|| invalid("notes slide lacks slide backlink"))?;
    let notes_master =
        notes_master.ok_or_else(|| invalid("notes slide lacks notes-master relationship"))?;
    Ok((backlink, notes_master))
}

fn required_theme<L: PartLoader>(
    loader: &L,
    master: &dyn Part,
    conformance: Conformance,
) -> Result<PackURI> {
    let mut theme = None;
    for relationship in master.rels().iter() {
        validate_id(relationship.r_id(), "notes-master relationship ID")?;
        if is_relationship_type(relationship.reltype(), conformance.theme_rel(), "theme") {
            if relationship.reltype() != conformance.theme_rel() {
                return Err(invalid(
                    "notes-master theme relationship has mixed conformance",
                ));
            }
            if theme.replace(relationship).is_some() {
                return Err(invalid("notes master has multiple theme relationships"));
            }
            if relationship.is_external() {
                return Err(invalid("notes-master theme relationship is external"));
            }
            theme = Some(relationship);
        } else {
            validate_opaque_relationship(loader, relationship)?;
        }
    }

    let theme = theme.ok_or_else(|| invalid("notes master lacks a theme relationship"))?;
    internal_target(theme, "notes-master theme relationship")
}

fn validate_opaque_relationship<L: PartLoader>(
    loader: &L,
    relationship: &Relationship,
) -> Result<()> {
    validate_id(relationship.r_id(), "notes relationship ID")?;
    if relationship.is_external() {
        return Ok(());
    }
    let target = relationship.target_partname()?;
    loader.exists(&target)
}

fn validate_opaque_relationships<L: PartLoader>(loader: &L, source: &dyn Part) -> Result<()> {
    for relationship in source.rels().iter() {
        validate_opaque_relationship(loader, relationship)?;
    }
    Ok(())
}

fn internal_target(relationship: &Relationship, label: &'static str) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(invalid(label));
    }
    Ok(relationship.target_partname()?)
}

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: expected.to_owned(),
            actual: part.content_type().to_owned(),
        })
    }
}

fn validate_leaf_path(uri: &PackURI, prefix: &str) -> Result<()> {
    let Some(rest) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(
            "notes relationship target has an invalid part path",
        ));
    };
    if rest.is_empty()
        || rest.contains('/')
        || !rest
            .rsplit_once('.')
            .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("xml"))
    {
        return Err(invalid(
            "notes relationship target has an invalid part path",
        ));
    }
    Ok(())
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.to_owned())
}
