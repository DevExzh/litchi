//! OPC binding validation and atomic publication.

use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};

use super::codec::{self, Located};
use super::model::{Binding, Snapshot};
use super::{Commit, Limits, Patch, Tags};
use crate::{Error, Result};

const PRESENTATION_PART: &str = "/ppt/presentation.xml";

/// Load a singular optional tag list under safe default bounds.
pub fn load(package: &OpcPackage, slide_id: u32) -> Result<Option<Tags>> {
    let snapshot = load_snapshot(package, slide_id)?;
    snapshot
        .tags()?
        .map(|tags| codec::clone_tags(tags, Limits::default()))
        .transpose()
}

/// Load a source-bound inventory under safe default bounds.
pub fn load_snapshot(package: &OpcPackage, slide_id: u32) -> Result<Snapshot> {
    load_snapshot_with_limits(package, slide_id, Limits::default())
}

/// Load a source-bound inventory under caller-supplied bounds.
pub fn load_snapshot_with_limits(
    package: &OpcPackage,
    slide_id: u32,
    limits: Limits,
) -> Result<Snapshot> {
    let presentation = package.main_document_part()?;
    if presentation.partname().as_str() != PRESENTATION_PART {
        return Err(Error::Invalid(
            "Designer tags require the /ppt/presentation.xml owner".into(),
        ));
    }
    if !crate::parts::expected_main_content_type(presentation.content_type()) {
        return Err(Error::ContentType {
            expected: "PresentationML main content type".into(),
            actual: presentation.content_type().into(),
        });
    }
    if presentation.blob().len() > limits.xml_bytes() {
        return Err(Error::Limit {
            resource: "Designer-tag owner XML bytes",
            limit: limits.xml_bytes(),
        });
    }
    let source = presentation.blob_arc();
    let located = codec::locate(source.as_slice(), slide_id, limits)?;
    let binding = validate_binding(package, presentation, &located.layout.relationship_id)?;
    snapshot_from_located(
        presentation.partname().to_string(),
        presentation.content_type().to_owned(),
        source,
        slide_id,
        binding,
        located,
        limits,
    )
}

/// Store a singular tag list atomically and return the previous value.
pub fn store(package: &mut OpcPackage, slide_id: u32, value: &Tags) -> Result<Option<Tags>> {
    let limits = Limits::default();
    let value = codec::clone_tags(value, limits)?;
    let snapshot = load_snapshot(package, slide_id)?;
    let previous = snapshot
        .tags()?
        .map(|tags| codec::clone_tags(tags, limits))
        .transpose()?;
    let mut edit = snapshot.edit()?;
    edit.set(value)?;
    apply_commit(package, edit.commit()?)?;
    Ok(previous)
}

/// Remove a singular tag list atomically and return the previous value.
pub fn remove(package: &mut OpcPackage, slide_id: u32) -> Result<Option<Tags>> {
    let snapshot = load_snapshot(package, slide_id)?;
    let previous = snapshot
        .tags()?
        .map(|tags| codec::clone_tags(tags, Limits::default()))
        .transpose()?;
    if previous.is_none() {
        return Ok(None);
    }
    let mut edit = snapshot.edit()?;
    edit.remove();
    apply_commit(package, edit.commit()?)?;
    Ok(previous)
}

/// Apply a committed edit atomically.
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

/// Apply a patch after re-resolving its stable slide identity and binding.
///
/// Reordering sibling `p:sldId` elements is allowed because the source check
/// is scoped to the selected host. Deletion, rebinding, duplicate identity,
/// and concurrent edits of that host are refused.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let current = load_snapshot_with_limits(package, patch.before.slide_id, patch.before.limits)?;
    if current.presentation_part_name != patch.before.presentation_part_name
        || current.presentation_content_type != patch.before.presentation_content_type
        || current.binding != patch.before.binding
        || current.layout.host_bytes(current.source_xml.as_slice())
            != patch
                .before
                .layout
                .host_bytes(patch.before.source_xml.as_slice())
    {
        return Err(Error::Invalid(
            "Designer-tag selected slide-ID source is stale or rebound".into(),
        ));
    }
    if !patch.is_changed() {
        return Ok(current);
    }
    let desired = patch.after.singular()?;
    let staged = codec::rewrite(
        current.source_xml.as_slice(),
        &current.layout,
        desired,
        current.limits,
    )?;
    let part_name = PackURI::new(&current.presentation_part_name).map_err(Error::Uri)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&part_name)?
        .set_blob_shared(Arc::new(staged));
    let published = load_snapshot_with_limits(&candidate, current.slide_id, current.limits)?;
    if published.binding != patch.after.binding
        || published.singular()? != patch.after.singular()?
        || published.layout.host_bytes(published.source_xml.as_slice())
            != patch
                .after
                .layout
                .host_bytes(patch.after.source_xml.as_slice())
    {
        return Err(Error::Invalid(
            "published Designer-tag candidate differs from the committed host".into(),
        ));
    }
    candidate.unsign();
    *package = candidate;
    Ok(published)
}

pub(crate) fn snapshot_from_located(
    presentation_part_name: String,
    presentation_content_type: String,
    source_xml: Arc<Vec<u8>>,
    slide_id: u32,
    binding: Binding,
    located: Located,
    limits: Limits,
) -> Result<Snapshot> {
    let revision = revision(
        located.layout.host_bytes(source_xml.as_slice()),
        slide_id,
        &binding,
        &presentation_part_name,
        &presentation_content_type,
    );
    Ok(Snapshot {
        presentation_part_name,
        presentation_content_type,
        source_xml,
        slide_id,
        binding,
        occurrences: located.tags,
        layout: located.layout,
        limits,
        revision,
    })
}

fn validate_binding(
    package: &OpcPackage,
    presentation: &dyn litchi_opc::Part,
    relationship_id: &str,
) -> Result<Binding> {
    let relationship = presentation.rels().get(relationship_id).ok_or_else(|| {
        Error::Relationship(format!(
            "selected slide ID references missing relationship '{relationship_id}'"
        ))
    })?;
    if relationship.is_external() {
        return Err(Error::Relationship(
            "selected slide relationship must be internal".into(),
        ));
    }
    if !crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide") {
        return Err(Error::Relationship(format!(
            "selected slide relationship has unexpected type '{}'",
            relationship.reltype()
        )));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    Ok(Binding {
        relationship_id: relationship_id.to_owned(),
        relationship_type: relationship.reltype().to_owned(),
        relationship_target: relationship.target_ref().to_owned(),
        part_name: target.to_string(),
        part_content_type: part.content_type().to_owned(),
    })
}

fn revision(
    host: &[u8],
    slide_id: u32,
    binding: &Binding,
    part_name: &str,
    content_type: &str,
) -> super::Revision {
    let mut hash = 0xcbf29ce484222325u64;
    for bytes in [
        host,
        slide_id.to_le_bytes().as_slice(),
        binding.relationship_id.as_bytes(),
        binding.relationship_type.as_bytes(),
        binding.relationship_target.as_bytes(),
        binding.part_name.as_bytes(),
        binding.part_content_type.as_bytes(),
        part_name.as_bytes(),
        content_type.as_bytes(),
    ] {
        for byte in bytes {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(0x100000001b3);
        }
    }
    hash
}
