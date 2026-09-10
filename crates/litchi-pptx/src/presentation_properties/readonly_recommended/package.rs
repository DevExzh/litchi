//! OPC lifecycle for the presentation-properties read-only recommendation.

use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcError, OpcPackage, PackURI};

use super::transaction::{Commit, Patch, Snapshot};
use super::{EXTENSION_URI, NAMESPACE};
use crate::presentation_properties::{P_NS, P_STRICT};
use crate::{Error, Result};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/presProps";
const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";

/// Read the optional typed recommendation from the presentation-properties
/// owner. `None` means the extension is absent.
pub fn load(package: &OpcPackage) -> Result<Option<bool>> {
    Ok(load_snapshot(package)?.and_then(|snapshot| snapshot.value()))
}

/// Capture the exact presentation-properties source, including an absent
/// recommendation when the owner part itself exists.
pub fn load_snapshot(package: &OpcPackage) -> Result<Option<Snapshot>> {
    let Some(uri) = find_properties_part(package)? else {
        return Ok(None);
    };
    let part = package.get_part(&uri)?;
    Snapshot::from_part(uri.to_string(), part.blob_arc()).map(Some)
}

/// Apply a source-checked recommendation patch atomically.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let name = PackURI::new(patch.before().source_part_name()).map_err(Error::Uri)?;
    let current = load_snapshot(package)?
        .ok_or_else(|| invalid("presentation-properties source part is absent"))?;
    if current.source_part_name() != patch.before().source_part_name()
        || current.source_xml() != patch.before().source_xml()
        || current.revision() != patch.before().revision()
    {
        return Err(invalid("presentation-properties source is stale"));
    }
    if patch.is_empty() {
        return Ok(current);
    }
    ensure_signature_policy(package)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&name)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate)?
        .ok_or_else(|| invalid("published presentation-properties source disappeared"))?;
    if result.source_part_name() != patch.after().source_part_name()
        || result.source_xml() != patch.after().source_xml()
        || result.value() != patch.after().value()
    {
        return Err(invalid(
            "published readonlyRecommended source differs from the patch",
        ));
    }
    *package = candidate;
    Ok(result)
}

/// Apply a committed recommendation edit atomically.
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

/// Replace or create the typed recommendation.
pub fn put(package: &mut OpcPackage, value: bool) -> Result<Option<bool>> {
    let previous = load(package)?;
    if previous == Some(value) {
        return Ok(previous);
    }
    if let Some(snapshot) = load_snapshot(package)? {
        let mut edit = snapshot.edit();
        edit.set_readonly_recommended(value)?;
        let commit = edit.commit()?;
        let previous = commit.before_value();
        apply_commit(package, commit)?;
        return Ok(previous);
    }

    ensure_signature_policy(package)?;
    let mut candidate = package.clone();
    let presentation = candidate.main_document_part()?.partname().clone();
    if candidate
        .iter_parts()
        .any(|part| part.content_type() == CONTENT_TYPE)
    {
        return Err(invalid(
            "orphan presentation-properties part exists without its relationship",
        ));
    }
    let strict = strict_presentation_root(candidate.main_document_part()?.blob())?;
    let p = if strict { P_STRICT } else { P_NS };
    let lexical = if value { "true" } else { "false" };
    let xml = format!(
        r#"<p:presentationPr xmlns:p="{p}"><p:extLst><p:ext uri="{EXTENSION_URI}"><p1710:readonlyRecommended xmlns:p1710="{NAMESPACE}" val="{lexical}"/></p:ext></p:extLst></p:presentationPr>"#
    )
    .into_bytes();
    let uri = candidate.next_partname("/ppt/presProps%d.xml")?;
    candidate.try_add_part(Box::new(BlobPart::new(
        uri.clone(),
        ct::PML_PRES_PROPS.to_owned(),
        xml,
    )))?;
    let target = uri.relative_ref(presentation.base_uri());
    let relationship_id = next_relationship_id(candidate.main_document_part()?);
    candidate
        .get_part_mut(&presentation)?
        .rels_mut()
        .add_relationship(
            if strict { STRICT_REL } else { REL }.to_owned(),
            target,
            relationship_id,
            false,
        );
    if load(&candidate)? != Some(value) {
        return Err(invalid(
            "created readonlyRecommended source did not round-trip",
        ));
    }
    *package = candidate;
    Ok(previous)
}

/// Remove the typed recommendation while retaining the owning properties part.
pub fn remove(package: &mut OpcPackage) -> Result<Option<bool>> {
    let previous = load(package)?;
    let Some(previous_value) = previous else {
        return Ok(None);
    };
    let snapshot = load_snapshot(package)?
        .ok_or_else(|| invalid("presentation-properties source part disappeared"))?;
    let mut edit = snapshot.edit();
    edit.remove()?;
    let commit = edit.commit()?;
    if commit.before_value() != Some(previous_value) {
        return Err(invalid("readonlyRecommended source changed during removal"));
    }
    apply_commit(package, commit)?;
    Ok(Some(previous_value))
}

fn find_properties_part(package: &OpcPackage) -> Result<Option<PackURI>> {
    let presentation = package.main_document_part()?;
    let strict = strict_presentation_root(presentation.blob())?;
    let mut owner_parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let owner_part = owner_parts.next();
    if owner_parts.next().is_some() {
        return Err(invalid(
            "presentation has multiple presentation-properties parts",
        ));
    }
    let mut relationships = presentation
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL));
    let Some(relationship) = relationships.next() else {
        if owner_part.is_some() {
            return Err(invalid(
                "orphan presentation-properties part exists without its relationship",
            ));
        }
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "presentation has multiple presentation-properties relationships",
        ));
    }
    let expected_relationship = if strict { STRICT_REL } else { REL };
    if relationship.reltype() != expected_relationship {
        return Err(invalid(
            "presentation-properties relationship profile does not match the presentation root",
        ));
    }
    if relationship.is_external() {
        return Err(invalid(
            "presentation-properties relationship cannot be external",
        ));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(invalid(
            "presentation-properties relationship target must be a part URI",
        ));
    }
    let uri = relationship.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "presentation-properties part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if owner_part.is_some_and(|owner| owner.partname() != &uri) {
        return Err(invalid(
            "orphan presentation-properties part is not the relationship target",
        ));
    }
    Ok(Some(uri))
}

fn strict_presentation_root(xml: &[u8]) -> Result<bool> {
    let mut reader = quick_xml::reader::NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element) => {
                if element.local_name().as_ref() != b"presentation" {
                    return Err(invalid("presentation root is not a presentation element"));
                }
                return match namespace {
                    quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value))
                        if value == P_STRICT.as_bytes() =>
                    {
                        Ok(true)
                    },
                    quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value))
                        if value == P_NS.as_bytes() =>
                    {
                        Ok(false)
                    },
                    _ => Err(invalid("presentation root has an unsupported namespace")),
                };
            },
            quick_xml::events::Event::Decl(_) | quick_xml::events::Event::Comment(_) => {},
            quick_xml::events::Event::Text(text)
                if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            _ => return Err(invalid("presentation root is missing")),
        }
    }
}

fn next_relationship_id(presentation: &dyn litchi_opc::Part) -> String {
    let mut index = 1usize;
    loop {
        let candidate = format!("rIdReadonlyRecommended{index}");
        if presentation.rels().get(&candidate).is_none() {
            return candidate;
        }
        index = index.saturating_add(1);
    }
}

fn ensure_signature_policy(package: &OpcPackage) -> Result<()> {
    if package.is_signed() || package.requires_signature_edit_policy() {
        return Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
