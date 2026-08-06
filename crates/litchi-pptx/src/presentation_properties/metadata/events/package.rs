//! Slide-show event OPC ownership and atomic publication.

use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};

use super::codec::{self, LoadLimits};
use super::model::{Draft, Event};
use super::transaction::{Commit, Patch, Snapshot};
use super::validation::validate_events;
use crate::{Error, Result};

/// Discover the typed event records on one slide without executing them.
pub fn load(package: &OpcPackage, slide_part_name: &PackURI) -> Result<Option<Vec<Event>>> {
    load_snapshot(package, slide_part_name)
        .map(|value| value.map(|snapshot| snapshot.events().to_vec()))
}

/// Capture one slide's event list and exact source context.
pub fn load_snapshot(package: &OpcPackage, slide_part_name: &PackURI) -> Result<Option<Snapshot>> {
    let slide = package.get_part(slide_part_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_owned(),
            actual: slide.content_type().to_owned(),
        });
    }
    let Some(located) = codec::locate(0, slide.blob(), &mut LoadLimits::default())? else {
        return Ok(None);
    };
    Snapshot::from_located(slide_part_name.to_string(), slide.blob_arc(), located).map(Some)
}

/// Add a new slide-show event extension atomically.
pub fn store(package: &mut OpcPackage, slide_part_name: &PackURI, events: &[Draft]) -> Result<()> {
    validate_events(events)?;
    let mut candidate = package.clone();
    codec::store(&mut candidate, slide_part_name, events)?;
    if load_snapshot(&candidate, slide_part_name)?.is_none() {
        return Err(invalid("stored slide-show events cannot be read back"));
    }
    candidate.unsign();
    *package = candidate;
    Ok(())
}

/// Remove only the MS-PPTX show-event extension, retaining neighboring opaque
/// extension blocks and all other slide XML.
pub fn remove(package: &mut OpcPackage, slide_part_name: &PackURI) -> Result<Option<Snapshot>> {
    let Some(before) = load_snapshot(package, slide_part_name)? else {
        return Ok(None);
    };
    let located = codec::locate(0, before.source_xml(), &mut LoadLimits::default())?
        .ok_or_else(|| invalid("slide-show event source disappeared during removal"))?;
    let updated = codec::remove_extension(before.source_xml(), &located)?;
    let mut candidate = package.clone();
    candidate.get_part_mut(slide_part_name)?.set_blob(updated);
    if load_snapshot(&candidate, slide_part_name)?.is_some() {
        return Err(invalid("slide-show event removal left an event extension"));
    }
    candidate.unsign();
    *package = candidate;
    Ok(Some(before))
}

/// Apply a committed source-checked event patch atomically.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let slide_name = PackURI::new(patch.before().slide_part_name()).map_err(Error::Uri)?;
    let current = load_snapshot(package, &slide_name)?
        .ok_or_else(|| invalid("slide-show event source is absent"))?;
    if !current.same_source(patch.before()) {
        return Err(invalid("slide-show event source is stale"));
    }
    if patch.is_empty() {
        return Ok(current);
    }
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&slide_name)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_snapshot(&candidate, &slide_name)?
        .ok_or_else(|| invalid("published slide-show event list cannot be read"))?;
    if !result.same_source(patch.after()) {
        return Err(invalid(
            "published slide-show event source differs from the patch",
        ));
    }
    candidate.unsign();
    *package = candidate;
    Ok(result)
}

/// Apply a committed edit and return its validated post-publication snapshot.
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
