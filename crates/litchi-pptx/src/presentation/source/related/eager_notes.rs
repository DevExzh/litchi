//! Eager per-slide notes snapshots.
//!
//! This leaf follows the selected slide's notes relationship and its bounded
//! owning-slide, notes-master, theme, and presentation-root closure. It does
//! not invoke the presentation-wide notes graph or serialize and reopen the
//! package. The returned [`super::Notes`] is detached from the package and
//! retains the same owned scene and semantic text projection as a
//! source-backed notes read.

use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use std::sync::Arc;

use crate::parts::{SlidePart, related_by_type};
use crate::shape::owned::Scene as OwnedScene;
use crate::{Result, shape};

use super::{Notes, RelatedPart, notes_text};

/// Read the selected slide's optional notes slide as a detached snapshot.
pub(crate) fn load<'a>(
    package: &'a OpcPackage,
    slide: &SlidePart<'a>,
    shape_limits: shape::Limits,
    notes_limits: crate::notes::Limits,
) -> Result<Option<Notes>> {
    let Some((relationship, part)) = related_by_type(
        package,
        slide.part(),
        rt::NOTES_SLIDE,
        "notesSlide",
        ct::PML_NOTES_SLIDE,
    )?
    else {
        return Ok(None);
    };

    super::notes::validate_eager_with_limits(package, slide.part(), part, notes_limits)?;
    let processed =
        super::notes::processed_eager_notes_with_limits(part, package.read_limits(), notes_limits)?;
    let parsed = shape::Scene::read(processed.as_slice(), shape_limits)?;
    let scene = OwnedScene::from_parsed(parsed, Arc::clone(&processed))?;
    super::strict_scene(&scene, "notes")?;
    let text = notes_text(&scene)?;
    Ok(Some(Notes::from_parts(
        RelatedPart::from_eager(part, relationship.r_id()),
        text,
        scene,
    )))
}
