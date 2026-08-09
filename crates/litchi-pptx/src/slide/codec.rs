//! Bounded PresentationML/DrawingML adapters for the slide owner.
//!
//! The actual bounded XML root checks, strict/transitional namespace handling,
//! MCE processing, text limits, and zero-copy shape indexing remain owned by
//! the [`crate::parts::SlidePart`] family. This layer is intentionally an
//! explicit semantic adapter: it selects the appropriate validated part
//! reader for `Slide`, `SlideLayout`, or `SlideMaster` and does not duplicate
//! XML parsing or copy shape subtrees. Keeping that boundary makes unknown
//! `DrawingML` and all existing reader limits flow through the established
//! codec unchanged.

use crate::Result;
use crate::parts::{SlideLayoutPart, SlideMasterPart, SlidePart};
use crate::shape::Scene;

pub(crate) fn slide_name(part: &SlidePart<'_>) -> Result<String> {
    part.name()
}

pub(crate) fn slide_hidden(part: &SlidePart<'_>) -> Result<bool> {
    part.is_hidden()
}

pub(crate) fn slide_text(part: &SlidePart<'_>) -> Result<String> {
    part.text()
}

pub(crate) fn slide_shapes<'a>(part: &SlidePart<'a>) -> Result<Scene<'a>> {
    part.shapes()
}

pub(crate) fn layout_name(part: &SlideLayoutPart<'_>) -> Result<String> {
    part.name()
}

pub(crate) fn layout_kind(part: &SlideLayoutPart<'_>) -> Result<Option<String>> {
    part.kind()
}

pub(crate) fn layout_shapes<'a>(part: &SlideLayoutPart<'a>) -> Result<Scene<'a>> {
    part.shapes()
}

pub(crate) fn master_name(part: &SlideMasterPart<'_>) -> Result<String> {
    part.name()
}

pub(crate) fn master_preserved(part: &SlideMasterPart<'_>) -> Result<bool> {
    part.is_preserved()
}

pub(crate) fn master_shapes<'a>(part: &SlideMasterPart<'a>) -> Result<Scene<'a>> {
    part.shapes()
}
