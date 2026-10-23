//! Speaker-notes text through the native eager edit (record 0252).
//!
//! `/slide[N]/notes` replaces an existing notes page's text; a slide's
//! `notes` property also creates the page when the slide has none. Source
//! owners are promoted to an eager package for the edit, since notes pages
//! are separate parts the source-backed slide edit does not stage.
use super::{Commit, Editable, Pptx, Proof, allocation, bound, error, validate_text};
use crate::format::pptx::Nav;
use crate::format::pptx::node::Item;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::{Code, PropValue, Props, Result};
use std::sync::Arc;

/// Whether a mutation addresses speaker notes: a `notes` path, or a slide
/// path whose only property is `notes`.
pub(super) fn accepts(target: &TargetArg<'_>, props: &Props<'_>) -> bool {
    let Some(last) = target.path().and_then(|path| path.path().steps().last()) else {
        return false;
    };
    last.name().eq_ignore_ascii_case("notes")
        || (last.name().eq_ignore_ascii_case("slide")
            && props
                .iter()
                .any(|(key, _)| key.as_str().eq_ignore_ascii_case("notes")))
}

pub(super) fn admitted_text(
    props: &Props<'_>,
    key: &'static str,
    limits: &SemanticLimits,
) -> Result<String> {
    let mut text = None;
    for (name, value) in props.iter() {
        if !name.as_str().eq_ignore_ascii_case(key) {
            return Err(error(
                Code::UnsupportedProperty,
                "speaker notes accept only their text in this call",
            )
            .with_property(name.as_str().to_owned()));
        }
        if text.is_some() {
            return Err(error(
                Code::InvalidArgument,
                "notes text was assigned more than once",
            )
            .with_property(key));
        }
        let value = match value {
            PropValue::Null => "",
            PropValue::Text(value) => value.as_ref(),
            _ => {
                return Err(error(Code::InvalidValue, "notes text must be text").with_property(key));
            },
        };
        bound(
            LimitKind::StagedBytes,
            value.len(),
            limits.max_staged_bytes(),
        )?;
        validate_text(value)?;
        let mut owned = String::new();
        owned.try_reserve_exact(value.len()).map_err(allocation)?;
        owned.push_str(value);
        text = Some(owned);
    }
    text.ok_or_else(|| error(Code::InvalidArgument, "no notes text was assigned"))
}

impl Pptx {
    pub(super) fn set_notes(
        &self,
        target: TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let Some(path) = target.path() else {
            return Err(error(
                Code::UnsupportedCapability,
                "notes writes take one slide or notes path",
            ));
        };
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let (slide, text) = match node.item() {
            Item::Notes(notes) => (notes.slide, admitted_text(props, "text", limits)?),
            Item::Slide(slide) => (slide.position, admitted_text(props, "notes", limits)?),
            _ => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "notes writes take one slide or notes path",
                ));
            },
        };
        let source = self.editable_owner()?;
        let base = match &source {
            Editable::Eager(package) => Arc::clone(package),
            Editable::Source(package) => {
                let bytes = package.to_bytes()?;
                Arc::new(litchi_pptx::Package::from_reader(std::io::Cursor::new(
                    bytes,
                ))?)
            },
        };
        let mut edit = base.edit()?;
        let changed = edit.set_notes_text(slide as usize, &text)?;
        let (target, patch) = edit.commit()?.into_parts();
        let changed = changed && !patch.is_noop();
        let target = Arc::new(target);
        let after = if changed {
            self.adopt_eager(Arc::clone(&target))
        } else {
            self.clone()
        };
        Ok(Commit {
            snapshot_bytes: 0,
            retained_cache: std::sync::OnceLock::new(),
            after,
            source,
            target: Editable::Eager(target),
            proof: Proof::Eager(patch),
            operations: 1,
        })
    }
}
