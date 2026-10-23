//! Slide metadata dispatch through the existing native transaction family.
use super::{
    Commit, Editable, Pptx, Proof, allocation, bound, error, native_error, native_snapshot_budget,
    opened_edit_options, source_edit_options, validate_text,
};
use crate::format::pptx::Nav;
use crate::format::pptx::node::Item;
use crate::format::pptx::spec::{SLIDE_WRITE_TABLE, SlideProperty};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, Props, Result};
use litchi_core::Position;
use std::collections::HashSet;
use std::sync::Arc;

pub(super) fn accepts(target: &TargetArg<'_>, props: &Props<'_>) -> bool {
    target.path().is_some_and(|path| {
        path.path()
            .steps()
            .last()
            .is_some_and(|step| step.name().eq_ignore_ascii_case("slide"))
    }) || props.iter().any(|(key, _)| {
        key.as_str().eq_ignore_ascii_case("name") || key.as_str().eq_ignore_ascii_case("hidden")
    })
}

impl Pptx {
    pub(super) fn set_slide_properties(
        &self,
        target: TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let mut name = None;
        let mut hidden = None;
        // The layout retargets the slide's layout relationship (record 0346).
        let mut layout = None;
        // `Some(None)` removes the slide's direct transition.
        let mut transition: Option<Option<litchi_pptx::transition::Transition>> = None;
        let mut properties = 0_u8;
        for (key, value) in props.iter() {
            let (index, entry) = SLIDE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|error| error.with_property(key.as_str().to_owned()))?;
            let bit = 1_u8 << index;
            if properties & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a slide property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            properties |= bit;
            entry.prepare(Operation::Set, value, |setter, value| {
                match setter {
                    SlideProperty::Name => {
                        let value = match value {
                            Some(Scalar::Text(value)) => value,
                            None => "",
                            _ => unreachable!("slide name coercion"),
                        };
                        bound(
                            LimitKind::StagedBytes,
                            value.len(),
                            limits.max_staged_bytes(),
                        )?;
                        validate_text(value).map_err(|error| error.with_property("name"))?;
                        let mut owned = String::new();
                        owned.try_reserve_exact(value.len()).map_err(allocation)?;
                        owned.push_str(value);
                        name = Some(owned);
                    },
                    SlideProperty::Hidden => {
                        let Some(Scalar::Bool(value)) = value else {
                            unreachable!("slide visibility coercion")
                        };
                        hidden = Some(value);
                    },
                    SlideProperty::Transition => {
                        transition = Some(match value {
                            Some(Scalar::Text(value)) => super::transition::parse(value)?,
                            None => None,
                            _ => unreachable!("slide transition coercion"),
                        });
                    },
                    SlideProperty::Layout => {
                        let Some(Scalar::Text(value)) = value else {
                            unreachable!("slide layout coercion")
                        };
                        bound(
                            LimitKind::StagedBytes,
                            value.len(),
                            limits.max_staged_bytes(),
                        )?;
                        let mut owned = String::new();
                        owned.try_reserve_exact(value.len()).map_err(allocation)?;
                        owned.push_str(value);
                        layout = Some(owned);
                    },
                    SlideProperty::Notes => unreachable!("notes are routed before slide metadata"),
                    SlideProperty::Title | SlideProperty::Text => {
                        unreachable!("the title and text shorthand is add only")
                    },
                }
                Ok(())
            })?;
        }
        let mut targets = Vec::new();
        let mut seen = HashSet::new();
        let charge = name
            .as_ref()
            .map_or(0, String::len)
            .saturating_add(size_of::<usize>() + size_of::<Option<bool>>())
            .saturating_add(if transition.is_some() { 64 } else { 0 })
            .saturating_add(layout.as_ref().map_or(0, String::len));
        let metadata = name.is_some() || hidden.is_some();
        let mut append = |node: crate::format::pptx::Node<'_>| -> Result<()> {
            let Item::Slide(slide) = node.item() else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "slide properties require slide targets",
                ));
            };
            if seen.contains(&slide.position) {
                return Ok(());
            }
            let count = targets.len().saturating_add(1);
            bound(
                LimitKind::ExpandedTargets,
                count,
                limits.max_expanded_targets(),
            )?;
            bound(
                LimitKind::StagedBytes,
                count.saturating_mul(charge),
                limits.max_staged_bytes(),
            )?;
            targets.try_reserve(1).map_err(allocation)?;
            seen.try_reserve(1).map_err(allocation)?;
            targets.push(Position::new(slide.position as usize));
            seen.insert(slide.position);
            Ok(())
        };
        let nav = Nav::new(self);
        if let Some(path) = target.path() {
            append(crate::path::resolve::get(&nav, path)?)?;
        } else if let Some(query) = target.query() {
            let walk = crate::query::run(nav, query)?;
            if !walk.plan().diagnostics().is_empty() {
                return Err(error(
                    Code::InvalidSelector,
                    "mutation selector contains an unsupported predicate",
                ));
            }
            for node in walk {
                append(node?)?;
            }
        } else {
            return Err(error(Code::InvalidArgument, "mutation target is empty"));
        }
        if targets.is_empty() {
            return Err(error(
                Code::ZeroMatches,
                "mutation selector matched no element",
            ));
        }
        let source = self.editable_owner_for_write()?;
        // A layout change edits the slide's relationships, which the
        // source-backed slide edit does not stage: a source owner is promoted
        // (record 0346).
        let promoted = match (&source, &layout) {
            (Editable::Source(_), Some(_)) => Some(Editable::Eager(self.eager_base(&source)?)),
            _ => None,
        };
        let owner = promoted.as_ref().unwrap_or(&source);
        let native_limit = native_snapshot_budget(limits);
        let (target, proof, changed) = match owner {
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_limit))
                    .map_err(native_error)?;
                for position in &targets {
                    if metadata {
                        edit.set_slide_properties(*position, name.as_deref(), hidden)?;
                    }
                    if let Some(value) = &transition {
                        edit.set_slide_transition(*position, value.as_ref())?;
                    }
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (Editable::Source(target), Proof::Source(patch), changed)
            },
            Editable::Eager(package) => {
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(native_error)?;
                for position in &targets {
                    if metadata {
                        edit.set_slide_properties(*position, name.as_deref(), hidden)?;
                    }
                    if let Some(value) = &transition {
                        edit.set_slide_transition(*position, value.as_ref())?;
                    }
                    if let Some(layout) = &layout {
                        edit.set_slide_layout(*position, layout)?;
                    }
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (
                    Editable::Eager(Arc::new(target)),
                    Proof::Eager(patch),
                    changed,
                )
            },
        };
        let after = if changed {
            match &target {
                Editable::Source(package) => self.adopt_source(package.clone())?,
                Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
            }
        } else {
            self.clone()
        };
        Ok(Commit::new(after, source, target, proof, targets.len()))
    }
}
