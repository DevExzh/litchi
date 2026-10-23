//! Slide master and slide layout writes (record 0321).
//!
//! - `set /slidemaster[N] name=...` and `set /slidemaster[N]/slidelayout[M]
//!   name=...` (or `/slidelayout[M]`) rename the master or layout
//!   (`p:cSld/@name`); slides stay bound by relationship.
//! - `set` of `text` on a direct shape or placeholder of a master or layout
//!   (`/slidemaster[N]/placeholder[title]`, `/slidelayout[M]/shape[K]`)
//!   replaces its text, keeping the first run's formatting.
//!
//! Other master and layout properties (background, direction) and other
//! shape properties of masters and layouts are gaps (`capabilities()`).
//! Source owners are promoted to an eager package, as for theme writes.

use std::sync::Arc;

use litchi_core::Position;

use super::node::{Item, TreeContainer};
use super::spec::{LAYOUT_WRITE_TABLE, MASTER_WRITE_TABLE};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options};
use super::{Editable, Nav, Pptx};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, Props, Result};

/// What a master or layout `set` changes.
enum Change {
    MasterName {
        master: usize,
        name: String,
    },
    LayoutName {
        master: usize,
        layout: usize,
        name: String,
    },
    MasterText {
        master: usize,
        shape: usize,
        text: String,
    },
    LayoutText {
        master: usize,
        layout: usize,
        shape: usize,
        text: String,
    },
}

fn owned_text(value: Option<Scalar<'_>>, limits: &SemanticLimits) -> Result<String> {
    let text = match value {
        Some(Scalar::Text(text) | Scalar::Token(text)) => text,
        _ => "",
    };
    bound(
        LimitKind::StagedBytes,
        text.len(),
        limits.max_staged_bytes(),
    )?;
    super::write::validate_text(text)?;
    Ok(text.to_owned())
}

impl Pptx {
    /// A `set` on a master, a layout, or a shape of either; `None` for
    /// other targets.
    pub(super) fn try_set_master(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let Some(first) = path.path().steps().first() else {
            return Ok(None);
        };
        if !["slidemaster", "slidelayout"]
            .iter()
            .any(|name| first.name().eq_ignore_ascii_case(name))
        {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let change = match node.item() {
            Item::Master(master) => Change::MasterName {
                master: master.position as usize,
                name: self.admit_name(&MASTER_WRITE_TABLE, props, limits)?,
            },
            Item::Layout(layout) => Change::LayoutName {
                master: layout.master as usize,
                layout: layout.position as usize,
                name: self.admit_name(&LAYOUT_WRITE_TABLE, props, limits)?,
            },
            Item::Tree(tree) if tree.handle.parent().is_none() => {
                let text = admit_text(props, limits)?;
                let shape = tree.handle.position().get();
                match tree.container {
                    TreeContainer::Master(master) => Change::MasterText {
                        master: master as usize,
                        shape,
                        text,
                    },
                    TreeContainer::Layout { master, position } => Change::LayoutText {
                        master: master as usize,
                        layout: position as usize,
                        shape,
                        text,
                    },
                    TreeContainer::Slide(_) | TreeContainer::Notes(_) => return Ok(None),
                }
            },
            Item::Tree(_) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "master and layout text writes require a direct shape",
                ));
            },
            _ => return Ok(None),
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        drop(node);
        let source = self.editable_owner_for_write()?;
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        match &change {
            Change::MasterName { master, name } => {
                native.set_master_name(Position::new(*master), name)?;
            },
            Change::LayoutName {
                master,
                layout,
                name,
            } => {
                native.set_layout_name(Position::new(*master), Position::new(*layout), name)?;
            },
            Change::MasterText {
                master,
                shape,
                text,
            } => {
                native.set_master_shape_text(Position::new(*master), *shape, text)?;
            },
            Change::LayoutText {
                master,
                layout,
                shape,
                text,
            } => {
                native.set_layout_shape_text(
                    Position::new(*master),
                    Position::new(*layout),
                    *shape,
                    text,
                )?;
            },
        }
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Some(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        )))
    }

    fn admit_name<const N: usize>(
        &self,
        table: &crate::registry::Table<(), N, 0>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<String> {
        let mut name = None;
        for (key, value) in props.iter() {
            let (_, entry) = table
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            entry.prepare(Operation::Set, value, |(), value| {
                if name.is_some() {
                    return Err(error(
                        Code::InvalidArgument,
                        "a property was assigned more than once",
                    )
                    .with_property("name"));
                }
                name = Some(owned_text(value, limits)?);
                Ok(())
            })?;
        }
        name.ok_or_else(|| {
            error(Code::MissingProperty, "nothing to set").with_suggestion("set name=...")
        })
    }
}

/// The `text` of a master or layout shape `set`; other shape properties
/// are refused on masters and layouts.
fn admit_text(props: &Props<'_>, limits: &SemanticLimits) -> Result<String> {
    let mut text = None;
    for (key, value) in props.iter() {
        if !key.as_str().eq_ignore_ascii_case("text") {
            return Err(error(
                Code::UnsupportedCapability,
                "master and layout shapes take text writes only",
            )
            .with_property(key.as_str().to_owned()));
        }
        text = Some(match value {
            crate::PropValue::Null => String::new(),
            _ => value.with_coerced(crate::value::ValueType::Text, |value| {
                owned_text(value, limits)
            })?,
        });
    }
    text.ok_or_else(|| error(Code::MissingProperty, "nothing to set").with_property("text"))
}
