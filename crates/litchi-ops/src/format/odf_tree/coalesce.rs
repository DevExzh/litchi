//! Coalesced object-format property sets for explicit transactions and
//! batches (record 0288).
//!
//! Consecutive direct-path sets on one drawing, chart, image, database or
//! formula snapshot are staged on one leaf `Edit` and published by one
//! commit, instead of one commit (and one rebuilt object tree) per set. A
//! set joins the pending run only when no pending set can change what it
//! addresses or reads: its path is positional (a name predicate could meet
//! a pending rename), every key is one whose write leaves every other
//! node's address and source values alone, and no pending set writes the
//! same property of the same node (a frame's or shape's position and size
//! count as one property: the leaf stages them as one geometry, merged
//! with the source values). Anything else closes the run first and
//! resolves against the published result.

use std::collections::HashSet;

use litchi_core::FileFormat;

use super::{Node, OdfTree, write};
use crate::error::Result;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// Whether setting `key` on a `kind` node may share a leaf edit with other
/// sets: it changes no address and no value another node's write reads.
fn coalescable(format: FileFormat, kind: &str, key: &str) -> bool {
    match format {
        FileFormat::Odg => matches!(
            (kind, key),
            ("page", "name") | ("shape", "text" | "name" | "x" | "y" | "width" | "height")
        ),
        FileFormat::Odc => matches!(kind, "chart" | "axis" | "series"),
        FileFormat::Odi => matches!(kind, "frame" | "metadata"),
        FileFormat::Odb => matches!((kind, key), ("query", "command")),
        FileFormat::Odf => matches!((kind, key), ("element", "text")),
        _ => false,
    }
}

/// The property family a key writes: position and size are one geometry.
fn family(key: &str) -> &str {
    match key {
        "x" | "y" | "width" | "height" => "geometry",
        other => other,
    }
}

/// A pending run of coalesced sets on one base snapshot.
pub(crate) struct Run {
    base: OdfTree,
    groups: Vec<(u32, Props<'static>)>,
    /// The (node, property family) pairs the run writes.
    targets: HashSet<(u32, String)>,
    first_index: Option<usize>,
    operations: usize,
}

impl OdfTree {
    /// Plan a coalescable set: the tree node a direct positional path names,
    /// when every key may share a leaf edit. `None` when the set must take
    /// the ordinary route.
    pub(crate) fn set_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<u32>> {
        if crate::format::odf_metadata::is_target(target) || props.is_empty() {
            return Ok(None);
        }
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let positional = path.path().steps().iter().all(|step| {
            step.predicates()
                .iter()
                .all(|predicate| matches!(predicate, litchi_core::selector::Predicate::Index(_)))
        });
        if !positional {
            return Ok(None);
        }
        let node = super::object(self, path)?;
        let kind = node.item().type_name();
        let format = self.format();
        if !props
            .iter()
            .all(|(key, _)| coalescable(format, kind, key.as_str()))
        {
            return Ok(None);
        }
        Ok(Some(node.index()))
    }
}

impl Run {
    pub(crate) fn new(base: &OdfTree) -> Self {
        Self {
            base: base.clone(),
            groups: Vec::new(),
            targets: HashSet::new(),
            first_index: None,
            operations: 0,
        }
    }

    /// The snapshot the run's sets resolve against.
    pub(crate) const fn base(&self) -> &OdfTree {
        &self.base
    }

    /// The wire index of the first staged set.
    pub(crate) const fn first_index(&self) -> Option<usize> {
        self.first_index
    }

    /// Whether a set of `node` must see the published result of the pending
    /// run, or would exceed its operation bound.
    pub(crate) fn requires_barrier(
        &self,
        node: u32,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> bool {
        !self.groups.is_empty()
            && (props.iter().any(|(key, _)| {
                self.targets
                    .contains(&(node, family(key.as_str()).to_owned()))
            }) || self.operations.saturating_add(props.len()) > limits.max_operations())
    }

    /// Stage a set admitted by [`Self::requires_barrier`].
    pub(crate) fn stage(&mut self, node: u32, props: &Props<'_>, index: usize) {
        self.first_index.get_or_insert(index);
        self.operations = self.operations.saturating_add(props.len());
        for (key, _) in props.iter() {
            self.targets.insert((node, family(key.as_str()).to_owned()));
        }
        self.groups.push((node, props.clone().into_owned()));
    }

    /// Publish the run as one leaf commit.
    pub(crate) fn commit(self, limits: &SemanticLimits) -> Result<(OdfTree, write::Commit)> {
        let groups: Vec<write::Group<'_, '_, '_>> = self
            .groups
            .iter()
            .map(|(node, props)| (vec![Node::new(&self.base, *node)], props))
            .collect();
        let commit =
            write::set_groups(&self.base, &groups, limits).map_err(|error| self.base.tag(error))?;
        drop(groups);
        Ok((self.base, commit))
    }
}
