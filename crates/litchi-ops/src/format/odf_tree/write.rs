//! Unified writes over the object formats' leaf edits.
//!
//! Every verb resolves its targets against one immutable handle's tree,
//! hands the flavor the nodes (with the leaf positions the tree recorded)
//! and the unified properties, and the flavor stages its leaf `Edit`
//! operations and commits them once. The committed package becomes a new
//! handle with a freshly built tree; an exact no-op returns the source
//! handle.

use litchi_core::FileFormat;
use litchi_core::selector::Path;

use super::{Flavor, Nav, Node, OdfTree, Owner};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::props::{IntoProps, PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
#[cfg(any(feature = "odg", feature = "odc", feature = "odi"))]
use crate::value::{Length, LengthSpec};
use crate::value::{Scalar, ValueType};

/// The leaf proof of one write.
#[allow(
    clippy::large_enum_variant,
    reason = "a write holds one proof; boxing would add an allocation per commit"
)]
pub(crate) enum Patch {
    /// The write changed nothing.
    Unchanged,
    #[cfg(feature = "odg")]
    Odg(litchi_odg::edit::Patch),
    #[cfg(feature = "odc")]
    Odc(litchi_odc::Patch),
    #[cfg(feature = "odi")]
    Odi(litchi_odi::Patch),
    #[cfg(feature = "odb")]
    Odb(litchi_odb::Patch),
    #[cfg(feature = "odf-formula")]
    Formula(litchi_odf_formula::Patch),
    /// A `/metadata` commit of the leaf metadata family (record 0288).
    Metadata(MetadataPatch),
}

/// The metadata-family proof of one object format.
pub(crate) enum MetadataPatch {
    #[cfg(feature = "odg")]
    Odg(litchi_odg::metadata::Patch),
    #[cfg(feature = "odc")]
    Odc(litchi_odc::metadata::Patch),
    #[cfg(feature = "odi")]
    Odi(litchi_odi::metadata::Patch),
    #[cfg(feature = "odb")]
    Odb(litchi_odb::metadata::Patch),
    #[cfg(feature = "odf-formula")]
    Formula(Box<litchi_odf_formula::metadata::Patch>),
}

impl MetadataPatch {
    fn is_noop(&self) -> bool {
        match self {
            #[cfg(feature = "odg")]
            Self::Odg(patch) => patch.is_noop(),
            #[cfg(feature = "odc")]
            Self::Odc(patch) => patch.is_noop(),
            #[cfg(feature = "odi")]
            Self::Odi(patch) => patch.is_noop(),
            #[cfg(feature = "odb")]
            Self::Odb(patch) => patch.is_noop(),
            #[cfg(feature = "odf-formula")]
            Self::Formula(patch) => patch.is_noop(),
        }
    }
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Unchanged => true,
            #[cfg(feature = "odg")]
            Self::Odg(patch) => patch.is_noop(),
            #[cfg(feature = "odc")]
            Self::Odc(patch) => patch.is_noop(),
            #[cfg(feature = "odi")]
            Self::Odi(patch) => patch.is_noop(),
            #[cfg(feature = "odb")]
            Self::Odb(patch) => patch.is_noop(),
            #[cfg(feature = "odf-formula")]
            Self::Formula(patch) => patch.is_noop(),
            Self::Metadata(patch) => patch.is_noop(),
        }
    }
}

/// One set of a write: the nodes it resolved and the properties it assigns.
pub(crate) type Group<'f, 'p, 'v> = (Vec<Node<'f>>, &'p Props<'v>);

/// The leaf proof and immutable endpoints of one write.
pub(crate) struct Commit {
    pub(crate) before: OdfTree,
    pub(crate) after: OdfTree,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

/// A refusal the caller's request caused, before the format is known.
pub(crate) fn refusal(code: Code, message: impl Into<std::borrow::Cow<'static, str>>) -> Error {
    Error::new(code, message)
}

/// A leaf staging failure: the caller's request, not the file, is at fault.
/// Authored invalid values and generic unsupported errors are recoded for the
/// write call site; a typed refusal keeps its `RefusalKind` classification.
pub(crate) fn staging_error(error: litchi_core::Error, format: FileFormat) -> Error {
    match error {
        error @ (litchi_core::Error::InvalidFormat(_) | litchi_core::Error::ParseError(_)) => {
            super::core_error(error, format).caused_by_argument(Code::InvalidValue)
        },
        error @ litchi_core::Error::Unsupported(_) => {
            super::core_error(error, format).caused_by_argument(Code::UnsupportedCapability)
        },
        other => super::core_error(other, format),
    }
}

/// A leaf commit failure. The source opened, so a candidate the leaf
/// refuses at publication (a failed readback or audit) is the edit's fault,
/// not a corrupt file; typed limits keep their code.
pub(crate) fn commit_error(error: litchi_odf_common::Error, format: FileFormat) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => staging_error(error, format),
        other => super::odf_error(other, format),
    }
}

pub(crate) fn bound(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TransactionError::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

/// Adopt a committed leaf package, charging the retained endpoint.
pub(crate) fn finish(
    file: &OdfTree,
    changed: bool,
    owner: Owner,
    patch: Patch,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    if !changed {
        return Ok(Commit {
            before: file.clone(),
            after: file.clone(),
            patch: Patch::Unchanged,
            operations,
            snapshot_bytes: 0,
        });
    }
    let snapshot_bytes = owner.as_bytes().len();
    bound(
        LimitKind::SnapshotBytes,
        snapshot_bytes,
        limits.max_snapshot_bytes(),
    )?;
    let after = file.republish(owner)?;
    Ok(Commit {
        before: file.clone(),
        after,
        patch,
        operations,
        snapshot_bytes,
    })
}

/// The nodes a set or remove targets: one path, or every match of a query.
pub(crate) fn resolve<'f, F: Flavor>(
    file: &'f OdfTree,
    target: &TargetArg<'_>,
    maximum: usize,
) -> Result<Vec<Node<'f>>> {
    let mut nodes = Vec::new();
    if let Some(path) = target.path() {
        nodes.push(crate::path::resolve::get(&Nav::<F>::new(file), path)?);
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::<F>::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(refusal(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            ));
        }
        for node in walk {
            nodes.push(node?);
            bound(LimitKind::ExpandedTargets, nodes.len(), maximum)?;
        }
        if nodes.is_empty() {
            return Err(refusal(
                Code::ZeroMatches,
                "the mutation selector matched nothing",
            ));
        }
    } else {
        return Err(refusal(
            Code::InvalidArgument,
            "the mutation target is empty",
        ));
    }
    bound(LimitKind::ExpandedTargets, nodes.len(), maximum)?;
    Ok(nodes)
}

/// Resolve the parent of an `add` (or any single element path).
pub(crate) fn parent<'f, F: Flavor>(file: &'f OdfTree, path: &PathArg<'_>) -> Result<Node<'f>> {
    crate::path::resolve::get(&Nav::<F>::new(file), path)
}

#[cfg(any(
    feature = "odg",
    feature = "odc",
    feature = "odb",
    feature = "odf-formula"
))]
/// The path of the `ordinal`th child of `kind` under `parent` in `file`'s
/// tree, after an add committed.
pub(crate) fn added_path(
    file: &OdfTree,
    parent: &[crate::node::PathStep<'_>],
    kind: &str,
    ordinal: Option<usize>,
) -> Result<Path<'static>> {
    let tree = file.tree();
    let parent_item = tree
        .items()
        .iter()
        .position(|item| item.path() == parent)
        .ok_or_else(|| {
            refusal(
                Code::InternalError,
                "the added element's parent disappeared",
            )
        })?;
    let children: Vec<u32> = tree
        .item(u32::try_from(parent_item).unwrap_or(u32::MAX))
        .children()
        .iter()
        .copied()
        .filter(|child| tree.item(*child).type_name() == kind)
        .collect();
    let index = match ordinal {
        Some(ordinal) => children.get(ordinal),
        None => children.last(),
    }
    .ok_or_else(|| {
        refusal(
            Code::InternalError,
            "the added element is missing after commit",
        )
    })?;
    Ok(crate::node::NodePath::new(tree.item(*index).path())
        .to_path()
        .into_owned())
}

/// The value of `key` coerced to text; `null` is `None`.
pub(crate) fn text(value: &PropValue<'_>, key: &'static str) -> Result<Option<String>> {
    value
        .with_coerced(ValueType::Text, |scalar| {
            Ok(scalar.map(|scalar| match scalar {
                Scalar::Text(text) => text.to_owned(),
                other => other.to_string(),
            }))
        })
        .map_err(|error| error.with_property(key))
}

#[cfg(any(feature = "odg", feature = "odc", feature = "odi"))]
/// The value of `key` coerced to a length; `null` is `None`.
pub(crate) fn length(
    value: &PropValue<'_>,
    key: &'static str,
    spec: LengthSpec,
) -> Result<Option<Length>> {
    value
        .with_coerced(ValueType::Length(spec), |scalar| {
            Ok(scalar.and_then(|scalar| match scalar {
                Scalar::Length(length) => Some(length),
                _ => None,
            }))
        })
        .map_err(|error| error.with_property(key))
}

#[cfg(feature = "odi")]
/// The value of `key` coerced to an integer; `null` is `None`.
pub(crate) fn int(value: &PropValue<'_>, key: &'static str) -> Result<Option<i64>> {
    value
        .with_coerced(ValueType::Int, |scalar| {
            Ok(scalar.and_then(|scalar| match scalar {
                Scalar::Int(value) => Some(value),
                _ => None,
            }))
        })
        .map_err(|error| error.with_property(key))
}

#[cfg(any(
    feature = "odg",
    feature = "odc",
    feature = "odb",
    feature = "odf-formula"
))]
/// A required (non-null) value.
pub(crate) fn required<T>(value: Option<T>, key: &'static str) -> Result<T> {
    value.ok_or_else(|| {
        refusal(Code::InvalidValue, "the property cannot be cleared")
            .with_property(key.to_owned())
            .with_suggestion("set a value; clearing this property is not supported")
    })
}

/// The refusal of a key the element does not write.
pub(crate) fn unwritable(element: &'static str, key: &str) -> Error {
    let readable = false;
    let _ = readable;
    refusal(
        Code::UnsupportedProperty,
        format!("`{key}` is not a writable property of `{element}`"),
    )
    .with_property(key.to_owned())
    .with_suggestion("list the writable properties with capabilities()")
}

/// The refusal of a read-only property, or an unknown key.
pub(crate) fn key_refusal(node: &Node<'_>, key: &str) -> Error {
    let item = node.item();
    if item
        .props()
        .iter()
        .any(|prop| prop.name.eq_ignore_ascii_case(key))
    {
        refusal(
            Code::UnsupportedCapability,
            format!("`{key}` of `{}` is read-only", item.type_name()),
        )
        .with_property(key.to_owned())
    } else {
        unwritable(item.type_name(), key)
    }
}

/// The refusal of a verb on an element the flavor does not write.
pub(crate) fn element_refusal(verb: &'static str, node: &Node<'_>) -> Error {
    refusal(
        Code::UnsupportedCapability,
        format!(
            "`{verb}` is not supported on `{}` elements",
            node.item().type_name()
        ),
    )
}

/// Set properties on every target.
pub(crate) fn set<'p, 'v>(
    file: &OdfTree,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let props = Props::from_props(props)?;
    if props.is_empty() {
        return Err(refusal(
            Code::InvalidArgument,
            "a set needs at least one property",
        ));
    }
    if crate::format::odf_metadata::is_target(&target) {
        return set_metadata(file, &props, limits);
    }
    let maximum = limits.max_expanded_targets();
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => {
            let nodes = resolve::<super::odg::Odg>(file, &target, maximum)?;
            super::odg::set(file, package, &[(nodes, &props)], limits)
        },
        #[cfg(feature = "odc")]
        Owner::Odc(package) => {
            let nodes = resolve::<super::odc::Odc>(file, &target, maximum)?;
            super::odc::set(file, package, &[(nodes, &props)], limits)
        },
        #[cfg(feature = "odi")]
        Owner::Odi(package) => {
            let nodes = resolve::<super::odi::Odi>(file, &target, maximum)?;
            super::odi::set(file, package, &[(nodes, &props)], limits)
        },
        #[cfg(feature = "odb")]
        Owner::Odb(package) => {
            let nodes = resolve::<super::odb::Odb>(file, &target, maximum)?;
            super::odb::set(file, package, &[(nodes, &props)], limits)
        },
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => {
            let nodes = resolve::<super::formula::Formula>(file, &target, maximum)?;
            super::formula::set(file, package, &[(nodes, &props)], limits)
        },
    }
}

/// Stage every group of a set (or of a coalesced run of sets) on one leaf
/// edit and commit once.
pub(crate) fn set_groups(
    file: &OdfTree,
    groups: &[Group<'_, '_, '_>],
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => super::odg::set(file, package, groups, limits),
        #[cfg(feature = "odc")]
        Owner::Odc(package) => super::odc::set(file, package, groups, limits),
        #[cfg(feature = "odi")]
        Owner::Odi(package) => super::odi::set(file, package, groups, limits),
        #[cfg(feature = "odb")]
        Owner::Odb(package) => super::odb::set(file, package, groups, limits),
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => super::formula::set(file, package, groups, limits),
    }
}

/// Set `/metadata` through the leaf metadata family (record 0288).
fn set_metadata(file: &OdfTree, props: &Props<'_>, limits: &SemanticLimits) -> Result<Commit> {
    let format = file.format();
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => metadata_commit(file, package, props, limits, |package, patch| {
            (Owner::Odg(package), MetadataPatch::Odg(patch))
        }),
        #[cfg(feature = "odc")]
        Owner::Odc(package) => metadata_commit(file, package, props, limits, |package, patch| {
            (Owner::Odc(package), MetadataPatch::Odc(patch))
        }),
        #[cfg(feature = "odi")]
        Owner::Odi(package) => metadata_commit(file, package, props, limits, |package, patch| {
            (Owner::Odi(package), MetadataPatch::Odi(patch))
        }),
        #[cfg(feature = "odb")]
        Owner::Odb(package) => metadata_commit(file, package, props, limits, |package, patch| {
            (Owner::Odb(package), MetadataPatch::Odb(patch))
        }),
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => {
            metadata_commit(file, package, props, limits, |package, patch| {
                (
                    Owner::Formula(package),
                    MetadataPatch::Formula(Box::new(patch)),
                )
            })
        },
    }
    .map_err(|error| error.with_format(format))
}

/// Stage and commit `props` on `package`'s metadata family.
fn metadata_commit<P: litchi_odf_common::metadata::Owner>(
    file: &OdfTree,
    package: &P,
    props: &Props<'_>,
    limits: &SemanticLimits,
    adopt: fn(P, litchi_odf_common::metadata::Patch<P>) -> (Owner, MetadataPatch),
) -> Result<Commit> {
    let format = file.format();
    let mut edit = litchi_odf_common::metadata::Edit::new(package);
    crate::format::odf_metadata::stage(&mut edit, props, format, limits)?;
    let (next, patch, changed) = crate::format::odf_metadata::commit(edit, format)?;
    let (owner, patch) = adopt(next, patch);
    finish(
        file,
        changed,
        owner,
        Patch::Metadata(patch),
        props.len(),
        limits,
    )
}

/// Add one element.
pub(crate) fn add<'p, 'n>(
    file: &OdfTree,
    parent: PathArg<'p>,
    node: NewNode<'n>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => super::odg::add(file, package, &parent, &node, limits),
        #[cfg(feature = "odc")]
        Owner::Odc(package) => super::odc::add(file, package, &parent, &node, limits),
        #[cfg(feature = "odi")]
        Owner::Odi(package) => super::odi::add(file, package, &parent, &node, limits),
        #[cfg(feature = "odb")]
        Owner::Odb(package) => super::odb::add(file, package, &parent, &node, limits),
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => super::formula::add(file, package, &parent, &node, limits),
    }
}

/// Remove every target.
pub(crate) fn remove<'p>(
    file: &OdfTree,
    target: TargetArg<'p>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => super::odg::remove(file, package, &target, limits),
        #[cfg(feature = "odc")]
        Owner::Odc(package) => super::odc::remove(file, package, &target, limits),
        #[cfg(feature = "odi")]
        Owner::Odi(package) => super::odi::remove(file, package, &target, limits),
        #[cfg(feature = "odb")]
        Owner::Odb(package) => super::odb::remove(file, package, &target, limits),
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => super::formula::remove(file, package, &target, limits),
    }
}

/// Move one element.
pub(crate) fn move_to<'p, 'q>(
    file: &OdfTree,
    path: PathArg<'p>,
    placement: crate::Placement<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => super::odg::move_to(file, package, &path, &placement, limits),
        #[cfg(feature = "odc")]
        Owner::Odc(package) => super::odc::move_to(file, package, &path, &placement, limits),
        #[cfg(feature = "odi")]
        Owner::Odi(_) => super::odi::move_to(file, &path, &placement, limits),
        #[cfg(feature = "odb")]
        Owner::Odb(package) => super::odb::move_to(file, package, &path, &placement, limits),
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => {
            super::formula::move_to(file, package, &path, &placement, limits)
        },
    }
}

/// Exchange two elements.
pub(crate) fn swap<'p, 'q>(
    file: &OdfTree,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => super::odg::swap(file, package, &left, &right, limits),
        #[cfg(feature = "odc")]
        Owner::Odc(package) => super::odc::swap(file, package, &left, &right, limits),
        #[cfg(feature = "odi")]
        Owner::Odi(_) => super::odi::swap(file, &left, &right, limits),
        #[cfg(feature = "odb")]
        Owner::Odb(package) => super::odb::swap(file, package, &left, &right, limits),
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => super::formula::swap(file, package, &left, &right, limits),
    }
}

/// The same-kind siblings of one element: its parent item, the sibling
/// items in document order, and the element's position among them.
pub(crate) struct Siblings {
    pub(crate) parent: u32,
    pub(crate) items: Vec<u32>,
    pub(crate) from: usize,
}

/// The same-kind siblings of `node` in `file`'s tree.
pub(crate) fn siblings(file: &OdfTree, node: &Node<'_>) -> Result<Siblings> {
    let tree = file.tree();
    let kind = node.item().type_name();
    let parent = (0..tree.len())
        .map(|index| u32::try_from(index).unwrap_or(u32::MAX))
        .find(|index| tree.item(*index).children().contains(&node.index()))
        .ok_or_else(|| {
            refusal(
                Code::UnsupportedCapability,
                format!("the root `{kind}` element has no siblings to reorder"),
            )
        })?;
    let items: Vec<u32> = tree
        .item(parent)
        .children()
        .iter()
        .copied()
        .filter(|child| tree.item(*child).type_name() == kind)
        .collect();
    let from = items
        .iter()
        .position(|child| *child == node.index())
        .ok_or_else(|| refusal(Code::InternalError, "the moved element is not a child"))?;
    Ok(Siblings {
        parent,
        items,
        from,
    })
}

/// Where a move places an element: among its own siblings (its final
/// sibling position), or under another parent.
#[cfg_attr(
    not(any(feature = "odg", feature = "odf-formula")),
    allow(
        dead_code,
        reason = "only the drawing and formula flavors move elements across parents"
    )
)]
pub(crate) enum Destination<'f> {
    /// The final position among the element's current siblings.
    #[cfg_attr(
        not(any(
            feature = "odg",
            feature = "odc",
            feature = "odb",
            feature = "odf-formula"
        )),
        allow(dead_code, reason = "a single-element image has no other position")
    )]
    Within(usize),
    /// Under `parent` (another item), before the same-kind sibling at
    /// `position` of that parent's children (`None`: last).
    Across {
        parent: Node<'f>,
        position: Option<usize>,
    },
}

/// Resolve `placement` for the element at `siblings.from`.
pub(crate) fn destination<'f, F: Flavor>(
    file: &'f OdfTree,
    node: &Node<'f>,
    siblings: &Siblings,
    placement: &crate::Placement<'_>,
) -> Result<Destination<'f>> {
    use crate::Placement;
    let count = siblings.items.len();
    let last = count.saturating_sub(1);
    let kind = node.item().type_name();
    let within = |to: usize| -> Result<Destination<'f>> {
        if to >= count {
            return Err(refusal(
                Code::OutOfRange,
                format!("the move position is outside the `{kind}` siblings"),
            ));
        }
        Ok(Destination::Within(to))
    };
    let anchored = |path: &PathArg<'_>, after: bool| -> Result<Destination<'f>> {
        let anchor = parent::<F>(file, path)?;
        if anchor.item().type_name() != kind {
            return Err(refusal(
                Code::UnsupportedType,
                format!("the placement anchor is not a `{kind}` element"),
            ));
        }
        if let Some(index) = siblings
            .items
            .iter()
            .position(|item| *item == anchor.index())
        {
            let from = siblings.from;
            let to = match (after, from < index) {
                (false, true) => index - 1,
                (false, false) | (true, true) => index,
                (true, false) => index + 1,
            };
            return within(to);
        }
        let anchor_siblings = self::siblings(file, &anchor)?;
        let position = anchor_siblings.from + usize::from(after);
        Ok(Destination::Across {
            parent: Node::new(file, anchor_siblings.parent),
            position: Some(position),
        })
    };
    match placement {
        Placement::Start => within(0),
        Placement::End => within(last),
        Placement::Position(position) => within(position.get()),
        Placement::Before(path) => anchored(path, false),
        Placement::After(path) => anchored(path, true),
        Placement::Into(path, position) => {
            let target = parent::<F>(file, path)?;
            if target.index() == siblings.parent {
                within(position.map_or(last, litchi_core::Position::get))
            } else {
                Ok(Destination::Across {
                    parent: target,
                    position: position.map(litchi_core::Position::get),
                })
            }
        },
    }
}

/// Resolve both sides of a swap; they must be distinct elements of one kind.
pub(crate) fn swap_pair<'f, F: Flavor>(
    file: &'f OdfTree,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
) -> Result<(Node<'f>, Node<'f>)> {
    let left = parent::<F>(file, left)?;
    let right = parent::<F>(file, right)?;
    if left.item().type_name() != right.item().type_name() {
        return Err(refusal(
            Code::UnsupportedType,
            "a swap exchanges two elements of the same kind",
        ));
    }
    Ok((left, right))
}

/// A fresh minimal package of `format`, with create-time metadata where
/// the format writes metadata.
pub(crate) fn create(format: FileFormat, options: &crate::CreateOptions) -> Result<OdfTree> {
    use crate::{Conformance, Flavor as CreateFlavor};
    if options.flavor_value() != CreateFlavor::Document {
        return Err(refusal(
            Code::UnsupportedCapability,
            "template creation is not supported for this OpenDocument format",
        )
        .with_format(format));
    }
    if options.conformance_value() == Conformance::Strict {
        return Err(refusal(
            Code::UnsupportedCapability,
            "OpenDocument formats have no OOXML Strict conformance",
        )
        .with_format(format));
    }
    let owner = match format {
        #[cfg(feature = "odg")]
        FileFormat::Odg => {
            Owner::Odg(litchi_odg::Package::new().map_err(|error| super::odf_error(error, format))?)
        },
        #[cfg(feature = "odc")]
        FileFormat::Odc => {
            Owner::Odc(litchi_odc::Package::new().map_err(|error| super::odf_error(error, format))?)
        },
        #[cfg(feature = "odi")]
        // A fresh image keeps an empty `meta.xml`: the leaf edits metadata
        // only in an existing part, and `/metadata` is settable.
        FileFormat::Odi => Owner::Odi(
            litchi_odi::Builder::new()
                .meta_xml(EMPTY_ODI_META)
                .build()
                .map_err(|error| super::odf_error(error, format))?,
        ),
        #[cfg(feature = "odb")]
        FileFormat::Odb => {
            Owner::Odb(litchi_odb::Package::new().map_err(|error| super::odf_error(error, format))?)
        },
        #[cfg(feature = "odf-formula")]
        FileFormat::Odf => Owner::Formula(
            litchi_odf_formula::Package::new().map_err(|error| super::odf_error(error, format))?,
        ),
        other => return Err(Error::unsupported_format(Some(other))),
    };
    let file = OdfTree::from_owner(owner, None, litchi_odf_common::OpenOptions::new())?;
    match options.metadata_value() {
        None => Ok(file),
        Some(metadata) => create_metadata(file, metadata),
    }
}

/// The `meta.xml` of a fresh image: an empty `office:meta` whose
/// namespaces the leaf's metadata edit writes into.
#[cfg(feature = "odi")]
const EMPTY_ODI_META: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" "#,
    r#"xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.4"><office:meta/></office:document-meta>"#,
);

/// Commit create-time metadata through the leaf metadata family.
fn create_metadata(file: OdfTree, metadata: &litchi_core::Metadata) -> Result<OdfTree> {
    let format = file.format();
    match file.owner() {
        #[cfg(feature = "odg")]
        Owner::Odg(package) => created(package, metadata, format, Owner::Odg),
        #[cfg(feature = "odc")]
        Owner::Odc(package) => created(package, metadata, format, Owner::Odc),
        #[cfg(feature = "odi")]
        Owner::Odi(package) => created(package, metadata, format, Owner::Odi),
        #[cfg(feature = "odb")]
        Owner::Odb(package) => created(package, metadata, format, Owner::Odb),
        #[cfg(feature = "odf-formula")]
        Owner::Formula(package) => created(package, metadata, format, Owner::Formula),
    }
}

/// A fresh handle over `package` with create-time `metadata` committed.
fn created<P: litchi_odf_common::metadata::Owner>(
    package: &P,
    metadata: &litchi_core::Metadata,
    format: FileFormat,
    wrap: fn(P) -> Owner,
) -> Result<OdfTree> {
    let mut edit = litchi_odf_common::metadata::Edit::new(package);
    crate::format::odf_metadata::stage_created(&mut edit, metadata, format)?;
    let (next, _, _) = crate::format::odf_metadata::commit(edit, format)?;
    OdfTree::from_owner(wrap(next), None, litchi_odf_common::OpenOptions::new())
}
