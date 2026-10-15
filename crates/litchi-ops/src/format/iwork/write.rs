//! Unified writes over the iWork leaves' focused transactions.
//!
//! Every verb resolves its targets against one immutable handle's tree,
//! hands the flavor the nodes (with the leaf positions the tree recorded)
//! and the unified properties, and the flavor runs its leaf transactions in
//! order, each over the package the previous one published. The last
//! package becomes a new handle whose tree is built on its first read; an
//! exact no-op returns the source handle.

#[cfg(any(feature = "keynote", feature = "pages", feature = "numbers"))]
use litchi_core::Position;
use litchi_core::selector::Path;

use super::{Flavor, Iwork, Nav, Node, Owner};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
use crate::props::PropValue;
use crate::props::{IntoProps, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
use crate::value::{Scalar, ValueType};

/// One leaf transaction's reversible proof.
#[allow(
    clippy::large_enum_variant,
    reason = "a write holds its proofs once; boxing would add an allocation per commit"
)]
#[allow(
    clippy::enum_variant_names,
    reason = "in a one-format build the remaining variants share that format's prefix"
)]
pub(crate) enum LeafPatch {
    #[cfg(feature = "pages")]
    Pages(super::pages::Patch),
    #[cfg(feature = "keynote")]
    Keynote(super::keynote::Patch),
    #[cfg(feature = "numbers")]
    NumbersCells(litchi_numbers::table::cells::Patch),
    /// One of the other Numbers transactions (record 0297).
    #[cfg(feature = "numbers")]
    Numbers(super::numbers::Patch),
}

impl LeafPatch {
    fn is_noop(&self) -> bool {
        match self {
            #[cfg(feature = "pages")]
            Self::Pages(patch) => patch.is_noop(),
            #[cfg(feature = "keynote")]
            Self::Keynote(patch) => patch.is_noop(),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(patch) => patch.is_noop(),
            #[cfg(feature = "numbers")]
            Self::Numbers(patch) => patch.is_noop(),
        }
    }
}

/// The leaf proof of one write: the leaf transactions it ran, in order.
pub(crate) enum Patch {
    /// The write changed nothing.
    Unchanged,
    /// The transactions, each over the package the previous one published.
    Steps(Vec<LeafPatch>),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Unchanged => true,
            Self::Steps(steps) => steps.iter().all(LeafPatch::is_noop),
        }
    }
}

/// The leaf proof and immutable endpoints of one write.
pub(crate) struct Commit {
    pub(crate) before: Iwork,
    pub(crate) after: Iwork,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

/// A refusal the caller's request caused, before the format is known.
pub(crate) fn refusal(code: Code, message: impl Into<std::borrow::Cow<'static, str>>) -> Error {
    Error::new(code, message)
}

/// A leaf transaction failure: the request's fault unless the leaf names
/// the file or a resource ceiling.
pub(crate) fn leaf_refusal(error: super::LeafError, format: litchi_core::FileFormat) -> Error {
    super::leaf_error(error, format)
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

/// The running state of one write: the package the next leaf transaction
/// starts from and the proofs so far.
pub(crate) struct Chain<P> {
    pub(crate) package: P,
    pub(crate) patches: Vec<LeafPatch>,
    pub(crate) changed: bool,
}

impl<P> Chain<P> {
    pub(crate) const fn new(package: P) -> Self {
        Self {
            package,
            patches: Vec::new(),
            changed: false,
        }
    }

    /// Record one committed transaction.
    pub(crate) fn push(&mut self, package: P, patch: LeafPatch, changed: bool) {
        self.package = package;
        self.patches.push(patch);
        self.changed |= changed;
    }
}

/// Counts bytes without keeping them.
struct Count(usize);

impl std::io::Write for Count {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.0 = self.0.saturating_add(buffer.len());
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// Adopt the last committed leaf package, charging the retained endpoint.
pub(crate) fn finish(
    file: &Iwork,
    changed: bool,
    owner: Owner,
    patches: Vec<LeafPatch>,
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
    let mut count = Count(0);
    owner.write_to(&mut count)?;
    let snapshot_bytes = count.0;
    bound(
        LimitKind::SnapshotBytes,
        snapshot_bytes,
        limits.max_snapshot_bytes(),
    )?;
    let after = Iwork::from_owner(owner, file.origin().cloned());
    Ok(Commit {
        before: file.clone(),
        after,
        patch: Patch::Steps(patches),
        operations,
        snapshot_bytes,
    })
}

/// The nodes a set or remove targets: one path, or every match of a query.
pub(crate) fn resolve<'f, F: Flavor>(
    file: &'f Iwork,
    target: &TargetArg<'_>,
    maximum: usize,
) -> Result<Vec<Node<'f>>> {
    file.checked_tree()?;
    let mut nodes = Vec::new();
    if let Some(path) = target.path() {
        nodes.push(crate::path::resolve::get(&Nav::<F>::new(file), path)?);
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::<F>::new(file), query.clone())?;
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

/// Resolve one element by path.
#[cfg(any(feature = "keynote", feature = "numbers", feature = "pages"))]
pub(crate) fn node_at<'f, F: Flavor>(file: &'f Iwork, path: &PathArg<'_>) -> Result<Node<'f>> {
    file.checked_tree()?;
    crate::path::resolve::get(&Nav::<F>::new(file), path)
}

/// The value of `key` coerced to text; `null` is `None`.
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
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

/// The value of `key` coerced to a boolean; `null` is refused.
#[cfg(feature = "keynote")]
pub(crate) fn boolean(value: &PropValue<'_>, key: &'static str) -> Result<bool> {
    value
        .with_coerced(ValueType::Bool, |scalar| {
            Ok(scalar.and_then(|scalar| match scalar {
                Scalar::Bool(value) => Some(value),
                _ => None,
            }))
        })
        .map_err(|error| error.with_property(key))?
        .ok_or_else(|| {
            refusal(Code::InvalidValue, "the property cannot be cleared")
                .with_property(key.to_owned())
        })
}

/// The refusal of a key the element does not write.
pub(crate) fn unwritable(element: &'static str, key: &str) -> Error {
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

/// The zero-based destination of a move among `count` siblings, for an
/// element now at `from`; `anchor` resolves a `before`/`after` path to its
/// sibling position.
#[cfg(any(feature = "keynote", feature = "pages", feature = "numbers"))]
pub(crate) fn destination(
    from: usize,
    count: usize,
    placement: &crate::Placement<'_>,
    anchor: impl Fn(&PathArg<'_>) -> Result<usize>,
) -> Result<usize> {
    use crate::Placement;
    let last = count.saturating_sub(1);
    let to = match placement {
        Placement::Start => 0,
        Placement::End => last,
        Placement::Position(position) => position.get(),
        Placement::Before(path) => {
            let anchor = anchor(path)?;
            if from < anchor { anchor - 1 } else { anchor }
        },
        Placement::After(path) => {
            let anchor = anchor(path)?;
            if from < anchor { anchor } else { anchor + 1 }
        },
        Placement::Into(_, position) => position.map_or(last, Position::get),
    };
    if to >= count {
        return Err(refusal(
            Code::OutOfRange,
            "the move position is outside the element's siblings",
        ));
    }
    Ok(to)
}

/// Set properties on every target.
pub(crate) fn set<'p, 'v>(
    file: &Iwork,
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
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::set(file, package, &target, &props, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::set(file, package, &target, &props, limits),
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => super::numbers::set(file, package, &target, &props, limits),
    }
}

/// Add one element (Pages paragraphs; Numbers merged ranges, record 0297).
pub(crate) fn add<'p, 'n>(
    file: &Iwork,
    parent: PathArg<'p>,
    node: NewNode<'n>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::add(file, package, &parent, &node, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::add(file, package, &parent, &node, limits),
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => super::numbers::add(file, package, &parent, &node, limits),
        #[allow(
            unreachable_patterns,
            reason = "every iWork leaf adds elements; the arm covers a build without them"
        )]
        _ => {
            let _ = (parent, node, limits);
            Err(refusal(
                Code::UnsupportedCapability,
                "this build cannot add elements to the format",
            ))
        },
    }
}

/// Remove every target.
pub(crate) fn remove<'p>(
    file: &Iwork,
    target: TargetArg<'p>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::remove(file, package, &target, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::remove(file, package, &target, limits),
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => super::numbers::remove(file, package, &target, limits),
        #[allow(unreachable_patterns, reason = "every iWork leaf removes elements")]
        _ => {
            let _ = (target, limits);
            let _ = file;
            Err(refusal(
                Code::UnsupportedCapability,
                "this iWork element cannot be removed",
            ))
        },
    }
}

/// Move one element (Keynote slides).
pub(crate) fn move_to<'p, 'q>(
    file: &Iwork,
    path: PathArg<'p>,
    placement: crate::Placement<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::move_to(file, package, &path, &placement, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => {
            super::keynote::move_to(file, package, &path, &placement, limits)
        },
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => {
            super::numbers::move_to(file, package, &path, &placement, limits)
        },
        #[allow(unreachable_patterns, reason = "every iWork leaf orders elements")]
        _ => {
            let _ = (path, placement, limits);
            Err(refusal(
                Code::UnsupportedCapability,
                "this iWork element cannot be moved",
            ))
        },
    }
}

/// Exchange two elements (Pages paragraphs).
pub(crate) fn swap<'p, 'q>(
    file: &Iwork,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::swap(file, package, &left, &right, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::swap(file, package, &left, &right, limits),
        #[allow(
            unreachable_patterns,
            reason = "the Pages and Keynote leaves exchange elements; Numbers does not yet"
        )]
        _ => {
            let _ = (left, right, limits);
            Err(refusal(
                Code::UnsupportedCapability,
                "swapping Numbers elements is not wired yet (follow-up: record 0297)",
            ))
        },
    }
}
