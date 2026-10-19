//! Batch coalescing of Keynote writes (record 0298).
//!
//! Consecutive coalescible items of one `File::edit` or `batch` share one
//! [`Run`]: text, paragraph, frame and visibility sets, and new text boxes,
//! shapes and pictures, publish through one slide-drawable commit and one
//! commit per changed text; consecutive new slides through one
//! slide-creation commit. A run holds either new slides or the other
//! changes, never both: a new slide moves every later slide, so the paths
//! of later sets would no longer name what the run's base tree names. Any
//! other operation, a read, or an item the run cannot plan publishes the
//! run first, so every item still sees the effects of the items before it.

use litchi_core::FileFormat;
use litchi_core::selector::{Path, Step};

use super::super::write::{Chain, Commit, finish, refusal};
use super::super::{At, Iwork, Nav, Node, Owner};
use super::Keynote;
use super::write::{Pending, SlideSource, drawable_request, insertion, slide_at, slide_position};
use crate::error::{Code, Error};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::{Placement, Result};

/// Whether `key` of an element of type `kind` is a set a run stages.
fn coalescible(kind: &str, key: &str) -> bool {
    matches!(
        (kind, key),
        ("slide", "hidden")
            | ("placeholder" | "notes" | "shape" | "paragraph", "text")
            | (
                "placeholder" | "shape" | "picture",
                "x" | "y" | "width" | "height" | "rotation"
            )
    )
}

/// The tree item a set of `props` on `target` stages, when a run can stage
/// it: a direct path to an element whose every key coalesces.
pub(crate) fn plan_set(file: &Iwork, target: &TargetArg<'_>, props: &Props<'_>) -> Option<At> {
    let path = target.path()?;
    file.checked_tree().ok()?;
    let node = crate::path::resolve::get(&Nav::<Keynote>::new(file), path).ok()?;
    let kind = node.item().type_name();
    (!props.is_empty() && props.iter().all(|(key, _)| coalescible(kind, key.as_str())))
        .then_some(node.at())
}

/// The zero-based slide a plain positional `/slide[N]` path names.
fn evolving_slide(path: &str) -> Option<usize> {
    let number = path
        .trim()
        .strip_prefix("/slide[")?
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()?;
    number.checked_sub(1)
}

/// Whether `held` and `path` are equal, compared from the last step: a
/// batch's distinct targets differ there, so a miss in the run's memo costs
/// one step comparison per entry instead of a whole path (record 0340).
fn same_path(held: &Path<'_>, path: &Path<'_>) -> bool {
    held.is_absolute() == path.is_absolute()
        && held.steps().len() == path.steps().len()
        && held.last().map(Step::index) == path.last().map(Step::index)
        && held.steps().iter().rev().eq(path.steps().iter().rev())
}

/// A parent the run resolved, and the paragraph step it resolved under it.
#[derive(Debug)]
struct Parent {
    /// The parent's path.
    path: Path<'static>,
    at: At,
    /// The name of a positional `name[N]` step that resolved, under this
    /// parent, to the paragraph the parent's N-th paragraph child is
    /// (record 0340): later steps of that name and form take that child
    /// directly.
    paragraphs: Option<String>,
}

/// Consecutive coalesced writes of one transaction.
#[derive(Debug)]
pub(crate) struct Run {
    base: Iwork,
    pending: Pending,
    /// Paths the run resolved against its base, with the item and its type
    /// (record 0314): a batch often sets the same few elements again.
    resolved: Vec<(Path<'static>, At, &'static str)>,
    /// Parents of the paths the run resolved (all but the last step), with
    /// their item (record 0340): a batch over one text's paragraphs resolves
    /// only each path's last step.
    parents: Vec<Parent>,
    /// New slides: source, destination among the evolving slides, item.
    slides: Vec<(SlideSource, usize, usize)>,
    /// The evolving slide list: each slide's base position, `None` for a
    /// new slide.
    order: Vec<Option<usize>>,
    limits: SemanticLimits,
    first_index: Option<usize>,
    operations: usize,
}

impl Run {
    /// A run over the immutable handle `base`.
    pub(crate) fn new(base: &Iwork, limits: &SemanticLimits) -> Self {
        let slides = base.kids(At::ROOT).len();
        Self {
            base: base.clone(),
            pending: Pending::default(),
            resolved: Vec::new(),
            parents: Vec::new(),
            slides: Vec::new(),
            order: (0..slides).map(Some).collect(),
            limits: *limits,
            first_index: None,
            operations: 0,
        }
    }

    /// The transaction item index of the run's first item.
    pub(crate) const fn first_index(&self) -> Option<usize> {
        self.first_index
    }

    /// Whether the run can stage sets and new drawables (it holds no new
    /// slide).
    pub(crate) fn accepts_sets(&self) -> bool {
        self.slides.is_empty()
    }

    fn admit(&mut self, index: usize) -> Result<()> {
        let operations = self.operations.saturating_add(1);
        super::super::write::bound(
            LimitKind::Operations,
            operations,
            self.limits.max_operations(),
        )?;
        self.operations = operations;
        self.first_index.get_or_insert(index);
        Ok(())
    }

    /// [`plan_set`] against the run's base, through its memo of resolved
    /// paths.
    pub(crate) fn plan_set(&mut self, target: &TargetArg<'_>, props: &Props<'_>) -> Option<At> {
        /// Paths the memo keeps.
        const RESOLVED: usize = 16;
        let path = target.path()?;
        let held = self
            .resolved
            .iter()
            .find(|(held, _, _)| same_path(held, path.path()))
            .map(|(_, at, kind)| (*at, *kind));
        let (at, kind) = match held {
            Some(found) => found,
            None => {
                let found = self.resolve(path)?;
                if self.resolved.len() < RESOLVED {
                    self.resolved
                        .push((path.path().clone().into_owned(), found.0, found.1));
                }
                found
            },
        };
        (!props.is_empty() && props.iter().all(|(key, _)| coalescible(kind, key.as_str())))
            .then_some(at)
    }

    /// Resolve `path` against the run's base: from a memoized parent when
    /// the path's steps but the last name one, else from the root (keeping
    /// the parent for later paths).
    fn resolve(&mut self, path: &PathArg<'_>) -> Option<(At, &'static str)> {
        /// Parents the memo keeps.
        const PARENTS: usize = 4;
        let nav = Nav::<Keynote>::new(&self.base);
        let steps = path.path().steps();
        let split = (path.path().is_absolute() && steps.len() >= 2)
            .then(|| steps.split_last())
            .flatten();
        let held = split.and_then(|(_, prefix)| {
            self.parents
                .iter()
                .position(|parent| parent.path.steps().iter().rev().eq(prefix.iter().rev()))
        });
        if let (Some((last, _)), Some(held)) = (split, held) {
            let parent = &self.parents[held];
            if parent.paragraphs.as_deref() == Some(last.name())
                && let Some(position) = last.index()
                && let Some(at) = self.paragraph(parent.at, position.get())
            {
                return Some((at, "paragraph"));
            }
            let start = Node::new(&self.base, parent.at);
            let kind = crate::query::inspect::Inspect::kind(&nav, &start);
            let step = PathArg::from(Path::relative([last.clone()]));
            if let Ok(node) = crate::path::resolve::get_from(&nav, start, kind, &step) {
                let found = (node.at(), node.item().type_name());
                self.learn(held, last, found);
                return Some(found);
            }
        }
        self.base.checked_tree().ok()?;
        let node = crate::path::resolve::get(&nav, path).ok()?;
        let found = (node.at(), node.item().type_name());
        if let Some((last, prefix)) = split
            && held.is_none()
            && self.parents.len() < PARENTS
        {
            let parent = PathArg::from(Path::from_steps(prefix.iter().cloned()));
            if let Ok(at) = crate::path::resolve::get(&nav, &parent).map(|node| node.at()) {
                self.parents.push(Parent {
                    path: Path::from_steps(prefix.iter().cloned()).into_owned(),
                    at,
                    paragraphs: None,
                });
                self.learn(self.parents.len() - 1, last, found);
            }
        }
        Some(found)
    }

    /// The `position`-th (zero-based) paragraph child of `parent`, when the
    /// paragraphs are contiguous from the first of them.
    fn paragraph(&self, parent: At, position: usize) -> Option<At> {
        let ordinal = crate::path::navigate::ElementKind::ordinal(super::Kind::Paragraph);
        let kids = self.base.kids(parent);
        let first = (0..kids.len()).find(|slot| {
            kids.get(*slot)
                .is_some_and(|child| self.base.kind_at(child) == ordinal)
        })?;
        kids.get(first + position).filter(|child| {
            self.base.kind_at(*child) == ordinal
                && usize::try_from(self.base.ordinal_at(*child)).is_ok_and(|held| held == position)
        })
    }

    /// Remember that `last` resolved to `found` under parent `held`: a
    /// positional step that named the paragraph [`Self::paragraph`] finds.
    fn learn(&mut self, held: usize, last: &Step<'_>, found: (At, &'static str)) {
        let Some(parent) = self.parents.get(held) else {
            return;
        };
        if parent.paragraphs.is_some() || found.1 != "paragraph" {
            return;
        }
        let Some(position) = last.index() else {
            return;
        };
        if self.paragraph(parent.at, position.get()) == Some(found.0) {
            let name = last.name().to_owned();
            if let Some(parent) = self.parents.get_mut(held) {
                parent.paragraphs = Some(name);
            }
        }
    }

    /// Whether a set of `props` on tree item `node` must wait for the run to
    /// publish: joined to the run it would not publish what it publishes
    /// after the run's items, one at a time (record 0314).
    pub(crate) fn requires_barrier(&mut self, node: At, props: &Props<'_>) -> bool {
        let Self { base, pending, .. } = self;
        let node = Node::new(base, node);
        !pending.joins(base, &node, props)
    }

    /// Stage the set planned by [`plan_set`] on tree item `node`.
    ///
    /// # Errors
    ///
    /// A refused value or key, or the run's operation limit.
    pub(crate) fn stage_set(&mut self, node: At, props: &Props<'_>, index: usize) -> Result<()> {
        if !self.accepts_sets() {
            return Err(refusal(
                Code::InternalError,
                "a Keynote run with new slides staged a set",
            ));
        }
        self.admit(index)?;
        let Self { base, pending, .. } = self;
        let node = Node::new(base, node);
        for (key, value) in props.iter() {
            pending.stage_set(base, &node, key.as_str(), value, index)?;
        }
        Ok(())
    }

    /// Stage a new slide, text box, shape or picture, returning its path;
    /// `Ok(None)` when the run cannot stage it (the caller publishes the run
    /// and adds it alone). Nothing is staged on `Ok(None)` or an error.
    ///
    /// # Errors
    ///
    /// A refused property of the new element, or the run's operation limit.
    pub(crate) fn stage_add(
        &mut self,
        parent: &PathArg<'_>,
        node: &NewNode<'_>,
        index: usize,
    ) -> Result<Option<Path<'static>>> {
        if self.base.checked_tree().is_err() {
            return Ok(None);
        }
        let Ok(parent) = crate::path::resolve::get(&Nav::<Keynote>::new(&self.base), parent) else {
            return Ok(None);
        };
        let kind = node.kind().to_ascii_lowercase();
        match (parent.item().type_name(), kind.as_str()) {
            ("presentation", "slide") if self.pending.is_empty() => {
                let structural = !self.slides.is_empty();
                if structural
                    && matches!(
                        node.placement(),
                        Some(Placement::Before(_) | Placement::After(_))
                    )
                {
                    return Ok(None);
                }
                let (layout, from) = super::write::slide_request(node)?;
                let source = match from {
                    // A copy's source path names the evolving slides: a run
                    // with new slides maps a positional `/slide[N]` to its
                    // base slide, and publishes first for anything else.
                    Some(from) if structural => {
                        match evolving_slide(&from).and_then(|slide| self.order.get(slide)) {
                            Some(Some(base)) => SlideSource::Duplicate(*base),
                            _ => return Ok(None),
                        }
                    },
                    Some(from) => SlideSource::Duplicate(slide_position(&self.base, &from)?),
                    None => SlideSource::Layout(layout),
                };
                let destination = insertion(self.order.len(), node.placement(), |anchor| {
                    slide_at(&self.base, anchor)
                })?;
                self.admit(index)?;
                self.order.insert(destination, None);
                self.slides.push((source, destination, index));
                Ok(Some(
                    Path::parse(&format!("/slide[{}]", destination + 1))?.into_owned(),
                ))
            },
            ("slide", "shape" | "textbox" | "picture" | "image" | "img" | "pic")
                if self.accepts_sets()
                    && matches!(node.placement(), None | Some(Placement::End)) =>
            {
                let slide = parent.item().leaf()[0] as usize;
                let request = drawable_request(slide, node, &self.limits)?;
                let existing = parent
                    .child_nodes()
                    .filter(|child| child.item().type_name() == request.element)
                    .count();
                let staged = self
                    .pending
                    .adds_on(slide)
                    .filter(|add| add.element() == request.element)
                    .count();
                self.admit(index)?;
                self.pending.stage_add(request.add, index);
                // A new drawable joins the end of its slide's drawables.
                Ok(Some(
                    Path::parse(&format!(
                        "/slide[{}]/{}[{}]",
                        slide + 1,
                        request.element,
                        existing + staged + 1
                    ))?
                    .into_owned(),
                ))
            },
            _ => Ok(None),
        }
    }

    /// Publish the run: one slide-creation commit, or one slide-drawable
    /// commit plus one commit per changed text and slide visibility.
    ///
    /// # Errors
    ///
    /// The leaf's refusal with the item it belongs to, or the snapshot
    /// limit with the run's first item.
    pub(crate) fn commit(mut self) -> std::result::Result<(Iwork, Commit), (Error, usize)> {
        let first = self.first_index.unwrap_or(0);
        let Some(package) = self.base.keynote().cloned() else {
            return Err((
                refusal(Code::InternalError, "the Keynote run lost its package")
                    .with_format(FileFormat::Keynote),
                first,
            ));
        };
        let mut chain = Chain::new(package);
        if !self.slides.is_empty() {
            super::write::create_slides(&mut chain, &self.slides)
                .map_err(|(error, index)| (error, index.unwrap_or(first)))?;
        }
        self.pending
            .flush(&mut chain)
            .map_err(|(error, index)| (error, index.unwrap_or(first)))?;
        let commit = finish(
            &self.base,
            chain.changed,
            Owner::Keynote(chain.package),
            chain.patches,
            self.operations,
            &self.limits,
        )
        .map_err(|error| (error, first))?;
        Ok((self.base, commit))
    }
}
