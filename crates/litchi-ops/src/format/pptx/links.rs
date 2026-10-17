//! Shape and run hyperlinks as elements (record 0321).
//!
//! `/slide[N]/shape[M]/hyperlink` is a shape's click action (its
//! `p:cNvPr`'s, else its first run's, as OfficeCLI's `ReadShapeHyperlink`);
//! `/slide[N]/shape[M]/paragraph[K]/run[L]/hyperlink` is one run's. Both
//! read `link` (an external URL, `slide[N]` for a jump to another slide, or
//! a named action: `firstslide`, `lastslide`, `nextslide`,
//! `previousslide`, `endshow`) and `tooltip`.
//!
//! - `add(shape, hyperlink, {link, tooltip})` attaches a shape-wide link:
//!   on the shape's `p:cNvPr` and every text run (OfficeCLI's
//!   `ApplyShapeHyperlink`); under a run it attaches the run's.
//! - `set` replaces `link` and/or `tooltip` (`link=none` removes);
//! - `remove` deletes the click actions.
//!
//! Source owners are promoted to an eager package, as for theme writes.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::Path;
use litchi_pptx::hyperlinks::{Click, Link, Target};

use super::node::{Item, TextItem, TreeItem};
use super::spec::{HYPERLINK_WRITE_TABLE, HyperlinkKey, Kind};
use super::structure::{bound, error};
use super::write::{Commit, Proof};
use super::{Editable, Nav, Pptx, PptxOwner};
use crate::node::Value;
use crate::path::PathArg;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Props, Result};

const SHOW_JUMP: &str = "ppaction://hlinkshowjump?jump=";
const NAMED: [&str; 5] = [
    "firstslide",
    "lastslide",
    "nextslide",
    "previousslide",
    "endshow",
];

/// What a hyperlink belongs to.
#[derive(Clone)]
pub(super) enum Owner {
    Shape(TreeItem),
    Run(TextItem),
}

/// A resolved hyperlink and its owner.
#[derive(Clone)]
pub(super) struct LinkItem {
    pub(super) owner: Owner,
    pub(super) link: Arc<Link>,
}

/// OfficeCLI's `link` readback of a target.
pub(super) fn link_text(link: &Link) -> Cow<'_, str> {
    match &link.target {
        Target::Url(url) => Cow::Borrowed(url.as_str()),
        Target::Slide(position) => Cow::Owned(format!("slide[{}]", position.get() + 1)),
        Target::Action(action) => match action.strip_prefix(SHOW_JUMP) {
            Some(jump) if NAMED.iter().any(|name| jump.eq_ignore_ascii_case(name)) => {
                Cow::Owned(jump.to_ascii_lowercase())
            },
            _ => Cow::Borrowed(action.as_str()),
        },
    }
}

/// The value of the hyperlink property at `index`.
pub(super) fn value(item: &LinkItem, index: usize) -> Option<Value<'_>> {
    match HyperlinkKey::at(index)? {
        HyperlinkKey::Link => Some(Value::Text(link_text(&item.link))),
        HyperlinkKey::Tooltip => item
            .link
            .tooltip
            .as_deref()
            .map(|tooltip| Value::Text(Cow::Borrowed(tooltip))),
    }
}

impl Pptx {
    /// Resolve a click action of the slide at `position`.
    fn slide_hyperlink(&self, position: usize, click: &Click) -> Result<Option<Link>> {
        self.finish(|| {
            if self.0.revision.is_some() {
                let Some(slide) = self.revised_slide_for_links(position)? else {
                    return Ok(None);
                };
                return Ok(slide.hyperlink(click)?);
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(None);
                    };
                    Ok(slide.hyperlink(click)?)
                },
                PptxOwner::Eager(package) => {
                    let Some(slide) = super::eager_slide(self, package, position)? else {
                        return Ok(None);
                    };
                    Ok(slide.hyperlink(click)?)
                },
            }
        })
    }

    fn revised_slide_for_links(
        &self,
        position: usize,
    ) -> Result<Option<litchi_pptx::source::Slide>> {
        let Some(revision) = &self.0.revision else {
            return Ok(None);
        };
        let presentation = revision.presentation()?;
        if position >= presentation.slide_count()? {
            return Ok(None);
        }
        Ok(Some(presentation.slide(position)?))
    }
}

/// A shape's hyperlink: its `p:cNvPr` click, else its first run's.
pub(super) fn of_shape(file: &Pptx, tree: &TreeItem) -> Result<Option<LinkItem>> {
    if tree.kind != Kind::Shape {
        return Ok(None);
    }
    let Some(slide) = tree.container.slide() else {
        return Ok(None);
    };
    let click = match Click::of_shape(tree.handle.view())? {
        Some(click) => Some(click),
        None => tree.body()?.and_then(|body| {
            body.paragraphs
                .iter()
                .flat_map(|paragraph| paragraph.runs.iter())
                .find(|run| run.kind == litchi_drawingml::text::body::RunKind::Text)
                .and_then(|run| run.properties.hyperlink.as_ref())
                .map(Click::from_text)
        }),
    };
    let Some(click) = click else {
        return Ok(None);
    };
    Ok(file
        .slide_hyperlink(slide as usize, &click)?
        .map(|link| LinkItem {
            owner: Owner::Shape(tree.clone()),
            link: Arc::new(link),
        }))
}

/// A run's hyperlink.
pub(super) fn of_run(file: &Pptx, text: &TextItem) -> Result<Option<LinkItem>> {
    let Some(slide) = text.shape.container.slide() else {
        return Ok(None);
    };
    let Some(link) = text.run().and_then(|run| run.properties.hyperlink.as_ref()) else {
        return Ok(None);
    };
    Ok(file
        .slide_hyperlink(slide as usize, &Click::from_text(link))?
        .map(|link| LinkItem {
            owner: Owner::Run(text.clone()),
            link: Arc::new(link),
        }))
}

/// The hyperlink child of a shape or run, as a list.
pub(super) fn children(file: &Pptx, parent: &Item) -> Result<Vec<Item>> {
    let found = match parent {
        Item::Tree(tree) => of_shape(file, tree)?,
        Item::Run(text) => of_run(file, text)?,
        _ => None,
    };
    Ok(found.map(Item::Hyperlink).into_iter().collect())
}

/// Parse OfficeCLI's `link` input; `None` removes.
fn parse_link(value: &str, slides: usize) -> Result<Option<Target>> {
    let value = value.trim();
    if value.is_empty() || value.eq_ignore_ascii_case("none") {
        return Ok(None);
    }
    if let Some(name) = NAMED.iter().find(|name| value.eq_ignore_ascii_case(name)) {
        return Ok(Some(Target::Action(format!("{SHOW_JUMP}{name}"))));
    }
    let lower = value.to_ascii_lowercase();
    if let Some(number) = lower
        .strip_prefix("slide[")
        .and_then(|rest| rest.strip_suffix(']'))
    {
        let number: usize = number.parse().map_err(|_| {
            error(Code::InvalidValue, "a slide jump is slide[N]").with_property("link")
        })?;
        if number == 0 || number > slides {
            return Err(error(
                Code::OutOfRange,
                "the slide-jump target is past the last slide",
            )
            .with_property("link")
            .with_suggestion("add the slide first, then set the link"));
        }
        return Ok(Some(Target::Slide(Position::new(number - 1))));
    }
    // OfficeCLI's scheme allow-list: refuse scripting and data schemes.
    let scheme = lower.split_once(':').map(|(scheme, _)| scheme);
    if matches!(
        scheme,
        Some("javascript" | "vbscript" | "data" | "file" | "jar")
    ) {
        return Err(
            error(Code::InvalidValue, "the hyperlink scheme is not allowed")
                .with_property("link")
                .with_suggestion(
                    "use an http(s), mailto or relative target, slide[N], or a named action",
                ),
        );
    }
    if value.starts_with("ppaction://") {
        return Ok(Some(Target::Action(value.to_owned())));
    }
    Ok(Some(Target::Url(value.to_owned())))
}

/// The admitted `link` and `tooltip` of an add or set.
struct Assignment {
    link: Option<Option<Target>>,
    tooltip: Option<Option<String>>,
}

fn admit(
    props: &Props<'_>,
    operation: Operation,
    slides: usize,
    limits: &SemanticLimits,
) -> Result<Assignment> {
    let mut out = Assignment {
        link: None,
        tooltip: None,
    };
    let mut staged = 0_usize;
    for (key, value) in props.iter() {
        let (_, entry) = HYPERLINK_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|e| e.with_property(key.as_str().to_owned()))?;
        entry.prepare(operation, value, |key, value| {
            let text = match value {
                Some(Scalar::Text(text) | Scalar::Token(text)) => text,
                _ => "",
            };
            staged = staged.saturating_add(text.len());
            bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
            super::write::validate_text(text)?;
            let duplicate = match key {
                HyperlinkKey::Link => out.link.replace(parse_link(text, slides)?).is_some(),
                HyperlinkKey::Tooltip => out
                    .tooltip
                    .replace((!text.is_empty()).then(|| text.to_owned()))
                    .is_some(),
            };
            if duplicate {
                return Err(error(
                    Code::InvalidArgument,
                    "a hyperlink property was assigned more than once",
                ));
            }
            Ok(())
        })?;
    }
    Ok(out)
}

impl Pptx {
    /// Write `link` (`None` removes) on `owner` in one native edit.
    fn write_link(&self, owner: &Owner, link: Option<&Link>) -> Result<Commit> {
        let tree = match owner {
            Owner::Shape(tree) => tree,
            Owner::Run(text) => &text.shape,
        };
        let Some(slide) = tree.container.slide() else {
            return Err(error(
                Code::UnsupportedCapability,
                "master and layout hyperlinks are not editable",
            ));
        };
        if tree.kind != Kind::Shape {
            return Err(error(
                Code::UnsupportedType,
                "hyperlinks are written on shapes and their runs",
            ));
        }
        let shape = tree.handle.position().get();
        let source = self.editable_owner()?;
        let base = super::smartart::eager(&source)?;
        let mut native = base.edit()?;
        match owner {
            Owner::Shape(_) => native.set_shape_hyperlink(slide as usize, shape, link)?,
            Owner::Run(text) => native.set_run_hyperlink(
                slide as usize,
                shape,
                Position::new(text.paragraph as usize),
                Position::new(text.run as usize),
                link,
            )?,
        };
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Commit {
            after,
            source,
            target: Editable::Eager(target),
            proof: Proof::Eager(patch),
            operations: 1,
        })
    }

    /// `add(shape or run, hyperlink, {link, tooltip})`.
    pub(super) fn add_hyperlink(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let resolved = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let owner = match resolved.item() {
            Item::Tree(tree) if tree.kind == Kind::Shape => Owner::Shape(tree.clone()),
            Item::Run(text) => Owner::Run(text.clone()),
            _ => {
                return Err(error(
                    Code::UnsupportedType,
                    "a hyperlink requires a shape or run parent",
                ));
            },
        };
        let path = crate::node::Project::path(&resolved).to_string();
        drop(resolved);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let assignment = admit(
            node.props(),
            Operation::Add,
            self.slides_in_catalog()?,
            limits,
        )?;
        let Some(Some(target)) = assignment.link else {
            return Err(error(Code::MissingProperty, "a hyperlink needs its target")
                .with_property("link")
                .with_suggestion("link=https://example.com"));
        };
        let link = Link {
            target,
            tooltip: assignment.tooltip.flatten(),
        };
        let commit = self.write_link(&owner, Some(&link))?;
        let text = format!("{path}/hyperlink");
        let path = Path::parse(&text)
            .map(Path::into_owned)
            .map_err(|_| error(Code::InternalError, "the hyperlink path does not parse"))?;
        Ok((commit, path))
    }

    /// A `set` on a hyperlink path; `None` for other targets.
    pub(super) fn try_set_hyperlink(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self.hyperlink_target(target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let assignment = admit(props, Operation::Set, self.slides_in_catalog()?, limits)?;
        let link = match assignment.link {
            Some(None) => None,
            Some(Some(target)) => Some(Link {
                target,
                tooltip: assignment
                    .tooltip
                    .unwrap_or_else(|| item.link.tooltip.clone()),
            }),
            None => Some(Link {
                target: item.link.target.clone(),
                tooltip: assignment.tooltip.flatten(),
            }),
        };
        self.write_link(&item.owner, link.as_ref()).map(Some)
    }

    /// A `remove` of a hyperlink path; `None` for other targets.
    pub(super) fn try_remove_hyperlink(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self.hyperlink_target(target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        self.write_link(&item.owner, None).map(Some)
    }

    fn hyperlink_target(&self, target: &TargetArg<'_>) -> Result<Option<LinkItem>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            step.name().eq_ignore_ascii_case("hyperlink")
                || step.name().eq_ignore_ascii_case("hlink")
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        Ok(match node.item() {
            Item::Hyperlink(item) => Some(item.clone()),
            _ => None,
        })
    }
}
