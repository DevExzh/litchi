//! Legacy slide comments through the native eager edit (record 0252).
//!
//! `set(/slide[N]/comment[K], {text, author, initials, date, x, y})`
//! replaces one comment's text, location or date, or re-attributes it
//! (record 0346: the author is renamed when the comment is their only one,
//! otherwise the comment moves to an author of that name and initials, so
//! other comments keep their attribution); `add(/slide[N], comment, {text,
//! author, initials, date, x, y, direction})` appends one, creating the
//! named author when absent (`direction=rtl` leads the text with a
//! right-to-left mark, as OfficeCLI does);
//! `remove(/slide[N]/comment[K])` deletes one (record 0333), as do
//! `remove(/slide[N]/moderncomment[K])` (with its replies) and
//! `remove(/slide[N]/moderncomment[K]/reply[R])` for threaded comments.
//! Removal keeps the comment authors (other slides may name them; record
//! 0319's rule: no edit of content the caller did not name) and drops a
//! slide's comment part with its last comment, as OfficeCLI does. Source
//! owners are promoted to an eager package for the edit, since comment
//! parts are separate parts the source-backed slide edit does not stage.
use super::node::Item;
use super::spec::{
    COMMENT_WRITE_TABLE, CommentSetter, MODERN_COMMENT_WRITE_TABLE, ModernSetter, REPLY_WRITE_TABLE,
};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options, validate_text};
use super::{Editable, Nav, Pptx};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Props, Result};
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};
use std::sync::Arc;

/// U+200F, prefixed to right-to-left comment text (record 0346).
const RIGHT_TO_LEFT_MARK: char = '\u{200F}';

#[derive(Default)]
struct Values {
    text: Option<String>,
    author: Option<String>,
    initials: Option<String>,
    date: Option<String>,
    x: Option<i64>,
    y: Option<i64>,
    /// `direction=rtl` on `add` (record 0346).
    rtl: bool,
}

fn admit(props: &Props<'_>, operation: Operation, limits: &SemanticLimits) -> Result<Values> {
    let mut values = Values::default();
    let mut assigned = 0_u32;
    let mut staged = 0_usize;
    for (key, value) in props.iter() {
        let (index, entry) = COMMENT_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|e| e.with_property(key.as_str().to_owned()))?;
        let bit = 1_u32 << index;
        if assigned & bit != 0 {
            return Err(error(
                Code::InvalidArgument,
                "a comment property was assigned more than once",
            )
            .with_property(entry.spec.name));
        }
        assigned |= bit;
        entry.prepare(operation, value, |setter, scalar| {
            match (setter, scalar) {
                (CommentSetter::Direction, Some(Scalar::Token(token))) => {
                    values.rtl = token == "rtl";
                },
                (CommentSetter::X | CommentSetter::Y, Some(Scalar::Length(length))) => {
                    if matches!(setter, CommentSetter::X) {
                        values.x = Some(length.emu());
                    } else {
                        values.y = Some(length.emu());
                    }
                },
                (setter, Some(Scalar::Text(text))) => {
                    staged = staged.saturating_add(text.len());
                    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
                    validate_text(text)?;
                    if matches!(setter, CommentSetter::Date) {
                        validate_date_time(text, "date")?;
                    }
                    let owned = Some(text.to_owned());
                    match setter {
                        CommentSetter::Text => values.text = owned,
                        CommentSetter::Author => values.author = owned,
                        CommentSetter::Initials => values.initials = owned,
                        CommentSetter::Date => values.date = owned,
                        CommentSetter::X | CommentSetter::Y | CommentSetter::Direction => {
                            unreachable!("length and token setters")
                        },
                    }
                },
                _ => unreachable!("comment property coercion"),
            }
            Ok(())
        })?;
    }
    Ok(values)
}

/// A comment `date` (or a modern comment's `created`) is an XML
/// `dateTime`, as the comments part requires; anything else is refused here
/// rather than failing as a corrupt part when the edit serializes it
/// (record 0330). The accepted forms are the leaf codec's.
fn validate_date_time(value: &str, property: &'static str) -> Result<()> {
    if chrono::DateTime::parse_from_rfc3339(value).is_ok()
        || chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(error(
            Code::InvalidValue,
            "a comment date is an XML dateTime such as 2026-10-17T09:30:00Z",
        )
        .with_property(property))
    }
}

impl Pptx {
    /// The eager base for an edit of separate parts: the eager owner itself,
    /// or the source owner's current bytes promoted to an eager package.
    pub(super) fn eager_base(&self, source: &Editable) -> Result<Arc<litchi_pptx::Package>> {
        Ok(match source {
            Editable::Eager(package) => Arc::clone(package),
            Editable::Source(package) => {
                super::eager_package_from_source(package.to_bytes()?, self.0.ooxml_limits)?
            },
        })
    }

    pub(super) fn set_comment(
        &self,
        slide: u32,
        position: u32,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let values = admit(props, Operation::Set, limits)?;
        let location = match (values.x, values.y) {
            (None, None) => None,
            (x, y) => {
                let current = self
                    .slide_comments(slide as usize)?
                    .and_then(|list| list.entries().get(position as usize).map(|c| c.location()))
                    .ok_or_else(|| error(Code::PathNotFound, "the comment is missing"))?;
                Some((x.unwrap_or(current.0), y.unwrap_or(current.1)))
            },
        };
        let source = self.editable_owner_for_write()?;
        let base = self.eager_base(&source)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let mut changed = edit.set_comment_at(
            slide as usize,
            Position::new(position as usize),
            values.text.as_deref(),
            location,
        )?;
        // Each setter addresses the comment by its slide position, which
        // re-attribution keeps (record 0346).
        if let Some(date) = values.date.as_deref() {
            changed |=
                edit.set_comment_date_at(slide as usize, Position::new(position as usize), date)?;
        }
        if values.author.is_some() || values.initials.is_some() {
            changed |= edit.set_comment_author_at(
                slide as usize,
                Position::new(position as usize),
                values.author.as_deref(),
                values.initials.as_deref(),
            )?;
        }
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if changed {
            self.adopt_eager(Arc::clone(&target))
        } else {
            self.clone()
        };
        Ok(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        ))
    }

    pub(super) fn add_comment(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let parent = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Slide(slide) = parent.item() else {
            return Err(error(
                Code::UnsupportedType,
                "comments require a slide parent",
            ));
        };
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "comment insertion supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let values = admit(node.props(), Operation::Add, limits)?;
        let author = values.author.ok_or_else(|| {
            error(Code::MissingProperty, "comments require an author").with_property("author")
        })?;
        let initials = values.initials.unwrap_or_else(|| {
            author
                .split_whitespace()
                .filter_map(|word| word.chars().next())
                .collect()
        });
        let slide = slide.position as usize;
        // OfficeCLI's convention for right-to-left comment text: `p:cm` has
        // no direction attribute, so a right-to-left mark leads the text
        // (once; an empty text stays empty).
        let mut text = values.text.unwrap_or_default();
        if values.rtl && !text.is_empty() && !text.starts_with(RIGHT_TO_LEFT_MARK) {
            text.insert(0, RIGHT_TO_LEFT_MARK);
        }
        let source = self.editable_owner()?;
        let base = self.eager_base(&source)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let position = edit.add_comment_text(
            slide,
            &author,
            &initials,
            &text,
            (values.x.unwrap_or(0), values.y.unwrap_or(0)),
            values.date.as_deref(),
        )?;
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide))),
            Step::new("comment").with(Predicate::Index(position)),
        ]);
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }
}

impl Pptx {
    /// A `remove` of a legacy comment, modern comment or reply path; `None`
    /// for other targets (record 0333).
    pub(super) fn try_remove_comment(
        &self,
        target: &crate::target::TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            [
                "comment",
                "moderncomment",
                "modern-comment",
                "thread",
                "threadedcomment",
                "reply",
            ]
            .iter()
            .any(|name| step.name().eq_ignore_ascii_case(name))
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        enum Target {
            Legacy(usize, usize),
            Modern(usize, usize),
            Reply(usize, usize, usize),
        }
        let target = match node.item() {
            Item::Comment(comment) => {
                Target::Legacy(comment.slide as usize, comment.index as usize)
            },
            Item::ModernComment(comment) => {
                Target::Modern(comment.slide as usize, comment.index as usize)
            },
            Item::Reply(reply) => Target::Reply(
                reply.slide as usize,
                reply.comment as usize,
                reply.index as usize,
            ),
            _ => return Ok(None),
        };
        drop(node);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let source = self.editable_owner()?;
        let base = self.eager_base(&source)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let changed = match target {
            Target::Legacy(slide, index) => edit.remove_comment_at(slide, Position::new(index))?,
            Target::Modern(slide, index) => {
                edit.remove_modern_comment_at(slide, Position::new(index))?
            },
            Target::Reply(slide, comment, reply) => edit.remove_modern_comment_reply_at(
                slide,
                Position::new(comment),
                Position::new(reply),
            )?,
        };
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if changed {
            self.adopt_eager(Arc::clone(&target))
        } else {
            self.clone()
        };
        Ok(Some(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        )))
    }
}

/// The admitted assignments of a modern comment or reply (record 0346).
#[derive(Default)]
struct ModernValues {
    text: Option<String>,
    author: Option<String>,
    initials: Option<String>,
    created: Option<String>,
    resolved: Option<bool>,
    parent: Option<String>,
}

fn admit_modern<const N: usize>(
    table: &crate::registry::Table<ModernSetter, N, 0>,
    props: &Props<'_>,
    operation: Operation,
    limits: &SemanticLimits,
) -> Result<ModernValues> {
    let mut values = ModernValues::default();
    let mut assigned = 0_u32;
    let mut staged = 0_usize;
    for (key, value) in props.iter() {
        let (index, entry) = table
            .resolve(key.as_str())
            .map_err(|e| e.with_property(key.as_str().to_owned()))?;
        let bit = 1_u32 << index;
        if assigned & bit != 0 {
            return Err(error(
                Code::InvalidArgument,
                "a modern comment property was assigned more than once",
            )
            .with_property(entry.spec.name));
        }
        assigned |= bit;
        entry.prepare(operation, value, |setter, scalar| {
            match (setter, scalar) {
                (ModernSetter::Resolved, Some(Scalar::Bool(resolved))) => {
                    values.resolved = Some(resolved);
                },
                (setter, Some(Scalar::Text(text))) => {
                    staged = staged.saturating_add(text.len());
                    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
                    validate_text(text)?;
                    if matches!(setter, ModernSetter::Created) {
                        validate_date_time(text, "created")?;
                    }
                    let owned = Some(text.to_owned());
                    match setter {
                        ModernSetter::Text => values.text = owned,
                        ModernSetter::Author => values.author = owned,
                        ModernSetter::Initials => values.initials = owned,
                        ModernSetter::Created => values.created = owned,
                        ModernSetter::Parent => values.parent = owned,
                        ModernSetter::Resolved => unreachable!("a boolean setter"),
                    }
                },
                _ => unreachable!("modern comment property coercion"),
            }
            Ok(())
        })?;
    }
    Ok(values)
}

/// OfficeCLI's default initials: the first character of each word.
fn derived_initials(author: &str) -> String {
    author
        .split_whitespace()
        .filter_map(|word| word.chars().next())
        .collect()
}

impl Pptx {
    /// A `set` of a modern comment or reply path; `None` for other targets
    /// (record 0346).
    pub(super) fn try_set_modern_comment(
        &self,
        target: &crate::target::TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            MODERN_STEPS
                .iter()
                .chain(&["reply"])
                .any(|name| step.name().eq_ignore_ascii_case(name))
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let (slide, comment, reply, current) = match node.item() {
            Item::ModernComment(item) => (
                item.slide,
                item.index,
                None,
                item.comment().and_then(|comment| comment.text()),
            ),
            Item::Reply(item) => (
                item.slide,
                item.comment,
                Some(item.index),
                item.reply().and_then(|reply| reply.text()),
            ),
            _ => return Ok(None),
        };
        let mut values = if reply.is_some() {
            admit_modern(&REPLY_WRITE_TABLE, props, Operation::Set, limits)?
        } else {
            admit_modern(&MODERN_COMMENT_WRITE_TABLE, props, Operation::Set, limits)?
        };
        // An equal text keeps the body (and its formatting) byte for byte.
        if values.text.is_some() && values.text.as_deref() == current {
            values.text = None;
        }
        drop(node);
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let source = self.editable_owner_for_write()?;
        let base = self.eager_base(&source)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let change = litchi_pptx::opened::ModernCommentChange {
            text: values.text.as_deref(),
            author: values.author.as_deref(),
            initials: values.initials.as_deref(),
            created: values.created.as_deref(),
            resolved: values.resolved,
        };
        let changed = edit.set_modern_comment_at(
            slide as usize,
            Position::new(comment as usize),
            reply.map(|reply| Position::new(reply as usize)),
            &change,
        )?;
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if changed {
            self.adopt_eager(Arc::clone(&target))
        } else {
            self.clone()
        };
        Ok(Some(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        )))
    }

    /// `add(/slide[N], modernComment, {text, author, initials, created,
    /// resolved, parent})`: a new thread, or with `parent` a reply to that
    /// thread (record 0346).
    pub(super) fn add_modern_comment(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let nav = Nav::new(self);
        let parent = crate::path::resolve::get(&nav, &parent)?;
        let Item::Slide(slide) = parent.item() else {
            return Err(error(
                Code::UnsupportedType,
                "modern comments require a slide parent",
            ));
        };
        let slide = slide.position;
        drop(parent);
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "modern comment insertion supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let values = admit_modern(
            &MODERN_COMMENT_WRITE_TABLE,
            node.props(),
            Operation::Add,
            limits,
        )?;
        let author = values.author.ok_or_else(|| {
            error(Code::MissingProperty, "modern comments require an author")
                .with_property("author")
        })?;
        // The editor reads no clock (ADR 0034): the stamp is the caller's.
        let created = values.created.ok_or_else(|| {
            error(
                Code::MissingProperty,
                "modern comments require a creation stamp",
            )
            .with_property("created")
            .with_suggestion("pass created=<xsd:dateTime>, for example 2026-10-19T09:30:00Z")
        })?;
        let initials = values.initials.unwrap_or_else(|| derived_initials(&author));
        let thread = match values.parent.as_deref() {
            None => None,
            Some(path) => {
                let path = PathArg::parse(path).map_err(|error| error.with_property("parent"))?;
                let node = crate::path::resolve::get(&nav, &path)
                    .map_err(|error| error.with_property("parent"))?;
                let Item::ModernComment(thread) = node.item() else {
                    return Err(error(
                        Code::InvalidValue,
                        "a reply's parent is a top-level modern comment",
                    )
                    .with_property("parent"));
                };
                if thread.slide != slide {
                    return Err(error(
                        Code::InvalidValue,
                        "a reply's parent is a modern comment of the same slide",
                    )
                    .with_property("parent"));
                }
                if values.resolved.is_some() {
                    return Err(error(
                        Code::InvalidArgument,
                        "resolution is a thread-level state; a reply has none",
                    )
                    .with_property("resolved"));
                }
                Some(thread.index)
            },
        };
        let text = values.text.unwrap_or_default();
        let source = self.editable_owner()?;
        let base = self.eager_base(&source)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let slide_step = Step::new("slide").with(Predicate::Index(Position::new(slide as usize)));
        let path = match thread {
            None => {
                let position = edit.add_modern_comment_text(
                    slide as usize,
                    &author,
                    Some(&initials),
                    &text,
                    &created,
                    values.resolved == Some(true),
                )?;
                Path::from_steps([
                    slide_step,
                    Step::new("moderncomment").with(Predicate::Index(position)),
                ])
            },
            Some(thread) => {
                let position = edit.add_modern_comment_reply_text(
                    slide as usize,
                    Position::new(thread as usize),
                    &author,
                    Some(&initials),
                    &text,
                    &created,
                )?;
                Path::from_steps([
                    slide_step,
                    Step::new("moderncomment")
                        .with(Predicate::Index(Position::new(thread as usize))),
                    Step::new("reply").with(Predicate::Index(position)),
                ])
            },
        };
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }
}

/// The step names of a modern comment path.
const MODERN_STEPS: &[&str] = &[
    "moderncomment",
    "modern-comment",
    "thread",
    "threadedcomment",
];
