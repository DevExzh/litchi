//! Legacy slide comments through the native eager edit (record 0252).
//!
//! `set(/slide[N]/comment[K], {text, x, y})` replaces one comment's text or
//! location; `add(/slide[N], comment, {text, author, initials, date, x, y})`
//! appends one, creating the named author when absent. Source owners are
//! promoted to an eager package for the edit, since comment parts are
//! separate parts the source-backed slide edit does not stage.
use super::node::Item;
use super::spec::{COMMENT_WRITE_TABLE, CommentSetter};
use super::structure::{bound, error};
use super::write::{Commit, Proof, validate_text};
use super::{Editable, Nav, Pptx};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Props, Result};
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};
use std::sync::Arc;

#[derive(Default)]
struct Values {
    text: Option<String>,
    author: Option<String>,
    initials: Option<String>,
    date: Option<String>,
    x: Option<i64>,
    y: Option<i64>,
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
                        validate_date(text)?;
                    }
                    let owned = Some(text.to_owned());
                    match setter {
                        CommentSetter::Text => values.text = owned,
                        CommentSetter::Author => values.author = owned,
                        CommentSetter::Initials => values.initials = owned,
                        CommentSetter::Date => values.date = owned,
                        CommentSetter::X | CommentSetter::Y => unreachable!("length setters"),
                    }
                },
                _ => unreachable!("comment property coercion"),
            }
            Ok(())
        })?;
    }
    Ok(values)
}

/// A comment `date` is an XML `dateTime`, as the comments part requires;
/// anything else is refused here rather than failing as a corrupt part when
/// the edit serializes it (record 0330). The accepted forms are the leaf
/// codec's.
fn validate_date(value: &str) -> Result<()> {
    if chrono::DateTime::parse_from_rfc3339(value).is_ok()
        || chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(error(
            Code::InvalidValue,
            "a comment date is an XML dateTime such as 2026-10-17T09:30:00Z",
        )
        .with_property("date"))
    }
}

impl Pptx {
    /// The eager base for an edit of separate parts: the eager owner itself,
    /// or the source owner's current bytes promoted to an eager package.
    fn eager_base(source: &Editable) -> Result<Arc<litchi_pptx::Package>> {
        Ok(match source {
            Editable::Eager(package) => Arc::clone(package),
            Editable::Source(package) => Arc::new(litchi_pptx::Package::from_reader(
                std::io::Cursor::new(package.to_bytes()?),
            )?),
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
        let source = self.editable_owner()?;
        let base = Self::eager_base(&source)?;
        let mut edit = base.edit()?;
        let changed = edit.set_comment_at(
            slide as usize,
            Position::new(position as usize),
            values.text.as_deref(),
            location,
        )?;
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if changed {
            self.adopt_eager(Arc::clone(&target))
        } else {
            self.clone()
        };
        Ok(Commit {
            after,
            source,
            target: Editable::Eager(target),
            proof: Proof::Eager(patch),
            operations: usize::from(changed),
        })
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
        let source = self.editable_owner()?;
        let base = Self::eager_base(&source)?;
        let mut edit = base.edit()?;
        let position = edit.add_comment_text(
            slide,
            &author,
            &initials,
            values.text.as_deref().unwrap_or(""),
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
            Commit {
                after,
                source,
                target: Editable::Eager(target),
                proof: Proof::Eager(patch),
                operations: 1,
            },
            path,
        ))
    }
}
