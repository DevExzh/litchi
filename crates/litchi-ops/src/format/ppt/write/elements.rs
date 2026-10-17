//! Unified writes of the record-0323 element surface: slide comments (and,
//! as they land, table cells and hyperlinks).
//!
//! Every write resolves its targets against the read navigation, admits its
//! properties through the element's static registry, and stages one native
//! primary edit (`litchi_ppt::package::edit::Edit`), which owns record
//! placement, publication and read-back.

use litchi_core::{FileFormat, Position};
use litchi_ppt::comments::{CommentChange, DateTime, NewComment};

use super::super::Ppt;
use super::super::elements::Element;
use super::super::node::Node;
use super::super::spec::{COMMENT_WRITE_TABLE, CommentSetter, Kind};
use super::{WriteCommit, allocation, bound, finish_commit, master_from_emu, stage_structural};
use crate::error::{Code, Error, Result};
use crate::placement::Placement;
use crate::props::Props;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The assignments of one comment write.
#[derive(Default)]
struct CommentValues {
    text: Option<String>,
    author: Option<String>,
    initials: Option<String>,
    date: Option<DateTime>,
    x: Option<i32>,
    y: Option<i32>,
}

fn invalid(message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message).with_format(FileFormat::Ppt)
}

fn admit_comment(
    props: &Props<'_>,
    operation: Operation,
    limits: &SemanticLimits,
) -> Result<CommentValues> {
    let mut values = CommentValues::default();
    let mut assigned = 0_u32;
    let mut staged = 0_usize;
    for (key, value) in props.iter() {
        let (index, entry) = COMMENT_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ppt)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u32 << index;
        if assigned & bit != 0 {
            return Err(Error::new(
                Code::InvalidArgument,
                "a PPT comment property was assigned more than once",
            )
            .with_format(FileFormat::Ppt)
            .with_property(entry.spec.name));
        }
        assigned |= bit;
        entry
            .prepare(operation, value, |setter, scalar| {
                match (setter, scalar) {
                    (CommentSetter::X | CommentSetter::Y, Some(Scalar::Length(length))) => {
                        let master = master_from_emu(length.emu(), entry.spec.name)?;
                        let master = i32::try_from(master).map_err(|_error| {
                            invalid("the PPT comment anchor is outside the native range")
                                .with_property(entry.spec.name)
                        })?;
                        if matches!(setter, CommentSetter::X) {
                            values.x = Some(master);
                        } else {
                            values.y = Some(master);
                        }
                    },
                    (CommentSetter::Date, Some(Scalar::Text(text))) => {
                        values.date = Some(DateTime::parse(text).ok_or_else(|| {
                            invalid("a PPT comment date is YYYY-MM-DD[THH:MM[:SS[.fff]]]")
                                .with_property("date")
                                .with_suggestion("for example 2026-10-17T09:30:00")
                        })?);
                    },
                    (setter, Some(Scalar::Text(text))) => {
                        staged = staged.saturating_add(text.len());
                        bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
                        let mut owned = String::new();
                        owned
                            .try_reserve_exact(text.len())
                            .map_err(|source| allocation(source, "PPT comment text"))?;
                        owned.push_str(text);
                        match setter {
                            CommentSetter::Text => values.text = Some(owned),
                            CommentSetter::Author => values.author = Some(owned),
                            CommentSetter::Initials => values.initials = Some(owned),
                            CommentSetter::Date | CommentSetter::X | CommentSetter::Y => {
                                unreachable!("typed comment setters are matched above")
                            },
                        }
                    },
                    _ => unreachable!("PPT comment coercion returned an unexpected scalar"),
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ppt))?;
    }
    Ok(values)
}

/// The native target of one comment node.
fn comment_target(node: &Node<'_>) -> Result<(Position, Position)> {
    match node.element_item() {
        Some(Element::Comment { slide, index, .. }) => {
            Ok((Position::new(*slide), Position::new(*index)))
        },
        _ => Err(
            Error::new(Code::InternalError, "PPT comment target lost its comment")
                .with_format(FileFormat::Ppt),
        ),
    }
}

/// Set the text or anchor of every selected comment in one native edit.
pub(super) fn set_comments(
    file: &Ppt,
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let values = admit_comment(props, Operation::Set, limits)?;
    let change = CommentChange {
        text: values.text,
        x: values.x,
        y: values.y,
    };
    let mut targets = nodes
        .iter()
        .map(comment_target)
        .collect::<Result<Vec<_>>>()?;
    targets.sort_unstable();
    let native = stage_structural(file, targets.len(), limits, |edit| {
        for (slide, comment) in targets {
            edit.set_comment(slide, comment, &change)?;
        }
        Ok(())
    })?;
    file.check_source()?;
    finish_commit(file, native, limits)
}

/// Remove every selected comment in one native edit, last first so each
/// position still names the comment the caller selected.
pub(super) fn remove_comments(
    file: &Ppt,
    nodes: &[Node<'_>],
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let mut targets = nodes
        .iter()
        .map(comment_target)
        .collect::<Result<Vec<_>>>()?;
    targets.sort_unstable_by(|left, right| right.cmp(left));
    let native = stage_structural(file, targets.len(), limits, |edit| {
        for (slide, comment) in targets {
            edit.remove_comment(slide, comment)?;
        }
        Ok(())
    })?;
    file.check_source()?;
    finish_commit(file, native, limits)
}

/// Append one comment under a slide, as the PPTX `add comment` verb does:
/// `author` is required; `initials` default to the author's word initials,
/// `x`/`y` to the slide origin; `date` is never taken from the clock.
pub(super) fn add_comment(
    file: &Ppt,
    parent: &Node<'_>,
    node: &crate::NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, crate::path::Path<'static>)> {
    if parent.element() != Kind::Slide {
        return Err(
            Error::new(Code::UnsupportedType, "PPT comments require a slide parent")
                .with_format(FileFormat::Ppt),
        );
    }
    if !matches!(node.placement(), None | Some(Placement::End)) {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "PPT comment insertion supports end placement",
        )
        .with_format(FileFormat::Ppt)
        .with_suggestion("omit placement or use Placement::End"));
    }
    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    let values = admit_comment(node.props(), Operation::Add, limits)?;
    let author = values.author.ok_or_else(|| {
        Error::new(Code::MissingProperty, "a PPT comment requires an author")
            .with_format(FileFormat::Ppt)
            .with_property("author")
    })?;
    let initials = values.initials.unwrap_or_else(|| {
        author
            .split_whitespace()
            .filter_map(|word| word.chars().next())
            .collect()
    });
    let comment = NewComment {
        author,
        initials,
        text: values.text.unwrap_or_default(),
        date: values.date,
        x: values.x.unwrap_or(0),
        y: values.y.unwrap_or(0),
    };
    let slide = parent.slide_position().ok_or_else(|| {
        Error::new(Code::InternalError, "PPT slide parent has no position")
            .with_format(FileFormat::Ppt)
    })?;
    let mut position = Position::new(0);
    let native = stage_structural(file, 1, limits, |edit| {
        position = edit.add_comment(Position::new(slide), &comment)?;
        Ok(())
    })?;
    file.check_source()?;
    let commit = finish_commit(file, native, limits)?;
    Ok((
        commit,
        crate::path::Path::from_steps([
            crate::path::Step::new("slide")
                .with(crate::path::Predicate::Index(Position::new(slide))),
            crate::path::Step::new("comment").with(crate::path::Predicate::Index(position)),
        ]),
    ))
}
