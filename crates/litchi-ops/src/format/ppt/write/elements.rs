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
use super::super::spec::{CELL_WRITE_TABLE, COMMENT_WRITE_TABLE, CommentSetter, Kind};
use super::{
    WriteCommit, admit_text_prop, allocation, bound, finish_commit, master_from_emu,
    stage_structural, unsupported_target,
};
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

/// The refusal of a table row, column or cell insertion or removal.
///
/// A legacy table is a group of cell shapes and border lines laid out by
/// position; inserting or removing a band would re-lay out every cell and
/// line after it and the group's anchor, which the leaf does not model.
pub(super) fn table_layout_refusal(verb: &'static str) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        match verb {
            "add" => "PPT table rows and columns cannot be inserted: the leaf does not re-lay out a table's cell shapes and border lines",
            _ => "PPT table rows, columns and cells cannot be removed: the leaf does not re-lay out a table's cell shapes and border lines",
        },
    )
    .with_format(FileFormat::Ppt)
    .with_suggestion("set cell text instead; table layout edits remain a follow-up (record 0323)")
}

/// Set the text of every selected table cell in one native edit.
pub(super) fn set_cells(
    file: &Ppt,
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    if props.is_empty() {
        return finish_commit(file, stage_structural(file, 0, limits, |_| Ok(()))?, limits);
    }
    let text = admit_text_prop(props, &CELL_WRITE_TABLE, limits)?;
    let mut targets = Vec::with_capacity(nodes.len());
    for node in nodes {
        let (Some(target), Some(Element::Cell { row, column, .. })) =
            (node.native_shape_target(), node.element_item())
        else {
            return Err(unsupported_target(
                "nested PPT group table writes are not integrated",
            ));
        };
        targets.push((target, Position::new(*row), Position::new(*column)));
    }
    let native = stage_structural(file, targets.len(), limits, |edit| {
        for (target, row, column) in targets {
            edit.set_table_cell_text(target, row, column, &text)?;
        }
        Ok(())
    })?;
    file.check_source()?;
    finish_commit(file, native, limits)
}

const NAMED_JUMPS: [(&str, litchi_ppt::hyperlink::InteractionJump); 5] = [
    (
        "firstslide",
        litchi_ppt::hyperlink::InteractionJump::FirstSlide,
    ),
    (
        "lastslide",
        litchi_ppt::hyperlink::InteractionJump::LastSlide,
    ),
    (
        "nextslide",
        litchi_ppt::hyperlink::InteractionJump::NextSlide,
    ),
    (
        "previousslide",
        litchi_ppt::hyperlink::InteractionJump::PreviousSlide,
    ),
    ("endshow", litchi_ppt::hyperlink::InteractionJump::EndShow),
];

/// Parse the PPTX `link` vocabulary; `None` removes.
fn parse_link(value: &str, slides: usize) -> Result<Option<litchi_ppt::hyperlink::Target>> {
    use litchi_ppt::hyperlink::Target;

    let value = value.trim();
    if value.is_empty() || value.eq_ignore_ascii_case("none") {
        return Ok(None);
    }
    let lower = value.to_ascii_lowercase();
    let named = lower
        .strip_prefix("ppaction://hlinkshowjump?jump=")
        .unwrap_or(&lower);
    if let Some((_, jump)) = NAMED_JUMPS.iter().find(|(name, _)| *name == named) {
        return Ok(Some(Target::Jump(*jump)));
    }
    if let Some(number) = lower
        .strip_prefix("slide[")
        .and_then(|rest| rest.strip_suffix(']'))
    {
        let number: usize = number
            .parse()
            .map_err(|_error| invalid("a slide jump is slide[N]").with_property("link"))?;
        if number == 0 || number > slides {
            return Err(Error::new(
                Code::OutOfRange,
                "the slide-jump target is past the last slide",
            )
            .with_format(FileFormat::Ppt)
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
        return Err(invalid("the hyperlink scheme is not allowed")
            .with_property("link")
            .with_suggestion(
                "use an http(s), mailto or relative target, slide[N], or a named action",
            ));
    }
    if lower.starts_with("ppaction://") {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "PPT hyperlinks support URL, slide and named slide-show jump targets",
        )
        .with_format(FileFormat::Ppt)
        .with_property("link"));
    }
    if value.contains('\0') {
        return Err(invalid("a hyperlink holds no NUL character").with_property("link"));
    }
    Ok(Some(Target::Url(value.to_owned())))
}

/// The admitted `link` and `tooltip` of a hyperlink add or set.
struct LinkAssignment {
    link: Option<Option<litchi_ppt::hyperlink::Target>>,
    tooltip: Option<Option<String>>,
}

fn admit_link(
    props: &Props<'_>,
    operation: Operation,
    slides: usize,
    limits: &SemanticLimits,
) -> Result<LinkAssignment> {
    use super::super::spec::{HYPERLINK_WRITE_TABLE, HyperlinkSetter};

    let mut out = LinkAssignment {
        link: None,
        tooltip: None,
    };
    let mut staged = 0_usize;
    for (key, value) in props.iter() {
        let (_, entry) = HYPERLINK_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| {
                error
                    .with_format(FileFormat::Ppt)
                    .with_property(key.as_str().to_owned())
            })?;
        entry
            .prepare(operation, value, |setter, scalar| {
                let text = match scalar {
                    Some(Scalar::Text(text)) => text,
                    _ => "",
                };
                staged = staged.saturating_add(text.len());
                bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
                let duplicate = match setter {
                    HyperlinkSetter::Link => out.link.replace(parse_link(text, slides)?).is_some(),
                    HyperlinkSetter::Tooltip => {
                        if text.contains('\0') {
                            return Err(invalid("a tooltip holds no NUL character"));
                        }
                        out.tooltip
                            .replace((!text.is_empty()).then(|| text.to_owned()))
                            .is_some()
                    },
                };
                if duplicate {
                    return Err(Error::new(
                        Code::InvalidArgument,
                        "a hyperlink property was assigned more than once",
                    ));
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ppt))?;
    }
    Ok(out)
}

/// What one hyperlink write targets: a shape's click, or a click over a
/// UTF-16 range of its text.
#[derive(Clone, Copy)]
struct LinkOwner {
    target: litchi_ppt::text_edit::Target,
    range: Option<(usize, usize)>,
}

/// The native owner of one hyperlink node, or of a shape or run parent.
fn link_owner(node: &Node<'_>) -> Result<LinkOwner> {
    let target = node.native_shape_target().ok_or_else(|| {
        unsupported_target("nested PPT group hyperlink writes are not integrated")
    })?;
    let range = match node.element_item() {
        Some(Element::Hyperlink { run, .. }) => *run,
        Some(Element::Run { piece, .. }) => Some((piece.start, piece.end)),
        _ => None,
    };
    Ok(LinkOwner { target, range })
}

/// Write `click` (`None` removes) on every owner in one native edit.
fn write_links(
    file: &Ppt,
    owners: Vec<(LinkOwner, Option<litchi_ppt::hyperlink::Click>)>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let native = stage_structural(file, owners.len(), limits, |edit| {
        for (owner, click) in &owners {
            match owner.range {
                None => edit.set_shape_hyperlink(owner.target, click.as_ref())?,
                Some((start, end)) => {
                    edit.set_text_hyperlink(owner.target, start, end - start, click.as_ref())?;
                },
            }
        }
        Ok(())
    })?;
    file.check_source()?;
    finish_commit(file, native, limits)
}

/// Set `link` and/or `tooltip` of every selected hyperlink: `link=none`
/// removes, a tooltip-only set keeps the target.
pub(super) fn set_hyperlinks(
    file: &Ppt,
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let assignment = admit_link(props, Operation::Set, file.slide_count()?, limits)?;
    let mut owners = Vec::with_capacity(nodes.len());
    for node in nodes {
        let Some(Element::Hyperlink { click, .. }) = node.element_item() else {
            return Err(unsupported_target("hyperlink target lost its hyperlink"));
        };
        let click = match &assignment.link {
            Some(None) => None,
            Some(Some(target)) => Some(litchi_ppt::hyperlink::Click {
                target: target.clone(),
                tooltip: assignment
                    .tooltip
                    .clone()
                    .unwrap_or_else(|| click.tooltip.clone()),
            }),
            None => Some(litchi_ppt::hyperlink::Click {
                target: click.target.clone(),
                tooltip: assignment.tooltip.clone().flatten(),
            }),
        };
        owners.push((link_owner(node)?, click));
    }
    write_links(file, owners, limits)
}

/// Remove every selected hyperlink.
pub(super) fn remove_hyperlinks(
    file: &Ppt,
    nodes: &[Node<'_>],
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let owners = nodes
        .iter()
        .map(|node| link_owner(node).map(|target| (target, None)))
        .collect::<Result<Vec<_>>>()?;
    write_links(file, owners, limits)
}

/// `add(shape, hyperlink, {link, tooltip})`: attach a shape's click link.
pub(super) fn add_hyperlink(
    file: &Ppt,
    parent: &Node<'_>,
    node: &crate::NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, crate::path::Path<'static>)> {
    let shape_parent = parent.element() == Kind::Shape && parent.element_item().is_none();
    if !shape_parent && parent.element() != Kind::Run {
        return Err(Error::new(
            Code::UnsupportedType,
            "a PPT hyperlink requires a shape or run parent",
        )
        .with_format(FileFormat::Ppt));
    }
    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    let assignment = admit_link(node.props(), Operation::Add, file.slide_count()?, limits)?;
    let Some(Some(target)) = assignment.link else {
        return Err(
            Error::new(Code::MissingProperty, "a PPT hyperlink needs its target")
                .with_format(FileFormat::Ppt)
                .with_property("link")
                .with_suggestion("link=https://example.com"),
        );
    };
    let click = litchi_ppt::hyperlink::Click {
        target,
        tooltip: assignment.tooltip.flatten(),
    };
    let commit = write_links(file, vec![(link_owner(parent)?, Some(click))], limits)?;
    let text = format!("{}/hyperlink", crate::node::Project::path(parent));
    let path = crate::path::Path::parse(&text)
        .map(crate::path::Path::into_owned)
        .map_err(|_error| {
            Error::new(Code::InternalError, "the hyperlink path does not parse")
                .with_format(FileFormat::Ppt)
        })?;
    Ok((commit, path))
}

/// Set the formatting and/or text of every selected run in one native edit.
///
/// Formatting splits the body's character runs at the run's ends. A text
/// replaces the run's characters within its shape's whole text, so a changed
/// length follows the shape-text closure rules (refused over several
/// formatting runs).
pub(super) fn set_runs(
    file: &Ppt,
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    use super::super::spec::{RUN_WRITE_TABLE, RunSetter};
    use litchi_ppt::text_edit::RunFormat;

    let mut format = RunFormat::default();
    let mut text = None;
    let mut assigned = 0_u32;
    for (key, value) in props.iter() {
        let (index, entry) = RUN_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ppt)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u32 << index;
        if assigned & bit != 0 {
            return Err(Error::new(
                Code::InvalidArgument,
                "a PPT run property was assigned more than once",
            )
            .with_format(FileFormat::Ppt)
            .with_property(entry.spec.name));
        }
        assigned |= bit;
        entry
            .prepare(Operation::Set, value, |setter, scalar| {
                match (setter, scalar) {
                    (RunSetter::Text, scalar) => {
                        let value = match scalar {
                            Some(Scalar::Text(value)) => value,
                            _ => "",
                        };
                        bound(
                            LimitKind::StagedBytes,
                            value.len(),
                            limits.max_staged_bytes(),
                        )?;
                        text = Some(value.to_owned());
                    },
                    (RunSetter::Bold, Some(Scalar::Bool(value))) => format.bold = Some(value),
                    (RunSetter::Italic, Some(Scalar::Bool(value))) => format.italic = Some(value),
                    (RunSetter::Underline, Some(Scalar::Token(value))) => {
                        format.underline = Some(value != "none");
                    },
                    (RunSetter::Size, Some(Scalar::Length(length))) => {
                        let points = length.points();
                        if !(1.0..=4000.0).contains(&points) || points.fract() != 0.0 {
                            return Err(invalid(
                                "a PPT run size is a whole number of points from 1 to 4000",
                            ));
                        }
                        // In range by the check above.
                        format.size = Some(points as u16);
                    },
                    (RunSetter::Color, Some(Scalar::Color(color))) => match color.base() {
                        crate::value::Base::Rgb(rgba) if rgba.alpha() == 255 => {
                            format.color = Some([rgba.red(), rgba.green(), rgba.blue()]);
                        },
                        _ => {
                            return Err(invalid("a PPT run color is an opaque RGB color")
                                .with_suggestion("for example 1F4E79"));
                        },
                    },
                    _ => unreachable!("PPT run coercion returned an unexpected scalar"),
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ppt))?;
    }
    let formatted = format != RunFormat::default();
    let mut formats = Vec::new();
    let mut texts = std::collections::BTreeMap::<
        (usize, usize),
        (litchi_ppt::text_edit::Target, Vec<u16>, Vec<(usize, usize)>),
    >::new();
    for node in nodes {
        let (Some(target), Some(Element::Run { body, piece, .. })) =
            (node.native_shape_target(), node.element_item())
        else {
            return Err(unsupported_target(
                "nested PPT group run writes are not integrated",
            ));
        };
        if formatted {
            formats.push((target, piece.start, piece.end - piece.start));
        }
        if text.is_some() {
            let entry = texts
                .entry((target.slide().get(), target.shape().get()))
                .or_insert_with(|| (target, body.text().encode_utf16().collect(), Vec::new()));
            entry.2.push((piece.start, piece.end));
        }
    }
    let mut replacements = Vec::with_capacity(texts.len());
    if let Some(text) = &text {
        let units = text.encode_utf16().collect::<Vec<_>>();
        for (target, mut body, mut ranges) in texts.into_values() {
            ranges.sort_unstable_by(|left, right| right.cmp(left));
            for (start, end) in ranges {
                body.splice(start..end, units.iter().copied());
            }
            replacements.push((target, String::from_utf16_lossy(&body)));
        }
    }
    let operations = formats.len() + replacements.len();
    let native = stage_structural(file, operations, limits, |edit| {
        for (target, start, length) in &formats {
            edit.set_run_format(*target, *start, *length, &format)?;
        }
        for (target, text) in replacements {
            edit.set_shape_text(target, text)?;
        }
        Ok(())
    })?;
    file.check_source()?;
    finish_commit(file, native, limits)
}

/// Dispatch a `set` whose targets are record-0323 elements; `None` for other
/// targets. One call's targets share one element kind.
pub(super) fn set(
    file: &Ppt,
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Option<Result<WriteCommit>> {
    let kind = nodes.first()?.element();
    if !matches!(
        kind,
        Kind::Comment | Kind::Cell | Kind::Hyperlink | Kind::Run
    ) {
        return None;
    }
    if nodes.iter().any(|node| node.element() != kind) {
        return Some(Err(unsupported_target(
            "one PPT write cannot mix element kinds",
        )));
    }
    Some(match kind {
        Kind::Comment => set_comments(file, nodes, props, limits),
        Kind::Cell => set_cells(file, nodes, props, limits),
        Kind::Hyperlink => set_hyperlinks(file, nodes, props, limits),
        _ => set_runs(file, nodes, props, limits),
    })
}
