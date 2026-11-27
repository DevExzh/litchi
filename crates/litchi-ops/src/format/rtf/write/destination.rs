//! Unified writes to retained RTF destination stories and pictures.
//!
//! Table cells, header/footer paragraphs, footnotes/endnotes and comments are
//! separate native text stories. Their text is replaced through the native
//! destination editor, which splices the replacement into the retained
//! source; pictures are removed through the native exact-source picture
//! editor. Every selection publishes as one native transaction.

use litchi_rtf::HeaderFooterType;
use litchi_rtf::edit::{EditOptions, HeaderFooterParagraph, SemanticLimits as NativeLimits};
use litchi_rtf::table::TableCellPath;

use super::super::node::Item;
use super::super::{Node, Rtf};
use super::{Commit, bound, patch_snapshot_bytes, snapshot_limit, staged_limit};
use crate::error::{Code, Error, Result};
use crate::props::Props;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_core::FileFormat;

/// One retained destination story selected by a mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) enum Target {
    /// A top-level table cell: table, row and cell positions.
    Cell(usize, usize, usize),
    /// One header/footer paragraph: section, story position, paragraph.
    StoryParagraph(usize, usize, usize),
    /// A whole single-paragraph header/footer: section and story position.
    Story(usize, usize),
    /// A footnote or endnote by native note index.
    Note(usize),
    /// A comment by native annotation index.
    Comment(usize),
}

impl Target {
    /// The destination story `node` stands for, if any.
    pub(super) fn of(node: &Node<'_>) -> Option<Self> {
        Some(match node.item {
            Item::Cell {
                table, row, cell, ..
            } => Self::Cell(table, row, cell),
            Item::StoryParagraph {
                section,
                story,
                index,
                ..
            } => Self::StoryParagraph(section, story, index),
            Item::Story { section, story, .. } => Self::Story(section, story),
            Item::Note { index, .. } => Self::Note(index),
            Item::Comment { index, .. } => Self::Comment(index),
            _ => return None,
        })
    }
}

pub(super) fn append(targets: &mut Vec<Target>, node: &Node<'_>, maximum: usize) -> Result<()> {
    let target = Target::of(node).ok_or_else(|| {
        Error::new(
            Code::UnsupportedType,
            "RTF destination writes target cells, header/footer stories, notes or comments",
        )
        .with_format(FileFormat::Rtf)
    })?;
    if targets.len() >= maximum {
        return Err(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: targets.len().checked_add(1).unwrap_or(usize::MAX),
            maximum,
        }
        .into());
    }
    targets
        .try_reserve(1)
        .map_err(|source| TransactionError::Allocation { source })?;
    targets.push(target);
    Ok(())
}

pub(super) fn append_picture(
    targets: &mut Vec<usize>,
    node: &Node<'_>,
    maximum: usize,
) -> Result<()> {
    let Item::Picture { index, .. } = node.item else {
        return Err(
            Error::new(Code::UnsupportedType, "RTF picture writes target pictures")
                .with_format(FileFormat::Rtf),
        );
    };
    if targets.len() >= maximum {
        return Err(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: targets.len().checked_add(1).unwrap_or(usize::MAX),
            maximum,
        }
        .into());
    }
    targets
        .try_reserve(1)
        .map_err(|source| TransactionError::Allocation { source })?;
    targets.push(index);
    Ok(())
}

/// The replacement text of a destination `set`: exactly one `text` key.
pub(super) fn replacement_text<'a>(props: &'a Props<'_>) -> Result<&'a str> {
    let mut text = None;
    for (key, value) in props.iter() {
        if !key.eq_ignore_ascii_case("text") {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "RTF destination stories accept only a `text` property",
            )
            .with_property(key.as_str().to_owned())
            .with_format(FileFormat::Rtf));
        }
        text = Some(value.as_text().ok_or_else(|| {
            Error::new(
                Code::InvalidValue,
                "RTF destination text requires a text value",
            )
            .with_property("text")
            .with_format(FileFormat::Rtf)
        })?);
    }
    text.ok_or_else(|| {
        Error::new(
            Code::MissingProperty,
            "RTF destination writes require a `text` property",
        )
        .with_property("text")
        .with_format(FileFormat::Rtf)
    })
}

fn native_options(limits: &SemanticLimits) -> Result<EditOptions> {
    let native = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    Ok(EditOptions::new().semantic_limits(native))
}

/// The native header/footer kind of story `story` in section `section`.
/// The native header/footer kind of story `story` in section `section`. The
/// native editor addresses a story by its section and kind, so a story that
/// repeats an earlier kind in its section cannot be addressed and is refused.
fn story_kind(file: &Rtf, section: usize, story: usize) -> Result<HeaderFooterType> {
    let stories = file
        .0
        .package
        .sections()
        .get(section)
        .map(|value| value.headers_footers.as_slice())
        .unwrap_or_default();
    let kind = stories
        .get(story)
        .map(|value| value.header_type)
        .ok_or_else(|| {
            Error::new(Code::PathNotFound, "the RTF header/footer story is missing")
                .with_format(FileFormat::Rtf)
        })?;
    if stories.iter().position(|value| value.header_type == kind) != Some(story) {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "the RTF section repeats this header/footer kind; only its first story is editable",
        )
        .with_format(FileFormat::Rtf));
    }
    Ok(kind)
}

/// The text a destination currently holds, for admission accounting.
pub(super) fn current_len(file: &Rtf, target: Target) -> usize {
    let package = &file.0.package;
    match target {
        Target::Cell(table, row, cell) => package
            .tables()
            .get(table)
            .and_then(|value| value.rows().get(row))
            .and_then(|value| value.cells().get(cell))
            .map_or(0, |value| value.text().len()),
        Target::StoryParagraph(section, story, paragraph) => package
            .sections()
            .get(section)
            .and_then(|value| value.headers_footers.get(story))
            .and_then(|value| value.paragraphs.get(paragraph))
            .map_or(0, |value| value.text.len()),
        Target::Story(section, story) => package
            .sections()
            .get(section)
            .and_then(|value| value.headers_footers.get(story))
            .map_or(0, |value| {
                value
                    .paragraphs
                    .iter()
                    .map(|paragraph| paragraph.text.len())
                    .sum()
            }),
        Target::Note(index) => package
            .notes()
            .get(index)
            .map_or(0, |value| value.content.len()),
        Target::Comment(index) => package
            .annotations()
            .get(index)
            .map_or(0, |value| value.text.len()),
    }
}

/// A destination named by an exact path, before native resolution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DirectPath {
    /// `/body/tbl[N]/tr[M]/tc[K]` as zero-based positions.
    Cell(usize, usize, usize),
    /// `/header[N]` or `/footer[N]` with an optional `p[M]`.
    Story {
        headers: bool,
        story: usize,
        paragraph: Option<usize>,
    },
    /// `/footnote[N]` or `/endnote[N]` as a zero-based ordinal.
    Note { footnotes: bool, ordinal: usize },
    /// `/comments/comment[N]` as a zero-based ordinal.
    Comment(usize),
}

fn missing(what: &'static str) -> Error {
    Error::new(Code::PathNotFound, what).with_format(FileFormat::Rtf)
}

/// Resolve an exact destination path against the snapshot. A whole
/// header/footer resolves to its single paragraph so a later paragraph path
/// to the same story is recognized as the same native destination.
pub(super) fn resolve(file: &Rtf, path: DirectPath) -> Result<Target> {
    let package = &file.0.package;
    Ok(match path {
        DirectPath::Cell(table, row, cell) => {
            package
                .tables()
                .get(table)
                .and_then(|value| value.rows().get(row))
                .and_then(|value| value.cells().get(cell))
                .ok_or_else(|| missing("the RTF table cell does not exist"))?;
            Target::Cell(table, row, cell)
        },
        DirectPath::Story {
            headers,
            story,
            paragraph,
        } => {
            let found = super::super::stories::header_footers(file, headers)
                .nth(story)
                .ok_or_else(|| missing("the RTF header/footer story does not exist"))?;
            let paragraph = match paragraph {
                Some(paragraph) => paragraph,
                None if found.value.paragraphs.len() == 1 => 0,
                None => return Err(single_paragraph_error()),
            };
            if paragraph >= found.value.paragraphs.len() {
                return Err(missing("the RTF header/footer paragraph does not exist"));
            }
            Target::StoryParagraph(found.section, found.story, paragraph)
        },
        DirectPath::Note { footnotes, ordinal } => Target::Note(
            super::super::stories::notes(file, footnotes)
                .nth(ordinal)
                .ok_or_else(|| missing("the RTF note does not exist"))?
                .index,
        ),
        DirectPath::Comment(index) => {
            package
                .annotations()
                .get(index)
                .ok_or_else(|| missing("the RTF comment does not exist"))?;
            Target::Comment(index)
        },
    })
}

fn single_paragraph_error() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "only a single-paragraph RTF header/footer accepts `text`; set its paragraph instead",
    )
    .with_suggestion("address the paragraph, e.g. /header[1]/p[1]")
    .with_format(FileFormat::Rtf)
}

/// The text a destination currently holds.
pub(super) fn current_text(file: &Rtf, target: Target) -> &str {
    let package = &file.0.package;
    match target {
        Target::Cell(table, row, cell) => package
            .tables()
            .get(table)
            .and_then(|value| value.rows().get(row))
            .and_then(|value| value.cells().get(cell))
            .map_or("", |value| value.text()),
        Target::StoryParagraph(section, story, paragraph) => package
            .sections()
            .get(section)
            .and_then(|value| value.headers_footers.get(story))
            .and_then(|value| value.paragraphs.get(paragraph))
            .map_or("", |value| value.text.as_ref()),
        Target::Story(section, story) => package
            .sections()
            .get(section)
            .and_then(|value| value.headers_footers.get(story))
            .and_then(|value| value.paragraphs.first())
            .map_or("", |value| value.text.as_ref()),
        Target::Note(index) => package
            .notes()
            .get(index)
            .map_or("", |value| value.content.as_ref()),
        Target::Comment(index) => package
            .annotations()
            .get(index)
            .map_or("", |value| value.text.as_ref()),
    }
}

/// Stage one destination text replacement on a native edit.
pub(super) fn stage(
    edit: &mut litchi_rtf::edit::Edit,
    file: &Rtf,
    target: Target,
    text: &str,
) -> Result<()> {
    let staged = match target {
        Target::Cell(table, row, cell) => {
            edit.set_table_cell_text(TableCellPath::outer(table, row, cell), text)
        },
        Target::StoryParagraph(section, story, paragraph) => {
            let kind = story_kind(file, section, story)?;
            edit.set_header_footer_text(HeaderFooterParagraph::new(section, kind, paragraph), text)
        },
        Target::Story(section, story) => {
            let paragraphs = file
                .0
                .package
                .sections()
                .get(section)
                .and_then(|value| value.headers_footers.get(story))
                .map_or(0, |value| value.paragraphs.len());
            if paragraphs != 1 {
                return Err(single_paragraph_error());
            }
            let kind = story_kind(file, section, story)?;
            edit.set_header_footer_text(HeaderFooterParagraph::new(section, kind, 0), text)
        },
        Target::Note(index) => edit.set_note_text(index, text),
        Target::Comment(index) => edit.set_annotation_text(index, text),
    };
    staged.map(|_| ()).map_err(Error::from)
}

/// Replace the text of every selected destination story in one native edit.
pub(super) fn publish_texts(
    file: &Rtf,
    targets: &[Target],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let text = replacement_text(props)?;
    bound(
        LimitKind::Operations,
        targets.len(),
        limits.max_operations(),
    )?;
    // Native staging retains both the replacement and the replaced text for
    // inverse history; charge both before any owned copy is made.
    let mut staged = 0usize;
    for target in targets {
        staged = staged
            .checked_add(text.len())
            .and_then(|bytes| bytes.checked_add(current_len(file, *target)))
            .and_then(|bytes| bytes.checked_add(size_of::<Target>()))
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    }
    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;

    let mut edit = file.0.package.edit_with_options(native_options(limits)?);
    for target in targets {
        stage(&mut edit, file, *target, text)?;
    }
    finish(file, edit.commit().map_err(Error::from)?, limits)
}

/// Remove the selected pictures through the native exact-source editor.
pub(super) fn remove_pictures(
    file: &Rtf,
    indices: &[usize],
    limits: &SemanticLimits,
) -> Result<Commit> {
    bound(
        LimitKind::Operations,
        indices.len(),
        limits.max_operations(),
    )?;
    let staged = indices.iter().try_fold(0usize, |total, index| {
        let bytes = file
            .0
            .package
            .pictures()
            .get(*index)
            .map_or(0, |picture| picture.data.len());
        total
            .checked_add(bytes.saturating_mul(2))
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
    })?;
    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
    let mut edit = file.0.package.edit_with_options(native_options(limits)?);
    edit.remove_pictures(indices).map_err(Error::from)?;
    finish(file, edit.commit().map_err(Error::from)?, limits)
}

fn finish(file: &Rtf, native: litchi_rtf::edit::Commit, limits: &SemanticLimits) -> Result<Commit> {
    let changed = native.diagnostics().changed();
    let operations = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let after = if changed {
        file.with_package(package)
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before: file.clone(),
        after,
        patch,
        operations,
        snapshot_bytes,
    })
}
