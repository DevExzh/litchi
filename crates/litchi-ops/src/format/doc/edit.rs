//! Unified DOC writes over the primary native body transaction
//! (`litchi_doc::Package::edit`).
//!
//! Targets resolve against the caller's immutable handle. One call stages
//! every selected operation into one native edit and publishes one commit:
//!
//! - `set` on `/body/p[N]`: `text`, and run formatting (`bold`, `italic`,
//!   `underline`, `size`, `color`, `font`) over the whole paragraph;
//! - `set` on `/body/p[N]/r[M]`: the same keys on one run;
//! - `add` of a paragraph under `/body`, `remove` of paragraphs or runs,
//!   `move_to` and `swap` of paragraphs.
//!
//! A formatting key whose value every selected run already reads as is not
//! staged, so an assignment that changes nothing publishes the source handle
//! unchanged. Signed handles refuse before staging, and the native editor
//! refuses encrypted sources.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_doc::body_text::{CharacterProperty, CharacterValue, TextTarget};

use super::{Doc, Nav, Node, StepPatch, WriteCommit};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::{IntoProps, PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::{Base, ColorContext, LengthSpec, Scalar, ValueType};

/// One formatting assignment of a `set` call.
#[derive(Debug, Clone, PartialEq, Eq)]
enum Format {
    Bold(bool),
    Italic(bool),
    Underline(bool),
    Size(u16),
    Color(Option<[u8; 3]>),
    Font(String),
    /// A character style, by stylesheet index (record 0324).
    Style(u16),
}

impl Format {
    /// Whether `run` already reads as this value.
    fn holds(&self, run: &litchi_doc::Run, font: Option<&str>) -> bool {
        match self {
            Self::Bold(value) => run.bold().unwrap_or(false) == *value,
            Self::Italic(value) => run.italic().unwrap_or(false) == *value,
            Self::Underline(value) => run.underline().is_some() == *value,
            Self::Size(value) => run.font_size() == Some(*value),
            Self::Color(value) => {
                run.color().map(|(red, green, blue)| [red, green, blue]) == *value
            },
            Self::Font(value) => font == Some(value.as_str()),
            Self::Style(istd) => {
                run.properties()
                    .style_index
                    .unwrap_or(super::parts::DEFAULT_CHARACTER_STYLE)
                    == *istd
            },
        }
    }

    fn stage(&self, edit: &mut litchi_doc::Edit<'static>, target: TextTarget) -> Result<()> {
        match self {
            Self::Bold(value) => {
                edit.set_character_property(target, CharacterProperty::Bold, *value)
            },
            Self::Italic(value) => {
                edit.set_character_property(target, CharacterProperty::Italic, *value)
            },
            Self::Underline(value) => {
                edit.set_character_property(target, CharacterProperty::Underline, *value)
            },
            Self::Size(value) => edit.set_character_value(target, CharacterValue::Size(*value)),
            Self::Color(value) => edit.set_character_value(target, CharacterValue::Color(*value)),
            Self::Font(value) => edit.set_character_value(target, CharacterValue::Font(value)),
            Self::Style(istd) => edit.set_character_value(target, CharacterValue::Style(*istd)),
        }
        .map_err(Error::from)
    }
}

/// The parsed properties of one `set` call.
#[derive(Debug, Default)]
struct Assignment {
    text: Option<String>,
    formats: Vec<Format>,
    /// `style`, `styleId` or `styleName`: a paragraph style, by id or name.
    style: Option<String>,
    /// `rStyle`: a character style, by id or name.
    run_style: Option<String>,
}

impl Assignment {
    fn parse(props: &Props<'_>) -> Result<Self> {
        let mut assignment = Self::default();
        for (key, value) in props.iter() {
            let name = key.as_str();
            if name.eq_ignore_ascii_case("text") {
                assignment.text = Some(
                    value
                        .as_text()
                        .ok_or_else(|| invalid("DOC text requires a text property value", "text"))?
                        .to_owned(),
                );
            } else if name.eq_ignore_ascii_case("bold") {
                assignment
                    .formats
                    .push(Format::Bold(boolean(value, "bold")?));
            } else if name.eq_ignore_ascii_case("italic") {
                assignment
                    .formats
                    .push(Format::Italic(boolean(value, "italic")?));
            } else if name.eq_ignore_ascii_case("underline") {
                assignment
                    .formats
                    .push(Format::Underline(underline(value)?));
            } else if name.eq_ignore_ascii_case("size") {
                assignment.formats.push(Format::Size(size(value)?));
            } else if name.eq_ignore_ascii_case("color") {
                assignment.formats.push(Format::Color(color(value)?));
            } else if ["style", "styleId", "styleName"]
                .iter()
                .any(|key| name.eq_ignore_ascii_case(key))
            {
                assignment.style = Some(style_value(value, "style")?);
            } else if name.eq_ignore_ascii_case("rStyle") {
                assignment.run_style = Some(style_value(value, "rStyle")?);
            } else if name.eq_ignore_ascii_case("font") {
                let font = value
                    .as_text()
                    .filter(|font| !font.is_empty())
                    .ok_or_else(|| invalid("DOC font requires a non-empty font name", "font"))?;
                assignment.formats.push(Format::Font(font.to_owned()));
            } else {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "the requested DOC property is not writable",
                )
                .with_property(name.to_owned())
                .with_format(FileFormat::Doc));
            }
        }
        Ok(assignment)
    }

    const fn len(&self) -> usize {
        self.formats.len()
            + if self.text.is_some() { 1 } else { 0 }
            + if self.style.is_some() { 1 } else { 0 }
            + if self.run_style.is_some() { 1 } else { 0 }
    }

    /// Resolve the style names against the document: a paragraph style
    /// becomes its index, a character style a [`Format::Style`].
    fn resolve_styles(&mut self, doc: &Doc) -> Result<Option<u16>> {
        if let Some(name) = self.run_style.take() {
            let istd = resolve_style(doc, &name, litchi_doc::parts::styles::StyleKind::Character)?;
            self.formats.push(Format::Style(istd));
        }
        self.style
            .as_deref()
            .map(|name| resolve_style(doc, name, litchi_doc::parts::styles::StyleKind::Paragraph))
            .transpose()
    }
}

fn style_value(value: &PropValue<'_>, key: &'static str) -> Result<String> {
    value
        .as_text()
        .filter(|name| !name.is_empty())
        .map(str::to_owned)
        .ok_or_else(|| invalid("DOC styles are named by a non-empty id or name", key))
}

/// The index of the style of `kind` named `name` (its id, its name or an
/// alternate name).
fn resolve_style(doc: &Doc, name: &str, kind: litchi_doc::parts::styles::StyleKind) -> Result<u16> {
    let not_found = || {
        Error::new(Code::NotFound, "the DOC stylesheet has no such style")
            .with_format(FileFormat::Doc)
            .with_suggestion("name a style listed under `/styles` by its id or name")
    };
    let istd = super::parts::find_style(doc, name).ok_or_else(not_found)?;
    let matches = doc
        .reader()?
        .stylesheet()
        .and_then(|sheet| sheet.get(istd))
        .is_some_and(|style| style.kind == kind);
    if !matches {
        return Err(Error::new(
            Code::InvalidValue,
            match kind {
                litchi_doc::parts::styles::StyleKind::Character => {
                    "`rStyle` names a character style"
                },
                _ => "a paragraph `style` names a paragraph style",
            },
        )
        .with_format(FileFormat::Doc));
    }
    Ok(istd)
}

fn invalid(message: &'static str, key: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_property(key)
        .with_format(FileFormat::Doc)
}

/// A direct paragraph `bold` or `italic` assignment a coalesced body run
/// stages (record 0359a): the leaf property, the requested value, and
/// whether every run of the resolved paragraph already reads as it (the
/// isolated setter's no-op test).
pub(super) fn paragraph_toggle(
    node: &Node<'_>,
    key: &str,
    value: &PropValue<'_>,
) -> Result<Option<(CharacterProperty, bool, bool)>> {
    let (property, format) = if key.eq_ignore_ascii_case("bold") {
        let enabled = boolean(value, "bold")?;
        (CharacterProperty::Bold, Format::Bold(enabled))
    } else if key.eq_ignore_ascii_case("italic") {
        let enabled = boolean(value, "italic")?;
        (CharacterProperty::Italic, Format::Italic(enabled))
    } else {
        return Ok(None);
    };
    let enabled = matches!(format, Format::Bold(true) | Format::Italic(true));
    let Target::Paragraph { .. } = Target::of(node, false)? else {
        return Ok(None);
    };
    let holds = Target::of(node, true)?.holds(&format);
    Ok(Some((property, enabled, holds)))
}

fn boolean(value: &PropValue<'_>, key: &'static str) -> Result<bool> {
    value
        .with_coerced(ValueType::Bool, |value| match value {
            Some(Scalar::Bool(value)) => Ok(value),
            _ => Err(invalid("DOC character formatting requires a boolean", key)),
        })
        .map_err(|error| error.with_property(key).with_format(FileFormat::Doc))
}

fn underline(value: &PropValue<'_>) -> Result<bool> {
    value
        .with_coerced(
            ValueType::Token(&super::spec::UNDERLINE),
            |value| match value {
                Some(Scalar::Token("none")) => Ok(false),
                Some(Scalar::Token("single")) => Ok(true),
                Some(Scalar::Token(_)) => Err(Error::new(
                    Code::UnsupportedCapability,
                    "DOC run writes accept the `none` and `single` underline styles",
                )
                .with_property("underline")
                .with_format(FileFormat::Doc)),
                _ => Err(invalid(
                    "DOC underline requires an underline token",
                    "underline",
                )),
            },
        )
        .map_err(|error| {
            error
                .with_property("underline")
                .with_format(FileFormat::Doc)
        })
}

fn size(value: &PropValue<'_>) -> Result<u16> {
    value
        .with_coerced(ValueType::Length(LengthSpec::FONT_SIZE), |value| {
            let Some(Scalar::Length(value)) = value else {
                return Err(invalid("DOC size requires a font-size length", "size"));
            };
            let emu = value.emu();
            if emu <= 0 || emu % 6_350 != 0 {
                return Err(invalid(
                    "DOC size must be a positive whole half-point",
                    "size",
                ));
            }
            u16::try_from(emu / 6_350)
                .ok()
                .filter(|half_points| (2..=3276).contains(half_points))
                .ok_or_else(|| invalid("DOC size must be 1 through 1638 points", "size"))
        })
        .map_err(|error| error.with_property("size").with_format(FileFormat::Doc))
}

fn color(value: &PropValue<'_>) -> Result<Option<[u8; 3]>> {
    value
        .with_coerced(ValueType::Color(ColorContext::Word), |value| {
            let Some(Scalar::Color(color)) = value else {
                return Err(invalid("DOC color requires a color value", "color"));
            };
            if !color.transforms().is_empty() {
                return Err(unsupported_color());
            }
            match color.base() {
                Base::Auto => Ok(None),
                Base::Rgb(rgba) if rgba.alpha() == 255 => {
                    Ok(Some([rgba.red(), rgba.green(), rgba.blue()]))
                },
                _ => Err(unsupported_color()),
            }
        })
        .map_err(|error| error.with_property("color").with_format(FileFormat::Doc))
}

fn unsupported_color() -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "DOC run colors are opaque RGB values or `auto`",
    )
    .with_property("color")
    .with_format(FileFormat::Doc)
}

/// One resolved write target.
#[derive(Debug, Clone)]
enum Target {
    Paragraph {
        position: Position,
        /// The formatting of each run, read when the assignment formats.
        runs: Vec<(litchi_doc::Run, Option<String>)>,
    },
    Run {
        paragraph: Position,
        start: u32,
        end: u32,
        run: Box<litchi_doc::Run>,
        font: Option<String>,
    },
}

impl Target {
    fn of(node: &Node<'_>, read_runs: bool) -> Result<Self> {
        if let Some(position) = node.paragraph_position() {
            let mut runs = Vec::new();
            if read_runs && let Some(paragraph) = node.native_paragraph()? {
                for run in paragraph.runs()? {
                    let font = font_name(node.doc, &run)?;
                    runs.push((run, font));
                }
            }
            return Ok(Self::Paragraph { position, runs });
        }
        if let (Some((paragraph, start, end)), Some(run)) = (node.body_run(), node.run()) {
            return Ok(Self::Run {
                paragraph,
                start,
                end,
                font: font_name(node.doc, run)?,
                run: Box::new(run.clone()),
            });
        }
        Err(Error::new(
            Code::UnsupportedCapability,
            "DOC writes target ordinary main-story paragraphs and their runs",
        )
        .with_format(FileFormat::Doc)
        .with_suggestion("select `/body/p[N]` or `/body/p[N]/r[M]`"))
    }

    const fn order(&self) -> (usize, u32) {
        match self {
            Self::Paragraph { position, .. } => (position.get(), 0),
            Self::Run {
                paragraph, start, ..
            } => (paragraph.get(), *start),
        }
    }

    const fn leaf(&self) -> TextTarget {
        match self {
            Self::Paragraph { position, .. } => TextTarget::body_paragraph(*position),
            Self::Run {
                paragraph,
                start,
                end,
                ..
            } => TextTarget::BodySpan {
                paragraph: *paragraph,
                start: *start,
                end: *end,
            },
        }
    }

    /// Whether every run of the target already reads as `format`.
    fn holds(&self, format: &Format) -> bool {
        match self {
            Self::Paragraph { runs, .. } => {
                !runs.is_empty()
                    && runs
                        .iter()
                        .all(|(run, font)| format.holds(run, font.as_deref()))
            },
            Self::Run { run, font, .. } => format.holds(run, font.as_deref()),
        }
    }
}

fn font_name(doc: &Doc, run: &litchi_doc::Run) -> Result<Option<String>> {
    Ok(match run.font_index() {
        Some(ftc) => doc.reader()?.font_name(ftc)?.map(str::to_owned),
        None => None,
    })
}

/// Resolve a path or scoped query to write targets, deduplicated.
fn resolve_targets(
    doc: &Doc,
    target: &TargetArg<'_>,
    maximum: usize,
    read_runs: bool,
) -> Result<Vec<Target>> {
    let mut targets: Vec<Target> = Vec::new();
    let mut push = |node: &Node<'_>| -> Result<()> {
        let target = Target::of(node, read_runs)?;
        if targets.iter().any(|seen| {
            seen.order() == target.order()
                && std::mem::discriminant(seen) == std::mem::discriminant(&target)
        }) {
            return Ok(());
        }
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
    };
    if let Some(path) = target.path() {
        push(&crate::path::resolve::get(&Nav::new(doc), path)?)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(doc), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the DOC mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Doc));
        }
        for node in walk {
            push(&node?)?;
        }
    }
    if targets.is_empty() {
        return Err(Error::new(
            Code::ZeroMatches,
            "the DOC mutation selector matched no paragraph or run",
        )
        .with_format(FileFormat::Doc));
    }
    let paragraphs = targets
        .iter()
        .filter(|target| matches!(target, Target::Paragraph { .. }))
        .count();
    if paragraphs != 0 && paragraphs != targets.len() {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "one DOC write selects either paragraphs or runs, not both",
        )
        .with_format(FileFormat::Doc));
    }
    // Later spans first, so a run's text change never moves a span staged
    // after it.
    targets.sort_by_key(|target| std::cmp::Reverse(target.order()));
    Ok(targets)
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
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

impl Doc {
    /// The retained native owner, or the typed refusal for a released one.
    fn edit_source(&self) -> Result<litchi_doc::Package> {
        self.editable_package_for_write()?.ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "the DOC document was opened without retaining its source for editing",
            )
            .with_format(FileFormat::Doc)
            .with_suggestion("reopen the document without `OpenOptions::retain_source(false)`")
        })
    }

    pub(super) fn native_edit(&self, limits: &SemanticLimits) -> Result<litchi_doc::Edit<'static>> {
        self.preflight_snapshot_bytes(limits.max_snapshot_bytes(), 0)?;
        let source_bytes = self
            .retained_source_len()
            .map_or(0, |bytes| usize::try_from(bytes).unwrap_or(usize::MAX));
        Ok(self.edit_source()?.edit_with_options(super::edit_options(
            super::body_transaction_limits_with_source(
                *limits,
                source_bytes,
                limits.max_snapshot_bytes(),
            ),
        ))?)
    }

    /// Commit a staged native edit into a write step. An edit that changed
    /// no byte keeps this handle.
    pub(super) fn publish(
        &self,
        edit: litchi_doc::Edit<'static>,
        operations: usize,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let staged_snapshot_bytes = super::patch_snapshot_bytes(commit.patch());
        limit(
            LimitKind::SnapshotBytes,
            staged_snapshot_bytes,
            limits.max_snapshot_bytes(),
        )?;
        if changed && self.signature() == super::Signature::Signed {
            return Err(super::signature::refusal(FileFormat::Doc));
        }
        let (package, patch) = if !changed && self.signature() == super::Signature::Signed {
            // The staged editor intentionally uses an unsigned private copy.
            // Rebuild an identity patch over the retained signed package so
            // native replay sees the same physical base as the facade.
            let original = self.original_package_for_noop()?;
            self.preflight_snapshot_bytes(limits.max_snapshot_bytes(), 0)?;
            original.edit()?.commit()?.into_parts()
        } else {
            commit.into_parts()
        };
        let snapshot_bytes = staged_snapshot_bytes.max(super::patch_snapshot_bytes(&patch));
        limit(
            LimitKind::SnapshotBytes,
            snapshot_bytes,
            limits.max_snapshot_bytes(),
        )?;
        Ok(WriteCommit {
            before: self.clone(),
            after: if changed {
                self.with_package(package)
            } else {
                self.clone()
            },
            patch: StepPatch::Primary(patch),
            operations,
            snapshot_bytes,
        })
    }

    /// Apply `props` to the selected paragraphs or runs in one native edit.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source_for_write()?;
        let result = (|| -> Result<WriteCommit> {
            let props = Props::from_props(props)?;
            if crate::format::ole_metadata::is_target(&target) {
                return self.set_metadata(&props, limits);
            }
            if let Some(commit) = self.set_mark(&target, &props, limits)? {
                return Ok(commit);
            }
            let mut assignment = Assignment::parse(&props)?;
            if assignment.len() == 0 {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "a DOC write needs at least one property",
                )
                .with_format(FileFormat::Doc));
            }
            let paragraph_style = assignment.resolve_styles(self)?;
            let targets = resolve_targets(
                self,
                &target,
                limits.max_expanded_targets(),
                !assignment.formats.is_empty(),
            )?;
            if paragraph_style.is_some()
                && targets
                    .iter()
                    .any(|target| !matches!(target, Target::Paragraph { .. }))
            {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "a DOC paragraph style applies to paragraphs; set `rStyle` on runs",
                )
                .with_property("style")
                .with_format(FileFormat::Doc));
            }
            limit(
                LimitKind::Operations,
                targets.len().saturating_mul(assignment.len()),
                limits.max_operations(),
            )?;
            if let Some(text) = &assignment.text {
                limit(
                    LimitKind::StagedBytes,
                    text.len().saturating_mul(targets.len()),
                    limits.max_staged_bytes(),
                )?;
            }
            let mut edit = self.native_edit(limits)?;
            let mut operations = 0usize;
            for target in &targets {
                for format in &assignment.formats {
                    if target.holds(format) {
                        continue;
                    }
                    format.stage(&mut edit, target.leaf())?;
                    operations += 1;
                }
                if let (Some(istd), Target::Paragraph { position, .. }) = (paragraph_style, target)
                {
                    edit.set_paragraph_style(*position, istd)?;
                    operations += 1;
                }
                if let Some(text) = &assignment.text {
                    match target {
                        Target::Paragraph { position, .. } => {
                            edit.replace_paragraph(*position, text)?;
                            operations += 1;
                        },
                        Target::Run { run, .. } => {
                            if run.text()? != text {
                                edit.replace_text(target.leaf(), text)?;
                                operations += 1;
                            }
                        },
                    }
                }
            }
            self.publish(edit, operations, limits)
        })();
        self.check_source_for_write()?;
        result
    }

    /// [`Self::add_with_limits`] under the default semantic limits.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
    ) -> Result<(WriteCommit, Path<'static>)> {
        self.add_with_limits(parent, node, &SemanticLimits::default())
    }

    /// [`Self::remove_with_limits`] under the default semantic limits.
    pub(crate) fn remove<'p>(&self, target: TargetArg<'p>) -> Result<WriteCommit> {
        self.remove_with_limits(target, &SemanticLimits::default())
    }

    /// [`Self::move_to_with_limits`] under the default semantic limits.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
    ) -> Result<WriteCommit> {
        self.move_to_with_limits(path, placement, &SemanticLimits::default())
    }

    /// [`Self::swap_with_limits`] under the default semantic limits.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
    ) -> Result<WriteCommit> {
        self.swap_with_limits(left, right, &SemanticLimits::default())
    }

    /// Insert one ordinary paragraph under `/body`.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        self.check_source()?;
        let result = (|| {
            let nav = Nav::new(self);
            let parent = crate::path::resolve::get(&nav, &parent)?;
            if let Some(added) = self.add_mark(&parent, &node, limits)? {
                return Ok(added);
            }
            let kind = node.kind();
            if !kind.eq_ignore_ascii_case("p") && !kind.eq_ignore_ascii_case("paragraph") {
                return Err(Error::new(
                    Code::UnsupportedElement,
                    "DOC adds ordinary paragraphs only",
                )
                .with_format(FileFormat::Doc));
            }
            if parent.element_kind() != super::spec::Kind::Body {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "DOC paragraphs can only be added directly under the body",
                )
                .with_format(FileFormat::Doc));
            }
            let text = add_text(node.props())?;
            let count = self.body_paragraph_count()?;
            let position = add_position(&nav, count, node.placement())?;
            limit(LimitKind::Operations, 1, limits.max_operations())?;
            limit(
                LimitKind::StagedBytes,
                text.len(),
                limits.max_staged_bytes(),
            )?;
            let mut edit = self.native_edit(limits)?;
            edit.insert_paragraph(Position::new(position), text)?;
            Ok((self.publish(edit, 1, limits)?, paragraph_path(position)))
        })();
        self.check_source()?;
        result
    }

    /// Remove the selected paragraphs, or the text of the selected runs.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: TargetArg<'p>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source()?;
        let result = (|| {
            if let Some(commit) = self.remove_mark(&target, limits)? {
                return Ok(commit);
            }
            let targets = resolve_targets(self, &target, limits.max_expanded_targets(), false)?;
            limit(
                LimitKind::Operations,
                targets.len(),
                limits.max_operations(),
            )?;
            let mut edit = self.native_edit(limits)?;
            for target in &targets {
                match target {
                    Target::Paragraph { position, .. } => edit.remove_paragraph(*position)?,
                    Target::Run { .. } => edit.replace_text(target.leaf(), "")?,
                }
            }
            self.publish(edit, targets.len(), limits)
        })();
        self.check_source()?;
        result
    }

    /// Move one ordinary paragraph.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source()?;
        let result = (|| {
            let nav = Nav::new(self);
            let from = paragraph_position(&crate::path::resolve::get(&nav, &path)?)?;
            let count = self.body_paragraph_count()?;
            let to = move_destination(&nav, from, count, &placement)?;
            limit(LimitKind::Operations, 1, limits.max_operations())?;
            let mut edit = self.native_edit(limits)?;
            edit.move_paragraph(Position::new(from), Position::new(to))?;
            self.publish(edit, 1, limits)
        })();
        self.check_source()?;
        result
    }

    /// Exchange two ordinary paragraphs, each keeping its runs and paragraph
    /// formatting.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source()?;
        let result = (|| {
            let nav = Nav::new(self);
            let left = paragraph_position(&crate::path::resolve::get(&nav, &left)?)?;
            let right = paragraph_position(&crate::path::resolve::get(&nav, &right)?)?;
            let (first, second) = (left.min(right), left.max(right));
            limit(LimitKind::Operations, 2, limits.max_operations())?;
            let mut edit = self.native_edit(limits)?;
            if first != second {
                edit.move_paragraph(Position::new(second), Position::new(first))?;
                edit.move_paragraph(Position::new(first + 1), Position::new(second))?;
            }
            self.publish(edit, 2, limits)
        })();
        self.check_source()?;
        result
    }
}

fn add_text<'a>(props: &'a Props<'_>) -> Result<&'a str> {
    let mut text = "";
    for (key, value) in props.iter() {
        if !key.eq_ignore_ascii_case("text") {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "DOC paragraph additions accept a `text` property only",
            )
            .with_property(key.as_str().to_owned())
            .with_format(FileFormat::Doc));
        }
        text = value
            .as_text()
            .ok_or_else(|| invalid("DOC paragraph text requires a text value", "text"))?;
    }
    Ok(text)
}

fn paragraph_position(node: &Node<'_>) -> Result<usize> {
    node.paragraph_position().map(Position::get).ok_or_else(|| {
        Error::new(
            Code::UnsupportedType,
            "DOC structural edits target ordinary main-story paragraphs",
        )
        .with_format(FileFormat::Doc)
    })
}

fn add_position(nav: &Nav<'_>, count: usize, placement: Option<&Placement<'_>>) -> Result<usize> {
    let position = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => {
            paragraph_position(&crate::path::resolve::get(nav, path)?)?
        },
        Some(Placement::After(path)) => {
            paragraph_position(&crate::path::resolve::get(nav, path)?)?.saturating_add(1)
        },
        Some(Placement::Into(path, position)) => {
            if crate::path::resolve::get(nav, path)?.element_kind() != super::spec::Kind::Body {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "DOC paragraphs can only be added within the body",
                )
                .with_format(FileFormat::Doc));
            }
            position.map_or(count, Position::get)
        },
    };
    if position > count {
        return Err(Error::new(
            Code::OutOfRange,
            "the DOC paragraph insertion position is outside the body",
        )
        .with_format(FileFormat::Doc));
    }
    Ok(position)
}

fn move_destination(
    nav: &Nav<'_>,
    from: usize,
    count: usize,
    placement: &Placement<'_>,
) -> Result<usize> {
    let last = count.checked_sub(1).ok_or_else(|| {
        Error::new(Code::PathNotFound, "the DOC body has no paragraphs")
            .with_format(FileFormat::Doc)
    })?;
    let anchor = |path: &PathArg<'_>| paragraph_position(&crate::path::resolve::get(nav, path)?);
    let destination = match placement {
        Placement::Start => 0,
        Placement::End => last,
        Placement::Position(position) => position.get(),
        Placement::Before(path) => {
            let anchor = anchor(path)?;
            if from < anchor { anchor - 1 } else { anchor }
        },
        Placement::After(path) => {
            let anchor = anchor(path)?;
            if from <= anchor { anchor } else { anchor + 1 }
        },
        Placement::Into(path, position) => {
            if crate::path::resolve::get(nav, path)?.element_kind() != super::spec::Kind::Body {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "DOC paragraphs can only be moved within the body",
                )
                .with_format(FileFormat::Doc));
            }
            position.map_or(last, Position::get)
        },
    };
    if destination >= count {
        return Err(Error::new(
            Code::OutOfRange,
            "the DOC paragraph move position is outside the body",
        )
        .with_format(FileFormat::Doc));
    }
    Ok(destination)
}

fn paragraph_path(position: usize) -> Path<'static> {
    Path::from_steps([
        Step::new("body"),
        Step::new("p").with(Predicate::Index(Position::new(position))),
    ])
}
