//! Writes of the DOCX part elements (record 0319): `set`, `add` and
//! `remove` of styles, numbering definitions, instances and levels,
//! headers, footers, footnotes and endnotes.
//!
//! Every property is admitted through the element's registry before an
//! editor is staged, so a refused write publishes nothing. The verbs stage
//! the leaf's catalog edits on the owning package's editor (record 0251b),
//! so a source-backed handle is materialized once.

use litchi_core::selector::{Path, Predicate, Step as PathStep};
use litchi_core::{FileFormat, Position};
use litchi_docx::document::{ElementStep, NewNote, NewStory, NoteKind};
use litchi_docx::header_footer::{Kind as StoryKind, Role};
use litchi_docx::numbering::catalog::{
    AbstractNumChange, LevelChange, NewAbstractNum, NewLevel, NewNum,
};
use litchi_docx::properties as direct;
use litchi_docx::styles::Type as StyleType;
use litchi_docx::styles::catalog::{NewStyle, StyleChange};

use super::super::element::{element_path, refusal, run_locations};
use super::super::node::{Node, Value};
use super::super::spec::Kind;
use super::super::write::{Commit, WorkEdit};
use super::super::{Docx, DocxEditor};
use super::{
    ABSTRACT_NUM_REGISTRY, AbstractKey, LEVEL_REGISTRY, LevelKey, MULTI_LEVEL_TYPE, NOTE_REGISTRY,
    NUM_REGISTRY, NumKey, PartSetter, PartValue, STORY_REGISTRY, STYLE_REGISTRY, SeedKey, StyleKey,
};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::placement::Placement;
use crate::props::{PropValue, Props};
use crate::registry::{Entry, Operation};
use crate::value::Scalar;

/// The style identifiers OfficeCLI treats as Word built-ins: a new style
/// with one of them is not marked `customStyle`.
const BUILT_IN_STYLES: &[&str] = &[
    "Normal",
    "Heading1",
    "Heading2",
    "Heading3",
    "Heading4",
    "Heading5",
    "Heading6",
    "Heading7",
    "Heading8",
    "Heading9",
    "Title",
    "Subtitle",
    "Quote",
    "IntenseQuote",
    "ListParagraph",
    "NoSpacing",
    "TOCHeading",
    "DefaultParagraphFont",
    "TableNormal",
    "NoList",
];

/// A new part element, by the type name `add` was given.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in super::super) enum NewPart {
    Style,
    AbstractNum,
    Num,
    Level,
    Header,
    Footer,
    Footnote,
    Endnote,
}

impl NewPart {
    /// The new part element `name` names, if it is one.
    pub(in super::super) fn of(name: &str) -> Option<Self> {
        let is = |names: &[&str]| names.iter().any(|known| name.eq_ignore_ascii_case(known));
        Some(if is(&["style"]) {
            Self::Style
        } else if is(&["abstractnum"]) {
            Self::AbstractNum
        } else if is(&["num"]) {
            Self::Num
        } else if is(&["level", "lvl"]) {
            Self::Level
        } else if is(&["header"]) {
            Self::Header
        } else if is(&["footer"]) {
            Self::Footer
        } else if is(&["footnote"]) {
            Self::Footnote
        } else if is(&["endnote"]) {
            Self::Endnote
        } else {
            return None;
        })
    }
}

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

/// One admitted property: its registry entry and value.
struct Assignment<'p, 'v> {
    key: &'p str,
    entry: Entry<PartSetter>,
    value: &'p PropValue<'v>,
}

/// Admit every property of `props` against `table` for `operation`; a
/// refused key fails before anything is staged.
fn admit<'p, 'v, const N: usize, const A: usize>(
    table: &crate::registry::Table<PartSetter, N, A>,
    operation: Operation,
    props: &'p Props<'v>,
) -> Result<Vec<Assignment<'p, 'v>>> {
    let mut admitted = Vec::with_capacity(props.len());
    for (key, value) in props.iter() {
        let (_, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_property(key.as_str().to_owned())
                .with_format(FileFormat::Docx)
        })?;
        // Coerce and check the value now; the staged write repeats it.
        entry
            .prepare(operation, value, |_, _| Ok(()))
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
        admitted.push(Assignment {
            key: key.as_str(),
            entry: *entry,
            value,
        });
    }
    Ok(admitted)
}

/// Run `stage` with the assignment's setter and coerced value.
fn stage<T>(
    assignment: &Assignment<'_, '_>,
    operation: Operation,
    stage: impl FnOnce(PartSetter, Option<Scalar<'_>>) -> Result<T>,
) -> Result<T> {
    assignment
        .entry
        .prepare(operation, assignment.value, stage)
        .map_err(|error| {
            error
                .with_property(assignment.key.to_owned())
                .with_format(FileFormat::Docx)
        })
}

fn text_of(value: Option<Scalar<'_>>) -> Option<String> {
    match value {
        Some(Scalar::Text(text) | Scalar::Token(text)) => Some(text.to_owned()),
        Some(other) => Some(other.to_string()),
        None => None,
    }
}

fn int_of(value: Option<Scalar<'_>>) -> Option<i64> {
    match value {
        Some(Scalar::Int(value)) => Some(value),
        _ => None,
    }
}

fn bool_of(value: Option<Scalar<'_>>) -> Option<bool> {
    match value {
        Some(Scalar::Bool(value)) => Some(value),
        _ => None,
    }
}

/// The style identifier `key` names: an identifier, else a display name,
/// else `key` itself (OfficeCLI's lenient style reference).
pub(in super::super) fn resolve_style_id(docx: &Docx, key: &str) -> Result<String> {
    Ok(match super::styles(docx)? {
        Some(catalog) => catalog
            .get(key)
            .or_else(|| catalog.resolve(key))
            .map_or_else(|| key.to_owned(), |entry| entry.id().to_owned()),
        None => key.to_owned(),
    })
}

/// Apply run and paragraph assignments to typed property targets.
fn apply_formatting(
    assignments: &[&Assignment<'_, '_>],
    operation: Operation,
    paragraph: &mut Option<direct::Paragraph>,
    run: &mut Option<direct::Run>,
    current_paragraph: impl FnOnce() -> Result<direct::Paragraph>,
    current_run: impl FnOnce() -> Result<direct::Run>,
) -> Result<()> {
    let mut current_paragraph = Some(current_paragraph);
    let mut current_run = Some(current_run);
    for assignment in assignments {
        match assignment.entry.access {
            crate::registry::Access::SetAndAdd(PartSetter::Paragraph(_))
            | crate::registry::Access::Set(PartSetter::Paragraph(_)) => {
                if paragraph.is_none()
                    && let Some(read) = current_paragraph.take()
                {
                    *paragraph = Some(read()?);
                }
                let target = paragraph
                    .as_mut()
                    .ok_or_else(|| error(Code::InternalError, "no paragraph target"))?;
                stage(assignment, operation, |setter, value| match setter {
                    PartSetter::Paragraph(setter) => setter.apply_paragraph(target, value),
                    _ => Ok(()),
                })?;
            },
            crate::registry::Access::SetAndAdd(PartSetter::Run(_))
            | crate::registry::Access::Set(PartSetter::Run(_)) => {
                if run.is_none()
                    && let Some(read) = current_run.take()
                {
                    *run = Some(read()?);
                }
                let target = run
                    .as_mut()
                    .ok_or_else(|| error(Code::InternalError, "no run target"))?;
                stage(assignment, operation, |setter, value| match setter {
                    PartSetter::Run(setter) => setter.apply_run(target, value),
                    _ => Ok(()),
                })?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn is_formatting(assignment: &Assignment<'_, '_>) -> bool {
    matches!(
        assignment.entry.access,
        crate::registry::Access::SetAndAdd(PartSetter::Paragraph(_) | PartSetter::Run(_))
            | crate::registry::Access::Set(PartSetter::Paragraph(_) | PartSetter::Run(_))
    )
}

/// Set a level's indentation (`indent` is the leading edge, `hanging` the
/// hanging indent), in twips.
fn set_indentation(paragraph: &mut direct::Paragraph, key: LevelKey, value: Option<i64>) {
    let mut indentation = paragraph.indentation.take().unwrap_or_default();
    let twips = value.map(direct::Twips);
    match key {
        LevelKey::Indent => {
            indentation.left = twips;
            indentation.start = None;
        },
        _ => indentation.hanging = twips,
    }
    paragraph.indentation = (indentation != direct::Indentation::default()).then_some(indentation);
}

impl Docx {
    /// `set` on part elements: `Err(editor)` hands the editor back when the
    /// targets are not part elements.
    pub(in super::super) fn try_set_parts(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty()
            || nodes
                .iter()
                .any(|node| !super::is_part_kind(node.element_kind()))
        {
            return Ok(Err(editor));
        }
        // Admit before staging.
        for node in nodes {
            match node.element_kind() {
                Kind::Style => drop(admit(&STYLE_REGISTRY, Operation::Set, props)?),
                Kind::AbstractNum => drop(admit(&ABSTRACT_NUM_REGISTRY, Operation::Set, props)?),
                Kind::Num => drop(admit(&NUM_REGISTRY, Operation::Set, props)?),
                Kind::Level => drop(admit(&LEVEL_REGISTRY, Operation::Set, props)?),
                Kind::Header | Kind::Footer => drop(admit(&STORY_REGISTRY, Operation::Set, props)?),
                Kind::Footnote | Kind::Endnote => {
                    drop(admit(&NOTE_REGISTRY, Operation::Set, props)?)
                },
                _ => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "the styles, numbering and notes parts have no settable properties",
                    ));
                },
            }
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("part edits need the owning package editor"));
        };
        for node in nodes {
            let Value::Part(value) = node.value else {
                return Err(error(
                    Code::InternalError,
                    "a part node holds no part value",
                ));
            };
            self.set_part(eager, value, props)?;
        }
        Ok(Ok(edit.commit(self)?))
    }

    fn set_part(
        &self,
        eager: &mut litchi_docx::document::Edit,
        value: PartValue,
        props: &Props<'_>,
    ) -> Result<()> {
        match value {
            PartValue::Style(index) => self.set_style(eager, index as usize, props),
            PartValue::AbstractNum(index) => {
                let id = super::numbering(self)?
                    .and_then(|catalog| catalog.abstract_nums().get(index as usize))
                    .map(litchi_docx::numbering::catalog::AbstractNum::id)
                    .ok_or_else(|| error(Code::PathNotFound, "the numbering definition is gone"))?;
                let mut change = AbstractNumChange::default();
                for assignment in admit(&ABSTRACT_NUM_REGISTRY, Operation::Set, props)? {
                    let (setter, text) = stage(&assignment, Operation::Set, |setter, value| {
                        Ok((setter, text_of(value)))
                    })?;
                    change = match setter {
                        PartSetter::AbstractNum(AbstractKey::Type) => {
                            change.with_multi_level_type(text)
                        },
                        PartSetter::AbstractNum(AbstractKey::Name) => change.with_name(text),
                        PartSetter::AbstractNum(AbstractKey::StyleLink) => {
                            change.with_style_link(text)
                        },
                        PartSetter::AbstractNum(AbstractKey::NumStyleLink) => {
                            change.with_num_style_link(text)
                        },
                        _ => change,
                    };
                }
                eager.set_abstract_num(id, &change).map_err(Error::from)?;
                Ok(())
            },
            PartValue::Num(index) => {
                let id = super::numbering(self)?
                    .and_then(|catalog| catalog.nums().get(index as usize))
                    .map(litchi_docx::numbering::catalog::Num::id)
                    .ok_or_else(|| error(Code::PathNotFound, "the numbering instance is gone"))?;
                for assignment in admit(&NUM_REGISTRY, Operation::Set, props)? {
                    let target = stage(&assignment, Operation::Set, |_, value| Ok(int_of(value)))?;
                    let Some(target) = target.and_then(|value| u32::try_from(value).ok()) else {
                        return Err(error(
                            Code::InvalidValue,
                            "a numbering definition identifier is required",
                        )
                        .with_property("abstractNumId"));
                    };
                    let known = super::numbering(self)?
                        .is_some_and(|catalog| catalog.abstract_num(target).is_some());
                    if !known {
                        return Err(error(
                            Code::InvalidValue,
                            "no numbering definition has this identifier",
                        )
                        .with_property("abstractNumId"));
                    }
                    eager.set_num_definition(id, target).map_err(Error::from)?;
                }
                Ok(())
            },
            PartValue::Level(definition, index) => self.set_level(eager, definition, index, props),
            PartValue::Story(index) => self.set_story(eager, index as usize, props),
            PartValue::Note(kind, index) => {
                let id = super::notes(self, kind)?
                    .get(index as usize)
                    .map(litchi_docx::footnote::Note::id)
                    .ok_or_else(|| error(Code::PathNotFound, "the note is gone"))?;
                let current = super::text(self, value)?;
                for assignment in admit(&NOTE_REGISTRY, Operation::Set, props)? {
                    let text = stage(&assignment, Operation::Set, |_, value| {
                        Ok(text_of(value).unwrap_or_default())
                    })?;
                    // Writing a note's own text back is an exact no-op.
                    if current.as_deref() != Some(text.as_str()) {
                        eager.set_note_text(kind, id, &text).map_err(Error::from)?;
                    }
                }
                Ok(())
            },
            PartValue::Styles | PartValue::Numbering | PartValue::Notes(_) => Err(error(
                Code::UnsupportedCapability,
                "the styles, numbering and notes parts have no settable properties",
            )),
        }
    }

    fn set_style(
        &self,
        eager: &mut litchi_docx::document::Edit,
        index: usize,
        props: &Props<'_>,
    ) -> Result<()> {
        let catalog = super::styles(self)?
            .ok_or_else(|| error(Code::PathNotFound, "the styles part is gone"))?;
        let entry = catalog
            .entries()
            .get(index)
            .ok_or_else(|| error(Code::PathNotFound, "the style is gone"))?;
        let id = entry.id().to_owned();
        let assignments = admit(&STYLE_REGISTRY, Operation::Set, props)?;
        let change = self.style_change(
            &assignments,
            Operation::Set,
            || catalog.paragraph_properties(entry).map_err(Error::from),
            || catalog.run_properties(entry).map_err(Error::from),
        )?;
        eager.set_style(&id, &change).map_err(Error::from)?;
        Ok(())
    }

    /// The style change the assignments describe.
    fn style_change(
        &self,
        assignments: &[Assignment<'_, '_>],
        operation: Operation,
        current_paragraph: impl FnOnce() -> Result<direct::Paragraph>,
        current_run: impl FnOnce() -> Result<direct::Run>,
    ) -> Result<StyleChange> {
        let mut change = StyleChange::default();
        for assignment in assignments
            .iter()
            .filter(|assignment| !is_formatting(assignment))
        {
            let next = stage(assignment, operation, |setter, value| {
                let current = std::mem::take(&mut change);
                Ok(match setter {
                    PartSetter::Style(StyleKey::Name) => match text_of(value) {
                        Some(name) => current.with_name(name),
                        None => {
                            return Err(error(
                                Code::InvalidValue,
                                "a style name cannot be cleared",
                            ));
                        },
                    },
                    PartSetter::Style(StyleKey::BasedOn) => current.with_based_on(
                        text_of(value)
                            .map(|key| resolve_style_id(self, &key))
                            .transpose()?,
                    ),
                    PartSetter::Style(StyleKey::Next) => current.with_next(
                        text_of(value)
                            .map(|key| resolve_style_id(self, &key))
                            .transpose()?,
                    ),
                    PartSetter::Style(StyleKey::Linked) => current.with_link(
                        text_of(value)
                            .map(|key| resolve_style_id(self, &key))
                            .transpose()?,
                    ),
                    PartSetter::Style(StyleKey::Flag(flag)) => {
                        current.with_flag(flag, bool_of(value))
                    },
                    PartSetter::Style(StyleKey::UiPriority) => {
                        current.with_ui_priority(int_of(value))
                    },
                    _ => current,
                })
            })?;
            change = next;
        }
        let formatting: Vec<&Assignment<'_, '_>> = assignments
            .iter()
            .filter(|assignment| is_formatting(assignment))
            .collect();
        let mut paragraph = None;
        let mut run = None;
        apply_formatting(
            &formatting,
            operation,
            &mut paragraph,
            &mut run,
            current_paragraph,
            current_run,
        )?;
        if let Some(paragraph) = paragraph {
            change = change.with_paragraph(paragraph);
        }
        if let Some(run) = run {
            change = change.with_run(run);
        }
        Ok(change)
    }

    fn set_level(
        &self,
        eager: &mut litchi_docx::document::Edit,
        definition: u32,
        index: u32,
        props: &Props<'_>,
    ) -> Result<()> {
        let catalog = super::numbering(self)?
            .ok_or_else(|| error(Code::PathNotFound, "the numbering part is gone"))?;
        let abstract_num = catalog
            .abstract_nums()
            .get(definition as usize)
            .ok_or_else(|| error(Code::PathNotFound, "the numbering definition is gone"))?;
        let level = abstract_num
            .levels()
            .get(index as usize)
            .ok_or_else(|| error(Code::PathNotFound, "the numbering level is gone"))?;
        let assignments = admit(&LEVEL_REGISTRY, Operation::Set, props)?;
        let mut change = LevelChange::default();
        let mut paragraph: Option<direct::Paragraph> = None;
        for assignment in assignments
            .iter()
            .filter(|assignment| !is_formatting(assignment))
        {
            stage(assignment, Operation::Set, |setter, value| {
                let current = std::mem::take(&mut change);
                change = match setter {
                    PartSetter::Level(LevelKey::Start) => current.with_start(int_of(value)),
                    PartSetter::Level(LevelKey::Format) => current.with_format(text_of(value)),
                    PartSetter::Level(LevelKey::Text) => current.with_text(text_of(value)),
                    PartSetter::Level(LevelKey::Justification) => {
                        current.with_justification(text_of(value))
                    },
                    PartSetter::Level(LevelKey::Suffix) => current.with_suffix(text_of(value)),
                    PartSetter::Level(LevelKey::Restart) => current.with_restart(int_of(value)),
                    PartSetter::Level(LevelKey::Legal) => current.with_legal(bool_of(value)),
                    PartSetter::Level(key @ (LevelKey::Indent | LevelKey::Hanging)) => {
                        if paragraph.is_none() {
                            paragraph = Some(
                                catalog
                                    .level_paragraph_properties(level)
                                    .map_err(Error::from)?,
                            );
                        }
                        if let Some(paragraph) = paragraph.as_mut() {
                            set_indentation(paragraph, key, int_of(value));
                        }
                        current
                    },
                    _ => current,
                };
                Ok(())
            })?;
        }
        let formatting: Vec<&Assignment<'_, '_>> = assignments
            .iter()
            .filter(|assignment| is_formatting(assignment))
            .collect();
        let mut run = None;
        let mut unused_paragraph = None;
        apply_formatting(
            &formatting,
            Operation::Set,
            &mut unused_paragraph,
            &mut run,
            || {
                catalog
                    .level_paragraph_properties(level)
                    .map_err(Error::from)
            },
            || catalog.level_run_properties(level).map_err(Error::from),
        )?;
        if let Some(paragraph) = paragraph {
            change = change.with_paragraph(paragraph);
        }
        if let Some(run) = run {
            change = change.with_run(run);
        }
        eager
            .set_level(abstract_num.id(), level.ilvl(), &change)
            .map_err(Error::from)?;
        Ok(())
    }

    fn set_story(
        &self,
        eager: &mut litchi_docx::document::Edit,
        index: usize,
        props: &Props<'_>,
    ) -> Result<()> {
        let stories = super::stories(self)?;
        let story = stories
            .get(index)
            .ok_or_else(|| error(Code::PathNotFound, "the header or footer is gone"))?;
        let name = story.part().name().to_owned();
        let assignments = admit(&STORY_REGISTRY, Operation::Set, props)?;
        for assignment in assignments
            .iter()
            .filter(|assignment| !is_formatting(assignment))
        {
            stage(assignment, Operation::Set, |setter, value| {
                match setter {
                    PartSetter::Text => {
                        // The text a story reads includes its field results;
                        // writing that text back is an exact no-op.
                        let text = text_of(value).unwrap_or_default();
                        if super::text(self, PartValue::Story(ordinal(index)))?.as_deref()
                            != Some(text.as_str())
                        {
                            eager.set_story_text(&name, &text).map_err(Error::from)?;
                        }
                    },
                    PartSetter::StoryKind => {
                        let kind = text_of(value)
                            .as_deref()
                            .and_then(StoryKind::from_xml)
                            .ok_or_else(|| {
                                error(
                                    Code::InvalidValue,
                                    "a header type is default, first or even",
                                )
                            })?;
                        eager.set_header_footer_kind(&name, kind).map_err(|leaf| {
                            let mapped = Error::from(leaf);
                            if mapped.code() == Code::InvalidValue {
                                error(
                                    Code::DuplicateName,
                                    "the section already has a header or footer of that type",
                                )
                            } else {
                                mapped
                            }
                        })?;
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        let formatting: Vec<&Assignment<'_, '_>> = assignments
            .iter()
            .filter(|assignment| is_formatting(assignment))
            .collect();
        if formatting.is_empty() {
            return Ok(());
        }
        // Check every value once on empty targets, so the closures below
        // cannot fail.
        let mut paragraph_probe = None;
        let mut run_probe = None;
        apply_formatting(
            &formatting,
            Operation::Set,
            &mut paragraph_probe,
            &mut run_probe,
            || Ok(direct::Paragraph::default()),
            || Ok(direct::Run::default()),
        )?;
        let paragraphs: Vec<&Assignment<'_, '_>> = formatting
            .iter()
            .copied()
            .filter(|assignment| {
                matches!(
                    assignment.entry.access,
                    crate::registry::Access::SetAndAdd(PartSetter::Paragraph(_))
                )
            })
            .collect();
        let runs: Vec<&Assignment<'_, '_>> = formatting
            .iter()
            .copied()
            .filter(|assignment| {
                matches!(
                    assignment.entry.access,
                    crate::registry::Access::SetAndAdd(PartSetter::Run(_))
                )
            })
            .collect();
        let paragraph_change = |paragraph: &mut direct::Paragraph| {
            for assignment in &paragraphs {
                let _ = stage(assignment, Operation::Set, |setter, value| match setter {
                    PartSetter::Paragraph(setter) => setter.apply_paragraph(paragraph, value),
                    _ => Ok(()),
                });
            }
        };
        let run_change = |run: &mut direct::Run| {
            for assignment in &runs {
                let _ = stage(assignment, Operation::Set, |setter, value| match setter {
                    PartSetter::Run(setter) => setter.apply_run(run, value),
                    _ => Ok(()),
                });
            }
        };
        let paragraph_ref: &dyn Fn(&mut direct::Paragraph) = &paragraph_change;
        let run_ref: &dyn Fn(&mut direct::Run) = &run_change;
        eager
            .format_story(
                &name,
                (!paragraphs.is_empty()).then_some(paragraph_ref),
                (!runs.is_empty()).then_some(run_ref),
            )
            .map_err(Error::from)?;
        Ok(())
    }

    /// `add` of a part element under `parent`.
    pub(in super::super) fn add_part(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        kind: NewPart,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        match kind {
            NewPart::Style => self.add_style(editor, parent, node),
            NewPart::AbstractNum => self.add_abstract_num(editor, parent, node),
            NewPart::Num => self.add_num(editor, parent, node),
            NewPart::Level => self.add_level(editor, parent, node),
            NewPart::Header | NewPart::Footer => self.add_story(editor, parent, kind, node),
            NewPart::Footnote | NewPart::Endnote => self.add_note(editor, parent, kind, node),
        }
    }

    fn add_style(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        if !matches!(parent.element_kind(), Kind::Styles | Kind::Document) {
            return Err(error(
                Code::UnsupportedCapability,
                "a style is added under /styles",
            ));
        }
        let assignments = admit(&STYLE_REGISTRY, Operation::Add, node.props())?;
        let mut id = None;
        let mut name = None;
        let mut kind = StyleType::Paragraph;
        let mut custom = None;
        for assignment in &assignments {
            stage(assignment, Operation::Add, |setter, value| {
                match setter {
                    PartSetter::Style(StyleKey::Id) => id = text_of(value),
                    PartSetter::Style(StyleKey::Name) => name = text_of(value),
                    PartSetter::Style(StyleKey::Type) => {
                        kind = text_of(value)
                            .as_deref()
                            .and_then(StyleType::from_xml)
                            .unwrap_or(StyleType::Paragraph);
                    },
                    PartSetter::Style(StyleKey::CustomStyle) => custom = bool_of(value),
                    _ => {},
                }
                Ok(())
            })?;
        }
        let taken = |candidate: &str| -> Result<bool> {
            Ok(super::styles(self)?.is_some_and(|catalog| catalog.get(candidate).is_some()))
        };
        let explicit = id.is_some();
        let mut style_id = id
            .or_else(|| name.clone())
            .unwrap_or_else(|| "CustomStyle".to_owned());
        if style_id.is_empty() {
            return Err(
                error(Code::InvalidValue, "a style identifier cannot be empty").with_property("id"),
            );
        }
        if taken(&style_id)? {
            if explicit {
                return Err(
                    error(Code::DuplicateName, "a style with this identifier exists")
                        .with_property("id"),
                );
            }
            let base = style_id.clone();
            let mut suffix = 2usize;
            while taken(&style_id)? {
                style_id = format!("{base}{suffix}");
                suffix += 1;
            }
        }
        let others: Vec<Assignment<'_, '_>> = assignments
            .into_iter()
            .filter(|assignment| {
                !matches!(
                    assignment.entry.access,
                    crate::registry::Access::Add(PartSetter::Style(
                        StyleKey::Id | StyleKey::Name | StyleKey::Type | StyleKey::CustomStyle
                    )) | crate::registry::Access::SetAndAdd(PartSetter::Style(StyleKey::Name))
                )
            })
            .collect();
        let mut change = self.style_change(
            &others,
            Operation::Add,
            || Ok(direct::Paragraph::default()),
            || Ok(direct::Run::default()),
        )?;
        change = change.with_name(name.unwrap_or_else(|| style_id.clone()));
        let custom = custom.unwrap_or(!BUILT_IN_STYLES.contains(&style_id.as_str()));
        let style = NewStyle::new(style_id.clone(), kind)
            .with_custom(custom)
            .with_change(change);
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("styles need the owning package editor"));
        };
        eager.add_style(&style).map_err(Error::from)?;
        let commit = edit.commit(self)?;
        Ok((
            commit,
            Path::from_steps([PathStep::new("styles"), PathStep::new(style_id)]),
        ))
    }

    fn numbering_parent(parent: &Node<'_>) -> Result<()> {
        if matches!(parent.element_kind(), Kind::Numbering | Kind::Document) {
            Ok(())
        } else {
            Err(error(
                Code::UnsupportedCapability,
                "numbering definitions and instances are added under /numbering",
            ))
        }
    }

    fn add_abstract_num(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        Self::numbering_parent(parent)?;
        let assignments = admit(&ABSTRACT_NUM_REGISTRY, Operation::Add, node.props())?;
        let mut id = None;
        let mut multi_level_type = None;
        let mut name = None;
        let mut style_link = None;
        let mut num_style_link = None;
        let mut format = None;
        let mut text = None;
        let mut start = None;
        let mut indent = None;
        for assignment in &assignments {
            stage(assignment, Operation::Add, |setter, value| {
                match setter {
                    PartSetter::AbstractNum(AbstractKey::Id) => id = int_of(value),
                    PartSetter::AbstractNum(AbstractKey::Type) => multi_level_type = text_of(value),
                    PartSetter::AbstractNum(AbstractKey::Name) => name = text_of(value),
                    PartSetter::AbstractNum(AbstractKey::StyleLink) => style_link = text_of(value),
                    PartSetter::AbstractNum(AbstractKey::NumStyleLink) => {
                        num_style_link = text_of(value)
                    },
                    PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Format)) => {
                        format = text_of(value)
                    },
                    PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Text)) => {
                        text = text_of(value)
                    },
                    PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Start)) => {
                        start = int_of(value)
                    },
                    PartSetter::AbstractNum(AbstractKey::Seed(LevelKey::Indent)) => {
                        indent = int_of(value)
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        let bare = Bare {
            format: format.is_some(),
            text: text.is_some(),
            start: start.is_some(),
            indent: indent.is_some(),
        };
        let definition = new_abstract_num(
            id,
            multi_level_type,
            name,
            style_link,
            num_style_link,
            format,
            text,
            start,
            indent,
        )?;
        let definition = seed_levels(definition, &assignments, bare)?;
        if let Some(wanted) = id
            && super::numbering(self)?.is_some_and(|catalog| {
                u32::try_from(wanted).is_ok_and(|wanted| catalog.abstract_num(wanted).is_some())
            })
        {
            return Err(error(
                Code::DuplicateName,
                "a numbering definition with this identifier exists",
            )
            .with_property("id"));
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("numbering needs the owning package editor"));
        };
        // The schema puts every definition before the instances, and the
        // leaf appends a new one after the last (record 0325): its ordinal
        // is the handle's definition count, read without reparsing the
        // committed part.
        let ordinal = super::numbering(self)?.map_or(0, |catalog| catalog.abstract_nums().len());
        let _ = eager.add_abstract_num(&definition).map_err(Error::from)?;
        let commit = edit.commit(self)?;
        Ok((commit, numbering_path("abstractNum", ordinal)))
    }

    fn add_num(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        Self::numbering_parent(parent)?;
        let assignments = admit(&NUM_REGISTRY, Operation::Add, node.props())?;
        let mut id = None;
        let mut abstract_num_id = None;
        let mut overrides: Vec<(u8, i64)> = Vec::new();
        let mut continues = false;
        // OfficeCLI's "Mode A" (record 0325): the definition the instance
        // creates when it names none.
        let mut template = false;
        let mut multi_level_type = None;
        let mut name = None;
        let mut style_link = None;
        let mut num_style_link = None;
        let mut format = None;
        let mut text = None;
        let mut indent = None;
        for assignment in &assignments {
            stage(assignment, Operation::Add, |setter, value| {
                match setter {
                    PartSetter::Num(NumKey::Id) => id = int_of(value),
                    PartSetter::Num(NumKey::AbstractNumId) => abstract_num_id = int_of(value),
                    PartSetter::Num(NumKey::Template(key)) => {
                        template = true;
                        match key {
                            AbstractKey::Type => multi_level_type = text_of(value),
                            AbstractKey::Name => name = text_of(value),
                            AbstractKey::StyleLink => style_link = text_of(value),
                            AbstractKey::NumStyleLink => num_style_link = text_of(value),
                            AbstractKey::Seed(LevelKey::Format) => format = text_of(value),
                            AbstractKey::Seed(LevelKey::Text) => text = text_of(value),
                            AbstractKey::Seed(LevelKey::Indent) => indent = int_of(value),
                            _ => {},
                        }
                    },
                    PartSetter::LevelSeed { .. } => template = true,
                    PartSetter::Num(NumKey::StartOverride(level)) => {
                        if let Some(start) = int_of(value) {
                            overrides.retain(|(other, _)| *other != level);
                            overrides.push((level, start));
                        }
                    },
                    PartSetter::Num(NumKey::Continue) => {
                        continues = bool_of(value).unwrap_or(false)
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        if template {
            if abstract_num_id.is_some() {
                return Err(error(
                    Code::InvalidArgument,
                    "a numbering instance names its definition or describes a new one, not both",
                )
                .with_property("abstractNumId"));
            }
            let start = overrides
                .iter()
                .find(|(level, _)| *level == 0)
                .map(|(_, start)| *start);
            let bare = Bare {
                format: format.is_some(),
                text: text.is_some(),
                start: start.is_some(),
                indent: indent.is_some(),
            };
            let definition = new_abstract_num(
                None,
                multi_level_type,
                name,
                style_link,
                num_style_link,
                format,
                text,
                start,
                indent,
            )?;
            let definition = seed_levels(definition, &assignments, bare)?;
            return self.add_templated_num(editor, definition, id, overrides, continues);
        }
        let Some(abstract_num_id) = abstract_num_id.and_then(|value| u32::try_from(value).ok())
        else {
            return Err(error(
                Code::MissingProperty,
                "a new numbering instance needs the definition it instantiates",
            )
            .with_property("abstractNumId")
            .with_suggestion("or describe a new definition with format, text or level<N>.* keys"));
        };
        let catalog = super::numbering(self)?;
        let Some(definition) = catalog.and_then(|catalog| catalog.abstract_num(abstract_num_id))
        else {
            return Err(error(
                Code::InvalidValue,
                "no numbering definition has this identifier",
            )
            .with_property("abstractNumId"));
        };
        if let Some(wanted) = id
            && catalog.is_some_and(|catalog| {
                u32::try_from(wanted).is_ok_and(|wanted| catalog.num(wanted).is_some())
            })
        {
            return Err(error(
                Code::DuplicateName,
                "a numbering instance with this identifier exists",
            )
            .with_property("id"));
        }
        // OfficeCLI restarts a new instance at its definition's first start
        // unless `continue` asks for Word's continued counting.
        if !continues && !overrides.iter().any(|(level, _)| *level == 0) {
            let start = definition
                .level(0)
                .and_then(|level| level.start())
                .and_then(|start| start.trim().parse::<i64>().ok())
                .unwrap_or(1);
            overrides.push((0, start));
        }
        let mut num = NewNum::new(abstract_num_id);
        if let Some(id) = id.and_then(|id| u32::try_from(id).ok()) {
            num = num.with_id(id);
        }
        for (level, start) in overrides {
            num = num.with_start_override(level, start);
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("numbering needs the owning package editor"));
        };
        let added = eager.add_num(&num).map_err(Error::from)?;
        let commit = edit.commit(self)?;
        let ordinal = super::numbering(commit.file())?
            .and_then(|catalog| catalog.nums().iter().position(|entry| entry.id() == added))
            .ok_or_else(|| refusal("the new numbering instance is not in the numbering part"))?;
        Ok((commit, numbering_path("num", ordinal)))
    }

    /// Add a new definition and an instance of it in one edit (OfficeCLI's
    /// `add num` "Mode A", record 0325).
    fn add_templated_num(
        &self,
        editor: DocxEditor,
        definition: NewAbstractNum,
        id: Option<i64>,
        mut overrides: Vec<(u8, i64)>,
        continues: bool,
    ) -> Result<(Commit, Path<'static>)> {
        if let Some(wanted) = id
            && super::numbering(self)?.is_some_and(|catalog| {
                u32::try_from(wanted).is_ok_and(|wanted| catalog.num(wanted).is_some())
            })
        {
            return Err(error(
                Code::DuplicateName,
                "a numbering instance with this identifier exists",
            )
            .with_property("id"));
        }
        if !continues && !overrides.iter().any(|(level, _)| *level == 0) {
            let start = definition.level(0).map_or(1, NewLevel::start);
            overrides.push((0, start));
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("numbering needs the owning package editor"));
        };
        let abstract_num_id = eager.add_abstract_num(&definition).map_err(Error::from)?;
        let mut num = NewNum::new(abstract_num_id);
        if let Some(id) = id.and_then(|id| u32::try_from(id).ok()) {
            num = num.with_id(id);
        }
        for (level, start) in overrides {
            num = num.with_start_override(level, start);
        }
        let added = eager.add_num(&num).map_err(Error::from)?;
        let commit = edit.commit(self)?;
        let ordinal = super::numbering(commit.file())?
            .and_then(|catalog| catalog.nums().iter().position(|entry| entry.id() == added))
            .ok_or_else(|| refusal("the new numbering instance is not in the numbering part"))?;
        Ok((commit, numbering_path("num", ordinal)))
    }

    fn add_level(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let Value::Part(PartValue::AbstractNum(definition)) = parent.value else {
            return Err(error(
                Code::UnsupportedCapability,
                "a numbering level is added under a numbering definition",
            ));
        };
        let catalog = super::numbering(self)?
            .ok_or_else(|| error(Code::PathNotFound, "the numbering part is gone"))?;
        let abstract_num = catalog
            .abstract_nums()
            .get(definition as usize)
            .ok_or_else(|| error(Code::PathNotFound, "the numbering definition is gone"))?;
        let assignments = admit(&LEVEL_REGISTRY, Operation::Add, node.props())?;
        let mut ilvl = None;
        let mut format = None;
        let mut text = None;
        let mut start = None;
        let mut justification = None;
        let mut suffix = None;
        let mut paragraph = direct::Paragraph::default();
        for assignment in assignments
            .iter()
            .filter(|assignment| !is_formatting(assignment))
        {
            stage(assignment, Operation::Add, |setter, value| {
                match setter {
                    PartSetter::Level(LevelKey::Ilvl) => ilvl = int_of(value),
                    PartSetter::Level(LevelKey::Format) => format = text_of(value),
                    PartSetter::Level(LevelKey::Text) => text = text_of(value),
                    PartSetter::Level(LevelKey::Start) => start = int_of(value),
                    PartSetter::Level(LevelKey::Justification) => justification = text_of(value),
                    PartSetter::Level(LevelKey::Suffix) => suffix = text_of(value),
                    PartSetter::Level(key @ (LevelKey::Indent | LevelKey::Hanging)) => {
                        set_indentation(&mut paragraph, key, int_of(value));
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        let Some(ilvl) = ilvl
            .and_then(|ilvl| u8::try_from(ilvl).ok())
            .filter(|ilvl| *ilvl <= 8)
        else {
            return Err(error(
                Code::MissingProperty,
                "a new numbering level needs its ilvl (0-8)",
            )
            .with_property("ilvl"));
        };
        let formatting: Vec<&Assignment<'_, '_>> = assignments
            .iter()
            .filter(|assignment| is_formatting(assignment))
            .collect();
        let mut run = None;
        let mut unused = None;
        apply_formatting(
            &formatting,
            Operation::Add,
            &mut unused,
            &mut run,
            || Ok(direct::Paragraph::default()),
            || Ok(direct::Run::default()),
        )?;
        let format = format.unwrap_or_else(|| "decimal".to_owned());
        let text = text.unwrap_or_else(|| format!("%{}.", ilvl + 1));
        let mut level = NewLevel::new(ilvl, format, text).with_paragraph(paragraph);
        if let Some(start) = start {
            level = level.with_start(start);
        }
        if let Some(justification) = justification {
            level = level.with_justification(justification);
        }
        if let Some(suffix) = suffix {
            level = level.with_suffix(suffix);
        }
        if let Some(run) = run {
            level = level.with_run(run);
        }
        let id = abstract_num.id();
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("numbering needs the owning package editor"));
        };
        eager.add_level(id, &level).map_err(Error::from)?;
        let commit = edit.commit(self)?;
        let (definition_ordinal, level_ordinal) = super::numbering(commit.file())?
            .and_then(|catalog| {
                let at = catalog
                    .abstract_nums()
                    .iter()
                    .position(|entry| entry.id() == id)?;
                let level = catalog.abstract_nums()[at]
                    .levels()
                    .iter()
                    .position(|level| level.ilvl() == ilvl)?;
                Some((at, level))
            })
            .ok_or_else(|| refusal("the new numbering level is not in its definition"))?;
        Ok((
            commit,
            Path::from_steps([
                PathStep::new("numbering"),
                PathStep::new("abstractNum")
                    .with(Predicate::Index(Position::new(definition_ordinal))),
                PathStep::new("level").with(Predicate::Index(Position::new(level_ordinal))),
            ]),
        ))
    }

    fn add_story(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        kind: NewPart,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        if !matches!(parent.element_kind(), Kind::Document | Kind::Body) {
            return Err(error(
                Code::UnsupportedCapability,
                "headers and footers are added under the document root",
            ));
        }
        let role = if kind == NewPart::Header {
            Role::Header
        } else {
            Role::Footer
        };
        let assignments = admit(&STORY_REGISTRY, Operation::Add, node.props())?;
        let mut text = String::new();
        let mut story_kind = StoryKind::Primary;
        let mut field = None;
        for assignment in assignments
            .iter()
            .filter(|assignment| !is_formatting(assignment))
        {
            stage(assignment, Operation::Add, |setter, value| {
                match setter {
                    PartSetter::Text => text = text_of(value).unwrap_or_default(),
                    PartSetter::StoryKind => {
                        story_kind = text_of(value)
                            .as_deref()
                            .and_then(StoryKind::from_xml)
                            .ok_or_else(|| {
                                error(
                                    Code::InvalidValue,
                                    "a header type is default, first or even",
                                )
                            })?;
                    },
                    PartSetter::StoryField => {
                        field = text_of(value).map(|name| match name.as_str() {
                            "page" => " PAGE ".to_owned(),
                            "numpages" => " NUMPAGES ".to_owned(),
                            "date" => " DATE \\@ \"yyyy-MM-dd\" ".to_owned(),
                            other => format!(" {} ", other.to_ascii_uppercase()),
                        });
                    },
                    _ => {},
                }
                Ok(())
            })?;
        }
        let formatting: Vec<&Assignment<'_, '_>> = assignments
            .iter()
            .filter(|assignment| is_formatting(assignment))
            .collect();
        let mut paragraph = None;
        let mut run = None;
        apply_formatting(
            &formatting,
            Operation::Add,
            &mut paragraph,
            &mut run,
            || Ok(direct::Paragraph::default()),
            || Ok(direct::Run::default()),
        )?;
        let mut story = NewStory::new(text);
        if let Some(paragraph) = paragraph {
            story = story.with_paragraph(paragraph);
        }
        if let Some(run) = run {
            story = story.with_run(run);
        }
        if let Some(field) = field {
            story = story.with_field(field);
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal(
                "headers and footers need the owning package editor",
            ));
        };
        let name = eager
            .add_header_footer(role, story_kind, &story)
            .map_err(|leaf| {
                let mapped = Error::from(leaf);
                if mapped.code() == Code::InvalidValue {
                    error(
                        Code::DuplicateName,
                        "the section already has a header or footer of that type",
                    )
                } else {
                    mapped
                }
            })?;
        let commit = edit.commit(self)?;
        let ordinal = super::stories(commit.file())?
            .iter()
            .filter(|story| story.part().role() == role)
            .position(|story| story.part().name() == name)
            .ok_or_else(|| refusal("the new header or footer is not related from the document"))?;
        let step = if role == Role::Header {
            "header"
        } else {
            "footer"
        };
        Ok((
            commit,
            Path::from_steps([PathStep::new(step).with(Predicate::Index(Position::new(ordinal)))]),
        ))
    }

    fn add_note(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        kind: NewPart,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let note_kind = if kind == NewPart::Footnote {
            NoteKind::Footnote
        } else {
            NoteKind::Endnote
        };
        let assignments = admit(&NOTE_REGISTRY, Operation::Add, node.props())?;
        let mut text = None;
        for assignment in &assignments {
            stage(assignment, Operation::Add, |setter, value| {
                if setter == PartSetter::Text {
                    text = Some(text_of(value).unwrap_or_default());
                }
                Ok(())
            })?;
        }
        let Some(text) = text else {
            return Err(
                error(Code::MissingProperty, "a new note needs its text").with_property("text")
            );
        };
        // A paragraph parent gets the reference after its last child, or
        // before a direct run a placement names; a run parent gets it right
        // after that run.
        let path = element_path(parent)?;
        let (paragraph, at) = match parent.element_kind() {
            Kind::Paragraph => {
                let Value::Paragraph { paragraph, .. } = &parent.value else {
                    return Err(refusal("the note parent holds no paragraph"));
                };
                let locations = run_locations(self, paragraph)?;
                let at = super::super::nested::sibling_position(
                    self,
                    parent,
                    Kind::Run,
                    locations.len(),
                    node.placement(),
                )?;
                let direct = match at {
                    None => None,
                    Some(index) if index >= locations.len() => None,
                    Some(index) => match locations.get(index).map(Vec::as_slice) {
                        Some([ElementStep::Run(run)]) => Some(*run),
                        _ => {
                            return Err(refusal(
                                "a note reference goes beside a direct run of the paragraph",
                            ));
                        },
                    },
                };
                (path, direct)
            },
            Kind::Run => match path.split_last() {
                Some((paragraph, ElementStep::Run(run))) => {
                    if !matches!(node.placement(), None | Some(Placement::End)) {
                        return Err(refusal("a note under a run goes right after the run"));
                    }
                    (paragraph, Some(Position::new(run.get() + 1)))
                },
                _ => return Err(refusal("a note is anchored on a direct run of a paragraph")),
            },
            _ => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "footnotes and endnotes are added under a paragraph or a run",
                ));
            },
        };
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("notes need the owning package editor"));
        };
        let id = eager
            .add_note_at(note_kind, &paragraph, at, &NewNote::new(text))
            .map_err(Error::from)?;
        let commit = edit.commit(self)?;
        let ordinal = super::notes(commit.file(), note_kind)?
            .iter()
            .position(|note| note.id() == id)
            .ok_or_else(|| refusal("the new note is not in its notes part"))?;
        let (container, step) = if note_kind == NoteKind::Footnote {
            ("footnotes", "footnote")
        } else {
            ("endnotes", "endnote")
        };
        Ok((
            commit,
            Path::from_steps([
                PathStep::new(container),
                PathStep::new(step).with(Predicate::Index(Position::new(ordinal))),
            ]),
        ))
    }

    /// `remove` of part elements.
    pub(in super::super) fn remove_parts(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
    ) -> Result<Commit> {
        enum Removal {
            Style(String),
            Num(u32),
            AbstractNum(u32),
            Level(u32, u8),
            Story(String),
            Note(NoteKind, u32),
        }
        let mut removals = Vec::with_capacity(nodes.len());
        for node in nodes {
            let Value::Part(value) = node.value else {
                return Err(error(
                    Code::InternalError,
                    "a part node holds no part value",
                ));
            };
            removals.push(match value {
                PartValue::Style(index) => Removal::Style(
                    super::styles(self)?
                        .and_then(|catalog| catalog.entries().get(index as usize))
                        .map(|entry| entry.id().to_owned())
                        .ok_or_else(|| error(Code::PathNotFound, "the style is gone"))?,
                ),
                PartValue::AbstractNum(index) => Removal::AbstractNum(
                    super::numbering(self)?
                        .and_then(|catalog| catalog.abstract_nums().get(index as usize))
                        .map(litchi_docx::numbering::catalog::AbstractNum::id)
                        .ok_or_else(|| {
                            error(Code::PathNotFound, "the numbering definition is gone")
                        })?,
                ),
                PartValue::Num(index) => Removal::Num(
                    super::numbering(self)?
                        .and_then(|catalog| catalog.nums().get(index as usize))
                        .map(litchi_docx::numbering::catalog::Num::id)
                        .ok_or_else(|| {
                            error(Code::PathNotFound, "the numbering instance is gone")
                        })?,
                ),
                PartValue::Level(definition, index) => {
                    let catalog = super::numbering(self)?
                        .ok_or_else(|| error(Code::PathNotFound, "the numbering part is gone"))?;
                    let definition = catalog
                        .abstract_nums()
                        .get(definition as usize)
                        .ok_or_else(|| {
                            error(Code::PathNotFound, "the numbering definition is gone")
                        })?;
                    let level = definition
                        .levels()
                        .get(index as usize)
                        .ok_or_else(|| error(Code::PathNotFound, "the numbering level is gone"))?;
                    Removal::Level(definition.id(), level.ilvl())
                },
                PartValue::Story(index) => Removal::Story(
                    super::stories(self)?
                        .get(index as usize)
                        .map(|story| story.part().name().to_owned())
                        .ok_or_else(|| error(Code::PathNotFound, "the header or footer is gone"))?,
                ),
                PartValue::Note(kind, index) => Removal::Note(
                    kind,
                    super::notes(self, kind)?
                        .get(index as usize)
                        .map(litchi_docx::footnote::Note::id)
                        .ok_or_else(|| error(Code::PathNotFound, "the note is gone"))?,
                ),
                PartValue::Styles | PartValue::Numbering | PartValue::Notes(_) => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "the styles, numbering and notes parts cannot be removed",
                    ));
                },
            });
        }
        // Instances go before the definitions they use.
        removals.sort_by_key(|removal| match removal {
            Removal::Num(_) => 0,
            Removal::Level(..) => 1,
            Removal::AbstractNum(_) => 2,
            _ => 3,
        });
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("part edits need the owning package editor"));
        };
        for removal in &removals {
            match removal {
                Removal::Style(id) => eager.remove_style(id).map(drop),
                Removal::Num(id) => eager.remove_num(*id).map(drop),
                Removal::AbstractNum(id) => eager.remove_abstract_num(*id).map(drop),
                Removal::Level(id, ilvl) => eager.remove_level(*id, *ilvl).map(drop),
                Removal::Story(name) => eager.remove_header_footer(name).map(drop),
                Removal::Note(kind, id) => eager.remove_note(*kind, *id).map(drop),
            }
            .map_err(Error::from)?;
        }
        edit.commit(self)
    }
}

fn ordinal(index: usize) -> u32 {
    u32::try_from(index).unwrap_or(u32::MAX)
}

fn numbering_path(step: &'static str, ordinal: usize) -> Path<'static> {
    Path::from_steps([
        PathStep::new("numbering"),
        PathStep::new(step).with(Predicate::Index(Position::new(ordinal))),
    ])
}

/// A new definition from OfficeCLI's `add abstractNum` keys: nine template
/// levels, an explicit `format` on every level, and `text`, `start` and
/// `indent` seeding level 0.
#[allow(clippy::too_many_arguments, reason = "one argument per OfficeCLI key")]
/// Which bare level-0 keys a new definition was given: they win over the
/// same `level0.*` key (OfficeCLI).
#[derive(Clone, Copy)]
struct Bare {
    format: bool,
    text: bool,
    start: bool,
    indent: bool,
}

/// The `level<N>.*` values of one level.
#[derive(Default)]
struct Seeds<'a, 'p, 'v> {
    format: Option<String>,
    text: Option<String>,
    start: Option<i64>,
    indent: Option<i64>,
    hanging: Option<i64>,
    justification: Option<String>,
    suffix: Option<String>,
    run: Vec<&'a Assignment<'p, 'v>>,
}

/// Apply the `level<N>.*` keys among `assignments` to the template levels
/// of `definition` (OfficeCLI's `BuildAbstractNumElement`, record 0325):
/// a level keeps its template value for every key it is not given, a bare
/// level-0 key wins over its `level0.*` spelling, a `bullet` format without
/// text takes the level's bullet character, and `indent` and `hanging`
/// default to `(N + 1) × 720` and 360 twips.
fn seed_levels(
    mut definition: NewAbstractNum,
    assignments: &[Assignment<'_, '_>],
    bare: Bare,
) -> Result<NewAbstractNum> {
    let mut levels: [Option<Seeds<'_, '_, '_>>; 9] = Default::default();
    for assignment in assignments {
        let crate::registry::Access::Add(PartSetter::LevelSeed { level, key }) =
            assignment.entry.access
        else {
            continue;
        };
        let seeds = levels
            .get_mut(usize::from(level))
            .ok_or_else(|| error(Code::InternalError, "a numbering level is above 8"))?
            .get_or_insert_with(Seeds::default);
        match key {
            SeedKey::Run(_) => seeds.run.push(assignment),
            SeedKey::Level(key) => stage(assignment, Operation::Add, |_, value| {
                match key {
                    LevelKey::Format => seeds.format = text_of(value),
                    LevelKey::Text => seeds.text = text_of(value),
                    LevelKey::Start => seeds.start = int_of(value),
                    LevelKey::Indent => seeds.indent = int_of(value),
                    LevelKey::Hanging => seeds.hanging = int_of(value),
                    LevelKey::Justification => seeds.justification = text_of(value),
                    LevelKey::Suffix => seeds.suffix = text_of(value),
                    _ => {},
                }
                Ok(())
            })?,
        }
    }
    for (ilvl, seeds) in (0_u8..).zip(levels) {
        let Some(seeds) = seeds else {
            continue;
        };
        let Some(mut level) = definition.level(ilvl).cloned() else {
            continue;
        };
        let first = ilvl == 0;
        if let Some(format) = seeds.format.filter(|_| !(first && bare.format)) {
            let bullet = format.eq_ignore_ascii_case("bullet");
            level = level.with_format(format);
            if bullet && seeds.text.is_none() && !(first && bare.text) {
                level =
                    level.with_text(["\u{2022}", "\u{25E6}", "\u{25AA}"][usize::from(ilvl % 3)]);
            }
        }
        if let Some(text) = seeds.text.filter(|_| !(first && bare.text)) {
            level = level.with_text(text);
        }
        if let Some(start) = seeds.start.filter(|_| !(first && bare.start)) {
            level = level.with_start(start);
        }
        if let Some(justification) = seeds.justification {
            level = level.with_justification(justification);
        }
        if let Some(suffix) = seeds.suffix {
            level = level.with_suffix(suffix);
        }
        let indent = seeds.indent.filter(|_| !(first && bare.indent));
        if indent.is_some() || seeds.hanging.is_some() {
            let current = level.paragraph().indentation.as_ref().map(|indentation| {
                (
                    indentation.left.map(|twips| twips.0),
                    indentation.hanging.map(|twips| twips.0),
                )
            });
            let mut paragraph = direct::Paragraph::default();
            let mut indentation = direct::Indentation::default();
            indentation.left = Some(direct::Twips(indent.unwrap_or_else(|| {
                current
                    .and_then(|(left, _)| left)
                    .unwrap_or(720 * (i64::from(ilvl) + 1))
            })));
            indentation.hanging =
                Some(direct::Twips(seeds.hanging.unwrap_or_else(|| {
                    current.and_then(|(_, hanging)| hanging).unwrap_or(360)
                })));
            paragraph.indentation = Some(indentation);
            level = level.with_paragraph(paragraph);
        }
        if !seeds.run.is_empty() {
            let mut run = direct::Run::default();
            for assignment in seeds.run {
                stage(assignment, Operation::Add, |setter, value| match setter {
                    PartSetter::LevelSeed {
                        key: SeedKey::Run(setter),
                        ..
                    } => setter.apply_run(&mut run, value),
                    _ => Ok(()),
                })?;
            }
            level = level.with_run(run);
        }
        definition = definition.with_level(level);
    }
    Ok(definition)
}

fn new_abstract_num(
    id: Option<i64>,
    multi_level_type: Option<String>,
    name: Option<String>,
    style_link: Option<String>,
    num_style_link: Option<String>,
    format: Option<String>,
    text: Option<String>,
    start: Option<i64>,
    indent: Option<i64>,
) -> Result<NewAbstractNum> {
    let bullet = format.as_deref().is_some_and(|format| {
        ["bullet", "unordered", "ul"]
            .iter()
            .any(|known| format.eq_ignore_ascii_case(known))
    });
    let mut definition = NewAbstractNum::new(bullet);
    if let Some(id) = id {
        definition = definition
            .with_id(u32::try_from(id).map_err(|_| {
                error(Code::InvalidValue, "a numbering identifier is out of range")
            })?);
    }
    if let Some(value) = multi_level_type {
        let canonical = MULTI_LEVEL_TYPE.get(&value).ok_or_else(|| {
            error(Code::InvalidValue, "unknown multi-level type").with_property("type")
        })?;
        definition = definition.with_multi_level_type(canonical);
    }
    if let Some(name) = name {
        definition = definition.with_name(name);
    }
    if let Some(link) = style_link {
        definition = definition.with_style_link(link);
    }
    if let Some(link) = num_style_link {
        definition = definition.with_num_style_link(link);
    }
    if let Some(format) = format.filter(|_| !bullet) {
        for ilvl in 0..9u8 {
            if let Some(level) = definition.level(ilvl) {
                let level = level.clone().with_format(format.clone());
                definition = definition.with_level(level);
            }
        }
    }
    if text.is_some() || start.is_some() || indent.is_some() {
        if let Some(level) = definition.level(0) {
            let mut level = level.clone();
            if let Some(text) = text {
                level = level.with_text(text);
            }
            if let Some(start) = start {
                level = level.with_start(start);
            }
            if let Some(indent) = indent {
                let mut paragraph = direct::Paragraph::default();
                let mut indentation = direct::Indentation::default();
                indentation.left = Some(direct::Twips(indent));
                indentation.hanging = Some(direct::Twips(360));
                paragraph.indentation = Some(indentation);
                level = level.with_paragraph(paragraph);
            }
            definition = definition.with_level(level);
        }
    }
    Ok(definition)
}
