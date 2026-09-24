//! Scalar publication against retained Survey XML.

use super::*;
use litchi_opc::{OwnedAttributeUpdate, OwnedXmlPart};
use std::ops::Range;

mod properties;
mod questions;

#[derive(Debug, Clone)]
pub(super) struct QuestionOrigin {
    lineage: Arc<()>,
    index: usize,
}

pub(super) fn assign_question_origins(questions: &mut [Question]) {
    let lineage = Arc::new(());
    for (index, question) in questions.iter_mut().enumerate() {
        question.origin = Some(QuestionOrigin {
            lineage: lineage.clone(),
            index,
        });
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Owner {
    Survey,
    Questions,
    Question(usize),
    SurveyPr,
    TitlePr,
    DescriptionPr,
    QuestionsPr,
    QuestionPr(usize),
    QuestionExt(usize),
    Opaque,
}

struct Pending {
    tag: Range<usize>,
    name: &'static str,
    value: Option<Vec<u8>>,
}

struct Changes {
    edits: Vec<Pending>,
    encoded_bytes: usize,
    maximum: usize,
}

impl Changes {
    fn text(
        &mut self,
        tag: &Range<usize>,
        name: &'static str,
        before: Option<&str>,
        after: Option<&str>,
        xstring: bool,
    ) -> Result<()> {
        if before != after {
            self.push(tag, name, after, xstring)?;
        }
        Ok(())
    }

    fn number<T: PartialEq + ToString>(
        &mut self,
        tag: &Range<usize>,
        name: &'static str,
        before: Option<T>,
        after: Option<T>,
    ) -> Result<()> {
        if before != after {
            let value = after.map(|value| value.to_string());
            self.push(tag, name, value.as_deref(), false)?;
        }
        Ok(())
    }

    fn push(
        &mut self,
        tag: &Range<usize>,
        name: &'static str,
        value: Option<&str>,
        xstring: bool,
    ) -> Result<()> {
        if self.edits.len() >= 65_536 {
            return Err(invalid("too many Survey attribute edits"));
        }
        let value = value
            .map(|value| {
                let mut size = 0usize;
                if xstring {
                    add_xstring_size(&mut size, value)?;
                } else {
                    size = escaped_len(value)?;
                }
                self.encoded_bytes = self
                    .encoded_bytes
                    .checked_add(size)
                    .ok_or_else(|| invalid("Survey attribute size overflow"))?;
                if self.encoded_bytes > self.maximum {
                    return Err(invalid("Survey attribute output exceeds limit"));
                }
                let mut bytes = Vec::new();
                bytes
                    .try_reserve_exact(size)
                    .map_err(|_| invalid("Survey attribute allocation failed"))?;
                if xstring {
                    crate::source_attributes::append_escaped_xstring(&mut bytes, value);
                } else {
                    escape_attribute(&mut bytes, value);
                }
                Ok::<_, Error>(bytes)
            })
            .transpose()?;
        self.edits
            .try_reserve(1)
            .map_err(|_| invalid("Survey edit allocation failed"))?;
        self.edits.push(Pending {
            tag: tag.clone(),
            name,
            value,
        });
        Ok(())
    }

    fn properties(
        &mut self,
        tag: &Range<usize>,
        before: &ElementProperties,
        after: &ElementProperties,
    ) -> Result<()> {
        self.text(
            tag,
            "cssClass",
            before.css_class.as_deref(),
            after.css_class.as_deref(),
            true,
        )?;
        self.number(tag, "bottom", before.bottom, after.bottom)?;
        self.number(tag, "top", before.top, after.top)?;
        self.number(tag, "left", before.left, after.left)?;
        self.number(tag, "right", before.right, after.right)?;
        self.number(tag, "width", before.width, after.width)?;
        self.number(tag, "height", before.height, after.height)?;
        self.text(
            tag,
            "position",
            before.position.map(Position::as_str),
            after.position.map(Position::as_str),
            false,
        )
    }

    fn question(&mut self, tag: &Range<usize>, before: &Question, after: &Question) -> Result<()> {
        self.number(
            tag,
            "binding",
            Some(before.binding.get()),
            Some(after.binding.get()),
        )?;
        self.text(
            tag,
            "text",
            before.text.as_deref(),
            after.text.as_deref(),
            true,
        )?;
        self.text(
            tag,
            "type",
            before.kind.map(QuestionType::as_str),
            after.kind.map(QuestionType::as_str),
            false,
        )?;
        self.text(
            tag,
            "format",
            before.format.map(QuestionFormat::as_str),
            after.format.map(QuestionFormat::as_str),
            false,
        )?;
        self.text(
            tag,
            "helpText",
            before.help_text.as_deref(),
            after.help_text.as_deref(),
            true,
        )?;
        self.number(tag, "required", Some(before.required), Some(after.required))?;
        self.text(
            tag,
            "defaultValue",
            before.default_value.as_deref(),
            after.default_value.as_deref(),
            true,
        )?;
        self.number(
            tag,
            "decimalPlaces",
            before.decimal_places,
            after.decimal_places,
        )?;
        self.text(
            tag,
            "rowSource",
            before.row_source.as_deref(),
            after.row_source.as_deref(),
            true,
        )
    }
}

fn same_properties(before: Option<&ElementProperties>, after: Option<&ElementProperties>) -> bool {
    match (before, after) {
        (None, None) => true,
        (Some(before), Some(after)) => {
            before.extension_xml == after.extension_xml
                && before.unknown_attributes == after.unknown_attributes
                && before.namespace_declarations == after.namespace_declarations
                && before.comments == after.comments
                && before.prefix == after.prefix
                && before.default_namespace == after.default_namespace
                && before.comment_slots == after.comment_slots
        },
        _ => false,
    }
}

fn same_structure(before: &Survey, after: &Survey) -> bool {
    before.extension_xml == after.extension_xml
        && before.unknown_attributes == after.unknown_attributes
        && before.namespace_declarations == after.namespace_declarations
        && before.comments == after.comments
        && before.comment_slots == after.comment_slots
        && before.default_namespace == after.default_namespace
        && before.root_prefix == after.root_prefix
        && before.questions.comments == after.questions.comments
        && (before.questions.comment_slots == after.questions.comment_slots
            || (before.questions.source_question_count != after.questions.source_question_count
                && normalized_questions_comment_slots(&before.questions)
                    == normalized_questions_comment_slots(&after.questions)))
        && same_default_namespace(
            before.questions.default_namespace.as_deref(),
            after.questions.default_namespace.as_deref(),
        )
        && before.questions.prefix == after.questions.prefix
        && before.questions.namespace_declarations == after.questions.namespace_declarations
        && before.questions.unknown_attributes == after.questions.unknown_attributes
        && before.questions.values.len() == after.questions.values.len()
        && before
            .questions
            .values
            .iter()
            .zip(after.questions.values.iter())
            .all(|(before, after)| {
                before.column_name == after.column_name
                    && before.extension_xml == after.extension_xml
                    && before.unknown_attributes == after.unknown_attributes
                    && before.namespace_declarations == after.namespace_declarations
                    && before.comments == after.comments
                    && before.comment_slots == after.comment_slots
                    && same_default_namespace(
                        before.default_namespace.as_deref(),
                        after.default_namespace.as_deref(),
                    )
                    && before.prefix == after.prefix
            })
}

fn classify(
    parent: Option<Owner>,
    namespace: &ResolveResult<'_>,
    name: &[u8],
    question: &mut usize,
) -> Owner {
    if !is_survey_namespace(namespace) {
        return Owner::Opaque;
    }
    match (parent, name) {
        (None, b"survey") => Owner::Survey,
        (Some(Owner::Survey), b"surveyPr") => Owner::SurveyPr,
        (Some(Owner::Survey), b"titlePr") => Owner::TitlePr,
        (Some(Owner::Survey), b"descriptionPr") => Owner::DescriptionPr,
        (Some(Owner::Survey), b"questions") => Owner::Questions,
        (Some(Owner::Questions), b"questionsPr") => Owner::QuestionsPr,
        (Some(Owner::Questions), b"question") => {
            let index = *question;
            *question += 1;
            Owner::Question(index)
        },
        (Some(Owner::Question(index)), b"questionPr") => Owner::QuestionPr(index),
        (Some(Owner::Question(index)), b"extLst") => Owner::QuestionExt(index),
        _ => Owner::Opaque,
    }
}

/// Scalar, property, and same-parent question sequence publication.
pub(super) fn rewrite(
    source: &OwnedXmlPart,
    before: &Survey,
    after: &Survey,
    limits: &Limits,
) -> Result<Option<OwnedXmlPart>> {
    if let Some(updated) = questions::rewrite(source, before, after, limits)? {
        let projected = parse_with_limits(updated.bytes(), limits)?;
        if projected == *after {
            return Ok(Some(updated));
        }
        return rewrite_aligned(&updated, &projected, after, limits)?
            .map(Some)
            .ok_or(Error::Unsupported {
                feature: "Survey question context replacement",
            });
    }
    rewrite_aligned(source, before, after, limits)
}

fn rewrite_aligned(
    source: &OwnedXmlPart,
    before: &Survey,
    after: &Survey,
    limits: &Limits,
) -> Result<Option<OwnedXmlPart>> {
    if !same_structure(before, after) {
        return Ok(None);
    }
    let mut changes = Changes {
        edits: Vec::new(),
        encoded_bytes: 0,
        maximum: limits.max_part_bytes(),
    };
    let mut reader = NsReader::from_reader(source.bytes());
    let mut stack = Vec::new();
    let mut question = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_error)?;
        let end = reader.buffer_position() as usize;
        let paired = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                let owner = classify(
                    stack.last().copied(),
                    &namespace,
                    element.local_name().as_ref(),
                    &mut question,
                );
                let tag = start..end;
                let properties = match owner {
                    Owner::Survey => {
                        changes.number(&tag, "id", Some(before.id.get()), Some(after.id.get()))?;
                        changes.text(
                            &tag,
                            "guid",
                            Some(before.guid.as_str()),
                            Some(after.guid.as_str()),
                            false,
                        )?;
                        changes.text(
                            &tag,
                            "title",
                            before.title.as_deref(),
                            after.title.as_deref(),
                            true,
                        )?;
                        changes.text(
                            &tag,
                            "description",
                            before.description.as_deref(),
                            after.description.as_deref(),
                            true,
                        )?;
                        None
                    },
                    Owner::Question(index) => {
                        let before = before
                            .questions
                            .values
                            .get(index)
                            .ok_or_else(|| invalid("Survey source question mismatch"))?;
                        let after = after
                            .questions
                            .values
                            .get(index)
                            .ok_or_else(|| invalid("Survey staged question mismatch"))?;
                        changes.question(&tag, before, after)?;
                        None
                    },
                    Owner::SurveyPr => {
                        Some((before.properties.as_ref(), after.properties.as_ref()))
                    },
                    Owner::TitlePr => Some((
                        before.title_properties.as_ref(),
                        after.title_properties.as_ref(),
                    )),
                    Owner::DescriptionPr => Some((
                        before.description_properties.as_ref(),
                        after.description_properties.as_ref(),
                    )),
                    Owner::QuestionsPr => Some((
                        before.questions.properties.as_ref(),
                        after.questions.properties.as_ref(),
                    )),
                    Owner::QuestionPr(index) => Some((
                        before
                            .questions
                            .values
                            .get(index)
                            .and_then(|question| question.properties.as_ref()),
                        after
                            .questions
                            .values
                            .get(index)
                            .and_then(|question| question.properties.as_ref()),
                    )),
                    Owner::Questions | Owner::QuestionExt(_) | Owner::Opaque => None,
                };
                if let Some((Some(before), Some(after))) = properties {
                    if same_properties(Some(before), Some(after)) {
                        changes.properties(&tag, before, after)?;
                    }
                }
                if paired {
                    stack.push(owner);
                }
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| invalid("unbalanced Survey source"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if question != before.questions.values.len() {
        return Err(invalid("Survey source question count mismatch"));
    }
    let mut updates = Vec::new();
    updates
        .try_reserve_exact(changes.edits.len())
        .map_err(|_| invalid("Survey edit allocation failed"))?;
    updates.extend(changes.edits.iter().map(|edit| OwnedAttributeUpdate {
        start_tag: edit.tag.clone(),
        name: edit.name,
        value: edit.value.as_deref(),
    }));
    let updated = source.update_attributes(&updates, limits.max_part_bytes())?;
    Ok(Some(properties::rewrite(&updated, before, after, limits)?))
}
