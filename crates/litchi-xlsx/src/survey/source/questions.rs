//! Stable question source identity and same-parent sequence publication.

use super::*;
use litchi_core::xml::ReaderOrigin;
use litchi_opc::OwnedChildElement;

fn original_index(before: &Survey, question: &Question) -> Option<usize> {
    let origin = question.origin.as_ref()?;
    let original = before.questions.values.get(origin.index)?.origin.as_ref()?;
    (original.index == origin.index && Arc::ptr_eq(&original.lineage, &origin.lineage))
        .then_some(origin.index)
}

fn has_opaque_properties(properties: &ElementProperties) -> bool {
    properties.extension_xml.is_some()
        || !properties.unknown_attributes.is_empty()
        || !properties.namespace_declarations.is_empty()
        || !properties.comments.is_empty()
        || properties.default_namespace.is_some()
        || properties.prefix.is_some()
}

fn authored_size(question: &Question) -> Result<usize> {
    let digits = |value: u32| value.checked_ilog10().map_or(1, |log| log as usize + 1);
    let mut size = "<question".len()
        + NAMESPACE.len()
        + 9
        + " binding=\"\"".len()
        + digits(question.binding.get());
    for (name, value) in [
        ("text", question.text.as_deref()),
        ("helpText", question.help_text.as_deref()),
        ("defaultValue", question.default_value.as_deref()),
        ("rowSource", question.row_source.as_deref()),
    ] {
        if let Some(value) = value {
            add_size(&mut size, name.len() + 4)?;
            add_xstring_size(&mut size, value)?;
        }
    }
    for (name, value) in [
        ("type", question.kind.map(QuestionType::as_str)),
        ("format", question.format.map(QuestionFormat::as_str)),
    ] {
        if let Some(value) = value {
            add_size(&mut size, name.len() + 4 + value.len())?;
        }
    }
    if question.required {
        add_size(&mut size, " required=\"true\"".len())?;
    }
    if let Some(value) = question.decimal_places {
        add_size(
            &mut size,
            " decimalPlaces=\"\"".len() + digits(u32::from(value)),
        )?;
    }
    if let Some(properties) = question.properties.as_ref() {
        add_size(
            &mut size,
            properties::authored_size("questionPr", properties)?,
        )?;
        add_size(&mut size, "></question>".len())?;
    } else {
        add_size(&mut size, 2)?;
    }
    Ok(size)
}

fn authored_question(question: &Question, remaining: usize) -> Result<Vec<u8>> {
    if question.extension_xml.is_some()
        || !question.unknown_attributes.is_empty()
        || !question.namespace_declarations.is_empty()
        || !question.comments.is_empty()
        || question.default_namespace.is_some()
        || question.prefix.is_some()
        || question
            .properties
            .as_ref()
            .is_some_and(has_opaque_properties)
    {
        return Err(Error::Unsupported {
            feature: "Survey question source transplantation",
        });
    }
    let size = authored_size(question)?;
    if size > remaining {
        return Err(invalid("authored Survey question exceeds output limit"));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(size)
        .map_err(|_| invalid("Survey question allocation failed"))?;
    write_question_with_namespace(&mut bytes, question, true)?;
    Ok(bytes)
}

pub(super) fn rewrite(
    source: &OwnedXmlPart,
    before: &Survey,
    after: &Survey,
    limits: &Limits,
) -> Result<Option<OwnedXmlPart>> {
    if before.questions.values.len() == after.questions.values.len()
        && after
            .questions
            .values
            .iter()
            .enumerate()
            .all(|(index, question)| original_index(before, question) == Some(index))
    {
        return Ok(None);
    }
    let mut reader = NsReader::from_reader(source.bytes());
    let origin = ReaderOrigin::of(source.bytes());
    let mut stack = Vec::new();
    let mut count = 0usize;
    let mut parent = None;
    let mut tags = Vec::new();
    tags.try_reserve_exact(before.questions.values.len())
        .map_err(|_| invalid("Survey question tag allocation failed"))?;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Survey XML position exceeds usize"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Survey XML position exceeds usize"))?;
        let paired = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                let owner = classify(
                    stack.last().copied(),
                    &namespace,
                    element.local_name().as_ref(),
                    &mut count,
                );
                match owner {
                    Owner::Questions => {
                        if parent.replace(start..end).is_some() {
                            return Err(invalid("duplicate Survey questions source"));
                        }
                    },
                    Owner::Question(_) => tags.push(start..end),
                    _ => {},
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
    if tags.len() != before.questions.values.len() {
        return Err(invalid("Survey source question count mismatch"));
    }
    let parent = parent.ok_or_else(|| invalid("Survey question parent is missing"))?;
    let mut authored = Vec::new();
    authored
        .try_reserve_exact(after.questions.values.len())
        .map_err(|_| invalid("Survey authored question allocation failed"))?;
    let mut remaining = limits.max_part_bytes();
    for question in &after.questions.values {
        if original_index(before, question).is_some() {
            authored.push(None);
        } else {
            let bytes = authored_question(question, remaining)?;
            remaining -= bytes.len();
            authored.push(Some(bytes));
        }
    }
    let mut children = Vec::new();
    children
        .try_reserve_exact(after.questions.values.len())
        .map_err(|_| invalid("Survey question sequence allocation failed"))?;
    for (question, authored) in after.questions.values.iter().zip(&authored) {
        children.push(match original_index(before, question) {
            Some(index) => OwnedChildElement::Retained(index),
            None => OwnedChildElement::Authored(
                authored
                    .as_deref()
                    .ok_or_else(|| invalid("missing authored Survey question"))?,
            ),
        });
    }
    Ok(Some(source.replace_child_sequence(
        parent,
        &tags,
        &children,
        limits.max_part_bytes(),
    )?))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn authored_question_precharges_exact_size_and_rejects_one_over_limit() {
        let mut question = Question::new(Binding::new(u32::MAX));
        question.set_text(Some("_x0041_\0\t&\"😀".into()));
        question.set_help_text(Some("help".into()));
        question.set_default_value(Some("value".into()));
        question.set_row_source(Some("a;b\"c".into())).unwrap();
        question.set_question_type(Some(QuestionType::Number));
        question.set_format(Some(QuestionFormat::Fixed));
        question.set_required(true);
        question.set_decimal_places(Some(15));
        let mut properties = ElementProperties::new();
        properties.set_css_class(Some("x&\0".into()));
        properties.set_bottom(Some(i32::MIN));
        properties.set_width(Some(u32::MAX));
        question.set_properties(Some(properties));
        for question in [&Question::new(Binding::new(0)), &question] {
            let size = authored_size(question).unwrap();
            assert_eq!(authored_question(question, size).unwrap().len(), size);
            assert!(authored_question(question, size - 1).is_err());
        }
    }
}
