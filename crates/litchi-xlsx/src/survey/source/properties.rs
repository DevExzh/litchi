//! Property-element presence edits with schema-ordered insertion.

use super::*;
use litchi_opc::{OwnedElementEdit, OwnedElementUpdate};

#[derive(Clone, Copy)]
enum Placement {
    Before,
    Append,
    Replace,
}

struct Pending {
    tag: Range<usize>,
    fragment: Option<Vec<u8>>,
    placement: Placement,
}

struct Plan {
    tags: HashMap<Owner, Range<usize>>,
    edits: Vec<Pending>,
    authored: usize,
    maximum: usize,
}

pub(super) fn authored_size(name: &str, properties: &ElementProperties) -> Result<usize> {
    let mut size = name.len() + NAMESPACE.len() + 12;
    if let Some(prefix) = properties.prefix.as_deref() {
        let prefix_markup = prefix
            .len()
            .checked_mul(2)
            .and_then(|size| size.checked_add(2))
            .ok_or_else(|| invalid("Survey property size overflow"))?;
        add_size(&mut size, prefix_markup)?;
    }
    if let Some(css) = properties.css_class.as_deref() {
        add_size(&mut size, "cssClass".len() + 4)?;
        add_xstring_size(&mut size, css)?;
    }
    for (name, value) in [
        ("bottom", properties.bottom.map(i64::from)),
        ("top", properties.top.map(i64::from)),
        ("left", properties.left.map(i64::from)),
        ("right", properties.right.map(i64::from)),
        ("width", properties.width.map(i64::from)),
        ("height", properties.height.map(i64::from)),
    ] {
        if let Some(value) = value {
            let digits = value
                .unsigned_abs()
                .checked_ilog10()
                .map_or(1, |log| log as usize + 1);
            add_size(&mut size, name.len() + 4 + digits + usize::from(value < 0))?;
        }
    }
    if let Some(position) = properties.position {
        add_size(&mut size, "position".len() + 4 + position.as_str().len())?;
    }
    Ok(size)
}

impl Plan {
    fn property(
        &mut self,
        owner: Owner,
        parent: Owner,
        name: &'static str,
        before: Option<&ElementProperties>,
        after: Option<&ElementProperties>,
        following: &[Owner],
    ) -> Result<()> {
        match (before, after) {
            (Some(_), None) => {
                let tag = self
                    .tags
                    .get(&owner)
                    .ok_or_else(|| invalid("Survey property source is missing"))?
                    .clone();
                self.push(Pending {
                    tag,
                    fragment: None,
                    placement: Placement::Replace,
                })?;
            },
            (old, Some(properties)) if !same_properties(old, Some(properties)) => {
                // Public constructors create scalar properties. Transplanting
                // retained opaque context needs source identity and bindings.
                if properties.extension_xml.is_some()
                    || !properties.unknown_attributes.is_empty()
                    || !properties.namespace_declarations.is_empty()
                    || !properties.comments.is_empty()
                    || properties.default_namespace.is_some()
                {
                    return Err(Error::Unsupported {
                        feature: "Survey property source transplantation",
                    });
                }
                let charge = authored_size(name, properties)?;
                self.authored = self
                    .authored
                    .checked_add(charge)
                    .ok_or_else(|| invalid("Survey property size overflow"))?;
                if self.authored > self.maximum {
                    return Err(invalid("Survey properties exceed output limit"));
                }
                let mut fragment = Vec::new();
                fragment
                    .try_reserve_exact(charge)
                    .map_err(|_| invalid("Survey property allocation failed"))?;
                write_properties_with_namespace(&mut fragment, name, properties, true)?;
                let next = following.iter().find_map(|owner| self.tags.get(owner));
                let (tag, placement) = if old.is_some() {
                    (self.tags.get(&owner), Placement::Replace)
                } else if let Some(next) = next {
                    (Some(next), Placement::Before)
                } else {
                    (self.tags.get(&parent), Placement::Append)
                };
                let tag = tag
                    .ok_or_else(|| invalid("Survey property parent is missing"))?
                    .clone();
                self.push(Pending {
                    tag,
                    fragment: Some(fragment),
                    placement,
                })?;
            },
            _ => {},
        }
        Ok(())
    }

    fn push(&mut self, edit: Pending) -> Result<()> {
        if self.edits.len() >= 65_536 {
            return Err(invalid("too many Survey property edits"));
        }
        self.edits
            .try_reserve(1)
            .map_err(|_| invalid("Survey property edit allocation failed"))?;
        self.edits.push(edit);
        Ok(())
    }
}

pub(super) fn rewrite(
    source: &OwnedXmlPart,
    before: &Survey,
    after: &Survey,
    limits: &Limits,
) -> Result<OwnedXmlPart> {
    let changed = !same_properties(before.properties.as_ref(), after.properties.as_ref())
        || !same_properties(
            before.title_properties.as_ref(),
            after.title_properties.as_ref(),
        )
        || !same_properties(
            before.description_properties.as_ref(),
            after.description_properties.as_ref(),
        )
        || !same_properties(
            before.questions.properties.as_ref(),
            after.questions.properties.as_ref(),
        )
        || before
            .questions
            .values
            .iter()
            .zip(after.questions.values.iter())
            .any(|(before, after)| {
                !same_properties(before.properties.as_ref(), after.properties.as_ref())
            });
    if !changed {
        return Ok(source.clone());
    }
    let mut plan = Plan {
        tags: HashMap::new(),
        edits: Vec::new(),
        authored: 0,
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
                if owner != Owner::Opaque {
                    plan.tags
                        .try_reserve(1)
                        .map_err(|_| invalid("Survey source tag allocation failed"))?;
                    if plan.tags.insert(owner, start..end).is_some() {
                        return Err(invalid("duplicate Survey source owner"));
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
    plan.property(
        Owner::SurveyPr,
        Owner::Survey,
        "surveyPr",
        before.properties.as_ref(),
        after.properties.as_ref(),
        &[Owner::TitlePr, Owner::DescriptionPr, Owner::Questions],
    )?;
    plan.property(
        Owner::TitlePr,
        Owner::Survey,
        "titlePr",
        before.title_properties.as_ref(),
        after.title_properties.as_ref(),
        &[Owner::DescriptionPr, Owner::Questions],
    )?;
    plan.property(
        Owner::DescriptionPr,
        Owner::Survey,
        "descriptionPr",
        before.description_properties.as_ref(),
        after.description_properties.as_ref(),
        &[Owner::Questions],
    )?;
    plan.property(
        Owner::QuestionsPr,
        Owner::Questions,
        "questionsPr",
        before.questions.properties.as_ref(),
        after.questions.properties.as_ref(),
        &[Owner::Question(0)],
    )?;
    for (index, (before, after)) in before
        .questions
        .values
        .iter()
        .zip(after.questions.values.iter())
        .enumerate()
    {
        plan.property(
            Owner::QuestionPr(index),
            Owner::Question(index),
            "questionPr",
            before.properties.as_ref(),
            after.properties.as_ref(),
            &[Owner::QuestionExt(index)],
        )?;
    }
    // Stable sorting retains schema order when several new properties share
    // one following source sibling.
    plan.edits.sort_by_key(|edit| edit.tag.start);
    let mut updates = Vec::new();
    updates
        .try_reserve_exact(plan.edits.len())
        .map_err(|_| invalid("Survey property update allocation failed"))?;
    updates.extend(plan.edits.iter().map(|edit| OwnedElementUpdate {
        start_tag: edit.tag.clone(),
        edit: match (&edit.fragment, edit.placement) {
            (Some(fragment), Placement::Append) => OwnedElementEdit::AppendChild(fragment),
            (Some(fragment), Placement::Replace) => OwnedElementEdit::Replace(fragment),
            (Some(fragment), Placement::Before) => OwnedElementEdit::InsertBefore(fragment),
            (None, _) => OwnedElementEdit::Remove,
        },
    }));
    Ok(source.update_elements(&updates, limits.max_part_bytes())?)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn authored_property_size_accounts_for_extremes_and_xstring_expansion() {
        let mut properties = ElementProperties::new();
        properties.set_css_class(Some("_x0041_\0\t&\"😀".into()));
        properties.set_bottom(Some(i32::MIN));
        properties.set_top(Some(i32::MAX));
        properties.set_left(Some(-1));
        properties.set_right(Some(0));
        properties.set_width(Some(u32::MAX));
        properties.set_height(Some(0));
        properties.set_position(Some(Position::Absolute));
        let mut prefixed = ElementProperties::new();
        prefixed.prefix = Some("q".into());
        for properties in [&ElementProperties::new(), &properties, &prefixed] {
            for name in [
                "surveyPr",
                "titlePr",
                "descriptionPr",
                "questionsPr",
                "questionPr",
            ] {
                let size = authored_size(name, properties).unwrap();
                let mut bytes = Vec::with_capacity(size);
                write_properties_with_namespace(&mut bytes, name, properties, true).unwrap();
                assert_eq!(bytes.len(), size);
            }
        }
    }
}
