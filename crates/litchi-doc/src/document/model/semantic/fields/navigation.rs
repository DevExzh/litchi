use super::super::prelude::*;

impl Document {
    /// Get the innermost stored field containing a story-relative character
    /// position.
    ///
    /// MS-DOC section 2.8.25 records field begin, separator, and end
    /// characters in a story-local `Plcfld`. The complete field range includes
    /// both its begin and end positions. When fields are nested, the field
    /// with the greatest stored nesting depth is returned. The result exposes
    /// only stored instruction and cached-result text; it never evaluates or
    /// refreshes a field.
    pub fn field_text_at_position(&self, story: FieldStory, cp: u32) -> Result<Option<FieldText>> {
        let Some(fields_table) = &self.fields_table else {
            return Ok(None);
        };
        let Some(field) = fields_table
            .fields(story)
            .iter()
            .filter(|field| field.start_cp <= cp && cp <= field.end_cp)
            .max_by_key(|field| (field.nesting_depth, field.start_cp))
        else {
            return Ok(None);
        };

        let mut text = self.field_text(field)?;
        let instruction_end = field.separator_cp.unwrap_or(field.end_cp);
        let nested = fields_table
            .fields(story)
            .iter()
            .filter(|nested| {
                nested.start_cp > field.start_cp
                    && nested.end_cp < instruction_end
                    && nested.start_cp < nested.end_cp
            })
            .map(|nested| (nested.start_cp, nested.end_cp));
        text.instruction =
            without_nested_fields(&text.instruction, field.start_cp.saturating_add(1), nested);
        Ok(Some(text))
    }
}

fn without_nested_fields(
    text: &str,
    base_cp: u32,
    nested: impl Iterator<Item = (u32, u32)>,
) -> String {
    let ranges = nested
        .filter_map(|(start, end)| {
            Some((
                start.checked_sub(base_cp)?,
                end.checked_sub(base_cp)?.checked_add(1)?,
            ))
        })
        .collect::<Vec<_>>();
    if ranges.is_empty() {
        return text.to_owned();
    }

    let mut output = String::with_capacity(text.len());
    let mut cp = 0u32;
    for character in text.chars() {
        let end = cp.saturating_add(character.len_utf16() as u32);
        if !ranges
            .iter()
            .any(|(start, end_range)| cp >= *start && cp < *end_range)
        {
            output.push(character);
        }
        cp = end;
    }
    output
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::Package;
    use crate::writer::Writer;
    use std::io::Cursor;

    #[test]
    fn field_text_at_position_selects_the_innermost_nested_field() {
        let mut writer = Writer::new();
        writer
            .add_paragraph(concat!(
                "\u{0013}IF 1 = ",
                "\u{0013}HYPERLINK \"target\"",
                "\u{0014}display",
                "\u{0015}",
                "\u{0014}yes",
                "\u{0015}",
            ))
            .unwrap();

        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        let mut package = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
        let document = package.document().unwrap();
        let fields = document.fields().unwrap();
        assert_eq!(fields.len(), 2);

        let outer = &fields[0].field;
        let inner = &fields[1].field;
        let selected = document
            .field_text_at_position(FieldStory::Main, inner.start_cp + 1)
            .unwrap()
            .unwrap();
        assert_eq!(selected.field, *inner);
        assert_eq!(selected.instruction.trim(), r#"HYPERLINK "target""#);

        let selected = document
            .field_text_at_position(FieldStory::Main, outer.start_cp + 1)
            .unwrap()
            .unwrap();
        assert_eq!(selected.field, *outer);
        assert_eq!(selected.instruction.trim(), "IF 1 =");

        let selected = document
            .field_text_at_position(FieldStory::Main, outer.end_cp)
            .unwrap()
            .unwrap();
        assert_eq!(selected.field, *outer);

        assert!(
            document
                .field_text_at_position(FieldStory::Main, outer.end_cp + 1)
                .unwrap()
                .is_none()
        );
    }
}
