use super::NameError;

pub(crate) fn validate_ncname(value: &str) -> Result<(), NameError> {
    if is_ncname(value) {
        Ok(())
    } else {
        Err(NameError::InvalidNcName(value.to_owned()))
    }
}

/// Return whether value follows the XML 1.0 Fifth Edition `NCName` grammar.
#[must_use]
pub fn is_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_character)
}

/// Return whether value follows the XML 1.0 Fifth Edition Name grammar.
#[must_use]
pub fn is_xml_name(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_name_start) && characters.all(is_name_character)
}

/// Return whether value follows the XML Schema `QName` lexical grammar.
#[must_use]
pub fn is_qualified_name(value: &str) -> bool {
    let mut components = value.split(':');
    let first = components.next().unwrap_or_default();
    let second = components.next();
    components.next().is_none() && is_ncname(first) && second.is_none_or(is_ncname)
}

fn is_ncname_start(character: char) -> bool {
    character != ':' && is_name_start(character)
}

fn is_ncname_character(character: char) -> bool {
    character != ':' && is_name_character(character)
}

fn is_name_start(character: char) -> bool {
    matches!(
        character,
        ':' | 'A'..='Z' | '_' | 'a'..='z'
            | '\u{C0}'..='\u{D6}'
            | '\u{D8}'..='\u{F6}'
            | '\u{F8}'..='\u{2FF}'
            | '\u{370}'..='\u{37D}'
            | '\u{37F}'..='\u{1FFF}'
            | '\u{200C}'..='\u{200D}'
            | '\u{2070}'..='\u{218F}'
            | '\u{2C00}'..='\u{2FEF}'
            | '\u{3001}'..='\u{D7FF}'
            | '\u{F900}'..='\u{FDCF}'
            | '\u{FDF0}'..='\u{FFFD}'
            | '\u{10000}'..='\u{EFFFF}'
    )
}

fn is_name_character(character: char) -> bool {
    is_name_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{B7}' | '\u{300}'..='\u{36F}' | '\u{203F}'..='\u{2040}'
        )
}
