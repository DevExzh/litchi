//! The header and footer parts a document relates, listed without decoding
//! them (record 0319, L-DOCX-R5).

use crate::error::Result;

use super::Role;

/// One header or footer part the main document relates: its role, its
/// relationship identifier and its part name.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Entry {
    role: Role,
    relationship_id: String,
    name: String,
}

impl Entry {
    pub(crate) const fn new(role: Role, relationship_id: String, name: String) -> Self {
        Self {
            role,
            relationship_id,
            name,
        }
    }

    /// Whether the part is a header or a footer.
    #[must_use]
    pub const fn role(&self) -> Role {
        self.role
    }

    /// The main part's relationship identifier of the part, which section
    /// properties reference.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// The part name (`/word/header1.xml`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// The header and footer parts `relationships` (the main part's) relate, in
/// relationship identifier order (`rId2` before `rId10`), each part once.
/// The order does not depend on how the relationships are stored, so both
/// owners list the same parts at the same positions.
pub(crate) fn entries_of(relationships: &litchi_opc::Relationships) -> Result<Vec<Entry>> {
    let mut entries: Vec<Entry> = Vec::new();
    for relationship in relationships.iter() {
        if relationship.is_external() {
            continue;
        }
        let role = if relationship.reltype().ends_with("/header") {
            Role::Header
        } else if relationship.reltype().ends_with("/footer") {
            Role::Footer
        } else {
            continue;
        };
        let name = relationship.target_partname()?.as_str().to_owned();
        entries.push(Entry::new(role, relationship.r_id().to_owned(), name));
    }
    entries.sort_by(|left, right| {
        natural(&left.relationship_id).cmp(&natural(&right.relationship_id))
    });
    let mut seen: Vec<String> = Vec::with_capacity(entries.len());
    entries.retain(|entry| {
        let first = !seen.contains(&entry.name);
        if first {
            seen.push(entry.name.clone());
        }
        first
    });
    Ok(entries)
}

/// A relationship identifier's sort key: its text before the trailing
/// digits, then their value, then the identifier itself.
fn natural(id: &str) -> (&str, u64, &str) {
    let digits = id.len()
        - id.trim_end_matches(|character: char| character.is_ascii_digit())
            .len();
    let (stem, number) = id.split_at(id.len() - digits);
    (stem, number.parse().unwrap_or(u64::MAX), id)
}

/// The header and footer references of a main-document payload, in
/// document order: each with its role, section kind and relationship
/// identifier. A reference without `w:type` is the default kind.
pub(crate) fn references_of(xml: &[u8]) -> Result<Vec<(Role, crate::section::Reference)>> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut found = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let role = match element.local_name().as_ref() {
                    b"headerReference" => Role::Header,
                    b"footerReference" => Role::Footer,
                    _ => continue,
                };
                let mut kind = super::Kind::Primary;
                let mut relationship_id = None;
                for attribute in element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                {
                    let value = String::from_utf8_lossy(&attribute.value).into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"type" => {
                            kind = super::Kind::from_xml(&value).unwrap_or(super::Kind::Primary)
                        },
                        b"id" => relationship_id = Some(value),
                        _ => {},
                    }
                }
                if let Some(relationship_id) = relationship_id {
                    found.push((
                        role,
                        crate::section::Reference {
                            kind,
                            relationship_id,
                        },
                    ));
                }
            },
            Event::Eof => return Ok(found),
            _ => {},
        }
    }
}
