use super::super::model::NamespaceDeclaration;
use super::super::{MAX_BYTES, MAX_STRING_BYTES};
use crate::{Error, Result};

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// An XML fragment retained without interpreting its application semantics.
///
/// The fragment is bounded and is emitted only as XML data by the owner. It
/// is never followed as a relationship, loaded as a resource, or executed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueXml {
    pub xml: Vec<u8>,
}

impl OpaqueXml {
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(xml: Vec<u8>) -> Result<Self> {
        if xml.len() > MAX_BYTES {
            return Err(invalid(
                "modern comment opaque XML exceeds implementation limit",
            ));
        }
        Ok(Self { xml })
    }

    #[inline]
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }
}

/// One `p:ext` entry in a modern comment extension list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    pub uri: String,
    pub payload: Payload,
}

/// Known versioned payloads and an inert preservation branch for all others.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Payload {
    TaskDetails(super::tasks::Details),
    Reactions(super::reactions::List),
    /// Contains the complete original `p:ext` element, including its URI and
    /// namespace declarations. `uri` is retained for inspection only.
    Opaque(OpaqueXml),
}

/// The complete `p188:extLst` envelope attached to a comment or reply.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct List {
    pub root_prefix: String,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
    pub entries: Vec<Entry>,
}

impl List {
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn validate(&self) -> Result<()> {
        if self.root_prefix.len() > MAX_STRING_BYTES {
            return Err(invalid("modern comment extension prefix is too long"));
        }
        let mut task_count = 0usize;
        let mut reaction_count = 0usize;
        for entry in &self.entries {
            if entry.uri.len() > MAX_STRING_BYTES || entry.uri.trim().is_empty() {
                return Err(invalid(
                    "modern comment extension URI must be nonempty and bounded",
                ));
            }
            match &entry.payload {
                Payload::TaskDetails(value) => {
                    task_count += 1;
                    value.validate()?;
                },
                Payload::Reactions(value) => {
                    reaction_count += 1;
                    value.validate()?;
                },
                Payload::Opaque(value) => {
                    if value.xml.len() > MAX_BYTES {
                        return Err(invalid("modern comment opaque extension is too large"));
                    }
                },
            }
        }
        if task_count > 1 || reaction_count > 1 {
            return Err(invalid(
                "modern comment extension list contains a duplicate typed payload",
            ));
        }
        Ok(())
    }

    #[must_use]
    pub fn task_details(&self) -> Option<&super::tasks::Details> {
        self.entries.iter().find_map(|entry| match &entry.payload {
            Payload::TaskDetails(value) => Some(value),
            _ => None,
        })
    }

    #[must_use]
    pub fn reactions(&self) -> Option<&super::reactions::List> {
        self.entries.iter().find_map(|entry| match &entry.payload {
            Payload::Reactions(value) => Some(value),
            _ => None,
        })
    }

    /// Replace or remove task details while retaining all opaque entries.
    /// `uri` is required only when inserting a new payload because the XML
    /// extension envelope has no safe universally applicable default URI.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_task_details(
        &mut self,
        uri: Option<&str>,
        value: Option<super::tasks::Details>,
    ) -> Result<()> {
        self.replace_known(uri, value.map(Payload::TaskDetails), "task details")
    }

    /// Replace or remove reactions while retaining all opaque entries.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_reactions(
        &mut self,
        uri: Option<&str>,
        value: Option<super::reactions::List>,
    ) -> Result<()> {
        self.replace_known(uri, value.map(Payload::Reactions), "reactions")
    }

    fn replace_known(
        &mut self,
        uri: Option<&str>,
        value: Option<Payload>,
        label: &str,
    ) -> Result<()> {
        let matches: Vec<usize> = self
            .entries
            .iter()
            .enumerate()
            .filter_map(|(index, entry)| {
                let known = match label {
                    "task details" => {
                        matches!(&entry.payload, Payload::TaskDetails(_))
                    },
                    "reactions" => matches!(&entry.payload, Payload::Reactions(_)),
                    _ => false,
                };
                known.then_some(index)
            })
            .collect();
        if matches.len() > 1 {
            return Err(invalid(format!("duplicate modern comment {label}")));
        }
        match (matches.first().copied(), value) {
            (Some(index), Some(payload)) => {
                self.entries[index].payload = payload;
            },
            (Some(index), None) => {
                self.entries.remove(index);
            },
            (None, Some(payload)) => {
                let uri = uri
                    .filter(|value| !value.trim().is_empty())
                    .ok_or_else(|| invalid(format!("a URI is required to insert {label}")))?;
                self.entries.push(Entry {
                    uri: uri.to_owned(),
                    payload,
                });
            },
            (None, None) => {},
        }
        self.validate()
    }
}
