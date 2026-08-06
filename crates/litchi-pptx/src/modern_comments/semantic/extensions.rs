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
    pub fn new(xml: Vec<u8>) -> Result<Self> {
        if xml.len() > MAX_BYTES {
            return Err(invalid(
                "modern comment opaque XML exceeds implementation limit",
            ));
        }
        Ok(Self { xml })
    }

    #[inline]
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }
}

/// One `p:ext` entry in a modern comment extension list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionEntry {
    pub uri: String,
    pub payload: ExtensionPayload,
}

/// Known versioned payloads and an inert preservation branch for all others.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExtensionPayload {
    TaskDetails(super::TaskDetails),
    Reactions(super::Reactions),
    /// Contains the complete original `p:ext` element, including its URI and
    /// namespace declarations. `uri` is retained for inspection only.
    Opaque(OpaqueXml),
}

/// The complete `p188:extLst` envelope attached to a comment or reply.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ExtensionList {
    pub root_prefix: String,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
    pub entries: Vec<ExtensionEntry>,
}

impl ExtensionList {
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
                ExtensionPayload::TaskDetails(value) => {
                    task_count += 1;
                    value.validate()?;
                },
                ExtensionPayload::Reactions(value) => {
                    reaction_count += 1;
                    value.validate()?;
                },
                ExtensionPayload::Opaque(value) => {
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

    pub fn task_details(&self) -> Option<&super::TaskDetails> {
        self.entries.iter().find_map(|entry| match &entry.payload {
            ExtensionPayload::TaskDetails(value) => Some(value),
            _ => None,
        })
    }

    pub fn reactions(&self) -> Option<&super::Reactions> {
        self.entries.iter().find_map(|entry| match &entry.payload {
            ExtensionPayload::Reactions(value) => Some(value),
            _ => None,
        })
    }

    /// Replace or remove task details while retaining all opaque entries.
    /// `uri` is required only when inserting a new payload because the XML
    /// extension envelope has no safe universally applicable default URI.
    pub fn replace_task_details(
        &mut self,
        uri: Option<&str>,
        value: Option<super::TaskDetails>,
    ) -> Result<()> {
        self.replace_known(
            uri,
            value.map(ExtensionPayload::TaskDetails),
            "task details",
        )
    }

    /// Replace or remove reactions while retaining all opaque entries.
    pub fn replace_reactions(
        &mut self,
        uri: Option<&str>,
        value: Option<super::Reactions>,
    ) -> Result<()> {
        self.replace_known(uri, value.map(ExtensionPayload::Reactions), "reactions")
    }

    fn replace_known(
        &mut self,
        uri: Option<&str>,
        value: Option<ExtensionPayload>,
        label: &str,
    ) -> Result<()> {
        let matches: Vec<usize> = self
            .entries
            .iter()
            .enumerate()
            .filter_map(|(index, entry)| {
                let known = match label {
                    "task details" => {
                        matches!(&entry.payload, ExtensionPayload::TaskDetails(_))
                    },
                    "reactions" => matches!(&entry.payload, ExtensionPayload::Reactions(_)),
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
                self.entries.push(ExtensionEntry {
                    uri: uri.to_owned(),
                    payload,
                });
            },
            (None, None) => {},
        }
        self.validate()
    }
}
