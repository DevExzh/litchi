use super::super::MAX_STRING_BYTES;
use super::extensions::OpaqueXml;
use crate::Error;
use crate::Result;
use litchi_ooxml_common::custom_xml::valid_guid;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(invalid("modern comment moniker ID is too long"))
    }
}

fn validate_guid(value: &str) -> Result<()> {
    if valid_guid(value) {
        Ok(())
    } else {
        Err(invalid(format!(
            "invalid modern comment moniker GUID '{value}'"
        )))
    }
}

/// The typed terminal expected by a 2.18 moniker list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Comment,
    Reply,
}

/// A known comment/reply moniker or an inert inherited moniker fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Node {
    Opaque(OpaqueXml),
    Comment { id: String },
    Reply { id: String },
}

/// Ordered content monikers from `cmMkLst` or `cmRplyMkLst`.
///
/// The slide, master, layout, drawing, and text monikers defined by earlier
/// command schemas are intentionally retained as opaque nodes. The terminal
/// comment and reply IDs remain typed and validated.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct List {
    pub kind: Kind,
    pub nodes: Vec<Node>,
}

impl List {
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn comment(nodes: Vec<Node>) -> Result<Self> {
        let value = Self {
            kind: Kind::Comment,
            nodes,
        };
        value.validate()?;
        Ok(value)
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reply(nodes: Vec<Node>) -> Result<Self> {
        let value = Self {
            kind: Kind::Reply,
            nodes,
        };
        value.validate()?;
        Ok(value)
    }

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn validate(&self) -> Result<()> {
        let mut comments = 0usize;
        let mut replies = 0usize;
        for node in &self.nodes {
            match node {
                Node::Opaque(value) => {
                    if value.xml.len() > super::super::MAX_BYTES {
                        return Err(invalid("modern comment moniker fragment is too large"));
                    }
                },
                Node::Comment { id } => {
                    validate_guid(id)?;
                    bounded(id)?;
                    comments += 1;
                },
                Node::Reply { id } => {
                    validate_guid(id)?;
                    bounded(id)?;
                    replies += 1;
                },
            }
        }
        match self.kind {
            Kind::Comment if comments == 1 && replies == 0 => Ok(()),
            Kind::Reply if comments == 1 && replies == 1 => Ok(()),
            Kind::Comment => Err(invalid(
                "comment moniker list requires exactly one comment moniker",
            )),
            Kind::Reply => Err(invalid(
                "reply moniker list requires one comment and one reply moniker",
            )),
        }
    }

    #[must_use]
    pub fn comment_id(&self) -> Option<&str> {
        self.nodes.iter().find_map(|node| match node {
            Node::Comment { id } => Some(id.as_str()),
            _ => None,
        })
    }

    #[must_use]
    pub fn reply_id(&self) -> Option<&str> {
        self.nodes.iter().find_map(|node| match node {
            Node::Reply { id } => Some(id.as_str()),
            _ => None,
        })
    }
}
