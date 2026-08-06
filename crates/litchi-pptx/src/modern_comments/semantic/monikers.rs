use super::super::MAX_STRING_BYTES;
use super::OpaqueXml;
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
pub enum MonikerKind {
    Comment,
    Reply,
}

/// A known comment/reply moniker or an inert inherited moniker fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MonikerNode {
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
pub struct MonikerList {
    pub kind: MonikerKind,
    pub nodes: Vec<MonikerNode>,
}

impl MonikerList {
    pub fn comment(nodes: Vec<MonikerNode>) -> Result<Self> {
        let value = Self {
            kind: MonikerKind::Comment,
            nodes,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn reply(nodes: Vec<MonikerNode>) -> Result<Self> {
        let value = Self {
            kind: MonikerKind::Reply,
            nodes,
        };
        value.validate()?;
        Ok(value)
    }

    pub fn validate(&self) -> Result<()> {
        let mut comments = 0usize;
        let mut replies = 0usize;
        for node in &self.nodes {
            match node {
                MonikerNode::Opaque(value) => {
                    if value.xml.len() > super::super::MAX_BYTES {
                        return Err(invalid("modern comment moniker fragment is too large"));
                    }
                },
                MonikerNode::Comment { id } => {
                    validate_guid(id)?;
                    bounded(id)?;
                    comments += 1;
                },
                MonikerNode::Reply { id } => {
                    validate_guid(id)?;
                    bounded(id)?;
                    replies += 1;
                },
            }
        }
        match self.kind {
            MonikerKind::Comment if comments == 1 && replies == 0 => Ok(()),
            MonikerKind::Reply if comments == 1 && replies == 1 => Ok(()),
            MonikerKind::Comment => Err(invalid(
                "comment moniker list requires exactly one comment moniker",
            )),
            MonikerKind::Reply => Err(invalid(
                "reply moniker list requires one comment and one reply moniker",
            )),
        }
    }

    pub fn comment_id(&self) -> Option<&str> {
        self.nodes.iter().find_map(|node| match node {
            MonikerNode::Comment { id } => Some(id.as_str()),
            _ => None,
        })
    }

    pub fn reply_id(&self) -> Option<&str> {
        self.nodes.iter().find_map(|node| match node {
            MonikerNode::Reply { id } => Some(id.as_str()),
            _ => None,
        })
    }
}
