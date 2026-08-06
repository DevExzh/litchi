use super::super::model::NamespaceDeclaration;
use super::{MonikerList, OpaqueXml};
use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::custom_xml::valid_guid;

const MAX_CHANGES: usize = 100_000;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn guid(value: &str, label: &str) -> Result<()> {
    if valid_guid(value) {
        Ok(())
    } else {
        Err(invalid(format!("invalid {label} GUID '{value}'")))
    }
}

fn date_time(value: &str, label: &str) -> Result<()> {
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid {label} XML dateTime '{value}'")))
    }
}

/// Typed `ac:CT_ChangesData` metadata. It is storage metadata only.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChangeMetadata {
    pub name: Option<String>,
    pub user_id: Option<String>,
    pub provider_id: Option<String>,
    pub client_id: Option<String>,
    pub email: Option<String>,
    pub date_time: Option<String>,
    pub version: Option<u32>,
    pub change_id: Option<String>,
    pub action_id: Option<i32>,
    pub extension_xml: Option<OpaqueXml>,
}

impl ChangeMetadata {
    pub fn validate(&self) -> Result<()> {
        for value in [
            self.name.as_deref(),
            self.user_id.as_deref(),
            self.provider_id.as_deref(),
            self.client_id.as_deref(),
            self.email.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            if value.len() > super::super::MAX_STRING_BYTES {
                return Err(invalid("modern comment change metadata is too long"));
            }
        }
        if let Some(value) = &self.date_time {
            date_time(value, "change metadata")?;
        }
        if let Some(value) = &self.change_id {
            guid(value, "change metadata")?;
        }
        if let Some(value) = &self.extension_xml {
            if value.xml.len() > super::super::MAX_BYTES {
                return Err(invalid(
                    "modern comment change metadata extension is too large",
                ));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CommentChange {
    Add,
    Delete,
    Modify,
    ModifyTask,
    ModifyReaction,
}

impl CommentChange {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "add" => Ok(Self::Add),
            "del" => Ok(Self::Delete),
            "mod" => Ok(Self::Modify),
            "modTsk" => Ok(Self::ModifyTask),
            "modRxn" => Ok(Self::ModifyReaction),
            _ => Err(invalid(format!(
                "unknown modern comment change bit '{value}'"
            ))),
        }
    }

    pub(crate) const fn token(self) -> &'static str {
        match self {
            Self::Add => "add",
            Self::Delete => "del",
            Self::Modify => "mod",
            Self::ModifyTask => "modTsk",
            Self::ModifyReaction => "modRxn",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ReplyChange {
    Add,
    Delete,
    Modify,
    ModifyReaction,
}

impl ReplyChange {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "add" => Ok(Self::Add),
            "del" => Ok(Self::Delete),
            "mod" => Ok(Self::Modify),
            "modRxn" => Ok(Self::ModifyReaction),
            _ => Err(invalid(format!(
                "unknown modern reply change bit '{value}'"
            ))),
        }
    }

    pub(crate) const fn token(self) -> &'static str {
        match self {
            Self::Add => "add",
            Self::Delete => "del",
            Self::Modify => "mod",
            Self::ModifyReaction => "modRxn",
        }
    }
}

fn validate_bits<T: Copy + Eq>(bits: &[T], label: &str) -> Result<()> {
    if bits.is_empty() {
        return Err(invalid(format!("{label} requires at least one change bit")));
    }
    for (index, bit) in bits.iter().enumerate() {
        if bits[..index].contains(bit) {
            return Err(invalid(format!("{label} contains a duplicate change bit")));
        }
    }
    Ok(())
}

/// A typed 2.19 `CT_CommentReplyV2Changes` descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReplyChanges {
    pub changes: Vec<ReplyChange>,
    pub metadata: Option<ChangeMetadata>,
    pub monikers: MonikerList,
    pub extension_xml: Option<OpaqueXml>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl ReplyChanges {
    pub fn validate(&self) -> Result<()> {
        validate_bits(&self.changes, "comment reply changes")?;
        if self.monikers.kind != super::MonikerKind::Reply {
            return Err(invalid("reply changes require a reply moniker list"));
        }
        self.monikers.validate()?;
        if let Some(value) = &self.metadata {
            value.validate()?;
        }
        if let Some(value) = &self.extension_xml {
            if value.xml.len() > super::super::MAX_BYTES {
                return Err(invalid("comment reply change extension is too large"));
            }
        }
        Ok(())
    }
}

/// A typed 2.19 `CT_CommentV2Changes` descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentChanges {
    pub changes: Vec<CommentChange>,
    pub metadata: Option<ChangeMetadata>,
    pub monikers: MonikerList,
    pub reply_changes: Vec<ReplyChanges>,
    pub extension_xml: Option<OpaqueXml>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl CommentChanges {
    pub fn validate(&self) -> Result<()> {
        validate_bits(&self.changes, "comment changes")?;
        if self.monikers.kind != super::MonikerKind::Comment {
            return Err(invalid("comment changes require a comment moniker list"));
        }
        self.monikers.validate()?;
        if self.reply_changes.len() > MAX_CHANGES {
            return Err(invalid("comment reply changes exceed implementation limit"));
        }
        if let Some(value) = &self.metadata {
            value.validate()?;
        }
        for value in &self.reply_changes {
            value.validate()?;
        }
        if let Some(value) = &self.extension_xml {
            if value.xml.len() > super::super::MAX_BYTES {
                return Err(invalid("comment change extension is too large"));
            }
        }
        Ok(())
    }
}
