use super::super::model::NamespaceDeclaration;
use super::extensions::OpaqueXml;
use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::custom_xml::valid_guid;

const MAX_REACTIONS: usize = 100_000;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn date_time(value: &str) -> Result<()> {
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid reaction XML dateTime '{value}'")))
    }
}

/// One user instance from `p223:reaction/instance`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Instance {
    pub time: String,
    pub author_id: String,
    pub extension_xml: Option<OpaqueXml>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl Instance {
    pub fn validate(&self) -> Result<()> {
        date_time(&self.time)?;
        if !valid_guid(&self.author_id) {
            return Err(invalid(format!(
                "invalid reaction author GUID '{}'",
                self.author_id
            )));
        }
        if let Some(value) = &self.extension_xml {
            if value.xml.len() > super::super::MAX_BYTES {
                return Err(invalid("reaction instance extension is too large"));
            }
        }
        Ok(())
    }
}

/// All instances of one reaction type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reaction {
    pub reaction_type: String,
    pub instances: Vec<Instance>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl Reaction {
    pub fn validate(&self) -> Result<()> {
        if self.reaction_type.is_empty()
            || self.reaction_type.len() > super::super::MAX_STRING_BYTES
        {
            return Err(invalid("reaction type must be nonempty and bounded"));
        }
        if self.instances.len() > MAX_REACTIONS {
            return Err(invalid("reaction instances exceed implementation limit"));
        }
        for value in &self.instances {
            value.validate()?;
        }
        Ok(())
    }
}

/// The 2.21 `p223:reactions` payload.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct List {
    pub reactions: Vec<Reaction>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl List {
    pub fn validate(&self) -> Result<()> {
        if self.reactions.len() > MAX_REACTIONS {
            return Err(invalid("reactions exceed implementation limit"));
        }
        for value in &self.reactions {
            value.validate()?;
        }
        Ok(())
    }
}
