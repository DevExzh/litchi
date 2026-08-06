use super::super::model::{NamespaceDeclaration, Progress};
use super::extensions::OpaqueXml;
use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::custom_xml::valid_guid;

const MAX_EVENTS: usize = 100_000;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn bounded(value: &str, label: &str) -> Result<()> {
    if value.len() <= super::super::MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(invalid(format!("{label} is too long")))
    }
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct User {
    pub author_id: String,
}

impl User {
    pub fn validate(&self) -> Result<()> {
        guid(&self.author_id, "task author")
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Anchor {
    pub comment_id: String,
    pub extension_xml: Option<OpaqueXml>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl Anchor {
    pub fn validate(&self) -> Result<()> {
        guid(&self.comment_id, "task anchor")?;
        if let Some(value) = &self.extension_xml {
            if value.xml.len() > super::super::MAX_BYTES {
                return Err(invalid("task anchor extension is too large"));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Assign {
    pub author_id: String,
}

impl Assign {
    pub fn validate(&self) -> Result<()> {
        guid(&self.author_id, "assigned task author")
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Title {
    pub value: String,
}

impl Title {
    pub fn validate(&self) -> Result<()> {
        bounded(&self.value, "task title")
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Schedule {
    pub start_date: Option<String>,
    pub end_date: Option<String>,
}

impl Schedule {
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = &self.start_date {
            date_time(value, "task start date")?;
        }
        if let Some(value) = &self.end_date {
            date_time(value, "task end date")?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Undo {
    pub event_id: String,
}

impl Undo {
    pub fn validate(&self) -> Result<()> {
        guid(&self.event_id, "undone task event")
    }
}

/// The schema's single task-history choice. `Unknown` retains future event
/// records without assigning them behavior.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Action {
    Assign(Assign),
    Add,
    Title(Title),
    Schedule(Schedule),
    Progress(Progress),
    UnassignAll,
    Undo(Undo),
    Unknown(OpaqueXml),
}

impl Action {
    pub fn validate(&self) -> Result<()> {
        match self {
            Self::Assign(value) => value.validate(),
            Self::Add | Self::UnassignAll => Ok(()),
            Self::Title(value) => value.validate(),
            Self::Schedule(value) => value.validate(),
            Self::Progress(_) => Ok(()),
            Self::Undo(value) => value.validate(),
            Self::Unknown(value) if value.xml.len() <= super::super::MAX_BYTES => Ok(()),
            Self::Unknown(_) => Err(invalid("unknown task event is too large")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Event {
    pub time: String,
    pub id: String,
    pub attributed_by: User,
    pub anchor: Option<Anchor>,
    pub action: Option<Action>,
    pub extension_xml: Option<OpaqueXml>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl Event {
    pub fn validate(&self) -> Result<()> {
        date_time(&self.time, "task history event")?;
        guid(&self.id, "task history event")?;
        self.attributed_by.validate()?;
        if let Some(value) = &self.anchor {
            value.validate()?;
        }
        if let Some(value) = &self.action {
            value.validate()?;
        }
        if let Some(value) = &self.extension_xml {
            if value.xml.len() > super::super::MAX_BYTES {
                return Err(invalid("task history event extension is too large"));
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct History {
    pub events: Vec<Event>,
}

impl History {
    pub fn validate(&self) -> Result<()> {
        if self.events.len() > MAX_EVENTS {
            return Err(invalid("task history events exceed implementation limit"));
        }
        for value in &self.events {
            value.validate()?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Details {
    pub history: History,
    pub extension_xml: Option<OpaqueXml>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
}

impl Details {
    pub fn validate(&self) -> Result<()> {
        self.history.validate()?;
        if let Some(value) = &self.extension_xml {
            if value.xml.len() > super::super::MAX_BYTES {
                return Err(invalid("task details extension is too large"));
            }
        }
        Ok(())
    }
}
