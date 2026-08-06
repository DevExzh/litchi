//! Canonical modern-comment values shared by the XML codec and package graph.

use crate::{Error, Result};
use std::fmt;
use std::num::NonZeroU32;
use std::str::FromStr;

const MAX_PROGRESS_THOUSANDTHS: u32 = 100_000;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamespaceDeclaration {
    /// Empty means the default namespace.
    pub prefix: String,
    pub uri: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Status {
    Active,
    Resolved,
    Closed,
}

impl Status {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "active" => Ok(Self::Active),
            "resolved" => Ok(Self::Resolved),
            "closed" => Ok(Self::Closed),
            _ => Err(invalid(format!("invalid modern comment status '{value}'"))),
        }
    }

    pub(super) fn token(self) -> &'static str {
        match self {
            Self::Active => "active",
            Self::Resolved => "resolved",
            Self::Closed => "closed",
        }
    }
}

/// Completion progress in Office's thousandths-of-a-percent units.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Progress(NonZeroU32);

impl Progress {
    /// No progress (0%).
    pub const ZERO: Self = Self(NonZeroU32::MIN);
    /// Complete (100%).
    pub const FULL: Self = match NonZeroU32::new(MAX_PROGRESS_THOUSANDTHS + 1) {
        Some(value) => Self(value),
        None => Self::ZERO,
    };

    /// Construct progress from a whole percentage in 0..=100.
    pub fn new(percent: u32) -> Result<Self> {
        if percent <= 100 {
            Self::from_thousandths(percent * 1_000)
        } else {
            Err(invalid(format!(
                "modern-comment progress {percent}% exceeds 100%"
            )))
        }
    }

    /// Construct progress from Office's thousandths-of-a-percent units.
    pub fn from_thousandths(value: u32) -> Result<Self> {
        if value <= MAX_PROGRESS_THOUSANDTHS {
            NonZeroU32::new(value + 1).map(Self).ok_or_else(|| {
                invalid("modern-comment progress could not be represented compactly")
            })
        } else {
            Err(invalid(format!(
                "modern-comment progress {value} exceeds {MAX_PROGRESS_THOUSANDTHS} thousandths"
            )))
        }
    }

    /// Return Office's thousandths-of-a-percent representation.
    #[inline]
    pub const fn thousandths(self) -> u32 {
        self.0.get() - 1
    }

    fn parse_percent(value: &str) -> Result<Self> {
        let (whole, fraction) = value
            .split_once('.')
            .map_or((value, None), |(whole, fraction)| (whole, Some(fraction)));
        let valid_whole = !whole.is_empty()
            && whole.len() <= 3
            && whole.bytes().all(|byte| byte.is_ascii_digit())
            && (whole.len() < 3 || whole == "100");
        let valid_fraction = fraction.is_none_or(|fraction| {
            (1..=2).contains(&fraction.len()) && fraction.bytes().all(|byte| byte.is_ascii_digit())
        });
        if !valid_whole || !valid_fraction {
            return Err(invalid(format!(
                "invalid positive fixed percentage '{value}%'"
            )));
        }

        let whole = whole
            .parse::<u32>()
            .map_err(|_| invalid(format!("invalid positive fixed percentage '{value}%'")))?;
        let fractional_thousandths = match fraction {
            Some(fraction) => {
                let mut digits = fraction.bytes().map(|digit| u32::from(digit - b'0'));
                let tenths = digits.next().ok_or_else(|| {
                    invalid(format!("invalid positive fixed percentage '{value}%'"))
                })?;
                tenths * 100 + digits.next().unwrap_or_default() * 10
            },
            None => 0,
        };
        let thousandths = whole * 1_000 + fractional_thousandths;
        Self::from_thousandths(thousandths)
            .map_err(|_| invalid(format!("invalid positive fixed percentage '{value}%'")))
    }
}

impl FromStr for Progress {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        if let Some(percent) = value.strip_suffix('%') {
            return Self::parse_percent(percent);
        }
        let digits = value.strip_prefix('+').unwrap_or(value);
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(invalid(format!(
                "invalid positive fixed percentage '{value}'"
            )));
        }
        let thousandths = value
            .parse::<u32>()
            .map_err(|_| invalid(format!("invalid positive fixed percentage '{value}'")))?;
        Self::from_thousandths(thousandths)
    }
}

impl fmt::Display for Progress {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(&self.thousandths(), formatter)
    }
}

impl Default for Progress {
    fn default() -> Self {
        Self::ZERO
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnchorKind {
    SlideMoniker,
    DrawingElementMoniker,
    TextRangeMoniker,
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Anchor {
    pub kind: AnchorKind,
    /// Complete imported moniker element retained inertly.
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Position {
    pub x: i64,
    pub y: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reply {
    pub id: String,
    pub author_id: String,
    /// None retains omission and the schema default active.
    pub status: Option<Status>,
    pub created: String,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
    /// Optional complete p188:txBody fragment retained inertly.
    pub text_body_xml: Option<Vec<u8>>,
    /// Optional complete p188:extLst fragment retained inertly.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub id: String,
    pub author_id: String,
    /// None retains omission and the schema default active.
    pub status: Option<Status>,
    pub created: String,
    pub start_date: Option<String>,
    pub due_date: Option<String>,
    /// None distinguishes omission from a present empty list.
    pub assigned_to: Option<Vec<String>>,
    /// Completion progress. None retains omission and the schema default 0%.
    pub complete: Option<Progress>,
    pub title: Option<String>,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
    pub anchors: Vec<Anchor>,
    pub position: Option<Position>,
    pub reply_list_namespace_declarations: Vec<NamespaceDeclaration>,
    pub replies: Vec<Reply>,
    /// Whether the optional replyLst wrapper was present, including when empty.
    pub reply_list_present: bool,
    /// Optional complete p188:txBody fragment retained inertly.
    pub text_body_xml: Option<Vec<u8>>,
    /// Optional complete p188:extLst fragment retained inertly.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct List {
    /// Prefix used for the 2018 PowerPoint namespace. Empty means default.
    pub root_prefix: String,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
    pub comments: Vec<Comment>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    pub slide_part_name: String,
    pub relationship_id: String,
    pub part_name: String,
    pub comments: List,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Author {
    pub id: String,
    pub name: String,
    pub initials: Option<String>,
    pub user_id: String,
    pub provider_id: String,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
    /// Optional complete p188:extLst fragment retained inertly.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Authors {
    /// Prefix used for the 2018 PowerPoint namespace. Empty means default.
    pub root_prefix: String,
    pub namespace_declarations: Vec<NamespaceDeclaration>,
    pub authors: Vec<Author>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthorPart {
    pub relationship_id: String,
    pub part_name: String,
    pub authors: Authors,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Graph {
    pub authors: Option<AuthorPart>,
    pub comments: Vec<Part>,
}

impl Comment {
    /// Decode the inert `p188:extLst` envelope into its typed known payloads
    /// and preserved unknown entries.
    pub fn extensions(&self) -> Result<super::semantic::ExtensionList> {
        super::wire::parse_extensions(self.extension_xml.as_deref())
    }

    /// Replace the extension envelope atomically after validating every typed
    /// payload. Empty lists remove the optional XML element.
    pub fn set_extensions(&mut self, value: super::semantic::ExtensionList) -> Result<()> {
        self.extension_xml = super::wire::write_extensions(&value)?;
        Ok(())
    }

    pub fn task_details(&self) -> Result<Option<super::semantic::TaskDetails>> {
        Ok(self.extensions()?.task_details().cloned())
    }

    pub fn reactions(&self) -> Result<Option<super::semantic::Reactions>> {
        Ok(self.extensions()?.reactions().cloned())
    }

    pub fn replace_task_details(
        &mut self,
        uri: Option<&str>,
        value: Option<super::semantic::TaskDetails>,
    ) -> Result<()> {
        let mut extensions = self.extensions()?;
        extensions.replace_task_details(uri, value)?;
        self.set_extensions(extensions)
    }

    pub fn replace_reactions(
        &mut self,
        uri: Option<&str>,
        value: Option<super::semantic::Reactions>,
    ) -> Result<()> {
        let mut extensions = self.extensions()?;
        extensions.replace_reactions(uri, value)?;
        self.set_extensions(extensions)
    }
}

impl Reply {
    pub fn extensions(&self) -> Result<super::semantic::ExtensionList> {
        super::wire::parse_extensions(self.extension_xml.as_deref())
    }

    pub fn set_extensions(&mut self, value: super::semantic::ExtensionList) -> Result<()> {
        self.extension_xml = super::wire::write_extensions(&value)?;
        Ok(())
    }

    pub fn reactions(&self) -> Result<Option<super::semantic::Reactions>> {
        Ok(self.extensions()?.reactions().cloned())
    }

    pub fn replace_reactions(
        &mut self,
        uri: Option<&str>,
        value: Option<super::semantic::Reactions>,
    ) -> Result<()> {
        let mut extensions = self.extensions()?;
        extensions.replace_reactions(uri, value)?;
        self.set_extensions(extensions)
    }
}
