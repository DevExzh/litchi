use crate::Result;

use super::support::{escape_attribute, invalid};

/// Type of document protection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProtectionType {
    /// No editing allowed.
    ReadOnly,
    /// Only comments allowed.
    Comments,
    /// Only tracked changes allowed.
    TrackedChanges,
    /// Only form fields allowed.
    Forms,
}

impl ProtectionType {
    /// Parse the optional `w:edit` token.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "readOnly" => Some(Self::ReadOnly),
            "comments" => Some(Self::Comments),
            "trackedChanges" => Some(Self::TrackedChanges),
            "forms" => Some(Self::Forms),
            _ => None,
        }
    }

    /// Get the XML value for this protection type.
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::ReadOnly => "readOnly",
            Self::Comments => "comments",
            Self::TrackedChanges => "trackedChanges",
            Self::Forms => "forms",
        }
    }
}

/// Document view mode from `w:view` (`ST_View`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum View {
    /// No explicit view is specified.
    None,
    /// Print layout view.
    Print,
    /// Outline view.
    Outline,
    /// Master pages view.
    MasterPages,
    /// Normal (draft) view.
    Normal,
    /// Web layout view.
    Web,
}

impl View {
    /// Parse the schema token.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "print" => Ok(Self::Print),
            "outline" => Ok(Self::Outline),
            "masterPages" => Ok(Self::MasterPages),
            "normal" => Ok(Self::Normal),
            "web" => Ok(Self::Web),
            _ => Err(invalid(format!("invalid document view value '{value}'"))),
        }
    }

    /// Get the XML value for this view mode.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Print => "print",
            Self::Outline => "outline",
            Self::MasterPages => "masterPages",
            Self::Normal => "normal",
            Self::Web => "web",
        }
    }
}

/// Proofing completion marker (`ST_ProofState`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProofState {
    /// Proofing for the region completed without errors.
    Clean,
    /// The region changed since proofing last ran.
    Dirty,
}

impl ProofState {
    /// Parse the schema token.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "clean" => Ok(Self::Clean),
            "dirty" => Ok(Self::Dirty),
            _ => Err(invalid(format!("invalid proof state value '{value}'"))),
        }
    }

    /// Get the XML value for this proof state.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Clean => "clean",
            Self::Dirty => "dirty",
        }
    }
}

/// Proofing completion markers from `w:proofState`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ProofingState {
    spelling: Option<ProofState>,
    grammar: Option<ProofState>,
}

impl ProofingState {
    /// Create a proofing state with no markers.
    pub const fn new() -> Self {
        Self {
            spelling: None,
            grammar: None,
        }
    }

    /// Set the spelling proofing marker.
    pub fn set_spelling(&mut self, value: Option<ProofState>) -> &mut Self {
        self.spelling = value;
        self
    }

    /// Set the grammar proofing marker.
    pub fn set_grammar(&mut self, value: Option<ProofState>) -> &mut Self {
        self.grammar = value;
        self
    }

    /// Return the spelling proofing marker, when specified.
    #[inline]
    pub const fn spelling(&self) -> Option<ProofState> {
        self.spelling
    }

    /// Return the grammar proofing marker, when specified.
    #[inline]
    pub const fn grammar(&self) -> Option<ProofState> {
        self.grammar
    }

    /// Serialize a standalone `w:proofState` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:proofState");
        if let Some(spelling) = self.spelling {
            xml.push_str(&format!(" {prefix}:spelling=\"{}\"", spelling.as_str()));
        }
        if let Some(grammar) = self.grammar {
            xml.push_str(&format!(" {prefix}:grammar=\"{}\"", grammar.as_str()));
        }
        xml.push_str("/>");
        xml
    }
}

/// Maximum accepted length of a `w:themeFontLang` language tag.
pub const MAX_LANGUAGE_TAG_LENGTH: usize = 255;

fn validate_language_tag(value: &str, description: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_LANGUAGE_TAG_LENGTH
        || value.chars().any(char::is_control)
    {
        return Err(invalid(format!(
            "invalid Word {description} language tag '{value}'"
        )));
    }
    Ok(())
}

/// Theme font language defaults from `w:themeFontLang`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ThemeFontLanguages {
    latin: Option<String>,
    east_asia: Option<String>,
    bidi: Option<String>,
}

impl ThemeFontLanguages {
    /// Create theme font language defaults with no languages set.
    pub const fn new() -> Self {
        Self {
            latin: None,
            east_asia: None,
            bidi: None,
        }
    }

    /// Set the Latin (`w:val`) theme language.
    pub fn set_latin(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "Latin theme font")?;
        }
        self.latin = value;
        Ok(self)
    }

    /// Set the East Asian (`w:eastAsia`) theme language.
    pub fn set_east_asia(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "East Asian theme font")?;
        }
        self.east_asia = value;
        Ok(self)
    }

    /// Set the complex-script (`w:bidi`) theme language.
    pub fn set_bidi(&mut self, value: Option<String>) -> Result<&mut Self> {
        if let Some(tag) = value.as_deref() {
            validate_language_tag(tag, "complex-script theme font")?;
        }
        self.bidi = value;
        Ok(self)
    }

    /// Return the Latin theme language, when specified.
    #[inline]
    pub fn latin(&self) -> Option<&str> {
        self.latin.as_deref()
    }

    /// Return the East Asian theme language, when specified.
    #[inline]
    pub fn east_asia(&self) -> Option<&str> {
        self.east_asia.as_deref()
    }

    /// Return the complex-script theme language, when specified.
    #[inline]
    pub fn bidi(&self) -> Option<&str> {
        self.bidi.as_deref()
    }

    /// Serialize a standalone `w:themeFontLang` fragment.
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:themeFontLang");
        for (name, value) in [
            ("val", &self.latin),
            ("eastAsia", &self.east_asia),
            ("bidi", &self.bidi),
        ] {
            if let Some(tag) = value {
                xml.push_str(&format!(" {prefix}:{name}=\""));
                escape_attribute(&mut xml, tag);
                xml.push('"');
            }
        }
        xml.push_str("/>");
        xml
    }
}
