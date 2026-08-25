use crate::Rect;
use crate::error::{Result, invalid};

pub(crate) const MAX_HYPERLINK_TEXT_BYTES: usize = 16 * 1024 * 1024;

/// One checked worksheet cell or rectangular range retained with its source
/// A1 spelling.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct HyperlinkReference {
    lexical: Box<str>,
    range: Rect,
}

impl HyperlinkReference {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        validate_text(value, "hyperlink reference")?;
        let range =
            Rect::from_a1(value).map_err(|_error| invalid("invalid XLSX hyperlink reference"))?;
        Ok(Self {
            lexical: value.into(),
            range,
        })
    }

    /// Original bounded A1 spelling from the worksheet.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    /// Checked semantic worksheet range.
    #[must_use]
    pub const fn range(&self) -> Rect {
        self.range
    }
}

/// One inert worksheet hyperlink.
///
/// An external target is returned verbatim from the worksheet relationship;
/// it is never opened or resolved by Litchi.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    reference: HyperlinkReference,
    location: Option<Box<str>>,
    external_target: Option<Box<str>>,
    display: Option<Box<str>>,
    tooltip: Option<Box<str>>,
}

impl Hyperlink {
    pub(crate) fn from_parts(
        reference: HyperlinkReference,
        location: Option<String>,
        external_target: Option<String>,
        display: Option<String>,
        tooltip: Option<String>,
    ) -> Result<Self> {
        if location.as_deref().is_none_or(str::is_empty) && external_target.is_none() {
            return Err(invalid(
                "XLSX hyperlink requires an internal location or external target",
            ));
        }
        for (value, label) in [
            (location.as_deref(), "hyperlink location"),
            (external_target.as_deref(), "hyperlink external target"),
            (display.as_deref(), "hyperlink display text"),
            (tooltip.as_deref(), "hyperlink tooltip"),
        ] {
            if let Some(value) = value {
                validate_text(value, label)?;
            }
        }
        if external_target.as_deref() == Some("") {
            return Err(invalid("XLSX hyperlink external target cannot be empty"));
        }
        Ok(Self {
            reference,
            location: location.map(String::into_boxed_str),
            external_target: external_target.map(String::into_boxed_str),
            display: display.map(String::into_boxed_str),
            tooltip: tooltip.map(String::into_boxed_str),
        })
    }

    /// Checked worksheet cell or range carrying the hyperlink.
    #[must_use]
    pub const fn reference(&self) -> &HyperlinkReference {
        &self.reference
    }

    /// Inert internal workbook location, when present.
    #[must_use]
    pub fn location(&self) -> Option<&str> {
        self.location.as_deref()
    }

    /// Verbatim inert external relationship target, when present.
    #[must_use]
    pub fn external_target(&self) -> Option<&str> {
        self.external_target.as_deref()
    }

    /// Producer-authored display text, when present.
    #[must_use]
    pub fn display(&self) -> Option<&str> {
        self.display.as_deref()
    }

    /// Producer-authored tooltip, when present.
    #[must_use]
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }
}

pub(crate) fn validate_text(value: &str, label: &str) -> Result<()> {
    if value.len() > MAX_HYPERLINK_TEXT_BYTES {
        return Err(invalid(format!(
            "XLSX {label} exceeds the {MAX_HYPERLINK_TEXT_BYTES} byte safety limit"
        )));
    }
    if value
        .bytes()
        .any(|byte| byte < 0x20 && !matches!(byte, b'\t' | b'\n' | b'\r'))
    {
        return Err(invalid(format!(
            "XLSX {label} contains an XML-forbidden control character"
        )));
    }
    Ok(())
}
