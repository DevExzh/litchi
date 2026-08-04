//! Cell hyperlink (`text:a`) data structures for ODS spreadsheets.
//!
//! ODF stores spreadsheet hyperlinks as `text:a` elements inside the
//! paragraph content of a table cell (ODF 1.3 §6.1.8). The target IRI is
//! carried by the mandatory `xlink:href` attribute, while the visible link
//! text is the character content of the element.

use litchi_core::{Error, Result, xml::escape_xml};
use std::ops::Range;

/// Window behavior requested by an inert cell hyperlink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyperlinkShow {
    New,
    Replace,
}

impl HyperlinkShow {
    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::New => "new",
            Self::Replace => "replace",
        }
    }

    pub(crate) fn parse(value: &str) -> Option<Self> {
        match value {
            "new" => Some(Self::New),
            "replace" => Some(Self::Replace),
            _ => None,
        }
    }
}

/// Activation behavior requested by an inert cell hyperlink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyperlinkActuate {
    OnRequest,
}

impl HyperlinkActuate {
    pub(crate) const fn as_str(self) -> &'static str {
        "onRequest"
    }
    pub(crate) fn parse(value: &str) -> Option<Self> {
        (value == "onRequest").then_some(Self::OnRequest)
    }
}

/// An inert hyperlink represented by a `text:a` element inside cell content.
///
/// The hyperlink is inert metadata: the target IRI is preserved verbatim and
/// is never dereferenced by this crate.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellHyperlink {
    /// Target IRI from the mandatory `xlink:href` attribute.
    pub href: String,
    /// Plain text of the link (character content of the `text:a` subtree).
    pub text: String,
    /// UTF-8 byte range occupied by this anchor in its parent cell's text.
    ///
    /// The range is assigned by a cell when the hyperlink is parsed or
    /// authored. It remains private so it cannot become detached from the
    /// cell text without the cell's validation path.
    pub(crate) range: Range<usize>,
    /// Optional `office:name` attribute naming the hyperlink.
    pub name: Option<String>,
    /// Optional `office:title` attribute with a short accessible title.
    pub title: Option<String>,
    /// Optional `office:target-frame-name` attribute.
    pub target_frame_name: Option<String>,
    /// Optional `xlink:show` behavior.
    pub show: Option<HyperlinkShow>,
    /// Optional explicit `xlink:actuate` behavior.
    pub actuate: Option<HyperlinkActuate>,
    /// Optional `text:style-name` applied to the unvisited link.
    pub style_name: Option<String>,
    /// Optional `text:visited-style-name` applied to the visited link.
    pub visited_style_name: Option<String>,
}

impl CellHyperlink {
    /// Create a hyperlink with the mandatory target IRI and empty metadata.
    pub fn new(href: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            text: String::new(),
            range: 0..0,
            name: None,
            title: None,
            target_frame_name: None,
            show: None,
            actuate: None,
            style_name: None,
            visited_style_name: None,
        }
    }

    /// Create a validated simple hyperlink with visible text.
    pub fn with_text(href: impl Into<String>, text: impl Into<String>) -> Result<Self> {
        let mut hyperlink = Self::new(href);
        hyperlink.text = text.into();
        hyperlink.range = 0..hyperlink.text.len();
        hyperlink.validate()?;
        Ok(hyperlink)
    }

    /// The target IRI of the hyperlink.
    pub fn href(&self) -> &str {
        &self.href
    }

    /// The visible plain-text content of the hyperlink.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return the UTF-8 byte range this anchor occupies in its parent cell.
    ///
    /// A range can be empty for a zero-width `text:a` anchor. The range is
    /// meaningful together with [`Cell::text`](super::Cell::text), not as an
    /// offset into the hyperlink's own text.
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    pub(crate) fn set_range(&mut self, range: Range<usize>) {
        self.range = range;
    }

    /// Validate data before it is serialized as an ODF `text:a` element.
    pub fn validate(&self) -> Result<()> {
        validate_href(&self.href)?;
        validate_xml_string(&self.text, "cell hyperlink text")?;
        for (value, label) in [
            (self.name.as_deref(), "cell hyperlink name"),
            (self.title.as_deref(), "cell hyperlink title"),
            (
                self.target_frame_name.as_deref(),
                "cell hyperlink target frame name",
            ),
            (self.style_name.as_deref(), "cell hyperlink style name"),
            (
                self.visited_style_name.as_deref(),
                "cell hyperlink visited style name",
            ),
        ] {
            if let Some(value) = value {
                validate_xml_string(value, label)?;
            }
        }
        Ok(())
    }

    /// Serialize this validated hyperlink as inert ODF inline content.
    pub(crate) fn write_xml(&self, output: &mut String) {
        output.push_str("<text:a xlink:type=\"simple\" xlink:href=\"");
        output.push_str(&escape_xml(&self.href));
        output.push('"');
        if let Some(actuate) = self.actuate {
            output.push_str(" xlink:actuate=\"");
            output.push_str(actuate.as_str());
            output.push('"');
        }
        if let Some(target_frame_name) = &self.target_frame_name {
            write_attribute(output, "office:target-frame-name", target_frame_name);
        }
        if let Some(show) = self.show {
            output.push_str(" xlink:show=\"");
            output.push_str(show.as_str());
            output.push('"');
        }
        if let Some(name) = &self.name {
            write_attribute(output, "office:name", name);
        }
        if let Some(title) = &self.title {
            write_attribute(output, "office:title", title);
        }
        if let Some(style_name) = &self.style_name {
            write_attribute(output, "text:style-name", style_name);
        }
        if let Some(visited_style_name) = &self.visited_style_name {
            write_attribute(output, "text:visited-style-name", visited_style_name);
        }
        output.push('>');
        output.push_str(&escape_xml(&self.text));
        output.push_str("</text:a>");
    }
}

fn write_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

fn validate_href(href: &str) -> Result<()> {
    if href.is_empty() {
        return Err(Error::InvalidFormat(
            "cell hyperlink href must not be empty".to_string(),
        ));
    }
    if href.chars().any(|character| character.is_control()) {
        return Err(Error::InvalidFormat(
            "cell hyperlink href must not contain control characters".to_string(),
        ));
    }
    Ok(())
}

fn validate_xml_string(value: &str, label: &str) -> Result<()> {
    if value
        .chars()
        .any(|character| matches!(character, '\0'..='\x08' | '\x0B'..='\x0C' | '\x0E'..='\x1F'))
    {
        return Err(Error::InvalidFormat(format!(
            "{label} must not contain XML control characters"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn new_hyperlink_keeps_href_and_defaults_metadata() {
        let link = CellHyperlink::new("https://example.com/");
        assert_eq!(link.href(), "https://example.com/");
        assert_eq!(link.text(), "");
        assert_eq!(link.range(), 0..0);
        assert!(link.name.is_none());
        assert!(link.title.is_none());
        assert!(link.target_frame_name.is_none());
        assert!(link.show.is_none());
        assert!(link.actuate.is_none());
        assert!(link.style_name.is_none());
        assert!(link.visited_style_name.is_none());
    }

    #[test]
    fn authoring_validates_and_serializes_complete_hyperlink_metadata() {
        let mut link = CellHyperlink::with_text("https://example.test/a?x=1&y=2", "A & B").unwrap();
        link.name = Some("example-link".to_string());
        link.title = Some("Example & more".to_string());
        link.target_frame_name = Some("_blank".to_string());
        link.show = Some(HyperlinkShow::New);
        link.actuate = Some(HyperlinkActuate::OnRequest);
        link.style_name = Some("Internet_20_link".to_string());
        link.visited_style_name = Some("Visited_20_Internet_20_link".to_string());
        link.validate().unwrap();

        let mut xml = String::new();
        link.write_xml(&mut xml);
        assert!(xml.contains("xlink:type=\"simple\""));
        assert!(xml.contains("xlink:href=\"https://example.test/a?x=1&amp;y=2\""));
        assert!(xml.contains("xlink:show=\"new\""));
        assert!(xml.contains("xlink:actuate=\"onRequest\""));
        assert!(xml.contains("office:title=\"Example &amp; more\""));
        assert!(xml.ends_with(">A &amp; B</text:a>"));

        assert_eq!(
            CellHyperlink::with_text("https://example.test/", "A & B")
                .unwrap()
                .range(),
            0..5
        );
        assert!(CellHyperlink::with_text("", "missing target").is_err());
        assert!(CellHyperlink::with_text("https://example.test/\nnext", "unsafe").is_err());
    }
}
