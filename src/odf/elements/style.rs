//! Style elements for ODF documents.
//!
//! This module provides comprehensive support for ODF style definitions,
//! including parsing, inheritance, and property resolution.

use super::element::{Element, ElementBase};
use crate::common::Result;
use std::collections::HashMap;

/// Style family types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleFamily {
    /// Paragraph styles
    Paragraph,
    /// Text/character styles
    Text,
    /// Table styles
    Table,
    /// Table column styles
    TableColumn,
    /// Table row styles
    TableRow,
    /// Table cell styles
    TableCell,
    /// Page layout styles
    PageLayout,
    /// Master page styles
    MasterPage,
    /// Graphic styles
    Graphic,
}

impl StyleFamily {
    /// Parse style family from string
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "paragraph" => Some(Self::Paragraph),
            "text" => Some(Self::Text),
            "table" => Some(Self::Table),
            "table-column" => Some(Self::TableColumn),
            "table-row" => Some(Self::TableRow),
            "table-cell" => Some(Self::TableCell),
            "page-layout" => Some(Self::PageLayout),
            "master-page" => Some(Self::MasterPage),
            "graphic" => Some(Self::Graphic),
            _ => None,
        }
    }

    /// Convert to string
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Paragraph => "paragraph",
            Self::Text => "text",
            Self::Table => "table",
            Self::TableColumn => "table-column",
            Self::TableRow => "table-row",
            Self::TableCell => "table-cell",
            Self::PageLayout => "page-layout",
            Self::MasterPage => "master-page",
            Self::Graphic => "graphic",
        }
    }
}

/// Style properties container
#[derive(Debug, Clone, Default)]
pub struct StyleProperties {
    /// Text properties
    pub text: TextProperties,
    /// Paragraph properties
    pub paragraph: ParagraphProperties,
    /// Table properties
    pub table: TableProperties,
    /// Graphic properties
    pub graphic: GraphicProperties,
}

/// Text/character style properties
#[derive(Debug, Clone, Default)]
pub struct TextProperties {
    pub font_name: Option<String>,
    pub font_size: Option<String>,
    pub font_weight: Option<String>,
    pub font_style: Option<String>,
    pub color: Option<String>,
    pub background_color: Option<String>,
    pub underline: Option<String>,
    pub strikethrough: Option<String>,
    pub text_shadow: Option<String>,
}

/// Paragraph style properties
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    pub margin_left: Option<String>,
    pub margin_right: Option<String>,
    pub margin_top: Option<String>,
    pub margin_bottom: Option<String>,
    pub text_align: Option<String>,
    pub line_height: Option<String>,
    pub background_color: Option<String>,
    pub border: Option<String>,
}

/// Table style properties
#[derive(Debug, Clone, Default)]
pub struct TableProperties {
    pub width: Option<String>,
    pub background_color: Option<String>,
    pub border: Option<String>,
    pub align: Option<String>,
}

/// Graphic style properties
#[derive(Debug, Clone, Default)]
pub struct GraphicProperties {
    pub background_color: Option<String>,
    pub border: Option<String>,
    pub shadow: Option<String>,
}

/// A style definition element
#[derive(Debug, Clone)]
pub struct Style {
    element: Element,
    properties: StyleProperties,
}

impl Default for Style {
    fn default() -> Self {
        Self::new()
    }
}

impl Style {
    /// Create a new style
    pub fn new() -> Self {
        Self {
            element: Element::new("style:style"),
            properties: StyleProperties::default(),
        }
    }

    /// Create style from element and parse properties
    pub fn from_element(element: Element) -> Result<Self> {
        let mut style = Self {
            element,
            properties: StyleProperties::default(),
        };
        style.parse_properties()?;
        Ok(style)
    }

    /// Parse style properties from the element
    fn parse_properties(&mut self) -> Result<()> {
        // Parse text properties
        if let Some(text_prop_elem) = self.find_property_element("style:text-properties") {
            self.properties.text = Self::parse_text_properties(&text_prop_elem);
        }

        // Parse paragraph properties
        if let Some(para_prop_elem) = self.find_property_element("style:paragraph-properties") {
            self.properties.paragraph = Self::parse_paragraph_properties(&para_prop_elem);
        }

        // Parse table properties
        if let Some(table_prop_elem) = self.find_property_element("style:table-properties") {
            self.properties.table = Self::parse_table_properties(&table_prop_elem);
        }

        // Parse graphic properties
        if let Some(graphic_prop_elem) = self.find_property_element("style:graphic-properties") {
            self.properties.graphic = Self::parse_graphic_properties(&graphic_prop_elem);
        }

        Ok(())
    }

    /// Find a property element by tag name
    fn find_property_element(&self, tag_name: &str) -> Option<Element> {
        for child in self.element.children() {
            if child.tag_name() == tag_name {
                return Some(unsafe { &*(child as *const _ as *const Element) }.clone());
            }
        }
        None
    }

    /// Parse text properties from element
    fn parse_text_properties(element: &Element) -> TextProperties {
        TextProperties {
            font_name: element
                .get_attribute("style:font-name")
                .map(|s| s.to_string()),
            font_size: element.get_attribute("fo:font-size").map(|s| s.to_string()),
            font_weight: element
                .get_attribute("fo:font-weight")
                .map(|s| s.to_string()),
            font_style: element
                .get_attribute("fo:font-style")
                .map(|s| s.to_string()),
            color: element.get_attribute("fo:color").map(|s| s.to_string()),
            background_color: element
                .get_attribute("fo:background-color")
                .map(|s| s.to_string()),
            underline: element
                .get_attribute("style:text-underline-style")
                .map(|s| s.to_string()),
            strikethrough: element
                .get_attribute("style:text-line-through-style")
                .map(|s| s.to_string()),
            text_shadow: element
                .get_attribute("fo:text-shadow")
                .map(|s| s.to_string()),
        }
    }

    /// Parse paragraph properties from element
    fn parse_paragraph_properties(element: &Element) -> ParagraphProperties {
        ParagraphProperties {
            margin_left: element
                .get_attribute("fo:margin-left")
                .map(|s| s.to_string()),
            margin_right: element
                .get_attribute("fo:margin-right")
                .map(|s| s.to_string()),
            margin_top: element
                .get_attribute("fo:margin-top")
                .map(|s| s.to_string()),
            margin_bottom: element
                .get_attribute("fo:margin-bottom")
                .map(|s| s.to_string()),
            text_align: element
                .get_attribute("fo:text-align")
                .map(|s| s.to_string()),
            line_height: element
                .get_attribute("fo:line-height")
                .map(|s| s.to_string()),
            background_color: element
                .get_attribute("fo:background-color")
                .map(|s| s.to_string()),
            border: element.get_attribute("fo:border").map(|s| s.to_string()),
        }
    }

    /// Parse table properties from element
    fn parse_table_properties(element: &Element) -> TableProperties {
        TableProperties {
            width: element.get_attribute("style:width").map(|s| s.to_string()),
            background_color: element
                .get_attribute("fo:background-color")
                .map(|s| s.to_string()),
            border: element.get_attribute("fo:border").map(|s| s.to_string()),
            align: element.get_attribute("table:align").map(|s| s.to_string()),
        }
    }

    /// Parse graphic properties from element
    fn parse_graphic_properties(element: &Element) -> GraphicProperties {
        GraphicProperties {
            background_color: element
                .get_attribute("draw:fill-color")
                .map(|s| s.to_string()),
            border: element.get_attribute("draw:stroke").map(|s| s.to_string()),
            shadow: element.get_attribute("draw:shadow").map(|s| s.to_string()),
        }
    }

    /// Get the style name
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("style:name")
    }

    /// Get the style family
    pub fn family(&self) -> Option<StyleFamily> {
        self.element
            .get_attribute("style:family")
            .and_then(StyleFamily::from_str)
    }

    /// Get the parent style name
    pub fn parent_style_name(&self) -> Option<&str> {
        self.element.get_attribute("style:parent-style-name")
    }

    /// Get style properties
    pub fn properties(&self) -> &StyleProperties {
        &self.properties
    }

    /// Check if this style is a default style
    pub fn is_default(&self) -> bool {
        self.name() == Some("")
    }
}

impl From<Style> for Element {
    fn from(style: Style) -> Element {
        style.element
    }
}

/// Style registry for managing document styles
#[derive(Debug, Clone, Default)]
pub struct StyleRegistry {
    pub styles: HashMap<String, Style>,
}

impl StyleRegistry {
    /// Create a new style registry
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a style to the registry
    pub fn add_style(&mut self, style: Style) {
        if let Some(name) = style.name() {
            self.styles.insert(name.to_string(), style);
        }
    }

    /// Get a style by name
    pub fn get_style(&self, name: &str) -> Option<&Style> {
        self.styles.get(name)
    }

    /// Get resolved properties for a style (with inheritance)
    pub fn get_resolved_properties(&self, style_name: &str) -> StyleProperties {
        let mut resolved = StyleProperties::default();

        // Walk up the inheritance chain
        let mut current_name = Some(style_name.to_string());
        while let Some(name) = current_name {
            if let Some(style) = self.styles.get(&name) {
                // Merge properties (child overrides parent)
                Self::merge_properties(&mut resolved, &style.properties);
                current_name = style.parent_style_name().map(|s| s.to_string());
            } else {
                break;
            }
        }

        resolved
    }

    /// Merge source properties into target (source takes precedence)
    fn merge_properties(target: &mut StyleProperties, source: &StyleProperties) {
        // Merge text properties
        if source.text.font_name.is_some() {
            target.text.font_name = source.text.font_name.clone();
        }
        if source.text.font_size.is_some() {
            target.text.font_size = source.text.font_size.clone();
        }
        if source.text.font_weight.is_some() {
            target.text.font_weight = source.text.font_weight.clone();
        }
        if source.text.font_style.is_some() {
            target.text.font_style = source.text.font_style.clone();
        }
        if source.text.color.is_some() {
            target.text.color = source.text.color.clone();
        }
        if source.text.background_color.is_some() {
            target.text.background_color = source.text.background_color.clone();
        }
        if source.text.underline.is_some() {
            target.text.underline = source.text.underline.clone();
        }
        if source.text.strikethrough.is_some() {
            target.text.strikethrough = source.text.strikethrough.clone();
        }
        if source.text.text_shadow.is_some() {
            target.text.text_shadow = source.text.text_shadow.clone();
        }

        // Merge paragraph properties
        if source.paragraph.margin_left.is_some() {
            target.paragraph.margin_left = source.paragraph.margin_left.clone();
        }
        if source.paragraph.margin_right.is_some() {
            target.paragraph.margin_right = source.paragraph.margin_right.clone();
        }
        if source.paragraph.margin_top.is_some() {
            target.paragraph.margin_top = source.paragraph.margin_top.clone();
        }
        if source.paragraph.margin_bottom.is_some() {
            target.paragraph.margin_bottom = source.paragraph.margin_bottom.clone();
        }
        if source.paragraph.text_align.is_some() {
            target.paragraph.text_align = source.paragraph.text_align.clone();
        }
        if source.paragraph.line_height.is_some() {
            target.paragraph.line_height = source.paragraph.line_height.clone();
        }
        if source.paragraph.background_color.is_some() {
            target.paragraph.background_color = source.paragraph.background_color.clone();
        }
        if source.paragraph.border.is_some() {
            target.paragraph.border = source.paragraph.border.clone();
        }

        // Similar merging for table and graphic properties...
    }

    /// Parse styles from XML content
    pub fn from_xml(xml_content: &str) -> Result<Self> {
        let mut registry = Self::new();

        // For now, use a simple approach that just parses style attributes
        // Full property parsing can be added later
        let mut reader = quick_xml::Reader::from_str(xml_content);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let tag_name =
                        String::from_utf8(e.name().as_ref().to_vec()).unwrap_or_default();

                    if tag_name == "style:style" {
                        let mut element = Element::new("style:style");

                        // Parse attributes
                        for attr_result in e.attributes() {
                            if let Ok(attr) = attr_result
                                && let (Ok(key), Ok(value)) = (
                                    String::from_utf8(attr.key.as_ref().to_vec()),
                                    String::from_utf8(attr.value.to_vec()),
                                )
                            {
                                element.set_attribute(&key, &value);
                            }
                        }

                        // Create style from element
                        if let Ok(style) = Style::from_element(element) {
                            registry.add_style(style);
                        }
                    }
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(registry)
    }
}

/// Collection of style elements for easy parsing
pub struct StyleElements;

impl StyleElements {
    /// Parse all styles from XML content
    pub fn parse_styles(xml_content: &str) -> Result<StyleRegistry> {
        StyleRegistry::from_xml(xml_content)
    }
}
