//! Style elements for ODF documents.
//!
//! This module provides comprehensive support for ODF style definitions,
//! including parsing, inheritance, and property resolution.

use super::element::{Element, ElementBase};
use crate::common::Result;
use std::borrow::Cow;
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

impl std::str::FromStr for StyleFamily {
    type Err = String;
    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        match s {
            "paragraph" => Ok(Self::Paragraph),
            "text" => Ok(Self::Text),
            "table" => Ok(Self::Table),
            "table-column" => Ok(Self::TableColumn),
            "table-row" => Ok(Self::TableRow),
            "table-cell" => Ok(Self::TableCell),
            "page-layout" => Ok(Self::PageLayout),
            "master-page" => Ok(Self::MasterPage),
            "graphic" => Ok(Self::Graphic),
            _ => Err(format!("Invalid style family: {}", s)),
        }
    }
}

impl StyleFamily {
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
pub struct StyleProperties<'a> {
    /// Text properties
    pub text: TextProperties<'a>,
    /// Paragraph properties
    pub paragraph: ParagraphProperties<'a>,
    /// Table properties
    pub table: TableProperties<'a>,
    /// Graphic properties
    pub graphic: GraphicProperties<'a>,
}

/// Text/character style properties
#[derive(Debug, Clone, Default)]
pub struct TextProperties<'a> {
    pub font_name: Option<Cow<'a, str>>,
    pub font_size: Option<Cow<'a, str>>,
    pub font_weight: Option<Cow<'a, str>>,
    pub font_style: Option<Cow<'a, str>>,
    pub color: Option<Cow<'a, str>>,
    pub background_color: Option<Cow<'a, str>>,
    pub underline: Option<Cow<'a, str>>,
    pub strikethrough: Option<Cow<'a, str>>,
    pub text_shadow: Option<Cow<'a, str>>,
}

/// Paragraph style properties
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties<'a> {
    pub margin_left: Option<Cow<'a, str>>,
    pub margin_right: Option<Cow<'a, str>>,
    pub margin_top: Option<Cow<'a, str>>,
    pub margin_bottom: Option<Cow<'a, str>>,
    pub text_align: Option<Cow<'a, str>>,
    pub line_height: Option<Cow<'a, str>>,
    pub background_color: Option<Cow<'a, str>>,
    pub border: Option<Cow<'a, str>>,
}

/// Table style properties
#[derive(Debug, Clone, Default)]
pub struct TableProperties<'a> {
    pub width: Option<Cow<'a, str>>,
    pub background_color: Option<Cow<'a, str>>,
    pub border: Option<Cow<'a, str>>,
    pub align: Option<Cow<'a, str>>,
}

/// Graphic style properties
#[derive(Debug, Clone, Default)]
pub struct GraphicProperties<'a> {
    pub background_color: Option<Cow<'a, str>>,
    pub border: Option<Cow<'a, str>>,
    pub shadow: Option<Cow<'a, str>>,
}

/// A style definition element
#[derive(Debug, Clone)]
pub struct Style {
    element: Element,
    properties: StyleProperties<'static>,
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

    /// Create a new style with name and family
    ///
    /// # Arguments
    ///
    /// * `name` - Name of the style
    /// * `family` - Style family (e.g., "text", "paragraph", "table")
    pub fn with_name_and_family(name: &str, family: &str) -> Self {
        let mut element = Element::new("style:style");
        element.set_attribute("style:name", name);
        element.set_attribute("style:family", family);
        Self {
            element,
            properties: StyleProperties::default(),
        }
    }

    /// Set a text property
    ///
    /// # Arguments
    ///
    /// * `property` - Property name (e.g., "fo:font-size", "fo:font-weight")
    /// * `value` - Property value
    pub fn set_text_property(&mut self, property: &str, value: &str) {
        // Create or update text-properties element
        let mut found = false;
        for child in &mut self.element.children {
            if child.tag_name() == "style:text-properties" {
                child.set_attribute(property, value);
                found = true;
                break;
            }
        }

        if !found {
            let mut text_props = Element::new("style:text-properties");
            text_props.set_attribute(property, value);
            self.element.children.push(text_props);
        }
    }

    /// Set a paragraph property
    ///
    /// # Arguments
    ///
    /// * `property` - Property name (e.g., "fo:text-align", "fo:margin-top")
    /// * `value` - Property value
    pub fn set_paragraph_property(&mut self, property: &str, value: &str) {
        // Create or update paragraph-properties element
        let mut found = false;
        for child in &mut self.element.children {
            if child.tag_name() == "style:paragraph-properties" {
                child.set_attribute(property, value);
                found = true;
                break;
            }
        }

        if !found {
            let mut para_props = Element::new("style:paragraph-properties");
            para_props.set_attribute(property, value);
            self.element.children.push(para_props);
        }
    }

    /// Set a table property
    ///
    /// # Arguments
    ///
    /// * `property` - Property name
    /// * `value` - Property value
    pub fn set_table_property(&mut self, property: &str, value: &str) {
        let mut found = false;
        for child in &mut self.element.children {
            if child.tag_name() == "style:table-properties" {
                child.set_attribute(property, value);
                found = true;
                break;
            }
        }

        if !found {
            let mut table_props = Element::new("style:table-properties");
            table_props.set_attribute(property, value);
            self.element.children.push(table_props);
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
            self.properties.text = Self::parse_text_properties(text_prop_elem);
        }

        // Parse paragraph properties
        if let Some(para_prop_elem) = self.find_property_element("style:paragraph-properties") {
            self.properties.paragraph = Self::parse_paragraph_properties(para_prop_elem);
        }

        // Parse table properties
        if let Some(table_prop_elem) = self.find_property_element("style:table-properties") {
            self.properties.table = Self::parse_table_properties(table_prop_elem);
        }

        // Parse graphic properties
        if let Some(graphic_prop_elem) = self.find_property_element("style:graphic-properties") {
            self.properties.graphic = Self::parse_graphic_properties(graphic_prop_elem);
        }

        Ok(())
    }

    /// Find a property element by tag name, returning a reference
    fn find_property_element(&self, tag_name: &str) -> Option<&Element> {
        self.element
            .children
            .iter()
            .find(|child| child.tag_name() == tag_name)
    }

    /// Parse text properties from element
    fn parse_text_properties(element: &Element) -> TextProperties<'static> {
        TextProperties {
            font_name: element
                .get_attribute("style:font-name")
                .map(|s| Cow::Owned(s.to_string())),
            font_size: element
                .get_attribute("fo:font-size")
                .map(|s| Cow::Owned(s.to_string())),
            font_weight: element
                .get_attribute("fo:font-weight")
                .map(|s| Cow::Owned(s.to_string())),
            font_style: element
                .get_attribute("fo:font-style")
                .map(|s| Cow::Owned(s.to_string())),
            color: element
                .get_attribute("fo:color")
                .map(|s| Cow::Owned(s.to_string())),
            background_color: element
                .get_attribute("fo:background-color")
                .map(|s| Cow::Owned(s.to_string())),
            underline: element
                .get_attribute("style:text-underline-style")
                .map(|s| Cow::Owned(s.to_string())),
            strikethrough: element
                .get_attribute("style:text-line-through-style")
                .map(|s| Cow::Owned(s.to_string())),
            text_shadow: element
                .get_attribute("fo:text-shadow")
                .map(|s| Cow::Owned(s.to_string())),
        }
    }

    /// Parse paragraph properties from element
    fn parse_paragraph_properties(element: &Element) -> ParagraphProperties<'static> {
        ParagraphProperties {
            margin_left: element
                .get_attribute("fo:margin-left")
                .map(|s| Cow::Owned(s.to_string())),
            margin_right: element
                .get_attribute("fo:margin-right")
                .map(|s| Cow::Owned(s.to_string())),
            margin_top: element
                .get_attribute("fo:margin-top")
                .map(|s| Cow::Owned(s.to_string())),
            margin_bottom: element
                .get_attribute("fo:margin-bottom")
                .map(|s| Cow::Owned(s.to_string())),
            text_align: element
                .get_attribute("fo:text-align")
                .map(|s| Cow::Owned(s.to_string())),
            line_height: element
                .get_attribute("fo:line-height")
                .map(|s| Cow::Owned(s.to_string())),
            background_color: element
                .get_attribute("fo:background-color")
                .map(|s| Cow::Owned(s.to_string())),
            border: element
                .get_attribute("fo:border")
                .map(|s| Cow::Owned(s.to_string())),
        }
    }

    /// Parse table properties from element
    fn parse_table_properties(element: &Element) -> TableProperties<'static> {
        TableProperties {
            width: element
                .get_attribute("style:width")
                .map(|s| Cow::Owned(s.to_string())),
            background_color: element
                .get_attribute("fo:background-color")
                .map(|s| Cow::Owned(s.to_string())),
            border: element
                .get_attribute("fo:border")
                .map(|s| Cow::Owned(s.to_string())),
            align: element
                .get_attribute("table:align")
                .map(|s| Cow::Owned(s.to_string())),
        }
    }

    /// Parse graphic properties from element
    fn parse_graphic_properties(element: &Element) -> GraphicProperties<'static> {
        GraphicProperties {
            background_color: element
                .get_attribute("draw:fill-color")
                .map(|s| Cow::Owned(s.to_string())),
            border: element
                .get_attribute("draw:stroke")
                .map(|s| Cow::Owned(s.to_string())),
            shadow: element
                .get_attribute("draw:shadow")
                .map(|s| Cow::Owned(s.to_string())),
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
            .and_then(|s| s.parse::<StyleFamily>().ok())
    }

    /// Get the parent style name
    pub fn parent_style_name(&self) -> Option<&str> {
        self.element.get_attribute("style:parent-style-name")
    }

    /// Get style properties
    pub fn properties(&self) -> &StyleProperties<'static> {
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
    pub fn get_resolved_properties(&self, style_name: &str) -> StyleProperties<'static> {
        let mut resolved = StyleProperties::default();

        // Walk up the inheritance chain
        let mut current_name = Some(style_name);
        while let Some(name) = current_name {
            if let Some(style) = self.styles.get(name) {
                // Merge properties (child overrides parent)
                Self::merge_properties(&mut resolved, &style.properties);
                current_name = style.parent_style_name();
            } else {
                break;
            }
        }

        resolved
    }

    /// Merge source properties into target (source takes precedence)
    ///
    /// Uses a macro to reduce boilerplate while maintaining zero-copy semantics
    /// where possible. The clone is necessary here because we're merging from
    /// a reference into a mutable target.
    fn merge_properties(target: &mut StyleProperties<'static>, source: &StyleProperties<'static>) {
        macro_rules! merge_prop {
            ($target_field:expr, $source_field:expr) => {
                if $source_field.is_some() {
                    $target_field = $source_field.clone();
                }
            };
        }

        // Merge text properties
        merge_prop!(target.text.font_name, source.text.font_name);
        merge_prop!(target.text.font_size, source.text.font_size);
        merge_prop!(target.text.font_weight, source.text.font_weight);
        merge_prop!(target.text.font_style, source.text.font_style);
        merge_prop!(target.text.color, source.text.color);
        merge_prop!(target.text.background_color, source.text.background_color);
        merge_prop!(target.text.underline, source.text.underline);
        merge_prop!(target.text.strikethrough, source.text.strikethrough);
        merge_prop!(target.text.text_shadow, source.text.text_shadow);

        // Merge paragraph properties
        merge_prop!(target.paragraph.margin_left, source.paragraph.margin_left);
        merge_prop!(target.paragraph.margin_right, source.paragraph.margin_right);
        merge_prop!(target.paragraph.margin_top, source.paragraph.margin_top);
        merge_prop!(
            target.paragraph.margin_bottom,
            source.paragraph.margin_bottom
        );
        merge_prop!(target.paragraph.text_align, source.paragraph.text_align);
        merge_prop!(target.paragraph.line_height, source.paragraph.line_height);
        merge_prop!(
            target.paragraph.background_color,
            source.paragraph.background_color
        );
        merge_prop!(target.paragraph.border, source.paragraph.border);

        // Merge table properties
        merge_prop!(target.table.width, source.table.width);
        merge_prop!(target.table.background_color, source.table.background_color);
        merge_prop!(target.table.border, source.table.border);
        merge_prop!(target.table.align, source.table.align);

        // Merge graphic properties
        merge_prop!(
            target.graphic.background_color,
            source.graphic.background_color
        );
        merge_prop!(target.graphic.border, source.graphic.border);
        merge_prop!(target.graphic.shadow, source.graphic.shadow);
    }

    /// Parse styles from XML content
    pub fn from_xml(xml_content: &str) -> Result<Self> {
        let mut registry = Self::default();

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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_style_family_from_str() {
        assert_eq!(
            "paragraph".parse::<StyleFamily>().unwrap(),
            StyleFamily::Paragraph
        );
        assert_eq!("text".parse::<StyleFamily>().unwrap(), StyleFamily::Text);
        assert_eq!("table".parse::<StyleFamily>().unwrap(), StyleFamily::Table);
        assert_eq!(
            "table-column".parse::<StyleFamily>().unwrap(),
            StyleFamily::TableColumn
        );
        assert_eq!(
            "table-row".parse::<StyleFamily>().unwrap(),
            StyleFamily::TableRow
        );
        assert_eq!(
            "table-cell".parse::<StyleFamily>().unwrap(),
            StyleFamily::TableCell
        );
        assert_eq!(
            "page-layout".parse::<StyleFamily>().unwrap(),
            StyleFamily::PageLayout
        );
        assert_eq!(
            "master-page".parse::<StyleFamily>().unwrap(),
            StyleFamily::MasterPage
        );
        assert_eq!(
            "graphic".parse::<StyleFamily>().unwrap(),
            StyleFamily::Graphic
        );

        // Invalid family
        assert!("invalid".parse::<StyleFamily>().is_err());
    }

    #[test]
    fn test_style_family_as_str() {
        assert_eq!(StyleFamily::Paragraph.as_str(), "paragraph");
        assert_eq!(StyleFamily::Text.as_str(), "text");
        assert_eq!(StyleFamily::Table.as_str(), "table");
        assert_eq!(StyleFamily::TableColumn.as_str(), "table-column");
        assert_eq!(StyleFamily::TableRow.as_str(), "table-row");
        assert_eq!(StyleFamily::TableCell.as_str(), "table-cell");
        assert_eq!(StyleFamily::PageLayout.as_str(), "page-layout");
        assert_eq!(StyleFamily::MasterPage.as_str(), "master-page");
        assert_eq!(StyleFamily::Graphic.as_str(), "graphic");
    }

    #[test]
    fn test_style_family_roundtrip() {
        let families = [
            StyleFamily::Paragraph,
            StyleFamily::Text,
            StyleFamily::Table,
            StyleFamily::TableColumn,
            StyleFamily::TableRow,
            StyleFamily::TableCell,
            StyleFamily::PageLayout,
            StyleFamily::MasterPage,
            StyleFamily::Graphic,
        ];

        for family in &families {
            let s = family.as_str();
            let parsed: StyleFamily = s.parse().unwrap();
            assert_eq!(*family, parsed);
        }
    }

    #[test]
    fn test_style_new() {
        let style = Style::new();
        assert!(style.name().is_none());
        assert!(style.family().is_none());
        assert!(style.parent_style_name().is_none());
        assert!(!style.is_default());
    }

    #[test]
    fn test_style_with_name_and_family() {
        let style = Style::with_name_and_family("Heading1", "paragraph");
        assert_eq!(style.name(), Some("Heading1"));
        assert_eq!(style.family(), Some(StyleFamily::Paragraph));
    }

    #[test]
    fn test_style_set_text_property() {
        let mut style = Style::new();
        style.set_text_property("fo:font-size", "12pt");

        // Property should be stored in element
        let text_props = style
            .element
            .children
            .iter()
            .find(|c| c.tag_name() == "style:text-properties");
        assert!(text_props.is_some());
        assert_eq!(
            text_props.unwrap().get_attribute("fo:font-size"),
            Some("12pt")
        );
    }

    #[test]
    fn test_style_set_paragraph_property() {
        let mut style = Style::new();
        style.set_paragraph_property("fo:text-align", "center");

        let para_props = style
            .element
            .children
            .iter()
            .find(|c| c.tag_name() == "style:paragraph-properties");
        assert!(para_props.is_some());
        assert_eq!(
            para_props.unwrap().get_attribute("fo:text-align"),
            Some("center")
        );
    }

    #[test]
    fn test_style_set_table_property() {
        let mut style = Style::new();
        style.set_table_property("style:width", "10cm");

        let table_props = style
            .element
            .children
            .iter()
            .find(|c| c.tag_name() == "style:table-properties");
        assert!(table_props.is_some());
        assert_eq!(
            table_props.unwrap().get_attribute("style:width"),
            Some("10cm")
        );
    }

    #[test]
    fn test_style_registry_default() {
        let registry = StyleRegistry::default();
        assert!(registry.styles.is_empty());
    }

    #[test]
    fn test_style_registry_add_style() {
        let mut registry = StyleRegistry::default();
        let style = Style::with_name_and_family("TestStyle", "text");
        registry.add_style(style);

        assert_eq!(registry.styles.len(), 1);
        assert!(registry.get_style("TestStyle").is_some());
    }

    #[test]
    fn test_style_registry_get_style() {
        let mut registry = StyleRegistry::default();
        let style = Style::with_name_and_family("MyStyle", "paragraph");
        registry.add_style(style);

        let retrieved = registry.get_style("MyStyle").unwrap();
        assert_eq!(retrieved.name(), Some("MyStyle"));

        assert!(registry.get_style("NonExistent").is_none());
    }

    #[test]
    fn test_style_registry_get_resolved_properties() {
        let mut registry = StyleRegistry::default();

        // Create parent style with properties set directly
        let mut parent = Style::with_name_and_family("Parent", "text");
        parent.properties.text.font_size = Some(Cow::Owned("12pt".to_string()));
        registry.add_style(parent);

        // Create child style with parent reference
        let mut child = Style::with_name_and_family("Child", "text");
        child
            .element
            .set_attribute("style:parent-style-name", "Parent");
        child.properties.text.color = Some(Cow::Owned("#ff0000".to_string()));
        registry.add_style(child);

        let resolved = registry.get_resolved_properties("Child");
        // Child properties take precedence
        assert_eq!(resolved.text.color.as_deref(), Some("#ff0000"));
        // Parent properties are inherited
        assert_eq!(resolved.text.font_size.as_deref(), Some("12pt"));
    }

    #[test]
    fn test_style_registry_from_xml() {
        // Note: The current implementation parses style attributes but not nested properties
        let xml = r#"
        <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                               xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
            <style:style style:name="Heading1" style:family="paragraph"></style:style>
            <style:style style:name="TextBody" style:family="paragraph"></style:style>
        </office:document-styles>"#;

        let registry = StyleRegistry::from_xml(xml).unwrap();
        assert_eq!(registry.styles.len(), 2);
        assert!(registry.get_style("Heading1").is_some());
        assert!(registry.get_style("TextBody").is_some());
    }

    #[test]
    fn test_style_elements_parse_styles() {
        let xml = r#"
        <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                               xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
            <style:style style:name="CustomStyle" style:family="text"></style:style>
        </office:document-styles>"#;

        let registry = StyleElements::parse_styles(xml).unwrap();
        assert_eq!(registry.styles.len(), 1);
        assert!(registry.get_style("CustomStyle").is_some());
    }

    #[test]
    fn test_style_properties_default() {
        let props = StyleProperties::default();
        assert!(props.text.font_name.is_none());
        assert!(props.text.font_size.is_none());
        assert!(props.paragraph.text_align.is_none());
        assert!(props.table.width.is_none());
        assert!(props.graphic.background_color.is_none());
    }

    #[test]
    fn test_style_is_default() {
        let style = Style::new();
        assert!(!style.is_default());

        let style = Style::with_name_and_family("", "paragraph");
        assert!(style.is_default());
    }

    #[test]
    fn test_style_into_element() {
        let style = Style::with_name_and_family("Test", "text");
        let element: Element = style.into();
        assert_eq!(element.tag_name(), "style:style");
        assert_eq!(element.get_attribute("style:name"), Some("Test"));
    }
}
