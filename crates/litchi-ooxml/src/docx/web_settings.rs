//! WordprocessingML web-settings support.

use crate::docx::namespace::{
    is_wordprocessing_namespace, normalize_xml_integer, word_attribute_value,
};
use crate::error::{OoxmlError, Result};
use litchi_core::xml::escape_xml;
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as _;

/// Scalar settings from a Word `webSettings.xml` part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WebSettings {
    frameset: Option<Frameset>,
    divs: Option<Vec<HtmlDiv>>,
    encoding: Option<String>,
    optimize_for_browser: Option<bool>,
    rely_on_vml: Option<bool>,
    allow_png: Option<bool>,
    do_not_rely_on_css: Option<bool>,
    do_not_save_as_single_file: Option<bool>,
    do_not_organize_in_folder: Option<bool>,
    do_not_use_long_file_names: Option<bool>,
    pixels_per_inch: Option<String>,
    target_screen_size: Option<TargetScreenSize>,
    save_smart_tags_as_xml: Option<bool>,
}

/// Fidelity information for one HTML `div`, `body`, or `blockquote` element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HtmlDiv {
    id: String,
    block_quote: Option<bool>,
    body_div: Option<bool>,
    margin_left_twips: Option<String>,
    margin_right_twips: Option<String>,
    margin_top_twips: Option<String>,
    margin_bottom_twips: Option<String>,
    borders: Option<HtmlDivBorders>,
    children: Vec<HtmlDiv>,
}

/// Borders around an HTML division.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HtmlDivBorders {
    top: Option<HtmlDivBorder>,
    left: Option<HtmlDivBorder>,
    bottom: Option<HtmlDivBorder>,
    right: Option<HtmlDivBorder>,
}

/// One border around an HTML division.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HtmlDivBorder {
    style: String,
    color: Option<String>,
    theme_color: Option<ThemeColor>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
    size_eighth_points: Option<u64>,
    space_points: Option<u64>,
    shadow: Option<bool>,
    frame: Option<bool>,
}

/// A recursive web frameset definition.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Frameset {
    size: Option<String>,
    split_bar: Option<FramesetSplitBar>,
    layout: Option<FrameLayout>,
    children: Vec<FramesetChild>,
}

/// A child of a web frameset, retained in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FramesetChild {
    /// A nested frameset.
    Frameset(Frameset),
    /// A leaf frame.
    Frame(Frame),
}

/// Properties for one frame in a web frameset.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Frame {
    size: Option<String>,
    name: Option<String>,
    source_file_relationship_id: Option<String>,
    margin_width: Option<u64>,
    margin_height: Option<u64>,
    scrollbar: Option<FrameScrollbarVisibility>,
    no_resize_allowed: Option<bool>,
    linked_to_file: Option<bool>,
}

/// Visual properties for the splitter bars of a frameset.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FramesetSplitBar {
    width_twips: Option<u64>,
    color: Option<FramesetColor>,
    no_border: Option<bool>,
    flat_borders: Option<bool>,
}

/// A frameset splitter color with optional theme modifiers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FramesetColor {
    value: String,
    theme_color: Option<ThemeColor>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
}

/// The direction in which a frameset stacks its children.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameLayout {
    Rows,
    Columns,
    None,
}

/// The scrollbar policy for one frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameScrollbarVisibility {
    On,
    Off,
    Auto,
}

/// A WordprocessingML theme color.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThemeColor {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
    None,
    Background1,
    Text1,
    Background2,
    Text2,
}

/// A spec-defined target size for generated web pages.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TargetScreenSize {
    Pixels544x376,
    Pixels640x480,
    Pixels720x512,
    Pixels800x600,
    Pixels1024x768,
    Pixels1152x882,
    Pixels1152x900,
    Pixels1280x1024,
    Pixels1600x1200,
    Pixels1800x1440,
    Pixels1920x1200,
}

impl TargetScreenSize {
    fn from_xml(value: &str) -> Option<Self> {
        match value {
            "544x376" => Some(Self::Pixels544x376),
            "640x480" => Some(Self::Pixels640x480),
            "720x512" => Some(Self::Pixels720x512),
            "800x600" => Some(Self::Pixels800x600),
            "1024x768" => Some(Self::Pixels1024x768),
            "1152x882" => Some(Self::Pixels1152x882),
            "1152x900" => Some(Self::Pixels1152x900),
            "1280x1024" => Some(Self::Pixels1280x1024),
            "1600x1200" => Some(Self::Pixels1600x1200),
            "1800x1440" => Some(Self::Pixels1800x1440),
            "1920x1200" => Some(Self::Pixels1920x1200),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Pixels544x376 => "544x376",
            Self::Pixels640x480 => "640x480",
            Self::Pixels720x512 => "720x512",
            Self::Pixels800x600 => "800x600",
            Self::Pixels1024x768 => "1024x768",
            Self::Pixels1152x882 => "1152x882",
            Self::Pixels1152x900 => "1152x900",
            Self::Pixels1280x1024 => "1280x1024",
            Self::Pixels1600x1200 => "1600x1200",
            Self::Pixels1800x1440 => "1800x1440",
            Self::Pixels1920x1200 => "1920x1200",
        }
    }
}

impl FrameLayout {
    fn from_xml(value: &str) -> Option<Self> {
        match value {
            "rows" => Some(Self::Rows),
            "cols" => Some(Self::Columns),
            "none" => Some(Self::None),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Rows => "rows",
            Self::Columns => "cols",
            Self::None => "none",
        }
    }
}

impl FrameScrollbarVisibility {
    fn from_xml(value: &str) -> Option<Self> {
        match value {
            "on" => Some(Self::On),
            "off" => Some(Self::Off),
            "auto" => Some(Self::Auto),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::On => "on",
            Self::Off => "off",
            Self::Auto => "auto",
        }
    }
}

impl ThemeColor {
    fn from_xml(value: &str) -> Option<Self> {
        match value {
            "dark1" => Some(Self::Dark1),
            "light1" => Some(Self::Light1),
            "dark2" => Some(Self::Dark2),
            "light2" => Some(Self::Light2),
            "accent1" => Some(Self::Accent1),
            "accent2" => Some(Self::Accent2),
            "accent3" => Some(Self::Accent3),
            "accent4" => Some(Self::Accent4),
            "accent5" => Some(Self::Accent5),
            "accent6" => Some(Self::Accent6),
            "hyperlink" => Some(Self::Hyperlink),
            "followedHyperlink" => Some(Self::FollowedHyperlink),
            "none" => Some(Self::None),
            "background1" => Some(Self::Background1),
            "text1" => Some(Self::Text1),
            "background2" => Some(Self::Background2),
            "text2" => Some(Self::Text2),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dark1",
            Self::Light1 => "light1",
            Self::Dark2 => "dark2",
            Self::Light2 => "light2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
            Self::None => "none",
            Self::Background1 => "background1",
            Self::Text1 => "text1",
            Self::Background2 => "background2",
            Self::Text2 => "text2",
        }
    }
}

impl Frameset {
    /// Return the size expression for this frameset, if explicitly present.
    pub fn size(&self) -> Option<&str> {
        self.size.as_deref()
    }

    /// Return splitter-bar properties, if explicitly present.
    pub fn split_bar(&self) -> Option<&FramesetSplitBar> {
        self.split_bar.as_ref()
    }

    /// Return the explicit child layout. Absence has the schema-defined row default.
    pub fn layout(&self) -> Option<FrameLayout> {
        self.layout
    }

    /// Return nested frames and framesets in document order.
    pub fn children(&self) -> &[FramesetChild] {
        &self.children
    }
}

impl Frame {
    pub fn size(&self) -> Option<&str> {
        self.size.as_deref()
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn source_file_relationship_id(&self) -> Option<&str> {
        self.source_file_relationship_id.as_deref()
    }

    pub fn margin_width(&self) -> Option<u64> {
        self.margin_width
    }

    pub fn margin_height(&self) -> Option<u64> {
        self.margin_height
    }

    pub fn scrollbar(&self) -> Option<FrameScrollbarVisibility> {
        self.scrollbar
    }

    pub fn no_resize_allowed(&self) -> Option<bool> {
        self.no_resize_allowed
    }

    pub fn linked_to_file(&self) -> Option<bool> {
        self.linked_to_file
    }
}

impl FramesetSplitBar {
    pub fn width_twips(&self) -> Option<u64> {
        self.width_twips
    }

    pub fn color(&self) -> Option<&FramesetColor> {
        self.color.as_ref()
    }

    pub fn no_border(&self) -> Option<bool> {
        self.no_border
    }

    pub fn flat_borders(&self) -> Option<bool> {
        self.flat_borders
    }
}

impl FramesetColor {
    pub fn value(&self) -> &str {
        &self.value
    }

    pub fn theme_color(&self) -> Option<ThemeColor> {
        self.theme_color
    }

    pub fn theme_tint(&self) -> Option<u8> {
        self.theme_tint
    }

    pub fn theme_shade(&self) -> Option<u8> {
        self.theme_shade
    }
}

impl HtmlDiv {
    pub fn id(&self) -> &str {
        &self.id
    }

    pub fn is_block_quote(&self) -> Option<bool> {
        self.block_quote
    }

    pub fn is_body_div(&self) -> Option<bool> {
        self.body_div
    }

    pub fn margin_left_twips(&self) -> Option<&str> {
        self.margin_left_twips.as_deref()
    }

    pub fn margin_right_twips(&self) -> Option<&str> {
        self.margin_right_twips.as_deref()
    }

    pub fn margin_top_twips(&self) -> Option<&str> {
        self.margin_top_twips.as_deref()
    }

    pub fn margin_bottom_twips(&self) -> Option<&str> {
        self.margin_bottom_twips.as_deref()
    }

    pub fn borders(&self) -> Option<&HtmlDivBorders> {
        self.borders.as_ref()
    }

    pub fn children(&self) -> &[HtmlDiv] {
        &self.children
    }
}

impl HtmlDivBorders {
    pub fn top(&self) -> Option<&HtmlDivBorder> {
        self.top.as_ref()
    }

    pub fn left(&self) -> Option<&HtmlDivBorder> {
        self.left.as_ref()
    }

    pub fn bottom(&self) -> Option<&HtmlDivBorder> {
        self.bottom.as_ref()
    }

    pub fn right(&self) -> Option<&HtmlDivBorder> {
        self.right.as_ref()
    }
}

impl HtmlDivBorder {
    pub fn style(&self) -> &str {
        &self.style
    }

    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    pub fn theme_color(&self) -> Option<ThemeColor> {
        self.theme_color
    }

    pub fn theme_tint(&self) -> Option<u8> {
        self.theme_tint
    }

    pub fn theme_shade(&self) -> Option<u8> {
        self.theme_shade
    }

    pub fn size_eighth_points(&self) -> Option<u64> {
        self.size_eighth_points
    }

    pub fn space_points(&self) -> Option<u64> {
        self.space_points
    }

    pub fn shadow(&self) -> Option<bool> {
        self.shadow
    }

    pub fn frame(&self) -> Option<bool> {
        self.frame
    }
}

impl WebSettings {
    /// Return the root frameset definition, if present.
    pub fn frameset(&self) -> Option<&Frameset> {
        self.frameset.as_ref()
    }

    /// Return the top-level HTML division definitions, preserving part absence.
    pub fn divs(&self) -> Option<&[HtmlDiv]> {
        self.divs.as_deref()
    }

    /// Return the requested output encoding, if declared.
    pub fn encoding(&self) -> Option<&str> {
        self.encoding.as_deref()
    }

    /// Return the browser-optimization setting, preserving absence.
    pub fn optimize_for_browser(&self) -> Option<bool> {
        self.optimize_for_browser
    }

    /// Return whether web output should rely on VML, preserving absence.
    pub fn rely_on_vml(&self) -> Option<bool> {
        self.rely_on_vml
    }

    /// Return whether PNG images are allowed, preserving absence.
    pub fn allow_png(&self) -> Option<bool> {
        self.allow_png
    }

    /// Return whether web output should avoid CSS, preserving absence.
    pub fn do_not_rely_on_css(&self) -> Option<bool> {
        self.do_not_rely_on_css
    }

    /// Return whether web output should use multiple files, preserving absence.
    pub fn do_not_save_as_single_file(&self) -> Option<bool> {
        self.do_not_save_as_single_file
    }

    /// Return whether supporting files should avoid a folder, preserving absence.
    pub fn do_not_organize_in_folder(&self) -> Option<bool> {
        self.do_not_organize_in_folder
    }

    /// Return whether web output should avoid long file names, preserving absence.
    pub fn do_not_use_long_file_names(&self) -> Option<bool> {
        self.do_not_use_long_file_names
    }

    /// Return the arbitrary-precision XML integer for web-output pixel density.
    pub fn pixels_per_inch(&self) -> Option<&str> {
        self.pixels_per_inch.as_deref()
    }

    /// Return the target screen size, if declared.
    pub fn target_screen_size(&self) -> Option<TargetScreenSize> {
        self.target_screen_size
    }

    /// Return whether smart tags should be saved as XML, preserving absence.
    pub fn save_smart_tags_as_xml(&self) -> Option<bool> {
        self.save_smart_tags_as_xml
    }

    /// Set the requested output encoding.
    pub fn set_encoding(&mut self, value: impl Into<String>) -> &mut Self {
        self.encoding = Some(value.into());
        self
    }

    /// Remove the requested output encoding.
    pub fn clear_encoding(&mut self) -> &mut Self {
        self.encoding = None;
        self
    }

    /// Set whether output should be optimized for browsers.
    pub fn set_optimize_for_browser(&mut self, value: bool) -> &mut Self {
        self.optimize_for_browser = Some(value);
        self
    }

    /// Restore schema-defined behavior for browser optimization.
    pub fn clear_optimize_for_browser(&mut self) -> &mut Self {
        self.optimize_for_browser = None;
        self
    }

    /// Set whether web output should rely on VML.
    pub fn set_rely_on_vml(&mut self, value: bool) -> &mut Self {
        self.rely_on_vml = Some(value);
        self
    }

    /// Restore schema-defined behavior for VML use.
    pub fn clear_rely_on_vml(&mut self) -> &mut Self {
        self.rely_on_vml = None;
        self
    }

    /// Set whether PNG images are allowed.
    pub fn set_allow_png(&mut self, value: bool) -> &mut Self {
        self.allow_png = Some(value);
        self
    }

    /// Restore schema-defined behavior for PNG images.
    pub fn clear_allow_png(&mut self) -> &mut Self {
        self.allow_png = None;
        self
    }

    /// Set whether web output should avoid CSS.
    pub fn set_do_not_rely_on_css(&mut self, value: bool) -> &mut Self {
        self.do_not_rely_on_css = Some(value);
        self
    }

    /// Restore schema-defined behavior for CSS output.
    pub fn clear_do_not_rely_on_css(&mut self) -> &mut Self {
        self.do_not_rely_on_css = None;
        self
    }

    /// Set whether web output should avoid a single-file representation.
    pub fn set_do_not_save_as_single_file(&mut self, value: bool) -> &mut Self {
        self.do_not_save_as_single_file = Some(value);
        self
    }

    /// Restore schema-defined behavior for single-file output.
    pub fn clear_do_not_save_as_single_file(&mut self) -> &mut Self {
        self.do_not_save_as_single_file = None;
        self
    }

    /// Set whether supporting files should avoid a dedicated folder.
    pub fn set_do_not_organize_in_folder(&mut self, value: bool) -> &mut Self {
        self.do_not_organize_in_folder = Some(value);
        self
    }

    /// Restore schema-defined behavior for supporting-file folders.
    pub fn clear_do_not_organize_in_folder(&mut self) -> &mut Self {
        self.do_not_organize_in_folder = None;
        self
    }

    /// Set whether web output should avoid long file names.
    pub fn set_do_not_use_long_file_names(&mut self, value: bool) -> &mut Self {
        self.do_not_use_long_file_names = Some(value);
        self
    }

    /// Restore schema-defined behavior for long file names.
    pub fn clear_do_not_use_long_file_names(&mut self) -> &mut Self {
        self.do_not_use_long_file_names = None;
        self
    }

    /// Set the arbitrary-precision XML integer for web-output pixel density.
    pub fn set_pixels_per_inch(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        self.pixels_per_inch = Some(normalize_xml_integer(value.into(), "pixels-per-inch")?);
        Ok(self)
    }

    /// Remove the explicit web-output pixel density.
    pub fn clear_pixels_per_inch(&mut self) -> &mut Self {
        self.pixels_per_inch = None;
        self
    }

    /// Set the target display size for generated web pages.
    pub fn set_target_screen_size(&mut self, value: TargetScreenSize) -> &mut Self {
        self.target_screen_size = Some(value);
        self
    }

    /// Remove the explicit target display size.
    pub fn clear_target_screen_size(&mut self) -> &mut Self {
        self.target_screen_size = None;
        self
    }

    /// Set whether smart tags should be retained in generated XML.
    pub fn set_save_smart_tags_as_xml(&mut self, value: bool) -> &mut Self {
        self.save_smart_tags_as_xml = Some(value);
        self
    }

    /// Restore schema-defined smart-tag serialization behavior.
    pub fn clear_save_smart_tags_as_xml(&mut self) -> &mut Self {
        self.save_smart_tags_as_xml = None;
        self
    }

    /// Serialize these web-output settings as canonical transitional WordprocessingML.
    ///
    /// The typed model is validated while parsing. Serialization additionally
    /// bounds recursive framesets and HTML divisions to the same safety limit
    /// used by the reader.
    pub fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        if let Some(frameset) = &self.frameset {
            write_frameset(&mut xml, frameset, 1)?;
        }
        if let Some(divs) = &self.divs {
            xml.push_str("<w:divs>");
            for div in divs {
                write_html_div(&mut xml, div, 1)?;
            }
            xml.push_str("</w:divs>");
        }
        if let Some(value) = &self.encoding {
            write_value_element(&mut xml, "encoding", value)?;
        }
        write_optional_on_off(&mut xml, "optimizeForBrowser", self.optimize_for_browser);
        write_optional_on_off(&mut xml, "relyOnVML", self.rely_on_vml);
        write_optional_on_off(&mut xml, "allowPNG", self.allow_png);
        write_optional_on_off(&mut xml, "doNotRelyOnCSS", self.do_not_rely_on_css);
        write_optional_on_off(
            &mut xml,
            "doNotSaveAsSingleFile",
            self.do_not_save_as_single_file,
        );
        write_optional_on_off(
            &mut xml,
            "doNotOrganizeInFolder",
            self.do_not_organize_in_folder,
        );
        write_optional_on_off(
            &mut xml,
            "doNotUseLongFileNames",
            self.do_not_use_long_file_names,
        );
        if let Some(value) = &self.pixels_per_inch {
            write_value_element(&mut xml, "pixelsPerInch", value)?;
        }
        if let Some(value) = self.target_screen_size {
            write_value_element(&mut xml, "targetScreenSz", value.as_str())?;
        }
        write_optional_on_off(&mut xml, "saveSmartTagsAsXml", self.save_smart_tags_as_xml);

        xml.push_str("</w:webSettings>");
        Ok(xml)
    }

    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let settings = Self::extract_from_xml(part.blob())?;
        validate_frame_relationships(part, &settings)?;
        Ok(settings)
    }

    fn extract_from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml);
        let mut settings = Self::default();
        let mut depth = 0usize;
        let mut saw_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "Word web-settings XML nesting is too deep".into(),
                        )
                    })?;
                    if depth == 1 {
                        validate_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if depth == 2 && saw_root && is_wordprocessing_namespace(&namespace) {
                        if element.local_name().as_ref() == b"frameset" {
                            let frameset = parse_frameset(&mut reader, 1)?;
                            set_once(&mut settings.frameset, frameset, "frameset")?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                OoxmlError::InvalidFormat(
                                    "invalid Word web-settings XML nesting".into(),
                                )
                            })?;
                        } else if element.local_name().as_ref() == b"divs" {
                            let divs = parse_div_container(&mut reader, b"divs", 1)?;
                            set_once(&mut settings.divs, divs, "divs")?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                OoxmlError::InvalidFormat(
                                    "invalid Word web-settings XML nesting".into(),
                                )
                            })?;
                        } else {
                            parse_setting(&element, decoder, &resolver, &mut settings)?;
                        }
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "Word web-settings XML nesting is too deep".into(),
                        )
                    })?;
                    if child_depth == 1 {
                        validate_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if child_depth == 2
                        && saw_root
                        && is_wordprocessing_namespace(&namespace)
                    {
                        if element.local_name().as_ref() == b"frameset" {
                            set_once(&mut settings.frameset, Frameset::default(), "frameset")?;
                        } else if element.local_name().as_ref() == b"divs" {
                            set_once(&mut settings.divs, Vec::new(), "divs")?;
                        } else {
                            parse_setting(&element, decoder, &resolver, &mut settings)?;
                        }
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word web-settings XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word web-settings XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(OoxmlError::InvalidFormat(
                "web-settings part has no webSettings root".into(),
            ));
        }
        Ok(settings)
    }
}

fn write_value_element(xml: &mut String, name: &str, value: &str) -> Result<()> {
    write!(xml, "<w:{name} w:val=\"{}\"/>", escape_xml(value))
        .map_err(|error| OoxmlError::Xml(error.to_string()))
}

fn write_optional_on_off(xml: &mut String, name: &str, value: Option<bool>) {
    match value {
        Some(true) => {
            write!(xml, "<w:{name}/>").expect("writing to a String cannot fail");
        },
        Some(false) => {
            write!(xml, "<w:{name} w:val=\"false\"/>").expect("writing to a String cannot fail");
        },
        None => {},
    }
}

fn write_frameset(xml: &mut String, frameset: &Frameset, nesting: usize) -> Result<()> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(OoxmlError::InvalidFormat(
            "Word web frameset nesting exceeds the supported safety limit".into(),
        ));
    }
    xml.push_str("<w:frameset>");
    if let Some(value) = &frameset.size {
        write_value_element(xml, "sz", value)?;
    }
    if let Some(split_bar) = &frameset.split_bar {
        write_frameset_split_bar(xml, split_bar)?;
    }
    if let Some(layout) = frameset.layout {
        write_value_element(xml, "frameLayout", layout.as_str())?;
    }
    for child in &frameset.children {
        match child {
            FramesetChild::Frameset(nested) => write_frameset(xml, nested, nesting + 1)?,
            FramesetChild::Frame(frame) => write_frame(xml, frame)?,
        }
    }
    xml.push_str("</w:frameset>");
    Ok(())
}

fn write_frame(xml: &mut String, frame: &Frame) -> Result<()> {
    xml.push_str("<w:frame>");
    if let Some(value) = &frame.size {
        write_value_element(xml, "sz", value)?;
    }
    if let Some(value) = &frame.name {
        write_value_element(xml, "name", value)?;
    }
    if let Some(value) = &frame.source_file_relationship_id {
        write!(xml, "<w:sourceFileName r:id=\"{}\"/>", escape_xml(value))
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(value) = frame.margin_width {
        write_value_element(xml, "marW", &value.to_string())?;
    }
    if let Some(value) = frame.margin_height {
        write_value_element(xml, "marH", &value.to_string())?;
    }
    if let Some(value) = frame.scrollbar {
        write_value_element(xml, "scrollbar", value.as_str())?;
    }
    write_optional_on_off(xml, "noResizeAllowed", frame.no_resize_allowed);
    write_optional_on_off(xml, "linkedToFile", frame.linked_to_file);
    xml.push_str("</w:frame>");
    Ok(())
}

fn write_frameset_split_bar(xml: &mut String, split_bar: &FramesetSplitBar) -> Result<()> {
    xml.push_str("<w:framesetSplitbar>");
    if let Some(value) = split_bar.width_twips {
        write_value_element(xml, "w", &value.to_string())?;
    }
    if let Some(color) = &split_bar.color {
        xml.push_str("<w:color");
        write_color_attributes(
            xml,
            &color.value,
            color.theme_color,
            color.theme_tint,
            color.theme_shade,
        )?;
        xml.push_str("/>");
    }
    write_optional_on_off(xml, "noBorder", split_bar.no_border);
    write_optional_on_off(xml, "flatBorders", split_bar.flat_borders);
    xml.push_str("</w:framesetSplitbar>");
    Ok(())
}

fn write_html_div(xml: &mut String, div: &HtmlDiv, nesting: usize) -> Result<()> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(OoxmlError::InvalidFormat(
            "Word HTML division nesting exceeds the supported safety limit".into(),
        ));
    }
    write!(xml, "<w:div w:id=\"{}\">", escape_xml(&div.id))
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    write_optional_on_off(xml, "blockQuote", div.block_quote);
    write_optional_on_off(xml, "bodyDiv", div.body_div);
    for (name, value) in [
        ("marLeft", &div.margin_left_twips),
        ("marRight", &div.margin_right_twips),
        ("marTop", &div.margin_top_twips),
        ("marBottom", &div.margin_bottom_twips),
    ] {
        if let Some(value) = value {
            write_value_element(xml, name, value)?;
        }
    }
    if let Some(borders) = &div.borders {
        write_html_div_borders(xml, borders)?;
    }
    if !div.children.is_empty() {
        xml.push_str("<w:divsChild>");
        for child in &div.children {
            write_html_div(xml, child, nesting + 1)?;
        }
        xml.push_str("</w:divsChild>");
    }
    xml.push_str("</w:div>");
    Ok(())
}

fn write_html_div_borders(xml: &mut String, borders: &HtmlDivBorders) -> Result<()> {
    xml.push_str("<w:divBdr>");
    for (name, border) in [
        ("top", &borders.top),
        ("left", &borders.left),
        ("bottom", &borders.bottom),
        ("right", &borders.right),
    ] {
        let Some(border) = border else {
            continue;
        };
        write!(xml, "<w:{name} w:val=\"{}\"", escape_xml(&border.style))
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if let Some(color) = &border.color {
            write!(xml, " w:color=\"{}\"", escape_xml(color))
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        write_theme_attributes(
            xml,
            border.theme_color,
            border.theme_tint,
            border.theme_shade,
        )?;
        if let Some(value) = border.size_eighth_points {
            write!(xml, " w:sz=\"{value}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        if let Some(value) = border.space_points {
            write!(xml, " w:space=\"{value}\"")
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        write_optional_on_off_attribute(xml, "shadow", border.shadow)?;
        write_optional_on_off_attribute(xml, "frame", border.frame)?;
        xml.push_str("/>");
    }
    xml.push_str("</w:divBdr>");
    Ok(())
}

fn write_color_attributes(
    xml: &mut String,
    value: &str,
    theme_color: Option<ThemeColor>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
) -> Result<()> {
    write!(xml, " w:val=\"{}\"", escape_xml(value))
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    write_theme_attributes(xml, theme_color, theme_tint, theme_shade)
}

fn write_theme_attributes(
    xml: &mut String,
    theme_color: Option<ThemeColor>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
) -> Result<()> {
    if let Some(value) = theme_color {
        write!(xml, " w:themeColor=\"{}\"", value.as_str())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(value) = theme_tint {
        write!(xml, " w:themeTint=\"{value:02X}\"")
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(value) = theme_shade {
        write!(xml, " w:themeShade=\"{value:02X}\"")
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    Ok(())
}

fn write_optional_on_off_attribute(
    xml: &mut String,
    name: &str,
    value: Option<bool>,
) -> Result<()> {
    if let Some(value) = value {
        write!(
            xml,
            " w:{name}=\"{}\"",
            if value { "true" } else { "false" }
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    Ok(())
}

fn validate_frame_relationships(part: &dyn Part, settings: &WebSettings) -> Result<()> {
    const FRAME_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame";
    const STRICT_FRAME_RELATIONSHIP: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/frame";

    fn validate(
        part: &dyn Part,
        frameset: &Frameset,
        transitional_type: &str,
        strict_type: &str,
    ) -> Result<()> {
        for child in &frameset.children {
            match child {
                FramesetChild::Frameset(nested) => {
                    validate(part, nested, transitional_type, strict_type)?;
                },
                FramesetChild::Frame(frame) => {
                    let Some(id) = &frame.source_file_relationship_id else {
                        continue;
                    };
                    let relationship = part.rels().get(id).ok_or_else(|| {
                        OoxmlError::InvalidFormat(format!(
                            "frame source relationship '{id}' does not exist"
                        ))
                    })?;
                    if relationship.reltype() != transitional_type
                        && relationship.reltype() != strict_type
                    {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "frame source relationship '{id}' has an invalid type"
                        )));
                    }
                },
            }
        }
        Ok(())
    }

    if let Some(frameset) = &settings.frameset {
        validate(
            part,
            frameset,
            FRAME_RELATIONSHIP,
            STRICT_FRAME_RELATIONSHIP,
        )?;
    }
    Ok(())
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"webSettings"
    {
        return Err(OoxmlError::InvalidFormat(
            "web-settings part has an invalid or trailing root element".into(),
        ));
    }
    Ok(())
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut WebSettings,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"encoding" => set_once(
            &mut settings.encoding,
            required_value(element, decoder, resolver, "web encoding")?,
            "encoding",
        ),
        b"optimizeForBrowser" => set_on_off(
            &mut settings.optimize_for_browser,
            element,
            decoder,
            resolver,
            "optimizeForBrowser",
        ),
        b"relyOnVML" => set_on_off(
            &mut settings.rely_on_vml,
            element,
            decoder,
            resolver,
            "relyOnVML",
        ),
        b"allowPNG" => set_on_off(
            &mut settings.allow_png,
            element,
            decoder,
            resolver,
            "allowPNG",
        ),
        b"doNotRelyOnCSS" => set_on_off(
            &mut settings.do_not_rely_on_css,
            element,
            decoder,
            resolver,
            "doNotRelyOnCSS",
        ),
        b"doNotSaveAsSingleFile" => set_on_off(
            &mut settings.do_not_save_as_single_file,
            element,
            decoder,
            resolver,
            "doNotSaveAsSingleFile",
        ),
        b"doNotOrganizeInFolder" => set_on_off(
            &mut settings.do_not_organize_in_folder,
            element,
            decoder,
            resolver,
            "doNotOrganizeInFolder",
        ),
        b"doNotUseLongFileNames" => set_on_off(
            &mut settings.do_not_use_long_file_names,
            element,
            decoder,
            resolver,
            "doNotUseLongFileNames",
        ),
        b"pixelsPerInch" => {
            let value = required_value(element, decoder, resolver, "pixels per inch")?;
            let value = normalize_xml_integer(value, "pixels-per-inch")?;
            set_once(&mut settings.pixels_per_inch, value, "pixelsPerInch")
        },
        b"targetScreenSz" => {
            let value = required_value(element, decoder, resolver, "target screen size")?;
            let value = TargetScreenSize::from_xml(&value).ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("invalid target-screen-size value '{value}'"))
            })?;
            set_once(&mut settings.target_screen_size, value, "targetScreenSz")
        },
        b"saveSmartTagsAsXml" => set_on_off(
            &mut settings.save_smart_tags_as_xml,
            element,
            decoder,
            resolver,
            "saveSmartTagsAsXml",
        ),
        _ => Ok(()),
    }
}

const MAX_FRAMESET_NESTING: usize = 128;

fn parse_frameset(reader: &mut NsReader<&[u8]>, nesting: usize) -> Result<Frameset> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(OoxmlError::InvalidFormat(
            "Word web frameset nesting exceeds the supported safety limit".into(),
        ));
    }
    let mut frameset = Frameset::default();
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    b"sz" => {
                        let value = required_value(&element, decoder, &resolver, "frame size")?;
                        set_once(&mut frameset.size, value, "frameset size")?;
                        finish_leaf(reader, "frameset size")?;
                    },
                    b"framesetSplitbar" => {
                        let split_bar = parse_frameset_split_bar(reader)?;
                        set_once(&mut frameset.split_bar, split_bar, "frameset split bar")?;
                    },
                    b"frameLayout" => {
                        let layout = parse_frame_layout(&element, decoder, &resolver)?;
                        set_once(&mut frameset.layout, layout, "frame layout")?;
                        finish_leaf(reader, "frame layout")?;
                    },
                    b"frameset" => frameset
                        .children
                        .push(FramesetChild::Frameset(parse_frameset(
                            reader,
                            nesting + 1,
                        )?)),
                    b"frame" => frameset
                        .children
                        .push(FramesetChild::Frame(parse_frame(reader)?)),
                    _ => skip_element(reader)?,
                }
            },
            Event::Empty(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    b"sz" => set_once(
                        &mut frameset.size,
                        required_value(&element, decoder, &resolver, "frame size")?,
                        "frameset size",
                    )?,
                    b"framesetSplitbar" => set_once(
                        &mut frameset.split_bar,
                        FramesetSplitBar::default(),
                        "frameset split bar",
                    )?,
                    b"frameLayout" => set_once(
                        &mut frameset.layout,
                        parse_frame_layout(&element, decoder, &resolver)?,
                        "frame layout",
                    )?,
                    b"frameset" => frameset
                        .children
                        .push(FramesetChild::Frameset(Frameset::default())),
                    b"frame" => frameset
                        .children
                        .push(FramesetChild::Frame(Frame::default())),
                    _ => {},
                }
            },
            Event::Start(_) => skip_element(reader)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"frameset" =>
            {
                return Ok(frameset);
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word web frameset".into(),
                ));
            },
            _ => {},
        }
    }
}

fn parse_div_container(
    reader: &mut NsReader<&[u8]>,
    end_name: &[u8],
    nesting: usize,
) -> Result<Vec<HtmlDiv>> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(OoxmlError::InvalidFormat(
            "Word HTML division nesting exceeds the supported safety limit".into(),
        ));
    }
    let mut divs = Vec::new();
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                divs.push(parse_html_div(
                    reader, &element, decoder, &resolver, nesting,
                )?);
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                divs.push(new_html_div(&element, decoder, &resolver)?);
            },
            Event::Start(_) => skip_element(reader)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == end_name =>
            {
                return Ok(divs);
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word HTML division container".into(),
                ));
            },
            _ => {},
        }
    }
}

fn new_html_div(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<HtmlDiv> {
    let id = word_attribute_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| OoxmlError::InvalidFormat("Word HTML division ID is required".into()))?;
    Ok(HtmlDiv {
        id,
        ..HtmlDiv::default()
    })
}

fn parse_html_div(
    reader: &mut NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    nesting: usize,
) -> Result<HtmlDiv> {
    let mut div = new_html_div(element, decoder, resolver)?;
    let mut children_present = false;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    name if is_html_div_leaf(name) => {
                        parse_html_div_leaf(&element, decoder, &resolver, &mut div)?;
                        finish_leaf(reader, "HTML division property")?;
                    },
                    b"divBdr" => {
                        let borders = parse_html_div_borders(reader)?;
                        set_once(&mut div.borders, borders, "HTML division borders")?;
                    },
                    b"divsChild" => {
                        if children_present {
                            return Err(OoxmlError::InvalidFormat(
                                "duplicate Word HTML child division container".into(),
                            ));
                        }
                        children_present = true;
                        div.children = parse_div_container(reader, b"divsChild", nesting + 1)?;
                    },
                    _ => skip_element(reader)?,
                }
            },
            Event::Empty(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    name if is_html_div_leaf(name) => {
                        parse_html_div_leaf(&element, decoder, &resolver, &mut div)?;
                    },
                    b"divBdr" => set_once(
                        &mut div.borders,
                        HtmlDivBorders::default(),
                        "HTML division borders",
                    )?,
                    b"divsChild" => {
                        mark_html_div_children_present(&mut children_present)?;
                    },
                    _ => {},
                }
            },
            Event::Start(_) => skip_element(reader)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                return Ok(div);
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word HTML division".into(),
                ));
            },
            _ => {},
        }
    }
}

fn mark_html_div_children_present(present: &mut bool) -> Result<()> {
    if std::mem::replace(present, true) {
        return Err(OoxmlError::InvalidFormat(
            "duplicate Word HTML child division container".into(),
        ));
    }
    Ok(())
}

fn is_html_div_leaf(name: &[u8]) -> bool {
    matches!(
        name,
        b"blockQuote" | b"bodyDiv" | b"marLeft" | b"marRight" | b"marTop" | b"marBottom"
    )
}

fn parse_html_div_leaf(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    div: &mut HtmlDiv,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"blockQuote" => set_on_off(
            &mut div.block_quote,
            element,
            decoder,
            resolver,
            "HTML blockQuote",
        ),
        b"bodyDiv" => set_on_off(
            &mut div.body_div,
            element,
            decoder,
            resolver,
            "HTML bodyDiv",
        ),
        b"marLeft" => set_signed_twips(
            &mut div.margin_left_twips,
            element,
            decoder,
            resolver,
            "HTML division left margin",
        ),
        b"marRight" => set_signed_twips(
            &mut div.margin_right_twips,
            element,
            decoder,
            resolver,
            "HTML division right margin",
        ),
        b"marTop" => set_signed_twips(
            &mut div.margin_top_twips,
            element,
            decoder,
            resolver,
            "HTML division top margin",
        ),
        b"marBottom" => set_signed_twips(
            &mut div.margin_bottom_twips,
            element,
            decoder,
            resolver,
            "HTML division bottom margin",
        ),
        _ => Ok(()),
    }
}

fn set_signed_twips(
    slot: &mut Option<String>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<()> {
    let value = required_value(element, decoder, resolver, description)?;
    let value = normalize_xml_integer(value, description)?;
    set_once(slot, value, description)
}

fn parse_html_div_borders(reader: &mut NsReader<&[u8]>) -> Result<HtmlDivBorders> {
    let mut borders = HtmlDivBorders::default();
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_html_div_border_side(element.local_name().as_ref()) =>
            {
                set_html_div_border(&mut borders, &element, decoder, &resolver)?;
                finish_leaf(reader, "HTML division border")?;
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_html_div_border_side(element.local_name().as_ref()) =>
            {
                set_html_div_border(&mut borders, &element, decoder, &resolver)?;
            },
            Event::Start(_) => skip_element(reader)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"divBdr" =>
            {
                return Ok(borders);
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word HTML division borders".into(),
                ));
            },
            _ => {},
        }
    }
}

fn is_html_div_border_side(name: &[u8]) -> bool {
    matches!(name, b"top" | b"left" | b"bottom" | b"right")
}

fn set_html_div_border(
    borders: &mut HtmlDivBorders,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let border = parse_html_div_border(element, decoder, resolver)?;
    let (slot, description) = match element.local_name().as_ref() {
        b"top" => (&mut borders.top, "top HTML division border"),
        b"left" => (&mut borders.left, "left HTML division border"),
        b"bottom" => (&mut borders.bottom, "bottom HTML division border"),
        b"right" => (&mut borders.right, "right HTML division border"),
        _ => return Ok(()),
    };
    set_once(slot, border, description)
}

fn parse_html_div_border(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<HtmlDivBorder> {
    let style = required_value(element, decoder, resolver, "HTML division border style")?;
    let color = word_attribute_value(element, b"color", decoder, resolver)?;
    if let Some(color) = &color
        && color != "auto"
        && (color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid HTML division border color '{color}'"
        )));
    }
    let theme_color = word_attribute_value(element, b"themeColor", decoder, resolver)?
        .map(|value| {
            ThemeColor::from_xml(&value)
                .ok_or_else(|| OoxmlError::InvalidFormat(format!("invalid theme color '{value}'")))
        })
        .transpose()?;
    Ok(HtmlDivBorder {
        style,
        color,
        theme_color,
        theme_tint: optional_hex_byte(element, b"themeTint", decoder, resolver)?,
        theme_shade: optional_hex_byte(element, b"themeShade", decoder, resolver)?,
        size_eighth_points: optional_unsigned_long_attribute(element, b"sz", decoder, resolver)?,
        space_points: optional_unsigned_long_attribute(element, b"space", decoder, resolver)?,
        shadow: optional_on_off_attribute(element, b"shadow", decoder, resolver)?,
        frame: optional_on_off_attribute(element, b"frame", decoder, resolver)?,
    })
}

fn parse_frame(reader: &mut NsReader<&[u8]>) -> Result<Frame> {
    let mut frame = Frame::default();
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                if is_frame_property(element.local_name().as_ref()) {
                    parse_frame_property(&element, decoder, &resolver, &mut frame)?;
                    finish_leaf(reader, "frame property")?;
                } else {
                    skip_element(reader)?;
                }
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_frame_property(element.local_name().as_ref()) =>
            {
                parse_frame_property(&element, decoder, &resolver, &mut frame)?;
            },
            Event::Start(_) => skip_element(reader)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"frame" =>
            {
                return Ok(frame);
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word web frame".into(),
                ));
            },
            _ => {},
        }
    }
}

fn is_frame_property(name: &[u8]) -> bool {
    matches!(
        name,
        b"sz"
            | b"name"
            | b"sourceFileName"
            | b"marW"
            | b"marH"
            | b"scrollbar"
            | b"noResizeAllowed"
            | b"linkedToFile"
    )
}

fn parse_frame_property(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    frame: &mut Frame,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"sz" => set_once(
            &mut frame.size,
            required_value(element, decoder, resolver, "frame size")?,
            "frame size",
        ),
        b"name" => set_once(
            &mut frame.name,
            required_value(element, decoder, resolver, "frame name")?,
            "frame name",
        ),
        b"sourceFileName" => set_once(
            &mut frame.source_file_relationship_id,
            required_relationship_id(element, decoder, resolver)?,
            "frame source file",
        ),
        b"marW" => set_once(
            &mut frame.margin_width,
            required_unsigned_long(element, decoder, resolver, "frame margin width")?,
            "frame margin width",
        ),
        b"marH" => set_once(
            &mut frame.margin_height,
            required_unsigned_long(element, decoder, resolver, "frame margin height")?,
            "frame margin height",
        ),
        b"scrollbar" => {
            let value = required_value(element, decoder, resolver, "frame scrollbar")?;
            let value = FrameScrollbarVisibility::from_xml(&value).ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("invalid frame scrollbar value '{value}'"))
            })?;
            set_once(&mut frame.scrollbar, value, "frame scrollbar")
        },
        b"noResizeAllowed" => set_on_off(
            &mut frame.no_resize_allowed,
            element,
            decoder,
            resolver,
            "frame noResizeAllowed",
        ),
        b"linkedToFile" => set_on_off(
            &mut frame.linked_to_file,
            element,
            decoder,
            resolver,
            "frame linkedToFile",
        ),
        _ => Ok(()),
    }
}

fn parse_frameset_split_bar(reader: &mut NsReader<&[u8]>) -> Result<FramesetSplitBar> {
    let mut split_bar = FramesetSplitBar::default();
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                if is_split_bar_property(element.local_name().as_ref()) {
                    parse_split_bar_property(&element, decoder, &resolver, &mut split_bar)?;
                    finish_leaf(reader, "frameset split-bar property")?;
                } else {
                    skip_element(reader)?;
                }
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_split_bar_property(element.local_name().as_ref()) =>
            {
                parse_split_bar_property(&element, decoder, &resolver, &mut split_bar)?;
            },
            Event::Start(_) => skip_element(reader)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"framesetSplitbar" =>
            {
                return Ok(split_bar);
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word frameset split bar".into(),
                ));
            },
            _ => {},
        }
    }
}

fn is_split_bar_property(name: &[u8]) -> bool {
    matches!(name, b"w" | b"color" | b"noBorder" | b"flatBorders")
}

fn parse_split_bar_property(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    split_bar: &mut FramesetSplitBar,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"w" => set_once(
            &mut split_bar.width_twips,
            required_unsigned_long(element, decoder, resolver, "split-bar width")?,
            "split-bar width",
        ),
        b"color" => set_once(
            &mut split_bar.color,
            parse_frameset_color(element, decoder, resolver)?,
            "split-bar color",
        ),
        b"noBorder" => set_on_off(
            &mut split_bar.no_border,
            element,
            decoder,
            resolver,
            "split-bar noBorder",
        ),
        b"flatBorders" => set_on_off(
            &mut split_bar.flat_borders,
            element,
            decoder,
            resolver,
            "split-bar flatBorders",
        ),
        _ => Ok(()),
    }
}

fn parse_frame_layout(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<FrameLayout> {
    let value = required_value(element, decoder, resolver, "frame layout")?;
    FrameLayout::from_xml(&value)
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("invalid frame-layout value '{value}'")))
}

fn parse_frameset_color(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<FramesetColor> {
    let value = required_value(element, decoder, resolver, "frameset splitter color")?;
    if value != "auto" && (value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid frameset splitter color '{value}'"
        )));
    }
    let theme_color = word_attribute_value(element, b"themeColor", decoder, resolver)?
        .map(|value| {
            ThemeColor::from_xml(&value)
                .ok_or_else(|| OoxmlError::InvalidFormat(format!("invalid theme color '{value}'")))
        })
        .transpose()?;
    let theme_tint = optional_hex_byte(element, b"themeTint", decoder, resolver)?;
    let theme_shade = optional_hex_byte(element, b"themeShade", decoder, resolver)?;
    Ok(FramesetColor {
        value,
        theme_color,
        theme_tint,
        theme_shade,
    })
}

fn required_unsigned_long(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<u64> {
    let value = required_value(element, decoder, resolver, description)?;
    value.trim().parse::<u64>().map_err(|_| {
        OoxmlError::InvalidFormat(format!("invalid unsigned {description} value '{value}'"))
    })
}

fn optional_unsigned_long_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<u64>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| {
            value.trim().parse::<u64>().map_err(|_| {
                OoxmlError::InvalidFormat(format!(
                    "invalid unsigned Word attribute value '{value}'"
                ))
            })
        })
        .transpose()
}

fn optional_on_off_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<bool>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| match value.as_str() {
            "true" | "1" | "on" => Ok(true),
            "false" | "0" | "off" => Ok(false),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid Word on/off value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_hex_byte(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<u8>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| {
            if value.len() != 2 {
                return Err(OoxmlError::InvalidFormat(format!(
                    "invalid hexadecimal byte '{value}'"
                )));
            }
            u8::from_str_radix(&value, 16).map_err(|_| {
                OoxmlError::InvalidFormat(format!("invalid hexadecimal byte '{value}'"))
            })
        })
        .transpose()
}

fn required_relationship_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    const RELATIONSHIPS: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(quick_xml::name::Namespace(namespace))
                if namespace == RELATIONSHIPS || namespace == STRICT_RELATIONSHIPS
        );
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "duplicate frame source relationship ID".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    value
        .ok_or_else(|| OoxmlError::InvalidFormat("frame source relationship ID is required".into()))
}

fn finish_leaf(reader: &mut NsReader<&[u8]>, description: &str) -> Result<()> {
    loop {
        match reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
        {
            Event::End(_) => return Ok(()),
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "unterminated Word {description} element"
                )));
            },
            _ => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "Word {description} must not contain child content"
                )));
            },
        }
    }
}

fn skip_element(reader: &mut NsReader<&[u8]>) -> Result<()> {
    let mut depth = 1usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
        {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("Word web XML nesting is too deep".into())
                })?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid Word web XML nesting".into())
                })?;
                if depth == 0 {
                    return Ok(());
                }
            },
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated Word web XML element".into(),
                ));
            },
            _ => {},
        }
    }
}

fn required_value(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, b"val", decoder, resolver)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("Word {description} value is required")))
}

fn set_on_off(
    slot: &mut Option<bool>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<()> {
    let value = match word_attribute_value(element, b"val", decoder, resolver)? {
        Some(value) => match value.as_str() {
            "true" | "1" | "on" => true,
            "false" | "0" | "off" => false,
            _ => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "invalid Word on/off value '{value}'"
                )));
            },
        },
        None => true,
    };
    set_once(slot, value, description)
}

fn set_once<T>(slot: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate Word web setting '{description}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    #[test]
    fn parses_all_scalar_web_settings_with_strict_namespaces() {
        let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:false="urn:not-wordprocessingml">
            <s:encoding s:val="utf-8"/>
            <s:optimizeForBrowser s:val="on"/>
            <s:relyOnVML s:val="off"/>
            <s:allowPNG/>
            <s:doNotRelyOnCSS s:val="0"/>
            <s:doNotSaveAsSingleFile s:val="1"/>
            <s:doNotOrganizeInFolder s:val="false"/>
            <s:doNotUseLongFileNames s:val="true"/>
            <s:pixelsPerInch s:val=" 123456789012345678901234567890 "/>
            <s:targetScreenSz s:val="1920x1200"/>
            <s:saveSmartTagsAsXml s:val="on"/>
            <false:saveSmartTagsAsXml false:val="off"/>
        </s:webSettings>"#;

        let settings = WebSettings::extract_from_xml(xml).unwrap();
        assert_eq!(settings.encoding(), Some("utf-8"));
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(settings.rely_on_vml(), Some(false));
        assert_eq!(settings.allow_png(), Some(true));
        assert_eq!(settings.do_not_rely_on_css(), Some(false));
        assert_eq!(settings.do_not_save_as_single_file(), Some(true));
        assert_eq!(settings.do_not_organize_in_folder(), Some(false));
        assert_eq!(settings.do_not_use_long_file_names(), Some(true));
        assert_eq!(
            settings.pixels_per_inch(),
            Some("123456789012345678901234567890")
        );
        assert_eq!(
            settings.target_screen_size(),
            Some(TargetScreenSize::Pixels1920x1200)
        );
        assert_eq!(settings.save_smart_tags_as_xml(), Some(true));
    }

    #[test]
    fn rejects_invalid_or_duplicate_scalar_web_settings() {
        let missing_value = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pixelsPerInch/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(missing_value).is_err());

        let invalid_on_off = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:saveSmartTagsAsXml w:val="maybe"/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(invalid_on_off).is_err());

        let duplicate = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG/><w:allowPNG/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(duplicate).is_err());

        let invalid_screen = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:targetScreenSz w:val="1366x768"/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(invalid_screen).is_err());
    }

    #[test]
    fn parses_recursive_framesets_and_all_frame_properties() {
        let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"
            xmlns:false="urn:not-wordprocessingml">
          <s:frameset>
            <s:sz s:val="2*"/>
            <s:framesetSplitbar>
              <s:w s:val="90"/>
              <s:color s:val="auto" s:themeColor="accent2" s:themeTint="7f" s:themeShade="00"/>
              <s:noBorder s:val="off"/>
              <s:flatBorders/>
            </s:framesetSplitbar>
            <s:frameLayout s:val="cols"/>
            <s:frame>
              <s:sz s:val="50%"/>
              <s:name s:val="navigation"/>
              <s:sourceFileName rel:id="rId7"/>
              <s:marW s:val="18446744073709551615"/>
              <s:marH s:val="24"/>
              <s:scrollbar s:val="auto"/>
              <s:noResizeAllowed/>
              <s:linkedToFile s:val="false"/>
              <s:futureExtension><s:nested/></s:futureExtension>
            </s:frame>
            <s:frameset>
              <s:frameLayout s:val="none"/>
              <s:frame><s:name s:val="content"/></s:frame>
            </s:frameset>
            <false:frame><false:name false:val="ignored"/></false:frame>
          </s:frameset>
        </s:webSettings>"#;

        let settings = WebSettings::extract_from_xml(xml).unwrap();
        let frameset = settings.frameset().unwrap();
        assert_eq!(frameset.size(), Some("2*"));
        assert_eq!(frameset.layout(), Some(FrameLayout::Columns));
        let split_bar = frameset.split_bar().unwrap();
        assert_eq!(split_bar.width_twips(), Some(90));
        assert_eq!(split_bar.no_border(), Some(false));
        assert_eq!(split_bar.flat_borders(), Some(true));
        let color = split_bar.color().unwrap();
        assert_eq!(color.value(), "auto");
        assert_eq!(color.theme_color(), Some(ThemeColor::Accent2));
        assert_eq!(color.theme_tint(), Some(0x7f));
        assert_eq!(color.theme_shade(), Some(0));
        assert_eq!(frameset.children().len(), 2);

        let FramesetChild::Frame(frame) = &frameset.children()[0] else {
            panic!("first frameset child must be a frame");
        };
        assert_eq!(frame.size(), Some("50%"));
        assert_eq!(frame.name(), Some("navigation"));
        assert_eq!(frame.source_file_relationship_id(), Some("rId7"));
        assert_eq!(frame.margin_width(), Some(u64::MAX));
        assert_eq!(frame.margin_height(), Some(24));
        assert_eq!(frame.scrollbar(), Some(FrameScrollbarVisibility::Auto));
        assert_eq!(frame.no_resize_allowed(), Some(true));
        assert_eq!(frame.linked_to_file(), Some(false));

        let FramesetChild::Frameset(nested) = &frameset.children()[1] else {
            panic!("second frameset child must be nested");
        };
        assert_eq!(nested.layout(), Some(FrameLayout::None));
        let FramesetChild::Frame(frame) = &nested.children()[0] else {
            panic!("nested child must be a frame");
        };
        assert_eq!(frame.name(), Some("content"));
    }

    #[test]
    fn validates_frame_values_and_source_relationships() {
        let invalid_layout = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frameLayout w:val="diagonal"/></w:frameset></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(invalid_layout).is_err());

        let overflowing_pixels = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:marW w:val="18446744073709551616"/></w:frame></w:frameset></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(overflowing_pixels).is_err());

        let child_in_leaf = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:name w:val="bad"><w:frame/></w:name></w:frame></w:frameset></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(child_in_leaf).is_err());

        let duplicate = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:name w:val="one"/><w:name w:val="two"/></w:frame></w:frameset></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(duplicate).is_err());

        let xml = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:frameset><w:frame><w:sourceFileName r:id="rId1"/></w:frame></w:frameset></w:webSettings>"#;
        let mut part = BlobPart::new(
            PackURI::new("/word/webSettings.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml".into(),
            xml.to_vec(),
        );
        assert!(WebSettings::extract_from_part(&part).is_err());
        part.rels_mut().add_relationship(
            "http://purl.oclc.org/ooxml/officeDocument/relationships/frame".into(),
            "https://example.test/frame.html".into(),
            "rId1".into(),
            true,
        );
        assert!(WebSettings::extract_from_part(&part).is_ok());
    }

    #[test]
    fn parses_recursive_html_divisions_and_border_properties() {
        let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:not-wordprocessingml">
          <s:divs>
            <s:div s:id="1785730240">
              <s:blockQuote/>
              <s:bodyDiv s:val="off"/>
              <s:marLeft s:val=" -123456789012345678901234567890 "/>
              <s:marRight s:val="+42"/>
              <s:marTop s:val="0"/>
              <s:marBottom s:val="700"/>
              <s:divBdr>
                <s:top s:val="single" s:color="A0b1C2" s:themeColor="text2" s:themeTint="10" s:themeShade="ff" s:sz="18446744073709551615" s:space="6" s:shadow="on" s:frame="0"/>
                <s:left s:val="zigZagStitch"/>
              </s:divBdr>
              <s:divsChild>
                <s:div s:id="child"><s:bodyDiv/></s:div>
              </s:divsChild>
            </s:div>
            <s:div s:id="second"/>
            <false:div false:id="ignored"/>
          </s:divs>
        </s:webSettings>"#;

        let settings = WebSettings::extract_from_xml(xml).unwrap();
        let divs = settings.divs().unwrap();
        assert_eq!(divs.len(), 2);
        let div = &divs[0];
        assert_eq!(div.id(), "1785730240");
        assert_eq!(div.is_block_quote(), Some(true));
        assert_eq!(div.is_body_div(), Some(false));
        assert_eq!(
            div.margin_left_twips(),
            Some("-123456789012345678901234567890")
        );
        assert_eq!(div.margin_right_twips(), Some("+42"));
        assert_eq!(div.margin_top_twips(), Some("0"));
        assert_eq!(div.margin_bottom_twips(), Some("700"));
        assert_eq!(div.children()[0].id(), "child");
        assert_eq!(div.children()[0].is_body_div(), Some(true));

        let borders = div.borders().unwrap();
        let top = borders.top().unwrap();
        assert_eq!(top.style(), "single");
        assert_eq!(top.color(), Some("A0b1C2"));
        assert_eq!(top.theme_color(), Some(ThemeColor::Text2));
        assert_eq!(top.theme_tint(), Some(0x10));
        assert_eq!(top.theme_shade(), Some(0xff));
        assert_eq!(top.size_eighth_points(), Some(u64::MAX));
        assert_eq!(top.space_points(), Some(6));
        assert_eq!(top.shadow(), Some(true));
        assert_eq!(top.frame(), Some(false));
        assert_eq!(borders.left().unwrap().style(), "zigZagStitch");
        assert!(borders.bottom().is_none());
        assert!(borders.right().is_none());
    }

    #[test]
    fn validates_html_division_structure_and_values() {
        const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let missing_id =
            format!(r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div/></w:divs></w:webSettings>"#);
        assert!(WebSettings::extract_from_xml(missing_id.as_bytes()).is_err());

        let invalid_margin = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marLeft w:val="1.5"/></w:div></w:divs></w:webSettings>"#
        );
        assert!(WebSettings::extract_from_xml(invalid_margin.as_bytes()).is_err());

        let invalid_color = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:divBdr><w:left w:val="single" w:color="xyz"/></w:divBdr></w:div></w:divs></w:webSettings>"#
        );
        assert!(WebSettings::extract_from_xml(invalid_color.as_bytes()).is_err());

        let duplicate_child_container = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:divsChild/><w:divsChild/></w:div></w:divs></w:webSettings>"#
        );
        assert!(WebSettings::extract_from_xml(duplicate_child_container.as_bytes()).is_err());
    }

    #[test]
    fn serializes_every_modeled_web_setting_for_round_trip() {
        let xml = br#"<w:webSettings
          xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:frameset>
            <w:sz w:val="2* &amp; 1*"/>
            <w:framesetSplitbar>
              <w:w w:val="18446744073709551615"/>
              <w:color w:val="A0b1C2" w:themeColor="accent4" w:themeTint="0a" w:themeShade="FF"/>
              <w:noBorder/>
              <w:flatBorders w:val="false"/>
            </w:framesetSplitbar>
            <w:frameLayout w:val="cols"/>
            <w:frame>
              <w:sz w:val="50%"/>
              <w:name w:val="main &amp; detail"/>
              <w:sourceFileName r:id="rId7"/>
              <w:marW w:val="42"/>
              <w:marH w:val="24"/>
              <w:scrollbar w:val="auto"/>
              <w:noResizeAllowed w:val="off"/>
              <w:linkedToFile/>
            </w:frame>
            <w:frameset><w:frameLayout w:val="none"/></w:frameset>
          </w:frameset>
          <w:divs>
            <w:div w:id="root&amp;division">
              <w:blockQuote/>
              <w:bodyDiv w:val="0"/>
              <w:marLeft w:val="-123456789012345678901234567890"/>
              <w:marRight w:val="+42"/>
              <w:marTop w:val="0"/>
              <w:marBottom w:val="700"/>
              <w:divBdr>
                <w:top w:val="single" w:color="auto" w:themeColor="text2" w:themeTint="10" w:themeShade="ff" w:sz="18446744073709551615" w:space="6" w:shadow="on" w:frame="0"/>
                <w:left w:val="zigZagStitch"/>
              </w:divBdr>
              <w:divsChild><w:div w:id="child"><w:bodyDiv/></w:div></w:divsChild>
            </w:div>
          </w:divs>
          <w:encoding w:val="utf&amp;8"/>
          <w:optimizeForBrowser/>
          <w:relyOnVML w:val="false"/>
          <w:allowPNG/>
          <w:doNotRelyOnCSS w:val="off"/>
          <w:doNotSaveAsSingleFile/>
          <w:doNotOrganizeInFolder w:val="0"/>
          <w:doNotUseLongFileNames/>
          <w:pixelsPerInch w:val="+123456789012345678901234567890"/>
          <w:targetScreenSz w:val="1920x1200"/>
          <w:saveSmartTagsAsXml w:val="false"/>
        </w:webSettings>"#;

        let settings = WebSettings::extract_from_xml(xml).unwrap();
        let serialized = settings.to_xml().unwrap();
        let reparsed = WebSettings::extract_from_xml(serialized.as_bytes()).unwrap();

        assert_eq!(reparsed, settings);
        assert!(serialized.contains("main &amp; detail"));
        assert!(serialized.contains("w:themeTint=\"0A\""));
        assert!(serialized.contains("w:themeShade=\"FF\""));
    }

    #[test]
    fn edits_and_clears_every_scalar_web_setting() {
        let mut settings = WebSettings::default();
        settings
            .set_encoding("utf&8")
            .set_optimize_for_browser(true)
            .set_rely_on_vml(false)
            .set_allow_png(true)
            .set_do_not_rely_on_css(false)
            .set_do_not_save_as_single_file(true)
            .set_do_not_organize_in_folder(false)
            .set_do_not_use_long_file_names(true);
        settings
            .set_pixels_per_inch(" +123456789012345678901234567890 ")
            .unwrap()
            .set_target_screen_size(TargetScreenSize::Pixels1800x1440)
            .set_save_smart_tags_as_xml(false);

        let serialized = settings.to_xml().unwrap();
        let reparsed = WebSettings::extract_from_xml(serialized.as_bytes()).unwrap();
        assert_eq!(reparsed, settings);
        assert_eq!(reparsed.encoding(), Some("utf&8"));
        assert_eq!(
            reparsed.pixels_per_inch(),
            Some("+123456789012345678901234567890")
        );
        assert_eq!(
            reparsed.target_screen_size(),
            Some(TargetScreenSize::Pixels1800x1440)
        );

        let previous_pixels = settings.pixels_per_inch().unwrap().to_owned();
        assert!(settings.set_pixels_per_inch("96.0").is_err());
        assert_eq!(settings.pixels_per_inch(), Some(previous_pixels.as_str()));

        settings
            .clear_encoding()
            .clear_optimize_for_browser()
            .clear_rely_on_vml()
            .clear_allow_png()
            .clear_do_not_rely_on_css()
            .clear_do_not_save_as_single_file()
            .clear_do_not_organize_in_folder()
            .clear_do_not_use_long_file_names()
            .clear_pixels_per_inch()
            .clear_target_screen_size()
            .clear_save_smart_tags_as_xml();
        assert_eq!(settings, WebSettings::default());
        assert_eq!(
            WebSettings::extract_from_xml(settings.to_xml().unwrap().as_bytes()).unwrap(),
            WebSettings::default()
        );
    }

    #[test]
    fn serialization_preserves_present_empty_top_level_containers() {
        let xml = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset/><w:divs/></w:webSettings>"#;
        let settings = WebSettings::extract_from_xml(xml).unwrap();
        let reparsed =
            WebSettings::extract_from_xml(settings.to_xml().unwrap().as_bytes()).unwrap();
        assert_eq!(reparsed, settings);
        assert!(reparsed.frameset().is_some());
        assert_eq!(reparsed.divs(), Some([].as_slice()));
    }

    #[test]
    fn serialization_rejects_excessive_recursive_nesting() {
        let mut frameset = Frameset::default();
        for _ in 0..=MAX_FRAMESET_NESTING {
            frameset = Frameset {
                children: vec![FramesetChild::Frameset(frameset)],
                ..Frameset::default()
            };
        }
        let settings = WebSettings {
            frameset: Some(frameset),
            ..WebSettings::default()
        };
        assert!(settings.to_xml().is_err());

        let mut div = HtmlDiv {
            id: "leaf".into(),
            ..HtmlDiv::default()
        };
        for id in 0..=MAX_FRAMESET_NESTING {
            div = HtmlDiv {
                id: id.to_string(),
                children: vec![div],
                ..HtmlDiv::default()
            };
        }
        let settings = WebSettings {
            divs: Some(vec![div]),
            ..WebSettings::default()
        };
        assert!(settings.to_xml().is_err());
    }
}
