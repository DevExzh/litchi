//! WordprocessingML web-settings support.

use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// Scalar settings from a Word `webSettings.xml` part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WebSettings {
    frameset: Option<Frameset>,
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

impl WebSettings {
    /// Return the root frameset definition, if present.
    pub fn frameset(&self) -> Option<&Frameset> {
        self.frameset.as_ref()
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
            let value = value.trim();
            if !is_xml_integer(value) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "invalid pixels-per-inch value '{value}'"
                )));
            }
            set_once(
                &mut settings.pixels_per_inch,
                value.to_owned(),
                "pixelsPerInch",
            )
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

fn is_xml_integer(value: &str) -> bool {
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit())
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
}
