//! PresentationML color-map parsing and resolution.

use crate::presentation_properties::metadata::is_presentationml_name;
use crate::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_ooxml_common::xml::{is_drawingml_name, unqualified_attribute_value};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

/// A role that can be mapped by a PresentationML color map.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Slot {
    /// Background color one.
    Background1,
    /// Text color one.
    Text1,
    /// Background color two.
    Background2,
    /// Text color two.
    Text2,
    /// Accent color one.
    Accent1,
    /// Accent color two.
    Accent2,
    /// Accent color three.
    Accent3,
    /// Accent color four.
    Accent4,
    /// Accent color five.
    Accent5,
    /// Accent color six.
    Accent6,
    /// Hyperlink color.
    Hyperlink,
    /// Followed-hyperlink color.
    FollowedHyperlink,
}

/// A color role defined by a DrawingML theme.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Role {
    /// Dark color one.
    Dark1,
    /// Light color one.
    Light1,
    /// Dark color two.
    Dark2,
    /// Light color two.
    Light2,
    /// Accent color one.
    Accent1,
    /// Accent color two.
    Accent2,
    /// Accent color three.
    Accent3,
    /// Accent color four.
    Accent4,
    /// Accent color five.
    Accent5,
    /// Accent color six.
    Accent6,
    /// Hyperlink color.
    Hyperlink,
    /// Followed-hyperlink color.
    FollowedHyperlink,
}

impl Role {
    /// Return the DrawingML theme color name.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value {
            "dk1" => Some(Self::Dark1),
            "lt1" => Some(Self::Light1),
            "dk2" => Some(Self::Dark2),
            "lt2" => Some(Self::Light2),
            "accent1" => Some(Self::Accent1),
            "accent2" => Some(Self::Accent2),
            "accent3" => Some(Self::Accent3),
            "accent4" => Some(Self::Accent4),
            "accent5" => Some(Self::Accent5),
            "accent6" => Some(Self::Accent6),
            "hlink" => Some(Self::Hyperlink),
            "folHlink" => Some(Self::FollowedHyperlink),
            _ => None,
        }
    }
}

/// A complete PresentationML color map.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Map {
    background1: Role,
    text1: Role,
    background2: Role,
    text2: Role,
    accent1: Role,
    accent2: Role,
    accent3: Role,
    accent4: Role,
    accent5: Role,
    accent6: Role,
    hyperlink: Role,
    followed_hyperlink: Role,
}

impl Map {
    /// Return the theme color role mapped from a PresentationML color slot.
    pub const fn color(&self, slot: Slot) -> Role {
        match slot {
            Slot::Background1 => self.background1,
            Slot::Text1 => self.text1,
            Slot::Background2 => self.background2,
            Slot::Text2 => self.text2,
            Slot::Accent1 => self.accent1,
            Slot::Accent2 => self.accent2,
            Slot::Accent3 => self.accent3,
            Slot::Accent4 => self.accent4,
            Slot::Accent5 => self.accent5,
            Slot::Accent6 => self.accent6,
            Slot::Hyperlink => self.hyperlink,
            Slot::FollowedHyperlink => self.followed_hyperlink,
        }
    }

    fn from_element(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        label: &str,
    ) -> Result<Self> {
        Ok(Self {
            background1: required_role(element, b"bg1", decoder, label)?,
            text1: required_role(element, b"tx1", decoder, label)?,
            background2: required_role(element, b"bg2", decoder, label)?,
            text2: required_role(element, b"tx2", decoder, label)?,
            accent1: required_role(element, b"accent1", decoder, label)?,
            accent2: required_role(element, b"accent2", decoder, label)?,
            accent3: required_role(element, b"accent3", decoder, label)?,
            accent4: required_role(element, b"accent4", decoder, label)?,
            accent5: required_role(element, b"accent5", decoder, label)?,
            accent6: required_role(element, b"accent6", decoder, label)?,
            hyperlink: required_role(element, b"hlink", decoder, label)?,
            followed_hyperlink: required_role(element, b"folHlink", decoder, label)?,
        })
    }
}

/// The color-map selection declared by a slide or slide layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Override {
    /// Use the color map defined by the owning slide master.
    Master,
    /// Use a color map declared directly by the slide or layout.
    Override(Map),
}

pub fn parse_master(xml: &[u8]) -> Result<Map> {
    let xml = process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut color_map = None;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::Invalid("color-map XML nesting is too deep".to_string())
                })?;
                if depth == 1 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, b"sldMaster", "slide master")?;
                    saw_root = true;
                } else if depth == 2
                    && is_presentationml_name(&namespace, element.name(), b"clrMap")
                {
                    store_color_map(
                        &mut color_map,
                        Map::from_element(&element, decoder, "slide-master color map")?,
                    )?;
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, b"sldMaster", "slide master")?;
                    saw_root = true;
                } else if depth == 1
                    && is_presentationml_name(&namespace, element.name(), b"clrMap")
                {
                    store_color_map(
                        &mut color_map,
                        Map::from_element(&element, decoder, "slide-master color map")?,
                    )?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid color-map XML nesting".to_string()))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || !saw_root {
        return Err(Error::Invalid(
            "unterminated slide-master color-map XML".to_string(),
        ));
    }
    color_map.ok_or_else(|| Error::Invalid("slide master is missing its color map".to_string()))
}

pub fn parse_override(xml: &[u8], root_name: &[u8], root_label: &str) -> Result<Option<Override>> {
    let xml = process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut saw_override = false;
    let mut override_depth = None;
    let mut mapping = None;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::Invalid("color-map XML nesting is too deep".to_string())
                })?;
                if depth == 1 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, root_name, root_label)?;
                    saw_root = true;
                } else if depth == 2
                    && is_presentationml_name(&namespace, element.name(), b"clrMapOvr")
                {
                    if saw_override {
                        return Err(Error::Invalid(format!(
                            "{root_label} has multiple color-map overrides"
                        )));
                    }
                    saw_override = true;
                    override_depth = Some(depth);
                } else if override_depth == Some(depth - 1) {
                    parse_override_mapping(&namespace, &element, decoder, &mut mapping)?;
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, root_name, root_label)?;
                    saw_root = true;
                } else if depth == 1
                    && is_presentationml_name(&namespace, element.name(), b"clrMapOvr")
                {
                    if saw_override {
                        return Err(Error::Invalid(format!(
                            "{root_label} has multiple color-map overrides"
                        )));
                    }
                    saw_override = true;
                } else if override_depth == Some(depth) {
                    parse_override_mapping(&namespace, &element, decoder, &mut mapping)?;
                }
            },
            Event::End(element) => {
                if override_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"clrMapOvr")
                {
                    override_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid color-map XML nesting".to_string()))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || !saw_root {
        return Err(Error::Invalid(format!(
            "unterminated {root_label} color-map XML"
        )));
    }
    if saw_override && mapping.is_none() {
        return Err(Error::Invalid(format!(
            "{root_label} color-map override has no mapping"
        )));
    }
    Ok(mapping)
}

fn require_root(
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_name: &[u8],
    label: &str,
) -> Result<()> {
    if is_presentationml_name(namespace, element.name(), expected_name) {
        Ok(())
    } else {
        Err(Error::Invalid(format!(
            "color-map XML must have a PresentationML {label} root"
        )))
    }
}

fn parse_override_mapping(
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    mapping: &mut Option<Override>,
) -> Result<()> {
    let value = if is_drawingml_name(namespace, element.name(), b"masterClrMapping") {
        Override::Master
    } else if is_drawingml_name(namespace, element.name(), b"overrideClrMapping") {
        Override::Override(Map::from_element(element, decoder, "color-map override")?)
    } else {
        return Ok(());
    };

    if mapping.replace(value).is_some() {
        return Err(Error::Invalid(
            "color-map override has multiple mappings".to_string(),
        ));
    }
    Ok(())
}

fn store_color_map(slot: &mut Option<Map>, value: Map) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::Invalid(
            "slide master has multiple color maps".to_string(),
        ));
    }
    Ok(())
}

fn required_role(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    label: &str,
) -> Result<Role> {
    let value = unqualified_attribute_value(element, name, decoder)?.ok_or_else(|| {
        Error::Invalid(format!(
            "{label} is missing its {} attribute",
            String::from_utf8_lossy(name)
        ))
    })?;
    Role::from_str(&value).ok_or_else(|| {
        Error::Invalid(format!(
            "{label} has unsupported {} value '{value}'",
            String::from_utf8_lossy(name)
        ))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    const MAP: &str = r#"bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
        accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4"
        accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink""#;

    #[test]
    fn parses_master_map_and_layout_override_by_namespace() {
        let master = parse_master(
            format!(r#"<p:sldMaster xmlns:p="{P}"><p:clrMap {MAP}/></p:sldMaster>"#).as_bytes(),
        )
        .unwrap();
        assert_eq!(master.color(Slot::Background1), Role::Light1);
        assert_eq!(master.color(Slot::Text1), Role::Dark1);

        let layout = parse_override(
            format!(
                r#"<q:sldLayout xmlns:q="{P}" xmlns:d="{A}"><q:clrMapOvr>
                <d:overrideClrMapping {MAP}/></q:clrMapOvr></q:sldLayout>"#
            )
            .as_bytes(),
            b"sldLayout",
            "slide layout",
        )
        .unwrap();
        assert_eq!(layout, Some(Override::Override(master)));
    }

    #[test]
    fn rejects_incomplete_or_duplicate_color_maps() {
        let incomplete =
            format!(r#"<p:sldMaster xmlns:p="{P}"><p:clrMap bg1="lt1"/></p:sldMaster>"#);
        assert!(parse_master(incomplete.as_bytes()).is_err());

        let duplicate = format!(
            r#"<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:clrMapOvr>
                <a:masterClrMapping/><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
        );
        assert!(parse_override(duplicate.as_bytes(), b"sld", "slide").is_err());

        let multiple_roots = format!(
            r#"<p:sldMaster xmlns:p="{P}"><p:clrMap {MAP}/></p:sldMaster>
            <p:sldMaster><p:clrMap {MAP}/></p:sldMaster>"#
        );
        assert!(parse_master(multiple_roots.as_bytes()).is_err());
    }

    #[test]
    fn supports_strict_color_map_namespaces() {
        const STRICT_P: &str = "http://purl.oclc.org/ooxml/presentationml/main";
        const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";

        let master = parse_master(
            format!(r#"<q:sldMaster xmlns:q="{STRICT_P}"><q:clrMap {MAP}/></q:sldMaster>"#)
                .as_bytes(),
        )
        .unwrap();
        assert_eq!(
            master.color(Slot::FollowedHyperlink),
            Role::FollowedHyperlink
        );

        let slide = parse_override(
            format!(
                r#"<q:sld xmlns:q="{STRICT_P}" xmlns:d="{STRICT_A}"><q:clrMapOvr>
                <d:masterClrMapping/></q:clrMapOvr></q:sld>"#
            )
            .as_bytes(),
            b"sld",
            "slide",
        )
        .unwrap();
        assert_eq!(slide, Some(Override::Master));
    }
}
