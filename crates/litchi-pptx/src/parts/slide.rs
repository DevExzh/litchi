//! Borrowed slide, layout, and master part views.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::Event;
use quick_xml::reader::{NsReader, Reader};

use super::{invalid, processed_xml, related_part_by_type, validate_content_type};
use crate::shape::Scene;
use crate::{Error, Result};

const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;

fn root_name(part: &dyn Part) -> Result<String> {
    let xml = processed_xml(part)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    loop {
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !crate::namespace::is_presentationml_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                ) {
                    return Err(invalid("PresentationML part has an invalid root namespace"));
                }
                return Ok(String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(|_| invalid("PresentationML root name is not UTF-8"))?);
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            _ => return Err(invalid("PresentationML part lacks an element root")),
        }
    }
}

fn c_sld_name(part: &dyn Part) -> Result<Option<String>> {
    let xml = processed_xml(part)?;
    crate::namespace::presentation_name(xml.as_ref())
}

/// Read the ordered `p:sldLayoutIdLst` relationship references owned by a
/// slide master. The OPC relationship collection may contain stale or
/// producer-private edges; the XML list is the semantic owner of the layout
/// inventory.
fn layout_relationship_ids(part: &dyn Part) -> Result<Vec<String>> {
    let xml = processed_xml(part)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut depth = 0usize;
    let mut in_list = false;
    let mut seen_list = false;
    let mut relationship_ids = Vec::new();

    loop {
        let (_namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-master XML nesting is too deep"))?;
                if depth == 2 && element.local_name().as_ref() == b"sldLayoutIdLst" {
                    if seen_list {
                        return Err(invalid("duplicate slide-layout ID list"));
                    }
                    seen_list = true;
                    in_list = true;
                } else if depth == 3 && in_list && element.local_name().as_ref() == b"sldLayoutId" {
                    relationship_ids.push(
                        crate::namespace::relationship_attribute_value(
                            &element,
                            b"id",
                            reader.decoder(),
                            reader.resolver(),
                        )?
                        .ok_or_else(|| {
                            invalid("slide-layout entry is missing its relationship ID")
                        })?,
                    );
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-master XML nesting is too deep"))?;
                if child_depth == 2 && element.local_name().as_ref() == b"sldLayoutIdLst" {
                    if seen_list {
                        return Err(invalid("duplicate slide-layout ID list"));
                    }
                    seen_list = true;
                } else if child_depth == 3
                    && in_list
                    && element.local_name().as_ref() == b"sldLayoutId"
                {
                    relationship_ids.push(
                        crate::namespace::relationship_attribute_value(
                            &element,
                            b"id",
                            reader.decoder(),
                            reader.resolver(),
                        )?
                        .ok_or_else(|| {
                            invalid("slide-layout entry is missing its relationship ID")
                        })?,
                    );
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected closing element in slide-master XML"));
                }
                if depth == 2 && element.local_name().as_ref() == b"sldLayoutIdLst" {
                    in_list = false;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 {
        return Err(invalid("unterminated slide-master XML"));
    }
    Ok(relationship_ids)
}

fn root_bool(part: &dyn Part, attribute: &[u8], field: &str, default: bool) -> Result<bool> {
    let xml = processed_xml(part)?;
    let mut reader = Reader::from_reader(xml.as_ref());
    loop {
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element) => {
                let value = litchi_ooxml_common::xml::unqualified_attribute_value(
                    &element,
                    attribute,
                    reader.decoder(),
                )?;
                return value.map_or(Ok(default), |value| super::parse_bool(&value, field));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            _ => return Err(invalid("PresentationML part lacks an element root")),
        }
    }
}

fn text_from_part(part: &dyn Part) -> Result<Option<String>> {
    let xml = processed_xml(part)?;
    let mut reader = Reader::from_reader(xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut in_text = false;
    let mut current = String::new();
    let mut value = String::new();
    loop {
        match reader.read_event()? {
            Event::Start(element) if element.local_name().as_ref() == b"t" => {
                in_text = true;
                current.clear();
            },
            Event::Text(text) if in_text => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let decoded = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if value
                    .len()
                    .saturating_add(current.len())
                    .saturating_add(decoded.len())
                    > MAX_TEXT_BYTES
                {
                    return Err(Error::Limit {
                        resource: "slide text",
                        limit: MAX_TEXT_BYTES,
                    });
                }
                current.push_str(&decoded);
            },
            Event::GeneralRef(reference) if in_text => {
                current.push_str(&litchi_ooxml_common::xml::decode_xml_reference(&reference)?);
            },
            Event::CData(text) if in_text => {
                current.push_str(
                    &text
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?,
                );
            },
            Event::End(element) if element.local_name().as_ref() == b"t" => {
                if !value.is_empty() {
                    value.push('\n');
                }
                value.push_str(&current);
                in_text = false;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok((!value.is_empty()).then_some(value))
}

/// Borrowed view of a PresentationML slide part.
#[derive(Clone, Copy)]
pub struct SlidePart<'a> {
    part: &'a dyn Part,
}

impl<'a> SlidePart<'a> {
    /// Validate and wrap a slide part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        validate_content_type(part, ct::PML_SLIDE)?;
        if root_name(part)? != "sld" {
            return Err(invalid("slide part does not have a p:sld root"));
        }
        Ok(Self { part })
    }

    /// The underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Producer-visible slide name, if present.
    pub fn name(&self) -> Result<String> {
        Ok(c_sld_name(self.part)?.unwrap_or_else(|| self.part.partname().to_string()))
    }

    /// Whether the slide is marked hidden by its root `show` attribute.
    pub fn is_hidden(&self) -> Result<bool> {
        Ok(!root_bool(self.part, b"show", "slide show", true)?)
    }

    /// Flatten DrawingML text runs in source order.
    pub fn text(&self) -> Result<String> {
        Ok(text_from_part(self.part)?.unwrap_or_default())
    }

    /// Build the bounded borrowed shape scene for this slide.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        Ok(Scene::read(self.part.blob())?)
    }

    /// Resolve ordinary DrawingML chart parts related to this slide.
    pub fn charts(&self, package: &'a OpcPackage) -> Result<Vec<crate::chart::Part<'a>>> {
        crate::chart::related(package, self.part)
    }

    /// Resolve Microsoft ChartEx parts related to this slide.
    pub fn chart_extensions(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Vec<crate::chart::extension::Part<'a>>> {
        crate::chart::extension::related(package, self.part)
    }

    /// Resolve the optional legacy comments list attached to this slide.
    pub fn comments(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Option<crate::comments::ListPart<'a>>> {
        let part = related_part_by_type(
            package,
            self.part,
            crate::comments::COMMENTS_REL,
            "comments",
            ct::PML_COMMENTS,
        )?;
        part.map(crate::comments::ListPart::from_part).transpose()
    }

    /// Resolve the slide's optional layout relationship.
    pub fn layout(&self, package: &'a OpcPackage) -> Result<Option<SlideLayoutPart<'a>>> {
        related_part_by_type(
            package,
            self.part,
            rt::SLIDE_LAYOUT,
            "slideLayout",
            ct::PML_SLIDE_LAYOUT,
        )?
        .map(SlideLayoutPart::from_part)
        .transpose()
    }
}

/// Borrowed view of a PresentationML slide-layout part.
#[derive(Clone, Copy)]
pub struct SlideLayoutPart<'a> {
    part: &'a dyn Part,
}

impl<'a> SlideLayoutPart<'a> {
    /// Validate and wrap a slide-layout part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        validate_content_type(part, ct::PML_SLIDE_LAYOUT)?;
        if root_name(part)? != "sldLayout" {
            return Err(invalid(
                "slide-layout part does not have a p:sldLayout root",
            ));
        }
        Ok(Self { part })
    }

    /// The underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Producer-visible layout name, if present.
    pub fn name(&self) -> Result<String> {
        Ok(c_sld_name(self.part)?.unwrap_or_else(|| self.part.partname().to_string()))
    }

    /// Layout kind token from `p:sldLayout@type`, if present.
    pub fn kind(&self) -> Result<Option<String>> {
        let xml = processed_xml(self.part)?;
        let mut reader = Reader::from_reader(xml.as_ref());
        loop {
            match reader.read_event()? {
                Event::Start(element) | Event::Empty(element) => {
                    return Ok(litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element,
                        b"type",
                        reader.decoder(),
                    )?);
                },
                Event::Decl(_) | Event::Comment(_) => {},
                _ => return Err(invalid("slide-layout part lacks an element root")),
            }
        }
    }

    /// Build the bounded borrowed shape scene for this layout.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        Ok(Scene::read(self.part.blob())?)
    }

    /// Read the optional theme override attached to this layout.
    pub fn theme_override(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Option<crate::shape::theme::Override>> {
        crate::shape::theme::package::load_override(package, self.part.partname().as_str())
    }

    /// Resolve the required slide-master relationship.
    pub fn master(&self, package: &'a OpcPackage) -> Result<SlideMasterPart<'a>> {
        let part = related_part_by_type(
            package,
            self.part,
            rt::SLIDE_MASTER,
            "slideMaster",
            ct::PML_SLIDE_MASTER,
        )?
        .ok_or_else(|| invalid("slide layout lacks its slide-master relationship"))?;
        SlideMasterPart::from_part(part)
    }
}

/// Borrowed view of a PresentationML slide-master part.
#[derive(Clone, Copy)]
pub struct SlideMasterPart<'a> {
    part: &'a dyn Part,
}

impl<'a> SlideMasterPart<'a> {
    /// Validate and wrap a slide-master part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        validate_content_type(part, ct::PML_SLIDE_MASTER)?;
        if root_name(part)? != "sldMaster" {
            return Err(invalid(
                "slide-master part does not have a p:sldMaster root",
            ));
        }
        Ok(Self { part })
    }

    /// The underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Producer-visible master name, if present.
    pub fn name(&self) -> Result<String> {
        Ok(c_sld_name(self.part)?.unwrap_or_else(|| self.part.partname().to_string()))
    }

    /// Whether the master is marked preserved.
    pub fn is_preserved(&self) -> Result<bool> {
        root_bool(self.part, b"preserve", "slide-master preserve", false)
    }

    /// Build the bounded borrowed shape scene for this master.
    pub fn shapes(&self) -> Result<Scene<'a>> {
        Ok(Scene::read(self.part.blob())?)
    }

    /// Read the theme reached from this slide master.
    pub fn theme(
        &self,
        package: &'a OpcPackage,
    ) -> Result<Option<crate::shape::theme::ThemeSummary>> {
        let part = related_part_by_type(package, self.part, rt::THEME, "theme", ct::OFC_THEME)?;
        part.map(|part| crate::shape::theme::part::Part::from_part(part)?.read())
            .transpose()
    }

    /// Resolve the slide layouts listed by this master in XML order.
    pub fn layouts(&self, package: &'a OpcPackage) -> Result<Vec<SlideLayoutPart<'a>>> {
        let relationship_ids = layout_relationship_ids(self.part)?;
        let mut layouts = Vec::with_capacity(relationship_ids.len());
        for relationship_id in relationship_ids {
            let relationship = self.part.rels().get(&relationship_id).ok_or_else(|| {
                Error::Relationship(format!(
                    "slide master references missing slide-layout relationship '{relationship_id}'"
                ))
            })?;
            if relationship.is_external() {
                return Err(Error::Relationship(
                    "slide-layout relationship must be internal".into(),
                ));
            }
            if !super::is_relationship_type(relationship.reltype(), rt::SLIDE_LAYOUT, "slideLayout")
            {
                return Err(Error::Relationship(format!(
                    "relationship '{relationship_id}' is not a slide-layout relationship"
                )));
            }
            let target = relationship.target_partname()?;
            let part = package.get_part(&target)?;
            validate_content_type(part, ct::PML_SLIDE_LAYOUT)?;
            layouts.push(SlideLayoutPart::from_part(part)?);
        }
        Ok(layouts)
    }
}
