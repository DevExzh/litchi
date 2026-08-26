//! Presentation-part graph decoding.

use std::collections::HashSet;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use super::{
    MAX_SLIDES, expected_main_content_type, invalid, processed_xml, relationship_attribute,
    validate_content_type,
};
use crate::{Error, Result};

/// One entry in the presentation's ordered slide list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideReference {
    id: u32,
    relationship_id: String,
}

impl SlideReference {
    /// The stable `p:sldId@id` value.
    #[inline]
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// The relationship ID used by the presentation part.
    #[inline]
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
}

/// Borrowed, validated view of `/ppt/presentation.xml`.
#[derive(Clone, Copy)]
pub struct PresentationPart<'a> {
    part: &'a dyn Part,
}

impl<'a> PresentationPart<'a> {
    /// Wrap the package main document after validating its `PresentationML` type.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_package(package: &'a OpcPackage) -> Result<Self> {
        let part = package.main_document_part()?;
        Self::from_part(part)
    }

    /// Wrap an already resolved main part.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        if !expected_main_content_type(part.content_type()) {
            return Err(Error::ContentType {
                expected: format!(
                    "{}, {}, {}, {}, {}, or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_SLIDESHOW_MAIN,
                    ct::PML_TEMPLATE_MAIN,
                    ct::PML_PRES_MACRO_MAIN,
                    ct::PML_SLIDESHOW_MACRO_MAIN,
                    ct::PML_TEMPLATE_MACRO_MAIN,
                ),
                actual: part.content_type().to_string(),
            });
        }

        let xml = processed_xml(part)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(element) | Event::Empty(element)
                    if crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"presentation",
                    ) =>
                {
                    return Ok(Self { part });
                },
                Event::Decl(_) | Event::Comment(_) => {},
                Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
                _ => {
                    return Err(invalid(
                        "PresentationML main part does not have a presentation root",
                    ));
                },
            }
        }
    }

    /// The underlying OPC part.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Read the ordered slide references without hydrating slide parts.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_references(&self) -> Result<Vec<SlideReference>> {
        let xml = processed_xml(self.part)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut references = Vec::new();
        let mut slide_ids = HashSet::new();
        let mut relationship_ids = HashSet::new();
        let mut target_part_names = HashSet::new();
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(element) | Event::Empty(element)
                    if crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldId",
                    ) =>
                {
                    if references.len() == MAX_SLIDES {
                        return Err(Error::Limit {
                            resource: "presentation slide references",
                            limit: MAX_SLIDES,
                        });
                    }
                    let id = litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element,
                        b"id",
                        reader.decoder(),
                    )?
                    .ok_or_else(|| invalid("presentation slide reference lacks an id"))?;
                    let relationship_id = relationship_attribute(&element, &reader)?
                        .ok_or_else(|| invalid("presentation slide reference lacks r:id"))?;
                    let id = super::parse_u32(&id, "slide ID")?;
                    if !(256..=2_147_483_647).contains(&id) {
                        return Err(invalid(
                            "presentation slide reference id is outside 256..=2147483647",
                        ));
                    }
                    slide_ids
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "presentation slide reference IDs",
                            source,
                        })?;
                    if !slide_ids.insert(id) {
                        return Err(invalid("presentation slide reference ids must be unique"));
                    }
                    relationship_ids
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "presentation slide reference relationship IDs",
                            source,
                        })?;
                    if !relationship_ids.insert(relationship_id.clone()) {
                        return Err(invalid(
                            "presentation slide reference relationship ids must be unique",
                        ));
                    }
                    if let Some(relationship) = self.part.rels().get(&relationship_id)
                        && !relationship.is_external()
                    {
                        let target = relationship.target_partname()?;
                        let target_key = target.as_str().to_ascii_lowercase();
                        target_part_names
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "presentation slide reference target names",
                                source,
                            })?;
                        if !target_part_names.insert(target_key) {
                            return Err(invalid(
                                "presentation slide reference targets must be unique",
                            ));
                        }
                    }
                    references.push(SlideReference {
                        id,
                        relationship_id,
                    });
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(references)
    }

    /// Read the ordered relationship IDs declared by `p:sldMasterIdLst`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_master_references(&self) -> Result<Vec<String>> {
        let xml = processed_xml(self.part)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut references = Vec::new();
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(element) | Event::Empty(element)
                    if crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldMasterId",
                    ) =>
                {
                    references.push(
                        relationship_attribute(&element, &reader)?
                            .ok_or_else(|| invalid("presentation master reference lacks r:id"))?,
                    );
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(references)
    }

    /// Return the presentation slide size in EMUs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_size(&self) -> Result<(i64, i64)> {
        let xml = processed_xml(self.part)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        loop {
            match reader.read_event()? {
                Event::Start(element) | Event::Empty(element)
                    if element.local_name().as_ref() == b"sldSz" =>
                {
                    let width = litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element,
                        b"cx",
                        reader.decoder(),
                    )?
                    .ok_or_else(|| invalid("presentation slide size lacks cx"))?;
                    let height = litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element,
                        b"cy",
                        reader.decoder(),
                    )?
                    .ok_or_else(|| invalid("presentation slide size lacks cy"))?;
                    return Ok((
                        super::parse_i64(&width, "slide width")?,
                        super::parse_i64(&height, "slide height")?,
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Err(invalid("presentation lacks a slide-size element"))
    }

    /// Return relationship IDs for the presentation's slide masters.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[must_use]
    pub fn slide_master_relationships(&self) -> Vec<String> {
        self.part
            .rels()
            .iter()
            .filter(|relationship| {
                super::is_relationship_type(relationship.reltype(), rt::SLIDE_MASTER, "slideMaster")
            })
            .map(|relationship| relationship.r_id().to_string())
            .collect()
    }

    /// Validate that a related part has the expected content type.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn validate_related_part(&self, part: &dyn Part, expected: &str) -> Result<()> {
        validate_content_type(part, expected)
    }
}
