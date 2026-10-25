//! Presentation-part graph decoding.

use std::collections::HashSet;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{
    MAX_SLIDES, expected_main_content_type, invalid, processed_xml, relationship_attribute,
    validate_content_type,
};
use crate::leniency::{PresentationDefect, ToleranceReport};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const MAX_SLIDE_MASTER_REFERENCES: usize = 4_096;
const MIN_NATIVE_ID: u32 = 2_147_483_648;
const MIN_SLIDE_SIZE_EMU: i64 = 914_400;
const MAX_SLIDE_SIZE_EMU: i64 = 51_206_400;

#[derive(Clone, Copy, PartialEq, Eq)]
enum PresentationNamespace {
    Transitional,
    Strict,
}

fn presentation_namespace(namespace: &ResolveResult<'_>) -> Option<PresentationNamespace> {
    match namespace {
        ResolveResult::Bound(Namespace(value))
            if *value == crate::namespace::PRESENTATIONML_NAMESPACE =>
        {
            Some(PresentationNamespace::Transitional)
        },
        ResolveResult::Bound(Namespace(value))
            if *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
        {
            Some(PresentationNamespace::Strict)
        },
        _ => None,
    }
}

fn same_presentation_namespace(
    root: Option<PresentationNamespace>,
    namespace: &ResolveResult<'_>,
) -> bool {
    let Some(root) = root else {
        return false;
    };
    Some(root) == presentation_namespace(namespace)
}

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
    /// Slide-list defects `PowerPoint` and LibreOffice read past are
    /// repaired (audit finding 6.7): a `p:sldId` repeating an earlier entry's
    /// slide ID, relationship ID, or target part is skipped (first wins), and
    /// foreign children, attributes, or character data inside `p:sldIdLst`
    /// are ignored. The owned handles record these repairs at open
    /// ([`crate::Package::tolerance_report`],
    /// [`crate::SourceBackedPresentation::tolerance_report`]); use
    /// [`Self::slide_references_with_tolerance`] to collect them here.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_references(&self) -> Result<Vec<SlideReference>> {
        self.slide_references_with_tolerance(&mut ToleranceReport::default())
    }

    /// [`Self::slide_references`], appending each repair to `tolerance`
    /// stamped with the presentation part name and the skipped entry's
    /// slide ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_references_with_tolerance(
        &self,
        tolerance: &mut ToleranceReport,
    ) -> Result<Vec<SlideReference>> {
        let mut repairs = ToleranceReport::default();
        let references = self.scan_slide_references(Some(&mut repairs))?;
        repairs.locate_in_part(self.part.partname().as_str());
        tolerance.append(&repairs);
        Ok(references)
    }

    /// The pre-leniency slide-list contract, for lossless edit planners:
    /// every defect [`Self::slide_references`] repairs is an error instead.
    ///
    /// # Errors
    ///
    /// Returns an error if the slide list is malformed or carries a defect.
    pub(crate) fn slide_references_strict(&self) -> Result<Vec<SlideReference>> {
        self.scan_slide_references(None)
    }

    /// Scan `p:sldIdLst`, repairing into `tolerance` or, without one,
    /// rejecting every repairable defect.
    fn scan_slide_references(
        &self,
        mut tolerance: Option<&mut ToleranceReport>,
    ) -> Result<Vec<SlideReference>> {
        let xml = processed_xml(self.part)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut references = Vec::new();
        let mut slide_ids = HashSet::new();
        let mut relationship_ids = HashSet::new();
        let mut target_part_names = HashSet::new();
        let mut depth = 0usize;
        let mut roots = 0usize;
        let mut lists = 0usize;
        let mut list_depth = None;
        // Depth of a foreign `p:sldIdLst` child whose subtree is skipped.
        let mut skipping: Option<usize> = None;
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            if let Some(skip_depth) = skipping {
                match event {
                    Event::Start(_) => {
                        depth = depth
                            .checked_add(1)
                            .ok_or_else(|| invalid("presentation XML depth overflow"))?;
                    },
                    Event::End(_) => {
                        depth = depth
                            .checked_sub(1)
                            .ok_or_else(|| invalid("presentation XML depth underflow"))?;
                        if depth == skip_depth {
                            skipping = None;
                        }
                    },
                    Event::Eof => break,
                    _ => {},
                }
                continue;
            }
            match event {
                Event::Start(element) => {
                    let is_presentation = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"presentation",
                    );
                    let is_list = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldIdLst",
                    );
                    let is_slide = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldId",
                    );
                    if depth == 0 {
                        roots = roots
                            .checked_add(1)
                            .ok_or_else(|| invalid("presentation XML root count overflow"))?;
                        if !is_presentation {
                            return Err(invalid(
                                "PresentationML main part does not have a presentation root",
                            ));
                        }
                    }
                    if is_list {
                        if depth != 1 {
                            return Err(invalid(
                                "presentation slide-ID list must be a direct root child",
                            ));
                        }
                        lists = lists
                            .checked_add(1)
                            .ok_or_else(|| invalid("presentation slide-ID list count overflow"))?;
                        if lists != 1 {
                            return Err(invalid("presentation has multiple slide-ID lists"));
                        }
                        reject_slide_id_list_attributes(&element, tolerance.as_deref_mut())?;
                        list_depth = Some(depth + 1);
                    } else if is_slide {
                        if list_depth != Some(depth) {
                            return Err(invalid(
                                "presentation slide reference is not a direct slide-ID-list child",
                            ));
                        }
                        append_slide_reference(
                            &mut references,
                            &mut slide_ids,
                            &mut relationship_ids,
                            &mut target_part_names,
                            self.part,
                            &element,
                            &reader,
                            tolerance.as_deref_mut(),
                        )?;
                    } else if list_depth == Some(depth) {
                        foreign_slide_list_content(tolerance.as_deref_mut(), "a non-slide child")?;
                        skipping = Some(depth);
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("presentation XML depth overflow"))?;
                },
                Event::Empty(element) => {
                    let is_presentation = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"presentation",
                    );
                    let is_list = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldIdLst",
                    );
                    let is_slide = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldId",
                    );
                    if depth == 0 {
                        roots = roots
                            .checked_add(1)
                            .ok_or_else(|| invalid("presentation XML root count overflow"))?;
                        if !is_presentation {
                            return Err(invalid(
                                "PresentationML main part does not have a presentation root",
                            ));
                        }
                    }
                    if is_list {
                        if depth != 1 || lists != 0 {
                            return Err(invalid(
                                "presentation has an additional or nested slide-ID list",
                            ));
                        }
                        lists = lists
                            .checked_add(1)
                            .ok_or_else(|| invalid("presentation slide-ID list count overflow"))?;
                        reject_slide_id_list_attributes(&element, tolerance.as_deref_mut())?;
                    } else if is_slide {
                        if list_depth != Some(depth) {
                            return Err(invalid(
                                "presentation slide reference is not a direct slide-ID-list child",
                            ));
                        }
                        append_slide_reference(
                            &mut references,
                            &mut slide_ids,
                            &mut relationship_ids,
                            &mut target_part_names,
                            self.part,
                            &element,
                            &reader,
                            tolerance.as_deref_mut(),
                        )?;
                    } else if list_depth == Some(depth) {
                        foreign_slide_list_content(tolerance.as_deref_mut(), "a non-slide child")?;
                    }
                },
                Event::End(element) => {
                    if crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldIdLst",
                    ) {
                        if list_depth != Some(depth) {
                            return Err(invalid(
                                "presentation slide-ID list is nested or not opened",
                            ));
                        }
                        list_depth = None;
                    } else if crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldId",
                    ) && list_depth == Some(depth)
                    {
                        return Err(invalid(
                            "presentation slide-ID entry unexpectedly remains open",
                        ));
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("presentation XML depth underflow"))?;
                },
                Event::Text(value) if list_depth == Some(depth) => {
                    if !value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .trim()
                        .is_empty()
                    {
                        foreign_slide_list_content(tolerance.as_deref_mut(), "text content")?;
                    }
                },
                Event::CData(_) | Event::GeneralRef(_) if list_depth == Some(depth) => {
                    foreign_slide_list_content(tolerance.as_deref_mut(), "unsupported content")?;
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if roots != 1 || depth != 0 || list_depth.is_some() || lists > 1 {
            return Err(invalid(
                "presentation must contain at most one direct slide-ID list",
            ));
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
        let mut depth = 0usize;
        let mut roots = 0usize;
        let mut root_namespace = None;
        let mut lists = 0usize;
        let mut list_depth: Option<usize> = None;
        let mut reference_depth: Option<usize> = None;
        let mut opaque_depth: Option<usize> = None;
        let mut extension_seen = false;
        let mut root_extension_seen = false;
        let mut references = Vec::new();
        let mut relationship_ids = HashSet::new();
        let mut native_ids = HashSet::new();
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(_) if opaque_depth.is_some() => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("presentation XML nesting is too deep"))?;
                },
                Event::Empty(_) if opaque_depth.is_some() => {},
                Event::End(_) if opaque_depth.is_some() => {
                    if opaque_depth == Some(depth) {
                        opaque_depth = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("presentation XML nesting underflow"))?;
                },
                Event::Text(_) | Event::CData(_) | Event::GeneralRef(_)
                    if opaque_depth.is_some() => {},
                Event::Start(element) => {
                    let is_root = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"presentation",
                    );
                    let is_list = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldMasterIdLst",
                    );
                    let is_reference = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldMasterId",
                    );
                    let is_extension = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"extLst",
                    );
                    if is_extension
                        && ((reference_depth == Some(depth))
                            || (reference_depth.is_none() && depth == 1))
                        && !same_presentation_namespace(root_namespace, &namespace)
                    {
                        return Err(invalid("presentation extension list has mixed conformance"));
                    }
                    if reference_depth == Some(depth) && is_extension {
                        if extension_seen {
                            return Err(invalid(
                                "presentation slide-master reference has duplicate extension lists",
                            ));
                        }
                        extension_seen = true;
                        opaque_depth = Some(
                            depth
                                .checked_add(1)
                                .ok_or_else(|| invalid("presentation XML nesting is too deep"))?,
                        );
                    } else if reference_depth.is_none() && depth == 1 && is_extension {
                        if root_extension_seen {
                            return Err(invalid("presentation has duplicate root extension lists"));
                        }
                        root_extension_seen = true;
                        opaque_depth = Some(
                            depth
                                .checked_add(1)
                                .ok_or_else(|| invalid("presentation XML nesting is too deep"))?,
                        );
                    } else {
                        if reference_depth.is_some() {
                            return Err(invalid(
                                "presentation slide-master reference contains an unexpected child",
                            ));
                        }
                        if depth == 0 {
                            roots = roots
                                .checked_add(1)
                                .ok_or_else(|| invalid("presentation root count overflow"))?;
                            if !is_root {
                                return Err(invalid(
                                    "PresentationML main part does not have a presentation root",
                                ));
                            }
                            root_namespace =
                                Some(presentation_namespace(&namespace).ok_or_else(|| {
                                    invalid(
                                        "PresentationML main part has an invalid root namespace",
                                    )
                                })?);
                        }
                        if is_list {
                            if depth != 1 || lists != 0 {
                                return Err(invalid(
                                    "presentation slide-master ID list must be one direct root child",
                                ));
                            }
                            if !same_presentation_namespace(root_namespace, &namespace) {
                                return Err(invalid(
                                    "presentation slide-master ID list has mixed conformance",
                                ));
                            }
                            lists = lists.checked_add(1).ok_or_else(|| {
                                invalid("presentation slide-master list count overflow")
                            })?;
                            reject_master_id_list_attributes(&element)?;
                            list_depth =
                                Some(depth.checked_add(1).ok_or_else(|| {
                                    invalid("presentation XML nesting is too deep")
                                })?);
                        } else if is_reference {
                            if list_depth != Some(depth) {
                                return Err(invalid(
                                    "presentation slide-master reference is not a direct list child",
                                ));
                            }
                            if !same_presentation_namespace(root_namespace, &namespace) {
                                return Err(invalid(
                                    "presentation slide-master reference has mixed conformance",
                                ));
                            }
                            append_slide_master_reference(
                                &mut references,
                                &mut relationship_ids,
                                &mut native_ids,
                                &element,
                                &reader,
                            )?;
                            reference_depth =
                                Some(depth.checked_add(1).ok_or_else(|| {
                                    invalid("presentation XML nesting is too deep")
                                })?);
                            extension_seen = false;
                        } else if element.local_name().as_ref() == b"sldMasterIdLst" {
                            return Err(invalid(
                                "presentation slide-master ID list has an invalid namespace",
                            ));
                        } else if depth == 1 && element.local_name().as_ref() == b"extLst" {
                            return Err(invalid(
                                "presentation extension list has an invalid namespace",
                            ));
                        } else if list_depth == Some(depth) {
                            return Err(invalid(
                                "presentation slide-master ID list contains an unexpected child",
                            ));
                        }
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("presentation XML nesting is too deep"))?;
                },
                Event::Empty(element) => {
                    let is_root = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"presentation",
                    );
                    let is_list = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldMasterIdLst",
                    );
                    let is_reference = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldMasterId",
                    );
                    let is_extension = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"extLst",
                    );
                    if is_extension
                        && ((reference_depth == Some(depth))
                            || (reference_depth.is_none() && depth == 1))
                        && !same_presentation_namespace(root_namespace, &namespace)
                    {
                        return Err(invalid("presentation extension list has mixed conformance"));
                    }
                    if reference_depth == Some(depth) && is_extension {
                        if extension_seen {
                            return Err(invalid(
                                "presentation slide-master reference has duplicate extension lists",
                            ));
                        }
                        extension_seen = true;
                    } else if reference_depth.is_none() && depth == 1 && is_extension {
                        if root_extension_seen {
                            return Err(invalid("presentation has duplicate root extension lists"));
                        }
                        root_extension_seen = true;
                    } else {
                        if reference_depth.is_some() {
                            return Err(invalid(
                                "presentation slide-master reference contains an unexpected child",
                            ));
                        }
                        if depth == 0 {
                            roots = roots
                                .checked_add(1)
                                .ok_or_else(|| invalid("presentation root count overflow"))?;
                            if !is_root {
                                return Err(invalid(
                                    "PresentationML main part does not have a presentation root",
                                ));
                            }
                            root_namespace =
                                Some(presentation_namespace(&namespace).ok_or_else(|| {
                                    invalid(
                                        "PresentationML main part has an invalid root namespace",
                                    )
                                })?);
                        }
                        if is_list {
                            if depth != 1 || lists != 0 {
                                return Err(invalid(
                                    "presentation slide-master ID list must be one direct root child",
                                ));
                            }
                            if !same_presentation_namespace(root_namespace, &namespace) {
                                return Err(invalid(
                                    "presentation slide-master ID list has mixed conformance",
                                ));
                            }
                            lists = lists.checked_add(1).ok_or_else(|| {
                                invalid("presentation slide-master list count overflow")
                            })?;
                            reject_master_id_list_attributes(&element)?;
                        } else if is_reference {
                            if list_depth != Some(depth) {
                                return Err(invalid(
                                    "presentation slide-master reference is not a direct list child",
                                ));
                            }
                            if !same_presentation_namespace(root_namespace, &namespace) {
                                return Err(invalid(
                                    "presentation slide-master reference has mixed conformance",
                                ));
                            }
                            append_slide_master_reference(
                                &mut references,
                                &mut relationship_ids,
                                &mut native_ids,
                                &element,
                                &reader,
                            )?;
                        } else if element.local_name().as_ref() == b"sldMasterIdLst" {
                            return Err(invalid(
                                "presentation slide-master ID list has an invalid namespace",
                            ));
                        } else if depth == 1 && element.local_name().as_ref() == b"extLst" {
                            return Err(invalid(
                                "presentation extension list has an invalid namespace",
                            ));
                        } else if list_depth == Some(depth) {
                            return Err(invalid(
                                "presentation slide-master ID list contains an unexpected child",
                            ));
                        }
                    }
                },
                Event::End(element) => {
                    if depth == 0 {
                        return Err(invalid("presentation XML nesting underflow"));
                    }
                    if reference_depth == Some(depth) {
                        if !same_presentation_namespace(root_namespace, &namespace)
                            || !crate::namespace::is_presentationml_name(
                                &namespace,
                                element.name(),
                                b"sldMasterId",
                            )
                        {
                            return Err(invalid(
                                "presentation slide-master reference has an invalid closing element",
                            ));
                        }
                        reference_depth = None;
                        extension_seen = false;
                    }
                    if list_depth == Some(depth)
                        && same_presentation_namespace(root_namespace, &namespace)
                        && crate::namespace::is_presentationml_name(
                            &namespace,
                            element.name(),
                            b"sldMasterIdLst",
                        )
                    {
                        list_depth = None;
                    }
                    depth -= 1;
                },
                Event::Text(value) if reference_depth == Some(depth) => {
                    if !value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .trim()
                        .is_empty()
                    {
                        return Err(invalid(
                            "presentation slide-master references contain text content",
                        ));
                    }
                },
                Event::Text(value) if list_depth == Some(depth) => {
                    if !value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .trim()
                        .is_empty()
                    {
                        return Err(invalid(
                            "presentation slide-master ID list contains text content",
                        ));
                    }
                },
                Event::CData(_) | Event::GeneralRef(_) if reference_depth == Some(depth) => {
                    return Err(invalid(
                        "presentation slide-master references contain unsupported content",
                    ));
                },
                Event::CData(_) | Event::GeneralRef(_) if list_depth == Some(depth) => {
                    return Err(invalid(
                        "presentation slide-master ID list contains unsupported content",
                    ));
                },
                Event::DocType(_) => return Err(invalid("DTDs are rejected")),
                Event::Eof => break,
                _ => {},
            }
        }
        if roots != 1
            || depth != 0
            || list_depth.is_some()
            || reference_depth.is_some()
            || opaque_depth.is_some()
        {
            return Err(invalid(
                "presentation must contain one balanced presentation tree",
            ));
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
        let mut depth = 0usize;
        let mut roots = 0usize;
        let mut root_namespace = None;
        let mut size_depth: Option<usize> = None;
        let mut opaque_depth: Option<usize> = None;
        let mut root_extension_seen = false;
        let mut size = None;
        loop {
            let (namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(_) if opaque_depth.is_some() => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("presentation XML nesting is too deep"))?;
                },
                Event::Empty(_) if opaque_depth.is_some() => {},
                Event::End(_) if opaque_depth.is_some() => {
                    if opaque_depth == Some(depth) {
                        opaque_depth = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("presentation XML nesting underflow"))?;
                },
                Event::Text(_) | Event::CData(_) | Event::GeneralRef(_)
                    if opaque_depth.is_some() => {},
                Event::Start(element) => {
                    let is_root = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"presentation",
                    );
                    let is_size = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldSz",
                    );
                    let is_extension = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"extLst",
                    );
                    if depth == 1
                        && is_extension
                        && !same_presentation_namespace(root_namespace, &namespace)
                    {
                        return Err(invalid("presentation extension list has mixed conformance"));
                    }
                    if depth == 1 && is_extension {
                        if root_extension_seen {
                            return Err(invalid("presentation has duplicate root extension lists"));
                        }
                        root_extension_seen = true;
                        opaque_depth = Some(
                            depth
                                .checked_add(1)
                                .ok_or_else(|| invalid("presentation XML nesting is too deep"))?,
                        );
                    } else {
                        if size_depth.is_some() {
                            return Err(invalid("presentation slide-size element must be empty"));
                        }
                        if depth == 0 {
                            roots = roots
                                .checked_add(1)
                                .ok_or_else(|| invalid("presentation root count overflow"))?;
                            if !is_root {
                                return Err(invalid(
                                    "PresentationML main part does not have a presentation root",
                                ));
                            }
                            root_namespace =
                                Some(presentation_namespace(&namespace).ok_or_else(|| {
                                    invalid(
                                        "PresentationML main part has an invalid root namespace",
                                    )
                                })?);
                        }
                        if is_size {
                            if depth != 1 {
                                return Err(invalid(
                                    "presentation slide-size element must be a direct root child",
                                ));
                            }
                            if !same_presentation_namespace(root_namespace, &namespace) {
                                return Err(invalid(
                                    "presentation slide-size element has mixed conformance",
                                ));
                            }
                            if size.is_some() {
                                return Err(invalid(
                                    "presentation has multiple slide-size elements",
                                ));
                            }
                            size = Some(parse_slide_size(&element, reader.decoder())?);
                            size_depth =
                                Some(depth.checked_add(1).ok_or_else(|| {
                                    invalid("presentation XML nesting is too deep")
                                })?);
                        } else if element.local_name().as_ref() == b"sldSz" {
                            return Err(invalid(
                                "presentation slide-size element has an invalid namespace",
                            ));
                        } else if depth == 1 && element.local_name().as_ref() == b"extLst" {
                            return Err(invalid(
                                "presentation extension list has an invalid namespace",
                            ));
                        }
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("presentation XML nesting is too deep"))?;
                },
                Event::Empty(element) => {
                    let is_root = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"presentation",
                    );
                    let is_size = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"sldSz",
                    );
                    let is_extension = crate::namespace::is_presentationml_name(
                        &namespace,
                        element.name(),
                        b"extLst",
                    );
                    if depth == 1
                        && is_extension
                        && !same_presentation_namespace(root_namespace, &namespace)
                    {
                        return Err(invalid("presentation extension list has mixed conformance"));
                    }
                    if depth == 1 && is_extension {
                        if root_extension_seen {
                            return Err(invalid("presentation has duplicate root extension lists"));
                        }
                        root_extension_seen = true;
                    } else {
                        if size_depth.is_some() {
                            return Err(invalid("presentation slide-size element must be empty"));
                        }
                        if depth == 0 {
                            roots = roots
                                .checked_add(1)
                                .ok_or_else(|| invalid("presentation root count overflow"))?;
                            if !is_root {
                                return Err(invalid(
                                    "PresentationML main part does not have a presentation root",
                                ));
                            }
                            root_namespace =
                                Some(presentation_namespace(&namespace).ok_or_else(|| {
                                    invalid(
                                        "PresentationML main part has an invalid root namespace",
                                    )
                                })?);
                        }
                        if is_size {
                            if depth != 1 {
                                return Err(invalid(
                                    "presentation slide-size element must be a direct root child",
                                ));
                            }
                            if !same_presentation_namespace(root_namespace, &namespace) {
                                return Err(invalid(
                                    "presentation slide-size element has mixed conformance",
                                ));
                            }
                            if size.is_some() {
                                return Err(invalid(
                                    "presentation has multiple slide-size elements",
                                ));
                            }
                            size = Some(parse_slide_size(&element, reader.decoder())?);
                        } else if element.local_name().as_ref() == b"sldSz" {
                            return Err(invalid(
                                "presentation slide-size element has an invalid namespace",
                            ));
                        } else if depth == 1 && element.local_name().as_ref() == b"extLst" {
                            return Err(invalid(
                                "presentation extension list has an invalid namespace",
                            ));
                        }
                    }
                },
                Event::End(element) => {
                    if depth == 0 {
                        return Err(invalid("presentation XML nesting underflow"));
                    }
                    if size_depth == Some(depth) {
                        if !same_presentation_namespace(root_namespace, &namespace)
                            || !crate::namespace::is_presentationml_name(
                                &namespace,
                                element.name(),
                                b"sldSz",
                            )
                        {
                            return Err(invalid(
                                "presentation slide-size element has an invalid closing element",
                            ));
                        }
                        size_depth = None;
                    }
                    depth -= 1;
                },
                Event::Text(value) if size_depth == Some(depth) => {
                    if !value
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .trim()
                        .is_empty()
                    {
                        return Err(invalid("presentation slide-size element contains text"));
                    }
                },
                Event::CData(_) | Event::GeneralRef(_) if size_depth == Some(depth) => {
                    return Err(invalid("presentation slide-size element contains content"));
                },
                Event::DocType(_) => return Err(invalid("DTDs are rejected")),
                Event::Eof => break,
                _ => {},
            }
        }
        if roots != 1 || depth != 0 || size_depth.is_some() || opaque_depth.is_some() {
            return Err(invalid("presentation XML is not a balanced tree"));
        }
        size.ok_or_else(|| invalid("presentation lacks a slide-size element"))
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

fn parse_slide_size(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<(i64, i64)> {
    let width = litchi_ooxml_common::xml::unqualified_attribute_value(element, b"cx", decoder)?
        .ok_or_else(|| invalid("presentation slide size lacks cx"))?;
    let height = litchi_ooxml_common::xml::unqualified_attribute_value(element, b"cy", decoder)?
        .ok_or_else(|| invalid("presentation slide size lacks cy"))?;
    let width = width
        .parse::<i64>()
        .ok()
        .filter(|value| (MIN_SLIDE_SIZE_EMU..=MAX_SLIDE_SIZE_EMU).contains(value))
        .ok_or_else(|| invalid("presentation slide width is outside the allowed range"))?;
    let height = height
        .parse::<i64>()
        .ok()
        .filter(|value| (MIN_SLIDE_SIZE_EMU..=MAX_SLIDE_SIZE_EMU).contains(value))
        .ok_or_else(|| invalid("presentation slide height is outside the allowed range"))?;
    Ok((width, height))
}

fn reject_master_id_list_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            return Err(invalid(
                "presentation slide-master ID list permits only namespace declarations",
            ));
        }
    }
    Ok(())
}

fn append_slide_master_reference(
    references: &mut Vec<String>,
    relationship_ids: &mut HashSet<String>,
    native_ids: &mut HashSet<u32>,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<()> {
    if references.len() >= MAX_SLIDE_MASTER_REFERENCES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::References),
            observed: references.len().checked_add(1).unwrap_or(usize::MAX),
            maximum: MAX_SLIDE_MASTER_REFERENCES,
        });
    }
    let native_id = parse_native_id(element, reader.decoder())?;
    if native_id.is_some_and(|id| native_ids.contains(&id)) {
        return Err(invalid(
            "presentation slide-master native IDs must be unique",
        ));
    }
    let relationship_id = relationship_attribute(element, reader)?
        .ok_or_else(|| invalid("presentation slide-master reference lacks r:id"))?;
    if relationship_ids.contains(&relationship_id) {
        return Err(invalid(
            "presentation slide-master relationship IDs must be unique",
        ));
    }
    relationship_ids
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "presentation slide-master relationship IDs",
            source,
        })?;
    if native_id.is_some() {
        native_ids
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "presentation slide-master native IDs",
                source,
            })?;
    }
    references
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "presentation slide-master references",
            source,
        })?;
    if let Some(native_id) = native_id {
        native_ids.insert(native_id);
    }
    relationship_ids.insert(relationship_id.clone());
    references.push(relationship_id);
    Ok(())
}

fn parse_native_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<u32>> {
    let Some(value) =
        litchi_ooxml_common::xml::unqualified_attribute_value(element, b"id", decoder)?
    else {
        return Ok(None);
    };
    let value = value
        .parse::<u32>()
        .ok()
        .filter(|value| *value >= MIN_NATIVE_ID)
        .ok_or_else(|| invalid("presentation slide-master native ID is invalid"))?;
    Ok(Some(value))
}

fn reject_slide_id_list_attributes(
    element: &BytesStart<'_>,
    tolerance: Option<&mut ToleranceReport>,
) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            let Some(tolerance) = tolerance else {
                return Err(invalid(
                    "presentation slide-ID list permits only namespace declarations",
                ));
            };
            // One record per list: the attributes are ignored together.
            tolerance.record(PresentationDefect::SlideListContent, 0);
            break;
        }
    }
    Ok(())
}

/// Ignore and record foreign `p:sldIdLst` content, or reject it strictly.
fn foreign_slide_list_content(tolerance: Option<&mut ToleranceReport>, what: &str) -> Result<()> {
    match tolerance {
        Some(tolerance) => {
            tolerance.record(PresentationDefect::SlideListContent, 0);
            Ok(())
        },
        None => Err(invalid(format!(
            "presentation slide-ID list contains {what}"
        ))),
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "the slide-list scan threads its uniqueness sets and repair sink"
)]
fn append_slide_reference(
    references: &mut Vec<SlideReference>,
    slide_ids: &mut HashSet<u32>,
    relationship_ids: &mut HashSet<String>,
    target_part_names: &mut HashSet<String>,
    part: &dyn Part,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    tolerance: Option<&mut ToleranceReport>,
) -> Result<()> {
    if references.len() == MAX_SLIDES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::References),
            observed: references.len().checked_add(1).unwrap_or(usize::MAX),
            maximum: MAX_SLIDES,
        });
    }
    let id =
        litchi_ooxml_common::xml::unqualified_attribute_value(element, b"id", reader.decoder())?
            .ok_or_else(|| invalid("presentation slide reference lacks an id"))?;
    let relationship_id = relationship_attribute(element, reader)?
        .ok_or_else(|| invalid("presentation slide reference lacks r:id"))?;
    let id = super::parse_u32(&id, "slide ID")?;
    if !(256..=2_147_483_647).contains(&id) {
        return Err(invalid(
            "presentation slide reference id is outside 256..=2147483647",
        ));
    }
    let target_key = match part.rels().get(&relationship_id) {
        Some(relationship) if !relationship.is_external() => Some(
            relationship
                .target_partname()?
                .as_str()
                .to_ascii_lowercase(),
        ),
        _ => None,
    };
    // The first entry carrying a slide ID, relationship ID, or target wins;
    // a later repeat is skipped whole so no set keeps half of it.
    let duplicate = if slide_ids.contains(&id) {
        Some((
            PresentationDefect::DuplicateSlideId,
            "presentation slide reference ids must be unique",
        ))
    } else if relationship_ids.contains(&relationship_id) {
        Some((
            PresentationDefect::DuplicateSlideRelationship,
            "presentation slide reference relationship ids must be unique",
        ))
    } else if target_key
        .as_ref()
        .is_some_and(|target| target_part_names.contains(target))
    {
        Some((
            PresentationDefect::DuplicateSlideTarget,
            "presentation slide reference targets must be unique",
        ))
    } else {
        None
    };
    if let Some((defect, message)) = duplicate {
        let Some(tolerance) = tolerance else {
            return Err(invalid(message));
        };
        tolerance.record(defect, id);
        return Ok(());
    }
    slide_ids
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "presentation slide reference IDs",
            source,
        })?;
    slide_ids.insert(id);
    relationship_ids
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "presentation slide reference relationship IDs",
            source,
        })?;
    relationship_ids.insert(relationship_id.clone());
    if let Some(target_key) = target_key {
        target_part_names
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "presentation slide reference target names",
                source,
            })?;
        target_part_names.insert(target_key);
    }
    references
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "presentation slide references",
            source,
        })?;
    references.push(SlideReference {
        id,
        relationship_id,
    });
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused low-level part tests use literal XML fixtures"
    )]

    use super::PresentationPart;
    use crate::Error;
    use litchi_opc::PackURI;
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::part::BlobPart;
    use std::fmt::Write as _;

    const TRANSITIONAL_PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const TRANSITIONAL_REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    fn presentation_part(
        body: &str,
        presentation_namespace: &str,
        relationship_namespace: &str,
    ) -> BlobPart {
        let xml = format!(
            r#"<p:presentation xmlns:p="{presentation_namespace}" xmlns:r="{relationship_namespace}" xmlns:s="{STRICT_PML}" xmlns:t="{TRANSITIONAL_PML}" xmlns:x="urn:foreign">{body}</p:presentation>"#
        );
        BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            ct::PML_PRESENTATION_MAIN.to_owned(),
            xml.into_bytes(),
        )
    }

    #[test]
    fn slide_master_references_accept_transitional_and_strict_direct_children() {
        for (presentation_namespace, relationship_namespace) in [
            (TRANSITIONAL_PML, TRANSITIONAL_REL),
            (STRICT_PML, STRICT_REL),
        ] {
            let part = presentation_part(
                r#"<p:sldMasterIdLst><p:sldMasterId r:id="rIdMaster"/></p:sldMasterIdLst>"#,
                presentation_namespace,
                relationship_namespace,
            );
            let presentation = PresentationPart::from_part(&part).unwrap();
            assert_eq!(
                presentation.slide_master_references().unwrap(),
                vec!["rIdMaster"]
            );
        }
    }

    #[test]
    fn slide_master_references_reject_mixed_root_and_core_namespaces() {
        for (presentation_namespace, body) in [
            (
                STRICT_PML,
                r#"<t:sldMasterIdLst><t:sldMasterId r:id="rIdMaster"/></t:sldMasterIdLst>"#,
            ),
            (
                TRANSITIONAL_PML,
                r#"<s:sldMasterIdLst><s:sldMasterId r:id="rIdMaster"/></s:sldMasterIdLst>"#,
            ),
        ] {
            let part = presentation_part(body, presentation_namespace, TRANSITIONAL_REL);
            assert!(
                PresentationPart::from_part(&part)
                    .unwrap()
                    .slide_master_references()
                    .is_err(),
                "core master references must use the root conformance"
            );
        }
    }

    #[test]
    fn core_presentation_scanners_reject_undeclared_p_prefix() {
        let references = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            ct::PML_PRESENTATION_MAIN.to_owned(),
            br#"<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId r:id="rIdMaster"/></p:sldMasterIdLst></p:presentation>"#.to_vec(),
        );
        assert!(
            PresentationPart::from_part(&references)
                .unwrap()
                .slide_master_references()
                .is_err()
        );

        let size = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            ct::PML_PRESENTATION_MAIN.to_owned(),
            br#"<p:presentation><p:sldSz cx="914400" cy="914401"/></p:presentation>"#.to_vec(),
        );
        assert!(
            PresentationPart::from_part(&size)
                .unwrap()
                .slide_size()
                .is_err()
        );
    }

    #[test]
    fn slide_master_references_reject_namespace_spoof_nested_and_duplicate_entries() {
        let cases = [
            r#"<p:sldMasterIdLst><x:sldMasterId r:id="rIdForeign"/></p:sldMasterIdLst>"#,
            r#"<p:sldMasterIdLst><p:wrapper><p:sldMasterId r:id="rIdNested"/></p:wrapper></p:sldMasterIdLst>"#,
            r#"<p:sldMasterIdLst><p:sldMasterId r:id="secret-id"/><p:sldMasterId r:id="secret-id"/></p:sldMasterIdLst>"#,
        ];
        for body in cases {
            let part = presentation_part(body, TRANSITIONAL_PML, TRANSITIONAL_REL);
            let presentation = PresentationPart::from_part(&part).unwrap();
            let error = presentation
                .slide_master_references()
                .expect_err("invalid master list must be refused")
                .to_string();
            assert!(!error.contains("secret-id"));
        }
    }

    #[test]
    fn slide_master_references_bound_collection_before_the_4097th_entry() {
        let mut body = String::new();
        for index in 0..=4_096 {
            write!(&mut body, r#"<p:sldMasterId r:id="rId{index}"/>"#).unwrap();
        }
        let body = format!(r#"<p:sldMasterIdLst>{body}</p:sldMasterIdLst>"#);
        let part = presentation_part(&body, TRANSITIONAL_PML, TRANSITIONAL_REL);
        let presentation = PresentationPart::from_part(&part).unwrap();
        assert!(matches!(
            presentation.slide_master_references(),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn slide_master_references_allow_one_opaque_extension_list_per_entry() {
        let part = presentation_part(
            r#"<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"><p:extLst><p:future><p:sldMasterId r:id="ignored"/><p:sldMasterIdLst/></p:future></p:extLst></p:sldMasterId></p:sldMasterIdLst>"#,
            TRANSITIONAL_PML,
            TRANSITIONAL_REL,
        );
        let presentation = PresentationPart::from_part(&part).unwrap();
        assert_eq!(
            presentation.slide_master_references().unwrap(),
            vec!["rIdMaster"]
        );

        let duplicate = presentation_part(
            r#"<p:sldMasterIdLst><p:sldMasterId r:id="rIdMaster"><p:extLst/><p:extLst/></p:sldMasterId></p:sldMasterIdLst>"#,
            TRANSITIONAL_PML,
            TRANSITIONAL_REL,
        );
        assert!(
            PresentationPart::from_part(&duplicate)
                .unwrap()
                .slide_master_references()
                .is_err()
        );
    }

    #[test]
    fn slide_master_references_ignore_opaque_root_extension_payloads() {
        let part = presentation_part(
            r#"<p:extLst><p:future><p:sldMasterId r:id="ignored"/><p:sldMasterIdLst><p:sldMasterId r:id="also-ignored"/></p:sldMasterIdLst></p:future></p:extLst><p:sldMasterIdLst><p:sldMasterId r:id="rIdMaster"/></p:sldMasterIdLst>"#,
            TRANSITIONAL_PML,
            TRANSITIONAL_REL,
        );
        assert_eq!(
            PresentationPart::from_part(&part)
                .unwrap()
                .slide_master_references()
                .unwrap(),
            vec!["rIdMaster"]
        );

        let duplicate = presentation_part(
            r#"<p:extLst/><p:extLst/><p:sldMasterIdLst><p:sldMasterId r:id="rIdMaster"/></p:sldMasterIdLst>"#,
            TRANSITIONAL_PML,
            TRANSITIONAL_REL,
        );
        assert!(
            PresentationPart::from_part(&duplicate)
                .unwrap()
                .slide_master_references()
                .is_err()
        );
    }

    #[test]
    fn slide_master_references_validate_optional_native_ids() {
        let valid = presentation_part(
            r#"<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdOne"/><p:sldMasterId id="2147483649" r:id="rIdTwo"/></p:sldMasterIdLst>"#,
            TRANSITIONAL_PML,
            TRANSITIONAL_REL,
        );
        assert_eq!(
            PresentationPart::from_part(&valid)
                .unwrap()
                .slide_master_references()
                .unwrap(),
            vec!["rIdOne", "rIdTwo"]
        );

        for body in [
            r#"<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdOne"/><p:sldMasterId id="2147483648" r:id="rIdTwo"/></p:sldMasterIdLst>"#,
            r#"<p:sldMasterIdLst><p:sldMasterId id="2147483647" r:id="rIdOne"/></p:sldMasterIdLst>"#,
        ] {
            let part = presentation_part(body, TRANSITIONAL_PML, TRANSITIONAL_REL);
            assert!(
                PresentationPart::from_part(&part)
                    .unwrap()
                    .slide_master_references()
                    .is_err()
            );
        }
    }

    #[test]
    fn slide_size_accepts_strict_schema_and_rejects_duplicates_or_nested_sizes() {
        let part = presentation_part(
            r#"<p:sldSz cx="9144000" cy="6858000"/>"#,
            STRICT_PML,
            STRICT_REL,
        );
        assert_eq!(
            PresentationPart::from_part(&part)
                .unwrap()
                .slide_size()
                .unwrap(),
            (9_144_000, 6_858_000)
        );

        let extension = presentation_part(
            r#"<p:extLst><p:future><p:sldSz cx="914399" cy="914399"/></p:future></p:extLst><p:sldSz cx="914400" cy="914401"/>"#,
            TRANSITIONAL_PML,
            TRANSITIONAL_REL,
        );
        assert_eq!(
            PresentationPart::from_part(&extension)
                .unwrap()
                .slide_size()
                .unwrap(),
            (914_400, 914_401)
        );

        for body in [
            r#"<p:sldSz cx="914400" cy="914401"/><p:sldSz cx="914402" cy="914403"/>"#,
            r#"<p:wrapper><p:sldSz cx="914400" cy="914401"/></p:wrapper>"#,
            r#"<x:sldSz cx="914400" cy="914401"/>"#,
        ] {
            let part = presentation_part(body, TRANSITIONAL_PML, TRANSITIONAL_REL);
            assert!(
                PresentationPart::from_part(&part)
                    .unwrap()
                    .slide_size()
                    .is_err(),
                "schema-invalid slide size must be refused"
            );
        }

        let duplicate_extension = presentation_part(
            r#"<p:extLst/><p:extLst/><p:sldSz cx="914400" cy="914401"/>"#,
            TRANSITIONAL_PML,
            TRANSITIONAL_REL,
        );
        assert!(
            PresentationPart::from_part(&duplicate_extension)
                .unwrap()
                .slide_size()
                .is_err()
        );
    }

    #[test]
    fn slide_size_rejects_mixed_root_and_core_namespaces() {
        for (presentation_namespace, body) in [
            (STRICT_PML, r#"<t:sldSz cx="914400" cy="914401"/>"#),
            (TRANSITIONAL_PML, r#"<s:sldSz cx="914400" cy="914401"/>"#),
        ] {
            let part = presentation_part(body, presentation_namespace, TRANSITIONAL_REL);
            assert!(
                PresentationPart::from_part(&part)
                    .unwrap()
                    .slide_size()
                    .is_err(),
                "slide size must use the root conformance"
            );
        }
    }

    #[test]
    fn slide_size_rejects_out_of_range_and_non_numeric_dimensions_without_echoing_input() {
        for body in [
            r#"<p:sldSz cx="0" cy="2"/>"#,
            r#"<p:sldSz cx="-1" cy="2"/>"#,
            r#"<p:sldSz cx="not-a-width" cy="2"/>"#,
            r#"<p:sldSz cx="1" cy="0"/>"#,
            r#"<p:sldSz cx="1" cy="-2"/>"#,
            r#"<p:sldSz cx="1" cy="not-a-height"/>"#,
            r#"<p:sldSz cx="914399" cy="914400"/>"#,
            r#"<p:sldSz cx="51206401" cy="914400"/>"#,
            r#"<p:sldSz cx="914400" cy="914399"/>"#,
            r#"<p:sldSz cx="914400" cy="51206401"/>"#,
        ] {
            let part = presentation_part(body, TRANSITIONAL_PML, TRANSITIONAL_REL);
            let error = PresentationPart::from_part(&part)
                .unwrap()
                .slide_size()
                .expect_err("invalid slide dimensions must be refused")
                .to_string();
            assert!(!error.contains("not-a-width"));
            assert!(!error.contains("not-a-height"));
        }
    }
}
