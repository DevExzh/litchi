#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
use crate::error::{Error, Result};
use crate::header_footer::Kind;
use crate::section::Start;
use litchi_ooxml_common::xml_name::is_ncname;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::fmt::Write;

use super::super::borders;
use super::super::model::{
    AuthoredChild, AuthoredChildRaw, ChapterSep, Color, Display, Endnotes, Footnotes, GridType,
    LineNumberRestart, NamespaceBinding, NoteNumberRestart, NotePos, OffsetFrom, PageNumberFormat,
    PageOrientation, SectionColumn, SectionColumns, SectionDocumentGrid,
    SectionHeaderFooterReference, SectionLineNumbering, SectionPageNumbering, SectionPaperSource,
    SectionProperties, SectionTextDirection, SectionVerticalAlignment, Style, ZOrder,
};
use super::package::{write_reference, write_references};

pub(super) const TRANSITIONAL_WORD_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
pub(super) const TRANSITIONAL_RELATIONSHIP_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_RELATIONSHIP_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

#[derive(Debug, Clone)]
pub(super) struct NamespacePlan {
    pub(super) word_prefix: Option<String>,
    pub(super) word_attribute_prefix: String,
    pub(super) relationship_prefix: String,
    pub(super) word_namespace: String,
    pub(super) relationship_namespace: String,
    pub(super) declare_word: bool,
    pub(super) declare_word_attribute: bool,
    pub(super) declare_relationship: bool,
}

impl NamespacePlan {
    pub(super) fn word_qname(&self, local: &str) -> String {
        self.word_prefix
            .as_deref()
            .map_or_else(|| local.to_owned(), |prefix| format!("{prefix}:{local}"))
    }

    pub(super) fn word_attribute_qname(&self, local: &str) -> String {
        format!("{}:{local}", self.word_attribute_prefix)
    }

    pub(super) fn relationship_qname(&self, local: &str) -> String {
        format!("{}:{local}", self.relationship_prefix)
    }

    pub(super) fn shadows_word_prefix(&self, attribute: &str, prefix: &str) -> bool {
        namespace_declaration_shadows(
            attribute,
            prefix,
            &self.word_namespace,
            &[TRANSITIONAL_WORD_NAMESPACE, STRICT_WORD_NAMESPACE],
        )
    }

    pub(super) fn shadows_relationship_prefix(&self, attribute: &str) -> bool {
        namespace_declaration_shadows(
            attribute,
            &self.relationship_prefix,
            &self.relationship_namespace,
            &[
                TRANSITIONAL_RELATIONSHIP_NAMESPACE,
                STRICT_RELATIONSHIP_NAMESPACE,
            ],
        )
    }
}

fn namespace_declaration_shadows(
    attribute: &str,
    prefix: &str,
    expected_uri: &str,
    detached_uris: &[&str],
) -> bool {
    let start = format!("xmlns:{prefix}=\"");
    let Some(uri) = attribute
        .strip_prefix(&start)
        .and_then(|value| value.strip_suffix('"'))
    else {
        return false;
    };
    if expected_uri.is_empty() {
        !detached_uris.contains(&uri)
    } else {
        uri != expected_uri
    }
}

fn namespace_plan(
    section: &SectionProperties,
    rels: Option<&super::super::super::relmap::RelationshipMapper>,
) -> Result<NamespacePlan> {
    if !section.root_namespace_resolved {
        let Some(word_prefix) = section.root_prefix.clone() else {
            return Err(Error::InvalidFormat(
                "unprefixed detached section namespace cannot be regenerated safely".into(),
            ));
        };
        let relationship_needed = !section.headers.is_empty()
            || !section.footers.is_empty()
            || section.printer_settings_relationship_id.is_some()
            || rels.is_some_and(|rels| {
                rels.get_header_id().is_some() || rels.get_footer_id().is_some()
            });
        let relationship_binding = section.namespace_bindings.iter().find(|binding| {
            binding.prefix.is_some() && is_relationship_namespace_uri(&binding.uri)
        });
        let (relationship_prefix, relationship_namespace) = if let Some(binding) =
            relationship_binding
        {
            (
                binding.prefix.clone().ok_or_else(|| {
                    Error::InvalidFormat("relationship namespace requires a named prefix".into())
                })?,
                binding.uri.clone(),
            )
        } else {
            if relationship_needed
                && (word_prefix == "r"
                    || section
                        .namespace_bindings
                        .iter()
                        .any(|binding| binding.prefix.as_deref() == Some("r")))
            {
                return Err(Error::InvalidFormat(
                    "detached section has a conflicting relationship prefix".into(),
                ));
            }
            ("r".to_owned(), String::new())
        };
        return Ok(NamespacePlan {
            word_prefix: Some(word_prefix.clone()),
            word_attribute_prefix: word_prefix,
            relationship_prefix,
            word_namespace: String::new(),
            relationship_namespace,
            declare_word: false,
            declare_word_attribute: false,
            declare_relationship: false,
        });
    }
    let bindings = &section.namespace_bindings;
    let word_binding = |prefix: &str| {
        bindings.iter().find(|binding| {
            binding.prefix.as_deref() == Some(prefix) && is_word_namespace_uri(&binding.uri)
        })
    };
    let first_word = bindings
        .iter()
        .find(|binding| is_word_namespace_uri(&binding.uri));
    let root_word_binding = if section.root_default_namespace {
        bindings
            .iter()
            .find(|binding| binding.prefix.is_none() && is_word_namespace_uri(&binding.uri))
    } else {
        section.root_prefix.as_deref().and_then(word_binding)
    };
    let word_namespace = root_word_binding
        .or(first_word)
        .map_or(TRANSITIONAL_WORD_NAMESPACE, |binding| binding.uri.as_str())
        .to_owned();
    let first_named_word = bindings
        .iter()
        .find(|binding| binding.prefix.is_some() && binding.uri == word_namespace);
    let generated_word_prefix = allocate_prefix(bindings, "w")?;
    let word_prefix = if section.root_default_namespace {
        None
    } else if let Some(prefix) = section.root_prefix.as_deref() {
        if root_word_binding.is_some()
            || !bindings
                .iter()
                .any(|binding| binding.prefix.as_deref() == Some(prefix))
        {
            Some(prefix.to_owned())
        } else {
            first_named_word
                .and_then(|binding| binding.prefix.clone())
                .or_else(|| Some(generated_word_prefix.clone()))
        }
    } else {
        first_named_word
            .and_then(|binding| binding.prefix.clone())
            .or_else(|| Some(generated_word_prefix.clone()))
    };
    let word_attribute_prefix =
        if let Some(prefix) = first_named_word.and_then(|binding| binding.prefix.clone()) {
            prefix
        } else if let Some(prefix) = word_prefix.clone() {
            prefix
        } else {
            allocate_prefix_avoiding(bindings, "w", std::iter::empty::<&str>())?
        };
    let relationship_binding = bindings.iter().find(|binding| {
        binding.prefix.as_deref() == Some("r") && is_relationship_namespace_uri(&binding.uri)
    });
    let root_relationship = bindings.iter().find(|binding| {
        is_relationship_namespace_uri(&binding.uri)
            && ((word_namespace == STRICT_WORD_NAMESPACE
                && binding.uri == STRICT_RELATIONSHIP_NAMESPACE)
                || (word_namespace != STRICT_WORD_NAMESPACE
                    && binding.uri == TRANSITIONAL_RELATIONSHIP_NAMESPACE))
    });
    let first_relationship = bindings
        .iter()
        .find(|binding| binding.prefix.is_some() && is_relationship_namespace_uri(&binding.uri));
    let relationship_namespace = relationship_binding
        .or(root_relationship)
        .or(first_relationship)
        .map_or(
            if word_namespace == STRICT_WORD_NAMESPACE {
                STRICT_RELATIONSHIP_NAMESPACE
            } else {
                TRANSITIONAL_RELATIONSHIP_NAMESPACE
            },
            |binding| binding.uri.as_str(),
        )
        .to_owned();
    let relationship_prefix = if let Some(prefix) = relationship_binding
        .or(root_relationship)
        .or(first_relationship)
        .and_then(|binding| binding.prefix.clone())
    {
        prefix
    } else {
        allocate_prefix_avoiding(
            bindings,
            "r",
            [word_prefix.as_deref(), Some(word_attribute_prefix.as_str())]
                .into_iter()
                .flatten(),
        )?
    };
    let declare_word = word_prefix.as_deref().is_some_and(|prefix| {
        !bindings.iter().any(|binding| {
            binding.prefix.as_deref() == Some(prefix) && binding.uri == word_namespace
        })
    });
    let declare_word_attribute = !bindings.iter().any(|binding| {
        binding.prefix.as_deref() == Some(word_attribute_prefix.as_str())
            && binding.uri == word_namespace
    }) && word_prefix.as_deref()
        != Some(word_attribute_prefix.as_str());
    let relationship_needed = !section.headers.is_empty()
        || !section.footers.is_empty()
        || section.printer_settings_relationship_id.is_some()
        || rels
            .is_some_and(|rels| rels.get_header_id().is_some() || rels.get_footer_id().is_some());
    let declare_relationship = relationship_needed
        && !bindings.iter().any(|binding| {
            binding.prefix.as_deref() == Some(relationship_prefix.as_str())
                && binding.uri == relationship_namespace
        });
    Ok(NamespacePlan {
        word_prefix,
        word_attribute_prefix,
        relationship_prefix,
        word_namespace,
        relationship_namespace,
        declare_word,
        declare_word_attribute,
        declare_relationship,
    })
}

fn allocate_prefix(bindings: &[NamespaceBinding], base: &str) -> Result<String> {
    if !bindings
        .iter()
        .any(|binding| binding.prefix.as_deref() == Some(base))
    {
        return Ok(base.to_owned());
    }
    for index in 1..=u16::MAX {
        let candidate = format!("{base}{index}");
        if !bindings
            .iter()
            .any(|binding| binding.prefix.as_deref() == Some(candidate.as_str()))
        {
            return Ok(candidate);
        }
    }
    Err(Error::InvalidFormat(
        "unable to allocate a collision-free XML namespace prefix".into(),
    ))
}

fn allocate_prefix_avoiding<'a>(
    bindings: &[NamespaceBinding],
    base: &str,
    forbidden: impl IntoIterator<Item = &'a str>,
) -> Result<String> {
    let forbidden = forbidden
        .into_iter()
        .collect::<std::collections::HashSet<_>>();
    let available = |candidate: &str| {
        !forbidden.contains(candidate)
            && !bindings
                .iter()
                .any(|binding| binding.prefix.as_deref() == Some(candidate))
    };
    if available(base) {
        return Ok(base.to_owned());
    }
    for index in 1..=u16::MAX {
        let candidate = format!("{base}{index}");
        if available(&candidate) {
            return Ok(candidate);
        }
    }
    Err(Error::InvalidFormat(
        "unable to allocate a collision-free XML namespace prefix".into(),
    ))
}

fn is_relationship_namespace_uri(uri: &str) -> bool {
    uri == TRANSITIONAL_RELATIONSHIP_NAMESPACE || uri == STRICT_RELATIONSHIP_NAMESPACE
}
impl SectionProperties {
    pub(crate) fn from_xml(xml: &str) -> Result<Self> {
        let metadata = root_metadata(xml)?;
        let children = direct_children(xml)?;
        let mut properties = Self {
            namespace_bindings: metadata.bindings.clone(),
            root_prefix: metadata.prefix.clone(),
            root_default_namespace: metadata.root_default_namespace,
            root_namespace_resolved: metadata.namespace_resolved,
            root_attributes: metadata.attributes.clone(),
            ..Self::default()
        };
        let mut seen = std::collections::HashSet::new();
        let mut last_rank = None;
        for (name, raw) in children {
            if let Some(rank) = section_child_rank(&name) {
                if last_rank.is_some_and(|last_rank| rank < last_rank) {
                    return Err(Error::InvalidFormat(format!(
                        "section property '{name}' is out of schema order"
                    )));
                }
                last_rank = Some(rank);
            }
            if section_child_rank(&name).is_some()
                && !seen.insert(name.clone())
                && !matches!(name.as_str(), "headerReference" | "footerReference")
            {
                return Err(Error::InvalidFormat(format!(
                    "section properties contain duplicate '{name}'"
                )));
            }
            if is_leaf_content_child(&name) {
                validate_leaf_content(&raw)?;
            }
            properties
                .authored_children
                .push(authored_child(&metadata, &name, &raw)?);
            let preserved_attributes = if name == "@raw" || name.starts_with("@foreign:") {
                Vec::new()
            } else {
                preserved_attribute_tokens(&metadata, &raw, &name)?
            };
            let original_on_off =
                if matches!(name.as_str(), "formProt" | "titlePg" | "bidi" | "rtlGutter") {
                    Some(parse_on_off(&metadata, &raw)?)
                } else {
                    None
                };
            properties.authored_child_raw.push(AuthoredChildRaw {
                raw: raw.clone(),
                preserved_attributes,
                original_on_off,
            });
            match name.as_str() {
                "headerReference" => properties
                    .headers
                    .push(parse_header_footer(&metadata, &raw)?),
                "footerReference" => properties
                    .footers
                    .push(parse_header_footer(&metadata, &raw)?),
                "footnotePr" => properties.footnotes = Some(parse_footnotes(&metadata, &raw)?),
                "endnotePr" => properties.endnotes = Some(parse_endnotes(&metadata, &raw)?),
                "type" => {
                    let value = required_attr(&metadata, &raw, b"val")?;
                    properties.start_type = Some(Start::from_xml(&value).ok_or_else(|| {
                        Error::InvalidFormat(format!("invalid section type '{value}'"))
                    })?);
                },
                "pgSz" => {
                    let attrs = attributes(&metadata, &raw)?;
                    if let Some(value) = attr(&attrs, "w") {
                        properties.page_width = parse_u32(value, "page width")?;
                    }
                    if let Some(value) = attr(&attrs, "h") {
                        properties.page_height = parse_u32(value, "page height")?;
                    }
                    if let Some(value) = attr(&attrs, "orient") {
                        properties.orientation = PageOrientation::parse(value)?;
                    }
                },
                "pgMar" => {
                    let attrs = attributes(&metadata, &raw)?;
                    assign_u32(&attrs, "top", &mut properties.margin_top)?;
                    assign_u32(&attrs, "bottom", &mut properties.margin_bottom)?;
                    assign_u32(&attrs, "left", &mut properties.margin_left)?;
                    assign_u32(&attrs, "right", &mut properties.margin_right)?;
                    assign_u32(&attrs, "header", &mut properties.header_distance)?;
                    assign_u32(&attrs, "footer", &mut properties.footer_distance)?;
                    assign_u32(&attrs, "gutter", &mut properties.gutter)?;
                },
                "pgNumType" => {
                    properties.page_numbering = Some(parse_page_numbering(&metadata, &raw)?)
                },
                "paperSrc" => {
                    let attrs = attributes(&metadata, &raw)?;
                    properties.paper_source = Some(SectionPaperSource {
                        first: attr(&attrs, "first")
                            .map(|value| parse_u32(value, "first paper source"))
                            .transpose()?,
                        other: attr(&attrs, "other")
                            .map(|value| parse_u32(value, "other paper source"))
                            .transpose()?,
                    });
                },
                "pgBorders" => properties.page_borders = Some(parse_page_borders(&metadata, &raw)?),
                "lnNumType" => {
                    properties.line_numbering = Some(parse_line_numbering(&metadata, &raw)?);
                },
                "cols" => properties.columns = Some(parse_columns(&metadata, &raw)?),
                "formProt" => properties.form_protection = parse_on_off(&metadata, &raw)?,
                "vAlign" => {
                    properties.vertical_alignment = Some(SectionVerticalAlignment::parse(
                        &required_attr(&metadata, &raw, b"val")?,
                    )?);
                },
                "titlePg" => properties.title_page = parse_on_off(&metadata, &raw)?,
                "textDirection" => {
                    properties.text_direction = Some(SectionTextDirection::parse(&required_attr(
                        &metadata, &raw, b"val",
                    )?)?);
                },
                "bidi" => properties.bidirectional = parse_on_off(&metadata, &raw)?,
                "rtlGutter" => properties.rtl_gutter = parse_on_off(&metadata, &raw)?,
                "docGrid" => properties.document_grid = Some(parse_grid(&metadata, &raw)?),
                "printerSettings" => {
                    let id = relationship_id(&metadata, &raw, "printer-settings")?;
                    properties.printer_settings_relationship_id = Some(id);
                },
                _ => properties.preserved_unknown_children.push(raw),
            }
        }
        properties.validate()?;
        Ok(properties)
    }

    pub(crate) fn write_xml(
        &self,
        xml: &mut String,
        rels: Option<&super::super::super::relmap::RelationshipMapper>,
    ) -> Result<()> {
        self.validate()?;
        let plan = namespace_plan(self, rels)?;
        write_root_open(self, xml, &plan)?;
        if self.authored_children.is_empty() {
            if !self.headers.is_empty() || !self.suppress_managed_headers {
                write_references(xml, "headerReference", &self.headers, rels, true, &plan)?;
            }
            if !self.footers.is_empty() || !self.suppress_managed_footers {
                write_references(xml, "footerReference", &self.footers, rels, false, &plan)?;
            }
            for index in 0..18 {
                write_field(self, xml, rels, &plan, index)?;
            }
            for child in &self.preserved_unknown_children {
                xml.push_str(child);
            }
        } else {
            write_authored_children(self, xml, rels, &plan)?;
        }
        write!(xml, "</{}>", plan.word_qname("sectPr"))
            .map_err(|error| Error::Xml(error.to_string()))?;
        Ok(())
    }
}

fn write_root_open(
    section: &SectionProperties,
    xml: &mut String,
    plan: &NamespacePlan,
) -> Result<()> {
    let root_name = plan.word_qname("sectPr");
    write!(xml, "<{root_name}").map_err(|error| Error::Xml(error.to_string()))?;
    for binding in &section.namespace_bindings {
        if let Some(prefix) = &binding.prefix {
            write!(xml, " xmlns:{prefix}=\"{}\"", escape(&binding.uri))
                .map_err(|error| Error::Xml(error.to_string()))?;
        } else {
            write!(xml, " xmlns=\"{}\"", escape(&binding.uri))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
    }
    for attribute in &section.root_attributes {
        write!(xml, " {attribute}").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if plan.declare_word {
        let prefix = plan
            .word_prefix
            .as_deref()
            .ok_or_else(|| Error::InvalidFormat("missing generated Word prefix".into()))?;
        write!(xml, " xmlns:{prefix}=\"{}\"", escape(&plan.word_namespace))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if plan.declare_word_attribute {
        write!(
            xml,
            " xmlns:{}=\"{}\"",
            plan.word_attribute_prefix,
            escape(&plan.word_namespace)
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if plan.declare_relationship {
        write!(
            xml,
            " xmlns:{}=\"{}\"",
            plan.relationship_prefix,
            escape(&plan.relationship_namespace)
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push('>');
    Ok(())
}

fn authored_rank(child: &AuthoredChild) -> u8 {
    match child {
        AuthoredChild::Header(_) | AuthoredChild::Footer(_) => 0,
        AuthoredChild::FootnotePr => 1,
        AuthoredChild::EndnotePr => 2,
        AuthoredChild::Type => 3,
        AuthoredChild::PageSize => 4,
        AuthoredChild::PageMargins => 5,
        AuthoredChild::PaperSource => 6,
        AuthoredChild::PageBorders => 7,
        AuthoredChild::LineNumbering => 8,
        AuthoredChild::PageNumbering => 9,
        AuthoredChild::Columns => 10,
        AuthoredChild::FormProtection => 11,
        AuthoredChild::VerticalAlignment => 12,
        AuthoredChild::NoEndnote(_) => 13,
        AuthoredChild::TitlePage => 14,
        AuthoredChild::TextDirection => 15,
        AuthoredChild::Bidirectional => 16,
        AuthoredChild::RtlGutter => 17,
        AuthoredChild::DocumentGrid => 18,
        AuthoredChild::PrinterSettings => 19,
        AuthoredChild::SectionChange(_) => 20,
        AuthoredChild::Unknown(_) => u8::MAX,
    }
}

fn write_authored_children(
    section: &SectionProperties,
    xml: &mut String,
    rels: Option<&super::super::super::relmap::RelationshipMapper>,
    plan: &NamespacePlan,
) -> Result<()> {
    let mut emitted = [false; 18];
    let mut headers = Vec::new();
    let mut footers = Vec::new();
    let mut header_managed = false;
    let mut footer_managed = false;
    let authored_headers: Vec<Kind> = section
        .authored_children
        .iter()
        .filter_map(|child| match child {
            AuthoredChild::Header(kind) => Some(*kind),
            _ => None,
        })
        .collect();
    let authored_footers: Vec<Kind> = section
        .authored_children
        .iter()
        .filter_map(|child| match child {
            AuthoredChild::Footer(kind) => Some(*kind),
            _ => None,
        })
        .collect();
    for (child_index, child) in section.authored_children.iter().enumerate() {
        let raw = section.authored_child_raw.get(child_index);
        if !matches!(child, AuthoredChild::Unknown(_)) {
            write_pending(
                section,
                xml,
                rels,
                plan,
                authored_rank(child),
                &mut emitted,
                &mut headers,
                &mut footers,
                &mut header_managed,
                &mut footer_managed,
                &authored_headers,
                &authored_footers,
            )?;
        }
        match child {
            AuthoredChild::Header(kind) => {
                if let Some(reference) = section
                    .headers
                    .iter()
                    .find(|reference| reference.kind == *kind)
                {
                    write_reference(
                        xml,
                        "headerReference",
                        reference,
                        rels,
                        true,
                        plan,
                        preserved_attributes(raw),
                    )?;
                }
                if !headers.contains(kind) {
                    headers.push(*kind);
                }
            },
            AuthoredChild::Footer(kind) => {
                if let Some(reference) = section
                    .footers
                    .iter()
                    .find(|reference| reference.kind == *kind)
                {
                    write_reference(
                        xml,
                        "footerReference",
                        reference,
                        rels,
                        false,
                        plan,
                        preserved_attributes(raw),
                    )?;
                }
                if !footers.contains(kind) {
                    footers.push(*kind);
                }
            },
            AuthoredChild::NoEndnote(raw) | AuthoredChild::SectionChange(raw) => {
                xml.push_str(raw);
            },
            AuthoredChild::Unknown(raw) => xml.push_str(raw),
            _ => {
                if let Some(index) = authored_field_index(child) {
                    if !emitted[index] {
                        write_field_preserving(section, xml, rels, plan, index, raw)?;
                        emitted[index] = true;
                    }
                }
            },
        }
    }
    write_pending(
        section,
        xml,
        rels,
        plan,
        u8::MAX,
        &mut emitted,
        &mut headers,
        &mut footers,
        &mut header_managed,
        &mut footer_managed,
        &authored_headers,
        &authored_footers,
    )?;
    Ok(())
}

fn write_pending(
    section: &SectionProperties,
    xml: &mut String,
    rels: Option<&super::super::super::relmap::RelationshipMapper>,
    plan: &NamespacePlan,
    before_rank: u8,
    emitted: &mut [bool; 18],
    headers: &mut Vec<Kind>,
    footers: &mut Vec<Kind>,
    header_managed: &mut bool,
    footer_managed: &mut bool,
    authored_headers: &[Kind],
    authored_footers: &[Kind],
) -> Result<()> {
    if before_rank > 0 {
        for reference in &section.headers {
            if !authored_headers.contains(&reference.kind) && !headers.contains(&reference.kind) {
                write_reference(xml, "headerReference", reference, rels, true, plan, &[])?;
                headers.push(reference.kind);
            }
        }
        for reference in &section.footers {
            if !authored_footers.contains(&reference.kind) && !footers.contains(&reference.kind) {
                write_reference(xml, "footerReference", reference, rels, false, plan, &[])?;
                footers.push(reference.kind);
            }
        }
        if section.headers.is_empty()
            && authored_headers.is_empty()
            && !section.suppress_managed_headers
            && !*header_managed
        {
            write_references(xml, "headerReference", &[], rels, true, plan)?;
            *header_managed = true;
        }
        if section.footers.is_empty()
            && authored_footers.is_empty()
            && !section.suppress_managed_footers
            && !*footer_managed
        {
            write_references(xml, "footerReference", &[], rels, false, plan)?;
            *footer_managed = true;
        }
    }
    for (index, rank) in [
        (0, 1),
        (1, 2),
        (2, 3),
        (3, 4),
        (4, 5),
        (5, 6),
        (6, 7),
        (7, 8),
        (8, 9),
        (9, 10),
        (10, 11),
        (11, 12),
        (12, 14),
        (13, 15),
        (14, 16),
        (15, 17),
        (16, 18),
        (17, 19),
    ] {
        if rank < before_rank && !emitted[index] {
            write_field_preserving(
                section,
                xml,
                rels,
                plan,
                index,
                raw_for_field(section, index),
            )?;
            emitted[index] = true;
        }
    }
    Ok(())
}

fn authored_field_index(child: &AuthoredChild) -> Option<usize> {
    Some(match child {
        AuthoredChild::FootnotePr => 0,
        AuthoredChild::EndnotePr => 1,
        AuthoredChild::Type => 2,
        AuthoredChild::PageSize => 3,
        AuthoredChild::PageMargins => 4,
        AuthoredChild::PaperSource => 5,
        AuthoredChild::PageBorders => 6,
        AuthoredChild::LineNumbering => 7,
        AuthoredChild::PageNumbering => 8,
        AuthoredChild::Columns => 9,
        AuthoredChild::FormProtection => 10,
        AuthoredChild::VerticalAlignment => 11,
        AuthoredChild::TitlePage => 12,
        AuthoredChild::TextDirection => 13,
        AuthoredChild::Bidirectional => 14,
        AuthoredChild::RtlGutter => 15,
        AuthoredChild::DocumentGrid => 16,
        AuthoredChild::PrinterSettings => 17,
        AuthoredChild::Header(_)
        | AuthoredChild::Footer(_)
        | AuthoredChild::NoEndnote(_)
        | AuthoredChild::SectionChange(_)
        | AuthoredChild::Unknown(_) => return None,
    })
}

fn preserved_attributes(raw: Option<&AuthoredChildRaw>) -> &[String] {
    raw.filter(|raw| !raw.raw.is_empty())
        .map_or(&[], |raw| raw.preserved_attributes.as_slice())
}

fn raw_for_field(section: &SectionProperties, index: usize) -> Option<&AuthoredChildRaw> {
    section
        .authored_children
        .iter()
        .enumerate()
        .find_map(|(child_index, child)| {
            (authored_field_index(child) == Some(index))
                .then(|| section.authored_child_raw.get(child_index))
                .flatten()
        })
}

fn write_field_preserving(
    section: &SectionProperties,
    xml: &mut String,
    rels: Option<&super::super::super::relmap::RelationshipMapper>,
    plan: &NamespacePlan,
    index: usize,
    raw: Option<&AuthoredChildRaw>,
) -> Result<()> {
    if raw.is_some_and(|raw| raw.original_on_off == Some(false))
        && current_on_off(section, index) == Some(false)
    {
        if let Some(raw) = raw {
            xml.push_str(&raw.raw);
            return Ok(());
        }
    }
    if let Some(raw) = raw
        && preserved_attributes(Some(raw)).iter().any(|attribute| {
            plan.word_prefix
                .as_deref()
                .is_some_and(|prefix| plan.shadows_word_prefix(attribute, prefix))
                || plan.shadows_word_prefix(attribute, &plan.word_attribute_prefix)
                || (index == 17 && plan.shadows_relationship_prefix(attribute))
        })
    {
        return Err(Error::InvalidFormat(
            "known section child shadows the generated Word attribute prefix".into(),
        ));
    }
    let mut rendered = String::new();
    write_field(section, &mut rendered, rels, plan, index)?;
    append_preserved_attributes(&mut rendered, preserved_attributes(raw))?;
    xml.push_str(&rendered);
    Ok(())
}

fn current_on_off(section: &SectionProperties, index: usize) -> Option<bool> {
    match index {
        10 => Some(section.form_protection),
        12 => Some(section.title_page),
        14 => Some(section.bidirectional),
        15 => Some(section.rtl_gutter),
        _ => None,
    }
}

fn append_preserved_attributes(xml: &mut String, attributes: &[String]) -> Result<()> {
    if attributes.is_empty() || xml.is_empty() {
        return Ok(());
    }
    let end = first_start_tag_end(xml)?;
    let insertion = match end.checked_sub(2) {
        Some(start) if xml.as_bytes().get(start..end) == Some(&b"/>"[..]) => start,
        _ => end
            .checked_sub(1)
            .filter(|position| xml.as_bytes().get(*position) == Some(&b'>'))
            .ok_or_else(|| Error::InvalidFormat("generated section start tag is invalid".into()))?,
    };
    let mut suffix = String::new();
    for attribute in attributes {
        write!(suffix, " {attribute}").map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.insert_str(insertion, &suffix);
    Ok(())
}

fn first_start_tag_end(xml: &str) -> Result<usize> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("generated section child is not valid XML: {error}"))
        })?;
        match event {
            Event::Start(_) | Event::Empty(_) => {
                return usize::try_from(reader.buffer_position()).map_err(|_source_error| {
                    Error::InvalidFormat("generated section child offset overflow".into())
                });
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "generated section child has no start tag".into(),
                ));
            },
            _ => {},
        }
    }
}

fn write_field(
    section: &SectionProperties,
    xml: &mut String,
    rels: Option<&super::super::super::relmap::RelationshipMapper>,
    plan: &NamespacePlan,
    index: usize,
) -> Result<()> {
    match index {
        0 => {
            if let Some(note) = &section.footnotes {
                write_footnotes(xml, note, plan)?;
            } else if rels.is_some_and(|rels| rels.get_footnotes_id().is_some()) {
                write!(
                    xml,
                    "<{}><{} {}=\"decimal\"/></{}>",
                    plan.word_qname("footnotePr"),
                    plan.word_qname("numFmt"),
                    plan.word_attribute_qname("val"),
                    plan.word_qname("footnotePr")
                )
                .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        1 => {
            if let Some(note) = &section.endnotes {
                write_endnotes(xml, note, plan)?;
            } else if rels.is_some_and(|rels| rels.get_endnotes_id().is_some()) {
                write!(
                    xml,
                    "<{}><{} {}=\"decimal\"/></{}>",
                    plan.word_qname("endnotePr"),
                    plan.word_qname("numFmt"),
                    plan.word_attribute_qname("val"),
                    plan.word_qname("endnotePr")
                )
                .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        2 => {
            if let Some(start_type) = section.start_type {
                write!(
                    xml,
                    "<{} {}=\"{}\"/>",
                    plan.word_qname("type"),
                    plan.word_attribute_qname("val"),
                    start_type.to_xml()
                )
                .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        3 => {
            write!(
                xml,
                "<{} {}=\"{}\" {}=\"{}\" {}=\"{}\"/>",
                plan.word_qname("pgSz"),
                plan.word_attribute_qname("w"),
                section.page_width,
                plan.word_attribute_qname("h"),
                section.page_height,
                plan.word_attribute_qname("orient"),
                section.orientation.as_str()
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
        },
        4 => {
            write!(
                xml,
                "<{} {}=\"{}\" {}=\"{}\" {}=\"{}\" {}=\"{}\" {}=\"{}\" {}=\"{}\" {}=\"{}\"/>",
                plan.word_qname("pgMar"),
                plan.word_attribute_qname("top"),
                section.margin_top,
                plan.word_attribute_qname("right"),
                section.margin_right,
                plan.word_attribute_qname("bottom"),
                section.margin_bottom,
                plan.word_attribute_qname("left"),
                section.margin_left,
                plan.word_attribute_qname("header"),
                section.header_distance,
                plan.word_attribute_qname("footer"),
                section.footer_distance,
                plan.word_attribute_qname("gutter"),
                section.gutter
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
        },
        5 => {
            if let Some(source) = &section.paper_source {
                write!(xml, "<{}", plan.word_qname("paperSrc"))
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if let Some(first) = source.first {
                    write!(xml, " {}=\"{first}\"", plan.word_attribute_qname("first"))
                        .map_err(|error| Error::Xml(error.to_string()))?;
                }
                if let Some(other) = source.other {
                    write!(xml, " {}=\"{other}\"", plan.word_attribute_qname("other"))
                        .map_err(|error| Error::Xml(error.to_string()))?;
                }
                xml.push_str("/>");
            }
        },
        6 => {
            if let Some(borders) = &section.page_borders {
                write_page_borders(xml, borders, plan)?;
            }
        },
        7 => {
            if let Some(numbering) = &section.line_numbering {
                write_line_numbering(xml, numbering, plan)?;
            }
        },
        8 => {
            if let Some(numbering) = &section.page_numbering {
                write_page_numbering(xml, numbering, plan)?;
            }
        },
        9 => {
            if let Some(columns) = &section.columns {
                write_columns(xml, columns, plan)?;
            }
        },
        10 => {
            if section.form_protection {
                write!(xml, "<{}/>", plan.word_qname("formProt"))
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        11 => {
            if let Some(alignment) = section.vertical_alignment {
                write!(
                    xml,
                    "<{} {}=\"{}\"/>",
                    plan.word_qname("vAlign"),
                    plan.word_attribute_qname("val"),
                    alignment.as_str()
                )
                .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        12 => {
            if section.title_page {
                write!(xml, "<{}/>", plan.word_qname("titlePg"))
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        13 => {
            if let Some(direction) = section.text_direction {
                write!(
                    xml,
                    "<{} {}=\"{}\"/>",
                    plan.word_qname("textDirection"),
                    plan.word_attribute_qname("val"),
                    direction.as_str()
                )
                .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        14 => {
            if section.bidirectional {
                write!(xml, "<{}/>", plan.word_qname("bidi"))
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        15 => {
            if section.rtl_gutter {
                write!(xml, "<{}/>", plan.word_qname("rtlGutter"))
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        16 => {
            if let Some(grid) = &section.document_grid {
                write_grid(xml, grid, plan)?;
            }
        },
        17 => {
            if let Some(id) = &section.printer_settings_relationship_id {
                validate_relationship_id(id, "printer-settings")?;
                write!(
                    xml,
                    "<{} {}=\"{}\"/>",
                    plan.word_qname("printerSettings"),
                    plan.relationship_qname("id"),
                    escape(id)
                )
                .map_err(|error| Error::Xml(error.to_string()))?;
            }
        },
        _ => {
            return Err(Error::InvalidFormat(
                "invalid authored section field".into(),
            ));
        },
    }
    Ok(())
}

fn section_child_rank(name: &str) -> Option<u8> {
    match name {
        "headerReference" | "footerReference" => Some(0),
        "footnotePr" => Some(1),
        "endnotePr" => Some(2),
        "type" => Some(3),
        "pgSz" => Some(4),
        "pgMar" => Some(5),
        "paperSrc" => Some(6),
        "pgBorders" => Some(7),
        "lnNumType" => Some(8),
        "pgNumType" => Some(9),
        "cols" => Some(10),
        "formProt" => Some(11),
        "vAlign" => Some(12),
        "noEndnote" => Some(13),
        "titlePg" => Some(14),
        "textDirection" => Some(15),
        "bidi" => Some(16),
        "rtlGutter" => Some(17),
        "docGrid" => Some(18),
        "printerSettings" => Some(19),
        "sectPrChange" => Some(20),
        _ => None,
    }
}

pub(super) fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

fn direct_children(xml: &str) -> Result<Vec<(String, String)>> {
    let metadata = root_metadata(xml)?;
    let root_prefix = metadata
        .prefix
        .as_deref()
        .map(str::as_bytes)
        .map(ToOwned::to_owned);
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut child: Option<(String, usize, usize)> = None;
    let mut stack = Vec::new();
    let mut children = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))?;
        let (word_element, event) = {
            let (namespace, event) = reader.read_resolved_event().map_err(|error| {
                Error::InvalidFormat(format!("section namespace resolution failed: {error}"))
            })?;
            let word_element = match &event {
                Event::Start(element) | Event::Empty(element) => is_word_element_at(
                    &namespace,
                    root_prefix.as_deref(),
                    metadata.root_default_namespace,
                    element,
                    depth > 0,
                ),
                _ => false,
            };
            (word_element, event)
        };
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))?;
        if depth > 0 && matches!(&event, Event::Start(_) | Event::Empty(_)) {
            validate_direct_child_attributes(&metadata, &xml[start..end])?;
        }
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen || element.local_name().as_ref() != b"sectPr" {
                        return Err(Error::InvalidFormat(
                            "section properties have an invalid root".into(),
                        ));
                    }
                    if !word_element {
                        return Err(Error::InvalidFormat(
                            "section properties use a foreign root namespace".into(),
                        ));
                    }
                    root_seen = true;
                } else if depth == 1 {
                    child = Some((child_name(&element, word_element), start, 1));
                } else if let Some((_, _, child_depth)) = child.as_mut() {
                    *child_depth += 1;
                }
                stack.push(element.name().as_ref().to_vec());
                depth += 1;
            },
            Event::Empty(element) if depth == 1 => {
                let name = child_name(&element, word_element);
                children.push((name, xml[start..end].to_string()));
            },
            Event::Empty(element) if depth == 0 => {
                if root_seen || element.local_name().as_ref() != b"sectPr" {
                    return Err(Error::InvalidFormat(
                        "section properties have an invalid root".into(),
                    ));
                }
                if !word_element {
                    return Err(Error::InvalidFormat(
                        "section properties use a foreign root namespace".into(),
                    ));
                }
                root_seen = true;
            },
            Event::End(element) => {
                let expected = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("section properties have an unmatched end element".into())
                })?;
                if expected != element.name().as_ref() {
                    return Err(Error::InvalidFormat(
                        "section properties have mismatched end elements".into(),
                    ));
                }
                if let Some((_, _, child_depth)) = child.as_mut() {
                    *child_depth -= 1;
                    if *child_depth == 0 {
                        let Some((name, child_start, _)) = child.take() else {
                            return Err(Error::InvalidFormat(
                                "section properties lost a direct child".into(),
                            ));
                        };
                        children.push((name, xml[child_start..end].to_string()));
                    }
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid section XML nesting".into()))?;
            },
            Event::Eof => break,
            Event::Text(text) if depth == 0 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "section properties contain trailing non-whitespace text".into(),
                    ));
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::Decl(_) if depth == 0 => {
                if root_seen {
                    return Err(Error::InvalidFormat(
                        "section properties contain trailing markup".into(),
                    ));
                }
            },
            Event::CData(_) | Event::DocType(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "section properties contain invalid content outside root".into(),
                ));
            },
            Event::Text(text) if depth == 1 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "section property contains unsupported text".into(),
                    ));
                }
                children.push(("@raw".to_owned(), xml[start..end].to_owned()));
            },
            Event::Decl(_) | Event::DocType(_) if depth > 0 => {
                return Err(Error::InvalidFormat(
                    "section property contains invalid nested markup".into(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 1 => {
                return Err(Error::InvalidFormat(
                    "section property contains unsupported nested content".into(),
                ));
            },
            Event::Text(_) | Event::Comment(_) | Event::PI(_) if depth == 1 => {
                children.push(("@raw".to_owned(), xml[start..end].to_owned()))
            },
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || depth != 0 || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "unterminated section properties".into(),
        ));
    }
    Ok(children)
}

fn validate_direct_child_attributes(metadata: &RootMetadata, xml: &str) -> Result<()> {
    for attribute in decode_attributes(metadata, xml)? {
        if attribute.qualified_name == "xmlns" || attribute.qualified_name.starts_with("xmlns:") {
            continue;
        }
        if attribute.unknown_prefix
            && (metadata.namespace_resolved || attribute.namespace != AttributeNamespace::Word)
            && !(attribute.namespace == AttributeNamespace::Relationship && attribute.name == "id")
        {
            return Err(Error::InvalidFormat(format!(
                "section child uses an undeclared attribute prefix '{}'",
                attribute.qualified_name
            )));
        }
    }
    Ok(())
}

struct RootMetadata {
    prefix: Option<String>,
    bindings: Vec<NamespaceBinding>,
    root_default_namespace: bool,
    namespace_resolved: bool,
    attributes: Vec<String>,
}

impl RootMetadata {
    fn has_binding(&self, prefix: &[u8]) -> bool {
        self.bindings.iter().any(|binding| {
            binding
                .prefix
                .as_deref()
                .is_some_and(|candidate| candidate.as_bytes() == prefix)
        })
    }
}

fn root_metadata(xml: &str) -> Result<RootMetadata> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let (namespace, element) = loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("section namespace resolution failed: {error}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => break (namespace, element),
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Text(text) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "section properties contain invalid text before root".into(),
                    ));
                }
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "section properties have no root element".into(),
                ));
            },
            Event::CData(_) | Event::DocType(_) | Event::GeneralRef(_) | Event::End(_) => {
                return Err(Error::InvalidFormat(
                    "section properties contain invalid content before root".into(),
                ));
            },
        }
    };
    let prefix = element
        .name()
        .prefix()
        .map(|prefix| String::from_utf8_lossy(prefix.into_inner()).into_owned());
    if element.local_name().as_ref() != b"sectPr"
        || !is_word_element(&namespace, prefix.as_deref().map(str::as_bytes))
    {
        return Err(Error::InvalidFormat(
            "section properties use an invalid root namespace".into(),
        ));
    }
    let namespace_resolved = crate::namespace::is_wordprocessing_namespace(&namespace);
    let mut bindings = Vec::new();
    let mut attributes = Vec::new();
    let resolver = reader.resolver().clone();
    let mut seen_attributes = std::collections::HashSet::new();
    let mut seen_raw_attributes = std::collections::HashSet::new();
    let mut seen_bindings = std::collections::HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        if !seen_raw_attributes.insert(name.to_vec()) {
            return Err(Error::InvalidFormat(
                "section root contains duplicate attributes".into(),
            ));
        }
        let prefix_binding = if name == b"xmlns" {
            Some(None)
        } else {
            name.strip_prefix(b"xmlns:")
                .map(|prefix| Some(String::from_utf8_lossy(prefix).into_owned()))
        };
        let Some(prefix_binding) = prefix_binding else {
            let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
            let key = (
                namespace_key(&attribute_namespace),
                attribute.key.local_name().as_ref().to_vec(),
            );
            if !seen_attributes.insert(key) {
                return Err(Error::InvalidFormat(
                    "section root contains duplicate attributes".into(),
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            attributes.push(format!(
                "{}=\"{}\"",
                String::from_utf8_lossy(name),
                escape(&value)
            ));
            continue;
        };
        let uri = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        if !seen_bindings.insert(prefix_binding.clone()) {
            return Err(Error::InvalidFormat(
                "section root contains duplicate namespace bindings".into(),
            ));
        }
        bindings.push(NamespaceBinding {
            prefix: prefix_binding,
            uri,
        });
    }
    let default_word_namespace = bindings
        .iter()
        .any(|binding| binding.prefix.is_none() && is_word_namespace_uri(&binding.uri));
    let root_default_namespace = prefix.is_none() && namespace_resolved && default_word_namespace;
    Ok(RootMetadata {
        prefix,
        bindings,
        root_default_namespace,
        namespace_resolved,
        attributes,
    })
}

fn is_word_namespace_uri(uri: &str) -> bool {
    matches!(
        uri,
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            | "http://purl.oclc.org/ooxml/wordprocessingml/main"
    )
}

fn namespace_key(namespace: &ResolveResult<'_>) -> String {
    match namespace {
        ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            format!("{{{}}}", String::from_utf8_lossy(value))
        },
        ResolveResult::Unknown(prefix) => {
            format!("?{}", String::from_utf8_lossy(prefix))
        },
        ResolveResult::Unbound => String::new(),
    }
}

fn is_word_element(namespace: &ResolveResult<'_>, root_prefix: Option<&[u8]>) -> bool {
    if crate::namespace::is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => root_prefix == Some(prefix.as_slice()),
        ResolveResult::Unbound => root_prefix.is_none(),
        ResolveResult::Bound(_) => false,
    }
}

fn is_word_element_at(
    namespace: &ResolveResult<'_>,
    root_prefix: Option<&[u8]>,
    root_default_namespace: bool,
    element: &quick_xml::events::BytesStart<'_>,
    child: bool,
) -> bool {
    if child
        && root_default_namespace
        && element.name().prefix().is_none()
        && element.attributes().any(|attribute| {
            attribute.ok().is_some_and(|attribute| {
                attribute.key.as_ref() == b"xmlns" && attribute.value.as_ref().is_empty()
            })
        })
    {
        return false;
    }
    is_word_element(namespace, root_prefix)
}

fn is_relationship_namespace(namespace: &ResolveResult<'_>) -> bool {
    const TRANSITIONAL: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
    matches!(
        namespace,
        ResolveResult::Bound(quick_xml::name::Namespace(value))
            if *value == TRANSITIONAL || *value == STRICT
    )
}

fn child_name(element: &quick_xml::events::BytesStart<'_>, word: bool) -> String {
    if word {
        String::from_utf8_lossy(element.local_name().as_ref()).into_owned()
    } else {
        format!(
            "@foreign:{}",
            String::from_utf8_lossy(element.local_name().as_ref())
        )
    }
}

fn authored_child(metadata: &RootMetadata, name: &str, raw: &str) -> Result<AuthoredChild> {
    let marker = match name {
        "headerReference" => AuthoredChild::Header(
            Kind::from_xml(&required_attr(metadata, raw, b"type")?)
                .ok_or_else(|| Error::InvalidFormat("invalid section header/footer type".into()))?,
        ),
        "footerReference" => AuthoredChild::Footer(
            Kind::from_xml(&required_attr(metadata, raw, b"type")?)
                .ok_or_else(|| Error::InvalidFormat("invalid section header/footer type".into()))?,
        ),
        "footnotePr" => AuthoredChild::FootnotePr,
        "endnotePr" => AuthoredChild::EndnotePr,
        "type" => AuthoredChild::Type,
        "pgSz" => AuthoredChild::PageSize,
        "pgMar" => AuthoredChild::PageMargins,
        "paperSrc" => AuthoredChild::PaperSource,
        "pgBorders" => AuthoredChild::PageBorders,
        "lnNumType" => AuthoredChild::LineNumbering,
        "pgNumType" => AuthoredChild::PageNumbering,
        "cols" => AuthoredChild::Columns,
        "formProt" => AuthoredChild::FormProtection,
        "vAlign" => AuthoredChild::VerticalAlignment,
        "noEndnote" => AuthoredChild::NoEndnote(raw.to_owned()),
        "titlePg" => AuthoredChild::TitlePage,
        "textDirection" => AuthoredChild::TextDirection,
        "bidi" => AuthoredChild::Bidirectional,
        "rtlGutter" => AuthoredChild::RtlGutter,
        "docGrid" => AuthoredChild::DocumentGrid,
        "printerSettings" => AuthoredChild::PrinterSettings,
        "sectPrChange" => AuthoredChild::SectionChange(raw.to_owned()),
        _ => AuthoredChild::Unknown(raw.to_owned()),
    };
    Ok(marker)
}

#[cfg(test)]
mod tests {
    use super::SectionProperties;

    fn section_xml(children: &str) -> String {
        format!(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{children}</w:sectPr>"#
        )
    }

    #[test]
    fn section_properties_accept_interleaved_header_footer_references() {
        let properties = SectionProperties::from_xml(&section_xml(
            r#"<w:headerReference w:type="default" r:id="rHdDefault"/><w:footerReference w:type="default" r:id="rFtDefault"/><w:headerReference w:type="first" r:id="rHdFirst"/><w:footerReference w:type="first" r:id="rFtFirst"/><w:headerReference w:type="even" r:id="rHdEven"/><w:footerReference w:type="even" r:id="rFtEven"/><w:pgSz w:w="12240" w:h="15840"/>"#,
        ))
        .expect("interleaved references are schema-valid");
        assert_eq!(properties.headers.len(), 3);
        assert_eq!(properties.footers.len(), 3);
    }

    #[test]
    fn section_properties_reject_reference_after_page_size() {
        assert!(SectionProperties::from_xml(&section_xml(
            r#"<w:pgSz w:w="12240" w:h="15840"/><w:headerReference w:type="default" r:id="rHdDefault"/>"#,
        ))
        .is_err());
    }

    #[test]
    fn section_properties_reject_duplicate_same_reference_type() {
        for children in [
            r#"<w:headerReference w:type="default" r:id="rHdOne"/><w:headerReference w:type="default" r:id="rHdTwo"/>"#,
            r#"<w:footerReference w:type="default" r:id="rFtOne"/><w:footerReference w:type="default" r:id="rFtTwo"/>"#,
        ] {
            assert!(SectionProperties::from_xml(&section_xml(children)).is_err());
        }
    }

    #[test]
    fn namespace_forms_and_detached_prefixes_parse() {
        for xml in [
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
            r#"<w:sectPr xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
            r#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><pgSz w:w="12240" w:h="15840"/></sectPr>"#,
            r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
        ] {
            assert!(SectionProperties::from_xml(xml).is_ok());
        }
        assert!(
            SectionProperties::from_xml(
                r#"<x:sectPr xmlns:x="urn:foreign"><x:pgSz x:w="1" x:h="1"/></x:sectPr>"#
            )
            .is_err()
        );
    }

    #[test]
    fn foreign_same_local_children_and_raw_positions_survive_write() {
        let section = SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><x:pgSz x:w="1"/><x:pgSz x:w="2"/><!--keep--><w:pgMar w:top="1"/></w:sectPr>"#,
        )
        .expect("foreign extensions parse");
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("foreign extensions write");
        assert!(output.contains(r#"<x:pgSz x:w="1"/>"#));
        assert!(output.contains(r#"<x:pgSz x:w="2"/>"#));
        assert!(output.contains("<!--keep-->"));
        assert!(output.find("<x:pgSz").unwrap() < output.find("<!--keep-->").unwrap());
    }

    #[test]
    fn interleaved_reference_slots_replace_in_place() {
        let mut section = SectionProperties::from_xml(
            &section_xml(
                r#"<w:headerReference w:type="default" r:id="rHd"/><w:footerReference w:type="default" r:id="rFt"/><w:headerReference w:type="first" r:id="rHdFirst"/><w:footerReference w:type="first" r:id="rFtFirst"/><w:pgSz w:w="12240" w:h="15840"/>"#,
            ),
        )
        .expect("interleaved references parse");
        section.headers[0].relationship_id = Some("newHeader".to_owned());
        section.footers[0].relationship_id = Some("newFooter".to_owned());
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("reference replacement write");
        assert!(
            output.find("r:id=\"newHeader\"").unwrap() < output.find("r:id=\"newFooter\"").unwrap()
        );
        assert!(
            output.find("r:id=\"newFooter\"").unwrap() < output.find("r:id=\"rHdFirst\"").unwrap()
        );
    }

    #[test]
    fn self_closing_and_mismatched_section_roots_are_handled() {
        assert!(SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
        )
        .is_ok());
        assert!(SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz></w:pgMar></w:sectPr>"#
        )
        .is_err());
    }

    #[test]
    fn relationship_ids_require_ncname_and_foreign_r_is_not_a_relationship() {
        assert!(SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="urn:foreign"><w:headerReference w:type="default" r:id="not valid"/></w:sectPr>"#
        )
        .is_err());
        assert!(SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>"#
        )
        .is_ok());
        assert!(SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:printerSettings r:id="bad id"/></w:sectPr>"#
        )
        .is_err());
    }

    #[test]
    fn default_namespace_root_writes_unprefixed_word_elements() {
        let section = SectionProperties::from_xml(
            r#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><pgSz w="12240" h="15840"/></sectPr>"#,
        )
        .expect("default namespace section parses");
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("default namespace section writes");
        assert!(output.starts_with("<sectPr xmlns=\""));
        assert!(output.contains("<pgSz "));
        assert!(
            output.contains(
                "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
            )
        );
        assert!(output.contains("<pgSz w:w=\"12240\" w:h=\"15840\""));
        assert!(!output.contains("<w:pgSz"));
        let reparsed = SectionProperties::from_xml(&output).expect("qualified output reparses");
        assert_eq!(reparsed.page_width, 12240);
    }

    #[test]
    fn mixed_word_bindings_use_the_root_namespace_for_generated_attributes() {
        for (xml, expected_namespace) in [
            (
                r#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:strict="http://purl.oclc.org/ooxml/wordprocessingml/main"><pgSz w="12240" h="15840"/></sectPr>"#,
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            ),
            (
                r#"<sectPr xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:transitional="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><pgSz w="12240" h="15840"/></sectPr>"#,
                "http://purl.oclc.org/ooxml/wordprocessingml/main",
            ),
        ] {
            let section = SectionProperties::from_xml(xml).expect("mixed section parses");
            let mut output = String::new();
            section
                .write_xml(&mut output, None)
                .expect("mixed section writes");
            assert!(output.contains(&format!("xmlns:w=\"{expected_namespace}\"")));
            let reparsed = SectionProperties::from_xml(&output).expect("mixed output reparses");
            assert_eq!(reparsed.page_width, 12240);
        }
    }

    #[test]
    fn prefixed_child_default_reset_is_not_replayed_on_regeneration() {
        let section = SectionProperties::from_xml(
            r#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz xmlns="" w:w="12240" w:h="15840"/></sectPr>"#,
        )
        .expect("prefixed child parses");
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("prefixed child writes");
        assert!(!output.contains("<pgSz xmlns=\"\""));
        let reparsed = SectionProperties::from_xml(&output).expect("reset-free output reparses");
        assert_eq!(reparsed.page_width, 12240);
    }

    #[test]
    fn alternate_word_prefix_and_foreign_w_r_bindings_remain_namespace_safe() {
        let mut section = SectionProperties::from_xml(
            r#"<x:sectPr xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w="urn:foreign" xmlns:r="urn:foreign" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:headerReference x:type="default" rel:id="rHeader"/><x:pgSz x:w="12240" x:h="15840"/><w:pgSz w:w="1"/></x:sectPr>"#,
        )
        .expect("alternate namespace section parses");
        section.page_width = 10000;
        section.headers[0].relationship_id = Some("rHeader2".to_owned());
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("alternate namespace section writes");
        assert!(output.contains("<x:sectPr"));
        assert!(output.contains("<x:pgSz"));
        assert!(output.contains("rel:id=\"rHeader2\""));
        assert!(output.contains("<w:pgSz w:w=\"1\"/>"));
        assert!(output.contains("xmlns:r=\"urn:foreign\""));
    }

    #[test]
    fn strict_namespaces_are_retained_for_generated_relationships() {
        let section = SectionProperties::from_xml(
            r#"<x:sectPr xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:q="http://purl.oclc.org/ooxml/officeDocument/relationships"><x:headerReference x:type="default" q:id="rHeader"/><x:pgSz x:w="12240" x:h="15840"/></x:sectPr>"#,
        )
        .expect("strict namespace section parses");
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("strict namespace section writes");
        assert!(output.contains("xmlns:x=\"http://purl.oclc.org/ooxml/wordprocessingml/main\""));
        assert!(output.contains("q:id=\"rHeader\""));
        let reparsed = SectionProperties::from_xml(&output).expect("strict output reparses");
        assert_eq!(reparsed.page_width, 12240);
    }

    #[test]
    fn relationship_decoder_rejects_unqualified_word_and_foreign_ids() {
        for xml in [
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:headerReference w:type="default" id="rHeader"/></w:sectPr>"#,
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:headerReference w:type="default" w:id="rHeader"/></w:sectPr>"#,
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="urn:foreign"><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>"#,
        ] {
            assert!(SectionProperties::from_xml(xml).is_err());
        }
    }

    #[test]
    fn foreign_same_local_id_is_preserved_with_valid_relationship_id() {
        let section = SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="urn:mc"><w:headerReference w:type="default" r:id="rHeader" mc:id="metadata"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
        )
        .expect("foreign id parses");
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("foreign id writes");
        assert!(output.contains("r:id=\"rHeader\""));
        assert!(output.contains("mc:id=\"metadata\""));
        SectionProperties::from_xml(&output).expect("foreign id output reparses");
    }

    #[test]
    fn generated_relationship_prefix_avoids_word_prefix_collision() {
        let mut section = SectionProperties::from_xml(
            r#"<r:sectPr xmlns:r="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><r:pgSz r:w="12240" r:h="15840"/></r:sectPr>"#,
        )
        .expect("section parses");
        section.headers.push(
            super::super::super::model::SectionHeaderFooterReference::existing(
                crate::header_footer::Kind::Primary,
                "rHeader",
            ),
        );
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("section writes");
        assert!(output.contains(
            "xmlns:r1=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        ));
        assert!(output.contains("r1:id=\"rHeader\""));
    }

    #[test]
    fn fresh_and_detached_namespace_serialization_are_safe() {
        let fresh = SectionProperties::default();
        let mut output = String::new();
        fresh
            .write_xml(&mut output, None)
            .expect("fresh section writes");
        assert!(
            output.contains(
                "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
            )
        );

        let detached = SectionProperties::from_xml(
            r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
        )
        .expect("detached section reads");
        let mut output = String::new();
        detached
            .write_xml(&mut output, None)
            .expect("detached section preserves its inherited lexical prefix");
        assert!(output.starts_with("<w:sectPr>"));
        assert!(!output.contains("xmlns:w="));
    }

    #[test]
    fn managed_authored_reference_deletion_does_not_resurrect() {
        let mut section = SectionProperties::from_xml(&section_xml(
            r#"<w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="12240" w:h="15840"/>"#,
        ))
        .expect("section parses");
        section.headers.clear();
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("section writes");
        assert!(!output.contains("headerReference"));
    }

    #[test]
    fn invalid_st_on_off_values_are_rejected() {
        for child in [
            r#"<w:formProt w:val="maybe"/>"#,
            r#"<w:titlePg w:val="maybe"/>"#,
            r#"<w:bidi w:val="maybe"/>"#,
            r#"<w:rtlGutter w:val="maybe"/>"#,
            r#"<w:cols w:equalWidth="maybe"/>"#,
            r#"<w:pgBorders><w:top w:val="single" w:shadow="maybe"/></w:pgBorders>"#,
        ] {
            assert!(SectionProperties::from_xml(&section_xml(child)).is_err());
        }
    }

    #[test]
    fn page_border_edges_follow_schema_order() {
        assert!(SectionProperties::from_xml(&section_xml(
            r#"<w:pgBorders><w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/></w:pgBorders>"#,
        ))
        .is_ok());
        assert!(SectionProperties::from_xml(&section_xml(
            r#"<w:pgBorders><w:top w:val="single"/><w:bottom w:val="single"/><w:left w:val="single"/></w:pgBorders>"#,
        ))
        .is_err());
    }

    #[test]
    fn explicit_false_on_off_values_are_preserved_until_enabled() {
        let mut section = SectionProperties::from_xml(&section_xml(
            r#"<w:pgSz w:w="12240" w:h="15840"/><w:formProt w:val="0"/><w:titlePg w:val="0"/><w:bidi w:val="0"/><w:rtlGutter w:val="0"/>"#,
        ))
        .expect("section parses");
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("section writes");
        assert_eq!(output.matches("w:val=\"0\"").count(), 4);
        section.form_protection = true;
        section.title_page = true;
        section.bidirectional = true;
        section.rtl_gutter = true;
        let mut enabled = String::new();
        section
            .write_xml(&mut enabled, None)
            .expect("enabled section writes");
        assert!(!enabled.contains("<w:formProt w:val=\"0\""));
        assert!(!enabled.contains("<w:titlePg w:val=\"0\""));
        assert!(!enabled.contains("<w:bidi w:val=\"0\""));
        assert!(!enabled.contains("<w:rtlGutter w:val=\"0\""));
    }

    #[test]
    fn root_and_known_child_unmodeled_attributes_survive_edits() {
        let mut section = SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="urn:mc" mc:Ignorable="x"><w:pgSz w:w="12240" w:h="15840" mc:custom="yes"/></w:sectPr>"#,
        )
        .expect("section parses");
        section.page_width = 10000;
        let mut output = String::new();
        section
            .write_xml(&mut output, None)
            .expect("section writes");
        assert!(output.contains("mc:Ignorable=\"x\""));
        assert!(output.contains("mc:custom=\"yes\""));
    }

    #[test]
    fn root_scanner_rejects_trailing_junk_and_preserves_explicit_unbinding() {
        assert!(SectionProperties::from_xml(
            r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>junk"#
        )
        .is_err());
        let section = SectionProperties::from_xml(
            r#"<sectPr xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><pgSz xmlns="" w="1" h="2"/></sectPr>"#,
        )
        .expect("foreign explicitly unbound child is preserved");
        assert_eq!(section.page_width, 12240);
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AttributeNamespace {
    Word,
    Relationship,
    Other,
}

#[derive(Debug)]
struct DecodedAttribute {
    name: String,
    qualified_name: String,
    value: String,
    namespace: AttributeNamespace,
    unknown_prefix: bool,
}

fn decode_attributes(metadata: &RootMetadata, xml: &str) -> Result<Vec<DecodedAttribute>> {
    let detached = detached_fragment(metadata, xml);
    let mut reader = NsReader::from_reader(detached.as_bytes());
    let element = loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("section namespace resolution failed: {error}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() != b"litchiRoot" =>
            {
                break element;
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "section property has no element".into(),
                ));
            },
            Event::Start(_) | Event::Empty(_) => {},
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    };
    let resolver = reader.resolver().clone();
    let mut result = Vec::new();
    let mut seen_attributes = std::collections::HashSet::new();
    let mut seen_raw_attributes = std::collections::HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if !seen_raw_attributes.insert(attribute.key.as_ref().to_vec()) {
            return Err(Error::InvalidFormat(
                "duplicate section property attributes".into(),
            ));
        }
        let name = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        let qualified_name = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let unknown_prefix = matches!(&namespace, ResolveResult::Unknown(_));
        if qualified_name != "xmlns" && !qualified_name.starts_with("xmlns:") {
            let key = (namespace_key(&namespace), name.clone());
            if !seen_attributes.insert(key) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate section property attribute '{qualified_name}'"
                )));
            }
        }
        let fragment_prefix = element
            .name()
            .prefix()
            .map(|prefix| prefix.into_inner().to_vec());
        let same_fragment_prefix = matches!(
            &namespace,
            ResolveResult::Unknown(prefix)
                if fragment_prefix.as_deref() == Some(prefix.as_slice())
                    && !metadata.has_binding(prefix.as_slice())
        );
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let relationship = name == "id"
            && (is_relationship_namespace(&namespace)
                || matches!(&namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r" && !metadata.has_binding(prefix.as_slice())));
        let word = !relationship
            && (crate::namespace::is_wordprocessing_namespace(&namespace)
                || matches!(namespace, ResolveResult::Unbound)
                || same_fragment_prefix);
        let namespace = if word {
            AttributeNamespace::Word
        } else if relationship {
            AttributeNamespace::Relationship
        } else {
            AttributeNamespace::Other
        };
        result.push(DecodedAttribute {
            name,
            qualified_name,
            value,
            namespace,
            unknown_prefix,
        });
    }
    Ok(result)
}

fn preserved_attribute_tokens(
    metadata: &RootMetadata,
    xml: &str,
    child_name: &str,
) -> Result<Vec<String>> {
    let modeled = modeled_attributes(child_name);
    let mut entries = Vec::new();
    let mut used_prefixes = std::collections::HashSet::new();
    for attribute in decode_attributes(metadata, xml)? {
        if attribute.qualified_name == "xmlns" {
            entries.push((
                format!(
                    "{}=\"{}\"",
                    attribute.qualified_name,
                    escape(&attribute.value)
                ),
                Some(String::new()),
            ));
            continue;
        }
        if let Some(prefix) = attribute.qualified_name.strip_prefix("xmlns:") {
            entries.push((
                format!(
                    "{}=\"{}\"",
                    attribute.qualified_name,
                    escape(&attribute.value)
                ),
                Some(prefix.to_owned()),
            ));
            continue;
        }
        let modeled_attribute = modeled
            .iter()
            .any(|(namespace, name)| *name == attribute.name && *namespace == attribute.namespace);
        if modeled_attribute {
            continue;
        }
        if let Some((prefix, _)) = attribute.qualified_name.split_once(':') {
            used_prefixes.insert(prefix.to_owned());
        }
        entries.push((
            format!(
                "{}=\"{}\"",
                attribute.qualified_name,
                escape(&attribute.value)
            ),
            None,
        ));
    }
    Ok(entries
        .into_iter()
        .filter_map(|(attribute, declaration)| match declaration {
            Some(prefix) if !prefix.is_empty() && !used_prefixes.contains(&prefix) => None,
            Some(_) | None => Some(attribute),
        })
        .collect())
}

fn reject_unmodeled_attributes(metadata: &RootMetadata, xml: &str, child_name: &str) -> Result<()> {
    let modeled = modeled_attributes(child_name);
    for attribute in decode_attributes(metadata, xml)? {
        if attribute.qualified_name == "xmlns" || attribute.qualified_name.starts_with("xmlns:") {
            continue;
        }
        if !modeled
            .iter()
            .any(|(namespace, name)| *name == attribute.name && *namespace == attribute.namespace)
        {
            return Err(Error::InvalidFormat(format!(
                "unsupported attribute '{}' in nested section property '{child_name}'",
                attribute.qualified_name
            )));
        }
    }
    Ok(())
}

fn modeled_attributes(child_name: &str) -> &'static [(AttributeNamespace, &'static str)] {
    match child_name {
        "headerReference" | "footerReference" => &[
            (AttributeNamespace::Word, "type"),
            (AttributeNamespace::Relationship, "id"),
        ],
        "printerSettings" => &[(AttributeNamespace::Relationship, "id")],
        "type" | "vAlign" | "textDirection" | "formProt" | "titlePg" | "bidi" | "rtlGutter" => {
            &[(AttributeNamespace::Word, "val")]
        },
        "pgSz" => &[
            (AttributeNamespace::Word, "w"),
            (AttributeNamespace::Word, "h"),
            (AttributeNamespace::Word, "orient"),
        ],
        "pgMar" => &[
            (AttributeNamespace::Word, "top"),
            (AttributeNamespace::Word, "right"),
            (AttributeNamespace::Word, "bottom"),
            (AttributeNamespace::Word, "left"),
            (AttributeNamespace::Word, "header"),
            (AttributeNamespace::Word, "footer"),
            (AttributeNamespace::Word, "gutter"),
        ],
        "paperSrc" => &[
            (AttributeNamespace::Word, "first"),
            (AttributeNamespace::Word, "other"),
        ],
        "pgNumType" => &[
            (AttributeNamespace::Word, "fmt"),
            (AttributeNamespace::Word, "start"),
            (AttributeNamespace::Word, "chapStyle"),
            (AttributeNamespace::Word, "chapSep"),
        ],
        "cols" => &[
            (AttributeNamespace::Word, "equalWidth"),
            (AttributeNamespace::Word, "num"),
            (AttributeNamespace::Word, "space"),
            (AttributeNamespace::Word, "sep"),
        ],
        "docGrid" => &[
            (AttributeNamespace::Word, "type"),
            (AttributeNamespace::Word, "linePitch"),
            (AttributeNamespace::Word, "charSpace"),
        ],
        "pgBorders" => &[
            (AttributeNamespace::Word, "offsetFrom"),
            (AttributeNamespace::Word, "zOrder"),
            (AttributeNamespace::Word, "display"),
        ],
        "lnNumType" => &[
            (AttributeNamespace::Word, "countBy"),
            (AttributeNamespace::Word, "start"),
            (AttributeNamespace::Word, "distance"),
            (AttributeNamespace::Word, "restart"),
        ],
        "numFmt" | "numStart" | "numRestart" | "pos" => &[(AttributeNamespace::Word, "val")],
        "col" => &[
            (AttributeNamespace::Word, "w"),
            (AttributeNamespace::Word, "space"),
        ],
        "top" | "left" | "bottom" | "right" => &[
            (AttributeNamespace::Word, "val"),
            (AttributeNamespace::Word, "sz"),
            (AttributeNamespace::Word, "space"),
            (AttributeNamespace::Word, "color"),
            (AttributeNamespace::Word, "shadow"),
            (AttributeNamespace::Word, "frame"),
        ],
        _ => &[],
    }
}

fn attributes(metadata: &RootMetadata, xml: &str) -> Result<Vec<(String, String)>> {
    let mut result = Vec::new();
    for attribute in decode_attributes(metadata, xml)? {
        if attribute.qualified_name == "xmlns" || attribute.qualified_name.starts_with("xmlns:") {
            continue;
        }
        if attribute.namespace != AttributeNamespace::Word {
            continue;
        }
        if result
            .iter()
            .any(|(candidate, _)| candidate == &attribute.name)
        {
            return Err(Error::InvalidFormat(format!(
                "duplicate section property attribute '{}'",
                attribute.name
            )));
        }
        result.push((attribute.name, attribute.value));
    }
    Ok(result)
}

fn detached_fragment(metadata: &RootMetadata, xml: &str) -> String {
    let mut detached = String::with_capacity(xml.len() + 128);
    detached.push_str("<litchiRoot");
    for binding in &metadata.bindings {
        detached.push_str(" xmlns");
        if let Some(prefix) = &binding.prefix {
            detached.push(':');
            detached.push_str(prefix);
        }
        detached.push_str("=\"");
        detached.push_str(&escape(&binding.uri));
        detached.push('"');
    }
    detached.push('>');
    detached.push_str(xml);
    detached.push_str("</litchiRoot>");
    detached
}

fn attr<'a>(attrs: &'a [(String, String)], name: &str) -> Option<&'a str> {
    attrs
        .iter()
        .find_map(|(candidate, value)| (candidate == name).then_some(value.as_str()))
}

fn required_attr(metadata: &RootMetadata, xml: &str, name: &[u8]) -> Result<String> {
    let name = String::from_utf8_lossy(name);
    let attrs = attributes(metadata, xml)?;
    attr(&attrs, &name)
        .map(ToOwned::to_owned)
        .ok_or_else(|| Error::InvalidFormat(format!("missing section attribute '{name}'")))
}

fn relationship_id(metadata: &RootMetadata, xml: &str, element: &str) -> Result<String> {
    let mut result = None;
    for attribute in decode_attributes(metadata, xml)? {
        if attribute.name != "id" {
            continue;
        }
        match attribute.namespace {
            AttributeNamespace::Relationship => {
                if result.is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "duplicate section {element} relationship ID"
                    )));
                }
                result = Some(attribute.value);
            },
            AttributeNamespace::Word => {
                return Err(Error::InvalidFormat(format!(
                    "section {element} relationship ID uses an invalid namespace"
                )));
            },
            AttributeNamespace::Other => {
                if attribute.qualified_name == "r:id" {
                    return Err(Error::InvalidFormat(format!(
                        "section {element} relationship ID uses a spoofed detached namespace"
                    )));
                }
            },
        }
    }
    let id = result.ok_or_else(|| {
        Error::InvalidFormat(format!("section {element} is missing a relationship ID"))
    })?;
    validate_relationship_id(&id, element)?;
    Ok(id)
}

fn parse_u32(value: &str, description: &str) -> Result<u32> {
    value.parse().map_err(|_source_error| {
        Error::InvalidFormat(format!("invalid {description} value '{value}'"))
    })
}

fn assign_u32(attrs: &[(String, String)], name: &str, slot: &mut u32) -> Result<()> {
    if let Some(value) = attr(attrs, name) {
        *slot = parse_u32(value, name)?;
    }
    Ok(())
}

fn parse_header_footer(metadata: &RootMetadata, xml: &str) -> Result<SectionHeaderFooterReference> {
    let relationship_id = relationship_id(metadata, xml, "header/footer")?;
    let kind = Kind::from_xml(&required_attr(metadata, xml, b"type")?)
        .ok_or_else(|| Error::InvalidFormat("invalid section header/footer type".to_string()))?;
    Ok(SectionHeaderFooterReference {
        kind,
        relationship_id: Some(relationship_id),
        part: None,
    })
}

fn validate_relationship_id(id: &str, element: &str) -> Result<()> {
    if !is_ncname(id) {
        return Err(Error::InvalidFormat(format!(
            "section {element} relationship ID is not an XML NCName"
        )));
    }
    Ok(())
}

fn parse_page_numbering(metadata: &RootMetadata, xml: &str) -> Result<SectionPageNumbering> {
    let attrs = attributes(metadata, xml)?;
    Ok(SectionPageNumbering {
        format: attr(&attrs, "fmt")
            .map(PageNumberFormat::parse)
            .transpose()?
            .unwrap_or(PageNumberFormat::Decimal),
        start: attr(&attrs, "start")
            .map(|value| parse_u32(value, "page number start"))
            .transpose()?,
        chapter_style: attr(&attrs, "chapStyle")
            .map(|value| {
                value
                    .parse::<u8>()
                    .map_err(|_source_error| Error::InvalidFormat("invalid chapter style".into()))
            })
            .transpose()?,
        chapter_separator: attr(&attrs, "chapSep").map(ChapterSep::parse).transpose()?,
    })
}

fn parse_columns(metadata: &RootMetadata, xml: &str) -> Result<SectionColumns> {
    let attrs = attributes(metadata, xml)?;
    let mut columns = SectionColumns {
        equal_width: parse_on_off_attr(&attrs, "equalWidth", true)?,
        count: attr(&attrs, "num")
            .map(|value| {
                value.parse::<u16>().map_err(|_source_error| {
                    Error::InvalidFormat("invalid section column count".into())
                })
            })
            .transpose()?
            .unwrap_or(1),
        space: attr(&attrs, "space")
            .map(|value| parse_u32(value, "column space"))
            .transpose()?,
        separator: parse_on_off_attr(&attrs, "sep", false)?,
        columns: Vec::new(),
    };
    for (name, raw) in direct_nested_children(metadata, xml)? {
        if name == "@raw" && ignorable_nested_raw(&raw) {
            continue;
        }
        if name != "col" {
            return Err(Error::InvalidFormat(format!(
                "invalid child '{name}' in section columns"
            )));
        }
        validate_leaf_content(&raw)?;
        reject_unmodeled_attributes(metadata, &raw, "col")?;
        let attrs = attributes(metadata, &raw)?;
        columns.columns.push(SectionColumn {
            width: parse_u32(
                attr(&attrs, "w")
                    .ok_or_else(|| Error::InvalidFormat("section column omits width".into()))?,
                "column width",
            )?,
            space: attr(&attrs, "space")
                .map(|value| parse_u32(value, "column space"))
                .transpose()?,
        });
    }
    Ok(columns)
}

fn direct_nested_children(metadata: &RootMetadata, xml: &str) -> Result<Vec<(String, String)>> {
    let Some(inner) = element_inner(xml)? else {
        return Ok(Vec::new());
    };
    if inner.is_empty() {
        return Ok(Vec::new());
    }
    let root_name = metadata
        .prefix
        .as_deref()
        .map_or_else(|| "sectPr".to_owned(), |prefix| format!("{prefix}:sectPr"));
    let mut wrapper = format!("<{root_name}");
    for binding in &metadata.bindings {
        wrapper.push_str(" xmlns");
        if let Some(prefix) = &binding.prefix {
            wrapper.push(':');
            wrapper.push_str(prefix);
        }
        wrapper.push_str("=\"");
        wrapper.push_str(&escape(&binding.uri));
        wrapper.push('"');
    }
    wrapper.push('>');
    wrapper.push_str(inner);
    wrapper.push_str(&format!("</{root_name}>"));
    direct_children(&wrapper)
}

fn is_leaf_content_child(name: &str) -> bool {
    matches!(
        name,
        "headerReference"
            | "footerReference"
            | "type"
            | "pgSz"
            | "pgMar"
            | "paperSrc"
            | "lnNumType"
            | "pgNumType"
            | "formProt"
            | "vAlign"
            | "titlePg"
            | "textDirection"
            | "bidi"
            | "rtlGutter"
            | "docGrid"
            | "printerSettings"
            | "noEndnote"
    )
}

fn validate_leaf_content(xml: &str) -> Result<()> {
    let Some(inner) = element_inner(xml)? else {
        return Ok(());
    };
    let mut reader = NsReader::from_reader(inner.as_bytes());
    loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("section leaf namespace resolution failed: {error}"))
        })?;
        match event {
            Event::Eof => return Ok(()),
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Text(_)
            | Event::CData(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Decl(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::Start(_)
            | Event::Empty(_)
            | Event::End(_) => {
                return Err(Error::InvalidFormat(
                    "section leaf contains unsupported nested content".into(),
                ));
            },
        }
    }
}

fn element_inner(xml: &str) -> Result<Option<&str>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let open_end = loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("section namespace resolution failed: {error}"))
        })?;
        match event {
            Event::Start(_) => {
                break usize::try_from(reader.buffer_position()).map_err(|_source_error| {
                    Error::InvalidFormat("section property offset overflow".into())
                })?;
            },
            Event::Empty(_) => return Ok(None),
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => {
                return Err(Error::InvalidFormat("invalid section property".into()));
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "invalid content before nested section property".into(),
                ));
            },
        }
    };
    let mut depth = 1usize;
    let close_start = loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("section property offset overflow".into())
        })?;
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("section namespace resolution failed: {error}"))
        })?;
        match event {
            Event::Start(_) => depth += 1,
            Event::Empty(_) => {},
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid section nesting".into()))?;
                if depth == 0 {
                    break event_start;
                }
            },
            Event::Eof => {
                return Err(Error::InvalidFormat("unterminated section property".into()));
            },
            _ => {},
        }
    };
    let final_position = usize::try_from(reader.buffer_position())
        .map_err(|_source_error| Error::InvalidFormat("section property offset overflow".into()))?;
    if !xml[final_position..].trim().is_empty() {
        return Err(Error::InvalidFormat(
            "section property has trailing content".into(),
        ));
    }
    Ok(Some(&xml[open_end..close_start]))
}

fn ignorable_nested_raw(raw: &str) -> bool {
    let trimmed = raw.trim();
    trimmed.is_empty()
}

#[derive(Debug, Clone, Copy)]
struct ParsedNote<P> {
    format: PageNumberFormat,
    start: Option<u32>,
    restart: Option<NoteNumberRestart>,
    position: Option<P>,
}

fn parse_note_properties<P: NotePos>(metadata: &RootMetadata, xml: &str) -> Result<ParsedNote<P>> {
    let mut result = ParsedNote {
        format: PageNumberFormat::Decimal,
        start: None,
        restart: None,
        position: None,
    };
    let mut seen = std::collections::HashSet::new();
    let mut last_rank = None;
    for (name, raw) in direct_nested_children(metadata, xml)? {
        if name == "@raw" && ignorable_nested_raw(&raw) {
            continue;
        }
        let rank = match name.as_str() {
            "numFmt" => 0,
            "numStart" => 1,
            "numRestart" => 2,
            "pos" => 3,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid note property '{name}'"
                )));
            },
        };
        if last_rank.is_some_and(|last| rank < last) || !seen.insert(name.clone()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate or out-of-order note property '{name}'"
            )));
        }
        last_rank = Some(rank);
        validate_leaf_content(&raw)?;
        reject_unmodeled_attributes(metadata, &raw, &name)?;
        let value = required_attr(metadata, &raw, b"val")?;
        match name.as_str() {
            "numFmt" => result.format = PageNumberFormat::parse(&value)?,
            "numStart" => result.start = Some(parse_u32(&value, "note number start")?),
            "numRestart" => result.restart = Some(NoteNumberRestart::parse(&value)?),
            "pos" => result.position = Some(P::parse(&value)?),
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid note property '{name}'"
                )));
            },
        }
    }
    Ok(result)
}

fn parse_footnotes(metadata: &RootMetadata, xml: &str) -> Result<Footnotes> {
    let parsed = parse_note_properties(metadata, xml)?;
    Ok(Footnotes {
        format: parsed.format,
        start: parsed.start,
        restart: parsed.restart,
        position: parsed.position,
    })
}

fn parse_endnotes(metadata: &RootMetadata, xml: &str) -> Result<Endnotes> {
    let parsed = parse_note_properties(metadata, xml)?;
    Ok(Endnotes {
        format: parsed.format,
        start: parsed.start,
        restart: parsed.restart,
        position: parsed.position,
    })
}

fn parse_grid(metadata: &RootMetadata, xml: &str) -> Result<SectionDocumentGrid> {
    let attrs = attributes(metadata, xml)?;
    Ok(SectionDocumentGrid {
        grid_type: attr(&attrs, "type")
            .map(GridType::parse)
            .transpose()?
            .unwrap_or(GridType::Default),
        line_pitch: attr(&attrs, "linePitch")
            .map(|value| parse_u32(value, "grid line pitch"))
            .transpose()?,
        char_space: attr(&attrs, "charSpace")
            .map(|value| {
                value.parse::<i32>().map_err(|_source_error| {
                    Error::InvalidFormat("invalid grid character space".into())
                })
            })
            .transpose()?,
    })
}

fn parse_page_borders(metadata: &RootMetadata, xml: &str) -> Result<borders::Borders> {
    let attrs = attributes(metadata, xml)?;
    let mut borders = borders::Borders {
        offset_from: attr(&attrs, "offsetFrom")
            .map(OffsetFrom::parse)
            .transpose()?
            .unwrap_or(OffsetFrom::Page),
        z_order: attr(&attrs, "zOrder")
            .map(ZOrder::parse)
            .transpose()?
            .unwrap_or(ZOrder::Back),
        display: attr(&attrs, "display")
            .map(Display::parse)
            .transpose()?
            .unwrap_or(Display::AllPages),
        ..borders::Borders::default()
    };
    let mut last_rank = None;
    for (name, raw) in direct_nested_children(metadata, xml)? {
        if name == "@raw" && ignorable_nested_raw(&raw) {
            continue;
        }
        let edge = match name.as_str() {
            "top" => &mut borders.top,
            "left" => &mut borders.left,
            "bottom" => &mut borders.bottom,
            "right" => &mut borders.right,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid child '{name}' in section page borders"
                )));
            },
        };
        let rank = match name.as_str() {
            "top" => 0,
            "left" => 1,
            "bottom" => 2,
            "right" => 3,
            _ => {
                return Err(Error::InvalidFormat("invalid page border edge".into()));
            },
        };
        if last_rank.is_some_and(|last| rank < last) {
            return Err(Error::InvalidFormat(
                "page border edges are out of schema order".into(),
            ));
        }
        last_rank = Some(rank);
        if edge.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate '{name}' page border edge"
            )));
        }
        reject_unmodeled_attributes(metadata, &raw, &name)?;
        *edge = Some(parse_page_border(metadata, &raw)?);
    }
    Ok(borders)
}

fn parse_page_border(metadata: &RootMetadata, xml: &str) -> Result<borders::Border> {
    validate_leaf_content(xml)?;
    let attrs = attributes(metadata, xml)?;
    let shadow = parse_on_off_attr(&attrs, "shadow", false)?;
    let frame = parse_on_off_attr(&attrs, "frame", false)?;
    Ok(borders::Border {
        style: Style::parse(
            attr(&attrs, "val")
                .ok_or_else(|| Error::InvalidFormat("page border omits style".into()))?,
        )?,
        size: attr(&attrs, "sz")
            .map(|value| parse_u32(value, "page border size"))
            .transpose()?,
        space: attr(&attrs, "space")
            .map(|value| parse_u32(value, "page border space"))
            .transpose()?,
        color: attr(&attrs, "color").map(Color::parse).transpose()?,
        shadow,
        frame,
    })
}

fn parse_line_numbering(metadata: &RootMetadata, xml: &str) -> Result<SectionLineNumbering> {
    let attrs = attributes(metadata, xml)?;
    Ok(SectionLineNumbering {
        count_by: attr(&attrs, "countBy")
            .map(|value| parse_u32(value, "line-number increment"))
            .transpose()?,
        start: attr(&attrs, "start")
            .map(|value| parse_u32(value, "line-number start"))
            .transpose()?,
        distance: attr(&attrs, "distance")
            .map(|value| parse_u32(value, "line-number distance"))
            .transpose()?,
        restart: attr(&attrs, "restart")
            .map(LineNumberRestart::parse)
            .transpose()?,
    })
}

fn parse_on_off(metadata: &RootMetadata, xml: &str) -> Result<bool> {
    let attrs = attributes(metadata, xml)?;
    parse_on_off_attr(&attrs, "val", true)
}

fn parse_on_off_attr(attrs: &[(String, String)], name: &str, default: bool) -> Result<bool> {
    let Some(value) = attr(attrs, name) else {
        return Ok(default);
    };
    match value {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid ST_OnOff value '{value}' for '{name}'"
        ))),
    }
}

fn write_note_properties<P: NotePos>(
    xml: &mut String,
    element: &str,
    format: PageNumberFormat,
    start: Option<u32>,
    restart: Option<NoteNumberRestart>,
    position: Option<P>,
    plan: &NamespacePlan,
) -> Result<()> {
    write!(
        xml,
        "<{}><{} {}=\"{}\"/>",
        plan.word_qname(element),
        plan.word_qname("numFmt"),
        plan.word_attribute_qname("val"),
        format.as_str()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(start) = start {
        write!(
            xml,
            "<{} {}=\"{start}\"/>",
            plan.word_qname("numStart"),
            plan.word_attribute_qname("val")
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(restart) = restart {
        write!(
            xml,
            "<{} {}=\"{}\"/>",
            plan.word_qname("numRestart"),
            plan.word_attribute_qname("val"),
            restart.as_str()
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(position) = position {
        write!(
            xml,
            "<{} {}=\"{}\"/>",
            plan.word_qname("pos"),
            plan.word_attribute_qname("val"),
            position.as_str()
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    write!(xml, "</{}>", plan.word_qname(element)).map_err(|error| Error::Xml(error.to_string()))
}

fn write_footnotes(xml: &mut String, note: &Footnotes, plan: &NamespacePlan) -> Result<()> {
    write_note_properties(
        xml,
        "footnotePr",
        note.format,
        note.start,
        note.restart,
        note.position,
        plan,
    )
}

fn write_endnotes(xml: &mut String, note: &Endnotes, plan: &NamespacePlan) -> Result<()> {
    write_note_properties(
        xml,
        "endnotePr",
        note.format,
        note.start,
        note.restart,
        note.position,
        plan,
    )
}

fn write_page_numbering(
    xml: &mut String,
    numbering: &SectionPageNumbering,
    plan: &NamespacePlan,
) -> Result<()> {
    write!(
        xml,
        "<{} {}=\"{}\"",
        plan.word_qname("pgNumType"),
        plan.word_attribute_qname("fmt"),
        numbering.format.as_str()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(start) = numbering.start {
        write!(xml, " {}=\"{start}\"", plan.word_attribute_qname("start"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(style) = numbering.chapter_style {
        write!(
            xml,
            " {}=\"{style}\"",
            plan.word_attribute_qname("chapStyle")
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(separator) = numbering.chapter_separator {
        write!(
            xml,
            " {}=\"{}\"",
            plan.word_attribute_qname("chapSep"),
            separator.as_str()
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn write_columns(xml: &mut String, columns: &SectionColumns, plan: &NamespacePlan) -> Result<()> {
    write!(
        xml,
        "<{} {}=\"{}\" {}=\"{}\"",
        plan.word_qname("cols"),
        plan.word_attribute_qname("equalWidth"),
        i32::from(columns.equal_width),
        plan.word_attribute_qname("num"),
        columns.count
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(space) = columns.space {
        write!(xml, " {}=\"{space}\"", plan.word_attribute_qname("space"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if columns.separator {
        write!(xml, " {}=\"1\"", plan.word_attribute_qname("sep"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if columns.columns.is_empty() {
        xml.push_str("/>");
    } else {
        xml.push('>');
        for column in &columns.columns {
            write!(
                xml,
                "<{} {}=\"{}\"",
                plan.word_qname("col"),
                plan.word_attribute_qname("w"),
                column.width
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
            if let Some(space) = column.space {
                write!(xml, " {}=\"{space}\"", plan.word_attribute_qname("space"))
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
            xml.push_str("/>");
        }
        write!(xml, "</{}>", plan.word_qname("cols"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}

fn write_grid(xml: &mut String, grid: &SectionDocumentGrid, plan: &NamespacePlan) -> Result<()> {
    write!(
        xml,
        "<{} {}=\"{}\"",
        plan.word_qname("docGrid"),
        plan.word_attribute_qname("type"),
        grid.grid_type.as_str()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(pitch) = grid.line_pitch {
        write!(
            xml,
            " {}=\"{pitch}\"",
            plan.word_attribute_qname("linePitch")
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(space) = grid.char_space {
        write!(
            xml,
            " {}=\"{space}\"",
            plan.word_attribute_qname("charSpace")
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn write_page_borders(
    xml: &mut String,
    borders: &borders::Borders,
    plan: &NamespacePlan,
) -> Result<()> {
    write!(
        xml,
        "<{} {}=\"{}\" {}=\"{}\" {}=\"{}\"",
        plan.word_qname("pgBorders"),
        plan.word_attribute_qname("offsetFrom"),
        borders.offset_from.as_str(),
        plan.word_attribute_qname("zOrder"),
        borders.z_order.as_str(),
        plan.word_attribute_qname("display"),
        borders.display.as_str()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    let edges = [
        ("top", &borders.top),
        ("left", &borders.left),
        ("bottom", &borders.bottom),
        ("right", &borders.right),
    ];
    if edges.iter().all(|(_, edge)| edge.is_none()) {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    for (name, edge) in edges {
        if let Some(border) = edge {
            write_page_border(xml, name, border, plan)?;
        }
    }
    write!(xml, "</{}>", plan.word_qname("pgBorders"))
        .map_err(|error| Error::Xml(error.to_string()))?;
    Ok(())
}

fn write_page_border(
    xml: &mut String,
    name: &str,
    border: &borders::Border,
    plan: &NamespacePlan,
) -> Result<()> {
    write!(
        xml,
        "<{} {}=\"{}\"",
        plan.word_qname(name),
        plan.word_attribute_qname("val"),
        escape(border.style.as_str())
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(size) = border.size {
        write!(xml, " {}=\"{size}\"", plan.word_attribute_qname("sz"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(space) = border.space {
        write!(xml, " {}=\"{space}\"", plan.word_attribute_qname("space"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(color) = border.color {
        match color {
            Color::Auto => write!(xml, " {}=\"auto\"", plan.word_attribute_qname("color"))
                .map_err(|error| Error::Xml(error.to_string()))?,
            Color::Rgb([red, green, blue]) => {
                write!(
                    xml,
                    " {}=\"{red:02X}{green:02X}{blue:02X}\"",
                    plan.word_attribute_qname("color")
                )
                .map_err(|error| Error::Xml(error.to_string()))?;
            },
        }
    }
    if border.shadow {
        write!(xml, " {}=\"1\"", plan.word_attribute_qname("shadow"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if border.frame {
        write!(xml, " {}=\"1\"", plan.word_attribute_qname("frame"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn write_line_numbering(
    xml: &mut String,
    numbering: &SectionLineNumbering,
    plan: &NamespacePlan,
) -> Result<()> {
    write!(xml, "<{}", plan.word_qname("lnNumType"))
        .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(count_by) = numbering.count_by {
        write!(
            xml,
            " {}=\"{count_by}\"",
            plan.word_attribute_qname("countBy")
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(start) = numbering.start {
        write!(xml, " {}=\"{start}\"", plan.word_attribute_qname("start"))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(distance) = numbering.distance {
        write!(
            xml,
            " {}=\"{distance}\"",
            plan.word_attribute_qname("distance")
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(restart) = numbering.restart {
        write!(
            xml,
            " {}=\"{}\"",
            plan.word_attribute_qname("restart"),
            restart.as_str()
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}
