//! Validated low-level workbook catalog records.
//!
//! Relationship IDs and native sheet IDs are deliberately isolated here.

mod catalog;
pub(crate) mod catalog_edit;
pub(crate) mod compact;
mod formula;
pub mod namespace;
pub(crate) mod properties_edit;
pub(crate) mod recalc;
pub(crate) mod reference_edit;
pub(crate) mod reference_scan;
pub mod shared_strings;
pub(crate) mod sheet_view_edit;
pub(crate) mod strings;
pub(crate) mod styles;
pub(crate) mod styles_edit;
pub mod web;
pub(crate) mod worksheet;
pub(crate) mod worksheet_property;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

/// Bounded source selection of one worksheet coordinate or rectangular range.
///
/// This surface is an internal format-owner API. `NotEligible` is a
/// capability result, not full worksheet validation; callers must fall back
/// to the materialized worksheet parser whenever it is returned. Eligible
/// range results are sparse physical records and include
/// [`selected_worksheet::SelectedDependencies`] for the zero-based
/// shared-string and direct-cell-style indexes observed across the complete
/// stream. The one-cell [`selected_worksheet::scan`] function is retained as
/// a compatibility wrapper around [`selected_worksheet::scan_range`].
pub mod selected_worksheet {
    pub use super::worksheet::selected::{
        NotEligibleReason, RangeScanOutcome, ScanOutcome, SelectedCell, SelectedCells,
        SelectedDependencies, SelectedPayload, SelectedRecord, StreamResult, scan, scan_range,
    };
}

pub(crate) use catalog::parse_catalog_repairing;
pub use catalog::{parse_catalog, parse_catalog_with_mce, parse_sheet};

/// Check the caller's source budget before a raw parser materializes XML.
pub(crate) fn check_source_bytes(
    actual: usize,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    let maximum = limits.max_part_bytes().min(usize::MAX as u64);
    let maximum = usize::try_from(maximum).unwrap_or(usize::MAX);
    if actual > maximum {
        return Err(crate::Error::Package(litchi_opc::OpcError::Limit {
            kind: litchi_opc::SourceLimitKind::PartBytes,
            observed: actual as u64,
            maximum: maximum as u64,
        }));
    }
    Ok(())
}

pub(crate) fn check_limit(
    kind: litchi_opc::SourceLimitKind,
    actual: u64,
    maximum: u64,
) -> crate::Result<()> {
    if actual > maximum {
        return Err(crate::Error::Package(litchi_opc::OpcError::Limit {
            kind,
            observed: actual,
            maximum,
        }));
    }
    Ok(())
}

pub(crate) fn check_output_bytes(
    actual: usize,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    check_limit(
        litchi_opc::SourceLimitKind::TotalPartBytes,
        actual as u64,
        limits.max_total_part_bytes(),
    )
}

/// Admit an internal relationship target before its relative spelling is
/// allocated. `PackURI::relative_ref_len` follows the writer without building
/// the target string, so a tight target policy cannot be bypassed by an
/// absolute native part name whose relative spelling is shorter.
pub(crate) fn check_relationship_target(
    target: &litchi_opc::PackURI,
    owner_base_uri: &str,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    check_limit(
        litchi_opc::SourceLimitKind::RelationshipTargetBytes,
        target.relative_ref_len(owner_base_uri) as u64,
        limits.max_relationship_target_bytes() as u64,
    )
}

/// Keep a raw operation inside both the caller's requested policy and the
/// policy retained when the package was opened. A later raw call may tighten
/// that policy, but it cannot widen the package's original admission.
pub(crate) fn effective_limits(
    package: &litchi_opc::OpcPackage,
    requested: litchi_opc::SourceLimits,
) -> litchi_opc::SourceLimits {
    requested.min(package.read_limits())
}

/// Admit one XML member under the caller's raw read policy before a format
/// parser materializes its model.  The format codecs still perform their
/// grammar and format-specific checks; this pass supplies the caller-owned
/// event, depth, and attribute ceilings that those codecs do not own.
pub(crate) fn check_xml_limits(
    content: &[u8],
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    check_source_bytes(content.len(), limits)?;
    let mut reader = quick_xml::Reader::from_reader(content);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut events = 0usize;
    loop {
        let event = reader.read_event().map_err(|error| {
            crate::Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
        })?;
        events = events.checked_add(1).ok_or_else(|| {
            limit_error(
                litchi_opc::SourceLimitKind::Events,
                u64::MAX,
                limits.max_events() as u64,
            )
        })?;
        if events > limits.max_events() {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::Events,
                events as u64,
                limits.max_events() as u64,
            ));
        }
        match event {
            quick_xml::events::Event::Start(element) => {
                check_xml_attributes(&element, limits)?;
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit_error(
                        litchi_opc::SourceLimitKind::Depth,
                        u64::MAX,
                        limits.max_depth() as u64,
                    )
                })?;
                if depth > limits.max_depth() {
                    return Err(limit_error(
                        litchi_opc::SourceLimitKind::Depth,
                        depth as u64,
                        limits.max_depth() as u64,
                    ));
                }
            },
            quick_xml::events::Event::Empty(element) => {
                check_xml_attributes(&element, limits)?;
                let observed = depth.checked_add(1).ok_or_else(|| {
                    limit_error(
                        litchi_opc::SourceLimitKind::Depth,
                        u64::MAX,
                        limits.max_depth() as u64,
                    )
                })?;
                if observed > limits.max_depth() {
                    return Err(limit_error(
                        litchi_opc::SourceLimitKind::Depth,
                        observed as u64,
                        limits.max_depth() as u64,
                    ));
                }
            },
            quick_xml::events::Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::Error::Invalid("XML closing element has no open parent".into())
                })?;
            },
            quick_xml::events::Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn check_xml_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| {
            crate::Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
        })?;
        let observed = attribute
            .key
            .as_ref()
            .len()
            .saturating_add(attribute.value.len());
        if observed > limits.max_xml_attribute_bytes() {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::XmlAttributeBytes,
                observed as u64,
                limits.max_xml_attribute_bytes() as u64,
            ));
        }
    }
    Ok(())
}

fn limit_error(kind: litchi_opc::SourceLimitKind, observed: u64, maximum: u64) -> crate::Error {
    crate::Error::Package(litchi_opc::OpcError::Limit {
        kind,
        observed,
        maximum,
    })
}

/// Admit a batch of newly authored internal relationships before any payload
/// serializer or relationship mutation runs. Existing source relationship
/// members retain their exact bytes; the affected owner uses the source-bound
/// fragment plan that the mutation will publish. The admission is scoped to
/// this owner so a raw edit does not add an unrelated package-wide scan to its
/// already bounded family read.
pub(crate) fn check_relationship_insertions(
    package: &litchi_opc::OpcPackage,
    owner: &litchi_opc::PackURI,
    additions: &[(String, String, String)],
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    if additions.len() > limits.max_relationships_per_part() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::RelationshipsPerPart,
            additions.len() as u64,
            limits.max_relationships_per_part() as u64,
        ));
    }

    let mut edits = Vec::new();
    edits.try_reserve_exact(additions.len()).map_err(|source| {
        crate::Error::Package(litchi_opc::OpcError::Allocation {
            resource: "XLSX raw relationship admission",
            source,
        })
    })?;
    for (id, reltype, target) in additions {
        edits.push(litchi_opc::RelationshipEdit {
            id,
            reltype,
            target,
            mode: litchi_opc::TargetMode::Internal,
        });
    }
    let source = package.source_relationships_with_limits(owner, *limits)?;
    let plan = source.plan_edit(&edits, &[], *limits)?;
    let relationship_parts = usize::from(plan.member_present());
    if relationship_parts > limits.max_relationship_parts() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::RelationshipParts,
            relationship_parts as u64,
            limits.max_relationship_parts() as u64,
        ));
    }
    if plan.final_len() > limits.max_total_relationship_xml_bytes() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::TotalRelationshipXmlBytes,
            plan.final_len() as u64,
            limits.max_total_relationship_xml_bytes() as u64,
        ));
    }
    if plan.event_count() > limits.max_total_relationship_xml_events() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::TotalRelationshipXmlEvents,
            plan.event_count() as u64,
            limits.max_total_relationship_xml_events() as u64,
        ));
    }
    if plan.relationship_count() > limits.max_total_relationships() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::TotalRelationships,
            plan.relationship_count() as u64,
            limits.max_total_relationships() as u64,
        ));
    }
    Ok(())
}

/// Parse only the direct worksheet-grid defaults from one `WorksheetPart`.
///
/// This focused read path shares the canonical [`crate::layout::Defaults`]
/// model and validation with the full worksheet parser without materializing
/// cells, rows, columns, or merges. Markup compatibility and the supported
/// `x14ac:dyDescent` extension are processed using the same bounded rules as
/// the full parser.
pub fn parse_worksheet_defaults(
    content: &[u8],
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<Option<crate::layout::Defaults>> {
    check_source_bytes(content.len(), limits)?;
    worksheet::parse_defaults(content)
}

/// Physical visibility value retained from `sheet/@state`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Visibility {
    Visible,
    Hidden,
    VeryHidden,
    Unknown(Box<str>),
}

/// One validated `workbook/sheets/sheet` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sheet {
    pub name: String,
    pub relationship_id: String,
    pub sheet_id: u32,
    pub visibility: Visibility,
}

/// One inert `SpreadsheetML` defined-name record.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DefinedName {
    pub name: String,
    pub reference: String,
    pub comment: Option<String>,
    pub local_sheet_id: Option<u32>,
    pub custom_menu: Option<String>,
    pub description: Option<String>,
    pub help: Option<String>,
    pub status_bar: Option<String>,
    pub shortcut_key: Option<String>,
    pub hidden: bool,
    pub function: bool,
    pub vb_procedure: bool,
    pub xlm: bool,
    pub function_group_id: Option<u32>,
    pub publish_to_server: bool,
    pub workbook_parameter: bool,
}

/// Workbook-level pivot-cache relationship record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotCache {
    pub cache_id: u32,
    pub relationship_id: String,
}

/// Validated catalog extracted from `workbook.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Catalog {
    pub sheets: Vec<Sheet>,
    pub active_sheet_index: usize,
    pub uses_1904_date_system: bool,
    pub defined_names: Vec<DefinedName>,
    pub pivot_caches: Vec<PivotCache>,
    pub external_reference_ids: Vec<String>,
}
