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

/// The strictest payload ceiling must be applied by the writer itself. This
/// keeps a caller's aggregate policy from being discovered only after the
/// serializer has already allocated the output buffer.
pub(crate) fn output_ceiling(limits: &litchi_opc::SourceLimits) -> usize {
    let ceiling = limits
        .max_part_bytes()
        .min(limits.max_total_part_bytes())
        .min(usize::MAX as u64);
    usize::try_from(ceiling).unwrap_or(usize::MAX)
}

/// Admit an OPC part URI under the ZIP member-name policy. The policy applies
/// to the archive spelling (`membername`), while relationship target limits
/// are checked separately on the relative reference.
pub(crate) fn check_part_member_name(
    part: &litchi_opc::PackURI,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    check_limit(
        litchi_opc::SourceLimitKind::ArchiveMemberNameBytes,
        part.membername().len() as u64,
        limits.max_archive_member_name_bytes(),
    )
}

/// Validate the lexical form used by relationship `Id` attributes before a
/// raw value is serialized or cloned into a staged model.
pub(crate) fn validate_relationship_id(value: &str) -> crate::Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(crate::Error::Invalid(
            "relationship ID cannot be empty".into(),
        ));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return Err(crate::Error::Invalid(
            "relationship ID is not a valid XML identifier".into(),
        ));
    }
    Ok(())
}

/// Check the identity slots that a raw insertion will occupy before its
/// typed payload is cloned or serialized.
pub(crate) fn check_new_relationship_identity(
    package: &litchi_opc::OpcPackage,
    owner: &litchi_opc::PackURI,
    relationship_id: &str,
    target: &litchi_opc::PackURI,
) -> crate::Result<()> {
    validate_relationship_id(relationship_id)?;
    let owner_part = package
        .part_metadata(owner)
        .ok_or_else(|| litchi_opc::OpcError::PartNotFound(owner.to_string()))?;
    if owner_part.rels().get(relationship_id).is_some() {
        return Err(crate::Error::Invalid(
            "relationship ID already exists on the owner".into(),
        ));
    }
    if package.part_metadata(target).is_some() {
        return Err(crate::Error::Invalid(
            "raw insertion target part already exists".into(),
        ));
    }
    Ok(())
}

/// Admit the graph-count portion of a raw insertion before the caller's
/// payload is cloned or serialized. The byte budget is checked after the
/// bounded writer knows the replacement length; part counts do not need that
/// materialization.
pub(crate) fn check_new_part_count(
    package: &litchi_opc::OpcPackage,
    added_parts: usize,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    let count = package
        .part_count()
        .checked_add(added_parts)
        .ok_or_else(|| {
            limit_error(
                litchi_opc::SourceLimitKind::Parts,
                u64::MAX,
                limits.max_parts() as u64,
            )
        })?;
    check_limit(
        litchi_opc::SourceLimitKind::Parts,
        count as u64,
        limits.max_parts() as u64,
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

/// The package-wide metrics needed by one raw insertion admission. This is
/// deliberately an operation-local value: it is never reused after the
/// package can change, so the relationship and payload proof cannot become
/// stale between admission and publication.
pub(crate) struct PackageBudget {
    relationship_parts: usize,
    relationship_bytes: u64,
    relationship_events: u64,
    relationship_count: usize,
    part_bytes: u64,
    owner_plan: litchi_opc::RelationshipSourcePlan,
}

impl PackageBudget {
    /// Refuse an already over-limit package before a raw operation captures
    /// source XML or stages its proposed relationship plan.
    fn check_baseline(
        &self,
        package: &litchi_opc::OpcPackage,
        limits: &litchi_opc::SourceLimits,
    ) -> crate::Result<()> {
        let part_count = package.part_count();
        if part_count > limits.max_parts() {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::Parts,
                part_count as u64,
                limits.max_parts() as u64,
            ));
        }
        if self.part_bytes > limits.max_total_part_bytes() {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::TotalPartBytes,
                self.part_bytes,
                limits.max_total_part_bytes(),
            ));
        }
        if self.relationship_parts > limits.max_relationship_parts() {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::RelationshipParts,
                self.relationship_parts as u64,
                limits.max_relationship_parts() as u64,
            ));
        }
        if self.relationship_bytes > limits.max_total_relationship_xml_bytes() as u64 {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::TotalRelationshipXmlBytes,
                self.relationship_bytes,
                limits.max_total_relationship_xml_bytes() as u64,
            ));
        }
        if self.relationship_events > limits.max_total_relationship_xml_events() as u64 {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::TotalRelationshipXmlEvents,
                self.relationship_events,
                limits.max_total_relationship_xml_events() as u64,
            ));
        }
        if self.relationship_count > limits.max_total_relationships() {
            return Err(limit_error(
                litchi_opc::SourceLimitKind::TotalRelationships,
                self.relationship_count as u64,
                limits.max_total_relationships() as u64,
            ));
        }
        Ok(())
    }
}

/// Admit one raw insertion's relationships and aggregate payload in one
/// package pass before any relationship mutation or payload publication runs.
/// Existing source relationship members retain their exact bytes; the
/// affected owner uses the source-bound fragment plan that the mutation will
/// publish. The ledger includes every relationship owner and every existing
/// part, preserving the cross-owner proof without paying for two independent
/// package scans.
pub(crate) fn prepare_package_budget(
    package: &litchi_opc::OpcPackage,
    owner: &litchi_opc::PackURI,
    additions: &[(String, String, String)],
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<PackageBudget> {
    if additions.len() > limits.max_relationships_per_part() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::RelationshipsPerPart,
            additions.len() as u64,
            limits.max_relationships_per_part() as u64,
        ));
    }

    let owner_plan = package.plan_source_relationships(owner, *limits)?;
    // The source plan above is output-free. Admit the current package ledger
    // before capturing a source token or cloning the proposed edit plan. A
    // package opened under the policy already satisfies this census, but the
    // low-level OPC authoring methods can add payloads or relationships without
    // re-running the aggregate checks; raw edits must refuse that state before
    // allocating another source-bound value.
    let budget = package_budget(package, owner, owner_plan, limits)?;
    budget.check_baseline(package, limits)?;
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
    let source = package.source_relationships(owner, *limits)?;
    let plan = source.plan_edit(&edits, &[], *limits)?;
    let relationship_parts = budget
        .relationship_parts
        .saturating_sub(usize::from(budget.owner_plan.member_present()))
        .saturating_add(usize::from(plan.member_present()));
    if relationship_parts > limits.max_relationship_parts() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::RelationshipParts,
            relationship_parts as u64,
            limits.max_relationship_parts() as u64,
        ));
    }
    let total_bytes = budget
        .relationship_bytes
        .saturating_sub(budget.owner_plan.final_len() as u64)
        .saturating_add(plan.final_len() as u64);
    if total_bytes > limits.max_total_relationship_xml_bytes() as u64 {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::TotalRelationshipXmlBytes,
            total_bytes,
            limits.max_total_relationship_xml_bytes() as u64,
        ));
    }
    let total_events = budget
        .relationship_events
        .saturating_sub(budget.owner_plan.event_count() as u64)
        .saturating_add(plan.event_count() as u64);
    if total_events > limits.max_total_relationship_xml_events() as u64 {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::TotalRelationshipXmlEvents,
            total_events,
            limits.max_total_relationship_xml_events() as u64,
        ));
    }
    let total_count = budget
        .relationship_count
        .saturating_sub(budget.owner_plan.relationship_count())
        .saturating_add(plan.relationship_count());
    if total_count > limits.max_total_relationships() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::TotalRelationships,
            total_count as u64,
            limits.max_total_relationships() as u64,
        ));
    }
    Ok(budget)
}

/// Check the aggregate payload budget using the package ledger prepared
/// before serializers or relationship mutations ran. The package is still
/// unchanged, so the ledger is valid for this final admission step.
pub(crate) fn check_part_budget_with_ledger(
    package: &litchi_opc::OpcPackage,
    budget: &PackageBudget,
    added_parts: usize,
    added_bytes: usize,
    replacements: &[(&litchi_opc::PackURI, usize)],
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<()> {
    let count = package
        .part_count()
        .checked_add(added_parts)
        .ok_or_else(|| {
            limit_error(
                litchi_opc::SourceLimitKind::Parts,
                u64::MAX,
                limits.max_parts() as u64,
            )
        })?;
    if count > limits.max_parts() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::Parts,
            count as u64,
            limits.max_parts() as u64,
        ));
    }
    let mut total = budget.part_bytes;
    for (name, replacement_len) in replacements {
        let old = package
            .part_metadata(name)
            .ok_or_else(|| litchi_opc::OpcError::PartNotFound(name.to_string()))?
            .payload_len()? as u64;
        total = total
            .saturating_sub(old)
            .saturating_add(*replacement_len as u64);
    }
    let total = total.saturating_add(added_bytes as u64);
    if total > limits.max_total_part_bytes() {
        return Err(limit_error(
            litchi_opc::SourceLimitKind::TotalPartBytes,
            total,
            limits.max_total_part_bytes(),
        ));
    }
    Ok(())
}

/// Return relationship and payload totals from one operation-local package
/// pass. The payload projection is intentionally size-only: aggregate
/// part-byte admission must account for the exact bytes that would remain
/// after the replacement without inflating deferred source members.
fn package_budget(
    package: &litchi_opc::OpcPackage,
    owner: &litchi_opc::PackURI,
    owner_plan: litchi_opc::RelationshipSourcePlan,
    limits: &litchi_opc::SourceLimits,
) -> crate::Result<PackageBudget> {
    let root =
        litchi_opc::PackURI::new("/").map_err(|error| crate::Error::Invalid(error.to_string()))?;
    let mut relationship_parts = 0usize;
    let mut relationship_bytes = 0u64;
    let mut relationship_events = 0u64;
    let mut relationship_count = 0usize;
    let mut part_bytes = 0u64;
    let mut owner_plan_seen = false;

    let mut add = |candidate: &litchi_opc::PackURI,
                   payload_len: Option<usize>|
     -> crate::Result<()> {
        if let Some(payload_len) = payload_len {
            if payload_len as u64 > limits.max_part_bytes() {
                return Err(limit_error(
                    litchi_opc::SourceLimitKind::PartBytes,
                    payload_len as u64,
                    limits.max_part_bytes(),
                ));
            }
        }
        let plan = if candidate == owner {
            owner_plan_seen = true;
            owner_plan
        } else {
            package.plan_source_relationships(candidate, *limits)?
        };
        relationship_parts = relationship_parts.saturating_add(usize::from(plan.member_present()));
        relationship_bytes = relationship_bytes.saturating_add(plan.final_len() as u64);
        relationship_events = relationship_events.saturating_add(plan.event_count() as u64);
        relationship_count = relationship_count.saturating_add(plan.relationship_count());
        if let Some(payload_len) = payload_len {
            part_bytes = part_bytes.saturating_add(payload_len as u64);
        }
        Ok(())
    };

    add(&root, None)?;
    for part in package.iter_parts() {
        let partname = part.partname();
        let payload_len = part.payload_len()?;
        add(partname, Some(payload_len))?;
    }
    if !owner_plan_seen {
        return Err(crate::Error::Invalid(
            "relationship owner is not present in package".into(),
        ));
    }
    Ok(PackageBudget {
        relationship_parts,
        relationship_bytes,
        relationship_events,
        relationship_count,
        part_bytes,
        owner_plan,
    })
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
