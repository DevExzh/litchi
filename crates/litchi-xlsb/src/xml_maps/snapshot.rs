#![cfg_attr(
    test,
    allow(
        clippy::expect_used,
        reason = "snapshot-only test constructors use panic-on-failure extraction for asserted valid fixtures"
    )
)]

//! Immutable XLSB XML Maps package snapshots.

use std::ops::Range;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::{
    Limits, MappedTable, SingleCellBinding, SingleCellTable, XmlMapConformance, XmlMapInfo,
    XmlMapLimits,
};
use crate::external_link::ExternalLinkLimits;
use crate::package::error::Result;

/// Finite package and codec ceilings for one XLSB XML Maps snapshot.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReadLimits {
    /// Limits for the shared SpreadsheetML MapInfo XML codec.
    pub xml_maps: XmlMapLimits,
    /// Limits for BIFF12 table and single-cell binding streams.
    pub bindings: Limits,
    /// Hierarchical core resource budget charged while traversing the package.
    pub core: litchi_core::Limits,
    /// Maximum OPC parts inspected by this package reader.
    pub max_parts: usize,
    /// Maximum root and part relationships inspected by this package reader.
    pub max_relationships: usize,
    /// Maximum sum of materialized OPC part bytes inspected by this package reader.
    pub max_total_bytes: usize,
    /// Maximum XML bindings accumulated across every table and single-cell part.
    pub max_total_bindings: usize,
    /// Maximum UTF-16 XPath units accumulated across every XML binding.
    pub max_total_xpath_units: usize,
}

impl ReadLimits {
    /// Conservative finite XML Maps package limits.
    pub const DEFAULT: Self = Self {
        xml_maps: XmlMapLimits::DEFAULT,
        bindings: Limits::DEFAULT,
        core: litchi_core::Limits::for_profile(litchi_core::Profile::Server),
        max_parts: 16_384,
        max_relationships: 65_536,
        max_total_bytes: 512 * 1024 * 1024,
        max_total_bindings: 1_000_000,
        max_total_xpath_units: 64 * 1024 * 1024,
    };
}

impl Default for ReadLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A read-only semantic and physical snapshot of XLSB XML Maps.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    map_info: Option<XmlMapInfo>,
    conformance: XmlMapConformance,
    mapped_tables: Vec<MappedTable>,
    single_cell_tables: Vec<SingleCellTable>,
    single_cell_ranges: Vec<Option<Range<usize>>>,
    limits: ReadLimits,
    source: SourceState,
}

impl Snapshot {
    /// Read an XLSB XML Maps snapshot using conservative finite defaults.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::read_with_limits(package, ReadLimits::DEFAULT)
    }

    /// Read an XLSB XML Maps snapshot with explicit finite limits.
    pub fn read_with_limits(package: &OpcPackage, limits: ReadLimits) -> Result<Self> {
        Self::read_with_limits_and_external_link_limits(
            package,
            limits,
            ExternalLinkLimits::default(),
        )
    }

    /// Read an XLSB XML Maps snapshot with independent XML Maps and
    /// external-link resource policies.
    pub fn read_with_limits_and_external_link_limits(
        package: &OpcPackage,
        limits: ReadLimits,
        external_link_limits: ExternalLinkLimits,
    ) -> Result<Self> {
        super::package::preflight(package, limits)?;
        let workbook = crate::Workbook::from_opc_package_with_external_link_limits(
            package.clone(),
            external_link_limits,
        )?;
        let worksheets = workbook.xml_maps_worksheet_parts()?;
        Self::read_for_worksheets(package, worksheets, limits)
    }

    pub(crate) fn read_for_worksheets(
        package: &OpcPackage,
        worksheets: Vec<PackURI>,
        limits: ReadLimits,
    ) -> Result<Self> {
        let loaded = super::package::read_with_limits(package, worksheets, limits)?;
        let workbook = package.main_document_part()?;
        let source = SourceState::capture(
            package,
            workbook,
            &loaded.worksheet_parts,
            &loaded.dependency_parts,
        )?;
        Ok(Self {
            map_info: loaded.map_info,
            conformance: loaded.conformance,
            mapped_tables: loaded.mapped_tables,
            single_cell_tables: loaded.single_cell_tables,
            single_cell_ranges: loaded.single_cell_ranges,
            limits,
            source,
        })
    }

    /// Borrow the workbook MapInfo catalog, if the workbook owns one.
    pub fn map_info(&self) -> Option<&XmlMapInfo> {
        self.map_info.as_ref()
    }

    /// Exact source XML for the workbook MapInfo part, when present.
    pub fn source_xml(&self) -> Option<&[u8]> {
        self.source
            .dependencies
            .iter()
            .find(|part| {
                part.content_type == litchi_ooxml_common::spreadsheet_xml_maps::CONTENT_TYPE
            })
            .map(SourcePart::bytes)
    }

    /// Start a detached, source-bound XML Maps transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Borrow the individual XML maps, or an empty slice when no catalog exists.
    pub fn maps(&self) -> &[super::XmlMap] {
        self.map_info.as_ref().map_or(&[], |info| &info.maps)
    }

    /// Borrow XML table bindings in workbook worksheet discovery order.
    pub fn mapped_tables(&self) -> &[MappedTable] {
        &self.mapped_tables
    }

    /// Borrow single-cell table bindings in workbook worksheet discovery order.
    pub fn single_cell_tables(&self) -> &[SingleCellTable] {
        &self.single_cell_tables
    }

    /// Borrow the optional single-cell binding collection for one zero-based worksheet.
    pub fn single_cell_bindings(&self, sheet_index: usize) -> Option<&[SingleCellBinding]> {
        self.single_cell_ranges
            .get(sheet_index)
            .and_then(Option::as_ref)
            .map(|range| &self.single_cell_tables[range.clone()])
    }

    /// Return the exact limits used to parse this snapshot.
    pub const fn limits(&self) -> ReadLimits {
        self.limits
    }

    /// Namespace conformance of the owned MapInfo relationship.
    pub const fn conformance(&self) -> XmlMapConformance {
        self.conformance
    }

    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub(crate) fn source(&self) -> &SourceState {
        &self.source
    }

    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(crate) fn materialized(
        before: &Self,
        map_info: Option<XmlMapInfo>,
        conformance: XmlMapConformance,
        mapped_tables: Vec<MappedTable>,
        single_cell_bindings: Vec<Option<Vec<SingleCellBinding>>>,
        source: SourceState,
    ) -> Result<Self> {
        let (single_cell_tables, single_cell_ranges) =
            flatten_binding_groups(single_cell_bindings, before.limits.max_total_bindings)?;
        Ok(Self {
            map_info,
            conformance,
            mapped_tables,
            single_cell_tables,
            single_cell_ranges,
            limits: before.limits,
            source,
        })
    }

    pub(crate) fn binding_groups(&self) -> Vec<Option<Vec<SingleCellBinding>>> {
        self.single_cell_ranges
            .iter()
            .map(|range| {
                range
                    .as_ref()
                    .map(|range| self.single_cell_tables[range.clone()].to_vec())
            })
            .collect()
    }

    pub(crate) fn binding_groups_equal(&self, groups: &[Option<Vec<SingleCellBinding>>]) -> bool {
        self.single_cell_ranges.len() == groups.len()
            && groups.iter().enumerate().all(|(sheet_index, group)| {
                self.single_cell_bindings(sheet_index) == group.as_deref()
            })
    }

    pub(crate) fn without_signatures(mut self) -> Self {
        self.source.remove_signatures();
        self
    }

    #[cfg(test)]
    pub(crate) fn empty_transaction_fixture(signed: bool) -> Self {
        let workbook_name = PackURI::new("/xl/workbook.bin").expect("fixture URI");
        let mut root_relationships = Vec::new();
        let mut part_names = vec![workbook_name.clone()];
        if signed {
            root_relationships.push(SourceRelationship {
                id: "rIdSignature".to_string(),
                relationship_type:
                    litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_string(),
                target: "_xmlsignatures/origin.sigs".to_string(),
                mode: TargetMode::Internal,
            });
            part_names
                .push(PackURI::new("/_xmlsignatures/origin.sigs").expect("signature fixture URI"));
        }
        Self {
            map_info: None,
            conformance: XmlMapConformance::Transitional,
            mapped_tables: Vec::new(),
            single_cell_tables: Vec::new(),
            single_cell_ranges: Vec::new(),
            limits: ReadLimits::DEFAULT,
            source: SourceState {
                root_relationships,
                workbook: SourcePart {
                    part_name: workbook_name,
                    content_type: litchi_opc::constants::content_type::XLSB_BIN.to_string(),
                    bytes: Arc::new(Vec::new()),
                    relationships: Vec::new(),
                },
                worksheets: Vec::new(),
                dependencies: Vec::new(),
                part_names,
            },
        }
    }

    #[cfg(test)]
    pub(crate) fn with_conformance_fixture(mut self, conformance: XmlMapConformance) -> Self {
        self.conformance = conformance;
        self
    }

    #[cfg(test)]
    pub(crate) fn non_xml_table_transaction_fixture() -> Self {
        let mut snapshot = Self::empty_transaction_fixture(false);
        let part_name = PackURI::new("/xl/tables/table1.bin").expect("table fixture URI");
        let table = crate::package::table::Table {
            id: 2,
            range: crate::package::table::Range {
                first_row: 0,
                last_row: 1,
                first_column: 0,
                last_column: 0,
            },
            table_type: crate::package::table::Type::Range,
            header_row_count: 1,
            columns: vec![crate::package::table::Column {
                id: 1,
                ..crate::package::table::Column::default()
            }],
            ..crate::package::table::Table::default()
        };
        let bytes =
            crate::package::table::write::write_table_part(&table).expect("non-XML table fixture");
        snapshot.source.dependencies.push(SourcePart {
            part_name: part_name.clone(),
            content_type: "application/vnd.ms-excel.table".to_string(),
            bytes: Arc::new(bytes),
            relationships: Vec::new(),
        });
        snapshot.source.part_names.push(part_name);
        snapshot
    }
}

fn flatten_binding_groups(
    groups: Vec<Option<Vec<SingleCellBinding>>>,
    max_total_bindings: usize,
) -> Result<(Vec<SingleCellBinding>, Vec<Option<Range<usize>>>)> {
    let total = groups
        .iter()
        .filter_map(Option::as_ref)
        .try_fold(0usize, |count, values| count.checked_add(values.len()))
        .ok_or(crate::package::error::Error::CapacityOverflow {
            resource: "single-cell snapshot binding count",
        })?;
    if total > max_total_bindings {
        return Err(crate::package::error::Error::InvalidLength {
            expected: max_total_bindings,
            found: total,
        });
    }
    let mut values = Vec::new();
    values
        .try_reserve_exact(total)
        .map_err(|source| crate::package::error::Error::Allocation {
            resource: "single-cell snapshot bindings",
            source,
        })?;
    let mut ranges = Vec::new();
    ranges.try_reserve_exact(groups.len()).map_err(|source| {
        crate::package::error::Error::Allocation {
            resource: "single-cell snapshot range index",
            source,
        }
    })?;
    for group in groups {
        match group {
            Some(mut group) => {
                let start = values.len();
                values.append(&mut group);
                ranges.push(Some(start..values.len()));
            },
            None => ranges.push(None),
        }
    }
    Ok((values, ranges))
}

/// Exact, cloneable package topology retained for future patch materialization.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    pub(crate) root_relationships: Vec<SourceRelationship>,
    pub(crate) workbook: SourcePart,
    pub(crate) worksheets: Vec<SourcePart>,
    pub(crate) dependencies: Vec<SourcePart>,
    pub(crate) part_names: Vec<PackURI>,
}

#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
impl SourceState {
    fn capture(
        package: &OpcPackage,
        workbook: &dyn Part,
        worksheets: &[PackURI],
        dependencies: &[PackURI],
    ) -> Result<Self> {
        Ok(Self {
            root_relationships: relationships(package.rels().iter()),
            workbook: SourcePart::from_part(workbook),
            worksheets: worksheets
                .iter()
                .map(|name| {
                    package
                        .get_part(name)
                        .map(SourcePart::from_part)
                        .map_err(Into::into)
                })
                .collect::<Result<Vec<_>>>()?,
            dependencies: dependencies
                .iter()
                .map(|name| {
                    package
                        .get_part(name)
                        .map(SourcePart::from_part)
                        .map_err(Into::into)
                })
                .collect::<Result<Vec<_>>>()?,
            part_names: {
                let mut names = package
                    .iter_parts()
                    .map(|part| part.partname().clone())
                    .collect::<Vec<_>>();
                names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
                names
            },
        })
    }

    pub(crate) fn workbook(&self) -> &SourcePart {
        &self.workbook
    }

    pub(crate) fn worksheets(&self) -> &[SourcePart] {
        &self.worksheets
    }

    pub(crate) fn dependencies(&self) -> &[SourcePart] {
        &self.dependencies
    }

    pub(crate) fn remove_signatures(&mut self) {
        self.root_relationships.retain(|relationship| {
            relationship.relationship_type
                != litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN
        });
        self.part_names
            .retain(|name| !name.as_str().starts_with("/_xmlsignatures/"));
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourcePart {
    pub(crate) part_name: PackURI,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
    pub(crate) relationships: Vec<SourceRelationship>,
}

#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
impl SourcePart {
    fn from_part(part: &dyn Part) -> Self {
        Self {
            part_name: part.partname().clone(),
            content_type: part.content_type().to_string(),
            bytes: part.blob_arc(),
            relationships: relationships(part.rels().iter()),
        }
    }

    pub(crate) fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    pub(crate) fn content_type(&self) -> &str {
        &self.content_type
    }

    pub(crate) fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    pub(crate) fn relationships(&self) -> &[SourceRelationship] {
        &self.relationships
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceRelationship {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target: String,
    pub(crate) mode: TargetMode,
}

#[allow(
    dead_code,
    reason = "retained for BIFF12 codec completeness and staged host integration"
)]
impl SourceRelationship {
    pub(crate) fn id(&self) -> &str {
        &self.id
    }
    pub(crate) fn relationship_type(&self) -> &str {
        &self.relationship_type
    }
    pub(crate) fn target(&self) -> &str {
        &self.target
    }
    pub(crate) const fn mode(&self) -> TargetMode {
        self.mode
    }
}

fn relationships<'a>(
    iter: impl Iterator<Item = &'a litchi_opc::Relationship>,
) -> Vec<SourceRelationship> {
    let mut values = iter
        .map(|relationship| SourceRelationship {
            id: relationship.r_id().to_string(),
            relationship_type: relationship.reltype().to_string(),
            target: relationship.target_ref().to_string(),
            mode: relationship.target_mode(),
        })
        .collect::<Vec<_>>();
    values.sort_by(|left, right| left.id.cmp(&right.id));
    values
}
