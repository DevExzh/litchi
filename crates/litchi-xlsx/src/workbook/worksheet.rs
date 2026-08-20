//! Worksheet-owned semantic reads and relationship projections.
//!
//! The immutable [`super::Worksheet`] handle owns the package boundary.  This
//! module keeps XML parsing and relationship traversal below that facade so
//! callers do not need to discover worksheet parts or duplicate host-side
//! orchestration.  Every operation is read-only and returns the canonical
//! model from its semantic owner module.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{Part, Relationships};

use super::{Worksheet, WorksheetKind};
use crate::error::{Result, invalid};

/// One worksheet table together with its physical relationship identity.
///
/// The table model itself remains [`crate::table::Table`]; this wrapper only
/// records the package edge needed to edit or inspect the owning relationship.
#[derive(Debug, Clone)]
pub struct TablePart {
    relationship_id: Box<str>,
    part_name: Box<str>,
    table: crate::table::Table,
}

impl TablePart {
    /// Relationship ID as stored by the worksheet part.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Absolute OPC part name of the table resource.
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    /// Parsed table declaration.
    #[must_use]
    pub fn table(&self) -> &crate::table::Table {
        &self.table
    }
}

/// One array-formula anchor discovered in a worksheet snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ArrayFormula {
    /// Checked one-based worksheet row of the formula anchor.
    pub row: u32,
    /// Checked one-based worksheet column of the formula anchor.
    pub column: u32,
    /// Formula expression without a leading `=`.
    pub formula: Box<str>,
    /// Inclusive A1 range supplied by the producer, when present.
    pub range: Option<Box<str>>,
}

fn part(sheet: &Worksheet) -> Result<&dyn Part> {
    if sheet.kind() != WorksheetKind::Worksheet {
        return Err(crate::Error::NotWorksheet {
            sheet: sheet.name().to_owned(),
        });
    }
    sheet
        .owner
        .package
        .get_part(&sheet.data.part_uri)
        .map_err(Into::into)
}

fn xml(sheet: &Worksheet) -> Result<&[u8]> {
    Ok(part(sheet)?.blob())
}

fn relationships(sheet: &Worksheet) -> Result<&Relationships> {
    Ok(part(sheet)?.rels())
}

/// Parse the worksheet's direct auto-filter declaration.
pub(crate) fn auto_filter(sheet: &Worksheet) -> Result<Option<crate::auto_filter::Definition>> {
    crate::auto_filter::parse_auto_filter(xml(sheet)?)
}

/// Parse all conditional-formatting containers and associate their differential
/// formats with the workbook styles resource.
pub(crate) fn conditional_formattings(
    sheet: &Worksheet,
) -> Result<Vec<crate::conditional_formatting::Formatting>> {
    let differential_formats =
        sheet
            .owner
            .styles_uri
            .as_ref()
            .map(|uri| {
                sheet.owner.package.get_part(uri).map(|part| {
                    crate::conditional_formatting::parse_differential_formats(part.blob())
                })
            })
            .transpose()?
            .transpose()?
            .unwrap_or_default();
    crate::conditional_formatting::parse_conditional_formattings(
        xml(sheet)?,
        differential_formats.len(),
    )
}

/// Parse worksheet data-validation collections.
pub(crate) fn data_validations(
    sheet: &Worksheet,
) -> Result<Vec<crate::data_validation::Collection>> {
    crate::data_validation::parse_data_validation_collections(xml(sheet)?)
}

/// Parse the worksheet's optional data-consolidation declaration.
pub(crate) fn data_consolidation(
    sheet: &Worksheet,
) -> Result<Option<crate::data_consolidation::DataConsolidation>> {
    crate::data_consolidation::parse_worksheet_data_consolidation(xml(sheet)?)
}

/// Parse core worksheet header/footer settings.
pub(crate) fn header_footer(sheet: &Worksheet) -> Result<Option<crate::header_footer::Settings>> {
    crate::header_footer::parse_worksheet_header_footer(xml(sheet)?)
}

/// Parse ignored-error declarations.
pub(crate) fn ignored_errors(
    sheet: &Worksheet,
) -> Result<Option<crate::ignored_errors::IgnoredErrors>> {
    crate::ignored_errors::parse_worksheet_ignored_errors(xml(sheet)?)
}

/// Parse inert worksheet smart-tag annotations.
pub(crate) fn smart_tags(sheet: &Worksheet) -> Result<Option<crate::smart_tags::Collection>> {
    crate::smart_tags::parse(xml(sheet)?)
}

/// Parse the worksheet's named-sheet-view relationship, when present.
pub(crate) fn named_sheet_views(
    sheet: &Worksheet,
) -> Result<Option<crate::named_sheet_view::Views>> {
    crate::named_sheet_view::load_worksheet_named_sheet_views(
        &sheet.owner.package,
        &sheet.data.part_uri,
    )
}

/// Parse worksheet outline properties.
pub(crate) fn outline_properties(
    sheet: &Worksheet,
) -> Result<Option<crate::outline_properties::OutlineProperties>> {
    crate::outline_properties::parse_outline_properties(xml(sheet)?)
}

/// Parse worksheet page margins.
pub(crate) fn page_margins(sheet: &Worksheet) -> Result<Option<crate::page_margins::Margins>> {
    crate::page_margins::parse_page_margins(xml(sheet)?)
}

/// Parse worksheet row and column page breaks.
pub(crate) fn page_breaks(sheet: &Worksheet) -> Result<crate::page_breaks::PageBreaks> {
    let value = sheet
        .data
        .page_breaks
        .get_or_try_init(|| crate::page_breaks::parse(xml(sheet)?))?;
    Ok(value.clone())
}

/// Parse worksheet page setup.
pub(crate) fn page_setup(sheet: &Worksheet) -> Result<Option<crate::page_setup::Setup>> {
    crate::page_setup::parse_worksheet_page_setup(xml(sheet)?)
}

/// Parse worksheet phonetic properties.
pub(crate) fn phonetic_properties(
    sheet: &Worksheet,
) -> Result<Option<crate::phonetic_properties::PhoneticProperties>> {
    crate::phonetic_properties::parse_phonetic_properties(xml(sheet)?)
}

/// Parse worksheet print options.
pub(crate) fn print_options(
    sheet: &Worksheet,
) -> Result<Option<crate::print_options::PrintOptions>> {
    crate::print_options::parse_print_options(xml(sheet)?)
}

/// Parse worksheet what-if scenarios.
pub(crate) fn scenarios(sheet: &Worksheet) -> Result<Option<crate::scenarios::Collection>> {
    crate::scenarios::parse_worksheet_scenarios(xml(sheet)?)
}

/// Parse worksheet-level calculation properties.
pub(crate) fn calculation_properties(
    sheet: &Worksheet,
) -> Result<Option<crate::sheet_calculation_properties::Properties>> {
    crate::sheet_calculation_properties::parse(xml(sheet)?)
}

/// Parse the complete worksheet protection projection.
pub(crate) fn protection(sheet: &Worksheet) -> Result<crate::sheet_protection::Metadata> {
    crate::sheet_protection::parse_protection(xml(sheet)?)
}

/// Parse the worksheet's ordinary sheet-view collection.
pub(crate) fn views(sheet: &Worksheet) -> Result<Option<crate::sheet_view::Collection>> {
    crate::sheet_view::parse_worksheet_views(xml(sheet)?)
}

/// Parse every table relationship owned by the worksheet.
pub(crate) fn tables(sheet: &Worksheet) -> Result<Vec<TablePart>> {
    let mut values = Vec::new();
    for relationship in relationships(sheet)?
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::TABLE | rt::STRICT_TABLE))
    {
        if relationship.is_external() {
            return Err(invalid("worksheet table relationship cannot be external"));
        }
        let part_name = relationship.target_partname()?;
        let table_part = sheet.owner.package.get_part(&part_name)?;
        if table_part.content_type() != ct::SML_TABLE {
            return Err(invalid(format!(
                "worksheet table part '{part_name}' has content type '{}', expected '{}'",
                table_part.content_type(),
                ct::SML_TABLE
            )));
        }
        let table = crate::table::parse_table_xml(table_part.blob())?
            .ok_or_else(|| invalid("worksheet table part is missing its table root"))?;
        values.push(TablePart {
            relationship_id: relationship.r_id().into(),
            part_name: part_name.to_string().into_boxed_str(),
            table,
        });
    }
    values.sort_unstable_by(|left, right| {
        left.relationship_id
            .cmp(&right.relationship_id)
            .then_with(|| left.part_name.cmp(&right.part_name))
    });
    Ok(values)
}

/// Load and validate every query-table part owned by the worksheet.
pub(crate) fn query_tables(sheet: &Worksheet) -> Result<Vec<crate::query_table::Part>> {
    crate::query_table::load_worksheet_query_tables(&sheet.owner.package, &sheet.data.part_uri)
}

/// Load the worksheet's inert `ActiveX` graph.
pub(crate) fn active_x(sheet: &Worksheet) -> Result<crate::active_x::ControlSet> {
    crate::active_x::load_from_worksheet(&sheet.owner.package, &sheet.data.part_uri)
}

/// Load timeline views associated with this worksheet.
pub(crate) fn timelines(sheet: &Worksheet) -> Result<Vec<crate::timeline::Part>> {
    Ok(
        crate::timeline::load_parts(&sheet.owner.package, &sheet.owner.workbook_uri)?
            .into_iter()
            .filter(|part| part.worksheet_part_name == sheet.data.part_uri.as_str())
            .collect(),
    )
}

/// Extract array-formula anchors from the already parsed sparse worksheet store.
pub(crate) fn array_formulas(sheet: &Worksheet) -> Result<Vec<ArrayFormula>> {
    let store = sheet.store()?;
    let mut values = Vec::new();
    for entry in store.entries() {
        let crate::cell::Cell::Formula(formula) = &entry.cell else {
            continue;
        };
        let crate::formula::Kind::Array { range } = formula.kind() else {
            continue;
        };
        values.push(ArrayFormula {
            row: entry.address.row().get(),
            column: entry.address.column().get(),
            formula: formula.text().into(),
            range: range.as_ref().map(|value| value.as_str().into()),
        });
    }
    Ok(values)
}

#[cfg(test)]
mod tests {
    use crate::Workbook;

    #[test]
    fn empty_snapshot_exposes_typed_worksheet_facades() {
        let workbook = Workbook::new().expect("minimal workbook");
        let sheet = workbook.active_sheet().expect("active worksheet");

        assert!(
            workbook
                .workbook_protection_metadata()
                .expect("workbook-protection read")
                .is_none()
        );
        assert!(sheet.auto_filter().expect("auto-filter read").is_none());
        assert!(
            sheet
                .conditional_formattings()
                .expect("conditional-formatting read")
                .is_empty()
        );
        assert!(
            sheet
                .data_validations()
                .expect("data-validation read")
                .is_empty()
        );
        assert!(sheet.header_footer().expect("header/footer read").is_none());
        assert!(sheet.views().expect("sheet-view read").is_none());
        assert!(sheet.tables().expect("table read").is_empty());
        assert!(sheet.query_tables().expect("query-table read").is_empty());
        assert!(
            sheet
                .array_formulas()
                .expect("array-formula read")
                .is_empty()
        );
    }
}
