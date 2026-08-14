//! Narrow dependency-free cross-workbook cell transfer planning.
//!
//! The ordinary same-workbook transfer intentionally remains in the semantic
//! transaction owner. This module owns the stricter provenance boundary needed
//! when values cross independent workbook lineages: only scalar cell values
//! are staged, and no donor package identity is allowed to enter the target
//! edit.

use litchi_ooxml_common::relationships::{STRICT_NAMESPACE, TRANSITIONAL_NAMESPACE};
use litchi_opc::{Part, constants::relationship_type as rt};
use litchi_sheet::{Area, At, Cell as Address, Rect};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::semantic::Edit;
use super::validation::pending_merge;
use crate::cell::{Cell, Content, Store, Stored};
use crate::error::{Error, Result, allocation, invalid};
use crate::formula::Kind as FormulaKind;
use crate::raw::namespace::{
    SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE, is_spreadsheetml_name,
};
use crate::raw::worksheet::edit::Action;
use crate::workbook::{Flavor, Selector, Workbook, Worksheet, WorksheetKind};

const MAX_SCALAR_TRANSFER: u64 = 65_536;
const MAX_SCALAR_TRANSFER_BYTES: usize = 16 * 1024 * 1024;
const SCALAR_TRANSFER_ENTRY_OVERHEAD: usize = 64;
const MAX_FORMULA_DEPENDENCY_SCAN: usize = 1_048_576;
const MAX_STRICT_WORKSHEET_BYTES: usize = 32 * 1024 * 1024;
const MAX_STRICT_WORKSHEET_EVENTS: usize = 1_000_000;
const MAX_STRICT_WORKSHEET_DEPTH: usize = 64;
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const STRICT_THEME_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";

/// Prepare and stage one dependency-free cross-workbook scalar copy.
///
/// All source and target validation happens before the target edit map is
/// touched. The returned actions carry owned scalar values only; style and
/// shared-string identities are deliberately not representable in this plan.
pub(super) fn copy_scalar_cells_from<'edit, 'source, 'range, 'target, 'anchor>(
    edit: &'edit mut Edit,
    donor: &Workbook,
    source: impl Into<Selector<'source>>,
    range: impl Into<Area<'range>>,
    target: impl Into<Selector<'target>>,
    anchor: impl Into<At<'anchor>>,
) -> Result<Option<&'edit mut Edit>> {
    let Some(source_sheet) = donor.sheet(source)? else {
        return Ok(None);
    };
    let Some(target_sheet) = edit.base.sheet(target)? else {
        return Ok(None);
    };
    ensure_worksheet(&source_sheet)?;
    ensure_worksheet(&target_sheet)?;
    super::codec::ensure_unsigned(donor)?;
    ensure_supported_flavor(donor)?;
    ensure_donor_catalog(donor)?;
    ensure_supported_flavor(&edit.base)?;
    ensure_workbook_relationships(&edit.base)?;

    let source_range = range.into().resolve()?;
    let target_start = anchor.into().resolve()?;
    let target_end_row = target_start
        .row()
        .get()
        .checked_add(source_range.rows())
        .ok_or_else(|| invalid("cross-workbook scalar target row overflow"))?;
    let target_end_column = target_start
        .column()
        .get()
        .checked_add(source_range.columns())
        .ok_or_else(|| invalid("cross-workbook scalar target column overflow"))?;
    let target_range = Rect::new(target_start, target_end_row, target_end_column)?;
    let cells = u64::from(source_range.rows())
        .checked_mul(u64::from(source_range.columns()))
        .ok_or_else(|| invalid("cross-workbook scalar area overflow"))?;
    if cells > MAX_SCALAR_TRANSFER {
        return Err(invalid(format!(
            "cross-workbook scalar transfer contains {cells} cells; limit is {MAX_SCALAR_TRANSFER}"
        )));
    }

    // A source or target relationship is never copied or retargeted by this
    // operation. Refusing the whole worksheet surface keeps comments,
    // hyperlinks, drawings, tables, and vendor relationship extensions out of
    // the dependency closure instead of guessing which ones are unrelated.
    ensure_worksheet_surface(donor, &source_sheet)?;
    ensure_worksheet_surface(&edit.base, &target_sheet)?;
    ensure_plain_unprotected(donor, &source_sheet)?;
    ensure_plain_unprotected(&edit.base, &target_sheet)?;
    if edit.drawings.contains_key(&target_sheet.position()) {
        return Err(Error::Unsupported {
            feature: "composing cross-workbook scalar transfer with a staged drawing graph",
        });
    }

    let source_store = source_sheet.store()?;
    ensure_formula_ownership_is_disjoint(source_store, source_range)?;
    ensure_formula_ownership_is_disjoint(target_sheet.store()?, target_range)?;
    ensure_pending_formula_ownership_is_disjoint(edit, &target_sheet, target_range)?;
    ensure_source_range_is_dependency_free(source_store, source_range)?;
    ensure_target_range_is_empty(edit, &target_sheet, target_range)?;

    let capacity = usize::try_from(cells)
        .map_err(|error| invalid(format!("scalar transfer count exceeds usize: {error}")))?;
    let mut staged = Vec::new();
    staged
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("cross-workbook scalar transfer plan", source))?;
    let mut staged_bytes = 0usize;

    for (source_address, _) in source_store.cells(source_range) {
        let entry = source_store
            .entry(source_address)
            .ok_or_else(|| invalid("source cell disappeared during scalar transfer planning"))?;
        let Some((content, value_bytes)) = scalar_content(entry)? else {
            continue;
        };
        let entry_bytes = value_bytes
            .checked_add(SCALAR_TRANSFER_ENTRY_OVERHEAD)
            .ok_or_else(|| invalid("cross-workbook scalar transfer plan size overflow"))?;
        staged_bytes = staged_bytes
            .checked_add(entry_bytes)
            .ok_or_else(|| invalid("cross-workbook scalar transfer plan size overflow"))?;
        if staged_bytes > MAX_SCALAR_TRANSFER_BYTES {
            return Err(invalid(format!(
                "cross-workbook scalar transfer plan exceeds {MAX_SCALAR_TRANSFER_BYTES} bytes"
            )));
        }
        let row_offset = source_address
            .row()
            .get()
            .checked_sub(source_range.start().row().get())
            .ok_or_else(|| invalid("source row precedes scalar transfer range"))?;
        let column_offset = source_address
            .column()
            .get()
            .checked_sub(source_range.start().column().get())
            .ok_or_else(|| invalid("source column precedes scalar transfer range"))?;
        let target_row = target_start
            .row()
            .get()
            .checked_add(row_offset)
            .ok_or_else(|| invalid("scalar transfer target row overflow"))?;
        let target_column = target_start
            .column()
            .get()
            .checked_add(column_offset)
            .ok_or_else(|| invalid("scalar transfer target column overflow"))?;
        let target_address = Address::at(target_row, target_column)?;
        staged.push((target_address, content));
    }

    if staged.is_empty() {
        return Ok(Some(edit));
    }

    // `SheetActions.cells` is a shared-MSRV BTreeMap without a stable
    // fallible reserve operation. The lexical staging vector and payload
    // budget above make every new allocation bounded and fallible; this
    // cardinality check bounds the final map insertion frontier before the
    // map is touched. All validation remains complete before this point, so
    // a failure cannot expose a partially staged transfer.
    let existing_cell_count = edit
        .sheets
        .get(&target_sheet.position())
        .map_or(0, |actions| actions.cells.len());
    let final_cell_count = existing_cell_count
        .checked_add(staged.len())
        .ok_or_else(|| invalid("cross-workbook scalar target action count overflow"))?;
    if final_cell_count > MAX_SCALAR_TRANSFER as usize {
        return Err(invalid(format!(
            "cross-workbook scalar target action count exceeds {MAX_SCALAR_TRANSFER}"
        )));
    }
    let target_actions = edit.sheets.entry(target_sheet.position()).or_default();
    for (address, content) in staged {
        target_actions.cells.insert(address, Action::set(content));
    }
    edit.cross_workbook_scalar = true;
    Ok(Some(edit))
}

fn ensure_supported_flavor(workbook: &Workbook) -> Result<()> {
    if matches!(
        workbook.flavor(),
        Flavor::MacroWorkbook | Flavor::MacroTemplate
    ) {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer for macro workbook flavors",
        });
    }
    Ok(())
}

fn ensure_plain_unprotected(workbook: &Workbook, sheet: &Worksheet) -> Result<()> {
    // Encrypted provenance is retained by the workbook only when the optional
    // encryption feature is enabled. This operation intentionally refuses it
    // instead of silently declassifying a donor into an ordinary target plan.
    #[cfg(feature = "encryption")]
    if workbook.encryption().is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with encrypted workbooks",
        });
    }
    if workbook.workbook_protection_metadata()?.is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with protected workbooks",
        });
    }
    let protection = sheet.protection()?;
    if protection.sheet_protection().is_some()
        || !protection.protected_range_collections().is_empty()
    {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with protected workbooks",
        });
    }
    Ok(())
}

fn ensure_worksheet(sheet: &Worksheet) -> Result<()> {
    if sheet.kind() != WorksheetKind::Worksheet {
        return Err(Error::NotWorksheet {
            sheet: sheet.name().to_owned(),
        });
    }
    Ok(())
}

fn ensure_donor_catalog(donor: &Workbook) -> Result<()> {
    if !donor.defined_names().is_empty() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with defined-name dependencies",
        });
    }
    if !donor.external_reference_ids().is_empty() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with external links",
        });
    }
    if !donor.pivot_caches().is_empty() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with pivot dependencies",
        });
    }
    ensure_workbook_relationships(donor)
}

fn ensure_workbook_relationships(workbook: &Workbook) -> Result<()> {
    let part = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    if part.rels().iter().any(|relationship| {
        relationship.is_external() || !supported_workbook_relationship(relationship.reltype())
    }) {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with unsupported workbook relationships",
        });
    }
    Ok(())
}

fn supported_workbook_relationship(reltype: &str) -> bool {
    matches!(
        reltype,
        rt::WORKSHEET
            | rt::STRICT_WORKSHEET
            | rt::STYLES
            | rt::STRICT_STYLES
            | rt::SHARED_STRINGS
            | rt::STRICT_SHARED_STRINGS
            | rt::THEME
            | STRICT_THEME_RELATIONSHIP
    )
}

fn ensure_worksheet_surface(workbook: &Workbook, sheet: &Worksheet) -> Result<()> {
    let part = workbook.inner.package.get_part(sheet.part_uri())?;
    // The typed worksheet model intentionally has extension fallbacks. A
    // cross-lineage transfer cannot use those fallbacks: an unmodeled direct
    // child, cell attribute, or MCE branch could carry a dependency that the
    // modeled store has already flattened. Scan the original lexical surface
    // before consulting any model accessor and fail closed on every element
    // outside the dependency-free grid vocabulary below.
    ensure_strict_worksheet_surface(part.blob())?;
    if !sheet.tables()?.is_empty() || !sheet.query_tables()?.is_empty() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with table dependencies",
        });
    }
    if part.rels().iter().next().is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with worksheet relationships",
        });
    }
    if !sheet.data_validations()?.is_empty() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with data validation",
        });
    }
    if !sheet.conditional_formattings()?.is_empty() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with conditional formatting",
        });
    }
    if sheet.auto_filter()?.is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with worksheet filters",
        });
    }
    if sheet.data_consolidation()?.is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with data consolidation",
        });
    }
    if sheet.smart_tags()?.is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with smart-tag metadata",
        });
    }
    if sheet.scenarios()?.is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with scenario metadata",
        });
    }
    if sheet.named_sheet_views()?.is_some() {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer with named-sheet-view metadata",
        });
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StrictWorksheetElement {
    Worksheet,
    Dimension,
    SheetData,
    Row,
    Cell,
    Formula,
    Value,
    Inline,
    Text,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StrictFormulaType {
    Normal,
    Array,
    DataTable,
    Shared,
}

impl StrictWorksheetElement {
    const fn local_name(self) -> &'static [u8] {
        match self {
            Self::Worksheet => b"worksheet",
            Self::Dimension => b"dimension",
            Self::SheetData => b"sheetData",
            Self::Row => b"row",
            Self::Cell => b"c",
            Self::Formula => b"f",
            Self::Value => b"v",
            Self::Inline => b"is",
            Self::Text => b"t",
        }
    }

    const fn allows_text(self) -> bool {
        matches!(self, Self::Formula | Self::Value | Self::Text)
    }
}

/// Validate the complete original worksheet XML surface accepted by the
/// narrow scalar-transfer contract.
///
/// The ordinary worksheet reader intentionally preserves some unknown
/// structures through modeled fallback records. That is useful for ordinary
/// round trips, but unsafe when values cross workbook lineages: a fallback
/// may hide a direct drawing, extension, MCE branch, or cell attribute. This
/// scanner therefore accepts only the worksheet/grid vocabulary whose
/// dependency closure is checked by this module. It is deliberately stricter
/// than the normal worksheet parser and rejects the whole operation when any
/// unsupported surface exists anywhere in either worksheet part.
fn ensure_strict_worksheet_surface(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_STRICT_WORKSHEET_BYTES {
        return Err(invalid(format!(
            "worksheet XML exceeds {MAX_STRICT_WORKSHEET_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML event count overflow"))?;
        if events > MAX_STRICT_WORKSHEET_EVENTS {
            return Err(invalid(format!(
                "worksheet XML exceeds {MAX_STRICT_WORKSHEET_EVENTS} events"
            )));
        }
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid worksheet XML: {error}")))?;
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if stack.is_empty() && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
                }
                if stack.len() >= MAX_STRICT_WORKSHEET_DEPTH {
                    return Err(invalid(format!(
                        "worksheet XML nesting exceeds {MAX_STRICT_WORKSHEET_DEPTH}"
                    )));
                }
                let resolver = reader.resolver().clone();
                let (namespace, _) = resolver.resolve_element(element.name());
                let kind = strict_begin_element(
                    &resolver,
                    reader.decoder(),
                    &element,
                    namespace,
                    stack.last().copied(),
                )?;
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("strict worksheet element stack", source))?;
                stack.push(kind);
                root_seen |= stack.len() == 1;
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if stack.is_empty() && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
                }
                let resolver = reader.resolver().clone();
                let (namespace, _) = resolver.resolve_element(element.name());
                let _kind = strict_begin_element(
                    &resolver,
                    reader.decoder(),
                    &element,
                    namespace,
                    stack.last().copied(),
                )?;
                if stack.is_empty() {
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(element) => {
                let kind = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected worksheet end element"))?;
                let resolver = reader.resolver().clone();
                let (namespace, _) = resolver.resolve_element(element.name());
                if !is_spreadsheetml_name(&namespace, element.name(), kind.local_name()) {
                    return Err(invalid("mismatched worksheet end element"));
                }
                if kind == StrictWorksheetElement::Worksheet {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                if stack.last().is_some_and(|kind| kind.allows_text()) {
                    // Formula, value, and inline text lexical forms are
                    // validated by the normal worksheet materializer.
                } else if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(unsupported_worksheet_surface());
                }
            },
            Event::Decl(_) if !root_seen && !declaration_seen => {
                declaration_seen = true;
            },
            Event::CData(_) | Event::Comment(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(unsupported_worksheet_surface());
            },
            Event::GeneralRef(_) => return Err(unsupported_worksheet_surface()),
            Event::Decl(_) => return Err(invalid("worksheet XML has an invalid declaration")),
            Event::Eof => break,
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("unterminated worksheet XML"));
    }
    Ok(())
}

fn strict_begin_element(
    resolver: &NamespaceResolver,
    decoder: Decoder,
    element: &BytesStart<'_>,
    namespace: ResolveResult<'_>,
    parent: Option<StrictWorksheetElement>,
) -> Result<StrictWorksheetElement> {
    if !matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if value == SPREADSHEETML_NAMESPACE || value == STRICT_SPREADSHEETML_NAMESPACE
    ) {
        return Err(unsupported_worksheet_surface());
    }
    let local = element.local_name();
    let kind = match parent {
        None if local.as_ref() == b"worksheet" => StrictWorksheetElement::Worksheet,
        Some(StrictWorksheetElement::Worksheet) => match local.as_ref() {
            b"dimension" => StrictWorksheetElement::Dimension,
            b"sheetData" => StrictWorksheetElement::SheetData,
            _ => return Err(unsupported_worksheet_surface()),
        },
        Some(StrictWorksheetElement::SheetData) if local.as_ref() == b"row" => {
            StrictWorksheetElement::Row
        },
        Some(StrictWorksheetElement::Row) if local.as_ref() == b"c" => StrictWorksheetElement::Cell,
        Some(StrictWorksheetElement::Cell) => match local.as_ref() {
            b"f" => StrictWorksheetElement::Formula,
            b"v" => StrictWorksheetElement::Value,
            b"is" => StrictWorksheetElement::Inline,
            _ => return Err(unsupported_worksheet_surface()),
        },
        Some(StrictWorksheetElement::Inline) if local.as_ref() == b"t" => {
            StrictWorksheetElement::Text
        },
        Some(StrictWorksheetElement::Dimension)
        | Some(StrictWorksheetElement::Formula)
        | Some(StrictWorksheetElement::Value)
        | Some(StrictWorksheetElement::Text)
        | Some(StrictWorksheetElement::Inline)
        | Some(StrictWorksheetElement::Row)
        | Some(StrictWorksheetElement::SheetData)
        | None => return Err(unsupported_worksheet_surface()),
    };
    validate_strict_attributes(resolver, decoder, element, kind).map(|()| kind)
}

fn validate_strict_attributes(
    resolver: &NamespaceResolver,
    decoder: Decoder,
    element: &BytesStart<'_>,
    kind: StrictWorksheetElement,
) -> Result<()> {
    let mut formula_type = StrictFormulaType::Normal;
    let mut formula_ref_seen = false;
    let mut formula_index_seen = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid worksheet attribute: {error}")))?;
        if is_namespace_declaration(attribute.key) {
            if kind != StrictWorksheetElement::Worksheet {
                return Err(unsupported_worksheet_surface());
            }
            validate_strict_namespace_declaration(decoder, &attribute)?;
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if kind == StrictWorksheetElement::Text
            && matches!(
                namespace,
                ResolveResult::Bound(Namespace(value))
                    if value == XML_NAMESPACE
            )
            && local.as_ref() == b"space"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| invalid(format!("invalid xml:space value: {error}")))?;
            if !matches!(value.as_ref(), "default" | "preserve") {
                return Err(invalid("worksheet xml:space must be default or preserve"));
            }
            continue;
        }
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(unsupported_worksheet_surface());
        }
        let allowed = match kind {
            StrictWorksheetElement::Worksheet => false,
            StrictWorksheetElement::Dimension => local.as_ref() == b"ref",
            StrictWorksheetElement::SheetData => false,
            StrictWorksheetElement::Row => local.as_ref() == b"r",
            StrictWorksheetElement::Cell => {
                matches!(local.as_ref(), b"r" | b"t" | b"s" | b"cm" | b"vm")
            },
            StrictWorksheetElement::Formula => {
                matches!(local.as_ref(), b"t" | b"ref" | b"si" | b"bx")
            },
            StrictWorksheetElement::Value | StrictWorksheetElement::Inline => false,
            StrictWorksheetElement::Text => false,
        };
        if !allowed {
            return Err(unsupported_worksheet_surface());
        }
        if matches!(kind, StrictWorksheetElement::Cell) && local.as_ref() == b"t" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| invalid(format!("invalid worksheet cell type: {error}")))?;
            if !matches!(
                value.as_ref(),
                "b" | "d" | "e" | "inlineStr" | "n" | "s" | "str"
            ) {
                return Err(unsupported_worksheet_surface());
            }
        }
        if matches!(kind, StrictWorksheetElement::Formula) && local.as_ref() == b"t" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| invalid(format!("invalid worksheet formula type: {error}")))?;
            formula_type = match value.as_ref() {
                "normal" => StrictFormulaType::Normal,
                "array" => StrictFormulaType::Array,
                "dataTable" => StrictFormulaType::DataTable,
                "shared" => StrictFormulaType::Shared,
                _ => return Err(unsupported_worksheet_surface()),
            };
        }
        if matches!(kind, StrictWorksheetElement::Formula) {
            if local.as_ref() == b"ref" {
                formula_ref_seen = true;
            } else if local.as_ref() == b"si" {
                formula_index_seen = true;
            }
        }
    }
    if kind == StrictWorksheetElement::Formula {
        match formula_type {
            StrictFormulaType::Normal if formula_ref_seen || formula_index_seen => {
                return Err(unsupported_worksheet_surface());
            },
            StrictFormulaType::Array | StrictFormulaType::DataTable if formula_index_seen => {
                return Err(unsupported_worksheet_surface());
            },
            StrictFormulaType::Shared if !formula_index_seen => {
                return Err(unsupported_worksheet_surface());
            },
            StrictFormulaType::Normal
            | StrictFormulaType::Array
            | StrictFormulaType::DataTable
            | StrictFormulaType::Shared => {},
        }
    }
    Ok(())
}

fn validate_strict_namespace_declaration(
    decoder: Decoder,
    attribute: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<()> {
    let value = attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
        .map_err(|error| invalid(format!("invalid worksheet namespace declaration: {error}")))?;
    let value = value.as_ref().as_bytes();
    let allowed = value == SPREADSHEETML_NAMESPACE
        || value == STRICT_SPREADSHEETML_NAMESPACE
        || value == TRANSITIONAL_NAMESPACE
        || value == STRICT_NAMESPACE;
    if !allowed {
        return Err(unsupported_worksheet_surface());
    }
    Ok(())
}

fn is_namespace_declaration(name: QName<'_>) -> bool {
    (name.prefix().is_none() && name.local_name().as_ref() == b"xmlns")
        || name
            .prefix()
            .as_ref()
            .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
}

const fn unsupported_worksheet_surface() -> Error {
    Error::Unsupported {
        feature: "cross-workbook scalar transfer with unsupported worksheet XML surface",
    }
}

fn ensure_source_range_is_dependency_free(store: &Store, range: Rect) -> Result<()> {
    if store
        .merge_ranges()
        .iter()
        .copied()
        .any(|merge| rectangles_overlap(merge, range))
    {
        return Err(Error::Unsupported {
            feature: "cross-workbook scalar transfer across merged cells",
        });
    }

    for row in store.row_entries() {
        let index = row.index.get();
        if index >= range.start().row().get()
            && index < range.end().0
            && row.properties.style.is_some()
        {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer with row styles",
            });
        }
    }
    for column in store.column_entries() {
        let first = column.first.get();
        let last = column.last.get();
        if first < range.end().1
            && range.start().column().get() <= last
            && column.properties.style.is_some()
        {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer with column styles",
            });
        }
    }

    for (address, _) in store.cells(range) {
        let entry = store
            .entry(address)
            .ok_or_else(|| invalid("source cell disappeared during dependency validation"))?;
        if entry.style.is_some() {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer with cell styles",
            });
        }
        if entry.shared_string.is_some() {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer with shared-string dependencies",
            });
        }
        if entry.inline_rich {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer with rich inline strings",
            });
        }
        if entry.cell_metadata.is_some() || entry.value_metadata.is_some() {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer with cell metadata",
            });
        }
        if matches!(entry.cell, Cell::Formula(_) | Cell::Unknown(_)) {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer with formula or unknown cells",
            });
        }
    }
    Ok(())
}

/// Refuse a transfer that intersects a formula-owned range whose anchor may
/// live outside the selected rectangle. Array/data-table followers and shared
/// formula members are materialized as ordinary-looking cells, so checking
/// only the entries physically present in `range` is not sufficient.
fn ensure_formula_ownership_is_disjoint(store: &Store, range: Rect) -> Result<()> {
    if store.entries().len() > MAX_FORMULA_DEPENDENCY_SCAN {
        return Err(invalid(format!(
            "formula dependency scan exceeds {MAX_FORMULA_DEPENDENCY_SCAN} cells"
        )));
    }
    for entry in store.entries() {
        let Some(owned) = formula_owned_range(entry)? else {
            continue;
        };
        if rectangles_overlap(owned, range) {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer across range-owned formulas",
            });
        }
    }
    Ok(())
}

fn ensure_pending_formula_ownership_is_disjoint(
    edit: &Edit,
    sheet: &Worksheet,
    range: Rect,
) -> Result<()> {
    let Some(actions) = edit.sheets.get(&sheet.position()) else {
        return Ok(());
    };
    if actions.cells.len() > MAX_FORMULA_DEPENDENCY_SCAN {
        return Err(invalid(format!(
            "pending formula dependency scan exceeds {MAX_FORMULA_DEPENDENCY_SCAN} cells"
        )));
    }
    for (address, action) in &actions.cells {
        let Action::Update {
            payload: Some(crate::raw::worksheet::edit::Payload::Set(Content::Formula(formula))),
            ..
        } = action
        else {
            continue;
        };
        let owned = pending_formula_range(*address, formula)?;
        if rectangles_overlap(owned, range) {
            return Err(Error::Unsupported {
                feature: "cross-workbook scalar transfer across pending range-owned formulas",
            });
        }
    }
    Ok(())
}

fn formula_owned_range(entry: &Stored) -> Result<Option<Rect>> {
    let Cell::Formula(formula) = &entry.cell else {
        return Ok(None);
    };
    let owned = match formula.kind() {
        FormulaKind::Scalar => entry
            .formula_range
            .unwrap_or_else(|| Rect::single(entry.address)),
        FormulaKind::Array { range } | FormulaKind::DataTable { range } => range
            .as_deref()
            .map(Rect::from_a1)
            .transpose()
            .map_err(Error::from)?
            .unwrap_or_else(|| Rect::single(entry.address)),
        FormulaKind::Unknown(_) => Rect::single(entry.address),
    };
    Ok(Some(owned))
}

fn pending_formula_range(address: Address, formula: &crate::formula::Formula) -> Result<Rect> {
    let range = match formula.kind() {
        FormulaKind::Array { range } | FormulaKind::DataTable { range } => range.as_deref(),
        FormulaKind::Scalar | FormulaKind::Unknown(_) => None,
    };
    range
        .map(Rect::from_a1)
        .transpose()
        .map(|range| range.unwrap_or_else(|| Rect::single(address)))
        .map_err(Error::from)
}

fn ensure_target_range_is_empty(edit: &Edit, sheet: &Worksheet, range: Rect) -> Result<()> {
    let store = sheet.store()?;
    let intents = edit
        .sheets
        .get(&sheet.position())
        .map_or(&[][..], |actions| actions.merges.as_slice());
    for row in range.start().row().get()..range.end().0 {
        for column in range.start().column().get()..range.end().1 {
            let address = Address::at(row, column)?;
            if pending_merge(store.merge_ranges(), intents, address).is_some() {
                return Err(Error::Unsupported {
                    feature: "cross-workbook scalar transfer into merged cells",
                });
            }
            let stored = store.entry(address);
            let occupied = edit
                .sheets
                .get(&sheet.position())
                .and_then(|actions| actions.cells.get(&address))
                .map_or(stored.is_some(), |action| match action {
                    Action::Remove => false,
                    Action::Update { .. } => stored.is_some() || action.creates_missing(),
                });
            if occupied {
                return Err(Error::Unsupported {
                    feature: "cross-workbook scalar transfer into occupied cells",
                });
            }
        }
    }
    Ok(())
}

fn scalar_content(entry: &Stored) -> Result<Option<(Content, usize)>> {
    let Cell::Value(value) = &entry.cell else {
        return Ok(None);
    };
    value.validate_for_write()?;
    let bytes = match value {
        crate::cell::Value::Bool(_) => 1,
        crate::cell::Value::Number(value) => value.as_str().len(),
        crate::cell::Value::Text(value) => value.as_str().len(),
        crate::cell::Value::Date(value) => value.as_str().len(),
        crate::cell::Value::Error(value) => value.as_str().len(),
    };
    Ok(Some((Content::Value(value.clone()), bytes)))
}

const fn rectangles_overlap(left: Rect, right: Rect) -> bool {
    left.start().row().get() < right.end().0
        && right.start().row().get() < left.end().0
        && left.start().column().get() < right.end().1
        && right.start().column().get() < left.end().1
}
