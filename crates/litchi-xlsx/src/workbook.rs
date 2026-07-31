//! Immutable workbook snapshots and selector-first sheet lookup.

mod edit;

pub use edit::{Change, Commit, Edit, Patch, SheetEdit, State};

use std::convert::Infallible;
use std::io::Read;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::Selector;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use litchi_sheet::{At, Rect};

use crate::cell::{Store, Text};
use crate::error::{Error, Result, invalid};
use crate::raw;
use crate::{Cell, Cells};

const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOGSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const STRICT_DIALOGSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet";
const MACROSHEET_REL: &str = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const INTL_MACROSHEET_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

/// Semantic selector accepted by [`Workbook::sheet`].
///
/// Names and checked zero-based positions are the ordinary entry points. The
/// uninhabited identity variant reserves room for a future lineage-checked
/// durable selector without exposing native SpreadsheetML IDs.
pub type SheetSelector<'a> = litchi_sheet::SheetSelector<'a, Infallible>;

/// Runtime workbook flavor derived from the main-part content type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Flavor {
    Workbook,
    Template,
    MacroWorkbook,
    MacroTemplate,
}

impl Flavor {
    fn from_content_type(value: &str) -> Option<Self> {
        match value {
            ct::SML_SHEET_MAIN => Some(Self::Workbook),
            ct::SML_TEMPLATE_MAIN => Some(Self::Template),
            ct::SML_SHEET_MACRO_MAIN => Some(Self::MacroWorkbook),
            ct::SML_TEMPLATE_MACRO_MAIN => Some(Self::MacroTemplate),
            _ => None,
        }
    }

    /// Whether this flavor can contain a VBA project without promotion.
    pub const fn allows_macros(self) -> bool {
        matches!(self, Self::MacroWorkbook | Self::MacroTemplate)
    }

    /// Whether opening the file is intended to create a new workbook.
    pub const fn is_template(self) -> bool {
        matches!(self, Self::Template | Self::MacroTemplate)
    }
}

/// Workbook date serial system.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DateSystem {
    Excel1900,
    Excel1904,
}

/// Semantic sheet kind resolved from the workbook relationship.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SheetKind {
    Worksheet,
    Chart,
    Dialog,
    Macro,
    Unknown,
}

/// Sheet visibility retained without approximating producer extensions.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Visibility {
    Visible,
    Hidden,
    VeryHidden,
    Unknown(Box<str>),
}

impl From<raw::Visibility> for Visibility {
    fn from(value: raw::Visibility) -> Self {
        match value {
            raw::Visibility::Visible => Self::Visible,
            raw::Visibility::Hidden => Self::Hidden,
            raw::Visibility::VeryHidden => Self::VeryHidden,
            raw::Visibility::Unknown(value) => Self::Unknown(value),
        }
    }
}

#[derive(Debug)]
struct SheetData {
    position: usize,
    name: String,
    kind: SheetKind,
    visibility: Visibility,
    part_uri: PackURI,
    cells: OnceLock<Store>,
    #[allow(dead_code)]
    native_id: u32,
    #[allow(dead_code)]
    relationship_id: String,
}

#[derive(Debug)]
struct Inner {
    package: OpcPackage,
    #[allow(dead_code)]
    workbook_uri: PackURI,
    shared_strings_uri: Option<PackURI>,
    shared_strings: OnceLock<Box<[Text]>>,
    flavor: Flavor,
    date_system: DateSystem,
    active_sheet: Option<usize>,
    sheets: Box<[Arc<SheetData>]>,
    defined_names: Box<[raw::DefinedName]>,
    pivot_caches: Box<[raw::PivotCache]>,
    external_reference_ids: Box<[String]>,
}

/// Immutable, cheap-to-share XLSX workbook snapshot.
#[derive(Debug, Clone)]
pub struct Workbook {
    inner: Arc<Inner>,
}

impl Workbook {
    /// Create a deterministic minimal workbook with one visible worksheet.
    pub fn new() -> Result<Self> {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").map_err(invalid)?;
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").map_err(invalid)?;

        let mut workbook = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.to_string(),
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>"#,
                r#"</workbook>"#
            )
            .as_bytes()
            .to_vec(),
        );
        workbook.rels_mut().try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )?;
        package.try_add_part(Box::new(workbook))?;
        package.try_add_part(Box::new(BlobPart::new(
            worksheet_uri,
            ct::SML_WORKSHEET.to_string(),
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
                r#"<sheetData/></worksheet>"#
            )
            .as_bytes()
            .to_vec(),
        )))?;
        package.rels_mut().try_add_relationship(
            rt::OFFICE_DOCUMENT.to_owned(),
            "xl/workbook.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )?;
        Self::from_package(package)
    }

    /// Open an XLSX-family package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(OpcPackage::open(path)?)
    }

    /// Move bytes into the XLSX parser.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OpcPackage::from_vec(bytes)?;
        Self::from_package(package)
    }

    /// Open a borrowed XLSX byte slice.
    pub fn from_slice(bytes: &[u8]) -> Result<Self> {
        Self::from_package(OpcPackage::from_bytes(bytes)?)
    }

    /// Read an XLSX package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_package(OpcPackage::from_reader(reader)?)
    }

    /// Build a snapshot from a validated OPC package without exposing it in
    /// ordinary sheet APIs.
    pub fn from_package(package: OpcPackage) -> Result<Self> {
        let (workbook_uri, flavor, catalog, sheet_parts, shared_strings_uri) = {
            let workbook = package.main_document_part()?;
            let flavor = Flavor::from_content_type(workbook.content_type()).ok_or_else(|| {
                invalid(format!(
                    "main part '{}' has non-XLSX content type '{}'",
                    workbook.partname(),
                    workbook.content_type()
                ))
            })?;
            let catalog = raw::parse_catalog(workbook.blob())?;
            let sheet_parts = validate_sheet_graph(&package, workbook, &catalog.sheets)?;
            let shared_strings_uri = validate_shared_strings(&package, workbook)?;
            (
                workbook.partname().clone(),
                flavor,
                catalog,
                sheet_parts,
                shared_strings_uri,
            )
        };

        let active_sheet = if catalog.sheets.is_empty() {
            None
        } else {
            Some(catalog.active_sheet_index)
        };
        let sheets = catalog
            .sheets
            .into_iter()
            .zip(sheet_parts)
            .enumerate()
            .map(|(position, (sheet, part))| {
                Arc::new(SheetData {
                    position,
                    name: sheet.name,
                    kind: part.kind,
                    visibility: sheet.visibility.into(),
                    part_uri: part.uri,
                    cells: OnceLock::new(),
                    native_id: sheet.sheet_id,
                    relationship_id: sheet.relationship_id,
                })
            })
            .collect::<Vec<_>>()
            .into_boxed_slice();

        Ok(Self {
            inner: Arc::new(Inner {
                package,
                workbook_uri,
                shared_strings_uri,
                shared_strings: OnceLock::new(),
                flavor,
                date_system: if catalog.uses_1904_date_system {
                    DateSystem::Excel1904
                } else {
                    DateSystem::Excel1900
                },
                active_sheet,
                sheets,
                defined_names: catalog.defined_names.into_boxed_slice(),
                pivot_caches: catalog.pivot_caches.into_boxed_slice(),
                external_reference_ids: catalog.external_reference_ids.into_boxed_slice(),
            }),
        })
    }

    /// Workbook flavor derived from package content, never its filename.
    pub fn flavor(&self) -> Flavor {
        self.inner.flavor
    }

    /// Date serial system used by the workbook.
    pub fn date_system(&self) -> DateSystem {
        self.inner.date_system
    }

    /// Number of logical workbook sheets, including chart and dialog sheets.
    pub fn len(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Whether the workbook catalog contains no sheets.
    pub fn is_empty(&self) -> bool {
        self.inner.sheets.is_empty()
    }

    /// Iterate lightweight sheet handles in workbook order.
    pub fn sheets(&self) -> impl ExactSizeIterator<Item = Sheet> + DoubleEndedIterator + '_ {
        self.inner.sheets.iter().cloned().map(|data| Sheet {
            owner: Arc::clone(&self.inner),
            data,
        })
    }

    /// Look up a sheet by developer-facing name or checked zero-based position.
    pub fn sheet<'a>(&self, selector: impl Into<SheetSelector<'a>>) -> Result<Option<Sheet>> {
        let data = match selector.into() {
            Selector::Position(position) => self.inner.sheets.get(position.get()).cloned(),
            Selector::Name(name) => {
                let mut found = None;
                let mut matches = 0usize;
                for sheet in &self.inner.sheets {
                    if sheet.name.eq_ignore_ascii_case(&name) {
                        matches = matches.saturating_add(1);
                        if found.is_none() {
                            found = Some(Arc::clone(sheet));
                        }
                    }
                }
                if matches > 1 {
                    return Err(Error::AmbiguousSheetName {
                        name: name.into_owned(),
                        matches,
                    });
                }
                found
            },
            Selector::Id(never) => match never {},
            _ => return Err(Error::UnsupportedSelector),
        };
        Ok(data.map(|data| Sheet {
            owner: Arc::clone(&self.inner),
            data,
        }))
    }

    /// Return the active sheet when the workbook contains sheets.
    pub fn active_sheet(&self) -> Option<Sheet> {
        let data = self
            .inner
            .active_sheet
            .and_then(|position| self.inner.sheets.get(position))
            .cloned()?;
        Some(Sheet {
            owner: Arc::clone(&self.inner),
            data,
        })
    }

    /// Low-level inert defined-name records retained by the catalog parser.
    pub fn defined_names(&self) -> &[raw::DefinedName] {
        &self.inner.defined_names
    }

    /// Low-level workbook pivot-cache references.
    pub fn pivot_caches(&self) -> &[raw::PivotCache] {
        &self.inner.pivot_caches
    }

    /// Inert external-workbook relationship IDs, for package diagnostics.
    pub fn external_reference_ids(&self) -> &[String] {
        &self.inner.external_reference_ids
    }

    /// Serialize the immutable package snapshot to bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(PackageWriter::to_bytes(&self.inner.package)?)
    }

    /// Start an isolated semantic transaction over this immutable snapshot.
    pub fn edit(&self) -> Result<Edit> {
        Edit::new(self.clone())
    }

    /// Apply a reversible patch after checking every expected source part.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        patch.apply_to(self)
    }
}

/// Lightweight lifetime-free handle to one sheet in a workbook snapshot.
#[derive(Debug, Clone)]
pub struct Sheet {
    owner: Arc<Inner>,
    data: Arc<SheetData>,
}

impl Sheet {
    /// Developer-facing sheet name.
    pub fn name(&self) -> &str {
        &self.data.name
    }

    /// Checked zero-based workbook position.
    pub fn position(&self) -> usize {
        self.data.position
    }

    /// Semantic sheet kind resolved from its relationship.
    pub fn kind(&self) -> SheetKind {
        self.data.kind
    }

    /// Retained visibility state.
    pub fn visibility(&self) -> &Visibility {
        &self.data.visibility
    }

    /// Whether two handles belong to the same immutable workbook snapshot.
    pub fn same_workbook(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.owner, &other.owner)
    }

    /// Look up a cell by raw zero-based `(row, column)` or a checked [`crate::Address`].
    ///
    /// `None` means no cell record is stored. [`Cell::Empty`] means a record is
    /// present but has no primary payload.
    pub fn cell<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<&Cell>> {
        let address = at.into().resolve()?;
        Ok(self.store()?.get(address))
    }

    /// Lazily traverse only stored cells inside a checked half-open range.
    pub fn cells(&self, range: Rect) -> Result<Cells<'_>> {
        Ok(self.store()?.cells(range))
    }

    /// Bounding rectangle of stored cell records, distinct from declared,
    /// formatted, and content extents.
    pub fn stored_extent(&self) -> Result<Option<Rect>> {
        Ok(self.store()?.extent())
    }

    fn store(&self) -> Result<&Store> {
        if self.data.kind != SheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        if let Some(store) = self.data.cells.get() {
            return Ok(store);
        }

        let part = self.owner.package.get_part(&self.data.part_uri)?;
        let parsed = raw::worksheet::parse(part.blob(), || self.owner.shared_strings())?;
        let _ = self.data.cells.set(parsed);
        self.data
            .cells
            .get()
            .ok_or_else(|| invalid("worksheet cache initialization did not publish a value"))
    }
}

impl Inner {
    fn shared_strings(&self) -> Result<Option<&[Text]>> {
        let Some(uri) = self.shared_strings_uri.as_ref() else {
            return Ok(None);
        };
        if let Some(strings) = self.shared_strings.get() {
            return Ok(Some(strings));
        }

        let part = self.package.get_part(uri)?;
        let parsed = raw::strings::parse(part.blob())?;
        let _ = self.shared_strings.set(parsed);
        self.shared_strings
            .get()
            .map(|strings| Some(strings.as_ref()))
            .ok_or_else(|| invalid("shared-string cache initialization did not publish a value"))
    }
}

struct SheetPart {
    kind: SheetKind,
    uri: PackURI,
}

fn validate_sheet_graph(
    package: &OpcPackage,
    workbook: &dyn Part,
    sheets: &[raw::Sheet],
) -> Result<Vec<SheetPart>> {
    let mut parts = Vec::with_capacity(sheets.len());
    for sheet in sheets {
        let relationship = workbook.rels().get(&sheet.relationship_id).ok_or_else(|| {
            invalid(format!(
                "sheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
        if relationship.is_external() {
            return Err(invalid(format!(
                "sheet '{}' relationship cannot be external",
                sheet.name
            )));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        let kind = match relationship.reltype() {
            rt::WORKSHEET | rt::STRICT_WORKSHEET => {
                require_content_type(sheet, part.content_type(), ct::SML_WORKSHEET)?;
                SheetKind::Worksheet
            },
            CHARTSHEET_REL | STRICT_CHARTSHEET_REL => {
                require_content_type(sheet, part.content_type(), CHARTSHEET_CONTENT_TYPE)?;
                SheetKind::Chart
            },
            DIALOGSHEET_REL | STRICT_DIALOGSHEET_REL => SheetKind::Dialog,
            MACROSHEET_REL | INTL_MACROSHEET_REL => SheetKind::Macro,
            _ => SheetKind::Unknown,
        };
        parts.push(SheetPart { kind, uri: target });
    }
    Ok(parts)
}

fn validate_shared_strings(package: &OpcPackage, workbook: &dyn Part) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::SHARED_STRINGS | rt::STRICT_SHARED_STRINGS
        )
    }) {
        if found.is_some() {
            return Err(invalid("workbook has multiple shared-string relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("shared-string relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.get_part(&uri)?;
        if part.content_type() != ct::SML_SHARED_STRINGS {
            return Err(invalid(format!(
                "shared-string part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_SHARED_STRINGS
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

fn require_content_type(sheet: &raw::Sheet, actual: &str, expected: &str) -> Result<()> {
    if actual != expected {
        return Err(invalid(format!(
            "sheet '{}' has content type '{actual}', expected '{expected}'",
            sheet.name
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::Value;
    use crate::formula::Cache;

    #[test]
    fn new_workbook_is_deterministic_and_selector_first() {
        let first = Workbook::new().expect("valid baseline");
        let second = Workbook::new().expect("valid baseline");

        assert_eq!(first.to_bytes().ok(), second.to_bytes().ok());
        assert_eq!(first.len(), 1);
        assert_eq!(first.flavor(), Flavor::Workbook);
        assert_eq!(first.date_system(), DateSystem::Excel1900);

        let by_name = first.sheet("sheet1").expect("lookup").expect("present");
        let by_position = first.sheet(0usize).expect("lookup").expect("present");
        assert_eq!(by_name.name(), "Sheet1");
        assert_eq!(by_position.position(), 0);
        assert!(by_name.same_workbook(&by_position));
        assert!(matches!(by_name.kind(), SheetKind::Worksheet));
        assert!(matches!(by_name.visibility(), Visibility::Visible));
        assert!(first.sheet(1usize).expect("lookup").is_none());

        let reopened = Workbook::from_bytes(first.to_bytes().expect("serialize"))
            .expect("reopen generated workbook");
        assert_eq!(
            reopened.active_sheet().map(|sheet| sheet.name().to_owned()),
            Some("Sheet1".into())
        );
    }

    #[test]
    fn clones_share_the_snapshot_and_handles_pin_it() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<Workbook>();
        assert_send_sync::<Sheet>();

        let workbook = Workbook::new().expect("valid baseline");
        let clone = workbook.clone();
        let sheet = workbook.active_sheet().expect("active sheet");
        drop(workbook);

        assert_eq!(sheet.name(), "Sheet1");
        assert_eq!(
            clone.active_sheet().map(|sheet| sheet.name().to_owned()),
            Some("Sheet1".into())
        );
        assert!(std::mem::size_of::<Workbook>() <= 2 * std::mem::size_of::<usize>());
    }

    #[test]
    fn flavor_is_content_derived() {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/custom/main.xml").expect("valid URI");
        let worksheet_uri = PackURI::new("/custom/sheet.xml").expect("valid URI");
        let mut workbook = BlobPart::new(
            workbook_uri,
            ct::SML_TEMPLATE_MAIN.into(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="8" state="veryHidden" r:id="rId1"/></sheets></workbook>"#.to_vec(),
        );
        workbook.relate_to("sheet.xml", rt::WORKSHEET);
        package.add_part(Box::new(workbook));
        package.add_part(Box::new(BlobPart::new(
            worksheet_uri,
            ct::SML_WORKSHEET.into(),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
        )));
        package.relate_to("custom/main.xml", rt::OFFICE_DOCUMENT);

        let workbook = Workbook::from_package(package).expect("valid template");
        let sheet = workbook.sheet("Data").expect("lookup").expect("present");
        assert_eq!(workbook.flavor(), Flavor::Template);
        assert!(workbook.flavor().is_template());
        assert!(matches!(sheet.visibility(), Visibility::VeryHidden));
    }

    #[test]
    fn ambiguous_names_and_dangling_relationships_are_typed_errors() {
        let duplicate_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="data" sheetId="2" r:id="rId2"/></sheets></workbook>"#;
        let mut package = package_with_workbook(duplicate_xml);
        let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
        let workbook = package.get_part_mut(&workbook_uri).expect("workbook part");
        workbook.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        workbook.relate_to("worksheets/sheet2.xml", rt::WORKSHEET);
        for index in 1..=2 {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new(format!("/xl/worksheets/sheet{index}.xml")).expect("valid URI"),
                ct::SML_WORKSHEET.into(),
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
            )));
        }
        let workbook = Workbook::from_package(package).expect("valid graph");
        assert!(matches!(
            workbook.sheet("DATA"),
            Err(Error::AmbiguousSheetName { matches: 2, .. })
        ));

        let dangling_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Missing" sheetId="1" r:id="absent"/></sheets></workbook>"#;
        assert!(matches!(
            Workbook::from_package(package_with_workbook(dangling_xml)),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn concurrent_snapshot_reads_need_no_public_locking() {
        let workbook = Workbook::new().expect("valid baseline");
        std::thread::scope(|scope| {
            for _ in 0..8 {
                let workbook = workbook.clone();
                scope.spawn(move || {
                    for _ in 0..1_000 {
                        let sheet = workbook.sheet("Sheet1").expect("lookup").expect("present");
                        assert_eq!(sheet.position(), 0);
                    }
                });
            }
        });
    }

    #[test]
    fn cell_facade_is_sparse_exact_and_non_mutating() {
        let workbook = Workbook::from_package(package_with_cells()).expect("valid workbook");
        let bytes_before = workbook.to_bytes().expect("serialize before lazy read");
        let sheet = workbook.sheet("data").expect("lookup").expect("present");

        assert!(matches!(
            sheet.cell("A1").expect("cell lookup"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "Office & Litchi"
        ));
        assert!(sheet.cell((0, 1)).expect("missing lookup").is_none());
        assert!(matches!(
            sheet.cell((1, 2)).expect("number lookup"),
            Some(Cell::Value(Value::Number(number))) if number.as_str() == "-0.000"
        ));
        let Some(Cell::Formula(formula)) = sheet.cell((2, 1)).expect("formula lookup") else {
            panic!("expected formula cell")
        };
        assert_eq!(formula.text(), "C2*2");
        assert!(matches!(
            formula.cached().map(Cache::value),
            Some(Value::Number(number)) if number.as_str() == "0"
        ));
        assert!(matches!(
            sheet.cell((4, 3)).expect("empty lookup"),
            Some(Cell::Empty)
        ));
        assert!(matches!(
            sheet.cell((litchi_sheet::ROWS, 0)),
            Err(Error::Coordinate(_))
        ));

        let range = Rect::at(0, 1, 4, 4).expect("valid range");
        let addresses = sheet
            .cells(range)
            .expect("sparse traversal")
            .map(|(address, _)| (address.row().get(), address.column().get()))
            .collect::<Vec<_>>();
        assert_eq!(addresses, [(1, 2), (2, 1)]);
        assert_eq!(
            sheet.stored_extent().expect("extent").map(Rect::end),
            Some((5, 4))
        );
        assert_eq!(
            workbook.to_bytes().expect("serialize after lazy read"),
            bytes_before
        );
    }

    #[test]
    fn concurrent_first_cell_read_publishes_one_safe_snapshot() {
        let workbook = Workbook::from_package(package_with_cells()).expect("valid workbook");
        let barrier = Arc::new(std::sync::Barrier::new(8));
        std::thread::scope(|scope| {
            for _ in 0..8 {
                let workbook = workbook.clone();
                let barrier = Arc::clone(&barrier);
                scope.spawn(move || {
                    barrier.wait();
                    let sheet = workbook.sheet("Data").expect("lookup").expect("present");
                    assert!(matches!(
                        sheet.cell((0, 0)).expect("cell lookup"),
                        Some(Cell::Value(Value::Text(text))) if text.as_str() == "Office & Litchi"
                    ));
                });
            }
        });
    }

    #[test]
    fn worksheet_operations_reject_other_sheet_kinds_without_parsing_them() {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
        let chart_uri = PackURI::new("/xl/chartsheets/sheet1.xml").expect("valid URI");
        let mut workbook = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.into(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Chart" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
        );
        workbook
            .rels_mut()
            .try_add_relationship(
                CHARTSHEET_REL.into(),
                "chartsheets/sheet1.xml".into(),
                "rId1".into(),
                TargetMode::Internal,
            )
            .expect("chart relationship");
        package.add_part(Box::new(workbook));
        package.add_part(Box::new(BlobPart::new(
            chart_uri,
            CHARTSHEET_CONTENT_TYPE.into(),
            b"not parsed by a worksheet operation".to_vec(),
        )));
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);

        let workbook = Workbook::from_package(package).expect("valid chart graph");
        let chart = workbook.sheet("Chart").expect("lookup").expect("present");
        assert!(matches!(
            chart.cell((0, 0)),
            Err(Error::NotWorksheet { .. })
        ));
    }

    #[test]
    fn poi_and_libreoffice_shared_formula_oracles_match() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let cases = [
            (
                root.join("test-data/poi/test-data/spreadsheet/shared_formulas.xlsx"),
                (40, 0),
                "B41",
            ),
            (
                root.join(
                    "test-data/libreoffice-core/sc/qa/unit/data/xlsx/shared-formula/basic.xlsx",
                ),
                (18, 1),
                "A19*10",
            ),
        ];
        for (path, address, expected) in cases {
            if !path.exists() {
                continue;
            }
            let workbook = Workbook::open(path).expect("corpus workbook");
            let sheet = workbook.sheet(0usize).expect("lookup").expect("present");
            let Some(Cell::Formula(formula)) = sheet.cell(address).expect("formula lookup") else {
                panic!("expected formula at {address:?}")
            };
            assert_eq!(formula.text(), expected);
        }
    }

    fn package_with_workbook(xml: &[u8]) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/workbook.xml").expect("valid URI"),
            ct::SML_SHEET_MAIN.into(),
            xml.to_vec(),
        )));
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package
    }

    fn package_with_cells() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").expect("valid URI");
        let strings_uri = PackURI::new("/xl/sharedStrings.xml").expect("valid URI");
        let mut workbook = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.into(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
        );
        workbook
            .rels_mut()
            .try_add_relationship(
                rt::WORKSHEET.into(),
                "worksheets/sheet1.xml".into(),
                "rId1".into(),
                TargetMode::Internal,
            )
            .expect("worksheet relationship");
        workbook
            .rels_mut()
            .try_add_relationship(
                rt::SHARED_STRINGS.into(),
                "sharedStrings.xml".into(),
                "rId2".into(),
                TargetMode::Internal,
            )
            .expect("shared-string relationship");
        package.add_part(Box::new(workbook));
        package.add_part(Box::new(BlobPart::new(
            worksheet_uri,
            ct::SML_WORKSHEET.into(),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="C2"><v>-0.000</v></c></row><row r="3"><c r="B3"><f>C2*2</f><v>0</v></c></row><row r="5"><c r="D5" s="2"/></row></sheetData></worksheet>"#.to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            strings_uri,
            ct::SML_SHARED_STRINGS.into(),
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><t>Office &amp; </t></r><r><t>Litchi</t></r></si></sst>"#.to_vec(),
        )));
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package
    }
}
