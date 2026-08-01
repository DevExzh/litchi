//! Immutable workbook snapshots and selector-first sheet lookup.

mod edit;

pub use edit::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DefaultsEdit, Edit, JoinError,
    JoinFailure, NewSheet, Patch, RowEdit, SheetEdit, State, TabEdit,
};

use std::collections::HashMap;
use std::convert::Infallible;
use std::io::{Read, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::Selector;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use litchi_sheet::{Area, At, ColumnAt, Rect, RowAt};

use crate::cell::{Extents, Store, Text};
use crate::error::{Error, Result, invalid};
use crate::raw;
use crate::style::StyleLineage;
use crate::{Cell, Cells, Column, Columns, LocalStyle, Row, Rows, Style, Styles};

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

impl Visibility {
    /// Whether Excel displays this sheet tab.
    pub const fn is_visible(&self) -> bool {
        matches!(self, Self::Visible)
    }

    /// Whether this tab is hidden by either recognized mechanism.
    pub const fn is_hidden(&self) -> bool {
        matches!(self, Self::Hidden | Self::VeryHidden)
    }

    /// Whether Excel omits this tab from its ordinary Unhide dialog.
    pub const fn is_very_hidden(&self) -> bool {
        matches!(self, Self::VeryHidden)
    }
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
    name_key: Box<str>,
    kind: SheetKind,
    visibility: Visibility,
    part_uri: PackURI,
    cells: OnceLock<Store>,
    #[allow(dead_code)]
    native_id: u32,
    relationship_id: String,
}

#[derive(Debug)]
pub(crate) struct Inner {
    package: OpcPackage,
    #[allow(dead_code)]
    workbook_uri: PackURI,
    shared_strings_uri: Option<PackURI>,
    shared_strings: OnceLock<Box<[Text]>>,
    styles_uri: Option<PackURI>,
    styles: OnceLock<raw::styles::Catalog>,
    pub(crate) style_lineage: Arc<StyleLineage>,
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
        let styles_uri = PackURI::new("/xl/styles.xml").map_err(invalid)?;

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
        workbook.rels_mut().try_add_relationship(
            rt::STYLES.to_owned(),
            "styles.xml".to_owned(),
            "rId2".to_owned(),
            TargetMode::Internal,
        )?;
        package.try_add_part(Box::new(workbook))?;
        package.try_add_part(Box::new(BlobPart::new(
            worksheet_uri,
            ct::SML_WORKSHEET.to_string(),
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
                r#"<dimension ref="A1"/><sheetData/></worksheet>"#
            )
            .as_bytes()
            .to_vec(),
        )))?;
        package.try_add_part(Box::new(BlobPart::new(
            styles_uri,
            ct::SML_STYLES.to_string(),
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
                r#"<fonts count="1"><font/></fonts>"#,
                r#"<fills count="2"><fill><patternFill patternType="none"/></fill>"#,
                r#"<fill><patternFill patternType="gray125"/></fill></fills>"#,
                r#"<borders count="1"><border/></borders>"#,
                r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
                r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
                r#"</styleSheet>"#
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
        Self::from_package_with_styles(package, None)
    }

    fn from_package_with_styles(package: OpcPackage, source: Option<&Workbook>) -> Result<Self> {
        let (workbook_uri, flavor, catalog, sheet_parts, shared_strings_uri, styles_uri) = {
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
            let styles_uri = validate_styles(&package, workbook)?;
            (
                workbook.partname().clone(),
                flavor,
                catalog,
                sheet_parts,
                shared_strings_uri,
                styles_uri,
            )
        };

        let active_sheet = if catalog.sheets.is_empty() {
            None
        } else {
            Some(catalog.active_sheet_index)
        };
        let style_lineage = match source {
            Some(source) if same_style_table(source, &package, styles_uri.as_ref())? => {
                Arc::clone(&source.inner.style_lineage)
            },
            Some(_) | None => Arc::new(StyleLineage),
        };
        let sheets = catalog
            .sheets
            .into_iter()
            .zip(sheet_parts)
            .enumerate()
            .map(|(position, (sheet, part))| {
                let name_key = crate::sheet::key(&sheet.name);
                Arc::new(SheetData {
                    position,
                    name: sheet.name,
                    name_key,
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
                styles_uri,
                styles: OnceLock::new(),
                style_lineage,
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
                let key = crate::sheet::key(&name);
                self.inner
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name_key == key)
                    .cloned()
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

    /// Shared immutable cell formats in this workbook snapshot.
    pub fn styles(&self) -> Result<Styles> {
        let len = self.inner.style_count()?;
        Ok(Styles::new(Arc::clone(&self.inner), len))
    }

    /// Serialize the immutable package snapshot to bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(PackageWriter::to_bytes(&self.inner.package)?)
    }

    /// Stream a finalized workbook to any sequential sink without seeking.
    ///
    /// A sink failure can leave caller-owned output incomplete. Use [`Self::save`]
    /// for atomic filesystem replacement.
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
        Ok(PackageWriter::write_to_stream(writer, &self.inner.package)?)
    }

    /// Atomically save through a finalized sibling temporary artifact.
    ///
    /// Serialization, flushing, and file synchronization finish before the
    /// destination is replaced. Existing symbolic-link destinations are
    /// refused instead of being followed or silently replaced.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        Ok(PackageWriter::write(path, &self.inner.package)?)
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

    /// Whether this is the active sheet in its immutable workbook snapshot.
    pub fn is_active(&self) -> bool {
        self.owner.active_sheet == Some(self.data.position)
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

    /// Exact local style state for a stored cell.
    ///
    /// `None` means no cell record exists. [`LocalStyle::Default`] means the
    /// record exists without an explicit shared-style reference.
    pub fn local_style<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<LocalStyle>> {
        let address = at.into().resolve()?;
        let Some(entry) = self.store()?.entry(address) else {
            return Ok(None);
        };
        entry.style.map_or(Ok(Some(LocalStyle::Default)), |key| {
            self.owner.require_style(key)?;
            Ok(Some(LocalStyle::Shared(Style::from_raw(
                Arc::clone(&self.owner),
                key,
            ))))
        })
    }

    /// Effective shared style for a stored cell.
    ///
    /// Cells without a local style resolve to the base shared format. If the
    /// workbook has no style part, an unstyled cell resolves to `None`.
    pub fn style<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<Style>> {
        let address = at.into().resolve()?;
        let Some(entry) = self.store()?.entry(address) else {
            return Ok(None);
        };
        let key = entry.style.unwrap_or(0);
        if self.owner.style_count()? == 0 {
            return Ok(None);
        }
        self.owner.require_style(key)?;
        Ok(Some(Style::from_raw(Arc::clone(&self.owner), key)))
    }

    /// Lazily traverse stored cells selected by A1 range, raw zero-based
    /// half-open bounds, or a reusable checked rectangle.
    pub fn cells<'a>(&self, area: impl Into<Area<'a>>) -> Result<Cells<'_>> {
        let range = area.into().resolve()?;
        Ok(self.store()?.cells(range))
    }

    /// Borrow one checked logical row, including an implicit default row.
    pub fn row(&self, at: impl Into<RowAt>) -> Result<Row<'_>> {
        let index = at.into().resolve()?;
        Ok(self.store()?.row(index))
    }

    /// Lazily traverse only explicit worksheet row records.
    pub fn rows(&self) -> Result<Rows<'_>> {
        Ok(self.store()?.rows())
    }

    /// Exact shared-style state contributed by a stored row record.
    ///
    /// `None` means the logical row is implicit. [`LocalStyle::Default`]
    /// means an explicit record applies without a shared-style reference.
    pub fn row_style(&self, at: impl Into<RowAt>) -> Result<Option<LocalStyle>> {
        let index = at.into().resolve()?;
        let Some(entry) = self.store()?.row_entry(index) else {
            return Ok(None);
        };
        entry
            .properties
            .style
            .map_or(Ok(Some(LocalStyle::Default)), |key| {
                self.owner.require_style(key)?;
                Ok(Some(LocalStyle::Shared(Style::from_raw(
                    Arc::clone(&self.owner),
                    key,
                ))))
            })
    }

    /// Borrow one checked logical column, including an implicit default
    /// column. A1 labels such as `"B"` are the primary entry; raw inputs are
    /// zero-based and validated before lookup.
    pub fn column<'a>(&self, at: impl Into<ColumnAt<'a>>) -> Result<Column<'_>> {
        let index = at.into().resolve()?;
        Ok(self.store()?.column(index))
    }

    /// Lazily traverse logical columns covered by explicit property records.
    /// Overlapping producer records have already been resolved using Excel's
    /// last-record-wins semantics.
    pub fn columns(&self) -> Result<Columns<'_>> {
        Ok(self.store()?.columns())
    }

    /// Stored worksheet-grid defaults, if the producer supplied them.
    ///
    /// Absence is preserved rather than guessing a font-dependent row height
    /// or column width.
    pub fn defaults(&self) -> Result<Option<&crate::layout::Defaults>> {
        Ok(self.store()?.defaults())
    }

    /// Exact shared-style state contributed by a column-property record.
    ///
    /// `None` means the logical column is implicit. [`LocalStyle::Default`]
    /// means an explicit record applies without a shared-style reference.
    pub fn column_style<'a>(&self, at: impl Into<ColumnAt<'a>>) -> Result<Option<LocalStyle>> {
        let index = at.into().resolve()?;
        let Some(entry) = self.store()?.column_entry(index) else {
            return Ok(None);
        };
        entry
            .properties
            .style
            .map_or(Ok(Some(LocalStyle::Default)), |key| {
                self.owner.require_style(key)?;
                Ok(Some(LocalStyle::Shared(Style::from_raw(
                    Arc::clone(&self.owner),
                    key,
                ))))
            })
    }

    /// Distinct declared, stored, content, and directly styled cell bounds.
    pub fn extents(&self) -> Result<&Extents> {
        Ok(self.store()?.extents())
    }

    /// Bounding rectangle of stored cell records, distinct from declared,
    /// formatted, and content extents.
    pub fn stored_extent(&self) -> Result<Option<Rect>> {
        Ok(self.store()?.extents().stored())
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
        self.owner.validate_styles(&parsed)?;
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

    fn style_catalog(&self) -> Result<Option<&raw::styles::Catalog>> {
        let Some(uri) = self.styles_uri.as_ref() else {
            return Ok(None);
        };
        if let Some(styles) = self.styles.get() {
            return Ok(Some(styles));
        }

        let part = self.package.get_part(uri)?;
        let parsed = raw::styles::parse(part.blob())?;
        let _ = self.styles.set(parsed);
        self.styles
            .get()
            .map(Some)
            .ok_or_else(|| invalid("style cache initialization did not publish a value"))
    }

    pub(crate) fn style_count(&self) -> Result<u32> {
        Ok(self.style_catalog()?.map_or(0, raw::styles::Catalog::len))
    }

    fn require_style(&self, key: u32) -> Result<()> {
        let len = self.style_count()?;
        if key >= len {
            return Err(invalid(format!(
                "worksheet cell references shared style {key}, but the workbook contains {len} cell formats"
            )));
        }
        Ok(())
    }

    fn validate_styles(&self, store: &Store) -> Result<()> {
        if !store.entries().iter().any(|entry| entry.style.is_some())
            && !store
                .row_entries()
                .iter()
                .any(|entry| entry.properties.style.is_some())
            && !store
                .column_entries()
                .iter()
                .any(|entry| entry.properties.style.is_some())
        {
            return Ok(());
        }
        let len = self.style_count()?;
        if let Some(entry) = store
            .entries()
            .iter()
            .find(|entry| entry.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet cell {} references shared style {}, but the workbook contains {len} cell formats",
                entry.address,
                entry.style.unwrap_or_default()
            )));
        }
        if let Some(entry) = store
            .row_entries()
            .iter()
            .find(|entry| entry.properties.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet row {} references shared style {}, but the workbook contains {len} cell formats",
                entry.index,
                entry.properties.style.unwrap_or_default()
            )));
        }
        if let Some(entry) = store
            .column_entries()
            .iter()
            .find(|entry| entry.properties.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet column {} references shared style {}, but the workbook contains {len} cell formats",
                entry.first,
                entry.properties.style.unwrap_or_default()
            )));
        }
        Ok(())
    }

    pub(crate) fn style_fan_out(self: &Arc<Self>, key: u32) -> Result<usize> {
        self.require_style(key)?;
        let mut count = 0usize;
        for data in &self.sheets {
            if data.kind != SheetKind::Worksheet {
                continue;
            }
            let sheet = Sheet {
                owner: Arc::clone(self),
                data: Arc::clone(data),
            };
            count = count
                .checked_add(
                    sheet
                        .store()?
                        .entries()
                        .iter()
                        .filter(|entry| entry.style.unwrap_or(0) == key)
                        .count(),
                )
                .ok_or_else(|| invalid("shared style fan-out count overflowed usize"))?;
        }
        Ok(count)
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
    let mut targets = HashMap::<PackURI, usize>::with_capacity(sheets.len());
    for (position, sheet) in sheets.iter().enumerate() {
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
        let uri = part.partname().clone();
        if let Some(previous) = targets.insert(uri.clone(), position) {
            return Err(invalid(format!(
                "sheet part '{uri}' is referenced by both '{}' and '{}'",
                sheets[previous].name, sheet.name
            )));
        }
        parts.push(SheetPart { kind, uri });
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

fn validate_styles(package: &OpcPackage, workbook: &dyn Part) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::STYLES | rt::STRICT_STYLES))
    {
        if found.is_some() {
            return Err(invalid("workbook has multiple styles relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("styles relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.get_part(&uri)?;
        if part.content_type() != ct::SML_STYLES {
            return Err(invalid(format!(
                "styles part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_STYLES
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

fn same_style_table(
    source: &Workbook,
    package: &OpcPackage,
    styles_uri: Option<&PackURI>,
) -> Result<bool> {
    let (Some(source_uri), Some(styles_uri)) = (source.inner.styles_uri.as_ref(), styles_uri)
    else {
        return Ok(source.inner.styles_uri.is_none() && styles_uri.is_none());
    };
    if source_uri != styles_uri {
        return Ok(false);
    }
    let before = source.inner.package.get_part(source_uri)?;
    let after = package.get_part(styles_uri)?;
    if before.content_type() != after.content_type() {
        return Ok(false);
    }
    let before_blob = before.blob_arc();
    let after_blob = after.blob_arc();
    Ok(Arc::ptr_eq(&before_blob, &after_blob) || before_blob.as_slice() == after_blob.as_slice())
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

    #[derive(Default)]
    struct WriteOnly(Vec<u8>);

    impl std::io::Write for WriteOnly {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.0.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn new_workbook_is_deterministic_and_selector_first() {
        let first = Workbook::new().expect("valid baseline");
        let second = Workbook::new().expect("valid baseline");

        assert_eq!(first.to_bytes().ok(), second.to_bytes().ok());

        let mut streamed = WriteOnly::default();
        first.write_to(&mut streamed).expect("stream workbook");
        assert_eq!(
            streamed.0,
            first.to_bytes().expect("buffered serialization")
        );
        let streamed = Workbook::from_slice(&streamed.0).expect("reopen streamed workbook");
        assert_eq!(streamed.len(), 1);
        assert_eq!(
            streamed
                .sheet("Sheet1")
                .expect("lookup")
                .expect("default sheet")
                .name(),
            "Sheet1"
        );
        assert_eq!(first.len(), 1);
        assert_eq!(first.flavor(), Flavor::Workbook);
        assert_eq!(first.date_system(), DateSystem::Excel1900);

        let by_name = first.sheet("sheet1").expect("lookup").expect("present");
        let by_position = first.sheet(0usize).expect("lookup").expect("present");
        let checked_name = crate::sheet::Name::new("SHEET1").expect("checked name");
        assert!(
            first
                .sheet(&checked_name)
                .expect("checked lookup")
                .is_some()
        );
        assert!(first.sheet(checked_name).expect("moved lookup").is_some());
        assert_eq!(by_name.name(), "Sheet1");
        assert_eq!(by_position.position(), 0);
        assert!(by_name.same_workbook(&by_position));
        assert!(matches!(by_name.kind(), SheetKind::Worksheet));
        assert!(matches!(by_name.visibility(), Visibility::Visible));
        assert!(first.sheet(1usize).expect("lookup").is_none());
        let extents = by_name.extents().expect("empty extents");
        assert_eq!(extents.declared().map(Rect::a1).as_deref(), Some("A1"));
        assert_eq!(extents.stored(), None);
        assert_eq!(extents.content(), None);
        assert_eq!(extents.styled(), None);

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
        assert_send_sync::<Style>();
        assert_send_sync::<Styles>();
        assert_send_sync::<crate::StyleKey>();
        assert_send_sync::<LocalStyle>();
        assert_send_sync::<Extents>();

        let workbook = Workbook::new().expect("valid baseline");
        let clone = workbook.clone();
        let sheet = workbook.active_sheet().expect("active sheet");
        let style = workbook
            .styles()
            .expect("styles")
            .base()
            .expect("base style");
        drop(workbook);

        assert_eq!(sheet.name(), "Sheet1");
        assert_eq!(style.fan_out().expect("fan-out"), 0);
        assert_eq!(
            clone.active_sheet().map(|sheet| sheet.name().to_owned()),
            Some("Sheet1".into())
        );
        assert!(std::mem::size_of::<Workbook>() <= 2 * std::mem::size_of::<usize>());
        assert!(std::mem::size_of::<Style>() <= 2 * std::mem::size_of::<usize>());
        assert!(std::mem::size_of::<Styles>() <= 2 * std::mem::size_of::<usize>());
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
    fn duplicate_names_and_dangling_relationships_are_typed_errors() {
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
        assert!(matches!(
            Workbook::from_package(package),
            Err(Error::SheetNameConflict {
                first: 0,
                second: 1,
                ..
            })
        ));

        let dangling_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Missing" sheetId="1" r:id="absent"/></sheets></workbook>"#;
        assert!(matches!(
            Workbook::from_package(package_with_workbook(dangling_xml)),
            Err(Error::Invalid(_))
        ));

        let aliased_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="rId1"/><sheet name="Two" sheetId="2" r:id="rId2"/></sheets></workbook>"#;
        let mut package = package_with_workbook(aliased_xml);
        let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
        let workbook = package.get_part_mut(&workbook_uri).expect("workbook part");
        for id in ["rId1", "rId2"] {
            workbook
                .rels_mut()
                .try_add_relationship(
                    rt::WORKSHEET.to_owned(),
                    "worksheets/sheet1.xml".to_owned(),
                    id.to_owned(),
                    TargetMode::Internal,
                )
                .expect("sheet relationship");
        }
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/worksheets/sheet1.xml").expect("valid URI"),
            ct::SML_WORKSHEET.into(),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
        )));
        assert!(matches!(
            Workbook::from_package(package),
            Err(Error::Invalid(message)) if message.contains("referenced by both 'One' and 'Two'")
        ));
    }

    #[test]
    fn styles_graph_table_and_cell_references_are_checked() {
        let baseline = Workbook::new().expect("baseline");

        let mut duplicate = baseline.inner.package.clone();
        duplicate
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .rels_mut()
            .try_add_relationship(
                rt::STYLES.into(),
                "styles.xml".into(),
                "rId3".into(),
                TargetMode::Internal,
            )
            .expect("second styles relationship");
        assert!(matches!(
            Workbook::from_package(duplicate),
            Err(Error::Invalid(message)) if message.contains("multiple styles relationships")
        ));

        let mut external = baseline.inner.package.clone();
        let rels = external
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .rels_mut();
        rels.remove("rId2").expect("styles relationship");
        rels.try_add_relationship(
            rt::STYLES.into(),
            "https://example.invalid/styles.xml".into(),
            "rId2".into(),
            TargetMode::External,
        )
        .expect("external styles relationship");
        assert!(matches!(
            Workbook::from_package(external),
            Err(Error::Invalid(message)) if message.contains("cannot be external")
        ));

        let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
        let mut wrong_type = baseline.inner.package.clone();
        wrong_type
            .get_part_mut(&styles_uri)
            .expect("styles part")
            .set_content_type("application/xml".into())
            .expect("replace content type");
        assert!(matches!(
            Workbook::from_package(wrong_type),
            Err(Error::Invalid(message)) if message.contains("styles part has content type")
        ));

        let mut malformed = baseline.inner.package.clone();
        malformed
            .get_part_mut(&styles_uri)
            .expect("styles part")
            .set_blob(
                br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf/></cellXfs></styleSheet>"#.to_vec(),
            );
        let malformed = Workbook::from_package(malformed).expect("graph remains lazy");
        assert!(matches!(malformed.styles(), Err(Error::Invalid(_))));

        let mut dangling = baseline.inner.package.clone();
        dangling
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>"#.to_vec(),
            );
        let dangling = Workbook::from_package(dangling).expect("lazy worksheet");
        assert!(matches!(
            dangling
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .cell("A1"),
            Err(Error::Invalid(message)) if message.contains("A1 references shared style 1")
        ));

        let mut dangling_column = baseline.inner.package.clone();
        dangling_column
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="2" max="2" style="1"/></cols><sheetData/></worksheet>"#.to_vec(),
            );
        let dangling_column = Workbook::from_package(dangling_column).expect("lazy column style");
        assert!(matches!(
            dangling_column
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .column(1),
            Err(Error::Invalid(message)) if message.contains("column 1 references shared style 1")
        ));

        let mut dangling_row = baseline.inner.package.clone();
        dangling_row
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("sheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2" s="1" customFormat="1"/></sheetData></worksheet>"#.to_vec(),
            );
        let dangling_row = Workbook::from_package(dangling_row).expect("lazy row style");
        assert!(matches!(
            dangling_row
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .row(1),
            Err(Error::Invalid(message)) if message.contains("row 1 references shared style 1")
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
        assert!(sheet.row(0).expect("stored row 1").stored());
        assert!(!sheet.row(3).expect("implicit row 4").stored());
        assert!(!sheet.row(4).expect("stored row 5").hidden());
        assert!(matches!(
            sheet.row(litchi_sheet::ROWS),
            Err(Error::Coordinate(_))
        ));
        assert_eq!(
            sheet
                .rows()
                .expect("stored rows")
                .map(|row| row.index().get())
                .collect::<Vec<_>>(),
            [0, 1, 2, 4]
        );

        let addresses = sheet
            .cells("B1:D4")
            .expect("sparse traversal")
            .map(|(address, _)| (address.row().get(), address.column().get()))
            .collect::<Vec<_>>();
        assert_eq!(addresses, [(1, 2), (2, 1)]);
        assert!(matches!(sheet.cells("B2:A1"), Err(Error::Range(_))));
        let extents = sheet.extents().expect("cell extents");
        assert_eq!(extents.declared(), None);
        assert_eq!(extents.stored().map(Rect::a1).as_deref(), Some("A1:D5"));
        assert_eq!(extents.content().map(Rect::a1).as_deref(), Some("A1:C3"));
        assert_eq!(extents.styled().map(Rect::a1).as_deref(), Some("D5"));
        assert_eq!(extents.used().map(Rect::a1).as_deref(), Some("A1:D5"));
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
        let styles_uri = PackURI::new("/xl/styles.xml").expect("valid URI");
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
        workbook
            .rels_mut()
            .try_add_relationship(
                rt::STYLES.into(),
                "styles.xml".into(),
                "rId3".into(),
                TargetMode::Internal,
            )
            .expect("styles relationship");
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
        package.add_part(Box::new(BlobPart::new(
            styles_uri,
            ct::SML_STYLES.into(),
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="3"><xf/><xf numFmtId="1"/><xf numFmtId="2"/></cellXfs></styleSheet>"#.to_vec(),
        )));
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package
    }
}
