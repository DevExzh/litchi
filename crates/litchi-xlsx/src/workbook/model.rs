//! Semantic workbook snapshot and worksheet models.

use std::convert::Infallible;
use std::io::{Read, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

use super::edit::{Commit, Edit, Patch};
use super::worksheet;
use super::{codec, package};
use litchi_core::Selector as CoreSelector;
use litchi_opc::{OpcPackage, PackURI, ReadLimits};
use litchi_sheet::{Area, At, ColumnAt, Rect, RowAt};
use once_cell::sync::OnceCell;

use crate::cell::{Extents, Store, Text};
use crate::error::{Error, Result, invalid};
use crate::raw;
use crate::style::StyleLineage;
use crate::{Cells, Column, Columns, LocalStyle, Row, Rows, Style, Styles};

#[cfg(feature = "encryption")]
use litchi_ooxml_common::package_encryption::PackageEncryption;

/// Semantic selector accepted by [`Workbook::sheet`].
///
/// Names and checked zero-based positions are the ordinary entry points. The
/// uninhabited identity variant reserves room for a future lineage-checked
/// durable selector without exposing native SpreadsheetML IDs.
pub type Selector<'a> = litchi_sheet::SheetSelector<'a, Infallible>;

/// Runtime workbook flavor derived from the main-part content type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Flavor {
    Workbook,
    Template,
    MacroWorkbook,
    MacroTemplate,
}

impl Flavor {
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
pub enum WorksheetKind {
    Worksheet,
    Chart,
    Dialog,
    Macro,
    Unknown,
}

/// Worksheet visibility retained without approximating producer extensions.
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

#[derive(Debug)]
pub(super) struct SheetData {
    pub(super) position: usize,
    pub(super) name: String,
    pub(super) name_key: Box<str>,
    pub(super) kind: WorksheetKind,
    pub(super) visibility: Visibility,
    pub(super) part_uri: PackURI,
    pub(super) cells: OnceLock<Store>,
    pub(super) web_bindings: OnceCell<crate::web::Bindings>,
    #[allow(dead_code)]
    pub(super) native_id: u32,
    pub(super) relationship_id: String,
}

#[derive(Debug)]
pub(crate) struct Inner {
    pub(super) package: OpcPackage,
    #[allow(dead_code)]
    pub(super) workbook_uri: PackURI,
    pub(super) shared_strings_uri: Option<PackURI>,
    pub(super) shared_strings: OnceLock<Box<[Text]>>,
    pub(super) styles_uri: Option<PackURI>,
    pub(super) styles: OnceLock<raw::styles::Catalog>,
    pub(super) task_panes: OnceCell<Option<litchi_ooxml_common::web::Panes>>,
    pub(crate) style_lineage: Arc<StyleLineage>,
    pub(super) flavor: Flavor,
    pub(super) date_system: DateSystem,
    pub(super) active_sheet: Option<usize>,
    pub(super) sheets: Box<[Arc<SheetData>]>,
    pub(super) defined_names: Box<[raw::DefinedName]>,
    pub(super) pivot_caches: Box<[raw::PivotCache]>,
    pub(super) external_reference_ids: Box<[String]>,
}

/// Immutable, cheap-to-share XLSX workbook snapshot.
#[derive(Debug, Clone)]
pub struct Workbook {
    pub(super) inner: Arc<Inner>,
    #[cfg(feature = "encryption")]
    encryption: PackageEncryption,
}

impl Workbook {
    /// Create a deterministic minimal workbook with one visible worksheet.
    pub fn new() -> Result<Self> {
        Self::from_package(crate::package::build_minimal_package()?)
    }

    /// Create a deterministic minimal workbook with one visible worksheet.
    pub fn create() -> Result<Self> {
        Self::new()
    }

    /// Open an XLSX-family package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, ReadLimits::default())
    }

    /// Open an XLSX-family package from a filesystem path with explicit OPC
    /// resource limits.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_package(OpcPackage::open_with_limits(path, limits)?)
    }

    /// Open an ordinary or encrypted workbook using safe independent resource
    /// policies.
    #[cfg(feature = "encryption")]
    pub fn open_with_password(path: impl AsRef<Path>, password: &str) -> Result<Self> {
        crate::Package::open_with_password(path, password)?.into_workbook()
    }

    /// Open an ordinary or encrypted workbook with independent outer
    /// encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn open_with_password_and_limits(
        path: impl AsRef<Path>,
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        crate::Package::open_with_password_and_limits(
            path,
            password,
            encryption_limits,
            opc_limits,
        )?
        .into_workbook()
    }

    /// Move bytes into the XLSX parser.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, ReadLimits::default())
    }

    /// Move bytes into the XLSX parser with explicit OPC resource limits.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        let package = OpcPackage::from_vec_with_limits(bytes, limits)?;
        Self::from_package(package)
    }

    /// Move ordinary or encrypted bytes into the workbook parser using safe
    /// independent resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: &str) -> Result<Self> {
        crate::Package::from_bytes_with_password(bytes, password)?.into_workbook()
    }

    /// Move ordinary or encrypted bytes into the workbook parser with
    /// independent outer encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_bytes_with_password_and_limits(
        bytes: Vec<u8>,
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        crate::Package::from_bytes_with_password_and_limits(
            bytes,
            password,
            encryption_limits,
            opc_limits,
        )?
        .into_workbook()
    }

    /// Open a borrowed XLSX byte slice.
    pub fn from_slice(bytes: &[u8]) -> Result<Self> {
        Self::from_slice_with_limits(bytes, ReadLimits::default())
    }

    /// Open a borrowed XLSX byte slice with explicit OPC resource limits.
    pub fn from_slice_with_limits(bytes: &[u8], limits: ReadLimits) -> Result<Self> {
        Self::from_package(OpcPackage::from_bytes_with_limits(bytes, limits)?)
    }

    /// Read an ordinary or encrypted borrowed byte slice using safe
    /// independent resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_slice_with_password(bytes: &[u8], password: &str) -> Result<Self> {
        crate::Package::from_slice_with_password(bytes, password)?.into_workbook()
    }

    /// Read an ordinary or encrypted borrowed byte slice with independent
    /// outer encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_slice_with_password_and_limits(
        bytes: &[u8],
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        crate::Package::from_slice_with_password_and_limits(
            bytes,
            password,
            encryption_limits,
            opc_limits,
        )?
        .into_workbook()
    }

    /// Read an XLSX package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limits(reader, ReadLimits::default())
    }

    /// Read an XLSX package from a synchronous reader with explicit OPC
    /// resource limits.
    pub fn from_reader_with_limits(reader: impl Read, limits: ReadLimits) -> Result<Self> {
        Self::from_package(OpcPackage::from_reader_with_limits(reader, limits)?)
    }

    /// Read an ordinary or encrypted workbook from a synchronous reader using
    /// safe independent resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password(reader: impl Read, password: &str) -> Result<Self> {
        crate::Package::from_reader_with_password(reader, password)?.into_workbook()
    }

    /// Read an ordinary or encrypted workbook from a synchronous reader with
    /// independent outer encryption and inner OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password_and_limits(
        reader: impl Read,
        password: &str,
        encryption_limits: &crate::encryption::Limits,
        opc_limits: ReadLimits,
    ) -> Result<Self> {
        crate::Package::from_reader_with_password_and_limits(
            reader,
            password,
            encryption_limits,
            opc_limits,
        )?
        .into_workbook()
    }

    /// Build a snapshot from a validated OPC package without exposing it in
    /// ordinary sheet APIs.
    ///
    /// This raw compatibility boundary deliberately treats the supplied clear
    /// OPC graph as plaintext and does not infer encryption provenance.
    pub fn from_package(package: OpcPackage) -> Result<Self> {
        Self::from_package_with_styles(package, None)
    }

    /// Adopt a parsed OPC package after validating its SpreadsheetML graph.
    pub fn from_opc(package: OpcPackage) -> Result<Self> {
        Self::from_package(package)
    }

    #[cfg(feature = "encryption")]
    pub(crate) fn from_package_with_encryption(
        package: OpcPackage,
        encryption: PackageEncryption,
    ) -> Result<Self> {
        let mut workbook = Self::from_package(package)?;
        workbook.encryption = encryption;
        Ok(workbook)
    }

    pub(super) fn from_package_with_styles(
        package: OpcPackage,
        source: Option<&Workbook>,
    ) -> Result<Self> {
        let (workbook_uri, flavor, catalog, sheet_parts, shared_strings_uri, styles_uri) = {
            let workbook = package.main_document_part()?;
            let flavor = codec::flavor(workbook.content_type()).ok_or_else(|| {
                invalid(format!(
                    "main part '{}' has non-XLSX content type '{}'",
                    workbook.partname(),
                    workbook.content_type()
                ))
            })?;
            let catalog = raw::parse_catalog(workbook.blob())?;
            let sheet_parts = package::validate_sheet_graph(&package, workbook, &catalog.sheets)?;
            let shared_strings_uri = package::validate_shared_strings(&package, workbook)?;
            let styles_uri = package::validate_styles(&package, workbook)?;
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
            Some(source) if package::same_style_table(source, &package, styles_uri.as_ref())? => {
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
                    visibility: codec::visibility(sheet.visibility),
                    part_uri: part.uri,
                    cells: OnceLock::new(),
                    web_bindings: OnceCell::new(),
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
                task_panes: OnceCell::new(),
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
            #[cfg(feature = "encryption")]
            encryption: source.map_or_else(PackageEncryption::plain, |source| source.encryption),
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
    pub fn sheets(&self) -> impl ExactSizeIterator<Item = Worksheet> + DoubleEndedIterator + '_ {
        self.inner.sheets.iter().cloned().map(|data| Worksheet {
            owner: Arc::clone(&self.inner),
            data,
        })
    }

    /// Look up a sheet by developer-facing name or checked zero-based position.
    pub fn sheet<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<Worksheet>> {
        let data = match selector.into() {
            CoreSelector::Position(position) => self.inner.sheets.get(position.get()).cloned(),
            CoreSelector::Name(name) => {
                let key = crate::sheet::key(&name);
                self.inner
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name_key == key)
                    .cloned()
            },
            CoreSelector::Id(never) => match never {},
            _ => return Err(Error::UnsupportedSelector),
        };
        Ok(data.map(|data| Worksheet {
            owner: Arc::clone(&self.inner),
            data,
        }))
    }

    /// Return the active sheet when the workbook contains sheets.
    pub fn active_sheet(&self) -> Option<Worksheet> {
        let data = self
            .inner
            .active_sheet
            .and_then(|position| self.inner.sheets.get(position))
            .cloned()?;
        Some(Worksheet {
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

    /// Read the workbook's source-bound calculation metadata.
    pub fn calculation_metadata(&self) -> Result<crate::calculation_properties::Snapshot> {
        crate::calculation_properties::Snapshot::load(&self.inner.package)
    }

    /// Read calculation metadata with a caller-supplied resource policy.
    pub fn calculation_metadata_with_limits(
        &self,
        limits: &crate::calculation_properties::Limits,
    ) -> Result<crate::calculation_properties::Snapshot> {
        crate::calculation_properties::Snapshot::load_with_limits(&self.inner.package, limits)
    }

    /// Persisted Office Add-in task panes, when this package contains them.
    ///
    /// The complete package graph is validated on first access. Later calls
    /// borrow the model retained by this immutable workbook snapshot.
    pub fn task_panes(&self) -> Result<Option<&litchi_ooxml_common::web::Panes>> {
        let panes = self
            .inner
            .task_panes
            .get_or_try_init(|| litchi_ooxml_common::web::load(&self.inner.package))?;
        Ok(panes.as_ref())
    }

    /// Workbook-level protection metadata, when `workbookProtection` is
    /// present. The values are passive verifier and lock metadata; they do
    /// not enforce an editing policy.
    pub fn workbook_protection_metadata(
        &self,
    ) -> Result<Option<crate::workbook_metadata::protection::Metadata>> {
        let workbook = self.inner.package.get_part(&self.inner.workbook_uri)?;
        crate::workbook_metadata::protection::parse_workbook_protection(workbook.blob())
    }

    /// Shared immutable cell formats in this workbook snapshot.
    pub fn styles(&self) -> Result<Styles> {
        let len = self.inner.style_count()?;
        Ok(Styles::new(Arc::clone(&self.inner), len))
    }

    /// Serialize the immutable package snapshot to bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.ensure_ordinary_output("to_bytes")?;
        self.to_plain_bytes()
    }

    /// Explicitly serialize a plaintext OPC package, declassifying encrypted
    /// source provenance for this output only.
    pub fn to_plain_bytes(&self) -> Result<Vec<u8>> {
        crate::writer::to_bytes(&self.inner.package)
    }

    /// Stream a finalized workbook to any sequential sink without seeking.
    ///
    /// A sink failure can leave caller-owned output incomplete. Use [`Self::save`]
    /// for atomic filesystem replacement.
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
        self.ensure_ordinary_output("write_to")?;
        self.write_plain_to(writer)
    }

    /// Explicitly stream the plaintext OPC package to a sequential sink.
    pub fn write_plain_to(&self, writer: impl Write) -> Result<()> {
        crate::writer::write_to(&self.inner.package, writer)
    }

    /// Atomically save through a finalized sibling temporary artifact.
    ///
    /// Serialization, flushing, and file synchronization finish before the
    /// destination is replaced. Existing symbolic-link destinations are
    /// refused instead of being followed or silently replaced.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.ensure_ordinary_output("save")?;
        self.save_plain(path)
    }

    /// Explicitly save the plaintext OPC package atomically.
    pub fn save_plain(&self, path: impl AsRef<Path>) -> Result<()> {
        crate::writer::save(&self.inner.package, path)
    }

    /// Encryption profile retained from source ingress or the latest
    /// successful encrypted save. In-memory byte generation is side-effect-free.
    #[cfg(feature = "encryption")]
    pub const fn encryption(&self) -> Option<crate::encryption::Mode> {
        self.encryption.mode()
    }

    /// Serialize and encrypt using an explicitly selected profile.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted(&self, password: &str, mode: crate::encryption::Mode) -> Result<Vec<u8>> {
        self.to_encrypted_with_limits(password, mode, &crate::encryption::Limits::default())
    }

    /// Serialize and encrypt using an explicit encryption resource policy.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted_with_limits(
        &self,
        password: &str,
        mode: crate::encryption::Mode,
        limits: &crate::encryption::Limits,
    ) -> Result<Vec<u8>> {
        self.encrypt_with_mode(password, mode, limits)
    }

    /// Serialize and encrypt using the source package's retained profile.
    #[cfg(feature = "encryption")]
    pub fn to_reencrypted(&self, password: &str) -> Result<Vec<u8>> {
        self.to_reencrypted_with_limits(password, &crate::encryption::Limits::default())
    }

    /// Re-encrypt using the retained profile and explicit resource policy.
    #[cfg(feature = "encryption")]
    pub fn to_reencrypted_with_limits(
        &self,
        password: &str,
        limits: &crate::encryption::Limits,
    ) -> Result<Vec<u8>> {
        let mode = self.retained_mode("to_reencrypted")?;
        self.to_encrypted_with_limits(password, mode, limits)
    }

    /// Atomically save with an explicitly selected encryption profile.
    #[cfg(feature = "encryption")]
    pub fn save_encrypted(
        &mut self,
        path: impl AsRef<Path>,
        password: &str,
        mode: crate::encryption::Mode,
    ) -> Result<()> {
        self.save_encrypted_with_limits(path, password, mode, &crate::encryption::Limits::default())
    }

    /// Atomically save with explicit encryption profile and resource policy.
    #[cfg(feature = "encryption")]
    pub fn save_encrypted_with_limits(
        &mut self,
        path: impl AsRef<Path>,
        password: &str,
        mode: crate::encryption::Mode,
        limits: &crate::encryption::Limits,
    ) -> Result<()> {
        let output = self.encrypt_with_mode(password, mode, limits)?;
        crate::writer::save_encrypted(&output, path)?;
        self.encryption.mark_encrypted(mode);
        Ok(())
    }

    /// Atomically save using the encrypted source's retained profile.
    #[cfg(feature = "encryption")]
    pub fn save_reencrypted(&mut self, path: impl AsRef<Path>, password: &str) -> Result<()> {
        self.save_reencrypted_with_limits(path, password, &crate::encryption::Limits::default())
    }

    /// Atomically save using the retained profile and explicit resource policy.
    #[cfg(feature = "encryption")]
    pub fn save_reencrypted_with_limits(
        &mut self,
        path: impl AsRef<Path>,
        password: &str,
        limits: &crate::encryption::Limits,
    ) -> Result<()> {
        let mode = self.retained_mode("save_reencrypted")?;
        self.save_encrypted_with_limits(path, password, mode, limits)
    }

    /// Start an isolated semantic transaction over this immutable snapshot.
    pub fn edit(&self) -> Result<Edit> {
        self.ensure_mutation_allowed("edit")?;
        Edit::new(self.clone())
    }

    /// Apply a reversible patch after checking every expected source part.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        self.ensure_mutation_allowed("apply")?;
        patch.apply_to(self)
    }

    fn ensure_ordinary_output(&self, _operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy {
                operation: _operation,
                source,
            })?;
        Ok(())
    }

    fn ensure_mutation_allowed(&self, operation: &'static str) -> Result<()> {
        self.ensure_ordinary_output(operation)
    }

    #[cfg(feature = "encryption")]
    fn retained_mode(&self, operation: &'static str) -> Result<crate::encryption::Mode> {
        self.encryption
            .require_retained_mode()
            .map_err(|source| Error::EncryptionPolicy { operation, source })
    }

    #[cfg(feature = "encryption")]
    fn encrypt_with_mode(
        &self,
        password: &str,
        mode: crate::encryption::Mode,
        limits: &crate::encryption::Limits,
    ) -> Result<Vec<u8>> {
        let plaintext = self.to_plain_bytes()?;
        Ok(crate::encryption::encrypt_with(
            plaintext, password, mode, limits,
        )?)
    }
}

/// Lightweight lifetime-free handle to one sheet in a workbook snapshot.
#[derive(Debug, Clone)]
pub struct Worksheet {
    pub(super) owner: Arc<Inner>,
    pub(super) data: Arc<SheetData>,
}

impl Worksheet {
    /// Developer-facing sheet name.
    pub fn name(&self) -> &str {
        &self.data.name
    }

    /// Checked zero-based workbook position.
    pub fn position(&self) -> usize {
        self.data.position
    }

    /// Semantic sheet kind resolved from its relationship.
    pub fn kind(&self) -> WorksheetKind {
        self.data.kind
    }

    pub(crate) fn part_uri(&self) -> &PackURI {
        &self.data.part_uri
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

    /// Borrow this worksheet's Office Add-in range bindings.
    ///
    /// Bindings are decoded and validated on first access, then retained by
    /// the immutable sheet snapshot. Non-worksheet handles return a typed
    /// error instead of attempting to interpret another sheet kind as XML.
    pub fn web_bindings(&self) -> Result<&crate::web::Bindings> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        self.data.web_bindings.get_or_try_init(|| {
            let part = self.owner.package.get_part(&self.data.part_uri)?;
            raw::web::read(part.blob())
        })
    }

    /// Parse the worksheet's direct auto-filter declaration.
    pub fn auto_filter(&self) -> Result<Option<crate::auto_filter::Definition>> {
        worksheet::auto_filter(self)
    }

    /// Parse all conditional-formatting containers and associate their
    /// differential formats with the workbook style resource.
    pub fn conditional_formattings(
        &self,
    ) -> Result<Vec<crate::conditional_formatting::Formatting>> {
        worksheet::conditional_formattings(self)
    }

    /// Parse worksheet data-validation collections.
    pub fn data_validations(&self) -> Result<Vec<crate::data_validation::Collection>> {
        worksheet::data_validations(self)
    }

    /// Parse the worksheet's optional data-consolidation declaration.
    pub fn data_consolidation(
        &self,
    ) -> Result<Option<crate::data_consolidation::DataConsolidation>> {
        worksheet::data_consolidation(self)
    }

    /// Parse core worksheet header/footer settings.
    pub fn header_footer(&self) -> Result<Option<crate::header_footer::Settings>> {
        worksheet::header_footer(self)
    }

    /// Parse ignored-error declarations.
    pub fn ignored_errors(&self) -> Result<Option<crate::ignored_errors::IgnoredErrors>> {
        worksheet::ignored_errors(self)
    }

    /// Parse this worksheet's inert smart-tag annotations.
    ///
    /// The metadata is never resolved to an action provider or executed.
    pub fn smart_tags(&self) -> Result<Option<crate::smart_tags::Collection>> {
        worksheet::smart_tags(self)
    }

    /// Parse the optional named-sheet-view relationship.
    pub fn named_sheet_views(&self) -> Result<Option<crate::named_sheet_view::Views>> {
        worksheet::named_sheet_views(self)
    }

    /// Parse worksheet outline properties.
    pub fn outline_properties(
        &self,
    ) -> Result<Option<crate::outline_properties::OutlineProperties>> {
        worksheet::outline_properties(self)
    }

    /// Parse worksheet page margins.
    pub fn page_margins(&self) -> Result<Option<crate::page_margins::Margins>> {
        worksheet::page_margins(self)
    }

    /// Parse worksheet page setup.
    pub fn page_setup(&self) -> Result<Option<crate::page_setup::Setup>> {
        worksheet::page_setup(self)
    }

    /// Parse worksheet phonetic properties.
    pub fn phonetic_properties(
        &self,
    ) -> Result<Option<crate::phonetic_properties::PhoneticProperties>> {
        worksheet::phonetic_properties(self)
    }

    /// Parse worksheet print options.
    pub fn print_options(&self) -> Result<Option<crate::print_options::PrintOptions>> {
        worksheet::print_options(self)
    }

    /// Parse worksheet what-if scenarios.
    pub fn scenarios(&self) -> Result<Option<crate::scenarios::Collection>> {
        worksheet::scenarios(self)
    }

    /// Parse worksheet-level calculation properties.
    pub fn calculation_properties(
        &self,
    ) -> Result<Option<crate::sheet_calculation_properties::Properties>> {
        worksheet::calculation_properties(self)
    }

    /// Parse the complete worksheet protection projection.
    pub fn protection(&self) -> Result<crate::sheet_protection::Metadata> {
        worksheet::protection(self)
    }

    /// Parse the worksheet's ordinary sheet-view collection.
    pub fn views(&self) -> Result<Option<crate::sheet_view::Collection>> {
        worksheet::views(self)
    }

    /// Parse every table relationship owned by this worksheet.
    pub fn tables(&self) -> Result<Vec<worksheet::TablePart>> {
        worksheet::tables(self)
    }

    /// Load and validate every query-table part owned by this worksheet.
    pub fn query_tables(&self) -> Result<Vec<crate::query_table::Part>> {
        worksheet::query_tables(self)
    }

    /// Load this worksheet's inert ActiveX graph.
    pub fn active_x(&self) -> Result<crate::active_x::ControlSet> {
        worksheet::active_x(self)
    }

    /// Load timeline views associated with this worksheet.
    pub fn timelines(&self) -> Result<Vec<crate::timeline::Part>> {
        worksheet::timelines(self)
    }

    /// Return array-formula anchors without materializing a dense worksheet.
    pub fn array_formulas(&self) -> Result<Vec<worksheet::ArrayFormula>> {
        worksheet::array_formulas(self)
    }

    /// Look up one exact logical cell state by A1 or checked coordinate.
    ///
    /// [`crate::cell::View::Missing`], [`crate::cell::View::Covered`], and a
    /// producer-stored [`crate::Cell`] are distinct without materializing follower
    /// cells for large merged ranges.
    pub fn cell<'a>(&self, at: impl Into<At<'a>>) -> Result<crate::cell::View<'_>> {
        let address = at.into().resolve()?;
        Ok(self.store()?.view(address))
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

    /// Lazily traverse validated, non-overlapping merged ranges.
    pub fn merges(&self) -> Result<crate::merge::Merges<'_>> {
        Ok(self.store()?.merges())
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

    pub(super) fn store(&self) -> Result<&Store> {
        if self.data.kind != WorksheetKind::Worksheet {
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
    pub(super) fn shared_strings(&self) -> Result<Option<&[Text]>> {
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

    pub(super) fn validate_styles(&self, store: &Store) -> Result<()> {
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
            if data.kind != WorksheetKind::Worksheet {
                continue;
            }
            let sheet = Worksheet {
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
