//! Source-bound flat `OpenDocument` Spreadsheet (`.fods`) support.

use crate::{Cell, CellValue, CellView, Row, Sheet};
use litchi_core::{Error as CoreError, Metadata, Result};
use litchi_odf_common::{SemanticLimitKind, SemanticLimits};
use litchi_odf_common::{
    compact_xml::{self, Limits as CompactLimits},
    core::metadata::Metadata as OdfMetadata,
};
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    collections::{BTreeMap, btree_map::Entry},
    io::Read,
    path::Path,
    sync::Arc,
};

/// Whether a commit changed the workbook bytes.
pub use litchi_core::edit::Diagnostics;
/// The family error: typed limits and refusals plus the common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;
pub use litchi_odf_common::OpenOptions;

type FamilyResult<T> = std::result::Result<T, Error>;

/// A committed flat edit: the published workbook, its patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Workbook, Patch>;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const CALCEXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 1_024;
const MAX_ELEMENTS: usize = 1_048_576;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 64 * 1024;
const MAX_RULES_PER_FORMAT: usize = 1_024;

/// The bounds one flat workbook is read and edited under, resolved from its
/// [`OpenOptions`]: the source `input_bytes` and the semantic `content_bytes`
/// bound the XML, `output_bytes` one published document.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Limits {
    input_bytes: usize,
    output_bytes: usize,
    semantic: SemanticLimits,
    options: OpenOptions,
}

impl Default for Limits {
    fn default() -> Self {
        Self::new(OpenOptions::new())
    }
}

impl Limits {
    fn new(options: OpenOptions) -> Self {
        let semantic = options.semantic_limits_value();
        let input =
            usize::try_from(options.source_limits_value().max_input_bytes()).unwrap_or(usize::MAX);
        Self {
            input_bytes: input
                .min(semantic.usize_value(SemanticLimitKind::ContentBytes))
                .min(MAX_XML_BYTES),
            output_bytes: semantic
                .usize_value(SemanticLimitKind::OutputBytes)
                .min(MAX_XML_BYTES),
            semantic,
            options,
        }
    }

    fn input_exceeded(self, observed: usize) -> CoreError {
        if observed > self.semantic.usize_value(SemanticLimitKind::ContentBytes) {
            return self
                .semantic
                .check_core(SemanticLimitKind::ContentBytes, observed)
                .err()
                .unwrap_or_else(|| invalid("flat ODS limit accounting failed"));
        }
        CoreError::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::InputBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(self.input_bytes).unwrap_or(u64::MAX),
            scope: "ODF owned package input".into(),
        })
    }

    fn output_exceeded(self, observed: usize) -> CoreError {
        self.semantic
            .check_core(SemanticLimitKind::OutputBytes, observed)
            .err()
            .unwrap_or_else(|| invalid("flat ODS output exceeds the byte limit"))
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SheetSelector<'a> {
    Name(&'a str),
    Index(usize),
}

impl<'a> From<&'a str> for SheetSelector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for SheetSelector<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

#[derive(Debug)]
struct State {
    limits: Limits,
    source: Arc<str>,
    sheets: Arc<[Sheet]>,
    metadata: Metadata,
    odf_metadata: OdfMetadata,
    tolerance: litchi_odf_common::leniency::ToleranceReport,
}

/// The immutable, Arc-backed, lossless owner of one flat spreadsheet
/// (`.fods`).
///
/// Cheap to clone; a commit publishes a new workbook and never changes its
/// source. The workbook keeps the [`OpenOptions`] it was opened with; its
/// edits and replays use them.
#[derive(Clone, Debug)]
pub struct Workbook {
    state: Arc<State>,
}

impl Workbook {
    /// Opens a flat spreadsheet from a path with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, Error> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Opens a flat spreadsheet from a path under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_options(file, options)
    }

    /// Reads a flat spreadsheet from a stream with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn from_reader(reader: impl Read) -> std::result::Result<Self, Error> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Reads a flat spreadsheet from a stream under caller-selected limits;
    /// the stream is read through an `input_bytes + 1` guard.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn from_reader_with_options(
        reader: impl Read,
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        let limits = Limits::new(options);
        let mut bytes = Vec::new();
        let guard = u64::try_from(limits.input_bytes)
            .unwrap_or(u64::MAX)
            .saturating_add(1);
        reader.take(guard).read_to_end(&mut bytes)?;
        Ok(Self::from_vec(bytes, limits)?)
    }

    /// Opens a flat spreadsheet with the default [`OpenOptions`]; the bytes
    /// are copied.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`].
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, Error> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Opens a flat spreadsheet under caller-selected limits; the bytes are
    /// copied after the size check.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] when a selected dimension is exceeded and
    /// [`Error::Core`] when the XML is not a valid flat spreadsheet.
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        let limits = Limits::new(options);
        if bytes.len() > limits.input_bytes {
            return Err(limits.input_exceeded(bytes.len()).into());
        }
        Ok(Self::from_vec(bytes.to_vec(), limits)?)
    }

    fn from_vec(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        if bytes.len() > limits.input_bytes {
            return Err(limits.input_exceeded(bytes.len()));
        }
        let source = String::from_utf8(bytes)
            .map_err(|_error| invalid("flat ODS XML is not valid UTF-8"))?;
        Self::from_source(Arc::from(source), limits)
    }

    /// The options this workbook was opened with.
    pub fn options(&self) -> OpenOptions {
        self.state.limits.options
    }

    /// Replays a patch onto this exact source workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this workbook is not the
    /// patch's exact source.
    pub fn apply(&self, patch: &Patch) -> std::result::Result<Commit, Error> {
        if self.state.source.as_ref() != patch.before.as_ref() {
            return Err(Error::patch_conflict(
                "flat.apply",
                "the workbook is not the patch source",
            ));
        }
        let limits = self.state.limits;
        if patch.after.len() > limits.output_bytes {
            return Err(limits.output_exceeded(patch.after.len()).into());
        }
        ensure_compact_or_unchanged(&patch.before, &patch.after, limits.output_bytes)?;
        let target = Self::from_source(patch.after.clone(), limits)?;
        Ok(Commit::new(
            target,
            patch.clone(),
            Diagnostics::new(patch.before != patch.after),
        ))
    }

    fn from_source(source: Arc<str>, limits: Limits) -> Result<Self> {
        if source.len() > limits.input_bytes {
            return Err(limits.input_exceeded(source.len()));
        }
        validate_flat_xml(&source)?;
        let mut tolerance = litchi_odf_common::leniency::ToleranceReport::default();
        let sheets = crate::worksheet::codec::parse_flat(&source, &mut tolerance)?;
        crate::dde::Snapshot::parse(&source).map_err(|error| {
            CoreError::InvalidFormat(format!("flat ODS DDE metadata inspection failed: {error}"))
        })?;
        crate::scenario::Snapshot::parse(&source).map_err(|error| {
            CoreError::InvalidFormat(format!(
                "flat ODS scenario metadata inspection failed: {error}"
            ))
        })?;
        let odf_metadata = OdfMetadata::from_xml(&source, limits.options.xml_limits_value())?;
        let metadata = odf_metadata.clone().into();
        Ok(Self {
            state: Arc::new(State {
                limits,
                source,
                sheets: Arc::from(sheets),
                metadata,
                odf_metadata,
                tolerance,
            }),
        })
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.state.source.as_bytes()
    }

    /// Everything a tolerant read repaired while this snapshot was parsed.
    ///
    /// An empty report means the document conformed; see
    /// [`litchi_odf_common::leniency::ToleranceReport`].
    #[inline]
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        &self.state.tolerance
    }

    /// Share the exact source owner without copying its bytes.
    #[must_use]
    pub fn to_bytes(&self) -> Arc<str> {
        self.state.source.clone()
    }

    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.state.sheets
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn sheet<'s, 'q>(
        &'s self,
        selector: impl Into<SheetSelector<'q>>,
    ) -> FamilyResult<Option<&'s Sheet>> {
        Ok(select(&self.state.sheets, selector.into())?)
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn cell<'q>(
        &self,
        selector: impl Into<SheetSelector<'q>>,
        row: usize,
        column: usize,
    ) -> FamilyResult<Option<CellView<'_>>> {
        Ok(self
            .sheet(selector)?
            .map(|sheet| sheet.cell_view(row, column)))
    }

    #[must_use]
    pub fn metadata(&self) -> &Metadata {
        &self.state.metadata
    }

    /// Extract displayed worksheet text with the packaged projection's
    /// layout: tab-separated cells, newline-separated rows, sheet order.
    ///
    /// # Errors
    /// Returns an error when the bounded projection cannot reserve its output.
    pub fn text(&self) -> FamilyResult<String> {
        Ok(crate::facade::source::project_text(&self.state.sheets)?)
    }

    #[must_use]
    pub fn odf_metadata(&self) -> &OdfMetadata {
        &self.state.odf_metadata
    }

    /// Inspects inert DDE metadata retained by this flat spreadsheet.
    ///
    /// # Errors
    ///
    /// Returns an error when the source XML has invalid or over-budget DDE
    /// metadata.
    pub fn dde(&self) -> FamilyResult<crate::dde::Snapshot> {
        Ok(
            crate::dde::Snapshot::parse(&self.state.source).map_err(|error| {
                CoreError::InvalidFormat(format!(
                    "flat ODS DDE metadata inspection failed: {error}"
                ))
            })?,
        )
    }

    /// Inspects scenario declarations without applying their values.
    ///
    /// # Errors
    ///
    /// Returns an error when the source XML has invalid or over-budget
    /// scenario metadata.
    pub fn scenarios(&self) -> FamilyResult<crate::scenario::Snapshot> {
        Ok(
            crate::scenario::Snapshot::parse(&self.state.source).map_err(|error| {
                CoreError::InvalidFormat(format!(
                    "flat ODS scenario metadata inspection failed: {error}"
                ))
            })?,
        )
    }

    /// Starts a detached, failure-atomic edit bound to this workbook.
    ///
    /// # Errors
    /// Returns an error when the staging buffer cannot be reserved.
    pub fn edit(&self) -> FamilyResult<Edit> {
        Ok(self.edit_core()?)
    }

    fn edit_core(&self) -> Result<Edit> {
        let limits = self.state.limits;
        Ok(Edit {
            base: self.clone(),
            staged: BTreeMap::new(),
            limits,
        })
    }
}

fn select<'s>(sheets: &'s [Sheet], selector: SheetSelector<'_>) -> Result<Option<&'s Sheet>> {
    match selector {
        SheetSelector::Index(index) => Ok(sheets.get(index)),
        SheetSelector::Name(name) => {
            let mut matches = sheets.iter().filter(|sheet| sheet.name == name);
            let first = matches.next();
            if first.is_some() && matches.next().is_some() {
                return Err(invalid(format!(
                    "flat ODS sheet selector '{name}' is ambiguous"
                )));
            }
            Ok(first)
        },
    }
}

/// A detached, failure-atomic edit bound to an immutable source workbook.
pub struct Edit {
    base: Workbook,
    staged: BTreeMap<usize, Sheet>,
    limits: Limits,
}

impl Edit {
    /// The source workbook this edit stages against.
    #[must_use]
    pub fn package(&self) -> &Workbook {
        &self.base
    }

    /// Discards every staged change in place.
    pub fn reset(&mut self) {
        self.staged.clear();
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell<'q>(
        &mut self,
        selector: impl Into<SheetSelector<'q>>,
        row: usize,
        column: usize,
        cell: Cell,
    ) -> FamilyResult<Option<()>> {
        let index = match selector.into() {
            SheetSelector::Index(index) => {
                if index >= self.base.sheets().len() {
                    return Ok(None);
                }
                index
            },
            SheetSelector::Name(name) => {
                let mut matches = self
                    .base
                    .sheets()
                    .iter()
                    .enumerate()
                    .filter(|(_, sheet)| sheet.name == name);
                let Some(index) = matches.next().map(|(index, _)| index) else {
                    return Ok(None);
                };
                if matches.next().is_some() {
                    return Err(Error::from(invalid(format!(
                        "flat ODS sheet selector '{name}' is ambiguous"
                    ))));
                }
                index
            },
        };
        let staged_sheet = match self.staged.entry(index) {
            Entry::Occupied(entry) => entry.into_mut(),
            Entry::Vacant(entry) => entry.insert(clone_sheet_fallible(&self.base.sheets()[index])?),
        };
        staged_sheet.set_cell(row, column, cell)?;
        Ok(Some(()))
    }

    /// Validates the staged cells and publishes one immutable workbook.
    ///
    /// An edit with no staged change commits the source itself:
    /// `diagnostics().changed()` is false and `patch().is_noop()` holds.
    ///
    /// # Errors
    /// Returns an error when the edit cannot be published losslessly within
    /// the workbook's limits.
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        Ok(self.commit_core()?)
    }

    fn commit_core(self) -> Result<Commit> {
        if self.staged.is_empty() {
            let patch = Patch {
                before: self.base.state.source.clone(),
                after: self.base.state.source.clone(),
            };
            return Ok(Commit::new(self.base, patch, Diagnostics::UNCHANGED));
        }
        ensure_compact(&self.base.state.source, self.limits.output_bytes)?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(self.base.sheets().len())
            .map_err(|_error| invalid("flat ODS transaction allocation failed"))?;
        staged.resize(self.base.sheets().len(), None);
        for (index, sheet) in &self.staged {
            staged[*index] = Some(sheet);
        }
        crate::worksheet::package::validate_owned_link_delta_partial(self.base.sheets(), &staged)?;
        let xml = crate::worksheet::package::replace_changed_rows(
            &self.base.state.source,
            self.base.sheets(),
            &staged,
            self.limits.output_bytes,
        )?;
        ensure_compact(&xml, self.limits.output_bytes)?;
        let after: Arc<str> = Arc::from(xml);
        let snapshot = Workbook::from_source(after.clone(), self.limits)?;
        for (index, expected) in &self.staged {
            if snapshot.sheets().get(*index) != Some(expected) {
                return Err(invalid(
                    "flat ODS typed readback does not match the staged edit",
                ));
            }
        }
        let changed = self.base.state.source != after;
        let patch = Patch {
            before: self.base.state.source.clone(),
            after,
        };
        Ok(Commit::new(snapshot, patch, Diagnostics::new(changed)))
    }
}

/// An exact-source, reversible flat workbook patch.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Arc<str>,
    after: Arc<str>,
}

impl Patch {
    /// Whether replaying this patch keeps the source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

fn ensure_compact_or_unchanged(before: &str, after: &str, max_bytes: usize) -> Result<()> {
    if before != after {
        ensure_compact(after, max_bytes)?;
    }
    Ok(())
}

fn ensure_compact(xml: &str, max_bytes: usize) -> Result<()> {
    let max_bytes = max_bytes.min(compact_xml::HARD_MAX_BYTES);
    let limits = CompactLimits::new(max_bytes, MAX_XML_DEPTH).map_err(CoreError::from)?;
    compact_xml::validate(xml.as_bytes(), limits).map_err(CoreError::from)
}

fn clone_string_fallible(value: &str) -> Result<String> {
    let mut cloned = String::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|_error| invalid("flat ODS string clone allocation failed"))?;
    cloned.push_str(value);
    Ok(cloned)
}

fn clone_option_string_fallible(value: Option<&String>) -> Result<Option<String>> {
    value.map(|value| clone_string_fallible(value)).transpose()
}

fn clone_value_fallible(value: &CellValue) -> Result<CellValue> {
    Ok(match value {
        CellValue::Empty => CellValue::Empty,
        CellValue::Text(value) => CellValue::Text(clone_string_fallible(value)?),
        CellValue::Number(value) => CellValue::Number(*value),
        CellValue::Currency { value, currency } => CellValue::Currency {
            value: *value,
            currency: clone_string_fallible(currency)?,
        },
        CellValue::Percentage(value) => CellValue::Percentage(*value),
        CellValue::Boolean(value) => CellValue::Boolean(*value),
        CellValue::Date(value) => CellValue::Date(clone_string_fallible(value)?),
        CellValue::Time(value) => CellValue::Time(clone_string_fallible(value)?),
        CellValue::Unknown { kind, value } => CellValue::Unknown {
            kind: clone_string_fallible(kind)?,
            value: clone_option_string_fallible(value.as_ref())?,
        },
    })
}

fn clone_cell_fallible(cell: &Cell) -> Result<Cell> {
    let mut cloned = Cell::repeated(
        clone_value_fallible(&cell.value)?,
        clone_string_fallible(&cell.text)?,
        cell.repeat(),
    )?;
    cloned.formula = clone_option_string_fallible(cell.formula.as_ref())?;
    cloned.style_name = clone_option_string_fallible(cell.style_name.as_ref())?;
    cloned.merge = cell.merge;
    Ok(cloned)
}

fn clone_row_fallible(row: &Row) -> Result<Row> {
    let mut cloned = Row::repeated(row.repeat())?;
    cloned.style_name = clone_option_string_fallible(row.style_name.as_ref())?;
    cloned.default_cell_style_name =
        clone_option_string_fallible(row.default_cell_style_name.as_ref())?;
    cloned
        .cells
        .try_reserve_exact(row.cells.len())
        .map_err(|_error| invalid("flat ODS row clone allocation failed"))?;
    for cell in &row.cells {
        cloned.cells.push(clone_cell_fallible(cell)?);
    }
    Ok(cloned)
}

fn clone_sheet_fallible(sheet: &Sheet) -> Result<Sheet> {
    let mut cloned = Sheet::new(clone_string_fallible(&sheet.name)?)?;
    cloned.style_name = clone_option_string_fallible(sheet.style_name.as_ref())?;
    cloned
        .rows
        .try_reserve_exact(sheet.rows.len())
        .map_err(|_error| invalid("flat ODS sheet clone allocation failed"))?;
    for row in &sheet.rows {
        cloned.rows.push(clone_row_fallible(row)?);
    }
    Ok(cloned)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum FlatNamespace {
    Office,
    Calcext,
    Other,
}

fn validate_flat_xml(xml: &str) -> Result<()> {
    if xml.is_empty() || xml.as_bytes().starts_with(b"PK\x03\x04") {
        return Err(invalid("flat ODS input is not XML"));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut elements = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_depth = None;
    let mut body_seen = false;
    let mut spreadsheet_seen = false;
    let mut conditional_depth = None;
    let mut conditional_rules = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid flat ODS XML: {error}")))?;
        let namespace = flat_namespace(&namespace);
        match event {
            Event::Start(element) => {
                elements = elements
                    .checked_add(1)
                    .ok_or_else(|| invalid("flat ODS element count overflow"))?;
                if elements > MAX_ELEMENTS || depth == MAX_XML_DEPTH {
                    return Err(invalid("flat ODS XML exceeds a structural limit"));
                }
                validate_attributes(&reader, &element)?;
                if depth == 0 {
                    if root_seen || root_closed || !office(namespace, &element, b"document") {
                        return Err(invalid("flat ODS root must be one office:document"));
                    }
                    root_seen = true;
                    if office_attribute(&reader, &element, b"mimetype")?.as_deref()
                        != Some(MIMETYPE)
                    {
                        return Err(invalid("flat ODS has the wrong document family"));
                    }
                } else if office(namespace, &element, b"body") {
                    if depth != 1 || body_seen {
                        return Err(invalid("office:body is misplaced or duplicated"));
                    }
                    body_seen = true;
                    body_depth = Some(depth + 1);
                } else if office(namespace, &element, b"spreadsheet") {
                    if body_depth != Some(depth) || spreadsheet_seen {
                        return Err(invalid("office:spreadsheet is misplaced or duplicated"));
                    }
                    spreadsheet_seen = true;
                }
                if calcext(
                    namespace,
                    element.local_name().as_ref(),
                    b"conditional-format",
                ) {
                    conditional_depth = Some(depth + 1);
                    conditional_rules = 0;
                } else if conditional_depth.is_some()
                    && calcext(namespace, element.local_name().as_ref(), b"condition")
                {
                    conditional_rules += 1;
                    if conditional_rules > MAX_RULES_PER_FORMAT {
                        return Err(invalid("flat ODS conditional rule limit exceeded"));
                    }
                }
                depth += 1;
            },
            Event::Empty(element) => {
                elements = elements
                    .checked_add(1)
                    .ok_or_else(|| invalid("flat ODS element count overflow"))?;
                if elements > MAX_ELEMENTS || depth == 0 {
                    return Err(invalid("flat ODS XML exceeds a structural limit"));
                }
                validate_attributes(&reader, &element)?;
                if office(namespace, &element, b"body")
                    || office(namespace, &element, b"spreadsheet")
                {
                    return Err(invalid("flat ODS body must contain a spreadsheet"));
                }
                if conditional_depth.is_some()
                    && calcext(namespace, element.local_name().as_ref(), b"condition")
                {
                    conditional_rules += 1;
                    if conditional_rules > MAX_RULES_PER_FORMAT {
                        return Err(invalid("flat ODS conditional rule limit exceeded"));
                    }
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("flat ODS XML element stack underflow"))?;
                if calcext(
                    namespace,
                    element.local_name().as_ref(),
                    b"conditional-format",
                ) {
                    conditional_depth = None;
                }
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::DocType(_) => {
                return Err(invalid("flat ODS XML document types are not accepted"));
            },
            Event::GeneralRef(reference) => {
                let reference: &[u8] = reference.as_ref();
                if !root_seen
                    || root_closed
                    || !matches!(reference, b"amp" | b"lt" | b"gt" | b"apos" | b"quot")
                {
                    return Err(invalid("flat ODS XML contains an unsupported entity"));
                }
            },
            Event::Text(text) => {
                // XML `Misc` around the root admits whitespace only; every
                // producer ends a flat document with a newline.
                let text: &[u8] = text.as_ref();
                let blank = text.iter().all(|byte| byte.is_ascii_whitespace());
                if root_closed && !blank {
                    return Err(invalid("flat ODS XML has trailing text after its root"));
                }
                if !root_seen && !blank {
                    return Err(invalid("flat ODS XML has text before its root"));
                }
            },
            Event::CData(_) => {
                if !root_seen || root_closed {
                    return Err(invalid("flat ODS XML has CDATA outside its root"));
                }
            },
            Event::Eof => break,
            Event::Decl(_) if !root_seen && !root_closed => {},
            // Comments and processing instructions are XML `Misc`, admitted
            // before and after the root as inside it; both stay inert.
            Event::PI(_) | Event::Comment(_) => {},
            Event::Decl(_) => {
                return Err(invalid("flat ODS XML has markup outside its root"));
            },
        }
        buffer.clear();
    }

    if depth != 0 || !root_seen || !root_closed || !body_seen || !spreadsheet_seen {
        return Err(invalid("flat ODS XML is structurally incomplete"));
    }
    Ok(())
}

fn validate_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes().with_checks(true) {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("flat ODS attribute count overflow"))?;
        if count > MAX_ATTRIBUTES {
            return Err(invalid(
                "flat ODS element exceeds the attribute count limit",
            ));
        }
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid flat ODS attribute: {error}")))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid flat ODS attribute value: {error}")))?;
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(invalid("flat ODS attribute exceeds the byte limit"));
        }
    }
    Ok(())
}

fn office(
    namespace: FlatNamespace,
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
) -> bool {
    namespace == FlatNamespace::Office && element.local_name().as_ref() == local
}

fn calcext(namespace: FlatNamespace, element_local: &[u8], local: &[u8]) -> bool {
    namespace == FlatNamespace::Calcext && element_local == local
}

fn flat_namespace(namespace: &ResolveResult<'_>) -> FlatNamespace {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE => FlatNamespace::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == CALCEXT_NAMESPACE => FlatNamespace::Calcext,
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            FlatNamespace::Other
        },
    }
}

fn office_attribute(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid flat ODS attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
            && name.as_ref() == local
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| invalid(format!("invalid flat ODS attribute value: {error}")));
        }
    }
    Ok(None)
}

fn invalid(message: impl Into<String>) -> CoreError {
    CoreError::InvalidFormat(message.into())
}
