//! Source-bound flat `OpenDocument` Spreadsheet (`.fods`) support.

use crate::{Cell, CellValue, CellView, Row, Sheet};
use litchi_core::{Error, Metadata, Result};
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
use std::sync::Arc;

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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    output_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_XML_BYTES,
            output_bytes: MAX_XML_BYTES,
        }
    }
}

impl Limits {
    #[must_use]
    pub fn with_input_bytes(mut self, input_bytes: usize) -> Self {
        self.input_bytes = input_bytes;
        self
    }

    #[must_use]
    pub fn with_output_bytes(mut self, output_bytes: usize) -> Self {
        self.output_bytes = output_bytes;
        self
    }

    fn validate(self) -> Result<Self> {
        if self.input_bytes > MAX_XML_BYTES || self.output_bytes > MAX_XML_BYTES {
            return Err(invalid("flat ODS limits exceed hard safety ceilings"));
        }
        Ok(self)
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
    source: Arc<str>,
    sheets: Arc<[Sheet]>,
    metadata: Metadata,
    odf_metadata: OdfMetadata,
}

/// Immutable, Arc-backed, lossless flat-spreadsheet snapshot.
#[derive(Clone, Debug)]
pub struct Snapshot {
    state: Arc<State>,
}

pub type FlatSpreadsheet = Snapshot;

impl Snapshot {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with(bytes, Limits::default())
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_bytes_with(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        if bytes.len() > limits.input_bytes {
            return Err(invalid("flat ODS XML exceeds the input byte limit"));
        }
        let source = String::from_utf8(bytes)
            .map_err(|_error| invalid("flat ODS XML is not valid UTF-8"))?;
        Self::from_source(Arc::from(source), limits)
    }

    fn from_source(source: Arc<str>, limits: Limits) -> Result<Self> {
        if source.len() > limits.input_bytes {
            return Err(invalid("flat ODS XML exceeds the input byte limit"));
        }
        validate_flat_xml(&source)?;
        let sheets = crate::worksheet::codec::parse_flat(&source)?;
        crate::dde::Snapshot::parse(&source).map_err(|error| {
            Error::InvalidFormat(format!("flat ODS DDE metadata inspection failed: {error}"))
        })?;
        crate::scenario::Snapshot::parse(&source).map_err(|error| {
            Error::InvalidFormat(format!(
                "flat ODS scenario metadata inspection failed: {error}"
            ))
        })?;
        let odf_metadata = OdfMetadata::from_xml(&source)?;
        let metadata = odf_metadata.clone().into();
        Ok(Self {
            state: Arc::new(State {
                source,
                sheets: Arc::from(sheets),
                metadata,
                odf_metadata,
            }),
        })
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.state.source.as_bytes()
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
    ) -> Result<Option<&'s Sheet>> {
        select(&self.state.sheets, selector.into())
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn cell<'q>(
        &self,
        selector: impl Into<SheetSelector<'q>>,
        row: usize,
        column: usize,
    ) -> Result<Option<CellView<'_>>> {
        Ok(self
            .sheet(selector)?
            .map(|sheet| sheet.cell_view(row, column)))
    }

    #[must_use]
    pub fn metadata(&self) -> &Metadata {
        &self.state.metadata
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
    pub fn dde(&self) -> Result<crate::dde::Snapshot> {
        crate::dde::Snapshot::parse(&self.state.source).map_err(|error| {
            Error::InvalidFormat(format!("flat ODS DDE metadata inspection failed: {error}"))
        })
    }

    /// Inspects scenario declarations without applying their values.
    ///
    /// # Errors
    ///
    /// Returns an error when the source XML has invalid or over-budget
    /// scenario metadata.
    pub fn scenarios(&self) -> Result<crate::scenario::Snapshot> {
        crate::scenario::Snapshot::parse(&self.state.source).map_err(|error| {
            Error::InvalidFormat(format!(
                "flat ODS scenario metadata inspection failed: {error}"
            ))
        })
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn transaction(&self) -> Result<Transaction> {
        self.transaction_with(Limits::default())
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn transaction_with(&self, limits: Limits) -> Result<Transaction> {
        let limits = limits.validate()?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(self.state.sheets.len())
            .map_err(|_error| invalid("flat ODS transaction allocation failed"))?;
        staged.resize_with(self.state.sheets.len(), || None);
        Ok(Transaction {
            base: self.clone(),
            staged,
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

/// Detached failure-atomic transaction bound to an immutable source snapshot.
pub struct Transaction {
    base: Snapshot,
    staged: Vec<Option<Sheet>>,
    limits: Limits,
}

pub type FlatEdit = Transaction;

impl Transaction {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell<'q>(
        &mut self,
        selector: impl Into<SheetSelector<'q>>,
        row: usize,
        column: usize,
        cell: Cell,
    ) -> Result<Option<()>> {
        let index = match selector.into() {
            SheetSelector::Index(index) => {
                if index >= self.staged.len() {
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
                    return Err(invalid(format!(
                        "flat ODS sheet selector '{name}' is ambiguous"
                    )));
                }
                index
            },
        };
        if self.staged[index].is_none() {
            self.staged[index] = Some(clone_sheet_fallible(&self.base.sheets()[index])?);
        }
        let staged_sheet = self.staged[index]
            .as_mut()
            .ok_or_else(|| invalid("flat ODS staged sheet initialization failed"))?;
        staged_sheet.set_cell(row, column, cell)?;
        Ok(Some(()))
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        if self.staged.iter().all(Option::is_none) {
            let patch = Patch {
                before: self.base.state.source.clone(),
                after: self.base.state.source.clone(),
            };
            return Ok(Commit {
                snapshot: self.base,
                patch,
            });
        }
        ensure_compact(&self.base.state.source, self.limits.output_bytes)?;
        let staged = self.staged.iter().map(Option::as_ref).collect::<Vec<_>>();
        crate::worksheet::package::validate_owned_link_delta_partial(self.base.sheets(), &staged)?;
        let xml = crate::worksheet::package::replace_changed_rows(
            &self.base.state.source,
            self.base.sheets(),
            &staged,
            self.limits.output_bytes,
        )?;
        ensure_compact(&xml, self.limits.output_bytes)?;
        let after: Arc<str> = Arc::from(xml);
        let snapshot = Snapshot::from_source(after.clone(), self.limits)?;
        for (index, staged) in self.staged.iter().enumerate() {
            if staged
                .as_ref()
                .is_some_and(|expected| snapshot.sheets().get(index) != Some(expected))
            {
                return Err(invalid(
                    "flat ODS typed readback does not match the staged edit",
                ));
            }
        }
        let patch = Patch {
            before: self.base.state.source.clone(),
            after,
        };
        Ok(Commit { snapshot, patch })
    }
}

#[derive(Clone, Debug)]
pub struct Patch {
    before: Arc<str>,
    after: Arc<str>,
}

impl Patch {
    #[must_use]
    pub fn changed(&self) -> bool {
        self.before != self.after
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Snapshot> {
        self.apply_with(snapshot, Limits::default())
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply_with(&self, snapshot: &Snapshot, limits: Limits) -> Result<Snapshot> {
        let limits = limits.validate()?;
        if snapshot.state.source.as_ref() != self.before.as_ref() {
            return Err(invalid("flat ODS patch source snapshot is stale"));
        }
        if self.after.len() > limits.output_bytes || self.after.len() > limits.input_bytes {
            return Err(invalid("flat ODS patch output exceeds the byte limit"));
        }
        ensure_compact_or_unchanged(&self.before, &self.after, limits.output_bytes)?;
        Snapshot::from_source(self.after.clone(), limits)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

pub type FlatCommit = Commit;

impl Commit {
    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.changed()
    }

    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn spreadsheet(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    #[must_use]
    pub fn into_spreadsheet(self) -> Snapshot {
        self.snapshot
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
    let limits = CompactLimits::new(max_bytes, MAX_XML_DEPTH).map_err(Error::from)?;
    compact_xml::validate_with_limits(xml.as_bytes(), limits).map_err(Error::from)
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
                let text: &[u8] = text.as_ref();
                if root_closed {
                    return Err(invalid("flat ODS XML has trailing text after its root"));
                }
                if !root_seen && text.iter().any(|byte| !byte.is_ascii_whitespace()) {
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
            Event::PI(_) | Event::Comment(_) if root_seen && !root_closed => {},
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {
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

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
