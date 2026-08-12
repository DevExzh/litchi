//! Bounded, source-backed inspection and editing of ODS content-validation ownership.
//!
//! The inventory covers the document-level validation catalog and the physical
//! cell runs that bind catalog names. Edits are clone-staged, preserve source
//! bytes outside checked owner ranges, and refuse opaque owners or operations
//! that would break the binding closure.

use core::fmt;
use litchi_odf_common::{
    constants,
    core::{AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication},
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    collections::{HashMap, HashSet},
    ops::Range,
};

struct AttributeNamespaceIds {
    namespace_ids: HashMap<Vec<u8>, usize>,
    prefix_ids: HashMap<Vec<u8>, usize>,
    #[cfg(test)]
    namespace_content_lookups: usize,
}

impl AttributeNamespaceIds {
    fn with_capacity(estimate: usize) -> Result<Self> {
        let capacity = estimate.min(64);
        let mut namespace_ids = HashMap::new();
        namespace_ids
            .try_reserve(capacity)
            .map_err(|_error| Error::AllocationFailed("expanded attribute namespaces"))?;
        let mut prefix_ids = HashMap::new();
        prefix_ids
            .try_reserve(capacity)
            .map_err(|_error| Error::AllocationFailed("attribute prefix index"))?;
        Ok(Self {
            namespace_ids,
            prefix_ids,
            #[cfg(test)]
            namespace_content_lookups: 0,
        })
    }

    fn id(&mut self, prefix: &[u8], namespace: &[u8]) -> Result<usize> {
        if let Some(id) = self.prefix_ids.get(prefix) {
            return Ok(*id);
        }
        #[cfg(test)]
        {
            self.namespace_content_lookups += 1;
        }
        let id = if let Some(id) = self.namespace_ids.get(namespace) {
            *id
        } else {
            self.namespace_ids
                .try_reserve(1)
                .map_err(|_error| Error::AllocationFailed("attribute namespace index"))?;
            let id = self.namespace_ids.len();
            self.namespace_ids
                .insert(copy_bytes(namespace, "attribute namespace")?, id);
            id
        };
        self.prefix_ids
            .try_reserve(1)
            .map_err(|_error| Error::AllocationFailed("attribute prefix index"))?;
        self.prefix_ids
            .insert(copy_bytes(prefix, "attribute prefix")?, id);
        Ok(id)
    }

    #[cfg(test)]
    fn namespace_content_lookups(&self) -> usize {
        self.namespace_content_lookups
    }
}

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const SCRIPT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

const MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_EVENTS: usize = 4_000_000;
const MAX_ATTRIBUTES: usize = 4_000_000;
const MAX_DEPTH: usize = 1_024;
const MAX_DEFINITIONS: usize = 65_536;
const MAX_SHEETS: usize = 65_536;
const MAX_BINDINGS: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_LOGICAL_ROWS: usize = 16_777_216;
const MAX_LOGICAL_COLUMNS: usize = 1_048_576;
const MAX_OPERATIONS: usize = 65_536;

/// A content-validation inspection result.
pub type Result<T> = std::result::Result<T, Error>;

/// The resource whose configured or observed value exceeded a limit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    Events,
    Attributes,
    Depth,
    Definitions,
    Sheets,
    Bindings,
    TextBytes,
    LogicalRows,
    LogicalColumns,
    OutputBytes,
    Operations,
}

/// Errors produced by source-backed content-validation inspection.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A configured or hard resource limit was exceeded.
    LimitExceeded {
        kind: LimitKind,
        observed: usize,
        maximum: usize,
    },
    /// An owned allocation could not be reserved.
    AllocationFailed(&'static str),
    /// The XML byte stream is malformed.
    InvalidXml(String),
    /// Validation ownership or placement is malformed or ambiguous.
    InvalidStructure(String),
    /// Markup Compatibility requires branch selection this reader does not perform.
    UnsupportedMarkupCompatibility,
    /// The requested definition name is already owned by the catalog.
    DuplicateDefinition(String),
    /// The requested definition is absent.
    DefinitionNotFound(String),
    /// A destructive edit would leave compact cell bindings dangling.
    DefinitionReferenced { name: String, bindings: usize },
    /// Renaming is unavailable because this editor does not rewrite bindings.
    UnsafeRename { from: String, to: String },
    /// The retained owner includes markup outside this editor's model.
    OpaqueOwner,
    /// A reversible patch was applied to a different exact XML source.
    SourceMismatch,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::LimitExceeded {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "ODS content-validation {kind:?} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::AllocationFailed(owner) => {
                write!(
                    formatter,
                    "could not allocate ODS content-validation {owner}"
                )
            },
            Self::InvalidXml(message) | Self::InvalidStructure(message) => {
                formatter.write_str(message)
            },
            Self::UnsupportedMarkupCompatibility => formatter.write_str(
                "ODS content-validation inspection does not select Markup Compatibility branches",
            ),
            Self::DuplicateDefinition(name) => {
                write!(formatter, "content-validation definition '{name}' already exists")
            },
            Self::DefinitionNotFound(name) => {
                write!(formatter, "content-validation definition '{name}' was not found")
            },
            Self::DefinitionReferenced { name, bindings } => write!(
                formatter,
                "content-validation definition '{name}' is retained by {bindings} compact binding(s)"
            ),
            Self::UnsafeRename { from, to } => write!(
                formatter,
                "content-validation rename from '{from}' to '{to}' requires an unavailable atomic binding rewrite"
            ),
            Self::OpaqueOwner => formatter.write_str(
                "ODS content-validation mutation refuses opaque catalog, definition, or bound-cell markup",
            ),
            Self::SourceMismatch => formatter.write_str(
                "ODS content-validation patch source snapshot does not match",
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Caller-selected resource limits, each capped by a fixed implementation ceiling.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    events: usize,
    attributes: usize,
    depth: usize,
    definitions: usize,
    sheets: usize,
    bindings: usize,
    text_bytes: usize,
    logical_rows: usize,
    logical_columns: usize,
    output_bytes: usize,
    operations: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            events: MAX_EVENTS,
            attributes: MAX_ATTRIBUTES,
            depth: MAX_DEPTH,
            definitions: MAX_DEFINITIONS,
            sheets: MAX_SHEETS,
            bindings: MAX_BINDINGS,
            text_bytes: MAX_TEXT_BYTES,
            logical_rows: MAX_LOGICAL_ROWS,
            logical_columns: MAX_LOGICAL_COLUMNS,
            output_bytes: MAX_INPUT_BYTES,
            operations: MAX_OPERATIONS,
        }
    }
}

macro_rules! limit_setter {
    ($name:ident, $field:ident) => {
        #[must_use]
        pub const fn $name(mut self, value: usize) -> Self {
            self.$field = value;
            self
        }
    };
}

impl Limits {
    limit_setter!(with_input_bytes, input_bytes);
    limit_setter!(with_events, events);
    limit_setter!(with_attributes, attributes);
    limit_setter!(with_depth, depth);
    limit_setter!(with_definitions, definitions);
    limit_setter!(with_sheets, sheets);
    limit_setter!(with_bindings, bindings);
    limit_setter!(with_text_bytes, text_bytes);
    limit_setter!(with_logical_rows, logical_rows);
    limit_setter!(with_logical_columns, logical_columns);
    limit_setter!(with_output_bytes, output_bytes);
    limit_setter!(with_operations, operations);

    fn validate(self) -> Result<Self> {
        for (kind, observed, maximum) in [
            (LimitKind::InputBytes, self.input_bytes, MAX_INPUT_BYTES),
            (LimitKind::Events, self.events, MAX_EVENTS),
            (LimitKind::Attributes, self.attributes, MAX_ATTRIBUTES),
            (LimitKind::Depth, self.depth, MAX_DEPTH),
            (LimitKind::Definitions, self.definitions, MAX_DEFINITIONS),
            (LimitKind::Sheets, self.sheets, MAX_SHEETS),
            (LimitKind::Bindings, self.bindings, MAX_BINDINGS),
            (LimitKind::TextBytes, self.text_bytes, MAX_TEXT_BYTES),
            (LimitKind::LogicalRows, self.logical_rows, MAX_LOGICAL_ROWS),
            (
                LimitKind::LogicalColumns,
                self.logical_columns,
                MAX_LOGICAL_COLUMNS,
            ),
            (LimitKind::OutputBytes, self.output_bytes, MAX_INPUT_BYTES),
            (LimitKind::Operations, self.operations, MAX_OPERATIONS),
        ] {
            if observed > maximum {
                return Err(Error::LimitExceeded {
                    kind,
                    observed,
                    maximum,
                });
            }
        }
        Ok(self)
    }

    const fn for_output_readback(mut self) -> Self {
        self.input_bytes = self.output_bytes;
        self
    }
}

/// How a validation list is presented by the producer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum DisplayList {
    None,
    Unsorted,
    SortAscending,
}

/// One document-level `table:content-validation` declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Definition {
    name: String,
    condition: Option<String>,
    base_cell_address: Option<String>,
    allow_empty_cell: Option<bool>,
    display_list: Option<DisplayList>,
    opaque_content: bool,
}

impl Definition {
    /// Construct one inert validation definition.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty name or an allocation failure.
    pub fn new(name: impl AsRef<str>) -> Result<Self> {
        let name = copy_string(name.as_ref(), "definition name")?;
        if name.is_empty() {
            return invalid("content-validation name must not be empty");
        }
        Ok(Self {
            name,
            condition: None,
            base_cell_address: None,
            allow_empty_cell: None,
            display_list: None,
            opaque_content: false,
        })
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn condition(&self) -> Option<&str> {
        self.condition.as_deref()
    }

    #[must_use]
    pub fn base_cell_address(&self) -> Option<&str> {
        self.base_cell_address.as_deref()
    }

    #[must_use]
    pub const fn allow_empty_cell(&self) -> Option<bool> {
        self.allow_empty_cell
    }

    #[must_use]
    pub const fn display_list(&self) -> Option<DisplayList> {
        self.display_list
    }

    /// Whether foreign markup, comments, or processing instructions occur in this owner.
    #[must_use]
    pub const fn has_opaque_content(&self) -> bool {
        self.opaque_content
    }

    /// Set or remove the condition lexical value.
    ///
    /// # Errors
    ///
    /// Returns an allocation error while retaining the previous value.
    pub fn set_condition(&mut self, value: Option<&str>) -> Result<()> {
        self.condition = copy_optional_string(value, "definition condition")?;
        Ok(())
    }

    /// Set or remove the base-cell-address lexical value.
    ///
    /// # Errors
    ///
    /// Returns an allocation error while retaining the previous value.
    pub fn set_base_cell_address(&mut self, value: Option<&str>) -> Result<()> {
        self.base_cell_address = copy_optional_string(value, "definition base cell address")?;
        Ok(())
    }

    pub fn set_allow_empty_cell(&mut self, value: Option<bool>) {
        self.allow_empty_cell = value;
    }

    pub fn set_display_list(&mut self, value: Option<DisplayList>) {
        self.display_list = value;
    }
}

/// A compact rectangular binding created by one physical cell run.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Binding {
    row: usize,
    column: usize,
    row_count: usize,
    column_count: usize,
    validation_name: String,
    definition_index: Option<usize>,
    covered: bool,
}

impl Binding {
    #[must_use]
    pub const fn row(&self) -> usize {
        self.row
    }

    #[must_use]
    pub const fn column(&self) -> usize {
        self.column
    }

    #[must_use]
    pub const fn row_count(&self) -> usize {
        self.row_count
    }

    #[must_use]
    pub const fn column_count(&self) -> usize {
        self.column_count
    }

    #[must_use]
    pub fn validation_name(&self) -> &str {
        &self.validation_name
    }

    /// The matching definition position, or `None` for a dangling reference.
    #[must_use]
    pub const fn definition_index(&self) -> Option<usize> {
        self.definition_index
    }

    #[must_use]
    pub const fn is_dangling(&self) -> bool {
        self.definition_index.is_none()
    }

    #[must_use]
    pub const fn is_covered_cell(&self) -> bool {
        self.covered
    }
}

/// Validation bindings belonging to one table, in physical source order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Sheet {
    name: String,
    logical_rows: usize,
    bindings: Vec<Binding>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct Layout {
    spreadsheet: Option<OwnerSpan>,
    catalog: Option<OwnerSpan>,
    catalog_insertion: usize,
    definitions: Vec<DefinitionSpan>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct OwnerSpan {
    range: Range<usize>,
    start_tag: Range<usize>,
    close_start: usize,
    qname: String,
    empty: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct DefinitionSpan {
    range: Range<usize>,
    start_tag: Range<usize>,
    empty: bool,
}

impl Sheet {
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub const fn logical_row_count(&self) -> usize {
        self.logical_rows
    }

    #[must_use]
    pub fn bindings(&self) -> &[Binding] {
        &self.bindings
    }
}

/// Immutable, read-only validation ownership projected from one exact XML source.
#[derive(Debug, PartialEq, Eq)]
pub struct Snapshot<'xml> {
    source_xml: &'xml str,
    definitions: Vec<Definition>,
    sheets: Vec<Sheet>,
    definition_index: Vec<usize>,
    sheet_index: Vec<usize>,
    catalog_opaque_content: bool,
    binding_opaque_content: bool,
    dangling_bindings: usize,
    limits: Limits,
    layout: Layout,
}

impl<'xml> Snapshot<'xml> {
    /// Parse a canonical ODS `content.xml` owner with default limits.
    pub fn parse(source_xml: &'xml str) -> Result<Self> {
        Self::parse_with_limits(source_xml, Limits::default())
    }

    /// Parse a canonical ODS `content.xml` owner with caller-selected limits.
    pub fn parse_with_limits(source_xml: &'xml str, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let mut snapshot = Parser::new(source_xml, limits).parse()?;
        snapshot.layout = scan_layout(source_xml, limits)?;
        Ok(snapshot)
    }

    /// The complete, unchanged XML source backing this inventory.
    #[must_use]
    pub const fn source_xml(&self) -> &'xml str {
        self.source_xml
    }

    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        &self.definitions
    }

    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    #[must_use]
    pub fn definition(&self, name: &str) -> Option<&Definition> {
        self.definition_index
            .binary_search_by(|index| self.definitions[*index].name.as_str().cmp(name))
            .ok()
            .map(|position| &self.definitions[self.definition_index[position]])
    }

    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheet_index
            .binary_search_by(|index| self.sheets[*index].name.as_str().cmp(name))
            .ok()
            .map(|position| &self.sheets[self.sheet_index[position]])
    }

    /// Foreign catalog children, comments, or processing instructions were retained.
    #[must_use]
    pub const fn has_opaque_catalog_content(&self) -> bool {
        self.catalog_opaque_content
    }

    #[must_use]
    pub const fn dangling_binding_count(&self) -> usize {
        self.dangling_bindings
    }

    /// Whether opaque content occurs inside a physical bound-cell owner.
    #[must_use]
    pub const fn has_opaque_binding_content(&self) -> bool {
        self.binding_opaque_content
    }

    /// Whether catalog ownership and every binding reference are fully understood.
    ///
    /// This is an inspection result, not permission to edit or publish the owner.
    #[must_use]
    pub fn has_complete_reference_closure(&self) -> bool {
        self.dangling_bindings == 0
            && !self.catalog_opaque_content
            && !self.binding_opaque_content
            && self
                .definitions
                .iter()
                .all(|definition| !definition.opaque_content)
    }

    /// Begin a bounded clone-staged edit.
    ///
    /// # Errors
    ///
    /// Returns an allocation error while cloning the modeled catalog.
    pub fn edit(&self) -> Result<Transaction<'_, 'xml>> {
        let definitions = try_clone_definitions(&self.definitions)?;
        Ok(Transaction {
            source: self,
            draft: definitions,
            operations: 0,
        })
    }
}

/// One clone-staged content-validation catalog transaction.
#[derive(Debug)]
pub struct Transaction<'snapshot, 'xml> {
    source: &'snapshot Snapshot<'xml>,
    draft: Vec<Definition>,
    operations: usize,
}

impl Transaction<'_, '_> {
    /// Borrow the current staged definition catalog.
    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        &self.draft
    }

    /// Add a new definition at the end of the catalog.
    ///
    /// # Errors
    ///
    /// Returns an error for a duplicate name, opaque ownership, a bound, or allocation failure.
    pub fn add(&mut self, definition: Definition) -> Result<()> {
        self.ensure_editable()?;
        if self.position(definition.name()).is_some() {
            return Err(Error::DuplicateDefinition(definition.name));
        }
        let mut candidate = try_clone_definitions(&self.draft)?;
        candidate
            .try_reserve(1)
            .map_err(|_error| Error::AllocationFailed("staged definition catalog"))?;
        candidate.push(definition);
        self.publish(candidate)
    }

    /// Insert a new definition or replace the same-named definition.
    ///
    /// # Errors
    ///
    /// Returns an error for opaque ownership, a bound, or allocation failure.
    pub fn set(&mut self, definition: Definition) -> Result<Option<Definition>> {
        self.ensure_editable()?;
        let mut candidate = try_clone_definitions(&self.draft)?;
        let previous = if let Some(position) = self.position(definition.name()) {
            if candidate[position] == definition {
                return Ok(Some(try_clone_definition(&candidate[position])?));
            }
            Some(core::mem::replace(&mut candidate[position], definition))
        } else {
            candidate
                .try_reserve(1)
                .map_err(|_error| Error::AllocationFailed("staged definition catalog"))?;
            candidate.push(definition);
            None
        };
        self.publish(candidate)?;
        Ok(previous)
    }

    /// Replace one existing definition without renaming it.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent definition, a rename, opaque ownership, or allocation.
    pub fn replace(&mut self, name: &str, definition: Definition) -> Result<Definition> {
        self.ensure_editable()?;
        if definition.name() != name {
            return Err(Error::UnsafeRename {
                from: copy_string(name, "rename source")?,
                to: definition.name,
            });
        }
        let Some(position) = self.position(name) else {
            return Err(Error::DefinitionNotFound(copy_string(
                name,
                "missing definition name",
            )?));
        };
        let mut candidate = try_clone_definitions(&self.draft)?;
        let previous = core::mem::replace(&mut candidate[position], definition);
        if candidate[position] == previous {
            return Ok(previous);
        }
        let returned = try_clone_definition(&previous)?;
        self.publish(candidate)?;
        Ok(returned)
    }

    /// Clone, update, validate, and atomically stage one existing definition.
    ///
    /// # Errors
    ///
    /// Returns an error from selection, the callback, validation, or allocation.
    pub fn update<F>(&mut self, name: &str, update: F) -> Result<()>
    where
        F: FnOnce(&mut Definition) -> Result<()>,
    {
        self.ensure_editable()?;
        let Some(position) = self.position(name) else {
            return Err(Error::DefinitionNotFound(copy_string(
                name,
                "missing definition name",
            )?));
        };
        let mut candidate = try_clone_definitions(&self.draft)?;
        update(&mut candidate[position])?;
        if candidate[position].name != name {
            return Err(Error::UnsafeRename {
                from: copy_string(name, "rename source")?,
                to: copy_string(&candidate[position].name, "rename target")?,
            });
        }
        if candidate[position] == self.draft[position] {
            return Ok(());
        }
        self.publish(candidate)
    }

    /// Remove an unbound definition. An absent name is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns an error when compact bindings retain the name or ownership is opaque.
    pub fn remove(&mut self, name: &str) -> Result<Option<Definition>> {
        self.ensure_editable()?;
        let Some(position) = self.position(name) else {
            return Ok(None);
        };
        self.refuse_referenced(name)?;
        let mut candidate = try_clone_definitions(&self.draft)?;
        let removed = candidate.remove(position);
        let returned = try_clone_definition(&removed)?;
        self.publish(candidate)?;
        Ok(Some(returned))
    }

    /// Remove the entire catalog only when no compact binding is retained.
    ///
    /// # Errors
    ///
    /// Returns an error for retained bindings or opaque ownership.
    pub fn clear(&mut self) -> Result<usize> {
        self.ensure_editable()?;
        let bindings = self.source.sheets.iter().try_fold(0usize, |count, sheet| {
            checked_add(count, sheet.bindings.len(), "binding closure")
        })?;
        if bindings != 0 {
            return Err(Error::DefinitionReferenced {
                name: copy_string("*", "clear binding label")?,
                bindings,
            });
        }
        let removed = self.draft.len();
        if removed != 0 {
            self.bump_operation()?;
            self.draft.clear();
        }
        Ok(removed)
    }

    /// Restore the exact source definition catalog.
    ///
    /// # Errors
    ///
    /// Returns an allocation error while rebuilding the staged catalog.
    pub fn rollback(&mut self) -> Result<()> {
        self.draft = try_clone_definitions(&self.source.definitions)?;
        self.operations = 0;
        Ok(())
    }

    /// Validate reference closure and materialize one reversible exact-source patch.
    ///
    /// # Errors
    ///
    /// Returns an error for dangling bindings, output bounds, invalid authored XML, or allocation.
    pub fn commit(self) -> Result<Commit> {
        if self.draft == self.source.definitions {
            return Commit::unchanged(self.source.source_xml);
        }
        validate_definition_catalog(&self.draft, self.source.limits)?;
        validate_binding_closure(self.source, &self.draft)?;
        let splices = build_splices(self.source, &self.draft, self.source.limits.output_bytes)?;
        let target = apply_splices(self.source.source_xml, &splices, self.source.limits)?;
        let readback =
            Snapshot::parse_with_limits(&target, self.source.limits.for_output_readback())?;
        if readback.definitions != self.draft || readback.dangling_bindings != 0 {
            return invalid("content-validation typed readback differs from the staged catalog");
        }
        Commit::from_changed(self.source.source_xml, target, splices)
    }

    fn ensure_editable(&self) -> Result<()> {
        if self.source.catalog_opaque_content
            || self.source.binding_opaque_content
            || self
                .source
                .definitions
                .iter()
                .any(Definition::has_opaque_content)
        {
            Err(Error::OpaqueOwner)
        } else {
            Ok(())
        }
    }

    fn position(&self, name: &str) -> Option<usize> {
        self.draft.iter().position(|value| value.name == name)
    }

    fn refuse_referenced(&self, name: &str) -> Result<()> {
        let bindings = self.source.sheets.iter().try_fold(0usize, |count, sheet| {
            let retained = sheet
                .bindings
                .iter()
                .filter(|binding| binding.validation_name == name)
                .count();
            checked_add(count, retained, "binding reference counter")
        })?;
        if bindings == 0 {
            Ok(())
        } else {
            Err(Error::DefinitionReferenced {
                name: copy_string(name, "referenced definition name")?,
                bindings,
            })
        }
    }

    fn publish(&mut self, candidate: Vec<Definition>) -> Result<()> {
        validate_definition_catalog(&candidate, self.source.limits)?;
        self.bump_operation()?;
        self.draft = candidate;
        Ok(())
    }

    fn bump_operation(&mut self) -> Result<()> {
        let observed = checked_add(self.operations, 1, "operation counter")?;
        check_limit(
            LimitKind::Operations,
            observed,
            self.source.limits.operations,
        )?;
        self.operations = observed;
        Ok(())
    }
}

/// A reversible exact-source content-validation patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source: String,
    target: String,
    splices: Vec<Splice>,
}

impl Patch {
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Return a patch that restores the exact accepted source.
    ///
    /// # Errors
    ///
    /// Returns an error for offset overflow or allocation failure.
    pub fn inverse(&self) -> Result<Self> {
        let mut inverse = Vec::new();
        inverse
            .try_reserve_exact(self.splices.len())
            .map_err(|_error| Error::AllocationFailed("inverse splice plan"))?;
        let mut removed = 0usize;
        let mut added = 0usize;
        for splice in &self.splices {
            let target_start = splice
                .range
                .start
                .checked_sub(removed)
                .and_then(|value| value.checked_add(added))
                .ok_or_else(|| {
                    Error::InvalidStructure("inverse splice offset overflow".to_string())
                })?;
            let target_end = checked_add(target_start, splice.after.len(), "inverse splice range")?;
            inverse.push(Splice {
                range: target_start..target_end,
                before: copy_string(&splice.after, "inverse expected source")?,
                after: copy_string(&splice.before, "inverse replacement")?,
                fragment: splice.inverse_fragment,
                inverse_fragment: splice.fragment,
            });
            removed = checked_add(
                removed,
                splice.range.end - splice.range.start,
                "inverse removed bytes",
            )?;
            added = checked_add(added, splice.after.len(), "inverse added bytes")?;
        }
        Ok(Self {
            source: copy_string(&self.target, "inverse patch source")?,
            target: copy_string(&self.source, "inverse patch target")?,
            splices: inverse,
        })
    }

    /// Apply only to the exact XML source that produced this patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::SourceMismatch`] for stale or foreign snapshots.
    pub fn apply(&self, snapshot: &Snapshot<'_>) -> Result<Commit> {
        if snapshot.source_xml != self.source {
            return Err(Error::SourceMismatch);
        }
        check_limit(
            LimitKind::OutputBytes,
            self.target.len(),
            snapshot.limits.output_bytes,
        )?;
        let target =
            Snapshot::parse_with_limits(&self.target, snapshot.limits.for_output_readback())?;
        validate_binding_closure(&target, &target.definitions)?;
        Ok(Commit {
            content_xml: copy_string(&self.target, "patch target XML")?,
            changed: !self.is_empty(),
            patch: try_clone_patch(self)?,
        })
    }
}

/// A validated semantic commit awaiting package publication.
#[derive(Debug)]
pub struct Commit {
    content_xml: String,
    changed: bool,
    patch: Patch,
}

impl Commit {
    fn unchanged(source: &str) -> Result<Self> {
        let source = copy_string(source, "unchanged patch source")?;
        Ok(Self {
            content_xml: copy_string(&source, "unchanged commit XML")?,
            changed: false,
            patch: Patch {
                target: copy_string(&source, "unchanged patch target")?,
                source,
                splices: Vec::new(),
            },
        })
    }

    fn from_changed(source: &str, target: String, splices: Vec<Splice>) -> Result<Self> {
        Ok(Self {
            content_xml: copy_string(&target, "changed commit XML")?,
            changed: true,
            patch: Patch {
                source: copy_string(source, "patch source XML")?,
                target,
                splices,
            },
        })
    }

    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        &self.content_xml
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FragmentKind {
    Markup,
    StartTag,
    Deletion,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Splice {
    range: Range<usize>,
    before: String,
    after: String,
    fragment: FragmentKind,
    inverse_fragment: FragmentKind,
}

fn try_clone_patch(source: &Patch) -> Result<Patch> {
    let mut splices = Vec::new();
    splices
        .try_reserve_exact(source.splices.len())
        .map_err(|_error| Error::AllocationFailed("patch splice clone"))?;
    for splice in &source.splices {
        splices.push(Splice {
            range: splice.range.clone(),
            before: copy_string(&splice.before, "patch expected source clone")?,
            after: copy_string(&splice.after, "patch replacement clone")?,
            fragment: splice.fragment,
            inverse_fragment: splice.inverse_fragment,
        });
    }
    Ok(Patch {
        source: copy_string(&source.source, "patch source clone")?,
        target: copy_string(&source.target, "patch target clone")?,
        splices,
    })
}

pub(crate) fn publish_package_commit(
    package: &crate::package::Package,
    commit: &Commit,
) -> litchi_core::Result<crate::package::Package> {
    if !commit.changed {
        return Err(litchi_core::Error::InvalidFormat(
            "unchanged content-validation commit must not rebuild the package".to_string(),
        ));
    }
    if package.content_xml() != commit.patch.source {
        return Err(litchi_core::Error::InvalidFormat(
            "content-validation package source snapshot does not match".to_string(),
        ));
    }
    let source_part = XmlSourcePart::load(package.package(), constants::ODF_CONTENT)?;
    let mut publication = XmlSplicePublication::new(source_part.clone());
    for splice in &commit.patch.splices {
        let proof = source_part.checked_range(splice.range.clone(), splice.before.as_bytes())?;
        let fragment = match splice.fragment {
            FragmentKind::Markup => AuthoredXmlFragment::markup(splice.after.as_bytes().to_vec())?,
            FragmentKind::StartTag => {
                AuthoredXmlFragment::start_tag(splice.after.as_bytes().to_vec())?
            },
            FragmentKind::Deletion => AuthoredXmlFragment::deletion(),
        };
        publication.replace(proof, fragment)?;
    }
    package.replace_spliced_content_xml(commit.content_xml(), publication)
}

fn validate_definition_catalog(definitions: &[Definition], limits: Limits) -> Result<()> {
    check_limit(
        LimitKind::Definitions,
        definitions.len(),
        limits.definitions,
    )?;
    let mut names = HashSet::new();
    names
        .try_reserve(definitions.len())
        .map_err(|_error| Error::AllocationFailed("staged definition name index"))?;
    let mut text_bytes = 0usize;
    for definition in definitions {
        if definition.name.is_empty() {
            return invalid("content-validation name must not be empty");
        }
        if definition.opaque_content {
            return Err(Error::OpaqueOwner);
        }
        if !names.insert(definition.name.as_str()) {
            return Err(Error::DuplicateDefinition(copy_string(
                &definition.name,
                "duplicate definition name",
            )?));
        }
        for value in [
            Some(definition.name.as_str()),
            definition.condition.as_deref(),
            definition.base_cell_address.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            text_bytes = checked_add(text_bytes, value.len(), "staged definition text bytes")?;
        }
    }
    check_limit(LimitKind::TextBytes, text_bytes, limits.text_bytes)
}

fn validate_binding_closure(snapshot: &Snapshot<'_>, definitions: &[Definition]) -> Result<()> {
    let mut names = HashSet::new();
    names
        .try_reserve(definitions.len())
        .map_err(|_error| Error::AllocationFailed("binding closure name index"))?;
    names.extend(
        definitions
            .iter()
            .map(|definition| definition.name.as_str()),
    );
    let mut dangling = 0usize;
    for binding in snapshot
        .sheets
        .iter()
        .flat_map(|sheet| sheet.bindings.iter())
    {
        if !names.contains(binding.validation_name.as_str()) {
            dangling = checked_add(dangling, 1, "dangling binding closure")?;
        }
    }
    if dangling == 0 {
        Ok(())
    } else {
        invalid(format!(
            "content-validation edit leaves {dangling} compact binding(s) dangling"
        ))
    }
}

fn try_clone_definition(source: &Definition) -> Result<Definition> {
    Ok(Definition {
        name: copy_string(&source.name, "definition name")?,
        condition: copy_optional_string(source.condition.as_deref(), "definition condition")?,
        base_cell_address: copy_optional_string(
            source.base_cell_address.as_deref(),
            "definition base cell address",
        )?,
        allow_empty_cell: source.allow_empty_cell,
        display_list: source.display_list,
        opaque_content: source.opaque_content,
    })
}

fn try_clone_definitions(source: &[Definition]) -> Result<Vec<Definition>> {
    let mut definitions = Vec::new();
    definitions
        .try_reserve_exact(source.len())
        .map_err(|_error| Error::AllocationFailed("staged definition catalog"))?;
    for definition in source {
        definitions.push(try_clone_definition(definition)?);
    }
    Ok(definitions)
}

fn build_splices(
    snapshot: &Snapshot<'_>,
    draft: &[Definition],
    output_limit: usize,
) -> Result<Vec<Splice>> {
    let removed = planned_removed_bytes(snapshot, draft)?;
    let retained = snapshot
        .source_xml
        .len()
        .checked_sub(removed)
        .ok_or_else(|| Error::InvalidStructure("splice removal exceeds source".to_string()))?;
    if retained > output_limit {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: retained,
            maximum: output_limit,
        });
    }
    let mut render_budget = RenderBudget {
        remaining: output_limit - retained,
        maximum: output_limit,
    };
    let mut splices = Vec::new();
    splices
        .try_reserve(snapshot.definitions.len().saturating_add(1))
        .map_err(|_error| Error::AllocationFailed("content-validation splice plan"))?;
    let source = snapshot.source_xml;
    let Some(catalog) = &snapshot.layout.catalog else {
        if draft.is_empty() {
            return Ok(splices);
        }
        let catalog_xml = render_catalog(draft, &mut render_budget)?;
        let spreadsheet = snapshot.layout.spreadsheet.as_ref().ok_or_else(|| {
            Error::InvalidStructure("content-validation spreadsheet layout is missing".to_string())
        })?;
        if spreadsheet.empty {
            let before = checked_source(source, spreadsheet.range.clone())?;
            let mut after = String::new();
            let opening = checked_source(source, spreadsheet.start_tag.clone())?
                .trim_end()
                .strip_suffix("/>")
                .ok_or_else(|| {
                    Error::InvalidStructure(
                        "empty spreadsheet owner has no self-closing start tag".to_string(),
                    )
                })?;
            render_budget.push_str(&mut after, opening)?;
            render_budget.push_char(&mut after, '>')?;
            after
                .try_reserve(catalog_xml.len())
                .map_err(|_error| Error::AllocationFailed("spreadsheet expansion"))?;
            after.push_str(&catalog_xml);
            render_budget.push_str(&mut after, "</")?;
            render_budget.push_str(&mut after, &spreadsheet.qname)?;
            render_budget.push_char(&mut after, '>')?;
            splices.push(Splice {
                range: spreadsheet.range.clone(),
                before: copy_string(before, "empty spreadsheet source")?,
                after,
                fragment: FragmentKind::Markup,
                inverse_fragment: FragmentKind::Markup,
            });
        } else {
            splices.push(Splice {
                range: snapshot.layout.catalog_insertion..snapshot.layout.catalog_insertion,
                before: String::new(),
                after: catalog_xml,
                fragment: FragmentKind::Markup,
                inverse_fragment: FragmentKind::Deletion,
            });
        }
        return Ok(splices);
    };

    if snapshot.layout.definitions.len() != snapshot.definitions.len() {
        return invalid("content-validation semantic and physical definition counts differ");
    }
    if draft.is_empty() {
        let before = checked_source(source, catalog.range.clone())?;
        splices.push(Splice {
            range: catalog.range.clone(),
            before: copy_string(before, "catalog deletion source")?,
            after: String::new(),
            fragment: FragmentKind::Deletion,
            inverse_fragment: FragmentKind::Markup,
        });
        return Ok(splices);
    }

    let mut target = HashMap::new();
    target
        .try_reserve(draft.len())
        .map_err(|_error| Error::AllocationFailed("target definition index"))?;
    for definition in draft {
        target.insert(definition.name.as_str(), definition);
    }
    let mut source_names = HashSet::new();
    source_names
        .try_reserve(snapshot.definitions.len())
        .map_err(|_error| Error::AllocationFailed("source definition index"))?;
    for (position, definition) in snapshot.definitions.iter().enumerate() {
        source_names.insert(definition.name.as_str());
        let span = &snapshot.layout.definitions[position];
        match target.get(definition.name.as_str()) {
            None => {
                let before = checked_source(source, span.range.clone())?;
                splices.push(Splice {
                    range: span.range.clone(),
                    before: copy_string(before, "definition deletion source")?,
                    after: String::new(),
                    fragment: FragmentKind::Deletion,
                    inverse_fragment: FragmentKind::Markup,
                });
            },
            Some(replacement) if *replacement != definition => {
                let before = checked_source(source, span.start_tag.clone())?;
                let after = render_existing_definition_start(
                    before,
                    replacement,
                    span.empty,
                    &mut render_budget,
                )?;
                let fragment = if span.empty {
                    FragmentKind::Markup
                } else {
                    FragmentKind::StartTag
                };
                splices.push(Splice {
                    range: span.start_tag.clone(),
                    before: copy_string(before, "definition start-tag source")?,
                    after,
                    fragment,
                    inverse_fragment: fragment,
                });
            },
            Some(_) => {},
        }
    }
    let mut additions = String::new();
    for definition in draft {
        if !source_names.contains(definition.name.as_str()) {
            render_new_definition_into(&mut additions, definition, &mut render_budget)?;
        }
    }
    if !additions.is_empty() {
        splices.push(Splice {
            range: catalog.close_start..catalog.close_start,
            before: String::new(),
            after: additions,
            fragment: FragmentKind::Markup,
            inverse_fragment: FragmentKind::Deletion,
        });
    }
    splices.sort_unstable_by_key(|splice| splice.range.start);
    for pair in splices.windows(2) {
        if pair[0].range.end > pair[1].range.start
            || (pair[0].range.is_empty()
                && pair[1].range.is_empty()
                && pair[0].range.start == pair[1].range.start)
        {
            return invalid("content-validation splice ranges overlap");
        }
    }
    Ok(splices)
}

fn planned_removed_bytes(snapshot: &Snapshot<'_>, draft: &[Definition]) -> Result<usize> {
    let Some(catalog) = &snapshot.layout.catalog else {
        return Ok(snapshot
            .layout
            .spreadsheet
            .as_ref()
            .filter(|spreadsheet| spreadsheet.empty && !draft.is_empty())
            .map_or(0, |spreadsheet| {
                spreadsheet.range.end - spreadsheet.range.start
            }));
    };
    if draft.is_empty() {
        return Ok(catalog.range.end - catalog.range.start);
    }
    let mut names = HashMap::new();
    names
        .try_reserve(draft.len())
        .map_err(|_error| Error::AllocationFailed("planned definition index"))?;
    for definition in draft {
        names.insert(definition.name.as_str(), definition);
    }
    snapshot
        .definitions
        .iter()
        .enumerate()
        .try_fold(0usize, |removed, (position, definition)| {
            let span = &snapshot.layout.definitions[position];
            let bytes = match names.get(definition.name.as_str()) {
                None => span.range.end - span.range.start,
                Some(target) if *target != definition => span.start_tag.end - span.start_tag.start,
                Some(_) => 0,
            };
            checked_add(removed, bytes, "planned splice removal bytes")
        })
}

fn apply_splices(source: &str, splices: &[Splice], limits: Limits) -> Result<String> {
    let removed = splices.iter().try_fold(0usize, |total, splice| {
        checked_add(
            total,
            splice.range.end - splice.range.start,
            "removed splice bytes",
        )
    })?;
    let added = splices.iter().try_fold(0usize, |total, splice| {
        checked_add(total, splice.after.len(), "added splice bytes")
    })?;
    let length = source
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| {
            Error::InvalidStructure("content-validation output size overflow".to_string())
        })?;
    check_limit(LimitKind::OutputBytes, length, limits.output_bytes)?;
    let mut output = String::new();
    reserve_string(&mut output, length, "content-validation output")?;
    let mut cursor = 0usize;
    for splice in splices {
        let actual = checked_source(source, splice.range.clone())?;
        if actual != splice.before {
            return Err(Error::SourceMismatch);
        }
        output.push_str(&source[cursor..splice.range.start]);
        output.push_str(&splice.after);
        cursor = splice.range.end;
    }
    output.push_str(&source[cursor..]);
    Ok(output)
}

fn render_catalog(definitions: &[Definition], budget: &mut RenderBudget) -> Result<String> {
    let mut output = String::new();
    budget.push_str(&mut output, "<table:content-validations xmlns:table=\"")?;
    budget.push_str(
        &mut output,
        core::str::from_utf8(TABLE).map_err(|_error| {
            Error::InvalidStructure("table namespace is not UTF-8".to_string())
        })?,
    )?;
    budget.push_str(&mut output, "\">")?;
    for definition in definitions {
        render_new_definition_into(&mut output, definition, budget)?;
    }
    budget.push_str(&mut output, "</table:content-validations>")?;
    Ok(output)
}

fn render_new_definition_into(
    output: &mut String,
    definition: &Definition,
    budget: &mut RenderBudget,
) -> Result<()> {
    budget.push_str(output, "<table:content-validation xmlns:table=\"")?;
    budget.push_str(
        output,
        core::str::from_utf8(TABLE).map_err(|_error| {
            Error::InvalidStructure("table namespace is not UTF-8".to_string())
        })?,
    )?;
    budget.push_char(output, '"')?;
    render_definition_attributes(output, "table", definition, budget)?;
    budget.push_str(output, "/>")
}

fn render_existing_definition_start(
    source_tag: &str,
    definition: &Definition,
    empty: bool,
    budget: &mut RenderBudget,
) -> Result<String> {
    let mut reader = quick_xml::Reader::from_str(source_tag);
    reader.config_mut().trim_text(false);
    let event = reader
        .read_event()
        .map_err(|error| Error::InvalidXml(format!("invalid definition start tag: {error}")))?;
    let start = match event {
        Event::Start(start) | Event::Empty(start) => start,
        _ => return invalid("definition source range is not a start tag"),
    };
    let start_name = start.name();
    let qname = core::str::from_utf8(start_name.as_ref()).map_err(|_error| {
        Error::InvalidXml("definition qualified name is not UTF-8".to_string())
    })?;
    let prefix = select_authored_prefix(&start)?;
    let mut output = String::new();
    budget.push_char(&mut output, '<')?;
    budget.push_str(&mut output, qname)?;
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidXml(format!("invalid definition start-tag attribute: {error}"))
        })?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            budget.push_char(&mut output, ' ')?;
            budget.push_str(
                &mut output,
                core::str::from_utf8(attribute.key.as_ref()).map_err(|_error| {
                    Error::InvalidXml("namespace declaration name is not UTF-8".to_string())
                })?,
            )?;
            budget.push_str(&mut output, "=\"")?;
            budget.push_str(
                &mut output,
                core::str::from_utf8(attribute.value.as_ref()).map_err(|_error| {
                    Error::InvalidXml("namespace declaration value is not UTF-8".to_string())
                })?,
            )?;
            budget.push_char(&mut output, '"')?;
        }
    }
    budget.push_str(&mut output, " xmlns:")?;
    budget.push_str(&mut output, &prefix)?;
    budget.push_str(&mut output, "=\"")?;
    budget.push_str(
        &mut output,
        core::str::from_utf8(TABLE).map_err(|_error| {
            Error::InvalidStructure("table namespace is not UTF-8".to_string())
        })?,
    )?;
    budget.push_char(&mut output, '"')?;
    render_definition_attributes(&mut output, &prefix, definition, budget)?;
    budget.push_str(&mut output, if empty { "/>" } else { ">" })?;
    Ok(output)
}

fn select_authored_prefix(start: &BytesStart<'_>) -> Result<String> {
    let mut used = HashSet::new();
    let estimate = start.attributes().count();
    used.try_reserve(estimate)
        .map_err(|_error| Error::AllocationFailed("definition namespace prefix index"))?;
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidXml(format!("invalid definition namespace declaration: {error}"))
        })?;
        if let Some(prefix) = attribute.key.as_ref().strip_prefix(b"xmlns:") {
            used.insert(copy_bytes(prefix, "definition namespace prefix")?);
        }
    }
    let base = b"litchicv".as_slice();
    if !used.contains(base) {
        return copy_string("litchicv", "authored namespace prefix");
    }
    for suffix in 1..=used.len() {
        let suffix = decimal_string(suffix)?;
        let mut candidate = String::new();
        reserve_string(
            &mut candidate,
            "litchicv".len().checked_add(suffix.len()).ok_or_else(|| {
                Error::InvalidStructure("namespace prefix size overflow".to_string())
            })?,
            "authored namespace prefix",
        )?;
        candidate.push_str("litchicv");
        candidate.push_str(&suffix);
        if !used.contains(candidate.as_bytes()) {
            return Ok(candidate);
        }
    }
    invalid("could not select a fresh content-validation namespace prefix")
}

fn decimal_string(mut value: usize) -> Result<String> {
    let mut buffer = [0u8; 20];
    let mut position = buffer.len();
    loop {
        position -= 1;
        buffer[position] = b'0'
            + u8::try_from(value % 10).map_err(|_error| {
                Error::InvalidStructure("namespace prefix digit overflow".to_string())
            })?;
        value /= 10;
        if value == 0 {
            break;
        }
    }
    copy_utf8(&buffer[position..], "namespace prefix suffix")
}

fn render_definition_attributes(
    output: &mut String,
    prefix: &str,
    definition: &Definition,
    budget: &mut RenderBudget,
) -> Result<()> {
    push_attribute(output, prefix, "name", &definition.name, budget)?;
    if let Some(value) = &definition.condition {
        push_attribute(output, prefix, "condition", value, budget)?;
    }
    if let Some(value) = &definition.base_cell_address {
        push_attribute(output, prefix, "base-cell-address", value, budget)?;
    }
    if let Some(value) = definition.allow_empty_cell {
        push_attribute(
            output,
            prefix,
            "allow-empty-cell",
            if value { "true" } else { "false" },
            budget,
        )?;
    }
    if let Some(value) = definition.display_list {
        push_attribute(
            output,
            prefix,
            "display-list",
            match value {
                DisplayList::None => "none",
                DisplayList::Unsorted => "unsorted",
                DisplayList::SortAscending => "sort-ascending",
            },
            budget,
        )?;
    }
    Ok(())
}

fn push_attribute(
    output: &mut String,
    prefix: &str,
    local: &str,
    value: &str,
    budget: &mut RenderBudget,
) -> Result<()> {
    budget.push_char(output, ' ')?;
    budget.push_str(output, prefix)?;
    budget.push_char(output, ':')?;
    budget.push_str(output, local)?;
    budget.push_str(output, "=\"")?;
    push_escaped_attribute(output, value, budget)?;
    budget.push_char(output, '"')?;
    Ok(())
}

fn push_escaped_attribute(
    output: &mut String,
    value: &str,
    budget: &mut RenderBudget,
) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => budget.push_str(output, "&amp;")?,
            '<' => budget.push_str(output, "&lt;")?,
            '"' => budget.push_str(output, "&quot;")?,
            '\r' => budget.push_str(output, "&#13;")?,
            '\n' => budget.push_str(output, "&#10;")?,
            '\t' => budget.push_str(output, "&#9;")?,
            value if is_xml_character(value) => budget.push_char(output, value)?,
            _ => return invalid("content-validation value contains an invalid XML character"),
        }
    }
    Ok(())
}

fn is_xml_character(value: char) -> bool {
    matches!(value, '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
}

fn checked_source(source: &str, range: Range<usize>) -> Result<&str> {
    source.get(range).ok_or_else(|| {
        Error::InvalidStructure("content-validation source range is invalid".to_string())
    })
}

fn reserve_string(value: &mut String, additional: usize, owner: &'static str) -> Result<()> {
    value
        .try_reserve(additional)
        .map_err(|_error| Error::AllocationFailed(owner))
}

struct RenderBudget {
    remaining: usize,
    maximum: usize,
}

impl RenderBudget {
    fn push_str(&mut self, output: &mut String, value: &str) -> Result<()> {
        if value.len() > self.remaining {
            return Err(Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                observed: self.maximum.saturating_add(1),
                maximum: self.maximum,
            });
        }
        output
            .try_reserve(value.len())
            .map_err(|_error| Error::AllocationFailed("authored XML output"))?;
        output.push_str(value);
        self.remaining -= value.len();
        Ok(())
    }

    fn push_char(&mut self, output: &mut String, value: char) -> Result<()> {
        let mut encoded = [0u8; 4];
        self.push_str(output, value.encode_utf8(&mut encoded))
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Text,
    Script,
    Presentation,
    Other,
}

#[derive(Debug)]
enum Element {
    Root,
    Body,
    Spreadsheet,
    Catalog(usize),
    Definition,
    Sheet(usize),
    RowContainer(usize),
    Row {
        sheet: usize,
        start: usize,
        count: usize,
        next_column: usize,
    },
    Cell(bool),
    Other,
}

struct Parser<'xml> {
    source_xml: &'xml str,
    limits: Limits,
    definitions: Vec<Definition>,
    sheets: Vec<Sheet>,
    stack: Vec<Element>,
    events: usize,
    attributes: usize,
    text_bytes: usize,
    binding_count: usize,
    seen_root: bool,
    seen_body: bool,
    seen_spreadsheet: bool,
    closed_root: bool,
    seen_catalog: bool,
    seen_table: bool,
    spreadsheet_prelude_complete: bool,
    catalog_active: bool,
    active_definition: Option<usize>,
    definition_state: Option<DefinitionState>,
    catalog_opaque_content: bool,
    binding_opaque_content: bool,
}

#[derive(Default)]
struct DefinitionState {
    seen_help: bool,
    seen_error_message: bool,
    seen_error_macro: bool,
}

impl<'xml> Parser<'xml> {
    fn new(source_xml: &'xml str, limits: Limits) -> Self {
        Self {
            source_xml,
            limits,
            definitions: Vec::new(),
            sheets: Vec::new(),
            stack: Vec::new(),
            events: 0,
            attributes: 0,
            text_bytes: 0,
            binding_count: 0,
            seen_root: false,
            seen_body: false,
            seen_spreadsheet: false,
            closed_root: false,
            seen_catalog: false,
            seen_table: false,
            spreadsheet_prelude_complete: false,
            catalog_active: false,
            active_definition: None,
            definition_state: None,
            catalog_opaque_content: false,
            binding_opaque_content: false,
        }
    }

    fn parse(mut self) -> Result<Snapshot<'xml>> {
        check_limit(
            LimitKind::InputBytes,
            self.source_xml.len(),
            self.limits.input_bytes,
        )?;
        self.stack
            .try_reserve(self.limits.depth.min(256))
            .map_err(|_error| Error::AllocationFailed("element stack"))?;
        self.definitions
            .try_reserve(self.limits.definitions.min(256))
            .map_err(|_error| Error::AllocationFailed("definition catalog"))?;
        self.sheets
            .try_reserve(self.limits.sheets.min(64))
            .map_err(|_error| Error::AllocationFailed("sheet inventory"))?;

        let mut reader = NsReader::from_str(self.source_xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().trim_text(false);
        loop {
            let (resolved, event) = reader.read_resolved_event().map_err(|error| {
                Error::InvalidXml(format!("invalid ODS content-validation XML: {error}"))
            })?;
            self.events = checked_add(self.events, 1, "XML event counter")?;
            check_limit(LimitKind::Events, self.events, self.limits.events)?;
            let namespace = namespace_kind(&resolved)?;
            match event {
                Event::Start(start) => self.start(namespace, &start, &reader, false)?,
                Event::Empty(start) => self.start(namespace, &start, &reader, true)?,
                Event::End(_) => self.end()?,
                Event::Text(text) => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidXml(format!("invalid XML text: {error}")))?;
                    if self.stack.is_empty() {
                        if !decoded.trim().is_empty() {
                            return Err(Error::InvalidStructure(
                                "ODS content-validation XML has text outside its root".to_string(),
                            ));
                        }
                    } else if !decoded.trim().is_empty() {
                        self.mark_opaque();
                    }
                },
                Event::Comment(_) | Event::PI(_) => self.mark_opaque(),
                Event::DocType(_) | Event::GeneralRef(_) => {
                    return Err(Error::InvalidStructure(
                        "ODS content-validation inspection rejects DTDs and entity references"
                            .to_string(),
                    ));
                },
                Event::CData(_) => self.mark_opaque(),
                Event::Decl(_) => {},
                Event::Eof => break,
            }
        }
        if !self.seen_root
            || !self.seen_body
            || !self.seen_spreadsheet
            || !self.closed_root
            || !self.stack.is_empty()
        {
            return Err(Error::InvalidStructure(
                "ODS content-validation XML lacks a complete document-content envelope".to_string(),
            ));
        }
        self.finish()
    }

    fn start(
        &mut self,
        namespace: NamespaceKind,
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
        empty: bool,
    ) -> Result<()> {
        check_limit(
            LimitKind::Depth,
            self.stack.len().saturating_add(1),
            self.limits.depth,
        )?;
        self.validate_attributes(start, reader)?;
        let local = start.local_name();
        let local = local.as_ref();
        if self.has_attribute(start, reader, TABLE, b"content-validation-name")?
            && !(namespace == NamespaceKind::Table
                && matches!(local, b"table-cell" | b"covered-table-cell"))
        {
            return invalid("table:content-validation-name is only valid on a table cell owner");
        }
        let element = self.classify(namespace, local, start, reader)?;
        if empty {
            self.finish_element(element)?;
        } else {
            self.stack
                .try_reserve(1)
                .map_err(|_error| Error::AllocationFailed("element stack"))?;
            self.stack.push(element);
        }
        Ok(())
    }

    fn classify(
        &mut self,
        namespace: NamespaceKind,
        local: &[u8],
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
    ) -> Result<Element> {
        if namespace == NamespaceKind::Office && local == b"document-content" {
            if self.seen_root || self.closed_root || !self.stack.is_empty() {
                return invalid("multiple or nested office:document-content roots");
            }
            self.seen_root = true;
            return Ok(Element::Root);
        }
        if namespace == NamespaceKind::Office && local == b"body" {
            if self.seen_body {
                return invalid("duplicate office:body owner");
            }
            self.seen_body = true;
            return direct(
                &self.stack,
                |parent| matches!(parent, Element::Root),
                Element::Body,
                "office:body must be the direct child of office:document-content",
            );
        }
        if namespace == NamespaceKind::Office && local == b"spreadsheet" {
            if self.seen_spreadsheet {
                return invalid("duplicate office:spreadsheet owner");
            }
            self.seen_spreadsheet = true;
            return direct(
                &self.stack,
                |parent| matches!(parent, Element::Body),
                Element::Spreadsheet,
                "office:spreadsheet must be the direct child of office:body",
            );
        }
        if self.stack.is_empty() || self.closed_root {
            return invalid("content occurs outside the canonical document root");
        }
        if namespace == NamespaceKind::Table && local == b"content-validations" {
            if !matches!(self.stack.last(), Some(Element::Spreadsheet))
                || self.seen_catalog
                || self.seen_table
                || self.spreadsheet_prelude_complete
            {
                return invalid(
                    "table:content-validations must be the unique catalog before worksheet tables",
                );
            }
            self.seen_catalog = true;
            self.catalog_active = true;
            if has_non_namespace_attributes(start) {
                self.catalog_opaque_content = true;
            }
            return Ok(Element::Catalog(self.definitions.len()));
        }
        if matches!(self.stack.last(), Some(Element::Spreadsheet))
            && namespace == NamespaceKind::Table
            && local != b"content-validations"
            && !matches!(local, b"tracked-changes" | b"calculation-settings")
        {
            self.spreadsheet_prelude_complete = true;
        }
        if namespace == NamespaceKind::Table && local == b"content-validation" {
            if !matches!(self.stack.last(), Some(Element::Catalog(_))) {
                return invalid(
                    "table:content-validation must be a direct child of table:content-validations",
                );
            }
            check_limit(
                LimitKind::Definitions,
                self.definitions.len().saturating_add(1),
                self.limits.definitions,
            )?;
            let definition = self.parse_definition(start, reader)?;
            self.definitions
                .try_reserve(1)
                .map_err(|_error| Error::AllocationFailed("definition catalog"))?;
            self.definitions.push(definition);
            let index = self.definitions.len() - 1;
            self.active_definition = Some(index);
            self.definition_state = Some(DefinitionState::default());
            return Ok(Element::Definition);
        }
        if let Some(index) = self.active_definition {
            if matches!(
                (namespace, local),
                (
                    NamespaceKind::Table,
                    b"help-message" | b"error-message" | b"error-macro"
                )
            ) {
                if !matches!(self.stack.last(), Some(Element::Definition)) {
                    return invalid("validation action must be a direct definition child");
                }
                let state = self.definition_state.as_mut().ok_or_else(|| {
                    Error::InvalidStructure("validation action state is missing".to_string())
                })?;
                match local {
                    b"help-message"
                        if !state.seen_help
                            && !state.seen_error_message
                            && !state.seen_error_macro =>
                    {
                        state.seen_help = true;
                    },
                    b"error-message" if !state.seen_error_message && !state.seen_error_macro => {
                        state.seen_error_message = true;
                    },
                    b"error-macro" if !state.seen_error_macro && !state.seen_error_message => {
                        state.seen_error_macro = true;
                    },
                    _ => return invalid("duplicate or out-of-order validation action"),
                }
            }
            // This prerequisite does not semantically own rich messages, events,
            // or foreign descendants; the exact source retains them instead.
            self.definitions[index].opaque_content = true;
        }
        if namespace == NamespaceKind::Table && local == b"table" {
            if !matches!(self.stack.last(), Some(Element::Spreadsheet)) {
                return Ok(Element::Other);
            }
            self.seen_table = true;
            check_limit(
                LimitKind::Sheets,
                self.sheets.len().saturating_add(1),
                self.limits.sheets,
            )?;
            let name = self.required_attribute(start, reader, TABLE, b"name", "table:name")?;
            let mut bindings = Vec::new();
            bindings
                .try_reserve(self.limits.bindings.min(16))
                .map_err(|_error| Error::AllocationFailed("cell bindings"))?;
            self.sheets
                .try_reserve(1)
                .map_err(|_error| Error::AllocationFailed("sheet inventory"))?;
            self.sheets.push(Sheet {
                name,
                logical_rows: 0,
                bindings,
            });
            return Ok(Element::Sheet(self.sheets.len() - 1));
        }
        if namespace == NamespaceKind::Table
            && matches!(
                local,
                b"table-header-rows" | b"table-row-group" | b"table-rows"
            )
        {
            let sheet = match self.stack.last() {
                Some(Element::Sheet(sheet) | Element::RowContainer(sheet)) => *sheet,
                _ => return Ok(Element::Other),
            };
            return Ok(Element::RowContainer(sheet));
        }
        if namespace == NamespaceKind::Table && local == b"table-row" {
            let sheet = match self.stack.last() {
                Some(Element::Sheet(sheet) | Element::RowContainer(sheet)) => *sheet,
                _ => return Ok(Element::Other),
            };
            let count = self
                .positive_attribute(
                    start,
                    reader,
                    TABLE,
                    b"number-rows-repeated",
                    "table:number-rows-repeated",
                )?
                .unwrap_or(1);
            let row = self.sheets[sheet].logical_rows;
            let end = checked_add(row, count, "repeated row extent")?;
            check_limit(LimitKind::LogicalRows, end, self.limits.logical_rows)?;
            return Ok(Element::Row {
                sheet,
                start: row,
                count,
                next_column: 0,
            });
        }
        if namespace == NamespaceKind::Table
            && matches!(local, b"table-cell" | b"covered-table-cell")
        {
            let (sheet, row, row_count, column) = match self.stack.last() {
                Some(Element::Row {
                    sheet,
                    start,
                    count,
                    next_column,
                }) => (*sheet, *start, *count, *next_column),
                _ => return Ok(Element::Other),
            };
            let column_count = self
                .positive_attribute(
                    start,
                    reader,
                    TABLE,
                    b"number-columns-repeated",
                    "table:number-columns-repeated",
                )?
                .unwrap_or(1);
            let end = checked_add(column, column_count, "repeated column extent")?;
            check_limit(LimitKind::LogicalColumns, end, self.limits.logical_columns)?;
            let validation_name =
                self.optional_attribute(start, reader, TABLE, b"content-validation-name")?;
            let Some(Element::Row { next_column, .. }) = self.stack.last_mut() else {
                return invalid("table cell lost its table-row owner");
            };
            *next_column = end;
            let bound = validation_name.is_some();
            if let Some(validation_name) = validation_name {
                if validation_name.is_empty() {
                    return invalid("table:content-validation-name must not be empty");
                }
                self.binding_count = checked_add(self.binding_count, 1, "binding counter")?;
                check_limit(
                    LimitKind::Bindings,
                    self.binding_count,
                    self.limits.bindings,
                )?;
                self.sheets[sheet]
                    .bindings
                    .try_reserve(1)
                    .map_err(|_error| Error::AllocationFailed("cell binding"))?;
                self.sheets[sheet].bindings.push(Binding {
                    row,
                    column,
                    row_count,
                    column_count,
                    validation_name,
                    definition_index: None,
                    covered: local == b"covered-table-cell",
                });
            }
            return Ok(Element::Cell(bound));
        }

        if self.active_definition.is_some() {
            if !known_definition_element(namespace, local) {
                self.mark_opaque();
            } else if let Some(index) = self.active_definition
                && has_unknown_definition_attributes(namespace, local, start, reader)?
            {
                self.definitions[index].opaque_content = true;
            }
        } else if matches!(self.stack.last(), Some(Element::Catalog(_))) {
            self.catalog_opaque_content = true;
        } else if self
            .stack
            .iter()
            .any(|element| matches!(element, Element::Cell(true)))
        {
            self.binding_opaque_content = true;
        }
        Ok(Element::Other)
    }

    fn end(&mut self) -> Result<()> {
        let element = self.stack.pop().ok_or_else(|| {
            Error::InvalidStructure("ODS content-validation element stack underflow".to_string())
        })?;
        self.finish_element(element)
    }

    fn finish_element(&mut self, element: Element) -> Result<()> {
        match element {
            Element::Root => self.closed_root = true,
            Element::Catalog(start) => {
                self.catalog_active = false;
                if self.definitions.len() == start {
                    return invalid("table:content-validations must contain a definition");
                }
            },
            Element::Definition => {
                self.active_definition = None;
                self.definition_state = None;
            },
            Element::Row {
                sheet,
                start,
                count,
                ..
            } => {
                self.sheets[sheet].logical_rows = checked_add(start, count, "row extent")?;
            },
            Element::Body
            | Element::Spreadsheet
            | Element::Sheet(_)
            | Element::RowContainer(_)
            | Element::Cell(_)
            | Element::Other => {},
        }
        Ok(())
    }

    fn finish(mut self) -> Result<Snapshot<'xml>> {
        let mut definitions = HashMap::new();
        definitions
            .try_reserve(self.definitions.len())
            .map_err(|_error| Error::AllocationFailed("definition name index"))?;
        for (index, definition) in self.definitions.iter().enumerate() {
            if definitions
                .insert(definition.name.as_str(), index)
                .is_some()
            {
                return invalid(format!(
                    "duplicate content-validation definition '{}'",
                    definition.name
                ));
            }
        }
        let mut sheet_names = HashSet::new();
        sheet_names
            .try_reserve(self.sheets.len())
            .map_err(|_error| Error::AllocationFailed("sheet name index"))?;
        let mut dangling_bindings = 0usize;
        for sheet in &mut self.sheets {
            if !sheet_names.insert(sheet.name.as_str()) {
                return invalid(format!("duplicate table name '{}'", sheet.name));
            }
            for binding in &mut sheet.bindings {
                binding.definition_index =
                    definitions.get(binding.validation_name.as_str()).copied();
                if binding.definition_index.is_none() {
                    dangling_bindings =
                        checked_add(dangling_bindings, 1, "dangling validation binding counter")?;
                }
            }
        }
        let mut definition_index = Vec::new();
        definition_index
            .try_reserve_exact(self.definitions.len())
            .map_err(|_error| Error::AllocationFailed("definition lookup index"))?;
        definition_index.extend(0..self.definitions.len());
        definition_index.sort_unstable_by(|left, right| {
            self.definitions[*left]
                .name
                .cmp(&self.definitions[*right].name)
        });
        let mut sheet_index = Vec::new();
        sheet_index
            .try_reserve_exact(self.sheets.len())
            .map_err(|_error| Error::AllocationFailed("sheet lookup index"))?;
        sheet_index.extend(0..self.sheets.len());
        sheet_index
            .sort_unstable_by(|left, right| self.sheets[*left].name.cmp(&self.sheets[*right].name));
        Ok(Snapshot {
            source_xml: self.source_xml,
            definitions: self.definitions,
            sheets: self.sheets,
            definition_index,
            sheet_index,
            catalog_opaque_content: self.catalog_opaque_content,
            binding_opaque_content: self.binding_opaque_content,
            dangling_bindings,
            limits: self.limits,
            layout: Layout::default(),
        })
    }

    fn parse_definition(
        &mut self,
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
    ) -> Result<Definition> {
        let name = self.required_attribute(start, reader, TABLE, b"name", "table:name")?;
        if name.is_empty() {
            return invalid("content-validation name must not be empty");
        }
        let condition = self.optional_attribute(start, reader, TABLE, b"condition")?;
        let base_cell_address =
            self.optional_attribute(start, reader, TABLE, b"base-cell-address")?;
        let allow_empty_cell = self
            .optional_attribute(start, reader, TABLE, b"allow-empty-cell")?
            .map(|value| parse_bool(&value, "table:allow-empty-cell"))
            .transpose()?;
        let display_list = self
            .optional_attribute(start, reader, TABLE, b"display-list")?
            .map(|value| match value.as_str() {
                "none" => Ok(DisplayList::None),
                "unsorted" => Ok(DisplayList::Unsorted),
                "sort-ascending" => Ok(DisplayList::SortAscending),
                _ => invalid(format!("invalid table:display-list value '{value}'")),
            })
            .transpose()?;
        Ok(Definition {
            name,
            condition,
            base_cell_address,
            allow_empty_cell,
            display_list,
            opaque_content: has_unknown_definition_attributes(
                NamespaceKind::Table,
                b"content-validation",
                start,
                reader,
            )?,
        })
    }

    fn validate_attributes(
        &mut self,
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
    ) -> Result<()> {
        let estimate = start.attributes().count();
        let observed = checked_add(self.attributes, estimate, "attribute counter")?;
        check_limit(LimitKind::Attributes, observed, self.limits.attributes)?;
        self.attributes = observed;
        let mut namespace_ids = AttributeNamespaceIds::with_capacity(estimate)?;
        let mut expanded = HashSet::<(usize, Vec<u8>)>::new();
        expanded
            .try_reserve(estimate)
            .map_err(|_error| Error::AllocationFailed("expanded attribute names"))?;
        for attribute in start.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| {
                Error::InvalidXml(format!("invalid ODS XML attribute: {error}"))
            })?;
            if is_namespace_declaration(attribute.key.as_ref()) {
                continue;
            }
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            let namespace = namespace_bytes(&namespace)?;
            if namespace == MCE {
                return Err(Error::UnsupportedMarkupCompatibility);
            }
            let namespace_id = match attribute.key.prefix() {
                Some(prefix) => namespace_ids.id(prefix.as_ref(), namespace)?,
                None => namespace_ids.id(&[], namespace)?,
            };
            let key = (
                namespace_id,
                copy_bytes(local.as_ref(), "attribute local name")?,
            );
            if !expanded.insert(key) {
                return invalid("duplicate expanded XML attribute name");
            }
        }
        Ok(())
    }

    fn required_attribute(
        &mut self,
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
        namespace: &[u8],
        local: &[u8],
        label: &str,
    ) -> Result<String> {
        self.optional_attribute(start, reader, namespace, local)?
            .ok_or_else(|| Error::InvalidStructure(format!("missing required {label}")))
    }

    fn optional_attribute(
        &mut self,
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
        namespace: &[u8],
        local: &[u8],
    ) -> Result<Option<String>> {
        for attribute in start.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| {
                Error::InvalidXml(format!("invalid ODS XML attribute: {error}"))
            })?;
            let (resolved, candidate) = reader.resolver().resolve_attribute(attribute.key);
            if namespace_bytes(&resolved)? == namespace && candidate.as_ref() == local {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                    .map_err(|error| {
                        Error::InvalidXml(format!("invalid ODS XML attribute value: {error}"))
                    })?;
                return self.copy_text(&value).map(Some);
            }
        }
        Ok(None)
    }

    fn has_attribute(
        &self,
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
        namespace: &[u8],
        local: &[u8],
    ) -> Result<bool> {
        for attribute in start.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| {
                Error::InvalidXml(format!("invalid ODS XML attribute: {error}"))
            })?;
            let (resolved, candidate) = reader.resolver().resolve_attribute(attribute.key);
            if namespace_bytes(&resolved)? == namespace && candidate.as_ref() == local {
                return Ok(true);
            }
        }
        Ok(false)
    }

    fn positive_attribute(
        &mut self,
        start: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
        namespace: &[u8],
        local: &[u8],
        label: &str,
    ) -> Result<Option<usize>> {
        self.optional_attribute(start, reader, namespace, local)?
            .map(|value| {
                value
                    .parse::<usize>()
                    .ok()
                    .filter(|value| *value > 0)
                    .ok_or_else(|| {
                        Error::InvalidStructure(format!("invalid positive {label} value '{value}'"))
                    })
            })
            .transpose()
    }

    fn copy_text(&mut self, value: &str) -> Result<String> {
        self.text_bytes = checked_add(self.text_bytes, value.len(), "retained text bytes")?;
        check_limit(
            LimitKind::TextBytes,
            self.text_bytes,
            self.limits.text_bytes,
        )?;
        let mut result = String::new();
        result
            .try_reserve_exact(value.len())
            .map_err(|_error| Error::AllocationFailed("retained text"))?;
        result.push_str(value);
        Ok(result)
    }

    fn inside_definition(&self) -> Option<usize> {
        self.active_definition
    }

    fn mark_opaque(&mut self) {
        if let Some(index) = self.inside_definition() {
            self.definitions[index].opaque_content = true;
        } else if self.catalog_active {
            self.catalog_opaque_content = true;
        } else if self
            .stack
            .iter()
            .any(|element| matches!(element, Element::Cell(true)))
        {
            self.binding_opaque_content = true;
        }
    }
}

#[derive(Debug)]
enum LayoutElement {
    Spreadsheet {
        start: usize,
        tag_end: usize,
        qname: String,
    },
    Catalog {
        start: usize,
        tag_end: usize,
        qname: String,
    },
    Definition {
        start: usize,
        tag_end: usize,
    },
    Prelude,
    Other,
}

fn scan_layout(source: &str, limits: Limits) -> Result<Layout> {
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut layout = Layout::default();
    let mut stack = Vec::new();
    stack
        .try_reserve(limits.depth.min(256))
        .map_err(|_error| Error::AllocationFailed("layout element stack"))?;
    layout
        .definitions
        .try_reserve(limits.definitions.min(256))
        .map_err(|_error| Error::AllocationFailed("definition layout"))?;
    let mut events = 0usize;
    let mut first_non_prelude = None;
    loop {
        let before = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::InvalidStructure("content-validation source offset exceeds usize".to_string())
        })?;
        let (resolved, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidXml(format!(
                "invalid ODS content-validation layout XML: {error}"
            ))
        })?;
        let namespace = namespace_kind(&resolved)?;
        let event = event.into_owned();
        drop(resolved);
        let after = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::InvalidStructure("content-validation source offset exceeds usize".to_string())
        })?;
        events = checked_add(events, 1, "layout event counter")?;
        check_limit(LimitKind::Events, events, limits.events)?;
        match event {
            Event::Start(start) => {
                let element = classify_layout_start(
                    namespace,
                    &start,
                    before,
                    after,
                    &stack,
                    &mut layout,
                    &mut first_non_prelude,
                )?;
                stack
                    .try_reserve(1)
                    .map_err(|_error| Error::AllocationFailed("layout element stack"))?;
                stack.push(element);
            },
            Event::Empty(start) => {
                let element = classify_layout_start(
                    namespace,
                    &start,
                    before,
                    after,
                    &stack,
                    &mut layout,
                    &mut first_non_prelude,
                )?;
                finish_layout_element(
                    element,
                    before,
                    after,
                    true,
                    &mut layout,
                    &mut first_non_prelude,
                )?;
            },
            Event::End(_) => {
                let element = stack.pop().ok_or_else(|| {
                    Error::InvalidStructure("content-validation layout stack underflow".to_string())
                })?;
                finish_layout_element(
                    element,
                    before,
                    after,
                    false,
                    &mut layout,
                    &mut first_non_prelude,
                )?;
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return invalid("content-validation layout rejects DTDs and entity references");
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    if !stack.is_empty() || layout.spreadsheet.is_none() {
        return invalid("content-validation layout lacks a complete spreadsheet owner");
    }
    Ok(layout)
}

fn classify_layout_start(
    namespace: NamespaceKind,
    start: &BytesStart<'_>,
    source_start: usize,
    tag_end: usize,
    stack: &[LayoutElement],
    layout: &mut Layout,
    first_non_prelude: &mut Option<usize>,
) -> Result<LayoutElement> {
    let local = start.local_name();
    let local = local.as_ref();
    if namespace == NamespaceKind::Office && local == b"spreadsheet" {
        let qname = copy_utf8(start.name().as_ref(), "spreadsheet qualified name")?;
        layout.catalog_insertion = tag_end;
        return Ok(LayoutElement::Spreadsheet {
            start: source_start,
            tag_end,
            qname,
        });
    }
    if matches!(stack.last(), Some(LayoutElement::Spreadsheet { .. })) {
        if namespace == NamespaceKind::Table && local == b"content-validations" {
            return Ok(LayoutElement::Catalog {
                start: source_start,
                tag_end,
                qname: copy_utf8(start.name().as_ref(), "catalog qualified name")?,
            });
        }
        if namespace == NamespaceKind::Table
            && matches!(local, b"tracked-changes" | b"calculation-settings")
        {
            return Ok(LayoutElement::Prelude);
        }
        if first_non_prelude.is_none() {
            *first_non_prelude = Some(source_start);
        }
    }
    if namespace == NamespaceKind::Table
        && local == b"content-validation"
        && matches!(stack.last(), Some(LayoutElement::Catalog { .. }))
    {
        return Ok(LayoutElement::Definition {
            start: source_start,
            tag_end,
        });
    }
    Ok(LayoutElement::Other)
}

fn finish_layout_element(
    element: LayoutElement,
    close_start: usize,
    end: usize,
    empty: bool,
    layout: &mut Layout,
    first_non_prelude: &mut Option<usize>,
) -> Result<()> {
    match element {
        LayoutElement::Spreadsheet {
            start,
            tag_end,
            qname,
        } => {
            let insertion = first_non_prelude.unwrap_or(close_start);
            if layout.catalog.is_none() {
                layout.catalog_insertion = insertion;
            }
            layout.spreadsheet = Some(OwnerSpan {
                range: start..end,
                start_tag: start..tag_end,
                close_start,
                qname,
                empty,
            });
        },
        LayoutElement::Catalog {
            start,
            tag_end,
            qname,
        } => {
            layout.catalog = Some(OwnerSpan {
                range: start..end,
                start_tag: start..tag_end,
                close_start,
                qname,
                empty,
            });
        },
        LayoutElement::Definition { start, tag_end } => {
            layout
                .definitions
                .try_reserve(1)
                .map_err(|_error| Error::AllocationFailed("definition layout"))?;
            layout.definitions.push(DefinitionSpan {
                range: start..end,
                start_tag: start..tag_end,
                empty,
            });
        },
        LayoutElement::Prelude => {
            if first_non_prelude.is_none() && layout.catalog.is_none() {
                layout.catalog_insertion = end;
            }
        },
        LayoutElement::Other => {},
    }
    Ok(())
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    let value = namespace_bytes(resolved)?;
    if value == MCE {
        return Err(Error::UnsupportedMarkupCompatibility);
    }
    Ok(match value {
        OFFICE => NamespaceKind::Office,
        TABLE => NamespaceKind::Table,
        TEXT => NamespaceKind::Text,
        SCRIPT => NamespaceKind::Script,
        PRESENTATION => NamespaceKind::Presentation,
        _ => NamespaceKind::Other,
    })
}

fn namespace_bytes<'a>(resolved: &'a ResolveResult<'a>) -> Result<&'a [u8]> {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => Ok(value),
        ResolveResult::Unbound => Ok(&[]),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidStructure(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn known_definition_element(namespace: NamespaceKind, local: &[u8]) -> bool {
    matches!(
        (namespace, local),
        (
            NamespaceKind::Table,
            b"help-message" | b"error-message" | b"error-macro"
        ) | (NamespaceKind::Office, b"event-listeners")
            | (
                NamespaceKind::Text,
                b"p" | b"span" | b"s" | b"tab" | b"line-break"
            )
            | (NamespaceKind::Script, b"event-listener")
            | (NamespaceKind::Presentation, b"event-listener" | b"sound")
    )
}

fn has_unknown_definition_attributes(
    namespace: NamespaceKind,
    local: &[u8],
    start: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<bool> {
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| Error::InvalidXml(format!("invalid ODS XML attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (resolved, attribute_local) = reader.resolver().resolve_attribute(attribute.key);
        let attribute_namespace = namespace_bytes(&resolved)?;
        let known = match (namespace, local) {
            (NamespaceKind::Table, b"content-validation") => {
                attribute_namespace == TABLE
                    && matches!(
                        attribute_local.as_ref(),
                        b"name"
                            | b"condition"
                            | b"base-cell-address"
                            | b"allow-empty-cell"
                            | b"display-list"
                    )
            },
            (NamespaceKind::Table, b"help-message" | b"error-message") => {
                attribute_namespace == TABLE
                    && matches!(
                        attribute_local.as_ref(),
                        b"title" | b"display" | b"message-type"
                    )
            },
            (NamespaceKind::Table, b"error-macro") => {
                attribute_namespace == TABLE && attribute_local.as_ref() == b"execute"
            },
            (NamespaceKind::Text, b"s") => {
                attribute_namespace == TEXT && attribute_local.as_ref() == b"c"
            },
            _ => false,
        };
        if !known {
            return Ok(true);
        }
    }
    Ok(false)
}

fn has_non_namespace_attributes(start: &BytesStart<'_>) -> bool {
    start
        .attributes()
        .filter_map(std::result::Result::ok)
        .any(|attribute| !is_namespace_declaration(attribute.key.as_ref()))
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn parse_bool(value: &str, label: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid Boolean {label} value '{value}'")),
    }
}

fn direct(
    stack: &[Element],
    predicate: impl FnOnce(&Element) -> bool,
    value: Element,
    message: &str,
) -> Result<Element> {
    if stack.last().is_some_and(predicate) {
        Ok(value)
    } else {
        invalid(message)
    }
}

fn checked_add(left: usize, right: usize, label: &str) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| Error::InvalidStructure(format!("{label} overflow")))
}

fn check_limit(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        Err(Error::LimitExceeded {
            kind,
            observed,
            maximum,
        })
    } else {
        Ok(())
    }
}

fn copy_bytes(value: &[u8], owner: &'static str) -> Result<Vec<u8>> {
    let mut result = Vec::new();
    result
        .try_reserve_exact(value.len())
        .map_err(|_error| Error::AllocationFailed(owner))?;
    result.extend_from_slice(value);
    Ok(result)
}

fn copy_string(value: &str, owner: &'static str) -> Result<String> {
    let mut result = String::new();
    result
        .try_reserve_exact(value.len())
        .map_err(|_error| Error::AllocationFailed(owner))?;
    result.push_str(value);
    Ok(result)
}

fn copy_optional_string(value: Option<&str>, owner: &'static str) -> Result<Option<String>> {
    value.map(|value| copy_string(value, owner)).transpose()
}

fn copy_utf8(value: &[u8], owner: &'static str) -> Result<String> {
    let value = core::str::from_utf8(value)
        .map_err(|_error| Error::InvalidXml(format!("{owner} is not UTF-8")))?;
    copy_string(value, owner)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidStructure(message.into()))
}

#[cfg(test)]
mod tests {
    use super::AttributeNamespaceIds;

    #[test]
    fn repeated_prefix_avoids_rehashing_namespace_uri() {
        let namespace = [b'n'; 128 * 1024];
        let mut ids = AttributeNamespaceIds::with_capacity(2_000).unwrap();
        let first = ids.id(b"vendor", &namespace).unwrap();
        for _ in 1..2_000 {
            assert_eq!(ids.id(b"vendor", &namespace).unwrap(), first);
        }
        assert_eq!(ids.namespace_content_lookups(), 1);

        assert_eq!(ids.id(b"alias", &namespace).unwrap(), first);
        assert_eq!(ids.namespace_content_lookups(), 2);
    }
}
