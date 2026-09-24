//! Canonical XLDM 140 table-local identity projection.
//!
//! MS-XLDM §2.2.3.7 gives generated names and containing dimension folders,
//! while §§2.5.2.1, 2.5.2.3, and 2.5.2.4 give the XML table, column, and
//! relationship objects. Section 2.6 supplies the corresponding Dimension
//! object and Attribute IDs. This module keeps those identities qualified by
//! their containing table. It deliberately does not flatten columns or
//! relationship endpoints into global name sets.
//!
//! The projection is source-bound and retains the ownership edges required by
//! a writer. Native/generated values, relationship indexes, file lists,
//! calculated members, and time-group dependencies are validated before a
//! change is admitted. Structural operations whose complete inverse is not
//! implemented remain explicit refusals; the typed table XML-name operation
//! below is the first supported descriptor edit.

use std::collections::{HashMap, HashSet, TryReserveError};
use std::error::Error as StdError;
use std::fmt;

use super::generated::{
    SystemGeneratedKind, SystemGeneratedModel, parse_system_generated_file,
    validate_system_generated_files,
};
use super::metadata::{MetadataFile, MetadataFileKind, MetadataModel, MetadataObject};
use super::native::NativeModel;
use super::olap::{OlapDocument, OlapError, OlapModel, OlapObjectKind};
use super::{
    FileEntry, FileGroupClass, GeneratedNameKind, Storage, StorageProfile, classify_generated_path,
};

const MAX_IDENTITY_ITEMS: usize = 500_000;
/// Aggregate bytes charged before any retained identity strings or indexes are
/// allocated.  The source models are already bounded by their section
/// parsers; this separate budget prevents a large number of individually
/// valid names from making the projection clone an unbounded graph.
const MAX_IDENTITY_BYTES: usize = 64 * 1024 * 1024;

/// Failure classes exposed by the bounded XLDM table-rename seam.
///
/// The ordinary identity APIs retain their historical [`OlapError`] return
/// type.  The caller-limited rename has this small diagnostic vocabulary so a
/// format host can preserve its own invalid-format, quota, allocation, and
/// unsupported-feature error contracts without parsing display strings.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Xldm140RenameErrorKind {
    /// The source, identity closure, XML, or requested operation is invalid.
    InvalidSource,
    /// A hard XLDM or caller-supplied bound was exceeded.
    LimitExceeded,
    /// A bounded vector or output buffer could not be reserved.
    Allocation,
    /// The source profile or requested operation is outside this owner.
    Unsupported,
}

/// Typed diagnostic returned by the caller-limited XLDM table rename.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140RenameError {
    kind: Xldm140RenameErrorKind,
    message: String,
    limit_bounds: Option<(usize, usize)>,
    /// The owner-specific quota label when the failure came from a nested
    /// proof.  Caller/output limits are deliberately left unset so hosts can
    /// apply their format-level fallback label.
    limit_resource: Option<&'static str>,
    allocation_resource: Option<String>,
    allocation_source: Option<TryReserveError>,
}

impl Xldm140RenameError {
    /// Return the stable failure class for host error mapping.
    #[must_use]
    pub const fn kind(&self) -> Xldm140RenameErrorKind {
        self.kind
    }

    /// Return the exact observed and maximum values for a bounded failure.
    ///
    /// Some structural overflow checks have no representable observed value;
    /// those retain `None` while preserving their historical display text.
    #[must_use]
    pub const fn limit_bounds(&self) -> Option<(usize, usize)> {
        self.limit_bounds
    }

    /// Return the observed value for a bounded failure, when available.
    #[must_use]
    pub const fn limit_actual(&self) -> Option<usize> {
        match self.limit_bounds {
            Some(bounds) => Some(bounds.0),
            None => None,
        }
    }

    /// Return the configured maximum for a bounded failure, when available.
    #[must_use]
    pub const fn limit_maximum(&self) -> Option<usize> {
        match self.limit_bounds {
            Some(bounds) => Some(bounds.1),
            None => None,
        }
    }

    /// Return the owner-specific resource label for a bounded failure, when
    /// the failure originated in a nested proof with one.  Caller-supplied
    /// output limits intentionally return `None`; the host owns that label.
    #[must_use]
    pub const fn limit_resource(&self) -> Option<&'static str> {
        self.limit_resource
    }

    /// Return the neutral resource label for a bounded allocation failure.
    ///
    /// The original reservation source, when available, is exposed separately
    /// through [`Self::allocation_source`].
    #[must_use]
    pub fn allocation_resource(&self) -> Option<&str> {
        self.allocation_resource.as_deref()
    }

    /// Return the original bounded reservation failure when the rename path
    /// received one from `try_reserve`.  Proof diagnostics that only carry a
    /// neutral detail string intentionally return `None`.
    #[must_use]
    pub fn allocation_source(&self) -> Option<&TryReserveError> {
        self.allocation_source.as_ref()
    }

    fn invalid(message: impl Into<String>) -> Self {
        Self {
            kind: Xldm140RenameErrorKind::InvalidSource,
            message: message.into(),
            limit_bounds: None,
            limit_resource: None,
            allocation_resource: None,
            allocation_source: None,
        }
    }

    fn limit(message: impl Into<String>) -> Self {
        Self::limit_with_bounds(message, None)
    }

    fn limit_with_bounds(message: impl Into<String>, bounds: Option<(usize, usize)>) -> Self {
        Self::limit_with_resource(message, bounds, None)
    }

    fn limit_with_resource(
        message: impl Into<String>,
        bounds: Option<(usize, usize)>,
        resource: Option<&'static str>,
    ) -> Self {
        Self {
            kind: Xldm140RenameErrorKind::LimitExceeded,
            message: message.into(),
            limit_bounds: bounds,
            limit_resource: resource,
            allocation_resource: None,
            allocation_source: None,
        }
    }

    fn unsupported(message: impl Into<String>) -> Self {
        Self {
            kind: Xldm140RenameErrorKind::Unsupported,
            message: message.into(),
            limit_bounds: None,
            limit_resource: None,
            allocation_resource: None,
            allocation_source: None,
        }
    }

    fn allocation(resource: impl Into<String>, error: TryReserveError) -> Self {
        let resource = resource.into();
        Self {
            kind: Xldm140RenameErrorKind::Allocation,
            message: format!("could not reserve XLDM rename {resource}: {error}"),
            limit_bounds: None,
            limit_resource: None,
            allocation_resource: Some(resource),
            allocation_source: Some(error),
        }
    }

    fn allocation_detail(resource: impl Into<String>, detail: impl fmt::Display) -> Self {
        let resource = resource.into();
        Self {
            kind: Xldm140RenameErrorKind::Allocation,
            message: format!("could not reserve XLDM rename {resource}: {detail}"),
            limit_bounds: None,
            limit_resource: None,
            allocation_resource: Some(resource),
            allocation_source: None,
        }
    }

    fn from_codec(error: crate::error::Error) -> Self {
        match error {
            crate::error::Error::Invalid(message) | crate::error::Error::Xml(message) => {
                Self::invalid(message)
            },
            crate::error::Error::Unsupported { feature } => Self {
                kind: Xldm140RenameErrorKind::Unsupported,
                message: format!("unsupported XLDM rename operation: {feature}"),
                limit_bounds: None,
                limit_resource: None,
                allocation_resource: None,
                allocation_source: None,
            },
            crate::error::Error::Allocation { resource, source } => {
                Self::allocation(resource, source)
            },
        }
    }

    fn from_variable_rewrite(error: super::codec::VariableRewriteError) -> Self {
        match error {
            super::codec::VariableRewriteError::Codec(error) => Self::from_codec(error),
            super::codec::VariableRewriteError::CallerLimit { actual, maximum } => {
                Self::limit_with_bounds(
                    format!(
                        "rewritten storage caller output bytes {actual} exceed limit {maximum}"
                    ),
                    Some((actual, maximum)),
                )
            },
        }
    }

    fn from_olap_proof(error: super::olapproof::OlapProofError) -> Self {
        match error {
            super::olapproof::OlapProofError::UnsupportedProfile => Self {
                kind: Xldm140RenameErrorKind::Unsupported,
                message: error.to_string(),
                limit_bounds: None,
                limit_resource: None,
                allocation_resource: None,
                allocation_source: None,
            },
            super::olapproof::OlapProofError::LimitExceeded {
                resource,
                actual,
                maximum,
            } => Self::limit_with_resource(
                error.to_string(),
                Some((actual, maximum)),
                Some(resource),
            ),
            super::olapproof::OlapProofError::Allocation { resource, detail } => {
                Self::allocation_detail(resource, detail)
            },
            super::olapproof::OlapProofError::Invalid { .. }
            | super::olapproof::OlapProofError::Unproven { .. } => Self::invalid(error.to_string()),
        }
    }

    fn into_olap_error(self) -> OlapError {
        OlapError::new(self.message)
    }
}

impl fmt::Display for Xldm140RenameError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl StdError for Xldm140RenameError {}

impl From<OlapError> for Xldm140RenameError {
    fn from(error: OlapError) -> Self {
        Self::invalid(error.to_string())
    }
}

/// One canonical table-local column identity.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140ColumnIdentity {
    /// The TableID from the validated containing `.dim` folder.
    pub table_id: String,
    /// The `XMRawColumn@name`/`ColID` value from the table metadata object.
    pub column_id: String,
    /// The explicit section-2.6 Dimension Attribute name, when present.
    ///
    /// This stays separate from the immutable raw-column identity so a
    /// workbook `columnName` can differ from `columnId` only when the source
    /// carries an explicit mapping. Sources that omit this optional base
    /// field conservatively use `column_id` for both descriptor fields.
    pub column_name: Option<String>,
    /// OLE DB DBType from the validated XMColumnStats member, when retained
    /// by the source projection. Time-group source binding requires Date (7)
    /// and calculated members require a validated integral result type.
    pub db_type: Option<u16>,
    /// The complete validated `XMRawColumn/Settings` value. Keeping the raw
    /// bitfield alongside the derived flag makes the calculated-column proof
    /// auditable: no adapter may replace a source column merely because a
    /// filename or display name happens to match.
    pub settings: u64,
    /// The partition/data-object storage name qualified to the metadata file.
    pub data_path: String,
    /// Whether the XMRawColumn is a calculated column.
    ///
    /// MS-XLDM 2.5.2.3 defines this on the `Settings` value (the calculated
    /// type is `0x2`, and the calculated-column modifier is `0x800`).  It is
    /// retained in the identity projection so an outer workbook reference
    /// cannot accidentally bind a source column to a generated column.
    pub is_calculated: bool,
}

/// A qualified column binding admitted by the complete XLDM identity closure.
///
/// The XLSX/XLSB workbook descriptor carries a user-facing column name and an
/// immutable column identifier. Version-140 XLDM exposes the raw-column
/// instance name as its table-local immutable identity. When the section-2.6
/// Dimension has an explicit Attribute name, that name is used for the
/// descriptor's display `columnName`; otherwise the raw identity is required
/// for both fields. A source whose independent mapping is not present in
/// XLDM is rejected instead of being guessed from a filename.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140ColumnBinding {
    /// Qualified XLDM TableID.
    pub table_id: String,
    /// Qualified XLDM XML table name.
    pub table_name: String,
    /// Workbook descriptor column name.
    pub column_name: String,
    /// Workbook descriptor immutable column identifier.
    pub column_id: String,
    /// Qualified XLDM column data member.
    pub data_path: String,
    /// Whether the bound XLDM raw column is calculated.
    pub is_calculated: bool,
    /// Validated OLE DB type retained from `XMColumnStats/DBType`.
    pub db_type: Option<u16>,
    /// The source `XMRawColumn/Settings` bitfield used for the calculated
    /// derivation classification.
    pub settings: u64,
    /// The typed outer time-grouping derivation, when the adapter supplied
    /// one.  This is retained as evidence of the content-type check rather
    /// than treating the descriptor token as an unvalidated string.
    pub content_type: Option<Xldm140TimeGroupingContentType>,
}

/// A validated source/calculated-column binding for one model time grouping.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140TimeGroupingBinding {
    /// Qualified XLDM TableID owning the grouping.
    pub table_id: String,
    /// XML table name used by the workbook descriptor.
    pub table_name: String,
    /// The non-calculated source column.
    pub source: Xldm140ColumnBinding,
    /// Calculated columns generated for the grouping.
    pub calculated_columns: Vec<Xldm140ColumnBinding>,
}

/// The nine content-type values defined by MS-XLSX
/// `ST_ModelTimeGroupingContentType` and MS-XLSB 2.4.716. Keeping the
/// granularity in the neutral proof prevents a typed XLSX/XLSB adapter from
/// validating only the outer token while ignoring the inner calculated
/// column closure.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum Xldm140TimeGroupingContentType {
    Years,
    Quarters,
    MonthsIndex,
    Months,
    DaysIndex,
    Days,
    Hours,
    Minutes,
    Seconds,
}

impl Xldm140TimeGroupingContentType {
    #[must_use]
    pub const fn code(self) -> u8 {
        match self {
            Self::Years => 0,
            Self::Quarters => 1,
            Self::MonthsIndex => 2,
            Self::Months => 3,
            Self::DaysIndex => 4,
            Self::Days => 5,
            Self::Hours => 6,
            Self::Minutes => 7,
            Self::Seconds => 8,
        }
    }
}

/// One canonical table identity assembled from section 2.2, 2.5, and 2.6.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140TableIdentity {
    /// The TableID from the containing dimension folder.
    pub table_id: String,
    /// The `XMSimpleTable@name` value.
    pub xml_name: String,
    /// The source `.tbl.xml` member that supplied this identity.
    pub metadata_path: String,
    /// The section 2.6 Dimension `ID`, retained separately from the generated
    /// TableID and XML table name.
    pub dimension_object_id: String,
    /// All section 2.6 Dimension Attribute IDs retained for this table.
    pub attribute_ids: Vec<String>,
    /// Columns qualified by this table.
    pub columns: Vec<Xldm140ColumnIdentity>,
}

/// One relationship identity with an explicit containing (foreign) table.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140RelationshipIdentity {
    /// The table whose dimension folder contains this relationship metadata.
    /// MS-XLDM has no `ForeignTable` XML property; containment supplies it.
    pub containing_table: String,
    /// The `XMRelationship@name` metadata value, retained separately from
    /// generated storage identity.
    pub relationship_name: Option<String>,
    /// The generated `RelId` token from the relationship metadata filename.
    pub relationship_id: String,
    /// The source relationship `.tbl.xml` member.
    pub metadata_path: String,
    /// `XMRelationship/PrimaryTable` (the one side).
    pub primary_table: String,
    /// `XMRelationship/PrimaryColumn` (the one-side column).
    pub primary_column: String,
    /// `XMRelationship/ForeignColumn` (the containing table column).
    pub foreign_column: String,
    /// The generated relationship-index key that must be closed by section 2.4.
    pub expected_index_key: String,
}

/// Canonical table-local XLDM 140 identities.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140IdentityProjection {
    pub tables: Vec<Xldm140TableIdentity>,
    pub relationships: Vec<Xldm140RelationshipIdentity>,
}

impl Xldm140IdentityProjection {
    /// Find a table by its qualified TableID.
    #[must_use]
    pub fn table(&self, table_id: &str) -> Option<&Xldm140TableIdentity> {
        self.tables.iter().find(|table| table.table_id == table_id)
    }

    /// Find a column only within its owning table.
    #[must_use]
    pub fn column(&self, table_id: &str, column_id: &str) -> Option<&Xldm140ColumnIdentity> {
        self.table(table_id)?
            .columns
            .iter()
            .find(|column| column.column_id == column_id)
    }
}

/// The section owning one admitted XLDM storage member.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Xldm140MemberSection {
    /// Section 2.2 control members (partitions, backup log, or key material).
    Outer,
    /// Section 2.5 table metadata.
    Metadata,
    /// Section 2.3 native data.
    Native,
    /// Section 2.4 generated data.
    Generated,
    /// A member admitted by both the native and generated views.
    NativeAndGenerated,
    /// Section 2.6 OLAP definitions and information files.
    Olap,
    /// A retained member not understood by the typed closure.
    Other,
}

/// A source-indexed member in an admitted XLDM closure.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Xldm140ClosureMember {
    /// Index into the file directory returned by [`Xldm140Closure::storage`].
    pub storage_index: usize,
    /// Typed owner, or [`Xldm140MemberSection::Other`] for an opaque member.
    pub section: Xldm140MemberSection,
}

/// A replacement for one already admitted XLDM allocation.
///
/// Replacements are deliberately keyed by the section 2.2 `StoragePath` and
/// contain the payload without its four-byte CRC marker. The closure writer
/// accepts only same-sized replacements whose parsed native/generated semantic
/// payload is unchanged, then reparses the complete cross-section closure.
/// This keeps all directory offsets, allocation gaps, and backup-log sizes
/// source-bound; native/generated value edits and member add/remove/rename
/// operations require a separate closure-aware writer.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Xldm140FileReplacement<'a> {
    pub storage_path: &'a str,
    pub payload: &'a [u8],
}

/// Bytes returned by a source-bound XLDM patch.
///
/// Exact no-ops borrow the inspected source directly.  A changed patch owns
/// its candidate allocation.  This distinction keeps a no-op from copying an
/// opaque model merely to satisfy a writer API.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum Xldm140PatchBytes<'a> {
    Borrowed(&'a [u8]),
    Owned(Vec<u8>),
}

impl Xldm140PatchBytes<'_> {
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Borrowed(bytes) => bytes,
            Self::Owned(bytes) => bytes,
        }
    }

    #[must_use]
    pub const fn is_borrowed(&self) -> bool {
        matches!(self, Self::Borrowed(_))
    }
}

/// A reversible source-checked XLDM rewrite.
///
/// The patch compares the complete source bytes before applying.  Projected
/// names alone are insufficient as a stale guard because unknown members and
/// lexical XML changes can be semantically invisible to the identity view.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Xldm140Patch<'source> {
    before: &'source [u8],
    after: Xldm140PatchBytes<'source>,
}

impl<'source> Xldm140Patch<'source> {
    #[must_use]
    pub fn before(&self) -> &'source [u8] {
        self.before
    }

    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after.as_bytes()
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.after.is_borrowed()
    }

    /// Apply against an exact source byte slice, borrowing the target on a
    /// no-op and allocating only for a changed patch.
    pub fn apply<'target>(
        &self,
        target: &'target [u8],
    ) -> Result<std::borrow::Cow<'target, [u8]>, OlapError> {
        if target != self.before {
            return Err(OlapError::new(
                "XLDM patch source is stale or has changed lexical bytes",
            ));
        }
        if self.is_noop() {
            Ok(std::borrow::Cow::Borrowed(target))
        } else {
            Ok(std::borrow::Cow::Owned(self.after.as_bytes().to_vec()))
        }
    }

    /// Return a source-bound inverse view. Applying it checks the exact
    /// candidate bytes and restores the original borrowed source.
    #[must_use]
    pub fn inverse(&self) -> Xldm140InversePatch<'_, 'source> {
        Xldm140InversePatch {
            expected: self.after(),
            restore: self.before,
        }
    }
}

/// The inverse half of [`Xldm140Patch`].
#[derive(Debug, Eq, PartialEq)]
pub struct Xldm140InversePatch<'patch, 'source> {
    expected: &'patch [u8],
    restore: &'source [u8],
}

impl<'patch, 'source> Xldm140InversePatch<'patch, 'source> {
    pub fn apply(&self, target: &[u8]) -> Result<&'source [u8], OlapError> {
        if target != self.expected {
            return Err(OlapError::new(
                "XLDM inverse patch candidate is stale or has changed lexical bytes",
            ));
        }
        Ok(self.restore)
    }
}

/// A source-bound, cross-section Xldm140 proof graph.
///
/// The graph retains the caller's inspected storage by borrow and stores only
/// bounded indexes into its directory. Unknown members remain visible through
/// [`Self::unknown_members`], so callers can preserve them while refusing any
/// structural edit whose closure could be affected.
pub struct Xldm140Closure<'storage, 'source> {
    storage: &'storage Storage<'source>,
    metadata: Option<&'source MetadataModel<'source>>,
    native: Option<&'source NativeModel<'source>>,
    generated: Option<&'source SystemGeneratedModel<'source>>,
    projection: Xldm140IdentityProjection,
    members: Vec<Xldm140ClosureMember>,
    unknown_members: Vec<usize>,
}

impl<'storage, 'source> Xldm140Closure<'storage, 'source> {
    /// Return the exact inspected storage to which this proof is bound.
    #[must_use]
    pub fn storage(&self) -> &'storage Storage<'source> {
        self.storage
    }

    /// Return the qualified table/column/relationship projection.
    #[must_use]
    pub fn projection(&self) -> &Xldm140IdentityProjection {
        &self.projection
    }

    /// Return every directory member and its typed section owner.
    #[must_use]
    pub fn members(&self) -> &[Xldm140ClosureMember] {
        &self.members
    }

    /// Return indexes of members retained but not understood by the proof.
    #[must_use]
    pub fn unknown_members(&self) -> &[usize] {
        &self.unknown_members
    }

    /// Whether every directory member belongs to a validated typed section.
    #[must_use]
    pub fn is_complete(&self) -> bool {
        self.unknown_members.is_empty()
    }

    /// Resolve a grouping and prove each calculated column's content type
    /// against the admitted XLDM column closure. The neutral API accepts the
    /// numeric enum as a typed value rather than trusting an adapter string.
    /// The caller's descriptor list supplies the source-to-calculated edge;
    /// the source proof checks that edge against distinct table-local column
    /// identities, calculated-column settings, integral
    /// `XMColumnStats/DBType`, required hierarchy mappings, admitted native
    /// data, and section-2.6 Attribute closure. XLDM 140 stores no
    /// model-time-grouping formula, so this method never evaluates or invents
    /// one; a calculated derivation absent from those source facts is refused.
    pub fn bind_time_grouping_with_content_types(
        &self,
        table_name: &str,
        source_name: &str,
        source_id: &str,
        calculated: &[(&str, &str, Xldm140TimeGroupingContentType)],
    ) -> Result<Xldm140TimeGroupingBinding, OlapError> {
        self.require_complete_writable_olap_proof()?;
        let calculated = calculated
            .iter()
            .map(|&(name, id, content_type)| (name, id, Some(content_type)))
            .collect::<Vec<_>>();
        self.bind_time_grouping_inner(table_name, source_name, source_id, &calculated)
    }

    /// A semantic projection may be useful to inspection callers, but it is
    /// insufficient evidence for a source-bound workbook write.  A typed
    /// time-grouping binding therefore requires all three retained section
    /// models and the standalone OLAP/file-group proof.  In particular, a
    /// hand-built projection or a source whose BackupLog does not advertise
    /// OLAP cannot accidentally be treated as a complete writable closure.
    fn require_complete_writable_olap_proof(&self) -> Result<(), OlapError> {
        if !self.storage.backup_log.is_olap {
            return Err(OlapError::new(
                "typed time-grouping binding requires an OLAP-advertised XLDM closure",
            ));
        }
        let metadata = self
            .metadata
            .ok_or_else(|| OlapError::new("typed time-grouping binding lacks section 2.5 proof"))?;
        let native = self
            .native
            .ok_or_else(|| OlapError::new("typed time-grouping binding lacks section 2.3 proof"))?;
        let generated = self
            .generated
            .ok_or_else(|| OlapError::new("typed time-grouping binding lacks section 2.4 proof"))?;
        let olap = super::olap::inspect(self.storage, metadata)
            .map_err(|error| OlapError::new(format!("section 2.6 proof is invalid: {error}")))?;
        let proof = super::olapproof::prove_xldm140_olap(
            self.storage,
            metadata,
            &olap,
            super::olapproof::OlapProofLimits::default(),
        )
        .map_err(|error| OlapError::new(format!("section 2.6 OLAP proof failed: {error}")))?;
        if !proof.is_complete() {
            return Err(OlapError::new(
                "typed time-grouping binding requires a complete section 2.6 OLAP proof",
            ));
        }
        // Keep the arguments live in this guard: a future proof implementation
        // must not silently drop native/generated closure ownership while the
        // public binding remains source-bound.
        if native.files.is_empty()
            && generated.files.is_empty()
            && !self.projection.tables.is_empty()
        {
            return Err(OlapError::new(
                "typed time-grouping binding lacks admitted native/generated members",
            ));
        }
        Ok(())
    }

    fn bind_time_grouping_inner(
        &self,
        table_name: &str,
        source_name: &str,
        source_id: &str,
        calculated: &[(&str, &str, Option<Xldm140TimeGroupingContentType>)],
    ) -> Result<Xldm140TimeGroupingBinding, OlapError> {
        if !self.is_complete() {
            return Err(OlapError::new(
                "cannot bind a time grouping in an XLDM closure containing unknown members",
            ));
        }
        let mut tables = self
            .projection
            .tables
            .iter()
            .filter(|table| table.xml_name == table_name);
        let table = tables.next().ok_or_else(|| {
            OlapError::new(format!(
                "time grouping table {table_name} is absent from the XLDM identity projection"
            ))
        })?;
        if tables.next().is_some() {
            return Err(OlapError::new(format!(
                "time grouping table {table_name} is ambiguous in the XLDM identity projection"
            )));
        }
        let source = bind_column(table, source_name, source_id, false, "source", None)?;
        let source_identity = table
            .columns
            .iter()
            .find(|column| column.column_id == source_id)
            .ok_or_else(|| {
                OlapError::new(format!(
                "time grouping source column {source_id} disappeared from the identity projection"
                ))
            })?;
        if self.generated.is_some() {
            self.require_position_to_identifier_mapping(table, source_identity, "source")?;
        }
        if source_identity.db_type != Some(7) {
            return Err(OlapError::new(format!(
                "time grouping source column {source_id} is not a validated OLE DB Date column"
            )));
        }
        if calculated.is_empty() {
            return Err(OlapError::new(
                "time grouping must contain at least one calculated column",
            ));
        }
        let mut seen = HashSet::new();
        seen.try_reserve(calculated.len()).map_err(|error| {
            OlapError::new(format!("cannot reserve time grouping column IDs: {error}"))
        })?;
        let mut calculated_columns = Vec::new();
        calculated_columns
            .try_reserve(calculated.len())
            .map_err(|error| {
                OlapError::new(format!("cannot reserve time grouping bindings: {error}"))
            })?;
        let mut seen_content_types = HashSet::new();
        seen_content_types
            .try_reserve(calculated.len())
            .map_err(|error| {
                OlapError::new(format!(
                    "cannot reserve time grouping content types: {error}"
                ))
            })?;
        for &(column_name, column_id, content_type) in calculated {
            if !seen.insert(column_id) {
                return Err(OlapError::new(format!(
                    "time grouping has duplicate calculated column ID {column_id}"
                )));
            }
            if column_id == source.column_id {
                return Err(OlapError::new(
                    "time grouping calculated column aliases its source column",
                ));
            }
            if let Some(content_type) = content_type {
                if !seen_content_types.insert(content_type.code()) {
                    return Err(OlapError::new(format!(
                        "time grouping has duplicate content type {}",
                        content_type.code()
                    )));
                }
            }
            let binding = bind_column(
                table,
                column_name,
                column_id,
                true,
                "calculated",
                content_type,
            )?;
            let calculated_identity = table
                .columns
                .iter()
                .find(|column| column.column_id == column_id)
                .ok_or_else(|| {
                    OlapError::new(format!(
                        "time grouping calculated column {column_id} disappeared from the identity projection"
                    ))
                })?;
            if self.generated.is_some() {
                self.require_position_to_identifier_mapping(
                    table,
                    calculated_identity,
                    "calculated",
                )?;
            }
            if !matches!(calculated_identity.db_type, Some(2 | 3 | 18 | 19 | 20 | 21)) {
                return Err(OlapError::new(format!(
                    "time grouping calculated column {column_id} has no validated integral derivation type"
                )));
            }
            calculated_columns.push(binding);
        }
        Ok(Xldm140TimeGroupingBinding {
            table_id: table.table_id.clone(),
            table_name: table.xml_name.clone(),
            source,
            calculated_columns,
        })
    }

    fn require_position_to_identifier_mapping(
        &self,
        table: &Xldm140TableIdentity,
        column: &Xldm140ColumnIdentity,
        role: &str,
    ) -> Result<(), OlapError> {
        let generated = self.generated.ok_or_else(|| {
            OlapError::new(format!(
                "time grouping {role} column {} has no admitted section 2.4 model",
                column.column_id
            ))
        })?;
        let (folder, basename) = column.data_path.rsplit_once('/').ok_or_else(|| {
            OlapError::new(format!(
                "time grouping {role} column {} has an unqualified data path",
                column.column_id
            ))
        })?;
        let ordinal = basename.split_once('.').map_or(basename, |value| value.0);
        let expected = format!(
            "{folder}/{ordinal}.H${}${}.POS_TO_ID.0.idf",
            table.table_id, column.column_id
        );
        let mut matches = generated.files.iter().filter(|file| {
            file.kind == SystemGeneratedKind::PositionToIdentifier && file.storage_path == expected
        });
        if matches.next().is_none() {
            return Err(OlapError::new(format!(
                "time grouping {role} column {} has no admitted POS_TO_ID member {expected}",
                column.column_id
            )));
        }
        if matches.next().is_some() {
            return Err(OlapError::new(format!(
                "time grouping {role} column {} has duplicate POS_TO_ID member {expected}",
                column.column_id
            )));
        }
        Ok(())
    }

    /// Resolve a closure member to its source directory record.
    #[must_use]
    pub fn file(&self, member: Xldm140ClosureMember) -> Option<&'storage FileEntry> {
        self.storage.files.get(member.storage_index)
    }

    /// Atomically rewrite same-sized section 2.3/2.4 members while retaining
    /// the outer XLDM directory and every admitted identity/dependency edge.
    ///
    /// The candidate is rebuilt in a bounded buffer, reparsed through the
    /// outer codec, and inspected through sections 2.3--2.6 before this
    /// method returns. Each changed native/generated payload must retain its
    /// complete parsed semantic value; its table-local identities and section
    /// owners must equal this source-bound proof graph. Unknown members, outer
    /// marker allocations, path changes, size changes, semantic value changes,
    /// and dependency changes are refused. An empty replacement list is an
    /// exact borrowed no-op even if the source contains members outside this
    /// writer's typed closure; a byte-identical typed list is also borrowed.
    pub fn rewrite_same_size(
        &self,
        replacements: &[Xldm140FileReplacement<'_>],
    ) -> Result<Xldm140Patch<'storage>, OlapError> {
        if replacements.is_empty() {
            return Ok(Xldm140Patch {
                before: self.storage.source_bytes(),
                after: Xldm140PatchBytes::Borrowed(self.storage.source_bytes()),
            });
        }
        if !self.is_complete() {
            return Err(OlapError::new(
                "cannot rewrite an XLDM closure containing unknown members",
            ));
        }
        let mut seen = HashSet::new();
        seen.try_reserve(replacements.len()).map_err(|error| {
            OlapError::new(format!("cannot reserve XLDM replacements: {error}"))
        })?;
        let mut has_byte_change = false;
        for replacement in replacements {
            if !seen.insert(replacement.storage_path) {
                return Err(OlapError::new(format!(
                    "duplicate XLDM replacement path {}",
                    replacement.storage_path
                )));
            }
            let index = self
                .storage
                .files
                .iter()
                .position(|entry| entry.path == replacement.storage_path)
                .ok_or_else(|| {
                    OlapError::new(format!(
                        "XLDM replacement path {} is absent from the source directory",
                        replacement.storage_path
                    ))
                })?;
            let section = self
                .members
                .get(index)
                .map(|member| member.section)
                .ok_or_else(|| OlapError::new("XLDM closure member index is out of range"))?;
            if !matches!(
                section,
                Xldm140MemberSection::Native
                    | Xldm140MemberSection::Generated
                    | Xldm140MemberSection::NativeAndGenerated
            ) {
                return Err(OlapError::new(format!(
                    "XLDM replacement {} is outside the writable native/generated closure",
                    replacement.storage_path
                )));
            }
            let source = self.storage.file_payload(index).ok_or_else(|| {
                OlapError::new(format!(
                    "cannot resolve source payload for XLDM replacement {}",
                    replacement.storage_path
                ))
            })?;
            if source.len() != replacement.payload.len() {
                return Err(OlapError::new(format!(
                    "XLDM replacement {} changes allocation size",
                    replacement.storage_path
                )));
            }
            has_byte_change |= source != replacement.payload;
            self.validate_replacement_payload(index, replacement.payload)?;
        }
        if !has_byte_change {
            return Ok(Xldm140Patch {
                before: self.storage.source_bytes(),
                after: Xldm140PatchBytes::Borrowed(self.storage.source_bytes()),
            });
        }

        let candidate_bytes = super::codec::rewrite_same_size_payloads(self.storage, replacements)
            .map_err(|error| OlapError::new(format!("XLDM outer rewrite failed: {error}")))?;
        let candidate = super::codec::inspect(&candidate_bytes).map_err(|error| {
            OlapError::new(format!("XLDM rewritten source is invalid: {error}"))
        })?;
        if candidate.profile() != StorageProfile::Xldm140
            || candidate
                .files
                .iter()
                .zip(self.storage.files.iter())
                .any(|(after, before)| {
                    after.path != before.path
                        || after.stored_size != before.stored_size
                        || after.offset != before.offset
                        || after.last_write_timestamp != before.last_write_timestamp
                })
        {
            return Err(OlapError::new(
                "XLDM rewrite changed the source directory closure",
            ));
        }

        let metadata = super::metadata::inspect(&candidate).map_err(|error| {
            OlapError::new(format!("rewritten section 2.5 is invalid: {error}"))
        })?;
        let native = super::native::inspect(&candidate, &metadata.native_parse_options()).map_err(
            |error| OlapError::new(format!("rewritten section 2.3 is invalid: {error}")),
        )?;
        let generated =
            super::generated::inspect_system_generated(&candidate).map_err(|error| {
                OlapError::new(format!("rewritten section 2.4 is invalid: {error}"))
            })?;
        let olap = super::olap::inspect(&candidate, &metadata).map_err(|error| {
            OlapError::new(format!("rewritten section 2.6 is invalid: {error}"))
        })?;
        let candidate_closure =
            prove_xldm140_closure(&candidate, &metadata, &olap, &native, &generated)?;
        if candidate_closure.projection != self.projection
            || candidate_closure.members != self.members
            || !candidate_closure.unknown_members.is_empty()
        {
            return Err(OlapError::new(
                "XLDM rewrite changed the admitted identity or dependency closure",
            ));
        }
        Ok(Xldm140Patch {
            before: self.storage.source_bytes(),
            after: Xldm140PatchBytes::Owned(candidate_bytes),
        })
    }

    /// Rename one validated `XMSimpleTable@name` while retaining its
    /// generated TableID, column data, relationships, OLAP object, and outer
    /// allocation.  This is the first writable typed descriptor operation:
    /// the lexical attribute must fit the existing allocation, and a
    /// relationship endpoint or relationship metadata root that depends on
    /// that XML table name is refused until the closure-aware relationship
    /// writer can update every dependency atomically.
    pub fn rename_table_name(
        &self,
        table_id: &str,
        new_name: &str,
    ) -> Result<Xldm140Patch<'storage>, OlapError> {
        if new_name.is_empty() || !new_name.chars().all(valid_xml10_char) {
            return Err(OlapError::new(
                "Xldm140 table XML name is empty or contains an XML 1.0-forbidden character",
            ));
        }
        if !self.is_complete() {
            return Err(OlapError::new(
                "cannot rename a table in an XLDM closure containing unknown members",
            ));
        }
        let table = self
            .projection
            .table(table_id)
            .ok_or_else(|| OlapError::new(format!("unknown XLDM table identity {table_id}")))?;
        if table.xml_name == new_name {
            return Ok(Xldm140Patch {
                before: self.storage.source_bytes(),
                after: Xldm140PatchBytes::Borrowed(self.storage.source_bytes()),
            });
        }
        if self.projection.relationships.iter().any(|relationship| {
            relationship.primary_table == table.xml_name || relationship.primary_table == new_name
        }) {
            return Err(OlapError::new(format!(
                "table {table_id} XML rename has an unrewritten relationship endpoint; metadata and relationship-index closure must be edited together"
            )));
        }
        let metadata = self
            .metadata
            .ok_or_else(|| OlapError::new("table rename lacks the source metadata model"))?;
        for file in &metadata.files {
            if relationship_metadata_owned_by(file, table_id)? {
                return Err(OlapError::new(
                    "table XML rename has relationship metadata roots owned by this TableID; use the closure-aware relationship rename so every root is rewritten atomically",
                ));
            }
        }
        let file = metadata
            .files
            .iter()
            .find(|file| file.storage_path == table.metadata_path)
            .ok_or_else(|| {
                OlapError::new(format!(
                    "table {table_id} metadata {} is absent from the source model",
                    table.metadata_path
                ))
            })?;
        let replacement = replace_table_name_attribute(file.bytes, &table.xml_name, new_name)?;
        let candidate_bytes = super::codec::rewrite_same_size_payloads(
            self.storage,
            &[Xldm140FileReplacement {
                storage_path: file.storage_path,
                payload: &replacement,
            }],
        )
        .map_err(|error| {
            OlapError::new(format!("XLDM table rename outer rewrite failed: {error}"))
        })?;
        let candidate = super::codec::inspect(&candidate_bytes).map_err(|error| {
            OlapError::new(format!(
                "XLDM table rename produced invalid storage: {error}"
            ))
        })?;
        let metadata_after = super::metadata::inspect(&candidate).map_err(|error| {
            OlapError::new(format!("renamed table metadata is invalid: {error}"))
        })?;
        let native_after =
            super::native::inspect(&candidate, &metadata_after.native_parse_options()).map_err(
                |error| OlapError::new(format!("renamed native closure is invalid: {error}")),
            )?;
        let generated_after =
            super::generated::inspect_system_generated(&candidate).map_err(|error| {
                OlapError::new(format!("renamed generated closure is invalid: {error}"))
            })?;
        let olap_after = super::olap::inspect(&candidate, &metadata_after)
            .map_err(|error| OlapError::new(format!("renamed OLAP closure is invalid: {error}")))?;
        let candidate_closure = prove_xldm140_closure(
            &candidate,
            &metadata_after,
            &olap_after,
            &native_after,
            &generated_after,
        )?;
        if candidate_closure.members != self.members
            || !candidate_closure.unknown_members.is_empty()
            || !same_directory_closure(self.storage, &candidate)
            || !same_projection_except_table_name(
                &self.projection,
                &candidate_closure.projection,
                table_id,
                new_name,
            )
        {
            return Err(OlapError::new(
                "XLDM table rename changed an admitted identity or dependency closure",
            ));
        }
        Ok(Xldm140Patch {
            before: self.storage.source_bytes(),
            after: Xldm140PatchBytes::Owned(candidate_bytes),
        })
    }

    /// Rename a table XML name and update every proven metadata relationship
    /// endpoint that explicitly refers to that name. The operation keeps the
    /// generated TableID, Dimension/Attribute IDs, native values, generated
    /// relationship-index paths, and OLAP object IDs unchanged. It is admitted
    /// only when the complete section-2.2/2.5/2.6 proof is present; changing a
    /// TableID or relationship identity still requires a different operation
    /// because those changes alter generated member paths and native/index
    /// payload identities.
    ///
    /// Unlike [`Self::rename_table_name`], this operation permits the affected
    /// metadata allocations and the outer directory's serial layout to grow.
    /// The candidate is reparsed, reproven, and returned with an exact inverse
    /// source guard. Unknown members, ambiguous XML-name references, and
    /// unproven OLAP dependencies are refused before the outer buffer grows.
    pub fn rename_table_name_with_relationships(
        &self,
        table_id: &str,
        new_name: &str,
    ) -> Result<Xldm140Patch<'storage>, OlapError> {
        self.rename_table_name_with_relationships_with_limit(
            table_id,
            new_name,
            super::model::MAX_STORAGE_BYTES,
        )
        .map_err(Xldm140RenameError::into_olap_error)
    }

    /// Rename a table XML name and its proven relationship endpoints under an
    /// exact caller output cap. The outer allocation and directory plan is
    /// computed from encoded lengths before changed metadata payloads or the
    /// rewritten storage buffer are materialized. Exact semantic no-ops keep
    /// borrowing the original source even when the caller cap is zero.
    pub fn rename_table_name_with_relationships_with_limit(
        &self,
        table_id: &str,
        new_name: &str,
        max_output_bytes: usize,
    ) -> Result<Xldm140Patch<'storage>, Xldm140RenameError> {
        if new_name.is_empty() || !new_name.chars().all(valid_xml10_char) {
            return Err(Xldm140RenameError::invalid(
                "XLDM table XML name is empty or contains an XML 1.0-forbidden character",
            ));
        }
        if new_name.len() > super::model::MAX_PATH_BYTES {
            return Err(Xldm140RenameError::limit_with_bounds(
                "XLDM table XML name exceeds the bounded path-text limit",
                Some((new_name.len(), super::model::MAX_PATH_BYTES)),
            ));
        }
        let table = self.projection.table(table_id).ok_or_else(|| {
            Xldm140RenameError::invalid(format!("unknown XLDM table identity {table_id}"))
        })?;
        if table.xml_name == new_name {
            return Ok(Xldm140Patch {
                before: self.storage.source_bytes(),
                after: Xldm140PatchBytes::Borrowed(self.storage.source_bytes()),
            });
        }
        if !self.is_complete() {
            return Err(Xldm140RenameError::unsupported(
                "cannot structurally rename a table in an XLDM closure containing unknown members",
            ));
        }
        if self
            .projection
            .tables
            .iter()
            .filter(|candidate| candidate.xml_name == table.xml_name)
            .count()
            != 1
        {
            return Err(Xldm140RenameError::invalid(
                "table XML rename is ambiguous because the source XML name is shared",
            ));
        }
        if self
            .projection
            .tables
            .iter()
            .any(|candidate| candidate.xml_name == new_name)
        {
            return Err(Xldm140RenameError::invalid(
                "table XML rename would collide with another table XML name",
            ));
        }
        let metadata = self.metadata.ok_or_else(|| {
            Xldm140RenameError::invalid("table rename lacks the source metadata model")
        })?;
        let source_olap = super::olap::inspect(self.storage, metadata).map_err(|error| {
            Xldm140RenameError::invalid(format!("source OLAP model is invalid: {error}"))
        })?;
        let source_olap_proof = super::olapproof::prove_xldm140_olap(
            self.storage,
            metadata,
            &source_olap,
            super::olapproof::OlapProofLimits::default(),
        )
        .map_err(Xldm140RenameError::from_olap_proof)?;
        if !source_olap_proof.is_complete() {
            return Err(Xldm140RenameError::unsupported(
                "table XML rename requires a complete OLAP/file-group proof",
            ));
        }

        // Parse every affected source span and compute each exact rewritten
        // payload length before retaining a source copy or an escaped name.
        // The section parser bounds each member, but a model may contain many
        // such members.  Charge the aggregate result before any payload
        // buffer is grown so allocation failure/refusal is deterministic.
        let mut rewrite_count = 0usize;
        let mut rewrite_output_bytes = 0usize;
        let mut payload_lengths = Vec::new();
        payload_lengths
            .try_reserve_exact(metadata.files.len())
            .map_err(|error| Xldm140RenameError::allocation("payload lengths", error))?;
        for file in &metadata.files {
            let relationship_root_owned = relationship_metadata_owned_by(file, table_id)?;
            let rewrites_root = file.storage_path == table.metadata_path || relationship_root_owned;
            let rewrites_relationships = file.table.collection("Relationships").is_some();
            if !rewrites_root && !rewrites_relationships {
                continue;
            }
            let mut output_len = file.bytes.len();
            let mut changed = false;
            if rewrites_root {
                output_len =
                    table_name_replacement_output_len(file.bytes, &table.xml_name, new_name)?;
                changed = true;
            }
            if rewrites_relationships {
                let (relationship_len, matches) =
                    relationship_replacement_output_len(file.bytes, &table.xml_name, new_name)?;
                if matches != 0 {
                    output_len = output_len
                        .checked_add(relationship_len)
                        .and_then(|value| value.checked_sub(file.bytes.len()))
                        .ok_or_else(|| {
                            Xldm140RenameError::limit("table rename output-byte budget overflow")
                        })?;
                    changed = true;
                }
            }
            if changed {
                rewrite_count = rewrite_count.checked_add(1).ok_or_else(|| {
                    Xldm140RenameError::limit("table rename member count overflow")
                })?;
                rewrite_output_bytes =
                    rewrite_output_bytes
                        .checked_add(output_len)
                        .ok_or_else(|| {
                            Xldm140RenameError::limit("table rename output-byte budget overflow")
                        })?;
                payload_lengths.push(super::codec::VariablePayloadLength {
                    storage_path: file.storage_path,
                    payload_len: output_len,
                });
            }
        }
        if rewrite_output_bytes > super::model::MAX_STORAGE_BYTES {
            return Err(Xldm140RenameError::limit_with_bounds(
                "table rename output-byte budget exceeds the XLDM storage limit",
                Some((rewrite_output_bytes, super::model::MAX_STORAGE_BYTES)),
            ));
        }
        super::codec::preflight_variable_size_payloads(
            self.storage,
            &payload_lengths,
            max_output_bytes,
        )
        .map_err(Xldm140RenameError::from_variable_rewrite)?;

        let mut owned_payloads: Vec<(&str, Vec<u8>)> = Vec::new();
        owned_payloads
            .try_reserve_exact(rewrite_count)
            .map_err(|error| Xldm140RenameError::allocation("payloads", error))?;
        for file in &metadata.files {
            let is_target_table = file.storage_path == table.metadata_path;
            let is_relationship_file = relationship_metadata_owned_by(file, table_id)?;
            let has_relationships = file.table.collection("Relationships").is_some();
            if !is_target_table && !is_relationship_file && !has_relationships {
                continue;
            }
            let mut payload = None;
            let mut changed = false;
            if is_target_table || is_relationship_file {
                payload = Some(replace_table_name_attribute_variable(
                    file.bytes,
                    &table.xml_name,
                    new_name,
                )?);
                changed = true;
            }
            if has_relationships {
                let (_, matches) =
                    relationship_replacement_output_len(file.bytes, &table.xml_name, new_name)?;
                if matches != 0 {
                    let source = payload.as_deref().unwrap_or(file.bytes);
                    payload = Some(replace_relationship_primary_table(
                        source,
                        &table.xml_name,
                        new_name,
                    )?);
                    changed = true;
                }
            }
            if changed {
                let payload = payload.ok_or_else(|| {
                    Xldm140RenameError::invalid("table rename planner lost a rewritten payload")
                })?;
                super::metadata::parse_file(file.storage_path, &payload).map_err(|error| {
                    Xldm140RenameError::invalid(format!(
                        "rewritten relationship metadata {} is invalid: {error}",
                        file.storage_path
                    ))
                })?;
                owned_payloads.push((file.storage_path, payload));
            }
        }
        if owned_payloads.is_empty() {
            return Err(Xldm140RenameError::invalid(
                "table XML rename found no source metadata member to rewrite",
            ));
        }
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(owned_payloads.len())
            .map_err(|error| Xldm140RenameError::allocation("replacements", error))?;
        for (storage_path, payload) in &owned_payloads {
            replacements.push(Xldm140FileReplacement {
                storage_path,
                payload,
            });
        }

        let candidate_bytes = super::codec::rewrite_variable_size_payloads_with_limit(
            self.storage,
            &replacements,
            max_output_bytes,
        )
        .map_err(Xldm140RenameError::from_codec)?;
        let candidate = super::codec::inspect(&candidate_bytes).map_err(|error| {
            Xldm140RenameError::invalid(format!(
                "XLDM structural table rename produced invalid storage: {error}"
            ))
        })?;
        let metadata_after = super::metadata::inspect(&candidate).map_err(|error| {
            Xldm140RenameError::invalid(format!("renamed table metadata is invalid: {error}"))
        })?;
        let native_after = super::native::inspect(
            &candidate,
            &metadata_after.native_parse_options(),
        )
        .map_err(|error| {
            Xldm140RenameError::invalid(format!("renamed native closure is invalid: {error}"))
        })?;
        let generated_after =
            super::generated::inspect_system_generated(&candidate).map_err(|error| {
                Xldm140RenameError::invalid(format!(
                    "renamed generated closure is invalid: {error}"
                ))
            })?;
        let olap_after = super::olap::inspect(&candidate, &metadata_after).map_err(|error| {
            Xldm140RenameError::invalid(format!("renamed OLAP closure is invalid: {error}"))
        })?;
        let candidate_closure = prove_xldm140_closure(
            &candidate,
            &metadata_after,
            &olap_after,
            &native_after,
            &generated_after,
        )?;
        if candidate_closure.members != self.members
            || !candidate_closure.unknown_members.is_empty()
            || !same_projection_except_table_name_and_relationships(
                &self.projection,
                &candidate_closure.projection,
                table_id,
                &table.xml_name,
                new_name,
            )
        {
            return Err(Xldm140RenameError::invalid(
                "XLDM structural table rename changed an admitted identity or dependency closure",
            ));
        }
        let candidate_olap_proof = super::olapproof::prove_xldm140_olap(
            &candidate,
            &metadata_after,
            &olap_after,
            super::olapproof::OlapProofLimits::default(),
        )
        .map_err(Xldm140RenameError::from_olap_proof)?;
        if !candidate_olap_proof.is_complete()
            || !same_olap_proof_except_table_name(
                &source_olap_proof,
                &candidate_olap_proof,
                table_id,
                &table.xml_name,
                new_name,
            )
        {
            return Err(Xldm140RenameError::invalid(
                "XLDM structural table rename changed the proven OLAP dependency closure",
            ));
        }
        Ok(Xldm140Patch {
            before: self.storage.source_bytes(),
            after: Xldm140PatchBytes::Owned(candidate_bytes),
        })
    }

    fn validate_replacement_payload(&self, index: usize, payload: &[u8]) -> Result<(), OlapError> {
        let Some(section) = self.members.get(index).map(|member| member.section) else {
            return Err(OlapError::new("XLDM closure member index is out of range"));
        };
        let path = self
            .storage
            .files
            .get(index)
            .map(|entry| entry.path.as_str())
            .ok_or_else(|| OlapError::new("XLDM replacement path index is out of range"))?;
        let generated = classify_generated_path(path)
            .map_err(|error| OlapError::new(format!("invalid replacement path {path}: {error}")))?;
        if matches!(
            section,
            Xldm140MemberSection::Generated | Xldm140MemberSection::NativeAndGenerated
        ) {
            let Some(model) = self.generated else {
                return Err(OlapError::new(
                    "generated replacement lacks the source semantic model",
                ));
            };
            let source = model
                .files
                .iter()
                .find(|file| file.storage_path == path)
                .ok_or_else(|| {
                    OlapError::new(format!(
                        "generated replacement {path} is absent from its source semantic model"
                    ))
                })?;
            let candidate = parse_system_generated_file(path, payload)
                .map_err(|error| OlapError::new(format!("generated replacement {path}: {error}")))?
                .ok_or_else(|| {
                    OlapError::new(format!("replacement {path} is not generated data"))
                })?;
            if candidate.kind != source.kind
                || candidate.object_key != source.object_key
                || candidate.version != source.version
                || candidate.data != source.data
            {
                return Err(OlapError::new(format!(
                    "generated replacement {path} changes its typed role, object identity, or semantic payload"
                )));
            }
        }
        if matches!(
            section,
            Xldm140MemberSection::Native | Xldm140MemberSection::NativeAndGenerated
        ) {
            let Some(model) = self.native else {
                return Err(OlapError::new(
                    "native replacement lacks the source semantic model",
                ));
            };
            let source = model
                .files
                .iter()
                .find(|file| file.storage_path == path)
                .ok_or_else(|| {
                    OlapError::new(format!(
                        "native replacement {path} is absent from its source semantic model"
                    ))
                })?;
            let candidate_data = match generated.kind {
                GeneratedNameKind::ColumnData
                | GeneratedNameKind::ColumnPositionToId
                | GeneratedNameKind::ColumnIdToPosition
                | GeneratedNameKind::UserHierarchyChildCount
                | GeneratedNameKind::UserHierarchyFirstChildPosition
                | GeneratedNameKind::UserHierarchyParentPosition
                | GeneratedNameKind::UserHierarchyMultilevelId => {
                    let candidate = super::native::parse_idf(payload).map_err(|error| {
                        OlapError::new(format!("native replacement {path}: {error}"))
                    })?;
                    if generated.kind == GeneratedNameKind::ColumnData
                        && (candidate.segments.len() < 2 || candidate.segments.len() % 2 != 0)
                    {
                        return Err(OlapError::new(format!(
                            "native replacement {path} lacks column primary/subsegment pairs"
                        )));
                    }
                    super::native::NativeData::Idf(candidate)
                },
                GeneratedNameKind::TableRelationshipIndex => {
                    super::native::parse_relationship_index(payload).map_err(|error| {
                        OlapError::new(format!(
                            "native replacement {path} has an invalid relationship index: {error}"
                        ))
                    })?
                },
                GeneratedNameKind::ColumnDictionary => {
                    let mode = self
                        .metadata
                        .map(|metadata| metadata.native_parse_options())
                        .map(|options| options.string_hash_mode(path))
                        .transpose()
                        .map_err(|error| {
                            OlapError::new(format!("native replacement {path}: {error}"))
                        })?
                        .unwrap_or_default();
                    super::native::NativeData::Dictionary(
                        super::native::parse_dictionary(payload, mode).map_err(|error| {
                            OlapError::new(format!("native replacement {path}: {error}"))
                        })?,
                    )
                },
                GeneratedNameKind::ColumnHashIndex => super::native::NativeData::HashIndex(
                    super::native::parse_hash_index(payload).map_err(|error| {
                        OlapError::new(format!("native replacement {path}: {error}"))
                    })?,
                ),
                _ => {
                    return Err(OlapError::new(format!(
                        "native replacement {path} has no supported typed role"
                    )));
                },
            };
            if candidate_data != source.data {
                return Err(OlapError::new(format!(
                    "native replacement {path} changes its semantic payload"
                )));
            }
        }
        Ok(())
    }
}

fn bind_column(
    table: &Xldm140TableIdentity,
    column_name: &str,
    column_id: &str,
    expected_calculated: bool,
    role: &str,
    content_type: Option<Xldm140TimeGroupingContentType>,
) -> Result<Xldm140ColumnBinding, OlapError> {
    if column_name.is_empty() || column_id.is_empty() {
        return Err(OlapError::new(format!(
            "time grouping {role} column name and immutable ID must be non-empty"
        )));
    }
    let column = table
        .columns
        .iter()
        .find(|column| column.column_id == column_id)
        .ok_or_else(|| {
            OlapError::new(format!(
                "time grouping {role} column {column_id} is absent from table {}",
                table.xml_name
            ))
        })?;
    let expected_name = column
        .column_name
        .as_deref()
        .unwrap_or(column.column_id.as_str());
    if column_name != expected_name {
        return Err(OlapError::new(format!(
            "time grouping {role} column name {column_name} is not the validated XLDM Attribute name {expected_name} for immutable ID {column_id}"
        )));
    }
    if column.is_calculated != expected_calculated {
        let actual = if column.is_calculated {
            "calculated"
        } else {
            "source"
        };
        let expected = if expected_calculated {
            "calculated"
        } else {
            "source"
        };
        return Err(OlapError::new(format!(
            "time grouping {role} column {column_id} is classified as {actual}, expected {expected}"
        )));
    }
    Ok(Xldm140ColumnBinding {
        table_id: table.table_id.clone(),
        table_name: table.xml_name.clone(),
        column_name: column_name.to_owned(),
        column_id: column_id.to_owned(),
        data_path: column.data_path.clone(),
        is_calculated: column.is_calculated,
        db_type: column.db_type,
        settings: column.settings,
        content_type,
    })
}

fn same_directory_closure(before: &Storage<'_>, after: &Storage<'_>) -> bool {
    before.profile() == after.profile()
        && before.files.len() == after.files.len()
        && before
            .files
            .iter()
            .zip(&after.files)
            .all(|(before, after)| {
                before.path == after.path
                    && before.kind == after.kind
                    && before.offset == after.offset
                    && before.stored_size == after.stored_size
                    && before.delete == after.delete
                    && before.created_timestamp == after.created_timestamp
                    && before.access_timestamp == after.access_timestamp
                    && before.last_write_timestamp == after.last_write_timestamp
            })
}

fn same_projection_except_table_name(
    before: &Xldm140IdentityProjection,
    after: &Xldm140IdentityProjection,
    changed_table_id: &str,
    changed_name: &str,
) -> bool {
    if before.relationships != after.relationships || before.tables.len() != after.tables.len() {
        return false;
    }
    before
        .tables
        .iter()
        .zip(&after.tables)
        .all(|(before, after)| {
            let expected_name = if before.table_id == changed_table_id {
                changed_name
            } else {
                before.xml_name.as_str()
            };
            before.table_id == after.table_id
                && after.xml_name == expected_name
                && before.metadata_path == after.metadata_path
                && before.dimension_object_id == after.dimension_object_id
                && before.attribute_ids == after.attribute_ids
                && before.columns == after.columns
        })
}

fn same_projection_except_table_name_and_relationships(
    before: &Xldm140IdentityProjection,
    after: &Xldm140IdentityProjection,
    changed_table_id: &str,
    old_name: &str,
    changed_name: &str,
) -> bool {
    if before.tables.len() != after.tables.len()
        || before.relationships.len() != after.relationships.len()
    {
        return false;
    }
    let tables_equal = before
        .tables
        .iter()
        .zip(&after.tables)
        .all(|(before, after)| {
            let expected_name = if before.table_id == changed_table_id {
                changed_name
            } else {
                before.xml_name.as_str()
            };
            before.table_id == after.table_id
                && after.xml_name == expected_name
                && before.metadata_path == after.metadata_path
                && before.dimension_object_id == after.dimension_object_id
                && before.attribute_ids == after.attribute_ids
                && before.columns == after.columns
        });
    if !tables_equal {
        return false;
    }
    before
        .relationships
        .iter()
        .zip(&after.relationships)
        .all(|(before, after)| {
            let expected_primary = if before.primary_table == old_name {
                changed_name
            } else {
                before.primary_table.as_str()
            };
            before.containing_table == after.containing_table
                && before.relationship_name == after.relationship_name
                && before.relationship_id == after.relationship_id
                && before.metadata_path == after.metadata_path
                && after.primary_table == expected_primary
                && before.primary_column == after.primary_column
                && before.foreign_column == after.foreign_column
                && before.expected_index_key == after.expected_index_key
        })
}

fn same_olap_proof_except_table_name(
    before: &super::olapproof::Xldm140OlapProof<'_, '_>,
    after: &super::olapproof::Xldm140OlapProof<'_, '_>,
    changed_table_id: &str,
    old_name: &str,
    changed_name: &str,
) -> bool {
    if before.file_groups() != after.file_groups()
        || before.cube() != after.cube()
        || before.measure_groups() != after.measure_groups()
        || before.partitions() != after.partitions()
        || before.unknown_members() != after.unknown_members()
        || before.tables().len() != after.tables().len()
        || before.relationships().len() != after.relationships().len()
    {
        return false;
    }
    let tables_equal = before
        .tables()
        .iter()
        .zip(after.tables())
        .all(|(before, after)| {
            let expected_name = if before.table_id == changed_table_id {
                changed_name
            } else {
                before.metadata_name.as_deref().unwrap_or("")
            };
            before.table_id == after.table_id
                && after.metadata_name.as_deref().unwrap_or("") == expected_name
                && before.metadata_path == after.metadata_path
                && before.dimension_path == after.dimension_path
                && before.dimension_id == after.dimension_id
                && before.dimension_object_id == after.dimension_object_id
                && before.dimension_name == after.dimension_name
                && before.attribute_ids == after.attribute_ids
                && before.data_files == after.data_files
        });
    if !tables_equal {
        return false;
    }
    before
        .relationships()
        .iter()
        .zip(after.relationships())
        .all(|(before, after)| {
            let expected_primary = if before.primary_table == old_name {
                changed_name
            } else {
                before.primary_table.as_str()
            };
            before.metadata_path == after.metadata_path
                && before.relationship_name == after.relationship_name
                && before.generated_relationship_id == after.generated_relationship_id
                && before.containing_table == after.containing_table
                && after.primary_table == expected_primary
                && before.primary_column == after.primary_column
                && before.foreign_column == after.foreign_column
                && before.dimension_path == after.dimension_path
                && before.dimension_reference == after.dimension_reference
                && before.relationship_index_paths == after.relationship_index_paths
        })
}

fn replace_table_name_attribute(
    source: &[u8],
    expected_name: &str,
    new_name: &str,
) -> Result<Vec<u8>, OlapError> {
    let output_len = table_name_replacement_output_len(source, expected_name, new_name)?;
    if output_len != source.len() {
        return Err(OlapError::new(
            "XLDM table XML rename changes its fixed allocation size",
        ));
    }
    let result = replace_table_name_attribute_variable(source, expected_name, new_name)?;
    debug_assert_eq!(result.len(), output_len);
    Ok(result)
}

fn replace_table_name_attribute_variable(
    source: &[u8],
    expected_name: &str,
    new_name: &str,
) -> Result<Vec<u8>, OlapError> {
    let output_len = table_name_replacement_output_len(source, expected_name, new_name)?;
    let root_start = find_xmobject_root_start(source)?;
    let root_end = find_start_tag_end(source, root_start)?;
    let (value_start, value_end, quote) = find_name_attribute(source, root_start, root_end)?;
    let encoded_old = std::str::from_utf8(&source[value_start..value_end])
        .map_err(|error| OlapError::new(format!("table XML name is not UTF-8: {error}")))?;
    let decoded_old = quick_xml::escape::unescape(encoded_old)
        .map_err(|error| OlapError::new(format!("table XML name is not well-formed: {error}")))?;
    if decoded_old != expected_name {
        return Err(OlapError::new(
            "table XML name changed since the identity snapshot was admitted",
        ));
    }
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|error| OlapError::new(format!("cannot reserve table rename bytes: {error}")))?;
    result.extend_from_slice(&source[..value_start]);
    append_escaped_xml_attribute(&mut result, new_name, quote)?;
    result.extend_from_slice(&source[value_end..]);
    Ok(result)
}

fn table_name_replacement_output_len(
    source: &[u8],
    expected_name: &str,
    new_name: &str,
) -> Result<usize, OlapError> {
    if !new_name.chars().all(valid_xml10_char) {
        return Err(OlapError::new(
            "XLDM table XML name contains an XML 1.0-forbidden character",
        ));
    }
    if new_name.len() > super::model::MAX_PATH_BYTES {
        return Err(OlapError::new(
            "XLDM table XML name exceeds the bounded path-text limit",
        ));
    }
    let root_start = find_xmobject_root_start(source)?;
    let root_end = find_start_tag_end(source, root_start)?;
    let (value_start, value_end, quote) = find_name_attribute(source, root_start, root_end)?;
    let encoded_old = std::str::from_utf8(&source[value_start..value_end])
        .map_err(|error| OlapError::new(format!("table XML name is not UTF-8: {error}")))?;
    let decoded_old = quick_xml::escape::unescape(encoded_old)
        .map_err(|error| OlapError::new(format!("table XML name is not well-formed: {error}")))?;
    if decoded_old != expected_name {
        return Err(OlapError::new(
            "table XML name changed since the identity snapshot was admitted",
        ));
    }
    let new_len = escaped_xml_attribute_len(new_name, quote)?;
    let output_len = source
        .len()
        .checked_sub(value_end - value_start)
        .and_then(|value| value.checked_add(new_len))
        .ok_or_else(|| OlapError::new("table XML rename allocation size overflow"))?;
    if output_len > super::model::MAX_STORAGE_BYTES {
        return Err(OlapError::new(
            "table XML rename output exceeds the XLDM storage limit",
        ));
    }
    Ok(output_len)
}

fn replace_relationship_primary_table(
    source: &[u8],
    expected_name: &str,
    new_name: &str,
) -> Result<Vec<u8>, OlapError> {
    // First collect lexical spans without allocating a rewritten copy. This
    // validates every PrimaryTable element, including nonmatching endpoints,
    // and computes the exact span/output bounds before any rewritten payload
    // buffer is reserved.
    let (planned_output_len, match_count) =
        relationship_replacement_output_len(source, expected_name, new_name)?;
    if match_count == 0 {
        return Ok(source.to_vec());
    }
    let mut spans = Vec::new();
    spans
        .try_reserve_exact(match_count)
        .map_err(|error| OlapError::new(format!("cannot reserve relationship spans: {error}")))?;
    scan_relationship_primary_table(source, expected_name, |value_start, value_end| {
        spans.push((value_start, value_end));
    })?;
    if spans.len() != match_count {
        return Err(OlapError::new(
            "relationship rename span preflight changed during validation",
        ));
    }
    let replacement_len = escaped_xml_text_len(new_name)?;

    let output_len = spans
        .iter()
        .try_fold(source.len(), |length, (start, end)| {
            length
                .checked_sub(end - start)
                .and_then(|value| value.checked_add(replacement_len))
                .ok_or_else(|| OlapError::new("relationship rename allocation size overflow"))
        })?;
    if output_len != planned_output_len {
        return Err(OlapError::new(
            "relationship rename output preflight changed during validation",
        ));
    }
    if output_len > super::model::MAX_STORAGE_BYTES {
        return Err(OlapError::new(
            "relationship rename output exceeds the XLDM storage limit",
        ));
    }
    let mut output = Vec::new();
    output.try_reserve_exact(output_len).map_err(|error| {
        OlapError::new(format!("cannot reserve relationship rename bytes: {error}"))
    })?;
    let mut previous = 0usize;
    for (start, end) in spans {
        output.extend_from_slice(&source[previous..start]);
        append_escaped_xml_text(&mut output, new_name)?;
        previous = end;
    }
    output.extend_from_slice(&source[previous..]);
    Ok(output)
}

fn relationship_replacement_output_len(
    source: &[u8],
    expected_name: &str,
    new_name: &str,
) -> Result<(usize, usize), OlapError> {
    let replacement_len = escaped_xml_text_len(new_name)?;
    let mut output_len = source.len();
    let mut matches = 0usize;
    let mut overflow = false;
    scan_relationship_primary_table(source, expected_name, |value_start, value_end| {
        if overflow {
            return;
        }
        let Some(next_matches) = matches.checked_add(1) else {
            overflow = true;
            return;
        };
        let Some(next_length) = output_len
            .checked_sub(value_end - value_start)
            .and_then(|value| value.checked_add(replacement_len))
        else {
            overflow = true;
            return;
        };
        matches = next_matches;
        output_len = next_length;
    })?;
    if overflow {
        return Err(OlapError::new(
            "relationship rename allocation size overflow",
        ));
    }
    Ok((output_len, matches))
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum RelationshipXmlFrame {
    Other,
    Object,
    Relationships,
    RelationshipObject,
    RelationshipProperties,
    PrimaryTable,
}

struct RelationshipPrimaryTableValue {
    content_start: usize,
}

fn relationship_xml_frame(
    start: &quick_xml::events::BytesStart<'_>,
    parent: Option<RelationshipXmlFrame>,
) -> Result<RelationshipXmlFrame, OlapError> {
    let name = start.name();
    let name = name.as_ref();
    if name == b"PrimaryTable"
        && matches!(
            parent,
            Some(
                RelationshipXmlFrame::Relationships
                    | RelationshipXmlFrame::RelationshipObject
                    | RelationshipXmlFrame::RelationshipProperties
                    | RelationshipXmlFrame::Object
            )
        )
    {
        return Ok(RelationshipXmlFrame::PrimaryTable);
    }
    if name == b"XMRelationship" {
        return Ok(RelationshipXmlFrame::RelationshipObject);
    }
    if name == b"XMObject" {
        let mut is_relationship = false;
        for attribute in start.attributes() {
            let attribute = attribute.map_err(|error| {
                OlapError::new(format!("relationship XML attribute is invalid: {error}"))
            })?;
            if attribute.key.as_ref() == b"class" {
                // The recognized metadata class is an ASCII token.  An
                // escaped spelling is opaque rather than another admitted
                // relationship owner, so it must not broaden the rewrite
                // scope.
                is_relationship = attribute.value.as_ref() == b"XMRelationship";
            }
        }
        return Ok(if is_relationship {
            RelationshipXmlFrame::RelationshipObject
        } else {
            RelationshipXmlFrame::Object
        });
    }
    if name == b"Relationships" && parent == Some(RelationshipXmlFrame::Object) {
        return Ok(RelationshipXmlFrame::Relationships);
    }
    if name == b"Properties" && parent == Some(RelationshipXmlFrame::RelationshipObject) {
        return Ok(RelationshipXmlFrame::RelationshipProperties);
    }
    Ok(RelationshipXmlFrame::Other)
}

fn relationship_xml_text_range(source: &[u8], start: usize, end: usize) -> (usize, usize) {
    let mut value_start = start;
    while value_start < end && source[value_start].is_ascii_whitespace() {
        value_start += 1;
    }
    let mut value_end = end;
    while value_end > value_start && source[value_end - 1].is_ascii_whitespace() {
        value_end -= 1;
    }
    (value_start, value_end)
}

/// Bytes of `source` before a quick-xml reader's position zero: the reader
/// drops one leading UTF-8 byte-order mark before its first event without
/// counting it, so each of its positions is this many bytes short of the byte
/// offset in `source`.
///
/// This crate deliberately has no `litchi-core` dependency; the rule is
/// `litchi_core::xml::ReaderOrigin`'s, and its tests pin the same reader
/// behaviour.
fn reader_origin(source: &[u8]) -> usize {
    if source.starts_with(b"\xEF\xBB\xBF") {
        3
    } else {
        0
    }
}

fn scan_relationship_primary_table(
    source: &[u8],
    expected_name: &str,
    mut on_match: impl FnMut(usize, usize),
) -> Result<(), OlapError> {
    let mut reader = quick_xml::reader::Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().enable_all_checks(true);
    let mut frames = Vec::new();
    frames
        .try_reserve(source.len().min(super::model::MAX_XML_DEPTH))
        .map_err(|error| {
            OlapError::new(format!("cannot reserve relationship XML stack: {error}"))
        })?;
    let mut candidate = None;
    let origin = reader_origin(source);
    let mut event_start = origin;
    loop {
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(error) => {
                if candidate.is_some() {
                    return Err(OlapError::new("PrimaryTable element is unclosed"));
                }
                return Err(OlapError::new(format!(
                    "relationship XML is not well-formed: {error}"
                )));
            },
        };
        let event_end = usize::try_from(reader.buffer_position())
            .ok()
            .and_then(|position| position.checked_add(origin))
            .ok_or_else(|| OlapError::new("relationship XML position exceeds host size"))?;
        if event_end > source.len() || event_start > event_end {
            return Err(OlapError::new("relationship XML event range is invalid"));
        }
        match event {
            quick_xml::events::Event::Start(start) => {
                if candidate.is_some() {
                    return Err(OlapError::new("PrimaryTable element is not scalar"));
                }
                let frame = relationship_xml_frame(&start, frames.last().copied())?;
                if frame == RelationshipXmlFrame::PrimaryTable {
                    candidate = Some(RelationshipPrimaryTableValue {
                        content_start: event_end,
                    });
                }
                frames.push(frame);
            },
            quick_xml::events::Event::Empty(empty) => {
                if candidate.is_some() {
                    return Err(OlapError::new("PrimaryTable element is not scalar"));
                }
                let frame = relationship_xml_frame(&empty, frames.last().copied())?;
                if frame == RelationshipXmlFrame::PrimaryTable {
                    // An empty scalar has no endpoint value.
                }
            },
            quick_xml::events::Event::End(end) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| OlapError::new("relationship XML has an unmatched end"))?;
                if frame == RelationshipXmlFrame::PrimaryTable {
                    let value = candidate
                        .take()
                        .ok_or_else(|| OlapError::new("PrimaryTable scalar state is invalid"))?;
                    let (value_start, value_end) =
                        relationship_xml_text_range(source, value.content_start, event_start);
                    if value_start < value_end {
                        let lexical = std::str::from_utf8(&source[value_start..value_end])
                            .map_err(|error| {
                                OlapError::new(format!("PrimaryTable value is not UTF-8: {error}"))
                            })?;
                        let decoded = quick_xml::escape::unescape(lexical).map_err(|error| {
                            OlapError::new(format!("PrimaryTable value is not XML: {error}"))
                        })?;
                        if decoded == expected_name {
                            on_match(value_start, value_end);
                        }
                    }
                }
                let _ = end;
            },
            quick_xml::events::Event::Text(_) | quick_xml::events::Event::GeneralRef(_) => {},
            quick_xml::events::Event::CData(_)
            | quick_xml::events::Event::Comment(_)
            | quick_xml::events::Event::PI(_)
            | quick_xml::events::Event::DocType(_) => {
                if candidate.is_some() {
                    return Err(OlapError::new("PrimaryTable element is not scalar"));
                }
            },
            quick_xml::events::Event::Decl(_) => {},
            quick_xml::events::Event::Eof => {
                if candidate.is_some() || !frames.is_empty() {
                    return Err(OlapError::new("PrimaryTable element is unclosed"));
                }
                break;
            },
        }
        event_start = event_end;
    }
    Ok(())
}

fn escaped_xml_text_len(value: &str) -> Result<usize, OlapError> {
    if !value.chars().all(valid_xml10_char) {
        return Err(OlapError::new(
            "XML text contains an XML 1.0-forbidden character",
        ));
    }
    value.chars().try_fold(0usize, |length, character| {
        let addition = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length
            .checked_add(addition)
            .ok_or_else(|| OlapError::new("XML text length overflow"))
    })
}

fn append_escaped_xml_text(output: &mut Vec<u8>, value: &str) -> Result<(), OlapError> {
    let _ = escaped_xml_text_len(value)?;
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\'' => output.extend_from_slice(b"&apos;"),
            _ => {
                let mut encoded = [0_u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
    Ok(())
}

fn find_start_tag_end(source: &[u8], start: usize) -> Result<usize, OlapError> {
    let mut quote = None;
    for (offset, byte) in source.iter().enumerate().skip(start + 1) {
        match (quote, *byte) {
            (None, b'\'' | b'"') => quote = Some(*byte),
            (Some(value), byte) if byte == value => quote = None,
            (None, b'>') => return Ok(offset),
            _ => {},
        }
    }
    Err(OlapError::new(
        "table metadata XMObject start tag is unclosed",
    ))
}

/// Locate the actual first element rather than a byte lookalike in a prolog
/// comment, processing instruction, or CDATA payload.  The root grammar for
/// table metadata admits an unprefixed `XMObject` element; any other first
/// element is refused before its attributes are inspected.
fn find_xmobject_root_start(source: &[u8]) -> Result<usize, OlapError> {
    let mut reader = quick_xml::reader::Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().enable_all_checks(true);
    let origin = reader_origin(source);
    let mut event_start = origin;
    loop {
        let event = reader.read_event().map_err(|error| {
            OlapError::new(format!("table metadata XML is not well-formed: {error}"))
        })?;
        let event_end = usize::try_from(reader.buffer_position())
            .ok()
            .and_then(|position| position.checked_add(origin))
            .ok_or_else(|| OlapError::new("table metadata XML position exceeds host size"))?;
        if event_end > source.len() || event_start > event_end {
            return Err(OlapError::new("table metadata XML event range is invalid"));
        }
        match event {
            quick_xml::events::Event::Start(start) => {
                if start.name().as_ref() == b"XMObject" {
                    return Ok(event_start);
                }
                return Err(OlapError::new("table metadata has a non-XMObject root"));
            },
            quick_xml::events::Event::Empty(empty) => {
                if empty.name().as_ref() == b"XMObject" {
                    return Ok(event_start);
                }
                return Err(OlapError::new("table metadata has a non-XMObject root"));
            },
            quick_xml::events::Event::End(_) => {
                return Err(OlapError::new("table metadata has an unmatched end"));
            },
            quick_xml::events::Event::Eof => {
                return Err(OlapError::new("table metadata has no XMObject root"));
            },
            quick_xml::events::Event::Text(_)
            | quick_xml::events::Event::GeneralRef(_)
            | quick_xml::events::Event::CData(_)
            | quick_xml::events::Event::Comment(_)
            | quick_xml::events::Event::PI(_)
            | quick_xml::events::Event::DocType(_)
            | quick_xml::events::Event::Decl(_) => {},
        }
        event_start = event_end;
    }
}

fn find_name_attribute(
    source: &[u8],
    start: usize,
    end: usize,
) -> Result<(usize, usize, u8), OlapError> {
    let mut cursor = start + b"<XMObject".len();
    while cursor < end {
        while cursor < end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end || source[cursor] == b'/' {
            break;
        }
        let key_start = cursor;
        while cursor < end
            && !source[cursor].is_ascii_whitespace()
            && !matches!(source[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let key = &source[key_start..cursor];
        while cursor < end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end || source[cursor] != b'=' {
            return Err(OlapError::new(
                "table metadata has a malformed XMObject attribute",
            ));
        }
        cursor += 1;
        while cursor < end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *source
            .get(cursor)
            .filter(|byte| matches!(byte, b'\'' | b'"'))
            .ok_or_else(|| OlapError::new("table metadata attribute value is unquoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < end && source[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= end {
            return Err(OlapError::new("table metadata attribute value is unclosed"));
        }
        cursor += 1;
        if key == b"name" {
            return Ok((value_start, value_end, quote));
        }
    }
    Err(OlapError::new(
        "table metadata XMObject has no name attribute",
    ))
}

fn escaped_xml_attribute_len(value: &str, quote: u8) -> Result<usize, OlapError> {
    if !value.chars().all(valid_xml10_char) {
        return Err(OlapError::new(
            "XML attribute contains an XML 1.0-forbidden character",
        ));
    }
    value.chars().try_fold(0usize, |length, character| {
        let addition = match character {
            '&' => 5,
            '<' => 4,
            c if (c == '"' && quote == b'"') || (c == '\'' && quote == b'\'') => 6,
            _ => character.len_utf8(),
        };
        length
            .checked_add(addition)
            .ok_or_else(|| OlapError::new("XML attribute length overflow"))
    })
}

fn append_escaped_xml_attribute(
    output: &mut Vec<u8>,
    value: &str,
    quote: u8,
) -> Result<(), OlapError> {
    let _ = escaped_xml_attribute_len(value, quote)?;
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' if quote == b'"' => output.extend_from_slice(b"&quot;"),
            '\'' if quote == b'\'' => output.extend_from_slice(b"&apos;"),
            _ => {
                let mut encoded = [0_u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
    Ok(())
}

fn valid_xml10_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

/// Project validated section 2.5 metadata and section 2.6 OLAP definitions
/// into table-local identities.
///
/// A matching filename token is only one input. The containing dimension
/// folder, the `XMSimpleTable@name`/`XMRawColumn@name` values, the generated
/// column-data path, and the section 2.6 Dimension/Attribute projection are
/// retained as separate identities. Relationship foreign ownership comes from
/// the containing folder; `PrimaryTable` is checked independently. Duplicate
/// qualified endpoints are rejected as ambiguous rather than resolved by
/// order. The source-bound [`prove_xldm140_closure`] additionally checks the
/// section 2.2 Dimension file-group/ObjectID mapping.
pub fn project_xldm140_identity(
    metadata: &MetadataModel<'_>,
    olap: &OlapModel<'_>,
) -> Result<Xldm140IdentityProjection, OlapError> {
    if metadata.files.len() > MAX_IDENTITY_ITEMS || olap.files.len() > MAX_IDENTITY_ITEMS {
        return Err(OlapError::new("Xldm140 identity input limit exceeded"));
    }
    preflight_projection_identity_budget(metadata, olap)?;
    let dimensions = dimension_definitions(olap)?;
    let table_count = metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::Table)
        .count();
    let relationship_count = metadata.files.iter().try_fold(0usize, |count, file| {
        count
            .checked_add(file.table.collection("Relationships").map_or(0, <[_]>::len))
            .ok_or_else(|| OlapError::new("Xldm140 relationship count overflow"))
    })?;
    if relationship_count > MAX_IDENTITY_ITEMS {
        return Err(OlapError::new("Xldm140 relationship count limit exceeded"));
    }
    let mut tables = Vec::new();
    tables
        .try_reserve(table_count)
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 tables: {error}")))?;
    let mut table_ids = HashSet::new();
    table_ids
        .try_reserve(table_count)
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 table IDs: {error}")))?;

    for file in metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::Table)
    {
        let table_id = containing_table_id(file)?;
        if !table_ids.insert(table_id.clone()) {
            return Err(OlapError::new(format!(
                "duplicate Xldm140 table identity {table_id}"
            )));
        }
        let xml_name = file.table.name.clone().ok_or_else(|| {
            OlapError::new(format!(
                "table metadata {} has no XMSimpleTable name",
                file.storage_path
            ))
        })?;
        let dimension = dimensions.for_table(table_id.as_str()).ok_or_else(|| {
            OlapError::new(format!(
                "table {table_id} has no section 2.6 Dimension definition"
            ))
        })?;
        let columns = table_columns(file, &table_id, dimension.definition)?;
        tables.push(Xldm140TableIdentity {
            table_id,
            xml_name,
            metadata_path: file.storage_path.to_owned(),
            dimension_object_id: dimension.definition.object_id.clone(),
            attribute_ids: dimension.definition.attribute_ids.clone(),
            columns,
        });
    }

    let mut table_lookup = HashMap::new();
    table_lookup
        .try_reserve(tables.len())
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 table lookup: {error}")))?;
    for table in &tables {
        table_lookup.insert(table.table_id.as_str(), table);
    }
    let mut relationships = Vec::new();
    relationships
        .try_reserve(relationship_count)
        .map_err(|error| {
            OlapError::new(format!("cannot reserve Xldm140 relationships: {error}"))
        })?;
    let mut endpoint_keys = HashSet::new();
    endpoint_keys
        .try_reserve(relationship_count)
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 endpoints: {error}")))?;
    for file in &metadata.files {
        let Some(values) = file.table.collection("Relationships") else {
            continue;
        };
        if values.is_empty() {
            continue;
        }
        let containing_table = containing_table_id(file)?;
        let containing = table_lookup.get(containing_table.as_str()).ok_or_else(|| {
            OlapError::new(format!(
                "relationship metadata {} has no containing table identity",
                file.storage_path
            ))
        })?;
        for relationship in values {
            let primary_table = scalar_property(relationship, "PrimaryTable", file)?;
            let primary_column = scalar_property(relationship, "PrimaryColumn", file)?;
            let foreign_column = scalar_property(relationship, "ForeignColumn", file)?;
            if containing
                .columns
                .iter()
                .all(|column| column.column_id != foreign_column)
            {
                return Err(OlapError::new(format!(
                    "relationship {} foreign column {foreign_column} is absent from containing table {containing_table}",
                    file.storage_path
                )));
            }
            let primary = resolve_table_reference(&table_lookup, &primary_table, file)?;
            if primary
                .columns
                .iter()
                .all(|column| column.column_id != primary_column)
            {
                return Err(OlapError::new(format!(
                    "relationship {} primary column {primary_table}.{primary_column} is absent",
                    file.storage_path
                )));
            }
            let relationship_name = relationship.name.clone();
            let relationship_id = generated_relationship_id(file)?;
            let endpoint_key = (
                containing_table.clone(),
                foreign_column.clone(),
                primary_table.clone(),
                primary_column.clone(),
            );
            if !endpoint_keys.insert(endpoint_key) {
                return Err(OlapError::new(format!(
                    "ambiguous Xldm140 relationship endpoint mapping for {containing_table}.{foreign_column} -> {primary_table}.{primary_column}"
                )));
            }
            let expected_index_key = format!("R${containing_table}${relationship_id}");
            if file.kind == MetadataFileKind::TableRelationship
                && !relationship_path_matches(file.storage_path, &expected_index_key)
            {
                return Err(OlapError::new(format!(
                    "relationship metadata {} does not match containing table or relationship ID {expected_index_key}",
                    file.storage_path
                )));
            }
            relationships.push(Xldm140RelationshipIdentity {
                containing_table: containing_table.clone(),
                relationship_name,
                relationship_id,
                metadata_path: file.storage_path.to_owned(),
                primary_table,
                primary_column,
                foreign_column,
                expected_index_key,
            });
        }
    }

    Ok(Xldm140IdentityProjection {
        tables,
        relationships,
    })
}

/// Verify the native and section 2.4 generated members that close an
/// already-projected identity graph.
///
/// This check is intentionally separate from [`project_xldm140_identity`]:
/// callers that have not admitted native/generated members cannot accidentally
/// treat a syntactically valid path as a complete writable identity.
pub fn validate_xldm140_identity_closure(
    projection: &Xldm140IdentityProjection,
    native: &NativeModel<'_>,
    generated: &SystemGeneratedModel<'_>,
) -> Result<(), OlapError> {
    let expected_column_count = projection.tables.iter().try_fold(0usize, |count, table| {
        count
            .checked_add(table.columns.len())
            .ok_or_else(|| OlapError::new("Xldm140 column count overflow"))
    })?;
    if expected_column_count > MAX_IDENTITY_ITEMS {
        return Err(OlapError::new("Xldm140 column count limit exceeded"));
    }
    let mut expected_column_data = HashSet::new();
    expected_column_data
        .try_reserve(expected_column_count)
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 columns: {error}")))?;
    for column in projection
        .tables
        .iter()
        .flat_map(|table| table.columns.iter())
    {
        expected_column_data.insert(column.data_path.as_str());
    }
    for table in &projection.tables {
        for column in &table.columns {
            let mut matches = native
                .files
                .iter()
                .filter(|file| file.storage_path == column.data_path);
            if matches.next().is_none() {
                return Err(OlapError::new(format!(
                    "column identity {}.{} has no admitted native data member {}",
                    column.table_id, column.column_id, column.data_path
                )));
            }
            if matches.next().is_some() {
                return Err(OlapError::new(format!(
                    "column identity {}.{} has ambiguous native data member {}",
                    column.table_id, column.column_id, column.data_path
                )));
            }
            let generated = classify_generated_path(&column.data_path).map_err(|error| {
                OlapError::new(format!(
                    "column identity {}.{} has invalid data path {}: {error}",
                    column.table_id, column.column_id, column.data_path
                ))
            })?;
            if generated.kind != GeneratedNameKind::ColumnData {
                return Err(OlapError::new(format!(
                    "column identity {}.{} data member {} is not ColumnData",
                    column.table_id, column.column_id, column.data_path
                )));
            }
        }
    }
    for file in &native.files {
        let generated = classify_generated_path(file.storage_path).map_err(|error| {
            OlapError::new(format!(
                "native member {} has an invalid generated path: {error}",
                file.storage_path
            ))
        })?;
        if generated.kind == GeneratedNameKind::ColumnData
            && !expected_column_data.contains(file.storage_path)
        {
            return Err(OlapError::new(format!(
                "native column data member {} is absent from metadata identity closure",
                file.storage_path
            )));
        }
    }
    for relationship in &projection.relationships {
        let mut indexes = generated.files.iter().filter(|file| {
            file.kind == SystemGeneratedKind::RelationshipIndex
                && relationship_generated_object_key_matches(
                    &file.object_key,
                    &relationship.metadata_path,
                    &relationship.expected_index_key,
                )
        });
        let Some(index) = indexes.next() else {
            return Err(OlapError::new(format!(
                "relationship {} has no admitted section 2.4 index {}",
                relationship.metadata_path, relationship.expected_index_key
            )));
        };
        if indexes.next().is_some() {
            return Err(OlapError::new(format!(
                "relationship {} has ambiguous admitted section 2.4 index {}",
                relationship.metadata_path, relationship.expected_index_key
            )));
        }
        if !relationship_index_path_matches(
            index.storage_path,
            &relationship.containing_table,
            &relationship.expected_index_key,
        ) {
            return Err(OlapError::new(format!(
                "relationship {} index {} is outside its containing table",
                relationship.metadata_path, index.storage_path
            )));
        }
    }
    for file in &generated.files {
        if file.kind == SystemGeneratedKind::RelationshipIndex {
            let mut owners = projection.relationships.iter().filter(|relationship| {
                relationship_generated_object_key_matches(
                    &file.object_key,
                    &relationship.metadata_path,
                    &relationship.expected_index_key,
                )
            });
            let Some(_) = owners.next() else {
                return Err(OlapError::new(format!(
                    "generated relationship index {} is absent from metadata identity closure",
                    file.storage_path
                )));
            };
            if owners.next().is_some() {
                return Err(OlapError::new(format!(
                    "generated relationship index {} has ambiguous metadata ownership",
                    file.storage_path
                )));
            }
        }
    }
    Ok(())
}

/// Build a table-local projection and require its native/data and
/// relationship-index closure before reporting it as complete.
pub fn project_xldm140_identity_with_closure(
    metadata: &MetadataModel<'_>,
    olap: &OlapModel<'_>,
    native: &NativeModel<'_>,
    generated: &SystemGeneratedModel<'_>,
) -> Result<Xldm140IdentityProjection, OlapError> {
    let projection = project_xldm140_identity(metadata, olap)?;
    validate_xldm140_identity_closure(&projection, native, generated)?;
    validate_native_generated_member_layouts(native, generated)?;
    Ok(projection)
}

const MAX_CLOSURE_REFERENCES: usize = MAX_IDENTITY_ITEMS;

/// Prove the bounded cross-section closure and bind it to one inspected source.
///
/// This composes section 2.2 directory membership with the existing section
/// 2.3, 2.4, 2.5, and 2.6 inspectors. It retains indexes into the caller's
/// storage rather than copying member payloads. Unknown members are reported
/// in the returned graph and make [`Xldm140Closure::is_complete`] false; a
/// structural writer must refuse such a graph until it can preserve the
/// affected members.
pub fn prove_xldm140_closure<'storage, 'source>(
    storage: &'storage Storage<'source>,
    metadata: &'source MetadataModel<'source>,
    olap: &'source OlapModel<'source>,
    native: &'source NativeModel<'source>,
    generated: &'source SystemGeneratedModel<'source>,
) -> Result<Xldm140Closure<'storage, 'source>, OlapError> {
    if storage.profile() != StorageProfile::Xldm140 {
        return Err(OlapError::new(
            "Xldm140 identity closure requires the canonical version-140 profile",
        ));
    }
    let reference_count = metadata
        .files
        .len()
        .checked_add(native.files.len())
        .and_then(|count| count.checked_add(generated.files.len()))
        .and_then(|count| count.checked_add(olap.files.len()))
        .ok_or_else(|| OlapError::new("XLDM identity closure reference count overflow"))?;
    if reference_count > MAX_CLOSURE_REFERENCES {
        return Err(OlapError::new(
            "XLDM identity closure reference limit exceeded",
        ));
    }
    if storage.files.len() > MAX_CLOSURE_REFERENCES {
        return Err(OlapError::new(
            "XLDM identity closure directory limit exceeded",
        ));
    }
    preflight_closure_identity_budget(storage, metadata, olap, native, generated)?;
    // Hand-built unit fixtures may omit a BackupLog. A validated XLDM source
    // cannot: section 2.1.2.3 requires FileGroups, so every production source
    // takes the object-ID/file-group mapping path below.
    if !storage.backup_log.file_groups.is_empty() {
        validate_dimension_file_groups(storage, metadata, olap)?;
    }
    super::metadata::validate_files(metadata, &native.files, &generated.files)
        .map_err(|error| OlapError::new(format!("section 2.5 closure: {error}")))?;
    validate_system_generated_files(&generated.files)
        .map_err(|error| OlapError::new(format!("section 2.4 closure: {error}")))?;
    validate_native_generated_member_layouts(native, generated)?;
    super::olap::validate_with_storage(olap, metadata, storage)
        .map_err(|error| OlapError::new(format!("section 2.6 closure: {error}")))?;
    // The generic section-2.6 validator checks XML shape and the table-local
    // projection, but it cannot by itself see duplicate or unlinked
    // Dimension relationship references that are owned by the standalone
    // OLAP definition graph. A real OLAP source advertises that graph through
    // BackupLog.OlapInfo; run the complete source-bound proof here before the
    // closure becomes writable. Small hand-built projection fixtures leave
    // that flag unset and retain their intentionally narrower test contract.
    if storage.backup_log.is_olap {
        let olap_proof = super::olapproof::prove_xldm140_olap(
            storage,
            metadata,
            olap,
            super::olapproof::OlapProofLimits::default(),
        )
        .map_err(|error| OlapError::new(format!("section 2.6 OLAP proof: {error}")))?;
        if !olap_proof.is_complete() {
            return Err(OlapError::new(
                "section 2.6 OLAP proof contains unlinked or unknown members",
            ));
        }
    }
    let projection = project_xldm140_identity(metadata, olap)?;
    validate_xldm140_identity_closure(&projection, native, generated)?;

    let mut storage_paths = HashMap::new();
    storage_paths
        .try_reserve(storage.files.len())
        .map_err(|error| {
            OlapError::new(format!("cannot reserve XLDM storage path index: {error}"))
        })?;
    for (index, entry) in storage.files.iter().enumerate() {
        if storage_paths.insert(entry.path.as_str(), index).is_some() {
            return Err(OlapError::new(format!(
                "section 2.2 directory contains duplicate path {}",
                entry.path
            )));
        }
    }
    let mut sections = Vec::new();
    sections
        .try_reserve_exact(storage.files.len())
        .map_err(|error| OlapError::new(format!("cannot reserve XLDM closure owners: {error}")))?;
    sections.resize(storage.files.len(), None);
    for (index, entry) in storage.files.iter().enumerate() {
        if matches!(
            entry.kind,
            super::FileKind::Partitions
                | super::FileKind::BackupLog
                | super::FileKind::CryptographicKey
        ) {
            sections[index] = Some(Xldm140MemberSection::Outer);
        }
    }
    record_paths(
        metadata.files.iter().map(|file| file.storage_path),
        Xldm140MemberSection::Metadata,
        &storage_paths,
        &mut sections,
    )?;
    record_paths(
        native.files.iter().map(|file| file.storage_path),
        Xldm140MemberSection::Native,
        &storage_paths,
        &mut sections,
    )?;
    record_paths(
        generated.files.iter().map(|file| file.storage_path),
        Xldm140MemberSection::Generated,
        &storage_paths,
        &mut sections,
    )?;
    record_paths(
        olap.files.iter().map(|file| file.storage_path),
        Xldm140MemberSection::Olap,
        &storage_paths,
        &mut sections,
    )?;
    let mut members = Vec::new();
    members
        .try_reserve(storage.files.len())
        .map_err(|error| OlapError::new(format!("cannot reserve XLDM closure members: {error}")))?;
    let mut unknown_members = Vec::new();
    unknown_members
        .try_reserve(storage.files.len())
        .map_err(|error| OlapError::new(format!("cannot reserve XLDM unknown members: {error}")))?;
    for (storage_index, section) in sections.into_iter().enumerate() {
        let section = section.unwrap_or(Xldm140MemberSection::Other);
        let member = Xldm140ClosureMember {
            storage_index,
            section,
        };
        if section == Xldm140MemberSection::Other {
            unknown_members.push(storage_index);
        }
        members.push(member);
    }
    Ok(Xldm140Closure {
        storage,
        metadata: Some(metadata),
        native: Some(native),
        generated: Some(generated),
        projection,
        members,
        unknown_members,
    })
}

fn validate_native_generated_member_layouts(
    native: &NativeModel<'_>,
    generated: &SystemGeneratedModel<'_>,
) -> Result<(), OlapError> {
    for file in &native.files {
        let kind = classify_generated_path(file.storage_path)
            .map_err(|error| {
                OlapError::new(format!("native member {}: {error}", file.storage_path))
            })?
            .kind;
        let valid = match kind {
            GeneratedNameKind::ColumnData
            | GeneratedNameKind::ColumnPositionToId
            | GeneratedNameKind::ColumnIdToPosition
            | GeneratedNameKind::UserHierarchyChildCount
            | GeneratedNameKind::UserHierarchyFirstChildPosition
            | GeneratedNameKind::UserHierarchyParentPosition
            | GeneratedNameKind::UserHierarchyMultilevelId => {
                matches!(file.data, super::native::NativeData::Idf(_))
            },
            GeneratedNameKind::TableRelationshipIndex => matches!(
                file.data,
                super::native::NativeData::Idf(_) | super::native::NativeData::HashIndex(_)
            ),
            GeneratedNameKind::ColumnDictionary => {
                matches!(file.data, super::native::NativeData::Dictionary(_))
            },
            GeneratedNameKind::ColumnHashIndex => {
                matches!(file.data, super::native::NativeData::HashIndex(_))
            },
            _ => false,
        };
        if !valid {
            return Err(OlapError::new(format!(
                "native member {} has a typed layout incompatible with generated role {kind:?}",
                file.storage_path
            )));
        }
    }
    for file in &generated.files {
        let valid = match file.kind {
            SystemGeneratedKind::RelationshipIndex => matches!(
                file.data,
                super::generated::SystemGeneratedData::Idf(_)
                    | super::generated::SystemGeneratedData::HashIndex(_)
            ),
            SystemGeneratedKind::PositionToIdentifier
            | SystemGeneratedKind::IdentifierToPosition
            | SystemGeneratedKind::UserHierarchyChildCount
            | SystemGeneratedKind::UserHierarchyFirstChildPosition
            | SystemGeneratedKind::UserHierarchyMultilevelIdentifier
            | SystemGeneratedKind::UserHierarchyParentPosition => {
                matches!(file.data, super::generated::SystemGeneratedData::Idf(_))
            },
        };
        if !valid {
            return Err(OlapError::new(format!(
                "generated member {} has a typed layout incompatible with role {:?}",
                file.storage_path, file.kind
            )));
        }
    }
    Ok(())
}

fn validate_dimension_file_groups(
    storage: &Storage<'_>,
    metadata: &MetadataModel<'_>,
    olap: &OlapModel<'_>,
) -> Result<(), OlapError> {
    let dimensions = dimension_definitions(olap)?;
    let database = olap
        .files
        .iter()
        .filter_map(|file| match &file.document {
            OlapDocument::Definition(definition) if definition.kind == OlapObjectKind::Database => {
                Some(definition)
            },
            _ => None,
        })
        .next()
        .ok_or_else(|| {
            OlapError::new("section 2.2 has Dimension file groups but no OLAP Database")
        })?;
    let database_storage_path = database
        .object
        .scalar("DbStorageLocation")
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OlapError::new(
                "section 2.2 Dimension file groups cannot be mapped without Database DbStorageLocation",
            )
        })?;
    let database_storage_path =
        normalize_file_group_path(database_storage_path, "Database DbStorageLocation")?;
    let mut matched_dimensions = HashSet::new();
    matched_dimensions
        .try_reserve(dimensions.by_object_id.len())
        .map_err(|error| {
            OlapError::new(format!(
                "cannot reserve XLDM Dimension file-group identities: {error}"
            ))
        })?;
    for file in metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::Table)
    {
        let table_id = containing_table_id(file)?;
        let dimension = dimensions.for_table(table_id.as_str()).ok_or_else(|| {
            OlapError::new(format!(
                "table {table_id} has no section 2.6 Dimension file mapping"
            ))
        })?;
        let object_id = dimension.definition.object_id.as_str();
        if !matched_dimensions.insert(object_id) {
            return Err(OlapError::new(format!(
                "table {table_id} maps to a duplicate section 2.6 Dimension ObjectID {object_id}"
            )));
        }
        let group = find_dimension_file_group(storage, object_id, &table_id)?;
        validate_dimension_file_group(
            group,
            dimension,
            database,
            &database_storage_path,
            &table_id,
        )?;
    }
    for group in storage
        .backup_log
        .file_groups
        .iter()
        .filter(|group| group.class == FileGroupClass::Dimension)
    {
        let Some(dimension) = dimensions
            .by_object_id
            .get(group.object_id.as_str())
            .copied()
        else {
            return Err(OlapError::new(format!(
                "orphan section 2.2 Dimension file group ObjectID {}",
                group.object_id
            )));
        };
        let table_id = dimension_table_id(dimension.path)?;
        validate_dimension_file_group(
            group,
            dimension,
            database,
            &database_storage_path,
            table_id,
        )?;
        matched_dimensions.insert(group.object_id.as_str());
    }
    for dimension in dimensions.by_object_id.values() {
        if !matched_dimensions.contains(dimension.definition.object_id.as_str()) {
            return Err(OlapError::new(format!(
                "section 2.6 Dimension {} has no section 2.2 file group",
                dimension.definition.object_id
            )));
        }
    }
    Ok(())
}

fn find_dimension_file_group<'a>(
    storage: &'a Storage<'_>,
    object_id: &str,
    table_id: &str,
) -> Result<&'a super::FileGroup, OlapError> {
    let mut groups =
        storage.backup_log.file_groups.iter().filter(|group| {
            group.class == FileGroupClass::Dimension && group.object_id == object_id
        });
    let group = groups.next().ok_or_else(|| {
        OlapError::new(format!(
            "table {table_id} has no section 2.2 Dimension file group for ObjectID {object_id}"
        ))
    })?;
    if groups.next().is_some() {
        return Err(OlapError::new(format!(
            "table {table_id} has ambiguous section 2.2 Dimension file groups for ObjectID {object_id}"
        )));
    }
    Ok(group)
}

fn validate_dimension_file_group(
    group: &super::FileGroup,
    dimension: DimensionBinding<'_>,
    database: &super::olap::OlapDefinition,
    database_storage_path: &str,
    table_id: &str,
) -> Result<(), OlapError> {
    if group.object_id != dimension.definition.object_id {
        return Err(OlapError::new(format!(
            "table {table_id} Dimension file-group ObjectID {} disagrees with section 2.6 ObjectID {}",
            group.object_id, dimension.definition.object_id
        )));
    }
    if group.object_version != dimension.definition.extension.object_version {
        return Err(OlapError::new(format!(
            "table {table_id} Dimension file-group ObjectVersion {} disagrees with section 2.6 ObjectVersion {}",
            group.object_version, dimension.definition.extension.object_version
        )));
    }
    if group.persist_location != database.extension.persist_location {
        return Err(OlapError::new(format!(
            "table {table_id} Dimension file-group PersistLocation {} disagrees with Database PersistLocation {}",
            group.persist_location, database.extension.persist_location
        )));
    }
    let group_storage_path = normalize_file_group_path(
        &group.persist_location_path,
        "Dimension file-group PersistLocationPath",
    )?;
    if group_storage_path != database_storage_path {
        return Err(OlapError::new(format!(
            "table {table_id} Dimension file-group PersistLocationPath {group_storage_path} disagrees with Database DbStorageLocation {database_storage_path}"
        )));
    }
    let mut definitions = group
        .files
        .iter()
        .filter(|file| file.storage_path == dimension.path);
    if definitions.next().is_none() {
        return Err(OlapError::new(format!(
            "table {table_id} Dimension file group does not enumerate {}",
            dimension.path
        )));
    }
    if definitions.next().is_some() {
        return Err(OlapError::new(format!(
            "table {table_id} Dimension file group duplicates {}",
            dimension.path
        )));
    }
    Ok(())
}

fn normalize_file_group_path(path: &str, label: &str) -> Result<String, OlapError> {
    let path = path.trim_end_matches(['/', '\\']);
    if path.is_empty() {
        return Err(OlapError::new(format!("{label} cannot be empty")));
    }
    let marker = format!("{path}/placeholder.0.db.xml");
    let normalized = super::validation::normalize_generated_path(&marker)
        .map_err(|error| OlapError::new(format!("{label} is invalid: {error}")))?;
    Ok(normalized
        .rsplit_once('/')
        .map_or(normalized.as_str(), |(folder, _)| folder)
        .to_owned())
}

fn record_paths<'a>(
    paths: impl IntoIterator<Item = &'a str>,
    section: Xldm140MemberSection,
    storage_paths: &HashMap<&str, usize>,
    sections: &mut [Option<Xldm140MemberSection>],
) -> Result<(), OlapError> {
    for path in paths {
        let index = *storage_paths.get(path).ok_or_else(|| {
            OlapError::new(format!(
                "typed XLDM closure member {path} is absent from section 2.2 directory"
            ))
        })?;
        let slot = &mut sections[index];
        *slot = Some(match (*slot, section) {
            (None, section) => section,
            (Some(previous), same) if previous == same => {
                return Err(OlapError::new(format!(
                    "typed XLDM closure member {path} is duplicated in {section:?}"
                )));
            },
            (Some(Xldm140MemberSection::Native), Xldm140MemberSection::Generated)
            | (Some(Xldm140MemberSection::Generated), Xldm140MemberSection::Native) => {
                Xldm140MemberSection::NativeAndGenerated
            },
            (Some(previous), current) => {
                return Err(OlapError::new(format!(
                    "typed XLDM closure member {path} has conflicting owners {previous:?} and {current:?}"
                )));
            },
        });
    }
    Ok(())
}

#[derive(Default)]
struct IdentityStringBudget {
    bytes: usize,
}

impl IdentityStringBudget {
    fn charge(&mut self, value: &str, label: &str) -> Result<(), OlapError> {
        self.bytes = self
            .bytes
            .checked_add(value.len())
            .ok_or_else(|| OlapError::new(format!("Xldm140 {label} byte count overflow")))?;
        if self.bytes > MAX_IDENTITY_BYTES {
            return Err(OlapError::new(format!(
                "Xldm140 identity string budget exceeded while charging {label}"
            )));
        }
        Ok(())
    }
}

fn preflight_projection_identity_budget(
    metadata: &MetadataModel<'_>,
    olap: &OlapModel<'_>,
) -> Result<(), OlapError> {
    let mut budget = IdentityStringBudget::default();
    charge_projection_identity_strings(metadata, olap, &mut budget)
}

fn charge_projection_identity_strings(
    metadata: &MetadataModel<'_>,
    olap: &OlapModel<'_>,
    budget: &mut IdentityStringBudget,
) -> Result<(), OlapError> {
    charge_metadata_identity_strings(metadata, budget)?;
    charge_olap_identity_strings(olap, budget)?;
    Ok(())
}

fn preflight_closure_identity_budget(
    storage: &Storage<'_>,
    metadata: &MetadataModel<'_>,
    olap: &OlapModel<'_>,
    native: &NativeModel<'_>,
    generated: &SystemGeneratedModel<'_>,
) -> Result<(), OlapError> {
    let mut budget = IdentityStringBudget::default();
    for value in [
        storage.backup_log.server_root.as_str(),
        storage.backup_log.object_name.as_str(),
        storage.backup_log.object_id.as_str(),
    ] {
        budget.charge(value, "backup-log identity")?;
    }
    for group in &storage.backup_log.file_groups {
        for value in [
            group.id.as_str(),
            group.name.as_str(),
            group.persist_location_path.as_str(),
            group.storage_location_path.as_str(),
            group.object_id.as_str(),
        ] {
            budget.charge(value, "file-group identity")?;
        }
        for file in &group.files {
            for value in [
                file.source_path.as_str(),
                file.storage_path.as_str(),
                file.generated.normalized_path.as_str(),
            ] {
                budget.charge(value, "logged-member identity")?;
            }
        }
    }
    for entry in &storage.files {
        budget.charge(entry.path.as_str(), "directory path")?;
    }
    charge_projection_identity_strings(metadata, olap, &mut budget)?;
    for file in &native.files {
        budget.charge(file.storage_path, "native path")?;
    }
    for file in &generated.files {
        budget.charge(file.storage_path, "generated path")?;
        budget.charge(file.object_key.as_str(), "generated object identity")?;
    }
    Ok(())
}

fn charge_metadata_identity_strings(
    metadata: &MetadataModel<'_>,
    budget: &mut IdentityStringBudget,
) -> Result<(), OlapError> {
    for file in &metadata.files {
        budget.charge(file.storage_path, "metadata path")?;
        charge_metadata_object_identity_strings(&file.table, budget)?;
    }
    for column in &metadata.columns {
        budget.charge(&column.name, "column policy name")?;
        budget.charge(&column.data_file, "column data path")?;
        if let Some(dictionary) = &column.dictionary {
            budget.charge(&dictionary.storage_name, "dictionary path")?;
            budget.charge(dictionary.class.as_str(), "dictionary class")?;
        }
    }
    for relationship in &metadata.relationships {
        if let Some(name) = &relationship.name {
            budget.charge(name, "relationship name")?;
        }
        for value in [
            relationship.primary_table.as_str(),
            relationship.primary_column.as_str(),
            relationship.foreign_column.as_str(),
        ] {
            budget.charge(value, "relationship endpoint")?;
        }
    }
    for hierarchy in &metadata.hierarchies {
        budget.charge(&hierarchy.table_store, "hierarchy store")?;
        for level in &hierarchy.level_ids {
            budget.charge(level, "hierarchy level")?;
        }
    }
    Ok(())
}

fn charge_metadata_object_identity_strings(
    object: &MetadataObject,
    budget: &mut IdentityStringBudget,
) -> Result<(), OlapError> {
    budget.charge(object.class.as_str(), "metadata class")?;
    if let Some(name) = &object.name {
        budget.charge(name, "metadata object name")?;
    }
    for property in &object.properties {
        budget.charge(&property.name, "metadata property name")?;
        budget.charge(&property.value, "metadata property value")?;
    }
    for member in &object.members {
        budget.charge(&member.name, "metadata member name")?;
        charge_metadata_object_identity_strings(&member.object, budget)?;
    }
    for collection in &object.collections {
        budget.charge(&collection.name, "metadata collection name")?;
        for item in &collection.objects {
            charge_metadata_object_identity_strings(item, budget)?;
        }
    }
    for data_object in &object.data_objects {
        charge_metadata_object_identity_strings(&data_object.object, budget)?;
    }
    Ok(())
}

fn charge_olap_identity_strings(
    olap: &OlapModel<'_>,
    budget: &mut IdentityStringBudget,
) -> Result<(), OlapError> {
    for file in &olap.files {
        budget.charge(file.storage_path, "OLAP path")?;
        if let OlapDocument::Definition(definition) = &file.document {
            budget.charge(&definition.object_id, "OLAP object ID")?;
            if let Some(name) = &definition.object_name {
                budget.charge(name, "OLAP object name")?;
            }
            if let Some(database) = &definition.parent.database_id {
                budget.charge(database, "OLAP database parent")?;
            }
            if let Some(cube) = &definition.parent.cube_id {
                budget.charge(cube, "OLAP cube parent")?;
            }
            charge_olap_element_identity_strings(&definition.object, budget)?;
            for value in definition.extension.data_files.iter().chain(
                definition
                    .extension
                    .permission_files
                    .iter()
                    .chain(definition.extension.measure_group_files.iter())
                    .chain(definition.extension.perspective_files.iter())
                    .chain(definition.extension.assembly_files.iter())
                    .chain(definition.extension.aggregation_design_files.iter())
                    .chain(definition.extension.partition_files.iter()),
            ) {
                budget.charge(value, "OLAP file-list path")?;
            }
            for attribute in &definition.attribute_ids {
                budget.charge(attribute, "OLAP attribute ID")?;
            }
            for hierarchy in &definition.hierarchies {
                budget.charge(&hierarchy.id, "OLAP hierarchy ID")?;
                for level in &hierarchy.level_ids {
                    budget.charge(level, "OLAP hierarchy level")?;
                }
            }
        }
    }
    Ok(())
}

fn charge_olap_element_identity_strings(
    element: &super::olap::OlapElement,
    budget: &mut IdentityStringBudget,
) -> Result<(), OlapError> {
    budget.charge(&element.name, "OLAP element name")?;
    budget.charge(&element.text, "OLAP element text")?;
    for (name, value) in &element.attributes {
        budget.charge(name, "OLAP attribute name")?;
        budget.charge(value, "OLAP attribute value")?;
    }
    for child in &element.children {
        charge_olap_element_identity_strings(child, budget)?;
    }
    Ok(())
}

fn dimension_definitions<'a>(
    model: &'a OlapModel<'_>,
) -> Result<DimensionDefinitions<'a>, OlapError> {
    let mut by_path = HashMap::new();
    let mut by_object_id = HashMap::new();
    let alias_count = model
        .files
        .len()
        .checked_mul(3)
        .ok_or_else(|| OlapError::new("Xldm140 Dimension alias count overflow"))?;
    by_path
        .try_reserve(alias_count)
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 dimensions: {error}")))?;
    by_object_id
        .try_reserve(model.files.len())
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 object IDs: {error}")))?;
    for file in &model.files {
        let OlapDocument::Definition(definition) = &file.document else {
            continue;
        };
        if definition.kind != OlapObjectKind::Dimension {
            continue;
        }
        let table_id = dimension_table_id(file.storage_path)?;
        let binding = DimensionBinding {
            path: file.storage_path,
            definition,
        };
        insert_unique_dimension_path(&mut by_path, table_id, binding)?;
        // The generated basename is the section-2.2 TableID and stays a
        // separate identity from the OLAP ObjectID/object name. Keep the
        // object-ID index for file-group ownership checks, but never use a
        // coincidental ObjectID alias to resolve a table's Dimension.
        if let Some(object_id) =
            (!definition.object_id.is_empty()).then_some(definition.object_id.as_str())
        {
            insert_unique_dimension_object_id(&mut by_object_id, object_id, binding)?;
        }
    }
    Ok(DimensionDefinitions {
        by_path,
        by_object_id,
    })
}

fn insert_unique_dimension_path<'a>(
    result: &mut HashMap<&'a str, DimensionBinding<'a>>,
    alias: &'a str,
    binding: DimensionBinding<'a>,
) -> Result<(), OlapError> {
    if let Some(previous) = result.get(alias).copied() {
        if std::ptr::eq(previous.definition, binding.definition) {
            return Ok(());
        }
        return Err(OlapError::new(format!(
            "duplicate section 2.6 Dimension path identity {alias}"
        )));
    }
    result.insert(alias, binding);
    Ok(())
}

fn insert_unique_dimension_object_id<'a>(
    result: &mut HashMap<&'a str, DimensionBinding<'a>>,
    alias: &'a str,
    binding: DimensionBinding<'a>,
) -> Result<(), OlapError> {
    if let Some(previous) = result.get(alias).copied()
        && !std::ptr::eq(previous.definition, binding.definition)
    {
        return Err(OlapError::new(format!(
            "duplicate section 2.6 Dimension ObjectID {alias}"
        )));
    }
    result.insert(alias, binding);
    Ok(())
}

struct DimensionDefinitions<'a> {
    by_path: HashMap<&'a str, DimensionBinding<'a>>,
    by_object_id: HashMap<&'a str, DimensionBinding<'a>>,
}

impl<'a> DimensionDefinitions<'a> {
    fn for_table(&self, table_id: &str) -> Option<DimensionBinding<'a>> {
        self.by_path.get(table_id).copied()
    }
}

#[derive(Clone, Copy)]
struct DimensionBinding<'a> {
    path: &'a str,
    definition: &'a super::olap::OlapDefinition,
}

fn dimension_table_id(path: &str) -> Result<&str, OlapError> {
    let generated = classify_generated_path(path)
        .map_err(|error| OlapError::new(format!("invalid Dimension path {path}: {error}")))?;
    if generated.kind != GeneratedNameKind::DataSourceOrDimensionDefinition {
        return Err(OlapError::new(format!(
            "section 2.6 Dimension path {path} has the wrong generated role"
        )));
    }
    let basename = path.rsplit('/').next().unwrap_or(path);
    let stem = basename.strip_suffix(".dim.xml").ok_or_else(|| {
        OlapError::new(format!(
            "section 2.6 Dimension path {path} has no .dim.xml suffix"
        ))
    })?;
    let (table_id, version) = stem.rsplit_once('.').ok_or_else(|| {
        OlapError::new(format!("section 2.6 Dimension path {path} has no version"))
    })?;
    if table_id.is_empty()
        || version.is_empty()
        || !version.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(OlapError::new(format!(
            "section 2.6 Dimension path {path} has an invalid generated name"
        )));
    }
    Ok(table_id)
}

fn resolve_table_reference<'a>(
    tables: &HashMap<&str, &'a Xldm140TableIdentity>,
    reference: &str,
    file: &MetadataFile<'_>,
) -> Result<&'a Xldm140TableIdentity, OlapError> {
    let mut matches = tables
        .values()
        .copied()
        .filter(|table| table.xml_name == reference);
    let Some(table) = matches.next() else {
        return Err(OlapError::new(format!(
            "relationship {} primary table {reference} is absent",
            file.storage_path
        )));
    };
    if matches.next().is_some() {
        return Err(OlapError::new(format!(
            "relationship {} primary table reference {reference} is ambiguous",
            file.storage_path
        )));
    }
    Ok(table)
}

fn table_columns(
    file: &MetadataFile<'_>,
    table_id: &str,
    dimension: &super::olap::OlapDefinition,
) -> Result<Vec<Xldm140ColumnIdentity>, OlapError> {
    let collection = file.table.collection("Columns").ok_or_else(|| {
        OlapError::new(format!(
            "table metadata {} has no Columns collection",
            file.storage_path
        ))
    })?;
    let mut result = Vec::new();
    result
        .try_reserve(collection.len())
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 columns: {error}")))?;
    let mut seen = HashSet::new();
    seen.try_reserve(collection.len())
        .map_err(|error| OlapError::new(format!("cannot reserve Xldm140 column IDs: {error}")))?;
    for column in collection {
        let column_id = column.name.clone().ok_or_else(|| {
            OlapError::new(format!(
                "table metadata {} contains an unnamed XMRawColumn",
                file.storage_path
            ))
        })?;
        if !seen.insert(column_id.clone()) {
            return Err(OlapError::new(format!(
                "duplicate column identity {table_id}.{column_id}"
            )));
        }
        // MS-XLDM 2.6.6 requires one DimensionAttribute for every table
        // column.  Keep this explicit section-2.5-to-2.6 edge in the
        // projection; a matching generated data path alone cannot prove the
        // column belongs to this table's OLAP dimension.
        if !dimension
            .attribute_ids
            .iter()
            .any(|attribute| attribute == &column_id)
        {
            return Err(OlapError::new(format!(
                "table column {table_id}.{column_id} is absent from its section 2.6 Dimension attributes"
            )));
        }
        let column_name = dimension_attribute_name(dimension, &column_id)?;
        let db_type = column
            .member("ColumnStats")
            .map(|stats| {
                let value = stats.property("DBType").ok_or_else(|| {
                    OlapError::new(format!(
                        "column identity {table_id}.{column_id} has ColumnStats without DBType"
                    ))
                })?;
                value.parse::<u16>().map_err(|error| {
                    OlapError::new(format!(
                        "column identity {table_id}.{column_id} has invalid DBType: {error}"
                    ))
                })
            })
            .transpose()?;
        let mut partitions = column
            .data_objects
            .iter()
            .filter(|item| item.object.class.as_str() == "XMRawColumnPartitionDataObject");
        let partition = partitions.next().ok_or_else(|| {
            OlapError::new(format!(
                "column identity {table_id}.{column_id} has no partition data object"
            ))
        })?;
        if partitions.next().is_some() {
            return Err(OlapError::new(format!(
                "column identity {table_id}.{column_id} has ambiguous partition data objects"
            )));
        }
        let storage_name = partition.object.name.as_deref().ok_or_else(|| {
            OlapError::new(format!(
                "column identity {table_id}.{column_id} has no partition storage name"
            ))
        })?;
        let data_path = qualify_storage_name(file.storage_path, storage_name);
        validate_column_data_path(&data_path, table_id, &column_id)?;
        // `MetadataModel` instances produced by the section-2.5 inspector
        // always carry Settings. Small projection fixtures used by callers
        // may intentionally omit optional metadata, in which case the
        // conservative classification is a source column. A present value is
        // still parsed strictly; malformed validated input cannot be bound.
        let settings = column
            .property("Settings")
            .map(|value| {
                value.parse::<u64>().map_err(|error| {
                    OlapError::new(format!(
                        "column identity {table_id}.{column_id} has invalid Settings: {error}"
                    ))
                })
            })
            .transpose()?
            .unwrap_or(0);
        result.push(Xldm140ColumnIdentity {
            table_id: table_id.to_owned(),
            column_id,
            column_name,
            db_type,
            data_path,
            settings,
            is_calculated: settings & 0x1f == 0x2 || settings & 0x800 != 0,
        });
    }
    for attribute in &dimension.attribute_ids {
        if !seen.contains(attribute) {
            return Err(OlapError::new(format!(
                "section 2.6 Dimension attribute {table_id}.{attribute} has no table column"
            )));
        }
    }
    Ok(result)
}

/// Resolve the optional display name of one Dimension Attribute without
/// treating a generated or XML object name as an outer immutable ID. The
/// canonical extractor always validates the Attribute IDs separately; this
/// helper only adds the explicit base-OLAP name edge when it is present.
fn dimension_attribute_name(
    dimension: &super::olap::OlapDefinition,
    column_id: &str,
) -> Result<Option<String>, OlapError> {
    let Some(attributes) = dimension.object.child("Attributes") else {
        return Ok(None);
    };
    let mut matched = None;
    for attribute in &attributes.children {
        if attribute.name != "Attribute" {
            return Err(OlapError::new(format!(
                "section 2.6 Dimension {} contains a non-Attribute member",
                dimension.object_id
            )));
        }
        let Some(id) = attribute.scalar("ID") else {
            return Err(OlapError::new(format!(
                "section 2.6 Dimension {} has an Attribute without an ID",
                dimension.object_id
            )));
        };
        if id != column_id {
            continue;
        }
        if matched.is_some() {
            return Err(OlapError::new(format!(
                "section 2.6 Dimension {} duplicates Attribute ID {column_id}",
                dimension.object_id
            )));
        }
        let name = attribute.scalar("Name").map(str::to_owned);
        if name.as_deref().is_some_and(str::is_empty) {
            return Err(OlapError::new(format!(
                "section 2.6 Dimension {} has an empty Attribute name for {column_id}",
                dimension.object_id
            )));
        }
        matched = Some(name);
    }
    matched.ok_or_else(|| {
        OlapError::new(format!(
            "section 2.6 Dimension {} has no Attribute with ID {column_id}",
            dimension.object_id
        ))
    })
}

fn validate_column_data_path(path: &str, table_id: &str, column_id: &str) -> Result<(), OlapError> {
    let generated = classify_generated_path(path)
        .map_err(|error| OlapError::new(format!("invalid column data path {path}: {error}")))?;
    if generated.kind != GeneratedNameKind::ColumnData {
        return Err(OlapError::new(format!(
            "column data path {path} is not a section 2.3 column member"
        )));
    }
    let Some(folder) = path.split('/').next_back() else {
        return Err(OlapError::new("column data path has no containing folder"));
    };
    let expected_folder = format!("{table_id}.0.dim");
    let parent = path
        .rsplit_once('/')
        .map(|(parent, _)| parent.rsplit('/').next().unwrap_or(parent));
    if parent != Some(expected_folder.as_str()) {
        return Err(OlapError::new(format!(
            "column data path {path} is outside containing table {table_id}"
        )));
    }
    let Some(stem) = folder.strip_suffix(".idf") else {
        return Err(OlapError::new(format!(
            "column data path {path} has no .idf suffix"
        )));
    };
    let Some((version, rest)) = stem.split_once('.') else {
        return Err(OlapError::new(format!(
            "column data path {path} has no version"
        )));
    };
    if version.is_empty() || !version.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(OlapError::new(format!(
            "column data path {path} has an invalid version"
        )));
    }
    let prefix = format!("{table_id}.");
    let Some(rest) = rest.strip_prefix(&prefix) else {
        return Err(OlapError::new(format!(
            "column data path {path} names the wrong table"
        )));
    };
    let Some(actual_column) = rest.strip_suffix(".0") else {
        return Err(OlapError::new(format!(
            "column data path {path} has an invalid partition"
        )));
    };
    if actual_column != column_id {
        return Err(OlapError::new(format!(
            "column data path {path} names column {actual_column}, expected {column_id}"
        )));
    }
    Ok(())
}

fn containing_table_id(file: &MetadataFile<'_>) -> Result<String, OlapError> {
    let generated = classify_generated_path(file.storage_path).map_err(|error| {
        OlapError::new(format!(
            "invalid metadata path {}: {error}",
            file.storage_path
        ))
    })?;
    if !matches!(
        generated.kind,
        GeneratedNameKind::TableMetadata
            | GeneratedNameKind::TableRelationshipMetadata
            | GeneratedNameKind::ColumnHierarchyMetadata
            | GeneratedNameKind::UserHierarchyMetadata
    ) {
        return Err(OlapError::new(format!(
            "{} is not a table-local metadata path",
            file.storage_path
        )));
    }
    let parent = file
        .storage_path
        .rsplit_once('/')
        .map(|(parent, _)| parent.rsplit('/').next().unwrap_or(parent))
        .ok_or_else(|| {
            OlapError::new(format!(
                "metadata {} has no containing dimension folder",
                file.storage_path
            ))
        })?;
    let table_id = parent.strip_suffix(".0.dim").ok_or_else(|| {
        OlapError::new(format!(
            "metadata {} has no canonical dimension folder",
            file.storage_path
        ))
    })?;
    if table_id.is_empty() {
        return Err(OlapError::new(format!(
            "metadata {} has an empty TableID",
            file.storage_path
        )));
    }
    Ok(table_id.to_owned())
}

fn relationship_metadata_owned_by(
    file: &MetadataFile<'_>,
    table_id: &str,
) -> Result<bool, OlapError> {
    if file.kind != MetadataFileKind::TableRelationship {
        return Ok(false);
    }
    // Relationship metadata is rooted in the dimension folder for the table
    // on the many/foreign side. The root XMSimpleTable name is a descriptor
    // value and can be stale or equal to another XML table name, so use the
    // generated containing folder to select every root owned by the renamed
    // TableID. Endpoint rewrites below separately cover a renamed table used
    // as PrimaryTable (the secondary/one side).
    Ok(containing_table_id(file)? == table_id)
}

fn scalar_property(
    object: &MetadataObject,
    name: &str,
    file: &MetadataFile<'_>,
) -> Result<String, OlapError> {
    object
        .property(name)
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
        .ok_or_else(|| {
            OlapError::new(format!(
                "relationship metadata {} has no {name} property",
                file.storage_path
            ))
        })
}

fn generated_relationship_id(file: &MetadataFile<'_>) -> Result<String, OlapError> {
    if file.kind != MetadataFileKind::TableRelationship {
        return Err(OlapError::new(format!(
            "relationship metadata {} is not a generated TableRelationship member",
            file.storage_path
        )));
    }
    let name = file
        .storage_path
        .rsplit('/')
        .next()
        .unwrap_or(file.storage_path);
    let stem = name
        .strip_prefix("R$")
        .and_then(|value| value.strip_suffix(".tbl.xml"))
        .ok_or_else(|| {
            OlapError::new(format!(
                "relationship metadata {} has no generated RelId filename",
                file.storage_path
            ))
        })?;
    let (_, rest) = stem.split_once('$').ok_or_else(|| {
        OlapError::new(format!(
            "relationship metadata {} has no generated RelId",
            file.storage_path
        ))
    })?;
    let (relation, version) = rest.rsplit_once('.').ok_or_else(|| {
        OlapError::new(format!(
            "relationship metadata {} has no generated RelId version",
            file.storage_path
        ))
    })?;
    if relation.is_empty()
        || version.is_empty()
        || !version.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(OlapError::new(format!(
            "relationship metadata {} has an invalid generated RelId version",
            file.storage_path
        )));
    }
    Ok(relation.to_owned())
}

fn relationship_path_matches(path: &str, expected_key: &str) -> bool {
    path.rsplit('/').next().is_some_and(|name| {
        name.strip_suffix(".tbl.xml").is_some_and(|stem| {
            stem.rsplit_once('.')
                .is_some_and(|(key, _)| key == expected_key)
        })
    })
}

fn relationship_index_path_matches(path: &str, containing_table: &str, expected_key: &str) -> bool {
    let Some((parent, basename)) = path.rsplit_once('/') else {
        return false;
    };
    let parent = parent.rsplit('/').next().unwrap_or(parent);
    if parent != format!("{containing_table}.0.dim") {
        return false;
    }
    let Some(stem) = basename.strip_suffix(".idf") else {
        return false;
    };
    let Some((version, suffix)) = stem.split_once('.') else {
        return false;
    };
    !version.is_empty()
        && version.bytes().all(|byte| byte.is_ascii_digit())
        && suffix == format!("{expected_key}.INDEX.0")
}

/// Match a generated relationship index to its qualified metadata owner.
///
/// `SystemGeneratedFile::object_key` is qualified for files discovered from
/// storage (the parser retains the generated member's parent and ordinal),
/// while a few older callers construct typed models with the legacy bare
/// `R$...` identity.  The qualified form is required to retain the containing
/// dimension ownership; the bare form remains accepted only for that legacy
/// representation and is still checked against the physical index path by
/// the closure validator.
fn relationship_generated_object_key_matches(
    object_key: &str,
    metadata_path: &str,
    expected_key: &str,
) -> bool {
    let Some((object_parent, object_name)) = object_key.rsplit_once('/') else {
        return object_key == expected_key;
    };
    let Some((metadata_parent, _)) = metadata_path.rsplit_once('/') else {
        return false;
    };
    if object_parent != metadata_parent {
        return false;
    }
    if object_name == expected_key {
        return true;
    }
    let Some((ordinal, identity)) = object_name.split_once('.') else {
        return false;
    };
    !ordinal.is_empty()
        && ordinal.bytes().all(|byte| byte.is_ascii_digit())
        && identity == expected_key
}

fn qualify_storage_name(metadata_path: &str, name: &str) -> String {
    if name.contains('/') {
        name.to_owned()
    } else if let Some((parent, _)) = metadata_path.rsplit_once('/') {
        format!("{parent}/{name}")
    } else {
        name.to_owned()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::generated::{SystemGeneratedData, SystemGeneratedFile, SystemGeneratedModel};
    use crate::metadata::{ColumnPolicy, MetadataClass, MetadataCollection, MetadataDataObject};
    use crate::model::{BOM, CRC_SIZE, test_xldm140_storage};
    use crate::native::{
        DictionaryBody, DictionaryFile, DictionaryType, IdfFile, NativeData, NativeFile,
        NativeModel, NumericDictionary,
    };
    use crate::olap::{
        OlapDefinition, OlapElement, OlapFile, OlapFileKind, OlapParentReference, TabularExtension,
    };
    use crate::{FileGroup, FileKind, GeneratedPath, LoggedFile};

    #[test]
    fn rename_error_preserves_nested_proof_limit_resource() {
        let error = Xldm140RenameError::from_olap_proof(crate::OlapProofError::LimitExceeded {
            resource: "graph work",
            actual: 17,
            maximum: 16,
        });

        assert_eq!(error.kind(), Xldm140RenameErrorKind::LimitExceeded);
        assert_eq!(error.limit_bounds(), Some((17, 16)));
        assert_eq!(error.limit_resource(), Some("graph work"));
    }

    fn object(
        class: &str,
        name: Option<&str>,
        properties: Vec<(&str, &str)>,
        collections: Vec<(&str, Vec<MetadataObject>)>,
        data_objects: Vec<MetadataObject>,
    ) -> MetadataObject {
        MetadataObject {
            class: MetadataClass(class.into()),
            name: name.map(str::to_owned),
            provider_version: None,
            properties: properties
                .into_iter()
                .map(|(name, value)| crate::metadata::MetadataProperty {
                    name: name.into(),
                    value: value.into(),
                })
                .collect(),
            members: Vec::new(),
            collections: collections
                .into_iter()
                .map(|(name, objects)| MetadataCollection {
                    name: name.into(),
                    objects,
                })
                .collect(),
            data_objects: data_objects
                .into_iter()
                .map(|object| MetadataDataObject {
                    object: Box::new(object),
                })
                .collect(),
        }
    }

    fn column(table: &str, name: &str) -> MetadataObject {
        object(
            "XMRawColumn",
            Some(name),
            vec![],
            vec![],
            vec![object(
                "XMRawColumnPartitionDataObject",
                Some(&format!("1.{table}.{name}.0.idf")),
                vec![],
                vec![],
                vec![],
            )],
        )
    }

    fn table(path_table: &str, columns: Vec<MetadataObject>) -> MetadataFile<'static> {
        let table = object(
            "XMSimpleTable",
            Some(path_table),
            vec![],
            vec![("Columns", columns), ("Relationships", Vec::new())],
            vec![],
        );
        MetadataFile {
            storage_path: Box::leak(
                format!("Model.1.db/{path_table}.0.dim/{path_table}.1.tbl.xml").into_boxed_str(),
            ),
            bytes: &[],
            kind: MetadataFileKind::Table,
            table,
        }
    }

    fn relationship_file(
        containing: &str,
        relationship: &str,
        primary: &str,
    ) -> MetadataFile<'static> {
        let relation = object(
            "XMRelationship",
            Some(relationship),
            vec![
                ("PrimaryTable", primary),
                ("PrimaryColumn", "Key"),
                ("ForeignColumn", "Key"),
            ],
            vec![],
            vec![],
        );
        let table = object(
            "XMSimpleTable",
            Some(containing),
            vec![],
            vec![("Relationships", vec![relation])],
            vec![],
        );
        MetadataFile {
            storage_path: Box::leak(
                format!("Model.1.db/{containing}.0.dim/R${containing}${relationship}.1.tbl.xml")
                    .into_boxed_str(),
            ),
            bytes: &[],
            kind: MetadataFileKind::TableRelationship,
            table,
        }
    }

    fn dimension(id: &str) -> OlapDefinition {
        dimension_with_attributes(id, vec!["Key"])
    }

    fn dimension_with_attributes(id: &str, attribute_ids: Vec<&str>) -> OlapDefinition {
        OlapDefinition {
            parent: OlapParentReference::default(),
            kind: OlapObjectKind::Dimension,
            object_id: id.into(),
            object_name: Some(id.into()),
            object: OlapElement {
                name: "Dimension".into(),
                attributes: Vec::new(),
                text: String::new(),
                children: Vec::new(),
            },
            extension: TabularExtension {
                ordinal: 0,
                object_version: 1,
                persist_location: 0,
                data_files: Vec::new(),
                permission_files: Vec::new(),
                measure_group_files: Vec::new(),
                perspective_files: Vec::new(),
                assembly_files: Vec::new(),
                aggregation_design_files: Vec::new(),
                partition_files: Vec::new(),
                default_collation_version: None,
            },
            attribute_ids: attribute_ids.into_iter().map(str::to_owned).collect(),
            hierarchies: Vec::new(),
        }
    }

    fn dimension_with_named_attributes(id: &str, attributes: &[(&str, &str)]) -> OlapDefinition {
        let mut definition =
            dimension_with_attributes(id, attributes.iter().map(|(id, _)| *id).collect());
        definition.object.children.push(OlapElement {
            name: "Attributes".into(),
            attributes: Vec::new(),
            text: String::new(),
            children: attributes
                .iter()
                .map(|(attribute_id, name)| OlapElement {
                    name: "Attribute".into(),
                    attributes: Vec::new(),
                    text: String::new(),
                    children: vec![
                        OlapElement {
                            name: "ID".into(),
                            attributes: Vec::new(),
                            text: (*attribute_id).into(),
                            children: Vec::new(),
                        },
                        OlapElement {
                            name: "Name".into(),
                            attributes: Vec::new(),
                            text: (*name).into(),
                            children: Vec::new(),
                        },
                    ],
                })
                .collect(),
        });
        definition
    }

    fn olap(dimensions: Vec<OlapDefinition>) -> OlapModel<'static> {
        let files = dimensions
            .into_iter()
            .map(|definition| {
                let path = Box::leak(
                    format!("Model.1.db/{}.1.dim.xml", definition.object_id).into_boxed_str(),
                );
                OlapFile {
                    storage_path: path,
                    bytes: &[],
                    kind: OlapFileKind::Definition(OlapObjectKind::Dimension),
                    document: OlapDocument::Definition(definition),
                }
            })
            .collect();
        OlapModel { files }
    }

    fn closure_definition(
        storage_path: &'static str,
        kind: OlapObjectKind,
        object_id: &'static str,
        parent: OlapParentReference,
    ) -> OlapFile<'static> {
        OlapFile {
            storage_path,
            bytes: &[],
            kind: OlapFileKind::Definition(kind),
            document: OlapDocument::Definition(OlapDefinition {
                parent,
                kind,
                object_id: object_id.into(),
                object_name: Some(object_id.into()),
                object: OlapElement {
                    name: "Object".into(),
                    attributes: Vec::new(),
                    text: String::new(),
                    children: Vec::new(),
                },
                extension: TabularExtension {
                    ordinal: 0,
                    object_version: 1,
                    persist_location: 0,
                    data_files: Vec::new(),
                    permission_files: Vec::new(),
                    measure_group_files: Vec::new(),
                    perspective_files: Vec::new(),
                    assembly_files: Vec::new(),
                    aggregation_design_files: Vec::new(),
                    partition_files: Vec::new(),
                    default_collation_version: None,
                },
                attribute_ids: Vec::new(),
                hierarchies: Vec::new(),
            }),
        }
    }

    #[test]
    fn qualifies_same_named_columns_by_containing_table() {
        let metadata = MetadataModel {
            files: vec![
                table("T1", vec![column("T1", "Key")]),
                table("T2", vec![column("T2", "Key")]),
            ],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let projection =
            project_xldm140_identity(&metadata, &olap(vec![dimension("T1"), dimension("T2")]))
                .unwrap();
        assert!(projection.column("T1", "Key").is_some());
        assert!(projection.column("T2", "Key").is_some());
        assert!(projection.column("T3", "Key").is_none());
    }

    #[test]
    fn binds_time_grouping_source_and_calculated_columns_by_qualified_identity() {
        let bytes = [1_u8];
        let storage = test_xldm140_storage(&bytes, &["Partitions"]);
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: None,
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: vec![Xldm140TableIdentity {
                    table_id: "T1-ID".into(),
                    xml_name: "Dates".into(),
                    metadata_path: "Dates.tbl.xml".into(),
                    dimension_object_id: "D1".into(),
                    attribute_ids: vec!["Date".into(), "Date.Year".into()],
                    columns: vec![
                        Xldm140ColumnIdentity {
                            table_id: "T1-ID".into(),
                            column_id: "Date".into(),
                            column_name: None,
                            db_type: Some(7),
                            settings: 0,
                            data_path: "Dates.Date.idf".into(),
                            is_calculated: false,
                        },
                        Xldm140ColumnIdentity {
                            table_id: "T1-ID".into(),
                            column_id: "Date.Year".into(),
                            column_name: None,
                            db_type: Some(20),
                            settings: 2,
                            data_path: "Dates.Date.Year.idf".into(),
                            is_calculated: true,
                        },
                    ],
                }],
                relationships: Vec::new(),
            },
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Outer,
            }],
            unknown_members: Vec::new(),
        };
        let binding = closure
            .bind_time_grouping_inner("Dates", "Date", "Date", &[("Date.Year", "Date.Year", None)])
            .expect("qualified source and calculated identities should bind");
        assert_eq!(binding.table_id, "T1-ID");
        assert_eq!(binding.source.data_path, "Dates.Date.idf");
        assert!(binding.calculated_columns[0].is_calculated);

        let error = closure
            .bind_time_grouping_inner(
                "Dates",
                "Date",
                "wrong-id",
                &[("Date.Year", "Date.Year", None)],
            )
            .unwrap_err();
        assert!(error.to_string().contains("absent from table"));
        let error = closure
            .bind_time_grouping_inner("Dates", "Date.Year", "Date.Year", &[("Date", "Date", None)])
            .unwrap_err();
        assert!(error.to_string().contains("expected source"));
    }

    #[test]
    fn binds_distinct_descriptor_names_to_explicit_attribute_ids() {
        let mut source = column("Dates", "order-date-id");
        source.members.push(crate::metadata::MetadataMember {
            name: "ColumnStats".into(),
            object: Box::new(object(
                "XMColumnStats",
                None,
                vec![("DBType", "7")],
                vec![],
                vec![],
            )),
        });
        let mut calculated = column("Dates", "order-year-id");
        calculated
            .properties
            .push(crate::metadata::MetadataProperty {
                name: "Settings".into(),
                value: "2".into(),
            });
        calculated.members.push(crate::metadata::MetadataMember {
            name: "ColumnStats".into(),
            object: Box::new(object(
                "XMColumnStats",
                None,
                vec![("DBType", "20")],
                vec![],
                vec![],
            )),
        });
        let metadata = MetadataModel {
            files: vec![table("Dates", vec![source, calculated])],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let projection = project_xldm140_identity(
            &metadata,
            &olap(vec![dimension_with_named_attributes(
                "Dates",
                &[
                    ("order-date-id", "Order Date"),
                    ("order-year-id", "Order Year"),
                ],
            )]),
        )
        .expect("the explicit Dimension Attribute mapping is a complete identity fixture");
        let native = NativeModel {
            files: projection
                .tables
                .iter()
                .flat_map(|table| table.columns.iter())
                .map(|column| NativeFile {
                    storage_path: Box::leak(column.data_path.clone().into_boxed_str()),
                    bytes: &[],
                    data: NativeData::Idf(IdfFile {
                        segments: Vec::new(),
                        trailing_zero_padding: &[],
                    }),
                })
                .collect(),
        };
        let generated = SystemGeneratedModel { files: Vec::new() };
        validate_xldm140_identity_closure(&projection, &native, &generated)
            .expect("distinct metadata IDs, native data paths, and OLAP Attributes form a complete identity closure");
        let bytes = [1_u8];
        let storage = test_xldm140_storage(&bytes, &["Partitions"]);
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: None,
            generated: None,
            projection,
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Outer,
            }],
            unknown_members: Vec::new(),
        };
        let binding = closure
            .bind_time_grouping_inner(
                "Dates",
                "Order Date",
                "order-date-id",
                &[(
                    "Order Year",
                    "order-year-id",
                    Some(Xldm140TimeGroupingContentType::Years),
                )],
            )
            .expect("explicit display names, IDs, date source, and granularity should bind");
        assert_eq!(binding.source.column_name, "Order Date");
        assert_eq!(binding.source.column_id, "order-date-id");
        assert_eq!(binding.source.db_type, Some(7));
        assert_eq!(binding.source.settings, 0);
        assert_eq!(binding.calculated_columns[0].column_name, "Order Year");
        assert_eq!(binding.calculated_columns[0].column_id, "order-year-id");
        assert_eq!(binding.calculated_columns[0].db_type, Some(20));
        assert_eq!(binding.calculated_columns[0].settings, 2);
        assert_eq!(
            binding.calculated_columns[0].content_type,
            Some(Xldm140TimeGroupingContentType::Years)
        );
    }

    #[test]
    fn reparses_complete_distinct_id_fixture_for_changed_table_name_and_inverse() {
        let entries = complete_distinct_id_storage_entries();
        let source = build_identity_storage(&entries);
        let storage = crate::inspect(&source).expect("canonical XLDM fixture should inspect");
        let metadata = crate::metadata::inspect(&storage).expect("table metadata should inspect");
        let native = crate::native::inspect(&storage, &metadata.native_parse_options())
            .expect("empty native closure should inspect");
        let generated = crate::generated::inspect_system_generated(&storage)
            .expect("empty generated closure should inspect");
        let olap = crate::olap::inspect(&storage, &metadata).expect("OLAP closure should inspect");
        let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .expect("distinct TableID/XML/Dimension identities should form a complete closure");
        assert!(closure.is_complete());
        let table = closure
            .projection()
            .table("T1")
            .expect("generated TableID should remain addressable");
        assert_eq!(table.xml_name, "OldName");
        assert_eq!(
            table.dimension_object_id,
            "11111111-2222-3333-4444-555555555555"
        );

        let patch = closure
            .rename_table_name("T1", "NewName")
            .expect("a closed metadata-only name edit should be admitted");
        assert!(!patch.is_noop());
        assert_ne!(patch.after(), source.as_slice());
        let changed = patch
            .apply(&source)
            .expect("source-bound patch should apply");
        let changed_storage =
            crate::inspect(changed.as_ref()).expect("changed source should reparse");
        let changed_metadata =
            crate::metadata::inspect(&changed_storage).expect("changed metadata should reparse");
        assert_eq!(
            changed_metadata.files[0].table.name.as_deref(),
            Some("NewName")
        );
        assert_eq!(
            patch.inverse().apply(changed.as_ref()).unwrap(),
            source.as_slice()
        );
    }

    #[test]
    fn relationship_rename_with_limit_accepts_exact_output_and_refuses_one_under() {
        let source = build_identity_storage(&complete_distinct_id_storage_entries());
        let storage = crate::inspect(&source).expect("canonical XLDM fixture should inspect");
        let metadata = crate::metadata::inspect(&storage).expect("table metadata should inspect");
        let native = crate::native::inspect(&storage, &metadata.native_parse_options())
            .expect("native closure should inspect");
        let generated = crate::generated::inspect_system_generated(&storage)
            .expect("generated closure should inspect");
        let olap = crate::olap::inspect(&storage, &metadata).expect("OLAP should inspect");
        let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .expect("fixture should form a complete closure");

        let grown = closure
            .rename_table_name_with_relationships_with_limit(
                "T1",
                "A&B<escaped-name>",
                super::super::model::MAX_STORAGE_BYTES,
            )
            .expect("escaped growth should be admitted");
        assert!(
            grown
                .after()
                .windows(b"A&amp;B&lt;escaped-name>".len())
                .any(|window| window == b"A&amp;B&lt;escaped-name>")
        );
        let exact = grown.after().len();
        assert!(exact > 0);
        let repeated = closure
            .rename_table_name_with_relationships_with_limit("T1", "A&B<escaped-name>", exact)
            .expect("the exact final output cap should be accepted");
        assert_eq!(repeated.after(), grown.after());
        let under = closure
            .rename_table_name_with_relationships_with_limit("T1", "A&B<escaped-name>", exact - 1)
            .expect_err("one byte below the final size must report a caller limit");
        assert_eq!(under.kind(), Xldm140RenameErrorKind::LimitExceeded);
        assert_eq!(under.limit_bounds(), Some((exact, exact - 1)));
        assert_eq!(
            closure
                .rename_table_name_with_relationships_with_limit("T1", "", exact)
                .expect_err("an empty table name is an invalid operation")
                .kind(),
            Xldm140RenameErrorKind::InvalidSource
        );
        assert_eq!(grown.inverse().apply(grown.after()).unwrap(), source);

        let shrunk = closure
            .rename_table_name_with_relationships_with_limit(
                "T1",
                "T",
                super::super::model::MAX_STORAGE_BYTES,
            )
            .expect("a shorter XML name should be admitted");
        assert_ne!(shrunk.after(), source.as_slice());
        assert_eq!(shrunk.inverse().apply(shrunk.after()).unwrap(), source);

        let no_op = closure
            .rename_table_name_with_relationships_with_limit("T1", "OldName", 0)
            .expect("an exact no-op should not require an output allocation");
        assert!(no_op.is_noop());
        assert!(std::ptr::eq(no_op.after().as_ptr(), source.as_ptr()));
    }

    #[test]
    fn physical_complete_fixture_binds_source_and_calculated_time_grouping() {
        let source = build_identity_storage(&complete_distinct_id_storage_entries());
        let storage = crate::inspect(&source).expect("physical XLDM fixture should inspect");
        let metadata = crate::metadata::inspect(&storage).expect("metadata should inspect");
        let native = crate::native::inspect(&storage, &metadata.native_parse_options())
            .expect("native column closure should inspect");
        assert!(native.files.iter().any(|file| {
            file.storage_path.ends_with("1.T1.Key.0.idf") && matches!(file.data, NativeData::Idf(_))
        }));
        assert!(native.files.iter().any(|file| {
            file.storage_path.ends_with("1.T1.Year.0.idf")
                && matches!(file.data, NativeData::Idf(_))
        }));
        let generated = crate::generated::inspect_system_generated(&storage)
            .expect("generated closure should inspect");
        assert!(generated.files.iter().any(|file| {
            file.kind == SystemGeneratedKind::PositionToIdentifier
                && file.storage_path.ends_with("H$T1$Year.POS_TO_ID.0.idf")
        }));
        let olap = crate::olap::inspect(&storage, &metadata).expect("OLAP should inspect");
        assert!(olap.files.iter().any(|file| {
            matches!(
                &file.document,
                OlapDocument::Definition(definition)
                    if definition.kind == OlapObjectKind::Dimension
            )
        }));
        let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .expect("the complete physical closure should be proven");
        let binding = closure
            .bind_time_grouping_with_content_types(
                "OldName",
                "Key",
                "Key",
                &[("Year", "Year", Xldm140TimeGroupingContentType::Years)],
            )
            .expect("source and calculated columns should bind through physical closure");
        assert_eq!(
            binding.source.data_path,
            "Model.1.db/T1.0.dim/1.T1.Key.0.idf"
        );
        assert_eq!(
            binding.calculated_columns[0].data_path,
            "Model.1.db/T1.0.dim/1.T1.Year.0.idf"
        );
        assert!(binding.calculated_columns[0].is_calculated);
        assert_eq!(
            binding.calculated_columns[0].content_type,
            Some(Xldm140TimeGroupingContentType::Years)
        );
    }

    fn column_data_fixture() -> Vec<u8> {
        let mut bytes = Vec::with_capacity(32);
        for value in [0x0102_0304_0506_0708_u64, 0_u64] {
            bytes.extend_from_slice(&1_u64.to_le_bytes());
            bytes.extend_from_slice(&value.to_le_bytes());
        }
        bytes
    }

    fn generated_mapping_fixture() -> Vec<u8> {
        let mut bytes = Vec::with_capacity(16);
        bytes.extend_from_slice(&1_u64.to_le_bytes());
        bytes.extend_from_slice(&0_u64.to_le_bytes());
        bytes
    }

    fn column_metadata_fixture() -> String {
        let mut metadata = r#"<XMObject class="XMSimpleTable" name="OldName"><Properties><Version>1</Version><Settings>0</Settings><RIViolationCount>0</RIViolationCount></Properties><Members><Member><Name>SegmentMap</Name><XMObject class="XMSegment1Map"><Properties><Records>1</Records></Properties></XMObject></Member><Member><Name>TableStats</Name><XMObject class="XMTableStats"><Properties><SegmentSize>1</SegmentSize><Usage>0</Usage></Properties></XMObject></Member></Members><Collections><Collection><Name>Partitions</Name></Collection><Collection><Name>Columns</Name><XMObject class="XMRawColumn" name="Key"><Properties><Settings>0</Settings><ColumnFlags>8</ColumnFlags><Collation></Collation><OrderByColumn></OrderByColumn><Locale>0</Locale><BinaryCharacters>0</BinaryCharacters></Properties><Members><Member><Name>IntrinsicHierarchy</Name><XMObject class="XMHierarchy"><Properties><SortOrder>0</SortOrder><IsProcessed>false</IsProcessed><TypeMaterialization>0</TypeMaterialization><ColumnPosition2DataID>-1</ColumnPosition2DataID><ColumnDataID2Position>-1</ColumnDataID2Position><DistinctDataIDs>1</DistinctDataIDs><TableStore>Key</TableStore></Properties></XMObject></Member><Member><Name>ColumnStats</Name><XMObject class="XMColumnStats"><Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns><Usage>0</Usage><DBType>7</DBType><XMType>0</XMType><CompressionType>0</CompressionType><CompressionParam>0</CompressionParam><EncodingHint>0</EncodingHint><AggCounter>0</AggCounter><WhereCounter>0</WhereCounter><OrderByCounter>0</OrderByCounter></Properties></XMObject></Member></Members><Collections><Collection><Name>Segments</Name><XMObject class="XMColumnSegment"><Properties><Records>1</Records><Mask>0</Mask></Properties><Members><Member><Name>SubSegment</Name><XMObject class="XMColumnSegment"><Properties><Records>1</Records><Mask>0</Mask></Properties><Members><Member><Name>CompressionInfo</Name><XMObject class="XM123CompressionInfo"><Properties><Min>0</Min></Properties></XMObject></Member><Member><Name>ColumnSegmentStats</Name><XMObject class="XMColumnSegmentStats"><Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns></Properties></XMObject></Member></Members></XMObject></Member><Member><Name>CompressionInfo</Name><XMObject class="XMHybridRLECompressionInfo&lt;class XM123CompressionInfo&gt;"><Members><Member><Name>RLECompression</Name><XMObject class="XMRLECompressionInfo"><Properties><BookmarkBits>0</BookmarkBits><StorageAllocSize>0</StorageAllocSize><StorageUsedSize>0</StorageUsedSize><SegmentNeedsResizing>false</SegmentNeedsResizing></Properties></XMObject></Member><Member><Name>SubCompression</Name><XMObject class="XM123CompressionInfo"><Properties><Min>0</Min></Properties></XMObject></Member></Members></XMObject></Member><Member><Name>ColumnSegmentStats</Name><XMObject class="XMColumnSegmentStats"><Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns></Properties></XMObject></Member></Members></XMObject></Collection></Collections><DataObjects><DataObject><XMObject class="XMRawColumnPartitionDataObject" name="1.T1.Key.0.idf"><Properties><DataVersion>0</DataVersion><Partition>0</Partition><SegmentCount>1</SegmentCount></Properties></XMObject></DataObject><DataObject><XMObject class="XMValueDataDictionary&lt;XM_Long&gt;" name="1.T1.Key.dictionary"><Properties><DataVersion>0</DataVersion><BaseId>0</BaseId><Magnitude>0</Magnitude></Properties></XMObject></DataObject></DataObjects></XMObject></Collection><Collection><Name>Relationships</Name></Collection><Collection><Name>UserHierarchies</Name></Collection></Collections></XMObject>"#.to_owned();
        metadata = metadata
            .replace(
                "<IsProcessed>false</IsProcessed><TypeMaterialization>0</TypeMaterialization><ColumnPosition2DataID>-1</ColumnPosition2DataID>",
                "<IsProcessed>true</IsProcessed><TypeMaterialization>0</TypeMaterialization><ColumnPosition2DataID>0</ColumnPosition2DataID>",
            );
        let column_start = metadata.find("<XMObject class=\"XMRawColumn\"").unwrap();
        let column_end = metadata
            .find("</XMObject></Collection><Collection><Name>Relationships")
            .unwrap();
        let calculated = metadata[column_start..column_end + "</XMObject>".len()]
            .replace("name=\"Key\"", "name=\"Year\"")
            .replace("1.T1.Key.0.idf", "1.T1.Year.0.idf")
            .replace("1.T1.Key.dictionary", "1.T1.Year.dictionary")
            .replace("<Settings>0</Settings>", "<Settings>2</Settings>")
            .replace("<DBType>7</DBType>", "<DBType>20</DBType>")
            .replace(
                "<TableStore>Key</TableStore>",
                "<TableStore>Year</TableStore>",
            );
        metadata.insert_str(column_end + "</XMObject>".len(), &calculated);
        metadata
    }

    fn complete_distinct_id_storage_entries() -> Vec<(&'static str, Vec<u8>)> {
        let metadata = column_metadata_fixture().into_bytes();
        let mut entries = vec![
            ("Partitions", partitions_fixture().into_bytes()),
            ("Model.1.db.xml", database_definition_fixture().into_bytes()),
            (
                "Model.1.db/Source.1.ds.xml",
                datasource_definition_fixture().into_bytes(),
            ),
            (
                "Model.1.db/View.1.dsv.xml",
                datasource_view_definition_fixture().into_bytes(),
            ),
            (
                "Model.1.db/C.1.cub.xml",
                cube_definition_fixture().into_bytes(),
            ),
            (
                "Model.1.db/T1.1.dim.xml",
                dimension_definition_fixture().into_bytes(),
            ),
            (
                "Model.1.db/C.0.cub/MdxScript.0.scr.xml",
                mdx_script_definition_fixture().into_bytes(),
            ),
            (
                "Model.1.db/C.0.cub/T1.1.det.xml",
                measure_group_definition_fixture().into_bytes(),
            ),
            (
                "Model.1.db/C.0.cub/T1.0.det/T1.1.prt.xml",
                partition_definition_fixture().into_bytes(),
            ),
            ("Model.1.db/T1.0.dim/T1.1.tbl.xml", metadata),
            ("Model.1.db/T1.0.dim/1.T1.Key.0.idf", column_data_fixture()),
            ("Model.1.db/T1.0.dim/1.T1.Year.0.idf", column_data_fixture()),
            (
                "Model.1.db/T1.0.dim/1.H$T1$Key.POS_TO_ID.0.idf",
                generated_mapping_fixture(),
            ),
            (
                "Model.1.db/T1.0.dim/1.H$T1$Year.POS_TO_ID.0.idf",
                generated_mapping_fixture(),
            ),
        ];
        let log = backup_log_fixture(&entries);
        entries.push(("BackupLog", log.into_bytes()));
        entries
    }

    fn build_identity_storage(entries: &[(&'static str, Vec<u8>)]) -> Vec<u8> {
        let mut bytes = vec![0; super::super::XLDM_PAGE_SIZE];
        bytes.extend_from_slice(&BOM);
        let mut allocations = Vec::new();
        for (index, (_, payload)) in entries.iter().enumerate() {
            if index + 1 == entries.len() {
                bytes.extend_from_slice(&BOM);
            }
            let offset = bytes.len();
            bytes.extend_from_slice(payload);
            bytes.extend_from_slice(&crate::codec::crc32(payload).to_le_bytes());
            allocations.push((offset, payload.len() + CRC_SIZE));
        }
        let directory_offset =
            bytes.len().div_ceil(super::super::XLDM_PAGE_SIZE) * super::super::XLDM_PAGE_SIZE;
        bytes.resize(directory_offset, 0);
        let mut directory = String::from("<VirtualDirectory>");
        for ((path, _), (offset, size)) in entries.iter().zip(&allocations) {
            directory.push_str(&format!(
                "<BackupFile><Path>{path}</Path><Size>{size}</Size><m_cbOffsetHeader>{offset}</m_cbOffsetHeader><Delete>false</Delete><CreatedTimestamp>0</CreatedTimestamp><Access>0</Access><LastWriteTime>0</LastWriteTime></BackupFile>"
            ));
        }
        directory.push_str("</VirtualDirectory>");
        let directory_bytes = crate::codec::utf16le(&directory);
        bytes.extend_from_slice(&directory_bytes);
        bytes.resize(
            bytes.len().div_ceil(super::super::XLDM_PAGE_SIZE) * super::super::XLDM_PAGE_SIZE,
            0,
        );
        let header = format!(
            "<BackupLog><BackupRestoreSyncVersion>140</BackupRestoreSyncVersion><Fault>false</Fault><faultcode>0</faultcode><ErrorCode>true</ErrorCode><EncryptionFlag>false</EncryptionFlag><EncryptionKey>0</EncryptionKey><ApplyCompression>true</ApplyCompression><m_cbOffsetHeader>{directory_offset}</m_cbOffsetHeader><DataSize>{}</DataSize><Files>{}</Files><ObjectID>11111111-2222-3333-4444-555555555500</ObjectID><m_cbOffsetData>4096</m_cbOffsetData></BackupLog>",
            directory_bytes.len(),
            entries.len()
        );
        let mut page = Vec::new();
        page.extend_from_slice(&BOM);
        page.extend_from_slice(&crate::codec::utf16le(super::super::XLDM_STREAM_SIGNATURE));
        page.extend_from_slice(&crate::codec::utf16le(&header));
        page.resize(super::super::XLDM_PAGE_SIZE, 0);
        bytes[..super::super::XLDM_PAGE_SIZE].copy_from_slice(&page);
        bytes
    }

    fn backup_log_fixture(entries: &[(&'static str, Vec<u8>)]) -> String {
        let groups: [(&str, i32, &str, &str, &[&str]); 8] = [
            (
                "100002",
                100_002,
                "DB",
                "11111111-2222-3333-4444-555555555551",
                &["Model.1.db.xml"],
            ),
            (
                "100003",
                100_003,
                "DS",
                "11111111-2222-3333-4444-555555555552",
                &["Model.1.db/Source.1.ds.xml"],
            ),
            (
                "100053",
                100_053,
                "DSV",
                "11111111-2222-3333-4444-555555555553",
                &["Model.1.db/View.1.dsv.xml"],
            ),
            (
                "100010",
                100_010,
                "CUBE",
                "11111111-2222-3333-4444-555555555554",
                &["Model.1.db/C.1.cub.xml"],
            ),
            (
                "100006",
                100_006,
                "DIM",
                "11111111-2222-3333-4444-555555555555",
                &[
                    "Model.1.db/T1.1.dim.xml",
                    "Model.1.db/T1.0.dim/T1.1.tbl.xml",
                    "Model.1.db/T1.0.dim/1.T1.Key.0.idf",
                    "Model.1.db/T1.0.dim/1.T1.Year.0.idf",
                    "Model.1.db/T1.0.dim/1.H$T1$Key.POS_TO_ID.0.idf",
                    "Model.1.db/T1.0.dim/1.H$T1$Year.POS_TO_ID.0.idf",
                ],
            ),
            (
                "100060",
                100_060,
                "SCRIPT",
                "11111111-2222-3333-4444-555555555556",
                &["Model.1.db/C.0.cub/MdxScript.0.scr.xml"],
            ),
            (
                "100016",
                100_016,
                "MG",
                "11111111-2222-3333-4444-555555555557",
                &["Model.1.db/C.0.cub/T1.1.det.xml"],
            ),
            (
                "100021",
                100_021,
                "PART",
                "11111111-2222-3333-4444-555555555558",
                &["Model.1.db/C.0.cub/T1.0.det/T1.1.prt.xml"],
            ),
        ];
        let mut result = String::from(
            "<BackupLog><BackupRestoreSyncVersion>1153</BackupRestoreSyncVersion><ServerRoot>C:\\inert</ServerRoot><SvrEncryptPwdFlag>true</SvrEncryptPwdFlag><ServerEnableBinaryXML>false</ServerEnableBinaryXML><ServerEnableCompression>false</ServerEnableCompression><CompressionFlag>false</CompressionFlag><EncryptionFlag>false</EncryptionFlag><ObjectName>Model</ObjectName><ObjectId>11111111-2222-3333-4444-555555555500</ObjectId><Write>ReadWrite</Write><OlapInfo>true</OlapInfo><Collations><Collation>Latin1_General</Collation></Collations><Languages><Language>1033</Language></Languages><FileGroups>",
        );
        for (_, class, name, object_id, paths) in groups {
            let persist = "Model.1.db";
            let version = if name == "SCRIPT" { 0 } else { 1 };
            let persist_location = 1;
            let object_name = match name {
                "DB" => "Database",
                "DS" => "Source",
                "DSV" => "View",
                "CUBE" => "Cube",
                "DIM" => "T1",
                "SCRIPT" => "MdxScript",
                "MG" | "PART" => "T1",
                _ => name,
            };
            result.push_str(&format!(
                "<FileGroup><Class>{class}</Class><ID>{object_id}</ID><Name>{object_name}</Name><ObjectVersion>{version}</ObjectVersion><PersistLocation>{persist_location}</PersistLocation><PersistLocationPath>{persist}</PersistLocationPath><StorageLocationPath></StorageLocationPath><ObjectID>{object_id}</ObjectID><FileList>"
            ));
            for path in paths {
                let size = entries
                    .iter()
                    .find(|(entry_path, _)| *entry_path == *path)
                    .map_or(0, |(_, payload)| payload.len());
                result.push_str(&format!(
                    "<BackupFile><Path>C:\\inert\\{path}</Path><StoragePath>{path}</StoragePath><LastWriteTime>0</LastWriteTime><Size>{size}</Size></BackupFile>"
                ));
            }
            result.push_str("</FileList></FileGroup>");
        }
        result.push_str("</FileGroups></BackupLog>");
        result
    }

    fn partitions_fixture() -> String {
        "<Partitions><Partition><ObjectPath></ObjectPath><Name></Name><DataSize>0</DataSize><Location></Location><DataSourceID></DataSourceID><ConnectionString></ConnectionString></Partition></Partitions>".into()
    }

    fn olap_definition_fixture(
        kind: &str,
        path_id: &str,
        id: &str,
        parent: &str,
        cube: Option<&str>,
        prefix: &str,
        data_files: &str,
        extras: &str,
    ) -> String {
        let object_id = match id {
            "11111111-2222-3333-4444-555555555551"
            | "11111111-2222-3333-4444-555555555552"
            | "11111111-2222-3333-4444-555555555553"
            | "11111111-2222-3333-4444-555555555554"
            | "11111111-2222-3333-4444-555555555555"
            | "11111111-2222-3333-4444-555555555556"
            | "11111111-2222-3333-4444-555555555557"
            | "11111111-2222-3333-4444-555555555558" => id,
            _ => id,
        };
        let parent = match cube {
            Some(cube) => format!(
                "<ParentObject><DatabaseID>{parent}</DatabaseID><CubeID>{cube}</CubeID></ParentObject>"
            ),
            None if parent.is_empty() => "<ParentObject></ParentObject>".into(),
            None => format!("<ParentObject><DatabaseID>{parent}</DatabaseID></ParentObject>"),
        };
        format!(
            "<Load>{parent}<ObjectDefinition><{kind}><ID>{id}</ID><ObjectID>{object_id}</ObjectID><Name>{path_id}</Name>{prefix}<Ordinal>0</Ordinal><ObjectVersion>1</ObjectVersion><PersistLocation>0</PersistLocation><System>false</System><DataFileList>{data_files}</DataFileList>{extras}</{kind}></ObjectDefinition></Load>"
        )
    }

    fn database_definition_fixture() -> String {
        olap_definition_fixture(
            "Database",
            "Database",
            "11111111-2222-3333-4444-555555555551",
            "",
            None,
            "<DbStorageLocation>Model.1.db</DbStorageLocation>",
            "",
            "",
        )
        .replace(
            "<PersistLocation>0</PersistLocation>",
            "<PersistLocation>1</PersistLocation>",
        )
    }

    fn datasource_definition_fixture() -> String {
        olap_definition_fixture(
            "DataSource",
            "Source",
            "11111111-2222-3333-4444-555555555552",
            "11111111-2222-3333-4444-555555555551",
            None,
            "",
            "",
            "<PermissionFileList></PermissionFileList>",
        )
    }

    fn datasource_view_definition_fixture() -> String {
        olap_definition_fixture(
            "DataSourceView",
            "View",
            "11111111-2222-3333-4444-555555555553",
            "11111111-2222-3333-4444-555555555551",
            None,
            "",
            "",
            "",
        )
    }

    fn cube_definition_fixture() -> String {
        olap_definition_fixture(
            "Cube",
            "Cube",
            "11111111-2222-3333-4444-555555555554",
            "11111111-2222-3333-4444-555555555551",
            None,
            "<Dimensions><Dimension xsi:type=\"CubeDimension\"><DimensionID>11111111-2222-3333-4444-555555555555</DimensionID><Attributes><Attribute xsi:type=\"CubeAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"CubeAttribute\"><AttributeID>Year</AttributeID></Attribute></Attributes></Dimension></Dimensions>",
            "",
            "<PermissionFileList></PermissionFileList><MeasureGroupFileList>T1.1.det.xml</MeasureGroupFileList><PerspectiveFileList></PerspectiveFileList><AssemblyFileList></AssemblyFileList>",
        )
    }

    fn dimension_definition_fixture() -> String {
        olap_definition_fixture(
            "Dimension",
            "T1",
            "11111111-2222-3333-4444-555555555555",
            "11111111-2222-3333-4444-555555555551",
            None,
            "<Attributes><Attribute xsi:type=\"DimensionAttribute\"><ID>Key</ID></Attribute><Attribute xsi:type=\"DimensionAttribute\"><ID>Year</ID></Attribute></Attributes>",
            "1.T1.Key.0.idf;1.T1.Year.0.idf;1.H$T1$Key.POS_TO_ID.0.idf;1.H$T1$Year.POS_TO_ID.0.idf",
            "<PermissionFileList></PermissionFileList>",
        )
    }

    fn mdx_script_definition_fixture() -> String {
        olap_definition_fixture(
            "MdxScript",
            "MdxScript",
            "11111111-2222-3333-4444-555555555556",
            "11111111-2222-3333-4444-555555555551",
            Some("11111111-2222-3333-4444-555555555554"),
            "",
            "",
            "",
        )
        .replace(
            "<ObjectVersion>1</ObjectVersion>",
            "<ObjectVersion>0</ObjectVersion>",
        )
    }

    fn measure_group_definition_fixture() -> String {
        olap_definition_fixture(
            "MeasureGroup",
            "T1",
            "11111111-2222-3333-4444-555555555557",
            "11111111-2222-3333-4444-555555555551",
            Some("11111111-2222-3333-4444-555555555554"),
            "<Dimensions><Dimension xsi:type=\"DegenerateMeasureGroupDimension\"><CubeDimensionID>11111111-2222-3333-4444-555555555555</CubeDimensionID><Attributes><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Year</AttributeID></Attribute></Attributes></Dimension></Dimensions>",
            "",
            "<AggregationDesignFileList></AggregationDesignFileList><PartitionFileList>T1.1.prt.xml</PartitionFileList>",
        )
    }

    fn partition_definition_fixture() -> String {
        olap_definition_fixture(
            "Partition",
            "T1",
            "11111111-2222-3333-4444-555555555558",
            "11111111-2222-3333-4444-555555555551",
            Some("11111111-2222-3333-4444-555555555554"),
            "",
            "",
            "",
        )
    }

    #[test]
    fn projects_calculated_column_from_settings_type_and_modifier_bits() {
        let mut calculated = column("T1", "Date.Year");
        calculated
            .properties
            .push(crate::metadata::MetadataProperty {
                name: "Settings".into(),
                value: "2".into(),
            });
        let mut modified = column("T1", "Date.Month");
        modified.properties.push(crate::metadata::MetadataProperty {
            name: "Settings".into(),
            value: "2048".into(),
        });
        let projection = project_xldm140_identity(
            &MetadataModel {
                files: vec![table("T1", vec![calculated, modified])],
                columns: Vec::new(),
                relationships: Vec::new(),
                hierarchies: Vec::new(),
            },
            &olap(vec![dimension_with_attributes(
                "T1",
                vec!["Date.Year", "Date.Month"],
            )]),
        )
        .expect("valid Settings values should project");
        assert!(projection.column("T1", "Date.Year").unwrap().is_calculated);
        assert!(projection.column("T1", "Date.Month").unwrap().is_calculated);
    }

    #[test]
    fn derives_foreign_side_from_containing_folder() {
        let metadata = MetadataModel {
            files: vec![
                table("T1", vec![column("T1", "Key")]),
                table("T2", vec![column("T2", "Key")]),
                relationship_file("T1", "RelA", "T2"),
            ],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let projection =
            project_xldm140_identity(&metadata, &olap(vec![dimension("T1"), dimension("T2")]))
                .unwrap();
        let relation = &projection.relationships[0];
        assert_eq!(relation.containing_table, "T1");
        assert_eq!(relation.primary_table, "T2");
        assert_eq!(relation.expected_index_key, "R$T1$RelA");
    }

    #[test]
    fn relationship_rename_owns_root_by_dimension_folder_not_xml_name() {
        let mut relation = relationship_file("T1", "RelA", "T2");
        // The root descriptor is a lexical value and is not the ownership
        // key. A valid rename planner must still select this file from its
        // generated T1.0.dim folder; the endpoint's T2 value is independent.
        relation.table.name = Some("T2".into());
        assert!(relationship_metadata_owned_by(&relation, "T1").unwrap());
        assert!(!relationship_metadata_owned_by(&relation, "T2").unwrap());
    }

    #[test]
    fn keeps_primary_xml_name_and_generated_relid_in_separate_namespaces() {
        let mut metadata = MetadataModel {
            files: vec![
                table("T1", vec![column("T1", "Key")]),
                table("T2", vec![column("T2", "Key")]),
                relationship_file("T1", "RelA", "T2"),
            ],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        metadata.files[1].table.name = Some("InnerT2".into());
        let relationship = &mut metadata.files[2].table.collections[0].objects[0];
        relationship.name = Some("DisplayRelationship".into());
        relationship.properties[0].value = "InnerT2".into();

        let projection =
            project_xldm140_identity(&metadata, &olap(vec![dimension("T1"), dimension("T2")]))
                .expect("PrimaryTable uses the XML metadata name");
        let relation = &projection.relationships[0];
        assert_eq!(relation.primary_table, "InnerT2");
        assert_eq!(relation.relationship_id, "RelA");
        assert_eq!(
            relation.relationship_name.as_deref(),
            Some("DisplayRelationship")
        );

        let mut generated_id_reference = metadata;
        generated_id_reference.files[2].table.collections[0].objects[0].properties[0].value =
            "T2".into();
        let error = project_xldm140_identity(
            &generated_id_reference,
            &olap(vec![dimension("T1"), dimension("T2")]),
        )
        .unwrap_err();
        assert!(error.to_string().contains("primary table T2 is absent"));
    }

    #[test]
    fn refuses_duplicate_endpoint_resolution_and_owner_mismatch() {
        let mut metadata = MetadataModel {
            files: vec![
                table("T1", vec![column("T1", "Key")]),
                table("T2", vec![column("T2", "Key")]),
                relationship_file("T1", "RelA", "T2"),
                relationship_file("T1", "RelB", "T2"),
            ],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let error =
            project_xldm140_identity(&metadata, &olap(vec![dimension("T1"), dimension("T2")]))
                .unwrap_err();
        assert!(error.to_string().contains("ambiguous"));

        let mut mismatched = relationship_file("T2", "RelA", "T1");
        mismatched.table.name = Some("T1".into());
        metadata.files[2] = mismatched;
        let projection =
            project_xldm140_identity(&metadata, &olap(vec![dimension("T1"), dimension("T2")]))
                .expect("XML object name is retained separately from containing ownership");
        assert_eq!(projection.relationships[0].containing_table, "T2");
    }

    #[test]
    fn closure_requires_admitted_native_data_and_relationship_index_members() {
        let metadata = MetadataModel {
            files: vec![
                table("T1", vec![column("T1", "Key")]),
                table("T2", vec![column("T2", "Key")]),
                relationship_file("T1", "RelA", "T2"),
            ],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let model = olap(vec![dimension("T1"), dimension("T2")]);
        let projection = project_xldm140_identity(&metadata, &model).unwrap();
        let native = NativeModel { files: Vec::new() };
        let generated = SystemGeneratedModel { files: Vec::new() };
        let error =
            validate_xldm140_identity_closure(&projection, &native, &generated).unwrap_err();
        assert!(error.to_string().contains("native data member"));

        let native = NativeModel {
            files: projection
                .tables
                .iter()
                .flat_map(|table| table.columns.iter())
                .map(|column| NativeFile {
                    storage_path: column.data_path.as_str(),
                    bytes: &[],
                    data: NativeData::Idf(IdfFile {
                        segments: Vec::new(),
                        trailing_zero_padding: &[],
                    }),
                })
                .collect(),
        };
        let error =
            validate_xldm140_identity_closure(&projection, &native, &generated).unwrap_err();
        assert!(error.to_string().contains("relationship"));
    }

    #[test]
    fn rejects_orphan_column_data_and_wrong_table_relationship_index() {
        let metadata = MetadataModel {
            files: vec![
                table("T1", vec![column("T1", "Key")]),
                table("T2", vec![column("T2", "Key")]),
                relationship_file("T1", "RelA", "T2"),
            ],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let model = olap(vec![dimension("T1"), dimension("T2")]);
        let projection = project_xldm140_identity(&metadata, &model).unwrap();
        let mut native_files: Vec<_> = projection
            .tables
            .iter()
            .flat_map(|table| table.columns.iter())
            .map(|column| NativeFile {
                storage_path: column.data_path.as_str(),
                bytes: &[],
                data: NativeData::Idf(IdfFile {
                    segments: Vec::new(),
                    trailing_zero_padding: &[],
                }),
            })
            .collect();
        native_files.push(NativeFile {
            storage_path: "Model.1.db/T1.0.dim/1.T1.Other.0.idf",
            bytes: &[],
            data: NativeData::Idf(IdfFile {
                segments: Vec::new(),
                trailing_zero_padding: &[],
            }),
        });
        let generated = SystemGeneratedModel {
            files: vec![SystemGeneratedFile {
                storage_path: "Model.1.db/T2.0.dim/1.R$T1$RelA.INDEX.0.idf",
                kind: SystemGeneratedKind::RelationshipIndex,
                object_key: "R$T1$RelA".into(),
                version: 1,
                bytes: &[],
                data: SystemGeneratedData::Idf(IdfFile {
                    segments: Vec::new(),
                    trailing_zero_padding: &[],
                }),
            }],
        };
        let error = validate_xldm140_identity_closure(
            &projection,
            &NativeModel {
                files: native_files,
            },
            &generated,
        )
        .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("absent from metadata identity closure")
        );

        let native = NativeModel {
            files: projection
                .tables
                .iter()
                .flat_map(|table| table.columns.iter())
                .map(|column| NativeFile {
                    storage_path: column.data_path.as_str(),
                    bytes: &[],
                    data: NativeData::Idf(IdfFile {
                        segments: Vec::new(),
                        trailing_zero_padding: &[],
                    }),
                })
                .collect(),
        };
        let error =
            validate_xldm140_identity_closure(&projection, &native, &generated).unwrap_err();
        assert!(error.to_string().contains("outside its containing table"));
    }

    #[test]
    fn generated_relationship_index_key_retains_qualified_owner() {
        let metadata_path = "Model.1.db/T1.0.dim/R$T1$RelA.1.tbl.xml";
        let expected = "R$T1$RelA";
        assert!(relationship_generated_object_key_matches(
            "Model.1.db/T1.0.dim/1.R$T1$RelA",
            metadata_path,
            expected,
        ));
        assert!(relationship_generated_object_key_matches(
            "Model.1.db/T1.0.dim/R$T1$RelA",
            metadata_path,
            expected,
        ));
        // Hand-built typed models from the pre-qualified API remain accepted;
        // the physical path validator still binds this legacy form to T1.
        assert!(relationship_generated_object_key_matches(
            expected,
            metadata_path,
            expected,
        ));
        assert!(!relationship_generated_object_key_matches(
            "Model.1.db/T2.0.dim/1.R$T1$RelA",
            metadata_path,
            expected,
        ));
        assert!(!relationship_generated_object_key_matches(
            "Model.1.db/T1.0.dim/x.R$T1$RelA",
            metadata_path,
            expected,
        ));
        assert!(!relationship_generated_object_key_matches(
            "Model.1.db/T1.0.dim/1.R$T1$RelB",
            metadata_path,
            expected,
        ));
    }

    #[test]
    fn closure_member_index_merges_shared_native_generated_paths() {
        let mut paths = HashMap::new();
        paths.insert("native-and-generated", 0);
        paths.insert("unknown", 1);
        let mut sections = vec![None, None];
        record_paths(
            ["native-and-generated"],
            Xldm140MemberSection::Native,
            &paths,
            &mut sections,
        )
        .unwrap();
        record_paths(
            ["native-and-generated"],
            Xldm140MemberSection::Generated,
            &paths,
            &mut sections,
        )
        .unwrap();
        assert_eq!(sections[0], Some(Xldm140MemberSection::NativeAndGenerated));
        assert!(
            record_paths(
                ["missing"],
                Xldm140MemberSection::Metadata,
                &paths,
                &mut sections,
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_filename_only_table_or_column_spoofs() {
        let mut metadata = MetadataModel {
            files: vec![table("T1", vec![column("T1", "Key")])],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let mut spoof = metadata.clone();
        spoof.files[0] = table("T1", vec![column("T2", "Key")]);
        let error = project_xldm140_identity(&spoof, &olap(vec![dimension("T1")])).unwrap_err();
        assert!(error.to_string().contains("column data path"));
        metadata.files[0].table.name = Some("T2".into());
        let projection = project_xldm140_identity(&metadata, &olap(vec![dimension("T1")]))
            .expect("an independent XML table name does not change the generated owner");
        assert_eq!(projection.tables[0].xml_name, "T2");
    }

    #[test]
    fn proves_source_bound_olap_members_and_reports_unknown_directory_members() {
        let database = closure_definition(
            "Model.1.db.xml",
            OlapObjectKind::Database,
            "D",
            OlapParentReference::default(),
        );
        let database_parent = OlapParentReference {
            database_id: Some("D".into()),
            cube_id: None,
        };
        let cube_parent = OlapParentReference {
            database_id: Some("D".into()),
            cube_id: Some("C".into()),
        };
        let olap = OlapModel {
            files: vec![
                database,
                closure_definition(
                    "Model.1.db/Source.1.ds.xml",
                    OlapObjectKind::DataSource,
                    "S",
                    database_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/View.1.dsv.xml",
                    OlapObjectKind::DataSourceView,
                    "V",
                    database_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/Cube.1.cub.xml",
                    OlapObjectKind::Cube,
                    "C",
                    database_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/Table.1.dim.xml",
                    OlapObjectKind::Dimension,
                    "T",
                    database_parent,
                ),
                closure_definition(
                    "Model.1.db/Script.1.scr.xml",
                    OlapObjectKind::MdxScript,
                    "M",
                    cube_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/Group.1.det.xml",
                    OlapObjectKind::MeasureGroup,
                    "G",
                    cube_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/Partition.1.prt.xml",
                    OlapObjectKind::Partition,
                    "P",
                    cube_parent,
                ),
            ],
        };
        let paths = [
            "Model.1.db.xml",
            "Model.1.db/Source.1.ds.xml",
            "Model.1.db/View.1.dsv.xml",
            "Model.1.db/Cube.1.cub.xml",
            "Model.1.db/Table.1.dim.xml",
            "Model.1.db/Script.1.scr.xml",
            "Model.1.db/Group.1.det.xml",
            "Model.1.db/Partition.1.prt.xml",
            "vendor.extra.bin",
        ];
        let bytes = [7_u8];
        let storage = test_xldm140_storage(&bytes, &paths);
        let metadata = MetadataModel {
            files: Vec::new(),
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let native = NativeModel { files: Vec::new() };
        let generated = SystemGeneratedModel { files: Vec::new() };
        let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .expect("synthetic complete OLAP section should be admitted");
        assert_eq!(closure.members().len(), paths.len());
        assert_eq!(closure.unknown_members(), &[8]);
        assert!(!closure.is_complete());
        assert!(std::ptr::eq(
            closure.storage().source_bytes().as_ptr(),
            bytes.as_ptr()
        ));
        assert!(
            closure
                .members()
                .iter()
                .take(8)
                .all(|member| member.section == Xldm140MemberSection::Olap)
        );
        assert_eq!(closure.members()[8].section, Xldm140MemberSection::Other);

        let missing_paths = [
            "Model.1.db.xml",
            "Model.1.db/Source.1.ds.xml",
            "Model.1.db/View.1.dsv.xml",
            "Model.1.db/Cube.1.cub.xml",
            "Model.1.db/Table.1.dim.xml",
            "Model.1.db/Script.1.scr.xml",
            "Model.1.db/Group.1.det.xml",
            "vendor.extra.bin",
        ];
        let missing_storage = test_xldm140_storage(&bytes, &missing_paths);
        let Err(error) =
            prove_xldm140_closure(&missing_storage, &metadata, &olap, &native, &generated)
        else {
            panic!("a typed OLAP member missing from the directory was admitted");
        };
        assert!(
            error
                .to_string()
                .contains("absent from section 2.2 directory")
        );
    }

    #[test]
    fn proves_metadata_native_and_olap_path_closure_together() {
        let database_parent = OlapParentReference {
            database_id: Some("D".into()),
            cube_id: None,
        };
        let cube_parent = OlapParentReference {
            database_id: Some("D".into()),
            cube_id: Some("C".into()),
        };
        let mut dimension = closure_definition(
            "Model.1.db/T1.1.dim.xml",
            OlapObjectKind::Dimension,
            "T1",
            database_parent.clone(),
        );
        if let OlapDocument::Definition(definition) = &mut dimension.document {
            definition.attribute_ids.push("Key".into());
        }
        let olap = OlapModel {
            files: vec![
                closure_definition(
                    "Model.1.db.xml",
                    OlapObjectKind::Database,
                    "D",
                    OlapParentReference::default(),
                ),
                closure_definition(
                    "Model.1.db/Source.1.ds.xml",
                    OlapObjectKind::DataSource,
                    "S",
                    database_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/View.1.dsv.xml",
                    OlapObjectKind::DataSourceView,
                    "V",
                    database_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/Cube.1.cub.xml",
                    OlapObjectKind::Cube,
                    "C",
                    database_parent.clone(),
                ),
                dimension,
                closure_definition(
                    "Model.1.db/Script.1.scr.xml",
                    OlapObjectKind::MdxScript,
                    "M",
                    cube_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/Group.1.det.xml",
                    OlapObjectKind::MeasureGroup,
                    "G",
                    cube_parent.clone(),
                ),
                closure_definition(
                    "Model.1.db/Partition.1.prt.xml",
                    OlapObjectKind::Partition,
                    "P",
                    cube_parent,
                ),
            ],
        };
        let metadata_file = table("T1", vec![column("T1", "Key")]);
        let data_path = "Model.1.db/T1.0.dim/1.T1.Key.0.idf";
        let metadata = MetadataModel {
            files: vec![metadata_file],
            columns: vec![ColumnPolicy {
                name: "Key".into(),
                data_file: data_path.into(),
                segment_count: 0,
                row_count: 0,
                compression_type: 0,
                settings: 0,
                dictionary: None,
            }],
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let native = NativeModel {
            files: vec![NativeFile {
                storage_path: data_path,
                bytes: &[],
                data: NativeData::Idf(IdfFile {
                    segments: Vec::new(),
                    trailing_zero_padding: &[],
                }),
            }],
        };
        let generated = SystemGeneratedModel { files: Vec::new() };
        let paths = [
            "Partitions",
            "Model.1.db.xml",
            "Model.1.db/Source.1.ds.xml",
            "Model.1.db/View.1.dsv.xml",
            "Model.1.db/Cube.1.cub.xml",
            "Model.1.db/T1.1.dim.xml",
            "Model.1.db/Script.1.scr.xml",
            "Model.1.db/Group.1.det.xml",
            "Model.1.db/Partition.1.prt.xml",
            "BackupLog",
            "Model.1.db/T1.0.dim/T1.1.tbl.xml",
            data_path,
            "vendor.extra.bin",
        ];
        let bytes = [9_u8];
        let mut storage = test_xldm140_storage(&bytes, &paths);
        storage.files[0].kind = FileKind::Partitions;
        storage.files[9].kind = FileKind::BackupLog;
        let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .expect("metadata, native, and OLAP members should form one closure");
        assert_eq!(closure.projection().tables.len(), 1);
        assert_eq!(closure.projection().tables[0].columns[0].column_id, "Key");
        assert_eq!(closure.members()[0].section, Xldm140MemberSection::Outer);
        assert_eq!(closure.members()[1].section, Xldm140MemberSection::Olap);
        assert_eq!(closure.members()[9].section, Xldm140MemberSection::Outer);
        assert_eq!(
            closure.members()[10].section,
            Xldm140MemberSection::Metadata
        );
        assert_eq!(closure.members()[11].section, Xldm140MemberSection::Native);
        assert_eq!(closure.unknown_members(), &[12]);
        assert!(!closure.is_complete());

        // Once the source advertises an OLAP BackupLog, the closure must also
        // prove the standalone Dimension relationship graph; the narrow
        // hand-built fixture has no FileGroups and is therefore refused.
        let mut flagged_storage = storage.clone();
        flagged_storage.backup_log.is_olap = true;
        let result = prove_xldm140_closure(&flagged_storage, &metadata, &olap, &native, &generated);
        let error = match result {
            Ok(_) => panic!("OLAP-advertised fixture was admitted without a standalone proof"),
            Err(error) => error,
        };
        assert!(error.to_string().contains("section 2.6 OLAP proof"));
    }

    #[test]
    fn closure_writer_refuses_unknown_and_outer_members_before_rewrite() {
        let bytes = [7_u8];
        let storage = test_xldm140_storage(&bytes, &["inner"]);
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: None,
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: Vec::new(),
                relationships: Vec::new(),
            },
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Other,
            }],
            unknown_members: vec![0],
        };
        let payload = [7_u8];
        let error = closure
            .rewrite_same_size(&[Xldm140FileReplacement {
                storage_path: "inner",
                payload: &payload,
            }])
            .unwrap_err();
        assert!(error.to_string().contains("unknown members"));

        let outer_storage = test_xldm140_storage(&bytes, &["Partitions"]);
        let outer = Xldm140Closure {
            storage: &outer_storage,
            metadata: None,
            native: None,
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: Vec::new(),
                relationships: Vec::new(),
            },
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Outer,
            }],
            unknown_members: Vec::new(),
        };
        let error = outer
            .rewrite_same_size(&[Xldm140FileReplacement {
                storage_path: "Partitions",
                payload: &payload,
            }])
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("outside the writable native/generated closure")
        );
    }

    #[test]
    fn exact_noop_patch_borrows_unknown_source_without_validation_or_copy() {
        let bytes = [3_u8, 4, 5];
        let storage = test_xldm140_storage(&bytes, &["vendor.extra.bin"]);
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: None,
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: Vec::new(),
                relationships: Vec::new(),
            },
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Other,
            }],
            unknown_members: vec![0],
        };
        let patch = closure
            .rewrite_same_size(&[])
            .expect("an exact no-op does not need a complete typed closure");
        assert!(patch.is_noop());
        assert!(std::ptr::eq(patch.before().as_ptr(), bytes.as_ptr()));
        assert!(std::ptr::eq(patch.after().as_ptr(), bytes.as_ptr()));
        let applied = patch.apply(bytes.as_slice()).unwrap();
        assert!(matches!(applied, std::borrow::Cow::Borrowed(_)));
    }

    #[test]
    fn structural_rename_noop_borrows_even_before_full_closure_validation() {
        let bytes = [8_u8, 9, 10];
        let storage = test_xldm140_storage(&bytes, &["vendor.extra.bin"]);
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: None,
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: vec![Xldm140TableIdentity {
                    table_id: "T1".into(),
                    xml_name: "Table".into(),
                    metadata_path: "missing.tbl.xml".into(),
                    dimension_object_id: "D1".into(),
                    attribute_ids: Vec::new(),
                    columns: Vec::new(),
                }],
                relationships: Vec::new(),
            },
            members: Vec::new(),
            unknown_members: vec![0],
        };
        let patch = closure
            .rename_table_name_with_relationships("T1", "Table")
            .expect("an exact structural rename no-op borrows the source");
        assert!(patch.is_noop());
        assert!(std::ptr::eq(patch.after().as_ptr(), bytes.as_ptr()));
    }

    #[test]
    fn fixed_size_rename_refuses_owned_relationship_root_instead_of_leaving_it_stale() {
        let bytes = [8_u8, 9, 10];
        let storage = test_xldm140_storage(&bytes, &["Model.1.db/T1.0.dim/T1.1.tbl.xml"]);
        let metadata_file = relationship_file("T1", "RelA", "Other");
        let metadata = MetadataModel {
            files: vec![metadata_file],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: Some(&metadata),
            native: None,
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: vec![Xldm140TableIdentity {
                    table_id: "T1".into(),
                    xml_name: "T1".into(),
                    metadata_path: "Model.1.db/T1.0.dim/T1.1.tbl.xml".into(),
                    dimension_object_id: "D1".into(),
                    attribute_ids: Vec::new(),
                    columns: Vec::new(),
                }],
                relationships: Vec::new(),
            },
            members: Vec::new(),
            unknown_members: Vec::new(),
        };
        let error = closure.rename_table_name("T1", "T2").unwrap_err();
        assert!(error.to_string().contains("relationship metadata roots"));
    }

    #[test]
    fn reversible_patch_rejects_stale_bytes_and_restores_borrowed_source() {
        let before = [1_u8, 2, 3];
        let after = [3_u8, 2, 1];
        let patch = Xldm140Patch {
            before: &before,
            after: Xldm140PatchBytes::Owned(after.to_vec()),
        };
        assert_eq!(patch.apply(&before).unwrap().as_ref(), after);
        let inverse = patch.inverse();
        assert_eq!(inverse.apply(&after).unwrap(), before);
        assert!(patch.apply(&[1, 2, 4]).is_err());
        assert!(inverse.apply(&[3, 2, 4]).is_err());
    }

    #[test]
    fn table_name_replacement_is_lexical_bounded_and_xml10_checked() {
        let source = br#"<?xml version="1.0"?><XMObject name="T1" class="XMSimpleTable"><Properties/></XMObject>"#;
        let changed = replace_table_name_attribute(source, "T1", "T2").unwrap();
        assert_eq!(changed, br#"<?xml version="1.0"?><XMObject name="T2" class="XMSimpleTable"><Properties/></XMObject>"#);
        let error = replace_table_name_attribute(source, "T1", "T22").unwrap_err();
        assert!(error.to_string().contains("allocation size"));
        let error = replace_table_name_attribute(source, "T1", "T\0").unwrap_err();
        assert!(error.to_string().contains("XML 1.0-forbidden"));
        let error = replace_table_name_attribute(source, "other", "T2").unwrap_err();
        assert!(error.to_string().contains("identity snapshot"));
    }

    #[test]
    fn table_name_replacement_ignores_pre_root_xml_lookalikes() {
        for prefix in [
            b"<!-- <XMObject name=\"Old\"> -->".as_slice(),
            b"<?vendor <XMObject name=\"Old\"> ?>".as_slice(),
            b"<![CDATA[<XMObject name=\"Old\">]]>".as_slice(),
        ] {
            let mut source = prefix.to_vec();
            source.extend_from_slice(
                br#"<XMObject class="XMSimpleTable" name="Old"><Properties/></XMObject>"#,
            );
            let changed = replace_table_name_attribute_variable(&source, "Old", "New")
                .expect("the actual XMObject root should be selected");
            let mut expected = prefix.to_vec();
            expected.extend_from_slice(
                br#"<XMObject class="XMSimpleTable" name="New"><Properties/></XMObject>"#,
            );
            assert_eq!(changed, expected);
        }
    }

    #[test]
    fn variable_table_and_relationship_replacements_preserve_unknown_markup() {
        let table = br#"<XMObject class="XMSimpleTable" name="Old"><Unknown a="1"/></XMObject>"#;
        let changed = replace_table_name_attribute_variable(table, "Old", "Longer&Name")
            .expect("variable table names should be rewritten lexically");
        assert_eq!(
            changed,
            br#"<XMObject class="XMSimpleTable" name="Longer&amp;Name"><Unknown a="1"/></XMObject>"#
        );

        let relationships = br#"<XMObject><Relationships><PrimaryTable> Old </PrimaryTable><Extension><PrimaryTable>Other</PrimaryTable></Extension><PrimaryTable>Old&amp;More</PrimaryTable></Relationships><Unknown/></XMObject>"#;
        let changed = replace_relationship_primary_table(relationships, "Old", "New&Name")
            .expect("matching relationship endpoints should be rewritten");
        assert_eq!(
            changed,
            br#"<XMObject><Relationships><PrimaryTable> New&amp;Name </PrimaryTable><Extension><PrimaryTable>Other</PrimaryTable></Extension><PrimaryTable>Old&amp;More</PrimaryTable></Relationships><Unknown/></XMObject>"#
        );
    }

    /// Change 0765: quick-xml drops a leading UTF-8 byte-order mark before
    /// its first event without counting it, so both scanners convert reader
    /// positions with `reader_origin`. A marked part — with or without a
    /// declaration before its root — is rewritten exactly like the unmarked
    /// part, behind the same mark.
    #[test]
    fn byte_order_marked_sources_are_rewritten_like_unmarked_ones() {
        const MARK: &[u8] = b"\xEF\xBB\xBF";
        for (source, reader_skips) in [
            (&b"<a/>"[..], 0),
            (b"\xEF\xBB\xBF<a/>", 3),
            (b"\xEF\xBB\xBF\xEF\xBB\xBF<a/>", 3),
            (b"\xEF\xBB<a/>", 0),
        ] {
            assert_eq!(reader_origin(source), reader_skips, "{source:?}");
            let mut reader = quick_xml::reader::Reader::from_reader(source);
            let _ = reader.read_event().expect("event");
            let first_event_bytes = usize::try_from(reader.buffer_position()).expect("fits");
            assert!(first_event_bytes + reader_skips <= source.len());
        }
        let declaration = br#"<?xml version="1.0" encoding="UTF-8"?>"#;
        for prefix in [&b""[..], declaration] {
            let relationships = [
                prefix,
                br#"<XMObject><Relationships><PrimaryTable> Old </PrimaryTable><PrimaryTable>Old&amp;More</PrimaryTable></Relationships></XMObject>"#,
            ]
            .concat();
            let marked = [MARK, relationships.as_slice()].concat();
            let plain = replace_relationship_primary_table(&relationships, "Old", "New&Name")
                .expect("unmarked relationships");
            let changed = replace_relationship_primary_table(&marked, "Old", "New&Name")
                .expect("marked relationships");
            assert_eq!(changed, [MARK, plain.as_slice()].concat());

            let table = [
                prefix,
                br#"<XMObject class="XMSimpleTable" name="Old"><Unknown a="1"/></XMObject>"#,
            ]
            .concat();
            let marked = [MARK, table.as_slice()].concat();
            let plain = replace_table_name_attribute_variable(&table, "Old", "Longer&Name")
                .expect("unmarked table");
            let changed = replace_table_name_attribute_variable(&marked, "Old", "Longer&Name")
                .expect("marked table");
            assert_eq!(changed, [MARK, plain.as_slice()].concat());
        }
    }

    #[test]
    fn relationship_replacement_ignores_opaque_same_name_markup() {
        let relationships = br#"<XMObject><Relationships><PrimaryTable>Old</PrimaryTable><Extension><PrimaryTable>Old</PrimaryTable><![CDATA[<PrimaryTable>Old</PrimaryTable>]]><!-- <PrimaryTable>Old</PrimaryTable> --></Extension></Relationships></XMObject>"#;
        let changed = replace_relationship_primary_table(relationships, "Old", "New")
            .expect("opaque relationship descendants must not become rewrite candidates");
        assert_eq!(
            changed,
            br#"<XMObject><Relationships><PrimaryTable>New</PrimaryTable><Extension><PrimaryTable>Old</PrimaryTable><![CDATA[<PrimaryTable>Old</PrimaryTable>]]><!-- <PrimaryTable>Old</PrimaryTable> --></Extension></Relationships></XMObject>"#
        );

        let canonical = br#"<XMObject><Collections><Collection><Name>Relationships</Name><XMObject class="XMRelationship"><Properties><PrimaryTable>Old</PrimaryTable></Properties></XMObject></Collection></Collections></XMObject>"#;
        let changed = replace_relationship_primary_table(canonical, "Old", "New")
            .expect("canonical XMRelationship ownership should be admitted");
        assert_eq!(
            changed,
            br#"<XMObject><Collections><Collection><Name>Relationships</Name><XMObject class="XMRelationship"><Properties><PrimaryTable>New</PrimaryTable></Properties></XMObject></Collection></Collections></XMObject>"#
        );
    }

    #[test]
    fn relationship_replacement_validates_all_scalar_spans_before_output_allocation() {
        let malformed = br#"<XMObject><PrimaryTable><Nested/></PrimaryTable></XMObject>"#;
        let error = replace_relationship_primary_table(malformed, "Old", "New").unwrap_err();
        assert!(error.to_string().contains("not scalar"));

        let unclosed = br#"<XMObject><PrimaryTable>Old</XMObject>"#;
        let error = replace_relationship_primary_table(unclosed, "Old", "New").unwrap_err();
        assert!(error.to_string().contains("unclosed"));
    }

    #[test]
    fn relationship_rename_preserves_comment_lookalikes() {
        for comment in [
            "<!-- <PrimaryTable>Old</PrimaryTable> -->",
            "<!-- <PrimaryTable>Old -->",
        ] {
            let source = format!(
                "<XMObject><Relationships>{comment}<PrimaryTable>Old</PrimaryTable></Relationships></XMObject>"
            );
            let expected = format!(
                "<XMObject><Relationships>{comment}<PrimaryTable>New</PrimaryTable></Relationships></XMObject>"
            );
            let actual = replace_relationship_primary_table(source.as_bytes(), "Old", "New")
                .expect("comment text must not affect relationship matching");
            assert_eq!(std::str::from_utf8(&actual).unwrap(), expected);
        }
    }

    #[test]
    fn aggregate_identity_string_budget_refuses_before_projection_reserve() {
        let metadata = MetadataModel {
            files: vec![MetadataFile {
                storage_path: "Model.1.db/T1.0.dim/T1.1.tbl.xml",
                bytes: &[],
                kind: MetadataFileKind::Table,
                table: MetadataObject {
                    class: MetadataClass("XMSimpleTable".into()),
                    name: Some("x".repeat(MAX_IDENTITY_BYTES + 1)),
                    provider_version: None,
                    properties: Vec::new(),
                    members: Vec::new(),
                    collections: Vec::new(),
                    data_objects: Vec::new(),
                },
            }],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        let error =
            project_xldm140_identity(&metadata, &OlapModel { files: Vec::new() }).unwrap_err();
        assert!(error.to_string().contains("identity string budget"));
    }

    #[test]
    fn native_generated_role_swaps_are_refused_before_any_outer_rewrite() {
        let native = NativeModel {
            files: vec![NativeFile {
                storage_path: "Model.1.db/T1.0.dim/1.T1.Key.0.idf",
                bytes: &[],
                data: NativeData::Dictionary(DictionaryFile {
                    dictionary_type: DictionaryType::Long,
                    hash: None,
                    body: DictionaryBody::Numeric(NumericDictionary {
                        element_count: 0,
                        element_size: 4,
                        values: &[],
                    }),
                    trailing_zero_padding: &[],
                }),
            }],
        };
        let error = validate_native_generated_member_layouts(
            &native,
            &SystemGeneratedModel { files: Vec::new() },
        )
        .unwrap_err();
        assert!(error.to_string().contains("typed layout"));
    }

    #[test]
    fn same_size_valid_native_payload_substitution_is_refused_by_semantic_guard() {
        let path = "Model.1.db/T1.0.dim/1.T1.Key.0.idf";
        let mut source_payload = Vec::new();
        source_payload.extend_from_slice(&1_u64.to_le_bytes());
        source_payload.extend_from_slice(&[0x11; 8]);
        source_payload.extend_from_slice(&1_u64.to_le_bytes());
        source_payload.extend_from_slice(&[0x12; 8]);
        let mut replacement = source_payload.clone();
        replacement[24] ^= 0x01;
        let source_data = crate::native::parse_idf(&source_payload).unwrap();
        let native = NativeModel {
            files: vec![NativeFile {
                storage_path: path,
                bytes: &source_payload,
                data: NativeData::Idf(source_data),
            }],
        };
        let storage_bytes = [0_u8];
        let storage = test_xldm140_storage(&storage_bytes, &[path]);
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: Some(&native),
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: Vec::new(),
                relationships: Vec::new(),
            },
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Native,
            }],
            unknown_members: Vec::new(),
        };
        let error = closure
            .validate_replacement_payload(0, &replacement)
            .unwrap_err();
        assert!(error.to_string().contains("semantic payload"));
    }

    #[test]
    fn same_size_identical_typed_replacement_borrows_source_without_outer_rewrite() {
        let path = "Model.1.db/T1.0.dim/1.T1.Key.0.idf";
        let mut source_payload = Vec::new();
        source_payload.extend_from_slice(&1_u64.to_le_bytes());
        source_payload.extend_from_slice(&[0x31; 8]);
        source_payload.extend_from_slice(&1_u64.to_le_bytes());
        source_payload.extend_from_slice(&[0x32; 8]);
        let source_data = crate::native::parse_idf(&source_payload).unwrap();
        let native = NativeModel {
            files: vec![NativeFile {
                storage_path: path,
                bytes: &source_payload,
                data: NativeData::Idf(source_data),
            }],
        };
        let mut storage_bytes = source_payload.clone();
        storage_bytes.extend_from_slice(&[0_u8; 4]);
        let mut storage = test_xldm140_storage(&storage_bytes, &[path]);
        storage.files[0].offset = super::super::Offset(0);
        storage.files[0].stored_size = super::super::Size(
            u64::try_from(storage_bytes.len()).expect("test storage length fits u64"),
        );
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: Some(&native),
            generated: None,
            projection: Xldm140IdentityProjection {
                tables: Vec::new(),
                relationships: Vec::new(),
            },
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Native,
            }],
            unknown_members: Vec::new(),
        };
        let patch = closure
            .rewrite_same_size(&[Xldm140FileReplacement {
                storage_path: path,
                payload: &source_payload,
            }])
            .expect("an identical typed replacement is an exact no-op");
        assert!(patch.is_noop());
        assert!(std::ptr::eq(
            patch.after().as_ptr(),
            storage.source_bytes().as_ptr()
        ));
    }

    #[test]
    fn same_size_valid_generated_payload_substitution_is_refused_by_semantic_guard() {
        let path = "Model.1.db/T1.0.dim/1.R$T1$Rel.INDEX.0.idf";
        let mut source_payload = Vec::new();
        source_payload.extend_from_slice(&1_u64.to_le_bytes());
        source_payload.extend_from_slice(&[0x22; 8]);
        let mut replacement = source_payload.clone();
        replacement[8] ^= 0x01;
        let source_data = crate::native::parse_idf(&source_payload).unwrap();
        let generated = SystemGeneratedModel {
            files: vec![SystemGeneratedFile {
                storage_path: path,
                kind: SystemGeneratedKind::RelationshipIndex,
                object_key: "R$T1$Rel".into(),
                version: 1,
                bytes: &source_payload,
                data: SystemGeneratedData::Idf(source_data),
            }],
        };
        let storage_bytes = [0_u8];
        let storage = test_xldm140_storage(&storage_bytes, &[path]);
        let closure = Xldm140Closure {
            storage: &storage,
            metadata: None,
            native: None,
            generated: Some(&generated),
            projection: Xldm140IdentityProjection {
                tables: Vec::new(),
                relationships: Vec::new(),
            },
            members: vec![Xldm140ClosureMember {
                storage_index: 0,
                section: Xldm140MemberSection::Generated,
            }],
            unknown_members: Vec::new(),
        };
        let error = closure
            .validate_replacement_payload(0, &replacement)
            .unwrap_err();
        assert!(error.to_string().contains("semantic payload"));
    }

    #[test]
    fn maps_dimension_file_groups_by_object_id_and_database_storage_location() {
        let database_path = "Model.7.db.xml";
        let dimension_path = "Model.7.db/TableToken.9.dim.xml";
        let mut database = closure_definition(
            database_path,
            OlapObjectKind::Database,
            "DATABASE",
            OlapParentReference::default(),
        );
        if let OlapDocument::Definition(definition) = &mut database.document {
            definition.extension.persist_location = 7;
            definition.object.children.push(OlapElement {
                name: "DbStorageLocation".into(),
                attributes: Vec::new(),
                text: "Model.7.db".into(),
                children: Vec::new(),
            });
        }
        let mut dimension = dimension("DIMENSION-OBJECT");
        dimension.extension.object_version = 9;
        let olap = OlapModel {
            files: vec![
                database,
                OlapFile {
                    storage_path: dimension_path,
                    bytes: &[],
                    kind: OlapFileKind::Definition(OlapObjectKind::Dimension),
                    document: OlapDocument::Definition(dimension),
                },
            ],
        };
        let bytes = [0_u8];
        let mut storage = test_xldm140_storage(&bytes, &[dimension_path]);
        storage.backup_log.file_groups.push(FileGroup {
            class: FileGroupClass::Dimension,
            id: "TableToken".into(),
            name: "A display name independent of the generated folder".into(),
            object_version: 9,
            persist_location: 7,
            persist_location_path: "Model.7.db".into(),
            storage_location_path: "ignored".into(),
            object_id: "DIMENSION-OBJECT".into(),
            files: vec![LoggedFile {
                source_path: dimension_path.into(),
                storage_path: dimension_path.into(),
                last_write_timestamp: 0,
                size: 0,
                generated: GeneratedPath {
                    normalized_path: dimension_path.into(),
                    kind: GeneratedNameKind::DataSourceOrDimensionDefinition,
                },
            }],
        });
        let metadata = MetadataModel {
            files: Vec::new(),
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        validate_dimension_file_groups(&storage, &metadata, &olap)
            .expect("dimension groups use the database storage path, not a table folder");

        storage.backup_log.file_groups[0].persist_location_path =
            "Model.7.db/TableToken.9.dim".into();
        let error = validate_dimension_file_groups(&storage, &metadata, &olap).unwrap_err();
        assert!(error.to_string().contains("DbStorageLocation"));
    }
}
