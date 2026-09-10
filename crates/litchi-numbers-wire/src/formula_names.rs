//! Bounded borrowed projections for native Numbers formula names.
//!
//! Formula owner records contain only identity edges needed by a renderer:
//! the owner UUID and a local reference to the selected table-info record.
//! This module validates those edges without decoding an AST or retaining a
//! generated dependency archive.  Category labels use the existing borrowed
//! `GroupNode` projection and are likewise returned without native objects.
//!
//! Owner admission preserves the focused Numbers name-map profile: only the
//! selected UUID and local-reference edges are schema-validated. Other owner
//! fields remain opaque; this is not validation of a complete dependency
//! archive for mutation. Canonical deprecated reference metadata remains
//! accepted under that compatibility profile, while external targets are
//! refused.

use core::fmt;
use std::str;

use litchi_iwa_common::wire::{
    RawWireField, RawWireFields, RawWireLimits, WireDescent, WirePreflight,
    preflight_wire_tree_with_limits,
};
use litchi_iwa_common::{Error, LimitKind, WireLimits};
use litchi_iwa_protos::group_node_category_codec;

const FORMULA_OWNER_UUID_FIELD: u32 = 1;
const FORMULA_OWNER_TABLE_INFO_FIELD: u32 = 11;
const LOCAL_REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const LOCAL_REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const LOCAL_REFERENCE_EXTERNAL_FIELD: u32 = 3;
const MIN_SIGN_EXTENDED_I32: u64 = 0xffff_ffff_8000_0000;

/// Aggregate limits for one formula-owner projection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ReadLimits {
    /// Common bounded wire limits shared by the selected owner and children.
    pub wire: WireLimits,
}

/// Resource totals from a successful formula-owner projection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ReadReport {
    input_bytes: usize,
    fields: usize,
    work: usize,
}

impl ReadReport {
    /// Total bytes in each selected message scanned by the projection.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Total wire fields visited across the owner, UUID, and reference.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Total structural parser work charged by the projection.
    #[must_use]
    pub const fn work(self) -> usize {
        self.work
    }
}

/// Cost observed before a formula-owner projection succeeds or refuses.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AttemptedCost {
    /// Selected-message bytes scanned before refusal.
    pub input_bytes: usize,
    /// Selected wire fields visited before refusal.
    pub fields: usize,
    /// Structural parser work charged before refusal.
    pub work: usize,
}

impl From<ReadReport> for AttemptedCost {
    fn from(report: ReadReport) -> Self {
        Self {
            input_bytes: report.input_bytes,
            fields: report.fields,
            work: report.work,
        }
    }
}

/// Borrow-free identity projection of one `FormulaOwnerDependenciesArchive`.
///
/// The four words use the same order as the focused Numbers owner key:
/// little-endian words of the UUID's lower 64-bit half followed by its upper
/// 64-bit half.  The table-info reference is a local, nonzero object
/// identifier; package owners resolve it in their own selected object index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaOwnerDependencies {
    owner_uuid_words: [u32; 4],
    table_info_ref: u64,
    report: ReadReport,
}

impl FormulaOwnerDependencies {
    /// Return the four native CFUUID-compatible owner words.
    #[must_use]
    pub const fn owner_uuid_words(self) -> [u32; 4] {
        self.owner_uuid_words
    }

    /// Return the owner UUID words under the CFUUID terminology used by the
    /// native formula renderer.
    #[must_use]
    pub const fn cfuuid_words(self) -> [u32; 4] {
        self.owner_uuid_words
    }

    /// Return the selected table-info local reference identifier.
    #[must_use]
    pub const fn table_info_ref(self) -> u64 {
        self.table_info_ref
    }

    /// Return the successful bounded scan totals.
    #[must_use]
    pub const fn report(self) -> ReadReport {
        self.report
    }

    /// Return the successful totals in failure-accounting form.
    #[must_use]
    pub const fn cost(self) -> AttemptedCost {
        AttemptedCost {
            input_bytes: self.report.input_bytes,
            fields: self.report.fields,
            work: self.report.work,
        }
    }
}

/// Failure while reading a formula-owner identity projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaOwnerReadError {
    error: Error,
    attempted: AttemptedCost,
}

impl FormulaOwnerReadError {
    /// Return the underlying common wire failure.
    #[must_use]
    pub const fn error(&self) -> &Error {
        &self.error
    }

    /// Return work observed before refusal.
    #[must_use]
    pub const fn attempted(&self) -> AttemptedCost {
        self.attempted
    }
}

impl fmt::Display for FormulaOwnerReadError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.error.fmt(formatter)
    }
}

impl std::error::Error for FormulaOwnerReadError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.error)
    }
}

/// Read one strict borrowed Numbers formula-owner dependency envelope.
///
/// Only the owner UUID and field-11 local table reference are interpreted.
/// Unknown root/reference fields remain opaque, while known fields retain the
/// focused reader's canonical framing and local-reference checks. No AST or
/// generated dependency tree is decoded.
pub fn read_formula_owner_dependencies(
    source: &[u8],
    limits: ReadLimits,
) -> Result<FormulaOwnerDependencies, FormulaOwnerReadError> {
    let mut decoder = Decoder::new(limits);
    let result = decoder.read_owner(source);
    match result {
        Ok((owner_uuid_words, table_info_ref)) => Ok(FormulaOwnerDependencies {
            owner_uuid_words,
            table_info_ref,
            report: decoder.report,
        }),
        Err(error) => Err(decoder.failure(error)),
    }
}

/// Alias for callers that describe this operation as a preflight.
pub fn preflight_formula_owner_dependencies(
    source: &[u8],
    limits: ReadLimits,
) -> Result<FormulaOwnerDependencies, FormulaOwnerReadError> {
    read_formula_owner_dependencies(source, limits)
}

#[derive(Debug)]
struct Decoder {
    limits: ReadLimits,
    report: ReadReport,
}

impl Decoder {
    const fn new(limits: ReadLimits) -> Self {
        Self {
            limits,
            report: ReadReport {
                input_bytes: 0,
                fields: 0,
                work: 0,
            },
        }
    }

    fn read_owner(&mut self, source: &[u8]) -> Result<([u32; 4], u64), Error> {
        let mut owner_uuid_words = None;
        let mut table_info_ref = None;
        self.scan(source, 0, |decoder, field| {
            match field.number() {
                FORMULA_OWNER_UUID_FIELD => {
                    if owner_uuid_words.is_some() {
                        return Err(duplicate(
                            "FormulaOwnerDependenciesArchive.formula_owner_uid",
                        ));
                    }
                    let payload =
                        length_payload(field, "FormulaOwnerDependenciesArchive.formula_owner_uid")?;
                    owner_uuid_words = Some(decoder.scan_uuid(payload, 1)?);
                },
                FORMULA_OWNER_TABLE_INFO_FIELD => {
                    if table_info_ref.is_some() {
                        return Err(duplicate("FormulaOwnerDependenciesArchive.formula_owner"));
                    }
                    let payload =
                        length_payload(field, "FormulaOwnerDependenciesArchive.formula_owner")?;
                    table_info_ref = Some(decoder.scan_local_reference(payload, 1)?);
                },
                _ => {},
            }
            Ok(())
        })?;
        let owner_uuid_words = owner_uuid_words.ok_or_else(|| {
            invalid("FormulaOwnerDependenciesArchive is missing formula_owner_uid")
        })?;
        let table_info_ref = table_info_ref
            .ok_or_else(|| invalid("FormulaOwnerDependenciesArchive is missing formula_owner"))?;
        Ok((owner_uuid_words, table_info_ref))
    }

    fn scan_uuid(&mut self, source: &[u8], depth: usize) -> Result<[u32; 4], Error> {
        let mut lower = None;
        let mut upper = None;
        self.scan(source, depth, |_, field| {
            if field.wire_type() != 0 || !field.key_is_canonical() || !field.value_is_canonical() {
                return Err(invalid(
                    "FormulaOwnerDependenciesArchive UUID has invalid framing",
                ));
            }
            let value = varint_u64(field, "FormulaOwnerDependenciesArchive UUID")?;
            match field.number() {
                1 if lower.replace(value).is_none() => {},
                2 if upper.replace(value).is_none() => {},
                _ => return Err(invalid("FormulaOwnerDependenciesArchive UUID is malformed")),
            }
            Ok(())
        })?;
        let lower = lower
            .ok_or_else(|| invalid("FormulaOwnerDependenciesArchive UUID is missing lower"))?;
        let upper = upper
            .ok_or_else(|| invalid("FormulaOwnerDependenciesArchive UUID is missing upper"))?;
        Ok([
            lower as u32,
            (lower >> 32) as u32,
            upper as u32,
            (upper >> 32) as u32,
        ])
    }

    fn scan_local_reference(&mut self, source: &[u8], depth: usize) -> Result<u64, Error> {
        let mut identifier = None;
        let mut deprecated_type = None;
        let mut external = None;
        self.scan(source, depth, |_, field| {
            match field.number() {
                LOCAL_REFERENCE_IDENTIFIER_FIELD => {
                    if identifier.is_some() || field.wire_type() != 0 {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference is malformed"));
                    }
                    if !field.key_is_canonical() || !field.value_is_canonical() {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference has invalid framing"));
                    }
                    identifier = Some(varint_u64(field, "FormulaOwnerDependenciesArchive table reference")?);
                },
                LOCAL_REFERENCE_DEPRECATED_TYPE_FIELD => {
                    if deprecated_type.is_some() || field.wire_type() != 0 {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference is malformed"));
                    }
                    if !field.key_is_canonical() || !field.value_is_canonical() {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference has invalid framing"));
                    }
                    let value = varint_u64(field, "FormulaOwnerDependenciesArchive table reference")?;
                    if value > u64::from(i32::MAX as u32) && value < MIN_SIGN_EXTENDED_I32 {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference has an invalid deprecated type"));
                    }
                    deprecated_type = Some(value);
                },
                LOCAL_REFERENCE_EXTERNAL_FIELD => {
                    if external.is_some() || field.wire_type() != 0 {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference is malformed"));
                    }
                    if !field.key_is_canonical() || !field.value_is_canonical() {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference has invalid framing"));
                    }
                    let value = varint_u64(field, "FormulaOwnerDependenciesArchive table reference")?;
                    if value > 1 {
                        return Err(invalid("FormulaOwnerDependenciesArchive table reference has an invalid external flag"));
                    }
                    external = Some(value != 0);
                },
                _ => {},
            }
            Ok(())
        })?;
        let identifier = identifier.ok_or_else(|| {
            invalid("FormulaOwnerDependenciesArchive table reference is missing identifier")
        })?;
        if identifier == 0 || external == Some(true) {
            return Err(invalid(
                "FormulaOwnerDependenciesArchive table reference is not local",
            ));
        }
        Ok(identifier)
    }

    fn scan<'source, F>(
        &mut self,
        source: &'source [u8],
        depth: usize,
        mut visitor: F,
    ) -> Result<(), Error>
    where
        F: FnMut(&mut Self, RawWireField<'source>) -> Result<(), Error>,
    {
        if depth > self.limits.wire.max_nesting() {
            return Err(limit(
                LimitKind::Nesting,
                depth,
                self.limits.wire.max_nesting(),
            ));
        }
        self.charge_input(source.len())?;
        let remaining_fields = self
            .limits
            .wire
            .max_fields()
            .checked_sub(self.report.fields)
            .ok_or_else(|| {
                limit(
                    LimitKind::Fields,
                    self.report.fields,
                    self.limits.wire.max_fields(),
                )
            })?;
        let remaining_work = self
            .limits
            .wire
            .max_rewrite_work()
            .checked_sub(self.report.work)
            .ok_or_else(|| {
                limit(
                    LimitKind::RewriteWork,
                    self.report.work,
                    self.limits.wire.max_rewrite_work(),
                )
            })?;
        if remaining_fields == 0 {
            return Err(limit(
                LimitKind::Fields,
                self.report.fields.saturating_add(1),
                self.limits.wire.max_fields(),
            ));
        }
        if remaining_work == 0 {
            return Err(limit(
                LimitKind::RewriteWork,
                self.report.work.saturating_add(1),
                self.limits.wire.max_rewrite_work(),
            ));
        }
        let raw_limits =
            RawWireLimits::new(source.len().max(1), remaining_fields, 0, remaining_work)?;
        let mut fields = RawWireFields::with_limits(source, raw_limits);
        loop {
            let before_fields = fields.fields();
            let before_work = fields.work();
            let next = fields.next();
            self.charge_fields(fields.fields().saturating_sub(before_fields))?;
            self.charge_work(fields.work().saturating_sub(before_work))?;
            let field = match next? {
                Some(field) => field,
                None => break,
            };
            if field.is_group_start() || field.is_group_end() {
                return Err(invalid(
                    "FormulaOwnerDependenciesArchive uses unsupported protobuf groups",
                ));
            }
            visitor(self, field)?;
        }
        Ok(())
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self
            .report
            .input_bytes
            .checked_add(amount)
            .ok_or_else(|| invalid("formula-owner input-byte count overflow"))?;
        if observed > self.limits.wire.max_input_bytes() {
            return Err(limit(
                LimitKind::InputBytes,
                observed,
                self.limits.wire.max_input_bytes(),
            ));
        }
        self.report.input_bytes = observed;
        Ok(())
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self
            .report
            .fields
            .checked_add(amount)
            .ok_or_else(|| invalid("formula-owner field count overflow"))?;
        if observed > self.limits.wire.max_fields() {
            return Err(limit(
                LimitKind::Fields,
                observed,
                self.limits.wire.max_fields(),
            ));
        }
        self.report.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self
            .report
            .work
            .checked_add(amount)
            .ok_or_else(|| invalid("formula-owner work count overflow"))?;
        if observed > self.limits.wire.max_rewrite_work() {
            return Err(limit(
                LimitKind::RewriteWork,
                observed,
                self.limits.wire.max_rewrite_work(),
            ));
        }
        self.report.work = observed;
        Ok(())
    }

    fn failure(&self, error: Error) -> FormulaOwnerReadError {
        FormulaOwnerReadError {
            error,
            attempted: self.report.into(),
        }
    }
}

fn length_payload<'source>(
    field: RawWireField<'source>,
    name: &str,
) -> Result<&'source [u8], Error> {
    if field.wire_type() != 2 || !field.key_is_canonical() || !field.length_is_canonical() {
        return Err(invalid(format!(
            "protobuf {name} field has invalid wire framing"
        )));
    }
    Ok(field.payload())
}

fn wire_length_payload<'source>(
    field: litchi_iwa_common::wire::WireFieldView<'source>,
    name: &str,
) -> Result<&'source [u8], Error> {
    if field.wire_type() != 2 {
        return Err(invalid(format!(
            "protobuf {name} field has invalid wire framing"
        )));
    }
    field
        .canonical_payload()
        .map_err(|_| invalid(format!("protobuf {name} field has invalid wire framing")))
}

fn wire_varint_canonical(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    name: &str,
) -> Result<u64, Error> {
    if field.wire_type() != 0 {
        return Err(invalid(format!(
            "protobuf {name} field has invalid wire framing"
        )));
    }
    field
        .validate_canonical_key()
        .map_err(|_| invalid(format!("protobuf {name} field has invalid wire framing")))?;
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
        .map_err(|error| invalid(format!("protobuf {name} field has invalid value: {error}")))?;
    if width != field.payload().len() || width != litchi_iwa_common::varint::encoded_len(value) {
        return Err(invalid(format!("protobuf {name} field has invalid value")));
    }
    Ok(value)
}

fn varint_u64(field: RawWireField<'_>, name: &str) -> Result<u64, Error> {
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
        .map_err(|error| invalid(format!("protobuf {name} field has invalid value: {error}")))?;
    if width != field.payload().len() || !field.value_is_canonical() {
        return Err(invalid(format!("protobuf {name} field has invalid value")));
    }
    Ok(value)
}

fn duplicate(name: &str) -> Error {
    invalid(format!("protobuf {name} field is duplicated"))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        kind,
        observed,
        limit: maximum,
    }
}

/// Limits for one bounded native category-name projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CategoryReadLimits {
    /// Common bounded wire limits for the selected category tree.
    pub wire: WireLimits,
    /// Maximum number of `GroupNode` records visited, including the root.
    pub max_nodes: usize,
}

impl Default for CategoryReadLimits {
    fn default() -> Self {
        Self {
            wire: WireLimits::default(),
            max_nodes: WireLimits::MAX_FIELDS,
        }
    }
}

/// Native scalar label retained by a category projection.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CategoryLabel<'source> {
    /// UTF-8 label borrowed directly from the category source.
    String(&'source str),
    /// Native finite or non-finite category number.
    Number(f64),
    /// Native category Boolean.
    Boolean(bool),
    /// Native category date scalar.
    Date(f64),
}

impl CategoryLabel<'_> {
    /// Counts display bytes without allocating, before a reader reserves text.
    ///
    /// Floating-point display can exceed 300 bytes for subnormal values, so a
    /// fixed short reservation is insufficient even for a single scalar.
    #[must_use]
    pub fn formatted_len(self) -> usize {
        struct Counter(usize);
        impl fmt::Write for Counter {
            fn write_str(&mut self, value: &str) -> fmt::Result {
                self.0 = self.0.saturating_add(value.len());
                Ok(())
            }
        }
        let mut counter = Counter(0);
        match fmt::write(&mut counter, format_args!("{self}")) {
            Ok(()) => counter.0,
            Err(_) => usize::MAX,
        }
    }
}

impl fmt::Display for CategoryLabel<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::String(value) => formatter.write_str(value),
            Self::Number(value) | Self::Date(value) => fmt::Display::fmt(value, formatter),
            Self::Boolean(value) => formatter.write_str(if *value { "TRUE" } else { "FALSE" }),
        }
    }
}

/// One category UUID and its borrowed scalar display label.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CategoryName<'source> {
    id: [u64; 2],
    label: CategoryLabel<'source>,
}

impl<'source> CategoryName<'source> {
    /// Return the category UUID as lower and upper 64-bit halves.
    #[must_use]
    pub const fn id(self) -> [u64; 2] {
        self.id
    }

    /// Return the borrowed or copyable native category label.
    #[must_use]
    pub const fn label(self) -> CategoryLabel<'source> {
        self.label
    }
}

/// Successful bounded category-name projection.
#[derive(Debug, PartialEq)]
pub struct CategoryRead<'source> {
    /// Category entries in native group-node traversal order.
    pub entries: Vec<CategoryName<'source>>,
    /// Resource totals from the selected group-node preflight.
    pub report: ReadReport,
}

/// Failure while reading native category names.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CategoryReadError {
    error: Error,
    attempted: AttemptedCost,
}

impl CategoryReadError {
    /// Return the underlying common wire failure.
    #[must_use]
    pub const fn error(&self) -> &Error {
        &self.error
    }

    /// Return work observed before refusal.
    #[must_use]
    pub const fn attempted(&self) -> AttemptedCost {
        self.attempted
    }
}

impl fmt::Display for CategoryReadError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.error.fmt(formatter)
    }
}

impl std::error::Error for CategoryReadError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.error)
    }
}

/// Read bounded native category labels from one `GroupNode` payload.
///
/// Category names borrow source strings and copy only scalar UUID/value data.
/// The generated category codec is used only after the shared wire preflight
/// has validated the recursive topology and scalar wrapper framing.
pub fn read_formula_category_names<'source>(
    source: &'source [u8],
    limits: CategoryReadLimits,
) -> Result<CategoryRead<'source>, CategoryReadError> {
    if limits.max_nodes == 0 || limits.max_nodes > WireLimits::MAX_FIELDS {
        return Err(CategoryReadError {
            error: Error::InvalidLimit {
                field: "formula category nodes",
                value: limits.max_nodes,
                maximum: WireLimits::MAX_FIELDS,
            },
            attempted: AttemptedCost::default(),
        });
    }
    let mut attempted_fields = 0usize;
    let mut attempted_work = source.len();
    let (preflight, nodes) =
        match preflight_categories(source, limits, &mut attempted_fields, &mut attempted_work) {
            Ok(value) => value,
            Err(error) => {
                return Err(CategoryReadError {
                    error,
                    attempted: AttemptedCost {
                        input_bytes: source.len(),
                        fields: attempted_fields,
                        work: attempted_work,
                    },
                });
            },
        };
    let recursion_limit =
        u32::try_from(limits.wire.max_nesting().saturating_add(3)).unwrap_or(u32::MAX);
    let root = group_node_category_codec::decode_group_node(
        source,
        group_node_category_codec::DecodeOptions::new(source.len().max(1), recursion_limit),
    )
    .map_err(|error| CategoryReadError {
        error: invalid(format!("invalid Numbers formula category payload: {error}")),
        attempted: AttemptedCost {
            input_bytes: source.len(),
            fields: attempted_fields,
            work: attempted_work,
        },
    })?;

    let mut entries = Vec::new();
    let mut pending = Vec::new();
    pending.try_reserve(1).map_err(|_| CategoryReadError {
        error: Error::Allocation {
            resource: "Numbers formula category traversal",
            amount: 1,
        },
        attempted: AttemptedCost {
            input_bytes: source.len(),
            fields: attempted_fields,
            work: attempted_work,
        },
    })?;
    let mut visited = 0usize;
    retain_category_name(root, &mut entries, &mut visited, nodes).map_err(|error| {
        CategoryReadError {
            error,
            attempted: AttemptedCost {
                input_bytes: source.len(),
                fields: attempted_fields,
                work: attempted_work,
            },
        }
    })?;
    pending.push(root.children());
    while let Some(children) = pending.last_mut() {
        let Some(child_result) = children.next() else {
            pending.pop();
            continue;
        };
        let child = child_result.map_err(|error| CategoryReadError {
            error: invalid(format!("invalid Numbers formula category child: {error}")),
            attempted: AttemptedCost {
                input_bytes: source.len(),
                fields: attempted_fields,
                work: attempted_work,
            },
        })?;
        pending.try_reserve(1).map_err(|_| CategoryReadError {
            error: Error::Allocation {
                resource: "Numbers formula category traversal",
                amount: pending.len().saturating_add(1),
            },
            attempted: AttemptedCost {
                input_bytes: source.len(),
                fields: attempted_fields,
                work: attempted_work,
            },
        })?;
        retain_category_name(child, &mut entries, &mut visited, nodes).map_err(|error| {
            CategoryReadError {
                error,
                attempted: AttemptedCost {
                    input_bytes: source.len(),
                    fields: attempted_fields,
                    work: attempted_work,
                },
            }
        })?;
        pending.push(child.children());
    }

    Ok(CategoryRead {
        entries,
        report: ReadReport {
            input_bytes: preflight.scanned_bytes(),
            fields: preflight.fields(),
            work: attempted_work,
        },
    })
}

fn preflight_categories(
    source: &[u8],
    limits: CategoryReadLimits,
    attempted_fields: &mut usize,
    attempted_work: &mut usize,
) -> Result<(WirePreflight, usize), Error> {
    if *attempted_work > limits.wire.max_rewrite_work() {
        return Err(limit(
            LimitKind::RewriteWork,
            *attempted_work,
            limits.wire.max_rewrite_work(),
        ));
    }
    let mut nodes = 1usize;
    let report = preflight_wire_tree_with_limits(source, limits.wire, |visit| {
        *attempted_fields = attempted_fields
            .checked_add(1)
            .ok_or_else(|| invalid("formula category field count overflow"))?;
        if *attempted_fields > limits.wire.max_fields() {
            return Err(limit(
                LimitKind::Fields,
                *attempted_fields,
                limits.wire.max_fields(),
            ));
        }
        *attempted_work = attempted_work
            .checked_add(1)
            .ok_or_else(|| invalid("formula category work count overflow"))?;
        if *attempted_work > limits.wire.max_rewrite_work() {
            return Err(limit(
                LimitKind::RewriteWork,
                *attempted_work,
                limits.wire.max_rewrite_work(),
            ));
        }
        let field = visit.field();
        if !visit.path().iter().all(|number| *number == 3) {
            return Err(invalid("formula category topology left the child path"));
        }
        match field.number() {
            1 => {
                let payload = wire_length_payload(field, "formula category UUID")?;
                validate_category_uuid(payload)?;
                Ok(WireDescent::Skip)
            },
            3 => {
                let _ = wire_length_payload(field, "formula category child")?;
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit(LimitKind::Fields, usize::MAX, limits.max_nodes))?;
                if nodes > limits.max_nodes {
                    return Err(limit(LimitKind::Fields, nodes, limits.max_nodes));
                }
                Ok(WireDescent::Descend)
            },
            7 => {
                let payload = wire_length_payload(field, "formula category value")?;
                validate_category_value(payload)?;
                Ok(WireDescent::Skip)
            },
            _ => Ok(WireDescent::Skip),
        }
    })?;
    Ok((report, nodes))
}

fn validate_category_uuid(source: &[u8]) -> Result<(), Error> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .and_then(|limits| limits.with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_nesting(1))?;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        if field.number() == 1 || field.number() == 2 {
            if field.wire_type() != 0 {
                return Err(invalid("formula category UUID has the wrong wire type"));
            }
            field.validate_canonical_key()?;
            let _ = wire_varint_canonical(field, "formula category UUID")?;
        }
        Ok(WireDescent::Skip)
    })?;
    Ok(())
}

fn validate_category_value(source: &[u8]) -> Result<(), Error> {
    let limits = WireLimits::default()
        .with_input_bytes(
            source
                .len()
                .saturating_mul(2)
                .clamp(1, WireLimits::MAX_INPUT_BYTES),
        )
        .and_then(|limits| limits.with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_nesting(2))?;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        match (visit.path(), field.number()) {
            ([], 2..=5) => {
                if field.wire_type() != 2 {
                    return Err(invalid("formula category value has the wrong wire type"));
                }
                field.validate_canonical_framing()?;
                Ok(WireDescent::Descend)
            },
            ([2], 1) => {
                let value = wire_varint_canonical(field, "formula category Boolean")?;
                if value > 1 {
                    return Err(invalid("formula category Boolean has a noncanonical value"));
                }
                Ok(WireDescent::Skip)
            },
            ([3 | 4], 1) => {
                if field.wire_type() != 1 {
                    return Err(invalid("formula category scalar has the wrong wire type"));
                }
                field.validate_canonical_key()?;
                Ok(WireDescent::Skip)
            },
            ([5], 1) => {
                let payload = wire_length_payload(field, "formula category string")?;
                str::from_utf8(payload)
                    .map_err(|_| invalid("formula category string is not UTF-8"))?;
                Ok(WireDescent::Skip)
            },
            _ => Ok(WireDescent::Skip),
        }
    })?;
    Ok(())
}

fn retain_category_name<'source>(
    node: group_node_category_codec::GroupNodeView<'source>,
    entries: &mut Vec<CategoryName<'source>>,
    visited: &mut usize,
    expected_nodes: usize,
) -> Result<(), Error> {
    *visited = visited
        .checked_add(1)
        .ok_or_else(|| limit(LimitKind::Fields, usize::MAX, expected_nodes))?;
    if *visited > expected_nodes {
        return Err(invalid(
            "formula category projection exceeded its preflight",
        ));
    }
    let id = node
        .group_uid()
        .map_err(|error| invalid(format!("invalid formula category UUID: {error}")))?
        .map_or([0, 0], |uid| [uid.lower(), uid.upper()]);
    let Some(value) = node
        .category_value()
        .map_err(|error| invalid(format!("invalid formula category value: {error}")))?
    else {
        return Ok(());
    };
    let label = if let Some(value) = value
        .string()
        .map_err(|error| invalid(format!("invalid formula category string: {error}")))?
    {
        CategoryLabel::String(value)
    } else if let Some(value) = value
        .number()
        .map_err(|error| invalid(format!("invalid formula category number: {error}")))?
    {
        CategoryLabel::Number(value)
    } else if let Some(value) = value
        .boolean()
        .map_err(|error| invalid(format!("invalid formula category Boolean: {error}")))?
    {
        CategoryLabel::Boolean(value)
    } else if let Some(value) = value
        .date()
        .map_err(|error| invalid(format!("invalid formula category date: {error}")))?
    {
        CategoryLabel::Date(value)
    } else {
        return Ok(());
    };
    entries.try_reserve(1).map_err(|_| Error::Allocation {
        resource: "Numbers formula category names",
        amount: entries.len().saturating_add(1),
    })?;
    entries.push(CategoryName { id, label });
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn category_display_counts_extreme_scalars_before_allocation() {
        for label in [
            CategoryLabel::String("Café 北京"),
            CategoryLabel::Boolean(true),
            CategoryLabel::Boolean(false),
            CategoryLabel::Number(f64::MAX),
            CategoryLabel::Number(-f64::from_bits(1)),
            CategoryLabel::Date(f64::from_bits(1)),
            CategoryLabel::Number(f64::NAN),
        ] {
            assert_eq!(label.formatted_len(), label.to_string().len());
        }
        assert!(CategoryLabel::Number(f64::from_bits(1)).formatted_len() > 300);
        assert_eq!(CategoryLabel::Boolean(true).to_string(), "TRUE");
        assert_eq!(CategoryLabel::Boolean(false).to_string(), "FALSE");
    }

    fn push_varint(mut value: u64, output: &mut Vec<u8>) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(u64::from(number) << 3, &mut output);
        push_varint(value, &mut output);
        output
    }

    fn length_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint((u64::from(number) << 3) | 2, &mut output);
        push_varint(payload.len() as u64, &mut output);
        output.extend_from_slice(payload);
        output
    }

    fn owner(lower: u64, upper: u64, table: u64) -> Vec<u8> {
        let mut uuid = varint_field(1, lower);
        uuid.extend(varint_field(2, upper));
        let reference = varint_field(1, table);
        let mut output = length_field(1, &uuid);
        output.extend(length_field(11, &reference));
        output
    }

    #[test]
    fn owner_projection_returns_uuid_words_and_local_reference() {
        let source = owner(0x1122_3344_5566_7788, 0x99aa_bbcc_ddee_ff00, 42);
        let projection = read_formula_owner_dependencies(&source, ReadLimits::default()).unwrap();
        assert_eq!(
            projection.owner_uuid_words(),
            [0x5566_7788, 0x1122_3344, 0xddee_ff00, 0x99aa_bbcc]
        );
        assert_eq!(projection.cfuuid_words(), projection.owner_uuid_words());
        assert_eq!(projection.table_info_ref(), 42);
        assert!(projection.report().input_bytes() >= source.len());
        assert!(projection.report().fields() >= 4);
        assert!(projection.report().work() > 0);
    }

    #[test]
    fn owner_projection_rejects_missing_or_foreign_reference() {
        let mut missing = owner(1, 2, 42);
        missing.truncate(missing.len() - 2);
        let failure = read_formula_owner_dependencies(&missing, ReadLimits::default()).unwrap_err();
        assert!(matches!(failure.error(), Error::InvalidFormat(_)));
        assert!(failure.attempted().work > 0);

        let mut foreign = varint_field(1, 42);
        foreign.extend(varint_field(3, 1));
        let mut owner_uuid = varint_field(1, 1);
        owner_uuid.extend(varint_field(2, 2));
        let mut source = length_field(1, &owner_uuid);
        source.extend(length_field(11, &foreign));
        let failure = read_formula_owner_dependencies(&source, ReadLimits::default()).unwrap_err();
        assert!(failure.to_string().contains("not local"));
    }

    #[test]
    fn owner_projection_rejects_noncanonical_known_fields() {
        let mut uuid = vec![0x08, 0x81, 0x00, 0x10, 0x02];
        let reference = varint_field(1, 42);
        let mut source = length_field(1, &uuid);
        source.extend(length_field(11, &reference));
        let failure = read_formula_owner_dependencies(&source, ReadLimits::default()).unwrap_err();
        assert!(matches!(failure.error(), Error::InvalidFormat(_)));

        uuid = varint_field(1, 1);
        let mut source = vec![0x0a, 0x80, 0x00];
        source.extend_from_slice(&uuid);
        source.extend(length_field(11, &reference));
        let failure = read_formula_owner_dependencies(&source, ReadLimits::default()).unwrap_err();
        assert!(matches!(failure.error(), Error::InvalidFormat(_)));
    }

    #[test]
    fn owner_projection_honors_aggregate_limits_and_reports_attempted_cost() {
        let source = owner(1, 2, 42);
        let limits = ReadLimits {
            wire: WireLimits::default().with_fields(3).unwrap(),
        };
        let failure = read_formula_owner_dependencies(&source, limits).unwrap_err();
        assert!(matches!(
            failure.error(),
            Error::LimitExceeded {
                kind: LimitKind::Fields,
                ..
            }
        ));
        assert!(failure.attempted().fields > 0);
    }

    #[test]
    fn category_projection_borrows_native_string_labels() {
        let uuid = {
            let mut bytes = varint_field(1, 7);
            bytes.extend(varint_field(2, 8));
            bytes
        };
        let string_wrapper = length_field(1, b"Cities");
        let cell_value = length_field(5, &string_wrapper);
        let mut source = length_field(1, &uuid);
        source.extend(length_field(7, &cell_value));
        let read = read_formula_category_names(&source, CategoryReadLimits::default()).unwrap();
        assert_eq!(read.entries.len(), 1);
        assert_eq!(read.entries[0].id(), [7, 8]);
        assert_eq!(read.entries[0].label(), CategoryLabel::String("Cities"));
    }

    #[test]
    fn category_projection_rejects_invalid_utf8_and_depth_limits() {
        let string_wrapper = length_field(1, &[0xff]);
        let cell_value = length_field(5, &string_wrapper);
        let source = length_field(7, &cell_value);
        let failure =
            read_formula_category_names(&source, CategoryReadLimits::default()).unwrap_err();
        assert!(failure.to_string().contains("UTF-8"));

        let child = length_field(3, &[]);
        let mut deep = child.clone();
        for _ in 0..3 {
            deep = length_field(3, &deep);
        }
        let limits = CategoryReadLimits {
            wire: WireLimits::default().with_nesting(2).unwrap(),
            max_nodes: 16,
        };
        let failure = read_formula_category_names(&deep, limits).unwrap_err();
        assert!(matches!(
            failure.error(),
            Error::LimitExceeded {
                kind: LimitKind::Nesting,
                ..
            }
        ));
    }
}
