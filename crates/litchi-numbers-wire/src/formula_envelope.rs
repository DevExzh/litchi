//! Bounded, generated-free preflight for Numbers `FormulaArchive` envelopes.
//!
//! This module validates the complete formula wire tree without retaining its
//! repeated AST representation.  The concrete Numbers adapters own byte
//! retention, aggregate budget mutation, decoding, and format error mapping;
//! this crate only reports successful work and the cost attempted by a failed
//! scan.

use litchi_iwa_common::wire::{WireDescent, WirePreflight, preflight_wire_tree_with_limits};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::numbers_formula_codec;

/// Residual limits and already-spent aggregate costs for one formula scan.
///
/// The adapter supplies the residual values from its own format budget.  The
/// preflight never mutates that budget; a successful report is charged by the
/// adapter, while a failed scan exposes [`FormulaEnvelopeFailure::attempted`]
/// for monotonic failure accounting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaEnvelopeLimits {
    /// Aggregate field ceiling, including `base_fields`.
    pub max_fields: usize,
    /// Per-scanner input-byte ceiling.
    pub max_input_bytes: usize,
    /// Aggregate wire-work ceiling, including `base_work`.
    pub max_work: usize,
    /// Fields already charged by the caller before this scan.
    pub base_fields: usize,
    /// Wire work already charged by the caller before this scan.
    pub base_work: usize,
}

impl FormulaEnvelopeLimits {
    fn validate(self) -> litchi_iwa_common::Result<()> {
        if self.max_fields > WireLimits::MAX_FIELDS {
            return Err(litchi_iwa_common::Error::InvalidLimit {
                field: "formula envelope fields",
                value: self.max_fields,
                maximum: WireLimits::MAX_FIELDS,
            });
        }
        if self.max_input_bytes > WireLimits::MAX_INPUT_BYTES {
            return Err(litchi_iwa_common::Error::InvalidLimit {
                field: "formula envelope input bytes",
                value: self.max_input_bytes,
                maximum: WireLimits::MAX_INPUT_BYTES,
            });
        }
        if self.max_work > WireLimits::MAX_REWRITE_WORK {
            return Err(litchi_iwa_common::Error::InvalidLimit {
                field: "formula envelope work",
                value: self.max_work,
                maximum: WireLimits::MAX_REWRITE_WORK,
            });
        }
        if self.base_fields > self.max_fields {
            return Err(litchi_iwa_common::Error::InvalidLimit {
                field: "formula envelope base fields",
                value: self.base_fields,
                maximum: self.max_fields,
            });
        }
        if self.base_work > self.max_work {
            return Err(litchi_iwa_common::Error::InvalidLimit {
                field: "formula envelope base work",
                value: self.base_work,
                maximum: self.max_work,
            });
        }
        Ok(())
    }
}

/// Cost observed while admitting or rejecting one formula envelope.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AttemptedFormulaEnvelopeCost {
    fields: usize,
    work: usize,
}

impl AttemptedFormulaEnvelopeCost {
    /// Number of fields charged before the scan stopped.
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Number of bytes charged before the scan stopped.
    pub const fn work(self) -> usize {
        self.work
    }
}

/// Result of a complete formula-envelope preflight.
///
/// `root_ast_present` is reported separately because the focused and host
/// adapters intentionally map a missing root to different format errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaEnvelopeReport {
    wire: Option<WirePreflight>,
    fields: usize,
    scanned_bytes: usize,
    root_ast_present: bool,
    scalar_visitor_eligible: bool,
    scalar_visitor_node_count: usize,
    lazy_traversal_entry_count: usize,
}

impl FormulaEnvelopeReport {
    /// Number of fields visited by the complete wire scan.
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Number of bytes charged by the selected wire tree scan.
    ///
    /// The bounded required-field checks rescan individual child payloads but
    /// remain validation work rather than additional aggregate admission cost.
    pub const fn scanned_bytes(self) -> usize {
        self.scanned_bytes
    }

    /// Whether the root archive contained its AST-node array.
    pub const fn root_ast_present(self) -> bool {
        self.root_ast_present
    }

    /// Whether all visited AST nodes are admitted by the scalar renderer.
    pub const fn scalar_visitor_eligible(self) -> bool {
        self.scalar_visitor_eligible
    }

    /// Exact count of AST nodes observed by the preflight.
    pub const fn scalar_visitor_node_count(self) -> usize {
        self.scalar_visitor_node_count
    }

    /// Count of non-AST repeated entries traversed by compatibility rendering.
    pub const fn lazy_traversal_entry_count(self) -> usize {
        self.lazy_traversal_entry_count
    }

    /// Return the report's successful aggregate cost.
    pub const fn cost(self) -> AttemptedFormulaEnvelopeCost {
        AttemptedFormulaEnvelopeCost {
            fields: self.fields,
            work: self.scanned_bytes,
        }
    }

    /// Return the common wire report used by aggregate adapter budgets.
    ///
    /// Empty input is the historical default archive and has no wire scan,
    /// so it returns `None` and carries zero aggregate cost.
    pub const fn wire_preflight(self) -> Option<WirePreflight> {
        self.wire
    }
}

/// Error returned after a formula preflight has spent part of its bounded
/// field/work allowance.
#[derive(Debug, PartialEq, Eq)]
pub struct FormulaEnvelopeFailure {
    error: litchi_iwa_common::Error,
    attempted: AttemptedFormulaEnvelopeCost,
}

impl FormulaEnvelopeFailure {
    /// Construct a failure with its attempted aggregate cost.
    const fn new(error: litchi_iwa_common::Error, attempted: AttemptedFormulaEnvelopeCost) -> Self {
        Self { error, attempted }
    }

    /// Borrow the format-neutral wire error for adapter-specific mapping.
    pub const fn error(&self) -> &litchi_iwa_common::Error {
        &self.error
    }

    /// Return the cost spent before the scan stopped.
    pub const fn attempted(&self) -> AttemptedFormulaEnvelopeCost {
        self.attempted
    }

    /// Split the failure into its wire error and attempted cost.
    pub fn into_parts(self) -> (litchi_iwa_common::Error, AttemptedFormulaEnvelopeCost) {
        (self.error, self.attempted)
    }
}

impl std::fmt::Display for FormulaEnvelopeFailure {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "{} (after {} fields and {} bytes)",
            self.error, self.attempted.fields, self.attempted.work
        )
    }
}

impl std::error::Error for FormulaEnvelopeFailure {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.error)
    }
}

/// Strictly preflight one serialized `TSCE.FormulaArchive` envelope.
///
/// The scan retains no decoded message tree or per-field span collection;
/// bounded traversal-path storage and diagnostic strings may still allocate.
/// Empty input is accepted as the historical default archive and reports no
/// root AST. A non-empty input is returned successfully only after every known
/// field has passed its wire, UTF-8, scalar, required-field, duplicate, and
/// canonical-framing checks. Unknown fields remain opaque and make the scalar
/// admission bit conservative. A successful report may have no root AST; the
/// adapter decides whether that format-specific case is admissible.
///
/// On failure, [`FormulaEnvelopeFailure::attempted`] includes source bytes and
/// every field/nested payload charge reached before the failing operation.
pub fn preflight_formula_envelope(
    source: &[u8],
    envelope_limits: FormulaEnvelopeLimits,
) -> Result<FormulaEnvelopeReport, FormulaEnvelopeFailure> {
    if let Err(error) = envelope_limits.validate() {
        return Err(FormulaEnvelopeFailure::new(
            error,
            AttemptedFormulaEnvelopeCost::default(),
        ));
    }
    // Prost accepts an empty proto2 message even when its schema marks field
    // 1 as required.  The former eager FormulaArchive decoder therefore
    // admitted the serialized default archive, whose empty AST rendered as
    // `=`.  Preserve that compatibility case while retaining the required
    // root field check for every non-empty archive.
    if source.is_empty() {
        return Ok(FormulaEnvelopeReport {
            wire: None,
            fields: 0,
            scanned_bytes: 0,
            root_ast_present: false,
            scalar_visitor_eligible: false,
            scalar_visitor_node_count: 0,
            lazy_traversal_entry_count: 0,
        });
    }
    if envelope_limits.max_input_bytes == 0 {
        return Err(FormulaEnvelopeFailure::new(
            litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed: source.len(),
                limit: 0,
            },
            AttemptedFormulaEnvelopeCost {
                fields: 0,
                work: source.len(),
            },
        ));
    }

    // The wire preflight report is aggregate: every selected child message is
    // charged in addition to its parent.  A per-message input/field ceiling
    // would reject an otherwise small, valid formula as soon as the first
    // ASTNode is descended.  Keep the scanner finite with the same package
    // ceilings used by `FormulaEnvelopeCost`; the latter folds the caller's
    // current aggregate offsets into the admission decision.
    let max_fields = envelope_limits
        .max_fields
        .saturating_sub(envelope_limits.base_fields)
        .clamp(1, WireLimits::MAX_FIELDS);
    let max_input_bytes = envelope_limits
        .max_input_bytes
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let scan_limits = WireLimits::default()
        .with_input_bytes(max_input_bytes)
        .and_then(|limits| limits.with_fields(max_fields))
        .and_then(|limits| limits.with_nesting(WireLimits::MAX_NESTING))
        .and_then(|limits| {
            limits.with_rewrite_work(
                envelope_limits
                    .max_work
                    .clamp(1, WireLimits::MAX_REWRITE_WORK),
            )
        })
        .map_err(|error| {
            FormulaEnvelopeFailure::new(error, AttemptedFormulaEnvelopeCost::default())
        })?;

    let mut attempted = match FormulaEnvelopeCost::new(
        source.len(),
        envelope_limits.base_fields,
        envelope_limits.base_work,
        envelope_limits.max_work,
    ) {
        Ok(cost) => cost,
        Err(error) => {
            return Err(FormulaEnvelopeFailure::new(
                error,
                AttemptedFormulaEnvelopeCost {
                    fields: 0,
                    work: source.len(),
                },
            ));
        },
    };
    let mut root_ast_present = false;
    let mut root_ast_count = 0usize;
    let mut scalar_visitor_eligible = true;
    let mut scalar_visitor_node_count = 0usize;
    let mut lazy_traversal_entry_count = 0usize;
    let mut root_known_fields = [0u32; 9];
    let mut root_known_field_count = 0usize;
    let preflight = preflight_wire_tree_with_limits(source, scan_limits, |visit| {
        let field = visit.field();
        attempted.charge_field(envelope_limits.max_fields)?;
        field.validate_canonical_framing()?;
        let schema = formula_envelope_field(visit.path(), field.number());
        // Unknown fields are intentionally retained as opaque wire data by
        // the envelope preflight, but they are not represented by the compact
        // scalar visitor.  Keep these archives on the full compatibility
        // renderer so a future extension cannot silently disappear on the
        // generated-free route.  The strict wire walk and bounded fallback
        // still apply exactly as before.
        if schema.wire_type.is_none() {
            scalar_visitor_eligible = false;
        }
        if let Some(expected_wire_type) = schema.wire_type
            && field.wire_type() != expected_wire_type
        {
            return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                "Numbers FormulaArchive field {} has wire type {}, expected {}",
                field.number(),
                field.wire_type(),
                expected_wire_type
            )));
        }
        if visit.path().is_empty()
            && schema.wire_type.is_some()
            && !formula_field_is_repeated(visit.path(), field.number())
        {
            if root_known_fields[..root_known_field_count].contains(&field.number()) {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} occurs more than once",
                    field.number()
                )));
            }
            root_known_fields[root_known_field_count] = field.number();
            root_known_field_count += 1;
        }
        if visit.path().is_empty() && field.number() == 1 {
            root_ast_count = root_ast_count.saturating_add(1);
            if root_ast_count > 1 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive AST node array occurs more than once".to_owned(),
                ));
            }
            root_ast_present = true;
        }
        if schema.nested {
            attempted.charge_nested(field.payload().len(), envelope_limits.max_work)?;
            let mut nested_path = [0u32; WireLimits::MAX_NESTING + 1];
            let Some(path_end) = visit.path().len().checked_add(1) else {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive nesting path overflows".to_owned(),
                ));
            };
            if path_end > nested_path.len() {
                return Err(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::Nesting,
                    observed: path_end,
                    limit: WireLimits::MAX_NESTING,
                });
            }
            nested_path[..visit.path().len()].copy_from_slice(visit.path());
            nested_path[visit.path().len()] = field.number();
            require_formula_fields(
                field.payload(),
                &nested_path[..path_end],
                schema.required_fields,
            )?;
        }
        if formula_ast_array_path(visit.path()) && field.number() == 1 {
            required_formula_ast_node_type(field.payload())?;
        }
        // Preserve the legacy admission charge for every repeated native
        // vector. AST node entries are charged separately by the formula-render
        // work budget, so count only the other entries (UID lists, ranges,
        // lambda identifiers, and similar vectors) before lazy compatibility
        // traversal begins.
        if formula_field_is_repeated(visit.path(), field.number())
            && !(formula_ast_array_path(visit.path()) && field.number() == 1)
        {
            lazy_traversal_entry_count =
                lazy_traversal_entry_count.checked_add(1).ok_or_else(|| {
                    litchi_iwa_common::Error::InvalidFormat(
                        "Numbers FormulaArchive repeated-entry count overflows host usize"
                            .to_owned(),
                    )
                })?;
        }
        if formula_ast_node_path(visit.path()) {
            if field.number() == 1 {
                scalar_visitor_node_count =
                    scalar_visitor_node_count.checked_add(1).ok_or_else(|| {
                        litchi_iwa_common::Error::InvalidFormat(
                            "Numbers FormulaArchive AST node count overflows host usize".to_owned(),
                        )
                    })?;
                let node_type = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .ok()
                    .and_then(|(value, width)| (width == field.payload().len()).then_some(value))
                    .and_then(|value| u32::try_from(value).ok());
                if node_type.is_none_or(|value| {
                    // The standalone LocalCellReferenceNode (27) is a
                    // compatibility-only node: unlike the nested local
                    // reference in CellReferenceNode, the scalar codec
                    // does not preserve its sticky flags.
                    value == 27 || !numbers_formula_codec::is_scalar_visitor_node_type(value)
                }) {
                    scalar_visitor_eligible = false;
                }
            } else if field.number() == 2 {
                // The streaming evaluator admits only the five function
                // identifiers below.  Unknown/native function IDs are valid
                // compatibility-renderer input, but attempting the scalar
                // visitor first would allocate its full node vector only to
                // fall back after the codec rejects the function.
                let function_identifier =
                    litchi_iwa_common::decode_varint_from_bytes(field.payload())
                        .ok()
                        .and_then(|(value, width)| {
                            (width == field.payload().len()).then_some(value)
                        })
                        .and_then(|value| u32::try_from(value).ok());
                if function_identifier.is_none_or(|value| !matches!(value, 15 | 30 | 84 | 88 | 168))
                {
                    scalar_visitor_eligible = false;
                }
            } else if !matches!(field.number(), 2 | 3 | 4 | 5 | 10 | 15 | 25 | 42 | 43) {
                // These are the only direct AST-node fields the compact
                // visitor can represent. Decimal128 sidecars (fields 42/43)
                // are validated by the scalar codec and do not change the
                // rendered number. Nested messages are checked by the
                // surrounding schema walk; any other direct field (strings,
                // dates, arrays, thunks, ranges, UIDs, or owner metadata)
                // falls back to the lossless compatibility renderer.
                scalar_visitor_eligible = false;
            }
        }
        if schema.utf8 {
            std::str::from_utf8(field.payload()).map_err(|_error| {
                litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive string field is not valid UTF-8".to_owned(),
                )
            })?;
        }
        validate_formula_scalar(field, schema.scalar)?;
        Ok(if schema.nested {
            WireDescent::Descend
        } else {
            WireDescent::Skip
        })
    });
    match preflight {
        Ok(report) => {
            debug_assert_eq!(attempted.fields, report.fields());
            debug_assert_eq!(attempted.work, report.scanned_bytes());
            Ok(FormulaEnvelopeReport {
                wire: Some(report),
                fields: report.fields(),
                scanned_bytes: report.scanned_bytes(),
                root_ast_present,
                scalar_visitor_eligible,
                scalar_visitor_node_count,
                lazy_traversal_entry_count,
            })
        },
        Err(error) => Err(FormulaEnvelopeFailure::new(
            error,
            AttemptedFormulaEnvelopeCost {
                fields: attempted.fields,
                work: attempted.work,
            },
        )),
    }
}

#[derive(Debug, Clone, Copy)]
struct FormulaEnvelopeCost {
    base_fields: usize,
    base_work: usize,
    fields: usize,
    work: usize,
}

impl FormulaEnvelopeCost {
    fn new(
        source_len: usize,
        base_fields: usize,
        base_work: usize,
        max_work: usize,
    ) -> litchi_iwa_common::Result<Self> {
        let mut cost = Self {
            base_fields,
            base_work,
            fields: 0,
            work: 0,
        };
        cost.charge_work(source_len, max_work)?;
        Ok(cost)
    }

    fn charge_field(&mut self, max_fields: usize) -> litchi_iwa_common::Result<()> {
        self.fields =
            self.fields
                .checked_add(1)
                .ok_or(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::Fields,
                    observed: usize::MAX,
                    limit: max_fields,
                })?;
        let observed = self.base_fields.saturating_add(self.fields);
        if observed > max_fields {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: max_fields,
            });
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize, max_work: usize) -> litchi_iwa_common::Result<()> {
        self.work =
            self.work
                .checked_add(amount)
                .ok_or(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::RewriteWork,
                    observed: usize::MAX,
                    limit: max_work,
                })?;
        let observed = self.base_work.saturating_add(self.work);
        if observed > max_work {
            return Err(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                observed,
                limit: max_work,
            });
        }
        Ok(())
    }

    fn charge_nested(&mut self, bytes: usize, max_work: usize) -> litchi_iwa_common::Result<()> {
        self.charge_work(bytes, max_work)
    }
}

fn required_formula_ast_node_type(source: &[u8]) -> litchi_iwa_common::Result<()> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .and_then(|limits| limits.with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_nesting(1))?;
    let mut found = false;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        field.validate_canonical_framing()?;
        if field.number() == 1 {
            if found {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive AST node type occurs more than once".to_owned(),
                ));
            }
            found = true;
            if field.wire_type() != 0 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "Numbers FormulaArchive AST node type has the wrong wire type".to_owned(),
                ));
            }
        }
        Ok(WireDescent::Skip)
    })?;
    if !found {
        return Err(litchi_iwa_common::Error::InvalidFormat(
            "Numbers FormulaArchive AST node type is missing required node type".to_owned(),
        ));
    }
    Ok(())
}

/// Check the required direct children of a schema-known nested message.
///
/// Prost's generated proto2 decoder does not consistently enforce required
/// fields on all of the deferred formula messages.  Keep this check local to
/// the nested payload selected by the schema so unknown opaque fields remain
/// untouched and repeated fields retain their normal semantics.
fn require_formula_fields(
    source: &[u8],
    message_path: &[u32],
    required_fields: &[u32],
) -> litchi_iwa_common::Result<()> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .and_then(|limits| limits.with_fields(source.len().clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_nesting(1))?;
    let mut seen = [false; 8];
    let mut known_fields = [0u32; 64];
    let mut known_field_count = 0usize;
    preflight_wire_tree_with_limits(source, limits, |visit| {
        let field = visit.field();
        field.validate_canonical_framing()?;
        let schema = formula_envelope_field(message_path, field.number());
        if schema.wire_type.is_some() && !formula_field_is_repeated(message_path, field.number()) {
            if known_fields[..known_field_count].contains(&field.number()) {
                if required_fields.contains(&field.number()) {
                    return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                        "Numbers FormulaArchive required field {} occurs more than once",
                        field.number()
                    )));
                }
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} occurs more than once",
                    field.number()
                )));
            }
            if known_field_count == known_fields.len() {
                return Err(litchi_iwa_common::Error::LimitExceeded {
                    kind: LimitKind::Fields,
                    observed: known_field_count.saturating_add(1),
                    limit: known_fields.len(),
                });
            }
            known_fields[known_field_count] = field.number();
            known_field_count += 1;
        }
        let Some(index) = required_fields
            .iter()
            .position(|required| *required == field.number())
        else {
            return Ok(WireDescent::Skip);
        };
        if seen[index] {
            return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                "Numbers FormulaArchive required field {} occurs more than once",
                field.number()
            )));
        }
        seen[index] = true;
        Ok(WireDescent::Skip)
    })?;
    for (index, field_number) in required_fields.iter().enumerate() {
        if !seen[index] {
            return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                "Numbers FormulaArchive required field {} is missing",
                field_number
            )));
        }
    }
    Ok(())
}

/// Return whether a schema-known field is allowed to occur more than once.
///
/// Prost accepts duplicate singular fields with last-value-wins semantics.
/// The formula archive ingress is stricter: only fields declared `repeated` in
/// the native schema may repeat. The path distinguishes repeated AST nodes
/// from the singular field-1 envelopes used by several nearby messages.
fn formula_field_is_repeated(path: &[u32], number: u32) -> bool {
    // Nested formula messages are visited with the complete path from the
    // FormulaArchive root (for example `[1, 1, 39, 1, 6]` for
    // CategoryReferenceArchive.CatRefUidList).  The repeated-field table is
    // expressed relative to the AST node, so strip that prefix before
    // matching the message-specific suffix.  Root AST-node arrays remain
    // addressable by their original path.
    let suffix = formula_ast_node_prefix_len(path).map_or(path, |prefix_len| &path[prefix_len..]);
    if number == 1 {
        formula_ast_array_path(path)
            || suffix == [38]
            || matches!(suffix, [38, 1, 1] | [38, 1, 2])
            || suffix == [39, 1, 6]
            || suffix == [40]
            || suffix == [45]
    } else {
        suffix == [40] && (2..=4).contains(&number)
    }
}

fn validate_canonical_formula_varint(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> litchi_iwa_common::Result<()> {
    let payload = field.payload();
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(payload).map_err(|error| {
        litchi_iwa_common::Error::InvalidFormat(format!(
            "Numbers FormulaArchive field {} has an invalid scalar varint: {error}",
            field.number()
        ))
    })?;
    if width != payload.len() || width != litchi_iwa_common::varint::encoded_len(value) {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "Numbers FormulaArchive field {} has a noncanonical scalar varint",
            field.number()
        )));
    }
    Ok(())
}

fn validate_formula_scalar(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    scalar: FormulaScalar,
) -> litchi_iwa_common::Result<()> {
    match scalar {
        FormulaScalar::Unknown => Ok(()),
        FormulaScalar::Varint => validate_canonical_formula_varint(field),
        FormulaScalar::Bool | FormulaScalar::U32 | FormulaScalar::Int32 | FormulaScalar::SInt32 => {
            let payload = field.payload();
            let (value, width) =
                litchi_iwa_common::decode_varint_from_bytes(payload).map_err(|error| {
                    litchi_iwa_common::Error::InvalidFormat(format!(
                        "Numbers FormulaArchive field {} has an invalid scalar varint: {error}",
                        field.number()
                    ))
                })?;
            if width != payload.len() || width != litchi_iwa_common::varint::encoded_len(value) {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} has a noncanonical scalar varint",
                    field.number()
                )));
            }
            let valid = match scalar {
                FormulaScalar::Bool => value <= 1,
                FormulaScalar::U32 | FormulaScalar::SInt32 => value <= u64::from(u32::MAX),
                FormulaScalar::Int32 => {
                    i32::try_from(value).is_ok() || value >= 0xffff_ffff_8000_0000
                },
                FormulaScalar::Unknown | FormulaScalar::Varint => true,
            };
            if !valid {
                return Err(litchi_iwa_common::Error::InvalidFormat(format!(
                    "Numbers FormulaArchive field {} has a noncanonical scalar value",
                    field.number()
                )));
            }
            Ok(())
        },
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaScalar {
    Unknown,
    Varint,
    Bool,
    U32,
    Int32,
    SInt32,
}

#[derive(Debug, Clone, Copy)]
struct FormulaEnvelopeField {
    wire_type: Option<u8>,
    nested: bool,
    utf8: bool,
    scalar: FormulaScalar,
    required_fields: &'static [u32],
}

const FORMULA_REQUIRED_NONE: &[u32] = &[];
const FORMULA_REQUIRED_AST_STICKY_BITS: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_AST_UID_TRACT: &[u32] = &[1, 2];
const FORMULA_REQUIRED_AST_UID: &[u32] = &[1, 2];
const FORMULA_REQUIRED_AST_CATEGORY_REFERENCE: &[u32] = &[1];
const FORMULA_REQUIRED_CATEGORY_REFERENCE: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_PRESERVE_FLAGS: &[u32] = &[1, 2];
const FORMULA_REQUIRED_LOCAL_REFERENCE: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_CROSS_REFERENCE: &[u32] = &[1, 2, 3, 4, 5];
const FORMULA_REQUIRED_COORDINATE: &[u32] = &[1];
const FORMULA_REQUIRED_UID_COORDINATE: &[u32] = &[1, 2, 3, 4];
const FORMULA_REQUIRED_AST_CATEGORY_LEVELS: &[u32] = &[1, 2];
const FORMULA_REQUIRED_CROSS_EXTRA: &[u32] = &[1];
const FORMULA_REQUIRED_RANGE_BEGIN: &[u32] = &[1];

const fn formula_envelope_scalar(wire_type: u8) -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(wire_type),
        nested: false,
        utf8: false,
        scalar: if wire_type == 0 {
            FormulaScalar::Varint
        } else {
            FormulaScalar::Unknown
        },
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_bool() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::Bool,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_u32() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::U32,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_int32() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::Int32,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_sint32() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(0),
        nested: false,
        utf8: false,
        scalar: FormulaScalar::SInt32,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_nested() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(2),
        nested: true,
        utf8: false,
        scalar: FormulaScalar::Unknown,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const fn formula_envelope_nested_required(required_fields: &'static [u32]) -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(2),
        nested: true,
        utf8: false,
        scalar: FormulaScalar::Unknown,
        required_fields,
    }
}

const fn formula_envelope_utf8() -> FormulaEnvelopeField {
    FormulaEnvelopeField {
        wire_type: Some(2),
        nested: false,
        utf8: true,
        scalar: FormulaScalar::Unknown,
        required_fields: FORMULA_REQUIRED_NONE,
    }
}

const FORMULA_ENVELOPE_UNKNOWN: FormulaEnvelopeField = FormulaEnvelopeField {
    wire_type: None,
    nested: false,
    utf8: false,
    scalar: FormulaScalar::Unknown,
    required_fields: FORMULA_REQUIRED_NONE,
};

/// Return the schema shape for a FormulaArchive field.  ASTNodeArray/ASTNode
/// paths are recognized structurally, so thunk arrays can recurse to any
/// bounded depth without treating arbitrary length-delimited strings/bytes as
/// messages.
fn formula_envelope_field(path: &[u32], number: u32) -> FormulaEnvelopeField {
    if path.is_empty() {
        return match number {
            1 => formula_envelope_nested(),
            2..=3 => formula_envelope_u32(),
            4..=5 => formula_envelope_bool(),
            6 => formula_envelope_nested(),
            7..=9 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        };
    }
    if path == [6] {
        return if (1..=5).contains(&number) {
            formula_envelope_bool()
        } else {
            FORMULA_ENVELOPE_UNKNOWN
        };
    }
    if matches!(path, [7] | [8] | [9]) {
        return if (1..=2).contains(&number) {
            formula_envelope_scalar(0)
        } else {
            FORMULA_ENVELOPE_UNKNOWN
        };
    }
    if formula_ast_array_path(path) {
        return if number == 1 {
            formula_envelope_nested()
        } else {
            FORMULA_ENVELOPE_UNKNOWN
        };
    }
    let Some(node_prefix_len) = formula_ast_node_prefix_len(path) else {
        return FORMULA_ENVELOPE_UNKNOWN;
    };
    formula_ast_suffix_field(&path[node_prefix_len..], number)
}

/// `ASTNodeArrayArchive` appears at `[1]`, then at a repeated `1,14` suffix
/// for every thunk (`ASTNodeArray.ast_node`, `ASTNode.thunk_array`).
fn formula_ast_array_path(path: &[u32]) -> bool {
    if path == [1] {
        return true;
    }
    if path.len() < 3 || path[0] != 1 {
        return false;
    }
    let mut index = 1;
    while index + 1 < path.len() && path[index..index + 2] == [1, 14] {
        index += 2;
    }
    index == path.len()
}

/// Return whether `path` identifies an ASTNodeArchive itself rather than one
/// of its nested coordinate/reference messages. The scalar visitor eligibility
/// pass uses this to distinguish direct node fields from nested wire fields.
fn formula_ast_node_path(path: &[u32]) -> bool {
    formula_ast_node_prefix_len(path) == Some(path.len())
}

/// Return the ASTNode path prefix length.  A node is reached through `1,1`,
/// then zero or more `14,1` thunk/array edges.  The remainder identifies a
/// child message (for example `[16, 5]` is a cross-table CFUUID).
fn formula_ast_node_prefix_len(path: &[u32]) -> Option<usize> {
    if path.len() < 2 || path[..2] != [1, 1] {
        return None;
    }
    let mut index = 2;
    while index + 1 < path.len() && path[index..index + 2] == [14, 1] {
        index += 2;
    }
    Some(index)
}

fn formula_ast_suffix_field(suffix: &[u32], number: u32) -> FormulaEnvelopeField {
    match suffix {
        // TSCE.ASTNodeArchive
        [] => match number {
            1 | 9 | 47 => formula_envelope_int32(),
            2..=3 | 11..=13 | 18 | 22..=24 | 37 | 46 => formula_envelope_u32(),
            5 | 10 | 19..=20 | 29 | 36 => formula_envelope_bool(),
            42..=43 => formula_envelope_scalar(0),
            4 | 7 | 8 => formula_envelope_scalar(1),
            6 | 17 | 21 | 25 | 34 | 35 => formula_envelope_utf8(),
            14 | 40 | 45 => formula_envelope_nested(),
            15 => formula_envelope_nested_required(FORMULA_REQUIRED_LOCAL_REFERENCE),
            16 => formula_envelope_nested_required(FORMULA_REQUIRED_CROSS_REFERENCE),
            26 | 27 => formula_envelope_nested_required(FORMULA_REQUIRED_COORDINATE),
            28 => formula_envelope_nested_required(FORMULA_REQUIRED_CROSS_EXTRA),
            30 => formula_envelope_nested_required(FORMULA_REQUIRED_UID_COORDINATE),
            33 | 41 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_STICKY_BITS),
            38 => formula_envelope_nested_required(&[2]),
            39 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_CATEGORY_REFERENCE),
            44 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_CATEGORY_LEVELS),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTNodeArrayArchive (ASTNode.thunk_array)
        [14] => {
            if number == 1 {
                formula_envelope_nested()
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // Local and cross-table cell reference messages are deliberately
        // separate: only the cross-table variant has field 5 (CFUUID) and
        // fields 6..9 (whitespace strings).
        [15] => match number {
            1..=4 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        [16] => match number {
            1..=4 => formula_envelope_u32(),
            5 => formula_envelope_nested(),
            6..=9 => formula_envelope_utf8(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSP.CFUUIDArchive
        [16, 5] | [28, 1] => match number {
            1 => formula_envelope_scalar(2),
            2..=5 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTColumn/RowCoordinateArchive
        [26] | [27] => match number {
            1 => formula_envelope_sint32(),
            2 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTStickyBits
        [33] | [41] | [38, 2] => match number {
            1..=4 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTCrossTableReferenceExtraInfoArchive
        [28] => match number {
            1 => formula_envelope_nested(),
            2..=5 => formula_envelope_utf8(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTUidCoordinateArchive
        [30] => match number {
            1 | 2 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID),
            3..=4 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSP.UUID
        [30, 1]
        | [30, 2]
        | [38, 1, 1, 1]
        | [38, 1, 2, 1]
        | [39, 1, 1]
        | [39, 1, 2]
        | [39, 1, 9]
        | [39, 1, 10]
        | [39, 1, 6, 1] => {
            if (1..=2).contains(&number) {
                formula_envelope_scalar(0)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.ASTUidTractList
        [38] => match number {
            1 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID_TRACT),
            2 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_STICKY_BITS),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTUidTract
        [38, 1] => match number {
            1..=2 => formula_envelope_nested(),
            3 | 5 => formula_envelope_bool(),
            4 => formula_envelope_int32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTUidList
        [38, 1, 1] | [38, 1, 2] => {
            if number == 1 {
                formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.ASTCategoryReferenceArchive
        [39] => {
            if number == 1 {
                formula_envelope_nested_required(FORMULA_REQUIRED_CATEGORY_REFERENCE)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.CategoryReferenceArchive
        [39, 1] => match number {
            1 | 2 | 9 | 10 => formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID),
            3 | 13 => formula_envelope_u32(),
            4 => formula_envelope_sint32(),
            8 => formula_envelope_int32(),
            11 | 12 | 14 => formula_envelope_bool(),
            6 => formula_envelope_nested(),
            7 => formula_envelope_nested_required(FORMULA_REQUIRED_PRESERVE_FLAGS),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.CategoryReferenceArchive.CatRefUidList
        [39, 1, 6] => {
            if number == 1 {
                formula_envelope_nested_required(FORMULA_REQUIRED_AST_UID)
            } else {
                FORMULA_ENVELOPE_UNKNOWN
            }
        },
        // TSCE.PreserveColumnRowFlagsArchive
        [39, 1, 7] => match number {
            1..=4 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTCategoryLevels
        [44] => match number {
            1..=3 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTColonTractArchive and its four range message variants.
        [40] => match number {
            1..=4 => formula_envelope_nested_required(FORMULA_REQUIRED_RANGE_BEGIN),
            5 => formula_envelope_bool(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        [40, 1] | [40, 2] => match number {
            1 | 2 => formula_envelope_int32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        [40, 3] | [40, 4] => match number {
            1 | 2 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        // TSCE.ASTLambdaIdentsListArchive
        [45] => match number {
            1 | 3 | 4 => formula_envelope_utf8(),
            2 => formula_envelope_u32(),
            _ => FORMULA_ENVELOPE_UNKNOWN,
        },
        _ => FORMULA_ENVELOPE_UNKNOWN,
    }
}
