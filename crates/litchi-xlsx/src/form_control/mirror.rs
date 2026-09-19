//! Source-backed scalar mirrors for the worksheet VML view of a form control.
//!
//! The leaf [`super::SourceView`] owns the `x14:formControlPr` splice.  This
//! module owns only the paired VML `x:ClientData` splice.  The VML range is
//! supplied by the worksheet owner together with the prefix whose namespace
//! binding that owner has already proved to be the Excel VML namespace.  This
//! keeps this small writer from guessing at prefixes, namespaces, shape
//! identity, or graph membership.
//!
//! A mirror edit is deliberately conservative.  It accepts a field only when
//! the mapping and lexical vocabulary are documented for this profile, both
//! current values agree, the VML field occurs exactly once, and the field has
//! no comments or nested opaque content that a scalar splice would discard.
//! Missing VML fields, unknown lexical values, duplicate fields, disagreements,
//! and object-type transitions are typed refusals.  A caller can therefore
//! compose the returned pair into its package transaction without observing a
//! one-sided edit.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_lines,
    reason = "the mapping table and bounded source scanner follow the XML field order"
)]

use std::fmt;
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

use super::budget::{
    RETAINED_BUDGET_HOLD_STORAGE_BYTES, RetainedBudgetHold, SOURCE_PAYLOAD_ARC_STORAGE_BYTES,
    retained_budget_hold,
};
use super::codec::SourceView;
use super::model::{
    Checked, DropStyle, EditValidation, FormControlFormula, KnownOrUnknown, ObjectType,
    ScalarField, ScalarValue, SelectionType, TextHAlign, TextVAlign,
};
use super::{
    FormControlError, MAX_ATTRIBUTES, MAX_FORMULA_BYTES, MAX_GENERATED_BYTES, MAX_PART_BYTES,
    MAX_XML_DEPTH, MAX_XML_EVENTS,
};
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};

const OBJECT_TYPE: &[u8] = b"ObjectType";

/// A source-qualified VML `ClientData` range.
///
/// `range` must start at the `<prefix:ClientData` start tag and end immediately
/// after its matching end tag.  `prefix` is compared byte-for-byte with every
/// `ClientData` and direct child field name.  The caller must supply the
/// namespace proof for that prefix; the helper intentionally does not inspect
/// ancestors outside `range`, because doing so would make an apparently local
/// edit depend on unrelated source bytes.
#[derive(Clone, Debug)]
pub(crate) struct ClientDataSource<'a> {
    source: &'a [u8],
    range: Range<usize>,
    prefix: Box<[u8]>,
}

impl<'a> ClientDataSource<'a> {
    /// Construct a source-qualified range.
    pub(crate) fn new(source: &'a [u8], range: Range<usize>, prefix: &[u8]) -> MirrorResult<Self> {
        if range.start >= range.end || range.end > source.len() {
            return Err(MirrorError::Invalid(
                "VML ClientData range is outside its source".to_owned(),
            ));
        }
        if prefix.len() > 128 || !valid_name(prefix, true) {
            return Err(MirrorError::Invalid(
                "VML ClientData prefix is not an XML name".to_owned(),
            ));
        }
        let mut owned_prefix = Vec::new();
        owned_prefix
            .try_reserve_exact(prefix.len())
            .map_err(|source| MirrorError::Allocation {
                resource: "VML ClientData prefix",
                source,
            })?;
        owned_prefix.extend_from_slice(prefix);
        Ok(Self {
            source,
            range,
            prefix: owned_prefix.into_boxed_slice(),
        })
    }
}

/// Bounded policy for one scalar mirror edit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct MirrorLimits {
    max_source_bytes: usize,
    max_output_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_fields: usize,
    max_value_bytes: usize,
    max_attributes: usize,
}

impl Default for MirrorLimits {
    fn default() -> Self {
        Self {
            max_source_bytes: MAX_PART_BYTES,
            max_output_bytes: MAX_GENERATED_BYTES,
            max_depth: MAX_XML_DEPTH,
            max_events: MAX_XML_EVENTS,
            max_fields: 65_536,
            max_value_bytes: MAX_FORMULA_BYTES,
            max_attributes: MAX_ATTRIBUTES,
        }
    }
}

impl MirrorLimits {
    /// Construct the standard bounded policy.
    #[must_use]
    pub(crate) const fn new() -> Self {
        Self {
            max_source_bytes: MAX_PART_BYTES,
            max_output_bytes: MAX_GENERATED_BYTES,
            max_depth: MAX_XML_DEPTH,
            max_events: MAX_XML_EVENTS,
            max_fields: 65_536,
            max_value_bytes: MAX_FORMULA_BYTES,
            max_attributes: MAX_ATTRIBUTES,
        }
    }

    /// Lower the source-member ceiling.
    #[must_use]
    pub(crate) fn with_max_source_bytes(mut self, value: usize) -> Self {
        self.max_source_bytes = value.min(MAX_PART_BYTES);
        self
    }

    /// Lower the combined candidate-part ceiling.
    #[must_use]
    pub(crate) fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = value.min(MAX_GENERATED_BYTES);
        self
    }

    /// Lower the XML nesting ceiling.
    #[must_use]
    pub(crate) fn with_max_depth(mut self, value: usize) -> Self {
        self.max_depth = value.min(MAX_XML_DEPTH);
        self
    }

    /// Lower the XML event ceiling.
    #[must_use]
    pub(crate) fn with_max_events(mut self, value: usize) -> Self {
        self.max_events = value.min(MAX_XML_EVENTS);
        self
    }

    /// Lower the number of retained direct VML fields.
    #[must_use]
    pub(crate) fn with_max_fields(mut self, value: usize) -> Self {
        self.max_fields = value.min(65_536);
        self
    }

    /// Lower the decoded scalar value ceiling.
    #[must_use]
    pub(crate) fn with_max_value_bytes(mut self, value: usize) -> Self {
        self.max_value_bytes = value.min(MAX_FORMULA_BYTES);
        self
    }

    /// Lower the per-tag attribute ceiling.
    #[must_use]
    pub(crate) fn with_max_attributes(mut self, value: usize) -> Self {
        self.max_attributes = value.min(MAX_ATTRIBUTES);
        self
    }

    pub(crate) const fn max_source_bytes(self) -> usize {
        self.max_source_bytes
    }

    pub(crate) const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    pub(crate) const fn max_depth(self) -> usize {
        self.max_depth
    }

    pub(crate) const fn max_events(self) -> usize {
        self.max_events
    }

    pub(crate) const fn max_fields(self) -> usize {
        self.max_fields
    }

    pub(crate) const fn max_value_bytes(self) -> usize {
        self.max_value_bytes
    }

    pub(crate) const fn max_attributes(self) -> usize {
        self.max_attributes
    }
}

/// A paired source result.  Both members are complete source-preserving
/// outputs and are intended to be committed together by the worksheet owner.
#[derive(Clone, Debug)]
pub(crate) struct MirroredScalarEdit {
    properties: Vec<u8>,
    vml: Vec<u8>,
    changed: bool,
    retained_budget: Option<Arc<RetainedBudgetHold>>,
}

impl PartialEq for MirroredScalarEdit {
    fn eq(&self, other: &Self) -> bool {
        self.properties == other.properties
            && self.vml == other.vml
            && self.changed == other.changed
    }
}

impl Eq for MirroredScalarEdit {}

impl MirroredScalarEdit {
    /// Exact candidate `x14:formControlPr` bytes.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn properties(&self) -> &[u8] {
        &self.properties
    }

    /// Exact candidate VML bytes.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn vml(&self) -> &[u8] {
        &self.vml
    }

    /// Whether either source member changed.
    #[must_use]
    pub(crate) const fn changed(&self) -> bool {
        self.changed
    }

    /// Consume the pair and return both owned source members.
    #[must_use]
    pub(crate) fn into_parts(self) -> (Vec<u8>, Vec<u8>) {
        (self.properties, self.vml)
    }

    pub(crate) fn into_parts_with_budget(
        self,
    ) -> (Vec<u8>, Vec<u8>, Option<Arc<RetainedBudgetHold>>) {
        (self.properties, self.vml, self.retained_budget)
    }
}

/// A refusal from the scalar mirror boundary.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub(crate) enum MirrorError {
    /// The caller's execution context cancelled or bounded the operation.
    #[error(transparent)]
    Execution(#[from] ExecutionError),
    /// The x14 source-backed leaf rejected its paired scalar splice.
    #[error(transparent)]
    Leaf(#[from] FormControlError),
    /// The supplied source range or XML is malformed.
    #[error("invalid VML scalar mirror source: {0}")]
    Invalid(String),
    /// The mapping is intentionally outside this bounded writer profile.
    #[error("unsupported VML scalar mirror for {field:?}: {reason}")]
    Unsupported {
        /// x14 field whose mapping was refused.
        field: ScalarField,
        /// Stable refusal reason.
        reason: &'static str,
    },
    /// The required VML occurrence is absent.
    #[error("missing VML scalar mirror {element} for {field:?}")]
    MissingSource {
        /// x14 field.
        field: ScalarField,
        /// VML direct-child name.
        element: &'static str,
    },
    /// A target VML field or attribute appears more than once.
    #[error("ambiguous VML scalar mirror {element} for {field:?}")]
    Ambiguous {
        /// x14 field.
        field: ScalarField,
        /// VML direct-child or attribute name.
        element: &'static str,
    },
    /// The two current source-backed views disagree.
    #[error("x14 and VML scalar mirrors disagree for {field:?}")]
    Disagreement {
        /// x14 field whose current values disagree.
        field: ScalarField,
    },
    /// A bounded resource was exceeded before construction.
    #[error("VML scalar mirror {resource} exceeds {maximum} (observed {observed})")]
    Limit {
        /// Stable resource name.
        resource: &'static str,
        /// Observed amount.
        observed: usize,
        /// Effective maximum.
        maximum: usize,
    },
    /// A fallible output or metadata reservation failed.
    #[error("could not reserve memory for VML scalar mirror {resource}: {source}")]
    Allocation {
        /// Allocation subject.
        resource: &'static str,
        /// Recoverable allocator error.
        #[source]
        source: std::collections::TryReserveError,
    },
}

/// Result type for this source-backed mirror helper.
pub(crate) type MirrorResult<T> = Result<T, MirrorError>;

/// Replace one scalar in both source-backed views.
///
/// The leaf `SourceView::replace_scalar` performs the x14 splice.  This
/// helper performs the corresponding VML splice only after it has parsed and
/// compared the current source values.  `None` is accepted only for an exact
/// already-absent effective default; clearing an authored field requires a
/// profile-specific insertion/removal policy and is refused here.
#[cfg(test)]
pub(crate) fn replace_scalar_pair(
    properties: &SourceView<'_>,
    client_data: ClientDataSource<'_>,
    field: ScalarField,
    value: Option<ScalarValue>,
) -> MirrorResult<MirroredScalarEdit> {
    replace_scalar_pair_with_limits(
        properties,
        client_data,
        field,
        value,
        MirrorLimits::default(),
        None,
    )
}

/// Replace one scalar in both source-backed views under a local policy.
pub(crate) fn replace_scalar_pair_with_limits(
    properties: &SourceView<'_>,
    client_data: ClientDataSource<'_>,
    field: ScalarField,
    value: Option<ScalarValue>,
    limits: MirrorLimits,
    context: Option<&ExecutionContext>,
) -> MirrorResult<MirroredScalarEdit> {
    let property_source = properties.source();
    let client_range = client_data.range.clone();
    if property_source.len() > limits.max_source_bytes() {
        return Err(MirrorError::Limit {
            resource: "x14 source bytes",
            observed: property_source.len(),
            maximum: limits.max_source_bytes(),
        });
    }
    let client_len = client_range
        .end
        .checked_sub(client_range.start)
        .ok_or_else(|| MirrorError::Invalid("VML ClientData range underflow".to_owned()))?;
    if client_len > limits.max_source_bytes() {
        return Err(MirrorError::Limit {
            resource: "VML ClientData source bytes",
            observed: client_len,
            maximum: limits.max_source_bytes(),
        });
    }
    if client_data.source.len() > limits.max_source_bytes() {
        return Err(MirrorError::Limit {
            resource: "VML source bytes",
            observed: client_data.source.len(),
            maximum: limits.max_source_bytes(),
        });
    }
    let input_len = property_source
        .len()
        .checked_add(client_data.source.len())
        .ok_or_else(|| MirrorError::Invalid("mirror input size overflow".to_owned()))?;
    let _input = reserve(
        context,
        Resource::InputBytes,
        input_len,
        "mirror input bytes",
    )?;
    let _depth = reserve(
        context,
        Resource::Depth,
        limits.max_depth(),
        "mirror XML depth",
    )?;
    let field_upper = limits
        .max_fields()
        .min(client_len.saturating_div(3).saturating_add(1));
    let _objects = reserve(context, Resource::Objects, field_upper, "mirror VML fields")?;
    let metadata_bytes = field_upper
        .checked_mul(size_of::<FieldRef>())
        .and_then(|value| {
            limits
                .max_depth()
                .checked_mul(size_of::<OpenFrame>())
                .and_then(|depth| value.checked_add(depth))
        })
        .ok_or_else(|| MirrorError::Invalid("mirror metadata size overflow".to_owned()))?;
    let _memory = reserve(context, Resource::Memory, metadata_bytes, "mirror metadata")?;
    let mut scan = ScanBudget::new(limits, context);
    let parsed = scan_client_data(&client_data, limits, &mut scan)?;
    scan.check()?;

    let spec = ScalarSpec::for_field(field)?;
    let props = properties.properties();
    let x14 = current_x14(&props, field)?;
    let vml = current_vml(&parsed, field, spec, limits)?;
    if let Some(vml_value) = vml.value.as_ref() {
        if *vml_value != x14.value {
            return Err(MirrorError::Disagreement { field });
        }
    } else if !spec.effective_default_matches(field, &x14.value) {
        if spec.is_effective() {
            return Err(MirrorError::Disagreement { field });
        }
        return Err(MirrorError::MissingSource {
            field,
            element: spec.element,
        });
    }

    validate_requested_value(field, value.as_ref())?;
    let Some(value) = value else {
        if !x14.authored && !vml.present && spec.is_effective() {
            return copy_pair(property_source, client_data.source, limits, context, false);
        }
        return Err(MirrorError::Unsupported {
            field,
            reason: "clearing an authored mirror requires an insertion/removal profile",
        });
    };

    if scalar_equal(field, &x14.value, &value)? {
        return copy_pair(property_source, client_data.source, limits, context, false);
    }
    if field == ScalarField::FmlaLink
        && (source_only_formula(&x14.value) || vml.value.as_ref().is_some_and(source_only_formula))
    {
        return Err(MirrorError::Unsupported {
            field,
            reason: "source-only formula values cannot be replaced",
        });
    }
    if spec.object_type && field == ScalarField::ObjectType {
        return Err(MirrorError::Unsupported {
            field,
            reason: "object-type transitions require shape and graph closure",
        });
    }
    let Some(vml_field) = vml.field else {
        return Err(MirrorError::MissingSource {
            field,
            element: spec.element,
        });
    };
    if vml_field.opaque {
        return Err(MirrorError::Unsupported {
            field,
            reason: "VML field contains comments or opaque nested content",
        });
    }
    if !x14.authored && !spec.is_effective() {
        return Err(MirrorError::MissingSource {
            field,
            element: spec.element,
        });
    }

    let x14_value_len = x14_encoded_len(field, &value, limits)?;
    let x14_edit_len = x14_edit_replacement_len(property_source, field, x14_value_len)?;
    let x14_output_len = x14_output_len(property_source, field, &value, limits)?;
    if x14_output_len > limits.max_output_bytes() {
        return Err(MirrorError::Limit {
            resource: "x14 generated output bytes",
            observed: x14_output_len,
            maximum: limits.max_output_bytes(),
        });
    }
    let vml_value_len = vml_replacement_len(client_data.source, &vml_field, field, &value, limits)?;
    let vml_edit_len = vml_edit_replacement_len(client_data.source, &vml_field, vml_value_len)?;
    let vml_output_len = vml_output_len(client_data.source, &vml_field, vml_edit_len)?;
    if vml_output_len > limits.max_output_bytes() {
        return Err(MirrorError::Limit {
            resource: "VML generated output bytes",
            observed: vml_output_len,
            maximum: limits.max_output_bytes(),
        });
    }
    let output_bytes = x14_output_len
        .checked_add(vml_output_len)
        .ok_or_else(|| MirrorError::Invalid("mirror output size overflow".to_owned()))?;
    let _output = reserve(
        context,
        Resource::OutputBytes,
        output_bytes,
        "mirror output bytes",
    )?;
    let wrapper_bytes = if vml_field.empty { vml_edit_len } else { 0 };
    let x14_wrapper_bytes = x14_edit_len
        .checked_sub(x14_value_len)
        .ok_or_else(|| MirrorError::Invalid("x14 replacement size underflow".to_owned()))?;
    let temporary_bytes = x14_value_len
        .checked_add(x14_wrapper_bytes)
        .and_then(|value| value.checked_add(vml_value_len))
        .and_then(|value| value.checked_add(wrapper_bytes))
        .and_then(|value| value.checked_add(output_bytes))
        .ok_or_else(|| MirrorError::Invalid("mirror staging size overflow".to_owned()))?;
    let retained_budget = reserve_retained_output(context, output_bytes, 2)?;
    let staging_bytes = temporary_bytes
        .checked_sub(output_bytes)
        .ok_or_else(|| MirrorError::Invalid("mirror staging size underflow".to_owned()))?;
    let _memory_output = reserve(
        context,
        Resource::Memory,
        staging_bytes,
        "mirror output staging",
    )?;
    scan.check()?;
    let replacement = render_vml_value(field, &value, limits)?;
    if replacement.len() != vml_value_len {
        return Err(MirrorError::Invalid(
            "VML encoded replacement length changed after preflight".to_owned(),
        ));
    }
    scan.check()?;
    let vml_edit = make_vml_edit(client_data.source, &vml_field, replacement, limits)?;
    scan.check()?;
    let x14_output = properties.replace_scalar(field, Some(value))?;
    if x14_output.len() != x14_output_len {
        return Err(MirrorError::Limit {
            resource: "x14 generated output bytes",
            observed: x14_output.len(),
            maximum: x14_output_len,
        });
    }
    scan.check()?;
    let vml_output = splice_one(
        client_data.source,
        vml_edit.range,
        &vml_edit.replacement,
        limits.max_output_bytes(),
    )?;
    if vml_output.len() != vml_output_len {
        return Err(MirrorError::Limit {
            resource: "VML generated output bytes",
            observed: vml_output.len(),
            maximum: vml_output_len,
        });
    }
    scan.check()?;
    Ok(MirroredScalarEdit {
        properties: x14_output,
        vml: vml_output,
        changed: true,
        retained_budget,
    })
}

#[derive(Clone, Copy, Debug)]
struct ScalarSpec {
    element: &'static str,
    object_type: bool,
    effective: bool,
}

impl ScalarSpec {
    fn for_field(field: ScalarField) -> MirrorResult<Self> {
        let spec = match field {
            ScalarField::ObjectType => Self {
                element: "ObjectType",
                object_type: true,
                effective: false,
            },
            ScalarField::Checked => Self::field("Checked"),
            ScalarField::Colored => Self::field("Colored"),
            ScalarField::DropLines => Self::field("DropLines"),
            ScalarField::DropStyle => Self::field("DropStyle"),
            ScalarField::Dx => Self::field("Dx"),
            ScalarField::FirstButton => Self::field("FirstButton"),
            ScalarField::FmlaLink => Self::field("FmlaLink"),
            ScalarField::Horiz => Self::effective("Horiz"),
            ScalarField::Inc => Self::field("Inc"),
            ScalarField::JustLastX => Self::field("JustLastX"),
            ScalarField::LockText => Self::effective("LockText"),
            ScalarField::Max => Self::field("Max"),
            ScalarField::Min => Self::effective("Min"),
            ScalarField::MultiSel => Self::field("MultiSel"),
            ScalarField::NoThreeD => Self::field("NoThreeD"),
            ScalarField::NoThreeD2 => Self::field("NoThreeD2"),
            ScalarField::Page => Self::field("Page"),
            ScalarField::Sel => Self::field("Sel"),
            ScalarField::SelType => Self::effective("SelType"),
            ScalarField::TextHAlign => Self::effective("TextHAlign"),
            ScalarField::TextVAlign => Self::effective("TextVAlign"),
            ScalarField::Val => Self::effective("Val"),
            ScalarField::WidthMin => Self::field("WidthMin"),
            // Both specifications define omitted VTEdit/editVal as Text.
            // Keep this as an effective row so an untouched pair with both
            // fields absent remains an exact no-op, while a real change still
            // requires an authored VML field (we do not invent one here).
            ScalarField::EditVal => Self::effective("VTEdit"),
            ScalarField::MultiLine => Self::field("MultiLine"),
            ScalarField::VerticalBar => Self::effective("VScroll"),
            ScalarField::PasswordEdit => Self::field("SecretEdit"),
            ScalarField::FmlaGroup | ScalarField::FmlaRange | ScalarField::FmlaTxbx => {
                return Err(MirrorError::Unsupported {
                    field,
                    reason: "formula/list graph mapping is read-only in this profile",
                });
            },
        };
        Ok(spec)
    }

    const fn field(element: &'static str) -> Self {
        Self {
            element,
            object_type: false,
            effective: false,
        }
    }

    const fn effective(element: &'static str) -> Self {
        Self {
            element,
            object_type: false,
            effective: true,
        }
    }

    const fn is_effective(self) -> bool {
        self.effective
    }

    fn effective_default_matches(self, field: ScalarField, value: &ScalarValue) -> bool {
        if !self.effective {
            return false;
        }
        match (field, value) {
            (ScalarField::Horiz | ScalarField::VerticalBar, ScalarValue::Boolean(false)) => true,
            (ScalarField::Min | ScalarField::Val, ScalarValue::Unsigned(0)) => true,
            (ScalarField::SelType, ScalarValue::SelectionType(SelectionType::Single)) => true,
            (ScalarField::TextHAlign, ScalarValue::TextHAlign(TextHAlign::Left)) => true,
            (ScalarField::TextVAlign, ScalarValue::TextVAlign(TextVAlign::Top)) => true,
            (ScalarField::EditVal, ScalarValue::EditValidation(EditValidation::Text)) => true,
            // VML lockText defaults true; x14's effective value must be true
            // only when it is authored explicitly.
            (ScalarField::LockText, ScalarValue::Boolean(true)) => true,
            _ => false,
        }
    }
}

#[derive(Clone, Debug)]
struct FieldRef {
    element: &'static str,
    span: Range<usize>,
    start_tag: Range<usize>,
    content: Range<usize>,
    qname: Range<usize>,
    empty: bool,
    empty_slash: Option<usize>,
    opaque: bool,
}

#[derive(Clone, Debug)]
struct AttrRef {
    value: Range<usize>,
}

#[derive(Clone, Debug)]
struct ParsedClientData<'a> {
    source: &'a [u8],
    fields: Vec<FieldRef>,
    object_type: Option<AttrRef>,
    object_type_duplicate: bool,
}

#[derive(Clone, Debug)]
struct OpenFrame {
    name: Range<usize>,
    direct_field: Option<usize>,
    opaque: bool,
}

struct ScanBudget<'a> {
    limits: MirrorLimits,
    context: Option<&'a ExecutionContext>,
    events: usize,
}

impl<'a> ScanBudget<'a> {
    fn new(limits: MirrorLimits, context: Option<&'a ExecutionContext>) -> Self {
        Self {
            limits,
            context,
            events: 0,
        }
    }

    fn event(&mut self) -> MirrorResult<()> {
        self.events = self
            .events
            .checked_add(1)
            .ok_or_else(|| MirrorError::Limit {
                resource: "VML XML events",
                observed: usize::MAX,
                maximum: self.limits.max_events(),
            })?;
        if self.events > self.limits.max_events() {
            return Err(MirrorError::Limit {
                resource: "VML XML events",
                observed: self.events,
                maximum: self.limits.max_events(),
            });
        }
        if let Some(context) = self.context {
            context.consume(Resource::Work, 1)?;
        }
        Ok(())
    }

    fn check(&self) -> MirrorResult<()> {
        if let Some(context) = self.context {
            context.check()?;
        }
        Ok(())
    }
}

fn scan_client_data<'a>(
    client_data: &'a ClientDataSource<'a>,
    limits: MirrorLimits,
    budget: &mut ScanBudget<'_>,
) -> MirrorResult<ParsedClientData<'a>> {
    let source = client_data.source;
    let range = client_data.range.clone();
    let bytes = source.get(range.clone()).ok_or_else(|| {
        MirrorError::Invalid("VML ClientData range is outside its source".to_owned())
    })?;
    if bytes.first() != Some(&b'<') {
        return Err(MirrorError::Invalid(
            "VML ClientData range does not start at an element".to_owned(),
        ));
    }
    let field_cap = limits
        .max_fields()
        .min(bytes.len().saturating_div(3).saturating_add(1));
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(field_cap)
        .map_err(|source| MirrorError::Allocation {
            resource: "VML field metadata",
            source,
        })?;
    let frame_cap = limits.max_depth().max(1);
    let mut stack = Vec::new();
    stack
        .try_reserve_exact(frame_cap)
        .map_err(|source| MirrorError::Allocation {
            resource: "VML element metadata",
            source,
        })?;

    let root_end = tag_end(bytes, 0)?;
    let root = parse_start_tag(bytes, 0, root_end, limits.max_attributes())?;
    check_qname(
        bytes,
        &root.name,
        client_data.prefix.as_ref(),
        b"ClientData",
    )?;
    if root.self_closing {
        if root_end != bytes.len() {
            return Err(MirrorError::Invalid(
                "trailing bytes follow self-closing VML ClientData".to_owned(),
            ));
        }
        return Ok(ParsedClientData {
            source,
            fields,
            object_type: root.object_type.map(|attribute| AttrRef {
                value: offset_range(attribute.value, range.start),
            }),
            object_type_duplicate: root.object_type_duplicate,
        });
    }
    stack.push(OpenFrame {
        name: root.name,
        direct_field: None,
        opaque: false,
    });
    let mut cursor = root_end;
    let mut root_closed = false;
    while cursor < bytes.len() {
        budget.event()?;
        if bytes[cursor] != b'<' {
            let text_end = text_end(bytes, cursor);
            if !bytes[cursor..text_end].iter().all(u8::is_ascii_whitespace) {
                let is_root_content = stack.len() == 1;
                if is_root_content {
                    return Err(MirrorError::Invalid(
                        "non-whitespace text occurs directly in VML ClientData".to_owned(),
                    ));
                }
            }
            let stack_depth = stack.len();
            let direct_index = stack.last().and_then(|frame| frame.direct_field);
            if let Some(index) = direct_index {
                if stack_depth == 2 {
                    let field = fields.get_mut(index).ok_or_else(|| {
                        MirrorError::Invalid("VML field frame is invalid".to_owned())
                    })?;
                    if field.content.start == field.content.end {
                        field.content = offset_range(cursor..text_end, range.start);
                    } else {
                        field.opaque = true;
                    }
                } else if let Some(frame) = stack.last_mut() {
                    frame.opaque = true;
                }
            }
            cursor = text_end;
            continue;
        }
        if bytes[cursor..].starts_with(b"<!--") {
            let end = bytes[cursor + 4..]
                .windows(3)
                .position(|window| window == b"-->")
                .map(|offset| cursor + 4 + offset + 3)
                .ok_or_else(|| MirrorError::Invalid("unterminated VML comment".to_owned()))?;
            if let Some(frame) = stack.last_mut() {
                if frame.direct_field.is_some() {
                    frame.opaque = true;
                }
            }
            if stack.len() > 1 {
                for frame in stack.iter_mut().rev().skip(1) {
                    if let Some(index) = frame.direct_field {
                        if let Some(field) = fields.get_mut(index) {
                            field.opaque = true;
                        }
                    }
                }
            }
            cursor = end;
            continue;
        }
        if bytes[cursor..].starts_with(b"<?") {
            let end = bytes[cursor + 2..]
                .windows(2)
                .position(|window| window == b"?>")
                .map(|offset| cursor + 2 + offset + 2)
                .ok_or_else(|| {
                    MirrorError::Invalid("unterminated VML processing instruction".to_owned())
                })?;
            mark_opaque(&mut stack, &mut fields);
            cursor = end;
            continue;
        }
        if bytes[cursor..].starts_with(b"<![CDATA[") {
            let end = bytes[cursor + 9..]
                .windows(3)
                .position(|window| window == b"]]>")
                .map(|offset| cursor + 9 + offset + 3)
                .ok_or_else(|| MirrorError::Invalid("unterminated VML CDATA".to_owned()))?;
            mark_opaque(&mut stack, &mut fields);
            cursor = end;
            continue;
        }
        if bytes[cursor..].starts_with(b"<!") {
            return Err(MirrorError::Invalid(
                "unsupported declaration inside VML ClientData".to_owned(),
            ));
        }
        if bytes[cursor..].starts_with(b"</") {
            let end = tag_end(bytes, cursor)?;
            let end_name = parse_end_name(bytes, cursor, end)?;
            let Some(frame) = stack.pop() else {
                return Err(MirrorError::Invalid("unexpected VML end tag".to_owned()));
            };
            if bytes[frame.name.clone()] != bytes[end_name.clone()] {
                return Err(MirrorError::Invalid(
                    "VML element end tag does not match start tag".to_owned(),
                ));
            }
            if let Some(index) = frame.direct_field {
                let field = fields
                    .get_mut(index)
                    .ok_or_else(|| MirrorError::Invalid("VML field frame is invalid".to_owned()))?;
                field.content.end = cursor.saturating_add(range.start);
                field.opaque |= frame.opaque;
                if field.content.start > field.content.end {
                    return Err(MirrorError::Invalid(
                        "VML field content range underflow".to_owned(),
                    ));
                }
            }
            cursor = end;
            if stack.is_empty() {
                root_closed = true;
                break;
            }
            continue;
        }
        let end = tag_end(bytes, cursor)?;
        let start = parse_start_tag(bytes, cursor, end, limits.max_attributes())?;
        let depth = stack.len();
        if depth >= limits.max_depth() {
            return Err(MirrorError::Limit {
                resource: "VML XML depth",
                observed: depth.saturating_add(1),
                maximum: limits.max_depth(),
            });
        }
        let direct = if depth == 1 {
            let prefix = qname_prefix(bytes, &start.name)?;
            if prefix != client_data.prefix.as_ref() {
                None
            } else {
                field_name(bytes, &start.name)
                    .map(|element| {
                        if fields.len() >= limits.max_fields() {
                            Err(MirrorError::Limit {
                                resource: "VML direct fields",
                                observed: fields.len().saturating_add(1),
                                maximum: limits.max_fields(),
                            })
                        } else {
                            let index = fields.len();
                            fields.push(FieldRef {
                                element,
                                span: offset_range(cursor..end, range.start),
                                start_tag: offset_range(cursor..end, range.start),
                                content: offset_range(end..end, range.start),
                                qname: offset_range(start.name.clone(), range.start),
                                empty: start.self_closing,
                                empty_slash: start
                                    .self_closing_slash
                                    .map(|value| value.saturating_add(range.start)),
                                opaque: false,
                            });
                            Ok(index)
                        }
                    })
                    .transpose()?
            }
        } else {
            None
        };
        if let Some(frame) = stack.last_mut() {
            if frame.direct_field.is_some() && depth >= 2 {
                frame.opaque = true;
            }
        }
        if start.self_closing {
            if let Some(index) = direct {
                if let Some(field) = fields.get_mut(index) {
                    field.span = offset_range(cursor..end, range.start);
                    field.empty = true;
                }
            }
            cursor = end;
            continue;
        }
        stack.push(OpenFrame {
            name: start.name,
            direct_field: direct,
            opaque: false,
        });
        cursor = end;
    }
    if !root_closed || !stack.is_empty() || cursor != bytes.len() {
        return Err(MirrorError::Invalid(
            "VML ClientData range does not contain one complete element".to_owned(),
        ));
    }
    Ok(ParsedClientData {
        source,
        fields,
        object_type: root.object_type.map(|attribute| AttrRef {
            value: offset_range(attribute.value, range.start),
        }),
        object_type_duplicate: root.object_type_duplicate,
    })
}

#[derive(Clone, Debug)]
struct ParsedStart {
    name: Range<usize>,
    self_closing: bool,
    self_closing_slash: Option<usize>,
    object_type: Option<AttrRef>,
    object_type_duplicate: bool,
}

fn parse_start_tag(
    source: &[u8],
    start: usize,
    end: usize,
    max_attributes: usize,
) -> MirrorResult<ParsedStart> {
    if end <= start + 2 || source.get(start) != Some(&b'<') || source.get(end - 1) != Some(&b'>') {
        return Err(MirrorError::Invalid("malformed VML start tag".to_owned()));
    }
    let mut cursor = start + 1;
    let name_start = cursor;
    while cursor + 1 < end && is_name_byte(source[cursor]) {
        cursor += 1;
    }
    if cursor == name_start {
        return Err(MirrorError::Invalid("VML start tag has no name".to_owned()));
    }
    let name = name_start..cursor;
    let mut attrs = 0usize;
    let mut object_type = None;
    let mut object_type_duplicate = false;
    let mut self_closing = false;
    let mut self_closing_slash = None;
    while cursor < end - 1 {
        while cursor < end - 1 && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end - 1 {
            break;
        }
        if source[cursor] == b'/' {
            if cursor + 1 != end - 1 {
                return Err(MirrorError::Invalid(
                    "VML start tag has trailing bytes".to_owned(),
                ));
            }
            self_closing = true;
            self_closing_slash = Some(cursor);
            cursor += 1;
            break;
        }
        attrs = attrs.checked_add(1).ok_or(MirrorError::Limit {
            resource: "VML attributes",
            observed: usize::MAX,
            maximum: max_attributes,
        })?;
        if attrs > max_attributes {
            return Err(MirrorError::Limit {
                resource: "VML attributes",
                observed: attrs,
                maximum: max_attributes,
            });
        }
        let attr_start = cursor;
        while cursor < end - 1 && is_name_byte(source[cursor]) {
            cursor += 1;
        }
        if cursor == attr_start {
            return Err(MirrorError::Invalid("VML attribute has no name".to_owned()));
        }
        let attr_name = attr_start..cursor;
        while cursor < end - 1 && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if source.get(cursor) != Some(&b'=') {
            return Err(MirrorError::Invalid(
                "VML attribute has no equals sign".to_owned(),
            ));
        }
        cursor += 1;
        while cursor < end - 1 && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *source
            .get(cursor)
            .ok_or_else(|| MirrorError::Invalid("VML attribute has no quote".to_owned()))?;
        if quote != b'\'' && quote != b'"' {
            return Err(MirrorError::Invalid(
                "VML attribute is not quoted".to_owned(),
            ));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < end - 1 && source[cursor] != quote {
            cursor += 1;
        }
        if cursor >= end - 1 {
            return Err(MirrorError::Invalid(
                "VML attribute quote is unterminated".to_owned(),
            ));
        }
        let value = value_start..cursor;
        cursor += 1;
        if source[attr_name.clone()] == *OBJECT_TYPE {
            if object_type.is_some() {
                object_type_duplicate = true;
            } else {
                object_type = Some(AttrRef { value });
            }
        }
    }
    if cursor != end - 1 {
        return Err(MirrorError::Invalid(
            "VML start tag is not terminated".to_owned(),
        ));
    }
    Ok(ParsedStart {
        name,
        self_closing,
        self_closing_slash,
        object_type,
        object_type_duplicate,
    })
}

fn tag_end(source: &[u8], start: usize) -> MirrorResult<usize> {
    if source.get(start) != Some(&b'<') {
        return Err(MirrorError::Invalid(
            "VML token does not start with '<'".to_owned(),
        ));
    }
    let mut quote = None;
    let mut cursor = start + 1;
    while cursor < source.len() {
        let byte = source[cursor];
        if let Some(expected) = quote {
            if byte == expected {
                quote = None;
            }
        } else if byte == b'\'' || byte == b'"' {
            quote = Some(byte);
        } else if byte == b'>' {
            return Ok(cursor + 1);
        }
        cursor += 1;
    }
    Err(MirrorError::Invalid("VML tag is not terminated".to_owned()))
}

fn parse_end_name(source: &[u8], start: usize, end: usize) -> MirrorResult<Range<usize>> {
    if end <= start + 3 || source.get(start..start + 2) != Some(b"</") {
        return Err(MirrorError::Invalid("malformed VML end tag".to_owned()));
    }
    let mut cursor = start + 2;
    let name_start = cursor;
    while cursor < end - 1 && is_name_byte(source[cursor]) {
        cursor += 1;
    }
    if cursor == name_start {
        return Err(MirrorError::Invalid("VML end tag has no name".to_owned()));
    }
    while cursor < end - 1 && source[cursor].is_ascii_whitespace() {
        cursor += 1;
    }
    if cursor != end - 1 {
        return Err(MirrorError::Invalid(
            "VML end tag has trailing bytes".to_owned(),
        ));
    }
    Ok(name_start..cursor)
}

fn text_end(source: &[u8], start: usize) -> usize {
    source[start..]
        .iter()
        .position(|byte| *byte == b'<')
        .map_or(source.len(), |offset| start + offset)
}

fn offset_range(range: Range<usize>, base: usize) -> Range<usize> {
    range.start.saturating_add(base)..range.end.saturating_add(base)
}

fn mark_opaque(stack: &mut [OpenFrame], fields: &mut [FieldRef]) {
    for frame in stack.iter_mut().rev() {
        if let Some(index) = frame.direct_field {
            frame.opaque = true;
            if let Some(field) = fields.get_mut(index) {
                field.opaque = true;
            }
            break;
        }
    }
}

fn valid_name(value: &[u8], allow_empty: bool) -> bool {
    if value.is_empty() {
        return allow_empty;
    }
    value.iter().enumerate().all(|(index, byte)| {
        if index == 0 {
            is_name_start(*byte)
        } else {
            is_name_byte(*byte)
        }
    })
}

fn is_name_start(byte: u8) -> bool {
    byte.is_ascii_alphabetic() || byte == b'_' || byte == b':'
}

fn is_name_byte(byte: u8) -> bool {
    is_name_start(byte) || byte.is_ascii_digit() || matches!(byte, b'-' | b'.')
}

fn qname_prefix<'a>(source: &'a [u8], qname: &Range<usize>) -> MirrorResult<&'a [u8]> {
    let value = source
        .get(qname.clone())
        .ok_or_else(|| MirrorError::Invalid("VML qualified name range is invalid".to_owned()))?;
    Ok(value
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&value[..0], |index| &value[..index]))
}

fn qname_local<'a>(source: &'a [u8], qname: &Range<usize>) -> MirrorResult<&'a [u8]> {
    let value = source
        .get(qname.clone())
        .ok_or_else(|| MirrorError::Invalid("VML qualified name range is invalid".to_owned()))?;
    Ok(value
        .iter()
        .position(|byte| *byte == b':')
        .map_or(value, |index| &value[index + 1..]))
}

fn check_qname(
    source: &[u8],
    qname: &Range<usize>,
    expected_prefix: &[u8],
    expected_local: &[u8],
) -> MirrorResult<()> {
    if qname_prefix(source, qname)? != expected_prefix
        || qname_local(source, qname)? != expected_local
    {
        return Err(MirrorError::Invalid(
            "VML ClientData QName is not source-qualified".to_owned(),
        ));
    }
    Ok(())
}

fn field_name(source: &[u8], qname: &Range<usize>) -> Option<&'static str> {
    let local = qname_local(source, qname).ok()?;
    match local {
        b"Checked" => Some("Checked"),
        b"Colored" => Some("Colored"),
        b"DropLines" => Some("DropLines"),
        b"DropStyle" => Some("DropStyle"),
        b"Dx" => Some("Dx"),
        b"FirstButton" => Some("FirstButton"),
        b"FmlaLink" => Some("FmlaLink"),
        b"Horiz" => Some("Horiz"),
        b"Inc" => Some("Inc"),
        b"JustLastX" => Some("JustLastX"),
        b"LockText" => Some("LockText"),
        b"Max" => Some("Max"),
        b"Min" => Some("Min"),
        b"MultiSel" => Some("MultiSel"),
        b"NoThreeD" => Some("NoThreeD"),
        b"NoThreeD2" => Some("NoThreeD2"),
        b"Page" => Some("Page"),
        b"Sel" => Some("Sel"),
        b"SelType" => Some("SelType"),
        b"TextHAlign" => Some("TextHAlign"),
        b"TextVAlign" => Some("TextVAlign"),
        b"Val" => Some("Val"),
        b"WidthMin" => Some("WidthMin"),
        b"VTEdit" => Some("VTEdit"),
        b"MultiLine" => Some("MultiLine"),
        b"VScroll" => Some("VScroll"),
        b"SecretEdit" => Some("SecretEdit"),
        _ => None,
    }
}

#[derive(Clone, Debug)]
struct CurrentVml {
    value: Option<ScalarValue>,
    field: Option<FieldRef>,
    present: bool,
}

#[derive(Clone, Debug)]
struct CurrentX14 {
    value: ScalarValue,
    authored: bool,
}

fn current_x14(
    properties: &super::codec::SourceProperties<'_>,
    field: ScalarField,
) -> MirrorResult<CurrentX14> {
    let authored = authored(properties, field);
    let value = match field {
        ScalarField::Horiz => ScalarValue::Boolean(properties.effective_horiz()),
        ScalarField::LockText => ScalarValue::Boolean(properties.effective_lock_text()),
        ScalarField::Min => ScalarValue::Unsigned(properties.effective_min()),
        ScalarField::SelType => match properties.effective_seltype() {
            KnownOrUnknown::Known(value) => ScalarValue::SelectionType(value),
            KnownOrUnknown::Unknown(_) => {
                return Err(MirrorError::Unsupported {
                    field,
                    reason: "x14 source has an unknown token",
                });
            },
        },
        ScalarField::TextHAlign => match properties.effective_text_h_align() {
            KnownOrUnknown::Known(value) => ScalarValue::TextHAlign(value),
            KnownOrUnknown::Unknown(_) => {
                return Err(MirrorError::Unsupported {
                    field,
                    reason: "x14 source has an unknown token",
                });
            },
        },
        ScalarField::TextVAlign => match properties.effective_text_v_align() {
            KnownOrUnknown::Known(value) => ScalarValue::TextVAlign(value),
            KnownOrUnknown::Unknown(_) => {
                return Err(MirrorError::Unsupported {
                    field,
                    reason: "x14 source has an unknown token",
                });
            },
        },
        ScalarField::Val => ScalarValue::Unsigned(properties.effective_val()),
        ScalarField::VerticalBar => ScalarValue::Boolean(properties.effective_vertical_bar()),
        ScalarField::EditVal => match properties.effective_edit_val() {
            KnownOrUnknown::Known(value) => ScalarValue::EditValidation(value),
            KnownOrUnknown::Unknown(_) => {
                return Err(MirrorError::Unsupported {
                    field,
                    reason: "x14 source has an unknown token",
                });
            },
        },
        _ => properties.scalar(field).ok_or_else(|| {
            if authored {
                MirrorError::Unsupported {
                    field,
                    reason: "x14 source has an unknown token",
                }
            } else {
                MirrorError::MissingSource {
                    field,
                    element: ScalarSpec::for_field(field)
                        .map(|spec| spec.element)
                        .unwrap_or("scalar"),
                }
            }
        })?,
    };
    Ok(CurrentX14 { value, authored })
}

fn authored(properties: &super::codec::SourceProperties<'_>, field: ScalarField) -> bool {
    match field {
        ScalarField::ObjectType => properties.object_type().is_some(),
        ScalarField::Checked => properties.checked().is_some(),
        ScalarField::Colored => properties.colored().is_some(),
        ScalarField::DropLines => properties.drop_lines().is_some(),
        ScalarField::DropStyle => properties.drop_style().is_some(),
        ScalarField::Dx => properties.dx().is_some(),
        ScalarField::FirstButton => properties.first_button().is_some(),
        ScalarField::FmlaGroup => properties.fmla_group().is_some(),
        ScalarField::FmlaLink => properties.fmla_link().is_some(),
        ScalarField::FmlaRange => properties.fmla_range().is_some(),
        ScalarField::FmlaTxbx => properties.fmla_txbx().is_some(),
        ScalarField::Horiz => properties.horiz().is_some(),
        ScalarField::Inc => properties.inc().is_some(),
        ScalarField::JustLastX => properties.just_last_x().is_some(),
        ScalarField::LockText => properties.lock_text().is_some(),
        ScalarField::Max => properties.max().is_some(),
        ScalarField::Min => properties.min().is_some(),
        ScalarField::MultiSel => properties.multi_sel().is_some(),
        ScalarField::NoThreeD => properties.no_three_d().is_some(),
        ScalarField::NoThreeD2 => properties.no_three_d2().is_some(),
        ScalarField::Page => properties.page().is_some(),
        ScalarField::Sel => properties.sel().is_some(),
        ScalarField::SelType => properties.seltype().is_some(),
        ScalarField::TextHAlign => properties.text_h_align().is_some(),
        ScalarField::TextVAlign => properties.text_v_align().is_some(),
        ScalarField::Val => properties.val().is_some(),
        ScalarField::WidthMin => properties.width_min().is_some(),
        ScalarField::EditVal => properties.edit_val().is_some(),
        ScalarField::MultiLine => properties.multi_line().is_some(),
        ScalarField::VerticalBar => properties.vertical_bar().is_some(),
        ScalarField::PasswordEdit => properties.password_edit().is_some(),
    }
}

fn current_vml(
    parsed: &ParsedClientData<'_>,
    field: ScalarField,
    spec: ScalarSpec,
    limits: MirrorLimits,
) -> MirrorResult<CurrentVml> {
    if spec.object_type {
        if parsed.object_type_duplicate {
            return Err(MirrorError::Ambiguous {
                field,
                element: "ObjectType",
            });
        }
        let Some(attribute) = parsed.object_type.as_ref() else {
            return Err(MirrorError::MissingSource {
                field,
                element: "ObjectType",
            });
        };
        let raw = decoded(
            parsed.source,
            attribute.value.clone(),
            limits.max_value_bytes(),
        )?;
        let value = match raw.as_str() {
            "Button" => ObjectType::Button,
            // MS-XLSX spells the x14 token `CheckBox`, while the VML
            // `ST_ObjectType` token is the fixture/spec-backed `Checkbox`.
            // Do not broaden this to case-folded or otherwise guessed names.
            "Checkbox" => ObjectType::CheckBox,
            "Radio" => ObjectType::Radio,
            _ => {
                return Err(MirrorError::Unsupported {
                    field,
                    reason: "VML ObjectType is outside the proven control mappings",
                });
            },
        };
        return Ok(CurrentVml {
            value: Some(ScalarValue::ObjectType(value)),
            field: None,
            present: true,
        });
    }
    let mut found = None;
    for candidate in &parsed.fields {
        if candidate.element == spec.element {
            if found.is_some() {
                return Err(MirrorError::Ambiguous {
                    field,
                    element: spec.element,
                });
            }
            found = Some(candidate.clone());
        }
    }
    let Some(candidate) = found else {
        return Ok(CurrentVml {
            value: None,
            field: None,
            present: false,
        });
    };
    let value = decode_vml_scalar(parsed.source, &candidate, field, limits)?;
    Ok(CurrentVml {
        value: Some(value),
        field: Some(candidate),
        present: true,
    })
}

fn decode_vml_scalar(
    source: &[u8],
    field: &FieldRef,
    x14_field: ScalarField,
    limits: MirrorLimits,
) -> MirrorResult<ScalarValue> {
    let raw = decoded(source, field.content.clone(), limits.max_value_bytes())?;
    match x14_field {
        ScalarField::Checked => match raw.as_str() {
            "0" => Ok(ScalarValue::Checked(Checked::Unchecked)),
            "1" => Ok(ScalarValue::Checked(Checked::Checked)),
            "2" => Ok(ScalarValue::Checked(Checked::Mixed)),
            _ => Err(MirrorError::Unsupported {
                field: x14_field,
                reason: "VML checked token is not decimal 0, 1, or 2",
            }),
        },
        ScalarField::Colored
        | ScalarField::FirstButton
        | ScalarField::Horiz
        | ScalarField::JustLastX
        | ScalarField::LockText
        | ScalarField::NoThreeD
        | ScalarField::NoThreeD2
        | ScalarField::MultiLine
        | ScalarField::VerticalBar
        | ScalarField::PasswordEdit => parse_vml_bool(&raw, field.empty, x14_field),
        ScalarField::DropStyle => match raw.as_str() {
            "Combo" => Ok(ScalarValue::DropStyle(DropStyle::Combo)),
            "ComboEdit" => Ok(ScalarValue::DropStyle(DropStyle::ComboEdit)),
            "Simple" => Ok(ScalarValue::DropStyle(DropStyle::Simple)),
            _ => Err(MirrorError::Unsupported {
                field: x14_field,
                reason: "VML drop style token is unknown",
            }),
        },
        ScalarField::SelType => match raw.as_str() {
            "Single" => Ok(ScalarValue::SelectionType(SelectionType::Single)),
            "Multi" => Ok(ScalarValue::SelectionType(SelectionType::Multi)),
            "Extend" => Ok(ScalarValue::SelectionType(SelectionType::Extended)),
            _ => Err(MirrorError::Unsupported {
                field: x14_field,
                reason: "VML selection token is unknown",
            }),
        },
        ScalarField::TextHAlign => match raw.as_str() {
            "Left" => Ok(ScalarValue::TextHAlign(TextHAlign::Left)),
            "Justify" => Ok(ScalarValue::TextHAlign(TextHAlign::Justify)),
            "Center" => Ok(ScalarValue::TextHAlign(TextHAlign::Center)),
            "Right" => Ok(ScalarValue::TextHAlign(TextHAlign::Right)),
            "Distributed" => Ok(ScalarValue::TextHAlign(TextHAlign::Distributed)),
            _ => Err(MirrorError::Unsupported {
                field: x14_field,
                reason: "VML horizontal alignment token is unknown",
            }),
        },
        ScalarField::TextVAlign => match raw.as_str() {
            "Top" => Ok(ScalarValue::TextVAlign(TextVAlign::Top)),
            "Justify" => Ok(ScalarValue::TextVAlign(TextVAlign::Justify)),
            "Center" => Ok(ScalarValue::TextVAlign(TextVAlign::Center)),
            "Bottom" => Ok(ScalarValue::TextVAlign(TextVAlign::Bottom)),
            "Distributed" => Ok(ScalarValue::TextVAlign(TextVAlign::Distributed)),
            _ => Err(MirrorError::Unsupported {
                field: x14_field,
                reason: "VML vertical alignment token is unknown",
            }),
        },
        ScalarField::EditVal => match raw.as_str() {
            "0" => Ok(ScalarValue::EditValidation(EditValidation::Text)),
            "1" => Ok(ScalarValue::EditValidation(EditValidation::Integer)),
            "2" => Ok(ScalarValue::EditValidation(EditValidation::Number)),
            "3" => Ok(ScalarValue::EditValidation(EditValidation::Reference)),
            "4" => Ok(ScalarValue::EditValidation(EditValidation::Formula)),
            _ => Err(MirrorError::Unsupported {
                field: x14_field,
                reason: "VML edit validation token is not decimal 0 through 4",
            }),
        },
        ScalarField::DropLines
        | ScalarField::Dx
        | ScalarField::Inc
        | ScalarField::Max
        | ScalarField::Min
        | ScalarField::Page
        | ScalarField::Sel
        | ScalarField::Val
        | ScalarField::WidthMin => parse_vml_unsigned(&raw, x14_field),
        ScalarField::FmlaLink => Ok(ScalarValue::Formula(
            FormControlFormula::from_source(raw).map_err(MirrorError::Leaf)?,
        )),
        ScalarField::MultiSel => Ok(ScalarValue::String(raw)),
        ScalarField::FmlaGroup
        | ScalarField::FmlaRange
        | ScalarField::FmlaTxbx
        | ScalarField::ObjectType => Err(MirrorError::Unsupported {
            field: x14_field,
            reason: "VML field is outside the scalar mirror profile",
        }),
    }
}

fn parse_vml_bool(raw: &str, empty: bool, field: ScalarField) -> MirrorResult<ScalarValue> {
    if empty || raw.is_empty() {
        return Ok(ScalarValue::Boolean(true));
    }
    match raw {
        "t" | "true" | "True" => Ok(ScalarValue::Boolean(true)),
        "f" | "false" | "False" => Ok(ScalarValue::Boolean(false)),
        _ => Err(MirrorError::Unsupported {
            field,
            reason: "VML boolean token is outside the proven lexical set",
        }),
    }
}

fn parse_vml_unsigned(raw: &str, field: ScalarField) -> MirrorResult<ScalarValue> {
    if raw.is_empty() || !raw.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(MirrorError::Unsupported {
            field,
            reason: "VML numeric token is not an unsigned decimal",
        });
    }
    let value = raw.parse::<u32>().map_err(|_| MirrorError::Unsupported {
        field,
        reason: "VML numeric token exceeds the x14 unsigned bound",
    })?;
    Ok(ScalarValue::Unsigned(value))
}

fn decoded(source: &[u8], range: Range<usize>, max_bytes: usize) -> MirrorResult<String> {
    let raw = source.get(range).ok_or_else(|| {
        MirrorError::Invalid("VML scalar value range is outside its source".to_owned())
    })?;
    if raw.len() > max_bytes {
        return Err(MirrorError::Limit {
            resource: "VML scalar value bytes",
            observed: raw.len(),
            maximum: max_bytes,
        });
    }
    let raw = std::str::from_utf8(raw)
        .map_err(|_| MirrorError::Invalid("VML scalar text is not UTF-8".to_owned()))?;
    decode_vml_text(raw, max_bytes)
}

fn decode_vml_text(value: &str, maximum: usize) -> MirrorResult<String> {
    let bytes = value.as_bytes();
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| MirrorError::Allocation {
            resource: "decoded VML scalar value",
            source,
        })?;
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        if bytes[cursor] != b'&' {
            output.push(bytes[cursor]);
            cursor += 1;
            continue;
        }
        let entity_start = cursor + 1;
        let entity_end = bytes[entity_start..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|offset| entity_start + offset)
            .ok_or_else(|| {
                MirrorError::Invalid("VML scalar text has an unterminated XML entity".to_owned())
            })?;
        let entity = &bytes[entity_start..entity_end];
        let decoded = match entity {
            b"amp" => b"&".as_slice(),
            b"lt" => b"<".as_slice(),
            b"gt" => b">".as_slice(),
            b"quot" => b"\"".as_slice(),
            b"apos" => b"'".as_slice(),
            _ if entity.starts_with(b"#x") || entity.starts_with(b"#X") => {
                let value = std::str::from_utf8(&entity[2..])
                    .ok()
                    .and_then(|value| u32::from_str_radix(value, 16).ok())
                    .ok_or_else(|| {
                        MirrorError::Invalid(
                            "VML scalar text has an invalid numeric entity".to_owned(),
                        )
                    })?;
                append_codepoint(&mut output, value, maximum)?;
                cursor = entity_end + 1;
                continue;
            },
            _ if entity.starts_with(b"#") => {
                let value = std::str::from_utf8(&entity[1..])
                    .ok()
                    .and_then(|value| value.parse::<u32>().ok())
                    .ok_or_else(|| {
                        MirrorError::Invalid(
                            "VML scalar text has an invalid numeric entity".to_owned(),
                        )
                    })?;
                append_codepoint(&mut output, value, maximum)?;
                cursor = entity_end + 1;
                continue;
            },
            _ => {
                return Err(MirrorError::Invalid(
                    "VML scalar text has an unknown XML entity".to_owned(),
                ));
            },
        };
        output.extend_from_slice(decoded);
        if output.len() > maximum {
            return Err(MirrorError::Limit {
                resource: "decoded VML scalar value bytes",
                observed: output.len(),
                maximum,
            });
        }
        cursor = entity_end + 1;
    }
    String::from_utf8(output)
        .map_err(|_| MirrorError::Invalid("decoded VML scalar text is not UTF-8".to_owned()))
}

fn append_codepoint(output: &mut Vec<u8>, value: u32, maximum: usize) -> MirrorResult<()> {
    let character = char::from_u32(value).ok_or_else(|| {
        MirrorError::Invalid("VML scalar numeric entity is not a Unicode scalar".to_owned())
    })?;
    let mut encoded = [0u8; 4];
    let bytes = character.encode_utf8(&mut encoded).as_bytes();
    let new_len = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| MirrorError::Invalid("decoded VML scalar size overflow".to_owned()))?;
    if new_len > maximum {
        return Err(MirrorError::Limit {
            resource: "decoded VML scalar value bytes",
            observed: new_len,
            maximum,
        });
    }
    output
        .try_reserve(bytes.len())
        .map_err(|source| MirrorError::Allocation {
            resource: "decoded VML scalar value",
            source,
        })?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn validate_requested_value(field: ScalarField, value: Option<&ScalarValue>) -> MirrorResult<()> {
    if let Some(value) = value {
        let valid = match field {
            ScalarField::ObjectType => matches!(value, ScalarValue::ObjectType(_)),
            ScalarField::Checked => matches!(value, ScalarValue::Checked(_)),
            ScalarField::Colored
            | ScalarField::FirstButton
            | ScalarField::Horiz
            | ScalarField::JustLastX
            | ScalarField::LockText
            | ScalarField::NoThreeD
            | ScalarField::NoThreeD2
            | ScalarField::MultiLine
            | ScalarField::VerticalBar
            | ScalarField::PasswordEdit => matches!(value, ScalarValue::Boolean(_)),
            ScalarField::DropLines
            | ScalarField::Dx
            | ScalarField::Inc
            | ScalarField::Max
            | ScalarField::Min
            | ScalarField::Page
            | ScalarField::Sel
            | ScalarField::Val
            | ScalarField::WidthMin => matches!(value, ScalarValue::Unsigned(_)),
            ScalarField::DropStyle => matches!(value, ScalarValue::DropStyle(_)),
            ScalarField::FmlaLink => matches!(value, ScalarValue::Formula(_)),
            ScalarField::MultiSel => matches!(value, ScalarValue::String(_)),
            ScalarField::SelType => matches!(value, ScalarValue::SelectionType(_)),
            ScalarField::TextHAlign => matches!(value, ScalarValue::TextHAlign(_)),
            ScalarField::TextVAlign => matches!(value, ScalarValue::TextVAlign(_)),
            ScalarField::EditVal => matches!(value, ScalarValue::EditValidation(_)),
            ScalarField::FmlaGroup | ScalarField::FmlaRange | ScalarField::FmlaTxbx => false,
        };
        if !valid {
            return Err(MirrorError::Invalid(
                "scalar value has the wrong field type".to_owned(),
            ));
        }
    }
    Ok(())
}

fn scalar_equal(
    field: ScalarField,
    current: &ScalarValue,
    requested: &ScalarValue,
) -> MirrorResult<bool> {
    validate_requested_value(field, Some(requested))?;
    Ok(current == requested)
}

fn source_only_formula(value: &ScalarValue) -> bool {
    matches!(value, ScalarValue::Formula(formula) if formula.source_only())
}

fn x14_output_len(
    source: &[u8],
    field: ScalarField,
    value: &ScalarValue,
    limits: MirrorLimits,
) -> MirrorResult<usize> {
    let encoded_len = x14_encoded_len(field, value, limits)?;
    let name = x14_wire_name(field);
    let output_len = if let Some(range) = x14_attribute_value_range(source, name, field)? {
        source
            .len()
            .checked_sub(range.end - range.start)
            .and_then(|value| value.checked_add(encoded_len))
            .ok_or_else(|| MirrorError::Invalid("x14 scalar output size overflow".to_owned()))?
    } else {
        source
            .len()
            .checked_add(name.len())
            .and_then(|value| value.checked_add(encoded_len))
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| MirrorError::Invalid("x14 scalar output size overflow".to_owned()))?
    };
    Ok(output_len)
}

fn x14_edit_replacement_len(
    source: &[u8],
    field: ScalarField,
    encoded_len: usize,
) -> MirrorResult<usize> {
    if x14_attribute_value_range(source, x14_wire_name(field), field)?.is_some() {
        return Ok(encoded_len);
    }
    x14_wire_name(field)
        .len()
        .checked_add(encoded_len)
        .and_then(|value| value.checked_add(4))
        .ok_or_else(|| MirrorError::Invalid("x14 scalar replacement size overflow".to_owned()))
}

fn x14_encoded_len(
    field: ScalarField,
    value: &ScalarValue,
    limits: MirrorLimits,
) -> MirrorResult<usize> {
    let length = match (field, value) {
        (ScalarField::ObjectType, ScalarValue::ObjectType(value)) => value.wire().len(),
        (ScalarField::Checked, ScalarValue::Checked(value)) => value.wire().len(),
        (ScalarField::DropStyle, ScalarValue::DropStyle(value)) => value.wire().len(),
        (ScalarField::SelType, ScalarValue::SelectionType(value)) => value.wire().len(),
        (ScalarField::TextHAlign, ScalarValue::TextHAlign(value)) => value.wire().len(),
        (ScalarField::TextVAlign, ScalarValue::TextVAlign(value)) => value.wire().len(),
        (ScalarField::EditVal, ScalarValue::EditValidation(value)) => value.wire().len(),
        (
            ScalarField::Colored
            | ScalarField::FirstButton
            | ScalarField::Horiz
            | ScalarField::JustLastX
            | ScalarField::LockText
            | ScalarField::NoThreeD
            | ScalarField::NoThreeD2
            | ScalarField::MultiLine
            | ScalarField::VerticalBar
            | ScalarField::PasswordEdit,
            ScalarValue::Boolean(_),
        ) => 1,
        (
            ScalarField::DropLines
            | ScalarField::Dx
            | ScalarField::Inc
            | ScalarField::Max
            | ScalarField::Min
            | ScalarField::Page
            | ScalarField::Sel
            | ScalarField::Val
            | ScalarField::WidthMin,
            ScalarValue::Unsigned(value),
        ) => {
            let mut buffer = itoa::Buffer::new();
            buffer.format(*value).len()
        },
        (ScalarField::FmlaGroup | ScalarField::FmlaLink, ScalarValue::Formula(value)) => {
            escaped_x14_len(value.as_str(), limits.max_value_bytes())?
        },
        (ScalarField::MultiSel, ScalarValue::String(value)) => {
            escaped_x14_len(value.as_str(), limits.max_value_bytes())?
        },
        _ => {
            return Err(MirrorError::Invalid(
                "scalar field/value pair is malformed".to_owned(),
            ));
        },
    };
    if length > limits.max_value_bytes() {
        return Err(MirrorError::Limit {
            resource: "generated x14 scalar value bytes",
            observed: length,
            maximum: limits.max_value_bytes(),
        });
    }
    Ok(length)
}

fn escaped_x14_len(value: &str, maximum: usize) -> MirrorResult<usize> {
    let mut length = 0usize;
    for byte in value.bytes() {
        let additional = match byte {
            b'<' | b'>' => 3,
            b'&' => 4,
            b'\'' | b'"' => 5,
            b'\t' | b'\n' | b'\r' => 4,
            _ => 0,
        };
        length = length
            .checked_add(1 + additional)
            .ok_or_else(|| MirrorError::Invalid("x14 scalar output size overflow".to_owned()))?;
    }
    if length > maximum {
        return Err(MirrorError::Limit {
            resource: "generated x14 scalar value bytes",
            observed: length,
            maximum,
        });
    }
    Ok(length)
}

fn x14_wire_name(field: ScalarField) -> &'static [u8] {
    match field {
        ScalarField::ObjectType => b"objectType",
        ScalarField::Checked => b"checked",
        ScalarField::Colored => b"colored",
        ScalarField::DropLines => b"dropLines",
        ScalarField::DropStyle => b"dropStyle",
        ScalarField::Dx => b"dx",
        ScalarField::FirstButton => b"firstButton",
        ScalarField::FmlaGroup => b"fmlaGroup",
        ScalarField::FmlaLink => b"fmlaLink",
        ScalarField::FmlaRange => b"fmlaRange",
        ScalarField::FmlaTxbx => b"fmlaTxbx",
        ScalarField::Horiz => b"horiz",
        ScalarField::Inc => b"inc",
        ScalarField::JustLastX => b"justLastX",
        ScalarField::LockText => b"lockText",
        ScalarField::Max => b"max",
        ScalarField::Min => b"min",
        ScalarField::MultiSel => b"multiSel",
        ScalarField::NoThreeD => b"noThreeD",
        ScalarField::NoThreeD2 => b"noThreeD2",
        ScalarField::Page => b"page",
        ScalarField::Sel => b"sel",
        ScalarField::SelType => b"seltype",
        ScalarField::TextHAlign => b"textHAlign",
        ScalarField::TextVAlign => b"textVAlign",
        ScalarField::Val => b"val",
        ScalarField::WidthMin => b"widthMin",
        ScalarField::EditVal => b"editVal",
        ScalarField::MultiLine => b"multiLine",
        ScalarField::VerticalBar => b"verticalBar",
        ScalarField::PasswordEdit => b"passwordEdit",
    }
}

fn x14_attribute_value_range(
    source: &[u8],
    expected: &[u8],
    field: ScalarField,
) -> MirrorResult<Option<Range<usize>>> {
    let (start, end) = x14_root_start(source)?;
    let mut cursor = start + 1;
    while cursor < end - 1 && is_name_byte(source[cursor]) {
        cursor += 1;
    }
    if cursor == start + 1 {
        return Err(MirrorError::Invalid(
            "x14 root has no element name".to_owned(),
        ));
    }
    let mut found = None;
    while cursor < end - 1 {
        while cursor < end - 1 && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end - 1 {
            break;
        }
        if source[cursor] == b'/' {
            if cursor + 1 != end - 1 {
                return Err(MirrorError::Invalid(
                    "x14 root has trailing bytes".to_owned(),
                ));
            }
            break;
        }
        let name_start = cursor;
        while cursor < end - 1 && is_name_byte(source[cursor]) {
            cursor += 1;
        }
        if cursor == name_start {
            return Err(MirrorError::Invalid(
                "x14 root attribute has no name".to_owned(),
            ));
        }
        let name = name_start..cursor;
        while cursor < end - 1 && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if source.get(cursor) != Some(&b'=') {
            return Err(MirrorError::Invalid(
                "x14 root attribute has no equals sign".to_owned(),
            ));
        }
        cursor += 1;
        while cursor < end - 1 && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *source
            .get(cursor)
            .ok_or_else(|| MirrorError::Invalid("x14 root attribute has no quote".to_owned()))?;
        if quote != b'\'' && quote != b'"' {
            return Err(MirrorError::Invalid(
                "x14 root attribute is not quoted".to_owned(),
            ));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < end - 1 && source[cursor] != quote {
            cursor += 1;
        }
        if cursor >= end - 1 {
            return Err(MirrorError::Invalid(
                "x14 root attribute quote is unterminated".to_owned(),
            ));
        }
        let value = value_start..cursor;
        cursor += 1;
        if source[name.clone()] == *expected {
            if found.is_some() {
                return Err(MirrorError::Ambiguous {
                    field,
                    element: "x14 scalar attribute",
                });
            }
            found = Some(value);
        }
    }
    Ok(found)
}

fn x14_root_start(source: &[u8]) -> MirrorResult<(usize, usize)> {
    let mut cursor = 0usize;
    while let Some(offset) = source[cursor..].iter().position(|byte| *byte == b'<') {
        let start = cursor + offset;
        let next = source.get(start + 1).copied();
        if !matches!(next, Some(b'?') | Some(b'!') | Some(b'/')) {
            return Ok((start, tag_end(source, start)?));
        }
        if source[start..].starts_with(b"<!--") {
            let end = source[start + 4..]
                .windows(3)
                .position(|window| window == b"-->")
                .map(|offset| start + 4 + offset + 3)
                .ok_or_else(|| MirrorError::Invalid("unterminated x14 comment".to_owned()))?;
            cursor = end;
        } else {
            cursor = tag_end(source, start)?;
        }
    }
    Err(MirrorError::Invalid(
        "x14 root element is absent".to_owned(),
    ))
}

fn vml_replacement_len(
    _source: &[u8],
    _field: &FieldRef,
    field: ScalarField,
    value: &ScalarValue,
    limits: MirrorLimits,
) -> MirrorResult<usize> {
    vml_encoded_len(field, value, limits)
}

fn vml_encoded_len(
    field: ScalarField,
    value: &ScalarValue,
    limits: MirrorLimits,
) -> MirrorResult<usize> {
    let length = match (field, value) {
        (
            ScalarField::DropLines
            | ScalarField::Dx
            | ScalarField::Inc
            | ScalarField::Max
            | ScalarField::Min
            | ScalarField::Page
            | ScalarField::Sel
            | ScalarField::Val
            | ScalarField::WidthMin,
            ScalarValue::Unsigned(value),
        ) => {
            let mut buffer = itoa::Buffer::new();
            buffer.format(*value).len()
        },
        (ScalarField::FmlaLink, ScalarValue::Formula(value)) => {
            escaped_vml_len(value.as_str(), limits.max_value_bytes())?
        },
        (ScalarField::MultiSel, ScalarValue::String(value)) => {
            escaped_vml_len(value.as_str(), limits.max_value_bytes())?
        },
        _ => escaped_vml_len(vml_static_lexical(field, value)?, limits.max_value_bytes())?,
    };
    if length > limits.max_value_bytes() {
        return Err(MirrorError::Limit {
            resource: "generated VML scalar value bytes",
            observed: length,
            maximum: limits.max_value_bytes(),
        });
    }
    Ok(length)
}

fn vml_static_lexical(field: ScalarField, value: &ScalarValue) -> MirrorResult<&'static str> {
    match (field, value) {
        (ScalarField::Checked, ScalarValue::Checked(value)) => Ok(match value {
            Checked::Unchecked => "0",
            Checked::Checked => "1",
            Checked::Mixed => "2",
        }),
        (
            ScalarField::Colored
            | ScalarField::FirstButton
            | ScalarField::Horiz
            | ScalarField::JustLastX
            | ScalarField::LockText
            | ScalarField::NoThreeD
            | ScalarField::NoThreeD2
            | ScalarField::MultiLine
            | ScalarField::VerticalBar
            | ScalarField::PasswordEdit,
            ScalarValue::Boolean(value),
        ) => Ok(if *value { "True" } else { "False" }),
        (ScalarField::DropStyle, ScalarValue::DropStyle(value)) => Ok(match value {
            DropStyle::Combo => "Combo",
            DropStyle::ComboEdit => "ComboEdit",
            DropStyle::Simple => "Simple",
        }),
        (ScalarField::SelType, ScalarValue::SelectionType(value)) => Ok(match value {
            SelectionType::Single => "Single",
            SelectionType::Multi => "Multi",
            SelectionType::Extended => "Extend",
        }),
        (ScalarField::TextHAlign, ScalarValue::TextHAlign(value)) => Ok(match value {
            TextHAlign::Left => "Left",
            TextHAlign::Justify => "Justify",
            TextHAlign::Center => "Center",
            TextHAlign::Right => "Right",
            TextHAlign::Distributed => "Distributed",
        }),
        (ScalarField::TextVAlign, ScalarValue::TextVAlign(value)) => Ok(match value {
            TextVAlign::Top => "Top",
            TextVAlign::Justify => "Justify",
            TextVAlign::Center => "Center",
            TextVAlign::Bottom => "Bottom",
            TextVAlign::Distributed => "Distributed",
        }),
        (ScalarField::EditVal, ScalarValue::EditValidation(value)) => Ok(match value {
            EditValidation::Text => "0",
            EditValidation::Integer => "1",
            EditValidation::Number => "2",
            EditValidation::Reference => "3",
            EditValidation::Formula => "4",
        }),
        (ScalarField::ObjectType, ScalarValue::ObjectType(value)) => match value {
            ObjectType::Button => Ok("Button"),
            ObjectType::CheckBox => Ok("Checkbox"),
            ObjectType::Radio => Ok("Radio"),
            _ => Err(MirrorError::Unsupported {
                field,
                reason: "VML ObjectType is outside the proven control mappings",
            }),
        },
        _ => Err(MirrorError::Invalid(
            "scalar field/value pair is malformed".to_owned(),
        )),
    }
}

fn vml_edit_replacement_len(
    source: &[u8],
    field: &FieldRef,
    encoded_len: usize,
) -> MirrorResult<usize> {
    if !field.empty {
        return Ok(encoded_len);
    }
    let slash = field
        .empty_slash
        .ok_or_else(|| MirrorError::Invalid("VML empty field has no slash".to_owned()))?;
    let qname = source
        .get(field.qname.clone())
        .ok_or_else(|| MirrorError::Invalid("VML field QName range is invalid".to_owned()))?;
    slash
        .checked_sub(field.start_tag.start)
        .and_then(|value| value.checked_add(1))
        .and_then(|value| value.checked_add(encoded_len))
        .and_then(|value| value.checked_add(3))
        .and_then(|value| value.checked_add(qname.len()))
        .ok_or_else(|| MirrorError::Invalid("VML field replacement size overflow".to_owned()))
}

fn vml_output_len(source: &[u8], field: &FieldRef, replacement_len: usize) -> MirrorResult<usize> {
    let removed = if field.empty {
        field.span.end.checked_sub(field.span.start)
    } else {
        field.content.end.checked_sub(field.content.start)
    }
    .ok_or_else(|| MirrorError::Invalid("VML field span underflow".to_owned()))?;
    source
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or_else(|| MirrorError::Invalid("VML output size overflow".to_owned()))
}

struct VmlEdit {
    range: Range<usize>,
    replacement: Vec<u8>,
}

fn render_vml_value(
    field: ScalarField,
    value: &ScalarValue,
    limits: MirrorLimits,
) -> MirrorResult<Vec<u8>> {
    let lexical = match (field, value) {
        (
            ScalarField::DropLines
            | ScalarField::Dx
            | ScalarField::Inc
            | ScalarField::Max
            | ScalarField::Min
            | ScalarField::Page
            | ScalarField::Sel
            | ScalarField::Val
            | ScalarField::WidthMin,
            ScalarValue::Unsigned(value),
        ) => {
            return ascii_u32(*value, limits.max_value_bytes());
        },
        (ScalarField::FmlaLink, ScalarValue::Formula(value)) => value.as_str(),
        (ScalarField::MultiSel, ScalarValue::String(value)) => value.as_str(),
        _ => vml_static_lexical(field, value)?,
    };
    escape_vml_text(lexical, limits.max_value_bytes())
}

fn ascii_u32(value: u32, maximum: usize) -> MirrorResult<Vec<u8>> {
    let mut buffer = itoa::Buffer::new();
    let text = buffer.format(value);
    if text.len() > maximum {
        return Err(MirrorError::Limit {
            resource: "generated VML scalar value bytes",
            observed: text.len(),
            maximum,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(text.len())
        .map_err(|source| MirrorError::Allocation {
            resource: "VML numeric scalar value",
            source,
        })?;
    output.extend_from_slice(text.as_bytes());
    Ok(output)
}

fn escape_vml_text(value: &str, maximum: usize) -> MirrorResult<Vec<u8>> {
    let length = escaped_vml_len(value, maximum)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| MirrorError::Allocation {
            resource: "VML scalar value",
            source,
        })?;
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\'' => output.extend_from_slice(b"&apos;"),
            _ => output.push(byte),
        }
    }
    Ok(output)
}

fn escaped_vml_len(value: &str, maximum: usize) -> MirrorResult<usize> {
    let mut length = 0usize;
    for byte in value.bytes() {
        let extra = match byte {
            b'&' => 4,
            b'<' | b'>' => 3,
            b'"' | b'\'' => 5,
            _ => 0,
        };
        length = length
            .checked_add(1)
            .and_then(|value| value.checked_add(extra))
            .ok_or_else(|| MirrorError::Invalid("VML scalar escaping size overflow".to_owned()))?;
    }
    if length > maximum {
        return Err(MirrorError::Limit {
            resource: "generated VML scalar value bytes",
            observed: length,
            maximum,
        });
    }
    Ok(length)
}

fn make_vml_edit(
    source: &[u8],
    field: &FieldRef,
    replacement: Vec<u8>,
    limits: MirrorLimits,
) -> MirrorResult<VmlEdit> {
    if field.element == "ObjectType" {
        return Ok(VmlEdit {
            range: field.content.clone(),
            replacement,
        });
    }
    if field.opaque {
        return Err(MirrorError::Unsupported {
            field: ScalarField::MultiSel,
            reason: "VML field contains comments or opaque nested content",
        });
    }
    if field.empty {
        let slash = field
            .empty_slash
            .ok_or_else(|| MirrorError::Invalid("VML empty field has no slash".to_owned()))?;
        let qname = source
            .get(field.qname.clone())
            .ok_or_else(|| MirrorError::Invalid("VML field QName range is invalid".to_owned()))?;
        let replacement_len = slash
            .checked_sub(field.start_tag.start)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| value.checked_add(replacement.len()))
            .and_then(|value| value.checked_add(3))
            .and_then(|value| value.checked_add(qname.len()))
            .ok_or_else(|| {
                MirrorError::Invalid("VML field replacement size overflow".to_owned())
            })?;
        if replacement_len > limits.max_output_bytes() {
            return Err(MirrorError::Limit {
                resource: "VML field replacement bytes",
                observed: replacement_len,
                maximum: limits.max_output_bytes(),
            });
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(replacement_len)
            .map_err(|source| MirrorError::Allocation {
                resource: "VML field replacement",
                source,
            })?;
        output.extend_from_slice(&source[field.start_tag.start..slash]);
        output.push(b'>');
        output.extend_from_slice(&replacement);
        output.extend_from_slice(b"</");
        output.extend_from_slice(qname);
        output.push(b'>');
        return Ok(VmlEdit {
            range: field.span.clone(),
            replacement: output,
        });
    }
    Ok(VmlEdit {
        range: field.content.clone(),
        replacement,
    })
}

fn splice_one(
    source: &[u8],
    range: Range<usize>,
    replacement: &[u8],
    maximum: usize,
) -> MirrorResult<Vec<u8>> {
    if range.start > range.end || range.end > source.len() {
        return Err(MirrorError::Invalid(
            "VML splice range is outside its source".to_owned(),
        ));
    }
    let output_len = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|value| value.checked_add(replacement.len()))
        .ok_or_else(|| MirrorError::Invalid("VML splice output size overflow".to_owned()))?;
    if output_len > maximum {
        return Err(MirrorError::Limit {
            resource: "VML generated output bytes",
            observed: output_len,
            maximum,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| MirrorError::Allocation {
            resource: "VML generated output",
            source,
        })?;
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[range.end..]);
    Ok(output)
}

fn copy_pair(
    properties: &[u8],
    vml: &[u8],
    limits: MirrorLimits,
    context: Option<&ExecutionContext>,
    changed: bool,
) -> MirrorResult<MirroredScalarEdit> {
    let total = properties
        .len()
        .checked_add(vml.len())
        .ok_or_else(|| MirrorError::Invalid("mirror copy size overflow".to_owned()))?;
    if properties.len() > limits.max_output_bytes() {
        return Err(MirrorError::Limit {
            resource: "x14 generated output bytes",
            observed: properties.len(),
            maximum: limits.max_output_bytes(),
        });
    }
    if vml.len() > limits.max_output_bytes() {
        return Err(MirrorError::Limit {
            resource: "VML generated output bytes",
            observed: vml.len(),
            maximum: limits.max_output_bytes(),
        });
    }
    let _output = reserve(context, Resource::OutputBytes, total, "mirror copy output")?;
    let retained_budget = reserve_retained_output(context, total, 2)?;
    let staging = total
        .checked_sub(total)
        .ok_or_else(|| MirrorError::Invalid("mirror copy staging size underflow".to_owned()))?;
    let _memory = reserve(context, Resource::Memory, staging, "mirror copy staging")?;
    let mut property_copy = Vec::new();
    property_copy
        .try_reserve_exact(properties.len())
        .map_err(|source| MirrorError::Allocation {
            resource: "x14 mirror copy",
            source,
        })?;
    property_copy.extend_from_slice(properties);
    let mut vml_copy = Vec::new();
    vml_copy
        .try_reserve_exact(vml.len())
        .map_err(|source| MirrorError::Allocation {
            resource: "VML mirror copy",
            source,
        })?;
    vml_copy.extend_from_slice(vml);
    if let Some(context) = context {
        context.check()?;
    }
    Ok(MirroredScalarEdit {
        properties: property_copy,
        vml: vml_copy,
        changed,
        retained_budget,
    })
}

fn reserve_retained_output(
    context: Option<&ExecutionContext>,
    output_bytes: usize,
    objects: usize,
) -> MirrorResult<Option<Arc<RetainedBudgetHold>>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let payload_storage = SOURCE_PAYLOAD_ARC_STORAGE_BYTES
        .checked_mul(2)
        .ok_or_else(|| {
            MirrorError::Invalid("mirror retained payload storage overflow".to_owned())
        })?;
    let memory_bytes = output_bytes
        .checked_add(payload_storage)
        .and_then(|bytes| bytes.checked_add(RETAINED_BUDGET_HOLD_STORAGE_BYTES))
        .ok_or_else(|| MirrorError::Invalid("mirror retained storage overflow".to_owned()))?;
    let memory = reserve(
        Some(context),
        Resource::Memory,
        memory_bytes,
        "mirror retained output",
    )?;
    let objects = reserve(
        Some(context),
        Resource::Objects,
        objects,
        "mirror retained output objects",
    )?;
    Ok(Some(retained_budget_hold(memory, objects)))
}

fn reserve(
    context: Option<&ExecutionContext>,
    resource: Resource,
    amount: usize,
    name: &'static str,
) -> MirrorResult<Option<Reservation>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let amount = u64::try_from(amount).map_err(|_| MirrorError::Limit {
        resource: name,
        observed: usize::MAX,
        maximum: u64::MAX as usize,
    })?;
    if amount == 0 {
        return Ok(None);
    }
    Ok(Some(context.reserve(resource, amount)?))
}

impl fmt::Display for ClientDataSource<'_> {
    fn fmt(&self, output: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            output,
            "ClientData[{}..{}]",
            self.range.start, self.range.end
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::form_control::{FormControlFormula, inspect};

    const EXCEL: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    fn x14(checked: &str, extra: &str) -> Vec<u8> {
        format!(
            r#"<x14:formControlPr xmlns:x14="{EXCEL}" checked="{checked}"{extra}/>"#,
            EXCEL = EXCEL
        )
        .into_bytes()
    }

    fn client_data(body: &str) -> Vec<u8> {
        format!(
            r#"<x:ClientData ObjectType="Checkbox">{body}</x:ClientData>"#,
            body = body
        )
        .into_bytes()
    }

    #[test]
    fn checked_pair_splice_preserves_vml_lexical_context() {
        let x14_source = x14("Checked", "");
        let view = inspect(&x14_source).expect("x14 fixture");
        let vml = client_data("<!-- keep --><x:Checked>1</x:Checked><x:Opaque a=\"b\"/>");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let edited = replace_scalar_pair(
            &view,
            source,
            ScalarField::Checked,
            Some(ScalarValue::Checked(Checked::Unchecked)),
        )
        .expect("paired edit");
        assert!(edited.changed());
        assert_eq!(edited.properties(), x14("Unchecked", ""));
        assert_eq!(
            edited.vml(),
            b"<x:ClientData ObjectType=\"Checkbox\"><!-- keep --><x:Checked>0</x:Checked><x:Opaque a=\"b\"/></x:ClientData>"
        );
    }

    #[test]
    fn boolean_empty_element_uses_proven_lexical_set() {
        let x14 = x14("Checked", " colored=\"1\"");
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:Colored/><x:Checked>1</x:Checked>");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let edited = replace_scalar_pair(
            &view,
            source,
            ScalarField::Colored,
            Some(ScalarValue::Boolean(false)),
        )
        .expect("paired edit");
        assert_eq!(
            edited.vml(),
            b"<x:ClientData ObjectType=\"Checkbox\"><x:Colored>False</x:Colored><x:Checked>1</x:Checked></x:ClientData>"
        );
    }

    #[test]
    fn unknown_boolean_is_refused_without_touching_x14() {
        let x14 = x14("Checked", " colored=\"1\"");
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:Colored>yes</x:Colored><x:Checked>1</x:Checked>");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let error = replace_scalar_pair(
            &view,
            source,
            ScalarField::Colored,
            Some(ScalarValue::Boolean(false)),
        )
        .expect_err("unknown VML bool must refuse");
        assert!(matches!(
            error,
            MirrorError::Unsupported {
                field: ScalarField::Colored,
                ..
            }
        ));
    }

    #[test]
    fn disagreement_is_refused() {
        let x14 = x14("Checked", " horiz=\"1\"");
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:Horiz>False</x:Horiz><x:Checked>1</x:Checked>");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let error = replace_scalar_pair(
            &view,
            source,
            ScalarField::Horiz,
            Some(ScalarValue::Boolean(false)),
        )
        .expect_err("disagreement must refuse");
        assert!(matches!(
            error,
            MirrorError::Disagreement {
                field: ScalarField::Horiz
            }
        ));
    }

    #[test]
    fn object_type_transition_is_refused() {
        let x14 = x14("Checked", " objectType=\"CheckBox\"");
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:Checked>1</x:Checked>");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let error = replace_scalar_pair(
            &view,
            source,
            ScalarField::ObjectType,
            Some(ScalarValue::ObjectType(ObjectType::Radio)),
        )
        .expect_err("object type transition must refuse");
        assert!(matches!(
            error,
            MirrorError::Unsupported {
                field: ScalarField::ObjectType,
                ..
            }
        ));
    }

    #[test]
    fn opaque_target_content_is_refused() {
        let x14 = x14("Checked", " colored=\"1\"");
        let view = inspect(&x14).expect("x14 fixture");
        let vml =
            client_data("<x:Colored><!-- keep --><x:Nested/></x:Colored><x:Checked>1</x:Checked>");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let error = replace_scalar_pair(
            &view,
            source,
            ScalarField::Colored,
            Some(ScalarValue::Boolean(false)),
        )
        .expect_err("opaque target must refuse");
        assert!(matches!(
            error,
            MirrorError::Unsupported {
                field: ScalarField::Colored,
                ..
            }
        ));
    }

    #[test]
    fn source_noop_copies_exact_bytes() {
        let x14 = x14("Checked", "");
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:Checked>1</x:Checked><!-- after -->");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let edited = replace_scalar_pair(
            &view,
            source,
            ScalarField::Checked,
            Some(ScalarValue::Checked(Checked::Checked)),
        )
        .expect("no-op");
        assert!(!edited.changed());
        assert_eq!(edited.properties(), x14.as_slice());
        assert_eq!(edited.vml(), vml.as_slice());
    }

    #[test]
    fn nonzero_client_data_range_splices_the_complete_vml_member() {
        let x14 = x14("Checked", "");
        let view = inspect(&x14).expect("x14 fixture");
        let vml = b"<!-- before --><x:ClientData ObjectType=\"Checkbox\"><x:Checked>1</x:Checked></x:ClientData><!-- after -->".to_vec();
        let start = vml
            .windows(b"<x:ClientData".len())
            .position(|window| window == b"<x:ClientData")
            .expect("ClientData start");
        let end = vml[start..]
            .windows(b"</x:ClientData>".len())
            .position(|window| window == b"</x:ClientData>")
            .map(|offset| start + offset + b"</x:ClientData>".len())
            .expect("ClientData end");
        let source = ClientDataSource::new(&vml, start..end, b"x").expect("VML fixture");
        let edited = replace_scalar_pair(
            &view,
            source,
            ScalarField::Checked,
            Some(ScalarValue::Checked(Checked::Unchecked)),
        )
        .expect("paired edit");
        assert_eq!(
            edited.vml(),
            b"<!-- before --><x:ClientData ObjectType=\"Checkbox\"><x:Checked>0</x:Checked></x:ClientData><!-- after -->"
        );
    }

    #[test]
    fn formula_source_value_is_escaped_and_bounded() {
        let formula = FormControlFormula::new("'Sheet&A'!A1").expect("formula");
        let x14 = format!(
            r#"<x14:formControlPr xmlns:x14="{EXCEL}" fmlaLink="Sheet1!A1"/>"#,
            EXCEL = EXCEL
        )
        .into_bytes();
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:FmlaLink>Sheet1!A1</x:FmlaLink>");
        let source = ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let edited = replace_scalar_pair(
            &view,
            source,
            ScalarField::FmlaLink,
            Some(ScalarValue::Formula(formula)),
        )
        .expect("paired formula edit");
        assert!(edited.vml().windows(5).any(|window| window == b"&amp;"));
    }

    #[test]
    fn exact_output_limit_accepts_numeric_shrink_and_one_under_refuses() {
        let x14 = x14("Checked", " dropLines=\"10\"");
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:DropLines>10</x:DropLines><x:Checked>1</x:Checked>");
        let make_source = || ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let target = ScalarValue::Unsigned(2);

        let default_edit = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::DropLines,
            Some(target.clone()),
            MirrorLimits::default(),
            None,
        )
        .expect("default budget accepts a shrinking edit");
        assert!(default_edit.properties().len() < x14.len());
        assert!(default_edit.vml().len() < vml.len());
        let exact = default_edit
            .properties()
            .len()
            .max(default_edit.vml().len());

        let exact_edit = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::DropLines,
            Some(target.clone()),
            MirrorLimits::default().with_max_output_bytes(exact),
            None,
        )
        .expect("the exact final output ceiling accepts the edit");
        assert_eq!(exact_edit.properties(), default_edit.properties());
        assert_eq!(exact_edit.vml(), default_edit.vml());

        let error = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::DropLines,
            Some(target),
            MirrorLimits::default().with_max_output_bytes(exact - 1),
            None,
        )
        .expect_err("one byte below the final output must refuse before rendering");
        assert!(matches!(error, MirrorError::Limit { .. }));
    }

    #[test]
    fn exact_output_limit_accounts_for_effective_x14_attribute_insertion() {
        let x14_source = x14("Checked", "");
        let view = inspect(&x14_source).expect("x14 fixture");
        let vml = client_data("<x:Horiz>False</x:Horiz><x:Checked>1</x:Checked>");
        let make_source = || ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let target = ScalarValue::Boolean(true);

        let default_edit = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::Horiz,
            Some(target.clone()),
            MirrorLimits::default(),
            None,
        )
        .expect("default budget accepts an effective-value insertion");
        assert!(
            default_edit
                .properties()
                .windows(9)
                .any(|window| window == b"horiz=\"1\"")
        );
        let exact = default_edit
            .properties()
            .len()
            .max(default_edit.vml().len());

        let exact_edit = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::Horiz,
            Some(target.clone()),
            MirrorLimits::default().with_max_output_bytes(exact),
            None,
        )
        .expect("the exact insertion output ceiling accepts the edit");
        assert_eq!(exact_edit.properties(), default_edit.properties());
        assert_eq!(exact_edit.vml(), default_edit.vml());

        let error = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::Horiz,
            Some(target),
            MirrorLimits::default().with_max_output_bytes(exact - 1),
            None,
        )
        .expect_err("one byte below inserted output must refuse before rendering");
        assert!(matches!(error, MirrorError::Limit { .. }));
    }

    #[test]
    fn exact_output_limit_accounts_for_escaped_formula_and_one_under_refuses() {
        let formula = FormControlFormula::new("'Sheet&A'!A1").expect("formula");
        let x14 = format!(
            r#"<x14:formControlPr xmlns:x14="{EXCEL}" fmlaLink="Sheet1!A1"/>"#,
            EXCEL = EXCEL
        )
        .into_bytes();
        let view = inspect(&x14).expect("x14 fixture");
        let vml = client_data("<x:FmlaLink>Sheet1!A1</x:FmlaLink>");
        let make_source = || ClientDataSource::new(&vml, 0..vml.len(), b"x").expect("VML fixture");
        let target = ScalarValue::Formula(formula.clone());

        let default_edit = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::FmlaLink,
            Some(target.clone()),
            MirrorLimits::default(),
            None,
        )
        .expect("default budget accepts the escaped formula");
        assert!(
            default_edit
                .properties()
                .windows(5)
                .any(|window| window == b"&amp;")
        );
        assert!(
            default_edit
                .vml()
                .windows(5)
                .any(|window| window == b"&amp;")
        );
        let exact = default_edit
            .properties()
            .len()
            .max(default_edit.vml().len());

        let exact_edit = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::FmlaLink,
            Some(target.clone()),
            MirrorLimits::default().with_max_output_bytes(exact),
            None,
        )
        .expect("the exact escaped output ceiling accepts the edit");
        assert_eq!(exact_edit.properties(), default_edit.properties());
        assert_eq!(exact_edit.vml(), default_edit.vml());

        let error = replace_scalar_pair_with_limits(
            &view,
            make_source(),
            ScalarField::FmlaLink,
            Some(target),
            MirrorLimits::default().with_max_output_bytes(exact - 1),
            None,
        )
        .expect_err("one byte below escaped output must refuse before rendering");
        assert!(matches!(error, MirrorError::Limit { .. }));
    }
}
