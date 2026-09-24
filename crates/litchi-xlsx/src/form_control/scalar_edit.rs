//! Selector-first, source-checked scalar edits for existing worksheet controls.

use std::io::Write;
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use litchi_core::{ExecutionContext, ReadAt, Resource};
use litchi_opc::{
    AuthoredXmlFragment, OpcPackage, PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits,
    SourceTopologyPlan, SourceXmlPart,
};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::{NsReader, Reader};

use super::{
    ClientDataSource, ControlSelector, FormControlCollection, FormControlView, MirrorLimits,
    OwnerLimits, OwnerProfile, Properties, ScalarField, ScalarValue,
    budget::{ARC_HEADER_BYTES, reserve_generated},
    replace_scalar_pair_with_limits, source_form_controls_for_sheet_with_limits,
};
use crate::error::{Error, Result, invalid};
use crate::raw;
use crate::source_payload::SourcePayload;
use crate::workbook::source::validate_sheet_graph;
use crate::{Selector, WorksheetKind};

const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";
const EXCEL_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:office:excel";

/// Immutable owner state captured before a scalar edit.
#[derive(Clone, Debug)]
pub struct FormControlSnapshot {
    sheet_name: Box<str>,
    sheet_position: usize,
    collection: FormControlCollection,
    limits: OwnerLimits,
    source_lineage: Option<litchi_opc::SourceLineage>,
    execution_context: Option<ExecutionContext>,
    generated_vml: Option<SourcePayload>,
}

impl FormControlSnapshot {
    fn new(
        name: String,
        position: usize,
        collection: FormControlCollection,
        limits: OwnerLimits,
        source_lineage: Option<litchi_opc::SourceLineage>,
        execution_context: Option<ExecutionContext>,
    ) -> Self {
        Self {
            sheet_name: name.into_boxed_str(),
            sheet_position: position,
            collection,
            limits,
            source_lineage,
            execution_context,
            generated_vml: None,
        }
    }
    /// Workbook-catalog sheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }
    /// Workbook-catalog zero-based sheet position.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.sheet_position
    }
    /// Immutable effective owner projection.
    #[must_use]
    pub const fn form_controls(&self) -> &FormControlCollection {
        &self.collection
    }
    /// Alias for [`Self::form_controls`].
    #[must_use]
    pub const fn collection(&self) -> &FormControlCollection {
        &self.collection
    }
    /// Resolve one control by an ordinary selector.
    pub fn control<'a>(
        &self,
        selector: impl Into<ControlSelector<'a>>,
    ) -> Result<Option<&FormControlView>> {
        self.collection.get(selector)
    }
    /// Pinned owner profile.
    #[must_use]
    pub const fn profile(&self) -> OwnerProfile {
        self.collection.profile()
    }
    /// Complete retained source read set, when source-backed.
    #[must_use]
    pub fn read_set(&self) -> Option<&super::FormControlReadSet> {
        self.collection.read_set()
    }
    /// Source version captured by this snapshot.
    #[must_use]
    pub fn source_version(&self) -> Option<litchi_core::SourceVersion> {
        self.collection.source_version()
    }
    /// Whether caller execution reservations remain attached.
    #[must_use]
    pub fn has_execution_budget(&self) -> bool {
        self.collection.has_execution_budget() || self.generated_vml.is_some()
    }
    fn same_source(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
            && self.sheet_position == other.sheet_position
            && self.limits == other.limits
            && self.source_lineage == other.source_lineage
            && self.collection.same_source(&other.collection)
    }
    fn with_properties(&self, position: usize, properties: Properties) -> Result<Self> {
        let collection = self
            .collection
            .with_replaced_properties(position, properties)
            .ok_or_else(|| invalid("form-control target disappeared during scalar edit"))?;
        Ok(Self {
            sheet_name: self.sheet_name.clone(),
            sheet_position: self.sheet_position,
            collection,
            limits: self.limits,
            source_lineage: self.source_lineage.clone(),
            execution_context: self.execution_context.clone(),
            generated_vml: self.generated_vml.clone(),
        })
    }

    fn retain_generated_vml(&mut self, payload: SourcePayload) {
        self.generated_vml = Some(payload);
    }
}

#[derive(Clone, Debug)]
pub(crate) enum OwnedSelector {
    Position(usize),
    Name(String),
}

impl OwnedSelector {
    pub(crate) fn selector(&self) -> ControlSelector<'_> {
        match self {
            Self::Position(position) => ControlSelector::Position(*position),
            Self::Name(name) => ControlSelector::Name(name),
        }
    }
}

pub(crate) fn owned_selector<'a>(selector: impl Into<ControlSelector<'a>>) -> OwnedSelector {
    match selector.into() {
        ControlSelector::Position(position) => OwnedSelector::Position(position),
        ControlSelector::Name(name) => OwnedSelector::Name(name.to_owned()),
    }
}

pub(crate) fn validate_scalar_operation_count(count: usize, limits: OwnerLimits) -> Result<()> {
    if count > limits.max_scalar_operations() {
        return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::Objects,
            observed: u64::try_from(count).unwrap_or(u64::MAX),
            limit: u64::try_from(limits.max_scalar_operations()).unwrap_or(u64::MAX),
            scope: Arc::from("form-control scalar operations"),
        }));
    }
    Ok(())
}

pub(crate) fn validate_changed_part_count(count: usize, limits: OwnerLimits) -> Result<()> {
    if count > limits.max_changed_parts() {
        return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::Objects,
            observed: u64::try_from(count).unwrap_or(u64::MAX),
            limit: u64::try_from(limits.max_changed_parts()).unwrap_or(u64::MAX),
            scope: Arc::from("form-control changed parts"),
        }));
    }
    Ok(())
}

pub(crate) fn validate_pair_input_limits(
    properties: &[u8],
    vml: &[u8],
    limits: OwnerLimits,
) -> Result<()> {
    let staging = properties
        .len()
        .checked_add(vml.len())
        .ok_or_else(|| invalid("form-control scalar staging size overflow"))?;
    if staging > limits.max_staging_bytes() {
        return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::Memory,
            observed: u64::try_from(staging).unwrap_or(u64::MAX),
            limit: u64::try_from(limits.max_staging_bytes()).unwrap_or(u64::MAX),
            scope: Arc::from("form-control scalar source staging"),
        }));
    }
    Ok(())
}

pub(crate) fn validate_pair_output_limits(
    before_properties: &[u8],
    before_vml: &[u8],
    after_properties: &[u8],
    after_vml: &[u8],
    limits: OwnerLimits,
) -> Result<()> {
    let output = after_properties
        .len()
        .checked_add(after_vml.len())
        .ok_or_else(|| invalid("form-control scalar output size overflow"))?;
    if output > limits.max_output_bytes() {
        return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::OutputBytes,
            observed: u64::try_from(output).unwrap_or(u64::MAX),
            limit: u64::try_from(limits.max_output_bytes()).unwrap_or(u64::MAX),
            scope: Arc::from("form-control scalar output"),
        }));
    }
    let staging = before_properties
        .len()
        .checked_add(before_vml.len())
        .and_then(|size| size.checked_add(output))
        .ok_or_else(|| invalid("form-control scalar staging size overflow"))?;
    if staging > limits.max_staging_bytes() {
        return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::Memory,
            observed: u64::try_from(staging).unwrap_or(u64::MAX),
            limit: u64::try_from(limits.max_staging_bytes()).unwrap_or(u64::MAX),
            scope: Arc::from("form-control scalar staging"),
        }));
    }
    Ok(())
}

#[derive(Clone, Debug)]
struct Operation {
    selector: OwnedSelector,
    field: ScalarField,
    value: Option<ScalarValue>,
}

/// One owned scalar operation staged by the ordinary worksheet transaction.
///
/// Physical package identities remain private to the owner helper.  The
/// transaction keeps only a selector, a typed field, and an owned value until
/// commit recomputes the complete paired properties/VML overlay.
#[derive(Clone, Debug)]
pub(crate) struct FormControlScalarAction {
    selector: OwnedSelector,
    field: ScalarField,
    value: Option<ScalarValue>,
}

impl FormControlScalarAction {
    pub(crate) fn new<'a>(
        selector: impl Into<ControlSelector<'a>>,
        field: ScalarField,
        value: Option<ScalarValue>,
    ) -> Self {
        Self {
            selector: owned_selector(selector),
            field,
            value,
        }
    }

    pub(crate) fn selector(&self) -> ControlSelector<'_> {
        self.selector.selector()
    }

    pub(crate) const fn field(&self) -> ScalarField {
        self.field
    }

    pub(crate) fn value(&self) -> Option<ScalarValue> {
        self.value.clone()
    }
}

/// Isolated source-backed edit for one worksheet.
pub struct FormControlSourceEdit {
    before: FormControlSnapshot,
    operations: Vec<Operation>,
}
impl FormControlSourceEdit {
    fn new(before: FormControlSnapshot) -> Self {
        Self {
            before,
            operations: Vec::new(),
        }
    }
    /// Exact owner projection captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &FormControlSnapshot {
        &self.before
    }
    /// Stage one scalar by position or exact authored name.
    pub fn set_scalar<'a>(
        &mut self,
        selector: impl Into<ControlSelector<'a>>,
        field: ScalarField,
        value: Option<ScalarValue>,
    ) -> Result<&mut Self> {
        let operation_count = self
            .operations
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("form-control scalar operation count overflow"))?;
        validate_scalar_operation_count(operation_count, self.before.limits)?;
        let selector = owned_selector(selector);
        let control = self
            .before
            .collection
            .get(selector.selector())?
            .ok_or_else(|| invalid("form-control selector did not resolve"))?;
        validate_scalar(control, field, value.as_ref())?;
        self.operations.push(Operation {
            selector,
            field,
            value,
        });
        Ok(self)
    }
    /// Selector-first alias matching the worksheet edit facade.
    pub fn set_form_control_scalar<'a>(
        &mut self,
        selector: impl Into<ControlSelector<'a>>,
        field: ScalarField,
        value: Option<ScalarValue>,
    ) -> Result<&mut Self> {
        self.set_scalar(selector, field, value)
    }
    /// Whether any scalar operation is staged.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.operations.is_empty()
    }
    /// Validate and freeze the paired properties/VML transaction.
    pub fn commit(self) -> Result<FormControlCommit> {
        let mut current = self.before.clone();
        let mut overlay: Option<Overlay> = None;
        let mut target_position = None;
        for operation in self.operations {
            let control = current
                .collection
                .get(operation.selector.selector())?
                .ok_or_else(|| invalid("form-control selector did not resolve during commit"))?;
            let position = control.position();
            if let Some(previous) = target_position {
                if previous != position {
                    return Err(invalid(
                        "form-control scalar batch targets multiple controls",
                    ));
                }
            } else {
                target_position = Some(position);
            }
            validate_scalar(control, operation.field, operation.value.as_ref())?;
            let property_source_payload = control
                .source_property_payload()
                .ok_or_else(|| invalid("form-control properties source bytes were not retained"))?;
            let property_source = property_source_payload.as_bytes();
            let property_uri = control.property_part_name().cloned().ok_or_else(|| {
                invalid("form-control properties source identity was not retained")
            })?;
            let vml_uri = control
                .vml_part_name()
                .cloned()
                .ok_or_else(|| invalid("form-control VML source identity was not retained"))?;
            let read_set = current
                .collection
                .read_set()
                .ok_or_else(|| invalid("form-control source read set was not retained"))?;
            let vml_source_payload = if let Some(existing) = overlay
                .as_ref()
                .filter(|existing| existing.position == position)
            {
                existing.after_vml.clone()
            } else {
                read_set
                    .vml()
                    .ok_or_else(|| invalid("form-control VML source was not retained"))?
                    .source_payload()
            };
            let vml_source = vml_source_payload.as_bytes();
            validate_pair_input_limits(property_source, vml_source, self.before.limits)?;
            let client_data = find_client_data(vml_source, control.shape().vml_id())?;
            let properties = super::inspect(property_source)?;
            let pair = replace_scalar_pair_with_limits(
                &properties,
                client_data,
                operation.field,
                operation.value.clone(),
                mirror_limits(self.before.limits),
                self.before.execution_context.as_ref(),
            )
            .map_err(mirror_error)?;
            if !pair.changed() {
                continue;
            }
            validate_changed_part_count(2, self.before.limits)?;
            let (properties_after, vml_after, pair_budget) = pair.into_parts_with_budget();
            validate_pair_output_limits(
                property_source,
                vml_source,
                &properties_after,
                &vml_after,
                self.before.limits,
            )?;
            let properties_after = Arc::new(properties_after);
            let projection_budget = reserve_generated(
                self.before.execution_context.as_ref(),
                generated_projection_storage_bytes(
                    current.collection.len(),
                    current.sheet_name.len(),
                    current
                        .collection
                        .generated_clone_uri_storage_bytes()
                        .ok_or_else(|| invalid("form-control generated URI storage overflow"))?,
                )?,
                current.collection.len().saturating_add(1),
                "form-control generated collection",
            )?;
            let properties_payload =
                SourcePayload::owned_budgeted(Arc::clone(&properties_after), pair_budget.clone());
            let generated_properties = super::codec::parse_source_with_limits_and_context(
                properties_payload.clone(),
                &super::owner::leaf_limits(&self.before.limits),
                self.before.execution_context.as_ref(),
            )?;
            generated_properties
                .validate_retained_with_limits(super::owner::leaf_limits(&self.before.limits))?;
            current = current.with_properties(position, generated_properties)?;
            let vml_after = Arc::new(vml_after);
            let vml_payload = SourcePayload::owned_budgeted(vml_after, pair_budget);
            current.collection = current.collection.with_generated_budget(projection_budget);
            current.retain_generated_vml(vml_payload.clone());
            match overlay.as_mut() {
                Some(existing) if existing.position == position => {
                    existing.after_properties = properties_payload;
                    existing.after_vml = vml_payload;
                },
                Some(_) => {
                    return Err(invalid(
                        "form-control scalar batch targets multiple controls",
                    ));
                },
                None => {
                    overlay = Some(Overlay {
                        position,
                        property_uri,
                        vml_uri,
                        before_properties: property_source_payload,
                        after_properties: properties_payload,
                        before_vml: vml_source_payload,
                        after_vml: vml_payload,
                    })
                },
            }
        }
        let Some(overlay) = overlay else {
            let patch = FormControlPatch::empty(self.before.clone());
            return Ok(FormControlCommit::new(self.before, patch, false));
        };
        let patch = FormControlPatch::from_overlay(self.before, current, overlay);
        Ok(FormControlCommit::new(patch.after.clone(), patch, true))
    }
}

/// Exact paired source overlays planned for an ordinary in-memory workbook
/// transaction.  The workbook edit owner turns these into its private
/// `PartChange` records; this type deliberately carries no OPC handles.
#[derive(Clone, Debug)]
pub(crate) struct OrdinaryFormControlOverlay {
    /// Zero-based control position within the selected worksheet owner.
    ///
    /// The ordinary workbook transaction uses this semantic position for its
    /// conflict/readback record; package identities remain private to this
    /// owner module.
    pub(crate) position: usize,
    pub(crate) property_uri: PackURI,
    pub(crate) before_properties: Arc<Vec<u8>>,
    pub(crate) after_properties: Arc<Vec<u8>>,
    pub(crate) vml_uri: PackURI,
    pub(crate) before_vml: Arc<Vec<u8>>,
    pub(crate) after_vml: Arc<Vec<u8>>,
    /// Complete semantic owner projection expected after both source members
    /// are installed. The ordinary transaction uses this for final paired
    /// owner readback after all other edits.
    pub(crate) expected_collection: FormControlCollection,
}

/// Validate one selector against the eager owner used by `WorksheetEdit`.
pub(crate) fn validate_ordinary_form_control_selector(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    selector: ControlSelector<'_>,
) -> Result<()> {
    let collection = super::eager_form_controls_for_sheet_with_limits(
        package,
        worksheet_uri,
        OwnerLimits::from_read_limits(package.read_limits()),
    )
    .map_err(super::owner_to_xlsx)?;
    if collection.get(selector)?.is_none() {
        return Err(invalid("form-control selector did not resolve"));
    }
    Ok(())
}

/// Plan all staged ordinary scalar operations as one paired properties/VML
/// write set.  The operation list is replayed from the original source bytes
/// so multiple fields on one control compose without dropping the first
/// overlay.  A batch that targets more than one control is refused until the
/// host has a proven multi-control transaction shape.
pub(crate) fn ordinary_form_control_overlays(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    actions: &[FormControlScalarAction],
) -> Result<Vec<OrdinaryFormControlOverlay>> {
    if actions.is_empty() {
        return Ok(Vec::new());
    }
    let limits = OwnerLimits::from_read_limits(package.read_limits());
    validate_scalar_operation_count(actions.len(), limits)?;
    let collection =
        super::eager_form_controls_for_sheet_with_limits(package, worksheet_uri, limits)
            .map_err(super::owner_to_xlsx)?;
    let mut target_position = None;
    let mut property_uri = None;
    let mut vml_uri = None;
    let mut property_source = None;
    let mut vml_source = None;
    let mut property_after = None;
    let mut vml_after = None;

    for action in actions {
        let control = collection
            .get(action.selector())?
            .ok_or_else(|| invalid("form-control selector did not resolve"))?;
        let position = control.position();
        if let Some(previous) = target_position {
            if previous != position {
                return Err(invalid(
                    "form-control scalar batch targets multiple controls",
                ));
            }
        } else {
            target_position = Some(position);
            let property_name = control.property_part_name().cloned().ok_or_else(|| {
                invalid("form-control properties source identity was not retained")
            })?;
            let vml_name = control
                .vml_part_name()
                .cloned()
                .ok_or_else(|| invalid("form-control VML source identity was not retained"))?;
            let properties = package.get_part(&property_name)?.blob();
            let vml = package.get_part(&vml_name)?.blob();
            validate_pair_input_limits(properties, vml, limits)?;
            property_uri = Some(property_name);
            vml_uri = Some(vml_name);
            property_source = Some(properties.to_vec());
            vml_source = Some(vml.to_vec());
            property_after = Some(properties.to_vec());
            vml_after = Some(vml.to_vec());
        }

        let property_bytes = property_after
            .as_deref()
            .ok_or_else(|| invalid("form-control properties source disappeared"))?;
        let vml_bytes = vml_after
            .as_deref()
            .ok_or_else(|| invalid("form-control VML source disappeared"))?;
        validate_pair_input_limits(property_bytes, vml_bytes, limits)?;
        let properties = super::inspect(property_bytes)?;
        let client_data = find_client_data(vml_bytes, control.shape().vml_id())?;
        let pair = replace_scalar_pair_with_limits(
            &properties,
            client_data,
            action.field(),
            action.value(),
            mirror_limits(limits),
            None,
        )
        .map_err(mirror_error)?;
        if pair.changed() {
            validate_changed_part_count(2, limits)?;
            let (properties, vml) = pair.into_parts();
            validate_pair_output_limits(property_bytes, vml_bytes, &properties, &vml, limits)?;
            property_after = Some(properties);
            vml_after = Some(vml);
        }
    }

    let property_source = property_source
        .ok_or_else(|| invalid("form-control properties source was not retained"))?;
    let vml_source =
        vml_source.ok_or_else(|| invalid("form-control VML source was not retained"))?;
    let property_after =
        property_after.ok_or_else(|| invalid("form-control properties output was not retained"))?;
    let vml_after = vml_after.ok_or_else(|| invalid("form-control VML output was not retained"))?;
    if property_after == property_source && vml_after == vml_source {
        return Ok(Vec::new());
    }
    let position =
        target_position.ok_or_else(|| invalid("form-control target position was not retained"))?;
    let expected_collection = collection
        .with_replaced_properties(position, super::parse(&property_after)?)
        .ok_or_else(|| invalid("form-control target disappeared during ordinary edit"))?;
    Ok(vec![OrdinaryFormControlOverlay {
        position,
        property_uri: property_uri
            .ok_or_else(|| invalid("form-control properties identity was not retained"))?,
        before_properties: Arc::new(property_source),
        after_properties: Arc::new(property_after),
        vml_uri: vml_uri.ok_or_else(|| invalid("form-control VML identity was not retained"))?,
        before_vml: Arc::new(vml_source),
        after_vml: Arc::new(vml_after),
        expected_collection,
    }])
}

/// Re-scan one ordinary candidate after all worksheet/package changes have
/// been staged. This validates the properties and VML pair through the same
/// owner graph used for reads and checks every effective control so an
/// unrelated candidate transition cannot be hidden by a leaf-only assertion.
pub(crate) fn validate_ordinary_form_control_candidate(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    overlay: &OrdinaryFormControlOverlay,
) -> Result<()> {
    let actual = super::eager_form_controls_for_sheet_with_limits(
        package,
        worksheet_uri,
        OwnerLimits::from_read_limits(package.read_limits()),
    )
    .map_err(super::owner_to_xlsx)?;
    if actual != overlay.expected_collection {
        return Err(invalid(
            "ordinary form-control paired owner readback did not match its target",
        ));
    }
    Ok(())
}

#[derive(Clone, Debug)]
struct Overlay {
    position: usize,
    property_uri: PackURI,
    vml_uri: PackURI,
    before_properties: SourcePayload,
    after_properties: SourcePayload,
    before_vml: SourcePayload,
    after_vml: SourcePayload,
}

/// Exact reversible paired properties/VML patch.
#[derive(Clone, Debug)]
pub struct FormControlPatch {
    before: FormControlSnapshot,
    after: FormControlSnapshot,
    property_uri: Option<PackURI>,
    vml_uri: Option<PackURI>,
    before_properties: Option<SourcePayload>,
    after_properties: Option<SourcePayload>,
    before_vml: Option<SourcePayload>,
    after_vml: Option<SourcePayload>,
}
impl FormControlPatch {
    fn empty(snapshot: FormControlSnapshot) -> Self {
        Self {
            before: snapshot.clone(),
            after: snapshot,
            property_uri: None,
            vml_uri: None,
            before_properties: None,
            after_properties: None,
            before_vml: None,
            after_vml: None,
        }
    }
    fn from_overlay(
        before: FormControlSnapshot,
        after: FormControlSnapshot,
        overlay: Overlay,
    ) -> Self {
        Self {
            before,
            after,
            property_uri: Some(overlay.property_uri),
            vml_uri: Some(overlay.vml_uri),
            before_properties: Some(overlay.before_properties),
            after_properties: Some(overlay.after_properties),
            before_vml: Some(overlay.before_vml),
            after_vml: Some(overlay.after_vml),
        }
    }
    /// Required exact source state.
    #[must_use]
    pub const fn before(&self) -> &FormControlSnapshot {
        &self.before
    }
    /// Exact target semantic state.
    #[must_use]
    pub const fn after(&self) -> &FormControlSnapshot {
        &self.after
    }
    /// Whether both paired source members are unchanged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        match (
            &self.before_properties,
            &self.after_properties,
            &self.before_vml,
            &self.after_vml,
        ) {
            (Some(a), Some(b), Some(c), Some(d)) => {
                a.as_bytes() == b.as_bytes() && c.as_bytes() == d.as_bytes()
            },
            _ => true,
        }
    }
    /// Swap exact before/after source bytes and semantic projections.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            property_uri: self.property_uri.clone(),
            vml_uri: self.vml_uri.clone(),
            before_properties: self.after_properties.clone(),
            after_properties: self.before_properties.clone(),
            before_vml: self.after_vml.clone(),
            after_vml: self.before_vml.clone(),
        }
    }
    /// Apply after checking both source members and eager owner readback.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        if self.is_empty() {
            if let Some(read_set) = self.before.collection.read_set() {
                if !super::read_set_matches_package_except(package, read_set, &[])? {
                    return Err(Error::PatchConflict {
                        part: read_set.worksheet_part_name().map_or_else(
                            || "form-control read set".to_owned(),
                            ToString::to_string,
                        ),
                    });
                }
            }
            return Ok(());
        }
        let property_uri = self
            .property_uri
            .as_ref()
            .ok_or_else(|| invalid("form-control patch has no properties identity"))?;
        let vml_uri = self
            .vml_uri
            .as_ref()
            .ok_or_else(|| invalid("form-control patch has no VML identity"))?;
        let before_properties = self
            .before_properties
            .as_ref()
            .ok_or_else(|| invalid("form-control patch has no properties source"))?;
        let before_vml = self
            .before_vml
            .as_ref()
            .ok_or_else(|| invalid("form-control patch has no VML source"))?;
        validate_changed_part_count(2, self.before.limits)?;
        validate_pair_input_limits(
            before_properties.as_bytes(),
            before_vml.as_bytes(),
            self.before.limits,
        )?;
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let property_current = match package.get_part(property_uri) {
            Ok(part) => part,
            Err(litchi_opc::OpcError::PartNotFound(_)) => {
                return Err(Error::PatchConflict {
                    part: property_uri.to_string(),
                });
            },
            Err(error) => return Err(error.into()),
        };
        if property_current.blob() != before_properties.as_bytes() {
            return Err(Error::PatchConflict {
                part: property_uri.to_string(),
            });
        }
        let vml_current = match package.get_part(vml_uri) {
            Ok(part) => part,
            Err(litchi_opc::OpcError::PartNotFound(_)) => {
                return Err(Error::PatchConflict {
                    part: vml_uri.to_string(),
                });
            },
            Err(error) => return Err(error.into()),
        };
        if vml_current.blob() != before_vml.as_bytes() {
            return Err(Error::PatchConflict {
                part: vml_uri.to_string(),
            });
        }
        let read_set = self
            .before
            .collection
            .read_set()
            .ok_or_else(|| invalid("form-control patch has no source read set"))?;
        if !super::read_set_matches_package_except(package, read_set, &[property_uri, vml_uri])? {
            return Err(Error::PatchConflict {
                part: read_set
                    .worksheet_part_name()
                    .map_or_else(|| "form-control read set".to_owned(), ToString::to_string),
            });
        }
        let worksheet_uri = read_set
            .worksheet_part_name()
            .cloned()
            .ok_or_else(|| invalid("form-control patch has no worksheet identity"))?;
        let after_properties = self
            .after_properties
            .as_ref()
            .ok_or_else(|| invalid("form-control patch has no properties target"))?;
        let after_vml = self
            .after_vml
            .as_ref()
            .ok_or_else(|| invalid("form-control patch has no VML target"))?;
        validate_pair_output_limits(
            before_properties.as_bytes(),
            before_vml.as_bytes(),
            after_properties.as_bytes(),
            after_vml.as_bytes(),
            self.before.limits,
        )?;
        let after_properties_bytes = after_properties.materialized_arc(
            self.before.limits.max_output_bytes(),
            "form-control eager properties publication",
        )?;
        let after_vml_bytes = after_vml.materialized_arc(
            self.before.limits.max_output_bytes(),
            "form-control eager VML publication",
        )?;
        let mut candidate = package.clone();
        // Keep the generated x14 payload source-authorized when this eager
        // patch is later serialized.  A direct `set_blob_shared` revokes the
        // package's retained XML provenance, so noncompact source formatting
        // would fail at the writer boundary even though this edit is a
        // source-preserving scalar splice.  The byte-oriented replacement
        // API checks the stale source and installs the replacement proof on
        // the candidate before either paired member is committed.
        candidate.try_replace_owned_xml_part_bytes(
            property_uri,
            before_properties.as_bytes(),
            after_properties_bytes,
        )?;
        candidate
            .get_part_mut(vml_uri)?
            .set_blob_shared(after_vml_bytes);
        let resulting = super::eager_form_controls_for_sheet_with_limits(
            &candidate,
            &worksheet_uri,
            self.before.limits,
        )
        .map_err(super::owner_to_xlsx)?;
        if resulting != self.after.collection {
            return Err(invalid(
                "form-control patch paired owner readback did not match its target",
            ));
        }
        *package = candidate;
        Ok(())
    }
}

/// Successful paired scalar transaction.
#[derive(Debug)]
pub struct FormControlCommit {
    snapshot: FormControlSnapshot,
    patch: FormControlPatch,
    changed: bool,
}
impl FormControlCommit {
    fn new(snapshot: FormControlSnapshot, patch: FormControlPatch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }
    /// Whether either paired source member changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
    /// Resulting immutable projection.
    #[must_use]
    pub const fn snapshot(&self) -> &FormControlSnapshot {
        &self.snapshot
    }
    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &FormControlPatch {
        &self.patch
    }
    /// Consume into snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (FormControlSnapshot, FormControlPatch) {
        (self.snapshot, self.patch)
    }
}

/// Source-backed form-control scalar editor.
pub struct SourceBackedFormControlEditor {
    package: SourceBackedPackage,
    limits: OwnerLimits,
}
impl SourceBackedFormControlEditor {
    /// Open with standard bounded policies.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }
    /// Open with explicit OPC read limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }
    /// Open with an explicit source cache policy.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_cache_limits(
            source, limits,
        )?)
    }
    /// Open with explicit read and cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                read_limits,
                cache_limits,
            )?,
        )
    }
    /// Open with a caller-owned execution context.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_execution_context(
            source, limits, context,
        )?)
    }
    /// Open from an already validated deferred package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        if package.has_encrypted_entries() {
            return Err(invalid(
                "encrypted XLSX form-control source is not admitted",
            ));
        }
        Ok(Self {
            limits: OwnerLimits::from_read_limits(package.read_limits()),
            package,
        })
    }
    /// Lower the owner policy.
    #[must_use]
    pub fn with_limits(mut self, limits: OwnerLimits) -> Self {
        self.limits = limits;
        self
    }
    /// Capture one worksheet's source-backed owner projection.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<FormControlSnapshot> {
        self.package.check_execution()?;
        let (position, name, uri) = source_sheet(&self.package, selector.into())?;
        let collection =
            source_form_controls_for_sheet_with_limits(&self.package, &uri, self.limits)?;
        Ok(FormControlSnapshot::new(
            name,
            position,
            collection,
            self.limits,
            Some(self.package.source_lineage()),
            self.package.execution_context(),
        ))
    }
    /// Start an isolated source-backed edit.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<FormControlSourceEdit> {
        Ok(FormControlSourceEdit::new(self.snapshot(selector)?))
    }
    /// Publish while raw-copying every unrelated physical ZIP member.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        mut writer: W,
        commit: &FormControlCommit,
    ) -> Result<FormControlSnapshot> {
        self.package.check_execution()?;
        let current = self.snapshot(commit.patch.before.sheet_position)?;
        if !current.same_source(&commit.patch.before) {
            return Err(Error::PatchConflict {
                part: current.sheet_name().to_owned(),
            });
        }
        if commit.patch.is_empty() {
            self.package
                .write_part_overlays_shared_to_stream(&mut writer, Vec::new())?;
            return Ok(current);
        }
        let property_uri = commit
            .patch
            .property_uri
            .clone()
            .ok_or_else(|| invalid("form-control commit has no properties identity"))?;
        let vml_uri = commit
            .patch
            .vml_uri
            .clone()
            .ok_or_else(|| invalid("form-control commit has no VML identity"))?;
        let properties_before = commit
            .patch
            .before_properties
            .as_ref()
            .ok_or_else(|| invalid("form-control commit has no properties source"))?;
        let properties_after = commit
            .patch
            .after_properties
            .as_ref()
            .ok_or_else(|| invalid("form-control commit has no properties output"))?;
        let vml_after = commit
            .patch
            .after_vml
            .as_ref()
            .ok_or_else(|| invalid("form-control commit has no VML output"))?;
        let properties_source = source_xml_splice(
            &self.package,
            &property_uri,
            properties_before.as_bytes(),
            properties_after.as_bytes(),
        )?;
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_source_xml_part(property_uri, properties_source)?;
        // VML is an XML-shaped legacy part whose OPC content type is not
        // classified as XML by the generic topology publisher.  Keep its
        // complete mirror output in the same atomic topology plan; the
        // source package still supplies the stale-byte and source-version
        // guards above, while the changed payload is copied as an opaque
        // legacy part.
        plan.try_replace_part(vml_uri, vml_after.as_bytes().to_vec())?;
        let worksheet_uri = commit
            .patch
            .before
            .collection
            .read_set()
            .and_then(|read_set| read_set.worksheet_part_name())
            .cloned()
            .ok_or_else(|| invalid("form-control commit has no worksheet identity"))?;
        let prepared = self.package.prepare_topology(plan)?;
        // `PreparedTopology::with_candidate` deliberately exposes the OPC
        // error type at its callback boundary.  Keep the owner readback
        // error in this XLSX layer instead of stringifying it into an OPC
        // error merely to cross that boundary; publication remains blocked
        // whenever the candidate cannot be re-read or does not match.
        let mut readback_error = None;
        prepared
            .with_candidate(|candidate| {
                match super::eager_form_controls_for_effective_topology(
                    candidate,
                    &worksheet_uri,
                    commit.patch.before.limits,
                ) {
                    Ok(resulting) => {
                        if resulting != commit.patch.after.collection {
                            readback_error = Some(invalid(
                                "form-control source publication paired owner readback did not match its target",
                            ));
                        }
                    }
                    Err(error) => {
                        readback_error = Some(super::owner_to_xlsx(error));
                    }
                }
                Ok(())
            })
            .map_err(Error::Package)?;
        if let Some(error) = readback_error {
            return Err(error);
        }
        prepared.publish_to_stream(&mut writer)?;
        Ok(commit.snapshot.clone())
    }
    /// Content-free source cache diagnostics.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }
}

fn source_xml_splice(
    package: &SourceBackedPackage,
    part_name: &PackURI,
    before: &[u8],
    after: &[u8],
) -> Result<SourceXmlPart> {
    let source = package.part(part_name)?.source_xml()?;
    if source.bytes() != before {
        return Err(Error::PatchConflict {
            part: part_name.to_string(),
        });
    }
    if before == after {
        return Ok(source);
    }
    let mut edits = text_splice_ranges(before, after);
    if edits.is_empty() {
        let (source_range, replacement_range) = minimal_diff(before, after)?;
        let replacement = after
            .get(replacement_range)
            .ok_or_else(|| invalid("form-control replacement XML range disappeared"))?;
        if replacement.contains(&b'<') {
            return Err(Error::Unsupported {
                feature: "form-control scalar source splice is not a direct text replacement",
            });
        }
        edits.push((source_range, replacement.to_vec()));
    }
    let mut proofs = Vec::new();
    for (range, replacement) in &edits {
        if replacement.contains(&b'<') {
            return Err(Error::Unsupported {
                feature: "form-control scalar source splice is not a direct text replacement",
            });
        }
        let expected = before
            .get(range.clone())
            .ok_or_else(|| invalid("form-control source XML splice range disappeared"))?;
        let proof = source.checked_range(range.clone(), expected)?;
        proofs.push((proof, replacement.clone()));
    }
    let mut publication = source.into_publication()?;
    for (proof, replacement) in proofs {
        let fragment = if replacement.is_empty() {
            AuthoredXmlFragment::empty()
        } else {
            AuthoredXmlFragment::text(replacement)?
        };
        publication.replace(proof, fragment)?;
    }
    Ok(publication.finish()?)
}

fn mirror_limits(owner: OwnerLimits) -> MirrorLimits {
    MirrorLimits::new()
        .with_max_source_bytes(owner.max_vml_bytes())
        .with_max_output_bytes(owner.max_vml_bytes().min(owner.max_output_bytes()))
        .with_max_depth(owner.max_mce_depth())
        .with_max_events(owner.max_mce_events())
        .with_max_fields(owner.max_mirror_nodes())
        .with_max_value_bytes(owner.max_mce_bytes())
        .with_max_attributes(owner.max_mirror_nodes())
}

fn generated_projection_storage_bytes(
    control_count: usize,
    sheet_name_bytes: usize,
    uri_storage_bytes: usize,
) -> Result<usize> {
    let controls = control_count
        .checked_mul(size_of::<FormControlView>())
        .ok_or_else(|| invalid("form-control generated collection storage overflow"))?;
    let controls_peak = controls
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(ARC_HEADER_BYTES))
        .ok_or_else(|| invalid("form-control generated collection storage overflow"))?;
    let sheet_name = sheet_name_bytes
        .checked_add(size_of::<usize>())
        .and_then(|bytes| bytes.checked_mul(2))
        .ok_or_else(|| invalid("form-control generated sheet-name storage overflow"))?;
    controls_peak
        .checked_add(uri_storage_bytes)
        .and_then(|bytes| bytes.checked_add(sheet_name))
        .ok_or_else(|| invalid("form-control generated snapshot storage overflow"))
}

fn text_splice_ranges(before: &[u8], after: &[u8]) -> Vec<(Range<usize>, Vec<u8>)> {
    let before_ranges = xml_text_ranges(before);
    let after_ranges = xml_text_ranges(after);
    if before_ranges.len() != after_ranges.len() {
        return Vec::new();
    }
    let mut edits = Vec::new();
    for ((before_range, before_text), (after_range, after_text)) in
        before_ranges.iter().zip(after_ranges.iter())
    {
        if before_text != after_text {
            edits.push((before_range.clone(), after[after_range.clone()].to_vec()));
        }
    }
    edits
}

fn xml_text_ranges(source: &[u8]) -> Vec<(Range<usize>, Vec<u8>)> {
    let mut reader = Reader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut ranges = Vec::new();
    loop {
        let Some(start) = origin.offset(reader.buffer_position()) else {
            return Vec::new();
        };
        let Ok(event) = reader.read_event_into(&mut buffer) else {
            return Vec::new();
        };
        let Some(end) = origin.offset(reader.buffer_position()) else {
            return Vec::new();
        };
        if let Event::Text(_) = event {
            if let Some(value) = source.get(start..end) {
                ranges.push((start..end, value.to_vec()));
            } else {
                return Vec::new();
            }
        } else if matches!(event, Event::Eof) {
            break;
        }
        buffer.clear();
    }
    ranges
}

fn minimal_diff(before: &[u8], after: &[u8]) -> Result<(Range<usize>, Range<usize>)> {
    let prefix = before
        .iter()
        .zip(after.iter())
        .take_while(|(left, right)| left == right)
        .count();
    let mut suffix = 0usize;
    while suffix < before.len().saturating_sub(prefix)
        && suffix < after.len().saturating_sub(prefix)
        && before[before.len() - suffix - 1] == after[after.len() - suffix - 1]
    {
        suffix += 1;
    }
    let before_end = before
        .len()
        .checked_sub(suffix)
        .ok_or_else(|| invalid("form-control source diff underflow"))?;
    let after_end = after
        .len()
        .checked_sub(suffix)
        .ok_or_else(|| invalid("form-control replacement diff underflow"))?;
    Ok((prefix..before_end, prefix..after_end))
}

fn source_sheet(
    package: &SourceBackedPackage,
    selector: Selector<'_>,
) -> Result<(usize, String, PackURI)> {
    let workbook = package.main_document_part()?;
    let data = workbook.data()?;
    let catalog = raw::parse_catalog(data.as_bytes())?;
    let parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
    let position = match selector {
        litchi_core::Selector::Position(position) => position.get(),
        litchi_core::Selector::Name(name) => catalog
            .sheets
            .iter()
            .position(|sheet| sheet.name.eq_ignore_ascii_case(name.as_ref()))
            .ok_or_else(|| invalid("worksheet selector did not resolve"))?,
        litchi_core::Selector::Id(_) => return Err(Error::UnsupportedSelector),
        _ => return Err(Error::UnsupportedSelector),
    };
    let sheet = catalog
        .sheets
        .get(position)
        .ok_or_else(|| invalid("worksheet selector did not resolve"))?;
    let part = parts
        .get(position)
        .ok_or_else(|| invalid("worksheet graph position is absent"))?;
    if part.kind != WorksheetKind::Worksheet {
        return Err(Error::NotWorksheet {
            sheet: sheet.name.clone(),
        });
    }
    Ok((position, sheet.name.clone(), part.uri.clone()))
}

fn validate_scalar(
    control: &FormControlView,
    field: ScalarField,
    value: Option<&ScalarValue>,
) -> Result<()> {
    let source = control
        .source_property_bytes()
        .ok_or_else(|| invalid("form-control properties source bytes were not retained"))?;
    let view = super::inspect(source)?;
    // Preserve an exact source-only formula (or any other unknown source
    // scalar) when the requested value is the same typed projection.  The
    // general codec writer intentionally refuses to author source-only
    // formulas, but an exact no-op must still be stageable so publishing can
    // copy the original bytes unchanged.
    if view.properties().scalar(field).as_ref() == value {
        return Ok(());
    }
    let _ = view.replace_scalar(field, value.cloned())?;
    Ok(())
}

fn mirror_error(error: super::mirror::MirrorError) -> Error {
    match error {
        super::mirror::MirrorError::Leaf(error) => Error::FormControl(error),
        super::mirror::MirrorError::Execution(error) => {
            Error::Package(litchi_opc::OpcError::Execution(error))
        },
        super::mirror::MirrorError::Limit {
            resource,
            observed,
            maximum,
        } => Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: Resource::Memory,
            observed: observed as u64,
            limit: maximum as u64,
            scope: Arc::from(format!("form-control mirror {resource}")),
        }),
        super::mirror::MirrorError::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
        super::mirror::MirrorError::Invalid(message) => Error::Invalid(message),
        super::mirror::MirrorError::Unsupported { reason, .. } => {
            Error::Unsupported { feature: reason }
        },
        super::mirror::MirrorError::MissingSource { .. } => Error::Unsupported {
            feature: "form-control scalar mirror is missing",
        },
        super::mirror::MirrorError::Ambiguous { .. } => Error::Unsupported {
            feature: "form-control scalar mirror is ambiguous",
        },
        super::mirror::MirrorError::Disagreement { .. } => Error::Unsupported {
            feature: "form-control scalar mirror disagrees",
        },
    }
}

fn find_client_data<'a>(source: &'a [u8], vml_id: &str) -> Result<ClientDataSource<'a>> {
    let mut reader = NsReader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut target_shape: Option<Vec<u8>> = None;
    let mut target_shape_seen = false;
    let mut client: Option<(usize, Vec<u8>, Vec<u8>)> = None;
    let mut found_client: Option<(usize, usize, Vec<u8>)> = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("form-control XML offset exceeds usize"))?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        let namespace = finder_namespace(resolved)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("form-control XML offset exceeds usize"))?;
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                let name = element.name().as_ref().to_vec();
                let depth = stack.len();
                if depth == 1 && namespace == VML_NAMESPACE && local.as_ref() == b"shape" {
                    if shape_id_matches(&reader, &element, vml_id)? {
                        if target_shape_seen {
                            return Err(invalid("form-control VML shape id is ambiguous"));
                        }
                        target_shape_seen = true;
                        target_shape = Some(name.clone());
                    }
                } else if target_shape.is_some()
                    && depth == 2
                    && namespace == EXCEL_NAMESPACE
                    && local.as_ref() == b"ClientData"
                {
                    if client.is_some() || found_client.is_some() {
                        return Err(invalid(
                            "form-control VML shape has duplicate Excel ClientData elements",
                        ));
                    }
                    let prefix = element
                        .name()
                        .prefix()
                        .map_or_else(Vec::new, |prefix| prefix.as_ref().to_vec());
                    client = Some((start, name.clone(), prefix));
                }
                stack.push(name);
            },
            Event::Empty(element) => {
                let local = element.local_name();
                let depth = stack.len();
                if depth == 1 && namespace == VML_NAMESPACE && local.as_ref() == b"shape" {
                    if shape_id_matches(&reader, &element, vml_id)? {
                        if target_shape_seen {
                            return Err(invalid("form-control VML shape id is ambiguous"));
                        }
                        return Err(invalid(
                            "form-control VML target shape has no ClientData child",
                        ));
                    }
                } else if target_shape.is_some()
                    && depth == 2
                    && namespace == EXCEL_NAMESPACE
                    && local.as_ref() == b"ClientData"
                {
                    if client.is_some() || found_client.is_some() {
                        return Err(invalid(
                            "form-control VML shape has duplicate Excel ClientData elements",
                        ));
                    }
                    let prefix = element
                        .name()
                        .prefix()
                        .map_or_else(Vec::new, |prefix| prefix.as_ref().to_vec());
                    found_client = Some((start, end, prefix));
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                let name = element.name().as_ref().to_vec();
                let depth = stack.len();
                let expected = stack
                    .pop()
                    .ok_or_else(|| invalid("form-control VML closing element is unmatched"))?;
                if expected.as_slice() != name.as_slice() {
                    return Err(invalid(
                        "form-control VML closing QName does not match its opening QName",
                    ));
                }
                if let Some((client_start, client_name, prefix)) = client.take() {
                    if depth == 3
                        && client_name.as_slice() == name.as_slice()
                        && namespace == EXCEL_NAMESPACE
                        && local.as_ref() == b"ClientData"
                    {
                        if found_client.is_some() {
                            return Err(invalid(
                                "form-control VML shape has duplicate Excel ClientData elements",
                            ));
                        }
                        found_client = Some((client_start, end, prefix));
                    } else {
                        client = Some((client_start, client_name, prefix));
                    }
                }
                if target_shape.as_deref() == Some(name.as_slice()) && depth == 2 {
                    target_shape = None;
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    let (client_start, client_end, prefix) =
        found_client.ok_or_else(|| invalid("form-control VML ClientData range was not found"))?;
    ClientDataSource::new(source, client_start..client_end, &prefix).map_err(mirror_error)
}

fn finder_namespace(resolved: ResolveResult<'_>) -> Result<&'static [u8]> {
    match resolved {
        ResolveResult::Unbound => Ok(&[]),
        ResolveResult::Bound(namespace) if namespace.as_ref() == VML_NAMESPACE => Ok(VML_NAMESPACE),
        ResolveResult::Bound(namespace) if namespace.as_ref() == EXCEL_NAMESPACE => {
            Ok(EXCEL_NAMESPACE)
        },
        ResolveResult::Bound(_) => Ok(&[]),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "form-control VML element uses an unbound namespace prefix {}",
            String::from_utf8_lossy(&prefix)
        ))),
    }
}

fn shape_id_matches(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    expected: &str,
) -> Result<bool> {
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unknown(_)) {
            return Err(invalid(
                "form-control VML attribute uses an unbound namespace prefix",
            ));
        }
        if !matches!(namespace, ResolveResult::Unbound) || local.as_ref() != b"id" {
            continue;
        }
        if value.is_some() {
            return Err(invalid(
                "form-control VML shape has duplicate unqualified id attributes",
            ));
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(error.to_string()))?;
        value = Some(decoded.into_owned());
    }
    Ok(value.as_deref() == Some(expected))
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}
