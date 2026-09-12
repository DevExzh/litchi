//! Source-preserving deletion of exact current PackageMetadata components.
//!
//! A component registration is a repeated `PackageMetadata.components` field
//! (field 3), while incoming references can live in either the current
//! component's field 6 or a current/versioned component's field 18.  This
//! module keeps that structural operation separate from object UUID and
//! external-reference ownership removals: deleting a registration is allowed
//! to remove all incoming records for its identifier, including versioned
//! records, whereas an ordinary external-reference removal is deliberately
//! current-only.

use core::mem::size_of;

use super::{
    Budget, ComponentSelector, InvalidReason, RewriteError, RewriteExecutionLimits,
    RewriteExecutionRequirements, RewriteLimit, RewriteOptions, RewriteOutput, RewriteReport,
    add_reports, checked_add, checked_append, checked_put_key, checked_put_varint,
    component_header, decode_external_reference, next_field, preflight_execution, repeated_counter,
    set_once,
};

/// Borrowed exact selectors for current component registrations to remove.
///
/// Each selector must match one current component by identifier and effective
/// locator.  The identifier must not occur in `versioned_components`; all
/// incoming current and versioned external references to a selected
/// identifier are removed in the same candidate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentRemovalBatch<'source> {
    components: &'source [ComponentSelector<'source>],
}

impl<'source> ComponentRemovalBatch<'source> {
    /// Construct a borrowed batch of exact current-component selectors.
    #[must_use]
    pub const fn new(components: &'source [ComponentSelector<'source>]) -> Self {
        Self { components }
    }

    /// Return the selectors covered by this batch.
    #[must_use]
    pub const fn components(self) -> &'source [ComponentSelector<'source>] {
        self.components
    }
}

/// Output-free, semantically validated current-component removal rewrite.
pub struct PreparedPackageMetadataComponentRemovalRewrite<'source, 'batch> {
    source: &'source [u8],
    batch: ComponentRemovalBatch<'batch>,
    budget: Budget,
    prepare_report: RewriteReport,
    requirements: RewriteExecutionRequirements,
    output_size: usize,
    planned_fields: usize,
    planned_work: usize,
    planned_components: usize,
    planned_references: usize,
}

impl PreparedPackageMetadataComponentRemovalRewrite<'_, '_> {
    /// Return the exact source-side preparation report.
    #[must_use]
    pub const fn prepare_report(&self) -> RewriteReport {
        self.prepare_report
    }

    /// Return the exact resources required by [`Self::execute`].
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Emit and strictly verify the one candidate after checking execution
    /// limits.  No candidate allocation occurs when a supplied limit is too
    /// small.
    pub fn execute(
        mut self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, RewriteError> {
        preflight_execution(self.requirements, limits)?;
        let before = self.budget.report();

        let mut candidate = Vec::new();
        #[cfg(test)]
        super::record_output_allocation();
        candidate
            .try_reserve_exact(self.output_size)
            .map_err(|_error| RewriteError::allocation(self.output_size))?;
        if candidate.capacity() != self.output_size {
            return Err(RewriteError::allocation(self.output_size));
        }
        self.budget.allocation(0)?;
        rewrite_component_removals_into(self.source, self.batch, &mut candidate, &mut self.budget)?;
        if candidate.len() != self.output_size {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }

        self.budget.source_phase = false;
        let mut verified = ComponentRemovalScanState::new(self.batch, &mut self.budget)?;
        scan_component_removal_metadata(
            &candidate,
            self.batch,
            &mut verified,
            &mut self.budget,
            ScanMode::Verification,
        )?;
        verified.validate_candidate()?;
        self.budget.pad_repeated_counters(
            self.planned_fields,
            self.planned_work,
            self.planned_components,
            self.planned_references,
        )?;
        self.budget.output_bytes = candidate.len();
        self.budget.retained_bytes = candidate.len();
        let report = super::subtract_report(self.budget.report(), before)?;
        super::validate_execution_report(report, self.requirements)?;
        Ok(RewriteOutput {
            bytes: candidate,
            report,
        })
    }
}

/// Prepare removal of exact current component registrations without allocating
/// candidate output.
pub fn prepare_package_metadata_component_removals<'source, 'batch>(
    source: &'source [u8],
    batch: ComponentRemovalBatch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataComponentRemovalRewrite<'source, 'batch>, RewriteError> {
    validate_component_removal_batch(batch, options)?;
    let mut budget = Budget::new_inspection(source, options)?;
    budget.removals = batch.components.len();
    validate_component_removal_duplicates(batch, &mut budget)?;

    let mut source_state = ComponentRemovalScanState::new(batch, &mut budget)?;
    scan_component_removal_metadata(
        source,
        batch,
        &mut source_state,
        &mut budget,
        ScanMode::Source,
    )?;
    source_state.validate_source()?;

    let output_size = component_removal_output_size(source, batch, &mut budget)?;
    budget.output_size(output_size)?;

    // Charge the same bounded traversals that execution will perform before
    // any candidate buffer is reserved.  The source shape is a conservative
    // bound for candidate verification because this operation only removes
    // fields; the candidate pads counters to the planned shape after its
    // strict readback.
    let measured = budget.clone();
    budget.source_phase = false;
    charge_component_removal_rewrite(source, batch, &mut budget)?;
    charge_component_removal_candidate_verification(source, batch, output_size, &mut budget)?;
    budget.preflight_repeat_delta(&measured)?;
    let planned_fields = repeated_counter(measured.fields, budget.fields)?;
    let planned_work = repeated_counter(measured.work_bytes, budget.work_bytes)?;
    let planned_components =
        repeated_counter(measured.components_scanned, budget.components_scanned)?;
    let planned_references =
        repeated_counter(measured.references_scanned, budget.references_scanned)?;

    let prepare_report = budget.report();
    let execution = RewriteReport {
        input_bytes: 0,
        output_bytes: 0,
        fields: planned_fields
            .checked_sub(prepare_report.fields())
            .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
        work_bytes: planned_work
            .checked_sub(prepare_report.work_bytes())
            .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
        max_depth: prepare_report.max_depth(),
        components_scanned: planned_components
            .checked_sub(prepare_report.components_scanned())
            .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
        components_changed: 0,
        references_scanned: planned_references
            .checked_sub(prepare_report.references_scanned())
            .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
        source_references_scanned: 0,
        additions: 0,
        removals: 0,
        allocations: 0,
        retained_bytes: 0,
        scratch_bytes: 0,
    };
    let requirements = component_removal_execution_requirements(batch, output_size, execution)?;
    Ok(PreparedPackageMetadataComponentRemovalRewrite {
        source,
        batch,
        budget,
        prepare_report,
        requirements,
        output_size,
        planned_fields,
        planned_work,
        planned_components,
        planned_references,
    })
}

/// Rewrite one exact current component registration and strictly verify the
/// resulting source-preserving candidate.
pub fn rewrite_package_metadata_component_removal(
    source: &[u8],
    component: ComponentSelector<'_>,
    options: RewriteOptions,
) -> Result<RewriteOutput, RewriteError> {
    let selectors = [component];
    rewrite_package_metadata_component_removals(
        source,
        ComponentRemovalBatch::new(&selectors),
        options,
    )
}

/// Rewrite exact current component registrations and strictly verify the
/// resulting source-preserving candidate.
pub fn rewrite_package_metadata_component_removals(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    options: RewriteOptions,
) -> Result<RewriteOutput, RewriteError> {
    let prepared = prepare_package_metadata_component_removals(source, batch, options)?;
    let prepare_report = prepared.prepare_report();
    let limits = prepared.execution_requirements().exact_limits();
    let mut output = prepared.execute(limits)?;
    output.report = add_reports(prepare_report, output.report())?;
    Ok(output)
}

#[derive(Clone, Copy, Default)]
struct ComponentRemovalMatchCount {
    current_identifier: usize,
    current_locator: usize,
    current_exact: usize,
    versioned_identifier: usize,
}

struct ComponentRemovalScanState {
    selectors: Vec<ComponentRemovalMatchCount>,
}

impl ComponentRemovalScanState {
    fn new(batch: ComponentRemovalBatch<'_>, budget: &mut Budget) -> Result<Self, RewriteError> {
        Ok(Self {
            selectors: super::zeroed_vec(batch.components.len(), budget)?,
        })
    }

    fn validate_source(&self) -> Result<(), RewriteError> {
        for matched in &self.selectors {
            if matched.versioned_identifier != 0 {
                return Err(RewriteError::invalid(InvalidReason::VersionedComponent));
            }
            if matched.current_identifier > 1 || matched.current_exact > 1 {
                return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
            }
            if matched.current_identifier != 1
                || matched.current_locator != 1
                || matched.current_exact != 1
            {
                return Err(RewriteError::invalid(InvalidReason::ComponentMismatch));
            }
        }
        Ok(())
    }

    fn validate_candidate(&self) -> Result<(), RewriteError> {
        if self
            .selectors
            .iter()
            .any(|matched| matched.current_identifier != 0 || matched.current_exact != 0)
        {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum ScanMode {
    Source,
    Verification,
}

fn validate_component_removal_batch(
    batch: ComponentRemovalBatch<'_>,
    options: RewriteOptions,
) -> Result<(), RewriteError> {
    if batch.components.is_empty() {
        return Err(RewriteError::invalid(InvalidReason::RemovalNotFound));
    }
    if batch.components.len() > options.max_additions() {
        return Err(RewriteError::limited(RewriteLimit::Additions {
            observed: batch.components.len(),
            maximum: options.max_additions(),
        }));
    }
    for component in batch.components.iter().copied() {
        // Component registration removal has to retain the host's legacy
        // acceptance of an empty effective locator.  The selector still has
        // an exact nonzero identifier and is matched against the source
        // locator before the registration is discarded.  Ordinary UUID,
        // external-reference, and save-token selectors remain strict through
        // `super::validate_selector`.
        if component.identifier() == 0 {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
    }
    Ok(())
}

fn validate_component_removal_duplicates(
    batch: ComponentRemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    for (index, selector) in batch.components.iter().copied().enumerate() {
        for prior in batch.components[..index].iter().copied() {
            budget.work(
                prior
                    .locator()
                    .len()
                    .checked_add(selector.locator().len())
                    .and_then(|value| value.checked_add(1))
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )?;
            if prior.identifier() == selector.identifier() {
                return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
            }
        }
    }
    Ok(())
}

fn component_removal_execution_requirements(
    batch: ComponentRemovalBatch<'_>,
    output_size: usize,
    predicted: RewriteReport,
) -> Result<RewriteExecutionRequirements, RewriteError> {
    let state_scratch = batch
        .components
        .len()
        .checked_mul(size_of::<ComponentRemovalMatchCount>())
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let state_allocations = usize::from(!batch.components.is_empty());
    Ok(RewriteExecutionRequirements {
        output_bytes: output_size,
        fields: predicted.fields(),
        work_bytes: predicted.work_bytes(),
        components: predicted.components_scanned(),
        references: predicted.references_scanned(),
        allocations: state_allocations
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        retained_bytes: output_size,
        scratch_bytes: output_size
            .checked_add(state_scratch)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
    })
}

fn scan_component_removal_metadata(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    state: &mut ComponentRemovalScanState,
    budget: &mut Budget,
    mode: ScanMode,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut last = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            1 => set_once(&mut last, field.varint()?)?,
            3 | 11 => scan_component_removal_component(
                field.bytes()?,
                field.number == 3,
                batch,
                state,
                budget,
                mode,
                2,
            )?,
            _ => {},
        }
    }
    let last = last
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    budget.message(source, 1)?;
    let view: super::projection::PackageMetadataArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_last_object_identifier() || view.last_object_identifier != last {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(())
}

fn scan_component_removal_component(
    source: &[u8],
    current: bool,
    batch: ComponentRemovalBatch<'_>,
    state: &mut ComponentRemovalScanState,
    budget: &mut Budget,
    mode: ScanMode,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.component()?;
    let (identifier, locator) = component_removal_header(source, budget, depth)?;
    if matches!(mode, ScanMode::Source) {
        for (index, selector) in batch.components.iter().copied().enumerate() {
            budget.work(1)?;
            let matched = &mut state.selectors[index];
            if current {
                if identifier == selector.identifier() {
                    matched.current_identifier = checked_add(matched.current_identifier, 1)?;
                }
                if locator == selector.locator() {
                    matched.current_locator = checked_add(matched.current_locator, 1)?;
                }
                if identifier == selector.identifier() && locator == selector.locator() {
                    matched.current_exact = checked_add(matched.current_exact, 1)?;
                }
            } else if identifier == selector.identifier() {
                matched.versioned_identifier = checked_add(matched.versioned_identifier, 1)?;
            }
        }
    } else if current && component_removal_selected(identifier, locator, batch, budget)? {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }

    budget.message(source, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if !matches!(field.number, 6 | 18) {
            continue;
        }
        let reference = decode_external_reference(field.bytes()?, budget, child_depth)?;
        let selected = component_removal_target_selected(reference.target, batch, budget)?;
        if !selected {
            continue;
        }
        if matches!(mode, ScanMode::Source) {
            if reference.unknown_fields {
                return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
            }
        } else {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
    }
    Ok(())
}

fn component_removal_header<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(u64, &'source str), RewriteError> {
    let (identifier, locator) = component_header(source, budget, depth)?;
    budget.message(source, depth)?;
    let view: super::projection::ComponentInfoArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let effective_locator = view.locator.unwrap_or(view.preferred_locator);
    if !view.has_identifier()
        || !view.has_preferred_locator()
        || view.identifier != identifier
        || effective_locator != locator
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok((identifier, locator))
}

fn component_removal_selected(
    identifier: u64,
    locator: &str,
    batch: ComponentRemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<bool, RewriteError> {
    for selector in batch.components.iter().copied() {
        budget.work(1)?;
        if selector.identifier() == identifier && selector.locator() == locator {
            return Ok(true);
        }
    }
    Ok(false)
}

fn component_removal_target_selected(
    identifier: u64,
    batch: ComponentRemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<bool, RewriteError> {
    for selector in batch.components.iter().copied() {
        budget.work(1)?;
        if selector.identifier() == identifier {
            return Ok(true);
        }
    }
    Ok(false)
}

fn component_removal_output_size(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<usize, RewriteError> {
    budget.message(source, 1)?;
    let mut size = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            3 | 11 => {
                let payload = field.bytes()?;
                let (identifier, locator) = component_removal_header(payload, budget, 2)?;
                if field.number == 3
                    && component_removal_selected(identifier, locator, batch, budget)?
                {
                    continue;
                }
                let new_len = component_removal_component_size(payload, batch, budget, 2)?;
                size = checked_add(
                    size,
                    if new_len == payload.len() {
                        field.raw.len()
                    } else {
                        super::length_delimited_field_len(field.number, new_len)?
                    },
                )?;
            },
            _ => size = checked_add(size, field.raw.len())?,
        }
    }
    Ok(size)
}

fn component_removal_component_size(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<usize, RewriteError> {
    budget.message(source, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut size = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if matches!(field.number, 6 | 18) {
            let reference = decode_external_reference(field.bytes()?, budget, child_depth)?;
            if component_removal_target_selected(reference.target, batch, budget)? {
                if reference.unknown_fields {
                    return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
                }
                continue;
            }
        }
        size = checked_add(size, field.raw.len())?;
    }
    Ok(size)
}

fn charge_component_removal_rewrite(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if !matches!(field.number, 3 | 11) {
            continue;
        }
        let payload = field.bytes()?;
        let (identifier, locator) = component_removal_header(payload, budget, 2)?;
        if field.number == 3 && component_removal_selected(identifier, locator, batch, budget)? {
            budget.changed_component()?;
            continue;
        }
        let new_len = component_removal_component_size(payload, batch, budget, 2)?;
        if new_len != payload.len() {
            budget.changed_component()?;
            charge_component_removal_payload_rewrite(payload, batch, budget, 2)?;
        }
    }
    Ok(())
}

fn charge_component_removal_payload_rewrite(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.message(source, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if matches!(field.number, 6 | 18) {
            let reference = decode_external_reference(field.bytes()?, budget, child_depth)?;
            let _ = component_removal_target_selected(reference.target, batch, budget)?;
        }
    }
    Ok(())
}

fn charge_component_removal_candidate_verification(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    output_size: usize,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message_len(output_size, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if !matches!(field.number, 3 | 11) {
            continue;
        }
        let payload = field.bytes()?;
        let (identifier, locator) = component_removal_header(payload, budget, 2)?;
        if field.number == 3 && component_removal_selected(identifier, locator, batch, budget)? {
            continue;
        }
        budget.component()?;
        budget.message(payload, 2)?;
        let child_depth = 3;
        let mut nested = payload;
        while let Some(nested_field) = next_field(&mut nested, budget, 2)? {
            if matches!(nested_field.number, 6 | 18) {
                let reference =
                    decode_external_reference(nested_field.bytes()?, budget, child_depth)?;
                let _ = component_removal_target_selected(reference.target, batch, budget)?;
            }
        }
    }
    budget.message_len(output_size, 1)?;
    Ok(())
}

fn rewrite_component_removals_into(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    output: &mut Vec<u8>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if !matches!(field.number, 3 | 11) {
            checked_append(output, field.raw)?;
            continue;
        }
        let payload = field.bytes()?;
        let (identifier, locator) = component_removal_header(payload, budget, 2)?;
        if field.number == 3 && component_removal_selected(identifier, locator, batch, budget)? {
            budget.changed_component()?;
            continue;
        }
        let new_len = component_removal_component_size(payload, batch, budget, 2)?;
        if new_len == payload.len() {
            checked_append(output, field.raw)?;
            continue;
        }
        budget.changed_component()?;
        checked_put_key(output, field.number, 2)?;
        checked_put_varint(
            output,
            u64::try_from(new_len)
                .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
        )?;
        rewrite_component_removal_payload(payload, batch, output, budget, 2)?;
    }
    Ok(())
}

fn rewrite_component_removal_payload(
    source: &[u8],
    batch: ComponentRemovalBatch<'_>,
    output: &mut Vec<u8>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.message(source, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if matches!(field.number, 6 | 18) {
            let reference = decode_external_reference(field.bytes()?, budget, child_depth)?;
            if component_removal_target_selected(reference.target, batch, budget)? {
                if reference.unknown_fields {
                    return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
                }
                continue;
            }
        }
        checked_append(output, field.raw)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package_metadata_codec::PackageMetadataVisitor;

    fn key(output: &mut Vec<u8>, number: u32, wire: u8) {
        varint(output, (u64::from(number) << 3) | u64::from(wire));
    }

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        key(&mut output, number, 0);
        varint(&mut output, value);
        output
    }

    fn bytes_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        key(&mut output, number, 2);
        varint(
            &mut output,
            u64::try_from(payload.len()).expect("test payload length fits"),
        );
        output.extend_from_slice(payload);
        output
    }

    fn external_reference(
        target: u64,
        object: Option<u64>,
        weak: Option<bool>,
        unknown: bool,
    ) -> Vec<u8> {
        let mut output = varint_field(1, target);
        if let Some(object) = object {
            output.extend(varint_field(2, object));
        }
        if let Some(weak) = weak {
            output.extend(varint_field(3, u64::from(weak)));
        }
        if unknown {
            output.extend(varint_field(50, 9));
        }
        output
    }

    fn component(
        identifier: u64,
        locator: &str,
        current_references: &[Vec<u8>],
        versioned_references: &[Vec<u8>],
        unknown: bool,
    ) -> Vec<u8> {
        component_with_locator(
            identifier,
            locator,
            None,
            current_references,
            versioned_references,
            unknown,
        )
    }

    fn component_with_locator(
        identifier: u64,
        preferred_locator: &str,
        locator: Option<&str>,
        current_references: &[Vec<u8>],
        versioned_references: &[Vec<u8>],
        unknown: bool,
    ) -> Vec<u8> {
        let mut output = varint_field(1, identifier);
        output.extend(bytes_field(2, preferred_locator.as_bytes()));
        if let Some(locator) = locator {
            output.extend(bytes_field(3, locator.as_bytes()));
        }
        for reference in current_references {
            output.extend(bytes_field(6, reference));
        }
        for reference in versioned_references {
            output.extend(bytes_field(18, reference));
        }
        if unknown {
            output.extend(varint_field(50, 0x42));
        }
        output
    }

    fn metadata(current: &[Vec<u8>], versioned: &[Vec<u8>]) -> Vec<u8> {
        let mut output = varint_field(1, 100);
        for component in current {
            output.extend(bytes_field(3, component));
        }
        for component in versioned {
            output.extend(bytes_field(11, component));
        }
        output.extend(varint_field(60, 7));
        output
    }

    fn options(source: &[u8]) -> RewriteOptions {
        RewriteOptions::new(
            source.len(),
            source.len().saturating_add(10_000),
            100_000,
            10_000_000,
            16,
            10_000,
            10_000,
            10_000,
        )
    }

    struct NoopVisitor;

    impl PackageMetadataVisitor for NoopVisitor {}

    #[test]
    fn removes_registration_and_current_and_versioned_incoming_references() {
        let current_reference_to_removed = external_reference(2, Some(22), Some(false), false);
        let current_versioned_reference_to_removed = external_reference(2, Some(23), None, false);
        let versioned_reference_to_removed = external_reference(2, None, Some(true), false);
        let versioned_current_reference_to_removed = external_reference(2, Some(24), None, false);
        let unrelated_reference = external_reference(9, Some(99), None, false);
        let removed = component(2, "b.iwa", &[], &[], true);
        let survivor = component(
            1,
            "a.iwa",
            &[
                current_reference_to_removed.clone(),
                current_versioned_reference_to_removed.clone(),
                unrelated_reference.clone(),
            ],
            &[],
            true,
        );
        let versioned = component(
            3,
            "old.iwa",
            &[versioned_current_reference_to_removed.clone()],
            &[versioned_reference_to_removed.clone()],
            false,
        );
        let source = metadata(&[survivor, removed.clone()], &[versioned]);
        let selector = ComponentSelector::new(2, "b.iwa");
        let output =
            rewrite_package_metadata_component_removal(&source, selector, options(&source))
                .expect("component removal should succeed");

        assert!(
            !output
                .bytes()
                .windows(removed.len())
                .any(|window| window == removed)
        );
        assert!(
            !output
                .bytes()
                .windows(current_reference_to_removed.len())
                .any(|window| window == current_reference_to_removed)
        );
        assert!(
            !output
                .bytes()
                .windows(versioned_reference_to_removed.len())
                .any(|window| window == versioned_reference_to_removed)
        );
        assert!(
            !output
                .bytes()
                .windows(current_versioned_reference_to_removed.len())
                .any(|window| window == current_versioned_reference_to_removed)
        );
        assert!(
            !output
                .bytes()
                .windows(versioned_current_reference_to_removed.len())
                .any(|window| window == versioned_current_reference_to_removed)
        );
        assert!(
            output
                .bytes()
                .windows(unrelated_reference.len())
                .any(|window| window == unrelated_reference)
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| { window == [0x90, 0x03, 0x42] })
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| { window == [0xe0, 0x03, 0x07] })
        );

        let mut visitor = NoopVisitor;
        super::super::inspect_package_metadata_with_visitor(
            output.bytes(),
            options(output.bytes()),
            &mut visitor,
        )
        .expect("candidate must strictly re-decode");
    }

    #[test]
    fn prepared_execution_checks_limits_before_candidate_allocation() {
        let source = metadata(
            &[
                component(1, "a.iwa", &[], &[], false),
                component(2, "b.iwa", &[], &[], false),
            ],
            &[],
        );
        let selectors = [ComponentSelector::new(2, "b.iwa")];
        let prepared = prepare_package_metadata_component_removals(
            &source,
            ComponentRemovalBatch::new(&selectors),
            options(&source),
        )
        .expect("preparation should succeed");
        let requirements = prepared.execution_requirements();
        let mut limited = requirements.exact_limits();
        limited.max_output_bytes = requirements.output_bytes() - 1;
        let before = super::super::output_allocations();
        assert!(prepared.execute(limited).is_err());
        assert_eq!(super::super::output_allocations(), before);

        for axis in 0..7 {
            let prepared = prepare_package_metadata_component_removals(
                &source,
                ComponentRemovalBatch::new(&selectors),
                options(&source),
            )
            .expect("preparation should succeed");
            let requirements = prepared.execution_requirements();
            let mut limited = requirements.exact_limits();
            let can_limit = match axis {
                0 if requirements.fields() != 0 => {
                    limited.max_fields = requirements.fields() - 1;
                    true
                },
                1 if requirements.work_bytes() != 0 => {
                    limited.max_work_bytes = requirements.work_bytes() - 1;
                    true
                },
                2 if requirements.components() != 0 => {
                    limited.max_components = requirements.components() - 1;
                    true
                },
                3 if requirements.references() != 0 => {
                    limited.max_references = requirements.references() - 1;
                    true
                },
                4 if requirements.allocations() != 0 => {
                    limited.max_allocations = requirements.allocations() - 1;
                    true
                },
                5 if requirements.retained_bytes() != 0 => {
                    limited.max_retained_bytes = requirements.retained_bytes() - 1;
                    true
                },
                6 if requirements.scratch_bytes() != 0 => {
                    limited.max_scratch_bytes = requirements.scratch_bytes() - 1;
                    true
                },
                _ => false,
            };
            if can_limit {
                let before = super::super::output_allocations();
                assert!(prepared.execute(limited).is_err());
                assert_eq!(super::super::output_allocations(), before);
            }
        }
    }

    #[test]
    fn rejects_duplicate_current_registration() {
        let duplicate_a = component(2, "b.iwa", &[], &[], false);
        let duplicate_b = component(2, "b.iwa", &[], &[], false);
        let source = metadata(&[duplicate_a, duplicate_b], &[]);
        let error = rewrite_package_metadata_component_removal(
            &source,
            ComponentSelector::new(2, "b.iwa"),
            options(&source),
        )
        .expect_err("duplicate current registration must be refused");
        assert_eq!(
            error.invalid_reason(),
            Some(InvalidReason::DuplicateRemoval)
        );
    }

    #[test]
    fn rejects_versioned_alias() {
        let current = component(2, "b.iwa", &[], &[], false);
        let versioned = component(2, "old-b.iwa", &[], &[], false);
        let source = metadata(&[current], &[versioned]);
        let error = rewrite_package_metadata_component_removal(
            &source,
            ComponentSelector::new(2, "b.iwa"),
            options(&source),
        )
        .expect_err("versioned alias must be refused");
        assert_eq!(
            error.invalid_reason(),
            Some(InvalidReason::VersionedComponent)
        );
    }

    #[test]
    fn rejects_unknown_selected_reference_payload() {
        let reference = external_reference(2, Some(22), Some(false), true);
        let survivor = component(1, "a.iwa", &[reference], &[], false);
        let removed = component(2, "b.iwa", &[], &[], false);
        let source = metadata(&[survivor, removed], &[]);
        let error = rewrite_package_metadata_component_removal(
            &source,
            ComponentSelector::new(2, "b.iwa"),
            options(&source),
        )
        .expect_err("unknown selected reference must be refused");
        assert_eq!(error.invalid_reason(), Some(InvalidReason::RemovalMismatch));
    }

    #[test]
    fn removes_component_with_explicit_empty_effective_locator_and_preserves_unknowns() {
        let selected = component_with_locator(2, "legacy-fallback", Some(""), &[], &[], false);
        let survivor = component(1, "a.iwa", &[], &[], true);
        let source = metadata(&[survivor, selected.clone()], &[]);
        let output = rewrite_package_metadata_component_removal(
            &source,
            ComponentSelector::new(2, ""),
            options(&source),
        )
        .expect("explicit empty effective locator should be removable");

        assert!(
            !output
                .bytes()
                .windows(selected.len())
                .any(|window| window == selected)
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| window == [0x90, 0x03, 0x42])
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| window == [0xe0, 0x03, 0x07])
        );
    }

    #[test]
    fn removes_component_with_empty_preferred_locator_fallback_and_preserves_unknowns() {
        let selected = component_with_locator(2, "", None, &[], &[], false);
        let survivor = component(1, "a.iwa", &[], &[], true);
        let source = metadata(&[survivor, selected.clone()], &[]);
        let output = rewrite_package_metadata_component_removal(
            &source,
            ComponentSelector::new(2, ""),
            options(&source),
        )
        .expect("empty preferred locator should be removable");

        assert!(
            !output
                .bytes()
                .windows(selected.len())
                .any(|window| window == selected)
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| window == [0x90, 0x03, 0x42])
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| window == [0xe0, 0x03, 0x07])
        );
    }

    #[test]
    fn rejects_malformed_selected_reference_payload() {
        let survivor = component(1, "a.iwa", &[vec![0x08, 0x80]], &[], false);
        let removed = component(2, "b.iwa", &[], &[], false);
        let source = metadata(&[survivor, removed], &[]);
        let error = rewrite_package_metadata_component_removal(
            &source,
            ComponentSelector::new(2, "b.iwa"),
            options(&source),
        )
        .expect_err("malformed selected reference must be refused");
        assert_eq!(error.invalid_reason(), Some(InvalidReason::MalformedWire));
    }
}
