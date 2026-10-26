//! The package metadata registry side of a structural transaction.
//!
//! `Index/Metadata.iwa` registers every component, the UUIDs of objects
//! other components address by UUID, and every cross-component object
//! reference. A structural edit that adds or removes objects must keep those
//! records exact: removed objects lose their UUID bindings and the external
//! references that pointed at them, new objects that address another
//! component gain an external reference, and the issued-identifier watermark
//! covers every new identifier. The transition is applied through the
//! shared strict codec, which preserves every unselected raw record.

use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::package_metadata_codec::{
    Batch, CombinedBatch, CombinedSaveTokenBatch, ComponentDescriptor, ComponentRemovalBatch,
    ComponentSelector, DataReferenceOwnerDescriptor, ExternalReferenceAddition,
    ExternalReferenceDescriptor, ExternalReferenceRemoval, ObjectUuidAddition,
    ObjectUuidDescriptor, ObjectUuidRemoval, PackageMetadataVisitor, RemovalBatch,
    RemovalSaveTokenBatch, RewriteError, RewriteLimit, RewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor, rewrite_package_metadata,
    rewrite_package_metadata_combined_additions_and_removals_and_save_tokens,
    rewrite_package_metadata_component_removals, rewrite_package_metadata_removals_and_save_tokens,
};

use super::super::metadata::{ENTRY_NAME, MESSAGE_TYPE, normalized_locator};
use super::{Fault, FaultLimit, FaultResult, Working, allocation, reserve};

/// One current component registration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::package) struct Component {
    pub(in crate::package) identifier: u64,
    pub(in crate::package) locator: String,
}

/// One object UUID binding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) struct Uuid {
    pub(in crate::package) component: u64,
    pub(in crate::package) object: u64,
    pub(in crate::package) uuid: UuidBits,
}

/// One cross-component object reference record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::package) struct External {
    pub(in crate::package) source: u64,
    pub(in crate::package) target: u64,
    pub(in crate::package) object: u64,
    pub(in crate::package) weak: Option<bool>,
}

/// The registry facts of one metadata payload.
#[derive(Debug, Default)]
pub(in crate::package) struct Registry {
    pub(in crate::package) components: Vec<Component>,
    pub(in crate::package) uuids: Vec<Uuid>,
    pub(in crate::package) externals: Vec<External>,
    /// Objects that own a data (media) reference.
    pub(in crate::package) data_owners: Vec<u64>,
    pub(in crate::package) last_object_identifier: u64,
    unknown: bool,
}

impl Registry {
    /// The current component registered for member `name`.
    pub(in crate::package) fn component_for_entry(&self, name: &str) -> FaultResult<&Component> {
        let locator = normalized_locator(name);
        let mut found = None;
        for component in &self.components {
            if component.locator == locator && found.replace(component).is_some() {
                return Err(Fault::Invalid);
            }
        }
        found.ok_or(Fault::Invalid)
    }

    fn component(&self, identifier: u64) -> FaultResult<&Component> {
        self.components
            .iter()
            .find(|component| component.identifier == identifier)
            .ok_or(Fault::Invalid)
    }
}

#[derive(Default)]
struct Visitor {
    registry: Registry,
}

fn remember<T>(items: &mut Vec<T>, item: T) -> Result<(), RewriteError> {
    items
        .try_reserve(1)
        .map_err(|_error| RewriteError::allocation(1))?;
    items.push(item);
    Ok(())
}

impl PackageMetadataVisitor for Visitor {
    fn visit_unknown_field(&mut self) -> Result<(), RewriteError> {
        self.registry.unknown = true;
        Ok(())
    }

    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if !component.is_current() {
            return Ok(());
        }
        let mut locator = String::new();
        locator
            .try_reserve_exact(component.effective_locator().len())
            .map_err(|_error| RewriteError::allocation(component.effective_locator().len()))?;
        locator.push_str(component.effective_locator());
        remember(
            &mut self.registry.components,
            Component {
                identifier: component.identifier(),
                locator,
            },
        )
    }

    fn visit_object_uuid(&mut self, binding: ObjectUuidDescriptor<'_>) -> Result<(), RewriteError> {
        remember(
            &mut self.registry.uuids,
            Uuid {
                component: binding.component().identifier(),
                object: binding.object_identifier(),
                uuid: binding.uuid(),
            },
        )
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        if let Some(object) = reference.object_identifier() {
            remember(
                &mut self.registry.externals,
                External {
                    source: reference.source().identifier(),
                    target: reference.target_component_identifier(),
                    object,
                    weak: reference.is_weak(),
                },
            )?;
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        remember(&mut self.registry.data_owners, owner.object_identifier())
    }
}

fn map_rewrite(error: &RewriteError) -> Fault {
    if let Some(amount) = error.allocation_request() {
        return allocation(amount);
    }
    let figures = |observed: usize, maximum: usize| Fault::Limit {
        kind: FaultLimit::PayloadBytes,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    };
    match error.resource_limit() {
        Some(
            RewriteLimit::InputBytes { observed, maximum }
            | RewriteLimit::OutputBytes { observed, maximum }
            | RewriteLimit::Fields { observed, maximum }
            | RewriteLimit::Work { observed, maximum }
            | RewriteLimit::Components { observed, maximum }
            | RewriteLimit::References { observed, maximum }
            | RewriteLimit::Additions { observed, maximum },
        ) => figures(observed, maximum),
        Some(RewriteLimit::Nesting { observed, maximum }) => {
            figures(observed as usize, maximum as usize)
        },
        Some(_) => figures(0, 0),
        None => Fault::Invalid,
    }
}

/// The metadata object and message in the working set, if the package has
/// a registry.
fn route(working: &Working<'_>) -> FaultResult<Option<(u64, usize)>> {
    let Ok(archive) = working.archive(ENTRY_NAME) else {
        return Ok(None);
    };
    let mut found = None;
    for object in &archive.objects {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == MESSAGE_TYPE {
                let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                if found.replace((identifier, index)).is_some() {
                    return Err(Fault::Invalid);
                }
            }
        }
    }
    found.map(Some).ok_or(Fault::Invalid)
}

fn payload<'w>(working: &'w Working<'_>, (object, message): (u64, usize)) -> FaultResult<&'w [u8]> {
    working
        .archive(ENTRY_NAME)?
        .object(object)
        .and_then(|object| object.messages.get(message))
        .map(|message| message.data.as_slice())
        .ok_or(Fault::Invalid)
}

fn options(working: &Working<'_>, payload: usize, additions: usize) -> RewriteOptions {
    let components = working
        .source()
        .state
        .components
        .catalog()
        .len()
        .max(payload)
        .max(1);
    let output = payload.saturating_mul(2).saturating_add(4_096);
    RewriteOptions::new(
        payload.max(1),
        output,
        payload
            .saturating_mul(4)
            .saturating_add(additions.saturating_mul(16))
            .max(64),
        payload
            .saturating_mul(256)
            .saturating_add(
                components
                    .saturating_mul(additions.max(1))
                    .saturating_mul(64),
            )
            .clamp(1, litchi_iwa_common::WireLimits::MAX_REWRITE_WORK),
        64,
        components.saturating_add(additions),
        working
            .source()
            .state
            .options
            .semantic()
            .max_references()
            .max(1),
        additions.max(1),
    )
}

/// The registry of the package, or `None` without `Index/Metadata.iwa`.
pub(in crate::package) fn registry(working: &Working<'_>) -> FaultResult<Option<Registry>> {
    let Some(route) = route(working)? else {
        return Ok(None);
    };
    let source = payload(working, route)?;
    let mut visitor = Visitor::default();
    let inspection = inspect_package_metadata_with_visitor(
        source,
        options(working, source.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_rewrite(&error))?;
    visitor.registry.last_object_identifier = inspection.last_object_identifier();
    if visitor.registry.unknown {
        return Err(Fault::Unsupported);
    }
    Ok(Some(visitor.registry))
}

/// The issued-identifier watermark, or `None` without a registry.
pub(super) fn last_object_identifier(working: &Working<'_>) -> FaultResult<Option<u64>> {
    Ok(registry(working)?.map(|registry| registry.last_object_identifier))
}

/// The registry records one transaction adds and removes.
#[derive(Debug, Default)]
pub(in crate::package) struct Transition {
    pub(in crate::package) uuid_additions: Vec<Uuid>,
    pub(in crate::package) uuid_removals: Vec<Uuid>,
    pub(in crate::package) external_removals: Vec<External>,
    pub(in crate::package) external_additions: Vec<External>,
    /// Removed objects whose identifiers may lower the watermark when they
    /// are its suffix (the host editor's `release_package_identifier_suffix`).
    pub(in crate::package) released: Vec<u64>,
}

impl Transition {
    /// Lower the watermark over `objects` when the last registered
    /// identifier is one of them (record 0370).
    pub(in crate::package) fn release_suffix(&mut self, objects: &[u64]) -> FaultResult<()> {
        reserve(&mut self.released, objects.len())?;
        self.released.extend_from_slice(objects);
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.uuid_additions.is_empty()
            && self.uuid_removals.is_empty()
            && self.external_removals.is_empty()
            && self.external_additions.is_empty()
    }

    /// Record the registry consequences of removing `objects`: their UUID
    /// bindings and every external reference that addresses them.
    pub(in crate::package) fn remove_objects(
        &mut self,
        registry: &Registry,
        objects: &[u64],
    ) -> FaultResult<()> {
        for binding in &registry.uuids {
            if objects.contains(&binding.object) {
                reserve(&mut self.uuid_removals, 1)?;
                self.uuid_removals.push(*binding);
            }
        }
        for external in &registry.externals {
            if objects.contains(&external.object) {
                reserve(&mut self.external_removals, 1)?;
                self.external_removals.push(*external);
            }
        }
        if registry
            .data_owners
            .iter()
            .any(|owner| objects.contains(owner))
        {
            // Media ownership would need its data member collected too.
            return Err(Fault::Unsupported);
        }
        Ok(())
    }

    /// Register a fresh UUID for new object `object` of component
    /// `component`, derived from `seed` and distinct from every registered
    /// UUID (no clock or randomness is consulted).
    pub(in crate::package) fn register_uuid(
        &mut self,
        registry: &Registry,
        component: u64,
        object: u64,
        seed: u64,
    ) -> FaultResult<()> {
        let mut state = seed ^ object.rotate_left(17);
        loop {
            let lower = mix(state);
            let upper = mix(lower ^ 0x6c69_7463_6869_0301);
            let uuid = UuidBits::new(lower, upper);
            let taken = registry.uuids.iter().any(|binding| binding.uuid == uuid)
                || self
                    .uuid_additions
                    .iter()
                    .any(|binding| binding.uuid == uuid);
            if !taken {
                reserve(&mut self.uuid_additions, 1)?;
                self.uuid_additions.push(Uuid {
                    component,
                    object,
                    uuid,
                });
                return Ok(());
            }
            state = state.wrapping_add(0x9e37_79b9_7f4a_7c15);
        }
    }

    /// Record that component `source` now references `object` in component
    /// `target`, unless the registry already holds that edge.
    pub(in crate::package) fn reference(
        &mut self,
        registry: &Registry,
        source: u64,
        target: u64,
        object: u64,
    ) -> FaultResult<()> {
        if source == target {
            return Ok(());
        }
        let known = |external: &External| {
            external.source == source && external.target == target && external.object == object
        };
        if registry.externals.iter().any(known) || self.external_additions.iter().any(known) {
            return Ok(());
        }
        reserve(&mut self.external_additions, 1)?;
        self.external_additions.push(External {
            source,
            target,
            object,
            weak: None,
        });
        Ok(())
    }
}

/// Apply `transition` and raise the watermark to `allocated`.
pub(super) fn apply(
    working: &mut Working<'_>,
    transition: &Transition,
    allocated: Option<u64>,
    emptied: &[String],
) -> FaultResult<()> {
    let Some(route) = route(working)? else {
        if transition.is_empty() && emptied.is_empty() {
            return Ok(());
        }
        return Err(Fault::Unsupported);
    };
    let registry = registry(working)?.ok_or(Fault::Invalid)?;
    let current = registry.last_object_identifier;
    let mut last = allocated.map_or(current, |allocated| allocated.max(current));
    if allocated.is_none() && transition.released.contains(&current) {
        last = working.maximum_identifier();
    }
    if transition.is_empty() && last == current && emptied.is_empty() {
        return Ok(());
    }
    let source = payload(working, route)?;
    let additions = transition
        .uuid_removals
        .len()
        .saturating_add(transition.uuid_additions.len())
        .saturating_add(transition.external_removals.len())
        .saturating_add(transition.external_additions.len());
    let options = options(working, source.len(), additions);
    let output = if transition.is_empty() && last == current {
        None
    } else if transition.is_empty() {
        Some(
            rewrite_package_metadata(source, Batch::new(current, last, &[], &[]), options)
                .map_err(|error| map_rewrite(&error))?,
        )
    } else {
        let selector = |identifier: u64| -> FaultResult<ComponentSelector<'_>> {
            let component = registry.component(identifier)?;
            Ok(ComponentSelector::new(
                component.identifier,
                component.locator.as_str(),
            ))
        };
        let mut uuid_removals = Vec::new();
        reserve(&mut uuid_removals, transition.uuid_removals.len())?;
        for binding in &transition.uuid_removals {
            uuid_removals.push(ObjectUuidRemoval::new(
                selector(binding.component)?,
                binding.object,
                binding.uuid,
            ));
        }
        let mut external_removals = Vec::new();
        reserve(&mut external_removals, transition.external_removals.len())?;
        for external in &transition.external_removals {
            external_removals.push(ExternalReferenceRemoval::new(
                selector(external.source)?,
                selector(external.target)?,
                external.object,
                external.weak,
            ));
        }
        let mut uuid_additions = Vec::new();
        reserve(&mut uuid_additions, transition.uuid_additions.len())?;
        for binding in &transition.uuid_additions {
            uuid_additions.push(ObjectUuidAddition::new(
                selector(binding.component)?,
                binding.object,
                binding.uuid,
            ));
        }
        let mut external_additions = Vec::new();
        reserve(&mut external_additions, transition.external_additions.len())?;
        for external in &transition.external_additions {
            external_additions.push(ExternalReferenceAddition::new(
                selector(external.source)?,
                selector(external.target)?,
                external.object,
                external.weak,
            ));
        }
        // Every component whose registry record changes advances its save
        // token, as the native app does when it rewrites a component.
        let mut tokens: Vec<ComponentSelector<'_>> = Vec::new();
        let touched = transition
            .uuid_removals
            .iter()
            .chain(&transition.uuid_additions)
            .map(|binding| binding.component)
            .chain(
                transition
                    .external_removals
                    .iter()
                    .chain(&transition.external_additions)
                    .flat_map(|external| [external.source, external.target]),
            );
        for identifier in touched {
            let selector = selector(identifier)?;
            if !tokens.contains(&selector) {
                reserve(&mut tokens, 1)?;
                tokens.push(selector);
            }
        }
        let removes = !uuid_removals.is_empty() || !external_removals.is_empty();
        let adds = !external_additions.is_empty() || !uuid_additions.is_empty();
        let output = if removes && !adds && last <= current {
            rewrite_package_metadata_removals_and_save_tokens(
                source,
                RemovalSaveTokenBatch::new(
                    RemovalBatch::new(current, &uuid_removals, &external_removals, &[])
                        .with_new_last_object_identifier(last),
                    SaveTokenBatch::new(&tokens),
                ),
                options,
            )
        } else if adds && !removes {
            rewrite_package_metadata(
                source,
                Batch::new(current, last, &uuid_additions, &external_additions),
                options,
            )
        } else {
            rewrite_package_metadata_combined_additions_and_removals_and_save_tokens(
                source,
                CombinedSaveTokenBatch::new(
                    CombinedBatch::new(
                        current,
                        last,
                        &uuid_additions,
                        &external_additions,
                        &uuid_removals,
                        &external_removals,
                        &[],
                    ),
                    SaveTokenBatch::new(&tokens),
                ),
                options,
            )
        };
        Some(output.map_err(|error| map_rewrite(&error))?)
    };
    let mut data = match output {
        Some(output) => output.into_bytes(),
        None => super::copy_bytes(source)?,
    };
    if !emptied.is_empty() {
        let mut selectors = Vec::new();
        reserve(&mut selectors, emptied.len())?;
        for name in emptied {
            let component = registry.component_for_entry(name)?;
            selectors.push(ComponentSelector::new(
                component.identifier,
                component.locator.as_str(),
            ));
        }
        let options = self::options(working, data.len(), emptied.len());
        data = rewrite_package_metadata_component_removals(
            &data,
            ComponentRemovalBatch::new(&selectors),
            options,
        )
        .map_err(|error| map_rewrite(&error))?
        .into_bytes();
    }
    let limits = working.archive_limits();
    let object = working
        .archive_mut(ENTRY_NAME)?
        .object_mut(route.0)
        .ok_or(Fault::Invalid)?;
    object
        .replace_message_preserving_header(
            route.1,
            RawMessage {
                type_: MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|error| super::map_core(&error))?;
    Ok(())
}

/// A 64-bit finalizer (SplitMix64) for deterministic identity derivation.
pub(in crate::package) const fn mix(mut value: u64) -> u64 {
    value = value.wrapping_add(0x9e37_79b9_7f4a_7c15);
    value = (value ^ (value >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value = (value ^ (value >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}

/// A deterministic seed for the fresh identities of one transaction: the
/// registry payload (which carries the document's own identity) and the
/// first identifier the transaction allocates.
pub(in crate::package) fn identity_seed(working: &Working<'_>, first: u64) -> FaultResult<u64> {
    let mut seed = mix(first);
    if let Some(route) = route(working)? {
        for chunk in payload(working, route)?.chunks(8) {
            let mut word = [0_u8; 8];
            word[..chunk.len()].copy_from_slice(chunk);
            seed = mix(seed ^ u64::from_le_bytes(word));
        }
    }
    Ok(seed)
}
