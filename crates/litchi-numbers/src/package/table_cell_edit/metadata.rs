//! Strict PackageMetadata routing for sparse table-cell publication.

use core::mem::size_of;

use litchi_iwa_protos::package_metadata_codec::{
    ComponentDescriptor, ComponentSelector, ObjectUuidDescriptor, PackageMetadataVisitor,
    RewriteError, RewriteOptions, RewriteReport, UuidBits, inspect_package_metadata_with_visitor,
};

use crate::package::table_cells::{Error, LimitKind, Path};

use super::{Package, resolve::MessageRoute};

const PACKAGE_METADATA_TYPE: u32 = 11_006;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ComponentIdentity {
    pub(super) component_index: usize,
    pub(super) identifier: u64,
}

#[derive(Debug)]
pub(super) struct Inspection {
    pub(super) route: MessageRoute,
    pub(super) last_object_identifier: u64,
    pub(super) maximum_object_identifier: u64,
    pub(super) components: Vec<ComponentIdentity>,
    pub(super) uuids: Vec<UuidBits>,
    pub(super) first_report: RewriteReport,
    pub(super) second_report: RewriteReport,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct InspectionRequirements {
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) allocation_events: usize,
    pub(super) work: usize,
    pub(super) uuids: usize,
}

pub(super) struct InspectionPlan {
    route: MessageRoute,
    last_object_identifier: u64,
    maximum_object_identifier: u64,
    expected: Counts,
    first_report: RewriteReport,
}

impl InspectionPlan {
    pub(super) const fn first_report(&self) -> RewriteReport {
        self.first_report
    }

    pub(super) fn requirements(&self, path: Path) -> Result<InspectionRequirements, Error> {
        let retained_elements = self
            .expected
            .current_components
            .checked_add(self.expected.uuids)
            .ok_or(Error::InvalidSource { path })?;
        let retained_bytes = self
            .expected
            .current_components
            .checked_mul(size_of::<ComponentIdentity>())
            .and_then(|bytes| {
                self.expected
                    .uuids
                    .checked_mul(size_of::<UuidBits>())
                    .and_then(|uuids| bytes.checked_add(uuids))
            })
            .ok_or(Error::InvalidSource { path })?;
        let work = sort_work(self.expected.current_components, path)?
            .checked_add(sort_work(self.expected.uuids, path)?)
            .ok_or(Error::InvalidSource { path })?;
        Ok(InspectionRequirements {
            retained_elements,
            retained_bytes,
            allocation_events: usize::from(self.expected.current_components != 0)
                .checked_add(usize::from(self.expected.uuids != 0))
                .ok_or(Error::InvalidSource { path })?,
            work,
            uuids: self.expected.uuids,
        })
    }

    pub(super) fn collect(
        self,
        source: &Package,
        options: RewriteOptions,
        path: Path,
    ) -> Result<Inspection, Error> {
        let payload = message_payload(source, self.route, path)?;
        let mut collector = CollectVisitor::new(source, self.expected, false, path)?;
        let second = inspect_package_metadata_with_visitor(payload, options, &mut collector)
            .map_err(|error| map_error(error, path))?;
        validate_inspection_report(second.report(), path)?;
        if self.last_object_identifier != second.last_object_identifier()
            || collector.components.len() > self.expected.current_components
            || collector.uuids.len() != self.expected.uuids
        {
            return Err(Error::InvalidSource { path });
        }
        collector
            .components
            .sort_unstable_by_key(|component| component.component_index);
        if collector
            .components
            .windows(2)
            .any(|pair| pair[0].component_index >= pair[1].component_index)
        {
            return Err(Error::InvalidSource { path });
        }
        collector
            .uuids
            .sort_unstable_by_key(|uuid| (uuid.lower(), uuid.upper()));
        if collector.uuids.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Error::InvalidSource { path });
        }
        Ok(Inspection {
            route: self.route,
            last_object_identifier: self.last_object_identifier,
            maximum_object_identifier: self.maximum_object_identifier,
            components: collector.components,
            uuids: collector.uuids,
            first_report: self.first_report,
            second_report: second.report(),
        })
    }
}

impl Inspection {
    pub(super) fn component(
        &self,
        component_index: usize,
        path: Path,
    ) -> Result<ComponentIdentity, Error> {
        self.components
            .binary_search_by_key(&component_index, |component| component.component_index)
            .ok()
            .and_then(|index| self.components.get(index).copied())
            .ok_or(Error::InvalidSource { path })
    }

    pub(super) fn selector<'source>(
        &self,
        source: &'source Package,
        component_index: usize,
        path: Path,
    ) -> Result<ComponentSelector<'source>, Error> {
        let identity = self.component(component_index, path)?;
        let locator = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .map(|component| normalized_locator(component.name()))
            .ok_or(Error::InvalidSource { path })?;
        Ok(ComponentSelector::new(identity.identifier, locator))
    }

    pub(super) fn allocate_identifiers(
        &self,
        count: usize,
        source_bytes: &[u8],
        path: Path,
    ) -> Result<(Vec<u64>, Vec<UuidBits>), Error> {
        let mut identifiers = Vec::new();
        identifiers
            .try_reserve_exact(count)
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::Objects,
                amount: count,
            })?;
        let mut uuids = Vec::new();
        uuids
            .try_reserve_exact(count)
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: count,
            })?;
        if identifiers.capacity() != count || uuids.capacity() != count {
            return Err(Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: count,
            });
        }
        // Some valid Numbers packages contain source objects whose identifiers
        // are newer than PackageMetadata's advisory last-object watermark.
        // Allocate above both authorities; the metadata rewrite still proves
        // the exact prior watermark before advancing it to the final ID.
        let mut identifier = self
            .last_object_identifier
            .max(self.maximum_object_identifier);
        let source_seed = fingerprint(source_bytes, 0xcbf2_9ce4_8422_2325);
        for ordinal in 0..count {
            identifier = identifier.checked_add(1).ok_or(Error::LimitExceeded {
                kind: LimitKind::Objects,
                observed: u64::MAX,
                maximum: u64::MAX - 1,
                path,
            })?;
            let ordinal = u64::try_from(ordinal).map_err(|_error| Error::InvalidSource { path })?;
            let lower = identifier;
            let mut salt = 0u64;
            let uuid = loop {
                let upper = mix(source_seed.rotate_left(29)
                    ^ identifier.rotate_left(17)
                    ^ ordinal
                    ^ salt.rotate_left(7));
                let candidate = UuidBits::new(lower, upper);
                if (lower != 0 || upper != 0)
                    && self
                        .uuids
                        .binary_search_by_key(&(lower, upper), |uuid| (uuid.lower(), uuid.upper()))
                        .is_err()
                {
                    break candidate;
                }
                salt = salt.checked_add(1).ok_or(Error::InvalidSource { path })?;
                let maximum_attempts = self
                    .uuids
                    .len()
                    .checked_add(1)
                    .ok_or(Error::InvalidSource { path })?;
                if usize::try_from(salt).map_or(true, |attempts| attempts > maximum_attempts) {
                    return Err(Error::InvalidSource { path });
                }
            };
            identifiers.push(identifier);
            uuids.push(uuid);
        }
        Ok((identifiers, uuids))
    }
}

pub(super) fn inspect(
    source: &Package,
    options: RewriteOptions,
    path: Path,
) -> Result<Inspection, Error> {
    plan_inspection(source, options, path)?.collect(source, options, path)
}

pub(super) fn plan_inspection(
    source: &Package,
    options: RewriteOptions,
    path: Path,
) -> Result<InspectionPlan, Error> {
    let route = unique_metadata_route(source, path)?;
    let object = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .ok_or(Error::InvalidSource { path })?;
    crate::package::table_headers::resolve::validate_message_metadata(object, route.message_index)
        .map_err(|_error| Error::InvalidSource { path })?;
    let payload = message_payload(source, route, path)?;
    let mut counter = CountVisitor::default();
    let first = inspect_package_metadata_with_visitor(payload, options, &mut counter)
        .map_err(|error| map_error(error, path))?;
    validate_inspection_report(first.report(), path)?;
    if counter.overflow {
        return Err(Error::InvalidSource { path });
    }
    Ok(InspectionPlan {
        route,
        last_object_identifier: first.last_object_identifier(),
        maximum_object_identifier: maximum_object_identifier(source, path)?,
        expected: counter.counts,
        first_report: first.report(),
    })
}

fn validate_inspection_report(report: RewriteReport, path: Path) -> Result<(), Error> {
    // Inspection is deliberately output-free and the codec itself owns no
    // storage after returning. The count/collect visitors are the only live
    // allocations and are accounted independently by `InspectionRequirements`.
    if report.output_bytes() != 0 || report.retained_bytes() != 0 || report.allocations() != 0 {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn sort_work(elements: usize, path: Path) -> Result<usize, Error> {
    if elements < 2 {
        return Ok(elements);
    }
    let log = usize::try_from(usize::BITS - (elements - 1).leading_zeros())
        .map_err(|_| Error::InvalidSource { path })?;
    elements
        .checked_mul(log)
        .and_then(|work| work.checked_add(elements))
        .ok_or(Error::InvalidSource { path })
}

#[derive(Clone, Copy, Default)]
struct Counts {
    current_components: usize,
    uuids: usize,
}

#[derive(Default)]
struct CountVisitor {
    counts: Counts,
    overflow: bool,
}

impl PackageMetadataVisitor for CountVisitor {
    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if component.is_current() {
            match self.counts.current_components.checked_add(1) {
                Some(value) => self.counts.current_components = value,
                None => self.overflow = true,
            }
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        _binding: ObjectUuidDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        match self.counts.uuids.checked_add(1) {
            Some(value) => self.counts.uuids = value,
            None => self.overflow = true,
        }
        Ok(())
    }
}

struct CollectVisitor<'source> {
    source: &'source Package,
    components: Vec<ComponentIdentity>,
    uuids: Vec<UuidBits>,
}

impl<'source> CollectVisitor<'source> {
    fn new(
        source: &'source Package,
        counts: Counts,
        overflow: bool,
        path: Path,
    ) -> Result<Self, Error> {
        if overflow {
            return Err(Error::InvalidSource { path });
        }
        let mut components = Vec::new();
        components
            .try_reserve_exact(counts.current_components)
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: counts.current_components,
            })?;
        let mut uuids = Vec::new();
        uuids
            .try_reserve_exact(counts.uuids)
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: counts.uuids,
            })?;
        if components.capacity() != counts.current_components || uuids.capacity() != counts.uuids {
            return Err(Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: counts.current_components.max(counts.uuids),
            });
        }
        Ok(Self {
            source,
            components,
            uuids,
        })
    }
}

impl PackageMetadataVisitor for CollectVisitor<'_> {
    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if component.is_current() {
            if let Some(component_index) =
                component_index(self.source, component.effective_locator())
            {
                self.components.push(ComponentIdentity {
                    component_index,
                    identifier: component.identifier(),
                });
            }
        }
        Ok(())
    }

    fn visit_object_uuid(&mut self, binding: ObjectUuidDescriptor<'_>) -> Result<(), RewriteError> {
        self.uuids.push(binding.uuid());
        Ok(())
    }
}

fn component_index(source: &Package, locator: &str) -> Option<usize> {
    let catalog = source.state.components.catalog();
    let mut low = 0usize;
    let mut high = catalog.len();
    while low < high {
        let middle = low + (high - low) / 2;
        let component = catalog.get_index(middle)?;
        match normalized_locator(component.name()).cmp(locator) {
            std::cmp::Ordering::Less => low = middle + 1,
            std::cmp::Ordering::Greater => high = middle,
            std::cmp::Ordering::Equal => return Some(middle),
        }
    }
    None
}

fn normalized_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

fn unique_metadata_route(source: &Package, path: Path) -> Result<MessageRoute, Error> {
    let mut found = None;
    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != PACKAGE_METADATA_TYPE {
                    continue;
                }
                if found.is_some() {
                    return Err(Error::InvalidSource { path });
                }
                found = Some(MessageRoute {
                    component_index,
                    object_index,
                    message_index,
                    message_type: PACKAGE_METADATA_TYPE,
                });
            }
        }
    }
    found.ok_or(Error::InvalidSource { path })
}

fn message_payload(source: &Package, route: MessageRoute, path: Path) -> Result<&[u8], Error> {
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == route.message_type)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource { path })
}

fn maximum_object_identifier(source: &Package, path: Path) -> Result<u64, Error> {
    source
        .state
        .components
        .catalog()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .try_fold(0u64, |maximum, object| {
            object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .map(|identifier| maximum.max(identifier))
                .ok_or(Error::InvalidSource { path })
        })
}

fn fingerprint(bytes: &[u8], seed: u64) -> u64 {
    bytes.iter().fold(seed, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x1000_0000_01b3)
    })
}

const fn mix(mut value: u64) -> u64 {
    value ^= value >> 30;
    value = value.wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value ^= value >> 27;
    value = value.wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}

fn map_error(error: RewriteError, path: Path) -> Error {
    if let Some(amount) = error.allocation_request() {
        Error::Allocation {
            kind: LimitKind::RetainedBytes,
            amount,
        }
    } else if error.resource_limit().is_some() {
        Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
            path,
        }
    } else {
        Error::InvalidSource { path }
    }
}
