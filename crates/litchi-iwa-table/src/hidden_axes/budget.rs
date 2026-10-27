//! The default wire budget of a hidden-axis transaction: the Pages
//! table-lock ledger (record 0379), for leaves without a budget of their
//! own. It charges the same figures under the same ceilings, derived from
//! the source's archive ingress limits.

use litchi_iwa_archive::{ComponentCatalog, Limits, SourceCatalog};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::Archive;

use super::{Error, LimitKind, map_archive_error, map_wire_error};

/// The Pages document root type; a reopen decodes it twice. Other leaves
/// never store it, so their reopen precharge is unaffected.
const ROOT_MESSAGE_TYPE: u32 = 10_000;

fn parsed_archive_source_length(archive: &Archive) -> Result<usize, Error> {
    let Some(last) = archive.objects.last() else {
        return Ok(0);
    };
    let length = last
        .data_offset
        .checked_add(last.data_length)
        .ok_or(Error::InvalidSource)?;
    usize::try_from(length).map_err(|_| Error::InvalidSource)
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

/// The figures one hidden-axis transaction charges.
pub struct WireBudget {
    limits: WireLimits,
    physical_limits: Limits,
    total_bytes: usize,
    total_fields: usize,
    total_work: usize,
    payload_work: usize,
    payload_references: usize,
    input_bytes: usize,
    output_bytes: usize,
    entries: usize,
    total_entry_bytes: usize,
    package_bytes: usize,
    payload_bytes: usize,
    total_payload_bytes: usize,
    payload_objects: usize,
    payload_messages: usize,
    payload_items: usize,
    source_keys: [Option<(usize, usize)>; 2],
    catalog_keys: [Option<(usize, usize)>; 2],
    source_key_count: usize,
    catalog_key_count: usize,
    maximum_payload_work: usize,
    maximum_payload_references: usize,
}

impl WireBudget {
    pub fn new(physical_limits: Limits) -> Result<Self, Error> {
        let archive = physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let common = WireLimits::default();
        let input = archive
            .max_message_bytes()
            .min(common.max_input_bytes())
            .max(1);
        let fields = archive.max_header_fields().min(common.max_fields()).max(1);
        let nesting = archive
            .max_header_nesting()
            .min(common.max_nesting())
            .max(1);
        let wire_limits = common
            .with_input_bytes(input)
            .and_then(|value| value.with_fields(fields))
            .and_then(|value| {
                value.with_output_bytes(
                    archive
                        .max_archive_bytes()
                        .min(common.max_output_bytes())
                        .max(1),
                )
            })
            .and_then(|value| value.with_nesting(nesting))
            .map_err(map_wire_error)?;
        let maximum_payload_work = physical_limits
            .max_iwa_stream_bytes()
            .checked_mul(16)
            .unwrap_or(usize::MAX);
        let maximum_payload_references = archive
            .max_metadata_items()
            .checked_mul(physical_limits.max_entries().max(1))
            .unwrap_or(usize::MAX);
        Ok(Self {
            limits: wire_limits,
            physical_limits,
            total_bytes: 0,
            total_fields: 0,
            total_work: 0,
            payload_work: 0,
            payload_references: 0,
            input_bytes: 0,
            output_bytes: 0,
            entries: 0,
            total_entry_bytes: 0,
            package_bytes: 0,
            payload_bytes: 0,
            total_payload_bytes: 0,
            payload_objects: 0,
            payload_messages: 0,
            payload_items: 0,
            source_keys: [None; 2],
            catalog_keys: [None; 2],
            source_key_count: 0,
            catalog_key_count: 0,
            maximum_payload_work,
            maximum_payload_references,
        })
    }

    pub fn wire_limits(&self) -> WireLimits {
        self.limits
    }

    pub fn remaining_wire_fields(&self) -> usize {
        self.limits.max_fields().saturating_sub(self.total_fields)
    }

    pub fn remaining_wire_work(&self) -> usize {
        self.limits
            .max_rewrite_work()
            .saturating_sub(self.total_work)
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "candidate reservation needs the exact selected archive/message bounds"
    )]
    pub fn precharge_candidate_reopen(
        &mut self,
        source: &SourceCatalog,
        candidate_package_bound: usize,
        changed_component_index: usize,
        changed_entry_bound: usize,
        changed_archive_bound: usize,
        changed_object_index: usize,
        changed_message_index: usize,
        changed_message_bound: usize,
    ) -> Result<(), Error> {
        self.charge_input_bytes(candidate_package_bound)?;
        self.charge_entries(source.package().len())?;
        let changed_component = source
            .components()
            .get_index(changed_component_index)
            .ok_or(Error::InvalidSource)?;
        for entry in source.package().iter() {
            let entry_bytes = if entry.name() == changed_component.name() {
                changed_entry_bound
            } else {
                entry.data().len()
            };
            self.charge_entry_bytes(entry_bytes)?;
            self.charge_total_entry_bytes(entry_bytes)?;
            self.charge_payload_work(entry_bytes)?;
            let metadata_bytes = entry
                .raw_name()
                .len()
                .checked_add(entry.metadata().local().name().len())
                .and_then(|value| value.checked_add(entry.metadata().local().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().local().comment().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().name().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().comment().len()))
                .ok_or(Error::Limit {
                    kind: LimitKind::PackageBytes,
                    observed: u64::MAX,
                    maximum: self.physical_limits.max_input_bytes(),
                })?;
            self.charge_package_bytes(metadata_bytes)?;
            self.charge_payload_work(metadata_bytes)?;
        }

        let components = source.components();
        self.charge_payload_work(components.len())?;
        for (component_index, component) in components.iter().enumerate() {
            let archive = component.archive();
            let archive_bytes = if component_index == changed_component_index {
                changed_archive_bound
            } else {
                parsed_archive_source_length(archive)?
            };
            self.charge_archive_inventory(archive_bytes, archive)?;
            self.charge_payload_work(archive.objects.len())?;
            for (object_index, object) in archive.objects.iter().enumerate() {
                self.charge_payload_work(object.messages.len())?;
                for (message_index, (message, info)) in object
                    .messages
                    .iter()
                    .zip(&object.archive_info.message_infos)
                    .enumerate()
                {
                    let message_bytes = if component_index == changed_component_index
                        && object_index == changed_object_index
                        && message_index == changed_message_index
                    {
                        changed_message_bound
                    } else {
                        message.data.len()
                    };
                    self.charge_payload_work(message_bytes)?;
                    if message.type_ == ROOT_MESSAGE_TYPE {
                        self.charge_payload_work(message_bytes)?;
                    } else if matches!(message.type_, 2_001 | 2_022) {
                        self.charge_payload_work(message_bytes)?;
                        self.charge_payload_work(message_bytes)?;
                    }
                    self.charge_payload_work(
                        info.field_infos
                            .len()
                            .saturating_add(info.object_references.len())
                            .saturating_add(info.data_references.len()),
                    )?;
                }
            }
        }
        Ok(())
    }

    pub fn maximum_payload_references(&self) -> usize {
        self.maximum_payload_references
    }

    pub fn remaining_payload_references(&self) -> usize {
        self.maximum_payload_references
            .saturating_sub(self.payload_references)
    }

    pub fn charge_codec_report(
        &mut self,
        fields: usize,
        work_bytes: usize,
        max_depth: u32,
        references: usize,
    ) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.total_fields,
            fields,
            usize_as_u64(self.limits.max_fields()),
            LimitKind::WireFields,
        )?;
        self.charge_wire_work(work_bytes)?;
        if usize::try_from(max_depth).unwrap_or(usize::MAX) > self.limits.max_nesting() {
            return Err(Error::Limit {
                kind: LimitKind::WireNesting,
                observed: u64::from(max_depth),
                maximum: usize_as_u64(self.limits.max_nesting()),
            });
        }
        self.charge_payload_references(references)
    }

    fn charge_counter(
        counter: &mut usize,
        amount: usize,
        maximum: u64,
        kind: LimitKind,
    ) -> Result<(), Error> {
        let observed = counter.checked_add(amount).ok_or(Error::Limit {
            kind,
            observed: u64::MAX,
            maximum,
        })?;
        let observed_u64 = usize_as_u64(observed);
        if observed_u64 > maximum {
            return Err(Error::Limit {
                kind,
                observed: observed_u64,
                maximum,
            });
        }
        *counter = observed;
        Ok(())
    }

    fn charge_wire_work(&mut self, amount: usize) -> Result<(), Error> {
        let work = self.total_work.checked_add(amount).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.limits.max_rewrite_work()),
        })?;
        if work > self.limits.max_rewrite_work() {
            return Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed: usize_as_u64(work),
                maximum: usize_as_u64(self.limits.max_rewrite_work()),
            });
        }
        self.total_work = work;
        Ok(())
    }

    pub fn charge_payload_work(&mut self, amount: usize) -> Result<(), Error> {
        let work = self.payload_work.checked_add(amount).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.maximum_payload_work),
        })?;
        if work > self.maximum_payload_work {
            return Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed: usize_as_u64(work),
                maximum: usize_as_u64(self.maximum_payload_work),
            });
        }
        self.charge_wire_work(amount)?;
        self.payload_work = work;
        Ok(())
    }

    pub fn charge_payload_references(&mut self, amount: usize) -> Result<(), Error> {
        self.charge_wire_work(amount)?;
        Self::charge_counter(
            &mut self.payload_references,
            amount,
            usize_as_u64(self.maximum_payload_references),
            LimitKind::PayloadReferences,
        )
    }

    fn charge_input_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.input_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            LimitKind::InputBytes,
        )
    }

    /// Admit a caller-owned selector name before retaining its bytes.
    ///
    /// This is intentionally limited to scalar byte ceilings.  Rooted table
    /// discovery remains deferred until the edit is first used.
    pub fn admit_selector_name(&mut self, amount: usize) -> Result<(), Error> {
        self.charge_input_bytes(amount)?;
        self.charge_payload_bytes(amount)?;
        self.charge_total_payload_bytes(amount)
    }

    pub fn charge_output_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.output_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            LimitKind::OutputBytes,
        )
    }

    fn charge_entries(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.entries,
            amount,
            usize_as_u64(self.physical_limits.max_entries()),
            LimitKind::Entries,
        )
    }

    fn charge_entry_bytes(&mut self, amount: usize) -> Result<(), Error> {
        let observed = usize_as_u64(amount);
        let maximum = self.physical_limits.max_entry_bytes();
        if observed > maximum {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed,
                maximum,
            });
        }
        Ok(())
    }

    fn charge_total_entry_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.total_entry_bytes,
            amount,
            self.physical_limits.max_total_bytes(),
            LimitKind::TotalEntryBytes,
        )
    }

    fn charge_package_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.package_bytes,
            amount,
            self.physical_limits.max_input_bytes(),
            LimitKind::PackageBytes,
        )
    }

    pub fn charge_payload_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.payload_bytes,
            amount,
            usize_as_u64(self.physical_limits.max_iwa_stream_bytes()),
            LimitKind::PayloadBytes,
        )
    }

    pub fn charge_total_payload_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            &mut self.total_payload_bytes,
            amount,
            self.physical_limits.max_total_bytes(),
            LimitKind::TotalPayloadBytes,
        )
    }

    pub fn charge_payload_objects(&mut self, amount: usize) -> Result<(), Error> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_objects,
            amount,
            usize_as_u64(limits.max_objects()),
            LimitKind::PayloadObjects,
        )
    }

    pub fn charge_payload_messages(&mut self, amount: usize) -> Result<(), Error> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_messages,
            amount,
            usize_as_u64(limits.max_messages()),
            LimitKind::PayloadMessages,
        )
    }

    pub fn charge_payload_items(&mut self, amount: usize) -> Result<(), Error> {
        let limits = self
            .physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        Self::charge_counter(
            &mut self.payload_items,
            amount,
            usize_as_u64(limits.max_metadata_items()),
            LimitKind::PayloadItems,
        )
    }

    fn remember_source_key(&mut self, bytes: &[u8]) -> bool {
        let key = (bytes.as_ptr() as usize, bytes.len());
        if self.source_keys[..self.source_key_count].contains(&Some(key)) {
            return false;
        }
        if self.source_key_count < self.source_keys.len() {
            self.source_keys[self.source_key_count] = Some(key);
            self.source_key_count += 1;
        }
        true
    }

    fn remember_catalog_key(&mut self, bytes: &[u8]) -> bool {
        let key = (bytes.as_ptr() as usize, bytes.len());
        if self.catalog_keys[..self.catalog_key_count].contains(&Some(key)) {
            return false;
        }
        if self.catalog_key_count < self.catalog_keys.len() {
            self.catalog_keys[self.catalog_key_count] = Some(key);
            self.catalog_key_count += 1;
        }
        true
    }

    pub fn charge_input_source(&mut self, bytes: &[u8]) -> Result<(), Error> {
        if self.remember_source_key(bytes) {
            self.charge_input_bytes(bytes.len())?;
        }
        Ok(())
    }

    pub fn charge_source_catalog(&mut self, source: &SourceCatalog) -> Result<(), Error> {
        let source_bytes = source.source_bytes();
        self.charge_input_source(source_bytes)?;
        if !self.remember_catalog_key(source_bytes) {
            return Ok(());
        }
        let package = source.package();
        self.charge_entries(package.len())?;
        for entry in package.iter() {
            let entry_bytes =
                usize::try_from(entry.metadata().uncompressed_size()).map_err(|_| {
                    Error::Limit {
                        kind: LimitKind::EntryBytes,
                        observed: u64::MAX,
                        maximum: self.physical_limits.max_entry_bytes(),
                    }
                })?;
            // EntryBytes is a per-member ceiling. The aggregate is charged
            // separately so several members can each fit individually while
            // still being bounded by TotalEntryBytes.
            self.charge_entry_bytes(entry_bytes)?;
            self.charge_total_entry_bytes(entry_bytes)?;
            self.charge_payload_work(entry.data().len())?;
            let metadata_bytes = entry
                .raw_name()
                .len()
                .checked_add(entry.metadata().local().name().len())
                .and_then(|value| value.checked_add(entry.metadata().local().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().local().comment().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().name().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().extra().len()))
                .and_then(|value| value.checked_add(entry.metadata().central().comment().len()))
                .ok_or(Error::Limit {
                    kind: LimitKind::PackageBytes,
                    observed: u64::MAX,
                    maximum: self.physical_limits.max_input_bytes(),
                })?;
            self.charge_package_bytes(metadata_bytes)?;
            self.charge_payload_work(metadata_bytes)?;
        }
        Ok(())
    }

    /// Charge the parsed IWA storage retained by a component-only reader.
    ///
    /// A shared component catalog has already passed physical ZIP, Snappy, and
    /// neutral archive limits at its originating ingress.  The focused query
    /// still needs an operation-local accounting point for the immutable
    /// archive bytes it can inspect, while object/message/index work is
    /// charged by the rooted ownership walk itself.
    pub fn charge_component_catalog(&mut self, components: &ComponentCatalog) -> Result<(), Error> {
        self.charge_payload_work(components.len())?;
        for component in components.iter() {
            let archive_bytes = parsed_archive_source_length(component.archive())?;
            if archive_bytes > self.physical_limits.max_iwa_stream_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::PayloadBytes,
                    observed: usize_as_u64(archive_bytes),
                    maximum: usize_as_u64(self.physical_limits.max_iwa_stream_bytes()),
                });
            }
            self.charge_total_payload_bytes(archive_bytes)?;
            self.charge_payload_work(archive_bytes)?;
        }
        Ok(())
    }

    fn charge_archive_inventory(
        &mut self,
        archive_bytes: usize,
        archive: &Archive,
    ) -> Result<(), Error> {
        self.charge_payload_bytes(archive_bytes)?;
        self.charge_total_payload_bytes(archive_bytes)?;
        self.charge_payload_objects(archive.objects.len())?;
        let mut message_count = 0usize;
        let mut item_count = 0usize;
        for object in &archive.objects {
            message_count =
                message_count
                    .checked_add(object.messages.len())
                    .ok_or(Error::Limit {
                        kind: LimitKind::PayloadMessages,
                        observed: u64::MAX,
                        maximum: u64::MAX,
                    })?;
            for info in &object.archive_info.message_infos {
                item_count = item_count
                    .checked_add(1)
                    .and_then(|value| value.checked_add(info.field_infos.len()))
                    .and_then(|value| value.checked_add(info.object_references.len()))
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(Error::Limit {
                        kind: LimitKind::PayloadItems,
                        observed: u64::MAX,
                        maximum: u64::MAX,
                    })?;
            }
        }
        self.charge_payload_messages(message_count)?;
        self.charge_payload_items(item_count)?;
        let work = archive_bytes
            .checked_add(archive.objects.len())
            .and_then(|value| value.checked_add(message_count))
            .and_then(|value| value.checked_add(item_count))
            .ok_or(Error::Limit {
                kind: LimitKind::WireWork,
                observed: u64::MAX,
                maximum: usize_as_u64(self.maximum_payload_work),
            })?;
        self.charge_payload_work(work)
    }

    pub fn charge_sort_work(&mut self, length: usize) -> Result<(), Error> {
        let levels = if length <= 1 {
            0
        } else {
            (usize::BITS - (length - 1).leading_zeros()) as usize
        };
        let amount = length.checked_mul(levels).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.maximum_payload_work),
        })?;
        self.charge_payload_work(amount)
    }

    pub fn parse<'a>(&mut self, source: &'a [u8], depth: usize) -> Result<WireView<'a>, Error> {
        if depth > self.limits.max_nesting() {
            return Err(Error::Limit {
                kind: LimitKind::WireNesting,
                observed: usize_as_u64(depth),
                maximum: usize_as_u64(self.limits.max_nesting()),
            });
        }
        let bytes = self
            .total_bytes
            .checked_add(source.len())
            .ok_or(Error::Limit {
                kind: LimitKind::WireBytes,
                observed: u64::MAX,
                maximum: usize_as_u64(self.limits.max_input_bytes()),
            })?;
        if bytes > self.limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::WireBytes,
                observed: usize_as_u64(bytes),
                maximum: usize_as_u64(self.limits.max_input_bytes()),
            });
        }
        let view = WireView::parse(source, self.limits).map_err(map_wire_error)?;
        let fields = self
            .total_fields
            .checked_add(view.len())
            .ok_or(Error::Limit {
                kind: LimitKind::WireFields,
                observed: u64::MAX,
                maximum: usize_as_u64(self.limits.max_fields()),
            })?;
        if fields > self.limits.max_fields() {
            return Err(Error::Limit {
                kind: LimitKind::WireFields,
                observed: usize_as_u64(fields),
                maximum: usize_as_u64(self.limits.max_fields()),
            });
        }
        let work = source.len().checked_add(view.len()).ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_as_u64(self.limits.max_rewrite_work()),
        })?;
        self.charge_wire_work(work)?;
        self.total_bytes = bytes;
        self.total_fields = fields;
        Ok(view)
    }
}

impl super::Budget for WireBudget {
    fn wire_limits(&self) -> WireLimits {
        Self::wire_limits(self)
    }
    fn remaining_wire_fields(&self) -> usize {
        Self::remaining_wire_fields(self)
    }
    fn remaining_wire_work(&self) -> usize {
        Self::remaining_wire_work(self)
    }
    fn maximum_payload_references(&self) -> usize {
        Self::maximum_payload_references(self)
    }
    fn charge_payload_work(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_payload_work(self, amount)
    }
    fn charge_payload_references(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_payload_references(self, amount)
    }
    fn charge_payload_objects(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_payload_objects(self, amount)
    }
    fn charge_payload_items(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_payload_items(self, amount)
    }
    fn charge_payload_messages(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_payload_messages(self, amount)
    }
    fn charge_payload_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_payload_bytes(self, amount)
    }
    fn charge_total_payload_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_total_payload_bytes(self, amount)
    }
    fn charge_output_bytes(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_output_bytes(self, amount)
    }
    fn charge_codec_report(
        &mut self,
        fields: usize,
        work_bytes: usize,
        max_depth: u32,
        references: usize,
    ) -> Result<(), Error> {
        Self::charge_codec_report(self, fields, work_bytes, max_depth, references)
    }
    fn parse<'a>(&mut self, source: &'a [u8], depth: usize) -> Result<WireView<'a>, Error> {
        Self::parse(self, source, depth)
    }
    fn precharge_candidate_reopen(
        &mut self,
        source: &SourceCatalog,
        candidate_package_bound: usize,
        changed_component_index: usize,
        changed_entry_bound: usize,
        changed_archive_bound: usize,
        changed_object_index: usize,
        changed_message_index: usize,
        changed_message_bound: usize,
    ) -> Result<(), Error> {
        Self::precharge_candidate_reopen(
            self,
            source,
            candidate_package_bound,
            changed_component_index,
            changed_entry_bound,
            changed_archive_bound,
            changed_object_index,
            changed_message_index,
            changed_message_bound,
        )
    }
}
