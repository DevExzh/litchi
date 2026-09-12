//! Flat-document snapshot and transactional XML edits.

use super::codec::{
    classify_mimetype, validate_flat_document_with_budget, validate_flat_document_with_context,
};
use super::model::{Family, FlatDocument};
use crate::core::Meta;
use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits as BudgetLimits, Metadata, Reservation, Resource, ResourceLimit, Result,
};
use std::cell::Cell;
use std::io::Read;
use std::mem::size_of;
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::Path;

/// Default maximum encoded size for a generic flat OpenDocument snapshot.
pub(super) const DEFAULT_MAX_FLAT_DOCUMENT_BYTES: u64 = 256 * 1024 * 1024;
/// Hard maximum encoded size accepted by the generic flat-document reader.
pub(super) const HARD_MAX_FLAT_DOCUMENT_BYTES: u64 = 512 * 1024 * 1024;
/// Finite memory budget used by the default generic flat mutation context.
///
/// This is deliberately independent of the encoded-document limit. A single
/// mutation may retain its source while constructing a candidate and several
/// bounded XML intermediates, so its aggregate memory reservation is charged
/// against this execution budget rather than inferred from the output cap.
const DEFAULT_FLAT_MUTATION_MEMORY_BYTES: u64 = 4 * 1024 * 1024 * 1024;
const DEFAULT_FLAT_MUTATION_OBJECTS: u64 = 1_000_000;
const DEFAULT_FLAT_MUTATION_WORK: u64 = 1_000_000_000;

/// An XML allocation charged to one generic flat mutation.
///
/// The reservation is kept beside the owned `String`; callers transfer the
/// candidate reservation into the snapshot only after validation succeeds.
pub(crate) struct ChargedXml {
    pub(crate) xml: String,
    pub(crate) memory: Option<MemoryLease>,
}

impl ChargedXml {
    pub(crate) fn into_string(self) -> String {
        self.xml
    }
}

/// An outstanding memory reservation owned by a retained value.
///
/// A single `String` can have a larger allocator capacity than the requested
/// length, and parser projections contain several independent vectors and
/// strings. Reservations from one budget node chain are coalesced into one
/// token so the ledger itself does not allocate an uncharged `Vec`.
#[derive(Default)]
pub(crate) struct MemoryLease {
    reservation: Option<Reservation>,
}

impl MemoryLease {
    pub(crate) fn new(reservation: Reservation) -> Self {
        Self {
            reservation: Some(reservation),
        }
    }

    pub(crate) fn try_merge(
        &mut self,
        reservation: Reservation,
    ) -> std::result::Result<(), Reservation> {
        if let Some(existing) = &mut self.reservation {
            existing.try_merge(reservation)
        } else {
            self.reservation = Some(reservation);
            Ok(())
        }
    }

    /// Reserve bytes before growing an owned buffer and reconcile allocator
    /// over-allocation after the fallible growth succeeds.
    pub(crate) fn reserve_vec<T>(
        &mut self,
        budget: &FlatMutationBudget,
        items: &mut Vec<T>,
        additional: usize,
        resource: &'static str,
    ) -> Result<()> {
        if additional == 0 {
            return Ok(());
        }
        let required = items
            .len()
            .checked_add(additional)
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        if required <= items.capacity() {
            return Ok(());
        }
        // A Vec growth keeps the old allocation live until the new allocation
        // has succeeded. Reserve the complete requested destination while the
        // old capacity remains charged by its owning lease; this is the
        // aggregate peak, rather than only the retained delta.
        let requested_bytes = required
            .checked_mul(size_of::<T>())
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        self.reserve_bytes(budget, requested_bytes, resource)?;
        items
            .try_reserve_exact(additional)
            .map_err(|source| Error::Allocation { resource, source })?;
        let actual_capacity = items.capacity();
        let actual_bytes = actual_capacity
            .checked_mul(size_of::<T>())
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        if actual_bytes > requested_bytes {
            self.reserve_bytes(budget, actual_bytes - requested_bytes, resource)?;
        }
        Ok(())
    }

    /// Reserve a standalone byte allocation and merge it into this lease.
    pub(crate) fn reserve_bytes(
        &mut self,
        budget: &FlatMutationBudget,
        amount: usize,
        resource: &'static str,
    ) -> Result<()> {
        if amount == 0 {
            return Ok(());
        }
        let reservation = budget.reserve_bytes(amount, resource)?;
        self.try_merge(reservation).map_err(|_reservation| {
            Error::InvalidFormat(format!("{resource} reservation chain changed"))
        })
    }
}

/// Operation-local aggregate memory budget for generic flat XML mutation.
pub(crate) struct FlatMutationBudget {
    budget: Budget,
    context: ExecutionContext,
    maximum_depth: Cell<u64>,
}

impl FlatMutationBudget {
    pub(crate) fn new(context: &ExecutionContext, operation: &'static str) -> Result<Self> {
        context.check().map_err(map_execution_error)?;
        let parent = context.budget();
        let child = parent.child(
            operation,
            BudgetLimits::new(
                parent.limit(Resource::Memory),
                parent.limit(Resource::InputBytes),
                parent.limit(Resource::OutputBytes),
                parent.limit(Resource::Objects),
                parent.limit(Resource::Depth),
                parent.limit(Resource::Work),
            ),
        );
        Ok(Self {
            budget: child,
            context: context.clone(),
            maximum_depth: Cell::new(0),
        })
    }

    pub(crate) fn check(&self) -> Result<()> {
        self.context.check().map_err(map_execution_error)?;
        self.budget
            .consume(Resource::Work, 1)
            .map_err(Error::ResourceLimit)
    }

    pub(crate) fn consume_objects(&self, amount: u64) -> Result<()> {
        self.context.check().map_err(map_execution_error)?;
        self.budget
            .consume(Resource::Objects, amount)
            .map_err(Error::ResourceLimit)
    }

    /// Charge the peak parser nesting observed by one operation. Depth is a
    /// peak dimension, so returning from a nested element does not release or
    /// consume it a second time.
    pub(crate) fn observe_depth(&self, depth: usize) -> Result<()> {
        let observed = u64::try_from(depth).map_err(|_| {
            Error::InvalidFormat("flat OpenDocument XML depth exceeds platform limits".to_string())
        })?;
        let previous = self.maximum_depth.get();
        if observed <= previous {
            return Ok(());
        }
        self.context.check().map_err(map_execution_error)?;
        self.budget
            .consume(Resource::Depth, observed - previous)
            .map_err(Error::ResourceLimit)?;
        self.maximum_depth.set(observed);
        Ok(())
    }

    /// Check cancellation/work and charge one decoded XML event plus its
    /// current nesting depth before a parser allocates a projection value.
    pub(crate) fn event(&self, depth: usize) -> Result<()> {
        self.check()?;
        self.consume_objects(1)?;
        self.observe_depth(depth)
    }

    pub(crate) fn reserve_bytes(
        &self,
        amount: usize,
        resource: &'static str,
    ) -> Result<Reservation> {
        self.check()?;
        let amount = u64::try_from(amount).map_err(|_| {
            Error::InvalidFormat(format!("{resource} exceeds platform resource limits"))
        })?;
        self.budget
            .reserve(Resource::Memory, amount)
            .map_err(Error::ResourceLimit)
    }
}

/// Allocate an XML `String` after reserving its actual allocator capacity.
pub(crate) fn allocate_xml(
    budget: Option<&FlatMutationBudget>,
    length: usize,
    resource: &'static str,
) -> Result<(String, Option<MemoryLease>)> {
    if let Some(budget) = budget {
        budget.check()?;
    }
    let mut memory = budget
        .map(|budget| budget.reserve_bytes(length, resource).map(MemoryLease::new))
        .transpose()?;
    let mut output = String::new();
    if let Err(source) = output.try_reserve_exact(length) {
        drop(memory);
        return Err(Error::Allocation { resource, source });
    }
    if let Some(budget) = budget {
        let actual = output.capacity();
        if actual > length {
            let extra = budget.reserve_bytes(actual - length, resource)?;
            memory
                .as_mut()
                .expect("a budgeted XML allocation has a memory lease")
                .try_merge(extra)
                .map_err(|_reservation| {
                    Error::InvalidFormat(format!(
                        "{resource} reservation chain changed during allocation"
                    ))
                })?;
        }
    }
    Ok((output, memory))
}

impl FlatDocument {
    /// Parses the optional flat-document `office:settings` inventory.
    pub fn settings(&self) -> Result<crate::Settings> {
        crate::settings::parse_flat(self.xml())
    }

    /// Open and validate a flat `OpenDocument` XML file.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Open and validate a flat `OpenDocument` XML file with a
    /// caller-supplied execution budget.
    pub fn open_with_execution_context(
        path: impl AsRef<Path>,
        context: ExecutionContext,
    ) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_execution_context(file, DEFAULT_MAX_FLAT_DOCUMENT_BYTES, context)
    }

    /// Read and validate a flat `OpenDocument` XML stream.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limit(reader, DEFAULT_MAX_FLAT_DOCUMENT_BYTES)
    }

    /// Read and validate a flat `OpenDocument` XML stream under a finite
    /// encoded-byte limit.
    pub fn from_reader_with_limit(reader: impl Read, maximum: u64) -> Result<Self> {
        Self::from_reader_with_execution_context(reader, maximum, default_flat_execution_context()?)
    }

    /// Read and validate a flat `OpenDocument` XML stream under a finite
    /// encoded-byte limit and caller-supplied execution budget.
    pub fn from_reader_with_execution_context(
        reader: impl Read,
        maximum: u64,
        context: ExecutionContext,
    ) -> Result<Self> {
        validate_flat_limit(maximum)?;
        context.check().map_err(map_execution_error)?;
        let (bytes, memory) = read_flat_input(reader, maximum, &context)?;
        Self::from_bytes_with_execution_context_and_memory(
            bytes,
            maximum,
            context,
            Some(memory),
            true,
        )
    }

    /// Validate flat `OpenDocument` XML from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limit(bytes, DEFAULT_MAX_FLAT_DOCUMENT_BYTES)
    }

    /// Validate flat `OpenDocument` XML from owned bytes under an explicit
    /// finite encoded-byte limit.
    pub fn from_bytes_with_limit(bytes: Vec<u8>, maximum: u64) -> Result<Self> {
        Self::from_bytes_with_execution_context(bytes, maximum, default_flat_execution_context()?)
    }

    /// Validate owned flat `OpenDocument` XML under an encoded-byte limit and
    /// caller-supplied execution budget.
    ///
    /// The context's `Memory` budget is used by every generic XML mutation.
    /// Its limit is aggregate operation memory and is intentionally separate
    /// from `maximum`, which remains the encoded input and output cap.
    pub fn from_bytes_with_execution_context(
        bytes: Vec<u8>,
        maximum: u64,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_bytes_with_execution_context_and_memory(bytes, maximum, context, None, false)
    }

    fn from_bytes_with_execution_context_and_memory(
        bytes: Vec<u8>,
        maximum: u64,
        context: ExecutionContext,
        retained_memory: Option<MemoryLease>,
        input_already_admitted: bool,
    ) -> Result<Self> {
        validate_flat_limit(maximum)?;
        context.check().map_err(map_execution_error)?;
        let maximum_usize = usize::try_from(maximum).map_err(|_| {
            Error::InvalidFormat(
                "flat OpenDocument input limit exceeds platform limits".to_string(),
            )
        })?;
        let observed = u64::try_from(bytes.len()).map_err(|_| {
            Error::InvalidFormat("flat OpenDocument input exceeds platform limits".to_string())
        })?;
        if observed > maximum {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::InputBytes,
                observed,
                limit: maximum,
                scope: "flat OpenDocument input".into(),
            }));
        }
        // Keep the caller's input lease live while the owned XML is created.
        // The retained snapshot is charged separately below, so this lease
        // is released before the constructor returns.
        let input_memory = if input_already_admitted {
            None
        } else {
            Some(
                context
                    .reserve(Resource::InputBytes, observed)
                    .map_err(map_execution_error)?,
            )
        };
        let retained_capacity = bytes.capacity();
        let mut source_memory = match retained_memory {
            Some(memory) => memory,
            None => MemoryLease::new(
                context
                    .reserve(
                        Resource::Memory,
                        u64::try_from(retained_capacity).map_err(|_| {
                            Error::InvalidFormat(
                                "flat OpenDocument source exceeds platform limits".to_string(),
                            )
                        })?,
                    )
                    .map_err(map_execution_error)?,
            ),
        };
        let mimetype = crate::detect::flat_mime(&bytes)
            .ok_or_else(|| Error::InvalidFormat("invalid flat OpenDocument root".to_string()))?;
        let (family, template) = classify_mimetype(&mimetype).ok_or_else(|| {
            Error::InvalidFormat(format!("unsupported OpenDocument mimetype '{mimetype}'"))
        })?;
        if (template && !matches!(family, Family::Text))
            || matches!(family, Family::Master | Family::Web | Family::Database)
        {
            return Err(Error::InvalidFormat(format!(
                "mimetype '{mimetype}' has no standard flat OpenDocument form"
            )));
        }
        let xml = String::from_utf8(bytes).map_err(|_error| {
            Error::InvalidFormat("invalid UTF-8 in flat OpenDocument".to_string())
        })?;
        debug_assert_eq!(
            xml.capacity(),
            retained_capacity,
            "String::from_utf8 must retain the input allocation"
        );
        if xml.capacity() > retained_capacity {
            source_memory
                .try_merge(
                    context
                        .reserve(
                            Resource::Memory,
                            u64::try_from(xml.capacity() - retained_capacity).map_err(|_| {
                                Error::InvalidFormat(
                                    "flat OpenDocument source exceeds platform limits".to_string(),
                                )
                            })?,
                        )
                        .map_err(map_execution_error)?,
                )
                .map_err(|_reservation| {
                    Error::InvalidFormat(
                        "flat OpenDocument source reservation chain changed during admission"
                            .to_string(),
                    )
                })?;
        }
        context.check().map_err(map_execution_error)?;
        validate_flat_document_with_context(&xml, family, &context)?;
        drop(input_memory);
        Ok(Self {
            xml,
            family,
            template,
            mimetype,
            max_document_bytes: maximum_usize,
            execution_context: context,
            source_memory,
        })
    }

    /// Attach a caller-supplied execution context to this snapshot.
    ///
    /// The context is consulted by subsequent generic mutations. Replacing a
    /// context does not alter the immutable XML snapshot or its encoded-byte
    /// limit.
    pub fn with_execution_context(&mut self, context: ExecutionContext) -> Result<()> {
        context.check().map_err(map_execution_error)?;
        let source_memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(self.xml.capacity()).map_err(|_| {
                    Error::InvalidFormat(
                        "flat OpenDocument source exceeds platform limits".to_string(),
                    )
                })?,
            )
            .map(MemoryLease::new)
            .map_err(map_execution_error)?;
        let previous_memory = std::mem::replace(&mut self.source_memory, source_memory);
        self.execution_context = context;
        drop(previous_memory);
        Ok(())
    }

    /// Return the document family.
    pub fn family(&self) -> Family {
        self.family
    }

    /// Return whether the root uses the text-template MIME.
    pub fn is_template(&self) -> bool {
        self.template
    }

    /// Return the root `office:mimetype` value.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Return the conventional flat `OpenDocument` extension.
    pub fn extension(&self) -> &'static str {
        match self.family {
            Family::Text => {
                if self.template {
                    "fott"
                } else {
                    "fodt"
                }
            },
            Family::Spreadsheet => "fods",
            Family::Presentation => "fodp",
            Family::Drawing => "fodg",
            Family::Chart => "fodc",
            Family::Formula => "fodf",
            Family::Image => "fodi",
            Family::Master | Family::Web | Family::Database => {
                unreachable!("master and web flat documents are rejected")
            },
        }
    }

    /// Return the complete flat XML document.
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Extract common document metadata from the combined XML document.
    pub fn metadata(&self) -> Result<Metadata> {
        Meta::from_bytes(self.xml.as_bytes())?.try_extract_metadata()
    }

    /// Extract the complete format-specific metadata model.
    pub fn odf_metadata(&self) -> Result<crate::Metadata> {
        Meta::from_bytes(self.xml.as_bytes())?.odf_metadata()
    }

    /// Discover inline and inert linked images in the flat document.
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        crate::media::scan_flat(&self.xml)
    }

    /// Inspect classic forms without executing bindings, events, or external resources.
    pub fn forms(&self) -> Result<crate::form::Forms> {
        crate::form::parse_form_parts(&[(self.xml(), crate::form::Part::Flat)])
    }

    /// Inspect in-content RDFa and inline `text:meta` values in the flat XML.
    pub fn in_content_metadata(&self) -> Result<crate::ContentMetadata> {
        crate::content_metadata::parse_parts(&[(self.xml(), crate::MetadataPart::Content)])
    }

    /// Inspect inert XForms model declarations in the flat XML.
    pub fn xforms_models(&self) -> Result<Vec<crate::xforms::Model>> {
        crate::xforms::parse_models(self.xml())
    }

    /// Set or clear RDFa on one paragraph in the flat XML.
    pub fn set_paragraph_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        let scratch =
            FlatMutationBudget::new(&self.execution_context, "flat OpenDocument paragraph RDFa")?;
        let updated = crate::content_metadata::set_paragraph_rdfa_with_limit_and_budget(
            &self.xml,
            position,
            value,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::content_metadata::parse_parts_with_budget(
            &[(updated.xml.as_str(), crate::MetadataPart::Content)],
            &scratch,
        )?;
        self.commit_candidate(updated)
    }

    /// Set or clear RDFa on a uniquely named bookmark start.
    pub fn set_bookmark_rdfa(&mut self, name: &str, value: &crate::RdfaAttributes) -> Result<()> {
        let scratch =
            FlatMutationBudget::new(&self.execution_context, "flat OpenDocument bookmark RDFa")?;
        let updated = crate::content_metadata::set_bookmark_rdfa_with_limit_and_budget(
            &self.xml,
            name,
            value,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::content_metadata::parse_parts_with_budget(
            &[(updated.xml.as_str(), crate::MetadataPart::Content)],
            &scratch,
        )?;
        self.commit_candidate(updated)
    }

    /// Set or clear RDFa on one inline `text:meta` occurrence.
    pub fn set_text_meta_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        let scratch =
            FlatMutationBudget::new(&self.execution_context, "flat OpenDocument text:meta RDFa")?;
        let updated = crate::content_metadata::set_text_meta_rdfa_with_limit_and_budget(
            &self.xml,
            position,
            value,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::content_metadata::parse_parts_with_budget(
            &[(updated.xml.as_str(), crate::MetadataPart::Content)],
            &scratch,
        )?;
        self.commit_candidate(updated)
    }

    /// Insert an inline `text:meta` into one paragraph in the flat XML.
    pub fn insert_text_meta(
        &mut self,
        paragraph: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> Result<()> {
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument text:meta insertion",
        )?;
        let updated = crate::content_metadata::insert_text_meta_with_limit_and_budget(
            &self.xml,
            paragraph,
            value,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::content_metadata::parse_parts_with_budget(
            &[(updated.xml.as_str(), crate::MetadataPart::Content)],
            &scratch,
        )?;
        self.commit_candidate(updated)
    }

    /// Replace one inline `text:meta` in the flat XML.
    pub fn replace_text_meta(
        &mut self,
        position: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> Result<()> {
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument text:meta replacement",
        )?;
        let updated = crate::content_metadata::replace_text_meta_with_limit_and_budget(
            &self.xml,
            position,
            value,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::content_metadata::parse_parts_with_budget(
            &[(updated.xml.as_str(), crate::MetadataPart::Content)],
            &scratch,
        )?;
        self.commit_candidate(updated)
    }

    /// Remove one inline `text:meta` from the flat XML.
    pub fn remove_text_meta(&mut self, position: litchi_core::Position) -> Result<()> {
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument text:meta removal",
        )?;
        let updated = crate::content_metadata::remove_text_meta_with_limit_and_budget(
            &self.xml,
            position,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::content_metadata::parse_parts_with_budget(
            &[(updated.xml.as_str(), crate::MetadataPart::Content)],
            &scratch,
        )?;
        self.commit_candidate(updated)
    }

    /// Replace one inert XForms model in the flat XML.
    pub fn replace_xforms_model(
        &mut self,
        position: litchi_core::Position,
        model: &crate::xforms::Model,
    ) -> Result<()> {
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument XForms model replacement",
        )?;
        let updated = crate::xforms::replace_model_with_limit_and_budget(
            &self.xml,
            position.get(),
            model,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::xforms::validate_models_with_budget(&updated.xml, &scratch)?;
        self.commit_candidate(updated)
    }

    /// Insert an inert XForms model into the existing `office:forms` container.
    pub fn insert_xforms_model(&mut self, model: &crate::xforms::Model) -> Result<()> {
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument XForms model insertion",
        )?;
        let updated = crate::xforms::insert_model_with_limit_and_budget(
            &self.xml,
            model,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::xforms::validate_models_with_budget(&updated.xml, &scratch)?;
        self.commit_candidate(updated)
    }

    /// Remove one inert XForms model from `office:forms`.
    pub fn remove_xforms_model(&mut self, position: litchi_core::Position) -> Result<()> {
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument XForms model removal",
        )?;
        let updated = crate::xforms::remove_model_with_limit_and_budget(
            &self.xml,
            position.get(),
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        crate::xforms::validate_models_with_budget(&updated.xml, &scratch)?;
        self.commit_candidate(updated)
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::variable_declaration::Declarations> {
        crate::variable_declaration::parse_parts(&[(
            self.xml(),
            crate::variable_declaration::Part::Flat,
        )])
    }

    /// Return the unnamed fallback page layout, when one is declared.
    pub fn default_page_layout(&self) -> Result<Option<crate::page_layout::PageLayout>> {
        crate::page_layout::parse_default_page_layout(self.xml())
    }

    /// Atomically insert or replace one variable declaration container.
    ///
    /// The group must target the flat part. Formulas and cached values remain
    /// inert; this method only updates XML metadata and never evaluates fields.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &crate::variable_declaration::Group,
    ) -> Result<Option<crate::variable_declaration::Group>> {
        if group.part != crate::variable_declaration::Part::Flat {
            return Err(Error::InvalidFormat(
                "FlatDocument requires Part::Flat".to_string(),
            ));
        }
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument variable declaration replacement",
        )?;
        scratch.check()?;
        let (current, source_memory) = crate::variable_declaration::parse_parts_with_budget(
            &[(self.xml(), crate::variable_declaration::Part::Flat)],
            &scratch,
        )?;
        let old = current
            .groups
            .iter()
            .find(|candidate| candidate.scope == group.scope && candidate.kind == group.kind)
            .cloned();
        drop(current);
        drop(source_memory);
        let updated = crate::variable_declaration::set_xml_with_limit_and_budget(
            &self.xml,
            group,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        let (parsed, readback_memory) = crate::variable_declaration::parse_parts_with_budget(
            &[(
                updated.xml.as_str(),
                crate::variable_declaration::Part::Flat,
            )],
            &scratch,
        )?;
        drop(parsed);
        drop(readback_memory);
        self.commit_candidate(updated)?;
        Ok(old)
    }

    /// Atomically remove one variable declaration container.
    ///
    /// Removal fails without mutation if any remaining field references a
    /// declaration owned by the container.
    pub fn remove_variable_declaration_group(
        &mut self,
        scope: &crate::variable_declaration::Scope,
        kind: crate::variable_declaration::Kind,
    ) -> Result<Option<crate::variable_declaration::Group>> {
        let scratch = FlatMutationBudget::new(
            &self.execution_context,
            "flat OpenDocument variable declaration removal",
        )?;
        scratch.check()?;
        let (current, source_memory) = crate::variable_declaration::parse_parts_with_budget(
            &[(self.xml(), crate::variable_declaration::Part::Flat)],
            &scratch,
        )?;
        let Some(old) = current
            .groups
            .iter()
            .find(|candidate| candidate.scope == *scope && candidate.kind == kind)
            .cloned()
        else {
            return Ok(None);
        };
        drop(current);
        drop(source_memory);
        let updated = crate::variable_declaration::remove_xml_with_limit_and_budget(
            &self.xml,
            scope,
            kind,
            self.max_document_bytes,
            Some(&scratch),
        )?;
        self.validate_candidate(&updated.xml, &scratch)?;
        let (parsed, readback_memory) = crate::variable_declaration::parse_parts_with_budget(
            &[(
                updated.xml.as_str(),
                crate::variable_declaration::Part::Flat,
            )],
            &scratch,
        )?;
        drop(parsed);
        drop(readback_memory);
        self.commit_candidate(updated)?;
        Ok(Some(old))
    }

    /// Discover inert inline and linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::Object>> {
        crate::embedded::scan_flat(&self.xml)
    }

    /// Return the exact original bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.xml.as_bytes()
    }

    /// Clone the exact original bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Consume this wrapper and return the exact original bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.xml.into_bytes()
    }

    /// Save the flat document without reconstructing its XML.
    ///
    /// Unix and Windows use a same-directory temporary and atomic replacement;
    /// other targets return [`Error::Unsupported`].
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_execution_context_and_scopes(
            path,
            self.execution_context.clone(),
            "flat OpenDocument output",
            "flat OpenDocument symbolic-link or non-file destination",
        )
    }

    /// Save the flat document with a caller-supplied cancellation and resource
    /// context for staging and publication.
    ///
    /// The context is checked before staging, while bytes are written, and at
    /// the final publication fence. A cancellation observed before that fence
    /// leaves the destination untouched; once replacement has committed, a
    /// later durability failure is reported as [`Error::Committed`].
    pub fn save_with_execution_context(
        &self,
        path: impl AsRef<Path>,
        context: ExecutionContext,
    ) -> Result<()> {
        self.save_with_execution_context_and_scopes(
            path,
            context,
            "flat OpenDocument output",
            "flat OpenDocument symbolic-link or non-file destination",
        )
    }

    pub(crate) fn save_with_attached_context_and_scopes(
        &self,
        path: impl AsRef<Path>,
        output_scope: &'static str,
        destination_scope: &'static str,
    ) -> Result<()> {
        self.save_with_execution_context_and_scopes(
            path,
            self.execution_context.clone(),
            output_scope,
            destination_scope,
        )
    }

    pub(crate) fn save_with_execution_context_and_scopes(
        &self,
        path: impl AsRef<Path>,
        context: ExecutionContext,
        output_scope: &'static str,
        destination_scope: &'static str,
    ) -> Result<()> {
        crate::flat::atomic::save_with_context(
            path.as_ref(),
            self.as_bytes(),
            self.max_document_bytes,
            output_scope,
            destination_scope,
            &context,
        )
    }

    fn validate_candidate(&self, updated: &str, budget: &FlatMutationBudget) -> Result<()> {
        budget.check()?;
        if updated.len() > self.max_document_bytes {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::OutputBytes,
                observed: u64::try_from(updated.len()).unwrap_or(u64::MAX),
                limit: u64::try_from(self.max_document_bytes).unwrap_or(u64::MAX),
                scope: "flat OpenDocument edited output".into(),
            }));
        }
        let mimetype = crate::detect::flat_mime(updated.as_bytes()).ok_or_else(|| {
            Error::InvalidFormat("flat OpenDocument mutation changed its root MIME".to_string())
        })?;
        let (family, template) = classify_mimetype(&mimetype).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "unsupported OpenDocument mimetype '{mimetype}' after flat mutation"
            ))
        })?;
        if family != self.family || template != self.template {
            return Err(Error::InvalidFormat(
                "flat OpenDocument mutation changed its family or template classification"
                    .to_string(),
            ));
        }
        if mimetype != self.mimetype {
            return Err(Error::InvalidFormat(
                "flat OpenDocument mutation changed its canonical MIME".to_string(),
            ));
        }
        validate_flat_document_with_budget(updated, self.family, budget)
    }

    fn commit_candidate(&mut self, candidate: ChargedXml) -> Result<()> {
        // A cancellation arriving after readback validation still prevents
        // publication.  The candidate remains owned by this stack frame, so
        // dropping it here also releases every temporary reservation.
        self.execution_context
            .check()
            .map_err(map_execution_error)?;
        if candidate.xml == self.xml {
            // A mutator may have materialized an exact no-op candidate while
            // validating its readback.  Preserve the existing String and its
            // reservation so a no-op cannot perturb retained snapshot usage.
            drop(candidate);
            return Ok(());
        }
        let Some(memory) = candidate.memory else {
            return Err(Error::InvalidFormat(
                "flat OpenDocument candidate lacks an execution reservation".to_string(),
            ));
        };
        let previous_memory = std::mem::replace(&mut self.source_memory, memory);
        self.xml = candidate.xml;
        drop(previous_memory);
        Ok(())
    }
}

fn default_flat_execution_context() -> Result<ExecutionContext> {
    let workers = NonZeroUsize::new(1)
        .ok_or_else(|| Error::InvalidFormat("flat execution worker limit is zero".to_string()))?;
    let tasks = NonZeroUsize::new(1)
        .ok_or_else(|| Error::InvalidFormat("flat execution task limit is zero".to_string()))?;
    let bytes = NonZeroU64::new(DEFAULT_FLAT_MUTATION_MEMORY_BYTES)
        .ok_or_else(|| Error::InvalidFormat("flat execution byte limit is zero".to_string()))?;
    let limits = ExecutionLimits::new(workers, tasks, bytes, 0)
        .map_err(|error| Error::Other(format!("invalid flat execution defaults: {error}")))?;
    let (_source, cancellation) = CancellationSource::pair();
    Ok(ExecutionContext::new(
        Budget::root(
            "flat OpenDocument execution",
            BudgetLimits::new(
                DEFAULT_FLAT_MUTATION_MEMORY_BYTES,
                HARD_MAX_FLAT_DOCUMENT_BYTES,
                HARD_MAX_FLAT_DOCUMENT_BYTES.saturating_mul(4),
                DEFAULT_FLAT_MUTATION_OBJECTS,
                4_096,
                DEFAULT_FLAT_MUTATION_WORK,
            ),
        ),
        cancellation,
        limits,
    ))
}

fn map_execution_error(error: ExecutionError) -> Error {
    match error {
        ExecutionError::ResourceLimit(limit) => Error::ResourceLimit(limit),
        ExecutionError::Cancelled => {
            Error::Other("flat OpenDocument operation cancelled".to_string())
        },
        other => Error::Other(format!("flat OpenDocument execution failed: {other}")),
    }
}

fn validate_flat_limit(maximum: u64) -> Result<()> {
    if maximum == 0 || maximum > HARD_MAX_FLAT_DOCUMENT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "flat OpenDocument input limit must be between 1 and {HARD_MAX_FLAT_DOCUMENT_BYTES}"
        )));
    }
    usize::try_from(maximum).map(|_| ()).map_err(|_| {
        Error::InvalidFormat("flat OpenDocument input limit exceeds platform limits".to_string())
    })
}

pub(super) fn read_flat_input(
    reader: impl Read,
    maximum: u64,
    context: &ExecutionContext,
) -> Result<(Vec<u8>, MemoryLease)> {
    context.check().map_err(map_execution_error)?;
    let read_limit = maximum.checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("flat OpenDocument input limit overflows".to_string())
    })?;
    let reserve = usize::try_from(read_limit.min(64 * 1024)).map_err(|_| {
        Error::InvalidFormat("flat OpenDocument input exceeds platform limits".to_string())
    })?;
    let mut bytes = Vec::new();
    let initial_reservation = context
        .reserve(
            Resource::Memory,
            u64::try_from(reserve).map_err(|_| {
                Error::InvalidFormat("flat OpenDocument input exceeds platform limits".to_string())
            })?,
        )
        .map_err(map_execution_error)?;
    bytes
        .try_reserve_exact(reserve)
        .map_err(|source| Error::Allocation {
            resource: "flat OpenDocument input",
            source,
        })?;
    let mut memory = MemoryLease::new(initial_reservation);
    if bytes.capacity() > reserve {
        memory
            .try_merge(
                context
                    .reserve(
                        Resource::Memory,
                        u64::try_from(bytes.capacity() - reserve).map_err(|_| {
                            Error::InvalidFormat(
                                "flat OpenDocument input exceeds platform limits".to_string(),
                            )
                        })?,
                    )
                    .map_err(map_execution_error)?,
            )
            .map_err(|_reservation| {
                Error::InvalidFormat(
                    "flat OpenDocument input reservation chain changed during read".to_string(),
                )
            })?;
    }

    let mut reader = reader.take(read_limit);
    let mut chunk = [0_u8; 8 * 1024];
    loop {
        context.check().map_err(map_execution_error)?;
        context
            .consume(Resource::Work, 1)
            .map_err(map_execution_error)?;
        let read = reader.read(&mut chunk)?;
        if read == 0 {
            break;
        }
        let required = bytes
            .len()
            .checked_add(read)
            .ok_or_else(|| Error::InvalidFormat("flat OpenDocument input overflows".to_string()))?;
        let observed = u64::try_from(required).map_err(|_| {
            Error::InvalidFormat("flat OpenDocument input exceeds platform limits".to_string())
        })?;
        if observed > maximum {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::InputBytes,
                observed,
                limit: maximum,
                scope: "flat OpenDocument input".into(),
            }));
        }
        // Input bytes are admitted cumulatively at the point the reader
        // delivers them. This prevents an unknown-size stream from being
        // consumed into an uncharged `Vec` before a later constructor check;
        // the bounded stack chunk may contain at most one small lookahead.
        context
            .consume(
                Resource::InputBytes,
                u64::try_from(read).map_err(|_| {
                    Error::InvalidFormat(
                        "flat OpenDocument input exceeds platform limits".to_string(),
                    )
                })?,
            )
            .map_err(map_execution_error)?;
        if required > bytes.capacity() {
            // `try_reserve_exact` keeps the old allocation live until the new
            // allocation has succeeded. Keep the old lease in `memory`, and
            // reserve the complete requested destination in a temporary lease
            // before growing. The old lease is released only after the Vec has
            // switched allocations; this charges the actual old-plus-new
            // peak rather than only the retained growth delta.
            let requested_destination = required;
            let destination_reservation = context
                .reserve(
                    Resource::Memory,
                    u64::try_from(requested_destination).map_err(|_| {
                        Error::InvalidFormat(
                            "flat OpenDocument input exceeds platform limits".to_string(),
                        )
                    })?,
                )
                .map_err(map_execution_error)?;
            bytes
                .try_reserve_exact(read)
                .map_err(|source| Error::Allocation {
                    resource: "flat OpenDocument input",
                    source,
                })?;
            let actual_capacity = bytes.capacity();
            if actual_capacity < required {
                return Err(Error::InvalidFormat(
                    "flat OpenDocument input allocation is smaller than requested".to_string(),
                ));
            }
            let mut replacement = MemoryLease::new(destination_reservation);
            if actual_capacity > requested_destination {
                replacement
                    .try_merge(
                        context
                            .reserve(
                                Resource::Memory,
                                u64::try_from(actual_capacity - requested_destination).map_err(
                                    |_| {
                                        Error::InvalidFormat(
                                            "flat OpenDocument input exceeds platform limits"
                                                .to_string(),
                                        )
                                    },
                                )?,
                            )
                            .map_err(map_execution_error)?,
                    )
                    .map_err(|_reservation| {
                        Error::InvalidFormat(
                            "flat OpenDocument input reservation chain changed during read"
                                .to_string(),
                        )
                    })?;
            }
            let old_memory = std::mem::replace(&mut memory, replacement);
            drop(old_memory);
        }
        bytes.extend_from_slice(&chunk[..read]);
        context.check().map_err(map_execution_error)?;
    }
    Ok((bytes, memory))
}
