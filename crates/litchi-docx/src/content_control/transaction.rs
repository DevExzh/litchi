#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Failure-atomic source-coordinate content-control edits.

use std::collections::HashMap;
use std::fmt::Write as _;
use std::sync::Arc;

use crate::{Error, Result};

use super::snapshot::{AttributeSpan, BindingSpan, Source, Span};
use super::{
    BindingFlavor, Checksum, DataBinding, FORMATTING_ALLOWED_NAMESPACE, FormattingAllowed, Lock,
    Patch, STORE_ITEM_CHECKSUM_NAMESPACE, Snapshot,
};

const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// One typed, source-order edit.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Edit {
    /// Set, repair, or remove the checksum on one exact binding occurrence.
    SetChecksum {
        /// Source-order `w:sdtPr` ordinal.
        occurrence: usize,
        /// Canonical value, or `None` to preserve checksum absence.
        value: Option<Checksum>,
    },
    /// Set, repair, or remove the checksum on one source-order binding.
    SetBindingChecksum {
        /// Source-order `w:sdtPr` ordinal.
        occurrence: usize,
        /// Source-order binding index within that properties element.
        binding: usize,
        /// Canonical value, or `None` to preserve checksum absence.
        value: Option<Checksum>,
    },
    /// Set or remove Word's formatting exception on an applicable lock.
    SetFormattingAllowed {
        /// Source-order `w:sdtPr` ordinal.
        occurrence: usize,
        /// Semantic value, or `None` to remove the attribute.
        value: Option<FormattingAllowed>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Key {
    Checksum(usize, usize),
    Formatting(usize),
}

/// Isolated edit queue over one exact XML snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    edits: Vec<Option<Edit>>,
    index: HashMap<Key, usize>,
}

impl Transaction {
    pub(crate) fn new(base: Snapshot) -> Self {
        Self {
            base,
            edits: Vec::new(),
            index: HashMap::new(),
        }
    }

    /// Immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Coalesced edits in request order.
    pub fn edits(&self) -> impl Iterator<Item = &Edit> {
        self.edits.iter().filter_map(Option::as_ref)
    }

    /// Whether a semantic change remains queued.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.index.is_empty()
    }

    /// Apply one typed edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&mut self, edit: Edit) -> Result<&mut Self> {
        match edit {
            Edit::SetChecksum { occurrence, value } => self.set_checksum(occurrence, value),
            Edit::SetBindingChecksum {
                occurrence,
                binding,
                value,
            } => self.set_binding_checksum(occurrence, binding, value),
            Edit::SetFormattingAllowed { occurrence, value } => {
                self.set_formatting_allowed(occurrence, value)
            },
        }
    }

    /// Set, repair, or remove the core-preferred binding checksum.
    ///
    /// Use [`Self::set_binding_checksum`] when a control owns multiple
    /// bindings.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_checksum(
        &mut self,
        occurrence: usize,
        value: Option<Checksum>,
    ) -> Result<&mut Self> {
        let semantic = semantic_at(&self.base, occurrence)?;
        let bindings = semantic.control().data_bindings();
        let binding = bindings
            .iter()
            .position(|binding| binding.flavor() == BindingFlavor::Core)
            .or_else(|| (!bindings.is_empty()).then_some(0))
            .ok_or_else(|| Error::Invalid("content control has no active binding".into()))?;
        self.set_binding_checksum(occurrence, binding, value)
    }

    /// Set one exact source-order binding checksum.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_binding_checksum(
        &mut self,
        occurrence: usize,
        binding: usize,
        value: Option<Checksum>,
    ) -> Result<&mut Self> {
        let (_, semantic_binding) = exact_binding_at(&self.base, occurrence, binding)?;
        let current = semantic_binding.checksum();
        let unchanged = match (&value, current) {
            (None, None) => semantic_binding.checksum_value().is_none(),
            (Some(next), Some(current)) => next.as_bytes() == current.as_bytes(),
            _ => false,
        };
        replace_edit(
            &mut self.edits,
            &mut self.index,
            Key::Checksum(occurrence, binding),
            (!unchanged).then_some(Edit::SetBindingChecksum {
                occurrence,
                binding,
                value,
            }),
        )?;
        Ok(self)
    }

    /// Alias emphasizing the stable source-order binding coordinate.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_checksum_at(
        &mut self,
        occurrence: usize,
        binding: usize,
        value: Option<Checksum>,
    ) -> Result<&mut Self> {
        self.set_binding_checksum(occurrence, binding, value)
    }

    pub(crate) fn validate_checksum_target(&self, occurrence: usize, binding: usize) -> Result<()> {
        exact_binding_at(&self.base, occurrence, binding).map(|_| ())
    }

    pub(crate) fn try_reserve_edits(&mut self, additional: usize) -> Result<()> {
        self.edits
            .try_reserve(additional)
            .map_err(alloc("content-control edits"))?;
        self.index
            .try_reserve(additional)
            .map_err(alloc("content-control edit index"))
    }

    /// Set or remove the formatting exception on one exact active lock.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_formatting_allowed(
        &mut self,
        occurrence: usize,
        value: Option<FormattingAllowed>,
    ) -> Result<&mut Self> {
        let source = occurrence_at(&self.base, occurrence)?;
        if source.locks().len() != 1 || source.locks()[0].formatting_count() > 1 {
            return Err(Error::Invalid(
                "formatting edits require exactly one unambiguous active lock".into(),
            ));
        }
        let semantic = semantic_at(&self.base, occurrence)?;
        if value.is_some()
            && !matches!(
                semantic.control().lock(),
                Lock::ContentLocked | Lock::SdtContentLocked
            )
        {
            return Err(Error::Invalid(
                "formattingAllowed applies only to contentLocked or sdtContentLocked".into(),
            ));
        }
        replace_edit(
            &mut self.edits,
            &mut self.index,
            Key::Formatting(occurrence),
            (value != semantic.control().formatting_allowed())
                .then_some(Edit::SetFormattingAllowed { occurrence, value }),
        )?;
        Ok(self)
    }

    /// Materialize and fully reparse a candidate without mutating the source.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(&self) -> Result<Commit> {
        let before = self.base.source_owner();
        if self.index.is_empty() {
            return Ok(Commit::new(
                self.base.clone(),
                Patch::new(before.clone(), before, self.base.limits().clone()),
            ));
        }
        let mut splices = Vec::new();
        splices
            .try_reserve(self.index.len().saturating_mul(2))
            .map_err(alloc("content-control XML splices"))?;
        let mut prefixes = PrefixAllocator::new(self.base.source());
        for edit in self.edits() {
            plan(&self.base, edit, &mut splices, &mut prefixes)?;
        }
        let output = apply_splices(
            self.base.source(),
            &mut splices,
            self.base.limits().max_output_bytes,
        )?;
        let after = Source::Package(Arc::new(output));
        let snapshot = Snapshot::from_source(after.clone(), self.base.limits().clone())?;
        Ok(Commit::new(
            snapshot,
            Patch::new(before, after, self.base.limits().clone()),
        ))
    }
}

/// A fully reparsed candidate and reversible patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(crate) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// Whether any source byte changes.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Fully reparsed candidate.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split the publication values.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

#[derive(Debug)]
struct Splice {
    span: Span,
    replacement: Vec<u8>,
}

fn plan(
    snapshot: &Snapshot,
    edit: &Edit,
    splices: &mut Vec<Splice>,
    prefixes: &mut PrefixAllocator,
) -> Result<()> {
    match edit {
        Edit::SetChecksum { occurrence, value } => {
            let semantic = semantic_at(snapshot, *occurrence)?;
            let bindings = semantic.control().data_bindings();
            let binding = bindings
                .iter()
                .position(|binding| binding.flavor() == BindingFlavor::Core)
                .or_else(|| (!bindings.is_empty()).then_some(0))
                .ok_or_else(|| Error::Invalid("content control has no active binding".into()))?;
            plan_checksum(
                snapshot,
                *occurrence,
                binding,
                value.as_ref(),
                splices,
                prefixes,
            )
        },
        Edit::SetBindingChecksum {
            occurrence,
            binding,
            value,
        } => plan_checksum(
            snapshot,
            *occurrence,
            *binding,
            value.as_ref(),
            splices,
            prefixes,
        ),
        Edit::SetFormattingAllowed { occurrence, value } => {
            let source = occurrence_at(snapshot, *occurrence)?;
            let lock = source
                .locks()
                .first()
                .ok_or_else(|| Error::Invalid("content control has no source lock".into()))?;
            let lexical = value.map(|value| if value.as_bool() { "1" } else { "0" });
            replace_extension(
                snapshot.source(),
                lock.start_tag(),
                lock.formatting_allowed(),
                lock.ignorable,
                lock.ignorable_count,
                "formattingAllowed",
                FORMATTING_ALLOWED_NAMESPACE,
                PrefixKind::Formatting,
                lexical,
                splices,
                prefixes,
            )
        },
    }
}

fn plan_checksum(
    snapshot: &Snapshot,
    occurrence: usize,
    binding: usize,
    value: Option<&Checksum>,
    splices: &mut Vec<Splice>,
    prefixes: &mut PrefixAllocator,
) -> Result<()> {
    let (binding, _) = exact_binding_at(snapshot, occurrence, binding)?;
    replace_extension(
        snapshot.source(),
        binding.start_tag(),
        binding.checksum(),
        binding.ignorable,
        binding.ignorable_count,
        "storeItemChecksum",
        STORE_ITEM_CHECKSUM_NAMESPACE,
        PrefixKind::Checksum,
        value.map(Checksum::to_base64).as_deref(),
        splices,
        prefixes,
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "signature mirrors the corresponding OOXML record"
)]
fn replace_extension(
    source: &[u8],
    tag: Span,
    current: Option<AttributeSpan>,
    ignorable: Option<AttributeSpan>,
    ignorable_count: usize,
    local: &str,
    namespace: &str,
    prefix_kind: PrefixKind,
    value: Option<&str>,
    splices: &mut Vec<Splice>,
    prefixes: &mut PrefixAllocator,
) -> Result<()> {
    match (current, value) {
        (Some(current), Some(value)) => splices.push(Splice {
            span: current.value,
            replacement: escape_attribute(value)?,
        }),
        (Some(current), None) => splices.push(Splice {
            span: current.attribute,
            replacement: Vec::new(),
        }),
        (None, None) => {},
        (None, Some(value)) => {
            if ignorable_count > 1 {
                return Err(Error::Invalid(
                    "content-control element has ambiguous mc:Ignorable ownership".into(),
                ));
            }
            let extension_prefix = prefixes.prefix(prefix_kind)?;
            let insertion = tag_insertion(source, tag)?;
            let escaped = String::from_utf8(escape_attribute(value)?).map_err(|_source_error| {
                Error::Invalid("escaped content-control attribute is not UTF-8".into())
            })?;
            let authored = if let Some(ignorable) = ignorable {
                let separator = if ignorable.value.len() == 0 { "" } else { " " };
                splices.push(Splice {
                    span: Span::new(ignorable.value.end(), ignorable.value.end())?,
                    replacement: format!("{separator}{extension_prefix}").into_bytes(),
                });
                format!(
                    " xmlns:{extension_prefix}=\"{namespace}\" {extension_prefix}:{local}=\"{escaped}\""
                )
                .into_bytes()
            } else {
                let mc_prefix = prefixes.prefix(PrefixKind::Mce)?;
                format!(
                    " xmlns:{extension_prefix}=\"{namespace}\" xmlns:{mc_prefix}=\"{MCE_NAMESPACE}\" {mc_prefix}:Ignorable=\"{extension_prefix}\" {extension_prefix}:{local}=\"{escaped}\""
                )
                .into_bytes()
            };
            splices.push(Splice {
                span: Span::new(insertion, insertion)?,
                replacement: authored,
            });
        },
    }
    Ok(())
}

const PREFIX_CANDIDATES: usize = 1_025;

#[derive(Debug, Clone, Copy)]
enum PrefixKind {
    Checksum,
    Formatting,
    Mce,
}

#[derive(Debug)]
struct PrefixSlot {
    used: [bool; PREFIX_CANDIDATES],
    chosen: Option<usize>,
}

impl PrefixSlot {
    const fn new() -> Self {
        Self {
            used: [false; PREFIX_CANDIDATES],
            chosen: None,
        }
    }
}

#[derive(Debug)]
struct PrefixAllocator {
    checksum: PrefixSlot,
    formatting: PrefixSlot,
    mce: PrefixSlot,
}

impl PrefixAllocator {
    fn new(source: &[u8]) -> Self {
        let mut value = Self {
            checksum: PrefixSlot::new(),
            formatting: PrefixSlot::new(),
            mce: PrefixSlot::new(),
        };
        let mut cursor = 0usize;
        while let Some(relative) = source[cursor..]
            .windows(b"xmlns:".len())
            .position(|window| window == b"xmlns:")
        {
            let start = cursor + relative + b"xmlns:".len();
            let mut end = start;
            while source.get(end).is_some_and(|byte| {
                byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')
            }) {
                end += 1;
            }
            value.mark(&source[start..end]);
            cursor = end.max(start.saturating_add(1));
            if cursor >= source.len() {
                break;
            }
        }
        value
    }

    fn mark(&mut self, prefix: &[u8]) {
        for (kind, stem) in [
            (PrefixKind::Checksum, b"litchiDHash".as_slice()),
            (PrefixKind::Formatting, b"litchiFormatLock".as_slice()),
            (PrefixKind::Mce, b"litchiMc".as_slice()),
        ] {
            if let Some(index) = prefix_candidate(prefix, stem) {
                self.slot_mut(kind).used[index] = true;
            }
        }
    }

    fn prefix(&mut self, kind: PrefixKind) -> Result<String> {
        let stem = match kind {
            PrefixKind::Checksum => "litchiDHash",
            PrefixKind::Formatting => "litchiFormatLock",
            PrefixKind::Mce => "litchiMc",
        };
        let slot = self.slot_mut(kind);
        let suffix = if let Some(value) = slot.chosen {
            value
        } else {
            let value = slot.used.iter().position(|used| !*used).ok_or_else(|| {
                Error::Invalid("content-control namespace prefix space is exhausted".into())
            })?;
            slot.chosen = Some(value);
            value
        };
        let mut prefix = String::new();
        prefix
            .try_reserve(stem.len().saturating_add(4))
            .map_err(alloc("content-control namespace prefix"))?;
        prefix.push_str(stem);
        if suffix != 0 {
            write!(&mut prefix, "{suffix}").map_err(|_source_error| {
                Error::Invalid("namespace prefix formatting failed".into())
            })?;
        }
        Ok(prefix)
    }

    fn slot_mut(&mut self, kind: PrefixKind) -> &mut PrefixSlot {
        match kind {
            PrefixKind::Checksum => &mut self.checksum,
            PrefixKind::Formatting => &mut self.formatting,
            PrefixKind::Mce => &mut self.mce,
        }
    }
}

fn prefix_candidate(prefix: &[u8], stem: &[u8]) -> Option<usize> {
    let suffix = prefix.strip_prefix(stem)?;
    if suffix.is_empty() {
        return Some(0);
    }
    if suffix[0] == b'0' || !suffix.iter().all(u8::is_ascii_digit) {
        return None;
    }
    let mut value = 0usize;
    for digit in suffix {
        value = value
            .checked_mul(10)?
            .checked_add(usize::from(digit - b'0'))?;
        if value >= PREFIX_CANDIDATES {
            return None;
        }
    }
    Some(value)
}

fn tag_insertion(source: &[u8], span: Span) -> Result<usize> {
    let bytes = source
        .get(span.start()..span.end())
        .ok_or_else(|| Error::Invalid("content-control start tag is stale".into()))?;
    if bytes.last() != Some(&b'>') {
        return Err(Error::Invalid(
            "content-control start tag has no closing delimiter".into(),
        ));
    }
    let relative = if bytes.get(bytes.len().saturating_sub(2)) == Some(&b'/') {
        bytes.len() - 2
    } else {
        bytes.len() - 1
    };
    span.start()
        .checked_add(relative)
        .ok_or_else(|| Error::Invalid("content-control tag offset overflow".into()))
}

fn escape_attribute(value: &str) -> Result<Vec<u8>> {
    let capacity = value
        .len()
        .checked_mul(6)
        .ok_or_else(|| Error::Invalid("content-control attribute size overflow".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve(capacity)
        .map_err(alloc("escaped content-control attribute"))?;
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

fn apply_splices(source: &[u8], splices: &mut [Splice], limit: usize) -> Result<Vec<u8>> {
    splices.sort_unstable_by_key(|splice| (splice.span.start(), splice.span.end()));
    let mut cursor = 0usize;
    let mut removed = 0usize;
    let mut added = 0usize;
    for splice in splices.iter() {
        if splice.span.end() > source.len() || splice.span.start() < cursor {
            return Err(Error::Invalid(
                "content-control edits overlap or have stale coordinates".into(),
            ));
        }
        cursor = splice.span.end();
        removed = removed
            .checked_add(splice.span.len())
            .ok_or_else(|| Error::Invalid("content-control splice size overflow".into()))?;
        added = added
            .checked_add(splice.replacement.len())
            .ok_or_else(|| Error::Invalid("content-control splice size overflow".into()))?;
    }
    let output_len = source
        .len()
        .checked_sub(removed)
        .and_then(|size| size.checked_add(added))
        .ok_or_else(|| Error::Invalid("content-control output size overflow".into()))?;
    if output_len > limit {
        return Err(Error::Invalid(format!(
            "content-control output exceeds the {limit}-byte limit"
        )));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(alloc("content-control XML output"))?;
    cursor = 0;
    for splice in splices {
        output.extend_from_slice(&source[cursor..splice.span.start()]);
        output.extend_from_slice(&splice.replacement);
        cursor = splice.span.end();
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn exact_binding_at(
    snapshot: &Snapshot,
    occurrence: usize,
    binding: usize,
) -> Result<(&BindingSpan, &DataBinding)> {
    let source = occurrence_at(snapshot, occurrence)?;
    let source_binding = source.bindings().get(binding).ok_or(Error::OutOfBounds {
        object: "content-control binding",
        index: binding,
        len: source.bindings().len(),
    })?;
    if source_binding.checksum_count() > 1 {
        return Err(Error::Invalid(
            "binding has ambiguous checksum attribute ownership".into(),
        ));
    }
    let semantic = semantic_at(snapshot, occurrence)?;
    let semantic_binding =
        semantic
            .control()
            .data_bindings()
            .get(binding)
            .ok_or(Error::OutOfBounds {
                object: "semantic content-control binding",
                index: binding,
                len: semantic.control().data_bindings().len(),
            })?;
    if source.bindings().len() != semantic.control().data_bindings().len()
        || source_binding.flavor() != semantic_binding.flavor()
    {
        return Err(Error::Invalid(
            "semantic and exact-source binding inventories are misaligned".into(),
        ));
    }
    Ok((source_binding, semantic_binding))
}

fn occurrence_at(snapshot: &Snapshot, index: usize) -> Result<&super::SourceOccurrence> {
    snapshot.occurrences().get(index).ok_or(Error::OutOfBounds {
        object: "content-control occurrence",
        index,
        len: snapshot.occurrences().len(),
    })
}

fn semantic_at(snapshot: &Snapshot, index: usize) -> Result<&super::Occurrence> {
    snapshot
        .inventory()
        .occurrences()
        .get(index)
        .ok_or(Error::OutOfBounds {
            object: "content-control semantic occurrence",
            index,
            len: snapshot.inventory().occurrences().len(),
        })
}

fn replace_edit(
    edits: &mut Vec<Option<Edit>>,
    index: &mut HashMap<Key, usize>,
    key: Key,
    edit: Option<Edit>,
) -> Result<()> {
    match (index.get(&key).copied(), edit) {
        (Some(position), Some(edit)) => {
            // Replacing an already-reserved slot is infallible and does not
            // transiently erase the prior retryable request.
            edits[position] = Some(edit);
        },
        (Some(position), None) => {
            edits[position] = None;
            index.remove(&key);
        },
        (None, Some(edit)) => {
            // Reserve every fallible owner before making either collection
            // semantically observable. Allocation failure leaves both intact.
            edits
                .try_reserve(1)
                .map_err(alloc("content-control edit queue"))?;
            index
                .try_reserve(1)
                .map_err(alloc("content-control edit index"))?;
            let position = edits.len();
            edits.push(Some(edit));
            index.insert(key, position);
        },
        (None, None) => {},
    }
    Ok(())
}

fn alloc(resource: &'static str) -> impl FnOnce(std::collections::TryReserveError) -> Error {
    move |source| Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;

    const OPEN: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtPr><w:id w:val="7"/>"#;
    const BINDING: &str = r#"<w:dataBinding w:xpath="/root/value" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/>"#;
    const CLOSE: &str = "</w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>";

    fn xml(property: &str) -> Vec<u8> {
        format!("{OPEN}{property}{CLOSE}").into_bytes()
    }

    #[test]
    fn checksum_insert_reparses_and_noop_preserves_exact_source() {
        let source = Snapshot::from_xml(xml(BINDING)).unwrap();
        let mut noop = source.edit();
        noop.set_checksum(0, None).unwrap();
        let noop = noop.commit().unwrap();
        assert!(!noop.changed());
        assert_eq!(noop.snapshot().source(), source.source());

        let checksum = Checksum::from_word_value(0x7fa1_7c35);
        let mut transaction = source.edit();
        transaction.set_checksum(0, Some(checksum.clone())).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        let output = std::str::from_utf8(commit.snapshot().source()).unwrap();
        assert!(output.contains(STORE_ITEM_CHECKSUM_NAMESPACE));
        assert!(output.contains("Ignorable=\"litchiDHash\""));
        assert!(output.contains(&format!("storeItemChecksum=\"{}\"", checksum.to_base64())));
        assert_eq!(
            commit.snapshot().inventory().occurrences()[0]
                .control()
                .data_binding()
                .unwrap()
                .checksum()
                .unwrap()
                .as_bytes(),
            checksum.as_bytes()
        );
    }

    #[test]
    fn formatting_insert_reuses_local_ignorable_without_duplicate_attribute() {
        let lock = r#"<w:lock xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:future" mc:Ignorable="future" w:val="contentLocked"/>"#;
        let source = Snapshot::from_xml(xml(lock)).unwrap();
        let mut transaction = source.edit();
        transaction
            .set_formatting_allowed(0, Some(FormattingAllowed::Allowed))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let output = std::str::from_utf8(commit.snapshot().source()).unwrap();
        assert_eq!(output.matches("Ignorable=").count(), 1);
        assert!(output.contains("Ignorable=\"future litchiFormatLock\""));
        assert!(output.contains("formattingAllowed=\"1\""));
    }

    #[test]
    fn malformed_checksum_can_be_explicitly_repaired() {
        let source = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:h="{STORE_ITEM_CHECKSUM_NAMESPACE}" mc:Ignorable="h"><w:body><w:sdt><w:sdtPr><w:dataBinding w:xpath="/x" w:storeItemID="{{00000000-0000-0000-0000-000000000001}}" h:storeItemChecksum="bad"/></w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>"#
        );
        let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
        let expected = Checksum::from_word_value(1);
        let mut transaction = snapshot.edit();
        transaction.set_checksum(0, Some(expected.clone())).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(
            std::str::from_utf8(commit.snapshot().source())
                .unwrap()
                .contains(&format!("h:storeItemChecksum=\"{}\"", expected.to_base64()))
        );
    }

    #[test]
    fn stale_patch_failure_is_retryable_and_inverse_is_fresh() {
        let source = Snapshot::from_xml(xml(BINDING)).unwrap();
        let mut transaction = source.edit();
        transaction
            .set_checksum(0, Some(Checksum::from_word_value(2)))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let stale = Snapshot::from_xml(xml("<w:lock w:val=\"contentLocked\"/>")).unwrap();
        assert!(commit.patch().apply(&stale).is_err());
        let applied = commit.patch().apply(&source).unwrap();
        assert!(commit.patch().is_applied());
        let inverse = commit.patch().inverse();
        let restored = inverse.apply(&applied).unwrap();
        assert_eq!(restored.source(), source.source());
    }

    #[test]
    fn edits_only_the_mce_selected_binding() {
        let source = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="u"><w:sdt><w:sdtPr><w:dataBinding w:xpath="/inactive" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/></w:sdtPr><w:sdtContent/></w:sdt></mc:Choice><mc:Fallback><w:sdt><w:sdtPr><w:dataBinding w:xpath="/active" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/></w:sdtPr><w:sdtContent/></w:sdt></mc:Fallback></mc:AlternateContent></w:body></w:document>"#.to_owned();
        let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
        assert_eq!(snapshot.occurrences().len(), 1);
        assert_eq!(
            snapshot.inventory().occurrences()[0]
                .control()
                .data_binding()
                .unwrap()
                .xpath(),
            "/active"
        );
        let mut transaction = snapshot.edit();
        transaction
            .set_checksum(0, Some(Checksum::from_word_value(3)))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let output = std::str::from_utf8(commit.snapshot().source()).unwrap();
        assert_eq!(output.matches("storeItemChecksum=").count(), 1);
        assert!(output.contains("w:xpath=\"/inactive\" w:storeItemID="));
    }

    #[test]
    fn refuses_formatting_on_non_content_lock_without_consuming_transaction() {
        let snapshot = Snapshot::from_xml(xml(r#"<w:lock w:val="sdtLocked"/>"#)).unwrap();
        let mut transaction = snapshot.edit();
        assert!(
            transaction
                .set_formatting_allowed(0, Some(FormattingAllowed::Allowed))
                .is_err()
        );
        assert!(!transaction.is_changed());
    }

    #[test]
    fn output_limit_failure_leaves_transaction_retryable() {
        let source = xml(BINDING);
        let limits = super::super::Limits {
            max_output_bytes: source.len(),
            ..Default::default()
        };
        let snapshot = Snapshot::from_xml_with_limits(source, limits).unwrap();
        let mut transaction = snapshot.edit();
        transaction
            .set_checksum(0, Some(Checksum::from_word_value(4)))
            .unwrap();
        assert!(transaction.commit().is_err());
        assert!(transaction.commit().is_err());
        assert!(transaction.is_changed());
    }

    #[test]
    fn standalone_properties_preserve_legal_epilog_whitespace() {
        let mut source = format!(
            r#"<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{BINDING}</w:sdtPr>"#
        )
        .into_bytes();
        source.extend_from_slice(b" \r\n\t");
        let snapshot = Snapshot::from_xml(source.clone()).unwrap();
        assert_eq!(snapshot.occurrences().len(), 1);
        assert_eq!(snapshot.source(), source);
        let mut transaction = snapshot.edit();
        transaction
            .set_checksum(0, Some(Checksum::from_word_value(5)))
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.snapshot().source().ends_with(b" \r\n\t"));
    }

    #[test]
    fn edits_only_selected_property_nested_in_alternate_content() {
        let source = format!(
            r#"<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="{MCE_NAMESPACE}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><mc:AlternateContent><mc:Choice Requires="w15"><w:dataBinding w:xpath="/active" w:storeItemID="{{00000000-0000-0000-0000-000000000001}}"/></mc:Choice><mc:Fallback><w:dataBinding w:xpath="/inactive" w:storeItemID="{{00000000-0000-0000-0000-000000000001}}"/></mc:Fallback></mc:AlternateContent></w:sdtPr>"#,
        );
        let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
        assert_eq!(snapshot.occurrences()[0].bindings().len(), 1);
        let span = snapshot.occurrences()[0].bindings()[0].start_tag();
        assert!(
            std::str::from_utf8(&snapshot.source()[span.start()..span.end()])
                .unwrap()
                .contains("/active")
        );

        let mut transaction = snapshot.edit();
        transaction
            .set_checksum(0, Some(Checksum::from_word_value(7)))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let output = std::str::from_utf8(commit.snapshot().source()).unwrap();
        assert_eq!(output.matches("storeItemChecksum=").count(), 1);
        assert!(output.find("storeItemChecksum=").unwrap() < output.find("/inactive").unwrap());
        assert!(output.contains(r#"<w:dataBinding w:xpath="/inactive""#));
    }

    #[test]
    fn formats_only_selected_fallback_lock_nested_in_alternate_content() {
        let source = format!(
            r#"<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="{MCE_NAMESPACE}" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><w:lock w:val="sdtLocked"/></mc:Choice><mc:Fallback><w:lock w:val="contentLocked"/></mc:Fallback></mc:AlternateContent></w:sdtPr>"#,
        );
        let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
        assert_eq!(snapshot.occurrences()[0].locks().len(), 1);
        assert_eq!(
            snapshot.occurrences()[0].locks()[0].lock(),
            Lock::ContentLocked
        );

        let mut transaction = snapshot.edit();
        transaction
            .set_formatting_allowed(0, Some(FormattingAllowed::Allowed))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let output = std::str::from_utf8(commit.snapshot().source()).unwrap();
        assert_eq!(output.matches("formattingAllowed=").count(), 1);
        assert!(output.contains(r#"<w:lock w:val="sdtLocked"/>"#));
        assert!(output.find("formattingAllowed=").unwrap() > output.find("mc:Fallback").unwrap());
    }

    #[test]
    fn exact_source_process_content_discovers_inherited_binding() {
        let source = format!(
            r#"<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="{MCE_NAMESPACE}" xmlns:x="urn:unsupported" mc:Ignorable="x" mc:ProcessContent="x:wrap"><x:wrap>{BINDING}</x:wrap></w:sdtPr>"#,
        );
        let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
        assert_eq!(snapshot.occurrences()[0].bindings().len(), 1);
        let span = snapshot.occurrences()[0].bindings()[0].start_tag();
        assert_eq!(
            &snapshot.source()[span.start()..span.end()],
            BINDING.as_bytes()
        );
    }

    #[test]
    fn exact_source_process_content_discovers_self_declared_wildcard_lock() {
        let source = format!(
            r#"<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="{MCE_NAMESPACE}" xmlns:x="urn:unsupported"><x:wrap mc:Ignorable="x" mc:ProcessContent="x:*"><x:inner><w:lock w:val="contentLocked"/></x:inner></x:wrap></w:sdtPr>"#,
        );
        let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
        assert_eq!(snapshot.occurrences()[0].locks().len(), 1);
        let lock = &snapshot.occurrences()[0].locks()[0];
        assert_eq!(lock.lock(), Lock::ContentLocked);
        let span = lock.start_tag();
        assert_eq!(
            &snapshot.source()[span.start()..span.end()],
            br#"<w:lock w:val="contentLocked"/>"#
        );
    }

    #[test]
    fn cached_prefix_selection_is_reused_across_many_edits() {
        let bindings = format!(
            r#"<w:sdt><w:sdtPr>{BINDING}</w:sdtPr><w:sdtContent/></w:sdt><w:sdt><w:sdtPr>{BINDING}</w:sdtPr><w:sdtContent/></w:sdt>"#,
        );
        let source = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:litchiDHash="urn:occupied"><w:body>{bindings}</w:body></w:document>"#,
        );
        let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
        let mut transaction = snapshot.edit();
        for occurrence in 0..2 {
            transaction
                .set_checksum(occurrence, Some(Checksum::from_word_value(8)))
                .unwrap();
        }
        let commit = transaction.commit().unwrap();
        let output = std::str::from_utf8(commit.snapshot().source()).unwrap();
        assert_eq!(output.matches("litchiDHash1:storeItemChecksum=").count(), 2);
        assert_eq!(
            output
                .matches(&format!(
                    r#"xmlns:litchiDHash1="{STORE_ITEM_CHECKSUM_NAMESPACE}""#
                ))
                .count(),
            2
        );
    }

    #[test]
    fn mce_output_limit_is_independent_from_authored_output_limit() {
        let source = format!(
            r#"<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">{BINDING}</w:sdtPr>"#
        )
        .into_bytes();

        let authored = super::super::Limits {
            max_output_bytes: source.len(),
            ..Default::default()
        };
        let snapshot = Snapshot::from_xml_with_limits(source.clone(), authored).unwrap();
        let mut transaction = snapshot.edit();
        transaction
            .set_checksum(0, Some(Checksum::from_word_value(6)))
            .unwrap();
        assert!(transaction.commit().is_err());

        let marked = super::super::Limits {
            max_mce_marked_bytes: source.len(),
            ..Default::default()
        };
        assert!(Snapshot::from_xml_with_limits(source.clone(), marked).is_err());

        let mce = super::super::Limits {
            max_mce_output_bytes: 1,
            ..Default::default()
        };
        assert!(Snapshot::from_xml_with_limits(source, mce).is_err());
    }

    #[test]
    fn detached_input_and_depth_stack_limits_are_admitted_before_ownership() {
        let source = xml(BINDING);
        let bytes = super::super::Limits {
            max_input_bytes: source.len() - 1,
            ..Default::default()
        };
        assert!(Snapshot::from_xml_with_limits(source.clone(), bytes).is_err());

        let depth = super::super::Limits {
            max_depth: usize::MAX,
            ..Default::default()
        };
        let snapshot = Snapshot::from_xml_with_limits(source, depth).unwrap();
        assert_eq!(snapshot.occurrences().len(), 1);
    }
}
