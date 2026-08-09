use std::fmt;
use std::sync::Arc;

use crate::escape;

use super::model::{Dialect, Error, ReadLimits, Snapshot};

#[derive(Clone, Debug)]
enum Intent {
    Replace {
        position: usize,
        replacement: String,
    },
    Append {
        block: String,
    },
}

/// A bounded one-operation edit against an immutable Markdown snapshot.
#[derive(Debug)]
pub struct Edit<'snapshot> {
    source: &'snapshot Snapshot,
    intent: Option<Intent>,
}

impl<'snapshot> Edit<'snapshot> {
    pub(crate) const fn new(source: &'snapshot Snapshot) -> Self {
        Self {
            source,
            intent: None,
        }
    }

    /// Append exactly one parsed Markdown block.
    ///
    /// The adapter inserts the minimum deterministic blank-line separator:
    /// none for an empty source, two LF bytes after non-newline-terminated
    /// source, and one LF byte after newline-terminated source.
    ///
    /// # Errors
    ///
    /// Returns a typed error if an operation is already staged or `block` does
    /// not parse as exactly one top-level block under the snapshot policy.
    pub fn append_block(&mut self, block: &str) -> Result<&mut Self, Error> {
        self.ensure_empty()?;
        validate_replacement(self.source, block)?;
        self.intent = Some(Intent::Append {
            block: copy_source(block)?,
        });
        Ok(self)
    }

    /// Append literal text as one safely escaped paragraph block.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::append_block`]. In particular, an
    /// empty value or literal blank-line sequence is refused because it cannot
    /// represent exactly one paragraph without changing the supplied text.
    pub fn append_text(&mut self, text: &str) -> Result<&mut Self, Error> {
        let escaped = escape::text(text);
        self.append_block(&escaped)
    }

    /// Atomically validate and publish the staged candidate.
    ///
    /// # Errors
    ///
    /// Returns a typed error without modifying the source when no operation is
    /// staged or the complete candidate violates its retained read policy.
    pub fn commit(self) -> Result<Commit, Error> {
        let intent = self.intent.ok_or(Error::NoStagedOperation)?;
        let target = render(self.source, &intent)?;
        publish(self.source, &target)
    }

    fn ensure_empty(&self) -> Result<(), Error> {
        if self.intent.is_some() {
            return Err(Error::OperationAlreadyStaged);
        }
        Ok(())
    }

    /// Remove one selected top-level block without normalizing surrounding
    /// whitespace.
    ///
    /// # Errors
    ///
    /// Returns a typed missing-position or already-staged error.
    pub fn remove_block(&mut self, position: usize) -> Result<&mut Self, Error> {
        self.replace_block(position, "")
    }

    /// Replace one selected top-level block with zero or one parsed block.
    ///
    /// An empty replacement removes the selected block. Nonempty input must
    /// parse as exactly one top-level block, including a link definition.
    /// Untouched source bytes, including surrounding blank lines, remain exact.
    ///
    /// # Errors
    ///
    /// Returns a typed missing-position, already-staged, input, allocation, or
    /// resource-limit error.
    pub fn replace_block(
        &mut self,
        position: usize,
        replacement: &str,
    ) -> Result<&mut Self, Error> {
        self.ensure_empty()?;
        if self.source.block(position).is_none() {
            return Err(Error::BlockNotFound { position });
        }
        if !replacement.is_empty() {
            validate_replacement(self.source, replacement)?;
        }
        self.intent = Some(Intent::Replace {
            position,
            replacement: copy_source(replacement)?,
        });
        Ok(self)
    }

    /// Replace one block with one safely escaped literal paragraph.
    ///
    /// Markdown delimiters in `text` cannot become active syntax. Use
    /// [`Self::replace_block`] when the replacement is intentionally Markdown.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::replace_block`]. Empty text and a
    /// literal blank-line sequence are refused rather than treated as removal
    /// or silently collapsed; use [`Self::remove_block`] for removal.
    pub fn replace_block_with_text(
        &mut self,
        position: usize,
        text: &str,
    ) -> Result<&mut Self, Error> {
        let escaped = escape::text(text);
        if escaped.is_empty() {
            return Err(Error::ReplacementBlockCount { actual: 0 });
        }
        self.replace_block(position, &escaped)
    }
}

/// Diagnostics for one Markdown publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_blocks: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Whether the exact source changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Whether the changed candidate required full parsing.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }

    /// Number of semantic blocks directly targeted by the operation.
    #[must_use]
    pub const fn touched_blocks(self) -> usize {
        self.touched_blocks
    }
}

/// A validated Markdown snapshot, reversible patch, and diagnostics.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Newly validated immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Consume this commit into its snapshot, patch, and diagnostics.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, Diagnostics) {
        (self.snapshot, self.patch, self.diagnostics)
    }
}

/// An in-memory reversible patch guarded by an exact complete before-image.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    before: Arc<str>,
    after: Arc<str>,
    dialect: Dialect,
    limits: ReadLimits,
    source_fingerprint: u64,
    target_fingerprint: u64,
}

impl Patch {
    /// Return an exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            dialect: self.dialect,
            limits: self.limits,
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
        }
    }

    /// Whether the before- and after-images are byte-identical.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Deterministic non-cryptographic fingerprint of the exact before-image.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Deterministic non-cryptographic fingerprint of the exact after-image.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("changed", &!self.is_empty())
            .field("dialect", &self.dialect)
            .finish_non_exhaustive()
    }
}

pub(crate) fn apply(source: &Snapshot, patch: &Patch) -> Result<Commit, Error> {
    if source.state.source != patch.before
        || source.dialect() != patch.dialect
        || source.limits() != patch.limits
    {
        return Err(Error::PatchConflict);
    }
    if patch.is_empty() {
        return Ok(Commit {
            snapshot: source.clone(),
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: false,
                touched_blocks: 0,
                full_reparse_performed: false,
            },
        });
    }
    let snapshot = Snapshot::read_with(&patch.after, patch.dialect, patch.limits)?;
    Ok(Commit {
        snapshot,
        patch: patch.clone(),
        diagnostics: Diagnostics {
            changed: true,
            touched_blocks: 1,
            full_reparse_performed: true,
        },
    })
}

fn fingerprint(source: &str) -> u64 {
    const OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
    const PRIME: u64 = 0x0000_0100_0000_01b3;
    source.bytes().fold(OFFSET, |hash, byte| {
        (hash ^ u64::from(byte)).wrapping_mul(PRIME)
    })
}

fn copy_source(source: &str) -> Result<String, Error> {
    let mut retained = String::new();
    retained
        .try_reserve_exact(source.len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown staged replacement",
            source: allocation_error,
        })?;
    retained.push_str(source);
    Ok(retained)
}

fn publish(source: &Snapshot, target: &str) -> Result<Commit, Error> {
    if target == source.source() {
        let exact = Arc::clone(&source.state.source);
        let fingerprint = fingerprint(&exact);
        return Ok(Commit {
            snapshot: source.clone(),
            patch: Patch {
                before: Arc::clone(&exact),
                after: exact,
                dialect: source.dialect(),
                limits: source.limits(),
                source_fingerprint: fingerprint,
                target_fingerprint: fingerprint,
            },
            diagnostics: Diagnostics {
                changed: false,
                touched_blocks: 0,
                full_reparse_performed: false,
            },
        });
    }
    let snapshot = Snapshot::read_with(target, source.dialect(), source.limits())?;
    let before = Arc::clone(&source.state.source);
    let after = Arc::clone(&snapshot.state.source);
    let source_fingerprint = fingerprint(&before);
    let target_fingerprint = fingerprint(&after);
    Ok(Commit {
        snapshot,
        patch: Patch {
            before,
            after,
            dialect: source.dialect(),
            limits: source.limits(),
            source_fingerprint,
            target_fingerprint,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_blocks: 1,
            full_reparse_performed: true,
        },
    })
}

fn render(source: &Snapshot, intent: &Intent) -> Result<String, Error> {
    let (before, prefix, replacement, suffix, after) = match intent {
        Intent::Replace {
            position,
            replacement,
        } => {
            let block = source.block(*position).ok_or(Error::BlockNotFound {
                position: *position,
            })?;
            let range = block.range();
            let suffix = if replacement.is_empty()
                || after_is_empty(source, range.end)
                || replacement.ends_with(['\r', '\n'])
            {
                ""
            } else if block.source().ends_with("\r\n") {
                "\r\n"
            } else if block.source().ends_with(['\r', '\n']) {
                "\n"
            } else {
                ""
            };
            (
                &source.source()[..range.start],
                "",
                replacement,
                suffix,
                &source.source()[range.end..],
            )
        },
        Intent::Append { block } => {
            let separator = if source.source().is_empty() {
                ""
            } else if source.source().ends_with('\n') {
                "\n"
            } else {
                "\n\n"
            };
            (source.source(), separator, block, "", "")
        },
    };
    let capacity = before
        .len()
        .checked_add(prefix.len())
        .and_then(|length| length.checked_add(replacement.len()))
        .and_then(|length| length.checked_add(suffix.len()))
        .and_then(|length| length.checked_add(after.len()))
        .ok_or(Error::SourceTooLarge {
            actual: usize::MAX,
            limit: source.limits().max_source_bytes,
        })?;
    if capacity > source.limits().max_source_bytes {
        return Err(Error::SourceTooLarge {
            actual: capacity,
            limit: source.limits().max_source_bytes,
        });
    }
    let mut target = String::new();
    target
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown edit candidate",
            source: allocation_error,
        })?;
    target.push_str(before);
    target.push_str(prefix);
    target.push_str(replacement);
    target.push_str(suffix);
    target.push_str(after);
    Ok(target)
}

fn after_is_empty(source: &Snapshot, range_end: usize) -> bool {
    range_end == source.source().len()
}

fn validate_replacement(source: &Snapshot, replacement: &str) -> Result<(), Error> {
    let parsed = Snapshot::read_with(replacement, source.dialect(), source.limits())?;
    let actual = parsed.blocks().len();
    if actual != 1 {
        return Err(Error::ReplacementBlockCount { actual });
    }
    Ok(())
}
