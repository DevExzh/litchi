//! Handle for Word binary (`.doc`) inputs (record 0003 §2.3, M0a step S5d).
//!
//! Every facade route that reaches the DOC reader today ends in one place:
//! detection has parsed the input bytes into an in-memory
//! [`OleFile`], and the reader builds its
//! [`litchi_doc::Document`] from that parse. The path routes read the file
//! into memory first (change 0609 kept the facade's `.doc` route eager by
//! measurement), so there is no positional DOC owner to retain, and the
//! handle has a single crate-internal constructor, `Doc::from_ole_file`.
//!
//! The handle keeps what the facade kept (the reader and the metadata it read
//! at open) and, unlike the facade, the source: the same bytes and the same
//! CFB index, sealed into a [`SharedOleFile`] without a copy or a second
//! parse. An editor is built from that source at the first edit, never at
//! open.

use std::io::{Cursor, Write};
use std::sync::{Arc, OnceLock};

use litchi_cfb::{OleFile, SharedOleFile, SharedOleFileLimits};
use litchi_core::{FileFormat, Metadata, OwnedSource, Position, ReadAt, SourceVersion};
use litchi_doc::body_text::source::SourceSnapshot;
use litchi_doc::body_text::{
    OpenOptions as BodyOpenOptions, Snapshot, TransactionLimits as BodyTransactionLimits,
};
use litchi_ole_common::property_set::PropertySetReader;

use super::LegacyOpen;
use super::signature::{self, Opened, Signature};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{IntoProps, OpenOptions, Props, Result, TargetArg};

// The DOC adapter deliberately starts with the smallest complete document
// vocabulary: root, main body and ordinary main-story paragraphs.  Keeping
// these nodes leaf-backed means the generic path/query/view surfaces never
// manufacture a Pending value for an element that the binary reader already
// knows how to answer.
pub(crate) mod children;
mod nav;
mod node;
mod spec;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use spec::SPEC;

/// An opened legacy Word (`.doc`) document.
///
/// One `Arc` wide: cloning is one atomic increment, and the handle is
/// `Send + Sync`. Reads go to the reader the open built. The handle also
/// retains the input (unless [`OpenOptions::retain_source`] was `false`), so
/// a later edit needs no second open.
#[derive(Clone)]
pub struct Doc(Arc<Opened<DocState>>);

struct DocState {
    /// The reader every M0a read goes through.  The original open installs
    /// the already-built reader; an edited body snapshot leaves this empty so
    /// a generic DOC read pays for one reader build only when it is requested.
    reader: OnceLock<Result<litchi_doc::Document>>,
    /// The OLE property-set projection, read at open as the facade reads it
    /// (record 0003 §1.6).
    metadata: Metadata,
    /// Package and reader policy retained for a lazy reader built from an
    /// edited body snapshot.
    package_limits: litchi_doc::package::Limits,
    doc_leniency: litchi_doc::Leniency,
    /// The source an editor is built from.
    source: DocSource,
}

/// The source state of a [`Doc`].
///
/// Only the in-memory state exists today, because no route opens a DOC over
/// a positional source; a positional route would add a variant here.
enum DocSource {
    /// The input bytes and the CFB index detection built, sealed together.
    /// The body snapshot is initialized at the first edit and then reused by
    /// subsequent direct sets and transaction runs.
    Retained {
        shared: Arc<SharedOleFile>,
        snapshot: OnceLock<Result<Arc<Snapshot>>>,
    },
    /// A native variable-width body edit result.  Its bytes and policy stay
    /// alive without reopening or copying the compound file at commit.
    Edited(Arc<Snapshot>),
    /// The input was released after open (`retain_source(false)`).
    Released,
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Doc>();
    assert!(size_of::<Doc>() == size_of::<usize>());
};

impl Doc {
    /// Create a canonical blank DOC through the leaf writer.  The empty body
    /// paragraph is structural document content, not a visible placeholder;
    /// it gives the native paragraph editor a real formatting run to replace.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        if options.flavor != crate::Flavor::Document {
            return Err(crate::Error::new(
                crate::Code::UnsupportedCapability,
                "DOC creation does not support template flavor",
            )
            .with_format(FileFormat::Doc));
        }
        if options.conformance != crate::Conformance::Transitional {
            return Err(crate::Error::new(
                crate::Code::UnsupportedCapability,
                "DOC creation does not support Strict conformance",
            )
            .with_format(FileFormat::Doc));
        }
        if options.metadata.is_some() {
            return Err(crate::Error::new(
                crate::Code::UnsupportedCapability,
                "DOC creation does not yet author initial metadata",
            )
            .with_format(FileFormat::Doc));
        }
        let mut builder = litchi_doc::Builder::new();
        builder.add_paragraph("")?;
        let package = builder.build()?;
        match Self::from_ole_file(package.into_ole_file(), &OpenOptions::default(), None)? {
            LegacyOpen::Opened(doc) => Ok(doc),
            LegacyOpen::PasswordRequired(_) => Err(litchi_doc::Error::PasswordRequired.into()),
        }
    }

    pub(crate) fn with_origin(&self, origin: Option<Arc<std::path::Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }

    pub(crate) fn origin(&self) -> Option<Arc<std::path::Path>> {
        self.0.origin()
    }

    /// Whether this handle retains an edit source, without reading any payload.
    pub(crate) fn retains_source(&self) -> bool {
        !matches!(self.0.source, DocSource::Released)
    }

    /// Fence a source-backed read or edit against a change made after open.
    /// Retained-source payload reads perform before/after version checks; an
    /// already released source has no external revision to fence.
    pub(crate) fn check_source(&self) -> Result<()> {
        if let DocSource::Retained { shared, .. } = &self.0.source {
            check_shared_source(shared)?;
        }
        Ok(())
    }

    /// Whether two handles point at the same retained source snapshot.  Byte
    /// equality is intentionally insufficient for transaction history: a
    /// separately opened document with identical bytes is still foreign.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        match (&self.0.source, &other.0.source) {
            (
                DocSource::Retained { shared: left, .. },
                DocSource::Retained { shared: right, .. },
            ) => Arc::ptr_eq(left, right),
            (DocSource::Edited(left), DocSource::Edited(right)) => Arc::ptr_eq(left, right),
            (DocSource::Released, DocSource::Released) => Arc::ptr_eq(&self.0, &other.0),
            _ => false,
        }
    }

    /// Opens the document from the OLE file detection already parsed.
    ///
    /// This is the facade's `DetectedFormat::Doc` arm: the DOC package adopts
    /// the parse, the reader is built (with the caller's [`DocOptions`] when
    /// they are set, and with `password` when [`Locked::unlock`] supplies
    /// one), and the metadata streams are projected, with an unreadable
    /// projection reported as empty metadata. Then the parse moves into the
    /// retained source; no byte is copied and the container is not indexed
    /// again.
    ///
    /// An encrypted document opened without a password comes back as
    /// [`LegacyOpen::PasswordRequired`] with the container unchanged.
    ///
    /// [`DocOptions`]: crate::DocOptions
    /// [`Locked::unlock`]: crate::Locked
    pub(crate) fn from_ole_file(
        ole: OleFile<Cursor<Vec<u8>>>,
        options: &OpenOptions,
        password: Option<&str>,
    ) -> Result<LegacyOpen<Self>> {
        let mut package = litchi_doc::Package::from_ole_file_with_limits(
            ole,
            options.doc.and_then(|doc| doc.limits).unwrap_or_default(),
        )?;
        let opened = if options.doc.is_none() && password.is_none() {
            package.document()
        } else {
            let mut leaf_options = litchi_doc::OpenOptions::default();
            if let Some(doc) = options.doc {
                leaf_options = leaf_options.with_leniency(doc.leniency);
            }
            if let Some(password) = password {
                leaf_options =
                    leaf_options.with_password(litchi_doc::Password::new(password.to_owned()));
            }
            package.document_with_options(leaf_options)
        };
        let reader = match opened {
            Ok(reader) => reader,
            Err(litchi_doc::Error::PasswordRequired) => {
                return Ok(LegacyOpen::PasswordRequired(Box::new(
                    package.into_ole_file(),
                )));
            },
            Err(error) => return Err(error.into()),
        };
        let package_limits = package.limits();
        let doc_leniency = options
            .doc
            .map_or(litchi_doc::Leniency::Strict, |doc| doc.leniency);
        let metadata = package
            .ole_file()
            .get_metadata()
            .map(Metadata::from)
            .unwrap_or_default();
        let source = if options.retain_source {
            DocSource::Retained {
                shared: retain(package.into_ole_file())?,
                snapshot: OnceLock::new(),
            }
        } else {
            DocSource::Released
        };
        let reader_cell = OnceLock::new();
        if reader_cell.set(Ok(reader)).is_err() {
            return Err(crate::Error::new(
                crate::Code::InternalError,
                "DOC reader initialization was repeated",
            )
            .with_format(FileFormat::Doc));
        }
        Ok(LegacyOpen::Opened(Self(Opened::new(DocState {
            reader: reader_cell,
            metadata,
            package_limits,
            doc_leniency,
            source,
        }))))
    }

    /// The same document under another signature disposition (record 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    /// All text of the document, as the facade's `Document::text` returns it.
    ///
    /// # Errors
    ///
    /// Returns a typed DOC error when a text story cannot be decoded.
    pub fn text(&self) -> Result<String> {
        let result = (|| -> Result<String> { Ok(self.reader()?.text()?) })();
        self.check_source()?;
        result
    }

    /// The document's OLE property-set metadata, read at open. Always
    /// `Some`; absent or unreadable property streams read as empty metadata,
    /// as the facade reports them.
    ///
    /// # Errors
    ///
    /// Returns a source error if the retained input changed after opening.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        self.check_source()?;
        let metadata = self.0.metadata.clone();
        self.check_source()?;
        Ok(Some(metadata))
    }

    /// Everything the reader repaired while the document was opened.
    ///
    /// # Errors
    ///
    /// Returns a source error if the retained input changed after opening.
    pub fn tolerance_report(&self) -> Result<litchi_doc::ToleranceReport> {
        let result = (|| -> Result<litchi_doc::ToleranceReport> {
            Ok(self.reader()?.tolerance_report().clone())
        })();
        self.check_source()?;
        result
    }

    /// The reader the family views read paragraphs, tables and elements
    /// from.
    pub(crate) fn reader(&self) -> Result<&litchi_doc::Document> {
        if let DocSource::Retained { shared, .. } = &self.0.source {
            shared.source_version()?;
        }
        self.0
            .reader
            .get_or_init(|| self.build_reader())
            .as_ref()
            .map_err(Clone::clone)
    }

    /// Borrow the ordinary main-story paragraph stream used by the unified
    /// body axis.  The generic DOC reader's `Document::paragraphs` includes
    /// headers, footnotes, comments and other subdocuments, so it must never
    /// be used to resolve `/body/p[N]`.
    pub(crate) fn body_paragraphs(&self) -> Result<litchi_doc::BodyParagraphs<'_>> {
        Ok(self.reader()?.body_paragraphs())
    }

    /// Collect the ordinary main-story paragraph text for a whole-document
    /// view.  Collection is local to the view request; the handle retains no
    /// second paragraph model or full-document text cache.
    pub(crate) fn body_paragraph_texts(&self) -> Result<Vec<String>> {
        let result = (|| -> Result<Vec<String>> {
            self.body_paragraphs()?
                .map(|paragraph| paragraph.map(str::to_owned).map_err(Into::into))
                .collect()
        })();
        self.check_source()?;
        result
    }

    /// Read one ordinary main-story paragraph without crossing into any DOC
    /// subdocument.  This is the identity boundary shared by read and write.
    pub(crate) fn body_paragraph_at(&self, position: Position) -> Result<Option<&str>> {
        let result =
            (|| -> Result<Option<&str>> { Ok(self.reader()?.body_paragraph_text(position)?) })();
        self.check_source()?;
        result
    }

    pub(crate) fn body_paragraph_count(&self) -> Result<usize> {
        let result = (|| -> Result<usize> { Ok(self.reader()?.body_paragraph_count()?) })();
        self.check_source()?;
        result
    }

    /// Builds the source-backed main-story editor over the retained source
    /// (L-DOC-1), without indexing the container again.
    ///
    /// Returns `Ok(None)` when the source was released at open. The editor's
    /// open runs its FIB, CLX and fingerprint checks here, at the first
    /// edit. After [`Signed::unsign`](crate::Signed::unsign) the editor is
    /// built over a copy of the input whose signature storages the signature
    /// leaf removed (`litchi_sign::cfb::Editor::clear`, record 0013), so a
    /// commit from it carries no signature.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// while the document is still signed, with a suggestion naming
    /// `Signed::unsign`, and the leaf's typed refusal otherwise.
    #[cfg_attr(
        not(test),
        expect(
            dead_code,
            reason = "the narrow equal-width source route is retained for native regression tests"
        )
    )]
    pub(crate) fn body_text_editor(&self) -> Result<Option<SourceSnapshot>> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Doc));
        }
        let DocSource::Retained { shared, .. } = &self.0.source else {
            return Ok(None);
        };
        if self.0.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                read_shared_bytes(shared)?,
                litchi_sign::cfb::Format::Doc,
            )?;
            return Ok(Some(SourceSnapshot::from_bytes(bytes)?));
        }
        Ok(Some(SourceSnapshot::from_shared(Arc::clone(shared))?))
    }

    /// Opens the variable-width body-text owner used by unified DOC writes.
    ///
    /// The narrow [`SourceSnapshot`] remains useful for source-backed
    /// equal-width callers, but paragraph writes need the leaf's full
    /// `body_text::Edit`: it can rebuild the CLX and shift dependent CPs when
    /// replacement text changes width.  This owner is created only at the
    /// first edit and retains the exact source bytes for no-op publication.
    pub(crate) fn body_text_snapshot(&self) -> Result<Option<Snapshot>> {
        if let DocSource::Retained { shared, .. } = &self.0.source {
            shared.source_version()?;
        }
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Doc));
        }
        match &self.0.source {
            DocSource::Retained { shared, snapshot } => snapshot
                .get_or_init(|| {
                    let bytes = read_shared_bytes(shared)?;
                    let bytes = if self.0.signature() == Signature::RemovalAuthorized {
                        signature::unsign_compound_file(bytes, litchi_sign::cfb::Format::Doc)?
                    } else {
                        bytes
                    };
                    Snapshot::from_bytes_with_options(
                        bytes,
                        body_open_options(self.0.package_limits),
                    )
                    .map(Arc::new)
                    .map_err(Into::into)
                })
                .clone()
                .map(|snapshot| (*snapshot).clone())
                .map(Some),
            DocSource::Edited(snapshot) => Ok(Some((**snapshot).clone())),
            DocSource::Released => Ok(None),
        }
    }

    /// Apply text to one paragraph or every paragraph matched by a scoped
    /// query.  Resolution happens against this immutable handle; all selected
    /// replacements then commit through one native body-text transaction.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<WriteCommit> {
        self.set_with_limits(target, props, &SemanticLimits::default())
    }

    /// Apply a DOC paragraph text assignment under one unified semantic
    /// resource profile. Query expansion and the native body transaction use
    /// the same profile, so a large selector cannot bypass the facade limit.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source()?;
        let result = (|| -> Result<WriteCommit> {
            let text = paragraph_text_assignment(&Props::from_props(props)?)?;
            let positions =
                resolve_paragraph_targets(self, &target, limits.max_expanded_targets())?;
            if positions.len() > limits.max_operations() {
                return Err(TransactionError::Limit {
                    kind: LimitKind::Operations,
                    observed: positions.len(),
                    maximum: limits.max_operations(),
                }
                .into());
            }
            let Some(snapshot) = self.body_text_snapshot()? else {
                return Err(crate::Error::new(
                    crate::Code::UnsupportedCapability,
                    "the DOC document was opened without retaining its source for editing",
                )
                .with_format(FileFormat::Doc)
                .with_suggestion(
                    "reopen the document without `OpenOptions::retain_source(false)`",
                ));
            };
            let mut edit = snapshot.edit_with_limits(body_transaction_limits(
                *limits,
                limits.max_expanded_targets(),
            ))?;
            let operations = positions.len();
            for position in positions {
                edit.replace_paragraph(position, &text)?;
            }
            let commit = edit.commit()?;
            let changed = commit.changed();
            let (snapshot, patch) = commit.into_parts();
            let snapshot_bytes = patch_snapshot_bytes(&patch);
            if snapshot_bytes > limits.max_snapshot_bytes() {
                return Err(TransactionError::Limit {
                    kind: LimitKind::SnapshotBytes,
                    observed: snapshot_bytes,
                    maximum: limits.max_snapshot_bytes(),
                }
                .into());
            }
            let after = if changed {
                self.with_snapshot(snapshot)
            } else {
                self.clone()
            };
            Ok(WriteCommit {
                before: self.clone(),
                after,
                patch,
                operations,
                snapshot_bytes,
            })
        })();
        self.check_source()?;
        result
    }

    /// Publish the exact compound source, preserving signatures until an
    /// explicitly authorized unsigned handle is written.
    pub(crate) fn write_to<W: Write>(&self, sink: &mut W) -> Result<u64> {
        match &self.0.source {
            DocSource::Retained { shared, .. }
                if self.0.signature() == Signature::RemovalAuthorized =>
            {
                let bytes = signature::unsign_compound_file(
                    read_shared_bytes(shared)?,
                    litchi_sign::cfb::Format::Doc,
                )?;
                sink.write_all(&bytes).map_err(crate::Error::from)?;
                Ok(u64::try_from(bytes.len()).unwrap_or(u64::MAX))
            },
            DocSource::Retained { shared, .. } => stream_shared(shared, sink),
            DocSource::Edited(snapshot) => {
                sink.write_all(snapshot.bytes())
                    .map_err(crate::Error::from)?;
                Ok(u64::try_from(snapshot.bytes().len()).unwrap_or(u64::MAX))
            },
            DocSource::Released => {
                return Err(crate::Error::new(
                    crate::Code::UnsupportedCapability,
                    "the DOC source was released before publication",
                )
                .with_format(FileFormat::Doc));
            },
        }
    }

    fn with_snapshot(&self, snapshot: Snapshot) -> Self {
        Self(
            Opened::new(DocState {
                reader: OnceLock::new(),
                metadata: self.0.metadata.clone(),
                package_limits: snapshot.package_limits(),
                doc_leniency: self.0.doc_leniency,
                source: DocSource::Edited(Arc::new(snapshot)),
            })
            .with_origin(self.origin()),
        )
    }

    fn build_reader(&self) -> Result<litchi_doc::Document> {
        let bytes = match &self.0.source {
            DocSource::Retained { shared, .. } => read_shared_bytes(shared)?,
            DocSource::Edited(snapshot) => snapshot.bytes().to_vec(),
            DocSource::Released => {
                return Err(crate::Error::new(
                    crate::Code::UnsupportedCapability,
                    "the DOC reader was released before lazy materialization",
                )
                .with_format(FileFormat::Doc));
            },
        };
        let mut package = litchi_doc::Package::from_bytes_with_options(
            &bytes,
            litchi_doc::package::PackageOpenOptions::default().with_limits(self.0.package_limits),
        )?;
        let options = litchi_doc::OpenOptions::default().with_leniency(self.0.doc_leniency);
        Ok(package.document_with_options(options)?)
    }

    /// The length of the retained input, or `None` when it was released.
    #[cfg(test)]
    fn retained_source_len(&self) -> Option<u64> {
        match &self.0.source {
            DocSource::Retained { shared, .. } => Some(shared.file_size()),
            DocSource::Edited(snapshot) => Some(snapshot.bytes().len() as u64),
            DocSource::Released => None,
        }
    }
}

/// The format-specific result retained by the generic transaction history.
pub(crate) struct WriteCommit {
    pub(crate) before: Doc,
    pub(crate) after: Doc,
    pub(crate) patch: litchi_doc::body_text::Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

fn paragraph_text_assignment(props: &Props<'_>) -> Result<String> {
    if props.len() != 1 {
        return Err(crate::Error::new(
            crate::Code::UnsupportedCapability,
            "DOC paragraph writes accept exactly one `text` property",
        )
        .with_format(FileFormat::Doc));
    }
    let Some((key, value)) = props.iter().next() else {
        unreachable!("the length check above admits one property")
    };
    if !key.eq_ignore_ascii_case("text") {
        return Err(crate::Error::new(
            crate::Code::UnsupportedCapability,
            "the requested DOC property is not writable",
        )
        .with_property(key.as_str().to_owned())
        .with_format(FileFormat::Doc));
    }
    value.as_text().map(str::to_owned).ok_or_else(|| {
        crate::Error::new(
            crate::Code::InvalidValue,
            "DOC paragraph text requires a text property value",
        )
        .with_property(key.as_str().to_owned())
        .with_format(FileFormat::Doc)
    })
}

/// The facade's body owner must not inherit the native leaf default of 256
/// staged operations.  The outer `SemanticLimits` profile is the authority
/// for unified transactions; this opening profile is deliberately broad and
/// is narrowed by `Snapshot::edit_with_limits` for each operation.
fn body_open_options(package_limits: litchi_doc::package::Limits) -> BodyOpenOptions {
    BodyOpenOptions::default()
        .with_package_limits(package_limits)
        .with_transaction_limits(
            BodyTransactionLimits::new(isize::MAX as usize, usize::MAX, usize::MAX)
                .with_max_retained_render_bytes(usize::MAX),
        )
}

fn body_transaction_limits(
    limits: SemanticLimits,
    maximum_targets: usize,
) -> BodyTransactionLimits {
    BodyTransactionLimits::new(
        limits.max_operations().min(maximum_targets),
        usize::MAX,
        usize::MAX,
    )
    .with_max_retained_render_bytes(limits.max_snapshot_bytes())
}

fn resolve_paragraph_targets(
    doc: &Doc,
    target: &TargetArg<'_>,
    maximum: usize,
) -> Result<Vec<Position>> {
    let mut positions = Vec::new();
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(doc), path)?;
        let Some(position) = node.paragraph_position() else {
            return Err(crate::Error::new(
                crate::Code::UnsupportedCapability,
                "DOC text writes target paragraph nodes",
            )
            .with_format(FileFormat::Doc));
        };
        push_position(&mut positions, position, maximum)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(doc), query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(crate::Error::new(
                crate::Code::InvalidSelector,
                "the DOC mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Doc));
        }
        for node in walk {
            let node = node?;
            let Some(position) = node.paragraph_position() else {
                return Err(crate::Error::new(
                    crate::Code::UnsupportedCapability,
                    "DOC text writes target paragraph nodes",
                )
                .with_format(FileFormat::Doc));
            };
            push_position(&mut positions, position, maximum)?;
        }
    }
    if positions.is_empty() {
        return Err(crate::Error::new(
            crate::Code::ZeroMatches,
            "the DOC mutation selector matched no paragraph",
        )
        .with_format(FileFormat::Doc));
    }
    Ok(positions)
}

fn push_position(positions: &mut Vec<Position>, position: Position, maximum: usize) -> Result<()> {
    if positions.len() >= maximum {
        return Err(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: positions.len().saturating_add(1),
            maximum,
        }
        .into());
    }
    positions
        .try_reserve(1)
        .map_err(|source| TransactionError::Allocation { source })?;
    positions.push(position);
    Ok(())
}

impl std::fmt::Debug for Doc {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Doc")
            .field(
                "source_retained",
                &matches!(
                    self.0.source,
                    DocSource::Retained { .. } | DocSource::Edited(_)
                ),
            )
            .finish_non_exhaustive()
    }
}

/// Seals a parsed in-memory OLE file as the retained source.
fn retain(ole: OleFile<Cursor<Vec<u8>>>) -> Result<Arc<SharedOleFile>> {
    let shared =
        SharedOleFile::from_owned_ole_file(ole, sealed_version()?, SharedOleFileLimits::default())?;
    Ok(Arc::new(shared))
}

/// A process-unique identity for a sealed in-memory source, minted by the
/// same counter as every other [`OwnedSource`].
fn sealed_version() -> Result<SourceVersion> {
    Ok(OwnedSource::new(Vec::new()).version()?)
}

/// Check the positional source identity without materializing the compound
/// file.  Payload reads below repeat the check after the read as well.
fn check_shared_source(shared: &SharedOleFile) -> Result<()> {
    let source = shared.source_arc();
    let expected = shared.captured_source_version();
    let observed = source.version()?;
    if observed != expected {
        return Err(litchi_core::Error::SourceChanged { expected, observed }.into());
    }
    Ok(())
}

/// Materialize a retained source only for a native snapshot or a signature
/// transformation.  Ordinary publication uses [`stream_shared`] instead.
fn read_shared_bytes(shared: &SharedOleFile) -> Result<Vec<u8>> {
    signature::compound_file_bytes(shared)
}

fn stream_shared<W: Write>(shared: &SharedOleFile, sink: &mut W) -> Result<u64> {
    let source = shared.source_arc();
    let expected = shared.captured_source_version();
    let before = source.version()?;
    if before != expected {
        return Err(litchi_core::Error::SourceChanged {
            expected,
            observed: before,
        }
        .into());
    }
    let length = source.len()?;
    let mut offset = 0_u64;
    let mut buffer = [0_u8; 64 * 1024];
    while offset < length {
        let remaining = length.saturating_sub(offset);
        let size = usize::try_from(remaining)
            .unwrap_or(buffer.len())
            .min(buffer.len());
        source.read_exact_at(offset, &mut buffer[..size])?;
        sink.write_all(&buffer[..size])
            .map_err(crate::Error::from)?;
        offset = offset.saturating_add(u64::try_from(size).unwrap_or(u64::MAX));
    }
    let after = source.version()?;
    if after != expected {
        return Err(litchi_core::Error::SourceChanged {
            expected,
            observed: after,
        }
        .into());
    }
    Ok(length)
}

/// Charge both immutable endpoint snapshots and the semantic patch payload.
/// Text is retained in the patch in addition to the endpoint bytes; fixed
/// change descriptors account for formatting/resource operations.
pub(crate) fn patch_snapshot_bytes(patch: &litchi_doc::body_text::Patch) -> usize {
    let endpoints = patch
        .before()
        .bytes()
        .len()
        .saturating_add(patch.after().bytes().len());
    let descriptors = size_of::<litchi_doc::body_text::ChangeRef<'static>>()
        .saturating_mul(patch.changes().len());
    let text = patch.changes().fold(0usize, |total, change| match change {
        litchi_doc::body_text::ChangeRef::Text { before, after, .. } => total
            .saturating_add(before.len())
            .saturating_add(after.len()),
        _ => total,
    });
    size_of::<litchi_doc::body_text::Patch>()
        .saturating_add(endpoints)
        .saturating_add(descriptors)
        .saturating_add(text)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use crate::detect::{DetectedFormat, detect_format_smart};
    use crate::{Code, DocOptions};
    use litchi_doc::{Leniency, Package};
    use std::path::PathBuf;

    fn fixture_path(name: &str) -> PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc")
            .join(name)
    }

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(fixture_path(name)).expect("DOC fixture")
    }

    /// The facade's byte route: detect, then hand the parse to the handle.
    /// A password-protected document reports the leaf's refusal, as the
    /// facade did; the unified opener turns it into `File::Locked`.
    fn open_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Result<Doc> {
        match detect_format_smart(bytes) {
            Some(DetectedFormat::Doc(ole)) => match Doc::from_ole_file(ole, options, None)? {
                LegacyOpen::Opened(doc) => Ok(doc),
                LegacyOpen::PasswordRequired(_) => Err(litchi_doc::Error::PasswordRequired.into()),
            },
            other => panic!("detected {other:?}, not DOC"),
        }
    }

    /// The facade's path route reads the file into memory and then takes the
    /// byte route (`detect_document_source_path*` returns `Bytes` for DOC).
    fn open_path(name: &str, options: &OpenOptions) -> Result<Doc> {
        open_bytes(fixture(name), options)
    }

    /// Every DOC in `test-data/ole/doc`.
    fn corpus() -> Vec<String> {
        let mut names = std::fs::read_dir(fixture_path(""))
            .unwrap()
            .map(|entry| entry.unwrap().file_name().into_string().unwrap())
            .filter(|name| name.ends_with(".doc"))
            .collect::<Vec<_>>();
        names.sort();
        assert!(names.len() > 30, "DOC corpus went missing");
        names
    }

    /// The same result, or the same typed failure.
    fn assert_same<T: PartialEq + std::fmt::Debug, E: std::fmt::Display>(
        handle: Result<T>,
        leaf: std::result::Result<T, E>,
        what: &str,
    ) {
        match (handle, leaf) {
            (Ok(handle), Ok(leaf)) => assert_eq!(handle, leaf, "{what}"),
            (Err(handle), Err(leaf)) => assert_eq!(handle.message(), leaf.to_string(), "{what}"),
            (handle, leaf) => panic!(
                "{what}: handle {:?} but leaf {:?}",
                handle.map_err(|error| error.message().to_owned()),
                leaf.map_err(|error| error.to_string())
            ),
        }
    }

    /// The facade's leaf pipeline, as an independent oracle, against both
    /// routes, over the whole corpus: equal reads where the leaf reads, and
    /// the same typed failure where it does not.
    #[test]
    fn path_and_byte_routes_match_the_leaf_reader() {
        let mut opened = 0_usize;
        for name in corpus() {
            let mut package = match Package::from_reader(Cursor::new(fixture(&name))) {
                Ok(package) => package,
                Err(_) => continue,
            };
            let document = package.document();
            let metadata = package
                .ole_file()
                .get_metadata()
                .map(Metadata::from)
                .unwrap_or_default();
            for route in [
                open_path(&name, &OpenOptions::default()),
                open_bytes(fixture(&name), &OpenOptions::default()),
            ] {
                let (doc, document) = match (route, &document) {
                    (Ok(doc), Ok(document)) => (doc, document),
                    (Err(error), Err(leaf)) => {
                        assert_eq!(error.message(), leaf.to_string(), "{name}");
                        assert!(matches!(error.leaf(), crate::error::Source::Doc(_)));
                        continue;
                    },
                    (route, leaf) => panic!(
                        "{name}: handle open {} but leaf open {}",
                        route.is_ok(),
                        leaf.is_ok()
                    ),
                };
                opened += 1;
                assert_same(doc.text(), document.text(), &name);
                assert_eq!(
                    format!("{:?}", doc.metadata().unwrap()),
                    format!("{:?}", Some(&metadata)),
                    "{name}"
                );
                assert_eq!(
                    &doc.tolerance_report().unwrap(),
                    document.tolerance_report(),
                    "{name}"
                );
                assert_eq!(
                    doc.reader()
                        .ok()
                        .and_then(|reader| reader.paragraph_count().ok()),
                    document.paragraph_count().ok(),
                    "{name}"
                );
            }
        }
        assert!(opened > 40, "too few DOC fixtures opened: {opened}");
    }

    #[test]
    fn known_metadata_survives_the_handle() {
        let doc = open_path("documentProperties.doc", &OpenOptions::default()).unwrap();
        let metadata = doc
            .metadata()
            .unwrap()
            .expect("DOC metadata is always reported");
        assert!(metadata.has_data());
    }

    #[test]
    fn doc_options_reach_the_reader() {
        let options = OpenOptions::default()
            .doc(DocOptions::new().leniency(Leniency::TolerateStylesheetDefects));
        let name = "duplicate-style-names.doc";
        let mut package = Package::from_reader(Cursor::new(fixture(name))).unwrap();
        let document = package
            .document_with_options(
                litchi_doc::OpenOptions::default()
                    .with_leniency(Leniency::TolerateStylesheetDefects),
            )
            .unwrap();
        let doc = open_path(name, &options).unwrap();
        assert_eq!(doc.text().unwrap(), document.text().unwrap());
        assert_eq!(
            &doc.tolerance_report().unwrap(),
            document.tolerance_report()
        );
    }

    #[test]
    fn encrypted_input_keeps_todays_typed_refusal() {
        let error = match open_path("PasswordProtected.doc", &OpenOptions::default()) {
            Ok(_) => panic!("an encrypted DOC opened without a password"),
            Err(error) => error,
        };
        assert_eq!(error.code(), Code::DocumentProtected);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Doc(litchi_doc::Error::PasswordRequired)
        ));
    }

    #[test]
    fn the_source_is_retained_without_a_copy_and_released_on_request() {
        let bytes = fixture("documentProperties.doc");
        let length = u64::try_from(bytes.len()).unwrap();
        let retained = open_bytes(bytes.clone(), &OpenOptions::default()).unwrap();
        // Risk R5: retention keeps exactly the input alive, nothing more.
        assert_eq!(retained.retained_source_len(), Some(length));

        let released = open_bytes(bytes, &OpenOptions::default().retain_source(false)).unwrap();
        assert_eq!(released.retained_source_len(), None);
        assert!(released.body_text_editor().unwrap().is_none());
        assert_eq!(released.text().unwrap(), retained.text().unwrap());
    }

    #[test]
    fn the_editor_is_built_from_the_retained_source() {
        let bytes = fixture("documentProperties.doc");
        let doc = open_bytes(bytes.clone(), &OpenOptions::default()).unwrap();
        let editor = doc.body_text_editor().unwrap().expect("retained source");
        let oracle = SourceSnapshot::from_bytes(bytes).unwrap();
        assert_eq!(editor.fingerprint(), oracle.fingerprint());
        let first = Position::new(0);
        assert_eq!(
            editor.paragraph(first).unwrap().text(),
            oracle.paragraph(first).unwrap().text()
        );
        // The editor is built over the retained source, not over a copy:
        // the sealed source's process-unique identity carries over.
        let DocSource::Retained { shared, .. } = &doc.0.source else {
            panic!("source not retained");
        };
        assert_eq!(editor.source_version(), shared.captured_source_version());
        assert_ne!(editor.source_version(), oracle.source_version());
    }

    #[test]
    fn handles_are_cheap_shared_clones() {
        let doc = open_path("documentProperties.doc", &OpenOptions::default()).unwrap();
        let clone = doc.clone();
        assert!(Arc::ptr_eq(&doc.0, &clone.0));
        std::thread::scope(|scope| {
            scope.spawn(|| assert_eq!(clone.text().unwrap(), doc.text().unwrap()));
        });
    }
}
