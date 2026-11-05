//! Unified legacy Word handle over the native immutable DOC package.
//!
//! Detection transfers its CFB parse into the native owner. Reads borrow the
//! owner's shared document projection; edits publish a new native package.
//! Source release keeps only the detached read projection and metadata.

use std::io::{Cursor, Write};
use std::mem::size_of;
use std::sync::Arc;

use litchi_cfb::OleFile;
use litchi_core::{FileFormat, Metadata, Position};
use litchi_doc::body_text::TransactionLimits as BodyTransactionLimits;
use litchi_doc::source_backed::Package as SourcePackage;

use super::LegacyOpen;
use super::signature::{self, Opened, Signature};
use crate::transaction::SemanticLimits;
use crate::{IntoProps, OpenOptions, Result, TargetArg};

// The DOC adapter deliberately starts with the smallest complete document
// vocabulary: root, main body and ordinary main-story paragraphs.  Keeping
// these nodes leaf-backed means the generic path/query/view surfaces never
// manufacture a Pending value for an element that the binary reader already
// knows how to answer.
pub(crate) mod children;
mod edit;
mod find;
mod marks;
mod marks_write;
mod metadata;
mod nav;
mod node;
mod parts;
pub(crate) mod refresh;
mod spec;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::NodeProps;
pub(crate) use spec::{ADDABLE, REMOVABLE, SETTABLE, SPEC};

/// An opened legacy Word (`.doc`) document.
///
/// One `Arc` wide: cloning is one atomic increment, and the handle is
/// `Send + Sync`. Reads go to the reader the open built. The handle also
/// retains the input (unless [`OpenOptions::retain_source`] was `false`), so
/// a later edit needs no second open.
#[derive(Clone)]
pub struct Doc(Arc<Opened<DocState>>);

struct DocState {
    reader: litchi_doc::Document,
    metadata: Metadata,
    source: DocSource,
    /// The DOCX-style identifier of each style slot, derived on first use
    /// (record 0324).
    style_ids: std::sync::OnceLock<Box<[Box<str>]>>,
}

/// Retained native owner or explicitly released source.
enum DocSource {
    Retained(litchi_doc::Package),
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
        let mut builder = litchi_doc::Builder::new();
        if let Some(metadata) = &options.metadata {
            metadata::author(&mut builder, metadata)?;
        }
        builder.add_paragraph("")?;
        Ok(Self::from_package(builder.build()?, true))
    }

    fn from_package(package: litchi_doc::Package, retain_source: bool) -> Self {
        Self(Opened::new(DocState {
            reader: package.document(),
            metadata: package.metadata().clone(),
            style_ids: std::sync::OnceLock::new(),
            source: if retain_source {
                DocSource::Retained(package)
            } else {
                DocSource::Released
            },
        }))
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
        if let DocSource::Retained(package) = &self.0.source {
            package.source_version()?;
        }
        Ok(())
    }

    /// Check retained-source freshness for a mutation that may still prove an
    /// exact signed no-op. A changed native result is rejected at publish,
    /// after the editor establishes whether the request changed bytes.
    pub(crate) fn check_source_for_write(&self) -> Result<()> {
        self.check_source()
    }

    /// Whether two handles point at the same retained source snapshot.  Byte
    /// equality is intentionally insufficient for transaction history: a
    /// separately opened document with identical bytes is still foreign.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        match (&self.0.source, &other.0.source) {
            (DocSource::Retained(left), DocSource::Retained(right)) => {
                Arc::ptr_eq(&left.shared_source(), &right.shared_source())
            },
            (DocSource::Released, DocSource::Released) => Arc::ptr_eq(&self.0, &other.0),
            _ => false,
        }
    }

    /// Opens the document from the OLE file detection already parsed.
    ///
    /// This is the facade's `DetectedFormat::Doc` arm: the DOC package adopts
    /// the parse, the reader is built (with the caller's [`DocOptions`] when
    /// they are set, and with `password` when [`Locked::unlock`] supplies
    /// one), and metadata streams are projected under the native typed policy. The parse moves into the
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
        let mut leaf_options = litchi_doc::OpenOptions::default().source_limits(
            options
                .doc
                .and_then(|doc| doc.source_limits)
                .unwrap_or_default(),
        );
        if let Some(doc) = options.doc {
            leaf_options = leaf_options.leniency(doc.leniency);
        }
        if let Some(password) = password {
            leaf_options = leaf_options.password(litchi_doc::Password::new(password.to_owned()));
        }
        match litchi_doc::raw::open_package(ole, leaf_options)? {
            litchi_doc::raw::OpenOutcome::Opened(package) => Ok(LegacyOpen::Opened(
                Self::from_package(package, options.retain_source),
            )),
            litchi_doc::raw::OpenOutcome::PasswordRequired(ole) => {
                Ok(LegacyOpen::PasswordRequired(ole))
            },
        }
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
    /// `Some`; absent property streams read as empty metadata. Invalid streams
    /// return the native typed error during ingress.
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
        self.check_source()?;
        Ok(&self.0.reader)
    }

    /// The parsed reader, without the retained-source fence: the detached
    /// in-memory projection whose values a resolved node's `format` reads
    /// (record 0324). Every path resolution and query still goes through
    /// [`Self::reader`], which fences the source.
    pub(crate) fn parsed(&self) -> &litchi_doc::Document {
        &self.0.reader
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
        let result = (|| -> Result<Option<&str>> {
            let reader = self.reader()?;
            // The memoized story index answers in constant time; a paragraph
            // it cannot vouch for (a trailing fragment, an empty story) falls
            // back to the borrowed body scan.
            if let Some(text) = reader.body_blocks()?.paragraph_text(position) {
                return Ok(Some(text));
            }
            Ok(reader.body_paragraph_text(position)?)
        })();
        self.check_source()?;
        result
    }

    /// The DOCX-style identifier of every style slot, by `istd` (record
    /// 0324).
    pub(crate) fn style_ids(&self) -> &[Box<str>] {
        self.0
            .style_ids
            .get_or_init(|| parts::style_ids(self.0.reader.stylesheet()))
    }

    /// The identifier of style `istd`, when the slot holds a style.
    pub(crate) fn style_id(&self, istd: u16) -> Option<&str> {
        self.style_ids()
            .get(usize::from(istd))
            .map(AsRef::as_ref)
            .filter(|id| !id.is_empty())
    }

    /// The body's children: its ordinary paragraphs and top-level tables.
    pub(crate) fn body_child_count(&self) -> Result<usize> {
        let tables = self
            .reader()?
            .body_blocks()?
            .count(litchi_doc::StoryBlockKind::Table);
        Ok(self.body_paragraph_count()?.saturating_add(tables))
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
    pub(crate) fn body_text_editor(&self) -> Result<Option<SourcePackage>> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Doc));
        }
        let DocSource::Retained(package) = &self.0.source else {
            return Ok(None);
        };
        if self.0.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                package.to_bytes()?,
                litchi_sign::cfb::Format::Doc,
            )?;
            return Ok(Some(SourcePackage::from_bytes(bytes)?));
        }
        Ok(Some(SourcePackage::from_shared(package.shared_source())?))
    }

    /// Return the retained native owner under the explicit signature disposition.
    pub(crate) fn editable_package(&self) -> Result<Option<litchi_doc::Package>> {
        self.check_source()?;
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Doc));
        }
        let DocSource::Retained(package) = &self.0.source else {
            return Ok(None);
        };
        if self.0.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                package.to_bytes()?,
                litchi_sign::cfb::Format::Doc,
            )?;
            return Ok(Some(litchi_doc::Package::from_bytes_with_options(
                &bytes,
                litchi_doc::OpenOptions::default()
                    .source_limits(package.source_limits())
                    .leniency(package.leniency()),
            )?));
        }
        Ok(Some(package.clone()))
    }

    /// Return the retained native owner for a setter that may prove an exact
    /// signed no-op. Removal-authorized handles still explicitly remove the
    /// signature before publishing; a signed handle keeps its original
    /// compound source until a changed result is refused.
    pub(crate) fn editable_package_for_write(&self) -> Result<Option<litchi_doc::Package>> {
        self.check_source_for_write()?;
        let DocSource::Retained(package) = &self.0.source else {
            return Ok(None);
        };
        if matches!(
            self.0.signature(),
            Signature::Signed | Signature::RemovalAuthorized
        ) {
            let bytes = signature::unsign_compound_file(
                package.to_bytes()?,
                litchi_sign::cfb::Format::Doc,
            )?;
            return Ok(Some(litchi_doc::Package::from_bytes_with_options(
                &bytes,
                litchi_doc::OpenOptions::default()
                    .source_limits(package.source_limits())
                    .leniency(package.leniency()),
            )?));
        }
        Ok(Some(package.clone()))
    }

    /// Return the retained native package for a canonical signed no-op patch.
    ///
    /// The write editor above deliberately removes signature storages so a
    /// changed publication can never carry them forward.  An exact logical
    /// cancellation needs an empty native patch whose source is the original
    /// signed compound artifact, however; keeping that source identity is
    /// what lets native replay and the unified history agree.
    pub(crate) fn original_package_for_noop(&self) -> Result<litchi_doc::Package> {
        self.check_source_for_write()?;
        match &self.0.source {
            DocSource::Retained(package) => Ok(package.clone()),
            DocSource::Released => Err(crate::Error::new(
                crate::Code::UnsupportedCapability,
                "the DOC source was released before signed no-op replay",
            )
            .with_format(FileFormat::Doc)
            .with_suggestion("reopen the document with source retention enabled")),
        }
    }

    /// Apply properties to one target or every match of a scoped query in
    /// one native body transaction (see [`edit`]).
    pub(crate) fn set<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<WriteCommit> {
        self.set_with_limits(target, props, &SemanticLimits::default())
    }

    /// Publish the exact compound source, preserving signatures until an
    /// explicitly authorized unsigned handle is written.
    pub(crate) fn write_to<W: Write>(&self, sink: &mut W) -> Result<u64> {
        let DocSource::Retained(package) = &self.0.source else {
            return Err(crate::Error::new(
                crate::Code::UnsupportedCapability,
                "the DOC source was released before publication",
            )
            .with_format(FileFormat::Doc));
        };
        if self.0.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                package.to_bytes()?,
                litchi_sign::cfb::Format::Doc,
            )?;
            sink.write_all(&bytes)?;
            return Ok(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
        }
        package.write_to(sink)?;
        Ok(package.shared_source().file_size())
    }

    /// [`Self::write_to`] into an owned buffer: the leaf's single sized
    /// copy of the retained compound file.
    pub(crate) fn to_bytes(&self) -> Result<Vec<u8>> {
        let DocSource::Retained(package) = &self.0.source else {
            return Err(crate::Error::new(
                crate::Code::UnsupportedCapability,
                "the DOC source was released before publication",
            )
            .with_format(FileFormat::Doc));
        };
        if self.0.signature() == Signature::RemovalAuthorized {
            return signature::unsign_compound_file(
                package.to_bytes()?,
                litchi_sign::cfb::Format::Doc,
            );
        }
        Ok(package.to_bytes()?)
    }

    fn with_package(&self, package: litchi_doc::Package) -> Self {
        Self::from_package(package, true).with_origin(self.origin())
    }

    /// The length of the retained input, or `None` when it was released.
    /// This is an O(1) lower-bound probe used before a source-backed native
    /// editor is allocated; the commit path still checks the exact retained
    /// patch after staging.
    pub(crate) fn retained_source_len(&self) -> Option<u64> {
        match &self.0.source {
            DocSource::Retained(package) => Some(package.shared_source().file_size()),
            DocSource::Released => None,
        }
    }

    /// Check the source, edit owner and caller-authored staging against the
    /// remaining snapshot budget before a DOC body editor can copy the CFB or
    /// build its immutable body snapshot.  The exact patch charge remains the
    /// publication check; this bound covers the live native edit handoff too.
    pub(crate) fn preflight_snapshot_bytes(&self, maximum: usize, additional: usize) -> Result<()> {
        let source = self
            .retained_source_len()
            .map_or(0, |bytes| usize::try_from(bytes).unwrap_or(usize::MAX));
        let observed = source
            .saturating_mul(2)
            .saturating_add(size_of::<litchi_doc::Patch>())
            .saturating_add(size_of::<litchi_doc::Edit<'static>>())
            .saturating_add(additional);
        if observed > maximum {
            return Err(crate::transaction::Error::Limit {
                kind: crate::transaction::LimitKind::SnapshotBytes,
                observed,
                maximum,
            }
            .into());
        }
        Ok(())
    }
}

/// The patch a DOC write step retains: a body edit or a metadata edit.
pub(crate) type StepPatch =
    crate::format::ole_metadata::StepPatch<litchi_doc::Patch, litchi_doc::package::metadata::Patch>;

/// The format-specific result retained by the generic transaction history.
pub(crate) struct WriteCommit {
    pub(crate) before: Doc,
    pub(crate) after: Doc,
    pub(crate) patch: StepPatch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

fn edit_options(transaction_limits: BodyTransactionLimits) -> litchi_doc::EditOptions {
    litchi_doc::EditOptions::default().transaction_limits(transaction_limits)
}

/// Keep the native DOC render handoff inside the transaction's remaining
/// snapshot budget. The body editor retains its source image while a rendered
/// candidate is live, so giving it the complete remaining ceiling would allow
/// the source and candidate to coexist above the facade limit before commit's
/// exact patch check runs.
pub(crate) fn body_transaction_limits_with_source(
    limits: SemanticLimits,
    source_bytes: usize,
    snapshot_budget: usize,
) -> BodyTransactionLimits {
    BodyTransactionLimits::new(limits.max_operations(), usize::MAX, usize::MAX)
        .with_max_retained_render_bytes(snapshot_budget.saturating_sub(source_bytes))
}

impl std::fmt::Debug for Doc {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Doc")
            .field(
                "source_retained",
                &matches!(self.0.source, DocSource::Retained(_)),
            )
            .finish_non_exhaustive()
    }
}

/// Charge both immutable endpoint snapshots and the semantic patch payload.
/// Text is retained in the patch in addition to the endpoint bytes; fixed
/// change descriptors account for formatting/resource operations.
pub(crate) fn patch_snapshot_bytes(patch: &litchi_doc::Patch) -> usize {
    let endpoints = patch
        .before_bytes()
        .len()
        .saturating_add(patch.after_bytes().len());
    let descriptors = size_of::<litchi_doc::body_text::ChangeRef<'static>>()
        .saturating_mul(patch.changes().len());
    let text = patch.changes().fold(0usize, |total, change| match change {
        litchi_doc::body_text::ChangeRef::Text { before, after, .. } => total
            .saturating_add(before.len())
            .saturating_add(after.len()),
        _ => total,
    });
    size_of::<litchi_doc::Patch>()
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
    use litchi_doc::{Leniency, raw::Reader};
    use litchi_ole_common::property_set::PropertySetReader;
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
        match detect_format_smart(bytes, &crate::detect::SourceLimits::default()) {
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
            let mut package = match Reader::from_reader(Cursor::new(fixture(&name))) {
                Ok(package) => package,
                Err(_) => continue,
            };
            let document = package.document().and_then(|document| {
                let metadata = package
                    .ole_file()
                    .get_metadata()
                    .map(Metadata::from)
                    .map_err(litchi_doc::Error::from)?;
                Ok((document, metadata))
            });
            for route in [
                open_path(&name, &OpenOptions::default()),
                open_bytes(fixture(&name), &OpenOptions::default()),
            ] {
                let (doc, document, metadata) = match (route, &document) {
                    (Ok(doc), Ok((document, metadata))) => (doc, document, metadata),
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
        let mut package = Reader::from_reader(Cursor::new(fixture(name))).unwrap();
        let document = package
            .document_with_options(
                litchi_doc::OpenOptions::default().leniency(Leniency::TolerateStylesheetDefects),
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
        let oracle = SourcePackage::from_bytes(bytes).unwrap();
        assert_eq!(editor.fingerprint(), oracle.fingerprint());
        let first = Position::new(0);
        assert_eq!(
            editor.paragraph(first).unwrap().text(),
            oracle.paragraph(first).unwrap().text()
        );
        // The editor is built over the retained source, not over a copy:
        // the sealed source's process-unique identity carries over.
        let DocSource::Retained(package) = &doc.0.source else {
            panic!("source not retained");
        };
        assert_eq!(
            editor.source_version(),
            package.shared_source().captured_source_version()
        );
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
