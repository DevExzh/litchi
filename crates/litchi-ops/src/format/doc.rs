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

use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::{OleFile, SharedOleFile, SharedOleFileLimits};
use litchi_core::{Metadata, OwnedSource, ReadAt, SourceVersion};
use litchi_doc::body_text::source::{self as body_source, SourceSnapshot};
use litchi_ole_common::property_set::PropertySetReader;

use super::LegacyOpen;
use crate::{OpenOptions, Result};

/// An opened legacy Word (`.doc`) document.
///
/// One `Arc` wide: cloning is one atomic increment, and the handle is
/// `Send + Sync`. Reads go to the reader the open built. The handle also
/// retains the input (unless [`OpenOptions::retain_source`] was `false`), so
/// a later edit needs no second open.
#[derive(Clone)]
pub struct Doc(Arc<DocState>);

struct DocState {
    /// The reader every M0a read goes through, built at open exactly as the
    /// facade builds it.
    reader: litchi_doc::Document,
    /// The OLE property-set projection, read at open as the facade reads it
    /// (record 0003 §1.6).
    metadata: Metadata,
    /// The source an editor is built from.
    source: DocSource,
}

/// The source state of a [`Doc`].
///
/// Only the in-memory state exists today, because no route opens a DOC over
/// a positional source; a positional route would add a variant here.
enum DocSource {
    /// The input bytes and the CFB index detection built, sealed together.
    Retained(Arc<SharedOleFile>),
    /// The input was released after open (`retain_source(false)`).
    Released,
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Doc>();
    assert!(size_of::<Doc>() == size_of::<usize>());
};

impl Doc {
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
        let mut package = litchi_doc::Package::from_ole_file(ole)?;
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
        let metadata = package
            .ole_file()
            .get_metadata()
            .map(Metadata::from)
            .unwrap_or_default();
        let source = if options.retain_source {
            retain(package.into_ole_file())?
        } else {
            DocSource::Released
        };
        Ok(LegacyOpen::Opened(Self(Arc::new(DocState {
            reader,
            metadata,
            source,
        }))))
    }

    /// All text of the document, as the facade's `Document::text` returns it.
    ///
    /// # Errors
    ///
    /// Returns a typed DOC error when a text story cannot be decoded.
    pub fn text(&self) -> Result<String> {
        Ok(self.0.reader.text()?)
    }

    /// The document's OLE property-set metadata, read at open. Always
    /// `Some`; absent or unreadable property streams read as empty metadata,
    /// as the facade reports them.
    ///
    /// # Errors
    ///
    /// Never fails today; the `Result` matches the unified read surface.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(Some(self.0.metadata.clone()))
    }

    /// Everything the reader repaired while the document was opened.
    ///
    /// # Errors
    ///
    /// Never fails today; the `Result` matches the unified read surface.
    pub fn tolerance_report(&self) -> Result<litchi_doc::ToleranceReport> {
        Ok(self.0.reader.tolerance_report().clone())
    }

    /// The reader the family views read paragraphs, tables and elements
    /// from.
    pub(crate) fn reader(&self) -> &litchi_doc::Document {
        &self.0.reader
    }

    /// Builds the source-backed main-story editor over the retained source
    /// (L-DOC-1), without indexing the container again.
    ///
    /// Returns `Ok(None)` when the source was released at open. The editor's
    /// open runs its FIB, CLX and fingerprint checks here, at the first
    /// edit. M2 maps the leaf error into the unified one.
    #[cfg_attr(not(test), expect(dead_code, reason = "edit verbs arrive in M2"))]
    pub(crate) fn body_text_editor(
        &self,
    ) -> std::result::Result<Option<SourceSnapshot>, body_source::Error> {
        match &self.0.source {
            DocSource::Retained(shared) => {
                SourceSnapshot::from_shared(Arc::clone(shared)).map(Some)
            },
            DocSource::Released => Ok(None),
        }
    }

    /// The length of the retained input, or `None` when it was released.
    #[cfg(test)]
    fn retained_source_len(&self) -> Option<u64> {
        match &self.0.source {
            DocSource::Retained(shared) => Some(shared.file_size()),
            DocSource::Released => None,
        }
    }
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

/// Seals a parsed in-memory OLE file as the retained source.
fn retain(ole: OleFile<Cursor<Vec<u8>>>) -> Result<DocSource> {
    let shared =
        SharedOleFile::from_owned_ole_file(ole, sealed_version()?, SharedOleFileLimits::default())?;
    Ok(DocSource::Retained(Arc::new(shared)))
}

/// A process-unique identity for a sealed in-memory source, minted by the
/// same counter as every other [`OwnedSource`].
fn sealed_version() -> Result<SourceVersion> {
    Ok(OwnedSource::new(Vec::new()).version()?)
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
                    doc.reader().paragraph_count().ok(),
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
        let first = litchi_core::Position::new(0);
        assert_eq!(
            editor.paragraph(first).unwrap().text(),
            oracle.paragraph(first).unwrap().text()
        );
        // The editor is built over the retained source, not over a copy:
        // the sealed source's process-unique identity carries over.
        let DocSource::Retained(shared) = &doc.0.source else {
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
