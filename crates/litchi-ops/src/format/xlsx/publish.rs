//! Publication through the existing XLSX leaf owners.

use super::{Xlsx, XlsxOwner};
#[cfg(feature = "encryption")]
use crate::Code;
use crate::format::signature::Signature;
use crate::save::{Overwrite, SaveOptions};
use crate::{Error, Result};
#[cfg(feature = "encryption")]
use litchi_core::Durability;
use litchi_core::ExecutionContext;
use std::io::{self, Write};
use std::path::Path;

impl Xlsx {
    /// Publish an XLSX snapshot through the shared atomic OPC writer.
    ///
    /// A source-backed owner streams its retained source artifact directly;
    /// it is never promoted to an eager workbook for an ordinary save. Eager
    /// owners clone their immutable workbook at the publication boundary.
    pub(crate) fn save_to(&self, path: &Path, options: &SaveOptions) -> Result<u64> {
        self.0.owner.ensure_current()?;
        let result = (|| {
            if self.signature() == Signature::RemovalAuthorized {
                let workbook = self.unsigned_workbook()?;
                let context = source_context(&self.0.owner);
                publish_owned(&workbook, path, options, context.as_ref(), || {
                    self.0.owner.ensure_current()
                })
            } else {
                match &self.0.owner {
                    XlsxOwner::Source(workbook) => publish_source(workbook, path, options),
                    XlsxOwner::Eager(workbook) => {
                        let workbook = workbook.clone();
                        publish_owned(&workbook, path, options, None, || Ok(()))
                    },
                }
            }
        })();
        self.0.owner.finish(|| result)
    }

    /// Stream a finalized XLSX snapshot to a sequential sink.
    pub(crate) fn write_to<W: Write>(&self, mut sink: W) -> Result<()> {
        self.0.owner.ensure_current()?;
        let result = (|| {
            if self.signature() == Signature::RemovalAuthorized {
                return self
                    .unsigned_workbook()?
                    .write_to(&mut sink)
                    .map_err(Error::from);
            }
            match &self.0.owner {
                XlsxOwner::Source(workbook) => workbook.write_to(&mut sink).map_err(Error::from),
                XlsxOwner::Eager(workbook) => {
                    workbook.clone().write_to(&mut sink).map_err(Error::from)
                },
            }
        })();
        self.0.owner.finish(|| result)
    }
}

/// Publish a retained source-backed package without materializing it.
///
/// The source artifact's own bounded writer performs source and execution
/// checks while staging. The owner fence is repeated immediately before the
/// atomic replacement and by the outer handle after publication.
fn publish_source(
    workbook: &litchi_xlsx::SourceBackedWorkbook,
    path: &Path,
    options: &SaveOptions,
) -> Result<u64> {
    #[cfg(feature = "encryption")]
    if options.reencrypt.is_some() {
        return Err(reencrypt_source_refused());
    }

    let package = workbook.source_backed_package()?;
    let context = package.execution_context();
    let mut bytes = 0_u64;
    let write = |sink: &mut litchi_opc::atomic::AtomicSink<'_>| {
        let mut counted = Counted::new(sink);
        package
            .source_artifact()
            .write_to(&mut counted)
            .map_err(Error::from)?;
        bytes = counted.written;
        Ok(())
    };
    let fence = || {
        workbook.source_version().map(drop).map_err(Error::from)?;
        package.check_execution().map_err(Error::from)
    };
    publish_atomic(path, options, context.as_ref(), write, fence)?;
    Ok(bytes)
}

/// Publish an owning workbook. Cloning occurs before entering the helper so
/// the immutable unified handle is unaffected by any leaf bookkeeping.
fn publish_owned(
    workbook: &litchi_xlsx::Package,
    path: &Path,
    options: &SaveOptions,
    context: Option<&ExecutionContext>,
    fence: impl FnOnce() -> Result<()>,
) -> Result<u64> {
    #[cfg(feature = "encryption")]
    {
        match (workbook.encryption(), options.reencrypt.as_ref()) {
            (Some(_), Some(password)) => {
                return publish_reencrypted(
                    workbook,
                    path,
                    options,
                    context,
                    fence,
                    password.as_str(),
                );
            },
            (Some(_), None) => return Err(sensitive_save_refused()),
            (None, Some(_)) => return Err(reencrypt_plain_refused()),
            (None, None) => {},
        }
    }

    let mut bytes = 0_u64;
    let write = |sink: &mut litchi_opc::atomic::AtomicSink<'_>| {
        let mut counted = Counted::new(sink);
        workbook.write_to(&mut counted).map_err(Error::from)?;
        bytes = counted.written;
        Ok(())
    };
    publish_atomic(path, options, context, write, fence)?;
    Ok(bytes)
}

#[cfg(feature = "encryption")]
fn publish_reencrypted(
    workbook: &litchi_xlsx::Package,
    path: &Path,
    options: &SaveOptions,
    context: Option<&ExecutionContext>,
    fence: impl FnOnce() -> Result<()>,
    password: &str,
) -> Result<u64> {
    // The leaf's re-encryption API currently promises full durability only.
    // Refuse weaker policies before generating output.
    if options.durability != Durability::Full {
        return Err(Error::new(
            Code::InvalidArgument,
            "re-encrypted XLSX publication requires full durability",
        )
        .with_format(litchi_core::FileFormat::Xlsx));
    }

    let output = workbook.to_reencrypted(password).map_err(Error::from)?;
    let bytes = output.len() as u64;
    let write = move |sink: &mut litchi_opc::atomic::AtomicSink<'_>| {
        sink.write_all(&output).map_err(Error::from)
    };
    publish_atomic(path, options, context, write, fence)?;
    Ok(bytes)
}

fn source_context(owner: &XlsxOwner) -> Option<ExecutionContext> {
    match owner {
        XlsxOwner::Source(workbook) => workbook
            .source_backed_package()
            .ok()
            .and_then(litchi_opc::SourceBackedPackage::execution_context),
        XlsxOwner::Eager(_) => None,
    }
}

fn publish_atomic<W, F>(
    path: &Path,
    options: &SaveOptions,
    context: Option<&ExecutionContext>,
    write: W,
    fence: F,
) -> Result<()>
where
    W: FnOnce(&mut litchi_opc::atomic::AtomicSink<'_>) -> std::result::Result<(), Error>,
    F: FnOnce() -> Result<()>,
{
    match options.overwrite {
        Overwrite::Replace => litchi_opc::atomic::replace_guarded_with_durability(
            path,
            options.durability,
            context,
            write,
            fence,
        ),
        Overwrite::Refuse => litchi_opc::atomic::replace_guarded_with_durability_noclobber(
            path,
            options.durability,
            context,
            write,
            fence,
        ),
    }
}

#[cfg(feature = "encryption")]
fn sensitive_save_refused() -> Error {
    Error::new(
        Code::DocumentProtected,
        "the document is encrypted; pass a password to re-encrypt it or consume it as plaintext",
    )
    .with_format(litchi_core::FileFormat::Xlsx)
    .with_suggestion("use `SaveOptions::reencrypt` for encrypted publication")
}

#[cfg(feature = "encryption")]
fn reencrypt_plain_refused() -> Error {
    Error::new(
        Code::InvalidArgument,
        "re-encryption requires an encrypted document",
    )
    .with_format(litchi_core::FileFormat::Xlsx)
}

#[cfg(feature = "encryption")]
fn reencrypt_source_refused() -> Error {
    Error::new(
        Code::InvalidArgument,
        "re-encryption is unavailable for a source-backed XLSX publication",
    )
    .with_format(litchi_core::FileFormat::Xlsx)
}

struct Counted<W> {
    inner: W,
    written: u64,
}

impl<W> Counted<W> {
    const fn new(inner: W) -> Self {
        Self { inner, written: 0 }
    }
}

impl<W: Write> Write for Counted<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let accepted = self.inner.write(bytes)?;
        self.written = self
            .written
            .checked_add(
                u64::try_from(accepted)
                    .map_err(|_error| io::Error::other("output byte count exceeds u64"))?,
            )
            .ok_or_else(|| io::Error::other("output byte count exceeds u64"))?;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::options::OpenOptions;

    #[test]
    fn write_to_uses_the_owned_leaf_writer() {
        let workbook = litchi_xlsx::Package::new().expect("create workbook");
        let handle = Xlsx::from_workbook(workbook, &OpenOptions::new()).expect("adopt workbook");
        let mut output = Vec::new();
        handle.write_to(&mut output).expect("stream workbook");
        let reopened = litchi_xlsx::Package::from_bytes(&output).expect("parse output");
        assert_eq!(reopened.workbook().len(), 1);
    }

    #[test]
    fn source_write_to_copies_the_exact_retained_artifact() {
        let source = litchi_xlsx::Package::new()
            .expect("create workbook")
            .to_bytes()
            .expect("serialize source");
        let package = litchi_opc::SourceBackedPackage::from_vec(
            source.clone(),
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            None,
        )
        .expect("source package");
        let handle =
            Xlsx::from_source_package(package, &OpenOptions::new()).expect("adopt source package");
        let mut output = Vec::new();
        handle.write_to(&mut output).expect("stream source");
        assert_eq!(output, source);
    }
}
