//! Publication through the existing leaf owners.

use super::{Docx, DocxOwner};
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

impl Docx {
    /// Publish a DOCX snapshot through the leaf's atomic writer.
    ///
    /// The unified layer supplies only destination and overwrite policy.  The
    /// package owner remains responsible for source preservation, encryption
    /// policy and its final source fence; no filesystem rename or
    /// synchronization is implemented here.
    pub(crate) fn save_to(&self, path: &Path, options: &SaveOptions) -> Result<u64> {
        self.owner().check_source()?;
        let result = (|| {
            if self.signature() == Signature::RemovalAuthorized {
                let mut package = self.unsigned_package()?;
                let context = source_context(self.owner());
                publish_owned(&mut package, path, options, context.as_ref(), || {
                    self.owner().check_source()
                })
            } else {
                match self.owner() {
                    DocxOwner::Source(package) => publish_source(package, path, options),
                    DocxOwner::Eager(package) => {
                        let mut package = package.clone();
                        publish_owned(&mut package, path, options, None, || Ok(()))
                    },
                }
            }
        })();
        self.owner().finish(result)
    }

    pub(crate) fn write_to<W: Write>(&self, sink: W) -> Result<()> {
        self.owner().check_source()?;
        let result = (|| {
            if self.signature() == Signature::RemovalAuthorized {
                self.unsigned_package()?.write_to(sink).map_err(Into::into)
            } else {
                match self.owner() {
                    DocxOwner::Source(package) => package.write_to(sink).map_err(Into::into),
                    DocxOwner::Eager(package) => package.clone().write_to(sink).map_err(Into::into),
                }
            }
        })();
        self.owner().finish(result)
    }
}

/// Publish a retained source-backed package.  The source fence is run after
/// staging and immediately before the leaf atomic operation, including for
/// `Overwrite::Refuse`.
fn publish_source(
    package: &litchi_docx::source_backed::Package,
    path: &Path,
    options: &SaveOptions,
) -> Result<u64> {
    #[cfg(feature = "encryption")]
    if options.reencrypt.is_some() {
        return Err(reencrypt_source_refused());
    }

    let mut bytes = 0_u64;
    let context = package.execution_context();
    let write = |sink: &mut litchi_opc::atomic::AtomicSink<'_>| {
        let mut counted = Counted::new(sink);
        package.write_to(&mut counted).map_err(Error::from)?;
        bytes = counted.written;
        Ok(())
    };
    let fence = || package.check_source().map_err(Error::from);
    publish_atomic(path, options, context.as_ref(), write, fence)?;
    Ok(bytes)
}

/// Publish an owning package.  Cloning the package at the publication
/// boundary keeps the immutable unified handle unchanged even when the leaf
/// writer updates internal publication bookkeeping.
fn publish_owned(
    package: &mut litchi_docx::Package,
    path: &Path,
    options: &SaveOptions,
    context: Option<&ExecutionContext>,
    fence: impl FnOnce() -> Result<()>,
) -> Result<u64> {
    #[cfg(feature = "encryption")]
    {
        match (package.encryption(), options.reencrypt.as_ref()) {
            (Some(_), Some(password)) => {
                return publish_reencrypted(
                    package,
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
        package.write_to(&mut counted).map_err(Error::from)?;
        bytes = counted.written;
        Ok(())
    };
    publish_atomic(path, options, context, write, fence)?;
    Ok(bytes)
}

#[cfg(feature = "encryption")]
fn publish_reencrypted(
    package: &mut litchi_docx::Package,
    path: &Path,
    options: &SaveOptions,
    context: Option<&ExecutionContext>,
    fence: impl FnOnce() -> Result<()>,
    password: &str,
) -> Result<u64> {
    // The leaf's re-encryption API currently promises full durability only.
    // Refuse weaker policies before generating any output so the destination
    // cannot be changed by an unsupported combination.
    if options.durability != Durability::Full {
        return Err(Error::new(
            Code::InvalidArgument,
            "re-encrypted DOCX publication requires full durability",
        )
        .with_format(litchi_core::FileFormat::Docx));
    }

    // Generate ciphertext in memory through the leaf.  Only ciphertext is
    // handed to the atomic staging sink; no plaintext temporary is created.
    let output = package.to_reencrypted(password).map_err(Error::from)?;
    let bytes = output.len() as u64;
    let write = move |sink: &mut litchi_opc::atomic::AtomicSink<'_>| {
        sink.write_all(&output).map_err(Error::from)
    };
    publish_atomic(path, options, context, write, fence)?;
    Ok(bytes)
}

fn source_context(owner: &DocxOwner) -> Option<ExecutionContext> {
    match owner {
        DocxOwner::Source(package) => package.execution_context(),
        DocxOwner::Eager(_) => None,
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
    .with_format(litchi_core::FileFormat::Docx)
    .with_suggestion("use `SaveOptions::reencrypt` for encrypted publication")
}

#[cfg(feature = "encryption")]
fn reencrypt_plain_refused() -> Error {
    Error::new(
        Code::InvalidArgument,
        "re-encryption requires an encrypted document",
    )
    .with_format(litchi_core::FileFormat::Docx)
}

#[cfg(feature = "encryption")]
fn reencrypt_source_refused() -> Error {
    Error::new(
        Code::InvalidArgument,
        "re-encryption is unavailable for a source-backed DOCX publication",
    )
    .with_format(litchi_core::FileFormat::Docx)
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
