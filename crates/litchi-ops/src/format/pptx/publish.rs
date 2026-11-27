//! Guarded atomic publication through native PPTX owners.
use super::{Pptx, PptxOwner};
#[cfg(feature = "encryption")]
use crate::Error;
use crate::Result;
use crate::save::{Overwrite, SaveOptions};
use std::io::{self, Write};
use std::path::Path;
impl Pptx {
    pub(crate) fn save_to(&self, path: &Path, options: &SaveOptions) -> Result<u64> {
        self.check_owner()?;
        #[cfg(feature = "encryption")]
        if options.reencrypt.is_some() && !self.requires_reencryption() {
            return Err(Error::new(
                crate::Code::InvalidArgument,
                "re-encryption requires an encrypted presentation",
            )
            .with_format(litchi_core::FileFormat::Pptx));
        }
        #[cfg(feature = "encryption")]
        if options.reencrypt.is_some() && self.signature() == super::Signature::RemovalAuthorized {
            return Err(Error::new(
                crate::Code::DocumentProtected,
                "encrypted PPTX signature removal requires an unsupported native edit",
            )
            .with_format(litchi_core::FileFormat::Pptx));
        }
        let context = match &self.0.owner {
            PptxOwner::Source(source) => source.source_backed_package().execution_context(),
            PptxOwner::Eager(_) => self
                .0
                .source_guard
                .as_ref()
                .and_then(|source| source.source_backed_package().execution_context()),
        };
        self.finish(|| {
            let mut bytes = 0;
            let write = |sink: &mut litchi_opc::atomic::AtomicSink<'_>| -> Result<()> {
                let mut counted = Counted { sink, bytes: 0 };
                #[cfg(feature = "encryption")]
                if let Some(password) = &options.reencrypt {
                    let PptxOwner::Eager(package) = &self.0.owner else {
                        unreachable!("encrypted PPTX owners are eager");
                    };
                    package.write_reencrypted(&mut counted, password.as_str())?;
                } else {
                    self.write_to(&mut counted)?;
                }
                #[cfg(not(feature = "encryption"))]
                self.write_to(&mut counted)?;
                bytes = counted.bytes;
                Ok(())
            };
            let fence = || self.check_owner();
            match options.overwrite {
                Overwrite::Replace => litchi_opc::atomic::replace_guarded_with_durability(
                    path,
                    options.durability,
                    context.as_ref(),
                    write,
                    fence,
                )?,
                Overwrite::Refuse => litchi_opc::atomic::replace_guarded_with_durability_noclobber(
                    path,
                    options.durability,
                    context.as_ref(),
                    write,
                    fence,
                )?,
            }
            Ok(bytes)
        })
    }
}
struct Counted<W> {
    sink: W,
    bytes: u64,
}
impl<W: Write> Write for Counted<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = self.sink.write(bytes)?;
        self.bytes = self
            .bytes
            .checked_add(count as u64)
            .ok_or_else(|| io::Error::other("output byte count exceeds u64"))?;
        Ok(count)
    }
    fn flush(&mut self) -> io::Result<()> {
        self.sink.flush()
    }
}
