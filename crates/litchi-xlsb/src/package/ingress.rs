//! Classify unsupported compound-file containers after normal OPC admission.

use super::{PackageError, PackageResult};
use litchi_core::{ExecutionContext, ReadAt};
use litchi_opc::{OpcError, OpcPackage, SourceBackedPackage, SourceCacheLimits, SourceLimits};
use std::io::{Read, Result as IoResult};
use std::sync::Arc;

const CFB: &[u8] = litchi_core::detection::utils::OLE2_SIGNATURE;
const UNSUPPORTED: &str = "CFB containers at the XLSB boundary (including encrypted packages)";

struct PrefixReader<R> {
    inner: R,
    prefix: [u8; 8],
    used: usize,
}
impl<R> PrefixReader<R> {
    fn new(inner: R) -> Self {
        Self {
            inner,
            prefix: [0; 8],
            used: 0,
        }
    }
}
impl<R: Read> Read for PrefixReader<R> {
    fn read(&mut self, output: &mut [u8]) -> IoResult<usize> {
        let read = self.inner.read(output)?;
        let copied = read.min(self.prefix.len() - self.used);
        self.prefix[self.used..self.used + copied].copy_from_slice(&output[..copied]);
        self.used += copied;
        Ok(read)
    }
}

fn classify(error: OpcError, prefix: &[u8]) -> PackageError {
    if matches!(error, OpcError::ZipError(_)) && prefix.starts_with(CFB) {
        PackageError::UnsupportedFeature(UNSUPPORTED)
    } else {
        PackageError::Opc(error)
    }
}

pub(crate) fn from_reader(reader: impl Read, limits: SourceLimits) -> PackageResult<OpcPackage> {
    let mut reader = PrefixReader::new(reader);
    OpcPackage::from_reader(&mut reader, limits)
        .map_err(|error| classify(error, &reader.prefix[..reader.used]))
}

pub(crate) fn from_vec(bytes: Vec<u8>, limits: SourceLimits) -> PackageResult<OpcPackage> {
    let compound = bytes.starts_with(CFB);
    OpcPackage::from_vec(bytes, limits)
        .map_err(|error| classify(error, if compound { CFB } else { &[] }))
}

pub(crate) fn source_from_reader(
    reader: impl Read,
    limits: SourceLimits,
) -> PackageResult<SourceBackedPackage> {
    let mut reader = PrefixReader::new(reader);
    SourceBackedPackage::from_reader(&mut reader, limits, SourceCacheLimits::default(), None)
        .map_err(|error| classify(error, &reader.prefix[..reader.used]))
}

pub(crate) fn source_from_read_at(
    source: Arc<dyn ReadAt>,
    limits: SourceLimits,
    cache: SourceCacheLimits,
    context: Option<ExecutionContext>,
) -> PackageResult<SourceBackedPackage> {
    check_context(context.as_ref())?;
    let expected = source.version().map_err(OpcError::from)?;
    let result = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        limits,
        cache,
        litchi_opc::SourceReadPolicy::exact(),
        context.clone(),
    );
    if let Err(error @ OpcError::ZipError(_)) = result {
        // Limit, cancellation and source errors from OPC keep precedence.
        // Sniff only a failed archive, then fence the diagnostic read too.
        check_context(context.as_ref())?;
        let mut prefix = [0; 8];
        let count = usize::try_from(source.len().map_err(OpcError::from)?.min(8)).map_err(
            |_source_error| PackageError::InvalidFormat("container prefix length exceeds usize"),
        )?;
        let read = source.read_exact_at(0, &mut prefix[..count]);
        let observed = source.version().map_err(OpcError::from)?;
        if observed != expected {
            return Err(PackageError::Opc(OpcError::SourceChanged {
                expected,
                actual: observed,
            }));
        }
        check_context(context.as_ref())?;
        read.map_err(OpcError::from)?;
        Err(classify(error, &prefix[..count]))
    } else {
        result.map_err(Into::into)
    }
}

fn check_context(context: Option<&ExecutionContext>) -> PackageResult<()> {
    if let Some(context) = context {
        context.check().map_err(|error| {
            if matches!(error, litchi_core::ExecutionError::Cancelled) {
                OpcError::Cancelled
            } else {
                OpcError::Execution(error)
            }
        })?;
    }
    Ok(())
}
