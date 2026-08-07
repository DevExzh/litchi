use std::fs::File;
use std::io::Read;
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::{Error, LimitKind, Limits, Result};

pub(crate) fn read_package(path: &Path, limits: Limits) -> Result<Arc<[u8]>> {
    let file = File::open(path)?;
    let metadata = file.metadata()?;
    let maximum = limits.max_input_bytes();
    if metadata.len() > maximum {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: metadata.len(),
            maximum,
        });
    }

    let expected = usize::try_from(metadata.len()).map_err(|_error| {
        Error::InvalidBundle("iWork package length does not fit usize".to_owned())
    })?;
    let mut source = Vec::new();
    source
        .try_reserve_exact(expected)
        .map_err(|_error| Error::Allocation {
            resource: "example package source",
            amount: expected,
        })?;
    file.take(maximum.saturating_add(1))
        .read_to_end(&mut source)?;

    let observed = u64::try_from(source.len()).map_err(|_error| {
        Error::InvalidBundle("iWork package length does not fit u64".to_owned())
    })?;
    if observed > maximum {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
        });
    }
    Ok(source.into())
}
