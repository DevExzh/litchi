//! `Package::from_prepared_source` adopts the container detection parsed.

use std::path::PathBuf;
use std::sync::Arc;

use litchi_iwa_detect::{Format, PreparedSource};
use litchi_numbers::Package;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

#[test]
fn prepared_source_handoff_preserves_bytes_and_semantics() -> Result<(), Box<dyn std::error::Error>>
{
    let source: Arc<[u8]> = std::fs::read(fixture_path())?.into();
    let direct = Package::from_bytes(&source)?;
    let prepared = PreparedSource::from_shared_bytes(
        Arc::clone(&source),
        litchi_iwa_detect::Limits::default(),
    )?
    .ok_or_else(|| std::io::Error::other("native Numbers fixture was not detected"))?;
    assert_eq!(prepared.format(), Format::Numbers);

    let package = Package::from_prepared_source(prepared, litchi_numbers::SemanticLimits::DEFAULT)?;

    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    assert_eq!(bytes, source.as_ref());
    assert_eq!(package.sheets(), direct.sheets());
    Ok(())
}
