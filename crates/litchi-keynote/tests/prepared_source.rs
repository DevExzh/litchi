#![cfg(feature = "internal-iwork-source")]

use std::path::PathBuf;
use std::sync::Arc;

use litchi_iwa_detect::{Format, PreparedSource};
use litchi_keynote::{Package, SemanticLimits};

fn exact_bytes(package: &Package) -> Result<Vec<u8>, litchi_keynote::WriteError> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
}

#[test]
fn prepared_source_handoff_preserves_bytes_and_semantics() -> Result<(), Box<dyn std::error::Error>>
{
    let source: Arc<[u8]> = std::fs::read(fixture_path())?.into();
    let prepared = PreparedSource::from_shared_bytes(Arc::clone(&source))?
        .ok_or_else(|| std::io::Error::other("native Keynote fixture was not detected"))?;
    assert_eq!(prepared.format(), Format::Keynote);

    let package = Package::__from_prepared_source(prepared, SemanticLimits::default())?;
    let direct = Package::from_bytes(&source)?;

    assert_eq!(exact_bytes(&package)?, source.as_ref());
    assert_eq!(package.show()?, direct.show()?);
    assert_eq!(
        package.semantic_snapshot()?.slides().as_ptr(),
        package.slides()?.as_ptr()
    );
    Ok(())
}
