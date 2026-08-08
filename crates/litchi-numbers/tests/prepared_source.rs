#![cfg(feature = "internal-iwork-source")]

use std::path::PathBuf;
use std::sync::Arc;

use litchi_iwa_detect::{Format, PreparedSource};
use litchi_numbers::{
    __compatibility_tables_from_prepared_source, PackageSemanticLimits,
    compatibility_tables_from_bytes,
};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

#[test]
fn prepared_source_handoff_matches_global_compatibility_projection()
-> Result<(), Box<dyn std::error::Error>> {
    let source: Arc<[u8]> = std::fs::read(fixture_path())?.into();
    let expected = compatibility_tables_from_bytes(&source)?;
    let prepared = PreparedSource::from_shared_bytes(Arc::clone(&source))?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture was not detected"))?;
    assert_eq!(prepared.format(), Format::Numbers);

    let actual =
        __compatibility_tables_from_prepared_source(prepared, PackageSemanticLimits::default())?;

    assert_eq!(actual, expected);
    assert_eq!(actual.len(), 1);
    assert_eq!(actual[0].name(), "Table 1");
    Ok(())
}
