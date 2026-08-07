use std::error::Error as StdError;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_archive::{
    ComponentCatalog, Error, LimitKind, Limits, Result, inspect_detection_root,
};
use soapberry_zip::ZipArchive as RawZipArchive;
use soapberry_zip::office::StreamingArchiveWriter;

fn nested_index_with_compressible_payload() -> Result<Vec<u8>> {
    let payload = vec![0u8; 64 * 1024];
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("Index/Document.iwa", &payload)?;
    Ok(writer.finish_to_bytes()?)
}

fn corrupt_only_compressed_member(bytes: &mut [u8]) -> Result<()> {
    let (raw_start, raw_end) = {
        let archive = RawZipArchive::from_slice(&*bytes)?;
        let header = archive.entries().next().ok_or_else(|| {
            Error::InvalidBundle("test ZIP unexpectedly contains no member".to_owned())
        })??;
        archive
            .get_entry(header.wayfinder())?
            .compressed_data_range()
    };
    let start = usize::try_from(raw_start).map_err(|_error| {
        Error::InvalidBundle("test compressed offset does not fit usize".to_owned())
    })?;
    let end = usize::try_from(raw_end).map_err(|_error| {
        Error::InvalidBundle("test compressed end does not fit usize".to_owned())
    })?;
    if start >= end {
        return Err(Error::InvalidBundle(
            "test compressed member has an invalid range".to_owned(),
        ));
    }
    let first = bytes.get_mut(start).ok_or_else(|| {
        Error::InvalidBundle("test compressed member is unexpectedly empty".to_owned())
    })?;
    *first ^= 0xff;
    Ok(())
}

fn assert_nested_input_limit<T>(label: &str, result: Result<T>, observed: u64, maximum: u64) {
    match result {
        Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: actual,
            maximum: limit,
        }) if actual == observed && limit == maximum => {},
        Err(error) => panic!("{label} returned the wrong error: {error:?}"),
        Ok(_) => panic!("{label} accepted an oversized nested index"),
    }
}

#[test]
fn declared_nested_index_size_wins_over_corrupt_deflate_for_every_ingress_api()
-> std::result::Result<(), Box<dyn StdError>> {
    let index = nested_index_with_compressible_payload()?;
    let mut writer = StreamingArchiveWriter::new();
    writer.write_deflated("legacy.pages/Index.zip", &index)?;
    let mut bytes = writer.finish_to_bytes()?;
    assert!(bytes.len() < index.len());
    corrupt_only_compressed_member(&mut bytes)?;

    let input = u64::try_from(bytes.len())?;
    let nested = u64::try_from(index.len())?;
    let limits = Limits::new(input, 10, nested, nested, 1024)?;

    assert_nested_input_limit(
        "physical catalog",
        Catalog::from_bytes_with_limits(&bytes, limits),
        nested,
        input,
    );
    assert_nested_input_limit(
        "component catalog",
        ComponentCatalog::from_bytes_with_limits(&bytes, limits),
        nested,
        input,
    );
    assert_nested_input_limit(
        "detection root",
        inspect_detection_root(&bytes, limits),
        nested,
        input,
    );
    Ok(())
}
