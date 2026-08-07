use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_core::{Archive, SnappyStream};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CorpusSummary {
    components: usize,
    objects: usize,
    messages: usize,
    decompressed_bytes: usize,
}

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(name)
}

fn verify_fixture(name: &str) -> Result<CorpusSummary, Box<dyn Error>> {
    let source = fs::read(fixture_path(name))?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut summary = CorpusSummary {
        components: 0,
        objects: 0,
        messages: 0,
        decompressed_bytes: 0,
    };

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        assert!(
            !entry.is_opaque(),
            "native IWA component unexpectedly remained opaque: {}",
            entry.name()
        );
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        let encoded = archive.to_bytes()?;
        assert_eq!(
            encoded,
            stream.as_bytes(),
            "decompressed IWA no-op changed component {} in {name}",
            entry.name()
        );

        summary.components += 1;
        summary.objects += archive.objects.len();
        summary.messages += archive
            .objects
            .iter()
            .map(|object| object.messages.len())
            .sum::<usize>();
        summary.decompressed_bytes += stream.as_bytes().len();
    }

    assert!(summary.components > 0, "{name} contained no IWA components");
    Ok(summary)
}

#[test]
fn native_iwa_headers_and_payloads_are_exactly_preserved() -> Result<(), Box<dyn Error>> {
    let corpus = [
        (
            "numbers/basic.numbers",
            CorpusSummary {
                components: 37,
                objects: 622,
                messages: 631,
                decompressed_bytes: 373_043,
            },
        ),
        (
            "pages/basic.pages",
            CorpusSummary {
                components: 7,
                objects: 570,
                messages: 576,
                decompressed_bytes: 360_855,
            },
        ),
        (
            "keynote/basic.key",
            CorpusSummary {
                components: 25,
                objects: 959,
                messages: 965,
                decompressed_bytes: 443_469,
            },
        ),
    ];
    for (fixture, expected) in corpus {
        assert_eq!(
            verify_fixture(fixture)?,
            expected,
            "corpus drift in {fixture}"
        );
    }
    Ok(())
}
