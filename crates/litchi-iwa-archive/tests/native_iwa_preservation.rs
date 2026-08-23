use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::package::{Catalog, EntryEdit, ReassemblyPatchError};
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
    assert!(
        catalog.source_is_exact(),
        "native flat package unexpectedly lost exact-source authority: {name}"
    );
    assert_eq!(
        catalog.to_bytes()?,
        source,
        "catalog no-op changed package bytes in {name}"
    );
    let mut streamed = Vec::new();
    catalog.write_to(&mut streamed)?;
    assert_eq!(
        streamed, source,
        "streaming catalog no-op changed package bytes in {name}"
    );
    assert_eq!(
        catalog.reassemble_to_bytes(&[], Limits::default())?,
        source,
        "empty reassembly changed package bytes in {name}"
    );
    let mut summary = CorpusSummary {
        components: 0,
        objects: 0,
        messages: 0,
        decompressed_bytes: 0,
    };

    for entry in catalog.iter().filter(|entry| {
        #[allow(
            clippy::case_sensitive_file_extension_comparisons,
            reason = "IWA member names are case-sensitive protocol names."
        )]
        {
            entry.name().ends_with(".iwa")
        }
    }) {
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

fn verify_native_reassembly_patch(name: &str) -> Result<(), Box<dyn Error>> {
    let source = fs::read(fixture_path(name))?;
    let catalog = Catalog::from_bytes(&source)?;
    let selected = catalog
        .iter()
        .find(|entry| {
            #[allow(
                clippy::case_sensitive_file_extension_comparisons,
                reason = "IWA member names are case-sensitive protocol names."
            )]
            {
                !entry.is_opaque() && !entry.name().ends_with(".iwa")
            }
        })
        .ok_or("native fixture has no decoded non-IWA member")?;
    let selected_name = selected.name().to_owned();
    let patch = catalog.reassemble_to_patch(
        &[EntryEdit::new(
            &selected_name,
            b"patched non-IWA physical member",
        )],
        Limits::default(),
    )?;
    let target = patch.apply(&source)?;
    let target_catalog = Catalog::from_bytes(&target)?;
    assert_eq!(target_catalog.len(), catalog.len());
    let source_archive = soapberry_zip::ZipArchive::from_slice(&source)?;
    let target_archive = soapberry_zip::ZipArchive::from_slice(&target)?;
    assert_eq!(target_archive.entries_hint(), source_archive.entries_hint());

    let target_selected = target_catalog
        .iter()
        .find(|entry| entry.name() == selected_name)
        .ok_or("reassembled fixture lost the selected member")?;
    assert_eq!(target_selected.data(), b"patched non-IWA physical member");

    for target_entry in target_catalog
        .iter()
        .filter(|entry| entry.name() != selected_name)
    {
        let source_entry = catalog
            .iter()
            .find(|entry| entry.name() == target_entry.name())
            .ok_or("reassembled fixture lost an untouched member")?;
        assert_eq!(target_entry.data(), source_entry.data());
        assert_eq!(target_entry.raw_name(), source_entry.raw_name());
        assert_eq!(
            target_entry.raw_record().compressed_data(),
            source_entry.raw_record().compressed_data()
        );
        let source_central = source_entry.raw_record().central_directory_record();
        let target_central = target_entry.raw_record().central_directory_record();
        assert_eq!(&target_central[..42], &source_central[..42]);
        assert_eq!(&target_central[46..], &source_central[46..]);
        assert_eq!(
            target_entry.metadata().local(),
            source_entry.metadata().local()
        );
        assert_eq!(
            target_entry.metadata().central().name(),
            source_entry.metadata().central().name()
        );
        assert_eq!(
            target_entry.metadata().central().extra(),
            source_entry.metadata().central().extra()
        );
        assert_eq!(
            target_entry.metadata().central().comment(),
            source_entry.metadata().central().comment()
        );
    }

    assert_eq!(patch.inverse().apply(&target)?, source);
    let mut foreign = source.clone();
    foreign.push(0);
    assert!(matches!(
        patch.apply(&foreign),
        Err(ReassemblyPatchError::Conflict)
    ));
    Ok(())
}

#[test]
fn native_reassembly_patch_round_trips_all_iwork_families() -> Result<(), Box<dyn Error>> {
    for fixture in [
        "numbers/basic.numbers",
        "pages/basic.pages",
        "keynote/basic.key",
    ] {
        verify_native_reassembly_patch(fixture)?;
    }
    Ok(())
}
