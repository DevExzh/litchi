//! Host owner I/O against the leaf owners (wave W4-R, record 0417).
//!
//! The host's mutable `Draft` owners (Pages, Numbers, the format-neutral
//! text and comments editors), its raw `Package` and its format-neutral
//! `Document` reader each open, validate, write and save a package. The
//! leaf owners are immutable `Package`s (`litchi_{pages,keynote,numbers}`)
//! and, for physical access, `litchi_iwa_archive::SourceCatalog`. These
//! read-only cases prove, on native-saved fixtures, that every host owner's
//! unedited output equals the leaf owner's, through every host entry point,
//! so the host rows retire with these tests as their pointers.

use std::sync::Arc;

use litchi_core::SliceSource;
use litchi_iwa_archive::{Limits, SourceCatalog};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const PAGES: &str = "iwork/pages/body-table-merges-native.pages";
const NUMBERS: &str = "iwork/numbers/table-merges-native.numbers";
const KEYNOTE: &str = "iwork/keynote/slide-table-merges-native.key";

/// A scratch directory unique to one test.
fn scratch(name: &str) -> TestResult<std::path::PathBuf> {
    let directory = std::env::temp_dir().join(format!("w4r-{name}-{}", std::process::id()));
    std::fs::create_dir_all(&directory)?;
    Ok(directory)
}

/// The leaf owner's unedited bytes.
fn leaf_output(format: Format, source: &[u8]) -> TestResult<Vec<u8>> {
    Ok(match format {
        Format::Numbers => leaf_bytes!(litchi_numbers::Package::from_bytes(source)?),
        Format::Pages => leaf_bytes!(litchi_pages::Package::from_bytes(source)?),
        Format::Keynote => leaf_bytes!(litchi_keynote::Package::from_bytes(source)?),
    })
}

/// The leaf owner's `save` output.
fn leaf_saved(format: Format, source: &[u8], path: &std::path::Path) -> TestResult<Vec<u8>> {
    match format {
        Format::Numbers => litchi_numbers::Package::from_bytes(source)?.save(path)?,
        Format::Pages => litchi_pages::Package::from_bytes(source)?.save(path)?,
        Format::Keynote => litchi_keynote::Package::from_bytes(source)?.save(path)?,
    }
    Ok(std::fs::read(path)?)
}

fn read_only_case(name: &'static str, format: Format, input: &Fixture) -> TestResult {
    Case::new(name, format, Tier::A, Evidence::Queued, input)
        .read_only()
        .host(|source| Ok(litchi_iwa::Package::from_bytes(source)?.to_bytes()?))
        .leaf(move |source| Ok(LeafRun::output(leaf_output(format, source)?)))
        .run()?;
    Ok(())
}

/// Host raw `Package`: every constructor (`open`, `from_bytes`,
/// `from_shared_bytes`, `from_read_at`), `len`, `is_empty`, `validate`,
/// `to_bytes`, `write_to` and `save`, against `SourceCatalog::{from_bytes,
/// from_shared_bytes, from_read_at}` and the leaf owner's output; `new` is
/// the empty package.
#[test]
fn physical_package_owner_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let empty = litchi_iwa::Package::new();
    assert_eq!(empty.len(), 0);
    assert!(empty.is_empty());
    let directory = scratch("physical")?;
    for (path, format) in [
        (PAGES, Format::Pages),
        (NUMBERS, Format::Numbers),
        (KEYNOTE, Format::Keynote),
    ] {
        let input = fixture(path)?;
        let bytes = input.bytes.as_slice();
        let catalogs = [
            SourceCatalog::from_bytes(bytes, Limits::default())?,
            SourceCatalog::from_shared_bytes(Arc::from(bytes), Limits::default())?,
            SourceCatalog::from_read_at(&SliceSource::new(bytes), Limits::default())?,
        ];
        let source_path = directory.join("source");
        std::fs::write(&source_path, bytes)?;
        let hosts = [
            litchi_iwa::Package::open(&source_path)?,
            litchi_iwa::Package::from_bytes(bytes)?,
            litchi_iwa::Package::from_shared_bytes(Arc::from(bytes))?,
            litchi_iwa::Package::from_read_at(&SliceSource::new(bytes))?,
        ];
        let leaf = leaf_output(format, bytes)?;
        let names = |catalog: &SourceCatalog| {
            catalog
                .package()
                .iter()
                .map(|entry| entry.name().to_owned())
                .collect::<Vec<_>>()
        };
        for catalog in &catalogs {
            assert_eq!(names(catalog), names(&catalogs[0]));
        }
        for host in &hosts {
            host.validate()?;
            assert_eq!(host.len(), catalogs[0].package().len(), "{path}");
            assert_eq!(host.is_empty(), catalogs[0].package().is_empty());
            assert_eq!(
                host.entry_names().map(str::to_owned).collect::<Vec<_>>(),
                names(&catalogs[0])
            );
            assert_eq!(host.to_bytes()?, leaf, "{path}: to_bytes");
            let mut written = Vec::new();
            host.write_to(&mut written)?;
            assert_eq!(written, leaf, "{path}: write_to");
            let saved = directory.join("host-saved");
            host.save(&saved)?;
            assert_eq!(std::fs::read(&saved)?, leaf, "{path}: save");
        }
        assert_eq!(
            leaf_saved(format, bytes, &directory.join("leaf-saved"))?,
            leaf
        );
        read_only_case("physical_package_owner_parity", format, &input)?;
    }
    std::fs::remove_dir_all(&directory)?;
    Ok(())
}

/// Host `pages::Draft::{open, from_bytes, from_package, package,
/// into_package, to_bytes, save}` against `litchi_pages::Package::{open,
/// from_bytes, to_bytes, write_to, save}`.
#[test]
fn pages_package_lifecycle_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = fixture(PAGES)?;
    let directory = scratch("pages")?;
    let source = directory.join("source.pages");
    std::fs::write(&source, &input.bytes)?;
    let leaf = litchi_pages::Package::open(&source)?;
    let leaf_bytes = leaf.to_bytes()?;
    assert_eq!(leaf_bytes, leaf_bytes!(&leaf));
    for host in [
        litchi_iwa::pages::Draft::open(&source)?,
        litchi_iwa::pages::Draft::from_bytes(&input.bytes)?,
        litchi_iwa::pages::Draft::from_package(litchi_iwa::Package::from_bytes(&input.bytes)?)?,
    ] {
        assert_eq!(host.to_bytes()?, leaf_bytes);
        assert_eq!(host.package().to_bytes()?, leaf_bytes);
        let saved = directory.join("host.pages");
        host.save(&saved)?;
        assert_eq!(std::fs::read(&saved)?, leaf_bytes);
        assert_eq!(host.into_package().to_bytes()?, leaf_bytes);
    }
    let saved = directory.join("leaf.pages");
    leaf.save(&saved)?;
    assert_eq!(std::fs::read(&saved)?, leaf_bytes);
    std::fs::remove_dir_all(&directory)?;
    Case::new(
        "pages_package_lifecycle_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(litchi_iwa::pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| {
        Ok(LeafRun::output(
            litchi_pages::Package::from_bytes(source)?.to_bytes()?,
        ))
    })
    .run()?;
    Ok(())
}

/// Host `numbers::Draft::{open, from_bytes, from_package, package,
/// into_package, to_bytes, save}` against `litchi_numbers::Package::{open,
/// from_bytes, write_to, save}`.
#[test]
fn numbers_package_lifecycle_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = fixture(NUMBERS)?;
    let directory = scratch("numbers")?;
    let source = directory.join("source.numbers");
    std::fs::write(&source, &input.bytes)?;
    let leaf = litchi_numbers::Package::open(&source)?;
    let leaf_bytes = leaf_bytes!(&leaf);
    for host in [
        litchi_iwa::numbers::Draft::open(&source)?,
        litchi_iwa::numbers::Draft::from_bytes(&input.bytes)?,
        litchi_iwa::numbers::Draft::from_package(litchi_iwa::Package::from_bytes(&input.bytes)?)?,
    ] {
        assert_eq!(host.to_bytes()?, leaf_bytes);
        assert_eq!(host.package().to_bytes()?, leaf_bytes);
        let saved = directory.join("host.numbers");
        host.save(&saved)?;
        assert_eq!(std::fs::read(&saved)?, leaf_bytes);
        assert_eq!(host.into_package().to_bytes()?, leaf_bytes);
    }
    let saved = directory.join("leaf.numbers");
    leaf.save(&saved)?;
    assert_eq!(std::fs::read(&saved)?, leaf_bytes);
    std::fs::remove_dir_all(&directory)?;
    Case::new(
        "numbers_package_lifecycle_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(litchi_iwa::numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_output(Format::Numbers, source)?)))
    .run()?;
    Ok(())
}

/// The format-neutral host text and comments editors
/// (`text::Draft::{open, from_bytes, to_bytes, save}`,
/// `comments::Draft::{open, from_bytes, from_package, package,
/// into_package, to_bytes, save}`) on all three formats, against the leaf
/// owner of each format.
#[test]
fn shared_draft_lifecycle_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let directory = scratch("shared")?;
    for (path, format) in [
        (PAGES, Format::Pages),
        (NUMBERS, Format::Numbers),
        (KEYNOTE, Format::Keynote),
    ] {
        let input = fixture(path)?;
        let source = directory.join("source");
        std::fs::write(&source, &input.bytes)?;
        let leaf = leaf_output(format, &input.bytes)?;
        let saved = directory.join("saved");
        for text in [
            litchi_iwa::text::Draft::open(&source)?,
            litchi_iwa::text::Draft::from_bytes(&input.bytes)?,
        ] {
            assert_eq!(text.to_bytes()?, leaf, "{path}: text to_bytes");
            text.save(&saved)?;
            assert_eq!(std::fs::read(&saved)?, leaf, "{path}: text save");
        }
        for comments in [
            litchi_iwa::comments::Draft::open(&source)?,
            litchi_iwa::comments::Draft::from_bytes(&input.bytes)?,
            litchi_iwa::comments::Draft::from_package(litchi_iwa::Package::from_bytes(
                &input.bytes,
            )?)?,
        ] {
            assert_eq!(comments.to_bytes()?, leaf, "{path}: comments to_bytes");
            assert_eq!(comments.package().to_bytes()?, leaf);
            comments.save(&saved)?;
            assert_eq!(std::fs::read(&saved)?, leaf, "{path}: comments save");
            assert_eq!(comments.into_package().to_bytes()?, leaf);
        }
        Case::new(
            "shared_draft_lifecycle_parity",
            format,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .read_only()
        .host(|source| Ok(litchi_iwa::text::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(move |source| Ok(LeafRun::output(leaf_output(format, source)?)))
        .run()?;
    }
    std::fs::remove_dir_all(&directory)?;
    Ok(())
}

/// The host `Document` reader's application, validation, snapshot and
/// media reads against the leaf: the format-typed leaf owner, the leaf
/// `Package::validate`, `litchi_iwa_common::media::Catalog` statistics over
/// the package's `Data/` members and the members' bytes. Text: Pages
/// equals `litchi_pages::Package::text`; the host's Keynote dump of every
/// storage contains every line of `litchi_keynote::Package::text` (the leaf
/// reads only text reachable from the show, in presentation order).
#[test]
fn document_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let directory = scratch("document")?;
    for (path, format) in [
        (PAGES, Format::Pages),
        (NUMBERS, Format::Numbers),
        (KEYNOTE, Format::Keynote),
        (
            "iwork/keynote/image-adjustments-native.key",
            Format::Keynote,
        ),
        ("iwork/pages/basic.pages", Format::Pages),
    ] {
        let input = fixture(path)?;
        let source = directory.join("source");
        std::fs::write(&source, &input.bytes)?;
        let opened = litchi_iwa::Document::open(&source)?;
        let document = litchi_iwa::Document::from_bytes(&input.bytes)?;
        let snapshot = document.snapshot();
        // The leaf classification is `litchi_iwa_detect::bytes`, whose
        // `Format` has the variant names of the harness `Format`.
        let expected = format;
        for host in [&opened, &document, &snapshot] {
            host.validate()?;
            let application = format!("{:?}", host.application());
            assert_eq!(application, format!("{expected:?}"), "{path}");
        }
        match format {
            Format::Pages => litchi_pages::Package::from_bytes(&input.bytes)?.validate()?,
            Format::Keynote => litchi_keynote::Package::from_bytes(&input.bytes)?.validate()?,
            Format::Numbers => drop(litchi_numbers::Package::from_bytes(&input.bytes)?),
        }

        let catalog = SourceCatalog::from_bytes(&input.bytes, Limits::default())?;
        let members = catalog
            .package()
            .iter()
            .map(|entry| (entry.name(), entry.data().len() as u64));
        let media = litchi_iwa_common::media::Catalog::from_members(members, Default::default())?;
        let stats = media.stats();
        match document.media_stats() {
            Some(host) => assert_eq!(
                format!("{host:?}").trim_start_matches("Media"),
                format!("{stats:?}"),
                "{path}: stats"
            ),
            None => assert!(media.is_empty(), "{path}: host media manager missing"),
        }
        for asset in media.assets() {
            let member = format!("Data/{}", asset.file_name());
            let expected = catalog
                .package()
                .iter()
                .find(|entry| entry.name() == member)
                .ok_or("member")?
                .data()
                .to_vec();
            assert_eq!(document.extract_media(asset.file_name())?, expected);
            let mut written = Vec::new();
            document.extract_media_to_writer(asset.file_name(), &mut written)?;
            assert_eq!(written, expected);
            let file = directory.join("asset");
            let _ = std::fs::remove_file(&file);
            document.extract_media_to_file(asset.file_name(), &file)?;
            assert_eq!(std::fs::read(&file)?, expected);
        }

        let host_text = document.text()?;
        match format {
            Format::Pages => assert_eq!(
                host_text,
                litchi_pages::Package::from_bytes(&input.bytes)?.text()?,
                "{path}: text"
            ),
            Format::Keynote => {
                let leaf_text = litchi_keynote::Package::from_bytes(&input.bytes)?.text()?;
                for line in leaf_text.lines() {
                    assert!(
                        host_text.lines().any(|host| host == line),
                        "{path}: {line:?}"
                    );
                }
            },
            Format::Numbers => {},
        }
    }
    std::fs::remove_dir_all(&directory)?;
    Ok(())
}

/// Host `numbers::Draft::sheet_text_boxes` (the ordinary text boxes of a
/// sheet, in drawable order, with their text) against
/// `litchi_numbers::Package::sheet_drawable_list` filtered to
/// `Kind::TextBox` and `sheet_drawable_plain_text`.
#[test]
fn numbers_sheet_text_boxes_read_parity() -> TestResult {
    use litchi_numbers::sheet::drawable::Kind;
    let _ = (Tier::A, Evidence::Queued);
    let input = crate::numbers_drawable_cases::authored_fixture()?;
    let host = litchi_iwa::numbers::Draft::from_bytes(&input.bytes)?;
    let sheet = crate::numbers_drawable_cases::first_sheet(&host)?;
    let host_texts = host
        .sheet_text_boxes(sheet)?
        .into_iter()
        .map(|info| info.storage.storage.text().to_owned())
        .collect::<Vec<_>>();
    let package = litchi_numbers::Package::from_bytes(&input.bytes)?;
    let mut leaf_texts = Vec::new();
    for item in package.sheet_drawable_list(0)? {
        if item.kind() == Kind::TextBox {
            leaf_texts.push(package.sheet_drawable_plain_text(0, item.selector())?);
        }
    }
    assert!(!host_texts.is_empty());
    assert_eq!(host_texts, leaf_texts);
    Case::new(
        "numbers_sheet_text_boxes_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(litchi_iwa::numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_output(Format::Numbers, source)?)))
    .run()?;
    Ok(())
}
