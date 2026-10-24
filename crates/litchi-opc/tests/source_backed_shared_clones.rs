#![allow(
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused sharing assertions intentionally fail on fixture errors"
)]

//! A cloned [`SourceBackedPackage`] is the same opened package (record 0012).
//!
//! These tests drive the public API on real-producer fixtures and on one
//! package built here. They check that a clone reads what the original reads,
//! that every clone publishes the bytes a sole owner would publish, any number
//! of times and from several threads at once, that the original stays
//! readable afterwards, and that lineage and source-change fences hold for
//! every clone.

use std::collections::BTreeMap;
use std::io;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, RwLock};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{OpcError, PackURI, Relationships, SourceBackedPackage, SourceTopologyPlan};
use litchi_opc::{SourceCacheLimits, SourceLimits, SourceReadPolicy};
use soapberry_zip::office::StreamingArchiveWriter;

/// Unsigned real-producer fixtures in four formats.
const FIXTURES: &[&str] = &[
    "ooxml/docx/drawing.docx",
    "ooxml/docx/comment.docx",
    "ooxml/pptx/SmartArt.pptx",
    "poi/test-data/slideshow/bug62513.pptx",
    "ooxml/xlsx/ExcelTables.xlsx",
    "ooxml/xlsb/Simple.xlsb",
];
const SIGNED_FIXTURE: &str = "poi/test-data/xmldsign/ms-office-2010-signed.docx";
/// One byte follows this fixture's archive, so the overlay publisher refuses
/// any edit of it before output, for a sole owner and for every clone alike.
const TRAILING_BYTE_FIXTURE: &str = "poi/test-data/slideshow/bug62513.pptx";
const BUILT: &str = "built here";
const COMMENT: &[u8] = b"<!-- record 0012 -->";

fn fixture(name: &str) -> Vec<u8> {
    if name == BUILT {
        return built_package();
    }
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(name);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

/// A package written by this crate's own ZIP writer: stored structural
/// members and deflated parts.
fn built_package() -> Vec<u8> {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored(
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#,
        )
        .unwrap();
    writer
        .write_stored(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        )
        .unwrap();
    writer
        .write_deflated(
            "word/document.xml",
            b"<document><p>shared clones</p></document>",
        )
        .unwrap();
    for index in 0..24 {
        writer
            .write_deflated(
                &format!("custom/item-{index:02}.xml"),
                format!("<item n=\"{index}\">{}</item>", "x".repeat(64 * index)).as_bytes(),
            )
            .unwrap();
    }
    writer
        .write_stored("custom/blob.bin", &[0x5a; 4096])
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn open(bytes: &[u8]) -> SourceBackedPackage {
    SourceBackedPackage::from_vec(
        bytes.to_vec(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .expect("fixture must open")
}

/// One relationship collection in a canonical order.
fn relationships(rels: &Relationships) -> Vec<(String, String, String, bool)> {
    let mut list: Vec<_> = rels
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_string(),
                relationship.reltype().to_string(),
                relationship.target_ref().to_string(),
                relationship.is_external(),
            )
        })
        .collect();
    list.sort();
    list
}

type PartRecord = (String, Vec<(String, String, String, bool)>, Vec<u8>);

/// Every part's name, content type, relationships and payload.
fn read_all(package: &SourceBackedPackage) -> BTreeMap<String, PartRecord> {
    package
        .iter_parts()
        .map(|part| {
            (
                part.partname().to_string(),
                (
                    part.content_type().to_string(),
                    relationships(part.rels()),
                    part.data().expect("part must read").as_bytes().to_vec(),
                ),
            )
        })
        .collect()
}

/// The XML part this test edits: the first by name that is not a
/// relationships part, with a comment appended after its root element.
fn edit(package: &SourceBackedPackage) -> (PackURI, Vec<u8>) {
    let mut parts: Vec<_> = package
        .iter_parts()
        .filter(|part| {
            let content_type = part.content_type();
            content_type.ends_with("xml") && !content_type.contains("relationships")
        })
        .map(|part| part.partname().clone())
        .collect();
    parts.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    let target = parts.first().expect("fixture has an XML part").clone();
    let mut replacement = package
        .part(&target)
        .unwrap()
        .data()
        .unwrap()
        .as_bytes()
        .to_vec();
    replacement.extend_from_slice(COMMENT);
    (target, replacement)
}

/// The bytes an edit publishes, or the text of its refusal. A refused
/// publication must not have written anything.
fn publish_edit(
    package: SourceBackedPackage,
    target: &PackURI,
    replacement: &[u8],
) -> Result<Vec<u8>, String> {
    let mut output = Vec::new();
    match package.write_part_overlay_to_stream(&mut output, target, replacement.to_vec()) {
        Ok(()) => Ok(output),
        Err(error) => {
            assert!(
                output.is_empty(),
                "a refused edit wrote {} bytes",
                output.len()
            );
            Err(error.to_string())
        },
    }
}

/// The bytes a prepared topology publishes, or the text of its refusal.
fn publish_topology(
    package: SourceBackedPackage,
    plan: SourceTopologyPlan,
) -> Result<Vec<u8>, String> {
    let mut output = Vec::new();
    match package
        .prepare_topology(plan)
        .and_then(|prepared| prepared.publish_to_stream(&mut output))
    {
        Ok(()) => Ok(output),
        Err(error) => {
            assert!(
                output.is_empty(),
                "a refused topology wrote {} bytes",
                output.len()
            );
            Err(error.to_string())
        },
    }
}

fn publish_noop(package: SourceBackedPackage) -> Vec<u8> {
    let mut output = Vec::new();
    package
        .write_part_overlays_to_stream(&mut output, Vec::new())
        .expect("an empty overlay set must publish");
    output
}

fn all_fixtures() -> impl Iterator<Item = &'static str> {
    FIXTURES.iter().copied().chain([BUILT])
}

#[test]
fn a_package_is_clone_send_and_sync() {
    fn shareable<T: Clone + Send + Sync + 'static>() {}
    shareable::<SourceBackedPackage>();

    let package = open(&built_package());
    let lineage = package.source_lineage();
    let moved = package.clone();
    let lineage_elsewhere = std::thread::spawn(move || moved.source_lineage())
        .join()
        .unwrap();
    assert_eq!(lineage, lineage_elsewhere);
}

#[test]
fn a_clone_reads_what_the_original_reads() {
    for name in all_fixtures() {
        let bytes = fixture(name);
        let original = open(&bytes);
        let clone = original.clone();
        assert_eq!(original.source_lineage(), clone.source_lineage(), "{name}");
        assert_eq!(
            original.source_version().unwrap(),
            clone.source_version().unwrap(),
            "{name}"
        );
        assert_eq!(original.read_limits(), clone.read_limits(), "{name}");
        assert_eq!(
            relationships(original.rels()),
            relationships(clone.rels()),
            "{name}"
        );
        assert_eq!(
            original.non_part_members(),
            clone.non_part_members(),
            "{name}"
        );
        assert_eq!(
            original.tolerance_report(),
            clone.tolerance_report(),
            "{name}"
        );

        // Read through the clone first (cold loads), then through the original
        // (cache hits), and compare both with a separate open: sharing the
        // cache must not change a value.
        let through_clone = read_all(&clone);
        let through_original = read_all(&original);
        assert_eq!(through_clone, through_original, "{name}");
        let reference = read_all(&open(&bytes));
        assert_eq!(through_original, reference, "{name}");
        let late_clone = original.clone();
        assert_eq!(read_all(&late_clone), reference, "{name}");
        assert_eq!(
            original.cache_diagnostics(),
            late_clone.cache_diagnostics(),
            "{name}"
        );
    }
}

#[test]
fn an_exact_no_op_from_any_clone_is_the_source() {
    for name in all_fixtures().chain([SIGNED_FIXTURE]) {
        let bytes = fixture(name);
        let original = open(&bytes);
        let clone = original.clone();
        assert_eq!(publish_noop(clone.clone()), bytes, "{name}");
        assert_eq!(publish_noop(clone.clone().clone()), bytes, "{name}");
        let mut prepared = Vec::new();
        clone
            .clone()
            .prepare_topology(SourceTopologyPlan::default())
            .unwrap()
            .publish_to_stream(&mut prepared)
            .unwrap();
        assert_eq!(prepared, bytes, "{name}");
        let mut artifact = Vec::new();
        clone.source_artifact().write_to(&mut artifact).unwrap();
        assert_eq!(artifact, bytes, "{name}");
        drop(clone);
        assert_eq!(publish_noop(original), bytes, "{name}");
    }
}

/// Record 0009 T1–T3 at the OPC level: a publication from a clone leaves the
/// original readable and byte-identical, a second publication is not stale,
/// and the outcome does not depend on how many clones or views are alive.
///
/// A fixture the overlay publisher refuses for a sole owner must be refused
/// identically through every clone.
#[test]
fn clones_publish_independently_and_as_a_sole_owner_would() {
    let mut refused = Vec::new();
    for name in all_fixtures() {
        let bytes = fixture(name);
        let (target, replacement) = edit(&open(&bytes));
        let sole = publish_edit(open(&bytes), &target, &replacement);
        if sole.is_err() {
            refused.push(name);
        }
        if let Ok(sole) = &sole {
            assert_ne!(sole, &bytes, "{name}: the edit must change the package");
            let republished = read_all(&open(sole));
            assert_eq!(
                republished[target.as_str()].2,
                replacement,
                "{name}: the output carries the edit"
            );
        }

        let original = open(&bytes);
        let before = read_all(&original);
        let view = original.part(&target).unwrap();
        let first = publish_edit(original.clone(), &target, &replacement);
        let second = publish_edit(original.clone(), &target, &replacement);
        let kept = original.clone();
        let third = publish_edit(original.clone(), &target, &replacement);
        assert_eq!(first, sole, "{name}: first publication from a clone");
        assert_eq!(second, sole, "{name}: second publication is not stale");
        assert_eq!(third, sole, "{name}: a live clone does not change output");

        assert_eq!(
            view.data().unwrap().as_bytes(),
            before[target.as_str()].2.as_slice()
        );
        assert_eq!(read_all(&original), before, "{name}: original unchanged");
        assert_eq!(read_all(&kept), before, "{name}: kept clone unchanged");
        assert_eq!(publish_noop(kept), bytes, "{name}");
        assert_eq!(publish_noop(original.clone()), bytes, "{name}");
        assert_eq!(
            publish_edit(original, &target, &replacement),
            sole,
            "{name}"
        );
    }
    assert_eq!(refused, [TRAILING_BYTE_FIXTURE]);
}

#[test]
fn clones_share_one_lineage_and_a_reopen_mints_another() {
    let bytes = built_package();
    let original = open(&bytes);
    let clone = original.clone();
    let reopened = open(&bytes);
    assert_eq!(original.source_lineage(), clone.source_lineage());
    assert_eq!(original.source_lineage(), clone.clone().source_lineage());
    assert_ne!(original.source_lineage(), reopened.source_lineage());
    assert_ne!(clone.source_lineage(), reopened.source_lineage());
    drop(original);
    assert_ne!(clone.source_lineage(), reopened.source_lineage());
}

/// A positional source that a test can change in place.
struct MutableSource {
    bytes: RwLock<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn change_in_place(&self) {
        let mut bytes = self.bytes.write().unwrap();
        let last = bytes.len() - 1;
        bytes[last] ^= 1;
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.read().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.read().unwrap();
        let start = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if start >= bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - start);
        output[..count].copy_from_slice(&bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x0012_0012_0012_0012,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

fn is_source_changed(error: &OpcError) -> bool {
    match error {
        OpcError::SourceChanged { .. } => true,
        OpcError::IncompleteOutput { source, .. } => is_source_changed(source),
        _ => false,
    }
}

#[test]
fn an_in_place_source_change_refuses_every_clone() {
    let source = Arc::new(MutableSource {
        bytes: RwLock::new(built_package()),
        revision: AtomicU64::new(0),
    });
    let original = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let clone = original.clone();
    let (target, replacement) = edit(&original);
    // Warm the shared cache first: a cached payload is still fenced.
    let cached = original.part(&target).unwrap().data().unwrap();
    assert!(!cached.as_bytes().is_empty());
    drop(cached);

    source.change_in_place();
    for package in [&original, &clone] {
        let error = package
            .part(&target)
            .and_then(|part| part.data())
            .unwrap_err();
        assert!(is_source_changed(&error), "{error:?}");
        let error = package.source_version().unwrap_err();
        assert!(is_source_changed(&error), "{error:?}");

        let mut output = Vec::new();
        let error = package
            .clone()
            .write_part_overlays_to_stream(&mut output, Vec::new())
            .unwrap_err();
        assert!(is_source_changed(&error), "{error:?}");
        let mut output = Vec::new();
        let error = package
            .clone()
            .write_part_overlay_to_stream(&mut output, &target, replacement.clone())
            .unwrap_err();
        assert!(is_source_changed(&error), "{error:?}");
        assert!(output.is_empty(), "a refused edit publishes nothing");
    }
}

#[test]
fn clones_read_and_publish_concurrently() {
    for name in ["poi/test-data/slideshow/45545_Comment.pptx", BUILT] {
        let bytes = fixture(name);
        let reference = read_all(&open(&bytes));
        let (target, replacement) = edit(&open(&bytes));
        let sole = publish_edit(open(&bytes), &target, &replacement);
        assert!(sole.is_ok(), "{name}: {sole:?}");
        let package = open(&bytes);
        std::thread::scope(|scope| {
            for worker in 0..8 {
                let clone = package.clone();
                let (bytes, reference, target, replacement, sole) =
                    (&bytes, &reference, &target, &replacement, &sole);
                scope.spawn(move || {
                    for round in 0..3 {
                        if (worker + round) % 2 == 0 {
                            assert_eq!(&read_all(&clone), reference, "{name}");
                            assert_eq!(&publish_edit(clone.clone(), target, replacement), sole);
                        } else {
                            assert_eq!(&publish_edit(clone.clone(), target, replacement), sole);
                            assert_eq!(&read_all(&clone), reference, "{name}");
                        }
                        assert_eq!(&publish_noop(clone.clone()), bytes, "{name}");
                    }
                });
            }
        });
        assert_eq!(read_all(&package), reference, "{name}");
        assert_eq!(publish_noop(package), bytes, "{name}");
    }
}

/// Record 0009 T6 at the OPC level: replacing the path atomically leaves every
/// clone on the file it opened.
#[cfg(unix)]
#[test]
fn an_atomic_replacement_leaves_every_clone_on_the_opened_file() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("package.docx");
    let bytes = fixture("ooxml/docx/drawing.docx");
    std::fs::write(&path, &bytes).unwrap();
    let original = SourceBackedPackage::from_path(&path, SourceLimits::default()).unwrap();
    let clone = original.clone();
    let before = read_all(&original);

    let replacement = built_package();
    let staged = directory.path().join("package.docx.tmp");
    std::fs::write(&staged, &replacement).unwrap();
    std::fs::rename(&staged, &path).unwrap();

    for package in [&original, &clone] {
        assert_eq!(read_all(package), before);
        assert_eq!(publish_noop(package.clone()), bytes);
    }
    let reopened = SourceBackedPackage::from_path(&path, SourceLimits::default()).unwrap();
    assert_eq!(read_all(&reopened), read_all(&open(&replacement)));
    assert_ne!(reopened.source_lineage(), clone.source_lineage());
}

#[test]
fn materializing_a_clone_leaves_the_package_readable() {
    for name in all_fixtures() {
        let bytes = fixture(name);
        let original = open(&bytes);
        let before = read_all(&original);
        let materialized = original.clone().into_opc_package().unwrap();
        assert_eq!(materialized.part_count(), before.len(), "{name}");
        assert_eq!(
            materialized.non_part_members(),
            original.non_part_members(),
            "{name}"
        );
        assert_eq!(read_all(&original), before, "{name}");
        let borrowed = original.to_opc_package().unwrap();
        assert_eq!(borrowed.part_count(), before.len(), "{name}");
        assert_eq!(publish_noop(original.clone()), bytes, "{name}");
        let last = original.into_opc_package().unwrap();
        assert_eq!(last.non_part_members(), borrowed.non_part_members());
    }
}

#[test]
fn a_prepared_topology_from_a_clone_leaves_the_package_readable() {
    let mut refused = Vec::new();
    for name in all_fixtures() {
        let bytes = fixture(name);
        let (target, replacement) = edit(&open(&bytes));
        let plan = || {
            let mut plan = SourceTopologyPlan::default();
            plan.try_replace_part(target.clone(), replacement.clone())
                .unwrap();
            plan
        };
        let sole = publish_topology(open(&bytes), plan());
        if sole.is_err() {
            refused.push(name);
        }

        let original = open(&bytes);
        let before = read_all(&original);
        if sole.is_ok() {
            // Two preparations hold two clones at once; the candidate of one
            // shows the edit while the package still reads the source.
            let first = original.clone().prepare_topology(plan()).unwrap();
            let second = original.clone().prepare_topology(plan()).unwrap();
            first
                .with_candidate(|candidate| {
                    let part = candidate.part(&target)?;
                    assert_eq!(part.data()?.as_bytes(), replacement.as_slice());
                    Ok(())
                })
                .unwrap();
            assert_eq!(read_all(&original), before, "{name}");
            let mut output = Vec::new();
            second.publish_to_stream(&mut output).unwrap();
            assert_eq!(Ok(output), sole, "{name}");
            let mut output = Vec::new();
            first.publish_to_stream(&mut output).unwrap();
            assert_eq!(Ok(output), sole, "{name}");
        }
        assert_eq!(publish_topology(original.clone(), plan()), sole, "{name}");
        assert_eq!(read_all(&original), before, "{name}");
        assert_eq!(publish_noop(original), bytes, "{name}");
    }
    assert_eq!(refused, [TRAILING_BYTE_FIXTURE]);
}
