use super::*;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Resource,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, ReadLimits};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{Arc, Barrier};
use std::thread;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const FINITE_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OBJECTS: u64 = 1_000_000;
const FINITE_DEPTH: u64 = 1024;
const FINITE_WORK: u64 = 1 << 30;

fn fixture_bytes() -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>managed</w:t></w:r></w:p><w:p><w:r><w:t>read</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document,
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn managed_package() -> (Budget, CancellationSource, Package) {
    managed_package_with_memory(1 << 20)
}

fn managed_package_with_memory(memory: u64) -> (Budget, CancellationSource, Package) {
    let budget = Budget::root(
        "docx-paragraph-index-private-test",
        Limits::new(
            memory,
            FINITE_INPUT_BYTES,
            FINITE_OUTPUT_BYTES,
            FINITE_OBJECTS,
            FINITE_DEPTH,
            FINITE_WORK,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let package = Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(fixture_bytes())),
        ReadLimits::default(),
        ExecutionContext::new(budget.clone(), cancellation, execution_limits),
    )
    .unwrap();
    (budget, cancellation_source, package)
}

#[test]
fn managed_document_views_reuse_one_index_after_the_first_build() {
    let (budget, _cancellation_source, package) = managed_package();

    let first = package.document().unwrap();
    assert_eq!(first.paragraph_count().unwrap(), 2);
    drop(first);
    let work_after_first = budget.used(Resource::Work);

    let second = package.document().unwrap();
    assert_eq!(second.paragraph_count().unwrap(), 2);
    drop(second);
    assert_eq!(budget.used(Resource::Work), work_after_first);

    let diagnostics = package.paragraph_index_cache.diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert_eq!(diagnostics.hits, 1);
    assert_eq!(diagnostics.misses, 1);
    assert_eq!(diagnostics.clean_entries, 1);
    assert_eq!(diagnostics.pinned_entries, 0);

    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn concurrent_first_managed_views_coalesce_to_one_index_build() {
    let (budget, _cancellation_source, package) = managed_package();
    let package = Arc::new(package);
    let barrier = Arc::new(Barrier::new(4));

    let handles = (0..4)
        .map(|_| {
            let package = Arc::clone(&package);
            let barrier = Arc::clone(&barrier);
            thread::spawn(move || {
                barrier.wait();
                let document = package.document().unwrap();
                assert_eq!(document.paragraph_count().unwrap(), 2);
            })
        })
        .collect::<Vec<_>>();
    for handle in handles {
        handle.join().unwrap();
    }

    let diagnostics = package.paragraph_index_cache.diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert_eq!(diagnostics.hits, 3);
    assert_eq!(diagnostics.misses, 1);
    assert_eq!(diagnostics.clean_entries, 1);
    assert_eq!(diagnostics.pinned_entries, 0);

    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_index_reservations_last_until_the_final_view_drops() {
    let (budget, _cancellation_source, package) = managed_package();
    let first = package.document().unwrap();
    let second = first.clone();
    assert!(budget.used(Resource::Memory) > 0);

    drop(package);
    assert!(budget.used(Resource::Memory) > 0);
    drop(first);
    assert!(budget.used(Resource::Memory) > 0);
    drop(second);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_after_warm_managed_index_evicts_the_clean_entry() {
    let (budget, cancellation_source, package) = managed_package();
    let document = package.document().unwrap();
    assert_eq!(document.paragraph_count().unwrap(), 2);
    drop(document);

    let warm = package.paragraph_index_cache.diagnostics();
    assert_eq!(warm.builds, 1);
    assert_eq!(warm.clean_entries, 1);
    assert_eq!(warm.pinned_entries, 0);

    cancellation_source.cancel();
    assert!(matches!(
        package.document(),
        Err(Error::Opc(OpcError::Cancelled))
    ));

    let after_cancel = package.paragraph_index_cache.diagnostics();
    assert_eq!(after_cancel.builds, 1);
    assert_eq!(after_cancel.clean_entries, 0);
    assert!(after_cancel.clean_evictions >= 1);
    assert_eq!(after_cancel.pinned_entries, 0);

    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn delayed_managed_tail_append_publishes_after_clean_index_warm() {
    let (budget, _cancellation_source, package) = managed_package_with_memory(8 << 20);
    let plan = package
        .tail_append_plain_paragraph("tail")
        .with_limits(tail_append::Limits::new(
            4096,
            1024,
            4096,
            8192,
            1000,
            16,
            10,
            4096,
            1 << 20,
            1 << 20,
            1024,
        ))
        .prepare()
        .unwrap();

    let warm = package.document().unwrap();
    assert_eq!(warm.paragraph_count().unwrap(), 2);
    drop(warm);
    let before_publish = package.paragraph_index_cache.diagnostics();
    assert_eq!(before_publish.clean_entries, 1);
    assert_eq!(before_publish.pinned_entries, 0);

    let mut output = Vec::new();
    plan.write_to_stream(&mut output).unwrap();
    assert!(!output.is_empty());

    let reopened = Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    let document = reopened.document().unwrap();
    assert_eq!(document.paragraph_count().unwrap(), 3);
    assert_eq!(document.extract_text().unwrap(), "managedreadtail");

    let after_publish = package.paragraph_index_cache.diagnostics();
    assert_eq!(after_publish.clean_entries, 0);
    assert_eq!(after_publish.pinned_entries, 0);
    drop(reopened);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}
