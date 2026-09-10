#![allow(
    clippy::expect_used,
    clippy::pedantic,
    clippy::shadow_reuse,
    clippy::unwrap_used,
    reason = "source theme integration tests use checked fixtures and panic-on-failure assertions"
)]

use std::io::{self, Cursor};
use std::path::Path;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits, ReadAt,
    SourceVersion,
};
use litchi_opc::{OpcError, OpcPackage, PackURI};
use litchi_xlsb::package::PackageError;
use litchi_xlsb::{ReadLimits, SourceBackedWorkbook};

#[derive(Debug)]
struct Source {
    bytes: Vec<u8>,
    reads: AtomicUsize,
    revision: AtomicU64,
}

impl Source {
    fn fixture() -> Arc<Self> {
        Arc::new(Self {
            bytes: std::fs::read(
                Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ooxml/xlsb/62815.xlsb"),
            )
            .expect("native fixture"),
            reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
        })
    }
}

impl ReadAt for Source {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5448_454d,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "theme-source-test",
        BudgetLimits::new(
            64 << 20,
            64 << 20,
            64 << 20,
            1_000_000,
            1_000_000,
            1_000_000,
        ),
    );
    let (source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).unwrap(),
        std::num::NonZeroUsize::new(1).unwrap(),
        std::num::NonZeroU64::new(1_000_000).unwrap(),
        0,
    )
    .unwrap();
    (source, ExecutionContext::new(budget, cancellation, limits))
}

#[test]
fn stale_source_refuses_theme_without_payload_reads() {
    let source = Source::fixture();
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let before = source.reads.load(Ordering::SeqCst);
    source.revision.fetch_add(1, Ordering::SeqCst);
    assert!(matches!(
        workbook.theme(),
        Err(PackageError::Opc(OpcError::SourceChanged { .. }))
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), before);
}

#[test]
fn cancelled_source_refuses_theme_without_payload_reads() {
    let source = Source::fixture();
    let (cancellation, context) = context();
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let before = source.reads.load(Ordering::SeqCst);
    cancellation.cancel();
    assert!(matches!(
        workbook.theme(),
        Err(PackageError::Opc(OpcError::Cancelled))
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), before);
}

#[test]
fn theme_access_materializes_only_the_selected_theme_part() {
    let source = Source::fixture();
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let before = workbook.cache_diagnostics();
    assert!(workbook.theme().unwrap().is_some());
    let after = workbook.cache_diagnostics();
    assert_eq!(after.cold_loads - before.cold_loads, 1);
    assert_eq!(after.retained_entries - before.retained_entries, 1);
    assert_eq!(after.retained_bytes - before.retained_bytes, 7_646);
    let reads = source.reads.load(Ordering::SeqCst);
    assert!(workbook.theme().unwrap().is_some());
    assert_eq!(source.reads.load(Ordering::SeqCst), reads);
    assert_eq!(workbook.cache_diagnostics().cold_loads, after.cold_loads);
}

#[test]
fn malformed_unselected_worksheet_does_not_block_theme_access() {
    let source = Source::fixture();
    let mut package = OpcPackage::from_reader(Cursor::new(&source.bytes)).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
        .unwrap()
        .set_blob(vec![0xff]);
    let mut bytes = Vec::new();
    package.to_stream(&mut bytes).unwrap();
    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
    assert!(workbook.theme().unwrap().is_some());
    assert!(
        workbook
            .worksheet_by_index(0)
            .unwrap()
            .unwrap()
            .materialize()
            .is_err()
    );
}

#[test]
fn theme_byte_limit_is_checked_before_payload_reads() {
    let source = Source::fixture();
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let before = source.reads.load(Ordering::SeqCst);
    let limits = litchi_xlsb::theme::Limits::new(7_645, 100_000, 128);
    assert!(matches!(
        workbook.theme_with_limits(limits),
        Err(PackageError::LimitExceeded {
            actual: 7_646,
            maximum: 7_645,
            ..
        })
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), before);
    let exact = litchi_xlsb::theme::Limits::new(7_646, 100_000, 128);
    assert!(workbook.theme_with_limits(exact).unwrap().is_some());
}

#[test]
fn managed_theme_view_retains_captured_data_after_workbook_drop() {
    let source = Source::fixture();
    let (_cancellation, context) = context();
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let view = workbook.theme().unwrap().unwrap();
    assert!(workbook.cache_diagnostics().budget_managed);
    let retained = view.clone();
    drop(view);
    drop(workbook);
    let reads = source.reads.load(Ordering::SeqCst);
    assert_eq!(retained.source_xml().len(), 7_646);
    assert_eq!(retained.theme().name, "Office Theme");
    assert_eq!(retained.theme().fonts.major().latin, "Cambria");
    assert_eq!(retained.theme().fonts.minor().latin, "Calibri");
    assert_eq!(source.reads.load(Ordering::SeqCst), reads);
}

#[test]
fn theme_image_relationships_do_not_materialize_image_payloads() {
    let fixture = Source::fixture();
    let mut package = OpcPackage::from_reader(Cursor::new(&fixture.bytes)).unwrap();
    let image_name = PackURI::new("/xl/media/theme-test.png").unwrap();
    package
        .try_add_part(Box::new(litchi_opc::BlobPart::new(
            image_name,
            "image/png".to_owned(),
            vec![0xa5; 1 << 20],
        )))
        .unwrap();
    let theme = package
        .get_part_mut(&PackURI::new("/xl/theme/theme1.xml").unwrap())
        .unwrap();
    theme
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::IMAGE.to_owned(),
            "../media/theme-test.png".to_owned(),
            "rIdImage".to_owned(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();
    theme
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::IMAGE.to_owned(),
            "https://example.invalid/theme-image.png".to_owned(),
            "rIdExternal".to_owned(),
            litchi_opc::TargetMode::External,
        )
        .unwrap();
    let mut bytes = Vec::new();
    package.to_stream(&mut bytes).unwrap();
    let source = Arc::new(Source {
        bytes,
        reads: AtomicUsize::new(0),
        revision: AtomicU64::new(0),
    });
    let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
    let before = workbook.cache_diagnostics();
    let view = workbook.theme().unwrap().unwrap();
    let after = workbook.cache_diagnostics();
    assert_eq!(view.theme().name, "Office Theme");
    assert_eq!(after.cold_loads - before.cold_loads, 1);
    assert_eq!(after.retained_entries - before.retained_entries, 1);
    assert_eq!(after.retained_bytes - before.retained_bytes, 7_646);
}
