//! Publication and execution-policy regressions for inert DDE facades.

mod support;

use std::{
    io,
    num::{NonZeroU64, NonZeroUsize},
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
};

use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionLimits, Limits, OwnedSource,
    Profile, ReadAt, Resource, SourceVersion,
};
use litchi_odf_common::{
    core::{
        SourceContentPublicationError, SourceContentPublicationOptions,
        SourceContentPublicationProgress,
    },
    package::raw_identical_members,
};
use litchi_ods::{MutableSpreadsheet, SourceBackedSpreadsheet, Spreadsheet, dde};

const CONTENT: &str = concat!(
    "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
    "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" office:version=\"1.4\">",
    "<office:body><office:spreadsheet><table:table table:name=\"Data\">",
    "<office:dde-source office:dde-application=\"litchi-inert-fixture\" ",
    "office:dde-topic=\"never-opened\" office:dde-item=\"cached-values\" ",
    "office:automatic-update=\"false\"/>",
    "<table:table-column/>",
    "<table:table-row><table:table-cell office:value-type=\"float\" office:value=\"42\"/>",
    "</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"
);

fn package() -> Vec<u8> {
    support::raw_package(&[
        ("content.xml", CONTENT.as_bytes(), "text/xml"),
        (
            "Pictures/unrelated.bin",
            b"inert-payload",
            "application/octet-stream",
        ),
    ])
}

fn context() -> (CancellationSource, ExecutionContext) {
    let (cancel, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        Budget::root("dde-facade-tests", Limits::for_profile(Profile::Server)),
        token,
        ExecutionLimits::new(
            NonZeroUsize::new(1).expect("worker"),
            NonZeroUsize::new(1).expect("task"),
            NonZeroU64::new(1024).expect("bytes"),
            0,
        )
        .expect("execution limits"),
    );
    (cancel, execution)
}

fn replacement_source() -> dde::Source {
    dde::Source::new("litchi-inert-fixture", "never-opened", "replacement-values")
        .expect("source descriptor")
        .named("DataSource")
        .expect("source name")
        .with_automatic_update(dde::AutomaticUpdate::Disabled)
}

#[test]
fn bom_prefixed_content_keeps_exact_transaction_coordinates() {
    let content = format!("\u{feff}{CONTENT}");
    let source = dde::Snapshot::parse(&content).expect("BOM-prefixed content");
    let mut edit = source.edit();
    edit.set_sheet_source("Data", replacement_source())
        .expect("stage source change");
    let (_cancel, execution) = context();
    let commit = edit.commit(&execution).expect("BOM-prefixed transaction");
    assert!(commit.snapshot().source_xml().starts_with('\u{feff}'));
    assert_eq!(
        commit.snapshot().sheet_sources()[0].source().item(),
        "replacement-values"
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("exact inverse");
    assert_eq!(restored.snapshot().source_xml(), content);
}

#[test]
fn ordinary_and_mutable_noops_preserve_exact_content() {
    let mut ordinary = Spreadsheet::from_bytes(package()).expect("ordinary fixture");
    let snapshot = ordinary.dde().expect("inert snapshot");
    assert_eq!(snapshot.sheet_sources().len(), 1);
    ordinary.edit_dde(|_| Ok(())).expect("ordinary no-op");
    assert_eq!(ordinary.content_xml(), CONTENT);

    let mut mutable = MutableSpreadsheet::from_bytes(package()).expect("mutable fixture");
    mutable.edit_dde(|_| Ok(())).expect("mutable no-op");
    assert_eq!(mutable.spreadsheet().content_xml(), CONTENT);
    assert_eq!(
        mutable.dde().expect("mutable snapshot").source_xml(),
        CONTENT
    );
}

#[test]
fn closure_failure_leaves_both_facades_unchanged() {
    let mut ordinary = Spreadsheet::from_bytes(package()).expect("ordinary fixture");
    let error = ordinary.edit_dde(|_| Err(Error::Unsupported("caller abort".to_owned())));
    assert!(matches!(error, Err(Error::Unsupported(message)) if message == "caller abort"));
    assert_eq!(ordinary.content_xml(), CONTENT);

    let mut mutable = MutableSpreadsheet::from_bytes(package()).expect("mutable fixture");
    assert!(
        mutable
            .edit_dde(|_| Err(Error::Unsupported("caller abort".to_owned())))
            .is_err()
    );
    assert_eq!(mutable.spreadsheet().content_xml(), CONTENT);
}

#[test]
fn cancellation_during_closure_refuses_even_noop_publication() {
    let mut ordinary = Spreadsheet::from_bytes(package()).expect("ordinary fixture");
    let (cancel, execution) = context();
    let result = ordinary.edit_dde_with_context(dde::Limits::default(), &execution, |_| {
        cancel.cancel();
        Ok(())
    });
    assert!(result.is_err());
    assert_eq!(ordinary.content_xml(), CONTENT);
}

#[test]
fn explicit_input_limits_keep_structured_resource_errors() {
    let ordinary = Spreadsheet::from_bytes(package()).expect("ordinary fixture");
    let (_cancel, execution) = context();
    let error = ordinary
        .dde_with(
            dde::Limits::default().with_input_bytes(CONTENT.len() - 1),
            &execution,
        )
        .expect_err("input ceiling");
    assert!(matches!(error, Error::ResourceLimit(_)));
}

#[test]
fn execution_limit_errors_retain_the_callers_resource_and_scope() {
    let ordinary = Spreadsheet::from_bytes(package()).expect("ordinary fixture");
    let (_cancel, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        Budget::root(
            "dde-input-admission",
            Limits::new(1_000_000, 1, 1_000_000, 1000, 100, 1_000_000),
        ),
        token,
        ExecutionLimits::new(
            NonZeroUsize::new(1).expect("worker"),
            NonZeroUsize::new(1).expect("task"),
            NonZeroU64::new(1024).expect("bytes"),
            0,
        )
        .expect("execution limits"),
    );
    let error = ordinary
        .dde_with(dde::Limits::default(), &execution)
        .expect_err("caller input budget must refuse admission");
    let Error::ResourceLimit(limit) = error else {
        panic!("expected structured resource error, got {error:?}");
    };
    assert_eq!(limit.resource, Resource::InputBytes);
    assert_eq!(limit.observed, CONTENT.len() as u64);
    assert_eq!(limit.limit, 1);
    assert_eq!(limit.scope.as_ref(), "dde-input-admission");
}

#[test]
fn ordinary_patch_and_inverse_rehydrate_the_complete_facade() {
    let mut ordinary = Spreadsheet::from_bytes(package()).expect("ordinary fixture");
    let (_cancel, execution) = context();
    let before = ordinary
        .dde_with(dde::Limits::default(), &execution)
        .expect("snapshot");
    let mut edit = before.edit();
    edit.set_sheet_source("Data", replacement_source())
        .expect("stage source");
    let commit = edit.commit(&execution).expect("commit source");
    assert!(commit.changed());
    ordinary
        .apply_dde_patch(commit.patch())
        .expect("apply patch");
    assert_eq!(
        ordinary.dde().expect("readback").sheet_sources()[0]
            .source()
            .item(),
        "replacement-values"
    );
    assert!(ordinary.content_xml().contains("office:value=\"42\""));
    ordinary
        .apply_dde_patch(&commit.patch().inverse())
        .expect("apply inverse");
    assert_eq!(ordinary.content_xml(), CONTENT);

    let mut mutable = MutableSpreadsheet::from_bytes(package()).expect("mutable fixture");
    mutable
        .edit_dde(|edit| edit.set_sheet_source("Data", replacement_source()))
        .expect("mutable edit");
    assert_eq!(
        mutable.dde().expect("mutable readback").sheet_sources()[0]
            .source()
            .item(),
        "replacement-values"
    );
}

#[test]
fn failed_closure_after_staging_does_not_publish() {
    let mut ordinary = Spreadsheet::from_bytes(package()).expect("ordinary fixture");
    let error = ordinary.edit_dde(|edit| {
        edit.set_sheet_source("Data", replacement_source())?;
        Err(Error::Unsupported("abort staged source".to_owned()))
    });
    assert!(error.is_err());
    assert_eq!(ordinary.content_xml(), CONTENT);
}

#[test]
fn signature_owner_refuses_changed_dde_but_retains_exact_noops() {
    let signed = support::raw_package(&[
        ("content.xml", CONTENT.as_bytes(), "text/xml"),
        (
            "META-INF/documentsignatures.xml",
            br#"<ds:document-signatures xmlns:ds="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
            "application/vnd.oasis.opendocument.digital-signature",
        ),
    ]);
    let mut ordinary = Spreadsheet::from_bytes(signed.clone()).expect("signed fixture");
    let error = ordinary.edit_dde(|edit| edit.set_sheet_source("Data", replacement_source()));
    assert!(matches!(error, Err(Error::Unsupported(_))));
    assert_eq!(ordinary.content_xml(), CONTENT);
    ordinary.edit_dde(|_| Ok(())).expect("signed no-op");

    let mut mutable = MutableSpreadsheet::from_bytes(signed).expect("signed mutable fixture");
    assert!(
        mutable
            .edit_dde(|edit| edit.set_sheet_source("Data", replacement_source()))
            .is_err()
    );
    assert_eq!(mutable.spreadsheet().content_xml(), CONTENT);
}

#[test]
fn source_backed_commit_reopens_preserves_members_and_inverts() {
    let original = package();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(original.clone())))
        .expect("source-backed fixture");
    let (_cancel, execution) = context();
    let before = owner
        .dde_with(dde::Limits::default(), &execution)
        .expect("source snapshot");
    let mut edit = before.edit().expect("source edit");
    edit.set_sheet_source("Data", replacement_source())
        .expect("stage source");
    let commit = edit.commit(&execution).expect("source commit");
    assert!(commit.changed());
    let applied = commit.patch().apply(&before).expect("apply source patch");
    let via_owner = owner
        .apply_dde_patch(commit.patch())
        .expect("owner patch forwarding");
    assert_eq!(
        via_owner.snapshot().source_xml(),
        applied.snapshot().source_xml()
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(applied.snapshot())
        .expect("source inverse");
    assert_eq!(restored.snapshot().source_xml(), CONTENT);

    let mut output = Vec::new();
    commit
        .write_to(&mut output, SourceContentPublicationOptions::new())
        .expect("sequential publish");
    let identical = raw_identical_members(&original, &output).expect("member comparison");
    assert!(identical.contains("Pictures/unrelated.bin"));
    assert!(identical.contains("META-INF/manifest.xml"));
    let reopened = Spreadsheet::from_bytes(output).expect("reopen publication");
    assert_eq!(
        reopened.dde().expect("readback").sheet_sources()[0]
            .source()
            .item(),
        "replacement-values"
    );
}

#[test]
fn source_patch_rechecks_destination_output_limit_and_live_owner() {
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package())))
        .expect("source owner");
    let (_cancel, execution) = context();
    let before = owner
        .dde_with(dde::Limits::default(), &execution)
        .expect("source snapshot");
    let mut edit = before.edit().expect("source edit");
    edit.set_sheet_source("Data", replacement_source())
        .expect("stage source");
    let commit = edit.commit(&execution).expect("source commit");
    let limited = owner
        .dde_with(dde::Limits::default().with_output_bytes(1), &execution)
        .expect("limited destination");
    assert!(matches!(
        commit.patch().apply(&limited),
        Err(Error::ResourceLimit(_))
    ));

    let second = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package())))
        .expect("second owner");
    let second_snapshot = second.dde().expect("second snapshot");
    assert!(commit.patch().apply(&second_snapshot).is_err());
}

struct MutableSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let Some(bytes) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let count = bytes.len().min(output.len());
        output[..count].copy_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4444_4553,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[test]
fn stale_source_refuses_reads_patch_application_and_publication() {
    let source = Arc::new(MutableSource {
        bytes: package(),
        revision: AtomicU64::new(0),
    });
    let owner =
        SourceBackedSpreadsheet::from_read_at(source.clone()).expect("mutable source owner");
    let (_cancel, execution) = context();
    let before = owner
        .dde_with(dde::Limits::default(), &execution)
        .expect("source snapshot");
    let mut edit = before.edit().expect("source edit");
    edit.set_sheet_source("Data", replacement_source())
        .expect("stage source");
    let commit = edit.commit(&execution).expect("source commit");
    source.revision.fetch_add(1, Ordering::Relaxed);
    assert!(before.sheet_sources().is_err());
    assert!(commit.patch().apply(&before).is_err());
    let mut output = Vec::new();
    let error = commit
        .write_to(&mut output, SourceContentPublicationOptions::new())
        .expect_err("stale publication");
    assert!(matches!(
        error,
        SourceContentPublicationError::SourceChanged {
            progress: SourceContentPublicationProgress::Untouched,
            ..
        }
    ));
    assert!(output.is_empty());
}
