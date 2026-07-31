use litchi_ooxml::{OoxmlError, pptx, xlsx};
use tempfile::{NamedTempFile, TempDir};

#[test]
fn opened_xlsx_cannot_enter_the_empty_legacy_writer() {
    let source = NamedTempFile::with_suffix(".xlsx").expect("temporary source");
    let mut created = xlsx::Workbook::create().expect("fresh workbook");
    created
        .worksheet_mut(0)
        .expect("fresh worksheet")
        .set_cell_value(1, 1, "preserve me");
    created.save(source.path()).expect("save source");

    let mut opened = xlsx::Workbook::open(source.path()).expect("open workbook");
    let error = opened
        .worksheet_mut(0)
        .expect_err("opened writer must fail safely");
    assert!(matches!(
        error.downcast_ref::<OoxmlError>(),
        Some(OoxmlError::UnsafeEdit {
            operation: "worksheet_mut",
            ..
        })
    ));
}

#[test]
fn opened_xlsx_rebuild_is_rejected_before_destination_creation() {
    let source = NamedTempFile::with_suffix(".xlsx").expect("temporary source");
    let mut created = xlsx::Workbook::create().expect("fresh workbook");
    created.save(source.path()).expect("save source");

    let mut opened = xlsx::Workbook::open(source.path()).expect("open workbook");
    opened.add_worksheet("would discard the source");
    let directory = TempDir::new().expect("temporary destination directory");
    let destination = directory.path().join("guarded.xlsx");
    let error = opened
        .save(&destination)
        .expect_err("destructive rebuild must be rejected");

    assert!(matches!(
        error.downcast_ref::<OoxmlError>(),
        Some(OoxmlError::UnsafeEdit {
            operation: "save",
            ..
        })
    ));
    assert!(!destination.exists());
}

#[test]
fn opened_pptx_cannot_enter_the_empty_legacy_writer() {
    let source = NamedTempFile::with_suffix(".pptx").expect("temporary source");
    let mut created = pptx::Package::new().expect("fresh presentation");
    created
        .presentation_mut()
        .expect("fresh presentation writer")
        .add_slide()
        .expect("add slide");
    created.save(source.path()).expect("save source");

    let mut opened = pptx::Package::open(source.path()).expect("open presentation");
    assert!(matches!(
        opened.presentation_mut(),
        Err(OoxmlError::UnsafeEdit {
            operation: "presentation_mut",
            ..
        })
    ));
    assert_eq!(
        opened
            .presentation()
            .expect("read presentation")
            .slide_count()
            .expect("read slide count"),
        1
    );
}
