use std::io::{self, Cursor};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};

use litchi_cfb::{OleError, SharedOleFile, writer::OleWriter};
use litchi_core::{CheckStatus, OwnedSource, ReadAt, SourceVersion};
use litchi_xls::ClipboardFormat;
use litchi_xls::validation::{
    XlsValidationError, XlsValidationLimits, validate_source, validate_source_with_limits,
};
use litchi_xls::writer::{
    AddInFunctionOptions, DdeOrOleItemOptions, DdeOrOleLinkOptions, ExternalDefinedNameOptions,
    ExternalSheetOptions, ExternalWorkbookOptions, WeakEncryptionPolicy, Writer,
};

fn authored(protected: bool, encrypted: bool) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").expect("worksheet");
    writer.write_number(sheet, 0, 0, 42.0).expect("number");
    if protected {
        writer.protect_workbook(Some("secret"), true, false);
        writer
            .protect_sheet(sheet, Some("secret"), true, false)
            .expect("sheet protection");
    }
    if encrypted {
        writer
            .set_xor_obfuscation_password("secret", WeakEncryptionPolicy::allow_xor_obfuscation())
            .expect("XOR protection");
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write XLS");
    output.into_inner()
}

fn authored_external() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_worksheet("Sheet1").expect("worksheet");
    let book = writer
        .add_external_workbook_link(ExternalWorkbookOptions {
            encoded_virtual_path: "\u{1}\u{2}Book.xls".to_string(),
            sheets: vec![ExternalSheetOptions {
                name: "Remote".to_string(),
                cache_rows: Vec::new(),
            }],
        })
        .expect("external workbook");
    writer
        .add_external_defined_name(
            book,
            ExternalDefinedNameOptions {
                name: "RemoteName".to_string(),
                sheet_index: Some(0),
                built_in: false,
                formula_bytes: vec![0x1c, 0x17],
            },
        )
        .expect("external name");
    writer
        .add_add_in_function(AddInFunctionOptions {
            name: "ISODD".to_string(),
            unused_data: vec![0x1c, 0x17],
        })
        .expect("add-in function");
    writer
        .add_dde_or_ole_link(DdeOrOleLinkOptions {
            encoded_virtual_path: "\u{6587}\u{3}System".to_string(),
            items: vec![DdeOrOleItemOptions {
                name: "Object".to_string(),
                automatic: false,
                picture: false,
                standard_document_name: false,
                ole_link: false,
                clipboard_format: ClipboardFormat::Text,
                displayed_as_icon: false,
                storage_id: 0,
                matrix: None,
            }],
        })
        .expect("DDE link");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write XLS");
    output.into_inner()
}

fn authored_encrypted_with_write_reservation() -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").expect("worksheet");
    writer.write_number(sheet, 0, 0, 42.0).expect("number");
    writer
        .set_file_sharing(true, Some("write"), "reviewer")
        .expect("file sharing");
    writer
        .set_xor_obfuscation_password("secret", WeakEncryptionPolicy::allow_xor_obfuscation())
        .expect("XOR protection");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write XLS");
    output.into_inner()
}

fn source(bytes: Vec<u8>) -> Arc<dyn ReadAt> {
    Arc::new(OwnedSource::new(bytes))
}

#[derive(Clone)]
struct ReadWidthSource {
    bytes: Arc<Vec<u8>>,
    maximum_request: usize,
    maximum_observed: Arc<Mutex<usize>>,
    ranges: Arc<Mutex<Vec<(u64, usize)>>>,
    failure_range: Option<(u64, usize)>,
    change_range: Option<(u64, usize)>,
    revision: Arc<AtomicU64>,
}

impl ReadWidthSource {
    fn new(bytes: Vec<u8>, maximum_request: usize) -> Self {
        Self {
            bytes: Arc::new(bytes),
            maximum_request,
            maximum_observed: Arc::new(Mutex::new(0)),
            ranges: Arc::new(Mutex::new(Vec::new())),
            failure_range: None,
            change_range: None,
            revision: Arc::new(AtomicU64::new(0)),
        }
    }

    fn with_failure_range(mut self, range: (u64, usize)) -> Self {
        self.failure_range = Some(range);
        self
    }

    fn with_change_range(mut self, range: (u64, usize)) -> Self {
        self.change_range = Some(range);
        self
    }

    fn maximum_observed(&self) -> usize {
        *self.maximum_observed.lock().unwrap()
    }

    fn ranges(&self) -> Vec<(u64, usize)> {
        self.ranges.lock().unwrap().clone()
    }
}

impl ReadAt for ReadWidthSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let mut maximum_observed = self.maximum_observed.lock().unwrap();
        *maximum_observed = (*maximum_observed).max(output.len());
        if !output.is_empty() {
            self.ranges.lock().unwrap().push((offset, output.len()));
        }
        if output.len() > self.maximum_request {
            return Err(io::Error::other("read request exceeded test width"));
        }
        let requested = (offset, output.len());
        if self
            .change_range
            .is_some_and(|range| ranges_overlap(requested, range))
        {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
        if self
            .failure_range
            .is_some_and(|range| ranges_overlap(requested, range))
        {
            return Err(io::Error::other("injected streaming read failure"));
        }
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if start >= self.bytes.len() || output.is_empty() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x584c_535f_5445_5354,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn ranges_overlap(left: (u64, usize), right: (u64, usize)) -> bool {
    let left_end = left.0.checked_add(left.1 as u64).unwrap();
    let right_end = right.0.checked_add(right.1 as u64).unwrap();
    left.0 < right_end && right.0 < left_end
}

fn cfb_with_streams(streams: &[(&str, Vec<u8>)], storages: &[&str]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for storage in storages {
        writer.create_storage(&[*storage]).expect("storage");
    }
    for (name, data) in streams {
        writer.create_stream(&[*name], data).expect("stream");
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("CFB");
    output.into_inner()
}

fn workbook_stream(bytes: Vec<u8>) -> Vec<u8> {
    SharedOleFile::open(source(bytes))
        .expect("CFB")
        .open_stream(&["Workbook"])
        .expect("Workbook stream")
}

fn frame(kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(4 + payload.len());
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(u16::try_from(payload.len()).expect("BIFF payload")).to_le_bytes());
    output.extend_from_slice(payload);
    output
}

fn bof_payload(dt: u16, length: usize) -> Vec<u8> {
    let mut payload = vec![0; length];
    payload[0..2].copy_from_slice(&0x0600u16.to_le_bytes());
    payload[2..4].copy_from_slice(&dt.to_le_bytes());
    payload
}

fn external_supbook_payload() -> Vec<u8> {
    vec![1, 0, 3, 0, 0, b'r', b'e', b'm', 1, 0, 0, b'S']
}

fn external_cache_row(first_column: u8, last_column: u8, row: u16) -> Vec<u8> {
    let mut payload = vec![last_column, first_column];
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&[0; 9]);
    payload
}

fn unsupported_cryptoapi_filepass() -> Vec<u8> {
    let mut data = vec![1, 0];
    data.extend_from_slice(&4u16.to_le_bytes());
    data.extend_from_slice(&2u16.to_le_bytes());
    data.extend_from_slice(&4u32.to_le_bytes());
    data.extend_from_slice(&32u32.to_le_bytes());
    data.extend_from_slice(&4u32.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&0x660eu32.to_le_bytes());
    data.extend_from_slice(&0x8004u32.to_le_bytes());
    data.extend_from_slice(&128u32.to_le_bytes());
    data.extend_from_slice(&1u32.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&16u32.to_le_bytes());
    data.extend_from_slice(&[0x11; 16]);
    data.extend_from_slice(&[0x22; 16]);
    data.extend_from_slice(&20u32.to_le_bytes());
    data.extend_from_slice(&[0x33; 20]);
    data
}

fn dde_supbook_payload() -> Vec<u8> {
    vec![0, 0, 3, 0, 0, b'd', b'd', b'e']
}

fn conflicting_dde_name_payload() -> Vec<u8> {
    let mut payload = vec![0x18, 0, 0, 0, 0, 0, 15, 0];
    payload.extend_from_slice(b"StdDocumentName");
    payload
}

fn dde_pending_name_payload() -> Vec<u8> {
    let mut payload = vec![0; 6];
    payload.extend_from_slice(&[4, 0, b'I', b't', b'e', b'm']);
    payload.extend_from_slice(&[0, 0, 0]);
    payload
}

fn minimal_workbook(
    before_bound: &[Vec<u8>],
    after_bound: &[Vec<u8>],
    worksheet_bof: &[u8],
    worksheet_tail: &[Vec<u8>],
    trailing: &[Vec<u8>],
) -> Vec<u8> {
    let mut preamble = frame(0x0809, &bof_payload(0x0005, 16));
    for record in before_bound {
        preamble.extend_from_slice(record);
    }
    let bound_payload_len = 9;
    let after_bound_len = after_bound.iter().map(Vec::len).sum::<usize>();
    let worksheet_offset = preamble.len() + 4 + bound_payload_len + after_bound_len + 4;
    let mut bound_payload = Vec::with_capacity(bound_payload_len);
    bound_payload.extend_from_slice(
        &u32::try_from(worksheet_offset)
            .expect("worksheet offset")
            .to_le_bytes(),
    );
    bound_payload.extend_from_slice(&[0, 0, 1, 0, b'S']);
    let mut output = preamble;
    output.extend_from_slice(&frame(0x0085, &bound_payload));
    for record in after_bound {
        output.extend_from_slice(record);
    }
    output.extend_from_slice(&frame(0x000A, &[]));
    output.extend_from_slice(&frame(0x0809, worksheet_bof));
    for record in worksheet_tail {
        output.extend_from_slice(record);
    }
    output.extend_from_slice(&frame(0x000A, &[]));
    for record in trailing {
        output.extend_from_slice(record);
    }
    output
}

fn status<'a>(report: &'a litchi_core::ValidateReport, id: &str) -> &'a CheckStatus {
    report
        .checks()
        .iter()
        .find(|check| check.id().as_str() == id)
        .expect("declared check")
        .status()
}

#[test]
fn clear_workbook_reports_only_provable_semantics() {
    let report = validate_source(source(authored(false, false))).expect("validation");

    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.cfb.ingress"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.workbook.stream"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.worksheet.inventory"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.protection.presence"),
        CheckStatus::NotApplicable
    ));
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::NotApplicable
    ));
    assert!(matches!(
        status(&report, "xls.signature.presence"),
        CheckStatus::NotApplicable
    ));
    assert!(matches!(
        status(&report, "xls.drm.presence"),
        CheckStatus::NotApplicable
    ));
    assert!(matches!(
        status(&report, "xls.external_reference.presence"),
        CheckStatus::NotApplicable
    ));
}

#[test]
fn external_metadata_is_reported_without_resolving_targets() {
    let report = validate_source(source(authored_external())).expect("validation");

    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.external_reference.presence"),
        CheckStatus::Complete
    ));
    assert!(report.issues().iter().any(|issue| {
        issue.code() == "xls.external_reference.present"
            && issue
                .message()
                .contains("targets were not resolved or fetched")
    }));
}

#[test]
fn all_space_external_path_remains_external_presence() {
    let payload = [1, 0, 2, 0, 0, b' ', b' ', 1, 0, 0, b' '];
    let stream = minimal_workbook(
        &[frame(0x01AE, &payload)],
        &[],
        &bof_payload(0x0010, 16),
        &[],
        &[],
    );
    let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
        .expect("validation report");

    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.external_reference.presence"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.external_reference.present")
    );
}

#[test]
fn rejects_duplicate_external_cache_owners_and_cells() {
    let duplicate_xct = [
        frame(0x01AE, &external_supbook_payload()),
        frame(0x0059, &[1, 0, 0, 0]),
        frame(0x005A, &external_cache_row(0, 0, 0)),
        frame(0x0059, &[1, 0, 0, 0]),
        frame(0x005A, &external_cache_row(0, 0, 1)),
    ];
    let overlapping_cells = [
        frame(0x01AE, &external_supbook_payload()),
        frame(0x0059, &[2, 0, 0, 0]),
        frame(0x005A, &external_cache_row(0, 0, 0)),
        frame(0x005A, &external_cache_row(0, 0, 0)),
    ];

    for records in [&duplicate_xct[..], &overlapping_cells[..]] {
        let stream = minimal_workbook(records, &[], &bof_payload(0x0010, 16), &[], &[]);
        let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
            .expect("validation report");
        assert!(report.has_errors(), "duplicate external cache was accepted");
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| { issue.code() == "xls.external_reference.invalid" })
        );
    }
}

#[test]
fn rejects_conflicting_dde_standard_document_flags() {
    let records = [
        frame(0x01AE, &dde_supbook_payload()),
        frame(0x0023, &conflicting_dde_name_payload()),
    ];
    let stream = minimal_workbook(&records, &[], &bof_payload(0x0010, 16), &[], &[]);
    let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
        .expect("validation report");

    assert!(report.has_errors());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| { issue.code() == "xls.external_reference.invalid" })
    );
}

#[test]
fn protection_is_reported_without_retaining_password_material() {
    let report = validate_source(source(authored(true, false))).expect("validation");

    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.protection.presence"),
        CheckStatus::Complete
    ));
    assert!(report.issues().iter().all(|issue| {
        !issue.message().contains("secret") && !issue.message().contains("Sheet1")
    }));
}

#[test]
fn encrypted_workbook_stops_clear_semantic_checks() {
    let report = validate_source(source(authored(false, true))).expect("validation");

    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.worksheet.inventory"),
        CheckStatus::StoppedBy { .. }
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.password_to_open_present")
    );
}

#[test]
fn writer_places_filepass_after_optional_writeprotect() {
    let report =
        validate_source(source(authored_encrypted_with_write_reservation())).expect("validation");

    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::Complete
    ));
    assert!(
        !report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.filepass_placement")
    );
}

#[test]
fn unsupported_filepass_is_not_reported_as_malformed() {
    let mut stream = frame(0x0809, &bof_payload(0x0005, 16));
    stream.extend_from_slice(&frame(0x002F, &unsupported_cryptoapi_filepass()));
    stream.extend_from_slice(&[0xFF; 5]);
    let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
        .expect("validation report");

    assert!(report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.unsupported")
    );
    assert!(
        !report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.filepass_invalid")
    );
}

#[test]
fn preserves_metadata_errors_before_filepass_stop() {
    let mut malformed_protection = frame(0x0809, &bof_payload(0x0005, 16));
    malformed_protection.extend_from_slice(&frame(0x0012, &[1, 0]));
    malformed_protection.extend_from_slice(&frame(0x0012, &[1, 0]));
    malformed_protection.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
    malformed_protection.extend_from_slice(&[0xFF; 5]);

    let mut malformed_external = frame(0x0809, &bof_payload(0x0005, 16));
    malformed_external.extend_from_slice(&frame(0x01AE, &external_supbook_payload()));
    malformed_external.extend_from_slice(&frame(0x0023, &[]));
    malformed_external.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
    malformed_external.extend_from_slice(&[0xFF; 5]);

    for (stream, check_id, issue_code) in [
        (
            malformed_protection,
            "xls.protection.presence",
            "xls.protection.invalid",
        ),
        (
            malformed_external,
            "xls.external_reference.presence",
            "xls.external_reference.invalid",
        ),
    ] {
        let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
            .expect("validation report");

        assert!(report.has_errors());
        assert!(matches!(
            status(&report, check_id),
            CheckStatus::Blocked { .. }
        ));
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == issue_code)
        );
    }
}

#[test]
fn malformed_cfb_is_structured_and_downstream_is_blocked() {
    let report = validate_source(source(vec![0u8; 512])).expect("validation report");

    assert!(!report.is_complete());
    assert!(report.has_errors());
    assert_eq!(report.issues().len(), 1);
    assert_eq!(report.issues()[0].code(), "xls.cfb.not_ole");
    assert!(matches!(
        status(&report, "xls.cfb.ingress"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.workbook.stream"),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn cfb_directory_byte_ceiling_is_blocked_without_validation_errors() {
    let workbook = minimal_workbook(&[], &[], &bof_payload(0x0010, 16), &[], &[]);
    let bytes = cfb_with_streams(
        &[
            ("Workbook", workbook),
            ("One", Vec::new()),
            ("Two", Vec::new()),
            ("Three", Vec::new()),
            ("Four", Vec::new()),
        ],
        &[],
    );

    let exact = XlsValidationLimits::default()
        .with_max_directory_bytes(1024)
        .expect("two directory sectors fit the exact ceiling");
    let report = validate_source_with_limits(source(bytes.clone()), exact)
        .expect("exact directory ceiling should validate");
    assert!(report.is_complete());
    assert!(!report.has_errors());

    let limited = XlsValidationLimits::default()
        .with_max_directory_bytes(512)
        .expect("one directory sector is a valid configured ceiling");
    let report = validate_source_with_limits(source(bytes), limited)
        .expect("directory ceiling should be represented in the report");
    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.cfb.ingress"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.workbook.stream"),
        CheckStatus::StoppedBy { .. }
    ));
    assert!(report.issues().is_empty());
}

#[test]
fn xls_directory_byte_limit_builder_rejects_invalid_values() {
    assert!(matches!(
        XlsValidationLimits::default().with_max_directory_bytes(0),
        Err(XlsValidationError::Ingress(OleError::InvalidLimit {
            resource: "CFB directory bytes",
            value: 0,
            ..
        }))
    ));
    assert!(matches!(
        XlsValidationLimits::default()
            .with_max_directory_bytes(litchi_cfb::SharedOleFileLimits::MAX_DIRECTORY_BYTES + 1),
        Err(XlsValidationError::Ingress(OleError::InvalidLimit {
            resource: "CFB directory bytes",
            ..
        }))
    ));
}

#[test]
fn workbook_byte_ceiling_blocks_before_materialization() {
    let limits = XlsValidationLimits::new(
        2 * 1024 * 1024 * 1024,
        1,
        1_000_000,
        65_535,
        65_536,
        100_000,
        litchi_core::ValidationLimits::default(),
    );
    let report = validate_source_with_limits(source(authored(false, false)), limits)
        .expect("validation report");

    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.workbook.stream"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::StoppedBy { .. }
    ));
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));
}

#[test]
fn streaming_record_limits_match_biff_boundaries() {
    let stream = minimal_workbook(
        &[],
        &[],
        &bof_payload(0x0010, 16),
        &[frame(0x7777, &[])],
        &[],
    );
    let limits = |max_biff_records| {
        XlsValidationLimits::new(
            2 * 1024 * 1024 * 1024,
            128 * 1024 * 1024,
            max_biff_records,
            65_535,
            65_536,
            100_000,
            litchi_core::ValidationLimits::default(),
        )
    };

    let preterminal = validate_source_with_limits(
        source(cfb_with_streams(&[("Workbook", stream.clone())], &[])),
        limits(1),
    )
    .expect("validation report");
    assert!(!preterminal.has_errors());
    assert!(matches!(
        status(&preterminal, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));
    assert!(matches!(
        status(&preterminal, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));

    for maximum in [6, 7] {
        let report = validate_source_with_limits(
            source(cfb_with_streams(&[("Workbook", stream.clone())], &[])),
            limits(maximum),
        )
        .expect("validation report");
        assert!(report.is_complete());
        assert!(!report.has_errors());
    }

    let limited = validate_source_with_limits(
        source(cfb_with_streams(&[("Workbook", stream.clone())], &[])),
        limits(5),
    )
    .expect("validation report");
    assert!(!limited.has_errors());
    assert!(matches!(
        status(&limited, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&limited, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));

    let zero = validate_source_with_limits(
        source(cfb_with_streams(&[("Workbook", stream)], &[])),
        limits(0),
    )
    .expect("validation report");
    assert!(!zero.has_errors());
    assert!(matches!(
        status(&zero, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn streaming_record_payload_and_truncation_boundaries_match_biff() {
    let maximum_payload = vec![0xA5; litchi_biff::MAX_RECORD_BYTES];
    let accepted = minimal_workbook(
        &[frame(0x7777, &maximum_payload)],
        &[],
        &bof_payload(0x0010, 16),
        &[],
        &[],
    );
    let report = validate_source(source(cfb_with_streams(&[("Workbook", accepted)], &[])))
        .expect("validation report");
    assert!(report.is_complete());
    assert!(!report.has_errors());

    let oversized_payload = vec![0xA5; litchi_biff::MAX_RECORD_BYTES + 1];
    let oversized = minimal_workbook(
        &[frame(0x7777, &oversized_payload)],
        &[],
        &bof_payload(0x0010, 16),
        &[],
        &[],
    );
    let report = validate_source(source(cfb_with_streams(&[("Workbook", oversized)], &[])))
        .expect("validation report");
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));

    let mut oversized_filepass = frame(0x0809, &bof_payload(0x0005, 16));
    oversized_filepass.extend_from_slice(&0x002Fu16.to_le_bytes());
    oversized_filepass.extend_from_slice(
        &u16::try_from(litchi_biff::MAX_RECORD_BYTES + 1)
            .expect("oversized BIFF payload")
            .to_le_bytes(),
    );
    let report = validate_source(source(cfb_with_streams(
        &[("Workbook", oversized_filepass)],
        &[],
    )))
    .expect("validation report");
    assert!(report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.filepass_invalid")
    );

    let mut truncated_filepass = frame(0x0809, &bof_payload(0x0005, 16));
    truncated_filepass.extend_from_slice(&0x002Fu16.to_le_bytes());
    truncated_filepass.extend_from_slice(&6u16.to_le_bytes());
    truncated_filepass.extend_from_slice(&[0; 2]);
    let report = validate_source(source(cfb_with_streams(
        &[("Workbook", truncated_filepass)],
        &[],
    )))
    .expect("validation report");
    assert!(report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.filepass_invalid")
    );

    let valid = minimal_workbook(&[], &[], &bof_payload(0x0010, 16), &[], &[]);
    let mut oversized_after_terminal = valid.clone();
    oversized_after_terminal.extend_from_slice(&0x7777u16.to_le_bytes());
    oversized_after_terminal.extend_from_slice(
        &u16::try_from(litchi_biff::MAX_RECORD_BYTES + 1)
            .expect("oversized BIFF payload")
            .to_le_bytes(),
    );
    let report = validate_source(source(cfb_with_streams(
        &[("Workbook", oversized_after_terminal)],
        &[],
    )))
    .expect("validation report");
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));

    for tail_len in 1..=3 {
        let mut stream = valid.clone();
        stream.extend(std::iter::repeat_n(0xAA, tail_len));
        let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
            .expect("validation report");
        assert!(report.has_errors());
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == "xls.biff.invalid")
        );
        assert!(matches!(
            status(&report, "xls.biff.parse"),
            CheckStatus::Blocked { .. }
        ));
        assert!(matches!(
            status(&report, "xls.encryption.presence"),
            CheckStatus::StoppedBy { .. }
        ));
    }

    let mut truncated = valid;
    truncated.extend_from_slice(&0x7777u16.to_le_bytes());
    truncated.extend_from_slice(&10u16.to_le_bytes());
    truncated.extend_from_slice(&[0xAA; 3]);
    let report = validate_source(source(cfb_with_streams(&[("Workbook", truncated)], &[])))
        .expect("validation report");
    assert!(report.has_errors());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.biff.invalid")
    );
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));

    let mut early_truncated = frame(0x0809, &bof_payload(0x0005, 16));
    early_truncated.extend_from_slice(&[0x77, 0x77, 0x0A]);
    let report = validate_source(source(cfb_with_streams(
        &[("Workbook", early_truncated)],
        &[],
    )))
    .expect("validation report");
    assert!(report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));
}

#[test]
fn streaming_validation_stops_before_large_ciphertext_tail() {
    let mut stream = frame(0x0809, &bof_payload(0x0005, 16));
    stream.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
    let sentinel: Vec<u8> = (0_usize..64)
        .map(|index| (index.wrapping_mul(37).wrapping_add(11) & 0xFF) as u8)
        .collect();
    stream.extend_from_slice(&sentinel);
    stream.extend(std::iter::repeat_n(0xFF, 64 * 1024));
    let cfb = cfb_with_streams(&[("Workbook", stream)], &[]);
    let mut matches = cfb
        .windows(sentinel.len())
        .enumerate()
        .filter_map(|item| (item.1 == sentinel.as_slice()).then_some(item.0));
    let tail_start = matches.next().expect("ciphertext sentinel");
    assert!(
        matches.next().is_none(),
        "ciphertext sentinel is not unique"
    );
    let tail_range = (tail_start as u64, sentinel.len());
    let source = Arc::new(ReadWidthSource::new(cfb, 16 * 1024));

    let report = validate_source(source.clone()).expect("validation report");
    assert!(!report.has_errors());
    assert!(source.maximum_observed() <= 16 * 1024);
    assert!(
        !source
            .ranges()
            .into_iter()
            .any(|range| ranges_overlap(range, tail_range)),
        "streaming validation touched the ciphertext tail"
    );
}

#[test]
fn streaming_validation_propagates_read_and_source_change_errors() {
    let marker: Vec<u8> = (0_u8..64)
        .map(|index| index.wrapping_mul(37).wrapping_add(11))
        .collect();
    let stream = minimal_workbook(
        &[frame(0x7777, &marker)],
        &[],
        &bof_payload(0x0010, 16),
        &[],
        &[],
    );
    let cfb = cfb_with_streams(&[("Workbook", stream)], &[]);
    let mut matches = cfb
        .windows(marker.len())
        .enumerate()
        .filter_map(|item| (item.1 == marker.as_slice()).then_some(item.0));
    let marker_start = matches.next().expect("streaming marker");
    assert!(matches.next().is_none(), "streaming marker is not unique");
    let marker_range = (marker_start as u64, marker.len());

    let source =
        Arc::new(ReadWidthSource::new(cfb.clone(), 16 * 1024).with_failure_range(marker_range));
    let error = validate_source(source).expect_err("injected read failure");
    assert!(matches!(
        error,
        XlsValidationError::Ingress(OleError::Io(_))
    ));

    let source =
        Arc::new(ReadWidthSource::new(cfb.clone(), 16 * 1024).with_change_range(marker_range));
    let error = validate_source(source).expect_err("injected source change");
    assert!(matches!(
        error,
        XlsValidationError::Ingress(OleError::SourceChanged { .. })
    ));

    let source = Arc::new(
        ReadWidthSource::new(cfb, 16 * 1024)
            .with_failure_range(marker_range)
            .with_change_range(marker_range),
    );
    let error = validate_source(source).expect_err("source change precedence");
    assert!(matches!(
        error,
        XlsValidationError::Ingress(OleError::SourceChanged { .. })
    ));
}

#[test]
fn incomplete_streaming_scans_fail_closed_for_encryption() {
    let limits = |max_workbook_stream_bytes, max_biff_records| {
        XlsValidationLimits::new(
            2 * 1024 * 1024 * 1024,
            max_workbook_stream_bytes,
            max_biff_records,
            65_535,
            65_536,
            100_000,
            litchi_core::ValidationLimits::default(),
        )
    };

    let mut record_limited = frame(0x0809, &bof_payload(0x0005, 16));
    record_limited.extend_from_slice(&frame(0x7777, &[]));
    let report = validate_source_with_limits(
        source(cfb_with_streams(&[("Workbook", record_limited)], &[])),
        limits(128 * 1024 * 1024, 1),
    )
    .expect("validation report");
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));

    let mut payload_limited = frame(0x0809, &bof_payload(0x0005, 16));
    payload_limited.extend_from_slice(&0x7777u16.to_le_bytes());
    payload_limited.extend_from_slice(
        &u16::try_from(litchi_biff::MAX_RECORD_BYTES + 1)
            .expect("oversized BIFF payload")
            .to_le_bytes(),
    );
    let report = validate_source_with_limits(
        source(cfb_with_streams(&[("Workbook", payload_limited)], &[])),
        limits(128 * 1024 * 1024, 1_000_000),
    )
    .expect("validation report");
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));

    let mut header_truncated = frame(0x0809, &bof_payload(0x0005, 16));
    header_truncated.extend_from_slice(&[0x77, 0x77, 0x0A]);
    let report = validate_source_with_limits(
        source(cfb_with_streams(&[("Workbook", header_truncated)], &[])),
        limits(128 * 1024 * 1024, 1_000_000),
    )
    .expect("validation report");
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));

    let stream = frame(0x0809, &bof_payload(0x0005, 16));
    let report = validate_source_with_limits(
        source(cfb_with_streams(&[("Workbook", stream)], &[])),
        limits(1, 1_000_000),
    )
    .expect("validation report");
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::StoppedBy { .. }
    ));
}

#[test]
fn finite_input_and_record_ceilings_stop_without_errors() {
    let input_limits = XlsValidationLimits::new(
        0,
        128 * 1024 * 1024,
        1_000_000,
        65_535,
        65_536,
        100_000,
        litchi_core::ValidationLimits::default(),
    );
    let report = validate_source_with_limits(source(authored(false, false)), input_limits)
        .expect("validation report");
    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.cfb.ingress"),
        CheckStatus::Blocked { .. }
    ));

    let record_limits = XlsValidationLimits::new(
        2 * 1024 * 1024 * 1024,
        128 * 1024 * 1024,
        1,
        65_535,
        65_536,
        100_000,
        litchi_core::ValidationLimits::default(),
    );
    let report = validate_source_with_limits(source(authored(false, false)), record_limits)
        .expect("validation report");
    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.worksheet.inventory"),
        CheckStatus::StoppedBy { .. }
    ));

    let worksheet_limits = XlsValidationLimits::new(
        2 * 1024 * 1024 * 1024,
        128 * 1024 * 1024,
        1_000_000,
        0,
        65_536,
        100_000,
        litchi_core::ValidationLimits::default(),
    );
    let report = validate_source_with_limits(source(authored(false, false)), worksheet_limits)
        .expect("validation report");
    assert!(!report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.worksheet.inventory"),
        CheckStatus::StoppedBy { .. }
    ));
}

#[test]
fn rejects_ambiguous_and_non_stream_workbook_entries() {
    let workbook = workbook_stream(authored(false, false));

    let book_fallback = cfb_with_streams(&[("Book", workbook.clone())], &[]);
    let report = validate_source(source(book_fallback)).expect("validation report");
    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.workbook.stream"),
        CheckStatus::Complete
    ));

    let both_names = cfb_with_streams(
        &[("Workbook", workbook.clone()), ("Book", workbook.clone())],
        &[],
    );
    let report = validate_source(source(both_names)).expect("validation report");
    assert_eq!(report.issues()[0].code(), "xls.workbook.stream_ambiguous");
    assert!(matches!(
        status(&report, "xls.workbook.stream"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));

    let non_stream = cfb_with_streams(&[("Book", workbook)], &["Workbook"]);
    let report = validate_source(source(non_stream)).expect("validation report");
    assert_eq!(report.issues()[0].code(), "xls.workbook.stream_ambiguous");
    assert!(report.has_errors());
}

#[test]
fn rejects_hostile_biff_ownership_and_global_ordering() {
    let cases = [
        (frame(0x0809, &[0; 8]), false),
        (frame(0x0809, &bof_payload(0x0010, 16)), false),
        (
            minimal_workbook(&[], &[], &bof_payload(0x0010, 8), &[], &[]),
            false,
        ),
        (
            minimal_workbook(&[], &[], &bof_payload(0x0020, 16), &[], &[]),
            false,
        ),
        (
            {
                let mut worksheet_bof = bof_payload(0x0010, 8);
                worksheet_bof[0..2].copy_from_slice(&0x0500u16.to_le_bytes());
                minimal_workbook(&[], &[], &worksheet_bof, &[], &[])
            },
            false,
        ),
        (
            minimal_workbook(
                &[],
                &[],
                &bof_payload(0x0010, 16),
                &[frame(0x000A, &[1])],
                &[],
            ),
            false,
        ),
        (
            minimal_workbook(
                &[],
                &[],
                &bof_payload(0x0010, 16),
                &[],
                &[frame(0x7777, &[0xAA])],
            ),
            false,
        ),
        (
            minimal_workbook(
                &[],
                &[],
                &bof_payload(0x0010, 16),
                &[],
                &[frame(0x0809, &bof_payload(0x0010, 16))],
            ),
            false,
        ),
        (
            minimal_workbook(
                &[],
                &[],
                &bof_payload(0x0010, 16),
                &[],
                &[frame(0x000A, &[])],
            ),
            false,
        ),
        (
            minimal_workbook(
                &[frame(0x0809, &bof_payload(0x0005, 16))],
                &[],
                &bof_payload(0x0010, 16),
                &[],
                &[],
            ),
            false,
        ),
        (
            minimal_workbook(
                &[frame(0x0042, &0x04B0u16.to_le_bytes())],
                &[frame(0x0042, &0x04B0u16.to_le_bytes())],
                &bof_payload(0x0010, 16),
                &[],
                &[],
            ),
            false,
        ),
        (
            {
                let mut bytes = frame(0x0809, &bof_payload(0x0005, 16));
                bytes.extend_from_slice(&[0x77, 0x77, 0x0A]);
                bytes
            },
            true,
        ),
        (
            {
                let mut bytes = frame(0x0809, &bof_payload(0x0005, 16));
                bytes.extend_from_slice(&0x7777u16.to_le_bytes());
                bytes.extend_from_slice(&10u16.to_le_bytes());
                bytes.extend_from_slice(&[0xAA; 3]);
                bytes
            },
            true,
        ),
    ];

    for (stream, expect_blocked) in cases {
        let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
            .expect("validation report");
        assert!(report.has_errors(), "hostile BIFF was accepted");
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == "xls.biff.invalid"
                    || issue.code() == "xls.worksheet.invalid")
        );
        if expect_blocked {
            assert!(matches!(
                status(&report, "xls.biff.parse"),
                CheckStatus::Blocked { .. }
            ));
        } else {
            assert!(matches!(
                status(&report, "xls.biff.parse"),
                CheckStatus::Complete
            ));
        }
    }
}

#[test]
fn stops_before_ciphertext_after_bounded_filepass() {
    let mut stream = frame(0x0809, &bof_payload(0x0005, 16));
    stream.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
    stream.extend_from_slice(&[0xFF, 0xFF, 0xFF, 0xFF, 0xFF]);

    let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
        .expect("validation report");
    assert!(!report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "xls.worksheet.inventory"),
        CheckStatus::StoppedBy { .. }
    ));
    assert!(
        !report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.biff.invalid")
    );
}

#[test]
fn reports_filepass_placement_errors_before_stopping_ciphertext() {
    let mut first = frame(0x002F, &[0, 0, 0, 0, 0, 0]);
    first.extend_from_slice(&[0xFF; 5]);

    let mut after_global_eof = frame(0x0809, &bof_payload(0x0005, 16));
    after_global_eof.extend_from_slice(&frame(0x000A, &[]));
    after_global_eof.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
    after_global_eof.extend_from_slice(&[0xFF; 5]);

    let mut inside_sheet = minimal_workbook(
        &[],
        &[],
        &bof_payload(0x0010, 16),
        &[frame(0x002F, &[0, 0, 0, 0, 0, 0]), frame(0x000A, &[])],
        &[],
    );
    inside_sheet.extend_from_slice(&[0xFF; 5]);

    for stream in [first, after_global_eof, inside_sheet] {
        let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
            .expect("validation report");
        assert!(
            report.has_errors(),
            "invalid FILEPASS placement was accepted"
        );
        assert!(matches!(
            status(&report, "xls.encryption.presence"),
            CheckStatus::Complete
        ));
        assert!(matches!(
            status(&report, "xls.biff.parse"),
            CheckStatus::Blocked { .. }
        ));
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| { issue.code() == "xls.encryption.filepass_placement" })
        );
    }
}

#[test]
fn rejects_filepass_after_codepage_before_ciphertext() {
    let mut stream = frame(0x0809, &bof_payload(0x0005, 16));
    stream.extend_from_slice(&frame(0x0042, &0x04B0u16.to_le_bytes()));
    stream.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
    stream.extend_from_slice(&[0xFF; 5]);

    let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
        .expect("validation report");

    assert!(report.has_errors());
    assert!(matches!(
        status(&report, "xls.encryption.presence"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "xls.biff.parse"),
        CheckStatus::Blocked { .. }
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.filepass_placement")
    );
}

#[test]
fn accepts_filepass_only_after_early_writeprotect() {
    let mut valid = frame(0x0809, &bof_payload(0x0005, 16));
    valid.extend_from_slice(&frame(0x0086, &[]));
    valid.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
    valid.extend_from_slice(&[0xFF; 5]);
    let report = validate_source(source(cfb_with_streams(&[("Workbook", valid)], &[])))
        .expect("validation report");
    assert!(!report.has_errors());
    assert!(
        !report
            .issues()
            .iter()
            .any(|issue| issue.code() == "xls.encryption.filepass_placement")
    );
}

#[test]
fn rejects_filepass_after_other_global_records() {
    let prelude_records = [
        vec![frame(0x01AE, &external_supbook_payload())],
        vec![frame(0x005B, &[0; 6])],
    ];

    for records in prelude_records {
        let mut stream = frame(0x0809, &bof_payload(0x0005, 16));
        for record in records {
            stream.extend_from_slice(&record);
        }
        stream.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));
        stream.extend_from_slice(&[0xFF; 5]);
        let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
            .expect("validation report");
        assert!(report.has_errors());
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == "xls.encryption.filepass_placement")
        );
    }
}

#[test]
fn preserves_incomplete_metadata_before_filepass_stop() {
    let mut incomplete_workbook_protection = frame(0x0809, &bof_payload(0x0005, 16));
    incomplete_workbook_protection.extend_from_slice(&frame(0x01AF, &[1, 0]));
    incomplete_workbook_protection.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));

    let mut incomplete_external_cache = frame(0x0809, &bof_payload(0x0005, 16));
    incomplete_external_cache.extend_from_slice(&frame(0x01AE, &external_supbook_payload()));
    incomplete_external_cache.extend_from_slice(&frame(0x0059, &[1, 0, 0, 0]));
    incomplete_external_cache.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));

    let mut incomplete_dde_values = frame(0x0809, &bof_payload(0x0005, 16));
    incomplete_dde_values.extend_from_slice(&frame(0x01AE, &dde_supbook_payload()));
    incomplete_dde_values.extend_from_slice(&frame(0x0023, &dde_pending_name_payload()));
    incomplete_dde_values.extend_from_slice(&frame(0x002F, &[0, 0, 0, 0, 0, 0]));

    let mut incomplete_sheet_protection = minimal_workbook(
        &[],
        &[],
        &bof_payload(0x0010, 16),
        &[frame(0x0063, &[1, 0]), frame(0x002F, &[0, 0, 0, 0, 0, 0])],
        &[],
    );
    incomplete_sheet_protection.extend_from_slice(&[0xFF; 5]);

    for (stream, check, expected_issue) in [
        (
            incomplete_workbook_protection,
            "xls.protection.presence",
            "xls.protection.invalid",
        ),
        (
            incomplete_external_cache,
            "xls.external_reference.presence",
            "xls.external_reference.invalid",
        ),
        (
            incomplete_dde_values,
            "xls.external_reference.presence",
            "xls.external_reference.invalid",
        ),
        (
            incomplete_sheet_protection,
            "xls.protection.presence",
            "xls.protection.invalid",
        ),
    ] {
        let report = validate_source(source(cfb_with_streams(&[("Workbook", stream)], &[])))
            .expect("validation report");
        assert!(report.has_errors());
        assert!(matches!(
            status(&report, check),
            CheckStatus::Blocked { .. }
        ));
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == expected_issue)
        );
    }
}
