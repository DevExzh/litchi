//! Native-input statistics, validation and forward-only plaintext authoring.
use litchi_core::{
    Budget, CancellationSource, CheckStatus, EvidenceValue, ExecutionContext, ExecutionLimits,
    Limits as CoreLimits, OwnedSource, Resource, ValidateReport, ValidationLimits,
};
use litchi_docx::validation::{Limits as ValidationPolicy, validate_read_at};
use litchi_docx::{
    Package, StreamingDocumentError, StreamingDocumentLimits, StreamingDocumentWriter,
};
use std::{
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
    sync::Arc,
};
const TEXT: &str = "word-text-structure.docx";
const BIDI: &str = "word-bidi-eastasia.docx";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn package(name: &str) -> Package {
    Package::from_bytes(&native(name)).unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = reader.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}

fn assert_members_equal(left: &[u8], right: &[u8]) {
    let a = members(left);
    let b = members(right);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        assert!(b[&name] == bytes, "changed {name}");
    }
}

fn status<'a>(r: &'a ValidateReport, id: &str) -> &'a CheckStatus {
    r.checks()
        .iter()
        .find(|c| c.id().as_str() == id)
        .unwrap()
        .status()
}
fn evidence(r: &ValidateReport, code: &str, key: &str) -> Option<u64> {
    r.issues()
        .iter()
        .find(|i| i.code() == code)
        .and_then(|i| i.evidence().iter().find(|e| e.key() == key))
        .and_then(|e| {
            if let EvidenceValue::Count(v) = e.value() {
                Some(v)
            } else {
                None
            }
        })
}
fn validate(bytes: &[u8], limits: ValidationPolicy) -> ValidateReport {
    validate_read_at(
        Arc::new(OwnedSource::new(bytes.to_vec())),
        Default::default(),
        limits,
        ValidationLimits::default(),
    )
    .unwrap()
}
#[test]
fn native_statistics_count_words_across_paragraph_and_table_cell_boundaries() {
    for (name, words, chars, without, paragraphs, tables, lines) in [
        (TEXT, 51, 334, 295, 12, 0, 5),
        ("word-tables-layout.docx", 57, 390, 365, 36, 1, 5),
        (BIDI, 24, 239, 222, 7, 0, 3),
    ] {
        let p = package(name);
        let s = p.document().unwrap().statistics().unwrap();
        assert_eq!(
            (
                s.word_count(),
                s.character_count(),
                s.character_count_no_spaces()
            ),
            (words, chars, without),
            "{name}"
        );
        assert_eq!((s.paragraph_count(), s.table_count()), (paragraphs, tables));
        assert_eq!((s.line_count(), s.page_count()), (lines, 1));
        assert_eq!((s.image_count(), s.drawing_count()), (0, 0));
        assert_members_equal(&p.to_bytes().unwrap(), &native(name));
    }
    let p = package("word-visual-objects.docx");
    let s = p.document().unwrap().statistics().unwrap();
    assert_eq!(
        (
            s.paragraph_count(),
            s.table_count(),
            s.image_count(),
            s.drawing_count()
        ),
        (9, 0, 2, 4)
    );
    // Drawing and image inventories overlap; they are not disjoint physical object counts.
}
#[test]
fn native_validation_reports_inert_macro_external_and_embedded_presence_without_false_svg_errors() {
    for (name, blocks, external, macros, embedded, unsupported) in [
        (TEXT, 12, Some(1), None, None, None),
        ("word-tables-layout.docx", 13, None, None, None, None),
        (BIDI, 7, None, None, None, None),
        ("word-visual-objects.docx", 9, None, None, Some(2), None),
        ("word-vector-diagram.docx", 7, None, None, None, None),
        ("word-recorded-macro.docm", 5, None, Some(5), None, None),
        ("word-recorded-macro.dotm", 5, None, Some(5), None, None),
        ("word-math-navigation.dotx", 17, None, None, None, Some(1)),
        ("word-document-protection.docx", 5, None, None, None, None),
        ("word-layout-extensions.docx", 11, None, None, None, None),
    ] {
        let bytes = native(name);
        let r = validate(&bytes, ValidationPolicy::default());
        assert!(r.is_complete(), "{name}: {r:?}");
        assert!(!r.has_errors(), "{name}: {r:?}");
        assert!(!r.has_fatal());
        assert_eq!(r.checks().len(), 9);
        for id in [
            "docx.package.ingress",
            "docx.package.main_document",
            "docx.main_document.relationship_closure",
            "docx.main_document.markup_compatibility",
            "docx.main_document.semantics",
        ] {
            assert_eq!(status(&r, id), &CheckStatus::Complete);
        }
        for (capability, code, expected) in [
            (
                "docx.relationships.external_target_presence",
                "docx.external_target.present",
                external,
            ),
            (
                "docx.security.macro_presence",
                "docx.macro.storage_present",
                macros,
            ),
            (
                "docx.security.embedded_content_presence",
                "docx.embedded_content.present",
                embedded,
            ),
        ] {
            assert_eq!(
                status(&r, capability),
                if expected.is_some() {
                    &CheckStatus::Complete
                } else {
                    &CheckStatus::NotApplicable
                }
            );
            assert_eq!(evidence(&r, code, "observations"), expected);
        }
        assert_eq!(
            status(&r, "docx.security.signature_presence"),
            &CheckStatus::NotApplicable
        );
        assert_eq!(
            evidence(
                &r,
                "docx.main_document.supported_blocks",
                "supported_paragraphs_and_tables"
            ),
            Some(blocks)
        );
        assert_eq!(
            evidence(
                &r,
                "docx.main_document.unsupported_body_children",
                "unsupported_body_children"
            ),
            unsupported
        );
        assert_eq!(
            evidence(
                &r,
                "docx.main_document.relationship_target_content_type",
                "content_type_mismatches"
            ),
            None
        );
        let diagnostic = format!("{r:?}");
        for text in ["Ryker Zhu", "example.com", "WORD_NATIVE_END", "ARABIC_015"] {
            assert!(!diagnostic.contains(text));
        }
        assert_eq!(native(name), bytes);
    }
}
#[test]
fn native_validation_exact_main_payload_bound_and_one_under_report_distinct_states() {
    let bytes = native("word-vector-diagram.docx");
    let length = members(&bytes)["word/document.xml"].len() as u64;
    let exact = validate(
        &bytes,
        ValidationPolicy::default().with_max_main_document_bytes(length),
    );
    assert!(exact.is_complete());
    assert!(!exact.has_errors());
    let short = validate(
        &bytes,
        ValidationPolicy::default().with_max_main_document_bytes(length - 1),
    );
    assert!(!short.is_complete());
    assert!(matches!(
        status(&short, "docx.main_document.markup_compatibility"),
        CheckStatus::Blocked { .. }
    ));
    assert!(
        matches!(status(&short,"docx.main_document.semantics"),CheckStatus::StoppedBy{check} if check.as_str()=="docx.main_document.markup_compatibility")
    );
    assert_eq!(native("word-vector-diagram.docx"), bytes);
}
fn context() -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "native-docx-stream",
        CoreLimits::new(1024 * 1024, 1024 * 1024, 1024 * 1024, 10000, 64, 10000000),
    );
    let (cancel, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(1024 * 1024).unwrap(),
        1,
    )
    .unwrap();
    let execution = ExecutionContext::new(budget.clone(), token, limits);
    (budget, cancel, execution)
}
#[derive(Default)]
struct ShortSink {
    bytes: Vec<u8>,
}
impl Write for ShortSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let n = bytes.len().min(3);
        self.bytes.extend_from_slice(&bytes[..n]);
        Ok(n)
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
const BIDI_PARAGRAPHS: [&str; 7] = [
    "WORD_BIDI_EASTASIA_015",
    "Fixture ID: word-bidi-eastasia-015",
    "ARABIC_015: مرحبا بالعالم — اختبار النص العربي ١٢٣٤٥",
    "HEBREW_015: שלום עולם — בדיקת טקסט 12345",
    "JAPANESE_015: 日本語の禁則処理、句読点。「試験」１２３４５",
    "CHINESE_015: 中文排版测试，标点符号。《测试》一二三四五",
    "END_BIDI_EASTASIA_015",
];
#[test]
fn native_unicode_paragraphs_stream_through_short_writes_with_exact_text_and_counts() {
    for name in [TEXT, BIDI] {
        let p = package(name);
        let original = p.to_bytes().unwrap();
        let paragraphs = p
            .document()
            .unwrap()
            .paragraphs()
            .unwrap()
            .iter()
            .map(|p| p.text().unwrap())
            .collect::<Vec<_>>();
        if name == BIDI {
            assert_eq!(paragraphs, BIDI_PARAGRAPHS);
        } else {
            assert_eq!(paragraphs.len(), 12);
            assert_eq!(
                paragraphs[2],
                "Unicode: café naïve résumé — Ελληνικά — 中文 — 日本語 — العربية — 😀"
            );
        }
        let input = paragraphs.iter().map(String::len).sum::<usize>() as u64;
        let (budget, _cancel, execution) = context();
        let limits = StreamingDocumentLimits {
            max_input_bytes: input,
            max_paragraphs: paragraphs.len() as u64,
            max_runs: paragraphs.len() as u64,
            ..Default::default()
        };
        let mut writer =
            StreamingDocumentWriter::new(ShortSink::default(), execution, limits).unwrap();
        for text in &paragraphs {
            writer.start_paragraph().unwrap();
            writer.start_run().unwrap();
            for c in text.chars() {
                writer.write_text(&c.to_string()).unwrap();
            }
            writer.finish_run().unwrap();
            writer.finish_paragraph().unwrap();
        }
        assert_eq!(writer.input_bytes(), input);
        let output = writer.finish().unwrap().bytes;
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(
            members(&output)
                .keys()
                .map(String::as_str)
                .collect::<Vec<_>>(),
            ["[Content_Types].xml", "_rels/.rels", "word/document.xml"]
        );
        let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
        std::fs::write(f.path(), &output).unwrap();
        let reopened = Package::open(f.path()).unwrap();
        let d = reopened.document().unwrap();
        assert_eq!(
            d.paragraphs()
                .unwrap()
                .iter()
                .map(|p| p.text().unwrap())
                .collect::<Vec<_>>(),
            paragraphs
        );
        let s = d.statistics().unwrap();
        assert_eq!(s.word_count(), if name == TEXT { 51 } else { 24 });
        assert_eq!(s.character_count(), if name == TEXT { 334 } else { 239 });
        assert!(validate(&output, ValidationPolicy::default()).is_complete());
        assert_members_equal(&p.to_bytes().unwrap(), &original);
        assert_members_equal(&original, &native(name));
    }
}
#[test]
fn native_stream_input_limits_and_cancellation_preserve_failure_progress() {
    let p = package(BIDI);
    let text = p.document().unwrap().paragraphs().unwrap()[2]
        .text()
        .unwrap();
    assert_eq!(text, BIDI_PARAGRAPHS[2]);
    let n = text.len() as u64;
    let (budget, _cancel, execution) = context();
    let mut writer = StreamingDocumentWriter::new(
        Vec::new(),
        execution,
        StreamingDocumentLimits {
            max_input_bytes: n - 1,
            ..Default::default()
        },
    )
    .unwrap();
    writer.start_paragraph().unwrap();
    writer.start_run().unwrap();
    let before = writer.output_bytes();
    let error = writer.write_text(&text).unwrap_err();
    assert!(
        matches!(error,StreamingDocumentError::LimitExceeded{resource:"input bytes",observed,limit,written} if observed==n && limit==n-1 && written==before)
    );
    assert_eq!(writer.input_bytes(), 0);
    assert_eq!(writer.output_bytes(), before);
    assert!(writer.is_poisoned());
    assert_eq!(writer.finish().unwrap_err().written(), before);
    assert_eq!(budget.used(Resource::Memory), 0);
    let (budget, cancel, execution) = context();
    let mut writer =
        StreamingDocumentWriter::new(Vec::new(), execution, Default::default()).unwrap();
    writer.start_paragraph().unwrap();
    writer.start_run().unwrap();
    writer.write_text(&text).unwrap();
    writer.finish_run().unwrap();
    writer.finish_paragraph().unwrap();
    let accepted = writer.output_bytes();
    cancel.cancel();
    assert!(
        matches!(writer.start_paragraph(),Err(StreamingDocumentError::Cancelled{written}) if written==accepted)
    );
    assert!(writer.is_poisoned());
    assert_eq!(writer.finish().unwrap_err().written(), accepted);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_members_equal(&p.to_bytes().unwrap(), &native(BIDI));
}
