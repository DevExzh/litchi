//! Native inputs for bounded append, transfer and fresh plaintext derivation.
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource,
};
use litchi_rtf::{
    Package, TableCellPath,
    tail_append::{SemanticLimits, TailSelector},
};
use std::{
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
    sync::Arc,
};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn context() -> ExecutionContext {
    let budget = Budget::root(
        "native-rtf-workflow",
        CoreLimits::new(64 << 20, 64 << 20, 64 << 20, 1_000_000, 256, 64 << 20),
    );
    let (_, token) = CancellationSource::pair();
    ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(4096).unwrap(),
            0,
        )
        .unwrap(),
    )
}
const APPEND: [&str; 2] = [
    "NATIVE_APPEND_035: café 中文 😀",
    "Escapes: {left} \\right\tend",
];
fn check_splice(before: &[u8], after: &[u8]) {
    let close = before.iter().rposition(|b| *b == b'}').unwrap();
    let added = after.len() - before.len();
    assert!(after[..close] == before[..close]);
    assert!(after[close + added..] == before[close..]);
}
fn save(p: &Package) -> Vec<u8> {
    let mut file = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
    p.write_to(&mut file).unwrap();
    std::fs::read(file.path()).unwrap()
}
#[test]
fn native_logical_tail_append_preserves_source_bytes_and_replays_durable_inverse() {
    for (name, count) in [
        ("word-rtf-minimal.rtf", 1),
        ("word-bidi-eastasia.rtf", 7),
        ("word-rtf-table-grid.rtf", 2),
        ("word-tables-layout.rtf", 10),
        ("word-write-reservation.rtf", 5),
        ("word-tracked-moves.rtf", 6),
        ("word-layout-extensions.rtf", 9),
    ] {
        let original = native(name);
        let p = Package::from_bytes(&original).unwrap();
        assert_eq!(p.paragraph_count(), count);
        let mut edit = p.edit_tail_append(TailSelector::Body);
        edit.append_text_paragraphs(&APPEND).unwrap();
        let commit = edit.commit().unwrap();
        let output = save(commit.package());
        check_splice(&original, &output);
        let read = Package::from_bytes(&output).unwrap();
        assert_eq!(read.paragraph_count(), count + 2);
        assert_eq!(
            read.text(),
            format!(
                "{}{}{}\n{}\n",
                p.text(),
                if p.text().ends_with('\n') { "" } else { "\n" },
                APPEND[0],
                APPEND[1]
            )
        );
        assert_eq!(read.raw().data_store(), p.raw().data_store());
        assert_eq!(
            read.raw().write_reservations(),
            p.raw().write_reservations()
        );
        assert_eq!(read.tables().len(), p.tables().len());
        let limits = SemanticLimits::default();
        let wire = commit
            .patch()
            .to_durable(limits)
            .unwrap()
            .to_deterministic_json()
            .unwrap();
        let durable =
            litchi_rtf::tail_append::DurableTailAppendPatch::from_deterministic_json(&wire, limits)
                .unwrap();
        assert_eq!(durable.to_deterministic_json().unwrap(), wire);
        assert!(durable.apply(&p).unwrap().to_bytes().unwrap() == output);
        assert!(durable.inverse().apply(&read).unwrap().to_bytes().unwrap() == original);
        assert!(p.to_bytes().unwrap() == original);
    }
}
fn source_backed(bytes: &[u8]) -> litchi_rtf::source_backed::Package {
    let p = Package::from_bytes(bytes).unwrap();
    litchi_rtf::source_backed::Package::open_with_options(
        &p,
        Arc::new(OwnedSource::new(bytes.to_vec())),
        context(),
        litchi_rtf::source_backed::OpenOptions::new().publication_limits(
            litchi_rtf::tail_append::TailAppendPublicationLimits::new(257, 17),
        ),
    )
    .unwrap()
}
#[derive(Default)]
struct SmallSink {
    bytes: Vec<u8>,
    largest: usize,
}
impl Write for SmallSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.largest = self.largest.max(bytes.len());
        let n = bytes.len().min(7);
        self.bytes.extend_from_slice(&bytes[..n]);
        Ok(n)
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
#[test]
fn native_source_backed_tail_append_respects_write_windows_and_reopened_inverse() {
    for name in ["word-bidi-eastasia.rtf", "word-rtf-table-grid.rtf"] {
        let original = native(name);
        let source = source_backed(&original);
        let mut edit = source.edit_tail_append().unwrap();
        edit.append_text_paragraphs(&APPEND).unwrap();
        let commit = edit.commit().unwrap();
        let mut sink = SmallSink::default();
        commit.package().write_to(&mut sink).unwrap();
        assert!(sink.largest <= 17);
        check_splice(&original, &sink.bytes);
        let p = Package::from_bytes(&original).unwrap();
        let mut logical = p.edit_tail_append(TailSelector::Body);
        logical.append_text_paragraphs(&APPEND).unwrap();
        assert!(logical.commit().unwrap().package().to_bytes().unwrap() == sink.bytes);
        let file = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
        commit.package().save(file.path()).unwrap();
        assert!(std::fs::read(file.path()).unwrap() == sink.bytes);
        let reopened = source_backed(&sink.bytes);
        let inverse = reopened
            .apply_tail_append(&commit.patch().inverse())
            .unwrap();
        let mut bytes = Vec::new();
        inverse.package().write_to(&mut bytes).unwrap();
        assert!(bytes == original);
        let mut retained = Vec::new();
        source.write_to(&mut retained).unwrap();
        assert!(retained == original);
    }
}
#[test]
fn native_tail_append_refusals_preserve_protected_and_dependent_sources() {
    use litchi_core::edit::RefusalKind::{Protected, UnsupportedContent};
    for (name, kind, reason) in [
        (
            "word-legacy-form.rtf",
            Protected,
            "document protection refuses changed publication",
        ),
        (
            "word-document-protection.rtf",
            Protected,
            "document protection refuses changed publication",
        ),
        (
            "word-review-controls.rtf",
            UnsupportedContent,
            "dropped unknown syntax prevents a safe root-tail splice",
        ),
        (
            "word-visual-objects.rtf",
            UnsupportedContent,
            "dropped unknown syntax prevents a safe root-tail splice",
        ),
        (
            "word-mailmerge-fields.rtf",
            UnsupportedContent,
            "external document or transformation content is present",
        ),
    ] {
        let original = native(name);
        let p = Package::from_bytes(&original).unwrap();
        assert!(
            p.edit_tail_append(TailSelector::Body)
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == original
        );
        let mut edit = p.edit_tail_append(TailSelector::Body);
        edit.append_text_paragraphs(&APPEND).unwrap();
        assert!(
            matches!(edit.commit(),Err(litchi_rtf::tail_append::Error::Refused(r)) if r.kind()==kind && r.reason()==reason),
            "{name}"
        );
        assert!(p.to_bytes().unwrap() == original);
    }
}
#[test]
fn native_table_cell_transfer_changes_only_the_recipient_literal() {
    let source_bytes = native("word-tables-layout.rtf");
    let target_bytes = native("word-rtf-table-grid.rtf");
    let source = Package::from_bytes(&source_bytes).unwrap();
    let target = Package::from_bytes(&target_bytes).unwrap();
    assert_eq!(source.tables()[0].rows()[1].cells()[0].text(), "Apples");
    assert_eq!(
        target.tables()[0].rows()[1].cells()[0].text(),
        "RTF_CELL_006"
    );
    let plan = litchi_rtf::edit::TransferPlan::table_cell_text(
        &source,
        &TableCellPath::outer(0, 1, 0),
        &target,
        TableCellPath::outer(0, 1, 0),
    )
    .unwrap();
    let commit = plan.commit().unwrap();
    let after = save(commit.package());
    let text = std::str::from_utf8(&target_bytes).unwrap();
    assert_eq!(text.matches("RTF_CELL_006").count(), 1);
    assert!(after == text.replace("RTF_CELL_006", "Apples").into_bytes());
    let read = Package::from_bytes(&after).unwrap();
    assert_eq!(read.tables()[0].rows()[1].cells()[0].text(), "Apples");
    assert_eq!(read.tables()[0].rows()[1].cells()[1].text(), "42");
    assert!(
        read.apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == target_bytes
    );
    assert!(
        source.to_bytes().unwrap() == source_bytes && target.to_bytes().unwrap() == target_bytes
    );
}
#[test]
fn native_multilingual_text_can_feed_a_bounded_fresh_stream_in_single_byte_chunks() {
    // This intentionally authors a plaintext derivative. It does not copy
    // native run formatting, layout, fonts or producer metadata.
    let original = native("word-bidi-eastasia.rtf");
    let p = Package::from_bytes(&original).unwrap();
    let paragraphs = p
        .body()
        .paragraphs()
        .map(|p| p.to_text())
        .collect::<Vec<_>>();
    let expected = [
        "WORD_BIDI_EASTASIA_015",
        "Fixture ID: word-bidi-eastasia-015",
        "ARABIC_015: مرحبا بالعالم — اختبار النص العربي ١٢٣٤٥",
        "HEBREW_015: שלום עולם — בדיקת טקסט 12345",
        "JAPANESE_015: 日本語の禁則処理、句読点。「試験」１２３４５",
        "CHINESE_015: 中文排版测试，标点符号。《测试》一二三四五",
        "END_BIDI_EASTASIA_015",
    ];
    assert_eq!(paragraphs, expected);
    let input = paragraphs.iter().map(String::len).sum::<usize>();
    let mut writer = litchi_rtf::streaming::Writer::new(
        Vec::new(),
        context(),
        litchi_rtf::streaming::Limits::new(input as u64, 1 << 20, 7, 7, 64),
    )
    .unwrap();
    for paragraph in &paragraphs {
        writer.start_paragraph().unwrap();
        writer.start_run().unwrap();
        for byte in paragraph.as_bytes() {
            writer.write_all(&[*byte]).unwrap();
        }
        writer.finish_run().unwrap();
        writer.finish_paragraph().unwrap();
    }
    let output = writer.finish().unwrap();
    let read = Package::from_bytes(&output).unwrap();
    assert_eq!(
        read.body()
            .paragraphs()
            .map(|p| p.to_text())
            .collect::<Vec<_>>(),
        expected
    );
    assert!(read.fields().is_empty() && read.tables().is_empty() && read.objects().is_empty());
    assert!(read.validation_report().is_conservatively_clean());
    assert!(p.to_bytes().unwrap() == original);
}
