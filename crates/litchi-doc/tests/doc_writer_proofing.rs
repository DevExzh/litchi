use litchi_doc::{
    Package, ProofingEntry, ProofingFeature, ProofingState, ProofingStateTable, ProofingStatus,
    ProofingTables, Writer,
};
use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};

static NEXT_FILE: AtomicUsize = AtomicUsize::new(0);

fn temporary_doc_path() -> std::path::PathBuf {
    std::env::temp_dir().join(format!(
        "litchi-proofing-{}-{}.doc",
        std::process::id(),
        NEXT_FILE.fetch_add(1, Ordering::Relaxed)
    ))
}

fn proofing_tables() -> ProofingTables {
    let spelling_clean = ProofingStatus::try_new(
        ProofingFeature::Spelling,
        ProofingState::Clean,
        false,
        false,
        false,
    )
    .unwrap();
    let unknown_word = ProofingStatus::try_new(
        ProofingFeature::Spelling,
        ProofingState::UnknownWord,
        true,
        false,
        false,
    )
    .unwrap();
    let grammar_clean = ProofingStatus::try_new(
        ProofingFeature::Grammar,
        ProofingState::Clean,
        false,
        false,
        false,
    )
    .unwrap();
    let grammar_error = ProofingStatus::try_new(
        ProofingFeature::Grammar,
        ProofingState::Dirty,
        true,
        true,
        true,
    )
    .unwrap();

    let spelling = ProofingStateTable::try_new(
        ProofingFeature::Spelling,
        vec![
            ProofingEntry::new(0, spelling_clean),
            ProofingEntry::new(6, unknown_word),
        ],
        11,
    )
    .unwrap();
    let grammar = ProofingStateTable::try_new(
        ProofingFeature::Grammar,
        vec![
            ProofingEntry::new(0, grammar_clean),
            ProofingEntry::new(6, grammar_error),
            ProofingEntry::new(6, grammar_clean),
        ],
        11,
    )
    .unwrap();
    ProofingTables::try_new(Some(spelling), Some(grammar)).unwrap()
}

fn assert_authored_tables(package: &mut Package<Cursor<Vec<u8>>>, expected: &ProofingTables) {
    let document = package.document().unwrap();
    assert_eq!(document.proofing_tables().unwrap(), expected);
    assert!(
        document
            .proofing_tables()
            .unwrap()
            .grammar()
            .unwrap()
            .range(1)
            .unwrap()
            .is_point()
    );
}

#[test]
fn spelling_and_grammar_tables_round_trip_through_write_to() {
    let expected = proofing_tables();
    let mut writer = Writer::new();
    writer.add_paragraph("alpha beta").unwrap();
    writer.set_proofing_tables(expected.clone());

    assert_eq!(
        writer
            .proofing_table(ProofingFeature::Spelling)
            .unwrap()
            .terminal_cp(),
        11
    );

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    assert_authored_tables(&mut package, &expected);
}

#[test]
fn proofing_tables_round_trip_through_file_save_and_can_be_cleared() {
    let expected = proofing_tables();
    let mut writer = Writer::new();
    writer.add_paragraph("alpha beta").unwrap();
    writer.set_proofing_tables(expected.clone());
    let removed = writer
        .clear_proofing_table(ProofingFeature::Grammar)
        .unwrap();
    assert_eq!(removed.feature(), ProofingFeature::Grammar);
    writer.set_proofing_table(removed);

    let path = temporary_doc_path();
    writer.save(&path).unwrap();
    let bytes = std::fs::read(&path).unwrap();
    std::fs::remove_file(path).unwrap();
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    assert_authored_tables(&mut package, &expected);
}

#[test]
fn output_rejects_out_of_document_proofing_cps_before_writing() {
    let status = ProofingStatus::try_new(
        ProofingFeature::Spelling,
        ProofingState::Clean,
        false,
        false,
        false,
    )
    .unwrap();
    let table = ProofingStateTable::try_new(
        ProofingFeature::Spelling,
        vec![ProofingEntry::new(0, status)],
        100,
    )
    .unwrap();

    let mut writer = Writer::new();
    writer.add_paragraph("x").unwrap();
    writer.set_proofing_table(table);

    let original = vec![0xA5; 16];
    let mut output = Cursor::new(original.clone());
    let error = writer.write_to(&mut output).unwrap_err();
    assert!(error.to_string().contains("proofing terminal CP"));
    assert_eq!(output.into_inner(), original);
}

#[test]
fn paired_constructor_rejects_feature_slot_mismatches() {
    let grammar = ProofingStateTable::try_new(ProofingFeature::Grammar, vec![], 0).unwrap();
    let spelling = ProofingStateTable::try_new(ProofingFeature::Spelling, vec![], 0).unwrap();

    assert!(ProofingTables::try_new(Some(grammar.clone()), None).is_err());
    assert!(ProofingTables::try_new(None, Some(spelling)).is_err());
}
