use litchi_core::Position;
use litchi_ods::{
    Builder, MutableSpreadsheet, Spreadsheet,
    definitions::{Key, Snapshot},
    names::{Definition, Expression, Range, Scope},
};

fn range(name: &str, column: &str) -> litchi_core::Result<Definition> {
    Range::new(name, format!("$Sheet1.${column}$1"), Scope::Global).map(Into::into)
}

#[test]
fn definition_patch_round_trips_inverse_and_exact_source() -> litchi_core::Result<()> {
    let mut builder = Builder::new();
    builder.add_definition(range("First", "A")?)?;
    let source = builder.build()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;

    let mut edit = snapshot.edit();
    edit.add(Expression::new("Second", "of:=1", Scope::Global)?.into())?;
    edit.move_to(Key::new("Second", &Scope::Global), Position::new(0))?;
    let commit = edit.commit()?;

    assert!(commit.changed());
    assert_eq!(
        commit
            .snapshot()
            .definitions()
            .iter()
            .map(Definition::name)
            .collect::<Vec<_>>(),
        ["Second", "First"]
    );
    let spreadsheet = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    litchi_odf_common::compact_xml::validate(spreadsheet.content_xml().as_bytes())?;

    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), source);
    assert!(commit.patch().apply(&snapshot).is_ok());

    let unrelated = Snapshot::from_bytes(Builder::new().build()?)?;
    assert!(commit.patch().apply(&unrelated).is_err());
    Ok(())
}

#[test]
fn failed_definition_staging_does_not_mutate_the_draft() -> litchi_core::Result<()> {
    let source = Builder::new().build()?;
    let snapshot = Snapshot::from_bytes(source)?;
    let mut edit = snapshot.edit();
    edit.add(range("Total", "A")?)?;
    let before = edit.definitions().to_vec();

    assert!(edit.add(range("Total", "B")?).is_err());
    assert_eq!(edit.definitions(), before);
    assert!(edit.replace(9_usize, range("Missing", "C")?)?.is_none());
    assert!(edit.remove(9_usize)?.is_none());
    Ok(())
}

#[test]
fn mutable_facade_publishes_one_transaction_and_replays_its_patch() -> litchi_core::Result<()> {
    let source = Builder::new().build()?;
    let source_snapshot = Snapshot::from_bytes(source.clone())?;
    let mut staged = source_snapshot.edit();
    staged.add(range("Input", "A")?)?;
    let staged_commit = staged.commit()?;
    let patch = staged_commit.patch().clone();

    let mut mutable = MutableSpreadsheet::from_bytes(source.clone())?;
    mutable.apply_definitions_patch(&patch)?;
    assert_eq!(mutable.definitions()[0].name(), "Input");

    let mut direct = MutableSpreadsheet::from_bytes(source)?;
    direct.edit_definitions(|edit| edit.add(range("Input", "A")?))?;
    assert_eq!(direct.definitions(), mutable.definitions());
    Ok(())
}
