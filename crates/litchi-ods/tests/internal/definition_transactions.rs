use crate::{
    Builder,
    definitions::Snapshot,
    names::{Definition, Range, Scope},
};

fn range(name: &str, column: &str) -> litchi_core::Result<Definition> {
    Range::new(name, format!("$Sheet1.${column}$1"), Scope::Global).map(Into::into)
}

#[test]
fn failed_definition_staging_does_not_mutate_the_draft() -> litchi_core::Result<()> {
    let source = Builder::new().build()?.to_bytes()?;
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
