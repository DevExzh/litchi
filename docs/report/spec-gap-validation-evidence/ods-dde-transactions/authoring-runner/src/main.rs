//! Emit one content.xml candidate through the public inert DDE transaction API.
//!
//! The output contains a deliberately nonexistent DDE target and disabled
//! automatic updates.  It never opens a package, starts a process, or resolves
//! the external topic.  The resulting XML is intended for owner-level RNG
//! validation with the sibling `validate_schema.py` script.

use std::{
    env, fs,
    num::{NonZeroU64, NonZeroUsize},
    path::PathBuf,
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits, Profile,
};
use litchi_ods::dde::{
    AutomaticUpdate, CachedCell, CachedRow, CachedTable, CachedValue, ConversionMode, LinkSpec,
    Snapshot, Source,
};

const SEED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

fn context() -> ExecutionContext {
    let (_cancellation_source, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker limit"),
        NonZeroUsize::new(1).expect("task limit"),
        NonZeroU64::new(8 * 1024 * 1024).expect("byte limit"),
        0,
    )
    .expect("finite execution limits");
    let budget = Budget::root(
        "ods-dde-authoring-capture",
        BudgetLimits::for_profile(Profile::TrustedBatch),
    );
    ExecutionContext::new(budget, token, limits)
}

fn authored_link() -> LinkSpec {
    let source = Source::new(
        "never\tcontacted\nDDE\rapp",
        "file:///never\tcontacted\nDDE\rtopic.ods",
        "Data.A1:B2",
    )
    .expect("inert source")
    .named("Captured\tname\nwith\rcontrols")
    .expect("source name")
    .with_conversion_mode(ConversionMode::KeepText)
    .with_automatic_update(AutomaticUpdate::Disabled);
    let row = CachedRow::from_cells(vec![
        CachedCell::new(CachedValue::Number(7.0)).expect("number cell"),
        CachedCell::new(CachedValue::Text(
            "captured\ttext\nwith\rcontrols".to_owned(),
        ))
        .expect("text cell"),
    ])
    .expect("cache row");
    let cache = CachedTable::from_rows(vec![row]).expect("cache table");
    LinkSpec::new(source, cache).expect("complete link")
}

fn empty_link() -> LinkSpec {
    let source = Source::new(
        "never-contacted-dde",
        "file:///never-contacted-dde.ods",
        "Data.C3",
    )
    .expect("inert empty-cache source")
    .named("EmptyCapture")
    .expect("empty-cache source name")
    .with_conversion_mode(ConversionMode::KeepText)
    .with_automatic_update(AutomaticUpdate::Disabled);
    LinkSpec::new(source, CachedTable::new()).expect("empty cache link")
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("api-authored-content.xml"));
    let snapshot = Snapshot::parse(SEED)?;
    let mut edit = snapshot.edit();
    edit.add_link(authored_link())?;
    edit.add_link(empty_link())?;
    let commit = edit.commit(&context())?;
    fs::write(&output, commit.snapshot().source_xml())?;
    println!(
        "wrote {} bytes to {} (changed={}, links={})",
        commit.snapshot().source_xml().len(),
        output.display(),
        commit.changed(),
        commit.snapshot().links().len(),
    );
    Ok(())
}
