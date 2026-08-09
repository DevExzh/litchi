mod support;

use std::{fs, path::Path};

use litchi_core::{HistoryLimits, MergeChoice, Result};
use litchi_ods::{
    Builder, Cell, CellValue, Sheet, Spreadsheet,
    annotations::Annotation,
    document::{Collision, History, JoinFailure, Limits, Resource, Snapshot, TransferDisposition},
    names::{Definition, Expression, Scope},
    rdf::{Object, Subject, Triple},
};

fn source() -> Result<Vec<u8>> {
    let mut builder = Builder::new();
    builder.add_sheet(Sheet::new("Data")?)?;
    builder.set_cell(
        "Data",
        0,
        0,
        Cell::new(CellValue::Text("seed".to_string()), "seed"),
    )?;
    builder.build()
}

fn label(value: &str) -> Triple {
    Triple {
        subject: Subject::Iri("urn:litchi:test:sheet".to_string()),
        predicate: "https://example.invalid/schema#label".to_string(),
        object: Object::Literal {
            value: value.to_string(),
            datatype: None,
            language: None,
        },
    }
}

#[test]
fn one_commit_composes_semantic_owners_and_resource_crud() -> Result<()> {
    let source = source()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.worksheets(|worksheets| {
        let _changed = worksheets.set_formula("Data", 0, 0, "of:=1+1")?;
        let _changed = worksheets.set_cell_style("Data", 0, 0, "InputCell")?;
        Ok(())
    })?;
    edit.definitions(|definitions| {
        definitions.add(Expression::new("Input", "of:=1+1", Scope::Global)?.into())
    })?;
    edit.annotations(|annotations| annotations.add("Data", 0, 0, Annotation::new("reviewed")))?;
    edit.rdf(|rdf| {
        let _path = rdf.add_graph(None, &[label("Data")])?;
        Ok(())
    })?;
    edit.protection(|protection| {
        protection.document_mut().structure_protected = Some(false);
        Ok(())
    })?;
    edit.data_pilot(|_data_pilot| Ok(()))?;
    edit.tracked_changes(|_tracked_changes| Ok(()))?;
    edit.charts(|_charts| Ok(()))?;
    let resource = Resource::new(
        "Pictures/transaction.bin",
        "application/octet-stream",
        vec![1, 3, 3, 7],
    )?;
    let _disposition = edit.put_resource(resource, Collision::Reject)?;

    let commit = edit.commit()?;
    assert!(commit.changed());
    assert!(commit.patch().operations().len() >= 5);
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    let cell = reopened.cell("Data", 0, 0);
    assert!(matches!(
        cell,
        Some(litchi_ods::CellView::Stored(value))
            if value.formula.as_deref() == Some("of:=1+1")
                && value.style_name.as_deref() == Some("InputCell")
    ));
    assert!(matches!(
        reopened.definitions(),
        [Definition::Expression(value)] if value.name == "Input"
    ));
    assert!(reopened.annotations()?.cell("Data", 0, 0)?.is_some());
    assert_eq!(reopened.rdf_graphs()?.len(), 1);
    let retained = commit
        .snapshot()
        .resource("Pictures/transaction.bin")?
        .ok_or_else(|| litchi_core::Error::InvalidFormat("resource is missing".to_string()))?;
    assert_eq!(retained.as_bytes(), [1, 3, 3, 7]);

    let wire = commit.patch().to_deterministic_json()?;
    assert_eq!(wire, commit.patch().to_deterministic_json()?);
    let decoded = litchi_ods::document::Patch::from_deterministic_json(&wire, snapshot.limits())?;
    assert_eq!(
        decoded.apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        decoded
            .inverse()
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );
    let inverse_wire = decoded.inverse().to_deterministic_json()?;
    let decoded_inverse =
        litchi_ods::document::Patch::from_deterministic_json(&inverse_wire, snapshot.limits())?;
    assert_eq!(
        decoded_inverse
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );
    let stale = Snapshot::from_bytes(Builder::new().build()?)?;
    assert!(decoded.apply(&stale).is_err());
    Ok(())
}

fn formula_patch(snapshot: &Snapshot, formula: &str) -> Result<litchi_ods::document::Patch> {
    let mut edit = snapshot.edit();
    edit.worksheets(|worksheets| {
        let _changed = worksheets.set_formula("Data", 0, 0, formula)?;
        Ok(())
    })?;
    Ok(edit.commit()?.patch().clone())
}

#[test]
fn join_is_deterministic_and_three_way_conflicts_are_explicit() -> Result<()> {
    let snapshot = Snapshot::from_bytes(source()?)?;
    let left = formula_patch(&snapshot, "of:=1")?;
    let mut resource_edit = snapshot.edit();
    let _disposition = resource_edit.put_resource(
        Resource::new("Data/sidecar.bin", "application/octet-stream", vec![9])?,
        Collision::Reject,
    )?;
    let right = resource_edit.commit()?.patch().clone();
    let left_right = left
        .join(&right)
        .map_err(|error| litchi_core::Error::InvalidFormat(error.detail().to_string()))?;
    let right_left = right
        .join(&left)
        .map_err(|error| litchi_core::Error::InvalidFormat(error.detail().to_string()))?;
    assert_eq!(
        left_right.to_deterministic_json()?,
        right_left.to_deterministic_json()?
    );
    let merged = left_right.apply(&snapshot)?;
    assert!(merged.snapshot().resource("Data/sidecar.bin")?.is_some());

    let competing = formula_patch(&snapshot, "of:=2")?;
    let conflict = left.join(&competing).err();
    assert!(matches!(
        conflict
            .as_ref()
            .map(litchi_ods::document::JoinError::failure),
        Some(JoinFailure::Conflict(_))
    ));
    let mut plan = litchi_ods::document::Patch::three_way(&left, &competing)?;
    assert!(
        plan.conflicts()
            .is_some_and(|conflicts| !conflicts.is_empty())
    );
    let _plan = plan.resolve(MergeChoice::Right);
    let resolved = plan.finish()?.apply(&snapshot)?;
    let reopened = Spreadsheet::from_bytes(resolved.snapshot().as_bytes().to_vec())?;
    assert!(matches!(
        reopened.cell("Data", 0, 0),
        Some(litchi_ods::CellView::Stored(cell)) if cell.formula.as_deref() == Some("of:=2")
    ));
    Ok(())
}

#[test]
fn bounded_history_evicts_oldest_transition_and_supports_redo() -> Result<()> {
    let defaults = Limits::default();
    let limits = Limits::new(
        defaults.max_package_bytes(),
        defaults.max_resources(),
        defaults.max_resource_bytes(),
        defaults.patch(),
        defaults.composition(),
        HistoryLimits::new(1, 512 * 1024 * 1024),
    );
    let base = Snapshot::from_bytes_with(source()?, limits)?;
    let mut history = History::new(base);

    let mut first_edit = history.current().edit();
    first_edit.worksheets(|worksheets| {
        let _changed = worksheets.set_formula("Data", 0, 0, "of:=1")?;
        Ok(())
    })?;
    let first = first_edit.commit()?;
    let first_bytes = first.snapshot().as_bytes().to_vec();
    assert!(history.record(first)?.is_empty());

    let mut second_edit = history.current().edit();
    second_edit.worksheets(|worksheets| {
        let _changed = worksheets.set_formula("Data", 0, 0, "of:=2")?;
        Ok(())
    })?;
    let second = second_edit.commit()?;
    assert_eq!(history.record(second)?.len(), 1);
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), first_bytes);
    assert!(!history.undo());
    assert!(history.redo());
    Ok(())
}

#[test]
fn resource_transfer_has_explicit_collisions_and_conservative_removal() -> Result<()> {
    let base = Snapshot::from_bytes(source()?)?;
    let mut source_edit = base.edit();
    assert_eq!(
        source_edit.put_resource(
            Resource::new(
                "Pictures/source.bin",
                "application/octet-stream",
                vec![5, 8]
            )?,
            Collision::Reject,
        )?,
        TransferDisposition::Added
    );
    let transfer_source = source_edit.commit()?.into_snapshot();

    let mut destination_edit = base.edit();
    assert_eq!(
        destination_edit.transfer_resource(
            &transfer_source,
            "Pictures/source.bin",
            "Pictures/copied.bin",
            Collision::Reject,
        )?,
        TransferDisposition::Added
    );
    let destination = destination_edit.commit()?.into_snapshot();
    let mut collision_edit = destination.edit();
    assert_eq!(
        collision_edit.transfer_resource(
            &transfer_source,
            "Pictures/source.bin",
            "Pictures/copied.bin",
            Collision::ReuseEquivalent,
        )?,
        TransferDisposition::Reused
    );
    assert!(
        collision_edit
            .transfer_resource(
                &transfer_source,
                "Pictures/source.bin",
                "Pictures/copied.bin",
                Collision::Reject,
            )
            .is_err()
    );
    collision_edit.remove_resource("Pictures/copied.bin")?;
    let removed = collision_edit.commit()?;
    assert!(
        removed
            .snapshot()
            .resource("Pictures/copied.bin")?
            .is_none()
    );
    Ok(())
}

#[test]
fn signed_and_protected_sources_refuse_changed_publication() -> Result<()> {
    let base = source()?;
    let content = Spreadsheet::from_bytes(base)?.content_xml().to_string();
    let signed_bytes = support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        (
            "META-INF/documentsignatures.xml",
            br#"<ds:document-signatures xmlns:ds="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
            "text/xml",
        ),
    ]);
    let signed = Snapshot::from_bytes(signed_bytes)?;
    assert!(!signed.edit().commit()?.changed());
    let mut changed_signed = signed.edit();
    let _disposition = changed_signed.put_resource(
        Resource::new("Data/new.bin", "application/octet-stream", vec![1])?,
        Collision::Reject,
    )?;
    assert!(changed_signed.commit().is_err());

    let unprotected = Snapshot::from_bytes(source()?)?;
    let mut protect = unprotected.edit();
    protect.protection(|metadata| {
        metadata.document_mut().structure_protected = Some(true);
        Ok(())
    })?;
    let protected = protect.commit()?.into_snapshot();
    let mut changed_protected = protected.edit();
    let _disposition = changed_protected.put_resource(
        Resource::new("Data/new.bin", "application/octet-stream", vec![1])?,
        Collision::Reject,
    )?;
    assert!(changed_protected.commit().is_err());
    Ok(())
}

#[test]
fn real_calc_fixture_reopens_after_resource_transfer_and_inverse() -> Result<()> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/odf/corpus/calc-two-sheets.ods");
    let source = fs::read(path)?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    let _disposition = edit.put_resource(
        Resource::new(
            "Data/roundtrip.bin",
            "application/octet-stream",
            vec![2, 4, 6, 8],
        )?,
        Collision::Reject,
    )?;
    let commit = edit.commit()?;
    assert_eq!(
        Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
            .sheets()
            .len(),
        2
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );
    Ok(())
}
