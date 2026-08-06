//! Task-pane, binding, and cross-graph edit tests.

use litchi_ooxml_common::web as common_web;

use super::super::*;
use super::support::task_panes;

#[test]
fn web_bindings_and_task_panes_commit_join_and_reverse_atomically() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");

    let mut pane_edit = source.edit().expect("pane edit");
    pane_edit
        .put_task_panes(
            task_panes("sheet-table"),
            common_web::Conformance::Transitional,
        )
        .expect("stage panes");
    let mut binding_edit = source.edit().expect("binding edit");
    binding_edit
        .sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!$A$1:$B$4").expect("binding"))
        .expect("stage binding");
    pane_edit.join(binding_edit).expect("dependent edits join");

    let committed = pane_edit.commit().expect("atomic web commit");
    assert_eq!(committed.patch().package_changes().len(), 1);
    assert!(committed.patch().changes().iter().any(|change| {
        matches!(change, Change::Web { sheet, .. } if sheet.as_ref() == "Sheet1")
    }));
    assert_eq!(
        committed
            .workbook()
            .task_panes()
            .expect("task panes")
            .expect("present")
            .len(),
        1
    );
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet");
    assert_eq!(
        sheet
            .web_bindings()
            .expect("bindings")
            .get("sheet-table")
            .map(crate::web::Binding::formula),
        Some("Sheet1!$A$1:$B$4")
    );

    let before_rename = committed.workbook().to_bytes().expect("pre-rename bytes");
    let mut rename = committed.workbook().edit().expect("rename edit");
    rename
        .tab("Sheet1")
        .expect("tab lookup")
        .expect("tab")
        .rename("Data")
        .expect("rename");
    let renamed = rename.commit().expect("dependency-aware rename");
    assert_eq!(
        renamed
            .workbook()
            .sheet("Data")
            .expect("renamed lookup")
            .expect("renamed worksheet")
            .web_bindings()
            .expect("renamed bindings")
            .get("sheet-table")
            .map(crate::web::Binding::formula),
        Some("Data!$A$1:$B$4")
    );
    assert_eq!(
        renamed
            .workbook()
            .apply(&renamed.patch().inverse())
            .expect("inverse rename")
            .workbook()
            .to_bytes()
            .expect("inverse rename bytes"),
        before_rename
    );

    let mut left = committed.workbook().edit().expect("left web edit");
    left.sheet("Sheet1")
        .expect("left lookup")
        .expect("left worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!A1").expect("left binding"))
        .expect("left stage");
    let mut right = committed.workbook().edit().expect("right web edit");
    right
        .sheet("Sheet1")
        .expect("right lookup")
        .expect("right worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!B2").expect("right binding"))
        .expect("right stage");
    let conflict = left.join(right).expect_err("whole binding sets overlap");
    assert!(
        conflict
            .conflicts()
            .expect("web conflict set")
            .conflicts()
            .iter()
            .any(Conflict::is_web)
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse web patch");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut unsafe_remove = committed.workbook().edit().expect("remove edit");
    unsafe_remove
        .remove_task_panes()
        .expect("stage unsafe removal");
    assert!(matches!(
        unsafe_remove.commit(),
        Err(Error::DanglingWebBinding { sheet, app_ref })
            if sheet == "Sheet1" && app_ref == "sheet-table"
    ));

    let mut safe_remove = committed.workbook().edit().expect("safe remove edit");
    safe_remove.remove_task_panes().expect("stage pane removal");
    assert!(
        safe_remove
            .sheet("Sheet1")
            .expect("lookup")
            .expect("worksheet")
            .clear_bindings()
            .expect("clear bindings")
    );
    let removed = safe_remove.commit().expect("combined removal");
    assert!(
        removed
            .workbook()
            .task_panes()
            .expect("task panes")
            .is_none()
    );
    assert!(
        removed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("worksheet")
            .web_bindings()
            .expect("bindings")
            .is_empty()
    );
}
#[test]
fn task_panes_and_sheet_removal_join_conflicts_are_symmetric() {
    let baseline = Workbook::new().expect("baseline");
    let mut create = baseline.edit().expect("create edit");
    create.add("Delete").expect("new worksheet");
    let source = create.commit().expect("source").into_workbook();

    for panes_first in [true, false] {
        let mut panes = source.edit().expect("pane edit");
        panes
            .put_task_panes(
                task_panes("sheet-table"),
                common_web::Conformance::Transitional,
            )
            .expect("stage panes");
        let mut removal = source.edit().expect("removal edit");
        removal
            .remove("Delete")
            .expect("lookup")
            .expect("worksheet");

        let error = if panes_first {
            panes.join(removal).expect_err("panes then removal")
        } else {
            removal.join(panes).expect_err("removal then panes")
        };
        assert!(matches!(
            error.failure(),
            JoinFailure::Overlap(conflicts)
                if conflicts.conflicts().iter().any(Conflict::is_remove)
        ));
    }
}

#[test]
fn web_patch_cross_graph_validation_is_atomic_on_alternate_snapshots() {
    fn with_bindings(workbook: &Workbook, bindings: WebBindings) -> Workbook {
        let mut package = workbook.inner.package.clone();
        let uri = workbook.inner.sheets[0].part_uri.clone();
        let source = package
            .get_part(&uri)
            .expect("worksheet part")
            .blob()
            .to_vec();
        let after = raw::web::replace(&source, &bindings).expect("write bindings");
        package
            .get_part_mut(&uri)
            .expect("worksheet part")
            .set_blob(after);
        Workbook::from_package(package).expect("reopen alternate workbook")
    }

    let baseline = Workbook::new().expect("baseline");
    let mut install = baseline.edit().expect("install edit");
    install
        .put_task_panes(
            task_panes("sheet-table"),
            common_web::Conformance::Transitional,
        )
        .expect("stage panes");
    let with_panes = install.commit().expect("install panes").into_workbook();

    let mut bind = with_panes.edit().expect("binding edit");
    bind.sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet")
        .bind(WebBinding::new("sheet-table", "Sheet1!A1").expect("binding"))
        .expect("stage binding");
    let binding_commit = bind.commit().expect("binding commit");
    let alternate = Workbook::new().expect("alternate without panes");
    let before = alternate.to_bytes().expect("alternate bytes");
    assert!(matches!(
        alternate.apply(binding_commit.patch()),
        Err(Error::DanglingWebBinding { sheet, app_ref })
            if sheet == "Sheet1" && app_ref == "sheet-table"
    ));
    assert_eq!(alternate.to_bytes().expect("unchanged bytes"), before);

    let old_bindings = WebBindings::try_from(vec![
        WebBinding::new("sheet-table", "Sheet1!A1").expect("old binding"),
    ])
    .expect("binding collection");
    let alternate = with_bindings(&with_panes, old_bindings);
    let before = alternate.to_bytes().expect("alternate bytes");
    let mut replace = with_panes.edit().expect("pane replacement");
    replace
        .put_task_panes(
            task_panes("replacement-table"),
            common_web::Conformance::Transitional,
        )
        .expect("stage replacement panes");
    let pane_commit = replace.commit().expect("pane replacement commit");
    assert!(matches!(
        alternate.apply(pane_commit.patch()),
        Err(Error::DanglingWebBinding { sheet, app_ref })
            if sheet == "Sheet1" && app_ref == "sheet-table"
    ));
    assert_eq!(alternate.to_bytes().expect("unchanged bytes"), before);
}
