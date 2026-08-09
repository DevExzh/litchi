//! Cell, merge, style, and package-edit regression tests.

use super::super::codec::calculation_chain_removal;
use super::super::*;
use super::support::{merged_workbook, part_text, styled_workbook, two_sheet_workbook};

use crate::StyleState;
use crate::cell::{Number, Value};
use crate::error::MergeEditBlock;
use crate::formula::Formula;
use litchi_opc::{BlobPart, PackURI, TargetMode};

#[test]
fn cell_crud_is_atomic_reversible_and_source_preserving() {
    let source = Workbook::new().expect("source workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    {
        let mut sheet = edit.sheet("Sheet1").expect("sheet lookup").expect("sheet");
        sheet
            .set("A1", "hello")
            .and_then(|sheet| sheet.set("B2", 42_i32))
            .and_then(|sheet| sheet.set("C3", Formula::new("B2*2").expect("formula")))
            .expect("cell changes");
    }
    let committed = edit.commit().expect("commit");
    assert_eq!(
        source.to_bytes().expect("source remains valid"),
        source_bytes
    );
    assert_eq!(committed.patch().len(), 3);

    let book = committed.workbook();
    let sheet = book.sheet("Sheet1").expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("A1").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "hello"
    ));
    assert!(matches!(
        sheet.cell("B2").expect("B2").stored(),
        Some(Cell::Value(Value::Number(number))) if number == &Number::new("42").expect("number")
    ));
    assert!(matches!(
        sheet.cell("C3").expect("C3").stored(),
        Some(Cell::Formula(_))
    ));
    let extents = sheet.extents().expect("committed extents");
    assert_eq!(extents.declared().map(Rect::a1).as_deref(), Some("A1:C3"));
    assert_eq!(extents.content().map(Rect::a1).as_deref(), Some("A1:C3"));

    let restored = book
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    assert!(matches!(
        source.apply(committed.patch()),
        Ok(applied) if applied.workbook().sheet("Sheet1").expect("lookup").expect("sheet").cell("A1").expect("cell").stored().is_some()
    ));
    assert!(matches!(
        book.apply(committed.patch()),
        Err(Error::PatchConflict { .. })
    ));
}
#[test]
fn merged_range_crud_is_sparse_safe_reversible_and_composable() {
    let source = merged_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let source_sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
    assert_eq!(
        source_sheet
            .merges()
            .expect("merged ranges")
            .map(Rect::a1)
            .collect::<Vec<_>>(),
        ["A1:B2"]
    );
    assert!(matches!(
        source_sheet.cell("A1").expect("anchor"),
        crate::cell::View::Stored(Cell::Value(_))
    ));
    assert!(matches!(
        source_sheet.cell("B2").expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("A1:B2").expect("range")
    ));

    let mut edit = source.edit().expect("edit");
    edit.sheet("Sheet1")
        .expect("lookup")
        .expect("sheet")
        .unmerge("B2")
        .and_then(|sheet| sheet.set("B2", "revealed"))
        .and_then(|sheet| sheet.merge("C3:D4"))
        .expect("merged-range changes");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().len(), 3);
    let sheet = committed
        .workbook()
        .sheet("Sheet1")
        .expect("lookup")
        .expect("sheet");
    assert_eq!(
        sheet
            .merges()
            .expect("merged ranges")
            .map(Rect::a1)
            .collect::<Vec<_>>(),
        ["C3:D4"]
    );
    assert!(matches!(
        sheet.cell("B2").expect("uncovered cell").stored(),
        Some(Cell::Value(Value::Text(value))) if value.as_str() == "revealed"
    ));
    assert!(matches!(
        sheet.cell("D4").expect("covered cell"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("C3:D4").expect("range")
    ));
    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let mut blocked = source.edit().expect("blocked edit");
    blocked
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D1:E2")
        .expect("plan merge");
    assert!(matches!(
        blocked.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::FollowerContent { address },
            ..
        }) if address == Address::from_a1("E2").expect("address")
    ));

    let mut cleared = source.edit().expect("clear edit");
    cleared
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .clear("E2")
        .and_then(|sheet| sheet.merge("D1:E2"))
        .expect("clear then merge");
    let cleared = cleared.commit().expect("safe merge");
    assert!(matches!(
        cleared
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("E2")
            .expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("D1:E2").expect("range")
    ));

    let mut merge = source.edit().expect("independent merge");
    merge
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D1:E2")
        .expect("merge");
    let mut clear = source.edit().expect("independent clear");
    clear
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .clear("E2")
        .expect("clear");
    merge.join(clear).expect("clear makes the merge safe");
    assert!(matches!(
        merge
            .commit()
            .expect("joined clear and merge")
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("E2")
            .expect("covered"),
        crate::cell::View::Covered(_)
    ));

    let mut merge = source.edit().expect("independent merge");
    merge
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("merge");
    let mut write = source.edit().expect("independent follower write");
    write
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("D4", "hidden")
        .expect("write");
    assert!(matches!(
        merge.join(write).expect_err("follower content must conflict").failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| conflict.merges().is_some())
    ));

    let mut overlap = source.edit().expect("overlap edit");
    overlap
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("B2:C3")
        .expect("plan overlap");
    assert!(matches!(
        overlap.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::Overlap { existing },
            ..
        }) if existing == Rect::from_a1("A1:B2").expect("existing")
    ));

    let mut created = source.edit().expect("create edit");
    created
        .add("Merged")
        .expect("new sheet")
        .set("A1", "anchor")
        .and_then(|sheet| sheet.merge("A1:C2"))
        .expect("new merged range");
    let created = created.commit().expect("create merged sheet");
    assert!(matches!(
        created
            .workbook()
            .sheet("Merged")
            .expect("lookup")
            .expect("sheet")
            .cell("C2")
            .expect("covered"),
        crate::cell::View::Covered(range) if range == Rect::from_a1("A1:C2").expect("range")
    ));

    let mut left = source.edit().expect("left edit");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("left merge");
    let mut right = source.edit().expect("right edit");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("F1:G2")
        .expect("right merge");
    left.join(right).expect("disjoint merge join");
    assert_eq!(
        left.commit()
            .expect("joined commit")
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .merges()
            .expect("merges")
            .count(),
        3
    );

    let mut left = source.edit().expect("left overlap");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("C3:D4")
        .expect("left merge");
    let mut right = source.edit().expect("right overlap");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .merge("D4:E5")
        .expect("right merge");
    let error = left.join(right).expect_err("overlapping joins must fail");
    assert!(matches!(
        error.failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| conflict.merges().is_some())
    ));

    let mut left = source.edit().expect("left unmerge");
    left.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("A1")
        .expect("left unmerge");
    let mut right = source.edit().expect("right unmerge");
    right
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("B2")
        .expect("right unmerge");
    let error = left
        .join(right)
        .expect_err("two selectors for one merge must conflict");
    assert!(matches!(
        error.failure(),
        JoinFailure::Overlap(conflicts)
            if conflicts.conflicts().iter().any(|conflict| {
                conflict.merges().is_some_and(|ranges| {
                    ranges == [Rect::from_a1("A1:B2").expect("range")]
                })
            })
    ));

    let mut missing = source.edit().expect("missing unmerge");
    missing
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .unmerge("H8")
        .expect("missing range is a no-op");
    assert!(missing.commit().expect("no-op commit").patch().is_empty());
}

#[test]
fn merged_range_reads_share_one_snapshot_without_public_locks() {
    let workbook = Arc::new(merged_workbook());
    std::thread::scope(|scope| {
        for _ in 0..8 {
            let workbook = Arc::clone(&workbook);
            scope.spawn(move || {
                let sheet = workbook.sheet(0usize).expect("lookup").expect("sheet");
                assert!(matches!(
                    sheet.cell("B2").expect("covered"),
                    crate::cell::View::Covered(range)
                        if range == Rect::from_a1("A1:B2").expect("range")
                ));
                assert_eq!(sheet.merges().expect("merges").count(), 1);
            });
        }
    });
}

#[test]
fn clear_and_remove_have_distinct_missing_and_empty_semantics() {
    let source = Workbook::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", "value")
        .expect("set");
    let first = edit.commit().expect("first commit").into_workbook();

    let mut edit = first.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet.clear("A1").expect("clear");
    sheet.clear("B1").expect("clear missing");
    let cleared = edit.commit().expect("clear commit");
    assert_eq!(cleared.patch().len(), 1);
    let sheet = cleared
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("cell").stored(),
        Some(Cell::Empty)
    ));
    assert!(sheet.cell("B1").expect("cell").is_missing());

    let mut edit = cleared.workbook().edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .remove("A1")
        .expect("remove");
    let removed = edit.commit().expect("remove commit");
    assert!(
        removed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("A1")
            .expect("cell")
            .is_missing()
    );
}

#[test]
fn clearing_a_cell_created_in_the_same_transaction_keeps_an_empty_record() {
    let source = Workbook::new().expect("source workbook");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet.set("A1", "temporary").expect("set");
    sheet.clear("A1").expect("clear");
    let committed = edit.commit().expect("commit");
    assert!(matches!(
        committed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("A1")
            .expect("cell")
            .stored(),
        Some(Cell::Empty)
    ));
}

#[test]
fn calculation_chain_removal_is_atomic_and_reversible() {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    let chain_uri = PackURI::new("/xl/calcChain.xml").expect("chain URI");
    package
            .try_add_part(Box::new(BlobPart::new(
                chain_uri.clone(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                    .to_owned(),
                br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>"#.to_vec(),
            )))
            .expect("chain part");
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::CALC_CHAIN.to_owned(),
            "calcChain.xml".to_owned(),
            "rId3".to_owned(),
            TargetMode::Internal,
        )
        .expect("chain relationship");
    let source = Workbook::from_package(package).expect("workbook with chain");
    let source_bytes = source.to_bytes().expect("source bytes");
    let workbook_before = source
        .inner
        .package
        .get_part(&source.inner.workbook_uri)
        .expect("workbook part")
        .blob_arc();

    let mut visibility = source.edit().expect("visibility edit");
    visibility
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .row(1)
        .expect("row 2")
        .hide();
    let visibility = visibility.commit().expect("visibility commit");
    assert!(visibility.patch().graph.is_empty());
    assert!(
        visibility
            .workbook()
            .inner
            .package
            .get_part(&chain_uri)
            .is_ok()
    );
    assert_eq!(
        visibility
            .workbook()
            .inner
            .package
            .get_part(&source.inner.workbook_uri)
            .expect("unchanged workbook part")
            .blob(),
        workbook_before.as_slice()
    );

    let mut edit = source.edit().expect("edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("set");
    let committed = edit.commit().expect("commit");
    assert!(
        committed
            .workbook()
            .inner
            .package
            .get_part(&chain_uri)
            .is_err()
    );
    assert!(
        calculation_chain_removal(committed.workbook())
            .expect("chain query")
            .is_empty()
    );

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored"),
        source_bytes
    );
    assert!(
        !calculation_chain_removal(restored.workbook())
            .expect("restored chain")
            .is_empty()
    );

    let mut shared_target = source.inner.package.clone();
    let mut referrer = BlobPart::new(
        PackURI::new("/xl/custom.xml").expect("custom URI"),
        "application/xml".to_owned(),
        b"<custom/>".to_vec(),
    );
    referrer
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test-reference".to_owned(),
            "calcChain.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("extra incoming relationship");
    shared_target
        .try_add_part(Box::new(referrer))
        .expect("referrer part");
    let shared_target = Workbook::from_package(shared_target).expect("shared target workbook");
    assert!(matches!(
        calculation_chain_removal(&shared_target),
        Err(Error::Invalid(message)) if message.contains("another incoming relationship")
    ));

    let mut outbound = source.inner.package.clone();
    outbound
        .get_part_mut(&chain_uri)
        .expect("chain part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("chain outbound relationship");
    let outbound = Workbook::from_package(outbound).expect("outbound chain workbook");
    let mut edit = outbound.edit().expect("outbound chain edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("set");
    assert!(matches!(
        edit.commit(),
        Err(Error::Invalid(message)) if message.contains("calculation-chain part cannot have relationships")
    ));
    assert!(outbound.inner.package.get_part(&chain_uri).is_ok());
}

#[test]
fn shared_style_crud_is_lineage_checked_reversible_and_exact() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let styles = source.styles().expect("styles");
    assert_eq!(styles.len(), 2);
    let base = styles.base().expect("base style");
    let accent = styles.get(1).expect("accent style");
    assert_eq!(base.fan_out().expect("base fan-out"), 1);
    assert_eq!(accent.fan_out().expect("accent fan-out"), 1);

    let sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.local_style("B1").expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    let Some(crate::LocalStyle::Shared(local)) = sheet.local_style("A1").expect("local style")
    else {
        panic!("A1 must have an explicit shared style")
    };
    assert!(local.same(&accent));
    assert!(
        sheet
            .style("B1")
            .expect("resolved style")
            .is_some_and(|style| style.same(&base))
    );

    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
    sheet
        .set("C1", 42_i32)
        .and_then(|sheet| sheet.style("C1", &accent))
        .and_then(|sheet| sheet.style("D1", &accent))
        .expect("style changes");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().len(), 2);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, State::Missing, _))
    ));
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, _, State::Cell {
            content: Cell::Value(Value::Number(number)),
            style: StyleState::Shared(_),
            ..
        })) if number.as_str() == "42"
    ));

    let book = committed.workbook();
    let sheet = book.sheet(0usize).expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.cell("D1").expect("D1").stored(),
        Some(Cell::Empty)
    ));
    let styles = book.styles().expect("styles");
    let inherited = styles.find(&accent.key()).expect("inherited style key");
    assert!(inherited.same(&accent));
    assert!(!inherited.same_workbook(&accent));
    assert_eq!(accent.fan_out().expect("source fan-out"), 1);
    assert_eq!(inherited.fan_out().expect("descendant fan-out"), 3);
    assert!(
        sheet
            .style("C1")
            .expect("style")
            .is_some_and(|style| style.same(&inherited))
    );

    let mut descendant = book.edit().expect("descendant edit");
    descendant
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet")
        .style("E1", &accent)
        .expect("reuse inherited style lineage");
    let descendant = descendant.commit().expect("descendant commit");
    assert!(matches!(
        descendant
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .local_style("E1")
            .expect("local style"),
        Some(crate::LocalStyle::Shared(_))
    ));

    let reopened = Workbook::from_bytes(source_bytes.clone()).expect("reopened source");
    assert!(
        reopened
            .styles()
            .expect("reopened styles")
            .find(&accent.key())
            .is_none()
    );
    let replayed = reopened
        .apply(committed.patch())
        .expect("replay onto byte-identical source");
    let (
        _,
        _,
        State::Cell {
            style: StyleState::Shared(replayed_key),
            ..
        },
    ) = replayed.patch().changes()[0].cell().expect("cell change")
    else {
        panic!("replayed change must retain its shared style")
    };
    assert!(
        replayed
            .workbook()
            .styles()
            .expect("replayed styles")
            .find(replayed_key)
            .is_some()
    );
    assert!(
        book.styles()
            .expect("original lineage")
            .find(replayed_key)
            .is_none()
    );

    let restored = book
        .apply(&committed.patch().inverse())
        .expect("inverse patch");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );

    let foreign = Workbook::new()
        .expect("other workbook")
        .styles()
        .expect("other styles")
        .base()
        .expect("other base style");
    let mut edit = source.edit().expect("edit");
    assert!(matches!(
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("A1", &foreign),
        Err(Error::ForeignStyle)
    ));

    let mut changed_package = source.inner.package.clone();
    let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
    let changed_xml = {
        let styles = changed_package.get_part(&styles_uri).expect("styles part");
        std::str::from_utf8(styles.blob())
            .expect("UTF-8 styles")
            .replace("FFFFFF00", "FFFF0000")
            .into_bytes()
    };
    changed_package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(changed_xml);
    let changed = Workbook::from_package_with_styles(changed_package, Some(&source))
        .expect("changed style table");
    assert!(
        changed
            .styles()
            .expect("changed styles")
            .find(&accent.key())
            .is_none()
    );
    let mut edit = changed.edit().expect("changed edit");
    assert!(matches!(
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("A1", &accent),
        Err(Error::ForeignStyle)
    ));
    assert!(matches!(
        changed.apply(committed.patch()),
        Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
    ));
}

#[test]
fn payload_and_style_effects_on_one_cell_join_without_locks() {
    let source = styled_workbook();
    let accent = source.styles().expect("styles").get(1).expect("accent");
    let (mut payload, style) = std::thread::scope(|scope| {
        let payload_source = source.clone();
        let style_source = source.clone();
        let accent = accent.clone();
        let payload = scope.spawn(move || {
            let mut edit = payload_source.edit().expect("payload edit");
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .set("B1", 9_i32)
                .expect("payload");
            edit
        });
        let style = scope.spawn(move || {
            let mut edit = style_source.edit().expect("style edit");
            edit.sheet("Sheet1")
                .expect("lookup")
                .expect("sheet")
                .style("B1", &accent)
                .expect("style");
            edit
        });
        (
            payload.join().expect("payload worker"),
            style.join().expect("style worker"),
        )
    });
    payload.join(style).expect("disjoint cell facets");
    assert_eq!(payload.len(), 1);
    let committed = payload.commit().expect("commit");
    assert_eq!(committed.patch().len(), 1);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((_, _, State::Cell {
            content: Cell::Value(Value::Number(number)),
            style: StyleState::Shared(_),
            ..
        })) if number.as_str() == "9"
    ));

    let sheet = committed
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.cell("B1").expect("B1").stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "9"
    ));
    assert!(matches!(
        sheet.local_style("B1").expect("style"),
        Some(crate::LocalStyle::Shared(_))
    ));
}

#[test]
fn rich_shared_string_transfer_keeps_exact_dependency_through_join_replay_and_inverse() {
    let baseline = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = baseline.inner.package.clone();
    let shared_uri = PackURI::new("/xl/sharedStrings.xml").expect("shared strings URI");
    package
        .try_add_part(Box::new(BlobPart::new(
            shared_uri.clone(),
            litchi_opc::constants::content_type::SML_SHARED_STRINGS.to_owned(),
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Rich</t></r><r><rPr><i/></rPr><t> text</t></r></si></sst>"#.to_vec(),
        )))
        .expect("shared strings part");
    package
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::SHARED_STRINGS.to_owned(),
            "sharedStrings.xml".to_owned(),
            "rIdSharedStrings".to_owned(),
            TargetMode::Internal,
        )
        .expect("shared strings relationship");
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#.to_vec(),
        );
    let source = Workbook::from_package(package).expect("rich-string workbook");
    let source_bytes = source.to_bytes().expect("source bytes");
    let shared_before = part_text(&source, "/xl/sharedStrings.xml").to_owned();

    let mut transfer = source.edit().expect("transfer edit");
    transfer
        .copy_cells("Sheet1", "A1", "Sheet2", "B2")
        .expect("copy rich string");
    let mut disjoint = source.edit().expect("disjoint edit");
    disjoint
        .sheet("Sheet2")
        .expect("sheet lookup")
        .expect("worksheet")
        .set("C3", 7_i32)
        .expect("disjoint value");
    transfer.join(disjoint).expect("join disjoint edit");
    let committed = transfer.commit().expect("commit rich transfer");

    let target_xml = part_text(committed.workbook(), "/xl/worksheets/sheet2.xml");
    assert!(target_xml.contains(r#"<c r="B2" t="s"><v>0</v></c>"#));
    assert_eq!(
        part_text(committed.workbook(), "/xl/sharedStrings.xml"),
        shared_before
    );
    assert!(committed.patch().changes().iter().any(|change| matches!(
        change.cell(),
        Some((
            _,
            _,
            State::Cell {
                shared_string: Some(_),
                ..
            }
        ))
    )));

    let replayed = source
        .apply(committed.patch())
        .expect("replay rich transfer");
    assert!(
        part_text(replayed.workbook(), "/xl/worksheets/sheet2.xml")
            .contains(r#"<c r="B2" t="s"><v>0</v></c>"#)
    );
    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse rich transfer");
    assert_eq!(
        restored.workbook().to_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn cell_transfer_refuses_drawing_graphs_without_staging_partial_work() {
    let baseline = two_sheet_workbook(WorksheetKind::Worksheet);
    let mut package = baseline.inner.package.clone();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/drawings/drawing1.xml").expect("drawing URI"),
            litchi_opc::constants::content_type::OFC_DRAWING.to_owned(),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#.to_vec(),
        )))
        .expect("drawing part");
    package
        .get_part_mut(&baseline.inner.sheets[0].part_uri)
        .expect("source worksheet")
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DRAWING.to_owned(),
            "../drawings/drawing1.xml".to_owned(),
            "rIdDrawing".to_owned(),
            TargetMode::Internal,
        )
        .expect("drawing relationship");
    let source = Workbook::from_package(package).expect("drawing workbook");
    let mut edit = source.edit().expect("edit");
    assert!(matches!(
        edit.copy_cells("Sheet1", "A1", "Sheet2", "A1"),
        Err(Error::Unsupported { feature })
            if feature == "copying cells on a worksheet with drawing dependencies"
    ));
    assert!(edit.is_empty());
}

#[test]
fn resetting_style_is_distinct_from_removal_and_missing_is_a_no_op() {
    let source = styled_workbook();
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("edit");
    let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
    sheet
        .reset_style("A1")
        .and_then(|sheet| sheet.reset_style("Z99"))
        .expect("style resets");
    let committed = edit.commit().expect("commit");
    assert_eq!(committed.patch().len(), 1);
    assert!(matches!(
        committed.patch().changes()[0].cell(),
        Some((
            _,
            _,
            State::Cell {
                style: StyleState::Default,
                ..
            }
        ))
    ));
    let sheet = committed
        .workbook()
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    assert!(matches!(
        sheet.local_style("A1").expect("local style"),
        Some(crate::LocalStyle::Default)
    ));
    assert!(sheet.cell("Z99").expect("missing").is_missing());

    let restored = committed
        .workbook()
        .apply(&committed.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
}

#[test]
fn signed_packages_refuse_edits_before_mutation() {
    let baseline = Workbook::new().expect("baseline");
    let mut package = baseline.inner.package.clone();
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    let signed = Workbook::from_package(package).expect("signed workbook snapshot");

    assert!(matches!(signed.edit(), Err(Error::Signed)));
    assert!(matches!(
        signed.apply(&Patch::default()),
        Err(Error::Signed)
    ));
}
