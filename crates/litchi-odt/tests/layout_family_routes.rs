//! The layout and protection capabilities of the removed part-level families
//! (`header_footer::{DurablePatch, History, MergePlan}` and the protection
//! policy transaction's durable patch, transfer and merge) on the package
//! route (record 0358a): durable patches, bounded history, joined and
//! three-way merged edits of disjoint layout owners, and transfers.

use litchi_core::{HistoryLimits, MergeChoice};
use litchi_odt::header_footer::{Kind, SectionLayout, Snapshot};
use litchi_odt::protection::Policy;
use litchi_odt::section_properties::{SectionLength, SectionProperties, SectionStyleProperties};
use litchi_odt::style::columns::Columns;
use litchi_odt::transaction::{DurablePatch, MergePlan};
use litchi_odt::{Builder, Package};

fn source() -> Package {
    let mut builder = Builder::new();
    builder.add_paragraph("body").unwrap();
    let package = builder.build().unwrap();
    let mut edit = package.edit();
    edit.add_master_page("Report", "pm1").unwrap();
    edit.insert_section_layout(&section("Sect", 2, "1cm"))
        .unwrap();
    edit.commit().unwrap().into_package()
}

fn section(name: &str, columns: u8, margin: &str) -> SectionLayout {
    let style = SectionStyleProperties::new(
        name,
        SectionProperties {
            margin_left: Some(SectionLength::new(margin).unwrap()),
            ..Default::default()
        },
    )
    .unwrap();
    SectionLayout::new(style, Some(Columns::new(columns).unwrap())).unwrap()
}

fn layouts(package: &Package) -> Snapshot {
    package.document().unwrap().layouts().unwrap()
}

fn columns_of(package: &Package, name: &str) -> u8 {
    layouts(package)
        .section_layouts()
        .iter()
        .find(|layout| layout.name() == name)
        .and_then(SectionLayout::columns)
        .map(|columns| columns.column_count)
        .unwrap()
}

#[test]
fn a_layout_edit_has_a_durable_reversible_patch() {
    let source = source();
    let mut edit = source.edit();
    edit.replace_section_layout(&section("Sect", 3, "1.5cm"))
        .unwrap();
    edit.set_header_footer_text("Report", Kind::Header, "changed")
        .unwrap();
    let commit = edit.commit().unwrap();

    let wire = commit
        .patch()
        .durable()
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let durable = DurablePatch::from_deterministic_json(&wire).unwrap();
    let replayed = durable.apply(&source).unwrap();
    assert_eq!(replayed.as_bytes(), commit.package().as_bytes());
    assert_eq!(columns_of(&replayed, "Sect"), 3);
    let restored = durable.inverse().apply(&replayed).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
    assert!(
        durable.apply(&replayed).is_err(),
        "a stale source is refused"
    );
}

#[test]
fn layout_edits_record_in_bounded_history() {
    let source = source();
    let mut history = source.history(HistoryLimits::new(4, 64 * 1024 * 1024));
    let mut edit = history.edit();
    edit.replace_section_layout(&section("Sect", 3, "1.5cm"))
        .unwrap();
    history.commit(edit).unwrap();
    assert_eq!(columns_of(history.current(), "Sect"), 3);
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), source.as_bytes());
    assert!(history.redo());
    assert_eq!(columns_of(history.current(), "Sect"), 3);
}

#[test]
fn disjoint_layout_owners_merge_and_one_owner_conflicts() {
    let source = source();
    let mut left = source.joined_edits();
    let mut left_edit = source.edit();
    left_edit
        .replace_section_layout(&section("Sect", 3, "1.5cm"))
        .unwrap();
    left.join("left", left_edit).unwrap();
    let mut right = source.joined_edits();
    let mut right_edit = source.edit();
    right_edit
        .set_header_footer_text("Report", Kind::Footer, "footer")
        .unwrap();
    right.join("right", right_edit).unwrap();

    let plan = MergePlan::new(left, right).unwrap();
    assert!(plan.conflicts().is_empty());
    let merged = match plan.finish() {
        Ok(joined) => joined.commit().unwrap().into_package(),
        Err(_) => panic!("disjoint layout owners merge without conflict"),
    };
    assert_eq!(columns_of(&merged, "Sect"), 3);
    let master = layouts(&merged)
        .master_pages()
        .iter()
        .find(|master| master.name == "Report")
        .cloned()
        .unwrap();
    assert_eq!(master.region(Kind::Footer).unwrap().text, "footer");

    let mut left = source.joined_edits();
    let mut one = source.edit();
    one.replace_section_layout(&section("Sect", 3, "1.5cm"))
        .unwrap();
    left.join("left", one).unwrap();
    let mut right = source.joined_edits();
    let mut other = source.edit();
    other
        .replace_section_layout(&section("Sect", 4, "2cm"))
        .unwrap();
    right.join("right", other).unwrap();
    let mut plan = MergePlan::new(left, right).unwrap();
    assert_eq!(plan.conflicts().len(), 1);
    plan.resolve(MergeChoice::Right);
    let merged = match plan.finish() {
        Ok(joined) => joined.commit().unwrap().into_package(),
        Err(_) => panic!("a resolved layout conflict finishes"),
    };
    assert_eq!(columns_of(&merged, "Sect"), 4);
}

#[test]
fn layouts_and_protection_transfer_to_another_document() {
    let source = source();
    let mut protect = source.edit();
    let policy = Policy {
        forms: Some(true),
        ..Policy::default()
    };
    protect.set_protection(&policy).unwrap();
    let source = protect.commit().unwrap().into_package();

    let transfer = layouts(&source)
        .prepare_section_layout_transfer("Sect")
        .unwrap();
    let mut builder = Builder::new();
    builder.add_paragraph("destination").unwrap();
    let destination = builder.build().unwrap();
    let mut edit = destination.edit();
    edit.insert_section_layout_transfer(&transfer, &[]).unwrap();
    edit.set_protection(&source.document().unwrap().protection().unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(columns_of(commit.package(), "Sect"), 2);
    assert_eq!(
        commit.package().document().unwrap().protection().unwrap(),
        policy
    );
    // The protection change has a semantic durable form too.
    let wire = commit
        .patch()
        .durable()
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let durable = DurablePatch::from_deterministic_json(&wire).unwrap();
    assert_eq!(
        durable.apply(&destination).unwrap().as_bytes(),
        commit.package().as_bytes()
    );
}

#[test]
fn conflicting_protection_branches_need_an_explicit_choice() {
    let source = source();
    let branch = |name: &str, forms: bool| {
        let mut joined = source.joined_edits();
        let mut edit = source.edit();
        edit.set_protection(&Policy {
            forms: Some(forms),
            ..Policy::default()
        })
        .unwrap();
        joined.join(name, edit).unwrap();
        joined
    };
    let mut plan = MergePlan::new(branch("left", true), branch("right", false)).unwrap();
    assert_eq!(plan.conflicts().len(), 1);
    plan.resolve(MergeChoice::Left);
    let merged = match plan.finish() {
        Ok(joined) => joined.commit().unwrap().into_package(),
        Err(_) => panic!("a resolved protection conflict finishes"),
    };
    assert_eq!(
        merged.document().unwrap().protection().unwrap().forms,
        Some(true)
    );
}
