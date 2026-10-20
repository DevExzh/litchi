//! The capabilities of the removed ODS part-level patches on the package
//! route (record 0358a): every part closure of `Package::edit` publishes a
//! package `Patch` that serializes deterministically, replays only onto its
//! exact source, inverts, and joins with disjoint part edits.

use litchi_ods::database_range::Range;
use litchi_ods::styles::table_template::{Region, Style, Template};
use litchi_ods::{Builder, Package, Sheet};

fn source() -> Package {
    let mut builder = Builder::new();
    builder.add_sheet(Sheet::new("Data").unwrap()).unwrap();
    builder.build().unwrap()
}

fn with_ranges(package: &Package) -> litchi_ods::document::Commit {
    let mut edit = package.edit().unwrap();
    edit.database_ranges(|ranges| {
        let mut range = Range::new("Data.A1:Data.B2");
        range.name = Some("Sales".to_string());
        ranges.editor().add(range)
    })
    .unwrap();
    edit.commit().unwrap()
}

fn with_templates(package: &Package) -> litchi_ods::document::Commit {
    let mut edit = package.edit().unwrap();
    edit.table_templates(|templates| {
        templates.add(Template::new("Bands").with_region(Region::Body, Style::new("Body")))
    })
    .unwrap();
    edit.commit().unwrap()
}

#[test]
fn part_closures_publish_durable_reversible_patches() {
    let source = source();
    for commit in [with_ranges(&source), with_templates(&source)] {
        let wire = commit.patch().to_deterministic_json().unwrap();
        let patch = litchi_ods::document::Patch::from_deterministic_json(&wire, Default::default())
            .unwrap();
        let replayed = source.apply(&patch).unwrap();
        assert_eq!(replayed.package().as_bytes(), commit.package().as_bytes());
        let restored = commit.package().apply(&patch.inverse()).unwrap();
        assert_eq!(restored.package().as_bytes(), source.as_bytes());
        assert!(
            commit.package().apply(&patch).is_err(),
            "a stale source is refused"
        );
    }
}

#[test]
fn part_edits_plan_a_three_way_merge() {
    let source = source();
    let ranges = with_ranges(&source);
    let mut settings = source.edit().unwrap();
    settings
        .set_settings(Some(&litchi_ods::Settings::default()))
        .unwrap();
    let settings = settings.commit().unwrap();
    // Both branches rewrite content.xml: the plan reports the overlap and
    // an explicit choice finishes it.
    let mut plan =
        litchi_ods::document::Patch::three_way(ranges.patch(), settings.patch()).unwrap();
    plan.resolve(litchi_core::MergeChoice::Left);
    let merged = plan.finish().unwrap();
    let workbook = source
        .apply(&merged)
        .unwrap()
        .into_package()
        .workbook()
        .unwrap();
    assert_eq!(workbook.database_ranges().unwrap().len(), 1);
}
