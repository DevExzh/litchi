//! Unified reads and writes of the OpenDocument object formats (record
//! 0282): drawings, charts, images, databases and formulas, through the
//! object-tree adapter over each leaf's `Package` and `Edit`.

#![cfg(any(
    feature = "odg",
    feature = "odc",
    feature = "odi",
    feature = "odb",
    feature = "odf-formula"
))]
// The shared helpers serve whichever format modules the build compiles.
#![cfg_attr(
    not(all(
        feature = "odg",
        feature = "odc",
        feature = "odi",
        feature = "odb",
        feature = "odf-formula"
    )),
    allow(dead_code, unused_imports)
)]

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, File, create, open, open_bytes, props};

fn fixture(path: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path)
}

fn read(path: &str) -> File {
    open_bytes(std::fs::read(fixture(path)).expect("read fixture")).expect("open")
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).expect("path");
    let format = node.view().format().expect("format");
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("path")
        .view()
        .text()
        .expect("text")
        .map(std::borrow::Cow::into_owned)
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

fn assert_supported(format: FileFormat, verbs: &[Verb]) {
    let report = Capabilities::of(format);
    for verb in verbs {
        assert!(report.verb(*verb).is_supported(), "{format:?} {verb:?}");
    }
}

#[cfg(feature = "odg")]
mod odg {
    use super::*;
    use litchi_ops::NewNode;

    const THREE_PAGES: &str = "libreoffice-core/desktop/qa/data/3page.odg";

    #[test]
    fn reads_pages_and_shapes() {
        let file = read(THREE_PAGES);
        assert_eq!(file.format(), Some(FileFormat::Odg));
        assert!(matches!(file, File::Odg(_)));
        assert_eq!(prop(&file, "/", "pageCount").as_deref(), Some("3"));
        assert_eq!(prop(&file, "/page[2]", "name").as_deref(), Some("page2"));
        assert_eq!(
            file.get("/page[page3]").unwrap().path().to_string(),
            "/page[3]"
        );
        assert_eq!(file.query("shape").unwrap().count(), 3);
        assert_eq!(
            file.get("/page[1]/shape[1]").unwrap().format(),
            FileFormat::Odg
        );
    }

    /// A producer drawing whose shape text is one replaceable span.
    const FONT_WORK: &str = "libreoffice-core/svx/qa/unit/data/FontWork.odg";

    #[test]
    fn writes_names_text_geometry_and_structure() {
        let file = read(FONT_WORK);
        let before = file.to_bytes().unwrap();
        let file = file
            .set("/page[1]", props! { "name" => "Cover" })
            .expect("page name")
            .set(
                "/page[1]/shape[1]",
                props! { "text" => "Hello", "name" => "Box", "x" => "2cm" },
            )
            .expect("shape set");
        let file = reopen(&file);
        assert_eq!(prop(&file, "/page[1]", "name").as_deref(), Some("Cover"));
        assert_eq!(text(&file, "/page[1]/shape[1]").as_deref(), Some("Hello"));
        assert_eq!(
            prop(&file, "/page[1]/shape[1]", "x").as_deref(),
            Some("2cm")
        );
        // Structural writes on a LibreOffice-authored drawing publish as
        // exact-source splices (record 0288).
        let producer = file
            .add(
                "/page[1]",
                NewNode::new("shape")
                    .prop("x", "1cm")
                    .prop("y", "1cm")
                    .prop("width", "4cm")
                    .prop("height", "2cm")
                    .prop("text", "Producer"),
            )
            .expect("producer insert");
        let producer = reopen(&producer.file);
        assert_eq!(
            text(&producer, "/page[1]/shape[last()]").as_deref(),
            Some("Producer")
        );
        let fresh = create(FileFormat::Odg)
            .unwrap()
            .add("/", NewNode::new("page").prop("name", "One"))
            .expect("add page")
            .file;
        let added = fresh
            .add(
                "/page[1]",
                NewNode::new("shape")
                    .prop("x", "1cm")
                    .prop("y", "1cm")
                    .prop("width", "4cm")
                    .prop("height", "2cm")
                    .prop("text", "Added"),
            )
            .expect("add shape");
        assert_eq!(added.path.to_string(), "/page[1]/shape[1]");
        let fresh = reopen(&added.file);
        assert_eq!(text(&fresh, "/page[1]/shape[1]").as_deref(), Some("Added"));
        let fresh = fresh
            .set(
                "/page[1]/shape[1]",
                props! { "text" => "Changed", "width" => "6cm" },
            )
            .expect("set added shape");
        let fresh = reopen(&fresh);
        assert_eq!(
            text(&fresh, "/page[1]/shape[1]").as_deref(),
            Some("Changed")
        );
        assert_eq!(
            prop(&fresh, "/page[1]/shape[1]", "width").as_deref(),
            Some("6cm")
        );
        let fresh = fresh.remove("/page[1]/shape[1]").expect("remove");
        assert_eq!(
            prop(&reopen(&fresh), "/page[1]", "shapeCount").as_deref(),
            Some("0")
        );
        let fresh = fresh
            .add("/", NewNode::new("page").prop("name", "Two"))
            .expect("add page")
            .file;
        assert_eq!(
            prop(&reopen(&fresh), "/", "pageCount").as_deref(),
            Some("2")
        );
        assert_ne!(file.to_bytes().unwrap(), before);
    }

    fn page_names(file: &File) -> Vec<String> {
        let count: usize = prop(file, "/", "pageCount").unwrap().parse().unwrap();
        (1..=count)
            .map(|index| prop(file, &format!("/page[{index}]"), "name").unwrap())
            .collect()
    }

    #[test]
    fn adds_removes_moves_and_swaps_producer_pages() {
        let file = read(THREE_PAGES);
        let added = file
            .add(
                "/",
                NewNode::new("page")
                    .prop("name", "first")
                    .at(litchi_ops::Placement::Start),
            )
            .expect("insert page");
        assert_eq!(added.path.to_string(), "/page[1]");
        let file = reopen(&added.file);
        assert_eq!(page_names(&file), ["first", "page1", "page2", "page3"]);
        let file = file.remove("/page[page2]").expect("remove page");
        assert_eq!(page_names(&reopen(&file)), ["first", "page1", "page3"]);
        let file = file
            .move_to("/page[1]", litchi_ops::Placement::End)
            .expect("move page");
        assert_eq!(page_names(&reopen(&file)), ["page1", "page3", "first"]);
        let file = file.swap("/page[1]", "/page[3]").expect("swap pages");
        assert_eq!(page_names(&reopen(&file)), ["first", "page3", "page1"]);
        let added = file
            .add(
                "/page[2]",
                NewNode::new("shape")
                    .prop("x", "1cm")
                    .prop("y", "1cm")
                    .prop("width", "2cm")
                    .prop("height", "2cm")
                    .prop("text", "Front")
                    .at(litchi_ops::Placement::Start),
            )
            .expect("insert shape first");
        assert_eq!(added.path.to_string(), "/page[2]/shape[1]");
        let file = reopen(&added.file);
        assert_eq!(text(&file, "/page[2]/shape[1]").as_deref(), Some("Front"));
        let count: usize = prop(&file, "/page[2]", "shapeCount")
            .unwrap()
            .parse()
            .unwrap();
        assert!(count >= 2);
        let file = file
            .move_to("/page[2]/shape[1]", litchi_ops::Placement::End)
            .expect("move shape");
        let file = reopen(&file);
        assert_eq!(
            text(&file, &format!("/page[2]/shape[{count}]")).as_deref(),
            Some("Front")
        );
        // A shape moves into another page, here the empty first page
        // (record 0292), and swaps across pages.
        let moved_text = text(&file, &format!("/page[2]/shape[{count}]"));
        let file = file
            .move_to(
                format!("/page[2]/shape[{count}]").as_str(),
                litchi_ops::Placement::Into(
                    litchi_ops::IntoPath::into_path("/page[1]").unwrap(),
                    None,
                ),
            )
            .expect("cross-page move");
        let file = reopen(&file);
        assert_eq!(prop(&file, "/page[1]", "shapeCount").as_deref(), Some("1"));
        assert_eq!(text(&file, "/page[1]/shape[1]"), moved_text);
        let other = text(&file, "/page[2]/shape[1]");
        let file = file
            .swap("/page[1]/shape[1]", "/page[2]/shape[1]")
            .expect("cross-page swap");
        let file = reopen(&file);
        assert_eq!(text(&file, "/page[1]/shape[1]"), other);
        assert_eq!(text(&file, "/page[2]/shape[1]"), moved_text);
        let file = file
            .move_to(
                "/page[1]/shape[1]",
                litchi_ops::Placement::Before(
                    litchi_ops::IntoPath::into_path("/page[2]/shape[1]").unwrap(),
                ),
            )
            .expect("anchored cross-page move");
        let file = reopen(&file);
        assert_eq!(prop(&file, "/page[1]", "shapeCount").as_deref(), Some("0"));
        assert_eq!(text(&file, "/page[2]/shape[1]"), other);
        let error = file
            .move_to(
                "/page[2]",
                litchi_ops::Placement::Position(litchi_core::Position::new(9)),
            )
            .expect_err("out of range");
        assert_eq!(error.code(), Code::OutOfRange);
    }

    /// A drawing with one page of `count` text rectangles.
    fn drawing_with_shapes(count: usize) -> File {
        let mut file = create(FileFormat::Odg)
            .unwrap()
            .add("/", NewNode::new("page").prop("name", "One"))
            .unwrap()
            .file;
        file = file
            .edit(|tx| {
                for index in 0..count {
                    tx.add(
                        "/page[1]",
                        NewNode::new("shape")
                            .prop("x", "1cm")
                            .prop("y", "1cm")
                            .prop("width", "2cm")
                            .prop("height", "1cm")
                            .prop("text", format!("s{index}").as_str())
                            .prop("name", format!("n{index}").as_str()),
                    )?;
                }
                Ok(())
            })
            .unwrap()
            .file()
            .clone();
        file
    }

    #[test]
    fn batches_coalesce_into_one_leaf_commit() {
        let file = drawing_with_shapes(6);
        let commit = file
            .edit(|tx| {
                for index in 1..=6 {
                    tx.set(
                        format!("/page[1]/shape[{index}]").as_str(),
                        props! { "text" => format!("t{index}").as_str() },
                    )?;
                }
                // Another property family of an already-written shape.
                tx.set(
                    "/page[1]/shape[1]",
                    props! { "name" => "First", "x" => "3cm" },
                )?;
                tx.set("/page[1]", props! { "name" => "Renamed" })?;
                Ok(())
            })
            .unwrap();
        assert_eq!(commit.diagnostics().operations(), 8);
        assert_eq!(commit.diagnostics().commits(), 1);
        let after = reopen(commit.file());
        assert_eq!(text(&after, "/page[1]/shape[6]").as_deref(), Some("t6"));
        assert_eq!(
            prop(&after, "/page[1]/shape[1]", "name").as_deref(),
            Some("First")
        );
        assert_eq!(
            prop(&after, "/page[1]/shape[1]", "x").as_deref(),
            Some("3cm")
        );
        assert_eq!(prop(&after, "/page[1]", "name").as_deref(), Some("Renamed"));
        let restored = commit.patch().inverse().apply(commit.file()).unwrap();
        assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());

        // A second write of one shape's text, a second geometry write and a
        // name-addressed set each see the published result.
        let commit = file
            .edit(|tx| {
                tx.set("/page[1]/shape[1]", props! { "text" => "a", "x" => "4cm" })?;
                tx.set("/page[1]/shape[1]", props! { "width" => "5cm" })?;
                tx.set("/page[1]/shape[1]", props! { "text" => "b" })?;
                tx.set("/page[One]/shape[2]", props! { "text" => "c" })?;
                Ok(())
            })
            .unwrap();
        assert!(commit.diagnostics().commits() >= 3);
        let after = reopen(commit.file());
        assert_eq!(text(&after, "/page[1]/shape[1]").as_deref(), Some("b"));
        assert_eq!(
            prop(&after, "/page[1]/shape[1]", "x").as_deref(),
            Some("4cm")
        );
        assert_eq!(
            prop(&after, "/page[1]/shape[1]", "width").as_deref(),
            Some("5cm")
        );
        assert_eq!(text(&after, "/page[1]/shape[2]").as_deref(), Some("c"));

        // A failed batch leaves the input unchanged and names the item.
        let before = file.to_bytes().unwrap();
        let error = file
            .edit(|tx| {
                tx.set("/page[1]/shape[1]", props! { "text" => "ok" })?;
                tx.set("/page[1]/shape[99]", props! { "text" => "missing" })?;
                Ok(())
            })
            .unwrap_err();
        assert_eq!(error.item(), Some(1));
        assert_eq!(file.to_bytes().unwrap(), before);
    }

    #[test]
    fn metadata_sets_through_the_family() {
        let file = read(THREE_PAGES)
            .set(
                "/metadata",
                props! { "title" => "Drawn", "author" => "Ann", "keywords" => "a, b" },
            )
            .expect("metadata");
        let file = reopen(&file);
        let metadata = file.metadata().unwrap().unwrap();
        assert_eq!(metadata.title.as_deref(), Some("Drawn"));
        assert_eq!(prop(&file, "/", "title").as_deref(), Some("Drawn"));
        assert_eq!(prop(&file, "/", "author").as_deref(), Some("Ann"));
        let error = file
            .set("/metadata", props! { "created" => "2020-01-01" })
            .expect_err("dates are producer-maintained");
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }

    #[test]
    fn unchanged_values_are_exact_no_ops() {
        let file = read(THREE_PAGES);
        let before = file.to_bytes().unwrap();
        let file = file.set("/page[2]", props! { "name" => "page2" }).unwrap();
        assert_eq!(file.to_bytes().unwrap(), before);
    }

    #[test]
    fn refusals_are_typed() {
        let file = read(THREE_PAGES);
        let error = file
            .set("/page[1]/layer[1]", props! { "name" => "x" })
            .expect_err("layers are read-only");
        assert_eq!(error.format(), Some(FileFormat::Odg));
        let error = file
            .set("/page[1]/shape[1]", props! { "shapeType" => "ellipse" })
            .expect_err("read-only");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file.get("/slide[1]").expect_err("no slides");
        assert_ne!(error.code(), Code::InternalError);
    }

    #[test]
    fn create_save_and_capabilities() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("new.odg");
        let file = create(FileFormat::Odg).expect("create");
        assert_eq!(file.format(), Some(FileFormat::Odg));
        let file = file
            .add("/", NewNode::new("page").prop("name", "First"))
            .expect("add page")
            .file;
        file.save_as(&path).expect("save");
        let reopened = open(&path).expect("reopen");
        let file = reopened
            .set("/page[1]", props! { "name" => "Named" })
            .expect("set");
        file.save().expect("save in place");
        assert_eq!(
            prop(&open(&path).unwrap(), "/page[1]", "name").as_deref(),
            Some("Named")
        );
        assert_supported(
            FileFormat::Odg,
            &[
                Verb::Get,
                Verb::Query,
                Verb::Set,
                Verb::Add,
                Verb::Remove,
                Verb::Move,
                Verb::Swap,
                Verb::Save,
                Verb::Create,
            ],
        );
    }
}

#[cfg(feature = "odc")]
mod odc {
    use super::*;
    use litchi_ops::NewNode;

    const CHART: &str = "odf/odc-producer-evidence/odfdom-created.odc";

    #[test]
    fn reads_axes_and_series() {
        let file = read(CHART);
        assert_eq!(file.format(), Some(FileFormat::Odc));
        assert_eq!(prop(&file, "/axis[1]", "name").as_deref(), Some("axis-x"));
        assert_eq!(prop(&file, "/axis[1]", "dimension").as_deref(), Some("x"));
        assert_eq!(file.query("series").unwrap().count(), 1);
        let named = file
            .query("axis[name=axis-y]")
            .unwrap()
            .next()
            .unwrap()
            .unwrap();
        assert_eq!(named.path().to_string(), "/axis[2]");
    }

    #[test]
    fn renames_axes_and_removes_series() {
        let file = read(CHART);
        let file = file
            .set("/axis[1]", props! { "name" => "horizontal" })
            .expect("rename axis");
        let file = reopen(&file);
        assert_eq!(
            prop(&file, "/axis[1]", "name").as_deref(),
            Some("horizontal")
        );
        // This producer chart holds content outside the leaf's lossless
        // typed surface, so a structural removal and a title write are typed
        // refusals; attribute writes stay exact-span splices (record 0288).
        let error = file.remove("/series[1]").expect_err("lossy typed edit");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .set("/", props! { "title" => "x" })
            .expect_err("producer chart title");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let file = file
            .set("/", props! { "chartType" => "bar" })
            .expect("chart class");
        let file = file
            .set("/series[1]", props! { "chartType" => "chart:line" })
            .expect("series class");
        let file = reopen(&file);
        assert_eq!(prop(&file, "/", "chartType").as_deref(), Some("chart:bar"));
        assert_eq!(
            prop(&file, "/series[1]", "chartType").as_deref(),
            Some("chart:line")
        );
        let unchanged = file
            .set("/", props! { "chartType" => "chart:bar" })
            .expect("no-op");
        assert_eq!(unchanged.to_bytes().unwrap(), file.to_bytes().unwrap());
    }

    #[test]
    fn writes_titles_types_series_and_axes_of_a_typed_chart() {
        let file = create(FileFormat::Odc).expect("create");
        let file = file
            .set(
                "/",
                props! { "title" => "Sales", "subtitle" => "FY", "chartType" => "bar" },
            )
            .expect("chart props");
        let added = file
            .add(
                "/",
                NewNode::new("axis")
                    .prop("dimension", "x")
                    .prop("name", "primary-x")
                    .prop("categories", "local-table.A2:A5"),
            )
            .expect("add x axis");
        assert_eq!(added.path.to_string(), "/axis[1]");
        let added = added
            .file
            .add(
                "/",
                NewNode::new("axis")
                    .prop("dimension", "y")
                    .prop("name", "primary-y")
                    .prop("title", "Units"),
            )
            .expect("add y axis");
        let added = added
            .file
            .add(
                "/",
                NewNode::new("series")
                    .prop("values", "local-table.B2:B5")
                    .prop("label", "local-table.B1")
                    .prop("axis", "primary-y"),
            )
            .expect("add series");
        assert_eq!(added.path.to_string(), "/series[1]");
        let file = added
            .file
            .set(
                "/series[1]",
                props! { "values" => "local-table.C2:C5", "chartType" => "line" },
            )
            .expect("series props")
            .set("/axis[2]", props! { "title" => "Count" })
            .expect("axis title")
            .set("/", props! { "width" => "12cm", "height" => "8cm" })
            .expect("chart size");
        let file = reopen(&file);
        assert_eq!(text(&file, "/").as_deref(), Some("Sales"));
        assert_eq!(prop(&file, "/", "subtitle").as_deref(), Some("FY"));
        assert_eq!(prop(&file, "/", "chartType").as_deref(), Some("chart:bar"));
        assert_eq!(prop(&file, "/", "axisCount").as_deref(), Some("2"));
        assert_eq!(
            prop(&file, "/series[1]", "values").as_deref(),
            Some("local-table.C2:C5")
        );
        assert_eq!(
            prop(&file, "/series[1]", "chartType").as_deref(),
            Some("chart:line")
        );
        assert_eq!(prop(&file, "/axis[2]", "title").as_deref(), Some("Count"));
        assert_eq!(prop(&file, "/", "width").as_deref(), Some("12cm"));
        assert_eq!(
            prop(&file, "/axis[1]", "categories").as_deref(),
            Some("local-table.A2:A5")
        );
        let file = file
            .set("/", props! { "title" => litchi_ops::PropValue::Null })
            .expect("clear");
        assert_eq!(prop(&reopen(&file), "/", "title"), None);
        let file = file.remove("/series[1]").expect("remove series");
        assert_eq!(
            prop(&reopen(&file), "/", "seriesCount").as_deref(),
            Some("0")
        );
        let error = file
            .set("/", props! { "chartType" => "pyramid" })
            .expect_err("unknown type");
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .add("/", NewNode::new("series").prop("label", "x"))
            .expect_err("values required");
        assert_ne!(error.code(), Code::InternalError);
    }

    #[test]
    fn moves_and_swaps_series_and_axes_and_edits_the_legend() {
        let mut file = create(FileFormat::Odc).expect("create");
        for (values, label) in [
            ("t.B2:B5", "t.B1"),
            ("t.C2:C5", "t.C1"),
            ("t.D2:D5", "t.D1"),
        ] {
            file = file
                .add(
                    "/",
                    NewNode::new("series")
                        .prop("values", values)
                        .prop("label", label),
                )
                .expect("add series")
                .file;
        }
        for (dimension, name) in [("x", "ax"), ("y", "ay")] {
            file = file
                .add(
                    "/",
                    NewNode::new("axis")
                        .prop("dimension", dimension)
                        .prop("name", name),
                )
                .expect("add axis")
                .file;
        }
        let file = file
            .move_to("/series[1]", litchi_ops::Placement::End)
            .expect("move series");
        let file = reopen(&file);
        assert_eq!(
            prop(&file, "/series[3]", "values").as_deref(),
            Some("t.B2:B5")
        );
        assert_eq!(
            prop(&file, "/series[1]", "values").as_deref(),
            Some("t.C2:C5")
        );
        let file = file.swap("/series[1]", "/series[3]").expect("swap series");
        let file = file.swap("/axis[1]", "/axis[2]").expect("swap axes");
        let file = reopen(&file);
        assert_eq!(
            prop(&file, "/series[1]", "values").as_deref(),
            Some("t.B2:B5")
        );
        assert_eq!(prop(&file, "/axis[1]", "name").as_deref(), Some("ay"));
        let error = file
            .swap("/series[1]", "/axis[1]")
            .expect_err("mixed kinds");
        assert_eq!(error.code(), Code::UnsupportedType);

        // The legend: add, set, remove (record 0292).
        assert!(file.find("/legend").unwrap().is_none());
        let added = file
            .add("/", NewNode::new("legend").prop("position", "bottom"))
            .expect("add legend");
        assert_eq!(added.path.to_string(), "/legend");
        let file = reopen(&added.file);
        assert_eq!(
            prop(&file, "/legend", "position").as_deref(),
            Some("bottom")
        );
        let error = file
            .add("/", NewNode::new("legend"))
            .expect_err("one legend");
        assert_eq!(error.code(), Code::InvalidArgument);
        let file = file
            .set("/legend", props! { "position" => "top" })
            .expect("legend position");
        assert_eq!(
            prop(&reopen(&file), "/legend", "position").as_deref(),
            Some("top")
        );
        let error = file
            .set("/legend", props! { "position" => "left" })
            .expect_err("unknown position");
        assert_eq!(error.code(), Code::InvalidValue);
        let file = file.remove("/legend").expect("remove legend");
        assert!(reopen(&file).find("/legend").unwrap().is_none());
    }

    #[test]
    fn producer_chart_moves_are_typed_refusals() {
        let file = read(CHART);
        let error = file
            .move_to("/axis[1]", litchi_ops::Placement::End)
            .expect_err("producer chart reorder");
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }

    #[test]
    fn create_and_capabilities() {
        let file = create(FileFormat::Odc).expect("create");
        assert_eq!(reopen(&file).format(), Some(FileFormat::Odc));
        assert_supported(
            FileFormat::Odc,
            &[
                Verb::Get,
                Verb::Query,
                Verb::Set,
                Verb::Add,
                Verb::Remove,
                Verb::Save,
                Verb::Create,
            ],
        );
    }
}

#[cfg(feature = "odi")]
mod odi {
    use super::*;

    #[test]
    fn creates_reads_and_writes_the_frame_and_metadata() {
        let file = create(FileFormat::Odi).expect("create");
        assert_eq!(file.format(), Some(FileFormat::Odi));
        assert!(matches!(file, File::Odi(_)));
        let file = reopen(&file);
        assert_eq!(prop(&file, "/", "frameCount").as_deref(), Some("1"));
        let file = file
            .set("/frame[1]", props! { "name" => "Logo", "title" => "Mark" })
            .expect("frame")
            .set(
                "/metadata",
                props! { "title" => "Image title", "author" => "Ada" },
            )
            .expect("metadata");
        let file = reopen(&file);
        assert_eq!(prop(&file, "/frame[1]", "name").as_deref(), Some("Logo"));
        assert_eq!(prop(&file, "/frame[1]", "title").as_deref(), Some("Mark"));
        // A fresh image frame has no geometry and its compact content keeps
        // no `svg` namespace declaration: the leaf declares it with the
        // inserted geometry (record 0288).
        let file = file
            .set(
                "/frame[1]",
                props! { "x" => "1cm", "y" => "2cm", "width" => "5cm", "height" => "3cm" },
            )
            .expect("fresh geometry");
        let file = reopen(&file);
        assert_eq!(prop(&file, "/frame[1]", "width").as_deref(), Some("5cm"));
        assert_eq!(prop(&file, "/frame[1]", "y").as_deref(), Some("2cm"));
        let file = file
            .set("/frame[1]", props! { "width" => "6cm" })
            .expect("one geometry value");
        assert_eq!(
            prop(&reopen(&file), "/frame[1]", "width").as_deref(),
            Some("6cm")
        );
        assert_eq!(
            prop(&file, "/metadata", "title").as_deref(),
            Some("Image title")
        );
        assert_eq!(
            file.metadata()
                .unwrap()
                .and_then(|metadata| metadata.author),
            Some("Ada".to_owned())
        );
        let error = file
            .remove("/frame[1]")
            .expect_err("frames are not removable");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .add("/", litchi_ops::NewNode::new("frame"))
            .expect_err("one frame");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        // The singletons move and swap only onto themselves (record 0292).
        let same = file
            .move_to("/frame[1]", litchi_ops::Placement::Start)
            .expect("self move");
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
        let same = file.swap("/frame[1]", "/frame[1]").expect("self swap");
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
        let error = file
            .move_to(
                "/frame[1]",
                litchi_ops::Placement::Position(litchi_core::Position::new(1)),
            )
            .expect_err("no second position");
        assert_eq!(error.code(), Code::OutOfRange);
        let error = file
            .swap("/frame[1]", "/metadata")
            .expect_err("different kinds");
        assert_eq!(error.code(), Code::UnsupportedType);
        // Removing `/metadata` clears its fields and keeps the frame.
        let file = reopen(&file.remove("/metadata").expect("remove metadata"));
        assert_eq!(prop(&file, "/metadata", "title"), None);
        assert_eq!(prop(&file, "/metadata", "author"), None);
        assert_eq!(prop(&file, "/frame[1]", "name").as_deref(), Some("Logo"));
        assert_supported(
            FileFormat::Odi,
            &[
                Verb::Get,
                Verb::Query,
                Verb::Set,
                Verb::Add,
                Verb::Remove,
                Verb::Move,
                Verb::Swap,
                Verb::Save,
                Verb::Create,
            ],
        );
    }
}

#[cfg(feature = "odb")]
mod odb {
    use super::*;
    use litchi_ops::NewNode;

    const DATABASE: &str = "libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb";

    #[test]
    fn reads_tables_columns_and_queries() {
        let file = read(DATABASE);
        assert_eq!(file.format(), Some(FileFormat::Odb));
        assert_eq!(
            prop(&file, "/table[test]", "columnCount").as_deref(),
            Some("2")
        );
        assert_eq!(
            prop(&file, "/table[1]/column[2]", "name").as_deref(),
            Some("tname")
        );
        assert_eq!(file.query("query").unwrap().count(), 1);
    }

    #[test]
    fn writes_queries_tables_and_columns() {
        let file = read(DATABASE);
        let added = file
            .add(
                "/",
                NewNode::new("query")
                    .prop("name", "Recent")
                    .prop("command", "SELECT * FROM \"test\""),
            )
            .expect("add query");
        assert_eq!(added.path.to_string(), "/query[2]");
        let file = added
            .file
            .set("/query[Recent]", props! { "command" => "SELECT 1" })
            .expect("set command")
            .set("/table[1]/column[2]", props! { "name" => "label" })
            .expect("rename column")
            .set("/table[1]", props! { "name" => "items" })
            .expect("rename table");
        let file = reopen(&file);
        assert_eq!(
            prop(&file, "/query[Recent]", "command").as_deref(),
            Some("SELECT 1")
        );
        assert_eq!(prop(&file, "/table[1]", "name").as_deref(), Some("items"));
        assert_eq!(
            prop(&file, "/table[1]/column[2]", "name").as_deref(),
            Some("label")
        );
        let file = file.remove("/query[Recent]").expect("remove query");
        assert_eq!(reopen(&file).query("query").unwrap().count(), 1);
        let error = file
            .set("/query[1]", props! { "name" => "Renamed" })
            .expect_err("query rename is a gap");
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }

    #[test]
    fn moves_and_swaps_queries_tables_and_columns() {
        let mut file = create(FileFormat::Odb).expect("create");
        for index in 0..3 {
            file = file
                .add(
                    "/",
                    NewNode::new("query")
                        .prop("name", format!("q{index}").as_str())
                        .prop("command", format!("SELECT {index}").as_str()),
                )
                .expect("add query")
                .file;
        }
        let names = |file: &File| -> Vec<String> {
            (1..=3)
                .map(|index| prop(file, &format!("/query[{index}]"), "name").unwrap())
                .collect()
        };
        let file = file
            .move_to("/query[q0]", litchi_ops::Placement::End)
            .expect("move query");
        assert_eq!(names(&reopen(&file)), ["q1", "q2", "q0"]);
        let file = file
            .move_to(
                "/query[q0]",
                litchi_ops::Placement::After(litchi_ops::IntoPath::into_path("/query[1]").unwrap()),
            )
            .expect("anchored move");
        assert_eq!(names(&reopen(&file)), ["q1", "q0", "q2"]);
        let file = file.swap("/query[1]", "/query[3]").expect("swap queries");
        assert_eq!(names(&reopen(&file)), ["q2", "q0", "q1"]);
        let same = file.swap("/query[2]", "/query[2]").expect("self swap");
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());

        // A producer database: a column moves within its table.
        let file = read(DATABASE);
        let file = file
            .move_to("/table[1]/column[2]", litchi_ops::Placement::Start)
            .expect("move column");
        let file = reopen(&file);
        assert_eq!(
            prop(&file, "/table[1]/column[1]", "name").as_deref(),
            Some("tname")
        );
        let file = file
            .swap("/table[1]/column[1]", "/table[1]/column[2]")
            .expect("swap columns");
        assert_eq!(
            prop(&reopen(&file), "/table[1]/column[2]", "name").as_deref(),
            Some("tname")
        );
        let error = file
            .swap("/table[1]", "/query[1]")
            .expect_err("mixed kinds");
        assert_eq!(error.code(), Code::UnsupportedType);
    }

    #[test]
    fn query_command_batches_coalesce() {
        let mut file = create(FileFormat::Odb).expect("create");
        for index in 0..3 {
            file = file
                .add(
                    "/",
                    NewNode::new("query")
                        .prop("name", format!("q{index}").as_str())
                        .prop("command", format!("SELECT {index}").as_str()),
                )
                .expect("add query")
                .file;
        }
        let commit = file
            .edit(|tx| {
                for index in 1..=3 {
                    tx.set(
                        format!("/query[{index}]").as_str(),
                        props! { "command" => format!("SELECT {index} FROM u").as_str() },
                    )?;
                }
                Ok(())
            })
            .unwrap_or_else(|error| panic!("{error} {:?}", std::error::Error::source(&error)));
        assert_eq!(commit.diagnostics().commits(), 1);
        assert_eq!(
            prop(&reopen(commit.file()), "/query[3]", "command").as_deref(),
            Some("SELECT 3 FROM u")
        );
    }

    #[test]
    fn create_and_capabilities() {
        let file = create(FileFormat::Odb).expect("create");
        assert_eq!(reopen(&file).format(), Some(FileFormat::Odb));
        assert_supported(
            FileFormat::Odb,
            &[
                Verb::Get,
                Verb::Query,
                Verb::Set,
                Verb::Add,
                Verb::Remove,
                Verb::Move,
                Verb::Swap,
                Verb::Save,
                Verb::Create,
            ],
        );
    }
}

#[cfg(feature = "odf-formula")]
mod formula {
    use super::*;

    const FORMULA: &str = "odf/native-resave/source/font-styles.odf";

    #[test]
    fn reads_the_mathml_tree() {
        let file = read(FORMULA);
        assert_eq!(file.format(), Some(FileFormat::Odf));
        assert!(matches!(file, File::Odf(_)));
        assert!(file.query("element[tag=mi]").unwrap().count() > 0);
        assert!(prop(&file, "/", "starmath").is_some());
    }

    #[test]
    fn writes_token_text_and_starmath() {
        let file = read(FORMULA);
        let token = file
            .query("element[tag=mi]")
            .unwrap()
            .next()
            .unwrap()
            .unwrap()
            .path()
            .to_string();
        let file = file
            .set(token.as_str(), props! { "text" => "g" })
            .expect("token text")
            .set("/", props! { "starmath" => "g" })
            .expect("starmath");
        let file = reopen(&file);
        assert_eq!(text(&file, &token).as_deref(), Some("g"));
        assert_eq!(prop(&file, "/", "starmath").as_deref(), Some("g"));
        let error = file
            .set("/element[1]", props! { "text" => "x" })
            .expect_err("layout elements hold no text");
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }

    #[test]
    fn adds_moves_and_swaps_mathml_elements() {
        use litchi_ops::NewNode;
        let file = create(FileFormat::Odf).expect("create");
        let added = file
            .add("/", NewNode::new("element").prop("tag", "mrow"))
            .expect("add row");
        let row = added.path.to_string();
        let mut file = added.file;
        for (tag, content) in [("mi", "a"), ("mo", "+"), ("mi", "b")] {
            file = file
                .add(
                    row.as_str(),
                    NewNode::new("element")
                        .prop("tag", tag)
                        .prop("text", content),
                )
                .expect("add token")
                .file;
        }
        let tokens = |file: &File| -> Vec<String> {
            file.query("element[tag=mi], element[tag=mo]")
                .unwrap()
                .map(|node| node.unwrap().view().text().unwrap().unwrap().into_owned())
                .collect()
        };
        let file = reopen(&file);
        assert_eq!(tokens(&file), ["a", "+", "b"]);
        let first = format!("{row}/element[1]");
        let third = format!("{row}/element[3]");
        let file = file.swap(first.as_str(), third.as_str()).expect("swap");
        assert_eq!(tokens(&reopen(&file)), ["b", "+", "a"]);
        let file = file
            .move_to(first.as_str(), litchi_ops::Placement::End)
            .expect("move");
        assert_eq!(tokens(&reopen(&file)), ["+", "a", "b"]);
        let added = file
            .add(
                row.as_str(),
                NewNode::new("element")
                    .prop("tag", "mn")
                    .prop("text", "2")
                    .at(litchi_ops::Placement::Start),
            )
            .expect("add first");
        assert_eq!(added.path.to_string(), format!("{row}/element[1]"));
        let error = added
            .file
            .add(
                row.as_str(),
                NewNode::new("element")
                    .prop("tag", "mrow")
                    .prop("text", "x"),
            )
            .expect_err("layout text");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = added
            .file
            .move_to(
                row.as_str(),
                litchi_ops::Placement::Into(
                    litchi_ops::IntoPath::into_path(format!("{row}/element[2]").as_str()).unwrap(),
                    None,
                ),
            )
            .expect_err("into own subtree");
        assert_eq!(error.code(), Code::InvalidArgument);
    }

    #[test]
    fn create_and_capabilities() {
        let file = create(FileFormat::Odf).expect("create");
        assert_eq!(reopen(&file).format(), Some(FileFormat::Odf));
        assert_supported(
            FileFormat::Odf,
            &[
                Verb::Get,
                Verb::Query,
                Verb::Set,
                Verb::Add,
                Verb::Remove,
                Verb::Move,
                Verb::Swap,
                Verb::Save,
                Verb::Create,
            ],
        );
    }
}
