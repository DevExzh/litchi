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
        // Structural writes on a LibreOffice-authored drawing are refused by
        // the leaf's compact-XML audit (a gap); they run on a fresh drawing.
        let error = file
            .add(
                "/page[1]",
                NewNode::new("shape")
                    .prop("x", "1cm")
                    .prop("y", "1cm")
                    .prop("width", "4cm")
                    .prop("height", "2cm"),
            )
            .expect_err("producer whitespace");
        assert_ne!(error.code(), Code::InternalError);
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
            .remove("/page[1]")
            .expect_err("pages are not removable");
        assert_eq!(error.code(), Code::UnsupportedCapability);
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
                Verb::Save,
                Verb::Create,
            ],
        );
    }
}

#[cfg(feature = "odc")]
mod odc {
    use super::*;

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
        // typed surface, so a structural removal is a typed refusal.
        let error = file.remove("/series[1]").expect_err("lossy typed edit");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let error = file
            .set("/", props! { "title" => "x" })
            .expect_err("chart title is a gap");
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
        // no `svg` namespace declaration: inserting geometry is refused by
        // the leaf's readback (a gap), as a typed caller error.
        let error = file
            .set("/frame[1]", props! { "width" => "5cm" })
            .expect_err("no svg namespace");
        assert_eq!(error.code(), Code::InvalidValue);
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
        assert_supported(
            FileFormat::Odi,
            &[Verb::Get, Verb::Query, Verb::Set, Verb::Save, Verb::Create],
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
    fn create_and_capabilities() {
        let file = create(FileFormat::Odf).expect("create");
        assert_eq!(reopen(&file).format(), Some(FileFormat::Odf));
        assert_supported(
            FileFormat::Odf,
            &[
                Verb::Get,
                Verb::Query,
                Verb::Set,
                Verb::Remove,
                Verb::Save,
                Verb::Create,
            ],
        );
    }
}
