//! D7-A unified chart point-size routes.
//!
//! These cases keep the full `File` route beside the direct focused leaf
//! route.  They are intentionally source-only until the coordinator releases
//! the quiet build/performance hold.

#![cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]

use std::path::PathBuf;

use litchi_ops::{File, PropValue, open_bytes, props};

fn fixture(format: &str) -> PathBuf {
    let extension = match format {
        "pages" => "pages",
        "keynote" => "key",
        "numbers" => "numbers",
        _ => panic!("unsupported iWork fixture format"),
    };
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(format)
        .join(format!("chart-grid-authored-native.{extension}"))
}

fn read(path: PathBuf) -> File {
    open_bytes(std::fs::read(path).expect("read chart fixture")).expect("open chart fixture")
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("serialize chart package")).expect("reopen chart package")
}

fn named_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .expect("chart path")
        .view()
        .format()
        .expect("chart format")
        .find(|property| property.key() == key)
        .map(|property| property.to_string())
}

fn property(file: &File, path: &str) -> Option<String> {
    named_property(file, path, "font.size")
}

fn format_values(file: &File, path: &str) -> Vec<(String, String)> {
    file.get(path)
        .expect("chart path")
        .view()
        .format()
        .expect("chart format")
        .map(|property| (property.key().to_owned(), property.to_string()))
        .collect()
}

fn points(value: Option<String>) -> f32 {
    value
        .expect("font.size property")
        .strip_suffix("pt")
        .expect("point-size unit")
        .parse()
        .expect("point-size value")
}

#[cfg(feature = "pages")]
mod pages {
    use super::*;

    const PATH: &str = "/body/chart[1]";

    fn source() -> File {
        read(fixture("pages"))
    }

    fn original_mixed_source() -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/pages/chart-data-native.pages"),
        )
        .expect("read original Pages chart fixture")
    }

    #[inline(never)]
    fn direct_set(source: &[u8]) -> Vec<u8> {
        let package = litchi_pages::Package::from_bytes(source).expect("Pages leaf source");
        let commit = package
            .edit_body_chart_font_size(litchi_pages::BodyChartSelector::index(0))
            .expect("Pages chart")
            .set_size(litchi_pages::chart::font::TextPointSize::from_points(17.0).unwrap())
            .expect("Pages size")
            .commit()
            .expect("Pages leaf set");
        let mut bytes = Vec::new();
        commit
            .package()
            .write_to(&mut bytes)
            .expect("Pages leaf bytes");
        bytes
    }

    #[inline(never)]
    fn direct_reset(source: &[u8]) -> Vec<u8> {
        let package = litchi_pages::Package::from_bytes(source).expect("Pages leaf source");
        let commit = package
            .edit_body_chart_font_size(litchi_pages::BodyChartSelector::index(0))
            .expect("Pages chart")
            .reset_size()
            .expect("Pages reset")
            .commit()
            .expect("Pages leaf reset");
        let mut bytes = Vec::new();
        commit
            .package()
            .write_to(&mut bytes)
            .expect("Pages leaf bytes");
        bytes
    }

    #[inline(never)]
    fn unified_set_route(source: &[u8]) -> Vec<u8> {
        open_bytes(source.to_vec())
            .expect("Pages unified open")
            .set(PATH, props! { "font.size" => "17pt" })
            .expect("Pages unified set")
            .to_bytes()
            .expect("Pages unified bytes")
    }

    #[inline(never)]
    fn unified_reset_route(source: &[u8]) -> Vec<u8> {
        open_bytes(source.to_vec())
            .expect("Pages unified open")
            .set(PATH, props! { "font.size" => PropValue::Null })
            .expect("Pages unified reset")
            .to_bytes()
            .expect("Pages unified bytes")
    }

    #[test]
    fn chart_font_size_read_set_reset_and_exact_source_controls() {
        let file = source();
        let original = file.to_bytes().expect("Pages source bytes");
        let before = points(property(&file, PATH));
        assert_eq!(file.to_bytes().expect("read preserves source"), original);
        let noop = file
            .set(
                PATH,
                props! { "font.size" => property(&file, PATH).expect("baseline size") },
            )
            .expect("Pages exact no-op");
        assert_eq!(noop.to_bytes().expect("Pages no-op bytes"), original);

        let direct = direct_set(&original);
        let unified = file
            .set(PATH, props! { "font.size" => "17pt" })
            .expect("Pages unified set");
        assert_eq!(points(property(&reopen(&unified), PATH)), 17.0);
        assert_eq!(
            points(property(&reopen(&unified), PATH)),
            points(property(
                &reopen(&open_bytes(direct.clone()).unwrap()),
                PATH
            ))
        );

        let direct_restored = direct_reset(&direct);
        let restored = reopen(&unified)
            .set(PATH, props! { "font.size" => PropValue::Null })
            .expect("Pages unified reset");
        assert_eq!(points(property(&reopen(&restored), PATH)), before);
        assert_eq!(reopen(&restored).to_bytes().unwrap(), direct_restored);
        let commit = file
            .edit(|tx| tx.set(PATH, props! { "font.size" => "17pt" }))
            .expect("Pages transaction set");
        let inverse = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("Pages inverse");
        assert_eq!(inverse.to_bytes().expect("Pages inverse bytes"), original);
        let stale = commit
            .patch()
            .apply(&unified)
            .expect_err("stale Pages chart patch");
        assert_eq!(stale.code(), litchi_ops::Code::InvalidInput);
        let failed = file
            .edit(|tx| {
                tx.set(PATH, props! { "font.size" => "17pt" })?;
                tx.set(PATH, props! { "font.size" => "0pt" })
            })
            .expect_err("failed Pages chart batch is atomic");
        assert_eq!(failed.code(), litchi_ops::Code::InvalidValue);
        assert_eq!(file.to_bytes().expect("Pages batch source"), original);

        let invalid = source()
            .set(PATH, props! { "font.size" => "0pt" })
            .expect_err("zero size must refuse");
        assert_eq!(invalid.code(), litchi_ops::Code::InvalidValue);

        let file = source();
        let original = file.to_bytes().expect("Pages source bytes");
        let extreme = file
            .set(
                PATH,
                props! {
                    "font.size" => PropValue::from(litchi_ops::value::Length::from_emu(i64::MAX)),
                },
            )
            .expect_err("unrepresentable EMU size must refuse before mutation");
        assert!(matches!(
            extreme.code(),
            litchi_ops::Code::InvalidValue | litchi_ops::Code::LimitExceeded
        ));
        assert_eq!(file.to_bytes().expect("Pages source bytes"), original);
    }

    #[test]
    fn original_mixed_chart_fixture_keeps_typed_leaf_refusal() {
        let bytes = original_mixed_source();
        let package = litchi_pages::Package::from_bytes(&bytes).expect("Pages leaf source");
        assert!(matches!(
            package.body_chart_font_size(litchi_pages::BodyChartSelector::index(0)),
            Err(litchi_pages::chart::font::Error::InvalidSource)
        ));

        let file = open_bytes(bytes).expect("Pages unified source");
        let error = file
            .get(PATH)
            .expect("Pages chart path")
            .view()
            .format()
            .expect_err("unified chart projection must retain the typed refusal");
        assert_eq!(error.code(), litchi_ops::Code::InvalidInput);
    }

    #[test]
    fn canonical_chart_path_matches_generic_projection_and_missing_behavior() {
        let file = source();
        assert_eq!(
            format_values(&file, PATH),
            format_values(&file, "/body/chart[last()]")
        );
        let fast = file.get(PATH).expect("canonical chart path");
        let generic = file.get("/body/chart[last()]").expect("generic chart path");
        assert_eq!(fast.path().to_string(), generic.path().to_string());
        assert_eq!(
            file.get("/body/chart[999]").unwrap_err().code(),
            litchi_ops::Code::PathNotFound
        );
    }

    #[inline(never)]
    fn unified_read_route(source: &[u8]) -> f32 {
        points(property(
            &open_bytes(source.to_vec()).expect("Pages unified open"),
            PATH,
        ))
    }

    #[inline(never)]
    fn direct_read_route(source: &[u8]) -> f32 {
        litchi_pages::Package::from_bytes(source)
            .expect("Pages leaf open")
            .body_chart_font_size(litchi_pages::BodyChartSelector::index(0))
            .expect("Pages leaf read")
            .points()
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn pages_chart_font_size_read_full_file_cost_abba() {
        let source = source().to_bytes().expect("Pages probe source");
        let source = std::hint::black_box(source);
        std::hint::black_box(unified_read_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_read_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_read_route(std::hint::black_box(&source)));
        std::hint::black_box(unified_read_route(std::hint::black_box(&source)));
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn pages_chart_font_size_set_full_file_cost_abba() {
        let source = source().to_bytes().expect("Pages probe source");
        let source = std::hint::black_box(source);
        std::hint::black_box(unified_set_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_set(std::hint::black_box(&source)));
        std::hint::black_box(direct_set(std::hint::black_box(&source)));
        std::hint::black_box(unified_set_route(std::hint::black_box(&source)));
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn pages_chart_font_size_reset_full_file_cost_abba() {
        let source = std::hint::black_box(direct_set(std::hint::black_box(
            &source().to_bytes().expect("Pages probe source"),
        )));
        std::hint::black_box(unified_reset_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_reset(std::hint::black_box(&source)));
        std::hint::black_box(direct_reset(std::hint::black_box(&source)));
        std::hint::black_box(unified_reset_route(std::hint::black_box(&source)));
    }
}

#[cfg(feature = "keynote")]
mod keynote {
    use super::*;

    const PATH: &str = "/slide[1]/chart[1]";

    fn source() -> File {
        read(fixture("keynote"))
    }

    #[inline(never)]
    fn direct_set(source: &[u8]) -> Vec<u8> {
        let package = litchi_keynote::Package::from_bytes(source).expect("Keynote leaf source");
        let commit = package
            .edit_slide_chart_font_size(0usize, 0usize)
            .expect("Keynote chart")
            .set(litchi_keynote::chart::font::size::TextPointSize::from_points(17.0).unwrap())
            .commit()
            .expect("Keynote leaf set");
        let mut bytes = Vec::new();
        commit
            .package()
            .write_to(&mut bytes)
            .expect("Keynote leaf bytes");
        bytes
    }

    #[inline(never)]
    fn direct_reset(source: &[u8]) -> Vec<u8> {
        let package = litchi_keynote::Package::from_bytes(source).expect("Keynote leaf source");
        let commit = package
            .edit_slide_chart_font_size(0usize, 0usize)
            .expect("Keynote chart")
            .reset()
            .commit()
            .expect("Keynote leaf reset");
        let mut bytes = Vec::new();
        commit
            .package()
            .write_to(&mut bytes)
            .expect("Keynote leaf bytes");
        bytes
    }

    #[inline(never)]
    fn unified_set_route(source: &[u8]) -> Vec<u8> {
        open_bytes(source.to_vec())
            .expect("Keynote unified open")
            .set(PATH, props! { "font.size" => "17pt" })
            .expect("Keynote unified set")
            .to_bytes()
            .expect("Keynote unified bytes")
    }

    #[inline(never)]
    fn unified_reset_route(source: &[u8]) -> Vec<u8> {
        open_bytes(source.to_vec())
            .expect("Keynote unified open")
            .set(PATH, props! { "font.size" => PropValue::Null })
            .expect("Keynote unified reset")
            .to_bytes()
            .expect("Keynote unified bytes")
    }

    #[test]
    fn chart_font_size_read_set_reset_and_exact_source_controls() {
        let file = source();
        let original = file.to_bytes().expect("Keynote source bytes");
        let before = points(property(&file, PATH));
        assert_eq!(file.to_bytes().expect("read preserves source"), original);
        let frame_before = (
            named_property(&file, PATH, "name"),
            named_property(&file, PATH, "x"),
            named_property(&file, PATH, "zorder"),
        );
        let _ = property(&file, PATH);
        assert_eq!(
            (
                named_property(&file, PATH, "name"),
                named_property(&file, PATH, "x"),
                named_property(&file, PATH, "zorder"),
            ),
            frame_before,
            "lazy font read must preserve Keynote frame/name/zorder values",
        );
        let noop = file
            .set(
                PATH,
                props! { "font.size" => property(&file, PATH).expect("baseline size") },
            )
            .expect("Keynote exact no-op");
        assert_eq!(noop.to_bytes().expect("Keynote no-op bytes"), original);

        let direct = direct_set(&original);
        let unified = file
            .set(PATH, props! { "font.size" => "17pt" })
            .expect("Keynote unified set");
        assert_eq!(points(property(&reopen(&unified), PATH)), 17.0);
        assert_eq!(
            points(property(&reopen(&unified), PATH)),
            points(property(
                &reopen(&open_bytes(direct.clone()).unwrap()),
                PATH
            ))
        );

        let direct_restored = direct_reset(&direct);
        let restored = reopen(&unified)
            .set(PATH, props! { "font.size" => PropValue::Null })
            .expect("Keynote unified reset");
        assert_eq!(points(property(&reopen(&restored), PATH)), before);
        assert_eq!(reopen(&restored).to_bytes().unwrap(), direct_restored);
        let commit = file
            .edit(|tx| tx.set(PATH, props! { "font.size" => "17pt" }))
            .expect("Keynote transaction set");
        let inverse = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("Keynote inverse");
        assert_eq!(inverse.to_bytes().expect("Keynote inverse bytes"), original);
        let stale = commit
            .patch()
            .apply(&unified)
            .expect_err("stale Keynote chart patch");
        assert_eq!(stale.code(), litchi_ops::Code::InvalidInput);
        let failed = file
            .edit(|tx| {
                tx.set(PATH, props! { "font.size" => "17pt" })?;
                tx.set(PATH, props! { "font.size" => "0pt" })
            })
            .expect_err("failed Keynote chart batch is atomic");
        assert_eq!(failed.code(), litchi_ops::Code::InvalidValue);
        assert_eq!(file.to_bytes().expect("Keynote batch source"), original);

        let invalid = source()
            .set(PATH, props! { "font.size" => "0pt" })
            .expect_err("zero size must refuse");
        assert_eq!(invalid.code(), litchi_ops::Code::InvalidValue);
    }

    #[test]
    fn chart_font_size_queries_use_typed_property_reads() {
        let changed = source()
            .set(PATH, props! { "font.size" => "17pt" })
            .expect("seed Keynote chart size");
        let single = changed
            .query("chart[font.size=17pt]")
            .expect("single chart-size query")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("single chart-size query read");
        assert_eq!(single.len(), 1);
        let compound = changed
            .query("chart[font.size=17pt][font.size!=18pt]")
            .expect("compound chart-size query")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("compound chart-size query read");
        assert_eq!(compound.len(), 1);

        let refused = read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/chart-data-native.key"),
        );
        let frame_only = refused
            .query("chart[zorder>=0]")
            .expect("frame-only query")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("frame-only query must not read font.size");
        assert_eq!(frame_only.len(), 1);
        let font_error = refused
            .query("chart[font.size=17pt]")
            .expect("font query plan")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect_err("font query retains the focused refusal");
        assert_eq!(font_error.code(), litchi_ops::Code::InvalidInput);
    }

    #[test]
    fn canonical_chart_path_matches_generic_projection_and_missing_behavior() {
        let file = source();
        assert_eq!(
            format_values(&file, PATH),
            format_values(&file, "/slide[1]/chart[last()]")
        );
        let fast = file.get(PATH).expect("canonical chart path");
        let generic = file
            .get("/slide[1]/chart[last()]")
            .expect("generic chart path");
        assert_eq!(fast.path().to_string(), generic.path().to_string());

        let refused = read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/chart-data-native.key"),
        );
        assert_eq!(
            refused
                .get(PATH)
                .expect("fast chart path")
                .view()
                .format()
                .unwrap_err()
                .code(),
            refused
                .get("/slide[1]/chart[last()]")
                .expect("generic chart path")
                .view()
                .format()
                .unwrap_err()
                .code()
        );
        assert_eq!(
            file.get("/slide[1]/chart[999]").unwrap_err().code(),
            litchi_ops::Code::PathNotFound
        );
    }

    #[inline(never)]
    fn unified_read_route(source: &[u8]) -> f32 {
        points(property(
            &open_bytes(source.to_vec()).expect("Keynote unified open"),
            PATH,
        ))
    }

    #[inline(never)]
    fn direct_read_route(source: &[u8]) -> f32 {
        litchi_keynote::Package::from_bytes(source)
            .expect("Keynote leaf open")
            .slide_chart_font_size(0usize, 0usize)
            .expect("Keynote leaf read")
            .points()
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn keynote_chart_font_size_read_full_file_cost_abba() {
        let source = source().to_bytes().expect("Keynote probe source");
        let source = std::hint::black_box(source);
        std::hint::black_box(unified_read_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_read_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_read_route(std::hint::black_box(&source)));
        std::hint::black_box(unified_read_route(std::hint::black_box(&source)));
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn keynote_chart_font_size_set_full_file_cost_abba() {
        let source = source().to_bytes().expect("Keynote probe source");
        let source = std::hint::black_box(source);
        std::hint::black_box(unified_set_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_set(std::hint::black_box(&source)));
        std::hint::black_box(direct_set(std::hint::black_box(&source)));
        std::hint::black_box(unified_set_route(std::hint::black_box(&source)));
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn keynote_chart_font_size_reset_full_file_cost_abba() {
        let source = std::hint::black_box(direct_set(std::hint::black_box(
            &source().to_bytes().expect("Keynote probe source"),
        )));
        std::hint::black_box(unified_reset_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_reset(std::hint::black_box(&source)));
        std::hint::black_box(direct_reset(std::hint::black_box(&source)));
        std::hint::black_box(unified_reset_route(std::hint::black_box(&source)));
    }
}

#[cfg(feature = "numbers")]
mod numbers {
    use super::*;

    const PATH: &str = "/sheet[1]/chart[1]";

    fn source() -> File {
        read(fixture("numbers"))
    }

    #[inline(never)]
    fn direct_set(source: &[u8]) -> Vec<u8> {
        let package = litchi_numbers::Package::from_bytes(source).expect("Numbers leaf source");
        let mut edit = package
            .edit_sheet_chart_font(0usize, 0usize)
            .expect("Numbers chart");
        edit.set(litchi_numbers::sheet::chart::font::Property::Size(
            litchi_numbers::sheet::chart::font::TextPointSize::from_points(17.0).unwrap(),
        ))
        .expect("Numbers size");
        let commit = edit.commit().expect("Numbers leaf set");
        let mut bytes = Vec::new();
        commit
            .package()
            .write_to(&mut bytes)
            .expect("Numbers leaf bytes");
        bytes
    }

    #[inline(never)]
    fn direct_reset(source: &[u8]) -> Vec<u8> {
        let package = litchi_numbers::Package::from_bytes(source).expect("Numbers leaf source");
        let mut edit = package
            .edit_sheet_chart_font(0usize, 0usize)
            .expect("Numbers chart");
        edit.reset(litchi_numbers::sheet::chart::font::PropertyKind::Size)
            .expect("Numbers reset");
        let commit = edit.commit().expect("Numbers leaf reset");
        let mut bytes = Vec::new();
        commit
            .package()
            .write_to(&mut bytes)
            .expect("Numbers leaf bytes");
        bytes
    }

    #[inline(never)]
    fn unified_set_route(source: &[u8]) -> Vec<u8> {
        open_bytes(source.to_vec())
            .expect("Numbers unified open")
            .set(PATH, props! { "font.size" => "17pt" })
            .expect("Numbers unified set")
            .to_bytes()
            .expect("Numbers unified bytes")
    }

    #[inline(never)]
    fn unified_reset_route(source: &[u8]) -> Vec<u8> {
        open_bytes(source.to_vec())
            .expect("Numbers unified open")
            .set(PATH, props! { "font.size" => PropValue::Null })
            .expect("Numbers unified reset")
            .to_bytes()
            .expect("Numbers unified bytes")
    }

    #[test]
    fn chart_font_size_read_set_reset_and_exact_source_controls() {
        let file = source();
        let original = file.to_bytes().expect("Numbers source bytes");
        let before = points(property(&file, PATH));
        assert_eq!(file.to_bytes().expect("read preserves source"), original);
        let noop = file
            .set(
                PATH,
                props! { "font.size" => property(&file, PATH).expect("baseline size") },
            )
            .expect("Numbers exact no-op");
        assert_eq!(noop.to_bytes().expect("Numbers no-op bytes"), original);

        let direct = direct_set(&original);
        let unified = file
            .set(PATH, props! { "font.size" => "17pt" })
            .expect("Numbers unified set");
        assert_eq!(points(property(&reopen(&unified), PATH)), 17.0);
        assert_eq!(
            points(property(&reopen(&unified), PATH)),
            points(property(
                &reopen(&open_bytes(direct.clone()).unwrap()),
                PATH
            ))
        );

        let direct_restored = direct_reset(&direct);
        let restored = reopen(&unified)
            .set(PATH, props! { "font.size" => PropValue::Null })
            .expect("Numbers unified reset");
        assert_eq!(points(property(&reopen(&restored), PATH)), before);
        assert_eq!(reopen(&restored).to_bytes().unwrap(), direct_restored);
        let commit = file
            .edit(|tx| tx.set(PATH, props! { "font.size" => "17pt" }))
            .expect("Numbers transaction set");
        let inverse = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("Numbers inverse");
        assert_eq!(inverse.to_bytes().expect("Numbers inverse bytes"), original);
        let stale = commit
            .patch()
            .apply(&unified)
            .expect_err("stale Numbers chart patch");
        assert_eq!(stale.code(), litchi_ops::Code::InvalidInput);
        let failed = file
            .edit(|tx| {
                tx.set(PATH, props! { "font.size" => "17pt" })?;
                tx.set(PATH, props! { "font.size" => "0pt" })
            })
            .expect_err("failed Numbers chart batch is atomic");
        assert_eq!(failed.code(), litchi_ops::Code::InvalidValue);
        assert_eq!(file.to_bytes().expect("Numbers batch source"), original);

        let invalid = source()
            .set(PATH, props! { "font.size" => "0pt" })
            .expect_err("zero size must refuse");
        assert_eq!(invalid.code(), litchi_ops::Code::InvalidValue);
    }

    #[test]
    fn canonical_chart_path_matches_generic_projection_and_missing_behavior() {
        let file = source();
        assert_eq!(
            format_values(&file, PATH),
            format_values(&file, "/sheet[1]/chart[last()]")
        );
        let fast = file.get(PATH).expect("canonical chart path");
        let generic = file
            .get("/sheet[1]/chart[last()]")
            .expect("generic chart path");
        assert_eq!(fast.path().to_string(), generic.path().to_string());
        assert_eq!(
            file.get("/sheet[1]/chart[999]").unwrap_err().code(),
            litchi_ops::Code::PathNotFound
        );
    }

    #[inline(never)]
    fn unified_read_route(source: &[u8]) -> f32 {
        points(property(
            &open_bytes(source.to_vec()).expect("Numbers unified open"),
            PATH,
        ))
    }

    #[inline(never)]
    fn direct_read_route(source: &[u8]) -> f32 {
        match litchi_numbers::Package::from_bytes(source)
            .expect("Numbers leaf open")
            .sheet_chart_font(
                0usize,
                0usize,
                litchi_numbers::sheet::chart::font::PropertyKind::Size,
            )
            .expect("Numbers leaf read")
        {
            litchi_numbers::sheet::chart::font::Property::Size(size) => size.points(),
            litchi_numbers::sheet::chart::font::Property::Font(_) => {
                panic!("Numbers chart font read returned font for size route")
            },
            _ => panic!("Numbers chart font read returned an unknown property"),
        }
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn numbers_chart_font_size_read_full_file_cost_abba() {
        let source = source().to_bytes().expect("Numbers probe source");
        let source = std::hint::black_box(source);
        std::hint::black_box(unified_read_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_read_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_read_route(std::hint::black_box(&source)));
        std::hint::black_box(unified_read_route(std::hint::black_box(&source)));
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn numbers_chart_font_size_set_full_file_cost_abba() {
        let source = source().to_bytes().expect("Numbers probe source");
        let source = std::hint::black_box(source);
        std::hint::black_box(unified_set_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_set(std::hint::black_box(&source)));
        std::hint::black_box(direct_set(std::hint::black_box(&source)));
        std::hint::black_box(unified_set_route(std::hint::black_box(&source)));
    }

    #[test]
    #[ignore = "coordinator CPU slot: full-file ABBA qualification"]
    fn numbers_chart_font_size_reset_full_file_cost_abba() {
        let source = std::hint::black_box(direct_set(std::hint::black_box(
            &source().to_bytes().expect("Numbers probe source"),
        )));
        std::hint::black_box(unified_reset_route(std::hint::black_box(&source)));
        std::hint::black_box(direct_reset(std::hint::black_box(&source)));
        std::hint::black_box(direct_reset(std::hint::black_box(&source)));
        std::hint::black_box(unified_reset_route(std::hint::black_box(&source)));
    }
}
