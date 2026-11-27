//! Native Excel chart substreams through the shared inert BIFF model.
use litchi_ograph::Limits;
use litchi_ograph::chart::{
    Cache, Context, Family, Model, Owner, Refs, Role, Source, XlValue, cache,
};
use litchi_ograph::record::of_pie;

fn workbook() -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel/excel-chart-families.xls");
    let mut ole = litchi_cfb::OleFile::open(
        std::fs::File::open(path).unwrap(),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    ole.open_stream(&["Workbook"]).unwrap()
}
fn family(group: &Family) -> &'static str {
    match group {
        Family::Line { .. } => "line",
        Family::Area { .. } => "area",
        Family::Bar { .. } => "bar",
        Family::Pie { hole: 0, .. } => "pie",
        Family::Pie { .. } => "doughnut",
        Family::OfPie { .. } => "ofPie",
        Family::Radar { .. } => "radar",
        Family::Surface { .. } => "surface",
        Family::Scatter { flags, .. } if flags & 1 != 0 => "bubble",
        Family::Scatter { .. } => "scatter",
    }
}
#[test]
fn native_seventeen_biff_charts_keep_families_series_owners_and_exact_replay() {
    let bytes = workbook();
    let charts = Refs::open_workbook(&bytes, Limits::default())
        .unwrap()
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    assert_eq!(charts.len(), 17);
    // BIFF's 3-D variants share their family's record, while stock uses Line
    // with CrtLine(HighLow). This asserts wire semantics, not rendered appearance.
    let families: [&[&str]; 17] = [
        &["line"],
        &["area"],
        &["pie"],
        &["radar"],
        &["surface"],
        &["scatter"],
        &["bubble"],
        &["bar"],
        &["bar"],
        &["area"],
        &["line"],
        &["surface"],
        &["doughnut"],
        &["ofPie"],
        &["pie"],
        &["line"],
        &["bar", "line"],
    ];
    for (i, source) in charts.into_iter().enumerate() {
        let model = Model::parse(source, Context::excel()).unwrap();
        assert_eq!(
            model
                .groups()
                .iter()
                .map(|g| family(&g.family))
                .collect::<Vec<_>>(),
            families[i]
        );
        assert_eq!(
            model.axes().len(),
            [2, 2, 0, 2, 3, 2, 2, 2, 2, 3, 3, 3, 0, 0, 0, 2, 4][i]
        );
        let regular = model
            .series()
            .iter()
            .filter(|s| matches!(s.owner, Owner::Group(_)))
            .collect::<Vec<_>>();
        assert_eq!(
            regular.len(),
            if matches!(i, 2 | 5 | 6 | 12..=14) {
                1
            } else {
                3
            }
        );
        for (n, s) in regular.iter().enumerate() {
            assert_eq!((s.category_count.get(), s.value_count.get()), (5, 5));
            assert_eq!(s.ai.get(Role::Values).link().source(), Source::Cells);
            if !matches!(i, 5 | 6 | 15) {
                assert_eq!(
                    s.ai.get(Role::Name).text(),
                    Some(["SERIES_A_013", "SERIES_B_013", "SERIES_C_013"][n])
                );
            }
        }
        let errors = model
            .series()
            .iter()
            .filter_map(|s| match &s.owner {
                Owner::ErrorBar { parent, .. } => Some(parent.series().get()),
                _ => None,
            })
            .collect::<Vec<_>>();
        let expected_errors: &[u32] = match i {
            0 => &[1, 1, 2, 2, 3, 3],
            5 => &[1, 1, 1, 1],
            _ => &[],
        };
        assert_eq!(errors, expected_errors);
        let trends = model
            .series()
            .iter()
            .filter_map(|s| match &s.owner {
                Owner::Trend { parent, .. } => Some(parent.series().get()),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(trends, if matches!(i, 0 | 5) { vec![1] } else { vec![] });
        assert_eq!(
            model.series().len(),
            regular.len() + errors.len() + trends.len()
        );
        if i == 13 {
            let Family::OfPie { options } = model.groups()[0].family else {
                unreachable!()
            };
            assert_eq!(options.kind(), of_pie::Kind::Pie);
            assert_eq!(options.split(), of_pie::Split::Automatic);
            assert_eq!(
                (
                    options.secondary_size(),
                    options.gap(),
                    options.has_shadow()
                ),
                (75, 150, false)
            );
        }
        if i == 16 {
            assert_eq!(
                regular
                    .iter()
                    .map(|s| s.owner.group().unwrap().get())
                    .collect::<Vec<_>>(),
                [0, 0, 1]
            );
            assert_eq!(model.parents().len(), 2);
        }
        assert!(model.is_pristine());
        assert!(model.encode().unwrap().as_bytes() == source.as_bytes());
    }
}

#[test]
fn native_scatter_error_cache_values_are_bound_to_auxiliary_series() {
    let bytes = workbook();
    let source = Refs::open_workbook(&bytes, Limits::default())
        .unwrap()
        .nth(5)
        .unwrap()
        .unwrap();
    let model = Model::parse(source, Context::excel()).unwrap();
    assert_eq!(model.caches().len(), 10);
    let mut values = Vec::new();
    for cell in model.caches() {
        let Cache::Excel {
            section,
            row,
            col,
            value,
            ..
        } = cell
        else {
            panic!("wrong producer cache")
        };
        assert_eq!(*section, cache::Index::Values);
        let XlValue::Number(value) = value else {
            panic!("numeric custom error expected")
        };
        values.push((*row, *col, *value));
    }
    values.sort_by_key(|(row, col, _)| (*row, *col));
    assert_eq!(
        values,
        [
            (0, 1, 0.5),
            (0, 2, 0.25),
            (1, 1, 1.),
            (1, 2, 0.5),
            (2, 1, 1.5),
            (2, 2, 0.75),
            (3, 1, 2.),
            (3, 2, 1.),
            (4, 1, 2.5),
            (4, 2, 1.25)
        ]
    );
    for (i, s) in model.series()[1..5].iter().enumerate() {
        let Owner::ErrorBar { parent, data } = &s.owner else {
            panic!("auxiliary error-bar owner missing")
        };
        assert_eq!(parent.series().get(), 1);
        assert_eq!(data.len(), 14);
        assert_eq!((data[0], data[1]), [(3, 4), (4, 4), (1, 2), (2, 2)][i]);
    }
    let Owner::Trend { parent, data } = &model.series()[5].owner else {
        panic!("trendline owner missing")
    };
    assert_eq!(parent.series().get(), 1);
    assert_eq!((data.len(), data[0], data[10], data[11]), (28, 0, 1, 1));
}

#[test]
fn native_of_pie_full_model_reauthoring_has_a_typed_preservation_refusal() {
    let bytes = workbook();
    let source = Refs::open_workbook(&bytes, Limits::default())
        .unwrap()
        .nth(13)
        .unwrap()
        .unwrap();
    let mut model = Model::parse(source, Context::excel()).unwrap();
    model.groups_mut()[0].family = Family::OfPie {
        options: of_pie::Options::new(
            of_pie::Kind::Pie,
            of_pie::Split::Position(2),
            75,
            150,
            false,
        )
        .unwrap(),
    };
    // Full model reauthoring cannot place the source's opaque records. The
    // source-preserving fixed-record transactions remain available separately.
    let error = model.encode().unwrap_err();
    assert!(matches!(error, litchi_ograph::Error::Refused(ref refusal)
        if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent));
}
