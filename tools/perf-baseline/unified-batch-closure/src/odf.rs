//! ODF batches: ODT and the ODT-engine text formats (OTH, ODM), ODS, ODP,
//! and the object formats (ODG, ODC, ODI, ODB, ODF formula).

#[allow(unused_imports)]
use crate::{Case, Result, check_pair, set, timed, unified_leg};
#[allow(unused_imports)]
use litchi_core::Position;
#[allow(unused_imports)]
use litchi_ops::{BatchItem, File};

pub const BUILDERS: &[(&str, fn(&[usize]) -> Result<Vec<Case<'static>>>)] = &[
    #[cfg(feature = "odt")]
    ("odt", odt),
    #[cfg(feature = "oth")]
    ("oth", oth),
    #[cfg(feature = "odm")]
    ("odm", odm),
    #[cfg(feature = "ods")]
    ("ods", ods),
    #[cfg(feature = "odp")]
    ("odp", odp),
    #[cfg(feature = "odg")]
    ("odg", odg),
    #[cfg(feature = "odc")]
    ("odc", odc),
    #[cfg(feature = "odi")]
    ("odi", odi),
    #[cfg(feature = "odb")]
    ("odb", odb),
    #[cfg(feature = "odf-formula")]
    ("odf", formula),
];

fn leak<T>(value: T) -> &'static T {
    Box::leak(Box::new(value))
}

#[allow(dead_code)]
fn open(bytes: Vec<u8>) -> Result<&'static File> {
    Ok(leak(litchi_ops::open_bytes(bytes)?))
}

/// The ODT body engine's paragraph text sets, shared by ODT, OTH and ODM
/// (`File::Oth`/`File::Odm` hold the ODT handle, record 0282): one leaf
/// `Edit::set_text` per paragraph against `set /body/p[i] text`.
#[cfg(any(feature = "odt", feature = "oth", feature = "odm"))]
fn odt_text_cases(
    format: &'static str,
    bytes: Vec<u8>,
    counts: &[usize],
) -> Result<Vec<Case<'static>>> {
    let native = leak(litchi_odt::Package::from_bytes(&bytes)?);
    let ids = {
        let document = native.document()?;
        let body = document.body()?;
        body.root()
            .children()
            .filter(|n| n.kind() == litchi_odt::body::Kind::Paragraph)
            .map(|n| n.id())
            .collect::<Vec<_>>()
    };
    let ids = leak(ids);
    let file = open(bytes)?;
    fn leaf(
        native: &litchi_odt::Package,
        ids: &[litchi_odt::body::NodeId],
        texts: &[String],
    ) -> Result<litchi_odt::Package> {
        let mut edit = native.edit();
        for (id, text) in ids.iter().zip(texts) {
            edit.set_text(*id, text.as_str())?;
        }
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        if ids.len() < n {
            return Err(
                format!("{format}: fixture has {} paragraphs, {n} wanted", ids.len()).into(),
            );
        }
        let texts = leak((0..n).map(|i| format!("set {i}")).collect::<Vec<_>>());
        let items = leak(
            texts
                .iter()
                .enumerate()
                .map(|(i, t)| set(format!("/body/p[{}]", i + 1), "text", t.clone()))
                .collect::<Vec<_>>(),
        );
        let ids = &ids[..n];
        cases.push(Case {
            format,
            scenario: "paragraph_text_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, ids, texts)?.to_bytes()?, 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, ids, texts))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

#[cfg(feature = "odt")]
fn odt(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut builder = litchi_odt::Builder::new();
    for i in 0..max {
        builder.add_paragraph(&format!(
            "Paragraph {i} holds a few ordinary words of body text."
        ))?;
    }
    let bytes = builder.build()?.to_bytes()?;
    odt_text_cases("odt", bytes, counts)
}

/// OTH (HTML template): the leaf's own builder writes the fixture; the leaf
/// leg is the ODT engine the unified handle holds.
#[cfg(feature = "oth")]
fn oth(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut builder = litchi_oth::Builder::new();
    for i in 0..max {
        builder = builder.paragraph(litchi_oth::paragraph::Paragraph::new(format!(
            "Paragraph {i} holds a few ordinary words of body text."
        )));
    }
    let bytes = builder.build()?.to_bytes()?;
    odt_text_cases("oth", bytes, counts)
}

/// ODM (master document): the leaf builder with the whole content part
/// (`Builder::body_item` re-appends the part per item).
#[cfg(feature = "odm")]
fn odm(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:text>"#,
    );
    for i in 0..max {
        xml.push_str(&format!(
            "<text:p>Paragraph {i} holds a few ordinary words of body text.</text:p>"
        ));
    }
    xml.push_str("</office:text></office:body></office:document-content>");
    let bytes = litchi_odm::Builder::new()
        .content_xml(xml)
        .build()?
        .as_bytes()
        .to_vec();
    odt_text_cases("odm", bytes, counts)
}

/// ODS: N numeric cell sets on a sheet of 1,000 rows × 10 columns (or more
/// rows for larger N); the leaf stages `MAX_CELL_CHANGES` (4,096) changes
/// per `set_cells` call in one edit, as the unified run does.
#[cfg(feature = "ods")]
fn ods(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_ods::worksheet::{CellChange, MAX_CELL_CHANGES, Selector};
    use litchi_ods::{Builder, Cell, CellValue, Package, Row, Sheet};
    const COLUMNS: usize = 10;
    let max = counts.iter().copied().max().unwrap_or(0);
    let rows = max.div_ceil(COLUMNS).max(1);
    let mut sheet = Sheet::new("Data")?;
    for row in 0..rows {
        let mut cells = Row::new();
        for column in 0..COLUMNS {
            let value = (row * COLUMNS + column) as f64;
            cells.push_cell(Cell::new(CellValue::Number(value), value.to_string()))?;
        }
        sheet.push_row(cells)?;
    }
    let mut builder = Builder::new();
    builder.add_sheet(sheet)?;
    let bytes = builder.build()?.to_bytes()?;
    let native = leak(Package::from_bytes(&bytes)?);
    let file = open(bytes)?;
    fn column(index: usize) -> String {
        let mut letters = Vec::new();
        let mut number = index + 1;
        while number > 0 {
            letters.push(b'A' + ((number - 1) % 26) as u8);
            number = (number - 1) / 26;
        }
        letters.reverse();
        String::from_utf8(letters).expect("ascii")
    }
    fn leaf(native: &Package, n: usize) -> Result<Package> {
        let mut edit = native.edit()?;
        edit.worksheets(|worksheets| {
            let mut batch = Vec::new();
            for index in 0..n {
                batch.push(CellChange::new(
                    index / COLUMNS,
                    index % COLUMNS,
                    Cell::new(CellValue::Number(-1.0), "-1"),
                ));
                if batch.len() == MAX_CELL_CHANGES {
                    worksheets.set_cells(
                        Selector::Position(Position::new(0)),
                        std::mem::take(&mut batch),
                    )?;
                }
            }
            if !batch.is_empty() {
                worksheets.set_cells(Selector::Position(Position::new(0)), batch)?;
            }
            Ok(())
        })?;
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let items = leak(
            (0..n)
                .map(|i| {
                    set(
                        format!("/Data/{}{}", column(i % COLUMNS), i / COLUMNS + 1),
                        "value",
                        -1,
                    )
                })
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "ods",
            scenario: "cell_value_set",
            owner: "default",
            count: n,
            check: check_pair(move || Ok((leaf(native, n)?.to_bytes()?, 1)), file, items),
            leaf: Box::new(move || timed(|| leaf(native, n))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// ODP: N paragraph text sets over 100-paragraph text boxes on N/100 slides;
/// the leaf stages every change in one `splice_scene` call.
#[cfg(feature = "odp")]
fn odp(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_odp::scene::{Change, ObjectRef, StoryRef};
    const LINES: usize = 100;
    let max = counts.iter().copied().max().unwrap_or(0);
    let slides = max.div_ceil(LINES).max(1);
    // Authored through the unified API (record 0278's deck).
    let mut deck = litchi_ops::create(litchi_core::FileFormat::Odp)?;
    let text = (0..LINES)
        .map(|l| format!("line {l} of the slide body"))
        .collect::<Vec<_>>()
        .join("\n");
    for slide in 0..slides {
        if slide > 0 {
            deck = deck.add("/", litchi_ops::NewNode::new("slide"))?.file;
        }
        deck = deck
            .add(
                format!("/slide[{}]", slide + 1).as_str(),
                litchi_ops::NewNode::new("shape")
                    .prop("text", text.as_str())
                    .prop("width", "20cm")
                    .prop("height", "12cm"),
            )?
            .file;
    }
    let bytes = deck.to_bytes()?;
    let native = leak(litchi_odp::Package::from_bytes(&bytes)?);
    let file = open(bytes)?;
    fn leaf(native: &litchi_odp::Package, changes: &[Change]) -> Result<litchi_odp::Package> {
        let mut edit = native.edit()?;
        edit.splice_scene(changes)?;
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let targets: Vec<(usize, usize)> = (0..n).map(|i| (i / LINES, i % LINES)).collect();
        let changes = leak(
            targets
                .iter()
                .map(|&(s, p)| Change::SetParagraphText {
                    story: StoryRef::object(ObjectRef::new(s, 0)),
                    paragraph: p,
                    text: format!("set {s} {p}"),
                })
                .collect::<Vec<_>>(),
        );
        let items = leak(
            targets
                .iter()
                .map(|&(s, p)| {
                    set(
                        format!("/slide[{}]/shape[1]/paragraph[{}]", s + 1, p + 1),
                        "text",
                        format!("set {s} {p}"),
                    )
                })
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "odp",
            scenario: "paragraph_text_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, changes)?.as_bytes().to_vec(), 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, changes))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// ODG: N shape text sets on one page of N named rectangles.
#[cfg(feature = "odg")]
fn odg(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_odg::shape::{Shape, ShapeKind};
    let max = counts.iter().copied().max().unwrap_or(0);
    let package = litchi_odg::Package::new()?;
    let mut edit = package.edit();
    edit.add_page(litchi_odg::page::Page::new("page1"))?;
    let package = edit.commit()?.into_package();
    let mut edit = package.edit();
    for i in 0..max {
        edit.add_shape(
            0,
            Shape::new(ShapeKind::Rectangle)
                .with_geometry("1cm", "1cm", "2cm", "1cm")
                .with_text(format!("shape {i}"))
                .with_name(format!("n{i}")),
        )?;
    }
    let bytes = edit.commit()?.into_package().to_bytes()?;
    let native = leak(litchi_odg::Package::from_bytes(&bytes)?);
    let file = open(bytes)?;
    fn leaf(native: &litchi_odg::Package, texts: &[String]) -> Result<litchi_odg::Package> {
        let mut edit = native.edit();
        for (i, text) in texts.iter().enumerate() {
            edit.set_shape_text(0, i, text.as_str())?;
        }
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let texts = leak((0..n).map(|i| format!("edited {i}")).collect::<Vec<_>>());
        let items = leak(
            texts
                .iter()
                .enumerate()
                .map(|(i, t)| set(format!("/page[1]/shape[{}]", i + 1), "text", t.clone()))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "odg",
            scenario: "shape_text_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, texts)?.as_bytes().to_vec(), 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, texts))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// ODC: N series label sets on a typed chart of N series; the leaf rewrites
/// the definition and replaces the chart once, as the unified run does.
#[cfg(feature = "odc")]
fn odc(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let max = counts.iter().copied().max().unwrap_or(0);
    let package = litchi_odc::Package::new()?;
    let mut definition = package.chart().definition()?;
    for dimension in [
        litchi_odc::chart::Dimension::X,
        litchi_odc::chart::Dimension::Y,
    ] {
        definition
            .plot_area
            .axes
            .push(litchi_odc::AxisSpec::new(dimension));
    }
    for i in 0..max {
        definition.plot_area.series.push(litchi_odc::SeriesSpec {
            values_cell_range_address: Some(format!("local-table.B{}:B{}", i + 2, i + 9)),
            ..litchi_odc::SeriesSpec::default()
        });
    }
    let mut edit = package.edit();
    edit.replace_chart(&definition)?;
    let bytes = edit.commit()?.into_package().to_bytes()?;
    let native = leak(litchi_odc::Package::from_bytes(&bytes)?);
    let file = open(bytes)?;
    fn leaf(native: &litchi_odc::Package, labels: &[String]) -> Result<litchi_odc::Package> {
        let mut definition = native.chart().definition()?;
        for (series, label) in definition.plot_area.series.iter_mut().zip(labels) {
            series.label_cell_address = Some(label.clone());
        }
        let mut edit = native.edit();
        edit.replace_chart(&definition)?;
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let labels = leak(
            (0..n)
                .map(|i| format!("local-table.A{}", i + 2))
                .collect::<Vec<_>>(),
        );
        let items = leak(
            labels
                .iter()
                .enumerate()
                .map(|(i, l)| set(format!("/series[{}]", i + 1), "label", l.clone()))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "odc",
            scenario: "series_label_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, labels)?.as_bytes().to_vec(), 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, labels))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// ODI holds exactly one frame, so no batch has 1,000 distinct targets: the
/// case is every independently settable non-geometry key of the frame (6
/// items, one leaf edit), whatever count is asked for. A new image has no
/// `meta.xml`, so its metadata keys are refused by the leaf and left out.
#[cfg(feature = "odi")]
fn odi(_counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let bytes = litchi_odi::Package::new()?.to_bytes()?;
    let native = leak(litchi_odi::Package::from_bytes(&bytes)?);
    let file = open(bytes)?;
    fn leaf(native: &litchi_odi::Package) -> Result<litchi_odi::Package> {
        let mut edit = native.edit();
        edit.set_frame_name(0, Some("frame name".into()))?;
        edit.set_frame_title(0, Some("frame title".into()))?;
        edit.set_frame_description(0, Some("frame description".into()))?;
        edit.set_anchor_type(0, Some("paragraph".into()))?;
        edit.set_z_index(0, Some(3))?;
        edit.set_layer(0, Some("layout".into()))?;
        Ok(edit.commit()?.into_package())
    }
    let items = leak(vec![
        set("/frame[1]".into(), "name", "frame name"),
        set("/frame[1]".into(), "title", "frame title"),
        set("/frame[1]".into(), "description", "frame description"),
        set("/frame[1]".into(), "anchorType", "paragraph"),
        set("/frame[1]".into(), "zIndex", 3),
        set("/frame[1]".into(), "layer", "layout"),
    ]);
    Ok(vec![Case {
        format: "odi",
        scenario: "frame_property_set",
        owner: "default",
        count: items.len(),
        check: check_pair(
            move || Ok((leaf(native)?.as_bytes().to_vec(), 1)),
            file,
            items,
        ),
        leaf: Box::new(move || timed(|| leaf(native))),
        unified: unified_leg(file, items),
    }])
}

/// ODB: N query command sets on a database of N queries.
#[cfg(feature = "odb")]
fn odb(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let max = counts.iter().copied().max().unwrap_or(0);
    let base = litchi_odb::Package::new()?;
    let mut edit = base.edit();
    for i in 0..max {
        edit.add_query(litchi_odb::Query::new(
            format!("q{i}"),
            format!("SELECT {i} FROM t"),
        ))?;
    }
    let bytes = edit.commit()?.into_package().to_bytes()?;
    let native = leak(litchi_odb::Package::from_bytes(&bytes)?);
    let file = open(bytes)?;
    fn leaf(
        native: &litchi_odb::Package,
        names: &[String],
        commands: &[String],
    ) -> Result<litchi_odb::Package> {
        let mut edit = native.edit();
        for (name, command) in names.iter().zip(commands) {
            edit.set_query_command(name, command.as_str())?;
        }
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let names = leak((0..n).map(|i| format!("q{i}")).collect::<Vec<_>>());
        let commands = leak(
            (0..n)
                .map(|i| format!("SELECT {i} FROM u"))
                .collect::<Vec<_>>(),
        );
        let items = leak(
            commands
                .iter()
                .enumerate()
                .map(|(i, c)| set(format!("/query[{}]", i + 1), "command", c.clone()))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "odb",
            scenario: "query_command_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, names, commands)?.as_bytes().to_vec(), 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, names, commands))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// ODF formula: N token text sets in one row of N identifiers.
#[cfg(feature = "odf-formula")]
fn formula(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_odf_formula::NodePath;
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut mathml = String::from("<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow>");
    for i in 0..max {
        if i > 0 {
            mathml.push_str("<mo>+</mo>");
        }
        mathml.push_str(&format!("<mi>a{i}</mi>"));
    }
    mathml.push_str("</mrow></math>");
    let bytes = litchi_odf_formula::Builder::from_mathml(mathml)
        .build()?
        .to_bytes()?;
    let native = leak(litchi_odf_formula::Package::from_bytes(&bytes)?);
    let file = open(bytes)?;
    fn leaf(
        native: &litchi_odf_formula::Package,
        routes: &[NodePath],
        texts: &[String],
    ) -> Result<litchi_odf_formula::Package> {
        let mut edit = native.edit();
        for (route, text) in routes.iter().zip(texts) {
            edit.set_text(route, text)?;
        }
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let texts = leak((0..n).map(|i| format!("b{i}")).collect::<Vec<_>>());
        let routes = leak(
            (0..n)
                .map(|i| NodePath::new(vec![0, i * 2]))
                .collect::<Vec<_>>(),
        );
        let items = leak(
            texts
                .iter()
                .enumerate()
                .map(|(i, t)| {
                    set(
                        format!("/element[1]/element[{}]", i * 2 + 1),
                        "text",
                        t.clone(),
                    )
                })
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "odf",
            scenario: "token_text_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, routes, texts)?.as_bytes().to_vec(), 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, routes, texts))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}
