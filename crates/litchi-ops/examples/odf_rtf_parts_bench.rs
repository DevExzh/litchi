//! Paired ODT, ODP and RTF timing for the parts of record 0328: header and
//! footer text, tab stops, hyperlinks, line breaks, field instructions,
//! bookmark names and section page setup, and the header, hyperlink,
//! section and tab-stop reads.
//!
//! Each write case times one edit against the same opened owner: the leaf
//! leg stages the leaf verb (`litchi_odt::transaction::Edit`, a
//! `litchi_odp` scene splice, a `litchi_rtf::edit::Edit` mark) and commits;
//! the unified leg runs the `File` verb on the matching path, including
//! path resolution, property admission and the returned path. Both legs
//! must publish the same bytes, checked outside the clock. Each read case
//! times a warm `get` and its `format` against the leaf reads the node
//! answers from. Rounds alternate ABBA.
//!
//! Run with `cargo run --release -p litchi-ops --example
//! odf_rtf_parts_bench --features odt,odp,rtf` (pin the CPU with
//! `taskset`). `PARTS_BENCH_CASE` restricts a run to one case.

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
use std::hint::black_box;
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
use std::time::Instant;

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
use litchi_ops::{NewNode, props};
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
use serde_json::json;

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
const ROUNDS: usize = 9;
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
const WARMUP: usize = 2;
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
const ITERATIONS: usize = 64;

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
type Outcome<T> = Result<T, Box<dyn std::error::Error>>;

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn fixture(name: &str) -> Outcome<Vec<u8>> {
    Ok(std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/unified-ops/0294")
            .join(name),
    )?)
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn main() -> Outcome<()> {
    let mut cases = Vec::new();
    odt_cases(&mut cases)?;
    odp_cases(&mut cases)?;
    rtf_cases(&mut cases)?;
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "fixtures": "test-data/unified-ops/0294/parts.{odt,odp,rtf}",
            "rounds": ROUNDS,
            "iterations": ITERATIONS,
            "cases": cases,
        }))?
    );
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn odt_cases(cases: &mut Vec<serde_json::Value>) -> Outcome<()> {
    use litchi_odt::body::parts::{Parts, RegionKind, tab_stops};
    use litchi_odt::body::{Kind, ParagraphPropertyUpdate, TabStopSpec};

    let bytes = fixture("parts.odt")?;
    let package = litchi_odt::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes(bytes)?;
    let document = package.document()?;
    let body = document.body()?;
    let parts = Parts::build(body)?;
    let view = parts.view(body);
    let footer = view
        .regions()
        .find(|region| {
            matches!(
                region.kind(),
                RegionKind::Footer | RegionKind::FooterLeft | RegionKind::FooterFirst
            )
        })
        .ok_or("footer")?;
    cases.push(write(
        "odt_footer_text_set",
        || {
            let mut edit = package.edit();
            edit.set_region_text(&footer, "Bottom line")?;
            Ok(edit.commit()?.into_package())
        },
        || Ok(file.set("/footer[1]", props! { "text" => "Bottom line" })?),
    )?);

    let paragraph = body
        .nodes()
        .find(|node| node.kind() == Kind::Paragraph && tab_stops(*node).len() == 2)
        .ok_or("tab paragraph")?;
    let paragraph_path = file
        .query("p")?
        .map(Result::unwrap)
        .find(|node| node.view().text().ok().flatten().as_deref() == Some("Name\tValue"))
        .ok_or("tab path")?
        .path()
        .to_string();
    let id = paragraph.id();
    let held: Vec<TabStopSpec> = tab_stops(paragraph)
        .iter()
        .map(|stop| {
            let mut edit = TabStopSpec::new(stop.position());
            edit.kind = stop.kind_attribute().map(str::to_owned);
            edit.leader_style = stop.leader_style().map(str::to_owned);
            edit.leader_text = stop.leader_text().map(str::to_owned);
            edit.character = stop.character().map(str::to_owned);
            edit
        })
        .collect();
    cases.push(write(
        "odt_tab_add",
        || {
            let mut stops = held.clone();
            let mut stop = TabStopSpec::new("144pt");
            stop.kind = Some("char".to_owned());
            stop.character = Some(".".to_owned());
            stops.insert(1, stop);
            let mut properties = ParagraphPropertyUpdate::new();
            properties.tab_stops = Some(stops);
            let mut edit = package.edit();
            edit.set_paragraph_properties(id, &properties)?;
            Ok(edit.commit()?.into_package())
        },
        || {
            Ok(file
                .add(
                    paragraph_path.as_str(),
                    NewNode::new("tab").prop("pos", 2880).prop("val", "decimal"),
                )?
                .file)
        },
    )?);

    let header = view.regions().next().ok_or("header")?;
    let _ = file.get("/header[1]")?.view().format()?.count();
    cases.push(read(
        "odt_header_get_format",
        || {
            black_box((header.kind(), header.master_name(), header.displayed()));
            Ok(())
        },
        || {
            let node = file.get("/header[1]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn odp_cases(cases: &mut Vec<serde_json::Value>) -> Outcome<()> {
    use litchi_odp::scene::{Change, ObjectRef, StoryRef};

    let bytes = fixture("parts.odp")?;
    let package = litchi_odp::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes(bytes)?;
    let story = StoryRef::object(ObjectRef::new(0, 0));
    let splice = |change: Change| -> Outcome<litchi_odp::Package> {
        let mut edit = package.edit()?;
        edit.splice_scene(&[change])?;
        Ok(edit.commit()?.into_package())
    };
    let run = "/slide[1]/shape[1]/paragraph[1]/run[2]/hyperlink";
    cases.push(write(
        "odp_link_set",
        || {
            splice(Change::SetRunLink {
                story: story.clone(),
                paragraph: 0,
                run: 1,
                href: Some("https://litchi.example/".to_owned()),
                title: None,
            })
        },
        || Ok(file.set(run, props! { "link" => "https://litchi.example/" })?),
    )?);
    cases.push(write(
        "odp_linebreak_add",
        || {
            splice(Change::InsertLineBreak {
                story: story.clone(),
                paragraph: 1,
            })
        },
        || {
            Ok(file
                .add("/slide[1]/shape[1]/paragraph[2]", NewNode::new("linebreak"))?
                .file)
        },
    )?);
    let presentation = package.presentation()?;
    let content = presentation.content_xml().to_owned();
    let scene = litchi_odp::scene::Scene::parse(&content)?;
    let _ = file.get(run)?.view().format()?.count();
    cases.push(read(
        "odp_link_get_format",
        || {
            let shape = scene.object(&ObjectRef::new(0, 0)).ok_or("shape")?;
            let link = shape
                .story()
                .and_then(|story| story.paragraphs().first())
                .and_then(|paragraph| paragraph.runs().get(1))
                .ok_or("run")?;
            black_box(litchi_odp::scene::parts::link_title(&content, link));
            black_box(link.href());
            Ok(())
        },
        || {
            let node = file.get(run)?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn rtf_cases(cases: &mut Vec<serde_json::Value>) -> Outcome<()> {
    use litchi_rtf::edit::SectionValue;
    use litchi_rtf::{PageOrientation, TabAlignment, TabLeader, TabStop};

    let bytes = fixture("parts.rtf")?;
    let package = litchi_rtf::Package::from_bytes(&bytes)?;
    let file = litchi_ops::open_bytes(bytes)?;
    let mark = |stage: &dyn Fn(
        &mut litchi_rtf::edit::Edit,
    ) -> Result<(), litchi_rtf::edit::Error>|
     -> Outcome<litchi_rtf::Package> {
        let mut edit = package.edit();
        stage(&mut edit)?;
        Ok(edit.commit()?.into_package())
    };
    cases.push(write(
        "rtf_field_instruction_set",
        || {
            mark(&|edit| {
                edit.set_field_instruction(1, "SECTIONPAGES")?;
                Ok(())
            })
        },
        || Ok(file.set("/field[2]", props! { "instruction" => "SECTIONPAGES" })?),
    )?);
    cases.push(write(
        "rtf_bookmark_rename",
        || {
            mark(&|edit| {
                edit.rename_bookmark(0, "Opening")?;
                Ok(())
            })
        },
        || Ok(file.set("/bookmark[1]", props! { "name" => "Opening" })?),
    )?);
    let held: Vec<TabStop> = package
        .body()
        .paragraphs()
        .next()
        .ok_or("p[1]")?
        .format()
        .tab_stops()
        .to_vec();
    cases.push(write(
        "rtf_tab_add",
        || {
            let mut stops = held.clone();
            stops.insert(
                1,
                TabStop {
                    position: 2880,
                    alignment: TabAlignment::Decimal,
                    leader: TabLeader::None,
                },
            );
            mark(&|edit| {
                edit.set_paragraph_tab_stops(0, &stops)?;
                Ok(())
            })
        },
        || {
            Ok(file
                .add(
                    "/body/p[1]",
                    NewNode::new("tab").prop("pos", 2880).prop("val", "decimal"),
                )?
                .file)
        },
    )?);
    cases.push(write(
        "rtf_section_set",
        || {
            mark(&|edit| {
                edit.set_section_values(
                    1,
                    &[
                        SectionValue::Orientation(PageOrientation::Portrait),
                        SectionValue::Columns(1),
                    ],
                )?;
                Ok(())
            })
        },
        || {
            Ok(file.set(
                "/section[2]",
                props! { "orientation" => "portrait", "columns" => 1 },
            )?)
        },
    )?);
    let _ = file.get("/section[2]")?.view().format()?.count();
    cases.push(read(
        "rtf_section_get_format",
        || {
            black_box(&package.sections().get(1).ok_or("section")?.properties);
            Ok(())
        },
        || {
            let node = file.get("/section[2]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    let _ = file.get("/body/p[1]/tab[2]")?.view().format()?.count();
    cases.push(read(
        "rtf_tab_get_format",
        || {
            let paragraph = package.body().paragraphs().next().ok_or("p[1]")?;
            black_box(paragraph.format().tab_stops().get(1));
            Ok(())
        },
        || {
            let node = file.get("/body/p[1]/tab[2]")?;
            black_box(node.view().format()?.count());
            Ok(())
        },
    )?);
    Ok(())
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    samples[samples.len() / 2]
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
#[allow(clippy::cast_precision_loss, reason = "nanosecond medians")]
fn paired(
    name: &str,
    mut leaf: impl FnMut() -> Outcome<()>,
    mut unified: impl FnMut() -> Outcome<()>,
) -> Outcome<serde_json::Value> {
    let sample = |leg: &mut dyn FnMut() -> Outcome<()>| -> Outcome<f64> {
        let start = Instant::now();
        for _ in 0..ITERATIONS {
            leg()?;
        }
        Ok(start.elapsed().as_nanos() as f64 / ITERATIONS as f64)
    };
    for _ in 0..WARMUP {
        sample(&mut leaf)?;
        sample(&mut unified)?;
    }
    let mut leaf_samples = Vec::new();
    let mut unified_samples = Vec::new();
    for round in 0..ROUNDS {
        if round % 2 == 0 {
            leaf_samples.push(sample(&mut leaf)?);
            unified_samples.push(sample(&mut unified)?);
            unified_samples.push(sample(&mut unified)?);
            leaf_samples.push(sample(&mut leaf)?);
        } else {
            unified_samples.push(sample(&mut unified)?);
            leaf_samples.push(sample(&mut leaf)?);
            leaf_samples.push(sample(&mut leaf)?);
            unified_samples.push(sample(&mut unified)?);
        }
    }
    let leaf_median = median(&mut leaf_samples);
    let unified_median = median(&mut unified_samples);
    Ok(json!({
        "case": name,
        "leaf_median_ns": leaf_median,
        "unified_median_ns": unified_median,
        "ratio": unified_median / leaf_median,
        "leaf_samples_ns": leaf_samples,
        "unified_samples_ns": unified_samples,
    }))
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn skipped(name: &str) -> bool {
    std::env::var("PARTS_BENCH_CASE").is_ok_and(|only| only != name)
}

/// The bytes a leaf package publishes.
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
trait Published {
    fn published(&self) -> Outcome<Vec<u8>>;
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
impl Published for litchi_odt::Package {
    fn published(&self) -> Outcome<Vec<u8>> {
        Ok(self.to_bytes()?)
    }
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
impl Published for litchi_odp::Package {
    fn published(&self) -> Outcome<Vec<u8>> {
        Ok(self.to_bytes()?)
    }
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
impl Published for litchi_rtf::Package {
    fn published(&self) -> Outcome<Vec<u8>> {
        Ok(self.to_bytes()?)
    }
}

/// Both legs publish the same bytes, checked outside the clock; the timed
/// legs stop at the published package or handle.
#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn write<P: Published>(
    name: &str,
    mut leaf: impl FnMut() -> Outcome<P>,
    mut unified: impl FnMut() -> Outcome<litchi_ops::File>,
) -> Outcome<serde_json::Value> {
    if skipped(name) {
        return Ok(json!({ "case": name, "skipped": true }));
    }
    if leaf()?.published()? != unified()?.to_bytes()? {
        return Err(format!("{name}: the legs publish different bytes").into());
    }
    paired(
        name,
        || {
            black_box(leaf()?);
            Ok(())
        },
        || {
            black_box(unified()?);
            Ok(())
        },
    )
}

#[cfg(all(feature = "odt", feature = "odp", feature = "rtf"))]
fn read(
    name: &str,
    leaf: impl FnMut() -> Outcome<()>,
    unified: impl FnMut() -> Outcome<()>,
) -> Outcome<serde_json::Value> {
    if skipped(name) {
        return Ok(json!({ "case": name, "skipped": true }));
    }
    paired(name, leaf, unified)
}

#[cfg(not(all(feature = "odt", feature = "odp", feature = "rtf")))]
fn main() {
    eprintln!("enable the `odt`, `odp` and `rtf` features");
}
