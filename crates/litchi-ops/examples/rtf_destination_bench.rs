//! Paired RTF destination-story write timing (record 0275).
//!
//! Each case compares the native destination editor with two unified routes:
//! one direct path `set` per target inside an edit closure (coalesced into one
//! native publication) and, where the selector grammar reaches the story,
//! one query fan-out `set` (also with the selector compiled once into a
//! `Plan`, record 0287; and with the query passed by reference, record 0305,
//! so neither leg clones an owned `Query` per call). Picture removal compares the native exact-source
//! remover with one unified query `remove`. Sources, paths, properties and
//! native operations are prepared outside timing; staging and commit are
//! timed. Every sample's output is checked byte-for-byte against the native
//! result outside the clock, and source immutability is checked per case.

#[cfg(feature = "rtf")]
use litchi_core::selector::{Path, Query, Scope};
#[cfg(feature = "rtf")]
use litchi_ops::query::Plan;
#[cfg(feature = "rtf")]
use litchi_ops::transaction::SemanticLimits;
#[cfg(feature = "rtf")]
use litchi_ops::{File, Props, open_bytes};
#[cfg(feature = "rtf")]
use litchi_rtf::edit::{EditOptions, HeaderFooterParagraph};
#[cfg(feature = "rtf")]
use serde_json::json;
#[cfg(feature = "rtf")]
use std::hint::black_box;
#[cfg(feature = "rtf")]
use std::time::Instant;

#[cfg(feature = "rtf")]
const PNG: &str = "89504e470d0a1a0a0000000d49484452";

#[cfg(feature = "rtf")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Family {
    Cell,
    Note,
    HeaderParagraph,
    Comment,
    PictureRemoval,
}

#[cfg(feature = "rtf")]
impl Family {
    const ALL: [Self; 5] = [
        Self::Cell,
        Self::Note,
        Self::HeaderParagraph,
        Self::Comment,
        Self::PictureRemoval,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::Cell => "table_cell_text",
            Self::Note => "footnote_text",
            Self::HeaderParagraph => "header_paragraph_text",
            Self::Comment => "comment_text",
            Self::PictureRemoval => "picture_removal",
        }
    }

    /// The native picture remover admits at most 64 pictures per transaction.
    const fn counts(self) -> &'static [usize] {
        match self {
            Self::PictureRemoval => &[1, 64],
            _ => &[1, 100, 1_000],
        }
    }

    /// The unified selector that fans out to every target, if the query
    /// walker reaches the story.
    const fn selector(self) -> Option<&'static str> {
        match self {
            Self::Cell => Some("tc"),
            Self::Note => Some("footnote"),
            Self::Comment => Some("comment"),
            Self::PictureRemoval => Some("picture"),
            Self::HeaderParagraph => None,
        }
    }

    fn source(self, count: usize) -> String {
        let mut body = String::from(r"{\rtf1\ansi");
        match self {
            Self::Cell => {
                body.push_str(r"\pard Lead\par ");
                for index in 0..count {
                    body.push_str(&format!(
                        r"\trowd\cellx2000\pard\intbl cell{index:05}\cell\row "
                    ));
                }
                body.push_str(r"\pard Tail\par}");
            },
            Self::Note => {
                body.push_str(r"\pard ");
                for index in 0..count {
                    body.push_str(&format!(
                        r"Text{index:05}{{\super\chftn{{\footnote\pard\plain\chftn note{index:05}}}}} "
                    ));
                }
                body.push_str(r"\par}");
            },
            Self::HeaderParagraph => {
                body.push_str(r"{\header\pard ");
                for index in 0..count {
                    body.push_str(&format!(r"head{index:05}\par "));
                }
                body.push_str(r"}\pard Body\par}");
            },
            Self::Comment => {
                body.push_str(r"\pard ");
                for index in 0..count {
                    body.push_str(&format!(
                        r"A{index:05}{{\*\atnid AM}}{{\*\atnauthor Ada}}\chatn{{\*\annotation comment{index:05}}} "
                    ));
                }
                body.push_str(r"\par}");
            },
            Self::PictureRemoval => {
                body.push_str(r"\pard ");
                for index in 0..count {
                    body.push_str(&format!(r"P{index:05}{{\pict\pngblip {PNG}}}"));
                }
                body.push_str(r"\par}");
            },
        }
        body
    }

    fn path(self, index: usize) -> String {
        match self {
            Self::Cell => format!("/body/tbl[1]/tr[{}]/tc[1]", index + 1),
            Self::Note => format!("/footnote[{}]", index + 1),
            Self::HeaderParagraph => format!("/header[1]/p[{}]", index + 1),
            Self::Comment => format!("/comments/comment[{}]", index + 1),
            Self::PictureRemoval => format!("/body/p[1]/picture[{}]", index + 1),
        }
    }
}

#[cfg(feature = "rtf")]
struct Prepared {
    family: Family,
    count: usize,
    bytes: Vec<u8>,
    native: litchi_rtf::Package,
    unified: File,
    native_options: EditOptions,
    limits: SemanticLimits,
    paths: Vec<Path<'static>>,
    props: Vec<Props<'static>>,
    texts: Vec<String>,
    fanout: Option<Query<'static>>,
    /// The fan-out selector compiled once for RTF (record 0287).
    fanout_plan: Option<Plan<'static>>,
}

#[cfg(feature = "rtf")]
impl Prepared {
    fn new(family: Family, count: usize) -> Result<Self, Box<dyn std::error::Error>> {
        let bytes = family.source(count).into_bytes();
        let native = litchi_rtf::Package::from_bytes(&bytes)?;
        let unified = open_bytes(bytes.as_slice())?;
        let native_limits = litchi_rtf::SemanticLimits::builder()
            .max_operations(100_000)?
            .max_replacement_bytes(64 * 1024 * 1024)?
            .build()?;
        let limits = SemanticLimits::builder()
            .max_operations(100_000)?
            .max_expanded_targets(100_000)?
            .max_staged_bytes(64 * 1024 * 1024)?
            .max_snapshot_bytes(512 * 1024 * 1024)?
            .build()?;
        let paths = (0..count)
            .map(|index| Path::parse(&family.path(index)).map(Path::into_owned))
            .collect::<Result<Vec<_>, _>>()?;
        let texts = (0..count)
            .map(|index| format!("changed{index:05}"))
            .collect::<Vec<_>>();
        let props = texts
            .iter()
            .map(|text| Props::new().with("text", text.clone()))
            .collect();
        let fanout = family.selector().map(|selector| {
            Query::parse(selector)
                .map(Query::into_owned)
                .map(|query| {
                    query.with_scope(Scope::Path(Box::new(
                        Path::parse("/").expect("root scope").into_owned(),
                    )))
                })
                .expect("fan-out selector")
        });
        let fanout_plan = fanout
            .as_ref()
            .map(|query| Plan::compile(query, litchi_core::FileFormat::Rtf))
            .transpose()?;
        Ok(Self {
            fanout_plan,
            family,
            count,
            bytes,
            native,
            unified,
            native_options: EditOptions::new().semantic_limits(native_limits),
            limits,
            paths,
            props,
            texts,
            fanout,
        })
    }

    /// The native destination edit: one staged operation per target. The
    /// fan-out comparison writes one shared value, so it has its own oracle.
    fn native_once(
        &self,
        shared: Option<&str>,
    ) -> Result<litchi_rtf::Package, litchi_rtf::edit::Error> {
        let mut edit = self.native.edit_with_options(self.native_options);
        let text = |index: usize| shared.unwrap_or(self.texts[index].as_str());
        match self.family {
            Family::Cell => {
                for index in 0..self.count {
                    edit.set_table_cell_text(
                        litchi_rtf::TableCellPath::outer(0, index, 0),
                        text(index),
                    )?;
                }
            },
            Family::Note => {
                for index in 0..self.count {
                    edit.set_note_text(index, text(index))?;
                }
            },
            Family::HeaderParagraph => {
                for index in 0..self.count {
                    edit.set_header_footer_text(
                        HeaderFooterParagraph::new(0, litchi_rtf::HeaderFooterType::Header, index),
                        text(index),
                    )?;
                }
            },
            Family::Comment => {
                for index in 0..self.count {
                    edit.set_annotation_text(index, text(index))?;
                }
            },
            Family::PictureRemoval => {
                let positions = (0..self.count).collect::<Vec<_>>();
                edit.remove_pictures(&positions)?;
            },
        }
        Ok(edit.commit()?.into_package())
    }

    fn unified_batch(&self) -> litchi_ops::Result<File> {
        if self.family == Family::PictureRemoval {
            // Positional picture paths shift after each removal; remove from
            // the end so every direct path stays valid.
            return self
                .unified
                .edit_with_options(
                    litchi_ops::EditOptions::new().semantic_limits(self.limits),
                    |edit| {
                        for path in self.paths.iter().rev() {
                            edit.remove(path)?;
                        }
                        Ok(())
                    },
                )
                .map(|commit| commit.into_file());
        }
        self.unified
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(self.limits),
                |edit| {
                    for (path, props) in self.paths.iter().zip(&self.props) {
                        edit.set(path, props)?;
                    }
                    Ok(())
                },
            )
            .map(|commit| commit.into_file())
    }

    /// One direct `File::set`/`remove` without an edit closure (single target).
    fn unified_direct(&self) -> litchi_ops::Result<File> {
        let path = &self.paths[0];
        if self.family == Family::PictureRemoval {
            self.unified.remove(path)
        } else {
            self.unified.set(path, &self.props[0])
        }
    }

    /// The fan-out with the selector compiled once outside the clock.
    fn unified_fanout_plan(&self, plan: &Plan<'static>) -> litchi_ops::Result<File> {
        self.unified
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(self.limits),
                |edit| {
                    if self.family == Family::PictureRemoval {
                        edit.remove(plan)
                    } else {
                        edit.set(plan, Props::new().with("text", "shared"))
                    }
                },
            )
            .map(|commit| commit.into_file())
    }

    /// The fan-out with the query passed by reference: the plan cache
    /// compiles it once, and no owned `Query` is cloned per call (record
    /// 0305).
    fn unified_fanout_ref(&self, query: &Query<'static>) -> litchi_ops::Result<File> {
        self.unified
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(self.limits),
                |edit| {
                    if self.family == Family::PictureRemoval {
                        edit.remove(query)
                    } else {
                        edit.set(query, Props::new().with("text", "shared"))
                    }
                },
            )
            .map(|commit| commit.into_file())
    }

    fn unified_fanout(&self, query: &Query<'static>) -> litchi_ops::Result<File> {
        self.unified
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(self.limits),
                |edit| {
                    if self.family == Family::PictureRemoval {
                        edit.remove(query.clone())
                    } else {
                        edit.set(query.clone(), Props::new().with("text", "shared"))
                    }
                },
            )
            .map(|commit| commit.into_file())
    }
}

#[cfg(feature = "rtf")]
enum Output {
    Native(litchi_rtf::Package),
    Unified(File),
}

#[cfg(feature = "rtf")]
fn median(values: &[u128]) -> f64 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    (sorted[sorted.len() / 2 - 1] + sorted[sorted.len() / 2]) as f64 / 2.0
}

#[cfg(feature = "rtf")]
fn run_case(case: &Prepared) -> Result<serde_json::Value, Box<dyn std::error::Error>> {
    let per_target = case.native_once(None)?.to_bytes()?;
    let shared = case.native_once(Some("shared"))?.to_bytes()?;
    let fanout_oracle = if case.family == Family::PictureRemoval {
        &per_target
    } else {
        &shared
    };
    let fanout_query = case.fanout.as_ref();
    let mut native = Vec::with_capacity(18);
    let mut batch = Vec::with_capacity(18);
    let mut native_shared = Vec::with_capacity(18);
    let mut fanout = Vec::with_capacity(18);
    let mut direct = Vec::with_capacity(18);
    let mut fanout_plan = Vec::with_capacity(18);
    let mut fanout_ref = Vec::with_capacity(18);
    // Two warm-up rounds, then nine A-B-C-D-E-F-G-G-F-E-D-C-B-A rounds: 18
    // samples per route. C/D (native and unified fan-out with one shared
    // value), F (the same fan-out with a precompiled plan, record 0287) and
    // G (the fan-out with the query by reference, record 0305) run only
    // where the query walker reaches the story; E (one direct
    // `File::set`/`remove`) only for a single target.
    for round in 0..11 {
        for route in [0u8, 1, 2, 3, 4, 5, 6, 6, 5, 4, 3, 2, 1, 0] {
            if matches!(route, 2 | 3 | 5 | 6) && fanout_query.is_none() {
                continue;
            }
            if route == 4 && case.count != 1 {
                continue;
            }
            let start = Instant::now();
            let (output, oracle) = match route {
                0 => (
                    Output::Native(black_box(case.native_once(None)?)),
                    &per_target,
                ),
                1 => (
                    Output::Unified(black_box(case.unified_batch()?)),
                    &per_target,
                ),
                2 if case.family == Family::PictureRemoval => (
                    Output::Native(black_box(case.native_once(None)?)),
                    fanout_oracle,
                ),
                2 => (
                    Output::Native(black_box(case.native_once(Some("shared"))?)),
                    fanout_oracle,
                ),
                4 => (
                    Output::Unified(black_box(case.unified_direct()?)),
                    &per_target,
                ),
                5 => (
                    Output::Unified(black_box(
                        case.unified_fanout_plan(case.fanout_plan.as_ref().ok_or("plan")?)?,
                    )),
                    fanout_oracle,
                ),
                6 => (
                    Output::Unified(black_box(
                        case.unified_fanout_ref(fanout_query.ok_or("query")?)?,
                    )),
                    fanout_oracle,
                ),
                _ => (
                    Output::Unified(black_box(
                        case.unified_fanout(fanout_query.ok_or("query")?)?,
                    )),
                    fanout_oracle,
                ),
            };
            let elapsed = start.elapsed().as_nanos();
            let bytes = match output {
                Output::Native(package) => package.to_bytes()?,
                Output::Unified(file) => file.to_bytes()?,
            };
            assert_eq!(&bytes, oracle);
            if round >= 2 {
                match route {
                    0 => native.push(elapsed),
                    1 => batch.push(elapsed),
                    2 => native_shared.push(elapsed),
                    3 => fanout.push(elapsed),
                    4 => direct.push(elapsed),
                    5 => fanout_plan.push(elapsed),
                    _ => fanout_ref.push(elapsed),
                }
            }
        }
    }
    assert_eq!(case.native.to_bytes()?, case.bytes);
    assert_eq!(case.unified.to_bytes()?, case.bytes);
    let native_median = median(&native);
    let batch_median = median(&batch);
    let fanout_median = (!fanout.is_empty()).then(|| median(&fanout));
    let shared_median = (!native_shared.is_empty()).then(|| median(&native_shared));
    let direct_median = (!direct.is_empty()).then(|| median(&direct));
    let fanout_plan_median = (!fanout_plan.is_empty()).then(|| median(&fanout_plan));
    let fanout_ref_median = (!fanout_ref.is_empty()).then(|| median(&fanout_ref));
    Ok(json!({
        "family": case.family.name(),
        "count": case.count,
        "source_bytes": case.bytes.len(),
        "native_ns": native,
        "unified_batch_ns": batch,
        "native_shared_ns": native_shared,
        "unified_fanout_ns": fanout,
        "native_median_ns": native_median,
        "unified_batch_median_ns": batch_median,
        "batch_ratio": batch_median / native_median,
        "native_shared_median_ns": shared_median,
        "unified_fanout_median_ns": fanout_median,
        "fanout_ratio": fanout_median.zip(shared_median).map(|(fanout, native)| fanout / native),
        "unified_direct_ns": direct,
        "unified_direct_median_ns": direct_median,
        "direct_ratio": direct_median.map(|value| value / native_median),
        "unified_fanout_plan_ns": fanout_plan,
        "unified_fanout_plan_median_ns": fanout_plan_median,
        "fanout_plan_ratio": fanout_plan_median.zip(shared_median).map(|(fanout, native)| fanout / native),
        "unified_fanout_ref_ns": fanout_ref,
        "unified_fanout_ref_median_ns": fanout_ref_median,
        "fanout_ref_ratio": fanout_ref_median.zip(shared_median).map(|(fanout, native)| fanout / native),
        "fanout_ref_over_plan": fanout_ref_median.zip(fanout_plan_median).map(|(by_ref, plan)| by_ref / plan),
        "exact_bytes_and_source_immutability": true
    }))
}

#[cfg(feature = "rtf")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut cases = Vec::new();
    for family in Family::ALL {
        for &count in family.counts() {
            let case = Prepared::new(family, count)?;
            cases.push(run_case(&case)?);
        }
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": "rtf_destination_benchmark.v1",
            "rounds": 9,
            "warmup_rounds": 2,
            "order": "ABCDEFGGFEDCBA",
            "timing": "warmed immutable owners; paths, properties and native operations prepared outside timing; staging and commit timed; serialization and validation excluded",
            "cases": cases
        }))?
    );
    Ok(())
}

#[cfg(not(feature = "rtf"))]
fn main() {
    eprintln!("enable the rtf feature");
}
