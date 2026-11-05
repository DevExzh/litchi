//! Paired RTF ordered-batch timing for native and unified direct sets.
//!
//! The timed region stages one prepared batch and commits it.  Serialization,
//! reopening, semantic checks, and source immutability checks are outside the
//! clock.  The unified side deliberately uses one direct path set per target;
//! its commit diagnostics prove whether compatible calls shared one native
//! commit.
//!
//! `RTF_BATCH_SET_COUNT` (record 0310) keeps only the cases of one target
//! count.
//!
//! Record 0315 adds the typed-props routes `paragraph_layout_alignment_typed`
//! and `run_compound_format_typed`: the same targets, native work and output,
//! with every length passed as a typed [`litchi_ops::value::Length`] instead
//! of the text `"6pt"`, so the run separates the cost of the text form of a
//! length from the rest of the write.

#[cfg(feature = "rtf")]
use litchi_core::Position;
#[cfg(feature = "rtf")]
use litchi_core::selector::Path;
#[cfg(feature = "rtf")]
use litchi_ops::transaction::{Commit, SemanticLimits};
#[cfg(feature = "rtf")]
use litchi_ops::value::Length;
#[cfg(feature = "rtf")]
use litchi_ops::{File, Props, open_bytes};
#[cfg(feature = "rtf")]
use litchi_rtf::Alignment;
#[cfg(feature = "rtf")]
use litchi_rtf::edit::{
    EditOptions, ParagraphAlignmentUpdate, ParagraphLayoutPatch, ParagraphLayoutUpdate,
    RunFormatPatch, RunFormatUpdate, RunTextReplacement,
};
#[cfg(feature = "rtf")]
use serde_json::json;
#[cfg(feature = "rtf")]
use std::hint::black_box;
#[cfg(feature = "rtf")]
use std::num::NonZeroU16;
#[cfg(feature = "rtf")]
use std::time::Instant;

#[cfg(feature = "rtf")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    const COUNTS: [usize; 5] = [1, 10, 100, 1_000, 10_000];
    // Record 0310: `RTF_BATCH_SET_COUNT` keeps one target count, for
    // profiling the fixed cost of the single-target cases.
    let only_count = std::env::var("RTF_BATCH_SET_COUNT")
        .ok()
        .and_then(|count| count.parse::<usize>().ok());
    let mut cases = Vec::new();
    for family in Family::ALL {
        for &shape in family.shapes() {
            for count in COUNTS {
                if only_count.is_some_and(|only| only != count) {
                    continue;
                }
                let case = Prepared::new(family, shape, count, false)?;
                cases.push(run_case(&case)?);
                assert_eq!(case.native.to_bytes()?, case.bytes);
                assert_eq!(case.unified.to_bytes()?, case.bytes);
            }
        }
    }
    // Record 0315: the typed-props routes, after every text route so the
    // earlier routes run as before.
    for family in [Family::ParagraphLayoutAlignment, Family::RunCompoundFormat] {
        for &shape in family.shapes() {
            for count in COUNTS {
                if only_count.is_some_and(|only| only != count) {
                    continue;
                }
                let case = Prepared::new(family, shape, count, true)?;
                cases.push(run_case(&case)?);
                assert_eq!(case.native.to_bytes()?, case.bytes);
                assert_eq!(case.unified.to_bytes()?, case.bytes);
            }
        }
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": "rtf_batch_set_benchmark.v1",
            "rounds": 9,
            "warmup_rounds": 2,
            "order": "ABBA",
            "families": [
                "paragraph_text",
                "paragraph_layout_alignment",
                "run_text",
                "run_compound_format",
                "paragraph_layout_alignment_typed",
                "run_compound_format_typed"
            ],
            "timing": "warmed immutable owners; ordered native updates, direct unified paths and properties prepared outside timing; staging and commit timed; serialization, reopening, semantic validation and source immutability excluded",
            "limits": {
                "operations": 100_000,
                "replacement_bytes": 64 * 1024 * 1024,
                "expanded_targets": 100_000,
                "staged_bytes": 64 * 1024 * 1024,
                "snapshot_bytes": 512 * 1024 * 1024
            },
            "cases": cases
        }))?
    );
    Ok(())
}

#[cfg(not(feature = "rtf"))]
fn main() {
    eprintln!("enable the rtf feature");
}

#[cfg(feature = "rtf")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Family {
    ParagraphText,
    ParagraphLayoutAlignment,
    RunText,
    RunCompoundFormat,
}

#[cfg(feature = "rtf")]
impl Family {
    const ALL: [Self; 4] = [
        Self::ParagraphText,
        Self::ParagraphLayoutAlignment,
        Self::RunText,
        Self::RunCompoundFormat,
    ];

    fn name(self) -> &'static str {
        match self {
            Self::ParagraphText => "paragraph_text",
            Self::ParagraphLayoutAlignment => "paragraph_layout_alignment",
            Self::RunText => "run_text",
            Self::RunCompoundFormat => "run_compound_format",
        }
    }

    fn shapes(self) -> &'static [Shape] {
        match self {
            Self::ParagraphText | Self::ParagraphLayoutAlignment => &[Shape::Paragraphs],
            Self::RunText | Self::RunCompoundFormat => &[Shape::Paragraphs, Shape::DenseRuns],
        }
    }
}

#[cfg(feature = "rtf")]
#[derive(Clone, Copy, Debug)]
enum Shape {
    Paragraphs,
    DenseRuns,
}

#[cfg(feature = "rtf")]
impl Shape {
    fn name(self) -> &'static str {
        match self {
            Self::Paragraphs => "separate_paragraphs",
            Self::DenseRuns => "dense_runs",
        }
    }
}

#[cfg(feature = "rtf")]
enum NativeWork {
    ParagraphText(Vec<litchi_rtf::edit::ParagraphTextReplacement>),
    ParagraphLayout {
        alignments: Vec<ParagraphAlignmentUpdate>,
        layouts: Vec<ParagraphLayoutUpdate>,
    },
    RunText(Vec<RunTextReplacement>),
    RunFormat(Vec<RunFormatUpdate>),
}

#[cfg(feature = "rtf")]
struct Prepared {
    family: Family,
    /// Lengths are typed values (record 0315's typed-props routes).
    typed: bool,
    shape: Shape,
    count: usize,
    bytes: Vec<u8>,
    native: litchi_rtf::Package,
    unified: File,
    native_options: EditOptions,
    limits: SemanticLimits,
    paths: Vec<Path<'static>>,
    props: Vec<Props<'static>>,
    native_work: NativeWork,
    expected_native_operations: usize,
}

#[cfg(feature = "rtf")]
impl Prepared {
    fn new(
        family: Family,
        shape: Shape,
        count: usize,
        typed: bool,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        let bytes = source(family, shape, count).into_bytes();
        let native = litchi_rtf::Package::from_bytes(&bytes)?;
        let unified = open_bytes(bytes.as_slice())?;
        let (paths, mut props, native_work, target_count) = prepare_work(family, shape, count)?;
        if typed {
            let typed_props = match family {
                Family::ParagraphLayoutAlignment => typed_paragraph_layout_props(),
                Family::RunCompoundFormat => typed_run_format_props(),
                Family::ParagraphText | Family::RunText => {
                    return Err("typed-props routes set lengths".into());
                },
            };
            props = vec![typed_props; props.len()];
        }

        // Both adapters receive the same finite operation/payload profile. The
        // native operation ceiling also admits the two native leaves in a
        // paragraph layout/alignment batch.
        let native_limits = litchi_rtf::SemanticLimits::builder()
            .max_operations(100_000)?
            .max_replacement_bytes(64 * 1024 * 1024)?
            .build()?;
        let native_options = EditOptions::new().semantic_limits(native_limits);
        let limits = SemanticLimits::builder()
            .max_operations(100_000)?
            .max_expanded_targets(100_000)?
            .max_staged_bytes(64 * 1024 * 1024)?
            .max_snapshot_bytes(512 * 1024 * 1024)?
            .build()?;
        let expected_native_operations = match family {
            Family::ParagraphLayoutAlignment => target_count * 2,
            Family::ParagraphText | Family::RunText | Family::RunCompoundFormat => target_count,
        };
        Ok(Self {
            family,
            typed,
            shape,
            count,
            bytes,
            native,
            unified,
            native_options,
            limits,
            paths,
            props,
            native_work,
            expected_native_operations,
        })
    }
}

#[cfg(feature = "rtf")]
fn source(family: Family, shape: Shape, count: usize) -> String {
    let groups = (0..count)
        .map(|index| {
            if matches!(
                family,
                Family::ParagraphText | Family::ParagraphLayoutAlignment
            ) {
                format!("paragraph{index:05}")
            } else {
                format!(r"{{\b0 plain{index:05}}}{{\i italic{index:05}}}{{\b0 tail{index:05}}}")
            }
        })
        .collect::<Vec<_>>();
    let body = match shape {
        Shape::Paragraphs => groups.join(r"\par "),
        Shape::DenseRuns => groups.join(""),
    };
    format!(r"{{\rtf1\ansi {body}}}")
}

#[cfg(feature = "rtf")]
fn prepare_work(
    family: Family,
    shape: Shape,
    count: usize,
) -> Result<(Vec<Path<'static>>, Vec<Props<'static>>, NativeWork, usize), Box<dyn std::error::Error>>
{
    let paragraph_count = match shape {
        Shape::Paragraphs => count,
        Shape::DenseRuns => 1,
    };
    match family {
        Family::ParagraphText => {
            let paths = (0..count)
                .map(|index| {
                    let value = format!("/body/p[{}]", index + 1);
                    Path::parse(&value).map(Path::into_owned)
                })
                .collect::<Result<Vec<_>, _>>()?;
            let props = (0..count)
                .map(|index| Props::new().with("text", format!("changed{index:05}")))
                .collect::<Vec<_>>();
            let replacements = (0..count)
                .map(|index| {
                    litchi_rtf::edit::ParagraphTextReplacement::new(
                        Position::new(index),
                        format!("changed{index:05}"),
                    )
                })
                .collect();
            Ok((paths, props, NativeWork::ParagraphText(replacements), count))
        },
        Family::ParagraphLayoutAlignment => {
            let paths = (0..count)
                .map(|index| {
                    let value = format!("/body/p[{}]", index + 1);
                    Path::parse(&value).map(Path::into_owned)
                })
                .collect::<Result<Vec<_>, _>>()?;
            let paragraph_props = paragraph_layout_props();
            let props = vec![paragraph_props; count];
            let alignments = (0..count)
                .map(|index| ParagraphAlignmentUpdate::new(Position::new(index), Alignment::Center))
                .collect();
            let layout = paragraph_layout_patch();
            let layouts = (0..count)
                .map(|index| ParagraphLayoutUpdate::new(Position::new(index), layout))
                .collect();
            Ok((
                paths,
                props,
                NativeWork::ParagraphLayout {
                    alignments,
                    layouts,
                },
                count,
            ))
        },
        Family::RunText | Family::RunCompoundFormat => {
            let target_count = count;
            let paths = (0..target_count)
                .map(|index| {
                    let paragraph = match shape {
                        Shape::Paragraphs => index,
                        Shape::DenseRuns => 0,
                    };
                    let run = match shape {
                        Shape::Paragraphs => 1,
                        Shape::DenseRuns => index * 3 + 1,
                    };
                    let value = format!("/body/p[{}]/r[{}]", paragraph + 1, run + 1);
                    Path::parse(&value).map(Path::into_owned)
                })
                .collect::<Result<Vec<_>, _>>()?;
            let props = match family {
                Family::RunText => (0..target_count)
                    .map(|index| Props::new().with("text", format!("changed{index:05}")))
                    .collect(),
                Family::RunCompoundFormat => vec![run_format_props(); target_count],
                Family::ParagraphText | Family::ParagraphLayoutAlignment => unreachable!(),
            };
            let native_work = match family {
                Family::RunText => NativeWork::RunText(
                    (0..target_count)
                        .map(|index| {
                            let paragraph = match shape {
                                Shape::Paragraphs => index,
                                Shape::DenseRuns => 0,
                            };
                            let run = match shape {
                                Shape::Paragraphs => 1,
                                Shape::DenseRuns => index * 3 + 1,
                            };
                            RunTextReplacement::new(
                                Position::new(paragraph),
                                Position::new(run),
                                format!("changed{index:05}"),
                            )
                        })
                        .collect(),
                ),
                Family::RunCompoundFormat => NativeWork::RunFormat(
                    (0..target_count)
                        .map(|index| {
                            let paragraph = match shape {
                                Shape::Paragraphs => index,
                                Shape::DenseRuns => 0,
                            };
                            let run = match shape {
                                Shape::Paragraphs => 1,
                                Shape::DenseRuns => index * 3 + 1,
                            };
                            RunFormatUpdate::new(
                                Position::new(paragraph),
                                Position::new(run),
                                run_format_patch(),
                            )
                        })
                        .collect(),
                ),
                Family::ParagraphText | Family::ParagraphLayoutAlignment => unreachable!(),
            };
            Ok((paths, props, native_work, target_count))
        },
    }
    .map(|(paths, props, work, targets)| {
        debug_assert!(paragraph_count >= 1);
        (paths, props, work, targets)
    })
}

#[cfg(feature = "rtf")]
fn paragraph_layout_patch() -> ParagraphLayoutPatch {
    ParagraphLayoutPatch::new()
        .with_space_before(120)
        .with_space_after(60)
        .with_left_indent(240)
        .with_right_indent(80)
        .with_first_line_indent(-40)
        .with_keep_with_next(true)
        .with_keep_together(true)
        .with_page_break_before(true)
}

#[cfg(feature = "rtf")]
fn paragraph_layout_props() -> Props<'static> {
    Props::new()
        .with("align", "center")
        .with("spaceBefore", "6pt")
        .with("spaceAfter", "3pt")
        .with("indent", "12pt")
        .with("rightIndent", "4pt")
        .with("firstLineIndent", "-2pt")
        .with("keepNext", true)
        .with("keepLines", true)
        .with("pageBreakBefore", true)
}

/// [`paragraph_layout_props`] with typed lengths (record 0315).
#[cfg(feature = "rtf")]
fn typed_paragraph_layout_props() -> Props<'static> {
    Props::new()
        .with("align", "center")
        .with("spaceBefore", Length::from_points(6))
        .with("spaceAfter", Length::from_points(3))
        .with("indent", Length::from_points(12))
        .with("rightIndent", Length::from_points(4))
        .with("firstLineIndent", Length::from_points(-2))
        .with("keepNext", true)
        .with("keepLines", true)
        .with("pageBreakBefore", true)
}

#[cfg(feature = "rtf")]
fn run_format_patch() -> RunFormatPatch {
    RunFormatPatch::new()
        .with_bold(true)
        .with_italic(true)
        .with_underline(litchi_rtf::UnderlineStyle::Double)
        .with_size(NonZeroU16::new(28).expect("nonzero size"))
        .with_strike(true)
        .with_double_strike(true)
        .with_hidden(true)
        .with_all_caps(true)
        .with_small_caps(true)
        .with_outline(true)
}

#[cfg(feature = "rtf")]
fn run_format_props() -> Props<'static> {
    Props::new()
        .with("bold", true)
        .with("italic", true)
        .with("underline", "double")
        .with("size", "14pt")
        .with("strike", true)
        .with("dstrike", true)
        .with("vanish", true)
        .with("caps", true)
        .with("smallcaps", true)
        .with("outline", true)
}

/// [`run_format_props`] with a typed size (record 0315).
#[cfg(feature = "rtf")]
fn typed_run_format_props() -> Props<'static> {
    Props::new()
        .with("bold", true)
        .with("italic", true)
        .with("underline", "double")
        .with("size", Length::from_points(14))
        .with("strike", true)
        .with("dstrike", true)
        .with("vanish", true)
        .with("caps", true)
        .with("smallcaps", true)
        .with("outline", true)
}

#[cfg(feature = "rtf")]
fn native_once(case: &Prepared) -> Result<(litchi_rtf::Package, usize), litchi_rtf::edit::Error> {
    let mut edit = case.native.edit_with_options(case.native_options);
    match &case.native_work {
        NativeWork::ParagraphText(replacements) => {
            edit.replace_body_paragraph_texts(replacements)?;
        },
        NativeWork::ParagraphLayout {
            alignments,
            layouts,
        } => {
            edit.set_body_paragraph_alignments(alignments)?;
            edit.patch_body_paragraph_layouts(layouts)?;
        },
        NativeWork::RunText(replacements) => {
            edit.replace_body_run_texts(replacements)?;
        },
        NativeWork::RunFormat(updates) => {
            edit.patch_body_run_formats(updates)?;
        },
    }
    let commit = edit.commit()?;
    let operations = commit.diagnostics().operation_count();
    Ok((commit.into_package(), operations))
}

#[cfg(feature = "rtf")]
fn unified_once(case: &Prepared) -> litchi_ops::Result<Commit<()>> {
    case.unified.edit_with_options(
        litchi_ops::EditOptions::new().semantic_limits(case.limits),
        |edit| {
            for (path, props) in case.paths.iter().zip(&case.props) {
                edit.set(path, props)?;
            }
            Ok(())
        },
    )
}

#[cfg(feature = "rtf")]
fn run_case(case: &Prepared) -> Result<serde_json::Value, Box<dyn std::error::Error>> {
    let (expected_native, expected_operations) = native_once(case)?;
    assert_eq!(expected_operations, case.expected_native_operations);
    let expected_bytes = expected_native.to_bytes()?;
    let mut native_ns = Vec::with_capacity(18);
    let mut unified_ns = Vec::with_capacity(18);
    for round in 0..11 {
        for native_route in [true, false, false, true] {
            if native_route {
                let start = Instant::now();
                let (output, operations) = black_box(native_once(case)?);
                let elapsed = start.elapsed().as_nanos();
                assert_eq!(operations, case.expected_native_operations);
                let output = output.to_bytes()?;
                assert_eq!(output, expected_bytes);
                validate_output(case, &output)?;
                if round >= 2 {
                    native_ns.push(elapsed);
                }
            } else {
                let start = Instant::now();
                let commit = black_box(unified_once(case)?);
                let elapsed = start.elapsed().as_nanos();
                assert_eq!(commit.diagnostics().commits(), 1);
                assert_eq!(commit.diagnostics().operations(), case.paths.len());
                assert_eq!(
                    commit.diagnostics().leaf_operations(),
                    case.expected_native_operations
                );
                let output = commit.file().to_bytes()?;
                assert_eq!(output, expected_bytes);
                validate_output(case, &output)?;
                if round >= 2 {
                    unified_ns.push(elapsed);
                }
            }
        }
    }
    assert_eq!(native_ns.len(), 18);
    assert_eq!(unified_ns.len(), 18);
    let native_median_ns = median(&native_ns);
    let unified_median_ns = median(&unified_ns);
    Ok(json!({
        "family": if case.typed {
            format!("{}_typed", case.family.name())
        } else {
            case.family.name().to_owned()
        },
        "shape": case.shape.name(),
        "count": case.count,
        "targets": case.paths.len(),
        "source_bytes": case.bytes.len(),
        "native_operations": case.expected_native_operations,
        "unified_operations": case.paths.len(),
        "unified_commits": 1,
        "native_ns": native_ns,
        "unified_ns": unified_ns,
        "native_median_ns": native_median_ns,
        "unified_median_ns": unified_median_ns,
        "ratio": unified_median_ns / native_median_ns,
        "exact_bytes_reopened_semantics_and_source_immutability": true
    }))
}

#[cfg(feature = "rtf")]
fn median(values: &[u128]) -> f64 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    (sorted[sorted.len() / 2 - 1] + sorted[sorted.len() / 2]) as f64 / 2.0
}

#[cfg(feature = "rtf")]
fn validate_output(case: &Prepared, output: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let reopened = litchi_rtf::Package::from_bytes(output)?;
    let paragraph_count = match case.shape {
        Shape::Paragraphs => case.count,
        Shape::DenseRuns => 1,
    };
    assert_eq!(reopened.paragraph_count(), paragraph_count);
    match case.family {
        Family::ParagraphText => {
            for (index, paragraph) in reopened.body().paragraphs().enumerate() {
                assert_eq!(paragraph.to_text(), format!("changed{index:05}"));
            }
        },
        Family::ParagraphLayoutAlignment => {
            for paragraph in reopened.body().paragraphs() {
                assert!(paragraph.to_text().starts_with("paragraph"));
                assert_eq!(paragraph.format().alignment(), Alignment::Center);
                let layout = paragraph.format().layout();
                assert_eq!(layout.space_before(), 120);
                assert_eq!(layout.space_after(), 60);
                assert_eq!(layout.left_indent(), 240);
                assert_eq!(layout.right_indent(), 80);
                assert_eq!(layout.first_line_indent(), -40);
                assert!(layout.keep_with_next());
                assert!(layout.keep_together());
                assert!(layout.page_break_before());
            }
        },
        Family::RunText | Family::RunCompoundFormat => {
            for (paragraph_index, paragraph) in reopened.body().paragraphs().enumerate() {
                let runs: Vec<_> = paragraph.runs().collect();
                let expected_runs = match case.shape {
                    Shape::Paragraphs => 3,
                    Shape::DenseRuns => case.count * 3,
                };
                assert_eq!(runs.len(), expected_runs);
                for (run_index, run) in runs.iter().enumerate() {
                    let unit = match case.shape {
                        Shape::Paragraphs => paragraph_index,
                        Shape::DenseRuns => run_index / 3,
                    };
                    let middle = run_index % 3 == 1;
                    let selected = middle;
                    let expected_text = if middle {
                        match case.family {
                            Family::RunText => format!("changed{unit:05}"),
                            Family::RunCompoundFormat => format!("italic{unit:05}"),
                            Family::ParagraphText | Family::ParagraphLayoutAlignment => {
                                unreachable!()
                            },
                        }
                    } else if run_index % 3 == 0 {
                        format!("plain{unit:05}")
                    } else {
                        format!("tail{unit:05}")
                    };
                    assert_eq!(run.text(), expected_text);
                    if case.family == Family::RunText {
                        assert_eq!(run.format().italic(), middle);
                        assert!(!run.format().bold());
                        assert_eq!(run.format().size().get(), 24);
                    } else {
                        assert_eq!(run.format().bold(), selected);
                        assert_eq!(run.format().italic(), selected);
                        assert_eq!(
                            run.format().underline(),
                            if selected {
                                litchi_rtf::text::Underline::Double
                            } else {
                                litchi_rtf::text::Underline::None
                            }
                        );
                        assert_eq!(run.format().size().get(), if selected { 28 } else { 24 });
                        assert_eq!(run.format().strike(), selected);
                        assert_eq!(run.format().double_strike(), selected);
                        assert_eq!(run.format().hidden(), selected);
                        assert_eq!(run.format().all_caps(), selected);
                        assert_eq!(run.format().small_caps(), selected);
                        assert_eq!(run.format().outline(), selected);
                    }
                }
            }
        },
    }
    Ok(())
}
