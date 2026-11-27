//! Differential checks for the symbolic generated-name plan language.
//!
//! The implementation deliberately does not expand indexed ranges while it
//! constructs a plan.  These tests use small finite ranges and an independent
//! expansion oracle to check that every plan the public builder accepts is
//! conflict-free.  Refusing a plan is allowed by the public contract, so the
//! differential property is one-way: accepted plans must agree with the
//! finite reference model.

use soapberry_zip::generated_names::{GeneratedNamePlanBuilder, GeneratedNamePlanLimits};
use soapberry_zip::path::ZipFilePath;

#[derive(Clone, Copy)]
enum Spec {
    Literal(&'static str),
    Indexed {
        first: u64,
        count: u64,
        patterns: &'static [(&'static str, &'static str)],
    },
}

const ONE_XML: &[(&str, &str)] = &[("a/entry", ".xml")];
const ONE_XML_RELS: &[(&str, &str)] = &[("a/entry", ".xml.rels")];
const SLIDE_XML_AND_RELS: &[(&str, &str)] = &[
    ("ppt/slides/slide", ".xml"),
    ("ppt/slides/slide", ".xml.rels"),
];
const LAYOUT_XML_AND_RELS: &[(&str, &str)] = &[
    ("ppt/slideLayouts/slideLayout", ".xml"),
    ("ppt/slideLayouts/slideLayout", ".xml.rels"),
];
const INTERNAL_DIGITS: &[(&str, &str)] = &[
    ("ppt/slides/slide2-", ".xml"),
    ("ppt/notes/note-", "-2.xml"),
];

fn limits() -> GeneratedNamePlanLimits {
    GeneratedNamePlanLimits::new(64, 4 * 1024, 4 * 1024)
}

fn build(
    specs: &[Spec],
) -> Result<soapberry_zip::generated_names::GeneratedNamePlan, soapberry_zip::Error> {
    let mut builder = GeneratedNamePlanBuilder::new(limits()).expect("valid test limits");
    for spec in specs {
        match *spec {
            Spec::Literal(name) => builder.push_literal(name)?,
            Spec::Indexed {
                first,
                count,
                patterns,
            } => builder.push_indexed(first, count, patterns)?,
        }
    }
    builder.finish()
}

fn expand(specs: &[Spec]) -> Vec<String> {
    let mut names = Vec::new();
    for spec in specs {
        match *spec {
            Spec::Literal(name) => names.push(name.to_owned()),
            Spec::Indexed {
                first,
                count,
                patterns,
            } => {
                for value in first..first + count {
                    for &(prefix, suffix) in patterns {
                        names.push(format!("{prefix}{value}{suffix}"));
                    }
                }
            },
        }
    }
    names
}

fn fold_eq(left: &str, right: &str) -> bool {
    left.bytes()
        .map(|byte| byte.to_ascii_lowercase())
        .eq(right.bytes().map(|byte| byte.to_ascii_lowercase()))
}

fn component_ancestor(left: &str, right: &str) -> bool {
    let left_components: Vec<_> = left.split('/').collect();
    let right_components: Vec<_> = right.split('/').collect();
    left_components.len() < right_components.len()
        && left_components
            .iter()
            .zip(right_components.iter())
            .all(|(left, right)| fold_eq(left, right))
}

fn reference_conflicts(names: &[String]) -> Vec<(usize, usize)> {
    let mut conflicts = Vec::new();
    for (left_index, left) in names.iter().enumerate() {
        for (right_index, right) in names.iter().enumerate().skip(left_index + 1) {
            if fold_eq(left, right)
                || component_ancestor(left, right)
                || component_ancestor(right, left)
            {
                conflicts.push((left_index, right_index));
            }
        }
    }
    conflicts
}

fn assert_accepted_plans_are_conflict_free(specs: &[Spec]) {
    if let Ok(plan) = build(specs) {
        let names = expand(specs);
        for name in &names {
            assert_eq!(
                ZipFilePath::from_str(name).as_str(),
                name,
                "accepted generated name is changed by the production normalizer"
            );
        }
        assert_eq!(
            plan.entry_count(),
            names.len() as u64,
            "symbolic entry count disagrees with finite expansion"
        );
        assert_eq!(
            reference_conflicts(&names),
            Vec::new(),
            "accepted symbolic plan has a reference conflict: {specs_len} descriptors",
            specs_len = specs.len()
        );
    }
}

#[test]
fn exhaustive_small_literal_and_family_pairs_have_no_false_acceptance() {
    const FAMILIES: &[Spec] = &[
        Spec::Indexed {
            first: 1,
            count: 2,
            patterns: ONE_XML,
        },
        Spec::Indexed {
            first: 3,
            count: 2,
            patterns: ONE_XML,
        },
        Spec::Indexed {
            first: 2,
            count: 2,
            patterns: ONE_XML,
        },
        Spec::Indexed {
            first: 1,
            count: 2,
            patterns: ONE_XML_RELS,
        },
        Spec::Indexed {
            first: 9,
            count: 2,
            patterns: ONE_XML,
        },
        Spec::Indexed {
            first: 11,
            count: 2,
            patterns: ONE_XML,
        },
        Spec::Indexed {
            first: 1,
            count: 2,
            patterns: &[("a/entry", ".XML")],
        },
        Spec::Indexed {
            first: 1,
            count: 2,
            patterns: &[("a/", ".xml")],
        },
        Spec::Indexed {
            first: 1,
            count: 2,
            patterns: &[("a/entry/", ".xml")],
        },
        Spec::Indexed {
            first: 1,
            count: 2,
            patterns: INTERNAL_DIGITS,
        },
        Spec::Indexed {
            first: 1,
            count: 2,
            patterns: &[("a/b/entry", ".xml")],
        },
    ];
    const LITERALS: &[Spec] = &[
        Spec::Literal("a"),
        Spec::Literal("A"),
        Spec::Literal("a/entry1.xml"),
        Spec::Literal("a/entry"),
        Spec::Literal("a/entry/child"),
        Spec::Literal("a/b/entry1.xml"),
        Spec::Literal("a/b/entry1.xml/child"),
        Spec::Literal("a/entry10.xml"),
        Spec::Literal("ppt/slides/slide1.xml"),
        Spec::Literal("ppt/slides/slide1.xml/child"),
    ];

    for first in FAMILIES {
        for second in FAMILIES {
            assert_accepted_plans_are_conflict_free(&[*first, *second]);
            assert_accepted_plans_are_conflict_free(&[*second, *first]);
        }
        for literal in LITERALS {
            assert_accepted_plans_are_conflict_free(&[*literal, *first]);
            assert_accepted_plans_are_conflict_free(&[*first, *literal]);
        }
    }
    for first in LITERALS {
        for second in LITERALS {
            assert_accepted_plans_are_conflict_free(&[*first, *second]);
            assert_accepted_plans_are_conflict_free(&[*second, *first]);
        }
    }
}

#[test]
fn exact_folded_and_whole_component_conflicts_are_rejected() {
    let rejected = [
        vec![
            Spec::Literal("ppt/slides/a.xml"),
            Spec::Literal("ppt/slides/a.xml"),
        ],
        vec![
            Spec::Literal("ppt/slides/a.xml"),
            Spec::Literal("PPT/SLIDES/A.XML"),
        ],
        vec![
            Spec::Literal("ppt/slides"),
            Spec::Literal("ppt/slides/a.xml"),
        ],
        vec![
            Spec::Literal("ppt/slides/a.xml"),
            Spec::Literal("ppt/slides/a.xml/x"),
        ],
        vec![
            Spec::Literal("ppt/slides"),
            Spec::Indexed {
                first: 1,
                count: 2,
                patterns: SLIDE_XML_AND_RELS,
            },
        ],
        vec![
            Spec::Indexed {
                first: 1,
                count: 2,
                patterns: &[("ppt/slides/slide", "")],
            },
            Spec::Literal("ppt/slides/slide1/child"),
        ],
        vec![
            Spec::Indexed {
                first: 1,
                count: 2,
                patterns: &[("ppt/slides/slide", "")],
            },
            Spec::Indexed {
                first: 1,
                count: 2,
                patterns: &[("ppt/slides/slide1/", ".xml")],
            },
        ],
        vec![
            Spec::Indexed {
                first: 1,
                count: 2,
                patterns: &[("a/", "")],
            },
            Spec::Indexed {
                first: 1,
                count: 2,
                patterns: &[("a/1/", ".xml")],
            },
        ],
        vec![
            Spec::Indexed {
                first: 1,
                count: 2,
                patterns: ONE_XML,
            },
            Spec::Indexed {
                first: 2,
                count: 2,
                patterns: &[("A/ENTRY", ".XML")],
            },
        ],
    ];
    for specs in rejected {
        assert!(
            build(&specs).is_err(),
            "conflicting descriptors were accepted: {} names",
            expand(&specs).len()
        );
    }
}

#[test]
fn common_pptx_families_and_decimal_range_transitions_succeed() {
    let mut builder =
        GeneratedNamePlanBuilder::new(GeneratedNamePlanLimits::new(32, 4096, 1024)).unwrap();
    builder.push_literal("[Content_Types].xml").unwrap();
    builder.push_literal("_rels/.rels").unwrap();
    builder.push_indexed(1, 9, SLIDE_XML_AND_RELS).unwrap();
    builder.push_indexed(10, 247, SLIDE_XML_AND_RELS).unwrap();
    builder.push_indexed(1, 11, LAYOUT_XML_AND_RELS).unwrap();
    builder.push_indexed(12, 11, LAYOUT_XML_AND_RELS).unwrap();
    let plan = builder.finish().unwrap();
    assert_eq!(plan.entry_count(), 2 + 256 * 2 + 22 * 2);
    assert_eq!(
        plan.max_name_bytes(),
        "ppt/slideLayouts/slideLayout11.xml.rels".len()
    );
}

#[test]
fn internal_digits_are_supported_but_ambiguous_numeric_slots_are_not() {
    let mut builder = GeneratedNamePlanBuilder::new(limits()).unwrap();
    builder.push_indexed(1, 3, INTERNAL_DIGITS).unwrap();
    let plan = builder.finish().unwrap();
    assert_eq!(plan.entry_count(), 6);

    let mut ambiguous_suffix = GeneratedNamePlanBuilder::new(limits()).unwrap();
    assert!(ambiguous_suffix.push_indexed(1, 1, &[("a", "2")]).is_err());

    let mut ambiguous_prefix = GeneratedNamePlanBuilder::new(limits()).unwrap();
    assert!(ambiguous_prefix.push_indexed(1, 1, &[("a1", "")]).is_err());

    // Even when each static fragment has a safe boundary, two families with
    // different numeric slots are conservatively refused if their prefixes
    // could align after a decimal expansion. The finite oracle below only
    // constrains accepted plans, so this is a deliberate supported-language
    // boundary rather than a soundness assumption.
    let mut ambiguous_slots = GeneratedNamePlanBuilder::new(limits()).unwrap();
    assert!(
        ambiguous_slots
            .push_indexed(1, 2, &[("a", "x2"), ("a1x", "")])
            .is_err()
    );
}

#[test]
fn empty_plans_and_rejected_insertions_are_atomic() {
    let empty = GeneratedNamePlanBuilder::new(GeneratedNamePlanLimits::new(0, 0, 0))
        .unwrap()
        .finish()
        .unwrap();
    assert_eq!(empty.entry_count(), 0);
    assert_eq!(empty.max_name_bytes(), 0);

    let mut builder = GeneratedNamePlanBuilder::new(limits()).unwrap();
    builder.push_literal("a/first").unwrap();
    assert!(builder.push_literal("A/FIRST").is_err());
    assert!(builder.push_literal("a/first/child").is_err());
    builder.push_literal("a/second").unwrap();
    let plan = builder.finish().unwrap();
    assert_eq!(plan.entry_count(), 2);

    let mut range_builder =
        GeneratedNamePlanBuilder::new(GeneratedNamePlanLimits::new(4, 64, 8)).unwrap();
    range_builder.push_indexed(1, 2, ONE_XML).unwrap();
    assert!(range_builder.push_indexed(2, 2, ONE_XML).is_err());
    range_builder.push_indexed(3, 2, ONE_XML_RELS).unwrap();
    assert_eq!(range_builder.finish().unwrap().entry_count(), 4);
}

#[test]
fn huge_symbolic_ranges_use_bounded_representation() {
    let mut builder =
        GeneratedNamePlanBuilder::new(GeneratedNamePlanLimits::new(1, 32, u64::MAX)).unwrap();
    builder
        .push_indexed(0, u64::MAX, &[("ppt/slides/slide", ".xml")])
        .unwrap();
    let plan = builder.finish().unwrap();
    assert_eq!(plan.entry_count(), u64::MAX);
    assert_eq!(
        plan.max_name_bytes(),
        "ppt/slides/slide18446744073709551614.xml".len()
    );

    let mut max_value =
        GeneratedNamePlanBuilder::new(GeneratedNamePlanLimits::new(1, 16, 1)).unwrap();
    max_value
        .push_indexed(u64::MAX, 1, &[("x", ".bin")])
        .unwrap();
    assert_eq!(max_value.finish().unwrap().max_name_bytes(), 1 + 20 + 4);
}
