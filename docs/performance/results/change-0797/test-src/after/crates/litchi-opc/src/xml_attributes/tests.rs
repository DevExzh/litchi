use core::cell::Cell;
use std::borrow::Cow;

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::AttrError;

use super::{BytesStartExt, QUICK_XML_LINEAR_NAMES};

thread_local! {
    static COMPARISONS: Cell<usize> = const { Cell::new(0) };
}

pub(super) fn count_comparison() {
    COMPARISONS.with(|count| count.set(count.get() + 1));
}

fn counted<T>(run: impl FnOnce() -> T) -> (T, usize) {
    COMPARISONS.with(|count| count.set(0));
    let result = run();
    (result, COMPARISONS.with(Cell::get))
}

type Item = Result<(Vec<u8>, Vec<u8>), AttrError>;

fn owned(item: Result<quick_xml::events::attributes::Attribute<'_>, AttrError>) -> Item {
    item.map(|attribute| (attribute.key.as_ref().to_vec(), attribute.value.to_vec()))
}

/// quick-xml's checked iterator, up to and including its first error: what a
/// reader that stops at the first error sees.
#[allow(clippy::disallowed_methods)]
fn quick_xml_until_error(tag: &BytesStart<'_>) -> Vec<Item> {
    let mut items = Vec::new();
    for item in tag.attributes() {
        let error = item.is_err();
        items.push(owned(item));
        if error {
            break;
        }
    }
    items
}

fn checked(tag: &BytesStart<'_>) -> Vec<Item> {
    tag.checked_attributes().map(owned).collect()
}

fn tag(content: &str) -> BytesStart<'static> {
    let name_len = content
        .find([' ', '\t', '\r', '\n'])
        .unwrap_or(content.len());
    BytesStart::from_content(content.to_owned(), name_len)
}

fn assert_same(content: &str) {
    let tag = tag(content);
    assert_eq!(checked(&tag), quick_xml_until_error(&tag), "{content}");
}

fn names(count: usize) -> Vec<String> {
    (0..count)
        .map(|index| format!("n{index}=\"{index}\""))
        .collect()
}

#[test]
fn well_formed_tags_yield_what_quick_xml_yields() {
    for content in [
        "e",
        "e ",
        "e a=\"1\"",
        "w:rPr w:val=\"x\" w:eastAsia='y' xmlns:w=\"urn:w\"",
        "e a = \"1\"\tb='2'\n c=\"\" \r\n",
        "e a=\"&amp;\" b=\"x y z\" c='\"' d=\"'\"",
        "e a=\"1\"b=\"2\"c='3'",
    ] {
        assert_same(content);
    }
    for count in [31, 32, 33, 34, 64, 500] {
        assert_same(&format!("e {}", names(count).join(" ")));
        assert_same(&format!("e {}", names(count).join("")));
        assert_same(&format!("e\n{}\n", names(count).join("\n\t ")));
    }
}

#[test]
fn a_duplicate_anywhere_is_reported_where_quick_xml_reports_it() {
    let values = [
        "\"r\"",
        "'r'",
        " = \"r\"",
        "=x",
        "=",
        "=\"unclosed",
        "='unclosed",
        " =  ",
    ];
    for count in [2, 3, 31, 32, 33, 34, 40, 70] {
        for repeated in [0, 1, count / 2, 30, 31, 32, 33, count - 1] {
            if repeated >= count {
                continue;
            }
            for at in [1, 2, 31, 32, 33, 34, count] {
                if at <= repeated || at > count {
                    continue;
                }
                for value in values {
                    let mut attributes = names(count);
                    let duplicate = if value.starts_with(['=', ' ']) {
                        format!("n{repeated}{value}")
                    } else {
                        format!("n{repeated}={value}")
                    };
                    attributes.insert(at, duplicate);
                    assert_same(&format!("e {}", attributes.join(" ")));
                }
            }
        }
    }
}

#[test]
fn malformed_attributes_around_the_switch_yield_what_quick_xml_yields() {
    let tails = [
        "flag",
        "flag next=\"1\"",
        "k=x",
        "k=x next=\"1\"",
        "k=",
        "k =",
        "k=\"open",
        "k='open",
        "n0=x",
        "n0=",
        "n0 = ",
        "n0=\"open",
        "n0",
        "n0 next=\"1\"",
        "=\"1\"",
        "k=\"1\"k=\"2\"",
        "k=\"1\" n0=\"x y\" z=\"3\"",
        "k=\"x\" k=\"a b='c'\" d=\"4\"",
    ];
    for count in 28..=36 {
        for tail in tails {
            assert_same(&format!("e {} {tail}", names(count).join(" ")));
            assert_same(&format!("e {} {tail} tail=\"t\"", names(count).join(" ")));
        }
    }
}

/// Every sequence of three tail items after 30 to 34 distinct names.
#[test]
fn every_short_tail_after_the_switch_yields_what_quick_xml_yields() {
    let items = [
        "fresh{}=\"v\"",
        "n0=\"v\"",
        "n{last}=\"v\"",
        "fresh0=\"again\"",
        "fresh{}=x",
        "n0=x",
        "bare{}",
        "n1 = 'v'",
        "fresh{}=\"a b\"",
    ];
    let mut cases = 0usize;
    for count in 30..=34 {
        let prefix = names(count);
        for first in items {
            for second in items {
                for third in items {
                    let mut attributes = prefix.clone();
                    for (index, item) in [first, second, third].into_iter().enumerate() {
                        attributes.push(
                            item.replace("{last}", &(count - 1).to_string())
                                .replace("{}", &index.to_string()),
                        );
                    }
                    assert_same(&format!("e {}", attributes.join(" ")));
                    cases += 1;
                }
            }
        }
    }
    assert_eq!(cases, 5 * 9 * 9 * 9);
}

/// A small deterministic generator (xorshift64*) for random tags.
struct Random(u64);

impl Random {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 >> 12;
        self.0 ^= self.0 << 25;
        self.0 ^= self.0 >> 27;
        self.0.wrapping_mul(0x2545_f491_4f6c_dd1d)
    }

    fn below(&mut self, bound: usize) -> usize {
        usize::try_from(self.next() % u64::try_from(bound).unwrap()).unwrap()
    }
}

#[test]
fn random_tags_yield_what_quick_xml_yields() {
    let separators = [" ", "  ", "\t", "\n", "\r\n", ""];
    let mut random = Random(0x0770_0770_0770_0770);
    // Outcomes of the tags whose first error (or end) comes after the 32nd
    // name, where this iterator checks the names: every kind must be seen.
    let mut late = std::collections::BTreeMap::<&str, usize>::new();
    for _ in 0..20_000 {
        let attributes = random.below(90);
        let alphabet = 1 + random.below(120);
        let mut content = String::from("e");
        // Half the tags start with distinct names that reach the switch.
        if random.below(2) == 0 {
            for index in 0..28 + random.below(12) {
                content.push_str(&format!(" p{index}=\"{index}\""));
            }
        }
        for _ in 0..attributes {
            content.push_str(separators[random.below(separators.len())]);
            let name = format!("a{}", random.below(alphabet));
            let item = match random.below(40) {
                0 => name,
                1 => format!("{name}=unquoted"),
                2 => format!("{name}="),
                3 => format!("{name}=\"open"),
                4 => format!("{name} = 'x y'"),
                5 => format!("{name}=\"a b='c'\""),
                6 => "=\"v\"".to_owned(),
                _ => format!("{name}=\"{}\"", random.below(10)),
            };
            content.push_str(&item);
        }
        assert_same(&content);
        let items = checked(&tag(&content));
        if items.iter().filter(|item| item.is_ok()).count() >= QUICK_XML_LINEAR_NAMES {
            let outcome = match items.last() {
                Some(Err(AttrError::Duplicated(..))) => "duplicate",
                Some(Err(AttrError::UnquotedValue(_))) => "unquoted value",
                Some(Err(AttrError::ExpectedValue(_))) => "no value",
                Some(Err(AttrError::ExpectedQuote(..))) => "no closing quote",
                Some(Err(AttrError::ExpectedEq(_))) => "no equals sign",
                _ => "end",
            };
            *late.entry(outcome).or_default() += 1;
        }
    }
    assert_eq!(late.len(), 6, "{late:?}");
    // The errors that consume the rest of the tag are the rarest.
    assert!(late.values().all(|count| *count >= 5), "{late:?}");
}

#[test]
fn nothing_is_yielded_after_the_first_error() {
    for content in [
        "e a=\"1\" a=\"2\" b=\"3\"",
        "e a=x b=\"2\"",
        "e flag b=\"2\"",
    ] {
        let tag = tag(content);
        let mut attributes = tag.checked_attributes();
        let error = attributes.find(Result::is_err);
        assert!(error.is_some(), "{content}");
        assert!(attributes.next().is_none(), "{content}");
        assert!(attributes.next().is_none(), "{content}");
    }
    let mut attributes = names(40);
    attributes.push("n3=\"again\"".to_owned());
    attributes.push("late=\"1\"".to_owned());
    let tag = tag(&format!("e {}", attributes.join(" ")));
    let items = checked(&tag);
    assert_eq!(items.len(), 41);
    assert!(matches!(items[40], Err(AttrError::Duplicated(..))));
    assert!(tag.checked_attributes().nth(41).is_none());
}

#[test]
fn a_single_attribute_checks_the_tail_and_then_stays_fused() {
    for content in ["e a=\"1\"", "e a='1' ", "e a = \"1\"\t\r\n"] {
        let tag = tag(content);
        let mut attributes = tag.checked_attributes();
        assert!(attributes.next().is_some(), "{content}");
        assert!(attributes.next().is_none(), "{content}");
        assert!(attributes.next().is_none(), "{content}");
    }
    for content in ["e", "e ", "e\t\r\n"] {
        let tag = tag(content);
        let mut attributes = tag.checked_attributes();
        assert!(attributes.next().is_none(), "{content}");
        assert!(attributes.next().is_none(), "{content}");
    }
}

#[test]
fn a_second_attribute_replays_the_first_before_checking_it() {
    for content in [
        "e a=\"1\" b=\"2\"",
        "e a=\"1\" a=\"2\"",
        "e a=\"1\" a=unquoted tail=\"ok\"",
        "e a=\"1\" b=",
    ] {
        assert_same(content);
        let tag = tag(content);
        let mut attributes = tag.checked_attributes();
        assert!(attributes.next().is_some(), "{content}");
        assert!(attributes.next().is_some(), "{content}");
        assert!(attributes.next().is_none(), "{content}");
        assert!(attributes.next().is_none(), "{content}");
    }
}

#[test]
fn clone_matches_from_before_and_after_the_replay_boundary() {
    for content in [
        "e",
        "e a=\"1\" ",
        "e a=\"1\" b=\"2\" c=\"3\"",
        "e a=\"1\" a=unquoted tail=\"ok\"",
    ] {
        for advance in [0, 1] {
            let tag = tag(content);
            let mut left = tag.checked_attributes();
            for _ in 0..advance {
                let _ = left.next();
            }
            let mut right = left.clone();
            loop {
                let left_item = left.next().map(owned);
                let right_item = right.next().map(owned);
                assert_eq!(left_item, right_item, "{content}, advance={advance}");
                if left_item.is_none() {
                    break;
                }
            }
            assert!(left.next().is_none(), "{content}, advance={advance}");
            assert!(right.next().is_none(), "{content}, advance={advance}");
        }
    }
}

#[test]
fn an_error_with_a_valid_recovery_tail_remains_the_terminal_item() {
    let content = "e a=\"1\" a=unquoted tail=\"ok\"";
    let tag = tag(content);
    let expected = quick_xml_until_error(&tag);
    assert!(matches!(
        expected.last(),
        Some(Err(AttrError::Duplicated(..)))
    ));
    let actual = checked(&tag);
    assert_eq!(actual, expected);
    let mut attributes = tag.checked_attributes();
    assert!(attributes.by_ref().any(|item| item.is_err()));
    assert!(attributes.next().is_none());
    assert!(attributes.next().is_none());
}

#[test]
#[allow(clippy::disallowed_methods)]
fn unchecked_attributes_are_quick_xml_without_its_check() {
    let tag = tag("e a=\"1\" a=\"2\" b=x c=\"3\" d");
    let mut expected = tag.attributes();
    expected.with_checks(false);
    let expected: Vec<Item> = expected.map(owned).collect();
    let actual: Vec<Item> = tag.unchecked_attributes().map(owned).collect();
    assert_eq!(actual, expected);
    assert_eq!(
        actual.iter().filter(|item| item.is_ok()).count(),
        3,
        "{actual:?}"
    );
}

/// The iterator finds where the next attribute starts from the end of the
/// last value yielded, which it can do because quick-xml yields values
/// borrowed from the tag.
#[test]
fn quick_xml_yields_values_borrowed_from_the_tag() {
    let tag = tag("e a=\"1\" b='two' c=\"\" d=\"&amp;\"");
    for attribute in tag.unchecked_attributes() {
        let attribute = attribute.unwrap();
        assert!(matches!(attribute.value, Cow::Borrowed(_)));
        let end = super::end_of(&tag, &attribute);
        assert!(matches!(tag.get(end - 1), Some(b'"' | b'\'')), "{end}");
    }
}

fn comparisons_for(content: &str) -> (usize, usize) {
    let tag = tag(content);
    let (items, comparisons) = counted(|| tag.checked_attributes().count());
    (items, comparisons)
}

fn flood(names: impl Iterator<Item = String>) -> String {
    let mut content = String::from("e");
    for name in names {
        content.push(' ');
        content.push_str(&name);
        content.push_str("=\"\"");
    }
    content
}

/// Upper bound on the comparisons an ordered map makes to insert `n` names:
/// every insertion descends at most `log2(n) + 1` levels and compares with at
/// most eleven keys on each (a B-tree node holds at most eleven).
fn ordered_bound(n: usize) -> usize {
    let levels = (usize::BITS - n.leading_zeros()) as usize + 1;
    n * 11 * levels
}

#[test]
fn a_tag_of_distinct_names_costs_n_log_n_comparisons_in_any_order() {
    for n in [1 << 10, 1 << 12, 1 << 14] {
        let shapes: [(&str, Vec<String>); 4] = [
            (
                "ascending",
                (0..n).map(|index| format!("n{index:06}")).collect(),
            ),
            (
                "descending",
                (0..n).rev().map(|index| format!("n{index:06}")).collect(),
            ),
            (
                "interleaved",
                (0..n)
                    .map(|index| {
                        let index = if index % 2 == 0 {
                            index / 2
                        } else {
                            n - 1 - index / 2
                        };
                        format!("n{index:06}")
                    })
                    .collect(),
            ),
            (
                "shared prefix",
                (0..n)
                    .map(|index| format!("{}{index:06}", "p".repeat(200)))
                    .collect(),
            ),
        ];
        for (shape, names) in shapes {
            let (items, comparisons) = comparisons_for(&flood(names.into_iter()));
            assert_eq!(items, n, "{shape}");
            assert!(
                comparisons <= ordered_bound(n),
                "{shape}, {n} names: {comparisons} comparisons, bound {}",
                ordered_bound(n)
            );
            // A scan of the earlier names would compare about n^2 / 2 times.
            if n == 1 << 14 {
                assert!(
                    comparisons * 40 < n * n / 2,
                    "{shape}, {n} names: {comparisons}"
                );
            }
        }
    }
}

#[test]
fn a_late_duplicate_costs_n_log_n_comparisons() {
    let n = 1 << 14;
    let mut content = flood((0..n).map(|index| format!("n{index:06}")));
    content.push_str(" n000000=\"again\"");
    let tag = tag(&content);
    let (items, comparisons) = counted(|| checked(&tag));
    assert_eq!(items.len(), n + 1);
    assert_eq!(
        items[n],
        Err(AttrError::Duplicated(content.len() - 15, 2)),
        "{:?}",
        items[n]
    );
    assert!(comparisons <= ordered_bound(n + 1), "{comparisons}");
}

#[test]
fn comparisons_grow_by_about_the_factor_of_names() {
    // Doubling the names doubles the insertions and deepens the map by at
    // most one level; a quadratic check would quadruple its work.
    let mut previous = None;
    for n in [1 << 11, 1 << 12, 1 << 13, 1 << 14, 1 << 15] {
        let (items, comparisons) =
            comparisons_for(&flood((0..n).map(|index| format!("n{index:06}"))));
        assert_eq!(items, n);
        if let Some(previous) = previous {
            assert!(
                comparisons * 10 <= previous * 25,
                "{n} names: {comparisons} comparisons after {previous}"
            );
        }
        previous = Some(comparisons);
    }
}

#[test]
fn quick_xml_checks_the_first_names_and_this_iterator_the_rest() {
    // Up to 32 names the map is never touched.
    let (items, comparisons) = comparisons_for(&format!("e {}", names(32).join(" ")));
    assert_eq!((items, comparisons), (32, 0));
    // The 33rd name records the first 32 and checks itself against them.
    let (items, comparisons) = comparisons_for(&format!("e {}", names(33).join(" ")));
    assert_eq!(items, 33);
    assert!(
        comparisons > 0 && comparisons <= ordered_bound(33),
        "{comparisons}"
    );
    assert_eq!(QUICK_XML_LINEAR_NAMES, 32);
}

/// The crates that may not depend on `litchi-opc` keep copies of its
/// `xml_attributes` module, and each compiles this file as its tests. Their
/// code must be the canonical code: only the module docs, the visibility of
/// the two public items and the path of this test module may differ.
#[test]
fn every_copy_of_the_module_is_the_canonical_code() {
    fn code(text: &str) -> String {
        let start = text.find("use std::borrow::Cow;").expect("module imports");
        text[start..]
            .replace("pub(crate) trait BytesStartExt", "pub trait BytesStartExt")
            .replace(
                "pub(crate) struct CheckedAttributes",
                "pub struct CheckedAttributes",
            )
            .replace(
                "#[path = \"../../litchi-opc/src/xml_attributes/tests.rs\"]\n",
                "",
            )
    }
    let crates = concat!(env!("CARGO_MANIFEST_DIR"), "/..");
    let canonical = std::fs::read_to_string(format!("{crates}/litchi-opc/src/xml_attributes.rs"))
        .expect("canonical module");
    for copy in [
        "litchi-ole-common",
        "litchi-sign",
        "litchi-xldm",
        "xml-minifier",
    ] {
        let text = std::fs::read_to_string(format!("{crates}/{copy}/src/xml_attributes.rs"))
            .expect("module copy");
        assert!(code(&text) == code(&canonical), "{copy} is out of step");
    }
}
