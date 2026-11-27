//! Property-style round-trip and robustness tests for the path/selector
//! grammar.
//!
//! The workspace has no `proptest` dependency, so these tests drive a seeded
//! deterministic generator: random typed ASTs must satisfy
//! `parse(display(ast)) == ast`, and randomly mutated corpus strings must
//! either fail with a typed error or parse to a value that round-trips.

use std::borrow::Cow;

use litchi_core::selector::{
    Combinator, CompareOp, Complex, Compound, Condition, Element, Expr, Key, Path, Position,
    Predicate, Pseudo, Query, Scope, Step, Value,
};

/// SplitMix64: tiny, deterministic, good enough for test-case generation.
struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9e37_79b9_7f4a_7c15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
        z ^ (z >> 31)
    }

    fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }

    fn chance(&mut self, percent: usize) -> bool {
        self.below(100) < percent
    }

    fn pick<'p, T>(&mut self, pool: &'p [T]) -> &'p T {
        &pool[self.below(pool.len())]
    }
}

const NAMES: &[&str] = &[
    "body",
    "p",
    "r",
    "slide",
    "shape",
    "Sheet1",
    "Sheet 1",
    "1-Basics & Quartile",
    "P&L",
    "O'Brien",
    "a/b",
    "x[1]",
    "'q",
    "\"d",
    "",
    "say \"hi\"",
    "A1",
    "A1:C5",
    "金额",
    "back\\slash",
    "tab\there",
    "My Data (2024)",
    "(open",
    "close)",
    "header",
    "notes",
    "@lead",
    " spaced ",
];
const KEYS: &[&str] = &[
    "a",
    "style",
    "Full Name",
    "12",
    "and",
    "or",
    "not",
    "last",
    "",
    "金额",
    "x.y",
    "a\"b",
    "paraId",
    "col.Salary",
    "a b",
    "_x",
];
const TEXTS: &[&str] = &[
    "1",
    "Heading 1",
    "a]b",
    "it's",
    "x*",
    "==",
    "~x",
    "!x",
    "back\\",
    "and",
    "or",
    "a and b",
    " pad ",
    "#FF0000",
    "r\"x\"",
    "",
    "金额",
    "x/y",
    "f(x)",
    "5000",
    "true",
    "a,b",
    "tab\tin",
];
const PATTERNS: &[&str] = &["^Q[1-4]", "a\"b", "it's", "\\d+", "x\\\"y", "[a-z]*"];
const ELEMENTS: &[&str] = &[
    "p", "cell", "B", "A1", "A1:C5", "3dmodel", "row", "_x", "no-alt", "tc",
];
const SHEETS: &[&str] = &[
    "Sheet1",
    "My Data (2024)",
    "O'Brien",
    "1-Hero",
    "a!b",
    "",
    "金额",
    "x.y",
];
const CONTAINS: &[&str] = &[
    "TODO",
    "a/b",
    "Use of Funds",
    " pad ",
    "#REF!",
    "x)y",
    "q\"t",
    "",
    "back\\slash",
    "(x",
];

fn key(rng: &mut Rng) -> Key<'static> {
    let name = *rng.pick(KEYS);
    if rng.chance(30) {
        Key::attribute(name)
    } else {
        Key::new(name)
    }
}

fn condition(rng: &mut Rng) -> Condition<'static> {
    if rng.chance(20) {
        return Condition::Exists(key(rng));
    }
    if rng.chance(15) {
        return Condition::Compare {
            key: key(rng),
            op: CompareOp::Contains,
            value: Value::regex(*rng.pick(PATTERNS)).expect("representable pattern"),
        };
    }
    let op = *rng.pick(&[
        CompareOp::Eq,
        CompareOp::Ne,
        CompareOp::Contains,
        CompareOp::Gt,
        CompareOp::Lt,
        CompareOp::Ge,
        CompareOp::Le,
    ]);
    Condition::Compare {
        key: key(rng),
        op,
        value: Value::text(*rng.pick(TEXTS)),
    }
}

fn expr(rng: &mut Rng, depth: usize) -> Expr<'static> {
    if depth == 0 || rng.chance(50) {
        return Expr::Condition(condition(rng));
    }
    match rng.below(3) {
        0 => Expr::And(
            (0..2 + rng.below(3))
                .map(|_| expr(rng, depth - 1))
                .collect(),
        ),
        1 => Expr::Or(
            (0..2 + rng.below(3))
                .map(|_| expr(rng, depth - 1))
                .collect(),
        ),
        _ => Expr::Not(Box::new(expr(rng, depth - 1))),
    }
}

fn predicate(rng: &mut Rng) -> Predicate<'static> {
    match rng.below(4) {
        0 => Predicate::Index(Position::new(if rng.chance(10) {
            usize::MAX - 1
        } else {
            rng.below(1000)
        })),
        1 => Predicate::Last,
        _ => Predicate::Filter(expr(rng, 3)),
    }
}

fn step(rng: &mut Rng) -> Step<'static> {
    let mut step = Step::new(*rng.pick(NAMES));
    for _ in 0..rng.below(4) {
        step = step.with(predicate(rng));
    }
    step
}

fn path(rng: &mut Rng) -> Path<'static> {
    let count = rng.below(9);
    let steps = (0..count).map(|_| step(rng)).collect::<Vec<_>>();
    if count > 0 && rng.chance(20) {
        Path::relative(steps)
    } else {
        Path::from_steps(steps)
    }
}

fn compound(rng: &mut Rng, depth: usize) -> Compound<'static> {
    let element = match rng.below(4) {
        0 => Element::Implicit,
        1 => Element::Any,
        _ => Element::Named(Cow::Borrowed(*rng.pick(ELEMENTS))),
    };
    let mut compound = Compound::new(element.clone());
    for _ in 0..rng.below(3) {
        compound = compound.with(predicate(rng));
    }
    for _ in 0..rng.below(3) {
        compound = compound.with_pseudo(pseudo(rng, depth));
    }
    if element == Element::Implicit
        && compound.predicates().is_empty()
        && compound.pseudos().is_empty()
    {
        compound = compound.with_pseudo(Pseudo::Empty);
    }
    compound
}

fn pseudo(rng: &mut Rng, depth: usize) -> Pseudo<'static> {
    match rng.below(if depth == 0 { 3 } else { 5 }) {
        0 => Pseudo::Contains(Cow::Borrowed(*rng.pick(CONTAINS))),
        1 => Pseudo::Empty,
        2 => Pseudo::NoAlt,
        3 => Pseudo::Has(list(rng, depth - 1)),
        _ => Pseudo::Not(list(rng, depth - 1)),
    }
}

fn complex(rng: &mut Rng, depth: usize) -> Complex<'static> {
    let mut complex = Complex::new(compound(rng, depth));
    for _ in 0..rng.below(3) {
        let combinator = if rng.chance(50) {
            Combinator::Child
        } else {
            Combinator::Descendant
        };
        complex = complex.then(combinator, compound(rng, depth));
    }
    complex
}

fn list(rng: &mut Rng, depth: usize) -> Vec<Complex<'static>> {
    (0..1 + rng.below(3)).map(|_| complex(rng, depth)).collect()
}

fn query(rng: &mut Rng) -> Query<'static> {
    let query = Query::new(list(rng, 2));
    match rng.below(3) {
        0 => query,
        1 => query.with_scope(Scope::Sheet(Cow::Borrowed(*rng.pick(SHEETS)))),
        _ => {
            let steps = (0..rng.below(4)).map(|_| step(rng)).collect::<Vec<_>>();
            query.with_scope(Scope::Path(Box::new(Path::from_steps(steps))))
        },
    }
}

#[test]
fn random_paths_round_trip() {
    let mut rng = Rng(0x5eed_0005);
    for case in 0..20_000 {
        let original = path(&mut rng);
        let shown = original.to_string();
        let parsed = if original.is_absolute() {
            Path::parse(&shown)
        } else {
            Path::parse_relative(&shown)
        }
        .unwrap_or_else(|error| panic!("case {case}: {shown:?}: {error}\n{original:#?}"));
        assert_eq!(parsed, original, "case {case}: {shown:?}");
        assert_eq!(
            parsed.to_string(),
            shown,
            "case {case}: display is canonical"
        );
    }
}

#[test]
fn random_queries_round_trip() {
    let mut rng = Rng(0x5eed_0b0b);
    for case in 0..20_000 {
        let original = query(&mut rng);
        let shown = original.to_string();
        let parsed = Query::parse(&shown)
            .unwrap_or_else(|error| panic!("case {case}: {shown:?}: {error}\n{original:#?}"));
        assert_eq!(parsed, original, "case {case}: {shown:?}");
        assert_eq!(
            parsed.to_string(),
            shown,
            "case {case}: display is canonical"
        );
    }
}

/// Grammar-significant characters plus a few ordinary ones.
const ALPHABET: &[char] = &[
    '/', '[', ']', '(', ')', '"', '\'', '\\', '@', '=', '!', '~', '<', '>', ':', ',', ' ', '*',
    '+', 'a', 'n', 'd', 'o', 'r', 't', '1', '0', '-', 'r', '金', '\u{1}', '\t',
];

fn mutate(rng: &mut Rng, seed: &str) -> String {
    let mut chars: Vec<char> = seed.chars().collect();
    for _ in 0..1 + rng.below(4) {
        let at = if chars.is_empty() {
            0
        } else {
            rng.below(chars.len() + 1)
        };
        match rng.below(3) {
            0 => chars.insert(at, *rng.pick(ALPHABET)),
            1 if at < chars.len() => {
                chars.remove(at);
            },
            _ if at < chars.len() => chars[at] = *rng.pick(ALPHABET),
            _ => chars.push(*rng.pick(ALPHABET)),
        }
    }
    chars.into_iter().collect()
}

fn corpus() -> Vec<&'static str> {
    include_str!("fixtures/officecli_paths.txt")
        .lines()
        .chain(include_str!("fixtures/officecli_selectors.txt").lines())
        .filter(|line| !line.is_empty() && !line.starts_with('#'))
        .collect()
}

#[test]
fn mutated_corpus_inputs_fail_typed_or_round_trip() {
    let seeds = corpus();
    let mut rng = Rng(0x5eed_f022);
    let (mut accepted, mut rejected) = (0usize, 0usize);
    for _ in 0..60_000 {
        let seed = *rng.pick(&seeds);
        let input = mutate(&mut rng, seed);
        match Path::parse(&input) {
            Ok(path) => {
                accepted += 1;
                let shown = path.to_string();
                assert_eq!(
                    Path::parse(&shown).as_ref(),
                    Ok(&path),
                    "{input:?} -> {shown:?}"
                );
            },
            Err(error) => {
                rejected += 1;
                assert!(error.offset() <= input.len(), "{input:?}: {error}");
                assert!(input.is_char_boundary(error.offset()), "{input:?}: {error}");
            },
        }
        match Query::parse(&input) {
            Ok(query) => {
                let shown = query.to_string();
                assert_eq!(
                    Query::parse(&shown).as_ref(),
                    Ok(&query),
                    "{input:?} -> {shown:?}"
                );
            },
            Err(error) => {
                assert!(error.offset() <= input.len(), "{input:?}: {error}");
                assert!(input.is_char_boundary(error.offset()), "{input:?}: {error}");
            },
        }
    }
    // Both outcomes are exercised.
    assert!(
        accepted > 1_000 && rejected > 1_000,
        "{accepted} / {rejected}"
    );
}

#[test]
fn hostile_inputs_are_bounded() {
    // Deep nesting, long unions, and oversized input fail with typed errors
    // without exhausting the stack.
    for input in [
        "/a[".to_owned() + &"(".repeat(100_000),
        "a".to_owned() + &":not(".repeat(100_000),
        "a".to_owned() + &",a".repeat(10_000),
        "/a".repeat(10_000),
        "x".repeat(1 << 20),
        "/a[not(".repeat(5_000),
        "/".to_owned() + &"金".repeat(10_000),
    ] {
        for error in [
            Path::parse(&input).unwrap_err(),
            Query::parse(&input).unwrap_err(),
        ] {
            assert!(input.is_char_boundary(error.offset()), "{error}");
        }
    }
}
