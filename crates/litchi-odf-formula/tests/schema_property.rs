#![allow(
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected failures"
)]

use litchi_odf_common::core::OwnedPackage;
use litchi_odf_formula::{Formula, NodePath, OpaqueStarMath, Patch, StarMathVersion, codec};

const NS: &str = "http://www.w3.org/1998/Math/MathML";

#[test]
fn w3c_content_signatures_accept_broader_safe_forms() {
    // These are independent examples derived from the MathML 2 Appendix C
    // signatures and the W3C MathML 2 XML Schema modules, not trees produced
    // by this crate's constructors.
    let schema_examples = [
        format!(
            r#"<math xmlns="{NS}"><cn type="e-notation">2<sep/>5</cn><cn type="integer" base="36">z</cn></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><declare><plus/></declare><declare><ci>y</ci><apply><plus/><ci>x</ci><cn>3</cn></apply></declare></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><set><bvar><ci>x</ci></bvar><condition><true/></condition><ci>x</ci></set></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><interval closure="open"><bvar><ci>x</ci></bvar><condition><apply><lt/><cn>0</cn><ci>x</ci></apply></condition></interval></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><apply><divergence/><bvar><ci>x</ci></bvar><bvar><ci>y</ci></bvar><vector><ci>f</ci><ci>g</ci></vector></apply></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><apply><ci type="function">f</ci><ci>x</ci></apply><apply><csymbol definitionURL="urn:operator">op</csymbol><cn>1</cn><cn>2</cn></apply></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><matrix><matrixrow><cn>1</cn><cn>2</cn></matrixrow><matrixrow><cn>3</cn><cn>4</cn></matrixrow></matrix></math>"#
        ),
    ];

    for xml in schema_examples {
        let formula = Formula::create(&xml).expect("W3C schema/signature example");
        let reopened = Formula::from_bytes(formula.to_bytes()).expect("full package reopen");
        codec::validate(reopened.root()).expect("reopened schema projection");
    }
}

#[test]
fn independent_recursive_grammar_satisfies_parse_package_and_reopen_properties() {
    let mut grammar = Grammar::new(0x6a09_e667_f3bc_c909);
    for case in 0..1_024_usize {
        let declaration = if case.is_multiple_of(3) {
            "<declare><plus/></declare>"
        } else {
            ""
        };
        let expression = grammar.expression(3);
        let xml = format!(r#"<math xmlns="{NS}">{declaration}{expression}</math>"#);
        let parsed = codec::parse(&xml).unwrap_or_else(|error| {
            panic!("independently generated XML case {case} failed: {error}: {xml}")
        });
        codec::validate(&parsed).expect("independent grammar must satisfy the checked projection");
        let compact = codec::serialize(&parsed);
        assert!(!compact.contains('\n'));
        let formula = Formula::create(&xml).expect("generated package");
        let reopened = Formula::from_bytes(formula.to_bytes()).expect("generated full reopen");
        assert_eq!(reopened.root(), &parsed, "generated case {case}");
    }
}

#[test]
fn independent_schema_breakers_reject_one_rule_at_a_time() {
    let unary = ["abs", "sin", "factorial", "transpose"];
    let binary = ["divide", "power", "quotient", "vectorproduct"];
    let mut malformed = Vec::new();
    for operator in unary {
        malformed.push(format!(
            r#"<math xmlns="{NS}"><apply><{operator}/></apply></math>"#
        ));
        malformed.push(format!(
            r#"<math xmlns="{NS}"><apply><{operator}/><ci>x</ci><ci>y</ci></apply></math>"#
        ));
    }
    for operator in binary {
        malformed.push(format!(
            r#"<math xmlns="{NS}"><apply><{operator}/><ci>x</ci></apply></math>"#
        ));
        malformed.push(format!(
            r#"<math xmlns="{NS}"><apply><{operator}/><ci>x</ci><ci>y</ci><ci>z</ci></apply></math>"#
        ));
    }
    malformed.extend([
        format!(r#"<math xmlns="{NS}"><cn type="e-notation">2</cn></math>"#),
        format!(
            r#"<math xmlns="{NS}"><cn type="e-notation">2<sep/>5<sep/>7</cn></math>"#
        ),
        format!(r#"<math xmlns="{NS}"><cn type="integer" base="1">0</cn></math>"#),
        format!(r#"<math xmlns="{NS}"><cn type="integer" base="37">0</cn></math>"#),
        format!(
            r#"<math xmlns="{NS}"><interval><condition><true/></condition><bvar><ci>x</ci></bvar></interval></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><interval><bvar><ci>x</ci></bvar></interval></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><set><bvar><ci>x</ci></bvar><condition><mi>x</mi></condition><ci>x</ci></set></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><declare><ci>x</ci><cn>1</cn><cn>2</cn></declare></math>"#
        ),
        format!(r#"<math xmlns="{NS}"><declare><mi>x</mi></declare></math>"#),
        format!(r#"<math xmlns="{NS}"><matrix/></math>"#),
        format!(r#"<math xmlns="{NS}"><matrix><matrixrow/></matrix></math>"#),
        format!(r#"<math xmlns="{NS}"><vector/></math>"#),
        format!(
            r#"<math xmlns="{NS}"><matrix><matrixrow><cn>1</cn></matrixrow><matrixrow><cn>2</cn><cn>3</cn></matrixrow></matrix></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><apply><divergence/><bvar><ci>x</ci></bvar></apply></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><apply><sum/><uplimit><cn>2</cn></uplimit><lowlimit><cn>0</cn></lowlimit><ci>x</ci></apply></math>"#
        ),
    ]);

    for (case, xml) in malformed.into_iter().enumerate() {
        let first = Formula::create(&xml).is_err();
        let second = Formula::create(&xml).is_err();
        assert!(
            first && second,
            "accepted independent breaker {case}: {xml}"
        );
    }
}

#[test]
fn arbitrary_byte_stream_fuzz_corpus_is_panic_free_and_repeatable() {
    let mut state = 0xbb67_ae85_84ca_a73b_u64;
    for case in 0..4_096_usize {
        state = xorshift(state);
        let length = usize::try_from(state & 0x7ff).unwrap_or(0);
        let mut bytes = Vec::with_capacity(length);
        for _index in 0..length {
            state = xorshift(state);
            bytes.push(state.to_le_bytes()[0]);
        }

        let package_first = Formula::from_bytes(bytes.clone()).is_ok();
        let package_second = Formula::from_bytes(bytes.clone()).is_ok();
        assert_eq!(package_first, package_second, "package fuzz case {case}");
        if let Ok(xml) = std::str::from_utf8(&bytes) {
            let parse_first = codec::parse(xml).is_ok();
            let parse_second = codec::parse(xml).is_ok();
            assert_eq!(parse_first, parse_second, "XML fuzz case {case}");
            let formula_first = Formula::create(xml).is_ok();
            let formula_second = Formula::create(xml).is_ok();
            assert_eq!(formula_first, formula_second, "formula fuzz case {case}");
        }
    }
}

#[test]
fn second_changed_libreoffice_package_reopens_with_auxiliary_payloads_exact() {
    let bytes = include_bytes!(
        "../../../3rdparty/libreoffice-core/starmath/qa/cppunit/data/font-styles.odf"
    );
    let source = Formula::from_bytes(bytes.to_vec()).expect("LibreOffice source");
    let original = OwnedPackage::from_bytes(bytes.to_vec()).expect("raw source package");

    let mut presentation_edit = source.edit();
    presentation_edit
        .set_text(&NodePath::new([0, 0, 0]), "g")
        .expect("presentation token");
    let mut annotation_edit = source.edit();
    let opaque = OpaqueStarMath::new(StarMathVersion::V6, "g(x)").expect("opaque StarMath");
    annotation_edit
        .set_starmath(&opaque)
        .expect("opaque annotation");
    presentation_edit
        .join(&annotation_edit)
        .expect("independent edit join");
    let commit = presentation_edit
        .commit()
        .expect("changed producer package");
    assert!(commit.diagnostics().candidate_reopened());
    let durable = Patch::from_bytes(&commit.patch().to_bytes().expect("patch bytes"))
        .expect("durable patch reopen");
    let changed = commit.into_formula();
    let reopened = Formula::from_bytes(changed.to_bytes()).expect("changed Formula reopen");
    assert_eq!(reopened.starmath().expect("StarMath").opaque(), &opaque);
    assert_eq!(
        durable.apply(&source).expect("durable replay").as_bytes(),
        reopened.as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(&reopened)
            .expect("inverse")
            .as_bytes(),
        source.as_bytes()
    );

    let changed_package =
        OwnedPackage::from_bytes(reopened.to_bytes()).expect("raw changed package");
    let mut original_names = original.files().expect("source members");
    let mut changed_names = changed_package.files().expect("changed members");
    original_names.retain(|name| !name.ends_with('/'));
    changed_names.retain(|name| !name.ends_with('/'));
    original_names.sort_unstable();
    changed_names.sort_unstable();
    assert_eq!(changed_names, original_names);
    for name in original_names {
        if !matches!(name.as_str(), "content.xml" | "META-INF/manifest.xml") {
            assert_eq!(
                changed_package.get_file(&name).expect("changed member"),
                original.get_file(&name).expect("source member"),
                "member {name} changed"
            );
        }
    }
}

struct Grammar {
    state: u64,
}

impl Grammar {
    const fn new(state: u64) -> Self {
        Self { state }
    }

    fn expression(&mut self, depth: usize) -> String {
        if depth == 0 {
            return self.leaf();
        }
        match self.bounded(14) {
            0 => {
                let argument = self.expression(depth - 1);
                format!("<apply><sin/>{argument}</apply>")
            },
            1 => {
                let left = self.expression(depth - 1);
                let right = self.expression(depth - 1);
                format!("<apply><power/>{left}{right}</apply>")
            },
            2 => {
                let count = self.bounded(4);
                let mut arguments = String::new();
                for _index in 0..count {
                    arguments.push_str(&self.expression(depth - 1));
                }
                format!("<apply><plus/>{arguments}</apply>")
            },
            3 => {
                let left = self.expression(depth - 1);
                let right = self.expression(depth - 1);
                format!("<apply><lt/>{left}{right}</apply>")
            },
            4 => {
                let lower = self.expression(depth - 1);
                let upper = self.expression(depth - 1);
                format!(r#"<interval closure="closed-open">{lower}{upper}</interval>"#)
            },
            5 => "<interval><bvar><ci>x</ci></bvar><condition><true/></condition></interval>"
                .to_string(),
            6 => {
                let body = self.expression(depth - 1);
                format!("<lambda><bvar><ci>x</ci></bvar><apply><plus/>{body}</apply></lambda>")
            },
            7 => self.collection("set", depth),
            8 => self.collection("list", depth),
            9 => self.collection("vector", depth),
            10 => self.matrix(depth),
            11 => {
                let value = self.expression(depth - 1);
                format!(
                    "<piecewise><piece>{value}<true/></piece><otherwise><cn>0</cn></otherwise></piecewise>"
                )
            },
            12 => {
                let primary = self.expression(depth - 1);
                format!(
                    r#"<semantics>{primary}<annotation encoding="application/x-grammar">opaque</annotation></semantics>"#
                )
            },
            _ => {
                let argument = self.expression(depth - 1);
                format!("<apply><ci type=\"function\">f</ci>{argument}</apply>")
            },
        }
    }

    fn leaf(&mut self) -> String {
        match self.bounded(8) {
            0 => "<ci>x</ci>".to_string(),
            1 => "<cn type=\"integer\" base=\"16\">a</cn>".to_string(),
            2 => "<cn type=\"e-notation\">2<sep/>5</cn>".to_string(),
            3 => "<csymbol definitionURL=\"urn:grammar\">f</csymbol>".to_string(),
            4 => "<true/>".to_string(),
            5 => "<reals/>".to_string(),
            6 => "<pi/>".to_string(),
            _ => "<cn>1</cn>".to_string(),
        }
    }

    fn collection(&mut self, name: &str, depth: usize) -> String {
        if self.bounded(3) == 0 {
            return format!(
                "<{name}><bvar><ci>x</ci></bvar><condition><true/></condition><ci>x</ci></{name}>"
            );
        }
        let count = self
            .bounded(4)
            .saturating_add(usize::from(name == "vector"));
        let mut children = String::new();
        for _index in 0..count {
            children.push_str(&self.expression(depth - 1));
        }
        format!("<{name}>{children}</{name}>")
    }

    fn matrix(&mut self, depth: usize) -> String {
        let height = self.bounded(3).saturating_add(1);
        let width = self.bounded(3).saturating_add(1);
        let mut rows = String::new();
        for _row in 0..height {
            rows.push_str("<matrixrow>");
            for _column in 0..width {
                rows.push_str(&self.expression(depth - 1));
            }
            rows.push_str("</matrixrow>");
        }
        format!("<matrix>{rows}</matrix>")
    }

    fn bounded(&mut self, upper: usize) -> usize {
        self.state = xorshift(self.state);
        usize::try_from(self.state).unwrap_or(0) % upper
    }
}

const fn xorshift(mut state: u64) -> u64 {
    state ^= state << 13;
    state ^= state >> 7;
    state ^ (state << 17)
}
