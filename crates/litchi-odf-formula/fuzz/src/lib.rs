use litchi_odf_formula::{Formula, codec};

const NS: &str = "http://www.w3.org/1998/Math/MathML";

/// Exercise raw XML, raw package, generated-valid, or generated-invalid input.
///
/// The low two bits of the first byte select the mode. Structured modes decode
/// the remaining bytes through an independent MathML 2 signature grammar so
/// coverage guidance can mutate tree shape without first discovering XML.
pub fn exercise(data: &[u8]) {
    let Some((&selector, payload)) = data.split_first() else {
        exercise_package(data);
        return;
    };
    match selector & 3 {
        0 => exercise_xml(payload),
        1 => exercise_package(payload),
        2 => exercise_generated_valid(payload),
        _ => exercise_generated_invalid(payload),
    }
}

fn exercise_generated_invalid(data: &[u8]) {
    let xml = Oracle::new(data).invalid_document();
    assert!(
        Formula::create(&xml).is_err(),
        "independent grammar breaker was accepted: {xml}"
    );
}

fn exercise_generated_valid(data: &[u8]) {
    let xml = Oracle::new(data).valid_document();
    let root = codec::parse(&xml)
        .unwrap_or_else(|error| panic!("generated valid XML failed to parse: {error}: {xml}"));
    codec::validate(&root)
        .unwrap_or_else(|error| panic!("generated valid tree failed validation: {error}: {xml}"));
    let compact = codec::serialize(&root);
    assert!(!compact.contains('\n'));
    let reparsed = codec::parse(&compact).expect("serialized generated tree must parse");
    assert_eq!(reparsed, root);
    let formula = Formula::create(&xml).expect("generated valid Formula must package");
    assert_eq!(
        formula
            .content_xml()
            .expect("generated package content must read"),
        xml
    );
    let reopened = Formula::from_bytes(formula.to_bytes()).expect("generated package must reopen");
    assert_eq!(reopened.root(), &root);
}

fn exercise_package(data: &[u8]) {
    if let Ok(formula) = Formula::from_bytes(data.to_vec()) {
        assert_eq!(formula.as_bytes(), data);
        let content = formula
            .content_xml()
            .expect("accepted package content must read");
        let parsed = codec::parse(&content).expect("accepted package content must parse");
        codec::validate(&parsed).expect("accepted package content must validate");
        assert_eq!(formula.root(), &parsed);
        let reopened =
            Formula::from_bytes(formula.to_bytes()).expect("accepted package must reopen");
        assert_eq!(reopened.as_bytes(), data);
        assert_eq!(reopened.root(), formula.root());
    }
}

fn exercise_xml(data: &[u8]) {
    let Ok(xml) = std::str::from_utf8(data) else {
        return;
    };
    let Ok(root) = codec::parse(xml) else {
        return;
    };
    let valid = codec::validate(&root).is_ok();
    let compact = codec::serialize(&root);
    let reparsed = codec::parse(&compact).expect("serialized parsed tree must parse");
    assert_eq!(reparsed, root);
    assert_eq!(codec::validate(&reparsed).is_ok(), valid);
    if valid {
        let formula = Formula::create(&compact).expect("compact validated XML must package");
        assert_eq!(
            formula
                .content_xml()
                .expect("authored package content must read"),
            compact
        );
        let reopened =
            Formula::from_bytes(formula.to_bytes()).expect("authored package must reopen");
        assert_eq!(reopened.root(), &root);
    }
}

struct Oracle<'data> {
    data: &'data [u8],
    position: usize,
}

impl<'data> Oracle<'data> {
    const fn new(data: &'data [u8]) -> Self {
        Self { data, position: 0 }
    }

    fn collection(&mut self, name: &str, depth: usize) -> String {
        if self.bounded(4) == 0 {
            return format!(
                "<{name}><bvar><ci>x</ci></bvar><condition><true/></condition><ci>x</ci></{name}>"
            );
        }
        let minimum = usize::from(name == "vector");
        let count = self.bounded(4).saturating_add(minimum);
        let mut children = String::new();
        for _index in 0..count {
            children.push_str(&self.expression(depth.saturating_sub(1)));
        }
        format!("<{name}>{children}</{name}>")
    }

    fn expression(&mut self, depth: usize) -> String {
        if depth == 0 {
            return self.leaf();
        }
        match self.bounded(17) {
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
                format!("<lambda><bvar><ci>x</ci></bvar>{body}</lambda>")
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
                    r#"<semantics>{primary}<annotation encoding="application/x-fuzz-oracle">opaque</annotation></semantics>"#
                )
            },
            13 => {
                let argument = self.expression(depth - 1);
                format!("<apply><ci type=\"function\">f</ci>{argument}</apply>")
            },
            14 => {
                let body = self.expression(depth - 1);
                format!(
                    "<apply><divergence/><bvar><ci>x</ci></bvar><bvar><ci>y</ci></bvar>{body}</apply>"
                )
            },
            15 => {
                let body = self.expression(depth - 1);
                format!(
                    "<apply><sum/><bvar><ci>i</ci></bvar><lowlimit><cn>0</cn></lowlimit><uplimit><cn>9</cn></uplimit>{body}</apply>"
                )
            },
            _ => {
                let value = self.expression(depth - 1);
                format!(
                    r#"<semantics>{value}<annotation encoding="StarMath 6">x + 1</annotation></semantics>"#
                )
            },
        }
    }

    fn invalid_document(mut self) -> String {
        let case = self.bounded(14);
        let body = match case {
            0 => "<apply><sin/></apply>",
            1 => "<apply><power/><ci>x</ci></apply>",
            2 => "<cn type=\"e-notation\">2</cn>",
            3 => "<cn type=\"integer\" base=\"37\">0</cn>",
            4 => "<interval><condition><true/></condition><bvar><ci>x</ci></bvar></interval>",
            5 => {
                "<matrix><matrixrow><cn>1</cn></matrixrow><matrixrow><cn>2</cn><cn>3</cn></matrixrow></matrix>"
            },
            6 => "<vector/>",
            7 => "<piecewise><piece><cn>1</cn><cn>0</cn></piece></piecewise>",
            8 => {
                "<apply><sum/><uplimit><cn>9</cn></uplimit><lowlimit><cn>0</cn></lowlimit><ci>x</ci></apply>"
            },
            9 => "<declare><ci>x</ci><cn>1</cn><cn>2</cn></declare>",
            10 => "<mrow><ci>x</ci></mrow>",
            11 => "<set><mi>x</mi></set>",
            12 => "<apply><divergence/><bvar><ci>x</ci></bvar></apply>",
            _ => {
                "<semantics><mi>x</mi><annotation encoding=\"StarMath 6\"><mi>active</mi></annotation></semantics>"
            },
        };
        format!(r#"<math xmlns="{NS}">{body}</math>"#)
    }

    fn leaf(&mut self) -> String {
        match self.bounded(10) {
            0 => "<ci>x</ci>".to_string(),
            1 => "<cn type=\"integer\" base=\"16\">a</cn>".to_string(),
            2 => "<cn type=\"e-notation\">2<sep/>5</cn>".to_string(),
            3 => "<cn type=\"rational\">1<sep/>3</cn>".to_string(),
            4 => "<csymbol definitionURL=\"urn:fuzz\">f</csymbol>".to_string(),
            5 => "<true/>".to_string(),
            6 => "<reals/>".to_string(),
            7 => "<pi/>".to_string(),
            8 => "<ci><msub><mi>x</mi><mn>1</mn></msub></ci>".to_string(),
            _ => "<cn>1</cn>".to_string(),
        }
    }

    fn matrix(&mut self, depth: usize) -> String {
        let height = self.bounded(3).saturating_add(1);
        let width = self.bounded(3).saturating_add(1);
        let mut rows = String::new();
        for _row in 0..height {
            rows.push_str("<matrixrow>");
            for _column in 0..width {
                rows.push_str(&self.expression(depth.saturating_sub(1)));
            }
            rows.push_str("</matrixrow>");
        }
        format!("<matrix>{rows}</matrix>")
    }

    fn next(&mut self) -> u8 {
        let value = self.data.get(self.position).copied().unwrap_or_else(|| {
            self.position
                .wrapping_mul(73)
                .wrapping_add(41)
                .to_le_bytes()[0]
        });
        self.position = self.position.saturating_add(1);
        value
    }

    fn bounded(&mut self, upper: usize) -> usize {
        usize::from(self.next()) % upper
    }

    fn valid_document(mut self) -> String {
        let display = if self.next() & 1 == 0 {
            "inline"
        } else {
            "block"
        };
        let declaration = match self.bounded(3) {
            0 => "",
            1 => "<declare><plus/></declare>",
            _ => "<declare><ci>y</ci><apply><plus/><ci>x</ci><cn>3</cn></apply></declare>",
        };
        let depth = self.bounded(4).saturating_add(1);
        let expression = self.expression(depth);
        format!(r#"<math xmlns="{NS}" display="{display}">{declaration}{expression}</math>"#)
    }
}
