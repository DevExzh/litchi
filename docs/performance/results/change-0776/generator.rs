use std::fmt::Write as _;
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
fn generated_document(seed: u64) -> String {
    let mut random = Xorshift(seed.wrapping_mul(0x9e37_79b9_7f4a_7c15) | 1);
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:w="urn:w" xmlns:i="urn:i" xmlns:j="urn:j""#);
    match random.below(4) {
        0 => xml.push_str(r#" mc:Ignorable="i j" mc:ProcessContent="i:u j:*" mc:PreserveElements="i:keep" mc:PreserveAttributes="i:*""#),
        1 => xml.push_str(r#" mc:Ignorable="i" mc:PreserveAttributes="i:k""#),
        2 => xml.push_str(r#" mc:Ignorable="j""#),
        _ => {},
    }
    xml.push('>');
    children(&mut random, &mut xml, 0);
    xml.push_str("</r>");
    xml
}

struct Xorshift(u64);

impl Xorshift {
    fn next(&mut self) -> u64 {
        let mut value = self.0;
        value ^= value << 13;
        value ^= value >> 7;
        value ^= value << 17;
        self.0 = value;
        value
    }

    fn below(&mut self, bound: u64) -> u64 {
        self.next() % bound
    }
}

const PREFIXES: [&str; 6] = ["p0", "p1", "p2", "i", "j", "w"];

fn declarations(random: &mut Xorshift, xml: &mut String, depth: usize) {
    let count = random.below(4);
    let mut used = Vec::new();
    for _ in 0..count {
        let prefix = PREFIXES[random.below(PREFIXES.len() as u64) as usize];
        if used.contains(&prefix) {
            continue;
        }
        used.push(prefix);
        let namespace = match random.below(4) {
            0 => "urn:i".to_owned(),
            1 => "urn:w".to_owned(),
            2 => "urn:j".to_owned(),
            _ => format!("urn:{prefix}:{depth}"),
        };
        let _ = write!(xml, r#" xmlns:{prefix}="{namespace}""#);
    }
    match random.below(8) {
        0 => xml.push_str(r#" xmlns="urn:d""#),
        1 => xml.push_str(r#" xmlns="""#),
        _ => {},
    }
}

fn attributes(random: &mut Xorshift, xml: &mut String) {
    let names = ["w:a", "i:b", "i:k", "j:c", "d", "p0:e", "p1:f", "xml:space"];
    let count = random.below(4);
    let mut used = Vec::new();
    for _ in 0..count {
        let name = names[random.below(names.len() as u64) as usize];
        if used.contains(&name) {
            continue;
        }
        used.push(name);
        let value = if name == "xml:space" {
            "preserve"
        } else {
            "v&amp;1"
        };
        let _ = write!(xml, r#" {name}="{value}""#);
    }
}

fn children(random: &mut Xorshift, xml: &mut String, depth: usize) {
    if depth >= 6 {
        return;
    }
    let count = random.below(4);
    for _ in 0..count {
        match random.below(9) {
            0 | 1 => {
                xml.push_str("<w:x");
                declarations(random, xml, depth);
                attributes(random, xml);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</w:x>");
            },
            2 => {
                xml.push_str("<mc:AlternateContent");
                declarations(random, xml, depth);
                xml.push('>');
                let requires = ["w", "i", "j", "p0"][random.below(4) as usize];
                let _ = write!(xml, r#"<mc:Choice Requires="{requires}""#);
                declarations(random, xml, depth);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</mc:Choice>");
                if random.below(2) == 0 {
                    xml.push_str("<mc:Fallback");
                    declarations(random, xml, depth);
                    xml.push('>');
                    children(random, xml, depth + 1);
                    xml.push_str("</mc:Fallback>");
                }
                xml.push_str("</mc:AlternateContent>");
            },
            3 => {
                xml.push_str("<i:u");
                if random.below(3) == 0 {
                    declarations(random, xml, depth);
                }
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</i:u>");
            },
            4 => {
                xml.push_str("<i:keep");
                attributes(random, xml);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</i:keep>");
            },
            5 => {
                xml.push_str("<i:other");
                declarations(random, xml, depth);
                xml.push('>');
                children(random, xml, depth + 1);
                xml.push_str("</i:other>");
            },
            6 => {
                xml.push_str("<j:y");
                attributes(random, xml);
                xml.push_str("/>");
            },
            7 => xml.push_str("text &lt; t"),
            _ => {
                xml.push_str("<w:e");
                declarations(random, xml, depth);
                attributes(random, xml);
                xml.push_str("/>");
            },
        }
    }
}

// --------------------------------------------------------------------------
// Record 0771: aliased namespaces
// --------------------------------------------------------------------------

/// A deterministic pseudo-random document whose prefixes alias one another
/// and the fixed namespaces: `b` may be bound to `a`'s URI, `x` to the `xml`
/// namespace, `n` to the `xmlns` namespace and `m` to the markup
/// compatibility namespace, at the root or again below it, beside
/// default-namespace resets. Its attributes, elements and directive tokens
/// name those namespaces through either prefix, so two names or targets that
/// differ only in their prefix are the same name.
fn aliasing_document(seed: u64) -> String {
    let mut random = Xorshift(seed.wrapping_mul(0xd1b5_4a32_d192_ed03) | 1);
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:a="urn:a" xmlns:w="urn:w""#);
    aliased_bindings(&mut random, &mut xml, true);
    aliased_directives(&mut random, &mut xml);
    aliased_attributes(&mut random, &mut xml);
    xml.push('>');
    aliased_children(&mut random, &mut xml, 0);
    xml.push_str("</r>");
    xml
}

/// Bind `b`, `x`, `n` and `m`, each to an alias or to a URI of its own: all
/// four at the root, some of them on a descendant.
fn aliased_bindings(random: &mut Xorshift, xml: &mut String, root: bool) {
    let choices: [(&str, &str, &str); 4] = [
        ("b", "urn:a", "urn:b"),
        ("x", XML_NAMESPACE, "urn:x"),
        ("n", XMLNS_NAMESPACE, "urn:n"),
        ("m", MC, "urn:m"),
    ];
    for (prefix, alias, own) in choices {
        if !root && random.below(3) != 0 {
            continue;
        }
        let namespace = if random.below(2) == 0 { alias } else { own };
        let _ = write!(xml, r#" xmlns:{prefix}="{namespace}""#);
    }
    if !root {
        match random.below(6) {
            0 => xml.push_str(r#" xmlns="""#),
            1 => xml.push_str(r#" xmlns="urn:a""#),
            2 => {
                let _ = write!(xml, r#" xmlns="{XML_NAMESPACE}""#);
            },
            _ => {},
        }
    }
}

/// Between one and `most` distinct items of `pool`.
fn pick<'a>(random: &mut Xorshift, pool: &[&'a str], most: u64) -> Vec<&'a str> {
    let count = 1 + random.below(most);
    let mut chosen: Vec<&str> = Vec::new();
    for _ in 0..count {
        let item = pool[random.below(pool.len() as u64) as usize];
        if !chosen.contains(&item) {
            chosen.push(item);
        }
    }
    chosen
}

/// The prefix the root may bind to the same namespace as `prefix`.
fn partner(prefix: &str) -> &str {
    match prefix {
        "a" => "b",
        "b" => "a",
        "x" => "xml",
        "xml" => "x",
        other => other,
    }
}

/// Compatibility directives whose targets name an ignorable prefix, its
/// alias, or now and then any prefix.
fn aliased_directives(random: &mut Xorshift, xml: &mut String) {
    const PREFIXES: [&str; 6] = ["a", "b", "x", "n", "xml", "w"];
    let directive = if random.below(3) == 0 { "m" } else { "mc" };
    if random.below(4) != 0 {
        let ignorable = pick(random, &PREFIXES, 3);
        let _ = write!(xml, r#" {directive}:Ignorable="{}""#, ignorable.join(" "));
        let target = |random: &mut Xorshift, locals: &[&str]| -> String {
            let chosen = ignorable[random.below(ignorable.len() as u64) as usize];
            let prefix = match random.below(6) {
                0 | 1 => partner(chosen),
                2 => PREFIXES[random.below(PREFIXES.len() as u64) as usize],
                _ => chosen,
            };
            let local = locals[random.below(locals.len() as u64) as usize];
            format!("{prefix}:{local}")
        };
        for (name, locals, probability) in [
            (
                "PreserveAttributes",
                &["*", "k", "lang", "space", "q"][..],
                2,
            ),
            ("PreserveElements", &["keep", "*"][..], 3),
            ("ProcessContent", &["u", "*"][..], 3),
        ] {
            if random.below(probability) != 0 {
                continue;
            }
            let count = 1 + random.below(2);
            let targets: Vec<String> = (0..count).map(|_| target(random, locals)).collect();
            let _ = write!(xml, r#" {directive}:{name}="{}""#, targets.join(" "));
        }
    }
    if random.below(12) == 0 {
        let tokens = pick(random, &["a", "b", "w", "x", "xml"], 2);
        let _ = write!(xml, r#" {directive}:MustUnderstand="{}""#, tokens.join(" "));
    }
}

fn aliased_attributes(random: &mut Xorshift, xml: &mut String) {
    let names = [
        "a:k",
        "b:k",
        "x:lang",
        "xml:lang",
        "xml:space",
        "n:q",
        "d",
        "w:v",
        "b:v",
    ];
    // Two names that differ only in aliased prefixes are the same name, so
    // most lists avoid the pair and a few keep it.
    let alias = |name: &str| match name {
        "a:k" => "b:k",
        "b:k" => "a:k",
        "x:lang" => "xml:lang",
        "xml:lang" => "x:lang",
        _ => "",
    };
    let count = random.below(4);
    let mut used = Vec::new();
    for _ in 0..count {
        let name = names[random.below(names.len() as u64) as usize];
        if used.contains(&name) || (used.contains(&alias(name)) && random.below(4) != 0) {
            continue;
        }
        used.push(name);
        let value = if name.ends_with(":space") {
            "preserve"
        } else {
            "v"
        };
        let _ = write!(xml, r#" {name}="{value}""#);
    }
}

fn aliased_children(random: &mut Xorshift, xml: &mut String, depth: usize) {
    if depth >= 4 {
        return;
    }
    let count = random.below(4);
    for _ in 0..count {
        let name = match random.below(12) {
            0 => "a:u",
            1 => "b:u",
            2 => "x:u",
            3 => "a:keep",
            4 => "b:keep",
            5 => "n:e",
            6 => "w:e",
            7 => "e",
            8 => "x:keep",
            9 => {
                aliased_alternate(random, xml, depth);
                continue;
            },
            10 => {
                xml.push_str("t &amp; u");
                continue;
            },
            _ => "b:other",
        };
        let _ = write!(xml, "<{name}");
        if random.below(3) == 0 {
            aliased_bindings(random, xml, false);
        }
        if random.below(4) == 0 {
            aliased_directives(random, xml);
        }
        aliased_attributes(random, xml);
        if random.below(3) == 0 {
            xml.push_str("/>");
        } else {
            xml.push('>');
            aliased_children(random, xml, depth + 1);
            let _ = write!(xml, "</{name}>");
        }
    }
}

/// An `AlternateContent` whose markup names the compatibility namespace
/// through `mc` or `m`, with `Requires` naming aliased prefixes.
fn aliased_alternate(random: &mut Xorshift, xml: &mut String, depth: usize) {
    let container = if random.below(2) == 0 { "mc" } else { "m" };
    let choice = if random.below(2) == 0 { "mc" } else { "m" };
    let requires = pick(random, &["a", "b", "w", "x", "xml", "n"], 2).join(" ");
    let _ = write!(
        xml,
        r#"<{container}:AlternateContent><{choice}:Choice Requires="{requires}">"#
    );
    aliased_children(random, xml, depth + 1);
    let _ = write!(xml, "</{choice}:Choice>");
    if random.below(2) == 0 {
        let fallback = if random.below(2) == 0 { "mc" } else { "m" };
        let _ = write!(xml, "<{fallback}:Fallback>");
        aliased_children(random, xml, depth + 1);
        let _ = write!(xml, "</{fallback}:Fallback>");
    }
    let _ = write!(xml, "</{container}:AlternateContent>");
}


fn main() -> Result<(), Box<dyn std::error::Error>> {
 let a: Vec<String> = std::env::args().collect();
 let seed: u64 = a[2].parse()?;
 let xml = if a[1] == "aliasing" { aliasing_document(seed) } else { generated_document(seed) };
 std::fs::write(&a[3],xml)?;
 Ok(())
}
