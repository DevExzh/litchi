//! Parser robustness sweeps: truncation and single-byte mutation of
//! feature-rich seeds must never panic, only yield typed results or errors.

use litchi_rtf::RtfDocument;

/// Feature-rich seeds, one per feature family.
const SEEDS: &[(&str, &str)] = &[
    (
        "table",
        concat!(
            r"{\rtf1\ansi{\colortbl;\red255\green0\blue0;}",
            r"\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10",
            r"\clbrdrt\brdrs\brdrw10\clbrdrl\brdrs\brdrw10\clvertalt\cltxlrtb",
            r"\cellx1440\cellx2880\pard\intbl A1\cell B1\cell\row",
            r"\trowd\clbrdrt\brdrs\brdrw10\clbrdrl\brdrs\brdrw10",
            r"\cellx1440\cellx2880\intbl A2\cell B2\cell\row\pard Tail\par}",
        ),
    ),
    (
        "nested-table",
        concat!(
            r"{\rtf1\ansi\trowd\cellx5000\intbl\itap1 Before \intbl\itap2 Inner\nestcell",
            r"\intbl\itap2\nestcell",
            r"{\*\nesttableprops\trowd\cellx1000\cellx2000\nestrow}",
            r"{\nonesttables ignored fallback}\intbl\itap1 After\cell\row\pard Tail\par}",
        ),
    ),
    (
        "shape",
        concat!(
            r#"{\rtf1 A{\shp{\*\shpinst\shpleft100\shptop50\shpright1100\shpbottom550\shpz3"#,
            r#"{\sp{\sn shapeType}{\sv 202}}{\sp{\sn fillColor}{\sv 255}}"#,
            r#"{\sp{\sn hyperlink}{\sv }{\hl {\hlloc http://example.test/x}}}"#,
            r#"{\shptxt Shape text}}}}"#,
        ),
    ),
    (
        "legacy-drawing",
        concat!(
            r"{\rtf1 A{\*\do\dobxpage\dobypara\dodhgt1\dpline",
            r"\dpptx0\dppty0\dpptx10\dppty10\dpx0\dpy0\dpxsize10\dpysize10}B}",
        ),
    ),
    (
        "field",
        concat!(
            r#"{\rtf1 before{\field{\*\fldinst HYPERLINK "http://example.test" }"#,
            r#"{\fldrslt link text}} after}"#,
        ),
    ),
    (
        "math-zone",
        concat!(
            r"{\rtf1 x{\mmath{\mf{\mfPr{\mtype bar}}{\mnum{\mr 1}}{\mden{\mr 2}}}",
            r"{\msSup{\msup{\mr 3}}{\me{\mr x}}}}{\mmathPara{\mmathParaPr{\mjc center}}{\mr y}} z}",
        ),
    ),
    (
        "custom-xml",
        concat!(
            r#"{\rtf1\ansi{\*\xmlnstbl {\xmlns1 urn:example:test}}"#,
            r#"{\xmlopen \xmlns1 employee}{\*\xmlattrname id}{\*\xmlattrvalue 7}"#,
            r#"Body{\xmlopen inner}x{\xmlclose inner}{\xmlclose employee}}"#,
        ),
    ),
    (
        "lists",
        concat!(
            r#"{\rtf1\ansi{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0\leveljc0"#,
            r#"\levelfollow0{\leveltext \'02\'00.;}{\levelnumbers \'01;}\fi-360\li720}\listid1}}"#,
            r#"{\*\listoverridetable{\listoverride\listid1\listoverridecount0\ls1}}"#,
            r#"\pard\ls1\ilvl0 Item one\par Item two\par}"#,
        ),
    ),
    (
        "picture",
        concat!(
            r"{\rtf1 A{\pict\pngblip\picw2\pich2\picwgoal120\pichgoal120\piccropl10",
            r" 89504e470d0a1a0a0000000d4948445200000001000000010806000000}",
            r" B}",
        ),
    ),
    (
        "footnote",
        r"{\rtf1 body{\footnote\chftn note \b text\b0 with {\field{\*\fldinst PAGE }{\fldrslt 1}}} more}",
    ),
    (
        "stylesheet",
        concat!(
            r#"{\rtf1\ansi{\stylesheet{\s1\ql\b Heading;}{\*\cs2\i Emph;}"#,
            r#"{\*\ts16\tsrowd\b \tscfirstrow\tsclastcol Table List;}}"#,
            r#"\pard\s1\brdrt\brdrs\brdrw20 Styled\par\cs2 emph\par}"#,
        ),
    ),
    (
        "mixed-full",
        concat!(
            r"{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fprq2 Arial;}}",
            r"{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}",
            r#"{\stylesheet{\s1\qc Center;}}"#,
            r#"{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc23\leveljc0"#,
            r#"{\leveltext \'01\u-3913 ?;}{\levelnumbers;}\fi-360\li360}\listid1}}"#,
            r#"{\*\xmlnstbl {\xmlns1 urn:x}}"#,
            r"\paperw12240\paperh15840\margl1440\margr1440\sectd\sbkodd\pgnstarts1",
            r"\pard\s1\ls1\ilvl0\f0\fs24\b Title\par",
            r"{\field{\*\fldinst PAGE }{\fldrslt 1}}",
            r"{\mmath{\mrad{\mdeg{\mr 3}}{\me{\mr x}}}}",
            r"{\*\protstart 0a0b}prot{\*\protend 0a0b}",
            r"\ebcstart edit\ebcend",
            r"{\pict\wmetafile8 0100090000030c0000000000}\par}",
        ),
    ),
];

fn parse_all(prefixes: &[u8]) {
    for end in 0..=prefixes.len() {
        let _ = RtfDocument::parse_bytes(&prefixes[..end]);
    }
}

fn mutate_all(seed: &[u8]) {
    const REPLACEMENTS: &[u8] = b"{}\\\x00\xff'\x7f";
    for index in 0..seed.len() {
        for &replacement in REPLACEMENTS {
            let mut mutated = seed.to_vec();
            mutated[index] = replacement;
            let _ = RtfDocument::parse_bytes(&mutated);
        }
        // Also flip the byte through all values that change control-word shape.
        for replacement in *b"a9- *" {
            let mut mutated = seed.to_vec();
            mutated[index] = replacement;
            let _ = RtfDocument::parse_bytes(&mutated);
        }
    }
}

#[test]
fn truncation_sweeps_never_panic() {
    for (name, seed) in SEEDS {
        parse_all(seed.as_bytes());
        // The seed itself must parse: sweeps only guard against panics.
        RtfDocument::parse(seed)
            .unwrap_or_else(|error| panic!("seed {name} does not parse: {error}"));
    }
}

#[test]
fn single_byte_mutation_sweeps_never_panic() {
    for (_name, seed) in SEEDS {
        mutate_all(seed.as_bytes());
    }
}

/// Regression: a multi-byte scalar right after `\'` split the lexer's raw
/// two-byte hex slice and panicked instead of producing a typed error.
#[test]
fn hex_escape_at_char_boundary_is_a_typed_error() {
    for rtf in ["{\rtf1\\'ÿ9}", "{\rtf1\\'é}", "{\rtf1\\'ÿ}", "{\rtf1\\'0ÿ}"] {
        let error = match RtfDocument::parse(rtf) {
            Ok(_) => panic!("accepted boundary-split hex escape {rtf:?}"),
            Err(error) => error,
        };
        assert!(
            matches!(error, litchi_rtf::RtfError::InvalidUnicode(_)),
            "expected typed InvalidUnicode for {rtf:?}, got {error}"
        );
    }
}

#[test]
fn targeted_malformed_documents_fail_typed() {
    let cases: &[&str] = &[
        // Unterminated root group.
        r"{\rtf1",
        // Unterminated font table.
        r"{\rtf1\ansi{\fonttbl{\f0\fnil Arial;}}",
        // Unterminated body group.
        r"{\rtf1{unclosed",
        // Invalid hex escape digits.
        r"{\rtf1\'zz}",
        // Incomplete hex escape.
        "{\rtf1\\'0",
        // Oversized Unicode escape.
        r"{\rtf1\u999999999999 x}",
        // Missing \u parameter.
        r"{\rtf1\u x}",
        // Bad picture hex run.
        r"{\rtf1{\pict\pngblip 0zz0}}",
        // Odd-length picture hex run.
        r"{\rtf1{\pict\pngblip 012}}",
        // Truncated \bin payload.
        r"{\rtf1\bin5 ab}",
        // Negative binary length.
        r"{\rtf1\bin-4 abcd}",
        // Oversized numeric parameter.
        r"{\rtf1\deftab999999999999999 Body}",
        // Negative document property parameter.
        r"{\rtf1\deftab-5 Body}",
        // Misplaced header destination after body text.
        r"{\rtf1 Body{\fonttbl{\f0\fnil X;}}}",
        // Misplaced body-only destination.
        r"{\rtf1{\*\xmlnstbl {\xmlns1 urn:x}}{\footnote{\xmlopen t}n{\xmlclose t}}}",
        // Invalid control symbol after backslash.
        "{\rtf1\\\x01}",
        // Binary length beyond the safety cap.
        r"{\rtf1{\object\objemb{\*\objdata 00}{\*\objdata 01}}}",
        // Invalid UCS-2 escape (lone surrogate half).
        r"{\rtf1\u55296 ?}",
        // Empty document.
        r"",
        // Not an RTF document at all.
        r"plain text, no groups",
    ];
    for rtf in cases {
        assert!(
            RtfDocument::parse(rtf).is_err(),
            "accepted malformed input {rtf:?}"
        );
    }
}

#[test]
fn deep_nesting_beyond_limits_fails_typed() {
    // Group nesting past the parser depth limit.
    let mut rtf = "{\\rtf1".to_string();
    for _ in 0..300 {
        rtf.push('{');
    }
    for _ in 0..300 {
        rtf.push('}');
    }
    rtf.push('}');
    assert!(RtfDocument::parse(&rtf).is_err());

    // Math nesting past the math depth limit.
    let mut math = "{\\rtf1{\\mmath".to_string();
    for _ in 0..65 {
        math.push_str(r"{\mbox{\me");
    }
    math.push_str(r"{\mr 1}");
    for _ in 0..65 {
        math.push_str("}}");
    }
    math.push_str("}}");
    assert!(RtfDocument::parse(&math).is_err());

    // Custom XML nesting past its depth limit.
    let mut xml = "{\\rtf1".to_string();
    for _ in 0..65 {
        xml.push_str(r"{\xmlopen t}");
    }
    xml.push('B');
    for _ in 0..65 {
        xml.push_str(r"{\xmlclose t}");
    }
    xml.push('}');
    assert!(RtfDocument::parse(&xml).is_err());
}
