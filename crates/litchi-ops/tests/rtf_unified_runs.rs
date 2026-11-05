#![cfg(feature = "rtf")]

use litchi_core::{FileFormat, Position};
use litchi_ops::capabilities::Subject;
use litchi_ops::query::{Plan, Query};
use litchi_ops::{Code, File, NewNode, Placement, open_bytes, props};
use std::borrow::Cow;

const INPUT: &[u8] =
    br"{\rtf1\ansi{\header Hidden header}First\line second\par\par A{\b B}{\i C}D\par last}";

fn texts(file: &File, query: &str) -> Vec<String> {
    file.query(query)
        .unwrap()
        .map(|node| node.unwrap().view().text().unwrap().unwrap().into_owned())
        .collect()
}
fn property(file: &File, path: &str, key: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|p| p.key() == key)
        .unwrap()
        .to_string()
}

#[test]
fn borrowed_run_paths_children_and_positions_match_native_paragraph_axes() {
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let file = open_bytes(INPUT).unwrap();
    let mut expected_paths = Vec::new();
    for (p, paragraph) in native.body().paragraphs().enumerate() {
        let parent = file.get(format!("/body/p[{}]", p + 1).as_str()).unwrap();
        let expected: Vec<_> = paragraph.runs().collect();
        assert_eq!(
            parent.view().child_count().unwrap(),
            Some(expected.len() as u32)
        );
        assert_eq!(parent.children().count(), expected.len());
        for (r, run) in expected.iter().enumerate() {
            let path = format!("/body/p[{}]/r[{}]", p + 1, r + 1);
            let node = file.get(path.as_str()).unwrap();
            assert_eq!(node.kind(), "run");
            assert_eq!(node.path().to_string(), path);
            assert_eq!(
                node.as_rtf().unwrap().paragraph_position(),
                Some(Position::new(p))
            );
            assert_eq!(
                node.as_rtf().unwrap().run_position(),
                Some(Position::new(r))
            );
            assert!(
                matches!(node.view().text().unwrap(), Some(Cow::Borrowed(text)) if text==run.text())
            );
            assert_eq!(node.children().count(), 0);
            assert_eq!(node.view().child_count().unwrap(), Some(0));
            expected_paths.push(path);
        }
        assert!(
            file.find(format!("/body/p[{}]/r[{}]", p + 1, expected.len() + 1).as_str())
                .unwrap()
                .is_none()
        );
        let last = file
            .find(format!("/body/paragraph[{}]/run[last()]", p + 1).as_str())
            .unwrap();
        assert_eq!(
            last.map(|node| node.view().text().unwrap().unwrap().into_owned()),
            expected.last().map(|run| run.text().to_owned())
        );
    }
    let actual = file
        .query("body > p > r")
        .unwrap()
        .map(|n| n.unwrap().path().to_string())
        .collect::<Vec<_>>();
    assert_eq!(actual, expected_paths);
    assert_eq!(
        file.get("/body/p[1]/r[0]").unwrap_err().code(),
        Code::InvalidPath
    );
    assert_eq!(file.to_bytes().unwrap(), INPUT);
}

#[test]
fn native_run_boundaries_omit_break_tokens_and_keep_paragraph_text_intact() {
    let file = open_bytes(INPUT).unwrap();
    assert_eq!(
        texts(&file, "p > r"),
        ["First", "second", "A", "B", "C", "D", "last"]
    );
    assert_eq!(texts(&file, "p"), ["First\nsecond", "", "ABCD", "last"]);
    assert!(file.find("/body/p[2]/r[1]").unwrap().is_none());
    assert_eq!(texts(&file, "r[bold=true]"), ["B"]);
    assert_eq!(texts(&file, "run[italic=true]"), ["C"]);
    assert_eq!(texts(&file, "p:has(r[bold=true])"), ["ABCD"]);
    assert_eq!(file.query("r").unwrap().limit(2).count(), 2);
    assert_eq!(
        file.query("r")
            .unwrap()
            .first()
            .unwrap()
            .unwrap()
            .view()
            .text()
            .unwrap()
            .unwrap(),
        "First"
    );
}

#[test]
fn character_properties_use_typed_values_and_native_formatting() {
    let input = br"{\rtf1\ansi\fs28\b\i\uldb\strike\striked\v\caps\scaps\outl Styled}";
    let file = open_bytes(input.as_slice()).unwrap();
    for (key, value) in [
        ("bold", "true"),
        ("italic", "true"),
        ("size", "14pt"),
        ("underline", "double"),
        ("strike", "true"),
        ("dstrike", "true"),
        ("vanish", "true"),
        ("caps", "true"),
        ("smallcaps", "true"),
        ("outline", "true"),
    ] {
        assert_eq!(property(&file, "/body/p[1]/r[1]", key), value, "{key}");
        assert_eq!(
            file.query(format!("r[{key}={value}]").as_str())
                .unwrap()
                .count(),
            1,
            "{key}"
        );
    }
    assert_eq!(texts(&file, "r[text=Styled]"), ["Styled"]);
}

#[test]
fn compiled_run_queries_are_reusable_and_bound_to_current_immutable_owner() {
    let original = open_bytes(br"{\rtf1\ansi Before}").unwrap();
    let plan = Plan::compile(&Query::parse("r[text=After]").unwrap(), FileFormat::Rtf).unwrap();
    assert_eq!(original.query(&plan).unwrap().count(), 0);
    let changed = original
        .set("/body/p[1]", props! {"text"=>"After"})
        .unwrap();
    assert_eq!(changed.query(&plan).unwrap().count(), 1);
    assert_eq!(original.query(&plan).unwrap().count(), 0);
    assert_eq!(texts(&original, "r"), ["Before"]);
}

#[test]
fn run_capabilities_admit_sets_adds_and_removal_with_ordering_gaps() {
    let file = open_bytes(br"{\rtf1\ansi First\par Second}").unwrap();
    let before = file.to_bytes().unwrap();
    let run = file.capabilities().element("r").unwrap();
    assert!(run.gap().is_none());
    let ops = run.operations();
    assert!(ops.get && ops.query);
    assert!(ops.set && ops.add && ops.remove);
    let writable = [
        "text",
        "size",
        "bold",
        "italic",
        "underline",
        "strike",
        "dstrike",
        "vanish",
        "caps",
        "smallcaps",
        "outline",
    ];
    for property in run.properties() {
        assert!(property.get());
        assert_eq!(property.set(), writable.contains(&property.name()));
        assert_eq!(property.add(), writable.contains(&property.name()));
    }
    assert!(
        file.capabilities()
            .gaps()
            .any(|gap| matches!(gap.subject(), Subject::Operation { element: "run", .. }))
    );
    assert!(
        file.set("/body/p[1]/r[1]", props! {"font"=>"Other"})
            .is_err()
    );
    assert!(
        file.set(Query::parse("p, r").unwrap(), props! {"text"=>"bad"})
            .is_err()
    );
    assert!(file.remove(Query::parse("r").unwrap()).is_err()); // unscoped mutation
    assert!(file.move_to("/body/p[1]/r[1]", Placement::End).is_err());
    assert!(file.swap("/body/p[1]/r[1]", "/body/p[2]").is_err());
    let added = file
        .add("/body/p[1]", NewNode::new("r").prop("text", "added"))
        .unwrap();
    assert_eq!(
        added
            .file
            .get(&added.path)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("added")
    );
    assert_eq!(file.to_bytes().unwrap(), before);
}

#[test]
fn compressed_transports_and_retained_corpus_match_native_run_text_and_flags() {
    let mut inputs = vec![INPUT.to_vec()];
    for lzfu in [false, true] {
        inputs.push(
            litchi_rtf::transport::compress(
                INPUT,
                lzfu,
                litchi_rtf::transport::CompressionLimits::default(),
            )
            .unwrap(),
        );
    }
    for name in [
        "typed_hyperlinks_and_references.rtf",
        "testNegativeUnicode.rtf",
        "test950Encoding.rtf",
        "chtoutline.rtf",
    ] {
        inputs.push(
            std::fs::read(
                std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../../test-data/rtf")
                    .join(name),
            )
            .unwrap(),
        );
    }
    for bytes in inputs {
        let native = litchi_rtf::Package::from_bytes(&bytes).unwrap();
        let file = open_bytes(bytes.as_slice()).unwrap();
        let expected = native
            .body()
            .paragraphs()
            .flat_map(|p| p.runs())
            .map(|r| r.text().to_owned())
            .collect::<Vec<_>>();
        assert_eq!(texts(&file, "r"), expected);
        let bold = native
            .body()
            .paragraphs()
            .flat_map(|p| p.runs())
            .filter(|r| r.format().bold())
            .map(|r| r.text().to_owned())
            .collect::<Vec<_>>();
        assert_eq!(texts(&file, "r[bold=true]"), bold);
        assert_eq!(file.to_bytes().unwrap(), native.to_bytes().unwrap());
    }
}

#[test]
fn every_native_underline_style_stays_distinct_and_queryable() {
    let mut tokens = std::collections::BTreeSet::new();
    let mut native_styles = std::collections::BTreeSet::new();
    for control in [
        "ulnone",
        "ul",
        "uldb",
        "uld",
        "uldash",
        "uldashd",
        "uldashdd",
        "ulw",
        "ulth",
        "ulwave",
        "ulhair",
        "ulthd",
        "ulthdash",
        "ulthdashd",
        "ulthdashdd",
        "ulthldash",
        "ulldash",
        "ulhwave",
        "ululdbwave",
    ] {
        let input = format!(r"{{\rtf1\ansi\{control} Styled}}");
        let native = litchi_rtf::Package::from_bytes(input.as_bytes()).unwrap();
        let native_style = native
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .runs()
            .next()
            .unwrap()
            .format()
            .underline();
        native_styles.insert(format!("{native_style:?}"));
        let file = open_bytes(input.as_bytes()).unwrap();
        let token = property(&file, "/body/p[1]/r[1]", "underline");
        assert_eq!(
            file.query(format!("r[underline={token}]").as_str())
                .unwrap()
                .count(),
            1
        );
        tokens.insert(token);
    }
    assert_eq!(native_styles.len(), 19);
    assert_eq!(tokens.len(), native_styles.len());
}

#[test]
fn resolved_fonts_and_colors_are_native_local_values() {
    let input=br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}\f0\cf1 Red\cf0 Auto}";
    let file = open_bytes(input.as_slice()).unwrap();
    assert_eq!(property(&file, "/body/p[1]/r[1]", "font"), "Example Sans");
    assert_eq!(property(&file, "/body/p[1]/r[1]", "color"), "#FF0000");
    assert_eq!(property(&file, "/body/p[1]/r[2]", "color"), "auto");
    assert_eq!(texts(&file, "r[color=#FF0000]"), ["Red"]);
    assert_eq!(texts(&file, "r[color=auto]"), ["Auto"]);
}
