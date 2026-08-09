//! Project-authored CommonMark/GFM regression corpus.
//!
//! These cases are original compact fixtures checked into this repository.
//! They are organized by normative syntax family but do not claim provenance
//! from, or verbatim identity with, any external/network corpus.

use litchi_markdown::reader::{BlockKind, Dialect, Error, InlineKind, ReadLimits, Snapshot};

const COMMONMARK_BLOCKS: &[(&str, BlockKind)] = &[
    ("plain", BlockKind::Paragraph),
    ("plain\ncontinuation", BlockKind::Paragraph),
    ("# h1", BlockKind::Heading { level: 1 }),
    ("## h2 ##", BlockKind::Heading { level: 2 }),
    ("### h3", BlockKind::Heading { level: 3 }),
    ("#### h4", BlockKind::Heading { level: 4 }),
    ("##### h5", BlockKind::Heading { level: 5 }),
    ("###### h6", BlockKind::Heading { level: 6 }),
    ("setext one\n=", BlockKind::Heading { level: 1 }),
    ("setext two\n---", BlockKind::Heading { level: 2 }),
    ("> quote", BlockKind::BlockQuote),
    ("> quote\n> next", BlockKind::BlockQuote),
    ("> outer\n>> inner", BlockKind::BlockQuote),
    ("- bullet", BlockKind::List { start: None }),
    ("+ bullet", BlockKind::List { start: None }),
    ("* bullet", BlockKind::List { start: None }),
    ("1. ordered", BlockKind::List { start: Some(1) }),
    ("7) ordered", BlockKind::List { start: Some(7) }),
    ("999. ordered", BlockKind::List { start: Some(999) }),
    ("- outer\n  - inner", BlockKind::List { start: None }),
    ("1. outer\n   1. inner", BlockKind::List { start: Some(1) }),
    ("    code", BlockKind::CodeBlock { fenced: false }),
    ("```\ncode\n```", BlockKind::CodeBlock { fenced: true }),
    (
        "~~~~\n``` nested\n~~~~",
        BlockKind::CodeBlock { fenced: true },
    ),
    (
        "``` rust\nfn f() {}\n```",
        BlockKind::CodeBlock { fenced: true },
    ),
    ("***", BlockKind::ThematicBreak),
    ("___", BlockKind::ThematicBreak),
    ("- - -", BlockKind::ThematicBreak),
    ("<div>block</div>", BlockKind::Html),
    ("<!-- comment -->", BlockKind::Html),
    ("<?instruction?>", BlockKind::Html),
    ("<![CDATA[value]]>", BlockKind::Html),
    ("[a]: /one", BlockKind::LinkDefinition),
    ("[A B]: <a b> 'title'", BlockKind::LinkDefinition),
    ("\\# escaped", BlockKind::Paragraph),
    ("&copy; and &#169;", BlockKind::Paragraph),
    ("`inline`", BlockKind::Paragraph),
    ("<https://example.test>", BlockKind::Paragraph),
    ("<name@example.test>", BlockKind::Paragraph),
    ("Unicode 🍋 漢字", BlockKind::Paragraph),
];

const INLINE_CASES: &[(&str, InlineKind)] = &[
    ("*em*", InlineKind::Emphasis),
    ("_em_", InlineKind::Emphasis),
    ("**strong**", InlineKind::Strong),
    ("__strong__", InlineKind::Strong),
    ("***both***", InlineKind::Emphasis),
    ("`code`", InlineKind::Code),
    ("`` code ` tick ``", InlineKind::Code),
    ("[link](/target)", InlineKind::Link),
    ("[link][id]\n\n[id]: /target", InlineKind::Link),
    ("![alt](/image.png)", InlineKind::Image),
    ("<b>raw</b>", InlineKind::Html),
    ("line  \nbreak", InlineKind::HardBreak),
    ("line\\\nbreak", InlineKind::HardBreak),
    ("line\nbreak", InlineKind::SoftBreak),
];

const GFM_BLOCKS: &[(&str, BlockKind)] = &[
    ("|a|b|\n|-|-|\n|1|2|", BlockKind::Table),
    ("a | b\n---|---\nx | y", BlockKind::Table),
    ("[^a]: note", BlockKind::FootnoteDefinition),
    ("> [!NOTE]\n> body", BlockKind::BlockQuote),
    ("> [!WARNING]\n> body", BlockKind::BlockQuote),
    ("- [ ] open", BlockKind::List { start: None }),
    ("- [x] done", BlockKind::List { start: None }),
    ("- [X] done", BlockKind::List { start: None }),
];

#[test]
fn authored_block_goldens_are_exact_and_deterministic() -> Result<(), Error> {
    for &(source, kind) in COMMONMARK_BLOCKS {
        assert_one_block(source, Dialect::CommonMark, kind)?;
    }
    for &(source, kind) in GFM_BLOCKS {
        assert_one_block(source, Dialect::GitHubFlavored, kind)?;
    }
    Ok(())
}

#[test]
fn authored_inline_goldens_retain_delimiter_ranges() -> Result<(), Error> {
    for &(source, expected) in INLINE_CASES {
        let snapshot = Snapshot::read(source)?;
        let found = snapshot
            .blocks()
            .flat_map(|block| block.inlines())
            .any(|inline| inline.kind() == expected && &source[inline.range()] == inline.source());
        assert!(found, "missing {expected:?} in {source:?}");
    }
    let gfm = Snapshot::read_with(
        "~~gone~~ [^note]\n\n[^note]: retained",
        Dialect::GitHubFlavored,
        ReadLimits::DEFAULT,
    )?;
    assert!(
        gfm.blocks()
            .flat_map(|block| block.inlines())
            .any(|inline| {
                matches!(
                    inline.kind(),
                    InlineKind::Strikethrough | InlineKind::FootnoteReference
                )
            })
    );
    Ok(())
}

#[test]
fn generated_cross_product_corpus_is_stable_and_bounded() -> Result<(), Error> {
    let markers = ["*", "_", "**", "__", "`"];
    let payloads = ["ascii", "two words", "punctuation!", "漢字", "🍋"];
    let prefixes = ["", "before ", "(prefix) ", "escaped \\# "];
    let suffixes = ["", " after", ".", " and tail"];
    let mut count = 0usize;
    for marker in markers {
        for payload in payloads {
            for prefix in prefixes {
                for suffix in suffixes {
                    let source = format!("{prefix}{marker}{payload}{marker}{suffix}");
                    let first = Snapshot::read(&source)?;
                    let second = Snapshot::read(&source)?;
                    assert_eq!(first, second);
                    assert_eq!(first.source(), source);
                    assert_eq!(first.blocks().len(), 1);
                    count = count.saturating_add(1);
                }
            }
        }
    }
    assert_eq!(count, 400);
    Ok(())
}

fn assert_one_block(source: &str, dialect: Dialect, expected: BlockKind) -> Result<(), Error> {
    let first = Snapshot::read_with(source, dialect, ReadLimits::DEFAULT)?;
    let second = Snapshot::read_with(source, dialect, ReadLimits::DEFAULT)?;
    assert_eq!(first, second);
    assert_eq!(first.source(), source);
    assert_eq!(first.blocks().len(), 1, "{source:?}");
    assert_eq!(first.block(0).map(|block| block.kind()), Some(expected));
    Ok(())
}
