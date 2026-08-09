//! Offline upstream conformance examples plus project-authored roundtrip files.
//!
//! See `tests/data/PROVENANCE.md` for the exact source and limits of the claim.

use litchi_markdown::reader::{Dialect, Error, Patch, PatchEnvelopeLimits, ReadLimits, Snapshot};

const COMMONMARK_0312_EXAMPLES: &[&str] = &[
    "\tfoo\tbaz\t\tbim\n",
    "  \tfoo\tbaz\t\tbim\n",
    "    a\ta\n    ὐ\ta\n",
    "  - foo\n\n\tbar\n",
    "- foo\n\n\t\tbar\n",
    ">\t\tfoo\n",
    "-\t\tfoo\n",
    "    foo\n\tbar\n",
    " - foo\n   - bar\n\t - baz\n",
    "#\tFoo\n",
    "*\t*\t*\t\n",
    "\\!\\\"\\#\\$\\%\\&\\'\\(\\)\\*\\+\\,\\-\\.\\/\\:\\;\\<\\=\\>\\?\\@\\[\\\\\\]\\^\\_\\`\\{\\|\\}\\~\n",
    "\\\t\\A\\a\\ \\3\\φ\\«\n",
    "\\*not emphasized*\n\\<br/> not a tag\n\\[not a link](/foo)\n\\`not code`\n1\\. not a list\n\\* not a list\n\\# not a heading\n\\[foo]: /url \"not a reference\"\n\\&ouml; not a character entity\n",
    "\\\\*emphasis*\n",
    "foo\\\nbar\n",
    "`` \\[\\` ``\n",
    "    \\[\\]\n",
    "~~~\n\\[\\]\n~~~\n",
];

const UPSTREAM_GFM_EXAMPLES: &[&str] = &[
    "~~Hi~~ Hello, ~there~ world!\n",
    "This ~~has a\n\nnew paragraph~~.\n",
    "This will ~~~not~~~ strike.\n",
    "- [ ] foo\n- [x] bar\n",
    "- [x] foo\n  - [ ] bar\n  - [x] baz\n- [ ] bim\n",
    "| foo | bar |\n| --- | --- |\n| baz | bim |\n",
    "| abc | defghi |\n:-: | -----------:\nbar | baz\n",
    "| f\\|oo  |\n| ------ |\n| b `\\|` az |\n| b **\\|** im |\n",
    "| abc | def |\n| --- | --- |\n| bar | baz |\n> bar\n",
    "| abc | def |\n| --- | --- |\n| bar | baz |\nbar\n\nbar\n",
    "| abc | def |\n| --- |\n| bar |\n",
    "| abc | def |\n| --- | --- |\n| bar |\n| bar | baz | boo |\n",
    "| abc | def |\n| --- | --- |\n",
    "Hello World\n| abc | def |\n| --- | --- |\n| bar | baz |\n",
];

#[test]
fn selected_upstream_examples_are_available_offline_and_lossless() -> Result<(), Error> {
    for source in COMMONMARK_0312_EXAMPLES {
        let snapshot = Snapshot::read(source)?;
        assert_eq!(snapshot.source().as_bytes(), source.as_bytes());
        assert_eq!(Snapshot::read(source)?, snapshot);
    }
    for source in UPSTREAM_GFM_EXAMPLES {
        let snapshot = Snapshot::read_with(source, Dialect::GitHubFlavored, ReadLimits::DEFAULT)?;
        assert_eq!(snapshot.source().as_bytes(), source.as_bytes());
        assert_eq!(
            Snapshot::read_with(source, Dialect::GitHubFlavored, ReadLimits::DEFAULT)?,
            snapshot
        );
    }
    Ok(())
}

#[test]
fn real_document_fixtures_roundtrip_through_durable_patches() -> Result<(), Error> {
    for (source, dialect) in [
        (
            include_str!("data/roundtrip-commonmark.md"),
            Dialect::CommonMark,
        ),
        (
            include_str!("data/roundtrip-gfm.md"),
            Dialect::GitHubFlavored,
        ),
    ] {
        let snapshot = Snapshot::read_with(source, dialect, ReadLimits::DEFAULT)?;
        let mut edit = snapshot.edit();
        edit.append_block("Roundtrip sentinel")?;
        let commit = edit.commit()?;
        let json = commit.patch().to_json(PatchEnvelopeLimits::DEFAULT)?;
        let durable = Patch::from_json(&json, PatchEnvelopeLimits::DEFAULT)?;
        let applied = snapshot.apply(&durable)?;
        let restored = applied.snapshot().apply(&durable.inverse())?;
        assert_eq!(restored.snapshot().source().as_bytes(), source.as_bytes());
    }
    Ok(())
}
