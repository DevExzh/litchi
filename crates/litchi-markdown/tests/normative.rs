//! Offline release gates over pinned normative Markdown corpora.
//!
//! Exact versions, hashes, generation commands, attribution, and licenses live
//! in `tests/data/PROVENANCE.md` and alongside the vendored corpora.

use std::ops::Range;

use litchi_markdown::reader::{Dialect, Error, ReadLimits, Snapshot};
use pulldown_cmark::{Options, Parser, html};
use serde::Deserialize;

const COMMONMARK_JSON: &str = include_str!("corpus/commonmark-0.31.2/spec.json");
const GFM_JSON: &str = include_str!("corpus/gfm-0.29.0.gfm.13/spec.json");
const COMMONMARK_EXAMPLE_COUNT: usize = 652;
const GFM_EXAMPLE_COUNT: usize = 670;
const GFM_EXTENSION_EXAMPLE_COUNT: usize = 22;
const GFM_RENDERED_EXTENSION_EXAMPLE_COUNT: usize = 10;

#[derive(Debug, Deserialize)]
struct Example {
    markdown: String,
    html: String,
    example: usize,
    section: String,
    #[serde(default)]
    extensions: Vec<String>,
}

#[test]
fn commonmark_0312_complete_normative_release_gate() -> Result<(), Box<dyn std::error::Error>> {
    let examples: Vec<Example> = serde_json::from_str(COMMONMARK_JSON)?;
    assert_eq!(examples.len(), COMMONMARK_EXAMPLE_COUNT);
    for example in &examples {
        assert_rendering(example, Options::empty());
        assert_exact_model_and_reversible_edit(example, Dialect::CommonMark)?;
    }
    Ok(())
}

#[test]
fn official_gfm_corpus_release_gate() -> Result<(), Box<dyn std::error::Error>> {
    let examples: Vec<Example> = serde_json::from_str(GFM_JSON)?;
    assert_eq!(examples.len(), GFM_EXAMPLE_COUNT);
    assert_eq!(
        examples
            .iter()
            .filter(|example| !example.extensions.is_empty())
            .count(),
        GFM_EXTENSION_EXAMPLE_COUNT
    );
    let mut rendered_extensions = 0usize;
    for example in &examples {
        // The pinned GFM specification is based on CommonMark 0.29. Its 648
        // unmarked examples are retained as an independent parse/edit corpus,
        // while current CommonMark rendering is governed by the complete
        // 0.31.2 gate above. Supported GFM-extension examples remain
        // normative rendering assertions.
        if let Some(options) = supported_gfm_options(&example.extensions) {
            assert_rendering(example, options);
            rendered_extensions = rendered_extensions.saturating_add(1);
        }
        assert_exact_model_and_reversible_edit(example, Dialect::GitHubFlavored)?;
    }
    assert_eq!(rendered_extensions, GFM_RENDERED_EXTENSION_EXAMPLE_COUNT);
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
        let json = commit
            .patch()
            .to_json(litchi_markdown::reader::PatchEnvelopeLimits::DEFAULT)?;
        let durable = litchi_markdown::reader::Patch::from_json(
            &json,
            litchi_markdown::reader::PatchEnvelopeLimits::DEFAULT,
        )?;
        let applied = snapshot.apply(&durable)?;
        let restored = applied.snapshot().apply(&durable.inverse())?;
        assert_eq!(restored.snapshot().source().as_bytes(), source.as_bytes());
    }
    Ok(())
}

fn assert_rendering(example: &Example, options: Options) {
    let mut actual = String::new();
    html::push_html(&mut actual, Parser::new_ext(&example.markdown, options));
    assert_eq!(
        normalized_html(&actual),
        normalized_html(&example.html),
        "normative example {} ({})",
        example.example,
        example.section
    );
}

fn assert_exact_model_and_reversible_edit(
    example: &Example,
    dialect: Dialect,
) -> Result<(), Error> {
    let snapshot = Snapshot::read_with(&example.markdown, dialect, ReadLimits::DEFAULT)?;
    assert_eq!(snapshot.source().as_bytes(), example.markdown.as_bytes());
    assert_eq!(
        snapshot,
        Snapshot::read_with(&example.markdown, dialect, ReadLimits::DEFAULT)?
    );
    assert_ranges(&snapshot, example);

    let mut edit = snapshot.edit();
    edit.append_block("release-gate-sentinel")?;
    let committed = edit.commit()?;
    let restored = committed.snapshot().apply(&committed.patch().inverse())?;
    assert_eq!(restored.snapshot().source(), snapshot.source());
    Ok(())
}

fn assert_ranges(snapshot: &Snapshot, example: &Example) {
    let source = snapshot.source();
    let mut prior_end = 0usize;
    for block in snapshot.blocks() {
        let block_range = block.range();
        assert_valid_range(source, &block_range);
        assert!(
            prior_end <= block_range.start,
            "overlapping top-level block in example {} ({}): prior end {prior_end}, current {block_range:?} {:?}",
            example.example,
            example.section,
            block.kind()
        );
        assert_eq!(&source[block_range.clone()], block.source());
        prior_end = block_range.end;
        for nested in block.descendants() {
            let range = nested.range();
            assert_valid_range(source, &range);
            assert!(block_range.start <= range.start && range.end <= block_range.end);
            assert_eq!(&source[range], nested.source());
        }
        for inline in block.inlines() {
            let range = inline.range();
            assert_valid_range(source, &range);
            assert!(block_range.start <= range.start && range.end <= block_range.end);
            assert_eq!(&source[range], inline.source());
        }
    }
    for reference in snapshot.references() {
        let range = reference.range();
        assert_valid_range(source, &range);
        assert_eq!(&source[range], reference.source());
    }
}

fn assert_valid_range(source: &str, range: &Range<usize>) {
    assert!(range.start <= range.end && range.end <= source.len());
    assert!(source.is_char_boundary(range.start));
    assert!(source.is_char_boundary(range.end));
}

fn supported_gfm_options(extensions: &[String]) -> Option<Options> {
    if extensions.is_empty() {
        return None;
    }
    let mut options = Options::empty();
    for extension in extensions {
        match extension.as_str() {
            "table" => options.insert(Options::ENABLE_TABLES),
            "strikethrough" => options.insert(Options::ENABLE_STRIKETHROUGH),
            // pulldown-cmark 0.13.4 does not implement GFM extended
            // autolinks or the disallowed-raw-HTML tag filter. These cases
            // still run through the exact-source/range/edit gate above.
            "autolink" | "tagfilter" => return None,
            unknown => panic!("unrecognized pinned GFM extension {unknown}"),
        }
    }
    Some(options)
}

fn normalized_html(source: &str) -> String {
    source
        // Both spellings represent the same text node. The upstream
        // CommonMark/GFM harnesses normalize character references before
        // comparing renderer output.
        .replace("&quot;", "\"")
        .replace("align=\"left\"", "style=\"text-align: left\"")
        .replace("align=\"center\"", "style=\"text-align: center\"")
        .replace("align=\"right\"", "style=\"text-align: right\"")
        .replace("<br>", "<br />")
        .replace("<br/>", "<br />")
        .replace("<hr>", "<hr />")
        .replace("<hr/>", "<hr />")
        .replace(">\n<", "><")
        .replace("<tbody></tbody>", "")
}
