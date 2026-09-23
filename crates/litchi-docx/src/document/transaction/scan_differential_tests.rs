//! Change 0754: the tracker-driven layout scanner against the `NsReader`
//! scanner it replaced.
//!
//! `scan_document_with_context` is the admission check of every main-document
//! snapshot, so its verdict is a refusal surface: the layout it returns and
//! every error it raises, with its text, must be what the retained
//! `scan_document_with_context_nsreader_oracle` returns for the same bytes.
//! These tests compare the two over hand-written edge cases, every DOCX
//! fixture's main part, deterministic mutations of those parts, the resource
//! limits, and the managed scan's work charge and budget refusals.

use super::*;
use litchi_core::Budget;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

/// One scan's complete observable outcome.
fn outcome(result: TransactionResult<Layout>) -> String {
    match result {
        Ok(layout) => format!(
            "ok paragraphs={:?} tables={:?} controls={:?} end={} conformance={}",
            layout.paragraphs,
            layout.tables,
            layout.block_controls,
            layout.content_end,
            layout.conformance.namespace()
        ),
        Err(error) => format!("err {error} | {error:?}"),
    }
}

fn preview(xml: &[u8]) -> String {
    String::from_utf8_lossy(&xml[..xml.len().min(160)]).into_owned()
}

fn assert_scan_parity(xml: &[u8]) -> String {
    let scanned = outcome(scan_document_with_context(xml, None));
    let oracle = outcome(scan_document_with_context_nsreader_oracle(xml, None));
    assert_eq!(
        scanned,
        oracle,
        "layout scan diverged from the NsReader oracle on {}",
        preview(xml)
    );
    scanned
}

fn document(body: &str) -> Vec<u8> {
    format!(r#"<w:document xmlns:w="{WORD}"><w:body>{body}</w:body></w:document>"#).into_bytes()
}

/// Hand-written documents for every verdict the scanner reaches.
fn edge_cases() -> Vec<Vec<u8>> {
    let mut cases = vec![
        document("<w:p><w:r><w:t>a</w:t></w:r></w:p><w:tbl/><w:sdt/><w:sectPr/>"),
        document("<w:p/><w:tbl><w:tr/></w:tbl><w:sdt><w:sdtContent/></w:sdt><w:sectPr></w:sectPr>"),
        format!(r#"<w:document xmlns:w="{WORD_STRICT}"><w:body><w:p/></w:body></w:document>"#)
            .into_bytes(),
        // Default namespace documents, and a default namespace unset inside.
        format!(r#"<document xmlns="{WORD}"><body><p/><tbl/><sectPr/></body></document>"#)
            .into_bytes(),
        format!(r#"<document xmlns="{WORD}"><body><p xmlns=""/><p/></body></document>"#)
            .into_bytes(),
        // Shadowed prefixes on direct-body children, on the body and root.
        document(r#"<w:p xmlns:w="urn:foreign"/><w:p/><w:tbl xmlns:w="urn:foreign"></w:tbl>"#),
        document(r#"<w:sectPr xmlns:w="urn:foreign"/><w:p/>"#),
        format!(
            r#"<w:document xmlns:w="{WORD}"><w:body xmlns:w="urn:foreign"><w:p/></w:body></w:document>"#
        )
        .into_bytes(),
        format!(
            r#"<w:document xmlns:w="urn:foreign"><w:body xmlns:w="{WORD}"><w:p/></w:body></w:document>"#
        )
        .into_bytes(),
        // `xmlns:w=""` on a child: its names resolve `Unknown`.
        document(r#"<w:p xmlns:w=""/><w:p/>"#),
        // Another prefix bound to the Word namespace, and five prefixes so
        // the tracker's resolution cache evicts.
        format!(
            r#"<a:document xmlns:a="{WORD}" xmlns:b="{WORD}" xmlns:c="{WORD}" xmlns:d="{WORD}" xmlns:e="{WORD}"><b:body><a:p/><b:p/><c:tbl/><d:sdt/><e:p/><a:p/><c:sectPr/></b:body></a:document>"#
        )
        .into_bytes(),
        // A body nested too deep, a second body, a body after the root closed.
        document("<w:p><w:body/></w:p>"),
        document("<w:p><w:body></w:body></w:p>"),
        format!(
            r#"<w:document xmlns:w="{WORD}"><w:body/><w:body></w:body></w:document>"#
        )
        .into_bytes(),
        format!(
            r#"<w:document xmlns:w="{WORD}"><w:body></w:body><w:body></w:body></w:document>"#
        )
        .into_bytes(),
        format!(r#"<w:document xmlns:w="{WORD}"><w:x><w:body></w:body></w:x></w:document>"#)
            .into_bytes(),
        // A body without a document root, or under a foreign root.
        format!(r#"<w:body xmlns:w="{WORD}"></w:body>"#).into_bytes(),
        format!(
            r#"<x:document xmlns:x="urn:foreign" xmlns:w="{WORD}"><w:body/></x:document>"#
        )
        .into_bytes(),
        // Section properties not last, empty or started.
        document("<w:sectPr/><w:p/>"),
        document("<w:sectPr></w:sectPr><w:p/>"),
        document("<w:sectPr/><w:sectPr/>"),
        document("<w:sectPr/><x:foreign xmlns:x=\"urn:x\"/>"),
        // Missing pieces and forbidden markup.
        format!(r#"<w:document xmlns:w="{WORD}"></w:document>"#).into_bytes(),
        format!(r#"<w:document xmlns:w="{WORD}"><w:body/></w:document>"#).into_bytes(),
        br#"<w:document xmlns:w="urn:x"><w:body></w:body></w:document>"#.to_vec(),
        format!(r#"<!DOCTYPE x><w:document xmlns:w="{WORD}"><w:body/></w:document>"#)
            .into_bytes(),
        format!(r#"<w:document xmlns:w="{WORD}"><w:body><?pi x?></w:body></w:document>"#)
            .into_bytes(),
        format!(r#"<w:document xmlns:w="{WORD}"><w:body><w:p>"#).into_bytes(),
        format!(r#"<w:document xmlns:w="{WORD}"><w:body><w:p></w:q></w:body></w:document>"#)
            .into_bytes(),
        Vec::new(),
        b"<".to_vec(),
        // Namespace declaration errors, on start and empty tags.
        format!(r#"<w:document xmlns:w="{WORD}" xmlns:xml="urn:x"><w:body/></w:document>"#)
            .into_bytes(),
        document(r#"<w:p xmlns:xmlns="urn:x"/>"#),
        document(r#"<w:p><w:r xmlns:q="http://www.w3.org/XML/1998/namespace"/></w:p>"#),
        document(r#"<w:p xmlns:q="http://www.w3.org/2000/xmlns/"></w:p>"#),
        document(r#"<w:p xmlns:xml="http://www.w3.org/XML/1998/namespace"/>"#),
        // Unusual names: two colons, an empty prefix, a trailing colon.
        document("<w:p:x/><:p/><w:/><w:p/>"),
        // Text, CDATA, comments, references and a declaration.
        format!(
            "<?xml version=\"1.0\"?><w:document xmlns:w=\"{WORD}\"><!--c--><w:body><w:p>&amp;<![CDATA[x]]></w:p></w:body></w:document>"
        )
        .into_bytes(),
    ];
    // A byte-order mark shifts every range.
    let mut marked = UTF8_BYTE_ORDER_MARK.to_vec();
    marked.extend_from_slice(&document("<w:p/><w:tbl></w:tbl><w:sectPr/>"));
    cases.push(marked);
    // 256 declarations on one tag pass; 257 fail inside the tokenizer step.
    let declarations = |count: usize| {
        (0..count)
            .map(|index| format!(r#" xmlns:d{index}="urn:{index}""#))
            .collect::<String>()
    };
    cases.push(document(&format!("<w:p{}/>", declarations(256))));
    cases.push(document(&format!("<w:p{}/>", declarations(257))));
    // The depth limit, exactly at and just past it.
    for extra in [MAX_DOCUMENT_DEPTH - 3, MAX_DOCUMENT_DEPTH - 2] {
        cases.push(document(&format!(
            "<w:p>{}{}</w:p>",
            "<w:x>".repeat(extra),
            "</w:x>".repeat(extra)
        )));
    }
    cases
}

fn fixture_documents() -> Vec<(std::path::PathBuf, Vec<u8>)> {
    let mut paths = Vec::new();
    let mut stack = vec![
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("..")
            .join("..")
            .join("test-data"),
    ];
    while let Some(directory) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&directory) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else if path.extension().is_some_and(|extension| {
                ["docx", "docm", "dotx", "dotm"]
                    .iter()
                    .any(|wanted| extension.eq_ignore_ascii_case(wanted))
            }) {
                paths.push(path);
            }
        }
    }
    paths.sort();
    paths
        .into_iter()
        .filter_map(|path| {
            let bytes = std::fs::read(&path).ok()?;
            let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
            let xml = reader.read("word/document.xml").ok()?;
            Some((path, xml))
        })
        .collect()
}

/// A small deterministic generator, so a failure names a reproducible case.
struct Mutator(u64);

impl Mutator {
    fn next(&mut self) -> u64 {
        // xorshift64*
        self.0 ^= self.0 >> 12;
        self.0 ^= self.0 << 25;
        self.0 ^= self.0 >> 27;
        self.0.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    fn below(&mut self, bound: usize) -> usize {
        usize::try_from(self.next() % u64::try_from(bound.max(1)).unwrap()).unwrap()
    }

    /// One mutation of `xml`: a byte replaced, a range deleted or
    /// duplicated, a markup snippet inserted, or a truncation.
    fn mutate(&mut self, xml: &[u8]) -> Vec<u8> {
        const BYTES: &[u8] = b"<>/=\"':x \0\xff&;!?-[]w";
        const SNIPPETS: &[&str] = &[
            "<w:p/>",
            "<w:p>",
            "</w:p>",
            "<w:body>",
            "</w:body>",
            "<w:sectPr/>",
            "<w:tbl/>",
            "<w:sdt>",
            " xmlns:w=\"urn:foreign\"",
            " xmlns=\"\"",
            " xmlns:w=\"\"",
            " xmlns:xml=\"urn:x\"",
            " xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
            " xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\"",
            "<?pi?>",
            "<!DOCTYPE d>",
            "<!--c-->",
            "<![CDATA[x]]>",
            "&amp;",
            "&bogus;",
            "\u{feff}",
        ];
        let mut out = xml.to_vec();
        if out.is_empty() {
            return out;
        }
        match self.below(6) {
            0 | 1 => {
                let at = self.below(out.len());
                out[at] = BYTES[self.below(BYTES.len())];
            },
            2 => {
                let at = self.below(out.len());
                let end = (at + 1 + self.below(24)).min(out.len());
                out.drain(at..end);
            },
            3 => {
                let at = self.below(out.len());
                let end = (at + 1 + self.below(24)).min(out.len());
                let copy = out[at..end].to_vec();
                out.splice(at..at, copy);
            },
            4 => {
                let at = self.below(out.len() + 1);
                let snippet = SNIPPETS[self.below(SNIPPETS.len())].as_bytes();
                out.splice(at..at, snippet.iter().copied());
            },
            _ => out.truncate(self.below(out.len())),
        }
        out
    }
}

#[test]
fn tracker_scan_matches_the_nsreader_oracle_on_edge_cases() {
    let mut accepted = 0usize;
    for case in edge_cases() {
        if assert_scan_parity(&case).starts_with("ok") {
            accepted += 1;
        }
    }
    // The edge cases reach both verdicts.
    assert!(accepted >= 8, "only {accepted} edge cases were accepted");
}

#[test]
fn tracker_scan_matches_the_nsreader_oracle_on_every_docx_fixture() {
    let fixtures = fixture_documents();
    assert!(fixtures.len() >= 40, "found {} fixtures", fixtures.len());
    for (path, xml) in &fixtures {
        let scanned = outcome(scan_document_with_context(xml, None));
        let oracle = outcome(scan_document_with_context_nsreader_oracle(xml, None));
        assert_eq!(scanned, oracle, "diverged on {}", path.display());
    }
}

#[test]
fn tracker_scan_matches_the_nsreader_oracle_on_mutated_fixtures() {
    let fixtures = fixture_documents();
    let mut mutator = Mutator(0x0754_5CA7_D1FF_0001);
    let mut outcomes = [0usize; 2];
    for (path, xml) in fixtures.iter().filter(|(_, xml)| xml.len() <= 256 * 1024) {
        for round in 0..24 {
            let mut mutated = mutator.mutate(xml);
            // Stack a second mutation on every other round.
            if round % 2 == 1 {
                mutated = mutator.mutate(&mutated);
            }
            let scanned = outcome(scan_document_with_context(&mutated, None));
            let oracle = outcome(scan_document_with_context_nsreader_oracle(&mutated, None));
            assert_eq!(
                scanned,
                oracle,
                "diverged on mutation {round} of {}",
                path.display()
            );
            outcomes[usize::from(scanned.starts_with("ok"))] += 1;
        }
    }
    // The mutations reach both verdicts.
    assert!(outcomes[0] > 100 && outcomes[1] > 100, "{outcomes:?}");
}

#[test]
fn tracker_scan_matches_the_nsreader_oracle_on_mutated_edge_cases() {
    let mut mutator = Mutator(0x0754_5CA7_D1FF_0002);
    for case in edge_cases() {
        for _ in 0..40 {
            assert_scan_parity(&mutator.mutate(&case));
        }
    }
}

#[test]
fn tracker_scan_matches_the_nsreader_oracle_at_the_element_limit() {
    for count in [MAX_DOCUMENT_NODES - 2, MAX_DOCUMENT_NODES - 1] {
        // The document root and the body count too.
        let xml = document(&"<w:p/>".repeat(count));
        let verdict = assert_scan_parity(&xml);
        assert_eq!(
            verdict.starts_with("ok"),
            count + 2 <= MAX_DOCUMENT_NODES,
            "{}",
            &verdict[..verdict.len().min(120)]
        );
    }
}

fn managed_context(work: u64) -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "scan-differential",
        litchi_core::Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let (_source, cancellation) = litchi_core::CancellationSource::pair();
    let limits = litchi_core::ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::MAX,
        0,
    )
    .unwrap();
    (
        budget.clone(),
        ExecutionContext::new(budget, cancellation, limits),
    )
}

#[test]
fn managed_tracker_scan_charges_and_refuses_exactly_as_the_oracle() {
    let mut documents = edge_cases();
    documents.extend(
        fixture_documents()
            .into_iter()
            .map(|(_, xml)| xml)
            .filter(|xml| xml.len() <= 128 * 1024)
            .take(12),
    );
    for xml in &documents {
        let (budget, context) = managed_context(u64::MAX);
        let scanned = outcome(scan_document_with_context(xml, Some(&context)));
        let charged = budget.used(Resource::Work);
        let (oracle_budget, oracle_context) = managed_context(u64::MAX);
        let oracle = outcome(scan_document_with_context_nsreader_oracle(
            xml,
            Some(&oracle_context),
        ));
        assert_eq!(scanned, oracle, "managed scan diverged on {}", preview(xml));
        assert_eq!(
            charged,
            oracle_budget.used(Resource::Work),
            "managed work charge diverged on {}",
            preview(xml)
        );
        // Budgets that run out part of the way through refuse at the same
        // event with the same error.
        for divisor in [2, 3, 7] {
            let work = charged / divisor;
            let (_, context) = managed_context(work);
            let (_, oracle_context) = managed_context(work);
            assert_eq!(
                outcome(scan_document_with_context(xml, Some(&context))),
                outcome(scan_document_with_context_nsreader_oracle(
                    xml,
                    Some(&oracle_context)
                )),
                "managed budget refusal diverged on {}",
                preview(xml)
            );
        }
    }
}

#[test]
fn an_exact_no_op_commit_still_shares_its_source_allocation() {
    let xml = document("<w:p><w:r><w:t>a</w:t></w:r></w:p><w:sectPr/>");
    let base = Snapshot::from_shared_xml(Arc::new(xml)).unwrap();
    let commit = base.edit().commit().unwrap();
    assert!(!commit.patch().changed());
    assert!(std::ptr::eq(
        commit.snapshot().xml_bytes(),
        base.xml_bytes()
    ));
}
