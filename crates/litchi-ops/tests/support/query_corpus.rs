//! Representative selectors over the workbook, document and presentation
//! corpus, with one digest per format and selector of everything the
//! queries return (records 0262 and 0267).
//!
//! The digest of a selector hashes, file by file in path order, each match's
//! kind, path, positional path and text, or the error code that refused or
//! ended the walk, so two builds that return the same result sets on every
//! file produce the same digests. `tests/query_corpus_differential.rs`
//! compares them with the digests the build before record 0262 printed
//! (`tests/golden/query_corpus/digests.json`); the example
//! `query_corpus_digest` prints them, so any two builds can be compared.
//!
//! Only public API is used, so the same file builds against older code. The
//! hash is FNV-1a (two 64-bit lanes), which is stable across toolchains.

use std::collections::BTreeMap;
use std::fmt::Write as _;
use std::path::{Path, PathBuf};

use litchi_core::FileFormat;
use litchi_ops::File;

/// Workbook selectors: the walk alone, every comparison route (text values,
/// numbers, booleans, tokens, colours, lengths), structure, header rows
/// (refused on files without such a table, which is part of the result),
/// positions and nested selectors.
pub(crate) const XLSX: &[&str] = &[
    "cell",
    "cell[value>1000]",
    "cell[value<=0]",
    "cell[value=0]",
    "cell[value=1]",
    "cell[value!=1]",
    "cell[value>=450.5]",
    "cell[value<-1.5]",
    "cell[value=1e3]",
    "cell[value~=a]",
    "cell[value>a]",
    "cell:contains(a)",
    "cell:empty",
    "cell[type=Number]",
    "cell[bold=true]",
    "cell[font.size>=12]",
    "cell[font.size=11pt]",
    "cell[fill=#FFFF00]",
    "cell[formula]",
    "row",
    "row[height>15]",
    "row[height=15pt]",
    "row[hidden=true]",
    "col",
    "col[width>10]",
    "B",
    "A1:C10",
    "cell[1]",
    "cell[last()]",
    "row[2], cell[value=1]",
    "sheet",
    "sheet cell[value>100]",
    "row:has(cell[value>100])",
    "cell:not([type=SharedString])",
    "cell[value>1000 and bold=true]",
    "table",
    "namedrange",
];

/// Legacy and binary workbook selectors (record 0267): the XLSX selectors,
/// so the three worksheet adapters are held to the same routes. XLS and
/// XLSB digests were recorded from trunk `4d1afe92bc` before the streaming
/// cell cursors.
pub(crate) const XLS: &[&str] = XLSX;

/// Binary workbook selectors (record 0267); see [`XLS`].
pub(crate) const XLSB: &[&str] = XLSX;

/// OpenDocument spreadsheet selectors (record 0277): the XLSX selectors,
/// over packaged `.ods` and flat `.fods` files. Recorded from the ODS read
/// adapter before its writers landed.
pub(crate) const ODS: &[&str] = XLSX;

/// Document selectors.
pub(crate) const DOCX: &[&str] = &[
    "paragraph",
    "p[style=Heading1]",
    "r",
    "r[bold=true]",
    "r[bold!=true]",
    "r[italic=true]",
    "r[size>=12pt]",
    "r[size<10]",
    "r[size=12]",
    "r[size=0.5cm]",
    "r[color=#FF0000]",
    "r[color]",
    "p:contains(the)",
    "p:empty",
    "tbl",
    "tc",
    "tbl p",
    "tbl > tr",
    "p[1]",
    "p[last()]",
    "p:has(r[bold=true])",
    "p:not(:empty)",
    "hyperlink",
    "sdt",
    "/body/r[bold=true]",
];

/// Presentation selectors.
pub(crate) const PPTX: &[&str] = &[
    "shape",
    "shape[width>=10cm]",
    "shape[width=11cm]",
    "shape[width>3960000]",
    "shape[x<2cm]",
    "shape[x>=0.5in]",
    "shape[height=1in]",
    "shape[y!=0]",
    "shape:contains(a)",
    "shape:empty",
    "shape[name~=box]",
    "shape[rotation>0]",
    "shape[fill=#FF0000]",
    "shape[fill]",
    "picture",
    "table",
    "tr",
    "tc",
    "title",
    "textbox",
    "placeholder",
    "paragraph",
    "run[bold=true]",
    "run[size>=18pt]",
    "run[size=18]",
    "slide",
    "group shape",
    "shape[1]",
    "shape[last()]",
    "shape[width>=10cm and height<2cm]",
    "slide > shape",
    "notes",
    "comment",
];

/// OpenDocument text selectors (record 0276): the element kinds of the body
/// index, character formatting predicates, structure, positions and nested
/// selectors.
#[cfg(feature = "odt")]
pub(crate) const ODT: &[&str] = &[
    "p",
    "h",
    "r",
    "r[bold=true]",
    "r[italic=true]",
    "r[size>=12pt]",
    "r[color]",
    "p:contains(the)",
    "p:empty",
    "p[style=Standard]",
    "p[align=center]",
    "tbl",
    "tr",
    "tc",
    "tbl p",
    "tbl > tr",
    "p[1]",
    "p[last()]",
    "p:has(r[bold=true])",
    "p:not(:empty)",
    "hyperlink",
    "list",
    "listItem",
    "section",
    "frame",
    "picture",
    "textbox",
    "footnote",
    "endnote",
    "comment",
    "bookmark",
    "style",
    "/body/p",
    "/body/r[bold=true]",
];

/// OpenDocument drawing selectors (record 0282), recorded from the
/// object-tree adapter's reads.
#[cfg(feature = "odg")]
pub(crate) const ODG: &[&str] = &[
    "page",
    "shape",
    "layer",
    "shape[name]",
    "page > shape",
    "shape:contains(a)",
    "shape[width>=5cm]",
    "shape[shapeType=rect]",
    "page[1]",
    "shape[last()]",
    "/page[1]/shape",
];

/// OpenDocument chart selectors (record 0282).
#[cfg(feature = "odc")]
pub(crate) const ODC: &[&str] = &[
    "axis",
    "series",
    "legend",
    "axis[dimension=x]",
    "series[values]",
];

/// OpenDocument database selectors (record 0282).
#[cfg(feature = "odb")]
pub(crate) const ODB: &[&str] = &[
    "table",
    "column",
    "query",
    "table > column",
    "column[typeName]",
    "query:contains(SELECT)",
    "table[columnCount>2]",
];

/// OpenDocument formula selectors (record 0282).
#[cfg(feature = "odf-formula")]
pub(crate) const ODF_FORMULA: &[&str] = &[
    "element",
    "element[tag=mi]",
    "element[text]",
    "element > element",
    "element:contains(x)",
];

/// OpenDocument presentation selectors (record 0278), recorded from the
/// read adapter before its write handlers landed.
#[cfg(feature = "odp")]
pub(crate) const ODP: &[&str] = &[
    "shape",
    "shape[width>=10cm]",
    "shape[x<2cm]",
    "shape[y!=0]",
    "shape:contains(a)",
    "shape:empty",
    "shape[name~=Title]",
    "shape[isTitle=true]",
    "title",
    "textbox",
    "picture",
    "table",
    "tr",
    "tc",
    "tc:contains(a)",
    "paragraph",
    "paragraph[level>=1]",
    "p:contains(the)",
    "run",
    "run[style]",
    "slide",
    "slide[name~=page]",
    "shape[1]",
    "shape[last()]",
    "slide > shape",
    "notes",
    "notes paragraph",
    "comment",
    "ole",
    "chart",
    "group",
    "connector",
    "slide:has(picture)",
    "paragraph:not(:empty)",
    "shape[width>=10cm and height<2cm]",
    "slideMaster",
    "slideLayout",
];

/// Apple Pages selectors (record 0296), recorded from the iWork read
/// adapter.
#[cfg(feature = "pages")]
pub(crate) const PAGES: &[&str] = &[
    "p",
    "section",
    "body > p",
    "p:contains(e)",
    "p[section>=2]",
    "section[paragraphCount>1]",
    "p[last()]",
    "p:not(:empty)",
    // Record 0299: anchored objects, body tables, headers, footers and
    // footnotes.
    "p[objects>0]",
    "tbl",
    "tbl > tr",
    "tc[type=text]",
    "header",
    "footer[template=odd]",
    "footnote",
];

/// Apple Keynote selectors (records 0296 and 0298).
#[cfg(feature = "keynote")]
pub(crate) const KEYNOTE: &[&str] = &[
    "slide",
    "placeholder",
    "notes",
    "placeholder[type=title]",
    "slide[hidden=true]",
    "slide > placeholder",
    "placeholder:contains(a)",
    "slide[last()]",
    // Record 0298: paragraphs, runs, drawables, tables and layouts.
    "paragraph",
    "placeholder > p",
    "notes > p",
    "shape",
    "shape:contains(a)",
    "picture",
    "video",
    "table",
    "tc",
    "run",
    "shape[zorder>0]",
    "slide[layout]",
];

/// Apple Numbers selectors (record 0296).
#[cfg(feature = "numbers")]
pub(crate) const NUMBERS: &[&str] = &[
    "sheet",
    "table",
    "cell",
    "cell[type=number]",
    "cell[type=text]",
    "table[rowCount>10]",
    "cell:contains(a)",
    "table > cell",
    "sheet[name~=Sheet]",
    // Record 0297: merged ranges.
    "mergeCell",
    "table[mergeCount>0]",
    "cell[mergeAnchor=true]",
];

/// FNV-1a over two 64-bit lanes with different offsets.
struct Fnv([u64; 2]);

impl Fnv {
    const PRIME: u64 = 0x0000_0100_0000_01B3;

    fn new() -> Self {
        Self([0xCBF2_9CE4_8422_2325, 0x6C62_272E_07BB_0142])
    }

    fn update(&mut self, bytes: &[u8]) {
        for lane in &mut self.0 {
            for &byte in bytes {
                *lane ^= u64::from(byte);
                *lane = lane.wrapping_mul(Self::PRIME);
            }
        }
    }

    fn hex(&self) -> String {
        format!("{:016x}{:016x}", self.0[0], self.0[1])
    }
}

/// Appends one record per match or error of `matches` to `record`, and
/// returns the number of matches.
fn record_matches(
    record: &mut String,
    matches: litchi_ops::Result<litchi_ops::query::Matches<'_, '_>>,
) -> usize {
    let mut count = 0;
    match matches {
        Err(error) => {
            let _ = writeln!(record, "refused {:?}", error.code());
        },
        Ok(matches) => {
            for node in matches {
                match node {
                    Ok(node) => {
                        count += 1;
                        let text = node.view().text().ok().flatten();
                        let _ = writeln!(
                            record,
                            "{} {} {} {:?}",
                            node.kind(),
                            node.path(),
                            node.positional_path(),
                            text.as_deref()
                        );
                    },
                    Err(error) => {
                        let _ = writeln!(record, "failed {:?}", error.code());
                    },
                }
            }
        },
    }
    count
}

/// Early-stopping queries on cold handles (record 0267): for every
/// selector, a freshly opened file-backed handle runs `limit(3)` and a
/// freshly opened in-memory handle runs `first()`, so the first matches of
/// a sheet that no earlier query has read are compared too. Keys are
/// `format limit3 selector` and `format first selector`.
fn early_stop_digests(
    root: &Path,
    extension: &str,
    format: FileFormat,
    selectors: &[&str],
    out: &mut BTreeMap<String, String>,
) {
    let wanted = |file: &File| file.format() == Some(format) && !matches!(file, File::Locked(_));
    let paths = corpus(root, extension);
    for selector in selectors {
        let (mut limited, mut first) = (Fnv::new(), Fnv::new());
        let (mut limited_count, mut first_count) = (0_usize, 0_usize);
        for path in &paths {
            let name = path
                .strip_prefix(root)
                .unwrap_or(path)
                .to_string_lossy()
                .replace('\\', "/");
            let Ok(bytes) = std::fs::read(path) else {
                continue;
            };
            if let Ok(file) = litchi_ops::open(path)
                && wanted(&file)
            {
                let mut record = format!("{name}\n");
                limited_count += record_matches(
                    &mut record,
                    file.query(*selector).map(|matches| matches.limit(3)),
                );
                limited.update(record.as_bytes());
            }
            if let Ok(file) = litchi_ops::open_bytes(bytes)
                && wanted(&file)
            {
                let mut record = format!("{name}\n");
                match file
                    .query(*selector)
                    .and_then(litchi_ops::query::Matches::first)
                {
                    Err(error) => {
                        let _ = writeln!(record, "failed {:?}", error.code());
                    },
                    Ok(None) => record.push_str("none\n"),
                    Ok(Some(node)) => {
                        first_count += 1;
                        let text = node.view().text().ok().flatten();
                        let _ = writeln!(
                            record,
                            "{} {} {} {:?}",
                            node.kind(),
                            node.path(),
                            node.positional_path(),
                            text.as_deref()
                        );
                    },
                }
                first.update(record.as_bytes());
            }
        }
        out.insert(
            format!("{extension} limit3 {selector}"),
            format!("{limited_count} {}", limited.hex()),
        );
        out.insert(
            format!("{extension} first {selector}"),
            format!("{first_count} {}", first.hex()),
        );
    }
}

/// The corpus files with `extension` under `root`, in path order.
pub(crate) fn corpus(root: &Path, extension: &str) -> Vec<PathBuf> {
    fn walk(directory: &Path, extension: &str, out: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, extension, out);
            } else if path
                .extension()
                .is_some_and(|found| found.eq_ignore_ascii_case(extension))
            {
                out.push(path);
            }
        }
    }
    let mut out = Vec::new();
    walk(root, extension, &mut out);
    out.sort();
    out
}

/// The digest of every selector of each format over the corpus under
/// `root`, keyed `format selector` with the value `matches digest`; plus,
/// per format, the number of files opened and skipped (refused, encrypted,
/// or read as another format).
pub(crate) fn digests(root: &Path) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();
    for (extension, format, selectors) in [
        ("xlsx", FileFormat::Xlsx, XLSX),
        ("docx", FileFormat::Docx, DOCX),
        ("pptx", FileFormat::Pptx, PPTX),
        ("xls", FileFormat::Xls, XLS),
        ("xlsb", FileFormat::Xlsb, XLSB),
        ("ods", FileFormat::Ods, ODS),
        ("fods", FileFormat::Ods, ODS),
        #[cfg(feature = "odt")]
        ("odt", FileFormat::Odt, ODT),
        // Flat text documents read through the ODT adapter (record 0292).
        #[cfg(feature = "odt")]
        ("fodt", FileFormat::Odt, ODT),
        // ODM master documents run the ODT selectors (record 0282).
        #[cfg(feature = "odm")]
        ("odm", FileFormat::Odm, ODT),
        #[cfg(feature = "odg")]
        ("odg", FileFormat::Odg, ODG),
        // Flat drawings read through the object-tree adapter (record 0292).
        #[cfg(feature = "odg")]
        ("fodg", FileFormat::Odg, ODG),
        #[cfg(feature = "odc")]
        ("odc", FileFormat::Odc, ODC),
        #[cfg(feature = "odb")]
        ("odb", FileFormat::Odb, ODB),
        #[cfg(feature = "odf-formula")]
        ("odf", FileFormat::Odf, ODF_FORMULA),
        #[cfg(feature = "odp")]
        ("odp", FileFormat::Odp, ODP),
        #[cfg(feature = "odp")]
        ("fodp", FileFormat::Odp, ODP),
        #[cfg(feature = "pages")]
        ("pages", FileFormat::Pages, PAGES),
        #[cfg(feature = "keynote")]
        ("key", FileFormat::Keynote, KEYNOTE),
        #[cfg(feature = "numbers")]
        ("numbers", FileFormat::Numbers, NUMBERS),
    ] {
        let mut hashers: Vec<Fnv> = selectors.iter().map(|_| Fnv::new()).collect();
        let mut counts = vec![0_usize; selectors.len()];
        let (mut files, mut skipped) = (0_usize, 0_usize);
        for path in corpus(root, extension) {
            let name = path
                .strip_prefix(root)
                .unwrap_or(&path)
                .to_string_lossy()
                .replace('\\', "/");
            let file = match litchi_ops::open(&path) {
                // An encrypted file opens as `Locked` only with the crypto
                // feature; skipping it keeps the digests feature-independent.
                Ok(file) if file.format() == Some(format) && !matches!(file, File::Locked(_)) => {
                    file
                },
                other => {
                    if std::env::var_os("LITCHI_QUERY_CORPUS_SKIPPED").is_some() {
                        eprintln!("skipped {name}: {:?}", other.map(|file| file.format()));
                    }
                    skipped += 1;
                    continue;
                },
            };
            files += 1;
            for (index, selector) in selectors.iter().enumerate() {
                let mut record = format!("{name}\n");
                counts[index] += record_matches(&mut record, file.query(*selector));
                hashers[index].update(record.as_bytes());
            }
        }
        if matches!(extension, "xlsx" | "xls" | "xlsb" | "ods" | "fods") {
            early_stop_digests(root, extension, format, selectors, &mut out);
        }
        for ((selector, hasher), count) in selectors.iter().zip(hashers).zip(counts) {
            out.insert(
                format!("{extension} {selector}"),
                format!("{count} {}", hasher.hex()),
            );
        }
        out.insert(
            format!("{extension} (files)"),
            format!("{files} opened, {skipped} skipped"),
        );
    }
    out
}
