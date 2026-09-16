//! Scratch correctness driver for change 0640 (not part of the tree).
//!
//! digest-doc <root>   : change 0596's differential oracle, verbatim
//! fld-dump <path>     : independent decode of every `Plcfld` in the file,
//!                       from the raw Table-stream bytes, with the `FieldList`
//!                       walk annotated so the disagreeing marker is visible
//! stsh-dump <path>    : every style's index, name and aliases, read through a
//!                       tolerant stylesheet parse, with the duplicates marked
//! open-leniency <path>: strict vs. `TolerateStylesheetDefects` open outcome
//! provenance <path>   : FIB version/product identification for the fixture
//! text-dump <path>    : `Document::text()`, for the independent-reading check
//! text-dump-lenient <path> : the same through `Leniency::TolerateStylesheetDefects`
//! profile <path> <warmups> <samples> : eager-open loop for callgrind

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::parts::text::TextExtractor;

type BoxError = Box<dyn std::error::Error>;

fn walk(root: &Path, ext: &str, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(root) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, ext, out);
            } else if path
                .extension()
                .and_then(|e| e.to_str())
                .map(|e| e.eq_ignore_ascii_case(ext))
                .unwrap_or(false)
            {
                out.push(path);
            }
        }
    }
}

fn short(err: impl std::fmt::Debug) -> String {
    let s = format!("{err:?}");
    let s = s.replace('\n', " ");
    if s.len() > 200 { format!("{}...", &s[..200]) } else { s }
}

/// FNV-1a 64-bit over a byte stream, used as a stable cross-binary digest.
const FNV_OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
fn fnv1a64(seed: u64, bytes: &[u8]) -> u64 {
    let mut hash = seed;
    for &byte in bytes {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

/// Differential oracle: one line per admitted `.doc`, covering the decoded text,
/// every resolved paragraph (which carries its runs' character properties), the
/// section table, the FIB bytes the parser retains, and the exact refusal text
/// for anything that does not open. Copied verbatim from change 0596's probe so
/// the two digests are directly comparable.
fn digest_doc(root: &Path) {
    let mut files = Vec::new();
    walk(root, "doc", &mut files);
    files.sort();
    println!(
        "file\topen\ttext_len\ttext_hash\tpara_count\tcount_api\tpara_hash\tsection_count\tsection_hash\tsubdoc\tfib_len\tfib_hash"
    );
    for path in files {
        let rel = path.display().to_string();
        let mut package = match litchi_doc::Package::open(&path) {
            Ok(package) => package,
            Err(error) => {
                println!("{rel}\tPKGERR {}\t\t\t\t\t\t\t\t\t\t", short(error));
                continue;
            },
        };
        let document = match package.document() {
            Ok(document) => document,
            Err(error) => {
                println!("{rel}\tDOCERR {}\t\t\t\t\t\t\t\t\t\t", short(error));
                continue;
            },
        };
        let (text_len, text_hash) = match document.text() {
            Ok(text) => (
                text.len() as i64,
                fnv1a64(FNV_OFFSET, text.as_bytes()).to_string(),
            ),
            Err(error) => (-1, format!("ERR {}", short(error))),
        };
        let (para_count, para_hash) = match document.paragraphs() {
            Ok(paragraphs) => {
                let mut hash = FNV_OFFSET;
                for paragraph in &paragraphs {
                    hash = fnv1a64(hash, format!("{paragraph:?}").as_bytes());
                }
                (paragraphs.len() as i64, hash.to_string())
            },
            Err(error) => (-1, format!("ERR {}", short(error))),
        };
        let count_api = match document.paragraph_count() {
            Ok(count) => count as i64,
            Err(_) => -1,
        };
        let sections = document.sections();
        let section_hash = fnv1a64(FNV_OFFSET, format!("{sections:?}").as_bytes());
        let fib = document.fib();
        let subdoc = fnv1a64(
            FNV_OFFSET,
            format!("{:?}", fib.get_all_subdoc_ranges()).as_bytes(),
        );
        let raw = fib.raw_data();
        println!(
            "{rel}\tOK\t{text_len}\t{text_hash}\t{para_count}\t{count_api}\t{para_hash}\t{}\t{section_hash}\t{subdoc}\t{}\t{}",
            sections.len(),
            raw.len(),
            fnv1a64(FNV_OFFSET, raw)
        );
    }
}

// ---------------------------------------------------------------------------
// Raw-byte access shared by the structural dumps.
// ---------------------------------------------------------------------------

struct Streams {
    fib: FileInformationBlock,
    word_document: Vec<u8>,
    table: Vec<u8>,
}

fn load(path: &Path) -> Result<Streams, BoxError> {
    let file = std::fs::File::open(path)?;
    let mut ole = litchi_cfb::OleFile::open(file)?;
    let word_document = ole.open_stream(&["WordDocument"])?;
    let fib = FileInformationBlock::parse(&word_document)?;
    let name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table = ole.open_stream(&[name])?;
    Ok(Streams {
        fib,
        word_document,
        table,
    })
}

/// FIB pointer index and label of every field-bearing story (MS-DOC 2.5.1).
const STORIES: [(usize, &str); 7] = [
    (16, "Main/fcPlcfFldMom"),
    (17, "Header/fcPlcfFldHdr"),
    (18, "Footnote/fcPlcfFldFtn"),
    (19, "Comment/fcPlcfFldAtn"),
    (48, "Endnote/fcPlcfFldEdn"),
    (57, "Textbox/fcPlcfFldTxbx"),
    (59, "HeaderTextbox/fcPlcffldHdrTxbx"),
];

fn flag_names(byte: u8) -> String {
    let mut names = Vec::new();
    for (mask, name) in [
        (0x01u8, "fDiffer"),
        (0x02, "fZombieEmbed"),
        (0x04, "fResultDirty"),
        (0x08, "fResultEdited"),
        (0x10, "fLocked"),
        (0x20, "fPrivateResult"),
        (0x40, "fNested"),
        (0x80, "fHasSep"),
    ] {
        if byte & mask != 0 {
            names.push(name);
        }
    }
    if names.is_empty() {
        "-".to_string()
    } else {
        names.join("|")
    }
}

/// Decode one `Plcfld` straight from the Table-stream bytes and walk the
/// `FieldList` grammar (MS-DOC 2.8.25), printing where each end marker's
/// `grffldEnd` agrees or disagrees with the structure the markers describe.
fn fld_dump(path: &Path) -> Result<(), BoxError> {
    let streams = load(path)?;
    let extractor = TextExtractor::new(&streams.fib, &streams.word_document, &streams.table).ok();
    println!("# {}", path.display());
    println!(
        "fib.nFib={:#06x} ({})",
        streams.fib.version(),
        streams.fib.version_name()
    );
    for (index, label) in STORIES {
        let Some((offset, length)) = streams.fib.get_table_pointer(index) else {
            continue;
        };
        if length == 0 {
            continue;
        }
        let start = offset as usize;
        let end = start + length as usize;
        let Some(data) = streams.table.get(start..end) else {
            println!("{label}: range {start}..{end} outside the Table stream");
            continue;
        };
        let count = (data.len() - 4) / 6;
        println!(
            "\n## story {label} (pointer {index}) fc={offset} lcb={length} markers={count}"
        );
        // CP array, then the FLD array (MS-DOC 2.8.25: a PLC of CPs plus a
        // parallel array of `Fld` structures, two bytes each).
        let cp_bytes = (count + 1) * 4;
        let mut stack: Vec<(usize, u32, u8)> = Vec::new();
        for i in 0..count {
            let cp = u32::from_le_bytes(data[i * 4..i * 4 + 4].try_into().unwrap());
            let fld = &data[cp_bytes + i * 2..cp_bytes + i * 2 + 2];
            let ch = fld[0] & 0x1F;
            let reserved = fld[0] >> 5;
            let second = fld[1];
            let kind = match ch {
                0x13 => "begin",
                0x14 => "separator",
                0x15 => "end  ",
                _ => "?????",
            };
            let text_char = extractor
                .as_ref()
                .map(|e| e.text_at_range(cp, cp + 1).chars().next().unwrap_or('\u{0}'))
                .map(|c| format!("U+{:04X}", c as u32))
                .unwrap_or_else(|| "-".to_string());
            let mut note = String::new();
            match ch {
                0x13 => {
                    stack.push((i, cp, second));
                    note = format!("flt={second:#04x} depth_after_push={}", stack.len());
                },
                0x14 => {
                    note = format!("ignored={second:#04x} depth={}", stack.len());
                },
                0x15 => {
                    let open = stack.pop();
                    let enclosing = !stack.is_empty();
                    let nested_bit = second & 0x40 != 0;
                    let has_sep_bit = second & 0x80 != 0;
                    // Was there a separator between this end and its begin?
                    let has_sep_struct = if let Some((begin_index, _, _)) = open {
                        (begin_index + 1..i).any(|j| {
                            let f = &data[cp_bytes + j * 2..cp_bytes + j * 2 + 2];
                            // Only separators that belong to this field, i.e.
                            // at the same nesting depth, count. Recompute by
                            // replaying between the begin and this end.
                            let mut depth = 0i32;
                            for k in begin_index + 1..j {
                                let g = &data[cp_bytes + k * 2..cp_bytes + k * 2 + 2];
                                match g[0] & 0x1F {
                                    0x13 => depth += 1,
                                    0x15 => depth -= 1,
                                    _ => {},
                                }
                            }
                            depth == 0 && f[0] & 0x1F == 0x14
                        })
                    } else {
                        false
                    };
                    note = format!(
                        "grffldEnd={second:#04x} [{}] fNested={nested_bit} enclosing_field={enclosing} {}   fHasSep={has_sep_bit} separator_present={has_sep_struct} {}",
                        flag_names(second),
                        if nested_bit == enclosing { "AGREE" } else { "DISAGREE" },
                        if has_sep_bit == has_sep_struct { "AGREE" } else { "DISAGREE" },
                    );
                    if let Some((begin_index, begin_cp, flt)) = open {
                        note.push_str(&format!(
                            "   begin=#{begin_index}@cp{begin_cp} flt={flt:#04x}"
                        ));
                    }
                },
                _ => {},
            }
            println!("#{i:<4} cp={cp:<8} fld=[{:#04x},{:#04x}] ch={ch:#04x} {kind} reserved={reserved} text={text_char}  {note}", fld[0], fld[1]);
        }
        let terminal = u32::from_le_bytes(data[count * 4..count * 4 + 4].try_into().unwrap());
        println!("terminal_cp={terminal} residual_open={}", stack.len());
        if let Some(extractor) = extractor.as_ref() {
            // Print the raw instruction text of each top-level field so the
            // structure can be read against what the document actually says.
            let mut open: Vec<(u32, usize)> = Vec::new();
            for i in 0..count {
                let cp = u32::from_le_bytes(data[i * 4..i * 4 + 4].try_into().unwrap());
                let fld = &data[cp_bytes + i * 2..cp_bytes + i * 2 + 2];
                match fld[0] & 0x1F {
                    0x13 => open.push((cp, i)),
                    0x15 => {
                        if let Some((begin_cp, begin_index)) = open.pop() {
                            let raw = extractor.text_at_range(begin_cp, cp + 1);
                            let shown: String = raw
                                .chars()
                                .map(|c| {
                                    if (c as u32) < 0x20 {
                                        char::from_u32(0x2400 + c as u32).unwrap_or('?')
                                    } else {
                                        c
                                    }
                                })
                                .take(240)
                                .collect();
                            println!(
                                "field #{begin_index}..#{i} cp {begin_cp}..{cp} depth={} text={shown:?}",
                                open.len()
                            );
                        }
                    },
                    _ => {},
                }
            }
        }
    }
    Ok(())
}

/// Every style's index, name and aliases, through a tolerant parse so that a
/// duplicate does not hide the rest of the sheet.
fn stsh_dump(path: &Path) -> Result<(), BoxError> {
    let streams = load(path)?;
    println!("# {}", path.display());
    let strict = litchi_doc::parts::styles::StyleSheet::parse(&streams.fib, &streams.table);
    println!("strict parse: {}", match &strict {
        Ok(_) => "OK".to_string(),
        Err(error) => format!("ERR {}", short(error)),
    });
    let sheet = litchi_doc::parts::styles::StyleSheet::parse_with_leniency(
        &streams.fib,
        &streams.table,
        litchi_doc::Leniency::TolerateStylesheetDefects,
    )?;
    println!(
        "tolerant parse: OK, cstd={} tolerance_report={:?}",
        sheet.header().style_count,
        sheet.tolerance_report()
    );
    let mut seen: HashMap<String, Vec<u16>> = HashMap::new();
    for style in sheet.styles().iter().flatten() {
        for name in std::iter::once(&style.name).chain(style.aliases.iter()) {
            seen.entry(name.clone()).or_default().push(style.index);
        }
    }
    for style in sheet.styles().iter().flatten() {
        let mut marks = Vec::new();
        for name in std::iter::once(&style.name).chain(style.aliases.iter()) {
            if seen[name].len() > 1 {
                marks.push(format!("DUP {name:?} on styles {:?}", seen[name]));
            }
        }
        println!(
            "istd={:<4} sti={:<5} kind={:?} name={:?} aliases={:?} base={:?} next={} {}",
            style.index,
            style.invariant_id,
            style.kind,
            style.name,
            style.aliases,
            style.base_style,
            style.next_style,
            marks.join("; ")
        );
    }
    let duplicates: Vec<_> = seen.iter().filter(|(_, v)| v.len() > 1).collect();
    println!("duplicate name count = {}", duplicates.len());
    for (name, indices) in duplicates {
        println!("DUPLICATE {name:?} -> istds {indices:?}");
    }
    Ok(())
}

/// Does `OpenOptions::with_leniency` alone admit this file?
fn open_leniency(path: &Path) -> Result<(), BoxError> {
    let strict = {
        let mut package = litchi_doc::Package::open(path)?;
        match package.document() {
            Ok(document) => format!("OK text_len={}", document.text().map(|t| t.len() as i64).unwrap_or(-1)),
            Err(error) => format!("ERR {}", short(error)),
        }
    };
    let lenient = {
        let mut package = litchi_doc::Package::open(path)?;
        let options = litchi_doc::OpenOptions::default()
            .with_leniency(litchi_doc::Leniency::TolerateStylesheetDefects);
        match package.document_with_options(options) {
            Ok(document) => format!(
                "OK text_len={} paragraphs={}",
                document.text().map(|t| t.len() as i64).unwrap_or(-1),
                document.paragraphs().map(|p| p.len() as i64).unwrap_or(-1)
            ),
            Err(error) => format!("ERR {}", short(error)),
        }
    };
    println!("{}\tstrict={strict}\tlenient={lenient}", path.display());
    Ok(())
}

fn provenance(path: &Path) -> Result<(), BoxError> {
    let streams = load(path)?;
    let raw = streams.fib.raw_data();
    println!("# {}", path.display());
    println!("bytes={}", std::fs::metadata(path)?.len());
    println!(
        "nFib={:#06x} ({})  lid={:#06x}  which_table_stream={}  encrypted={}",
        streams.fib.version(),
        streams.fib.version_name(),
        streams.fib.language_id(),
        streams.fib.which_table_stream(),
        streams.fib.is_encrypted()
    );
    // FibBase (MS-DOC 2.5.2): wIdent, nFib, unused, lid, pnNext, flags, nFibBack,
    // lKey, envr, flags2, reserved3/4, reserved5/6.
    println!("FibBase[0..32] = {:02x?}", &raw[..32.min(raw.len())]);
    // nProduct is in the FibBase `nFibBack`/`lKey` neighbourhood only for old
    // files; the reliable build stamp is FibRgLw97.lProductCreated/Revised.
    Ok(())
}

fn text_dump(path: &Path, lenient: bool) -> Result<(), BoxError> {
    let mut package = litchi_doc::Package::open(path)?;
    let document = if lenient {
        package.document_with_options(
            litchi_doc::OpenOptions::default()
                .with_leniency(litchi_doc::Leniency::TolerateStylesheetDefects),
        )?
    } else {
        package.document()?
    };
    let text = document.text()?;
    print!("{text}");
    Ok(())
}

/// Open the eager document `warmups + samples` times, for a callgrind
/// isolation pair (profile N and N+M, difference, divide by M).
fn profile(path: &Path, warmups: usize, samples: usize) -> Result<(), BoxError> {
    let mut checksum = 0usize;
    for _ in 0..(warmups + samples) {
        let mut package = litchi_doc::Package::open(path)?;
        let document = package.document()?;
        checksum = checksum.wrapping_add(std::hint::black_box(&document).sections().len());
    }
    eprintln!("checksum={checksum}");
    Ok(())
}

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("digest-doc") => digest_doc(Path::new(&args[1])),
        Some("fld-dump") => fld_dump(Path::new(&args[1]))?,
        Some("stsh-dump") => stsh_dump(Path::new(&args[1]))?,
        Some("open-leniency") => {
            for p in &args[1..] {
                open_leniency(Path::new(p))?;
            }
        },
        Some("provenance") => {
            for p in &args[1..] {
                provenance(Path::new(p))?;
            }
        },
        Some("text-dump") => text_dump(Path::new(&args[1]), false)?,
        Some("text-dump-lenient") => text_dump(Path::new(&args[1]), true)?,
        Some("profile") => profile(Path::new(&args[1]), args[2].parse()?, args[3].parse()?)?,
        _ => {
            return Err("usage: digest-doc <root> | fld-dump <path> | stsh-dump <path> | open-leniency <path...> | provenance <path...> | text-dump <path> | profile <path> <warmups> <samples>".into());
        },
    }
    Ok(())
}
