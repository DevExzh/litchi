//! Semantic-read probe for change 0653.
//!
//! One fixture path in, a deterministic line-oriented digest of documented
//! public reads out. Every fallible read is caught: a refusal is printed as
//! `ERR<TAB><Debug>` in place of the value so that it can be compared between
//! the two legs. The process never panics out: each section is wrapped in
//! `catch_unwind` and a panic is printed as `PANIC` in place of the value.
//!
//! Usage: `readprobe <fixture-path> <display-path> [--dump]`
//!
//! In `--dump` mode only the markup accessors are printed, and their value is
//! the base64 of the returned bytes rather than a length and a hash, so that a
//! differing hash can be canonicalized outside the probe.

use std::fmt::Debug;
use std::io::Write as _;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::OnceLock;

// ---------------------------------------------------------------- SHA-256 --

const K: [u32; 64] = [
    0x428a_2f98, 0x7137_4491, 0xb5c0_fbcf, 0xe9b5_dba5, 0x3956_c25b, 0x59f1_11f1, 0x923f_82a4,
    0xab1c_5ed5, 0xd807_aa98, 0x1283_5b01, 0x2431_85be, 0x550c_7dc3, 0x72be_5d74, 0x80de_b1fe,
    0x9bdc_06a7, 0xc19b_f174, 0xe49b_69c1, 0xefbe_4786, 0x0fc1_9dc6, 0x240c_a1cc, 0x2de9_2c6f,
    0x4a74_84aa, 0x5cb0_a9dc, 0x76f9_88da, 0x983e_5152, 0xa831_c66d, 0xb003_27c8, 0xbf59_7fc7,
    0xc6e0_0bf3, 0xd5a7_9147, 0x06ca_6351, 0x1429_2967, 0x27b7_0a85, 0x2e1b_2138, 0x4d2c_6dfc,
    0x5338_0d13, 0x650a_7354, 0x766a_0abb, 0x81c2_c92e, 0x9272_2c85, 0xa2bf_e8a1, 0xa81a_664b,
    0xc24b_8b70, 0xc76c_51a3, 0xd192_e819, 0xd699_0624, 0xf40e_3585, 0x106a_a070, 0x19a4_c116,
    0x1e37_6c08, 0x2748_774c, 0x34b0_bcb5, 0x391c_0cb3, 0x4ed8_aa4a, 0x5b9c_ca4f, 0x682e_6ff3,
    0x748f_82ee, 0x78a5_636f, 0x84c8_7814, 0x8cc7_0208, 0x90be_fffa, 0xa450_6ceb, 0xbef9_a3f7,
    0xc671_78f2,
];

fn sha256_hex(data: &[u8]) -> String {
    let mut h: [u32; 8] = [
        0x6a09_e667,
        0xbb67_ae85,
        0x3c6e_f372,
        0xa54f_f53a,
        0x510e_527f,
        0x9b05_688c,
        0x1f83_d9ab,
        0x5be0_cd19,
    ];
    let bit_len = (data.len() as u64).wrapping_mul(8);
    let mut message = Vec::with_capacity(data.len() + 72);
    message.extend_from_slice(data);
    message.push(0x80);
    while message.len() % 64 != 56 {
        message.push(0);
    }
    message.extend_from_slice(&bit_len.to_be_bytes());

    let mut w = [0u32; 64];
    for chunk in message.chunks_exact(64) {
        for (index, slot) in w.iter_mut().enumerate().take(16) {
            let base = index * 4;
            *slot = u32::from_be_bytes([
                chunk[base],
                chunk[base + 1],
                chunk[base + 2],
                chunk[base + 3],
            ]);
        }
        for index in 16..64 {
            let s0 = w[index - 15].rotate_right(7)
                ^ w[index - 15].rotate_right(18)
                ^ (w[index - 15] >> 3);
            let s1 = w[index - 2].rotate_right(17)
                ^ w[index - 2].rotate_right(19)
                ^ (w[index - 2] >> 10);
            w[index] = w[index - 16]
                .wrapping_add(s0)
                .wrapping_add(w[index - 7])
                .wrapping_add(s1);
        }
        let (mut a, mut b, mut c, mut d, mut e, mut f, mut g, mut hh) =
            (h[0], h[1], h[2], h[3], h[4], h[5], h[6], h[7]);
        for index in 0..64 {
            let s1 = e.rotate_right(6) ^ e.rotate_right(11) ^ e.rotate_right(25);
            let ch = (e & f) ^ ((!e) & g);
            let temp1 = hh
                .wrapping_add(s1)
                .wrapping_add(ch)
                .wrapping_add(K[index])
                .wrapping_add(w[index]);
            let s0 = a.rotate_right(2) ^ a.rotate_right(13) ^ a.rotate_right(22);
            let maj = (a & b) ^ (a & c) ^ (b & c);
            let temp2 = s0.wrapping_add(maj);
            hh = g;
            g = f;
            f = e;
            e = d.wrapping_add(temp1);
            d = c;
            c = b;
            b = a;
            a = temp1.wrapping_add(temp2);
        }
        h[0] = h[0].wrapping_add(a);
        h[1] = h[1].wrapping_add(b);
        h[2] = h[2].wrapping_add(c);
        h[3] = h[3].wrapping_add(d);
        h[4] = h[4].wrapping_add(e);
        h[5] = h[5].wrapping_add(f);
        h[6] = h[6].wrapping_add(g);
        h[7] = h[7].wrapping_add(hh);
    }
    let mut out = String::with_capacity(64);
    for word in h {
        out.push_str(&format!("{word:08x}"));
    }
    out
}

// ------------------------------------------------------------------ base64 --

fn base64(data: &[u8]) -> String {
    const ALPHABET: &[u8; 64] =
        b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(data.len().div_ceil(3) * 4);
    for chunk in data.chunks(3) {
        let b0 = u32::from(chunk[0]);
        let b1 = chunk.get(1).copied().map_or(0, u32::from);
        let b2 = chunk.get(2).copied().map_or(0, u32::from);
        let triple = (b0 << 16) | (b1 << 8) | b2;
        out.push(ALPHABET[((triple >> 18) & 0x3f) as usize] as char);
        out.push(ALPHABET[((triple >> 12) & 0x3f) as usize] as char);
        out.push(if chunk.len() > 1 {
            ALPHABET[((triple >> 6) & 0x3f) as usize] as char
        } else {
            '='
        });
        out.push(if chunk.len() > 2 {
            ALPHABET[(triple & 0x3f) as usize] as char
        } else {
            '='
        });
    }
    out
}

// ------------------------------------------------------------------ output --

static DISPLAY: OnceLock<String> = OnceLock::new();
static DUMP: OnceLock<bool> = OnceLock::new();

fn display() -> &'static str {
    DISPLAY.get().map_or("?", String::as_str)
}

fn dump_mode() -> bool {
    DUMP.get().copied().unwrap_or(false)
}

/// One printable line. Control characters are escaped so a value can never
/// break the line-oriented digest.
fn escape(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '\\' => out.push_str("\\\\"),
            '\t' => out.push_str("\\t"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            other if (other as u32) < 0x20 || other as u32 == 0x7f => {
                out.push_str(&format!("\\x{:02x}", other as u32));
            },
            other => out.push(other),
        }
    }
    out
}

fn emit(key: &str, value: &str) {
    if dump_mode() {
        return;
    }
    let mut stdout = std::io::stdout().lock();
    let _ = writeln!(stdout, "{}\t{}\t{}", display(), key, value);
}

fn emit_error<E: Debug>(key: &str, error: &E) {
    if dump_mode() {
        return;
    }
    let mut stdout = std::io::stdout().lock();
    let _ = writeln!(
        stdout,
        "{}\t{}\tERR\t{}",
        display(),
        key,
        escape(&format!("{error:?}"))
    );
}

/// A markup accessor's result. In the ordinary mode the length and a SHA-256
/// of the bytes; in dump mode the bytes themselves, base64-encoded, so the
/// comparison script can canonicalize them.
fn emit_markup(key: &str, bytes: &[u8]) {
    let mut stdout = std::io::stdout().lock();
    if dump_mode() {
        let _ = writeln!(stdout, "{}\t{}\tb64={}", display(), key, base64(bytes));
    } else {
        let _ = writeln!(
            stdout,
            "{}\t{}\tlen={};sha={}",
            display(),
            key,
            bytes.len(),
            sha256_hex(bytes)
        );
    }
}

/// Run one section, printing `PANIC` in place of its values if it unwinds.
fn section<F: FnOnce()>(key: &str, body: F) {
    if catch_unwind(AssertUnwindSafe(body)).is_err() {
        let mut stdout = std::io::stdout().lock();
        let _ = writeln!(stdout, "{}\t{}\tPANIC", display(), key);
    }
}

fn opt<T: std::fmt::Display>(value: Option<T>) -> String {
    value.map_or_else(|| "none".to_owned(), |inner| format!("{inner}"))
}

// -------------------------------------------------------------------- DOCX --

const MAX_PARAGRAPHS: usize = 50_000;
const MAX_TABLES: usize = 5_000;
const MAX_CELLS_PER_SHEET: usize = 5_000;

fn docx(path: &str) {
    use litchi_docx::Package;

    let package = match Package::open(path) {
        Ok(package) => {
            emit("docx.open", "OK");
            package
        },
        Err(error) => {
            emit_error("docx.open", &error);
            return;
        },
    };
    let document = match package.document() {
        Ok(document) => {
            emit("docx.document", "OK");
            document
        },
        Err(error) => {
            emit_error("docx.document", &error);
            return;
        },
    };

    section("docx.paragraphs", || match document.paragraph_count() {
        Err(error) => emit_error("docx.paragraph_count", &error),
        Ok(count) => {
            emit("docx.paragraph_count", &count.to_string());
            match document.paragraphs() {
                Err(error) => emit_error("docx.paragraphs", &error),
                Ok(paragraphs) => {
                    emit("docx.paragraphs.len", &paragraphs.len().to_string());
                    if paragraphs.len() > MAX_PARAGRAPHS {
                        emit(
                            "docx.paragraphs.capped",
                            &format!("at={MAX_PARAGRAPHS}"),
                        );
                    }
                    for (index, paragraph) in paragraphs.iter().take(MAX_PARAGRAPHS).enumerate() {
                        match paragraph.text() {
                            Ok(text) => emit(&format!("docx.par.{index}.text"), &escape(&text)),
                            Err(error) => emit_error(&format!("docx.par.{index}.text"), &error),
                        }
                        match paragraph.extensions() {
                            Ok(extensions) => {
                                let ids = extensions.ids();
                                emit(
                                    &format!("docx.par.{index}.ext"),
                                    &format!(
                                        "para_id={};text_id={};no_spell_err={}",
                                        opt(ids.para_id().map(|id| format!("{id:?}"))),
                                        opt(ids.text_id().map(|id| format!("{id:?}"))),
                                        opt(extensions.no_spell_err())
                                    ),
                                );
                            },
                            Err(error) => emit_error(&format!("docx.par.{index}.ext"), &error),
                        }
                        match paragraph.inlines() {
                            Ok(inlines) => {
                                emit(
                                    &format!("docx.par.{index}.inlines"),
                                    &inlines.len().to_string(),
                                );
                                for (position, inline) in inlines.iter().enumerate() {
                                    if let litchi_docx::Inline::Unknown(opaque) = inline {
                                        emit_markup(
                                            &format!("docx.par.{index}.inline.{position}.xml"),
                                            opaque.xml_bytes(),
                                        );
                                    }
                                }
                            },
                            Err(error) => emit_error(&format!("docx.par.{index}.inlines"), &error),
                        }
                        let runs = match paragraph.runs() {
                            Ok(runs) => runs,
                            Err(error) => {
                                emit_error(&format!("docx.par.{index}.runs"), &error);
                                continue;
                            },
                        };
                        emit(&format!("docx.par.{index}.runs"), &runs.len().to_string());
                        for (run_position, run) in runs.iter().enumerate() {
                            match run.contents() {
                                Ok(contents) => {
                                    for (content_position, content) in contents.iter().enumerate() {
                                        if let litchi_docx::RunContent::Unknown(opaque) = content {
                                            emit_markup(
                                                &format!(
                                                    "docx.par.{index}.run.{run_position}.content.{content_position}.xml"
                                                ),
                                                opaque.xml_bytes(),
                                            );
                                        }
                                    }
                                },
                                Err(error) => emit_error(
                                    &format!("docx.par.{index}.run.{run_position}.contents"),
                                    &error,
                                ),
                            }
                        }
                    }
                },
            }
        },
    });

    section("docx.tables", || match document.table_count() {
        Err(error) => emit_error("docx.table_count", &error),
        Ok(count) => {
            emit("docx.table_count", &count.to_string());
            match document.tables() {
                Err(error) => emit_error("docx.tables", &error),
                Ok(tables) => {
                    emit("docx.tables.len", &tables.len().to_string());
                    for (index, table) in tables.iter().take(MAX_TABLES).enumerate() {
                        match table.row_count() {
                            Ok(rows) => emit(&format!("docx.tbl.{index}.rows"), &rows.to_string()),
                            Err(error) => emit_error(&format!("docx.tbl.{index}.rows"), &error),
                        }
                        match table.column_count() {
                            Ok(columns) => {
                                emit(&format!("docx.tbl.{index}.cols"), &columns.to_string());
                            },
                            Err(error) => emit_error(&format!("docx.tbl.{index}.cols"), &error),
                        }
                        match table.rows() {
                            Err(error) => emit_error(&format!("docx.tbl.{index}.rowlist"), &error),
                            Ok(rows) => {
                                for (row_index, row) in rows.iter().enumerate() {
                                    match row.cell_count() {
                                        Ok(cells) => emit(
                                            &format!("docx.tbl.{index}.row.{row_index}.cells"),
                                            &cells.to_string(),
                                        ),
                                        Err(error) => emit_error(
                                            &format!("docx.tbl.{index}.row.{row_index}.cells"),
                                            &error,
                                        ),
                                    }
                                    match row.extension_ids() {
                                        Ok(ids) => emit(
                                            &format!("docx.tbl.{index}.row.{row_index}.extids"),
                                            &format!(
                                                "para_id={};text_id={}",
                                                opt(ids.para_id().map(|id| format!("{id:?}"))),
                                                opt(ids.text_id().map(|id| format!("{id:?}")))
                                            ),
                                        ),
                                        Err(error) => emit_error(
                                            &format!("docx.tbl.{index}.row.{row_index}.extids"),
                                            &error,
                                        ),
                                    }
                                    match row.cells() {
                                        Err(error) => emit_error(
                                            &format!("docx.tbl.{index}.row.{row_index}.celllist"),
                                            &error,
                                        ),
                                        Ok(cells) => {
                                            for (cell_index, cell) in cells.iter().enumerate() {
                                                match cell.text() {
                                                    Ok(text) => emit(
                                                        &format!(
                                                            "docx.tbl.{index}.row.{row_index}.cell.{cell_index}.text"
                                                        ),
                                                        &escape(&text),
                                                    ),
                                                    Err(error) => emit_error(
                                                        &format!(
                                                            "docx.tbl.{index}.row.{row_index}.cell.{cell_index}.text"
                                                        ),
                                                        &error,
                                                    ),
                                                }
                                            }
                                        },
                                    }
                                }
                            },
                        }
                    }
                },
            }
        },
    });

    section("docx.blocks", || match document.blocks() {
        Err(error) => emit_error("docx.blocks", &error),
        Ok(blocks) => {
            emit("docx.blocks.len", &blocks.len().to_string());
            for (index, block) in blocks.iter().enumerate() {
                let kind = match block {
                    litchi_docx::Block::Paragraph(_) => "paragraph",
                    litchi_docx::Block::Table(_) => "table",
                    litchi_docx::Block::Alt(_) => "alt",
                    litchi_docx::Block::Unknown(opaque) => {
                        emit_markup(&format!("docx.block.{index}.xml"), opaque.xml_bytes());
                        "unknown"
                    },
                };
                emit(&format!("docx.block.{index}.kind"), kind);
            }
        },
    });

    section("docx.text", || match document.text() {
        Ok(text) => emit("docx.text", &escape(&text)),
        Err(error) => emit_error("docx.text", &error),
    });

    section("docx.sections", || match document.sections() {
        Ok(sections) => emit("docx.sections.len", &sections.len().to_string()),
        Err(error) => emit_error("docx.sections", &error),
    });

    section("docx.comments", || match document.comments() {
        Err(error) => emit_error("docx.comments", &error),
        Ok(comments) => {
            emit("docx.comments.len", &comments.len().to_string());
            for (index, comment) in comments.iter().enumerate() {
                emit(&format!("docx.cmt.{index}.id"), &comment.id().to_string());
                emit(&format!("docx.cmt.{index}.author"), &escape(comment.author()));
                emit(
                    &format!("docx.cmt.{index}.initials"),
                    &escape(comment.initials().unwrap_or("none")),
                );
                emit(
                    &format!("docx.cmt.{index}.date"),
                    &escape(comment.date().unwrap_or("none")),
                );
                match comment.text() {
                    Ok(text) => emit(&format!("docx.cmt.{index}.text"), &escape(&text)),
                    Err(error) => emit_error(&format!("docx.cmt.{index}.text"), &error),
                }
                emit_markup(&format!("docx.cmt.{index}.xml"), comment.xml_bytes());
            }
        },
    });

    for (kind, notes) in [
        ("fn", document.footnotes()),
        ("en", document.endnotes()),
    ] {
        section(&format!("docx.{kind}"), || match &notes {
            Err(error) => emit_error(&format!("docx.{kind}s"), error),
            Ok(notes) => {
                emit(&format!("docx.{kind}s.len"), &notes.len().to_string());
                for (index, note) in notes.iter().enumerate() {
                    emit(&format!("docx.{kind}.{index}.id"), &note.id().to_string());
                    emit(
                        &format!("docx.{kind}.{index}.type"),
                        &format!("{:?}", note.note_type()),
                    );
                    match note.text() {
                        Ok(text) => emit(&format!("docx.{kind}.{index}.text"), &escape(&text)),
                        Err(error) => emit_error(&format!("docx.{kind}.{index}.text"), &error),
                    }
                    emit_markup(&format!("docx.{kind}.{index}.xml"), note.xml_bytes());
                }
            },
        });
    }

    section("docx.source_backed", || {
        let source = match litchi_docx::source_backed::Package::from_path(path) {
            Ok(source) => {
                emit("docxsb.open", "OK");
                source
            },
            Err(error) => {
                emit_error("docxsb.open", &error);
                return;
            },
        };
        let document = match source.document() {
            Ok(document) => {
                emit("docxsb.document", "OK");
                document
            },
            Err(error) => {
                emit_error("docxsb.document", &error);
                return;
            },
        };
        match document.extract_text() {
            Ok(text) => emit("docxsb.text", &escape(&text)),
            Err(error) => emit_error("docxsb.text", &error),
        }
        match document.paragraph_count() {
            Err(error) => emit_error("docxsb.paragraph_count", &error),
            Ok(count) => {
                emit("docxsb.paragraph_count", &count.to_string());
                for index in 0..count.min(MAX_PARAGRAPHS) {
                    match document.paragraph_text(index) {
                        Ok(Some(text)) => {
                            emit(&format!("docxsb.par.{index}.text"), &escape(&text));
                        },
                        Ok(None) => emit(&format!("docxsb.par.{index}.text"), "none"),
                        Err(error) => emit_error(&format!("docxsb.par.{index}.text"), &error),
                    }
                }
            },
        }
    });
}

// -------------------------------------------------------------------- PPTX --

fn shape_kind(shape: &litchi_pptx::shape::Shape<'_>) -> &'static str {
    use litchi_pptx::shape::Shape;
    match shape {
        Shape::Auto(_) => "auto",
        Shape::Picture(_) => "picture",
        Shape::Table(_) => "table",
        Shape::Chart(_) => "chart",
        Shape::Diagram(_) => "diagram",
        Shape::Ole(_) => "ole",
        Shape::Frame(_) => "frame",
        Shape::Group(_) => "group",
        Shape::Connector(_) => "connector",
        Shape::Content(_) => "content",
        Shape::Unknown(_) => "unknown",
        _ => "other",
    }
}

fn pptx(path: &str) {
    use litchi_pptx::Package;

    let package = match Package::open(path) {
        Ok(package) => {
            emit("pptx.open", "OK");
            package
        },
        Err(error) => {
            emit_error("pptx.open", &error);
            return;
        },
    };
    let presentation = match package.presentation() {
        Ok(presentation) => {
            emit("pptx.presentation", "OK");
            presentation
        },
        Err(error) => {
            emit_error("pptx.presentation", &error);
            return;
        },
    };

    section("pptx.slide_count", || match presentation.slide_count() {
        Ok(count) => emit("pptx.slide_count", &count.to_string()),
        Err(error) => emit_error("pptx.slide_count", &error),
    });

    section("pptx.text", || match presentation.text() {
        Ok(text) => emit("pptx.text", &escape(&text)),
        Err(error) => emit_error("pptx.text", &error),
    });

    section("pptx.slides", || match presentation.slides() {
        Err(error) => emit_error("pptx.slides", &error),
        Ok(slides) => {
            emit("pptx.slides.len", &slides.len().to_string());
            for (index, slide) in slides.iter().enumerate() {
                match slide.name() {
                    Ok(name) => emit(&format!("pptx.sld.{index}.name"), &escape(&name)),
                    Err(error) => emit_error(&format!("pptx.sld.{index}.name"), &error),
                }
                match slide.text() {
                    Ok(text) => emit(&format!("pptx.sld.{index}.text"), &escape(&text)),
                    Err(error) => emit_error(&format!("pptx.sld.{index}.text"), &error),
                }
                match slide.shape_count() {
                    Ok(count) => {
                        emit(&format!("pptx.sld.{index}.shape_count"), &count.to_string());
                    },
                    Err(error) => emit_error(&format!("pptx.sld.{index}.shape_count"), &error),
                }
                match slide.shapes() {
                    Err(error) => emit_error(&format!("pptx.sld.{index}.shapes"), &error),
                    Ok(scene) => {
                        emit(&format!("pptx.sld.{index}.scene.len"), &scene.len().to_string());
                        emit(
                            &format!("pptx.sld.{index}.scene.rewritten"),
                            &scene.is_rewritten().to_string(),
                        );
                        for (position, shape) in scene.iter().enumerate() {
                            let prefix = format!("pptx.sld.{index}.shp.{position}");
                            emit(&format!("{prefix}.kind"), shape_kind(&shape));
                            emit(&format!("{prefix}.name"), &escape(shape.name().unwrap_or("none")));
                            emit(&format!("{prefix}.id"), &opt(shape.id()));
                            emit(&format!("{prefix}.text"), &escape(shape.text().unwrap_or("none")));
                            match shape.xml() {
                                Ok(xml) => {
                                    emit_markup(&format!("{prefix}.xml"), xml);
                                    if matches!(shape, litchi_pptx::shape::Shape::Table(_)) {
                                        match litchi_pptx::table::Table::from_graphic_frame(xml) {
                                            Err(error) => {
                                                emit_error(&format!("{prefix}.tbl"), &error);
                                            },
                                            Ok(table) => {
                                                match table.row_count() {
                                                    Ok(rows) => emit(
                                                        &format!("{prefix}.tbl.rows"),
                                                        &rows.to_string(),
                                                    ),
                                                    Err(error) => emit_error(
                                                        &format!("{prefix}.tbl.rows"),
                                                        &error,
                                                    ),
                                                }
                                                match table.column_count() {
                                                    Ok(columns) => emit(
                                                        &format!("{prefix}.tbl.cols"),
                                                        &columns.to_string(),
                                                    ),
                                                    Err(error) => emit_error(
                                                        &format!("{prefix}.tbl.cols"),
                                                        &error,
                                                    ),
                                                }
                                                match table.rows() {
                                                    Err(error) => emit_error(
                                                        &format!("{prefix}.tbl.rowlist"),
                                                        &error,
                                                    ),
                                                    Ok(rows) => {
                                                        for (row_index, row) in
                                                            rows.iter().enumerate()
                                                        {
                                                            match row.cells() {
                                                                Err(error) => emit_error(
                                                                    &format!(
                                                                        "{prefix}.tbl.row.{row_index}.cells"
                                                                    ),
                                                                    &error,
                                                                ),
                                                                Ok(cells) => {
                                                                    emit(
                                                                        &format!(
                                                                            "{prefix}.tbl.row.{row_index}.cellcount"
                                                                        ),
                                                                        &cells.len().to_string(),
                                                                    );
                                                                    for (cell_index, cell) in
                                                                        cells.iter().enumerate()
                                                                    {
                                                                        match cell.text() {
                                                                            Ok(text) => emit(
                                                                                &format!(
                                                                                    "{prefix}.tbl.row.{row_index}.cell.{cell_index}.text"
                                                                                ),
                                                                                &escape(&text),
                                                                            ),
                                                                            Err(error) => {
                                                                                emit_error(
                                                                                    &format!(
                                                                                        "{prefix}.tbl.row.{row_index}.cell.{cell_index}.text"
                                                                                    ),
                                                                                    &error,
                                                                                );
                                                                            },
                                                                        }
                                                                    }
                                                                },
                                                            }
                                                        }
                                                    },
                                                }
                                            },
                                        }
                                    }
                                },
                                Err(error) => emit_error(&format!("{prefix}.xml"), &error),
                            }
                        }
                    },
                }
            }
        },
    });

    section("pptx.source_backed", || {
        use litchi_pptx::presentation::SourceBackedPresentation;
        let source = match SourceBackedPresentation::from_path(path) {
            Ok(source) => {
                emit("pptxsb.open", "OK");
                source
            },
            Err(error) => {
                emit_error("pptxsb.open", &error);
                return;
            },
        };
        emit("pptxsb.slide_count", &source.slide_count().to_string());
        for (index, slide) in source.slides().enumerate() {
            match slide.name() {
                Ok(name) => emit(&format!("pptxsb.sld.{index}.name"), &escape(&name)),
                Err(error) => emit_error(&format!("pptxsb.sld.{index}.name"), &error),
            }
            match slide.text() {
                Ok(text) => emit(&format!("pptxsb.sld.{index}.text"), &escape(&text)),
                Err(error) => emit_error(&format!("pptxsb.sld.{index}.text"), &error),
            }
            match slide.images() {
                Err(error) => emit_error(&format!("pptxsb.sld.{index}.images"), &error),
                Ok(images) => {
                    emit(
                        &format!("pptxsb.sld.{index}.images.len"),
                        &images.len().to_string(),
                    );
                    for (image_index, image) in images.iter().enumerate() {
                        emit(
                            &format!("pptxsb.sld.{index}.img.{image_index}"),
                            &escape(&format!(
                                "position={};shape_position={};id={};name={};rel={};external={}",
                                image.position(),
                                image.shape_position(),
                                opt(image.id()),
                                image.name().unwrap_or("none"),
                                image.relationship_id(),
                                image.target().is_external()
                            )),
                        );
                    }
                },
            }
        }
    });
}

// -------------------------------------------------------------------- XLSX --

fn xlsx(path: &str) {
    use litchi_xlsx::Workbook;

    let workbook = match Workbook::open(path) {
        Ok(workbook) => {
            emit("xlsx.open", "OK");
            workbook
        },
        Err(error) => {
            emit_error("xlsx.open", &error);
            return;
        },
    };
    emit("xlsx.sheets.len", &workbook.len().to_string());

    for (index, sheet) in workbook.sheets().enumerate() {
        let prefix = format!("xlsx.sheet.{index}");
        emit(&format!("{prefix}.name"), &escape(sheet.name()));
        emit(&format!("{prefix}.kind"), &format!("{:?}", sheet.kind()));

        section(&format!("{prefix}.extents"), || match sheet.extents() {
            Err(error) => emit_error(&format!("{prefix}.extents"), &error),
            Ok(extents) => {
                emit(
                    &format!("{prefix}.extents"),
                    &escape(&format!(
                        "declared={:?};stored={:?};content={:?};styled={:?}",
                        extents.declared(),
                        extents.stored(),
                        extents.content(),
                        extents.styled()
                    )),
                );
            },
        });

        section(&format!("{prefix}.cells"), || {
            match sheet.cells("A1:XFD1048576") {
                Err(error) => emit_error(&format!("{prefix}.cells"), &error),
                Ok(cells) => {
                    let mut count = 0usize;
                    let mut truncated = false;
                    for (address, cell) in cells {
                        if count >= MAX_CELLS_PER_SHEET {
                            truncated = true;
                            break;
                        }
                        count += 1;
                        emit(
                            &format!("{prefix}.cell.{address}"),
                            &escape(&format!("{cell:?}")),
                        );
                    }
                    emit(&format!("{prefix}.cells.emitted"), &count.to_string());
                    if truncated {
                        emit(
                            &format!("{prefix}.cells.capped"),
                            &format!("at={MAX_CELLS_PER_SHEET}"),
                        );
                    }
                },
            }
        });

        section(&format!("{prefix}.views"), || match sheet.views() {
            Err(error) => emit_error(&format!("{prefix}.views"), &error),
            Ok(None) => emit(&format!("{prefix}.views"), "none"),
            Ok(Some(collection)) => {
                emit(
                    &format!("{prefix}.views.entries"),
                    &collection.entries().len().to_string(),
                );
                for (entry_index, entry) in collection.entries().iter().enumerate() {
                    emit_markup(
                        &format!("{prefix}.view.{entry_index}.retained_xml"),
                        entry.retained_xml(),
                    );
                    for (extension_index, extension) in entry.extensions().iter().enumerate() {
                        emit(
                            &format!("{prefix}.view.{entry_index}.ext.{extension_index}.uri"),
                            &escape(extension.uri()),
                        );
                        emit_markup(
                            &format!("{prefix}.view.{entry_index}.ext.{extension_index}.markup"),
                            extension.markup(),
                        );
                    }
                    for (selection_index, selection) in entry.pivot_selections().iter().enumerate() {
                        emit_markup(
                            &format!(
                                "{prefix}.view.{entry_index}.pivot.{selection_index}.area_markup"
                            ),
                            selection.area().markup(),
                        );
                    }
                }
                for (extension_index, extension) in collection.extensions().iter().enumerate() {
                    emit(
                        &format!("{prefix}.views.ext.{extension_index}.uri"),
                        &escape(extension.uri()),
                    );
                    emit_markup(
                        &format!("{prefix}.views.ext.{extension_index}.markup"),
                        extension.markup(),
                    );
                }
            },
        });

        section(&format!("{prefix}.ignored_errors"), || {
            match sheet.ignored_errors() {
                Err(error) => emit_error(&format!("{prefix}.ignored_errors"), &error),
                Ok(None) => emit(&format!("{prefix}.ignored_errors"), "none"),
                Ok(Some(errors)) => {
                    emit(
                        &format!("{prefix}.ignored_errors.entries"),
                        &errors.entries().len().to_string(),
                    );
                    for (entry_index, entry) in errors.entries().iter().enumerate() {
                        emit(
                            &format!("{prefix}.ierr.{entry_index}.ranges"),
                            &escape(&format!(
                                "{:?}",
                                entry
                                    .ranges()
                                    .iter()
                                    .map(litchi_xlsx::IgnoredErrorRangeReference::as_str)
                                    .collect::<Vec<_>>()
                            )),
                        );
                    }
                    for (extension_index, extension) in errors.extensions().iter().enumerate() {
                        emit(
                            &format!("{prefix}.ierr.ext.{extension_index}.uri"),
                            &escape(extension.uri()),
                        );
                        emit_markup(
                            &format!("{prefix}.ierr.ext.{extension_index}.markup"),
                            extension.markup(),
                        );
                    }
                },
            }
        });

        section(&format!("{prefix}.named_sheet_views"), || {
            match sheet.named_sheet_views() {
                Err(error) => emit_error(&format!("{prefix}.named_sheet_views"), &error),
                Ok(None) => emit(&format!("{prefix}.named_sheet_views"), "none"),
                Ok(Some(views)) => {
                    emit(
                        &format!("{prefix}.nsv.views"),
                        &views.views().len().to_string(),
                    );
                    for (extension_index, extension) in views.extensions().iter().enumerate() {
                        emit(
                            &format!("{prefix}.nsv.ext.{extension_index}.uri"),
                            &escape(extension.uri()),
                        );
                        emit_markup(
                            &format!("{prefix}.nsv.ext.{extension_index}.markup"),
                            extension.markup().xml(),
                        );
                    }
                },
            }
        });

        section(&format!("{prefix}.conditional_formattings"), || {
            match sheet.conditional_formattings() {
                Err(error) => emit_error(&format!("{prefix}.cf"), &error),
                Ok(formattings) => {
                    emit(&format!("{prefix}.cf.len"), &formattings.len().to_string());
                    for (formatting_index, formatting) in formattings.iter().enumerate() {
                        emit(
                            &format!("{prefix}.cf.{formatting_index}.ranges"),
                            &escape(&format!("{:?}", formatting.ranges)),
                        );
                        emit(
                            &format!("{prefix}.cf.{formatting_index}.rules"),
                            &formatting.rules.len().to_string(),
                        );
                        for (rule_index, rule) in formatting.rules.iter().enumerate() {
                            emit(
                                &format!("{prefix}.cf.{formatting_index}.rule.{rule_index}"),
                                &escape(&format!(
                                    "type={:?};priority={:?};formulas={:?};text={:?}",
                                    rule.rule_type, rule.priority, rule.formulas, rule.text
                                )),
                            );
                            if let Some(litchi_xlsx::DifferentialRef::Inline(differential)) =
                                rule.differential_format.as_ref()
                            {
                                emit_markup(
                                    &format!(
                                        "{prefix}.cf.{formatting_index}.rule.{rule_index}.dxf"
                                    ),
                                    differential.raw_xml(),
                                );
                            }
                        }
                    }
                },
            }
        });
    }

    section("xlsx.chain", || {
        let package = match litchi_opc::OpcPackage::open(path) {
            Ok(package) => package,
            Err(error) => {
                emit_error("xlsx.chain.open", &error);
                return;
            },
        };
        match litchi_xlsx::chain::load(&package) {
            Err(error) => emit_error("xlsx.chain", &error),
            Ok(None) => emit("xlsx.chain", "none"),
            Ok(Some((chain, conformance))) => {
                emit("xlsx.chain.len", &chain.len().to_string());
                emit("xlsx.chain.conformance", &format!("{conformance:?}"));
                match chain.extension_list_xml() {
                    None => emit("xlsx.chain.extlst", "none"),
                    Some(xml) => emit_markup("xlsx.chain.extlst", xml.as_bytes()),
                }
                for (index, cell) in chain.cells().iter().enumerate().take(MAX_CELLS_PER_SHEET) {
                    emit(
                        &format!("xlsx.chain.cell.{index}"),
                        &escape(&format!("ref={}", cell.reference())),
                    );
                }
            },
        }
    });

    section("xlsx.source_backed", || {
        use litchi_xlsx::SourceBackedWorkbook;
        let workbook = match SourceBackedWorkbook::from_path(path) {
            Ok(workbook) => {
                emit("xlsxsb.open", "OK");
                workbook
            },
            Err(error) => {
                emit_error("xlsxsb.open", &error);
                return;
            },
        };
        emit("xlsxsb.sheets.len", &workbook.len().to_string());
        for (index, sheet) in workbook.sheets().enumerate() {
            let prefix = format!("xlsxsb.sheet.{index}");
            emit(&format!("{prefix}.name"), &escape(sheet.name()));
            match sheet.stored_extent() {
                Ok(extent) => emit(
                    &format!("{prefix}.stored_extent"),
                    &escape(&format!("{extent:?}")),
                ),
                Err(error) => emit_error(&format!("{prefix}.stored_extent"), &error),
            }
            match sheet.cells("A1:XFD1048576") {
                Err(error) => emit_error(&format!("{prefix}.cells"), &error),
                Ok(cells) => {
                    emit(
                        &format!("{prefix}.cells.emitted"),
                        &cells.len().min(MAX_CELLS_PER_SHEET).to_string(),
                    );
                    if cells.len() > MAX_CELLS_PER_SHEET {
                        emit(
                            &format!("{prefix}.cells.capped"),
                            &format!("at={MAX_CELLS_PER_SHEET}"),
                        );
                    }
                    for entry in cells.iter().take(MAX_CELLS_PER_SHEET) {
                        emit(
                            &format!("{prefix}.cell.{}", entry.address),
                            &escape(&format!("{:?}", entry.cell)),
                        );
                    }
                },
            }
        }
    });
}

// -------------------------------------------------------------------- main --

fn main() {
    let arguments: Vec<String> = std::env::args().collect();
    if arguments.len() < 3 {
        eprintln!("usage: readprobe <fixture-path> <display-path> [--dump]");
        std::process::exit(2);
    }
    let path = arguments[1].clone();
    let _ = DISPLAY.set(arguments[2].clone());
    let _ = DUMP.set(arguments.iter().any(|argument| argument == "--dump"));

    // Panics are an outcome, not a crash: suppress the default hook's output so
    // that stderr stays comparable between the two legs.
    std::panic::set_hook(Box::new(|_| {}));

    let lower = path.to_ascii_lowercase();
    section("probe", || {
        if lower.ends_with(".docx") {
            docx(&path);
        } else if lower.ends_with(".pptx") {
            pptx(&path);
        } else if lower.ends_with(".xlsx") {
            xlsx(&path);
        } else {
            emit("probe.skipped", "unsupported extension");
        }
    });
    let _ = std::io::stdout().flush();
}
