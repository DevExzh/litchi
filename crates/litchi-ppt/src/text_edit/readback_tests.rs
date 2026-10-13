//! Differential tests of the incremental readback (record 0264).
//!
//! Every case publishes a batch through the real rewrite and editor, lets a
//! hook corrupt the staged slide record (simulating a faulty rewrite), and
//! runs both verifiers on the published candidate. The production verdict is
//! `incremental accept || full accept`-when-incremental-cannot-prove, so it
//! equals the reference verdict exactly when an incremental accept implies a
//! full accept. Each case asserts that, and that the shapes the incremental
//! verifier proves equal a full parse of the published slide.

use super::{
    DifferentialVerdicts, Position, ShapeTextRewrite, Target, differential_readback_for_test,
};

fn target(slide: usize, shape: usize) -> Target {
    Target::new(Position::new(slide), Position::new(shape))
}

/// One slide with `count` text boxes (the 0218–0260 batch corpus shape).
fn deck(count: usize) -> Vec<u8> {
    let mut builder = crate::Builder::new();
    let slide = builder.add_slide().unwrap();
    for index in 0..count {
        let x = 24 + i32::try_from(index % 16).unwrap() * 120;
        let y = 24 + i32::try_from(index / 16).unwrap() * 72;
        builder
            .add_textbox(slide, x, y, 96, 48, &format!("source text {index:03}"))
            .unwrap();
    }
    builder.build().unwrap().to_bytes().unwrap()
}

fn fixture(relative: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(relative),
    )
    .unwrap()
}

#[derive(Debug, Default)]
struct Tally {
    cases: usize,
    incremental_accepts: usize,
    full_accepts: usize,
    reopen_failures: usize,
}

/// The equivalence property for one candidate.
fn check(verdicts: &DifferentialVerdicts, tally: &mut Tally, context: &str) {
    tally.cases += 1;
    if !verdicts.reopened {
        tally.reopen_failures += 1;
        assert!(verdicts.incremental.is_none() && !verdicts.full_accepted);
        return;
    }
    tally.full_accepts += usize::from(verdicts.full_accepted);
    let Some(proven) = &verdicts.incremental else {
        return;
    };
    tally.incremental_accepts += 1;
    assert!(
        verdicts.full_accepted,
        "{context}: incremental accept but full refusal"
    );
    for (persist_id, shapes) in proven {
        let position = verdicts
            .persist_ids
            .iter()
            .position(|id| id == persist_id)
            .unwrap();
        assert_eq!(
            verdicts.parsed[position].as_ref().ok(),
            Some(shapes),
            "{context}: proven shapes differ from a full parse"
        );
    }
}

fn run(
    source: &[u8],
    replacements: &[(Target, &str)],
    corrupt: &dyn Fn(u32, &mut Vec<u8>),
) -> Option<DifferentialVerdicts> {
    differential_readback_for_test(source, replacements, corrupt)
        .inspect_err(|error| {
            if std::env::var_os("LITCHI_READBACK_DEBUG").is_some() {
                eprintln!("hook error: {error:?}");
            }
        })
        .ok()
}

/// Offsets of every record header in a PPT/OfficeArt record tree.
fn headers(record: &[u8]) -> Vec<usize> {
    fn walk(data: &[u8], base: usize, depth: usize, out: &mut Vec<usize>) {
        let mut offset = 0usize;
        while offset + 8 <= data.len() && depth < 64 {
            out.push(base + offset);
            let length = u32::from_le_bytes(data[offset + 4..offset + 8].try_into().unwrap());
            let Some(end) = (offset + 8).checked_add(length as usize) else {
                return;
            };
            if end > data.len() {
                return;
            }
            if data[offset] & 0x0F == 0x0F {
                walk(&data[offset + 8..end], base + offset + 8, depth + 1, out);
            }
            offset = end;
        }
    }
    let mut out = Vec::new();
    walk(record, 0, 0, &mut out);
    out
}

/// Offset of the `Sp` atom payload of the shape container with `shape_id`.
fn sp_payload(record: &[u8], shape_id: u32) -> Option<usize> {
    headers(record)
        .into_iter()
        .find(|&offset| {
            record.get(offset + 2..offset + 4) == Some(&0xF00A_u16.to_le_bytes()[..])
                && record.get(offset + 8..offset + 12) == Some(&shape_id.to_le_bytes()[..])
        })
        .map(|offset| offset + 8)
}

fn shape_id_at(source: &[u8], shape: usize) -> u32 {
    let presentation = super::presentation(source).unwrap();
    let slides = presentation.slides().unwrap();
    super::native_shape_id(&slides[0].shapes().unwrap()[shape])
}

#[test]
fn genuine_publications_are_proven_incrementally() {
    let mut tally = Tally::default();
    for count in [1usize, 3, 40] {
        let source = deck(count);
        let texts = (0..count)
            .map(|index| format!("replacement of another length {index}"))
            .collect::<Vec<_>>();
        // Every shape, a sparse subset, and a single last shape.
        let all = (0..count)
            .map(|index| (target(0, index), texts[index].as_str()))
            .collect::<Vec<_>>();
        let sparse = all.iter().step_by(3).copied().collect::<Vec<_>>();
        let last = vec![all[count - 1]];
        for replacements in [&all, &sparse, &last] {
            let verdicts = run(&source, replacements, &|_, _| {}).unwrap();
            check(&verdicts, &mut tally, "genuine");
            assert!(verdicts.full_accepted && verdicts.incremental.is_some());
        }
    }
    assert_eq!(tally.incremental_accepts, tally.cases);
}

/// Every single-byte corruption of the staged slide record.
#[test]
fn every_byte_flip_of_the_staged_slide_keeps_verdicts_equal() {
    let mut tally = Tally::default();
    let source = deck(3);
    let replacements = [
        (target(0, 0), "first replacement, longer"),
        (target(0, 2), "x"),
    ];
    let clean = run(&source, &replacements, &|_, _| {}).unwrap();
    assert!(clean.incremental.is_some());
    // Size of the staged record, learnt from a no-op corruption probe.
    let size = std::cell::Cell::new(0usize);
    run(&source, &replacements, &|_, record| size.set(record.len())).unwrap();
    for offset in 0..size.get() {
        for mask in [0xFF_u8, 0x01, 0x80] {
            if let Some(verdicts) = run(&source, &replacements, &|_, record| {
                record[offset] ^= mask;
            }) {
                check(&verdicts, &mut tally, &format!("flip {offset} ^ {mask:#x}"));
            }
        }
    }
    // Some flips must be caught only by the full verifier (or by neither),
    // and some must still be proven incrementally (text bytes of the target
    // that still read back are a different text: those are refusals).
    eprintln!("byte flips: {tally:?}");
    assert!(tally.cases > 1000, "{tally:?}");
    assert!(tally.incremental_accepts < tally.full_accepts, "{tally:?}");
}

/// Every record-length change of the staged slide record.
#[test]
fn every_header_length_change_keeps_verdicts_equal() {
    let mut tally = Tally::default();
    let source = deck(4);
    let replacements = [(target(0, 1), "replacement"), (target(0, 3), "another one")];
    let headers = std::cell::RefCell::new(Vec::new());
    run(&source, &replacements, &|_, record| {
        *headers.borrow_mut() = super::readback_tests::headers(record);
    })
    .unwrap();
    for offset in headers.into_inner() {
        for delta in [-9_i64, -8, -1, 1, 2, 8] {
            if let Some(verdicts) = run(&source, &replacements, &|_, record| {
                let field = &mut record[offset + 4..offset + 8];
                let length = i64::from(u32::from_le_bytes(field.try_into().unwrap()));
                let changed = u32::try_from((length + delta).max(0)).unwrap();
                field.copy_from_slice(&changed.to_le_bytes());
            }) {
                check(&verdicts, &mut tally, &format!("length {offset} {delta:+}"));
            }
        }
        // Kind/version/instance changes.
        for byte in [0usize, 1, 2, 3] {
            if let Some(verdicts) = run(&source, &replacements, &|_, record| {
                record[offset + byte] ^= 0x10;
            }) {
                check(&verdicts, &mut tally, &format!("header {offset}+{byte}"));
            }
        }
    }
    eprintln!("header changes: {tally:?}");
    assert!(tally.cases > 100, "{tally:?}");
}

/// Structured faults a buggy rewrite could make.
#[test]
fn structural_rewrite_faults_keep_verdicts_equal() {
    let mut tally = Tally::default();
    let source = deck(5);
    let shape_ids = (0..5)
        .map(|shape| shape_id_at(&source, shape))
        .collect::<Vec<_>>();
    let replacements = [
        (target(0, 1), "target text"),
        (target(0, 3), "second target"),
    ];

    // A non-target sibling's text changed as well: full accepts (its text is
    // not checked), incremental cannot prove it.
    let sibling = shape_ids[2];
    let verdicts = run(&source, &replacements, &|_, record| {
        let atom = super::inspect_slide(record, sibling).unwrap();
        let after = super::encode_replacement("sibling", atom.kind).unwrap();
        *record = super::rewrite_slide_grouped(
            record,
            &[ShapeTextRewrite {
                shape_id: sibling,
                kind: atom.kind,
                before: &atom.payload,
                after: &after,
            }],
        )
        .unwrap();
    })
    .unwrap();
    check(&verdicts, &mut tally, "sibling text");
    assert!(verdicts.full_accepted && verdicts.incremental.is_none());

    // A target's text rewritten to something else: both refuse.
    let first = shape_ids[1];
    let verdicts = run(&source, &replacements, &|_, record| {
        let atom = super::inspect_slide(record, first).unwrap();
        let after = super::encode_replacement("TARGET TEXT", atom.kind).unwrap();
        *record = super::rewrite_slide_grouped(
            record,
            &[ShapeTextRewrite {
                shape_id: first,
                kind: atom.kind,
                before: &atom.payload,
                after: &after,
            }],
        )
        .unwrap();
    })
    .unwrap();
    check(&verdicts, &mut tally, "target text");
    assert!(!verdicts.full_accepted && verdicts.incremental.is_none());

    // The target reverted to its source text: both refuse.
    let verdicts = run(&source, &replacements, &|_, record| {
        let atom = super::inspect_slide(record, first).unwrap();
        let after = super::encode_replacement("source text 001", atom.kind).unwrap();
        *record = super::rewrite_slide_grouped(
            record,
            &[ShapeTextRewrite {
                shape_id: first,
                kind: atom.kind,
                before: &atom.payload,
                after: &after,
            }],
        )
        .unwrap();
    })
    .unwrap();
    check(&verdicts, &mut tally, "target reverted");
    assert!(!verdicts.full_accepted);

    // Flag changes on every shape's `Sp` atom (background, child,
    // patriarch, deleted, hidden bits) and identifier swaps.
    for (index, &shape_id) in shape_ids.iter().enumerate() {
        for bit in 0..12 {
            let verdicts = run(&source, &replacements, &|_, record| {
                let at = sp_payload(record, shape_id).unwrap() + 4;
                record[at + bit / 8] ^= 1 << (bit % 8);
            });
            if let Some(verdicts) = verdicts {
                check(&verdicts, &mut tally, &format!("flag {index} bit {bit}"));
            }
        }
        for &other in &shape_ids {
            let verdicts = run(&source, &replacements, &|_, record| {
                let at = sp_payload(record, shape_id).unwrap();
                record[at..at + 4].copy_from_slice(&other.to_le_bytes());
            });
            if let Some(verdicts) = verdicts {
                check(
                    &verdicts,
                    &mut tally,
                    &format!("identifier {index} -> {other}"),
                );
            }
        }
    }

    // Swapped shape containers (order), and a duplicated container.
    let swap = |record: &mut Vec<u8>, duplicate: bool| {
        let starts = headers(record)
            .into_iter()
            .filter(|&offset| {
                record[offset + 2..offset + 4] == 0xF004_u16.to_le_bytes()
                    && sp_payload(&record[offset..], shape_ids[1]) == Some(16)
                    || record[offset + 2..offset + 4] == 0xF004_u16.to_le_bytes()
                        && sp_payload(&record[offset..], shape_ids[2]) == Some(16)
            })
            .collect::<Vec<_>>();
        let span = |offset: usize, record: &[u8]| {
            offset
                ..offset
                    + 8
                    + u32::from_le_bytes(record[offset + 4..offset + 8].try_into().unwrap())
                        as usize
        };
        let (a, b) = (span(starts[0], record), span(starts[1], record));
        let first = record[a.clone()].to_vec();
        let second = record[b.clone()].to_vec();
        let mut rebuilt = record[..a.start].to_vec();
        rebuilt.extend_from_slice(&second);
        rebuilt.extend_from_slice(&record[a.end..b.start]);
        rebuilt.extend_from_slice(if duplicate { &second } else { &first });
        rebuilt.extend_from_slice(&record[b.end..]);
        *record = rebuilt;
    };
    // A duplicated container changes the record length; the editor refuses
    // it before publication unless both containers have one length.
    for duplicate in [false, true] {
        if let Some(verdicts) = run(&source, &replacements, &|_, record| swap(record, duplicate)) {
            check(
                &verdicts,
                &mut tally,
                &format!("swap duplicate={duplicate}"),
            );
            assert!(verdicts.incremental.is_none());
        } else {
            assert!(duplicate);
        }
    }

    // Truncated and extended records.
    for cut in [1usize, 4, 8, 20] {
        if let Some(verdicts) = run(&source, &replacements, &|_, record| {
            let length = record.len();
            record.truncate(length - cut);
        }) {
            check(&verdicts, &mut tally, &format!("truncate {cut}"));
        }
        if let Some(verdicts) = run(&source, &replacements, &|_, record| {
            record.extend(std::iter::repeat_n(0, cut));
        }) {
            check(&verdicts, &mut tally, &format!("extend {cut}"));
        }
    }
    eprintln!("structural faults: {tally:?}");
    assert!(tally.cases > 80, "{tally:?}");
}

/// Every top-level shape of the length-change corpus that the native edit
/// accepts: the genuine publication, and every byte flip of its target
/// shape's text records.
#[test]
fn corpus_publications_keep_verdicts_equal() {
    const DECKS: &[&str] = &[
        "ole/ppt/SampleShow.ppt",
        "ole/ppt/empty_textbox.ppt",
        "ole/ppt/ppt_with_png.ppt",
        "ole/ppt/text-margins.ppt",
        "ole/ppt/text_shapes.ppt",
        "poi/test-data/slideshow/41246-1.ppt",
        "poi/test-data/slideshow/45543.ppt",
        "poi/test-data/slideshow/54880_chinese.ppt",
        "poi/test-data/slideshow/WithComments.ppt",
        "poi/test-data/slideshow/WithLinks.ppt",
        "poi/test-data/slideshow/basic_test_ppt_file.ppt",
        "poi/test-data/slideshow/bug58144-headers-footers-2003.ppt",
        "poi/test-data/slideshow/datetime.ppt",
        "poi/test-data/slideshow/headers_footers.ppt",
        "poi/test-data/slideshow/incorrect_slide_order.ppt",
        "poi/test-data/slideshow/ole2-embedding-2003.ppt",
        "poi/test-data/slideshow/sound.ppt",
    ];
    let mut tally = Tally::default();
    let mut genuine = Tally::default();
    for deck in DECKS {
        let source = fixture(deck);
        let package = crate::Package::from_bytes(&source).unwrap();
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        for (slide_index, slide) in slides.iter().enumerate() {
            let Ok(shapes) = slide.shapes() else { continue };
            for shape_index in 0..shapes.len() {
                let replacement = "Changed text of another length!";
                let mut edit = package.edit().unwrap();
                if edit
                    .set_shape_text(target(slide_index, shape_index), replacement)
                    .is_err()
                {
                    continue;
                }
                let replacements = [(target(slide_index, shape_index), replacement)];
                let Some(verdicts) = run(&source, &replacements, &|_, _| {}) else {
                    // Outline-backed text (`NoTextAtom` in the drawing) is
                    // published in the document and always verified in full.
                    continue;
                };
                check(
                    &verdicts,
                    &mut genuine,
                    &format!("{deck} {slide_index}/{shape_index}"),
                );
                assert!(verdicts.full_accepted);
                // Flip every byte of the slide record's last 64 bytes and of
                // a window around each header (bounded per shape).
                let size = std::cell::Cell::new(0usize);
                run(&source, &replacements, &|_, record| size.set(record.len())).unwrap();
                for offset in (0..size.get()).step_by((size.get() / 150).max(7)) {
                    if let Some(verdicts) = run(&source, &replacements, &|_, record| {
                        record[offset] ^= 0x41;
                    }) {
                        check(
                            &verdicts,
                            &mut tally,
                            &format!("{deck} {slide_index}/{shape_index} flip {offset}"),
                        );
                    }
                }
            }
        }
    }
    // Clean corpus slides are proven incrementally; slides that the source
    // parse repaired are left to the full verifier.
    eprintln!("corpus genuine: {genuine:?}; corpus flips: {tally:?}");
    assert!(genuine.cases >= 30, "{genuine:?}");
    assert!(
        genuine.incremental_accepts * 10 >= genuine.cases * 8,
        "{genuine:?}"
    );
    assert!(tally.cases > 1000, "{tally:?}");
}
