//! 0768 diagnostic: locate the CHPX FKP page an edited output fails on.
use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::parts::fkp::ChpxFkp;
use litchi_doc::tracked_revision::{Limits, RevisionEditor, RevisionKind, RevisionMetadata};
use std::io::Cursor;

fn u32_at(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

fn scan(label: &str, bytes: &[u8]) {
    let mut ole = OleFile::open(Cursor::new(bytes.to_vec())).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word).unwrap();
    let table = ole
        .open_stream(&[if fib.which_table_stream() { "1Table" } else { "0Table" }])
        .unwrap();
    let (fc, lcb) = fib.get_table_pointer(12).unwrap();
    let plc = &table[fc as usize..(fc + lcb) as usize];
    let n = (plc.len() - 4) / 8;
    println!("{label}: word={} bte_chpx pages={n}", word.len());
    for i in 0..n {
        let pn = u32_at(plc, (n + 1) * 4 + i * 4) as usize;
        let page = &word[pn * 512..pn * 512 + 512];
        let crun = page[511] as usize;
        let ok = ChpxFkp::parse(page, &word).is_some();
        let fcs: Vec<u32> = (0..=crun.min(101)).map(|k| u32_at(page, k * 4)).collect();
        let bxs: Vec<u8> = (0..crun.min(101)).map(|k| page[(crun + 1) * 4 + k]).collect();
        let bad_fc = fcs.windows(2).position(|w| w[0] >= w[1]);
        let bad_bx = bxs.iter().position(|&bx| {
            let off = usize::from(bx) * 2;
            bx != 0
                && (off < (crun + 1) * 4 + crun
                    || off >= 511
                    || off + 1 + usize::from(page[off.min(511)]) > 511)
        });
        if !ok || i + 1 == n {
            println!(
                "  page {i} pn={pn} crun={crun} ok={ok} bte=[{}..{}) first_fc={} last_fc={} bad_fc_at={bad_fc:?} bad_bx_at={bad_bx:?}",
                u32_at(plc, i * 4),
                u32_at(plc, (i + 1) * 4),
                fcs.first().unwrap_or(&0),
                fcs.last().unwrap_or(&0),
            );
            if !ok {
                let prop_start = (crun + 1) * 4 + crun;
                for (k, &bx) in bxs.iter().enumerate() {
                    let off = usize::from(bx) * 2;
                    let cb = if bx != 0 && off < 512 { page[off] } else { 0 };
                    if bx == 0 || off < prop_start || off + 1 + usize::from(cb) > 511 {
                        println!("    bad entry {k}: bx={bx} off={off} prop_start={prop_start} cb={cb}");
                    }
                }
                let mut offs: Vec<usize> = bxs.iter().filter(|&&b| b != 0).map(|&b| usize::from(b) * 2).collect();
                offs.sort();
                println!("    lowest grpprl offsets: {:?}", &offs[..offs.len().min(4)]);
            }
            if let Some(at) = bad_fc {
                println!("    fcs around: {:?}", &fcs[at.saturating_sub(2)..(at + 3).min(fcs.len())]);
            }
        }
    }
}

fn main() {
    let path = std::env::args().nth(1).expect("path");
    let bytes = std::fs::read(&path).unwrap();
    scan("input", &bytes);
    let mut editor = RevisionEditor::open(bytes, Limits::default()).unwrap();
    editor
        .add_text(0, "x", RevisionKind::Insertion, RevisionMetadata::new("diag"))
        .unwrap();
    let output = editor.finish().unwrap();
    scan("output", &output);
}
