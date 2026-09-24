//! 0768 probe: per-fixture DOC protection classification and tracked-edit outcome.
//!
//! Run as a temporary `litchi-doc` example (copied to `crates/litchi-doc/examples/`)
//! identically on the base and the candidate tree:
//! `cargo run -p litchi-doc --example probe_0768 --locked --offline -- <dir>`.
//! Prints one JSON object per `.doc` fixture on stdout.

use litchi_cfb::OleFile;
use litchi_doc::package::property_set::Snapshot;
use litchi_doc::tracked_revision::{Limits, RevisionEditor, RevisionKind, RevisionMetadata};
use std::io::Cursor;

fn u16_at(data: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(data.get(offset..offset + 2)?.try_into().ok()?))
}

fn u32_at(data: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(data.get(offset..offset + 4)?.try_into().ok()?))
}

fn json_str(value: &str) -> String {
    serde_json::to_string(value).unwrap()
}

fn edit(bytes: &[u8], cp: u32) -> String {
    let mut editor = match RevisionEditor::open(bytes.to_vec(), Limits::default()) {
        Ok(editor) => editor,
        Err(error) => return format!("open: {error}"),
    };
    if let Err(error) = editor.add_text(
        cp,
        "x",
        RevisionKind::Insertion,
        RevisionMetadata::new("probe-0768"),
    ) {
        return format!("add_text: {error}");
    }
    match editor.finish() {
        Ok(output) => {
            // The edited output must reopen and edit-free finish must be byte exact.
            match RevisionEditor::open(output.clone(), Limits::default()) {
                Ok(reopened) => match reopened.finish() {
                    Ok(again) if again == output => "ok".to_string(),
                    Ok(_) => "reopen-finish-differs".to_string(),
                    Err(error) => format!("reopen-finish: {error}"),
                },
                Err(error) => format!("reopen: {error}"),
            }
        },
        Err(error) => format!("finish: {error}"),
    }
}

fn main() {
    let dir = std::env::args().nth(1).expect("fixture directory");
    let mut paths = std::fs::read_dir(&dir)
        .unwrap()
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| path.extension().is_some_and(|ext| ext == "doc"))
        .collect::<Vec<_>>();
    paths.sort();
    for path in paths {
        let name = path.file_name().unwrap().to_string_lossy().to_string();
        let bytes = std::fs::read(&path).unwrap();
        let mut fields = vec![format!("\"file\":{}", json_str(&name))];
        let shape = (|| -> Option<Vec<String>> {
            let mut ole = OleFile::open(Cursor::new(bytes.clone())).ok()?;
            let word = ole.open_stream(&["WordDocument"]).ok()?;
            let nfib = u16_at(&word, 2)?;
            let flags = u16_at(&word, 10)?;
            let count = usize::from(u16_at(&word, 152)?);
            let pointer_end = 154 + count * 8;
            let csw_new = u16_at(&word, pointer_end);
            let nfib_new = u16_at(&word, pointer_end + 2);
            let dop_pointer = 154 + 31 * 8;
            let (fc_dop, lcb_dop) = if count > 31 {
                (u32_at(&word, dop_pointer)?, u32_at(&word, dop_pointer + 4)?)
            } else {
                (0, 0)
            };
            let mut out = vec![
                format!("\"nfib\":\"0x{nfib:04X}\""),
                format!("\"magic34\":\"0x{:04X}\"", u16_at(&word, 34)?),
                format!("\"magic36\":\"0x{:04X}\"", u16_at(&word, 36)?),
                format!("\"cb_rg_fc_lcb\":\"0x{count:04X}\""),
                format!(
                    "\"csw_new\":{}",
                    csw_new.map_or("null".to_string(), |v| v.to_string())
                ),
                format!(
                    "\"nfib_new\":{}",
                    nfib_new.map_or("null".to_string(), |v| format!("\"0x{v:04X}\""))
                ),
                format!("\"encrypted\":{}", flags & 0x0100 != 0),
                format!("\"lcb_dop\":{lcb_dop}"),
            ];
            let table_name = if flags & 0x0200 != 0 { "1Table" } else { "0Table" };
            if let Ok(table) = ole.open_stream(&[table_name]) {
                let start = fc_dop as usize;
                if let Some(dop) = table.get(start..start + lcb_dop as usize)
                    && dop.len() >= 84
                {
                    out.push(format!("\"f_rev_marking\":{}", dop[5] & 0x80 != 0));
                    out.push(format!("\"f_form_no_fields\":{}", dop[5] & 0x20 != 0));
                    out.push(format!("\"f_lock_atn\":{}", dop[6] & 0x10 != 0));
                    out.push(format!("\"f_prot_enabled\":{}", dop[7] & 0x02 != 0));
                    out.push(format!("\"f_lock_rev\":{}", dop[7] & 0x40 != 0));
                    out.push(format!("\"l_key_prot_doc\":{}", u32_at(dop, 78)?));
                    out.push(format!("\"wvko_saved\":{}", u16_at(dop, 82)? & 0x07));
                    out.push(format!("\"w_spare2\":{}", u16_at(dop, 18)?));
                    if dop.len() >= 600 {
                        let word598 = u16_at(dop, 598)?;
                        out.push(format!("\"word598\":\"0x{word598:04X}\""));
                        out.push(format!("\"f_enforce_doc_prot\":{}", word598 & 0x0008 != 0));
                        out.push(format!("\"i_doc_prot_cur\":{}", (word598 >> 4) & 0x7));
                    }
                }
                let wss = 154 + 30 * 8;
                if count > 30 {
                    let fc = u32_at(&word, wss)? as usize;
                    let lcb = u32_at(&word, wss + 4)? as usize;
                    if lcb == 36
                        && let Some(selsf) = table.get(fc..fc + 36)
                    {
                        out.push(format!("\"selsf_flags\":\"0x{:04X}\"", u16_at(selsf, 0)?));
                        out.push(format!("\"selsf_cp_first\":{}", u32_at(selsf, 4)?));
                        out.push(format!("\"selsf_cp_lim\":{}", u32_at(selsf, 8)?));
                        out.push(format!("\"selsf_cp_anchor\":{}", u32_at(selsf, 20)?));
                    } else {
                        out.push(format!("\"lcb_wss\":{lcb}"));
                    }
                }
            }
            Some(out)
        })();
        if let Some(shape) = shape {
            fields.extend(shape);
        }
        let classification = match Snapshot::from_bytes(bytes.clone()) {
            Ok(snapshot) => format!("{:?}", snapshot.protection()),
            Err(error) => format!("error: {error}"),
        };
        fields.push(format!("\"classification\":{}", json_str(&classification)));
        fields.push(format!("\"edit_cp0\":{}", json_str(&edit(&bytes, 0))));
        fields.push(format!("\"edit_cp1\":{}", json_str(&edit(&bytes, 1))));
        println!("{{{}}}", fields.join(","));
    }
}
