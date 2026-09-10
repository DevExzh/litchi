use litchi_xls::{FtCf, OleObjectRecord};

fn record(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(4 + body.len());
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(body.len() as u16).to_le_bytes());
    out.extend_from_slice(body);
    out
}

fn ordinary_body(with_extra_cf: bool) -> Vec<u8> {
    let mut body = record(
        0x0015,
        &[
            0x08, 0x00, 0x07, 0x00, 0x00, 0x00, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        ],
    );
    body.extend_from_slice(&record(0x0007, &FtCf::UNSPECIFIED.to_le_bytes()));
    if with_extra_cf {
        body.extend_from_slice(&record(0x0007, &[0x00]));
    }
    body.extend_from_slice(&record(0x0008, &[0x00, 0x00]));
    let formula = [
        0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00,
    ];
    let mut fmla = (formula.len() as u16).to_le_bytes().to_vec();
    fmla.extend_from_slice(&formula);
    fmla.extend_from_slice(&0x2Au32.to_le_bytes());
    body.extend_from_slice(&record(0x0009, &fmla));
    body.extend_from_slice(&record(0x0000, &[]));
    body
}

fn ordinary_body_with_reordered_picture_prefix() -> Vec<u8> {
    let mut body = record(
        0x0015,
        &[
            0x08, 0x00, 0x07, 0x00, 0x00, 0x00, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        ],
    );
    body.extend_from_slice(&record(0x0008, &[0x00, 0x00]));
    body.extend_from_slice(&record(0x0007, &FtCf::UNSPECIFIED.to_le_bytes()));
    let formula = [
        0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00,
    ];
    let mut fmla = (formula.len() as u16).to_le_bytes().to_vec();
    fmla.extend_from_slice(&formula);
    fmla.extend_from_slice(&0x2Au32.to_le_bytes());
    body.extend_from_slice(&record(0x0009, &fmla));
    body.extend_from_slice(&record(0x0000, &[]));
    body
}

fn ordinary_body_with_forbidden_checkbox_data() -> Vec<u8> {
    let mut body = ordinary_body(false);
    // FtCblsData is valid on its own, but MS-XLS 2.4.181 permits it only for
    // cmo.ot 0x000B/0x000C, never for an ordinary OLE picture (0x0008).
    let checkbox = record(0x0012, &[0, 0, 0, 0, 0, 0, 0, 0]);
    let end = body.len() - 4;
    body.splice(end..end, checkbox);
    body
}

#[test]
fn v6_probe_extra_malformed_ftcf_is_rejected() {
    assert!(OleObjectRecord::parse(&ordinary_body(false), None).is_ok());
    assert!(OleObjectRecord::parse(&ordinary_body(true), None).is_err());
}

#[test]
fn v6_probe_reordered_picture_prefix_is_rejected() {
    assert!(OleObjectRecord::parse(&ordinary_body_with_reordered_picture_prefix(), None).is_err());
}

#[test]
fn v7_probe_forbidden_control_subrecord_is_rejected() {
    assert!(OleObjectRecord::parse(&ordinary_body_with_forbidden_checkbox_data(), None).is_err());
}
