use litchi::ole::OleFile;
use litchi::ole::consts::PptRecordType;
use litchi::ole::ppt::PptRecord;

fn dump_children(rec: &PptRecord, indent: usize, depth: usize) {
    if depth == 0 || rec.children.is_empty() {
        return;
    }

    let pad = " ".repeat(indent);
    println!("{}children ({} records):", pad, rec.children.len());
    for (idx, child) in rec.children.iter().enumerate() {
        println!(
            "{}  [{}] type_raw=0x{:04X}, type={:?}, ver={}, inst={}, len={}",
            pad,
            idx,
            child.record_type_raw,
            child.record_type,
            child.version,
            child.instance,
            child.data_length,
        );

        if child.record_type == PptRecordType::PPDrawingGroup {
            dump_escher_dgg(&child.data);
        }
        if child.record_type == PptRecordType::PPDrawing {
            dump_escher_dg(&child.data);
        }

        // Recurse one level deeper for container-like children
        if !child.children.is_empty() {
            dump_children(child, indent + 4, depth - 1);
        }
    }
}

fn dump_persist_dir(data: &[u8]) {
    println!("  PersistPtr payload ({} bytes):", data.len());
    for (i, chunk) in data.chunks(4).take(8).enumerate() {
        if chunk.len() < 4 {
            break;
        }
        let v = u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
        if i == 0 {
            let base = v & 0x000F_FFFF;
            let count = (v >> 20) & 0x0FFF;
            println!(
                "    info[0]=0x{:08X} (base_id={}, count={})",
                v, base, count
            );
        } else {
            println!("    dword[{}]=0x{:08X}", i, v);
        }
    }
}

fn dump_user_edit(data: &[u8]) {
    if data.len() < 28 {
        println!("  UserEditAtom data too short: {} bytes", data.len());
        return;
    }
    let last_viewed = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let ppt_version = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
    let off_last_edit = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
    let off_persist = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
    let doc_ref = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
    let max_persist = u32::from_le_bytes([data[20], data[21], data[22], data[23]]);
    let last_view = u16::from_le_bytes([data[24], data[25]]);
    let unused = u16::from_le_bytes([data[26], data[27]]);
    println!("  UserEditAtom:");
    println!("    lastViewedSlideID={}", last_viewed);
    println!("    pptVersion=0x{:08X}", ppt_version);
    println!("    offsetLastEdit=0x{:08X}", off_last_edit);
    println!("    offsetPersistDir=0x{:08X}", off_persist);
    println!("    docPersistRef={}", doc_ref);
    println!("    maxPersistWritten={}", max_persist);
    println!("    lastViewType={}", last_view);
    println!("    unused=0x{:04X}", unused);
}

fn dump_current_user_stream(data: &[u8]) {
    println!("Current User stream length: {} bytes", data.len());
    if data.len() < 28 {
        println!("  CurrentUser stream too short to parse header");
        return;
    }

    let ver_inst = u16::from_le_bytes([data[0], data[1]]);
    let rec_type = u16::from_le_bytes([data[2], data[3]]);
    let atom_size = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
    let detail_size = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
    let header_token = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
    let cur_edit = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
    let username_len = u16::from_le_bytes([data[20], data[21]]) as usize;
    let doc_final_version = u16::from_le_bytes([data[22], data[23]]);
    let doc_major = data[24];
    let doc_minor = data[25];
    let reserved = u16::from_le_bytes([data[26], data[27]]);

    let ascii_start = 28usize;
    let ascii_end = ascii_start.saturating_add(username_len);
    let mut ascii_username = String::new();
    if username_len > 0 && ascii_end <= data.len() {
        ascii_username = String::from_utf8_lossy(&data[ascii_start..ascii_end]).to_string();
    }

    let mut release_version = 0u32;
    let mut unicode_username = String::new();
    let rel_pos = ascii_end;
    if rel_pos + 4 <= data.len() {
        release_version = u32::from_le_bytes([
            data[rel_pos],
            data[rel_pos + 1],
            data[rel_pos + 2],
            data[rel_pos + 3],
        ]);

        let uni_start = rel_pos + 4;
        let uni_len_bytes = username_len.saturating_mul(2);
        let uni_end = uni_start.saturating_add(uni_len_bytes);
        if uni_len_bytes > 0 && uni_end <= data.len() {
            let mut chars = Vec::with_capacity(username_len);
            let mut i = uni_start;
            while i + 1 < uni_end {
                let cu = u16::from_le_bytes([data[i], data[i + 1]]);
                if cu == 0 {
                    break;
                }
                if let Some(ch) = char::from_u32(cu as u32) {
                    chars.push(ch);
                }
                i += 2;
            }
            unicode_username = chars.into_iter().collect();
        }
    }

    println!("  CurrentUser atom:");
    println!(
        "    ver_inst=0x{:04X}, rec_type=0x{:04X}",
        ver_inst, rec_type
    );
    println!("    atomSize={} (bytes after header)", atom_size);
    println!("    detailSize={}", detail_size);
    println!("    headerToken=0x{:08X}", header_token);
    println!("    currentEditOffset=0x{:08X}", cur_edit);
    println!("    usernameLen={} (bytes, ASCII)", username_len);
    println!("    docFinalVersion=0x{:04X}", doc_final_version);
    println!("    docVersion={}.{}", doc_major, doc_minor);
    println!("    reserved=0x{:04X}", reserved);
    println!("    releaseVersion={}", release_version);
    if !ascii_username.is_empty() {
        println!("    asciiUsername={:?}", ascii_username);
    }
    if !unicode_username.is_empty() {
        println!("    unicodeUsername={:?}", unicode_username);
    }
}

fn dump_escher_dgg(data: &[u8]) {
    if data.len() < 8 {
        println!(
            "  [Escher] PPDrawingGroup payload too short for DggContainer: {} bytes",
            data.len()
        );
        return;
    }

    let ver_inst = u16::from_le_bytes([data[0], data[1]]);
    let r#type = u16::from_le_bytes([data[2], data[3]]);
    let len = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;
    println!(
        "  [Escher DggContainer] ver_inst=0x{:04X}, type=0x{:04X}, len={}",
        ver_inst, r#type, len
    );

    if r#type != 0xF000 {
        println!("    (unexpected Escher type for DggContainer, expected 0xF000)");
        return;
    }

    if data.len() < 8 + 8 {
        println!("    (too short for child Dgg record)");
        return;
    }

    let dgg_header_start = 8usize;
    let dgg_ver_inst = u16::from_le_bytes([data[dgg_header_start], data[dgg_header_start + 1]]);
    let dgg_type = u16::from_le_bytes([data[dgg_header_start + 2], data[dgg_header_start + 3]]);
    let dgg_len = u32::from_le_bytes([
        data[dgg_header_start + 4],
        data[dgg_header_start + 5],
        data[dgg_header_start + 6],
        data[dgg_header_start + 7],
    ]) as usize;
    println!(
        "    Dgg: ver_inst=0x{:04X}, type=0x{:04X}, len={}",
        dgg_ver_inst, dgg_type, dgg_len
    );

    if dgg_type != 0xF006 {
        println!("    (unexpected Dgg record type, expected 0xF006)");
        return;
    }

    let dgg_body_start = dgg_header_start + 8;
    let dgg_body_end = dgg_body_start.saturating_add(dgg_len);
    if data.len() < dgg_body_end || dgg_len < 16 {
        println!("    (Dgg body too short: len={})", dgg_len);
        return;
    }

    let shape_id_max = u32::from_le_bytes([
        data[dgg_body_start],
        data[dgg_body_start + 1],
        data[dgg_body_start + 2],
        data[dgg_body_start + 3],
    ]);
    let num_id_clusters_field = u32::from_le_bytes([
        data[dgg_body_start + 4],
        data[dgg_body_start + 5],
        data[dgg_body_start + 6],
        data[dgg_body_start + 7],
    ]);
    let num_shapes_saved = u32::from_le_bytes([
        data[dgg_body_start + 8],
        data[dgg_body_start + 9],
        data[dgg_body_start + 10],
        data[dgg_body_start + 11],
    ]);
    let drawings_saved = u32::from_le_bytes([
        data[dgg_body_start + 12],
        data[dgg_body_start + 13],
        data[dgg_body_start + 14],
        data[dgg_body_start + 15],
    ]);

    println!("    shapeIdMax={} (next shape id)", shape_id_max);
    println!(
        "    numIdClustersField={} (numIdClusters field)",
        num_id_clusters_field
    );
    println!("    numShapesSaved={}", num_shapes_saved);
    println!("    drawingsSaved={}", drawings_saved);

    let mut pos = dgg_body_start + 16;
    let mut idx = 0usize;
    while pos + 8 <= dgg_body_end {
        let dg_id = u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
        let num_shape_ids_used =
            u32::from_le_bytes([data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]]);
        println!(
            "    cluster[{}]: dgId={}, numShapeIdsUsed={}",
            idx, dg_id, num_shape_ids_used
        );
        idx += 1;
        pos += 8;
    }
}

fn dump_escher_dg(data: &[u8]) {
    if data.len() < 8 {
        println!(
            "  [Escher] PPDrawing payload too short for DgContainer: {} bytes",
            data.len()
        );
        return;
    }

    let ver_inst = u16::from_le_bytes([data[0], data[1]]);
    let r#type = u16::from_le_bytes([data[2], data[3]]);
    let len = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;
    println!(
        "  [Escher DgContainer] ver_inst=0x{:04X}, type=0x{:04X}, len={}",
        ver_inst, r#type, len
    );

    if r#type != 0xF002 {
        println!("    (unexpected Escher type for DgContainer, expected 0xF002)");
        return;
    }

    if data.len() < 8 + 8 {
        println!("    (too short for child Dg record)");
        return;
    }

    let dg_header_start = 8usize;
    let dg_type = u16::from_le_bytes([data[dg_header_start + 2], data[dg_header_start + 3]]);
    let dg_len = u32::from_le_bytes([
        data[dg_header_start + 4],
        data[dg_header_start + 5],
        data[dg_header_start + 6],
        data[dg_header_start + 7],
    ]) as usize;
    println!("    Dg: type=0x{:04X}, len={}", dg_type, dg_len);

    if dg_type != 0xF008 {
        println!("    (unexpected Dg record type, expected 0xF008)");
        return;
    }

    let dg_body_start = dg_header_start + 8;
    let dg_body_end = dg_body_start.saturating_add(dg_len);
    if data.len() < dg_body_end || dg_len < 8 {
        println!("    (Dg body too short: len={})", dg_len);
        return;
    }

    let csp = u32::from_le_bytes([
        data[dg_body_start],
        data[dg_body_start + 1],
        data[dg_body_start + 2],
        data[dg_body_start + 3],
    ]);
    let spid_cur = u32::from_le_bytes([
        data[dg_body_start + 4],
        data[dg_body_start + 5],
        data[dg_body_start + 6],
        data[dg_body_start + 7],
    ]);
    println!("    csp={} (shapes in drawing)", csp);
    println!("    spidCur={} (last shape id)", spid_cur);
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "minimal.ppt".to_string());

    println!("Analyzing PPT structure: {}", path);

    let file = std::fs::File::open(&path)?;
    let mut ole = OleFile::open(file)?;

    // List all OLE streams for high-level comparison
    let streams = ole.list_streams();
    println!("OLE streams ({}):", streams.len());
    for s in &streams {
        println!("  - {}", s.join("/"));
    }

    if let Ok(cur) = ole.open_stream(&["Current User"]) {
        dump_current_user_stream(&cur);
    } else {
        println!("Current User stream not found");
    }

    let data = if let Ok(d) = ole.open_stream(&["PowerPoint Document"]) {
        d
    } else {
        eprintln!("PowerPoint Document stream not found");
        return Ok(());
    };

    println!("PowerPoint Document stream length: {} bytes", data.len());

    let mut offset: usize = 0;
    let len = data.len();
    let mut index = 0usize;

    while offset + 8 <= len {
        match PptRecord::parse(&data, offset) {
            Ok((rec, consumed)) => {
                println!(
                    "[{}] @0x{:08X}: ver={}, inst={}, type_raw=0x{:04X}, type={:?}, len={}",
                    index,
                    offset,
                    rec.version,
                    rec.instance,
                    rec.record_type_raw,
                    rec.record_type,
                    rec.data_length,
                );

                if matches!(
                    rec.record_type,
                    PptRecordType::Document | PptRecordType::MainMaster | PptRecordType::Slide
                ) {
                    dump_children(&rec, 2, 3);
                }
                if rec.record_type == PptRecordType::PPDrawingGroup {
                    dump_escher_dgg(&rec.data);
                }
                if rec.record_type == PptRecordType::PPDrawing {
                    dump_escher_dg(&rec.data);
                }
                if rec.record_type_raw == 6001 || rec.record_type_raw == 6002 {
                    dump_persist_dir(&rec.data);
                }
                if rec.record_type_raw == 4085 {
                    dump_user_edit(&rec.data);
                }

                if consumed == 0 {
                    break;
                }
                offset += consumed;
                index += 1;
            },
            Err(e) => {
                println!("Parse error at 0x{:08X}: {}", offset, e);
                offset += 1;
            },
        }
    }

    Ok(())
}
