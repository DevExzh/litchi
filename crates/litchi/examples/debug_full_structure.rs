use litchi_cfb::OleFile;
use std::env;
use std::fs::File;
use std::io::BufReader;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Accept path argument; default to headers_footers.doc
    let path = env::args()
        .nth(1)
        .unwrap_or_else(|| "headers_footers.doc".to_string());
    let file = File::open(&path)?;
    let reader = BufReader::new(file);
    let mut ole = OleFile::open(reader)?;

    let table_data = ole.open_stream(&["1Table"])?;
    let wd_data = ole.open_stream(&["WordDocument"])?;

    println!("=== COMPLETE STRUCTURE ANALYSIS: {} ===\n", path);

    // 1) FIB base fields
    println!("1) FIB BASE:");
    let magic = u16::from_le_bytes([wd_data[0], wd_data[1]]);
    let n_fib = u16::from_le_bytes([wd_data[2], wd_data[3]]);
    let fc_min = u32::from_le_bytes([wd_data[24], wd_data[25], wd_data[26], wd_data[27]]);
    let fc_mac = u32::from_le_bytes([wd_data[28], wd_data[29], wd_data[30], wd_data[31]]);
    println!("   magic=0x{:04X} nFib=0x{:04X}", magic, n_fib);
    println!("   fcMin={} fcMac={}", fc_min, fc_mac);

    // FibRgLw (starts at 64)
    println!("\n2) FibRgLw COUNTS:");
    let cb_mac = u32::from_le_bytes([wd_data[64], wd_data[65], wd_data[66], wd_data[67]]);
    let ccp_text = u32::from_le_bytes([wd_data[76], wd_data[77], wd_data[78], wd_data[79]]);
    let ccp_ftn = u32::from_le_bytes([wd_data[80], wd_data[81], wd_data[82], wd_data[83]]);
    let ccp_hdd = u32::from_le_bytes([wd_data[84], wd_data[85], wd_data[86], wd_data[87]]);
    println!("   cbMac={} (total WordDocument bytes)", cb_mac);
    println!(
        "   ccpText={} ccpFtn={} ccpHdd={}",
        ccp_text, ccp_ftn, ccp_hdd
    );

    // FibRgFcLcb97 pairs begin at offset 154 in our FIB
    let fclcb_base = 154usize;
    let read_pair = |idx: usize| -> (u32, u32) {
        let off = fclcb_base + idx * 8;
        let fc = u32::from_le_bytes([
            wd_data[off],
            wd_data[off + 1],
            wd_data[off + 2],
            wd_data[off + 3],
        ]);
        let lcb = u32::from_le_bytes([
            wd_data[off + 4],
            wd_data[off + 5],
            wd_data[off + 6],
            wd_data[off + 7],
        ]);
        (fc, lcb)
    };

    // Indices must match what our FibBuilder writes
    let (fc_stshf, lcb_stshf) = read_pair(1); // StyleSheet
    let (fc_plcfsed, lcb_plcfsed) = read_pair(6); // Section table
    let (fc_plcfhdd, lcb_plcfhdd) = read_pair(11); // PlcfHdd
    let (fc_plcfbte_chpx, lcb_plcfbte_chpx) = read_pair(12);
    let (fc_plcfbte_papx, lcb_plcfbte_papx) = read_pair(13);
    let (fc_sttbfffn, lcb_sttbfffn) = read_pair(15); // Font table
    let (fc_dop, lcb_dop) = read_pair(31);
    let (fc_clx, lcb_clx) = read_pair(33);

    println!("\n3) FibRgFcLcb97 POINTERS (1Table):");
    println!("   STSHF   @{} (lcb={})", fc_stshf, lcb_stshf);
    println!("   PLCFSED @{} (lcb={})", fc_plcfsed, lcb_plcfsed);
    println!("   PLCFHDD @{} (lcb={})", fc_plcfhdd, lcb_plcfhdd);
    println!("   BTE CHP @{} (lcb={})", fc_plcfbte_chpx, lcb_plcfbte_chpx);
    println!("   BTE PAP @{} (lcb={})", fc_plcfbte_papx, lcb_plcfbte_papx);
    println!("   STTBFFFN@{} (lcb={})", fc_sttbfffn, lcb_sttbfffn);
    println!("   DOP     @{} (lcb={})", fc_dop, lcb_dop);
    println!("   CLX     @{} (lcb={})", fc_clx, lcb_clx);

    // 4) CLX (Piece Table)
    println!("\n4) CLX (Piece Table):");
    if lcb_clx >= 5 && (fc_clx as usize + 5) <= table_data.len() {
        let clx_type = table_data[fc_clx as usize];
        let clx_size = u32::from_le_bytes([
            table_data[fc_clx as usize + 1],
            table_data[fc_clx as usize + 2],
            table_data[fc_clx as usize + 3],
            table_data[fc_clx as usize + 4],
        ]);
        println!("   type=0x{:02X} size={} bytes", clx_type, clx_size);
        let plcpcd_start = fc_clx as usize + 5;
        let n = ((clx_size - 4) / 12) as usize; // size = 12n + 4
        println!("   pieces={} (derived)", n);
        if n > 0 {
            let cp0 = u32::from_le_bytes([
                table_data[plcpcd_start],
                table_data[plcpcd_start + 1],
                table_data[plcpcd_start + 2],
                table_data[plcpcd_start + 3],
            ]);
            let last_cp_off = plcpcd_start + n * 4;
            let cp_n = u32::from_le_bytes([
                table_data[last_cp_off],
                table_data[last_cp_off + 1],
                table_data[last_cp_off + 2],
                table_data[last_cp_off + 3],
            ]);
            println!("   CP[0]={} .. CP[n]={} (len={})", cp0, cp_n, cp_n - cp0);

            // Parse PCDs
            let pcd_start = plcpcd_start + (n + 1) * 4;
            println!("   -- PCD list --");
            for i in 0..n {
                let cps = u32::from_le_bytes([
                    table_data[plcpcd_start + i * 4],
                    table_data[plcpcd_start + i * 4 + 1],
                    table_data[plcpcd_start + i * 4 + 2],
                    table_data[plcpcd_start + i * 4 + 3],
                ]);
                let cpe = u32::from_le_bytes([
                    table_data[plcpcd_start + (i + 1) * 4],
                    table_data[plcpcd_start + (i + 1) * 4 + 1],
                    table_data[plcpcd_start + (i + 1) * 4 + 2],
                    table_data[plcpcd_start + (i + 1) * 4 + 3],
                ]);
                let p_off = pcd_start + i * 8;
                if p_off + 8 > table_data.len() {
                    break;
                }
                let fc_enc = u32::from_le_bytes([
                    table_data[p_off + 2],
                    table_data[p_off + 3],
                    table_data[p_off + 4],
                    table_data[p_off + 5],
                ]);
                let is_ansi = (fc_enc & 0x4000_0000) != 0;
                let fc = if is_ansi {
                    (fc_enc & 0x3FFF_FFFF) / 2
                } else {
                    fc_enc
                };
                println!(
                    "     #{:02}: CP[{}..{}) len={} -> FC={} type={}",
                    i,
                    cps,
                    cpe,
                    cpe.saturating_sub(cps),
                    fc,
                    if is_ansi { "ANSI" } else { "Unicode" }
                );
            }
        }
    } else {
        println!("   CLX not present or invalid pointers");
    }

    // 5) PLCFBTE (CHPX/PAPX) check - single entry expected, 12 bytes each
    println!("\n5) PLCFBTE (Bin Tables):");
    if lcb_plcfbte_chpx == 12 && (fc_plcfbte_chpx as usize + 12) <= table_data.len() {
        let s = fc_plcfbte_chpx as usize;
        let start_fc = u32::from_le_bytes([
            table_data[s],
            table_data[s + 1],
            table_data[s + 2],
            table_data[s + 3],
        ]);
        let end_fc = u32::from_le_bytes([
            table_data[s + 4],
            table_data[s + 5],
            table_data[s + 6],
            table_data[s + 7],
        ]);
        let pn = u32::from_le_bytes([
            table_data[s + 8],
            table_data[s + 9],
            table_data[s + 10],
            table_data[s + 11],
        ]);
        println!("   CHPX: [{}, {}) -> page {}", start_fc, end_fc, pn);
    } else {
        println!("   CHPX: unexpected size (lcb={})", lcb_plcfbte_chpx);
    }
    if lcb_plcfbte_papx == 12 && (fc_plcfbte_papx as usize + 12) <= table_data.len() {
        let s = fc_plcfbte_papx as usize;
        let start_fc = u32::from_le_bytes([
            table_data[s],
            table_data[s + 1],
            table_data[s + 2],
            table_data[s + 3],
        ]);
        let end_fc = u32::from_le_bytes([
            table_data[s + 4],
            table_data[s + 5],
            table_data[s + 6],
            table_data[s + 7],
        ]);
        let pn = u32::from_le_bytes([
            table_data[s + 8],
            table_data[s + 9],
            table_data[s + 10],
            table_data[s + 11],
        ]);
        println!("   PAPX: [{}, {}) -> page {}", start_fc, end_fc, pn);
    } else {
        println!("   PAPX: unexpected size (lcb={})", lcb_plcfbte_papx);
    }

    // 6) PlcfHdd (header/footer PLCF of CPs)
    println!("\n6) PLCFHDD (Header/Footer PLCF):");
    if lcb_plcfhdd > 0 && (fc_plcfhdd as usize + lcb_plcfhdd as usize) <= table_data.len() {
        let count = (lcb_plcfhdd / 4) as usize; // CPs only (no data elements)
        let mut cps: Vec<u32> = Vec::with_capacity(count);
        for i in 0..count {
            let o = fc_plcfhdd as usize + i * 4;
            cps.push(u32::from_le_bytes([
                table_data[o],
                table_data[o + 1],
                table_data[o + 2],
                table_data[o + 3],
            ]));
        }
        println!("   entries={} (expected 13 CPs for 12 slots)", count);
        // Map indices 6..11 to names
        let labels = [
            "(0) FtnSep",
            "(1) FtnCont",
            "(2) FtnContNote",
            "(3) EndSep",
            "(4) EndCont",
            "(5) EndContNote",
            "(6) EvenHdr",
            "(7) OddHdr",
            "(8) EvenFtr",
            "(9) OddFtr",
            "(10) FirstHdr",
            "(11) FirstFtr",
        ];
        for i in 0..12.min(cps.len().saturating_sub(1)) {
            let start = cps[i];
            let end = cps[i + 1];
            if i >= 6 {
                println!(
                    "   {:<12}: CP[{}..{}) len={}{}",
                    labels[i],
                    i,
                    i + 1,
                    end.saturating_sub(start),
                    if end > start { "" } else { " (empty)" }
                );
            }
        }
        if let Some(last) = cps.last() {
            println!("   ccpHdd (from PLCF end) = {}", last);
        }
        println!("   ccpHdd (from FIB)     = {}", ccp_hdd);

        // Map header CP to global CP/FC and preview text for 6..11
        let hdd_base_cp = ccp_text + ccp_ftn; // Header story starts after main and footnotes
        println!("\n   -- Map HDD CPs to global FC --");
        // Re-parse PCDs for mapping
        if lcb_clx >= 5 && (fc_clx as usize + 5) <= table_data.len() {
            let clx_size = u32::from_le_bytes([
                table_data[fc_clx as usize + 1],
                table_data[fc_clx as usize + 2],
                table_data[fc_clx as usize + 3],
                table_data[fc_clx as usize + 4],
            ]) as usize;
            let plcpcd_start = fc_clx as usize + 5;
            let n = (clx_size - 4) / 12;
            let pcd_start = plcpcd_start + (n + 1) * 4;
            // Helper to find FC for a global CP
            let find_fc = |gcp: u32| -> Option<(u32, bool)> {
                for i in 0..n {
                    let cps = u32::from_le_bytes([
                        table_data[plcpcd_start + i * 4],
                        table_data[plcpcd_start + i * 4 + 1],
                        table_data[plcpcd_start + i * 4 + 2],
                        table_data[plcpcd_start + i * 4 + 3],
                    ]);
                    let cpe = u32::from_le_bytes([
                        table_data[plcpcd_start + (i + 1) * 4],
                        table_data[plcpcd_start + (i + 1) * 4 + 1],
                        table_data[plcpcd_start + (i + 1) * 4 + 2],
                        table_data[plcpcd_start + (i + 1) * 4 + 3],
                    ]);
                    if gcp >= cps && gcp < cpe {
                        let off = pcd_start + i * 8;
                        let fc_enc = u32::from_le_bytes([
                            table_data[off + 2],
                            table_data[off + 3],
                            table_data[off + 4],
                            table_data[off + 5],
                        ]);
                        let is_ansi = (fc_enc & 0x4000_0000) != 0;
                        let base_fc = if is_ansi {
                            (fc_enc & 0x3FFF_FFFF) / 2
                        } else {
                            fc_enc
                        };
                        let delta = gcp - cps;
                        let fc = base_fc + if is_ansi { delta } else { delta * 2 };
                        return Some((fc, is_ansi));
                    }
                }
                None
            };

            for i in 6..12.min(cps.len().saturating_sub(1)) {
                let local_s = cps[i];
                let local_e = cps[i + 1];
                if local_e > local_s {
                    let gcp_s = hdd_base_cp + local_s;
                    if let Some((fc, is_ansi)) = find_fc(gcp_s) {
                        let bytes = if is_ansi {
                            (local_e - local_s) as usize
                        } else {
                            (local_e - local_s) as usize * 2
                        };
                        let end = (fc as usize).saturating_add(bytes).min(wd_data.len());
                        let start = fc as usize;
                        let preview_len = end.saturating_sub(start).min(64);
                        let slice = &wd_data[start..start + preview_len];
                        print!(
                            "   {} -> gCP={} FC={} bytes={} preview=",
                            labels[i], gcp_s, fc, bytes
                        );
                        for b in slice {
                            print!("{:02X} ", b);
                        }
                        println!();
                    }
                }
            }
        }
    } else {
        println!("   Not present (lcb={})", lcb_plcfhdd);
    }

    // 7) Summary checks
    println!("\n7) SUMMARY:");
    println!("   ✓ FIB magic 0xA5EC, nFib 0x{:04X}", n_fib);
    println!("   ✓ fcMin={} fcMac={} cbMac={}", fc_min, fc_mac, cb_mac);
    if lcb_clx != 0 {
        println!("   ✓ CLX present ({} bytes)", lcb_clx);
    }
    if lcb_plcfhdd != 0 {
        println!("   ✓ PlcfHdd present ({} bytes)", lcb_plcfhdd);
    }

    Ok(())
}
