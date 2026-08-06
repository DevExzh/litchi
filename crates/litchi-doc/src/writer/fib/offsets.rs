//! Typed fc/lcb references from FibRgFcLcb.

use super::header::FibBuilder;

#[derive(Debug, Default, Clone, Copy)]
pub(super) struct Range {
    offset: u32,
    size: u32,
}

impl Range {
    const fn new(offset: u32, size: u32) -> Self {
        Self { offset, size }
    }
}

/// All table references emitted by the writer.
#[derive(Debug, Default)]
pub(super) struct Offsets {
    pub(super) stshf: Range,
    pub(super) dop: Range,
    pub(super) sttbf_assoc: Range,
    pub(super) sttbf_glsy: Range,
    pub(super) plcf_glsy: Range,
    pub(super) sttb_glsy_style: Range,
    pub(super) clx: Range,
    pub(super) plcfbte_chpx: Range,
    pub(super) plcfbte_papx: Range,
    pub(super) plcfsed: Range,
    pub(super) sttbfffn: Range,
    pub(super) plcfhdd: Range,
    pub(super) plcffnd_ref: Range,
    pub(super) plcffnd_txt: Range,
    pub(super) plcfend_ref: Range,
    pub(super) plcfend_txt: Range,
    pub(super) plcfand_ref: Range,
    pub(super) plcfand_txt: Range,
    pub(super) grp_xst_atn_owners: Range,
    pub(super) sttbf_bkmk: Range,
    pub(super) plcf_bkf: Range,
    pub(super) plcf_bkl: Range,
    pub(super) cmds: Range,
    pub(super) sttbf_bkmk_factoid: Range,
    pub(super) plcf_bkf_factoid: Range,
    pub(super) plcf_bkl_factoid: Range,
    pub(super) factoid_data: Range,
    pub(super) plcf_factoid: Range,
    pub(super) sttbf_rmark: Range,
    pub(super) plcfspl: Range,
    pub(super) plcfgram: Range,
    pub(super) sttb_saved_by: Range,
    pub(super) sttbf_atn_bkmk: Range,
    pub(super) plcf_atn_bkf: Range,
    pub(super) plcf_atn_bkl: Range,
    pub(super) atrd_extra: Range,
    pub(super) plflst: Range,
    pub(super) plflfo: Range,
    pub(super) sttb_list_names: Range,
    pub(super) sttb_rgtplc: Range,
    pub(super) plcffld_mom: Range,
    pub(super) plcffld_hdr: Range,
    pub(super) plc_spa_mom: Range,
    pub(super) plcftxbx_txt: Range,
    pub(super) plcf_hdr_txbx_txt: Range,
    pub(super) plc_spa_hdr: Range,
    pub(super) dgg_info: Range,
}

macro_rules! define_setter {
    ($name:ident, $field:ident) => {
        pub fn $name(&mut self, offset: u32, size: u32) {
            self.offsets.$field = Range::new(offset, size);
        }
    };
}

impl FibBuilder {
    define_setter!(set_stshf, stshf);
    define_setter!(set_dop, dop);
    define_setter!(set_sttbf_assoc, sttbf_assoc);
    define_setter!(set_sttbf_glsy, sttbf_glsy);
    define_setter!(set_plcf_glsy, plcf_glsy);
    define_setter!(set_sttb_glsy_style, sttb_glsy_style);
    define_setter!(set_clx, clx);
    define_setter!(set_plcfbte_chpx, plcfbte_chpx);
    define_setter!(set_plcfbte_papx, plcfbte_papx);
    define_setter!(set_plcfsed, plcfsed);
    define_setter!(set_sttbfffn, sttbfffn);
    define_setter!(set_plcfhdd, plcfhdd);
    define_setter!(set_plcffnd_ref, plcffnd_ref);
    define_setter!(set_plcffnd_txt, plcffnd_txt);
    define_setter!(set_plcfend_ref, plcfend_ref);
    define_setter!(set_plcfend_txt, plcfend_txt);
    define_setter!(set_plcfand_ref, plcfand_ref);
    define_setter!(set_plcfand_txt, plcfand_txt);
    define_setter!(set_grp_xst_atn_owners, grp_xst_atn_owners);
    define_setter!(set_sttbf_bkmk, sttbf_bkmk);
    define_setter!(set_plcf_bkf, plcf_bkf);
    define_setter!(set_plcf_bkl, plcf_bkl);
    define_setter!(set_cmds, cmds);
    define_setter!(set_sttbf_bkmk_factoid, sttbf_bkmk_factoid);
    define_setter!(set_plcf_bkf_factoid, plcf_bkf_factoid);
    define_setter!(set_plcf_bkl_factoid, plcf_bkl_factoid);
    define_setter!(set_factoid_data, factoid_data);
    define_setter!(set_plcf_factoid, plcf_factoid);
    define_setter!(set_sttbf_rmark, sttbf_rmark);
    define_setter!(set_plcfspl, plcfspl);
    define_setter!(set_plcfgram, plcfgram);
    define_setter!(set_sttb_saved_by, sttb_saved_by);
    define_setter!(set_sttbf_atn_bkmk, sttbf_atn_bkmk);
    define_setter!(set_plcf_atn_bkf, plcf_atn_bkf);
    define_setter!(set_plcf_atn_bkl, plcf_atn_bkl);
    define_setter!(set_atrd_extra, atrd_extra);
    define_setter!(set_plflst, plflst);
    define_setter!(set_plflfo, plflfo);
    define_setter!(set_sttb_list_names, sttb_list_names);
    define_setter!(set_sttb_rgtplc, sttb_rgtplc);
    define_setter!(set_plcffld_mom, plcffld_mom);
    define_setter!(set_plcffld_hdr, plcffld_hdr);
    define_setter!(set_plc_spa_mom, plc_spa_mom);
    define_setter!(set_plcftxbx_txt, plcftxbx_txt);
    define_setter!(set_plcf_hdr_txbx_txt, plcf_hdr_txbx_txt);
    define_setter!(set_plc_spa_hdr, plc_spa_hdr);
    define_setter!(set_dgg_info, dgg_info);
}

impl Offsets {
    /// Encode the 136 fixed-width fc/lcb pairs.
    pub(super) fn write_into(&self, buf: &mut [u8]) {
        buf.fill(0);

        let set = |buf: &mut [u8], index: usize, range: Range| {
            let offset = index * 8;
            if offset + 8 <= buf.len() {
                buf[offset..offset + 4].copy_from_slice(&range.offset.to_le_bytes());
                buf[offset + 4..offset + 8].copy_from_slice(&range.size.to_le_bytes());
            }
        };

        // Field indices are the FibRgFcLcb order from [MS-DOC].
        set(buf, 1, self.stshf);
        set(buf, 9, self.sttbf_glsy);
        set(buf, 10, self.plcf_glsy);
        set(buf, 83, self.sttb_glsy_style);
        set(buf, 32, self.sttbf_assoc);
        set(buf, 2, self.plcffnd_ref);
        set(buf, 3, self.plcffnd_txt);
        set(buf, 4, self.plcfand_ref);
        set(buf, 5, self.plcfand_txt);
        set(buf, 21, self.sttbf_bkmk);
        set(buf, 22, self.plcf_bkf);
        set(buf, 23, self.plcf_bkl);
        set(buf, 24, self.cmds);
        set(buf, 46, self.plcfend_ref);
        set(buf, 47, self.plcfend_txt);
        set(buf, 51, self.sttbf_rmark);
        set(buf, 55, self.plcfspl);
        set(buf, 90, self.plcfgram);
        set(buf, 71, self.sttb_saved_by);
        set(buf, 6, self.plcfsed);
        set(buf, 11, self.plcfhdd);
        set(buf, 12, self.plcfbte_chpx);
        set(buf, 37, self.sttbf_atn_bkmk);
        set(buf, 42, self.plcf_atn_bkf);
        set(buf, 43, self.plcf_atn_bkl);
        set(buf, 112, self.atrd_extra);
        set(buf, 114, self.sttbf_bkmk_factoid);
        set(buf, 115, self.plcf_bkf_factoid);
        set(buf, 117, self.plcf_bkl_factoid);
        set(buf, 118, self.factoid_data);
        set(buf, 132, self.plcf_factoid);
        set(buf, 13, self.plcfbte_papx);
        set(buf, 15, self.sttbfffn);
        set(buf, 16, self.plcffld_mom);
        set(buf, 17, self.plcffld_hdr);
        set(buf, 36, self.grp_xst_atn_owners);
        set(buf, 73, self.plflst);
        set(buf, 74, self.plflfo);
        set(buf, 91, self.sttb_list_names);
        set(buf, 96, self.sttb_rgtplc);
        set(buf, 31, self.dop);
        set(buf, 33, self.clx);
        set(buf, 40, self.plc_spa_mom);
        set(buf, 41, self.plc_spa_hdr);
        set(buf, 56, self.plcftxbx_txt);
        set(buf, 58, self.plcf_hdr_txbx_txt);
        set(buf, 50, self.dgg_info);
    }
}
