#!/usr/bin/env python3
"""Track C survey: CFB FAT geometry and BIFF globals read schedules.

Reads every file under ``test-data/ole`` (xls, doc, ppt) and, as an appendix,
every ``*.xls`` fixture elsewhere under ``test-data``.  Produces ``survey.json``
and ``survey.md`` next to this script (or under ``--out``).

Everything reported about reads is a MODEL of stream-level (logical) reads and
of the physically contiguous spans those reads would split into.  Nothing here
is a measured syscall count.

Deterministic; stdlib + ``olefile`` only.
"""

import argparse
import collections
import json
import os
import statistics
import struct
import sys

import olefile

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
MAXREGSECT = 0xFFFFFFFA
HEADER_DIFAT_ENTRIES = 109
HEADER_DIFAT_OFFSET = 0x4C
MINI_SECTOR_SIZE = 64

BOF = 0x0809
OLD_BOFS = {0x0009: "BIFF2", 0x0209: "BIFF3", 0x0409: "BIFF4"}
EOF = 0x000A
FILEPASS = 0x002F
BOUND_SHEET = 0x0085
BOUNDSHEET8 = 0x0085

WINDOW_FIRST = 512
EXACT_PROLOGUE_RECORDS = 4
MAX_GLOBAL_BYTES = 128 * 1024 * 1024
WINDOW_MAX = 65536

RECORD_NAMES = {
    0x0006: "Formula", 0x000A: "EOF", 0x000C: "CalcCount", 0x000D: "CalcMode",
    0x000E: "CalcPrecision", 0x000F: "CalcRefMode", 0x0010: "CalcDelta",
    0x0011: "CalcIter", 0x0012: "Protect", 0x0013: "Password", 0x0014: "Header",
    0x0015: "Footer", 0x0017: "ExternSheet", 0x0018: "Lbl", 0x0019: "WinProtect",
    0x001A: "VerticalPageBreaks", 0x001B: "HorizontalPageBreaks", 0x001C: "Note",
    0x001D: "Selection", 0x0022: "Date1904", 0x0023: "ExternName",
    0x0026: "LeftMargin", 0x0027: "RightMargin", 0x0028: "TopMargin",
    0x0029: "BottomMargin", 0x002A: "PrintRowCol", 0x002B: "PrintGrid",
    0x002F: "FilePass", 0x0031: "Font", 0x0033: "PrintSize", 0x003C: "Continue",
    0x003D: "Window1", 0x0040: "Backup", 0x0041: "Pane", 0x0042: "CodePage",
    0x004D: "Pls", 0x0050: "DCon", 0x0051: "DConRef", 0x0052: "DConName",
    0x0055: "DefColWidth", 0x0059: "XCT", 0x005A: "CRN", 0x005B: "FileSharing",
    0x005C: "WriteAccess", 0x005D: "Obj", 0x005E: "Uncalced",
    0x005F: "CalcSaveRecalc", 0x0060: "Template", 0x0061: "Intl",
    0x0063: "ObjProtect", 0x007D: "ColInfo", 0x0080: "Guts", 0x0081: "WsBool",
    0x0082: "GridSet", 0x0083: "HCenter", 0x0084: "VCenter",
    0x0085: "BoundSheet8", 0x0086: "WriteProtect", 0x008C: "Country",
    0x008D: "HideObj", 0x0090: "Sort", 0x0092: "Palette", 0x009B: "FilterMode",
    0x009C: "BuiltInFnGroupCount", 0x009D: "AutoFilterInfo", 0x009E: "AutoFilter",
    0x00A0: "Scl", 0x00A1: "Setup", 0x00AE: "ScenMan", 0x00AF: "Scenario",
    0x00B0: "SxView", 0x00B1: "Sxvd", 0x00B2: "SXVI", 0x00B4: "SxIvd",
    0x00B5: "SXLI", 0x00B6: "SXPI", 0x00B8: "DocRoute", 0x00B9: "RecipName",
    0x00BD: "MulRk", 0x00BE: "MulBlank", 0x00C1: "Mms", 0x00C5: "SXDI",
    0x00C6: "SXDB", 0x00C7: "SXFDB", 0x00C8: "SXDBB", 0x00C9: "SXNum",
    0x00CA: "SxBool", 0x00CB: "SxErr", 0x00CC: "SXInt", 0x00CD: "SXString",
    0x00CE: "SXDtr", 0x00CF: "SxNil", 0x00D0: "SXTbl", 0x00D1: "SXTBRGIITM",
    0x00D2: "SxTbpg", 0x00D3: "ObProj", 0x00D5: "SXStreamID", 0x00D7: "DBCell",
    0x00D8: "SXRng", 0x00D9: "SxIsxoper", 0x00DA: "BookBool",
    0x00DC: "DbOrParamQry", 0x00DD: "ScenarioProtect", 0x00DE: "OleObjectSize",
    0x00E0: "XF", 0x00E1: "InterfaceHdr", 0x00E2: "InterfaceEnd", 0x00E3: "SXVS",
    0x00E5: "MergeCells", 0x00E9: "BkHim", 0x00EB: "MsoDrawingGroup",
    0x00EC: "MsoDrawing", 0x00ED: "MsoDrawingSelection", 0x00EF: "PhoneticInfo",
    0x00F0: "SxRule", 0x00F1: "SXEx", 0x00F2: "SxFilt", 0x00F4: "SxDXF",
    0x00F5: "SxItm", 0x00F6: "SxName", 0x00F7: "SxSelect", 0x00F8: "SXPair",
    0x00F9: "SxFmla", 0x00FB: "SxFormat", 0x00FC: "SST", 0x00FD: "LabelSst",
    0x00FF: "ExtSST", 0x0100: "SXVDEx", 0x0103: "SXFormula", 0x0122: "SXDBEx",
    0x0137: "RRDInsDel", 0x0138: "RRDHead", 0x013B: "RRDChgCell", 0x013D: "TabId",
    0x013E: "RRDRenSheet", 0x013F: "RRSort", 0x0140: "RRDMove",
    0x014A: "RRFormat", 0x014B: "RRAutoFmt", 0x014D: "RRInsertSh",
    0x014E: "RRDMoveBegin", 0x014F: "RRDMoveEnd", 0x0150: "RRDInsDelBegin",
    0x0151: "RRDInsDelEnd", 0x0152: "RRDConflict", 0x0153: "RRDDefName",
    0x0154: "RRDRstEtxp", 0x015F: "LRng", 0x0160: "UsesELFs", 0x0161: "DSF",
    0x0191: "CUsr", 0x0192: "CbUsr", 0x0193: "UsrInfo", 0x0194: "UsrExcl",
    0x0195: "FileLock", 0x0196: "RRDInfo", 0x0197: "BCUsrs", 0x0198: "UsrChk",
    0x01A9: "UserBView", 0x01AA: "UserSViewBegin", 0x01AB: "UserSViewEnd",
    0x01AC: "RRDUserView", 0x01AD: "Qsi", 0x01AE: "SupBook", 0x01AF: "Prot4Rev",
    0x01B0: "CondFmt", 0x01B1: "CF", 0x01B2: "DVal", 0x01B5: "DConBin",
    0x01B6: "TxO", 0x01B7: "RefreshAll", 0x01B8: "HLink", 0x01B9: "Lel",
    0x01BA: "CodeName", 0x01BB: "SXFDBType", 0x01BC: "Prot4RevPass",
    0x01BD: "ObNoMacros", 0x01BE: "Dv", 0x01C0: "Excel9File", 0x01C1: "RecalcId",
    0x01C2: "EntExU2", 0x0200: "Dimensions", 0x0201: "Blank", 0x0203: "Number",
    0x0204: "Label", 0x0205: "BoolErr", 0x0207: "String", 0x0208: "Row",
    0x020B: "Index", 0x0221: "Array", 0x0225: "DefaultRowHeight", 0x0236: "Table",
    0x023E: "Window2", 0x027E: "RK", 0x0293: "Style", 0x0418: "BigName",
    0x041E: "Format", 0x043C: "ContinueBigName", 0x04BC: "ShrFmla",
    0x0800: "HLinkTooltip", 0x0801: "WebPub", 0x0802: "QsiSXTag",
    0x0803: "DBQueryExt", 0x0804: "ExtString", 0x0805: "TxtQry", 0x0806: "Qsir",
    0x0807: "Qsif", 0x0808: "RRDTQSIF", 0x0809: "BOF", 0x080A: "OleDbConn",
    0x080B: "WOpt", 0x080C: "SXViewEx", 0x080D: "SXTH", 0x080E: "SXPIEx",
    0x080F: "SXVDTEx", 0x0810: "SXViewEx9", 0x0812: "ContinueFrt",
    0x0813: "RealTimeData", 0x0850: "ChartFrtInfo", 0x0851: "FrtWrapper",
    0x0852: "StartBlock", 0x0853: "EndBlock", 0x0854: "StartObject",
    0x0855: "EndObject", 0x0856: "CatLab", 0x0857: "YMult", 0x0858: "SXViewLink",
    0x0859: "PivotChartBits", 0x085A: "FrtFontList", 0x0862: "SheetExt",
    0x0863: "BookExt", 0x0864: "SXAddl", 0x0865: "CrErr", 0x0866: "HFPicture",
    0x0867: "FeatHdr", 0x0868: "Feat", 0x086A: "DataLabExt",
    0x086B: "DataLabExtContents", 0x086C: "CellWatch", 0x0871: "FeatHdr11",
    0x0872: "Feature11", 0x0874: "DropDownObjIds", 0x0875: "ContinueFrt11",
    0x0876: "DConn", 0x0877: "List12", 0x0878: "Feature12", 0x0879: "CondFmt12",
    0x087A: "CF12", 0x087B: "CFEx", 0x087C: "XFCRC", 0x087D: "XFExt",
    0x087E: "AutoFilter12", 0x087F: "ContinueFrt12", 0x0884: "MDTInfo",
    0x0885: "MDXStr", 0x0886: "MDXTuple", 0x0887: "MDXSet", 0x0888: "MDXProp",
    0x0889: "MDXKPI", 0x088A: "MDB", 0x088B: "PLV", 0x088C: "Compat12",
    0x088D: "DXF", 0x088E: "TableStyles", 0x088F: "TableStyle",
    0x0890: "TableStyleElement", 0x0892: "StyleExt", 0x0893: "NamePublish",
    0x0894: "NameCmt", 0x0895: "SortData", 0x0896: "Theme", 0x0897: "GUIDTypeLib",
    0x0898: "FnGrp12", 0x0899: "NameFnGrp12", 0x089A: "MTRSettings",
    0x089B: "CompressPictures", 0x089C: "HeaderFooter", 0x089D: "CrtLayout12",
    0x089E: "CrtMlFrt", 0x089F: "CrtMlFrtContinue", 0x08A3: "ForceFullCalculation",
    0x08A4: "ShapePropsStream", 0x08A5: "TextPropsStream", 0x08A6: "RichTextStream",
    0x08A7: "CrtLayout12A", 0x1001: "Units", 0x1002: "Chart", 0x1003: "Series",
    0x1006: "DataFormat", 0x1007: "LineFormat", 0x1009: "MarkerFormat",
    0x100A: "AreaFormat", 0x100B: "PieFormat", 0x100C: "AttachedLabel",
    0x100D: "SeriesText", 0x1014: "ChartFormat", 0x1015: "Legend",
    0x1016: "SeriesList", 0x1017: "Bar", 0x1018: "Line", 0x1019: "Pie",
    0x101A: "Area", 0x101B: "Scatter", 0x101C: "CrtLine", 0x101D: "Axis",
    0x101E: "Tick", 0x101F: "ValueRange", 0x1020: "CatSerRange",
    0x1021: "AxisLine", 0x1022: "CrtLink", 0x1024: "DefaultText", 0x1025: "Text",
    0x1026: "FontX", 0x1027: "ObjectLink", 0x1032: "Frame", 0x1033: "Begin",
    0x1034: "End", 0x1035: "PlotArea", 0x103A: "Chart3d", 0x103C: "PicF",
    0x103D: "DropBar", 0x103E: "Radar", 0x103F: "Surf", 0x1040: "RadarArea",
    0x1041: "AxisParent", 0x1043: "LegendException", 0x1044: "ShtProps",
    0x1045: "SerToCrt", 0x1046: "AxesUsed", 0x1048: "SBaseRef", 0x104A: "SerParent",
    0x104B: "SerAuxTrend", 0x104E: "IFmtRecord", 0x104F: "Pos", 0x1050: "AlRuns",
    0x1051: "BRAI", 0x105B: "SerAuxErrBar", 0x105C: "ClrtClient",
    0x105D: "SerFmt", 0x105F: "Chart3DBarShape", 0x1060: "Fbi", 0x1061: "BopPop",
    0x1062: "AxcExt", 0x1063: "Dat", 0x1064: "PlotGrowth", 0x1065: "SIIndex",
    0x1066: "GelFrame", 0x1067: "BopPopCustom", 0x1068: "Fbi2",
}


def record_name(kind):
    name = RECORD_NAMES.get(kind)
    return "0x%04X %s" % (kind, name) if name else "0x%04X" % kind


def u16(buf, off):
    return struct.unpack_from("<H", buf, off)[0]


def u32(buf, off):
    return struct.unpack_from("<I", buf, off)[0]


def runs_of(sectors):
    """Number of maximal runs of physically consecutive sector ids, in list order."""
    if not sectors:
        return 0
    runs = 1
    for prev, cur in zip(sectors, sectors[1:]):
        if cur != prev + 1:
            runs += 1
    return runs


def run_lengths(sectors):
    if not sectors:
        return []
    lengths = [1]
    for prev, cur in zip(sectors, sectors[1:]):
        if cur == prev + 1:
            lengths[-1] += 1
        else:
            lengths.append(1)
    return lengths


class CfbGeometry:
    """Header, DIFAT, FAT, MiniFAT and directory geometry parsed from raw bytes.

    Mirrors what litchi-cfb's OleFile::open_with_limits reads, so that per-sector
    read counts can be modelled from the same structures.
    """

    def __init__(self, data):
        self.data = data
        self.issues = []
        header = data[:512]
        if header[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("not a CFB file")
        self.dll_version = u16(header, 0x1A)
        self.sector_shift = u16(header, 0x1E)
        self.mini_sector_shift = u16(header, 0x20)
        self.num_dir_sectors = u32(header, 0x28)
        self.num_fat_sectors = u32(header, 0x2C)
        self.first_dir_sector = u32(header, 0x30)
        self.mini_stream_cutoff = u32(header, 0x38)
        self.first_minifat_sector = u32(header, 0x3C)
        self.num_minifat_sectors = u32(header, 0x40)
        self.first_difat_sector = u32(header, 0x44)
        self.num_difat_sectors = u32(header, 0x48)
        self.sector_size = 1 << self.sector_shift
        self.mini_sector_size = 1 << self.mini_sector_shift
        self.file_size = len(data)
        # litchi: physical_sector_count = file_size / sector_size - 1 (floor)
        self.physical_sectors = self.file_size // self.sector_size - 1
        # olefile: ceil-based count
        self.physical_sectors_ceil = (self.file_size + self.sector_size - 1) // self.sector_size - 1
        self.file_is_whole_sectors = (self.file_size % self.sector_size) == 0

        # FAT sector locations, in DIFAT order.
        fat_sectors = []
        header_count = min(HEADER_DIFAT_ENTRIES, self.num_fat_sectors)
        for i in range(HEADER_DIFAT_ENTRIES):
            entry = u32(header, HEADER_DIFAT_OFFSET + 4 * i)
            if i < header_count:
                fat_sectors.append(entry)
            elif entry != FREESECT:
                self.issues.append("header DIFAT entry %d past the FAT count is not FREESECT" % i)
        difat_sectors = []
        sector = self.first_difat_sector
        per_sector = self.sector_size // 4 - 1
        for _ in range(self.num_difat_sectors):
            if sector > MAXREGSECT or sector in difat_sectors:
                self.issues.append("DIFAT chain broken at %r" % sector)
                break
            difat_sectors.append(sector)
            block = self.sector_bytes(sector)
            for j in range(per_sector):
                if len(fat_sectors) < self.num_fat_sectors:
                    fat_sectors.append(u32(block, 4 * j))
            sector = u32(block, 4 * per_sector)
        if len(fat_sectors) != self.num_fat_sectors:
            self.issues.append("expected %d FAT sectors, DIFAT lists %d" % (self.num_fat_sectors, len(fat_sectors)))
        self.fat_sectors = fat_sectors
        self.difat_sectors = difat_sectors

        fat = []
        for sid in fat_sectors:
            block = self.sector_bytes(sid)
            fat.extend(struct.unpack_from("<%dI" % (self.sector_size // 4), block, 0))
        self.fat = fat
        for sid in fat_sectors:
            if sid >= len(fat) or fat[sid] != FATSECT:
                self.issues.append("FAT sector %d is not marked FATSECT" % sid)
        for sid in difat_sectors:
            if sid >= len(fat) or fat[sid] != DIFSECT:
                self.issues.append("DIFAT sector %d is not marked DIFSECT" % sid)

        # MiniFAT chain (litchi follows exactly num_minifat_sectors links).
        self.minifat_sectors = self.chain(self.first_minifat_sector, exact=self.num_minifat_sectors, label="MiniFAT") if self.num_minifat_sectors else []
        # Directory chain: V3 walks to ENDOFCHAIN, V4 uses the declared count.
        if self.dll_version == 4 and self.num_dir_sectors:
            self.dir_sectors = self.chain(self.first_dir_sector, exact=self.num_dir_sectors, label="directory")
        else:
            self.dir_sectors = self.chain(self.first_dir_sector, label="directory")
        dir_data = b"".join(self.sector_bytes(s) for s in self.dir_sectors)
        self.dir_data = dir_data
        self.root_start = u32(dir_data, 0x74) if len(dir_data) >= 128 else ENDOFCHAIN
        self.root_size = u32(dir_data, 0x78) if len(dir_data) >= 128 else 0
        if self.dll_version == 4 and len(dir_data) >= 128:
            self.root_size |= u32(dir_data, 0x7C) << 32
        self.ministream_sectors = self.chain(self.root_start, label="mini stream") if self.root_size else []
        minifat = []
        for sid in self.minifat_sectors:
            block = self.sector_bytes(sid)
            minifat.extend(struct.unpack_from("<%dI" % (self.sector_size // 4), block, 0))
        self.minifat = minifat

    def sector_bytes(self, sid):
        pos = (sid + 1) * self.sector_size
        block = self.data[pos:pos + self.sector_size]
        if len(block) < self.sector_size:
            block = block + b"\0" * (self.sector_size - len(block))
        return block

    def chain(self, start, exact=None, label="chain"):
        out = []
        seen = set()
        sector = start
        limit = exact if exact is not None else len(self.fat) + 1
        while len(out) < limit:
            if sector == ENDOFCHAIN:
                if exact is not None:
                    self.issues.append("%s chain ends before its declared length" % label)
                break
            if sector > MAXREGSECT or sector >= len(self.fat) or sector in seen:
                self.issues.append("%s chain broken at %r" % (label, sector))
                break
            seen.add(sector)
            out.append(sector)
            sector = self.fat[sector]
        return out

    def mini_chain(self, start, needed):
        out = []
        seen = set()
        sector = start
        while len(out) < needed:
            if sector == ENDOFCHAIN or sector > MAXREGSECT or sector >= len(self.minifat) or sector in seen:
                self.issues.append("mini chain broken at %r" % sector)
                break
            seen.add(sector)
            out.append(sector)
            sector = self.minifat[sector]
        return out

    def stream_pieces(self, start, size, is_mini):
        """Maximal physically-contiguous pieces of a stream: (logical, length, physical)."""
        pieces = []
        if size == 0:
            return pieces
        if not is_mini:
            needed = (size + self.sector_size - 1) // self.sector_size
            sectors = self.chain(start, exact=needed, label="stream")
            for k, sid in enumerate(sectors):
                logical = k * self.sector_size
                length = min(self.sector_size, size - logical)
                physical = (sid + 1) * self.sector_size
                pieces.append([logical, length, physical])
        else:
            needed = (size + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE
            minis = self.mini_chain(start, needed)
            root = self.ministream_sectors
            for k, mid in enumerate(minis):
                logical = k * MINI_SECTOR_SIZE
                length = min(MINI_SECTOR_SIZE, size - logical)
                mini_off = mid * MINI_SECTOR_SIZE
                root_index = mini_off // self.sector_size
                if root_index >= len(root):
                    self.issues.append("mini sector %d is outside the root chain" % mid)
                    break
                physical = (root[root_index] + 1) * self.sector_size + mini_off % self.sector_size
                pieces.append([logical, length, physical])
        merged = []
        for piece in pieces:
            if merged and merged[-1][2] + merged[-1][1] == piece[2] and merged[-1][0] + merged[-1][1] == piece[0]:
                merged[-1][1] += piece[1]
            else:
                merged.append(list(piece))
        return merged


def spans_in(pieces, start, length):
    """Physically contiguous spans a logical range [start, start+length) splits into."""
    if length <= 0:
        return 0
    end = start + length
    count = 0
    for logical, plen, _physical in pieces:
        if logical < end and logical + plen > start:
            count += 1
    return count


def cfb_record(path, geo, ole):
    fat_runs = runs_of(geo.fat_sectors)
    minifat_runs = runs_of(geo.minifat_sectors)
    dir_runs = runs_of(geo.dir_sectors)
    ministream_runs = runs_of(geo.ministream_sectors)
    first_fat_is_sector0 = bool(geo.fat_sectors) and geo.fat_sectors[0] == 0
    # Sector 0 starts at file offset sector_size; the header occupies [0, 512).
    header_adjacent = first_fat_is_sector0 and geo.sector_size == 512
    fat_reads_today = len(geo.fat_sectors)
    fat_reads_batched = fat_runs
    open_reads_today = 1 + len(geo.difat_sectors) + len(geo.fat_sectors) + len(geo.minifat_sectors) + dir_runs
    open_reads_batched = 1 + len(geo.difat_sectors) + fat_runs + minifat_runs + dir_runs
    open_reads_batched_adjacent = open_reads_batched - (1 if header_adjacent else 0)
    return {
        "cfb_version": geo.dll_version,
        "sector_size": geo.sector_size,
        "file_size": geo.file_size,
        "file_is_whole_sectors": geo.file_is_whole_sectors,
        "physical_sectors": geo.physical_sectors,
        "fat_sectors": geo.fat_sectors,
        "fat_sector_count": len(geo.fat_sectors),
        "fat_runs": fat_runs,
        "fat_run_lengths": run_lengths(geo.fat_sectors),
        "fat_gaps_between_fat_sectors": [b - a for a, b in zip(geo.fat_sectors, geo.fat_sectors[1:])],
        "first_fat_sector_is_0": first_fat_is_sector0,
        "header_fat0_adjacent": header_adjacent,
        "difat_sector_count": len(geo.difat_sectors),
        "difat_sectors": geo.difat_sectors,
        "minifat_sector_count": len(geo.minifat_sectors),
        "minifat_sectors": geo.minifat_sectors,
        "minifat_runs": minifat_runs,
        "directory_sector_count": len(geo.dir_sectors),
        "directory_sectors": geo.dir_sectors,
        "directory_runs": dir_runs,
        "directory_entries": len(geo.dir_data) // 128,
        "ministream_size": geo.root_size,
        "ministream_sector_count": len(geo.ministream_sectors),
        "ministream_runs": ministream_runs,
        "mini_stream_cutoff": geo.mini_stream_cutoff,
        "stream_count": len(ole.listdir(streams=True, storages=False)),
        "load_fat_reads_today": fat_reads_today,
        "load_fat_reads_batched": fat_reads_batched,
        "load_fat_reads_saved": fat_reads_today - fat_reads_batched,
        "load_minifat_reads_today": len(geo.minifat_sectors),
        "load_minifat_reads_batched": minifat_runs,
        "cfb_open_reads_today": open_reads_today,
        "cfb_open_reads_batched": open_reads_batched,
        "cfb_open_reads_batched_and_header_merge": open_reads_batched_adjacent,
        "geometry_issues": list(geo.issues),
    }


def walk_globals(stream):
    """Walk BIFF records from offset 0 until the first EOF (production rule).

    Returns (records, globals_end, depth0_end, issues).  ``records`` includes
    the BOF and the terminating EOF.  ``depth0_end`` is the end computed with
    BOF/EOF depth tracking, for comparison with the production rule that stops
    at the first EOF regardless of nesting.
    """
    records = []
    issues = []
    off = 0
    globals_end = None
    depth = 0
    depth0_end = None
    n = len(stream)
    while off + 4 <= n:
        kind = u16(stream, off)
        length = u16(stream, off + 2)
        end = off + 4 + length
        if end > n:
            issues.append("record 0x%04X at %d exceeds the stream" % (kind, off))
            break
        records.append((off, kind, length))
        if kind == BOF or kind in OLD_BOFS:
            depth += 1
        if kind == EOF:
            if length != 0:
                issues.append("EOF at %d has a non-empty payload (%d bytes)" % (off, length))
            depth -= 1
            if globals_end is None:
                globals_end = end
            if depth <= 0:
                depth0_end = end
                break
        off = end
    if globals_end is None:
        issues.append("no EOF found before the end of the stream")
    if depth0_end is None and globals_end is not None:
        # Nested BOFs never closed before the stream ended; production stops at first EOF anyway.
        issues.append("BOF/EOF depth never returned to 0 (nested BOF inside globals)")
    return records, globals_end, depth0_end, issues


def model_new_schedule(records, globals_end, boundsheets, filepass_index, stream_len,
                       max_global_bytes=MAX_GLOBAL_BYTES):
    """Logical reads under today's pre-pass and under the single-pass schedule.

    The single-pass model mirrors the implemented algorithm: an exact prologue of
    ``EXACT_PROLOGUE_RECORDS`` records, each costing one header read and one payload
    read, then fills of ``WINDOW_FIRST`` bytes doubling to ``WINDOW_MAX``, each fill
    clamped by the stream length, by ``max_global_bytes``, and by the smallest
    BoundSheet8 ``lbPlyPos`` seen so far.  A fill is never shorter than the bytes the
    next check needs.
    """
    n = len(records)
    result = {
        "current_header_reads": None,
        "current_bulk_reads": None,
        "current_logical_reads": None,
        "new_prologue_reads": 0,
        "new_window_reads": 0,
        "new_logical_reads": None,
        "window_sizes": [],
        "first_window_offset": None,
        "clamp": None,
        "clamp_known_at_offset": None,
        "overread_bytes_past_globals_end": 0,
        "phase_note": "",
        "read_ranges": [],
    }
    if globals_end is None and filepass_index is None:
        result["phase_note"] = "globals never terminate; not modelled"
        return result

    # Today's schedule: one 4-byte header read per record, stopping at a FILEPASS
    # header, then one bulk read of [0, globals_end) when the globals framed cleanly.
    if filepass_index is not None:
        result["current_header_reads"] = filepass_index + 1
        result["current_bulk_reads"] = 0
        result["current_logical_reads"] = filepass_index + 1
    else:
        result["current_header_reads"] = n
        result["current_bulk_reads"] = 1
        result["current_logical_reads"] = n + 1

    filled = 0
    reads = []
    window = WINDOW_FIRST
    clamp = None
    clamp_at = None
    clamp_dropped = False
    first_window_offset = None
    prologue = 0
    windows = 0
    fills_after_clamp_dropped = [None]

    def fill(need, index, prefetch=False):
        nonlocal filled, window, windows, prologue, first_window_offset
        if need <= filled:
            return
        if index < EXACT_PROLOGUE_RECORDS:
            # The prologue fetches a record's payload together with the next
            # record's four-byte header, so an encrypted workbook is refused at
            # a header that arrived with the previous record's read and no byte
            # of its payload is read.  The extra four bytes are clamped only by
            # the stream length, matching the pre-change header read, which also
            # ran before the record's own max_global_bytes check.
            end = need + 4 if (prefetch and need + 4 <= stream_len) else need
            reads.append((filled, end - filled))
            filled = end
            prologue += 1
            return
        nonlocal clamp_dropped
        if first_window_offset is None:
            first_window_offset = filled
        if clamp is not None and not clamp_dropped and need > clamp:
            # Globals frame past the smallest declared sheet position, so
            # lbPlyPos is corrupt; production drops the clamp for the rest of
            # the scan.  Framing that merely overtakes the clamp does not reach
            # here, because no fill is required once the bytes are resident.
            clamp_dropped = True
            fills_after_clamp_dropped[0] = windows
        # max_global_bytes bounds retained globals, not the four header bytes
        # that prove a record crosses it, so the cap is never below `need`.
        cap = min(stream_len, max(max_global_bytes, need))
        if clamp is not None and not clamp_dropped:
            cap = min(cap, clamp)
        end = min(max(need, filled + window), max(cap, need))
        reads.append((filled, end - filled))
        filled = end
        windows += 1
        window = min(window * 2, WINDOW_MAX)

    for index, (offset, kind, length) in enumerate(records):
        fill(offset + 4, index)                      # the record header
        if kind == FILEPASS:
            result["phase_note"] = (
                "FILEPASS at record %d: refused at its header%s"
                % (index, "" if index < EXACT_PROLOGUE_RECORDS
                   else "; payload bytes may be resident in a fill buffer and are never framed"))
            break
        if kind == EOF:
            break
        # The payload; in the prologue this also carries the next record's header.
        fill(offset + 4 + length, index, prefetch=True)
        if kind == BOUNDSHEET8 and length >= 4:
            for b in boundsheets:
                if b["index"] == index:
                    clamp = b["lbPlyPos"] if clamp is None else min(clamp, b["lbPlyPos"])
                    if clamp_at is None:
                        clamp_at = filled
                    break

    if clamp_dropped:
        result["phase_note"] = (
            "globals frame past the smallest BoundSheet8 position, so the clamp was dropped; "
            "%d later fill(s) ran bounded only by the stream length and max_global_bytes"
            % max(0, windows - (fills_after_clamp_dropped[0] or 0)))
    result.update({
        "new_prologue_reads": prologue,
        "new_window_reads": windows,
        "new_logical_reads": prologue + windows,
        "window_sizes": [length for _offset, length in reads[prologue:]],
        "first_window_offset": first_window_offset,
        "clamp": clamp,
        "clamp_known_at_offset": clamp_at,
        "overread_bytes_past_globals_end": (
            max(0, filled - globals_end) if globals_end is not None else 0),
        "read_ranges": reads,
    })
    return result

def biff_record(path, geo, ole):
    names = ["Workbook", "Book"]
    present = [n for n in names if ole.exists(n)]
    if not present:
        return {"workbook_stream": None, "note": "no Workbook or Book stream"}
    name = present[0]
    sid = ole._find(name)
    entry = ole.direntries[sid]
    stream = ole.openstream(name).read()
    stream_len = len(stream)
    is_mini = entry.size < geo.mini_stream_cutoff
    pieces = geo.stream_pieces(entry.isectStart, entry.size, is_mini)
    records, globals_end, depth0_end, issues = walk_globals(stream)
    if not records:
        return {"workbook_stream": name, "stream_len": stream_len, "note": "empty stream", "issues": issues}
    first_off, first_kind, first_len = records[0]
    version = None
    substream = None
    if first_kind == BOF and first_len >= 4:
        vers = u16(stream, first_off + 4)
        substream = u16(stream, first_off + 6)
        version = {0x0600: "BIFF8", 0x0500: "BIFF5/7"}.get(vers, "BOF vers=0x%04X" % vers)
    elif first_kind in OLD_BOFS:
        version = OLD_BOFS[first_kind]
    else:
        version = "no BOF at offset 0 (0x%04X)" % first_kind
        issues.append("stream does not start with BOF")
    boundsheets = []
    for i, (off, kind, length) in enumerate(records):
        if kind == BOUNDSHEET8:
            lb = u32(stream, off + 4) if length >= 4 else None
            if lb is None:
                issues.append("BoundSheet8 at %d is shorter than 4 bytes" % off)
            boundsheets.append({"index": i, "offset": off, "lbPlyPos": lb})
    filepass = [(i, off) for i, (off, kind, _l) in enumerate(records) if kind == FILEPASS]
    filepass_index = filepass[0][0] if filepass else None
    filepass_offset = filepass[0][1] if filepass else None
    nested_bofs = sum(1 for off, kind, _l in records[1:] if kind == BOF or kind in OLD_BOFS)
    lbs = [b["lbPlyPos"] for b in boundsheets if b["lbPlyPos"] is not None]
    first_bs_index = boundsheets[0]["index"] if boundsheets else None
    last_bs_index = boundsheets[-1]["index"] if boundsheets else None
    contiguous = bool(boundsheets) and (last_bs_index - first_bs_index + 1 == len(boundsheets))
    ascending = bool(lbs) and all(a < b for a, b in zip(lbs, lbs[1:]))
    non_decreasing = bool(lbs) and all(a <= b for a, b in zip(lbs, lbs[1:]))
    min_lb = min(lbs) if lbs else None
    slack = (min_lb - globals_end) if (min_lb is not None and globals_end is not None) else None
    corrupt = bool(lbs) and globals_end is not None and (any(lb < globals_end for lb in lbs) or any(lb >= stream_len for lb in lbs))
    hist = collections.Counter(kind for _o, kind, _l in records[:first_bs_index]) if first_bs_index is not None else collections.Counter(kind for _o, kind, _l in records)
    top = sorted(hist.items(), key=lambda kv: (-kv[1], kv[0]))[:8]
    reads = model_new_schedule(records, globals_end, boundsheets, filepass_index, stream_len)

    # Physical-span model (extra): how many contiguous spans each schedule's logical reads split into.
    stop = filepass_index + 1 if filepass_index is not None else len(records)
    current_phys = sum(spans_in(pieces, records[i][0], 4) for i in range(stop))
    if reads["current_bulk_reads"]:
        current_phys += spans_in(pieces, 0, globals_end)
    new_phys = sum(spans_in(pieces, s, l) for s, l in reads["read_ranges"])
    reads.pop("read_ranges")

    ratio = None
    if reads["current_logical_reads"] and reads["new_logical_reads"] is not None:
        ratio = reads["new_logical_reads"] / reads["current_logical_reads"]
    return {
        "workbook_stream": name,
        "both_book_and_workbook": len(present) == 2,
        "stream_len": stream_len,
        "stream_in_ministream": is_mini,
        "stream_physical_pieces": len(pieces),
        "biff_version": version,
        "bof_substream_type": substream,
        "globals_end": globals_end,
        "globals_end_depth0": depth0_end,
        "nested_bofs_before_first_eof": nested_bofs,
        "records_in_globals": len(records),
        "records_before_first_boundsheet8": first_bs_index,
        "first_boundsheet8_offset": boundsheets[0]["offset"] if boundsheets else None,
        "boundsheet8_count": len(boundsheets),
        "boundsheet8_contiguous": contiguous,
        "last_boundsheet8_index": last_bs_index,
        "record_after_boundsheet8_run": record_name(records[last_bs_index + 1][1]) if boundsheets and last_bs_index + 1 < len(records) else None,
        "lbPlyPos": lbs,
        "lbPlyPos_ascending": ascending,
        "lbPlyPos_non_decreasing": non_decreasing,
        "min_lbPlyPos": min_lb,
        "slack_min_lbPlyPos_minus_globals_end": slack,
        "lbPlyPos_corrupt": corrupt,
        "filepass": filepass_index is not None,
        "filepass_index": filepass_index,
        "filepass_offset": filepass_offset,
        "record_type_histogram_before_first_boundsheet8_top8": [[record_name(k), c] for k, c in top],
        "record_types_before_first_boundsheet8_distinct": len(hist),
        "reads": reads,
        "ratio_new_over_current": ratio,
        "physical_spans_current": current_phys,
        "physical_spans_new": new_phys,
        "issues": issues,
    }


def survey_file(path, want_biff, repo="."):
    # Record the repository-relative path so the artifact is byte-identical
    # regardless of how --repo was spelled on the command line.
    rec = {"path": os.path.relpath(path, repo), "size": os.path.getsize(path)}
    if not olefile.isOleFile(path):
        rec["skipped"] = "olefile.isOleFile is false"
        return rec
    try:
        ole = olefile.OleFileIO(path)
    except Exception as error:  # noqa: BLE001 - report and continue
        rec["skipped"] = "olefile cannot open: %s: %s" % (type(error).__name__, error)
        return rec
    try:
        with open(path, "rb") as handle:
            data = handle.read()
        geo = CfbGeometry(data)
        # Cross-check the independently parsed FAT against olefile's.
        ole_fat = list(ole.fat)
        if ole_fat[:len(geo.fat)] != geo.fat[:len(ole_fat)]:
            geo.issues.append("FAT differs from olefile's FAT")
        rec["cfb"] = cfb_record(path, geo, ole)
        rec["olefile_defects"] = ["%s: %s" % (lvl, msg) for lvl, msg in ole.parsing_issues]
        if want_biff:
            rec["biff"] = biff_record(path, geo, ole)
    except Exception as error:  # noqa: BLE001
        rec["skipped"] = "survey error: %s: %s" % (type(error).__name__, error)
    finally:
        ole.close()
    return rec


def fmt_bool(value):
    return "yes" if value else "no"


def fmt(value):
    return "-" if value is None else str(value)


def median(values):
    return statistics.median(values) if values else None


def write_markdown(out, main, extra, repo):
    lines = []
    w = lines.append
    w("# Track C survey: CFB FAT geometry and BIFF globals read schedules")
    w("")
    w("All read counts below are a **model of stream-level (logical) reads** derived from")
    w("each file's CFB and BIFF geometry, mirroring `parse_globals` in")
    w("`crates/litchi-xls/src/workbook/source.rs` (today: one 4-byte header read per globals")
    w("record including EOF, abort at FILEPASS before its payload, then one bulk range read of")
    w("`[0, globals_end)`) and the single-pass schedule of change 0565: an exact prologue of the")
    w("first %d records, each costing one header read and one payload read, then fills of %d bytes" % (EXACT_PROLOGUE_RECORDS, WINDOW_FIRST))
    w("doubling to %d, every fill clamped by the stream length, by `max_global_bytes`, and by the" % WINDOW_MAX)
    w("smallest BoundSheet8 `lbPlyPos` seen so far. The `phys` columns count the physically contiguous")
    w("spans those logical reads would split into under `SharedOleFile`'s run grouping. **Nothing")
    w("here is a measured syscall count.** Generated by `survey.py`; per-file records are in")
    w("`survey.json`.")
    w("")
    xls = [r for r in main if "biff" in r and r["biff"].get("workbook_stream")]
    skipped = [r for r in main if "skipped" in r]
    cfb_ok = [r for r in main if "cfb" in r]
    w("Scope: %d files under `test-data/ole/`; %d opened as CFB; %d skipped (listed at the end)." % (len(main), len(cfb_ok), len(skipped)))
    w("Appendix: %d `*.xls` fixtures elsewhere under `test-data/` (POI, LibreOffice, interop), surveyed the same way." % len(extra))
    w("")

    def xls_table(rows, title):
        w("## %s" % title)
        w("")
        w("| file | ver | stream | mini | globals end | recs | recs<BS8 | BS8 | contig | asc | slack | FILEPASS | current | new | ratio | windows | phys cur→new |")
        w("| --- | --- | ---: | --- | ---: | ---: | ---: | ---: | --- | --- | ---: | --- | ---: | ---: | ---: | --- | --- |")
        for r in rows:
            b = r["biff"]
            rd = b["reads"]
            fp = "rec %s" % b["filepass_index"] if b["filepass"] else "no"
            ratio = "%.3f" % b["ratio_new_over_current"] if b["ratio_new_over_current"] is not None else "-"
            windows = ",".join(str(s) for s in rd["window_sizes"]) if rd["window_sizes"] else ("-" if rd["new_window_reads"] == 0 else "?")

            w("| %s | %s | %d | %s | %s | %d | %s | %d | %s | %s | %s | %s | %s | %s | %s | %s | %d→%d |" % (
                r["path"].replace("test-data/", ""), b["biff_version"], b["stream_len"], fmt_bool(b["stream_in_ministream"]),
                fmt(b["globals_end"]), b["records_in_globals"], fmt(b["records_before_first_boundsheet8"]), b["boundsheet8_count"],
                fmt_bool(b["boundsheet8_contiguous"]), fmt_bool(b["lbPlyPos_ascending"]), fmt(b["slack_min_lbPlyPos_minus_globals_end"]), fp,
                fmt(rd["current_logical_reads"]), fmt(rd["new_logical_reads"]), ratio, windows, b["physical_spans_current"], b["physical_spans_new"]))
        w("")

    xls_table(xls, "XLS: BIFF globals geometry and modelled reads (test-data/ole/xls)")

    def cfb_table(rows, title):
        w("## %s" % title)
        w("")
        w("| file | sect | sectors | FAT n | FAT runs | FAT gaps | FAT0=sect0 | DIFAT | MiniFAT n/runs | dir n/runs | ministream bytes | open reads today | batched | +hdr merge |")
        w("| --- | ---: | ---: | ---: | ---: | --- | --- | ---: | --- | --- | ---: | ---: | ---: | ---: |")
        for r in rows:
            c = r["cfb"]
            gaps = c["fat_gaps_between_fat_sectors"]
            gap_text = "-" if not gaps else (",".join(str(g) for g in gaps[:6]) + ("…" if len(gaps) > 6 else ""))
            w("| %s | %d | %d | %d | %d | %s | %s | %d | %d/%d | %d/%d | %d | %d | %d | %d |" % (
                r["path"].replace("test-data/", ""), c["sector_size"], c["physical_sectors"], c["fat_sector_count"], c["fat_runs"], gap_text,
                fmt_bool(c["first_fat_sector_is_0"]), c["difat_sector_count"], c["minifat_sector_count"], c["minifat_runs"],
                c["directory_sector_count"], c["directory_runs"], c["ministream_size"], c["cfb_open_reads_today"], c["cfb_open_reads_batched"],
                c["cfb_open_reads_batched_and_header_merge"]))
        w("")

    cfb_table(cfb_ok, "CFB FAT geometry, all files under test-data/ole")

    # Summary
    w("## Summary")
    w("")
    n_cfb = len(cfb_ok)
    fat_save_files = [r for r in cfb_ok if r["cfb"]["load_fat_reads_saved"] > 0]
    fat_saved_total = sum(r["cfb"]["load_fat_reads_saved"] for r in cfb_ok)
    multi_fat = [r for r in cfb_ok if r["cfb"]["fat_sector_count"] > 1]
    minifat_save_files = [r for r in cfb_ok if r["cfb"]["load_minifat_reads_today"] - r["cfb"]["load_minifat_reads_batched"] > 0]
    minifat_saved_total = sum(r["cfb"]["load_minifat_reads_today"] - r["cfb"]["load_minifat_reads_batched"] for r in cfb_ok)
    adjacent = [r for r in cfb_ok if r["cfb"]["header_fat0_adjacent"]]
    sector0 = [r for r in cfb_ok if r["cfb"]["first_fat_sector_is_0"]]
    sizes = collections.Counter(r["cfb"]["sector_size"] for r in cfb_ok)
    open_today = sum(r["cfb"]["cfb_open_reads_today"] for r in cfb_ok)
    open_batched = sum(r["cfb"]["cfb_open_reads_batched"] for r in cfb_ok)
    open_merged = sum(r["cfb"]["cfb_open_reads_batched_and_header_merge"] for r in cfb_ok)
    w("### CFB-level follow-ups")
    w("")
    w("- Sector sizes: %s." % ", ".join("%d bytes x%d" % (k, v) for k, v in sorted(sizes.items())))
    w("- **FAT batching (`load_fat`)**: %d of %d files have more than one FAT sector; batching contiguous runs saves reads on **%d of %d files (%.0f%%)**, %d reads in total across the corpus." % (
        len(multi_fat), n_cfb, len(fat_save_files), n_cfb, 100.0 * len(fat_save_files) / max(1, n_cfb), fat_saved_total))
    for r in sorted(fat_save_files, key=lambda r: -r["cfb"]["load_fat_reads_saved"]):
        c = r["cfb"]
        w("  - `%s`: %d FAT sectors in %d runs (%d → %d reads, saves %d); run lengths %s." % (
            r["path"], c["fat_sector_count"], c["fat_runs"], c["load_fat_reads_today"], c["load_fat_reads_batched"], c["load_fat_reads_saved"], c["fat_run_lengths"]))
    no_save_multi = [r for r in multi_fat if r["cfb"]["load_fat_reads_saved"] == 0]
    if no_save_multi:
        w("  - Multi-FAT files where batching saves nothing (every FAT sector isolated): %s." % ", ".join(
            "`%s` (%d sectors, gaps %s)" % (os.path.basename(r["path"]), r["cfb"]["fat_sector_count"], r["cfb"]["fat_gaps_between_fat_sectors"][:4]) for r in no_save_multi))
    w("- **MiniFAT batching (`load_minifat`)**: saves reads on %d of %d files, %d reads in total." % (len(minifat_save_files), n_cfb, minifat_saved_total))
    for r in minifat_save_files:
        c = r["cfb"]
        w("  - `%s`: %d MiniFAT sectors in %d runs." % (r["path"], c["minifat_sector_count"], c["minifat_runs"]))
    w("- **Header / FAT-sector-0 adjacency**: the first FAT sector is physical sector 0 in **%d of %d files (%.0f%%)**; with 512-byte sectors that sector starts at file offset 512, immediately after the header, so a speculative 1,024-byte first read would cover both in those files. It does not hold in: %s." % (
        len(sector0), n_cfb, 100.0 * len(sector0) / max(1, n_cfb),
        ", ".join("`%s` (first FAT sector %s)" % (os.path.basename(r["path"]), r["cfb"]["fat_sectors"][:1]) for r in cfb_ok if not r["cfb"]["first_fat_sector_is_0"]) or "none"))
    w("- Modelled CFB open reads (header + DIFAT + FAT + MiniFAT + batched directory) across the corpus: **%d today → %d with FAT/MiniFAT batching → %d with the header merge too** (%d files)." % (open_today, open_batched, open_merged, n_cfb))
    dir_frag = [r for r in cfb_ok if r["cfb"]["directory_runs"] > 1]
    w("- Directory chains are already batched; %d of %d files have a fragmented directory chain (%s)." % (
        len(dir_frag), n_cfb, ", ".join("`%s` %d sectors/%d runs" % (os.path.basename(r["path"]), r["cfb"]["directory_sector_count"], r["cfb"]["directory_runs"]) for r in dir_frag) or "none"))
    w("")

    w("### BIFF globals schedule (test-data/ole/xls)")
    w("")
    modelled = [r for r in xls if r["biff"]["ratio_new_over_current"] is not None]
    ratios = [r["biff"]["ratio_new_over_current"] for r in modelled]
    cur_total = sum(r["biff"]["reads"]["current_logical_reads"] for r in modelled)
    new_total = sum(r["biff"]["reads"]["new_logical_reads"] for r in modelled)
    w("- %d XLS files carry a `Workbook`/`Book` stream; %d modelled. Aggregate logical reads for one open of every file: **%d today → %d new (%.1f%%)**." % (
        len(xls), len(modelled), cur_total, new_total, 100.0 * new_total / max(1, cur_total)))
    if ratios:
        best = min(modelled, key=lambda r: r["biff"]["ratio_new_over_current"])
        worst = max(modelled, key=lambda r: r["biff"]["ratio_new_over_current"])
        w("- Ratio new/current per file: **median %.3f**, best %.3f (`%s`), **worst %.3f (`%s`, %s)**." % (
            median(ratios), best["biff"]["ratio_new_over_current"], os.path.basename(best["path"]),
            worst["biff"]["ratio_new_over_current"], os.path.basename(worst["path"]), worst["biff"]["reads"]["phase_note"] or "see table"))
        non_fp = [r for r in modelled if not r["biff"]["filepass"]]
        if non_fp:
            nf_ratios = [r["biff"]["ratio_new_over_current"] for r in non_fp]
            worst_nf = max(non_fp, key=lambda r: r["biff"]["ratio_new_over_current"])
            w("- Excluding FILEPASS carriers (where both schedules stop at the same header): median %.3f, worst %.3f (`%s`: %d → %d)." % (
                median(nf_ratios), worst_nf["biff"]["ratio_new_over_current"], os.path.basename(worst_nf["path"]),
                worst_nf["biff"]["reads"]["current_logical_reads"], worst_nf["biff"]["reads"]["new_logical_reads"]))
    before = [r["biff"]["records_before_first_boundsheet8"] for r in xls if r["biff"]["records_before_first_boundsheet8"] is not None]
    if before:
        dist = collections.Counter(before)
        w("- Records before the first BoundSheet8: min %d, median %s, max %d; distribution %s." % (
            min(before), median(before), max(before), ", ".join("%d×%d" % (k, v) for k, v in sorted(dist.items()))))
        exact_reads = [r["biff"]["reads"]["new_prologue_reads"] for r in modelled if r["biff"]["reads"]["new_prologue_reads"] is not None]
        w("- Prologue reads (at most two per prologue record, fewer when a fill already covered the bytes): min %d, median %s, max %d. The prologue is the floor of the new schedule; the fill phase adds %s reads per file (median %s)." % (
            min(exact_reads), median(exact_reads), max(exact_reads),
            "%d–%d" % (min(r["biff"]["reads"]["new_window_reads"] for r in modelled), max(r["biff"]["reads"]["new_window_reads"] for r in modelled)),
            median([r["biff"]["reads"]["new_window_reads"] for r in modelled])))
    contig = [r for r in xls if r["biff"]["boundsheet8_count"] and r["biff"]["boundsheet8_contiguous"]]
    asc = [r for r in xls if r["biff"]["lbPlyPos"] and r["biff"]["lbPlyPos_ascending"]]
    with_bs = [r for r in xls if r["biff"]["boundsheet8_count"]]
    w("- BoundSheet8 runs are contiguous in %d of %d files with BoundSheet8; `lbPlyPos` strictly ascending in %d of %d." % (len(contig), len(with_bs), len(asc), len(with_bs)))
    after = collections.Counter(r["biff"]["record_after_boundsheet8_run"] for r in with_bs)
    w("- Record following the BoundSheet8 run: %s." % ", ".join("%s ×%d" % (k, v) for k, v in sorted(after.items(), key=lambda kv: -kv[1])))
    slacks = [r["biff"]["slack_min_lbPlyPos_minus_globals_end"] for r in with_bs if r["biff"]["slack_min_lbPlyPos_minus_globals_end"] is not None]
    if slacks:
        sl = collections.Counter(slacks)
        overreads = [r["biff"]["reads"]["overread_bytes_past_globals_end"] for r in modelled]
        w("- Slack (min `lbPlyPos` − globals end): %s. Zero slack means the first sheet BOF immediately follows the globals EOF, so once the clamp is known a fill stops exactly at the globals end. Every byte read past the globals end is therefore read before the first BoundSheet8 is framed: %s." % (
            ", ".join("%d×%d" % (k, v) for k, v in sorted(sl.items())),
            "0 bytes in every file" if not any(overreads) else
            "%d bytes across %d of %d files, worst %d (`%s`)" % (
                sum(overreads), sum(1 for o in overreads if o), len(overreads), max(overreads),
                os.path.basename(max(modelled, key=lambda r: r["biff"]["reads"]["overread_bytes_past_globals_end"])["path"]))))
        dropped = [r for r in modelled if "clamp was dropped" in (r["biff"]["reads"]["phase_note"] or "")]
        w("- Files where the sheet clamp was dropped because the globals frame past the smallest `lbPlyPos`, which is the only condition under which production drops it: **%d of %d**. A fill that merely overtakes the clamp does not drop it, because once the bytes are resident no fill is required." % (
            len(dropped), len(modelled)))
    corrupt = [r for r in xls if r["biff"]["lbPlyPos_corrupt"]]
    w("- Corrupt `lbPlyPos` (below globals end or beyond the stream): %s." % (", ".join("`%s`" % os.path.basename(r["path"]) for r in corrupt) or "none"))
    fps = [r for r in xls if r["biff"]["filepass"]]
    w("- FILEPASS carriers: %s." % (", ".join("`%s` (record %d, offset %d)" % (r["path"], r["biff"]["filepass_index"], r["biff"]["filepass_offset"]) for r in fps) or "none"))
    extra_fps = [r for r in extra if "biff" in r and r["biff"].get("filepass")]
    if extra_fps:
        w("  Outside `test-data/ole`: %s." % ", ".join("`%s` (record %d, offset %d)" % (r["path"], r["biff"]["filepass_index"], r["biff"]["filepass_offset"]) for r in extra_fps))
    versions = collections.Counter(r["biff"]["biff_version"] for r in xls)
    w("- BIFF versions: %s." % ", ".join("%s ×%d" % (k, v) for k, v in sorted(versions.items())))
    mini = [r for r in xls if r["biff"]["stream_in_ministream"]]
    w("- Workbook streams resident in the mini stream (below the 4,096-byte cutoff): %s." % (", ".join("`%s` (%d bytes)" % (os.path.basename(r["path"]), r["biff"]["stream_len"]) for r in mini) or "none"))
    both = [r for r in xls if r["biff"]["both_book_and_workbook"]]
    if both:
        w("- Files carrying both a `Book` and a `Workbook` stream (production selects `Workbook`): %s." % ", ".join("`%s`" % os.path.basename(r["path"]) for r in both))
    nested = [r for r in xls if r["biff"]["nested_bofs_before_first_eof"]]
    w("- Nested BOF before the first EOF (production's first-EOF rule vs depth-0 rule would differ): %s." % (", ".join("`%s`" % os.path.basename(r["path"]) for r in nested) or "none"))
    issues = [(r, r["biff"]["issues"]) for r in xls if r["biff"].get("issues")]
    if issues:
        w("- BIFF walk issues: %s." % "; ".join("`%s`: %s" % (os.path.basename(r["path"]), "; ".join(i)) for r, i in issues))
    geo_issues = [(r, r["cfb"]["geometry_issues"]) for r in cfb_ok if r["cfb"]["geometry_issues"]]
    w("- CFB geometry issues: %s." % ("; ".join("`%s`: %s" % (os.path.basename(r["path"]), "; ".join(i)) for r, i in geo_issues) or "none"))
    defects = [(r, r["olefile_defects"]) for r in cfb_ok if r.get("olefile_defects")]
    if defects:
        w("- olefile parsing issues: %s." % "; ".join("`%s`: %s" % (os.path.basename(r["path"]), "; ".join(i)) for r, i in defects))
    phys_cur = sum(r["biff"]["physical_spans_current"] for r in modelled)
    phys_new = sum(r["biff"]["physical_spans_new"] for r in modelled)
    w("- Physical-span model over the same logical reads: %d → %d spans across the modelled XLS files (straddle splits and fragmented chains included)." % (phys_cur, phys_new))
    w("")

    # Histogram aggregate
    w("### Record types before the first BoundSheet8 (aggregate over test-data/ole/xls)")
    w("")
    agg = collections.Counter()
    for r in xls:
        for name, count in r["biff"]["record_type_histogram_before_first_boundsheet8_top8"]:
            agg[name] += count
    w("| record type | occurrences (sum of per-file top-8) | files |")
    w("| --- | ---: | ---: |")
    files_with = collections.Counter()
    for r in xls:
        for name, _count in r["biff"]["record_type_histogram_before_first_boundsheet8_top8"]:
            files_with[name] += 1
    for name, count in sorted(agg.items(), key=lambda kv: (-kv[1], kv[0]))[:12]:
        w("| %s | %d | %d |" % (name, count, files_with[name]))
    w("")
    w("Per-file top-8 histograms are in `survey.json` under `biff.record_type_histogram_before_first_boundsheet8_top8`.")
    w("")

    # Appendix: extra xls
    extra_xls = [r for r in extra if "biff" in r and r["biff"].get("workbook_stream")]
    if extra_xls:
        xls_table(extra_xls, "Appendix A: XLS fixtures outside test-data/ole (POI, LibreOffice, interop)")
        ex_mod = [r for r in extra_xls if r["biff"]["ratio_new_over_current"] is not None]
        ex_ratios = [r["biff"]["ratio_new_over_current"] for r in ex_mod]
        ex_before = [r["biff"]["records_before_first_boundsheet8"] for r in extra_xls if r["biff"]["records_before_first_boundsheet8"] is not None]
        ex_cur = sum(r["biff"]["reads"]["current_logical_reads"] for r in ex_mod)
        ex_new = sum(r["biff"]["reads"]["new_logical_reads"] for r in ex_mod)
        worst = max(ex_mod, key=lambda r: r["biff"]["ratio_new_over_current"]) if ex_mod else None
        w("Appendix aggregate: %d files, %d modelled; logical reads %d → %d; ratio median %.3f, worst %.3f (`%s`); records before first BoundSheet8 min %d / median %s / max %d; contiguous BoundSheet8 runs in %d of %d; ascending `lbPlyPos` in %d of %d; corrupt `lbPlyPos` in %s; nested BOF in %s; FAT batching would save reads in %d of %d." % (
            len(extra_xls), len(ex_mod), ex_cur, ex_new, median(ex_ratios) if ex_ratios else float("nan"),
            worst["biff"]["ratio_new_over_current"] if worst else float("nan"), os.path.basename(worst["path"]) if worst else "-",
            min(ex_before), median(ex_before), max(ex_before),
            sum(1 for r in extra_xls if r["biff"]["boundsheet8_count"] and r["biff"]["boundsheet8_contiguous"]), sum(1 for r in extra_xls if r["biff"]["boundsheet8_count"]),
            sum(1 for r in extra_xls if r["biff"]["lbPlyPos"] and r["biff"]["lbPlyPos_ascending"]), sum(1 for r in extra_xls if r["biff"]["lbPlyPos"]),
            ", ".join("`%s`" % os.path.basename(r["path"]) for r in extra_xls if r["biff"]["lbPlyPos_corrupt"]) or "none",
            ", ".join("`%s`" % os.path.basename(r["path"]) for r in extra_xls if r["biff"]["nested_bofs_before_first_eof"]) or "none",
            sum(1 for r in extra if "cfb" in r and r["cfb"]["load_fat_reads_saved"] > 0), sum(1 for r in extra if "cfb" in r)))
        w("")
        cfb_table([r for r in extra if "cfb" in r], "Appendix B: CFB FAT geometry of the XLS fixtures outside test-data/ole")
    ex_skipped = [r for r in extra if "skipped" in r]

    w("## Skipped files")
    w("")
    for r in skipped:
        w("- `%s`: %s" % (r["path"], r["skipped"]))
    for r in ex_skipped:
        w("- `%s` (appendix): %s" % (r["path"], r["skipped"]))
    if not skipped and not ex_skipped:
        w("- none")
    w("")
    with open(out, "w") as handle:
        handle.write("\n".join(lines))


def main():
    global WINDOW_FIRST, WINDOW_MAX, EXACT_PROLOGUE_RECORDS
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--repo", default="/home/zhuhe/code/litchi")
    parser.add_argument("--first-window", type=int, default=WINDOW_FIRST,
                        help="first fill size in bytes, doubling to --max-window; "
                             "the rejected 4096-byte draft is reproduced with --first-window 4096")
    parser.add_argument("--max-window", type=int, default=WINDOW_MAX)
    parser.add_argument("--prologue-records", type=int, default=EXACT_PROLOGUE_RECORDS)
    parser.add_argument("--out", default=os.path.dirname(os.path.abspath(__file__)))
    parser.add_argument("--no-extra", action="store_true", help="skip the *.xls fixtures outside test-data/ole")
    args = parser.parse_args()
    WINDOW_FIRST = args.first_window
    WINDOW_MAX = args.max_window
    EXACT_PROLOGUE_RECORDS = args.prologue_records
    root = os.path.join(args.repo, "test-data", "ole")
    main_paths = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        for name in sorted(filenames):
            main_paths.append(os.path.join(dirpath, name))
    main_paths.sort()
    main = [survey_file(p, want_biff=p.lower().endswith(".xls"), repo=args.repo) for p in main_paths]
    extra = []
    if not args.no_extra:
        extra_paths = []
        for dirpath, dirnames, filenames in os.walk(os.path.join(args.repo, "test-data")):
            dirnames.sort()
            if os.path.abspath(dirpath).startswith(os.path.abspath(root)):
                continue
            for name in sorted(filenames):
                if name.lower().endswith(".xls"):
                    extra_paths.append(os.path.join(dirpath, name))
        extra_paths.sort()
        extra = [survey_file(p, want_biff=True, repo=args.repo) for p in extra_paths]
    os.makedirs(args.out, exist_ok=True)
    payload = {
        "model_note": "All read counts are a model of stream-level logical reads (and the physically contiguous spans they split into); none are measured syscalls.",
        "schedule": {
            "current": "one 4-byte header read per globals record including EOF; abort at FILEPASS before its payload; then one bulk range read of [0, globals_end)",
            "new": "first 4-byte header, then one read of payload + next 4-byte header per record through the last BoundSheet8; then windows min(filled + W, min lbPlyPos), W = 4096 doubling to 65536, until globals_end",
            "window_first": WINDOW_FIRST, "window_max": WINDOW_MAX,
        },
        "olefile_version": olefile.__version__,
        "python_version": sys.version.split()[0],
        "files": main,
        "extra_xls": extra,
    }
    with open(os.path.join(args.out, "survey.json"), "w") as handle:
        json.dump(payload, handle, indent=1, sort_keys=True)
    write_markdown(os.path.join(args.out, "survey.md"), main, extra, args.repo)
    print("wrote", os.path.join(args.out, "survey.json"), "and survey.md;", len(main), "files,", len(extra), "extra")


if __name__ == "__main__":
    main()
