//! Unicode 17.0.0 predicates generated from the official UCD files.
//!
//! Do not edit manually; rerun unicode-data/generate.py after changing
//! the pinned source receipts.  The tables contain only Unicode scalar
//! values and are searched without allocation or ambient I/O.
//! Source SHA-256: UnicodeData.txt=2e1efc1dcb59c575eedf5ccae60f95229f706ee6d031835247d843c11d96470c, DerivedCoreProperties.txt=24c7fed1195c482faaefd5c1e7eb821c5ee1fb6de07ecdbaa64b56a99da22c08, CaseFolding.txt=ff8d8fefbf123574205085d6714c36149eb946d717a0c585c27f0f4ef58c4183, SpecialCasing.txt=efc25faf19de21b92c1194c111c932e03d2a5eaf18194e33f1156e96de4c9588, license.txt=e7a93b009565cfce55919a381437ac4db883e9da2126fa28b91d12732bc53d96

#[derive(Clone, Copy)]
struct Range {
    start: u32,
    end: u32,
}

const CLEAN_REMOVED: &[Range] = &[
    Range {
        start: 0x000000,
        end: 0x00001F,
    },
    Range {
        start: 0x00007F,
        end: 0x00009F,
    },
    Range {
        start: 0x000378,
        end: 0x000379,
    },
    Range {
        start: 0x000380,
        end: 0x000383,
    },
    Range {
        start: 0x00038B,
        end: 0x00038B,
    },
    Range {
        start: 0x00038D,
        end: 0x00038D,
    },
    Range {
        start: 0x0003A2,
        end: 0x0003A2,
    },
    Range {
        start: 0x000530,
        end: 0x000530,
    },
    Range {
        start: 0x000557,
        end: 0x000558,
    },
    Range {
        start: 0x00058B,
        end: 0x00058C,
    },
    Range {
        start: 0x000590,
        end: 0x000590,
    },
    Range {
        start: 0x0005C8,
        end: 0x0005CF,
    },
    Range {
        start: 0x0005EB,
        end: 0x0005EE,
    },
    Range {
        start: 0x0005F5,
        end: 0x0005FF,
    },
    Range {
        start: 0x00070E,
        end: 0x00070E,
    },
    Range {
        start: 0x00074B,
        end: 0x00074C,
    },
    Range {
        start: 0x0007B2,
        end: 0x0007BF,
    },
    Range {
        start: 0x0007FB,
        end: 0x0007FC,
    },
    Range {
        start: 0x00082E,
        end: 0x00082F,
    },
    Range {
        start: 0x00083F,
        end: 0x00083F,
    },
    Range {
        start: 0x00085C,
        end: 0x00085D,
    },
    Range {
        start: 0x00085F,
        end: 0x00085F,
    },
    Range {
        start: 0x00086B,
        end: 0x00086F,
    },
    Range {
        start: 0x000892,
        end: 0x000896,
    },
    Range {
        start: 0x000984,
        end: 0x000984,
    },
    Range {
        start: 0x00098D,
        end: 0x00098E,
    },
    Range {
        start: 0x000991,
        end: 0x000992,
    },
    Range {
        start: 0x0009A9,
        end: 0x0009A9,
    },
    Range {
        start: 0x0009B1,
        end: 0x0009B1,
    },
    Range {
        start: 0x0009B3,
        end: 0x0009B5,
    },
    Range {
        start: 0x0009BA,
        end: 0x0009BB,
    },
    Range {
        start: 0x0009C5,
        end: 0x0009C6,
    },
    Range {
        start: 0x0009C9,
        end: 0x0009CA,
    },
    Range {
        start: 0x0009CF,
        end: 0x0009D6,
    },
    Range {
        start: 0x0009D8,
        end: 0x0009DB,
    },
    Range {
        start: 0x0009DE,
        end: 0x0009DE,
    },
    Range {
        start: 0x0009E4,
        end: 0x0009E5,
    },
    Range {
        start: 0x0009FF,
        end: 0x000A00,
    },
    Range {
        start: 0x000A04,
        end: 0x000A04,
    },
    Range {
        start: 0x000A0B,
        end: 0x000A0E,
    },
    Range {
        start: 0x000A11,
        end: 0x000A12,
    },
    Range {
        start: 0x000A29,
        end: 0x000A29,
    },
    Range {
        start: 0x000A31,
        end: 0x000A31,
    },
    Range {
        start: 0x000A34,
        end: 0x000A34,
    },
    Range {
        start: 0x000A37,
        end: 0x000A37,
    },
    Range {
        start: 0x000A3A,
        end: 0x000A3B,
    },
    Range {
        start: 0x000A3D,
        end: 0x000A3D,
    },
    Range {
        start: 0x000A43,
        end: 0x000A46,
    },
    Range {
        start: 0x000A49,
        end: 0x000A4A,
    },
    Range {
        start: 0x000A4E,
        end: 0x000A50,
    },
    Range {
        start: 0x000A52,
        end: 0x000A58,
    },
    Range {
        start: 0x000A5D,
        end: 0x000A5D,
    },
    Range {
        start: 0x000A5F,
        end: 0x000A65,
    },
    Range {
        start: 0x000A77,
        end: 0x000A80,
    },
    Range {
        start: 0x000A84,
        end: 0x000A84,
    },
    Range {
        start: 0x000A8E,
        end: 0x000A8E,
    },
    Range {
        start: 0x000A92,
        end: 0x000A92,
    },
    Range {
        start: 0x000AA9,
        end: 0x000AA9,
    },
    Range {
        start: 0x000AB1,
        end: 0x000AB1,
    },
    Range {
        start: 0x000AB4,
        end: 0x000AB4,
    },
    Range {
        start: 0x000ABA,
        end: 0x000ABB,
    },
    Range {
        start: 0x000AC6,
        end: 0x000AC6,
    },
    Range {
        start: 0x000ACA,
        end: 0x000ACA,
    },
    Range {
        start: 0x000ACE,
        end: 0x000ACF,
    },
    Range {
        start: 0x000AD1,
        end: 0x000ADF,
    },
    Range {
        start: 0x000AE4,
        end: 0x000AE5,
    },
    Range {
        start: 0x000AF2,
        end: 0x000AF8,
    },
    Range {
        start: 0x000B00,
        end: 0x000B00,
    },
    Range {
        start: 0x000B04,
        end: 0x000B04,
    },
    Range {
        start: 0x000B0D,
        end: 0x000B0E,
    },
    Range {
        start: 0x000B11,
        end: 0x000B12,
    },
    Range {
        start: 0x000B29,
        end: 0x000B29,
    },
    Range {
        start: 0x000B31,
        end: 0x000B31,
    },
    Range {
        start: 0x000B34,
        end: 0x000B34,
    },
    Range {
        start: 0x000B3A,
        end: 0x000B3B,
    },
    Range {
        start: 0x000B45,
        end: 0x000B46,
    },
    Range {
        start: 0x000B49,
        end: 0x000B4A,
    },
    Range {
        start: 0x000B4E,
        end: 0x000B54,
    },
    Range {
        start: 0x000B58,
        end: 0x000B5B,
    },
    Range {
        start: 0x000B5E,
        end: 0x000B5E,
    },
    Range {
        start: 0x000B64,
        end: 0x000B65,
    },
    Range {
        start: 0x000B78,
        end: 0x000B81,
    },
    Range {
        start: 0x000B84,
        end: 0x000B84,
    },
    Range {
        start: 0x000B8B,
        end: 0x000B8D,
    },
    Range {
        start: 0x000B91,
        end: 0x000B91,
    },
    Range {
        start: 0x000B96,
        end: 0x000B98,
    },
    Range {
        start: 0x000B9B,
        end: 0x000B9B,
    },
    Range {
        start: 0x000B9D,
        end: 0x000B9D,
    },
    Range {
        start: 0x000BA0,
        end: 0x000BA2,
    },
    Range {
        start: 0x000BA5,
        end: 0x000BA7,
    },
    Range {
        start: 0x000BAB,
        end: 0x000BAD,
    },
    Range {
        start: 0x000BBA,
        end: 0x000BBD,
    },
    Range {
        start: 0x000BC3,
        end: 0x000BC5,
    },
    Range {
        start: 0x000BC9,
        end: 0x000BC9,
    },
    Range {
        start: 0x000BCE,
        end: 0x000BCF,
    },
    Range {
        start: 0x000BD1,
        end: 0x000BD6,
    },
    Range {
        start: 0x000BD8,
        end: 0x000BE5,
    },
    Range {
        start: 0x000BFB,
        end: 0x000BFF,
    },
    Range {
        start: 0x000C0D,
        end: 0x000C0D,
    },
    Range {
        start: 0x000C11,
        end: 0x000C11,
    },
    Range {
        start: 0x000C29,
        end: 0x000C29,
    },
    Range {
        start: 0x000C3A,
        end: 0x000C3B,
    },
    Range {
        start: 0x000C45,
        end: 0x000C45,
    },
    Range {
        start: 0x000C49,
        end: 0x000C49,
    },
    Range {
        start: 0x000C4E,
        end: 0x000C54,
    },
    Range {
        start: 0x000C57,
        end: 0x000C57,
    },
    Range {
        start: 0x000C5B,
        end: 0x000C5B,
    },
    Range {
        start: 0x000C5E,
        end: 0x000C5F,
    },
    Range {
        start: 0x000C64,
        end: 0x000C65,
    },
    Range {
        start: 0x000C70,
        end: 0x000C76,
    },
    Range {
        start: 0x000C8D,
        end: 0x000C8D,
    },
    Range {
        start: 0x000C91,
        end: 0x000C91,
    },
    Range {
        start: 0x000CA9,
        end: 0x000CA9,
    },
    Range {
        start: 0x000CB4,
        end: 0x000CB4,
    },
    Range {
        start: 0x000CBA,
        end: 0x000CBB,
    },
    Range {
        start: 0x000CC5,
        end: 0x000CC5,
    },
    Range {
        start: 0x000CC9,
        end: 0x000CC9,
    },
    Range {
        start: 0x000CCE,
        end: 0x000CD4,
    },
    Range {
        start: 0x000CD7,
        end: 0x000CDB,
    },
    Range {
        start: 0x000CDF,
        end: 0x000CDF,
    },
    Range {
        start: 0x000CE4,
        end: 0x000CE5,
    },
    Range {
        start: 0x000CF0,
        end: 0x000CF0,
    },
    Range {
        start: 0x000CF4,
        end: 0x000CFF,
    },
    Range {
        start: 0x000D0D,
        end: 0x000D0D,
    },
    Range {
        start: 0x000D11,
        end: 0x000D11,
    },
    Range {
        start: 0x000D45,
        end: 0x000D45,
    },
    Range {
        start: 0x000D49,
        end: 0x000D49,
    },
    Range {
        start: 0x000D50,
        end: 0x000D53,
    },
    Range {
        start: 0x000D64,
        end: 0x000D65,
    },
    Range {
        start: 0x000D80,
        end: 0x000D80,
    },
    Range {
        start: 0x000D84,
        end: 0x000D84,
    },
    Range {
        start: 0x000D97,
        end: 0x000D99,
    },
    Range {
        start: 0x000DB2,
        end: 0x000DB2,
    },
    Range {
        start: 0x000DBC,
        end: 0x000DBC,
    },
    Range {
        start: 0x000DBE,
        end: 0x000DBF,
    },
    Range {
        start: 0x000DC7,
        end: 0x000DC9,
    },
    Range {
        start: 0x000DCB,
        end: 0x000DCE,
    },
    Range {
        start: 0x000DD5,
        end: 0x000DD5,
    },
    Range {
        start: 0x000DD7,
        end: 0x000DD7,
    },
    Range {
        start: 0x000DE0,
        end: 0x000DE5,
    },
    Range {
        start: 0x000DF0,
        end: 0x000DF1,
    },
    Range {
        start: 0x000DF5,
        end: 0x000E00,
    },
    Range {
        start: 0x000E3B,
        end: 0x000E3E,
    },
    Range {
        start: 0x000E5C,
        end: 0x000E80,
    },
    Range {
        start: 0x000E83,
        end: 0x000E83,
    },
    Range {
        start: 0x000E85,
        end: 0x000E85,
    },
    Range {
        start: 0x000E8B,
        end: 0x000E8B,
    },
    Range {
        start: 0x000EA4,
        end: 0x000EA4,
    },
    Range {
        start: 0x000EA6,
        end: 0x000EA6,
    },
    Range {
        start: 0x000EBE,
        end: 0x000EBF,
    },
    Range {
        start: 0x000EC5,
        end: 0x000EC5,
    },
    Range {
        start: 0x000EC7,
        end: 0x000EC7,
    },
    Range {
        start: 0x000ECF,
        end: 0x000ECF,
    },
    Range {
        start: 0x000EDA,
        end: 0x000EDB,
    },
    Range {
        start: 0x000EE0,
        end: 0x000EFF,
    },
    Range {
        start: 0x000F48,
        end: 0x000F48,
    },
    Range {
        start: 0x000F6D,
        end: 0x000F70,
    },
    Range {
        start: 0x000F98,
        end: 0x000F98,
    },
    Range {
        start: 0x000FBD,
        end: 0x000FBD,
    },
    Range {
        start: 0x000FCD,
        end: 0x000FCD,
    },
    Range {
        start: 0x000FDB,
        end: 0x000FFF,
    },
    Range {
        start: 0x0010C6,
        end: 0x0010C6,
    },
    Range {
        start: 0x0010C8,
        end: 0x0010CC,
    },
    Range {
        start: 0x0010CE,
        end: 0x0010CF,
    },
    Range {
        start: 0x001249,
        end: 0x001249,
    },
    Range {
        start: 0x00124E,
        end: 0x00124F,
    },
    Range {
        start: 0x001257,
        end: 0x001257,
    },
    Range {
        start: 0x001259,
        end: 0x001259,
    },
    Range {
        start: 0x00125E,
        end: 0x00125F,
    },
    Range {
        start: 0x001289,
        end: 0x001289,
    },
    Range {
        start: 0x00128E,
        end: 0x00128F,
    },
    Range {
        start: 0x0012B1,
        end: 0x0012B1,
    },
    Range {
        start: 0x0012B6,
        end: 0x0012B7,
    },
    Range {
        start: 0x0012BF,
        end: 0x0012BF,
    },
    Range {
        start: 0x0012C1,
        end: 0x0012C1,
    },
    Range {
        start: 0x0012C6,
        end: 0x0012C7,
    },
    Range {
        start: 0x0012D7,
        end: 0x0012D7,
    },
    Range {
        start: 0x001311,
        end: 0x001311,
    },
    Range {
        start: 0x001316,
        end: 0x001317,
    },
    Range {
        start: 0x00135B,
        end: 0x00135C,
    },
    Range {
        start: 0x00137D,
        end: 0x00137F,
    },
    Range {
        start: 0x00139A,
        end: 0x00139F,
    },
    Range {
        start: 0x0013F6,
        end: 0x0013F7,
    },
    Range {
        start: 0x0013FE,
        end: 0x0013FF,
    },
    Range {
        start: 0x00169D,
        end: 0x00169F,
    },
    Range {
        start: 0x0016F9,
        end: 0x0016FF,
    },
    Range {
        start: 0x001716,
        end: 0x00171E,
    },
    Range {
        start: 0x001737,
        end: 0x00173F,
    },
    Range {
        start: 0x001754,
        end: 0x00175F,
    },
    Range {
        start: 0x00176D,
        end: 0x00176D,
    },
    Range {
        start: 0x001771,
        end: 0x001771,
    },
    Range {
        start: 0x001774,
        end: 0x00177F,
    },
    Range {
        start: 0x0017DE,
        end: 0x0017DF,
    },
    Range {
        start: 0x0017EA,
        end: 0x0017EF,
    },
    Range {
        start: 0x0017FA,
        end: 0x0017FF,
    },
    Range {
        start: 0x00181A,
        end: 0x00181F,
    },
    Range {
        start: 0x001879,
        end: 0x00187F,
    },
    Range {
        start: 0x0018AB,
        end: 0x0018AF,
    },
    Range {
        start: 0x0018F6,
        end: 0x0018FF,
    },
    Range {
        start: 0x00191F,
        end: 0x00191F,
    },
    Range {
        start: 0x00192C,
        end: 0x00192F,
    },
    Range {
        start: 0x00193C,
        end: 0x00193F,
    },
    Range {
        start: 0x001941,
        end: 0x001943,
    },
    Range {
        start: 0x00196E,
        end: 0x00196F,
    },
    Range {
        start: 0x001975,
        end: 0x00197F,
    },
    Range {
        start: 0x0019AC,
        end: 0x0019AF,
    },
    Range {
        start: 0x0019CA,
        end: 0x0019CF,
    },
    Range {
        start: 0x0019DB,
        end: 0x0019DD,
    },
    Range {
        start: 0x001A1C,
        end: 0x001A1D,
    },
    Range {
        start: 0x001A5F,
        end: 0x001A5F,
    },
    Range {
        start: 0x001A7D,
        end: 0x001A7E,
    },
    Range {
        start: 0x001A8A,
        end: 0x001A8F,
    },
    Range {
        start: 0x001A9A,
        end: 0x001A9F,
    },
    Range {
        start: 0x001AAE,
        end: 0x001AAF,
    },
    Range {
        start: 0x001ADE,
        end: 0x001ADF,
    },
    Range {
        start: 0x001AEC,
        end: 0x001AFF,
    },
    Range {
        start: 0x001B4D,
        end: 0x001B4D,
    },
    Range {
        start: 0x001BF4,
        end: 0x001BFB,
    },
    Range {
        start: 0x001C38,
        end: 0x001C3A,
    },
    Range {
        start: 0x001C4A,
        end: 0x001C4C,
    },
    Range {
        start: 0x001C8B,
        end: 0x001C8F,
    },
    Range {
        start: 0x001CBB,
        end: 0x001CBC,
    },
    Range {
        start: 0x001CC8,
        end: 0x001CCF,
    },
    Range {
        start: 0x001CFB,
        end: 0x001CFF,
    },
    Range {
        start: 0x001F16,
        end: 0x001F17,
    },
    Range {
        start: 0x001F1E,
        end: 0x001F1F,
    },
    Range {
        start: 0x001F46,
        end: 0x001F47,
    },
    Range {
        start: 0x001F4E,
        end: 0x001F4F,
    },
    Range {
        start: 0x001F58,
        end: 0x001F58,
    },
    Range {
        start: 0x001F5A,
        end: 0x001F5A,
    },
    Range {
        start: 0x001F5C,
        end: 0x001F5C,
    },
    Range {
        start: 0x001F5E,
        end: 0x001F5E,
    },
    Range {
        start: 0x001F7E,
        end: 0x001F7F,
    },
    Range {
        start: 0x001FB5,
        end: 0x001FB5,
    },
    Range {
        start: 0x001FC5,
        end: 0x001FC5,
    },
    Range {
        start: 0x001FD4,
        end: 0x001FD5,
    },
    Range {
        start: 0x001FDC,
        end: 0x001FDC,
    },
    Range {
        start: 0x001FF0,
        end: 0x001FF1,
    },
    Range {
        start: 0x001FF5,
        end: 0x001FF5,
    },
    Range {
        start: 0x001FFF,
        end: 0x001FFF,
    },
    Range {
        start: 0x002065,
        end: 0x002065,
    },
    Range {
        start: 0x002072,
        end: 0x002073,
    },
    Range {
        start: 0x00208F,
        end: 0x00208F,
    },
    Range {
        start: 0x00209D,
        end: 0x00209F,
    },
    Range {
        start: 0x0020C2,
        end: 0x0020CF,
    },
    Range {
        start: 0x0020F1,
        end: 0x0020FF,
    },
    Range {
        start: 0x00218C,
        end: 0x00218F,
    },
    Range {
        start: 0x00242A,
        end: 0x00243F,
    },
    Range {
        start: 0x00244B,
        end: 0x00245F,
    },
    Range {
        start: 0x002B74,
        end: 0x002B75,
    },
    Range {
        start: 0x002CF4,
        end: 0x002CF8,
    },
    Range {
        start: 0x002D26,
        end: 0x002D26,
    },
    Range {
        start: 0x002D28,
        end: 0x002D2C,
    },
    Range {
        start: 0x002D2E,
        end: 0x002D2F,
    },
    Range {
        start: 0x002D68,
        end: 0x002D6E,
    },
    Range {
        start: 0x002D71,
        end: 0x002D7E,
    },
    Range {
        start: 0x002D97,
        end: 0x002D9F,
    },
    Range {
        start: 0x002DA7,
        end: 0x002DA7,
    },
    Range {
        start: 0x002DAF,
        end: 0x002DAF,
    },
    Range {
        start: 0x002DB7,
        end: 0x002DB7,
    },
    Range {
        start: 0x002DBF,
        end: 0x002DBF,
    },
    Range {
        start: 0x002DC7,
        end: 0x002DC7,
    },
    Range {
        start: 0x002DCF,
        end: 0x002DCF,
    },
    Range {
        start: 0x002DD7,
        end: 0x002DD7,
    },
    Range {
        start: 0x002DDF,
        end: 0x002DDF,
    },
    Range {
        start: 0x002E5E,
        end: 0x002E7F,
    },
    Range {
        start: 0x002E9A,
        end: 0x002E9A,
    },
    Range {
        start: 0x002EF4,
        end: 0x002EFF,
    },
    Range {
        start: 0x002FD6,
        end: 0x002FEF,
    },
    Range {
        start: 0x003040,
        end: 0x003040,
    },
    Range {
        start: 0x003097,
        end: 0x003098,
    },
    Range {
        start: 0x003100,
        end: 0x003104,
    },
    Range {
        start: 0x003130,
        end: 0x003130,
    },
    Range {
        start: 0x00318F,
        end: 0x00318F,
    },
    Range {
        start: 0x0031E6,
        end: 0x0031EE,
    },
    Range {
        start: 0x00321F,
        end: 0x00321F,
    },
    Range {
        start: 0x00A48D,
        end: 0x00A48F,
    },
    Range {
        start: 0x00A4C7,
        end: 0x00A4CF,
    },
    Range {
        start: 0x00A62C,
        end: 0x00A63F,
    },
    Range {
        start: 0x00A6F8,
        end: 0x00A6FF,
    },
    Range {
        start: 0x00A7DD,
        end: 0x00A7F0,
    },
    Range {
        start: 0x00A82D,
        end: 0x00A82F,
    },
    Range {
        start: 0x00A83A,
        end: 0x00A83F,
    },
    Range {
        start: 0x00A878,
        end: 0x00A87F,
    },
    Range {
        start: 0x00A8C6,
        end: 0x00A8CD,
    },
    Range {
        start: 0x00A8DA,
        end: 0x00A8DF,
    },
    Range {
        start: 0x00A954,
        end: 0x00A95E,
    },
    Range {
        start: 0x00A97D,
        end: 0x00A97F,
    },
    Range {
        start: 0x00A9CE,
        end: 0x00A9CE,
    },
    Range {
        start: 0x00A9DA,
        end: 0x00A9DD,
    },
    Range {
        start: 0x00A9FF,
        end: 0x00A9FF,
    },
    Range {
        start: 0x00AA37,
        end: 0x00AA3F,
    },
    Range {
        start: 0x00AA4E,
        end: 0x00AA4F,
    },
    Range {
        start: 0x00AA5A,
        end: 0x00AA5B,
    },
    Range {
        start: 0x00AAC3,
        end: 0x00AADA,
    },
    Range {
        start: 0x00AAF7,
        end: 0x00AB00,
    },
    Range {
        start: 0x00AB07,
        end: 0x00AB08,
    },
    Range {
        start: 0x00AB0F,
        end: 0x00AB10,
    },
    Range {
        start: 0x00AB17,
        end: 0x00AB1F,
    },
    Range {
        start: 0x00AB27,
        end: 0x00AB27,
    },
    Range {
        start: 0x00AB2F,
        end: 0x00AB2F,
    },
    Range {
        start: 0x00AB6C,
        end: 0x00AB6F,
    },
    Range {
        start: 0x00ABEE,
        end: 0x00ABEF,
    },
    Range {
        start: 0x00ABFA,
        end: 0x00ABFF,
    },
    Range {
        start: 0x00D7A4,
        end: 0x00D7AF,
    },
    Range {
        start: 0x00D7C7,
        end: 0x00D7CA,
    },
    Range {
        start: 0x00D7FC,
        end: 0x00D7FF,
    },
    Range {
        start: 0x00FA6E,
        end: 0x00FA6F,
    },
    Range {
        start: 0x00FADA,
        end: 0x00FAFF,
    },
    Range {
        start: 0x00FB07,
        end: 0x00FB12,
    },
    Range {
        start: 0x00FB18,
        end: 0x00FB1C,
    },
    Range {
        start: 0x00FB37,
        end: 0x00FB37,
    },
    Range {
        start: 0x00FB3D,
        end: 0x00FB3D,
    },
    Range {
        start: 0x00FB3F,
        end: 0x00FB3F,
    },
    Range {
        start: 0x00FB42,
        end: 0x00FB42,
    },
    Range {
        start: 0x00FB45,
        end: 0x00FB45,
    },
    Range {
        start: 0x00FDD0,
        end: 0x00FDEF,
    },
    Range {
        start: 0x00FE1A,
        end: 0x00FE1F,
    },
    Range {
        start: 0x00FE53,
        end: 0x00FE53,
    },
    Range {
        start: 0x00FE67,
        end: 0x00FE67,
    },
    Range {
        start: 0x00FE6C,
        end: 0x00FE6F,
    },
    Range {
        start: 0x00FE75,
        end: 0x00FE75,
    },
    Range {
        start: 0x00FEFD,
        end: 0x00FEFE,
    },
    Range {
        start: 0x00FF00,
        end: 0x00FF00,
    },
    Range {
        start: 0x00FFBF,
        end: 0x00FFC1,
    },
    Range {
        start: 0x00FFC8,
        end: 0x00FFC9,
    },
    Range {
        start: 0x00FFD0,
        end: 0x00FFD1,
    },
    Range {
        start: 0x00FFD8,
        end: 0x00FFD9,
    },
    Range {
        start: 0x00FFDD,
        end: 0x00FFDF,
    },
    Range {
        start: 0x00FFE7,
        end: 0x00FFE7,
    },
    Range {
        start: 0x00FFEF,
        end: 0x00FFF8,
    },
    Range {
        start: 0x00FFFE,
        end: 0x00FFFF,
    },
    Range {
        start: 0x01000C,
        end: 0x01000C,
    },
    Range {
        start: 0x010027,
        end: 0x010027,
    },
    Range {
        start: 0x01003B,
        end: 0x01003B,
    },
    Range {
        start: 0x01003E,
        end: 0x01003E,
    },
    Range {
        start: 0x01004E,
        end: 0x01004F,
    },
    Range {
        start: 0x01005E,
        end: 0x01007F,
    },
    Range {
        start: 0x0100FB,
        end: 0x0100FF,
    },
    Range {
        start: 0x010103,
        end: 0x010106,
    },
    Range {
        start: 0x010134,
        end: 0x010136,
    },
    Range {
        start: 0x01018F,
        end: 0x01018F,
    },
    Range {
        start: 0x01019D,
        end: 0x01019F,
    },
    Range {
        start: 0x0101A1,
        end: 0x0101CF,
    },
    Range {
        start: 0x0101FE,
        end: 0x01027F,
    },
    Range {
        start: 0x01029D,
        end: 0x01029F,
    },
    Range {
        start: 0x0102D1,
        end: 0x0102DF,
    },
    Range {
        start: 0x0102FC,
        end: 0x0102FF,
    },
    Range {
        start: 0x010324,
        end: 0x01032C,
    },
    Range {
        start: 0x01034B,
        end: 0x01034F,
    },
    Range {
        start: 0x01037B,
        end: 0x01037F,
    },
    Range {
        start: 0x01039E,
        end: 0x01039E,
    },
    Range {
        start: 0x0103C4,
        end: 0x0103C7,
    },
    Range {
        start: 0x0103D6,
        end: 0x0103FF,
    },
    Range {
        start: 0x01049E,
        end: 0x01049F,
    },
    Range {
        start: 0x0104AA,
        end: 0x0104AF,
    },
    Range {
        start: 0x0104D4,
        end: 0x0104D7,
    },
    Range {
        start: 0x0104FC,
        end: 0x0104FF,
    },
    Range {
        start: 0x010528,
        end: 0x01052F,
    },
    Range {
        start: 0x010564,
        end: 0x01056E,
    },
    Range {
        start: 0x01057B,
        end: 0x01057B,
    },
    Range {
        start: 0x01058B,
        end: 0x01058B,
    },
    Range {
        start: 0x010593,
        end: 0x010593,
    },
    Range {
        start: 0x010596,
        end: 0x010596,
    },
    Range {
        start: 0x0105A2,
        end: 0x0105A2,
    },
    Range {
        start: 0x0105B2,
        end: 0x0105B2,
    },
    Range {
        start: 0x0105BA,
        end: 0x0105BA,
    },
    Range {
        start: 0x0105BD,
        end: 0x0105BF,
    },
    Range {
        start: 0x0105F4,
        end: 0x0105FF,
    },
    Range {
        start: 0x010737,
        end: 0x01073F,
    },
    Range {
        start: 0x010756,
        end: 0x01075F,
    },
    Range {
        start: 0x010768,
        end: 0x01077F,
    },
    Range {
        start: 0x010786,
        end: 0x010786,
    },
    Range {
        start: 0x0107B1,
        end: 0x0107B1,
    },
    Range {
        start: 0x0107BB,
        end: 0x0107FF,
    },
    Range {
        start: 0x010806,
        end: 0x010807,
    },
    Range {
        start: 0x010809,
        end: 0x010809,
    },
    Range {
        start: 0x010836,
        end: 0x010836,
    },
    Range {
        start: 0x010839,
        end: 0x01083B,
    },
    Range {
        start: 0x01083D,
        end: 0x01083E,
    },
    Range {
        start: 0x010856,
        end: 0x010856,
    },
    Range {
        start: 0x01089F,
        end: 0x0108A6,
    },
    Range {
        start: 0x0108B0,
        end: 0x0108DF,
    },
    Range {
        start: 0x0108F3,
        end: 0x0108F3,
    },
    Range {
        start: 0x0108F6,
        end: 0x0108FA,
    },
    Range {
        start: 0x01091C,
        end: 0x01091E,
    },
    Range {
        start: 0x01093A,
        end: 0x01093E,
    },
    Range {
        start: 0x01095A,
        end: 0x01097F,
    },
    Range {
        start: 0x0109B8,
        end: 0x0109BB,
    },
    Range {
        start: 0x0109D0,
        end: 0x0109D1,
    },
    Range {
        start: 0x010A04,
        end: 0x010A04,
    },
    Range {
        start: 0x010A07,
        end: 0x010A0B,
    },
    Range {
        start: 0x010A14,
        end: 0x010A14,
    },
    Range {
        start: 0x010A18,
        end: 0x010A18,
    },
    Range {
        start: 0x010A36,
        end: 0x010A37,
    },
    Range {
        start: 0x010A3B,
        end: 0x010A3E,
    },
    Range {
        start: 0x010A49,
        end: 0x010A4F,
    },
    Range {
        start: 0x010A59,
        end: 0x010A5F,
    },
    Range {
        start: 0x010AA0,
        end: 0x010ABF,
    },
    Range {
        start: 0x010AE7,
        end: 0x010AEA,
    },
    Range {
        start: 0x010AF7,
        end: 0x010AFF,
    },
    Range {
        start: 0x010B36,
        end: 0x010B38,
    },
    Range {
        start: 0x010B56,
        end: 0x010B57,
    },
    Range {
        start: 0x010B73,
        end: 0x010B77,
    },
    Range {
        start: 0x010B92,
        end: 0x010B98,
    },
    Range {
        start: 0x010B9D,
        end: 0x010BA8,
    },
    Range {
        start: 0x010BB0,
        end: 0x010BFF,
    },
    Range {
        start: 0x010C49,
        end: 0x010C7F,
    },
    Range {
        start: 0x010CB3,
        end: 0x010CBF,
    },
    Range {
        start: 0x010CF3,
        end: 0x010CF9,
    },
    Range {
        start: 0x010D28,
        end: 0x010D2F,
    },
    Range {
        start: 0x010D3A,
        end: 0x010D3F,
    },
    Range {
        start: 0x010D66,
        end: 0x010D68,
    },
    Range {
        start: 0x010D86,
        end: 0x010D8D,
    },
    Range {
        start: 0x010D90,
        end: 0x010E5F,
    },
    Range {
        start: 0x010E7F,
        end: 0x010E7F,
    },
    Range {
        start: 0x010EAA,
        end: 0x010EAA,
    },
    Range {
        start: 0x010EAE,
        end: 0x010EAF,
    },
    Range {
        start: 0x010EB2,
        end: 0x010EC1,
    },
    Range {
        start: 0x010EC8,
        end: 0x010ECF,
    },
    Range {
        start: 0x010ED9,
        end: 0x010EF9,
    },
    Range {
        start: 0x010F28,
        end: 0x010F2F,
    },
    Range {
        start: 0x010F5A,
        end: 0x010F6F,
    },
    Range {
        start: 0x010F8A,
        end: 0x010FAF,
    },
    Range {
        start: 0x010FCC,
        end: 0x010FDF,
    },
    Range {
        start: 0x010FF7,
        end: 0x010FFF,
    },
    Range {
        start: 0x01104E,
        end: 0x011051,
    },
    Range {
        start: 0x011076,
        end: 0x01107E,
    },
    Range {
        start: 0x0110C3,
        end: 0x0110CC,
    },
    Range {
        start: 0x0110CE,
        end: 0x0110CF,
    },
    Range {
        start: 0x0110E9,
        end: 0x0110EF,
    },
    Range {
        start: 0x0110FA,
        end: 0x0110FF,
    },
    Range {
        start: 0x011135,
        end: 0x011135,
    },
    Range {
        start: 0x011148,
        end: 0x01114F,
    },
    Range {
        start: 0x011177,
        end: 0x01117F,
    },
    Range {
        start: 0x0111E0,
        end: 0x0111E0,
    },
    Range {
        start: 0x0111F5,
        end: 0x0111FF,
    },
    Range {
        start: 0x011212,
        end: 0x011212,
    },
    Range {
        start: 0x011242,
        end: 0x01127F,
    },
    Range {
        start: 0x011287,
        end: 0x011287,
    },
    Range {
        start: 0x011289,
        end: 0x011289,
    },
    Range {
        start: 0x01128E,
        end: 0x01128E,
    },
    Range {
        start: 0x01129E,
        end: 0x01129E,
    },
    Range {
        start: 0x0112AA,
        end: 0x0112AF,
    },
    Range {
        start: 0x0112EB,
        end: 0x0112EF,
    },
    Range {
        start: 0x0112FA,
        end: 0x0112FF,
    },
    Range {
        start: 0x011304,
        end: 0x011304,
    },
    Range {
        start: 0x01130D,
        end: 0x01130E,
    },
    Range {
        start: 0x011311,
        end: 0x011312,
    },
    Range {
        start: 0x011329,
        end: 0x011329,
    },
    Range {
        start: 0x011331,
        end: 0x011331,
    },
    Range {
        start: 0x011334,
        end: 0x011334,
    },
    Range {
        start: 0x01133A,
        end: 0x01133A,
    },
    Range {
        start: 0x011345,
        end: 0x011346,
    },
    Range {
        start: 0x011349,
        end: 0x01134A,
    },
    Range {
        start: 0x01134E,
        end: 0x01134F,
    },
    Range {
        start: 0x011351,
        end: 0x011356,
    },
    Range {
        start: 0x011358,
        end: 0x01135C,
    },
    Range {
        start: 0x011364,
        end: 0x011365,
    },
    Range {
        start: 0x01136D,
        end: 0x01136F,
    },
    Range {
        start: 0x011375,
        end: 0x01137F,
    },
    Range {
        start: 0x01138A,
        end: 0x01138A,
    },
    Range {
        start: 0x01138C,
        end: 0x01138D,
    },
    Range {
        start: 0x01138F,
        end: 0x01138F,
    },
    Range {
        start: 0x0113B6,
        end: 0x0113B6,
    },
    Range {
        start: 0x0113C1,
        end: 0x0113C1,
    },
    Range {
        start: 0x0113C3,
        end: 0x0113C4,
    },
    Range {
        start: 0x0113C6,
        end: 0x0113C6,
    },
    Range {
        start: 0x0113CB,
        end: 0x0113CB,
    },
    Range {
        start: 0x0113D6,
        end: 0x0113D6,
    },
    Range {
        start: 0x0113D9,
        end: 0x0113E0,
    },
    Range {
        start: 0x0113E3,
        end: 0x0113FF,
    },
    Range {
        start: 0x01145C,
        end: 0x01145C,
    },
    Range {
        start: 0x011462,
        end: 0x01147F,
    },
    Range {
        start: 0x0114C8,
        end: 0x0114CF,
    },
    Range {
        start: 0x0114DA,
        end: 0x01157F,
    },
    Range {
        start: 0x0115B6,
        end: 0x0115B7,
    },
    Range {
        start: 0x0115DE,
        end: 0x0115FF,
    },
    Range {
        start: 0x011645,
        end: 0x01164F,
    },
    Range {
        start: 0x01165A,
        end: 0x01165F,
    },
    Range {
        start: 0x01166D,
        end: 0x01167F,
    },
    Range {
        start: 0x0116BA,
        end: 0x0116BF,
    },
    Range {
        start: 0x0116CA,
        end: 0x0116CF,
    },
    Range {
        start: 0x0116E4,
        end: 0x0116FF,
    },
    Range {
        start: 0x01171B,
        end: 0x01171C,
    },
    Range {
        start: 0x01172C,
        end: 0x01172F,
    },
    Range {
        start: 0x011747,
        end: 0x0117FF,
    },
    Range {
        start: 0x01183C,
        end: 0x01189F,
    },
    Range {
        start: 0x0118F3,
        end: 0x0118FE,
    },
    Range {
        start: 0x011907,
        end: 0x011908,
    },
    Range {
        start: 0x01190A,
        end: 0x01190B,
    },
    Range {
        start: 0x011914,
        end: 0x011914,
    },
    Range {
        start: 0x011917,
        end: 0x011917,
    },
    Range {
        start: 0x011936,
        end: 0x011936,
    },
    Range {
        start: 0x011939,
        end: 0x01193A,
    },
    Range {
        start: 0x011947,
        end: 0x01194F,
    },
    Range {
        start: 0x01195A,
        end: 0x01199F,
    },
    Range {
        start: 0x0119A8,
        end: 0x0119A9,
    },
    Range {
        start: 0x0119D8,
        end: 0x0119D9,
    },
    Range {
        start: 0x0119E5,
        end: 0x0119FF,
    },
    Range {
        start: 0x011A48,
        end: 0x011A4F,
    },
    Range {
        start: 0x011AA3,
        end: 0x011AAF,
    },
    Range {
        start: 0x011AF9,
        end: 0x011AFF,
    },
    Range {
        start: 0x011B0A,
        end: 0x011B5F,
    },
    Range {
        start: 0x011B68,
        end: 0x011BBF,
    },
    Range {
        start: 0x011BE2,
        end: 0x011BEF,
    },
    Range {
        start: 0x011BFA,
        end: 0x011BFF,
    },
    Range {
        start: 0x011C09,
        end: 0x011C09,
    },
    Range {
        start: 0x011C37,
        end: 0x011C37,
    },
    Range {
        start: 0x011C46,
        end: 0x011C4F,
    },
    Range {
        start: 0x011C6D,
        end: 0x011C6F,
    },
    Range {
        start: 0x011C90,
        end: 0x011C91,
    },
    Range {
        start: 0x011CA8,
        end: 0x011CA8,
    },
    Range {
        start: 0x011CB7,
        end: 0x011CFF,
    },
    Range {
        start: 0x011D07,
        end: 0x011D07,
    },
    Range {
        start: 0x011D0A,
        end: 0x011D0A,
    },
    Range {
        start: 0x011D37,
        end: 0x011D39,
    },
    Range {
        start: 0x011D3B,
        end: 0x011D3B,
    },
    Range {
        start: 0x011D3E,
        end: 0x011D3E,
    },
    Range {
        start: 0x011D48,
        end: 0x011D4F,
    },
    Range {
        start: 0x011D5A,
        end: 0x011D5F,
    },
    Range {
        start: 0x011D66,
        end: 0x011D66,
    },
    Range {
        start: 0x011D69,
        end: 0x011D69,
    },
    Range {
        start: 0x011D8F,
        end: 0x011D8F,
    },
    Range {
        start: 0x011D92,
        end: 0x011D92,
    },
    Range {
        start: 0x011D99,
        end: 0x011D9F,
    },
    Range {
        start: 0x011DAA,
        end: 0x011DAF,
    },
    Range {
        start: 0x011DDC,
        end: 0x011DDF,
    },
    Range {
        start: 0x011DEA,
        end: 0x011EDF,
    },
    Range {
        start: 0x011EF9,
        end: 0x011EFF,
    },
    Range {
        start: 0x011F11,
        end: 0x011F11,
    },
    Range {
        start: 0x011F3B,
        end: 0x011F3D,
    },
    Range {
        start: 0x011F5B,
        end: 0x011FAF,
    },
    Range {
        start: 0x011FB1,
        end: 0x011FBF,
    },
    Range {
        start: 0x011FF2,
        end: 0x011FFE,
    },
    Range {
        start: 0x01239A,
        end: 0x0123FF,
    },
    Range {
        start: 0x01246F,
        end: 0x01246F,
    },
    Range {
        start: 0x012475,
        end: 0x01247F,
    },
    Range {
        start: 0x012544,
        end: 0x012F8F,
    },
    Range {
        start: 0x012FF3,
        end: 0x012FFF,
    },
    Range {
        start: 0x013456,
        end: 0x01345F,
    },
    Range {
        start: 0x0143FB,
        end: 0x0143FF,
    },
    Range {
        start: 0x014647,
        end: 0x0160FF,
    },
    Range {
        start: 0x01613A,
        end: 0x0167FF,
    },
    Range {
        start: 0x016A39,
        end: 0x016A3F,
    },
    Range {
        start: 0x016A5F,
        end: 0x016A5F,
    },
    Range {
        start: 0x016A6A,
        end: 0x016A6D,
    },
    Range {
        start: 0x016ABF,
        end: 0x016ABF,
    },
    Range {
        start: 0x016ACA,
        end: 0x016ACF,
    },
    Range {
        start: 0x016AEE,
        end: 0x016AEF,
    },
    Range {
        start: 0x016AF6,
        end: 0x016AFF,
    },
    Range {
        start: 0x016B46,
        end: 0x016B4F,
    },
    Range {
        start: 0x016B5A,
        end: 0x016B5A,
    },
    Range {
        start: 0x016B62,
        end: 0x016B62,
    },
    Range {
        start: 0x016B78,
        end: 0x016B7C,
    },
    Range {
        start: 0x016B90,
        end: 0x016D3F,
    },
    Range {
        start: 0x016D7A,
        end: 0x016E3F,
    },
    Range {
        start: 0x016E9B,
        end: 0x016E9F,
    },
    Range {
        start: 0x016EB9,
        end: 0x016EBA,
    },
    Range {
        start: 0x016ED4,
        end: 0x016EFF,
    },
    Range {
        start: 0x016F4B,
        end: 0x016F4E,
    },
    Range {
        start: 0x016F88,
        end: 0x016F8E,
    },
    Range {
        start: 0x016FA0,
        end: 0x016FDF,
    },
    Range {
        start: 0x016FE5,
        end: 0x016FEF,
    },
    Range {
        start: 0x016FF7,
        end: 0x016FFF,
    },
    Range {
        start: 0x018CD6,
        end: 0x018CFE,
    },
    Range {
        start: 0x018D1F,
        end: 0x018D7F,
    },
    Range {
        start: 0x018DF3,
        end: 0x01AFEF,
    },
    Range {
        start: 0x01AFF4,
        end: 0x01AFF4,
    },
    Range {
        start: 0x01AFFC,
        end: 0x01AFFC,
    },
    Range {
        start: 0x01AFFF,
        end: 0x01AFFF,
    },
    Range {
        start: 0x01B123,
        end: 0x01B131,
    },
    Range {
        start: 0x01B133,
        end: 0x01B14F,
    },
    Range {
        start: 0x01B153,
        end: 0x01B154,
    },
    Range {
        start: 0x01B156,
        end: 0x01B163,
    },
    Range {
        start: 0x01B168,
        end: 0x01B16F,
    },
    Range {
        start: 0x01B2FC,
        end: 0x01BBFF,
    },
    Range {
        start: 0x01BC6B,
        end: 0x01BC6F,
    },
    Range {
        start: 0x01BC7D,
        end: 0x01BC7F,
    },
    Range {
        start: 0x01BC89,
        end: 0x01BC8F,
    },
    Range {
        start: 0x01BC9A,
        end: 0x01BC9B,
    },
    Range {
        start: 0x01BCA4,
        end: 0x01CBFF,
    },
    Range {
        start: 0x01CCFD,
        end: 0x01CCFF,
    },
    Range {
        start: 0x01CEB4,
        end: 0x01CEB9,
    },
    Range {
        start: 0x01CED1,
        end: 0x01CEDF,
    },
    Range {
        start: 0x01CEF1,
        end: 0x01CEFF,
    },
    Range {
        start: 0x01CF2E,
        end: 0x01CF2F,
    },
    Range {
        start: 0x01CF47,
        end: 0x01CF4F,
    },
    Range {
        start: 0x01CFC4,
        end: 0x01CFFF,
    },
    Range {
        start: 0x01D0F6,
        end: 0x01D0FF,
    },
    Range {
        start: 0x01D127,
        end: 0x01D128,
    },
    Range {
        start: 0x01D1EB,
        end: 0x01D1FF,
    },
    Range {
        start: 0x01D246,
        end: 0x01D2BF,
    },
    Range {
        start: 0x01D2D4,
        end: 0x01D2DF,
    },
    Range {
        start: 0x01D2F4,
        end: 0x01D2FF,
    },
    Range {
        start: 0x01D357,
        end: 0x01D35F,
    },
    Range {
        start: 0x01D379,
        end: 0x01D3FF,
    },
    Range {
        start: 0x01D455,
        end: 0x01D455,
    },
    Range {
        start: 0x01D49D,
        end: 0x01D49D,
    },
    Range {
        start: 0x01D4A0,
        end: 0x01D4A1,
    },
    Range {
        start: 0x01D4A3,
        end: 0x01D4A4,
    },
    Range {
        start: 0x01D4A7,
        end: 0x01D4A8,
    },
    Range {
        start: 0x01D4AD,
        end: 0x01D4AD,
    },
    Range {
        start: 0x01D4BA,
        end: 0x01D4BA,
    },
    Range {
        start: 0x01D4BC,
        end: 0x01D4BC,
    },
    Range {
        start: 0x01D4C4,
        end: 0x01D4C4,
    },
    Range {
        start: 0x01D506,
        end: 0x01D506,
    },
    Range {
        start: 0x01D50B,
        end: 0x01D50C,
    },
    Range {
        start: 0x01D515,
        end: 0x01D515,
    },
    Range {
        start: 0x01D51D,
        end: 0x01D51D,
    },
    Range {
        start: 0x01D53A,
        end: 0x01D53A,
    },
    Range {
        start: 0x01D53F,
        end: 0x01D53F,
    },
    Range {
        start: 0x01D545,
        end: 0x01D545,
    },
    Range {
        start: 0x01D547,
        end: 0x01D549,
    },
    Range {
        start: 0x01D551,
        end: 0x01D551,
    },
    Range {
        start: 0x01D6A6,
        end: 0x01D6A7,
    },
    Range {
        start: 0x01D7CC,
        end: 0x01D7CD,
    },
    Range {
        start: 0x01DA8C,
        end: 0x01DA9A,
    },
    Range {
        start: 0x01DAA0,
        end: 0x01DAA0,
    },
    Range {
        start: 0x01DAB0,
        end: 0x01DEFF,
    },
    Range {
        start: 0x01DF1F,
        end: 0x01DF24,
    },
    Range {
        start: 0x01DF2B,
        end: 0x01DFFF,
    },
    Range {
        start: 0x01E007,
        end: 0x01E007,
    },
    Range {
        start: 0x01E019,
        end: 0x01E01A,
    },
    Range {
        start: 0x01E022,
        end: 0x01E022,
    },
    Range {
        start: 0x01E025,
        end: 0x01E025,
    },
    Range {
        start: 0x01E02B,
        end: 0x01E02F,
    },
    Range {
        start: 0x01E06E,
        end: 0x01E08E,
    },
    Range {
        start: 0x01E090,
        end: 0x01E0FF,
    },
    Range {
        start: 0x01E12D,
        end: 0x01E12F,
    },
    Range {
        start: 0x01E13E,
        end: 0x01E13F,
    },
    Range {
        start: 0x01E14A,
        end: 0x01E14D,
    },
    Range {
        start: 0x01E150,
        end: 0x01E28F,
    },
    Range {
        start: 0x01E2AF,
        end: 0x01E2BF,
    },
    Range {
        start: 0x01E2FA,
        end: 0x01E2FE,
    },
    Range {
        start: 0x01E300,
        end: 0x01E4CF,
    },
    Range {
        start: 0x01E4FA,
        end: 0x01E5CF,
    },
    Range {
        start: 0x01E5FB,
        end: 0x01E5FE,
    },
    Range {
        start: 0x01E600,
        end: 0x01E6BF,
    },
    Range {
        start: 0x01E6DF,
        end: 0x01E6DF,
    },
    Range {
        start: 0x01E6F6,
        end: 0x01E6FD,
    },
    Range {
        start: 0x01E700,
        end: 0x01E7DF,
    },
    Range {
        start: 0x01E7E7,
        end: 0x01E7E7,
    },
    Range {
        start: 0x01E7EC,
        end: 0x01E7EC,
    },
    Range {
        start: 0x01E7EF,
        end: 0x01E7EF,
    },
    Range {
        start: 0x01E7FF,
        end: 0x01E7FF,
    },
    Range {
        start: 0x01E8C5,
        end: 0x01E8C6,
    },
    Range {
        start: 0x01E8D7,
        end: 0x01E8FF,
    },
    Range {
        start: 0x01E94C,
        end: 0x01E94F,
    },
    Range {
        start: 0x01E95A,
        end: 0x01E95D,
    },
    Range {
        start: 0x01E960,
        end: 0x01EC70,
    },
    Range {
        start: 0x01ECB5,
        end: 0x01ED00,
    },
    Range {
        start: 0x01ED3E,
        end: 0x01EDFF,
    },
    Range {
        start: 0x01EE04,
        end: 0x01EE04,
    },
    Range {
        start: 0x01EE20,
        end: 0x01EE20,
    },
    Range {
        start: 0x01EE23,
        end: 0x01EE23,
    },
    Range {
        start: 0x01EE25,
        end: 0x01EE26,
    },
    Range {
        start: 0x01EE28,
        end: 0x01EE28,
    },
    Range {
        start: 0x01EE33,
        end: 0x01EE33,
    },
    Range {
        start: 0x01EE38,
        end: 0x01EE38,
    },
    Range {
        start: 0x01EE3A,
        end: 0x01EE3A,
    },
    Range {
        start: 0x01EE3C,
        end: 0x01EE41,
    },
    Range {
        start: 0x01EE43,
        end: 0x01EE46,
    },
    Range {
        start: 0x01EE48,
        end: 0x01EE48,
    },
    Range {
        start: 0x01EE4A,
        end: 0x01EE4A,
    },
    Range {
        start: 0x01EE4C,
        end: 0x01EE4C,
    },
    Range {
        start: 0x01EE50,
        end: 0x01EE50,
    },
    Range {
        start: 0x01EE53,
        end: 0x01EE53,
    },
    Range {
        start: 0x01EE55,
        end: 0x01EE56,
    },
    Range {
        start: 0x01EE58,
        end: 0x01EE58,
    },
    Range {
        start: 0x01EE5A,
        end: 0x01EE5A,
    },
    Range {
        start: 0x01EE5C,
        end: 0x01EE5C,
    },
    Range {
        start: 0x01EE5E,
        end: 0x01EE5E,
    },
    Range {
        start: 0x01EE60,
        end: 0x01EE60,
    },
    Range {
        start: 0x01EE63,
        end: 0x01EE63,
    },
    Range {
        start: 0x01EE65,
        end: 0x01EE66,
    },
    Range {
        start: 0x01EE6B,
        end: 0x01EE6B,
    },
    Range {
        start: 0x01EE73,
        end: 0x01EE73,
    },
    Range {
        start: 0x01EE78,
        end: 0x01EE78,
    },
    Range {
        start: 0x01EE7D,
        end: 0x01EE7D,
    },
    Range {
        start: 0x01EE7F,
        end: 0x01EE7F,
    },
    Range {
        start: 0x01EE8A,
        end: 0x01EE8A,
    },
    Range {
        start: 0x01EE9C,
        end: 0x01EEA0,
    },
    Range {
        start: 0x01EEA4,
        end: 0x01EEA4,
    },
    Range {
        start: 0x01EEAA,
        end: 0x01EEAA,
    },
    Range {
        start: 0x01EEBC,
        end: 0x01EEEF,
    },
    Range {
        start: 0x01EEF2,
        end: 0x01EFFF,
    },
    Range {
        start: 0x01F02C,
        end: 0x01F02F,
    },
    Range {
        start: 0x01F094,
        end: 0x01F09F,
    },
    Range {
        start: 0x01F0AF,
        end: 0x01F0B0,
    },
    Range {
        start: 0x01F0C0,
        end: 0x01F0C0,
    },
    Range {
        start: 0x01F0D0,
        end: 0x01F0D0,
    },
    Range {
        start: 0x01F0F6,
        end: 0x01F0FF,
    },
    Range {
        start: 0x01F1AE,
        end: 0x01F1E5,
    },
    Range {
        start: 0x01F203,
        end: 0x01F20F,
    },
    Range {
        start: 0x01F23C,
        end: 0x01F23F,
    },
    Range {
        start: 0x01F249,
        end: 0x01F24F,
    },
    Range {
        start: 0x01F252,
        end: 0x01F25F,
    },
    Range {
        start: 0x01F266,
        end: 0x01F2FF,
    },
    Range {
        start: 0x01F6D9,
        end: 0x01F6DB,
    },
    Range {
        start: 0x01F6ED,
        end: 0x01F6EF,
    },
    Range {
        start: 0x01F6FD,
        end: 0x01F6FF,
    },
    Range {
        start: 0x01F7DA,
        end: 0x01F7DF,
    },
    Range {
        start: 0x01F7EC,
        end: 0x01F7EF,
    },
    Range {
        start: 0x01F7F1,
        end: 0x01F7FF,
    },
    Range {
        start: 0x01F80C,
        end: 0x01F80F,
    },
    Range {
        start: 0x01F848,
        end: 0x01F84F,
    },
    Range {
        start: 0x01F85A,
        end: 0x01F85F,
    },
    Range {
        start: 0x01F888,
        end: 0x01F88F,
    },
    Range {
        start: 0x01F8AE,
        end: 0x01F8AF,
    },
    Range {
        start: 0x01F8BC,
        end: 0x01F8BF,
    },
    Range {
        start: 0x01F8C2,
        end: 0x01F8CF,
    },
    Range {
        start: 0x01F8D9,
        end: 0x01F8FF,
    },
    Range {
        start: 0x01FA58,
        end: 0x01FA5F,
    },
    Range {
        start: 0x01FA6E,
        end: 0x01FA6F,
    },
    Range {
        start: 0x01FA7D,
        end: 0x01FA7F,
    },
    Range {
        start: 0x01FA8B,
        end: 0x01FA8D,
    },
    Range {
        start: 0x01FAC7,
        end: 0x01FAC7,
    },
    Range {
        start: 0x01FAC9,
        end: 0x01FACC,
    },
    Range {
        start: 0x01FADD,
        end: 0x01FADE,
    },
    Range {
        start: 0x01FAEB,
        end: 0x01FAEE,
    },
    Range {
        start: 0x01FAF9,
        end: 0x01FAFF,
    },
    Range {
        start: 0x01FB93,
        end: 0x01FB93,
    },
    Range {
        start: 0x01FBFB,
        end: 0x01FFFF,
    },
    Range {
        start: 0x02A6E0,
        end: 0x02A6FF,
    },
    Range {
        start: 0x02B81E,
        end: 0x02B81F,
    },
    Range {
        start: 0x02CEAE,
        end: 0x02CEAF,
    },
    Range {
        start: 0x02EBE1,
        end: 0x02EBEF,
    },
    Range {
        start: 0x02EE5E,
        end: 0x02F7FF,
    },
    Range {
        start: 0x02FA1E,
        end: 0x02FFFF,
    },
    Range {
        start: 0x03134B,
        end: 0x03134F,
    },
    Range {
        start: 0x03347A,
        end: 0x0E0000,
    },
    Range {
        start: 0x0E0002,
        end: 0x0E001F,
    },
    Range {
        start: 0x0E0080,
        end: 0x0E00FF,
    },
    Range {
        start: 0x0E01F0,
        end: 0x0EFFFF,
    },
    Range {
        start: 0x0FFFFE,
        end: 0x0FFFFF,
    },
    Range {
        start: 0x10FFFE,
        end: 0x10FFFF,
    },
];

const LETTER: &[Range] = &[
    Range {
        start: 0x000041,
        end: 0x00005A,
    },
    Range {
        start: 0x000061,
        end: 0x00007A,
    },
    Range {
        start: 0x0000AA,
        end: 0x0000AA,
    },
    Range {
        start: 0x0000B5,
        end: 0x0000B5,
    },
    Range {
        start: 0x0000BA,
        end: 0x0000BA,
    },
    Range {
        start: 0x0000C0,
        end: 0x0000D6,
    },
    Range {
        start: 0x0000D8,
        end: 0x0000F6,
    },
    Range {
        start: 0x0000F8,
        end: 0x0002C1,
    },
    Range {
        start: 0x0002C6,
        end: 0x0002D1,
    },
    Range {
        start: 0x0002E0,
        end: 0x0002E4,
    },
    Range {
        start: 0x0002EC,
        end: 0x0002EC,
    },
    Range {
        start: 0x0002EE,
        end: 0x0002EE,
    },
    Range {
        start: 0x000370,
        end: 0x000374,
    },
    Range {
        start: 0x000376,
        end: 0x000377,
    },
    Range {
        start: 0x00037A,
        end: 0x00037D,
    },
    Range {
        start: 0x00037F,
        end: 0x00037F,
    },
    Range {
        start: 0x000386,
        end: 0x000386,
    },
    Range {
        start: 0x000388,
        end: 0x00038A,
    },
    Range {
        start: 0x00038C,
        end: 0x00038C,
    },
    Range {
        start: 0x00038E,
        end: 0x0003A1,
    },
    Range {
        start: 0x0003A3,
        end: 0x0003F5,
    },
    Range {
        start: 0x0003F7,
        end: 0x000481,
    },
    Range {
        start: 0x00048A,
        end: 0x00052F,
    },
    Range {
        start: 0x000531,
        end: 0x000556,
    },
    Range {
        start: 0x000559,
        end: 0x000559,
    },
    Range {
        start: 0x000560,
        end: 0x000588,
    },
    Range {
        start: 0x0005D0,
        end: 0x0005EA,
    },
    Range {
        start: 0x0005EF,
        end: 0x0005F2,
    },
    Range {
        start: 0x000620,
        end: 0x00064A,
    },
    Range {
        start: 0x00066E,
        end: 0x00066F,
    },
    Range {
        start: 0x000671,
        end: 0x0006D3,
    },
    Range {
        start: 0x0006D5,
        end: 0x0006D5,
    },
    Range {
        start: 0x0006E5,
        end: 0x0006E6,
    },
    Range {
        start: 0x0006EE,
        end: 0x0006EF,
    },
    Range {
        start: 0x0006FA,
        end: 0x0006FC,
    },
    Range {
        start: 0x0006FF,
        end: 0x0006FF,
    },
    Range {
        start: 0x000710,
        end: 0x000710,
    },
    Range {
        start: 0x000712,
        end: 0x00072F,
    },
    Range {
        start: 0x00074D,
        end: 0x0007A5,
    },
    Range {
        start: 0x0007B1,
        end: 0x0007B1,
    },
    Range {
        start: 0x0007CA,
        end: 0x0007EA,
    },
    Range {
        start: 0x0007F4,
        end: 0x0007F5,
    },
    Range {
        start: 0x0007FA,
        end: 0x0007FA,
    },
    Range {
        start: 0x000800,
        end: 0x000815,
    },
    Range {
        start: 0x00081A,
        end: 0x00081A,
    },
    Range {
        start: 0x000824,
        end: 0x000824,
    },
    Range {
        start: 0x000828,
        end: 0x000828,
    },
    Range {
        start: 0x000840,
        end: 0x000858,
    },
    Range {
        start: 0x000860,
        end: 0x00086A,
    },
    Range {
        start: 0x000870,
        end: 0x000887,
    },
    Range {
        start: 0x000889,
        end: 0x00088F,
    },
    Range {
        start: 0x0008A0,
        end: 0x0008C9,
    },
    Range {
        start: 0x000904,
        end: 0x000939,
    },
    Range {
        start: 0x00093D,
        end: 0x00093D,
    },
    Range {
        start: 0x000950,
        end: 0x000950,
    },
    Range {
        start: 0x000958,
        end: 0x000961,
    },
    Range {
        start: 0x000971,
        end: 0x000980,
    },
    Range {
        start: 0x000985,
        end: 0x00098C,
    },
    Range {
        start: 0x00098F,
        end: 0x000990,
    },
    Range {
        start: 0x000993,
        end: 0x0009A8,
    },
    Range {
        start: 0x0009AA,
        end: 0x0009B0,
    },
    Range {
        start: 0x0009B2,
        end: 0x0009B2,
    },
    Range {
        start: 0x0009B6,
        end: 0x0009B9,
    },
    Range {
        start: 0x0009BD,
        end: 0x0009BD,
    },
    Range {
        start: 0x0009CE,
        end: 0x0009CE,
    },
    Range {
        start: 0x0009DC,
        end: 0x0009DD,
    },
    Range {
        start: 0x0009DF,
        end: 0x0009E1,
    },
    Range {
        start: 0x0009F0,
        end: 0x0009F1,
    },
    Range {
        start: 0x0009FC,
        end: 0x0009FC,
    },
    Range {
        start: 0x000A05,
        end: 0x000A0A,
    },
    Range {
        start: 0x000A0F,
        end: 0x000A10,
    },
    Range {
        start: 0x000A13,
        end: 0x000A28,
    },
    Range {
        start: 0x000A2A,
        end: 0x000A30,
    },
    Range {
        start: 0x000A32,
        end: 0x000A33,
    },
    Range {
        start: 0x000A35,
        end: 0x000A36,
    },
    Range {
        start: 0x000A38,
        end: 0x000A39,
    },
    Range {
        start: 0x000A59,
        end: 0x000A5C,
    },
    Range {
        start: 0x000A5E,
        end: 0x000A5E,
    },
    Range {
        start: 0x000A72,
        end: 0x000A74,
    },
    Range {
        start: 0x000A85,
        end: 0x000A8D,
    },
    Range {
        start: 0x000A8F,
        end: 0x000A91,
    },
    Range {
        start: 0x000A93,
        end: 0x000AA8,
    },
    Range {
        start: 0x000AAA,
        end: 0x000AB0,
    },
    Range {
        start: 0x000AB2,
        end: 0x000AB3,
    },
    Range {
        start: 0x000AB5,
        end: 0x000AB9,
    },
    Range {
        start: 0x000ABD,
        end: 0x000ABD,
    },
    Range {
        start: 0x000AD0,
        end: 0x000AD0,
    },
    Range {
        start: 0x000AE0,
        end: 0x000AE1,
    },
    Range {
        start: 0x000AF9,
        end: 0x000AF9,
    },
    Range {
        start: 0x000B05,
        end: 0x000B0C,
    },
    Range {
        start: 0x000B0F,
        end: 0x000B10,
    },
    Range {
        start: 0x000B13,
        end: 0x000B28,
    },
    Range {
        start: 0x000B2A,
        end: 0x000B30,
    },
    Range {
        start: 0x000B32,
        end: 0x000B33,
    },
    Range {
        start: 0x000B35,
        end: 0x000B39,
    },
    Range {
        start: 0x000B3D,
        end: 0x000B3D,
    },
    Range {
        start: 0x000B5C,
        end: 0x000B5D,
    },
    Range {
        start: 0x000B5F,
        end: 0x000B61,
    },
    Range {
        start: 0x000B71,
        end: 0x000B71,
    },
    Range {
        start: 0x000B83,
        end: 0x000B83,
    },
    Range {
        start: 0x000B85,
        end: 0x000B8A,
    },
    Range {
        start: 0x000B8E,
        end: 0x000B90,
    },
    Range {
        start: 0x000B92,
        end: 0x000B95,
    },
    Range {
        start: 0x000B99,
        end: 0x000B9A,
    },
    Range {
        start: 0x000B9C,
        end: 0x000B9C,
    },
    Range {
        start: 0x000B9E,
        end: 0x000B9F,
    },
    Range {
        start: 0x000BA3,
        end: 0x000BA4,
    },
    Range {
        start: 0x000BA8,
        end: 0x000BAA,
    },
    Range {
        start: 0x000BAE,
        end: 0x000BB9,
    },
    Range {
        start: 0x000BD0,
        end: 0x000BD0,
    },
    Range {
        start: 0x000C05,
        end: 0x000C0C,
    },
    Range {
        start: 0x000C0E,
        end: 0x000C10,
    },
    Range {
        start: 0x000C12,
        end: 0x000C28,
    },
    Range {
        start: 0x000C2A,
        end: 0x000C39,
    },
    Range {
        start: 0x000C3D,
        end: 0x000C3D,
    },
    Range {
        start: 0x000C58,
        end: 0x000C5A,
    },
    Range {
        start: 0x000C5C,
        end: 0x000C5D,
    },
    Range {
        start: 0x000C60,
        end: 0x000C61,
    },
    Range {
        start: 0x000C80,
        end: 0x000C80,
    },
    Range {
        start: 0x000C85,
        end: 0x000C8C,
    },
    Range {
        start: 0x000C8E,
        end: 0x000C90,
    },
    Range {
        start: 0x000C92,
        end: 0x000CA8,
    },
    Range {
        start: 0x000CAA,
        end: 0x000CB3,
    },
    Range {
        start: 0x000CB5,
        end: 0x000CB9,
    },
    Range {
        start: 0x000CBD,
        end: 0x000CBD,
    },
    Range {
        start: 0x000CDC,
        end: 0x000CDE,
    },
    Range {
        start: 0x000CE0,
        end: 0x000CE1,
    },
    Range {
        start: 0x000CF1,
        end: 0x000CF2,
    },
    Range {
        start: 0x000D04,
        end: 0x000D0C,
    },
    Range {
        start: 0x000D0E,
        end: 0x000D10,
    },
    Range {
        start: 0x000D12,
        end: 0x000D3A,
    },
    Range {
        start: 0x000D3D,
        end: 0x000D3D,
    },
    Range {
        start: 0x000D4E,
        end: 0x000D4E,
    },
    Range {
        start: 0x000D54,
        end: 0x000D56,
    },
    Range {
        start: 0x000D5F,
        end: 0x000D61,
    },
    Range {
        start: 0x000D7A,
        end: 0x000D7F,
    },
    Range {
        start: 0x000D85,
        end: 0x000D96,
    },
    Range {
        start: 0x000D9A,
        end: 0x000DB1,
    },
    Range {
        start: 0x000DB3,
        end: 0x000DBB,
    },
    Range {
        start: 0x000DBD,
        end: 0x000DBD,
    },
    Range {
        start: 0x000DC0,
        end: 0x000DC6,
    },
    Range {
        start: 0x000E01,
        end: 0x000E30,
    },
    Range {
        start: 0x000E32,
        end: 0x000E33,
    },
    Range {
        start: 0x000E40,
        end: 0x000E46,
    },
    Range {
        start: 0x000E81,
        end: 0x000E82,
    },
    Range {
        start: 0x000E84,
        end: 0x000E84,
    },
    Range {
        start: 0x000E86,
        end: 0x000E8A,
    },
    Range {
        start: 0x000E8C,
        end: 0x000EA3,
    },
    Range {
        start: 0x000EA5,
        end: 0x000EA5,
    },
    Range {
        start: 0x000EA7,
        end: 0x000EB0,
    },
    Range {
        start: 0x000EB2,
        end: 0x000EB3,
    },
    Range {
        start: 0x000EBD,
        end: 0x000EBD,
    },
    Range {
        start: 0x000EC0,
        end: 0x000EC4,
    },
    Range {
        start: 0x000EC6,
        end: 0x000EC6,
    },
    Range {
        start: 0x000EDC,
        end: 0x000EDF,
    },
    Range {
        start: 0x000F00,
        end: 0x000F00,
    },
    Range {
        start: 0x000F40,
        end: 0x000F47,
    },
    Range {
        start: 0x000F49,
        end: 0x000F6C,
    },
    Range {
        start: 0x000F88,
        end: 0x000F8C,
    },
    Range {
        start: 0x001000,
        end: 0x00102A,
    },
    Range {
        start: 0x00103F,
        end: 0x00103F,
    },
    Range {
        start: 0x001050,
        end: 0x001055,
    },
    Range {
        start: 0x00105A,
        end: 0x00105D,
    },
    Range {
        start: 0x001061,
        end: 0x001061,
    },
    Range {
        start: 0x001065,
        end: 0x001066,
    },
    Range {
        start: 0x00106E,
        end: 0x001070,
    },
    Range {
        start: 0x001075,
        end: 0x001081,
    },
    Range {
        start: 0x00108E,
        end: 0x00108E,
    },
    Range {
        start: 0x0010A0,
        end: 0x0010C5,
    },
    Range {
        start: 0x0010C7,
        end: 0x0010C7,
    },
    Range {
        start: 0x0010CD,
        end: 0x0010CD,
    },
    Range {
        start: 0x0010D0,
        end: 0x0010FA,
    },
    Range {
        start: 0x0010FC,
        end: 0x001248,
    },
    Range {
        start: 0x00124A,
        end: 0x00124D,
    },
    Range {
        start: 0x001250,
        end: 0x001256,
    },
    Range {
        start: 0x001258,
        end: 0x001258,
    },
    Range {
        start: 0x00125A,
        end: 0x00125D,
    },
    Range {
        start: 0x001260,
        end: 0x001288,
    },
    Range {
        start: 0x00128A,
        end: 0x00128D,
    },
    Range {
        start: 0x001290,
        end: 0x0012B0,
    },
    Range {
        start: 0x0012B2,
        end: 0x0012B5,
    },
    Range {
        start: 0x0012B8,
        end: 0x0012BE,
    },
    Range {
        start: 0x0012C0,
        end: 0x0012C0,
    },
    Range {
        start: 0x0012C2,
        end: 0x0012C5,
    },
    Range {
        start: 0x0012C8,
        end: 0x0012D6,
    },
    Range {
        start: 0x0012D8,
        end: 0x001310,
    },
    Range {
        start: 0x001312,
        end: 0x001315,
    },
    Range {
        start: 0x001318,
        end: 0x00135A,
    },
    Range {
        start: 0x001380,
        end: 0x00138F,
    },
    Range {
        start: 0x0013A0,
        end: 0x0013F5,
    },
    Range {
        start: 0x0013F8,
        end: 0x0013FD,
    },
    Range {
        start: 0x001401,
        end: 0x00166C,
    },
    Range {
        start: 0x00166F,
        end: 0x00167F,
    },
    Range {
        start: 0x001681,
        end: 0x00169A,
    },
    Range {
        start: 0x0016A0,
        end: 0x0016EA,
    },
    Range {
        start: 0x0016F1,
        end: 0x0016F8,
    },
    Range {
        start: 0x001700,
        end: 0x001711,
    },
    Range {
        start: 0x00171F,
        end: 0x001731,
    },
    Range {
        start: 0x001740,
        end: 0x001751,
    },
    Range {
        start: 0x001760,
        end: 0x00176C,
    },
    Range {
        start: 0x00176E,
        end: 0x001770,
    },
    Range {
        start: 0x001780,
        end: 0x0017B3,
    },
    Range {
        start: 0x0017D7,
        end: 0x0017D7,
    },
    Range {
        start: 0x0017DC,
        end: 0x0017DC,
    },
    Range {
        start: 0x001820,
        end: 0x001878,
    },
    Range {
        start: 0x001880,
        end: 0x001884,
    },
    Range {
        start: 0x001887,
        end: 0x0018A8,
    },
    Range {
        start: 0x0018AA,
        end: 0x0018AA,
    },
    Range {
        start: 0x0018B0,
        end: 0x0018F5,
    },
    Range {
        start: 0x001900,
        end: 0x00191E,
    },
    Range {
        start: 0x001950,
        end: 0x00196D,
    },
    Range {
        start: 0x001970,
        end: 0x001974,
    },
    Range {
        start: 0x001980,
        end: 0x0019AB,
    },
    Range {
        start: 0x0019B0,
        end: 0x0019C9,
    },
    Range {
        start: 0x001A00,
        end: 0x001A16,
    },
    Range {
        start: 0x001A20,
        end: 0x001A54,
    },
    Range {
        start: 0x001AA7,
        end: 0x001AA7,
    },
    Range {
        start: 0x001B05,
        end: 0x001B33,
    },
    Range {
        start: 0x001B45,
        end: 0x001B4C,
    },
    Range {
        start: 0x001B83,
        end: 0x001BA0,
    },
    Range {
        start: 0x001BAE,
        end: 0x001BAF,
    },
    Range {
        start: 0x001BBA,
        end: 0x001BE5,
    },
    Range {
        start: 0x001C00,
        end: 0x001C23,
    },
    Range {
        start: 0x001C4D,
        end: 0x001C4F,
    },
    Range {
        start: 0x001C5A,
        end: 0x001C7D,
    },
    Range {
        start: 0x001C80,
        end: 0x001C8A,
    },
    Range {
        start: 0x001C90,
        end: 0x001CBA,
    },
    Range {
        start: 0x001CBD,
        end: 0x001CBF,
    },
    Range {
        start: 0x001CE9,
        end: 0x001CEC,
    },
    Range {
        start: 0x001CEE,
        end: 0x001CF3,
    },
    Range {
        start: 0x001CF5,
        end: 0x001CF6,
    },
    Range {
        start: 0x001CFA,
        end: 0x001CFA,
    },
    Range {
        start: 0x001D00,
        end: 0x001DBF,
    },
    Range {
        start: 0x001E00,
        end: 0x001F15,
    },
    Range {
        start: 0x001F18,
        end: 0x001F1D,
    },
    Range {
        start: 0x001F20,
        end: 0x001F45,
    },
    Range {
        start: 0x001F48,
        end: 0x001F4D,
    },
    Range {
        start: 0x001F50,
        end: 0x001F57,
    },
    Range {
        start: 0x001F59,
        end: 0x001F59,
    },
    Range {
        start: 0x001F5B,
        end: 0x001F5B,
    },
    Range {
        start: 0x001F5D,
        end: 0x001F5D,
    },
    Range {
        start: 0x001F5F,
        end: 0x001F7D,
    },
    Range {
        start: 0x001F80,
        end: 0x001FB4,
    },
    Range {
        start: 0x001FB6,
        end: 0x001FBC,
    },
    Range {
        start: 0x001FBE,
        end: 0x001FBE,
    },
    Range {
        start: 0x001FC2,
        end: 0x001FC4,
    },
    Range {
        start: 0x001FC6,
        end: 0x001FCC,
    },
    Range {
        start: 0x001FD0,
        end: 0x001FD3,
    },
    Range {
        start: 0x001FD6,
        end: 0x001FDB,
    },
    Range {
        start: 0x001FE0,
        end: 0x001FEC,
    },
    Range {
        start: 0x001FF2,
        end: 0x001FF4,
    },
    Range {
        start: 0x001FF6,
        end: 0x001FFC,
    },
    Range {
        start: 0x002071,
        end: 0x002071,
    },
    Range {
        start: 0x00207F,
        end: 0x00207F,
    },
    Range {
        start: 0x002090,
        end: 0x00209C,
    },
    Range {
        start: 0x002102,
        end: 0x002102,
    },
    Range {
        start: 0x002107,
        end: 0x002107,
    },
    Range {
        start: 0x00210A,
        end: 0x002113,
    },
    Range {
        start: 0x002115,
        end: 0x002115,
    },
    Range {
        start: 0x002119,
        end: 0x00211D,
    },
    Range {
        start: 0x002124,
        end: 0x002124,
    },
    Range {
        start: 0x002126,
        end: 0x002126,
    },
    Range {
        start: 0x002128,
        end: 0x002128,
    },
    Range {
        start: 0x00212A,
        end: 0x00212D,
    },
    Range {
        start: 0x00212F,
        end: 0x002139,
    },
    Range {
        start: 0x00213C,
        end: 0x00213F,
    },
    Range {
        start: 0x002145,
        end: 0x002149,
    },
    Range {
        start: 0x00214E,
        end: 0x00214E,
    },
    Range {
        start: 0x002183,
        end: 0x002184,
    },
    Range {
        start: 0x002C00,
        end: 0x002CE4,
    },
    Range {
        start: 0x002CEB,
        end: 0x002CEE,
    },
    Range {
        start: 0x002CF2,
        end: 0x002CF3,
    },
    Range {
        start: 0x002D00,
        end: 0x002D25,
    },
    Range {
        start: 0x002D27,
        end: 0x002D27,
    },
    Range {
        start: 0x002D2D,
        end: 0x002D2D,
    },
    Range {
        start: 0x002D30,
        end: 0x002D67,
    },
    Range {
        start: 0x002D6F,
        end: 0x002D6F,
    },
    Range {
        start: 0x002D80,
        end: 0x002D96,
    },
    Range {
        start: 0x002DA0,
        end: 0x002DA6,
    },
    Range {
        start: 0x002DA8,
        end: 0x002DAE,
    },
    Range {
        start: 0x002DB0,
        end: 0x002DB6,
    },
    Range {
        start: 0x002DB8,
        end: 0x002DBE,
    },
    Range {
        start: 0x002DC0,
        end: 0x002DC6,
    },
    Range {
        start: 0x002DC8,
        end: 0x002DCE,
    },
    Range {
        start: 0x002DD0,
        end: 0x002DD6,
    },
    Range {
        start: 0x002DD8,
        end: 0x002DDE,
    },
    Range {
        start: 0x002E2F,
        end: 0x002E2F,
    },
    Range {
        start: 0x003005,
        end: 0x003006,
    },
    Range {
        start: 0x003031,
        end: 0x003035,
    },
    Range {
        start: 0x00303B,
        end: 0x00303C,
    },
    Range {
        start: 0x003041,
        end: 0x003096,
    },
    Range {
        start: 0x00309D,
        end: 0x00309F,
    },
    Range {
        start: 0x0030A1,
        end: 0x0030FA,
    },
    Range {
        start: 0x0030FC,
        end: 0x0030FF,
    },
    Range {
        start: 0x003105,
        end: 0x00312F,
    },
    Range {
        start: 0x003131,
        end: 0x00318E,
    },
    Range {
        start: 0x0031A0,
        end: 0x0031BF,
    },
    Range {
        start: 0x0031F0,
        end: 0x0031FF,
    },
    Range {
        start: 0x003400,
        end: 0x004DBF,
    },
    Range {
        start: 0x004E00,
        end: 0x00A48C,
    },
    Range {
        start: 0x00A4D0,
        end: 0x00A4FD,
    },
    Range {
        start: 0x00A500,
        end: 0x00A60C,
    },
    Range {
        start: 0x00A610,
        end: 0x00A61F,
    },
    Range {
        start: 0x00A62A,
        end: 0x00A62B,
    },
    Range {
        start: 0x00A640,
        end: 0x00A66E,
    },
    Range {
        start: 0x00A67F,
        end: 0x00A69D,
    },
    Range {
        start: 0x00A6A0,
        end: 0x00A6E5,
    },
    Range {
        start: 0x00A717,
        end: 0x00A71F,
    },
    Range {
        start: 0x00A722,
        end: 0x00A788,
    },
    Range {
        start: 0x00A78B,
        end: 0x00A7DC,
    },
    Range {
        start: 0x00A7F1,
        end: 0x00A801,
    },
    Range {
        start: 0x00A803,
        end: 0x00A805,
    },
    Range {
        start: 0x00A807,
        end: 0x00A80A,
    },
    Range {
        start: 0x00A80C,
        end: 0x00A822,
    },
    Range {
        start: 0x00A840,
        end: 0x00A873,
    },
    Range {
        start: 0x00A882,
        end: 0x00A8B3,
    },
    Range {
        start: 0x00A8F2,
        end: 0x00A8F7,
    },
    Range {
        start: 0x00A8FB,
        end: 0x00A8FB,
    },
    Range {
        start: 0x00A8FD,
        end: 0x00A8FE,
    },
    Range {
        start: 0x00A90A,
        end: 0x00A925,
    },
    Range {
        start: 0x00A930,
        end: 0x00A946,
    },
    Range {
        start: 0x00A960,
        end: 0x00A97C,
    },
    Range {
        start: 0x00A984,
        end: 0x00A9B2,
    },
    Range {
        start: 0x00A9CF,
        end: 0x00A9CF,
    },
    Range {
        start: 0x00A9E0,
        end: 0x00A9E4,
    },
    Range {
        start: 0x00A9E6,
        end: 0x00A9EF,
    },
    Range {
        start: 0x00A9FA,
        end: 0x00A9FE,
    },
    Range {
        start: 0x00AA00,
        end: 0x00AA28,
    },
    Range {
        start: 0x00AA40,
        end: 0x00AA42,
    },
    Range {
        start: 0x00AA44,
        end: 0x00AA4B,
    },
    Range {
        start: 0x00AA60,
        end: 0x00AA76,
    },
    Range {
        start: 0x00AA7A,
        end: 0x00AA7A,
    },
    Range {
        start: 0x00AA7E,
        end: 0x00AAAF,
    },
    Range {
        start: 0x00AAB1,
        end: 0x00AAB1,
    },
    Range {
        start: 0x00AAB5,
        end: 0x00AAB6,
    },
    Range {
        start: 0x00AAB9,
        end: 0x00AABD,
    },
    Range {
        start: 0x00AAC0,
        end: 0x00AAC0,
    },
    Range {
        start: 0x00AAC2,
        end: 0x00AAC2,
    },
    Range {
        start: 0x00AADB,
        end: 0x00AADD,
    },
    Range {
        start: 0x00AAE0,
        end: 0x00AAEA,
    },
    Range {
        start: 0x00AAF2,
        end: 0x00AAF4,
    },
    Range {
        start: 0x00AB01,
        end: 0x00AB06,
    },
    Range {
        start: 0x00AB09,
        end: 0x00AB0E,
    },
    Range {
        start: 0x00AB11,
        end: 0x00AB16,
    },
    Range {
        start: 0x00AB20,
        end: 0x00AB26,
    },
    Range {
        start: 0x00AB28,
        end: 0x00AB2E,
    },
    Range {
        start: 0x00AB30,
        end: 0x00AB5A,
    },
    Range {
        start: 0x00AB5C,
        end: 0x00AB69,
    },
    Range {
        start: 0x00AB70,
        end: 0x00ABE2,
    },
    Range {
        start: 0x00AC00,
        end: 0x00D7A3,
    },
    Range {
        start: 0x00D7B0,
        end: 0x00D7C6,
    },
    Range {
        start: 0x00D7CB,
        end: 0x00D7FB,
    },
    Range {
        start: 0x00F900,
        end: 0x00FA6D,
    },
    Range {
        start: 0x00FA70,
        end: 0x00FAD9,
    },
    Range {
        start: 0x00FB00,
        end: 0x00FB06,
    },
    Range {
        start: 0x00FB13,
        end: 0x00FB17,
    },
    Range {
        start: 0x00FB1D,
        end: 0x00FB1D,
    },
    Range {
        start: 0x00FB1F,
        end: 0x00FB28,
    },
    Range {
        start: 0x00FB2A,
        end: 0x00FB36,
    },
    Range {
        start: 0x00FB38,
        end: 0x00FB3C,
    },
    Range {
        start: 0x00FB3E,
        end: 0x00FB3E,
    },
    Range {
        start: 0x00FB40,
        end: 0x00FB41,
    },
    Range {
        start: 0x00FB43,
        end: 0x00FB44,
    },
    Range {
        start: 0x00FB46,
        end: 0x00FBB1,
    },
    Range {
        start: 0x00FBD3,
        end: 0x00FD3D,
    },
    Range {
        start: 0x00FD50,
        end: 0x00FD8F,
    },
    Range {
        start: 0x00FD92,
        end: 0x00FDC7,
    },
    Range {
        start: 0x00FDF0,
        end: 0x00FDFB,
    },
    Range {
        start: 0x00FE70,
        end: 0x00FE74,
    },
    Range {
        start: 0x00FE76,
        end: 0x00FEFC,
    },
    Range {
        start: 0x00FF21,
        end: 0x00FF3A,
    },
    Range {
        start: 0x00FF41,
        end: 0x00FF5A,
    },
    Range {
        start: 0x00FF66,
        end: 0x00FFBE,
    },
    Range {
        start: 0x00FFC2,
        end: 0x00FFC7,
    },
    Range {
        start: 0x00FFCA,
        end: 0x00FFCF,
    },
    Range {
        start: 0x00FFD2,
        end: 0x00FFD7,
    },
    Range {
        start: 0x00FFDA,
        end: 0x00FFDC,
    },
    Range {
        start: 0x010000,
        end: 0x01000B,
    },
    Range {
        start: 0x01000D,
        end: 0x010026,
    },
    Range {
        start: 0x010028,
        end: 0x01003A,
    },
    Range {
        start: 0x01003C,
        end: 0x01003D,
    },
    Range {
        start: 0x01003F,
        end: 0x01004D,
    },
    Range {
        start: 0x010050,
        end: 0x01005D,
    },
    Range {
        start: 0x010080,
        end: 0x0100FA,
    },
    Range {
        start: 0x010280,
        end: 0x01029C,
    },
    Range {
        start: 0x0102A0,
        end: 0x0102D0,
    },
    Range {
        start: 0x010300,
        end: 0x01031F,
    },
    Range {
        start: 0x01032D,
        end: 0x010340,
    },
    Range {
        start: 0x010342,
        end: 0x010349,
    },
    Range {
        start: 0x010350,
        end: 0x010375,
    },
    Range {
        start: 0x010380,
        end: 0x01039D,
    },
    Range {
        start: 0x0103A0,
        end: 0x0103C3,
    },
    Range {
        start: 0x0103C8,
        end: 0x0103CF,
    },
    Range {
        start: 0x010400,
        end: 0x01049D,
    },
    Range {
        start: 0x0104B0,
        end: 0x0104D3,
    },
    Range {
        start: 0x0104D8,
        end: 0x0104FB,
    },
    Range {
        start: 0x010500,
        end: 0x010527,
    },
    Range {
        start: 0x010530,
        end: 0x010563,
    },
    Range {
        start: 0x010570,
        end: 0x01057A,
    },
    Range {
        start: 0x01057C,
        end: 0x01058A,
    },
    Range {
        start: 0x01058C,
        end: 0x010592,
    },
    Range {
        start: 0x010594,
        end: 0x010595,
    },
    Range {
        start: 0x010597,
        end: 0x0105A1,
    },
    Range {
        start: 0x0105A3,
        end: 0x0105B1,
    },
    Range {
        start: 0x0105B3,
        end: 0x0105B9,
    },
    Range {
        start: 0x0105BB,
        end: 0x0105BC,
    },
    Range {
        start: 0x0105C0,
        end: 0x0105F3,
    },
    Range {
        start: 0x010600,
        end: 0x010736,
    },
    Range {
        start: 0x010740,
        end: 0x010755,
    },
    Range {
        start: 0x010760,
        end: 0x010767,
    },
    Range {
        start: 0x010780,
        end: 0x010785,
    },
    Range {
        start: 0x010787,
        end: 0x0107B0,
    },
    Range {
        start: 0x0107B2,
        end: 0x0107BA,
    },
    Range {
        start: 0x010800,
        end: 0x010805,
    },
    Range {
        start: 0x010808,
        end: 0x010808,
    },
    Range {
        start: 0x01080A,
        end: 0x010835,
    },
    Range {
        start: 0x010837,
        end: 0x010838,
    },
    Range {
        start: 0x01083C,
        end: 0x01083C,
    },
    Range {
        start: 0x01083F,
        end: 0x010855,
    },
    Range {
        start: 0x010860,
        end: 0x010876,
    },
    Range {
        start: 0x010880,
        end: 0x01089E,
    },
    Range {
        start: 0x0108E0,
        end: 0x0108F2,
    },
    Range {
        start: 0x0108F4,
        end: 0x0108F5,
    },
    Range {
        start: 0x010900,
        end: 0x010915,
    },
    Range {
        start: 0x010920,
        end: 0x010939,
    },
    Range {
        start: 0x010940,
        end: 0x010959,
    },
    Range {
        start: 0x010980,
        end: 0x0109B7,
    },
    Range {
        start: 0x0109BE,
        end: 0x0109BF,
    },
    Range {
        start: 0x010A00,
        end: 0x010A00,
    },
    Range {
        start: 0x010A10,
        end: 0x010A13,
    },
    Range {
        start: 0x010A15,
        end: 0x010A17,
    },
    Range {
        start: 0x010A19,
        end: 0x010A35,
    },
    Range {
        start: 0x010A60,
        end: 0x010A7C,
    },
    Range {
        start: 0x010A80,
        end: 0x010A9C,
    },
    Range {
        start: 0x010AC0,
        end: 0x010AC7,
    },
    Range {
        start: 0x010AC9,
        end: 0x010AE4,
    },
    Range {
        start: 0x010B00,
        end: 0x010B35,
    },
    Range {
        start: 0x010B40,
        end: 0x010B55,
    },
    Range {
        start: 0x010B60,
        end: 0x010B72,
    },
    Range {
        start: 0x010B80,
        end: 0x010B91,
    },
    Range {
        start: 0x010C00,
        end: 0x010C48,
    },
    Range {
        start: 0x010C80,
        end: 0x010CB2,
    },
    Range {
        start: 0x010CC0,
        end: 0x010CF2,
    },
    Range {
        start: 0x010D00,
        end: 0x010D23,
    },
    Range {
        start: 0x010D4A,
        end: 0x010D65,
    },
    Range {
        start: 0x010D6F,
        end: 0x010D85,
    },
    Range {
        start: 0x010E80,
        end: 0x010EA9,
    },
    Range {
        start: 0x010EB0,
        end: 0x010EB1,
    },
    Range {
        start: 0x010EC2,
        end: 0x010EC7,
    },
    Range {
        start: 0x010F00,
        end: 0x010F1C,
    },
    Range {
        start: 0x010F27,
        end: 0x010F27,
    },
    Range {
        start: 0x010F30,
        end: 0x010F45,
    },
    Range {
        start: 0x010F70,
        end: 0x010F81,
    },
    Range {
        start: 0x010FB0,
        end: 0x010FC4,
    },
    Range {
        start: 0x010FE0,
        end: 0x010FF6,
    },
    Range {
        start: 0x011003,
        end: 0x011037,
    },
    Range {
        start: 0x011071,
        end: 0x011072,
    },
    Range {
        start: 0x011075,
        end: 0x011075,
    },
    Range {
        start: 0x011083,
        end: 0x0110AF,
    },
    Range {
        start: 0x0110D0,
        end: 0x0110E8,
    },
    Range {
        start: 0x011103,
        end: 0x011126,
    },
    Range {
        start: 0x011144,
        end: 0x011144,
    },
    Range {
        start: 0x011147,
        end: 0x011147,
    },
    Range {
        start: 0x011150,
        end: 0x011172,
    },
    Range {
        start: 0x011176,
        end: 0x011176,
    },
    Range {
        start: 0x011183,
        end: 0x0111B2,
    },
    Range {
        start: 0x0111C1,
        end: 0x0111C4,
    },
    Range {
        start: 0x0111DA,
        end: 0x0111DA,
    },
    Range {
        start: 0x0111DC,
        end: 0x0111DC,
    },
    Range {
        start: 0x011200,
        end: 0x011211,
    },
    Range {
        start: 0x011213,
        end: 0x01122B,
    },
    Range {
        start: 0x01123F,
        end: 0x011240,
    },
    Range {
        start: 0x011280,
        end: 0x011286,
    },
    Range {
        start: 0x011288,
        end: 0x011288,
    },
    Range {
        start: 0x01128A,
        end: 0x01128D,
    },
    Range {
        start: 0x01128F,
        end: 0x01129D,
    },
    Range {
        start: 0x01129F,
        end: 0x0112A8,
    },
    Range {
        start: 0x0112B0,
        end: 0x0112DE,
    },
    Range {
        start: 0x011305,
        end: 0x01130C,
    },
    Range {
        start: 0x01130F,
        end: 0x011310,
    },
    Range {
        start: 0x011313,
        end: 0x011328,
    },
    Range {
        start: 0x01132A,
        end: 0x011330,
    },
    Range {
        start: 0x011332,
        end: 0x011333,
    },
    Range {
        start: 0x011335,
        end: 0x011339,
    },
    Range {
        start: 0x01133D,
        end: 0x01133D,
    },
    Range {
        start: 0x011350,
        end: 0x011350,
    },
    Range {
        start: 0x01135D,
        end: 0x011361,
    },
    Range {
        start: 0x011380,
        end: 0x011389,
    },
    Range {
        start: 0x01138B,
        end: 0x01138B,
    },
    Range {
        start: 0x01138E,
        end: 0x01138E,
    },
    Range {
        start: 0x011390,
        end: 0x0113B5,
    },
    Range {
        start: 0x0113B7,
        end: 0x0113B7,
    },
    Range {
        start: 0x0113D1,
        end: 0x0113D1,
    },
    Range {
        start: 0x0113D3,
        end: 0x0113D3,
    },
    Range {
        start: 0x011400,
        end: 0x011434,
    },
    Range {
        start: 0x011447,
        end: 0x01144A,
    },
    Range {
        start: 0x01145F,
        end: 0x011461,
    },
    Range {
        start: 0x011480,
        end: 0x0114AF,
    },
    Range {
        start: 0x0114C4,
        end: 0x0114C5,
    },
    Range {
        start: 0x0114C7,
        end: 0x0114C7,
    },
    Range {
        start: 0x011580,
        end: 0x0115AE,
    },
    Range {
        start: 0x0115D8,
        end: 0x0115DB,
    },
    Range {
        start: 0x011600,
        end: 0x01162F,
    },
    Range {
        start: 0x011644,
        end: 0x011644,
    },
    Range {
        start: 0x011680,
        end: 0x0116AA,
    },
    Range {
        start: 0x0116B8,
        end: 0x0116B8,
    },
    Range {
        start: 0x011700,
        end: 0x01171A,
    },
    Range {
        start: 0x011740,
        end: 0x011746,
    },
    Range {
        start: 0x011800,
        end: 0x01182B,
    },
    Range {
        start: 0x0118A0,
        end: 0x0118DF,
    },
    Range {
        start: 0x0118FF,
        end: 0x011906,
    },
    Range {
        start: 0x011909,
        end: 0x011909,
    },
    Range {
        start: 0x01190C,
        end: 0x011913,
    },
    Range {
        start: 0x011915,
        end: 0x011916,
    },
    Range {
        start: 0x011918,
        end: 0x01192F,
    },
    Range {
        start: 0x01193F,
        end: 0x01193F,
    },
    Range {
        start: 0x011941,
        end: 0x011941,
    },
    Range {
        start: 0x0119A0,
        end: 0x0119A7,
    },
    Range {
        start: 0x0119AA,
        end: 0x0119D0,
    },
    Range {
        start: 0x0119E1,
        end: 0x0119E1,
    },
    Range {
        start: 0x0119E3,
        end: 0x0119E3,
    },
    Range {
        start: 0x011A00,
        end: 0x011A00,
    },
    Range {
        start: 0x011A0B,
        end: 0x011A32,
    },
    Range {
        start: 0x011A3A,
        end: 0x011A3A,
    },
    Range {
        start: 0x011A50,
        end: 0x011A50,
    },
    Range {
        start: 0x011A5C,
        end: 0x011A89,
    },
    Range {
        start: 0x011A9D,
        end: 0x011A9D,
    },
    Range {
        start: 0x011AB0,
        end: 0x011AF8,
    },
    Range {
        start: 0x011BC0,
        end: 0x011BE0,
    },
    Range {
        start: 0x011C00,
        end: 0x011C08,
    },
    Range {
        start: 0x011C0A,
        end: 0x011C2E,
    },
    Range {
        start: 0x011C40,
        end: 0x011C40,
    },
    Range {
        start: 0x011C72,
        end: 0x011C8F,
    },
    Range {
        start: 0x011D00,
        end: 0x011D06,
    },
    Range {
        start: 0x011D08,
        end: 0x011D09,
    },
    Range {
        start: 0x011D0B,
        end: 0x011D30,
    },
    Range {
        start: 0x011D46,
        end: 0x011D46,
    },
    Range {
        start: 0x011D60,
        end: 0x011D65,
    },
    Range {
        start: 0x011D67,
        end: 0x011D68,
    },
    Range {
        start: 0x011D6A,
        end: 0x011D89,
    },
    Range {
        start: 0x011D98,
        end: 0x011D98,
    },
    Range {
        start: 0x011DB0,
        end: 0x011DDB,
    },
    Range {
        start: 0x011EE0,
        end: 0x011EF2,
    },
    Range {
        start: 0x011F02,
        end: 0x011F02,
    },
    Range {
        start: 0x011F04,
        end: 0x011F10,
    },
    Range {
        start: 0x011F12,
        end: 0x011F33,
    },
    Range {
        start: 0x011FB0,
        end: 0x011FB0,
    },
    Range {
        start: 0x012000,
        end: 0x012399,
    },
    Range {
        start: 0x012480,
        end: 0x012543,
    },
    Range {
        start: 0x012F90,
        end: 0x012FF0,
    },
    Range {
        start: 0x013000,
        end: 0x01342F,
    },
    Range {
        start: 0x013441,
        end: 0x013446,
    },
    Range {
        start: 0x013460,
        end: 0x0143FA,
    },
    Range {
        start: 0x014400,
        end: 0x014646,
    },
    Range {
        start: 0x016100,
        end: 0x01611D,
    },
    Range {
        start: 0x016800,
        end: 0x016A38,
    },
    Range {
        start: 0x016A40,
        end: 0x016A5E,
    },
    Range {
        start: 0x016A70,
        end: 0x016ABE,
    },
    Range {
        start: 0x016AD0,
        end: 0x016AED,
    },
    Range {
        start: 0x016B00,
        end: 0x016B2F,
    },
    Range {
        start: 0x016B40,
        end: 0x016B43,
    },
    Range {
        start: 0x016B63,
        end: 0x016B77,
    },
    Range {
        start: 0x016B7D,
        end: 0x016B8F,
    },
    Range {
        start: 0x016D40,
        end: 0x016D6C,
    },
    Range {
        start: 0x016E40,
        end: 0x016E7F,
    },
    Range {
        start: 0x016EA0,
        end: 0x016EB8,
    },
    Range {
        start: 0x016EBB,
        end: 0x016ED3,
    },
    Range {
        start: 0x016F00,
        end: 0x016F4A,
    },
    Range {
        start: 0x016F50,
        end: 0x016F50,
    },
    Range {
        start: 0x016F93,
        end: 0x016F9F,
    },
    Range {
        start: 0x016FE0,
        end: 0x016FE1,
    },
    Range {
        start: 0x016FE3,
        end: 0x016FE3,
    },
    Range {
        start: 0x016FF2,
        end: 0x016FF3,
    },
    Range {
        start: 0x017000,
        end: 0x018CD5,
    },
    Range {
        start: 0x018CFF,
        end: 0x018D1E,
    },
    Range {
        start: 0x018D80,
        end: 0x018DF2,
    },
    Range {
        start: 0x01AFF0,
        end: 0x01AFF3,
    },
    Range {
        start: 0x01AFF5,
        end: 0x01AFFB,
    },
    Range {
        start: 0x01AFFD,
        end: 0x01AFFE,
    },
    Range {
        start: 0x01B000,
        end: 0x01B122,
    },
    Range {
        start: 0x01B132,
        end: 0x01B132,
    },
    Range {
        start: 0x01B150,
        end: 0x01B152,
    },
    Range {
        start: 0x01B155,
        end: 0x01B155,
    },
    Range {
        start: 0x01B164,
        end: 0x01B167,
    },
    Range {
        start: 0x01B170,
        end: 0x01B2FB,
    },
    Range {
        start: 0x01BC00,
        end: 0x01BC6A,
    },
    Range {
        start: 0x01BC70,
        end: 0x01BC7C,
    },
    Range {
        start: 0x01BC80,
        end: 0x01BC88,
    },
    Range {
        start: 0x01BC90,
        end: 0x01BC99,
    },
    Range {
        start: 0x01D400,
        end: 0x01D454,
    },
    Range {
        start: 0x01D456,
        end: 0x01D49C,
    },
    Range {
        start: 0x01D49E,
        end: 0x01D49F,
    },
    Range {
        start: 0x01D4A2,
        end: 0x01D4A2,
    },
    Range {
        start: 0x01D4A5,
        end: 0x01D4A6,
    },
    Range {
        start: 0x01D4A9,
        end: 0x01D4AC,
    },
    Range {
        start: 0x01D4AE,
        end: 0x01D4B9,
    },
    Range {
        start: 0x01D4BB,
        end: 0x01D4BB,
    },
    Range {
        start: 0x01D4BD,
        end: 0x01D4C3,
    },
    Range {
        start: 0x01D4C5,
        end: 0x01D505,
    },
    Range {
        start: 0x01D507,
        end: 0x01D50A,
    },
    Range {
        start: 0x01D50D,
        end: 0x01D514,
    },
    Range {
        start: 0x01D516,
        end: 0x01D51C,
    },
    Range {
        start: 0x01D51E,
        end: 0x01D539,
    },
    Range {
        start: 0x01D53B,
        end: 0x01D53E,
    },
    Range {
        start: 0x01D540,
        end: 0x01D544,
    },
    Range {
        start: 0x01D546,
        end: 0x01D546,
    },
    Range {
        start: 0x01D54A,
        end: 0x01D550,
    },
    Range {
        start: 0x01D552,
        end: 0x01D6A5,
    },
    Range {
        start: 0x01D6A8,
        end: 0x01D6C0,
    },
    Range {
        start: 0x01D6C2,
        end: 0x01D6DA,
    },
    Range {
        start: 0x01D6DC,
        end: 0x01D6FA,
    },
    Range {
        start: 0x01D6FC,
        end: 0x01D714,
    },
    Range {
        start: 0x01D716,
        end: 0x01D734,
    },
    Range {
        start: 0x01D736,
        end: 0x01D74E,
    },
    Range {
        start: 0x01D750,
        end: 0x01D76E,
    },
    Range {
        start: 0x01D770,
        end: 0x01D788,
    },
    Range {
        start: 0x01D78A,
        end: 0x01D7A8,
    },
    Range {
        start: 0x01D7AA,
        end: 0x01D7C2,
    },
    Range {
        start: 0x01D7C4,
        end: 0x01D7CB,
    },
    Range {
        start: 0x01DF00,
        end: 0x01DF1E,
    },
    Range {
        start: 0x01DF25,
        end: 0x01DF2A,
    },
    Range {
        start: 0x01E030,
        end: 0x01E06D,
    },
    Range {
        start: 0x01E100,
        end: 0x01E12C,
    },
    Range {
        start: 0x01E137,
        end: 0x01E13D,
    },
    Range {
        start: 0x01E14E,
        end: 0x01E14E,
    },
    Range {
        start: 0x01E290,
        end: 0x01E2AD,
    },
    Range {
        start: 0x01E2C0,
        end: 0x01E2EB,
    },
    Range {
        start: 0x01E4D0,
        end: 0x01E4EB,
    },
    Range {
        start: 0x01E5D0,
        end: 0x01E5ED,
    },
    Range {
        start: 0x01E5F0,
        end: 0x01E5F0,
    },
    Range {
        start: 0x01E6C0,
        end: 0x01E6DE,
    },
    Range {
        start: 0x01E6E0,
        end: 0x01E6E2,
    },
    Range {
        start: 0x01E6E4,
        end: 0x01E6E5,
    },
    Range {
        start: 0x01E6E7,
        end: 0x01E6ED,
    },
    Range {
        start: 0x01E6F0,
        end: 0x01E6F4,
    },
    Range {
        start: 0x01E6FE,
        end: 0x01E6FF,
    },
    Range {
        start: 0x01E7E0,
        end: 0x01E7E6,
    },
    Range {
        start: 0x01E7E8,
        end: 0x01E7EB,
    },
    Range {
        start: 0x01E7ED,
        end: 0x01E7EE,
    },
    Range {
        start: 0x01E7F0,
        end: 0x01E7FE,
    },
    Range {
        start: 0x01E800,
        end: 0x01E8C4,
    },
    Range {
        start: 0x01E900,
        end: 0x01E943,
    },
    Range {
        start: 0x01E94B,
        end: 0x01E94B,
    },
    Range {
        start: 0x01EE00,
        end: 0x01EE03,
    },
    Range {
        start: 0x01EE05,
        end: 0x01EE1F,
    },
    Range {
        start: 0x01EE21,
        end: 0x01EE22,
    },
    Range {
        start: 0x01EE24,
        end: 0x01EE24,
    },
    Range {
        start: 0x01EE27,
        end: 0x01EE27,
    },
    Range {
        start: 0x01EE29,
        end: 0x01EE32,
    },
    Range {
        start: 0x01EE34,
        end: 0x01EE37,
    },
    Range {
        start: 0x01EE39,
        end: 0x01EE39,
    },
    Range {
        start: 0x01EE3B,
        end: 0x01EE3B,
    },
    Range {
        start: 0x01EE42,
        end: 0x01EE42,
    },
    Range {
        start: 0x01EE47,
        end: 0x01EE47,
    },
    Range {
        start: 0x01EE49,
        end: 0x01EE49,
    },
    Range {
        start: 0x01EE4B,
        end: 0x01EE4B,
    },
    Range {
        start: 0x01EE4D,
        end: 0x01EE4F,
    },
    Range {
        start: 0x01EE51,
        end: 0x01EE52,
    },
    Range {
        start: 0x01EE54,
        end: 0x01EE54,
    },
    Range {
        start: 0x01EE57,
        end: 0x01EE57,
    },
    Range {
        start: 0x01EE59,
        end: 0x01EE59,
    },
    Range {
        start: 0x01EE5B,
        end: 0x01EE5B,
    },
    Range {
        start: 0x01EE5D,
        end: 0x01EE5D,
    },
    Range {
        start: 0x01EE5F,
        end: 0x01EE5F,
    },
    Range {
        start: 0x01EE61,
        end: 0x01EE62,
    },
    Range {
        start: 0x01EE64,
        end: 0x01EE64,
    },
    Range {
        start: 0x01EE67,
        end: 0x01EE6A,
    },
    Range {
        start: 0x01EE6C,
        end: 0x01EE72,
    },
    Range {
        start: 0x01EE74,
        end: 0x01EE77,
    },
    Range {
        start: 0x01EE79,
        end: 0x01EE7C,
    },
    Range {
        start: 0x01EE7E,
        end: 0x01EE7E,
    },
    Range {
        start: 0x01EE80,
        end: 0x01EE89,
    },
    Range {
        start: 0x01EE8B,
        end: 0x01EE9B,
    },
    Range {
        start: 0x01EEA1,
        end: 0x01EEA3,
    },
    Range {
        start: 0x01EEA5,
        end: 0x01EEA9,
    },
    Range {
        start: 0x01EEAB,
        end: 0x01EEBB,
    },
    Range {
        start: 0x020000,
        end: 0x02A6DF,
    },
    Range {
        start: 0x02A700,
        end: 0x02B81D,
    },
    Range {
        start: 0x02B820,
        end: 0x02CEAD,
    },
    Range {
        start: 0x02CEB0,
        end: 0x02EBE0,
    },
    Range {
        start: 0x02EBF0,
        end: 0x02EE5D,
    },
    Range {
        start: 0x02F800,
        end: 0x02FA1D,
    },
    Range {
        start: 0x030000,
        end: 0x03134A,
    },
    Range {
        start: 0x031350,
        end: 0x033479,
    },
];

const CASED: &[Range] = &[
    Range {
        start: 0x000041,
        end: 0x00005A,
    },
    Range {
        start: 0x000061,
        end: 0x00007A,
    },
    Range {
        start: 0x0000AA,
        end: 0x0000AA,
    },
    Range {
        start: 0x0000B5,
        end: 0x0000B5,
    },
    Range {
        start: 0x0000BA,
        end: 0x0000BA,
    },
    Range {
        start: 0x0000C0,
        end: 0x0000D6,
    },
    Range {
        start: 0x0000D8,
        end: 0x0000F6,
    },
    Range {
        start: 0x0000F8,
        end: 0x0001BA,
    },
    Range {
        start: 0x0001BC,
        end: 0x0001BF,
    },
    Range {
        start: 0x0001C4,
        end: 0x000293,
    },
    Range {
        start: 0x000296,
        end: 0x0002B8,
    },
    Range {
        start: 0x0002C0,
        end: 0x0002C1,
    },
    Range {
        start: 0x0002E0,
        end: 0x0002E4,
    },
    Range {
        start: 0x000345,
        end: 0x000345,
    },
    Range {
        start: 0x000370,
        end: 0x000373,
    },
    Range {
        start: 0x000376,
        end: 0x000377,
    },
    Range {
        start: 0x00037A,
        end: 0x00037D,
    },
    Range {
        start: 0x00037F,
        end: 0x00037F,
    },
    Range {
        start: 0x000386,
        end: 0x000386,
    },
    Range {
        start: 0x000388,
        end: 0x00038A,
    },
    Range {
        start: 0x00038C,
        end: 0x00038C,
    },
    Range {
        start: 0x00038E,
        end: 0x0003A1,
    },
    Range {
        start: 0x0003A3,
        end: 0x0003F5,
    },
    Range {
        start: 0x0003F7,
        end: 0x000481,
    },
    Range {
        start: 0x00048A,
        end: 0x00052F,
    },
    Range {
        start: 0x000531,
        end: 0x000556,
    },
    Range {
        start: 0x000560,
        end: 0x000588,
    },
    Range {
        start: 0x0010A0,
        end: 0x0010C5,
    },
    Range {
        start: 0x0010C7,
        end: 0x0010C7,
    },
    Range {
        start: 0x0010CD,
        end: 0x0010CD,
    },
    Range {
        start: 0x0010D0,
        end: 0x0010FA,
    },
    Range {
        start: 0x0010FC,
        end: 0x0010FF,
    },
    Range {
        start: 0x0013A0,
        end: 0x0013F5,
    },
    Range {
        start: 0x0013F8,
        end: 0x0013FD,
    },
    Range {
        start: 0x001C80,
        end: 0x001C8A,
    },
    Range {
        start: 0x001C90,
        end: 0x001CBA,
    },
    Range {
        start: 0x001CBD,
        end: 0x001CBF,
    },
    Range {
        start: 0x001D00,
        end: 0x001DBF,
    },
    Range {
        start: 0x001E00,
        end: 0x001F15,
    },
    Range {
        start: 0x001F18,
        end: 0x001F1D,
    },
    Range {
        start: 0x001F20,
        end: 0x001F45,
    },
    Range {
        start: 0x001F48,
        end: 0x001F4D,
    },
    Range {
        start: 0x001F50,
        end: 0x001F57,
    },
    Range {
        start: 0x001F59,
        end: 0x001F59,
    },
    Range {
        start: 0x001F5B,
        end: 0x001F5B,
    },
    Range {
        start: 0x001F5D,
        end: 0x001F5D,
    },
    Range {
        start: 0x001F5F,
        end: 0x001F7D,
    },
    Range {
        start: 0x001F80,
        end: 0x001FB4,
    },
    Range {
        start: 0x001FB6,
        end: 0x001FBC,
    },
    Range {
        start: 0x001FBE,
        end: 0x001FBE,
    },
    Range {
        start: 0x001FC2,
        end: 0x001FC4,
    },
    Range {
        start: 0x001FC6,
        end: 0x001FCC,
    },
    Range {
        start: 0x001FD0,
        end: 0x001FD3,
    },
    Range {
        start: 0x001FD6,
        end: 0x001FDB,
    },
    Range {
        start: 0x001FE0,
        end: 0x001FEC,
    },
    Range {
        start: 0x001FF2,
        end: 0x001FF4,
    },
    Range {
        start: 0x001FF6,
        end: 0x001FFC,
    },
    Range {
        start: 0x002071,
        end: 0x002071,
    },
    Range {
        start: 0x00207F,
        end: 0x00207F,
    },
    Range {
        start: 0x002090,
        end: 0x00209C,
    },
    Range {
        start: 0x002102,
        end: 0x002102,
    },
    Range {
        start: 0x002107,
        end: 0x002107,
    },
    Range {
        start: 0x00210A,
        end: 0x002113,
    },
    Range {
        start: 0x002115,
        end: 0x002115,
    },
    Range {
        start: 0x002119,
        end: 0x00211D,
    },
    Range {
        start: 0x002124,
        end: 0x002124,
    },
    Range {
        start: 0x002126,
        end: 0x002126,
    },
    Range {
        start: 0x002128,
        end: 0x002128,
    },
    Range {
        start: 0x00212A,
        end: 0x00212D,
    },
    Range {
        start: 0x00212F,
        end: 0x002134,
    },
    Range {
        start: 0x002139,
        end: 0x002139,
    },
    Range {
        start: 0x00213C,
        end: 0x00213F,
    },
    Range {
        start: 0x002145,
        end: 0x002149,
    },
    Range {
        start: 0x00214E,
        end: 0x00214E,
    },
    Range {
        start: 0x002160,
        end: 0x00217F,
    },
    Range {
        start: 0x002183,
        end: 0x002184,
    },
    Range {
        start: 0x0024B6,
        end: 0x0024E9,
    },
    Range {
        start: 0x002C00,
        end: 0x002CE4,
    },
    Range {
        start: 0x002CEB,
        end: 0x002CEE,
    },
    Range {
        start: 0x002CF2,
        end: 0x002CF3,
    },
    Range {
        start: 0x002D00,
        end: 0x002D25,
    },
    Range {
        start: 0x002D27,
        end: 0x002D27,
    },
    Range {
        start: 0x002D2D,
        end: 0x002D2D,
    },
    Range {
        start: 0x00A640,
        end: 0x00A66D,
    },
    Range {
        start: 0x00A680,
        end: 0x00A69D,
    },
    Range {
        start: 0x00A722,
        end: 0x00A787,
    },
    Range {
        start: 0x00A78B,
        end: 0x00A78E,
    },
    Range {
        start: 0x00A790,
        end: 0x00A7DC,
    },
    Range {
        start: 0x00A7F1,
        end: 0x00A7F6,
    },
    Range {
        start: 0x00A7F8,
        end: 0x00A7FA,
    },
    Range {
        start: 0x00AB30,
        end: 0x00AB5A,
    },
    Range {
        start: 0x00AB5C,
        end: 0x00AB69,
    },
    Range {
        start: 0x00AB70,
        end: 0x00ABBF,
    },
    Range {
        start: 0x00FB00,
        end: 0x00FB06,
    },
    Range {
        start: 0x00FB13,
        end: 0x00FB17,
    },
    Range {
        start: 0x00FF21,
        end: 0x00FF3A,
    },
    Range {
        start: 0x00FF41,
        end: 0x00FF5A,
    },
    Range {
        start: 0x010400,
        end: 0x01044F,
    },
    Range {
        start: 0x0104B0,
        end: 0x0104D3,
    },
    Range {
        start: 0x0104D8,
        end: 0x0104FB,
    },
    Range {
        start: 0x010570,
        end: 0x01057A,
    },
    Range {
        start: 0x01057C,
        end: 0x01058A,
    },
    Range {
        start: 0x01058C,
        end: 0x010592,
    },
    Range {
        start: 0x010594,
        end: 0x010595,
    },
    Range {
        start: 0x010597,
        end: 0x0105A1,
    },
    Range {
        start: 0x0105A3,
        end: 0x0105B1,
    },
    Range {
        start: 0x0105B3,
        end: 0x0105B9,
    },
    Range {
        start: 0x0105BB,
        end: 0x0105BC,
    },
    Range {
        start: 0x010780,
        end: 0x010780,
    },
    Range {
        start: 0x010783,
        end: 0x010785,
    },
    Range {
        start: 0x010787,
        end: 0x0107B0,
    },
    Range {
        start: 0x0107B2,
        end: 0x0107BA,
    },
    Range {
        start: 0x010C80,
        end: 0x010CB2,
    },
    Range {
        start: 0x010CC0,
        end: 0x010CF2,
    },
    Range {
        start: 0x010D50,
        end: 0x010D65,
    },
    Range {
        start: 0x010D70,
        end: 0x010D85,
    },
    Range {
        start: 0x0118A0,
        end: 0x0118DF,
    },
    Range {
        start: 0x016E40,
        end: 0x016E7F,
    },
    Range {
        start: 0x016EA0,
        end: 0x016EB8,
    },
    Range {
        start: 0x016EBB,
        end: 0x016ED3,
    },
    Range {
        start: 0x01D400,
        end: 0x01D454,
    },
    Range {
        start: 0x01D456,
        end: 0x01D49C,
    },
    Range {
        start: 0x01D49E,
        end: 0x01D49F,
    },
    Range {
        start: 0x01D4A2,
        end: 0x01D4A2,
    },
    Range {
        start: 0x01D4A5,
        end: 0x01D4A6,
    },
    Range {
        start: 0x01D4A9,
        end: 0x01D4AC,
    },
    Range {
        start: 0x01D4AE,
        end: 0x01D4B9,
    },
    Range {
        start: 0x01D4BB,
        end: 0x01D4BB,
    },
    Range {
        start: 0x01D4BD,
        end: 0x01D4C3,
    },
    Range {
        start: 0x01D4C5,
        end: 0x01D505,
    },
    Range {
        start: 0x01D507,
        end: 0x01D50A,
    },
    Range {
        start: 0x01D50D,
        end: 0x01D514,
    },
    Range {
        start: 0x01D516,
        end: 0x01D51C,
    },
    Range {
        start: 0x01D51E,
        end: 0x01D539,
    },
    Range {
        start: 0x01D53B,
        end: 0x01D53E,
    },
    Range {
        start: 0x01D540,
        end: 0x01D544,
    },
    Range {
        start: 0x01D546,
        end: 0x01D546,
    },
    Range {
        start: 0x01D54A,
        end: 0x01D550,
    },
    Range {
        start: 0x01D552,
        end: 0x01D6A5,
    },
    Range {
        start: 0x01D6A8,
        end: 0x01D6C0,
    },
    Range {
        start: 0x01D6C2,
        end: 0x01D6DA,
    },
    Range {
        start: 0x01D6DC,
        end: 0x01D6FA,
    },
    Range {
        start: 0x01D6FC,
        end: 0x01D714,
    },
    Range {
        start: 0x01D716,
        end: 0x01D734,
    },
    Range {
        start: 0x01D736,
        end: 0x01D74E,
    },
    Range {
        start: 0x01D750,
        end: 0x01D76E,
    },
    Range {
        start: 0x01D770,
        end: 0x01D788,
    },
    Range {
        start: 0x01D78A,
        end: 0x01D7A8,
    },
    Range {
        start: 0x01D7AA,
        end: 0x01D7C2,
    },
    Range {
        start: 0x01D7C4,
        end: 0x01D7CB,
    },
    Range {
        start: 0x01DF00,
        end: 0x01DF09,
    },
    Range {
        start: 0x01DF0B,
        end: 0x01DF1E,
    },
    Range {
        start: 0x01DF25,
        end: 0x01DF2A,
    },
    Range {
        start: 0x01E030,
        end: 0x01E06D,
    },
    Range {
        start: 0x01E900,
        end: 0x01E943,
    },
    Range {
        start: 0x01F130,
        end: 0x01F149,
    },
    Range {
        start: 0x01F150,
        end: 0x01F169,
    },
    Range {
        start: 0x01F170,
        end: 0x01F189,
    },
];

const CASE_IGNORABLE: &[Range] = &[
    Range {
        start: 0x000027,
        end: 0x000027,
    },
    Range {
        start: 0x00002E,
        end: 0x00002E,
    },
    Range {
        start: 0x00003A,
        end: 0x00003A,
    },
    Range {
        start: 0x00005E,
        end: 0x00005E,
    },
    Range {
        start: 0x000060,
        end: 0x000060,
    },
    Range {
        start: 0x0000A8,
        end: 0x0000A8,
    },
    Range {
        start: 0x0000AD,
        end: 0x0000AD,
    },
    Range {
        start: 0x0000AF,
        end: 0x0000AF,
    },
    Range {
        start: 0x0000B4,
        end: 0x0000B4,
    },
    Range {
        start: 0x0000B7,
        end: 0x0000B8,
    },
    Range {
        start: 0x0002B0,
        end: 0x00036F,
    },
    Range {
        start: 0x000374,
        end: 0x000375,
    },
    Range {
        start: 0x00037A,
        end: 0x00037A,
    },
    Range {
        start: 0x000384,
        end: 0x000385,
    },
    Range {
        start: 0x000387,
        end: 0x000387,
    },
    Range {
        start: 0x000483,
        end: 0x000489,
    },
    Range {
        start: 0x000559,
        end: 0x000559,
    },
    Range {
        start: 0x00055F,
        end: 0x00055F,
    },
    Range {
        start: 0x000591,
        end: 0x0005BD,
    },
    Range {
        start: 0x0005BF,
        end: 0x0005BF,
    },
    Range {
        start: 0x0005C1,
        end: 0x0005C2,
    },
    Range {
        start: 0x0005C4,
        end: 0x0005C5,
    },
    Range {
        start: 0x0005C7,
        end: 0x0005C7,
    },
    Range {
        start: 0x0005F4,
        end: 0x0005F4,
    },
    Range {
        start: 0x000600,
        end: 0x000605,
    },
    Range {
        start: 0x000610,
        end: 0x00061A,
    },
    Range {
        start: 0x00061C,
        end: 0x00061C,
    },
    Range {
        start: 0x000640,
        end: 0x000640,
    },
    Range {
        start: 0x00064B,
        end: 0x00065F,
    },
    Range {
        start: 0x000670,
        end: 0x000670,
    },
    Range {
        start: 0x0006D6,
        end: 0x0006DD,
    },
    Range {
        start: 0x0006DF,
        end: 0x0006E8,
    },
    Range {
        start: 0x0006EA,
        end: 0x0006ED,
    },
    Range {
        start: 0x00070F,
        end: 0x00070F,
    },
    Range {
        start: 0x000711,
        end: 0x000711,
    },
    Range {
        start: 0x000730,
        end: 0x00074A,
    },
    Range {
        start: 0x0007A6,
        end: 0x0007B0,
    },
    Range {
        start: 0x0007EB,
        end: 0x0007F5,
    },
    Range {
        start: 0x0007FA,
        end: 0x0007FA,
    },
    Range {
        start: 0x0007FD,
        end: 0x0007FD,
    },
    Range {
        start: 0x000816,
        end: 0x00082D,
    },
    Range {
        start: 0x000859,
        end: 0x00085B,
    },
    Range {
        start: 0x000888,
        end: 0x000888,
    },
    Range {
        start: 0x000890,
        end: 0x000891,
    },
    Range {
        start: 0x000897,
        end: 0x00089F,
    },
    Range {
        start: 0x0008C9,
        end: 0x000902,
    },
    Range {
        start: 0x00093A,
        end: 0x00093A,
    },
    Range {
        start: 0x00093C,
        end: 0x00093C,
    },
    Range {
        start: 0x000941,
        end: 0x000948,
    },
    Range {
        start: 0x00094D,
        end: 0x00094D,
    },
    Range {
        start: 0x000951,
        end: 0x000957,
    },
    Range {
        start: 0x000962,
        end: 0x000963,
    },
    Range {
        start: 0x000971,
        end: 0x000971,
    },
    Range {
        start: 0x000981,
        end: 0x000981,
    },
    Range {
        start: 0x0009BC,
        end: 0x0009BC,
    },
    Range {
        start: 0x0009C1,
        end: 0x0009C4,
    },
    Range {
        start: 0x0009CD,
        end: 0x0009CD,
    },
    Range {
        start: 0x0009E2,
        end: 0x0009E3,
    },
    Range {
        start: 0x0009FE,
        end: 0x0009FE,
    },
    Range {
        start: 0x000A01,
        end: 0x000A02,
    },
    Range {
        start: 0x000A3C,
        end: 0x000A3C,
    },
    Range {
        start: 0x000A41,
        end: 0x000A42,
    },
    Range {
        start: 0x000A47,
        end: 0x000A48,
    },
    Range {
        start: 0x000A4B,
        end: 0x000A4D,
    },
    Range {
        start: 0x000A51,
        end: 0x000A51,
    },
    Range {
        start: 0x000A70,
        end: 0x000A71,
    },
    Range {
        start: 0x000A75,
        end: 0x000A75,
    },
    Range {
        start: 0x000A81,
        end: 0x000A82,
    },
    Range {
        start: 0x000ABC,
        end: 0x000ABC,
    },
    Range {
        start: 0x000AC1,
        end: 0x000AC5,
    },
    Range {
        start: 0x000AC7,
        end: 0x000AC8,
    },
    Range {
        start: 0x000ACD,
        end: 0x000ACD,
    },
    Range {
        start: 0x000AE2,
        end: 0x000AE3,
    },
    Range {
        start: 0x000AFA,
        end: 0x000AFF,
    },
    Range {
        start: 0x000B01,
        end: 0x000B01,
    },
    Range {
        start: 0x000B3C,
        end: 0x000B3C,
    },
    Range {
        start: 0x000B3F,
        end: 0x000B3F,
    },
    Range {
        start: 0x000B41,
        end: 0x000B44,
    },
    Range {
        start: 0x000B4D,
        end: 0x000B4D,
    },
    Range {
        start: 0x000B55,
        end: 0x000B56,
    },
    Range {
        start: 0x000B62,
        end: 0x000B63,
    },
    Range {
        start: 0x000B82,
        end: 0x000B82,
    },
    Range {
        start: 0x000BC0,
        end: 0x000BC0,
    },
    Range {
        start: 0x000BCD,
        end: 0x000BCD,
    },
    Range {
        start: 0x000C00,
        end: 0x000C00,
    },
    Range {
        start: 0x000C04,
        end: 0x000C04,
    },
    Range {
        start: 0x000C3C,
        end: 0x000C3C,
    },
    Range {
        start: 0x000C3E,
        end: 0x000C40,
    },
    Range {
        start: 0x000C46,
        end: 0x000C48,
    },
    Range {
        start: 0x000C4A,
        end: 0x000C4D,
    },
    Range {
        start: 0x000C55,
        end: 0x000C56,
    },
    Range {
        start: 0x000C62,
        end: 0x000C63,
    },
    Range {
        start: 0x000C81,
        end: 0x000C81,
    },
    Range {
        start: 0x000CBC,
        end: 0x000CBC,
    },
    Range {
        start: 0x000CBF,
        end: 0x000CBF,
    },
    Range {
        start: 0x000CC6,
        end: 0x000CC6,
    },
    Range {
        start: 0x000CCC,
        end: 0x000CCD,
    },
    Range {
        start: 0x000CE2,
        end: 0x000CE3,
    },
    Range {
        start: 0x000D00,
        end: 0x000D01,
    },
    Range {
        start: 0x000D3B,
        end: 0x000D3C,
    },
    Range {
        start: 0x000D41,
        end: 0x000D44,
    },
    Range {
        start: 0x000D4D,
        end: 0x000D4D,
    },
    Range {
        start: 0x000D62,
        end: 0x000D63,
    },
    Range {
        start: 0x000D81,
        end: 0x000D81,
    },
    Range {
        start: 0x000DCA,
        end: 0x000DCA,
    },
    Range {
        start: 0x000DD2,
        end: 0x000DD4,
    },
    Range {
        start: 0x000DD6,
        end: 0x000DD6,
    },
    Range {
        start: 0x000E31,
        end: 0x000E31,
    },
    Range {
        start: 0x000E34,
        end: 0x000E3A,
    },
    Range {
        start: 0x000E46,
        end: 0x000E4E,
    },
    Range {
        start: 0x000EB1,
        end: 0x000EB1,
    },
    Range {
        start: 0x000EB4,
        end: 0x000EBC,
    },
    Range {
        start: 0x000EC6,
        end: 0x000EC6,
    },
    Range {
        start: 0x000EC8,
        end: 0x000ECE,
    },
    Range {
        start: 0x000F18,
        end: 0x000F19,
    },
    Range {
        start: 0x000F35,
        end: 0x000F35,
    },
    Range {
        start: 0x000F37,
        end: 0x000F37,
    },
    Range {
        start: 0x000F39,
        end: 0x000F39,
    },
    Range {
        start: 0x000F71,
        end: 0x000F7E,
    },
    Range {
        start: 0x000F80,
        end: 0x000F84,
    },
    Range {
        start: 0x000F86,
        end: 0x000F87,
    },
    Range {
        start: 0x000F8D,
        end: 0x000F97,
    },
    Range {
        start: 0x000F99,
        end: 0x000FBC,
    },
    Range {
        start: 0x000FC6,
        end: 0x000FC6,
    },
    Range {
        start: 0x00102D,
        end: 0x001030,
    },
    Range {
        start: 0x001032,
        end: 0x001037,
    },
    Range {
        start: 0x001039,
        end: 0x00103A,
    },
    Range {
        start: 0x00103D,
        end: 0x00103E,
    },
    Range {
        start: 0x001058,
        end: 0x001059,
    },
    Range {
        start: 0x00105E,
        end: 0x001060,
    },
    Range {
        start: 0x001071,
        end: 0x001074,
    },
    Range {
        start: 0x001082,
        end: 0x001082,
    },
    Range {
        start: 0x001085,
        end: 0x001086,
    },
    Range {
        start: 0x00108D,
        end: 0x00108D,
    },
    Range {
        start: 0x00109D,
        end: 0x00109D,
    },
    Range {
        start: 0x0010FC,
        end: 0x0010FC,
    },
    Range {
        start: 0x00135D,
        end: 0x00135F,
    },
    Range {
        start: 0x001712,
        end: 0x001714,
    },
    Range {
        start: 0x001732,
        end: 0x001733,
    },
    Range {
        start: 0x001752,
        end: 0x001753,
    },
    Range {
        start: 0x001772,
        end: 0x001773,
    },
    Range {
        start: 0x0017B4,
        end: 0x0017B5,
    },
    Range {
        start: 0x0017B7,
        end: 0x0017BD,
    },
    Range {
        start: 0x0017C6,
        end: 0x0017C6,
    },
    Range {
        start: 0x0017C9,
        end: 0x0017D3,
    },
    Range {
        start: 0x0017D7,
        end: 0x0017D7,
    },
    Range {
        start: 0x0017DD,
        end: 0x0017DD,
    },
    Range {
        start: 0x00180B,
        end: 0x00180F,
    },
    Range {
        start: 0x001843,
        end: 0x001843,
    },
    Range {
        start: 0x001885,
        end: 0x001886,
    },
    Range {
        start: 0x0018A9,
        end: 0x0018A9,
    },
    Range {
        start: 0x001920,
        end: 0x001922,
    },
    Range {
        start: 0x001927,
        end: 0x001928,
    },
    Range {
        start: 0x001932,
        end: 0x001932,
    },
    Range {
        start: 0x001939,
        end: 0x00193B,
    },
    Range {
        start: 0x001A17,
        end: 0x001A18,
    },
    Range {
        start: 0x001A1B,
        end: 0x001A1B,
    },
    Range {
        start: 0x001A56,
        end: 0x001A56,
    },
    Range {
        start: 0x001A58,
        end: 0x001A5E,
    },
    Range {
        start: 0x001A60,
        end: 0x001A60,
    },
    Range {
        start: 0x001A62,
        end: 0x001A62,
    },
    Range {
        start: 0x001A65,
        end: 0x001A6C,
    },
    Range {
        start: 0x001A73,
        end: 0x001A7C,
    },
    Range {
        start: 0x001A7F,
        end: 0x001A7F,
    },
    Range {
        start: 0x001AA7,
        end: 0x001AA7,
    },
    Range {
        start: 0x001AB0,
        end: 0x001ADD,
    },
    Range {
        start: 0x001AE0,
        end: 0x001AEB,
    },
    Range {
        start: 0x001B00,
        end: 0x001B03,
    },
    Range {
        start: 0x001B34,
        end: 0x001B34,
    },
    Range {
        start: 0x001B36,
        end: 0x001B3A,
    },
    Range {
        start: 0x001B3C,
        end: 0x001B3C,
    },
    Range {
        start: 0x001B42,
        end: 0x001B42,
    },
    Range {
        start: 0x001B6B,
        end: 0x001B73,
    },
    Range {
        start: 0x001B80,
        end: 0x001B81,
    },
    Range {
        start: 0x001BA2,
        end: 0x001BA5,
    },
    Range {
        start: 0x001BA8,
        end: 0x001BA9,
    },
    Range {
        start: 0x001BAB,
        end: 0x001BAD,
    },
    Range {
        start: 0x001BE6,
        end: 0x001BE6,
    },
    Range {
        start: 0x001BE8,
        end: 0x001BE9,
    },
    Range {
        start: 0x001BED,
        end: 0x001BED,
    },
    Range {
        start: 0x001BEF,
        end: 0x001BF1,
    },
    Range {
        start: 0x001C2C,
        end: 0x001C33,
    },
    Range {
        start: 0x001C36,
        end: 0x001C37,
    },
    Range {
        start: 0x001C78,
        end: 0x001C7D,
    },
    Range {
        start: 0x001CD0,
        end: 0x001CD2,
    },
    Range {
        start: 0x001CD4,
        end: 0x001CE0,
    },
    Range {
        start: 0x001CE2,
        end: 0x001CE8,
    },
    Range {
        start: 0x001CED,
        end: 0x001CED,
    },
    Range {
        start: 0x001CF4,
        end: 0x001CF4,
    },
    Range {
        start: 0x001CF8,
        end: 0x001CF9,
    },
    Range {
        start: 0x001D2C,
        end: 0x001D6A,
    },
    Range {
        start: 0x001D78,
        end: 0x001D78,
    },
    Range {
        start: 0x001D9B,
        end: 0x001DFF,
    },
    Range {
        start: 0x001FBD,
        end: 0x001FBD,
    },
    Range {
        start: 0x001FBF,
        end: 0x001FC1,
    },
    Range {
        start: 0x001FCD,
        end: 0x001FCF,
    },
    Range {
        start: 0x001FDD,
        end: 0x001FDF,
    },
    Range {
        start: 0x001FED,
        end: 0x001FEF,
    },
    Range {
        start: 0x001FFD,
        end: 0x001FFE,
    },
    Range {
        start: 0x00200B,
        end: 0x00200F,
    },
    Range {
        start: 0x002018,
        end: 0x002019,
    },
    Range {
        start: 0x002024,
        end: 0x002024,
    },
    Range {
        start: 0x002027,
        end: 0x002027,
    },
    Range {
        start: 0x00202A,
        end: 0x00202E,
    },
    Range {
        start: 0x002060,
        end: 0x002064,
    },
    Range {
        start: 0x002066,
        end: 0x00206F,
    },
    Range {
        start: 0x002071,
        end: 0x002071,
    },
    Range {
        start: 0x00207F,
        end: 0x00207F,
    },
    Range {
        start: 0x002090,
        end: 0x00209C,
    },
    Range {
        start: 0x0020D0,
        end: 0x0020F0,
    },
    Range {
        start: 0x002C7C,
        end: 0x002C7D,
    },
    Range {
        start: 0x002CEF,
        end: 0x002CF1,
    },
    Range {
        start: 0x002D6F,
        end: 0x002D6F,
    },
    Range {
        start: 0x002D7F,
        end: 0x002D7F,
    },
    Range {
        start: 0x002DE0,
        end: 0x002DFF,
    },
    Range {
        start: 0x002E2F,
        end: 0x002E2F,
    },
    Range {
        start: 0x003005,
        end: 0x003005,
    },
    Range {
        start: 0x00302A,
        end: 0x00302D,
    },
    Range {
        start: 0x003031,
        end: 0x003035,
    },
    Range {
        start: 0x00303B,
        end: 0x00303B,
    },
    Range {
        start: 0x003099,
        end: 0x00309E,
    },
    Range {
        start: 0x0030FC,
        end: 0x0030FE,
    },
    Range {
        start: 0x00A015,
        end: 0x00A015,
    },
    Range {
        start: 0x00A4F8,
        end: 0x00A4FD,
    },
    Range {
        start: 0x00A60C,
        end: 0x00A60C,
    },
    Range {
        start: 0x00A66F,
        end: 0x00A672,
    },
    Range {
        start: 0x00A674,
        end: 0x00A67D,
    },
    Range {
        start: 0x00A67F,
        end: 0x00A67F,
    },
    Range {
        start: 0x00A69C,
        end: 0x00A69F,
    },
    Range {
        start: 0x00A6F0,
        end: 0x00A6F1,
    },
    Range {
        start: 0x00A700,
        end: 0x00A721,
    },
    Range {
        start: 0x00A770,
        end: 0x00A770,
    },
    Range {
        start: 0x00A788,
        end: 0x00A78A,
    },
    Range {
        start: 0x00A7F1,
        end: 0x00A7F4,
    },
    Range {
        start: 0x00A7F8,
        end: 0x00A7F9,
    },
    Range {
        start: 0x00A802,
        end: 0x00A802,
    },
    Range {
        start: 0x00A806,
        end: 0x00A806,
    },
    Range {
        start: 0x00A80B,
        end: 0x00A80B,
    },
    Range {
        start: 0x00A825,
        end: 0x00A826,
    },
    Range {
        start: 0x00A82C,
        end: 0x00A82C,
    },
    Range {
        start: 0x00A8C4,
        end: 0x00A8C5,
    },
    Range {
        start: 0x00A8E0,
        end: 0x00A8F1,
    },
    Range {
        start: 0x00A8FF,
        end: 0x00A8FF,
    },
    Range {
        start: 0x00A926,
        end: 0x00A92D,
    },
    Range {
        start: 0x00A947,
        end: 0x00A951,
    },
    Range {
        start: 0x00A980,
        end: 0x00A982,
    },
    Range {
        start: 0x00A9B3,
        end: 0x00A9B3,
    },
    Range {
        start: 0x00A9B6,
        end: 0x00A9B9,
    },
    Range {
        start: 0x00A9BC,
        end: 0x00A9BD,
    },
    Range {
        start: 0x00A9CF,
        end: 0x00A9CF,
    },
    Range {
        start: 0x00A9E5,
        end: 0x00A9E6,
    },
    Range {
        start: 0x00AA29,
        end: 0x00AA2E,
    },
    Range {
        start: 0x00AA31,
        end: 0x00AA32,
    },
    Range {
        start: 0x00AA35,
        end: 0x00AA36,
    },
    Range {
        start: 0x00AA43,
        end: 0x00AA43,
    },
    Range {
        start: 0x00AA4C,
        end: 0x00AA4C,
    },
    Range {
        start: 0x00AA70,
        end: 0x00AA70,
    },
    Range {
        start: 0x00AA7C,
        end: 0x00AA7C,
    },
    Range {
        start: 0x00AAB0,
        end: 0x00AAB0,
    },
    Range {
        start: 0x00AAB2,
        end: 0x00AAB4,
    },
    Range {
        start: 0x00AAB7,
        end: 0x00AAB8,
    },
    Range {
        start: 0x00AABE,
        end: 0x00AABF,
    },
    Range {
        start: 0x00AAC1,
        end: 0x00AAC1,
    },
    Range {
        start: 0x00AADD,
        end: 0x00AADD,
    },
    Range {
        start: 0x00AAEC,
        end: 0x00AAED,
    },
    Range {
        start: 0x00AAF3,
        end: 0x00AAF4,
    },
    Range {
        start: 0x00AAF6,
        end: 0x00AAF6,
    },
    Range {
        start: 0x00AB5B,
        end: 0x00AB5F,
    },
    Range {
        start: 0x00AB69,
        end: 0x00AB6B,
    },
    Range {
        start: 0x00ABE5,
        end: 0x00ABE5,
    },
    Range {
        start: 0x00ABE8,
        end: 0x00ABE8,
    },
    Range {
        start: 0x00ABED,
        end: 0x00ABED,
    },
    Range {
        start: 0x00FB1E,
        end: 0x00FB1E,
    },
    Range {
        start: 0x00FBB2,
        end: 0x00FBC2,
    },
    Range {
        start: 0x00FE00,
        end: 0x00FE0F,
    },
    Range {
        start: 0x00FE13,
        end: 0x00FE13,
    },
    Range {
        start: 0x00FE20,
        end: 0x00FE2F,
    },
    Range {
        start: 0x00FE52,
        end: 0x00FE52,
    },
    Range {
        start: 0x00FE55,
        end: 0x00FE55,
    },
    Range {
        start: 0x00FEFF,
        end: 0x00FEFF,
    },
    Range {
        start: 0x00FF07,
        end: 0x00FF07,
    },
    Range {
        start: 0x00FF0E,
        end: 0x00FF0E,
    },
    Range {
        start: 0x00FF1A,
        end: 0x00FF1A,
    },
    Range {
        start: 0x00FF3E,
        end: 0x00FF3E,
    },
    Range {
        start: 0x00FF40,
        end: 0x00FF40,
    },
    Range {
        start: 0x00FF70,
        end: 0x00FF70,
    },
    Range {
        start: 0x00FF9E,
        end: 0x00FF9F,
    },
    Range {
        start: 0x00FFE3,
        end: 0x00FFE3,
    },
    Range {
        start: 0x00FFF9,
        end: 0x00FFFB,
    },
    Range {
        start: 0x0101FD,
        end: 0x0101FD,
    },
    Range {
        start: 0x0102E0,
        end: 0x0102E0,
    },
    Range {
        start: 0x010376,
        end: 0x01037A,
    },
    Range {
        start: 0x010780,
        end: 0x010785,
    },
    Range {
        start: 0x010787,
        end: 0x0107B0,
    },
    Range {
        start: 0x0107B2,
        end: 0x0107BA,
    },
    Range {
        start: 0x010A01,
        end: 0x010A03,
    },
    Range {
        start: 0x010A05,
        end: 0x010A06,
    },
    Range {
        start: 0x010A0C,
        end: 0x010A0F,
    },
    Range {
        start: 0x010A38,
        end: 0x010A3A,
    },
    Range {
        start: 0x010A3F,
        end: 0x010A3F,
    },
    Range {
        start: 0x010AE5,
        end: 0x010AE6,
    },
    Range {
        start: 0x010D24,
        end: 0x010D27,
    },
    Range {
        start: 0x010D4E,
        end: 0x010D4E,
    },
    Range {
        start: 0x010D69,
        end: 0x010D6D,
    },
    Range {
        start: 0x010D6F,
        end: 0x010D6F,
    },
    Range {
        start: 0x010EAB,
        end: 0x010EAC,
    },
    Range {
        start: 0x010EC5,
        end: 0x010EC5,
    },
    Range {
        start: 0x010EFA,
        end: 0x010EFF,
    },
    Range {
        start: 0x010F46,
        end: 0x010F50,
    },
    Range {
        start: 0x010F82,
        end: 0x010F85,
    },
    Range {
        start: 0x011001,
        end: 0x011001,
    },
    Range {
        start: 0x011038,
        end: 0x011046,
    },
    Range {
        start: 0x011070,
        end: 0x011070,
    },
    Range {
        start: 0x011073,
        end: 0x011074,
    },
    Range {
        start: 0x01107F,
        end: 0x011081,
    },
    Range {
        start: 0x0110B3,
        end: 0x0110B6,
    },
    Range {
        start: 0x0110B9,
        end: 0x0110BA,
    },
    Range {
        start: 0x0110BD,
        end: 0x0110BD,
    },
    Range {
        start: 0x0110C2,
        end: 0x0110C2,
    },
    Range {
        start: 0x0110CD,
        end: 0x0110CD,
    },
    Range {
        start: 0x011100,
        end: 0x011102,
    },
    Range {
        start: 0x011127,
        end: 0x01112B,
    },
    Range {
        start: 0x01112D,
        end: 0x011134,
    },
    Range {
        start: 0x011173,
        end: 0x011173,
    },
    Range {
        start: 0x011180,
        end: 0x011181,
    },
    Range {
        start: 0x0111B6,
        end: 0x0111BE,
    },
    Range {
        start: 0x0111C9,
        end: 0x0111CC,
    },
    Range {
        start: 0x0111CF,
        end: 0x0111CF,
    },
    Range {
        start: 0x01122F,
        end: 0x011231,
    },
    Range {
        start: 0x011234,
        end: 0x011234,
    },
    Range {
        start: 0x011236,
        end: 0x011237,
    },
    Range {
        start: 0x01123E,
        end: 0x01123E,
    },
    Range {
        start: 0x011241,
        end: 0x011241,
    },
    Range {
        start: 0x0112DF,
        end: 0x0112DF,
    },
    Range {
        start: 0x0112E3,
        end: 0x0112EA,
    },
    Range {
        start: 0x011300,
        end: 0x011301,
    },
    Range {
        start: 0x01133B,
        end: 0x01133C,
    },
    Range {
        start: 0x011340,
        end: 0x011340,
    },
    Range {
        start: 0x011366,
        end: 0x01136C,
    },
    Range {
        start: 0x011370,
        end: 0x011374,
    },
    Range {
        start: 0x0113BB,
        end: 0x0113C0,
    },
    Range {
        start: 0x0113CE,
        end: 0x0113CE,
    },
    Range {
        start: 0x0113D0,
        end: 0x0113D0,
    },
    Range {
        start: 0x0113D2,
        end: 0x0113D2,
    },
    Range {
        start: 0x0113E1,
        end: 0x0113E2,
    },
    Range {
        start: 0x011438,
        end: 0x01143F,
    },
    Range {
        start: 0x011442,
        end: 0x011444,
    },
    Range {
        start: 0x011446,
        end: 0x011446,
    },
    Range {
        start: 0x01145E,
        end: 0x01145E,
    },
    Range {
        start: 0x0114B3,
        end: 0x0114B8,
    },
    Range {
        start: 0x0114BA,
        end: 0x0114BA,
    },
    Range {
        start: 0x0114BF,
        end: 0x0114C0,
    },
    Range {
        start: 0x0114C2,
        end: 0x0114C3,
    },
    Range {
        start: 0x0115B2,
        end: 0x0115B5,
    },
    Range {
        start: 0x0115BC,
        end: 0x0115BD,
    },
    Range {
        start: 0x0115BF,
        end: 0x0115C0,
    },
    Range {
        start: 0x0115DC,
        end: 0x0115DD,
    },
    Range {
        start: 0x011633,
        end: 0x01163A,
    },
    Range {
        start: 0x01163D,
        end: 0x01163D,
    },
    Range {
        start: 0x01163F,
        end: 0x011640,
    },
    Range {
        start: 0x0116AB,
        end: 0x0116AB,
    },
    Range {
        start: 0x0116AD,
        end: 0x0116AD,
    },
    Range {
        start: 0x0116B0,
        end: 0x0116B5,
    },
    Range {
        start: 0x0116B7,
        end: 0x0116B7,
    },
    Range {
        start: 0x01171D,
        end: 0x01171D,
    },
    Range {
        start: 0x01171F,
        end: 0x01171F,
    },
    Range {
        start: 0x011722,
        end: 0x011725,
    },
    Range {
        start: 0x011727,
        end: 0x01172B,
    },
    Range {
        start: 0x01182F,
        end: 0x011837,
    },
    Range {
        start: 0x011839,
        end: 0x01183A,
    },
    Range {
        start: 0x01193B,
        end: 0x01193C,
    },
    Range {
        start: 0x01193E,
        end: 0x01193E,
    },
    Range {
        start: 0x011943,
        end: 0x011943,
    },
    Range {
        start: 0x0119D4,
        end: 0x0119D7,
    },
    Range {
        start: 0x0119DA,
        end: 0x0119DB,
    },
    Range {
        start: 0x0119E0,
        end: 0x0119E0,
    },
    Range {
        start: 0x011A01,
        end: 0x011A0A,
    },
    Range {
        start: 0x011A33,
        end: 0x011A38,
    },
    Range {
        start: 0x011A3B,
        end: 0x011A3E,
    },
    Range {
        start: 0x011A47,
        end: 0x011A47,
    },
    Range {
        start: 0x011A51,
        end: 0x011A56,
    },
    Range {
        start: 0x011A59,
        end: 0x011A5B,
    },
    Range {
        start: 0x011A8A,
        end: 0x011A96,
    },
    Range {
        start: 0x011A98,
        end: 0x011A99,
    },
    Range {
        start: 0x011B60,
        end: 0x011B60,
    },
    Range {
        start: 0x011B62,
        end: 0x011B64,
    },
    Range {
        start: 0x011B66,
        end: 0x011B66,
    },
    Range {
        start: 0x011C30,
        end: 0x011C36,
    },
    Range {
        start: 0x011C38,
        end: 0x011C3D,
    },
    Range {
        start: 0x011C3F,
        end: 0x011C3F,
    },
    Range {
        start: 0x011C92,
        end: 0x011CA7,
    },
    Range {
        start: 0x011CAA,
        end: 0x011CB0,
    },
    Range {
        start: 0x011CB2,
        end: 0x011CB3,
    },
    Range {
        start: 0x011CB5,
        end: 0x011CB6,
    },
    Range {
        start: 0x011D31,
        end: 0x011D36,
    },
    Range {
        start: 0x011D3A,
        end: 0x011D3A,
    },
    Range {
        start: 0x011D3C,
        end: 0x011D3D,
    },
    Range {
        start: 0x011D3F,
        end: 0x011D45,
    },
    Range {
        start: 0x011D47,
        end: 0x011D47,
    },
    Range {
        start: 0x011D90,
        end: 0x011D91,
    },
    Range {
        start: 0x011D95,
        end: 0x011D95,
    },
    Range {
        start: 0x011D97,
        end: 0x011D97,
    },
    Range {
        start: 0x011DD9,
        end: 0x011DD9,
    },
    Range {
        start: 0x011EF3,
        end: 0x011EF4,
    },
    Range {
        start: 0x011F00,
        end: 0x011F01,
    },
    Range {
        start: 0x011F36,
        end: 0x011F3A,
    },
    Range {
        start: 0x011F40,
        end: 0x011F40,
    },
    Range {
        start: 0x011F42,
        end: 0x011F42,
    },
    Range {
        start: 0x011F5A,
        end: 0x011F5A,
    },
    Range {
        start: 0x013430,
        end: 0x013440,
    },
    Range {
        start: 0x013447,
        end: 0x013455,
    },
    Range {
        start: 0x01611E,
        end: 0x016129,
    },
    Range {
        start: 0x01612D,
        end: 0x01612F,
    },
    Range {
        start: 0x016AF0,
        end: 0x016AF4,
    },
    Range {
        start: 0x016B30,
        end: 0x016B36,
    },
    Range {
        start: 0x016B40,
        end: 0x016B43,
    },
    Range {
        start: 0x016D40,
        end: 0x016D42,
    },
    Range {
        start: 0x016D6B,
        end: 0x016D6C,
    },
    Range {
        start: 0x016F4F,
        end: 0x016F4F,
    },
    Range {
        start: 0x016F8F,
        end: 0x016F9F,
    },
    Range {
        start: 0x016FE0,
        end: 0x016FE1,
    },
    Range {
        start: 0x016FE3,
        end: 0x016FE4,
    },
    Range {
        start: 0x016FF2,
        end: 0x016FF3,
    },
    Range {
        start: 0x01AFF0,
        end: 0x01AFF3,
    },
    Range {
        start: 0x01AFF5,
        end: 0x01AFFB,
    },
    Range {
        start: 0x01AFFD,
        end: 0x01AFFE,
    },
    Range {
        start: 0x01BC9D,
        end: 0x01BC9E,
    },
    Range {
        start: 0x01BCA0,
        end: 0x01BCA3,
    },
    Range {
        start: 0x01CF00,
        end: 0x01CF2D,
    },
    Range {
        start: 0x01CF30,
        end: 0x01CF46,
    },
    Range {
        start: 0x01D167,
        end: 0x01D169,
    },
    Range {
        start: 0x01D173,
        end: 0x01D182,
    },
    Range {
        start: 0x01D185,
        end: 0x01D18B,
    },
    Range {
        start: 0x01D1AA,
        end: 0x01D1AD,
    },
    Range {
        start: 0x01D242,
        end: 0x01D244,
    },
    Range {
        start: 0x01DA00,
        end: 0x01DA36,
    },
    Range {
        start: 0x01DA3B,
        end: 0x01DA6C,
    },
    Range {
        start: 0x01DA75,
        end: 0x01DA75,
    },
    Range {
        start: 0x01DA84,
        end: 0x01DA84,
    },
    Range {
        start: 0x01DA9B,
        end: 0x01DA9F,
    },
    Range {
        start: 0x01DAA1,
        end: 0x01DAAF,
    },
    Range {
        start: 0x01E000,
        end: 0x01E006,
    },
    Range {
        start: 0x01E008,
        end: 0x01E018,
    },
    Range {
        start: 0x01E01B,
        end: 0x01E021,
    },
    Range {
        start: 0x01E023,
        end: 0x01E024,
    },
    Range {
        start: 0x01E026,
        end: 0x01E02A,
    },
    Range {
        start: 0x01E030,
        end: 0x01E06D,
    },
    Range {
        start: 0x01E08F,
        end: 0x01E08F,
    },
    Range {
        start: 0x01E130,
        end: 0x01E13D,
    },
    Range {
        start: 0x01E2AE,
        end: 0x01E2AE,
    },
    Range {
        start: 0x01E2EC,
        end: 0x01E2EF,
    },
    Range {
        start: 0x01E4EB,
        end: 0x01E4EF,
    },
    Range {
        start: 0x01E5EE,
        end: 0x01E5EF,
    },
    Range {
        start: 0x01E6E3,
        end: 0x01E6E3,
    },
    Range {
        start: 0x01E6E6,
        end: 0x01E6E6,
    },
    Range {
        start: 0x01E6EE,
        end: 0x01E6EF,
    },
    Range {
        start: 0x01E6F5,
        end: 0x01E6F5,
    },
    Range {
        start: 0x01E6FF,
        end: 0x01E6FF,
    },
    Range {
        start: 0x01E8D0,
        end: 0x01E8D6,
    },
    Range {
        start: 0x01E944,
        end: 0x01E94B,
    },
    Range {
        start: 0x01F3FB,
        end: 0x01F3FF,
    },
    Range {
        start: 0x0E0001,
        end: 0x0E0001,
    },
    Range {
        start: 0x0E0020,
        end: 0x0E007F,
    },
    Range {
        start: 0x0E0100,
        end: 0x0E01EF,
    },
];

#[derive(Clone, Copy)]
struct Mapping {
    source: u32,
    len: u8,
    target: [u32; 3],
}

const CASE_FOLD: &[Mapping] = &[
    Mapping {
        source: 0x000041,
        len: 1,
        target: [0x000061, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000042,
        len: 1,
        target: [0x000062, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000043,
        len: 1,
        target: [0x000063, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000044,
        len: 1,
        target: [0x000064, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000045,
        len: 1,
        target: [0x000065, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000046,
        len: 1,
        target: [0x000066, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000047,
        len: 1,
        target: [0x000067, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000048,
        len: 1,
        target: [0x000068, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000049,
        len: 1,
        target: [0x000069, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004A,
        len: 1,
        target: [0x00006A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004B,
        len: 1,
        target: [0x00006B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004C,
        len: 1,
        target: [0x00006C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004D,
        len: 1,
        target: [0x00006D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004E,
        len: 1,
        target: [0x00006E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004F,
        len: 1,
        target: [0x00006F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000050,
        len: 1,
        target: [0x000070, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000051,
        len: 1,
        target: [0x000071, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000052,
        len: 1,
        target: [0x000072, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000053,
        len: 1,
        target: [0x000073, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000054,
        len: 1,
        target: [0x000074, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000055,
        len: 1,
        target: [0x000075, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000056,
        len: 1,
        target: [0x000076, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000057,
        len: 1,
        target: [0x000077, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000058,
        len: 1,
        target: [0x000078, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000059,
        len: 1,
        target: [0x000079, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00005A,
        len: 1,
        target: [0x00007A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000B5,
        len: 1,
        target: [0x0003BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C0,
        len: 1,
        target: [0x0000E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C1,
        len: 1,
        target: [0x0000E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C2,
        len: 1,
        target: [0x0000E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C3,
        len: 1,
        target: [0x0000E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C4,
        len: 1,
        target: [0x0000E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C5,
        len: 1,
        target: [0x0000E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C6,
        len: 1,
        target: [0x0000E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C7,
        len: 1,
        target: [0x0000E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C8,
        len: 1,
        target: [0x0000E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C9,
        len: 1,
        target: [0x0000E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CA,
        len: 1,
        target: [0x0000EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CB,
        len: 1,
        target: [0x0000EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CC,
        len: 1,
        target: [0x0000EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CD,
        len: 1,
        target: [0x0000ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CE,
        len: 1,
        target: [0x0000EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CF,
        len: 1,
        target: [0x0000EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D0,
        len: 1,
        target: [0x0000F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D1,
        len: 1,
        target: [0x0000F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D2,
        len: 1,
        target: [0x0000F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D3,
        len: 1,
        target: [0x0000F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D4,
        len: 1,
        target: [0x0000F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D5,
        len: 1,
        target: [0x0000F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D6,
        len: 1,
        target: [0x0000F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D8,
        len: 1,
        target: [0x0000F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D9,
        len: 1,
        target: [0x0000F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DA,
        len: 1,
        target: [0x0000FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DB,
        len: 1,
        target: [0x0000FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DC,
        len: 1,
        target: [0x0000FC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DD,
        len: 1,
        target: [0x0000FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DE,
        len: 1,
        target: [0x0000FE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DF,
        len: 2,
        target: [0x000073, 0x000073, 0x000000],
    },
    Mapping {
        source: 0x000100,
        len: 1,
        target: [0x000101, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000102,
        len: 1,
        target: [0x000103, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000104,
        len: 1,
        target: [0x000105, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000106,
        len: 1,
        target: [0x000107, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000108,
        len: 1,
        target: [0x000109, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010A,
        len: 1,
        target: [0x00010B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010C,
        len: 1,
        target: [0x00010D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010E,
        len: 1,
        target: [0x00010F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000110,
        len: 1,
        target: [0x000111, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000112,
        len: 1,
        target: [0x000113, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000114,
        len: 1,
        target: [0x000115, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000116,
        len: 1,
        target: [0x000117, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000118,
        len: 1,
        target: [0x000119, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011A,
        len: 1,
        target: [0x00011B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011C,
        len: 1,
        target: [0x00011D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011E,
        len: 1,
        target: [0x00011F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000120,
        len: 1,
        target: [0x000121, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000122,
        len: 1,
        target: [0x000123, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000124,
        len: 1,
        target: [0x000125, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000126,
        len: 1,
        target: [0x000127, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000128,
        len: 1,
        target: [0x000129, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012A,
        len: 1,
        target: [0x00012B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012C,
        len: 1,
        target: [0x00012D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012E,
        len: 1,
        target: [0x00012F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000130,
        len: 2,
        target: [0x000069, 0x000307, 0x000000],
    },
    Mapping {
        source: 0x000132,
        len: 1,
        target: [0x000133, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000134,
        len: 1,
        target: [0x000135, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000136,
        len: 1,
        target: [0x000137, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000139,
        len: 1,
        target: [0x00013A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013B,
        len: 1,
        target: [0x00013C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013D,
        len: 1,
        target: [0x00013E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013F,
        len: 1,
        target: [0x000140, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000141,
        len: 1,
        target: [0x000142, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000143,
        len: 1,
        target: [0x000144, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000145,
        len: 1,
        target: [0x000146, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000147,
        len: 1,
        target: [0x000148, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000149,
        len: 2,
        target: [0x0002BC, 0x00006E, 0x000000],
    },
    Mapping {
        source: 0x00014A,
        len: 1,
        target: [0x00014B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00014C,
        len: 1,
        target: [0x00014D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00014E,
        len: 1,
        target: [0x00014F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000150,
        len: 1,
        target: [0x000151, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000152,
        len: 1,
        target: [0x000153, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000154,
        len: 1,
        target: [0x000155, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000156,
        len: 1,
        target: [0x000157, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000158,
        len: 1,
        target: [0x000159, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015A,
        len: 1,
        target: [0x00015B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015C,
        len: 1,
        target: [0x00015D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015E,
        len: 1,
        target: [0x00015F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000160,
        len: 1,
        target: [0x000161, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000162,
        len: 1,
        target: [0x000163, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000164,
        len: 1,
        target: [0x000165, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000166,
        len: 1,
        target: [0x000167, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000168,
        len: 1,
        target: [0x000169, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016A,
        len: 1,
        target: [0x00016B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016C,
        len: 1,
        target: [0x00016D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016E,
        len: 1,
        target: [0x00016F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000170,
        len: 1,
        target: [0x000171, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000172,
        len: 1,
        target: [0x000173, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000174,
        len: 1,
        target: [0x000175, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000176,
        len: 1,
        target: [0x000177, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000178,
        len: 1,
        target: [0x0000FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000179,
        len: 1,
        target: [0x00017A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017B,
        len: 1,
        target: [0x00017C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017D,
        len: 1,
        target: [0x00017E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017F,
        len: 1,
        target: [0x000073, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000181,
        len: 1,
        target: [0x000253, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000182,
        len: 1,
        target: [0x000183, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000184,
        len: 1,
        target: [0x000185, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000186,
        len: 1,
        target: [0x000254, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000187,
        len: 1,
        target: [0x000188, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000189,
        len: 1,
        target: [0x000256, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018A,
        len: 1,
        target: [0x000257, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018B,
        len: 1,
        target: [0x00018C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018E,
        len: 1,
        target: [0x0001DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018F,
        len: 1,
        target: [0x000259, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000190,
        len: 1,
        target: [0x00025B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000191,
        len: 1,
        target: [0x000192, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000193,
        len: 1,
        target: [0x000260, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000194,
        len: 1,
        target: [0x000263, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000196,
        len: 1,
        target: [0x000269, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000197,
        len: 1,
        target: [0x000268, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000198,
        len: 1,
        target: [0x000199, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019C,
        len: 1,
        target: [0x00026F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019D,
        len: 1,
        target: [0x000272, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019F,
        len: 1,
        target: [0x000275, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A0,
        len: 1,
        target: [0x0001A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A2,
        len: 1,
        target: [0x0001A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A4,
        len: 1,
        target: [0x0001A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A6,
        len: 1,
        target: [0x000280, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A7,
        len: 1,
        target: [0x0001A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A9,
        len: 1,
        target: [0x000283, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001AC,
        len: 1,
        target: [0x0001AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001AE,
        len: 1,
        target: [0x000288, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001AF,
        len: 1,
        target: [0x0001B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B1,
        len: 1,
        target: [0x00028A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B2,
        len: 1,
        target: [0x00028B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B3,
        len: 1,
        target: [0x0001B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B5,
        len: 1,
        target: [0x0001B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B7,
        len: 1,
        target: [0x000292, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B8,
        len: 1,
        target: [0x0001B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001BC,
        len: 1,
        target: [0x0001BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C4,
        len: 1,
        target: [0x0001C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C5,
        len: 1,
        target: [0x0001C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C7,
        len: 1,
        target: [0x0001C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C8,
        len: 1,
        target: [0x0001C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CA,
        len: 1,
        target: [0x0001CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CB,
        len: 1,
        target: [0x0001CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CD,
        len: 1,
        target: [0x0001CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CF,
        len: 1,
        target: [0x0001D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D1,
        len: 1,
        target: [0x0001D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D3,
        len: 1,
        target: [0x0001D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D5,
        len: 1,
        target: [0x0001D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D7,
        len: 1,
        target: [0x0001D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D9,
        len: 1,
        target: [0x0001DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DB,
        len: 1,
        target: [0x0001DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DE,
        len: 1,
        target: [0x0001DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E0,
        len: 1,
        target: [0x0001E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E2,
        len: 1,
        target: [0x0001E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E4,
        len: 1,
        target: [0x0001E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E6,
        len: 1,
        target: [0x0001E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E8,
        len: 1,
        target: [0x0001E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EA,
        len: 1,
        target: [0x0001EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EC,
        len: 1,
        target: [0x0001ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EE,
        len: 1,
        target: [0x0001EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F0,
        len: 2,
        target: [0x00006A, 0x00030C, 0x000000],
    },
    Mapping {
        source: 0x0001F1,
        len: 1,
        target: [0x0001F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F2,
        len: 1,
        target: [0x0001F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F4,
        len: 1,
        target: [0x0001F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F6,
        len: 1,
        target: [0x000195, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F7,
        len: 1,
        target: [0x0001BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F8,
        len: 1,
        target: [0x0001F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FA,
        len: 1,
        target: [0x0001FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FC,
        len: 1,
        target: [0x0001FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FE,
        len: 1,
        target: [0x0001FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000200,
        len: 1,
        target: [0x000201, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000202,
        len: 1,
        target: [0x000203, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000204,
        len: 1,
        target: [0x000205, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000206,
        len: 1,
        target: [0x000207, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000208,
        len: 1,
        target: [0x000209, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020A,
        len: 1,
        target: [0x00020B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020C,
        len: 1,
        target: [0x00020D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020E,
        len: 1,
        target: [0x00020F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000210,
        len: 1,
        target: [0x000211, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000212,
        len: 1,
        target: [0x000213, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000214,
        len: 1,
        target: [0x000215, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000216,
        len: 1,
        target: [0x000217, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000218,
        len: 1,
        target: [0x000219, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021A,
        len: 1,
        target: [0x00021B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021C,
        len: 1,
        target: [0x00021D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021E,
        len: 1,
        target: [0x00021F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000220,
        len: 1,
        target: [0x00019E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000222,
        len: 1,
        target: [0x000223, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000224,
        len: 1,
        target: [0x000225, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000226,
        len: 1,
        target: [0x000227, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000228,
        len: 1,
        target: [0x000229, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022A,
        len: 1,
        target: [0x00022B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022C,
        len: 1,
        target: [0x00022D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022E,
        len: 1,
        target: [0x00022F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000230,
        len: 1,
        target: [0x000231, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000232,
        len: 1,
        target: [0x000233, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023A,
        len: 1,
        target: [0x002C65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023B,
        len: 1,
        target: [0x00023C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023D,
        len: 1,
        target: [0x00019A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023E,
        len: 1,
        target: [0x002C66, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000241,
        len: 1,
        target: [0x000242, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000243,
        len: 1,
        target: [0x000180, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000244,
        len: 1,
        target: [0x000289, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000245,
        len: 1,
        target: [0x00028C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000246,
        len: 1,
        target: [0x000247, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000248,
        len: 1,
        target: [0x000249, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024A,
        len: 1,
        target: [0x00024B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024C,
        len: 1,
        target: [0x00024D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024E,
        len: 1,
        target: [0x00024F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000345,
        len: 1,
        target: [0x0003B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000370,
        len: 1,
        target: [0x000371, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000372,
        len: 1,
        target: [0x000373, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000376,
        len: 1,
        target: [0x000377, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00037F,
        len: 1,
        target: [0x0003F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000386,
        len: 1,
        target: [0x0003AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000388,
        len: 1,
        target: [0x0003AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000389,
        len: 1,
        target: [0x0003AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038A,
        len: 1,
        target: [0x0003AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038C,
        len: 1,
        target: [0x0003CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038E,
        len: 1,
        target: [0x0003CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038F,
        len: 1,
        target: [0x0003CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000390,
        len: 3,
        target: [0x0003B9, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x000391,
        len: 1,
        target: [0x0003B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000392,
        len: 1,
        target: [0x0003B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000393,
        len: 1,
        target: [0x0003B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000394,
        len: 1,
        target: [0x0003B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000395,
        len: 1,
        target: [0x0003B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000396,
        len: 1,
        target: [0x0003B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000397,
        len: 1,
        target: [0x0003B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000398,
        len: 1,
        target: [0x0003B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000399,
        len: 1,
        target: [0x0003B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039A,
        len: 1,
        target: [0x0003BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039B,
        len: 1,
        target: [0x0003BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039C,
        len: 1,
        target: [0x0003BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039D,
        len: 1,
        target: [0x0003BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039E,
        len: 1,
        target: [0x0003BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039F,
        len: 1,
        target: [0x0003BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A0,
        len: 1,
        target: [0x0003C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A1,
        len: 1,
        target: [0x0003C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A3,
        len: 1,
        target: [0x0003C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A4,
        len: 1,
        target: [0x0003C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A5,
        len: 1,
        target: [0x0003C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A6,
        len: 1,
        target: [0x0003C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A7,
        len: 1,
        target: [0x0003C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A8,
        len: 1,
        target: [0x0003C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A9,
        len: 1,
        target: [0x0003C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003AA,
        len: 1,
        target: [0x0003CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003AB,
        len: 1,
        target: [0x0003CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B0,
        len: 3,
        target: [0x0003C5, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x0003C2,
        len: 1,
        target: [0x0003C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003CF,
        len: 1,
        target: [0x0003D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D0,
        len: 1,
        target: [0x0003B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D1,
        len: 1,
        target: [0x0003B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D5,
        len: 1,
        target: [0x0003C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D6,
        len: 1,
        target: [0x0003C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D8,
        len: 1,
        target: [0x0003D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DA,
        len: 1,
        target: [0x0003DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DC,
        len: 1,
        target: [0x0003DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DE,
        len: 1,
        target: [0x0003DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E0,
        len: 1,
        target: [0x0003E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E2,
        len: 1,
        target: [0x0003E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E4,
        len: 1,
        target: [0x0003E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E6,
        len: 1,
        target: [0x0003E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E8,
        len: 1,
        target: [0x0003E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EA,
        len: 1,
        target: [0x0003EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EC,
        len: 1,
        target: [0x0003ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EE,
        len: 1,
        target: [0x0003EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F0,
        len: 1,
        target: [0x0003BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F1,
        len: 1,
        target: [0x0003C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F4,
        len: 1,
        target: [0x0003B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F5,
        len: 1,
        target: [0x0003B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F7,
        len: 1,
        target: [0x0003F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F9,
        len: 1,
        target: [0x0003F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FA,
        len: 1,
        target: [0x0003FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FD,
        len: 1,
        target: [0x00037B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FE,
        len: 1,
        target: [0x00037C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FF,
        len: 1,
        target: [0x00037D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000400,
        len: 1,
        target: [0x000450, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000401,
        len: 1,
        target: [0x000451, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000402,
        len: 1,
        target: [0x000452, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000403,
        len: 1,
        target: [0x000453, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000404,
        len: 1,
        target: [0x000454, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000405,
        len: 1,
        target: [0x000455, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000406,
        len: 1,
        target: [0x000456, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000407,
        len: 1,
        target: [0x000457, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000408,
        len: 1,
        target: [0x000458, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000409,
        len: 1,
        target: [0x000459, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040A,
        len: 1,
        target: [0x00045A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040B,
        len: 1,
        target: [0x00045B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040C,
        len: 1,
        target: [0x00045C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040D,
        len: 1,
        target: [0x00045D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040E,
        len: 1,
        target: [0x00045E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040F,
        len: 1,
        target: [0x00045F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000410,
        len: 1,
        target: [0x000430, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000411,
        len: 1,
        target: [0x000431, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000412,
        len: 1,
        target: [0x000432, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000413,
        len: 1,
        target: [0x000433, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000414,
        len: 1,
        target: [0x000434, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000415,
        len: 1,
        target: [0x000435, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000416,
        len: 1,
        target: [0x000436, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000417,
        len: 1,
        target: [0x000437, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000418,
        len: 1,
        target: [0x000438, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000419,
        len: 1,
        target: [0x000439, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041A,
        len: 1,
        target: [0x00043A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041B,
        len: 1,
        target: [0x00043B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041C,
        len: 1,
        target: [0x00043C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041D,
        len: 1,
        target: [0x00043D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041E,
        len: 1,
        target: [0x00043E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041F,
        len: 1,
        target: [0x00043F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000420,
        len: 1,
        target: [0x000440, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000421,
        len: 1,
        target: [0x000441, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000422,
        len: 1,
        target: [0x000442, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000423,
        len: 1,
        target: [0x000443, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000424,
        len: 1,
        target: [0x000444, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000425,
        len: 1,
        target: [0x000445, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000426,
        len: 1,
        target: [0x000446, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000427,
        len: 1,
        target: [0x000447, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000428,
        len: 1,
        target: [0x000448, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000429,
        len: 1,
        target: [0x000449, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042A,
        len: 1,
        target: [0x00044A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042B,
        len: 1,
        target: [0x00044B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042C,
        len: 1,
        target: [0x00044C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042D,
        len: 1,
        target: [0x00044D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042E,
        len: 1,
        target: [0x00044E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042F,
        len: 1,
        target: [0x00044F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000460,
        len: 1,
        target: [0x000461, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000462,
        len: 1,
        target: [0x000463, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000464,
        len: 1,
        target: [0x000465, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000466,
        len: 1,
        target: [0x000467, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000468,
        len: 1,
        target: [0x000469, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046A,
        len: 1,
        target: [0x00046B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046C,
        len: 1,
        target: [0x00046D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046E,
        len: 1,
        target: [0x00046F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000470,
        len: 1,
        target: [0x000471, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000472,
        len: 1,
        target: [0x000473, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000474,
        len: 1,
        target: [0x000475, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000476,
        len: 1,
        target: [0x000477, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000478,
        len: 1,
        target: [0x000479, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047A,
        len: 1,
        target: [0x00047B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047C,
        len: 1,
        target: [0x00047D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047E,
        len: 1,
        target: [0x00047F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000480,
        len: 1,
        target: [0x000481, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048A,
        len: 1,
        target: [0x00048B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048C,
        len: 1,
        target: [0x00048D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048E,
        len: 1,
        target: [0x00048F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000490,
        len: 1,
        target: [0x000491, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000492,
        len: 1,
        target: [0x000493, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000494,
        len: 1,
        target: [0x000495, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000496,
        len: 1,
        target: [0x000497, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000498,
        len: 1,
        target: [0x000499, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049A,
        len: 1,
        target: [0x00049B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049C,
        len: 1,
        target: [0x00049D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049E,
        len: 1,
        target: [0x00049F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A0,
        len: 1,
        target: [0x0004A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A2,
        len: 1,
        target: [0x0004A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A4,
        len: 1,
        target: [0x0004A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A6,
        len: 1,
        target: [0x0004A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A8,
        len: 1,
        target: [0x0004A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AA,
        len: 1,
        target: [0x0004AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AC,
        len: 1,
        target: [0x0004AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AE,
        len: 1,
        target: [0x0004AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B0,
        len: 1,
        target: [0x0004B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B2,
        len: 1,
        target: [0x0004B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B4,
        len: 1,
        target: [0x0004B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B6,
        len: 1,
        target: [0x0004B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B8,
        len: 1,
        target: [0x0004B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BA,
        len: 1,
        target: [0x0004BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BC,
        len: 1,
        target: [0x0004BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BE,
        len: 1,
        target: [0x0004BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C0,
        len: 1,
        target: [0x0004CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C1,
        len: 1,
        target: [0x0004C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C3,
        len: 1,
        target: [0x0004C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C5,
        len: 1,
        target: [0x0004C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C7,
        len: 1,
        target: [0x0004C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C9,
        len: 1,
        target: [0x0004CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CB,
        len: 1,
        target: [0x0004CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CD,
        len: 1,
        target: [0x0004CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D0,
        len: 1,
        target: [0x0004D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D2,
        len: 1,
        target: [0x0004D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D4,
        len: 1,
        target: [0x0004D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D6,
        len: 1,
        target: [0x0004D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D8,
        len: 1,
        target: [0x0004D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DA,
        len: 1,
        target: [0x0004DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DC,
        len: 1,
        target: [0x0004DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DE,
        len: 1,
        target: [0x0004DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E0,
        len: 1,
        target: [0x0004E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E2,
        len: 1,
        target: [0x0004E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E4,
        len: 1,
        target: [0x0004E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E6,
        len: 1,
        target: [0x0004E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E8,
        len: 1,
        target: [0x0004E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EA,
        len: 1,
        target: [0x0004EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EC,
        len: 1,
        target: [0x0004ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EE,
        len: 1,
        target: [0x0004EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F0,
        len: 1,
        target: [0x0004F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F2,
        len: 1,
        target: [0x0004F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F4,
        len: 1,
        target: [0x0004F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F6,
        len: 1,
        target: [0x0004F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F8,
        len: 1,
        target: [0x0004F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FA,
        len: 1,
        target: [0x0004FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FC,
        len: 1,
        target: [0x0004FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FE,
        len: 1,
        target: [0x0004FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000500,
        len: 1,
        target: [0x000501, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000502,
        len: 1,
        target: [0x000503, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000504,
        len: 1,
        target: [0x000505, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000506,
        len: 1,
        target: [0x000507, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000508,
        len: 1,
        target: [0x000509, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050A,
        len: 1,
        target: [0x00050B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050C,
        len: 1,
        target: [0x00050D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050E,
        len: 1,
        target: [0x00050F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000510,
        len: 1,
        target: [0x000511, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000512,
        len: 1,
        target: [0x000513, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000514,
        len: 1,
        target: [0x000515, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000516,
        len: 1,
        target: [0x000517, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000518,
        len: 1,
        target: [0x000519, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051A,
        len: 1,
        target: [0x00051B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051C,
        len: 1,
        target: [0x00051D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051E,
        len: 1,
        target: [0x00051F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000520,
        len: 1,
        target: [0x000521, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000522,
        len: 1,
        target: [0x000523, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000524,
        len: 1,
        target: [0x000525, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000526,
        len: 1,
        target: [0x000527, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000528,
        len: 1,
        target: [0x000529, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052A,
        len: 1,
        target: [0x00052B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052C,
        len: 1,
        target: [0x00052D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052E,
        len: 1,
        target: [0x00052F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000531,
        len: 1,
        target: [0x000561, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000532,
        len: 1,
        target: [0x000562, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000533,
        len: 1,
        target: [0x000563, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000534,
        len: 1,
        target: [0x000564, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000535,
        len: 1,
        target: [0x000565, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000536,
        len: 1,
        target: [0x000566, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000537,
        len: 1,
        target: [0x000567, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000538,
        len: 1,
        target: [0x000568, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000539,
        len: 1,
        target: [0x000569, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053A,
        len: 1,
        target: [0x00056A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053B,
        len: 1,
        target: [0x00056B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053C,
        len: 1,
        target: [0x00056C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053D,
        len: 1,
        target: [0x00056D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053E,
        len: 1,
        target: [0x00056E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053F,
        len: 1,
        target: [0x00056F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000540,
        len: 1,
        target: [0x000570, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000541,
        len: 1,
        target: [0x000571, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000542,
        len: 1,
        target: [0x000572, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000543,
        len: 1,
        target: [0x000573, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000544,
        len: 1,
        target: [0x000574, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000545,
        len: 1,
        target: [0x000575, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000546,
        len: 1,
        target: [0x000576, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000547,
        len: 1,
        target: [0x000577, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000548,
        len: 1,
        target: [0x000578, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000549,
        len: 1,
        target: [0x000579, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054A,
        len: 1,
        target: [0x00057A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054B,
        len: 1,
        target: [0x00057B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054C,
        len: 1,
        target: [0x00057C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054D,
        len: 1,
        target: [0x00057D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054E,
        len: 1,
        target: [0x00057E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054F,
        len: 1,
        target: [0x00057F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000550,
        len: 1,
        target: [0x000580, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000551,
        len: 1,
        target: [0x000581, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000552,
        len: 1,
        target: [0x000582, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000553,
        len: 1,
        target: [0x000583, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000554,
        len: 1,
        target: [0x000584, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000555,
        len: 1,
        target: [0x000585, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000556,
        len: 1,
        target: [0x000586, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000587,
        len: 2,
        target: [0x000565, 0x000582, 0x000000],
    },
    Mapping {
        source: 0x0010A0,
        len: 1,
        target: [0x002D00, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A1,
        len: 1,
        target: [0x002D01, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A2,
        len: 1,
        target: [0x002D02, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A3,
        len: 1,
        target: [0x002D03, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A4,
        len: 1,
        target: [0x002D04, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A5,
        len: 1,
        target: [0x002D05, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A6,
        len: 1,
        target: [0x002D06, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A7,
        len: 1,
        target: [0x002D07, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A8,
        len: 1,
        target: [0x002D08, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A9,
        len: 1,
        target: [0x002D09, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AA,
        len: 1,
        target: [0x002D0A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AB,
        len: 1,
        target: [0x002D0B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AC,
        len: 1,
        target: [0x002D0C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AD,
        len: 1,
        target: [0x002D0D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AE,
        len: 1,
        target: [0x002D0E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AF,
        len: 1,
        target: [0x002D0F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B0,
        len: 1,
        target: [0x002D10, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B1,
        len: 1,
        target: [0x002D11, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B2,
        len: 1,
        target: [0x002D12, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B3,
        len: 1,
        target: [0x002D13, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B4,
        len: 1,
        target: [0x002D14, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B5,
        len: 1,
        target: [0x002D15, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B6,
        len: 1,
        target: [0x002D16, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B7,
        len: 1,
        target: [0x002D17, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B8,
        len: 1,
        target: [0x002D18, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B9,
        len: 1,
        target: [0x002D19, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BA,
        len: 1,
        target: [0x002D1A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BB,
        len: 1,
        target: [0x002D1B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BC,
        len: 1,
        target: [0x002D1C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BD,
        len: 1,
        target: [0x002D1D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BE,
        len: 1,
        target: [0x002D1E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BF,
        len: 1,
        target: [0x002D1F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C0,
        len: 1,
        target: [0x002D20, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C1,
        len: 1,
        target: [0x002D21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C2,
        len: 1,
        target: [0x002D22, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C3,
        len: 1,
        target: [0x002D23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C4,
        len: 1,
        target: [0x002D24, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C5,
        len: 1,
        target: [0x002D25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C7,
        len: 1,
        target: [0x002D27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010CD,
        len: 1,
        target: [0x002D2D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F8,
        len: 1,
        target: [0x0013F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F9,
        len: 1,
        target: [0x0013F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FA,
        len: 1,
        target: [0x0013F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FB,
        len: 1,
        target: [0x0013F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FC,
        len: 1,
        target: [0x0013F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FD,
        len: 1,
        target: [0x0013F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C80,
        len: 1,
        target: [0x000432, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C81,
        len: 1,
        target: [0x000434, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C82,
        len: 1,
        target: [0x00043E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C83,
        len: 1,
        target: [0x000441, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C84,
        len: 1,
        target: [0x000442, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C85,
        len: 1,
        target: [0x000442, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C86,
        len: 1,
        target: [0x00044A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C87,
        len: 1,
        target: [0x000463, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C88,
        len: 1,
        target: [0x00A64B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C89,
        len: 1,
        target: [0x001C8A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C90,
        len: 1,
        target: [0x0010D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C91,
        len: 1,
        target: [0x0010D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C92,
        len: 1,
        target: [0x0010D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C93,
        len: 1,
        target: [0x0010D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C94,
        len: 1,
        target: [0x0010D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C95,
        len: 1,
        target: [0x0010D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C96,
        len: 1,
        target: [0x0010D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C97,
        len: 1,
        target: [0x0010D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C98,
        len: 1,
        target: [0x0010D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C99,
        len: 1,
        target: [0x0010D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9A,
        len: 1,
        target: [0x0010DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9B,
        len: 1,
        target: [0x0010DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9C,
        len: 1,
        target: [0x0010DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9D,
        len: 1,
        target: [0x0010DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9E,
        len: 1,
        target: [0x0010DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9F,
        len: 1,
        target: [0x0010DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA0,
        len: 1,
        target: [0x0010E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA1,
        len: 1,
        target: [0x0010E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA2,
        len: 1,
        target: [0x0010E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA3,
        len: 1,
        target: [0x0010E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA4,
        len: 1,
        target: [0x0010E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA5,
        len: 1,
        target: [0x0010E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA6,
        len: 1,
        target: [0x0010E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA7,
        len: 1,
        target: [0x0010E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA8,
        len: 1,
        target: [0x0010E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA9,
        len: 1,
        target: [0x0010E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAA,
        len: 1,
        target: [0x0010EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAB,
        len: 1,
        target: [0x0010EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAC,
        len: 1,
        target: [0x0010EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAD,
        len: 1,
        target: [0x0010ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAE,
        len: 1,
        target: [0x0010EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAF,
        len: 1,
        target: [0x0010EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB0,
        len: 1,
        target: [0x0010F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB1,
        len: 1,
        target: [0x0010F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB2,
        len: 1,
        target: [0x0010F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB3,
        len: 1,
        target: [0x0010F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB4,
        len: 1,
        target: [0x0010F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB5,
        len: 1,
        target: [0x0010F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB6,
        len: 1,
        target: [0x0010F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB7,
        len: 1,
        target: [0x0010F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB8,
        len: 1,
        target: [0x0010F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB9,
        len: 1,
        target: [0x0010F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBA,
        len: 1,
        target: [0x0010FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBD,
        len: 1,
        target: [0x0010FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBE,
        len: 1,
        target: [0x0010FE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBF,
        len: 1,
        target: [0x0010FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E00,
        len: 1,
        target: [0x001E01, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E02,
        len: 1,
        target: [0x001E03, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E04,
        len: 1,
        target: [0x001E05, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E06,
        len: 1,
        target: [0x001E07, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E08,
        len: 1,
        target: [0x001E09, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0A,
        len: 1,
        target: [0x001E0B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0C,
        len: 1,
        target: [0x001E0D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0E,
        len: 1,
        target: [0x001E0F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E10,
        len: 1,
        target: [0x001E11, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E12,
        len: 1,
        target: [0x001E13, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E14,
        len: 1,
        target: [0x001E15, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E16,
        len: 1,
        target: [0x001E17, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E18,
        len: 1,
        target: [0x001E19, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1A,
        len: 1,
        target: [0x001E1B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1C,
        len: 1,
        target: [0x001E1D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1E,
        len: 1,
        target: [0x001E1F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E20,
        len: 1,
        target: [0x001E21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E22,
        len: 1,
        target: [0x001E23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E24,
        len: 1,
        target: [0x001E25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E26,
        len: 1,
        target: [0x001E27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E28,
        len: 1,
        target: [0x001E29, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2A,
        len: 1,
        target: [0x001E2B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2C,
        len: 1,
        target: [0x001E2D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2E,
        len: 1,
        target: [0x001E2F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E30,
        len: 1,
        target: [0x001E31, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E32,
        len: 1,
        target: [0x001E33, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E34,
        len: 1,
        target: [0x001E35, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E36,
        len: 1,
        target: [0x001E37, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E38,
        len: 1,
        target: [0x001E39, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3A,
        len: 1,
        target: [0x001E3B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3C,
        len: 1,
        target: [0x001E3D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3E,
        len: 1,
        target: [0x001E3F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E40,
        len: 1,
        target: [0x001E41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E42,
        len: 1,
        target: [0x001E43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E44,
        len: 1,
        target: [0x001E45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E46,
        len: 1,
        target: [0x001E47, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E48,
        len: 1,
        target: [0x001E49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4A,
        len: 1,
        target: [0x001E4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4C,
        len: 1,
        target: [0x001E4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4E,
        len: 1,
        target: [0x001E4F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E50,
        len: 1,
        target: [0x001E51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E52,
        len: 1,
        target: [0x001E53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E54,
        len: 1,
        target: [0x001E55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E56,
        len: 1,
        target: [0x001E57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E58,
        len: 1,
        target: [0x001E59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5A,
        len: 1,
        target: [0x001E5B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5C,
        len: 1,
        target: [0x001E5D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5E,
        len: 1,
        target: [0x001E5F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E60,
        len: 1,
        target: [0x001E61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E62,
        len: 1,
        target: [0x001E63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E64,
        len: 1,
        target: [0x001E65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E66,
        len: 1,
        target: [0x001E67, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E68,
        len: 1,
        target: [0x001E69, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6A,
        len: 1,
        target: [0x001E6B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6C,
        len: 1,
        target: [0x001E6D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6E,
        len: 1,
        target: [0x001E6F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E70,
        len: 1,
        target: [0x001E71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E72,
        len: 1,
        target: [0x001E73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E74,
        len: 1,
        target: [0x001E75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E76,
        len: 1,
        target: [0x001E77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E78,
        len: 1,
        target: [0x001E79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7A,
        len: 1,
        target: [0x001E7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7C,
        len: 1,
        target: [0x001E7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7E,
        len: 1,
        target: [0x001E7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E80,
        len: 1,
        target: [0x001E81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E82,
        len: 1,
        target: [0x001E83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E84,
        len: 1,
        target: [0x001E85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E86,
        len: 1,
        target: [0x001E87, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E88,
        len: 1,
        target: [0x001E89, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8A,
        len: 1,
        target: [0x001E8B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8C,
        len: 1,
        target: [0x001E8D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8E,
        len: 1,
        target: [0x001E8F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E90,
        len: 1,
        target: [0x001E91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E92,
        len: 1,
        target: [0x001E93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E94,
        len: 1,
        target: [0x001E95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E96,
        len: 2,
        target: [0x000068, 0x000331, 0x000000],
    },
    Mapping {
        source: 0x001E97,
        len: 2,
        target: [0x000074, 0x000308, 0x000000],
    },
    Mapping {
        source: 0x001E98,
        len: 2,
        target: [0x000077, 0x00030A, 0x000000],
    },
    Mapping {
        source: 0x001E99,
        len: 2,
        target: [0x000079, 0x00030A, 0x000000],
    },
    Mapping {
        source: 0x001E9A,
        len: 2,
        target: [0x000061, 0x0002BE, 0x000000],
    },
    Mapping {
        source: 0x001E9B,
        len: 1,
        target: [0x001E61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E9E,
        len: 2,
        target: [0x000073, 0x000073, 0x000000],
    },
    Mapping {
        source: 0x001EA0,
        len: 1,
        target: [0x001EA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA2,
        len: 1,
        target: [0x001EA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA4,
        len: 1,
        target: [0x001EA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA6,
        len: 1,
        target: [0x001EA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA8,
        len: 1,
        target: [0x001EA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAA,
        len: 1,
        target: [0x001EAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAC,
        len: 1,
        target: [0x001EAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAE,
        len: 1,
        target: [0x001EAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB0,
        len: 1,
        target: [0x001EB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB2,
        len: 1,
        target: [0x001EB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB4,
        len: 1,
        target: [0x001EB5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB6,
        len: 1,
        target: [0x001EB7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB8,
        len: 1,
        target: [0x001EB9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBA,
        len: 1,
        target: [0x001EBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBC,
        len: 1,
        target: [0x001EBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBE,
        len: 1,
        target: [0x001EBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC0,
        len: 1,
        target: [0x001EC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC2,
        len: 1,
        target: [0x001EC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC4,
        len: 1,
        target: [0x001EC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC6,
        len: 1,
        target: [0x001EC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC8,
        len: 1,
        target: [0x001EC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECA,
        len: 1,
        target: [0x001ECB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECC,
        len: 1,
        target: [0x001ECD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECE,
        len: 1,
        target: [0x001ECF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED0,
        len: 1,
        target: [0x001ED1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED2,
        len: 1,
        target: [0x001ED3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED4,
        len: 1,
        target: [0x001ED5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED6,
        len: 1,
        target: [0x001ED7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED8,
        len: 1,
        target: [0x001ED9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDA,
        len: 1,
        target: [0x001EDB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDC,
        len: 1,
        target: [0x001EDD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDE,
        len: 1,
        target: [0x001EDF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE0,
        len: 1,
        target: [0x001EE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE2,
        len: 1,
        target: [0x001EE3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE4,
        len: 1,
        target: [0x001EE5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE6,
        len: 1,
        target: [0x001EE7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE8,
        len: 1,
        target: [0x001EE9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEA,
        len: 1,
        target: [0x001EEB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEC,
        len: 1,
        target: [0x001EED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEE,
        len: 1,
        target: [0x001EEF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF0,
        len: 1,
        target: [0x001EF1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF2,
        len: 1,
        target: [0x001EF3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF4,
        len: 1,
        target: [0x001EF5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF6,
        len: 1,
        target: [0x001EF7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF8,
        len: 1,
        target: [0x001EF9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFA,
        len: 1,
        target: [0x001EFB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFC,
        len: 1,
        target: [0x001EFD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFE,
        len: 1,
        target: [0x001EFF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F08,
        len: 1,
        target: [0x001F00, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F09,
        len: 1,
        target: [0x001F01, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0A,
        len: 1,
        target: [0x001F02, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0B,
        len: 1,
        target: [0x001F03, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0C,
        len: 1,
        target: [0x001F04, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0D,
        len: 1,
        target: [0x001F05, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0E,
        len: 1,
        target: [0x001F06, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0F,
        len: 1,
        target: [0x001F07, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F18,
        len: 1,
        target: [0x001F10, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F19,
        len: 1,
        target: [0x001F11, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1A,
        len: 1,
        target: [0x001F12, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1B,
        len: 1,
        target: [0x001F13, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1C,
        len: 1,
        target: [0x001F14, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1D,
        len: 1,
        target: [0x001F15, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F28,
        len: 1,
        target: [0x001F20, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F29,
        len: 1,
        target: [0x001F21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2A,
        len: 1,
        target: [0x001F22, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2B,
        len: 1,
        target: [0x001F23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2C,
        len: 1,
        target: [0x001F24, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2D,
        len: 1,
        target: [0x001F25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2E,
        len: 1,
        target: [0x001F26, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2F,
        len: 1,
        target: [0x001F27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F38,
        len: 1,
        target: [0x001F30, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F39,
        len: 1,
        target: [0x001F31, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3A,
        len: 1,
        target: [0x001F32, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3B,
        len: 1,
        target: [0x001F33, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3C,
        len: 1,
        target: [0x001F34, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3D,
        len: 1,
        target: [0x001F35, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3E,
        len: 1,
        target: [0x001F36, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3F,
        len: 1,
        target: [0x001F37, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F48,
        len: 1,
        target: [0x001F40, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F49,
        len: 1,
        target: [0x001F41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4A,
        len: 1,
        target: [0x001F42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4B,
        len: 1,
        target: [0x001F43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4C,
        len: 1,
        target: [0x001F44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4D,
        len: 1,
        target: [0x001F45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F50,
        len: 2,
        target: [0x0003C5, 0x000313, 0x000000],
    },
    Mapping {
        source: 0x001F52,
        len: 3,
        target: [0x0003C5, 0x000313, 0x000300],
    },
    Mapping {
        source: 0x001F54,
        len: 3,
        target: [0x0003C5, 0x000313, 0x000301],
    },
    Mapping {
        source: 0x001F56,
        len: 3,
        target: [0x0003C5, 0x000313, 0x000342],
    },
    Mapping {
        source: 0x001F59,
        len: 1,
        target: [0x001F51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F5B,
        len: 1,
        target: [0x001F53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F5D,
        len: 1,
        target: [0x001F55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F5F,
        len: 1,
        target: [0x001F57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F68,
        len: 1,
        target: [0x001F60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F69,
        len: 1,
        target: [0x001F61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6A,
        len: 1,
        target: [0x001F62, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6B,
        len: 1,
        target: [0x001F63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6C,
        len: 1,
        target: [0x001F64, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6D,
        len: 1,
        target: [0x001F65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6E,
        len: 1,
        target: [0x001F66, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6F,
        len: 1,
        target: [0x001F67, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F80,
        len: 2,
        target: [0x001F00, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F81,
        len: 2,
        target: [0x001F01, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F82,
        len: 2,
        target: [0x001F02, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F83,
        len: 2,
        target: [0x001F03, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F84,
        len: 2,
        target: [0x001F04, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F85,
        len: 2,
        target: [0x001F05, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F86,
        len: 2,
        target: [0x001F06, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F87,
        len: 2,
        target: [0x001F07, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F88,
        len: 2,
        target: [0x001F00, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F89,
        len: 2,
        target: [0x001F01, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F8A,
        len: 2,
        target: [0x001F02, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F8B,
        len: 2,
        target: [0x001F03, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F8C,
        len: 2,
        target: [0x001F04, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F8D,
        len: 2,
        target: [0x001F05, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F8E,
        len: 2,
        target: [0x001F06, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F8F,
        len: 2,
        target: [0x001F07, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F90,
        len: 2,
        target: [0x001F20, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F91,
        len: 2,
        target: [0x001F21, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F92,
        len: 2,
        target: [0x001F22, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F93,
        len: 2,
        target: [0x001F23, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F94,
        len: 2,
        target: [0x001F24, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F95,
        len: 2,
        target: [0x001F25, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F96,
        len: 2,
        target: [0x001F26, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F97,
        len: 2,
        target: [0x001F27, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F98,
        len: 2,
        target: [0x001F20, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F99,
        len: 2,
        target: [0x001F21, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F9A,
        len: 2,
        target: [0x001F22, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F9B,
        len: 2,
        target: [0x001F23, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F9C,
        len: 2,
        target: [0x001F24, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F9D,
        len: 2,
        target: [0x001F25, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F9E,
        len: 2,
        target: [0x001F26, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001F9F,
        len: 2,
        target: [0x001F27, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA0,
        len: 2,
        target: [0x001F60, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA1,
        len: 2,
        target: [0x001F61, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA2,
        len: 2,
        target: [0x001F62, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA3,
        len: 2,
        target: [0x001F63, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA4,
        len: 2,
        target: [0x001F64, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA5,
        len: 2,
        target: [0x001F65, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA6,
        len: 2,
        target: [0x001F66, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA7,
        len: 2,
        target: [0x001F67, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA8,
        len: 2,
        target: [0x001F60, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FA9,
        len: 2,
        target: [0x001F61, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FAA,
        len: 2,
        target: [0x001F62, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FAB,
        len: 2,
        target: [0x001F63, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FAC,
        len: 2,
        target: [0x001F64, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FAD,
        len: 2,
        target: [0x001F65, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FAE,
        len: 2,
        target: [0x001F66, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FAF,
        len: 2,
        target: [0x001F67, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FB2,
        len: 2,
        target: [0x001F70, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FB3,
        len: 2,
        target: [0x0003B1, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FB4,
        len: 2,
        target: [0x0003AC, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FB6,
        len: 2,
        target: [0x0003B1, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FB7,
        len: 3,
        target: [0x0003B1, 0x000342, 0x0003B9],
    },
    Mapping {
        source: 0x001FB8,
        len: 1,
        target: [0x001FB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FB9,
        len: 1,
        target: [0x001FB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FBA,
        len: 1,
        target: [0x001F70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FBB,
        len: 1,
        target: [0x001F71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FBC,
        len: 2,
        target: [0x0003B1, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FBE,
        len: 1,
        target: [0x0003B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FC2,
        len: 2,
        target: [0x001F74, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FC3,
        len: 2,
        target: [0x0003B7, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FC4,
        len: 2,
        target: [0x0003AE, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FC6,
        len: 2,
        target: [0x0003B7, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FC7,
        len: 3,
        target: [0x0003B7, 0x000342, 0x0003B9],
    },
    Mapping {
        source: 0x001FC8,
        len: 1,
        target: [0x001F72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FC9,
        len: 1,
        target: [0x001F73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FCA,
        len: 1,
        target: [0x001F74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FCB,
        len: 1,
        target: [0x001F75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FCC,
        len: 2,
        target: [0x0003B7, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FD2,
        len: 3,
        target: [0x0003B9, 0x000308, 0x000300],
    },
    Mapping {
        source: 0x001FD3,
        len: 3,
        target: [0x0003B9, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x001FD6,
        len: 2,
        target: [0x0003B9, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FD7,
        len: 3,
        target: [0x0003B9, 0x000308, 0x000342],
    },
    Mapping {
        source: 0x001FD8,
        len: 1,
        target: [0x001FD0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FD9,
        len: 1,
        target: [0x001FD1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FDA,
        len: 1,
        target: [0x001F76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FDB,
        len: 1,
        target: [0x001F77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FE2,
        len: 3,
        target: [0x0003C5, 0x000308, 0x000300],
    },
    Mapping {
        source: 0x001FE3,
        len: 3,
        target: [0x0003C5, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x001FE4,
        len: 2,
        target: [0x0003C1, 0x000313, 0x000000],
    },
    Mapping {
        source: 0x001FE6,
        len: 2,
        target: [0x0003C5, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FE7,
        len: 3,
        target: [0x0003C5, 0x000308, 0x000342],
    },
    Mapping {
        source: 0x001FE8,
        len: 1,
        target: [0x001FE0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FE9,
        len: 1,
        target: [0x001FE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FEA,
        len: 1,
        target: [0x001F7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FEB,
        len: 1,
        target: [0x001F7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FEC,
        len: 1,
        target: [0x001FE5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FF2,
        len: 2,
        target: [0x001F7C, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FF3,
        len: 2,
        target: [0x0003C9, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FF4,
        len: 2,
        target: [0x0003CE, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x001FF6,
        len: 2,
        target: [0x0003C9, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FF7,
        len: 3,
        target: [0x0003C9, 0x000342, 0x0003B9],
    },
    Mapping {
        source: 0x001FF8,
        len: 1,
        target: [0x001F78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FF9,
        len: 1,
        target: [0x001F79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FFA,
        len: 1,
        target: [0x001F7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FFB,
        len: 1,
        target: [0x001F7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FFC,
        len: 2,
        target: [0x0003C9, 0x0003B9, 0x000000],
    },
    Mapping {
        source: 0x002126,
        len: 1,
        target: [0x0003C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00212A,
        len: 1,
        target: [0x00006B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00212B,
        len: 1,
        target: [0x0000E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002132,
        len: 1,
        target: [0x00214E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002160,
        len: 1,
        target: [0x002170, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002161,
        len: 1,
        target: [0x002171, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002162,
        len: 1,
        target: [0x002172, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002163,
        len: 1,
        target: [0x002173, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002164,
        len: 1,
        target: [0x002174, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002165,
        len: 1,
        target: [0x002175, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002166,
        len: 1,
        target: [0x002176, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002167,
        len: 1,
        target: [0x002177, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002168,
        len: 1,
        target: [0x002178, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002169,
        len: 1,
        target: [0x002179, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216A,
        len: 1,
        target: [0x00217A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216B,
        len: 1,
        target: [0x00217B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216C,
        len: 1,
        target: [0x00217C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216D,
        len: 1,
        target: [0x00217D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216E,
        len: 1,
        target: [0x00217E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216F,
        len: 1,
        target: [0x00217F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002183,
        len: 1,
        target: [0x002184, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B6,
        len: 1,
        target: [0x0024D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B7,
        len: 1,
        target: [0x0024D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B8,
        len: 1,
        target: [0x0024D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B9,
        len: 1,
        target: [0x0024D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BA,
        len: 1,
        target: [0x0024D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BB,
        len: 1,
        target: [0x0024D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BC,
        len: 1,
        target: [0x0024D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BD,
        len: 1,
        target: [0x0024D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BE,
        len: 1,
        target: [0x0024D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BF,
        len: 1,
        target: [0x0024D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C0,
        len: 1,
        target: [0x0024DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C1,
        len: 1,
        target: [0x0024DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C2,
        len: 1,
        target: [0x0024DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C3,
        len: 1,
        target: [0x0024DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C4,
        len: 1,
        target: [0x0024DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C5,
        len: 1,
        target: [0x0024DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C6,
        len: 1,
        target: [0x0024E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C7,
        len: 1,
        target: [0x0024E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C8,
        len: 1,
        target: [0x0024E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C9,
        len: 1,
        target: [0x0024E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CA,
        len: 1,
        target: [0x0024E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CB,
        len: 1,
        target: [0x0024E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CC,
        len: 1,
        target: [0x0024E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CD,
        len: 1,
        target: [0x0024E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CE,
        len: 1,
        target: [0x0024E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CF,
        len: 1,
        target: [0x0024E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C00,
        len: 1,
        target: [0x002C30, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C01,
        len: 1,
        target: [0x002C31, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C02,
        len: 1,
        target: [0x002C32, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C03,
        len: 1,
        target: [0x002C33, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C04,
        len: 1,
        target: [0x002C34, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C05,
        len: 1,
        target: [0x002C35, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C06,
        len: 1,
        target: [0x002C36, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C07,
        len: 1,
        target: [0x002C37, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C08,
        len: 1,
        target: [0x002C38, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C09,
        len: 1,
        target: [0x002C39, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0A,
        len: 1,
        target: [0x002C3A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0B,
        len: 1,
        target: [0x002C3B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0C,
        len: 1,
        target: [0x002C3C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0D,
        len: 1,
        target: [0x002C3D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0E,
        len: 1,
        target: [0x002C3E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0F,
        len: 1,
        target: [0x002C3F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C10,
        len: 1,
        target: [0x002C40, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C11,
        len: 1,
        target: [0x002C41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C12,
        len: 1,
        target: [0x002C42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C13,
        len: 1,
        target: [0x002C43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C14,
        len: 1,
        target: [0x002C44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C15,
        len: 1,
        target: [0x002C45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C16,
        len: 1,
        target: [0x002C46, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C17,
        len: 1,
        target: [0x002C47, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C18,
        len: 1,
        target: [0x002C48, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C19,
        len: 1,
        target: [0x002C49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1A,
        len: 1,
        target: [0x002C4A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1B,
        len: 1,
        target: [0x002C4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1C,
        len: 1,
        target: [0x002C4C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1D,
        len: 1,
        target: [0x002C4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1E,
        len: 1,
        target: [0x002C4E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1F,
        len: 1,
        target: [0x002C4F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C20,
        len: 1,
        target: [0x002C50, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C21,
        len: 1,
        target: [0x002C51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C22,
        len: 1,
        target: [0x002C52, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C23,
        len: 1,
        target: [0x002C53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C24,
        len: 1,
        target: [0x002C54, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C25,
        len: 1,
        target: [0x002C55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C26,
        len: 1,
        target: [0x002C56, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C27,
        len: 1,
        target: [0x002C57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C28,
        len: 1,
        target: [0x002C58, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C29,
        len: 1,
        target: [0x002C59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2A,
        len: 1,
        target: [0x002C5A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2B,
        len: 1,
        target: [0x002C5B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2C,
        len: 1,
        target: [0x002C5C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2D,
        len: 1,
        target: [0x002C5D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2E,
        len: 1,
        target: [0x002C5E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2F,
        len: 1,
        target: [0x002C5F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C60,
        len: 1,
        target: [0x002C61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C62,
        len: 1,
        target: [0x00026B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C63,
        len: 1,
        target: [0x001D7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C64,
        len: 1,
        target: [0x00027D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C67,
        len: 1,
        target: [0x002C68, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C69,
        len: 1,
        target: [0x002C6A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6B,
        len: 1,
        target: [0x002C6C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6D,
        len: 1,
        target: [0x000251, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6E,
        len: 1,
        target: [0x000271, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6F,
        len: 1,
        target: [0x000250, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C70,
        len: 1,
        target: [0x000252, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C72,
        len: 1,
        target: [0x002C73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C75,
        len: 1,
        target: [0x002C76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C7E,
        len: 1,
        target: [0x00023F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C7F,
        len: 1,
        target: [0x000240, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C80,
        len: 1,
        target: [0x002C81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C82,
        len: 1,
        target: [0x002C83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C84,
        len: 1,
        target: [0x002C85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C86,
        len: 1,
        target: [0x002C87, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C88,
        len: 1,
        target: [0x002C89, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8A,
        len: 1,
        target: [0x002C8B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8C,
        len: 1,
        target: [0x002C8D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8E,
        len: 1,
        target: [0x002C8F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C90,
        len: 1,
        target: [0x002C91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C92,
        len: 1,
        target: [0x002C93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C94,
        len: 1,
        target: [0x002C95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C96,
        len: 1,
        target: [0x002C97, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C98,
        len: 1,
        target: [0x002C99, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9A,
        len: 1,
        target: [0x002C9B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9C,
        len: 1,
        target: [0x002C9D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9E,
        len: 1,
        target: [0x002C9F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA0,
        len: 1,
        target: [0x002CA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA2,
        len: 1,
        target: [0x002CA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA4,
        len: 1,
        target: [0x002CA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA6,
        len: 1,
        target: [0x002CA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA8,
        len: 1,
        target: [0x002CA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAA,
        len: 1,
        target: [0x002CAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAC,
        len: 1,
        target: [0x002CAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAE,
        len: 1,
        target: [0x002CAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB0,
        len: 1,
        target: [0x002CB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB2,
        len: 1,
        target: [0x002CB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB4,
        len: 1,
        target: [0x002CB5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB6,
        len: 1,
        target: [0x002CB7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB8,
        len: 1,
        target: [0x002CB9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBA,
        len: 1,
        target: [0x002CBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBC,
        len: 1,
        target: [0x002CBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBE,
        len: 1,
        target: [0x002CBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC0,
        len: 1,
        target: [0x002CC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC2,
        len: 1,
        target: [0x002CC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC4,
        len: 1,
        target: [0x002CC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC6,
        len: 1,
        target: [0x002CC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC8,
        len: 1,
        target: [0x002CC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCA,
        len: 1,
        target: [0x002CCB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCC,
        len: 1,
        target: [0x002CCD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCE,
        len: 1,
        target: [0x002CCF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD0,
        len: 1,
        target: [0x002CD1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD2,
        len: 1,
        target: [0x002CD3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD4,
        len: 1,
        target: [0x002CD5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD6,
        len: 1,
        target: [0x002CD7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD8,
        len: 1,
        target: [0x002CD9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDA,
        len: 1,
        target: [0x002CDB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDC,
        len: 1,
        target: [0x002CDD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDE,
        len: 1,
        target: [0x002CDF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CE0,
        len: 1,
        target: [0x002CE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CE2,
        len: 1,
        target: [0x002CE3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CEB,
        len: 1,
        target: [0x002CEC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CED,
        len: 1,
        target: [0x002CEE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CF2,
        len: 1,
        target: [0x002CF3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A640,
        len: 1,
        target: [0x00A641, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A642,
        len: 1,
        target: [0x00A643, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A644,
        len: 1,
        target: [0x00A645, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A646,
        len: 1,
        target: [0x00A647, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A648,
        len: 1,
        target: [0x00A649, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64A,
        len: 1,
        target: [0x00A64B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64C,
        len: 1,
        target: [0x00A64D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64E,
        len: 1,
        target: [0x00A64F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A650,
        len: 1,
        target: [0x00A651, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A652,
        len: 1,
        target: [0x00A653, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A654,
        len: 1,
        target: [0x00A655, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A656,
        len: 1,
        target: [0x00A657, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A658,
        len: 1,
        target: [0x00A659, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65A,
        len: 1,
        target: [0x00A65B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65C,
        len: 1,
        target: [0x00A65D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65E,
        len: 1,
        target: [0x00A65F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A660,
        len: 1,
        target: [0x00A661, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A662,
        len: 1,
        target: [0x00A663, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A664,
        len: 1,
        target: [0x00A665, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A666,
        len: 1,
        target: [0x00A667, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A668,
        len: 1,
        target: [0x00A669, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A66A,
        len: 1,
        target: [0x00A66B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A66C,
        len: 1,
        target: [0x00A66D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A680,
        len: 1,
        target: [0x00A681, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A682,
        len: 1,
        target: [0x00A683, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A684,
        len: 1,
        target: [0x00A685, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A686,
        len: 1,
        target: [0x00A687, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A688,
        len: 1,
        target: [0x00A689, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68A,
        len: 1,
        target: [0x00A68B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68C,
        len: 1,
        target: [0x00A68D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68E,
        len: 1,
        target: [0x00A68F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A690,
        len: 1,
        target: [0x00A691, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A692,
        len: 1,
        target: [0x00A693, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A694,
        len: 1,
        target: [0x00A695, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A696,
        len: 1,
        target: [0x00A697, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A698,
        len: 1,
        target: [0x00A699, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A69A,
        len: 1,
        target: [0x00A69B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A722,
        len: 1,
        target: [0x00A723, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A724,
        len: 1,
        target: [0x00A725, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A726,
        len: 1,
        target: [0x00A727, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A728,
        len: 1,
        target: [0x00A729, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72A,
        len: 1,
        target: [0x00A72B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72C,
        len: 1,
        target: [0x00A72D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72E,
        len: 1,
        target: [0x00A72F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A732,
        len: 1,
        target: [0x00A733, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A734,
        len: 1,
        target: [0x00A735, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A736,
        len: 1,
        target: [0x00A737, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A738,
        len: 1,
        target: [0x00A739, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73A,
        len: 1,
        target: [0x00A73B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73C,
        len: 1,
        target: [0x00A73D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73E,
        len: 1,
        target: [0x00A73F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A740,
        len: 1,
        target: [0x00A741, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A742,
        len: 1,
        target: [0x00A743, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A744,
        len: 1,
        target: [0x00A745, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A746,
        len: 1,
        target: [0x00A747, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A748,
        len: 1,
        target: [0x00A749, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74A,
        len: 1,
        target: [0x00A74B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74C,
        len: 1,
        target: [0x00A74D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74E,
        len: 1,
        target: [0x00A74F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A750,
        len: 1,
        target: [0x00A751, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A752,
        len: 1,
        target: [0x00A753, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A754,
        len: 1,
        target: [0x00A755, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A756,
        len: 1,
        target: [0x00A757, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A758,
        len: 1,
        target: [0x00A759, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75A,
        len: 1,
        target: [0x00A75B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75C,
        len: 1,
        target: [0x00A75D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75E,
        len: 1,
        target: [0x00A75F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A760,
        len: 1,
        target: [0x00A761, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A762,
        len: 1,
        target: [0x00A763, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A764,
        len: 1,
        target: [0x00A765, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A766,
        len: 1,
        target: [0x00A767, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A768,
        len: 1,
        target: [0x00A769, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76A,
        len: 1,
        target: [0x00A76B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76C,
        len: 1,
        target: [0x00A76D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76E,
        len: 1,
        target: [0x00A76F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A779,
        len: 1,
        target: [0x00A77A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77B,
        len: 1,
        target: [0x00A77C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77D,
        len: 1,
        target: [0x001D79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77E,
        len: 1,
        target: [0x00A77F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A780,
        len: 1,
        target: [0x00A781, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A782,
        len: 1,
        target: [0x00A783, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A784,
        len: 1,
        target: [0x00A785, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A786,
        len: 1,
        target: [0x00A787, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A78B,
        len: 1,
        target: [0x00A78C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A78D,
        len: 1,
        target: [0x000265, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A790,
        len: 1,
        target: [0x00A791, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A792,
        len: 1,
        target: [0x00A793, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A796,
        len: 1,
        target: [0x00A797, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A798,
        len: 1,
        target: [0x00A799, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79A,
        len: 1,
        target: [0x00A79B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79C,
        len: 1,
        target: [0x00A79D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79E,
        len: 1,
        target: [0x00A79F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A0,
        len: 1,
        target: [0x00A7A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A2,
        len: 1,
        target: [0x00A7A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A4,
        len: 1,
        target: [0x00A7A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A6,
        len: 1,
        target: [0x00A7A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A8,
        len: 1,
        target: [0x00A7A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AA,
        len: 1,
        target: [0x000266, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AB,
        len: 1,
        target: [0x00025C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AC,
        len: 1,
        target: [0x000261, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AD,
        len: 1,
        target: [0x00026C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AE,
        len: 1,
        target: [0x00026A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B0,
        len: 1,
        target: [0x00029E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B1,
        len: 1,
        target: [0x000287, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B2,
        len: 1,
        target: [0x00029D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B3,
        len: 1,
        target: [0x00AB53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B4,
        len: 1,
        target: [0x00A7B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B6,
        len: 1,
        target: [0x00A7B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B8,
        len: 1,
        target: [0x00A7B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BA,
        len: 1,
        target: [0x00A7BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BC,
        len: 1,
        target: [0x00A7BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BE,
        len: 1,
        target: [0x00A7BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C0,
        len: 1,
        target: [0x00A7C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C2,
        len: 1,
        target: [0x00A7C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C4,
        len: 1,
        target: [0x00A794, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C5,
        len: 1,
        target: [0x000282, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C6,
        len: 1,
        target: [0x001D8E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C7,
        len: 1,
        target: [0x00A7C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C9,
        len: 1,
        target: [0x00A7CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CB,
        len: 1,
        target: [0x000264, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CC,
        len: 1,
        target: [0x00A7CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CE,
        len: 1,
        target: [0x00A7CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D0,
        len: 1,
        target: [0x00A7D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D2,
        len: 1,
        target: [0x00A7D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D4,
        len: 1,
        target: [0x00A7D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D6,
        len: 1,
        target: [0x00A7D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D8,
        len: 1,
        target: [0x00A7D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7DA,
        len: 1,
        target: [0x00A7DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7DC,
        len: 1,
        target: [0x00019B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7F5,
        len: 1,
        target: [0x00A7F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB70,
        len: 1,
        target: [0x0013A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB71,
        len: 1,
        target: [0x0013A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB72,
        len: 1,
        target: [0x0013A2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB73,
        len: 1,
        target: [0x0013A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB74,
        len: 1,
        target: [0x0013A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB75,
        len: 1,
        target: [0x0013A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB76,
        len: 1,
        target: [0x0013A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB77,
        len: 1,
        target: [0x0013A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB78,
        len: 1,
        target: [0x0013A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB79,
        len: 1,
        target: [0x0013A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7A,
        len: 1,
        target: [0x0013AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7B,
        len: 1,
        target: [0x0013AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7C,
        len: 1,
        target: [0x0013AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7D,
        len: 1,
        target: [0x0013AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7E,
        len: 1,
        target: [0x0013AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7F,
        len: 1,
        target: [0x0013AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB80,
        len: 1,
        target: [0x0013B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB81,
        len: 1,
        target: [0x0013B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB82,
        len: 1,
        target: [0x0013B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB83,
        len: 1,
        target: [0x0013B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB84,
        len: 1,
        target: [0x0013B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB85,
        len: 1,
        target: [0x0013B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB86,
        len: 1,
        target: [0x0013B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB87,
        len: 1,
        target: [0x0013B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB88,
        len: 1,
        target: [0x0013B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB89,
        len: 1,
        target: [0x0013B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8A,
        len: 1,
        target: [0x0013BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8B,
        len: 1,
        target: [0x0013BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8C,
        len: 1,
        target: [0x0013BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8D,
        len: 1,
        target: [0x0013BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8E,
        len: 1,
        target: [0x0013BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8F,
        len: 1,
        target: [0x0013BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB90,
        len: 1,
        target: [0x0013C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB91,
        len: 1,
        target: [0x0013C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB92,
        len: 1,
        target: [0x0013C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB93,
        len: 1,
        target: [0x0013C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB94,
        len: 1,
        target: [0x0013C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB95,
        len: 1,
        target: [0x0013C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB96,
        len: 1,
        target: [0x0013C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB97,
        len: 1,
        target: [0x0013C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB98,
        len: 1,
        target: [0x0013C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB99,
        len: 1,
        target: [0x0013C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9A,
        len: 1,
        target: [0x0013CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9B,
        len: 1,
        target: [0x0013CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9C,
        len: 1,
        target: [0x0013CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9D,
        len: 1,
        target: [0x0013CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9E,
        len: 1,
        target: [0x0013CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9F,
        len: 1,
        target: [0x0013CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA0,
        len: 1,
        target: [0x0013D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA1,
        len: 1,
        target: [0x0013D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA2,
        len: 1,
        target: [0x0013D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA3,
        len: 1,
        target: [0x0013D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA4,
        len: 1,
        target: [0x0013D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA5,
        len: 1,
        target: [0x0013D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA6,
        len: 1,
        target: [0x0013D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA7,
        len: 1,
        target: [0x0013D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA8,
        len: 1,
        target: [0x0013D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA9,
        len: 1,
        target: [0x0013D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAA,
        len: 1,
        target: [0x0013DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAB,
        len: 1,
        target: [0x0013DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAC,
        len: 1,
        target: [0x0013DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAD,
        len: 1,
        target: [0x0013DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAE,
        len: 1,
        target: [0x0013DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAF,
        len: 1,
        target: [0x0013DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB0,
        len: 1,
        target: [0x0013E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB1,
        len: 1,
        target: [0x0013E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB2,
        len: 1,
        target: [0x0013E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB3,
        len: 1,
        target: [0x0013E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB4,
        len: 1,
        target: [0x0013E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB5,
        len: 1,
        target: [0x0013E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB6,
        len: 1,
        target: [0x0013E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB7,
        len: 1,
        target: [0x0013E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB8,
        len: 1,
        target: [0x0013E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB9,
        len: 1,
        target: [0x0013E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBA,
        len: 1,
        target: [0x0013EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBB,
        len: 1,
        target: [0x0013EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBC,
        len: 1,
        target: [0x0013EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBD,
        len: 1,
        target: [0x0013ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBE,
        len: 1,
        target: [0x0013EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBF,
        len: 1,
        target: [0x0013EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FB00,
        len: 2,
        target: [0x000066, 0x000066, 0x000000],
    },
    Mapping {
        source: 0x00FB01,
        len: 2,
        target: [0x000066, 0x000069, 0x000000],
    },
    Mapping {
        source: 0x00FB02,
        len: 2,
        target: [0x000066, 0x00006C, 0x000000],
    },
    Mapping {
        source: 0x00FB03,
        len: 3,
        target: [0x000066, 0x000066, 0x000069],
    },
    Mapping {
        source: 0x00FB04,
        len: 3,
        target: [0x000066, 0x000066, 0x00006C],
    },
    Mapping {
        source: 0x00FB05,
        len: 2,
        target: [0x000073, 0x000074, 0x000000],
    },
    Mapping {
        source: 0x00FB06,
        len: 2,
        target: [0x000073, 0x000074, 0x000000],
    },
    Mapping {
        source: 0x00FB13,
        len: 2,
        target: [0x000574, 0x000576, 0x000000],
    },
    Mapping {
        source: 0x00FB14,
        len: 2,
        target: [0x000574, 0x000565, 0x000000],
    },
    Mapping {
        source: 0x00FB15,
        len: 2,
        target: [0x000574, 0x00056B, 0x000000],
    },
    Mapping {
        source: 0x00FB16,
        len: 2,
        target: [0x00057E, 0x000576, 0x000000],
    },
    Mapping {
        source: 0x00FB17,
        len: 2,
        target: [0x000574, 0x00056D, 0x000000],
    },
    Mapping {
        source: 0x00FF21,
        len: 1,
        target: [0x00FF41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF22,
        len: 1,
        target: [0x00FF42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF23,
        len: 1,
        target: [0x00FF43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF24,
        len: 1,
        target: [0x00FF44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF25,
        len: 1,
        target: [0x00FF45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF26,
        len: 1,
        target: [0x00FF46, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF27,
        len: 1,
        target: [0x00FF47, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF28,
        len: 1,
        target: [0x00FF48, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF29,
        len: 1,
        target: [0x00FF49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2A,
        len: 1,
        target: [0x00FF4A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2B,
        len: 1,
        target: [0x00FF4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2C,
        len: 1,
        target: [0x00FF4C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2D,
        len: 1,
        target: [0x00FF4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2E,
        len: 1,
        target: [0x00FF4E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2F,
        len: 1,
        target: [0x00FF4F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF30,
        len: 1,
        target: [0x00FF50, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF31,
        len: 1,
        target: [0x00FF51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF32,
        len: 1,
        target: [0x00FF52, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF33,
        len: 1,
        target: [0x00FF53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF34,
        len: 1,
        target: [0x00FF54, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF35,
        len: 1,
        target: [0x00FF55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF36,
        len: 1,
        target: [0x00FF56, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF37,
        len: 1,
        target: [0x00FF57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF38,
        len: 1,
        target: [0x00FF58, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF39,
        len: 1,
        target: [0x00FF59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF3A,
        len: 1,
        target: [0x00FF5A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010400,
        len: 1,
        target: [0x010428, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010401,
        len: 1,
        target: [0x010429, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010402,
        len: 1,
        target: [0x01042A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010403,
        len: 1,
        target: [0x01042B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010404,
        len: 1,
        target: [0x01042C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010405,
        len: 1,
        target: [0x01042D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010406,
        len: 1,
        target: [0x01042E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010407,
        len: 1,
        target: [0x01042F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010408,
        len: 1,
        target: [0x010430, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010409,
        len: 1,
        target: [0x010431, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040A,
        len: 1,
        target: [0x010432, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040B,
        len: 1,
        target: [0x010433, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040C,
        len: 1,
        target: [0x010434, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040D,
        len: 1,
        target: [0x010435, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040E,
        len: 1,
        target: [0x010436, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040F,
        len: 1,
        target: [0x010437, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010410,
        len: 1,
        target: [0x010438, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010411,
        len: 1,
        target: [0x010439, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010412,
        len: 1,
        target: [0x01043A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010413,
        len: 1,
        target: [0x01043B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010414,
        len: 1,
        target: [0x01043C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010415,
        len: 1,
        target: [0x01043D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010416,
        len: 1,
        target: [0x01043E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010417,
        len: 1,
        target: [0x01043F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010418,
        len: 1,
        target: [0x010440, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010419,
        len: 1,
        target: [0x010441, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041A,
        len: 1,
        target: [0x010442, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041B,
        len: 1,
        target: [0x010443, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041C,
        len: 1,
        target: [0x010444, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041D,
        len: 1,
        target: [0x010445, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041E,
        len: 1,
        target: [0x010446, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041F,
        len: 1,
        target: [0x010447, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010420,
        len: 1,
        target: [0x010448, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010421,
        len: 1,
        target: [0x010449, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010422,
        len: 1,
        target: [0x01044A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010423,
        len: 1,
        target: [0x01044B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010424,
        len: 1,
        target: [0x01044C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010425,
        len: 1,
        target: [0x01044D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010426,
        len: 1,
        target: [0x01044E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010427,
        len: 1,
        target: [0x01044F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B0,
        len: 1,
        target: [0x0104D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B1,
        len: 1,
        target: [0x0104D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B2,
        len: 1,
        target: [0x0104DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B3,
        len: 1,
        target: [0x0104DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B4,
        len: 1,
        target: [0x0104DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B5,
        len: 1,
        target: [0x0104DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B6,
        len: 1,
        target: [0x0104DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B7,
        len: 1,
        target: [0x0104DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B8,
        len: 1,
        target: [0x0104E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B9,
        len: 1,
        target: [0x0104E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BA,
        len: 1,
        target: [0x0104E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BB,
        len: 1,
        target: [0x0104E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BC,
        len: 1,
        target: [0x0104E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BD,
        len: 1,
        target: [0x0104E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BE,
        len: 1,
        target: [0x0104E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BF,
        len: 1,
        target: [0x0104E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C0,
        len: 1,
        target: [0x0104E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C1,
        len: 1,
        target: [0x0104E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C2,
        len: 1,
        target: [0x0104EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C3,
        len: 1,
        target: [0x0104EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C4,
        len: 1,
        target: [0x0104EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C5,
        len: 1,
        target: [0x0104ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C6,
        len: 1,
        target: [0x0104EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C7,
        len: 1,
        target: [0x0104EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C8,
        len: 1,
        target: [0x0104F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C9,
        len: 1,
        target: [0x0104F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CA,
        len: 1,
        target: [0x0104F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CB,
        len: 1,
        target: [0x0104F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CC,
        len: 1,
        target: [0x0104F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CD,
        len: 1,
        target: [0x0104F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CE,
        len: 1,
        target: [0x0104F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CF,
        len: 1,
        target: [0x0104F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D0,
        len: 1,
        target: [0x0104F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D1,
        len: 1,
        target: [0x0104F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D2,
        len: 1,
        target: [0x0104FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D3,
        len: 1,
        target: [0x0104FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010570,
        len: 1,
        target: [0x010597, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010571,
        len: 1,
        target: [0x010598, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010572,
        len: 1,
        target: [0x010599, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010573,
        len: 1,
        target: [0x01059A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010574,
        len: 1,
        target: [0x01059B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010575,
        len: 1,
        target: [0x01059C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010576,
        len: 1,
        target: [0x01059D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010577,
        len: 1,
        target: [0x01059E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010578,
        len: 1,
        target: [0x01059F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010579,
        len: 1,
        target: [0x0105A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057A,
        len: 1,
        target: [0x0105A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057C,
        len: 1,
        target: [0x0105A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057D,
        len: 1,
        target: [0x0105A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057E,
        len: 1,
        target: [0x0105A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057F,
        len: 1,
        target: [0x0105A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010580,
        len: 1,
        target: [0x0105A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010581,
        len: 1,
        target: [0x0105A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010582,
        len: 1,
        target: [0x0105A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010583,
        len: 1,
        target: [0x0105AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010584,
        len: 1,
        target: [0x0105AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010585,
        len: 1,
        target: [0x0105AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010586,
        len: 1,
        target: [0x0105AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010587,
        len: 1,
        target: [0x0105AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010588,
        len: 1,
        target: [0x0105AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010589,
        len: 1,
        target: [0x0105B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058A,
        len: 1,
        target: [0x0105B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058C,
        len: 1,
        target: [0x0105B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058D,
        len: 1,
        target: [0x0105B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058E,
        len: 1,
        target: [0x0105B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058F,
        len: 1,
        target: [0x0105B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010590,
        len: 1,
        target: [0x0105B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010591,
        len: 1,
        target: [0x0105B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010592,
        len: 1,
        target: [0x0105B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010594,
        len: 1,
        target: [0x0105BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010595,
        len: 1,
        target: [0x0105BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C80,
        len: 1,
        target: [0x010CC0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C81,
        len: 1,
        target: [0x010CC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C82,
        len: 1,
        target: [0x010CC2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C83,
        len: 1,
        target: [0x010CC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C84,
        len: 1,
        target: [0x010CC4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C85,
        len: 1,
        target: [0x010CC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C86,
        len: 1,
        target: [0x010CC6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C87,
        len: 1,
        target: [0x010CC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C88,
        len: 1,
        target: [0x010CC8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C89,
        len: 1,
        target: [0x010CC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8A,
        len: 1,
        target: [0x010CCA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8B,
        len: 1,
        target: [0x010CCB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8C,
        len: 1,
        target: [0x010CCC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8D,
        len: 1,
        target: [0x010CCD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8E,
        len: 1,
        target: [0x010CCE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8F,
        len: 1,
        target: [0x010CCF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C90,
        len: 1,
        target: [0x010CD0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C91,
        len: 1,
        target: [0x010CD1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C92,
        len: 1,
        target: [0x010CD2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C93,
        len: 1,
        target: [0x010CD3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C94,
        len: 1,
        target: [0x010CD4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C95,
        len: 1,
        target: [0x010CD5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C96,
        len: 1,
        target: [0x010CD6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C97,
        len: 1,
        target: [0x010CD7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C98,
        len: 1,
        target: [0x010CD8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C99,
        len: 1,
        target: [0x010CD9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9A,
        len: 1,
        target: [0x010CDA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9B,
        len: 1,
        target: [0x010CDB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9C,
        len: 1,
        target: [0x010CDC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9D,
        len: 1,
        target: [0x010CDD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9E,
        len: 1,
        target: [0x010CDE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9F,
        len: 1,
        target: [0x010CDF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA0,
        len: 1,
        target: [0x010CE0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA1,
        len: 1,
        target: [0x010CE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA2,
        len: 1,
        target: [0x010CE2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA3,
        len: 1,
        target: [0x010CE3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA4,
        len: 1,
        target: [0x010CE4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA5,
        len: 1,
        target: [0x010CE5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA6,
        len: 1,
        target: [0x010CE6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA7,
        len: 1,
        target: [0x010CE7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA8,
        len: 1,
        target: [0x010CE8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA9,
        len: 1,
        target: [0x010CE9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAA,
        len: 1,
        target: [0x010CEA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAB,
        len: 1,
        target: [0x010CEB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAC,
        len: 1,
        target: [0x010CEC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAD,
        len: 1,
        target: [0x010CED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAE,
        len: 1,
        target: [0x010CEE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAF,
        len: 1,
        target: [0x010CEF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CB0,
        len: 1,
        target: [0x010CF0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CB1,
        len: 1,
        target: [0x010CF1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CB2,
        len: 1,
        target: [0x010CF2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D50,
        len: 1,
        target: [0x010D70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D51,
        len: 1,
        target: [0x010D71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D52,
        len: 1,
        target: [0x010D72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D53,
        len: 1,
        target: [0x010D73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D54,
        len: 1,
        target: [0x010D74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D55,
        len: 1,
        target: [0x010D75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D56,
        len: 1,
        target: [0x010D76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D57,
        len: 1,
        target: [0x010D77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D58,
        len: 1,
        target: [0x010D78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D59,
        len: 1,
        target: [0x010D79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5A,
        len: 1,
        target: [0x010D7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5B,
        len: 1,
        target: [0x010D7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5C,
        len: 1,
        target: [0x010D7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5D,
        len: 1,
        target: [0x010D7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5E,
        len: 1,
        target: [0x010D7E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5F,
        len: 1,
        target: [0x010D7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D60,
        len: 1,
        target: [0x010D80, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D61,
        len: 1,
        target: [0x010D81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D62,
        len: 1,
        target: [0x010D82, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D63,
        len: 1,
        target: [0x010D83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D64,
        len: 1,
        target: [0x010D84, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D65,
        len: 1,
        target: [0x010D85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A0,
        len: 1,
        target: [0x0118C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A1,
        len: 1,
        target: [0x0118C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A2,
        len: 1,
        target: [0x0118C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A3,
        len: 1,
        target: [0x0118C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A4,
        len: 1,
        target: [0x0118C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A5,
        len: 1,
        target: [0x0118C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A6,
        len: 1,
        target: [0x0118C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A7,
        len: 1,
        target: [0x0118C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A8,
        len: 1,
        target: [0x0118C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A9,
        len: 1,
        target: [0x0118C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AA,
        len: 1,
        target: [0x0118CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AB,
        len: 1,
        target: [0x0118CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AC,
        len: 1,
        target: [0x0118CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AD,
        len: 1,
        target: [0x0118CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AE,
        len: 1,
        target: [0x0118CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AF,
        len: 1,
        target: [0x0118CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B0,
        len: 1,
        target: [0x0118D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B1,
        len: 1,
        target: [0x0118D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B2,
        len: 1,
        target: [0x0118D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B3,
        len: 1,
        target: [0x0118D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B4,
        len: 1,
        target: [0x0118D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B5,
        len: 1,
        target: [0x0118D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B6,
        len: 1,
        target: [0x0118D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B7,
        len: 1,
        target: [0x0118D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B8,
        len: 1,
        target: [0x0118D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B9,
        len: 1,
        target: [0x0118D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BA,
        len: 1,
        target: [0x0118DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BB,
        len: 1,
        target: [0x0118DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BC,
        len: 1,
        target: [0x0118DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BD,
        len: 1,
        target: [0x0118DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BE,
        len: 1,
        target: [0x0118DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BF,
        len: 1,
        target: [0x0118DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E40,
        len: 1,
        target: [0x016E60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E41,
        len: 1,
        target: [0x016E61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E42,
        len: 1,
        target: [0x016E62, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E43,
        len: 1,
        target: [0x016E63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E44,
        len: 1,
        target: [0x016E64, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E45,
        len: 1,
        target: [0x016E65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E46,
        len: 1,
        target: [0x016E66, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E47,
        len: 1,
        target: [0x016E67, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E48,
        len: 1,
        target: [0x016E68, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E49,
        len: 1,
        target: [0x016E69, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4A,
        len: 1,
        target: [0x016E6A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4B,
        len: 1,
        target: [0x016E6B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4C,
        len: 1,
        target: [0x016E6C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4D,
        len: 1,
        target: [0x016E6D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4E,
        len: 1,
        target: [0x016E6E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4F,
        len: 1,
        target: [0x016E6F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E50,
        len: 1,
        target: [0x016E70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E51,
        len: 1,
        target: [0x016E71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E52,
        len: 1,
        target: [0x016E72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E53,
        len: 1,
        target: [0x016E73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E54,
        len: 1,
        target: [0x016E74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E55,
        len: 1,
        target: [0x016E75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E56,
        len: 1,
        target: [0x016E76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E57,
        len: 1,
        target: [0x016E77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E58,
        len: 1,
        target: [0x016E78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E59,
        len: 1,
        target: [0x016E79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5A,
        len: 1,
        target: [0x016E7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5B,
        len: 1,
        target: [0x016E7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5C,
        len: 1,
        target: [0x016E7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5D,
        len: 1,
        target: [0x016E7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5E,
        len: 1,
        target: [0x016E7E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5F,
        len: 1,
        target: [0x016E7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA0,
        len: 1,
        target: [0x016EBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA1,
        len: 1,
        target: [0x016EBC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA2,
        len: 1,
        target: [0x016EBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA3,
        len: 1,
        target: [0x016EBE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA4,
        len: 1,
        target: [0x016EBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA5,
        len: 1,
        target: [0x016EC0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA6,
        len: 1,
        target: [0x016EC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA7,
        len: 1,
        target: [0x016EC2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA8,
        len: 1,
        target: [0x016EC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA9,
        len: 1,
        target: [0x016EC4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAA,
        len: 1,
        target: [0x016EC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAB,
        len: 1,
        target: [0x016EC6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAC,
        len: 1,
        target: [0x016EC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAD,
        len: 1,
        target: [0x016EC8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAE,
        len: 1,
        target: [0x016EC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAF,
        len: 1,
        target: [0x016ECA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB0,
        len: 1,
        target: [0x016ECB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB1,
        len: 1,
        target: [0x016ECC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB2,
        len: 1,
        target: [0x016ECD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB3,
        len: 1,
        target: [0x016ECE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB4,
        len: 1,
        target: [0x016ECF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB5,
        len: 1,
        target: [0x016ED0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB6,
        len: 1,
        target: [0x016ED1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB7,
        len: 1,
        target: [0x016ED2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB8,
        len: 1,
        target: [0x016ED3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E900,
        len: 1,
        target: [0x01E922, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E901,
        len: 1,
        target: [0x01E923, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E902,
        len: 1,
        target: [0x01E924, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E903,
        len: 1,
        target: [0x01E925, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E904,
        len: 1,
        target: [0x01E926, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E905,
        len: 1,
        target: [0x01E927, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E906,
        len: 1,
        target: [0x01E928, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E907,
        len: 1,
        target: [0x01E929, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E908,
        len: 1,
        target: [0x01E92A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E909,
        len: 1,
        target: [0x01E92B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90A,
        len: 1,
        target: [0x01E92C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90B,
        len: 1,
        target: [0x01E92D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90C,
        len: 1,
        target: [0x01E92E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90D,
        len: 1,
        target: [0x01E92F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90E,
        len: 1,
        target: [0x01E930, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90F,
        len: 1,
        target: [0x01E931, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E910,
        len: 1,
        target: [0x01E932, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E911,
        len: 1,
        target: [0x01E933, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E912,
        len: 1,
        target: [0x01E934, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E913,
        len: 1,
        target: [0x01E935, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E914,
        len: 1,
        target: [0x01E936, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E915,
        len: 1,
        target: [0x01E937, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E916,
        len: 1,
        target: [0x01E938, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E917,
        len: 1,
        target: [0x01E939, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E918,
        len: 1,
        target: [0x01E93A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E919,
        len: 1,
        target: [0x01E93B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91A,
        len: 1,
        target: [0x01E93C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91B,
        len: 1,
        target: [0x01E93D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91C,
        len: 1,
        target: [0x01E93E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91D,
        len: 1,
        target: [0x01E93F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91E,
        len: 1,
        target: [0x01E940, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91F,
        len: 1,
        target: [0x01E941, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E920,
        len: 1,
        target: [0x01E942, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E921,
        len: 1,
        target: [0x01E943, 0x000000, 0x000000],
    },
];

const LOWERCASE: &[Mapping] = &[
    Mapping {
        source: 0x000041,
        len: 1,
        target: [0x000061, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000042,
        len: 1,
        target: [0x000062, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000043,
        len: 1,
        target: [0x000063, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000044,
        len: 1,
        target: [0x000064, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000045,
        len: 1,
        target: [0x000065, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000046,
        len: 1,
        target: [0x000066, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000047,
        len: 1,
        target: [0x000067, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000048,
        len: 1,
        target: [0x000068, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000049,
        len: 1,
        target: [0x000069, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004A,
        len: 1,
        target: [0x00006A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004B,
        len: 1,
        target: [0x00006B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004C,
        len: 1,
        target: [0x00006C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004D,
        len: 1,
        target: [0x00006D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004E,
        len: 1,
        target: [0x00006E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00004F,
        len: 1,
        target: [0x00006F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000050,
        len: 1,
        target: [0x000070, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000051,
        len: 1,
        target: [0x000071, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000052,
        len: 1,
        target: [0x000072, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000053,
        len: 1,
        target: [0x000073, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000054,
        len: 1,
        target: [0x000074, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000055,
        len: 1,
        target: [0x000075, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000056,
        len: 1,
        target: [0x000076, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000057,
        len: 1,
        target: [0x000077, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000058,
        len: 1,
        target: [0x000078, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000059,
        len: 1,
        target: [0x000079, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00005A,
        len: 1,
        target: [0x00007A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C0,
        len: 1,
        target: [0x0000E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C1,
        len: 1,
        target: [0x0000E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C2,
        len: 1,
        target: [0x0000E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C3,
        len: 1,
        target: [0x0000E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C4,
        len: 1,
        target: [0x0000E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C5,
        len: 1,
        target: [0x0000E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C6,
        len: 1,
        target: [0x0000E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C7,
        len: 1,
        target: [0x0000E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C8,
        len: 1,
        target: [0x0000E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000C9,
        len: 1,
        target: [0x0000E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CA,
        len: 1,
        target: [0x0000EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CB,
        len: 1,
        target: [0x0000EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CC,
        len: 1,
        target: [0x0000EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CD,
        len: 1,
        target: [0x0000ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CE,
        len: 1,
        target: [0x0000EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000CF,
        len: 1,
        target: [0x0000EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D0,
        len: 1,
        target: [0x0000F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D1,
        len: 1,
        target: [0x0000F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D2,
        len: 1,
        target: [0x0000F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D3,
        len: 1,
        target: [0x0000F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D4,
        len: 1,
        target: [0x0000F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D5,
        len: 1,
        target: [0x0000F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D6,
        len: 1,
        target: [0x0000F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D8,
        len: 1,
        target: [0x0000F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000D9,
        len: 1,
        target: [0x0000F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DA,
        len: 1,
        target: [0x0000FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DB,
        len: 1,
        target: [0x0000FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DC,
        len: 1,
        target: [0x0000FC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DD,
        len: 1,
        target: [0x0000FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DE,
        len: 1,
        target: [0x0000FE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000100,
        len: 1,
        target: [0x000101, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000102,
        len: 1,
        target: [0x000103, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000104,
        len: 1,
        target: [0x000105, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000106,
        len: 1,
        target: [0x000107, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000108,
        len: 1,
        target: [0x000109, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010A,
        len: 1,
        target: [0x00010B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010C,
        len: 1,
        target: [0x00010D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010E,
        len: 1,
        target: [0x00010F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000110,
        len: 1,
        target: [0x000111, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000112,
        len: 1,
        target: [0x000113, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000114,
        len: 1,
        target: [0x000115, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000116,
        len: 1,
        target: [0x000117, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000118,
        len: 1,
        target: [0x000119, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011A,
        len: 1,
        target: [0x00011B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011C,
        len: 1,
        target: [0x00011D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011E,
        len: 1,
        target: [0x00011F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000120,
        len: 1,
        target: [0x000121, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000122,
        len: 1,
        target: [0x000123, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000124,
        len: 1,
        target: [0x000125, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000126,
        len: 1,
        target: [0x000127, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000128,
        len: 1,
        target: [0x000129, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012A,
        len: 1,
        target: [0x00012B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012C,
        len: 1,
        target: [0x00012D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012E,
        len: 1,
        target: [0x00012F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000130,
        len: 2,
        target: [0x000069, 0x000307, 0x000000],
    },
    Mapping {
        source: 0x000132,
        len: 1,
        target: [0x000133, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000134,
        len: 1,
        target: [0x000135, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000136,
        len: 1,
        target: [0x000137, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000139,
        len: 1,
        target: [0x00013A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013B,
        len: 1,
        target: [0x00013C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013D,
        len: 1,
        target: [0x00013E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013F,
        len: 1,
        target: [0x000140, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000141,
        len: 1,
        target: [0x000142, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000143,
        len: 1,
        target: [0x000144, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000145,
        len: 1,
        target: [0x000146, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000147,
        len: 1,
        target: [0x000148, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00014A,
        len: 1,
        target: [0x00014B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00014C,
        len: 1,
        target: [0x00014D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00014E,
        len: 1,
        target: [0x00014F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000150,
        len: 1,
        target: [0x000151, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000152,
        len: 1,
        target: [0x000153, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000154,
        len: 1,
        target: [0x000155, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000156,
        len: 1,
        target: [0x000157, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000158,
        len: 1,
        target: [0x000159, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015A,
        len: 1,
        target: [0x00015B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015C,
        len: 1,
        target: [0x00015D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015E,
        len: 1,
        target: [0x00015F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000160,
        len: 1,
        target: [0x000161, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000162,
        len: 1,
        target: [0x000163, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000164,
        len: 1,
        target: [0x000165, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000166,
        len: 1,
        target: [0x000167, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000168,
        len: 1,
        target: [0x000169, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016A,
        len: 1,
        target: [0x00016B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016C,
        len: 1,
        target: [0x00016D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016E,
        len: 1,
        target: [0x00016F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000170,
        len: 1,
        target: [0x000171, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000172,
        len: 1,
        target: [0x000173, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000174,
        len: 1,
        target: [0x000175, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000176,
        len: 1,
        target: [0x000177, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000178,
        len: 1,
        target: [0x0000FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000179,
        len: 1,
        target: [0x00017A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017B,
        len: 1,
        target: [0x00017C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017D,
        len: 1,
        target: [0x00017E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000181,
        len: 1,
        target: [0x000253, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000182,
        len: 1,
        target: [0x000183, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000184,
        len: 1,
        target: [0x000185, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000186,
        len: 1,
        target: [0x000254, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000187,
        len: 1,
        target: [0x000188, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000189,
        len: 1,
        target: [0x000256, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018A,
        len: 1,
        target: [0x000257, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018B,
        len: 1,
        target: [0x00018C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018E,
        len: 1,
        target: [0x0001DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018F,
        len: 1,
        target: [0x000259, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000190,
        len: 1,
        target: [0x00025B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000191,
        len: 1,
        target: [0x000192, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000193,
        len: 1,
        target: [0x000260, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000194,
        len: 1,
        target: [0x000263, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000196,
        len: 1,
        target: [0x000269, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000197,
        len: 1,
        target: [0x000268, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000198,
        len: 1,
        target: [0x000199, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019C,
        len: 1,
        target: [0x00026F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019D,
        len: 1,
        target: [0x000272, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019F,
        len: 1,
        target: [0x000275, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A0,
        len: 1,
        target: [0x0001A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A2,
        len: 1,
        target: [0x0001A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A4,
        len: 1,
        target: [0x0001A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A6,
        len: 1,
        target: [0x000280, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A7,
        len: 1,
        target: [0x0001A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A9,
        len: 1,
        target: [0x000283, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001AC,
        len: 1,
        target: [0x0001AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001AE,
        len: 1,
        target: [0x000288, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001AF,
        len: 1,
        target: [0x0001B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B1,
        len: 1,
        target: [0x00028A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B2,
        len: 1,
        target: [0x00028B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B3,
        len: 1,
        target: [0x0001B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B5,
        len: 1,
        target: [0x0001B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B7,
        len: 1,
        target: [0x000292, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B8,
        len: 1,
        target: [0x0001B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001BC,
        len: 1,
        target: [0x0001BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C4,
        len: 1,
        target: [0x0001C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C5,
        len: 1,
        target: [0x0001C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C7,
        len: 1,
        target: [0x0001C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C8,
        len: 1,
        target: [0x0001C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CA,
        len: 1,
        target: [0x0001CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CB,
        len: 1,
        target: [0x0001CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CD,
        len: 1,
        target: [0x0001CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CF,
        len: 1,
        target: [0x0001D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D1,
        len: 1,
        target: [0x0001D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D3,
        len: 1,
        target: [0x0001D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D5,
        len: 1,
        target: [0x0001D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D7,
        len: 1,
        target: [0x0001D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D9,
        len: 1,
        target: [0x0001DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DB,
        len: 1,
        target: [0x0001DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DE,
        len: 1,
        target: [0x0001DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E0,
        len: 1,
        target: [0x0001E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E2,
        len: 1,
        target: [0x0001E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E4,
        len: 1,
        target: [0x0001E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E6,
        len: 1,
        target: [0x0001E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E8,
        len: 1,
        target: [0x0001E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EA,
        len: 1,
        target: [0x0001EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EC,
        len: 1,
        target: [0x0001ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EE,
        len: 1,
        target: [0x0001EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F1,
        len: 1,
        target: [0x0001F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F2,
        len: 1,
        target: [0x0001F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F4,
        len: 1,
        target: [0x0001F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F6,
        len: 1,
        target: [0x000195, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F7,
        len: 1,
        target: [0x0001BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F8,
        len: 1,
        target: [0x0001F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FA,
        len: 1,
        target: [0x0001FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FC,
        len: 1,
        target: [0x0001FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FE,
        len: 1,
        target: [0x0001FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000200,
        len: 1,
        target: [0x000201, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000202,
        len: 1,
        target: [0x000203, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000204,
        len: 1,
        target: [0x000205, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000206,
        len: 1,
        target: [0x000207, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000208,
        len: 1,
        target: [0x000209, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020A,
        len: 1,
        target: [0x00020B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020C,
        len: 1,
        target: [0x00020D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020E,
        len: 1,
        target: [0x00020F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000210,
        len: 1,
        target: [0x000211, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000212,
        len: 1,
        target: [0x000213, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000214,
        len: 1,
        target: [0x000215, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000216,
        len: 1,
        target: [0x000217, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000218,
        len: 1,
        target: [0x000219, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021A,
        len: 1,
        target: [0x00021B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021C,
        len: 1,
        target: [0x00021D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021E,
        len: 1,
        target: [0x00021F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000220,
        len: 1,
        target: [0x00019E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000222,
        len: 1,
        target: [0x000223, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000224,
        len: 1,
        target: [0x000225, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000226,
        len: 1,
        target: [0x000227, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000228,
        len: 1,
        target: [0x000229, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022A,
        len: 1,
        target: [0x00022B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022C,
        len: 1,
        target: [0x00022D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022E,
        len: 1,
        target: [0x00022F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000230,
        len: 1,
        target: [0x000231, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000232,
        len: 1,
        target: [0x000233, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023A,
        len: 1,
        target: [0x002C65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023B,
        len: 1,
        target: [0x00023C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023D,
        len: 1,
        target: [0x00019A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023E,
        len: 1,
        target: [0x002C66, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000241,
        len: 1,
        target: [0x000242, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000243,
        len: 1,
        target: [0x000180, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000244,
        len: 1,
        target: [0x000289, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000245,
        len: 1,
        target: [0x00028C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000246,
        len: 1,
        target: [0x000247, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000248,
        len: 1,
        target: [0x000249, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024A,
        len: 1,
        target: [0x00024B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024C,
        len: 1,
        target: [0x00024D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024E,
        len: 1,
        target: [0x00024F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000370,
        len: 1,
        target: [0x000371, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000372,
        len: 1,
        target: [0x000373, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000376,
        len: 1,
        target: [0x000377, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00037F,
        len: 1,
        target: [0x0003F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000386,
        len: 1,
        target: [0x0003AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000388,
        len: 1,
        target: [0x0003AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000389,
        len: 1,
        target: [0x0003AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038A,
        len: 1,
        target: [0x0003AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038C,
        len: 1,
        target: [0x0003CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038E,
        len: 1,
        target: [0x0003CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00038F,
        len: 1,
        target: [0x0003CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000391,
        len: 1,
        target: [0x0003B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000392,
        len: 1,
        target: [0x0003B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000393,
        len: 1,
        target: [0x0003B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000394,
        len: 1,
        target: [0x0003B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000395,
        len: 1,
        target: [0x0003B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000396,
        len: 1,
        target: [0x0003B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000397,
        len: 1,
        target: [0x0003B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000398,
        len: 1,
        target: [0x0003B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000399,
        len: 1,
        target: [0x0003B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039A,
        len: 1,
        target: [0x0003BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039B,
        len: 1,
        target: [0x0003BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039C,
        len: 1,
        target: [0x0003BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039D,
        len: 1,
        target: [0x0003BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039E,
        len: 1,
        target: [0x0003BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00039F,
        len: 1,
        target: [0x0003BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A0,
        len: 1,
        target: [0x0003C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A1,
        len: 1,
        target: [0x0003C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A3,
        len: 1,
        target: [0x0003C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A4,
        len: 1,
        target: [0x0003C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A5,
        len: 1,
        target: [0x0003C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A6,
        len: 1,
        target: [0x0003C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A7,
        len: 1,
        target: [0x0003C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A8,
        len: 1,
        target: [0x0003C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003A9,
        len: 1,
        target: [0x0003C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003AA,
        len: 1,
        target: [0x0003CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003AB,
        len: 1,
        target: [0x0003CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003CF,
        len: 1,
        target: [0x0003D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D8,
        len: 1,
        target: [0x0003D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DA,
        len: 1,
        target: [0x0003DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DC,
        len: 1,
        target: [0x0003DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DE,
        len: 1,
        target: [0x0003DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E0,
        len: 1,
        target: [0x0003E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E2,
        len: 1,
        target: [0x0003E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E4,
        len: 1,
        target: [0x0003E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E6,
        len: 1,
        target: [0x0003E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E8,
        len: 1,
        target: [0x0003E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EA,
        len: 1,
        target: [0x0003EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EC,
        len: 1,
        target: [0x0003ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EE,
        len: 1,
        target: [0x0003EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F4,
        len: 1,
        target: [0x0003B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F7,
        len: 1,
        target: [0x0003F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F9,
        len: 1,
        target: [0x0003F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FA,
        len: 1,
        target: [0x0003FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FD,
        len: 1,
        target: [0x00037B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FE,
        len: 1,
        target: [0x00037C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FF,
        len: 1,
        target: [0x00037D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000400,
        len: 1,
        target: [0x000450, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000401,
        len: 1,
        target: [0x000451, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000402,
        len: 1,
        target: [0x000452, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000403,
        len: 1,
        target: [0x000453, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000404,
        len: 1,
        target: [0x000454, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000405,
        len: 1,
        target: [0x000455, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000406,
        len: 1,
        target: [0x000456, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000407,
        len: 1,
        target: [0x000457, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000408,
        len: 1,
        target: [0x000458, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000409,
        len: 1,
        target: [0x000459, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040A,
        len: 1,
        target: [0x00045A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040B,
        len: 1,
        target: [0x00045B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040C,
        len: 1,
        target: [0x00045C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040D,
        len: 1,
        target: [0x00045D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040E,
        len: 1,
        target: [0x00045E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00040F,
        len: 1,
        target: [0x00045F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000410,
        len: 1,
        target: [0x000430, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000411,
        len: 1,
        target: [0x000431, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000412,
        len: 1,
        target: [0x000432, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000413,
        len: 1,
        target: [0x000433, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000414,
        len: 1,
        target: [0x000434, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000415,
        len: 1,
        target: [0x000435, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000416,
        len: 1,
        target: [0x000436, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000417,
        len: 1,
        target: [0x000437, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000418,
        len: 1,
        target: [0x000438, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000419,
        len: 1,
        target: [0x000439, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041A,
        len: 1,
        target: [0x00043A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041B,
        len: 1,
        target: [0x00043B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041C,
        len: 1,
        target: [0x00043C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041D,
        len: 1,
        target: [0x00043D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041E,
        len: 1,
        target: [0x00043E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00041F,
        len: 1,
        target: [0x00043F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000420,
        len: 1,
        target: [0x000440, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000421,
        len: 1,
        target: [0x000441, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000422,
        len: 1,
        target: [0x000442, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000423,
        len: 1,
        target: [0x000443, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000424,
        len: 1,
        target: [0x000444, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000425,
        len: 1,
        target: [0x000445, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000426,
        len: 1,
        target: [0x000446, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000427,
        len: 1,
        target: [0x000447, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000428,
        len: 1,
        target: [0x000448, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000429,
        len: 1,
        target: [0x000449, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042A,
        len: 1,
        target: [0x00044A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042B,
        len: 1,
        target: [0x00044B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042C,
        len: 1,
        target: [0x00044C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042D,
        len: 1,
        target: [0x00044D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042E,
        len: 1,
        target: [0x00044E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00042F,
        len: 1,
        target: [0x00044F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000460,
        len: 1,
        target: [0x000461, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000462,
        len: 1,
        target: [0x000463, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000464,
        len: 1,
        target: [0x000465, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000466,
        len: 1,
        target: [0x000467, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000468,
        len: 1,
        target: [0x000469, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046A,
        len: 1,
        target: [0x00046B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046C,
        len: 1,
        target: [0x00046D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046E,
        len: 1,
        target: [0x00046F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000470,
        len: 1,
        target: [0x000471, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000472,
        len: 1,
        target: [0x000473, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000474,
        len: 1,
        target: [0x000475, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000476,
        len: 1,
        target: [0x000477, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000478,
        len: 1,
        target: [0x000479, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047A,
        len: 1,
        target: [0x00047B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047C,
        len: 1,
        target: [0x00047D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047E,
        len: 1,
        target: [0x00047F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000480,
        len: 1,
        target: [0x000481, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048A,
        len: 1,
        target: [0x00048B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048C,
        len: 1,
        target: [0x00048D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048E,
        len: 1,
        target: [0x00048F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000490,
        len: 1,
        target: [0x000491, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000492,
        len: 1,
        target: [0x000493, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000494,
        len: 1,
        target: [0x000495, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000496,
        len: 1,
        target: [0x000497, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000498,
        len: 1,
        target: [0x000499, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049A,
        len: 1,
        target: [0x00049B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049C,
        len: 1,
        target: [0x00049D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049E,
        len: 1,
        target: [0x00049F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A0,
        len: 1,
        target: [0x0004A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A2,
        len: 1,
        target: [0x0004A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A4,
        len: 1,
        target: [0x0004A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A6,
        len: 1,
        target: [0x0004A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A8,
        len: 1,
        target: [0x0004A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AA,
        len: 1,
        target: [0x0004AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AC,
        len: 1,
        target: [0x0004AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AE,
        len: 1,
        target: [0x0004AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B0,
        len: 1,
        target: [0x0004B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B2,
        len: 1,
        target: [0x0004B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B4,
        len: 1,
        target: [0x0004B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B6,
        len: 1,
        target: [0x0004B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B8,
        len: 1,
        target: [0x0004B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BA,
        len: 1,
        target: [0x0004BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BC,
        len: 1,
        target: [0x0004BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BE,
        len: 1,
        target: [0x0004BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C0,
        len: 1,
        target: [0x0004CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C1,
        len: 1,
        target: [0x0004C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C3,
        len: 1,
        target: [0x0004C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C5,
        len: 1,
        target: [0x0004C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C7,
        len: 1,
        target: [0x0004C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C9,
        len: 1,
        target: [0x0004CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CB,
        len: 1,
        target: [0x0004CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CD,
        len: 1,
        target: [0x0004CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D0,
        len: 1,
        target: [0x0004D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D2,
        len: 1,
        target: [0x0004D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D4,
        len: 1,
        target: [0x0004D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D6,
        len: 1,
        target: [0x0004D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D8,
        len: 1,
        target: [0x0004D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DA,
        len: 1,
        target: [0x0004DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DC,
        len: 1,
        target: [0x0004DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DE,
        len: 1,
        target: [0x0004DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E0,
        len: 1,
        target: [0x0004E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E2,
        len: 1,
        target: [0x0004E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E4,
        len: 1,
        target: [0x0004E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E6,
        len: 1,
        target: [0x0004E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E8,
        len: 1,
        target: [0x0004E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EA,
        len: 1,
        target: [0x0004EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EC,
        len: 1,
        target: [0x0004ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EE,
        len: 1,
        target: [0x0004EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F0,
        len: 1,
        target: [0x0004F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F2,
        len: 1,
        target: [0x0004F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F4,
        len: 1,
        target: [0x0004F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F6,
        len: 1,
        target: [0x0004F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F8,
        len: 1,
        target: [0x0004F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FA,
        len: 1,
        target: [0x0004FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FC,
        len: 1,
        target: [0x0004FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FE,
        len: 1,
        target: [0x0004FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000500,
        len: 1,
        target: [0x000501, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000502,
        len: 1,
        target: [0x000503, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000504,
        len: 1,
        target: [0x000505, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000506,
        len: 1,
        target: [0x000507, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000508,
        len: 1,
        target: [0x000509, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050A,
        len: 1,
        target: [0x00050B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050C,
        len: 1,
        target: [0x00050D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050E,
        len: 1,
        target: [0x00050F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000510,
        len: 1,
        target: [0x000511, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000512,
        len: 1,
        target: [0x000513, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000514,
        len: 1,
        target: [0x000515, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000516,
        len: 1,
        target: [0x000517, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000518,
        len: 1,
        target: [0x000519, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051A,
        len: 1,
        target: [0x00051B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051C,
        len: 1,
        target: [0x00051D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051E,
        len: 1,
        target: [0x00051F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000520,
        len: 1,
        target: [0x000521, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000522,
        len: 1,
        target: [0x000523, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000524,
        len: 1,
        target: [0x000525, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000526,
        len: 1,
        target: [0x000527, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000528,
        len: 1,
        target: [0x000529, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052A,
        len: 1,
        target: [0x00052B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052C,
        len: 1,
        target: [0x00052D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052E,
        len: 1,
        target: [0x00052F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000531,
        len: 1,
        target: [0x000561, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000532,
        len: 1,
        target: [0x000562, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000533,
        len: 1,
        target: [0x000563, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000534,
        len: 1,
        target: [0x000564, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000535,
        len: 1,
        target: [0x000565, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000536,
        len: 1,
        target: [0x000566, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000537,
        len: 1,
        target: [0x000567, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000538,
        len: 1,
        target: [0x000568, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000539,
        len: 1,
        target: [0x000569, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053A,
        len: 1,
        target: [0x00056A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053B,
        len: 1,
        target: [0x00056B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053C,
        len: 1,
        target: [0x00056C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053D,
        len: 1,
        target: [0x00056D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053E,
        len: 1,
        target: [0x00056E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00053F,
        len: 1,
        target: [0x00056F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000540,
        len: 1,
        target: [0x000570, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000541,
        len: 1,
        target: [0x000571, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000542,
        len: 1,
        target: [0x000572, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000543,
        len: 1,
        target: [0x000573, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000544,
        len: 1,
        target: [0x000574, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000545,
        len: 1,
        target: [0x000575, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000546,
        len: 1,
        target: [0x000576, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000547,
        len: 1,
        target: [0x000577, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000548,
        len: 1,
        target: [0x000578, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000549,
        len: 1,
        target: [0x000579, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054A,
        len: 1,
        target: [0x00057A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054B,
        len: 1,
        target: [0x00057B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054C,
        len: 1,
        target: [0x00057C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054D,
        len: 1,
        target: [0x00057D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054E,
        len: 1,
        target: [0x00057E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00054F,
        len: 1,
        target: [0x00057F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000550,
        len: 1,
        target: [0x000580, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000551,
        len: 1,
        target: [0x000581, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000552,
        len: 1,
        target: [0x000582, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000553,
        len: 1,
        target: [0x000583, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000554,
        len: 1,
        target: [0x000584, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000555,
        len: 1,
        target: [0x000585, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000556,
        len: 1,
        target: [0x000586, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A0,
        len: 1,
        target: [0x002D00, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A1,
        len: 1,
        target: [0x002D01, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A2,
        len: 1,
        target: [0x002D02, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A3,
        len: 1,
        target: [0x002D03, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A4,
        len: 1,
        target: [0x002D04, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A5,
        len: 1,
        target: [0x002D05, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A6,
        len: 1,
        target: [0x002D06, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A7,
        len: 1,
        target: [0x002D07, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A8,
        len: 1,
        target: [0x002D08, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010A9,
        len: 1,
        target: [0x002D09, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AA,
        len: 1,
        target: [0x002D0A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AB,
        len: 1,
        target: [0x002D0B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AC,
        len: 1,
        target: [0x002D0C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AD,
        len: 1,
        target: [0x002D0D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AE,
        len: 1,
        target: [0x002D0E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010AF,
        len: 1,
        target: [0x002D0F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B0,
        len: 1,
        target: [0x002D10, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B1,
        len: 1,
        target: [0x002D11, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B2,
        len: 1,
        target: [0x002D12, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B3,
        len: 1,
        target: [0x002D13, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B4,
        len: 1,
        target: [0x002D14, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B5,
        len: 1,
        target: [0x002D15, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B6,
        len: 1,
        target: [0x002D16, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B7,
        len: 1,
        target: [0x002D17, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B8,
        len: 1,
        target: [0x002D18, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010B9,
        len: 1,
        target: [0x002D19, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BA,
        len: 1,
        target: [0x002D1A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BB,
        len: 1,
        target: [0x002D1B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BC,
        len: 1,
        target: [0x002D1C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BD,
        len: 1,
        target: [0x002D1D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BE,
        len: 1,
        target: [0x002D1E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010BF,
        len: 1,
        target: [0x002D1F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C0,
        len: 1,
        target: [0x002D20, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C1,
        len: 1,
        target: [0x002D21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C2,
        len: 1,
        target: [0x002D22, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C3,
        len: 1,
        target: [0x002D23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C4,
        len: 1,
        target: [0x002D24, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C5,
        len: 1,
        target: [0x002D25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010C7,
        len: 1,
        target: [0x002D27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010CD,
        len: 1,
        target: [0x002D2D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A0,
        len: 1,
        target: [0x00AB70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A1,
        len: 1,
        target: [0x00AB71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A2,
        len: 1,
        target: [0x00AB72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A3,
        len: 1,
        target: [0x00AB73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A4,
        len: 1,
        target: [0x00AB74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A5,
        len: 1,
        target: [0x00AB75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A6,
        len: 1,
        target: [0x00AB76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A7,
        len: 1,
        target: [0x00AB77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A8,
        len: 1,
        target: [0x00AB78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013A9,
        len: 1,
        target: [0x00AB79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013AA,
        len: 1,
        target: [0x00AB7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013AB,
        len: 1,
        target: [0x00AB7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013AC,
        len: 1,
        target: [0x00AB7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013AD,
        len: 1,
        target: [0x00AB7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013AE,
        len: 1,
        target: [0x00AB7E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013AF,
        len: 1,
        target: [0x00AB7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B0,
        len: 1,
        target: [0x00AB80, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B1,
        len: 1,
        target: [0x00AB81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B2,
        len: 1,
        target: [0x00AB82, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B3,
        len: 1,
        target: [0x00AB83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B4,
        len: 1,
        target: [0x00AB84, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B5,
        len: 1,
        target: [0x00AB85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B6,
        len: 1,
        target: [0x00AB86, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B7,
        len: 1,
        target: [0x00AB87, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B8,
        len: 1,
        target: [0x00AB88, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013B9,
        len: 1,
        target: [0x00AB89, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013BA,
        len: 1,
        target: [0x00AB8A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013BB,
        len: 1,
        target: [0x00AB8B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013BC,
        len: 1,
        target: [0x00AB8C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013BD,
        len: 1,
        target: [0x00AB8D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013BE,
        len: 1,
        target: [0x00AB8E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013BF,
        len: 1,
        target: [0x00AB8F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C0,
        len: 1,
        target: [0x00AB90, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C1,
        len: 1,
        target: [0x00AB91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C2,
        len: 1,
        target: [0x00AB92, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C3,
        len: 1,
        target: [0x00AB93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C4,
        len: 1,
        target: [0x00AB94, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C5,
        len: 1,
        target: [0x00AB95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C6,
        len: 1,
        target: [0x00AB96, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C7,
        len: 1,
        target: [0x00AB97, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C8,
        len: 1,
        target: [0x00AB98, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013C9,
        len: 1,
        target: [0x00AB99, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013CA,
        len: 1,
        target: [0x00AB9A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013CB,
        len: 1,
        target: [0x00AB9B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013CC,
        len: 1,
        target: [0x00AB9C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013CD,
        len: 1,
        target: [0x00AB9D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013CE,
        len: 1,
        target: [0x00AB9E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013CF,
        len: 1,
        target: [0x00AB9F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D0,
        len: 1,
        target: [0x00ABA0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D1,
        len: 1,
        target: [0x00ABA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D2,
        len: 1,
        target: [0x00ABA2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D3,
        len: 1,
        target: [0x00ABA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D4,
        len: 1,
        target: [0x00ABA4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D5,
        len: 1,
        target: [0x00ABA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D6,
        len: 1,
        target: [0x00ABA6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D7,
        len: 1,
        target: [0x00ABA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D8,
        len: 1,
        target: [0x00ABA8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013D9,
        len: 1,
        target: [0x00ABA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013DA,
        len: 1,
        target: [0x00ABAA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013DB,
        len: 1,
        target: [0x00ABAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013DC,
        len: 1,
        target: [0x00ABAC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013DD,
        len: 1,
        target: [0x00ABAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013DE,
        len: 1,
        target: [0x00ABAE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013DF,
        len: 1,
        target: [0x00ABAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E0,
        len: 1,
        target: [0x00ABB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E1,
        len: 1,
        target: [0x00ABB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E2,
        len: 1,
        target: [0x00ABB2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E3,
        len: 1,
        target: [0x00ABB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E4,
        len: 1,
        target: [0x00ABB4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E5,
        len: 1,
        target: [0x00ABB5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E6,
        len: 1,
        target: [0x00ABB6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E7,
        len: 1,
        target: [0x00ABB7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E8,
        len: 1,
        target: [0x00ABB8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013E9,
        len: 1,
        target: [0x00ABB9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013EA,
        len: 1,
        target: [0x00ABBA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013EB,
        len: 1,
        target: [0x00ABBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013EC,
        len: 1,
        target: [0x00ABBC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013ED,
        len: 1,
        target: [0x00ABBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013EE,
        len: 1,
        target: [0x00ABBE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013EF,
        len: 1,
        target: [0x00ABBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F0,
        len: 1,
        target: [0x0013F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F1,
        len: 1,
        target: [0x0013F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F2,
        len: 1,
        target: [0x0013FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F3,
        len: 1,
        target: [0x0013FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F4,
        len: 1,
        target: [0x0013FC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F5,
        len: 1,
        target: [0x0013FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C89,
        len: 1,
        target: [0x001C8A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C90,
        len: 1,
        target: [0x0010D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C91,
        len: 1,
        target: [0x0010D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C92,
        len: 1,
        target: [0x0010D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C93,
        len: 1,
        target: [0x0010D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C94,
        len: 1,
        target: [0x0010D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C95,
        len: 1,
        target: [0x0010D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C96,
        len: 1,
        target: [0x0010D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C97,
        len: 1,
        target: [0x0010D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C98,
        len: 1,
        target: [0x0010D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C99,
        len: 1,
        target: [0x0010D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9A,
        len: 1,
        target: [0x0010DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9B,
        len: 1,
        target: [0x0010DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9C,
        len: 1,
        target: [0x0010DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9D,
        len: 1,
        target: [0x0010DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9E,
        len: 1,
        target: [0x0010DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C9F,
        len: 1,
        target: [0x0010DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA0,
        len: 1,
        target: [0x0010E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA1,
        len: 1,
        target: [0x0010E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA2,
        len: 1,
        target: [0x0010E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA3,
        len: 1,
        target: [0x0010E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA4,
        len: 1,
        target: [0x0010E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA5,
        len: 1,
        target: [0x0010E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA6,
        len: 1,
        target: [0x0010E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA7,
        len: 1,
        target: [0x0010E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA8,
        len: 1,
        target: [0x0010E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CA9,
        len: 1,
        target: [0x0010E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAA,
        len: 1,
        target: [0x0010EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAB,
        len: 1,
        target: [0x0010EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAC,
        len: 1,
        target: [0x0010EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAD,
        len: 1,
        target: [0x0010ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAE,
        len: 1,
        target: [0x0010EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CAF,
        len: 1,
        target: [0x0010EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB0,
        len: 1,
        target: [0x0010F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB1,
        len: 1,
        target: [0x0010F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB2,
        len: 1,
        target: [0x0010F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB3,
        len: 1,
        target: [0x0010F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB4,
        len: 1,
        target: [0x0010F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB5,
        len: 1,
        target: [0x0010F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB6,
        len: 1,
        target: [0x0010F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB7,
        len: 1,
        target: [0x0010F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB8,
        len: 1,
        target: [0x0010F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CB9,
        len: 1,
        target: [0x0010F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBA,
        len: 1,
        target: [0x0010FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBD,
        len: 1,
        target: [0x0010FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBE,
        len: 1,
        target: [0x0010FE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001CBF,
        len: 1,
        target: [0x0010FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E00,
        len: 1,
        target: [0x001E01, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E02,
        len: 1,
        target: [0x001E03, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E04,
        len: 1,
        target: [0x001E05, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E06,
        len: 1,
        target: [0x001E07, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E08,
        len: 1,
        target: [0x001E09, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0A,
        len: 1,
        target: [0x001E0B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0C,
        len: 1,
        target: [0x001E0D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0E,
        len: 1,
        target: [0x001E0F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E10,
        len: 1,
        target: [0x001E11, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E12,
        len: 1,
        target: [0x001E13, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E14,
        len: 1,
        target: [0x001E15, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E16,
        len: 1,
        target: [0x001E17, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E18,
        len: 1,
        target: [0x001E19, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1A,
        len: 1,
        target: [0x001E1B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1C,
        len: 1,
        target: [0x001E1D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1E,
        len: 1,
        target: [0x001E1F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E20,
        len: 1,
        target: [0x001E21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E22,
        len: 1,
        target: [0x001E23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E24,
        len: 1,
        target: [0x001E25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E26,
        len: 1,
        target: [0x001E27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E28,
        len: 1,
        target: [0x001E29, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2A,
        len: 1,
        target: [0x001E2B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2C,
        len: 1,
        target: [0x001E2D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2E,
        len: 1,
        target: [0x001E2F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E30,
        len: 1,
        target: [0x001E31, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E32,
        len: 1,
        target: [0x001E33, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E34,
        len: 1,
        target: [0x001E35, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E36,
        len: 1,
        target: [0x001E37, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E38,
        len: 1,
        target: [0x001E39, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3A,
        len: 1,
        target: [0x001E3B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3C,
        len: 1,
        target: [0x001E3D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3E,
        len: 1,
        target: [0x001E3F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E40,
        len: 1,
        target: [0x001E41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E42,
        len: 1,
        target: [0x001E43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E44,
        len: 1,
        target: [0x001E45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E46,
        len: 1,
        target: [0x001E47, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E48,
        len: 1,
        target: [0x001E49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4A,
        len: 1,
        target: [0x001E4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4C,
        len: 1,
        target: [0x001E4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4E,
        len: 1,
        target: [0x001E4F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E50,
        len: 1,
        target: [0x001E51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E52,
        len: 1,
        target: [0x001E53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E54,
        len: 1,
        target: [0x001E55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E56,
        len: 1,
        target: [0x001E57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E58,
        len: 1,
        target: [0x001E59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5A,
        len: 1,
        target: [0x001E5B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5C,
        len: 1,
        target: [0x001E5D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5E,
        len: 1,
        target: [0x001E5F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E60,
        len: 1,
        target: [0x001E61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E62,
        len: 1,
        target: [0x001E63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E64,
        len: 1,
        target: [0x001E65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E66,
        len: 1,
        target: [0x001E67, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E68,
        len: 1,
        target: [0x001E69, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6A,
        len: 1,
        target: [0x001E6B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6C,
        len: 1,
        target: [0x001E6D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6E,
        len: 1,
        target: [0x001E6F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E70,
        len: 1,
        target: [0x001E71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E72,
        len: 1,
        target: [0x001E73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E74,
        len: 1,
        target: [0x001E75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E76,
        len: 1,
        target: [0x001E77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E78,
        len: 1,
        target: [0x001E79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7A,
        len: 1,
        target: [0x001E7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7C,
        len: 1,
        target: [0x001E7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7E,
        len: 1,
        target: [0x001E7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E80,
        len: 1,
        target: [0x001E81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E82,
        len: 1,
        target: [0x001E83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E84,
        len: 1,
        target: [0x001E85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E86,
        len: 1,
        target: [0x001E87, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E88,
        len: 1,
        target: [0x001E89, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8A,
        len: 1,
        target: [0x001E8B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8C,
        len: 1,
        target: [0x001E8D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8E,
        len: 1,
        target: [0x001E8F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E90,
        len: 1,
        target: [0x001E91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E92,
        len: 1,
        target: [0x001E93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E94,
        len: 1,
        target: [0x001E95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E9E,
        len: 1,
        target: [0x0000DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA0,
        len: 1,
        target: [0x001EA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA2,
        len: 1,
        target: [0x001EA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA4,
        len: 1,
        target: [0x001EA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA6,
        len: 1,
        target: [0x001EA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA8,
        len: 1,
        target: [0x001EA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAA,
        len: 1,
        target: [0x001EAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAC,
        len: 1,
        target: [0x001EAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAE,
        len: 1,
        target: [0x001EAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB0,
        len: 1,
        target: [0x001EB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB2,
        len: 1,
        target: [0x001EB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB4,
        len: 1,
        target: [0x001EB5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB6,
        len: 1,
        target: [0x001EB7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB8,
        len: 1,
        target: [0x001EB9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBA,
        len: 1,
        target: [0x001EBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBC,
        len: 1,
        target: [0x001EBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBE,
        len: 1,
        target: [0x001EBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC0,
        len: 1,
        target: [0x001EC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC2,
        len: 1,
        target: [0x001EC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC4,
        len: 1,
        target: [0x001EC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC6,
        len: 1,
        target: [0x001EC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC8,
        len: 1,
        target: [0x001EC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECA,
        len: 1,
        target: [0x001ECB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECC,
        len: 1,
        target: [0x001ECD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECE,
        len: 1,
        target: [0x001ECF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED0,
        len: 1,
        target: [0x001ED1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED2,
        len: 1,
        target: [0x001ED3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED4,
        len: 1,
        target: [0x001ED5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED6,
        len: 1,
        target: [0x001ED7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED8,
        len: 1,
        target: [0x001ED9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDA,
        len: 1,
        target: [0x001EDB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDC,
        len: 1,
        target: [0x001EDD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDE,
        len: 1,
        target: [0x001EDF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE0,
        len: 1,
        target: [0x001EE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE2,
        len: 1,
        target: [0x001EE3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE4,
        len: 1,
        target: [0x001EE5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE6,
        len: 1,
        target: [0x001EE7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE8,
        len: 1,
        target: [0x001EE9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEA,
        len: 1,
        target: [0x001EEB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEC,
        len: 1,
        target: [0x001EED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEE,
        len: 1,
        target: [0x001EEF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF0,
        len: 1,
        target: [0x001EF1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF2,
        len: 1,
        target: [0x001EF3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF4,
        len: 1,
        target: [0x001EF5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF6,
        len: 1,
        target: [0x001EF7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF8,
        len: 1,
        target: [0x001EF9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFA,
        len: 1,
        target: [0x001EFB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFC,
        len: 1,
        target: [0x001EFD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFE,
        len: 1,
        target: [0x001EFF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F08,
        len: 1,
        target: [0x001F00, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F09,
        len: 1,
        target: [0x001F01, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0A,
        len: 1,
        target: [0x001F02, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0B,
        len: 1,
        target: [0x001F03, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0C,
        len: 1,
        target: [0x001F04, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0D,
        len: 1,
        target: [0x001F05, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0E,
        len: 1,
        target: [0x001F06, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F0F,
        len: 1,
        target: [0x001F07, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F18,
        len: 1,
        target: [0x001F10, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F19,
        len: 1,
        target: [0x001F11, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1A,
        len: 1,
        target: [0x001F12, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1B,
        len: 1,
        target: [0x001F13, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1C,
        len: 1,
        target: [0x001F14, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F1D,
        len: 1,
        target: [0x001F15, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F28,
        len: 1,
        target: [0x001F20, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F29,
        len: 1,
        target: [0x001F21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2A,
        len: 1,
        target: [0x001F22, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2B,
        len: 1,
        target: [0x001F23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2C,
        len: 1,
        target: [0x001F24, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2D,
        len: 1,
        target: [0x001F25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2E,
        len: 1,
        target: [0x001F26, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F2F,
        len: 1,
        target: [0x001F27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F38,
        len: 1,
        target: [0x001F30, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F39,
        len: 1,
        target: [0x001F31, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3A,
        len: 1,
        target: [0x001F32, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3B,
        len: 1,
        target: [0x001F33, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3C,
        len: 1,
        target: [0x001F34, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3D,
        len: 1,
        target: [0x001F35, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3E,
        len: 1,
        target: [0x001F36, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F3F,
        len: 1,
        target: [0x001F37, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F48,
        len: 1,
        target: [0x001F40, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F49,
        len: 1,
        target: [0x001F41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4A,
        len: 1,
        target: [0x001F42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4B,
        len: 1,
        target: [0x001F43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4C,
        len: 1,
        target: [0x001F44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F4D,
        len: 1,
        target: [0x001F45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F59,
        len: 1,
        target: [0x001F51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F5B,
        len: 1,
        target: [0x001F53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F5D,
        len: 1,
        target: [0x001F55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F5F,
        len: 1,
        target: [0x001F57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F68,
        len: 1,
        target: [0x001F60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F69,
        len: 1,
        target: [0x001F61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6A,
        len: 1,
        target: [0x001F62, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6B,
        len: 1,
        target: [0x001F63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6C,
        len: 1,
        target: [0x001F64, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6D,
        len: 1,
        target: [0x001F65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6E,
        len: 1,
        target: [0x001F66, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F6F,
        len: 1,
        target: [0x001F67, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F88,
        len: 1,
        target: [0x001F80, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F89,
        len: 1,
        target: [0x001F81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F8A,
        len: 1,
        target: [0x001F82, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F8B,
        len: 1,
        target: [0x001F83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F8C,
        len: 1,
        target: [0x001F84, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F8D,
        len: 1,
        target: [0x001F85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F8E,
        len: 1,
        target: [0x001F86, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F8F,
        len: 1,
        target: [0x001F87, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F98,
        len: 1,
        target: [0x001F90, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F99,
        len: 1,
        target: [0x001F91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F9A,
        len: 1,
        target: [0x001F92, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F9B,
        len: 1,
        target: [0x001F93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F9C,
        len: 1,
        target: [0x001F94, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F9D,
        len: 1,
        target: [0x001F95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F9E,
        len: 1,
        target: [0x001F96, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F9F,
        len: 1,
        target: [0x001F97, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FA8,
        len: 1,
        target: [0x001FA0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FA9,
        len: 1,
        target: [0x001FA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FAA,
        len: 1,
        target: [0x001FA2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FAB,
        len: 1,
        target: [0x001FA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FAC,
        len: 1,
        target: [0x001FA4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FAD,
        len: 1,
        target: [0x001FA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FAE,
        len: 1,
        target: [0x001FA6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FAF,
        len: 1,
        target: [0x001FA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FB8,
        len: 1,
        target: [0x001FB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FB9,
        len: 1,
        target: [0x001FB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FBA,
        len: 1,
        target: [0x001F70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FBB,
        len: 1,
        target: [0x001F71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FBC,
        len: 1,
        target: [0x001FB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FC8,
        len: 1,
        target: [0x001F72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FC9,
        len: 1,
        target: [0x001F73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FCA,
        len: 1,
        target: [0x001F74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FCB,
        len: 1,
        target: [0x001F75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FCC,
        len: 1,
        target: [0x001FC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FD8,
        len: 1,
        target: [0x001FD0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FD9,
        len: 1,
        target: [0x001FD1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FDA,
        len: 1,
        target: [0x001F76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FDB,
        len: 1,
        target: [0x001F77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FE8,
        len: 1,
        target: [0x001FE0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FE9,
        len: 1,
        target: [0x001FE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FEA,
        len: 1,
        target: [0x001F7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FEB,
        len: 1,
        target: [0x001F7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FEC,
        len: 1,
        target: [0x001FE5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FF8,
        len: 1,
        target: [0x001F78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FF9,
        len: 1,
        target: [0x001F79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FFA,
        len: 1,
        target: [0x001F7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FFB,
        len: 1,
        target: [0x001F7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FFC,
        len: 1,
        target: [0x001FF3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002126,
        len: 1,
        target: [0x0003C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00212A,
        len: 1,
        target: [0x00006B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00212B,
        len: 1,
        target: [0x0000E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002132,
        len: 1,
        target: [0x00214E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002160,
        len: 1,
        target: [0x002170, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002161,
        len: 1,
        target: [0x002171, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002162,
        len: 1,
        target: [0x002172, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002163,
        len: 1,
        target: [0x002173, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002164,
        len: 1,
        target: [0x002174, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002165,
        len: 1,
        target: [0x002175, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002166,
        len: 1,
        target: [0x002176, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002167,
        len: 1,
        target: [0x002177, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002168,
        len: 1,
        target: [0x002178, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002169,
        len: 1,
        target: [0x002179, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216A,
        len: 1,
        target: [0x00217A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216B,
        len: 1,
        target: [0x00217B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216C,
        len: 1,
        target: [0x00217C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216D,
        len: 1,
        target: [0x00217D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216E,
        len: 1,
        target: [0x00217E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00216F,
        len: 1,
        target: [0x00217F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002183,
        len: 1,
        target: [0x002184, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B6,
        len: 1,
        target: [0x0024D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B7,
        len: 1,
        target: [0x0024D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B8,
        len: 1,
        target: [0x0024D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024B9,
        len: 1,
        target: [0x0024D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BA,
        len: 1,
        target: [0x0024D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BB,
        len: 1,
        target: [0x0024D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BC,
        len: 1,
        target: [0x0024D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BD,
        len: 1,
        target: [0x0024D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BE,
        len: 1,
        target: [0x0024D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024BF,
        len: 1,
        target: [0x0024D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C0,
        len: 1,
        target: [0x0024DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C1,
        len: 1,
        target: [0x0024DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C2,
        len: 1,
        target: [0x0024DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C3,
        len: 1,
        target: [0x0024DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C4,
        len: 1,
        target: [0x0024DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C5,
        len: 1,
        target: [0x0024DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C6,
        len: 1,
        target: [0x0024E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C7,
        len: 1,
        target: [0x0024E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C8,
        len: 1,
        target: [0x0024E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024C9,
        len: 1,
        target: [0x0024E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CA,
        len: 1,
        target: [0x0024E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CB,
        len: 1,
        target: [0x0024E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CC,
        len: 1,
        target: [0x0024E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CD,
        len: 1,
        target: [0x0024E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CE,
        len: 1,
        target: [0x0024E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024CF,
        len: 1,
        target: [0x0024E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C00,
        len: 1,
        target: [0x002C30, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C01,
        len: 1,
        target: [0x002C31, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C02,
        len: 1,
        target: [0x002C32, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C03,
        len: 1,
        target: [0x002C33, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C04,
        len: 1,
        target: [0x002C34, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C05,
        len: 1,
        target: [0x002C35, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C06,
        len: 1,
        target: [0x002C36, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C07,
        len: 1,
        target: [0x002C37, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C08,
        len: 1,
        target: [0x002C38, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C09,
        len: 1,
        target: [0x002C39, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0A,
        len: 1,
        target: [0x002C3A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0B,
        len: 1,
        target: [0x002C3B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0C,
        len: 1,
        target: [0x002C3C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0D,
        len: 1,
        target: [0x002C3D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0E,
        len: 1,
        target: [0x002C3E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C0F,
        len: 1,
        target: [0x002C3F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C10,
        len: 1,
        target: [0x002C40, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C11,
        len: 1,
        target: [0x002C41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C12,
        len: 1,
        target: [0x002C42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C13,
        len: 1,
        target: [0x002C43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C14,
        len: 1,
        target: [0x002C44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C15,
        len: 1,
        target: [0x002C45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C16,
        len: 1,
        target: [0x002C46, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C17,
        len: 1,
        target: [0x002C47, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C18,
        len: 1,
        target: [0x002C48, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C19,
        len: 1,
        target: [0x002C49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1A,
        len: 1,
        target: [0x002C4A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1B,
        len: 1,
        target: [0x002C4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1C,
        len: 1,
        target: [0x002C4C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1D,
        len: 1,
        target: [0x002C4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1E,
        len: 1,
        target: [0x002C4E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C1F,
        len: 1,
        target: [0x002C4F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C20,
        len: 1,
        target: [0x002C50, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C21,
        len: 1,
        target: [0x002C51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C22,
        len: 1,
        target: [0x002C52, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C23,
        len: 1,
        target: [0x002C53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C24,
        len: 1,
        target: [0x002C54, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C25,
        len: 1,
        target: [0x002C55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C26,
        len: 1,
        target: [0x002C56, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C27,
        len: 1,
        target: [0x002C57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C28,
        len: 1,
        target: [0x002C58, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C29,
        len: 1,
        target: [0x002C59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2A,
        len: 1,
        target: [0x002C5A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2B,
        len: 1,
        target: [0x002C5B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2C,
        len: 1,
        target: [0x002C5C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2D,
        len: 1,
        target: [0x002C5D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2E,
        len: 1,
        target: [0x002C5E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C2F,
        len: 1,
        target: [0x002C5F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C60,
        len: 1,
        target: [0x002C61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C62,
        len: 1,
        target: [0x00026B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C63,
        len: 1,
        target: [0x001D7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C64,
        len: 1,
        target: [0x00027D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C67,
        len: 1,
        target: [0x002C68, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C69,
        len: 1,
        target: [0x002C6A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6B,
        len: 1,
        target: [0x002C6C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6D,
        len: 1,
        target: [0x000251, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6E,
        len: 1,
        target: [0x000271, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6F,
        len: 1,
        target: [0x000250, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C70,
        len: 1,
        target: [0x000252, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C72,
        len: 1,
        target: [0x002C73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C75,
        len: 1,
        target: [0x002C76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C7E,
        len: 1,
        target: [0x00023F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C7F,
        len: 1,
        target: [0x000240, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C80,
        len: 1,
        target: [0x002C81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C82,
        len: 1,
        target: [0x002C83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C84,
        len: 1,
        target: [0x002C85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C86,
        len: 1,
        target: [0x002C87, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C88,
        len: 1,
        target: [0x002C89, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8A,
        len: 1,
        target: [0x002C8B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8C,
        len: 1,
        target: [0x002C8D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8E,
        len: 1,
        target: [0x002C8F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C90,
        len: 1,
        target: [0x002C91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C92,
        len: 1,
        target: [0x002C93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C94,
        len: 1,
        target: [0x002C95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C96,
        len: 1,
        target: [0x002C97, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C98,
        len: 1,
        target: [0x002C99, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9A,
        len: 1,
        target: [0x002C9B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9C,
        len: 1,
        target: [0x002C9D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9E,
        len: 1,
        target: [0x002C9F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA0,
        len: 1,
        target: [0x002CA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA2,
        len: 1,
        target: [0x002CA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA4,
        len: 1,
        target: [0x002CA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA6,
        len: 1,
        target: [0x002CA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA8,
        len: 1,
        target: [0x002CA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAA,
        len: 1,
        target: [0x002CAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAC,
        len: 1,
        target: [0x002CAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAE,
        len: 1,
        target: [0x002CAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB0,
        len: 1,
        target: [0x002CB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB2,
        len: 1,
        target: [0x002CB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB4,
        len: 1,
        target: [0x002CB5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB6,
        len: 1,
        target: [0x002CB7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB8,
        len: 1,
        target: [0x002CB9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBA,
        len: 1,
        target: [0x002CBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBC,
        len: 1,
        target: [0x002CBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBE,
        len: 1,
        target: [0x002CBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC0,
        len: 1,
        target: [0x002CC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC2,
        len: 1,
        target: [0x002CC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC4,
        len: 1,
        target: [0x002CC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC6,
        len: 1,
        target: [0x002CC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC8,
        len: 1,
        target: [0x002CC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCA,
        len: 1,
        target: [0x002CCB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCC,
        len: 1,
        target: [0x002CCD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCE,
        len: 1,
        target: [0x002CCF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD0,
        len: 1,
        target: [0x002CD1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD2,
        len: 1,
        target: [0x002CD3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD4,
        len: 1,
        target: [0x002CD5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD6,
        len: 1,
        target: [0x002CD7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD8,
        len: 1,
        target: [0x002CD9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDA,
        len: 1,
        target: [0x002CDB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDC,
        len: 1,
        target: [0x002CDD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDE,
        len: 1,
        target: [0x002CDF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CE0,
        len: 1,
        target: [0x002CE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CE2,
        len: 1,
        target: [0x002CE3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CEB,
        len: 1,
        target: [0x002CEC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CED,
        len: 1,
        target: [0x002CEE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CF2,
        len: 1,
        target: [0x002CF3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A640,
        len: 1,
        target: [0x00A641, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A642,
        len: 1,
        target: [0x00A643, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A644,
        len: 1,
        target: [0x00A645, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A646,
        len: 1,
        target: [0x00A647, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A648,
        len: 1,
        target: [0x00A649, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64A,
        len: 1,
        target: [0x00A64B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64C,
        len: 1,
        target: [0x00A64D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64E,
        len: 1,
        target: [0x00A64F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A650,
        len: 1,
        target: [0x00A651, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A652,
        len: 1,
        target: [0x00A653, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A654,
        len: 1,
        target: [0x00A655, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A656,
        len: 1,
        target: [0x00A657, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A658,
        len: 1,
        target: [0x00A659, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65A,
        len: 1,
        target: [0x00A65B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65C,
        len: 1,
        target: [0x00A65D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65E,
        len: 1,
        target: [0x00A65F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A660,
        len: 1,
        target: [0x00A661, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A662,
        len: 1,
        target: [0x00A663, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A664,
        len: 1,
        target: [0x00A665, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A666,
        len: 1,
        target: [0x00A667, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A668,
        len: 1,
        target: [0x00A669, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A66A,
        len: 1,
        target: [0x00A66B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A66C,
        len: 1,
        target: [0x00A66D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A680,
        len: 1,
        target: [0x00A681, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A682,
        len: 1,
        target: [0x00A683, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A684,
        len: 1,
        target: [0x00A685, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A686,
        len: 1,
        target: [0x00A687, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A688,
        len: 1,
        target: [0x00A689, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68A,
        len: 1,
        target: [0x00A68B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68C,
        len: 1,
        target: [0x00A68D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68E,
        len: 1,
        target: [0x00A68F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A690,
        len: 1,
        target: [0x00A691, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A692,
        len: 1,
        target: [0x00A693, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A694,
        len: 1,
        target: [0x00A695, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A696,
        len: 1,
        target: [0x00A697, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A698,
        len: 1,
        target: [0x00A699, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A69A,
        len: 1,
        target: [0x00A69B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A722,
        len: 1,
        target: [0x00A723, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A724,
        len: 1,
        target: [0x00A725, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A726,
        len: 1,
        target: [0x00A727, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A728,
        len: 1,
        target: [0x00A729, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72A,
        len: 1,
        target: [0x00A72B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72C,
        len: 1,
        target: [0x00A72D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72E,
        len: 1,
        target: [0x00A72F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A732,
        len: 1,
        target: [0x00A733, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A734,
        len: 1,
        target: [0x00A735, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A736,
        len: 1,
        target: [0x00A737, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A738,
        len: 1,
        target: [0x00A739, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73A,
        len: 1,
        target: [0x00A73B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73C,
        len: 1,
        target: [0x00A73D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73E,
        len: 1,
        target: [0x00A73F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A740,
        len: 1,
        target: [0x00A741, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A742,
        len: 1,
        target: [0x00A743, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A744,
        len: 1,
        target: [0x00A745, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A746,
        len: 1,
        target: [0x00A747, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A748,
        len: 1,
        target: [0x00A749, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74A,
        len: 1,
        target: [0x00A74B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74C,
        len: 1,
        target: [0x00A74D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74E,
        len: 1,
        target: [0x00A74F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A750,
        len: 1,
        target: [0x00A751, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A752,
        len: 1,
        target: [0x00A753, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A754,
        len: 1,
        target: [0x00A755, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A756,
        len: 1,
        target: [0x00A757, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A758,
        len: 1,
        target: [0x00A759, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75A,
        len: 1,
        target: [0x00A75B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75C,
        len: 1,
        target: [0x00A75D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75E,
        len: 1,
        target: [0x00A75F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A760,
        len: 1,
        target: [0x00A761, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A762,
        len: 1,
        target: [0x00A763, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A764,
        len: 1,
        target: [0x00A765, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A766,
        len: 1,
        target: [0x00A767, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A768,
        len: 1,
        target: [0x00A769, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76A,
        len: 1,
        target: [0x00A76B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76C,
        len: 1,
        target: [0x00A76D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76E,
        len: 1,
        target: [0x00A76F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A779,
        len: 1,
        target: [0x00A77A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77B,
        len: 1,
        target: [0x00A77C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77D,
        len: 1,
        target: [0x001D79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77E,
        len: 1,
        target: [0x00A77F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A780,
        len: 1,
        target: [0x00A781, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A782,
        len: 1,
        target: [0x00A783, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A784,
        len: 1,
        target: [0x00A785, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A786,
        len: 1,
        target: [0x00A787, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A78B,
        len: 1,
        target: [0x00A78C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A78D,
        len: 1,
        target: [0x000265, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A790,
        len: 1,
        target: [0x00A791, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A792,
        len: 1,
        target: [0x00A793, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A796,
        len: 1,
        target: [0x00A797, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A798,
        len: 1,
        target: [0x00A799, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79A,
        len: 1,
        target: [0x00A79B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79C,
        len: 1,
        target: [0x00A79D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79E,
        len: 1,
        target: [0x00A79F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A0,
        len: 1,
        target: [0x00A7A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A2,
        len: 1,
        target: [0x00A7A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A4,
        len: 1,
        target: [0x00A7A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A6,
        len: 1,
        target: [0x00A7A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A8,
        len: 1,
        target: [0x00A7A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AA,
        len: 1,
        target: [0x000266, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AB,
        len: 1,
        target: [0x00025C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AC,
        len: 1,
        target: [0x000261, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AD,
        len: 1,
        target: [0x00026C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7AE,
        len: 1,
        target: [0x00026A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B0,
        len: 1,
        target: [0x00029E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B1,
        len: 1,
        target: [0x000287, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B2,
        len: 1,
        target: [0x00029D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B3,
        len: 1,
        target: [0x00AB53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B4,
        len: 1,
        target: [0x00A7B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B6,
        len: 1,
        target: [0x00A7B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B8,
        len: 1,
        target: [0x00A7B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BA,
        len: 1,
        target: [0x00A7BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BC,
        len: 1,
        target: [0x00A7BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BE,
        len: 1,
        target: [0x00A7BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C0,
        len: 1,
        target: [0x00A7C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C2,
        len: 1,
        target: [0x00A7C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C4,
        len: 1,
        target: [0x00A794, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C5,
        len: 1,
        target: [0x000282, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C6,
        len: 1,
        target: [0x001D8E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C7,
        len: 1,
        target: [0x00A7C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C9,
        len: 1,
        target: [0x00A7CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CB,
        len: 1,
        target: [0x000264, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CC,
        len: 1,
        target: [0x00A7CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CE,
        len: 1,
        target: [0x00A7CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D0,
        len: 1,
        target: [0x00A7D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D2,
        len: 1,
        target: [0x00A7D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D4,
        len: 1,
        target: [0x00A7D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D6,
        len: 1,
        target: [0x00A7D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D8,
        len: 1,
        target: [0x00A7D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7DA,
        len: 1,
        target: [0x00A7DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7DC,
        len: 1,
        target: [0x00019B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7F5,
        len: 1,
        target: [0x00A7F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF21,
        len: 1,
        target: [0x00FF41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF22,
        len: 1,
        target: [0x00FF42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF23,
        len: 1,
        target: [0x00FF43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF24,
        len: 1,
        target: [0x00FF44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF25,
        len: 1,
        target: [0x00FF45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF26,
        len: 1,
        target: [0x00FF46, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF27,
        len: 1,
        target: [0x00FF47, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF28,
        len: 1,
        target: [0x00FF48, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF29,
        len: 1,
        target: [0x00FF49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2A,
        len: 1,
        target: [0x00FF4A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2B,
        len: 1,
        target: [0x00FF4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2C,
        len: 1,
        target: [0x00FF4C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2D,
        len: 1,
        target: [0x00FF4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2E,
        len: 1,
        target: [0x00FF4E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF2F,
        len: 1,
        target: [0x00FF4F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF30,
        len: 1,
        target: [0x00FF50, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF31,
        len: 1,
        target: [0x00FF51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF32,
        len: 1,
        target: [0x00FF52, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF33,
        len: 1,
        target: [0x00FF53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF34,
        len: 1,
        target: [0x00FF54, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF35,
        len: 1,
        target: [0x00FF55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF36,
        len: 1,
        target: [0x00FF56, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF37,
        len: 1,
        target: [0x00FF57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF38,
        len: 1,
        target: [0x00FF58, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF39,
        len: 1,
        target: [0x00FF59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF3A,
        len: 1,
        target: [0x00FF5A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010400,
        len: 1,
        target: [0x010428, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010401,
        len: 1,
        target: [0x010429, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010402,
        len: 1,
        target: [0x01042A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010403,
        len: 1,
        target: [0x01042B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010404,
        len: 1,
        target: [0x01042C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010405,
        len: 1,
        target: [0x01042D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010406,
        len: 1,
        target: [0x01042E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010407,
        len: 1,
        target: [0x01042F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010408,
        len: 1,
        target: [0x010430, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010409,
        len: 1,
        target: [0x010431, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040A,
        len: 1,
        target: [0x010432, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040B,
        len: 1,
        target: [0x010433, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040C,
        len: 1,
        target: [0x010434, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040D,
        len: 1,
        target: [0x010435, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040E,
        len: 1,
        target: [0x010436, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01040F,
        len: 1,
        target: [0x010437, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010410,
        len: 1,
        target: [0x010438, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010411,
        len: 1,
        target: [0x010439, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010412,
        len: 1,
        target: [0x01043A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010413,
        len: 1,
        target: [0x01043B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010414,
        len: 1,
        target: [0x01043C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010415,
        len: 1,
        target: [0x01043D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010416,
        len: 1,
        target: [0x01043E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010417,
        len: 1,
        target: [0x01043F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010418,
        len: 1,
        target: [0x010440, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010419,
        len: 1,
        target: [0x010441, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041A,
        len: 1,
        target: [0x010442, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041B,
        len: 1,
        target: [0x010443, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041C,
        len: 1,
        target: [0x010444, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041D,
        len: 1,
        target: [0x010445, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041E,
        len: 1,
        target: [0x010446, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01041F,
        len: 1,
        target: [0x010447, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010420,
        len: 1,
        target: [0x010448, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010421,
        len: 1,
        target: [0x010449, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010422,
        len: 1,
        target: [0x01044A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010423,
        len: 1,
        target: [0x01044B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010424,
        len: 1,
        target: [0x01044C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010425,
        len: 1,
        target: [0x01044D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010426,
        len: 1,
        target: [0x01044E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010427,
        len: 1,
        target: [0x01044F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B0,
        len: 1,
        target: [0x0104D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B1,
        len: 1,
        target: [0x0104D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B2,
        len: 1,
        target: [0x0104DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B3,
        len: 1,
        target: [0x0104DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B4,
        len: 1,
        target: [0x0104DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B5,
        len: 1,
        target: [0x0104DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B6,
        len: 1,
        target: [0x0104DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B7,
        len: 1,
        target: [0x0104DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B8,
        len: 1,
        target: [0x0104E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104B9,
        len: 1,
        target: [0x0104E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BA,
        len: 1,
        target: [0x0104E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BB,
        len: 1,
        target: [0x0104E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BC,
        len: 1,
        target: [0x0104E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BD,
        len: 1,
        target: [0x0104E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BE,
        len: 1,
        target: [0x0104E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104BF,
        len: 1,
        target: [0x0104E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C0,
        len: 1,
        target: [0x0104E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C1,
        len: 1,
        target: [0x0104E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C2,
        len: 1,
        target: [0x0104EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C3,
        len: 1,
        target: [0x0104EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C4,
        len: 1,
        target: [0x0104EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C5,
        len: 1,
        target: [0x0104ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C6,
        len: 1,
        target: [0x0104EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C7,
        len: 1,
        target: [0x0104EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C8,
        len: 1,
        target: [0x0104F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104C9,
        len: 1,
        target: [0x0104F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CA,
        len: 1,
        target: [0x0104F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CB,
        len: 1,
        target: [0x0104F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CC,
        len: 1,
        target: [0x0104F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CD,
        len: 1,
        target: [0x0104F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CE,
        len: 1,
        target: [0x0104F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104CF,
        len: 1,
        target: [0x0104F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D0,
        len: 1,
        target: [0x0104F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D1,
        len: 1,
        target: [0x0104F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D2,
        len: 1,
        target: [0x0104FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D3,
        len: 1,
        target: [0x0104FB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010570,
        len: 1,
        target: [0x010597, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010571,
        len: 1,
        target: [0x010598, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010572,
        len: 1,
        target: [0x010599, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010573,
        len: 1,
        target: [0x01059A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010574,
        len: 1,
        target: [0x01059B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010575,
        len: 1,
        target: [0x01059C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010576,
        len: 1,
        target: [0x01059D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010577,
        len: 1,
        target: [0x01059E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010578,
        len: 1,
        target: [0x01059F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010579,
        len: 1,
        target: [0x0105A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057A,
        len: 1,
        target: [0x0105A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057C,
        len: 1,
        target: [0x0105A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057D,
        len: 1,
        target: [0x0105A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057E,
        len: 1,
        target: [0x0105A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01057F,
        len: 1,
        target: [0x0105A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010580,
        len: 1,
        target: [0x0105A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010581,
        len: 1,
        target: [0x0105A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010582,
        len: 1,
        target: [0x0105A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010583,
        len: 1,
        target: [0x0105AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010584,
        len: 1,
        target: [0x0105AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010585,
        len: 1,
        target: [0x0105AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010586,
        len: 1,
        target: [0x0105AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010587,
        len: 1,
        target: [0x0105AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010588,
        len: 1,
        target: [0x0105AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010589,
        len: 1,
        target: [0x0105B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058A,
        len: 1,
        target: [0x0105B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058C,
        len: 1,
        target: [0x0105B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058D,
        len: 1,
        target: [0x0105B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058E,
        len: 1,
        target: [0x0105B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01058F,
        len: 1,
        target: [0x0105B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010590,
        len: 1,
        target: [0x0105B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010591,
        len: 1,
        target: [0x0105B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010592,
        len: 1,
        target: [0x0105B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010594,
        len: 1,
        target: [0x0105BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010595,
        len: 1,
        target: [0x0105BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C80,
        len: 1,
        target: [0x010CC0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C81,
        len: 1,
        target: [0x010CC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C82,
        len: 1,
        target: [0x010CC2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C83,
        len: 1,
        target: [0x010CC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C84,
        len: 1,
        target: [0x010CC4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C85,
        len: 1,
        target: [0x010CC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C86,
        len: 1,
        target: [0x010CC6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C87,
        len: 1,
        target: [0x010CC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C88,
        len: 1,
        target: [0x010CC8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C89,
        len: 1,
        target: [0x010CC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8A,
        len: 1,
        target: [0x010CCA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8B,
        len: 1,
        target: [0x010CCB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8C,
        len: 1,
        target: [0x010CCC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8D,
        len: 1,
        target: [0x010CCD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8E,
        len: 1,
        target: [0x010CCE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C8F,
        len: 1,
        target: [0x010CCF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C90,
        len: 1,
        target: [0x010CD0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C91,
        len: 1,
        target: [0x010CD1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C92,
        len: 1,
        target: [0x010CD2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C93,
        len: 1,
        target: [0x010CD3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C94,
        len: 1,
        target: [0x010CD4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C95,
        len: 1,
        target: [0x010CD5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C96,
        len: 1,
        target: [0x010CD6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C97,
        len: 1,
        target: [0x010CD7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C98,
        len: 1,
        target: [0x010CD8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C99,
        len: 1,
        target: [0x010CD9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9A,
        len: 1,
        target: [0x010CDA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9B,
        len: 1,
        target: [0x010CDB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9C,
        len: 1,
        target: [0x010CDC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9D,
        len: 1,
        target: [0x010CDD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9E,
        len: 1,
        target: [0x010CDE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010C9F,
        len: 1,
        target: [0x010CDF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA0,
        len: 1,
        target: [0x010CE0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA1,
        len: 1,
        target: [0x010CE1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA2,
        len: 1,
        target: [0x010CE2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA3,
        len: 1,
        target: [0x010CE3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA4,
        len: 1,
        target: [0x010CE4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA5,
        len: 1,
        target: [0x010CE5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA6,
        len: 1,
        target: [0x010CE6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA7,
        len: 1,
        target: [0x010CE7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA8,
        len: 1,
        target: [0x010CE8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CA9,
        len: 1,
        target: [0x010CE9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAA,
        len: 1,
        target: [0x010CEA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAB,
        len: 1,
        target: [0x010CEB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAC,
        len: 1,
        target: [0x010CEC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAD,
        len: 1,
        target: [0x010CED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAE,
        len: 1,
        target: [0x010CEE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CAF,
        len: 1,
        target: [0x010CEF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CB0,
        len: 1,
        target: [0x010CF0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CB1,
        len: 1,
        target: [0x010CF1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CB2,
        len: 1,
        target: [0x010CF2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D50,
        len: 1,
        target: [0x010D70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D51,
        len: 1,
        target: [0x010D71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D52,
        len: 1,
        target: [0x010D72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D53,
        len: 1,
        target: [0x010D73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D54,
        len: 1,
        target: [0x010D74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D55,
        len: 1,
        target: [0x010D75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D56,
        len: 1,
        target: [0x010D76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D57,
        len: 1,
        target: [0x010D77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D58,
        len: 1,
        target: [0x010D78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D59,
        len: 1,
        target: [0x010D79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5A,
        len: 1,
        target: [0x010D7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5B,
        len: 1,
        target: [0x010D7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5C,
        len: 1,
        target: [0x010D7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5D,
        len: 1,
        target: [0x010D7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5E,
        len: 1,
        target: [0x010D7E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D5F,
        len: 1,
        target: [0x010D7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D60,
        len: 1,
        target: [0x010D80, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D61,
        len: 1,
        target: [0x010D81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D62,
        len: 1,
        target: [0x010D82, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D63,
        len: 1,
        target: [0x010D83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D64,
        len: 1,
        target: [0x010D84, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D65,
        len: 1,
        target: [0x010D85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A0,
        len: 1,
        target: [0x0118C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A1,
        len: 1,
        target: [0x0118C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A2,
        len: 1,
        target: [0x0118C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A3,
        len: 1,
        target: [0x0118C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A4,
        len: 1,
        target: [0x0118C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A5,
        len: 1,
        target: [0x0118C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A6,
        len: 1,
        target: [0x0118C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A7,
        len: 1,
        target: [0x0118C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A8,
        len: 1,
        target: [0x0118C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118A9,
        len: 1,
        target: [0x0118C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AA,
        len: 1,
        target: [0x0118CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AB,
        len: 1,
        target: [0x0118CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AC,
        len: 1,
        target: [0x0118CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AD,
        len: 1,
        target: [0x0118CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AE,
        len: 1,
        target: [0x0118CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118AF,
        len: 1,
        target: [0x0118CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B0,
        len: 1,
        target: [0x0118D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B1,
        len: 1,
        target: [0x0118D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B2,
        len: 1,
        target: [0x0118D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B3,
        len: 1,
        target: [0x0118D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B4,
        len: 1,
        target: [0x0118D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B5,
        len: 1,
        target: [0x0118D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B6,
        len: 1,
        target: [0x0118D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B7,
        len: 1,
        target: [0x0118D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B8,
        len: 1,
        target: [0x0118D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118B9,
        len: 1,
        target: [0x0118D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BA,
        len: 1,
        target: [0x0118DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BB,
        len: 1,
        target: [0x0118DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BC,
        len: 1,
        target: [0x0118DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BD,
        len: 1,
        target: [0x0118DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BE,
        len: 1,
        target: [0x0118DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118BF,
        len: 1,
        target: [0x0118DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E40,
        len: 1,
        target: [0x016E60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E41,
        len: 1,
        target: [0x016E61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E42,
        len: 1,
        target: [0x016E62, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E43,
        len: 1,
        target: [0x016E63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E44,
        len: 1,
        target: [0x016E64, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E45,
        len: 1,
        target: [0x016E65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E46,
        len: 1,
        target: [0x016E66, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E47,
        len: 1,
        target: [0x016E67, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E48,
        len: 1,
        target: [0x016E68, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E49,
        len: 1,
        target: [0x016E69, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4A,
        len: 1,
        target: [0x016E6A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4B,
        len: 1,
        target: [0x016E6B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4C,
        len: 1,
        target: [0x016E6C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4D,
        len: 1,
        target: [0x016E6D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4E,
        len: 1,
        target: [0x016E6E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E4F,
        len: 1,
        target: [0x016E6F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E50,
        len: 1,
        target: [0x016E70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E51,
        len: 1,
        target: [0x016E71, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E52,
        len: 1,
        target: [0x016E72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E53,
        len: 1,
        target: [0x016E73, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E54,
        len: 1,
        target: [0x016E74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E55,
        len: 1,
        target: [0x016E75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E56,
        len: 1,
        target: [0x016E76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E57,
        len: 1,
        target: [0x016E77, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E58,
        len: 1,
        target: [0x016E78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E59,
        len: 1,
        target: [0x016E79, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5A,
        len: 1,
        target: [0x016E7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5B,
        len: 1,
        target: [0x016E7B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5C,
        len: 1,
        target: [0x016E7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5D,
        len: 1,
        target: [0x016E7D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5E,
        len: 1,
        target: [0x016E7E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E5F,
        len: 1,
        target: [0x016E7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA0,
        len: 1,
        target: [0x016EBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA1,
        len: 1,
        target: [0x016EBC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA2,
        len: 1,
        target: [0x016EBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA3,
        len: 1,
        target: [0x016EBE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA4,
        len: 1,
        target: [0x016EBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA5,
        len: 1,
        target: [0x016EC0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA6,
        len: 1,
        target: [0x016EC1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA7,
        len: 1,
        target: [0x016EC2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA8,
        len: 1,
        target: [0x016EC3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EA9,
        len: 1,
        target: [0x016EC4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAA,
        len: 1,
        target: [0x016EC5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAB,
        len: 1,
        target: [0x016EC6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAC,
        len: 1,
        target: [0x016EC7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAD,
        len: 1,
        target: [0x016EC8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAE,
        len: 1,
        target: [0x016EC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EAF,
        len: 1,
        target: [0x016ECA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB0,
        len: 1,
        target: [0x016ECB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB1,
        len: 1,
        target: [0x016ECC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB2,
        len: 1,
        target: [0x016ECD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB3,
        len: 1,
        target: [0x016ECE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB4,
        len: 1,
        target: [0x016ECF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB5,
        len: 1,
        target: [0x016ED0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB6,
        len: 1,
        target: [0x016ED1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB7,
        len: 1,
        target: [0x016ED2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EB8,
        len: 1,
        target: [0x016ED3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E900,
        len: 1,
        target: [0x01E922, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E901,
        len: 1,
        target: [0x01E923, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E902,
        len: 1,
        target: [0x01E924, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E903,
        len: 1,
        target: [0x01E925, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E904,
        len: 1,
        target: [0x01E926, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E905,
        len: 1,
        target: [0x01E927, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E906,
        len: 1,
        target: [0x01E928, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E907,
        len: 1,
        target: [0x01E929, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E908,
        len: 1,
        target: [0x01E92A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E909,
        len: 1,
        target: [0x01E92B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90A,
        len: 1,
        target: [0x01E92C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90B,
        len: 1,
        target: [0x01E92D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90C,
        len: 1,
        target: [0x01E92E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90D,
        len: 1,
        target: [0x01E92F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90E,
        len: 1,
        target: [0x01E930, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E90F,
        len: 1,
        target: [0x01E931, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E910,
        len: 1,
        target: [0x01E932, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E911,
        len: 1,
        target: [0x01E933, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E912,
        len: 1,
        target: [0x01E934, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E913,
        len: 1,
        target: [0x01E935, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E914,
        len: 1,
        target: [0x01E936, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E915,
        len: 1,
        target: [0x01E937, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E916,
        len: 1,
        target: [0x01E938, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E917,
        len: 1,
        target: [0x01E939, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E918,
        len: 1,
        target: [0x01E93A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E919,
        len: 1,
        target: [0x01E93B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91A,
        len: 1,
        target: [0x01E93C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91B,
        len: 1,
        target: [0x01E93D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91C,
        len: 1,
        target: [0x01E93E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91D,
        len: 1,
        target: [0x01E93F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91E,
        len: 1,
        target: [0x01E940, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E91F,
        len: 1,
        target: [0x01E941, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E920,
        len: 1,
        target: [0x01E942, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E921,
        len: 1,
        target: [0x01E943, 0x000000, 0x000000],
    },
];

const UPPERCASE: &[Mapping] = &[
    Mapping {
        source: 0x000061,
        len: 1,
        target: [0x000041, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000062,
        len: 1,
        target: [0x000042, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000063,
        len: 1,
        target: [0x000043, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000064,
        len: 1,
        target: [0x000044, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000065,
        len: 1,
        target: [0x000045, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000066,
        len: 1,
        target: [0x000046, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000067,
        len: 1,
        target: [0x000047, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000068,
        len: 1,
        target: [0x000048, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000069,
        len: 1,
        target: [0x000049, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00006A,
        len: 1,
        target: [0x00004A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00006B,
        len: 1,
        target: [0x00004B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00006C,
        len: 1,
        target: [0x00004C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00006D,
        len: 1,
        target: [0x00004D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00006E,
        len: 1,
        target: [0x00004E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00006F,
        len: 1,
        target: [0x00004F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000070,
        len: 1,
        target: [0x000050, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000071,
        len: 1,
        target: [0x000051, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000072,
        len: 1,
        target: [0x000052, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000073,
        len: 1,
        target: [0x000053, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000074,
        len: 1,
        target: [0x000054, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000075,
        len: 1,
        target: [0x000055, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000076,
        len: 1,
        target: [0x000056, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000077,
        len: 1,
        target: [0x000057, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000078,
        len: 1,
        target: [0x000058, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000079,
        len: 1,
        target: [0x000059, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00007A,
        len: 1,
        target: [0x00005A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000B5,
        len: 1,
        target: [0x00039C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000DF,
        len: 2,
        target: [0x000053, 0x000053, 0x000000],
    },
    Mapping {
        source: 0x0000E0,
        len: 1,
        target: [0x0000C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E1,
        len: 1,
        target: [0x0000C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E2,
        len: 1,
        target: [0x0000C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E3,
        len: 1,
        target: [0x0000C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E4,
        len: 1,
        target: [0x0000C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E5,
        len: 1,
        target: [0x0000C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E6,
        len: 1,
        target: [0x0000C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E7,
        len: 1,
        target: [0x0000C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E8,
        len: 1,
        target: [0x0000C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000E9,
        len: 1,
        target: [0x0000C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000EA,
        len: 1,
        target: [0x0000CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000EB,
        len: 1,
        target: [0x0000CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000EC,
        len: 1,
        target: [0x0000CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000ED,
        len: 1,
        target: [0x0000CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000EE,
        len: 1,
        target: [0x0000CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000EF,
        len: 1,
        target: [0x0000CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F0,
        len: 1,
        target: [0x0000D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F1,
        len: 1,
        target: [0x0000D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F2,
        len: 1,
        target: [0x0000D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F3,
        len: 1,
        target: [0x0000D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F4,
        len: 1,
        target: [0x0000D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F5,
        len: 1,
        target: [0x0000D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F6,
        len: 1,
        target: [0x0000D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F8,
        len: 1,
        target: [0x0000D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000F9,
        len: 1,
        target: [0x0000D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000FA,
        len: 1,
        target: [0x0000DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000FB,
        len: 1,
        target: [0x0000DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000FC,
        len: 1,
        target: [0x0000DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000FD,
        len: 1,
        target: [0x0000DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000FE,
        len: 1,
        target: [0x0000DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0000FF,
        len: 1,
        target: [0x000178, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000101,
        len: 1,
        target: [0x000100, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000103,
        len: 1,
        target: [0x000102, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000105,
        len: 1,
        target: [0x000104, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000107,
        len: 1,
        target: [0x000106, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000109,
        len: 1,
        target: [0x000108, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010B,
        len: 1,
        target: [0x00010A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010D,
        len: 1,
        target: [0x00010C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00010F,
        len: 1,
        target: [0x00010E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000111,
        len: 1,
        target: [0x000110, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000113,
        len: 1,
        target: [0x000112, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000115,
        len: 1,
        target: [0x000114, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000117,
        len: 1,
        target: [0x000116, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000119,
        len: 1,
        target: [0x000118, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011B,
        len: 1,
        target: [0x00011A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011D,
        len: 1,
        target: [0x00011C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00011F,
        len: 1,
        target: [0x00011E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000121,
        len: 1,
        target: [0x000120, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000123,
        len: 1,
        target: [0x000122, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000125,
        len: 1,
        target: [0x000124, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000127,
        len: 1,
        target: [0x000126, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000129,
        len: 1,
        target: [0x000128, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012B,
        len: 1,
        target: [0x00012A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012D,
        len: 1,
        target: [0x00012C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00012F,
        len: 1,
        target: [0x00012E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000131,
        len: 1,
        target: [0x000049, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000133,
        len: 1,
        target: [0x000132, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000135,
        len: 1,
        target: [0x000134, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000137,
        len: 1,
        target: [0x000136, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013A,
        len: 1,
        target: [0x000139, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013C,
        len: 1,
        target: [0x00013B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00013E,
        len: 1,
        target: [0x00013D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000140,
        len: 1,
        target: [0x00013F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000142,
        len: 1,
        target: [0x000141, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000144,
        len: 1,
        target: [0x000143, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000146,
        len: 1,
        target: [0x000145, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000148,
        len: 1,
        target: [0x000147, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000149,
        len: 2,
        target: [0x0002BC, 0x00004E, 0x000000],
    },
    Mapping {
        source: 0x00014B,
        len: 1,
        target: [0x00014A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00014D,
        len: 1,
        target: [0x00014C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00014F,
        len: 1,
        target: [0x00014E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000151,
        len: 1,
        target: [0x000150, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000153,
        len: 1,
        target: [0x000152, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000155,
        len: 1,
        target: [0x000154, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000157,
        len: 1,
        target: [0x000156, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000159,
        len: 1,
        target: [0x000158, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015B,
        len: 1,
        target: [0x00015A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015D,
        len: 1,
        target: [0x00015C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00015F,
        len: 1,
        target: [0x00015E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000161,
        len: 1,
        target: [0x000160, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000163,
        len: 1,
        target: [0x000162, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000165,
        len: 1,
        target: [0x000164, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000167,
        len: 1,
        target: [0x000166, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000169,
        len: 1,
        target: [0x000168, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016B,
        len: 1,
        target: [0x00016A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016D,
        len: 1,
        target: [0x00016C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00016F,
        len: 1,
        target: [0x00016E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000171,
        len: 1,
        target: [0x000170, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000173,
        len: 1,
        target: [0x000172, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000175,
        len: 1,
        target: [0x000174, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000177,
        len: 1,
        target: [0x000176, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017A,
        len: 1,
        target: [0x000179, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017C,
        len: 1,
        target: [0x00017B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017E,
        len: 1,
        target: [0x00017D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00017F,
        len: 1,
        target: [0x000053, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000180,
        len: 1,
        target: [0x000243, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000183,
        len: 1,
        target: [0x000182, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000185,
        len: 1,
        target: [0x000184, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000188,
        len: 1,
        target: [0x000187, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00018C,
        len: 1,
        target: [0x00018B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000192,
        len: 1,
        target: [0x000191, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000195,
        len: 1,
        target: [0x0001F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000199,
        len: 1,
        target: [0x000198, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019A,
        len: 1,
        target: [0x00023D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019B,
        len: 1,
        target: [0x00A7DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00019E,
        len: 1,
        target: [0x000220, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A1,
        len: 1,
        target: [0x0001A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A3,
        len: 1,
        target: [0x0001A2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A5,
        len: 1,
        target: [0x0001A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001A8,
        len: 1,
        target: [0x0001A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001AD,
        len: 1,
        target: [0x0001AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B0,
        len: 1,
        target: [0x0001AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B4,
        len: 1,
        target: [0x0001B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B6,
        len: 1,
        target: [0x0001B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001B9,
        len: 1,
        target: [0x0001B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001BD,
        len: 1,
        target: [0x0001BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001BF,
        len: 1,
        target: [0x0001F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C5,
        len: 1,
        target: [0x0001C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C6,
        len: 1,
        target: [0x0001C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C8,
        len: 1,
        target: [0x0001C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001C9,
        len: 1,
        target: [0x0001C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CB,
        len: 1,
        target: [0x0001CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CC,
        len: 1,
        target: [0x0001CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001CE,
        len: 1,
        target: [0x0001CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D0,
        len: 1,
        target: [0x0001CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D2,
        len: 1,
        target: [0x0001D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D4,
        len: 1,
        target: [0x0001D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D6,
        len: 1,
        target: [0x0001D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001D8,
        len: 1,
        target: [0x0001D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DA,
        len: 1,
        target: [0x0001D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DC,
        len: 1,
        target: [0x0001DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DD,
        len: 1,
        target: [0x00018E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001DF,
        len: 1,
        target: [0x0001DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E1,
        len: 1,
        target: [0x0001E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E3,
        len: 1,
        target: [0x0001E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E5,
        len: 1,
        target: [0x0001E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E7,
        len: 1,
        target: [0x0001E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001E9,
        len: 1,
        target: [0x0001E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EB,
        len: 1,
        target: [0x0001EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001ED,
        len: 1,
        target: [0x0001EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001EF,
        len: 1,
        target: [0x0001EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F0,
        len: 2,
        target: [0x00004A, 0x00030C, 0x000000],
    },
    Mapping {
        source: 0x0001F2,
        len: 1,
        target: [0x0001F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F3,
        len: 1,
        target: [0x0001F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F5,
        len: 1,
        target: [0x0001F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001F9,
        len: 1,
        target: [0x0001F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FB,
        len: 1,
        target: [0x0001FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FD,
        len: 1,
        target: [0x0001FC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0001FF,
        len: 1,
        target: [0x0001FE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000201,
        len: 1,
        target: [0x000200, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000203,
        len: 1,
        target: [0x000202, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000205,
        len: 1,
        target: [0x000204, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000207,
        len: 1,
        target: [0x000206, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000209,
        len: 1,
        target: [0x000208, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020B,
        len: 1,
        target: [0x00020A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020D,
        len: 1,
        target: [0x00020C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00020F,
        len: 1,
        target: [0x00020E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000211,
        len: 1,
        target: [0x000210, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000213,
        len: 1,
        target: [0x000212, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000215,
        len: 1,
        target: [0x000214, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000217,
        len: 1,
        target: [0x000216, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000219,
        len: 1,
        target: [0x000218, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021B,
        len: 1,
        target: [0x00021A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021D,
        len: 1,
        target: [0x00021C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00021F,
        len: 1,
        target: [0x00021E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000223,
        len: 1,
        target: [0x000222, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000225,
        len: 1,
        target: [0x000224, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000227,
        len: 1,
        target: [0x000226, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000229,
        len: 1,
        target: [0x000228, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022B,
        len: 1,
        target: [0x00022A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022D,
        len: 1,
        target: [0x00022C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00022F,
        len: 1,
        target: [0x00022E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000231,
        len: 1,
        target: [0x000230, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000233,
        len: 1,
        target: [0x000232, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023C,
        len: 1,
        target: [0x00023B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00023F,
        len: 1,
        target: [0x002C7E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000240,
        len: 1,
        target: [0x002C7F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000242,
        len: 1,
        target: [0x000241, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000247,
        len: 1,
        target: [0x000246, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000249,
        len: 1,
        target: [0x000248, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024B,
        len: 1,
        target: [0x00024A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024D,
        len: 1,
        target: [0x00024C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00024F,
        len: 1,
        target: [0x00024E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000250,
        len: 1,
        target: [0x002C6F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000251,
        len: 1,
        target: [0x002C6D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000252,
        len: 1,
        target: [0x002C70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000253,
        len: 1,
        target: [0x000181, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000254,
        len: 1,
        target: [0x000186, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000256,
        len: 1,
        target: [0x000189, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000257,
        len: 1,
        target: [0x00018A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000259,
        len: 1,
        target: [0x00018F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00025B,
        len: 1,
        target: [0x000190, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00025C,
        len: 1,
        target: [0x00A7AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000260,
        len: 1,
        target: [0x000193, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000261,
        len: 1,
        target: [0x00A7AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000263,
        len: 1,
        target: [0x000194, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000264,
        len: 1,
        target: [0x00A7CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000265,
        len: 1,
        target: [0x00A78D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000266,
        len: 1,
        target: [0x00A7AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000268,
        len: 1,
        target: [0x000197, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000269,
        len: 1,
        target: [0x000196, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00026A,
        len: 1,
        target: [0x00A7AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00026B,
        len: 1,
        target: [0x002C62, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00026C,
        len: 1,
        target: [0x00A7AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00026F,
        len: 1,
        target: [0x00019C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000271,
        len: 1,
        target: [0x002C6E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000272,
        len: 1,
        target: [0x00019D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000275,
        len: 1,
        target: [0x00019F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00027D,
        len: 1,
        target: [0x002C64, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000280,
        len: 1,
        target: [0x0001A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000282,
        len: 1,
        target: [0x00A7C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000283,
        len: 1,
        target: [0x0001A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000287,
        len: 1,
        target: [0x00A7B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000288,
        len: 1,
        target: [0x0001AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000289,
        len: 1,
        target: [0x000244, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00028A,
        len: 1,
        target: [0x0001B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00028B,
        len: 1,
        target: [0x0001B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00028C,
        len: 1,
        target: [0x000245, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000292,
        len: 1,
        target: [0x0001B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00029D,
        len: 1,
        target: [0x00A7B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00029E,
        len: 1,
        target: [0x00A7B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000345,
        len: 1,
        target: [0x000399, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000371,
        len: 1,
        target: [0x000370, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000373,
        len: 1,
        target: [0x000372, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000377,
        len: 1,
        target: [0x000376, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00037B,
        len: 1,
        target: [0x0003FD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00037C,
        len: 1,
        target: [0x0003FE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00037D,
        len: 1,
        target: [0x0003FF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000390,
        len: 3,
        target: [0x000399, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x0003AC,
        len: 1,
        target: [0x000386, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003AD,
        len: 1,
        target: [0x000388, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003AE,
        len: 1,
        target: [0x000389, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003AF,
        len: 1,
        target: [0x00038A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B0,
        len: 3,
        target: [0x0003A5, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x0003B1,
        len: 1,
        target: [0x000391, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B2,
        len: 1,
        target: [0x000392, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B3,
        len: 1,
        target: [0x000393, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B4,
        len: 1,
        target: [0x000394, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B5,
        len: 1,
        target: [0x000395, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B6,
        len: 1,
        target: [0x000396, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B7,
        len: 1,
        target: [0x000397, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B8,
        len: 1,
        target: [0x000398, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003B9,
        len: 1,
        target: [0x000399, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003BA,
        len: 1,
        target: [0x00039A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003BB,
        len: 1,
        target: [0x00039B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003BC,
        len: 1,
        target: [0x00039C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003BD,
        len: 1,
        target: [0x00039D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003BE,
        len: 1,
        target: [0x00039E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003BF,
        len: 1,
        target: [0x00039F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C0,
        len: 1,
        target: [0x0003A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C1,
        len: 1,
        target: [0x0003A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C2,
        len: 1,
        target: [0x0003A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C3,
        len: 1,
        target: [0x0003A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C4,
        len: 1,
        target: [0x0003A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C5,
        len: 1,
        target: [0x0003A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C6,
        len: 1,
        target: [0x0003A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C7,
        len: 1,
        target: [0x0003A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C8,
        len: 1,
        target: [0x0003A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003C9,
        len: 1,
        target: [0x0003A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003CA,
        len: 1,
        target: [0x0003AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003CB,
        len: 1,
        target: [0x0003AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003CC,
        len: 1,
        target: [0x00038C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003CD,
        len: 1,
        target: [0x00038E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003CE,
        len: 1,
        target: [0x00038F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D0,
        len: 1,
        target: [0x000392, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D1,
        len: 1,
        target: [0x000398, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D5,
        len: 1,
        target: [0x0003A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D6,
        len: 1,
        target: [0x0003A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D7,
        len: 1,
        target: [0x0003CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003D9,
        len: 1,
        target: [0x0003D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DB,
        len: 1,
        target: [0x0003DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DD,
        len: 1,
        target: [0x0003DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003DF,
        len: 1,
        target: [0x0003DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E1,
        len: 1,
        target: [0x0003E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E3,
        len: 1,
        target: [0x0003E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E5,
        len: 1,
        target: [0x0003E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E7,
        len: 1,
        target: [0x0003E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003E9,
        len: 1,
        target: [0x0003E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EB,
        len: 1,
        target: [0x0003EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003ED,
        len: 1,
        target: [0x0003EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003EF,
        len: 1,
        target: [0x0003EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F0,
        len: 1,
        target: [0x00039A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F1,
        len: 1,
        target: [0x0003A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F2,
        len: 1,
        target: [0x0003F9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F3,
        len: 1,
        target: [0x00037F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F5,
        len: 1,
        target: [0x000395, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003F8,
        len: 1,
        target: [0x0003F7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0003FB,
        len: 1,
        target: [0x0003FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000430,
        len: 1,
        target: [0x000410, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000431,
        len: 1,
        target: [0x000411, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000432,
        len: 1,
        target: [0x000412, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000433,
        len: 1,
        target: [0x000413, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000434,
        len: 1,
        target: [0x000414, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000435,
        len: 1,
        target: [0x000415, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000436,
        len: 1,
        target: [0x000416, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000437,
        len: 1,
        target: [0x000417, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000438,
        len: 1,
        target: [0x000418, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000439,
        len: 1,
        target: [0x000419, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00043A,
        len: 1,
        target: [0x00041A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00043B,
        len: 1,
        target: [0x00041B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00043C,
        len: 1,
        target: [0x00041C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00043D,
        len: 1,
        target: [0x00041D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00043E,
        len: 1,
        target: [0x00041E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00043F,
        len: 1,
        target: [0x00041F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000440,
        len: 1,
        target: [0x000420, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000441,
        len: 1,
        target: [0x000421, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000442,
        len: 1,
        target: [0x000422, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000443,
        len: 1,
        target: [0x000423, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000444,
        len: 1,
        target: [0x000424, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000445,
        len: 1,
        target: [0x000425, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000446,
        len: 1,
        target: [0x000426, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000447,
        len: 1,
        target: [0x000427, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000448,
        len: 1,
        target: [0x000428, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000449,
        len: 1,
        target: [0x000429, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00044A,
        len: 1,
        target: [0x00042A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00044B,
        len: 1,
        target: [0x00042B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00044C,
        len: 1,
        target: [0x00042C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00044D,
        len: 1,
        target: [0x00042D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00044E,
        len: 1,
        target: [0x00042E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00044F,
        len: 1,
        target: [0x00042F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000450,
        len: 1,
        target: [0x000400, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000451,
        len: 1,
        target: [0x000401, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000452,
        len: 1,
        target: [0x000402, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000453,
        len: 1,
        target: [0x000403, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000454,
        len: 1,
        target: [0x000404, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000455,
        len: 1,
        target: [0x000405, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000456,
        len: 1,
        target: [0x000406, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000457,
        len: 1,
        target: [0x000407, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000458,
        len: 1,
        target: [0x000408, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000459,
        len: 1,
        target: [0x000409, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00045A,
        len: 1,
        target: [0x00040A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00045B,
        len: 1,
        target: [0x00040B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00045C,
        len: 1,
        target: [0x00040C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00045D,
        len: 1,
        target: [0x00040D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00045E,
        len: 1,
        target: [0x00040E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00045F,
        len: 1,
        target: [0x00040F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000461,
        len: 1,
        target: [0x000460, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000463,
        len: 1,
        target: [0x000462, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000465,
        len: 1,
        target: [0x000464, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000467,
        len: 1,
        target: [0x000466, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000469,
        len: 1,
        target: [0x000468, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046B,
        len: 1,
        target: [0x00046A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046D,
        len: 1,
        target: [0x00046C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00046F,
        len: 1,
        target: [0x00046E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000471,
        len: 1,
        target: [0x000470, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000473,
        len: 1,
        target: [0x000472, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000475,
        len: 1,
        target: [0x000474, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000477,
        len: 1,
        target: [0x000476, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000479,
        len: 1,
        target: [0x000478, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047B,
        len: 1,
        target: [0x00047A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047D,
        len: 1,
        target: [0x00047C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00047F,
        len: 1,
        target: [0x00047E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000481,
        len: 1,
        target: [0x000480, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048B,
        len: 1,
        target: [0x00048A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048D,
        len: 1,
        target: [0x00048C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00048F,
        len: 1,
        target: [0x00048E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000491,
        len: 1,
        target: [0x000490, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000493,
        len: 1,
        target: [0x000492, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000495,
        len: 1,
        target: [0x000494, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000497,
        len: 1,
        target: [0x000496, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000499,
        len: 1,
        target: [0x000498, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049B,
        len: 1,
        target: [0x00049A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049D,
        len: 1,
        target: [0x00049C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00049F,
        len: 1,
        target: [0x00049E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A1,
        len: 1,
        target: [0x0004A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A3,
        len: 1,
        target: [0x0004A2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A5,
        len: 1,
        target: [0x0004A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A7,
        len: 1,
        target: [0x0004A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004A9,
        len: 1,
        target: [0x0004A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AB,
        len: 1,
        target: [0x0004AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AD,
        len: 1,
        target: [0x0004AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004AF,
        len: 1,
        target: [0x0004AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B1,
        len: 1,
        target: [0x0004B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B3,
        len: 1,
        target: [0x0004B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B5,
        len: 1,
        target: [0x0004B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B7,
        len: 1,
        target: [0x0004B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004B9,
        len: 1,
        target: [0x0004B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BB,
        len: 1,
        target: [0x0004BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BD,
        len: 1,
        target: [0x0004BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004BF,
        len: 1,
        target: [0x0004BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C2,
        len: 1,
        target: [0x0004C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C4,
        len: 1,
        target: [0x0004C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C6,
        len: 1,
        target: [0x0004C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004C8,
        len: 1,
        target: [0x0004C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CA,
        len: 1,
        target: [0x0004C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CC,
        len: 1,
        target: [0x0004CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CE,
        len: 1,
        target: [0x0004CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004CF,
        len: 1,
        target: [0x0004C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D1,
        len: 1,
        target: [0x0004D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D3,
        len: 1,
        target: [0x0004D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D5,
        len: 1,
        target: [0x0004D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D7,
        len: 1,
        target: [0x0004D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004D9,
        len: 1,
        target: [0x0004D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DB,
        len: 1,
        target: [0x0004DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DD,
        len: 1,
        target: [0x0004DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004DF,
        len: 1,
        target: [0x0004DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E1,
        len: 1,
        target: [0x0004E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E3,
        len: 1,
        target: [0x0004E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E5,
        len: 1,
        target: [0x0004E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E7,
        len: 1,
        target: [0x0004E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004E9,
        len: 1,
        target: [0x0004E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EB,
        len: 1,
        target: [0x0004EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004ED,
        len: 1,
        target: [0x0004EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004EF,
        len: 1,
        target: [0x0004EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F1,
        len: 1,
        target: [0x0004F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F3,
        len: 1,
        target: [0x0004F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F5,
        len: 1,
        target: [0x0004F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F7,
        len: 1,
        target: [0x0004F6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004F9,
        len: 1,
        target: [0x0004F8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FB,
        len: 1,
        target: [0x0004FA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FD,
        len: 1,
        target: [0x0004FC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0004FF,
        len: 1,
        target: [0x0004FE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000501,
        len: 1,
        target: [0x000500, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000503,
        len: 1,
        target: [0x000502, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000505,
        len: 1,
        target: [0x000504, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000507,
        len: 1,
        target: [0x000506, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000509,
        len: 1,
        target: [0x000508, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050B,
        len: 1,
        target: [0x00050A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050D,
        len: 1,
        target: [0x00050C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00050F,
        len: 1,
        target: [0x00050E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000511,
        len: 1,
        target: [0x000510, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000513,
        len: 1,
        target: [0x000512, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000515,
        len: 1,
        target: [0x000514, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000517,
        len: 1,
        target: [0x000516, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000519,
        len: 1,
        target: [0x000518, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051B,
        len: 1,
        target: [0x00051A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051D,
        len: 1,
        target: [0x00051C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00051F,
        len: 1,
        target: [0x00051E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000521,
        len: 1,
        target: [0x000520, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000523,
        len: 1,
        target: [0x000522, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000525,
        len: 1,
        target: [0x000524, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000527,
        len: 1,
        target: [0x000526, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000529,
        len: 1,
        target: [0x000528, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052B,
        len: 1,
        target: [0x00052A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052D,
        len: 1,
        target: [0x00052C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00052F,
        len: 1,
        target: [0x00052E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000561,
        len: 1,
        target: [0x000531, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000562,
        len: 1,
        target: [0x000532, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000563,
        len: 1,
        target: [0x000533, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000564,
        len: 1,
        target: [0x000534, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000565,
        len: 1,
        target: [0x000535, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000566,
        len: 1,
        target: [0x000536, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000567,
        len: 1,
        target: [0x000537, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000568,
        len: 1,
        target: [0x000538, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000569,
        len: 1,
        target: [0x000539, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00056A,
        len: 1,
        target: [0x00053A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00056B,
        len: 1,
        target: [0x00053B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00056C,
        len: 1,
        target: [0x00053C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00056D,
        len: 1,
        target: [0x00053D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00056E,
        len: 1,
        target: [0x00053E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00056F,
        len: 1,
        target: [0x00053F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000570,
        len: 1,
        target: [0x000540, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000571,
        len: 1,
        target: [0x000541, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000572,
        len: 1,
        target: [0x000542, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000573,
        len: 1,
        target: [0x000543, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000574,
        len: 1,
        target: [0x000544, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000575,
        len: 1,
        target: [0x000545, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000576,
        len: 1,
        target: [0x000546, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000577,
        len: 1,
        target: [0x000547, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000578,
        len: 1,
        target: [0x000548, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000579,
        len: 1,
        target: [0x000549, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00057A,
        len: 1,
        target: [0x00054A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00057B,
        len: 1,
        target: [0x00054B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00057C,
        len: 1,
        target: [0x00054C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00057D,
        len: 1,
        target: [0x00054D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00057E,
        len: 1,
        target: [0x00054E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00057F,
        len: 1,
        target: [0x00054F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000580,
        len: 1,
        target: [0x000550, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000581,
        len: 1,
        target: [0x000551, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000582,
        len: 1,
        target: [0x000552, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000583,
        len: 1,
        target: [0x000553, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000584,
        len: 1,
        target: [0x000554, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000585,
        len: 1,
        target: [0x000555, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000586,
        len: 1,
        target: [0x000556, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x000587,
        len: 2,
        target: [0x000535, 0x000552, 0x000000],
    },
    Mapping {
        source: 0x0010D0,
        len: 1,
        target: [0x001C90, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D1,
        len: 1,
        target: [0x001C91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D2,
        len: 1,
        target: [0x001C92, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D3,
        len: 1,
        target: [0x001C93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D4,
        len: 1,
        target: [0x001C94, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D5,
        len: 1,
        target: [0x001C95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D6,
        len: 1,
        target: [0x001C96, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D7,
        len: 1,
        target: [0x001C97, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D8,
        len: 1,
        target: [0x001C98, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010D9,
        len: 1,
        target: [0x001C99, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010DA,
        len: 1,
        target: [0x001C9A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010DB,
        len: 1,
        target: [0x001C9B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010DC,
        len: 1,
        target: [0x001C9C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010DD,
        len: 1,
        target: [0x001C9D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010DE,
        len: 1,
        target: [0x001C9E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010DF,
        len: 1,
        target: [0x001C9F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E0,
        len: 1,
        target: [0x001CA0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E1,
        len: 1,
        target: [0x001CA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E2,
        len: 1,
        target: [0x001CA2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E3,
        len: 1,
        target: [0x001CA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E4,
        len: 1,
        target: [0x001CA4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E5,
        len: 1,
        target: [0x001CA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E6,
        len: 1,
        target: [0x001CA6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E7,
        len: 1,
        target: [0x001CA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E8,
        len: 1,
        target: [0x001CA8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010E9,
        len: 1,
        target: [0x001CA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010EA,
        len: 1,
        target: [0x001CAA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010EB,
        len: 1,
        target: [0x001CAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010EC,
        len: 1,
        target: [0x001CAC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010ED,
        len: 1,
        target: [0x001CAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010EE,
        len: 1,
        target: [0x001CAE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010EF,
        len: 1,
        target: [0x001CAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F0,
        len: 1,
        target: [0x001CB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F1,
        len: 1,
        target: [0x001CB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F2,
        len: 1,
        target: [0x001CB2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F3,
        len: 1,
        target: [0x001CB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F4,
        len: 1,
        target: [0x001CB4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F5,
        len: 1,
        target: [0x001CB5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F6,
        len: 1,
        target: [0x001CB6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F7,
        len: 1,
        target: [0x001CB7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F8,
        len: 1,
        target: [0x001CB8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010F9,
        len: 1,
        target: [0x001CB9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010FA,
        len: 1,
        target: [0x001CBA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010FD,
        len: 1,
        target: [0x001CBD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010FE,
        len: 1,
        target: [0x001CBE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0010FF,
        len: 1,
        target: [0x001CBF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F8,
        len: 1,
        target: [0x0013F0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013F9,
        len: 1,
        target: [0x0013F1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FA,
        len: 1,
        target: [0x0013F2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FB,
        len: 1,
        target: [0x0013F3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FC,
        len: 1,
        target: [0x0013F4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0013FD,
        len: 1,
        target: [0x0013F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C80,
        len: 1,
        target: [0x000412, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C81,
        len: 1,
        target: [0x000414, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C82,
        len: 1,
        target: [0x00041E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C83,
        len: 1,
        target: [0x000421, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C84,
        len: 1,
        target: [0x000422, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C85,
        len: 1,
        target: [0x000422, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C86,
        len: 1,
        target: [0x00042A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C87,
        len: 1,
        target: [0x000462, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C88,
        len: 1,
        target: [0x00A64A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001C8A,
        len: 1,
        target: [0x001C89, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001D79,
        len: 1,
        target: [0x00A77D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001D7D,
        len: 1,
        target: [0x002C63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001D8E,
        len: 1,
        target: [0x00A7C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E01,
        len: 1,
        target: [0x001E00, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E03,
        len: 1,
        target: [0x001E02, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E05,
        len: 1,
        target: [0x001E04, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E07,
        len: 1,
        target: [0x001E06, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E09,
        len: 1,
        target: [0x001E08, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0B,
        len: 1,
        target: [0x001E0A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0D,
        len: 1,
        target: [0x001E0C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E0F,
        len: 1,
        target: [0x001E0E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E11,
        len: 1,
        target: [0x001E10, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E13,
        len: 1,
        target: [0x001E12, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E15,
        len: 1,
        target: [0x001E14, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E17,
        len: 1,
        target: [0x001E16, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E19,
        len: 1,
        target: [0x001E18, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1B,
        len: 1,
        target: [0x001E1A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1D,
        len: 1,
        target: [0x001E1C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E1F,
        len: 1,
        target: [0x001E1E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E21,
        len: 1,
        target: [0x001E20, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E23,
        len: 1,
        target: [0x001E22, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E25,
        len: 1,
        target: [0x001E24, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E27,
        len: 1,
        target: [0x001E26, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E29,
        len: 1,
        target: [0x001E28, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2B,
        len: 1,
        target: [0x001E2A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2D,
        len: 1,
        target: [0x001E2C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E2F,
        len: 1,
        target: [0x001E2E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E31,
        len: 1,
        target: [0x001E30, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E33,
        len: 1,
        target: [0x001E32, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E35,
        len: 1,
        target: [0x001E34, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E37,
        len: 1,
        target: [0x001E36, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E39,
        len: 1,
        target: [0x001E38, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3B,
        len: 1,
        target: [0x001E3A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3D,
        len: 1,
        target: [0x001E3C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E3F,
        len: 1,
        target: [0x001E3E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E41,
        len: 1,
        target: [0x001E40, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E43,
        len: 1,
        target: [0x001E42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E45,
        len: 1,
        target: [0x001E44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E47,
        len: 1,
        target: [0x001E46, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E49,
        len: 1,
        target: [0x001E48, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4B,
        len: 1,
        target: [0x001E4A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4D,
        len: 1,
        target: [0x001E4C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E4F,
        len: 1,
        target: [0x001E4E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E51,
        len: 1,
        target: [0x001E50, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E53,
        len: 1,
        target: [0x001E52, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E55,
        len: 1,
        target: [0x001E54, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E57,
        len: 1,
        target: [0x001E56, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E59,
        len: 1,
        target: [0x001E58, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5B,
        len: 1,
        target: [0x001E5A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5D,
        len: 1,
        target: [0x001E5C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E5F,
        len: 1,
        target: [0x001E5E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E61,
        len: 1,
        target: [0x001E60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E63,
        len: 1,
        target: [0x001E62, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E65,
        len: 1,
        target: [0x001E64, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E67,
        len: 1,
        target: [0x001E66, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E69,
        len: 1,
        target: [0x001E68, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6B,
        len: 1,
        target: [0x001E6A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6D,
        len: 1,
        target: [0x001E6C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E6F,
        len: 1,
        target: [0x001E6E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E71,
        len: 1,
        target: [0x001E70, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E73,
        len: 1,
        target: [0x001E72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E75,
        len: 1,
        target: [0x001E74, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E77,
        len: 1,
        target: [0x001E76, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E79,
        len: 1,
        target: [0x001E78, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7B,
        len: 1,
        target: [0x001E7A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7D,
        len: 1,
        target: [0x001E7C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E7F,
        len: 1,
        target: [0x001E7E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E81,
        len: 1,
        target: [0x001E80, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E83,
        len: 1,
        target: [0x001E82, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E85,
        len: 1,
        target: [0x001E84, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E87,
        len: 1,
        target: [0x001E86, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E89,
        len: 1,
        target: [0x001E88, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8B,
        len: 1,
        target: [0x001E8A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8D,
        len: 1,
        target: [0x001E8C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E8F,
        len: 1,
        target: [0x001E8E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E91,
        len: 1,
        target: [0x001E90, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E93,
        len: 1,
        target: [0x001E92, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E95,
        len: 1,
        target: [0x001E94, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001E96,
        len: 2,
        target: [0x000048, 0x000331, 0x000000],
    },
    Mapping {
        source: 0x001E97,
        len: 2,
        target: [0x000054, 0x000308, 0x000000],
    },
    Mapping {
        source: 0x001E98,
        len: 2,
        target: [0x000057, 0x00030A, 0x000000],
    },
    Mapping {
        source: 0x001E99,
        len: 2,
        target: [0x000059, 0x00030A, 0x000000],
    },
    Mapping {
        source: 0x001E9A,
        len: 2,
        target: [0x000041, 0x0002BE, 0x000000],
    },
    Mapping {
        source: 0x001E9B,
        len: 1,
        target: [0x001E60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA1,
        len: 1,
        target: [0x001EA0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA3,
        len: 1,
        target: [0x001EA2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA5,
        len: 1,
        target: [0x001EA4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA7,
        len: 1,
        target: [0x001EA6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EA9,
        len: 1,
        target: [0x001EA8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAB,
        len: 1,
        target: [0x001EAA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAD,
        len: 1,
        target: [0x001EAC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EAF,
        len: 1,
        target: [0x001EAE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB1,
        len: 1,
        target: [0x001EB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB3,
        len: 1,
        target: [0x001EB2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB5,
        len: 1,
        target: [0x001EB4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB7,
        len: 1,
        target: [0x001EB6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EB9,
        len: 1,
        target: [0x001EB8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBB,
        len: 1,
        target: [0x001EBA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBD,
        len: 1,
        target: [0x001EBC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EBF,
        len: 1,
        target: [0x001EBE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC1,
        len: 1,
        target: [0x001EC0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC3,
        len: 1,
        target: [0x001EC2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC5,
        len: 1,
        target: [0x001EC4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC7,
        len: 1,
        target: [0x001EC6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EC9,
        len: 1,
        target: [0x001EC8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECB,
        len: 1,
        target: [0x001ECA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECD,
        len: 1,
        target: [0x001ECC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ECF,
        len: 1,
        target: [0x001ECE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED1,
        len: 1,
        target: [0x001ED0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED3,
        len: 1,
        target: [0x001ED2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED5,
        len: 1,
        target: [0x001ED4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED7,
        len: 1,
        target: [0x001ED6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001ED9,
        len: 1,
        target: [0x001ED8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDB,
        len: 1,
        target: [0x001EDA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDD,
        len: 1,
        target: [0x001EDC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EDF,
        len: 1,
        target: [0x001EDE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE1,
        len: 1,
        target: [0x001EE0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE3,
        len: 1,
        target: [0x001EE2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE5,
        len: 1,
        target: [0x001EE4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE7,
        len: 1,
        target: [0x001EE6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EE9,
        len: 1,
        target: [0x001EE8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEB,
        len: 1,
        target: [0x001EEA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EED,
        len: 1,
        target: [0x001EEC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EEF,
        len: 1,
        target: [0x001EEE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF1,
        len: 1,
        target: [0x001EF0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF3,
        len: 1,
        target: [0x001EF2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF5,
        len: 1,
        target: [0x001EF4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF7,
        len: 1,
        target: [0x001EF6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EF9,
        len: 1,
        target: [0x001EF8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFB,
        len: 1,
        target: [0x001EFA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFD,
        len: 1,
        target: [0x001EFC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001EFF,
        len: 1,
        target: [0x001EFE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F00,
        len: 1,
        target: [0x001F08, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F01,
        len: 1,
        target: [0x001F09, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F02,
        len: 1,
        target: [0x001F0A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F03,
        len: 1,
        target: [0x001F0B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F04,
        len: 1,
        target: [0x001F0C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F05,
        len: 1,
        target: [0x001F0D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F06,
        len: 1,
        target: [0x001F0E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F07,
        len: 1,
        target: [0x001F0F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F10,
        len: 1,
        target: [0x001F18, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F11,
        len: 1,
        target: [0x001F19, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F12,
        len: 1,
        target: [0x001F1A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F13,
        len: 1,
        target: [0x001F1B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F14,
        len: 1,
        target: [0x001F1C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F15,
        len: 1,
        target: [0x001F1D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F20,
        len: 1,
        target: [0x001F28, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F21,
        len: 1,
        target: [0x001F29, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F22,
        len: 1,
        target: [0x001F2A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F23,
        len: 1,
        target: [0x001F2B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F24,
        len: 1,
        target: [0x001F2C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F25,
        len: 1,
        target: [0x001F2D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F26,
        len: 1,
        target: [0x001F2E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F27,
        len: 1,
        target: [0x001F2F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F30,
        len: 1,
        target: [0x001F38, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F31,
        len: 1,
        target: [0x001F39, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F32,
        len: 1,
        target: [0x001F3A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F33,
        len: 1,
        target: [0x001F3B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F34,
        len: 1,
        target: [0x001F3C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F35,
        len: 1,
        target: [0x001F3D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F36,
        len: 1,
        target: [0x001F3E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F37,
        len: 1,
        target: [0x001F3F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F40,
        len: 1,
        target: [0x001F48, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F41,
        len: 1,
        target: [0x001F49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F42,
        len: 1,
        target: [0x001F4A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F43,
        len: 1,
        target: [0x001F4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F44,
        len: 1,
        target: [0x001F4C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F45,
        len: 1,
        target: [0x001F4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F50,
        len: 2,
        target: [0x0003A5, 0x000313, 0x000000],
    },
    Mapping {
        source: 0x001F51,
        len: 1,
        target: [0x001F59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F52,
        len: 3,
        target: [0x0003A5, 0x000313, 0x000300],
    },
    Mapping {
        source: 0x001F53,
        len: 1,
        target: [0x001F5B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F54,
        len: 3,
        target: [0x0003A5, 0x000313, 0x000301],
    },
    Mapping {
        source: 0x001F55,
        len: 1,
        target: [0x001F5D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F56,
        len: 3,
        target: [0x0003A5, 0x000313, 0x000342],
    },
    Mapping {
        source: 0x001F57,
        len: 1,
        target: [0x001F5F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F60,
        len: 1,
        target: [0x001F68, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F61,
        len: 1,
        target: [0x001F69, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F62,
        len: 1,
        target: [0x001F6A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F63,
        len: 1,
        target: [0x001F6B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F64,
        len: 1,
        target: [0x001F6C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F65,
        len: 1,
        target: [0x001F6D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F66,
        len: 1,
        target: [0x001F6E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F67,
        len: 1,
        target: [0x001F6F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F70,
        len: 1,
        target: [0x001FBA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F71,
        len: 1,
        target: [0x001FBB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F72,
        len: 1,
        target: [0x001FC8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F73,
        len: 1,
        target: [0x001FC9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F74,
        len: 1,
        target: [0x001FCA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F75,
        len: 1,
        target: [0x001FCB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F76,
        len: 1,
        target: [0x001FDA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F77,
        len: 1,
        target: [0x001FDB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F78,
        len: 1,
        target: [0x001FF8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F79,
        len: 1,
        target: [0x001FF9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F7A,
        len: 1,
        target: [0x001FEA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F7B,
        len: 1,
        target: [0x001FEB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F7C,
        len: 1,
        target: [0x001FFA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F7D,
        len: 1,
        target: [0x001FFB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001F80,
        len: 2,
        target: [0x001F08, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F81,
        len: 2,
        target: [0x001F09, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F82,
        len: 2,
        target: [0x001F0A, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F83,
        len: 2,
        target: [0x001F0B, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F84,
        len: 2,
        target: [0x001F0C, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F85,
        len: 2,
        target: [0x001F0D, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F86,
        len: 2,
        target: [0x001F0E, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F87,
        len: 2,
        target: [0x001F0F, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F88,
        len: 2,
        target: [0x001F08, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F89,
        len: 2,
        target: [0x001F09, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F8A,
        len: 2,
        target: [0x001F0A, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F8B,
        len: 2,
        target: [0x001F0B, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F8C,
        len: 2,
        target: [0x001F0C, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F8D,
        len: 2,
        target: [0x001F0D, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F8E,
        len: 2,
        target: [0x001F0E, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F8F,
        len: 2,
        target: [0x001F0F, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F90,
        len: 2,
        target: [0x001F28, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F91,
        len: 2,
        target: [0x001F29, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F92,
        len: 2,
        target: [0x001F2A, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F93,
        len: 2,
        target: [0x001F2B, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F94,
        len: 2,
        target: [0x001F2C, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F95,
        len: 2,
        target: [0x001F2D, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F96,
        len: 2,
        target: [0x001F2E, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F97,
        len: 2,
        target: [0x001F2F, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F98,
        len: 2,
        target: [0x001F28, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F99,
        len: 2,
        target: [0x001F29, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F9A,
        len: 2,
        target: [0x001F2A, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F9B,
        len: 2,
        target: [0x001F2B, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F9C,
        len: 2,
        target: [0x001F2C, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F9D,
        len: 2,
        target: [0x001F2D, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F9E,
        len: 2,
        target: [0x001F2E, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001F9F,
        len: 2,
        target: [0x001F2F, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA0,
        len: 2,
        target: [0x001F68, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA1,
        len: 2,
        target: [0x001F69, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA2,
        len: 2,
        target: [0x001F6A, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA3,
        len: 2,
        target: [0x001F6B, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA4,
        len: 2,
        target: [0x001F6C, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA5,
        len: 2,
        target: [0x001F6D, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA6,
        len: 2,
        target: [0x001F6E, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA7,
        len: 2,
        target: [0x001F6F, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA8,
        len: 2,
        target: [0x001F68, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FA9,
        len: 2,
        target: [0x001F69, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FAA,
        len: 2,
        target: [0x001F6A, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FAB,
        len: 2,
        target: [0x001F6B, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FAC,
        len: 2,
        target: [0x001F6C, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FAD,
        len: 2,
        target: [0x001F6D, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FAE,
        len: 2,
        target: [0x001F6E, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FAF,
        len: 2,
        target: [0x001F6F, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FB0,
        len: 1,
        target: [0x001FB8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FB1,
        len: 1,
        target: [0x001FB9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FB2,
        len: 2,
        target: [0x001FBA, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FB3,
        len: 2,
        target: [0x000391, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FB4,
        len: 2,
        target: [0x000386, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FB6,
        len: 2,
        target: [0x000391, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FB7,
        len: 3,
        target: [0x000391, 0x000342, 0x000399],
    },
    Mapping {
        source: 0x001FBC,
        len: 2,
        target: [0x000391, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FBE,
        len: 1,
        target: [0x000399, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FC2,
        len: 2,
        target: [0x001FCA, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FC3,
        len: 2,
        target: [0x000397, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FC4,
        len: 2,
        target: [0x000389, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FC6,
        len: 2,
        target: [0x000397, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FC7,
        len: 3,
        target: [0x000397, 0x000342, 0x000399],
    },
    Mapping {
        source: 0x001FCC,
        len: 2,
        target: [0x000397, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FD0,
        len: 1,
        target: [0x001FD8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FD1,
        len: 1,
        target: [0x001FD9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FD2,
        len: 3,
        target: [0x000399, 0x000308, 0x000300],
    },
    Mapping {
        source: 0x001FD3,
        len: 3,
        target: [0x000399, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x001FD6,
        len: 2,
        target: [0x000399, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FD7,
        len: 3,
        target: [0x000399, 0x000308, 0x000342],
    },
    Mapping {
        source: 0x001FE0,
        len: 1,
        target: [0x001FE8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FE1,
        len: 1,
        target: [0x001FE9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FE2,
        len: 3,
        target: [0x0003A5, 0x000308, 0x000300],
    },
    Mapping {
        source: 0x001FE3,
        len: 3,
        target: [0x0003A5, 0x000308, 0x000301],
    },
    Mapping {
        source: 0x001FE4,
        len: 2,
        target: [0x0003A1, 0x000313, 0x000000],
    },
    Mapping {
        source: 0x001FE5,
        len: 1,
        target: [0x001FEC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x001FE6,
        len: 2,
        target: [0x0003A5, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FE7,
        len: 3,
        target: [0x0003A5, 0x000308, 0x000342],
    },
    Mapping {
        source: 0x001FF2,
        len: 2,
        target: [0x001FFA, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FF3,
        len: 2,
        target: [0x0003A9, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FF4,
        len: 2,
        target: [0x00038F, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x001FF6,
        len: 2,
        target: [0x0003A9, 0x000342, 0x000000],
    },
    Mapping {
        source: 0x001FF7,
        len: 3,
        target: [0x0003A9, 0x000342, 0x000399],
    },
    Mapping {
        source: 0x001FFC,
        len: 2,
        target: [0x0003A9, 0x000399, 0x000000],
    },
    Mapping {
        source: 0x00214E,
        len: 1,
        target: [0x002132, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002170,
        len: 1,
        target: [0x002160, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002171,
        len: 1,
        target: [0x002161, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002172,
        len: 1,
        target: [0x002162, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002173,
        len: 1,
        target: [0x002163, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002174,
        len: 1,
        target: [0x002164, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002175,
        len: 1,
        target: [0x002165, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002176,
        len: 1,
        target: [0x002166, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002177,
        len: 1,
        target: [0x002167, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002178,
        len: 1,
        target: [0x002168, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002179,
        len: 1,
        target: [0x002169, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00217A,
        len: 1,
        target: [0x00216A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00217B,
        len: 1,
        target: [0x00216B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00217C,
        len: 1,
        target: [0x00216C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00217D,
        len: 1,
        target: [0x00216D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00217E,
        len: 1,
        target: [0x00216E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00217F,
        len: 1,
        target: [0x00216F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002184,
        len: 1,
        target: [0x002183, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D0,
        len: 1,
        target: [0x0024B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D1,
        len: 1,
        target: [0x0024B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D2,
        len: 1,
        target: [0x0024B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D3,
        len: 1,
        target: [0x0024B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D4,
        len: 1,
        target: [0x0024BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D5,
        len: 1,
        target: [0x0024BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D6,
        len: 1,
        target: [0x0024BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D7,
        len: 1,
        target: [0x0024BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D8,
        len: 1,
        target: [0x0024BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024D9,
        len: 1,
        target: [0x0024BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024DA,
        len: 1,
        target: [0x0024C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024DB,
        len: 1,
        target: [0x0024C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024DC,
        len: 1,
        target: [0x0024C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024DD,
        len: 1,
        target: [0x0024C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024DE,
        len: 1,
        target: [0x0024C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024DF,
        len: 1,
        target: [0x0024C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E0,
        len: 1,
        target: [0x0024C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E1,
        len: 1,
        target: [0x0024C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E2,
        len: 1,
        target: [0x0024C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E3,
        len: 1,
        target: [0x0024C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E4,
        len: 1,
        target: [0x0024CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E5,
        len: 1,
        target: [0x0024CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E6,
        len: 1,
        target: [0x0024CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E7,
        len: 1,
        target: [0x0024CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E8,
        len: 1,
        target: [0x0024CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0024E9,
        len: 1,
        target: [0x0024CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C30,
        len: 1,
        target: [0x002C00, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C31,
        len: 1,
        target: [0x002C01, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C32,
        len: 1,
        target: [0x002C02, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C33,
        len: 1,
        target: [0x002C03, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C34,
        len: 1,
        target: [0x002C04, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C35,
        len: 1,
        target: [0x002C05, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C36,
        len: 1,
        target: [0x002C06, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C37,
        len: 1,
        target: [0x002C07, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C38,
        len: 1,
        target: [0x002C08, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C39,
        len: 1,
        target: [0x002C09, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C3A,
        len: 1,
        target: [0x002C0A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C3B,
        len: 1,
        target: [0x002C0B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C3C,
        len: 1,
        target: [0x002C0C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C3D,
        len: 1,
        target: [0x002C0D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C3E,
        len: 1,
        target: [0x002C0E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C3F,
        len: 1,
        target: [0x002C0F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C40,
        len: 1,
        target: [0x002C10, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C41,
        len: 1,
        target: [0x002C11, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C42,
        len: 1,
        target: [0x002C12, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C43,
        len: 1,
        target: [0x002C13, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C44,
        len: 1,
        target: [0x002C14, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C45,
        len: 1,
        target: [0x002C15, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C46,
        len: 1,
        target: [0x002C16, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C47,
        len: 1,
        target: [0x002C17, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C48,
        len: 1,
        target: [0x002C18, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C49,
        len: 1,
        target: [0x002C19, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C4A,
        len: 1,
        target: [0x002C1A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C4B,
        len: 1,
        target: [0x002C1B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C4C,
        len: 1,
        target: [0x002C1C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C4D,
        len: 1,
        target: [0x002C1D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C4E,
        len: 1,
        target: [0x002C1E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C4F,
        len: 1,
        target: [0x002C1F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C50,
        len: 1,
        target: [0x002C20, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C51,
        len: 1,
        target: [0x002C21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C52,
        len: 1,
        target: [0x002C22, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C53,
        len: 1,
        target: [0x002C23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C54,
        len: 1,
        target: [0x002C24, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C55,
        len: 1,
        target: [0x002C25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C56,
        len: 1,
        target: [0x002C26, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C57,
        len: 1,
        target: [0x002C27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C58,
        len: 1,
        target: [0x002C28, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C59,
        len: 1,
        target: [0x002C29, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C5A,
        len: 1,
        target: [0x002C2A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C5B,
        len: 1,
        target: [0x002C2B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C5C,
        len: 1,
        target: [0x002C2C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C5D,
        len: 1,
        target: [0x002C2D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C5E,
        len: 1,
        target: [0x002C2E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C5F,
        len: 1,
        target: [0x002C2F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C61,
        len: 1,
        target: [0x002C60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C65,
        len: 1,
        target: [0x00023A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C66,
        len: 1,
        target: [0x00023E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C68,
        len: 1,
        target: [0x002C67, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6A,
        len: 1,
        target: [0x002C69, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C6C,
        len: 1,
        target: [0x002C6B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C73,
        len: 1,
        target: [0x002C72, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C76,
        len: 1,
        target: [0x002C75, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C81,
        len: 1,
        target: [0x002C80, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C83,
        len: 1,
        target: [0x002C82, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C85,
        len: 1,
        target: [0x002C84, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C87,
        len: 1,
        target: [0x002C86, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C89,
        len: 1,
        target: [0x002C88, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8B,
        len: 1,
        target: [0x002C8A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8D,
        len: 1,
        target: [0x002C8C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C8F,
        len: 1,
        target: [0x002C8E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C91,
        len: 1,
        target: [0x002C90, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C93,
        len: 1,
        target: [0x002C92, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C95,
        len: 1,
        target: [0x002C94, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C97,
        len: 1,
        target: [0x002C96, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C99,
        len: 1,
        target: [0x002C98, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9B,
        len: 1,
        target: [0x002C9A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9D,
        len: 1,
        target: [0x002C9C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002C9F,
        len: 1,
        target: [0x002C9E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA1,
        len: 1,
        target: [0x002CA0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA3,
        len: 1,
        target: [0x002CA2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA5,
        len: 1,
        target: [0x002CA4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA7,
        len: 1,
        target: [0x002CA6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CA9,
        len: 1,
        target: [0x002CA8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAB,
        len: 1,
        target: [0x002CAA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAD,
        len: 1,
        target: [0x002CAC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CAF,
        len: 1,
        target: [0x002CAE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB1,
        len: 1,
        target: [0x002CB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB3,
        len: 1,
        target: [0x002CB2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB5,
        len: 1,
        target: [0x002CB4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB7,
        len: 1,
        target: [0x002CB6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CB9,
        len: 1,
        target: [0x002CB8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBB,
        len: 1,
        target: [0x002CBA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBD,
        len: 1,
        target: [0x002CBC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CBF,
        len: 1,
        target: [0x002CBE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC1,
        len: 1,
        target: [0x002CC0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC3,
        len: 1,
        target: [0x002CC2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC5,
        len: 1,
        target: [0x002CC4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC7,
        len: 1,
        target: [0x002CC6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CC9,
        len: 1,
        target: [0x002CC8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCB,
        len: 1,
        target: [0x002CCA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCD,
        len: 1,
        target: [0x002CCC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CCF,
        len: 1,
        target: [0x002CCE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD1,
        len: 1,
        target: [0x002CD0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD3,
        len: 1,
        target: [0x002CD2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD5,
        len: 1,
        target: [0x002CD4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD7,
        len: 1,
        target: [0x002CD6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CD9,
        len: 1,
        target: [0x002CD8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDB,
        len: 1,
        target: [0x002CDA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDD,
        len: 1,
        target: [0x002CDC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CDF,
        len: 1,
        target: [0x002CDE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CE1,
        len: 1,
        target: [0x002CE0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CE3,
        len: 1,
        target: [0x002CE2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CEC,
        len: 1,
        target: [0x002CEB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CEE,
        len: 1,
        target: [0x002CED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002CF3,
        len: 1,
        target: [0x002CF2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D00,
        len: 1,
        target: [0x0010A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D01,
        len: 1,
        target: [0x0010A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D02,
        len: 1,
        target: [0x0010A2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D03,
        len: 1,
        target: [0x0010A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D04,
        len: 1,
        target: [0x0010A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D05,
        len: 1,
        target: [0x0010A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D06,
        len: 1,
        target: [0x0010A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D07,
        len: 1,
        target: [0x0010A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D08,
        len: 1,
        target: [0x0010A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D09,
        len: 1,
        target: [0x0010A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D0A,
        len: 1,
        target: [0x0010AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D0B,
        len: 1,
        target: [0x0010AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D0C,
        len: 1,
        target: [0x0010AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D0D,
        len: 1,
        target: [0x0010AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D0E,
        len: 1,
        target: [0x0010AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D0F,
        len: 1,
        target: [0x0010AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D10,
        len: 1,
        target: [0x0010B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D11,
        len: 1,
        target: [0x0010B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D12,
        len: 1,
        target: [0x0010B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D13,
        len: 1,
        target: [0x0010B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D14,
        len: 1,
        target: [0x0010B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D15,
        len: 1,
        target: [0x0010B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D16,
        len: 1,
        target: [0x0010B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D17,
        len: 1,
        target: [0x0010B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D18,
        len: 1,
        target: [0x0010B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D19,
        len: 1,
        target: [0x0010B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D1A,
        len: 1,
        target: [0x0010BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D1B,
        len: 1,
        target: [0x0010BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D1C,
        len: 1,
        target: [0x0010BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D1D,
        len: 1,
        target: [0x0010BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D1E,
        len: 1,
        target: [0x0010BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D1F,
        len: 1,
        target: [0x0010BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D20,
        len: 1,
        target: [0x0010C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D21,
        len: 1,
        target: [0x0010C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D22,
        len: 1,
        target: [0x0010C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D23,
        len: 1,
        target: [0x0010C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D24,
        len: 1,
        target: [0x0010C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D25,
        len: 1,
        target: [0x0010C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D27,
        len: 1,
        target: [0x0010C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x002D2D,
        len: 1,
        target: [0x0010CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A641,
        len: 1,
        target: [0x00A640, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A643,
        len: 1,
        target: [0x00A642, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A645,
        len: 1,
        target: [0x00A644, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A647,
        len: 1,
        target: [0x00A646, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A649,
        len: 1,
        target: [0x00A648, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64B,
        len: 1,
        target: [0x00A64A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64D,
        len: 1,
        target: [0x00A64C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A64F,
        len: 1,
        target: [0x00A64E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A651,
        len: 1,
        target: [0x00A650, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A653,
        len: 1,
        target: [0x00A652, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A655,
        len: 1,
        target: [0x00A654, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A657,
        len: 1,
        target: [0x00A656, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A659,
        len: 1,
        target: [0x00A658, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65B,
        len: 1,
        target: [0x00A65A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65D,
        len: 1,
        target: [0x00A65C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A65F,
        len: 1,
        target: [0x00A65E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A661,
        len: 1,
        target: [0x00A660, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A663,
        len: 1,
        target: [0x00A662, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A665,
        len: 1,
        target: [0x00A664, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A667,
        len: 1,
        target: [0x00A666, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A669,
        len: 1,
        target: [0x00A668, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A66B,
        len: 1,
        target: [0x00A66A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A66D,
        len: 1,
        target: [0x00A66C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A681,
        len: 1,
        target: [0x00A680, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A683,
        len: 1,
        target: [0x00A682, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A685,
        len: 1,
        target: [0x00A684, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A687,
        len: 1,
        target: [0x00A686, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A689,
        len: 1,
        target: [0x00A688, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68B,
        len: 1,
        target: [0x00A68A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68D,
        len: 1,
        target: [0x00A68C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A68F,
        len: 1,
        target: [0x00A68E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A691,
        len: 1,
        target: [0x00A690, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A693,
        len: 1,
        target: [0x00A692, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A695,
        len: 1,
        target: [0x00A694, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A697,
        len: 1,
        target: [0x00A696, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A699,
        len: 1,
        target: [0x00A698, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A69B,
        len: 1,
        target: [0x00A69A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A723,
        len: 1,
        target: [0x00A722, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A725,
        len: 1,
        target: [0x00A724, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A727,
        len: 1,
        target: [0x00A726, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A729,
        len: 1,
        target: [0x00A728, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72B,
        len: 1,
        target: [0x00A72A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72D,
        len: 1,
        target: [0x00A72C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A72F,
        len: 1,
        target: [0x00A72E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A733,
        len: 1,
        target: [0x00A732, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A735,
        len: 1,
        target: [0x00A734, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A737,
        len: 1,
        target: [0x00A736, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A739,
        len: 1,
        target: [0x00A738, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73B,
        len: 1,
        target: [0x00A73A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73D,
        len: 1,
        target: [0x00A73C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A73F,
        len: 1,
        target: [0x00A73E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A741,
        len: 1,
        target: [0x00A740, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A743,
        len: 1,
        target: [0x00A742, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A745,
        len: 1,
        target: [0x00A744, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A747,
        len: 1,
        target: [0x00A746, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A749,
        len: 1,
        target: [0x00A748, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74B,
        len: 1,
        target: [0x00A74A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74D,
        len: 1,
        target: [0x00A74C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A74F,
        len: 1,
        target: [0x00A74E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A751,
        len: 1,
        target: [0x00A750, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A753,
        len: 1,
        target: [0x00A752, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A755,
        len: 1,
        target: [0x00A754, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A757,
        len: 1,
        target: [0x00A756, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A759,
        len: 1,
        target: [0x00A758, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75B,
        len: 1,
        target: [0x00A75A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75D,
        len: 1,
        target: [0x00A75C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A75F,
        len: 1,
        target: [0x00A75E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A761,
        len: 1,
        target: [0x00A760, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A763,
        len: 1,
        target: [0x00A762, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A765,
        len: 1,
        target: [0x00A764, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A767,
        len: 1,
        target: [0x00A766, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A769,
        len: 1,
        target: [0x00A768, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76B,
        len: 1,
        target: [0x00A76A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76D,
        len: 1,
        target: [0x00A76C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A76F,
        len: 1,
        target: [0x00A76E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77A,
        len: 1,
        target: [0x00A779, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77C,
        len: 1,
        target: [0x00A77B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A77F,
        len: 1,
        target: [0x00A77E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A781,
        len: 1,
        target: [0x00A780, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A783,
        len: 1,
        target: [0x00A782, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A785,
        len: 1,
        target: [0x00A784, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A787,
        len: 1,
        target: [0x00A786, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A78C,
        len: 1,
        target: [0x00A78B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A791,
        len: 1,
        target: [0x00A790, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A793,
        len: 1,
        target: [0x00A792, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A794,
        len: 1,
        target: [0x00A7C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A797,
        len: 1,
        target: [0x00A796, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A799,
        len: 1,
        target: [0x00A798, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79B,
        len: 1,
        target: [0x00A79A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79D,
        len: 1,
        target: [0x00A79C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A79F,
        len: 1,
        target: [0x00A79E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A1,
        len: 1,
        target: [0x00A7A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A3,
        len: 1,
        target: [0x00A7A2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A5,
        len: 1,
        target: [0x00A7A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A7,
        len: 1,
        target: [0x00A7A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7A9,
        len: 1,
        target: [0x00A7A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B5,
        len: 1,
        target: [0x00A7B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B7,
        len: 1,
        target: [0x00A7B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7B9,
        len: 1,
        target: [0x00A7B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BB,
        len: 1,
        target: [0x00A7BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BD,
        len: 1,
        target: [0x00A7BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7BF,
        len: 1,
        target: [0x00A7BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C1,
        len: 1,
        target: [0x00A7C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C3,
        len: 1,
        target: [0x00A7C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7C8,
        len: 1,
        target: [0x00A7C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CA,
        len: 1,
        target: [0x00A7C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CD,
        len: 1,
        target: [0x00A7CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7CF,
        len: 1,
        target: [0x00A7CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D1,
        len: 1,
        target: [0x00A7D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D3,
        len: 1,
        target: [0x00A7D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D5,
        len: 1,
        target: [0x00A7D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D7,
        len: 1,
        target: [0x00A7D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7D9,
        len: 1,
        target: [0x00A7D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7DB,
        len: 1,
        target: [0x00A7DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00A7F6,
        len: 1,
        target: [0x00A7F5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB53,
        len: 1,
        target: [0x00A7B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB70,
        len: 1,
        target: [0x0013A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB71,
        len: 1,
        target: [0x0013A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB72,
        len: 1,
        target: [0x0013A2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB73,
        len: 1,
        target: [0x0013A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB74,
        len: 1,
        target: [0x0013A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB75,
        len: 1,
        target: [0x0013A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB76,
        len: 1,
        target: [0x0013A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB77,
        len: 1,
        target: [0x0013A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB78,
        len: 1,
        target: [0x0013A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB79,
        len: 1,
        target: [0x0013A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7A,
        len: 1,
        target: [0x0013AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7B,
        len: 1,
        target: [0x0013AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7C,
        len: 1,
        target: [0x0013AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7D,
        len: 1,
        target: [0x0013AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7E,
        len: 1,
        target: [0x0013AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB7F,
        len: 1,
        target: [0x0013AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB80,
        len: 1,
        target: [0x0013B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB81,
        len: 1,
        target: [0x0013B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB82,
        len: 1,
        target: [0x0013B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB83,
        len: 1,
        target: [0x0013B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB84,
        len: 1,
        target: [0x0013B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB85,
        len: 1,
        target: [0x0013B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB86,
        len: 1,
        target: [0x0013B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB87,
        len: 1,
        target: [0x0013B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB88,
        len: 1,
        target: [0x0013B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB89,
        len: 1,
        target: [0x0013B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8A,
        len: 1,
        target: [0x0013BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8B,
        len: 1,
        target: [0x0013BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8C,
        len: 1,
        target: [0x0013BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8D,
        len: 1,
        target: [0x0013BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8E,
        len: 1,
        target: [0x0013BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB8F,
        len: 1,
        target: [0x0013BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB90,
        len: 1,
        target: [0x0013C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB91,
        len: 1,
        target: [0x0013C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB92,
        len: 1,
        target: [0x0013C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB93,
        len: 1,
        target: [0x0013C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB94,
        len: 1,
        target: [0x0013C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB95,
        len: 1,
        target: [0x0013C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB96,
        len: 1,
        target: [0x0013C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB97,
        len: 1,
        target: [0x0013C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB98,
        len: 1,
        target: [0x0013C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB99,
        len: 1,
        target: [0x0013C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9A,
        len: 1,
        target: [0x0013CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9B,
        len: 1,
        target: [0x0013CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9C,
        len: 1,
        target: [0x0013CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9D,
        len: 1,
        target: [0x0013CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9E,
        len: 1,
        target: [0x0013CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00AB9F,
        len: 1,
        target: [0x0013CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA0,
        len: 1,
        target: [0x0013D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA1,
        len: 1,
        target: [0x0013D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA2,
        len: 1,
        target: [0x0013D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA3,
        len: 1,
        target: [0x0013D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA4,
        len: 1,
        target: [0x0013D4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA5,
        len: 1,
        target: [0x0013D5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA6,
        len: 1,
        target: [0x0013D6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA7,
        len: 1,
        target: [0x0013D7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA8,
        len: 1,
        target: [0x0013D8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABA9,
        len: 1,
        target: [0x0013D9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAA,
        len: 1,
        target: [0x0013DA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAB,
        len: 1,
        target: [0x0013DB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAC,
        len: 1,
        target: [0x0013DC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAD,
        len: 1,
        target: [0x0013DD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAE,
        len: 1,
        target: [0x0013DE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABAF,
        len: 1,
        target: [0x0013DF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB0,
        len: 1,
        target: [0x0013E0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB1,
        len: 1,
        target: [0x0013E1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB2,
        len: 1,
        target: [0x0013E2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB3,
        len: 1,
        target: [0x0013E3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB4,
        len: 1,
        target: [0x0013E4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB5,
        len: 1,
        target: [0x0013E5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB6,
        len: 1,
        target: [0x0013E6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB7,
        len: 1,
        target: [0x0013E7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB8,
        len: 1,
        target: [0x0013E8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABB9,
        len: 1,
        target: [0x0013E9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBA,
        len: 1,
        target: [0x0013EA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBB,
        len: 1,
        target: [0x0013EB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBC,
        len: 1,
        target: [0x0013EC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBD,
        len: 1,
        target: [0x0013ED, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBE,
        len: 1,
        target: [0x0013EE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00ABBF,
        len: 1,
        target: [0x0013EF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FB00,
        len: 2,
        target: [0x000046, 0x000046, 0x000000],
    },
    Mapping {
        source: 0x00FB01,
        len: 2,
        target: [0x000046, 0x000049, 0x000000],
    },
    Mapping {
        source: 0x00FB02,
        len: 2,
        target: [0x000046, 0x00004C, 0x000000],
    },
    Mapping {
        source: 0x00FB03,
        len: 3,
        target: [0x000046, 0x000046, 0x000049],
    },
    Mapping {
        source: 0x00FB04,
        len: 3,
        target: [0x000046, 0x000046, 0x00004C],
    },
    Mapping {
        source: 0x00FB05,
        len: 2,
        target: [0x000053, 0x000054, 0x000000],
    },
    Mapping {
        source: 0x00FB06,
        len: 2,
        target: [0x000053, 0x000054, 0x000000],
    },
    Mapping {
        source: 0x00FB13,
        len: 2,
        target: [0x000544, 0x000546, 0x000000],
    },
    Mapping {
        source: 0x00FB14,
        len: 2,
        target: [0x000544, 0x000535, 0x000000],
    },
    Mapping {
        source: 0x00FB15,
        len: 2,
        target: [0x000544, 0x00053B, 0x000000],
    },
    Mapping {
        source: 0x00FB16,
        len: 2,
        target: [0x00054E, 0x000546, 0x000000],
    },
    Mapping {
        source: 0x00FB17,
        len: 2,
        target: [0x000544, 0x00053D, 0x000000],
    },
    Mapping {
        source: 0x00FF41,
        len: 1,
        target: [0x00FF21, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF42,
        len: 1,
        target: [0x00FF22, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF43,
        len: 1,
        target: [0x00FF23, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF44,
        len: 1,
        target: [0x00FF24, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF45,
        len: 1,
        target: [0x00FF25, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF46,
        len: 1,
        target: [0x00FF26, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF47,
        len: 1,
        target: [0x00FF27, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF48,
        len: 1,
        target: [0x00FF28, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF49,
        len: 1,
        target: [0x00FF29, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF4A,
        len: 1,
        target: [0x00FF2A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF4B,
        len: 1,
        target: [0x00FF2B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF4C,
        len: 1,
        target: [0x00FF2C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF4D,
        len: 1,
        target: [0x00FF2D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF4E,
        len: 1,
        target: [0x00FF2E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF4F,
        len: 1,
        target: [0x00FF2F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF50,
        len: 1,
        target: [0x00FF30, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF51,
        len: 1,
        target: [0x00FF31, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF52,
        len: 1,
        target: [0x00FF32, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF53,
        len: 1,
        target: [0x00FF33, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF54,
        len: 1,
        target: [0x00FF34, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF55,
        len: 1,
        target: [0x00FF35, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF56,
        len: 1,
        target: [0x00FF36, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF57,
        len: 1,
        target: [0x00FF37, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF58,
        len: 1,
        target: [0x00FF38, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF59,
        len: 1,
        target: [0x00FF39, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x00FF5A,
        len: 1,
        target: [0x00FF3A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010428,
        len: 1,
        target: [0x010400, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010429,
        len: 1,
        target: [0x010401, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01042A,
        len: 1,
        target: [0x010402, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01042B,
        len: 1,
        target: [0x010403, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01042C,
        len: 1,
        target: [0x010404, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01042D,
        len: 1,
        target: [0x010405, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01042E,
        len: 1,
        target: [0x010406, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01042F,
        len: 1,
        target: [0x010407, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010430,
        len: 1,
        target: [0x010408, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010431,
        len: 1,
        target: [0x010409, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010432,
        len: 1,
        target: [0x01040A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010433,
        len: 1,
        target: [0x01040B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010434,
        len: 1,
        target: [0x01040C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010435,
        len: 1,
        target: [0x01040D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010436,
        len: 1,
        target: [0x01040E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010437,
        len: 1,
        target: [0x01040F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010438,
        len: 1,
        target: [0x010410, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010439,
        len: 1,
        target: [0x010411, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01043A,
        len: 1,
        target: [0x010412, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01043B,
        len: 1,
        target: [0x010413, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01043C,
        len: 1,
        target: [0x010414, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01043D,
        len: 1,
        target: [0x010415, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01043E,
        len: 1,
        target: [0x010416, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01043F,
        len: 1,
        target: [0x010417, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010440,
        len: 1,
        target: [0x010418, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010441,
        len: 1,
        target: [0x010419, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010442,
        len: 1,
        target: [0x01041A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010443,
        len: 1,
        target: [0x01041B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010444,
        len: 1,
        target: [0x01041C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010445,
        len: 1,
        target: [0x01041D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010446,
        len: 1,
        target: [0x01041E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010447,
        len: 1,
        target: [0x01041F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010448,
        len: 1,
        target: [0x010420, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010449,
        len: 1,
        target: [0x010421, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01044A,
        len: 1,
        target: [0x010422, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01044B,
        len: 1,
        target: [0x010423, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01044C,
        len: 1,
        target: [0x010424, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01044D,
        len: 1,
        target: [0x010425, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01044E,
        len: 1,
        target: [0x010426, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01044F,
        len: 1,
        target: [0x010427, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D8,
        len: 1,
        target: [0x0104B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104D9,
        len: 1,
        target: [0x0104B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104DA,
        len: 1,
        target: [0x0104B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104DB,
        len: 1,
        target: [0x0104B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104DC,
        len: 1,
        target: [0x0104B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104DD,
        len: 1,
        target: [0x0104B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104DE,
        len: 1,
        target: [0x0104B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104DF,
        len: 1,
        target: [0x0104B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E0,
        len: 1,
        target: [0x0104B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E1,
        len: 1,
        target: [0x0104B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E2,
        len: 1,
        target: [0x0104BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E3,
        len: 1,
        target: [0x0104BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E4,
        len: 1,
        target: [0x0104BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E5,
        len: 1,
        target: [0x0104BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E6,
        len: 1,
        target: [0x0104BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E7,
        len: 1,
        target: [0x0104BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E8,
        len: 1,
        target: [0x0104C0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104E9,
        len: 1,
        target: [0x0104C1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104EA,
        len: 1,
        target: [0x0104C2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104EB,
        len: 1,
        target: [0x0104C3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104EC,
        len: 1,
        target: [0x0104C4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104ED,
        len: 1,
        target: [0x0104C5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104EE,
        len: 1,
        target: [0x0104C6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104EF,
        len: 1,
        target: [0x0104C7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F0,
        len: 1,
        target: [0x0104C8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F1,
        len: 1,
        target: [0x0104C9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F2,
        len: 1,
        target: [0x0104CA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F3,
        len: 1,
        target: [0x0104CB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F4,
        len: 1,
        target: [0x0104CC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F5,
        len: 1,
        target: [0x0104CD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F6,
        len: 1,
        target: [0x0104CE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F7,
        len: 1,
        target: [0x0104CF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F8,
        len: 1,
        target: [0x0104D0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104F9,
        len: 1,
        target: [0x0104D1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104FA,
        len: 1,
        target: [0x0104D2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0104FB,
        len: 1,
        target: [0x0104D3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010597,
        len: 1,
        target: [0x010570, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010598,
        len: 1,
        target: [0x010571, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010599,
        len: 1,
        target: [0x010572, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01059A,
        len: 1,
        target: [0x010573, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01059B,
        len: 1,
        target: [0x010574, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01059C,
        len: 1,
        target: [0x010575, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01059D,
        len: 1,
        target: [0x010576, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01059E,
        len: 1,
        target: [0x010577, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01059F,
        len: 1,
        target: [0x010578, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A0,
        len: 1,
        target: [0x010579, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A1,
        len: 1,
        target: [0x01057A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A3,
        len: 1,
        target: [0x01057C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A4,
        len: 1,
        target: [0x01057D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A5,
        len: 1,
        target: [0x01057E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A6,
        len: 1,
        target: [0x01057F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A7,
        len: 1,
        target: [0x010580, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A8,
        len: 1,
        target: [0x010581, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105A9,
        len: 1,
        target: [0x010582, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105AA,
        len: 1,
        target: [0x010583, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105AB,
        len: 1,
        target: [0x010584, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105AC,
        len: 1,
        target: [0x010585, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105AD,
        len: 1,
        target: [0x010586, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105AE,
        len: 1,
        target: [0x010587, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105AF,
        len: 1,
        target: [0x010588, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B0,
        len: 1,
        target: [0x010589, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B1,
        len: 1,
        target: [0x01058A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B3,
        len: 1,
        target: [0x01058C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B4,
        len: 1,
        target: [0x01058D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B5,
        len: 1,
        target: [0x01058E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B6,
        len: 1,
        target: [0x01058F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B7,
        len: 1,
        target: [0x010590, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B8,
        len: 1,
        target: [0x010591, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105B9,
        len: 1,
        target: [0x010592, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105BB,
        len: 1,
        target: [0x010594, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0105BC,
        len: 1,
        target: [0x010595, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC0,
        len: 1,
        target: [0x010C80, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC1,
        len: 1,
        target: [0x010C81, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC2,
        len: 1,
        target: [0x010C82, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC3,
        len: 1,
        target: [0x010C83, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC4,
        len: 1,
        target: [0x010C84, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC5,
        len: 1,
        target: [0x010C85, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC6,
        len: 1,
        target: [0x010C86, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC7,
        len: 1,
        target: [0x010C87, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC8,
        len: 1,
        target: [0x010C88, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CC9,
        len: 1,
        target: [0x010C89, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CCA,
        len: 1,
        target: [0x010C8A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CCB,
        len: 1,
        target: [0x010C8B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CCC,
        len: 1,
        target: [0x010C8C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CCD,
        len: 1,
        target: [0x010C8D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CCE,
        len: 1,
        target: [0x010C8E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CCF,
        len: 1,
        target: [0x010C8F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD0,
        len: 1,
        target: [0x010C90, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD1,
        len: 1,
        target: [0x010C91, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD2,
        len: 1,
        target: [0x010C92, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD3,
        len: 1,
        target: [0x010C93, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD4,
        len: 1,
        target: [0x010C94, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD5,
        len: 1,
        target: [0x010C95, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD6,
        len: 1,
        target: [0x010C96, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD7,
        len: 1,
        target: [0x010C97, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD8,
        len: 1,
        target: [0x010C98, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CD9,
        len: 1,
        target: [0x010C99, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CDA,
        len: 1,
        target: [0x010C9A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CDB,
        len: 1,
        target: [0x010C9B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CDC,
        len: 1,
        target: [0x010C9C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CDD,
        len: 1,
        target: [0x010C9D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CDE,
        len: 1,
        target: [0x010C9E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CDF,
        len: 1,
        target: [0x010C9F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE0,
        len: 1,
        target: [0x010CA0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE1,
        len: 1,
        target: [0x010CA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE2,
        len: 1,
        target: [0x010CA2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE3,
        len: 1,
        target: [0x010CA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE4,
        len: 1,
        target: [0x010CA4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE5,
        len: 1,
        target: [0x010CA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE6,
        len: 1,
        target: [0x010CA6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE7,
        len: 1,
        target: [0x010CA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE8,
        len: 1,
        target: [0x010CA8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CE9,
        len: 1,
        target: [0x010CA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CEA,
        len: 1,
        target: [0x010CAA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CEB,
        len: 1,
        target: [0x010CAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CEC,
        len: 1,
        target: [0x010CAC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CED,
        len: 1,
        target: [0x010CAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CEE,
        len: 1,
        target: [0x010CAE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CEF,
        len: 1,
        target: [0x010CAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CF0,
        len: 1,
        target: [0x010CB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CF1,
        len: 1,
        target: [0x010CB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010CF2,
        len: 1,
        target: [0x010CB2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D70,
        len: 1,
        target: [0x010D50, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D71,
        len: 1,
        target: [0x010D51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D72,
        len: 1,
        target: [0x010D52, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D73,
        len: 1,
        target: [0x010D53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D74,
        len: 1,
        target: [0x010D54, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D75,
        len: 1,
        target: [0x010D55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D76,
        len: 1,
        target: [0x010D56, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D77,
        len: 1,
        target: [0x010D57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D78,
        len: 1,
        target: [0x010D58, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D79,
        len: 1,
        target: [0x010D59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D7A,
        len: 1,
        target: [0x010D5A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D7B,
        len: 1,
        target: [0x010D5B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D7C,
        len: 1,
        target: [0x010D5C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D7D,
        len: 1,
        target: [0x010D5D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D7E,
        len: 1,
        target: [0x010D5E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D7F,
        len: 1,
        target: [0x010D5F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D80,
        len: 1,
        target: [0x010D60, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D81,
        len: 1,
        target: [0x010D61, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D82,
        len: 1,
        target: [0x010D62, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D83,
        len: 1,
        target: [0x010D63, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D84,
        len: 1,
        target: [0x010D64, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x010D85,
        len: 1,
        target: [0x010D65, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C0,
        len: 1,
        target: [0x0118A0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C1,
        len: 1,
        target: [0x0118A1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C2,
        len: 1,
        target: [0x0118A2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C3,
        len: 1,
        target: [0x0118A3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C4,
        len: 1,
        target: [0x0118A4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C5,
        len: 1,
        target: [0x0118A5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C6,
        len: 1,
        target: [0x0118A6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C7,
        len: 1,
        target: [0x0118A7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C8,
        len: 1,
        target: [0x0118A8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118C9,
        len: 1,
        target: [0x0118A9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118CA,
        len: 1,
        target: [0x0118AA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118CB,
        len: 1,
        target: [0x0118AB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118CC,
        len: 1,
        target: [0x0118AC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118CD,
        len: 1,
        target: [0x0118AD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118CE,
        len: 1,
        target: [0x0118AE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118CF,
        len: 1,
        target: [0x0118AF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D0,
        len: 1,
        target: [0x0118B0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D1,
        len: 1,
        target: [0x0118B1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D2,
        len: 1,
        target: [0x0118B2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D3,
        len: 1,
        target: [0x0118B3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D4,
        len: 1,
        target: [0x0118B4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D5,
        len: 1,
        target: [0x0118B5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D6,
        len: 1,
        target: [0x0118B6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D7,
        len: 1,
        target: [0x0118B7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D8,
        len: 1,
        target: [0x0118B8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118D9,
        len: 1,
        target: [0x0118B9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118DA,
        len: 1,
        target: [0x0118BA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118DB,
        len: 1,
        target: [0x0118BB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118DC,
        len: 1,
        target: [0x0118BC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118DD,
        len: 1,
        target: [0x0118BD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118DE,
        len: 1,
        target: [0x0118BE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x0118DF,
        len: 1,
        target: [0x0118BF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E60,
        len: 1,
        target: [0x016E40, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E61,
        len: 1,
        target: [0x016E41, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E62,
        len: 1,
        target: [0x016E42, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E63,
        len: 1,
        target: [0x016E43, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E64,
        len: 1,
        target: [0x016E44, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E65,
        len: 1,
        target: [0x016E45, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E66,
        len: 1,
        target: [0x016E46, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E67,
        len: 1,
        target: [0x016E47, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E68,
        len: 1,
        target: [0x016E48, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E69,
        len: 1,
        target: [0x016E49, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E6A,
        len: 1,
        target: [0x016E4A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E6B,
        len: 1,
        target: [0x016E4B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E6C,
        len: 1,
        target: [0x016E4C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E6D,
        len: 1,
        target: [0x016E4D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E6E,
        len: 1,
        target: [0x016E4E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E6F,
        len: 1,
        target: [0x016E4F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E70,
        len: 1,
        target: [0x016E50, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E71,
        len: 1,
        target: [0x016E51, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E72,
        len: 1,
        target: [0x016E52, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E73,
        len: 1,
        target: [0x016E53, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E74,
        len: 1,
        target: [0x016E54, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E75,
        len: 1,
        target: [0x016E55, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E76,
        len: 1,
        target: [0x016E56, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E77,
        len: 1,
        target: [0x016E57, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E78,
        len: 1,
        target: [0x016E58, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E79,
        len: 1,
        target: [0x016E59, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E7A,
        len: 1,
        target: [0x016E5A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E7B,
        len: 1,
        target: [0x016E5B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E7C,
        len: 1,
        target: [0x016E5C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E7D,
        len: 1,
        target: [0x016E5D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E7E,
        len: 1,
        target: [0x016E5E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016E7F,
        len: 1,
        target: [0x016E5F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EBB,
        len: 1,
        target: [0x016EA0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EBC,
        len: 1,
        target: [0x016EA1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EBD,
        len: 1,
        target: [0x016EA2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EBE,
        len: 1,
        target: [0x016EA3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EBF,
        len: 1,
        target: [0x016EA4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC0,
        len: 1,
        target: [0x016EA5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC1,
        len: 1,
        target: [0x016EA6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC2,
        len: 1,
        target: [0x016EA7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC3,
        len: 1,
        target: [0x016EA8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC4,
        len: 1,
        target: [0x016EA9, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC5,
        len: 1,
        target: [0x016EAA, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC6,
        len: 1,
        target: [0x016EAB, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC7,
        len: 1,
        target: [0x016EAC, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC8,
        len: 1,
        target: [0x016EAD, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016EC9,
        len: 1,
        target: [0x016EAE, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ECA,
        len: 1,
        target: [0x016EAF, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ECB,
        len: 1,
        target: [0x016EB0, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ECC,
        len: 1,
        target: [0x016EB1, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ECD,
        len: 1,
        target: [0x016EB2, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ECE,
        len: 1,
        target: [0x016EB3, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ECF,
        len: 1,
        target: [0x016EB4, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ED0,
        len: 1,
        target: [0x016EB5, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ED1,
        len: 1,
        target: [0x016EB6, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ED2,
        len: 1,
        target: [0x016EB7, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x016ED3,
        len: 1,
        target: [0x016EB8, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E922,
        len: 1,
        target: [0x01E900, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E923,
        len: 1,
        target: [0x01E901, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E924,
        len: 1,
        target: [0x01E902, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E925,
        len: 1,
        target: [0x01E903, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E926,
        len: 1,
        target: [0x01E904, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E927,
        len: 1,
        target: [0x01E905, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E928,
        len: 1,
        target: [0x01E906, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E929,
        len: 1,
        target: [0x01E907, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E92A,
        len: 1,
        target: [0x01E908, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E92B,
        len: 1,
        target: [0x01E909, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E92C,
        len: 1,
        target: [0x01E90A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E92D,
        len: 1,
        target: [0x01E90B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E92E,
        len: 1,
        target: [0x01E90C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E92F,
        len: 1,
        target: [0x01E90D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E930,
        len: 1,
        target: [0x01E90E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E931,
        len: 1,
        target: [0x01E90F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E932,
        len: 1,
        target: [0x01E910, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E933,
        len: 1,
        target: [0x01E911, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E934,
        len: 1,
        target: [0x01E912, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E935,
        len: 1,
        target: [0x01E913, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E936,
        len: 1,
        target: [0x01E914, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E937,
        len: 1,
        target: [0x01E915, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E938,
        len: 1,
        target: [0x01E916, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E939,
        len: 1,
        target: [0x01E917, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E93A,
        len: 1,
        target: [0x01E918, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E93B,
        len: 1,
        target: [0x01E919, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E93C,
        len: 1,
        target: [0x01E91A, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E93D,
        len: 1,
        target: [0x01E91B, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E93E,
        len: 1,
        target: [0x01E91C, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E93F,
        len: 1,
        target: [0x01E91D, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E940,
        len: 1,
        target: [0x01E91E, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E941,
        len: 1,
        target: [0x01E91F, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E942,
        len: 1,
        target: [0x01E920, 0x000000, 0x000000],
    },
    Mapping {
        source: 0x01E943,
        len: 1,
        target: [0x01E921, 0x000000, 0x000000],
    },
];

#[inline]
fn contains(ranges: &[Range], code_point: u32) -> bool {
    let mut low = 0_usize;
    let mut high = ranges.len();
    while low < high {
        let middle = low + (high - low) / 2;
        let range = ranges[middle];
        if code_point < range.start {
            high = middle;
        } else if code_point > range.end {
            low = middle + 1;
        } else {
            return true;
        }
    }
    false
}

#[inline]
pub(super) fn is_clean_removed(value: char) -> bool {
    contains(CLEAN_REMOVED, value as u32)
}

#[inline]
pub(super) fn is_letter(value: char) -> bool {
    contains(LETTER, value as u32)
}

#[inline]
pub(super) fn is_cased(value: char) -> bool {
    contains(CASED, value as u32)
}

#[inline]
pub(super) fn is_case_ignorable(value: char) -> bool {
    contains(CASE_IGNORABLE, value as u32)
}

#[inline]
fn find_mapping(table: &[Mapping], code_point: u32) -> Option<Mapping> {
    let mut low = 0_usize;
    let mut high = table.len();
    while low < high {
        let middle = low + (high - low) / 2;
        let mapping = table[middle];
        if code_point < mapping.source {
            high = middle;
        } else if code_point > mapping.source {
            low = middle + 1;
        } else {
            return Some(mapping);
        }
    }
    None
}

#[derive(Clone, Copy)]
pub(super) struct CaseMapping {
    chars: [char; 3],
    len: u8,
}

impl CaseMapping {
    #[inline]
    pub(super) fn iter(self) -> impl Iterator<Item = char> {
        self.chars.into_iter().take(usize::from(self.len))
    }
}

#[inline]
fn map_character(table: &[Mapping], value: char) -> CaseMapping {
    if let Some(mapping) = find_mapping(table, value as u32) {
        CaseMapping {
            chars: [
                char::from_u32(mapping.target[0]).expect("generated Unicode scalar"),
                char::from_u32(mapping.target[1]).expect("generated Unicode scalar"),
                char::from_u32(mapping.target[2]).expect("generated Unicode scalar"),
            ],
            len: mapping.len,
        }
    } else {
        CaseMapping {
            chars: [value, value, value],
            len: 1,
        }
    }
}

#[inline]
pub(super) fn case_fold(value: char) -> CaseMapping {
    map_character(CASE_FOLD, value)
}

#[inline]
pub(super) fn lowercase(value: char) -> CaseMapping {
    map_character(LOWERCASE, value)
}

#[inline]
pub(super) fn uppercase(value: char) -> CaseMapping {
    map_character(UPPERCASE, value)
}

#[cfg(test)]
mod tests {
    use super::{
        CASE_FOLD, CASE_IGNORABLE, CASED, CLEAN_REMOVED, LETTER, LOWERCASE, UPPERCASE, contains,
        find_mapping,
    };

    fn assert_table_boundaries(table: &[super::Range]) {
        let mut previous_end = None;
        for range in table {
            assert!(range.start <= range.end);
            assert!(range.end < 0xD800 || range.start > 0xDFFF);
            if let Some(end) = previous_end {
                assert!(range.start > end);
            }
            assert!(contains(table, range.start));
            assert!(contains(table, range.end));
            previous_end = Some(range.end);
        }
    }

    #[test]
    fn generated_tables_are_sorted_scalar_ranges() {
        assert_table_boundaries(CLEAN_REMOVED);
        assert_table_boundaries(LETTER);
        assert_table_boundaries(CASED);
        assert_table_boundaries(CASE_IGNORABLE);
    }

    #[test]
    fn unicode_boundary_classes_are_pinned() {
        assert!(super::is_clean_removed('\0'));
        assert!(super::is_clean_removed('\u{0378}'));
        assert!(!super::is_clean_removed(' '));
        assert!(!super::is_clean_removed('\u{E000}'));
        assert!(super::is_letter('A'));
        assert!(super::is_letter('\u{4E00}'));
        assert!(!super::is_letter('\u{0301}'));
        assert!(super::is_cased('\u{03A3}'));
        assert!(!super::is_cased('0'));
        assert!(super::is_case_ignorable('\u{0301}'));
        assert!(super::is_case_ignorable(char::from_u32(0x27).unwrap()));
        assert!(!super::is_case_ignorable('A'));
    }

    fn assert_mapping_boundaries(table: &[super::Mapping]) {
        let mut previous_source = None;
        for mapping in table {
            assert!((1..=3).contains(&mapping.len));
            if let Some(source) = previous_source {
                assert!(mapping.source > source);
            }
            assert_eq!(
                find_mapping(table, mapping.source).unwrap().source,
                mapping.source,
            );
            for target in mapping.target.iter().take(usize::from(mapping.len)) {
                assert!(char::from_u32(*target).is_some());
            }
            previous_source = Some(mapping.source);
        }
    }

    #[test]
    fn generated_case_mappings_are_sorted_bounded_scalars() {
        assert_mapping_boundaries(CASE_FOLD);
        assert_mapping_boundaries(LOWERCASE);
        assert_mapping_boundaries(UPPERCASE);
    }

    #[test]
    fn default_case_mappings_are_pinned() {
        assert_eq!(super::case_fold('A').iter().collect::<Vec<_>>(), vec!['a']);
        assert_eq!(
            super::case_fold('\u{00DF}').iter().collect::<Vec<_>>(),
            vec!['s', 's'],
        );
        assert_eq!(
            super::case_fold('\u{0130}').iter().collect::<Vec<_>>(),
            vec!['i', '\u{0307}'],
        );
        assert_eq!(super::case_fold('I').iter().collect::<Vec<_>>(), vec!['i']);
        assert_eq!(
            super::lowercase('\u{0130}').iter().collect::<Vec<_>>(),
            vec!['i', '\u{0307}'],
        );
        assert_eq!(
            super::uppercase('\u{00DF}').iter().collect::<Vec<_>>(),
            vec!['S', 'S'],
        );
        assert_eq!(
            super::uppercase('\u{FB03}').iter().collect::<Vec<_>>(),
            vec!['F', 'F', 'I'],
        );
        assert_eq!(super::lowercase('A').iter().collect::<Vec<_>>(), vec!['a']);
        assert_eq!(super::uppercase('a').iter().collect::<Vec<_>>(), vec!['A']);
    }
}
