//! Where document-level header items are accepted relative to body content.
//!
//! Each matrix row is one header item (a destination group or a control
//! word); each column is what precedes it in the document. Every cell pins
//! today's outcome: accepted, or rejected with one exact message. The
//! expectations are the parser's current behaviour, including the cells
//! marked "kept behaviour", which look inconsistent but are relied upon by
//! existing files.

#![allow(
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::raw::Tree;

use Cell::{Accepted as A, Rejected as R, RejectedWith as X};

/// What precedes (and, for the child-group columns, follows) the item.
struct Column {
    name: &'static str,
    before: &'static str,
    after: &'static str,
}

/// The expected outcome of one matrix cell.
#[derive(Clone, Copy)]
enum Cell {
    /// The document parses.
    Accepted,
    /// The document is rejected with the row's usual message.
    Rejected,
    /// The document is rejected with a different message.
    RejectedWith(&'static str),
}

/// One header item and its outcome in every column.
struct Row {
    name: &'static str,
    /// Header content the item depends on, placed before the column text.
    setup: &'static str,
    item: &'static str,
    /// The message of `Cell::Rejected`; empty when no cell uses it.
    rejection: &'static str,
    cells: [Cell; COLUMN_COUNT],
}

/// One standalone document and its outcome.
struct Case {
    name: &'static str,
    source: &'static str,
    expected: Result<(), &'static str>,
}

/// Parse `source` and reduce the result to accepted or the exact
/// `MalformedDocument` message. Any other error variant yields text that no
/// expectation matches.
fn outcome(source: &str) -> Result<(), String> {
    match Tree::parse(source, litchi_rtf::read::OpenOptions::default()) {
        Ok(_) => Ok(()),
        Err(litchi_rtf::Error::MalformedDocument(message)) => Err(message),
        Err(other) => Err(format!("unexpected error variant: {other:?}")),
    }
}

fn check_rows(rows: &[Row]) {
    let mut mismatches = Vec::new();
    for row in rows {
        for (column, cell) in COLUMNS.iter().zip(row.cells) {
            let source = format!(
                r"{{\rtf1\ansi{}{}{}{}}}",
                row.setup, column.before, row.item, column.after
            );
            let expected = match cell {
                Cell::Accepted => Ok(()),
                Cell::Rejected => {
                    assert!(!row.rejection.is_empty(), "{} has no message", row.name);
                    Err(row.rejection.to_owned())
                },
                Cell::RejectedWith(message) => Err(message.to_owned()),
            };
            let actual = outcome(&source);
            if actual != expected {
                mismatches.push(format!(
                    "{} after {}: {source}\n  expected {expected:?}\n  actual   {actual:?}",
                    row.name, column.name
                ));
            }
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

fn check_cases(cases: &[Case]) {
    let mut mismatches = Vec::new();
    for case in cases {
        let expected = case.expected.map_err(str::to_owned);
        let actual = outcome(case.source);
        if actual != expected {
            mismatches.push(format!(
                "{}: {}\n  expected {expected:?}\n  actual   {actual:?}",
                case.name, case.source
            ));
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

const COLUMN_COUNT: usize = 18;

const COLUMNS: [Column; COLUMN_COUNT] = [
    Column {
        name: r"nothing",
        before: r"",
        after: r"",
    },
    // Blank text is not a body event.
    Column {
        name: r"blank text",
        before: r"  ",
        after: r"",
    },
    Column {
        name: r"text",
        before: r" Body",
        after: r"",
    },
    // Kept behaviour: a unicode character closes the section prefix, and its flushed block
    // is what the body-text scans see.
    Column {
        name: r"unicode character",
        before: r"\u8364?",
        after: r"",
    },
    // Kept behaviour: table-cell text closes the section prefix but is never flushed into the
    // body blocks, so the header tables that only scan flushed text are still accepted.
    Column {
        name: r"text in an open table cell",
        before: r"\trowd\cellx1000\intbl Cell",
        after: r"",
    },
    Column {
        name: r"text in a finished table row",
        before: r"\trowd\cellx1000\intbl Cell\cell\row",
        after: r"",
    },
    // Kept behaviour: a unicode character closes the section prefix but not the document-wide
    // note-options window, so `\fet` is still accepted after one inside a table cell.
    Column {
        name: r"unicode character in an open table cell",
        before: r"\trowd\cellx1000\intbl\u8364?",
        after: r"",
    },
    Column {
        name: r"unicode character in a finished table row",
        before: r"\trowd\cellx1000\intbl\u8364?\cell\row",
        after: r"",
    },
    // Kept behaviour: a paragraph break, a field and a picture are not body events for the
    // section prefix, and the flushed-text scans do not see them either.
    Column {
        name: r"paragraph break only",
        before: r"\par",
        after: r"",
    },
    Column {
        name: r"field",
        before: r"{\field{\*\fldinst PAGE}{\fldrslt 1}}",
        after: r"",
    },
    Column {
        name: r"picture",
        before: r"{\pict\pngblip\picw1\pich1 89504E470D0A1A0A}",
        after: r"",
    },
    Column {
        name: r"page break",
        before: r"\page",
        after: r"",
    },
    Column {
        name: r"column break",
        before: r"\column",
        after: r"",
    },
    Column {
        name: r"soft line break",
        before: r"\softline",
        after: r"",
    },
    // Kept behaviour: a section break reopens the header window, so document-level header
    // controls are accepted again until the next body event. The flushed-text scans still
    // see the earlier text.
    Column {
        name: r"text, paragraph break, section break",
        before: r" Body\par\sect",
        after: r"",
    },
    Column {
        name: r"text, paragraph break, section break, text",
        before: r" Body\par\sect More",
        after: r"",
    },
    // One group deeper than the document root.
    Column {
        name: r"inside a child group",
        before: r"{",
        after: r"}",
    },
    Column {
        name: r"inside a child group after text",
        before: r" Body{",
        after: r"}",
    },
];

const DESTINATION_ROWS: &[Row] = &[
    Row {
        name: r"\fonttbl",
        setup: r"",
        item: r"{\fonttbl{\f0\fswiss Arial;}}",
        rejection: "RTF font table must occur exactly once at document scope before body text",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\colortbl",
        setup: r"",
        item: r"{\colortbl;\red255\green0\blue0;}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\stylesheet",
        setup: r"",
        item: r"{\stylesheet{\s0 Normal;}}",
        rejection: "RTF stylesheet must occur in the root document header",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\listtable",
        setup: r"",
        item: r"{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}{\listname L;}\listid3}}",
        rejection: "RTF list table must occur in the root header",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\listoverridetable",
        setup: r"{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}{\listname L;}\listid3}}",
        item: r"{\*\listoverridetable{\listoverride\listid3\listoverridecount0\ls1}}",
        rejection: "RTF list override table must follow listtable in the root header",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\info",
        setup: r"",
        item: r"{\info{\title T}}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\revtbl",
        setup: r"",
        item: r"{\*\revtbl{A;}}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\rsidtbl",
        setup: r"",
        item: r"{\*\rsidtbl \rsid7564464}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\xmlnstbl",
        setup: r"",
        item: r"{\*\xmlnstbl {\xmlns1 urn:x}}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\protusertbl",
        setup: r"",
        item: r"{\*\protusertbl{DOMAIN\'5calice}}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\userprops",
        setup: r"",
        item: r"{\*\userprops{\propname A}\proptype30{\staticval x}}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\themedata",
        setup: r"",
        item: r"{\*\themedata 0102}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\colorschememapping",
        setup: r"{\*\themedata 0102}",
        item: r"{\*\colorschememapping 3C3F786D6C3E}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\datastore",
        setup: r"",
        item: r"{\*\datastore 010500000200000018000000}",
        rejection: "",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A],
    },
    Row {
        name: r"\docvar",
        setup: r"",
        item: r"{\*\docvar {Name}{Value}}",
        rejection: "RTF docvar destination must occur in the root document header",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\filetbl",
        setup: r"",
        item: r"{\*\filetbl{\file\fid1 x;}}",
        rejection: "RTF filetbl must occur at document scope before body text",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\pgptbl",
        setup: r"",
        item: r"{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}",
        rejection: "RTF pgptbl must occur at document scope before body text",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\passwordhash",
        setup: r"",
        item: r"{\*\passwordhash 010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49}",
        rejection: "RTF passwordhash destination must occur in the root header",
        cells: [A, R, R, R, R, R, R, R, R, R, R, R, R, R, R, R, R, R],
    },
    Row {
        name: r"\defchp",
        setup: r"",
        item: r"{\*\defchp\fs20}",
        rejection: "RTF default-formatting destination must occur once in the root document header",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\ftnsep",
        setup: r"",
        item: r"{\*\ftnsep\chftnsep\par}",
        rejection: "RTF note separators must occur at document scope before body text",
        cells: [A, A, R, R, A, A, A, A, A, A, A, A, A, A, R, R, R, R],
    },
    Row {
        name: r"\pnseclvl",
        setup: r"",
        item: r"{\*\pnseclvl1\pndec{\pntxta .}}",
        rejection: "RTF late pnseclvl differs from the retained section defaults",
        cells: [
            A,
            A,
            R,
            R,
            A,
            A,
            A,
            A,
            A,
            A,
            A,
            A,
            A,
            A,
            R,
            R,
            X("RTF pnseclvl destinations must occur at document scope before body text"),
            X("RTF pnseclvl destinations must occur at document scope before body text"),
        ],
    },
    Row {
        name: r"\fchars",
        setup: r"",
        item: r"{\*\fchars !}",
        rejection: "RTF kinsoku destinations must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"starred \windowcaption",
        setup: r"",
        item: r"{\*\windowcaption W}",
        rejection: "RTF window caption must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"plain \windowcaption",
        setup: r"",
        item: r"{\windowcaption W}",
        rejection: "RTF window caption must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\xform",
        setup: r"",
        item: r"{\*\xform transform.xsl}",
        rejection: "RTF XSL transform must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\wgrffmtfilter",
        setup: r"",
        item: r"{\*\wgrffmtfilter 2002}",
        rejection: "RTF style-list filter must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\writereservation",
        setup: r"",
        item: r"{\*\writereservation legacy}",
        rejection: "RTF write reservation must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\nextfile",
        setup: r"",
        item: r"{\*\nextfile next.rtf}",
        rejection: "RTF external document references must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\template",
        setup: r"",
        item: r"{\*\template t.dot}",
        rejection: "RTF external document references must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"starred \viewkind",
        setup: r"",
        item: r"{\*\viewkind1}",
        rejection: "RTF document-view controls must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\background",
        setup: r"",
        item: r"{\*\background{\shp{\*\shpinst}}}",
        rejection: "RTF background destination must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
];

const CONTROL_WORD_ROWS: &[Row] = &[
    Row {
        name: r"\deff0",
        setup: r"",
        item: r"\deff0",
        rejection: "RTF default-font selectors must occur in the root document header",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\deftab720",
        setup: r"",
        item: r"\deftab720",
        rejection: "RTF deftab must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\ksulang1041",
        setup: r"",
        item: r"\ksulang1041",
        rejection: "RTF ksulang must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\makebackup",
        setup: r"",
        item: r"\makebackup",
        rejection: "RTF file-setting flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\muser",
        setup: r"",
        item: r"\muser",
        rejection: "RTF output-setting flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\horzdoc",
        setup: r"",
        item: r"\horzdoc",
        rejection: "RTF rendering flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\fracwidth",
        setup: r"",
        item: r"\fracwidth",
        rejection: "RTF processing property must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\dgsnap",
        setup: r"",
        item: r"\dgsnap",
        rejection: "RTF drawing-grid property must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\facingp",
        setup: r"",
        item: r"\facingp",
        rejection: "RTF print-layout setting must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\themelang1033",
        setup: r"",
        item: r"\themelang1033",
        rejection: "RTF theme language must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\validatexml1",
        setup: r"",
        item: r"\validatexml1",
        rejection: "RTF XML policy must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\donotembedsysfont1",
        setup: r"",
        item: r"\donotembedsysfont1",
        rejection: "RTF embedding policy must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\trackmoves1",
        setup: r"",
        item: r"\trackmoves1",
        rejection: "RTF revision policy must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\stylelocktheme",
        setup: r"",
        item: r"\stylelocktheme",
        rejection: "RTF style policy must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\stylelockenforced",
        setup: r"",
        item: r"\stylelockenforced",
        rejection: "RTF style restriction must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\bookfold",
        setup: r"",
        item: r"\bookfold",
        rejection: "RTF booklet-printing property must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\rempersonalinfo",
        setup: r"",
        item: r"\rempersonalinfo",
        rejection: "RTF privacy policy must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\spltpgpar",
        setup: r"",
        item: r"\spltpgpar",
        rejection: "RTF Word 2003 compatibility flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\nouicompat",
        setup: r"",
        item: r"\nouicompat",
        rejection: "RTF compatibility policy must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\wrppunct",
        setup: r"",
        item: r"\wrppunct",
        rejection: "RTF Asian grid compatibility flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\ftnlytwnine",
        setup: r"",
        item: r"\ftnlytwnine",
        rejection: "RTF legacy layout compatibility flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\lyttblrtgr",
        setup: r"",
        item: r"\lyttblrtgr",
        rejection: "RTF table-layout compatibility flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\dntblnsbdb",
        setup: r"",
        item: r"\dntblnsbdb",
        rejection: "RTF East Asian compatibility flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\sprstsp",
        setup: r"",
        item: r"\sprstsp",
        rejection: "RTF line-spacing compatibility flag must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\fromtext",
        setup: r"",
        item: r"\fromtext",
        rejection: "RTF document origin must occur in the header before font tables and visible text",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\doctype1",
        setup: r"",
        item: r"\doctype1",
        rejection: "RTF document type must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\readonlyrecommended",
        setup: r"",
        item: r"\readonlyrecommended",
        rejection: "RTF save preference must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\stylesortmethod4",
        setup: r"",
        item: r"\stylesortmethod4",
        rejection: "RTF style-sort method must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\usexform",
        setup: r"",
        item: r"\usexform",
        rejection: "RTF usexform must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\donotshowmarkup",
        setup: r"",
        item: r"\donotshowmarkup",
        rejection: "RTF review-display flags must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\viewkind1",
        setup: r"",
        item: r"\viewkind1",
        rejection: "RTF document-view controls must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\hyphauto1",
        setup: r"",
        item: r"\hyphauto1",
        rejection: "RTF document hyphenation controls must precede visible text at document root",
        cells: [A, A, R, R, R, R, R, R, A, A, A, R, R, R, A, R, R, R],
    },
    Row {
        name: r"\rsidroot7",
        setup: r"",
        item: r"\rsidroot7",
        rejection: "RTF rsidroot must occur exactly once at document scope",
        cells: [A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, A, R, R],
    },
    Row {
        name: r"\fet0",
        setup: r"",
        item: r"\fet0",
        rejection: "RTF note options must precede body text at document root",
        cells: [
            A,
            A,
            R,
            R,
            R,
            R,
            A,
            A,
            A,
            A,
            A,
            R,
            R,
            R,
            R,
            R,
            X("RTF note options must be root document-format controls"),
            X("RTF note options must be root document-format controls"),
        ],
    },
    Row {
        name: r"\sftnbj",
        setup: r"",
        item: r"\sftnbj",
        rejection: "RTF section note options must precede section content at document root",
        cells: [
            A,
            A,
            R,
            R,
            R,
            R,
            R,
            R,
            A,
            A,
            A,
            R,
            R,
            R,
            A,
            R,
            X("RTF section note options must be root controls or visible field-result formatting"),
            X("RTF section note options must be root controls or visible field-result formatting"),
        ],
    },
    Row {
        name: r"\formprot",
        setup: r"",
        item: r"\formprot",
        rejection: "RTF document protection controls must occur in the root header",
        cells: [A, R, R, R, A, A, A, A, R, A, A, A, A, A, R, R, R, R],
    },
];

#[test]
fn header_destinations_follow_each_preceding_event() {
    check_rows(DESTINATION_ROWS);
}

#[test]
fn header_control_words_follow_each_preceding_event() {
    check_rows(CONTROL_WORD_ROWS);
}

/// The matrix must be able to fail in both directions: every column that is
/// a body event has a rejected cell, and every column has an accepted one.
#[test]
fn matrix_holds_accepted_and_rejected_cells() {
    let rows = || DESTINATION_ROWS.iter().chain(CONTROL_WORD_ROWS);
    for index in 0..COLUMN_COUNT {
        let accepted = rows().any(|row| matches!(row.cells.get(index), Some(Cell::Accepted)));
        assert!(accepted, "column {index} has no accepted cell");
    }
    let rejected = rows()
        .flat_map(|row| row.cells)
        .filter(|cell| !matches!(cell, Cell::Accepted))
        .count();
    assert!(rejected > COLUMN_COUNT);
}

#[test]
fn once_only_items_given_twice() {
    check_cases(&[
        Case {
            name: r"\fonttbl",
            source: r"{\rtf1\ansi{\fonttbl{\f0\fswiss Arial;}}{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Err(
                "RTF font table must occur exactly once at document scope before body text",
            ),
        },
        Case {
            name: r"\colortbl",
            source: r"{\rtf1\ansi{\colortbl;\red255\green0\blue0;}{\colortbl;\red255\green0\blue0;}}",
            expected: Ok(()),
        },
        Case {
            name: r"\stylesheet",
            source: r"{\rtf1\ansi{\stylesheet{\s0 Normal;}}{\stylesheet{\s0 Normal;}}}",
            expected: Err("RTF document contains multiple stylesheet destinations"),
        },
        Case {
            name: r"\listtable",
            source: r"{\rtf1\ansi{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}{\listname L;}\listid3}}{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}{\listname L;}\listid3}}}",
            expected: Err("RTF document contains multiple list tables"),
        },
        Case {
            name: r"\listoverridetable",
            source: r"{\rtf1\ansi{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}{\listname L;}\listid3}}{\*\listoverridetable{\listoverride\listid3\listoverridecount0\ls1}}{\*\listoverridetable{\listoverride\listid3\listoverridecount0\ls1}}}",
            expected: Err("RTF document contains multiple list override tables"),
        },
        Case {
            name: r"\info",
            source: r"{\rtf1\ansi{\info{\title T}}{\info{\title T}}}",
            expected: Err("RTF contains multiple info groups"),
        },
        Case {
            name: r"\revtbl",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\*\revtbl{A;}}}",
            expected: Err("RTF contains multiple revision-author tables"),
        },
        Case {
            name: r"\rsidtbl",
            source: r"{\rtf1\ansi{\*\rsidtbl \rsid7564464}{\*\rsidtbl \rsid7564464}}",
            expected: Err("RTF contains multiple revision-save tables"),
        },
        Case {
            name: r"\xmlnstbl",
            source: r"{\rtf1\ansi{\*\xmlnstbl {\xmlns1 urn:x}}{\*\xmlnstbl {\xmlns1 urn:x}}}",
            expected: Err("RTF contains multiple XML namespace tables"),
        },
        Case {
            name: r"\protusertbl",
            source: r"{\rtf1\ansi{\*\protusertbl{DOMAIN\'5calice}}{\*\protusertbl{DOMAIN\'5calice}}}",
            expected: Err("RTF contains multiple protection-user tables"),
        },
        Case {
            name: r"\userprops",
            source: r"{\rtf1\ansi{\*\userprops{\propname A}\proptype30{\staticval x}}{\*\userprops{\propname A}\proptype30{\staticval x}}}",
            expected: Err("RTF document contains multiple userprops destinations"),
        },
        Case {
            name: r"\themedata",
            source: r"{\rtf1\ansi{\*\themedata 0102}{\*\themedata 0102}}",
            expected: Err("RTF contains multiple theme-data destinations"),
        },
        Case {
            name: r"\colorschememapping",
            source: r"{\rtf1\ansi{\*\themedata 0102}{\*\colorschememapping 3C3F786D6C3E}{\*\colorschememapping 3C3F786D6C3E}}",
            expected: Err("RTF contains multiple color-scheme mappings"),
        },
        Case {
            name: r"\datastore",
            source: r"{\rtf1\ansi{\*\datastore 010500000200000018000000}{\*\datastore 010500000200000018000000}}",
            expected: Err("RTF contains multiple datastore destinations"),
        },
        Case {
            name: r"\docvar",
            source: r"{\rtf1\ansi{\*\docvar {Name}{Value}}{\*\docvar {Name}{Value}}}",
            expected: Ok(()),
        },
        Case {
            name: r"\filetbl",
            source: r"{\rtf1\ansi{\*\filetbl{\file\fid1 x;}}{\*\filetbl{\file\fid1 x;}}}",
            expected: Err("RTF contains multiple filetbl destinations"),
        },
        Case {
            name: r"\pgptbl",
            source: r"{\rtf1\ansi{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}}",
            expected: Err("RTF contains multiple pgptbl destinations"),
        },
        Case {
            name: r"\passwordhash",
            source: r"{\rtf1\ansi{\*\passwordhash 010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49}{\*\passwordhash 010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49}}",
            expected: Err("duplicate RTF modern protection password hash"),
        },
        Case {
            name: r"\defchp",
            source: r"{\rtf1\ansi{\*\defchp\fs20}{\*\defchp\fs20}}",
            expected: Err("duplicate RTF default-formatting destination"),
        },
        Case {
            name: r"\ftnsep",
            source: r"{\rtf1\ansi{\*\ftnsep\chftnsep\par}{\*\ftnsep\chftnsep\par}}",
            expected: Err("RTF note-separator destinations are duplicated or out of order"),
        },
        Case {
            name: r"\pnseclvl",
            source: r"{\rtf1\ansi{\*\pnseclvl1\pndec{\pntxta .}}{\*\pnseclvl1\pndec{\pntxta .}}}",
            expected: Err("RTF pnseclvl destinations are duplicated or out of order"),
        },
        Case {
            name: r"\fchars",
            source: r"{\rtf1\ansi{\*\fchars !}{\*\fchars !}}",
            expected: Err("duplicate RTF kinsoku destination"),
        },
        Case {
            name: r"starred \windowcaption",
            source: r"{\rtf1\ansi{\*\windowcaption W}{\*\windowcaption W}}",
            expected: Err("duplicate RTF window caption destination"),
        },
        Case {
            name: r"plain \windowcaption",
            source: r"{\rtf1\ansi{\windowcaption W}{\windowcaption W}}",
            expected: Err("duplicate RTF window caption destination"),
        },
        Case {
            name: r"\xform",
            source: r"{\rtf1\ansi{\*\xform transform.xsl}{\*\xform transform.xsl}}",
            expected: Err("duplicate RTF XSL transform destination"),
        },
        Case {
            name: r"\wgrffmtfilter",
            source: r"{\rtf1\ansi{\*\wgrffmtfilter 2002}{\*\wgrffmtfilter 2002}}",
            expected: Err("duplicate RTF style-list filter destination"),
        },
        Case {
            name: r"\writereservation",
            source: r"{\rtf1\ansi{\*\writereservation legacy}{\*\writereservation legacy}}",
            expected: Err("duplicate RTF legacy write-reservation destination"),
        },
        Case {
            name: r"\nextfile",
            source: r"{\rtf1\ansi{\*\nextfile next.rtf}{\*\nextfile next.rtf}}",
            expected: Err("duplicate RTF external document reference destination"),
        },
        Case {
            name: r"\template",
            source: r"{\rtf1\ansi{\*\template t.dot}{\*\template t.dot}}",
            expected: Err("duplicate RTF external document reference destination"),
        },
        Case {
            name: r"starred \viewkind",
            source: r"{\rtf1\ansi{\*\viewkind1}{\*\viewkind1}}",
            expected: Err("duplicate RTF viewkind document property"),
        },
        Case {
            name: r"\background",
            source: r"{\rtf1\ansi{\*\background{\shp{\*\shpinst}}}{\*\background{\shp{\*\shpinst}}}}",
            expected: Err("RTF contains multiple background destinations"),
        },
        Case {
            name: r"\deff0",
            source: r"{\rtf1\ansi\deff0\deff0}",
            expected: Err("duplicate RTF deff default-font selector"),
        },
        Case {
            name: r"\deftab720",
            source: r"{\rtf1\ansi\deftab720\deftab720}",
            expected: Ok(()),
        },
        Case {
            name: r"\ksulang1041",
            source: r"{\rtf1\ansi\ksulang1041\ksulang1041}",
            expected: Err("duplicate RTF ksulang document property"),
        },
        Case {
            name: r"\makebackup",
            source: r"{\rtf1\ansi\makebackup\makebackup}",
            expected: Err("duplicate RTF makebackup document property"),
        },
        Case {
            name: r"\muser",
            source: r"{\rtf1\ansi\muser\muser}",
            expected: Err("duplicate RTF muser document property"),
        },
        Case {
            name: r"\horzdoc",
            source: r"{\rtf1\ansi\horzdoc\horzdoc}",
            expected: Err("duplicate or conflicting RTF horzdoc rendering property"),
        },
        Case {
            name: r"\fracwidth",
            source: r"{\rtf1\ansi\fracwidth\fracwidth}",
            expected: Err("duplicate RTF fracwidth document property"),
        },
        Case {
            name: r"\dgsnap",
            source: r"{\rtf1\ansi\dgsnap\dgsnap}",
            expected: Err("duplicate RTF dgsnap drawing-grid property"),
        },
        Case {
            name: r"\facingp",
            source: r"{\rtf1\ansi\facingp\facingp}",
            expected: Err("duplicate RTF facingp document property"),
        },
        Case {
            name: r"\themelang1033",
            source: r"{\rtf1\ansi\themelang1033\themelang1033}",
            expected: Err("duplicate RTF themelang document property"),
        },
        Case {
            name: r"\validatexml1",
            source: r"{\rtf1\ansi\validatexml1\validatexml1}",
            expected: Err("duplicate RTF validatexml document property"),
        },
        Case {
            name: r"\donotembedsysfont1",
            source: r"{\rtf1\ansi\donotembedsysfont1\donotembedsysfont1}",
            expected: Err("duplicate RTF donotembedsysfont document property"),
        },
        Case {
            name: r"\trackmoves1",
            source: r"{\rtf1\ansi\trackmoves1\trackmoves1}",
            expected: Err("duplicate RTF trackmoves document property"),
        },
        Case {
            name: r"\stylelocktheme",
            source: r"{\rtf1\ansi\stylelocktheme\stylelocktheme}",
            expected: Err("duplicate RTF stylelocktheme document property"),
        },
        Case {
            name: r"\stylelockenforced",
            source: r"{\rtf1\ansi\stylelockenforced\stylelockenforced}",
            expected: Err("duplicate RTF stylelockenforced document property"),
        },
        Case {
            name: r"\bookfold",
            source: r"{\rtf1\ansi\bookfold\bookfold}",
            expected: Err("duplicate RTF bookfold document property"),
        },
        Case {
            name: r"\rempersonalinfo",
            source: r"{\rtf1\ansi\rempersonalinfo\rempersonalinfo}",
            expected: Err("duplicate RTF rempersonalinfo document property"),
        },
        Case {
            name: r"\spltpgpar",
            source: r"{\rtf1\ansi\spltpgpar\spltpgpar}",
            expected: Err("duplicate RTF spltpgpar document property"),
        },
        Case {
            name: r"\nouicompat",
            source: r"{\rtf1\ansi\nouicompat\nouicompat}",
            expected: Err("duplicate RTF nouicompat document property"),
        },
        Case {
            name: r"\wrppunct",
            source: r"{\rtf1\ansi\wrppunct\wrppunct}",
            expected: Err("duplicate RTF wrppunct document property"),
        },
        Case {
            name: r"\ftnlytwnine",
            source: r"{\rtf1\ansi\ftnlytwnine\ftnlytwnine}",
            expected: Err("duplicate RTF ftnlytwnine document property"),
        },
        Case {
            name: r"\lyttblrtgr",
            source: r"{\rtf1\ansi\lyttblrtgr\lyttblrtgr}",
            expected: Err("duplicate RTF lyttblrtgr document property"),
        },
        Case {
            name: r"\dntblnsbdb",
            source: r"{\rtf1\ansi\dntblnsbdb\dntblnsbdb}",
            expected: Err("duplicate RTF dntblnsbdb document property"),
        },
        Case {
            name: r"\sprstsp",
            source: r"{\rtf1\ansi\sprstsp\sprstsp}",
            expected: Err("duplicate RTF sprstsp document property"),
        },
        Case {
            name: r"\fromtext",
            source: r"{\rtf1\ansi\fromtext\fromtext}",
            expected: Err("duplicate or conflicting RTF document-origin property"),
        },
        Case {
            name: r"\doctype1",
            source: r"{\rtf1\ansi\doctype1\doctype1}",
            expected: Err("duplicate RTF document type property"),
        },
        Case {
            name: r"\readonlyrecommended",
            source: r"{\rtf1\ansi\readonlyrecommended\readonlyrecommended}",
            expected: Err("duplicate RTF readonlyrecommended document property"),
        },
        Case {
            name: r"\stylesortmethod4",
            source: r"{\rtf1\ansi\stylesortmethod4\stylesortmethod4}",
            expected: Err("duplicate RTF style-sort method document property"),
        },
        Case {
            name: r"\usexform",
            source: r"{\rtf1\ansi\usexform\usexform}",
            expected: Err("duplicate RTF usexform document property"),
        },
        Case {
            name: r"\donotshowmarkup",
            source: r"{\rtf1\ansi\donotshowmarkup\donotshowmarkup}",
            expected: Err("duplicate RTF donotshowmarkup document property"),
        },
        Case {
            name: r"\viewkind1",
            source: r"{\rtf1\ansi\viewkind1\viewkind1}",
            expected: Err("duplicate RTF viewkind document property"),
        },
        Case {
            name: r"\hyphauto1",
            source: r"{\rtf1\ansi\hyphauto1\hyphauto1}",
            expected: Err("duplicate RTF hyphauto document property"),
        },
        Case {
            name: r"\rsidroot7",
            source: r"{\rtf1\ansi\rsidroot7\rsidroot7}",
            expected: Err("RTF rsidroot must occur exactly once at document scope"),
        },
        Case {
            name: r"\fet0",
            source: r"{\rtf1\ansi\fet0\fet0}",
            expected: Ok(()),
        },
        Case {
            name: r"\sftnbj",
            source: r"{\rtf1\ansi\sftnbj\sftnbj}",
            expected: Ok(()),
        },
        Case {
            name: r"\formprot",
            source: r"{\rtf1\ansi\formprot\formprot}",
            expected: Err("duplicate RTF formprot control"),
        },
    ]);
}

#[test]
fn list_override_table_needs_the_list_table_first() {
    check_cases(&[
        Case {
            name: r"\listoverridetable without \listtable",
            source: r"{\rtf1\ansi{\*\listoverridetable{\listoverride\listid3\listoverridecount0\ls1}}}",
            expected: Err("RTF list override table must follow listtable in the root header"),
        },
        Case {
            name: r"\listoverridetable before \listtable",
            source: r"{\rtf1\ansi{\*\listoverridetable}{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}{\listname L;}\listid3}}}",
            expected: Err("RTF list override table must follow listtable in the root header"),
        },
    ]);
}

#[test]
fn font_table_depth_and_unicode_alternate_groups() {
    check_cases(&[
        // The only place a font table is accepted two groups below the root.
        Case {
            name: r"\fonttbl two groups deep inside \upr",
            source: r"{\rtf1\ansi{\upr{\fonttbl{\f0\fswiss Arial;}}{\*\ud{\fonttbl{\f0\fswiss Arial;}}}}}",
            expected: Ok(()),
        },
        Case {
            name: r"\fonttbl two groups deep outside \upr",
            source: r"{\rtf1\ansi{{\fonttbl{\f0\fswiss Arial;}}}}",
            expected: Err(
                "RTF font table must occur exactly once at document scope before body text",
            ),
        },
        Case {
            name: r"\fonttbl three groups deep inside \upr",
            source: r"{\rtf1\ansi{\upr{{\fonttbl{\f0\fswiss Arial;}}}{\*\ud{{\fonttbl{\f0\fswiss Arial;}}}}}}",
            expected: Err(
                "RTF font table must occur exactly once at document scope before body text",
            ),
        },
        Case {
            name: r"\fromtext after \fonttbl",
            source: r"{\rtf1\ansi{\fonttbl{\f0\fswiss Arial;}}\fromtext}",
            expected: Err(
                "RTF document origin must occur in the header before font tables and visible text",
            ),
        },
        Case {
            name: r"\fromtext before \fonttbl",
            source: r"{\rtf1\ansi\fromtext{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Ok(()),
        },
    ]);
}

#[test]
fn section_note_options_and_the_late_section_format_run() {
    check_cases(&[
        Case {
            name: r"after \sectd",
            source: r"{\rtf1\ansi\sectd\sftnbj Body}",
            expected: Ok(()),
        },
        Case {
            name: r"after text",
            source: r"{\rtf1\ansi Body\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        // Kept behaviour: a root `\sectd` starts a late section-format run in which section note
        // options are accepted although the section already has content.
        Case {
            name: r"after text and \sectd",
            source: r"{\rtf1\ansi Body\sectd\sftnbj}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, \sectd and text",
            source: r"{\rtf1\ansi Body\sectd More\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"after text, \sectd and a unicode character",
            source: r"{\rtf1\ansi Body\sectd\u8364?\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        // Blank text does not end the run.
        Case {
            name: r"after text, \sectd and blank text",
            source: r"{\rtf1\ansi Body\sectd  \sftnbj}",
            expected: Ok(()),
        },
        // Opening a child group at the root ends the run.
        Case {
            name: r"after text, \sectd and a child group",
            source: r"{\rtf1\ansi Body\sectd{}\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        // `\sbk` ends the run; the other break-type controls do not.
        Case {
            name: r"after text, \sectd and \sbk",
            source: r"{\rtf1\ansi Body\sectd\sbk\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"after text, \sectd and \sbknone",
            source: r"{\rtf1\ansi Body\sectd\sbknone\sftnbj}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, \sectd and \page",
            source: r"{\rtf1\ansi Body\sectd\page\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"after text, \sectd and \column",
            source: r"{\rtf1\ansi Body\sectd\column\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"after text, \sectd and \softline",
            source: r"{\rtf1\ansi Body\sectd\softline\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        // A paragraph break is not a body event.
        Case {
            name: r"after text, \sectd and \par",
            source: r"{\rtf1\ansi Body\sectd\par\sftnbj}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, \sectd and \sect",
            source: r"{\rtf1\ansi Body\sectd\sect\sftnbj}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, \sectd, \sbk and \sectd",
            source: r"{\rtf1\ansi Body\sectd\sbk\sectd\sftnbj}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, in a child group opening with \sectd",
            source: r"{\rtf1\ansi Body{\sectd\sftnbj}}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, in a child group with a late \sectd",
            source: r"{\rtf1\ansi Body{\b\sectd\sftnbj}}",
            expected: Err(
                "RTF section note options must be root controls or visible field-result formatting",
            ),
        },
        Case {
            name: r"after text, in a child group without \sectd",
            source: r"{\rtf1\ansi Body{\sftnbj}}",
            expected: Err(
                "RTF section note options must be root controls or visible field-result formatting",
            ),
        },
        Case {
            name: r"after text and \sectd inside a closed child group",
            source: r"{\rtf1\ansi Body{\sectd}\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"in a child group before text",
            source: r"{\rtf1\ansi{\sftnbj}}",
            expected: Err(
                "RTF section note options must be root controls or visible field-result formatting",
            ),
        },
        // A header or footer in the current section closes the prefix for section note options
        // only; it is tracked on the section, not by the parser phase.
        Case {
            name: r"after a header",
            source: r"{\rtf1\ansi{\header H}\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"after a header and \sectd",
            source: r"{\rtf1\ansi{\header H}\sectd\sftnbj}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, \sect and a header",
            source: r"{\rtf1\ansi Body\par\sect{\header H}\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        // Kept behaviour: the section break reopens the section prefix.
        Case {
            name: r"after text and \sect",
            source: r"{\rtf1\ansi Body\par\sect\sftnbj More}",
            expected: Ok(()),
        },
        Case {
            name: r"after text, \sect and text",
            source: r"{\rtf1\ansi Body\par\sect More\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"after text, \sect and a unicode character",
            source: r"{\rtf1\ansi Body\par\sect\u8364?\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"after a header, text and \sect",
            source: r"{\rtf1\ansi{\header H} Body\par\sect\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
    ]);
}

#[test]
fn header_window_after_a_section_break() {
    check_cases(&[
        Case {
            name: r"\deff after text, \sect and a unicode character",
            source: r"{\rtf1\ansi Body\par\sect\u8364?\deff0}",
            expected: Err("RTF default-font selectors must occur in the root document header"),
        },
        Case {
            name: r"\deff after text, \sect and \page",
            source: r"{\rtf1\ansi Body\par\sect\page\deff0}",
            expected: Err("RTF default-font selectors must occur in the root document header"),
        },
        // Kept behaviour: after a section break the header window stays open until a body event.
        Case {
            name: r"\deff after text, \sect and \par",
            source: r"{\rtf1\ansi Body\par\sect\par\deff0}",
            expected: Ok(()),
        },
        Case {
            name: r"\deff after text, \sect and \sectd",
            source: r"{\rtf1\ansi Body\par\sect\sectd\deff0}",
            expected: Ok(()),
        },
        Case {
            name: r"\deff after text, \sect and blank text",
            source: r"{\rtf1\ansi Body\par\sect  \deff0}",
            expected: Ok(()),
        },
        Case {
            name: r"\deff after text, \sect and a child group",
            source: r"{\rtf1\ansi Body\par\sect{}\deff0}",
            expected: Ok(()),
        },
        Case {
            name: r"\deff after \sect alone",
            source: r"{\rtf1\ansi\sect\deff0}",
            expected: Ok(()),
        },
        Case {
            name: r"\defchp after \sect alone",
            source: r"{\rtf1\ansi\sect{\*\defchp\fs20}}",
            expected: Ok(()),
        },
        Case {
            name: r"\fet after \sect alone",
            source: r"{\rtf1\ansi\sect\fet0}",
            expected: Ok(()),
        },
        Case {
            name: r"\fonttbl after \sect alone",
            source: r"{\rtf1\ansi\sect{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Ok(()),
        },
    ]);
}

#[test]
fn body_events_in_other_positions() {
    check_cases(&[
        Case {
            name: r"\deff after text in a child group",
            source: r"{\rtf1\ansi{ Body}\deff0}",
            expected: Err("RTF default-font selectors must occur in the root document header"),
        },
        Case {
            name: r"\deff after a unicode character in a child group",
            source: r"{\rtf1\ansi{\u8364?}\deff0}",
            expected: Err("RTF default-font selectors must occur in the root document header"),
        },
        Case {
            name: r"\deff after a unicode character in a font name",
            source: r"{\rtf1\ansi{\fonttbl{\f0\fswiss A\u8364?;}}\deff0}",
            expected: Ok(()),
        },
        // Kept behaviour: a header does not close the header window for document-level controls.
        Case {
            name: r"\deff after a header",
            source: r"{\rtf1\ansi{\header H}\deff0}",
            expected: Ok(()),
        },
        // Deleted text is a body event although it adds no flushed body text.
        Case {
            name: r"\deff after a deleted text run",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0 gone}\deff0}",
            expected: Err("RTF default-font selectors must occur in the root document header"),
        },
        // Kept behaviour: the flushed-text scan does not see deleted text.
        Case {
            name: r"\fonttbl after a deleted text run",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0 gone}{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Ok(()),
        },
        Case {
            name: r"\fet after a deleted text run",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0 gone}\fet0}",
            expected: Err("RTF note options must precede body text at document root"),
        },
        // Kept behaviour: protection controls test the body text length, which deleted text does
        // not change.
        Case {
            name: r"\formprot after a deleted text run",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0 gone}\formprot}",
            expected: Ok(()),
        },
        Case {
            name: r"\fonttbl after a deleted unicode character",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0\u8364?}{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Ok(()),
        },
        // Kept behaviour: a unicode character closes the section prefix (see the `\deff` and
        // `\sftnbj` cases below) but not the document-wide note-options window, so with no
        // flushed block the document note options are still accepted.
        Case {
            name: r"\fet after a deleted unicode character",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0\u8364?}\fet0}",
            expected: Ok(()),
        },
        Case {
            name: r"\deff after a deleted unicode character",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0\u8364?}\deff0}",
            expected: Err("RTF default-font selectors must occur in the root document header"),
        },
        Case {
            name: r"\sftnbj after a deleted unicode character",
            source: r"{\rtf1\ansi{\*\revtbl{A;}}{\deleted\revauthdel0\u8364?}\sftnbj}",
            expected: Err("RTF section note options must precede section content at document root"),
        },
        Case {
            name: r"\fonttbl after a unicode character and trailing text",
            source: r"{\rtf1\ansi\uc1\u8364?xy{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Err(
                "RTF font table must occur exactly once at document scope before body text",
            ),
        },
        // Kept behaviour: the header-table scans trim, so a blank unicode character is not body text.
        Case {
            name: r"\fonttbl after a blank unicode character",
            source: r"{\rtf1\ansi\u32?{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Ok(()),
        },
        Case {
            name: r"\fet after a blank unicode character",
            source: r"{\rtf1\ansi\u32?\fet0}",
            expected: Err("RTF note options must precede body text at document root"),
        },
        Case {
            name: r"\formprot after a blank unicode character",
            source: r"{\rtf1\ansi\u32?\formprot}",
            expected: Err("RTF document protection controls must occur in the root header"),
        },
        Case {
            name: r"\fonttbl after text in a child group",
            source: r"{\rtf1\ansi{ Body}{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Err(
                "RTF font table must occur exactly once at document scope before body text",
            ),
        },
        Case {
            name: r"\fet after text in a child group",
            source: r"{\rtf1\ansi{ Body}\fet0}",
            expected: Err("RTF note options must precede body text at document root"),
        },
        // Kept behaviour: the note-options scan does not trim, so a flushed blank block closes it
        // while the header tables (which trim) stay open.
        Case {
            name: r"\fet after blank text in a child group",
            source: r"{\rtf1\ansi{  }\fet0}",
            expected: Err("RTF note options must precede body text at document root"),
        },
        Case {
            name: r"\fet after \softpage",
            source: r"{\rtf1\ansi\softpage\fet0}",
            expected: Err("RTF note options must precede body text at document root"),
        },
        Case {
            name: r"\fet after \softcol",
            source: r"{\rtf1\ansi\softcol\fet0}",
            expected: Err("RTF note options must precede body text at document root"),
        },
        Case {
            name: r"\fet after \softlheight",
            source: r"{\rtf1\ansi\softlheight240\fet0}",
            expected: Err("RTF note options must precede body text at document root"),
        },
        // Kept behaviour: a tab is not a body event.
        Case {
            name: r"\fonttbl after a tab",
            source: r"{\rtf1\ansi\tab{\fonttbl{\f0\fswiss Arial;}}}",
            expected: Ok(()),
        },
        Case {
            name: r"\fet after a tab",
            source: r"{\rtf1\ansi\tab\fet0}",
            expected: Ok(()),
        },
        Case {
            name: r"\deff after a tab",
            source: r"{\rtf1\ansi\tab\deff0}",
            expected: Ok(()),
        },
    ]);
}
