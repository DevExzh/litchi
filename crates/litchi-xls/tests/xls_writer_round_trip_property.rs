#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "property checks panic with the failing case"
)]

//! A writer-to-reader round-trip property over random workbooks (change
//! 0766).
//!
//! Each case drives one [`Writer`] through a random sequence of registration
//! calls — worksheet and defined names, strings, numbers, formulas,
//! hyperlinks, `AutoFilter` conditions, fonts, cell styles and number formats
//! — with lengths drawn around every BIFF8 limit and text drawn from ASCII,
//! Latin-1, other BMP scripts and supplementary-plane characters, plus the
//! characters BIFF8 forbids in names. Then:
//!
//! - every refused call returned a typed error and changed nothing: a second
//!   writer that received only the accepted calls writes the same bytes;
//! - the writer writes the same bytes twice;
//! - litchi's reader opens the workbook and reads back every accepted input.
//!
//! `LITCHI_XLS_PROPERTY_CASES` and `LITCHI_XLS_PROPERTY_SEED` widen or move
//! the search; the default is a fixed set of seeds.

use std::collections::BTreeMap;
use std::io::Cursor;

use litchi_core::sheet::{Cell as _, CellValue as ReadValue};
use litchi_xls::autofilter::FilterValue;
use litchi_xls::writer::formula::encode_ptg_tokens;
use litchi_xls::writer::{
    AutoFilterConditionWrite, CellStyle, Font, FormattingManager, FormulaTokenizer,
    HorizontalAlignment, VerticalAlignment, Writer,
};
use litchi_xls::{DefinedNameKind, Error, FontUnderline, NameScope};

type Workbook = litchi_xls::Workbook<Cursor<Vec<u8>>>;

/// xorshift64*: small, deterministic and good enough to pick test inputs.
struct Random(u64);

impl Random {
    fn next(&mut self) -> u64 {
        let mut value = self.0;
        value ^= value >> 12;
        value ^= value << 25;
        value ^= value >> 27;
        self.0 = value;
        value.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    fn below(&mut self, bound: usize) -> usize {
        usize::try_from(self.next() % u64::try_from(bound).unwrap()).unwrap()
    }

    fn chance(&mut self, numerator: usize, denominator: usize) -> bool {
        self.below(denominator) < numerator
    }

    fn pick<'a, T>(&mut self, values: &'a [T]) -> &'a T {
        &values[self.below(values.len())]
    }
}

/// Character classes the strings are drawn from.
#[derive(Clone, Copy)]
enum Class {
    Ascii,
    Latin1,
    Greek,
    Cjk,
    Astral,
    /// Characters BIFF8 forbids in worksheet names, NUL and U+0003.
    Hostile,
}

fn character(random: &mut Random, class: Class) -> char {
    let pick = |random: &mut Random, low: u32, high: u32| {
        char::from_u32(
            low + u32::try_from(random.below(usize::try_from(high - low + 1).unwrap())).unwrap(),
        )
        .unwrap()
    };
    match class {
        Class::Ascii => pick(random, 0x20, 0x7E),
        Class::Latin1 => pick(random, 0xA1, 0xFF),
        Class::Greek => pick(random, 0x03B1, 0x03C9),
        Class::Cjk => pick(random, 0x4E00, 0x9FA5),
        Class::Astral => *random.pick(&[
            '\u{1F600}',
            '\u{1F4A9}',
            '\u{1D11E}',
            '\u{20000}',
            '\u{2A6D6}',
            '\u{10348}',
        ]),
        Class::Hostile => *random.pick(&['\0', '\u{3}', ':', '\\', '/', '?', '*', '[', ']', '\'']),
    }
}

/// Random text of about `units` UTF-16 code units. A supplementary-plane
/// character drawn with one unit left makes the text one unit longer, so
/// targets near a limit also produce a surrogate pair straddling it.
fn text(random: &mut Random, units: usize, hostile: bool) -> String {
    let classes: &[Class] = if hostile {
        &[
            Class::Ascii,
            Class::Ascii,
            Class::Latin1,
            Class::Greek,
            Class::Cjk,
            Class::Astral,
            Class::Hostile,
        ]
    } else {
        &[
            Class::Ascii,
            Class::Ascii,
            Class::Latin1,
            Class::Greek,
            Class::Cjk,
            Class::Astral,
        ]
    };
    // Mostly one class, so ASCII-only and Latin-1-only strings occur too.
    let dominant = *random.pick(classes);
    let mut value = String::new();
    let mut length = 0;
    while length < units {
        let class = if random.chance(3, 4) {
            dominant
        } else {
            *random.pick(classes)
        };
        let character = character(random, class);
        length += character.len_utf16();
        value.push(character);
    }
    value
}

/// A length near `limit` (one of limit - 1, limit, limit + 1) or a short one.
fn length_near(random: &mut Random, limit: usize) -> usize {
    match random.below(6) {
        0 => limit - 1,
        1 => limit,
        2 => limit + 1,
        _ => 1 + random.below(limit.min(24)),
    }
}

/// A registration call; the second writer replays the accepted ones.
#[derive(Clone, Debug)]
enum Call {
    AddWorksheet(String),
    String(usize, u32, u16, String, u16),
    Number(usize, u32, u16, f64, u16),
    Boolean(usize, u32, u16, bool),
    Formula(usize, u32, u16, String),
    Hyperlink(usize, u32, u16, String),
    AutoFilter(usize, u16),
    FilterCondition(usize, u16, AutoFilterConditionWrite),
    CellStyle(CellStyle),
    NumberFormat(String),
    DefinedName(String, String),
}

/// What a successful call returned.
enum Accepted {
    Unit,
    Index(usize),
    Id(u16),
}

fn apply(writer: &mut Writer, call: &Call) -> Result<Accepted, Error> {
    Ok(match call.clone() {
        Call::AddWorksheet(name) => Accepted::Index(writer.add_worksheet(&name)?),
        Call::String(sheet, row, col, value, style) => {
            writer.write_string_with_format(sheet, row, col, &value, style)?;
            Accepted::Unit
        },
        Call::Number(sheet, row, col, value, style) => {
            writer.write_number_with_format(sheet, row, col, value, style)?;
            Accepted::Unit
        },
        Call::Boolean(sheet, row, col, value) => {
            writer.write_boolean(sheet, row, col, value)?;
            Accepted::Unit
        },
        Call::Formula(sheet, row, col, formula) => {
            writer.write_formula(sheet, row, col, &formula)?;
            Accepted::Unit
        },
        Call::Hyperlink(sheet, row, col, url) => {
            writer.set_hyperlink(sheet, row, col, &url)?;
            Accepted::Unit
        },
        Call::AutoFilter(sheet, last_col) => {
            writer.set_auto_filter(sheet, 0, 40, 0, last_col)?;
            Accepted::Unit
        },
        Call::FilterCondition(sheet, column, condition) => {
            writer.add_filter_condition(
                sheet,
                column,
                false,
                condition,
                AutoFilterConditionWrite::None,
            )?;
            Accepted::Unit
        },
        Call::CellStyle(style) => Accepted::Id(writer.add_cell_style(style)?),
        Call::NumberFormat(pattern) => Accepted::Id(writer.register_number_format(&pattern)?),
        Call::DefinedName(name, reference) => {
            writer.define_name(&name, &reference)?;
            Accepted::Unit
        },
    })
}

/// A refusal is one of the crate's typed errors, never a panic or an I/O
/// error.
fn assert_typed_refusal(call: &Call, error: &Error) {
    assert!(
        matches!(
            error,
            Error::StringTooLong { .. }
                | Error::TooMany { .. }
                | Error::RecordTooLong { .. }
                | Error::InvalidData(_)
                | Error::InvalidFormula(_)
                | Error::InvalidCellReference(_)
                | Error::InvalidFormat(_)
                | Error::WorksheetNotFound(_)
        ),
        "{call:?} was refused with an untyped error: {error:?}"
    );
}

/// What the reader must find.
#[derive(Default)]
struct Expected {
    sheets: Vec<String>,
    /// (sheet, row, col) -> value and style id.
    cells: BTreeMap<(usize, u32, u16), (ExpectedValue, u16)>,
    /// (sheet, row, col) -> the target as the reader reports it.
    hyperlinks: BTreeMap<(usize, u32, u16), Link>,
    /// sheet -> relative column -> condition.
    filters: BTreeMap<usize, BTreeMap<u16, FilterValue>>,
    styles: BTreeMap<u16, CellStyle>,
    names: Vec<String>,
}

#[derive(Clone, Debug)]
enum ExpectedValue {
    String(String),
    Number(f64),
    Boolean(bool),
    Formula(Vec<u8>),
}

#[derive(Clone, Debug, PartialEq)]
enum Link {
    Location(String),
    Url(String),
}

fn record(expected: &mut Expected, call: &Call, accepted: &Accepted) {
    match (call, accepted) {
        (Call::AddWorksheet(name), Accepted::Index(index)) => {
            assert_eq!(*index, expected.sheets.len());
            expected.sheets.push(name.clone());
        },
        (Call::String(sheet, row, col, value, style), _) => {
            expected.cells.insert(
                (*sheet, *row, *col),
                (ExpectedValue::String(value.clone()), *style),
            );
        },
        (Call::Number(sheet, row, col, value, style), _) => {
            expected.cells.insert(
                (*sheet, *row, *col),
                (ExpectedValue::Number(*value), *style),
            );
        },
        (Call::Boolean(sheet, row, col, value), _) => {
            expected
                .cells
                .insert((*sheet, *row, *col), (ExpectedValue::Boolean(*value), 0));
        },
        (Call::Formula(sheet, row, col, formula), _) => {
            let tokens = FormulaTokenizer::new().tokenize(formula).unwrap();
            let encoded = encode_ptg_tokens(&tokens).unwrap();
            expected
                .cells
                .insert((*sheet, *row, *col), (ExpectedValue::Formula(encoded), 0));
        },
        (Call::Hyperlink(sheet, row, col, url), _) => {
            let trimmed = url.trim();
            if trimmed.is_empty() {
                expected.hyperlinks.remove(&(*sheet, *row, *col));
            } else {
                let link = match trimmed.strip_prefix("internal:") {
                    Some(location) => Link::Location(location.to_string()),
                    None => Link::Url(trimmed.to_string()),
                };
                expected.hyperlinks.insert((*sheet, *row, *col), link);
            }
        },
        (Call::AutoFilter(sheet, _), _) => {
            expected.filters.entry(*sheet).or_default();
        },
        (Call::FilterCondition(sheet, column, condition), _) => {
            let value = match condition {
                AutoFilterConditionWrite::String { value, .. } => {
                    FilterValue::String(value.clone())
                },
                AutoFilterConditionWrite::Number { value, .. } => FilterValue::Number(*value),
                AutoFilterConditionWrite::Bool { value, .. } => FilterValue::Bool(*value),
                AutoFilterConditionWrite::MatchAll { .. } => FilterValue::MatchAll,
                AutoFilterConditionWrite::None => FilterValue::None,
            };
            expected
                .filters
                .get_mut(sheet)
                .unwrap()
                .insert(*column, value);
        },
        (Call::CellStyle(style), Accepted::Id(id)) => {
            expected.styles.insert(*id, style.clone());
        },
        (Call::DefinedName(name, _), _) => expected.names.push(name.clone()),
        _ => {},
    }
}

struct Generator<'r> {
    random: &'r mut Random,
    sheets: Vec<String>,
    styles: Vec<u16>,
    filters: Vec<(usize, u16)>,
}

impl Generator<'_> {
    fn sheet(&mut self) -> usize {
        self.random.below(self.sheets.len())
    }

    fn cell(&mut self) -> (u32, u16) {
        let row = u32::try_from(self.random.below(12)).unwrap();
        let col = u16::try_from(self.random.below(6)).unwrap();
        (row, col)
    }

    fn style(&mut self) -> u16 {
        if self.styles.is_empty() || self.random.chance(1, 2) {
            0
        } else {
            *self.random.pick(&self.styles)
        }
    }

    fn sheet_name(&mut self) -> String {
        let length = length_near(self.random, 31);
        let hostile = self.random.chance(1, 4);
        text(self.random, length, hostile)
    }

    fn formula(&mut self, depth: usize) -> String {
        let choice = if depth == 0 {
            self.random.below(4)
        } else {
            self.random.below(8)
        };
        match choice {
            0 => format!("{}", self.random.below(65_536)),
            1 => format!("{}.{}", self.random.below(1000), self.random.below(100)),
            2 => {
                let column = self.random.below(256);
                let letters = if column < 26 {
                    char::from(b'A' + u8::try_from(column).unwrap()).to_string()
                } else {
                    format!(
                        "{}{}",
                        char::from(b'A' + u8::try_from(column / 26 - 1).unwrap()),
                        char::from(b'A' + u8::try_from(column % 26).unwrap())
                    )
                };
                format!("{letters}{}", 1 + self.random.below(65_536))
            },
            3 => {
                let length = length_near(self.random, 255);
                let value = text(self.random, length, false).replace('"', "\"\"");
                format!("\"{value}\"")
            },
            4 => {
                let op = *self
                    .random
                    .pick(&["+", "-", "*", "/", "&", "=", "<", ">", "<=", ">=", "<>"]);
                format!("{}{op}{}", self.formula(depth - 1), self.formula(depth - 1))
            },
            5 => format!("({})", self.formula(depth - 1)),
            6 => {
                let function = *self
                    .random
                    .pick(&["SUM", "AVERAGE", "MIN", "MAX", "ABS", "LEN"]);
                format!("{function}({})", self.formula(depth - 1))
            },
            _ => {
                // Something the tokenizer refuses.
                (*self
                    .random
                    .pick(&["SUM(", "1+", "NOPE(1)", "(1", "\"open", "A1:"]))
                .to_string()
            },
        }
    }

    fn url(&mut self) -> String {
        let limit = if self.random.chance(1, 2) { 4093 } else { 4085 };
        let length = match self.random.below(4) {
            0 => limit - 1 - 9,
            1 => limit - 9,
            2 => limit + 1 - 9,
            _ => 1 + self.random.below(40),
        };
        let body = text(self.random, length, false).replace('\0', "");
        match self.random.below(4) {
            0 => {
                let sheet = self.random.pick(&self.sheets).clone();
                format!("internal:'{sheet}'!A1")
            },
            1 => format!("internal:{body}"),
            2 => format!("https://{body}"),
            _ => (*self
                .random
                .pick(&["", "  ", "https://x\0y", "mailto:a@b.c"]))
            .to_string(),
        }
    }

    fn filter_condition(&mut self) -> AutoFilterConditionWrite {
        let operator = u8::try_from(self.random.below(8)).unwrap();
        match self.random.below(4) {
            0 | 1 => {
                let length = length_near(self.random, 255);
                AutoFilterConditionWrite::String {
                    operator,
                    value: if self.random.chance(1, 12) {
                        String::new()
                    } else {
                        text(self.random, length, false)
                    },
                }
            },
            2 => AutoFilterConditionWrite::Number {
                operator,
                value: *self
                    .random
                    .pick(&[0.5, -3.0, 1e9, 0.0, -0.0, f64::NAN, 7.25]),
            },
            _ => AutoFilterConditionWrite::Bool {
                operator,
                value: self.random.chance(1, 2),
            },
        }
    }

    fn cell_style(&mut self) -> CellStyle {
        let name_length = length_near(self.random, 31);
        let name = if self.random.chance(1, 16) {
            String::new()
        } else {
            let hostile = self.random.chance(1, 8);
            text(self.random, name_length, hostile)
        };
        let number_format = match self.random.below(4) {
            0 => None,
            1 => Some((*self.random.pick(&["0.00", "@", "General", "0%"])).to_string()),
            _ => {
                let length = length_near(self.random, 255);
                Some(text(self.random, length, false))
            },
        };
        CellStyle {
            font: Font {
                height: *self.random.pick(&[0, 19, 20, 200, 240, 8191, 8192]),
                weight: *self.random.pick(&[0, 99, 100, 400, 700, 1000, 1001]),
                italic: self.random.chance(1, 2),
                underline: *self.random.pick(&[0x00, 0x01, 0x02, 0x21, 0x22, 0x03]),
                color_index: *self.random.pick(&[0x0008, 0x000A, 0x003F, 0x7FFF, 0x0042]),
                name,
            },
            h_align: *self.random.pick(&[
                HorizontalAlignment::General,
                HorizontalAlignment::Left,
                HorizontalAlignment::Center,
                HorizontalAlignment::Right,
            ]),
            v_align: *self.random.pick(&[
                VerticalAlignment::Top,
                VerticalAlignment::Center,
                VerticalAlignment::Bottom,
            ]),
            text_wrap: self.random.chance(1, 2),
            number_format,
            ..CellStyle::default()
        }
    }

    fn call(&mut self) -> Call {
        let (row, col) = self.cell();
        match self.random.below(20) {
            0 => Call::AddWorksheet(self.sheet_name()),
            1..=4 => {
                let length = if self.random.chance(1, 8) {
                    self.random.below(3000)
                } else {
                    self.random.below(40)
                };
                let sheet = self.sheet();
                let hostile = self.random.chance(1, 8);
                let value = text(self.random, length, hostile);
                let style = self.style();
                Call::String(sheet, row, col, value, style)
            },
            5 | 6 => {
                let sheet = self.sheet();
                let value = f64::from(u32::try_from(self.random.below(1_000_000)).unwrap()) / 8.0;
                let style = self.style();
                Call::Number(sheet, row, col, value, style)
            },
            7 => {
                let sheet = self.sheet();
                Call::Boolean(sheet, row, col, self.random.chance(1, 2))
            },
            8..=10 => {
                let sheet = self.sheet();
                let formula = self.formula(3);
                Call::Formula(sheet, row, col, formula)
            },
            11 | 12 => {
                let sheet = self.sheet();
                let url = self.url();
                Call::Hyperlink(sheet, row, col, url)
            },
            13 => {
                let sheet = self.sheet();
                // One AutoFilter range per worksheet: replacing a range keeps
                // the conditions of the old one, which this property does
                // not model.
                if self.filters.iter().any(|(filtered, _)| *filtered == sheet) {
                    return Call::Boolean(sheet, row, col, false);
                }
                let last = u16::try_from(1 + self.random.below(4)).unwrap();
                self.filters.push((sheet, last));
                Call::AutoFilter(sheet, last)
            },
            14 | 15 => match self.filters.last().copied() {
                Some((sheet, last)) => {
                    let column = u16::try_from(self.random.below(usize::from(last) + 1)).unwrap();
                    Call::FilterCondition(sheet, column, self.filter_condition())
                },
                None => Call::Boolean(0, row, col, true),
            },
            16 | 17 => Call::CellStyle(self.cell_style()),
            18 => {
                let length = length_near(self.random, 255);
                Call::NumberFormat(text(self.random, length, false))
            },
            _ => {
                let length = length_near(self.random, 255);
                let hostile = self.random.chance(1, 6);
                let name = text(self.random, length, hostile);
                Call::DefinedName(name, "A1:B2".to_string())
            },
        }
    }
}

fn write(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// Accepted and refused calls per kind, printed with `--nocapture`.
type Tally = BTreeMap<&'static str, (usize, usize)>;

const fn kind(call: &Call) -> &'static str {
    match call {
        Call::AddWorksheet(_) => "worksheet name",
        Call::String(..) => "string cell",
        Call::Number(..) => "number cell",
        Call::Boolean(..) => "boolean cell",
        Call::Formula(..) => "formula",
        Call::Hyperlink(..) => "hyperlink",
        Call::AutoFilter(..) => "AutoFilter range",
        Call::FilterCondition(..) => "AutoFilter condition",
        Call::CellStyle(_) => "cell style (font, format)",
        Call::NumberFormat(_) => "number format",
        Call::DefinedName(..) => "defined name",
    }
}

fn check_case(seed: u64, tally: &mut Tally) {
    let mut random = Random(seed | 1);
    let mut writer = Writer::new();
    let mut shadow = Writer::new();
    let mut expected = Expected::default();
    let mut generator = Generator {
        random: &mut random,
        sheets: Vec::new(),
        styles: Vec::new(),
        filters: Vec::new(),
    };
    // At least one worksheet, so every other call has a target.
    let first = Call::AddWorksheet(format!("S{seed:x}").chars().take(31).collect());
    let accepted = apply(&mut writer, &first).unwrap();
    apply(&mut shadow, &first).unwrap();
    record(&mut expected, &first, &accepted);
    generator.sheets.push(expected.sheets[0].clone());

    let calls = 30 + generator.random.below(50);
    for _ in 0..calls {
        let call = generator.call();
        let counts = tally.entry(kind(&call)).or_default();
        match apply(&mut writer, &call) {
            Ok(accepted) => {
                counts.0 += 1;
                let replayed = apply(&mut shadow, &call);
                assert!(
                    replayed.is_ok(),
                    "seed {seed:#x}: replay of {call:?} failed"
                );
                if let (Call::AddWorksheet(name), Accepted::Index(_)) = (&call, &accepted) {
                    generator.sheets.push(name.clone());
                }
                if let (Call::CellStyle(_), Accepted::Id(id)) = (&call, &accepted) {
                    generator.styles.push(*id);
                }
                record(&mut expected, &call, &accepted);
            },
            Err(error) => {
                counts.1 += 1;
                assert_typed_refusal(&call, &error);
                if let Call::AutoFilter(sheet, last) = call {
                    generator.filters.retain(|entry| *entry != (sheet, last));
                }
            },
        }
    }

    let bytes = write(&mut writer);
    assert_eq!(
        write(&mut writer),
        bytes,
        "seed {seed:#x}: two writes differ"
    );
    assert_eq!(
        write(&mut shadow),
        bytes,
        "seed {seed:#x}: a refused call changed the writer"
    );
    let workbook = Workbook::new(Cursor::new(bytes))
        .unwrap_or_else(|error| panic!("seed {seed:#x}: the reader refused the workbook: {error}"));
    compare(seed, &workbook, &expected);
}

fn compare(seed: u64, workbook: &Workbook, expected: &Expected) {
    let names = workbook
        .sheets()
        .iter()
        .map(|sheet| sheet.name().to_string())
        .collect::<Vec<_>>();
    assert_eq!(names, expected.sheets, "seed {seed:#x}: worksheet names");

    for ((sheet, row, col), (value, style)) in &expected.cells {
        let worksheet = workbook.xls_worksheet(*sheet).unwrap();
        let cell = worksheet
            .get_cell(*row, u32::from(*col))
            .unwrap_or_else(|| panic!("seed {seed:#x}: cell {sheet}/{row}/{col} is missing"));
        match value {
            ExpectedValue::String(text) => {
                assert_eq!(
                    cell.value(),
                    &ReadValue::String(text.clone()),
                    "seed {seed:#x}"
                );
            },
            ExpectedValue::Number(number) => match cell.value() {
                // A number under a date or time format reads as a serial date.
                ReadValue::Float(read) | ReadValue::DateTime(read) => {
                    assert_eq!(read, number, "seed {seed:#x}");
                },
                ReadValue::Int(read) => assert_eq!(*read as f64, *number, "seed {seed:#x}"),
                other => panic!("seed {seed:#x}: number read back as {other:?}"),
            },
            ExpectedValue::Boolean(flag) => {
                assert_eq!(cell.value(), &ReadValue::Bool(*flag), "seed {seed:#x}");
            },
            ExpectedValue::Formula(tokens) => {
                assert_eq!(
                    cell.formula_bytes(),
                    Some(tokens.as_slice()),
                    "seed {seed:#x}"
                );
            },
        }
        check_style(seed, workbook, cell.xf_index(), *style, expected);
    }

    for sheet in 0..expected.sheets.len() {
        let worksheet = workbook.xls_worksheet(sheet).unwrap();
        let mut links = worksheet
            .hyperlinks()
            .iter()
            .map(|link| {
                let range = link.range();
                assert_eq!(range.first_row(), range.last_row());
                assert_eq!(range.first_column(), range.last_column());
                let read = match link.location() {
                    Some(location) if link.moniker().is_none() => {
                        Link::Location(location.to_string())
                    },
                    _ => Link::Url(link.address().unwrap().to_string()),
                };
                (
                    (
                        sheet,
                        u32::from(range.first_row()),
                        u16::from(range.first_column()),
                    ),
                    read,
                )
            })
            .collect::<Vec<_>>();
        links.sort_by_key(|link| link.0);
        let wanted = expected
            .hyperlinks
            .iter()
            .filter(|((link_sheet, _, _), _)| *link_sheet == sheet)
            .map(|(key, link)| (*key, link.clone()))
            .collect::<Vec<_>>();
        assert_eq!(links, wanted, "seed {seed:#x}: hyperlinks of sheet {sheet}");

        let conditions = worksheet
            .autofilter()
            .map(|filter| {
                filter
                    .columns
                    .iter()
                    .map(|column| (column.column_index, column.condition1.value.clone()))
                    .collect::<BTreeMap<_, _>>()
            })
            .unwrap_or_default();
        let wanted = expected.filters.get(&sheet).cloned().unwrap_or_default();
        assert_eq!(
            conditions.len(),
            wanted.len(),
            "seed {seed:#x}: filters of sheet {sheet}"
        );
        for (column, value) in wanted {
            match (&value, &conditions[&column]) {
                (FilterValue::Number(left), FilterValue::Number(right)) => {
                    assert_eq!(left.to_bits(), right.to_bits(), "seed {seed:#x}");
                },
                (left, right) => assert_eq!(left, right, "seed {seed:#x}: filter {column}"),
            }
        }
    }

    let mut read_names = workbook
        .defined_names()
        .iter()
        .filter(|name| name.kind == DefinedNameKind::User && name.scope == NameScope::Workbook)
        .map(|name| name.name.clone())
        .collect::<Vec<_>>();
    let mut wanted = expected.names.clone();
    read_names.sort();
    wanted.sort();
    assert_eq!(read_names, wanted, "seed {seed:#x}: defined names");
}

fn check_style(seed: u64, workbook: &Workbook, xf_index: u16, style: u16, expected: &Expected) {
    if style == 0 {
        assert_eq!(xf_index, 15, "seed {seed:#x}: an unstyled cell uses XF 15");
        return;
    }
    // Style `id` is the id-th user cell XF, after the 21 fixed ones.
    assert_eq!(
        usize::from(xf_index),
        20 + usize::from(style),
        "seed {seed:#x}"
    );
    let wanted = &expected.styles[&style];
    let xf = &workbook.formatting().extended_formats()[usize::from(xf_index)];
    let font = workbook
        .fonts()
        .iter()
        .find(|font| font.index() == xf.font_index())
        .unwrap_or_else(|| panic!("seed {seed:#x}: XF font {} is missing", xf.font_index()));
    assert_eq!(font.name(), wanted.font.name, "seed {seed:#x}");
    assert_eq!(font.height_twips(), wanted.font.height, "seed {seed:#x}");
    assert_eq!(font.weight(), wanted.font.weight, "seed {seed:#x}");
    assert_eq!(font.is_italic(), wanted.font.italic, "seed {seed:#x}");
    assert_eq!(
        font.color_index(),
        wanted.font.color_index,
        "seed {seed:#x}"
    );
    let underline = match wanted.font.underline {
        0x00 => FontUnderline::None,
        0x01 => FontUnderline::Single,
        0x02 => FontUnderline::Double,
        0x21 => FontUnderline::SingleAccounting,
        _ => FontUnderline::DoubleAccounting,
    };
    assert_eq!(font.underline(), underline, "seed {seed:#x}");
    let format_id = xf.number_format_id();
    // A pattern a fresh manager maps below 164 is a built-in format, which
    // has no Format record.
    let builtin = wanted.number_format.as_deref().map_or(Some(0), |pattern| {
        FormattingManager::new()
            .register_number_format(pattern)
            .ok()
            .filter(|id| *id < 164)
    });
    match (builtin, wanted.number_format.as_deref()) {
        (Some(id), _) => assert_eq!(format_id, id, "seed {seed:#x}"),
        (None, None) => unreachable!(),
        (None, Some(pattern)) => {
            let format = workbook
                .formatting()
                .number_formats()
                .iter()
                .find(|format| format.id() == format_id)
                .unwrap_or_else(|| panic!("seed {seed:#x}: number format {format_id} is missing"));
            assert_eq!(format.code(), pattern, "seed {seed:#x}");
        },
    }
    let alignment = xf.alignment();
    let horizontal = match wanted.h_align {
        HorizontalAlignment::General => litchi_xls::HorizontalAlignment::General,
        HorizontalAlignment::Left => litchi_xls::HorizontalAlignment::Left,
        HorizontalAlignment::Center => litchi_xls::HorizontalAlignment::Center,
        HorizontalAlignment::Right => litchi_xls::HorizontalAlignment::Right,
        other => panic!("unexpected generated alignment {other:?}"),
    };
    assert_eq!(alignment.horizontal(), horizontal, "seed {seed:#x}");
    let vertical = match wanted.v_align {
        VerticalAlignment::Top => litchi_xls::VerticalAlignment::Top,
        VerticalAlignment::Center => litchi_xls::VerticalAlignment::Center,
        VerticalAlignment::Bottom => litchi_xls::VerticalAlignment::Bottom,
        VerticalAlignment::Justify => litchi_xls::VerticalAlignment::Justify,
    };
    assert_eq!(alignment.vertical(), vertical, "seed {seed:#x}");
    assert_eq!(alignment.wraps_text(), wanted.text_wrap, "seed {seed:#x}");
}

#[test]
fn random_workbooks_round_trip_through_the_reader() {
    let cases = std::env::var("LITCHI_XLS_PROPERTY_CASES")
        .ok()
        .and_then(|value| value.parse::<u64>().ok())
        .unwrap_or(64);
    let base = std::env::var("LITCHI_XLS_PROPERTY_SEED")
        .ok()
        .and_then(|value| u64::from_str_radix(value.trim_start_matches("0x"), 16).ok())
        .unwrap_or(0x0766_5eed_0000_0000);
    let mut tally = Tally::new();
    for case in 0..cases {
        check_case(
            base.wrapping_add(case.wrapping_mul(0x9E37_79B9_7F4A_7C15)),
            &mut tally,
        );
    }
    println!("{cases} random workbooks from seed {base:#x}; calls accepted / refused:");
    for (kind, (accepted, refused)) in &tally {
        println!("  {kind}: {accepted} / {refused}");
    }
}

/// Inputs the generator reaches only by chance, pinned: a limit reached
/// exactly by a surrogate pair, and one crossed by it.
#[test]
fn surrogate_pairs_at_each_limit_round_trip_or_are_refused() {
    let at = |limit: usize| format!("{}\u{1F600}", "a".repeat(limit - 2));
    let past = |limit: usize| format!("{}\u{1F600}", "a".repeat(limit - 1));
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet(&at(31)).unwrap();
    assert!(matches!(
        writer.add_worksheet(&past(31)),
        Err(Error::StringTooLong { .. })
    ));
    writer.write_string(sheet, 0, 0, &at(20)).unwrap();
    writer
        .write_formula(sheet, 1, 0, &format!("\"{}\"", at(255)))
        .unwrap();
    assert!(matches!(
        writer.write_formula(sheet, 1, 1, &format!("\"{}\"", past(255))),
        Err(Error::StringTooLong { .. })
    ));
    writer
        .set_hyperlink(sheet, 2, 0, &format!("internal:{}", at(4093)))
        .unwrap();
    writer
        .set_hyperlink(sheet, 2, 1, &format!("https://{}", at(4085 - 8)))
        .unwrap();
    writer.set_auto_filter(sheet, 0, 9, 0, 1).unwrap();
    writer
        .add_filter_condition(
            sheet,
            0,
            false,
            AutoFilterConditionWrite::String {
                operator: 2,
                value: at(255),
            },
            AutoFilterConditionWrite::None,
        )
        .unwrap();
    let style = writer
        .add_cell_style(CellStyle {
            font: Font {
                name: at(31),
                ..Font::default()
            },
            number_format: Some(at(255)),
            ..CellStyle::default()
        })
        .unwrap();
    writer
        .write_number_with_format(sheet, 3, 0, 1.5, style)
        .unwrap();
    writer.define_name(&at(255), "A1").unwrap();

    let workbook = Workbook::new(Cursor::new(write(&mut writer))).unwrap();
    assert_eq!(workbook.sheets()[0].name(), at(31));
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert_eq!(
        worksheet.get_cell(0, 0).unwrap().value(),
        &ReadValue::String(at(20))
    );
    let links = worksheet.hyperlinks();
    assert_eq!(links[0].location(), Some(at(4093).as_str()));
    assert_eq!(
        links[1].address(),
        Some(format!("https://{}", at(4085 - 8)).as_str())
    );
    assert_eq!(
        worksheet.autofilter().unwrap().columns[0].condition1.value,
        FilterValue::String(at(255))
    );
    let xf = &workbook.formatting().extended_formats()
        [usize::from(worksheet.get_cell(3, 0).unwrap().xf_index())];
    assert_eq!(workbook.fonts().last().unwrap().name(), at(31));
    assert_eq!(
        workbook
            .formatting()
            .number_formats()
            .iter()
            .find(|format| format.id() == xf.number_format_id())
            .unwrap()
            .code(),
        at(255)
    );
    assert!(
        workbook
            .defined_names()
            .iter()
            .any(|name| name.name == at(255))
    );
}
