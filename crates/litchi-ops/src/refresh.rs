//! Field and table-of-contents refresh (design §13 P0, record 0279).
//!
//! [`File::refresh`](crate::File::refresh) recomputes the derived results a
//! document caches: tables of contents and computable fields in DOCX,
//! computable fields in DOC, ODT, OTH and ODM, slide numbers and dates in
//! PPTX and ODP, and formula results in XLSX, XLSB and XLS (records 0279
//! and 0295). Every field the
//! refresh meets is reported in the [`Summary`] with a [`Status`]; the
//! changed results are published in one atomic unified operation (XLS may
//! use several native chunks). A document whose supported results are all
//! current is an exact no-op. Unsupported fields retain their saved results.
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! use litchi_core::FileFormat;
//! use litchi_ops::refresh::{Clock, Options, Status};
//! use litchi_ops::{NewNode, create};
//!
//! let file = create(FileFormat::Docx)?
//!     .add("/body", NewNode::new("paragraph").prop("text", "Report"))?
//!     .file;
//! // Nothing to refresh: the file comes back unchanged.
//! let commit = file.refresh(&Options::new().clock(Clock::parse("2026-10-14T09:30:00")?))?;
//! assert!(commit.patch().is_noop());
//! assert!(commit.value().fields().is_empty());
//! # }
//! # Ok::<(), litchi_ops::Error>(())
//! ```
//!
//! # Litchi has no layout engine
//!
//! Page numbers depend on pagination. `PAGE`, `NUMPAGES`, `SECTIONPAGES`
//! and `PAGEREF` fields keep their cached results and are reported as
//! [`Status::NeedsLayout`]. A table of contents is rebuilt from the
//! document's headings; each entry's page number is the cached number of the
//! old entry with the same bookmark (or the same level and text), and an
//! entry without one gets an empty page number, counted in
//! [`Summary::missing_page_numbers`]. Such a table is marked dirty (unless
//! [`Options::mark_dirty`] is off) so a word processor recomputes it on
//! open. [`PageNumbers::Require`] refuses such a refresh with
//! [`Code::RefreshFailed`] instead.
//!
//! # No ambient clock
//!
//! `DATE`, `TIME` and PPTX date fields are refreshed only with an explicit
//! [`Options::clock`]; without one they are reported as
//! [`Status::NeedsClock`]. The system clock is never read.
//!
//! # DOCX fields
//!
//! | Field | Result |
//! |---|---|
//! | `TOC` with `\o`, `\u` or `\t` | one `TOC1`..`TOC9` paragraph per heading of the main story (style outline levels, `HeadingN` styles, direct outline levels with `\u`, custom styles with `\t`), `\h` links, `\n` page-number omission, `\p` separators; headings gain `_Toc` bookmarks as needed |
//! | `SEQ` | document-order numbering per identifier: `\n`, `\c`, `\r N`, `\h`, `\* ARABIC/ROMAN/roman/ALPHABETIC/alphabetic` |
//! | `REF` | the bookmark's text, with `\* Upper/Lower/Caps/FirstCap` |
//! | `DATE`, `TIME` | the clock, with `\@` date pictures |
//! | `CREATEDATE`, `SAVEDATE`, `PRINTDATE` | the core properties' times, with `\@` |
//! | `TITLE`, `SUBJECT`, `AUTHOR`, `KEYWORDS`, `COMMENTS`, `LASTSAVEDBY`, `DOCPROPERTY` | the document properties |
//! | `NUMWORDS`, `NUMCHARS` | counted over the main story's text |
//! | `PAGE`, `NUMPAGES`, `SECTIONPAGES`, `PAGEREF` | kept ([`Status::NeedsLayout`]) |
//!
//! Locked fields are reported [`Status::Locked`]; fields nested in other
//! fields (except the `PAGEREF`s a rebuilt table of contents replaces), in
//! text boxes, with numeric pictures (`\#`) or other switches are
//! [`Status::Unsupported`]. Only the main story is refreshed.
//!
//! # Other formats (record 0295)
//!
//! | Format | Fields |
//! |---|---|
//! | DOC | the main story's top-level fields with a result, computed as DOCX computes them (`REF` is `unsupported`), rewritten through the native field-result edit; a story with a nested field or a field without a result reports every computable field `unsupported` |
//! | ODT, OTH, ODM | the body's `text:date`/`text:time` (clock), stored-time and document-property fields; page fields are [`Status::NeedsLayout`], fixed fields [`Status::Locked`], fields with an ODF data style `unsupported` |
//! | ODP | as ODT, and `text:page-number` (the slide position) and `text:page-count` (the number of slides) on every slide and notes page |
//! | XLSB, XLS | formula results as XLSX, written back with [`Options::write_cache`] as the formula records' cached results |
//!
//! PPT and RTF have no refresh adapter yet. Changed PPTX fields inherit the
//! native text editor's markup profile: fields on slides requiring unsupported
//! markup-compatibility branch selection refuse atomically, while unchanged
//! fields can still be reported without a write.
//!
//! # Limits
//!
//! The transaction's [`SemanticLimits`](crate::transaction::SemanticLimits):
//! the changed results are bounded by `max_expanded_targets`, and their text
//! by `max_staged_bytes`.

#![cfg_attr(
    not(feature = "docx"),
    allow(dead_code, reason = "the DOCX adapter plans most fields")
)]

use std::fmt;

use crate::error::{Code, Error, Result};

/// An explicit date and time for date fields, in the document's local
/// time. Litchi never reads the system clock.
///
/// ```
/// use litchi_ops::refresh::Clock;
///
/// let clock = Clock::parse("2026-10-14T09:05:00")?;
/// assert_eq!(clock.format("dddd, MMMM d, yyyy h:mm am/pm"), "Wednesday, October 14, 2026 9:05 am");
/// assert!(Clock::parse("2026-02-30").is_err());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Clock {
    year: i32,
    month: u8,
    day: u8,
    hour: u8,
    minute: u8,
    second: u8,
}

const MONTHS: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];
const DAYS: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

const fn days_in_month(year: i32, month: u8) -> u8 {
    match month {
        2 if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

impl Clock {
    /// A clock at the given local date and time.
    ///
    /// # Errors
    ///
    /// [`Code::InvalidValue`] for an impossible date or time, or a year
    /// outside 1..=9999.
    pub fn new(year: i32, month: u8, day: u8, hour: u8, minute: u8, second: u8) -> Result<Self> {
        if !(1..=9999).contains(&year)
            || !(1..=12).contains(&month)
            || day == 0
            || day > days_in_month(year, month)
            || hour > 23
            || minute > 59
            || second > 59
        {
            return Err(
                Error::new(Code::InvalidValue, "the clock is not a valid date and time")
                    .with_suggestion("give the clock as YYYY-MM-DD or YYYY-MM-DDTHH:MM[:SS]"),
            );
        }
        Ok(Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
        })
    }

    /// Parse `YYYY-MM-DD`, optionally followed by `T` or a space and
    /// `HH:MM` or `HH:MM:SS`. A zone suffix is not accepted.
    ///
    /// # Errors
    ///
    /// [`Code::InvalidValue`] for any other spelling or an impossible date.
    pub fn parse(text: &str) -> Result<Self> {
        let invalid = || {
            Error::new(
                Code::InvalidValue,
                "the clock is not YYYY-MM-DD[THH:MM[:SS]]",
            )
            .with_suggestion("give the clock as YYYY-MM-DD or YYYY-MM-DDTHH:MM[:SS]")
        };
        let number = |part: &str, digits: usize| -> Result<u32> {
            if part.len() != digits || !part.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid());
            }
            part.parse().map_err(|_| invalid())
        };
        let text = text.trim();
        let (date, time) = match text.find(['T', ' ']) {
            Some(at) => (&text[..at], Some(&text[at + 1..])),
            None => (text, None),
        };
        let mut date_parts = date.split('-');
        let (Some(year), Some(month), Some(day), None) = (
            date_parts.next(),
            date_parts.next(),
            date_parts.next(),
            date_parts.next(),
        ) else {
            return Err(invalid());
        };
        let (mut hour, mut minute, mut second) = (0, 0, 0);
        if let Some(time) = time {
            let mut parts = time.split(':');
            hour = number(parts.next().ok_or_else(invalid)?, 2)?;
            minute = number(parts.next().ok_or_else(invalid)?, 2)?;
            if let Some(seconds) = parts.next() {
                second = number(seconds, 2)?;
            }
            if parts.next().is_some() {
                return Err(invalid());
            }
        }
        let narrow = |value: u32| u8::try_from(value).map_err(|_| invalid());
        Self::new(
            i32::try_from(number(year, 4)?).map_err(|_| invalid())?,
            narrow(number(month, 2)?)?,
            narrow(number(day, 2)?)?,
            narrow(hour)?,
            narrow(minute)?,
            narrow(second)?,
        )
    }

    /// The day of the week, 0 for Sunday.
    fn weekday(self) -> usize {
        // Sakamoto's method.
        const OFFSETS: [i32; 12] = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
        let year = if self.month < 3 {
            self.year - 1
        } else {
            self.year
        };
        let day = (year + year / 4 - year / 100
            + year / 400
            + OFFSETS[usize::from(self.month - 1)]
            + i32::from(self.day))
            % 7;
        usize::try_from(day).unwrap_or_default()
    }

    /// Format with a Word date-time picture (the `\@` switch): `d`, `dd`,
    /// `ddd`, `dddd`, `M`, `MM`, `MMM`, `MMMM`, `yy`, `yyyy`, `h`, `hh`,
    /// `H`, `HH`, `m`, `mm`, `s`, `ss`, `am/pm`, `AM/PM`, and quoted text.
    /// Names are English.
    #[must_use]
    pub fn format(self, picture: &str) -> String {
        let chars: Vec<char> = picture.chars().collect();
        let mut output = String::new();
        let mut index = 0;
        let twelve = picture.to_ascii_lowercase().contains("am/pm");
        while index < chars.len() {
            let character = chars[index];
            let run = chars[index..]
                .iter()
                .take_while(|next| **next == character)
                .count();
            let rest: String = chars[index..].iter().take(5).collect();
            if rest.eq_ignore_ascii_case("am/pm") {
                let upper = rest.starts_with('A');
                let text = if self.hour < 12 { "am" } else { "pm" };
                output.push_str(&if upper {
                    text.to_ascii_uppercase()
                } else {
                    text.to_owned()
                });
                index += 5;
                continue;
            }
            match character {
                '\'' => {
                    index += 1;
                    while index < chars.len() && chars[index] != '\'' {
                        output.push(chars[index]);
                        index += 1;
                    }
                    index += 1;
                    continue;
                },
                'd' | 'D' => match run {
                    1 => output.push_str(&self.day.to_string()),
                    2 => output.push_str(&format!("{:02}", self.day)),
                    3 => output.push_str(&DAYS[self.weekday()][..3]),
                    _ => output.push_str(DAYS[self.weekday()]),
                },
                'M' => {
                    let name = MONTHS[usize::from(self.month - 1)];
                    match run {
                        1 => output.push_str(&self.month.to_string()),
                        2 => output.push_str(&format!("{:02}", self.month)),
                        3 => output.push_str(&name[..3]),
                        _ => output.push_str(name),
                    }
                },
                'y' | 'Y' => {
                    if run <= 2 {
                        output.push_str(&format!("{:02}", self.year % 100));
                    } else {
                        output.push_str(&format!("{:04}", self.year));
                    }
                },
                'h' | 'H' => {
                    let hour = if character == 'h' && twelve {
                        match self.hour % 12 {
                            0 => 12,
                            hour => hour,
                        }
                    } else {
                        self.hour
                    };
                    if run >= 2 {
                        output.push_str(&format!("{hour:02}"));
                    } else {
                        output.push_str(&hour.to_string());
                    }
                },
                'm' => {
                    if run >= 2 {
                        output.push_str(&format!("{:02}", self.minute));
                    } else {
                        output.push_str(&self.minute.to_string());
                    }
                },
                's' | 'S' => {
                    if run >= 2 {
                        output.push_str(&format!("{:02}", self.second));
                    } else {
                        output.push_str(&self.second.to_string());
                    }
                },
                _ => {
                    for _ in 0..run {
                        output.push(character);
                    }
                },
            }
            index += run;
        }
        output
    }

    /// A clock from a stored timestamp's `YYYY-MM-DD HH:MM:SS` rendering.
    pub(crate) fn from_timestamp(timestamp: litchi_core::metadata::Timestamp) -> Option<Self> {
        let text = timestamp.naive().to_string();
        let text = text.get(..19).unwrap_or(&text);
        Self::parse(text).ok()
    }
}

impl std::str::FromStr for Clock {
    type Err = Error;

    fn from_str(text: &str) -> Result<Self> {
        Self::parse(text)
    }
}

/// What a table of contents entry whose page number Litchi cannot know
/// gets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum PageNumbers {
    /// Keep the old entry's cached number where one matches; leave the
    /// others empty and mark the table dirty (the default).
    #[default]
    Keep,
    /// Refuse the refresh with [`Code::RefreshFailed`] when any page number
    /// is unknown.
    Require,
}

/// Refresh options: no clock, cached page numbers kept, dirty marking on,
/// no formula cache write-back.
///
/// ```
/// use litchi_ops::refresh::{Options, PageNumbers};
///
/// let options = Options::new().page_numbers(PageNumbers::Require).write_cache(true);
/// assert_eq!(options.page_number_policy(), PageNumbers::Require);
/// assert!(options.writes_cache() && options.marks_dirty());
/// ```
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Options {
    clock: Option<Clock>,
    page_numbers: PageNumbers,
    mark_dirty: bool,
    write_cache: bool,
    #[cfg(all(
        feature = "eval",
        any(feature = "xlsx", feature = "xlsb", feature = "xls")
    ))]
    calc: crate::calc::Options,
}

impl Default for Options {
    fn default() -> Self {
        Self::new()
    }
}

impl Options {
    /// The defaults.
    #[must_use]
    pub fn new() -> Self {
        Self {
            clock: None,
            page_numbers: PageNumbers::Keep,
            mark_dirty: true,
            write_cache: false,
            #[cfg(all(
                feature = "eval",
                any(feature = "xlsx", feature = "xlsb", feature = "xls")
            ))]
            calc: crate::calc::Options::new(),
        }
    }

    /// The date and time `DATE`, `TIME` and PPTX date fields show.
    #[must_use]
    pub const fn clock(mut self, clock: Clock) -> Self {
        self.clock = Some(clock);
        self
    }

    /// The table of contents page-number policy.
    #[must_use]
    pub const fn page_numbers(mut self, policy: PageNumbers) -> Self {
        self.page_numbers = policy;
        self
    }

    /// Whether a table of contents with unknown page numbers is marked
    /// dirty for its word processor (on by default).
    #[must_use]
    pub const fn mark_dirty(mut self, mark: bool) -> Self {
        self.mark_dirty = mark;
        self
    }

    /// Whether an XLS/XLSX/XLSB refresh writes recalculated formula results back as
    /// the cells' cached values (off by default: a workbook refresh then
    /// only reports).
    #[must_use]
    pub const fn write_cache(mut self, write: bool) -> Self {
        self.write_cache = write;
        self
    }

    /// The formula-evaluation bounds of a workbook refresh.
    #[cfg(all(
        feature = "eval",
        any(feature = "xlsx", feature = "xlsb", feature = "xls")
    ))]
    #[must_use]
    pub const fn calc(mut self, options: crate::calc::Options) -> Self {
        self.calc = options;
        self
    }

    /// The clock, if any.
    #[must_use]
    pub const fn clock_value(&self) -> Option<Clock> {
        self.clock
    }

    /// The page-number policy.
    #[must_use]
    pub const fn page_number_policy(&self) -> PageNumbers {
        self.page_numbers
    }

    /// Whether dirty marking is on.
    #[must_use]
    pub const fn marks_dirty(&self) -> bool {
        self.mark_dirty
    }

    /// Whether formula caches are written back.
    #[must_use]
    pub const fn writes_cache(&self) -> bool {
        self.write_cache
    }

    #[cfg(all(
        feature = "eval",
        any(feature = "xlsx", feature = "xlsb", feature = "xls")
    ))]
    pub(crate) const fn calc_options(&self) -> &crate::calc::Options {
        &self.calc
    }
}

/// What a refresh did with one field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Status {
    /// The result was recomputed and written.
    Refreshed,
    /// The result was recomputed and equals the cached one.
    Unchanged,
    /// A workbook formula's computed result differs from its cache, which
    /// was not written ([`Options::write_cache`] is off).
    Stale,
    /// The result depends on pagination; the cached one is kept.
    NeedsLayout,
    /// The result depends on the clock and no [`Options::clock`] was given;
    /// the cached one is kept.
    NeedsClock,
    /// The field is locked against updates.
    Locked,
    /// The field refers to something the document lacks (a bookmark, a
    /// property); the cached result is kept.
    Unresolved,
    /// The field type, a switch or its markup is outside what Litchi
    /// computes; the cached result is kept.
    Unsupported,
}

impl Status {
    /// The status name (`refreshed`, `needs_layout`, ...).
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Refreshed => "refreshed",
            Self::Unchanged => "unchanged",
            Self::Stale => "stale",
            Self::NeedsLayout => "needs_layout",
            Self::NeedsClock => "needs_clock",
            Self::Locked => "locked",
            Self::Unresolved => "unresolved",
            Self::Unsupported => "unsupported",
        }
    }
}

impl fmt::Display for Status {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// One field of a refresh report.
#[derive(Clone, PartialEq, Eq)]
pub struct Field {
    kind: String,
    location: Option<String>,
    status: Status,
    reason: Option<&'static str>,
}

impl Field {
    pub(crate) fn new(kind: impl Into<String>, location: Option<String>, status: Status) -> Self {
        Self {
            kind: kind.into(),
            location,
            status,
            reason: None,
        }
    }

    pub(crate) const fn because(mut self, reason: &'static str) -> Self {
        self.reason = Some(reason);
        self
    }

    #[cfg(all(
        feature = "eval",
        any(feature = "xlsx", feature = "xlsb", feature = "xls")
    ))]
    pub(crate) const fn with_status(mut self, status: Status, reason: &'static str) -> Self {
        self.status = status;
        self.reason = Some(reason);
        self
    }

    /// The field type in upper case: `TOC`, `SEQ`, `PAGE`, `SLIDENUM`,
    /// `DATETIME1`, `FORMULA`, ...
    #[must_use]
    pub fn kind(&self) -> &str {
        &self.kind
    }

    /// Where the field is, when it has an address: `/body/p[3]` (a direct
    /// body paragraph), `/slide[2]`, `Sheet1!B4`.
    #[must_use]
    pub fn location(&self) -> Option<&str> {
        self.location.as_deref()
    }

    /// What the refresh did.
    #[must_use]
    pub const fn status(&self) -> Status {
        self.status
    }

    /// Why a field was not refreshed, when there is more to say than its
    /// status.
    #[must_use]
    pub const fn reason(&self) -> Option<&'static str> {
        self.reason
    }
}

impl fmt::Debug for Field {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Field")
            .field("kind", &self.kind)
            .field("status", &self.status)
            .finish_non_exhaustive()
    }
}

/// What one refresh did, the value of the returned
/// [`Commit`](crate::transaction::Commit).
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Summary {
    pub(crate) fields: Vec<Field>,
    pub(crate) bookmarks: usize,
    pub(crate) missing_page_numbers: usize,
}

impl Summary {
    /// A summary of `fields` without bookmarks or missing page numbers.
    #[cfg_attr(
        not(any(
            all(feature = "eval", any(feature = "xlsb", feature = "xls")),
            feature = "doc",
            feature = "odt",
            feature = "odp"
        )),
        allow(
            dead_code,
            reason = "used by the adapters that build their report as a list"
        )
    )]
    pub(crate) const fn from_fields(fields: Vec<Field>) -> Self {
        Self {
            fields,
            bookmarks: 0,
            missing_page_numbers: 0,
        }
    }

    /// Every field met, in document order.
    #[must_use]
    pub fn fields(&self) -> &[Field] {
        &self.fields
    }

    /// The fields with `status`.
    #[must_use]
    pub fn count(&self, status: Status) -> usize {
        self.fields
            .iter()
            .filter(|field| field.status == status)
            .count()
    }

    /// Fields whose result was written.
    #[must_use]
    pub fn refreshed(&self) -> usize {
        self.count(Status::Refreshed)
    }

    /// Bookmarks added for table of contents entries.
    #[must_use]
    pub const fn bookmarks(&self) -> usize {
        self.bookmarks
    }

    /// Table of contents entries left without a page number.
    #[must_use]
    pub const fn missing_page_numbers(&self) -> usize {
        self.missing_page_numbers
    }
}

impl fmt::Debug for Summary {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Summary")
            .field("fields", &self.fields.len())
            .field("refreshed", &self.refreshed())
            .field("bookmarks", &self.bookmarks)
            .field("missing_page_numbers", &self.missing_page_numbers)
            .finish()
    }
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

/// A field instruction split into its type, arguments and switches.
#[derive(Debug, Default, PartialEq, Eq)]
pub(crate) struct Instruction {
    pub(crate) kind: String,
    pub(crate) arguments: Vec<String>,
    pub(crate) switches: Vec<(char, Option<String>)>,
}

impl Instruction {
    /// Split `text`: tokens are separated by white space; double quotes
    /// group a token (`\"` inside one is a quote); a switch is `\` and one
    /// character, taking the next token as its argument unless that token
    /// is another switch. The format switches `\*`, `\#` and `\@` always
    /// take one.
    pub(crate) fn parse(text: &str) -> Self {
        let mut tokens: Vec<(String, bool)> = Vec::new();
        let mut chars = text.chars().peekable();
        while let Some(&character) = chars.peek() {
            if character.is_whitespace() {
                chars.next();
                continue;
            }
            let mut token = String::new();
            if character == '"' {
                chars.next();
                while let Some(next) = chars.next() {
                    match next {
                        '\\' if chars.peek() == Some(&'"') => {
                            token.push('"');
                            chars.next();
                        },
                        '"' => break,
                        _ => token.push(next),
                    }
                }
                tokens.push((token, true));
            } else {
                while let Some(&next) = chars.peek() {
                    if next.is_whitespace() || (next == '"' && !token.is_empty()) {
                        break;
                    }
                    token.push(next);
                    chars.next();
                }
                tokens.push((token, false));
            }
        }
        let mut instruction = Self::default();
        let mut iter = tokens.into_iter().peekable();
        if let Some((kind, _)) = iter.next() {
            instruction.kind = kind.to_ascii_uppercase();
        }
        while let Some((token, quoted)) = iter.next() {
            let switch = (!quoted && token.starts_with('\\') && token.chars().count() == 2)
                .then(|| token.chars().nth(1))
                .flatten();
            match switch {
                Some(name) => {
                    let format = matches!(name, '*' | '#' | '@');
                    let argument = match iter.peek() {
                        Some((next, quoted)) if format || *quoted || !next.starts_with('\\') => {
                            let argument = next.clone();
                            iter.next();
                            Some(argument)
                        },
                        _ => None,
                    };
                    instruction
                        .switches
                        .push((name.to_ascii_lowercase(), argument));
                },
                None => instruction.arguments.push(token),
            }
        }
        instruction
    }

    pub(crate) fn has(&self, name: char) -> bool {
        self.switches.iter().any(|(switch, _)| *switch == name)
    }

    pub(crate) fn switch(&self, name: char) -> Option<&str> {
        self.switches
            .iter()
            .find(|(switch, _)| *switch == name)
            .and_then(|(_, argument)| argument.as_deref())
    }

    /// The `\*` arguments that are not `MERGEFORMAT`/`CHARFORMAT`.
    pub(crate) fn formats(&self) -> impl Iterator<Item = &str> {
        self.switches
            .iter()
            .filter(|(switch, _)| *switch == '*')
            .filter_map(|(_, argument)| argument.as_deref())
            .filter(|argument| {
                !argument.eq_ignore_ascii_case("mergeformat")
                    && !argument.eq_ignore_ascii_case("charformat")
            })
    }
}

/// Apply a `\*` case format; `None` for any other format.
pub(crate) fn apply_case(value: &str, format: &str) -> Option<String> {
    Some(match format.to_ascii_lowercase().as_str() {
        "upper" => value.to_uppercase(),
        "lower" => value.to_lowercase(),
        "caps" => {
            let mut output = String::with_capacity(value.len());
            let mut start = true;
            for character in value.chars() {
                if start && character.is_alphabetic() {
                    output.extend(character.to_uppercase());
                } else {
                    output.extend(character.to_lowercase());
                }
                start = character.is_whitespace();
            }
            output
        },
        "firstcap" => {
            let mut chars = value.chars();
            chars.next().map_or_else(String::new, |first| {
                first.to_uppercase().chain(chars).collect()
            })
        },
        _ => return None,
    })
}

/// Format `number` with a `\*` numbering format; `None` for an unknown
/// format.
pub(crate) fn format_number(number: u64, format: &str) -> Option<String> {
    match format {
        "ARABIC" | "Arabic" | "arabic" => Some(number.to_string()),
        "ROMAN" => roman(number),
        "roman" => roman(number).map(|text| text.to_lowercase()),
        "ALPHABETIC" => alphabetic(number),
        "alphabetic" => alphabetic(number).map(|text| text.to_lowercase()),
        _ => None,
    }
}

fn roman(mut number: u64) -> Option<String> {
    if number == 0 || number > 32_767 {
        return None;
    }
    const DIGITS: [(u64, &str); 13] = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut output = String::new();
    for (value, digits) in DIGITS {
        while number >= value {
            output.push_str(digits);
            number -= value;
        }
    }
    Some(output)
}

/// Word's alphabetic numbering: A..Z, then AA..ZZ, AAA...
fn alphabetic(number: u64) -> Option<String> {
    if number == 0 || number > 780 {
        return None;
    }
    let letter = char::from(b'A' + u8::try_from((number - 1) % 26).ok()?);
    let repeat = usize::try_from((number - 1) / 26 + 1).ok()?;
    Some(std::iter::repeat_n(letter, repeat).collect())
}

/// `A`, `B`, ..., `AA` for a zero-based column.
#[cfg(all(
    feature = "eval",
    any(feature = "xlsx", feature = "xlsb", feature = "xls")
))]
pub(crate) fn column_letters(mut column: u32) -> String {
    let mut letters = Vec::new();
    loop {
        letters.push(char::from(
            b'A' + u8::try_from(column % 26).unwrap_or_default(),
        ));
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    letters.iter().rev().collect()
}

/// One workbook formula's report entry. `written` says whether a stale
/// result was written back.
#[cfg(all(
    feature = "eval",
    any(feature = "xlsx", feature = "xlsb", feature = "xls")
))]
pub(crate) fn formula_report(
    names: &[String],
    formula: &crate::calc::Formula,
    written: bool,
) -> Field {
    use crate::calc::Status as Calc;
    let location = format!(
        "{}!{}{}",
        names.get(formula.sheet()).map_or("?", String::as_str),
        column_letters(formula.column()),
        formula.row() + 1
    );
    let (status, reason) = match formula.status() {
        Calc::Match => (Status::Unchanged, None),
        Calc::Mismatch | Calc::Uncached if written => (Status::Refreshed, None),
        Calc::Mismatch | Calc::Uncached => (Status::Stale, None),
        Calc::Volatile => (
            Status::NeedsClock,
            Some("the formula calls a volatile function"),
        ),
        _ => (
            Status::Unsupported,
            Some("the formula engine cannot evaluate the formula"),
        ),
    };
    let field = Field::new("FORMULA", Some(location), status);
    match reason {
        Some(reason) => field.because(reason),
        None => field,
    }
}

/// The report of a workbook recalculation that writes nothing.
#[cfg(all(
    feature = "eval",
    any(feature = "xlsx", feature = "xlsb", feature = "xls")
))]
pub(crate) fn workbook_summary(
    file: &crate::File,
    report: &crate::calc::Report,
    written: bool,
) -> Result<Summary> {
    let names = crate::calc::sheet_names(file)?;
    Ok(Summary {
        fields: report
            .formulas()
            .iter()
            .map(|formula| formula_report(&names, formula, written))
            .collect(),
        bookmarks: 0,
        missing_page_numbers: 0,
    })
}

/// Why a field was not computed.
#[cfg(any(feature = "docx", feature = "doc", feature = "odt", feature = "odp"))]
pub(crate) type Skip = (Status, Option<&'static str>);

/// The new displayed text of an ODF text field (record 0295): `name` is the
/// element's local name (`date`, `title`, ...), `fixed` its `text:fixed`,
/// `styled` whether it names a data style. Page fields are the caller's
/// (they need layout in a text document, the slide position in a
/// presentation).
#[cfg(any(feature = "odt", feature = "odp"))]
pub(crate) fn odf_field(
    name: &str,
    fixed: bool,
    styled: bool,
    metadata: &litchi_core::Metadata,
    clock: Option<Clock>,
) -> std::result::Result<String, Skip> {
    if fixed {
        return Err((Status::Locked, Some("the field is fixed (text:fixed)")));
    }
    let dated = |stamp: Option<litchi_core::metadata::Timestamp>, time: bool| {
        if styled {
            return Err((
                Status::Unsupported,
                Some("ODF data styles are not applied to field values"),
            ));
        }
        let clock = stamp.and_then(Clock::from_timestamp).ok_or((
            Status::Unresolved,
            Some("the document does not store this time"),
        ))?;
        Ok(clock.format(if time { "HH:mm:ss" } else { "yyyy-MM-dd" }))
    };
    let property = |value: &Option<String>| Ok(value.clone().unwrap_or_default());
    match name {
        "date" | "time" => {
            if styled {
                return Err((
                    Status::Unsupported,
                    Some("ODF data styles are not applied to field values"),
                ));
            }
            let clock = clock.ok_or((Status::NeedsClock, None))?;
            Ok(clock.format(if name == "date" {
                "yyyy-MM-dd"
            } else {
                "HH:mm:ss"
            }))
        },
        "creation-date" => dated(metadata.created, false),
        "creation-time" => dated(metadata.created, true),
        "modification-date" => dated(metadata.modified, false),
        "modification-time" => dated(metadata.modified, true),
        "print-date" => dated(metadata.last_printed, false),
        "print-time" => dated(metadata.last_printed, true),
        "title" => property(&metadata.title),
        "subject" => property(&metadata.subject),
        "description" => property(&metadata.description),
        // LibreOffice shows the `meta:keyword` elements joined with `, `.
        "keywords" => Ok(metadata.keywords.join(", ")),
        "initial-creator" => property(&metadata.author),
        "creator" => property(&metadata.last_modified_by),
        _ => Err((Status::Unsupported, None)),
    }
}

#[cfg(any(feature = "docx", feature = "doc"))]
mod word {
    use std::collections::HashMap;

    use litchi_core::Metadata;

    use super::{Clock, Instruction, Status, apply_case, format_number};

    use super::Skip;

    /// The new result of a Word field that is not a table of contents (DOCX,
    /// and DOC since record 0295): `bookmark` reads a bookmark's text for
    /// `REF`, `counts` the main story's words and non-space characters.
    pub(crate) fn word_field(
        kind: &str,
        instruction: &Instruction,
        metadata: &Metadata,
        clock: Option<Clock>,
        seq: &mut HashMap<String, (u64, bool)>,
        bookmark: impl FnOnce(&str) -> Option<String>,
        counts: impl FnOnce() -> (usize, usize),
    ) -> Result<String, Skip> {
        if instruction.has('#') {
            return Err((
                Status::Unsupported,
                Some("numeric pictures (\\#) are not applied"),
            ));
        }
        let text = match kind {
            "SEQ" => return sequence(instruction, seq),
            "DATE" | "TIME" => {
                let clock = clock.ok_or((Status::NeedsClock, None))?;
                let default = if kind == "DATE" {
                    "M/d/yyyy"
                } else {
                    "h:mm am/pm"
                };
                clock.format(instruction.switch('@').unwrap_or(default))
            },
            "CREATEDATE" | "SAVEDATE" | "PRINTDATE" => {
                let stamp = match kind {
                    "CREATEDATE" => metadata.created,
                    "SAVEDATE" => metadata.modified,
                    _ => metadata.last_printed,
                };
                let clock = stamp.and_then(Clock::from_timestamp).ok_or((
                    Status::Unresolved,
                    Some("the document does not store this time"),
                ))?;
                clock.format(instruction.switch('@').unwrap_or("M/d/yyyy h:mm:ss am/pm"))
            },
            "REF" => {
                let name = instruction
                    .arguments
                    .first()
                    .ok_or((Status::Unsupported, Some("REF names no bookmark")))?;
                if instruction
                    .switches
                    .iter()
                    .any(|(switch, _)| !matches!(switch, 'h' | '*'))
                {
                    return Err((
                        Status::Unsupported,
                        Some("REF paragraph-number and position switches are not computed"),
                    ));
                }
                bookmark(name).ok_or((Status::Unresolved, Some("the bookmark does not exist")))?
            },
            "TITLE" | "SUBJECT" | "AUTHOR" | "KEYWORDS" | "COMMENTS" | "LASTSAVEDBY" => {
                if !instruction.arguments.is_empty() {
                    return Err((
                        Status::Unsupported,
                        Some("setting a property from a field is not done"),
                    ));
                }
                property(metadata, kind).unwrap_or_default()
            },
            "DOCPROPERTY" => {
                let name = instruction
                    .arguments
                    .first()
                    .ok_or((Status::Unsupported, Some("DOCPROPERTY names no property")))?;
                let key = name.replace(' ', "").to_ascii_uppercase();
                property(metadata, &key).ok_or((
                    Status::Unresolved,
                    Some("the property is not a known document property"),
                ))?
            },
            "NUMWORDS" | "NUMCHARS" => {
                let (words, characters) = counts();
                if kind == "NUMWORDS" {
                    words
                } else {
                    characters
                }
                .to_string()
            },
            _ => return Err((Status::Unsupported, None)),
        };
        let mut text = text;
        for format in instruction.formats() {
            text = apply_case(&text, format).ok_or((
                Status::Unsupported,
                Some("the \\* format is not applied to this field"),
            ))?;
        }
        Ok(text)
    }

    fn property(metadata: &Metadata, key: &str) -> Option<String> {
        let value = match key {
            "TITLE" => &metadata.title,
            "SUBJECT" => &metadata.subject,
            "AUTHOR" => &metadata.author,
            // The keyword list joined with `, ` (record 0354): the
            // source's own separators are not kept in the list.
            "KEYWORDS" => return Some(metadata.keywords.join(", ")),
            "COMMENTS" => &metadata.description,
            "LASTSAVEDBY" => &metadata.last_modified_by,
            "CATEGORY" => &metadata.category,
            "COMPANY" => &metadata.company,
            "MANAGER" => &metadata.manager,
            "TEMPLATE" => &metadata.template,
            _ => return None,
        };
        Some(value.clone().unwrap_or_default())
    }

    /// `SEQ` numbering in document order, after OfficeCLI's evaluator.
    pub(crate) fn sequence(
        instruction: &Instruction,
        counters: &mut HashMap<String, (u64, bool)>,
    ) -> Result<String, Skip> {
        let identifier = instruction
            .arguments
            .first()
            .ok_or((Status::Unsupported, Some("SEQ names no sequence")))?
            .clone();
        let counter = counters.entry(identifier).or_insert((0, false));
        if counter.1 {
            return Err((
                Status::Unsupported,
                Some("an earlier field of this sequence was not computed"),
            ));
        }
        if instruction.has('s') {
            counter.1 = true;
            return Err((
                Status::Unsupported,
                Some("heading-relative SEQ resets (\\s) need heading numbering"),
            ));
        }
        let value = if let Some(reset) = instruction.switch('r') {
            let Ok(reset) = reset.trim().parse::<u64>() else {
                counter.1 = true;
                return Err((
                    Status::Unsupported,
                    Some("the SEQ \\r reset is not a number"),
                ));
            };
            reset
        } else if instruction.has('c') {
            counter.0
        } else {
            counter.0 + 1
        };
        counter.0 = value;
        if instruction.has('h') {
            return Ok(String::new());
        }
        let mut formats = instruction.formats();
        let text = match formats.next() {
            None => value.to_string(),
            Some(format) => {
                if formats.next().is_some() {
                    counter.1 = true;
                    return Err((Status::Unsupported, Some("several SEQ formats")));
                }
                if let Some(text) = format_number(value, format) {
                    text
                } else {
                    counter.1 = true;
                    return Err((
                        Status::Unsupported,
                        Some("the SEQ number format is not applied"),
                    ));
                }
            },
        };
        Ok(text)
    }
}
#[cfg(any(feature = "docx", feature = "doc"))]
#[cfg(all(test, feature = "docx"))]
pub(crate) use word::sequence;
#[cfg(any(feature = "docx", feature = "doc"))]
pub(crate) use word::word_field;

#[cfg(test)]
mod tests;
