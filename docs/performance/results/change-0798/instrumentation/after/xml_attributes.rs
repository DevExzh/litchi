//! Attribute iteration for readers that stop at a start tag's first attribute
//! error, with a worst case that stays bounded on hostile tags.
//!
//! quick-xml 0.41 checks a start tag's attribute names for duplicates by
//! default. It compares each name with the names before it while it has seen
//! at most 32; from the 33rd name on it pre-filters with a hash set and, on
//! every pre-filter hit, scans all the names before it. The pre-filter's
//! hasher is not keyed (`DefaultHasher::new()`), so a name hashes to the same
//! value in every process built with the same standard library: whoever writes
//! the tag chooses how often the pre-filter hits, and each hit costs a scan of
//! the tag's earlier names. The check's worst case therefore grows with the
//! square of the tag, and nothing but the tag's size bounds it where no
//! per-element limit applies.
//!
//! [`BytesStartExt::checked_attributes`] yields what quick-xml's checked
//! iterator (`BytesStart::attributes`) yields up to and including its first
//! error, and nothing after it. quick-xml checks the first 32 names with its
//! linear scan, exactly as before; from the 33rd on its check is off and this
//! iterator checks each name itself, in an ordered map: `O(log n)` name
//! comparisons per attribute and no hashing, so a tag of `n` attributes costs
//! `O(n log n)` comparisons whatever its names are. The first error is the one
//! quick-xml reports, at the same attribute, with the same positions, so a
//! reader that stops at an attribute's error (`?` on each item) behaves as it
//! did with quick-xml's iterator.
//!
//! A reader that skips attribute errors and reads on uses
//! `litchi_ooxml_common::xml::attributes::first_wins` instead (record 0764),
//! and a reader that needs no duplicate check uses
//! [`BytesStartExt::unchecked_attributes`]. The workspace's `clippy.toml`
//! disallows quick-xml's checked iteration in the crates that read untrusted
//! OOXML and OLE2 XML, so a new call site has to choose one of the three.

use core::cell::{Cell, RefCell};
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::collections::btree_map::Entry;
use std::thread::{self, ThreadId};

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::{AttrError, Attribute, Attributes};

/// Names quick-xml checks with a linear scan of the names before them. From
/// the next name on it would hash names with its unkeyed hasher, so this
/// module turns its check off there and checks the names itself.
const QUICK_XML_LINEAR_NAMES: usize = 32;

/// Bounded attribute iteration for quick-xml start tags.
pub trait BytesStartExt {
    /// Iterate the tag's attributes as `BytesStart::attributes` does, up to
    /// and including the first error; nothing is yielded after it.
    ///
    /// Items and errors are the ones quick-xml's checked iterator yields,
    /// including `AttrError::Duplicated(position, first_position)` at the
    /// first repeated name, and a repeated name whose value is malformed is
    /// reported as a duplicate, as quick-xml checks a name before reading its
    /// value. The cost is quick-xml's linear check for the first 32 names and
    /// `O(log n)` name comparisons for each later one; no name is hashed.
    fn checked_attributes(&self) -> CheckedAttributes<'_>;

    /// Iterate the tag's attributes without any duplicate check: quick-xml's
    /// iterator with `with_checks(false)`. Every occurrence of a repeated name
    /// is yielded as `Ok`, and lexical errors are yielded as quick-xml reports
    /// them, with its recovery after each.
    fn unchecked_attributes(&self) -> Attributes<'_>;
}

impl BytesStartExt for BytesStart<'_> {
    #[inline]
    fn checked_attributes(&self) -> CheckedAttributes<'_> {
        CheckedAttributes::new(self)
    }

    // The two calls quick-xml offers for an unchecked iterator.
    #[allow(clippy::disallowed_methods)]
    #[inline]
    fn unchecked_attributes(&self) -> Attributes<'_> {
        let mut attributes = self.attributes();
        attributes.with_checks(false);
        attributes
    }
}

/// The iterator [`BytesStartExt::checked_attributes`] returns.
#[derive(Debug)]
pub struct CheckedAttributes<'a> {
    tag: &'a BytesStart<'a>,
    attributes: Attributes<'a>,
    phase: Phase<'a>,
    census: Option<CensusToken0798>,
}

#[derive(Clone, Debug)]
enum Phase<'a> {
    /// quick-xml checks the names; this many have been yielded.
    QuickXml(usize),
    /// quick-xml's check is off; this iterator checks the names.
    Own(Box<OwnCheck<'a>>),
    /// An error or the end has been yielded.
    Done,
}

/// This iterator's check of the names, from the 33rd on.
#[derive(Clone, Debug)]
struct OwnCheck<'a> {
    /// Every name yielded so far, with its position.
    names: BTreeMap<Name<'a>, usize>,
    /// Where quick-xml starts to look for the next attribute: just after the
    /// last one yielded.
    next_start: usize,
}

impl<'a> CheckedAttributes<'a> {
    // quick-xml's checked iterator, used for the first 32 names only.
    #[allow(clippy::disallowed_methods)]
    #[inline]
    fn new(tag: &'a BytesStart<'a>) -> Self {
        Self {
            tag,
            attributes: tag.attributes(),
            phase: Phase::QuickXml(0),
            census: census_start_0798(tag),
        }
    }

    /// Everything after quick-xml's 32 names: the switch to this iterator's
    /// own check, its check of each later name, and the end.
    #[cold]
    #[inline(never)]
    fn next_after_quick_xml(&mut self) -> Option<Result<Attribute<'a>, AttrError>> {
        match self.phase {
            Phase::Done => return None,
            // quick-xml would hash the next name: turn its check off first.
            Phase::QuickXml(_) => self.take_over(),
            Phase::Own(_) => {},
        }
        let Some(item) = self.attributes.next() else {
            self.phase = Phase::Done;
            return None;
        };
        let checked = if let Phase::Own(check) = &mut self.phase {
            check.check(self.tag, item)
        } else {
            let mut check = OwnCheck::after_quick_xml(self.tag);
            let checked = check.check(self.tag, item);
            self.phase = Phase::Own(Box::new(check));
            checked
        };
        if checked.is_err() {
            self.phase = Phase::Done;
        }
        Some(checked)
    }

    // Turning quick-xml's check off before its hashed check can run.
    #[allow(clippy::disallowed_methods)]
    fn take_over(&mut self) {
        self.attributes.with_checks(false);
    }
}

impl<'a> OwnCheck<'a> {
    /// The 32 names quick-xml has checked, with their positions, and where
    /// the attribute after them starts. Reading them again without the check
    /// yields the same attributes: the check never changes what is read.
    fn after_quick_xml(tag: &'a BytesStart<'a>) -> Self {
        let mut check = Self {
            names: BTreeMap::new(),
            next_start: 0,
        };
        for attribute in tag
            .unchecked_attributes()
            .take(QUICK_XML_LINEAR_NAMES)
            .flatten()
        {
            let name = attribute.key.into_inner();
            check.names.insert(Name(name), offset_in(tag, name));
            check.next_start = end_of(tag, &attribute);
        }
        debug_assert_eq!(check.names.len(), QUICK_XML_LINEAR_NAMES);
        check
    }

    fn check(
        &mut self,
        tag: &'a BytesStart<'a>,
        item: Result<Attribute<'a>, AttrError>,
    ) -> Result<Attribute<'a>, AttrError> {
        let attribute = match item {
            Ok(attribute) => attribute,
            Err(error) => return Err(self.duplicate_before_value(tag, error)),
        };
        let name = attribute.key.into_inner();
        let position = offset_in(tag, name);
        match self.names.entry(Name(name)) {
            Entry::Occupied(first) => Err(AttrError::Duplicated(position, *first.get())),
            Entry::Vacant(entry) => {
                entry.insert(position);
                self.next_start = end_of(tag, &attribute);
                Ok(attribute)
            },
        }
    }

    /// quick-xml checks a name once it has read the name and the `=` after
    /// it, before the value: a repeated name followed by a malformed value is
    /// reported as a duplicate, not as the value's error.
    fn duplicate_before_value(&self, tag: &[u8], error: AttrError) -> AttrError {
        if !matches!(
            error,
            AttrError::UnquotedValue(_)
                | AttrError::ExpectedValue(_)
                | AttrError::ExpectedQuote(..)
        ) {
            return error;
        }
        let Some((position, name)) = name_at(tag, self.next_start) else {
            return error;
        };
        match self.names.get(&Name(name)) {
            Some(first) => AttrError::Duplicated(position, *first),
            None => error,
        }
    }
}

impl<'a> Iterator for CheckedAttributes<'a> {
    type Item = Result<Attribute<'a>, AttrError>;

    /// The first 32 names are quick-xml's checked iterator plus a counter.
    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        if let Phase::QuickXml(yielded) = &mut self.phase
            && *yielded < QUICK_XML_LINEAR_NAMES
        {
            let item = self.attributes.next();
            if matches!(item, Some(Ok(_))) {
                *yielded += 1;
            } else {
                self.phase = Phase::Done;
            }
            census_next_0798(self.census, &item);
            return item;
        }
        let item = self.next_after_quick_xml();
        census_next_0798(self.census, &item);
        item
    }
}

impl core::iter::FusedIterator for CheckedAttributes<'_> {}

impl<'a> Clone for CheckedAttributes<'a> {
    fn clone(&self) -> Self {
        Self {
            tag: self.tag,
            attributes: self.attributes.clone(),
            phase: self.phase.clone(),
            census: census_clone_0798(self.census),
        }
    }
}

impl Drop for CheckedAttributes<'_> {
    fn drop(&mut self) {
        census_drop_0798(self.tag, self.census.take());
    }
}

/// The byte offset of `part` within `base`, of which it is a subslice.
fn offset_in(base: &[u8], part: &[u8]) -> usize {
    (part.as_ptr() as usize).wrapping_sub(base.as_ptr() as usize)
}

/// Where quick-xml looks for the attribute after `attribute`: just past the
/// closing quote of its value. quick-xml yields values borrowed from the tag.
fn end_of(base: &[u8], attribute: &Attribute<'_>) -> usize {
    debug_assert!(matches!(attribute.value, Cow::Borrowed(_)));
    let value = attribute.value.as_ref();
    offset_in(base, value)
        .saturating_add(value.len())
        .saturating_add(1)
}

/// The name of the attribute that starts at or after `from`, read the way
/// quick-xml reads it: after any whitespace, up to `=` or whitespace; with
/// its position.
fn name_at(tag: &[u8], from: usize) -> Option<(usize, &[u8])> {
    let rest = tag.get(from..)?;
    let start = from + rest.iter().position(|byte| !is_whitespace(*byte))?;
    let name = &tag[start..];
    let length = name
        .iter()
        .position(|byte| *byte == b'=' || is_whitespace(*byte))
        .unwrap_or(name.len());
    Some((start, &name[..length]))
}

/// quick-xml's whitespace: space, tab, carriage return and line feed.
fn is_whitespace(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// The lifecycle state observed by the diagnostic iterator census.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AttributeCensusTermination0798 {
    /// No terminal `next` result or drop has been observed yet.
    Active,
    /// The checked iterator returned an attribute error.
    Error,
    /// The checked iterator returned `None`.
    Exhausted,
    /// The iterator was dropped before returning a terminal result.
    Dropped,
    /// The census was finished while the iterator was still live.
    LiveAtFinish,
}

/// One checked-iterator instance recorded by the diagnostic census.
///
/// The source and element name are copied only while the diagnostic census is
/// enabled. `successful_yields` and the other event counters belong to this
/// instance; a clone starts them at zero and records its inherited successful
/// prefix in `starting_successful_yields`.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct AttributeCensusRow0798 {
    /// Monotonic identifier within one census session.
    pub instance_id: u64,
    /// Identifier shared by an iterator and its clones.
    pub lineage_id: u64,
    /// Whether this instance was made by `Clone`.
    pub is_clone: bool,
    /// Successful attributes observed by the source iterator before this clone.
    pub starting_successful_yields: u64,
    /// Exact raw `BytesStart` content, including the element name.
    pub source: Vec<u8>,
    /// Exact raw element name bytes.
    pub element_name: Vec<u8>,
    /// Number of calls to this instance's `next` method.
    pub next_calls: u64,
    /// Number of `Some(Ok(_))` results observed by this instance.
    pub successful_yields: u64,
    /// Number of `Some(Err(_))` results observed by this instance.
    pub error_yields: u64,
    /// Number of `None` results observed by this instance.
    pub end_yields: u64,
    /// Number of `Ok` items found by the separate unchecked lexical scan.
    pub lexical_attribute_count: u64,
    /// Number of all items, including errors, found by that scan.
    pub lexical_item_count: u64,
    /// Number of errors found by that scan.
    pub lexical_error_count: u64,
    /// Whether the separate unchecked scan has completed.
    pub lexical_scan_completed: bool,
    /// Lifecycle state at report finalization.
    pub termination: AttributeCensusTermination0798,
    /// Whether this instance's `Drop` hook ran while the session was active.
    pub dropped: bool,
    /// Whether `Drop` ran before an error or `None` result was observed.
    pub early_drop: bool,
    /// Whether the checked iterator observed fewer successful attributes than
    /// the complete lexical scan, after accounting for a clone's prefix.
    pub partial_consumption: bool,
    /// Whether this instance was still live when `finish_attribute_census_0798`
    /// took the registry.
    pub live_at_finish: bool,
    /// Whether any per-row counter saturated at `u64::MAX`.
    pub counter_saturated: bool,
}

/// Result returned by [`finish_attribute_census_0798`].
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct AttributeCensusReport0798 {
    /// One row for every counted iterator start or clone.
    pub rows: Vec<AttributeCensusRow0798>,
    /// Number of `CheckedAttributes` constructors observed.
    pub iterator_starts: u64,
    /// Number of counted `Clone` calls.
    pub iterator_clones: u64,
    /// Number of counted `Drop` calls completed before finish.
    pub iterator_drops: u64,
    /// Number of counted instances still live when finish took the registry.
    pub live_instances_at_finish: u64,
    /// Whether any session-level counter saturated at `u64::MAX`.
    pub counter_saturated: bool,
}

#[derive(Clone, Copy, Debug)]
struct CensusToken0798 {
    generation: u64,
    thread: ThreadId,
    row: usize,
}

struct CensusRowState0798 {
    tag: BytesStart<'static>,
    report: AttributeCensusRow0798,
}

struct CensusSession0798 {
    generation: u64,
    thread: ThreadId,
    rows: Vec<CensusRowState0798>,
    next_instance_id: u64,
    next_lineage_id: u64,
    iterator_starts: u64,
    iterator_clones: u64,
    iterator_drops: u64,
    live_instances: u64,
    counter_saturated: bool,
}

impl Default for CensusSession0798 {
    fn default() -> Self {
        Self {
            generation: 0,
            thread: thread::current().id(),
            rows: Vec::new(),
            next_instance_id: 0,
            next_lineage_id: 0,
            iterator_starts: 0,
            iterator_clones: 0,
            iterator_drops: 0,
            live_instances: 0,
            counter_saturated: false,
        }
    }
}

#[derive(Default)]
struct LexicalCounts0798 {
    attributes: u64,
    items: u64,
    errors: u64,
    counter_saturated: bool,
}

thread_local! {
    static ATTRIBUTE_CENSUS_0798: RefCell<Option<CensusSession0798>> = const { RefCell::new(None) };
    static ATTRIBUTE_CENSUS_ENABLED_0798: Cell<bool> = const { Cell::new(false) };
    static ATTRIBUTE_CENSUS_GENERATION_0798: Cell<u64> = const { Cell::new(0) };
}

/// Start a fresh checked-iterator census on the calling thread.
///
/// Calling this function discards an unfinished session on the same thread.
/// Iterators made by an earlier session become stale and are ignored by later
/// iterator, clone, and drop hooks. The caller must keep the session on one
/// thread and call [`finish_attribute_census_0798`] after all measured owners
/// have been dropped.
pub fn begin_attribute_census_0798() {
    ATTRIBUTE_CENSUS_ENABLED_0798.with(|enabled| enabled.set(false));
    let (generation, generation_saturated) = ATTRIBUTE_CENSUS_GENERATION_0798.with(|serial| {
        match serial.get().checked_add(1) {
            Some(next) => {
                serial.set(next);
                (next, false)
            },
            None => (u64::MAX, true),
        }
    });
    ATTRIBUTE_CENSUS_0798.with(|registry| {
        *registry.borrow_mut() = Some(CensusSession0798 {
            generation,
            thread: thread::current().id(),
            next_instance_id: 1,
            next_lineage_id: 1,
            counter_saturated: generation_saturated,
            ..CensusSession0798::default()
        });
    });
    ATTRIBUTE_CENSUS_ENABLED_0798.with(|enabled| enabled.set(true));
}

/// Finish the calling thread's census and return its plain diagnostic rows.
///
/// The enabled bit is cleared before the registry is taken. Consequently,
/// iterators dropped after this function returns are stale and do not mutate a
/// later session. Live rows are retained with `live_at_finish = true`; this is
/// a qualification failure for the probe rather than an inferred drop.
pub fn finish_attribute_census_0798() -> AttributeCensusReport0798 {
    ATTRIBUTE_CENSUS_ENABLED_0798.with(|enabled| enabled.set(false));
    ATTRIBUTE_CENSUS_0798.with(|registry| {
        let Some(mut session) = registry.borrow_mut().take() else {
            return AttributeCensusReport0798::default();
        };
        session.finish_live_rows();
        AttributeCensusReport0798 {
            rows: session.rows.into_iter().map(|row| row.report).collect(),
            iterator_starts: session.iterator_starts,
            iterator_clones: session.iterator_clones,
            iterator_drops: session.iterator_drops,
            live_instances_at_finish: session.live_instances,
            counter_saturated: session.counter_saturated,
        }
    })
}

fn saturating_increment(value: &mut u64, saturated: &mut bool) {
    if let Some(next) = value.checked_add(1) {
        *value = next;
    } else {
        *value = u64::MAX;
        *saturated = true;
    }
}

fn saturating_add(value: &mut u64, amount: u64, saturated: &mut bool) {
    if let Some(next) = value.checked_add(amount) {
        *value = next;
    } else {
        *value = u64::MAX;
        *saturated = true;
    }
}

fn next_identifier(value: &mut u64, saturated: &mut bool) -> u64 {
    let current = *value;
    if let Some(next) = value.checked_add(1) {
        *value = next;
    } else {
        *value = u64::MAX;
        *saturated = true;
    }
    current
}

fn census_start_0798(tag: &BytesStart<'_>) -> Option<CensusToken0798> {
    let enabled = ATTRIBUTE_CENSUS_ENABLED_0798.with(Cell::get);
    if !enabled {
        return None;
    }
    ATTRIBUTE_CENSUS_0798.with(|registry| {
        registry
            .borrow_mut()
            .as_mut()
            .map(|session| session.start(tag))
    })
}

fn census_clone_0798(token: Option<CensusToken0798>) -> Option<CensusToken0798> {
    let token = token?;
    if !ATTRIBUTE_CENSUS_ENABLED_0798.with(Cell::get) {
        return None;
    }
    ATTRIBUTE_CENSUS_0798.with(|registry| {
        registry
            .borrow_mut()
            .as_mut()
            .and_then(|session| session.clone_instance(token))
    })
}

fn census_next_0798(
    token: Option<CensusToken0798>,
    item: &Option<Result<Attribute<'_>, AttrError>>,
) {
    let Some(token) = token else {
        return;
    };
    if !ATTRIBUTE_CENSUS_ENABLED_0798.with(Cell::get) {
        return;
    }
    ATTRIBUTE_CENSUS_0798.with(|registry| {
        if let Some(session) = registry.borrow_mut().as_mut() {
            session.record_next(token, item);
        }
    });
}

fn census_drop_0798(tag: &BytesStart<'_>, token: Option<CensusToken0798>) {
    let Some(token) = token else {
        return;
    };
    if !ATTRIBUTE_CENSUS_ENABLED_0798.with(Cell::get) {
        return;
    }
    let valid = ATTRIBUTE_CENSUS_0798.with(|registry| {
        registry
            .borrow()
            .as_ref()
            .is_some_and(|session| session.valid(token))
    });
    if !valid {
        return;
    }
    let lexical = lexical_counts_0798(tag);
    ATTRIBUTE_CENSUS_0798.with(|registry| {
        if let Some(session) = registry.borrow_mut().as_mut() {
            session.record_drop(token, lexical);
        }
    });
}

impl CensusSession0798 {
    fn valid(&self, token: CensusToken0798) -> bool {
        token.generation == self.generation
            && token.thread == self.thread
            && token.row < self.rows.len()
    }

    fn start(&mut self, tag: &BytesStart<'_>) -> CensusToken0798 {
        let instance_id = next_identifier(&mut self.next_instance_id, &mut self.counter_saturated);
        let lineage_id = next_identifier(&mut self.next_lineage_id, &mut self.counter_saturated);
        saturating_increment(&mut self.iterator_starts, &mut self.counter_saturated);
        saturating_increment(&mut self.live_instances, &mut self.counter_saturated);
        let tag = tag.to_owned();
        let row = self.rows.len();
        self.rows.push(CensusRowState0798 {
            report: AttributeCensusRow0798 {
                instance_id,
                lineage_id,
                is_clone: false,
                starting_successful_yields: 0,
                source: tag.as_ref().to_vec(),
                element_name: tag.name().as_ref().to_vec(),
                next_calls: 0,
                successful_yields: 0,
                error_yields: 0,
                end_yields: 0,
                lexical_attribute_count: 0,
                lexical_item_count: 0,
                lexical_error_count: 0,
                lexical_scan_completed: false,
                termination: AttributeCensusTermination0798::Active,
                dropped: false,
                early_drop: false,
                partial_consumption: false,
                live_at_finish: false,
                counter_saturated: false,
            },
            tag,
        });
        CensusToken0798 {
            generation: self.generation,
            thread: self.thread,
            row,
        }
    }

    fn clone_instance(&mut self, token: CensusToken0798) -> Option<CensusToken0798> {
        if !self.valid(token) {
            return None;
        }
        let (tag, lineage_id, source_saturated, mut starting_successful_yields, source_successes) = {
            let source = &self.rows[token.row];
            (
                source.tag.clone(),
                source.report.lineage_id,
                source.report.counter_saturated,
                source.report.starting_successful_yields,
                source.report.successful_yields,
            )
        };
        let mut prefix_saturated = false;
        saturating_add(
            &mut starting_successful_yields,
            source_successes,
            &mut prefix_saturated,
        );
        self.counter_saturated |= prefix_saturated || source_saturated;
        let instance_id = next_identifier(&mut self.next_instance_id, &mut self.counter_saturated);
        let row = self.rows.len();
        self.rows.push(CensusRowState0798 {
            report: AttributeCensusRow0798 {
                instance_id,
                lineage_id,
                is_clone: true,
                starting_successful_yields,
                source: tag.as_ref().to_vec(),
                element_name: tag.name().as_ref().to_vec(),
                next_calls: 0,
                successful_yields: 0,
                error_yields: 0,
                end_yields: 0,
                lexical_attribute_count: 0,
                lexical_item_count: 0,
                lexical_error_count: 0,
                lexical_scan_completed: false,
                termination: AttributeCensusTermination0798::Active,
                dropped: false,
                early_drop: false,
                partial_consumption: false,
                live_at_finish: false,
                counter_saturated: prefix_saturated || source_saturated,
            },
            tag,
        });
        saturating_increment(&mut self.iterator_clones, &mut self.counter_saturated);
        saturating_increment(&mut self.live_instances, &mut self.counter_saturated);
        Some(CensusToken0798 {
            generation: self.generation,
            thread: self.thread,
            row,
        })
    }

    fn record_next(
        &mut self,
        token: CensusToken0798,
        item: &Option<Result<Attribute<'_>, AttrError>>,
    ) {
        if !self.valid(token) {
            return;
        }
        let row_saturated = {
            let row = &mut self.rows[token.row].report;
            saturating_increment(&mut row.next_calls, &mut row.counter_saturated);
            match item {
                Some(Ok(_)) => {
                    saturating_increment(&mut row.successful_yields, &mut row.counter_saturated);
                },
                Some(Err(_)) => {
                    saturating_increment(&mut row.error_yields, &mut row.counter_saturated);
                    if row.termination == AttributeCensusTermination0798::Active {
                        row.termination = AttributeCensusTermination0798::Error;
                    }
                },
                None => {
                    saturating_increment(&mut row.end_yields, &mut row.counter_saturated);
                    if row.termination == AttributeCensusTermination0798::Active {
                        row.termination = AttributeCensusTermination0798::Exhausted;
                    }
                },
            }
            row.counter_saturated
        };
        self.counter_saturated |= row_saturated;
    }

    fn record_drop(&mut self, token: CensusToken0798, lexical: LexicalCounts0798) {
        if !self.valid(token) {
            return;
        }
        let row_saturated = {
            let row = &mut self.rows[token.row].report;
            if row.dropped {
                return;
            }
            row.lexical_attribute_count = lexical.attributes;
            row.lexical_item_count = lexical.items;
            row.lexical_error_count = lexical.errors;
            row.lexical_scan_completed = true;
            row.dropped = true;
            row.early_drop = row.termination == AttributeCensusTermination0798::Active;
            if row.early_drop {
                row.termination = AttributeCensusTermination0798::Dropped;
            }
            let mut consumed = row.starting_successful_yields;
            saturating_add(&mut consumed, row.successful_yields, &mut row.counter_saturated);
            row.partial_consumption = consumed < row.lexical_attribute_count;
            row.counter_saturated
        };
        self.counter_saturated |= row_saturated || lexical.counter_saturated;
        saturating_increment(&mut self.iterator_drops, &mut self.counter_saturated);
        self.live_instances = self.live_instances.saturating_sub(1);
    }

    fn finish_live_rows(&mut self) {
        let mut live = 0u64;
        for index in 0..self.rows.len() {
            let row_saturated = {
                let row = &mut self.rows[index];
                if row.report.dropped {
                    continue;
                }
                let mut live_saturated = false;
                saturating_increment(&mut live, &mut live_saturated);
                let lexical = lexical_counts_0798(&row.tag);
                row.report.lexical_attribute_count = lexical.attributes;
                row.report.lexical_item_count = lexical.items;
                row.report.lexical_error_count = lexical.errors;
                row.report.lexical_scan_completed = true;
                row.report.live_at_finish = true;
                let mut consumed = row.report.starting_successful_yields;
                saturating_add(
                    &mut consumed,
                    row.report.successful_yields,
                    &mut row.report.counter_saturated,
                );
                row.report.partial_consumption = consumed < row.report.lexical_attribute_count;
                if row.report.termination == AttributeCensusTermination0798::Active {
                    row.report.termination = AttributeCensusTermination0798::LiveAtFinish;
                }
                lexical.counter_saturated || row.report.counter_saturated || live_saturated
            };
            self.counter_saturated |= row_saturated;
        }
        self.live_instances = live;
    }
}

fn lexical_counts_0798(tag: &BytesStart<'_>) -> LexicalCounts0798 {
    let mut counts = LexicalCounts0798::default();
    for item in tag.unchecked_attributes() {
        saturating_increment(&mut counts.items, &mut counts.counter_saturated);
        match item {
            Ok(_) => saturating_increment(&mut counts.attributes, &mut counts.counter_saturated),
            Err(_) => {
                saturating_increment(&mut counts.errors, &mut counts.counter_saturated);
                break;
            },
        }
    }
    counts
}

/// A name ordered by its bytes, with every comparison counted in tests.
#[derive(Clone, Copy, Debug)]
struct Name<'a>(&'a [u8]);

impl PartialEq for Name<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.cmp(other).is_eq()
    }
}

impl Eq for Name<'_> {}

impl PartialOrd for Name<'_> {
    fn partial_cmp(&self, other: &Self) -> Option<core::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for Name<'_> {
    fn cmp(&self, other: &Self) -> core::cmp::Ordering {
        #[cfg(test)]
        tests::count_comparison();
        self.0.cmp(other.0)
    }
}

#[cfg(test)]
mod tests;

#[cfg(test)]
#[path = "census_tests_0798.rs"]
mod census_tests_0798;
