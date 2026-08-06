//! A bounded weighted cache for immutable parser values.
//!
//! [`WeightedCache`] stores only [`Arc`] handles to values supplied by its
//! caller. The cache itself is synchronized, while parsing runs outside the
//! cache mutex. Calls for the same missing key share one parser invocation;
//! calls for different keys can parse concurrently.
//!
//! The cache has no runtime or third-party dependencies. It deliberately
//! accepts an explicit weight for every value because a generic cache cannot
//! infer the retained memory cost of an arbitrary `V` safely.

#![forbid(unsafe_code)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The public error contracts and their state-machine helpers are kept together for auditability."
)]

use std::collections::HashMap;
use std::error::Error;
use std::fmt;
use std::hash::Hash;
use std::panic::{self, AssertUnwindSafe};
use std::sync::{Arc, Condvar, Mutex, MutexGuard};

/// A parser error that can be shared safely with single-flight waiters.
pub type ParseError = Box<dyn Error + Send + Sync + 'static>;

/// A shared parser error returned to every caller waiting on one failed flight.
pub type SharedParseError = Arc<dyn Error + Send + Sync + 'static>;

/// The cache collection whose allocation could not be reserved.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum AllocationKind {
    /// Completed value entries.
    Entries,
    /// Active single-flight parser entries.
    Flights,
}

impl fmt::Display for AllocationKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::Entries => "cache entries",
            Self::Flights => "parser flights",
        };
        formatter.write_str(name)
    }
}

/// A recoverable allocation failure from a cache-owned collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AllocationError {
    kind: AllocationKind,
    requested: usize,
}

impl AllocationError {
    /// Return the cache collection that could not reserve storage.
    #[must_use]
    pub const fn kind(self) -> AllocationKind {
        self.kind
    }

    /// Return the number of additional slots requested when allocation failed.
    #[must_use]
    pub const fn requested(self) -> usize {
        self.requested
    }
}

impl fmt::Display for AllocationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "could not reserve {requested} slot(s) for {kind}",
            requested = self.requested,
            kind = self.kind
        )
    }
}

impl Error for AllocationError {}

/// A rejected cache capacity or value-weight request.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WeightError {
    /// A cache cannot be constructed with no positive weight budget.
    ZeroCapacity,
    /// A cache cannot permit no active parser flights.
    ZeroFlightCapacity,
    /// Zero-weight values are rejected so the weight bound also bounds the
    /// number of retained entries in every usable cache.
    ZeroWeight,
    /// A value is larger than the complete cache budget and can never fit.
    ExceedsCapacity {
        /// Weight supplied for the value.
        weight: usize,
        /// Maximum total cache weight.
        capacity: usize,
    },
    /// A caller supplied a weight different from the active flight's weight.
    InFlightMismatch {
        /// Weight selected by the caller that started parsing.
        expected: usize,
        /// Weight supplied by the waiting caller.
        requested: usize,
    },
}

impl fmt::Display for WeightError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ZeroCapacity => formatter.write_str("cache weight capacity must be non-zero"),
            Self::ZeroFlightCapacity => {
                formatter.write_str("active parser flight capacity must be non-zero")
            },
            Self::ZeroWeight => formatter.write_str("cached value weight must be non-zero"),
            Self::ExceedsCapacity { weight, capacity } => write!(
                formatter,
                "cached value weight {weight} exceeds cache capacity {capacity}"
            ),
            Self::InFlightMismatch {
                expected,
                requested,
            } => write!(
                formatter,
                "active parser flight uses weight {expected}, not requested weight {requested}"
            ),
        }
    }
}

impl Error for WeightError {}

/// Errors raised while inserting or otherwise updating cache state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum CacheError {
    /// The cache could not reserve its next internal slot.
    Allocation(AllocationError),
    /// A supplied capacity or value weight is invalid.
    Weight(WeightError),
    /// The bounded active-parser budget is currently exhausted.
    FlightsLimit {
        /// Number of active parser generations when the request was rejected.
        active: usize,
        /// Maximum number of active parser generations.
        limit: usize,
    },
}

impl fmt::Display for CacheError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Allocation(error) => error.fmt(formatter),
            Self::Weight(error) => error.fmt(formatter),
            Self::FlightsLimit { active, limit } => write!(
                formatter,
                "active parser flight limit {limit} reached with {active} flight(s)"
            ),
        }
    }
}

impl Error for CacheError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        match self {
            Self::Allocation(error) => Some(error),
            Self::Weight(error) => Some(error),
            Self::FlightsLimit { .. } => None,
        }
    }
}

/// Errors raised by a single-flight lookup or parser invocation.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum GetOrInsertError {
    /// The cache could not retain the parsed value.
    Cache(CacheError),
    /// The parser returned an error. The shared error is identical for all
    /// callers waiting on that parser generation.
    Parse(SharedParseError),
    /// The parser panicked. The initiating caller resumes the panic; waiters
    /// receive this error so they cannot remain blocked indefinitely.
    ParserPanicked,
}

impl fmt::Display for GetOrInsertError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Cache(error) => error.fmt(formatter),
            Self::Parse(error) => write!(formatter, "parser failed: {error}"),
            Self::ParserPanicked => formatter.write_str("parser panicked"),
        }
    }
}

impl Error for GetOrInsertError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        match self {
            Self::Cache(error) => Some(error),
            Self::Parse(error) => Some(error.as_ref()),
            Self::ParserPanicked => None,
        }
    }
}

impl From<CacheError> for GetOrInsertError {
    fn from(error: CacheError) -> Self {
        Self::Cache(error)
    }
}

#[derive(Debug)]
struct Entry<V> {
    value: Arc<V>,
    weight: usize,
    last_used: usize,
    renumbered: bool,
}

impl<V> Clone for Entry<V> {
    fn clone(&self) -> Self {
        Self {
            value: Arc::clone(&self.value),
            weight: self.weight,
            last_used: self.last_used,
            renumbered: self.renumbered,
        }
    }
}

enum FlightOutcome<V> {
    Success(Arc<V>),
    Cache(CacheError),
    Parse(SharedParseError),
    ParserPanicked,
}

impl<V> Clone for FlightOutcome<V> {
    fn clone(&self) -> Self {
        match self {
            Self::Success(value) => Self::Success(Arc::clone(value)),
            Self::Cache(error) => Self::Cache(*error),
            Self::Parse(error) => Self::Parse(Arc::clone(error)),
            Self::ParserPanicked => Self::ParserPanicked,
        }
    }
}

struct Flight<V> {
    weight: Option<usize>,
    state: Mutex<Option<FlightOutcome<V>>>,
    completed: Condvar,
}

impl<V> Flight<V> {
    fn new(weight: Option<usize>) -> Self {
        Self {
            weight,
            state: Mutex::new(None),
            completed: Condvar::new(),
        }
    }

    fn complete(&self, outcome: FlightOutcome<V>) {
        let mut state = lock(&self.state);
        *state = Some(outcome);
        self.completed.notify_all();
    }

    fn wait(&self) -> FlightOutcome<V> {
        let mut state = lock(&self.state);
        loop {
            if let Some(outcome) = state.as_ref() {
                return outcome.clone();
            }
            state = self
                .completed
                .wait(state)
                .unwrap_or_else(std::sync::PoisonError::into_inner);
        }
    }
}

struct State<K, V> {
    entries: HashMap<K, Entry<V>>,
    flights: HashMap<K, Arc<Flight<V>>>,
    total_weight: usize,
    next_recency: usize,
}

impl<K, V> Default for State<K, V> {
    fn default() -> Self {
        Self {
            entries: HashMap::new(),
            flights: HashMap::new(),
            total_weight: 0,
            next_recency: 0,
        }
    }
}

impl<K, V> State<K, V>
where
    K: Eq + Hash,
{
    fn touch(&mut self, key: &K) {
        let recency = self.next_recency();
        if let Some(entry) = self.entries.get_mut(key) {
            entry.last_used = recency;
        }
    }

    fn next_recency(&mut self) -> usize {
        if self.next_recency == usize::MAX {
            self.renumber_recency();
        }
        let recency = self.next_recency;
        self.next_recency = self.next_recency.saturating_add(1);
        recency
    }

    fn renumber_recency(&mut self) {
        for entry in self.entries.values_mut() {
            entry.renumbered = false;
        }

        let mut rank = 0;
        while rank < self.entries.len() {
            let oldest_tick = self
                .entries
                .values()
                .filter(|entry| !entry.renumbered)
                .map(|entry| entry.last_used)
                .min();
            let Some(oldest) = oldest_tick else {
                break;
            };

            for entry in self.entries.values_mut() {
                if !entry.renumbered && entry.last_used == oldest {
                    entry.last_used = rank;
                    entry.renumbered = true;
                    rank = rank.saturating_add(1);
                    break;
                }
            }
        }

        for entry in self.entries.values_mut() {
            entry.renumbered = false;
        }
        self.next_recency = rank;
    }

    fn evict_oldest(&mut self) {
        let Some(oldest) = self.entries.values().map(|entry| entry.last_used).min() else {
            return;
        };
        let mut removed_weight = 0;
        self.entries.retain(|_, entry| {
            if entry.last_used == oldest {
                removed_weight += entry.weight;
                false
            } else {
                true
            }
        });
        self.total_weight -= removed_weight;
    }

    fn evict_until_fit(&mut self, capacity: usize) {
        while self.total_weight > capacity {
            self.evict_oldest();
        }
    }
}

enum Lookup<K, V> {
    Cached(Arc<V>),
    Wait(Arc<Flight<V>>),
    Parse { key: K, flight: Arc<Flight<V>> },
}

/// A thread-safe cache of immutable, explicitly weighted values.
///
/// `K` must be cloneable because an active key is held by both the flight
/// registry and the parser owner. `V` is never exposed mutably: successful
/// lookups return a cloned [`Arc`] handle.
pub struct WeightedCache<K, V> {
    max_weight: usize,
    max_flights: usize,
    state: Mutex<State<K, V>>,
}

impl<K, V> fmt::Debug for WeightedCache<K, V>
where
    K: Eq + Hash + Clone,
{
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("WeightedCache")
            .field("max_weight", &self.max_weight)
            .field("max_flights", &self.max_flights)
            .field("len", &self.len())
            .field("total_weight", &self.total_weight())
            .finish_non_exhaustive()
    }
}

impl<K, V> WeightedCache<K, V>
where
    K: Eq + Hash + Clone,
{
    /// Default bound for distinct active parser generations.
    pub const DEFAULT_MAX_FLIGHTS: usize = 64;

    /// Construct a cache with a positive total-weight budget.
    ///
    /// # Errors
    ///
    /// Returns [`WeightError::ZeroCapacity`] when `max_weight` is zero.
    pub fn new(max_weight: usize) -> Result<Self, WeightError> {
        Self::new_with_limits(max_weight, Self::DEFAULT_MAX_FLIGHTS)
    }

    /// Construct a cache with explicit value-weight and active-flight budgets.
    ///
    /// Bounding active flights prevents a burst of distinct slow parser keys
    /// from growing an unbounded wait registry before any value is cacheable.
    ///
    /// # Errors
    ///
    /// Returns [`WeightError::ZeroCapacity`] when `max_weight` is zero or
    /// [`WeightError::ZeroFlightCapacity`] when `max_flights` is zero.
    pub fn new_with_limits(max_weight: usize, max_flights: usize) -> Result<Self, WeightError> {
        if max_weight == 0 {
            return Err(WeightError::ZeroCapacity);
        }
        if max_flights == 0 {
            return Err(WeightError::ZeroFlightCapacity);
        }
        Ok(Self {
            max_weight,
            max_flights,
            state: Mutex::new(State::default()),
        })
    }

    /// Return the configured total-weight budget.
    #[must_use]
    pub const fn max_weight(&self) -> usize {
        self.max_weight
    }

    /// Return the maximum number of distinct active parser generations.
    #[must_use]
    pub const fn max_flights(&self) -> usize {
        self.max_flights
    }

    /// Fork the completed cache state for an immutable copy-on-write generation.
    ///
    /// Retained values keep sharing their [`Arc`] allocations, while active
    /// parser flights are deliberately omitted because they belong to the
    /// source generation's exact input state.
    #[must_use]
    pub fn fork(&self) -> Self {
        let state = lock(&self.state);
        Self {
            max_weight: self.max_weight,
            max_flights: self.max_flights,
            state: Mutex::new(State {
                entries: state.entries.clone(),
                flights: HashMap::new(),
                total_weight: state.total_weight,
                next_recency: state.next_recency,
            }),
        }
    }

    /// Return the number of completed values currently retained.
    #[must_use]
    pub fn len(&self) -> usize {
        lock(&self.state).entries.len()
    }

    /// Return whether no completed values are currently retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// Return the total weight of all completed values currently retained.
    #[must_use]
    pub fn total_weight(&self) -> usize {
        lock(&self.state).total_weight
    }

    /// Return a retained value and mark it as most recently used.
    ///
    /// A plain lookup never waits for an active parser. Use
    /// [`Self::get_or_insert_with`] when a missing value should be parsed.
    #[must_use]
    pub fn get(&self, key: &K) -> Option<Arc<V>> {
        let mut state = lock(&self.state);
        let value = state.entries.get(key).map(|entry| Arc::clone(&entry.value));
        if value.is_some() {
            state.touch(key);
        }
        value
    }

    /// Return whether a completed value is retained without changing LRU order.
    #[must_use]
    pub fn contains_key(&self, key: &K) -> bool {
        lock(&self.state).entries.contains_key(key)
    }

    /// Insert or replace one immutable value.
    ///
    /// The value is retained only while its weight fits the configured
    /// budget. Least-recently-used completed entries are evicted until the
    /// new value fits. Replacing a key makes that key most recently used.
    /// Any active parser for the key is detached, so it cannot overwrite this
    /// explicit value when it completes.
    ///
    /// # Errors
    ///
    /// Returns a typed weight error for an invalid weight or a typed
    /// allocation error if the entry map cannot reserve its next slot. On
    /// either error, the completed cache state is unchanged.
    pub fn insert(&self, key: K, value: Arc<V>, weight: usize) -> Result<(), CacheError> {
        self.validate_weight(weight)?;
        let mut state = lock(&self.state);
        let replacing = state.entries.contains_key(&key);
        if !replacing && state.entries.try_reserve(1).is_err() {
            return Err(allocation_error(AllocationKind::Entries));
        }

        state.flights.remove(&key);
        if let Some(previous) = state.entries.remove(&key) {
            state.total_weight -= previous.weight;
        }
        let capacity = self.max_weight - weight;
        state.evict_until_fit(capacity);
        let recency = state.next_recency();
        state.entries.insert(
            key,
            Entry {
                value,
                weight,
                last_used: recency,
                renumbered: false,
            },
        );
        state.total_weight += weight;
        Ok(())
    }

    /// Invalidate one completed value and any active parser generation for its key.
    ///
    /// An invalidated parser may still finish for callers that already joined
    /// it, but its result is not published into this cache generation.
    ///
    /// Returns `true` when a completed value or active parser was removed.
    pub fn invalidate(&self, key: &K) -> bool {
        let mut state = lock(&self.state);
        let removed_value = state.entries.remove(key).map(|entry| {
            state.total_weight -= entry.weight;
        });
        let removed_flight = state.flights.remove(key);
        removed_value.is_some() || removed_flight.is_some()
    }

    /// Remove every completed value and detach every active parser generation.
    ///
    /// Existing parser callers are still completed normally, but their values
    /// cannot repopulate the cache after this call. A later lookup can start a
    /// fresh parser generation even if an older one is still running.
    pub fn clear(&self) {
        let mut state = lock(&self.state);
        state.entries.clear();
        state.flights.clear();
        state.total_weight = 0;
    }

    /// Return a cached value or parse it once for all concurrent callers of the key.
    ///
    /// The parser runs without holding the cache mutex. Its returned value is
    /// wrapped in an [`Arc`] exactly once and shared by the initiating caller
    /// and all waiters. A parser panic is resumed on the initiating thread;
    /// waiters receive [`GetOrInsertError::ParserPanicked`] instead.
    ///
    /// # Errors
    ///
    /// Returns typed cache allocation/weight errors. The parser itself is
    /// infallible; use [`Self::get_or_try_insert_with`] for a fallible parser.
    pub fn get_or_insert_with<F>(
        &self,
        key: K,
        weight: usize,
        parse: F,
    ) -> Result<Arc<V>, GetOrInsertError>
    where
        F: FnOnce() -> V,
    {
        self.get_or_try_insert_with(key, weight, || Ok::<V, ParseError>(parse()))
    }

    /// Return a cached value or run one fallible parser for all concurrent callers of the key.
    ///
    /// `ParseError` is boxed so the completed parser failure can be shared
    /// safely with callers that were waiting on the same flight. Parsing
    /// failures are not retained, so a later lookup may retry. If invalidation
    /// or clear detaches the flight while parsing, its callers still receive
    /// that generation's result but the cache does not retain it.
    ///
    /// # Errors
    ///
    /// Returns [`GetOrInsertError::Parse`] for a parser error,
    /// [`GetOrInsertError::Cache`] for a cache error, or
    /// [`GetOrInsertError::ParserPanicked`] to a waiter when the initiating
    /// parser panicked.
    pub fn get_or_try_insert_with<F>(
        &self,
        key: K,
        weight: usize,
        parse: F,
    ) -> Result<Arc<V>, GetOrInsertError>
    where
        F: FnOnce() -> Result<V, ParseError>,
    {
        let lookup = self.begin_lookup(key, Some(weight))?;
        match lookup {
            Lookup::Cached(value) => Ok(value),
            Lookup::Wait(flight) => flight.wait().into_result(),
            Lookup::Parse {
                key: parse_key,
                flight,
            } => self.run_parser(parse_key, weight, &flight, parse),
        }
    }

    /// Return a cached value or run one fallible parser whose retained weight
    /// is known only after parsing.
    ///
    /// # Errors
    ///
    /// Returns [`GetOrInsertError::Parse`] for a parser error,
    /// [`GetOrInsertError::Cache`] for a cache error, or
    /// [`GetOrInsertError::ParserPanicked`] to a waiter when the initiating
    /// parser panicked.
    pub fn get_or_try_insert_with_weight<F>(
        &self,
        key: K,
        parse: F,
    ) -> Result<Arc<V>, GetOrInsertError>
    where
        F: FnOnce() -> Result<(V, usize), ParseError>,
    {
        let lookup = self.begin_lookup(key, None)?;
        match lookup {
            Lookup::Cached(value) => Ok(value),
            Lookup::Wait(flight) => flight.wait().into_result(),
            Lookup::Parse {
                key: parse_key,
                flight,
            } => self.run_weighted_parser(parse_key, &flight, parse),
        }
    }

    fn begin_lookup(
        &self,
        key: K,
        weight: Option<usize>,
    ) -> Result<Lookup<K, V>, GetOrInsertError> {
        let mut state = lock(&self.state);
        if let Some(value) = state
            .entries
            .get(&key)
            .map(|entry| Arc::clone(&entry.value))
        {
            state.touch(&key);
            return Ok(Lookup::Cached(value));
        }

        if let Some(flight) = state.flights.get(&key) {
            if let (Some(expected), Some(requested)) = (flight.weight, weight)
                && expected != requested
            {
                return Err(CacheError::Weight(WeightError::InFlightMismatch {
                    expected,
                    requested,
                })
                .into());
            }
            return Ok(Lookup::Wait(Arc::clone(flight)));
        }

        if let Some(requested_weight) = weight {
            self.validate_weight(requested_weight)?;
        }
        if state.flights.len() >= self.max_flights {
            return Err(CacheError::FlightsLimit {
                active: state.flights.len(),
                limit: self.max_flights,
            }
            .into());
        }
        if state.flights.try_reserve(1).is_err() {
            return Err(allocation_error(AllocationKind::Flights).into());
        }
        let flight = Arc::new(Flight::new(weight));
        state.flights.insert(key.clone(), Arc::clone(&flight));
        Ok(Lookup::Parse { key, flight })
    }

    fn run_parser<F>(
        &self,
        key: K,
        weight: usize,
        flight: &Arc<Flight<V>>,
        parse: F,
    ) -> Result<Arc<V>, GetOrInsertError>
    where
        F: FnOnce() -> Result<V, ParseError>,
    {
        let parsed = panic::catch_unwind(AssertUnwindSafe(parse));
        match parsed {
            Ok(Ok(parsed_value)) => {
                let shared_value = Arc::new(parsed_value);
                self.publish_success(key, weight, flight, &shared_value)
            },
            Ok(Err(parse_error)) => {
                let shared_error: SharedParseError = Arc::from(parse_error);
                self.complete_detached_flight(
                    &key,
                    flight,
                    FlightOutcome::Parse(Arc::clone(&shared_error)),
                );
                Err(GetOrInsertError::Parse(shared_error))
            },
            Err(payload) => {
                self.complete_detached_flight(&key, flight, FlightOutcome::ParserPanicked);
                panic::resume_unwind(payload);
            },
        }
    }

    fn run_weighted_parser<F>(
        &self,
        key: K,
        flight: &Arc<Flight<V>>,
        parse: F,
    ) -> Result<Arc<V>, GetOrInsertError>
    where
        F: FnOnce() -> Result<(V, usize), ParseError>,
    {
        let parsed = panic::catch_unwind(AssertUnwindSafe(parse));
        match parsed {
            Ok(Ok((parsed_value, weight))) => {
                if let Err(error) = self.validate_weight(weight) {
                    self.complete_detached_flight(&key, flight, FlightOutcome::Cache(error));
                    return Err(GetOrInsertError::Cache(error));
                }
                let shared_value = Arc::new(parsed_value);
                self.publish_success(key, weight, flight, &shared_value)
            },
            Ok(Err(parse_error)) => {
                let shared_error: SharedParseError = Arc::from(parse_error);
                self.complete_detached_flight(
                    &key,
                    flight,
                    FlightOutcome::Parse(Arc::clone(&shared_error)),
                );
                Err(GetOrInsertError::Parse(shared_error))
            },
            Err(payload) => {
                self.complete_detached_flight(&key, flight, FlightOutcome::ParserPanicked);
                panic::resume_unwind(payload);
            },
        }
    }

    fn publish_success(
        &self,
        key: K,
        weight: usize,
        flight: &Arc<Flight<V>>,
        value: &Arc<V>,
    ) -> Result<Arc<V>, GetOrInsertError> {
        let outcome = {
            let mut state = lock(&self.state);
            let current = state
                .flights
                .get(&key)
                .is_some_and(|registered| Arc::ptr_eq(registered, flight));
            if current {
                state.flights.remove(&key);
                if state.entries.try_reserve(1).is_err() {
                    FlightOutcome::Cache(CacheError::Allocation(AllocationError {
                        kind: AllocationKind::Entries,
                        requested: 1,
                    }))
                } else {
                    let capacity = self.max_weight - weight;
                    state.evict_until_fit(capacity);
                    let recency = state.next_recency();
                    state.entries.insert(
                        key,
                        Entry {
                            value: Arc::clone(value),
                            weight,
                            last_used: recency,
                            renumbered: false,
                        },
                    );
                    state.total_weight += weight;
                    FlightOutcome::Success(Arc::clone(value))
                }
            } else {
                FlightOutcome::Success(Arc::clone(value))
            }
        };
        flight.complete(outcome.clone());
        outcome.into_result()
    }

    fn complete_detached_flight(
        &self,
        key: &K,
        flight: &Arc<Flight<V>>,
        outcome: FlightOutcome<V>,
    ) {
        let mut state = lock(&self.state);
        if state
            .flights
            .get(key)
            .is_some_and(|registered| Arc::ptr_eq(registered, flight))
        {
            state.flights.remove(key);
        }
        drop(state);
        flight.complete(outcome);
    }

    fn validate_weight(&self, weight: usize) -> Result<(), CacheError> {
        if weight == 0 {
            return Err(CacheError::Weight(WeightError::ZeroWeight));
        }
        if weight > self.max_weight {
            return Err(CacheError::Weight(WeightError::ExceedsCapacity {
                weight,
                capacity: self.max_weight,
            }));
        }
        Ok(())
    }
}

impl<V> FlightOutcome<V> {
    fn into_result(self) -> Result<Arc<V>, GetOrInsertError> {
        match self {
            Self::Success(value) => Ok(value),
            Self::Cache(error) => Err(GetOrInsertError::Cache(error)),
            Self::Parse(error) => Err(GetOrInsertError::Parse(error)),
            Self::ParserPanicked => Err(GetOrInsertError::ParserPanicked),
        }
    }
}

fn allocation_error(kind: AllocationKind) -> CacheError {
    CacheError::Allocation(AllocationError { kind, requested: 1 })
}

fn lock<T>(mutex: &Mutex<T>) -> MutexGuard<'_, T> {
    mutex
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner)
}
