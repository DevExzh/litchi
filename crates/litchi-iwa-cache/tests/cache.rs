use std::error::Error;
use std::fmt;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Barrier};
use std::thread;

use litchi_iwa_cache::{CacheError, GetOrInsertError, ParseError, WeightError, WeightedCache};

#[derive(Debug)]
struct TestError(&'static str);

impl fmt::Display for TestError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.0)
    }
}

impl Error for TestError {}

fn boxed_error(message: &'static str) -> ParseError {
    Box::new(TestError(message))
}

#[test]
fn rejects_invalid_capacity_and_weights() {
    assert!(matches!(
        WeightedCache::<&str, usize>::new(0),
        Err(WeightError::ZeroCapacity)
    ));
    assert!(matches!(
        WeightedCache::<&str, usize>::new_with_limits(1, 0),
        Err(WeightError::ZeroFlightCapacity)
    ));

    let cache = cache(3);
    assert_eq!(
        cache.insert("zero", Arc::new(0), 0),
        Err(CacheError::Weight(WeightError::ZeroWeight))
    );
    assert_eq!(
        cache.insert("large", Arc::new(0), 4),
        Err(CacheError::Weight(WeightError::ExceedsCapacity {
            weight: 4,
            capacity: 3,
        }))
    );
    assert!(cache.is_empty());
    assert_eq!(cache.total_weight(), 0);
}

#[test]
fn weighted_lru_eviction_is_deterministic() {
    let cache = cache(5);
    insert(&cache, "a", 2, 2);
    insert(&cache, "b", 2, 2);
    insert(&cache, "c", 1, 1);

    assert_eq!(cache.get(&"a").as_deref(), Some(&2));
    insert(&cache, "d", 3, 3);

    assert!(cache.get(&"b").is_none());
    assert_eq!(cache.get(&"a").as_deref(), Some(&2));
    assert!(cache.get(&"c").is_none());
    assert_eq!(cache.get(&"d").as_deref(), Some(&3));
    assert_eq!(cache.len(), 2);
    assert_eq!(cache.total_weight(), 5);
}

#[test]
fn replacement_and_explicit_invalidation_update_weight() {
    let cache = cache(5);
    insert(&cache, "a", 1, 2);
    insert(&cache, "b", 2, 2);
    insert(&cache, "a", 7, 4);

    assert_eq!(cache.get(&"a").as_deref(), Some(&7));
    assert!(cache.get(&"b").is_none());
    assert_eq!(cache.total_weight(), 4);

    assert!(cache.invalidate(&"a"));
    assert!(!cache.invalidate(&"a"));
    assert!(cache.is_empty());

    insert(&cache, "c", 3, 1);
    cache.clear();
    assert!(cache.is_empty());
    assert_eq!(cache.total_weight(), 0);
}

#[test]
fn concurrent_callers_share_one_arc_and_one_parse() {
    const CALLERS: usize = 8;
    let cache = Arc::new(cache(4));
    let ready = Arc::new(Barrier::new(CALLERS));
    let parser_started = Arc::new(Barrier::new(2));
    let parse_count = Arc::new(AtomicUsize::new(0));
    let handles = (0..CALLERS)
        .map(|_| {
            let cache_handle = Arc::clone(&cache);
            let ready_barrier = Arc::clone(&ready);
            let parser_started_barrier = Arc::clone(&parser_started);
            let parse_count_ref = Arc::clone(&parse_count);
            thread::spawn(move || {
                ready_barrier.wait();
                cache_handle.get_or_insert_with("key", 2, || {
                    if parse_count_ref.fetch_add(1, Ordering::SeqCst) == 0 {
                        parser_started_barrier.wait();
                    }
                    42
                })
            })
        })
        .collect::<Vec<_>>();
    parser_started.wait();

    let results = handles.into_iter().map(join_success).collect::<Vec<_>>();
    let first = results.first().map_or_else(|| Arc::new(0), Arc::clone);
    assert!(results.iter().all(|value| Arc::ptr_eq(&first, value)));
    assert!(results.iter().all(|value| **value == 42));
    assert_eq!(parse_count.load(Ordering::SeqCst), 1);
    assert_eq!(cache.total_weight(), 2);
}

#[test]
fn failed_parse_is_not_cached_and_allows_retry() {
    let cache = cache(4);
    let parse_count = Arc::new(AtomicUsize::new(0));
    let first = cache.get_or_try_insert_with("key", 1, || {
        parse_count.fetch_add(1, Ordering::SeqCst);
        Err::<usize, ParseError>(boxed_error("synthetic parse failure"))
    });
    match first {
        Err(GetOrInsertError::Parse(error)) => {
            assert_eq!(error.to_string(), "synthetic parse failure");
        },
        other => panic!("expected parse error, got {other:?}"),
    }
    assert_eq!(parse_count.load(Ordering::SeqCst), 1);

    let retry = cache
        .get_or_try_insert_with("key", 1, || {
            parse_count.fetch_add(1, Ordering::SeqCst);
            Ok::<usize, ParseError>(9)
        })
        .unwrap_or_else(|error| panic!("retry unexpectedly failed: {error}"));
    assert_eq!(*retry, 9);
    assert_eq!(parse_count.load(Ordering::SeqCst), 2);
}

#[test]
fn active_parser_flights_are_bounded() {
    let cache = Arc::new(
        WeightedCache::<&'static str, usize>::new_with_limits(4, 1)
            .unwrap_or_else(|error| panic!("test cache construction failed: {error}")),
    );
    let parser_started = Arc::new(Barrier::new(2));
    let release_parser = Arc::new(Barrier::new(2));
    let owner_cache = Arc::clone(&cache);
    let owner_started = Arc::clone(&parser_started);
    let owner_release = Arc::clone(&release_parser);
    let owner = thread::spawn(move || {
        owner_cache.get_or_insert_with("first", 1, || {
            owner_started.wait();
            owner_release.wait();
            1
        })
    });

    parser_started.wait();
    assert!(matches!(
        cache.get_or_insert_with("second", 1, || 2),
        Err(GetOrInsertError::Cache(CacheError::FlightsLimit {
            active: 1,
            limit: 1,
        }))
    ));
    release_parser.wait();
    assert_eq!(*join_success(owner), 1);
}

#[test]
fn reported_weights_are_validated_after_parsing() {
    let cache = cache(4);
    assert!(matches!(
        cache.get_or_try_insert_with_weight("large", || {
            Ok::<(usize, usize), ParseError>((7, 5))
        }),
        Err(GetOrInsertError::Cache(CacheError::Weight(
            WeightError::ExceedsCapacity {
                weight: 5,
                capacity: 4,
            }
        )))
    ));
    let value = cache
        .get_or_try_insert_with_weight("large", || Ok::<(usize, usize), ParseError>((9, 4)))
        .unwrap_or_else(|error| panic!("weighted retry unexpectedly failed: {error}"));
    assert_eq!(*value, 9);
    assert_eq!(cache.total_weight(), 4);
}

#[test]
fn invalidation_detaches_old_parser_generation() {
    let cache = Arc::new(cache(4));
    let parser_started = Arc::new(Barrier::new(2));
    let release_parser = Arc::new(Barrier::new(2));
    let parse_count = Arc::new(AtomicUsize::new(0));
    let owner_cache = Arc::clone(&cache);
    let owner_started = Arc::clone(&parser_started);
    let owner_release = Arc::clone(&release_parser);
    let owner_count = Arc::clone(&parse_count);
    let owner = thread::spawn(move || {
        owner_cache.get_or_insert_with("key", 1, || {
            owner_count.fetch_add(1, Ordering::SeqCst);
            owner_started.wait();
            owner_release.wait();
            1
        })
    });

    parser_started.wait();
    assert!(cache.invalidate(&"key"));
    release_parser.wait();
    let stale = join_success(owner);
    assert_eq!(*stale, 1);
    assert!(cache.get(&"key").is_none());

    let fresh = cache
        .get_or_insert_with("key", 1, || 2)
        .unwrap_or_else(|error| panic!("fresh parse unexpectedly failed: {error}"));
    assert_eq!(*fresh, 2);
    assert_eq!(parse_count.load(Ordering::SeqCst), 1);
}

fn cache(max_weight: usize) -> WeightedCache<&'static str, usize> {
    WeightedCache::new(max_weight)
        .unwrap_or_else(|error| panic!("test cache construction failed: {error}"))
}

fn insert(
    cache: &WeightedCache<&'static str, usize>,
    key: &'static str,
    value: usize,
    weight: usize,
) {
    cache
        .insert(key, Arc::new(value), weight)
        .unwrap_or_else(|error| panic!("test insertion failed: {error}"));
}

fn join_success<T>(handle: thread::JoinHandle<Result<Arc<T>, GetOrInsertError>>) -> Arc<T> {
    match handle.join() {
        Ok(Ok(value)) => value,
        Ok(Err(error)) => panic!("thread returned cache error: {error}"),
        Err(payload) => {
            drop(payload);
            panic!("cache worker panicked");
        },
    }
}
