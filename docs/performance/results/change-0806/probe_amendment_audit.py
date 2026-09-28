"""Check the exact scope of the two retained baseline setup repairs."""
import hashlib
import custody as c

p = c.P
initial = p / 'probe-src-setup-0/src'
current = p / 'probe-src/src'

def code_lines(value):
    return [line.strip() for line in value.splitlines()
            if line.strip() and not line.lstrip().startswith('//')]

old = (initial / 'allocation_metrics.rs').read_text()
new = (current / 'allocation_metrics.rs').read_text()
allow = '#[cfg_attr(not(test), allow(dead_code))]'
fallback = '#[cfg(not(feature = "allocator-metrics"))]'
assert new.count(allow) == 7 and old.count(allow) == 0
assert new.count(fallback) == old.count(fallback) + 1
assert fallback + '\npub(crate) fn unavailable_sample' in new
assert code_lines(old) == code_lines(new.replace(allow, '').replace(
    fallback + '\npub(crate) fn unavailable_sample',
    'pub(crate) fn unavailable_sample'))

old = (initial / 'counting_allocator.rs').read_text()
new = (current / 'counting_allocator.rs').read_text()
assert old.count('#[global_allocator]') == 1
assert new.count('#[cfg_attr(not(test), global_allocator)]') == 1
assert code_lines(old) == code_lines(new.replace(
    '#[cfg_attr(not(test), global_allocator)]', '#[global_allocator]'))

old = (initial / 'main.rs').read_text()
new = (current / 'main.rs').read_text()
marker = '#[cfg(test)]\nmod tests {'
start = old.index(marker)
end = old.index('\nfn identity(', start)
old_tests = old[start:end].strip()
new_start = new.index(marker)
round_trip = new.index('    #[test]\n    fn cli_shape_identities_round_trip_through_serde')
assert hashlib.sha256(new[round_trip:].encode()).hexdigest() == (
    '942595b375a60d421c7638d8c8bbd6e9b8283350d1b7c0727471ffb211f81f98')
assert (new[new_start:round_trip].rstrip() + '\n}').strip() == old_tests
assert '"valid-4attr"' in new[round_trip:]
assert 'assert_eq!(encoded, format!("\\\"{cli_name}\\\""));' in new[round_trip:]
assert 'assert_eq!(decoded, shape);' in new[round_trip:]
test_derive = '#[cfg_attr(test, derive(serde::Deserialize, PartialEq, Eq))]'
shape_name = '#[serde(rename = "valid-4attr")]'
assert new.count(test_derive) == new.count(shape_name) == 1
assert code_lines(old[:start] + old[end:]) == code_lines(
    new[:new_start].replace(test_derive, '').replace(shape_name, ''))

c.write(p / 'probe-amendment-audit.json', {
    'schema': 'litchi.performance.0806.probe-amendment-audit.v1',
    'passed': True,
    'reader': c.artifact(__file__),
    'files': {name: {'initial': c.artifact(initial / name),
                     'current': c.artifact(current / name)}
              for name in ['main.rs', 'allocation_metrics.rs', 'counting_allocator.rs']},
    'changes': ['test-only allocator registration isolation',
                'feature gate on unused fallback sample helper',
                'seven non-test dead-code allowances on retained support items',
                'unchanged test module relocated to file end',
                'explicit valid-4attr serialization name',
                'six-shape round-trip test with test-only derives',
                'comments and module documentation whitespace'],
    'new_round_trip_test_sha256': hashlib.sha256(new[round_trip:].encode()).hexdigest(),
    'counter_logic_and_layout_preserved': True,
    'fixture_logic_and_tests_preserved': True,
})
print('Probe amendment scope verified: counter and fixture logic preserved')
