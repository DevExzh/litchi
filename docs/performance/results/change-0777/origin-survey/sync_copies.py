"""Regenerate the private copies of litchi-opc's xml_attributes module.

The copies differ from the canonical module only in their module docs, the
visibility of the trait and iterator (private copies), and the path of the
shared test module."""
import sys
W = '/home/zhuhe/code/litchi-worktrees/0770-quick-xml-fail-fast-attribute-checks/'
canon = open(W + 'crates/litchi-opc/src/xml_attributes.rs').read()
body = canon[canon.index('use std::borrow::Cow;'):]
tail = '#[cfg(test)]\nmod tests;\n'
assert body.endswith(tail)
body = body[:-len(tail)]
copies = {
    'crates/litchi-ole-common/src/xml_attributes.rs': ('pub', 'the OLE2 crates'),
    'crates/litchi-sign/src/xml_attributes.rs': ('pub(crate)', 'this crate'),
    'crates/litchi-xldm/src/xml_attributes.rs': ('pub(crate)', 'this crate'),
    'crates/xml-minifier/src/xml_attributes.rs': ('pub(crate)', 'this crate'),
}
check = '--check' in sys.argv
bad = 0
for path, (vis, users) in copies.items():
    doc = f"""//! Attribute iteration for readers that stop at a start tag's first attribute
//! error, with a worst case that stays bounded on hostile tags.
//!
//! This is a copy of `litchi_opc::xml_attributes` (record 0770) for {users},
//! which may not depend on `litchi-opc` (`tools/crate_boundaries.json`). Keep
//! the two in step: this module compiles the same tests, from
//! `litchi-opc/src/xml_attributes/tests.rs`, against its own code.
//!
//! quick-xml 0.41 checks a start tag's attribute names for duplicates by
//! default: a linear scan of the names before each one while there are at
//! most 32, then a hash pre-filter whose hasher is not keyed, with a scan of
//! every earlier name on each pre-filter hit, so its worst case grows with the
//! square of the tag. [`BytesStartExt::checked_attributes`] yields what
//! quick-xml's checked iterator yields up to and including its first error,
//! and nothing after it; quick-xml checks the first 32 names and this module
//! checks the rest in an ordered map, `O(log n)` comparisons per name and no
//! hashing.

"""
    text = doc + body
    if vis != 'pub':
        assert text.count('pub trait BytesStartExt') == 1 and text.count('pub struct CheckedAttributes') == 1
        text = text.replace('pub trait BytesStartExt', 'pub(crate) trait BytesStartExt')
        text = text.replace('pub struct CheckedAttributes', 'pub(crate) struct CheckedAttributes')
    text += '#[cfg(test)]\n#[path = "../../litchi-opc/src/xml_attributes/tests.rs"]\nmod tests;\n'
    if check:
        if open(W + path).read() != text:
            print('OUT OF STEP', path); bad += 1
    else:
        open(W + path, 'w').write(text)
print('copies', 'checked' if check else 'written', len(copies), 'bad', bad)
sys.exit(1 if bad else 0)
