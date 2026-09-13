const CAP: usize = 131_072;
const CHUNK: usize = 64 * 1024;
const SPARSE_HITS: usize = 16;

#[inline(never)]
pub fn counted(content: &[u8]) -> bool {
    let mut bound = 1 + usize::from(content.first().is_some_and(|&b| !matches!(b, b'<' | b'&')));
    let mut remaining = content;
    // Keep delimiter skipping for inputs with few marker bytes. Dense input
    // switches to the pinned memchr single-byte iterator's bulk count path.
    for _ in 0..SPARSE_HITS {
        let Some(index) = memchr::memchr2(b'<', b'&', remaining) else {
            remaining = &[];
            break;
        };
        if !add(&mut bound, 1) {
            return false;
        }
        remaining = &remaining[index + 1..];
    }
    for chunk in remaining.chunks(CHUNK) {
        // Matches are disjoint; the sum is at most CHUNK and cannot overflow.
        let subtotal =
            memchr::memchr_iter(b'<', chunk).count() + memchr::memchr_iter(b'&', chunk).count();
        if !add(&mut bound, subtotal) {
            return false;
        }
    }
    for index in memchr::memchr2_iter(b'>', b';', content) {
        if content
            .get(index + 1)
            .is_some_and(|&next| !matches!(next, b'<' | b'&'))
            && !add(&mut bound, 1)
        {
            return false;
        }
    }
    true
}

#[inline]
fn add(bound: &mut usize, amount: usize) -> bool {
    let Some(next) = bound.checked_add(amount) else {
        return false;
    };
    if next > CAP {
        return false;
    }
    *bound = next;
    true
}
