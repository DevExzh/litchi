const CAP: usize = 131_072;
const CHUNK: usize = 4096;

#[inline(never)]
pub fn chunked(content: &[u8]) -> bool {
    let Some((&last, prefix)) = content.split_last() else {
        return true;
    };
    let mut bound = 1usize
        + usize::from(!matches!(content[0], b'<' | b'&'))
        + usize::from(matches!(last, b'<' | b'&'));
    // Each position except the last has one adjacent successor. Separate
    // bounded reduction from the cap branch to permit compiler vectorization.
    for (current, next) in prefix.chunks(CHUNK).zip(content[1..].chunks(CHUNK)) {
        let mut subtotal = 0usize;
        for (&byte, &following) in current.iter().zip(next) {
            subtotal += usize::from(matches!(byte, b'<' | b'&'));
            subtotal +=
                usize::from(matches!(byte, b'>' | b';') && !matches!(following, b'<' | b'&'));
        }
        let Some(total) = bound.checked_add(subtotal) else {
            return false;
        };
        if total > CAP {
            return false;
        }
        bound = total;
    }
    true
}
