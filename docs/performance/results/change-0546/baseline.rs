const MAX_SHARED_PROVISIONAL_EVENTS: usize = 131_072;
#[inline(never)]
pub fn baseline(content: &[u8]) -> bool {
    let mut bound = 1usize; // The reader emits one terminal `Event::Eof`.

    if let Some(&first) = content.first()
        && !matches!(first, b'<' | b'&')
        && !add_shared_event_bound(&mut bound)
    {
        return false;
    }

    // Every markup or general reference event begins at one of these bytes.
    for _ in memchr::memchr2_iter(b'<', b'&', content) {
        if !add_shared_event_bound(&mut bound) {
            return false;
        }
    }

    // A text event can begin after markup (`>`) or a completed reference
    // (`;`) when the next source byte is neither another delimiter nor EOF.
    // Delimiters inside attributes, comments, CDATA, and ordinary text make
    // this deliberately over-count, which only causes a safe fallback.
    for index in memchr::memchr2_iter(b'>', b';', content) {
        if content
            .get(index + 1)
            .is_some_and(|&next| !matches!(next, b'<' | b'&'))
            && !add_shared_event_bound(&mut bound)
        {
            return false;
        }
    }

    true
}

#[inline]
fn add_shared_event_bound(bound: &mut usize) -> bool {
    let Some(next) = bound.checked_add(1) else {
        return false;
    };
    if next > MAX_SHARED_PROVISIONAL_EVENTS {
        return false;
    }
    *bound = next;
    true
}
