use litchi_iwa_detect::{__classify_application_payload, Error, Format};
use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

struct CountingAllocator;

thread_local! {
    static ALLOCATIONS: Cell<Option<usize>> = const { Cell::new(None) };
}

#[global_allocator]
static GLOBAL: CountingAllocator = CountingAllocator;

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let _ = ALLOCATIONS.try_with(|counter| {
            if let Some(count) = counter.get() {
                counter.set(Some(count.saturating_add(1)));
            }
        });
        // SAFETY: The test allocator forwards the layout unchanged to the
        // platform allocator.
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: The pointer and layout came from `alloc` above.
        unsafe { System.dealloc(pointer, layout) }
    }
}

fn append_length_delimited(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    litchi_iwa_common::encode_varint_into(output, (u64::from(number) << 3) | 2);
    litchi_iwa_common::encode_varint_into(output, u64::try_from(payload.len()).unwrap());
    output.extend_from_slice(payload);
}

fn reference(identifier: u8) -> [u8; 2] {
    [0x08, identifier]
}

fn shared_document() -> [u8; 2] {
    [0x0a, 0]
}

fn payload(format: Format) -> Vec<u8> {
    let shared = shared_document();
    let mut output = Vec::new();
    match format {
        Format::Pages => append_length_delimited(&mut output, 15, &shared),
        Format::Numbers => {
            let first = reference(1);
            let second = reference(2);
            let third = reference(3);
            append_length_delimited(&mut output, 4, &first);
            append_length_delimited(&mut output, 5, &second);
            append_length_delimited(&mut output, 6, &third);
            append_length_delimited(&mut output, 8, &shared);
        },
        Format::Keynote => {
            let show = reference(1);
            append_length_delimited(&mut output, 2, &show);
            append_length_delimited(&mut output, 3, &shared);
        },
    }
    output
}

#[test]
fn canonical_classification_is_allocation_free_and_preserves_wire_policy() {
    for format in [Format::Pages, Format::Numbers, Format::Keynote] {
        let source = payload(format);
        ALLOCATIONS.with(|counter| counter.set(Some(0)));
        let classified = __classify_application_payload(&source).unwrap();
        let allocations = ALLOCATIONS.with(|counter| counter.take().unwrap());
        assert_eq!(classified, Some(format));
        assert_eq!(
            allocations, 0,
            "{format:?} canonical classifier allocations"
        );
    }

    // High-field roots and nested shared documents must not allocate an
    // input-sized index. Unknown scalar values retain their original policy.
    for nested in [false, true] {
        let unknown_fields = [0x80, 0x01, 0x80, 0x00].repeat(if nested { 100_000 } else { 99_996 });
        let mut source = if nested {
            Vec::new()
        } else {
            payload(Format::Numbers)
        };
        if nested {
            let mut shared = Vec::new();
            append_length_delimited(&mut shared, 1, &unknown_fields);
            append_length_delimited(&mut source, 15, &shared);
        } else {
            source.extend_from_slice(&unknown_fields);
        }
        ALLOCATIONS.with(|counter| counter.set(Some(0)));
        let result = __classify_application_payload(&source).unwrap();
        let allocations = ALLOCATIONS.with(|counter| counter.take().unwrap());
        assert_eq!(
            result,
            Some(if nested {
                Format::Pages
            } else {
                Format::Numbers
            })
        );
        assert_eq!(allocations, 0);
        if !nested {
            source.push(0x80);
            assert!(matches!(
                __classify_application_payload(&source),
                Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
            ));
            source.extend_from_slice(&[0x01, 0x00]);
            assert!(matches!(
                __classify_application_payload(&source),
                Err(Error::LimitExceeded {
                    kind: litchi_iwa_detect::LimitKind::IwaFields,
                    observed: 100_001,
                    maximum: 100_000
                })
            ));
        }
    }

    let mut unknown = payload(Format::Pages);
    // Unknown canonical scalar fields remain opaque to the application shape.
    unknown.extend_from_slice(&[0x80, 0x01, 0x00]);
    assert_eq!(
        __classify_application_payload(&unknown).unwrap(),
        Some(Format::Pages)
    );

    let mut duplicate = payload(Format::Pages);
    duplicate.extend_from_slice(&payload(Format::Pages));
    assert!(matches!(
        __classify_application_payload(&duplicate),
        Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
    ));

    let noncanonical_key = [0xfa, 0x00, 0x02, 0x0a, 0x00];
    assert!(matches!(
        __classify_application_payload(&noncanonical_key),
        Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
    ));

    let noncanonical_length = [0x7a, 0x82, 0x00, 0x0a, 0x00];
    assert!(matches!(
        __classify_application_payload(&noncanonical_length),
        Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
    ));
}
