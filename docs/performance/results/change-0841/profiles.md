# Operation-scoped Callgrind diagnostic

Guest Ir is not native cycles, latency, physical copies or a causal share. Each row is one operation in a fresh process. No profile confidence interval is inferred from two repeats.

| Shape | Repeat | Before Ir | After Ir | Ratio |
| --- | ---: | ---: | ---: | ---: |
| many | 0 | 6168522 | 6084727 | 0.98641571 |
| many | 1 | 6168717 | 6084513 | 0.98634984 |
| unicode | 0 | 54554799 | 51564359 | 0.94518466 |
| unicode | 1 | 54555031 | 51564549 | 0.94518412 |
| payload | 0 | 6885505 | 4302783 | 0.62490449 |
| payload | 1 | 6885708 | 4302777 | 0.62488520 |
| rich | 0 | 236296 | 235970 | 0.99862037 |
| rich | 1 | 236300 | 235967 | 0.99859077 |

## many repeat 0: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `free` | 413752 | 380752 | -33000 |
| `malloc` | 351698 | 321698 | -30000 |
| `__memcpy_avx_unaligned_erms` | 458959 | 446981 | -11978 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 46000 | 42000 | -4000 |
| `_int_malloc` | 605538 | 601938 | -3600 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 519500 | 516700 | -2800 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 797000 | 798000 | 1000 |
| `core::slice::sort::unstable::quicksort::quicksort'2` | 3094 | 3436 | 342 |
| `core::slice::sort::unstable::quicksort::quicksort` | 2597 | 2837 | 240 |
| `hashbrown::map::HashMap<K,V,S,A>::insert` | 7591 | 7637 | 46 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 12906 | 12881 | -25 |
| `core::slice::sort::unstable::ipnsort` | 41 | 21 | -20 |
| `(below main)` | 0 | 0 | 0 |
| `0x0000000000023cc0` | 0 | 0 | 0 |
| `0x00000000049f3680` | 6176 | 6176 | 0 |
| `0x00000000049f36f0` | 2 | 2 | 0 |
| `<alloc::raw_vec::RawVecInner>::finish_grow` | 240 | 240 | 0 |
| `<alloc::vec::Vec<T,A> as core::clone::Clone>::clone` | 28000 | 28000 | 0 |
| `<alloc::vec::Vec<T> as alloc::vec::spec_from_iter_nested::SpecFromIterNested<T,I>>::from_iter` | 1379 | 1379 | 0 |
| `<core::hash::sip::Hasher<S> as core::hash::Hasher>::write` | 6328 | 6328 | 0 |

## many repeat 1: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `free` | 413752 | 380752 | -33000 |
| `malloc` | 351698 | 321698 | -30000 |
| `__memcpy_avx_unaligned_erms` | 458959 | 446959 | -12000 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 46000 | 42000 | -4000 |
| `_int_malloc` | 605538 | 601938 | -3600 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 519500 | 516700 | -2800 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 797000 | 798000 | 1000 |
| `core::slice::sort::unstable::quicksort::quicksort` | 2016 | 2715 | 699 |
| `core::slice::sort::unstable::quicksort::quicksort'2` | 3756 | 3241 | -515 |
| `hashbrown::map::HashMap<K,V,S,A>::insert` | 7647 | 7678 | 31 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 12974 | 12966 | -8 |
| `core::slice::sort::unstable::ipnsort` | 27 | 20 | -7 |
| `hashbrown::raw::RawTable<T,A>::reserve_rehash` | 4988 | 4984 | -4 |
| `(below main)` | 0 | 0 | 0 |
| `0x0000000000023cc0` | 0 | 0 | 0 |
| `0x00000000049f3680` | 6176 | 6176 | 0 |
| `0x00000000049f36f0` | 2 | 2 | 0 |
| `<alloc::raw_vec::RawVecInner>::finish_grow` | 240 | 240 | 0 |
| `<alloc::vec::Vec<T,A> as core::clone::Clone>::clone` | 28000 | 28000 | 0 |
| `<alloc::vec::Vec<T> as alloc::vec::spec_from_iter_nested::SpecFromIterNested<T,I>>::from_iter` | 1379 | 1379 | 0 |

## unicode repeat 0: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `__memcpy_avx_unaligned_erms` | 3603876 | 634572 | -2969304 |
| `_int_malloc` | 97460 | 86180 | -11280 |
| `_int_free_maybe_trim` | 19098 | 17108 | -1990 |
| `unlink_chunk.isra.0` | 10653 | 8923 | -1730 |
| `__libc_malloc2` | 10279 | 8830 | -1449 |
| `free` | 44707 | 43427 | -1280 |
| `_int_free_merge_chunk` | 55645 | 54373 | -1272 |
| `_int_free_chunk` | 13238 | 12133 | -1105 |
| `malloc` | 34546 | 33834 | -712 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 2944 | 2688 | -256 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 41840 | 41680 | -160 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 3045 | 3130 | 85 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 51008 | 51072 | 64 |
| `alloc_perturb` | 805 | 745 | -60 |
| `hashbrown::map::HashMap<K,V,S,A>::insert` | 1491 | 1530 | 39 |
| `mmap` | 78 | 52 | -26 |
| `__memset_avx2_unaligned_erms` | 44241 | 44255 | 14 |
| `__set_vma_name` | 42 | 28 | -14 |
| `_int_realloc` | 26793 | 26807 | 14 |
| `__memcmp_avx2_movbe` | 390 | 378 | -12 |

## unicode repeat 1: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `__memcpy_avx_unaligned_erms` | 3603876 | 634572 | -2969304 |
| `_int_malloc` | 97460 | 86180 | -11280 |
| `_int_free_maybe_trim` | 19098 | 17108 | -1990 |
| `unlink_chunk.isra.0` | 10653 | 8923 | -1730 |
| `__libc_malloc2` | 10279 | 8830 | -1449 |
| `free` | 44707 | 43427 | -1280 |
| `_int_free_merge_chunk` | 55645 | 54373 | -1272 |
| `_int_free_chunk` | 13238 | 12133 | -1105 |
| `malloc` | 34546 | 33834 | -712 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 2944 | 2688 | -256 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 41840 | 41680 | -160 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 3277 | 3354 | 77 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 51008 | 51072 | 64 |
| `alloc_perturb` | 805 | 745 | -60 |
| `mmap` | 78 | 52 | -26 |
| `__memset_avx2_unaligned_erms` | 44241 | 44255 | 14 |
| `__set_vma_name` | 42 | 28 | -14 |
| `_int_realloc` | 26793 | 26807 | 14 |
| `__memcmp_avx2_movbe` | 390 | 378 | -12 |
| `hashbrown::map::HashMap<K,V,S,A>::insert` | 1491 | 1496 | 5 |

## payload repeat 0: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `__memcpy_avx_unaligned_erms` | 5450619 | 2890159 | -2560460 |
| `_int_malloc` | 89927 | 80861 | -9066 |
| `_int_free_merge_chunk` | 55942 | 51213 | -4729 |
| `_int_free_maybe_trim` | 18164 | 16114 | -2050 |
| `unlink_chunk.isra.0` | 9483 | 7971 | -1512 |
| `__libc_malloc2` | 9359 | 7910 | -1449 |
| `free` | 43280 | 42000 | -1280 |
| `_int_free_chunk` | 12541 | 11436 | -1105 |
| `malloc` | 33677 | 32965 | -712 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 2944 | 2688 | -256 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 41840 | 41680 | -160 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 3268 | 3394 | 126 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 51008 | 51072 | 64 |
| `alloc_perturb` | 610 | 565 | -45 |
| `sbrk` | 25 | 0 | -25 |
| `hashbrown::map::HashMap<K,V,S,A>::insert` | 1504 | 1491 | -13 |
| `__memcmp_avx2_movbe` | 402 | 390 | -12 |
| `__glibc_morecore` | 10 | 0 | -10 |
| `_int_realloc` | 26817 | 26807 | -10 |
| `brk` | 9 | 0 | -9 |

## payload repeat 1: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `__memcpy_avx_unaligned_erms` | 5450619 | 2890159 | -2560460 |
| `_int_malloc` | 89927 | 80861 | -9066 |
| `_int_free_merge_chunk` | 55942 | 51213 | -4729 |
| `_int_free_maybe_trim` | 18164 | 16114 | -2050 |
| `unlink_chunk.isra.0` | 9483 | 7971 | -1512 |
| `__libc_malloc2` | 9359 | 7910 | -1449 |
| `free` | 43280 | 42000 | -1280 |
| `_int_free_chunk` | 12541 | 11436 | -1105 |
| `malloc` | 33677 | 32965 | -712 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 2944 | 2688 | -256 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 41840 | 41680 | -160 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 3432 | 3349 | -83 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 51008 | 51072 | 64 |
| `alloc_perturb` | 610 | 565 | -45 |
| `sbrk` | 25 | 0 | -25 |
| `hashbrown::map::HashMap<K,V,S,A>::insert` | 1543 | 1530 | -13 |
| `__memcmp_avx2_movbe` | 402 | 390 | -12 |
| `__glibc_morecore` | 10 | 0 | -10 |
| `_int_realloc` | 26817 | 26807 | -10 |
| `brk` | 9 | 0 | -9 |

## rich repeat 0: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `_int_malloc` | 33409 | 33108 | -301 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 3992 | 4064 | 72 |
| `__libc_malloc2` | 3059 | 3013 | -46 |
| `_int_realloc` | 11327 | 11308 | -19 |
| `unlink_chunk.isra.0` | 2207 | 2188 | -19 |
| `malloc` | 9419 | 9435 | 16 |
| `__memcpy_avx_unaligned_erms` | 9331 | 9321 | -10 |
| `alloc_perturb` | 270 | 260 | -10 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 184 | 176 | -8 |
| `_int_free_merge_chunk` | 14954 | 14948 | -6 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 3232 | 3228 | -4 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 754 | 758 | 4 |
| `realloc` | 14673 | 14677 | 4 |
| `free` | 12328 | 12331 | 3 |
| `0x00000000049f3680` | 256 | 254 | -2 |
| `(below main)` | 0 | 0 | 0 |
| `0x0000000000023cc0` | 0 | 0 | 0 |
| `<alloc::raw_vec::RawVecInner>::finish_grow` | 240 | 240 | 0 |
| `<alloc::string::String as core::fmt::Write>::write_str` | 448 | 448 | 0 |
| `<alloc::vec::Vec<T,A> as core::clone::Clone>::clone` | 112 | 112 | 0 |

## rich repeat 1: largest recorded self-cost differences

A missing separately emitted function is not zero logical work; names can move into callers through inlining. These rows partition recorded self costs; descendant inclusive costs are not added.

| Function | Before self Ir | After self Ir | Delta |
| --- | ---: | ---: | ---: |
| `_int_malloc` | 33409 | 33108 | -301 |
| `litchi_ppt::writer::records::model::InPlaceRecord::write` | 3992 | 4064 | 72 |
| `__libc_malloc2` | 3059 | 3013 | -46 |
| `_int_realloc` | 11327 | 11308 | -19 |
| `unlink_chunk.isra.0` | 2207 | 2188 | -19 |
| `malloc` | 9419 | 9435 | 16 |
| `__memcpy_avx_unaligned_erms` | 9331 | 9321 | -10 |
| `alloc_perturb` | 270 | 260 | -10 |
| `core::ptr::drop_in_place<litchi_ppt::writer::escher::model::UserShapeData>` | 184 | 176 | -8 |
| `_int_free_merge_chunk` | 14954 | 14948 | -6 |
| `litchi_ppt::writer::escher::codec::shapes::append_shape_container_children` | 3232 | 3228 | -4 |
| `realloc` | 14673 | 14677 | 4 |
| `free` | 12328 | 12331 | 3 |
| `litchi_ppt::writer::persist::PersistPtrBuilder::generate_record` | 758 | 755 | -3 |
| `0x00000000049f3680` | 256 | 254 | -2 |
| `(below main)` | 0 | 0 | 0 |
| `0x0000000000023cc0` | 0 | 0 | 0 |
| `<alloc::raw_vec::RawVecInner>::finish_grow` | 240 | 240 | 0 |
| `<alloc::string::String as core::fmt::Write>::write_str` | 448 | 448 | 0 |
| `<alloc::vec::Vec<T,A> as core::clone::Clone>::clone` | 112 | 112 | 0 |
