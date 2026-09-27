# 0795 multi event Callgrind analysis

This is an offline replay of the twelve retained Callgrind captures. The event columns are guest counter diagnostics and do not measure native latency.

| Repeat | Shape | Leg | Ir | Bc | Bcm | Bi | Bim | Owner qualified |
| ---: | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| 0 | tiny | before | 7,758,062 | 351,680 | 25,552 | 34,178 | 1,663 | yes |
| 0 | tiny | after | 7,745,872 | 344,394 | 23,584 | 33,682 | 2,094 | yes |
| 0 | medium | before | 13,724,056 | 872,978 | 58,128 | 90,953 | 4,361 | yes |
| 0 | medium | after | 13,706,121 | 859,366 | 54,455 | 89,691 | 5,460 | yes |
| 0 | large | before | 549,368,835 | 45,300,910 | 2,956,469 | 5,586,687 | 229,969 | yes |
| 0 | large | after | 547,207,829 | 44,356,083 | 2,753,078 | 5,495,777 | 310,633 | yes |
| 1 | large | after | 547,227,600 | 44,357,335 | 2,747,449 | 5,495,831 | 310,633 | yes |
| 1 | large | before | 549,351,049 | 45,296,412 | 2,955,275 | 5,586,916 | 230,149 | yes |
| 1 | medium | after | 13,706,059 | 859,113 | 54,609 | 89,807 | 5,459 | yes |
| 1 | medium | before | 13,731,167 | 874,197 | 58,075 | 90,995 | 4,354 | yes |
| 1 | tiny | after | 7,744,819 | 344,147 | 23,656 | 33,697 | 2,095 | yes |
| 1 | tiny | before | 7,761,900 | 352,389 | 25,668 | 34,225 | 1,658 | yes |

For every positive dump the parser sums every function self vector and checks it against both `summary:` and `totals:` for all five events. The termination dump is required to be zero for every event.

Immediate owner children are the only inclusive rows used in the disjoint partition. Descendant inclusive rows overlap and are retained for path inspection; no fractions are computed.

The `allocation_named_functions` group is a name-based function census. Its incoming and outgoing graph calls are not allocator API call counts.

## Owner qualification failures

None.

## Selected symbol accounting

### 0-tiny-before

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 379,901 | 9,048,753 | 2,240 | 14,113 |
| `inspect` | 1 | 163,178 | 999,589 | 730 | 9,899 |
| `checked_attributes` | 1 | 65,520 | 315,095 | 1,138 | 2,789 |
| `iter_state` | 2 | 380,247 | 557,452 | 9,985 | 8,661 |
| `allocation_named_functions` | 30 | 194,168 | 610,064 | 24,871 | 23,538 |
| `memcpy` | 2 | 28,354 | 28,354 | 5,835 | 0 |
| `opened_presentation` | 1 | 83 | 7,758,052 | 1 | 5 |

### 0-tiny-after

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 384,234 | 8,965,656 | 2,240 | 13,700 |
| `inspect` | 1 | 167,540 | 983,339 | 730 | 9,487 |
| `checked_attributes` | 1 | 68,301 | 309,785 | 1,138 | 3,764 |
| `iter_state` | 2 | 342,742 | 373,362 | 9,985 | 5,537 |
| `allocation_named_functions` | 30 | 149,794 | 504,650 | 17,474 | 18,396 |
| `memcpy` | 2 | 28,759 | 28,759 | 5,800 | 0 |
| `opened_presentation` | 1 | 83 | 7,745,862 | 1 | 5 |

### 0-medium-before

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 954,256 | 14,202,560 | 6,362 | 38,511 |
| `inspect` | 1 | 426,219 | 2,445,352 | 2,350 | 25,534 |
| `checked_attributes` | 1 | 144,963 | 664,857 | 2,533 | 12,761 |
| `iter_state` | 2 | 841,146 | 1,220,832 | 27,301 | 17,343 |
| `allocation_named_functions` | 30 | 414,750 | 1,335,631 | 45,055 | 42,916 |
| `memcpy` | 2 | 60,393 | 60,393 | 11,464 | 0 |
| `opened_presentation` | 1 | 83 | 13,724,046 | 1 | 5 |

### 0-medium-after

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 964,180 | 14,093,130 | 6,362 | 37,549 |
| `inspect` | 1 | 436,172 | 2,411,157 | 2,350 | 24,573 |
| `checked_attributes` | 1 | 150,417 | 656,083 | 2,533 | 16,202 |
| `iter_state` | 2 | 763,258 | 831,318 | 27,301 | 11,135 |
| `allocation_named_functions` | 30 | 316,883 | 1,140,165 | 30,866 | 33,381 |
| `memcpy` | 2 | 61,965 | 61,965 | 11,481 | 0 |
| `opened_presentation` | 1 | 83 | 13,706,111 | 1 | 5 |

### 0-large-before

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 61,568,120 | 515,248,503 | 455,386 | 2,665,801 |
| `inspect` | 1 | 29,165,311 | 158,609,830 | 181,678 | 1,731,532 |
| `checked_attributes` | 1 | 8,660,171 | 37,176,448 | 152,413 | 1,118,441 |
| `iter_state` | 2 | 42,805,826 | 63,319,582 | 1,804,245 | 883,343 |
| `allocation_named_functions` | 30 | 8,102,066 | 20,247,070 | 1,132,615 | 1,046,339 |
| `memcpy` | 2 | 1,620,404 | 1,620,404 | 395,945 | 0 |
| `opened_presentation` | 1 | 83 | 549,368,825 | 1 | 5 |

### 0-large-after

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 62,177,777 | 511,010,705 | 455,386 | 2,604,981 |
| `inspect` | 1 | 29,774,878 | 156,534,287 | 181,678 | 1,670,707 |
| `checked_attributes` | 1 | 8,936,993 | 36,682,147 | 152,413 | 1,391,226 |
| `iter_state` | 2 | 38,818,090 | 42,521,990 | 1,804,245 | 556,335 |
| `allocation_named_functions` | 30 | 1,961,540 | 7,285,482 | 396,876 | 554,653 |
| `memcpy` | 2 | 1,635,722 | 1,635,722 | 396,362 | 0 |
| `opened_presentation` | 1 | 83 | 547,207,819 | 1 | 5 |

### 1-large-after

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 62,177,639 | 511,008,423 | 455,386 | 2,604,978 |
| `inspect` | 1 | 29,774,878 | 156,536,077 | 181,678 | 1,670,707 |
| `checked_attributes` | 1 | 8,936,993 | 36,682,011 | 152,413 | 1,391,226 |
| `iter_state` | 2 | 38,818,090 | 42,521,990 | 1,804,245 | 556,335 |
| `allocation_named_functions` | 30 | 1,993,717 | 7,501,232 | 397,225 | 555,711 |
| `memcpy` | 2 | 1,631,558 | 1,631,558 | 396,369 | 0 |
| `opened_presentation` | 1 | 83 | 547,227,590 | 1 | 5 |

### 1-large-before

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 61,568,038 | 515,249,869 | 455,386 | 2,665,801 |
| `inspect` | 1 | 29,165,311 | 158,612,395 | 181,678 | 1,731,532 |
| `checked_attributes` | 1 | 8,660,171 | 37,177,006 | 152,413 | 1,118,441 |
| `iter_state` | 2 | 42,805,826 | 63,320,698 | 1,804,245 | 883,343 |
| `allocation_named_functions` | 30 | 8,088,324 | 20,171,090 | 1,132,624 | 1,046,256 |
| `memcpy` | 2 | 1,623,315 | 1,623,315 | 396,057 | 0 |
| `opened_presentation` | 1 | 83 | 549,351,039 | 1 | 5 |

### 1-medium-after

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 964,203 | 14,089,111 | 6,362 | 37,549 |
| `inspect` | 1 | 436,172 | 2,411,287 | 2,350 | 24,573 |
| `checked_attributes` | 1 | 150,417 | 655,837 | 2,533 | 16,202 |
| `iter_state` | 2 | 763,258 | 831,318 | 27,301 | 11,135 |
| `allocation_named_functions` | 30 | 316,389 | 1,129,837 | 30,683 | 33,195 |
| `memcpy` | 2 | 61,844 | 61,844 | 11,465 | 0 |
| `opened_presentation` | 1 | 83 | 13,706,049 | 1 | 5 |

### 1-medium-before

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 954,256 | 14,206,589 | 6,362 | 38,511 |
| `inspect` | 1 | 426,219 | 2,444,766 | 2,350 | 25,534 |
| `checked_attributes` | 1 | 144,963 | 664,816 | 2,533 | 12,761 |
| `iter_state` | 2 | 841,146 | 1,220,750 | 27,301 | 17,343 |
| `allocation_named_functions` | 30 | 420,183 | 1,368,627 | 45,142 | 43,131 |
| `memcpy` | 2 | 60,376 | 60,376 | 11,431 | 0 |
| `opened_presentation` | 1 | 83 | 13,731,157 | 1 | 5 |

### 1-tiny-after

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 384,234 | 8,957,610 | 2,240 | 13,700 |
| `inspect` | 1 | 167,540 | 983,329 | 730 | 9,487 |
| `checked_attributes` | 1 | 68,301 | 309,737 | 1,138 | 3,764 |
| `iter_state` | 2 | 342,742 | 373,362 | 9,985 | 5,537 |
| `allocation_named_functions` | 30 | 148,858 | 498,816 | 17,533 | 18,441 |
| `memcpy` | 2 | 28,710 | 28,710 | 5,808 | 0 |
| `opened_presentation` | 1 | 83 | 7,744,809 | 1 | 5 |

### 1-tiny-before

| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |
| --- | ---: | ---: | ---: | ---: | ---: |
| `notes` | 25 | 379,901 | 9,057,131 | 2,240 | 14,113 |
| `inspect` | 1 | 163,178 | 999,275 | 730 | 9,899 |
| `checked_attributes` | 1 | 65,520 | 314,909 | 1,138 | 2,789 |
| `iter_state` | 2 | 380,247 | 557,080 | 9,985 | 8,661 |
| `allocation_named_functions` | 30 | 196,918 | 623,191 | 24,948 | 23,603 |
| `memcpy` | 2 | 28,183 | 28,183 | 5,826 | 0 |
| `opened_presentation` | 1 | 83 | 7,761,890 | 1 | 5 |

Raw dumps, termination dumps, reports, and receipt artifact hashes are bound in `callgrind-analysis.json`.
