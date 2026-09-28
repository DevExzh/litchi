# 0827 ordinary-save comparison

Offline replay of full ordinary-save before/after evidence.
Native timing and observer counters are separate; this packet records flags and makes no adoption decision.

Reports: 216  Samples: 4488

| case | p50 before | p50 after | after/before | CI95 | RSS ratio |
|---|---:|---:|---:|---|---:|
| docx_real_file_ordinary_save_lifecycle | 5274724 | 5272876 | 0.999293521 | [0.993437332, 1.00235757] | 0.999985412 |
| docx_real_file_ordinary_save_edit | 56015 | 56167.5 | 1.00204899 | [0.994976443, 1.00772328] | 1.00083137 |
| docx_real_file_ordinary_save_atomic_publish | 5088365 | 5098148.5 | 1.00198876 | [0.996771702, 1.00587984] | 1.00062738 |
| docx_real_file_ordinary_save_counting_publish | 51980 | 52310 | 1.00635484 | [0.989117094, 1.02190312] | 0.999737456 |
| xlsx_real_file_ordinary_save_lifecycle | 5464522 | 5479965 | 1.00105041 | [0.994168442, 1.00628567] | 1.00068628 |
| xlsx_real_file_ordinary_save_edit | 263688.5 | 264493.5 | 0.999528245 | [0.973969636, 1.01940528] | 0.999737597 |
| xlsx_real_file_ordinary_save_atomic_publish | 5000232.5 | 5017995 | 1.00288731 | [1.00073304, 1.00470987] | 0.999985361 |
| xlsx_real_file_ordinary_save_counting_publish | 64937.5 | 66000 | 1.01253929 | [0.99653166, 1.0298117] | 1.00000009 |
| pptx_real_file_ordinary_save_lifecycle | 7474617.5 | 7304064 | 0.975058135 | [0.973040368, 0.982245857] | 1.00011672 |
| pptx_real_file_ordinary_save_edit | 1412057 | 1249406 | 0.884764806 | [0.882642279, 0.887018713] | 1.00002917 |
| pptx_real_file_ordinary_save_atomic_publish | 5693326 | 5689858 | 1.00051814 | [0.996739266, 1.00387571] | 1.00016049 |
| pptx_real_file_ordinary_save_counting_publish | 304328.5 | 302829 | 0.993334612 | [0.989624787, 1.00706987] | 1.00023327 |

Observer allocation vectors and diagnostic process counters are retained in analysis.json.
No unsupported general full-save, additive-phase, pooled-lane, or adoption claim is made.
