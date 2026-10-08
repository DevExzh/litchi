# 0976 offline decoder empty-chain transition

Status: **pass**. This is a retained-artifact/offline replay result; no workload or benchmark build was run.

The 0975 corrected stream has 3745 empty samples of 6940. The first is sample index 1389 at 303382.496301; 3734 of those empty samples are raw IPs in the target PF_X interval, with 11 outside. The 0973 corrected control has 0 empty samples.

The 0975 corrected DWFL log has 6940 samples in the script and its target report summary is 6940 samples; see JSON for exact report counts. The first target report failure is log line 70522; libdw error 21 is `address range overlaps an existing module` and build-id error 196610 is `No such file or directory`.

The external cache control replay produced 0 empty samples for 0975 and 0 for 0973. Its 0975 replay headers/times/periods match the retained corrected stream, and the second ASLR run has identical perf-script SHA-256. Target duplicate reuse is admitted only after an exact module probe; details and all artifact hashes are in the JSON receipt.
