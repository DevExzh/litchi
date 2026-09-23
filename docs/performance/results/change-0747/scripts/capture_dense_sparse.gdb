set pagination off
set confirm off
set print thread-events off
set print inferior-events off
set $n = 0
break xml_minifier::audit::verify_with_policy
commands 1
  silent
  set $n = $n + 1
  printf "HIT %d len=%lu policy=%lx\n", $n, $rdx, $r8
  bt 40
  eval "dump binary memory /home/zhuhe/code/litchi-worktrees/scratch/0747/gdbcap/dense-sparse/hit%03d.bin %lu %lu", $n, $rsi, $rsi + $rdx
  printf "ENDHIT\n"
  continue
end
run --case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape dense-sparse --samples 1 --warmup 0 --json /dev/null
quit
