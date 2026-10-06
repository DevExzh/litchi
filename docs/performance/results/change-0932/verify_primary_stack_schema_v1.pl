use strict;
use warnings;
use JSON::PP;

my $profile = shift @ARGV or die "profile required\n";
my $stack_out = $ENV{STACK_OUT};
my %strings;
my $next_string = 1;
my @alloc_size;
my @alloc_trace;
my %trace_parent;
my %trace_func;
my %ip_func;
my $next_ip = 1;
my $next_trace = 1;
my %trace_calls;
my %trace_bytes;
my %stack_calls;
my %stack_bytes;
my %markers;
my ($calls,$bytes,$minus,$miss,$defs) = (0,0,0,0,0);
open my $z, '-|', '/usr/bin/zstd', '-dc', $profile or die "zstd: $!\n";
while (my $line = <$z>) {
    chomp $line;
    next if $line eq '';
    my ($kind,@f) = split / /, $line;
    $markers{$kind}++;
    if ($kind eq 's') {
        # v3 sized-string format: first field is byte length, not an ID.
        my $text = $line;
        $text =~ s/^s [^ ]+ ?//;
        $strings{$next_string++} = $text;
    } elsif ($kind eq 'i') {
        # i <IP> <module> [function [file line] [inline function file line]...];
        # i records receive implicit sequential IP IDs.
        $ip_func{$next_ip++} = @f >= 3 ? ($strings{hex($f[2])} // "<string:$f[2]>") : '<no-frame>';
    } elsif ($kind eq 't') {
        # t <IP-ID> <parent-trace-index>; t records receive implicit sequential trace IDs.
        my $ip_id = hex($f[0]);
        my $tid = $next_trace++;
        $trace_parent{$tid} = hex($f[1] // 0);
        $trace_func{$tid} = $ip_func{$ip_id} // "<ip:$ip_id>";
    } elsif ($kind eq 'a') {
        # a <size> <trace>: next implicit AllocationInfoIndex.
        $alloc_size[$defs] = hex($f[0]);
        $alloc_trace[$defs] = hex($f[1]);
        $defs++;
    } elsif ($kind eq '+') {
        my $idx = hex($f[0]);
        if ($idx >= $defs || !defined $alloc_size[$idx]) {
            $miss++;
            next;
        }
        my $size = $alloc_size[$idx];
        my $tid = $alloc_trace[$idx];
        $calls++;
        $bytes += $size;
        $trace_calls{$tid}++;
        $trace_bytes{$tid} += $size;
    } elsif ($kind eq '-') {
        $minus++;
    }
}
close $z;
die "zstd failed\n" if $? != 0;
sub stack_for {
    my ($tid) = @_;
    my @stack;
    my %seen;
    while (defined($tid) && $tid != 0 && !$seen{$tid}++) {
        push @stack, ($trace_func{$tid} // "<trace:$tid>");
        last if !exists $trace_parent{$tid};
        $tid = $trace_parent{$tid};
    }
    return @stack ? join(';', @stack) : '<empty>';
}
for my $tid (keys %trace_bytes) {
    my $stack = stack_for(0+$tid);
    $stack_calls{$stack} += $trace_calls{$tid};
    $stack_bytes{$stack} += $trace_bytes{$tid};
}
sub top_rows {
    my ($h) = @_;
    my @keys = sort { $h->{$b} <=> $h->{$a} || $a cmp $b } keys %$h;
    splice @keys, 20 if @keys > 20;
    return \@keys;
}
if (defined $stack_out) {
    open my $sf, ">", $stack_out or die "stack output: $!\n";
    print $sf "stack\tcalls\trequested_bytes\n";
    for my $stack (sort { $stack_bytes{$b} <=> $stack_bytes{$a} || $a cmp $b } keys %stack_bytes) {
        print $sf $stack, "\t", ($stack_calls{$stack} // 0), "\t", ($stack_bytes{$stack} // 0), "\n";
    }
    close $sf or die "stack close: $!\n";
}
my $out = {
    profile => $profile,
    markers => \%markers,
    allocation_info_defs => 0+$defs,
    allocation_events => 0+$calls,
    requested_bytes_total => 0+$bytes,
    deallocation_events => 0+$minus,
    plus_index_misses => 0+$miss,
    trace_definitions => 0+scalar(keys %trace_parent),
    trace_ids_with_allocations => 0+scalar(keys %trace_bytes),
    top_trace_bytes => [map { {trace_id=>0+$_, calls=>0+($trace_calls{$_}//0), requested_bytes=>0+($trace_bytes{$_}//0), stack=>stack_for($_)} } (sort { $trace_bytes{$b}<=>$trace_bytes{$a} || $a<=>$b } keys %trace_bytes)[0..((scalar(keys %trace_bytes)<20?scalar(keys %trace_bytes):20)-1)]],
    top_stack_bytes => [map { {calls=>0+($stack_calls{$_}//0), requested_bytes=>0+$stack_bytes{$_}, stack=>$_} } @{top_rows(\%stack_bytes)}],
};
print JSON::PP->new->canonical->pretty->encode($out);
