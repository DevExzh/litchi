"""Capture final candidate binaries and complete the frozen native ABBA."""
import run
from inspect_assembly import inspect

run.configure('candidate')
run.build('normal')
run.build('alloc')
run.capture('native', 1)
run.capture('native', 2)
run.configure('baseline', 'candidate')
run.capture('native', 2)
run.configure('candidate')
inspect('candidate')
run.capture('profile')
run.capture('hardware')
run.capture('alloc')
