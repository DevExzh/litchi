"""Finish the baseline lanes before the candidate source is applied."""
import run
from inspect_assembly import inspect

run.configure('baseline')
run.build('alloc')
inspect('baseline')
run.capture('profile')
run.capture('hardware')
run.capture('alloc')
run.capture('native', 1)
