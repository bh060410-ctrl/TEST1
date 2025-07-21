
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import hello


def test_solution_simple():
    assert hello.solution(2, 3) == 5


def test_solution_large():
    assert hello.solution(100, 2) == 102
