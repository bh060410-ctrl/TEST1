def solution(num1: int, num2: int) -> int:
    """Return the sum of two integers."""
    return num1 + num2


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 3:
        a = int(sys.argv[1])
        b = int(sys.argv[2])
        print(solution(a, b))
    else:
        print("Usage: python hello.py <num1> <num2>")
