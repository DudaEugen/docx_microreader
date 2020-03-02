from typing import Tuple


class ContentInf:

    def __init__(self, text: str, position: Tuple[int, int]):
        self.content = text
        self.begin = position[0]
        self.end = position[1]

    def __str__(self):
        return self.content
