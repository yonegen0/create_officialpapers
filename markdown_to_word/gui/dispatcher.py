"""Workers enqueue callbacks; only the GUI thread touches Tk."""
from queue import Empty, SimpleQueue
from collections.abc import Callable
import tkinter as tk


class Dispatcher:
    def __init__(self, root: tk.Misc) -> None:
        self._root = root
        self._queue: SimpleQueue[Callable[[], None]] = SimpleQueue()
        self._running = True
        self._tick()

    def post(self, callback: Callable[[], None]) -> None:
        self._queue.put(callback)

    def _tick(self) -> None:
        if not self._running:
            return
        while True:
            try:
                callback = self._queue.get_nowait()
            except Empty:
                break
            callback()
        self._root.after(75, self._tick)

    def close(self) -> None:
        self._running = False
