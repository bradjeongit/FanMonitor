import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import gc  # noqa: E402

import pytest  # noqa: E402


@pytest.fixture(autouse=True)
def _collect_garbage_in_main_thread():
    """시험이 끝날 때마다 주 쓰레드에서 정리.
    화면 시험이 남긴 tkinter 객체가 다른 시험의 통신 쓰레드에서 정리되면, mainloop 가 없는 시험 환경에서는
    그 쓰레드가 객체마다 약 1초씩 멈춰(main thread is not in main loop) 시간에 민감한 시험이 실패함."""
    yield
    gc.collect()
