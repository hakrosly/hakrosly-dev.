"""外部通信を禁止する。

このツールは画像の読み込みから出力まですべてローカルで完結する。
万一ライブラリなどが通信しようとしても、ここで止めて例外にする。
"""

from __future__ import annotations

import socket


class NetworkBlockedError(RuntimeError):
    pass


def _refuse(*_args, **_kwargs):
    raise NetworkBlockedError("外部通信は禁止されています (tableocr はローカルでのみ動作します)")


_blocked = False


def block_network() -> None:
    """このプロセスからのネットワーク接続と名前解決をすべて禁止する。"""
    global _blocked
    if _blocked:
        return
    socket.socket.connect = _refuse
    socket.socket.connect_ex = _refuse
    socket.socket.sendto = _refuse
    socket.create_connection = _refuse
    socket.getaddrinfo = _refuse
    socket.gethostbyname = _refuse
    socket.gethostbyname_ex = _refuse
    _blocked = True
