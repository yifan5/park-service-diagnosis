#!/usr/bin/env python3
"""Start the local park analysis demo server."""

import os

from backend.app import serve


if __name__ == "__main__":
    serve(os.environ.get("PARK_HOST", "127.0.0.1"), int(os.environ.get("PARK_PORT", "8765")))
