"""Secure, size-limited download of file attachments delivered by MAX."""
from __future__ import annotations

from pathlib import Path
from urllib.parse import urlparse
import aiohttp


async def download_max_file(url: str, destination: Path, max_bytes: int) -> None:
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.hostname:
        raise ValueError("MAX не вернул корректную HTTPS-ссылку на файл")
    timeout = aiohttp.ClientTimeout(total=90, connect=20)
    # Do not forward the bot Authorization token to a CDN or redirect host.
    async with aiohttp.ClientSession(timeout=timeout) as session:
        async with session.get(url, allow_redirects=True) as response:
            response.raise_for_status()
            length = response.headers.get("Content-Length")
            if length and int(length) > max_bytes:
                raise ValueError(f"Файл больше лимита {max_bytes // (1024 * 1024)} МБ")
            total = 0
            with destination.open("wb") as handle:
                async for chunk in response.content.iter_chunked(64 * 1024):
                    total += len(chunk)
                    if total > max_bytes:
                        raise ValueError(f"Файл больше лимита {max_bytes // (1024 * 1024)} МБ")
                    handle.write(chunk)
