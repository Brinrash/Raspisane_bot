"""Small asynchronous client for the official MAX Bot API."""
from __future__ import annotations

import asyncio
import json
import logging
import ssl
from typing import Any
from urllib.parse import urlencode

import aiohttp

logger = logging.getLogger("schedule-bot.max-api")
API_BASE = "https://platform-api2.max.ru"


class MaxApiError(RuntimeError):
    pass


class MaxApi:
    def __init__(self, token: str, ca_bundle: str = "") -> None:
        if not token:
            raise ValueError("MAX_BOT_TOKEN не задан")
        self.token = token
        self.ca_bundle = ca_bundle
        self.session: aiohttp.ClientSession | None = None

    async def __aenter__(self) -> "MaxApi":
        ssl_context = ssl.create_default_context(cafile=self.ca_bundle or None)
        connector = aiohttp.TCPConnector(ssl=ssl_context)
        timeout = aiohttp.ClientTimeout(total=100, connect=20)
        self.session = aiohttp.ClientSession(
            headers={"Authorization": self.token},
            connector=connector,
            timeout=timeout,
        )
        return self

    async def __aexit__(self, exc_type, exc, tb) -> None:
        if self.session:
            await self.session.close()
            self.session = None

    async def request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        payload: dict[str, Any] | None = None,
        retries: int = 3,
    ) -> dict[str, Any]:
        if self.session is None:
            raise RuntimeError("MAX API session is not open")
        url = f"{API_BASE}{path}"
        for attempt in range(retries + 1):
            try:
                async with self.session.request(method, url, params=params, json=payload) as response:
                    body_text = await response.text()
                    if response.status == 429 and attempt < retries:
                        try:
                            body = json.loads(body_text) if body_text else {}
                        except json.JSONDecodeError:
                            body = {}
                        retry_after = response.headers.get("Retry-After")
                        delay = float(retry_after) if retry_after else min(2 ** attempt, 8)
                        logger.warning("MAX rate limit; retrying in %.1fs", delay)
                        await asyncio.sleep(delay)
                        continue
                    if response.status >= 400:
                        # Avoid logging request headers: they contain the access token.
                        raise MaxApiError(f"MAX API вернул HTTP {response.status}: {body_text[:700]}")
                    if not body_text:
                        return {}
                    try:
                        decoded = json.loads(body_text)
                    except json.JSONDecodeError as exc:
                        raise MaxApiError("MAX API вернул некорректный JSON") from exc
                    if isinstance(decoded, dict) and decoded.get("success") is False:
                        raise MaxApiError(str(decoded.get("message") or "MAX API сообщил об ошибке"))
                    return decoded if isinstance(decoded, dict) else {"result": decoded}
            except (aiohttp.ClientError, asyncio.TimeoutError) as exc:
                if attempt >= retries:
                    raise MaxApiError(f"Ошибка соединения с MAX API: {exc.__class__.__name__}") from exc
                await asyncio.sleep(min(2 ** attempt, 8))
        raise MaxApiError("Не удалось выполнить запрос к MAX API")

    async def get_me(self) -> dict[str, Any]:
        return await self.request("GET", "/me")

    async def get_updates(self, marker: int | None, timeout: int) -> dict[str, Any]:
        params: dict[str, Any] = {"limit": 100, "timeout": timeout}
        if marker is not None:
            params["marker"] = marker
        return await self.request("GET", "/updates", params=params, retries=0)

    @staticmethod
    def message_body(text: str, buttons: list[list[dict[str, str]]] | None = None) -> dict[str, Any]:
        body: dict[str, Any] = {"text": text[:4000]}
        if buttons:
            body["attachments"] = [{"type": "inline_keyboard", "payload": {"buttons": buttons}}]
        return body

    async def send_user(
        self,
        user_id: int,
        text: str,
        buttons: list[list[dict[str, str]]] | None = None,
    ) -> dict[str, Any]:
        return await self.request(
            "POST",
            "/messages",
            params={"user_id": int(user_id)},
            payload=self.message_body(text, buttons),
        )

    async def send_chat(
        self,
        chat_id: int,
        text: str,
        buttons: list[list[dict[str, str]]] | None = None,
    ) -> dict[str, Any]:
        return await self.request(
            "POST",
            "/messages",
            params={"chat_id": int(chat_id)},
            payload=self.message_body(text, buttons),
        )

    async def answer_callback(
        self,
        callback_id: str,
        text: str,
        buttons: list[list[dict[str, str]]] | None = None,
    ) -> dict[str, Any]:
        body = {"message": self.message_body(text, buttons)}
        if not buttons:
            body["message"]["attachments"] = []
        return await self.request(
            "POST",
            "/answers",
            params={"callback_id": callback_id},
            payload=body,
        )
