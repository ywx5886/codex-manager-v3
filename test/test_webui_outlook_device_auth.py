from __future__ import annotations

import unittest
from unittest.mock import patch

import httpx

from src.webui.server import app


class _MockResponse:
    def __init__(self, status_code: int, data: dict):
        self.status_code = status_code
        self._data = data
        self.text = str(data)

    def json(self):
        return self._data


class _MockAsyncClient:
    def __init__(self, responses: list[_MockResponse]):
        self._responses = responses

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def post(self, url, data=None):
        return self._responses.pop(0)


class OutlookDeviceAuthApiTests(unittest.IsolatedAsyncioTestCase):
    async def test_device_code_success(self):
        mocked = _MockAsyncClient([
            _MockResponse(200, {
                "user_code": "ABCDEF",
                "device_code": "dev-code",
                "verification_uri": "https://microsoft.com/devicelogin",
                "message": "Use code ABCDEF",
                "interval": 5,
                "expires_in": 900,
            }),
        ])
        transport = httpx.ASGITransport(app=app)
        async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
            with patch("src.webui.server.httpx.AsyncClient", return_value=mocked):
                resp = await client.post("/api/mail/outlook/device-code", json={
                    "client_id": "cid",
                    "tenant_id": "consumers",
                    "scope": "scope",
                })
        self.assertEqual(resp.status_code, 200)
        data = resp.json()
        self.assertEqual(data["user_code"], "ABCDEF")
        self.assertEqual(data["device_code"], "dev-code")

    async def test_device_token_pending(self):
        mocked = _MockAsyncClient([
            _MockResponse(400, {
                "error": "authorization_pending",
                "error_description": "still waiting",
            }),
        ])
        transport = httpx.ASGITransport(app=app)
        async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
            with patch("src.webui.server.httpx.AsyncClient", return_value=mocked):
                resp = await client.post("/api/mail/outlook/device-token", json={
                    "client_id": "cid",
                    "tenant_id": "consumers",
                    "device_code": "dc",
                })
        self.assertEqual(resp.status_code, 200)
        data = resp.json()
        self.assertEqual(data["status"], "pending")
        self.assertEqual(data["error"], "authorization_pending")

    async def test_device_token_success(self):
        mocked = _MockAsyncClient([
            _MockResponse(200, {
                "access_token": "access",
                "refresh_token": "refresh",
                "expires_in": 3600,
                "scope": "scope",
                "token_type": "Bearer",
            }),
        ])
        transport = httpx.ASGITransport(app=app)
        async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
            with patch("src.webui.server.httpx.AsyncClient", return_value=mocked):
                resp = await client.post("/api/mail/outlook/device-token", json={
                    "client_id": "cid",
                    "tenant_id": "consumers",
                    "device_code": "dc",
                })
        self.assertEqual(resp.status_code, 200)
        data = resp.json()
        self.assertEqual(data["status"], "success")
        self.assertEqual(data["refresh_token"], "refresh")


if __name__ == "__main__":
    unittest.main()
