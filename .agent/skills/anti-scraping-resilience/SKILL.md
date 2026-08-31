---
name: anti-scraping-resilience
description: Implement robust evasion, adaptive rate limiting, realistic browser fingerprinting (User-Agents, Sec-Ch-Ua headers), and resilient backoff handling for strict e-commerce platforms like Rakuten and Yahoo Japan.
---

# Anti-Scraping Resilience Skill

This skill provides patterns and techniques to prevent IP rate-limiting, CAPTCHAs, and 403/429 HTTP blocks when scraping high-traffic e-commerce marketplaces.

## 🎯 When to Use This Skill

- When scrapers encounter `429 Too Many Requests`, `403 Forbidden`, or slow response delays (>10 seconds per page).
- When configuring realistic browser headers for Japanese e-commerce platforms.
- Implementing exponential backoff with randomized jitter.
- Integrating session pooling, proxy rotation, or TLS fingerprinting.

---

## 🛡️ Core Defense Strategies

### 1. Modern Client Hints & Headers Pool

Never send bare `User-Agent` strings. Modern Chromium browsers send Client Hints (`Sec-Ch-Ua` headers).

```python
"""Realistic Japanese Browser Header Pool."""

import random
from typing import Dict

BROWSER_PROFILES = [
    {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/125.0.0.0 Safari/537.36"
        ),
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;q=0.9,"
            "image/avif,image/webp,image/apng,*/*;q=0.8"
        ),
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Sec-Ch-Ua": (
            '"Google Chrome";v="125", "Chromium";v="125", "Not.A/Brand";v="24"'
        ),
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"Windows"',
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "none",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
    },
    {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7",
        "Sec-Ch-Ua-Platform": '"macOS"',
    },
]


def get_random_headers() -> Dict[str, str]:
    """Selects a randomized browser header profile."""
    return random.choice(BROWSER_PROFILES).copy()
```

---

### 2. Adaptive Exponential Backoff with Jitter

When a request encounters connection errors, timeouts, or HTTP 429:

```python
import random
import time
import requests
from typing import Optional


def fetch_with_backoff(
    url: str,
    headers: Dict[str, str],
    max_retries: int = 4,
    base_delay: float = 2.0,
    timeout: int = 15,
) -> Optional[requests.Response]:
    """Fetches a URL with randomized exponential backoff on failure."""
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=headers, timeout=timeout)
            if response.status_code == 200:
                return response
            elif response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", "10"))
                delay = retry_after + random.uniform(1.0, 3.0)
                print(f"Rate limited (429). Backing off for {delay:.1f}s...")
                time.sleep(delay)
            elif response.status_code in [500, 502, 503, 504]:
                delay = base_delay * (2**attempt) + random.uniform(0.5, 1.5)
                time.sleep(delay)
            else:
                return response
        except (requests.Timeout, requests.ConnectionError) as exc:
            delay = base_delay * (2**attempt) + random.uniform(0.5, 1.5)
            print(f"Connection issue ({exc}). Retrying in {delay:.1f}s...")
            time.sleep(delay)

    return None
```

---

### 3. Polite Rate Limiting & Sleep Jitter

- Never use a static `time.sleep(1.0)`.
- Use dynamic sleep intervals: `time.sleep(COURTESY_PAUSE_SECONDS + random.uniform(0.2, 0.8))`.
