from __future__ import annotations

import os
import re
import warnings
from typing import Any
from urllib.parse import quote, unquote, urljoin, urlparse, urlunparse

import requests
from requests.auth import HTTPBasicAuth
from urllib3.exceptions import InsecureRequestWarning


class JenkinsError(Exception):
    """Raised when Jenkins API access fails."""


def _resolve_verify() -> bool | str:
    """Return requests verify= value from env.

    JENKINS_CA_BUNDLE / REQUESTS_CA_BUNDLE → path to CA cert file
    JENKINS_VERIFY_SSL=false → disable verification (corporate/internal Jenkins)
    """
    ca = (os.getenv("JENKINS_CA_BUNDLE") or os.getenv("REQUESTS_CA_BUNDLE") or "").strip()
    if ca:
        return ca
    flag = (os.getenv("JENKINS_VERIFY_SSL") or "true").strip().lower()
    if flag in {"0", "false", "no", "off"}:
        return False
    return True


class JenkinsClient:
    def __init__(
        self,
        username: str | None = None,
        token: str | None = None,
        timeout: float = 60.0,
        verify: bool | str | None = None,
    ) -> None:
        self.username = username or os.getenv("JENKINS_USER", "")
        self.token = (
            token
            or os.getenv("JENKINS_TOKEN")
            or os.getenv("JENKINS_API_TOKEN")
            or ""
        )
        self.timeout = timeout
        self.verify = _resolve_verify() if verify is None else verify
        self.session = requests.Session()
        self.session.verify = self.verify
        if self.verify is False:
            warnings.filterwarnings("ignore", category=InsecureRequestWarning)
        if self.username and self.token:
            self.session.auth = HTTPBasicAuth(self.username, self.token)

    def get_json(self, url: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        try:
            response = self.session.get(url, params=params, timeout=self.timeout)
        except requests.exceptions.SSLError as exc:
            raise JenkinsError(
                "SSL certificate verification failed talking to Jenkins.\n"
                "For internal Jenkins, add to .env:\n"
                "  JENKINS_VERIFY_SSL=false\n"
                "Or point to your corporate CA with:\n"
                "  JENKINS_CA_BUNDLE=C:\\path\\to\\corp-root-ca.pem\n"
                f"Details: {exc}"
            ) from exc
        except requests.exceptions.RequestException as exc:
            raise JenkinsError(f"Network error talking to Jenkins: {exc}") from exc

        if response.status_code == 401:
            raise JenkinsError(
                "Jenkins authentication failed (401). "
                "Set JENKINS_USER and JENKINS_TOKEN / JENKINS_API_TOKEN."
            )
        if response.status_code == 403:
            raise JenkinsError(
                "Jenkins access forbidden (403). Check permissions for this job."
            )
        if response.status_code == 404:
            raise JenkinsError(f"Jenkins resource not found (404): {url}")
        if not response.ok:
            raise JenkinsError(
                f"Jenkins request failed ({response.status_code}): {url}\n{response.text[:500]}"
            )
        return response.json()

    def get_json_or_none(
        self, url: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any] | None:
        try:
            return self.get_json(url, params=params)
        except JenkinsError as exc:
            if "404" in str(exc):
                return None
            raise

    def get_text(self, url: str) -> str:
        try:
            response = self.session.get(url, timeout=self.timeout)
        except requests.exceptions.RequestException as exc:
            raise JenkinsError(f"Network error talking to Jenkins: {exc}") from exc

        if response.status_code == 401:
            raise JenkinsError("Jenkins authentication failed (401).")
        if response.status_code == 403:
            raise JenkinsError("Jenkins access forbidden (403).")
        if response.status_code == 404:
            raise JenkinsError(f"Jenkins resource not found (404): {url}")
        if not response.ok:
            raise JenkinsError(f"Jenkins request failed ({response.status_code}): {url}")
        return response.text


_BUILD_URL_RE = re.compile(
    r"^(?P<root>https?://[^/]+)/(?P<path>(?:job/[^/]+/)+)(?P<number>\d+)/?(?:.*)?$",
    re.IGNORECASE,
)


def normalize_build_url(raw_url: str) -> str:
    """Normalize a Jenkins build URL to .../job/.../<number>/."""
    url = raw_url.strip()
    parsed = urlparse(url)
    if not parsed.scheme or not parsed.netloc:
        raise JenkinsError(f"Invalid Jenkins URL: {raw_url}")

    # Drop query/fragment and trailing junk after build number when possible.
    match = _BUILD_URL_RE.match(url.split("?", 1)[0].split("#", 1)[0])
    if match:
        root = match.group("root")
        path = match.group("path")
        number = match.group("number")
        return f"{root}/{path}{number}/"

    # Fallback: keep path up to first numeric segment that looks like a build.
    parts = [p for p in parsed.path.split("/") if p]
    build_idx = None
    for i, part in enumerate(parts):
        if part.isdigit() and i > 0 and parts[i - 1] != "job":
            # Prefer segment after job/<name>/
            if i >= 2 and parts[i - 2] == "job":
                build_idx = i
                break
    if build_idx is None:
        raise JenkinsError(
            "Could not parse a build number from the Jenkins URL. "
            "Expected something like https://jenkins/.../job/MyJob/123/"
        )

    build_path = "/".join(parts[: build_idx + 1]) + "/"
    return urlunparse((parsed.scheme, parsed.netloc, build_path, "", "", ""))


def job_slug_from_build_url(build_url: str) -> str:
    """Create a filesystem-safe slug from job path + build number."""
    normalized = normalize_build_url(build_url)
    match = _BUILD_URL_RE.match(normalized.rstrip("/") + "/")
    if not match:
        return "build"
    path = match.group("path")  # job/A/job/B/
    number = match.group("number")
    names = []
    segments = [s for s in path.split("/") if s]
    for i, seg in enumerate(segments):
        if seg == "job" and i + 1 < len(segments):
            names.append(unquote(segments[i + 1]))
    job_part = "_".join(names) if names else "job"
    safe = re.sub(r"[^\w.\-]+", "_", job_part).strip("_")
    return f"{safe}_{number}"


def api_json_url(resource_url: str) -> str:
    base = resource_url if resource_url.endswith("/") else resource_url + "/"
    return urljoin(base, "api/json")


def test_report_url(build_url: str) -> str:
    base = build_url if build_url.endswith("/") else build_url + "/"
    return urljoin(base, "testReport/api/json")


def console_text_url(build_url: str) -> str:
    base = build_url if build_url.endswith("/") else build_url + "/"
    return urljoin(base, "consoleText")


def build_test_case_url(
    build_url: str, suite_name: str, class_name: str, test_name: str
) -> str:
    """Best-effort Jenkins test case deep link."""
    base = build_url if build_url.endswith("/") else build_url + "/"
    suite = quote(suite_name or "(root)", safe="")
    # Jenkins often uses package/class URL encoding under testReport
    class_path = quote(class_name or suite_name or "(root)", safe="")
    name = quote(test_name, safe="")
    return urljoin(base, f"testReport/{suite}/{class_path}/{name}/")
