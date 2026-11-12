"""Publishing to Instagram and a Facebook Page through Meta's Graph API, and
reading the posts' stats. Runs on the BEAST: the videos / photos come from
Drive, the page token from SnappySjaak right before it is needed.

Instagram
  video     -> Reel: container with upload_type=resumable, the file sent to
               rupload.facebook.com (no public URL needed), wait, publish
  photo(s)  -> Instagram fetches photos from a URL: each photo is first
               uploaded to the Facebook Page as an UNPUBLISHED photo, its
               fbcdn link is given to Instagram, and removed afterwards
               (carousel for 2-10 photos)
Facebook Page
  9:16 video -> Reel (video_reels, resumable upload)
  other video -> page video
  photo(s)  -> page photo / one post with several photos
"""
from __future__ import annotations

import time
from pathlib import Path

import requests

TIMEOUT = 120


class MetaError(RuntimeError):
    pass


def _check(response: requests.Response) -> dict:
    try:
        payload = response.json()
    except ValueError:
        payload = {}
    if not response.ok or (isinstance(payload, dict) and payload.get("error")):
        error = payload.get("error") if isinstance(payload, dict) else None
        message = (error or {}).get("error_user_msg") or (error or {}).get("message") or response.text[:300]
        raise MetaError(f"Meta {response.status_code}: {message}")
    return payload


class Graph:
    def __init__(self, token: dict):
        self.token = token["access_token"]
        self.version = token.get("graph_version") or "v23.0"
        self.base = f"https://graph.facebook.com/{self.version}"

    def get(self, path: str, **params) -> dict:
        return _check(requests.get(f"{self.base}/{path}", params={**params, "access_token": self.token}, timeout=TIMEOUT))

    def post(self, path: str, files=None, base=None, **data) -> dict:
        return _check(requests.post(f"{base or self.base}/{path}", data={**data, "access_token": self.token}, files=files, timeout=600))

    def delete(self, path: str) -> None:
        try:
            requests.delete(f"{self.base}/{path}", params={"access_token": self.token}, timeout=TIMEOUT)
        except Exception:
            pass

    def rupload(self, url: str, path: Path) -> dict:
        size = path.stat().st_size
        with path.open("rb") as handle:
            response = requests.post(url, data=handle, timeout=1800, headers={
                "Authorization": f"OAuth {self.token}", "offset": "0", "file_size": str(size), "Content-Type": "application/octet-stream",
            })
        return _check(response)


def _ig_wait(graph: Graph, container_id: str, minutes: int = 15) -> None:
    deadline = time.time() + minutes * 60
    while time.time() < deadline:
        status = graph.get(container_id, fields="status_code,status")
        code = status.get("status_code")
        if code == "FINISHED":
            return
        if code in ("ERROR", "EXPIRED"):
            raise MetaError(f"Instagram could not process the media: {status.get('status') or code}")
        time.sleep(6)
    raise MetaError("Instagram took too long to process the media")


def _unpublished_page_photo(graph: Graph, page_id: str, path: Path) -> tuple[str, str]:
    """Upload to the Page without publishing; return (photo id, public url)."""
    with path.open("rb") as handle:
        photo = graph.post(f"{page_id}/photos", files={"source": (path.name, handle, "image/jpeg")}, published="false", temporary="true")
    photo_id = photo["id"]
    images = graph.get(photo_id, fields="images").get("images") or []
    if not images:
        raise MetaError("Facebook gave no link for the photo")
    best = max(images, key=lambda image: image.get("width", 0))
    return photo_id, best["source"]


def publish_instagram(graph: Graph, account: dict, files: list[Path], kind: str, caption: str) -> dict:
    ig = account["account_id"]
    caption = caption[:2200]
    if kind == "video":
        container = graph.post(f"{ig}/media", media_type="REELS", upload_type="resumable", caption=caption, share_to_feed="true", thumb_offset="1200")
        upload_uri = container.get("uri") or f"https://rupload.facebook.com/ig-api-upload/{graph.version}/{container['id']}"
        graph.rupload(upload_uri, files[0])
        _ig_wait(graph, container["id"])
        creation = container["id"]
        temp_photos = []
    else:
        temp_photos = []
        urls = []
        for path in files[:10]:
            photo_id, url = _unpublished_page_photo(graph, account["page_id"], path)
            temp_photos.append(photo_id)
            urls.append(url)
        if len(urls) == 1:
            container = graph.post(f"{ig}/media", image_url=urls[0], caption=caption)
            _ig_wait(graph, container["id"], minutes=5)
        else:
            children = []
            for url in urls:
                child = graph.post(f"{ig}/media", image_url=url, is_carousel_item="true")
                _ig_wait(graph, child["id"], minutes=5)
                children.append(child["id"])
            container = graph.post(f"{ig}/media", media_type="CAROUSEL", children=",".join(children), caption=caption)
            _ig_wait(graph, container["id"], minutes=5)
        creation = container["id"]
    try:
        published = graph.post(f"{ig}/media_publish", creation_id=creation)
    finally:
        for photo_id in temp_photos:
            graph.delete(photo_id)
    media_id = published["id"]
    permalink = ""
    try:
        permalink = graph.get(media_id, fields="permalink").get("permalink", "")
    except MetaError:
        pass
    return {"platform_post_id": media_id, "post_url": permalink}


def publish_facebook(graph: Graph, account: dict, files: list[Path], kind: str, caption: str, variant: str) -> dict:
    page = account["page_id"] or account["account_id"]
    if kind == "video" and variant == "9x16":
        start = graph.post(f"{page}/video_reels", upload_phase="start")
        video_id = start["video_id"]
        graph.rupload(start.get("upload_url") or f"https://rupload.facebook.com/video-upload/{graph.version}/{video_id}", files[0])
        graph.post(f"{page}/video_reels", upload_phase="finish", video_id=video_id, video_state="PUBLISHED", description=caption)
        return {"platform_post_id": f"video:{video_id}", "post_url": f"https://www.facebook.com/reel/{video_id}"}
    if kind == "video":
        with files[0].open("rb") as handle:
            video = graph.post(f"{page}/videos", files={"source": (files[0].name, handle, "video/mp4")}, description=caption,
                               base=f"https://graph-video.facebook.com/{graph.version}")
        return {"platform_post_id": f"video:{video['id']}", "post_url": f"https://www.facebook.com/{page}/videos/{video['id']}"}
    if len(files) == 1:
        with files[0].open("rb") as handle:
            photo = graph.post(f"{page}/photos", files={"source": (files[0].name, handle, "image/jpeg")}, message=caption)
        post_id = photo.get("post_id") or photo["id"]
        return {"platform_post_id": f"post:{post_id}", "post_url": f"https://www.facebook.com/{post_id}"}
    attached = {}
    for index, path in enumerate(files[:30]):
        with path.open("rb") as handle:
            photo = graph.post(f"{page}/photos", files={"source": (path.name, handle, "image/jpeg")}, published="false")
        attached[f"attached_media[{index}]"] = f'{{"media_fbid":"{photo["id"]}"}}'
    post = graph.post(f"{page}/feed", message=caption, **attached)
    return {"platform_post_id": f"post:{post['id']}", "post_url": f"https://www.facebook.com/{post['id']}"}


def publish(token: dict, payload: dict, files: list[Path]) -> dict:
    graph = Graph(token)
    account = {"account_id": token["account_id"], "page_id": token.get("page_id") or payload.get("page_id") or ""}
    media = payload["media"]
    caption = payload.get("caption") or ""
    if payload["platform"] == "instagram":
        return publish_instagram(graph, account, files, media["kind"], caption)
    if payload["platform"] == "facebook":
        return publish_facebook(graph, account, files, media["kind"], caption, (media["files"][0] or {}).get("variant", ""))
    raise MetaError(f"Not a Meta platform: {payload['platform']}")


# ---------------------------------------------------------------------------
# stats (best effort: metrics differ per media type and API version)

def _insights(graph: Graph, object_id: str, metrics: list[str], edge: str = "insights") -> dict:
    values = {}
    for attempt in (metrics, *[[metric] for metric in metrics]):
        try:
            data = graph.get(f"{object_id}/{edge}", metric=",".join(attempt)).get("data") or []
            for row in data:
                value = (row.get("values") or [{}])[0].get("value", row.get("total_value", {}).get("value"))
                if isinstance(value, (int, float)):
                    values[row["name"]] = value
            if len(attempt) == len(metrics):
                return values
        except MetaError:
            continue
    return values


def stats_for(token: dict, post: dict) -> dict:
    graph = Graph(token)
    pid = post["platform_post_id"]
    out = {"post_id": post["post_id"]}
    if post["platform"] == "instagram":
        found = _insights(graph, pid, ["views", "reach", "likes", "comments", "shares", "saved", "follows"])
        out.update({"views": found.get("views"), "reach": found.get("reach"), "likes": found.get("likes"), "comments": found.get("comments"),
                    "shares": found.get("shares"), "saves": found.get("saved"), "follows": found.get("follows")})
        if out["likes"] is None:
            try:
                basic = graph.get(pid, fields="like_count,comments_count")
                out["likes"], out["comments"] = basic.get("like_count"), basic.get("comments_count")
            except MetaError:
                pass
        return out
    kind, _, object_id = pid.partition(":")
    if kind == "video":
        found = _insights(graph, object_id, ["blue_reels_play_count", "total_video_views", "post_impressions_unique"], edge="video_insights")
        out["views"] = found.get("blue_reels_play_count") or found.get("total_video_views")
        out["reach"] = found.get("post_impressions_unique")
    else:
        found = _insights(graph, object_id, ["post_impressions_unique", "post_clicks"])
        out["reach"], out["clicks"] = found.get("post_impressions_unique"), found.get("post_clicks")
    try:
        basic = graph.get(object_id, fields="likes.summary(true).limit(0),comments.summary(true).limit(0)" + (",shares" if kind == "post" else ""))
        out["likes"] = (basic.get("likes") or {}).get("summary", {}).get("total_count")
        out["comments"] = (basic.get("comments") or {}).get("summary", {}).get("total_count")
        if kind == "post":
            out["shares"] = (basic.get("shares") or {}).get("count", 0)
    except MetaError:
        pass
    return out
