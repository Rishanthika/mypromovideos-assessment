import os
import io
import re
import math
import datetime
import traceback
import requests
import numpy as np

from flask import (
    Flask,
    request,
    jsonify,
    send_from_directory,
    send_file
)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

import matplotlib
matplotlib.use("Agg")

from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
import matplotlib.pyplot as plt

app = Flask(__name__)

# =========================================================
# ENV
# =========================================================
YOUTUBE_API_KEY = os.environ.get("YOUTUBE_API_KEY")

# =========================================================
# COLORS
# =========================================================
DARK_BG = RGBColor(0x0D, 0x1B, 0x2A)
ACCENT1 = RGBColor(0x00, 0xC2, 0xA0)
ACCENT2 = RGBColor(0xFF, 0x6B, 0x35)
ACCENT3 = RGBColor(0xF7, 0xC5, 0x9F)
MID_BG = RGBColor(0x1A, 0x2E, 0x44)
LIGHT_BG = RGBColor(0xF4, 0xF7, 0xFB)
TEXT_DARK = RGBColor(0x0D, 0x1B, 0x2A)
TEXT_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_MID = RGBColor(0x64, 0x74, 0x87)

# =========================================================
# CHART COLORS
# =========================================================
CHART_COLORS = [
    "#00C2A0",
    "#FF6B35",
    "#2563EB",
    "#9333EA",
    "#F59E0B"
]

# =========================================================
# HELPERS
# =========================================================
def format_number(n):

    if n >= 1_000_000:
        return f"{n / 1_000_000:.1f}M"

    if n >= 1_000:
        return f"{n / 1_000:.1f}K"

    return str(n)

# =========================================================
# YOUTUBE API
# =========================================================
def search_channel(company_name):

    try:

        url = "https://www.googleapis.com/youtube/v3/search"

        params = {
            "part": "snippet",
            "q": company_name,
            "type": "channel",
            "maxResults": 1,
            "key": YOUTUBE_API_KEY
        }

        r = requests.get(url, params=params, timeout=10)

        data = r.json()

        items = data.get("items", [])

        if not items:
            return None

        return items[0]["snippet"]["channelId"]

    except:
        return None


def get_channel_stats(channel_id):

    try:

        url = "https://www.googleapis.com/youtube/v3/channels"

        params = {
            "part": "statistics,snippet,contentDetails",
            "id": channel_id,
            "key": YOUTUBE_API_KEY
        }

        r = requests.get(url, params=params, timeout=10)

        data = r.json()

        items = data.get("items", [])

        if not items:
            return {}

        item = items[0]

        stats = item.get("statistics", {})
        snippet = item.get("snippet", {})
        content = item.get("contentDetails", {})

        return {
            "channel_title": snippet.get("title", ""),
            "subscribers": int(stats.get("subscriberCount", 0)),
            "total_videos": int(stats.get("videoCount", 0)),
            "total_views": int(stats.get("viewCount", 0)),
            "uploads_playlist": content.get(
                "relatedPlaylists",
                {}
            ).get("uploads", "")
        }

    except:
        return {}


def get_recent_videos(playlist_id):

    try:

        url = "https://www.googleapis.com/youtube/v3/playlistItems"

        params = {
            "part": "snippet,contentDetails",
            "playlistId": playlist_id,
            "maxResults": 25,
            "key": YOUTUBE_API_KEY
        }

        r = requests.get(url, params=params, timeout=10)

        data = r.json()

        items = data.get("items", [])

        return [
            x["contentDetails"]["videoId"]
            for x in items
        ]

    except:
        return []


def get_video_stats(video_ids):

    videos = []

    if not video_ids:
        return videos

    try:

        url = "https://www.googleapis.com/youtube/v3/videos"

        params = {
            "part": "statistics,snippet",
            "id": ",".join(video_ids),
            "key": YOUTUBE_API_KEY
        }

        r = requests.get(url, params=params, timeout=10)

        data = r.json()

        items = data.get("items", [])

        for item in items:

            stats = item.get("statistics", {})
            snippet = item.get("snippet", {})

            videos.append({
                "title": snippet.get("title", ""),
                "views": int(stats.get("viewCount", 0)),
                "likes": int(stats.get("likeCount", 0)),
                "comments": int(stats.get("commentCount", 0)),
                "published_at": snippet.get("publishedAt", ""),
                "tags": snippet.get("tags", [])[:10]
            })

    except:
        pass

    return videos

# =========================================================
# MOCK DATA
# =========================================================
def generate_mock_data(company_name):

    return {
        "company": company_name,

        "channel": {
            "channel_title": company_name,
            "subscribers": 175000,
            "total_videos": 360,
            "total_views": 24000000
        },

        "videos": [
            {
                "title": f"{company_name} Marketing Strategy",
                "views": 300000,
                "likes": 12000,
                "comments": 500,
                "published_at": datetime.datetime.now().isoformat(),
                "tags": ["marketing", "strategy", "business"]
            }
        ]
    }

# =========================================================
# FETCH DATA
# =========================================================
def fetch_company_data(company_name):

    if not YOUTUBE_API_KEY:
        return generate_mock_data(company_name)

    channel_id = search_channel(company_name)

    if not channel_id:
        return generate_mock_data(company_name)

    channel = get_channel_stats(channel_id)

    playlist_id = channel.get("uploads_playlist", "")

    video_ids = get_recent_videos(playlist_id)

    videos = get_video_stats(video_ids)

    return {
        "company": company_name,
        "channel": channel,
        "videos": videos
    }

# =========================================================
# ANALYSIS
# =========================================================
def analyze_data(companies_data):

    analyzed = []

    for data in companies_data:

        videos = data["videos"]

        total_views = sum(v["views"] for v in videos)
        total_likes = sum(v["likes"] for v in videos)
        total_comments = sum(v["comments"] for v in videos)

        count = len(videos) or 1

        avg_views = total_views // count

        engagement_rate = (
            (total_likes + total_comments)
            / max(total_views, 1)
        ) * 100

        # =============================
        # MONTH COUNTS FIXED
        # =============================
        month_counts = {}

        for v in videos:

            try:

                dt = datetime.datetime.fromisoformat(
                    v["published_at"].replace("Z", "+00:00")
                )

                key = dt.strftime("%Y-%m")

                month_counts[key] = (
                    month_counts.get(key, 0) + 1
                )

            except:
                pass

        # =============================
        # TOP THEMES
        # =============================
        all_tags = []

        for v in videos:
            all_tags.extend(v.get("tags", []))

        top_themes = list(set(all_tags))[:5]

        top_videos = sorted(
            videos,
            key=lambda x: x["views"],
            reverse=True
        )[:5]

        analyzed.append({

            "company": data["company"],

            "channel": data["channel"],

            "avg_views": avg_views,

            "engagement_rate": round(
                engagement_rate,
                2
            ),

            "month_counts": month_counts,

            "top_themes": top_themes,

            "top_videos": top_videos
        })

    return analyzed


def compute_scores(analyzed):

    max_views = max(
        [a["avg_views"] for a in analyzed]
    ) or 1

    for a in analyzed:

        score = (
            a["avg_views"] / max_views
        ) * 100

        a["score"] = round(score)

    return sorted(
        analyzed,
        key=lambda x: x["score"],
        reverse=True
    )

# =========================================================
# CHARTS
# =========================================================
def fig_to_png_bytes(fig):

    buf = io.BytesIO()

    fig.savefig(
        buf,
        format="png",
        dpi=150,
        bbox_inches="tight"
    )

    buf.seek(0)

    return buf


def make_bar_chart(labels, values, title):

    fig = Figure(figsize=(6, 3.5))

    canvas = FigureCanvasAgg(fig)

    ax = fig.add_subplot(111)

    ax.bar(
        labels,
        values,
        color=CHART_COLORS[:len(labels)]
    )

    ax.set_title(title)

    fig.tight_layout()

    return fig


def add_chart_image(slide, fig, x, y, w, h):

    buf = fig_to_png_bytes(fig)

    slide.shapes.add_picture(
        buf,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h)
    )

    fig.clf()

    plt.close("all")

# =========================================================
# PPT HELPERS
# =========================================================
def add_dark_slide(prs, title):

    slide = prs.slides.add_slide(
        prs.slide_layouts[6]
    )

    slide.background.fill.solid()

    slide.background.fill.fore_color.rgb = DARK_BG

    title_box = slide.shapes.add_textbox(
        Inches(0.5),
        Inches(0.3),
        Inches(8),
        Inches(0.5)
    )

    p = title_box.text_frame.paragraphs[0]

    p.text = title

    p.font.size = Pt(28)

    p.font.bold = True

    p.font.color.rgb = TEXT_LIGHT

    return slide


def add_text_box(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=12,
    bold=False,
    color=None
):

    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h)
    )

    p = box.text_frame.paragraphs[0]

    p.text = str(text)

    p.font.size = Pt(size)

    p.font.bold = bold

    p.font.color.rgb = color or TEXT_LIGHT

# =========================================================
# PPT GENERATOR
# =========================================================
def build_pptx(analyzed, your_company):

    prs = Presentation()

    prs.slide_width = Inches(10)

    prs.slide_height = Inches(5.625)

    # =====================================================
    # TITLE
    # =====================================================
    slide = add_dark_slide(
        prs,
        "VIDEO COMPETITOR INTELLIGENCE"
    )

    add_text_box(
        slide,
        f"Generated for {your_company}",
        0.5,
        1.2,
        6,
        0.5,
        16
    )

    # =====================================================
    # PERFORMANCE
    # =====================================================
    slide = add_dark_slide(
        prs,
        "CHANNEL PERFORMANCE"
    )

    companies = [
        a["company"]
        for a in analyzed
    ]

    views = [
        a["avg_views"]
        for a in analyzed
    ]

    fig = make_bar_chart(
        companies,
        views,
        "Average Views"
    )

    add_chart_image(
        slide,
        fig,
        0.5,
        1.1,
        8,
        3.5
    )

    # =====================================================
    # FINAL SCORES
    # =====================================================
    slide = add_dark_slide(
        prs,
        "FINAL SCORES"
    )

    y = 1.2

    for idx, a in enumerate(analyzed):

        add_text_box(
            slide,
            f"{idx + 1}. {a['company']} — {a['score']}/100",
            0.8,
            y,
            5,
            0.4,
            18,
            True
        )

        y += 0.5

    # =====================================================
    # SAVE
    # =====================================================
    buf = io.BytesIO()

    prs.save(buf)

    buf.seek(0)

    return buf

# =========================================================
# ROUTES
# =========================================================
@app.route("/")
def index():

    return send_from_directory(
        "static",
        "index.html"
    )

@app.route("/api/analyze", methods=["POST"])
def analyze():

    try:

        body = request.get_json()

        your_company = body.get(
            "your_company",
            ""
        ).strip()

        competitors = [
            c.strip()
            for c in body.get(
                "competitors",
                []
            )
            if c.strip()
        ]

        if not your_company:

            return jsonify({
                "error": "Company name required"
            }), 400

        companies = [your_company] + competitors[:4]

        companies_data = [
            fetch_company_data(c)
            for c in companies
        ]

        analyzed = compute_scores(
            analyze_data(companies_data)
        )

        safe_result = []

        for a in analyzed:

            safe_result.append({

                "company": a["company"],

                "score": a.get(
                    "score",
                    0
                ),

                "channel": {

                    "subscribers":
                        a["channel"].get(
                            "subscribers",
                            0
                        ),

                    "total_videos":
                        a["channel"].get(
                            "total_videos",
                            0
                        ),

                    "total_views":
                        a["channel"].get(
                            "total_views",
                            0
                        )
                },

                "avg_views":
                    a.get(
                        "avg_views",
                        0
                    ),

                "engagement_rate":
                    a.get(
                        "engagement_rate",
                        0
                    ),

                "top_themes":
                    a.get(
                        "top_themes",
                        []
                    ),

                "top_videos":
                    a.get(
                        "top_videos",
                        []
                    )
            })

        return jsonify({
            "companies": safe_result
        })

    except Exception as e:

        print(traceback.format_exc())

        return jsonify({
            "error": str(e)
        }), 500


@app.route("/api/download", methods=["POST"])
def download():

    try:

        body = request.get_json()

        your_company = body.get(
            "your_company",
            ""
        ).strip()

        competitors = [
            c.strip()
            for c in body.get(
                "competitors",
                []
            )
            if c.strip()
        ]

        companies = [your_company] + competitors[:4]

        companies_data = [
            fetch_company_data(c)
            for c in companies
        ]

        analyzed = compute_scores(
            analyze_data(companies_data)
        )

        pptx_buf = build_pptx(
            analyzed,
            your_company
        )

        safe_company = re.sub(
            r"[^A-Za-z0-9_]",
            "",
            your_company.replace(" ", "_")
        )

        filename = f"{safe_company}_report.pptx"

        pptx_buf.seek(0)

        return send_file(
            pptx_buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:

        print(traceback.format_exc())

        return jsonify({
            "error": str(e)
        }), 500

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":

    port = int(
        os.environ.get("PORT", 5000)
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
