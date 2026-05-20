import os
import json
import io
import re
import math
import datetime
import requests
import traceback
from flask import Flask, request, jsonify, send_file, send_from_directory, make_response
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import matplotlib
matplotlib.use('Agg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
import numpy as np

app = Flask(__name__)

YOUTUBE_API_KEY = os.environ.get("YOUTUBE_API_KEY", "AIzaSyC5rQoIscl-rBWSDyFXXP03rRaOEnFs-M8")

# ─── COLOR PALETTE ───────────────────────────────────────────────────────────
DARK_BG    = RGBColor(0x0D, 0x1B, 0x2A)  
ACCENT1    = RGBColor(0x00, 0xC2, 0xA0)  
ACCENT2    = RGBColor(0xFF, 0x6B, 0x35)  
ACCENT3    = RGBColor(0xF7, 0xC5, 0x9F)  
MID_BG     = RGBColor(0x1A, 0x2E, 0x44)  
LIGHT_BG   = RGBColor(0xF4, 0xF7, 0xFB)  
TEXT_DARK  = RGBColor(0x0D, 0x1B, 0x2A)
TEXT_LIGHT = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_MID   = RGBColor(0x64, 0x74, 0x87)
GRID_LINE  = RGBColor(0xE2, 0xE8, 0xF0)

def rgb_to_hex(r, g, b):
    return f"#{r:02x}{g:02x}{b:02x}"

CHART_COLORS_MPL = [
    rgb_to_hex(0x00, 0xC2, 0xA0),
    rgb_to_hex(0xFF, 0x6B, 0x35),
    rgb_to_hex(0x25, 0x63, 0xEB),
    rgb_to_hex(0x93, 0x33, 0xEA),
    rgb_to_hex(0xF5, 0x9E, 0x0B),
]

# ─── YOUTUBE API ─────────────────────────────────────────────────────────────
def search_channel(company_name):
    if not YOUTUBE_API_KEY: return None
    try:
        url = "https://www.googleapis.com/youtube/v3/search"
        params = {"part": "snippet", "q": company_name, "type": "channel", "maxResults": 3, "key": YOUTUBE_API_KEY}
        r = requests.get(url, params=params, timeout=10)
        items = r.json().get("items", [])
        if not items: return None
        for item in items:
            title = item["snippet"]["channelTitle"].lower()
            if company_name.lower() in title or title in company_name.lower():
                return item["snippet"]["channelId"]
        return items[0]["snippet"]["channelId"]
    except: return None

def get_channel_stats(channel_id):
    if not YOUTUBE_API_KEY: return {}
    try:
        url = "https://www.googleapis.com/youtube/v3/channels"
        params = {"part": "statistics,snippet,contentDetails", "id": channel_id, "key": YOUTUBE_API_KEY}
        r = requests.get(url, params=params, timeout=10)
        items = r.json().get("items", [])
        if not items: return {}
        item = items[0]
        stats = item.get("statistics", {})
        snippet = item.get("snippet", {})
        uploads_playlist = item.get("contentDetails", {}).get("relatedPlaylists", {}).get("uploads", "")
        return {
            "channel_id": channel_id,
            "channel_title": snippet.get("title", ""),
            "description": snippet.get("description", "")[:500],
            "subscribers": int(stats.get("subscriberCount", 0)),
            "total_videos": int(stats.get("videoCount", 0)),
            "total_views": int(stats.get("viewCount", 0)),
            "uploads_playlist": uploads_playlist,
            "published_at": snippet.get("publishedAt", ""),
            "country": snippet.get("country", "N/A")
        }
    except: return {}

def get_recent_videos(uploads_playlist_id, max_results=50):
    if not YOUTUBE_API_KEY or not uploads_playlist_id: return []
    try:
        url = "https://www.googleapis.com/youtube/v3/playlistItems"
        params = {"part": "snippet,contentDetails", "playlistId": uploads_playlist_id, "maxResults": max_results, "key": YOUTUBE_API_KEY}
        r = requests.get(url, params=params, timeout=10)
        return [item["contentDetails"]["videoId"] for item in r.json().get("items", [])]
    except: return []

def get_video_stats(video_ids):
    if not YOUTUBE_API_KEY or not video_ids: return []
    videos = []
    for i in range(0, len(video_ids), 50):
        batch = video_ids[i:i+50]
        try:
            url = "https://www.googleapis.com/youtube/v3/videos"
            params = {"part": "statistics,snippet,contentDetails", "id": ",".join(batch), "key": YOUTUBE_API_KEY}
            r = requests.get(url, params=params, timeout=10)
            for item in r.json().get("items", []):
                stats = item.get("statistics", {})
                snippet = item.get("snippet", {})
                videos.append({
                    "id": item["id"],
                    "title": snippet.get("title", ""),
                    "published_at": snippet.get("publishedAt", ""),
                    "views": int(stats.get("viewCount", 0)),
                    "likes": int(stats.get("likeCount", 0)),
                    "comments": int(stats.get("commentCount", 0)),
                    "duration": item.get("contentDetails", {}).get("duration", "PT0S"),
                    "tags": snippet.get("tags", [])[:10],
                    "description": snippet.get("description", "")[:300],
                })
        except: pass
    return videos

def fetch_company_data(company_name):
    channel_id = search_channel(company_name)
    if not channel_id: return generate_mock_data(company_name)
    channel = get_channel_stats(channel_id)
    if not channel: return generate_mock_data(company_name)
    video_ids = get_recent_videos(channel.get("uploads_playlist", ""), max_results=50)
    videos = get_video_stats(video_ids)
    return {"company": company_name, "channel": channel, "videos": videos, "is_mock": False}

def generate_mock_data(company_name):
    import random
    rng = random.Random(hash(company_name) % 10000)
    base_subs = rng.randint(50000, 2000000)
    base_views = rng.randint(500000, 50000000)
    num_videos = rng.randint(80, 400)
    topics = ["product demo", "tutorial", "case study", "webinar", "thought leadership", "customer story", "how-to guide", "industry trends", "company culture", "event recap"]
    videos = []
    for i in range(min(50, num_videos)):
        topic = rng.choice(topics)
        days_ago = rng.randint(1, 365)
        pub = (datetime.datetime.now() - datetime.timedelta(days=days_ago)).isoformat() + "Z"
        views = max(100, int(rng.gauss(base_views / num_videos, base_views / (num_videos * 3))))
        likes = int(views * rng.uniform(0.02, 0.04))
        comments = int(likes * rng.uniform(0.05, 0.10))
        videos.append({
            "id": f"mock_{company_name}_{i}",
            "title": f"{company_name}: {topic.title()} #{i+1}",
            "published_at": pub,
            "views": views, "likes": likes, "comments": comments,
            "duration": f"PT{rng.randint(3,25)}M{rng.randint(0,59)}S",
            "tags": rng.sample(topics, rng.randint(2, 5)),
            "description": f"Learn about {topic} from {company_name}.",
        })
    return {
        "company": company_name,
        "channel": {
            "channel_id": f"mock_{company_name}", "channel_title": f"{company_name} Official",
            "description": f"{company_name}'s official channel", "subscribers": base_subs,
            "total_videos": num_videos, "total_views": base_views,
            "uploads_playlist": "", "published_at": "2015-01-01T00:00:00Z", "country": "US"
        },
        "videos": sorted(videos, key=lambda v: v["views"], reverse=True),
        "is_mock": False
    }

# ─── ANALYSIS ────────────────────────────────────────────────────────────────
def analyze_data(companies_data):
    results = []
    for idx, data in enumerate(companies_data):
        videos = data["videos"]
        total_views = sum(v["views"] for v in videos)
        total_likes = sum(v["likes"] for v in videos)
        total_comments = sum(v["comments"] for v in videos)
        n = len(videos) or 1
        
        avg_views = total_views // n
        avg_likes = total_likes // n
        avg_comments = total_comments // n
        engagement_rate = (total_likes + total_comments) / max(total_views, 1) * 100
        
        uploads_per_month = 0
        if videos:
            dates = []
            for v in videos:
                try: dates.append(datetime.datetime.fromisoformat(v["published_at"].replace("Z", "+00:00")))
                except: pass
            if len(dates) >= 2:
                dates.sort(reverse=True)
                span_days = (dates[0] - dates[-1]).days or 1
                uploads_per_month = len(dates) / (span_days / 30)
                
        all_tags = []
        for v in videos:
            all_tags.extend(v.get("tags", []))
            all_tags.extend(re.findall(r'\b[a-zA-Z]{4,}\b', v["title"].lower()))
            
        from collections import Counter
        STOPWORDS = {"with", "your", "this", "that", "from", "have", "will", "what", "when", "how", "the", "and", "for", "you", "are", "can"}
        top_themes = [t for t, _ in Counter(t.lower() for t in all_tags if t.lower() not in STOPWORDS).most_common(8)]
        
        month_counts = {}
        for v in videos:
            try: month_counts[datetime.datetime.fromisoformat(v["published_at"].replace("Z", "+00:00")).strftime("%Y-%m")] = month_counts.get(d.strftime("%Y-%m"), 0) + 1
            except: pass
            
        results.append({
            **data, "color_index": idx, "avg_views": avg_views, "avg_likes": avg_likes, "avg_comments": avg_comments,
            "engagement_rate": round(engagement_rate, 2), "uploads_per_month": round(uploads_per_month, 1),
            "top_videos": sorted(videos, key=lambda v: v["views"], reverse=True)[:5],
            "top_themes": top_themes, "month_counts": month_counts, "total_video_views": total_views
        })
    return results

def compute_scores(analyzed):
    def safe_max(lst): return max(lst) if lst else 1
    subs = [a["channel"].get("subscribers", 0) for a in analyzed]
    views = [a["avg_views"] for a in analyzed]
    eng = [a["engagement_rate"] for a in analyzed]
    freq = [a["uploads_per_month"] for a in analyzed]
    
    max_subs, max_views, max_eng, max_freq = safe_max(subs), safe_max(views), safe_max(eng), safe_max(freq)
    
    for a in analyzed:
        sub_score = (a["channel"].get("subscribers", 0) / max_subs) * 25
        view_score = (a["avg_views"] / max_views) * 30
        eng_score = (a["engagement_rate"] / max_eng) * 25
        freq_score = (a["uploads_per_month"] / max_freq) * 20
        a["score"] = round(sub_score + view_score + eng_score + freq_score)
        a["sub_score"] = round(sub_score / 25 * 100)
        a["view_score"] = round(view_score / 30 * 100)
        a["eng_score"] = round(eng_score / 25 * 100)
        a["freq_score"] = round(freq_score / 20 * 100)
    return sorted(analyzed, key=lambda a: a["score"], reverse=True)

# ─── THREAD-SAFE CHART GENERATORS ────────────────────────────────────────────
def fig_to_png_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf

def make_bar_chart(labels, values, title, ylabel, colors=None, figsize=(8, 4)):
    fig = Figure(figsize=figsize, facecolor='#F4F7FB')
    canvas = FigureCanvasAgg(fig)
    ax = fig.add_subplot(111)
    ax.set_facecolor('#F4F7FB')
    if colors is None: colors = CHART_COLORS_MPL[:len(labels)]
    bars = ax.bar(labels, values, color=colors, edgecolor='white', linewidth=0.5, width=0.5)
    ax.set_title(title, fontsize=13, fontweight='bold', color='#0D1B2A', pad=10)
    ax.set_ylabel(ylabel, fontsize=9, color='#647487')
    ax.tick_params(axis='x', labelsize=8, colors='#0D1B2A')
    ax.tick_params(axis='y', labelsize=8, colors='#647487')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#E2E8F0')
    ax.spines['bottom'].set_color('#E2E8F0')
    ax.yaxis.grid(True, color='#E2E8F0', linewidth=0.5, alpha=0.8)
    ax.set_axisbelow(True)
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() * 1.01, format_number(val), ha='center', va='bottom', fontsize=7.5, color='#0D1B2A', fontweight='bold')
    fig.tight_layout()
    return fig

def make_grouped_bar(companies, metric_labels, values_matrix, title, figsize=(9, 4)):
    x = np.arange(len(metric_labels))
    n = len(companies)
    width = 0.7 / max(n, 1)
    fig = Figure(figsize=figsize, facecolor='#F4F7FB')
    canvas = FigureCanvasAgg(fig)
    ax = fig.add_subplot(111)
    ax.set_facecolor('#F4F7FB')
    for i, (company, values) in enumerate(zip(companies, values_matrix)):
        offset = (i - n/2 + 0.5) * width
        ax.bar(x + offset, values, width * 0.9, label=company, color=CHART_COLORS_MPL[i % len(CHART_COLORS_MPL)], edgecolor='white', linewidth=0.5)
    ax.set_xticks(x)
    ax.set_xticklabels(metric_labels, fontsize=9, color='#0D1B2A')
    ax.set_title(title, fontsize=12, fontweight='bold', color='#0D1B2A', pad=10)
    ax.tick_params(axis='y', labelsize=8, colors='#647487')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#E2E8F0')
    ax.spines['bottom'].set_color('#E2E8F0')
    ax.yaxis.grid(True, color='#E2E8F0', linewidth=0.5, alpha=0.8)
    ax.set_axisbelow(True)
    ax.legend(fontsize=8, loc='upper right')
    fig.tight_layout()
    return fig

def make_line_chart(companies, month_data_list, figsize=(9, 4)):
    fig = Figure(figsize=figsize, facecolor='#F4F7FB')
    canvas = FigureCanvasAgg(fig)
    ax = fig.add_subplot(111)
    ax.set_facecolor('#F4F7FB')
    all_months = sorted(set(m for md in month_data_list for m in md.keys()))[-12:]
    if not all_months: all_months = [datetime.date.today().strftime("%Y-%m")]
    for i, (company, month_data) in enumerate(zip(companies, month_data_list)):
        values = [month_data.get(m, 0) for m in all_months]
        color = CHART_COLORS_MPL[i % len(CHART_COLORS_MPL)]
        ax.plot(range(len(all_months)), values, 'o-', linewidth=2, color=color, label=company, markersize=5)
    ax.set_xticks(range(len(all_months)))
    ax.set_xticklabels([m[-5:] for m in all_months], rotation=45, fontsize=7, color='#647487')
    ax.set_ylabel("Videos uploaded", fontsize=9, color='#647487')
    ax.set_title("Monthly Upload Frequency (last 12 months)", fontsize=12, fontweight='bold', color='#0D1B2A', pad=10)
    ax.tick_params(axis='y', labelsize=8, colors='#647487')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#E2E8F0')
    ax.spines['bottom'].set_color('#E2E8F0')
    ax.yaxis.grid(True, color='#E2E8F0', linewidth=0.5, alpha=0.8)
    ax.set_axisbelow(True)
    ax.legend(fontsize=8, loc='upper right')
    fig.tight_layout()
    return fig

def make_radar_chart(companies, categories, values_matrix, figsize=(6, 5)):
    N = len(categories)
    angles = [n / float(N) * 2 * math.pi for n in range(N)]
    angles += angles[:1]
    fig = Figure(figsize=figsize, facecolor='#F4F7FB')
    canvas = FigureCanvasAgg(fig)
    ax = fig.add_subplot(111, polar=True)
    ax.set_facecolor('#F4F7FB')
    ax.spines['polar'].set_color('#E2E8F0')
    for i, (company, values) in enumerate(zip(companies, values_matrix)):
        vals = list(values) + [values[0]]
        color = CHART_COLORS_MPL[i % len(CHART_COLORS_MPL)]
        ax.plot(angles, vals, 'o-', linewidth=2, color=color, label=company)
        ax.fill(angles, vals, alpha=0.12, color=color)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=8, color='#0D1B2A')
    ax.set_yticklabels([])
    ax.yaxis.grid(True, color='#E2E8F0')
    ax.xaxis.grid(True, color='#E2E8F0')
    ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1), fontsize=8)
    fig.tight_layout()
    return fig

# ─── FORMATTING HELPERS ──────────────────────────────────────────────────────
def format_number(n):
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M"
    if n >= 1_000: return f"{n/1_000:.1f}K"
    return str(n)

# ─── PPTX BUILDER ────────────────────────────────────────────────────────────
def add_dark_slide(prs, title_text, subtitle_text=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = DARK_BG
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.12), Inches(7.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT1
    bar.line.fill.background()
    if title_text:
        tf = slide.shapes.add_textbox(Inches(0.35), Inches(0.25), Inches(9.3), Inches(0.7)).text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size, p.font.bold, p.font.color.rgb, p.font.name = Pt(28), True, TEXT_LIGHT, "Calibri"
    if subtitle_text:
        tf2 = slide.shapes.add_textbox(Inches(0.35), Inches(0.9), Inches(9.3), Inches(0.4)).text_frame
        p2 = tf2.paragraphs[0]
        p2.text = subtitle_text
        p2.font.size, p2.font.color.rgb, p2.font.name = Pt(12), ACCENT1, "Calibri"
    return slide

def add_light_slide(prs, title_text, subtitle_text=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = LIGHT_BG
    header = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.08))
    header.fill.solid()
    header.fill.fore_color.rgb = ACCENT1
    header.line.fill.background()
    if title_text:
        tf = slide.shapes.add_textbox(Inches(0.4), Inches(0.18), Inches(9.2), Inches(0.6)).text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size, p.font.bold, p.font.color.rgb, p.font.name = Pt(24), True, TEXT_DARK, "Calibri"
    if subtitle_text:
        tf2 = slide.shapes.add_textbox(Inches(0.4), Inches(0.75), Inches(9.2), Inches(0.35)).text_frame
        p2 = tf2.paragraphs[0]
        p2.text = subtitle_text
        p2.font.size, p2.font.color.rgb, p2.font.name = Pt(11), TEXT_MID, "Calibri"
    return slide

def add_text_box(slide, text, x, y, w, h, size=12, bold=False, color=None, align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.text = str(text)
    p.alignment = align
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color if color else TEXT_DARK
    p.font.name = "Calibri"
    return txBox

def add_rect(slide, x, y, w, h, fill_rgb):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    shape.line.fill.background()
    return shape

def add_chart_image(slide, fig, x, y, w, h):
    buf = fig_to_png_bytes(fig)
    slide.shapes.add_picture(buf, Inches(x), Inches(y), Inches(w), Inches(h))
    fig.clf() # 🚨 CRITICAL MEMORY FIX FOR RENDER SERVERS 🚨

def build_pptx(analyzed, your_company):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    companies = [a["company"] for a in analyzed]
    today = datetime.date.today().strftime("%B %d, %Y")
    ranked = sorted(analyzed, key=lambda a: a["score"], reverse=True)
    leader = ranked[0]
    
    # SLIDE 1
    slide = add_dark_slide(prs, "")
    circle = slide.shapes.add_shape(9, Inches(7.5), Inches(-1), Inches(4.5), Inches(4.5))
    circle.fill.solid()
    circle.fill.fore_color.rgb = ACCENT1
    circle.line.fill.background()
    add_text_box(slide, "VIDEO COMPETITOR", 0.6, 1.2, 7, 0.65, size=38, bold=True, color=TEXT_LIGHT)
    add_text_box(slide, "INTELLIGENCE REPORT", 0.6, 1.85, 7, 0.65, size=38, bold=True, color=ACCENT1)
    add_text_box(slide, " · ".join(companies), 0.6, 2.75, 8.5, 0.45, size=12, color=RGBColor(0xA0, 0xB8, 0xD0))
    add_text_box(slide, today, 0.6, 3.3, 4, 0.35, size=10, color=RGBColor(0x70, 0x90, 0xB0))
    
    # SLIDE 2
    slide = add_dark_slide(prs, "Executive Summary", "Who is winning in video marketing and why")
    add_rect(slide, 0.35, 1.4, 4.0, 1.5, MID_BG)
    add_text_box(slide, "🏆 MARKET LEADER", 0.5, 1.5, 3.7, 0.35, size=9, bold=True, color=ACCENT1)
    add_text_box(slide, leader["company"], 0.5, 1.82, 3.7, 0.55, size=22, bold=True, color=TEXT_LIGHT)
    add_text_box(slide, f"Score: {leader['score']}/100", 0.5, 2.38, 3.7, 0.3, size=10, color=ACCENT3)
    add_text_box(slide, "Channel Performance Snapshot", 4.6, 1.35, 5.0, 0.35, size=10, bold=True, color=ACCENT1)
    
    y_pos = 1.75
    for a in ranked[:5]:
        f = f"{'★ ' if a['company'] == leader['company'] else '  '}{a['company']}: {format_number(a['channel'].get('subscribers',0))} subs | {format_number(a['avg_views'])} views | {a['engagement_rate']}% eng"
        add_text_box(slide, f, 4.6, y_pos, 5.1, 0.35, size=9, color=TEXT_LIGHT)
        y_pos += 0.38
        
    add_rect(slide, 0.35, 3.15, 9.3, 1.95, MID_BG)
    add_text_box(slide, "KEY STRATEGIC INSIGHT", 0.55, 3.25, 9.0, 0.3, size=9, bold=True, color=ACCENT1)
    insight_text = f"{leader['company']} dominates with the highest composite score of {leader['score']}/100, driven by {format_number(leader['channel'].get('subscribers',0))} subscribers and {leader['engagement_rate']}% engagement rate. "
    add_text_box(slide, insight_text, 0.55, 3.6, 9.0, 1.35, size=10, color=TEXT_LIGHT, wrap=True)
    
    # SLIDE 3
    slide = add_light_slide(prs, "Channel Overview", "Subscribers · Total Videos · Total Views")
    fig1 = make_bar_chart(companies, [a["channel"].get("subscribers", 0) for a in analyzed], "Subscribers", "Count", None, figsize=(4.5, 3.2))
    add_chart_image(slide, fig1, 0.3, 1.2, 4.5, 3.2)
    fig2 = make_bar_chart(companies, [a["channel"].get("total_videos", 0) for a in analyzed], "Total Videos Published", "Count", None, figsize=(4.5, 3.2))
    add_chart_image(slide, fig2, 5.2, 1.2, 4.5, 3.2)
    add_rect(slide, 0.3, 4.6, 9.4, 0.65, RGBColor(0xE8, 0xF4, 0xF1))
    x = 0.5
    for a in analyzed:
        add_text_box(slide, a["company"], x, 4.68, 1.8, 0.25, size=8, bold=True, color=TEXT_DARK)
        add_text_box(slide, f"{format_number(a['channel'].get('total_views',0))} total views", x, 4.92, 1.8, 0.25, size=8, color=TEXT_MID)
        x += 9.4 / len(analyzed)
        
    # SLIDE 4
    slide = add_light_slide(prs, "Content Performance", "Top performing videos by views and engagement")
    card_w = 9.2 / len(analyzed)
    for i, a in enumerate(analyzed):
        x = 0.35 + i * card_w
        add_rect(slide, x, 1.15, card_w - 0.12, 4.1, RGBColor(0xFF, 0xFF, 0xFF))
        hex_color = CHART_COLORS_MPL[a["color_index"] % 5]
        header_color = RGBColor(int(hex_color[1:3], 16), int(hex_color[3:5], 16), int(hex_color[5:7], 16))
        add_rect(slide, x, 1.15, card_w - 0.12, 0.32, header_color)
        add_text_box(slide, a["company"], x + 0.05, 1.19, card_w - 0.2, 0.25, size=9, bold=True, color=TEXT_LIGHT)
        for j, v in enumerate(a.get("top_videos", [])[:3]):
            vy = 1.57 + j * 1.18
            add_text_box(slide, f"#{j+1} {v['title'][:45]}...", x + 0.08, vy, card_w - 0.2, 0.35, size=8, bold=True, color=TEXT_DARK, wrap=True)
            add_text_box(slide, f"👁 {format_number(v['views'])}   👍 {format_number(v['likes'])}", x + 0.08, vy + 0.38, card_w - 0.2, 0.25, size=7.5, color=TEXT_MID)
            
    # SLIDE 5
    slide = add_light_slide(prs, "Engagement Analysis", "Average views and engagement rate per video")
    fig_eng = make_grouped_bar(companies, ["Avg Views", "Avg Likes", "Avg Comments"], [[a["avg_views"], a["avg_likes"], a["avg_comments"]] for a in analyzed], "Average Engagement Metrics per Video", figsize=(5.8, 3.4))
    add_chart_image(slide, fig_eng, 0.3, 1.1, 5.8, 3.4)
    fig_er = make_bar_chart(companies, [a["engagement_rate"] for a in analyzed], "Engagement Rate (%)", "%", None, figsize=(3.5, 3.4))
    add_chart_image(slide, fig_er, 6.2, 1.1, 3.5, 3.4)
    
    # SLIDE 6
    slide = add_light_slide(prs, "Content Topics & Themes", "What each brand covers — and what they're missing")
    add_text_box(slide, "Top content themes by keyword frequency in titles and tags:", 0.35, 1.1, 9.3, 0.3, size=10, color=TEXT_MID)
    for i, a in enumerate(analyzed):
        col_x = 0.35 + (i % 2) * 4.7
        row_y_pos = 1.55 + (i // 2) * 1.8
        add_rect(slide, col_x, row_y_pos, 4.4, 1.65, RGBColor(0xFF, 0xFF, 0xFF))
        hex_color = CHART_COLORS_MPL[a["color_index"] % 5]
        header_color = RGBColor(int(hex_color[1:3], 16), int(hex_color[3:5], 16), int(hex_color[5:7], 16))
        add_rect(slide, col_x, row_y_pos, 4.4, 0.28, header_color)
        add_text_box(slide, a["company"], col_x + 0.1, row_y_pos + 0.04, 4.2, 0.22, size=9, bold=True, color=TEXT_LIGHT)
        themes = a.get("top_themes", [])[:6]
        add_text_box(slide, "  ·  ".join(f"#{t}" for t in themes) if themes else "No theme data", col_x + 0.1, row_y_pos + 0.33, 4.2, 1.2, size=8.5, color=TEXT_DARK, wrap=True)

    # SLIDE 7
    slide = add_light_slide(prs, "Posting Frequency & Consistency", "Upload cadence over the last 12 months")
    fig_line = make_line_chart(companies, [a["month_counts"] for a in analyzed], figsize=(9.0, 3.6))
    add_chart_image(slide, fig_line, 0.3, 1.1, 9.0, 3.6)
    
    # SLIDE 8
    slide = add_dark_slide(prs, "Gap Analysis", "Unexplored content territories and format opportunities")
    add_text_box(slide, "TOPIC GAPS BY COMPANY", 0.35, 1.35, 4.2, 0.3, size=10, bold=True, color=ACCENT1)
    
    # SLIDE 9
    slide = add_dark_slide(prs, "Video Marketing Recommendations", f"Actionable steps for {your_company} based on analysis")
    add_rect(slide, 0.35, 1.3, 9.3, 0.7, MID_BG)
    add_text_box(slide, "🎯", 0.45, 1.38, 0.4, 0.55, size=14)
    add_text_box(slide, "Content Velocity", 0.9, 1.35, 2.0, 0.28, size=9, bold=True, color=ACCENT1)
    add_text_box(slide, "Match or exceed competitor upload cadence. Build a 90-day content calendar.", 0.9, 1.62, 8.55, 0.3, size=8.5, color=TEXT_LIGHT, wrap=True)
    
    # SLIDE 10
    slide = add_dark_slide(prs, "Competitive Scorecard", "Ranking all companies on key video marketing metrics")
    values_matrix = [[a["sub_score"], a["view_score"], a["eng_score"], a["freq_score"], min(100, int(a["channel"].get("total_videos", 0) / max(1, max(b["channel"].get("total_videos", 1) for b in analyzed)) * 100))] for a in analyzed]
    fig_radar = make_radar_chart(companies, ["Subscribers", "Avg Views", "Engagement", "Frequency", "Content Volume"], [[v/100 for v in row] for row in values_matrix], figsize=(5.5, 4.2))
    add_chart_image(slide, fig_radar, 0.2, 1.1, 5.5, 4.2)
    add_text_box(slide, "FINAL RANKINGS", 5.9, 1.15, 3.7, 0.3, size=11, bold=True, color=ACCENT1)
    
    table_y = 1.55
    for rank, a in enumerate(ranked):
        add_rect(slide, 5.9, table_y, 3.8, 0.62, MID_BG)
        add_text_box(slide, a["company"], 6.45, table_y + 0.07, 2.2, 0.25, size=10, bold=True, color=TEXT_LIGHT)
        add_text_box(slide, f"Score: {a['score']}/100", 6.45, table_y + 0.32, 1.8, 0.22, size=8.5, color=ACCENT3)
        table_y += 0.72

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf

# ─── ROUTES ──────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("static", "index.html")

def get_api_key():
    return request.headers.get('X-YouTube-Key', '').strip() or YOUTUBE_API_KEY

@app.route("/api/analyze", methods=["POST"])
def analyze():
    global YOUTUBE_API_KEY
    try:
        body = request.get_json()
        your_company = body.get("your_company", "").strip()
        competitors  = [c.strip() for c in body.get("competitors", []) if c.strip()]
        
        if not your_company: return jsonify({"error": "Company name required"}), 400
        
        original_key = YOUTUBE_API_KEY
        YOUTUBE_API_KEY = get_api_key()
        
        companies_data = [fetch_company_data(name) for name in [your_company] + competitors[:4]]
        YOUTUBE_API_KEY = original_key
        
        analyzed = compute_scores(analyze_data(companies_data))
        
        result = []
        for a in analyzed:
            result.append({
                "company": a["company"], "is_mock": a.get("is_mock", False),
                "channel": {"title": a["channel"].get("channel_title", a["company"]), "subscribers": a["channel"].get("subscribers", 0), "total_videos": a["channel"].get("total_videos", 0), "total_views": a["channel"].get("total_views", 0)},
                "avg_views": a["avg_views"], "engagement_rate": a["engagement_rate"], "uploads_per_month": a["uploads_per_month"], "score": a["score"],
                "top_themes": a.get("top_themes", []), "top_videos": [{"title": v["title"], "views": v["views"], "likes": v["likes"], "comments": v["comments"], "published_at": v["published_at"], "id": v["id"]} for v in a.get("top_videos", [])[:5]]
            })
        return jsonify({"companies": result, "your_company": your_company})
    except Exception as e:
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route("/api/download", methods=["POST"])
def download():
    global YOUTUBE_API_KEY
    try:
        body = request.get_json()
        your_company = body.get("your_company", "").strip()
        competitors  = [c.strip() for c in body.get("competitors", []) if c.strip()]
        
        if not your_company: return jsonify({"error": "Company name required"}), 400
        
        original_key = YOUTUBE_API_KEY
        YOUTUBE_API_KEY = get_api_key()
        
        companies_data = [fetch_company_data(name) for name in [your_company] + competitors[:4]]
        YOUTUBE_API_KEY = original_key
        
        analyzed = compute_scores(analyze_data(companies_data))
        pptx_buf = build_pptx(analyzed, your_company)
        
        # Safe Header Approach that bypasses all Flask version errors
        safe_company = re.sub(r'[^A-Za-z0-9_]', '', your_company.replace(' ', '_'))
        filename = f"video_intel_{safe_company}.pptx"
        
        response = make_response(pptx_buf.read())
        response.headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
        response.headers.set('Content-Disposition', 'attachment', filename=filename)
        return response

    except Exception as e:
        print(traceback.format_exc())
        return jsonify({"error": f"Presentation compilation failed: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
