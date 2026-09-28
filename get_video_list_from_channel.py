import re
import os
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timezone

import pandas as pd
from tqdm import tqdm
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from api import API_KEY


# ============================================================
# CONFIG
# ============================================================

DEFAULT_START_YEAR = 2019
DEFAULT_START_MONTH = 1

DEFAULT_END_YEAR = 2021
DEFAULT_END_MONTH = 12


# ============================================================
# YOUTUBE API
# ============================================================

youtube = build(
    "youtube",
    "v3",
    developerKey=API_KEY
)


# ============================================================
# GLOBAL SETTINGS
# ============================================================

START_DATE = None
END_DATE = None
OUTPUT_FILE = None


# ============================================================
# DATE HELPERS
# ============================================================

def get_month_start(year, month):
    """
    Trả về ngày đầu tháng theo UTC.
    """
    return datetime(
        year,
        month,
        1,
        0,
        0,
        0,
        tzinfo=timezone.utc
    )


def get_month_end(year, month):
    """
    Trả về thời điểm cuối tháng theo UTC.
    """

    if month == 12:
        next_month = datetime(
            year + 1,
            1,
            1,
            0,
            0,
            0,
            tzinfo=timezone.utc
        )
    else:
        next_month = datetime(
            year,
            month + 1,
            1,
            0,
            0,
            0,
            tzinfo=timezone.utc
        )

    return next_month.timestamp() - 1


def format_date_for_filename(year1, month1, year2, month2):
    return (
        f"{year1}-{month1:02d}_"
        f"{year2}-{month2:02d}"
    )


# ============================================================
# GUI INPUT
# ============================================================

def show_input_dialog():

    result = {
        "channel": None,
        "start_year": None,
        "start_month": None,
        "end_year": None,
        "end_month": None,
        "cancelled": True
    }

    root = tk.Tk()

    root.title("YouTube Video Scanner")
    root.geometry("600x390")
    root.resizable(False, False)

    # --------------------------------------------------------
    # STYLE
    # --------------------------------------------------------

    style = ttk.Style()

    try:
        style.theme_use("vista")
    except:
        pass

    style.configure(
        "Title.TLabel",
        font=("Arial", 16, "bold")
    )

    style.configure(
        "Normal.TLabel",
        font=("Arial", 10)
    )

    style.configure(
        "Big.TButton",
        font=("Arial", 11, "bold"),
        padding=10
    )

    # --------------------------------------------------------
    # MAIN FRAME
    # --------------------------------------------------------

    main_frame = ttk.Frame(
        root,
        padding=25
    )

    main_frame.pack(
        fill="both",
        expand=True
    )

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    title = ttk.Label(
        main_frame,
        text="YOUTUBE VIDEO SCANNER",
        style="Title.TLabel"
    )

    title.pack(
        pady=(0, 20)
    )

    # --------------------------------------------------------
    # CHANNEL
    # --------------------------------------------------------

    channel_label = ttk.Label(
        main_frame,
        text="Link kênh / @handle / Channel ID:",
        style="Normal.TLabel"
    )

    channel_label.pack(
        anchor="w"
    )

    channel_var = tk.StringVar()

    channel_entry = ttk.Entry(
        main_frame,
        textvariable=channel_var,
        font=("Arial", 11)
    )

    channel_entry.pack(
        fill="x",
        pady=(5, 18)
    )

    channel_entry.focus()

    # --------------------------------------------------------
    # PERIOD TITLE
    # --------------------------------------------------------

    period_label = ttk.Label(
        main_frame,
        text="Khoảng thời gian video:",
        style="Normal.TLabel"
    )

    period_label.pack(
        anchor="w"
    )

    # --------------------------------------------------------
    # PERIOD FRAME
    # --------------------------------------------------------

    period_frame = ttk.Frame(main_frame)

    period_frame.pack(
        fill="x",
        pady=(8, 20)
    )

    # ========================================================
    # START
    # ========================================================

    start_frame = ttk.LabelFrame(
        period_frame,
        text="  Từ tháng  ",
        padding=12
    )

    start_frame.pack(
        side="left",
        fill="x",
        expand=True,
        padx=(0, 8)
    )

    start_month_var = tk.StringVar(
        value="01"
    )

    start_year_var = tk.StringVar(
        value=str(DEFAULT_START_YEAR)
    )

    ttk.Label(
        start_frame,
        text="Tháng:"
    ).grid(
        row=0,
        column=0,
        padx=(0, 5)
    )

    start_month_combo = ttk.Combobox(
        start_frame,
        textvariable=start_month_var,
        values=[
            f"{i:02d}"
            for i in range(1, 13)
        ],
        width=5,
        state="readonly"
    )

    start_month_combo.grid(
        row=0,
        column=1,
        padx=(0, 15)
    )

    ttk.Label(
        start_frame,
        text="Năm:"
    ).grid(
        row=0,
        column=2,
        padx=(0, 5)
    )

    years = [
        str(year)
        for year in range(
            2005,
            datetime.now().year + 1
        )
    ]

    start_year_combo = ttk.Combobox(
        start_frame,
        textvariable=start_year_var,
        values=years,
        width=7,
        state="readonly"
    )

    start_year_combo.grid(
        row=0,
        column=3
    )

    # ========================================================
    # END
    # ========================================================

    end_frame = ttk.LabelFrame(
        period_frame,
        text="  Đến tháng  ",
        padding=12
    )

    end_frame.pack(
        side="left",
        fill="x",
        expand=True,
        padx=(8, 0)
    )

    end_month_var = tk.StringVar(
        value="12"
    )

    end_year_var = tk.StringVar(
        value=str(DEFAULT_END_YEAR)
    )

    ttk.Label(
        end_frame,
        text="Tháng:"
    ).grid(
        row=0,
        column=0,
        padx=(0, 5)
    )

    end_month_combo = ttk.Combobox(
        end_frame,
        textvariable=end_month_var,
        values=[
            f"{i:02d}"
            for i in range(1, 13)
        ],
        width=5,
        state="readonly"
    )

    end_month_combo.grid(
        row=0,
        column=1,
        padx=(0, 15)
    )

    ttk.Label(
        end_frame,
        text="Năm:"
    ).grid(
        row=0,
        column=2,
        padx=(0, 5)
    )

    end_year_combo = ttk.Combobox(
        end_frame,
        textvariable=end_year_var,
        values=years,
        width=7,
        state="readonly"
    )

    end_year_combo.grid(
        row=0,
        column=3
    )

    # --------------------------------------------------------
    # EXAMPLE
    # --------------------------------------------------------

    info_label = ttk.Label(
        main_frame,
        text=(
            "Ví dụ: chọn 01/2019 → 12/2021\n"
            "sẽ quét toàn bộ video được đăng trong giai đoạn này."
        ),
        style="Normal.TLabel",
        justify="left"
    )

    info_label.pack(
        anchor="w",
        pady=(0, 20)
    )

    # --------------------------------------------------------
    # BUTTONS
    # --------------------------------------------------------

    button_frame = ttk.Frame(main_frame)

    button_frame.pack(
        fill="x"
    )

    def start_scan():

        channel_input = channel_var.get().strip()

        if not channel_input:

            messagebox.showwarning(
                "Thiếu thông tin",
                "Bạn chưa nhập link kênh YouTube."
            )

            channel_entry.focus()
            return

        try:

            start_year = int(
                start_year_var.get()
            )

            start_month = int(
                start_month_var.get()
            )

            end_year = int(
                end_year_var.get()
            )

            end_month = int(
                end_month_var.get()
            )

        except ValueError:

            messagebox.showerror(
                "Lỗi",
                "Tháng hoặc năm không hợp lệ."
            )

            return

        # ----------------------------------------------------
        # CHECK PERIOD
        # ----------------------------------------------------

        start_dt = get_month_start(
            start_year,
            start_month
        )

        end_dt = datetime.fromtimestamp(
            get_month_end(
                end_year,
                end_month
            ),
            tz=timezone.utc
        )

        if start_dt > end_dt:

            messagebox.showerror(
                "Khoảng thời gian không hợp lệ",
                "Tháng bắt đầu phải nhỏ hơn hoặc bằng tháng kết thúc."
            )

            return

        # ----------------------------------------------------
        # SAVE
        # ----------------------------------------------------

        result["channel"] = channel_input

        result["start_year"] = start_year
        result["start_month"] = start_month

        result["end_year"] = end_year
        result["end_month"] = end_month

        result["cancelled"] = False

        root.destroy()

    def cancel():

        result["cancelled"] = True

        root.destroy()

    start_button = ttk.Button(
        button_frame,
        text="▶  BẮT ĐẦU QUÉT",
        style="Big.TButton",
        command=start_scan
    )

    start_button.pack(
        side="left",
        fill="x",
        expand=True,
        padx=(0, 8)
    )

    cancel_button = ttk.Button(
        button_frame,
        text="HỦY",
        style="Big.TButton",
        command=cancel
    )

    cancel_button.pack(
        side="left",
        fill="x",
        expand=True,
        padx=(8, 0)
    )

    # --------------------------------------------------------
    # ENTER / ESC
    # --------------------------------------------------------

    root.bind(
        "<Return>",
        lambda event: start_scan()
    )

    root.bind(
        "<Escape>",
        lambda event: cancel()
    )

    # --------------------------------------------------------
    # CLOSE WINDOW
    # --------------------------------------------------------

    root.protocol(
        "WM_DELETE_WINDOW",
        cancel
    )

    root.mainloop()

    return result


# ============================================================
# RESOLVE CHANNEL
# ============================================================

def extract_channel_id(url_or_input):

    value = url_or_input.strip()

    # --------------------------------------------------------
    # CHANNEL ID trực tiếp
    # --------------------------------------------------------

    if re.fullmatch(
        r"UC[a-zA-Z0-9_-]{20,}",
        value
    ):
        return value

    # --------------------------------------------------------
    # /channel/UC...
    # --------------------------------------------------------

    m = re.search(
        r"/channel/(UC[a-zA-Z0-9_-]+)",
        value
    )

    if m:
        return m.group(1)

    # --------------------------------------------------------
    # @handle trong URL
    # --------------------------------------------------------

    m = re.search(
        r"youtube\.com/@([^/?]+)",
        value
    )

    if m:

        handle = m.group(1)

        print(
            f"\n[>] Tìm channel theo handle: @{handle}"
        )

        try:

            response = youtube.channels().list(
                part="id,snippet,contentDetails",
                forHandle=handle
            ).execute()

            if response.get("items"):

                return response["items"][0]["id"]

        except HttpError as e:

            print(
                "\n[X] YouTube API ERROR:"
            )

            print(e)

            return None

    # --------------------------------------------------------
    # /user/username
    # --------------------------------------------------------

    m = re.search(
        r"youtube\.com/user/([^/?]+)",
        value
    )

    if m:

        username = m.group(1)

        print(
            f"\n[>] Tìm channel theo username: {username}"
        )

        try:

            response = youtube.channels().list(
                part="id,snippet,contentDetails",
                forUsername=username
            ).execute()

            if response.get("items"):

                return response["items"][0]["id"]

        except HttpError as e:

            print(
                "\n[X] YouTube API ERROR:"
            )

            print(e)

            return None

    # --------------------------------------------------------
    # Nếu nhập chỉ @handle
    # --------------------------------------------------------

    if value.startswith("@"):

        handle = value[1:]

        print(
            f"\n[>] Tìm channel theo handle: @{handle}"
        )

        try:

            response = youtube.channels().list(
                part="id,snippet,contentDetails",
                forHandle=handle
            ).execute()

            if response.get("items"):

                return response["items"][0]["id"]

        except HttpError as e:

            print(
                "\n[X] YouTube API ERROR:"
            )

            print(e)

            return None

    return None


# ============================================================
# CHANNEL INFO
# ============================================================

def get_channel_info(channel_id):

    try:

        response = youtube.channels().list(
            part="snippet,contentDetails,statistics",
            id=channel_id
        ).execute()

    except HttpError as e:

        print(
            "\n[X] YouTube API ERROR:"
        )

        print(e)

        return None

    if not response.get("items"):

        return None

    item = response["items"][0]

    return {
        "channel_id":
            item["id"],

        "channel_title":
            item["snippet"]["title"],

        "channel_url":
            f"https://www.youtube.com/channel/{item['id']}",

        "uploads_playlist":
            item["contentDetails"]["relatedPlaylists"]["uploads"],

        "subscribers":
            item.get("statistics", {}).get(
                "subscriberCount"
            ),

        "video_count":
            item.get("statistics", {}).get(
                "videoCount"
            )
    }


# ============================================================
# GET ALL VIDEOS FROM UPLOADS PLAYLIST
# ============================================================

def get_all_video_ids(uploads_playlist):

    video_ids = []

    next_page_token = None

    print(
        "\n[+] Đang lấy danh sách video của kênh..."
    )

    while True:

        try:

            response = youtube.playlistItems().list(
                part="contentDetails,snippet",
                playlistId=uploads_playlist,
                maxResults=50,
                pageToken=next_page_token
            ).execute()

        except HttpError as e:

            print(
                "\n[X] YouTube API ERROR:"
            )

            print(e)

            break

        items = response.get(
            "items",
            []
        )

        for item in items:

            video_id = item[
                "contentDetails"
            ][
                "videoId"
            ]

            video_ids.append(
                video_id
            )

        print(
            f"\r[+] Đã lấy: {len(video_ids):,} video",
            end="",
            flush=True
        )

        next_page_token = response.get(
            "nextPageToken"
        )

        if not next_page_token:

            break

    print()

    return video_ids


# ============================================================
# GET VIDEO DETAILS
# ============================================================

def get_video_details(
    video_ids,
    start_dt,
    end_dt
):

    results = []

    if not video_ids:

        return results

    print(
        "\n[+] Đang lấy thông tin video..."
    )

    total = len(video_ids)

    for i in tqdm(
        range(0, total, 50),
        desc="VIDEOS",
        unit="batch"
    ):

        batch = video_ids[
            i:i + 50
        ]

        try:

            response = youtube.videos().list(
                part=(
                    "snippet,"
                    "contentDetails,"
                    "statistics,"
                    "status"
                ),
                id=",".join(batch)
            ).execute()

        except HttpError as e:

            print(
                "\n[X] YouTube API ERROR:"
            )

            print(e)

            continue

        for item in response.get(
            "items",
            []
        ):

            snippet = item.get(
                "snippet",
                {}
            )

            statistics = item.get(
                "statistics",
                {}
            )

            content = item.get(
                "contentDetails",
                {}
            )

            status = item.get(
                "status",
                {}
            )

            published_at = snippet.get(
                "publishedAt"
            )

            if not published_at:

                continue

            # ------------------------------------------------
            # Chuyển sang datetime
            # ------------------------------------------------

            try:

                dt = datetime.fromisoformat(
                    published_at.replace(
                        "Z",
                        "+00:00"
                    )
                )

            except Exception:

                continue

            # ------------------------------------------------
            # FILTER THEO KHOẢNG ĐƯỢC CHỌN
            # ------------------------------------------------

            if not (
                start_dt <= dt <= end_dt
            ):

                continue

            video_id = item["id"]

            duration = content.get(
                "duration",
                ""
            )

            results.append({

                "Published":
                    dt.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),

                "Year":
                    dt.year,

                "Month":
                    dt.month,

                "Title":
                    snippet.get(
                        "title",
                        ""
                    ),

                "Video ID":
                    video_id,

                "Video URL":
                    (
                        "https://www.youtube.com/watch?v="
                        + video_id
                    ),

                "Channel":
                    snippet.get(
                        "channelTitle",
                        ""
                    ),

                "Channel ID":
                    snippet.get(
                        "channelId",
                        ""
                    ),

                "Description":
                    snippet.get(
                        "description",
                        ""
                    ),

                "Duration":
                    duration,

                "Views":
                    statistics.get(
                        "viewCount",
                        ""
                    ),

                "Likes":
                    statistics.get(
                        "likeCount",
                        ""
                    ),

                "Comments":
                    statistics.get(
                        "commentCount",
                        ""
                    ),

                "Privacy":
                    status.get(
                        "privacyStatus",
                        ""
                    ),

                "Live Broadcast":
                    snippet.get(
                        "liveBroadcastContent",
                        ""
                    ),

            })

    return results


# ============================================================
# MAIN
# ============================================================

def main():

    global START_DATE
    global END_DATE
    global OUTPUT_FILE

    print(
        "=" * 80
    )

    print(
        "YOUTUBE VIDEO SCANNER"
    )

    print(
        "=" * 80
    )

    # ========================================================
    # GUI
    # ========================================================

    settings = show_input_dialog()

    if settings["cancelled"]:

        print(
            "\n[!] Đã hủy."
        )

        return

    # ========================================================
    # GET SETTINGS
    # ========================================================

    channel_input = settings[
        "channel"
    ]

    start_year = settings[
        "start_year"
    ]

    start_month = settings[
        "start_month"
    ]

    end_year = settings[
        "end_year"
    ]

    end_month = settings[
        "end_month"
    ]

    # --------------------------------------------------------
    # DATE RANGE
    # --------------------------------------------------------

    START_DATE = get_month_start(
        start_year,
        start_month
    )

    END_DATE = datetime.fromtimestamp(
        get_month_end(
            end_year,
            end_month
        ),
        tz=timezone.utc
    )

    # --------------------------------------------------------
    # OUTPUT NAME
    # --------------------------------------------------------

    period_name = format_date_for_filename(
        start_year,
        start_month,
        end_year,
        end_month
    )

    OUTPUT_FILE = (
        f"YOUTUBE_VIDEOS_{period_name}.xlsx"
    )

    # ========================================================
    # SHOW SETTINGS
    # ========================================================

    print(
        "\n" + "=" * 80
    )

    print(
        "SCAN SETTINGS"
    )

    print(
        "=" * 80
    )

    print(
        f"[+] Channel input : {channel_input}"
    )

    print(
        f"[+] From          : "
        f"{START_DATE.strftime('%Y-%m-%d %H:%M:%S')} UTC"
    )

    print(
        f"[+] To            : "
        f"{END_DATE.strftime('%Y-%m-%d %H:%M:%S')} UTC"
    )

    print(
        f"[+] Output        : {OUTPUT_FILE}"
    )

    # ========================================================
    # RESOLVE
    # ========================================================

    print(
        "\n[1] RESOLVE CHANNEL"
    )

    channel_id = extract_channel_id(
        channel_input
    )

    if not channel_id:

        print(
            "\n[X] Không tìm được Channel ID."
        )

        print(
            "\nBạn có thể nhập dạng:"
        )

        print(
            "  https://www.youtube.com/channel/UCxxxxxxxx"
        )

        print(
            "  https://www.youtube.com/@username"
        )

        print(
            "  @username"
        )

        print(
            "  UCxxxxxxxx"
        )

        return

    print(
        f"[+] Channel ID: {channel_id}"
    )

    # ========================================================
    # CHANNEL INFO
    # ========================================================

    print(
        "\n[2] CHANNEL INFO"
    )

    channel = get_channel_info(
        channel_id
    )

    if not channel:

        print(
            "[X] Không lấy được thông tin channel."
        )

        return

    print(
        f"[+] Channel : "
        f"{channel['channel_title']}"
    )

    print(
        f"[+] ID      : "
        f"{channel['channel_id']}"
    )

    print(
        f"[+] Videos  : "
        f"{channel['video_count']}"
    )

    print(
        f"[+] URL     : "
        f"{channel['channel_url']}"
    )

    # ========================================================
    # ALL VIDEO IDS
    # ========================================================

    print(
        "\n[3] GET VIDEO LIST"
    )

    video_ids = get_all_video_ids(
        channel["uploads_playlist"]
    )

    print(
        f"[+] Tổng video tìm thấy: "
        f"{len(video_ids):,}"
    )

    if not video_ids:

        print(
            "[X] Không có video."
        )

        return

    # ========================================================
    # VIDEO DETAILS
    # ========================================================

    print(
        "\n[4] FILTER VIDEO THEO THỜI GIAN"
    )

    print(
        f"[+] "
        f"{START_DATE.strftime('%Y-%m-%d')}"
        f" -> "
        f"{END_DATE.strftime('%Y-%m-%d')}"
    )

    videos = get_video_details(
        video_ids,
        START_DATE,
        END_DATE
    )

    # ========================================================
    # EXPORT
    # ========================================================

    print(
        "\n[5] EXPORT EXCEL"
    )

    if not videos:

        print(
            "\n[!] Không tìm thấy video nào "
            "trong khoảng thời gian đã chọn."
        )

        return

    df = pd.DataFrame(
        videos
    )

    # --------------------------------------------------------
    # Sort mới nhất -> cũ nhất
    # --------------------------------------------------------

    df = df.sort_values(
        by="Published",
        ascending=False
    )

    # --------------------------------------------------------
    # Reset index
    # --------------------------------------------------------

    df = df.reset_index(
        drop=True
    )

    # --------------------------------------------------------
    # Excel
    # --------------------------------------------------------

    output_path = os.path.join(
        os.getcwd(),
        OUTPUT_FILE
    )

    df.to_excel(
        output_path,
        index=False,
        engine="openpyxl"
    )

    # ========================================================
    # DONE
    # ========================================================

    print(
        "\n" + "=" * 80
    )

    print(
        "DONE"
    )

    print(
        "=" * 80
    )

    print(
        f"\n[+] Channel : "
        f"{channel['channel_title']}"
    )

    print(
        f"[+] Video  : "
        f"{len(df):,}"
    )

    print(
        f"[+] Period : "
        f"{START_DATE.strftime('%Y-%m-%d')}"
        f" -> "
        f"{END_DATE.strftime('%Y-%m-%d')}"
    )

    print(
        "\n[+] File Excel:"
    )

    print(
        output_path
    )

    print(
        "\n" + "=" * 80
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        print(
            "\n\n[!] Đã dừng chương trình."
        )

    except Exception as e:

        print(
            "\n[X] ERROR:"
        )

        print(e)

    input(
        "\nENTER để thoát..."
    )