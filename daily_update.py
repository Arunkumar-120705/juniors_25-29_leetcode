import requests
import gspread
import time

from datetime import datetime, timedelta, UTC
from oauth2client.service_account import ServiceAccountCredentials

# ==========================================================
# CONFIGURATION
# ==========================================================

SHEET_ID = "1q86fO_1AT3fesQAFFYE3sdSpm56BQ9fNZQdXDUtJfWk"
CREDENTIALS_FILE = "credentials.json"

REQUEST_TIMEOUT = 20
RETRY_COUNT = 3
RETRY_DELAY = 2


# ==========================================================
# HELPER FUNCTIONS
# ==========================================================

def safe_int(value):
    """
    Safely convert anything to integer.
    Empty strings or invalid values become 0.
    """
    try:
        return int(str(value).strip())
    except:
        return 0


def col_to_letter(col):
    """
    Convert column number to Excel/Google Sheet column.

    Example:
        1 -> A
        2 -> B
        27 -> AA
    """

    result = ""

    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result

    return result


def parse_today_cell(cell):
    """
    Converts

        5 (3/2/0)

    into

        total = 5
        easy = 3
        medium = 2
        hard = 0
    """

    if not cell:
        return 0, 0, 0, 0

    cell = cell.strip()

    if cell == "" or cell.upper() == "ERR":
        return 0, 0, 0, 0

    try:
        total_part, rest = cell.split(" ", 1)

        total = int(total_part)

        rest = rest.strip()[1:-1]

        easy, medium, hard = rest.split("/")

        return (
            safe_int(total),
            safe_int(easy),
            safe_int(medium),
            safe_int(hard)
        )

    except:
        return 0, 0, 0, 0


# ==========================================================
# DATE (IST)
# ==========================================================

def get_today():
    """
    Returns today's date in IST.

    Example:
        2026-07-29
    """

    utc = datetime.now(UTC)

    ist = utc + timedelta(hours=5, minutes=30)

    # Prevent next-day issue around midnight GitHub runs
    if ist.hour < 2:
        ist -= timedelta(days=1)

    return ist.strftime("%Y-%m-%d")

# ==========================================================
# LEETCODE GRAPHQL API
# ==========================================================

GRAPHQL_URL = "https://leetcode.com/graphql"

GRAPHQL_QUERY = """
query($u:String!){
  matchedUser(username:$u){
    submitStats{
      acSubmissionNum{
        difficulty
        count
      }
    }
  }
}
"""


def get_stats(username):
    """
    Returns:
        {
            "easy": int,
            "medium": int,
            "hard": int,
            "total": int
        }

    Returns None if the user does not exist
    or the API request fails.
    """

    for attempt in range(RETRY_COUNT):

        try:

            response = requests.post(
                GRAPHQL_URL,
                json={
                    "query": GRAPHQL_QUERY,
                    "variables": {
                        "u": username
                    }
                },
                timeout=REQUEST_TIMEOUT
            )

            data = response.json()

            user = data.get("data", {}).get("matchedUser")

            if not user:
                return None

            stats = user["submitStats"]["acSubmissionNum"]

            easy = next(
                (x["count"] for x in stats if x["difficulty"] == "Easy"),
                0
            )

            medium = next(
                (x["count"] for x in stats if x["difficulty"] == "Medium"),
                0
            )

            hard = next(
                (x["count"] for x in stats if x["difficulty"] == "Hard"),
                0
            )

            return {
                "easy": safe_int(easy),
                "medium": safe_int(medium),
                "hard": safe_int(hard),
                "total": safe_int(easy + medium + hard)
            }

        except Exception:

            if attempt < RETRY_COUNT - 1:
                time.sleep(RETRY_DELAY)

    return None


# ==========================================================
# GOOGLE SHEETS AUTHENTICATION
# ==========================================================

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

client = gspread.authorize(
    ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_FILE,
        scope
    )
)

sheet = client.open_by_key(SHEET_ID).sheet1


# ==========================================================
# LOAD SHEET
# ==========================================================

today = get_today()

values = sheet.get_all_values()

header = values[0]


# ==========================================================
# CREATE TODAY'S COLUMN IF NEEDED
# ==========================================================

if today not in header:

    header.append(today)

    sheet.update(
        range_name="1:1",
        values=[header]
    )

    values = sheet.get_all_values()

    header = values[0]


# ==========================================================
# COLUMN INDEXES
# ==========================================================

idx_name = header.index("Name") + 1
idx_user = header.index("LeetCodeUsername") + 1
idx_baseline = header.index("BaselineTotal") + 1
idx_total = header.index("TotalSolved") + 1

idx_prev_easy = header.index("PrevEasy") + 1
idx_prev_medium = header.index("PrevMedium") + 1
idx_prev_hard = header.index("PrevHard") + 1
idx_prev_total = header.index("PrevTotal") + 1

today_col = header.index(today) + 1


# ==========================================================
# LISTS FOR BATCH UPDATE
# ==========================================================

today_values = []

new_prev = []

new_total = []


# ==========================================================
# COUNTERS
# ==========================================================

updated = 0
unchanged = 0
failed = 0
recovered = 0


print("=" * 60)
print(f"📅 DAILY UPDATE : {today}")
print("=" * 60)

# ==========================================================
# START PROCESSING STUDENTS
# ==========================================================

for row_number, row in enumerate(values[1:], start=2):

    def cell(col):
        idx = col - 1

        if idx >= len(row):
            return ""

        return row[idx]

    name = cell(idx_name).strip()
    username = cell(idx_user).strip()

    # ------------------------------------------------------
    # Skip Empty Rows
    # ------------------------------------------------------

    if username == "":

        today_values.append([""])
        new_prev.append(["", "", "", ""])
        new_total.append([""])

        continue

    # ------------------------------------------------------
    # Skip Section Headers
    # ------------------------------------------------------

    if username.lower() == "leetcodeusername":

        today_values.append([""])
        new_prev.append(["", "", "", ""])
        new_total.append([""])

        continue

    # ------------------------------------------------------
    # Read Previous Sheet Values
    # ------------------------------------------------------

    baseline = safe_int(cell(idx_baseline))

    prev_easy = safe_int(cell(idx_prev_easy))
    prev_medium = safe_int(cell(idx_prev_medium))
    prev_hard = safe_int(cell(idx_prev_hard))
    prev_total = safe_int(cell(idx_prev_total))

    total_sheet = safe_int(cell(idx_total))

    existing_today = cell(today_col)

    old_total, old_easy, old_medium, old_hard = parse_today_cell(
        existing_today
    )

    # ------------------------------------------------------
    # Fetch Latest LeetCode Stats
    # ------------------------------------------------------

    stats = get_stats(username)

    # ------------------------------------------------------
    # API Failed
    # ------------------------------------------------------

    if stats is None:

        failed += 1

        today_values.append([
            existing_today if existing_today else "ERR"
        ])

        new_prev.append([
            prev_easy,
            prev_medium,
            prev_hard,
            prev_total
        ])

        new_total.append([
            total_sheet
        ])

        print(f"❌ {username}")

        continue

    easy_now = stats["easy"]
    medium_now = stats["medium"]
    hard_now = stats["hard"]
    total_now = stats["total"]

    # ------------------------------------------------------
    # Detect Recovery
    # ------------------------------------------------------

    if total_now > total_sheet:

        recovered += 1

        print(
            f"🔄 Recovering {username} "
            f"({total_sheet} → {total_now})"
        )

    # ------------------------------------------------------
    # Calculate Today's Progress
    # ------------------------------------------------------

    delta_easy = max(easy_now - prev_easy, 0)
    delta_medium = max(medium_now - prev_medium, 0)
    delta_hard = max(hard_now - prev_hard, 0)
    delta_total = max(total_now - prev_total, 0)

    # ------------------------------------------------------
    # Student Solved Problems Today
    # ------------------------------------------------------

    if delta_total > 0:

        new_today_total = old_total + delta_total
        new_today_easy = old_easy + delta_easy
        new_today_medium = old_medium + delta_medium
        new_today_hard = old_hard + delta_hard

        today_string = (
            f"{new_today_total} "
            f"({new_today_easy}/{new_today_medium}/{new_today_hard})"
        )

        today_values.append([today_string])

        updated += 1

        print(
            f"✅ {username:<30}"
            f"+{delta_total:<3}"
            f" ({delta_easy}/{delta_medium}/{delta_hard})"
        )

    else:

        # Keep Previous Value
        if existing_today:

            today_values.append([existing_today])

        else:

            today_values.append(["0 (0/0/0)"])

        unchanged += 1

        print(f"➖ {username}")

    # ------------------------------------------------------
    # Update Previous Stats
    # ------------------------------------------------------

    new_prev.append([
        easy_now,
        medium_now,
        hard_now,
        total_now
    ])

    # ------------------------------------------------------
    # Update TotalSolved
    # ------------------------------------------------------

    new_total.append([
        total_now
    ])

    # ------------------------------------------------------
    # Delay (Avoid LeetCode Rate Limit)
    # ------------------------------------------------------

    time.sleep(0.8)


# ==========================================================
# BATCH UPDATE GOOGLE SHEET
# ==========================================================

last_row = len(values)

today_letter = col_to_letter(today_col)
total_letter = col_to_letter(idx_total)

prev_start = col_to_letter(idx_prev_easy)
prev_end = col_to_letter(idx_prev_total)

print("\n📤 Updating Google Sheet...\n")

# ----------------------------------------------------------
# Update Today's Progress Column
# ----------------------------------------------------------

sheet.update(
    range_name=f"{today_letter}2:{today_letter}{last_row}",
    values=today_values
)

# ----------------------------------------------------------
# Update Previous Statistics
# ----------------------------------------------------------

sheet.update(
    range_name=f"{prev_start}2:{prev_end}{last_row}",
    values=new_prev
)

# ----------------------------------------------------------
# Update TotalSolved Column
# ----------------------------------------------------------

sheet.update(
    range_name=f"{total_letter}2:{total_letter}{last_row}",
    values=new_total
)

print("✅ Google Sheet Updated Successfully")

# ==========================================================
# SUMMARY REPORT
# ==========================================================

print("\n" + "=" * 65)
print("                DAILY UPDATE SUMMARY")
print("=" * 65)

print(f"📅 Date               : {today}")
print(f"✅ Updated Students   : {updated}")
print(f"➖ No Changes         : {unchanged}")
print(f"🔄 Recovered          : {recovered}")
print(f"❌ Failed             : {failed}")
print(f"👨‍🎓 Total Processed   : {updated + unchanged + failed}")

print("=" * 65)
print("🎉 Daily Update Completed Successfully!")
print("=" * 65)