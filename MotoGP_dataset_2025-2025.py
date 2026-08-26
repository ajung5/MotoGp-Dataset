import requests
import pandas as pd
import time
import re
from urllib.parse import quote

# ============================================================
# CONFIGURATION
# ============================================================

START_SEASON = 2025
END_SEASON = 2025

OUTPUT_FILE = "MotoGP_Historical_Race_Results_2021_2025.xlsx"

BASE_URL = "https://api.motogp.pulselive.com/motogp/v1"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json",
    "Origin": "https://www.motogp.com",
    "Referer": "https://www.motogp.com/"
}

# Delay antar request agar tidak terlalu agresif
REQUEST_DELAY = 0.3


# ============================================================
# HELPER REQUEST
# ============================================================

def get_json(url, params=None, retries=3):

    for attempt in range(retries):

        try:

            response = requests.get(
                url,
                params=params,
                headers=HEADERS,
                timeout=30
            )

            response.raise_for_status()

            return response.json()

        except requests.exceptions.RequestException as e:

            print(
                f"Request gagal ({attempt + 1}/{retries}): "
                f"{e}"
            )

            if attempt < retries - 1:
                time.sleep(2)

    return None


# ============================================================
# GET SEASON ID
# ============================================================

def get_season_id(year):

    url = f"{BASE_URL}/results/seasons"

    data = get_json(url)

    if not data:
        return None

    for season in data:

        if int(season.get("year", 0)) == year:
            return season.get("id")

    return None


# ============================================================
# GET MOTOGP CATEGORY ID
# ============================================================

def get_motogp_category_id(season_uuid):

    url = f"{BASE_URL}/results/categories"

    params = {
        "seasonUuid": season_uuid
    }

    data = get_json(url, params)

    if not data:
        return None

    for category in data:

        name = category.get("name", "").lower()

        if "motogp" in name:

            return category.get("id")

    return None


# ============================================================
# GET EVENTS
# ============================================================

def get_events(season_uuid):

    url = f"{BASE_URL}/results/events"

    params = {
        "seasonUuid": season_uuid,
        "isFinished": "true"
    }

    return get_json(url, params) or []


# ============================================================
# GET SESSIONS
# ============================================================

def get_sessions(event_uuid, category_uuid):

    url = f"{BASE_URL}/results/sessions"

    params = {
        "eventUuid": event_uuid,
        "categoryUuid": category_uuid
    }

    return get_json(url, params) or []


# ============================================================
# FIND MAIN RACE SESSION
# ============================================================

def find_race_session(sessions):

    """
    Mencari session Race / RAC.
    
    Sprint tidak diambil.
    """

    for session in sessions:

        session_type = str(
            session.get("type", "")
        ).upper()

        if session_type == "RAC":

            return session

    return None


# ============================================================
# GET CLASSIFICATION
# ============================================================

def get_classification(session_id, season):

    url = (
        f"{BASE_URL}/results/session/"
        f"{session_id}/classification"
    )

    params = {
        "seasonYear": season,
        "test": "false"
    }

    data = get_json(url, params)

    if not data:
        return []

    return data.get("classification", [])


# ============================================================
# GET GRID
# ============================================================

def get_grid(event_uuid, category_uuid):

    url = (
        f"{BASE_URL}/results/event/"
        f"{event_uuid}/category/"
        f"{category_uuid}/grid"
    )

    data = get_json(url)

    return data or []


# ============================================================
# FORMAT GRAND PRIX INITIAL
# ============================================================

def get_gp_initial(event):

    # Coba ambil short_name
    short_name = event.get("short_name")

    if short_name:
        return short_name.upper()

    # Coba sponsored name
    sponsored_name = event.get(
        "sponsored_name",
        ""
    )

    words = sponsored_name.split()

    # Ambil 3 huruf awal sebagai fallback
    if words:
        text = re.sub(
            r"[^A-Za-z]",
            "",
            words[-1]
        )

        return text[:3].upper()

    return ""


# ============================================================
# FORMAT GRAND PRIX NAME
# ============================================================

def get_grand_prix_name(event):

    name = event.get("name")

    if name:
        return name

    sponsored = event.get(
        "sponsored_name"
    )

    if sponsored:
        return sponsored

    return ""


# ============================================================
# GET CITY
# ============================================================

def get_city(event):

    circuit = event.get("circuit") or {}

    # Beberapa response menggunakan locality
    location = circuit.get("location")

    if location:
        return location

    locality = circuit.get("locality")

    if locality:
        return locality

    return ""


# ============================================================
# GET CIRCUIT NAME
# ============================================================

def get_circuit_name(event):

    circuit = event.get("circuit") or {}

    return (
        circuit.get("name")
        or event.get("circuit_name")
        or ""
    )


# ============================================================
# GET NATION
# ============================================================

def get_nation(event):

    country = event.get("country")

    if isinstance(country, dict):

        return (
            country.get("name")
            or country.get("iso")
            or ""
        )

    if isinstance(country, str):
        return country

    return ""


# ============================================================
# DRIVER NAME
# ============================================================

def get_driver_name(rider):

    return rider.get(
        "full_name",
        ""
    ).strip()


# ============================================================
# LAST NAME
# ============================================================

def get_last_name(full_name):

    if not full_name:
        return ""

    return full_name.split()[-1]


# ============================================================
# DRIVER INITIAL
# ============================================================

def get_driver_initial(full_name):

    if not full_name:
        return ""

    parts = full_name.split()

    if len(parts) == 1:
        return parts[0][0].upper()

    # Format: FN
    return (
        parts[0][0] +
        parts[-1][0]
    ).upper()


# ============================================================
# GRID DICTIONARY
# ============================================================

def create_grid_dictionary(grid_data):

    grid_dict = {}

    for item in grid_data:

        rider = item.get("rider") or {}

        rider_id = (
            rider.get("id")
            or rider.get("legacy_id")
        )

        if rider_id is None:
            continue

        grid_position = item.get(
            "qualifying_position"
        )

        grid_dict[str(rider_id)] = grid_position

    return grid_dict


# ============================================================
# MAIN SCRAPER
# ============================================================

all_results = []

print("=" * 70)
print("MOTOGP HISTORICAL RACE DATA SCRAPER")
print("=" * 70)

print(
    f"Season : {START_SEASON} - {END_SEASON}"
)

print(
    f"Output : {OUTPUT_FILE}"
)

print("=" * 70)


for season in range(
    START_SEASON,
    END_SEASON + 1
):

    print()
    print("=" * 70)
    print(f"SEASON {season}")
    print("=" * 70)

    # --------------------------------------------------------
    # SEASON ID
    # --------------------------------------------------------

    season_uuid = get_season_id(season)

    if not season_uuid:

        print(
            f"[WARNING] Season {season} tidak ditemukan."
        )

        continue

    print(
        f"Season UUID : {season_uuid}"
    )

    # --------------------------------------------------------
    # CATEGORY
    # --------------------------------------------------------

    category_uuid = get_motogp_category_id(
        season_uuid
    )

    if not category_uuid:

        print(
            "[WARNING] Category MotoGP tidak ditemukan."
        )

        continue

    print(
        f"MotoGP Category UUID : {category_uuid}"
    )

    # --------------------------------------------------------
    # EVENTS
    # --------------------------------------------------------

    events = get_events(
        season_uuid
    )

    print(
        f"Total Event : {len(events)}"
    )

    # Sort berdasarkan tanggal
    events = sorted(
        events,
        key=lambda x: (
            x.get("date")
            or x.get("event_date")
            or ""
        )
    )

    # --------------------------------------------------------
    # LOOP EVENT
    # --------------------------------------------------------

    for round_number, event in enumerate(
        events,
        start=1
    ):

        event_uuid = event.get("id")

        if not event_uuid:
            continue

        gp_name = get_grand_prix_name(
            event
        )

        print()
        print(
            f"[{season}] Round {round_number} - {gp_name}"
        )

        # ----------------------------------------------------
        # SESSION
        # ----------------------------------------------------

        sessions = get_sessions(
            event_uuid,
            category_uuid
        )

        time.sleep(REQUEST_DELAY)

        race_session = find_race_session(
            sessions
        )

        if not race_session:

            print(
                "  [SKIP] Main Race tidak ditemukan."
            )

            continue

        session_id = race_session.get("id")

        if not session_id:

            print(
                "  [SKIP] Session ID tidak ditemukan."
            )

            continue

        print(
            f"  Race Session : {session_id}"
        )

        # ----------------------------------------------------
        # CLASSIFICATION
        # ----------------------------------------------------

        classification = get_classification(
            session_id,
            season
        )

        time.sleep(REQUEST_DELAY)

        if not classification:

            print(
                "  [SKIP] Classification kosong."
            )

            continue

        print(
            f"  Riders : {len(classification)}"
        )

        # ----------------------------------------------------
        # GRID
        # ----------------------------------------------------

        grid_data = get_grid(
            event_uuid,
            category_uuid
        )

        time.sleep(REQUEST_DELAY)

        grid_dict = create_grid_dictionary(
            grid_data
        )

        # ----------------------------------------------------
        # FIND FASTEST LAP
        # ----------------------------------------------------

        fastest_lap_rider_id = None

        fastest_lap_time = None

        for result in classification:

            rider = result.get(
                "rider"
            ) or {}

            rider_id = (
                rider.get("id")
                or rider.get("legacy_id")
            )

            best_lap = (
                result.get("best_lap")
                or {}
            )

            lap_time = best_lap.get(
                "time"
            )

            if not lap_time:
                continue

            # Convert time menjadi milidetik
            try:

                parts = lap_time.split(":")

                if len(parts) == 3:

                    minutes = float(parts[0])

                    seconds = float(parts[1])

                    milliseconds = float(
                        parts[2]
                    )

                    total_ms = (
                        minutes * 60000
                        + seconds * 1000
                        + milliseconds
                    )

                elif len(parts) == 2:

                    minutes = float(parts[0])

                    seconds = float(parts[1])

                    total_ms = (
                        minutes * 60000
                        + seconds * 1000
                    )

                else:

                    continue

            except Exception:
                continue

            if (
                fastest_lap_time is None
                or total_ms < fastest_lap_time
            ):

                fastest_lap_time = total_ms

                fastest_lap_rider_id = (
                    rider_id
                )

        # ----------------------------------------------------
        # EVENT INFORMATION
        # ----------------------------------------------------

        circuit_name = get_circuit_name(
            event
        )

        city = get_city(
            event
        )

        nation = get_nation(
            event
        )

        gp_initial = get_gp_initial(
            event
        )

        # ----------------------------------------------------
        # CREATE ROW
        # ----------------------------------------------------

        for result in classification:

            rider = (
                result.get("rider")
                or {}
            )

            rider_id = (
                rider.get("id")
                or rider.get("legacy_id")
            )

            rider_name = get_driver_name(
                rider
            )

            last_name = get_last_name(
                rider_name
            )

            driver_initial = (
                get_driver_initial(
                    rider_name
                )
            )

            # ----------------------------------------------
            # GRID POSITION
            # ----------------------------------------------

            grid_position = grid_dict.get(
                str(rider_id)
            )

            # ----------------------------------------------
            # FINISH POSITION
            # ----------------------------------------------

            finish_position = result.get(
                "position"
            )

            # ----------------------------------------------
            # RACE POINTS
            # ----------------------------------------------

            race_points = result.get(
                "points"
            )

            # Jika API tidak menyediakan points
            # pada classification, default 0.
            if race_points is None:
                race_points = 0

            # ----------------------------------------------
            # FASTEST LAP FLAG
            # ----------------------------------------------

            if (
                rider_id is not None
                and fastest_lap_rider_id is not None
                and str(rider_id)
                == str(fastest_lap_rider_id)
            ):

                fastest_lap = 1

            else:

                fastest_lap = 0

            # ----------------------------------------------
            # APPEND
            # ----------------------------------------------

            row = {

                "Round":
                    round_number,

                "Season":
                    season,

                "IsLatestSeason":
                    1 if season == END_SEASON else 0,

                "Grand Prix":
                    gp_name,

                "Circuit Name":
                    circuit_name,

                "City":
                    city,

                "Nation":
                    nation,

                "GP Initial":
                    gp_initial,

                "Race Type":
                    "Main Race",

                "Driver Name":
                    rider_name,

                "Last Name":
                    last_name,

                "Driver Initial":
                    driver_initial,

                "Grid Position":
                    grid_position,

                "Finish Position":
                    finish_position,

                "Race Points":
                    race_points,

                "Fastest Lap":
                    fastest_lap
            }

            all_results.append(
                row
            )


# ============================================================
# CREATE DATAFRAME
# ============================================================

print()
print("=" * 70)
print("MEMBUAT DATAFRAME")
print("=" * 70)

df = pd.DataFrame(
    all_results
)


# ============================================================
# SORTING
# ============================================================

if not df.empty:

    df = df.sort_values(
        by=[
            "Season",
            "Round",
            "Finish Position"
        ],
        ascending=[
            True,
            True,
            True
        ]
    )

    df = df.reset_index(
        drop=True
    )


# ============================================================
# COLUMN ORDER
# ============================================================

columns = [

    "Round",
    "Season",
    "IsLatestSeason",
    "Grand Prix",
    "Circuit Name",
    "City",
    "Nation",
    "GP Initial",
    "Race Type",
    "Driver Name",
    "Last Name",
    "Driver Initial",
    "Grid Position",
    "Finish Position",
    "Race Points",
    "Fastest Lap"

]

df = df[columns]


# ============================================================
# SAVE TO EXCEL
# ============================================================

print()
print("=" * 70)
print("SAVE EXCEL")
print("=" * 70)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Driver",
        index=False
    )

    worksheet = writer.sheets[
        "Driver"
    ]

    # --------------------------------------------------------
    # AUTO WIDTH
    # --------------------------------------------------------

    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = (
            column_cells[0].column_letter
        )

        for cell in column_cells:

            try:

                cell_length = len(
                    str(cell.value)
                )

                if cell_length > max_length:
                    max_length = cell_length

            except Exception:
                pass

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            35
        )

    # --------------------------------------------------------
    # FREEZE HEADER
    # --------------------------------------------------------

    worksheet.freeze_panes = "A2"

    # --------------------------------------------------------
    # AUTO FILTER
    # --------------------------------------------------------

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )


# ============================================================
# SUMMARY
# ============================================================

print()
print("=" * 70)
print("SELESAI")
print("=" * 70)

print(
    f"Total Rows : {len(df):,}"
)

print(
    f"Total Race : "
    f"{df[['Season', 'Round']].drop_duplicates().shape[0]:,}"
)

print(
    f"File      : {OUTPUT_FILE}"
)

print("=" * 70)