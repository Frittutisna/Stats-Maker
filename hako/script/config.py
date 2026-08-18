import  re
from    PIL import Image, ImageChops, ImageOps

BROWSER_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

COLOR_0     = "#C83232"
COLOR_1     = "#7D327D"
COLOR_2     = "#3232C8"
CONST_CONV  = 3
DIR_CREDS   = "credential"
DIR_JSONS   = "json"
DIR_OUT     = "hako"
DIR_TOURS   = "tour"
FILE_CHANT  = "chant.txt"
FILE_CODES  = "code.txt"
FILE_ALIAS  = "alias.txt"
SCALE_PERF  = -1.5
TEAMS_RE    = r"([^\s(]+)\s*\(([-]?\d+(?:\.\d+)?)\)"
THRESH_CHRL = 50
THRESH_CHRM = 40
THRESH_CHRS = 30
THRESH_SONG = 35
THRESH_TEAM = 4
THRESH_TIME = 17.5
THRESH_WTCH = 5
TOKEN_NTLFY = "nfp_n58sF53EWEurX9VVkKTpL9ekLsbeDp2Xde73"

ANT_URL_ALIAS   = "https://docs.google.com/spreadsheets/d/1JjI2GaHjAGr6dABR9Mo_mZ0CdKbXrvlKHnIpttr22uo/export?format=csv&gid=0"
ANT_KEY_STATS   = "1WadyJ-rpY8AbjO8TW8rKBt1PerYOQH7tj2yk9hgNOzU"
ANT_MAP_STATS   = {
    "Usual"         : 1800755148,
    "Watched"       : 1294827541
}

TOUR_URL_ALIAS  = "https://docs.google.com/spreadsheets/d/10YBcZP_l5Tjf1MOiWeBlLg-ATuAWXgTPsj7bW79bU30/export?format=csv&gid=1934025140"
TOUR_KEY_STATS  = "1Fm6pMyXv7qhOQkLah4yX9HNow4WaDR4HJuAVMukQl34"
TOUR_MAP_STATS  = {
    "Usual"         : 0,
    "Watched"       : 2040874005,
    "Watched OP"    : 2122428774,
    "Watched ED"    : 1177334024,
    "Watched IN"    : 928352352,
    "Watched 2+8s"  : 41221104,
    "Watcehd 5s"    : 1525886733,
    "Random OP"     : 1093764794,
    "Random ED"     : 1863696842,
    "Random IN"     : 1919154942
}

EXCLUDED_TAGS = {
    "Female Protagonist",
    "Male Protagonist",
    "Primarily Female Cast",
    "Primarily Male Cast",
    "School",
    "Heterosexual",
    "Primarily Teen Cast",
    "Ensemble Cast"
}

def extract_year(vintage_str):
    years       = re.findall(r'\d{4}', str(vintage_str))
    year_val    = float(years[0])
    season_map  = {"winter": 0.00, "spring": 0.25, "summer": 0.50, "fall": 0.75}
    v_lower     = str(vintage_str).lower()
    decimal     = next((val for s, val in season_map.items() if s in v_lower), 0.0)

    return year_val + decimal

def format_year(val):
    year    = int(val)
    frac    = val - year
    season  = "Winter" if frac < 0.25 else "Spring" if frac < 0.50 else "Summer" if frac < 0.75 else "Fall"

    return f"{season} {year}"

def trim_whitespace(image_path):
    with Image.open(image_path) as img:
        img     = img.convert("RGB")
        bg      = Image.new(img.mode, img.size, "white")
        diff    = ImageChops.difference(img, bg)
        bbox    = diff.getbbox()

        if bbox:
            img = img.crop(bbox)
            img = ImageOps.expand(img, border = 10, fill = "white")

            img.save(image_path, compress_level = 1, optimize = False)