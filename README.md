# Hako Stats v1.5.0

## Changelog
1. Fixed auto-update
2. Added version tag in Tour Configuration prompt
3. Added `.gitignore` to public repo

## Note
1. To use with **Dry Stats**, move to the same level as `assets`
2. From `v1.3.0`, `Confirm` changes to `Update` whenever there's a newer release

## How to Install
* **Clear your directory**, then download and extract [Hako.zip](https://github.com/Frittutisna/Stats-Maker/releases/download/v1.5.0/Hako.zip)
* Download and extract [Ant.zip](https://github.com/Frittutisna/Stats-Maker/releases/download/v1.5.0/Ant.zip) to run stats for **Ant Tours**
* Download and extract [Sample.zip](https://github.com/Frittutisna/Stats-Maker/releases/download/v1.5.0/Sample.zip), then follow from **Step 2** of the [How to Use](#how-to-use-ant-in-brackets) section to try out **Hako Stats**

## How To Use (Ant in Brackets)
1. Run `hako_clean.py`, then do the following:

* **Single Tour:** Copy-paste from `#tour-information` (`#stuff`), to `codes.txt` and move the JSON files to `jsons`
* **Multiple Tours:** Copy-paste from `#tour-information` (`#stuff`) to each tour's `code.txt` in `hako/tour/[0-2]/txt` and move the JSON files to each tour's `hako/tour/[0-2]/json`

```
zental (112.564) floth (95.265) oolli (86.752) furlain (86.562) | Total = 381.143 | Guesses = [5555]
lika (113.654) mangoss (101.347) serozero (95.132) hubcioskubcio (71.401) | Total = 381.534 | Guesses = [5553]
emdiel (111.710) keysha (94.076) yellowrose (93.009) zomg (82.119) | Total = 380.914 | Guesses = [5554]
civiltwlight (128.325) i_dont_like_loli (91.408) 2skr (89.121) hakohoka (72.484) | Total = 381.338 | Guesses = [5554]

Average: 381.2322

https://challonge.com/amqtour2610222130
```

2. Run `hako_stats.py`. For **Multiple Tours**, confirm the Tour Selection prompt (example below)

![Tour Selection Prompt Example](https://files.catbox.moe/egq6gb.png)

3. Modify accordingly, then confirm the Tour Configuration prompt (example below)

![Tour Configuration Prompt Example](https://files.catbox.moe/x4gui0.png)

4. Do the following:
    * **Mid-Tour:** Post `Player.png` and `Extra.png` from `hako-[0-2]` to `#tour-talk` (`#yappa`)
    * **Post-Tour:** Copy-paste `Site.zip` from `hako-[0-2]` to [Netlify](https://app.netlify.com/drop), then post the link in `#export-stats` as `Site: [Link]` after **Dry Stats** PNGs (post the PNGs and link (as `Site: [Link]`) to `#export-stats-for-the-curious`)

## Tools
Copy-paste the relevant file(s) to `hako/tool` before running the script:
1. `edit.py` edits per-song JSON data
2. `stitch.py` stitches multiple JSONs from the same match into one
3. `swap.py` swaps old and new players from `code.txt`