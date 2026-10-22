# Hako Stats v1.4.2

## Changelog
1. Fixed auto-update
2. Added version tag in Tour Configuration prompt

## Note
1. To use with **Dry Stats**, move to the same level as `assets`
2. From `v1.3.0`, `Confirm` changes to `Update` whenever there's a newer release

## How to Install
[Download](https://github.com/Frittutisna/Stats-Maker/releases/latest) or update to the latest release

## How To Use
1. Run `hako_clean.py`, then do the following:

* **Single Tour:** Copy-paste from `#tour-information` (example below), to `codes.txt` and move the JSON files to `jsons`
* **Multiple Tours:** Copy-paste from `#tour-information` (example below) to each tour's `code.txt` in `hako/tour/[0-2]/txt` and move the JSON files to each tour's `hako/tour/[0-2]/json`

```
bofu (118.884) paperyoshi10 (98.474) justaweirdo (78.879) redrumyyy (72.540) | Total = 368.777 | Guesses = [5544]
kaededayo (114.178) torradinhas (94.353) serozero (91.831) rhummy (68.191) | Total = 368.553 | Guesses = [5553]
konomi (127.293) pile (110.347) foopypoopy (79.747) kazuyasouma1 (51.240) | Total = 368.627 | Guesses = [5551]
liljakatsuragi (103.454) deerparkboss (97.130) sabetin (95.981) whales6978 (71.819) | Total = 368.384 | Guesses = [5554]

Average: 368.585

https://challonge.com/dblow5ep
```

2. Run `hako_stats.py`. For **Multiple Tours**, confirm the Tour Selection prompt (example below)

![Tour Selection Prompt Example](https://files.catbox.moe/3h2qoh.webp)

3. Modify accordingly, then confirm the Tour Configuration prompt (example below)

![Tour Configuration Prompt Example](https://files.catbox.moe/s3hvth.png)

4. Do the following:
    * **Mid-Tour:** Post `Player.png` and `Extra.png` from `hako-[0-2]` to `#tour-talk`
    * **Post-Tour:** Copy-paste `Site.zip` from `hako-[0-2]` to [Netlify](https://app.netlify.com/drop), then post the link in `#export-stats` as `Site: [Link]` after **Dry Stats** PNGs

## Tools
Copy-paste the relevant file(s) to `hako/tool` before running the script:
1. `edit.py` edits per-song JSON data
2. `stitch.py` stitches multiple JSONs from the same match into one
3. `swap.py` swaps old and new players from `code.txt`