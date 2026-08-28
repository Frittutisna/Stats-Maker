# Hako Stats v1.2.1

## Changelog

1. Fixed alias assignment
2. Fixed 0/0 data
3. Moved from `html2image` to `playwright`

## How to Install

1. Sign up for a [Netlify](https://app.netlify.com/signup) account
2. Delete any folder and file with `hako` in their name from your directory
3. Download the latest release, then extract the downloaded file and move them to your directory. If you are working with **Dry Stats** as well, put them on the same level as the `assets` folder
4. Run `hako_clean.py` to clean your directory

## How To Use

### Single Tour

1. Before running stats for a new tour, run `hako_clean.py` to clean your directory
2. Copy-paste from `#tour-information` to `codes.txt`. **Hako Stats** takes the following format as an example:

```
bofu (118.884) paperyoshi10 (98.474) justaweirdo (78.879) redrumyyy (72.540) | Total = 368.777 | Guesses = [5544]
kaededayo (114.178) torradinhas (94.353) serozero (91.831) rhummy (68.191) | Total = 368.553 | Guesses = [5553]
konomi (127.293) pile (110.347) foopypoopy (79.747) kazuyasouma1 (51.240) | Total = 368.627 | Guesses = [5551]
liljakatsuragi (103.454) deerparkboss (97.130) sabetin (95.981) whales6978 (71.819) | Total = 368.384 | Guesses = [5554]

Average: 368.585

https://challonge.com/dblow5ep
```

3. Move the JSON files to the `jsons` folder
4. Run `hako_stats.py`. You will be presented with the following prompt as an example. While you can usually click `Confirm` without changing anything, feel free to change the answer to each question as you see fit

![Configuration Prompt Example](https://files.catbox.moe/s3hvth.png)

5. For mid-tour stats, post `Player.png` and `Extra.png` from the `hako_0` folder to `#tour-talk`
6. At the end of tour:
    1. Run `hako_stats.py` again, but change `Do you want to fetch Challonge data?` and `Do you want to use Dry's script?` to `Yes`
    2. Drag-and-drop `Site.zip` from the `hako_0` folder to [Netlify](https://app.netlify.com/drop)
    3. Post the Netlify link to `#export-stats` below **Dry Stats** outputs

### Multiple (Split) Tours

1. **Hako Stats** supports up to three tours simultaneously. Before running stats for a new batch of tours, run `hako_clean.py` to clean your directory
2. Copy-paste from `#tour-information` to each tour's respective `code.txt` in `hako/tour/[0-2]`. **Hako Stats** takes the following format as an example:

```
bofu (118.884) paperyoshi10 (98.474) justaweirdo (78.879) redrumyyy (72.540) | Total = 368.777 | Guesses = [5544]
kaededayo (114.178) torradinhas (94.353) serozero (91.831) rhummy (68.191) | Total = 368.553 | Guesses = [5553]
konomi (127.293) pile (110.347) foopypoopy (79.747) kazuyasouma1 (51.240) | Total = 368.627 | Guesses = [5551]
liljakatsuragi (103.454) deerparkboss (97.130) sabetin (95.981) whales6978 (71.819) | Total = 368.384 | Guesses = [5554]

Average: 368.585

https://challonge.com/dblow5ep
```

3. Move the JSON files for each tour to their `json` folder in `hako/tour/[0-2]`
4. Run `hako_stats.py`, then click `Confirm` if the prompt shown looks correct to you

![Selection Prompt Example](https://files.catbox.moe/3h2qoh.webp)

5. You will be presented with the following prompt for each tour as an example. While you can usually click `Confirm` without changing anything, feel free to change the answer to each question as you see fit

![Configuration Prompt Example](https://files.catbox.moe/s3hvth.png)

6. For mid-tour stats, post `Player.png` and `Extra.png` from their respective `hako_[0-2]` folder to `#tour-talk`
7. At the end of tour:
    1. Run `hako_stats.py` again, but change `Do you want to fetch Challonge data?` and `Do you want to use Dry's script?` to `Yes`
    2. Drag-and-drop `Site.zip` from their respective `hako_[0-2]` folder to [Netlify](https://app.netlify.com/drop)
    3. Post the Netlify link to `#export-stats` below **Dry Stats** outputs

## Tools

You will find `edit.py`, `stitch.py`, and `swap.py` in the `hako/tool` folder. You can use this to edit the JSON, stitch multiple JSONs from the same round into one JSON, and swap players within `codes.txt` respectively. Make sure to copy-paste the relevant JSON(s) or TXT file to the `hako/tool` folder before running the script.