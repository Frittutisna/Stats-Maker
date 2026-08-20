# How to Install

1. Delete the `hako` folder, as well as `hako_clean.py` and `hako_stats.py` from your directory if you have them
2. Download the latest release
3. Extract the downloaded file
4. Move the `hako` folder, `hako_clean.py`, `hako_stats.py` to your directory. If you are working with **Dry Stats** as well, make sure to put them on the same level as the `assets` folder
5. Run `hako_clean.py` to clean your directory

# How To Use
## Single Tour
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
4. Run `hako_stats.py`. You will be presented with the following prompt as an example:
![Configuration Prompt Example](https://files.catbox.moe/24lkpn.webp)
Change the answer to each question as you see fit:
    1. `Are there any new player?` detects players with round-number Elos that might need mid-tour guess adjustments
    2. Change `Do you want to fetch Challonge data as well?` to `Yes` and `Would you like to share the Stats site?` to `Yes, push this to Netlify ...` when running this at the end of tour
    3.  Change `Do you want to use Dry's script as well?` to `Yes, ...` to continue with running **Dry Stats** after **Hako Stats** ends
5. Post `hako_0_player.png` and `hako_0_extra.png` to `#tour-talk` for mid-tour stats, or copy-paste the link to the site from the command-line to `#export-stats` below **Dry Stats** outputs at the end of tour
## Multiple (Split) Tours
1. **Hako Stats** supports up to three tours simultaneously
2. Before running stats for a new batch of tours, run `hako_clean.py` to clean your directory
3. Copy-paste from `#tour-information` to each tour's respective `code.txt` in `hako/tour/[0-2]`. **Hako Stats** takes the following format as an example:
```
bofu (118.884) paperyoshi10 (98.474) justaweirdo (78.879) redrumyyy (72.540) | Total = 368.777 | Guesses = [5544]
kaededayo (114.178) torradinhas (94.353) serozero (91.831) rhummy (68.191) | Total = 368.553 | Guesses = [5553]
konomi (127.293) pile (110.347) foopypoopy (79.747) kazuyasouma1 (51.240) | Total = 368.627 | Guesses = [5551]
liljakatsuragi (103.454) deerparkboss (97.130) sabetin (95.981) whales6978 (71.819) | Total = 368.384 | Guesses = [5554]

Average: 368.585

https://challonge.com/dblow5ep
```
4. Move the JSON files for each tour to their `json` folder in `hako/tour/[0-2]`
5. Run `hako_stats.py`, then click `Confirm` if the prompt shown looks correct to you
![Selection Prompt Example](https://files.catbox.moe/3h2qoh.webp)
6. You will be presented with the following prompt for each tour as an example:
![Configuration Prompt Example](https://files.catbox.moe/24lkpn.webp)
Change the answer to each question as you see fit:
    1. `Are there any new player?` detects players with round-number Elos that might need mid-tour guess adjustments
    2. Change `Do you want to fetch Challonge data as well?` to `Yes` and `Would you like to share the Stats site?` to `Yes, push this to Netlify ...` when running this at the end of tour
    3.  Change `Do you want to use Dry's script as well?` to `Yes, ...` to continue with running **Dry Stats** after **Hako Stats** ends
7. Post `hako_[0-2]_player.png` and `hako_[0-2]_extra.png` to `#tour-talk` for mid-tour stats, or copy-paste the link to the site from the command-line to `#export-stats` below **Dry Stats** outputs at the end of tour