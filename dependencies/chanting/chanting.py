import  requests
from    models      import  Anime, SongLink, Song, SongCategory

def decode_master_list(master_list):
    for master_anime in master_list['animeMap'].values():
        anime = Anime(master_anime)
        for master_song_links in master_anime['songLinks'].values():
            for master_song_link in master_song_links:
                song_link       = SongLink(master_song_link)
                song_link.song  = Song.from_master_list(master_list, master_song_link['songId'])
                song_link.anime = anime
                yield song_link

def main():
    r = requests.get("https://animemusicquiz.com/libraryMasterList")
    if not r.ok:
        print(f"Failed to request master list from AMQ: {r.status_code}")
        return

    master_list = r.json()
    song_links  = [song_link for song_link in decode_master_list(master_list)]
    results     = []

    for song_link in song_links:
        if song_link.song.category == SongCategory.CHANTING: results.append(song_link.ann_song_id)

    results.sort()
    with open("chanting.txt", 'w') as f: f.write('\n'.join(map(str, results)))

if __name__ == "__main__": main()