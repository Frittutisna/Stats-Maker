import datetime, json, re
import tkinter as tk

from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox

class JSONEditor(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("JSON Editor")
        self.configure(bg = "white")

        self.file_path      = None
        self.output_path    = None
        self.data           = None
        self.players        = []
        self.row_states     = []
        self.hit_regions    = []

        self.font_regular   = tkfont.Font(family="Segoe UI", size = 10)
        self.font_bold      = tkfont.Font(family="Segoe UI", size = 10, weight = "bold")
        self.font_btn_sm    = tkfont.Font(family="Segoe UI", size = 8,  weight = "bold")

        if not self.find_and_load_file():
            self.destroy()
            return

        self._init_row_states   ()
        self.create_widgets     ()
        self.render_table       ()

    def find_and_load_file(self):
        current_dir = Path(__file__).parent.absolute()
        json_files  = [f for f in current_dir.glob("*.json") if not f.name.startswith("edited-")]

        if not json_files:
            messagebox.showerror("Error", "No JSON files found in folder")
            return False

        def get_start_time(file_path):
            try:
                with open(file_path, "r", encoding = "utf-8") as f: raw_time = json.load(f).get("startTime", "")
                return                  datetime.strptime       (raw_time.replace("GMT", ""), "%a %b %d %Y %H:%M:%S %z")
            except Exception: return    datetime.max.replace    (tzinfo = datetime.now().astimezone().tzinfo)

        if len(json_files) > 1:
            self.file_path = min(json_files, key = get_start_time)
            messagebox.showinfo("Multiple Files Found", f"Found multiple JSONs, automatically processing oldest: {self.file_path.name}")
        else:
            self.file_path = json_files[0]

        self.output_path = current_dir / f"edited-{self.file_path.stem}{self.file_path.suffix}"

        try:
            with open(self.file_path, "r", encoding = "utf-8") as f: self.data = json.load(f)
            self.extract_all_players()
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse {self.file_path.name}: {str(e)}")
            return False

    def extract_all_players(self):
        found = set()

        for song in self.data.get("songs", []):
            for player in song.get("correctGuessPlayers", []):
                if      isinstance(player, dict) and "name" in player   : found.add(player["name"])
                elif    isinstance(player, str)                         : found.add(player)

            for state in song.get("listStates", []):
                if isinstance(state, dict) and "name" in state: found.add(state["name"])

        self.players = sorted(found, key = str.lower)

    def _init_row_states(self):
        self.row_states.clear()

        for song in self.data.get("songs", []):
            correct_set = set()

            for p in song.get("correctGuessPlayers", []):
                if      isinstance(p, dict) and "name" in p : correct_set.add(p["name"])
                elif    isinstance(p, str)                  : correct_set.add(p)

            self.row_states.append({"song": song, "correct": {p: (p in correct_set) for p in self.players}})

    def create_widgets(self):
        self.top_bar = tk.Frame(self, bg = "white", padx = 15, pady = 5)
        self.top_bar.pack(fill = tk.X)

        self.file_lbl = tk.Label(self.top_bar, text = f"Editing {self.file_path.name}", bg = "white", fg = "black", font = self.font_regular)
        self.file_lbl.pack(side = tk.LEFT)

        action_frame = tk.Frame(self.top_bar, bg = "white")
        action_frame.pack(side = tk.RIGHT)

        control_h = 20

        self.add_player_frame = tk.Frame(action_frame, bg = "white")
        if len(self.players) < 8: self.add_player_frame.pack(side = tk.LEFT)

        tk.Label(self.add_player_frame, text = "New Player:", bg = "white", fg = "black", font = self.font_regular).pack(side = tk.LEFT, padx = (0, 5))

        entry_wrap = tk.Frame(self.add_player_frame, width = 100, height = control_h, bg = "white", highlightthickness = 1, highlightbackground = "black")
        entry_wrap.pack_propagate(False)
        entry_wrap.pack(side = tk.LEFT, padx = (0, 5))

        self.new_player_var     = tk.StringVar()
        self.new_player_entry   = tk.Entry(
            entry_wrap,
            textvariable        = self.new_player_var,
            bg                  = "white",
            fg                  = "black",
            insertbackground    = "black",
            selectbackground    = "black",
            selectforeground    = "white",
            font                = self.font_regular,
            bd                  = 0,
            highlightthickness = 0
        )

        self.new_player_entry.pack(fill = tk.BOTH, expand = True, padx = 5)
        self.new_player_entry.bind("<Return>", lambda _: self.add_player())

        self._make_fixed_button(self.add_player_frame, text = "Add", command = self.add_player, width = 50, height = control_h, inverted = False).pack(side = tk.LEFT, padx = (0, 5))

        self.save_btn = self._make_fixed_button(action_frame, text = "Save", command = self.save_changes, width = 50, height = control_h, inverted = True)
        self.save_btn.pack(side = tk.LEFT)

        container = tk.Frame(self, bg = "white")
        container.pack(fill = tk.BOTH, expand = True, padx = 15, pady = (0, 15))

        self.canvas = tk.Canvas(container, bg = "white", highlightthickness = 0, bd = 0, yscrollincrement = 20, xscrollincrement = 20)
        self.canvas.pack(side = tk.LEFT, fill = tk.BOTH, expand = True)

        self.canvas.bind("<Button-1>",  self._on_canvas_click)
        self.canvas.bind("<Motion>",    self._on_canvas_motion)
        self.canvas.bind("<Button-4>",  lambda _: self.canvas.yview_scroll(-1,  "units"))
        self.canvas.bind("<Button-5>",  lambda _: self.canvas.yview_scroll(1,   "units"))

    def _make_fixed_button(self, parent, text, command, width, height, inverted = False):
        bg = "black" if inverted else "white"
        fg = "white" if inverted else "black"

        wrap = tk.Frame(parent, width = width, height = height, bg = bg, highlightthickness = 1, highlightbackground = "black")
        wrap.pack_propagate(False)

        btn = tk.Button(
            wrap,
            text                = text,
            command             = command,
            bg                  = bg,
            fg                  = fg,
            activebackground    = fg,
            activeforeground    = bg,
            font                = self.font_btn_sm,
            bd                  = 0,
            relief              = "flat",
            highlightthickness  = 0,
            cursor              = "hand2",
        )

        btn.pack(fill = tk.BOTH, expand = True)
        return wrap

    def add_player(self):
        name = self.new_player_var.get().strip()
        if not name: return

        if name in self.players:
            messagebox.showinfo("Duplicate Player", f"'{name}' is already in the table")
            return

        self.players.append(name)
        self.players.sort(key = str.lower)

        for row in self.row_states: row["correct"].setdefault(name, False)

        self.new_player_var.set ("")
        self.render_table       ()

    def _format_song_type(self, song_info):
        song_type   = song_info.get("type",         3)
        type_num    = song_info.get("typeNumber",   0)

        if song_type == 1: return   f"OP{type_num}" if type_num else "OP"
        if song_type == 2: return   f"ED{type_num}" if type_num else "ED"
        return                      "IN"

    def _draw_delete_button(self, cx, cy):
        bw, bh = 50,            20
        x0, y0 = cx - bw / 2,   cy - bh / 2
        x1, y1 = cx + bw / 2,   cy + bh / 2

        self.canvas.create_text(cx, cy, text = "Delete", fill = "black", font = self.font_btn_sm)
        return x0, y0, x1, y1

    def render_table(self):
        self.canvas.delete      ("all")
        self.hit_regions.clear  ()

        num_players     = len(self.players)
        sub_cols        = max(1, num_players)
        pad_x, pad_y    = 4, 2
        line_h          = self.font_regular.metrics("linespace")
        prepared_rows   = []

        for row_data in self.row_states:
            song_info       = row_data["song"].get("songInfo", {})
            anime_lines     = [song_info.get("animeNames", {}).get("romaji", "Unknown") or "Unknown"]
            song_lines      = [song_info.get("songName",   "Unknown") or "Unknown"]
            raw_artists     = re.split(r"\s*[&,]\s*", song_info.get("artist", "Unknown"))
            artist_lines    = [part for part in raw_artists if part] or ["Unknown"]
            type_str        = self._format_song_type(song_info)

            prepared_rows.append((row_data, anime_lines, song_lines, artist_lines, type_str))

        def calc_col_width(header_title, rows_lines):
            max_w = self.font_bold.measure(header_title) + 15

            for lines in rows_lines:
                for line in lines: max_w = max(max_w, self.font_regular.measure(line) + pad_x * 2 + 4)

            return max_w

        w_num       = 20
        w_anime     = calc_col_width("Anime",   [r[1] for r in prepared_rows])
        w_type      = 40
        w_song      = calc_col_width("Song",    [r[2] for r in prepared_rows])
        w_artist    = calc_col_width("Artist",  [r[3] for r in prepared_rows])
        w_del       = 60

        player_widths   = [self.font_bold.measure(p) + 15 for p in self.players]
        w_players       = sum(player_widths)
        x_num           = 0
        x_anime         = x_num     + w_num
        x_type          = x_anime   + w_anime
        x_song          = x_type    + w_type
        x_artist        = x_song    + w_song
        x_players       = x_artist  + w_artist
        x_del           = x_players + w_players
        total_w         = x_del     + w_del
        hdr_h           = 20

        self.canvas.create_rectangle(0, 0, total_w, hdr_h, fill = "black", outline = "black")

        fixed_headers = [
            (x_num,     w_num,      "#"),
            (x_anime,   w_anime,    "Anime"),
            (x_type,    w_type,     "Type"),
            (x_song,    w_song,     "Song"),
            (x_artist,  w_artist,   "Artist"),
            (x_del,     w_del,      "Delete"),
        ]

        for x, w, title in fixed_headers                                            : self.canvas.create_text(x + w / 2, hdr_h / 2, text = title, fill = "white", font = self.font_bold)
        for vx          in (x_anime, x_type, x_song, x_artist, x_players, x_del)    : self.canvas.create_line(vx, 0, vx, hdr_h, fill = "white")

        cur_x = x_players

        for i in range(sub_cols):
            pw      = player_widths [i]
            ptitle  = self.players  [i] if num_players > 0 else "-"

            if i > 0: self.canvas.create_line(cur_x, 0, cur_x, hdr_h, fill = "white")
            self.canvas.create_text(cur_x + pw / 2, hdr_h / 2, text = ptitle, fill = "white", font = self.font_bold)

            cur_x += pw

        y = hdr_h

        for idx, (row_data, anime_lines, song_lines, artist_lines, type_str) in enumerate(prepared_rows):
            max_lines   = max(len(anime_lines), len(song_lines), len(artist_lines), 1)
            row_h       = max(10, max_lines * line_h + pad_y * 2)

            y1 = y + row_h
            cy = y + row_h / 2

            self.canvas.create_rectangle(0, y, total_w, y1, fill = "white", outline = "black")
            for vx in (x_num, x_anime, x_type, x_song, x_artist, x_players, x_del, total_w): self.canvas.create_line(vx, y, vx, y1, fill = "black")

            self.canvas.create_text(x_num       + w_num     / 2,    cy, text = str(idx + 1),            fill = "black", font = self.font_regular)
            self.canvas.create_text(x_anime     + pad_x,            cy, text = "\n".join(anime_lines),  fill = "black", font = self.font_regular, anchor = "w")
            self.canvas.create_text(x_type      + w_type    / 2,    cy, text = type_str,                fill = "black", font = self.font_regular)
            self.canvas.create_text(x_song      + pad_x,            cy, text = "\n".join(song_lines),   fill = "black", font = self.font_regular, anchor = "w")
            self.canvas.create_text(x_artist    + pad_x,            cy, text = "\n".join(artist_lines), fill = "black", font = self.font_regular, anchor = "w")

            cur_x = x_players

            for i in range(sub_cols):
                pw = player_widths[i]
                if i > 0: self.canvas.create_line(cur_x, y, cur_x, y1, fill = "black")

                if num_players > 0:
                    player  = self.players[i]
                    checked = row_data["correct"][player]
                    tag     = f"cb_{idx}_{i}"

                    self.canvas.create_rectangle    (cur_x, y, cur_x + pw, y1, fill = "black" if checked else "white", outline = "black", tags = (tag, ))
                    self.hit_regions.append         ((cur_x, y, cur_x + pw, y1, ("check", tag), idx, player))
                else: self.canvas.create_text       (cur_x + pw / 2, cy, text = "-", fill = "black", font = self.font_regular)

                cur_x += pw

            self._draw_delete_button    (x_del + w_del / 2, cy)
            self.hit_regions.append     ((x_del, y, total_w, y1, ("delete", None), idx, None))

            y = y1

        self.canvas.create_rectangle    (0, 0, total_w, hdr_h, outline = "black")
        self.canvas.create_rectangle    (0, hdr_h - 1, total_w, hdr_h + 1, fill = "white")
        self.canvas.configure           (scrollregion = (0, 0, total_w, max(y, 1)))
        self.canvas.xview_moveto        (0)
        self.canvas.yview_moveto        (0)
        self.update_idletasks           ()

        pad_w = 30 + 1
        pad_h = 20 + 1 + self.top_bar.winfo_reqheight()

        req_w = pad_w + total_w
        req_h = pad_h + max(y, 1)

        max_w = self.winfo_screenwidth  () - 80
        max_h = self.winfo_screenheight () - 100

        self.geometry(f"{min(req_w, max_w)}x{min(req_h, max_h)}")

    def _find_hit(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)

        for region in self.hit_regions:
            x0, y0, x1, y1, action, row_idx, player = region
            if x0 <= cx <= x1 and y0 <= cy <= y1: return action, row_idx, player

        return None

    def _on_canvas_motion(self, event):
        hit = self._find_hit(event)
        self.canvas.configure(cursor = "hand2" if hit else "")

    def _on_canvas_click(self, event):
        hit = self._find_hit(event)
        if not hit: return

        (kind, tag), row_idx, player = hit

        if kind == "check":
            new_val = not self.row_states[row_idx]["correct"][player]
            self.row_states[row_idx]["correct"][player] = new_val
            self.canvas.itemconfigure(tag, fill = "black" if new_val else "white")
        elif kind == "delete": self.delete_song(row_idx)

    def delete_song(self, row_idx):
        if not (0 <= row_idx < len(self.row_states)): return

        song_info   = self.row_states[row_idx]["song"].get("songInfo", {})
        song_name   = song_info.get("songName", "Unknown")
        artist      = song_info.get("artist", "Unknown")
        anime_name  = song_info.get("animeNames", {}).get("romaji", "Unknown")

        if not messagebox.askyesno("Delete Song", f"Delete {song_name} by {artist} in {anime_name}?"): return

        self.row_states.pop (row_idx)
        self.render_table   ()

    def save_changes(self):
        if not self.data: return
        updated_songs = []

        for idx, row_data in enumerate(self.row_states):
            song = row_data["song"]
            if "songNumber" in song: song["songNumber"] = idx + 1

            updated_correct     = [p for p in self.players if row_data["correct"].get(p, False)]
            old_correct_count   = len(song.get("correctGuessPlayers", []))
            new_correct_count   = len(updated_correct)
            diff                = new_correct_count - old_correct_count

            song["correctGuessPlayers"] = updated_correct
            song["correctCount"]        = song.get("correctCount", 0) + diff

            if      diff > 0: song["wrongCount"] = max(0, song  .get("wrongCount", 0) - diff)
            elif    diff < 0: song["wrongCount"] = song         .get("wrongCount", 0) + abs(diff)

            updated_songs.append(song)

        self.data["songs"] = updated_songs

        try:
            with open(self.output_path, "w", encoding = "utf-8") as f: json.dump(self.data, f, ensure_ascii = False)
            messagebox.showinfo("Success", f"Saved changes to {self.output_path.name}")

            self.file_path = self.output_path
            self.file_lbl.configure(text = f"Editing {self.file_path.name}")

            with open(self.file_path, "r", encoding = "utf-8") as f: self.data = json.load(f)

            self.extract_all_players    ()
            self._init_row_states       ()
            self.render_table           ()
        except Exception as e: messagebox.showerror("Error", f"Failed to save changes: {str(e)}")

if __name__ == "__main__":
    app = JSONEditor()
    if app.winfo_exists(): app.mainloop()