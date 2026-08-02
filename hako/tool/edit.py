import json
import tkinter as tk

from pathlib import Path
from tkinter import messagebox, ttk

class CustomSpinbox(tk.Frame):
    def __init__(self, parent, from_, to, initial_val = 1, command = None):
        super().__init__(parent, bg = "white")

        self.from_      = from_
        self.to         = to
        self.command    = command
        self.var        = tk.StringVar(value = str(initial_val))
        vcmd            = (self.register(self._validate_input), "%P")

        self.btn_dec, self.dec_arrow = self._create_arrow_button(column = 0, coords = (7,   9,  17, 9,  12, 15), delta = -1)
        self.btn_inc, self.inc_arrow = self._create_arrow_button(column = 2, coords = (12,  9,  7,  15, 17, 15), delta = 1)

        self.entry = tk.Entry(
            self,
            textvariable        = self.var,
            bg                  = "white",
            fg                  = "black",
            font                = ("Segoe UI", 10),
            justify             = "center",
            width               = 5,
            bd                  = 1,
            relief              = "solid",
            highlightthickness  = 0,
            validate            = "key",
            validatecommand     = vcmd,
        )

        self.entry.grid             (row = 0, column = 1, sticky = "ns")
        self.grid_columnconfigure   (1, minsize = 25)
        self._update_button_states  (initial_val)

    def _create_arrow_button(self, column, coords, delta):
        btn = tk.Canvas(
            self,
            width               = 25,
            height              = 25,
            bg                  = "black",
            highlightthickness  = 0,
            borderwidth         = 0,
            cursor              = "hand2",
        )

        arrow = btn.create_polygon(*coords, fill = "white")
        btn.grid(row = 0, column = column, sticky = "nsew")
        btn.bind("<Button-1>", lambda _: self._adjust_value(delta))

        return btn, arrow

    def _update_button_states(self, current_val):
        if current_val <= self.from_:
            self.btn_dec.configure(bg = "#f0f0f0", cursor = "arrow")
            self.btn_dec.itemconfig(self.dec_arrow, fill = "gray50")

        else:
            self.btn_dec.configure(bg = "black", cursor = "hand2")
            self.btn_dec.itemconfig(self.dec_arrow, fill = "white")

        if current_val >= self.to:
            self.btn_inc.configure(bg = "#f0f0f0", cursor = "arrow")
            self.btn_inc.itemconfig(self.inc_arrow, fill = "gray50")

        else:
            self.btn_inc.configure(bg = "black", cursor = "hand2")
            self.btn_inc.itemconfig(self.inc_arrow, fill = "white")

    def _validate_input(self, current_text):
        if current_text == "": return True

        if current_text.isdigit():
            val = int(current_text)

            if self.from_ <= val <= self.to:
                self._update_button_states(val)
                if self.command: self.after(10, lambda: self.command(val))
                return True

        return False

    def _adjust_value(self, delta):
        try                 : curr = int(self.var.get())
        except ValueError   : curr = self.from_

        if delta == -1  and curr <= self.from_  : return
        if delta == 1   and curr >= self.to     : return

        new_val = max(self.from_, min(self.to, curr + delta))
        self.var.set(str(new_val))
        self._update_button_states(new_val)
        if self.command: self.command(new_val)

    def get(self):
        try                 : return int(self.var.get())
        except ValueError   : return self.from_

    def set(self, val):
        bounded_val = max(self.from_, min(self.to, int(val)))
        self.var.set(str(bounded_val))
        self._update_button_states(bounded_val)
        if self.command: self.command(bounded_val)

class JSONEditor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("JSON Editor")
        self.configure(bg = "white")

        self.file_path      = None
        self.output_path    = None
        self.data           = None
        self.all_players    = set()
        self.player_vars    = {}
        self.list_vars      = {}
        self.manual_entries = []
        self.fill_color     = "#000000"
        self.style          = ttk.Style(self)

        self.style.configure(".", background = "white", foreground = "black")
        self.style.configure("TLabel", font = ("Segoe UI", 10))
        self.style.configure("TButton", font = ("Segoe UI", 10), padding = 5)

        if not self.find_and_load_file():
            self.destroy()
            return

        self.create_widgets     ()
        self.on_song_changed    (1)

    def find_and_load_file(self):
        current_dir = Path(__file__).parent.absolute()
        json_files  = list(current_dir.glob("*.json"))
        json_files  = [f for f in json_files if not f.name.startswith("edited-")]

        if not json_files:
            messagebox.showerror("Error", "No JSON files found in folder")
            return False

        if len(json_files) > 1:
            self.file_path = json_files[0]
            messagebox.showinfo("Multiple Files Found", f"Found multiple JSONs, automatically processing {self.file_path.name}")
        else: self.file_path = json_files[0]

        self.output_path = current_dir / f"edited-{self.file_path.stem}{self.file_path.suffix}"

        try:
            with open(self.file_path, "r", encoding = "utf-8") as f: self.data = json.load(f)
            self.extract_all_players()
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse {self.file_path.name}: {str(e)}")
            return False

    def extract_all_players(self):
        self.all_players.clear()

        if "songs" in self.data:
            for song in self.data["songs"]:
                for player in song.get("correctGuessPlayers", []): 
                    if      isinstance(player, dict) and "name" in player   : self.all_players.add(player["name"])
                    elif    isinstance(player, str)                         : self.all_players.add(player)

                for state in song.get("listStates", []):
                    if isinstance(state, dict) and "name" in state: self.all_players.add(state["name"])

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding = 20)
        main_frame.pack(fill = tk.BOTH, expand = True)

        ttk.Label(main_frame, text = f"Editing {self.file_path.name}",      font = ("Segoe UI", 10, "italic"), foreground = "gray50")   .pack(anchor = "w", pady = (0, 5))
        ttk.Label(main_frame, text = "Which song would you like to edit?",  font = ("Segoe UI", 10, "bold"))                            .pack(anchor = "w", pady = (0, 5))

        max_songs = len(self.data.get("songs", []))

        self.song_spinbox = CustomSpinbox(main_frame, from_ = 1, to = max_songs if max_songs > 0 else 1, initial_val = 1, command = self.on_song_changed)
        self.song_spinbox.pack(anchor = "w", pady = (0, 10))

        self.info_frame = ttk.Frame(main_frame)
        self.info_frame.pack(anchor = "w", pady = (0, 10))

        self.lbl_anime = ttk.Label(self.info_frame, text = "", font = ("Segoe UI", 10, "italic"), foreground = "gray50")
        self.lbl_anime.pack(anchor = "w")

        self.lbl_type = ttk.Label(self.info_frame, text = "", font = ("Segoe UI", 10, "italic"), foreground = "gray50")
        self.lbl_type.pack(anchor = "w")

        self.lbl_songname = ttk.Label(self.info_frame, text = "", font = ("Segoe UI", 10, "italic"), foreground = "gray50")
        self.lbl_songname.pack(anchor = "w")

        self.lbl_artist = ttk.Label(self.info_frame, text = "", font = ("Segoe UI", 10, "italic"), foreground = "gray50")
        self.lbl_artist.pack(anchor = "w")

        ttk.Label(main_frame, text = "Who got this song right?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (0, 5))
        self.player_frame = ttk.Frame(main_frame)

        self.player_frame.pack(side = "top", fill = "x", pady = (0, 10))
        ttk.Label(main_frame, text = "Who has this song on their list?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (0, 5))

        self.list_frame = ttk.Frame(main_frame)
        self.list_frame.pack(side = "top", fill = "x", pady = (0, 10))

        self.manual_container = ttk.Frame(main_frame)
        self.manual_container.pack(fill = "x", anchor = "w")

        self.build_manual_slots()
        ttk.Label(main_frame, text = "Would you like to delete the last song?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (0, 5))

        self.delete_var = tk.StringVar(value = "No")
        self.del_boxes  = {}

        for opt in ["No", "Yes"]:
            f_opt = ttk.Frame(main_frame)
            f_opt.pack(anchor = "w", pady = 1)

            is_sel      = self.delete_var.get() == opt
            bg_color    = self.fill_color if is_sel else "white"

            box = tk.Canvas(
                f_opt,
                width               = 10,
                height              = 10,
                bg                  = bg_color,
                highlightthickness  = 1,
                highlightbackground = "black",
            )

            box.pack(side = tk.LEFT, padx = (0, 5))

            self.del_boxes[opt] = box
            lbl                 = ttk.Label(f_opt, text = opt, font = ("Segoe UI", 10))

            lbl.pack(side = tk.LEFT)
            for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_del_opt(o))

        self.confirm_btn = ttk.Button(main_frame, text = "Confirm", command = self.on_confirm)
        self.confirm_btn.pack(side = tk.RIGHT)

    def build_manual_slots(self):
        for widget in self.manual_container.winfo_children(): widget.destroy()
        self.manual_entries.clear()
        found_count = len(self.all_players)

        if found_count < 8:
            slots_needed = 8 - found_count
            ttk.Label(self.manual_container, text = "Add new players: ", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (0, 5))

            for i in range(slots_needed):
                f_slot = ttk.Frame(self.manual_container)
                f_slot.pack(anchor = "w", fill = "x", pady = 5)
                ttk.Label(f_slot, text = f"Player {found_count + i + 1}: ", width = 10, font = ("Segoe UI", 10)).pack(side = tk.LEFT)

                ent = ttk.Entry(f_slot, width = 30)
                ent.pack(side = tk.LEFT, padx = 5)
                self.manual_entries.append(ent)

    def _select_del_opt(self, opt):
        self.delete_var.set(opt)
        for k, box in self.del_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")

    def toggle_custom_player(self, name, box, var_dict):
        new_val = not var_dict[name].get()
        var_dict[name].set(new_val)
        box.configure(bg = self.fill_color if new_val else "white")

    def on_song_changed(self, song_num):
        for widget in self.player_frame .winfo_children(): widget.destroy()
        for widget in self.list_frame   .winfo_children(): widget.destroy()

        if not self.data or "songs" not in self.data or not self.data["songs"]: return

        idx = song_num - 1
        if idx >= len(self.data["songs"]): return

        target_song     = self.data["songs"][idx]
        song_info       = target_song.get("songInfo",   {})
        anime_romaji    = song_info.get("animeNames",   {}).get("romaji", "Unknown")
        song_type       = song_info.get("type",         3)
        type_num        = song_info.get("typeNumber",   0)
        song_name       = song_info.get("songName",     "Unknown")
        artist_name     = song_info.get("artist",       "Unknown")

        self.lbl_anime      .configure(text = anime_romaji)
        self.lbl_songname   .configure(text = song_name)
        self.lbl_artist     .configure(text = artist_name)

        if      song_type == 1  : self.lbl_type.configure(text = f"Opening {type_num}")
        elif    song_type == 2  : self.lbl_type.configure(text = f"Ending {type_num}")
        else                    : self.lbl_type.configure(text = "Insert")

        correct_guess_raw   = target_song.get("correctGuessPlayers", [])
        correct_guessers    = set()

        for p in correct_guess_raw:
            if      isinstance(p, dict) and "name" in p : correct_guessers.add(p["name"])
            elif    isinstance(p, str)                  : correct_guessers.add(p)

        list_state_players  = set(state["name"] for state in target_song.get("listStates", []) if "name" in state)

        self.player_vars    .clear()
        self.list_vars      .clear()

        sorted_players  = sorted(list(self.all_players), key = str.lower)
        rows_per_col    = 8

        for i, name in enumerate(sorted_players):
            col                     = i //  rows_per_col
            row                     = i %   rows_per_col
            is_correct              = name in correct_guessers
            var                     = tk.BooleanVar(value = is_correct)
            self.player_vars[name]  = var
            item_frame              = ttk.Frame(self.player_frame)

            item_frame.grid(row = row, column = col, pady = 1, sticky = "w")
            initial_bg = self.fill_color if is_correct else "white"

            box = tk.Canvas(
                item_frame,
                width               = 10,
                height              = 10,
                bg                  = initial_bg,
                highlightthickness  = 1,
                highlightbackground = "black",
            )

            box.pack(side = tk.LEFT, padx = (0, 5))
            lbl = ttk.Label(item_frame, text = name, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)
            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, n = name, b = box: self.toggle_custom_player(n, b, self.player_vars))

        for i, name in enumerate(sorted_players):
            col                     = i // rows_per_col
            row                     = i % rows_per_col
            has_list                = name in list_state_players
            var                     = tk.BooleanVar(value = has_list)
            self.list_vars[name]    = var
            item_frame              = ttk.Frame(self.list_frame)

            item_frame.grid(row = row, column = col, pady = 1, sticky = "w")
            initial_bg = self.fill_color if has_list else "white"

            box = tk.Canvas(
                item_frame,
                width               = 10,
                height              = 10,
                bg                  = initial_bg,
                highlightthickness  = 1,
                highlightbackground = "black",
            )

            box.pack(side = tk.LEFT, padx = (0, 5))
            lbl = ttk.Label(item_frame, text = name, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)
            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, n = name, b = box: self.toggle_custom_player(n, b, self.list_vars))

    def on_confirm(self):
        if not self.data or "songs" not in self.data or not self.data["songs"]: return

        if self.delete_var.get() == "Yes":
            confirm = messagebox.askyesno("Deletion Confirmation", "Are you sure you want to delete the last song?")

            if confirm:
                self.data["songs"].pop()
                self.save_and_reload("Deleted the last song")
                return

        idx = self.song_spinbox.get() - 1
        if idx >= len(self.data["songs"]): return

        song            = self.data["songs"][idx]
        updated_correct = [name for name, var in self.player_vars.items() if var.get()]

        for ent in self.manual_entries:
            val = ent.get().strip()
            if val and val not in updated_correct: updated_correct.append(val)

        old_correct_count   = len(song.get("correctGuessPlayers", []))
        new_correct_count   = len(updated_correct)
        diff                = new_correct_count - old_correct_count

        song["correctGuessPlayers"] = updated_correct
        song["correctCount"]        = song.get("correctCount", 0) + diff

        if      diff > 0: song["wrongCount"] = max(0, song.get("wrongCount", 0) - diff)
        elif    diff < 0: song["wrongCount"] = song.get("wrongCount", 0) + abs(diff)

        current_list_states = song.get("listStates", [])
        existing_list_names = {state["name"] for state in current_list_states if "name" in state}
        new_list_states     = []

        for state in current_list_states:
            if "name" in state and self.list_vars.get(state["name"], tk.BooleanVar(value = True)).get(): new_list_states.append(state)

        for name, var in self.list_vars.items():
            if var.get() and name not in existing_list_names: new_list_states.append({
                "name"      : name,
                "status"    : 2,
                "score"     : None
            })

        song["listStates"] = new_list_states

        self.save_and_reload("Updated song records")

    def save_and_reload(self, success_message):
        try:
            with open(self.output_path, "w", encoding = "utf-8") as f: json.dump(self.data, f, ensure_ascii = False)
            messagebox.showinfo("Success", success_message)

            self.file_path = self.output_path
            with open(self.file_path, "r", encoding = "utf-8") as f: self.data = json.load(f)

            self.extract_all_players    ()
            self.build_manual_slots     ()

            max_songs = len(self.data.get("songs", []))

            self.song_spinbox.to = max_songs if max_songs > 0 else 1
            self.song_spinbox.set(min(self.song_spinbox.get(), max_songs))

            self._select_del_opt("No")

        except Exception as e: messagebox.showerror("Error", f"Failed to save changes:\n{str(e)}")

if __name__ == "__main__":
    app = JSONEditor()
    if app.winfo_exists(): app.mainloop()