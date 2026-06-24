import tkinter as tk

from help.config    import DIR_TOURS, DIR_JSONS, FILE_CODES
from pathlib        import Path
from tkinter        import ttk

class CustomSpinbox(tk.Frame):
    def __init__(self, parent, from_, to, initial_val = 1, state = "normal"):
        super().__init__(parent, bg = "white")

        self.from_  = from_
        self.to     = to
        self._state = state
        self.var    = tk.StringVar(value=str(initial_val))
        vcmd        = (self.register(self._validate_input), '%P')
        btn_width   = 25
        btn_height  = 25

        self.btn_dec = tk.Canvas(self, width = btn_width, height = btn_height, bg = "black", highlightthickness = 0, borderwidth = 0, cursor = "hand2")
        self.btn_dec.create_polygon(7, 9, 17, 9, 12, 15, fill = "white")
        self.btn_dec.grid(row = 0, column = 0, sticky = "nsew")
        self.btn_dec.bind("<Button-1>", lambda _: self._adjust_value(-1))

        self.entry = tk.Entry(
            self, 
            textvariable        = self.var, 
            bg                  = "white", 
            fg                  = "black", 
            font                = ("Segoe UI", 10), 
            justify             = "center", 
            width               = 4,
            bd                  = 1, 
            relief              = "solid", 
            highlightthickness  = 0, 
            validate            = "key", 
            validatecommand     = vcmd
        )

        self.entry.grid(row = 0, column = 1, padx = 0, sticky = "ns")
        self.grid_columnconfigure(1, minsize = btn_width)

        self.btn_inc = tk.Canvas(self, width = btn_width, height = btn_height, bg = "black", highlightthickness = 0, borderwidth = 0, cursor = "hand2")
        self.btn_inc.create_polygon(12, 9, 7, 15, 17, 15, fill = "white")
        self.btn_inc.grid(row = 0, column = 2, sticky = "nsew")
        self.btn_inc.bind("<Button-1>", lambda _: self._adjust_value(1))

        if state == "disabled": self.configure_state("disabled")

    def _validate_input(self, current_text):
        if current_text == ""       : return True
        if current_text.isdigit()   : return self.from_ <= int(current_text) <= self.to

        return False

    def _adjust_value(self, delta):
        if self._state == "disabled": return

        try                 : curr = int(self.var.get())
        except ValueError   : curr = self.from_

        new_val = max(self.from_, min(self.to, curr + delta))
        self.var.set(str(new_val))

    def get(self):
        try                 : return int(self.var.get())
        except ValueError   : return self.from_

    def set(self, val): self.var.set(str(max(self.from_, min(self.to, int(val)))))

    def configure_state(self, state):
        self._state = state

        if state == "disabled":
            self.entry      .configure  (state = "disabled", disabledbackground = "gray75", disabledforeground = "gray50")
            self.btn_dec    .configure  (bg = "gray75", cursor = "")
            self.btn_inc    .configure  (bg = "gray75", cursor = "")
            self.btn_dec    .itemconfig (1, fill = "gray50")
            self.btn_inc    .itemconfig (1, fill = "gray50")

        else:
            self.entry      .configure(state="normal")
            self.btn_dec    .configure(bg = "black", cursor = "hand2")
            self.btn_inc    .configure(bg = "black", cursor = "hand2")
            self.btn_dec    .itemconfig(1, fill = "white")
            self.btn_inc    .itemconfig(1, fill = "white")

class UnifiedDialog(tk.Toplevel):
    def __init__(self, parent, title, prompt):
        super().__init__(parent)

        self.title(title)
        self.result = None

        if parent is not None   : self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        else                    : self.geometry("+100+100")

        main_frame = ttk.Frame(self, padding = 8)
        main_frame.pack(fill = tk.BOTH, expand = True)

        if prompt: ttk.Label(main_frame, text = prompt, font = ("Segoe UI", 10)).pack(anchor = "w")

        self.container = ttk.Frame(main_frame)
        self.container.pack(fill = tk.BOTH, expand = True)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill = tk.X)

        self.confirm_btn = ttk.Button(btn_frame, text = "Confirm", command = self.on_confirm)
        self.confirm_btn.pack(side = tk.RIGHT)

        self.bind("<Return>", lambda: self.on_confirm())

    def on_confirm(self): self.destroy()

class TourSelectionDialog(UnifiedDialog):
    def __init__(self, parent, tour_ids):
        super().__init__(parent, "Tour Selection", "Which tours should be processed?")

        self.selected_tours = []
        self.vars           = {}
        self.fill_color     = "#000000"

        script_dir  = Path(__file__).parent.parent.absolute()
        states      = {}

        for tid in tour_ids:
            t_path          = script_dir / DIR_TOURS / str(tid)
            json_dir        = t_path / DIR_JSONS
            codes_file      = t_path / FILE_CODES
            is_recommended  = False

            if codes_file.exists() and json_dir.exists():
                codes_size = codes_file.stat().st_size
                json_count = len(list(json_dir.glob("*.json")))

                if codes_size > 0 and json_count > 1: is_recommended = True

            states[tid] = is_recommended

        for tid in tour_ids:
            is_active       = states[tid]
            var             = tk.BooleanVar(value = is_active)
            self.vars[tid]  = var
            item_frame      = ttk.Frame(self.container)

            item_frame.pack(anchor = "w")
            initial_bg = self.fill_color if is_active else "white"

            box = tk.Canvas(item_frame, width = 10, height = 10, bg = initial_bg, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 4))

            lbl = ttk.Label(item_frame, text = f"Tour {tid}", font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)

            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, t = tid, b = box: self.toggle_custom(t, b))

        self.protocol       ("WM_DELETE_WINDOW", lambda: [setattr(self, 'selected_tours', []), self.destroy()])
        self.grab_set       ()
        self.wait_window    ()

    def toggle_custom(self, tid, box):
        new_val = not self.vars[tid].get()
        self.vars[tid].set(new_val)
        color = self.fill_color if new_val else "white"
        box.configure(bg = color)

    def on_confirm(self):
        self.selected_tours = [tid for tid, var in self.vars.items() if var.get()]
        super().on_confirm()

class TourMetadataDialog(UnifiedDialog):
    def __init__(self, parent, tour_id, init_label, default_th, baseline_initial, active_players, elo_map = None, sub_candidates = None, original_players_list = None, tour_dir = None):
        super().__init__(parent, f"Tour {tour_id} Configuration", "")

        self.fill_color = "#000000"
        ttk.Label(self.container, text = "What tour is this?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w")

        self.lbl_var    = tk.StringVar(value = init_label if init_label in ["Watched", "Usual"] else "Others")
        self.lbl_boxes  = {}

        for opt in ["Watched", "Usual", "Others"]:
            f_opt = ttk.Frame(self.container)
            f_opt.pack(anchor = "w")

            is_sel      = (self.lbl_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"
            box         = tk.Canvas(f_opt, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")

            box.pack(side = tk.LEFT, padx = (0, 4))
            self.lbl_boxes[opt] = box

            if opt == "Others":
                lbl = ttk.Label(f_opt, text = "Others:", font = ("Segoe UI", 10))
                lbl.pack(side = tk.LEFT)

                self.lbl_entry = ttk.Entry(f_opt, width = 20)
                self.lbl_entry.insert(0, init_label)
                self.lbl_entry.pack(side = tk.LEFT, padx = (4, 0))

                for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_lbl_opt(o))

            else:
                lbl = ttk.Label(f_opt, text = opt, font = ("Segoe UI", 10))
                lbl.pack(side = tk.LEFT)

                for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_lbl_opt(o))

        self._update_lbl_state()
        ttk.Label(self.container, text = "What are the comma-separated guess rate threshold values?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w")

        self.th_var     = tk.StringVar(value = "default")
        self.th_boxes   = {}

        f_th1 = ttk.Frame(self.container)
        f_th1.pack(anchor = "w")

        box_th1 = tk.Canvas(f_th1, width = 10, height = 10, bg = self.fill_color, highlightthickness = 1, highlightbackground = "black")
        box_th1.pack(side = tk.LEFT, padx = (0, 4))
        self.th_boxes["default"] = box_th1

        lbl_th1 = ttk.Label(f_th1, text = "Use the default threshold values", font = ("Segoe UI", 10))
        lbl_th1.pack(side = tk.LEFT)

        for w in (box_th1, lbl_th1): w.bind("<Button-1>", lambda _: self._select_th_opt("default"))

        f_th2 = ttk.Frame(self.container)
        f_th2.pack(anchor = "w")

        box_th2 = tk.Canvas(f_th2, width = 10, height = 10, bg = "white", highlightthickness = 1, highlightbackground = "black")
        box_th2.pack(side = tk.LEFT, padx = (0, 4))

        self.th_boxes["custom"] = box_th2

        lbl_th2 = ttk.Label(f_th2, text = "Use custom threshold values:", font = ("Segoe UI", 10))
        lbl_th2.pack(side = tk.LEFT)

        self.th_entry = ttk.Entry(f_th2, width = 25)
        self.th_entry.insert(0, default_th)
        self.th_entry.pack(side = tk.LEFT, padx = (4, 0))

        for w in (box_th2, lbl_th2): w.bind("<Button-1>", lambda _: self._select_th_opt("custom"))
        self._update_th_state()
        ttk.Label(self.container, text = "How many rounds have elapsed?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (6, 4))

        self.spin = CustomSpinbox(self.container, from_ = 1, to = 6, initial_val = baseline_initial)
        self.spin.pack(anchor = "w")

        self.sub_vars = {}
        self.tour_dir = tour_dir

        if sub_candidates and original_players_list:
            sorted_subs         = sorted(list(sub_candidates),          key = str.lower)
            sorted_originals    = sorted(list(original_players_list),   key = str.lower)
            saved_subs_map      = {}

            if tour_dir and (tour_dir / "subs.txt").exists():
                with open(tour_dir / "subs.txt", "r", encoding = "utf-8") as f:
                    for line in f:
                        if "," in line:
                            s_name, o_name                          = line.strip().split(",", 1)
                            saved_subs_map[s_name.strip().lower()]  = o_name.strip()

            for sub_name in sorted_subs:
                f_sub = ttk.Frame(self.container)
                f_sub.pack(fill = tk.X, anchor = "w")

                lbl = ttk.Label(f_sub, text = f"Who is {sub_name} subbing for?", font = ("Segoe UI", 10, "bold"))
                lbl.pack(anchor = "w")

                choice_var = tk.StringVar()
                saved_orig = saved_subs_map.get(sub_name.lower())

                if saved_orig and any(p.lower() == saved_orig.lower() for p in sorted_originals)    : choice_var.set(next(p for p in sorted_originals   if p.lower() == saved_orig.lower()))
                else                                                                                : choice_var.set(sorted_originals[0]                if sorted_originals else "")

                self.sub_vars[sub_name] = choice_var

                combo_frame = tk.Frame(f_sub, bg = "white", bd = 1, relief = "solid")
                combo_frame.pack(fill = tk.X)

                entry = tk.Entry(combo_frame, textvariable = choice_var, bg = "white", fg = "black", font = ("Segoe UI", 10), justify = "center", bd = 0, state = "readonly")
                entry.pack(side = tk.LEFT, fill = tk.X, expand = True, padx = 4)

                arrow_btn = tk.Canvas(combo_frame, width = 25, height = 25, bg = "black", highlightthickness = 0, borderwidth = 0, cursor = "hand2")
                arrow_btn.create_polygon(7, 10, 17, 10, 12, 16, fill = "white")
                arrow_btn.pack(side = tk.RIGHT)

                def make_show_menu(s_players, c_var):
                    return lambda event: [
                        menu := tk.Menu(self, tearoff = 0),
                        *[menu.add_command(label = p, command = lambda val = p: c_var.set(val)) for p in s_players],
                        menu.post(event.x_root, event.y_root)
                    ]

                show_menu_func = make_show_menu(sorted_originals, choice_var)

                arrow_btn   .bind("<Button-1>", show_menu_func)
                entry       .bind("<Button-1>", show_menu_func)

        ttk.Label(self.container, text = "Are there any new players?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (6, 4))

        has_round_elo       = False
        round_elo_players   = set()

        if elo_map:
            for p in active_players:
                p_low = p.lower()

                if p_low in elo_map:
                    try:
                        val = float(elo_map[p_low])

                        if val.is_integer():
                            has_round_elo = True
                            round_elo_players.add(p_low)

                    except ValueError: pass

        self.np_var     = tk.StringVar(value = "Yes" if has_round_elo else "No")
        self.np_boxes   = {}

        for opt in ["No", "Yes"]:
            f_np = ttk.Frame(self.container)
            f_np.pack(anchor = "w", pady = 2)

            is_sel      = (self.np_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"
            box         = tk.Canvas(f_np, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")

            box.pack(side = tk.LEFT, padx = (0, 4))
            self.np_boxes[opt] = box

            lbl = ttk.Label(f_np, text = opt, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)

            for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_np_opt(o))

        self.player_container = ttk.Frame(self.container)
        self.player_container.pack(fill = tk.BOTH, expand = True, pady = (4, 0))

        self.player_vars    = {}
        player_list         = sorted(list(active_players), key = str.lower)
        num_players         = len(player_list)
        rows_per_col        = 8 if num_players >= 16 else num_players

        for i, name in enumerate(player_list):
            col                     = i //  rows_per_col
            row                     = i %   rows_per_col
            is_round                = name.lower() in round_elo_players
            var                     = tk.BooleanVar(value = is_round)
            self.player_vars[name]  = var

            item_frame              = ttk.Frame(self.player_container)
            item_frame.grid(row = row, column = col, padx = 2, pady = 2, sticky = "w")

            box = tk.Canvas(item_frame, width = 10, height = 10, bg = "white", highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 4))

            lbl = ttk.Label(item_frame, text = name, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)

            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, n=name, b = box: self.toggle_custom_player(n, b))

        self._update_np_state()

        def on_close_cancel():
            self.result = None
            self.destroy()

        self.protocol       ("WM_DELETE_WINDOW", on_close_cancel)
        self.grab_set       ()
        self.wait_window    ()

    def _select_lbl_opt(self, opt):
        self.lbl_var.set(opt)
        for k, box in self.lbl_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")
        self._update_lbl_state()

    def _select_th_opt(self, opt):
        self.th_var.set(opt)
        for k, box in self.th_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")
        self._update_th_state()

    def _select_np_opt(self, opt):
        self.np_var.set(opt)
        for k, box in self.np_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")
        self._update_np_state()

    def _update_lbl_state(self):
        if self.lbl_var.get() == "Others"   : self.lbl_entry.configure(state = "normal")
        else                                : self.lbl_entry.configure(state = "disabled")

    def _update_th_state(self):
        if self.th_var.get() == "custom"    : self.th_entry.configure(state = "normal")
        else                                : self.th_entry.configure(state = "disabled")

    def _update_np_state(self):
        state = "normal" if self.np_var.get() == "Yes" else "disabled"

        for child in self.player_container.winfo_children():
            for w in child.winfo_children():
                if isinstance(w, tk.Canvas):
                    if state == "disabled": w.configure(bg = "gray75")

                    else:
                        name = child.winfo_children()[1].cget("text")
                        w.configure(bg = self.fill_color if self.player_vars[name].get() else "white")

                elif isinstance(w, ttk.Label): w.configure(state = state)

    def toggle_custom_player(self, name, box):
        if self.np_var.get() != "Yes": self._select_np_opt("Yes")
        new_val = not self.player_vars[name].get()
        self.player_vars[name].set(new_val)
        color = self.fill_color if new_val else "white"
        box.configure(bg = color)

    def on_confirm(self):
        try                 : base_exp = int(self.spin.get())
        except ValueError   : base_exp = 1

        tour_label      = self.lbl_entry    .get() if self.lbl_var  .get() == "Others" else self.lbl_var.get()
        th_str          = self.th_entry     .get() if self.th_var   .get() == "custom" else "default"
        selected_new    = [name for name, var in self.player_vars.items() if var.get()] if self.np_var.get() == "Yes" else []
        sub_results     = {sub_name: var.get() for sub_name, var in self.sub_vars.items()}
        self.result     = {"tour_label": tour_label, "th_str": th_str, "base_exp": base_exp, "selected_new": selected_new, "sub_results": sub_results}

        if self.tour_dir and sub_results:
            existing_lines = []

            if (self.tour_dir / "subs.txt").exists():
                with open(self.tour_dir / "subs.txt", "r", encoding = "utf-8") as f:
                    for line in f:
                        if "," in line:
                            s_part, _ = line.split(",", 1)
                            if s_part.strip().lower() not in [s.lower() for s in sub_results]: existing_lines.append(line.strip())

            for s_name, r_name in sub_results.items()                           : existing_lines.append(f"{s_name}, {r_name}")
            with open(self.tour_dir / "subs.txt", "w", encoding = "utf-8") as f : f.write("\n".join(existing_lines) + "\n")

        super().on_confirm()

class MismatchedRoundsDialog(UnifiedDialog):
    def __init__(self, parent, mismatched_players, base_exp, subbed_players_set, tour_dir, is_watched = True):
        title_part = "These players appear" if len(mismatched_players) > 1 else "This player appears"

        prompt_text = (
            f"{title_part} in fewer JSONs than expected; how many rounds were they expected to be in?\n\n"
            '● "Use the current round count" is primarily used if the player has 0/0 round(s)\n'
            '● "Use the current JSON count" is primarily used if the player was subbed in/out\n'
            '● "Ignore mismatch" is primarily used for non-Watched tours\n'
        )

        super().__init__(parent, "Mismatched Round Counts", prompt_text)

        self.base_exp       = base_exp
        self.player_configs = {}
        self.fill_color     = "#000000"

        subs_txt_players    = set()
        subs_file           = tour_dir / "subs.txt"

        if subs_file.exists():
            with open(subs_file, "r", encoding = "utf-8") as f:
                for line in f:
                    if "," in line:
                        sub_player, original_player = line.split(",", 1)

                        subs_txt_players.add(sub_player         .strip().lower())
                        subs_txt_players.add(original_player    .strip().lower())

        for _, (name, act) in enumerate(sorted(mismatched_players.items())):
            p_frame = ttk.LabelFrame(self.container, text = f" {name} ", padding = 4)
            p_frame.pack(fill = tk.X, pady = 4, anchor = "w")

            is_subbed       = (name.lower() in subbed_players_set) or (name.lower() in subs_txt_players)
            initial_mode    = "custom" if not is_watched else ("json" if is_subbed else "round")
            mode_var        = tk.StringVar(value = initial_mode)
            boxes           = {}

            f_r1 = ttk.Frame(p_frame)
            f_r1.pack(anchor = "w", pady = 2)

            box_r1 = tk.Canvas(f_r1, width = 10, height = 10, bg = self.fill_color if initial_mode == "round" else "white", highlightthickness = 1, highlightbackground = "black")
            box_r1.pack(side = tk.LEFT, padx = (0, 4))

            boxes["round"] = box_r1

            lbl_r1 = ttk.Label(f_r1, text = "Use the current round count")
            lbl_r1.pack(side = tk.LEFT)

            f_r2 = ttk.Frame(p_frame)
            f_r2.pack(anchor = "w", pady = 2)

            box_r2 = tk.Canvas(f_r2, width = 10, height = 10, bg = self.fill_color if initial_mode == "json" else "white", highlightthickness = 1, highlightbackground = "black")
            box_r2.pack(side = tk.LEFT, padx = (0, 4))

            boxes["json"] = box_r2

            lbl_r2 = ttk.Label(f_r2, text = "Use the current JSON count")
            lbl_r2.pack(side = tk.LEFT)

            f_custom = ttk.Frame(p_frame)
            f_custom.pack(anchor = "w", pady = 2)

            box_r3 = tk.Canvas(f_custom, width = 10, height = 10, bg = self.fill_color if initial_mode == "custom" else "white", highlightthickness = 1, highlightbackground = "black")
            box_r3.pack(side = tk.LEFT, padx = (0, 4))

            boxes["custom"] = box_r3

            lbl_r3 = ttk.Label(f_custom, text = "Ignore mismatch")
            lbl_r3.pack(side = tk.LEFT)

            def make_selector(m_var, b_map, target_opt):
                return lambda _: [
                    m_var.set(target_opt),
                    *[b.configure(bg = self.fill_color if k == target_opt else "white") for k, b in b_map.items()]
                ]

            for w, opt in [(box_r1, "round"),   (lbl_r1, "round")]  : w.bind("<Button-1>", make_selector(mode_var, boxes, opt))
            for w, opt in [(box_r2, "json"),    (lbl_r2, "json")]   : w.bind("<Button-1>", make_selector(mode_var, boxes, opt))
            for w, opt in [(box_r3, "custom"),  (lbl_r3, "custom")] : w.bind("<Button-1>", make_selector(mode_var, boxes, opt))

            self.player_configs[name] = {"mode": mode_var, "act": act}
 
        def on_close_cancel():
            self.result = None
            self.destroy()

        self.protocol       ("WM_DELETE_WINDOW", on_close_cancel)
        self.grab_set       ()
        self.wait_window    ()

    def on_confirm(self):
        self.result = {}

        for name, cfg in self.player_configs.items():
            mode = cfg["mode"].get()

            if      mode == "round" : self.result[name] = self.base_exp
            elif    mode == "json"  : self.result[name] = cfg["act"]
            else                    : self.result[name] = "ignore"

        super().on_confirm()

class SubstitutePromptDialog(UnifiedDialog):
    def __init__(self, parent, sub_name, original_players_list, tour_dir):
        super().__init__(parent, "Substitute Setup", "")

        self.result         = None
        self._sub_name      = sub_name
        self.subs_txt_path  = tour_dir / "subs.txt"

        lbl = ttk.Label(self.container, text = f"Who is {sub_name} subbing for?", font = ("Segoe UI", 10))
        lbl.grid(row = 0, column = 0, padx = (0, 4), pady = 4, sticky = "w")

        self.choice_var     = tk.StringVar()
        self.sorted_players = sorted(list(original_players_list), key = str.lower)

        saved_original = None

        if self.subs_txt_path.exists():
            with open(self.subs_txt_path, "r", encoding = "utf-8") as f:
                for line in f:
                    if "," in line:
                        s_name, o_name = line.strip().split(",", 1)

                        if s_name.strip().lower() == sub_name.lower():
                            saved_original = next((p for p in self.sorted_players if p.lower() == o_name.strip().lower()), None)
                            if saved_original: break

        if      saved_original      : self.choice_var.set(saved_original)
        elif    self.sorted_players : self.choice_var.set(self.sorted_players[0])

        combo_frame = tk.Frame(self.container, bg = "white", bd = 1, relief = "solid")
        combo_frame.grid(row = 0, column = 1, pady = 4, sticky = "ew")

        self.container.grid_columnconfigure(1, weight = 1)

        self.entry = tk.Entry(combo_frame, textvariable = self.choice_var, bg = "white", fg = "black", font = ("Segoe UI", 10), justify = "center", bd = 0, state = "readonly")
        self.entry.pack(side = tk.LEFT, fill = tk.X, expand = True, padx = 4)

        arrow_btn = tk.Canvas(combo_frame, width = 25, height = 25, bg = "black", highlightthickness = 0, borderwidth = 0, cursor = "hand2")
        arrow_btn.create_polygon(7, 10, 17, 10, 12, 16, fill = "white")
        arrow_btn.pack(side = tk.RIGHT)

        def show_menu(event):
            menu = tk.Menu(self, tearoff = 0)
            for p in self.sorted_players: menu.add_command(label = p, command = lambda val = p: self.choice_var.set(val))
            menu.post(event.x_root, event.y_root)

        arrow_btn   .bind("<Button-1>", show_menu)
        self.entry  .bind("<Button-1>", show_menu)

        def on_close_cancel():
            self.result = None
            self.destroy()

        self.protocol       ("WM_DELETE_WINDOW", on_close_cancel)
        self.grab_set       ()
        self.wait_window    ()

    def on_confirm(self):
        self.result     = self.choice_var.get()
        existing_lines  = []

        if self.subs_txt_path.exists():
            with open(self.subs_txt_path, "r", encoding = "utf-8") as f:
                for line in f:
                    if "," in line:
                        s_part, _ = line.split(",", 1)
                        if s_part.strip().lower() != self._sub_name.lower(): existing_lines.append(line.strip())

        existing_lines.append(f"{self._sub_name}, {self.result}")        
        with open(self.subs_txt_path, "w", encoding = "utf-8") as f: f.write("\n".join(existing_lines) + "\n")
        super().on_confirm()