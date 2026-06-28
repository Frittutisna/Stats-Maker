import tkinter as tk

from help.config    import DIR_TOURS, DIR_JSONS, FILE_ALIAS
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

        self.entry.grid(row = 0, column = 1, sticky = "ns")
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
        main_frame.pack(fill = tk.BOTH)

        if prompt: ttk.Label(main_frame, text = prompt, font = ("Segoe UI", 10)).pack(anchor = "w")

        self.container = ttk.Frame(main_frame)
        self.container.pack(fill = tk.BOTH)

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
            is_recommended  = False

            if json_dir.exists() and len(list(json_dir.glob("*.json"))) > 1: is_recommended = True
            states[tid] = is_recommended

        for tid in tour_ids:
            is_active       = states[tid]
            var             = tk.BooleanVar(value = is_active)
            self.vars[tid]  = var
            item_frame      = ttk.Frame(self.container)

            item_frame.pack(anchor = "w")
            
            if is_active:
                initial_bg = self.fill_color
                lbl_state  = "normal"

            else:
                initial_bg = "gray75"
                lbl_state  = "disabled"

            box = tk.Canvas(item_frame, width = 10, height = 10, bg = initial_bg, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 4))

            lbl = ttk.Label(item_frame, text = f"Tour {tid}", font = ("Segoe UI", 10), state = lbl_state)
            lbl.pack(side = tk.LEFT)

            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, t = tid, b = box, l = lbl, a = is_active: self.toggle_custom(t, b, l, a))

        self.protocol       ("WM_DELETE_WINDOW", lambda: [setattr(self, 'selected_tours', []), self.destroy()])
        self.grab_set       ()
        self.wait_window    ()

    def toggle_custom(self, tid, box, lbl, is_recommended_tour):
        new_val = not self.vars[tid].get()
        self.vars[tid].set(new_val)
        
        if new_val:
            box.configure(bg    = self.fill_color)
            lbl.configure(state = "normal")
        else:
            if is_recommended_tour:
                box.configure(bg    = "white")
                lbl.configure(state = "normal")

            else:
                box.configure(bg    = "gray75")
                lbl.configure(state = "disabled")

    def on_confirm(self):
        self.selected_tours = [tid for tid, var in self.vars.items() if var.get()]
        super().on_confirm()

class TourMetadataDialog(UnifiedDialog):
    def __init__(self, parent, tour_id, init_label, default_th, baseline_initial, active_players, elo_map = None, sub_candidates = None, original_players_list = None, tour_dir = None):
        super().__init__(parent, f"Tour {tour_id} Configuration", "")

        self.fill_color = "#000000"
        ttk.Label(self.container, text = "What tour is this?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w")

        if      "Watched"   in init_label or init_label in ["Brute-force", "Masquerade", "Other Random", "Other Watched"]   : starting_lbl = init_label
        elif    "Random"    in init_label                                                                                   : starting_lbl = init_label
        elif    init_label == "Usual"                                                                                       : starting_lbl = "Usual"
        else                                                                                                                : starting_lbl = "Others"

        self.lbl_var = tk.StringVar(value=starting_lbl if starting_lbl in ["Watched", "Usual"] else "Others")
        self.lbl_boxes = {}

        for opt in ["Random", "Watched", "Others"]:
            f_opt = ttk.Frame(self.container)
            f_opt.pack(anchor = "w", pady = 1)

            is_sel      = (self.lbl_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"

            box = tk.Canvas(f_opt, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 4))
            self.lbl_boxes[opt] = box

            if opt == "Others":
                lbl = ttk.Label(f_opt, text = "Others:", font = ("Segoe UI", 10))
                lbl.pack(side = tk.LEFT)

                self.lbl_entry = ttk.Entry(f_opt, width = 20)
                if self.lbl_var.get() == "Others": self.lbl_entry.insert(0, init_label)
                self.lbl_entry.pack(side=tk.LEFT, padx=(4, 0))

                for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_lbl_opt(o))

            else:
                lbl = ttk.Label(f_opt, text = opt, font = ("Segoe UI", 10))
                lbl.pack(side = tk.LEFT)

                for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_lbl_opt(o))

        self.sub_lbl_container = ttk.Frame(self.container)
        self.sub_lbl_container.pack(fill = tk.BOTH)

        columns_layout = [
            ["Watched OP",      "Watched ED",       "Watched IN",   "Watched IN -Chanting", "Watched 2+8s",     "Watched 5s", "Watched -2009"],
            ["Random OP",       "Random ED",        "Random IN",    "Random OPED",          "Random Chanting"],
            ["Other Random",    "Other Watched",    "Brute-force",  "Masquerade"]
        ]

        self.sub_lbl_widgets = {}

        for col_idx, items in enumerate(columns_layout):
            for row_idx, name in enumerate(items):
                item_frame = ttk.Frame(self.sub_lbl_container)
                item_frame.grid(row = row_idx, column = col_idx, padx = 4, sticky = "w")

                is_active = (init_label == name)
                if is_active: self.lbl_var.set(name)

                bg_color    = self.fill_color if is_active else "white"
                box         = tk.Canvas(item_frame, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")
                box.pack(side = tk.LEFT, padx = (0, 4))

                self.lbl_boxes[name]    = box
                lbl                     = ttk.Label(item_frame, text = name, font = ("Segoe UI", 10))
                lbl.pack(side = tk.LEFT)
                
                self.sub_lbl_widgets[name] = lbl
                for widget in (box, lbl): widget.bind("<Button-1>", lambda _, n=name: self._select_lbl_opt(n))

        self._update_lbl_state()

        has_extended_delta_data = False
        script_root_dir         = Path(__file__).parent.parent.absolute()
        global_alias_path       = script_root_dir / DIR_TOURS / FILE_ALIAS

        if global_alias_path.exists():
            try:
                with open(global_alias_path, "r", encoding = "utf-8") as f_alias:
                    for line in f_alias:
                        if "," in line:
                            parts = line.strip().split(",")
                            if len(parts) > 2:
                                has_extended_delta_data = True
                                break

            except Exception: pass

        suggested_delta_default = "No" if has_extended_delta_data else "Yes"
        self.delta_var          = tk.StringVar(value = suggested_delta_default)
        self.delta_boxes        = {}

        ttk.Label(self.container, text = "Do you want to fetch Δ data as well?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (5, 0))

        for opt in ["No", "Yes"]:
            f_delta = ttk.Frame(self.container)
            f_delta.pack(anchor = "w", pady = 1)

            is_sel      = (self.delta_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"

            box = tk.Canvas(f_delta, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 4))
            self.delta_boxes[opt] = box

            lbl = ttk.Label(f_delta, text = opt, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)

            def make_delta_callback(choice_opt):
                return lambda _: self._select_delta_opt(choice_opt)

            for w in (box, lbl): w.bind("<Button-1>", make_delta_callback(opt))

        self.challonge_var      = tk.StringVar(value = "No")
        self.challonge_boxes    = {}

        ttk.Label(self.container, text = "Do you want to fetch Challonge data as well?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (5, 0))

        for opt in ["No", "Yes"]:
            f_chal = ttk.Frame(self.container)
            f_chal.pack(anchor = "w", pady = 1)

            is_sel      = (self.challonge_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"

            box = tk.Canvas(f_chal, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 4))
            self.challonge_boxes[opt] = box

            lbl = ttk.Label(f_chal, text = opt, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)

            for w in (box, lbl): w.bind("<Button-1>", lambda _, o=opt: self._select_challonge_opt(o))

        ttk.Label(self.container, text = "Do you want to use Dry's script as well?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (5, 0))

        self.dry_var    = tk.StringVar(value = "No")
        self.dry_boxes  = {}
        dry_options     = ["No", "Yes, but don't push it to the database", "Yes, and push it to the database"]

        for opt in dry_options:
            f_dry = ttk.Frame(self.container)
            f_dry.pack(anchor = "w", pady = 1)

            is_sel      = (self.dry_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"

            box = tk.Canvas(f_dry, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 4))
            self.dry_boxes[opt] = box

            lbl = ttk.Label(f_dry, text = opt, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)
            
            for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_dry_opt(o))

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
        ttk.Label(self.container, text = "How many rounds have elapsed?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w")

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
                f_sub.pack(anchor = "w") 

                lbl = ttk.Label(f_sub, text = f"Who is {sub_name} subbing for?", font = ("Segoe UI", 10, "bold"))
                lbl.pack(anchor = "w")

                choice_var = tk.StringVar()
                saved_orig = saved_subs_map.get(sub_name.lower())

                if saved_orig and any(p.lower() == saved_orig.lower() for p in sorted_originals)    : choice_var.set(next(p for p in sorted_originals   if p.lower() == saved_orig.lower()))
                else                                                                                : choice_var.set(sorted_originals[0]                if sorted_originals else "")

                self.sub_vars[sub_name] = choice_var

                combo_frame = tk.Frame(f_sub, bg = "white", bd = 1, relief = "solid")
                combo_frame.pack(anchor = "w")

                entry = tk.Entry(combo_frame, textvariable = choice_var, bg = "white", fg = "black", font = ("Segoe UI", 10), justify = "center", bd = 0, state = "readonly")
                entry.pack(side = tk.LEFT)

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

        ttk.Label(self.container, text = "Are there any new players?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w")

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
            f_np.pack(anchor = "w")

            is_sel      = (self.np_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"
            box         = tk.Canvas(f_np, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")

            box.pack(side = tk.LEFT, padx = (0, 4))
            self.np_boxes[opt] = box

            lbl = ttk.Label(f_np, text = opt, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)

            for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_np_opt(o))

        self.player_container = ttk.Frame(self.container)
        self.player_container.pack(fill = tk.BOTH)

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
            item_frame.grid(row = row, column = col, padx = 4, sticky = "w")

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

    def _select_dry_opt(self, opt):
        self.dry_var.set(opt)
        for k, box in self.dry_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")

    def _select_th_opt(self, opt):
        self.th_var.set(opt)
        for k, box in self.th_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")
        self._update_th_state()

    def _select_np_opt(self, opt):
        self.np_var.set(opt)
        for k, box in self.np_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")
        self._update_np_state()

    def _update_lbl_state(self):
        current_selection   = self.lbl_var.get()
        is_others_active    = (current_selection == "Others" or current_selection in getattr(self, 'sub_lbl_widgets', {}))
        state               = "normal" if is_others_active else "disabled"

        if current_selection == "Others"    : self.lbl_entry.configure(state = "normal")
        else                                : self.lbl_entry.configure(state = "disabled")

        for name, lbl in getattr(self, 'sub_lbl_widgets', {}).items():
            box = self.lbl_boxes.get(name)

            if state == "disabled":
                lbl.configure(state = "disabled")
                if box: box.configure(bg = "gray75")

            else:
                lbl.configure(state = "normal")
                if box: box.configure(bg = self.fill_color if current_selection == name else "white")

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

    def _select_delta_opt(self, opt):
        self.delta_var.set(opt)
        for k, box in self.delta_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")

    def _select_challonge_opt(self, opt):
        self.challonge_var.set(opt)
        for k, box in self.challonge_boxes.items(): box.configure(bg = self.fill_color if k == opt else "white")

    def on_confirm(self):
        try                 : base_exp = int(self.spin.get())
        except ValueError   : base_exp = 1

        tour_label      = self.lbl_entry    .get() if self.lbl_var  .get() == "Others" else self.lbl_var.get()
        th_str          = self.th_entry     .get() if self.th_var   .get() == "custom" else "default"
        selected_new    = [name for name, var in self.player_vars.items() if var.get()] if self.np_var.get() == "Yes" else []
        sub_results     = {sub_name: var.get() for sub_name, var in self.sub_vars.items()}
        dry_choice      = self.dry_var      .get()
        delta_choice    = self.delta_var    .get()
        self.result     = {
            "tour_label"        : tour_label, 
            "th_str"            : th_str, 
            "base_exp"          : base_exp, 
            "selected_new"      : selected_new, 
            "sub_results"       : sub_results, 
            "dry_choice"        : dry_choice,
            "delta_choice"      : delta_choice,
            "challonge_choice"  : self.challonge_var.get()
        }

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