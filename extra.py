import  json, os, re, math
import  matplotlib.pyplot   as      plt
import  matplotlib.colors   as      mcolors
import  numpy               as      np
import  pandas              as      pd
import  tkinter             as      tk
from    adjustText          import  adjust_text
from    collections         import  Counter,    defaultdict
from    html2image          import  Html2Image
from    pathlib             import  Path
from    PIL                 import  Image,      ImageChops,     ImageOps
from    tkinter             import  messagebox, ttk

BROWSER_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

DIR_DEPS            = "dependencies"
DIR_JSONS           = "jsons"
DIR_OUT             = "output"
DIR_TOURS           = "tours"
FILE_CODES          = "codes.txt"
FILE_ALIASES        = "aliases.txt"
RIG_GR_THRESHOLD    = 0.85
URL_ALIAS           = "https://docs.google.com/spreadsheets/d/10YBcZP_l5Tjf1MOiWeBlLg-ATuAWXgTPsj7bW79bU30/export?format=csv&gid=1934025140"

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
    if not vintage_str: return None
    years = re.findall(r'\d{4}', str(vintage_str))
    if not years: return None
    year_val    = float(years[0])
    season_map  = {"winter": 0.00, "spring": 0.25, "summer": 0.50, "fall": 0.75}
    v_lower     = str(vintage_str).lower()
    decimal     = next((val for s, val in season_map.items() if s in v_lower), 0.0)
    return year_val + decimal

def format_year(val):
    if val is None or val == "N/A": return "N/A"
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
            img.save(image_path)

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
        if self._state == "disabled"    : return

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
        self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        main_frame = ttk.Frame(self, padding = 15)
        main_frame.pack(fill = tk.BOTH, expand = True)
        if prompt: ttk.Label(main_frame, text = prompt, font = ("Segoe UI", 10)).pack(pady = (0, 10), anchor = "w")
        self.container = ttk.Frame(main_frame)
        self.container.pack(fill = tk.BOTH, expand = True)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill = tk.X, pady = (15, 0))
        self.confirm_btn = ttk.Button(btn_frame, text = "Confirm", command = self.on_confirm)
        self.confirm_btn.pack(side = tk.RIGHT, padx = 5)
        self.bind("<Return>", lambda: self.on_confirm())

    def on_confirm(self): self.destroy()

class TourSelectionDialog(UnifiedDialog):
    def __init__(self, parent, tour_ids):
        super().__init__(parent, "Tour Selection", "Which tours should be processed?")
        self.selected_tours = []
        self.vars           = {}
        self.fill_color     = "#000000"
        script_dir          = Path(__file__).parent.absolute()
        states              = {}
        any_recommended     = False

        for tid in tour_ids:
            t_path          = script_dir / DIR_TOURS / str(tid)
            json_dir        = t_path / DIR_JSONS
            codes_file      = t_path / FILE_CODES
            is_recommended  = False
            
            if codes_file.exists() and json_dir.exists():
                json_count = len(list(json_dir.glob("*.json")))

                if json_count > 0:
                    with open(codes_file, "r", encoding = "utf-8") as f:
                        content     = f.read()
                        main_part   = re.split(r'https://challonge\.com/\S+', content)[0]
                        players     = re.findall(r'[^\s(]+\s*\([-]?\d+\.\d+\)', main_part)
                        p           = len(players)

                        if p > 0:
                            divisor = p // 8

                            if divisor > 0 and json_count % divisor == 0: 
                                is_recommended  = True
                                any_recommended = True

            states[tid] = is_recommended

        if not any_recommended and "0" in states: states["0"] = True

        for tid in tour_ids:
            is_active       = states[tid]
            var             = tk.BooleanVar(value = is_active)
            self.vars[tid]  = var
            item_frame      = ttk.Frame(self.container)

            item_frame.pack(anchor = "w", pady = 2)
            initial_bg = self.fill_color if is_active else "white"
            box = tk.Canvas(item_frame, width = 10, height = 10, bg = initial_bg, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 5))
            lbl = ttk.Label(item_frame, text = f"Tour {tid}", font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)
            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, t = tid, b = box: self.toggle_custom(t, b))
            
        self.grab_set()
        self.wait_window()

    def toggle_custom(self, tid, box):
        new_val = not self.vars[tid].get()
        self.vars[tid].set(new_val)
        color = self.fill_color if new_val else "white"
        box.configure(bg = color)

    def on_confirm(self):
        self.selected_tours = [tid for tid, var in self.vars.items() if var.get()]
        super().on_confirm()

class TourMetadataDialog(UnifiedDialog):
    def __init__(self, parent, tour_id, init_label, default_th, baseline_initial, active_players):
        super().__init__(parent, f"Tour {tour_id} Configuration", "")
        self.fill_color = "#000000"
        ttk.Label(self.container, text = "What tour is this?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (0, 2))
        self.lbl_var    = tk.StringVar(value = init_label if init_label in ["Watched", "Usual"] else "Others")
        self.lbl_boxes  = {}

        for opt in ["Watched", "Usual", "Others"]:
            f_opt = ttk.Frame(self.container)
            f_opt.pack(anchor = "w", pady = 2)

            is_sel      = (self.lbl_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"
            box         = tk.Canvas(f_opt, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")

            box.pack(side = tk.LEFT, padx = (0, 5))
            self.lbl_boxes[opt] = box
            
            if opt == "Others":
                lbl = ttk.Label(f_opt, text = "Others:", font = ("Segoe UI", 10))
                lbl.pack(side = tk.LEFT)
                self.lbl_entry = ttk.Entry(f_opt, width = 20)
                self.lbl_entry.insert(0, init_label)
                self.lbl_entry.pack(side = tk.LEFT, padx = (5, 0))
                for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_lbl_opt(o))
            else:
                lbl = ttk.Label(f_opt, text = opt, font = ("Segoe UI", 10))
                lbl.pack(side = tk.LEFT)
                for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_lbl_opt(o))
                    
        self._update_lbl_state()

        ttk.Label(self.container, text = "What are the comma-separated guess rate threshold values?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (10, 2))
        self.th_var = tk.StringVar(value = "default")
        
        self.th_boxes   = {}
        f_th1           = ttk.Frame(self.container)
        f_th1.pack(anchor = "w", pady = 2)
        box_th1 = tk.Canvas(f_th1, width = 10, height = 10, bg = self.fill_color, highlightthickness = 1, highlightbackground = "black")
        box_th1.pack(side = tk.LEFT, padx = (0, 5))
        self.th_boxes["default"]    = box_th1
        lbl_th1                     = ttk.Label(f_th1, text = "Use the default threshold values", font = ("Segoe UI", 10))
        lbl_th1.pack(side = tk.LEFT)
        for w in (box_th1, lbl_th1): w.bind("<Button-1>", lambda _: self._select_th_opt("default"))
            
        f_th2 = ttk.Frame(self.container)
        f_th2.pack(anchor = "w", pady = 2)
        box_th2 = tk.Canvas(f_th2, width = 10, height = 10, bg = "white", highlightthickness = 1, highlightbackground = "black")
        box_th2.pack(side = tk.LEFT, padx = (0, 5))
        self.th_boxes["custom"] = box_th2
        lbl_th2                 = ttk.Label(f_th2, text = "Use custom threshold values:", font = ("Segoe UI", 10))
        lbl_th2.pack(side = tk.LEFT)
        self.th_entry = ttk.Entry(f_th2, width = 25)
        self.th_entry.insert(0, default_th)
        self.th_entry.pack(side = tk.LEFT, padx = (5, 0))
        for w in (box_th2, lbl_th2): w.bind("<Button-1>", lambda _: self._select_th_opt("custom"))
            
        self._update_th_state()

        ttk.Label(self.container, text = "How many rounds have elapsed?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (10, 2))
        self.spin = CustomSpinbox(self.container, from_ = 1, to = 6, initial_val = baseline_initial)
        self.spin.pack(anchor = "w", pady = (0, 10))

        ttk.Label(self.container, text = "Are there any new players?", font = ("Segoe UI", 10, "bold")).pack(anchor = "w", pady = (5, 2))
        self.np_var = tk.StringVar(value = "No")

        self.np_boxes = {}

        for opt in ["No", "Yes"]:
            f_np = ttk.Frame(self.container)
            f_np.pack(anchor = "w", pady = 2)
            is_sel      = (self.np_var.get() == opt)
            bg_color    = self.fill_color if is_sel else "white"
            box = tk.Canvas(f_np, width = 10, height = 10, bg = bg_color, highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 5))
            self.np_boxes[opt] = box
            lbl = ttk.Label(f_np, text = opt, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)
            for w in (box, lbl): w.bind("<Button-1>", lambda _, o = opt: self._select_np_opt(o))

        self.player_container = ttk.Frame(self.container)
        self.player_container.pack(fill = tk.BOTH, expand = True, pady = (5, 0))
        
        self.player_vars    = {}
        player_list         = sorted(list(active_players), key = str.lower)
        num_players         = len(player_list)
        rows_per_col        = 8 if num_players >= 16 else num_players

        for i, name in enumerate(player_list):
            col                     = i //  rows_per_col
            row                     = i %   rows_per_col
            var                     = tk.BooleanVar(value = False)
            self.player_vars[name]  = var
            item_frame              = ttk.Frame(self.player_container)

            item_frame.grid(row = row, column = col, padx = 5, pady = 1, sticky = "w")
            
            box = tk.Canvas(item_frame, width = 10, height = 10, bg = "white", highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 5))
            lbl = ttk.Label(item_frame, text = name, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)
            
            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, n=name, b = box: self.toggle_custom_player(n, b))

        self._update_np_state   ()
        self.grab_set           ()
        self.wait_window        ()

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
                    if state == "disabled": w.configure(bg="gray75")
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
        
        self.result = {"tour_label": tour_label, "th_str": th_str, "base_exp": base_exp, "selected_new": selected_new}
        super().on_confirm()

class MismatchedRoundsDialog(UnifiedDialog):
    def __init__(self, parent, mismatched_players, base_exp, subbed_players_set):
        title_part = "These players appear" if len(mismatched_players) > 1 else "This player appears"
        
        prompt_text = (
            f"{title_part} in fewer JSONs than expected; how many rounds were they expected to be in?\n"
            f'"Use the current round count" is primarily used if the player has 0/0 round(s)\n'
            f'"Use the current JSON count" is primarily used if the player was subbed in/out'
        )
        
        super().__init__(parent, "Mismatched Round Counts", prompt_text)
        self.base_exp       = base_exp
        self.player_configs = {}
        self.fill_color     = "#000000"

        for _, (name, act) in enumerate(mismatched_players.items()):
            p_frame = ttk.LabelFrame(self.container, text = f" {name} ", padding = 10)
            p_frame.pack(fill = tk.X, pady = 5, anchor = "w")
            
            is_subbed       = name.lower() in subbed_players_set
            initial_mode    = "json" if is_subbed else "round"
            mode_var        = tk.StringVar(value=initial_mode)
            boxes           = {}
            
            f_r1 = ttk.Frame(p_frame)
            f_r1.pack(anchor = "w", pady = 1)
            box_r1 = tk.Canvas(f_r1, width = 10, height = 10, bg = self.fill_color if initial_mode == "round" else "white", highlightthickness = 1, highlightbackground = "black")
            box_r1.pack(side = tk.LEFT, padx = (0, 5))
            boxes["round"] = box_r1
            
            lbl_r1 = ttk.Label(f_r1, text = "Use the current round count")
            lbl_r1.pack(side = tk.LEFT)
            
            f_r2 = ttk.Frame(p_frame)
            f_r2.pack(anchor = "w", pady = 1)
            box_r2 = tk.Canvas(f_r2, width = 10, height = 10, bg = self.fill_color if initial_mode == "json" else "white", highlightthickness = 1, highlightbackground = "black")
            box_r2.pack(side = tk.LEFT, padx = (0, 5))
            boxes["json"] = box_r2
            
            lbl_r2 = ttk.Label(f_r2, text = "Use the current JSON count")
            lbl_r2.pack(side = tk.LEFT)
            
            f_custom = ttk.Frame(p_frame)
            f_custom.pack(anchor = "w", pady = 1)
            box_r3 = tk.Canvas(f_custom, width = 10, height = 10, bg = "white", highlightthickness = 1, highlightbackground = "black")
            box_r3.pack(side = tk.LEFT, padx = (0, 5))
            boxes["custom"] = box_r3
            
            lbl_r3 = ttk.Label(f_custom, text = "Use a custom value:")
            lbl_r3.pack(side = tk.LEFT)
            
            spin = CustomSpinbox(f_custom, from_ = act, to = base_exp, initial_val = act)
            spin.pack(side = tk.LEFT, padx = 5)
            spin.configure_state("normal" if initial_mode == "custom" else "disabled")
            
            def make_selector(m_var, b_map, target_opt, s_box):
                return lambda _: [
                    m_var.set(target_opt),
                    s_box.configure_state("normal" if target_opt == "custom" else "disabled"),
                    *[b.configure(bg = self.fill_color if k == target_opt else "white") for k, b in b_map.items()]
                ]
                
            for w, opt in [(box_r1, "round"),   (lbl_r1, "round")]  : w.bind("<Button-1>", make_selector(mode_var, boxes, opt, spin))
            for w, opt in [(box_r2, "json"),    (lbl_r2, "json")]   : w.bind("<Button-1>", make_selector(mode_var, boxes, opt, spin))
            for w, opt in [(box_r3, "custom"),  (lbl_r3, "custom")] : w.bind("<Button-1>", make_selector(mode_var, boxes, opt, spin))
            
            self.player_configs[name] = {"mode": mode_var, "spin": spin, "act": act}
            
        self.grab_set       ()
        self.wait_window    ()

    def on_confirm(self):
        self.result = {}

        for name, cfg in self.player_configs.items():
            mode = cfg["mode"].get()

            if      mode == "round" : self.result[name] = self.base_exp
            elif    mode == "json"  : self.result[name] = cfg["act"]
            else:
                try                 : self.result[name] = int(cfg["spin"].get())
                except ValueError   : self.result[name] = cfg["act"]

        super().on_confirm()

class ManualMatchDialog(tk.Toplevel):
    def __init__(self, parent, unknown_name, available_pool):
        super().__init__(parent)
        self.title("Manual Match Required")
        self.result = None
        ttk.Label(self, text = f"Match required for: '{unknown_name}'", font = ("Arial", 10, "bold")).pack(pady = 10)
        self.listbox = tk.Listbox(self, height = 15)
        self.listbox.pack(padx = 10, fill = tk.BOTH)
        for name in sorted(available_pool): self.listbox.insert(tk.END, name)
        ttk.Button(self, text = "Match Selected", command = self.on_match).pack(pady = 10)
        self.grab_set(); self.wait_window()

    def on_match(self):
        sel = self.listbox.curselection()
        if sel: self.result = self.listbox.get(sel[0]); self.destroy()

class SubSelectionDialog(tk.Toplevel):
    def __init__(self, parent, missing_roster):
        super().__init__(parent)
        self.title("Substitute Resolution")
        self.result = None
        tk.Label(self, text = "Select the subbed player:").pack(padx = 20, pady = 10)
        self.listbox = tk.Listbox(self, height = len(missing_roster))
        self.listbox.pack(padx = 20, pady = 5, fill = tk.X)
        for m in missing_roster: self.listbox.insert(tk.END, m)
        ttk.Button(self, text = "Confirm", command = self.on_confirm).pack(pady = 10)
        self.grab_set(); self.wait_window()

    def on_confirm(self):
        sel = self.listbox.curselection()
        if sel: self.result = self.listbox.get(sel[0]); self.destroy()

# Main Processor
class TourAnalyzer:
    def __init__(self, tour_id):
        self.tour_id                    = str(tour_id)
        self.script_dir                 = Path(__file__).parent.absolute()
        self.tour_dir                   = self.script_dir / DIR_TOURS / self.tour_id
        self.browser_path               = self._find_browser()
        self.s_part                     = defaultdict(int)
        self.c_counts                   = defaultdict(int)
        self.e_counts                   = defaultdict(int)
        self.p_rev_e                    = defaultdict(int)
        self.p_two_e                    = defaultdict(int)
        self.p_pts                      = defaultdict(int)
        self.p_blks                     = defaultdict(int)
        self.p_type_c                   = defaultdict(lambda: defaultdict(int))
        self.p_type_s                   = defaultdict(lambda: defaultdict(int))
        self.p_rigs                     = defaultdict(int)
        self.p_rigs_h                   = defaultdict(int)
        self.p_l_vint                   = defaultdict(list)
        self.p_l_corr                   = defaultdict(list)
        self.p_m_erigs                  = defaultdict(int)
        self.p_l_solos                  = defaultdict(int)
        self.p_chan_c                   = defaultdict(int)
        self.p_chan_s                   = defaultdict(int)
        self.t_vint                     = defaultdict(list)
        self.t_c_ps                     = defaultdict(list)
        self.t_on_syn                   = defaultdict(list)
        self.t_off_syn                  = defaultdict(list)
        self.t_sh_rig                   = defaultdict(list)
        self.t_solos                    = defaultdict(int)
        self.t_sweeps                   = defaultdict(int)
        self.t_overs                    = defaultdict(list)
        self.genre_c                    = Counter()
        self.tag_c                      = Counter()
        self.global_stats               = Counter()
        self.all_diff, self.all_vint    = [], []
        self.song_history               = []
        self.chanting_ids               = set()
        self.subbed_players_set         = set()
        self.tour_label                 = ""
        self.id_database                = {}

    def _find_browser(self): return next((p for p in BROWSER_PATHS if os.path.exists(p)), None)

    def _load_player_ids(self):
        id_map = {}

        try:
            df = pd.read_csv(URL_ALIAS)

            for _, row in df.iterrows():
                name    = str(row.get('Player Name',    '')).strip().lower()
                pid     = str(row.get('Player ID',      '')).strip()

                if name and pid: id_map[name] = pid
        except: pass

        return id_map

    def run(self):
        chanting_path = self.script_dir / DIR_DEPS / "chanting" / "chanting.txt"

        if chanting_path.exists():
            with open(chanting_path, "r") as f:
                for line in f:
                    line = line.strip()
                    if line: self.chanting_ids.add(line)

        json_dir = self.tour_dir / DIR_JSONS

        if not json_dir.exists() or not any(json_dir.glob("*.json")):
            messagebox.showerror("Error", f"Folder not found or empty: {json_dir}")
            return

        json_paths                                                      = list(json_dir.glob("*.json"))
        all_known, appearances                                          = self._scan_players    (json_paths)
        use_teams, elo_map, assignments, t1_lookup, rosters, all_known  = self._load_team_data  (all_known)
        missing_list_count                                              = 0
        tour_types                                                      = set()

        for path in json_paths:
            with open(path, encoding = "utf-8") as f: data = json.load(f)
            
            songs = data.get("songs", [])
            if not songs: continue

            raw_f_players = {p for s in songs for p in s.get("correctGuessPlayers", [])} | {ls["name"] for s in songs for ls in s.get("listStates", [])}
            final_members = set(raw_f_players)

            if use_teams:
                t_in_f = {assignments[p.lower()][0] for p in raw_f_players if p.lower() in assignments}
                for tid in t_in_f:
                    ros     = rosters[tid]
                    missing = [p for p in ros if p not in raw_f_players]
                    if len([p for p in ros if p in raw_f_players]) == 3 and missing:
                        res = SubSelectionDialog(None, missing).result if len(missing) > 1 else missing[0]
                        if res:
                            final_members.add(res)
                            potential_subs = list(raw_f_players - rosters[tid])
                            for sub_candidate in potential_subs:
                                if sub_candidate.lower() not in assignments: assignments[sub_candidate.lower()] = assignments[res.lower()]

                if len(final_members) < 8:
                    for tid in t_in_f: final_members.update(rosters[tid])

            apply_rev       = (len(final_members) % 2 == 0)
            max_s           = max(s.get("songNumber", 0) for s in songs)
            f_type_totals   = defaultdict(int)

            for song in songs:
                st = song.get("songInfo", {}).get("type")
                if st in [1, 2, 3]: 
                    f_type_totals[st] += 1
                    tour_types.add(st)

            for name in final_members:
                if name in raw_f_players:
                    self.s_part[name] += max_s
                    for t in [1, 2, 3]: self.p_type_s[name][t] += f_type_totals[t]

            for song in songs:
                si      = song      .get("songInfo", {})
                st      = si        .get("type")
                ann_id  = str(si    .get("annSongId"))
                is_chan = ann_id in self.chanting_ids
                
                if isinstance(si.get("animeGenre"), list): self.genre_c .update(si.get("animeGenre"))
                if isinstance(si.get("animeTags"),  list): self.tag_c   .update([t for t in si.get("animeTags") if t not in EXCLUDED_TAGS])
                
                correct = set(song.get("correctGuessPlayers", []))
                self.song_history.append((correct, raw_f_players))
                ls                          =   song.get("listStates", [])
                self.global_stats["tot_c"]  +=  len(correct)
                
                yr = extract_year(si.get("vintage"))
                if yr is not None                                       : self.all_vint.append(yr)
                if isinstance(si.get("animeDifficulty"), (int, float))  : self.all_diff.append(si.get("animeDifficulty"))
                if not ls                                               : missing_list_count += 1

                s_riggers = {p["name"] for p in ls}
                if len(ls) == 1:
                    u                   =   ls[0]["name"]
                    self.p_l_solos[u]   +=  1
                    if not (len(correct) == 1 and list(correct)[0] == u): self.p_m_erigs[u] += 1

                if use_teams:
                    t_list = list({assignments[p.lower()][0] for p in raw_f_players if p.lower() in assignments})
                    if len(t_list) == 2:
                        tA, tB = t_list[0],             t_list[1]
                        cA, cB = correct & rosters[tA], correct & rosters[tB]

                        if len(cA) == 4 and not cB: self.t_sweeps[tA] += 1; self.global_stats["sweeps"] += 1
                        if len(cB) == 4 and not cA: self.t_sweeps[tB] += 1; self.global_stats["sweeps"] += 1

                        for cur, opp in [(tA, tB), (tB, tA)]:
                            cC, oC = correct & rosters[cur], correct & rosters[opp]
                            if not oC: 
                                for p in cC: self.p_pts[p] += 1
                            if len(cC) == 1 and len(oC) > 0: self.p_blks[list(cC)[0]] += 1
                    
                    for tid in t_list:
                        ros     = rosters[tid]
                        c_on_t  = correct & ros
                        self.t_c_ps[tid].append(len(c_on_t) / 4.0)
                        if yr is not None: self.t_vint[tid].append(yr)
                        if s_riggers & ros:
                            self.t_on_syn       [tid].append(len(c_on_t)                / 4.0)
                            self.t_sh_rig       [tid].append((len(s_riggers & ros) - 1) / 3.0)
                            self.t_overs        [tid].append((len(correct), len(s_riggers & ros)))
                        else: self.t_off_syn    [tid].append(len(c_on_t)                / 4.0)

                if len(final_members - correct) == 0: self.global_stats["fulls"] += 1
                elif apply_rev and len(final_members - correct) == 1:
                    self.global_stats   ["sevens"]                          += 1
                    self.p_rev_e        [list(final_members - correct)[0]]  += 1
                elif len(correct) == 2:
                    self.global_stats   ["doubles"]     += 1
                    for p in correct: self.p_two_e[p]   += 1
                elif len(correct) == 1:
                    self.global_stats["solos"]                                              +=  1
                    sw                                                                      =   list(correct)[0]
                    self.e_counts[sw]                                                       +=  1
                    if sw.lower() in assignments: self.t_solos[assignments[sw.lower()][0]]  +=  1
                elif len(correct) == 0: self.global_stats["blanks"] += 1

                for name in final_members:
                    if name in correct:
                        self                        .c_counts[name]     += 1
                        if st in [1, 2, 3]  : self  .p_type_c[name][st] += 1
                        if is_chan          : self  .p_chan_c[name]     += 1
                    if is_chan: self                .p_chan_s[name]     += 1

                if ls:
                    for p in ls:
                        n = p["name"]       ; self.p_rigs   [n] += 1
                        if n in correct     : self.p_rigs_h [n] += 1
                        if yr is not None   : self.p_l_vint [n].append(yr)
                        self.p_l_corr[n].append(len(correct))

        self._finalize_outputs(missing_list_count, appearances, use_teams, elo_map, assignments, t1_lookup, tour_types)

    def _scan_players(self, paths):
        players = set           ()
        apps    = defaultdict   (set)

        for p in paths:
            try:
                with open(p, encoding = "utf-8") as f:
                    data = json.load(f)
                    for s in data.get("songs", []):
                        for plyr    in s.get("correctGuessPlayers", []): players.add(plyr);         apps[plyr]          .add(str(p))
                        for ls      in s.get("listStates",          []): players.add(ls["name"]);   apps[ls["name"]]    .add(str(p))
            except: continue
        return players, apps

    def _load_team_data(self, all_known):
        codes = self.tour_dir / FILE_CODES
        if not codes.exists(): return False, {}, {}, {}, defaultdict(set)
        
        self.main_roster_names                      = set()
        elo_map, assignments, rosters, t1_lookup    = {}, {}, defaultdict(set), {}
        avail                                       = sorted(list(all_known)) 
        alias_path                                  = self.script_dir / DIR_TOURS / FILE_ALIASES
        local_aliases                               = {}

        if alias_path.exists():
            with open(alias_path, "r", encoding = "utf-8") as f:
                for line in f:
                    if "," in line:
                        k, v = line.strip().split(",", 1)
                        local_aliases[k.strip().lower()] = v.strip()
        
        new_aliases = {}

        def find_best_match(p_in, allow_manual = False, line_text = ""):
            p_low = p_in.lower()

            if p_low in local_aliases:
                m = local_aliases[p_low]
                if m in all_known: return m

            match = next((n for n in all_known if n.lower() == p_low), None)

            if not match:
                if not self.id_database: self.id_database = self._load_player_ids()

                if p_low in self.id_database:
                    target_id   = self.id_database[p_low]
                    match       = next((n for n in all_known if self.id_database.get(n.lower()) == target_id), None)

            if not match and allow_manual and ("[" in line_text or "Subs:" in line_text)    : match = ManualMatchDialog(None, p_in, avail).result
            if match                                                                        : new_aliases[p_in] = match

            return match

        with open(codes, "r", encoding = "utf-8") as f: lines = f.readlines()

        for line in lines:
            matches = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', line)
            for p_in, val in matches:
                match = find_best_match(p_in, allow_manual = True, line_text = line)
                if match: elo_map[match.lower()] = val

        idx = 1

        for line in lines:
            if "Subs:" in line or "subs:" in line:
                mems_subs = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', line)

                for p_sub, _ in mems_subs:
                    m_sub = find_best_match(p_sub)
                    if m_sub: self.subbed_players_set.add(m_sub.lower())

            if "[" not in line: continue

            pre     = re.match(r'^(?:\\s*)?([^:\[\d\(]+)\s*\([\d.-]+\):', line)
            ename   = pre.group(1).strip() if pre else None
            sec     = line.split(":", 1)[1] if ":" in line else line
            mems    = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', sec)

            for i, (p_in, _) in enumerate(mems[:4]):
                tier        = str(i + 1)
                match       = find_best_match(p_in)

                if match:
                    self.main_roster_names.add(match.lower())
                    assignments[match.lower()] = (idx, tier)
                    rosters[idx].add(match)
                    if match in avail: avail.remove(match)
                    t1_lookup[idx] = ename if ename else (match if tier == "1" else t1_lookup.get(idx))

            idx += 1

        if new_aliases:
            existing_entries = []

            if alias_path.exists():
                with open(alias_path, "r", encoding = "utf-8") as f: existing_entries = [l.strip().lower() for l in f.readlines()]

            with open(alias_path, "a", encoding = "utf-8") as f:
                for k, v in new_aliases.items():
                    entry = f"{k}, {v}".lower()

                    if entry not in existing_entries:
                        f.write(f"{k}, {v}\n")
                        existing_entries.append(entry)

        return True, elo_map, assignments, t1_lookup, rosters, all_known

    def _finalize_outputs(self, missing_count, appearances, use_teams, elo_map, assignments, t1_lookup, tour_types):
        watched_valid       = missing_count <= 5
        baseline_initial    = int(np.median([len(appearances.get(name, [])) for name in self.s_part]))
        
        if len(tour_types) == 1:
            t_map           = {1: "OP", 2: "ED", 3: "IN"}
            t_str           = t_map.get(list(tour_types)[0], "")
            init_label      = f"Watched {t_str}" if watched_valid else f"Random {t_str}"
        else: init_label    = "Watched" if watched_valid else "Usual"

        if "Eru" in init_label and use_teams                    : default_th = ""
        else:
            if      init_label == "Watched 2+8"                 : default_th = "25, 20, 15, 10, 5"
            elif    init_label in ["Watched",   "QuagWatched"]  : default_th = "28, 18, 12, 6"
            elif    init_label in ["Usual",     "Quagsual"]     : default_th = "28, 19, 8"
            elif    "Rigs" in self.s_part                       : default_th = "28, 18, 12, 6"
            else                                                : default_th = "28, 19, 8"

        meta_dialog     = TourMetadataDialog(root, self.tour_id, init_label, default_th, baseline_initial, list(self.s_part.keys()))
        meta_res        = meta_dialog.result if meta_dialog.result else {"tour_label": init_label, "th_str": "default", "base_exp": baseline_initial, "selected_new": []}
        self.tour_label = meta_res["tour_label"]

        if not self.tour_label: self.tour_label = init_label
        
        val_str     = meta_res["th_str"]
        base_exp    = meta_res["base_exp"]
        new_players = meta_res["selected_new"]

        if "Eru" in self.tour_label and use_teams:
            self.p_pts  .clear()
            self.p_blks .clear()

            for cor, raw_p in self.song_history:
                t_list = list({assignments[p.lower()][0] for p in raw_p if p.lower() in assignments})

                if len(t_list) == 2:
                    tA, tB  = t_list[0], t_list[1]
                    cA      = {assignments[p.lower()][1]: p for p in raw_p if p.lower() in assignments and assignments[p.lower()][0] == tA}
                    cB      = {assignments[p.lower()][1]: p for p in raw_p if p.lower() in assignments and assignments[p.lower()][0] == tB}

                    for tr in ["1", "2", "3", "4"]:
                        pA, pB = cA.get(tr), cB.get(tr)

                        if pA and pB:
                            rA, rB = pA in cor, pB in cor

                            if rA and not rB: self.p_pts[pA] += 1
                            if rB and not rA: self.p_pts[pB] += 1

                            if rA and rB:
                                self.p_blks[pA] += 0.50
                                self.p_blks[pB] += 0.50

        t_name              = self.tour_label.strip()
        tour_disp           = f"{t_name} Tour"    
        exp_map             = {}
        mismatched_players  = {}

        for name in list(self.s_part.keys()):
            act = len(appearances.get(name, []))

            if act < base_exp   : mismatched_players[name]  = act
            else                : exp_map[name]             = base_exp

        if mismatched_players:
            mismatch_dialog = MismatchedRoundsDialog(root, mismatched_players, base_exp, self.subbed_players_set)
            mismatch_res    = mismatch_dialog.result if mismatch_dialog.result else {k: base_exp for k in mismatched_players}

            for name, target in mismatch_res.items():
                act             = len(appearances.get(name, []))
                exp_map[name]   = target

                if target > act:
                    avg_songs_per_json  =   sum(self.s_part.values()) / sum(len(v) for v in appearances.values())
                    missing_rounds      =   target - act
                    self.s_part[name]   +=  int(missing_rounds * avg_songs_per_json)

        final_threshold = 6 if len(self.s_part) <= 20 else 5

        if      base_exp >= final_threshold     : stage = "Final"
        elif    base_exp == 3                   : stage = "Mid-Tour"
        else                                    : stage = f"R{base_exp}"

        prefix      = f"{tour_disp}, " 
        out_path    = self.tour_dir / DIR_OUT
        out_path.mkdir(parents = True, exist_ok = True)

        self._create_player_png (use_teams, elo_map, watched_valid, stage, out_path, appearances, prefix, exp_map, base_exp, assignments, new_players, t1_lookup, val_str)
        self._create_tour_png   (use_teams, watched_valid, out_path)

        if watched_valid and assignments        : self._create_team_png     (t1_lookup,     out_path)
        if assignments                          : self._create_tier_png     (assignments,   out_path,   watched_valid)
        if watched_valid                        : self._create_watched_png  (out_path, assignments, t1_lookup)
        if watched_valid and self.chanting_ids  : self._create_chanting_png (out_path)

        if watched_valid: self._fuse_and_clean(out_path)
        else:
            f = {"Tour": "Tour.png", "Team": "Team.png", "Tier": "Tier.png", "List": "List.png", "Chanting": "Chanting.png"}

            for v in f.values():
                p = out_path / v

                if p.exists():
                    try     : os.remove(p)
                    except  : pass

        messagebox.showinfo("Success", f"Saved the PNGs for the {t_name} tour to {DIR_OUT}/{self.tour_id}")

    def _create_player_png(self, use_teams, elo_map, watched, stage, path, apps, prefix, exp_map, base_exp, assigns, new_players, t1_lookup, val_str):
        rows, eligibility   = [], []
        t_labels            = {1: "OP GR", 2: "ED GR", 3: "IN GR"}
        active              = [t for t in [1, 2, 3] if any(self.p_type_s[p][t] > 0 for p in self.s_part)]
        if len(active) <= 1 : active = []

        for name in self.s_part:
            tot, cor    = self.s_part[name], self.c_counts[name]
            target      = exp_map.get(name, base_exp)
            d_name      = name

            if name in new_players: d_name += " ☆"

            if target < base_exp:
                if name.lower() in self.main_roster_names   : d_name += " ▼"
                else                                        : d_name += " ▲"
            
            is_eligible = not ("▼" in d_name or "▲" in d_name)
            eligibility.append(is_eligible)
            
            act = len(apps.get(name, []))
            if act < target:
                syms = ["", "(1)", "(2)", "(3)", "(4)", "(5)", "(6)"]
                if 0 < (target-act) < len(syms): d_name += f" {syms[target-act]}"

            row = {"Player": d_name}
            
            if use_teams:
                team_info = assigns.get(name.lower(), ("N/A", "N/A"))
                if team_info[0] != "N/A":
                    leader_name = t1_lookup.get(team_info[0], "")
                    clean_name  = "".join(filter(str.isalnum, leader_name))
                    row["Team"] = clean_name[:3].upper() if leader_name else f"T{team_info[0]}"
                    row["Tier"] = team_info[1]
                else:
                    row["Team"] = "N/A"
                    row["Tier"] = "N/A"
                row["Elo"]      = elo_map.get(name.lower(), "N/A")

            row.update({
                "Guess Rate"    : cor / tot if tot else 0,
                "1/8s"          : self.e_counts [name],
                "2/8s"          : self.p_two_e  [name],
                "7/8s"          : self.p_rev_e  [name]
            })

            if use_teams: row.update({"Lives Taken": self.p_pts[name], "Lives Saved": self.p_blks[name]})
            
            for tid in active:
                seen                = self.p_type_s[name][tid]
                row[t_labels[tid]]  = self.p_type_c[name][tid] / seen if seen else np.nan

            if watched:
                row.update({
                    "Rigs"              : self.p_rigs[name],
                    "Rig Rate"          : self.p_rigs[name]             / tot                       if tot                          else np.nan,
                    "Rig Delta"         : (cor - self.p_rigs[name])     / cor                       if cor                          else np.nan,
                    "Rig GR"            : self.p_rigs_h[name]           / self.p_rigs[name]         if self.p_rigs[name]            else np.nan,
                    "Off GR"            : (cor - self.p_rigs_h[name])   / (tot - self.p_rigs[name]) if (tot - self.p_rigs[name])    else np.nan,
                })

            rows.append(row)

        df      = pd.DataFrame(rows).sort_values("Guess Rate", ascending = False)
        mask    = pd.Series(eligibility, index = pd.DataFrame(rows).index).reindex(df.index).values
        pcts    = ["Guess Rate"] + [t_labels[t] for t in active] + (["Rig Rate", "Rig Delta", "Rig GR", "Off GR"] if watched else [])

        if "Elo" in df.columns: df["Elo"] = pd.to_numeric(df["Elo"], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        for c in pcts: df[c] = pd.to_numeric(df[c], errors = 'coerce').mul(100).map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        self._export_png(df, path, "Player.png", f"{prefix}Player Statistics, {stage}", mask, val_str)

    def _create_tour_png(self, use_teams, watched, path):
        def fmt_most(names, val):
            if not names: return "N/A"
            win = sorted(names, key = lambda x: (self.c_counts[x] / self.s_part[x]) if self.s_part[x] else 0)[0]
            gr  = (self.c_counts[win] / self.s_part[win]) * 100 if self.s_part[win] else 0
            return f"{win} ({val}{f', {gr:.2f}' if len(names) > 1 else ''})"

        stats = [
            ["Median Vintage",      format_year(round(np.median(self.all_vint), 2))                         if self.all_vint    else "N/A"],
            ["Average Difficulty",  f"{np.mean(self.all_diff):.2f}"                                         if self.all_diff    else "N/A"],
            ["Average GR",          f"{100 * (self.global_stats['tot_c'] / sum(self.s_part.values())):.2f}" if self.s_part      else "0.00"],
            ["Total 0/8s",          self.global_stats["blanks"]],
            ["Total 1/8s",          self.global_stats["solos"]],
            ["Total 2/8s",          self.global_stats["doubles"]],
            ["Total 7/8s",          self.global_stats["sevens"]],
            ["Total 8/8s",          self.global_stats["fulls"]]
        ]

        if use_teams: stats.append(["Total 4-0s", self.global_stats["sweeps"]])
        stats.extend([
            ["Most Popular Genre",  f"{self.genre_c .most_common(1)[0][0]} ({self.genre_c   .most_common(1)[0][1]})" if self.genre_c else "N/A"],
            ["Most Popular Tag",    f"{self.tag_c   .most_common(1)[0][0]} ({self.tag_c     .most_common(1)[0][1]})" if self.tag_c else "N/A"],
            ["Most 1/8s",           fmt_most([n for n, v in self.e_counts   .items() if v == max(self.e_counts  .values(), default = 0) and v > 0], max(self.e_counts   .values(), default = 0))],
            ["Most 2/8s",           fmt_most([n for n, v in self.p_two_e    .items() if v == max(self.p_two_e   .values(), default = 0) and v > 0], max(self.p_two_e    .values(), default = 0))],
            ["Most 7/8s",           fmt_most([n for n, v in self.p_rev_e    .items() if v == max(self.p_rev_e   .values(), default = 0) and v > 0], max(self.p_rev_e    .values(), default = 0))]
        ])
        
        plist   = list(self.s_part.keys())
        no_s    = sorted([n for n in plist if self.e_counts[n] ==   0 and self.s_part[n] > 0], key = lambda x: self.c_counts[x] / self.s_part[x], reverse = True)
        yes_s   = sorted([n for n in plist if self.e_counts[n] >    0 and self.s_part[n] > 0], key = lambda x: self.c_counts[x] / self.s_part[x])

        if no_s     : stats.append(["Highest GR Without 1/8s",  f"{no_s     [0]} ({100 * (self.c_counts[no_s    [0]] / self.s_part[no_s     [0]]):.2f})"])
        if yes_s    : stats.append(["Lowest GR With 1/8s",      f"{yes_s    [0]} ({100 * (self.c_counts[yes_s   [0]] / self.s_part[yes_s    [0]]):.2f}, {self.e_counts[yes_s[0]]})"])

        if watched:
            conv        = []
            eligible    = [p for p in plist if self.p_l_solos[p] > 0]
            
            if eligible:
                total_hits      = sum((self.p_l_solos[p] - self.p_m_erigs[p]) for p in eligible)
                total_attempts  = sum(self.p_l_solos[p] for p in eligible)
                global_avg      = total_hits / total_attempts if total_attempts > 0 else 0
                constant        = 3
                
                for n in eligible:
                    t               = self.p_l_solos[n]
                    h               = t - self.p_m_erigs[n]
                    weighted_score  = (h + constant * global_avg) / (t + constant)
                    conv.append({'n': n, 'score': weighted_score, 'p': 100 * h / t, 'h': h, 't': t})

                b = sorted(conv, key = lambda x: x['score'], reverse = True)    [0]
                w = sorted(conv, key = lambda x: x['score'])                    [0]

                stats.append(["Best Solo Rig Converter",    f"{b['n']} ({b['p']:.2f}%, {b['h']}/{b['t']})"])
                stats.append(["Worst Solo Rig Converter",   f"{w['n']} ({w['p']:.2f}%, {w['h']}/{w['t']})"])

        self._export_png(pd.DataFrame(stats, columns = ["Statistic", "Value"]), path, "Tour.png", "Tour Statistics")

    def _create_team_png(self, t1_lookup, path):
        res = []

        for tid in self.t_c_ps:
            leader_name = t1_lookup.get(tid, "")
            clean_name  = "".join(filter(str.isalnum, leader_name))
            t_lbl       = clean_name[:3].upper() if leader_name else f"T{tid}"
            
            res.append({
                "Team"              : t_lbl,
                "Median Vintage"    : format_year(np.median(self.t_vint[tid])),
                "Average GR"        : f"{np.mean(self.t_c_ps    [tid]) * 100:.2f}",
                "Rig Synergy"       : f"{np.mean(self.t_on_syn  [tid]) * 100:.2f}",
                "Off Synergy"       : f"{np.mean(self.t_off_syn [tid]) * 100:.2f}",
                "Shared Rigs"       : f"{np.mean(self.t_sh_rig  [tid]) * 100:.2f}",
                "Total 1/8s"        : self.t_solos[tid],
            })
        self._export_png(pd.DataFrame(res).sort_values("Average GR", ascending = False), path, "Team.png", "Team Statistics")

    def _create_tier_png(self, assigns, path, watched_valid):
        tiers   = sorted({v[1] for v in assigns.values() if v[1] != "N/A"}, reverse = True)
        res     = []

        for tr in tiers:
            tp = [n for n in self.s_part if n.lower() in assigns and assigns[n.lower()][1] == tr]
            if not tp: continue

            def get_generalist(plist):
                sorted_p        = sorted(plist, key = lambda x: (self.c_counts[x] / self.s_part[x] if self.s_part[x] else 0), reverse = True)
                name, value     = sorted_p[0], 100 * (self.c_counts[sorted_p[0]] / self.s_part[sorted_p[0]]) if self.s_part[sorted_p[0]] else 0
                return f"{name} ({value:.2f})"

            def get_attblk(plist, sdict):
                sorted_p        = sorted(plist, key = lambda x: (sdict[x], self.c_counts[x] / self.s_part[x] if self.s_part[x] else 0), reverse = True)
                name, value     = sorted_p[0], sdict[sorted_p[0]]
                return f"{name} ({value})"

            def get_contributor(plist):
                sorted_p        = sorted(plist, key = lambda x: ((self.p_pts[x] + self.p_blks[x]), (self.c_counts[x] / self.s_part[x] if self.s_part[x] else 0)), reverse = True)
                name, v1, v2    = sorted_p[0], self.p_pts[sorted_p[0]], self.p_blks[sorted_p[0]]
                return f"{name} ({v1 + v2})"

            def get_chanter(plist):
                pool = [n for n in plist if self.p_chan_s[n] > 0 and self.c_counts[n] > 0]
                if not pool: return "N/A"
                sorted_p    = sorted(pool, key = lambda x: (100 * self.p_chan_c[x] / self.p_chan_s[x], -(100 * self.c_counts[x] / self.s_part[x])), reverse = True)
                name        = sorted_p[0]                
                ratio       = 100 * self.p_chan_c[name] / self.p_chan_s[name]
                return f"{name} ({ratio:.2f})"

            row_data = [tr, get_generalist(tp), get_attblk(tp, self.p_pts), get_attblk(tp, self.p_blks), get_contributor (tp)]
            if watched_valid: row_data.append(get_chanter(tp))
            res.append(row_data)
            
        cols = ["Tier", "Generalist", "Attacker", "Blocker", "Contributor"]
        if watched_valid: cols.append("Chanter")
        self._export_png(pd.DataFrame(sorted(res, key = lambda x: x[0]), columns = cols), path, "Tier.png", "Tier Statistics")

    def _create_watched_png(self, path, assigns, t1_lookup):
        plist = [n for n in self.s_part if self.p_l_corr[n]]
        if not plist: return

        x_vals = [np.mean   (self.p_l_corr[name]) for name in plist]
        y_vals = [np.median (self.p_l_vint[name]) if self.p_l_vint[name] else np.nan for name in plist]

        valid_data = [(p, x, y) for p, x, y in zip(plist, x_vals, y_vals) if not np.isnan(y)]
        if not valid_data: return

        plist, x_vals, y_vals   = zip   (*valid_data)
        plist                   = list  (plist)
        x_vals                  = list  (x_vals)
        y_vals                  = list  (y_vals)

        rig_rates   = [self.p_rigs      [name] / self.s_part[name] if self.s_part[name] else 0 for name in plist]
        rig_grs     = [self.p_rigs_h    [name] / self.p_rigs[name] if self.p_rigs[name] else 0 for name in plist]

        fig, ax     = plt.subplots(figsize = (10, 10))
        cmap        = mcolors.LinearSegmentedColormap.from_list("rig_gr_cmap", [(0.0, "#D95400"), (0.5, "#D95400"), (RIG_GR_THRESHOLD, "#FFFFFF"), (1.0, "#0056B3")])
        sizes       = [rate ** 2 * 10000 for rate in rig_rates]
        sc          = ax.scatter(x_vals, y_vals, s = sizes, c = rig_grs, cmap = cmap, vmin = 0.0, vmax = 1.0, edgecolors = 'black', alpha = 0.9)

        x_min       = math.floor    ((min   (x_vals) - 0.50) * 2) / 2
        x_max       = math.ceil     ((max   (x_vals) + 0.50) * 2) / 2
        y_min       = math.floor    (min    (y_vals) - 1.00)
        y_max       = math.ceil     (max    (y_vals) + 1.00)

        while True:
            r = y_max - y_min

            if r % 4 == 0:
                step = r // 4
                break
            elif r % 3 == 0:
                step = r // 3
                break

            if r % 2 != 0 or y_max >= 2026  : y_min -= 1
            else                            : y_max += 1

        ax.set_xlim     (x_min, x_max)
        ax.set_ylim     (y_min, y_max)

        ax.set_xticks   (np.arange  (x_min, x_max + 0.5,    0.5))
        ax.set_yticks   (range      (y_min, y_max + 1,      step))

        texts = []

        for name, x, y in zip(plist, x_vals, y_vals):
            label       = ""
            team_info   = assigns.get(name.lower(), ("N/A", "N/A"))

            if team_info[0] != "N/A":
                leader_name = t1_lookup.get(team_info[0], "")
                clean_name  = "".join(filter(str.isalnum, leader_name))
                t_lbl       = clean_name[ : 3].upper() if leader_name else f"T{team_info[0]}"
                label       = f"{t_lbl}-{team_info[1]}"
            
            if not label: continue
            t = ax.text(x, y, label, fontsize = 10, fontname = "Segoe UI")
            texts.append(t)

        if texts: adjust_text(
            texts, 
            ax                      = ax, 
            objects                 = sc, 
            avoid_self              = True, 
            add_objects_to_edges    = True, 
            force_text              = (1.00, 1.00), 
            force_objects           = (1.00, 1.00), 
            expand                  = (2.00, 2.00), 
            arrowprops              = dict(arrowstyle = "-", color = 'black', shrinkA = 10)
        )

        ax          .set_title              ("List Statistics", weight = 'bold', fontname = "Segoe UI", fontsize = 22.5, pad      = 12.5)
        ax          .set_xlabel             ("Average Over-8",  weight = 'bold', fontname = "Segoe UI", fontsize = 15.0, labelpad = 2.5)
        ax          .set_ylabel             ("Median Vintage",  weight = 'bold', fontname = "Segoe UI", fontsize = 15.0, labelpad = 2.5)
        ax.yaxis    .set_major_formatter    (plt.FuncFormatter(lambda val, _: str(int(val))))
        plt         .setp                   (ax.get_yticklabels(), horizontalalignment = 'center', verticalalignment = 'center')
        ax          .tick_params            (axis = 'x', which = 'both', length = 0, pad = 5)
        ax          .tick_params            (axis = 'y', which = 'both', length = 0, pad = 15)
        
        cbar = fig.colorbar(sc, ax = ax, pad = 0.005, aspect = 40, ticks = [0.0, 0.5, RIG_GR_THRESHOLD, 1.0])

        cbar        .set_label          ("Rig GR", weight = 'bold', fontname = "Segoe UI", fontsize = 15, labelpad = -5)
        cbar.ax     .set_yticklabels    (['0', '50', f'{int(RIG_GR_THRESHOLD * 100)}', '100'])
        cbar.ax     .tick_params        (labelsize = 10, length = 0)

        ax.text(0.01, 0.99, "New\nHard", transform = ax.transAxes, color = "grey", fontsize = 10, va = "top",       ha = "left",    weight = "bold", alpha = 0.75)
        ax.text(0.99, 0.99, "New\nEasy", transform = ax.transAxes, color = "grey", fontsize = 10, va = "top",       ha = "right",   weight = "bold", alpha = 0.75)
        ax.text(0.01, 0.01, "Old\nHard", transform = ax.transAxes, color = "grey", fontsize = 10, va = "bottom",    ha = "left",    weight = "bold", alpha = 0.75)
        ax.text(0.99, 0.01, "Old\nEasy", transform = ax.transAxes, color = "grey", fontsize = 10, va = "bottom",    ha = "right",   weight = "bold", alpha = 0.75)

        ax  .grid           (False)
        plt .tight_layout   ()
        plt .savefig        (path / "List.png", dpi = 500)
        plt .close          (fig)

        try     : trim_whitespace(path / "List.png")
        except  : pass

    def _create_chanting_png(self, path):
        plist = [n for n in self.s_part if self.p_chan_s[n] > 0 and self.c_counts[n] > 0]
        if not plist: return

        def get_ratio   (p): return 100 * self.p_chan_c[p] / self.p_chan_s  [p]
        def get_gr      (p): return 100 * self.c_counts[p] / self.s_part    [p]

        best    = sorted(plist, key = lambda p: (get_ratio(p), -get_gr(p)), reverse = True) [ : 3]
        worst   = sorted(plist, key = lambda p: (get_ratio(p), -get_gr(p)))                 [ : 3]
        rows    = []

        for i in range(3):
            b_cell = "N/A"
            if i < len(best):
                p       = best[i]
                b_cell  = f"{p} ({get_ratio(p):.2f})"
            
            w_cell = "N/A"
            if i < len(worst):
                p       = worst[i]
                w_cell  = f"{p} ({get_ratio(p):.2f})"
            
            rows.append([f"{i + 1}", b_cell, w_cell])

        self._export_png(pd.DataFrame(rows, columns = ["Rank", "Best", "Worst"]), path, "Chanting.png", "Chanting Statistics")

    def _export_png(self, df, path, fname, title, mask = None, val_str = "default"):
        if not self.browser_path: return

        desc = [
            "Elo", 
            "Guess Rate", 
            "1/8s", 
            "2/8s", 
            "Rigs", 
            "Rig Delta", 
            "Lives Taken", 
            "Lives Saved", 
            "Rig Rate", 
            "OP GR", 
            "ED GR", 
            "IN GR", 
            "Rig GR", 
            "Off GR", 
            "Average GR", 
            "Rig Synergy", 
            "Off Synergy", 
            "Shared Rigs", 
            "Total 1/8s"
        ]

        asc     = ["7/8s"]
        rest    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs"]
        stats   = {}

        for col in df.columns:
            if col in desc or col in asc:
                num     = pd.to_numeric(df[col].astype(str).str.replace('%',''), errors = 'coerce')
                el_num  = num[mask].dropna() if mask is not None and col in rest else num.dropna()

                if not num.dropna().empty:
                    stats[col] = {
                        'max'   : num.dropna().max(),
                        'min'   : el_num.min() if not el_num.empty else None, 
                        's_max' : num.dropna().value_counts().get(num.dropna().max(), 0) <= 3, 
                        's_min' : el_num.value_counts().get(el_num.min(), 0) <= 3 if not el_num.empty else False
                    }

        df      = df.reset_index(drop = True)
        borders = []

        if "Guess Rate" in df.columns:
            if "Eru" in self.tour_label: th = []
            else:
                if val_str == "default":
                    if      self.tour_label == "Watched 2+8"                : th_val = "25, 20, 15, 10, 5"
                    elif    self.tour_label in ["Watched", "QuagWatched"]   : th_val = "28, 18, 12, 6"
                    elif    self.tour_label in ["Usual", "Quagsual"]        : th_val = "28, 19, 8"
                    elif    "Rigs"          in df.columns                   : th_val = "28, 18, 12, 6"
                    else                                                    : th_val = "28, 19, 8"
                else                                                        : th_val = val_str

                try         : th = [float(x.strip()) for x in th_val.split(",")] if th_val else []
                except      : th = [28.0, 18.0, 12.0, 6.0]
            
            gv = pd.to_numeric(df["Guess Rate"].astype(str).str.replace('%',''), errors = 'coerce').tolist()

            for t in th:
                f_idx = -1
                for i, v in enumerate(gv):
                    if pd.notnull(v) and v >= t: f_idx = i
                if f_idx != -1 and f_idx < len(df) - 1: borders.append(f_idx)

        html = f"<thead><tr>" + "".join([f"<th>{str(c).replace(' ','<br>')}</th>" for c in df.columns]) + "</tr></thead><tbody>"
        for idx, row in df.iterrows():
            b_s     =   "border-bottom: 3px solid black;" if idx in borders else ""
            html    +=  "<tr>"

            for i, (cname, cell) in enumerate(row.items()):
                style = [b_s] if b_s else []

                if cname in stats:
                    v = pd.to_numeric(str(cell).replace('%',''), errors = 'coerce')

                    if pd.notnull(v):
                        is_max, is_min  = (v == stats[cname]['max']) and stats[cname]['s_max'], (v == stats[cname]['min']) and stats[cname]['s_min']
                        elig            = True if mask is None or cname not in rest else mask[idx]

                        if cname in desc:
                            if      is_max          : style.append("color: #0056B3; font-weight: bold;")
                            elif    is_min and elig : style.append("color: #D95400; font-weight: bold;")
                        elif cname in asc:
                            if      is_max and elig : style.append("color: #D95400; font-weight: bold;")
                            elif    is_min          : style.append("color: #0056B3; font-weight: bold;")

                s_attr  =   f' style="{" ".join(style)}"' if style else ""
                cnt     =   f"<b>{cell}</b>" if i==0 else cell
                html    +=  f"<td{s_attr}>{cnt}</td>"

            html += "</tr>"

        full = f"<html><head><style>body {{font-family: 'Segoe UI', Arial, sans-serif; background: white; display: inline-block; margin: 0;}} h2 {{margin: 10px 0 10px 5px; font-size: 30px; text-align: center;}} table {{margin-left: 10px; border-collapse: collapse; width: auto;}} th {{font-weight: bold; font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}} td {{font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}}</style></head><body><h2>{title}</h2><table>{html}</table></body></html>"
        hti  = Html2Image(size = (max(2000, len(df.columns) * 120), max(2000, len(df) * 60)), browser_executable = self.browser_path, output_path = str(path), custom_flags = ['--log-level=3', '--silent'])

        hti.screenshot(html_str = full, save_as = fname)

        try     : trim_whitespace(path / fname)
        except  : pass

    def _fuse_and_clean(self, path):
        f       = {"Tour": "Tour.png", "Team": "Team.png", "Tier": "Tier.png", "List": "List.png", "Chanting": "Chanting.png"}
        ps      = {k: path / v      for k, v in f   .items() if (path / v).exists()}
        imgs    = {k: Image.open(v) for k, v in ps  .items()}

        if not imgs: return

        rk      = [k for k in ["Team", "Tier", "Chanting"] if k in imgs]
        tw, th  = (imgs["Tour"].width, imgs["Tour"].height) if "Tour" in imgs else (0, 0)
        
        if "List" in imgs:
            img_list        = imgs["List"]
            lw, lh          = img_list.width, img_list.height
            scale_f         = th / float(lh) if th else 1.0
            new_lw          = int(lw * scale_f)
            img_list        = img_list.resize((new_lw, th), Image.Resampling.LANCZOS)
            imgs["List"]    = img_list
            lw              = new_lw
        else: lw            = 0

        rw, rh = 0, 0

        if rk:
            if "Team" in imgs:
                rw  =   max(rw, imgs["Team"].width)
                rh  +=  imgs["Team"].height + 10

            if "Tier" in imgs:
                rw  =   max(rw, imgs["Tier"].width)
                rh  +=  imgs["Tier"].height + 10

            if "Chanting" in imgs:
                rw  =   max(rw, imgs["Chanting"].width)
                rh  +=  imgs["Chanting"].height + 10

            rh -= 10

        grid_w  = lw + (10 if lw and tw else 0) + tw + (10 if tw and rw else 0) + rw
        grid_h  = max(th, rh)
        total_w = grid_w
        total_h = grid_h
        fused   = Image.new("RGB", (total_w, total_h), "white")
        cx, cy  = 0, 0

        if "List" in imgs:
            fused.paste(imgs["List"], (cx, 0))
            cx += lw + 10

        if "Tour" in imgs:
            fused.paste(imgs["Tour"], (cx, 0))
            cx += tw + 10

        if "Team" in imgs:
            fused.paste(imgs["Team"], (cx, cy))
            cy += imgs["Team"].height + 10

        if "Tier" in imgs:
            fused.paste(imgs["Tier"], (cx, cy))
            cy += imgs["Tier"].height + 10

        if "Chanting" in imgs:
            fused.paste(imgs["Chanting"], (cx, cy))
            cy += imgs["Chanting"].height + 10
            
        if fused:
            f_p = path / "Extra.png"
            fused.save(f_p)

            try     : trim_whitespace(f_p)
            except  : pass
            
        for p in ps.values():
            try     : os.remove(p)
            except  : pass

if __name__ == "__main__":
    root                = tk.Tk(); root.withdraw()
    selection_dialog    = TourSelectionDialog(root, ["0", "1", "2"])
    selected_tours      = selection_dialog.selected_tours
    
    if selected_tours:
        for tour_id in selected_tours: TourAnalyzer(tour_id).run()