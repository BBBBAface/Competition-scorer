# strategy_game.py (MODIFIED WITH ENHANCED SIMULATION)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import random
import math

class StrategyGameTab(ttk.Frame):
    """
    A ttk.Frame that contains the UI and logic for the 4X Strategy Game module.
    This tab allows users to manage factions, create tank designs, and calculate
    a final 'Battle Rating' based on combat strength, morale, and situational multipliers.
    It also includes a battle simulator.
    """
    DATA_FILE = "strategy_game_data.json"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # --- UI Style Fix ---
        # Improve readability of Combobox widgets which can be difficult to read on some systems.
        style = ttk.Style(self)
        style.map('TCombobox', fieldbackground=[('readonly','white')])
        style.map('TCombobox', selectbackground=[('readonly', 'white')])
        style.map('TCombobox', selectforeground=[('readonly', 'black')])


        self.factions = {}
        self.tank_designs = {}

        self.vars = {
            'tank_name': tk.StringVar(),
            'faction': tk.StringVar(),
            'gun_type': tk.StringVar(value='Turreted'),
            'upper_forward_armor': tk.DoubleVar(value=50), 'lower_forward_armor': tk.DoubleVar(value=50),
            'upper_side_armor': tk.DoubleVar(value=50), 'lower_side_armor': tk.DoubleVar(value=50),
            'upper_rear_armor': tk.DoubleVar(value=50), 'lower_rear_armor': tk.DoubleVar(value=50),
            'upper_armor': tk.DoubleVar(value=50), # Top plate armor
            'gun_pen': tk.DoubleVar(value=50),
            'ammo_caliber': tk.DoubleVar(value=105),       # Renamed from gun_caliber
            'ammo_length': tk.DoubleVar(value=300),         # New
            'gun_barrel_length': tk.DoubleVar(value=4000),  # New
            'gun_reload_speed': tk.DoubleVar(value=50), 'gun_mobility': tk.DoubleVar(value=50),
            'gun_traverse_range': tk.DoubleVar(value=50), 'turret_mobility': tk.DoubleVar(value=50),
            'turret_profile': tk.DoubleVar(value=50), 'mobility': tk.DoubleVar(value=50),
            'awareness': tk.DoubleVar(value=50), 'design_rating': tk.DoubleVar(value=50),
            'crew_comfort': tk.DoubleVar(value=50),
            'extra_crew': tk.IntVar(value=0),
            'soldiers': tk.IntVar(value=0),                  # New
            'tech_level': tk.IntVar(value=0),
            'gun_elevation': tk.DoubleVar(value=15),
            'gun_depression': tk.DoubleVar(value=5),
            'is_amphibious': tk.BooleanVar(value=False),
            'is_rushed': tk.BooleanVar(value=False),
            'has_dual_transmission': tk.BooleanVar(value=False),
            'tank_length': tk.DoubleVar(value=6000),
            'tank_width': tk.DoubleVar(value=3000),
            'tank_height': tk.DoubleVar(value=2500),
            'belly_armor': tk.DoubleVar(value=20),
            'is_open_top': tk.BooleanVar(value=False),
            'has_spaced_armor': tk.BooleanVar(value=False),
            'is_stationary': tk.BooleanVar(value=False),
            'has_secondary_gun': tk.BooleanVar(value=False), # New
            'secondary_weapon': tk.DoubleVar(value=50),      # New
            'engine_position': tk.DoubleVar(value=50),       # New
            'tonnage': tk.DoubleVar(value=50),               # New
            'is_overworked': tk.BooleanVar(value=False),     # New
            'combat_strength': tk.StringVar(value="--"), 'morale_score': tk.StringVar(value="--"),
            'interaction_details': tk.StringVar(value="Bonuses/Penalties will appear here."),
            'battle_rating': tk.StringVar(value="--")
        }
        self.vars['gun_type'].trace_add("write", self.on_gun_type_change)
        self.vars['is_open_top'].trace_add("write", self.on_open_top_change)
        self.vars['is_stationary'].trace_add("write", self.on_stationary_change)
        self.vars['has_secondary_gun'].trace_add("write", self.on_secondary_gun_change)

        self.faction_vars = {
            'name': tk.StringVar(), 'leadership': tk.DoubleVar(value=50),
            'stability': tk.DoubleVar(value=50), 'wealth': tk.DoubleVar(value=50)
        }
        
        self.ui_elements = {}

        self.create_widgets()
        self.load_data_from_file()
        self.on_gun_type_change()
        self.on_open_top_change()
        self.on_stationary_change()
        self.on_secondary_gun_change()


    def create_widgets(self):
        top_frame = ttk.Frame(self, padding=(10, 10, 10, 0))
        top_frame.pack(fill=tk.X)
        ttk.Button(top_frame, text="💾 Save All Data", command=self.save_data_to_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="📂 Load All Data", command=self.load_data_from_file).pack(side=tk.LEFT)
        ttk.Button(top_frame, text="⚔️ Commence Battle!", command=self.open_battle_setup).pack(side=tk.RIGHT, padx=5)

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(expand=True, fill=tk.BOTH)
        main_frame.columnconfigure(0, weight=1, minsize=300)
        main_frame.columnconfigure(1, weight=2, minsize=500)
        main_frame.columnconfigure(2, weight=1, minsize=300)

        self.create_faction_widgets(ttk.LabelFrame(main_frame, text="Faction Management", padding=10))
        self.create_designer_widgets(ttk.LabelFrame(main_frame, text="Tank Designer", padding=10))
        self.create_saved_designs_widgets(ttk.LabelFrame(main_frame, text="Saved Tank Designs", padding=10))

    def create_faction_widgets(self, parent):
        parent.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        parent.rowconfigure(2, weight=1); parent.columnconfigure(0, weight=1)
        editor = ttk.Frame(parent); editor.grid(row=0, column=0, sticky="ew"); editor.columnconfigure(1, weight=1)
        ttk.Label(editor, text="Name:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(editor, textvariable=self.faction_vars['name']).grid(row=0, column=1, sticky="ew")
        self.create_slider_entry(editor, "Leadership:", self.faction_vars['leadership'], 1, tooltip="The quality and experience of the faction's commanders.")
        self.create_slider_entry(editor, "Stability:", self.faction_vars['stability'], 2, tooltip="The faction's political and social cohesion.")
        self.create_slider_entry(editor, "Wealth:", self.faction_vars['wealth'], 3, tooltip="The faction's economic strength and industrial capacity.")
        buttons = ttk.Frame(parent); buttons.grid(row=1, column=0, pady=10)
        ttk.Button(buttons, text="Add/Update", command=self.add_update_faction).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons, text="Delete", command=self.delete_faction).pack(side=tk.LEFT)
        
        self.faction_listbox = tk.Listbox(parent, selectmode=tk.SINGLE, exportselection=False)
        self.faction_listbox.grid(row=2, column=0, sticky="nsew")
        self.faction_listbox.bind("<<ListboxSelect>>", self.on_faction_select)

    def create_designer_widgets(self, parent):
        parent.grid(row=0, column=1, sticky="nsew", padx=5)
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        info_frame = ttk.Frame(parent); info_frame.grid(row=0, column=0, sticky="ew"); info_frame.columnconfigure(1, weight=1)
        ttk.Label(info_frame, text="Tank Name:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(info_frame, textvariable=self.vars['tank_name']).grid(row=0, column=1, sticky="ew")
        ttk.Label(info_frame, text="Faction:").grid(row=1, column=0, sticky="w", pady=2)
        self.faction_combo = ttk.Combobox(info_frame, textvariable=self.vars['faction'], state="readonly")
        self.faction_combo.grid(row=1, column=1, sticky="ew")
        
        notebook = ttk.Notebook(parent); notebook.grid(row=1, column=0, sticky="nsew", pady=5)
        chassis_tab = ttk.Frame(notebook, padding=10); armor_tab = ttk.Frame(notebook, padding=10)
        weapon_tab = ttk.Frame(notebook, padding=10); crew_tab = ttk.Frame(notebook, padding=10)
        notebook.add(chassis_tab, text="Chassis"); notebook.add(armor_tab, text="Armor");
        notebook.add(weapon_tab, text="Weapon"); notebook.add(crew_tab, text="Crew & Soft Factors")
        
        chassis_tab.columnconfigure(1, weight=1)
        self.ui_elements['mobility_widgets'] = self.create_slider_entry(chassis_tab, "Mobility:", self.vars['mobility'], 0, tooltip="The tank's top speed and acceleration.")
        self.create_slider_entry(chassis_tab, "Awareness:", self.vars['awareness'], 1, tooltip="The crew's ability to spot enemies, including optics quality.")
        ttk.Label(chassis_tab, text="Tech Level:").grid(row=2, column=0, sticky="w", pady=2)
        tech_spinbox = ttk.Spinbox(chassis_tab, from_=0, to=5, textvariable=self.vars['tech_level'], width=5)
        tech_spinbox.grid(row=2, column=1, sticky="w")
        ToolTip(tech_spinbox, "The technological advancement of the tank's components (0-5).")
        self.create_slider_entry(chassis_tab, "Tonnage:", self.vars['tonnage'], 3, from_=0, to=300, tooltip="The overall weight of the tank in tons.")
        self.create_slider_entry(chassis_tab, "Engine Position:", self.vars['engine_position'], 4, from_=0, to=100, tooltip="Location of the engine. 0=Front, 50=Mid, 100=Rear.")

        checkbox_frame = ttk.LabelFrame(chassis_tab, text="Special Properties", padding=5)
        checkbox_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)
        cb1 = ttk.Checkbutton(checkbox_frame, text="Stationary Gun", variable=self.vars['is_stationary']); cb1.pack(side=tk.LEFT, padx=5); ToolTip(cb1, "Is the gun fixed in a fortress or bunker?")
        cb2 = ttk.Checkbutton(checkbox_frame, text="Amphibious", variable=self.vars['is_amphibious']); cb2.pack(side=tk.LEFT, padx=5); ToolTip(cb2, "Can the tank cross deep water?")
        cb3 = ttk.Checkbutton(checkbox_frame, text="Dual Transmission", variable=self.vars['has_dual_transmission']); cb3.pack(side=tk.LEFT, padx=5); ToolTip(cb3, "are there 2 transmissions?")
        cb4 = ttk.Checkbutton(checkbox_frame, text="Rushed Production", variable=self.vars['is_rushed']); cb4.pack(side=tk.LEFT, padx=5); ToolTip(cb4, "Was the tank built hastily? (Applies a large BR penalty)")

        armor_tab.columnconfigure(1, weight=1)
        self.create_slider_entry(armor_tab, "Upper Fwd Armor:", self.vars['upper_forward_armor'], 0, tooltip="Armor thickness on the upper front plate.")
        self.create_slider_entry(armor_tab, "Lower Fwd Armor:", self.vars['lower_forward_armor'], 1, tooltip="Armor thickness on the lower front plate.")
        self.create_slider_entry(armor_tab, "Upper Side Armor:", self.vars['upper_side_armor'], 2, tooltip="Armor thickness on the upper side.")
        self.create_slider_entry(armor_tab, "Lower Side Armor:", self.vars['lower_side_armor'], 3, tooltip="Armor thickness on the lower side.")
        self.create_slider_entry(armor_tab, "Upper Rear Armor:", self.vars['upper_rear_armor'], 4, tooltip="Armor thickness on the upper rear.")
        self.create_slider_entry(armor_tab, "Lower Rear Armor:", self.vars['lower_rear_armor'], 5, tooltip="Armor thickness on the lower rear.")
        self.ui_elements['upper_armor_widgets'] = self.create_slider_entry(armor_tab, "Top Armor:", self.vars['upper_armor'], 6, tooltip="Armor thickness on the roof.")
        self.create_slider_entry(armor_tab, "Turret Profile:", self.vars['turret_profile'], 7, tooltip="Lower is better (represents height/visibility of the turret).")
        self.create_slider_entry(armor_tab, "Belly Armor:", self.vars['belly_armor'], 8, tooltip="Resistance to mines and underbelly explosions.")

        dims_frame = ttk.LabelFrame(armor_tab, text="Dimensions (mm)", padding=5)
        dims_frame.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(10,0))
        dims_frame.columnconfigure(1, weight=1); dims_frame.columnconfigure(3, weight=1); dims_frame.columnconfigure(5, weight=1)
        ttk.Label(dims_frame, text="Length:").grid(row=0, column=0, padx=(5,2))
        ttk.Entry(dims_frame, textvariable=self.vars['tank_length'], width=8).grid(row=0, column=1)
        ttk.Label(dims_frame, text="Width:").grid(row=0, column=2, padx=(10,2))
        ttk.Entry(dims_frame, textvariable=self.vars['tank_width'], width=8).grid(row=0, column=3)
        ttk.Label(dims_frame, text="Height:").grid(row=0, column=4, padx=(10,2))
        ttk.Entry(dims_frame, textvariable=self.vars['tank_height'], width=8).grid(row=0, column=5)

        armor_opts_frame = ttk.LabelFrame(armor_tab, text="Armor Options", padding=5)
        armor_opts_frame.grid(row=10, column=0, columnspan=3, sticky="ew", pady=10)
        cb5 = ttk.Checkbutton(armor_opts_frame, text="Open Top", variable=self.vars['is_open_top']); cb5.pack(side=tk.LEFT, padx=5); ToolTip(cb5, "Is the crew compartment open to the sky?")
        cb6 = ttk.Checkbutton(armor_opts_frame, text="Spaced Armor", variable=self.vars['has_spaced_armor']); cb6.pack(side=tk.LEFT, padx=5); ToolTip(cb6, "Does the tank have an extra layer of spaced armor?")

        weapon_tab.columnconfigure(1, weight=1)
        gun_type_frame = ttk.Frame(weapon_tab)
        gun_type_frame.grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(gun_type_frame, text="Turreted", variable=self.vars['gun_type'], value="Turreted").pack(side=tk.LEFT)
        ttk.Radiobutton(gun_type_frame, text="Fixed", variable=self.vars['gun_type'], value="Fixed").pack(side=tk.LEFT)
        ttk.Label(weapon_tab, text="Gun Type:").grid(row=0, column=0, sticky="w", pady=2)
        self.create_slider_entry(weapon_tab, "Gun Penetration:", self.vars['gun_pen'], 1, tooltip="The raw penetration power of the main gun.")
        self.create_slider_entry(weapon_tab, "Ammo Caliber (mm):", self.vars['ammo_caliber'], 2, from_=10, to=250, tooltip="The diameter of the ammunition.")
        self.create_slider_entry(weapon_tab, "Ammo Length (mm):", self.vars['ammo_length'], 3, from_=0, to=1200, tooltip="The length of the ammunition casing.")
        self.create_slider_entry(weapon_tab, "Barrel Length (mm):", self.vars['gun_barrel_length'], 4, from_=0, to=12000, tooltip="The length of the gun barrel.")
        self.create_slider_entry(weapon_tab, "Reload Speed:", self.vars['gun_reload_speed'], 5, tooltip="The time required to reload the main gun. Lower is better.")
        self.create_slider_entry(weapon_tab, "Gun Aim Speed:", self.vars['gun_mobility'], 6, tooltip="How quickly the gunner can aim the gun.")
        self.create_slider_entry(weapon_tab, "Turret Mobility:", self.vars['turret_mobility'], 7, tooltip="How quickly the turret can rotate.")
        self.create_slider_entry(weapon_tab, "Gun Traverse:", self.vars['gun_traverse_range'], 8, tooltip="The horizontal arc a Fixed Gun can cover.")
        self.create_slider_entry(weapon_tab, "Gun Depression:", self.vars['gun_depression'], 9, from_=0, to=45, tooltip="Degrees the gun can aim down.")
        self.create_slider_entry(weapon_tab, "Gun Elevation:", self.vars['gun_elevation'], 10, from_=0, to=45, tooltip="Degrees the gun can aim up.")
        
        sec_gun_frame = ttk.LabelFrame(weapon_tab, text="Secondary Weapon", padding=5)
        sec_gun_frame.grid(row=11, column=0, columnspan=3, sticky="ew", pady=10)
        sec_gun_frame.columnconfigure(1, weight=1)
        cb7 = ttk.Checkbutton(sec_gun_frame, text="Has Secondary Gun", variable=self.vars['has_secondary_gun']); cb7.grid(row=0, column=0, columnspan=3); ToolTip(cb7, "Does the tank have a secondary armament like a machine gun?")
        self.ui_elements['secondary_weapon_widgets'] = self.create_slider_entry(sec_gun_frame, "Effectiveness:", self.vars['secondary_weapon'], 1, tooltip="The overall combat effectiveness of the secondary weapon.")

        crew_tab.columnconfigure(1, weight=1)
        self.create_slider_entry(crew_tab, "Design Rating:", self.vars['design_rating'], 0, tooltip="The overall ergonomic and functional quality of the tank's design.")
        self.create_slider_entry(crew_tab, "Crew Comfort:", self.vars['crew_comfort'], 1, tooltip="How comfortable the tank is for the crew. Very low values can cause penalties.")
        
        ttk.Label(crew_tab, text="Extra Crew:").grid(row=2, column=0, sticky="w", pady=2)
        crew_spinbox = ttk.Spinbox(crew_tab, from_=0, to=10, textvariable=self.vars['extra_crew'], width=5)
        crew_spinbox.grid(row=2, column=1, sticky="w")
        ToolTip(crew_spinbox, "Number of additional crew members (e.g., loaders, commanders).")

        ttk.Label(crew_tab, text="Soldiers:").grid(row=3, column=0, sticky="w", pady=2)
        soldier_spinbox = ttk.Spinbox(crew_tab, from_=0, to=20, textvariable=self.vars['soldiers'], width=5)
        soldier_spinbox.grid(row=3, column=1, sticky="w")
        ToolTip(soldier_spinbox, "Number of infantry soldiers the tank can carry.")
        
        overworked_frame = ttk.Frame(crew_tab)
        overworked_frame.grid(row=4, column=0, columnspan=2, sticky='w', pady=5)
        cb8 = ttk.Checkbutton(overworked_frame, text="Overworked Crew", variable=self.vars['is_overworked'])
        cb8.pack(side=tk.LEFT)
        ToolTip(cb8, "Are the crew pushed past their limits? (Applies a flat penalty to Morale)")

        result_frame = ttk.LabelFrame(parent, text="Final Calculation", padding=10)
        result_frame.grid(row=2, column=0, sticky="ew", pady=10)
        result_frame.columnconfigure(0, weight=1)
        ttk.Button(result_frame, text="Calculate Battle Rating", command=self.perform_calculations).pack()
        scores_frame = ttk.Frame(result_frame, padding=(0, 5)); scores_frame.pack(fill=tk.X); scores_frame.columnconfigure(1, weight=1)
        ttk.Label(scores_frame, text="Battle Rating:", font=("", 14, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(scores_frame, textvariable=self.vars['battle_rating'], font=("Courier", 16, "bold"), foreground="orange").grid(row=0, column=1, sticky="w")
        details_frame = ttk.Frame(result_frame); details_frame.pack(fill=tk.X, pady=5)
        ttk.Label(details_frame, text="Combat Strength:").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Label(details_frame, textvariable=self.vars['combat_strength']).grid(row=0, column=1, sticky="w")
        ttk.Label(details_frame, text="Morale Score:").grid(row=1, column=0, sticky="w", padx=5)
        ttk.Label(details_frame, textvariable=self.vars['morale_score']).grid(row=1, column=1, sticky="w")
        ttk.Label(result_frame, text="Active Multipliers:", anchor="w").pack(fill=tk.X, pady=(5,0))
        ttk.Label(result_frame, textvariable=self.vars['interaction_details'], wraplength=450, justify=tk.LEFT).pack(anchor="w", fill=tk.X, padx=5)

    def create_saved_designs_widgets(self, parent):
        parent.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        parent.rowconfigure(0, weight=1); parent.columnconfigure(0, weight=1)
        
        self.designs_listbox = tk.Listbox(parent, selectmode=tk.SINGLE, exportselection=False)
        self.designs_listbox.grid(row=0, column=0, sticky="nsew")
        self.designs_listbox.bind("<<ListboxSelect>>", self.on_design_select)

        buttons = ttk.Frame(parent); buttons.grid(row=1, column=0, pady=10)
        ttk.Button(buttons, text="Save Current", command=self.save_tank_design).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons, text="Delete", command=self.delete_tank_design).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Clear", command=self.clear_designer_form).pack(side=tk.LEFT, padx=5)

    def on_secondary_gun_change(self, *args):
        state = 'normal' if self.vars['has_secondary_gun'].get() else 'disabled'
        if 'secondary_weapon_widgets' in self.ui_elements:
            for widget in self.ui_elements['secondary_weapon_widgets']:
                widget.configure(state=state)

    def on_stationary_change(self, *args):
        is_stationary = self.vars['is_stationary'].get()
        state = 'disabled' if is_stationary else 'normal'
        if 'mobility_widgets' in self.ui_elements:
            for widget in self.ui_elements['mobility_widgets']:
                widget.configure(state=state)
            if is_stationary:
                self.vars['mobility'].set(0)

    def on_open_top_change(self, *args):
        state = 'disabled' if self.vars['is_open_top'].get() else 'normal'
        if 'upper_armor_widgets' in self.ui_elements:
            for widget in self.ui_elements['upper_armor_widgets']:
                widget.configure(state=state)

    def on_gun_type_change(self, *args):
        is_turreted = self.vars['gun_type'].get() == 'Turreted'
        for child in self.winfo_children():
            self._find_and_set_state(child, 'turret_mobility', 'normal' if is_turreted else 'disabled')
            self._find_and_set_state(child, 'turret_profile', 'normal' if is_turreted else 'disabled')
            self._find_and_set_state(child, 'gun_traverse_range', 'disabled' if is_turreted else 'normal')

    def _find_and_set_state(self, parent, var_name, state):
        for widget in parent.winfo_children():
            if isinstance(widget, ttk.Scale) and widget.cget('variable') == str(self.vars[var_name]):
                widget.config(state=state)
                return
            self._find_and_set_state(widget, var_name, state)

    def perform_calculations(self):
        faction_name = self.vars['faction'].get()
        if not faction_name or faction_name not in self.factions:
            messagebox.showerror("Calc Error", "Please select a valid faction."); return
        
        cs = self.calculate_combat_strength()
        ms = self.calculate_morale_score(faction_name)
        if cs is None or ms is None:
            self.vars['battle_rating'].set("Error")
            return

        tank_stats = {k: v.get() for k, v in self.vars.items()}
        faction_stats = self.factions[faction_name]
        im, details = self.calculate_interaction_multipliers(tank_stats, faction_stats, ms)
        morale_mod = 0.8 + (ms / 100) * 0.4
        
        final_br = cs * morale_mod * im

        if tank_stats['is_rushed']:
            final_br *= 0.6
            details += "\n• Rushed Production (x0.6)"

        self.vars['combat_strength'].set(f"{cs:.2f}")
        self.vars['morale_score'].set(f"{ms:.2f}")
        self.vars['interaction_details'].set(details.strip() if details else "No special interactions.")
        self.vars['battle_rating'].set(f"{final_br:.2f}")

    def calculate_combat_strength(self):
        try:
            s = {k: v.get() for k, v in self.vars.items()}
            
            pen_score = s['gun_pen']
            cal_score = (s['ammo_caliber'] - 10) / (250 - 10) * 100
            reload_score = 100 - s['gun_reload_speed']
            firepower = (pen_score * 0.4) + (cal_score * 0.4) + (reload_score * 0.2)
            
            if s['has_secondary_gun']:
                main_gun_stats = [s['gun_pen'], cal_score, reload_score, s['gun_mobility']]
                if s['gun_type'] == 'Turreted': main_gun_stats.append(s['turret_mobility'])
                else: main_gun_stats.append(s['gun_traverse_range'])
                
                avg_gun_stat = sum(main_gun_stats) / len(main_gun_stats)
                reduced_avg = avg_gun_stat * 0.85
                secondary_bonus = reduced_avg * (s['secondary_weapon'] / 100.0)
                firepower += secondary_bonus

            fwd_armor = (s['upper_forward_armor'] + s['lower_forward_armor']) / 2.0
            side_armor = (s['upper_side_armor'] + s['lower_side_armor']) / 2.0
            rear_armor = (s['upper_rear_armor'] + s['lower_rear_armor']) / 2.0

            upper_armor = 0 if s['is_open_top'] else s['upper_armor']
            weighted_armor_sum = (fwd_armor * 2) + side_armor + rear_armor + upper_armor
            
            engine_armor_bonus = (100 - s['engine_position']) / 20.0
            weighted_armor_sum += engine_armor_bonus

            weighted_armor_avg = weighted_armor_sum / 5.0
            
            profile_area = s['tank_length'] * s['tank_height']
            profile_penalty = 1.0 - ((profile_area - 15000000) / 50000000.0)
            profile_penalty = max(0.7, min(1.3, profile_penalty))

            turret_profile_mod = 1 - (s['turret_profile'] / 200) if s['gun_type'] == 'Turreted' else 1
            survivability = weighted_armor_avg * turret_profile_mod * profile_penalty

            if s['has_spaced_armor'] and s['tech_level'] >= 4:
                survivability *= 1.15

            mobility_score = 0 if s['is_stationary'] else s['mobility'] * (1.15 if s['has_dual_transmission'] else 1.0)
            
            if s['tonnage'] > 100 and mobility_score < 30 and not s['is_stationary']:
                mobility_score *= 0.9

            gun_handling = (s['gun_mobility'] + (s['turret_mobility'] if s['gun_type']=='Turreted' else s['gun_traverse_range'])) / 2
            
            if s['tonnage'] < 40 and s['ammo_caliber'] > 120:
                gun_handling *= 0.85

            if s['gun_barrel_length'] > 8000 and s['gun_mobility'] < 30:
                gun_handling *= 0.90
            
            depression_score = min(s['gun_depression'] / 12.0, 1.2) * 100
            elevation_score = min(s['gun_elevation'] / 45.0, 1.0) * 100
            gun_arc_score = (depression_score * 0.7) + (elevation_score * 0.3)

            tactical = (mobility_score * 0.3) + (s['awareness'] * 0.3) + (gun_handling * 0.2) + (gun_arc_score * 0.2)
            
            base_cs = (firepower * 0.4) + (survivability * 0.35) + (tactical * 0.25)
            
            if s['is_open_top']:
                base_cs *= 1.05

            tech_mult = 1.0 + (s['tech_level'] * 0.10)
            return base_cs * tech_mult
        except: return None

    def calculate_morale_score(self, faction_name):
        try:
            fs = self.factions[faction_name]
            ts = {k: v.get() for k, v in self.vars.items()}
            faction_comp = (fs['leadership'] + fs['stability'] + fs['wealth']) / 3.0
            tank_comp = (ts['design_rating'] + ts['crew_comfort']) / 2.0
            crew_bonus = ts['extra_crew'] * 1.5
            
            morale = (faction_comp * 0.6) + (tank_comp * 0.4) + crew_bonus
            
            if ts['is_overworked']:
                morale -= 10
            
            if ts['is_open_top']:
                morale *= 1.10
                
            return min(100, max(0, morale))
        except: return None

    def calculate_interaction_multipliers(self, ts, fs, morale):
        m, d_list = 1.0, []
        HIGH, LOW = 75, 25
        
        fwd_armor_avg = (ts['upper_forward_armor'] + ts['lower_forward_armor']) / 2.0
        side_armor_avg = (ts['upper_side_armor'] + ts['lower_side_armor']) / 2.0

        # Positive Modifiers
        if fs['leadership'] > HIGH and ts['awareness'] > HIGH: m*=1.05; d_list.append("• Elite Crew (+5%)")
        if ts['mobility'] > HIGH and ts['has_dual_transmission']: m*=1.07; d_list.append("• Expert Flanker (+7%)")
        if fs['stability'] > HIGH and (ts['design_rating']<LOW or ts['crew_comfort']<LOW): m*=1.04; d_list.append("• Stoic Crew (+4%)")
        if ts['turret_profile'] < LOW and ts['awareness'] > HIGH and ts['gun_type']=='Turreted': m*=1.07; d_list.append("• Ambusher (+7%)")
        if ts['gun_depression'] > 10 and ts['turret_profile'] < LOW and ts['gun_type'] == 'Turreted': m*=1.08; d_list.append("• Hull-Down Expert (+8%)")
        if ts['gun_type'] == 'Fixed' and ts['has_dual_transmission']: m*=1.04; d_list.append("• Shoot & Scoot (+4%)")
        if fwd_armor_avg > 85 and ts['gun_pen'] > 85 and ts['mobility'] < LOW: m*=1.05; d_list.append("• Breakthrough Tank (+5%)")
        
        # Mutually Exclusive Artillery Modifiers
        artillery_doctrine_active = False
        if ts['gun_elevation'] > 30 and ts['ammo_caliber'] > 120:
            m*=1.05; d_list.append("• Artillery Doctrine (+5%)")
            artillery_doctrine_active = True
        if not artillery_doctrine_active and ts['gun_elevation'] >= 20 and ts['ammo_length'] < 300:
            m*=1.03; d_list.append("• Light Indirect Fire (+3%)")

        # Mutually Exclusive Loader Modifiers
        if ts['gun_reload_speed'] < LOW:
            if ts['extra_crew'] >= 3:
                m *= 1.06; d_list.append("• Coordinated Team (+6%)")
            elif ts['extra_crew'] == 2:
                m *= 1.03; d_list.append("• Loader Team (+3%)")

        # Negative Modifiers
        if morale < LOW and ts['mobility'] < LOW: m*=0.90; d_list.append("• Crew Panic (-10%)")
        if fs['wealth'] < LOW and ts['tech_level'] > 3: m*=0.95; d_list.append("• Poor Maint. (-5%)")
        if ts['gun_pen'] < LOW and ts['mobility'] < LOW and fwd_armor_avg < LOW: m*=0.93; d_list.append("• Obsolete (-7%)")
        if ts['gun_depression'] < 3 and ts['gun_type'] == 'Turreted': m*=0.95; d_list.append("• Limited Depression (-5%)")
        if ts['is_open_top'] and ts['awareness'] < 25: m*=0.90; d_list.append("• Vulnerable Open Top (-10%)")

        # Cramped Conditions Modifier
        cramped_cond1 = ts['crew_comfort'] < 30 and (ts['extra_crew'] + ts['soldiers']) > 5
        cramped_cond2 = ts['crew_comfort'] < 25 and ts['extra_crew'] == 0 and ts['soldiers'] == 0
        if cramped_cond1 or cramped_cond2:
            m*=0.96; d_list.append("• Dangerously Cramped (-4%)")

        # New Negative Modifiers
        if ts['tech_level'] == 0 and fwd_armor_avg > 70: m *= 0.94; d_list.append("• Brittle Armor (-6%)")
        if ts['mobility'] > 80 and ts['gun_mobility'] < 30: m *= 0.95; d_list.append("• Unstable Firing Platform (-5%)")
        if ts['gun_type'] == 'Turreted' and ts['tonnage'] > 150 and ts['turret_mobility'] < 25: m *= 0.96; d_list.append("• Underpowered Traverse (-4%)")
        if side_armor_avg < 20 and ts['ammo_caliber'] > 120: m *= 0.92; d_list.append("• Exposed Ammo Rack (-8%)")
        if ts['engine_position'] < 10 and ts['ammo_caliber'] > 150: m *= 0.97; d_list.append("• Poor Weight Distribution (-3%)")
        
        return m, "\n".join(d_list)
    
    def open_battle_setup(self):
        if not self.factions or not self.tank_designs:
            messagebox.showerror("Error", "Create at least one faction and one tank design before starting a battle.")
            return
        BattleSetupWindow(self, "Battle Setup", self.factions, self.tank_designs)

    def create_slider_entry(self, p, l, v, r, from_=0, to=100, tooltip=None):
        lbl = ttk.Label(p, text=l)
        lbl.grid(row=r, column=0, sticky="w", pady=2, padx=2)
        if tooltip:
            ToolTip(lbl, tooltip)

        s = ttk.Scale(p, from_=from_, to=to, orient=tk.HORIZONTAL, variable=v)
        s.grid(row=r, column=1, sticky="ew", padx=5)
        
        e = ttk.Entry(p, textvariable=v, width=6)
        e.grid(row=r, column=2, padx=(0, 2))

        def clamp_on_focus_out(event):
            try:
                val = float(v.get())
                if val < from_: v.set(from_)
                elif val > to: v.set(to)
            except (ValueError, tk.TclError):
                v.set(from_)

        e.bind("<FocusOut>", clamp_on_focus_out)
        e.bind("<Return>", lambda e: p.focus())

        return [lbl, s, e]

    def add_update_faction(self):
        n = self.faction_vars['name'].get().strip()
        if not n: messagebox.showerror("Input Error", "Faction name cannot be empty."); return
        self.factions[n] = {k: v.get() for k, v in self.faction_vars.items() if k != 'name'}
        self.populate_faction_list(); self.update_faction_combobox()
    def delete_faction(self):
        s = self.faction_listbox.curselection()
        if not s: messagebox.showwarning("Selection Error", "Please select a faction."); return
        n = self.faction_listbox.get(s[0])
        if messagebox.askyesno("Confirm Delete", f"Delete faction '{n}'?"):
            if n in self.factions:
                del self.factions[n]; self.populate_faction_list(); self.update_faction_combobox()
    def on_faction_select(self, e):
        s = self.faction_listbox.curselection()
        if not s: return
        n = self.faction_listbox.get(s[0]);
        if n not in self.factions: return
        stats = self.factions[n]
        self.faction_vars['name'].set(n)
        for k, v in stats.items(): self.faction_vars[k].set(v)
    def populate_faction_list(self):
        self.faction_listbox.delete(0, tk.END)
        for n in sorted(self.factions.keys()):
            self.faction_listbox.insert(tk.END, n)
    def update_faction_combobox(self):
        n = sorted(list(self.factions.keys())); self.faction_combo['values'] = n
        if n and self.vars['faction'].get() not in n: self.vars['faction'].set(n[0])
        elif not n: self.vars['faction'].set("")
    def save_tank_design(self):
        n = self.vars['tank_name'].get().strip()
        if not n: messagebox.showerror("Input Error", "Tank name cannot be empty."); return
        self.perform_calculations()
        self.tank_designs[n] = {k: v.get() for k, v in self.vars.items()}
        self.populate_tank_list()
    def delete_tank_design(self):
        s = self.designs_listbox.curselection()
        if not s: messagebox.showwarning("Selection Error", "Select a design."); return
        n = self.designs_listbox.get(s[0])
        if messagebox.askyesno("Confirm Delete", f"Delete design '{n}'?"):
            del self.tank_designs[n]; self.populate_tank_list()
    def on_design_select(self, e):
        s = self.designs_listbox.curselection()
        if not s: return
        n = self.designs_listbox.get(s[0]);
        if n not in self.tank_designs: return
        data = self.tank_designs[n]
        for k, v in data.items():
            if k in self.vars: self.vars[k].set(v)
    def clear_designer_form(self):
        self.vars['tank_name'].set(""); self.designs_listbox.selection_clear(0, tk.END)
    def populate_tank_list(self):
        self.designs_listbox.delete(0, tk.END)
        for n in sorted(self.tank_designs.keys()):
            self.designs_listbox.insert(tk.END, n)
    def save_data_to_file(self):
        try:
            with open(self.DATA_FILE, 'w') as f: json.dump({"factions": self.factions, "tank_designs": self.tank_designs}, f, indent=4)
            messagebox.showinfo("Save Successful", f"Data saved to {self.DATA_FILE}")
        except Exception as e: messagebox.showerror("Save Error", f"Could not save file:\n{e}")
    def load_data_from_file(self):
        if not os.path.exists(self.DATA_FILE): return
        try:
            with open(self.DATA_FILE, 'r') as f: data = json.load(f)
            self.factions = data.get("factions", {}); self.tank_designs = data.get("tank_designs", {})
            self.populate_faction_list(); self.update_faction_combobox(); self.populate_tank_list()
        except Exception as e: messagebox.showerror("Load Error", f"Could not load file:\n{e}")

class BattleSetupWindow(tk.Toplevel):
    def __init__(self, parent, title, factions, designs):
        super().__init__(parent)
        self.transient(parent); self.grab_set(); self.title(title)
        self.parent = parent
        self.factions = factions; self.designs = designs
        self.attacker_roster = []; self.defender_roster = []

        self.env_var = tk.StringVar(value="Grassy Fields")
        self.attacker_faction_var = tk.StringVar()
        self.defender_faction_var = tk.StringVar()
        self.flank_attack_var = tk.BooleanVar(value=False)
        self.both_attacking_var = tk.BooleanVar(value=False)
        self.def_entrenched_var = tk.BooleanVar(value=False)
        
        main_frame = ttk.Frame(self, padding=15); main_frame.pack(expand=True, fill=tk.BOTH)

        # Top frame for environment and battle conditions
        top_frame = ttk.LabelFrame(main_frame, text="Battle Conditions", padding=10);
        top_frame.pack(fill=tk.X, pady=(0, 15))
        
        env_frame = ttk.Frame(top_frame)
        env_frame.pack(fill=tk.X)
        ttk.Label(env_frame, text="Environment:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Combobox(env_frame, textvariable=self.env_var, state="readonly",
                     values=["Grassy Fields", "Sand Dunes", "Urban Environment", "Dense Forest", "Mountains", "Coastline", "Fortified Position"]).pack(side=tk.LEFT)

        options_frame = ttk.Frame(top_frame)
        options_frame.pack(fill=tk.X, pady=(5,0))

        cb_flank = ttk.Checkbutton(options_frame, text="Flank Attack (>1 Direction)", variable=self.flank_attack_var)
        cb_flank.pack(side=tk.LEFT, padx=(0, 10))

        cb_both_attacking = ttk.Checkbutton(options_frame, text="Both Teams Attacking", variable=self.both_attacking_var)
        cb_both_attacking.pack(side=tk.LEFT, padx=10)
        
        cb_entrenched = ttk.Checkbutton(options_frame, text="Defenders Entrenched", variable=self.def_entrenched_var)
        cb_entrenched.pack(side=tk.LEFT, padx=10)

        def _update_checkboxes(*args):
            if self.both_attacking_var.get():
                self.def_entrenched_var.set(False)
                cb_entrenched.config(state='disabled')
            else:
                cb_entrenched.config(state='normal')

            if self.def_entrenched_var.get():
                self.both_attacking_var.set(False)
                cb_both_attacking.config(state='disabled')
            else:
                cb_both_attacking.config(state='normal')
        
        self.both_attacking_var.trace_add("write", _update_checkboxes)
        self.def_entrenched_var.trace_add("write", _update_checkboxes)

        # Frame for Attacker and Defender rosters
        battle_frame = ttk.Frame(main_frame); battle_frame.pack(expand=True, fill=tk.BOTH)
        self.attacker_frame = self.create_team_frame(battle_frame, "Attacker", self.attacker_faction_var, self.attacker_roster)
        self.attacker_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0, 10))
        self.defender_frame = self.create_team_frame(battle_frame, "Defender", self.defender_faction_var, self.defender_roster)
        self.defender_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        bottom_frame = ttk.Frame(main_frame, padding=(0, 10)); bottom_frame.pack(fill=tk.X)
        ttk.Button(bottom_frame, text="Simulate Battle", command=self.run_simulation).pack(side=tk.RIGHT)
        ttk.Button(bottom_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=10)

    def create_team_frame(self, parent, name, faction_var, roster):
        frame = ttk.LabelFrame(parent, text=name, padding=10)
        
        controls = ttk.Frame(frame); controls.pack(fill=tk.X, pady=(0,10))
        ttk.Label(controls, text="Faction:").grid(row=0, column=0, padx=2)
        ttk.Combobox(controls, textvariable=faction_var, state="readonly", values=sorted(list(self.factions.keys()))).grid(row=0, column=1)
        
        add_frame = ttk.Frame(frame); add_frame.pack(fill=tk.X, pady=5)
        design_var = tk.StringVar()
        qty_var = tk.IntVar(value=1)
        design_combo = ttk.Combobox(add_frame, textvariable=design_var, state="readonly", values=sorted(list(self.designs.keys())))
        design_combo.pack(side=tk.LEFT, expand=True, fill=tk.X)
        ttk.Spinbox(add_frame, from_=1, to=100, textvariable=qty_var, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Button(add_frame, text="Add", command=lambda d=design_var, q=qty_var, r=roster, f=frame: self.add_to_roster(d, q, r, f)).pack(side=tk.LEFT)
        
        tree = ttk.Treeview(frame, columns=("qty", "br"), show="headings", selectmode="browse")
        tree.pack(expand=True, fill=tk.BOTH)
        tree.heading("#0", text="Design"); tree.column("#0", width=120)
        tree.heading("qty", text="Qty"); tree.column("qty", width=40, anchor="center")
        tree.heading("br", text="Base BR"); tree.column("br", width=60, anchor="center")
        
        frame.roster_tree = tree
        return frame

    def add_to_roster(self, design_var, qty_var, roster, frame):
        design_name = design_var.get()
        qty = qty_var.get()
        if not design_name or qty < 1: return
        
        roster.append({"name": design_name, "qty": qty})
        tree = frame.roster_tree
        tree.insert("", "end", text=design_name, values=(qty, self.designs[design_name].get('battle_rating', '--')))
        
    def run_simulation(self):
        if not self.attacker_roster or not self.defender_roster:
            messagebox.showerror("Error", "Both attacker and defender must have units."); return
        if not self.attacker_faction_var.get() or not self.defender_faction_var.get():
            messagebox.showerror("Error", "Both teams must have a faction selected."); return

        sim_engine = BattleSimulator(self.parent, self.factions, self.designs)
        results = sim_engine.simulate(
            self.attacker_roster, self.defender_roster,
            self.attacker_faction_var.get(), self.defender_faction_var.get(),
            self.env_var.get(),
            self.flank_attack_var.get(),
            self.both_attacking_var.get(),
            self.def_entrenched_var.get()
        )
        self.destroy()
        BattleResultsWindow(self.parent, "Battle Results", results)

class BattleSimulator:
    def __init__(self, parent, factions, designs):
        self.parent = parent
        self.factions = factions
        self.designs = designs
        self.log = []
        # Environment modifiers dictionary
        self.env_modifiers = {
            "Grassy Fields": {
                'attacker_power_mod': 1.0,
                'defender_defense_mod': 1.0,
                'ambient_danger_chance': 0.05
            },
            "Fortified Position": {
                'attacker_power_mod': 0.7,
                'defender_defense_mod': 1.6,
                'ambient_danger_chance': 0.35
            },
            "Mountains": {
                'attacker_power_mod': 0.75,
                'defender_defense_mod': 1.5,
                'ambient_danger_chance': 0.30
            },
            "Urban Environment": {
                'attacker_power_mod': 0.85,
                'defender_defense_mod': 1.4,
                'ambient_danger_chance': 0.25
            },
            "Dense Forest": {
                'attacker_power_mod': 0.8,
                'defender_defense_mod': 1.3,
                'ambient_danger_chance': 0.20
            },
            "Sand Dunes": {
                'attacker_power_mod': 0.9,
                'defender_defense_mod': 1.15,
                'ambient_danger_chance': 0.10
            },
            "Coastline": {
                'attacker_power_mod': 0.95,
                'defender_defense_mod': 1.0,
                'ambient_danger_chance': 0.08
            },
        }

    def _get_weighted_average_stat(self, roster, stat_key):
        total_stat, total_units = 0, sum(u['hp'] for u in roster)
        if total_units == 0: return 0
        for unit in roster:
            design = self.designs[unit['name']]
            total_stat += float(design.get(stat_key, 50)) * unit['hp']
        return total_stat / total_units

    def _apply_casualties(self, roster, num_casualties, log_message_prefix=""):
        if not roster or num_casualties <= 0: return roster, 0
        casualties_inflicted = 0
        roster.sort(key=lambda u: self.designs[u['name']].get('upper_forward_armor', 50))
        for _ in range(num_casualties):
            if not roster: break
            target_unit = roster[0]
            target_unit['hp'] -= 1
            casualties_inflicted += 1
            self.log.append(f"  {log_message_prefix} A {target_unit['name']} is destroyed!")
            if target_unit['hp'] <= 0: roster.pop(0)
        return roster, casualties_inflicted

    def _apply_surrender(self, roster, percentage):
        if not roster: return roster, []
        num_to_capture = math.ceil(sum(u['hp'] for u in roster) * percentage)
        if num_to_capture == 0 and sum(u['hp'] for u in roster) > 0:
             num_to_capture = 1 # Ensure at least one unit is captured if any remain
        
        captured_roster = []
        roster.sort(key=lambda u: self.designs[u['name']].get('upper_forward_armor', 50))
        
        captured_this_loop = 0
        while captured_this_loop < num_to_capture and roster:
            target_unit = roster[0]
            target_unit['hp'] -= 1
            captured_this_loop += 1
            
            found = False
            for captured_unit in captured_roster:
                if captured_unit['name'] == target_unit['name']:
                    captured_unit['qty'] += 1
                    found = True
                    break
            if not found:
                captured_roster.append({'name': target_unit['name'], 'qty': 1})
            
            self.log.append(f"  A {target_unit['name']} surrenders and is captured!")
            if target_unit['hp'] <= 0:
                roster.pop(0)
                
        return roster, captured_roster

    def _merge_capture_lists(self, main_list, new_items):
        for new_item in new_items:
            found = False
            for main_item in main_list:
                if main_item['name'] == new_item['name']:
                    main_item['qty'] += new_item['qty']
                    found = True
                    break
            if not found:
                main_list.append(new_item)
        return main_list

    def _resolve_combat_phase(self, attackers, defenders, phase, is_defender_flanked, environment, both_attacking, apply_entrenchment_bonus):
        num_attackers = sum(u['hp'] for u in attackers)
        if num_attackers == 0: return 0

        # --- 1. Get Environment Modifiers ---
        mods = self.env_modifiers.get(environment, self.env_modifiers["Grassy Fields"])
        attacker_mod = mods['attacker_power_mod']
        defender_mod = mods['defender_defense_mod']

        if both_attacking:
            defender_mod = 1.0 # In a meeting engagement, defensive terrain bonuses are nullified

        if environment == "Dense Forest" and phase == 'Opening Volley':
            attacker_mod *= 0.8

        attack_power = 0
        for unit in attackers:
            design = self.designs[unit['name']]
            tech_mod = 1 + (design.get('tech_level', 0) * 0.1)
            power_components = {}
            if phase == 'Opening Volley':
                power_components = {
                    'Penetration': design.get('gun_pen', 50) * 1.5,
                    'Barrel Length': math.sqrt(design.get('gun_barrel_length', 4000)),
                    'Awareness': design.get('awareness', 50)
                }
            elif phase == 'Early Engagement':
                 power_components = {
                    'Aim Speed': design.get('gun_mobility', 50) * 1.5,
                    'Penetration': design.get('gun_pen', 50) * 1.0,
                    'Awareness': design.get('awareness', 50) * 0.5
                }
            else: # Late Engagement
                power_components = {
                    'Reload Speed': (100 - design.get('gun_reload_speed', 50)) * 1.5,
                    'Ammo Caliber': design.get('ammo_caliber', 105) * 1.0,
                    'Turret Mobility': design.get('turret_mobility', 50) * 0.5
                }
            attack_power += sum(power_components.values()) * tech_mod * unit['hp']
        
        # --- 2. Apply Attacker Modifier ---
        attack_power *= attacker_mod

        entrenchment_bonus = 50.0 if apply_entrenchment_bonus else 0
        defense_value = 0
        for unit in defenders:
            design = self.designs[unit['name']]
            tech_mod = 1 + (design.get('tech_level', 0) * 0.1)
            value = 0
            if phase == 'Opening Volley':
                profile_area = design.get('tank_length', 6000) * design.get('tank_height', 2500)
                profile_mod = 15000000 / profile_area if profile_area > 0 else 1.0
                armor = (design.get('upper_forward_armor', 50) + design.get('lower_forward_armor', 50)) / 2
                value = (armor * 2.0 + design.get('mobility', 50) * profile_mod) * tech_mod
            elif phase == 'Early Engagement':
                armor = (design.get('upper_forward_armor', 50) + design.get('lower_forward_armor', 50)) / 2
                value = (armor * 2.0 + design.get('mobility', 50)) * tech_mod
            else: # Late Engagement
                armor_key = 'upper_side_armor' if is_defender_flanked else 'upper_forward_armor'
                armor = (design.get(armor_key, 50) + design.get(armor_key.replace('upper', 'lower'), 50)) / 2
                value = (armor * 1.5 + design.get('tonnage', 50) * 1.5) * tech_mod
            defense_value += (value + entrenchment_bonus) * unit['hp']
        
        # --- 3. Apply Defender Modifier ---
        defense_value *= defender_mod

        if attack_power <= 0: return 0
        hit_ratio = attack_power / defense_value if defense_value > 0 else 10.0
        kill_probability = 0.10
        effective_ratio = min(hit_ratio, 3.0)
        expected_kills = num_attackers * kill_probability * effective_ratio
        
        final_casualties = math.floor(expected_kills) + (1 if random.random() < (expected_kills - math.floor(expected_kills)) else 0)
        
        # --- 4. NEW "Ambient Danger" Mechanic ---
        if final_casualties == 0 and num_attackers > 0:
            ambient_chance = mods.get('ambient_danger_chance', 0.05)
            avg_attacker_awareness = self._get_weighted_average_stat(attackers, 'awareness')
            avg_defender_awareness = self._get_weighted_average_stat(defenders, 'awareness')
            if avg_defender_awareness > avg_attacker_awareness:
                ambient_chance *= 1.5
            if random.random() < ambient_chance:
                self.log.append(f"   A lucky shot gets through the chaos!")
                final_casualties = 1

        log_phase_str = f"({phase} | ATT Mod: {attacker_mod:.2f}, DEF Mod: {defender_mod:.2f}"
        if apply_entrenchment_bonus: log_phase_str += ", Entrenched"
        log_phase_str += ")"
        self.log.append(f"-> Attacker Power: {int(attack_power)} vs Defender Value: {int(defense_value)} {log_phase_str}")
        if final_casualties > 0:
            self.log.append(f"   The exchange is effective, inflicting {final_casualties} casualties!")
        else:
            self.log.append(f"   The defenders' armor holds! The attack is ineffective.")
        return final_casualties


    def _compile_final_roster(self, original, final):
        final_map = {u['name']: u['hp'] for u in final}
        return [{'name': u['name'], 'qty': u['qty'], 'survivors': final_map.get(u['name'], 0)} for u in original]

    def simulate(self, attacker_roster_orig, defender_roster_orig, attacker_faction_name, defender_faction_name, environment, is_flank_attack, both_attacking, defenders_entrenched):
        self.log = [f"Battle Begins! {attacker_faction_name} vs. {defender_faction_name} in {environment}."]
        attacker_roster = [dict(u, hp=u['qty']) for u in attacker_roster_orig]
        defender_roster = [dict(u, hp=u['qty']) for u in defender_roster_orig]
        defender_is_flanked = is_flank_attack
        surrendered_by_defender = []
        if is_flank_attack: self.log.append("The attackers begin with a flanking advantage!")
        if both_attacking: self.log.append("This is a meeting engagement; both sides are on the offensive!")
        if defenders_entrenched: self.log.append("The defenders are entrenched and will be harder to dislodge.")

        for turn in range(1, 11):
            if not attacker_roster or not defender_roster: break
            self.log.append(f"\n--- Turn {turn} ---")
            att_cas_turn, def_cas_turn = 0, 0

            # --- Opening Volley ---
            att_aware = self._get_weighted_average_stat(attacker_roster, 'awareness')
            def_aware = self._get_weighted_average_stat(defender_roster, 'awareness')
            self.log.append(f"Opening Volley Phase (Awareness: ATT {att_aware:.0f} vs DEF {def_aware:.0f})")
            if att_aware > def_aware:
                self.log.append("Attackers' superior awareness allows them to fire the first volley.")
                cas = self._resolve_combat_phase(attacker_roster, defender_roster, 'Opening Volley', defender_is_flanked, environment, both_attacking, defenders_entrenched)
                defender_roster, lost = self._apply_casualties(defender_roster, cas, "Attacker Volley:")
                def_cas_turn += lost
            else:
                self.log.append("Defenders' vigilance allows them to fire the first volley.")
                cas = self._resolve_combat_phase(defender_roster, attacker_roster, 'Opening Volley', False, environment, both_attacking, False)
                attacker_roster, lost = self._apply_casualties(attacker_roster, cas, "Defender Volley:")
                att_cas_turn += lost
            if not attacker_roster or not defender_roster: break

            # --- Maneuver Phase ---
            att_mob = self._get_weighted_average_stat(attacker_roster, 'mobility')
            def_mob = self._get_weighted_average_stat(defender_roster, 'mobility')
            self.log.append(f"Maneuver Phase (Mobility: ATT {att_mob:.0f} vs DEF {def_mob:.0f})")
            if att_mob > def_mob * 1.2 and not defender_is_flanked:
                self.log.append("Attackers outmaneuver the defenders, achieving a flanking position!")
                defender_is_flanked = True
            elif def_mob > att_mob * 1.2 and defender_is_flanked:
                self.log.append("Defenders use their mobility to counter-maneuver and secure their flanks.")
                defender_is_flanked = False

            # --- Early & Late Engagement Phases ---
            for phase in ["Early Engagement", "Late Engagement"]:
                if not attacker_roster or not defender_roster: break
                self.log.append(f"--- {phase} Phase ---")
                cas = self._resolve_combat_phase(attacker_roster, defender_roster, phase, defender_is_flanked, environment, both_attacking, defenders_entrenched)
                defender_roster, lost = self._apply_casualties(defender_roster, cas, "Attacker Fire:")
                def_cas_turn += lost
                if defender_roster:
                    cas = self._resolve_combat_phase(defender_roster, attacker_roster, phase, False, environment, both_attacking, False)
                    attacker_roster, lost = self._apply_casualties(attacker_roster, cas, "Defender Fire:")
                    att_cas_turn += lost

            # --- Morale Phase ---
            att_total, def_total = sum(u['qty'] for u in attacker_roster_orig), sum(u['qty'] for u in defender_roster_orig)
            att_morale_hit = (att_cas_turn / att_total) * 150 if att_total > 0 else 0
            def_morale_hit = (def_cas_turn / def_total) * 150 if def_total > 0 else 0
            att_morale = self._get_weighted_average_stat(attacker_roster, 'morale_score') - att_morale_hit
            def_morale = self._get_weighted_average_stat(defender_roster, 'morale_score') - def_morale_hit
            self.log.append(f"Morale Check: ATT ({att_morale:.0f}) vs DEF ({def_morale:.0f})")
            if att_morale < 0:
                self.log.append("Attackers' morale shatters! They are in retreat!"); break
            if def_morale < 0:
                self.log.append("Defenders' line breaks! They surrender and 20% of their force is captured!")
                defender_roster, new_surrenders = self._apply_surrender(defender_roster, 0.20)
                surrendered_by_defender = self._merge_capture_lists(surrendered_by_defender, new_surrenders)
                break
        else: self.log.append("\nThe battle becomes a stalemate after 10 turns.")

        self.log.append("\n--- BATTLE CONCLUSION ---")
        att_survivors, def_survivors = sum(u['hp'] for u in attacker_roster), sum(u['hp'] for u in defender_roster)
        winner = "Defender" if att_survivors == 0 or (att_survivors > 0 and def_survivors > 0) else "Attacker"
        self.log.append(f"Decisive victory for the {winner}!" if (att_survivors == 0 or def_survivors == 0) else "The defenders hold the field.")
        
        final_attacker = self._compile_final_roster(attacker_roster_orig, attacker_roster)
        final_defender = self._compile_final_roster(defender_roster_orig, defender_roster)
        capt_att, capt_def = [], []
        if winner == "Attacker":
            capt_att = self._calculate_captures(final_defender, final_attacker, self.factions[attacker_faction_name])
            capt_att = self._merge_capture_lists(capt_att, surrendered_by_defender)
        else:
            capt_def = self._calculate_captures(final_attacker, final_defender, self.factions[defender_faction_name])
        return {"log": self.log, "attacker_roster_final": final_attacker, "defender_roster_final": final_defender, "captured_by_attacker": capt_att, "captured_by_defender": capt_def}
    
    def _calculate_captures(self, loser_roster, victor_roster, victor_faction):
        capture_list = []
        total_soldiers = sum(self.designs[u['name']].get('soldiers', 0) * u['survivors'] for u in victor_roster)
        soldier_capture_bonus = total_soldiers * 0.001
        for unit in loser_roster:
            losses = unit['qty'] - unit['survivors']
            if losses == 0: continue
            design = self.designs[unit['name']]
            base_chance = (victor_faction['wealth'] / 1500.0) - (design['tech_level'] * 0.01)
            final_chance = max(0, base_chance + soldier_capture_bonus)
            captured_count = sum(1 for _ in range(losses) if random.random() < final_chance)
            if captured_count > 0: capture_list.append({"name": unit['name'], "qty": captured_count})
        return capture_list
        
class BattleResultsWindow(tk.Toplevel):
    def __init__(self, parent, title, results):
        super().__init__(parent)
        self.transient(parent); self.grab_set(); self.title(title)
        
        summary_str = ""
        total_att_start = sum(u['qty'] for u in results['attacker_roster_final'])
        total_att_survivors = sum(u['survivors'] for u in results['attacker_roster_final'])
        att_casualties = 100 * (1 - total_att_survivors / total_att_start) if total_att_start > 0 else 0
        summary_str += f"--- ATTACKER --- | Forces: {total_att_start} -> Survivors: {total_att_survivors} ({att_casualties:.1f}% casualties)\n"
        for unit in results['attacker_roster_final']:
            summary_str += f"  {unit['qty']}x {unit['name']:<15} | Losses: {unit['qty'] - unit['survivors']:<3} | Survivors: {unit['survivors']}\n"

        total_def_start = sum(u['qty'] for u in results['defender_roster_final'])
        total_def_survivors = sum(u['survivors'] for u in results['defender_roster_final'])
        def_casualties = 100 * (1 - total_def_survivors / total_def_start) if total_def_start > 0 else 0
        summary_str += f"\n--- DEFENDER --- | Forces: {total_def_start} -> Survivors: {total_def_survivors} ({def_casualties:.1f}% casualties)\n"
        for unit in results['defender_roster_final']:
            summary_str += f"  {unit['qty']}x {unit['name']:<15} | Losses: {unit['qty'] - unit['survivors']:<3} | Survivors: {unit['survivors']}\n"
            
        if results['captured_by_attacker'] or results['captured_by_defender']:
            summary_str += "\n--- SALVAGED & CAPTURED EQUIPMENT ---\n"
            if results['captured_by_attacker']:
                summary_str += " Acquired by Attacker:\n"
                for item in results['captured_by_attacker']:
                    summary_str += f"  {item['qty']}x {item['name']}\n"
            if results['captured_by_defender']:
                summary_str += " Acquired by Defender:\n"
                for item in results['captured_by_defender']:
                    summary_str += f"  {item['qty']}x {item['name']}\n"

        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(expand=True, fill=tk.BOTH)
        main_frame.rowconfigure(1, weight=1)
        main_frame.columnconfigure(0, weight=1)

        summary_frame = ttk.LabelFrame(main_frame, text="Battle Summary", padding=10)
        summary_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        summary_text = tk.Text(summary_frame, wrap=tk.WORD, font=("Courier", 10), height=14)
        summary_text.pack(expand=True, fill=tk.BOTH)
        summary_text.insert(tk.END, summary_str)
        summary_text.config(state=tk.DISABLED)

        log_frame = ttk.LabelFrame(main_frame, text="Combat Log", padding=10)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Courier", 10), height=20, width=90)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=log_text.yview)
        log_text['yscrollcommand'] = log_scroll.set
        
        log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")
        
        log_text.insert(tk.END, "\n".join(results['log']))
        log_text.config(state=tk.DISABLED)

        ttk.Button(main_frame, text="Close", command=self.destroy).grid(row=2, column=0, pady=(10, 0))

class ToolTip:
    def __init__(self, w, txt): self.w, self.txt, self.tw = w, txt, None; self.w.bind("<Enter>", self.show); self.w.bind("<Leave>", self.hide)
    def show(self, e): x, y, _, _ = self.w.bbox("insert"); x += self.w.winfo_rootx() + 25; y += self.w.winfo_rooty() + 25; self.tw = tk.Toplevel(self.w); self.tw.wm_overrideredirect(True); self.tw.wm_geometry(f"+{x}+{y}"); tk.Label(self.tw, text=self.txt, justify='left', background="#ffffe0", relief='solid', borderwidth=1).pack(ipadx=1)
    def hide(self, e):
        if self.tw: self.tw.destroy(); self.tw = None