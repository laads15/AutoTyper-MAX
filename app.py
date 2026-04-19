"""
Personal Auto Typer Application
================================

This simple Python application allows you to paste text and have it
typed into the currently focused application window with a configurable
delay between characters. The primary goal is convenience: it saves
you from manually re‑typing repetitive text.  This is provided for
personal use only — please do not use it to automate login forms or
otherwise circumvent anti‑automation checks.

The app uses Tkinter for its graphical interface and PyAutoGUI for
sending keystrokes to the operating system.  You will need to
install PyAutoGUI before running this app:

    pip install pyautogui

If you would like the application to copy text to the clipboard
instead of typing each character, you can modify the `type_text`
function to call `pyautogui.hotkey("ctrl", "v")` after copying the
contents to the clipboard using Python's `pyperclip` module.
"""

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
# Additional imports for file dialogs, configuration and JSON handling
import os
import json
from tkinter import filedialog, simpledialog
import threading
import random
import time
import datetime
import re
import string

# Optional dependencies for additional functionality.  These are
# imported conditionally so the program still runs even if they
# are not available.
try:
    import pyperclip  # type: ignore
except ImportError:
    pyperclip = None  # type: ignore

try:
    import winsound  # type: ignore
except ImportError:
    winsound = None  # type: ignore

try:
    # Optional dependency for global hotkeys. This library allows us to
    # register a system‑wide key combination to start the typing
    # process from anywhere.  On some platforms this may require
    # administrator privileges.
    import keyboard  # type: ignore
except ImportError:
    keyboard = None  # type: ignore


try:
    import pyautogui
except ImportError as e:
    raise SystemExit(
        "PyAutoGUI is required for this application. "
        "Please install it by running 'pip install pyautogui' in your terminal."
    ) from e


class AutoTyperApp(tk.Tk):
    """
    Main application class for the Auto Typer.

    In addition to the basic functionality of pasting text and typing it
    with configurable delays, this version supports launching the
    typing via a global hotkey and inserting regular pauses during the
    typing process.  The hotkey defaults to ``ctrl+alt+t`` but can be
    customized via the interface.  Pauses are defined by how many
    characters to type before pausing and how long the pause should
    last (in milliseconds).
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Personal Auto Typer")
        # Allow the window to be resizable so that users can view the
        # entire interface or adjust it to their screen.  Provide a
        # reasonable default size to accommodate all controls without
        # overflowing on smaller displays.
        self.geometry("750x700")
        self.resizable(True, True)

        # ------------------------------------------------------------------
        # Configuration and snippet paths and loading
        #
        # Determine file paths for persisted configuration and snippet data.
        # These files live in the user's home directory so settings and
        # snippets persist across launches of the application.
        self.config_path = os.path.join(os.path.expanduser("~"), ".auto_typer_config.json")
        self.snippets_path = os.path.join(os.path.expanduser("~"), ".auto_typer_snippets.json")

        # Initialise configuration and snippet lists.  load_config() and
        # load_snippets() will populate these from disk if the files exist.
        self.config = {}
        self.snippets = []

        # Load persisted settings and snippets before building the UI.  This
        # ensures that default values such as the "always on top" flag are
        # available when UI widgets are created.
        try:
            self.load_config()
        except Exception:
            self.config = {}
        try:
            self.load_snippets()
        except Exception:
            self.snippets = []

        # Create a notebook with two tabs: one for typing and one for help/instructions.
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, padx=0, pady=0)

        # ------------------------------------------------------------------
        # Create a scrollable container for the "Type" tab.  The content
        # of the first tab may grow beyond the height of the window as
        # additional features are added.  To ensure all controls remain
        # accessible on smaller displays, embed the content in a canvas
        # with a vertical scrollbar.  This pattern creates a frame
        # (``main_frame``) inside the canvas where all widgets are
        # placed, and automatically updates the scroll region when
        # content size changes.
        main_container = tk.Frame(self.notebook)
        # Canvas for scrollable content
        self.main_canvas = tk.Canvas(main_container, borderwidth=0, highlightthickness=0)
        # Vertical scrollbar linked to the canvas
        self.main_scrollbar = tk.Scrollbar(main_container, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        self.main_scrollbar.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)
        # Inner frame that holds the actual widgets for the "Type" tab
        main_frame = tk.Frame(self.main_canvas)
        self.main_canvas.create_window((0, 0), window=main_frame, anchor="nw")
        # As the content frame grows or shrinks, adjust the scroll region
        def on_main_frame_configure(event: tk.Event) -> None:
            try:
                self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
            except Exception:
                pass
        main_frame.bind("<Configure>", on_main_frame_configure)
        # Bind mouse wheel scrolling to the canvas so the user can
        # scroll through the controls with the wheel.  On Windows
        # ``event.delta`` is a multiple of 120.  Negative values
        # indicate scrolling down.  We divide by 120 to convert to
        # lines and invert the sign to scroll in the correct direction.
        def _on_mousewheel(event: tk.Event) -> None:
            try:
                # For Windows and MacOS
                if event.delta:
                    self.main_canvas.yview_scroll(-int(event.delta / 120), "units")
                # For Linux (delta is typically 0 and event.num used)
                elif event.num == 4:
                    self.main_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    self.main_canvas.yview_scroll(1, "units")
            except Exception:
                pass
        # Bind to both the canvas and the toplevel so that scrolling works
        # regardless of which widget the pointer is over.
        self.main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.main_canvas.bind_all("<Button-4>", _on_mousewheel)
        self.main_canvas.bind_all("<Button-5>", _on_mousewheel)
        # Allow the second column to expand when the window is resized
        main_frame.columnconfigure(1, weight=1)

        # Additional tabs: help and snippets
        help_frame = tk.Frame(self.notebook)
        snippets_frame = tk.Frame(self.notebook)
        # Add the tabs to the notebook (note: order matters)
        self.notebook.add(main_container, text="Type")
        self.notebook.add(help_frame, text="Help")
        self.notebook.add(snippets_frame, text="Snippets")
        # Build the snippet tab UI
        self.build_snippets_tab(snippets_frame)

        # Text input label and box on the main tab
        tk.Label(main_frame, text="Text to type:").grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        self.text_box = tk.Text(main_frame, width=40, height=10)
        self.text_box.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10))

        # Speed settings (WPM range) on the main tab
        tk.Label(main_frame, text="Min speed (WPM):").grid(row=2, column=0, padx=10, pady=(0, 0), sticky="w")
        tk.Label(main_frame, text="Max speed (WPM):").grid(row=2, column=1, padx=10, pady=(0, 0), sticky="w")
        self.min_speed_entry = tk.Entry(main_frame)
        self.min_speed_entry.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.min_speed_entry.insert(0, "30")  # default min speed 30 WPM
        self.max_speed_entry = tk.Entry(main_frame)
        self.max_speed_entry.grid(row=3, column=1, padx=10, pady=(0, 10), sticky="ew")
        self.max_speed_entry.insert(0, "60")  # default max speed 60 WPM

        # Pause settings on the main tab: define a random range of words between pauses and the pause length (ms).
        tk.Label(main_frame, text="Min words between pauses:").grid(row=4, column=0, padx=10, pady=(0, 0), sticky="w")
        tk.Label(main_frame, text="Max words between pauses:").grid(row=4, column=1, padx=10, pady=(0, 0), sticky="w")
        self.min_pause_words_entry = tk.Entry(main_frame)
        self.min_pause_words_entry.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.min_pause_words_entry.insert(0, "10")
        self.max_pause_words_entry = tk.Entry(main_frame)
        self.max_pause_words_entry.grid(row=5, column=1, padx=10, pady=(0, 10), sticky="ew")
        self.max_pause_words_entry.insert(0, "150")
        tk.Label(main_frame, text="Pause length (ms):").grid(row=6, column=0, padx=10, pady=(0, 0), sticky="w")
        self.pause_length_entry = tk.Entry(main_frame)
        self.pause_length_entry.grid(row=6, column=1, padx=10, pady=(0, 10), sticky="ew")
        self.pause_length_entry.insert(0, "1000")  # default: 1000 ms pause

        # Delay start setting on the main tab.  This allows the user to
        # schedule typing to begin after a brief countdown.  Specify the
        # delay in seconds; use 0.0 for immediate start.  Place this row
        # after the pause settings.
        tk.Label(main_frame, text="Delay start (sec):").grid(row=7, column=0, padx=10, pady=(0, 0), sticky="w")
        self.delay_start_entry = tk.Entry(main_frame)
        self.delay_start_entry.grid(row=7, column=1, padx=10, pady=(0, 10), sticky="ew")
        self.delay_start_entry.insert(0, "0.0")

        # Hotkey setting on the main tab (row shifted down by one)
        tk.Label(main_frame, text="Hotkey (e.g. ctrl+alt+t):").grid(row=8, column=0, padx=10, pady=(0, 0), sticky="w")
        self.hotkey_entry = tk.Entry(main_frame)
        self.hotkey_entry.grid(row=8, column=1, padx=10, pady=(0, 0), sticky="ew")
        self.hotkey_entry.insert(0, "ctrl+alt+t")

        # Start button and instructions on the main tab.  The start button now
        # resides on row 9 because rows 7–8 are used by the delay start and
        # hotkey settings.  The instructions follow on row 10.
        start_button = ttk.Button(main_frame, text="Type into focused app", command=self.start_typing, style="Rounded.TButton")
        start_button.grid(row=9, column=0, columnspan=2, padx=10, pady=(0, 10))

        self.instruction_label = tk.Label(
            main_frame,
            text=(
                "Click into the destination application and field before pressing "
                "'Type into focused app', or use the global hotkey to start.\n"
                "The hotkey can be customized above.\n"
                "While typing occurs, avoid moving the cursor or interacting with the keyboard."
            ),
            wraplength=300,
            justify="left",
            fg="#555555"
        )
        self.instruction_label.grid(row=10, column=0, columnspan=2, padx=10, pady=(0, 10))

        # Feature toggles
        # These BooleanVars drive options such as randomized delays,
        # natural pauses, newline handling, multi-field mode, typo
        # simulation, invisible paste and keyboard sounds.  The default
        # values mirror the screenshot provided by the user: random
        # delays and natural pauses enabled; keep line breaks off; multi-field off; typo
        # simulation off; invisible paste off; keyboard sounds off.
        self.random_delays_var = tk.BooleanVar(value=True)
        self.natural_pauses_var = tk.BooleanVar(value=True)
        self.keep_line_breaks_var = tk.BooleanVar(value=False)
        self.multi_field_var = tk.BooleanVar(value=False)
        self.typo_sim_var = tk.BooleanVar(value=False)
        self.invisible_paste_var = tk.BooleanVar(value=False)
        self.keyboard_sounds_var = tk.BooleanVar(value=False)

        # Alarm option: ring a sound when typing completes.  When
        # enabled, the app will play a short audible alert after the
        # typing thread finishes.  This is helpful for long typing
        # operations where you may be away from the screen.
        self.alarm_var = tk.BooleanVar(value=False)

        # Preserve formatting option: when enabled, the text will be typed
        # exactly as pasted, without replacing newlines or spaces.  This
        # overrides the "Keep line breaks" and "Multi-field" settings.
        self.preserve_formatting_var = tk.BooleanVar(value=False)

        # Always on top option: when enabled, keep the application
        # window above all others.  The default value is loaded from
        # the persisted configuration if available.
        self.always_on_top_var = tk.BooleanVar(value=bool(self.config.get("always_on_top", False)))

        # Finish time, finish-by time and countdown timer variables.
        # When a finish duration is provided, the application will
        # compute an average typing rate to complete the text within
        # that time.  A finish-by value specifies an absolute due
        # time of day (HH:MM) by which typing should finish.  These
        # values are restored from the configuration if present.
        self.finish_time_entry_value = self.config.get("finish_in", self.config.get("finish_time", ""))
        self.finish_by_entry_value = self.config.get("finish_by", "")
        self.countdown_timer_var_value = bool(self.config.get("countdown_timer", False))

        # Create a labelled frame to group the option checkboxes on the main tab.
        # Using a LabelFrame makes the UI more organised and self‑describing.
        options_frame = tk.LabelFrame(main_frame, text="Options")
        # Shift the options group down to row 11 to accommodate the delay start and hotkey rows
        options_frame.grid(row=11, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="ew")

        toggles = [
            ("Randomised delays", self.random_delays_var),
            ("Natural pauses", self.natural_pauses_var),
            ("Keep line breaks", self.keep_line_breaks_var),
            ("Multi-field", self.multi_field_var),
            ("Typo simulation", self.typo_sim_var),
            ("Invisible paste", self.invisible_paste_var),
            ("Keyboard sounds", self.keyboard_sounds_var),
            ("Preserve formatting", self.preserve_formatting_var),
            ("Alarm when done", self.alarm_var),
        ]
        # Arrange toggles in the options frame, two per row
        for idx, (label_text, var) in enumerate(toggles):
            col = idx % 2
            row = idx // 2
            cb = tk.Checkbutton(options_frame, text=label_text, variable=var)
            cb.grid(row=row, column=col, padx=5, pady=2, sticky="w")

        # Snippet selection drop‑down: allow the user to choose a saved snippet
        # for quick insertion.  This appears below the options frame.  When a
        # snippet is chosen, its content will load into the main text box.
        tk.Label(main_frame, text="Snippet:").grid(row=12, column=0, padx=10, pady=(0, 0), sticky="w")
        self.snippet_var = tk.StringVar(value="")
        self.snippet_combobox = ttk.Combobox(main_frame, textvariable=self.snippet_var, state="readonly")
        self.snippet_combobox.grid(row=12, column=1, padx=10, pady=(0, 0), sticky="ew")
        # Populate available snippet names
        self.populate_snippet_combobox()
        # Bind selection event
        self.snippet_combobox.bind("<<ComboboxSelected>>", lambda e: self.on_snippet_selected())

        # Progress display: show a progress bar and remaining time estimate.
        # This uses a DoubleVar to track the percent complete.  Place it at row 13.
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100.0)
        self.progress_bar.grid(row=13, column=0, columnspan=2, padx=10, pady=(5, 0), sticky="ew")
        self.remaining_time_label = tk.Label(main_frame, text="", fg="#555555")
        self.remaining_time_label.grid(row=14, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")

        # Finish in entry: specify a target duration in hours and minutes to complete typing.
        # Accept values like "1:30" for 1 hour 30 minutes or a single number for hours.
        tk.Label(main_frame, text="Finish in (h:mm):").grid(row=15, column=0, padx=10, pady=(0, 0), sticky="w")
        self.finish_time_entry = tk.Entry(main_frame)
        self.finish_time_entry.grid(row=15, column=1, padx=10, pady=(0, 10), sticky="ew")
        # Finish by entry: specify an absolute time of day (HH:MM) by which typing should finish.
        tk.Label(main_frame, text="Finish by (HH:MM):").grid(row=16, column=0, padx=10, pady=(0, 0), sticky="w")
        self.finish_by_entry = tk.Entry(main_frame)
        self.finish_by_entry.grid(row=16, column=1, padx=10, pady=(0, 10), sticky="ew")
        # Prepopulate finish-in and finish-by from prior configuration if available.
        fin_val = getattr(self, "finish_time_entry_value", "")
        if fin_val:
            self.finish_time_entry.insert(0, str(fin_val))
        fin_by_val = getattr(self, "finish_by_entry_value", "")
        if fin_by_val:
            self.finish_by_entry.insert(0, str(fin_by_val))
        # Countdown timer toggle: when enabled, show a countdown of remaining time.  Move to next row.
        self.countdown_timer_var = tk.BooleanVar(value=getattr(self, "countdown_timer_var_value", False))
        tk.Checkbutton(main_frame, text="Countdown timer", variable=self.countdown_timer_var).grid(row=17, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="w")

        # Flag to prevent multiple concurrent typing threads
        self._is_typing = False
        # Event used to signal the typing thread to stop early.  When
        # ``_stop_event`` is set, the typing loop will break out at
        # the next convenient point.  This allows the same hotkey
        # handler to toggle the typing on and off.
        self._stop_event = threading.Event()
        self._typing_thread: threading.Thread | None = None

        # Register the hotkey if the keyboard module is available
        self.register_hotkey()

        # Apply a dark mode theme to the interface
        self.apply_dark_mode()

        # Build the help tab with collapsible sections
        self.build_help_tab(help_frame)

        # ------------------------------------------------------------------
        # Restore saved UI state and register snippet hotkeys
        # After building all UI widgets we restore any persisted settings
        # from the configuration file.  This sets toggle states, text
        # contents, always-on-top flag and selects the last snippet.
        try:
            self.restore_state()
        except Exception:
            pass
        # Register snippet-specific hotkeys now that snippets are loaded
        self.register_snippet_hotkeys()
        # Apply always-on-top state from the saved configuration
        self.toggle_always_on_top()
        # Bind the window close event to persist configuration and snippets
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def start_typing(self) -> None:
        """
        Validate user input and start the typing thread.

        This method reads the text and delay values from the UI, validates
        them, and launches a new thread to perform the typing. Running
        the typing in a separate thread prevents the GUI from freezing.
        """
        text = self.text_box.get("1.0", "end-1c")
        if not text:
            messagebox.showerror("Error", "Please enter some text to type.")
            return
        # Resolve any placeholders in the text before starting
        text = self.resolve_placeholders(text)

        try:
            min_speed = float(self.min_speed_entry.get())
            max_speed = float(self.max_speed_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Min and Max speed must be numeric.")
            return

        # Validate WPM range
        if min_speed <= 0 or max_speed <= 0:
            messagebox.showerror("Error", "Speed values must be positive.")
            return
        if min_speed > max_speed:
            messagebox.showerror("Error", "Min speed must not exceed max speed.")
            return

        # Parse pause settings (words between pauses and pause length)
        try:
            min_pause_words = int(self.min_pause_words_entry.get())
            max_pause_words = int(self.max_pause_words_entry.get())
            pause_length = int(self.pause_length_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Pause settings must be numeric.")
            return
        if min_pause_words < 0 or max_pause_words < 0 or pause_length < 0:
            messagebox.showerror("Error", "Pause values must be non-negative.")
            return
        if min_pause_words > max_pause_words:
            messagebox.showerror("Error", "Min words between pauses must not exceed max words.")
            return

        # Check if a typing operation is already in progress
        if self._is_typing:
            messagebox.showinfo("Already typing", "A typing operation is already in progress.")
            return

        # Capture toggle states for advanced functionality
        random_delays = self.random_delays_var.get() if hasattr(self, "random_delays_var") else True
        natural_pauses = self.natural_pauses_var.get() if hasattr(self, "natural_pauses_var") else True
        keep_line_breaks = self.keep_line_breaks_var.get() if hasattr(self, "keep_line_breaks_var") else False
        multi_field = self.multi_field_var.get() if hasattr(self, "multi_field_var") else False
        typo_sim = self.typo_sim_var.get() if hasattr(self, "typo_sim_var") else False
        invisible_paste = self.invisible_paste_var.get() if hasattr(self, "invisible_paste_var") else False
        keyboard_sounds = self.keyboard_sounds_var.get() if hasattr(self, "keyboard_sounds_var") else False
        preserve_formatting = self.preserve_formatting_var.get() if hasattr(self, "preserve_formatting_var") else False
        alarm = self.alarm_var.get() if hasattr(self, "alarm_var") else False
        # Delay start in seconds
        try:
            delay_start = float(self.delay_start_entry.get())
        except Exception:
            delay_start = 0.0

        # Parse duration and due time inputs.  Durations are
        # specified in hours and minutes (e.g. "1:30" or "0:45").
        # Due times are absolute times of day (HH:MM).  Blank or
        # malformed inputs disable the corresponding modes.

        def parse_duration_str(val: str) -> float:
            """Convert a human‑readable duration string to seconds.

            Accepts formats like "1:30" (1 hour 30 minutes),
            "0:45" (45 minutes), or a single number (treated as
            hours).  Returns 0.0 for empty or invalid inputs.
            """
            val = val.strip()
            if not val:
                return 0.0
            try:
                # If the value contains a colon, treat it as H:MM
                if ":" in val:
                    parts = val.split(":")
                    if len(parts) != 2:
                        return 0.0
                    hours = float(parts[0]) if parts[0] else 0.0
                    minutes = float(parts[1]) if parts[1] else 0.0
                    return max(hours * 3600.0 + minutes * 60.0, 0.0)
                else:
                    # Treat single number as hours
                    hours = float(val)
                    return max(hours * 3600.0, 0.0)
            except Exception:
                return 0.0

        def parse_due_time_str(val: str) -> float:
            """Convert a due time string (HH:MM) to a Unix timestamp.

            Interprets the provided time as an absolute time of day
            today or tomorrow.  If the time has already passed today,
            schedule for the same time on the next day.  Returns 0.0
            for blank or invalid inputs.
            """
            val = val.strip()
            if not val:
                return 0.0
            try:
                # Extract hours and minutes
                if ":" not in val:
                    return 0.0
                hour_str, minute_str = val.split(":", 1)
                hour = int(hour_str)
                minute = int(minute_str)
                # Clamp values to valid ranges
                if hour < 0 or hour > 23 or minute < 0 or minute > 59:
                    return 0.0
                now = datetime.datetime.now()
                due_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                # If the due time has already passed today, schedule for the next day
                if due_time <= now:
                    due_time += datetime.timedelta(days=1)
                return due_time.timestamp()
            except Exception:
                return 0.0

        # Determine requested finish duration and due time
        finish_in = parse_duration_str(self.finish_time_entry.get())
        due_timestamp = parse_due_time_str(self.finish_by_entry.get() if hasattr(self, "finish_by_entry") else "")

        # Mark typing as active and reset stop event
        self._is_typing = True
        self._stop_event.clear()

        # Reset progress indicators
        self.progress_var.set(0.0)
        self.remaining_time_label.config(text="")
        # Launch the typing in a separate thread and save the reference
        self._typing_thread = threading.Thread(
            target=self.type_text,
            args=(
                text,
                min_speed,
                max_speed,
                min_pause_words,
                max_pause_words,
                pause_length,
                random_delays,
                natural_pauses,
                keep_line_breaks,
                multi_field,
                typo_sim,
                invisible_paste,
                keyboard_sounds,
                preserve_formatting,
                alarm,
                delay_start,
                finish_in,
                due_timestamp,
            ),
            daemon=True,
        )
        self._typing_thread.start()

        # Persist current state to disk
        self.save_config()

    def type_text(
        self,
        text: str,
        min_wpm: float,
        max_wpm: float,
        min_pause_words: int = 0,
        max_pause_words: int = 0,
        pause_length_ms: int = 0,
        random_delays: bool = True,
        natural_pauses: bool = True,
        keep_line_breaks: bool = False,
        multi_field: bool = False,
        typo_sim: bool = False,
        invisible_paste: bool = False,
        keyboard_sounds: bool = False,
        preserve_formatting: bool = False,
        alarm: bool = False,
        delay_start: float = 0.0,
        finish_in_seconds: float = 0.0,
        due_timestamp: float = 0.0,
    ) -> None:
        """
        Type the provided text into the currently focused application.

        Parameters
        ----------
        text : str
            The text to type.
        min_wpm, max_wpm : float
            Minimum and maximum typing speed in words per minute (WPM).
            If ``random_delays`` is true, the WPM is chosen uniformly at
            random between these values for each character. If false, a
            constant WPM equal to the average of these values is used.
        min_pause_words, max_pause_words : int
            Range for random word count between pauses (if ``natural_pauses``
            is true). A long pause will be inserted after a random number of
            words within this range. If either value is zero, pauses are disabled.
        pause_length_ms : int
            Length of each long pause in milliseconds.
        random_delays : bool
            Whether to vary the typing speed per character. If false,
            a constant speed equal to the average WPM is used.
        natural_pauses : bool
            Whether to insert regular pauses.  If false, pauses are disabled.
        keep_line_breaks : bool
            Preserve newline characters.  If false, newlines will be replaced
            with spaces (unless multi_field is true).
        multi_field : bool
            Treat each line of the text as input for separate fields.  When
            enabled, newline characters result in a ``Tab`` key press instead
            of a newline.
        typo_sim : bool
            Simulate occasional typos by inserting a wrong character and
            pressing backspace to correct it.  The probability of a typo is
            hard-coded to ~3% per character.
        invisible_paste : bool
            If true, the entire text is copied to the clipboard and
            pasted in one operation using Ctrl+V, ignoring other settings.
        keyboard_sounds : bool
            Play a simple beep sound for each character typed (Windows only).
        preserve_formatting : bool
            When true, type the text exactly as it was pasted.  This
            overrides ``keep_line_breaks`` and ``multi_field`` so that
            newline characters are preserved and will not be converted
            to tabs or spaces.  Use this option to maintain
            formatting, spacing and line breaks from the original text.
        alarm : bool
            Play a notification sound when typing completes.  If
            enabled, a short audible alert will sound after the
            typing thread finishes, regardless of whether it
            completed naturally or was stopped via the hotkey.

        finish_in_seconds : float
            If greater than zero, specifies the desired duration (in
            seconds) within which the entire text should be typed.  A
            dynamic scheduling algorithm varies the per‑character delay
            around the average rate needed to finish within this
            timeframe.  Natural pauses and random delays are
            disabled in this mode.

        due_timestamp : float
            Absolute Unix timestamp representing a target “finish by”
            time.  When provided, the application will pace typing so
            that the session completes at or before this timestamp,
            varying the speed subtly to maintain a natural feel.  If
            the scheduled time has already passed, this parameter
            should be zero.
        """
        try:
            # small initial delay to allow the user to refocus the target app
            # plus any user‑specified delay for scheduling the typing to begin.
            total_delay = 1.5 + max(0.0, delay_start)
            time.sleep(total_delay)

            # If invisible paste is selected and pyperclip is available
            if invisible_paste and pyperclip is not None:
                pyperclip.copy(text)
                time.sleep(0.1)
                pyautogui.hotkey("ctrl", "v")
                return

            # Preprocess text based on line break settings
            # If the user requests to preserve formatting exactly as pasted,
            # override the line break/multi‑field settings.  In this mode
            # we will not replace newline characters with spaces and we will
            # avoid converting newlines into TAB presses.  This ensures the
            # output reflects the original pasted text verbatim.
            if preserve_formatting:
                keep_line_breaks = True
                multi_field = False

            # Normalize CRLF to LF
            normalized_text = text.replace("\r", "")
            # Replace newlines with spaces only when the user has not
            # requested to keep line breaks and the multi‑field mode is off.
            if not keep_line_breaks and not multi_field:
                normalized_text = normalized_text.replace("\n", " ")

            # Compute the total number of characters and words that will be typed.
            # The total words are approximated by splitting on whitespace after
            # normalisation.  These values are stored on the ``self`` object
            # so the progress display can compute percentages and word counts.

            total_chars = len(normalized_text)
            # Count words using regex to match non‑whitespace sequences
            total_words = len(re.findall(r"\S+", normalized_text)) if normalized_text else 0
            self.total_chars = total_chars
            self.total_words = total_words
            # Initialise counters that track how many characters and words
            # have been typed so far.  These are updated inside the
            # character loop and referenced in the progress display.
            self.chars_typed = 0
            self.words_typed = 0

            # Helper to convert WPM to a per‑character delay.  One word is
            # standardised to five characters【524493561122292†L239-L244】.
            def wpm_to_delay(wpm: float) -> float:
                return 60.0 / (wpm * 5.0)

            # Determine whether we are operating under a target schedule.
            # A target schedule exists if a due timestamp is provided or
            # a finish‑in duration is greater than zero.  Under a target
            # schedule, delays are dynamically adjusted to finish on time.
            use_target_mode = False
            target_finish_timestamp: float | None = None
            predicted_total_time: float | None = None

            # Finish‑by mode: absolute due timestamp
            if due_timestamp > 0 and total_chars > 0:
                use_target_mode = True
                target_finish_timestamp = due_timestamp
                # Predicted total time for progress display equals the
                # difference between the due time and the current time.  It
                # will be refined after the initial delay.
                predicted_total_time = max(target_finish_timestamp - time.time(), 0.0)
                # Disable pauses and random delays under schedule to allow
                # fine‑grained control of the pace.
                min_pause_words = max_pause_words = 0
                pause_length_ms = 0
                natural_pauses = False
                random_delays = False
            # Finish‑in mode: relative duration
            elif finish_in_seconds > 0.0 and total_chars > 0:
                use_target_mode = True
                # Predicted total time equals the requested duration.
                predicted_total_time = finish_in_seconds
                # target_finish_timestamp will be set later after the
                # initial delay; it will be ``start_time + finish_in_seconds``.
                target_finish_timestamp = None
                # Disable pauses and random delays
                min_pause_words = max_pause_words = 0
                pause_length_ms = 0
                natural_pauses = False
                random_delays = False
            else:
                # No explicit schedule; compute a baseline predicted time
                # based on the user’s WPM settings.  Use the average of
                # the min and max speeds to estimate the total duration
                # and multiply by a random factor to avoid mechanical
                # predictability.  This value is used for progress
                # estimation but does not control the actual typing rate.
                avg_wpm_estimate = (min_wpm + max_wpm) / 2.0
                baseline_delay = wpm_to_delay(avg_wpm_estimate) if avg_wpm_estimate > 0 else 0.0
                predicted_total_time = baseline_delay * total_chars
                # Introduce a modest random factor (±10%) so that
                # repeated runs with the same text do not produce
                # identical estimated durations.
                if predicted_total_time > 0:
                    # Introduce a random factor to avoid mechanical
                    # predictability.  Use a slightly wider range (±15%)
                    # to reflect natural variation in human typing speed.
                    predicted_total_time *= random.uniform(0.85, 1.15)

            # Store predicted_total_time so the progress display can
            # calculate remaining time.  ``target_finish_timestamp`` may be
            # ``None`` until we determine the actual start time for
            # finish‑in mode.
            self.predicted_total_time = predicted_total_time or 0.0
            self.target_finish_time = target_finish_timestamp
            self.use_target_mode = use_target_mode

            # Determine the delay generator.  Under a target schedule we
            # adjust delays dynamically based on remaining characters and
            # time.  Otherwise we use either random or fixed delays based
            # on the WPM settings and user toggles.
            # ``average_delay`` is used for progress estimation in
            # non‑scheduled modes; under a schedule it is recomputed after
            # the initial delay when the start time is known.

            delay_generator: callable[[], float]
            average_delay: float = 0.0

            if use_target_mode:
                # The actual delay generator will be defined after we
                # establish the start time (below).  Here, we prepare a
                # placeholder; average_delay will be set later.
                def next_delay() -> float:
                    return 0.01
            else:
                # Non‑scheduled mode: choose WPM per character or use a
                # fixed average speed, depending on random_delays.
                if random_delays:
                    def next_delay() -> float:
                        chosen_wpm = random.uniform(min_wpm, max_wpm)
                        return wpm_to_delay(chosen_wpm)
                    # Estimate average delay using the mean of the WPM range
                    avg_wpm = (min_wpm + max_wpm) / 2.0
                    average_delay = wpm_to_delay(avg_wpm) if avg_wpm > 0 else 0.0
                else:
                    avg_wpm = (min_wpm + max_wpm) / 2.0
                    fixed_delay = wpm_to_delay(avg_wpm) if avg_wpm > 0 else 0.0
                    def next_delay() -> float:
                        return fixed_delay
                    average_delay = fixed_delay
                # If pauses are not requested or any values are zero,
                # disable them entirely.
                if (not natural_pauses) or min_pause_words == 0 or max_pause_words == 0 or pause_length_ms == 0:
                    min_pause_words = max_pause_words = 0
                    pause_length_ms = 0

            # Establish the start time just before typing begins and
            # compute any schedule‑dependent values.  For finish‑in
            # mode the target finish timestamp is set relative to the
            # start time; for due‑time mode it remains unchanged.  We
            # also compute the average delay per character for
            # progress estimation under a schedule and define the
            # dynamic ``next_delay`` generator when a schedule is
            # active.
            self.start_time = time.time()
            if self.use_target_mode:
                # If this is a finish‑in schedule, set the target
                # finish time relative to the start time.  For a due
                # schedule (target_finish_time is already set to
                # an absolute timestamp) no adjustment is needed.
                if self.target_finish_time is None and finish_in_seconds > 0:
                    self.target_finish_time = self.start_time + finish_in_seconds
                # Recompute predicted total time based on the actual start
                # time.  ``predicted_total_time`` represents the
                # difference between the target finish time and the start
                # time.  This value will be used to estimate remaining
                # time for the progress display.
                if self.target_finish_time is not None:
                    self.predicted_total_time = max(self.target_finish_time - self.start_time, 0.0)
                # Define a dynamic delay generator that adjusts the
                # pacing according to remaining time and remaining
                # characters.  A random factor introduces natural
                # variability while keeping the overall timing on
                # schedule.  If remaining_chars or remaining_time is
                # negative, a small minimum delay is used to avoid
                # division by zero.
                def next_delay() -> float:
                    remaining_chars = max(self.total_chars - self.chars_typed, 1)
                    remaining_time = (self.target_finish_time - time.time()) if self.target_finish_time is not None else 0.0
                    base = remaining_time / remaining_chars if remaining_chars > 0 else 0.001
                    # Constrain the random factor more tightly when the
                    # schedule is very constrained (very short base),
                    # otherwise allow wider variation.  The factor
                    # ranges between 0.85 and 1.15.
                    factor = random.uniform(0.85, 1.15)
                    delay = base * factor
                    return max(delay, 0.001)
                # When following a schedule, pauses are disabled; this
                # was already done above when configuring the schedule.
                average_delay = self.predicted_total_time / self.total_chars if self.total_chars > 0 else 0.0
            else:
                # In non‑scheduled mode record the start time and use
                # the precomputed ``average_delay`` and ``next_delay``.
                # ``predicted_total_time`` remains as previously
                # calculated based on the WPM settings.
                pass

            # Initialize pause and word tracking and progress
            chars_typed = 0
            words_since_pause = 0
            next_pause_threshold = random.randint(min_pause_words, max_pause_words) if max_pause_words > 0 else 0
            inside_word = False
            lines = normalized_text.split("\n")
            for idx, line in enumerate(lines):
                for ch in line:
                    # Check if typing should stop (for hotkey toggling)
                    if self._stop_event.is_set():
                        return
                    # Simulate typos if enabled
                    if typo_sim and random.random() < 0.03:
                        wrong_char = random.choice(string.ascii_lowercase)
                        pyautogui.typewrite(wrong_char)
                        if keyboard_sounds and winsound is not None:
                            winsound.Beep(800, 30)
                        time.sleep(next_delay())
                        pyautogui.press("backspace")
                        if keyboard_sounds and winsound is not None:
                            winsound.Beep(800, 30)
                    # Type the actual character
                    pyautogui.typewrite(ch)
                    if keyboard_sounds and winsound is not None:
                        winsound.Beep(1200, 30)
                    # Update character and word counters
                    self.chars_typed += 1
                    # Word boundary detection and word counter update
                    if ch.isspace():
                        if inside_word:
                            words_since_pause += 1
                            self.words_typed += 1
                        inside_word = False
                    else:
                        inside_word = True
                    # Send progress update to the main thread.  The
                    # update handler will compute percentages, time
                    # remaining and schedule status based on these
                    # counters and configuration stored on ``self``.
                    if self.total_chars > 0:
                        self.after(0, self.update_progress_display, self.chars_typed, self.words_typed)
                    # Insert pause when threshold reached
                    if next_pause_threshold > 0 and words_since_pause >= next_pause_threshold:
                        time.sleep(pause_length_ms / 1000.0)
                        words_since_pause = 0
                        next_pause_threshold = random.randint(min_pause_words, max_pause_words) if max_pause_words > 0 else 0
                    # Wait between characters
                    time.sleep(next_delay())
                # At end of line, handle newline or tab
                if idx < len(lines) - 1:
                    # If the line ended with a word (no trailing space), count it
                    if inside_word:
                        words_since_pause += 1
                        self.words_typed += 1
                        inside_word = False
                    if multi_field:
                        pyautogui.press("tab")
                    else:
                        if keep_line_breaks:
                            pyautogui.typewrite("\n")
                            if keyboard_sounds and winsound is not None:
                                winsound.Beep(1200, 30)
                    # Send progress update after newline
                    if self.total_chars > 0:
                        self.after(0, self.update_progress_display, self.chars_typed, self.words_typed)
                    # Insert pause after newline if needed
                    if next_pause_threshold > 0 and words_since_pause >= next_pause_threshold:
                        time.sleep(pause_length_ms / 1000.0)
                        words_since_pause = 0
                        next_pause_threshold = random.randint(min_pause_words, max_pause_words) if max_pause_words > 0 else 0
            
        finally:
            # Mark typing finished and clear stop flag
            self._is_typing = False
            self._stop_event.clear()

            # Ring an alarm if requested when typing completes.  We
            # perform the alarm in the typing thread to avoid
            # interfering with the main GUI loop.  On Windows, use
            # winsound.Beep if available; otherwise fall back to
            # Tkinter's bell() which triggers a system alert sound.
            if alarm:
                try:
                    # If winsound is available, play a series of beeps.
                    if winsound is not None:
                        for _ in range(3):
                            winsound.Beep(1000, 400)
                            time.sleep(0.2)
                    else:
                        # Use Tkinter's bell; schedule on main thread
                        self.after(0, self.bell)
                except Exception:
                    # As a last resort, emit the ASCII bell character
                    try:
                        import sys
                        sys.stdout.write("\a")
                        sys.stdout.flush()
                    except Exception:
                        pass

            # Ensure progress shows complete when finished.  This
            # update runs on the main thread.  Provide the total
            # character and word counts so the progress bar reaches
            # 100% and the status reflects completion.
            try:
                self.after(0, self.update_progress_display, getattr(self, "total_chars", 0), getattr(self, "total_words", 0))
            except Exception:
                pass

    def register_hotkey(self) -> None:
        """
        Register a global hotkey to start typing.

        This method reads the hotkey string from ``self.hotkey_entry``
        and registers it with the ``keyboard`` module.  If the
        ``keyboard`` module is not available, no hotkey will be
        registered and the application will display an informational
        message.  On some platforms this may require elevated
        privileges.
        """
        # If the keyboard module isn't available, we can't register hotkeys
        if keyboard is None:
            self.hotkey_entry.configure(state="disabled")
            self.hotkey_entry.insert(0, "Install 'keyboard' module for hotkey support")
            return

        # Clear any existing hotkeys, depending on available API.  Some
        # versions of the keyboard library provide clear_all_hotkeys(),
        # while others expose unhook_all_hotkeys().  We attempt the
        # available one to avoid AttributeError on older versions.
        try:
            if hasattr(keyboard, "clear_all_hotkeys"):
                keyboard.clear_all_hotkeys()
            elif hasattr(keyboard, "unhook_all_hotkeys"):
                keyboard.unhook_all_hotkeys()
        except Exception:
            # If clearing fails, ignore – registering a new hotkey will
            # still work, but multiple registrations may stack.
            pass

        hotkey_str = self.hotkey_entry.get().strip()
        if not hotkey_str:
            return

        def on_hotkey():
            """
            Callback triggered by the global hotkey.

            If a typing operation is currently running, this will stop
            the typing by signalling the stop event. Otherwise it
            starts a new typing operation.  We schedule the calls on
            the Tkinter main thread using ``after`` to avoid
            threading issues.
            """
            if self._is_typing:
                self.after(0, self.stop_typing)
            else:
                self.after(0, self.start_typing)

        try:
            keyboard.add_hotkey(hotkey_str, on_hotkey)
        except Exception as exc:  # Catch broad exceptions from keyboard library
            messagebox.showerror(
                "Hotkey error",
                f"Failed to register hotkey '{hotkey_str}'. Make sure the string is valid and you have the necessary privileges.\n\n{exc}",
            )

    def stop_typing(self) -> None:
        """
        Signal the current typing thread to stop early.

        This method sets the internal stop event so that the typing
        loop will exit at the next opportunity.  It does not block
        waiting for the thread to finish; the thread will clean up and
        reset the state in its ``finally`` block.  If no typing
        operation is in progress, this does nothing.
        """
        if self._is_typing:
            # Signal the typing loop to exit
            self._stop_event.set()

    def apply_dark_mode(self) -> None:
        """
        Apply a dark colour scheme to the application.

        This helper iterates through all child widgets and sets background
        and foreground colours appropriate for a dark theme.  It
        customizes text boxes, entries, labels, buttons and
        checkbuttons.  System window chrome is not affected.
        """
        # Define a palette for a dark, neumorphic style.  The base
        # background is quite dark; surfaces (frames and panels) are a
        # slightly lighter shade.  The accent colour provides contrast
        # for interactive elements such as buttons and progress bars.
        base_bg = "#1e1f26"         # very dark base for the window
        surface_bg = "#282a36"      # slightly lighter surface for frames
        light_fg = "#e6e6e6"        # light foreground for text
        accent_bg = "#474b7c"       # muted purple accent for buttons
        button_active = "#5b5f8c"   # slightly brighter accent on hover

        # Configure a rounded button style using ttk.  This style will
        # make the primary buttons appear with softer edges and
        # consistent colours in dark mode.  Note that full rounded
        # corners are not directly supported by Tkinter themes, but
        # reducing the border width and using padding yields a
        # smoother appearance.
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        # Configure a rounded button style with neumorphic colours
        style.configure(
            "Rounded.TButton",
            background=accent_bg,
            foreground=light_fg,
            padding=(10, 6),
            borderwidth=0,
            relief="flat"
        )
        style.map(
            "Rounded.TButton",
            background=[("active", button_active)],
            foreground=[("active", light_fg)],
        )

        # Style the notebook and tabs with neumorphic colours
        style.configure(
            "TNotebook",
            background=base_bg,
            borderwidth=0,
        )
        style.configure(
            "TNotebook.Tab",
            background=base_bg,
            foreground=light_fg,
            padding=(12, 6)
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", accent_bg)],
            foreground=[("selected", light_fg)],
        )

        # Apply base colour to root
        self.configure(bg=base_bg)

        def style_widget(widget: tk.Misc) -> None:
            """
            Recursively apply dark colours to supported widgets.

            Some themed ``ttk`` widgets (like ``ttk.Combobox`` and ``ttk.Progressbar``)
            do not accept ``bg`` or ``fg`` configuration options.  To avoid
            crashing the application, each ``configure`` call is wrapped in a
            ``try``/``except`` block.  Unsupported options are silently
            ignored.  All children are recursed into regardless of type.
            """
            for child in widget.winfo_children():
                # Text widgets: dark background, light text and caret.  Tkinter Text supports
                # the insertbackground option, while Entry widgets may not accept it.  We
                # separate the handling to avoid passing unsupported options to ttk
                # entries.  Only native Tk widgets receive these settings.
                if isinstance(child, tk.Text):
                    try:
                        child.configure(bg=surface_bg, fg=light_fg)
                        child.configure(insertbackground=light_fg)
                    except Exception:
                        pass
                elif type(child) is tk.Entry:
                    try:
                        child.configure(bg=surface_bg, fg=light_fg)
                    except Exception:
                        pass
                    try:
                        child.configure(insertbackground=light_fg)
                    except Exception:
                        pass
                # Standard Button widgets: use accent colour and update active states
                elif isinstance(child, tk.Button):
                    try:
                        child.configure(
                            bg=accent_bg,
                            fg=light_fg,
                            activebackground=button_active,
                            activeforeground=light_fg,
                            relief="flat",
                            bd=1
                        )
                    except Exception:
                        pass
                # Checkbuttons: dark background, light text, match check colours
                elif isinstance(child, tk.Checkbutton):
                    try:
                        child.configure(
                            bg=surface_bg,
                            fg=light_fg,
                            selectcolor=surface_bg,
                            activebackground=surface_bg,
                            activeforeground=light_fg,
                            relief="flat",
                            bd=1
                        )
                    except Exception:
                        pass
                # Labels: dark background with light text
                elif isinstance(child, tk.Label):
                    try:
                        child.configure(bg=surface_bg, fg=light_fg)
                    except Exception:
                        pass
                # Frames (including ttk.Frame and tk.Frame): set only the background.
                # Many ttk widgets inherit from Frame but may not accept ``bg``.
                elif isinstance(child, (tk.Frame, ttk.Frame)):
                    try:
                        child.configure(bg=surface_bg, bd=1, relief="flat")
                    except Exception:
                        pass
                # Listboxes: dark list background, light items and highlighted selection
                elif isinstance(child, tk.Listbox):
                    try:
                        child.configure(
                            bg=surface_bg,
                            fg=light_fg,
                            selectbackground=accent_bg,
                            selectforeground=light_fg,
                            relief="flat",
                            bd=1
                        )
                    except Exception:
                        pass
                else:
                    # For any other widget types (including ttk widgets) attempt to
                    # set common options.  This catches unhandled cases but is
                    # safely wrapped to ignore unsupported options.
                    try:
                        child.configure(bg=surface_bg, fg=light_fg)
                    except Exception:
                        pass
                # Recurse regardless of widget type
                style_widget(child)

        # Kick off the recursive styling on the root window.  Wrap in a
        # try/except to suppress any unexpected Tk errors while still
        # applying as much of the dark theme as possible.
        try:
            style_widget(self)
        except Exception:
            pass

    # ----------------------------------------------------------------------
    # Configuration and snippet persistence
    #
    def load_config(self) -> None:
        """Load user configuration from disk.

        This method populates the ``self.config`` dictionary from a JSON
        file on disk.  If the file does not exist or is invalid, the
        dictionary remains empty.  Configuration keys include saved
        values for speed settings, pause settings, delay start,
        toggle states, and whether the window should remain always on
        top.  The config file lives in the user's home directory and
        is named ``.auto_typer_config.json``.
        """
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                self.config = json.load(f)
        except Exception:
            self.config = {}

    def save_config(self) -> None:
        """Persist the current configuration to disk.

        Called when the application is closing or when the user
        initiates typing.  The configuration captures the current
        values of entry fields and toggle states, so they are restored
        the next time the app runs.
        """
        data = {
            "min_wpm": self.min_speed_entry.get(),
            "max_wpm": self.max_speed_entry.get(),
            "min_pause_words": self.min_pause_words_entry.get(),
            "max_pause_words": self.max_pause_words_entry.get(),
            "pause_length": self.pause_length_entry.get(),
            "delay_start": self.delay_start_entry.get(),
            "global_hotkey": self.hotkey_entry.get(),
            "random_delays": self.random_delays_var.get(),
            "natural_pauses": self.natural_pauses_var.get(),
            "keep_line_breaks": self.keep_line_breaks_var.get(),
            "multi_field": self.multi_field_var.get(),
            "typo_sim": self.typo_sim_var.get(),
            "invisible_paste": self.invisible_paste_var.get(),
            "keyboard_sounds": self.keyboard_sounds_var.get(),
            "preserve_formatting": self.preserve_formatting_var.get(),
            "alarm": self.alarm_var.get(),
            "always_on_top": self.always_on_top_var.get(),
            "last_text": self.text_box.get("1.0", "end-1c"),
            "last_snippet": self.snippet_var.get() if hasattr(self, "snippet_var") else "",
            # Persist finish‑in duration, finish‑by time and countdown timer state
            "finish_in": self.finish_time_entry.get(),
            "finish_time": self.finish_time_entry.get(),  # for backward compatibility
            "finish_by": self.finish_by_entry.get() if hasattr(self, "finish_by_entry") else "",
            "countdown_timer": self.countdown_timer_var.get(),
        }
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def load_snippets(self) -> None:
        """Load the snippet library from disk.

        The snippet file stores a list of objects with ``name``, ``text``
        and optional ``hotkey`` keys.  If the file does not exist,
        ``self.snippets`` remains empty.  Invalid JSON is ignored.
        """
        try:
            with open(self.snippets_path, "r", encoding="utf-8") as f:
                self.snippets = json.load(f)
            if not isinstance(self.snippets, list):
                self.snippets = []
        except Exception:
            self.snippets = []

    def save_snippets(self) -> None:
        """Persist the current snippet library to disk.

        Called whenever snippets are added, updated or deleted.  This
        writes the ``self.snippets`` list to the snippets file.
        """
        try:
            with open(self.snippets_path, "w", encoding="utf-8") as f:
                json.dump(self.snippets, f, indent=2)
        except Exception:
            pass

    def restore_state(self) -> None:
        """Restore saved state into UI controls.

        After constructing the interface, this method populates the
        widgets with values from the loaded configuration.  It sets
        toggle states, text contents, and selects the last used
        snippet if available.  It also toggles the window's
        always-on-top property.
        """
        cfg = self.config
        # Set toggles
        self.random_delays_var.set(cfg.get("random_delays", True))
        self.natural_pauses_var.set(cfg.get("natural_pauses", True))
        self.keep_line_breaks_var.set(cfg.get("keep_line_breaks", False))
        self.multi_field_var.set(cfg.get("multi_field", False))
        self.typo_sim_var.set(cfg.get("typo_sim", False))
        self.invisible_paste_var.set(cfg.get("invisible_paste", False))
        self.keyboard_sounds_var.set(cfg.get("keyboard_sounds", False))
        self.preserve_formatting_var.set(cfg.get("preserve_formatting", False))
        self.alarm_var.set(cfg.get("alarm", False))
        self.always_on_top_var.set(cfg.get("always_on_top", False))

        # Restore text content
        last_text = cfg.get("last_text", "")
        if last_text:
            self.text_box.delete("1.0", "end")
            self.text_box.insert("1.0", last_text)

        # Restore snippet selection
        last_snippet = cfg.get("last_snippet", "")
        if last_snippet and last_snippet in [s["name"] for s in self.snippets]:
            self.snippet_var.set(last_snippet)

        # Make window topmost if requested
        self.toggle_always_on_top()

        # Restore finish‑in and finish‑by settings and countdown timer
        try:
            finish_in_val = cfg.get("finish_in", cfg.get("finish_time", ""))
            finish_by_val = cfg.get("finish_by", "")
            if hasattr(self, "finish_time_entry"):
                self.finish_time_entry.delete(0, tk.END)
                if finish_in_val:
                    self.finish_time_entry.insert(0, str(finish_in_val))
            if hasattr(self, "finish_by_entry"):
                self.finish_by_entry.delete(0, tk.END)
                if finish_by_val:
                    self.finish_by_entry.insert(0, str(finish_by_val))
            self.countdown_timer_var.set(cfg.get("countdown_timer", False))
        except Exception:
            pass

    # ----------------------------------------------------------------------
    # Always-on-top handling
    def toggle_always_on_top(self) -> None:
        """Set or unset the window's topmost attribute based on the toggle."""
        try:
            self.attributes("-topmost", self.always_on_top_var.get())
        except Exception:
            pass

    # ----------------------------------------------------------------------
    # Snippet management UI
    def build_snippets_tab(self, frame: tk.Frame) -> None:
        """Construct the Snippets tab UI.

        This tab allows the user to create, edit, delete and manage
        named snippets.  Each snippet can have an optional hotkey
        assigned.  The UI consists of a listbox for snippet names,
        fields for snippet content and hotkey, and buttons for
        operations.  Import/export buttons allow saving and loading
        the entire snippet library to and from external files.
        """
        # Left side: list of snippet names
        list_frame = tk.Frame(frame)
        list_frame.pack(side="left", fill="y", padx=(10, 5), pady=10)
        tk.Label(list_frame, text="Saved snippets:").pack(anchor="w")
        self.snippet_listbox = tk.Listbox(list_frame, height=12)
        self.snippet_listbox.pack(fill="y", expand=True)
        self.snippet_listbox.bind("<<ListboxSelect>>", self.on_snippet_list_select)
        self.populate_snippet_list()

        # Right side: editor and controls
        editor_frame = tk.Frame(frame)
        editor_frame.pack(side="right", fill="both", expand=True, padx=(5, 10), pady=10)
        tk.Label(editor_frame, text="Snippet name:").grid(row=0, column=0, sticky="w")
        self.snippet_name_entry = tk.Entry(editor_frame)
        self.snippet_name_entry.grid(row=0, column=1, sticky="ew")
        tk.Label(editor_frame, text="Snippet hotkey:").grid(row=1, column=0, sticky="w")
        self.snippet_hotkey_entry = tk.Entry(editor_frame)
        self.snippet_hotkey_entry.grid(row=1, column=1, sticky="ew")
        tk.Label(editor_frame, text="Snippet content:").grid(row=2, column=0, sticky="nw")
        self.snippet_text_box = tk.Text(editor_frame, width=40, height=8)
        self.snippet_text_box.grid(row=2, column=1, sticky="nsew")
        # Buttons for snippet actions
        button_frame = tk.Frame(editor_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(5, 0), sticky="ew")
        add_btn = ttk.Button(button_frame, text="Add New", command=self.add_snippet)
        add_btn.pack(side="left", padx=2)
        save_btn = ttk.Button(button_frame, text="Save", command=self.save_snippet_action)
        save_btn.pack(side="left", padx=2)
        del_btn = ttk.Button(button_frame, text="Delete", command=self.delete_snippet)
        del_btn.pack(side="left", padx=2)
        # Import/export buttons
        import_btn = ttk.Button(button_frame, text="Import", command=self.import_snippets)
        import_btn.pack(side="left", padx=2)
        export_btn = ttk.Button(button_frame, text="Export", command=self.export_snippets)
        export_btn.pack(side="left", padx=2)
        # Configure editor frame grid weights
        editor_frame.columnconfigure(1, weight=1)
        editor_frame.rowconfigure(2, weight=1)

    # ------------------------------------------------------------------
    # Help tab construction and interaction
    def build_help_tab(self, frame: tk.Frame) -> None:
        """Build the help tab with collapsible sections and scrolling.

        The help tab presents a list of application features.  Each feature
        is represented by a header button; clicking the button toggles
        visibility of a descriptive body below it.  A vertical scrollbar
        allows the user to scroll through the list when the content
        exceeds the available height.  The dark theme and neumorphic
        styling are applied automatically via ``apply_dark_mode``.
        """
        # List of features and their descriptions.  Each tuple
        # contains a header title and the corresponding description.
        features = [
            (
                "Snippets",
                "Use the Snippets tab to create and manage named text snippets.\n"
                "Each snippet can have its own hotkey and will appear in the drop‑down on the Type tab.\n"
                "You can import/export snippets to share them between machines."
            ),
            (
                "Dynamic placeholders",
                "Include variables in your text using braces (e.g. {name}, {date}).\n"
                "When you start typing, the app will prompt you to fill in these values."
            ),
            (
                "Speed settings (WPM)",
                "Set the range of typing speed in words per minute.  A word is defined as five characters【524493561122292†L239-L244】.\n"
                "When 'Randomised delays' is on, a new speed is chosen for every character;\n"
                "when off, a fixed speed equal to the average of your range is used."
            ),
            (
                "Natural pauses",
                "Specify a range of words and a pause length (ms).\n"
                "The program inserts a pause after a random number of words within your range."
            ),
            (
                "Finish in time & countdown",
                "Enter a number of seconds in the 'Finish in (sec)' field to complete typing within that duration.\n"
                "The app disables randomised delays and pauses to achieve a consistent typing speed.\n"
                "Enable 'Countdown timer' to show the remaining time."
            ),
            (
                "Delay start",
                "Wait a specified number of seconds before typing begins."
            ),
            (
                "Line break handling",
                "Choose how newline characters are handled.  'Keep line breaks' preserves newlines,\n"
                "'Multi‑field' converts newlines to tabs, 'Preserve formatting' types the text exactly as pasted."
            ),
            (
                "Typo simulation",
                "Occasionally types a wrong character and backspaces to correct it."
            ),
            (
                "Invisible paste",
                "Copies the entire text to the clipboard and pastes it with Ctrl+V;\n"
                "this bypasses per‑character delays."
            ),
            (
                "Keyboard sounds",
                "Plays a beep for each keypress (Windows only)."
            ),
            (
                "Alarm when done",
                "Plays a short sound when typing completes."
            ),
            (
                "Always on top",
                "Keeps the app window above other windows for quick access."
            ),
            (
                "Progress bar & countdown",
                "Shows how much of the text has been typed and an estimated remaining time based on your settings."
            ),
            (
                "Global hotkey",
                "Set a hotkey to start/stop typing.  Snippets can also have their own hotkeys."
            ),
            (
                "How to start",
                "Click 'Type into focused app' or press your hotkey, then click into the destination application and field.\n"
                "The app waits briefly (plus any delay or countdown you specify) before typing begins.\n"
                "Avoid using the keyboard or mouse until typing finishes.  This tool is intended for personal convenience—\n"
                "do not use it to automate sensitive login forms or violate terms of service."
            ),
        ]

        # Create a canvas and a vertical scrollbar for scrolling the help content
        canvas = tk.Canvas(frame, borderwidth=0, highlightthickness=0)
        v_scroll = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Frame inside the canvas that will contain the help items
        content_frame = tk.Frame(canvas)
        # Make the window item in the canvas and anchor it to the top-left
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # Store references for updating scroll region and toggling sections
        self.help_canvas = canvas
        self.help_content_frame = content_frame
        self.help_section_bodies: list[tk.Widget] = []

        # Build each help section with a header button and a descriptive body
        for idx, (title, description) in enumerate(features):
            # Header uses the rounded button style for consistency
            header_btn = ttk.Button(content_frame, text=title, style="Rounded.TButton")
            header_btn.grid(row=idx * 2, column=0, sticky="ew", padx=10, pady=(10 if idx == 0 else 0, 2))
            # Body is a label that wraps text; initially hidden
            body_label = tk.Label(content_frame, text=description, wraplength=660, justify="left", anchor="w")
            body_label.grid(row=idx * 2 + 1, column=0, sticky="ew", padx=20, pady=(0, 10))
            body_label.grid_remove()
            # Bind header click to toggle the body
            header_btn.configure(command=lambda i=idx: self.toggle_help_section(i))
            self.help_section_bodies.append(body_label)
        # Make the content frame expand horizontally
        content_frame.columnconfigure(0, weight=1)

        # Update the scroll region whenever the content size changes
        def on_frame_configure(event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
        content_frame.bind("<Configure>", on_frame_configure)

    def toggle_help_section(self, index: int) -> None:
        """Toggle the visibility of a help section at the given index."""
        try:
            body = self.help_section_bodies[index]
        except Exception:
            return
        # Show or hide the body
        if body.winfo_viewable():
            body.grid_remove()
        else:
            body.grid()
        # After toggling, update the scroll region so the canvas
        # reflects the new size of the content
        if hasattr(self, "help_canvas"):
            try:
                self.help_canvas.configure(scrollregion=self.help_canvas.bbox("all"))
            except Exception:
                pass

    def populate_snippet_list(self) -> None:
        """Refresh the snippet listbox with current snippet names."""
        self.snippet_listbox.delete(0, tk.END)
        for snip in self.snippets:
            self.snippet_listbox.insert(tk.END, snip.get("name", "(unnamed)"))

    def populate_snippet_combobox(self) -> None:
        """Refresh the snippet drop‑down on the main tab."""
        names = [snip.get("name", "") for snip in self.snippets]
        self.snippet_combobox["values"] = names

    def on_snippet_list_select(self, event: tk.Event) -> None:
        """Handle selection of a snippet in the listbox.

        This loads the selected snippet's data into the editor fields.
        """
        if not self.snippet_listbox.curselection():
            return
        index = self.snippet_listbox.curselection()[0]
        snip = self.snippets[index]
        self.snippet_name_entry.delete(0, tk.END)
        self.snippet_name_entry.insert(0, snip.get("name", ""))
        self.snippet_hotkey_entry.delete(0, tk.END)
        self.snippet_hotkey_entry.insert(0, snip.get("hotkey", ""))
        self.snippet_text_box.delete("1.0", tk.END)
        self.snippet_text_box.insert("1.0", snip.get("text", ""))

    def add_snippet(self) -> None:
        """Prepare the editor for a new snippet."""
        self.snippet_listbox.selection_clear(0, tk.END)
        self.snippet_name_entry.delete(0, tk.END)
        self.snippet_hotkey_entry.delete(0, tk.END)
        self.snippet_text_box.delete("1.0", tk.END)

    def save_snippet_action(self) -> None:
        """Save or update the snippet currently in the editor."""
        name = self.snippet_name_entry.get().strip()
        text = self.snippet_text_box.get("1.0", "end-1c")
        hotkey = self.snippet_hotkey_entry.get().strip()
        if not name:
            messagebox.showerror("Name required", "Please enter a name for the snippet.")
            return
        # If snippet exists (matched by name), update it; else append new
        existing = next((s for s in self.snippets if s.get("name") == name), None)
        if existing:
            existing["text"] = text
            existing["hotkey"] = hotkey
        else:
            self.snippets.append({"name": name, "text": text, "hotkey": hotkey})
        # Refresh UI and persist
        self.save_snippets()
        self.populate_snippet_list()
        self.populate_snippet_combobox()
        # Re-register snippet hotkeys
        self.register_snippet_hotkeys()
        messagebox.showinfo("Snippet saved", f"Snippet '{name}' saved successfully.")

    def delete_snippet(self) -> None:
        """Delete the selected snippet from the list."""
        if not self.snippet_listbox.curselection():
            messagebox.showerror("No selection", "Please select a snippet to delete.")
            return
        idx = self.snippet_listbox.curselection()[0]
        snip_name = self.snippets[idx].get("name", "")
        if messagebox.askyesno("Confirm delete", f"Delete snippet '{snip_name}'?"):
            del self.snippets[idx]
            self.save_snippets()
            self.populate_snippet_list()
            self.populate_snippet_combobox()
            self.register_snippet_hotkeys()

    def import_snippets(self) -> None:
        """Import snippets from a JSON file selected by the user."""
        path = filedialog.askopenfilename(
            title="Import snippets", filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                self.snippets = data
                self.save_snippets()
                self.populate_snippet_list()
                self.populate_snippet_combobox()
                self.register_snippet_hotkeys()
                messagebox.showinfo("Import successful", "Snippets imported successfully.")
            else:
                messagebox.showerror("Import failed", "Invalid snippet file format.")
        except Exception as exc:
            messagebox.showerror("Import failed", f"Failed to import snippets:\n{exc}")

    def export_snippets(self) -> None:
        """Export the current snippet library to a JSON file."""
        path = filedialog.asksaveasfilename(
            title="Export snippets", defaultextension=".json", filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.snippets, f, indent=2)
            messagebox.showinfo("Export successful", "Snippets exported successfully.")
        except Exception as exc:
            messagebox.showerror("Export failed", f"Failed to export snippets:\n{exc}")

    # ----------------------------------------------------------------------
    # Snippet usage
    def populate_snippet_listbox_and_combobox(self):
        """Helper to repopulate both the snippet listbox and combobox."""
        self.populate_snippet_list()
        self.populate_snippet_combobox()

    def on_snippet_selected(self) -> None:
        """Load the selected snippet into the main text box."""
        name = self.snippet_var.get()
        for snip in self.snippets:
            if snip.get("name") == name:
                self.text_box.delete("1.0", tk.END)
                self.text_box.insert("1.0", snip.get("text", ""))
                break

    def register_snippet_hotkeys(self) -> None:
        """Register global hotkeys for each snippet that defines one."""
        # If keyboard module is unavailable, skip
        if keyboard is None:
            return
        # Clear existing snippet hotkeys first.  We use the same
        # clearing logic as register_hotkey() to avoid duplicate
        # callbacks.  This will not remove the main start/stop hotkey.
        try:
            if hasattr(keyboard, "clear_all_hotkeys"):
                keyboard.clear_all_hotkeys()
            elif hasattr(keyboard, "unhook_all_hotkeys"):
                keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        # Re-register the main global hotkey
        self.register_hotkey()
        # Now register snippet-specific hotkeys
        for snip in self.snippets:
            hk = snip.get("hotkey")
            name = snip.get("name")
            text = snip.get("text")
            if hk:
                def make_callback(snip=snip):
                    return lambda: self.after(0, lambda: self.start_typing_snippet(snip))
                try:
                    keyboard.add_hotkey(hk, make_callback())
                except Exception:
                    pass

    def start_typing_snippet(self, snip: dict[str, any]) -> None:
        """Start typing using the contents of the given snippet.

        This function constructs a temporary environment for the
        snippet: it resolves placeholders, then calls the same
        underlying typing logic as the main start button.  It
        preserves the current UI settings (speed, pauses, toggles).
        """
        text = snip.get("text", "")
        text = self.resolve_placeholders(text)
        if not text:
            return
        # Populate the text box (optional) and call start_typing
        self.text_box.delete("1.0", tk.END)
        self.text_box.insert("1.0", text)
        # Save snippet name for config
        self.snippet_var.set(snip.get("name", ""))
        self.start_typing()

    # ----------------------------------------------------------------------
    # Placeholder resolution
    def resolve_placeholders(self, text: str) -> str:
        """Prompt the user to fill in any placeholders in the text.

        Placeholders are denoted by curly braces (e.g. ``{name}``).  For
        each unique placeholder, the user is prompted for a value via
        a simple dialog.  All occurrences of the placeholder are then
        replaced with the provided value.  If the user cancels the
        dialog or provides an empty value, the placeholder is left
        unchanged.
        """
        pattern = re.compile(r"\{([^{}]+)\}")
        matches = pattern.findall(text)
        # Use an ordered dict to preserve prompt order but ensure uniqueness
        seen: dict[str, None] = {}
        for m in matches:
            seen[m] = None
        for placeholder in seen.keys():
            # Prompt for value; default to placeholder name
            value = simpledialog.askstring("Placeholder", f"Enter value for '{placeholder}':", parent=self)
            if value is None:
                # Cancelled; keep placeholder as-is
                continue
            # Replace all occurrences
            text = text.replace("{" + placeholder + "}", value)
        return text

    # ----------------------------------------------------------------------
    # Progress bar update
    def update_progress_display(self, chars_typed: int, words_typed: int) -> None:
        """Update the progress bar and status label based on progress.

        This method runs on the Tkinter main thread.  It receives the
        number of characters and words typed so far and computes the
        percentage complete, remaining time and schedule status.  The
        status label displays multiple pieces of information:

        - Words typed and total words (e.g. ``50/200``)
        - Percentage complete (e.g. ``25.0%``)
        - Estimated time remaining (``H:MM:SS`` or ``M:SS``)
        - Target finish time when a deadline is set
        - Whether the current pace is ahead of schedule, on pace or behind

        The countdown timer toggle controls whether the remaining time
        appears.  Even when the countdown is hidden, the status label
        still shows words and percentage complete.
        """
        try:
            # Protect against division by zero
            if getattr(self, "total_chars", 0) > 0:
                progress_pct = (chars_typed / self.total_chars) * 100.0
            else:
                progress_pct = 100.0
            self.progress_var.set(progress_pct)

            now = time.time()
            # Determine remaining time and finish information.  When
            # operating under a target schedule (finish‑in or finish‑by),
            # the remaining time equals the difference between the
            # target finish timestamp and the current time.  Otherwise
            # we estimate remaining time based on the predicted total
            # duration.  We also compute a predicted finish time for
            # display when a schedule is not provided.
            use_target = getattr(self, "use_target_mode", False)
            target_finish = getattr(self, "target_finish_time", None)
            predicted_total = getattr(self, "predicted_total_time", 0.0)
            if use_target and target_finish:
                # Scheduled mode: compute remaining time until the
                # user‑specified target finish.  Format the due time
                # for display (HH:MM).
                remaining_seconds = max(target_finish - now, 0.0)
                try:
                    target_dt = datetime.datetime.fromtimestamp(target_finish)
                    target_time_str = target_dt.strftime("%H:%M")
                except Exception:
                    target_time_str = ""
                # Determine the schedule pacing status.  If the
                # predicted finish (now + remaining) is earlier than
                # the target by more than 10 seconds we are ahead;
                # if later by more than 10 seconds we are behind.
                predicted_finish_time = now + remaining_seconds
                diff = predicted_finish_time - target_finish
                if diff < -10:
                    status_str = "Ahead of schedule"
                elif diff > 10:
                    status_str = "Behind schedule"
                else:
                    status_str = "On pace"
            else:
                # Unscheduled mode: estimate remaining time as the
                # difference between the predicted total typing time
                # and the elapsed time since typing began.  Compute
                # a predicted finish time (absolute) and format it
                # for display.  No schedule pacing status is needed.
                start_time = getattr(self, "start_time", now)
                time_elapsed = now - start_time
                remaining_seconds = max(predicted_total - time_elapsed, 0.0)
                predicted_finish_time = start_time + predicted_total
                try:
                    finish_dt = datetime.datetime.fromtimestamp(predicted_finish_time)
                    target_time_str = finish_dt.strftime("%H:%M") if predicted_total > 0 else ""
                except Exception:
                    target_time_str = ""
                status_str = ""

            # Construct the status string.  Always show words and percentage.
            word_info = f"{words_typed}/{self.total_words}" if getattr(self, "total_words", 0) > 0 else f"{words_typed}"
            pct_info = f"{progress_pct:.1f}%"

            # Format remaining time if the countdown timer is enabled and
            # there is still time left.  Otherwise leave it blank.
            remain_info = ""
            if self.countdown_timer_var.get() and remaining_seconds > 0:
                # Convert seconds to H:MM:SS or M:SS
                rem_int = int(round(remaining_seconds))
                rem_hours, rem_remainder = divmod(rem_int, 3600)
                rem_mins, rem_secs = divmod(rem_remainder, 60)
                if rem_hours > 0:
                    remain_info = f"{rem_hours}:{rem_mins:02d}:{rem_secs:02d}"
                else:
                    remain_info = f"{rem_mins}:{rem_secs:02d}"
            # Build components list
            parts = [f"Words: {word_info}", pct_info]
            if remain_info:
                parts.append(f"Remaining: {remain_info}")
            # Include a finish time indicator if available.  When a
            # schedule is active (finish‑in or finish‑by), prefix with
            # "Finish by"; otherwise prefix with "Finish at" to
            # emphasise that it is an estimate rather than a hard
            # deadline.
            if target_time_str:
                if getattr(self, "use_target_mode", False) and getattr(self, "target_finish_time", None):
                    parts.append(f"Finish by: {target_time_str}")
                else:
                    parts.append(f"Finish at: {target_time_str}")
            if status_str:
                parts.append(status_str)
            label_text = " | ".join(parts)
            self.remaining_time_label.config(text=label_text)
        except Exception:
            # Ignore errors if widgets are destroyed during update
            pass

    # ----------------------------------------------------------------------
    # Override window close handler to save state
    def on_close(self) -> None:
        """Handle application shutdown: persist configuration and snippets."""
        self.save_config()
        self.save_snippets()
        self.destroy()



def main() -> None:
    """Entry point for running the application."""
    app = AutoTyperApp()
    app.mainloop()


if __name__ == "__main__":
    main()