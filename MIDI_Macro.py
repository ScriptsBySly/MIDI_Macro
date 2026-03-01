import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import mido
import threading
import time
import ctypes
import json
import os
import difflib

current_port = None
listening = False
idle_port = None
idle_listening = False
feedback_port = None
feedback_output_name = None
feedback_warning_shown = False
macros = {}
active_macro_dialog = None
macro_list_index_map = {}

KEYEVENTF_KEYUP = 0x0002
user32 = ctypes.windll.user32

SPECIAL_VK_MAP = {
    "Return": 0x0D,
    "Tab": 0x09,
    "space": 0x20,
    "Escape": 0x1B,
    "BackSpace": 0x08,
    "Delete": 0x2E,
    "Insert": 0x2D,
    "Home": 0x24,
    "End": 0x23,
    "Prior": 0x21,
    "Next": 0x22,
    "Left": 0x25,
    "Up": 0x26,
    "Right": 0x27,
    "Down": 0x28,
    "Shift_L": 0xA0,
    "Shift_R": 0xA1,
    "Control_L": 0xA2,
    "Control_R": 0xA3,
    "Alt_L": 0xA4,
    "Alt_R": 0xA5,
}

def get_midi_devices():
    return mido.get_input_names()


def get_midi_output_devices():
    return mido.get_output_names()


def get_default_documents_dir():
    return os.path.join(os.path.expanduser("~"), "Documents")


def log_to_output(message):
    if "output_box" not in globals():
        return
    output_box.insert(tk.END, f"{message}\n")
    output_box.see(tk.END)


def normalize_port_name(name):
    lowered = name.lower()
    for token in ("midiin", "midiout", "input", "output", " in ", " out ", "port"):
        lowered = lowered.replace(token, " ")
    cleaned = []
    for char in lowered:
        if char.isalnum() or char.isspace():
            cleaned.append(char)
    normalized = " ".join("".join(cleaned).split())
    return normalized


def find_matching_output_port(input_name):
    outputs = get_midi_output_devices()
    if not outputs:
        return None

    if len(outputs) == 1:
        return outputs[0]

    for output_name in outputs:
        if output_name == input_name:
            return output_name

    input_normalized = normalize_port_name(input_name)
    for output_name in outputs:
        output_normalized = normalize_port_name(output_name)
        if output_normalized == input_normalized:
            return output_name

    for output_name in outputs:
        output_normalized = normalize_port_name(output_name)
        if input_normalized and (input_normalized in output_normalized or output_normalized in input_normalized):
            return output_name

    # Fuzzy fallback for devices that expose input/output with slightly different names.
    best_name = None
    best_score = 0.0
    for output_name in outputs:
        score = difflib.SequenceMatcher(None, input_normalized, normalize_port_name(output_name)).ratio()
        if score > best_score:
            best_score = score
            best_name = output_name
    if best_name and best_score >= 0.45:
        return best_name

    return None


def close_feedback_port():
    global feedback_port, feedback_output_name
    if feedback_port:
        feedback_port.close()
        feedback_port = None
    feedback_output_name = None


def ensure_feedback_port():
    global feedback_port, feedback_output_name, feedback_warning_shown

    input_name = dropdown.get()
    if not input_name or input_name == "No MIDI devices found":
        close_feedback_port()
        return False

    if feedback_port and feedback_output_name:
        if feedback_output_name in get_midi_output_devices():
            return True
        close_feedback_port()

    output_name = find_matching_output_port(input_name)
    if not output_name:
        if not feedback_warning_shown:
            outputs = get_midi_output_devices()
            if outputs:
                log_to_output(f"No matching MIDI output device found for LED feedback. Outputs: {outputs}")
            else:
                log_to_output("No MIDI output devices found for LED feedback.")
            feedback_warning_shown = True
        return False

    try:
        feedback_port = mido.open_output(output_name)
        feedback_output_name = output_name
        feedback_warning_shown = False
        log_to_output(f"LED feedback output connected: {output_name}")
        return True
    except Exception as e:
        close_feedback_port()
        log_to_output(f"Error opening MIDI output for LED feedback: {e}")
        return False


def parse_binding_key(binding_key):
    try:
        binding_type, raw_value = binding_key.split(":", 1)
        binding_value = int(raw_value)
        return binding_type, binding_value
    except Exception:
        return None, None


def set_led_for_binding(binding_key, turn_on=True):
    if not ensure_feedback_port():
        return

    binding_type, binding_value = parse_binding_key(binding_key)
    if binding_type is None:
        return

    try:
        if binding_type == "note":
            velocity = 127 if turn_on else 0
            message = mido.Message("note_on", note=binding_value, velocity=velocity)
        elif binding_type == "cc":
            value = 127 if turn_on else 0
            message = mido.Message("control_change", control=binding_value, value=value)
        else:
            return
        feedback_port.send(message)
    except Exception as e:
        log_to_output(f"Error sending LED feedback: {e}")


def turn_off_all_leds():
    if not ensure_feedback_port():
        return

    try:
        for note in range(128):
            feedback_port.send(mido.Message("note_on", note=note, velocity=0))
        for control in range(128):
            feedback_port.send(mido.Message("control_change", control=control, value=0))
    except Exception as e:
        log_to_output(f"Error turning off LEDs: {e}")


def sync_macro_leds():
    if not macros:
        return
    for binding_key in macros.keys():
        set_led_for_binding(binding_key, turn_on=True)


def keysym_to_vk(keysym):
    if keysym in SPECIAL_VK_MAP:
        return SPECIAL_VK_MAP[keysym]

    if keysym.startswith("F") and keysym[1:].isdigit():
        fn_number = int(keysym[1:])
        if 1 <= fn_number <= 24:
            return 0x70 + (fn_number - 1)

    if len(keysym) == 1:
        vk_scan = user32.VkKeyScanW(ord(keysym))
        if vk_scan != -1:
            return vk_scan & 0xFF
        return None

    return None


def run_macro_keys(keys):
    vk_codes = []
    for key in keys:
        vk_code = keysym_to_vk(key)
        if vk_code is None:
            output_box.insert(tk.END, f"Cannot type unsupported key: {key}\n")
            output_box.see(tk.END)
            return
        vk_codes.append(vk_code)

    for vk_code in vk_codes:
        user32.keybd_event(vk_code, 0, 0, 0)

    time.sleep(0.03)

    for vk_code in reversed(vk_codes):
        user32.keybd_event(vk_code, 0, KEYEVENTF_KEYUP, 0)


def trigger_macro(binding_key):
    keys = macros.get(binding_key)
    if not keys:
        return
    threading.Thread(target=run_macro_keys, args=(keys,), daemon=True).start()


def refresh_macro_list():
    global macro_list_index_map

    if "macro_listbox" not in globals():
        return

    macro_list_index_map = {}
    macro_listbox.delete(0, tk.END)
    for index, binding_key in enumerate(sorted(macros.keys())):
        keys = macros[binding_key]
        keys_text = " + ".join(keys) if keys else "(no keys)"
        macro_listbox.insert(tk.END, f"{binding_key} -> {keys_text}")
        macro_list_index_map[binding_key] = index


def delete_selected_macro():
    if "macro_listbox" not in globals():
        return

    selection = macro_listbox.curselection()
    if not selection:
        log_to_output("Select a macro first, then click Delete Macro.")
        return

    selected_text = macro_listbox.get(selection[0])
    binding_key = selected_text.split(" -> ", 1)[0]
    if binding_key not in macros:
        return

    del macros[binding_key]
    set_led_for_binding(binding_key, turn_on=False)
    refresh_macro_list()
    log_to_output(f"Deleted macro: {binding_key}")


def highlight_macro_entry(binding_key):
    if "macro_listbox" not in globals():
        return

    if tabs.select() != str(macros_tab):
        return

    focused_widget = root.focus_displayof()
    if focused_widget is None or focused_widget.winfo_toplevel() != root:
        return

    index = macro_list_index_map.get(binding_key)
    if index is None:
        return

    macro_listbox.selection_clear(0, tk.END)
    macro_listbox.selection_set(index)
    macro_listbox.activate(index)
    macro_listbox.see(index)


def process_macro_binding(message):
    global active_macro_dialog

    if active_macro_dialog and active_macro_dialog.midi_binding_key is None:
        if active_macro_dialog.handle_midi_message(message):
            return True

    binding_key, _ = get_midi_binding(message)
    if binding_key in macros:
        highlight_macro_entry(binding_key)
        trigger_macro(binding_key)

    return False


def idle_listen_to_midi():
    global idle_listening
    while idle_listening and not listening:
        for message in idle_port.iter_pending():
            process_macro_binding(message)
        time.sleep(0.01)


def stop_idle_macro_listener():
    global idle_listening, idle_port

    idle_listening = False
    if idle_port:
        idle_port.close()
        idle_port = None


def start_idle_macro_listener():
    global idle_port, idle_listening

    if listening or idle_listening:
        return True

    device_name = dropdown.get()
    if not device_name:
        return False

    try:
        idle_port = mido.open_input(device_name)
        idle_listening = True
        threading.Thread(target=idle_listen_to_midi, daemon=True).start()
        return True
    except Exception:
        idle_listening = False
        idle_port = None
        return False


def save_macro_mapping():
    file_path = filedialog.asksaveasfilename(
        title="Save Macro Mapping",
        initialdir=get_default_documents_dir(),
        defaultextension=".json",
        filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
    )
    if not file_path:
        return

    try:
        with open(file_path, "w", encoding="utf-8") as file:
            json.dump(macros, file, indent=2)
        output_box.insert(tk.END, f"Macro mapping saved: {file_path}\n")
        output_box.see(tk.END)
    except Exception as e:
        output_box.insert(tk.END, f"Error saving macro mapping: {e}\n")
        output_box.see(tk.END)


def load_macro_mapping():
    file_path = filedialog.askopenfilename(
        title="Load Macro Mapping",
        initialdir=get_default_documents_dir(),
        filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
    )
    if not file_path:
        return

    try:
        with open(file_path, "r", encoding="utf-8") as file:
            data = json.load(file)

        if not isinstance(data, dict):
            raise ValueError("Invalid format: expected an object/dictionary.")

        loaded_macros = {}
        for binding_key, keys in data.items():
            if not isinstance(binding_key, str):
                continue
            if not isinstance(keys, list):
                continue
            valid_keys = [key for key in keys if isinstance(key, str)]
            loaded_macros[binding_key] = valid_keys

        turn_off_all_leds()
        macros.clear()
        macros.update(loaded_macros)
        refresh_macro_list()
        sync_macro_leds()
        output_box.insert(tk.END, f"Macro mapping loaded: {file_path}\n")
        output_box.see(tk.END)
    except Exception as e:
        output_box.insert(tk.END, f"Error loading macro mapping: {e}\n")
        output_box.see(tk.END)

def refresh_devices():
    devices = get_midi_devices()
    dropdown['values'] = devices
    if devices:
        dropdown.current(0)
    else:
        dropdown.set("No MIDI devices found")

    close_feedback_port()
    sync_macro_leds()

    if not listening:
        stop_idle_macro_listener()
        start_idle_macro_listener()


def on_device_selected(_event=None):
    close_feedback_port()
    sync_macro_leds()

    if listening:
        return
    stop_idle_macro_listener()
    start_idle_macro_listener()

def start_listening():
    global current_port, listening

    if listening:
        return

    device_name = dropdown.get()
    if not device_name:
        return

    stop_idle_macro_listener()

    try:
        current_port = mido.open_input(device_name)
        listening = True
        status_label.config(text="Running...",  fg="green")
        start_button.config(state="disabled")
        stop_button.config(state="normal")
        threading.Thread(target=listen_to_midi, daemon=True).start()
    except Exception as e:
        start_idle_macro_listener()
        start_button.config(state="normal")
        stop_button.config(state="disabled")
        output_box.insert(tk.END, f"Error opening device: {e}\n")

def stop_listening():
    global listening, current_port
    listening = False
    if current_port:
        current_port.close()
        current_port = None
    status_label.config(text="Stopped",  fg="red")
    start_button.config(state="normal")
    stop_button.config(state="disabled")
    start_idle_macro_listener()
    sync_macro_leds()

def listen_to_midi():
    global listening
    while listening:
        for message in current_port.iter_pending():
            handle_message(message)


def get_midi_binding(message):
    if message.type == 'note_on' and message.velocity > 0:
        return f"note:{message.note}", f"Note {message.note}"
    if message.type == 'control_change' and message.value > 0:
        return f"cc:{message.control}", f"Control {message.control}"
    return None, None


class MacroDialog:
    def __init__(self, master):
        self.window = tk.Toplevel(master)
        self.window.title("Add Macro")
        self.window.geometry("500x380")
        self.window.resizable(False, False)
        self.window.protocol("WM_DELETE_WINDOW", self.cancel)

        self.running = True
        self.capture_key_mode = False
        self.midi_binding_key = None
        self.keys = []

        self.prompt_box = tk.Text(self.window, height=2, width=58, borderwidth=0)
        self.prompt_box.tag_configure("bold", font=("TkDefaultFont", 10, "bold"))
        self.prompt_box.insert(tk.END, "Press a button from your MIDI device to register it.", "bold")
        self.prompt_box.configure(state="disabled")
        self.prompt_box.pack(pady=10)

        self.midi_label = tk.Label(self.window, text="Registered MIDI button: None")
        self.midi_label.pack(pady=5)

        self.add_key_button = tk.Button(self.window, text="Add Key", state="disabled", command=self.start_key_capture)
        self.add_key_button.pack(pady=5)

        self.key_status_label = tk.Label(self.window, text="No key capture in progress.")
        self.key_status_label.pack(pady=5)

        self.key_limit_label = tk.Label(self.window, text="You can add up to 3 keys per macro.")
        self.key_limit_label.pack(pady=2)

        self.keys_listbox = tk.Listbox(self.window, height=5, width=40)
        self.keys_listbox.pack(pady=5)

        self.clear_macro_button = tk.Button(self.window, text="Clear Macro", command=self.clear_macro_keys)
        self.clear_macro_button.pack(pady=5)

        actions = tk.Frame(self.window)
        actions.pack(pady=10)

        self.save_button = tk.Button(actions, text="Save", state="disabled", command=self.save)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.reset_button = tk.Button(actions, text="Reset", command=self.reset)
        self.reset_button.pack(side=tk.LEFT, padx=5)

        self.cancel_button = tk.Button(actions, text="Cancel", command=self.cancel)
        self.cancel_button.pack(side=tk.LEFT, padx=5)

        self.window.bind("<KeyPress>", self.on_key_press)
        self.start_midi_capture()

    def start_midi_capture(self):
        global active_macro_dialog

        device_name = dropdown.get()
        if not device_name:
            active_macro_dialog = None
            self.update_prompt("No MIDI input selected. Select a device and try again.")
            return

        if not listening and not start_idle_macro_listener():
            active_macro_dialog = None
            self.update_prompt("Unable to open MIDI input. Check device and try again.")
            return

        active_macro_dialog = self

    def handle_midi_message(self, message):
        binding_key, binding_text = get_midi_binding(message)
        if binding_key:
            self.register_midi_button(binding_key, binding_text)
            return True
        return False

    def register_midi_button(self, binding_key, binding_text):
        if self.midi_binding_key is not None:
            return

        self.midi_binding_key = binding_key
        self.midi_label.config(text=f"Registered MIDI button: {binding_text}")
        if binding_key in macros:
            self.keys = list(macros[binding_key])
            self.refresh_keys()
            self.key_status_label.config(text="Loaded existing macro keys.")
            self.update_prompt("Existing macro loaded. Add/remove keys, then Save.")
        else:
            self.update_prompt("MIDI button registered. Click 'Add Key' and press a keyboard key. \nPress 'Save' when done!")
        self.update_buttons()

    def start_key_capture(self):
        if len(self.keys) >= 3:
            self.key_status_label.config(text="Key limit reached (3). Clear or save to continue.")
            return
        self.capture_key_mode = True
        self.key_status_label.config(text="Press a keyboard key now...")
        self.window.focus_force()

    def on_key_press(self, event):
        if not self.capture_key_mode:
            return

        if len(self.keys) < 3:
            self.keys.append(event.keysym)
            self.refresh_keys()

        self.capture_key_mode = False
        self.key_status_label.config(text="Key added.")
        self.update_buttons()

    def refresh_keys(self):
        self.keys_listbox.delete(0, tk.END)
        for key in self.keys:
            self.keys_listbox.insert(tk.END, key)

    def clear_macro_keys(self):
        self.keys = []
        self.refresh_keys()
        self.key_status_label.config(text="Macro keys cleared.")
        self.update_buttons()

    def update_prompt(self, message):
        self.prompt_box.configure(state="normal")
        self.prompt_box.delete("1.0", tk.END)
        self.prompt_box.insert(tk.END, message, "bold")
        self.prompt_box.configure(state="disabled")

    def update_buttons(self):
        can_add_key = self.midi_binding_key is not None and len(self.keys) < 3
        self.add_key_button.config(state="normal" if can_add_key else "disabled")

        can_save = self.midi_binding_key is not None and len(self.keys) > 0
        self.save_button.config(state="normal" if can_save else "disabled")

    def save(self):
        macros[self.midi_binding_key] = list(self.keys)
        refresh_macro_list()
        set_led_for_binding(self.midi_binding_key, turn_on=True)
        output_box.insert(tk.END, f"Saved macro: {self.midi_binding_key} -> {self.keys}\n")
        output_box.see(tk.END)
        self.cancel()

    def reset(self):
        self.capture_key_mode = False
        self.midi_binding_key = None
        self.keys = []
        self.refresh_keys()
        self.midi_label.config(text="Registered MIDI button: None")
        self.key_status_label.config(text="No key capture in progress.")
        self.update_prompt("Press a button from your MIDI device to register it.")
        self.update_buttons()
        self.start_midi_capture()

    def cancel(self):
        global active_macro_dialog
        self.running = False
        if active_macro_dialog is self:
            active_macro_dialog = None
        self.window.destroy()


def open_add_macro_dialog():
    MacroDialog(root)

def handle_message(message):
    if process_macro_binding(message):
        return

    text = None

    if message.type == 'note_on' and message.velocity > 0:
        text = f"Button-{message.note} Pressed!\n"
    elif message.type == 'control_change' and message.value > 0:
        text = f"CButton-{message.control} Pressed!\n"

    if text:
        output_box.insert(tk.END, text)
        output_box.see(tk.END)

# ---- GUI ----
root = tk.Tk()
root.title("MIDI Device Monitor")
root.geometry("600x400")
root.resizable(False, False)

style = ttk.Style(root)
style.theme_use("clam")
style.configure("Device.TCombobox", fieldbackground="white", background="white")
style.map("Device.TCombobox", fieldbackground=[("readonly", "white")], background=[("readonly", "white")])
style.configure("Custom.TNotebook", borderwidth=2)
style.configure(
    "Custom.TNotebook.Tab",
    borderwidth=2,
    background="#d0d0d0",
    foreground="#111111"
)
style.map(
    "Custom.TNotebook.Tab",
    background=[("selected", "#ffffff"), ("active", "#e2e2e2")],
    foreground=[("selected", "#000000"), ("active", "#000000")],
    expand=[("selected", (0, 0, 0, 0)), ("!selected", (0, 0, 0, 0))]
)

menu_bar = tk.Menu(root)
menu_bar.add_command(label="Save Mapping", command=save_macro_mapping)
menu_bar.add_command(label="Load Mapping", command=load_macro_mapping)
root.config(menu=menu_bar)

tk.Label(root, text="Select MIDI Input Device:").pack(pady=5)

device_row = tk.Frame(root)
device_row.pack(pady=5)

refresh_button = tk.Button(device_row, text="Refresh Devices", command=refresh_devices)
refresh_button.pack(side=tk.LEFT, padx=5)

dropdown = ttk.Combobox(device_row, state="readonly", width=52, style="Device.TCombobox")
dropdown.pack(side=tk.LEFT, padx=5)
dropdown.bind("<<ComboboxSelected>>", on_device_selected)

tabs_frame = tk.Frame(root, bd=3)
tabs_frame.pack(fill="both", expand=True, padx=8, pady=8)

tabs = ttk.Notebook(tabs_frame, style="Custom.TNotebook")
tabs.pack(fill="both", expand=True, padx=4, pady=4)

main_tab = tk.Frame(tabs)
macros_tab = tk.Frame(tabs)

tabs.add(main_tab, text="Main")
tabs.add(macros_tab, text="Macros")

control_buttons = tk.Frame(main_tab)
control_buttons.pack(pady=5)

start_button = tk.Button(control_buttons, text="Start", command=start_listening)
start_button.pack(side=tk.LEFT, padx=5)

stop_button = tk.Button(control_buttons, text="Stop", command=stop_listening)
stop_button.pack(side=tk.LEFT, padx=5)
stop_button.config(state="disabled")

macro_actions = tk.Frame(macros_tab)
macro_actions.pack(pady=5)

add_macro_button = tk.Button(macro_actions, text="Add Macro", command=open_add_macro_dialog)
add_macro_button.pack(side=tk.LEFT, padx=5)

delete_macro_button = tk.Button(macro_actions, text="Delete Macro", command=delete_selected_macro)
delete_macro_button.pack(side=tk.LEFT, padx=5)

tk.Label(macros_tab, text="Configured Macros:").pack(pady=(8, 4))
macro_listbox = tk.Listbox(macros_tab, height=12, width=65)
macro_listbox.pack(pady=5, padx=8, fill="both", expand=True)

status_label = tk.Label(main_tab, text="Stopped", fg="red")
status_label.pack(pady=5)

# Text box to display MIDI messages
output_box = tk.Text(main_tab, height=12, width=70)
output_box.pack(pady=10)

refresh_devices()
refresh_macro_list()
start_idle_macro_listener()

root.mainloop()
