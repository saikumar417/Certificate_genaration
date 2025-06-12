import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageTk, ImageFont, ImageDraw
import subprocess
import json
import os
import sys
import main2

from data_saparate import main as data_separate_main
from generate_cert_no import main as cert_no_main
from remove_duplicates import main as remove_duplicates_main
from email_module import main as email_main

# Constants
LOGO_DEFAULTS_FILE = "logo_defaults.json"
KNOWLEDGE_FILE = "knowledge_text.json"
PASSION_FILE = "passion_text.json"
SEMINAR_TITLE = "seminar_title.json"

def load_json_data(filename, default):
    if os.path.exists(filename):
        with open(filename, "r") as f:
            return json.load(f)
    return default

def save_json_data(filename, data):
    with open(filename, "w") as f:
        json.dump(data, f)

class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Generation SAA")
        self.root.geometry("1400x800")
        self.root.configure(bg="#f0f0f0")
            # Header Frame
        header_frame = tk.Frame(self.root, bg="#FFFFFF", height=100)
        header_frame.pack(side="top", fill="x")

        # Optional: College Logo
        try:
            logo_img = Image.open("input_files/sitams_logo.png").resize((80, 80))
            logo_img2 = Image.open("input_files/SAA_logo.png").resize((80,80))

            logo_photo = ImageTk.PhotoImage(logo_img)
            logo2_photo = ImageTk.PhotoImage(logo_img2)

            logo_label = tk.Label(header_frame, image=logo_photo, bg="#FFFFFF")
            logo2_label = tk.Label(header_frame, image=logo2_photo, bg="#FFFFFF")

            logo_label.image = logo_photo  # Keep a reference
            logo2_label.image = logo2_photo

            logo_label.pack(side="left", padx=20, pady=10)
            logo2_label.pack(side="right",padx=50,pady=10)
        except:
            pass  # Skip if logo image not found

        # College Name and Tagline
        text_frame = tk.Frame(header_frame, bg="#B9B6C9")
        text_frame.pack(side="left", expand=True, fill="both")

        college_name_label = tk.Label(
            text_frame,
            text="SITAMS ALUMNI ASSOCIATION(SAA)",
            font=("Arial", 24, "bold"),
            fg="#B91BC9",
            bg="#B9B6C9"
        )
        college_name_label.pack(anchor="center", pady=(10, 0))

        tagline_label = tk.Label(
            text_frame,
            text="(Reg. No. 329 / 2019)",
            font=("Arial", 12),
            fg="#0B67AB",
            bg="#B9B6C9"
        )

        tagline_label.pack(anchor="center", pady=(5, 10))

        district_label = tk.Label(
            text_frame,
            text="CHITTOOR - 517127",
            font=("Arial", 12),
            fg="#0B67AB",
            bg="#B9B6C9"
        )
        district_label.pack(anchor="center",pady=(4,11))

        self.knowledge_options = load_json_data(KNOWLEDGE_FILE, [
            "Student Knowledge Sharing Series",
            "An Alumni Knowledge Sharing Series"
        ])
        self.passion_options = load_json_data(PASSION_FILE, [
            "Student Sparks: Fostering Intellectual Exchange",
            "Passionate Towards Passion - An Approach To Career Path"
        ])

        self.seminar_title = load_json_data(SEMINAR_TITLE, [])
        self.create_layout()
        
        # Set icon
        # Absolute path approach (more reliable)
        icon_path = os.path.abspath("SAA_logo.ico")
       # icon_path = os.path.abspath(os.path.join("assets", "SAA_logo.ico"))  # Update path as needed
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(icon_path)
            except Exception as e:
                print("Failed to set icon:", e)
        else:
            print("⚠️ Icon file not found:", icon_path)

    def set_seminar_title(self):
        seminar_title = self.seminar_title_entry.get().strip()
        if seminar_title:
            self.seminar_title.append(seminar_title)
            save_json_data(SEMINAR_TITLE, self.seminar_title)
            messagebox.showinfo("Success", "Seminar title saved successfully!")
        else:
            messagebox.showerror("Error", "Please enter a seminar title.")

    def clear_seminar_title(self):
        self.seminar_title_entry.delete(0, tk.END)
        self.seminar_title = []
        save_json_data(SEMINAR_TITLE, self.seminar_title)
        messagebox.showinfo("Success", "Seminar title cleared successfully!")

    def create_seminar_title_frame(self, parent):
        frame = tk.LabelFrame(parent, text="Seminar Title", bg="#f0f0f0")
        frame.pack(pady=10, fill="x", padx=10)

        self.seminar_title_entry = tk.Entry(frame, width=50)
        self.seminar_title_entry.grid(row=0, column=0, padx=5, pady=5)

        tk.Button(frame, text="Set Title", command=self.set_seminar_title, width=15).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="Clear Title", command=self.clear_seminar_title, width=15).grid(row=0, column=2, padx=5)

    def set_logo_defaults(self, attr_name):
        logo_defaults = load_json_data(LOGO_DEFAULTS_FILE, {})
        entry = getattr(self, f"{attr_name}_entry").get()
        x = getattr(self, f"{attr_name}_x_entry").get()
        y = getattr(self, f"{attr_name}_y_entry").get()
        size = getattr(self, f"{attr_name}_size_entry").get()

        logo_defaults[attr_name] = {
            "path": entry,
            "x": int(x) if x else 0,
            "y": int(y) if y else 0,
            "size": int(size) if size else 100
        }
        save_json_data(LOGO_DEFAULTS_FILE, logo_defaults)
        messagebox.showinfo("Default Saved", f"Default values saved for {attr_name}")

    def clear_logo_fields(self, attr_name):
        getattr(self, f"{attr_name}_entry").delete(0, tk.END)
        getattr(self, f"{attr_name}_x_entry").delete(0, tk.END)
        getattr(self, f"{attr_name}_y_entry").delete(0, tk.END)
        getattr(self, f"{attr_name}_size_entry").delete(0, tk.END)

    def create_layout(self):
        left_frame = tk.Frame(self.root, bg="#f0f0f0")
        left_frame.pack(side="left", fill="y", padx=30, pady=10)

        # Right frame with canvas and scrollbar
        canvas = tk.Canvas(self.root, bg="#f0f0f0")
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Scrollable frame
        self.right_frame = tk.Frame(canvas, bg="#f0f0f0")
        canvas.create_window((0, 0), window=self.right_frame, anchor="nw")

        self.right_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        tk.Label(left_frame, text="Menu", font=("Arial", 17, "bold"), bg="#f0f0f0").pack(pady=15)
        modules = [
            ("Data Separate", self.run_data_separate),
            ("Generate Cert No", self.run_generate_cert_no),
            ("Remove Duplicates", self.run_remove_duplicates),
            ("Send Certificates Through Email", self.run_email_module),
            ("Generate Certificate", self.generate_certificate)
        ]
        for name, cmd in modules:
            tk.Button(left_frame, text=name,font=("Arial", 12, "bold"), command=cmd, width=27, height=2, bg="#C3D6D5", fg="#020E99").pack(pady=5)

        self.create_dropdown_with_entry(self.right_frame, "Series Name:", self.knowledge_options, KNOWLEDGE_FILE, "knowledge_text")
        self.create_dropdown_with_entry(self.right_frame, "Quotation:", self.passion_options, PASSION_FILE, "passion_text")
        self.create_seminar_title_frame(self.right_frame)
        self.create_file_selection(self.right_frame, "Select Certificate Template:", "Select Template", "template")
        
        for i in range(1, 5):
            self.create_file_selection(self.right_frame, f"Select Logo{i}:", f"Select Logo{i}", f"logo{i}")

        for i in range(1, 5):
            self.create_signature_frame(self.right_frame, f"Signature {i}", f"signature{i}")


        preview_frame = tk.LabelFrame(self.right_frame, text="Certificate Preview", bg="#f0f0f0")
        preview_frame.pack(pady=10, fill="both", expand=True)
        self.preview_canvas = tk.Canvas(preview_frame, bg="white", width=600, height=400)
        self.preview_canvas.pack(padx=10, pady=10)

    def create_dropdown_with_entry(self, parent, label_text, options, json_file, attr_name):
        frame = tk.LabelFrame(parent, text=label_text, bg="#f0f0f0")
        frame.pack(pady=10, fill="x", padx=10)
        frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="equal")

        var = tk.StringVar(value=options[0])
        menu = ttk.Combobox(frame, textvariable=var, values=options, state="readonly", width=30)
        menu.grid(row=0, column=0, padx=5, pady=10, sticky="w")

        entry = tk.Entry(frame, width=40)
        entry.insert(0, options[0])
        entry.grid(row=0, column=1, padx=5, pady=10, sticky="w")

        def update_entry(event):
            entry.delete(0, tk.END)
            entry.insert(0, var.get())

        def add_option():
            new_text = entry.get().strip()
            if new_text and new_text not in options:
                options.append(new_text)
                menu["values"] = options
                var.set(new_text)
                save_json_data(json_file, options)

        def delete_option():
            selected = var.get()
            if selected in options:
                options.remove(selected)
                menu["values"] = options
                if options:
                    var.set(options[0])
                    entry.delete(0, tk.END)
                    entry.insert(0, options[0])
                else:
                    var.set("")
                    entry.delete(0, tk.END)
                save_json_data(json_file, options)

        menu.bind("<<ComboboxSelected>>", update_entry)

        tk.Button(frame, text="Add", command=add_option, width=10).grid(row=0, column=2, padx=5)
        tk.Button(frame, text="Delete", command=delete_option, width=10).grid(row=0, column=3, padx=5)

        setattr(self, f"{attr_name}_var", var)
        setattr(self, f"{attr_name}_entry", entry)
    def create_signature_frame(self, parent, label_text, attr_name):
        frame = tk.LabelFrame(parent, text=label_text, bg="#f0f0f0")
        frame.pack(pady=5, fill="x", padx=10)

        # Image row
        entry = tk.Entry(frame, width=50)
        entry.grid(row=0, column=0, padx=5, pady=5)
        tk.Button(frame, text="Select Image", command=lambda: self.select_file(entry, attr_name), width=15).grid(row=0, column=1, padx=5)

        tk.Label(frame, text="X:", bg="#f0f0f0").grid(row=0, column=2)
        x_entry = tk.Entry(frame, width=5)
        x_entry.grid(row=0, column=3, padx=2)

        tk.Label(frame, text="Y:", bg="#f0f0f0").grid(row=0, column=4)
        y_entry = tk.Entry(frame, width=5)
        y_entry.grid(row=0, column=5, padx=2)

        tk.Label(frame, text="Width:", bg="#f0f0f0").grid(row=0, column=6)
        w_entry = tk.Entry(frame, width=5)
        w_entry.grid(row=0, column=7, padx=2)

        tk.Label(frame, text="Height:", bg="#f0f0f0").grid(row=0, column=8)
        h_entry = tk.Entry(frame, width=5)
        h_entry.grid(row=0, column=9, padx=2)

        # Name row
        tk.Label(frame, text="Name:", bg="#f0f0f0").grid(row=1, column=0, sticky="e", padx=5)
        name_entry = tk.Entry(frame, width=20)
        name_entry.grid(row=1, column=1, padx=2)

        tk.Label(frame, text="Font:", bg="#f0f0f0").grid(row=1, column=2)
        name_font_entry = tk.Entry(frame, width=5)
        name_font_entry.grid(row=1, column=3, padx=2)

        tk.Label(frame, text="X:", bg="#f0f0f0").grid(row=1, column=4)
        name_x_entry = tk.Entry(frame, width=5)
        name_x_entry.grid(row=1, column=5, padx=2)

        tk.Label(frame, text="Y:", bg="#f0f0f0").grid(row=1, column=6)
        name_y_entry = tk.Entry(frame, width=5)
        name_y_entry.grid(row=1, column=7, padx=2)

        # Designation row
        tk.Label(frame, text="Designation:", bg="#f0f0f0").grid(row=2, column=0, sticky="e", padx=5)
        desig_entry = tk.Entry(frame, width=20)
        desig_entry.grid(row=2, column=1, padx=2)

        tk.Label(frame, text="Font:", bg="#f0f0f0").grid(row=2, column=2)
        desig_font_entry = tk.Entry(frame, width=5)
        desig_font_entry.grid(row=2, column=3, padx=2)

        tk.Label(frame, text="X:", bg="#f0f0f0").grid(row=2, column=4)
        desig_x_entry = tk.Entry(frame, width=5)
        desig_x_entry.grid(row=2, column=5, padx=2)

        tk.Label(frame, text="Y:", bg="#f0f0f0").grid(row=2, column=6)
        desig_y_entry = tk.Entry(frame, width=5)
        desig_y_entry.grid(row=2, column=7, padx=2)

        # Load defaults if available
        defaults = load_json_data(LOGO_DEFAULTS_FILE, {}).get(attr_name, {})
        entry.insert(0, defaults.get("path", ""))
        x_entry.insert(0, str(defaults.get("x", "")))
        y_entry.insert(0, str(defaults.get("y", "")))
        w_entry.insert(0, str(defaults.get("width", "")))
        h_entry.insert(0, str(defaults.get("height", "")))
        name_entry.insert(0, defaults.get("name", ""))
        name_font_entry.insert(0, str(defaults.get("name_font", 20)))
        name_x_entry.insert(0, str(defaults.get("name_x", "")))
        name_y_entry.insert(0, str(defaults.get("name_y", "")))
        desig_entry.insert(0, defaults.get("designation", ""))
        desig_font_entry.insert(0, str(defaults.get("designation_font", 16)))
        desig_x_entry.insert(0, str(defaults.get("designation_x", "")))
        desig_y_entry.insert(0, str(defaults.get("designation_y", "")))

        # Bind preview update on entry change
        for widget in (
            x_entry, y_entry, w_entry, h_entry,
            name_font_entry, name_x_entry, name_y_entry,
            desig_font_entry, desig_x_entry, desig_y_entry
        ):
            widget.bind("<KeyRelease>", lambda e: self.update_preview())

        # Save as default
        def set_defaults():
            data = {
                "path": entry.get(),
                "x": int(x_entry.get() or 0),
                "y": int(y_entry.get() or 0),
                "width": int(w_entry.get() or 100),
                "height": int(h_entry.get() or 50),
                "name": name_entry.get(),
                "name_font": int(name_font_entry.get() or 20),
                "name_x": int(name_x_entry.get() or 0),
                "name_y": int(name_y_entry.get() or 0),
                "designation": desig_entry.get(),
                "designation_font": int(desig_font_entry.get() or 16),
                "designation_x": int(desig_x_entry.get() or 0),
                "designation_y": int(desig_y_entry.get() or 0)
            }
            all_defaults = load_json_data(LOGO_DEFAULTS_FILE, {})
            all_defaults[attr_name] = data
            save_json_data(LOGO_DEFAULTS_FILE, all_defaults)
            messagebox.showinfo("Default Saved", f"Default values saved for {attr_name}")

        # Clear fields
        def clear_fields():
            for widget in (
                entry, x_entry, y_entry, w_entry, h_entry,
                name_entry, name_font_entry, name_x_entry, name_y_entry,
                desig_entry, desig_font_entry, desig_x_entry, desig_y_entry
            ):
                widget.delete(0, tk.END)

        # Action buttons
        tk.Button(frame, text="Set as Default", command=set_defaults, width=15, bg="#2196F3", fg="white").grid(row=3, column=0, columnspan=2, pady=5)
        tk.Button(frame, text="Clear", command=clear_fields, width=10).grid(row=3, column=2, pady=5)

        # Attribute assignment for external access
        setattr(self, f"{attr_name}_entry", entry)
        setattr(self, f"{attr_name}_x_entry", x_entry)
        setattr(self, f"{attr_name}_y_entry", y_entry)
        setattr(self, f"{attr_name}_width_entry", w_entry)
        setattr(self, f"{attr_name}_height_entry", h_entry)
        setattr(self, f"{attr_name}_name_entry", name_entry)
        setattr(self, f"{attr_name}_name_font_entry", name_font_entry)
        setattr(self, f"{attr_name}_name_x_entry", name_x_entry)
        setattr(self, f"{attr_name}_name_y_entry", name_y_entry)
        setattr(self, f"{attr_name}_designation_entry", desig_entry)
        setattr(self, f"{attr_name}_designation_font_entry", desig_font_entry)
        setattr(self, f"{attr_name}_designation_x_entry", desig_x_entry)
        setattr(self, f"{attr_name}_designation_y_entry", desig_y_entry)

    def create_file_selection(self, parent, label_text, button_text, attr_name):
        frame = tk.LabelFrame(parent, text=label_text, bg="#f0f0f0")
        frame.pack(pady=5, fill="x")      

        entry = tk.Entry(frame, width=50)
        entry.grid(row=0, column=0, padx=5, pady=5)

        tk.Button(frame, text=button_text, command=lambda: self.select_file(entry, attr_name), width=15).grid(row=0, column=1, padx=5)

        logo_defaults = load_json_data(LOGO_DEFAULTS_FILE, {})
        defaults = logo_defaults.get(attr_name, {})

        if attr_name != "template":
            tk.Label(frame, text="X:", bg="#f0f0f0").grid(row=0, column=2)
            x_entry = tk.Entry(frame, width=5)
            x_entry.grid(row=0, column=3, padx=2)

            tk.Label(frame, text="Y:", bg="#f0f0f0").grid(row=0, column=4)
            y_entry = tk.Entry(frame, width=5)
            y_entry.grid(row=0, column=5, padx=2)

            tk.Label(frame, text="Size:", bg="#f0f0f0").grid(row=0, column=6)
            size_entry = tk.Entry(frame, width=5)
            size_entry.grid(row=0, column=7, padx=2)

            entry.insert(0, defaults.get("path", ""))
            x_entry.insert(0, str(defaults.get("x", "")))
            y_entry.insert(0, str(defaults.get("y", "")))
            size_entry.insert(0, str(defaults.get("size", "")))

            # ✅ Bind key release to update preview dynamically
            for widget in (x_entry, y_entry, size_entry):
                widget.bind("<KeyRelease>", lambda e: self.update_preview())

            setattr(self, f"{attr_name}_x_entry", x_entry)
            setattr(self, f"{attr_name}_y_entry", y_entry)
            setattr(self, f"{attr_name}_size_entry", size_entry)

            tk.Button(frame, text="Set as Default", command=lambda: self.set_logo_defaults(attr_name), width=15, bg="#2196F3", fg="white").grid(row=0, column=8, padx=5)
            tk.Button(frame, text="Clear", command=lambda: self.clear_logo_fields(attr_name), width=10).grid(row=0, column=9, padx=5)

        setattr(self, f"{attr_name}_entry", entry)

    def select_file(self, entry, attr_name):
        file = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg"), ("All Files", "*.*")])
        if file:
            entry.delete(0, tk.END)
            entry.insert(0, file)
            self.update_preview()

    def update_preview(self):
        try:
            template_path = self.template_entry.get()
            if not os.path.exists(template_path):
                return

            template = Image.open(template_path).convert("RGBA")
            draw = ImageDraw.Draw(template)

            # Paste logo images
            for i in range(1, 5):
                logo_entry = getattr(self, f"logo{i}_entry", None)
                if not logo_entry:
                    continue
                path = logo_entry.get()
                if os.path.exists(path):
                    x = int(getattr(self, f"logo{i}_x_entry").get() or 0)
                    y = int(getattr(self, f"logo{i}_y_entry").get() or 0)
                    size = int(getattr(self, f"logo{i}_size_entry").get() or 100)
                    logo_img = Image.open(path).convert("RGBA").resize((size, size))
                    template.paste(logo_img, (x, y), logo_img)

            # Paste signatures and add name + designation
            for i in range(1, 5):
                x = int(getattr(self, f"signature{i}_x_entry").get() or 0)
                y = int(getattr(self, f"signature{i}_y_entry").get() or 0)
                width = int(getattr(self, f"signature{i}_width_entry").get() or 100)
                height = int(getattr(self, f"signature{i}_height_entry").get() or 50)

                # Try pasting signature image
                sig_entry = getattr(self, f"signature{i}_entry", None)
                if sig_entry:
                    path = sig_entry.get()
                    if os.path.exists(path):
                        sig_img = Image.open(path).convert("RGBA").resize((width, height))
                        template.paste(sig_img, (x, y), sig_img)

                # Name & Designation drawing (independent of image)
                name = getattr(self, f"signature{i}_name_entry", None).get()
                name_x = int(getattr(self, f"signature{i}_name_x_entry").get() or 0)
                name_y = int(getattr(self, f"signature{i}_name_y_entry").get() or 0)
                name_size = int(getattr(self, f"signature{i}_name_font_entry").get() or 20)

                designation = getattr(self, f"signature{i}_designation_entry", None).get()
                desig_x = int(getattr(self, f"signature{i}_designation_x_entry").get() or 0)
                desig_y = int(getattr(self, f"signature{i}_designation_y_entry").get() or 0)
                desig_size = int(getattr(self, f"signature{i}_designation_font_entry").get() or 18)

                try:
                    name_font = ImageFont.truetype("arial.ttf", name_size)
                except:
                    name_font = ImageFont.load_default()

                try:
                    desig_font = ImageFont.truetype("arial.ttf", desig_size)
                except:
                    desig_font = ImageFont.load_default()

                # Draw name and designation
                if name:
                    draw.text((name_x, name_y), name, font=name_font, fill="black")
                if designation:
                    draw.text((desig_x, desig_y), designation, font=desig_font, fill="black")

            # Maintain preview ratio
            preview_width = 600
            preview_height = int(preview_width / (2000 / 1414))

            resized = template.resize((preview_width, preview_height))
            self.tk_preview = ImageTk.PhotoImage(resized)
            self.preview_canvas.config(width=preview_width, height=preview_height)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(0, 0, anchor="nw", image=self.tk_preview)

        except Exception as e:
            print("Preview error:", e)
       
    # In your button functions
    def run_data_separate(self):
        data_separate_main()

    def run_generate_cert_no(self):
        cert_no_main()

    def run_remove_duplicates(self):
        remove_duplicates_main()

    def run_email_module(self):
        email_main()
    def generate_certificate(self):
        knowledge_text = self.knowledge_text_entry.get().strip() or self.knowledge_text_var.get()
        passion_text = self.passion_text_entry.get().strip() or self.passion_text_var.get()
        template = self.template_entry.get().strip()
        seminar_title = self.seminar_title_entry.get().strip() or ""
        if not template:
            messagebox.showerror("Error", "Please select a certificate template.")
            return

        logos = []
        x_positions = []
        y_positions = []
        size_values = []

        for i in range(1, 5):
            logo_path = getattr(self, f"logo{i}_entry").get().strip()
            x = getattr(self, f"logo{i}_x_entry").get().strip() or "0"
            y = getattr(self, f"logo{i}_y_entry").get().strip() or "0"
            size = getattr(self, f"logo{i}_size_entry").get().strip() or "0"

            logos.append(logo_path)
            x_positions.append(int(x))
            y_positions.append(int(y))
            size_values.append(int(size))

        signatures = []
        names = []
        designations = []

        for i in range(1, 5):
            sig_path = getattr(self, f"signature{i}_entry").get().strip()
            name = getattr(self, f"signature{i}_name_entry").get().strip()
            desig = getattr(self, f"signature{i}_designation_entry").get().strip()

            signatures.append(sig_path if sig_path else None)
            names.append(name if name else None)
            designations.append(desig if desig else None)

        if not template:
            messagebox.showerror("Error", "Please select a certificate template.")
            return

        main2.knowledge_text = knowledge_text
        main2.passion_text = passion_text
        main2.template = template
        main2.seminar_title = seminar_title

        for i in range(4):
            setattr(main2, f"logo{i+1}", logos[i] if logos[i] else None)
            setattr(main2, f"x{i+1}", x_positions[i])
            setattr(main2, f"y{i+1}", y_positions[i])
            setattr(main2, f"size{i+1}", size_values[i])

        enabled_signatures = []  # Define once before the loop

        for i in range(4):
            setattr(main2, f"signature{i+1}", signatures[i])
            setattr(main2, f"name{i+1}", names[i])
            setattr(main2, f"designation{i+1}", designations[i])
            setattr(main2, f"signature{i+1}_x", int(getattr(self, f"signature{i+1}_x_entry").get().strip() or 0))
            setattr(main2, f"signature{i+1}_y", int(getattr(self, f"signature{i+1}_y_entry").get().strip() or 0))
            setattr(main2, f"signature{i+1}_width", int(getattr(self, f"signature{i+1}_width_entry").get().strip() or 100))
            setattr(main2, f"signature{i+1}_height", int(getattr(self, f"signature{i+1}_height_entry").get().strip() or 50))
            setattr(main2, f"name{i+1}_font", int(getattr(self, f"signature{i+1}_name_font_entry").get().strip() or 20))
            setattr(main2, f"designation{i+1}_font", int(getattr(self, f"signature{i+1}_designation_font_entry").get().strip() or 16))
            setattr(main2, f"name{i+1}_x", int(getattr(self, f"signature{i+1}_name_x_entry").get().strip() or 0))
            setattr(main2, f"name{i+1}_y", int(getattr(self, f"signature{i+1}_name_y_entry").get().strip() or 0))
            setattr(main2, f"designation{i+1}_x", int(getattr(self, f"signature{i+1}_designation_x_entry").get().strip() or 0))
            setattr(main2, f"designation{i+1}_y", int(getattr(self, f"signature{i+1}_designation_y_entry").get().strip() or 0))

            # Add to enabled_signatures if signature is provided
            if signatures[i]:
                enabled_signatures.append(i + 1)

        main2.enabled_signatures = enabled_signatures  # Set after loop

        try:
            main2.generate_certificates()
            messagebox.showinfo("Success", "Certificates generated successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error generating certificates: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()
