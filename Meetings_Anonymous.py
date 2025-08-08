import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
import json
import pyperclip
import os
import en_core_web_sm

spacy_nlp = en_core_web_sm.load()

class MeetingSummarizerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Teams Meeting Summarizer")

        # Menu bar
        menubar = tk.Menu(self.master)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Import Names CSV", command=self.import_names_csv)
        file_menu.add_command(label="Export Names CSV", command=self.export_names_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Edit Default Text", command=self.edit_defaults)
        menubar.add_cascade(label="File", menu=file_menu)
        self.master.config(menu=menubar)

        # Default text settings
        self.default_prompt_prefix = "Please summarize this meeting:\n\n"
        self.default_email_intro = "Hi Team,\n\nHere are the notes from the recent meeting:\n\n"
        self.default_email_signature = "\n\nPlease let me know if you have any additional questions"

        # Load saved defaults if available
        defaults_file = "defaults.json"
        if os.path.exists(defaults_file):
            with open(defaults_file, "r") as f:
                try:
                    defaults = json.load(f)
                    self.default_prompt_prefix = defaults.get("prompt_prefix", self.default_prompt_prefix)
                    self.default_email_intro = defaults.get("email_intro", self.default_email_intro)
                    self.default_email_signature = defaults.get("email_signature", self.default_email_signature)
                except json.JSONDecodeError:
                    pass
                
        self.transcript = ""
        self.people_names = []
        self.company_names = []
        self.name_map = {}
        self.names_file = "names.json"

        if os.path.exists(self.names_file):
            with open(self.names_file, 'r') as f:
                data = json.load(f)
                self.people_names = data.get("people", [])
                self.company_names = data.get("companies", [])

        self.transcript_text = tk.Text(master, wrap="word", height=15)
        self.transcript_text.pack(padx=10, pady=5, fill="both", expand=True)

        self.summary_text = tk.Text(master, wrap="word", height=15)
        self.summary_text.pack(padx=10, pady=5, fill="both", expand=True)

        button_frame = tk.Frame(master)
        button_frame.pack(pady=5)

        tk.Button(button_frame, text="Select Transcript", command=self.select_file, takefocus=False).pack(side="left", padx=5)
        tk.Button(button_frame, text="Suggest Names", command=self.suggest_names, takefocus=False).pack(side="left", padx=5)
        tk.Button(button_frame, text="Edit Names", command=self.edit_names, takefocus=False).pack(side="left", padx=5)
        tk.Button(button_frame, text="Anonymize + Copy", command=self.anonymize_and_copy, takefocus=False).pack(side="left", padx=5)
        tk.Button(button_frame, text="Paste Summary", command=self.paste_summary, takefocus=False).pack(side="left", padx=5)
    def save_names(self):
        with open(self.names_file, 'w') as f:
            json.dump({"people": self.people_names, "companies": self.company_names}, f)

    def import_names_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    self.people_names = [line.strip() for line in lines[0].split(',') if line.strip()]
                    self.company_names = [line.strip() for line in lines[1].split(',') if line.strip()]
                    self.save_names()
                    messagebox.showinfo("Import Complete", "Name lists imported successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import CSV: {e}")

    def export_names_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(','.join(self.people_names) + '\n')
                    f.write(','.join(self.company_names) + '\n')
                    messagebox.showinfo("Export Complete", "Name lists exported successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export CSV: {e}")


        with open(self.names_file, 'w') as f:
            json.dump({"people": self.people_names, "companies": self.company_names}, f)

    def edit_names(self):
        popup = Toplevel(self.master)
        popup.grab_set()
        popup.focus_set()
        popup.title("Edit People and Company Names")

        tk.Label(popup, text="People Names (comma-separated):").pack()
        people_entry = tk.Text(popup, height=4, width=60)
        people_entry.insert("1.0", ", ".join(self.people_names))
        people_entry.pack()

        tk.Label(popup, text="Company Names (comma-separated):").pack()
        company_entry = tk.Text(popup, height=4, width=60)
        company_entry.insert("1.0", ", ".join(self.company_names))
        company_entry.pack()

        def save_and_close():
            self.people_names = [p.strip() for p in people_entry.get("1.0", tk.END).split(",") if p.strip()]
            self.company_names = [c.strip() for c in company_entry.get("1.0", tk.END).split(",") if c.strip()]
            self.save_names()
            popup.destroy()

        tk.Button(popup, text="Save", command=save_and_close).pack(pady=5)

    def suggest_names(self):
        transcript = self.transcript_text.get("1.0", tk.END)
        doc = spacy_nlp(transcript)

        suggested_people = set()
        suggested_companies = set()

        for ent in doc.ents:
            if ent.label_ == "PERSON" and ent.text not in self.people_names and len(ent.text.split()) <= 2:
                suggested_people.add(ent.text)
            elif ent.label_ in ("ORG", "GPE") and ent.text not in self.company_names and len(ent.text.split()) <= 3:
                suggested_companies.add(ent.text)

        if suggested_people or suggested_companies:
            self.show_suggestions_popup(suggested_people, suggested_companies)
        else:
            messagebox.showinfo("No Suggestions", "No new names found.")
    
    def show_suggestions_popup(self, suggested_people, suggested_companies):
        popup = Toplevel(self.master)
        popup.title("Review Suggested Names")

        container = tk.Frame(popup)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tk.Label(scrollable_frame, text="Suggested People Names:").pack()
        people_vars = {}
        for name in sorted(suggested_people):
            var = tk.BooleanVar(value=True)
            tk.Checkbutton(scrollable_frame, text=name, variable=var).pack(anchor="w")
            people_vars[name] = var

        tk.Label(scrollable_frame, text="Suggested Company Names:").pack()
        company_vars = {}
        for name in sorted(suggested_companies):
            var = tk.BooleanVar(value=True)
            tk.Checkbutton(scrollable_frame, text=name, variable=var).pack(anchor="w")
            company_vars[name] = var

        def add_selected():
            self.people_names.extend([name for name, var in people_vars.items() if var.get()])
            self.company_names.extend([name for name, var in company_vars.items() if var.get()])
            self.people_names = list(set(self.people_names))
            self.company_names = list(set(self.company_names))
            self.save_names()
            popup.destroy()
            messagebox.showinfo("Saved", "Suggested names added to your list.")

        def select_all():
            for var in list(people_vars.values()) + list(company_vars.values()):
                var.set(True)

        def deselect_all():
            for var in list(people_vars.values()) + list(company_vars.values()):
                var.set(False)

        action_frame = tk.Frame(popup)
        action_frame.pack(pady=5)

        tk.Button(action_frame, text="Select All", command=select_all).pack(side="left", padx=5)
        tk.Button(action_frame, text="Deselect All", command=deselect_all).pack(side="left", padx=5)
        tk.Button(action_frame, text="Add Selected", command=add_selected).pack(side="left", padx=5)


    def select_file(self):
        file_path = filedialog.askopenfilename()
        if not file_path:
            return
        with open(file_path, 'r', encoding='utf-8') as f:
            self.transcript = f.read()
        self.transcript_text.delete("1.0", tk.END)
        self.transcript_text.insert(tk.END, self.transcript)

    def anonymize_text(self, text, people_names, company_names):
        name_map = {}

        for i, name in enumerate(people_names):
            placeholder = f"Person_{i}"
            name_map[placeholder] = name
            text = text.replace(name, placeholder)

        for i, company in enumerate(company_names):
            placeholder = f"Company_{i}"
            name_map[placeholder] = company
            text = text.replace(company, placeholder)

        with open("name_map.json", "w") as f:
            json.dump(name_map, f)

        self.name_map = name_map
        return text

    def deanonymize_text(self, text):
        text = text.replace(r"\_", "_")
        import re
        import unicodedata
    
        with open("name_map.json", "r") as f:
            name_map = json.load(f)
    
        def replace_placeholder(match):
            placeholder = match.group(0)
            return name_map.get(placeholder, placeholder)
    
        # Use regex to match placeholders like Person_0, Company_1
        text = re.sub(r'\b(Person|Company)_\d+\b', replace_placeholder, text)
    
        # Remove unwanted infographic icons (but keep bullets)
        unwanted_symbols = ['➤', '✔', '✅', '✏️', '📌', '🔸', '🔹', '➡️', '📝', '❗', '👉']
        for symbol in unwanted_symbols:
            text = text.replace(symbol, '')
    
        text = unicodedata.normalize("NFKC", text)
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\n{3,}', '\n\n', text)
    
        return text.strip()


    def export_draft_email(self, content):
        attach_transcript = messagebox.askyesno("Attach Transcript", "Include original transcript as attachment?")
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".eml", filetypes=[("Outlook Draft Email", "*.eml")])
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("To: \n")
                    f.write("Subject: Meeting Summary\n")
                    f.write("MIME-Version: 1.0\n")

                    if attach_transcript:
                        f.write("Content-Type: multipart/mixed; boundary=\"sep\"\n\n")
                        f.write("--sep\n")
                        f.write("Content-Type: text/plain; charset=UTF-8\n\n")
                        f.write(content + "\n")
                        f.write("--sep\n")
                        f.write("Content-Type: text/plain; name=transcript.txt\n")
                        f.write("Content-Disposition: attachment; filename=transcript.txt\n\n")
                        f.write(self.transcript)
                        f.write("\n--sep--")
                    else:
                        f.write("Content-Type: text/plain; charset=UTF-8\n\n")
                        f.write(content)

                messagebox.showinfo("Success", f"Draft email saved to: {file_path}")
                try:
                    os.startfile(file_path)
                except Exception as open_err:
                    messagebox.showwarning("Open Failed", f"Could not open the email draft automatically: {open_err}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export email: {e}")
  


    def review_and_edit(self, prompt_text, title="Edit Text"):
        popup = Toplevel(self.master)
        popup.title(title)

        text_widget = tk.Text(popup, wrap="word", width=80, height=30)
        text_widget.insert("1.0", prompt_text)
        text_widget.pack(expand=True, fill="both")

        def on_confirm():
            popup.edited_text = text_widget.get("1.0", tk.END)
            popup.destroy()

        confirm_btn = tk.Button(popup, text="Confirm", command=on_confirm)
        confirm_btn.pack(pady=5)

        popup.wait_window()
        return getattr(popup, 'edited_text', prompt_text)

    def anonymize_and_copy(self):
        transcript = self.transcript_text.get("1.0", tk.END)
        if not transcript.strip():
            messagebox.showerror("Error", "Transcript is empty.")
            return

        if not self.people_names or not self.company_names:
            messagebox.showerror("Error", "People or company names list is empty. Edit them first.")
            return

        try:
            anonymized = self.anonymize_text(transcript, self.people_names, self.company_names)
            edited_anonymized = self.review_and_edit(anonymized, title="Review and Edit Anonymized Transcript")

            prompt = f"{self.default_prompt_prefix}\n\n{edited_anonymized}"
            edited_prompt = self.review_and_edit(prompt, title="Edit Prompt to ChatGPT")

            pyperclip.copy(edited_prompt)
            messagebox.showinfo("Copied", "Prompt copied to clipboard. Paste it into ChatGPT to summarize.")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def edit_defaults(self):
        popup = Toplevel(self.master)
        popup.title("Edit Default Text")

        def make_text_section(label_text, default_val):
            tk.Label(popup, text=label_text).pack()
            text_widget = tk.Text(popup, height=4, width=60)
            text_widget.insert("1.0", default_val)
            text_widget.pack()
            return text_widget

        prompt_widget = make_text_section("ChatGPT Prompt Prefix:", self.default_prompt_prefix)
        intro_widget = make_text_section("Email Intro:", self.default_email_intro)
        sign_widget = make_text_section("Email Signature:", self.default_email_signature)

        def save_and_close():
            self.default_prompt_prefix = prompt_widget.get("1.0", tk.END).strip()
            self.default_email_intro = intro_widget.get("1.0", tk.END).strip()
            self.default_email_signature = sign_widget.get("1.0", tk.END).strip()

            with open("defaults.json", "w") as f:
                json.dump({
                    "prompt_prefix": self.default_prompt_prefix,
                    "email_intro": self.default_email_intro,
                    "email_signature": self.default_email_signature
                }, f)

            popup.destroy()

        tk.Button(popup, text="Save", command=save_and_close).pack(pady=5)

    def paste_summary(self):
        try:
            clipboard_content = pyperclip.paste()
            if not clipboard_content.strip():
                messagebox.showerror("Error", "Clipboard is empty.")
                return

            deanonymized = self.deanonymize_text(clipboard_content)
            email_draft = f"Subject: Meeting Summary\n\n{self.default_email_intro}{deanonymized}{self.default_email_signature}"

            self.summary_text.delete("1.0", tk.END)
            self.summary_text.insert(tk.END, email_draft)

            pyperclip.copy(email_draft)
            messagebox.showinfo("Copied", "Email draft copied to clipboard.")

        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = MeetingSummarizerApp(root)
    root.mainloop()
