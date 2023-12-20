import tkinter as tk
from tkinter import filedialog
from docx import Document
from collections import Counter
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string
import pandas as pd

# Download NLTK data if not already present
# import nltk
# nltk.download('punkt')
# nltk.download('stopwords')

class TermExtractorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Term Extractor")

        self.file_path_label = tk.Label(master, text="Select Word document:")
        self.file_path_label.pack()

        self.file_path_entry = tk.Entry(master, width=50, state='disabled')
        self.file_path_entry.pack()

        self.browse_button = tk.Button(master, text="Browse", command=self.browse_file)
        self.browse_button.pack()

        self.extract_button = tk.Button(master, text="Extract Terms", command=self.extract_terms)
        self.extract_button.pack()

        self.export_button = tk.Button(master, text="Export to Excel", command=self.export_to_excel)
        self.export_button.pack()

        self.result_label = tk.Label(master, text="Key Terms:")
        self.result_label.pack()

        self.result_text = tk.Text(master, height=10, width=50)
        self.result_text.pack()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self.file_path_entry.config(state='normal')
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)
            self.file_path_entry.config(state='disabled')

    def extract_terms(self):
        file_path = self.file_path_entry.get()

        if not file_path:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Please select a Word document.")
            return

        document = Document(file_path)
        text = " ".join([paragraph.text for paragraph in document.paragraphs])

        # Tokenize the text
        tokens = word_tokenize(text)

        # Remove punctuation and stopwords
        stop_words = set(stopwords.words('english') + list(string.punctuation))
        filtered_tokens = [word.lower() for word in tokens if word.lower() not in stop_words]

        # Extract key terms using Counter
        term_counter = Counter(filtered_tokens)
        key_terms = [term for term, count in term_counter.most_common(10)]

        # Display key terms in the UI
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "\n".join(key_terms))

        # Store key terms for export
        self.key_terms_for_export = key_terms

    def export_to_excel(self):
        if hasattr(self, 'key_terms_for_export') and self.key_terms_for_export:
            export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=[("Excel Files", "*.xlsx")])
            if export_path:
                df = pd.DataFrame({"Key Terms": self.key_terms_for_export})
                df.to_excel(export_path, index=False)
                tk.messagebox.showinfo("Export Successful", "Key terms exported to Excel file.")
        else:
            tk.messagebox.showwarning("No Key Terms", "Please extract key terms first.")

def main():
    root = tk.Tk()
    app = TermExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
