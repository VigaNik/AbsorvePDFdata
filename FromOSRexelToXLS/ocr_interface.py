import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import os
import threading
import csv
from pathlib import Path
from OSTtessToPDF import (
    footnoteConfig,
    footnoteProcessor,
    save_footnotes_to_xml,
    save_footnotes_to_csv,
    extract_issue_number_from_filename,
    extract_journal_name_from_path
)

# Import paper_abbrev functionality
try:
    from abbreviations import paper_abbrev

    ABBREV_AVAILABLE = True
except ImportError:
    paper_abbrev = None
    ABBREV_AVAILABLE = False
    print("Warning: abbreviations module not available. Paper abbreviation processing will be disabled.")


class JournalConfigManager:
    """Manages journal configurations to avoid duplication"""

    @staticmethod
    def get_journal_configs():
        """Returns the complete journal configuration dictionary"""
        return {
            "tarbiz": {
                "printed": {
                    "bottom_margin_min": 1605, "bottom_margin_max": 1670,
                    "left_margin_threshold_even": 195, "left_margin_threshold_odd": 295,
                    "width_threshold_even": 1070, "width_threshold_odd": 1160,
                    "merge_footnotes_threshold_even": 1050, "merge_footnotes_threshold_odd": 1080,
                    "footnotes_spleat_threshold_even": 1070, "footnotes_spleat_threshold_odd": 1180,
                    "total_left": 7200
                },
                "scanned": {
                    "bottom_margin_min": 1605, "bottom_margin_max": 1695,
                    "left_margin_threshold_even": 195, "left_margin_threshold_odd": 195,
                    "width_threshold_even": 1095, "width_threshold_odd": 1095,
                    "merge_footnotes_threshold_even": 1070, "merge_footnotes_threshold_odd": 1070,
                    "footnotes_spleat_threshold_even": 1085, "footnotes_spleat_threshold_odd": 1085,
                    "total_left": 7200
                }
            },
            "meghillot": {
                "printed": {
                    "bottom_margin_min": 1670, "bottom_margin_max": 1680,
                    "left_margin_threshold_even": 220, "left_margin_threshold_odd": 220,
                    "width_threshold_even": 1027, "width_threshold_odd": 1027,
                    "merge_footnotes_threshold_even": 1027, "merge_footnotes_threshold_odd": 1140,
                    "footnotes_spleat_threshold_even": 1080, "footnotes_spleat_threshold_odd": 1150,
                    "total_left": 6700
                },
                "scanned": {
                    "bottom_margin_min": 1670, "bottom_margin_max": 1680,
                    "left_margin_threshold_even": 220, "left_margin_threshold_odd": 220,
                    "width_threshold_even": 1027, "width_threshold_odd": 1027,
                    "merge_footnotes_threshold_even": 1027, "merge_footnotes_threshold_odd": 1027,
                    "footnotes_spleat_threshold_even": 1050, "footnotes_spleat_threshold_odd": 1050,
                    "total_left": 6700
                }
            },
            "shenmishivri": {

                "printed": {
                    "bottom_margin_min": 1645, "bottom_margin_max": 1685,
                    "left_margin_threshold_even": 205, "left_margin_threshold_odd": 205,
                    "width_threshold_even": 1075, "width_threshold_odd": 1075,
                    "merge_footnotes_threshold_even":1045, "merge_footnotes_threshold_odd": 1045,
                    "footnotes_spleat_threshold_even": 1055, "footnotes_spleat_threshold_odd": 1055,
                    "total_left": 7000
                },
                "scanned": {
                    "bottom_margin_min": 1645, "bottom_margin_max": 1685,
                    "left_margin_threshold_even": 205, "left_margin_threshold_odd": 205,
                    "width_threshold_even": 1075, "width_threshold_odd": 1075,
                    "merge_footnotes_threshold_even": 1045, "merge_footnotes_threshold_odd": 1045,
                    "footnotes_spleat_threshold_even": 1055, "footnotes_spleat_threshold_odd": 1055,
                    "total_left": 7000
                }
            },
            "sibra": {
                "printed": {
                    "bottom_margin_min": 1680, "bottom_margin_max": 1712,
                    "left_margin_threshold_even": 195, "left_margin_threshold_odd": 295,
                    "width_threshold_even": 1090, "width_threshold_odd": 1090,
                    "merge_footnotes_threshold_even": 1050, "merge_footnotes_threshold_odd": 1140,
                    "footnotes_spleat_threshold_even": 1080, "footnotes_spleat_threshold_odd": 1150,
                    "total_left": 7200
                },
                "scanned": {
                    "bottom_margin_min": 1680, "bottom_margin_max": 1712,
                    "left_margin_threshold_even": 195, "left_margin_threshold_odd": 195,
                    "width_threshold_even": 1090, "width_threshold_odd": 1090,
                    "merge_footnotes_threshold_even": 1070, "merge_footnotes_threshold_odd": 1070,
                    "footnotes_spleat_threshold_even": 1080, "footnotes_spleat_threshold_odd": 1080,
                    "total_left": 7200
                }
            },
            "zion": {
                "printed": {
                    "bottom_margin_min": 1680, "bottom_margin_max": 1698,
                    "left_margin_threshold_even": 225, "left_margin_threshold_odd": 225,
                    "width_threshold_even": 1080, "width_threshold_odd": 1080,
                    "merge_footnotes_threshold_even": 1050, "merge_footnotes_threshold_odd": 1140,
                    "footnotes_spleat_threshold_even": 1080, "footnotes_spleat_threshold_odd": 1150,
                    "total_left": 7200
                },
                "scanned": {
                    "bottom_margin_min": 1680, "bottom_margin_max": 1698,
                    "left_margin_threshold_even": 225, "left_margin_threshold_odd": 225,
                    "width_threshold_even": 1080, "width_threshold_odd": 1080,
                    "merge_footnotes_threshold_even": 1070, "merge_footnotes_threshold_odd": 1070,
                    "footnotes_spleat_threshold_even": 1080, "footnotes_spleat_threshold_odd": 1080,
                    "total_left": 7200
                }
            },
            "leshonenu": {
                "printed": {
                    "bottom_margin_min": 1680, "bottom_margin_max": 1665,
                    "left_margin_threshold_even": 220, "left_margin_threshold_odd": 220,
                    "width_threshold_even": 1077, "width_threshold_odd": 1077,
                    "merge_footnotes_threshold_even": 1050, "merge_footnotes_threshold_odd": 1140,
                    "footnotes_spleat_threshold_even": 1080, "footnotes_spleat_threshold_odd": 1150,
                    "total_left": 7200
                },
                "scanned": {
                    "bottom_margin_min": 1680, "bottom_margin_max": 1665,
                    "left_margin_threshold_even": 220, "left_margin_threshold_odd": 220,
                    "width_threshold_even": 1077, "width_threshold_odd": 1077,
                    "merge_footnotes_threshold_even": 1070, "merge_footnotes_threshold_odd": 1070,
                    "footnotes_spleat_threshold_even": 1080, "footnotes_spleat_threshold_odd": 1080,
                    "total_left": 7200
                }
            }
        }

    @staticmethod
    def get_config_for_journal(journal_key: str, doc_type: str) -> dict:
        """Get configuration for a specific journal and document type"""
        configs = JournalConfigManager.get_journal_configs()
        return configs.get(journal_key, {}).get(doc_type, {})

    @staticmethod
    def create_footnote_config(journal_key: str, doc_type: str) -> footnoteConfig:
        """Create a footnoteConfig object for the specified journal and type"""
        params = JournalConfigManager.get_config_for_journal(journal_key, doc_type)

        if not params:
            raise ValueError(f"No configuration found for journal '{journal_key}' type '{doc_type}'")

        return footnoteConfig(
            exclusion_phrases=[
                "https://about,jstor.org/terms",
                "[תרביץ", "(תרביץ",
                "https://about.jstor.org/terms",
                "https://aboutjstor.org/terms"
            ],
            start_row=1,
            **params  # Unpack all the parameter values
        )


class ProcessingTaskManager:
    """Manages different types of processing tasks"""

    def __init__(self, interface):
        self.interface = interface

    def process_footnotes_folder(self, input_dir, output_dir, meta_dir, journal_key, doc_type):
        """Process footnotes for an entire folder"""
        try:
            self.interface.log_message("Starting folder footnote processing...")

            xlsx_files = [f for f in os.listdir(input_dir) if f.endswith(".xlsx")]
            if not xlsx_files:
                self.interface.log_message("No .xlsx files found in input directory")
                messagebox.showinfo("Info", "No .xlsx files found")
                return

            # Create configuration using the manager
            config = JournalConfigManager.create_footnote_config(journal_key, doc_type)
            processor = footnoteProcessor(config)

            journal_name = extract_journal_name_from_path(input_dir)
            total_footnotes = 0
            total_meta_refs = 0
            processed_files = 0

            for i, xlsx_file in enumerate(xlsx_files):
                self.interface.progress_var.set(int((i / len(xlsx_files)) * 100))

                file_path = os.path.join(input_dir, xlsx_file)
                base_name = os.path.splitext(xlsx_file)[0]
                issue_number = extract_issue_number_from_filename(xlsx_file)

                # Extract meta information
                meta_file_path = os.path.join(meta_dir, f"{base_name}.json")
                meta_info = processor.extract_meta_info(meta_file_path)

                result_data = {
                    "issue_number": issue_number,
                    "filename": xlsx_file,
                    "meta_references": meta_info["number_of_references"],
                    "meta_labels": meta_info["biggest_label_number"],
                    "collected_footnotes": 0,
                    "has_meta_file": meta_info["has_meta_file"],
                    "status": "Processing..."
                }

                try:
                    self.interface.log_message(f"Processing file: {xlsx_file}")

                    # Process the file
                    all_footnotes, main_texts = processor.process_workbook(file_path)

                    # Save results
                    output_xml = os.path.join(output_dir, f"{base_name}_footnotes.xml")
                    save_footnotes_to_xml(all_footnotes, main_texts, output_xml)
                    save_footnotes_to_csv(all_footnotes, main_texts, output_xml)

                    ref_count = len(all_footnotes)
                    total_footnotes += ref_count
                    total_meta_refs += meta_info["number_of_references"]
                    processed_files += 1

                    result_data["collected_footnotes"] = ref_count
                    result_data["status"] = "Completed"

                    self.interface.log_message(f"Processed {xlsx_file}: {ref_count} footnotes")

                except Exception as e:
                    self.interface.log_message(f"Error processing {xlsx_file}: {e}")
                    result_data["status"] = f"Error: {str(e)[:50]}..."

                self.interface.update_results_table(result_data)
                self.interface.processing_results.append(result_data)

            self.interface.progress_var.set(100)
            self.interface.update_summary(processed_files, total_footnotes, total_meta_refs, journal_name)
            self.interface.save_csv_button.config(state=tk.NORMAL)

            messagebox.showinfo(
                "Processing Complete",
                f"Processing complete!\nFiles processed: {processed_files}\n"
                f"Footnotes collected: {total_footnotes}\nMeta references: {total_meta_refs}"
            )

        except Exception as e:
            self.interface.log_message(f"Error during processing: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.interface.start_button.config(state=tk.NORMAL)

    def process_footnotes_single(self, file_path, output_dir, journal_key, doc_type):
        """Process footnotes for a single file"""
        try:
            self.interface.log_message("Starting single file footnote processing...")

            filename = os.path.basename(file_path)
            base_name = os.path.splitext(filename)[0]
            issue_number = extract_issue_number_from_filename(filename)
            journal_name = extract_journal_name_from_path(file_path)

            # Create configuration using the manager
            config = JournalConfigManager.create_footnote_config(journal_key, doc_type)
            processor = footnoteProcessor(config)

            result_data = {
                "issue_number": issue_number,
                "filename": filename,
                "meta_references": 0,
                "meta_labels": 0,
                "collected_footnotes": 0,
                "has_meta_file": False,
                "status": "Processing..."
            }

            try:
                self.interface.log_message(f"Processing file: {filename}")
                self.interface.progress_var.set(50)

                # Process the file
                all_footnotes, main_texts = processor.process_workbook(file_path)

                # Save results
                output_xml = os.path.join(output_dir, f"{base_name}_footnotes.xml")
                save_footnotes_to_xml(all_footnotes, main_texts, output_xml)
                save_footnotes_to_csv(all_footnotes, main_texts, output_xml)

                ref_count = len(all_footnotes)
                result_data["collected_footnotes"] = ref_count
                result_data["status"] = "Completed"

                self.interface.progress_var.set(100)
                self.interface.update_summary(1, ref_count, 0, journal_name)

                messagebox.showinfo(
                    "Processing Complete",
                    f"Processing complete!\nFile processed: {filename}\n"
                    f"Footnotes collected: {ref_count}\nMain text pages: {len(main_texts)}"
                )

            except Exception as e:
                self.interface.log_message(f"Error processing {filename}: {e}")
                result_data["status"] = f"Error: {str(e)[:50]}..."

            self.interface.update_results_table(result_data)
            self.interface.processing_results.append(result_data)

            if self.interface.processing_results:
                self.interface.save_csv_button.config(state=tk.NORMAL)

        except Exception as e:
            self.interface.log_message(f"Error during processing: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.interface.start_button.config(state=tk.NORMAL)


class EnhancedOCRInterface(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("OCR Processing Interface - Footnotes & Bibliographic Abbreviations")
        self.geometry("900x700")  # Reduced initial height

        # Use the centralized journal configuration
        self.journals = JournalConfigManager.get_journal_configs()

        # Initialize processing task manager
        self.task_manager = ProcessingTaskManager(self)

        # Store processing results for CSV report
        self.processing_results = []

        self.create_widgets()

    def create_widgets(self):
        # Create main canvas with scrollbars
        self.canvas = tk.Canvas(self)

        # Create vertical scrollbar
        v_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Configure canvas
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create main frame inside canvas
        main_frame = ttk.Frame(self.canvas, padding="10")

        # Add main frame to canvas
        self.canvas_frame = self.canvas.create_window((0, 0), window=main_frame, anchor="nw")

        # Bind events for responsive scrolling
        main_frame.bind('<Configure>', self.on_frame_configure)
        self.canvas.bind('<Configure>', self.on_canvas_configure)

        # Bind mouse wheel events for scrolling
        self.bind_mousewheel_events()

        # Store reference to main frame for widget creation
        self.main_frame = main_frame

        # Processing type selection
        type_frame = ttk.LabelFrame(self.main_frame, text="Processing Type", padding="10")
        type_frame.pack(fill=tk.X, pady=5)

        self.processing_type = tk.StringVar(value="footnotes")
        ttk.Radiobutton(type_frame, text="Extract Footnotes", variable=self.processing_type,
                        value="footnotes", command=self.toggle_processing_type).pack(side=tk.LEFT, padx=10)

        if ABBREV_AVAILABLE:
            ttk.Radiobutton(type_frame, text="Extract Bibliographic Abbreviations", variable=self.processing_type,
                            value="abbreviations", command=self.toggle_processing_type).pack(side=tk.LEFT, padx=10)
        else:
            disabled_abbrev = ttk.Radiobutton(type_frame, text="Extract Bibliographic Abbreviations (Not Available)",
                                              variable=self.processing_type, value="abbreviations", state=tk.DISABLED)
            disabled_abbrev.pack(side=tk.LEFT, padx=10)

        # Processing mode selection
        mode_frame = ttk.LabelFrame(self.main_frame, text="Processing Mode", padding="10")
        mode_frame.pack(fill=tk.X, pady=5)

        self.processing_mode = tk.StringVar(value="folder")
        ttk.Radiobutton(mode_frame, text="Process Entire Folder", variable=self.processing_mode,
                        value="folder", command=self.toggle_mode).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="Process Single File", variable=self.processing_mode,
                        value="single", command=self.toggle_mode).pack(side=tk.LEFT, padx=10)

        # Directory/File selection section
        dir_frame = ttk.LabelFrame(self.main_frame, text="File Selection", padding="10")
        dir_frame.pack(fill=tk.X, pady=5)

        # Input directory/file
        self.input_label = ttk.Label(dir_frame, text="Input Directory:")
        self.input_label.grid(row=0, column=0, sticky="w", pady=5)
        self.input_dir_var = tk.StringVar()
        self.input_entry = ttk.Entry(dir_frame, textvariable=self.input_dir_var, width=50)
        self.input_entry.grid(row=0, column=1, padx=5, pady=5)
        self.input_browse_btn = ttk.Button(dir_frame, text="Browse...", command=self.browse_input)
        self.input_browse_btn.grid(row=0, column=2, pady=5)

        # Output directory
        ttk.Label(dir_frame, text="Output Directory:").grid(row=1, column=0, sticky="w", pady=5)
        self.output_dir_var = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.output_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(dir_frame, text="Browse...", command=self.browse_output_dir).grid(row=1, column=2, pady=5)

        # Metadata directory
        self.meta_label = ttk.Label(dir_frame, text="Metadata Directory:")
        self.meta_label.grid(row=2, column=0, sticky="w", pady=5)
        self.meta_dir_var = tk.StringVar()
        self.meta_entry = ttk.Entry(dir_frame, textvariable=self.meta_dir_var, width=50)
        self.meta_entry.grid(row=2, column=1, padx=5, pady=5)
        self.meta_browse_btn = ttk.Button(dir_frame, text="Browse...", command=self.browse_meta_dir)
        self.meta_browse_btn.grid(row=2, column=2, pady=5)

        # OCR directory (for abbreviations processing)
        self.ocr_label = ttk.Label(dir_frame, text="OCR Directory:")
        self.ocr_label.grid(row=3, column=0, sticky="w", pady=5)
        self.ocr_dir_var = tk.StringVar()
        self.ocr_entry = ttk.Entry(dir_frame, textvariable=self.ocr_dir_var, width=50)
        self.ocr_entry.grid(row=3, column=1, padx=5, pady=5)
        self.ocr_browse_btn = ttk.Button(dir_frame, text="Browse...", command=self.browse_ocr_dir)
        self.ocr_browse_btn.grid(row=3, column=2, pady=5)

        # Trace directory (for abbreviations processing)
        self.trace_label = ttk.Label(dir_frame, text="Trace Directory:")
        self.trace_label.grid(row=4, column=0, sticky="w", pady=5)
        self.trace_dir_var = tk.StringVar()
        self.trace_entry = ttk.Entry(dir_frame, textvariable=self.trace_dir_var, width=50)
        self.trace_entry.grid(row=4, column=1, padx=5, pady=5)
        self.trace_browse_btn = ttk.Button(dir_frame, text="Browse...", command=self.browse_trace_dir)
        self.trace_browse_btn.grid(row=4, column=2, pady=5)

        # Parameters section
        params_frame = ttk.LabelFrame(self.main_frame, text="Parameters", padding="10")
        params_frame.pack(fill=tk.X, pady=10)

        # Journal selection dropdown
        ttk.Label(params_frame, text="Journal:").grid(row=0, column=0, sticky="w", pady=5)
        self.journal_var = tk.StringVar(value=list(self.journals.keys())[0])
        journal_combo = ttk.Combobox(params_frame, textvariable=self.journal_var, values=list(self.journals.keys()),
                                     state="readonly")
        journal_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Type selection dropdown (printed/scanned) - only for footnotes
        self.type_label = ttk.Label(params_frame, text="Type:")
        self.type_label.grid(row=1, column=0, sticky="w", pady=5)
        self.type_var = tk.StringVar(value="scanned")
        self.type_combo = ttk.Combobox(params_frame, textvariable=self.type_var, values=["printed", "scanned"],
                                       state="readonly")
        self.type_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Results table section
        results_frame = ttk.LabelFrame(self.main_frame, text="Processing Results", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create Treeview for results
        self.results_columns = ("issue", "filename", "meta_refs", "meta_labels", "collected", "status")
        self.results_tree = ttk.Treeview(results_frame, columns=self.results_columns, show="headings", height=6)

        # Define headings
        self.results_tree.heading("issue", text="Issue #")
        self.results_tree.heading("filename", text="Filename")
        self.results_tree.heading("meta_refs", text="Meta Refs")
        self.results_tree.heading("meta_labels", text="Biggest Label")
        self.results_tree.heading("collected", text="Collected")
        self.results_tree.heading("status", text="Status")

        # Configure tags for highlighting rows with significant differences
        self.results_tree.tag_configure("bold_diff", font=("TkDefaultFont", 9, "bold"),
                                        foreground="red", background="#ffe6e6")

        # Define column widths
        self.results_tree.column("issue", width=60)
        self.results_tree.column("filename", width=150)
        self.results_tree.column("meta_refs", width=80)
        self.results_tree.column("meta_labels", width=80)
        self.results_tree.column("collected", width=80)
        self.results_tree.column("status", width=100)

        # Add scrollbar for results
        results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=results_scrollbar.set)

        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Summary section
        summary_frame = ttk.LabelFrame(self.main_frame, text="Summary", padding="10")
        summary_frame.pack(fill=tk.X, pady=5)

        self.summary_text = tk.Text(summary_frame, height=3, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True)

        # Frame for progress and buttons
        control_frame = ttk.Frame(self.main_frame)
        control_frame.pack(fill=tk.X, pady=10)

        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(control_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.start_button = ttk.Button(control_frame, text="Start Processing", command=self.start_processing)
        self.start_button.pack(side=tk.RIGHT, padx=5)

        self.save_csv_button = ttk.Button(control_frame, text="Save CSV Report", command=self.save_csv_report,
                                          state=tk.DISABLED)
        self.save_csv_button.pack(side=tk.RIGHT, padx=5)

        # Logs section
        log_frame = ttk.LabelFrame(self.main_frame, text="Logs", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Create scrolled text with both scrollbars
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Initialize display
        self.toggle_processing_type()
        self.toggle_mode()

    def on_frame_configure(self, event):
        """Update canvas scroll region when frame size changes"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        """Update frame width when canvas size changes for responsive design"""
        canvas_width = event.width
        # Update the frame width to match canvas width for responsive design
        self.canvas.itemconfig(self.canvas_frame, width=canvas_width)

    def bind_mousewheel_events(self):
        """Bind mouse wheel events for scrolling"""

        # Bind to canvas and all child widgets
        def bind_to_mousewheel(widget):
            widget.bind("<MouseWheel>", self.on_mousewheel)  # Windows
            widget.bind("<Button-4>", self.on_mousewheel)  # Linux
            widget.bind("<Button-5>", self.on_mousewheel)  # Linux

            # Recursively bind to all children
            for child in widget.winfo_children():
                bind_to_mousewheel(child)

        bind_to_mousewheel(self)

    def on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        # Check if vertical scrolling is needed
        if self.canvas.cget("scrollregion"):
            if event.num == 4 or event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                self.canvas.yview_scroll(1, "units")
        return "break"

    def on_horizontal_scroll(self, event):
        """Handle horizontal scrolling with Shift+mousewheel"""
        if event.state & 0x1:  # Shift key pressed
            if event.num == 4 or event.delta > 0:
                self.canvas.xview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                self.canvas.xview_scroll(1, "units")
            return "break"

    def toggle_processing_type(self):
        """Toggle between footnotes and abbreviations processing"""
        proc_type = self.processing_type.get()
        if proc_type == "footnotes":
            # Show footnote-specific controls
            self.type_label.grid()
            self.type_combo.grid()
            # Hide abbreviation-specific controls
            self.ocr_label.grid_remove()
            self.ocr_entry.grid_remove()
            self.ocr_browse_btn.grid_remove()
            self.trace_label.grid_remove()
            self.trace_entry.grid_remove()
            self.trace_browse_btn.grid_remove()
            # Update results columns for footnotes
            self.update_results_columns_for_footnotes()
        else:  # abbreviations
            # Hide footnote-specific controls
            self.type_label.grid_remove()
            self.type_combo.grid_remove()
            # Show abbreviation-specific controls
            self.ocr_label.grid()
            self.ocr_entry.grid()
            self.ocr_browse_btn.grid()
            self.trace_label.grid()
            self.trace_entry.grid()
            self.trace_browse_btn.grid()
            # Update results columns for abbreviations
            self.update_results_columns_for_abbreviations()

    def update_results_columns_for_footnotes(self):
        """Update results table for footnote processing"""
        self.results_tree.heading("issue", text="Issue #")
        self.results_tree.heading("filename", text="Filename")
        self.results_tree.heading("meta_refs", text="Meta Refs")
        self.results_tree.heading("meta_labels", text="Biggest Label")
        self.results_tree.heading("collected", text="Collected")
        self.results_tree.heading("status", text="Status")

    def update_results_columns_for_abbreviations(self):
        """Update results table for abbreviation processing"""
        self.results_tree.heading("issue", text="Issue #")
        self.results_tree.heading("filename", text="Filename")
        self.results_tree.heading("meta_refs", text="Has Abbrev")
        self.results_tree.heading("meta_labels", text="Abbrev Count")
        self.results_tree.heading("collected", text="Pages")
        self.results_tree.heading("status", text="Status")

    def toggle_mode(self):
        """Toggle between folder and single file mode"""
        mode = self.processing_mode.get()
        if mode == "folder":
            self.input_label.config(text="Input Directory:")
            self.meta_label.grid()
            self.meta_entry.grid()
            self.meta_browse_btn.grid()
            # For abbreviations, also handle OCR and trace directories
            if self.processing_type.get() == "abbreviations":
                self.ocr_label.grid()
                self.ocr_entry.grid()
                self.ocr_browse_btn.grid()
                self.trace_label.grid()
                self.trace_entry.grid()
                self.trace_browse_btn.grid()
        else:  # single file mode
            if self.processing_type.get() == "footnotes":
                self.input_label.config(text="Input File (XLSX):")
            else:
                self.input_label.config(text="Input File (PDF):")
            self.meta_label.grid_remove()
            self.meta_entry.grid_remove()
            self.meta_browse_btn.grid_remove()
            # For abbreviations, hide OCR and trace directories in single file mode
            if self.processing_type.get() == "abbreviations":
                self.ocr_label.grid_remove()
                self.ocr_entry.grid_remove()
                self.ocr_browse_btn.grid_remove()
                self.trace_label.grid_remove()
                self.trace_entry.grid_remove()
                self.trace_browse_btn.grid_remove()

    def browse_input(self):
        """Open file dialog to select input directory or file based on mode and type."""
        mode = self.processing_mode.get()
        proc_type = self.processing_type.get()

        if mode == "folder":
            directory = filedialog.askdirectory(title="Select Input Directory")
            if directory:
                self.input_dir_var.set(directory)
        else:  # single file mode
            if proc_type == "footnotes":
                file_path = filedialog.askopenfilename(
                    title="Select XLSX File",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
            else:  # abbreviations
                file_path = filedialog.askopenfilename(
                    title="Select PDF File",
                    filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
                )
            if file_path:
                self.input_dir_var.set(file_path)

    def browse_output_dir(self):
        """Open file dialog to select output directory."""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_var.set(directory)

    def browse_meta_dir(self):
        """Open file dialog to select metadata directory."""
        directory = filedialog.askdirectory(title="Select Metadata Directory")
        if directory:
            self.meta_dir_var.set(directory)

    def browse_ocr_dir(self):
        """Open file dialog to select OCR directory."""
        directory = filedialog.askdirectory(title="Select OCR Directory")
        if directory:
            self.ocr_dir_var.set(directory)

    def browse_trace_dir(self):
        """Open file dialog to select trace directory."""
        directory = filedialog.askdirectory(title="Select Trace Directory")
        if directory:
            self.trace_dir_var.set(directory)

    def log_message(self, message):
        """Append a log message to the text widget."""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def update_results_table(self, result_data):
        """Add a row to the results table with bold formatting for files with >8% difference."""
        proc_type = self.processing_type.get()

        if proc_type == "footnotes":
            # Calculate percentage difference between collected footnotes and meta references
            collected = result_data.get("collected_footnotes", 0)
            meta_refs = result_data.get("meta_references", 0)
            biggest_refs = result_data.get("meta_labels", 0)

            # Calculate percentage difference (avoid division by zero)
            if meta_refs > 0:
                percentage_diff = abs((collected - meta_refs) / meta_refs) * 100
            elif collected > 0:
                percentage_diff = 100  # If no meta refs but we have footnotes, that's 100% difference
            else:
                percentage_diff = 0  # Both are zero

            # Insert the row
            item = self.results_tree.insert("", "end", values=(
                result_data.get("issue_number", ""),
                result_data.get("filename", ""),
                result_data.get("meta_references", "0"),
                result_data.get("meta_labels", "0"),
                result_data.get("collected_footnotes", "0"),
                result_data.get("status", "")
            ))

            # Bold the row if difference is more than 8%
            if percentage_diff > 8:
                if abs(collected - biggest_refs) > 3 or abs(collected - meta_refs) > 3:
                    self.results_tree.set(item, "status",
                                          f"{result_data.get('status', '')} ({percentage_diff:.1f}% diff)")
                    self.results_tree.item(item, tags=("bold_diff",))

        else:  # abbreviations
            # Insert the row for abbreviations
            item = self.results_tree.insert("", "end", values=(
                result_data.get("issue_number", ""),
                result_data.get("filename", ""),
                result_data.get("has_abbreviations", "No"),
                result_data.get("abbreviation_count", "0"),
                result_data.get("pages_processed", "0"),
                result_data.get("status", "")
            ))

    def update_summary(self, total_files, total_items, total_meta_refs, journal_name):
        """Update the summary section."""
        proc_type = self.processing_type.get()
        self.summary_text.delete(1.0, tk.END)

        if proc_type == "footnotes":
            summary = f"""Journal: {journal_name}
Total Files Processed: {total_files}
Total Footnotes Collected: {total_items}
Total Meta References: {total_meta_refs}
Average per File: {total_items / max(total_files, 1):.1f} footnotes, {total_meta_refs / max(total_files, 1):.1f} meta refs"""
        else:  # abbreviations
            summary = f"""Journal: {journal_name}
Total Files Processed: {total_files}
Total Abbreviations Collected: {total_items}
Files with Abbreviations: {total_meta_refs}
Average per File: {total_items / max(total_files, 1):.1f} abbreviations"""

        self.summary_text.insert(tk.END, summary)

    def start_processing(self):
        """Validate inputs and start processing in a background thread."""
        mode = self.processing_mode.get()
        proc_type = self.processing_type.get()
        input_path = self.input_dir_var.get()
        output_dir = self.output_dir_var.get()

        # Basic validation
        if not input_path or not output_dir:
            messagebox.showerror("Error", "Please select input and output directories/files")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Error", f"Input path does not exist: {input_path}")
            return

        # Validate based on mode and type
        if mode == "folder":
            if proc_type == "footnotes":
                meta_dir = self.meta_dir_var.get()
                if not meta_dir or not os.path.exists(meta_dir):
                    messagebox.showerror("Error", "Please select a valid metadata directory")
                    return
            else:  # abbreviations
                ocr_dir = self.ocr_dir_var.get()
                trace_dir = self.trace_dir_var.get()
                if not ocr_dir or not trace_dir:
                    messagebox.showerror("Error", "Please select OCR and trace directories for abbreviation processing")
                    return
                if not os.path.exists(ocr_dir):
                    messagebox.showerror("Error", f"OCR directory does not exist: {ocr_dir}")
                    return
        else:  # single file mode
            if proc_type == "footnotes" and not input_path.lower().endswith('.xlsx'):
                messagebox.showerror("Error", "Please select an XLSX file for footnote processing")
                return
            elif proc_type == "abbreviations" and not input_path.lower().endswith('.pdf'):
                messagebox.showerror("Error", "Please select a PDF file for abbreviation processing")
                return

        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_message(f"Created output directory: {output_dir}")

        journal = self.journal_var.get()
        doc_type = self.type_var.get() if proc_type == "footnotes" else None

        # Validate journal and document type for footnotes
        if proc_type == "footnotes":
            if journal not in self.journals or doc_type not in self.journals[journal]:
                messagebox.showerror("Error", "Invalid journal or type selection")
                return

        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.processing_results = []

        # Disable buttons and reset progress
        self.start_button.config(state=tk.DISABLED)
        self.save_csv_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        # Start processing in a separate thread
        if proc_type == "footnotes":
            if mode == "folder":
                meta_dir = self.meta_dir_var.get()
                thread = threading.Thread(
                    target=self.task_manager.process_footnotes_folder,
                    args=(input_path, output_dir, meta_dir, journal, doc_type)
                )
            else:  # single file
                thread = threading.Thread(
                    target=self.task_manager.process_footnotes_single,
                    args=(input_path, output_dir, journal, doc_type)
                )
        else:  # abbreviations
            if mode == "folder":
                # For abbreviations folder processing, implement similar to footnotes
                meta_dir = self.meta_dir_var.get()
                ocr_dir = self.ocr_dir_var.get()
                trace_dir = self.trace_dir_var.get()
                # This would need to be implemented in ProcessingTaskManager
                messagebox.showinfo("Info", "Abbreviation processing not fully implemented in this version")
                self.start_button.config(state=tk.NORMAL)
                return
            else:  # single file abbreviations
                messagebox.showinfo("Info", "Single file abbreviation processing requires folder mode with OCR data")
                self.start_button.config(state=tk.NORMAL)
                return

        thread.daemon = True
        thread.start()

    def save_csv_report(self):
        """Save the processing results to a CSV file in the output directory."""
        if not self.processing_results:
            messagebox.showwarning("No Data", "No processing results to save")
            return

        # Get output directory and journal name
        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showerror("Error", "Output directory not specified")
            return

        journal_name = extract_journal_name_from_path(self.input_dir_var.get())
        proc_type = self.processing_type.get()

        # Create filename and full path
        filename = f"{journal_name}_{proc_type}_processing_report.csv"
        file_path = os.path.join(output_dir, filename)

        try:
            if proc_type == "footnotes":
                # Define CSV headers for footnotes
                headers = [
                    "Issue_Number",
                    "Filename",
                    "Meta_References_Count",
                    "Meta_Biggest_Label_Number",
                    "Collected_Footnotes_Count",
                    "Has_Meta_File",
                    "Processing_Status"
                ]

                with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=headers)
                    writer.writeheader()

                    for result in self.processing_results:
                        row = {
                            "Issue_Number": result.get("issue_number", ""),
                            "Filename": result.get("filename", ""),
                            "Meta_References_Count": result.get("meta_references", 0),
                            "Meta_Biggest_Label_Number": result.get("meta_labels", 0),
                            "Collected_Footnotes_Count": result.get("collected_footnotes", 0),
                            "Has_Meta_File": result.get("has_meta_file", False),
                            "Processing_Status": result.get("status", "")
                        }
                        writer.writerow(row)

            else:  # abbreviations
                # Define CSV headers for abbreviations
                headers = [
                    "Issue_Number",
                    "Filename",
                    "Has_Abbreviations",
                    "Abbreviation_Count",
                    "Pages_Processed",
                    "Processing_Status"
                ]

                with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=headers)
                    writer.writeheader()

                    for result in self.processing_results:
                        row = {
                            "Issue_Number": result.get("issue_number", ""),
                            "Filename": result.get("filename", ""),
                            "Has_Abbreviations": result.get("has_abbreviations", "No"),
                            "Abbreviation_Count": result.get("abbreviation_count", 0),
                            "Pages_Processed": result.get("pages_processed", 0),
                            "Processing_Status": result.get("status", "")
                        }
                        writer.writerow(row)

            self.log_message(f"CSV report saved to: {file_path}")
            messagebox.showinfo("Success", f"Report saved successfully to:\n{file_path}")

        except Exception as e:
            self.log_message(f"Error saving CSV report: {e}")
            messagebox.showerror("Error", f"Failed to save report: {e}")


if __name__ == "__main__":
    app = EnhancedOCRInterface()
    app.mainloop()