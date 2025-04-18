import fitz  # PyMuPDF
import os
import shutil
import uuid
import zipfile
import sys
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image, ImageTk
from xml.sax.saxutils import escape as xml_escape
import traceback # For detailed error logging

# --- Constants ---
TEMP_DIR_PREFIX = "pdf2epub_py_"

# --- NEW Modern GUI Colors ---
BG_COLOR = "#F0F0F0"          # Light grey background
FRAME_BG = "#FFFFFF"          # White frame background
ACCENT_COLOR = "#0078D4"     # Modern blue accent
ACCENT_ACTIVE_COLOR = "#005A9E" # Darker blue for active/hover
BUTTON_FG = "#FFFFFF"          # White button text
TEXT_COLOR = "#202020"         # Dark grey text
SECONDARY_TEXT_COLOR = "#606060" # Lighter grey text (e.g., for hints)
BORDER_COLOR = "#D0D0D0"      # Light grey border
ENTRY_BG = "#FFFFFF"          # White entry background
DISABLED_BG = "#EAEAEA"      # Background for readonly/disabled fields
DND_BG = "#E8F0FE"          # Light blue for Drag and Drop hint
PROGRESS_BG = ACCENT_COLOR    # Progress bar color
PROGRESS_TROUGH = "#DCDCDC"  # Progress bar trough color

# --- EPUB Structure Templates ---
CONTAINER_XML_CONTENT = """<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>"""

CSS_CONTENT = """body { margin: 0; padding: 0; }
.page-svg-container {
    width: 100vw;
    height: 100vh;
    display: flex;
    justify-content: center;
    align-items: center;
    overflow: hidden;
}
svg { display: block; width: 100%; height: 100%; }
image { width: 100%; height: 100%; object-fit: contain; }
"""

# --- Helper Functions for EPUB Generation (Restored) ---

def create_content_opf(title, image_files, page_dimensions):
    """Generates the content.opf XML string."""
    book_uuid = uuid.uuid4()
    now = fitz.get_pdf_now() # Get timestamp in PDF format
    manifest_items = [
        f'    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
        f'    <item id="css" href="css/styles.css" media-type="text/css"/>'
    ]
    spine_items = []

    for i, img_file in enumerate(image_files):
        page_num = i + 1
        page_id, image_id = f"page{page_num}", f"img{page_num}"
        xhtml_href, image_href = f"xhtml/{page_id}.xhtml", f"images/{img_file}"
        manifest_items.extend([
            f'    <item id="{page_id}" href="{xhtml_href}" media-type="application/xhtml+xml"/>',
            f'    <item id="{image_id}" href="{image_href}" media-type="image/png"/>' # Assuming PNG
        ])
        spine_items.append(f'    <itemref idref="{page_id}" properties="page-spread-left rendition:layout-pre-paginated rendition:orientation-auto rendition:spread-auto"/>')

    manifest_str = "\n".join(manifest_items)
    spine_str = "\n".join(spine_items)

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="BookId" version="3.0" prefix="rendition: http://www.idpf.org/vocab/rendition/#">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="BookId">urn:uuid:{book_uuid}</dc:identifier>
    <dc:title>{xml_escape(title)}</dc:title>
    <dc:language>en</dc:language>
    <meta property="dcterms:modified">{now}</meta>
    <meta name="cover" content="img1"/>
    <meta property="rendition:layout">pre-paginated</meta>
    <meta property="rendition:orientation">auto</meta>
    <meta property="rendition:spread">auto</meta>
  </metadata>
  <manifest>
{manifest_str}
  </manifest>
  <spine toc="nav">
{spine_str}
  </spine>
</package>"""

def create_nav_xhtml(title, image_files):
    """Generates the nav.xhtml (EPUB3 ToC/Page List) XML string."""
    toc_list_items, page_list_items = [], []
    for i, _ in enumerate(image_files):
        page_num = i + 1
        xhtml_href = f"xhtml/page{page_num}.xhtml"
        toc_list_items.append(f'      <li><a href="{xhtml_href}">Page {page_num}</a></li>')
        page_list_items.append(f'      <li><a href="{xhtml_href}">{page_num}</a></li>')

    toc_list_str = "\n".join(toc_list_items)
    page_list_str = "\n".join(page_list_items)

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"><head>
<title>{xml_escape(title)} - Contents</title><link rel="stylesheet" type="text/css" href="css/styles.css" /></head><body>
<nav epub:type="toc" id="toc"><h1>Table of Contents</h1><ol>{toc_list_str}</ol></nav>
<nav epub:type="page-list" id="page-list" hidden=""><h1>Page List</h1><ol>{page_list_str}</ol></nav>
</body></html>"""

def create_page_xhtml(page_num, img_file, dimensions):
    """Generates an XHTML string for a single EPUB page, wrapping the image."""
    if not dimensions: dimensions = {"width": 600, "height": 800}
    vp_width, vp_height = int(dimensions["width"]), int(dimensions["height"])
    image_href = f"../images/{img_file}"

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="en" lang="en"><head>
<title>Page {page_num}</title><meta charset="UTF-8"/><meta name="viewport" content="width={vp_width}, height={vp_height}"/>
<link rel="stylesheet" type="text/css" href="../css/styles.css"/></head><body>
<div class="page-svg-container"><svg xmlns="http://www.w3.org/2000/svg" version="1.1" width="{vp_width}" height="{vp_height}" viewBox="0 0 {vp_width} {vp_height}" preserveAspectRatio="xMidYMid meet">
<image width="{vp_width}" height="{vp_height}" xlink:href="{image_href}" xmlns:xlink="http://www.w3.org/1999/xlink"/></svg></div></body></html>"""

def write_to_file(filepath, content):
    """Helper to write content to a file using UTF-8 encoding."""
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
    except IOError as e:
        raise IOError(f"Failed to write to {filepath}: {e}") from e

# --- Core Conversion Logic (Worker Thread - Restored) ---

def pdf_to_epub_fxl_core(pdf_path, dpi, status_queue, output_dir=None):
    """Performs the PDF to EPUB conversion for a single file."""
    abs_pdf_path = os.path.abspath(pdf_path)
    pdf_basename = os.path.basename(abs_pdf_path)
    pdf_title = os.path.splitext(pdf_basename)[0].replace("_", " ")
    epub_filename = f"{os.path.splitext(pdf_basename)[0]}.epub"

    if output_dir is None:
        output_path = os.path.join(os.path.dirname(abs_pdf_path), epub_filename)
    else:
        output_path = os.path.join(output_dir, epub_filename)

    temp_dir = None
    try:
        temp_dir = os.path.join(os.path.dirname(output_path), TEMP_DIR_PREFIX + uuid.uuid4().hex)
        build_dir = os.path.join(temp_dir, "epub_build")
        raw_image_dir = os.path.join(temp_dir, "images_raw")
        os.makedirs(build_dir, exist_ok=True)
        os.makedirs(raw_image_dir, exist_ok=True)

        oebps_dir = os.path.join(build_dir, "OEBPS")
        meta_inf_dir = os.path.join(build_dir, "META-INF")
        images_dest_dir = os.path.join(oebps_dir, "images")
        xhtml_dir = os.path.join(oebps_dir, "xhtml")
        css_dir = os.path.join(oebps_dir, "css")
        for d in [oebps_dir, meta_inf_dir, images_dest_dir, xhtml_dir, css_dir]:
            os.makedirs(d, exist_ok=True)

        doc = fitz.open(abs_pdf_path)
        image_files = []
        page_dimensions = []
        total_pages = len(doc)
        status_queue.put(f"Processing {pdf_basename}: {total_pages} pages...")

        for i, page in enumerate(doc):
            page_num = i + 1
            # Minimal progress update to queue to avoid flooding
            if (i + 1) % 10 == 0 or (i + 1) == total_pages:
                 status_queue.put(f"  -> Rendering page {page_num}/{total_pages}...")
            rect = page.rect
            page_dimensions.append({"width": rect.width, "height": rect.height})
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_filename = f"page-{page_num}.png"
            img_raw_path = os.path.join(raw_image_dir, img_filename)
            pix.save(img_raw_path)
            image_files.append(img_filename)
        doc.close()
        status_queue.put(f"  -> Rendered {len(image_files)} pages.")
        if not image_files: raise RuntimeError("No images generated from PDF.")

        status_queue.put("  -> Generating EPUB structure...")
        write_to_file(os.path.join(build_dir, "mimetype"), "application/epub+zip")
        write_to_file(os.path.join(meta_inf_dir, "container.xml"), CONTAINER_XML_CONTENT)
        write_to_file(os.path.join(oebps_dir, "content.opf"), create_content_opf(pdf_title, image_files, page_dimensions))
        write_to_file(os.path.join(oebps_dir, "nav.xhtml"), create_nav_xhtml(pdf_title, image_files))
        write_to_file(os.path.join(css_dir, "styles.css"), CSS_CONTENT)

        status_queue.put("  -> Creating page files and copying images...")
        for i, img_filename in enumerate(image_files):
            page_num = i + 1
            xhtml_content = create_page_xhtml(page_num, img_filename, page_dimensions[i])
            write_to_file(os.path.join(xhtml_dir, f"page{page_num}.xhtml"), xhtml_content)
            shutil.copy2(os.path.join(raw_image_dir, img_filename), os.path.join(images_dest_dir, img_filename))

        status_queue.put(f"  -> Creating EPUB archive: {output_path}")
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as epub_zip:
            epub_zip.write(os.path.join(build_dir, "mimetype"), "mimetype", compress_type=zipfile.ZIP_STORED)
            for root, _, files in os.walk(build_dir):
                if os.path.basename(root) == os.path.basename(build_dir) and "mimetype" in files:
                    files.remove("mimetype")
                for file in files:
                    file_path = os.path.join(root, file)
                    archive_name = os.path.relpath(file_path, build_dir)
                    epub_zip.write(file_path, archive_name, compress_type=zipfile.ZIP_DEFLATED)

        status_queue.put("DONE_FILE")
        print(f"ERROR: Worker thread for {pdf_basename} finished successfully.")

    except Exception as e:
        error_traceback = traceback.format_exc()
        status_queue.put(f"\n❌ ERROR converting {pdf_basename}:")
        status_queue.put(f"   {type(e).__name__}: {e}")
        status_queue.put("--- Error Details ---")
        status_queue.put(error_traceback)
        status_queue.put("---------------------")
        status_queue.put("ERROR_FILE")
        # Keep console log for fatal errors in worker
        print(f"ERROR: Worker thread for {pdf_basename} hit error:\n{error_traceback}")

    finally:
        if temp_dir and os.path.exists(temp_dir):
            try:
                # print(f"DEBUG: Attempting cleanup of {temp_dir}") # Keep commented out or remove
                shutil.rmtree(temp_dir)
                # print(f"DEBUG: Cleanup successful for {temp_dir}")
            except Exception as cleanup_error:
                # Keep error messages for cleanup failures
                status_queue.put(f"⚠️ Error cleaning up temp dir: {cleanup_error}")
                print(f"ERROR: Cleanup FAILED for {temp_dir}: {cleanup_error}")
        # else:
            # print(f"DEBUG: No temp dir cleanup needed for {pdf_basename}.") # Remove

# --- GUI Application ---

class PdfToEpubApp(TkinterDnD.Tk):
    """Main application window for the PDF to EPUB FXL converter."""
    def __init__(self):
        """Initializes the application window, styles, variables, and widgets."""
        super().__init__()
        self.title("epub please! - PDF to EPUB FXL Converter")
        self.geometry("700x700")
        self.resizable(True, True)
        self.config(bg=BG_COLOR)

        # Configure resizing behavior
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0) # Logo row - fixed size
        self.rowconfigure(1, weight=2) # Input list gets more weight
        self.rowconfigure(6, weight=3) # Log area gets less weight than input
        # Other rows (2, 3, 4, 5) have default weight 0 (fixed size)

        self._configure_styles()
        self._initialize_variables()
        self._set_window_icon()
        self._create_widgets()
        self._configure_drag_drop()
        self.check_status_queue()

        # Force update after widgets are created
        self.update_idletasks()

    def _configure_styles(self):
        """Configures ttk styles with the modern theme."""
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # --- General Styles ---
        self.style.configure(".",
                           background=BG_COLOR,
                           foreground=TEXT_COLOR,
                           font=("Segoe UI", 9)) # Use Segoe UI if available
        self.style.configure("TFrame", background=BG_COLOR)
        self.style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
        self.style.configure("Accent.TLabel", foreground=ACCENT_COLOR) # Example accent label

        # --- Button Styles ---
        self.style.configure("TButton",
                           background=ACCENT_COLOR,
                           foreground=BUTTON_FG,
                           borderwidth=0,
                           padding=(10, 5),
                           font=("Segoe UI", 10, "bold"),
                           focuscolor=ACCENT_COLOR) # Keep focus color same
        self.style.map("TButton",
                       background=[("active", ACCENT_ACTIVE_COLOR), ("hover", ACCENT_ACTIVE_COLOR)],
                       foreground=[("active", BUTTON_FG)])

        # --- Entry Styles ---
        self.style.configure("TEntry",
                           fieldbackground=ENTRY_BG,
                           foreground=TEXT_COLOR,
                           borderwidth=1,
                           relief="solid",
                           bordercolor=BORDER_COLOR,
                           insertcolor=TEXT_COLOR, # Cursor color
                           padding=(5, 5))
        self.style.map("TEntry", bordercolor=[("focus", ACCENT_COLOR)])
        # Readonly entry style
        self.style.configure("Readonly.TEntry",
                           fieldbackground=DISABLED_BG,
                           foreground=SECONDARY_TEXT_COLOR,
                           borderwidth=1,
                           relief="solid",
                           bordercolor=BORDER_COLOR)

        # --- LabelFrame Styles ---
        self.style.configure("TLabelframe",
                           background=BG_COLOR,
                           borderwidth=1,
                           relief="solid",
                           bordercolor=BORDER_COLOR,
                           padding=(10, 10))
        self.style.configure("TLabelframe.Label",
                           background=BG_COLOR,
                           foreground=TEXT_COLOR,
                           font=("Segoe UI", 10, "bold"),
                           padding=(0, 0, 0, 5)) # Padding below label

        # --- Progress Bar Style ---
        self.style.configure("custom.Horizontal.TProgressbar",
                           troughcolor=PROGRESS_TROUGH,
                           background=PROGRESS_BG,
                           thickness=15, # Make it thicker
                           borderwidth=0,
                           relief="flat")

        # Add Tag for placeholder text in ScrolledText
        # Note: ScrolledText doesn't directly support ttk styles for tags,
        # so we configure the tag on the underlying tk.Text widget later.
        pass # Placeholder, tag configured in update_file_list_display

    def _initialize_variables(self):
        """Initializes Tkinter variables and application state flags."""
        self.output_dir_path = tk.StringVar()
        self.pdf_file_list = []
        self.current_conversion_index = -1
        self.dpi_var = tk.StringVar(value="150")
        self.status_queue = queue.Queue()
        self.is_converting = False
        self.logo_image = None
        self.icon_image = None
        # Widgets initialized in _create_widgets
        self.log_display_text = None

    def _set_window_icon(self):
        """Loads and sets the application window icon."""
        try:
            # Assumes icon.webp is in 'assets' subdirectory relative to the script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "assets", "icon.webp")
            img = Image.open(icon_path)
            self.icon_image = ImageTk.PhotoImage(img)
            # Set icon for this window and potentially future top-levels
            self.iconphoto(True, self.icon_image)
        except Exception as e:
            # Log warning if icon fails to load, but don't crash the app
            print(f"Warning: Could not set window icon - {e}")

    def _create_widgets(self):
        """Creates and arranges widgets with modern styling and padding."""

        # --- Logo (Row 0) - Placed directly on grid with explicit resize ---
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(script_dir, "assets", "icon.webp")
            # Load and resize specifically for the label display
            img = Image.open(logo_path)
            logo_display_size = (80, 80) # Smaller size for the label
            img.thumbnail(logo_display_size, Image.Resampling.LANCZOS)
            self.logo_image_display = ImageTk.PhotoImage(img) # Store reference

            logo_label = tk.Label(self, image=self.logo_image_display, bg=BG_COLOR)
            logo_label.grid(row=0, column=0, pady=(15, 5)) # Adjusted padding

        except FileNotFoundError:
            error_logo_label = ttk.Label(self, text="[Logo Not Found]", foreground="red", background=BG_COLOR)
            error_logo_label.grid(row=0, column=0, pady=(15,5))
        except Exception as e:
            error_logo_label = ttk.Label(self, text=f"[Logo Error: {e}]", foreground="red", background=BG_COLOR)
            error_logo_label.grid(row=0, column=0, pady=(15,5))

        # --- Input Files Frame (Row 1) ---
        input_frame = ttk.LabelFrame(self, text="Input PDF Files", padding=(15, 10))
        input_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(0, weight=1) # Text area row expands

        self.file_list_text = scrolledtext.ScrolledText(
            input_frame, height=8, width=70, wrap=tk.WORD,
            state="disabled", relief="flat", borderwidth=1,
            background=FRAME_BG, fg=TEXT_COLOR, font=("Segoe UI", 9),
            bd=1, highlightthickness=1, highlightcolor=BORDER_COLOR, highlightbackground=BORDER_COLOR
        )
        # Access the underlying tk.Text widget to configure tags
        self.file_list_text.tag_configure("placeholder", foreground=SECONDARY_TEXT_COLOR, justify="center", font=("Segoe UI", 9, "italic"))
        self.file_list_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew") # Spans the frame

        # --- Input Buttons (Row 2) ---
        input_button_frame = ttk.Frame(self, padding=(0, 5, 0, 10)) # Adjusted padding slightly
        input_button_frame.grid(row=2, column=0)
        self.browse_button = ttk.Button(input_button_frame, text="Add Files...", command=self.browse_input_pdfs)
        self.browse_button.pack(side=tk.LEFT, padx=10)
        self.clear_button = ttk.Button(input_button_frame, text="Clear List", command=self.clear_file_list)
        self.clear_button.pack(side=tk.LEFT, padx=10)

        # --- Options & Output Frame (Row 3) ---
        options_output_outer_frame = ttk.Frame(self, padding=(0, 0, 0, 10))
        options_output_outer_frame.grid(row=3, column=0, padx=20, sticky="ew")
        options_output_outer_frame.columnconfigure(0, weight=1)
        options_output_outer_frame.columnconfigure(1, weight=2) # Give output more space

        options_frame = ttk.LabelFrame(options_output_outer_frame, text="Options", padding=(15, 10))
        options_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        ttk.Label(options_frame, text="Rendering DPI:").pack(side=tk.LEFT, padx=(0, 5), pady=5)
        ttk.Entry(options_frame, textvariable=self.dpi_var, width=6).pack(side=tk.LEFT, pady=5)

        output_frame = ttk.LabelFrame(options_output_outer_frame, text="Output Directory (Optional)", padding=(15, 10))
        output_frame.grid(row=0, column=1, padx=(10, 0), sticky="nsew")
        output_frame.columnconfigure(0, weight=1)

        self.output_dir_entry = ttk.Entry(output_frame, textvariable=self.output_dir_path, state="readonly", style="Readonly.TEntry")
        self.output_dir_entry.grid(row=0, column=0, padx=(0, 10), pady=5, sticky="ew")
        browse_output_btn = ttk.Button(output_frame, text="Browse...", command=self.browse_output_dir)
        browse_output_btn.grid(row=0, column=1, pady=5)
        ttk.Label(output_frame, text="(Leave blank to save next to original PDFs)", foreground=SECONDARY_TEXT_COLOR, font=("Segoe UI", 8)).grid(row=1, column=0, columnspan=2, sticky="w", pady=(5,0))

        # --- Convert Button (Row 4) ---
        button_frame = ttk.Frame(self, padding=(0, 15, 0, 15))
        button_frame.grid(row=4, column=0)
        self.convert_button = ttk.Button(button_frame, text="Convert All to EPUB", command=self.start_batch_conversion, width=20)
        self.convert_button.pack()

        # --- Progress Bar (Row 5) ---
        progress_frame = ttk.Frame(self, padding=(0, 0, 0, 10))
        progress_frame.grid(row=5, column=0, padx=20, sticky="ew")
        progress_frame.columnconfigure(0, weight=1)
        self.progress_var = tk.DoubleVar()
        self.progressbar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", variable=self.progress_var, style="custom.Horizontal.TProgressbar")
        self.progressbar.grid(row=0, column=0, sticky="ew", ipady=2) # ipady for internal padding

        # --- Log Display Area (Row 6) - Now taller ---
        log_frame = ttk.LabelFrame(self, text="Logs & Status", padding=(15, 10))
        log_frame.grid(row=6, column=0, padx=20, pady=(10, 15), sticky="nsew") # Increase bottom pady
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_display_text = scrolledtext.ScrolledText(
            log_frame, height=20, width=70, wrap=tk.WORD, # Increased height to 20
            state="disabled", relief="flat", borderwidth=1,
            background=FRAME_BG, fg=TEXT_COLOR, font=("Segoe UI", 9),
            bd=1, highlightthickness=1, highlightcolor=BORDER_COLOR, highlightbackground=BORDER_COLOR
        )
        # Reconfigure tag here as well, just in case
        self.log_display_text.tag_configure("placeholder", foreground=SECONDARY_TEXT_COLOR, justify="center", font=("Segoe UI", 9, "italic"))
        self.log_display_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

    def _configure_drag_drop(self):
        """Registers the drop target and binds the drop event."""
        # Register drop target on the MAIN window
        self.drop_target_register(DND_FILES)
        # Bind event to the MAIN window
        self.dnd_bind("<<Drop>>", lambda event: self.handle_drop(event))
        # Removed bindings for self.file_list_text
        # self.file_list_text.drop_target_register(DND_FILES)
        # self.file_list_text.dnd_bind("<<Drop>>", lambda event: self.handle_drop(event))

    def handle_drop(self, event):
        """Handles the <<Drop>> event, parsing multiple PDF file paths with improved logic."""
        filepaths_str = event.data
        # print(f"DEBUG: Raw drop data (event.data): \n{filepaths_str}\n") # Debug line removed for production

        paths = []
        try:
            potential_paths = self.tk.splitlist(filepaths_str)
            paths.extend(potential_paths)
            # print(f"DEBUG: Parsed paths via splitlist: {paths}") # Debug line removed
        except tk.TclError:
            # print("DEBUG: tk.splitlist failed, falling back to basic split.") # Debug line removed
            paths = filepaths_str.split()

        if not paths:
            # print("DEBUG: No paths extracted from drop data.") # Debug line removed
            messagebox.showerror("Invalid Drop", "Could not extract any file paths from the dropped items.")
            return

        added_count = 0
        skipped_count = 0
        newly_added_paths = []

        for p in paths:
            cleaned_path = p.strip("{} ")
            abs_path = os.path.abspath(cleaned_path)
            # print(f"DEBUG: Checking path: \nOriginal: \n{p}\nCleaned: \n{cleaned_path}\nAbsolute: \n{abs_path}\n") # Debug line removed

            if os.path.isfile(abs_path) and abs_path.lower().endswith(".pdf"):
                if abs_path not in self.pdf_file_list:
                    self.pdf_file_list.append(abs_path)
                    newly_added_paths.append(abs_path)
                    added_count += 1
                    # print(f"DEBUG: Added valid PDF: {abs_path}") # Debug line removed
                else:
                    # print(f"DEBUG: PDF already in list: {abs_path}") # Debug line removed
                    skipped_count += 1
            else:
                skipped_count += 1
                if os.path.exists(abs_path):
                     self.update_status(f"Skipped (not a PDF or already listed): {os.path.basename(abs_path)}")
                # print(f"DEBUG: Skipped invalid/non-PDF path: {abs_path}") # Debug line removed

        if added_count > 0:
            self.update_file_list_display()
            self.update_status(f"Added {added_count} PDF file(s) via drop.")
            if skipped_count > 0:
                 self.update_status(f"Skipped {skipped_count} other item(s).")
        elif skipped_count > 0:
             messagebox.showwarning("Drop Info", f"Skipped {skipped_count} item(s) (not valid PDFs or already in list).")
             self.update_status(f"Skipped {skipped_count} item(s) during drop.")
        else:
            messagebox.showerror("Invalid Drop", "No valid new PDF files found in the dropped items.")

    def browse_input_pdfs(self):
        """Opens a file dialog to select multiple input PDF files."""
        # Ensure askopenfilenames is used and its result (tuple) is handled
        filepaths_tuple = filedialog.askopenfilenames(
            title="Select PDF File(s) to Add",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        added_count = 0
        skipped_count = 0
        if filepaths_tuple: # Check if the tuple is not empty
            # print(f"DEBUG: Browse selected: {filepaths_tuple}") # Remove
            for p in filepaths_tuple:
                abs_path = os.path.abspath(p)
                if os.path.isfile(abs_path) and abs_path.lower().endswith(".pdf"):
                     if abs_path not in self.pdf_file_list:
                         self.pdf_file_list.append(abs_path)
                         added_count += 1
                     else:
                          skipped_count +=1
                else:
                     skipped_count +=1 # Non-PDF file selected somehow?

            if added_count > 0:
                self.update_file_list_display()
                self.update_status(f"Added {added_count} PDF file(s).")
                if skipped_count > 0:
                     self.update_status(f"Skipped {skipped_count} selected item(s) (already listed or invalid).")
            elif skipped_count > 0:
                 self.update_status(f"Skipped {skipped_count} selected item(s) (already listed or invalid).")
            else:
                self.update_status("No new valid PDF files were added from selection.")
        else:
             # print("DEBUG: Browse cancelled or no files selected.") # Remove
             pass

    def browse_output_dir(self):
        """Opens a directory selection dialog for the output directory."""
        dir_path = filedialog.askdirectory(title="Select Output Directory")
        if dir_path:
            self.output_dir_path.set(dir_path)
            self.update_status(f"Selected output directory: {dir_path}")

    def update_status(self, message):
        """Appends a message to the main log/status text area."""
        target_widget = self.log_display_text

        if target_widget and target_widget.winfo_exists():
            target_widget.config(state="normal")
            target_widget.insert(tk.END, message + "\n")
            target_widget.see(tk.END)
            target_widget.config(state="disabled")
            self.update_idletasks()

    def check_status_queue(self):
        """Periodically checks the status queue for messages from the worker thread."""
        try:
            while True:
                message = self.status_queue.get_nowait()
                if message == "DONE_FILE":
                    # print("DEBUG: GUI received DONE_FILE") # Remove
                    completed_count = self.current_conversion_index + 1
                    if self.pdf_file_list:
                        progress = (completed_count / len(self.pdf_file_list)) * 100
                        self.progress_var.set(progress)
                    self.update_status(f"✅ File {completed_count} completed.")
                    self._start_next_conversion()
                elif message == "ERROR_FILE":
                     # print("DEBUG: GUI received ERROR_FILE") # Remove
                     completed_count = self.current_conversion_index + 1
                     if self.pdf_file_list:
                         progress = (completed_count / len(self.pdf_file_list)) * 100
                         self.progress_var.set(progress)
                     self.update_status(f"❌ Error processing file {completed_count}. See details above.")
                     self._start_next_conversion()
                else:
                    self.update_status(message)
        except queue.Empty:
            pass
        except Exception as e:
            error_msg = f"Error processing status queue: {e}"
            print(error_msg)
            self.update_status(f"GUI ERROR: {error_msg}")
        self.after(100, self.check_status_queue)

    def _start_next_conversion(self):
        """Starts the conversion thread for the next file in the list."""
        self.current_conversion_index += 1

        if self.current_conversion_index < len(self.pdf_file_list):
            pdf_path = self.pdf_file_list[self.current_conversion_index]
            filename = os.path.basename(pdf_path)
            self.update_status(f"\n[{self.current_conversion_index + 1}/{len(self.pdf_file_list)}] Starting: {filename}...")
            try:
                dpi = int(self.dpi_var.get())
                if dpi <= 0: raise ValueError("DPI must be positive")
            except ValueError:
                 self.update_status(f"Invalid DPI '{self.dpi_var.get()}' entered for {filename}. Using default 150.")
                 self.update_status(f"Using default DPI 150 for {filename}.")
                 dpi = 150

            output_dir = self.output_dir_path.get() or None

            conversion_thread = threading.Thread(
                target=pdf_to_epub_fxl_core,
                kwargs={ # Pass kwargs explicitly
                    'pdf_path': pdf_path,
                    'dpi': dpi,
                    'status_queue': self.status_queue,
                    'output_dir': output_dir
                    },
                daemon=True
            )
            conversion_thread.start()
        else:
            self.stop_batch_conversion("✅ Batch conversion finished.")

    def stop_batch_conversion(self, final_message="Conversion stopped."):
        """Handles UI changes when batch stops (completed or error)."""
        self.update_status(f"\n--- {final_message} ---")
        self.is_converting = False
        self.convert_button.config(state="normal")
        self.browse_button.config(state="normal")
        self.clear_button.config(state="normal")
        if "finished" in final_message.lower() or "completed" in final_message.lower():
            self.progress_var.set(100.0)
            self.progressbar['value'] = 100
        self.update_idletasks()

    def start_batch_conversion(self):
        """Validates inputs and starts the batch conversion process."""
        if self.is_converting:
            messagebox.showwarning("Busy", "Conversion already in progress.")
            return
        if not self.pdf_file_list:
             messagebox.showerror("Input Error", "Please add PDF files to the list first.")
             return

        try:
            dpi = int(self.dpi_var.get())
            if dpi <= 0: raise ValueError("DPI must be positive")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid positive integer for DPI (e.g., 150).")
            return

        output_dir = self.output_dir_path.get() or None
        if output_dir and not os.path.isdir(output_dir):
             messagebox.showerror("Output Error", f"Selected output directory does not exist:\n{output_dir}")
             return

        # Clear the single log area
        if self.log_display_text and self.log_display_text.winfo_exists():
             self.log_display_text.config(state="normal")
             self.log_display_text.delete(1.0, tk.END)
             self.log_display_text.config(state="disabled")

        self.update_status(f"Starting batch conversion for {len(self.pdf_file_list)} file(s)...")
        output_dir = self.output_dir_path.get() or None
        if output_dir:
            self.update_status(f"Output directory: {output_dir}")
        else:
            self.update_status("Output: Saving EPUBs next to original PDFs.")

        self.is_converting = True
        self.convert_button.config(state="disabled")
        self.browse_button.config(state="disabled")
        self.clear_button.config(state="disabled")
        self.progress_var.set(0.0)
        self.progressbar['value'] = 0
        self.progressbar['maximum'] = 100
        self.update_idletasks()

        self.current_conversion_index = -1
        self._start_next_conversion()

    def update_file_list_display(self):
        """Updates the text area showing the list of files or a placeholder."""
        self.file_list_text.config(state="normal")
        self.file_list_text.delete(1.0, tk.END)
        self.file_list_text.tag_remove("placeholder", "1.0", tk.END)

        if self.pdf_file_list:
            display_text = "\n".join([f"{i+1}. {os.path.basename(f_path)} ({os.path.dirname(f_path)})" for i, f_path in enumerate(self.pdf_file_list)])
            self.file_list_text.insert(tk.END, display_text)
        else:
            # Update placeholder text and center it
            placeholder = "\n\nDrag & Drop PDF's"
            self.file_list_text.insert(tk.END, placeholder, "placeholder")

        self.file_list_text.config(state="disabled")
        self.update_idletasks()

    def clear_file_list(self):
        """Clears the list of selected PDF files."""
        if self.is_converting:
            messagebox.showwarning("Busy", "Cannot clear list while conversion is in progress.")
            return
        self.pdf_file_list = []
        self.update_file_list_display()
        self.update_status("Selected files list cleared.")
        # Clear the single log area
        if self.log_display_text and self.log_display_text.winfo_exists():
             self.log_display_text.config(state="normal")
             self.log_display_text.delete(1.0, tk.END)
             self.log_display_text.config(state="disabled")

# --- Main Execution Guard --- Functions -----------

def check_dependencies():
    """Checks for required libraries and returns a list of missing ones."""
    missing = []
    try: import fitz
    except ImportError: missing.append("PyMuPDF")
    try: from PIL import Image, ImageTk
    except ImportError: missing.append("Pillow")
    try: from tkinterdnd2 import TkinterDnD
    except ImportError: missing.append("tkinterdnd2")
    return missing

def show_dependency_error(missing_libs):
    """Displays an error message about missing dependencies."""
    error_message = (
        f"The following required libraries are missing:\n\n"
        f"{', '.join(missing_libs)}\n\n"
        f"Please install them, ideally in a virtual environment, using:\n"
        f"pip install -r requirements.txt"
    )
    try:
        # Try showing graphical error message first
        root = tk.Tk()
        root.withdraw() # Hide the empty root window
        messagebox.showerror("Missing Dependencies", error_message)
        root.destroy()
    except tk.TclError:
        # Fallback to console if Tkinter is unavailable or fails
        print("ERROR: Missing Dependencies")
        print(error_message.replace("\n\n", "\n")) # Make console output more compact

if __name__ == "__main__":
    # 1. Check Dependencies before initializing GUI
    missing = check_dependencies()
    if missing:
        show_dependency_error(missing)
        sys.exit(1) # Exit if dependencies are missing

    # 2. Launch the GUI Application
    try:
        app = PdfToEpubApp()
        app.mainloop()
    except Exception as main_err:
        # Catch unexpected errors during GUI initialization or main loop
        print(f"FATAL ERROR: An unexpected error occurred in the application.")
        print(traceback.format_exc())
        try:
             messagebox.showerror("Fatal Error", f"An unexpected error occurred:\n{main_err}\n\nSee console for details.")
        except tk.TclError:
             pass # Console message already printed
        sys.exit(1) 
