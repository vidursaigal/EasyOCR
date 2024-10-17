import os
import pytesseract
import sv_ttk  # Import the Sun Valley theme
from tkinter import Tk, Canvas, Frame, Scrollbar, VERTICAL, RIGHT, Y, BOTH, BOTTOM, filedialog, Entry, messagebox, StringVar
from tkinter import ttk  # Use ttk for widgets for theme compatibility
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageTk
from threading import Thread
from pdf2image import convert_from_path
import io

# Additional imports for saving PDF and Word documents
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas
from docx import Document


class OCRApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # Set window title and fullscreen mode
        self.title("EasyOCR")
        self.state('zoomed')  # Open in fullscreen mode

        # Apply the Sun Valley theme
        sv_ttk.set_theme("light")  # You can change "light" to "dark" for dark mode

        # Aesthetic Font and Padding for Labels
        self.default_font = ("Helvetica", 12)

        # Create main frames
        self.left_frame = ttk.Frame(self)
        self.right_frame = ttk.Frame(self, width=200)  # Adjust width as needed

        self.left_frame.grid(row=0, column=0, sticky='nsew')
        self.right_frame.grid(row=0, column=1, sticky='ns')

        # Configure grid weights
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_rowconfigure(0, weight=1)

        # Instructions and controls on the left frame
        self.instructions_label = ttk.Label(self.left_frame, text=(
            "Instructions:\n\n"
            "1. Drag and drop images or PDF files here.\n"
            "   - Supported image formats: PNG, JPG, JPEG.\n"
            "   - PDFs will be converted to images.\n"
            "2. To rearrange images, type a new number in the field below each image and press Enter.\n"
            "3. Use the slider to zoom in/out of images.\n"
            "4. Click 'Run OCR and Save' to extract text from images.\n"
            "   - You can save the results as a TXT, PDF, or Word document."
        ), anchor="w", justify="left")
        self.instructions_label.pack(padx=10, pady=10, fill='x')

        # Create a scrollable frame for the image previews
        self.canvas = Canvas(self.left_frame)
        self.canvas.pack(side="left", fill=BOTH, expand=True, padx=20, pady=20)

        # Add vertical scrollbar to the canvas
        self.scrollbar = ttk.Scrollbar(self.left_frame, orient=VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a frame inside the canvas for the grid of images
        self.preview_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.preview_frame, anchor='nw')

        # Ensure the canvas scrolls when resized
        self.preview_frame.bind("<Configure>", lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Zoom slider
        self.zoom_label = ttk.Label(self.right_frame, text="Zoom (%)", font=self.default_font)
        self.zoom_label.pack(anchor='w', padx=10, pady=5)

        self.zoom_slider = ttk.Scale(self.right_frame, from_=100, to=300, orient='horizontal', command=self.update_zoom)
        self.zoom_slider.set(200)  # Increase the default zoom level
        self.zoom_slider.pack(fill='x', padx=10, pady=5)

        # Button to run OCR processing
        self.process_button = ttk.Button(self.right_frame, text="Run OCR and Save", command=self.run_ocr)
        self.process_button.pack(pady=10, padx=10, fill='x')

        # Progress bar
        self.progress_var = StringVar()
        self.progress_var.set("Progress: 0%")
        self.progress_label = ttk.Label(self.right_frame, textvariable=self.progress_var)
        self.progress_label.pack(pady=5, padx=10, anchor='w')

        self.progress_bar = ttk.Progressbar(self.right_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill='x', padx=10, pady=5)

        # Lists to store file paths and preview widgets
        self.dropped_files = []
        self.preview_widgets = []
        self.drag_data = {}

        # Image preview size
        self.thumbnail_size = 200  # Increased default size

        # Bind drag-and-drop functionality
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.drop_files)

    def drop_files(self, event):
        files = self.splitlist(event.data)
        for file in files:
            if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                self.dropped_files.append(file)
                self.show_image_preview(file)
            elif file.lower().endswith('.pdf'):
                self.process_pdf(file)
            else:
                messagebox.showerror("Invalid File", f"{file} is not a valid image or PDF.")
        self.update_labels()

    def process_pdf(self, pdf_file):
        try:
            # Convert PDF pages to images
            pages = convert_from_path(pdf_file)
            for i, page in enumerate(pages):
                # Save each page as an image in memory
                img_byte_arr = io.BytesIO()
                page.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                image = Image.open(img_byte_arr)

                # Store the image in a temporary file
                temp_img_path = f"temp_pdf_page_{len(self.dropped_files)}_{i}.png"
                image.save(temp_img_path)

                self.dropped_files.append(temp_img_path)
                self.show_image_preview(temp_img_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process PDF file {pdf_file}: {str(e)}")

    def show_image_preview(self, file):
        # Load and create a thumbnail for the image
        image = Image.open(file)
        image.thumbnail((self.thumbnail_size, self.thumbnail_size))
        thumbnail = ImageTk.PhotoImage(image)

        # Create a container frame for the image and its label (position number)
        container = ttk.Frame(self.preview_frame, width=self.thumbnail_size, height=self.thumbnail_size + 60)
        container.pack_propagate(False)

        # Create a label to display the image
        img_label = ttk.Label(container, image=thumbnail)
        img_label.image = thumbnail  # Keep a reference
        img_label.pack(pady=5)

        # Create an editable Entry for the image order (position number)
        pos_entry = ttk.Entry(container, font=("Arial", 14), justify='center', width=5)
        pos_entry.insert(0, str(len(self.preview_widgets) + 1))
        pos_entry.bind("<Return>", lambda event, widget=container: self.on_position_change(event, widget, pos_entry))
        pos_entry.pack(pady=10)  # Ensure space below the image

        # Add container to the grid
        container.grid_propagate(False)  # Prevent resizing
        self.preview_widgets.append(container)
        self.rearrange_grid()

    def rearrange_grid(self):
        """Arrange the image containers in a responsive grid layout."""
        max_cols = self.get_max_columns()
        for idx, widget in enumerate(self.preview_widgets):
            row = idx // max_cols
            col = idx % max_cols
            widget.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

    def get_max_columns(self):
        """Return the maximum number of columns based on the canvas width and thumbnail size."""
        self.update_idletasks()  # Ensure the canvas size is updated
        canvas_width = self.canvas.winfo_width()
        if canvas_width == 1:
            # Initial width before window is fully rendered
            canvas_width = self.winfo_screenwidth() - self.right_frame.winfo_width() - 40
        total_image_width = self.thumbnail_size + 40  # Image size plus padding
        max_cols = max(1, canvas_width // total_image_width)
        # Adjust max_cols to ensure images are not cut off
        while max_cols > 1 and max_cols * total_image_width > canvas_width:
            max_cols -= 1
        return max_cols

    def update_zoom(self, event=None):
        """Adjust the zoom level for all image previews and rearrange the grid."""
        self.thumbnail_size = int(self.zoom_slider.get())
        for widget in self.preview_widgets:
            file = self.dropped_files[self.preview_widgets.index(widget)]
            self.update_image_preview(widget, file)
        self.rearrange_grid()

    def update_image_preview(self, container, file):
        """Update the thumbnail in the container based on zoom level."""
        image = Image.open(file)
        image.thumbnail((self.thumbnail_size, self.thumbnail_size))
        thumbnail = ImageTk.PhotoImage(image)

        # Find the img_label in the container and update it
        img_label = container.winfo_children()[0]
        img_label.config(image=thumbnail)
        img_label.image = thumbnail  # Keep reference
        container.config(width=self.thumbnail_size, height=self.thumbnail_size + 60)  # Adjust container height for label

    def update_labels(self):
        """Update the number labels on each image container."""
        for i, widget in enumerate(self.preview_widgets):
            pos_entry = widget.winfo_children()[1]
            pos_entry.delete(0, 'end')
            pos_entry.insert(0, str(i + 1))

    def on_position_change(self, event, widget, pos_entry):
        """Handle changes in the position input field, swapping images based on the new position."""
        try:
            new_position = int(pos_entry.get()) - 1
            if new_position < 0 or new_position >= len(self.preview_widgets):
                raise ValueError("Invalid position")

            current_position = self.preview_widgets.index(widget)

            # Swap the widgets in the list
            self.preview_widgets[current_position], self.preview_widgets[new_position] = \
                self.preview_widgets[new_position], self.preview_widgets[current_position]

            # Swap the files in the list
            self.dropped_files[current_position], self.dropped_files[new_position] = \
                self.dropped_files[new_position], self.dropped_files[current_position]

            # Rearrange the grid and update labels
            self.rearrange_grid()
            self.update_labels()

        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid position number.")

    def run_ocr(self):
        if not self.dropped_files:
            messagebox.showerror("Error", "No files dropped!")
            return

        # Disable the button to prevent multiple clicks
        self.process_button.config(state='disabled')

        # Start OCR in a separate thread
        thread = Thread(target=self.perform_ocr)
        thread.start()

    def perform_ocr(self):
        total_files = len(self.dropped_files)
        result_text = ""
        for idx, file in enumerate(self.dropped_files):
            try:
                image = Image.open(file)
                text = pytesseract.image_to_string(image)
                result_text += text + "\n"
            except Exception as e:
                result_text += f"Error processing {file}: {str(e)}\n"

            # Update progress bar
            progress = int(((idx + 1) / total_files) * 100)
            self.progress_bar['value'] = progress
            self.progress_var.set(f"Progress: {progress}%")
            self.update_idletasks()

        # Enable the button after processing
        self.process_button.config(state='normal')

        # Reset progress bar
        self.progress_bar['value'] = 0
        self.progress_var.set("Progress: 0%")

        # Save the result
        self.save_result(result_text)

    def save_result(self, result_text):
        # Ask user for save format
        filetypes = [
            ("Text files", "*.txt"),
            ("PDF files", "*.pdf"),
            ("Word Document", "*.docx")
        ]
        output_file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=filetypes)

        if output_file:
            if output_file.lower().endswith('.txt'):
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(result_text)
            elif output_file.lower().endswith('.pdf'):
                self.save_as_pdf(output_file, result_text)
            elif output_file.lower().endswith('.docx'):
                self.save_as_word(output_file, result_text)
            else:
                messagebox.showerror("Error", "Unsupported file format!")
                return

            messagebox.showinfo("Success", "OCR results saved successfully!")

    def save_as_pdf(self, output_file, text):
        try:
            c = pdf_canvas.Canvas(output_file, pagesize=letter)
            textobject = c.beginText(40, 750)
            lines = text.split('\n')
            for line in lines:
                textobject.textLine(line)
            c.drawText(textobject)
            c.save()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")

    def save_as_word(self, output_file, text):
        try:
            doc = Document()
            doc.add_paragraph(text)
            doc.save(output_file)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Word document: {str(e)}")

# Main code to run the app
if __name__ == "__main__":
    app = OCRApp()
    app.mainloop()
