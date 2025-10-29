import time
import os
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image
import sys
import subprocess
import threading
import csv

# --- Ensure Pillow is installed ---
try:
    from PIL import Image
except ImportError:
    print("Pillow not found, installing...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pillow"])
    from PIL import Image  # try again after installing
from PIL import ImageTk 
import qrcode

#### --- ROOT WINDOW SETUP --- ####
root = tk.Tk()
root.title('TPB POSTER APP')
root.geometry('360x513')
root['bg'] = '#f0f0f0'
root.resizable(False, False)
root.attributes('-topmost', True)

#### --- INTIIAL VARIABLE DECLERATIONS
hotFolderDir = "C:/Users/jackl/OneDrive/Desktop/TPB/TPBMugsApplication/HotFolder/"
os.makedirs(hotFolderDir, exist_ok=True)  # create it if it doesn't exist
fileDumpDir = "C:/Users/jackl/OneDrive/Desktop/TPB/TPBMugsApplication/FileDump/"
os.makedirs(fileDumpDir, exist_ok=True)  # create it if it doesn't exist
mugApplicationDir = "C:/Users/jackl/OneDrive/Desktop/TPB/TPBMugsApplication/"
errorState = False
 
### --- VISUALS
#TPB LOGO
tpb_image_path = mugApplicationDir + "TPB_logo_f0f0f0.png"
tpb_original_image = Image.open(tpb_image_path)
tpb_banner_image = ImageTk.PhotoImage(tpb_original_image) 

#FILERIPPER LOGO
FR_image_path = mugApplicationDir + "fileRipperLogo2.5cm.png"
FR_original_image = Image.open(FR_image_path)
FR_banner_image = ImageTk.PhotoImage(FR_original_image) 

#SUCCESS GIF
gif = Image.open(mugApplicationDir + "ripper.gif")
frames = []
try:
    while True:
        frames.append(ImageTk.PhotoImage(gif.copy()))
        gif.seek(len(frames))  # move to next frame
except EOFError:
    pass

#FAIL ICON
failIcon = Image.open(mugApplicationDir + "fail_icon.png")
failIcon_tk = ImageTk.PhotoImage(failIcon)

#### --- STYLE SET UP
# BUTTONS
style = ttk.Style()
style.theme_use("clam") 
style.configure("My.TButton",
                foreground="black",
                background="yellow",
                font=("Arial", 12, "bold"))
style.map("My.TButton",
          background=[("active", "light blue"), ("!active", "darkgrey")])

style = ttk.Style()
style.theme_use("clam") 
style.configure("My.TFrame",
                background="#f0f0f0",
                bordercolor="black",
                borderwidth=2,
                relief="solid")

# LABELS
style = ttk.Style()
style.theme_use("clam")
style.configure("My.TLabel",
                foreground="black",
                background="#f0f0f0",
                font=("Arial", 10, "bold"))

style = ttk.Style()
style.theme_use("clam")
style.configure("Cursive.TLabel",
                foreground="black",
                background="#f0f0f0",
                font=("Arial", 10, "italic"))

style = ttk.Style()
style.theme_use("clam")
style.configure("HyperLink.TLabel",
                foreground="#196dff",
                background="#f0f0f0",
                font=("Arial", 10, "underline"))

style = ttk.Style()
style.theme_use("clam")
style.configure("HyperLink_Hover.TLabel",
                foreground="#6093ec",
                background="#f0f0f0",
                font=("Arial", 10, "underline"))

### - CLASS (FOR PLACEHOLDER TEXT)
class EntryWithPlaceholder(tk.Entry):
    def __init__(self, master=None, placeholder="PLACEHOLDER", color='grey', **kwargs):
        super().__init__(master, **kwargs)

        self.placeholder = placeholder
        self.placeholder_color = color
        self.default_fg_colour = self['fg']

        self.bind("<FocusIn>", self.foc_in)
        self.bind("<FocusOut>", self.foc_out)

        self.put_placeholder()

    def put_placeholder(self):
        self.insert(0, self.placeholder)
        self['fg'] = self.placeholder_color

    def foc_in(self, *args):
        if self['fg'] == self.placeholder_color:
            self.delete('0', 'end')
            self['fg'] = self.default_fg_colour

    def foc_out(self, *args):
        if not self.get():
            self.put_placeholder()

### - FUNCTIONS
def update(ind=0):
    frame = frames[ind]
    ripperGIF_label.configure(image=frame)
    ind = (ind + 1) % len(frames)
    root.after(100, update, ind)  

def toggle_on():
    if errorState==False:
        ripperGIF_label.pack()
        success_label.pack()
    if errorState==True:
        failIcon_label.pack()
        fail_label.pack()

def toggle_off():
    ripperGIF_label.pack_forget()     
    success_label.pack_forget()   
    failIcon_label.pack_forget()
    fail_label.pack_forget()

def flatten_to_rgb(img, bg_colour =(255,255,255)):
    # Force any image into RGB with a solid background - forcing transparent pixels to white
    #ensure that we have an ALPHA channel
    img = img.convert("RGBA") 
    #Create a white bg
    bg = Image.new("RGB", img.size, bg_colour)
    #iterate over all pixels and replace transparent ones with bg_colour
    alpha = img.getchannel("A")
    # Put image over background using alpha as mask
    bg.paste(img, mask=alpha)

    #Force fully transparent pizels to bg_colour
    pixels = bg.load()
    for y in range(bg.height):
        for x in range (bg.width):
            if alpha.getpixel((x,y)) == 0:
                pixels[x,y] = bg_colour
    return bg

def urlValidityChecker(url):
    global errorState
    # Check first if the imgLink_var text box is blank
    if not url or str(url).strip() == "":
        fail_label.config(text = "imgLink cannot be blank!")
        errorState = True
        return False
    try:
        # Then check if the string entered actually leads anywhere
        parsedURL = urllib.parse.urlparse(url)
        if not parsedURL.scheme or not parsedURL.netloc:
            fail_label.config(text="imgLink is not a valid URL!")
            errorState = True
            return False
        # If the string passes this check, we now check to see if its a url we can actually work with (has an easily pullable img file)
        req = urllib.request.Request(url, method='GET')
        with urllib.request.urlopen(req) as response:
            content_type = response.headers.get('Content-Type')
            if response.status == 200 and content_type and content_type.startswith('image/'):
                return True
            else:
                print(f"Invalid content type or status: {response.status}, {content_type}")
                return False
    except urllib.error.HTTPError as e:
        print(f"❌ HTTP Error: {e.code} {e.reason}")
        return False
    except urllib.error.URLError as e:
        print(f"❌ URL Error: {e.reason}")
        return False

def imgLink_webscrape():
    global errorState
    global fileCount
    global prodfile
    global prodfile_copy
    global prodfile_path
    global save_path
    #Check if a file by the same name already exists
    fileCount = 0
    while True:
        if fileCount == 0:
            filename = f"{shipmentNo_var.get()}.png"
        else:
            filename = f"{shipmentNo_var.get()}_{fileCount}.png"
        save_path = os.path.join(hotFolderDir, filename)
        if not os.path.exists(save_path):
            break
        fileCount += 1
    # Check to see if the link is valid
    img_url = str(imgLink_var.get())
    if urlValidityChecker(img_url):
        urllib.request.urlretrieve(str(imgLink_var.get()), save_path)
        # Need to open it this way, so it automatically closes, and Photoshop will be able to reference it
        with Image.open(save_path) as prodfile:
            # Convert the image from RGBA to RGB to replace the transparency with a white BG
            prodfile_copy = flatten_to_rgb(prodfile)
        prodfile_path = save_path
        print("prodfile_path is:" + prodfile_path)
        print("Prod Image Downloaded")
        errorState = False
    else:  
        errorState = True

def qrCode_generate():
    fileCount = 0
    global qrCodeFile
    while True:
        if fileCount == 0:
            filename = str(shipmentNo_var.get()) + "_QRcode.png"
        else:
            filename = str(shipmentNo_var.get()) + "_QRcode-" + str(fileCount) + ".png"
        save_path = os.path.join(hotFolderDir, filename)
        global qrCodeFile_Path
        qrCodeFile_Path = str(save_path)
        if not os.path.exists(save_path):
            break
        fileCount += 1
    img = qrcode.make(str(qrCode_var.get()))
    # Convert the image from RGBA to RGB to replace the transparency with a white BG
    qrCodeFile = flatten_to_rgb(img)
    type(img)
    img.save(save_path)
 
def generateButton():
    global errorState
    toggle_off()
    imgLink_webscrape()
    if errorState==False:
        qrCode_generate()
        success_label.config(text = "Print file pulled... generating Prod file now")
        toggle_on()
        # STEP 1: Create a white background template
        template_width = 2539
        template_height = 1032
        template_bg = Image.new("RGB", (template_width, template_height), (255, 255, 255))
        # STEP 2: Paste prodfile_copy onto the white background, centered
        x_offset = (template_width - prodfile_copy.width) // 2
        y_offset = (template_height - prodfile_copy.height) // 2
        template_bg.paste(prodfile_copy, (x_offset, y_offset))
        # 🔹 At this point, template_bg is your prodfile centered on white
        # STEP 3: Create new image to hold template_bg + QR code
        combined_width = template_bg.width + qrCodeFile.width + 10  # 10px margin between
        combined_height = max(template_bg.height, qrCodeFile.height)
        combined_image = Image.new("RGB", (combined_width, combined_height), (255, 255, 255))
        # STEP 4: Paste the white-backgrounded prodfile first
        combined_image.paste(template_bg, (0, 0))
        # STEP 5: Paste the QR code to the right, vertically centered
        qr_offset_x = template_bg.width + 10  # after prodfile + margin
        qr_offset_y = combined_height - qrCodeFile.height
        combined_image.paste(qrCodeFile, (qr_offset_x, qr_offset_y))
        # STEP 6: Save the result
        filename = f"{shipmentNo_var.get()}-{fileCount}.png"
        save_path = os.path.join(hotFolderDir, filename)
        combined_image.save(save_path)
        print(f"Saved combined image at: {save_path}")
        # STEP 7: Rotate image for your PSD layout
        rotated_image = combined_image.transpose(Image.ROTATE_90)
        rotated_image.save(save_path)
        print("Rotated and saved")
        # STEP 8: Clean up
        prodfile_copy.close()
        qrCodeFile.close()
        combined_image.close()
        template_bg.close()
        os.remove(prodfile_path)
        os.remove(qrCodeFile_Path)
        toggle_off()
        success_label.config(text = "Nice! Prod file now in HotFolder")
        toggle_on()
        print("✅ Finished generateButton()")
    if errorState==True: 
        toggle_on()
        return   

def process_poster_csv(file_path, save_dir):
    """Process poster CSV, download artwork, and generate posters with QR codes."""
    print(f"📂 Processing CSV: {file_path}")

    # --- Read CSV ---
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        rows = list(reader)

    total_rows = len(rows)
    print(f"🧾 Found {total_rows} rows in CSV\n")

    for i, row in enumerate(rows, start=1):
        print(f"➡️  [{i}/{total_rows}] Starting row {i}")

        # --- Extract required fields ---
        order_number = row.get("Id", "").strip()
        shipment_id = row.get("Shipment id", "").strip()
        shipment_item = row.get("Shipment item number", "").strip()
        artwork_url = row.get("Artwork 1 artwork file url", "").strip()

        if not order_number or not shipment_id or not artwork_url:
            print(f"⚠️  Skipping row {i}: Missing one or more required fields")
            continue

        file_name = f"{shipment_id}-{shipment_item}"
        print(f"   ➤ File name: {file_name}")
        print(f"   ➤ QR ID: {order_number}")
        print(f"   ➤ Artwork URL: {artwork_url}")

        # --- Generate unique temp name to prevent overwriting ---
        random_suffix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=6))
        temp_ext = ".pdf" if artwork_url.lower().endswith(".pdf") else ".png"
        temp_filename = f"{file_name}_{random_suffix}{temp_ext}"
        temp_path = os.path.join(save_dir, temp_filename)

        # --- Download artwork ---
        try:
            urllib.request.urlretrieve(artwork_url, temp_path)
            print(f"📥  Downloaded artwork to: {temp_path}")
        except Exception as e:
            print(f"❌  Failed to download artwork for {file_name}: {e}")
            continue

        # --- Process poster ---
        generate_dynamic_poster(temp_path, order_number, save_dir, i, total_rows)

        # small delay between downloads (helps prevent server throttling)
        time.sleep(0.25)

    print(f"\n✅ All {total_rows} rows processed.\n")

def generate_dynamic_poster(poster_path, order_number, save_dir, index, total):
    """Add cut line + QR code to poster, converting PDFs if necessary."""
    try:
        print(f"🖼️  [{index}/{total}] Processing poster for order: {order_number}")

        # --- Detect and handle PDF ---
        if poster_path.lower().endswith(".pdf"):
            print(f"   🧾 Converting PDF to image...")
            with Image.open(poster_path) as pdf:
                pdf.load()
                img = pdf.convert("RGB")
        else:
            img = Image.open(poster_path).convert("RGB")

        # --- Add white background and red cut line ---
        margin = 50
        line_thickness = 10
        template_width = img.width + (margin * 2)
        template_height = img.height + (margin * 2)
        template_bg = Image.new("RGB", (template_width, template_height), (255, 255, 255))

        x_offset = (template_width - img.width) // 2
        y_offset = (template_height - img.height) // 2
        template_bg.paste(img, (x_offset, y_offset))

        draw = ImageDraw.Draw(template_bg)
        draw.rectangle(
            [0, template_bg.height - line_thickness, template_bg.width, template_bg.height],
            fill=(255, 0, 0)
        )

        # --- Generate and place QR code ---
        qr_img = qrcode.make(order_number)
        qr_rgb = qr_img.convert("RGB")

        combined_height = template_bg.height + qr_rgb.height + 20
        combined_img = Image.new("RGB", (template_bg.width, combined_height), (255, 255, 255))
        combined_img.paste(template_bg, (0, 0))

        qr_x = (combined_img.width - qr_rgb.width) // 2
        qr_y = template_bg.height + 10
        combined_img.paste(qr_rgb, (qr_x, qr_y))

        # --- Save final image ---
        output_filename = f"{order_number}_{index}.png"
        output_path = os.path.join(save_dir, output_filename)
        combined_img.save(output_path)
        print(f"✅  Saved poster: {output_path}")

    except Exception as e:
        print(f"❌  Error generating poster for {order_number}: {e}")

    finally:
        # --- Clean up temp file ---
        try:
            os.remove(poster_path)
            print(f"🗑️  Deleted temp file: {poster_path}\n")
        except Exception as e:
            print(f"⚠️  Could not delete temp file: {e}\n")

def csvUpload_click(event=None):
    """Upload CSV file and process posters with progress bar."""
    file_path = filedialog.askopenfilename(
        title="Select .csv file", 
        filetypes=[("CSV files", "*.csv")]
    )
    if not file_path:
        print("No file selected.")
        return

    upload_button.config(state='disabled')

    def worker():
        try:
            print(f"⏳ Starting poster batch from: {file_path}")
            root.after(0, lambda: (
                progress_label.config(text="Processing… please wait"),
                progress_label.pack(pady=5),
                progress_bar.pack(pady=5),
                progress_bar.start(10)
            ))

            process_poster_csv(file_path, hotFolderDir)

        finally:
            root.after(0, lambda: (
                upload_button.config(state='normal'),
                progress_bar.stop(),
                progress_bar.pack_forget(),
                progress_label.config(text="Upload complete ✅")
            ))

    threading.Thread(target=worker, daemon=True).start()
    
#### --- WINDOW ELEMENT PROPERTIES --- ####
# FRAME
mainFrame = ttk.Frame(root, style='My.TFrame')
mainFrame.pack(padx=20, pady=20, fill="both", expand=True)
# FR banner LABEL
FR_banner_label = tk.Label(mainFrame, image=FR_banner_image, border=0)
FR_banner_label.pack(pady=3)
# TPB banner LABEL
tpb_banner_label = tk.Label(mainFrame, image=tpb_banner_image, border=0)
tpb_banner_label.pack()

#----
# imgLink ENTRY
imgLink_var = tk.StringVar()
imgLink_entry = EntryWithPlaceholder(mainFrame, "ImgLink", textvariable=imgLink_var)
imgLink_entry.pack(pady=5)

#----
# qrCode ENTRY
qrCode_var = tk.StringVar()
qrCode_entry = EntryWithPlaceholder(mainFrame, "QR Code", textvariable=qrCode_var)
qrCode_entry.pack(pady=5)
#----
# shipmentNo ENTRY
shipmentNo_var = tk.StringVar()
shipmentNo_entry = EntryWithPlaceholder(mainFrame,"Shipment Number", textvariable=shipmentNo_var)
shipmentNo_entry.pack(pady=5)
#----

# Generate BUTTON
Generate_button = ttk.Button(mainFrame, text='Generate Prod File', style='My.TButton', width=20 , command=generateButton)
Generate_button.pack(pady=3)
#----
# Upload BUTTON
upload_button = ttk.Button(mainFrame, text='Upload .csv File', style='My.TButton', command=csvUpload_click)
upload_button.pack(pady=5)
#---- 
# progress BAR
progress_bar = ttk.Progressbar(mainFrame, mode='indeterminate', length=250)
# Progress label 
progress_label = ttk.Label(mainFrame, text="")
# Ripper GIF LABEL 
ripperGIF_label = tk.Label(mainFrame, border=0)
# success LABEL
success_label = ttk.Label(mainFrame, text='nice!', style='Cursive.TLabel')
# Fail_icon LABEL
failIcon_label = ttk.Label(mainFrame, border=0, image=failIcon_tk)
# fail LABEL
fail_label = ttk.Label(mainFrame, text='Uh Oh - no good!', style='Cursive.TLabel')

# Has to be at the very end of the program. 
update()
root.mainloop()