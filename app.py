import streamlit as st
from PIL import Image, ImageDraw, ImageFont, ImageOps
from datetime import date
from dateutil.relativedelta import relativedelta
import io, math, os, tempfile, subprocess
from docxtpl import DocxTemplate
from pypdf import PdfWriter, PdfReader

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Wercore Compliance Engine", page_icon="💧", layout="wide")

# ── Paths (Flat Structure for Streamlit Cloud) ────────────────────────────────
# These paths expect all files to be in the exact same main folder as app.py
TPL_COVER       = "Cover_Template.docx"
TPL_INSPECTION  = "Inspection_Template.docx"
TPL_WARRANTY    = "Warranty_Template.docx"
TPL_JOB         = "Job_Completion_Template.docx"

STATIC_DM_CERT  = "DM_Certificate.pdf"
STATIC_MSDS     = "MSDS.pdf"

# ── Session state bootstrap ───────────────────────────────────────────────────
if "tank_count" not in st.session_state:
    st.session_state.tank_count = 1
if "tanks" not in st.session_state:
    st.session_state.tanks = [{}]

def add_tank():
    if st.session_state.tank_count < 20:
        st.session_state.tank_count += 1
        st.session_state.tanks.append({})

def remove_tank(idx):
    if st.session_state.tank_count > 1:
        st.session_state.tanks.pop(idx)
        st.session_state.tank_count -= 1

# ── Font Loader (Cloud Safe) ──────────────────────────────────────────────────
def load_font(size: int) -> ImageFont.FreeTypeFont:
    windows_fonts = [r"C:\Windows\Fonts\calibrib.ttf", r"C:\Windows\Fonts\arialbd.ttf"]
    linux_fonts = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/ubuntu/Ubuntu-B.ttf"
    ]
    for path in windows_fonts + linux_fonts:
        if os.path.exists(path):
            try: return ImageFont.truetype(path, size)
            except Exception: continue
    return ImageFont.load_default()

# ── Image Processing & PDF Canvas ─────────────────────────────────────────────
def build_photo_grid_pages(all_photos: list, photos_per_page: int = 4) -> list:
    n = math.ceil(len(all_photos) / photos_per_page)
    return [all_photos[i * photos_per_page:(i + 1) * photos_per_page] for i in range(n)]

def photo_grid_to_pdf_bytes(page_photos: list, page_num: int, total_pages: int, client_name: str, date_str: str) -> bytes:
    PAGE_W, PAGE_H = 1240, 1754 # A4 Portrait
    PADDING, HEADER_H, FOOTER_H, CELL_GAP = 40, 70, 40, 20
    
    canvas = Image.new("RGB", (PAGE_W, PAGE_H), (255, 255, 255))
    draw   = ImageDraw.Draw(canvas)
    font_header, font_footer, font_caption, font_wm = load_font(28), load_font(20), load_font(22), load_font(24)

    draw.rectangle([(0, 0), (PAGE_W, HEADER_H)], fill=(0, 84, 166))
    draw.text((PADDING, 18), f"Photo Documentation — {client_name}  |  {date_str}", font=font_header, fill=(255, 255, 255))

    caption_h = 35
    grid_top = HEADER_H + PADDING
    grid_bot = PAGE_H - FOOTER_H - PADDING
    cell_w = (PAGE_W - 2 * PADDING - CELL_GAP) // 2
    cell_h = (grid_bot - grid_top - CELL_GAP - caption_h * 2) // 2

    for idx, photo in enumerate(page_photos):
        col, row = idx % 2, idx // 2
        x0 = PADDING + col * (cell_w + CELL_GAP)
        y0 = grid_top + row * (cell_h + CELL_GAP + caption_h)

        img = Image.open(io.BytesIO(photo["image_bytes"])).convert("RGB")
        img = ImageOps.fit(img, (cell_w, cell_h), Image.LANCZOS)
        canvas.paste(img, (x0, y0))

        wm_margin = 15
        bbox = draw.textbbox((0, 0), date_str, font=font_wm)
        wm_x = x0 + cell_w - (bbox[2] - bbox[0]) - wm_margin
        wm_y = y0 + cell_h - (bbox[3] - bbox[1]) - wm_margin
        draw.text((wm_x + 2, wm_y + 2), date_str, font=font_wm, fill=(0, 0, 0, 180))
        draw.text((wm_x, wm_y), date_str, font=font_wm, fill=(255, 255, 255, 230))

        draw.text((x0, y0 + cell_h + 8), f"{photo['tank_name']}  |  {photo['label']}", font=font_caption, fill=(60, 60, 60))

    draw.text((PADDING, PAGE_H - FOOTER_H + 8), "CONFIDENTIAL — Water Tank Cleaning Compliance Report", font=font_footer, fill=(150, 150, 150))
    draw.text((PAGE_W - 180, PAGE_H - FOOTER_H + 8), f"Page {page_num} of {total_pages}", font=font_footer, fill=(150, 150, 150))

    buf = io.BytesIO()
    canvas.save(buf, format="PDF")
    return buf.getvalue()

def site_report_to_pdf_bytes(image_bytes: bytes, client_name: str, date_str: str) -> bytes:
    PAGE_W, PAGE_H = 1240, 1754 
    canvas = Image.new("RGB", (PAGE_W, PAGE_H), (255, 255, 255))
    
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    img = ImageOps.contain(img, (PAGE_W - 80, PAGE_H - 150), Image.LANCZOS)
    
    paste_x = (PAGE_W - img.width) // 2
    paste_y = (PAGE_H - img.height) // 2
    canvas.paste(img, (paste_x, paste_y))

    draw = ImageDraw.Draw(canvas)
    draw.rectangle([(0, 0), (PAGE_W, 70)], fill=(0, 84, 166))
    draw.text((40, 18), f"Handwritten Site Report — {client_name}  |  {date_str}", font=load_font(28), fill=(255, 255, 255))

    buf = io.BytesIO()
    canvas.save(buf, format="PDF")
    return buf.getvalue()

# ── Cloud-Safe Word to PDF Engine ─────────────────────────────────────────────
def fill_word_template(tpl_path: str, context: dict, output_path: str):
    if not os.path.exists(tpl_path):
        raise FileNotFoundError(f"Template not found: {tpl_path}. Ensure it is uploaded to the main GitHub folder.")
    tpl = DocxTemplate(tpl_path)
    tpl.render(context)
    tpl.save(output_path)

def convert_docx_to_pdf_cloud(docx_path: str, output_dir: str):
    """Uses LibreOffice to convert Word to PDF on Linux/Cloud environments."""
    try:
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf", 
            docx_path, "--outdir", output_dir
        ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        raise RuntimeError(f"PDF Conversion failed. Ensure LibreOffice is installed via packages.txt. Error: {e}")

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/water-tank.png", width=72)
    st.title("Project Details")
    st.markdown("---")
    client_name      = st.text_input("🏢 Client Name",       placeholder="e.g. Al Noor Properties")
    authorized_person= st.text_input("👤 Authorized Person", placeholder="e.g. Afzal Rahman")
    client_email     = st.text_input("📧 Client Email",      placeholder="e.g. info@client.com")
    project_location = st.text_input("📍 Project Location",  placeholder="e.g. Dubai Marina")
    service_date     = st.date_input("📅 Date of Service",   value=date.today())
    date_str         = service_date.strftime("%d-%m-%Y")
    
    expiry_date      = service_date + relativedelta(months=+6)
    expiry_date_str  = expiry_date.strftime("%d-%m-%Y")

    st.markdown("---")
    st.markdown("**Handwritten Site Reports**")
    handwritten_files = st.file_uploader("Upload physical reports (Optional)", type=["jpg","png","pdf"], accept_multiple_files=True)
    
    st.markdown("---")
    generate_btn = st.button("🚀 Generate Master Report", type="primary", use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN HEADER & TANK MODULES
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## 💧 Water Tank Cleaning — Master Report Builder")
if not client_name: st.info("👈 Please enter all Client details in the sidebar to begin.")
st.markdown("---")

tank_data_list = []
for i in range(st.session_state.tank_count):
    tank_label = st.session_state.tanks[i].get("name", f"Tank {i + 1}")
    with st.expander(f"🛢️ {tank_label}", expanded=(i == st.session_state.tank_count - 1)):
        col1, col2, col3 = st.columns([2, 1, 0.4])
        with col1:
            tank_name = st.text_input("Tank Name / Type", value=st.session_state.tanks[i].get("name", f"Roof Tank {i + 1}"), key=f"name_{i}")
            st.session_state.tanks[i]["name"] = tank_name
        with col2:
            capacity = st.number_input("Capacity (USG)", min_value=0, max_value=500000, value=st.session_state.tanks[i].get("capacity", 1000), step=100, key=f"cap_{i}")
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️", key=f"del_{i}", help="Remove tank"): remove_tank(i); st.rerun()
        
        c1, c2 = st.columns(2)
        with c1: ph = st.slider("pH Level", 0.0, 14.0, value=float(st.session_state.tanks[i].get("ph", 7.0)), step=0.1, key=f"ph_{i}")
        with c2: tds = st.number_input("TDS Level (ppm)", 0, 5000, value=st.session_state.tanks[i].get("tds", 200), step=10, key=f"tds_{i}")
        
        cc1, cc2, cc3 = st.columns(3)
        biofilm   = cc1.checkbox("🦠 Bio-Film",   key=f"bio_{i}")
        sediments = cc2.checkbox("🪨 Sediments",  key=f"sed_{i}")
        algae     = cc3.checkbox("🌿 Algae",      key=f"alg_{i}")
        
        pc1, pc2 = st.columns(2)
        with pc1: before_files = st.file_uploader("Upload Before Photos", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"before_{i}")
        with pc2: after_files = st.file_uploader("Upload After Photos", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"after_{i}")

        tank_data_list.append({
            "name": tank_name, "capacity_usg": capacity, "ph": ph, "tds": tds,
            "biofilm": "Yes" if biofilm else "No", "sediments": "Yes" if sediments else "No", "algae": "Yes" if algae else "No",
            "before_files": before_files or [], "after_files":  after_files  or []
        })

add_col, _ = st.columns([1, 3])
with add_col:
    if st.button("➕ Add Another Tank", use_container_width=True): add_tank(); st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
#  GENERATE REPORT PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
if generate_btn:
    if not client_name or not project_location:
        st.error("❌ Please enter Client Name and Location."); st.stop()
        
    st.markdown("---")
    progress_bar = st.progress(0, text="Initializing...")
    
    with tempfile.TemporaryDirectory() as tmp_dir:
        try:
            # Context for Word Templates
            context = {
                "client_name": client_name,
                "authorized_person": authorized_person,
                "client_email": client_email,
                "project_location": project_location,
                "date_of_service": date_str,
                "expiry_date_str": expiry_date_str,
                "total_tanks": len(tank_data_list),
                "tanks": [{k: v for k, v in t.items() if k not in ("before_files", "after_files")} for t in tank_data_list]
            }

            # 1. Process Photos & Hand-written Reports
            progress_bar.progress(10, text="Processing Images...")
            all_photos = []
            for t in tank_data_list:
                for lbl, files in [("Before Cleaning", t["before_files"]), ("After Cleaning", t["after_files"])]:
                    for f in files: all_photos.append({"tank_name": t["name"], "label": lbl, "image_bytes": f.read()})
            
            grid_pages = build_photo_grid_pages(all_photos, 4)
            grid_pdfs = [photo_grid_to_pdf_bytes(pg, i+1, len(grid_pages), client_name, date_str) for i, pg in enumerate(grid_pages)]
            
            hw_pdfs = []
            if handwritten_files:
                for f in handwritten_files:
                    if f.name.lower().endswith(".pdf"):
                        hw_pdfs.append(f.read())
                    else:
                        hw_pdfs.append(site_report_to_pdf_bytes(f.read(), client_name, date_str))

            # 2. Fill Word Templates & Convert to PDF
            progress_bar.progress(40, text="Generating Word Documents via LibreOffice...")
            word_files = [
                (TPL_COVER,      os.path.join(tmp_dir, "cover.docx")),
                (TPL_INSPECTION, os.path.join(tmp_dir, "inspection.docx")),
                (TPL_WARRANTY,   os.path.join(tmp_dir, "warranty.docx")),
                (TPL_JOB,        os.path.join(tmp_dir, "job.docx"))
            ]
            
            for tpl, docx_out in word_files:
                fill_word_template(tpl, context, docx_out)
                convert_docx_to_pdf_cloud(docx_out, tmp_dir)

            # 3. Merging Sequence
            progress_bar.progress(80, text="Stapling Final PDF...")
            writer = PdfWriter()
            
            # Sequence: Cover, Inspection, DM Cert, MSDS
            sequence = [
                os.path.join(tmp_dir, "cover.pdf"),
                os.path.join(tmp_dir, "inspection.pdf"),
                STATIC_DM_CERT,
                STATIC_MSDS
            ]
            
            for fpath in sequence:
                if os.path.exists(fpath): 
                    writer.append(fpath)
                else:
                    st.warning(f"⚠️ Could not find {fpath} to merge. Skipping.")
            
            for pdf_bytes in grid_pdfs: writer.append(PdfReader(io.BytesIO(pdf_bytes))) # Photo Grid
            for pdf_bytes in hw_pdfs: writer.append(PdfReader(io.BytesIO(pdf_bytes)))   # Handwritten Reports
            
            writer.append(os.path.join(tmp_dir, "warranty.pdf")) # Warranty
            writer.append(os.path.join(tmp_dir, "job.pdf"))      # Job Completion

            # 4. Output
            final_pdf_path = os.path.join(tmp_dir, "Final_Report.pdf")
            with open(final_pdf_path, "wb") as f: writer.write(f)
            with open(final_pdf_path, "rb") as f: pdf_bytes_final = f.read()

            progress_bar.progress(100, text="✅ Report Complete!")
            st.success("✅ Master Compliance Report generated successfully!")
            
            st.download_button(
                label="⬇️ Download Complete PDF Dossier",
                data=pdf_bytes_final,
                file_name=f"Wercore_Report_{client_name.replace(' ', '_')}_{date_str}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )

        except Exception as e:
            progress_bar.empty()
            st.error(f"❌ Error generating report: {e}")
