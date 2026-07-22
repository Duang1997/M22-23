import os
import streamlit as st
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
from PIL import Image

st.set_page_config(
    page_title="ระบบบันทึกภาพและจัดทำแฟ้มผู้ต้องหา", page_icon="📋", layout="centered"
)

st.title("ระบบจัดทำบันทึกแนบท้ายภาพถ่ายผู้ถูกจับและผู้ถูกควบคุม")
st.markdown("---")

# ส่วนที่ 1: ข้อมูลผู้ต้องหา
st.subheader("1. ข้อมูลผู้ต้องหา")
col1, col2 = st.columns(2)
with col1:
  first_name = st.text_input("ชื่อ")
  national_id = st.text_input(
      "เลขประจำตัวประชาชน (13 หลัก)", max_chars=13
  ).strip()
with col2:
  last_name = st.text_input("นามสกุล")

st.markdown("---")

# ส่วนที่ 2: บันทึกภาพถ่าย 4 ด้าน (บน 2 / ล่าง 2)
st.subheader("2. บันทึกภาพถ่าย 4 ด้าน")
angles = {
    "pic_front": "ภาพหน้าตรง (pic_front)",
    "pic_left": "ภาพด้านซ้าย (pic_left)",
    "pic_right": "ภาพด้านขวา (pic_right)",
    "pic_back": "ภาพด้านหลัง (pic_back)",
}

captured_images = {}
for key, label in angles.items():
  st.markdown(f"**{label}**")
  input_method = st.radio(
      f"วิธีนำเข้า {key}",
      ("ถ่ายภาพจากกล้อง", "อัปโหลดไฟล์ภาพ"),
      key=f"method_{key}",
      horizontal=True,
  )
  if input_method == "ถ่ายภาพจากกล้อง":
    img_file = st.camera_input(f"ถ่าย {key}", key=f"cam_{key}")
  else:
    img_file = st.file_uploader(
        f"เลือกไฟล์ {key}", type=["jpg", "jpeg", "png"], key=f"file_{key}"
    )

  captured_images[key] = img_file
  st.markdown("---")


# ฟังก์ชันสร้างเอกสารผ่าน Template
def generate_report_from_template(nid, fname, lname, images):
  template_path = "template.docx"
  if not os.path.exists(template_path):
    st.error("ไม่พบไฟล์ template.docx ในระบบ กรุณาอัปโหลดไฟล์ Template")
    return None

  doc = DocxTemplate(template_path)

  image_context = {}
  temp_files = []

  for key, img_file in images.items():
    if img_file is not None:
      temp_path = f"temp_{key}.jpg"
      img = Image.open(img_file)
      img.save(temp_path)
      temp_files.append(temp_path)
      image_context[key] = InlineImage(doc, temp_path, width=Inches(2.0))
    else:
      image_context[key] = ""

  # กำหนดค่าตัวแปรตาม Template
  context = {
      "suspect_name": f"{fname} {lname}",
      "suspect_id": nid,
      **image_context,
  }

  doc.render(context)

  output_filename = f"ภาพแนบ-{nid}-{fname} {lname}.docx"
  doc.save(output_filename)

  for tf in temp_files:
    if os.path.exists(tf):
      os.remove(tf)

  return output_filename


# ส่วนที่ 3: ปุ่มบันทึกและดาวน์โหลด
if st.button("บันทึกข้อมูลและสร้างไฟล์รายงาน", type="primary"):
  if not national_id or not first_name or not last_name:
    st.error("กรุณากรอกข้อมูล ชื่อ นามสกุล และเลขประจำตัวประชาชนให้ครบถ้วน")
  else:
    with st.spinner("กำลังประมวลผลข้อมูล..."):
      output_file = generate_report_from_template(
          national_id, first_name, last_name, captured_images
      )

      if output_file:
        st.success("สร้างเอกสารสำเร็จ")

        with open(output_file, "rb") as f:
          st.download_button(
              label="ดาวน์โหลดไฟล์ Word (.docx)",
              data=f,
              file_name=output_file,
              mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          )
