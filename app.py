import streamlit as st
import zipfile
import os
import io
import shutil
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE
from pptx.util import Inches
import random
from datetime import datetime, date
from PIL import Image
from PIL.ExifTags import TAGS
import tempfile
import streamlit.components.v1 as components

# --- Session State Initialization ---

def init_session():
    defaults = {
        'current_step': 1,
        'pptx_data': None,
        'slide_analysis': None,
        'placeholders_config': {},
        'processing_details': [],
        'show_details_needed': False,
        'selected_placeholder': None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def add_detail(message, detail_type="info"):
    st.session_state.processing_details.append({'message': message, 'type': detail_type})
    if detail_type in ['error', 'warning']:
        st.session_state.show_details_needed = True

def clear_details():
    st.session_state.processing_details = []
    st.session_state.show_details_needed = False

def show_details_section():
    if st.session_state.processing_details:
        with st.expander("📋 تفاصيل المعالجة", expanded=False):
            for detail in st.session_state.processing_details:
                if detail['type'] == 'success':
                    st.success(detail['message'])
                elif detail['type'] == 'warning':
                    st.warning(detail['message'])
                elif detail['type'] == 'error':
                    st.error(detail['message'])
                else:
                    st.info(detail['message'])

def analyze_slide_placeholders(prs):
    if len(prs.slides) == 0:
        return None
    first_slide = prs.slides[0]
    slide_width, slide_height = prs.slide_width, prs.slide_height
    placeholders = {
        'image_placeholders': [],
        'text_placeholders': [],
        'title_placeholders': [],
        'slide_dimensions': {
            'width': slide_width,
            'height': slide_height,
            'width_inches': slide_width / 914400,
            'height_inches': slide_height / 914400
        }
    }
    placeholder_id = 0

    def clamp_percent(val):
        return max(0, min(val, 100))

    for shape in first_slide.shapes:
        if shape.is_placeholder:
            placeholder_type = shape.placeholder_format.type
            left_percent = clamp_percent((shape.left / slide_width) * 100)
            top_percent = clamp_percent((shape.top / slide_height) * 100)
            width_percent = clamp_percent((shape.width / slide_width) * 100)
            height_percent = clamp_percent((shape.height / slide_height) * 100)
            placeholder_info = {
                'id': placeholder_id,
                'type': placeholder_type,
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height,
                'left_percent': left_percent,
                'top_percent': top_percent,
                'width_percent': width_percent,
                'height_percent': height_percent,
                'rotation': getattr(shape, 'rotation', 0)
            }
            if placeholder_type == PP_PLACEHOLDER.PICTURE:
                placeholder_info['current_content'] = "صورة"
                placeholders['image_placeholders'].append(placeholder_info)
            elif placeholder_type == PP_PLACEHOLDER.TITLE:
                placeholder_info['current_content'] = (
                    shape.text_frame.text if hasattr(shape, 'text_frame') and shape.text_frame.text else "العنوان"
                )
                placeholders['title_placeholders'].append(placeholder_info)
            else:
                if hasattr(shape, 'text_frame') and shape.text_frame:
                    placeholder_info['current_content'] = (
                        shape.text_frame.text if shape.text_frame.text else f"نص {placeholder_id + 1}"
                    )
                    placeholders['text_placeholders'].append(placeholder_info)
            placeholder_id += 1

    for shape in first_slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not shape.is_placeholder:
            left_percent = clamp_percent((shape.left / slide_width) * 100)
            top_percent = clamp_percent((shape.top / slide_height) * 100)
            width_percent = clamp_percent((shape.width / slide_width) * 100)
            height_percent = clamp_percent((shape.height / slide_height) * 100)
            image_info = {
                'id': placeholder_id,
                'type': 'regular_image',
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height,
                'left_percent': left_percent,
                'top_percent': top_percent,
                'width_percent': width_percent,
                'height_percent': height_percent,
                'rotation': getattr(shape, 'rotation', 0),
                'current_content': "صورة موجودة"
            }
            placeholders['image_placeholders'].append(image_info)
            placeholder_id += 1
    return placeholders

def render_slide_preview_interactive(slide_analysis):
    # Interactive preview: Show boxes for all placeholders with clickable buttons below
    dimensions = slide_analysis['slide_dimensions']
    max_width = 820
    aspect_ratio = dimensions['width'] / dimensions['height']
    display_width = max_width if aspect_ratio > 1 else max_width * aspect_ratio
    display_height = max_width / aspect_ratio if aspect_ratio > 1 else max_width

    def clamp_box(left, top, width, height):
        left = max(0, min(left, display_width-8))
        top = max(0, min(top, display_height-8))
        width = max(8, min(width, display_width-left))
        height = max(8, min(height, display_height-top))
        return left, top, width, height

    placeholder_html = ""
    btns = []
    # Images
    for i, placeholder in enumerate(slide_analysis['image_placeholders']):
        left = (placeholder['left_percent'] / 100) * display_width
        top = (placeholder['top_percent'] / 100) * display_height
        width = (placeholder['width_percent'] / 100) * display_width
        height = (placeholder['height_percent'] / 100) * display_height
        left, top, width, height = clamp_box(left, top, width, height)
        placeholder_html += f"""
        <button onclick="window.parent.postMessage({{'phType':'image','phIdx':{i}}}, '*')" style="
            position: absolute;
            left: {left}px;
            top: {top}px;
            width: {width}px;
            height: {height}px;
            border: 2px solid #ff6b6b;
            background: rgba(255, 107, 107, 0.22);
            font-size: 15px;
            color: #ff6b6b;
            font-weight: bold;
            border-radius: 5px;
            cursor:pointer;
            z-index:4;
            pointer-events: auto;
        ">🖼️ صورة {i+1}</button>
        """
        btns.append((f"🖼️ صورة {i+1}", ('image', i)))
    # Texts
    for i, placeholder in enumerate(slide_analysis['text_placeholders']):
        left = (placeholder['left_percent'] / 100) * display_width
        top = (placeholder['top_percent'] / 100) * display_height
        width = (placeholder['width_percent'] / 100) * display_width
        height = (placeholder['height_percent'] / 100) * display_height
        left, top, width, height = clamp_box(left, top, width, height)
        placeholder_html += f"""
        <button onclick="window.parent.postMessage({{'phType':'text','phIdx':{i}}}, '*')" style="
            position: absolute;
            left: {left}px;
            top: {top}px;
            width: {width}px;
            height: {height}px;
            border: 2px solid #4ecdc4;
            background: rgba(78, 205, 196, 0.18);
            font-size: 13px;
            color: #4ecdc4;
            font-weight: bold;
            border-radius: 5px;
            cursor:pointer;
            z-index:4;
            pointer-events: auto;
        ">📝 نص {i+1}</button>
        """
        btns.append((f"📝 نص {i+1}", ('text', i)))
    # Titles
    for i, placeholder in enumerate(slide_analysis['title_placeholders']):
        left = (placeholder['left_percent'] / 100) * display_width
        top = (placeholder['top_percent'] / 100) * display_height
        width = (placeholder['width_percent'] / 100) * display_width
        height = (placeholder['height_percent'] / 100) * display_height
        left, top, width, height = clamp_box(left, top, width, height)
        placeholder_html += f"""
        <button onclick="window.parent.postMessage({{'phType':'title','phIdx':{i}}}, '*')" style="
            position: absolute;
            left: {left}px;
            top: {top}px;
            width: {width}px;
            height: {height}px;
            border: 2px solid #45b7d1;
            background: rgba(69, 183, 209, 0.18);
            font-size: 13px;
            color: #45b7d1;
            font-weight: bold;
            border-radius: 5px;
            cursor:pointer;
            z-index:4;
            pointer-events: auto;
        ">📋 عنوان {i+1}</button>
        """
        btns.append((f"📋 عنوان {i+1}", ('title', i)))
    # Slide Frame
    html_code = f"""
    <div style="
        width: {display_width}px;
        height: {display_height}px;
        border: 2px solid #ddd;
        position: relative;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        margin: 20px auto;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        overflow: hidden;
        z-index: 5;">
        <div style="
            position: absolute;
            top: 5px;
            left: 5px;
            background: rgba(0,0,0,0.65);
            color: white;
            padding: 4px 10px;
            border-radius: 5px;
            font-size: 12px;
            z-index:10;">
            أبعاد الشريحة: {dimensions['width_inches']:.1f}" × {dimensions['height_inches']:.1f}"
        </div>
        {placeholder_html}
    </div>
    <script>
    window.addEventListener('message', function(e) {{
        const d = e.data;
        if (d && d.phType && d.phIdx !== undefined) {{
            window.parent.streamlitSendPh && window.parent.streamlitSendPh(d.phType, d.phIdx)
        }}
    }});
    </script>
    """
    # For Streamlit communication, we can't capture JS events directly, so below we use clickable buttons
    components.html(html_code, height=int(display_height)+60, scrolling=False)
    # As fallback: Also render a grid of clickable buttons below
    st.markdown("#### اختر موضع للتعديل:")
    btn_cols = st.columns(3)
    for idx, (label, tag) in enumerate(btns):
        with btn_cols[idx % 3]:
            if st.button(label, key=f"btn_{tag[0]}_{tag[1]}"):
                st.session_state.selected_placeholder = tag

def configure_selected_placeholder(analysis):
    # Show config panel for selected placeholder only
    selected = st.session_state.get('selected_placeholder')
    if not selected:
        st.info("اختر موضع Placeholder من المعاينة أو الأزرار أعلاه")
        return

    ptype, idx = selected
    if ptype == 'image':
        ph = analysis['image_placeholders'][idx]
        n_img_ph = len(analysis['image_placeholders'])
        st.markdown(f"#### إعداد صورة {idx+1}")
        use_image = st.checkbox("استبدال هذه الصورة", value=True, key=f"use_image_{ph['id']}")
        image_order = st.number_input(
            "ترتيب الصورة",
            min_value=1,
            max_value=n_img_ph,
            value=min(idx+1, n_img_ph),
            key=f"image_order_{ph['id']}"
        )
        st.markdown(f"""
        **معلومات الموضع:**
        - العرض: {ph['width_percent']:.1f}%
        - الارتفاع: {ph['height_percent']:.1f}%
        - الموقع: ({ph['left_percent']:.1f}%, {ph['top_percent']:.1f}%)
        """)
        # Save config
        images_conf = st.session_state.placeholders_config.get('images', {})
        images_conf[f"image_{ph['id']}"] = {
            'use': use_image,
            'order': image_order,
            'placeholder_info': ph
        }
        st.session_state.placeholders_config['images'] = images_conf

    elif ptype == 'text':
        ph = analysis['text_placeholders'][idx]
        st.markdown(f"#### إعداد نص {idx+1}")
        fill_option = st.radio(
            "كيف تريد ملء هذا النص؟",
            ("ترك فارغ", "نص ثابت", "تاريخ", "تاريخ الصورة", "اسم المجلد"),
            key=f"text_fill_option_{ph['id']}",
            index=0
        )
        placeholder_config = {'type': fill_option, 'value': None}
        if fill_option == "نص ثابت":
            custom_text = st.text_input(
                "أدخل النص المطلوب:",
                key=f"custom_text_{ph['id']}",
                placeholder="مثال: اسم المشروع، اسم الشركة، إلخ...")
            placeholder_config['value'] = custom_text
        elif fill_option == "تاريخ":
            date_option = st.radio(
                "اختر نوع التاريخ:",
                ("تاريخ اليوم", "تاريخ مخصص"),
                key=f"date_option_{ph['id']}"
            )
            if date_option == "تاريخ اليوم":
                placeholder_config['value'] = "today"
            else:
                custom_date = st.date_input(
                    "اختر التاريخ:",
                    key=f"custom_date_{ph['id']}",
                    value=date.today()
                )
                placeholder_config['value'] = custom_date.strftime('%Y-%m-%d')
        elif fill_option == "تاريخ الصورة":
            placeholder_config['value'] = "image_date"
            st.info("سيتم استخدام تاريخ التقاط الصورة الأولى في كل مجلد")
        elif fill_option == "اسم المجلد":
            placeholder_config['value'] = "folder_name"
            st.info("سيتم استخدام اسم المجلد كنص")
        texts_conf = st.session_state.placeholders_config.get('texts', {})
        texts_conf[f"text_{ph['id']}"] = placeholder_config
        st.session_state.placeholders_config['texts'] = texts_conf

    elif ptype == 'title':
        ph = analysis['title_placeholders'][idx]
        st.markdown(f"#### إعداد عنوان {idx+1}")
        # Optional: Allow customizing the title text, for now just info
        st.info("سيتم تعيين عنوان الشريحة باسم المجلد تلقائياً.")

def step1_upload_pptx():
    st.title("🔄 PowerPoint Image & Placeholder Replacer")
    st.markdown("---")
    st.markdown("### 📂 الخطوة 1: رفع ملف PowerPoint")
    st.info("ارفع ملف PowerPoint (.pptx) لتحليل القالب وإعداد الخيارات")
    uploaded_pptx = st.file_uploader(
        "اختر ملف PowerPoint (.pptx)",
        type=["pptx"],
        key="pptx_uploader",
        help="ارفع ملف PowerPoint الذي تريد استبدال الصور والنصوص فيه"
    )
    if uploaded_pptx:
        if st.button("📊 تحليل القالب والمتابعة", type="primary"):
            with st.spinner("🔍 جاري تحليل ملف PowerPoint..."):
                try:
                    st.session_state.pptx_data = uploaded_pptx.read()
                    prs = Presentation(io.BytesIO(st.session_state.pptx_data))
                    slide_analysis = analyze_slide_placeholders(prs)
                    if slide_analysis:
                        st.session_state.slide_analysis = slide_analysis
                        st.session_state.current_step = 2
                        st.rerun()
                    else:
                        st.error("❌ لا توجد شرائح في الملف أو حدث خطأ في التحليل")
                except Exception as e:
                    st.error(f"❌ خطأ في تحليل الملف: {e}")
    with st.expander("📖 تعليمات الاستخدام", expanded=False):
        st.markdown("""
        ### 🎯 كيفية استخدام التطبيق:
        #### **الخطوة 1: رفع ملف PowerPoint**
        - ارفع ملف PowerPoint (.pptx) يحتوي على القالب المطلوب
        - سيتم تحليل الشريحة الأولى واستخراج جميع placeholders
        #### **الخطوة 2: إعداد Placeholders**
        - ستظهر معاينة تفاعلية للشريحة مع جميع placeholders
        - يمكنك تخصيص كل placeholder حسب احتياجاتك بالنقر عليه
        - للصور: اختيار الترتيب أو عدم الاستبدال
        - للنصوص: اختيار نوع المحتوى (ثابت، تاريخ، إلخ)
        #### **الخطوة 3: رفع الصور والمعالجة**
        - ارفع ملف ZIP يحتوي على مجلدات الصور
        - ابدأ المعالجة وفقاً للإعدادات المحددة
        """)

def step2_configure_placeholders():
    st.title("✅ ⚙️ إعداد Placeholders")
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        if st.button("⬅️ العودة للخطوة السابقة"):
            st.session_state.current_step = 1
            st.rerun()
    with col3:
        if st.button("➡️ المتابعة للمعالجة", type="primary"):
            st.session_state.current_step = 3
            st.rerun()
    st.markdown("### 👁️ معاينة القالب")
    analysis = st.session_state.slide_analysis
    if analysis:
        render_slide_preview_interactive(analysis)
        st.markdown("---")
        configure_selected_placeholder(analysis)
        st.markdown("---")
        # Show summary
        if st.checkbox("📋 عرض ملخص الإعدادات", value=False):
            st.markdown("### 📋 ملخص الإعدادات الحالية")
            image_config = st.session_state.placeholders_config.get('images', {})
            text_config = st.session_state.placeholders_config.get('texts', {})
            if image_config:
                st.markdown("#### 🖼️ إعدادات الصور:")
                for key, config in image_config.items():
                    if config['use']:
                        st.success(f"✅ صورة {config['order']}: سيتم استبدالها بالصورة رقم {config['order']} من كل مجلد")
                    else:
                        st.info(f"⏭️ صورة: لن يتم استبدالها")
            if text_config:
                st.markdown("#### 📝 إعدادات النصوص:")
                for key, config in text_config.items():
                    if config['type'] == 'ترك فارغ':
                        st.info(f"⏭️ نص: سيترك فارغاً")
                    elif config['type'] == 'نص ثابت':
                        st.success(f"✅ نص ثابت: '{config['value']}'")
                    elif config['type'] == 'تاريخ':
                        st.success(f"📅 تاريخ: {config['value']}")
                    elif config['type'] == 'تاريخ الصورة':
                        st.success(f"📸 تاريخ الصورة: سيتم استخراجه من metadata")
                    elif config['type'] == 'اسم المجلد':
                        st.success(f"📁 اسم المجلد: سيتم استخدام اسم كل مجلد")

def get_image_date(image_path):
    try:
        with Image.open(image_path) as img:
            exifdata = img.getexif()
            for tag_id in exifdata:
                tag = TAGS.get(tag_id, tag_id)
                data = exifdata.get(tag_id)
                if tag in ['DateTime', 'DateTimeOriginal', 'DateTimeDigitized']:
                    try:
                        date_obj = datetime.strptime(str(data), '%Y:%m:%d %H:%M:%S')
                        return date_obj.strftime('%Y-%m-%d')
                    except Exception:
                        continue
        timestamp = os.path.getmtime(image_path)
        return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d')
    except Exception:
        return datetime.now().strftime('%Y-%m-%d')

def apply_configured_placeholders(slide, folder_path, folder_name, slide_analysis, placeholders_config):
    imgs = sorted(
        [f for f in os.listdir(folder_path)
         if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
    )
    image_config = placeholders_config.get('images', {})
    for config_key, config in image_config.items():
        if config['use'] and config['order'] and config['order'] <= len(imgs):
            image_path = os.path.join(folder_path, imgs[config['order'] - 1])
            placeholder_info = config['placeholder_info']
            target_shapes = []
            for shape in slide.shapes:
                if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                    target_shapes.append(shape)
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not shape.is_placeholder:
                    target_shapes.append(shape)
            for shape in target_shapes:
                shape_left_percent = (shape.left / slide_analysis['slide_dimensions']['width']) * 100
                shape_top_percent = (shape.top / slide_analysis['slide_dimensions']['height']) * 100
                if (abs(shape_left_percent - placeholder_info['left_percent']) < 5 and
                        abs(shape_top_percent - placeholder_info['top_percent']) < 5):
                    try:
                        img_bytes = open(image_path, "rb").read()
                        image_stream = io.BytesIO(img_bytes)
                        if shape.is_placeholder:
                            shape.insert_picture(image_stream)
                        else:
                            original_left = shape.left
                            original_top = shape.top
                            original_width = shape.width
                            original_height = shape.height
                            shape_element = shape._element
                            shape_element.getparent().remove(shape_element)
                            slide.shapes.add_picture(image_stream, original_left, original_top, original_width, original_height)
                        add_detail(f"✅ تم استبدال الصورة {config['order']}: {os.path.basename(image_path)}", "success")
                        break
                    except Exception as e:
                        add_detail(f"❌ Image replacement failed for {os.path.basename(image_path)}: {str(e)}", "error")
    text_config = placeholders_config.get('texts', {})
    text_shapes = [
        shape for shape in slide.shapes
        if (shape.is_placeholder and
            shape.placeholder_format.type not in [PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.TITLE] and
            hasattr(shape, 'text_frame') and shape.text_frame)
    ]
    text_index = 0
    for config_key, config in text_config.items():
        if text_index < len(text_shapes):
            shape = text_shapes[text_index]
            try:
                if config['type'] == "ترك فارغ":
                    shape.text_frame.text = ""
                elif config['type'] == "نص ثابت":
                    if config['value']:
                        shape.text_frame.text = config['value']
                elif config['type'] == "تاريخ":
                    if config['value'] == "today":
                        date_text = datetime.now().strftime('%Y-%m-%d')
                    else:
                        date_text = config['value']
                    shape.text_frame.text = date_text
                elif config['type'] == "تاريخ الصورة" and imgs:
                    first_image_path = os.path.join(folder_path, imgs[0])
                    image_date = get_image_date(first_image_path)
                    shape.text_frame.text = image_date
                elif config['type'] == "اسم المجلد":
                    shape.text_frame.text = folder_name
                add_detail(f"✅ تم تطبيق النص: {config['type']}", "success")
            except Exception as e:
                add_detail(f"⚠ خطأ في تطبيق النص: {e}", "warning")
            text_index += 1
    title_shapes = [
        shape for shape in slide.shapes
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE
    ]
    if title_shapes:
        title_shapes[0].text = folder_name
        add_detail(f"✅ تم تحديث العنوان: {folder_name}", "success")

def step3_process_files():
    st.title("🚀 معالجة الملفات")
    st.markdown("---")
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("⬅️ العودة لإعداد Placeholders"):
            st.session_state.current_step = 2
            st.rerun()
    st.markdown("### 📂 رفع ملف الصور")
    uploaded_zip = st.file_uploader(
        "اختر ملف ZIP يحتوي على مجلدات صور",
        type=["zip"],
        key="zip_uploader",
        help="ارفع ملف مضغوط يحتوي على مجلدات، كل مجلد يحتوي على صور لشريحة واحدة"
    )
    with st.expander("📋 ملخص الإعدادات المحددة", expanded=True):
        if st.session_state.placeholders_config:
            image_config = st.session_state.placeholders_config.get('images', {})
            text_config = st.session_state.placeholders_config.get('texts', {})
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### 🖼️ إعدادات الصور:")
                active_images = sum(1 for config in image_config.values() if config['use'])
                for key, config in image_config.items():
                    if config['use']:
                        st.success(f"✅ صورة {config['order']}: سيتم استبدالها")
                    else:
                        st.info(f"⏭️ صورة: لن يتم استبدالها")
                if active_images == 0:
                    st.warning("⚠️ لم يتم تحديد أي صور للاستبدال")
            with col2:
                st.markdown("#### 📝 إعدادات النصوص:")
                active_texts = sum(1 for config in text_config.values() if config['type'] != 'ترك فارغ')
                for key, config in text_config.items():
                    if config['type'] != 'ترك فارغ':
                        st.success(f"✅ {config['type']}: {config.get('value', 'تلقائي')}")
                    else:
                        st.info(f"⏭️ نص: سيترك فارغاً")
                if active_texts == 0:
                    st.info("ℹ️ جميع النصوص ستترك فارغة")
    st.markdown("### ⚙️ خيارات إضافية")
    col1, col2 = st.columns(2)
    with col1:
        image_order_option = st.radio(
            "ترتيب الصور في المجلدات:",
            ("بالترتيب الأبجدي", "عشوائي"),
            index=0,
            help="كيف تريد ترتيب الصور داخل كل مجلد قبل التطبيق"
        )
    with col2:
        skip_empty_folders = st.checkbox(
            "تخطي المجلدات الفارغة",
            value=True,
            help="تجاهل المجلدات التي لا تحتوي على صور"
        )
    if uploaded_zip:
        if st.button("🚀 بدء المعالجة", type="primary"):
            clear_details()
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    with st.spinner("📦 جاري استخراج الملفات..."):
                        zip_bytes = io.BytesIO(uploaded_zip.read())
                        with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                            for member in zip_ref.namelist():
                                if os.path.isabs(member) or ".." in member:
                                    raise Exception("ZIP contains unsafe paths!")
                            zip_ref.extractall(temp_dir)
                    add_detail("📂 تم استخراج الملف المضغوط بنجاح", "success")
                    all_items = os.listdir(temp_dir)
                    folder_paths = []
                    for item in all_items:
                        item_path = os.path.join(temp_dir, item)
                        if os.path.isdir(item_path):
                            imgs_in_folder = [
                                f for f in os.listdir(item_path)
                                if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))
                            ]
                            if imgs_in_folder:
                                folder_paths.append(item_path)
                                add_detail(f"📁 المجلد '{item}' يحتوي على {len(imgs_in_folder)} صورة", "info")
                            elif not skip_empty_folders:
                                add_detail(f"⚠ المجلد '{item}' فارغ من الصور", "warning")
                    if not folder_paths:
                        st.error("❌ لا توجد مجلدات تحتوي على صور في الملف المضغوط.")
                        add_detail("❌ لا توجد مجلدات تحتوي على صور", "error")
                        show_details_section()
                        st.stop()
                    folder_paths.sort()
                    add_detail(f"✅ تم العثور على {len(folder_paths)} مجلد يحتوي على صور", "success")
                    prs = Presentation(io.BytesIO(st.session_state.pptx_data))
                    if len(prs.slides) == 0:
                        st.error("❌ لا توجد شرائح في ملف PowerPoint")
                        st.stop()
                    first_slide = prs.slides[0]
                    slide_layout = first_slide.slide_layout
                    total_processed = 0
                    created_slides = 0
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    for folder_idx, folder_path in enumerate(folder_paths):
                        folder_name = os.path.basename(folder_path)
                        status_text.text(f"🔄 معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: {folder_name}")
                        try:
                            imgs = [
                                f for f in os.listdir(folder_path)
                                if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))
                            ]
                            if image_order_option == "عشوائي":
                                random.shuffle(imgs)
                                add_detail(f"🔀 تم ترتيب صور المجلد {folder_name} عشوائياً", "info")
                            else:
                                imgs.sort()
                                add_detail(f"📋 تم ترتيب صور المجلد {folder_name} أبجدياً", "info")
                            new_slide = prs.slides.add_slide(slide_layout)
                            created_slides += 1
                            apply_configured_placeholders(
                                new_slide,
                                folder_path,
                                folder_name,
                                st.session_state.slide_analysis,
                                st.session_state.placeholders_config
                            )
                            total_processed += len(imgs)
                            add_detail(f"✅ تم إنشاء شريحة للمجلد '{folder_name}' مع {len(imgs)} صورة", "success")
                        except Exception as e:
                            add_detail(f"❌ خطأ في معالجة المجلد {folder_name}: {e}", "error")
                        progress_bar.progress((folder_idx + 1) / len(folder_paths))
                    progress_bar.empty()
                    status_text.empty()
                    st.success("🎉 تم الانتهاء من المعالجة بنجاح!")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("الشرائح المُضافة", created_slides)
                    with col2:
                        st.metric("المجلدات المُعالجة", len(folder_paths))
                    with col3:
                        st.metric("إجمالي الصور", total_processed)
                    if created_slides == 0:
                        st.error("❌ لم يتم إضافة أي شرائح.")
                        show_details_section()
                        st.stop()
                    output_filename = f"PowerPoint_Updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    output_buffer = io.BytesIO()
                    prs.save(output_buffer)
                    output_buffer.seek(0)
                    st.download_button(
                        label="⬇️ تحميل الملف المُحدث",
                        data=output_buffer.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary"
                    )
                    if st.button("🔄 بدء عملية جديدة"):
                        for key in list(st.session_state.keys()):
                            del st.session_state[key]
                        st.rerun()
                    if not st.session_state.show_details_needed:
                        if st.button("📋 إظهار تفاصيل المعالجة"):
                            show_details_section()
                    else:
                        show_details_section()
                except Exception as e:
                    st.error(f"❌ خطأ أثناء المعالجة: {e}")
                    add_detail(f"❌ خطأ عام أثناء المعالجة: {e}", "error")
                    show_details_section()

def main():
    st.set_page_config(
        page_title="PowerPoint Image Replacer",
        layout="wide",
        page_icon="🔄"
    )
    init_session()
    st.markdown("""
    <style>
    .stExpander > div:first-child { background-color: #f0f2f6; }
    .stMetric { background-color: #ffffff; padding: 1rem; border-radius: 0.5rem; border: 1px solid #e1e5e9; }
    .slide-preview { border: 2px solid #ddd; border-radius: 10px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }
    </style>
    """, unsafe_allow_html=True)
    if st.session_state.current_step == 1:
        step1_upload_pptx()
    elif st.session_state.current_step == 2:
        step2_configure_placeholders()
    elif st.session_state.current_step == 3:
        step3_process_files()
    st.markdown("---")
    progress_labels = ["📂 رفع القالب", "⚙️ إعداد Placeholders", "🚀 المعالجة"]
    cols = st.columns(3)
    for i, (col, label) in enumerate(zip(cols, progress_labels)):
        with col:
            if i + 1 < st.session_state.current_step:
                st.success(f"✅ {label}")
            elif i + 1 == st.session_state.current_step:
                st.info(f"🔄 {label}")
            else:
                st.write(f"⏳ {label}")

if __name__ == '__main__':
    main()
