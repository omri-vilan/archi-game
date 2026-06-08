import streamlit as st
import pandas as pd
import random
import os
import zipfile
import re
import xml.etree.ElementTree as ET

# App Directories Setup
IMAGE_DIR = "images"
os.makedirs(IMAGE_DIR, exist_ok=True)
EXCEL_FILE = "data.xlsx"

# 1. OpenXML Extraction Pipeline
def extract_rich_value_images(zip_path):
    cell_to_image_map = {}
    if not os.path.exists(zip_path):
        return cell_to_image_map
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            file_list = z.namelist()
            sheet_path = 'xl/worksheets/sheet1.xml'
            if sheet_path not in file_list:
                return cell_to_image_map
                
            sheet_content = z.read(sheet_path).decode('utf-8')
            cell_vm_matches = re.findall(r'<c\s+[^>]*r="([A-Z]+[0-9]+)"[^>]*vm="([0-9]+)"', sheet_content)
            if not cell_vm_matches:
                cell_vm_matches = re.findall(r'vm="([0-9]+)"[^>]*r="([A-Z]+[0-9]+)"', sheet_content)
                cell_vm_matches = [(r, v) for v, r in cell_vm_matches]

            cell_to_vm = {cell: int(vm) for cell, vm in cell_vm_matches}

            rv_rel_path = 'xl/richData/richValueRel.xml'
            if rv_rel_path not in file_list:
                return cell_to_image_map
                
            rv_rel_content = z.read(rv_rel_path)
            root_rv = ET.fromstring(rv_rel_content)
            rIds = []
            for rel in root_rv.findall('.//rel', {'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}):
                rIds.append(rel.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'))
            if not rIds:
                rIds = re.findall(r'r:id="([^"]+)"', rv_rel_content.decode('utf-8'))

            rels_path = 'xl/richData/_rels/richValueRel.xml.rels'
            if rels_path not in file_list:
                return cell_to_image_map
                
            rels_content = z.read(rels_path)
            root_rels = ET.fromstring(rels_content)
            rid_to_target = {}
            for rel in root_rels.findall('.//Relationship', {'': 'http://schemas.openxmlformats.org/package/2006/relationships'}):
                rid_to_target[rel.get('Id')] = rel.get('Target')
            if not rid_to_target:
                rid_matches = re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_content.decode('utf-8'))
                rid_to_target = {rid: target for rid, target in rid_matches}

            for cell, vm in cell_to_vm.items():
                array_idx = vm - 1 
                if 0 <= array_idx < len(rIds):
                    target_rid = rIds[array_idx]
                    if target_rid in rid_to_target:
                        target_img_filename = os.path.basename(rid_to_target[target_rid])
                        zip_img_path = f"xl/media/{target_img_filename}"
                        if zip_img_path in file_list:
                            local_path = os.path.join(IMAGE_DIR, f"{cell}_{target_img_filename}")
                            with open(local_path, "wb") as img_out:
                                img_out.write(z.read(zip_img_path))
                            cell_to_image_map[cell] = local_path
    except Exception as e:
        st.error(f"Failed parsing modern Excel schema layer: {e}")
    return cell_to_image_map

@st.cache_data
def load_game_data():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.astype(str).str.strip().str.lower()
    
    base_cols = ['building name', 'architect name', 'location', 'year']
    for col in base_cols:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
            
    df['unified_option'] = df['building name'] + " / " + df['architect name'] + " / " + df['year']
    cell_img_map = extract_rich_value_images(EXCEL_FILE)
    col_to_name = {'E': 'photo 1', 'F': 'photo 2', 'G': 'photo 3', 'H': 'photo 4'}
    
    pool_records = []
    for cell_coord, local_img_path in cell_img_map.items():
        match = re.match(r"([A-Z]+)([0-9]+)", cell_coord)
        if match:
            col_letter, row_num = match.groups()
            df_row_idx = int(row_num) - 2
            if col_letter in col_to_name and 0 <= df_row_idx < len(df):
                row_data = df.iloc[df_row_idx]
                pool_records.append({
                    'index': len(pool_records), # Keep absolute track of each unique photo record
                    'building_name_clean': row_data['building name'],
                    'unified_option': row_data['unified_option'],
                    'photo_path': local_img_path
                })
                
    return df, pd.DataFrame(pool_records)

df, raw_game_pool = load_game_data()

# 2. Advanced Session State Control
if 'active_pool_indices' not in st.session_state and not raw_game_pool.empty:
    # Full deck mode on initial startup
    full_indices = list(range(len(raw_game_pool)))
    random.shuffle(full_indices)
    st.session_state.active_pool_indices = full_indices

if 'current_index_ptr' not in st.session_state:
    st.session_state.current_index_ptr = 0
if 'correct_count' not in st.session_state:
    st.session_state.correct_count = 0
if 'wrong_count' not in st.session_state:
    st.session_state.wrong_count = 0
if 'game_forced_stop' not in st.session_state:
    st.session_state.game_forced_stop = False
if 'wrong_answers_ledger' not in st.session_state:
    st.session_state.wrong_answers_ledger = []  # Running log of mistake objects
if 'last_round_mistake_pool' not in st.session_state:
    st.session_state.last_round_mistake_pool = [] # Retains photo IDs missed for re-testing
if 'answered' not in st.session_state:
    st.session_state.answered = False
if 'feedback' not in st.session_state:
    st.session_state.feedback = ""

total_photos = len(st.session_state.get('active_pool_indices', []))

def reset_question_state():
    st.session_state.answered = False
    st.session_state.feedback = ""

# Helper to prepare subsequent sessions
def initialize_round(selected_indices):
    random.shuffle(selected_indices)
    st.session_state.active_pool_indices = selected_indices
    st.session_state.current_index_ptr = 0
    st.session_state.correct_count = 0
    st.session_state.wrong_count = 0
    st.session_state.game_forced_stop = False
    st.session_state.wrong_answers_ledger = []
    reset_question_state()
    st.rerun()

# 3. UI Configurations Layout
st.set_page_config(page_title="Architecture Trivia", layout="wide") # Shifted to wide format to maximize separation
st.title("🏛️ Architecture Guessing Game")

if raw_game_pool.empty:
    st.error("⚠️ No playable images loaded from Excel schema layers.")
    st.stop()

# Evaluate Game State
is_game_over = st.session_state.game_forced_stop or (st.session_state.current_index_ptr >= total_photos)

if is_game_over:
    st.balloons()
    st.header("🏁 Game Summary")
    
    total_guessed = st.session_state.correct_count + st.session_state.wrong_count
    success_percentage = int((st.session_state.correct_count / total_guessed) * 100) if total_guessed > 0 else 0
    
    # Updated Visual Performance Metrics Dashboard
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Guessed Questions", f"📋 {total_guessed} / {total_photos}")
    c2.metric("Correct Answers", f"🟢 {st.session_state.correct_count}")
    c3.metric("Wrong Answers", f"🔴 {st.session_state.wrong_count}")
    c4.metric("Accuracy Rate (% of Guessed)", f"🎯 {success_percentage}%")
    
    st.write("---")
    
    # Display Missteps Grouped by Building Name
    if st.session_state.wrong_answers_ledger:
        st.subheader("🕵️ Review Your Missed Buildings:")
        errors_df = pd.DataFrame(st.session_state.wrong_answers_ledger)
        grouped = errors_df.groupby('building_name')
        
        for building, group in grouped:
            with st.expander(f"🏢 {building} ({len(group)} misidentified photo{'s' if len(group) > 1 else ''})", expanded=True):
                st.caption(f"**Correct Target Profile:** {group.iloc[0]['correct_profile']}")
                thumb_cols = st.columns(min(len(group), 4))
                for idx, (_, row_item) in enumerate(group.iterrows()):
                    with thumb_cols[idx % 4]:
                        st.image(row_item['photo_path'], use_container_width=True)
                        st.error(f"Guessed: {row_item['user_guess'].split(' / ')[0] if '/' in row_item['user_guess'] else row_item['user_guess']}")
    else:
        st.success("🏆 Perfect game! You didn't miss a single project.")
        
    st.write("### Choose Your Next Round:")
    btn_c1, btn_c2 = st.columns(2)
    
    with btn_c1:
        if st.button("🔄 Play New Full Shuffled Round", type="primary", use_container_width=True):
            initialize_round(list(range(len(raw_game_pool))))
            
    with btn_c2:
        # Build alternative mistake replay selection pool dynamically based on recent log saves
        has_mistakes = len(st.session_state.last_round_mistake_pool) > 0
        disable_replay = not has_mistakes
        button_text = f"🎯 Replay Mistakes Only ({len(st.session_state.last_round_mistake_pool)} photos)" if has_mistakes else "🎯 Replay Mistakes Only (No mistakes last round!)"
        
        if st.button(button_text, type="secondary", disabled=disable_replay, use_container_width=True):
            initialize_round(list(st.session_state.last_round_mistake_pool))
    st.stop()

# Regular Game Loop Active View Interface
progress_percentage = st.session_state.current_index_ptr / total_photos if total_photos > 0 else 0
st.progress(progress_percentage, text=f"Photo {st.session_state.current_index_ptr + 1} of {total_photos}")

# Sidebar controls column mapping
st.sidebar.metric("Current Score", f"{st.session_state.correct_count} / {st.session_state.current_index_ptr}")

if st.sidebar.button("Skip to Next Photo"):
    st.session_state.current_index_ptr += 1
    reset_question_state()
    st.rerun()

if st.sidebar.button("🛑 Stop Now & See Summary", type="secondary"):
    st.session_state.game_forced_stop = True
    st.rerun()

# Extract active photo row
active_idx = st.session_state.active_pool_indices[st.session_state.current_index_ptr]
q = raw_game_pool.iloc[active_idx]

# FIXED INTERFACE: Side-by-side split screen layout column allocations
view_left, view_right = st.columns([3, 2], gap="large")

with view_left:
    st.subheader("Visual Evaluation Display")
    try:
        st.image(q['photo_path'], use_container_width=True)
    except Exception:
        st.error("Error reading current active image asset.")

with view_right:
    st.subheader("Quiz Control Panel")
    st.write("Analyze the image on the left and select its matching architectural specifications profile below.")
    st.write("---")
    
    full_options_pool = sorted(df['unified_option'].dropna().unique().tolist())
    guess = st.selectbox(
        "Choose the correct combination (Building / Architect / Year)", 
        full_options_pool, 
        index=None, 
        placeholder="Select building, architect, and year..."
    )
    
    submit_btn = st.button("Submit Guess", disabled=st.session_state.answered, type="primary", use_container_width=True)

    if submit_btn:
        if not guess:
            st.warning("Please make a selection first!")
        else:
            st.session_state.answered = True
            is_correct = str(guess).strip() == str(q['unified_option']).strip()
            
            if is_correct:
                st.session_state.correct_count += 1
                st.session_state.feedback = "🎉 **Correct!** Excellent job."
            else:
                st.session_state.wrong_count += 1
                st.session_state.feedback = f"❌ **Not quite!**\n\nThe correct answer is:\n\n`{q['unified_option']}`"
                
                # Log mistake metadata down to display on the grouped final card layout
                st.session_state.wrong_answers_ledger.append({
                    'building_name': q['building_name_clean'],
                    'correct_profile': q['unified_option'],
                    'user_guess': guess,
                    'photo_path': q['photo_path']
                })
                
                # Append this unique photo's core index key to the mistake replay pool array
                photo_absolute_id = int(q['index'])
                if photo_absolute_id not in st.session_state.last_round_mistake_pool:
                    st.session_state.last_round_mistake_pool.append(photo_absolute_id)
            st.rerun()

    if st.session_state.answered:
        st.info(st.session_state.feedback)
        
        # If they guessed correctly, remove that photo from the mistake replay tracking pool
        is_correct = str(guess).strip() == str(q['unified_option']).strip()
        if is_correct and int(q['index']) in st.session_state.last_round_mistake_pool:
            st.session_state.last_round_mistake_pool.remove(int(q['index']))
            
        if st.button("Advance to Next Photo ➡️", use_container_width=True):
            st.session_state.current_index_ptr += 1
            reset_question_state()
            st.rerun()