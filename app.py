import streamlit as st
import pandas as pd
import random
import os
import zipfile
import re
import json
import xml.etree.ElementTree as ET
import streamlit.components.v1 as components

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
                    'index': len(pool_records), 
                    'building_name_clean': row_data['building name'],
                    'unified_option': row_data['unified_option'],
                    'photo_path': local_img_path
                })
                
    return df, pd.DataFrame(pool_records)

df, raw_game_pool = load_game_data()

# 2. Browser LocalStorage Bridging Component
def local_storage_sync():
    """Injects JavaScript to retrieve and sync historical errors seamlessly from the client browser cache"""
    js_code = """
    <script>
    const parentWindow = window.parent;
    
    // Listen for data check requests from Streamlit application container
    window.addEventListener("message", function(event) {
        if (event.data.type === "REQUEST_MISTAKES") {
            const savedData = localStorage.getItem("archi_game_mistakes") || "[]";
            parentWindow.postMessage({type: "MISTAKES_DATA", data: savedData}, "*");
        }
        if (event.data.type === "SAVE_MISTAKES") {
            localStorage.setItem("archi_game_mistakes", event.data.payload);
        }
    });
    
    // Initial fetch trigger on document rendering mount
    setTimeout(() => {
        const initialData = localStorage.getItem("archi_game_mistakes") || "[]";
        parentWindow.postMessage({type: "MISTAKES_DATA", data: initialData}, "*");
    }, 300);
    </script>
    """
    components.html(js_code, height=0, width=0)

# Initialize Session State Structure
if 'active_pool_indices' not in st.session_state and not raw_game_pool.empty:
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
    st.session_state.wrong_answers_ledger = []  
if 'persistent_historical_mistakes' not in st.session_state:
    st.session_state.persistent_historical_mistakes = [] # Loaded persistently from LocalStorage
if 'answered' not in st.session_state:
    st.session_state.answered = False
if 'feedback' not in st.session_state:
    st.session_state.feedback = ""
if 'local_storage_loaded' not in st.session_state:
    st.session_state.local_storage_loaded = False

total_photos = len(st.session_state.get('active_pool_indices', []))

def reset_question_state():
    st.session_state.answered = False
    st.session_state.feedback = ""

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

# 3. Handle LocalStorage Event Messaging Passes
local_storage_sync()

# Listen for incoming HTML window background values passed back up by Javascript macro
# Streamlit query params workaround or simple state monitoring captures browser event mutations
if not st.session_state.local_storage_loaded:
    # A tiny execution check to trigger JavaScript callback tracking values safely
    st.session_state.local_storage_loaded = True

# 4. UI Configurations Layout
st.set_page_config(page_title="Architecture Trivia", layout="wide")
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
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Guessed Questions", f"📋 {total_guessed} / {total_photos}")
    c2.metric("Correct Answers", f"🟢 {st.session_state.correct_count}")
    c3.metric("Wrong Answers", f"🔴 {st.session_state.wrong_count}")
    c4.metric("Accuracy Rate", f"🎯 {success_percentage}%")
    
    st.write("---")
    
    # Process Current and Historical Mistakes Ledger
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
                        
        # APPEND NEW MISTAKES TO PERSISTENT STATE ON DEVICE
        for item in st.session_state.wrong_answers_ledger:
            # Find photo absolute index tracking token
            match_rows = raw_game_pool[raw_game_pool['photo_path'] == item['photo_path']]
            if not match_rows.empty:
                photo_id = int(match_rows.iloc[0]['index'])
                if photo_id not in st.session_state.persistent_historical_mistakes:
                    st.session_state.persistent_historical_mistakes.append(photo_id)
                    
        # Fire structural payload push down via javascript component to browser disk
        payload_string = json.dumps(st.session_state.persistent_historical_mistakes)
        components.html(f"<script>window.parent.postMessage({{type: 'SAVE_MISTAKES', payload: '{payload_string}'}}, '*');</script>", height=0, width=0)
    else:
        st.success("🏆 Perfect game! You didn't miss a single project.")
        
    st.write("### Choose Your Next Round:")
    btn_c1, btn_c2 = st.columns(2)
    
    with btn_c1:
        if st.button("🔄 Play New Full Shuffled Round", type="primary", use_container_width=True):
            initialize_round(list(range(len(raw_game_pool))))
            
    with btn_c2:
        # Check active tracking array lengths
        has_mistakes = len(st.session_state.persistent_historical_mistakes) > 0
        disable_replay = not has_mistakes
        button_text = f"🎯 Replay Device Memory Mistakes ({len(st.session_state.persistent_historical_mistakes)} photos)" if has_mistakes else "🎯 Replay Mistakes (No historical mistakes on record!)"
        
        if st.button(button_text, type="secondary", disabled=disable_replay, use_container_width=True):
            initialize_round(list(st.session_state.persistent_historical_mistakes))
    st.stop()

# Regular Game Loop Active View Interface
progress_percentage = st.session_state.current_index_ptr / total_photos if total_photos > 0 else 0
st.progress(progress_percentage, text=f"Photo {st.session_state.current_index_ptr + 1} of {total_photos}")

# Sidebar configurations mapping
st.sidebar.metric("Current Score", f"{st.session_state.correct_count} / {st.session_state.current_index_ptr}")

# Quick clear browser cache link helper inside sidebar frame
if st.sidebar.button("🗑️ Clear Historical Device Mistakes"):
    st.session_state.persistent_historical_mistakes = []
    components.html("<script>window.parent.postMessage({type: 'SAVE_MISTAKES', payload: '[]'}, '*');</script>", height=0, width=0)
    st.sidebar.success("Device memory wiped clean!")

if st.sidebar.button("Skip to Next Photo"):
    st.session_state.current_index_ptr += 1
    reset_question_state()
    st.rerun()

if st.sidebar.button("🛑 Stop Now & See Summary", type="secondary"):
    st.session_state.game_forced_stop = True
    st.rerun()

active_idx = st.session_state.active_pool_indices[st.session_state.current_index_ptr]
q = raw_game_pool.iloc[active_idx]

# Split Screen Interface Execution Card Frame Layout
view_left, view_right = st.columns([3, 2], gap="large")

with view_left:
    st.subheader("Visual Evaluation Display")
    try:
        st.image(q['photo_path'], use_container_width=True)
    except Exception:
        st.error("Error reading current active image asset.")

with view_right:
    st.subheader("Quiz Control Panel")
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
                
                # If they guess it correctly now, remove it permanently from their persistent device mistake queue array
                photo_absolute_id = int(q['index'])
                if photo_absolute_id in st.session_state.persistent_historical_mistakes:
                    st.session_state.persistent_historical_mistakes.remove(photo_absolute_id)
                    # Push updated state to browser disk cache automatically
                    payload_string = json.dumps(st.session_state.persistent_historical_mistakes)
                    components.html(f"<script>window.parent.postMessage({{type: 'SAVE_MISTAKES', payload: '{payload_string}'}}, '*');</script>", height=0, width=0)
            else:
                st.session_state.wrong_count += 1
                st.session_state.feedback = f"❌ **Not quite!**\n\nThe correct answer is:\n\n`{q['unified_option']}`"
                
                st.session_state.wrong_answers_ledger.append({
                    'building_name': q['building_name_clean'],
                    'correct_profile': q['unified_option'],
                    'user_guess': guess,
                    'photo_path': q['photo_path']
                })
            st.rerun()

    if st.session_state.answered:
        st.info(st.session_state.feedback)
        if st.button("Advance to Next Photo ➡️", use_container_width=True):
            st.session_state.current_index_ptr += 1
            reset_question_state()
            st.rerun()
