import streamlit as st
import google.generativeai as genai
from audiorecorder import audiorecorder

# Page Configuration with Corporate Blue Identity
st.set_page_config(page_title="WNW Holidays Voice Engine", layout="wide")

# Theme styling anchors
st.markdown("""
    <style>
    .reportview-container { background: #f0f5fa; }
    .stButton>button { background-color: #0056b3; color: white; border-radius: 6px; }
    </style>
""", unsafe_allow_html=True)

st.title("🦅 Wings 'N' Wheels Holidays Voice Portal")
st.caption("Commercial Multi-Step Voice AI Itinerary Automation Engine")

# 1. INITIALIZE MASTER ITINERARY DATABASE
MASTER_PACKAGES = {
    "Chandigarh – Manali Tour (6N/7D)": {
        "route": "CAMPUS → AJMER → CHANDIGARH → MANALI → AJMER → CAMPUS",
        "inclusions": "- 01 Night Stay in CHANDIGARH\\n- 02 Night Stay in MANALI hotel\\n- 01 Night Stay in MANALI camp\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Visit By Luxury coach (Non AC)\\n- All Meals (veg) – As per Itinerary",
        "exclusions": "- All Personal Expenses\\n- RIVER RAFTING\\n- Tips\\n- Anything Which Is Not Mentioned in Inclusions",
        "days_count": 6
    },
    "Dalhousie – Dharamshala – Amritsar Tour (7N/8D)": {
        "route": "FALNA → AJMER → DALHOUSIE → DHARAMSHALA → AMRITSAR → AJMER → FALNA",
        "inclusions": "- 05 Night Stay in DALHOUSIE / DHARAMSHALA / AMRITSAR hotels\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Visit By Luxury coach (Non AC)\\n- All Meals (veg) – As per Itinerary\\n- All Transportation Expense",
        "exclusions": "- All Personal Expenses\\n- Tips\\n- Anything Which Is Not Mentioned in Inclusions",
        "days_count": 7
    }
}

if "step" not in st.session_state:
    st.session_state.step = 1
    st.session_state.trip_data = {}

# STEP 1: PARSE INITIAL REQUIREMENTS WITH MICROPHONE
if st.session_state.step == 1:
    st.subheader("📍 Step 1: Voice Verification Log")
    api_input = st.text_input("Enter your Gemini API Key to wake up the agent:", type="password")
    selected_pack = st.selectbox("Select Target Master Itinerary Template:", list(MASTER_PACKAGES.keys()))
    
    st.markdown("### 🎙️ Speak Your Basic Parameters:")
    st.write("Click record and say: *'Starting from Pali, train at 17:00 from Ajmer.'* (You can speak in Hindi or English)")
    
    # Renders an interactive voice layout button
    audio = audiorecorder("Click to Record Voice", "Click to Stop Recording")
    
    if len(audio) > 0:
        st.audio(audio.export().read())
        st.success("Audio captured! Processing voice transcripts...")
        
    start_place = st.text_input("Confirm Starting Location:", "Pali")
    boarding_hub = st.text_input("Confirm Boarding Hub:", "Ajmer")
    train_time = st.text_input("Confirm Train Departure Time:", "17:00")
    
    if st.button("Process Routing Logs"):
        if not api_input:
            st.error("Please enter your API key to continue.")
        else:
            st.session_state.api_key = api_input
            st.session_state.trip_data.update({
                "package": selected_pack, "start": start_place, 
                "hub": boarding_hub, "train_time": train_time
            })
            if start_place.lower() != boarding_hub.lower():
                st.session_state.step = 2
            else:
                st.session_state.step = 3
            st.rerun()

# STEP 2: VOICE DISCREPANCY RECONCILIATION
if st.session_state.step == 2:
    st.warning(f"⚠️ Logistics Discrepancy Flagged: {st.session_state.trip_data['start']} to {st.session_state.trip_data['hub']}.")
    
    st.markdown("### 🎙️ Speak Connecting Plan:")
    st.write("Click record and explain how you will go from Pali to Ajmer:")
    audio_gap = audiorecorder("Speak Connecting Plan", "Stop Recording")
    
    transfer_method = st.text_input("Or verify layout text here:", "By private road bus transfer via NH58")
    client_count = st.number_input("Total Student Headcount:", min_value=1, value=45)
    
    if st.button("Lock Final Systems"):
        st.session_state.trip_data.update({
            "transfer_method": transfer_method,
            "headcount": client_count,
            "buffer_calc": "3 Hours road transit"
        })
        st.session_state.step = 3
        st.rerun()

# STEP 3: OUTPUT PACKAGING ENGINE
if st.session_state.step == 3:
    st.subheader("⚙️ Final Step: Compiling Proposal Layout")
    if st.button("Compile PDF Proposal Documents"):
        genai.configure(api_key=st.session_state.api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
        pack_info = MASTER_PACKAGES[st.session_state.trip_data["package"]]
        
        ai_prompt = f"""
        You are a strict data-entry automation agent for Wings 'N' Wheels Holidays. 
        Take these custom parameters and map them neatly into our commercial routing rules.
        
        Inputs:
        - Origin: {st.session_state.trip_data['start']}
        - Train Station: {st.session_state.trip_data['hub']}
        - Train Departure: {st.session_state.trip_data['train_time']}
        - Interlock Transit Buffer: {st.session_state.trip_data.get('transfer_method', '')}
        
        Output three distinct markdown chapters: A spoken verification block in conversational Hindi text for customer follow-up validation calls, the day-by-day timestamps itinerary, and the hardcoded costing conditions.
        
        Static Inclusions Block:
        {pack_info['inclusions']}
        
        Static Exclusions Block:
        {pack_info['exclusions']}
        """
        with st.spinner("Assembling proposal frames..."):
            response = model.generate_content(ai_prompt)
            st.session_state.final_output = response.text
            st.session_state.step = 4
            st.rerun()

if st.session_state.step == 4:
    st.success("🎉 Document Created!")
    st.markdown(st.session_state.final_output)
    if st.button("Start New Run File"):
        st.session_state.clear()
        st.session_state.step = 1
        st.rerun()
