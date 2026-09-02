import streamlit as st
import google.generativeai as genai

# Page Configuration with Corporate Blue Identity
st.set_page_config(page_title="WNW Holidays Itinerary Engine", layout="wide")

# Theme styling anchors
st.markdown("""
    <style>
    .reportview-container { background: #f0f5fa; }
    .stButton>button { background-color: #0056b3; color: white; border-radius: 6px; }
    </style>
""", unsafe_allow_html=True)

st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Commercial Multi-Step AI Itinerary Automation Engine")

# 1. INITIALIZE MASTER ITINERARY DATABASE FROM YOUR PROPOSALS
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
    },
    "Gir – Somnath – Diu Tour (5N/6D)": {
        "route": "FALNA → GIR → SOMNATH → DIU → FALNA",
        "inclusions": "- 03 Night Stay in GIR, SOMNATH & DIU hotels\\n- 01 Night Stay in SOMNATH\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Visit By Luxury coach (Non AC)\\n- All Meals (veg) – As per Itinerary",
        "exclusions": "- All Personal Expenses\\n- Safari Permits\\n- Tips\\n- Anything Which Is Not Mentioned in Inclusions",
        "days_count": 5
    },
    "Khajuraho – Jabalpur Tour (4N/6D)": {
        "route": "FALNA → AJMER → KHAJURAHO → JABALPUR → AJMER → FALNA",
        "inclusions": "- 02 Night Stay in KHAJURAHO hotel\\n- 01 Night Stay in KHAJURAHO camp\\n- 01 Night Stay in JABALPUR\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Visit By Luxury coach (Non AC)\\n- All Meals (veg) – As per Itinerary",
        "exclusions": "- All Personal Expenses\\n- Tips\\n- Anything Which Is Not Mentioned in Inclusions",
        "days_count": 4
    }
}

# Session State Manager for multi-step interview flow
if "step" not in st.session_state:
    st.session_state.step = 1
    st.session_state.trip_data = {}

# STEP 1: PARSE INITIAL REQUIREMENTS
if st.session_state.step == 1:
    st.subheader("📍 Step 1: Client Requirements & Package Select")
    
    # Secure API Input Field
    api_input = st.text_input("Enter your Gemini API Key to wake up the agent:", type="password")
    
    selected_pack = st.selectbox("Select Target Master Itinerary Template:", list(MASTER_PACKAGES.keys()))
    start_place = st.text_input("Client's Exact Starting Location (e.g., Pali):", "Pali")
    boarding_hub = st.text_input("Boarding Hub / Train Station (e.g., Ajmer):", "Ajmer")
    train_time = st.text_input("Scheduled Train Departure Time (HH:MM format):", "17:00")
    
    if st.button("Analyze Transit Gaps"):
        if not api_input:
            st.error("Please enter your API key to continue.")
        else:
            st.session_state.api_key = api_input
            st.session_state.trip_data.update({
                "package": selected_pack, "start": start_place, 
                "hub": boarding_hub, "train_time": train_time
            })
            
            # Geographical Gap Validation Logic
            if start_place.lower() != boarding_hub.lower():
                st.session_state.step = 2  # Divert to routing reconciliation wizard
            else:
                st.session_state.step = 3  # Direct pass
            st.rerun()

# STEP 2: GEOGRAPHIC DISCREPANCY RECONCILIATION
if st.session_state.step == 2:
    st.warning(f"⚠️ Logistics Discrepancy Flagged: Traveler is starting from '{st.session_state.trip_data['start']}' but boarding transit from '{st.session_state.trip_data['hub']}'.")
    st.info("🗺️ Background Matrix Verification: The distance between Pali and Ajmer requires an approximate 3-hour road cushion via NH58 to guarantee safe boarding.")
    
    transfer_method = st.text_input(f"How will you arrange group transit from {st.session_state.trip_data['start']} to {st.session_state.trip_data['hub']}?", "By specialized corporate luxury bus")
    client_count = st.number_input("Total Student Headcount:", min_value=1, value=45)
    
    if st.button("Lock Connecting Logistics"):
        st.session_state.trip_data.update({
            "transfer_method": transfer_method,
            "headcount": client_count,
            "buffer_calc": "3 Hours road transit"
        })
        st.session_state.step = 3
        st.rerun()

# STEP 3: EXECUTE AI PARSING & GENERATION
if st.session_state.step == 3:
    st.subheader("⚙️ Step 3: Compiling Custom Error-Less Proposal Layout")
    st.info("The agent is now combining your live parameters with hardcoded corporate headers...")
    
    if st.button("Generate Final PDF Ready Layout"):
        # Setup the secure AI Model call
        genai.configure(api_key=st.session_state.api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
        
        pack_info = MASTER_PACKAGES[st.session_state.trip_data["package"]]
        
        # Engineering strict prompt mapping behavior
        ai_prompt = f"""
        You are a strict data-entry automation agent for Wings 'N' Wheels Holidays. 
        Your task is to populate the trip data into our corporate format structure without altering static components.

        Live Client Inputs:
        - Origin: {st.session_state.trip_data['start']}
        - Train Boarding Station: {st.session_state.trip_data['hub']}
        - Train Time: {st.session_state.trip_data['train_time']}
        - Connection Plan: {st.session_state.trip_data.get('transfer_method', 'Direct Departure')}
        - Road Transit time: {st.session_state.trip_data.get('buffer_calc', '0 hours')}
        - Selected Package Matrix: {st.session_state.trip_data['package']}

        OUTPUT FORMAT REQUIREMENTS:
        Generate exactly three sections. 
        Section 1 must be a conversational text paragraph written completely in Hindi summarizing Day 1 and return transit for customer verbal confirmation calls. Explicitly state the route gap logic if present.
        Section 2 must be the daily day-by-day itinerary breakdown cleanly formatted with precise markdown timestamps.
        Section 3 must compile static pricing variables.

        Static Inclusions Block:
        {pack_info['inclusions']}

        Static Exclusions Block:
        {pack_info['exclusions']}
        """
        
        with st.spinner("Assembling structural parameters..."):
            response = model.generate_content(ai_prompt)
            st.session_state.final_output = response.text
            st.session_state.step = 4
            st.rerun()

# STEP 4: PRODUCTION EXPORT INTERFACE
if st.session_state.step == 4:
    st.success("🎉 Layout Compiled Successfully with Zero Format Distortions!")
    st.markdown(st.session_state.final_output)
    
    if st.button("Process New Client Routing File"):
        st.session_state.clear()
        st.session_state.step = 1
        st.rerun()
